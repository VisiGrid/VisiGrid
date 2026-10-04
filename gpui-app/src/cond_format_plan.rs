//! Canonical conditional-format targets and private editor drafts. No cell writes.
use crate::{
    formatting::plan::{visible_rows, Range},
    history::UndoAction,
};
use std::collections::{BTreeMap, BTreeSet};
use visigrid_engine::{
    cond_format::{CondFormatRule, CondFormatStore, CondStyle},
    filter::RowView,
    formula::parser::adjust_formula_refs,
    sheet::{Sheet, SheetId},
    validation::CellRange,
    workbook::Workbook,
};

const MAX_CELLS: usize = 100_000;
const MAX_FRAGMENTS: usize = 4096;

pub(crate) fn targets(
    sheet: &Sheet,
    rows: &RowView,
    hidden_rows: Option<&BTreeSet<usize>>,
    hidden_cols: Option<&BTreeSet<usize>>,
    ranges: &[Range],
) -> Result<(Vec<CellRange>, (usize, usize)), String> {
    let mut cells = BTreeMap::<usize, BTreeSet<usize>>::new();
    let mut anchor = None;
    let mut count = 0;
    for &((r1, c1), (r2, c2)) in ranges {
        if r1 > r2 || c1 > c2 || r2 >= sheet.rows || c2 >= sheet.cols {
            return Err("The conditional-format selection is outside the worksheet.".into());
        }
        for (_, row) in visible_rows(rows, hidden_rows, r1, r2) {
            for col in c1..=c2 {
                if hidden_cols.is_some_and(|h| h.contains(&col)) {
                    continue;
                }
                anchor.get_or_insert((row, col));
                if cells.entry(row).or_default().insert(col) {
                    count += 1;
                }
                if count > MAX_CELLS {
                    return Err("Select at most 100,000 visible cells for conditional formatting through a Table view.".into());
                }
            }
        }
    }
    let anchor = anchor.ok_or("No visible cells are selected.")?;
    let mut result: Vec<CellRange> = Vec::new();
    let mut previous = BTreeMap::<(usize, usize), usize>::new();
    for (row, cols) in cells {
        let mut spans = Vec::<(usize, usize)>::new();
        for col in cols {
            if let Some(last) = spans.last_mut().filter(|p| p.1 + 1 == col) {
                last.1 = col;
            } else {
                spans.push((col, col));
            }
        }
        let mut current = BTreeMap::new();
        for (start_col, end_col) in spans {
            let key = (start_col, end_col);
            let index = if let Some(&i) = previous
                .get(&key)
                .filter(|&&i| result[i].end_row + 1 == row)
            {
                result[i].end_row = row;
                i
            } else {
                result.push(CellRange {
                    start_row: row,
                    end_row: row,
                    start_col,
                    end_col,
                });
                if result.len() > MAX_FRAGMENTS {
                    return Err(fragment_error());
                }
                result.len() - 1
            };
            current.insert(key, index);
        }
        previous = current;
    }
    Ok((result, anchor))
}

fn fragment_error() -> String {
    "This selection would create too many conditional-format ranges. Select a smaller range.".into()
}

pub(crate) fn validate_rule(sheet: &Sheet, rule: &CondFormatRule) -> Result<(), String> {
    for r in &rule.ranges {
        if r.start_row > r.end_row
            || r.start_col > r.end_col
            || r.end_row >= sheet.rows
            || r.end_col >= sheet.cols
        {
            return Err("A conditional-format range is outside the worksheet.".into());
        }
        if let Some(spec) = sheet.table_view_spec().filter(|s| s.has_criteria()) {
            let table = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The conditional-format Table no longer exists.")?;
            let t = table.range;
            if r.start_row <= t.end_row
                && r.end_row > t.start_row
                && (r.start_col < t.start_col || r.end_col > t.end_col)
            {
                return Err("Conditional formatting beside this Table would move or hide with its records. Select cells inside the Table, or clear its sorting and filters.".into());
            }
        }
    }
    Ok(())
}

#[derive(Clone)]
pub(crate) struct Draft {
    pub sheet_id: SheetId,
    revision: u64,
    pub before: CondFormatStore,
    pub editing: Option<(usize, CondFormatRule)>,
    pub anchor: (usize, usize),
    pub ranges: Vec<CellRange>,
    pub preview: Option<CondFormatStore>,
}
impl Draft {
    pub fn new(
        wb: &Workbook,
        index: usize,
        ranges: Vec<CellRange>,
        anchor: (usize, usize),
        editing: Option<(usize, CondFormatRule)>,
    ) -> Self {
        let sheet = wb.sheet(index).expect("existing editor sheet");
        Self {
            sheet_id: sheet.id,
            revision: wb.revision(),
            before: sheet.cond_formats.clone(),
            editing,
            anchor,
            ranges,
            preview: None,
        }
    }
    pub fn is_current(&self, wb: &Workbook) -> bool {
        wb.revision() == self.revision && wb.active_sheet_id() == self.sheet_id
    }
    pub fn validate(&self, wb: &Workbook) -> Result<usize, String> {
        wb.ensure_writable()?;
        let index = wb.sheet_index_by_id(self.sheet_id).ok_or(
            "The conditional-format sheet no longer exists. Copy your draft before closing.",
        )?;
        if !self.is_current(wb)
            || !wb
                .sheet(index)
                .unwrap()
                .cond_formats
                .iter()
                .eq(self.before.iter())
        {
            return Err("The workbook changed while this rule was open. Copy your draft, then reopen the rule before saving.".into());
        }
        Ok(index)
    }
    pub fn build(&self, predicate: &str, style: CondStyle) -> Result<CondFormatStore, String> {
        let mut store = self.before.clone();
        if let Some((pos, old)) = &self.editing {
            store.remove(old.id);
            let mut rule = CondFormatRule::new(old.id, self.ranges.clone(), predicate, style);
            rule.enabled = old.enabled;
            store.insert_at(*pos, rule);
        } else {
            for range in &self.ranges {
                let formula = rebase(predicate, self.anchor, range)?;
                store.add(vec![range.clone()], formula, style.clone());
            }
        }
        Ok(store)
    }
}
fn rebase(predicate: &str, anchor: (usize, usize), range: &CellRange) -> Result<String, String> {
    let shifted = adjust_formula_refs(
        predicate,
        range.start_row as i32 - anchor.0 as i32,
        range.start_col as i32 - anchor.1 as i32,
    );
    // A negative offset can destroy a reference that becomes valid later in
    // the fragment. Refuse that representation rather than freeze it as #REF!.
    if shifted.to_ascii_uppercase().matches("#REF!").count()
        > predicate.to_ascii_uppercase().matches("#REF!").count()
    {
        return Err("A relative reference would move outside the worksheet. Use absolute references or select a smaller range.".into());
    }
    Ok(shifted)
}

fn subtract(r: &CellRange, cut: &CellRange) -> Vec<CellRange> {
    if !r.overlaps(cut) {
        return vec![r.clone()];
    }
    let top = r.start_row.max(cut.start_row);
    let bottom = r.end_row.min(cut.end_row);
    let left = r.start_col.max(cut.start_col);
    let right = r.end_col.min(cut.end_col);
    let mut out = Vec::new();
    if r.start_row < top {
        out.push(CellRange {
            end_row: top - 1,
            ..r.clone()
        });
    }
    if bottom < r.end_row {
        out.push(CellRange {
            start_row: bottom + 1,
            ..r.clone()
        });
    }
    if r.start_col < left {
        out.push(CellRange {
            start_row: top,
            end_row: bottom,
            end_col: left - 1,
            ..r.clone()
        });
    }
    if right < r.end_col {
        out.push(CellRange {
            start_row: top,
            end_row: bottom,
            start_col: right + 1,
            ..r.clone()
        });
    }
    out
}
fn subtract_all(mut pieces: Vec<CellRange>, cuts: &[CellRange]) -> Result<Vec<CellRange>, String> {
    for cut in cuts {
        let mut next = Vec::new();
        for piece in pieces {
            next.extend(subtract(&piece, cut));
            if next.len() > MAX_FRAGMENTS {
                return Err(fragment_error());
            }
        }
        pieces = next;
    }
    Ok(pieces)
}

/// Clear only selected canonical cells; retain hidden records and predicate anchors.
/// Earlier overlapping ranges in a rule own their cells, as in the engine.
pub(crate) fn clear(
    store: &CondFormatStore,
    cuts: &[CellRange],
) -> Result<CondFormatStore, String> {
    let mut allocator = store.clone();
    let mut result = store.clone();
    let mut position = 0;
    let mut fragments = 0;
    for rule in store.iter() {
        if !rule
            .ranges
            .iter()
            .any(|r| cuts.iter().any(|c| r.overlaps(c)))
        {
            position += 1;
            continue;
        }
        result.remove(rule.id);
        let mut first = true;
        for (n, range) in rule.ranges.iter().enumerate() {
            let pieces = subtract_all(vec![range.clone()], &rule.ranges[..n])?;
            for piece in subtract_all(pieces, cuts)? {
                fragments += 1;
                if fragments > MAX_FRAGMENTS {
                    return Err(fragment_error());
                }
                let formula = rebase(&rule.predicate, (range.start_row, range.start_col), &piece)?;
                let id = if first {
                    first = false;
                    rule.id
                } else {
                    allocator.add(Vec::new(), "=TRUE", rule.style.clone())
                };
                let mut fragment =
                    CondFormatRule::new(id, vec![piece], formula, rule.style.clone());
                fragment.enabled = rule.enabled;
                result.insert_at(position, fragment);
                position += 1;
            }
        }
    }
    Ok(result)
}

pub(crate) fn validate_history(
    wb: &Workbook,
    action: &UndoAction,
    forward: bool,
) -> Result<(), String> {
    match action {
        UndoAction::CondFormatAdded { sheet_index, rule } => {
            let sheet = wb
                .sheet(*sheet_index)
                .ok_or("The conditional-format sheet no longer exists.")?;
            if forward {
                validate_rule(sheet, rule)?;
            }
        }
        UndoAction::CondFormatsCleared { sheet_index, rules } => {
            let sheet = wb
                .sheet(*sheet_index)
                .ok_or("The conditional-format sheet no longer exists.")?;
            if !forward {
                for rule in rules {
                    validate_rule(sheet, rule)?;
                }
            }
        }
        UndoAction::Group { actions, .. } => {
            for action in actions {
                validate_history(wb, action, forward)?;
            }
        }
        _ => (),
    }
    Ok(())
}

/// Call after validating the entire history entry. Replays canonical metadata
/// without copying cell storage or evaluating formulas.
pub(crate) fn apply(wb: &mut Workbook, index: usize, rules: &[CondFormatRule], add: bool) {
    let sheet = wb
        .sheet_mut(index)
        .expect("preflighted conditional-format sheet");
    for rule in rules {
        if add {
            let mut rule = rule.clone();
            rule.reparse();
            sheet.cond_formats.insert_at(usize::MAX, rule);
        } else {
            sheet.cond_formats.remove(rule.id);
        }
    }
    wb.bump_revision_for_structure();
}

#[cfg(test)]
#[path = "cond_format_plan_tests.rs"]
mod tests;
