//! Rewrite rule sources without enumerating their target cells. Local binding
//! and this-row validity are constant inside each Table-boundary rectangle.
use super::Workbook;
use crate::{
    cond_format::{CondFormatRule, CondFormatStore},
    formula::{
        parser::{self, Expr},
        structured::{self, TableSection},
    },
    sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS},
    table::DataTable,
    validation::{CellRange, ValidationRule, ValidationStore},
};
use std::collections::BTreeSet;

const MAX_FRAGMENTS: usize = 4096;
struct Budget {
    steps: usize,
    fragments: usize,
    bytes: usize,
}
impl Budget {
    fn step(&mut self, n: usize) -> Result<(), String> {
        self.steps = self.steps.checked_sub(n).ok_or_else(limit)?;
        Ok(())
    }
    fn fragment(&mut self, bytes: usize) -> Result<(), String> {
        self.fragments = self.fragments.checked_sub(1).ok_or_else(limit)?;
        self.bytes = self.bytes.checked_sub(bytes).ok_or_else(limit)?;
        Ok(())
    }
}
fn limit() -> String {
    "Converting this Table would create too much formatting or validation metadata. Simplify its rules first. Nothing was changed.".into()
}
fn without(
    mut pieces: Vec<CellRange>,
    cuts: &[CellRange],
    budget: &mut Budget,
) -> Result<Vec<CellRange>, String> {
    for cut in cuts {
        budget.step(pieces.len())?;
        let mut next = Vec::new();
        for r in pieces {
            if !r.overlaps(cut) {
                next.push(r);
                continue;
            }
            let (top, bottom, left, right) = (
                r.start_row.max(cut.start_row),
                r.end_row.min(cut.end_row),
                r.start_col.max(cut.start_col),
                r.end_col.min(cut.end_col),
            );
            if r.start_row < top {
                next.push(CellRange::new(r.start_row, r.start_col, top - 1, r.end_col));
            }
            if bottom < r.end_row {
                next.push(CellRange::new(
                    bottom + 1,
                    r.start_col,
                    r.end_row,
                    r.end_col,
                ));
            }
            if r.start_col < left {
                next.push(CellRange::new(top, r.start_col, bottom, left - 1));
            }
            if right < r.end_col {
                next.push(CellRange::new(top, right + 1, bottom, r.end_col));
            }
            if next.len() > MAX_FRAGMENTS {
                return Err(limit());
            }
        }
        pieces = next;
        if pieces.is_empty() {
            break;
        }
    }
    Ok(pieces)
}
fn sources(rule: &ValidationRule) -> Vec<String> {
    let mut sources = Vec::new();
    rule.clone().map_references(|s| {
        sources.push(s.to_owned());
        s.to_owned()
    });
    sources
}
struct Conversion<'a> {
    wb: &'a Workbook,
    owner: SheetId,
    table: &'a DataTable,
}
impl Conversion<'_> {
    fn touches(&self, source: &str, sheet: SheetId, range: CellRange) -> bool {
        let full = self.table.full_range();
        structured::source_references(source)
            .iter()
            .any(|(_, _, r)| match &r.table {
                Some(name) => name.eq_ignore_ascii_case(&self.table.name),
                None => {
                    sheet == self.owner
                        && range.overlaps(&CellRange::new(
                            full.start_row,
                            full.start_col,
                            full.end_row,
                            full.end_col,
                        ))
                }
            })
    }
    fn pieces(&self, range: CellRange, sheet: SheetId) -> Vec<CellRange> {
        if sheet != self.owner {
            return vec![range];
        }
        let full = self.table.full_range();
        let mut rows = BTreeSet::from([range.start_row, range.end_row + 1]);
        let mut cols = BTreeSet::from([range.start_col, range.end_col + 1]);
        for r in [
            full.start_row,
            full.start_row + 1,
            self.table.range.end_row + 1,
            full.end_row + 1,
        ] {
            if r > range.start_row && r <= range.end_row {
                rows.insert(r);
            }
        }
        for c in [full.start_col, full.end_col + 1] {
            if c > range.start_col && c <= range.end_col {
                cols.insert(c);
            }
        }
        let rows: Vec<_> = rows.into_iter().collect();
        let cols: Vec<_> = cols.into_iter().collect();
        rows.windows(2)
            .flat_map(|r| {
                cols.windows(2)
                    .map(move |c| CellRange::new(r[0], c[0], r[1] - 1, c[1] - 1))
            })
            .collect()
    }
    fn source(
        &self,
        source: &str,
        sheet: SheetId,
        row: usize,
        col: usize,
        anchor_row: usize,
    ) -> Result<String, String> {
        let mut rewritten = source.to_owned();
        for (start, end, reference) in structured::source_references(source).into_iter().rev() {
            let targets = reference.table.as_ref().map_or_else(
                || sheet == self.owner && self.table.full_range().contains(row, col),
                |name| name.eq_ignore_ascii_case(&self.table.name),
            );
            if !targets {
                continue;
            }
            let mut resolved = structured::resolve_region(
                self.table,
                self.owner,
                sheet,
                Some((row, col)),
                &reference,
            );
            if sheet == self.owner {
                match &mut resolved {
                    Expr::CellRef { sheet, .. } | Expr::Range { sheet, .. } => {
                        *sheet = crate::sheet::SheetRef::Current
                    }
                    _ => {}
                }
            }
            // The rule's stored anchor may lie outside its current fragment.
            // Keep this-row references relative to that anchor, while column
            // identities and all other structured rectangles become fixed A1.
            if reference.section == TableSection::ThisRow {
                match &mut resolved {
                    Expr::CellRef { row, row_abs, .. } => {
                        *row = anchor_row;
                        *row_abs = false;
                    }
                    Expr::Range {
                        start_row,
                        end_row,
                        start_row_abs,
                        end_row_abs,
                        ..
                    } => {
                        *start_row = anchor_row;
                        *end_row = anchor_row;
                        *start_row_abs = false;
                        *end_row_abs = false;
                    }
                    _ => {}
                }
            }
            let replacement = match &resolved {
                Expr::CellRef { .. } | Expr::Range { .. } => parser::format_expr(&resolved, |id| self.wb.sheet_by_id(id).map(|s| s.name.clone()))[1..].to_owned(),
                Expr::ReferenceError(error) => error.split_whitespace().next().unwrap_or("#REF!").to_owned(),
                Expr::EmptyRange { .. } => return Err("Cannot convert a referenced empty Table to a range; it has no A1 representation.".into()),
                _ => return Err("A rule's Table reference could not be converted. Nothing was changed.".into()),
            };
            rewritten.replace_range(start..end, &replacement);
        }
        Ok(rewritten)
    }
    fn formats(&self, sheet: &Sheet, budget: &mut Budget) -> Result<CondFormatStore, String> {
        let store = &sheet.cond_formats;
        let mut result = store.clone();
        let mut next_id = None;
        let mut position = 0;
        for rule in store.iter() {
            budget.step(rule.ranges.len())?;
            if !rule
                .ranges
                .iter()
                .any(|r| self.touches(&rule.predicate, sheet.id, *r))
            {
                position += 1;
                continue;
            }
            if next_id.is_none() {
                next_id = Some(store.fragment_id_start()?);
            }
            let mut earlier = Vec::new();
            let mut fragments = Vec::new();
            for range in &rule.ranges {
                for unique in without(vec![*range], &earlier, budget)? {
                    for piece in self.pieces(unique, sheet.id) {
                        let source = parser::adjust_formula_refs(
                            &rule.predicate,
                            (piece.start_row - range.start_row) as i32,
                            (piece.start_col - range.start_col) as i32,
                        );
                        let source = self.source(
                            &source,
                            sheet.id,
                            piece.start_row,
                            piece.start_col,
                            piece.start_row,
                        )?;
                        budget.fragment(source.len())?;
                        fragments.push((piece, source));
                    }
                }
                earlier.push(*range);
            }
            result.remove(rule.id);
            for (index, (range, source)) in fragments.into_iter().enumerate() {
                let id = if index == 0 {
                    rule.id
                } else {
                    let id = next_id.unwrap();
                    next_id = Some(id.checked_add(1).ok_or(
                        "Conditional-format rule IDs are exhausted. Nothing was changed.",
                    )?);
                    id
                };
                let mut converted =
                    CondFormatRule::new(id, vec![range], source, rule.style.clone());
                converted.enabled = rule.enabled;
                result.insert_at(position, converted);
                position += 1;
            }
        }
        Ok(result)
    }
    fn validation(
        &self,
        rule: &ValidationRule,
        sheet: SheetId,
        range: CellRange,
    ) -> Result<ValidationRule, String> {
        if !sources(rule).iter().any(|s| self.touches(s, sheet, range)) {
            return Ok(rule.clone());
        }
        let mut converted = rule.clone();
        let origin = rule
            .reference_origin
            .unwrap_or((range.start_row, range.start_col));
        let mut error = None;
        converted.map_references(|s| {
            // Native validation A1 references are fixed even without '$'.
            // Freeze those before introducing a relative this-row reference.
            let source = if rule.reference_origin.is_none() {
                parser::absolutize_formula_refs(s)
            } else {
                s.to_owned()
            };
            match self.source(&source, sheet, range.start_row, range.start_col, origin.0) {
                Ok(s) => s,
                Err(e) => {
                    error = Some(e);
                    source
                }
            }
        });
        if let Some(error) = error {
            return Err(error);
        }
        converted.reference_origin = Some(origin);
        Ok(converted)
    }
    fn validations(&self, sheet: &Sheet, budget: &mut Budget) -> Result<ValidationStore, String> {
        let store = &sheet.validations;
        let rules: Vec<_> = store.iter().map(|(r, rule)| (*r, rule.clone())).collect();
        let mut affected = BTreeSet::new();
        for (index, (range, rule)) in rules.iter().enumerate() {
            budget.step(1)?;
            if sources(rule)
                .iter()
                .any(|s| self.touches(s, sheet.id, *range))
            {
                affected.insert(index);
            }
        }
        // Splitting must preserve first-rule precedence, even if the new
        // rectangle keys would otherwise sort ahead of an overlapping rule.
        let mut pending: Vec<_> = affected.iter().copied().collect();
        while let Some(i) = pending.pop() {
            for (j, (range, _)) in rules.iter().enumerate() {
                budget.step(1)?;
                if !affected.contains(&j) && rules[i].0.overlaps(range) {
                    affected.insert(j);
                    pending.push(j);
                }
            }
        }
        let mut result = store.clone();
        for &index in &affected {
            result.remove(&rules[index].0);
        }
        let mut earlier = Vec::new();
        for index in affected {
            let (range, rule) = &rules[index];
            for unique in without(vec![*range], &earlier, budget)? {
                for piece in self.pieces(unique, sheet.id) {
                    let converted = self.validation(rule, sheet.id, piece)?;
                    budget.fragment(sources(&converted).iter().map(String::len).sum())?;
                    result.set(piece, converted);
                }
            }
            earlier.push(*range);
        }
        Ok(result)
    }
}
impl Workbook {
    pub(super) fn rewrite_conversion_rules(
        &mut self,
        owner: SheetId,
        table: &DataTable,
    ) -> Result<(), String> {
        let conversion = Conversion {
            wb: self,
            owner,
            table,
        };
        let mut budget = Budget {
            steps: 1_000_000,
            fragments: MAX_FRAGMENTS,
            bytes: 8 * 1024 * 1024,
        };
        let mut changes = Vec::new();
        for sheet in self.sheets() {
            let valid = |r: &CellRange| {
                r.start_row <= r.end_row
                    && r.start_col <= r.end_col
                    && r.end_row < sheet.rows.min(NUM_ROWS)
                    && r.end_col < sheet.cols.min(NUM_COLS)
            };
            if sheet
                .cond_formats
                .iter()
                .flat_map(|r| &r.ranges)
                .any(|r| !valid(r))
                || sheet.validations.iter().any(|(r, rule)| {
                    !valid(r)
                        || rule.reference_origin.is_some_and(|(r, c)| {
                            r >= sheet.rows.min(NUM_ROWS) || c >= sheet.cols.min(NUM_COLS)
                        })
                })
                || sheet.validations.exclusions_iter().any(|r| !valid(r))
            {
                return Err("A formatting or validation range is outside the worksheet. Nothing was changed.".into());
            }
            changes.push((
                sheet.id,
                conversion.formats(sheet, &mut budget)?,
                conversion.validations(sheet, &mut budget)?,
            ));
        }
        for (id, formats, validations) in changes {
            let sheet = self.sheet_by_id_mut(id).unwrap();
            sheet.cond_formats = formats;
            sheet.validations = validations;
        }
        Ok(())
    }
}
