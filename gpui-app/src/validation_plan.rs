//! Private validation drafts and metadata-only history. All ranges are stored
//! worksheet coordinates; screen positions must be resolved before capture.
use crate::history::UndoAction;
use std::collections::HashMap;
use visigrid_engine::{
    sheet::{Sheet, SheetId},
    validation::{CellRange, ValidationEdit, ValidationPatch, ValidationStore},
    workbook::Workbook,
};

#[derive(Clone, Debug)]
pub(crate) struct Draft {
    pub sheet_id: SheetId,
    revision: u64,
    before: ValidationStore,
    pub ranges: Vec<CellRange>,
    pub anchor: (usize, usize),
}

#[derive(Clone, Debug)]
pub(crate) struct Commit {
    pub sheet_id: SheetId,
    pub ranges: Vec<CellRange>,
    patch: ValidationPatch,
}

#[derive(Clone, Debug)]
pub(crate) struct DropdownTarget {
    pub sheet_id: SheetId,
    pub revision: u64,
    pub cell: (usize, usize),
}
impl DropdownTarget {
    pub fn capture(
        wb: &Workbook,
        rows: &visigrid_engine::filter::RowView,
        view: (usize, usize),
    ) -> Self {
        Self {
            sheet_id: wb.active_sheet_id(),
            revision: wb.revision(),
            cell: (rows.view_to_data(view.0), view.1),
        }
    }
    pub fn is_current(
        &self,
        wb: &Workbook,
        rows: &visigrid_engine::filter::RowView,
        selected: (usize, usize),
    ) -> bool {
        self.sheet_id == wb.active_sheet_id()
            && self.revision == wb.revision()
            && rows.data_to_view(self.cell.0) == Some(selected.0)
            && self.cell.1 == selected.1
            && !wb
                .active_sheet()
                .manual_hidden_rows()
                .contains(&self.cell.0)
    }
}

pub(crate) fn validate_targets(sheet: &Sheet, ranges: &[CellRange]) -> Result<(), String> {
    for r in ranges {
        if r.start_row > r.end_row
            || r.start_col > r.end_col
            || r.end_row >= sheet.rows
            || r.end_col >= sheet.cols
        {
            return Err("A validation range is outside the worksheet.".into());
        }
        if let Some(spec) = sheet.table_view_spec().filter(|s| s.has_criteria()) {
            let table = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The validation Table no longer exists.")?;
            let t = table.range;
            if r.start_row <= t.end_row
                && r.end_row > t.start_row
                && (r.start_col < t.start_col || r.end_col > t.end_col)
            {
                return Err("Validation beside this Table would move or hide with its records. Select cells inside the Table, or clear its sorting and filters.".into());
            }
        }
    }
    Ok(())
}

impl Draft {
    pub fn anchor_rule(&self) -> Option<visigrid_engine::validation::ValidationRule> {
        // Exclusion disables enforcement, not the underlying definition. The
        // editor must not mistake an excluded imported rule for Any Value.
        self.before
            .iter()
            .find(|(r, _)| r.contains(self.anchor.0, self.anchor.1))
            .map(|(_, r)| r.clone())
    }
    pub fn anchor_excluded(&self) -> bool {
        self.before.is_excluded(self.anchor.0, self.anchor.1)
    }

    pub fn new(
        wb: &Workbook,
        ranges: Vec<CellRange>,
        anchor: (usize, usize),
    ) -> Result<Self, String> {
        wb.ensure_writable()?;
        let sheet = wb.active_sheet();
        if ranges.is_empty() {
            return Err("No visible cells are selected.".into());
        }
        validate_targets(sheet, &ranges)?;
        Ok(Self {
            sheet_id: sheet.id,
            revision: wb.revision(),
            before: sheet.validations.clone(),
            ranges,
            anchor,
        })
    }
    pub fn prepare(&self, wb: &Workbook, edit: ValidationEdit) -> Result<Option<Commit>, String> {
        wb.ensure_writable()?;
        let sheet = wb
            .sheet_by_id(self.sheet_id)
            .ok_or("The validation sheet no longer exists.")?;
        if wb.revision() != self.revision
            || wb.active_sheet_id() != self.sheet_id
            || sheet.validations != self.before
        {
            return Err("The workbook changed while this rule was open. Copy your draft, then reopen the rule before saving.".into());
        }
        validate_targets(sheet, &self.ranges)?;
        let patch = self
            .before
            .plan_edit(&self.ranges, edit, sheet.rows, sheet.cols)?;
        Ok((!patch.is_empty()).then(|| Commit {
            sheet_id: self.sheet_id,
            ranges: self.ranges.clone(),
            patch,
        }))
    }
}

impl Commit {
    pub fn apply(&self, wb: &mut Workbook, forward: bool) -> Result<(), String> {
        wb.ensure_writable()?;
        let sheet = wb
            .sheet_by_id_mut(self.sheet_id)
            .ok_or("The validation sheet no longer exists.")?;
        validate_targets(sheet, &self.ranges)?;
        // Historical fragments may extend beyond the selected cells. Only
        // bounds matter for those: their effective unselected rules don't change.
        for r in self.patch.changed_ranges() {
            if r.end_row >= sheet.rows || r.end_col >= sheet.cols {
                return Err("A historical validation range is outside the worksheet.".into());
            }
        }
        self.patch.apply(&mut sheet.validations, forward)?;
        wb.bump_revision_for_structure();
        Ok(())
    }
}

/// Simulate every validation change in replay order before a group's first
/// mutation. Other metadata does not touch validation; mixed structural/value
/// groups need their own transaction and are refused rather than partially run.
pub(crate) fn validate_history(
    wb: &Workbook,
    action: &UndoAction,
    forward: bool,
) -> Result<(), String> {
    fn contains(action: &UndoAction) -> bool {
        match action {
            UndoAction::ValidationChanged { .. } => true,
            UndoAction::Group { actions, .. } => actions.iter().any(contains),
            _ => false,
        }
    }
    if !contains(action) {
        return Ok(());
    }
    wb.ensure_writable()?;
    fn walk(
        wb: &Workbook,
        action: &UndoAction,
        forward: bool,
        stores: &mut HashMap<SheetId, ValidationStore>,
    ) -> Result<(), String> {
        match action {
            UndoAction::ValidationChanged { commit, .. } => {
                let sheet = wb.sheet_by_id(commit.sheet_id).ok_or("The validation sheet no longer exists.")?;
                validate_targets(sheet, &commit.ranges)?;
                if commit.patch.changed_ranges().any(|r| r.end_row >= sheet.rows || r.end_col >= sheet.cols) {
                    return Err("A historical validation range is outside the worksheet.".into());
                }
                let store = stores.entry(sheet.id).or_insert_with(|| sheet.validations.clone());
                commit.patch.apply(store, forward)?;
            }
            UndoAction::Group { actions, .. } => {
                if forward { for a in actions { walk(wb, a, forward, stores)?; } }
                else { for a in actions.iter().rev() { walk(wb, a, forward, stores)?; } }
            }
            UndoAction::Format { .. } | UndoAction::Comments { .. } | UndoAction::CondFormatAdded { .. }
            | UndoAction::CondFormatsCleared { .. } | UndoAction::FreezePanesChanged { .. } => {}
            _ => return Err("This mixed validation history group cannot be replayed safely. Nothing was applied.".into()),
        }
        Ok(())
    }
    walk(wb, action, forward, &mut HashMap::new())
}

pub(crate) fn range_summary(ranges: &[CellRange]) -> String {
    let labels: Vec<_> = ranges
        .iter()
        .take(3)
        .map(|r| {
            let cell =
                |row, col| format!("{}{}", crate::app::Spreadsheet::col_letter(col), row + 1);
            let start = cell(r.start_row, r.start_col);
            if r.start_row == r.end_row && r.start_col == r.end_col {
                start
            } else {
                format!("{start}:{}", cell(r.end_row, r.end_col))
            }
        })
        .collect();
    let mut text = labels.join(", ");
    if ranges.len() > 3 {
        text.push_str(&format!(" (+{} ranges)", ranges.len() - 3));
    }
    text
}

pub(crate) fn failure_target(
    rows: &visigrid_engine::filter::RowView,
    hidden_rows: Option<&std::collections::BTreeSet<usize>>,
    hidden_cols: Option<&std::collections::BTreeSet<usize>>,
    failures: &[(usize, usize)],
    current: (usize, usize),
    backward: bool,
) -> Option<(usize, (usize, usize), usize, usize)> {
    let mut visible: Vec<_> = failures
        .iter()
        .enumerate()
        .filter_map(|(i, &(row, col))| {
            if hidden_rows.is_some_and(|h| h.contains(&row))
                || hidden_cols.is_some_and(|h| h.contains(&col))
            {
                return None;
            }
            rows.data_to_view(row).map(|r| (i, (r, col)))
        })
        .collect();
    visible.sort_unstable_by_key(|(_, pos)| *pos);
    if visible.is_empty() {
        return None;
    }
    let rank = if backward {
        visible
            .iter()
            .rposition(|(_, pos)| *pos < current)
            .unwrap_or(visible.len() - 1)
    } else {
        visible
            .iter()
            .position(|(_, pos)| *pos > current)
            .unwrap_or(0)
    };
    Some((visible[rank].0, visible[rank].1, rank + 1, visible.len()))
}

#[cfg(test)]
#[path = "validation_plan_tests.rs"]
mod tests;
