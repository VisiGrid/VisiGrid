//! Explicit column rules and sparse history through saved Table views.
use crate::app::Spreadsheet;
use gpui::Context;
use visigrid_engine::{
    sheet::SheetId,
    table::TableId,
    workbook::{TableCommit, Workbook},
};

fn finish(mut candidate: Workbook) -> Result<Workbook, String> {
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!("Could not recalculate the column formula: {error:?}"));
    }
    for sheet in candidate.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(candidate)
}

fn prepare_rule(
    wb: &Workbook, id: TableId, col: usize, row: usize, source: &str, replace: bool,
) -> Result<(Workbook, TableCommit), String> {
    let mut candidate = wb.clone();
    let commit = candidate.set_calculated_column(id, col, row, source, replace)?;
    Ok((finish(candidate)?, commit))
}

pub(crate) fn prepare_restore(
    wb: &Workbook, id: TableId, row: usize, col: usize,
) -> Result<(Workbook, TableCommit), String> {
    let mut candidate = wb.clone();
    let commit = candidate.restore_calculated_cell(id, row, col)?;
    Ok((finish(candidate)?, commit))
}

pub(crate) fn prepare_inferred(
    wb: &Workbook, sheet: SheetId, row: usize, col: usize, source: &str,
) -> Result<Option<(Workbook, TableCommit)>, String> {
    if !source.starts_with('=') { return Ok(None); }
    let sheet = wb.sheet_by_id(sheet).ok_or("Sheet no longer exists.")?;
    let Some(table) = sheet.table_at(row, col) else { return Ok(None); };
    if row <= table.range.start_row || row > table.range.end_row
        || table.columns[col - table.range.start_col].formula.is_some()
        || (table.range.start_row + 1..=table.range.end_row)
            .any(|r| r != row && !sheet.get_raw(r, col).is_empty()) {
        return Ok(None);
    }
    prepare_rule(wb, table.id, col, row, source, true).map(Some)
}

pub(crate) fn prepare_replay(
    wb: &Workbook, commit: &TableCommit, undo: bool,
) -> Result<Workbook, String> {
    if !commit.is_calculated_change() {
        return Err("This history entry is not a calculated-column change.".into());
    }
    let mut candidate = wb.clone();
    candidate.apply_table_commit(commit, undo)?;
    finish(candidate)
}

impl Spreadsheet {
    pub(crate) fn publish_calculated(
        &mut self, candidate: Workbook, commit: TableCommit, description: &str,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.sync_table_view(cx);
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered()) {
            return Err("Clear worksheet sorting and filters before changing a column formula.".into());
        }
        self.validate_saved_view_layout(self.wb(cx))?;
        self.validate_saved_view_layout(&candidate)?;
        self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.table_filter_dropdown = None;
        self.sync_table_view(cx);
        self.record_table_commit(commit, description.into(), cx);
        Ok(())
    }

    pub(crate) fn submit_column_formula(
        &mut self, id: TableId, col: usize, row: usize, source: &str, replace: bool,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        let (candidate, commit) = prepare_rule(self.wb(cx), id, col, row, source, replace)?;
        self.publish_calculated(candidate, commit,
            if replace { "Replace column formulas" } else { "Edit column formula" }, cx)
    }

    pub(crate) fn replay_calculated(
        &mut self, commit: &TableCommit, undo: bool, cx: &mut Context<Self>,
    ) -> bool {
        let result = prepare_replay(self.wb(cx), commit, undo).and_then(|candidate| {
            self.validate_saved_view_layout(&candidate)?;
            Ok(candidate)
        });
        match result {
            Ok(candidate) => {
                self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.bump_cells_rev();
                self.is_modified = true;
                cx.notify();
                true
            }
            Err(error) => { self.status_message = Some(error); cx.notify(); false }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{history::{History, UndoAction}, table_edit::tests::fixture};

    #[test]
    fn rules_with_totals_preserve_hidden_overrides_and_replay_through_views() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_calculated_column(id, 3, 3, "=[@Amount]*2", true).unwrap();
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_cell_value_tracked(0, 4, 3, "999");
        base.clear_cell_tracked(0, 5, 3);
        let view = base.active_sheet().table_view_spec().cloned();
        let (after, update) = prepare_rule(&base, id, 3, 3, "=[@Amount]*3", false).unwrap();
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().get_raw(5, 3), "");
        assert_eq!(after.active_sheet().get_display(7, 3), "210");
        assert_eq!(after.active_sheet().table_view_spec(), view.as_ref());
        let (replaced, replace) = prepare_rule(&after, id, 3, 3, "=[@Amount]*3", true).unwrap();
        assert_eq!(replaced.active_sheet().get_display(4, 3), "30");
        assert_eq!(replaced.active_sheet().get_display(7, 3), "270");
        assert_eq!(replaced.active_sheet().get_raw(7, 3), "=SUBTOTAL(109,[Result])");
        let undone = prepare_replay(&replaced, &replace, true).unwrap();
        let undone = prepare_replay(&undone, &update, true).unwrap();
        assert_eq!(undone.active_sheet().tables(), base.active_sheet().tables());
        assert_eq!(undone.active_sheet().get_display(7, 3), "140");
        let redone = prepare_replay(&undone, &update, false).unwrap();
        let redone = prepare_replay(&redone, &replace, false).unwrap();
        assert_eq!(redone.active_sheet().get_display(7, 3), "270");
        let mut history = History::new();
        for commit in [update, replace] {
            history.record_action_with_provenance(UndoAction::TableCommit {
                sheet_index: 0, commit: Box::new(commit), header_layout: None,
                description: "Column formula".into(),
            }, None);
        }
        let preview = history.build_workbook_before(2, Some(&base), 100, 10_000).unwrap();
        assert_eq!(preview.workbook.active_sheet().get_display(7, 3), "270");
        assert_eq!(preview.workbook.active_sheet().table_view_spec(), view.as_ref());
        let earlier = history.build_workbook_before(1, Some(&base), 100, 10_000).unwrap();
        assert_eq!(earlier.workbook.active_sheet().get_display(7, 3), "210");
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("calculated-totals.sheet");
        visigrid_io::native::save_workbook(&replaced, &path).unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(loaded.active_sheet().tables(), replaced.active_sheet().tables());
        assert_eq!(loaded.active_sheet().get_display(7, 3), "270");
    }

    #[test]
    fn rule_candidates_and_replay_refuse_new_spills_without_publishing() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_calculated_column(id, 3, 3, "=[@Amount]*2", true).unwrap();
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IF(D4=60,1,5))");
        let revision = base.revision();
        assert!(prepare_rule(&base, id, 3, 3, "=[@Amount]*3", false).is_err());
        assert_eq!(base.revision(), revision);
        assert_eq!(base.active_sheet().get_display(7, 3), "180");
        base.set_cell_value_tracked(0, 0, 0, "");
        let (mut after, commit) = prepare_rule(&base, id, 3, 3, "=[@Amount]*3", false).unwrap();
        after.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IF(D4=90,1,5))");
        let revision = after.revision();
        assert!(prepare_replay(&after, &commit, true).is_err());
        assert_eq!(after.revision(), revision);
        assert_eq!(after.active_sheet().get_display(7, 3), "270");
    }

    #[test]
    fn inferred_rule_and_restore_exclude_footer_and_reject_stale_cells() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        for row in 3..=6 { base.clear_cell_tracked(0, row, 3); }
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        assert!(prepare_inferred(&base, SheetId(7), 7, 3, "=99").unwrap().is_none());
        let (mut after, _) = prepare_inferred(&base, SheetId(7), 3, 3, "=[@Amount]*2").unwrap().unwrap();
        assert_eq!(after.active_sheet().get_display(7, 3), "180");
        after.clear_cell_tracked(0, 5, 3);
        let (mut restored, restore) = prepare_restore(&after, id, 5, 3).unwrap();
        assert_eq!(restored.active_sheet().get_display(7, 3), "180");
        assert!(prepare_restore(&after, id, 7, 3).is_err());
        restored.set_cell_value_tracked(0, 5, 3, "999");
        assert!(prepare_replay(&restored, &restore, true).is_err());
        assert_eq!(restored.active_sheet().get_raw(5, 3), "999");
    }
}
