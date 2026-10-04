//! Explicit canonical-bound resize with saved Table criteria kept intact.
use crate::app::Spreadsheet;
use gpui::Context;
use visigrid_engine::{
    table::{TableId, TableRange},
    workbook::{TableCommit, Workbook},
};

fn validate_views(wb: &Workbook) -> Result<(), String> {
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(())
}

// Only bound changes pass this history gate; creation, conversion and formula
// rule changes continue to use their existing command restrictions.
pub(crate) fn is_resize(commit: &TableCommit) -> bool {
    let (Some(before), Some(after)) = (commit.before_table(), commit.after_table()) else {
        return false;
    };
    commit.inserted_header_row().is_none()
        && before.range != after.range
        && before.range.start_row == after.range.start_row
        && before.range.start_col == after.range.start_col
        && before.name == after.name
        && before.style == after.style
}

fn check_removed_criteria(wb: &Workbook, id: TableId, range: TableRange) -> Result<(), String> {
    let (sheet_id, table) = wb.table(id).ok_or("The Table no longer exists.")?;
    if let Some(spec) = wb
        .sheet_by_id(sheet_id)
        .unwrap()
        .table_view_spec()
        .filter(|s| s.table == id)
    {
        for (offset, column) in table.columns.iter().enumerate() {
            if table.range.start_col + offset > range.end_col
                && (spec.sort.as_ref().is_some_and(|s| s.column == column.id)
                    || spec.filters.iter().any(|f| f.column == column.id))
            {
                return Err(format!(
                    "Clear the sort or filter on {} before removing that column from the Table.",
                    column.name
                ));
            }
        }
    }
    Ok(())
}

fn finish_candidate(mut candidate: Workbook) -> Result<Workbook, String> {
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!(
            "Could not recalculate the resized Table: {error:?}"
        ));
    }
    validate_views(&candidate)?;
    Ok(candidate)
}

fn prepare_resize(
    wb: &Workbook,
    id: TableId,
    range: TableRange,
) -> Result<Option<(Workbook, TableCommit)>, String> {
    wb.ensure_writable()?;
    validate_views(wb)?;
    let (_, table) = wb.table(id).ok_or("The Table no longer exists.")?;
    if range == table.range {
        return Ok(None);
    }
    check_removed_criteria(wb, id, range)?;
    let mut candidate = wb.clone();
    let commit = candidate.resize_table(id, range)?;
    Ok(Some((finish_candidate(candidate)?, commit)))
}

pub(crate) fn prepare_resize_replay(
    wb: &Workbook,
    commit: &TableCommit,
    undo: bool,
) -> Result<Workbook, String> {
    wb.ensure_writable()?;
    if !is_resize(commit) {
        return Err("This history entry is not a Table resize.".into());
    }
    validate_views(wb)?;
    let target = if undo {
        commit.before_table()
    } else {
        commit.after_table()
    }
    .unwrap();
    check_removed_criteria(wb, commit.table_id(), target.range)?;
    let mut candidate = wb.clone();
    candidate.apply_table_commit(commit, undo)?;
    finish_candidate(candidate)
}

impl Spreadsheet {
    pub(crate) fn submit_table_resize(
        &mut self,
        id: TableId,
        range: TableRange,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.sync_table_view(cx);
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            return Err(
                "Clear this sheet's worksheet sorting and filters before resizing a Table.".into(),
            );
        }
        self.validate_saved_view_layout(self.wb(cx))?;
        let Some((candidate, commit)) = prepare_resize(self.wb(cx), id, range)? else {
            return Ok(());
        };
        self.validate_saved_view_layout(&candidate)?;
        let name = commit.after_table().unwrap().name.clone();
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.table_filter_dropdown = None;
        self.sync_table_view(cx);
        // A resize edits membership rather than a record: anchor at the header,
        // which stays visible even when the body shrinks to zero records.
        self.view_state
            .select_cell(range.start_row, range.start_col);
        self.view_state.additional_selections.clear();
        self.ensure_visible(cx);
        self.record_table_commit(commit, format!("Resize Table: {name}"), cx);
        Ok(())
    }

    pub(crate) fn replay_table_resize(
        &mut self,
        commit: &TableCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = prepare_resize_replay(self.wb(cx), commit, undo).and_then(|candidate| {
            self.validate_saved_view_layout(&candidate)?;
            Ok(candidate)
        });
        match result {
            Ok(candidate) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.bump_cells_rev();
                self.is_modified = true;
                cx.notify();
                true
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                false
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        history::{History, UndoAction},
        table_edit::tests::fixture,
    };
    use visigrid_engine::{cell::CellComment, sheet::SheetId};

    fn resize(wb: &Workbook, row: usize, col: usize) -> Result<(Workbook, TableCommit), String> {
        let table = &wb.active_sheet().tables()[0];
        prepare_resize(
            wb,
            table.id,
            TableRange {
                end_row: row,
                end_col: col,
                ..table.range
            },
        )?
        .ok_or("Unexpected no-op".into())
    }

    #[test]
    fn resize_moves_totals_with_saved_criteria_and_replays() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        before.set_cell_value_tracked(0, 8, 1, "West");
        before.set_cell_value_tracked(0, 8, 3, "25");
        let (after, commit) = resize(&before, 8, 3).unwrap();
        assert_eq!(after.table(id).unwrap().1.totals_row(), Some(9));
        assert_eq!(after.active_sheet().get_display(9, 3), "205");
        assert_eq!(after.active_sheet().get_raw(7, 3), "");
        assert_eq!(after.active_sheet().table_view_spec(), before.active_sheet().table_view_spec());
        let restored = prepare_resize_replay(&after, &commit, true).unwrap();
        assert_eq!(restored.table(id).unwrap().1.totals_row(), Some(7));
        assert_eq!(restored.active_sheet().get_raw(8, 3), "25");
        let redone = prepare_resize_replay(&restored, &commit, false).unwrap();
        assert_eq!(redone.active_sheet().get_display(9, 3), "205");
    }

    #[test]
    fn includes_existing_records_without_filling_or_touching_hidden_overrides() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        before.set_cell_value_tracked(0, 4, 3, "999");
        before.set_cell_value_tracked(0, 7, 1, "West");
        before.set_cell_value_tracked(0, 7, 2, "5");
        before.active_sheet_mut().set_comment(
            7,
            2,
            Some(CellComment {
                text: "existing note".into(),
                author: "".into(),
            }),
        );
        before.set_cell_value_tracked(0, 8, 1, "East");
        before.set_cell_value_tracked(0, 8, 2, "70");
        let (after, commit) = resize(&before, 8, 3).unwrap();
        assert!(is_resize(&commit));
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().get_raw(7, 3), "");
        assert_eq!(
            after.active_sheet().comment(7, 2),
            before.active_sheet().comment(7, 2)
        );
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(view.rows().view_to_data(3), 7);
        assert!(view.rows().data_to_view(8).is_none());
        let undone = prepare_resize_replay(&after, &commit, true).unwrap();
        assert_eq!(
            undone.table(id).unwrap().1.range,
            before.table(id).unwrap().1.range
        );
        assert_eq!(undone.active_sheet().get_raw(7, 2), "5");
        let redone = prepare_resize_replay(&undone, &commit, false).unwrap();
        assert_eq!(redone.table(id).unwrap().1, after.table(id).unwrap().1);
    }

    #[test]
    fn shrink_releases_rows_and_can_leave_only_headers_with_criteria() {
        let before = fixture(true);
        for end in [4, 2] {
            let (after, commit) = resize(&before, end, 3).unwrap();
            assert_eq!(after.active_sheet().get_raw(6, 2), "40");
            assert_eq!(
                after.active_sheet().table_view_spec(),
                before.active_sheet().table_view_spec()
            );
            let undone = prepare_resize_replay(&after, &commit, true).unwrap();
            assert_eq!(
                undone.active_sheet().tables(),
                before.active_sheet().tables()
            );
            assert_eq!(undone.active_sheet().get_raw(4, 1), "East");
        }
    }

    #[test]
    fn width_growth_normalizes_headers_and_preserves_column_ids_and_formulas() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 2, 4, "Amount");
        before.set_cell_value_tracked(0, 7, 4, "existing");
        let (after, commit) = resize(&before, 7, 4).unwrap();
        let old = &before.active_sheet().tables()[0];
        let new = &after.active_sheet().tables()[0];
        assert_eq!(&new.columns[..3], &old.columns);
        assert_eq!(new.columns[3].name, "Amount2");
        assert_eq!(after.active_sheet().get_raw(7, 4), "existing");
        let undone = prepare_resize_replay(&after, &commit, true).unwrap();
        assert_eq!(undone.active_sheet().get_raw(2, 4), "Amount");
        assert_eq!(
            prepare_resize_replay(&undone, &commit, false)
                .unwrap()
                .active_sheet()
                .tables(),
            after.active_sheet().tables()
        );
    }

    #[test]
    fn removing_unused_column_rewrites_dependents_and_undo_restores_them() {
        let mut before = fixture(true);
        for row in 3..=6 {
            before.set_cell_value_tracked(0, row, 3, "");
        }
        let other = before.add_sheet();
        before.set_cell_value_tracked(other, 0, 0, "=SUM(Sales[Result])");
        let (after, commit) = resize(&before, 6, 2).unwrap();
        assert!(after.sheet(other).unwrap().get_raw(0, 0).contains("#REF!"));
        assert_eq!(after.active_sheet().get_raw(3, 3), "");
        let undone = prepare_resize_replay(&after, &commit, true).unwrap();
        assert_eq!(
            undone.sheet(other).unwrap().get_raw(0, 0),
            "=SUM(Sales[Result])"
        );
    }

    #[test]
    fn removing_sort_or_filter_column_refuses_atomically() {
        for filtered in [false, true] {
            let mut before = fixture(filtered);
            if filtered {
                let mut spec = before.active_sheet().table_view_spec().unwrap().clone();
                spec.sort = None;
                spec.filters[0].column = before.active_sheet().tables()[0].columns[2].id;
                before.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
            }
            let revision = before.revision();
            let error = resize(&before, 6, if filtered { 2 } else { 1 }).unwrap_err();
            assert!(error.contains("Clear the sort or filter"), "{error}");
            assert_eq!(before.revision(), revision);
            assert_eq!(before.active_sheet().tables()[0].columns.len(), 3);
        }
    }

    #[test]
    fn refuses_unsafe_new_layout_and_invalid_ranges_without_mutating() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 7, 0, "neighbor");
        let rev = before.revision();
        assert!(resize(&before, 7, 3).is_err());
        // Released populated columns would sit beside the remaining projection.
        assert!(resize(&before, 6, 2).is_err());
        assert!(resize(&before, 30, 3).is_err());
        let table = &before.active_sheet().tables()[0];
        assert!(prepare_resize(
            &before,
            table.id,
            TableRange {
                start_row: 1,
                ..table.range
            }
        )
        .is_err());
        assert_eq!(before.revision(), rev);
        assert_eq!(before.active_sheet().get_raw(7, 0), "neighbor");
    }

    #[test]
    fn dependent_spill_refuses_resize_and_replay() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 0, 0, "=IF(ROWS(Sales[Amount])>4,SEQUENCE(5),0)");
        let rev = before.revision();
        assert!(resize(&before, 7, 3).is_err());
        assert_eq!(before.revision(), rev);
        before.set_cell_value_tracked(0, 0, 0, "");
        let (mut after, commit) = resize(&before, 7, 3).unwrap();
        after.set_cell_value_tracked(0, 0, 0, "=IF(ROWS(Sales[Amount])<5,SEQUENCE(5),0)");
        let rev = after.revision();
        assert!(prepare_resize_replay(&after, &commit, true).is_err());
        assert_eq!(after.revision(), rev);
        assert_eq!(after.active_sheet().tables()[0].range.end_row, 7);
    }

    #[test]
    fn no_op_recovery_and_stale_history_guards() {
        let mut before = fixture(true);
        let table = before.active_sheet().tables()[0].clone();
        assert!(prepare_resize(&before, table.id, table.range)
            .unwrap()
            .is_none());
        let (mut after, commit) = resize(&before, 7, 3).unwrap();
        after.rename_table(table.id, "Changed").unwrap();
        assert!(prepare_resize_replay(&after, &commit, true).is_err());
        before.active_sheet_mut().read_only_reason = Some("Damaged metadata".into());
        assert!(resize(&before, 7, 3).unwrap_err().contains("Read-only"));
    }

    #[test]
    fn rewind_and_native_roundtrip_keep_bounds_and_criteria() {
        let before = fixture(true);
        let (after, commit) = resize(&before, 7, 4).unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableCommit { header_layout: None,
                sheet_index: 0,
                commit: Box::new(commit),
                description: "Resize Table".into(),
            },
            None,
        );
        assert_eq!(history.undo_count(), 1);
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(
            preview.workbook.active_sheet().tables(),
            after.active_sheet().tables()
        );
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("resize.sheet");
        visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
            .unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(
            loaded.active_sheet().tables(),
            after.active_sheet().tables()
        );
        assert_eq!(
            loaded.active_sheet().table_view_spec(),
            after.active_sheet().table_view_spec()
        );
    }
    #[test]
    fn validates_other_sheets_after_recalculation() {
        let mut before = fixture(true);
        let other = before.add_sheet_named("Other").unwrap();
        let sheet_id = before.sheet(other).unwrap().id;
        before.set_cell_value_tracked(other, 2, 1, "Value");
        before.set_cell_value_tracked(other, 3, 1, "10");
        before.set_cell_value_tracked(other, 4, 1, "20");
        let id = before
            .create_table(
                sheet_id,
                TableRange {
                    start_row: 2,
                    end_row: 4,
                    start_col: 1,
                    end_col: 1,
                },
                "OtherTable",
            )
            .unwrap()
            .table_id();
        let mut spec = visigrid_engine::table_view::TableViewSpec::new(id);
        spec.sort = Some(visigrid_engine::table_view::TableSort {
            column: before.table(id).unwrap().1.columns[0].id,
            direction: visigrid_engine::filter::SortDirection::Ascending,
        });
        before.set_table_view_spec(sheet_id, Some(spec)).unwrap();
        before.set_cell_value_tracked(other, 0, 0, "=IF(ROWS(Sales[Amount])>4,SEQUENCE(5),0)");
        validate_views(&before).unwrap();
        assert!(resize(&before, 7, 3).is_err());
        assert_eq!(before.sheet(other).unwrap().get_display(0, 0), "0");
        // The other Table can itself be resized without clearing the first view.
        let (after, commit) = prepare_resize(
            &before,
            id,
            TableRange {
                start_row: 2,
                end_row: 5,
                start_col: 1,
                end_col: 1,
            },
        )
        .unwrap()
        .unwrap();
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        assert!(prepare_resize_replay(&after, &commit, true).is_ok());
    }

    #[test]
    fn undo_cannot_remove_a_column_now_used_by_criteria() {
        let before = fixture(true);
        let (mut after, commit) = resize(&before, 6, 4).unwrap();
        let mut spec = after.active_sheet().table_view_spec().unwrap().clone();
        spec.sort.as_mut().unwrap().column = after.active_sheet().tables()[0].columns[3].id;
        after.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        let rev = after.revision();
        assert!(prepare_resize_replay(&after, &commit, true)
            .unwrap_err()
            .contains("Clear the sort or filter"));
        assert_eq!(after.revision(), rev);
    }

    #[test]
    fn resized_bounds_are_checked_against_desktop_row_presentation() {
        let before = fixture(true);
        let (after, _) = resize(&before, 8, 3).unwrap();
        let table = &after.active_sheet().tables()[0];
        let heights = [(8, 35.0)].into();
        let hidden = [8].into();
        assert!(
            crate::table_filter_ui::desktop_layout_error(table, Some(&heights), None, 0).is_some()
        );
        assert!(
            crate::table_filter_ui::desktop_layout_error(table, None, Some(&hidden), 0).is_some()
        );
        assert!(crate::table_filter_ui::desktop_layout_error(table, None, None, 8).is_some());
        assert!(crate::table_filter_ui::desktop_layout_error(table, None, None, 3).is_none());
    }
    #[test]
    fn totals_width_resize_preserves_criteria_and_rewinds_after_total_edit_undo() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_cell_value_tracked(0, 2, 4, "Extra");
        let spec = base.active_sheet().table_view_spec().cloned();
        let (mut after, commit) = resize(&base, 6, 4).unwrap();
        assert!(is_resize(&commit));
        assert_eq!(after.active_sheet().tables()[0].totals.as_ref().unwrap().columns.len(), 4);
        assert_eq!(after.active_sheet().get_display(7, 3), "180");
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        let total = after.set_table_total(id, 4, visigrid_engine::table::TableTotal {
            function: Some("sum".into()), ..Default::default()
        }).unwrap();
        after.apply_table_commit(&total, true).unwrap();
        let undone = prepare_resize_replay(&after, &commit, true).unwrap();
        assert_eq!(undone.active_sheet().get_display(7, 3), "180");
        assert_eq!(undone.active_sheet().tables()[0].columns.len(), 3);
        let redone = prepare_resize_replay(&undone, &commit, false).unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(UndoAction::TableCommit {
            header_layout: None, sheet_index: 0, commit: Box::new(commit), description: "Widen Table".into(),
        }, None);
        let preview = history.build_workbook_before(1, Some(&base), 100, 10_000).unwrap();
        assert_eq!(preview.workbook.active_sheet().tables(), redone.active_sheet().tables());
        assert_eq!(preview.workbook.active_sheet().get_display(7, 3), "180");
        let (shrunk, _) = resize(&redone, 6, 3).unwrap();
        assert_eq!(shrunk.active_sheet().table_view_spec(), spec.as_ref());
        assert_eq!(shrunk.active_sheet().tables()[0].columns.len(), 3);
        assert!(resize(&redone, 6, 1).is_err()); // Amount is the sort field.
    }

    #[test]
    fn totals_width_resize_refuses_occupied_footer_and_unsafe_released_records() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_cell_value_tracked(0, 7, 4, "Keep note");
        let revision = base.revision();
        assert!(resize(&base, 6, 4).is_err());
        assert_eq!(base.revision(), revision);
        assert_eq!(base.active_sheet().get_raw(7, 4), "Keep note");
        assert!(resize(&base, 6, 2).is_err()); // Released Result cells would sit beside a projected Table.
        assert_eq!(base.active_sheet().get_display(7, 3), "180");
    }

}
