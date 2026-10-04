//! Append through saved Table views without moving neighboring worksheet cells.
//! History retains a schema/value commit plus sparse pending-edit or paste patches.
use crate::{
    app::Spreadsheet,
    history::UndoAction,
    table_cell_history::TableCellsCommit,
    table_edit::{prepare_table_writes, TableCellWrite},
};
use gpui::Context;
use visigrid_engine::{
    filter::RowView,
    table::{TableId, TableRange},
    table_view::TableViewSpec,
    workbook::{TableCommit, Workbook},
};

#[derive(Clone, Debug)]
pub(crate) struct TableAppendHistory {
    pub table: TableCommit,
    edit: Option<TableCellsCommit>,
    view: Option<TableViewSpec>,
    paste: Option<TableCellsCommit>,
}

fn validate_views(wb: &Workbook) -> Result<(), String> {
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(())
}

/// The optional Tab edit uses canonical coordinates resolved before sorting.
/// It changes only that visible cell; existing calculated rules fill the new row.
fn prepare_append(
    wb: &Workbook,
    id: TableId,
    edit: Option<TableCellWrite>,
) -> Result<(Workbook, TableAppendHistory), String> {
    wb.ensure_writable()?;
    validate_views(wb)?;
    let (sheet_id, table) = wb.table(id).ok_or("The Table no longer exists.")?;
    let index = wb
        .sheet_index_by_id(sheet_id)
        .ok_or("The sheet no longer exists.")?;
    let sheet = wb.sheet(index).unwrap();
    if edit
        .as_ref()
        .is_some_and(|w| w.row <= table.range.start_row || !table.range.contains(w.row, w.col))
    {
        return Err("The pending edit is outside the Table body.".into());
    }
    let view = sheet.table_view_spec().cloned();
    let (mut candidate, edit) = if let Some(write) = edit {
        let candidate = prepare_table_writes(wb, index, std::slice::from_ref(&write))?;
        let patch = TableCellsCommit::capture(
            sheet,
            candidate.sheet(index).unwrap(),
            [(write.row, write.col)],
        );
        (candidate, Some(patch))
    } else {
        (wb.clone(), None)
    };
    let table = candidate.append_table_rows(id, 1, &[])?;
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!(
            "Could not recalculate the new Table row: {error:?}"
        ));
    }
    validate_views(&candidate)?;
    Ok((
        candidate,
        TableAppendHistory {
            table,
            edit,
            view,
            paste: None,
        },
    ))
}

impl TableAppendHistory {
    pub(crate) fn with_appended_writes(
        table: TableCommit,
        paste: TableCellsCommit,
        view: Option<TableViewSpec>,
    ) -> Self {
        Self {
            table,
            edit: None,
            view,
            paste: Some(paste),
        }
    }

    pub(crate) fn replay(&self, wb: &Workbook, undo: bool) -> Result<Workbook, String> {
        wb.ensure_writable()?;
        let sheet = wb
            .sheet_by_id(self.table.sheet_id())
            .ok_or("The history sheet no longer exists.")?;
        if sheet.table_view_spec() != self.view.as_ref() {
            return Err(
                "The Table view changed since this append. Undo/redo was not applied.".into(),
            );
        }
        validate_views(wb)?;
        let mut candidate = wb.clone();
        if undo {
            if let Some(paste) = &self.paste {
                paste.replay(&mut candidate, true)?;
            }
            candidate.apply_table_commit(&self.table, true)?;
            if let Some(edit) = &self.edit {
                edit.replay(&mut candidate, true)?;
            }
        } else {
            if let Some(edit) = &self.edit {
                edit.replay(&mut candidate, false)?;
            }
            candidate.apply_table_commit(&self.table, false)?;
            if let Some(paste) = &self.paste {
                paste.replay(&mut candidate, false)?;
            }
        }
        if let Some(error) = candidate.take_incremental_errors().first() {
            return Err(format!(
                "Could not recalculate Table append history: {error:?}"
            ));
        }
        validate_views(&candidate)?;
        Ok(candidate)
    }
}

fn last_visible_body_row(rows: &RowView, range: TableRange) -> Option<usize> {
    let visible = rows.visible_rows();
    let end = visible.partition_point(|r| *r <= range.end_row);
    end.checked_sub(1)
        .map(|i| visible[i])
        .filter(|r| *r > range.start_row)
}

pub(crate) fn is_last_visible_cell(
    rows: &RowView,
    range: TableRange,
    cell: (usize, usize),
) -> bool {
    cell.1 == range.end_col && last_visible_body_row(rows, range) == Some(cell.0)
}

/// Stay inside the Table when the new record is filtered out, including when
/// all records are hidden. Never select a hidden slot or neighboring notes.
fn append_focus(rows: &RowView, range: TableRange) -> (usize, bool) {
    match rows.data_to_view(range.end_row) {
        Some(row) => (row, false),
        None => (
            last_visible_body_row(rows, range).unwrap_or(range.start_row),
            true,
        ),
    }
}

/// Typing grows only the immediately adjacent row, within the Table's width.
/// Blank edits, existing body cells, side cells and gaps retain ordinary editing.
fn typed_append_target(
    wb: &Workbook,
    index: usize,
    write: &TableCellWrite,
) -> Result<Option<TableId>, String> {
    if write
        .value
        .as_ref()
        .is_none_or(|value| value.trim().is_empty())
    {
        return Ok(None);
    }
    let sheet = wb.sheet(index).ok_or("The sheet no longer exists.")?;
    wb.table_append_target(
        sheet.id,
        TableRange {
            start_row: write.row,
            end_row: write.row,
            start_col: write.col,
            end_col: write.col,
        },
    )
}

impl Spreadsheet {
    /// Some(false) leaves the editor open after a refused growth operation.
    /// Some(true) commits growth; the caller handles ordinary edit completion.
    pub(crate) fn try_typed_table_append(
        &mut self,
        write: &TableCellWrite,
        cx: &mut Context<Self>,
    ) -> Option<bool> {
        let id = match typed_append_target(self.wb(cx), self.sheet_index(cx), write) {
            Ok(None) => return None,
            Ok(Some(id)) => id,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return Some(false);
            }
        };
        if self.block_if_previewing_only(cx) {
            return Some(false);
        }
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            self.status_message = Some(
                "Clear this sheet's worksheet sorting and filters before appending Table rows."
                    .into(),
            );
            cx.notify();
            return Some(false);
        }
        let result = self
            .validate_saved_view_layout(self.wb(cx))
            .and_then(|_| {
                crate::table_bulk_append::prepare_append_writes(
                    self.wb(cx),
                    id,
                    1,
                    std::slice::from_ref(write),
                )
            })
            .and_then(|(candidate, history)| {
                self.validate_saved_view_layout(&candidate)?;
                Ok((candidate, history))
            });
        Some(match result {
            Ok((candidate, history)) => {
                let range = history.table.after_table().unwrap().range;
                let index = candidate
                    .sheet_index_by_id(history.table.sheet_id())
                    .unwrap();
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                let (row, hidden) = append_focus(&self.row_view, range);
                self.view_state.select_cell(row, write.col);
                self.view_state.additional_selections.clear();
                self.ensure_visible(cx);
                self.history.record_action_with_provenance(
                    UndoAction::TableAppend {
                        sheet_index: index,
                        history: Box::new(history),
                        description: "Type and append Table row".into(),
                    },
                    None,
                );
                self.bump_cells_rev();
                self.is_modified = true;
                self.clipboard_visual_range = None;
                self.status_message = Some(if hidden {
                    "Added 1 Table row, hidden by the current view. Unhide rows or clear filters to see it."
                        .into()
                } else {
                    "Added 1 Table row.".into()
                });
                cx.notify();
                true
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                false
            }
        })
    }

    pub(crate) fn append_table_row_in_view(
        &mut self,
        id: TableId,
        edit: Option<TableCellWrite>,
        cx: &mut Context<Self>,
    ) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        self.sync_table_view(cx);
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            self.status_message = Some(
                "Clear this sheet's worksheet sorting and filters before appending Table rows."
                    .into(),
            );
            cx.notify();
            return;
        }
        let result = self
            .validate_saved_view_layout(self.wb(cx))
            .and_then(|_| prepare_append(self.wb(cx), id, edit))
            .and_then(|(candidate, history)| {
                self.validate_saved_view_layout(&candidate)?;
                Ok((candidate, history))
            });
        match result {
            Ok((candidate, history)) => {
                let range = history.table.after_table().unwrap().range;
                let index = candidate
                    .sheet_index_by_id(history.table.sheet_id())
                    .unwrap();
                self.cancel_edit(cx);
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                let (row, hidden) = append_focus(&self.row_view, range);
                self.view_state.select_cell(row, range.start_col);
                self.view_state.additional_selections.clear();
                self.tab_chain_origin_col = Some(range.start_col);
                self.ensure_visible(cx);
                self.history.record_action_with_provenance(
                    UndoAction::TableAppend {
                        sheet_index: index,
                        history: Box::new(history),
                        description: "Append Table row".into(),
                    },
                    None,
                );
                self.bump_cells_rev();
                self.is_modified = true;
                self.clipboard_visual_range = None;
                self.maybe_show_cycle_banner(cx);
                self.surface_incremental_recalc_problems(cx);
                self.status_message = Some(if hidden {
                    "Added 1 Table row, hidden by the current view. Unhide rows or clear filters to enter its values.".into()
                } else {
                    "Added 1 Table row.".into()
                });
                cx.notify();
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
            }
        }
    }

    pub(crate) fn replay_table_append(
        &mut self,
        history: &TableAppendHistory,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = history.replay(self.wb(cx), undo).and_then(|candidate| {
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
    use crate::{history::History, table_edit::tests::fixture};
    use visigrid_engine::{
        cell::NumberFormat,
        filter::SortDirection,
        sheet::{MergedRegion, SheetId},
    };

    fn book(filtered: bool) -> (Workbook, TableId) {
        let mut wb = fixture(filtered);
        let id = wb.active_sheet().tables()[0].id;
        wb.set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        wb.set_cell_value_tracked(0, 4, 3, "999"); // Hidden explicit override.
        wb.set_cell_value_tracked(0, 12, 1, "Notes stay put");
        (wb, id)
    }

    #[test]
    fn footer_fixed_references_follow_tab_append_undo_and_rewind() {
        let (mut before, id) = book(true);
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        before.set_cell_value_tracked(0, 0, 0, "=D8");
        let (after, entry) = prepare_append(&before, id, Some(TableCellWrite::value(6, 3, "15".into()))).unwrap();
        assert_eq!(after.active_sheet().get_raw(0, 0), "=D9");
        assert_eq!(after.active_sheet().get_raw(6, 3), "15");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        let undone = entry.replay(&after, true).unwrap();
        assert_eq!(undone.active_sheet().get_raw(0, 0), "=D8");
        assert_eq!(undone.active_sheet().get_raw(6, 3), before.active_sheet().get_raw(6, 3));
        let redone = entry.replay(&undone, false).unwrap();
        assert_eq!(redone.active_sheet().get_raw(0, 0), "=D9");
        let mut history = History::new();
        history.record_action_with_provenance(UndoAction::TableAppend { sheet_index: 0, history: Box::new(entry), description: "Append with linked totals".into() }, None);
        for (end, reference) in [(0, "=D8"), (1, "=D9")] {
            let preview = history.build_workbook_before(end, Some(&before), 100, 10_000).unwrap();
            assert_eq!(preview.workbook.active_sheet().get_raw(0, 0), reference);
            assert_eq!(preview.workbook.active_sheet().table_view_spec(), before.active_sheet().table_view_spec());
        }
    }

    #[test]
    fn filtered_append_moves_footer_and_rewinds_without_touching_hidden_overrides() {
        let (mut before, id) = book(true);
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        let (after, entry) = prepare_append(&before, id, None).unwrap();
        assert_eq!(after.table(id).unwrap().1.totals_row(), Some(8));
        assert_eq!(after.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(after.active_sheet().get_raw(8, 3), before.active_sheet().get_raw(7, 3));
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().table_view_spec(), before.active_sheet().table_view_spec());
        let restored = entry.replay(&after, true).unwrap();
        assert_eq!(restored.table(id).unwrap().1.totals_row(), Some(7));
        let redone = entry.replay(&restored, false).unwrap();
        assert_eq!(redone.active_sheet().get_raw(8, 3), after.active_sheet().get_raw(8, 3));
        let mut history = History::new();
        history.record_action_with_provenance(UndoAction::TableAppend { sheet_index: 0, history: Box::new(entry), description: "Append with totals".into() }, None);
        let preview = history.build_workbook_before(1, Some(&before), 100, 10_000).unwrap();
        assert_eq!(preview.workbook.table(id).unwrap().1.totals_row(), Some(8));
        assert_eq!(preview.workbook.active_sheet().get_raw(4, 3), "999");
    }

    #[test]
    fn filtered_append_fills_formula_preserves_criteria_and_replays_in_one_step() {
        let (before, id) = book(true);
        let spec = before.active_sheet().table_view_spec().cloned();
        let (after, commit) = prepare_append(&before, id, None).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 7);
        assert_eq!(after.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().get_raw(12, 1), "Notes stay put");
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.rows().data_to_view(7).is_none());
        let (focus, hidden) = append_focus(view.rows(), after.table(id).unwrap().1.range);
        assert!(hidden);
        assert_eq!(view.rows().view_to_data(focus), 6);
        let undo = commit.replay(&after, true).unwrap();
        assert_eq!(undo.table(id).unwrap().1, before.table(id).unwrap().1);
        assert_eq!(undo.active_sheet().get_raw(7, 3), "");
        let redo = commit.replay(&undo, false).unwrap();
        assert_eq!(redo.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(redo.active_sheet().table_view_spec(), spec.as_ref());
    }

    #[test]
    fn tab_from_last_displayed_record_uses_canonical_target_and_restores_percent_format() {
        let (mut before, id) = book(false);
        let mut spec = before.active_sheet().table_view_spec().unwrap().clone();
        spec.sort.as_mut().unwrap().direction = SortDirection::Descending;
        before.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        let view = before
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        let range = before.table(id).unwrap().1.range;
        assert!(is_last_visible_cell(view.rows(), range, (6, 3)));
        assert_eq!(view.rows().view_to_data(6), 4); // Last displayed != last stored.
        assert!(!is_last_visible_cell(view.rows(), range, (5, 3)));
        assert!(!is_last_visible_cell(view.rows(), range, (6, 2)));
        let mut edit = TableCellWrite::value(4, 3, "25%".into());
        let mut format = before.active_sheet().get_format(4, 3);
        format.number_format = NumberFormat::Percent { decimals: 0 };
        edit.format = Some(format);
        let (after, commit) = prepare_append(&before, id, Some(edit)).unwrap();
        assert_eq!(after.active_sheet().get_display(4, 3), "0.25");
        assert_eq!(
            after.active_sheet().get_format(4, 3).number_format,
            NumberFormat::Percent { decimals: 0 }
        );
        assert_eq!(
            after.active_sheet().get_raw(6, 3),
            before.active_sheet().get_raw(6, 3)
        );
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            append_focus(view.rows(), after.table(id).unwrap().1.range),
            (7, false)
        );
        let undo = commit.replay(&after, true).unwrap();
        assert_eq!(undo.active_sheet().get_raw(4, 3), "999");
        assert_eq!(
            undo.active_sheet().get_format(4, 3),
            before.active_sheet().get_format(4, 3)
        );
        let redo = commit.replay(&undo, false).unwrap();
        assert_eq!(redo.active_sheet().get_display(4, 3), "0.25");
        assert_eq!(
            redo.active_sheet().get_format(4, 3).number_format,
            NumberFormat::Percent { decimals: 0 }
        );
    }

    #[test]
    fn new_record_is_selected_at_its_sorted_position() {
        let (mut before, id) = book(false);
        let mut spec = before.active_sheet().table_view_spec().unwrap().clone();
        spec.sort.as_mut().unwrap().column = before.table(id).unwrap().1.columns[2].id;
        before.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        let (after, _) = prepare_append(&before, id, None).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            append_focus(view.rows(), after.table(id).unwrap().1.range),
            (3, false)
        );
        assert_eq!(view.rows().view_to_data(3), 7);
    }

    #[test]
    fn all_hidden_and_header_only_tables_focus_the_header() {
        let (mut before, id) = book(true);
        for row in 3..=6 {
            before.set_cell_value_tracked(0, row, 1, "East");
        }
        let (after, _) = prepare_append(&before, id, None).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            append_focus(view.rows(), after.table(id).unwrap().1.range),
            (2, true)
        );
        assert!(!is_last_visible_cell(
            view.rows(),
            after.table(id).unwrap().1.range,
            (2, 3)
        ));
        // Remove records entirely, keeping the saved criterion.
        for row in 3..=6 {
            for col in 1..=3 {
                before.clear_cell_tracked(0, row, col);
            }
        }
        let mut range = before.table(id).unwrap().1.range;
        range.end_row = range.start_row;
        before.resize_table(id, range).unwrap();
        let (after, _) = prepare_append(&before, id, None).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 3);
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            append_focus(view.rows(), after.table(id).unwrap().1.range),
            (2, true)
        );
    }

    #[test]
    fn collision_rejects_pending_edit_and_append_together() {
        for collision in 0..4 {
            let (mut before, id) = book(true);
            match collision {
                0 => {
                    before.set_cell_value_tracked(0, 7, 1, "Do not replace");
                }
                1 => {
                    before.set_cell_value_tracked(0, 7, 5, "Beside new Table row");
                }
                2 => {
                    before
                        .active_sheet_mut()
                        .add_merge(MergedRegion {
                            start: (7, 1),
                            end: (7, 2),
                        })
                        .unwrap();
                }
                _ => {
                    before.active_sheet_mut().set_comment(
                        7,
                        2,
                        Some(visigrid_engine::cell::CellComment {
                            text: "Keep comment".into(),
                            author: "".into(),
                        }),
                    );
                }
            }
            let revision = before.revision();
            assert!(
                prepare_append(&before, id, Some(TableCellWrite::value(6, 3, "15".into())))
                    .is_err()
            );
            assert_eq!(before.revision(), revision);
            assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
            assert_eq!(before.active_sheet().get_raw(6, 3), "=[@Amount]*2");
        }
    }

    #[test]
    fn hidden_pending_edit_is_rejected() {
        let (before, id) = book(true);
        assert!(
            prepare_append(&before, id, Some(TableCellWrite::value(4, 3, "15".into())))
                .unwrap_err()
                .contains("hidden")
        );
        assert_eq!(before.active_sheet().get_raw(4, 3), "999");
    }

    #[test]
    fn dependent_spill_on_another_filtered_sheet_rejects_append() {
        let (mut before, id) = book(true);
        let other = before.add_sheet_named("Other").unwrap();
        let other_id = before.sheet(other).unwrap().id;
        for (r, a, b) in [(2, "Group", "Amount"), (3, "West", "10"), (4, "West", "20")] {
            before.set_cell_value_tracked(other, r, 1, a);
            before.set_cell_value_tracked(other, r, 2, b);
        }
        let t = before
            .create_table(
                other_id,
                TableRange {
                    start_row: 2,
                    end_row: 4,
                    start_col: 1,
                    end_col: 2,
                },
                "OtherTable",
            )
            .unwrap()
            .table_id();
        let mut spec = TableViewSpec::new(t);
        spec.sort = Some(visigrid_engine::table_view::TableSort {
            column: before.table(t).unwrap().1.columns[1].id,
            direction: SortDirection::Ascending,
        });
        before.set_table_view_spec(other_id, Some(spec)).unwrap();
        before.set_cell_value_tracked(other, 0, 0, "=IF(ROWS(Sales[Amount])>4,SEQUENCE(5),0)");
        validate_views(&before).unwrap();
        assert!(prepare_append(&before, id, None).is_err());
        assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
        assert_eq!(before.sheet(other).unwrap().get_display(0, 0), "0");
    }

    #[test]
    fn append_on_another_sheet_preserves_existing_view() {
        let (mut before, _) = book(true);
        let other = before.add_sheet_named("Other").unwrap();
        let sheet = before.sheet(other).unwrap().id;
        before.set_cell_value_tracked(other, 0, 0, "Value");
        let id = before
            .create_table(
                sheet,
                TableRange {
                    start_row: 0,
                    end_row: 1,
                    start_col: 0,
                    end_col: 0,
                },
                "OtherTable",
            )
            .unwrap()
            .table_id();
        let (after, history) = prepare_append(&before, id, None).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 2);
        assert_eq!(
            after.sheet(0).unwrap().table_view_spec(),
            before.sheet(0).unwrap().table_view_spec()
        );
        let undo = history.replay(&after, true).unwrap();
        assert_eq!(undo.table(id).unwrap().1.range.end_row, 1);
    }

    #[test]
    fn stale_replay_refuses_without_discarding_new_values_or_edit() {
        let (before, id) = book(true);
        let (mut after, history) =
            prepare_append(&before, id, Some(TableCellWrite::value(6, 3, "15".into()))).unwrap();
        after.set_cell_value_tracked(0, 7, 1, "New user input");
        assert!(history.replay(&after, true).is_err());
        assert_eq!(after.active_sheet().get_raw(6, 3), "15");
        assert_eq!(after.table(id).unwrap().1.range.end_row, 7);
        let (mut after, history) = prepare_append(&before, id, None).unwrap();
        after.set_table_view_spec(SheetId(7), None).unwrap();
        assert!(history
            .replay(&after, true)
            .unwrap_err()
            .contains("view changed"));
    }

    #[test]
    fn preview_replays_append_and_pending_edit_from_loaded_base() {
        let (before, id) = book(true);
        let (_, commit) =
            prepare_append(&before, id, Some(TableCellWrite::value(6, 3, "15".into()))).unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableAppend {
                sheet_index: 0,
                history: Box::new(commit),
                description: "Append Table row".into(),
            },
            None,
        );
        assert_eq!(history.undo_count(), 1);
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(preview.workbook.table(id).unwrap().1.range.end_row, 7);
        assert_eq!(preview.workbook.active_sheet().get_raw(6, 3), "15");
        assert_eq!(
            preview.workbook.active_sheet().get_raw(7, 3),
            "=[@Amount]*2"
        );
        let preview = history
            .build_workbook_before(0, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(preview.workbook.table(id).unwrap().1.range.end_row, 6);
        assert_eq!(
            preview.workbook.active_sheet().get_raw(6, 3),
            "=[@Amount]*2"
        );
    }

    #[test]
    fn expanded_table_checks_desktop_layout_at_new_row() {
        let (before, id) = book(true);
        let (after, _) = prepare_append(&before, id, None).unwrap();
        let heights = [(7, 40.0)].into_iter().collect();
        let hidden = [7].into_iter().collect();
        let old = before.table(id).unwrap().1;
        let new = after.table(id).unwrap().1;
        assert!(
            crate::table_filter_ui::desktop_layout_error(old, Some(&heights), None, 0).is_none()
        );
        assert!(
            crate::table_filter_ui::desktop_layout_error(new, Some(&heights), None, 0).is_some()
        );
        assert!(
            crate::table_filter_ui::desktop_layout_error(new, None, Some(&hidden), 0).is_none()
        );
        assert!(crate::table_filter_ui::desktop_layout_error(new, None, None, 7).is_some());
    }
    #[test]
    fn recovery_and_grid_boundary_refuse_before_changing_anything() {
        let (mut before, id) = book(true);
        before.active_sheet_mut().read_only_reason = Some("Damaged table metadata".into());
        assert!(prepare_append(&before, id, None)
            .unwrap_err()
            .contains("Read-only"));
        before.active_sheet_mut().read_only_reason = None;
        before.active_sheet_mut().rows = 7;
        assert!(prepare_append(&before, id, None).is_err());
        assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
    }

    #[test]
    fn appended_filtered_table_roundtrips_native_full_save() {
        let (before, id) = book(true);
        let (after, _) = prepare_append(&before, id, None).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("append.sheet");
        visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
            .unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(loaded.table(id).unwrap().1, after.table(id).unwrap().1);
        assert_eq!(
            loaded.active_sheet().table_view_spec(),
            after.active_sheet().table_view_spec()
        );
        assert_eq!(loaded.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(loaded.active_sheet().get_raw(4, 3), "999");
        let view = loaded
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.rows().data_to_view(7).is_none());
    }
    fn typed(
        wb: &Workbook,
        write: TableCellWrite,
    ) -> Result<(Workbook, TableAppendHistory), String> {
        let id = typed_append_target(wb, 0, &write)?.ok_or("No append target")?;
        crate::table_bulk_append::prepare_append_writes(wb, id, 1, &[write])
    }

    #[test]
    fn typed_growth_requires_nonblank_value_immediately_below_within_width() {
        let (wb, id) = book(true);
        assert_eq!(
            typed_append_target(&wb, 0, &TableCellWrite::value(7, 1, "West".into())).unwrap(),
            Some(id)
        );
        for (row, col, value) in [
            (7, 1, ""),
            (7, 1, "  "),
            (8, 1, "West"),
            (7, 0, "West"),
            (7, 4, "West"),
            (6, 1, "West"),
        ] {
            assert!(
                typed_append_target(&wb, 0, &TableCellWrite::value(row, col, value.into()))
                    .unwrap()
                    .is_none()
            );
        }
    }

    #[test]
    fn typing_below_filter_grows_fills_and_replays_without_touching_hidden_records() {
        let (before, id) = book(true);
        let (after, history) = typed(&before, TableCellWrite::value(7, 1, "West".into())).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 7);
        assert_eq!(after.active_sheet().get_raw(7, 1), "West");
        assert_eq!(after.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        let (focus, hidden) = append_focus(view.rows(), after.table(id).unwrap().1.range);
        assert!(!hidden);
        assert_eq!(view.rows().view_to_data(focus), 7);
        let undo = history.replay(&after, true).unwrap();
        assert_eq!(undo.table(id).unwrap().1.range.end_row, 6);
        assert_eq!(undo.active_sheet().get_raw(7, 1), "");
        assert_eq!(undo.active_sheet().get_raw(7, 3), "");
        assert_eq!(
            history
                .replay(&undo, false)
                .unwrap()
                .active_sheet()
                .get_raw(7, 1),
            "West"
        );
    }

    #[test]
    fn typed_percentage_hidden_by_filter_restores_its_value_and_format_together() {
        let (before, id) = book(true);
        let mut write = TableCellWrite::value(7, 2, "25%".into());
        let mut format = before.active_sheet().get_format(7, 2);
        format.number_format = NumberFormat::Percent { decimals: 0 };
        write.format = Some(format.clone());
        let (after, history) = typed(&before, write).unwrap();
        assert_eq!(after.active_sheet().get_display(7, 2), "0.25");
        assert_eq!(
            after.active_sheet().get_computed_value(7, 3),
            visigrid_engine::formula::eval::Value::Number(0.5)
        );
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(append_focus(view.rows(), after.table(id).unwrap().1.range).1);
        let undo = history.replay(&after, true).unwrap();
        assert_eq!(
            undo.active_sheet().get_format(7, 2),
            before.active_sheet().get_format(7, 2)
        );
        assert_eq!(undo.active_sheet().get_raw(7, 2), "");
        let redo = history.replay(&undo, false).unwrap();
        assert_eq!(redo.active_sheet().get_format(7, 2), format);
        assert_eq!(
            redo.active_sheet().get_computed_value(7, 3),
            visigrid_engine::formula::eval::Value::Number(0.5)
        );
    }

    #[test]
    fn typed_formula_is_an_override_without_replacing_the_column_rule() {
        let (before, id) = book(true);
        let (after, history) = typed(
            &before,
            TableCellWrite::value(7, 3, "=SUM(Sales[Amount])".into()),
        )
        .unwrap();
        assert_eq!(after.active_sheet().get_display(7, 3), "100");
        assert_eq!(
            after.table(id).unwrap().1.columns,
            before.table(id).unwrap().1.columns
        );
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .active_sheet()
                .get_raw(7, 3),
            ""
        );
    }

    #[test]
    fn typed_append_refuses_occupied_stripe_and_unsafe_recalculation_atomically() {
        for case in 0..3 {
            let (mut before, id) = book(true);
            match case {
                0 => {
                    before.set_cell_value_tracked(0, 7, 3, "Existing note");
                }
                1 => {
                    before.set_cell_value_tracked(0, 7, 5, "Beside Table");
                }
                _ => {
                    before.set_cell_value_tracked(
                        0,
                        0,
                        0,
                        "=IF(SUM(Sales[Amount])>100,SEQUENCE(5),0)",
                    );
                }
            }
            let revision = before.revision();
            assert!(typed(&before, TableCellWrite::value(7, 2, "50".into())).is_err());
            assert_eq!(before.revision(), revision);
            assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
            assert_eq!(before.active_sheet().get_raw(7, 2), "");
        }
    }

    #[test]
    fn typed_append_into_header_only_table_and_on_other_sheet() {
        let (mut before, _) = book(true);
        let index = before.add_sheet_named("Other").unwrap();
        let sid = before.sheet(index).unwrap().id;
        before.set_cell_value_tracked(index, 0, 0, "Value");
        let id = before
            .create_table(
                sid,
                TableRange {
                    start_row: 0,
                    end_row: 0,
                    start_col: 0,
                    end_col: 0,
                },
                "OtherTable",
            )
            .unwrap()
            .table_id();
        let write = TableCellWrite::value(1, 0, "First record".into());
        assert_eq!(
            typed_append_target(&before, index, &write).unwrap(),
            Some(id)
        );
        let (after, history) =
            crate::table_bulk_append::prepare_append_writes(&before, id, 1, &[write]).unwrap();
        assert_eq!(after.sheet(index).unwrap().get_raw(1, 0), "First record");
        assert_eq!(
            after.sheet(0).unwrap().table_view_spec(),
            before.sheet(0).unwrap().table_view_spec()
        );
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .table(id)
                .unwrap()
                .1
                .range
                .data_rows(),
            0
        );
    }

    #[test]
    fn typed_append_has_one_replayable_history_entry_and_persists() {
        let (before, id) = book(true);
        let (after, commit) = typed(&before, TableCellWrite::value(7, 1, "West".into())).unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableAppend {
                sheet_index: 0,
                history: Box::new(commit),
                description: "Type and append Table row".into(),
            },
            None,
        );
        assert_eq!(history.undo_count(), 1);
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(preview.workbook.table(id).unwrap().1.range.end_row, 7);
        assert_eq!(preview.workbook.active_sheet().get_raw(7, 1), "West");
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("typed.sheet");
        visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
            .unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(loaded.table(id).unwrap().1, after.table(id).unwrap().1);
        assert_eq!(loaded.active_sheet().get_raw(7, 1), "West");
        assert_eq!(loaded.active_sheet().get_raw(7, 3), "=[@Amount]*2");
        assert_eq!(
            loaded.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
    }
}
