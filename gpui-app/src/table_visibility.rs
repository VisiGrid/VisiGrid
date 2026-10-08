//! Manual visibility and totals metadata are one guarded history operation.
use crate::{app::Spreadsheet, history::UndoAction, table_structure::{StructureLayout, TableStructureHistory}};
use gpui::Context;
use visigrid_engine::{filter::RowView, sheet::SheetId, workbook::Workbook};
use std::collections::BTreeSet;

/// Resolve the current displayed selection once, before visibility changes.
/// Unhide spans invisible slots too, but only removes manual flags.
pub(crate) fn selected_row_visibility(
    rows: &RowView, manual: &BTreeSet<usize>, start: usize, end: usize, hidden: bool,
) -> Result<Vec<usize>, String> {
    if start > end || end >= rows.row_count() {
        return Err("The row selection is outside the worksheet.".into());
    }
    Ok((start..=end).filter_map(|slot| {
        let row = rows.view_to_data(slot);
        if hidden {
            (rows.is_view_row_visible(slot) && !manual.contains(&row)).then_some(row)
        } else {
            manual.contains(&row).then_some(row)
        }
    }).collect())
}

#[cfg(test)]
fn prepare(
    wb: &Workbook, id: SheetId, before: &StructureLayout, start: usize, end: usize, hidden: bool,
) -> Result<Option<(Workbook, TableStructureHistory, usize)>, String> {
    let sheet = wb.sheet_by_id(id).ok_or("Visibility sheet no longer exists.")?;
    let view = sheet.build_saved_table_view(sheet.rows)?;
    let identity;
    let rows = if let Some(view) = &view { view.rows() } else {
        identity = RowView::new(sheet.rows);
        &identity
    };
    prepare_in_view(wb, id, before, rows, start, end, hidden)
}

fn prepare_in_view(
    wb: &Workbook, id: SheetId, before: &StructureLayout, rows: &RowView,
    start: usize, end: usize, hidden: bool,
) -> Result<Option<(Workbook, TableStructureHistory, usize)>, String> {
    let sheet = wb.sheet_by_id(id).ok_or("Visibility sheet no longer exists.")?;
    if start > end || end >= sheet.rows.min(crate::app::NUM_ROWS) {
        return Err("The row selection is outside the worksheet.".into());
    }
    if let Some(table) = sheet.table_view_spec().filter(|spec| spec.has_criteria())
        .and_then(|spec| sheet.tables().iter().find(|t| t.id == spec.table)) {
        if crate::table_filter_ui::desktop_layout_error(
            table, Some(&before.heights), Some(&before.hidden_rows), sheet.frozen_panes.0,
        ).is_some() {
            return Err("Reset body row heights or unfreeze the Table body before changing row visibility.".into());
        }
    }
    let targets = selected_row_visibility(rows, &before.hidden_rows, start, end, hidden)?;
    if targets.iter().any(|row| *row >= sheet.rows) {
        return Err("The row selection is outside the worksheet.".into());
    }
    let mut after = before.clone();
    let count = targets.len();
    for row in targets {
        if hidden { after.hidden_rows.insert(row); } else { after.hidden_rows.remove(&row); }
    }
    if count == 0 { return Ok(None); }
    if let Some(table) = sheet.table_view_spec().filter(|spec| spec.has_criteria())
        .and_then(|spec| sheet.tables().iter().find(|t| t.id == spec.table)) {
        if crate::table_filter_ui::desktop_layout_error(
            table, Some(&after.heights), Some(&after.hidden_rows), sheet.frozen_panes.0,
        ).is_some() { return Err("Reset body row heights or unfreeze the Table body before changing row visibility.".into()); }
    }
    let (candidate, commit) = wb.prepare_table_row_visibility(id, after.hidden_rows.clone())?;
    Ok(Some((candidate, TableStructureHistory { commit, source_frozen: None, before: before.clone(), after }, count)))
}

impl Spreadsheet {
    /// Keep the engine's calculation/serialization state aligned with the host
    /// layout after ordinary visibility history or an imported layout install.
    pub(crate) fn sync_manual_row_visibility(&mut self, id: SheetId, cx: &mut Context<Self>) -> bool {
        let hidden = self.hidden_rows.get(&id).cloned().unwrap_or_default();
        let Some(sheet) = self.wb(cx).sheet_by_id(id) else { return false; };
        let old = sheet.manual_hidden_rows();
        if old == hidden { return true; }
        let result = self.wb(cx).prepare_table_row_visibility(id, hidden);
        let candidate = match result {
            Ok((candidate, _)) => candidate,
            Err(error) => {
                self.hidden_rows.insert(id, old);
                self.status_message = Some(error);
                return false;
            }
        };
        self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.table_view_sync_key = None;
        self.bump_cells_rev();
        true
    }

    pub(crate) fn change_table_row_visibility(&mut self, hidden: bool, cx: &mut Context<Self>) {
        self.sync_table_view(cx);
        let ((start, _), (end, _)) = self.selection_range();
        let id = self.cached_sheet_id();
        let before = self.structure_layout(id);
        let result = prepare_in_view(self.wb(cx), id, &before, &self.row_view, start, end, hidden).and_then(|result| {
            if let Some((candidate, history, _)) = &result {
                self.validate_structure_layout(candidate, id, &history.after)?;
            }
            Ok(result)
        });
        match result {
            Ok(Some((candidate, history, count))) => {
                let description = format!("{} {count} row(s)", if hidden { "Hide" } else { "Unhide" });
                let index = self.sheet_index(cx);
                self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.install_structure_layout(id, &history.after);
                self.record_action_with_provenance(cx, UndoAction::TableStructureChanged {
                    sheet_index: index, history: Box::new(history), description: description.clone(),
                }, None);
                self.sync_table_view(cx);
                self.table_filter_dropdown = None;
                self.bump_cells_rev();
                self.is_modified = true;
                self.status_message = Some(description);
            }
            Ok(None) => self.status_message = Some(if hidden { "Selected rows are already hidden." } else { "No manually hidden rows in selection." }.into()),
            Err(error) => self.status_message = Some(error),
        }
        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{history::History, table_edit::tests::fixture};

    fn book() -> Workbook {
        let mut wb = fixture(false);
        let id = wb.active_sheet().tables()[0].id;
        wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb
    }

    #[test]
    fn hide_unhide_rewinds_visibility_and_totals_together() {
        let wb = book();
        let id = wb.active_sheet_id();
        let layout = StructureLayout::default();
        let (after, history, count) = prepare(&wb, id, &layout, 4, 4, true).unwrap().unwrap();
        assert_eq!(count, 1);
        assert_eq!(after.active_sheet().get_display(7, 3), "180");
        assert_eq!(history.after.hidden_rows, [4].into_iter().collect());
        let undone = history.commit.candidate(&after, true).unwrap();
        assert_eq!(undone.active_sheet().get_display(7, 3), "200");
        let (shown, unhide, _) = prepare(&after, id, &history.after, 3, 5, false).unwrap().unwrap();
        assert_eq!(shown.active_sheet().get_display(7, 3), "200");
        assert!(unhide.after.hidden_rows.is_empty());
        let mut stack = History::new();
        for h in [history, unhide] {
            stack.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableStructureChanged {
                sheet_index: 0, history: Box::new(h), description: "Row visibility".into(),
            }, None);
        }
        for (end, expected, hidden) in [(0, "200", false), (1, "180", true), (2, "200", false)] {
            let preview = stack.build_workbook_before(end, Some(&wb), 100, 10_000).unwrap();
            assert_eq!(preview.workbook.active_sheet().get_display(7, 3), expected);
            assert_eq!(preview.view_state.per_sheet[0].structure_layout.as_ref().unwrap().hidden_rows.contains(&4), hidden);
        }
    }

    #[test]
    fn criteria_allow_manual_hides_but_keep_height_and_freeze_guards() {
        let mut wb = fixture(true);
        let id = wb.active_sheet_id();
        let table = wb.active_sheet().tables()[0].id;
        wb.set_table_totals_visible(table, true, Default::default()).unwrap();
        let (hidden, history, count) = prepare(&wb, id, &Default::default(), 3, 6, true).unwrap().unwrap();
        assert_eq!(count, 3); // the East record is already filter-hidden
        assert_eq!(history.after.hidden_rows, [3, 5, 6].into());
        assert_eq!(hidden.active_sheet().get_display(7, 3), "0");
        assert_eq!(hidden.active_sheet().table_view_spec(), wb.active_sheet().table_view_spec());
        let (shown, _, count) = prepare(&hidden, id, &history.after, 3, 6, false).unwrap().unwrap();
        assert_eq!(count, 3);
        assert_eq!(shown.active_sheet().get_display(7, 3), "180");
        assert!(!shown.active_sheet().build_saved_table_view(30).unwrap().unwrap().rows().is_data_row_visible(4));
        let mut invalid_layout = StructureLayout::default();
        invalid_layout.heights.insert(4, 42.0);
        assert!(prepare(&wb, id, &invalid_layout, 3, 6, false).unwrap_err().contains("row heights"));
        let (outside, history, _) = prepare(&wb, id, &Default::default(), 12, 13, true).unwrap().unwrap();
        assert_eq!(outside.active_sheet().get_display(7, 3), "180");
        assert_eq!(outside.active_sheet().table_view_spec(), wb.active_sheet().table_view_spec());
        assert!(history.commit.candidate(&outside, true).is_ok());
        let other = wb.add_sheet_named("Other").unwrap();
        let other_id = wb.sheet(other).unwrap().id;
        let (after, history, _) = prepare(&wb, other_id, &Default::default(), 3, 4, true).unwrap().unwrap();
        assert_eq!(history.commit.sheet, other_id);
        assert_eq!(history.after.hidden_rows, [3, 4].into_iter().collect());
        assert_eq!(after.sheet(0).unwrap().table_view_spec(), wb.sheet(0).unwrap().table_view_spec());
    }

    #[test]
    fn unhide_on_another_totals_sheet_keeps_saved_criteria() {
        let mut wb = fixture(true);
        let original_view = wb.active_sheet().table_view_spec().cloned();
        let other = wb.add_sheet_named("Other").unwrap();
        for (row, value) in ["Amount", "10", "20"].iter().enumerate() { wb.set_cell_value_tracked(other, row, 0, value); }
        let id = wb.sheet(other).unwrap().id;
        let table = wb.create_table(id, visigrid_engine::table::TableRange {
            start_row: 0, end_row: 2, start_col: 0, end_col: 0,
        }, "OtherData").unwrap().table_id();
        wb.set_table_totals_visible(table, true, [1].into_iter().collect()).unwrap();
        let mut layout = StructureLayout::default();
        layout.hidden_rows.insert(1);
        assert_eq!(wb.sheet(other).unwrap().get_display(3, 0), "20");
        let (after, history, _) = prepare(&wb, id, &layout, 0, 2, false).unwrap().unwrap();
        assert!(history.after.hidden_rows.is_empty());
        assert_eq!(after.sheet(other).unwrap().get_display(3, 0), "30");
        assert_eq!(after.active_sheet().table_view_spec(), original_view.as_ref());
        let undone = history.commit.candidate(&after, true).unwrap();
        assert_eq!(undone.sheet(other).unwrap().get_display(3, 0), "20");
    }

    #[test]
    fn totals_visibility_preflight_refuses_new_spills_on_another_view_sheet() {
        let mut wb = book();
        let other = wb.add_sheet_named("Other").unwrap();
        wb.set_cell_value_tracked(other, 2, 0, "Value");
        wb.set_cell_value_tracked(other, 3, 0, "1");
        let table = wb.create_table(wb.sheet(other).unwrap().id, visigrid_engine::table::TableRange {
            start_row: 2, end_row: 3, start_col: 0, end_col: 0,
        }, "OtherData").unwrap().table_id();
        let mut spec = visigrid_engine::table_view::TableViewSpec::new(table);
        spec.sort = Some(visigrid_engine::table_view::TableSort {
            column: wb.table(table).unwrap().1.columns[0].id,
            direction: visigrid_engine::filter::SortDirection::Ascending,
        });
        wb.set_table_view_spec(wb.sheet(other).unwrap().id, Some(spec)).unwrap();
        wb.set_cell_value_tracked(other, 0, 1, "=IF(SUM(Sales[#Totals])<200,SEQUENCE(6),0)");
        let revision = wb.revision();
        assert!(prepare(&wb, wb.active_sheet_id(), &Default::default(), 4, 4, true).is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.active_sheet().get_display(7, 3), "200");
    }
    #[test]
    fn manually_hidden_sorted_records_are_excluded_from_paste_fill_cut_and_row_deletion() {
        use crate::table_edit::{prepare_table_writes, view_safe_paste_targets, view_safe_selection_rows, TableCellWrite};
        let mut wb = fixture(true);
        let id = wb.active_sheet_id();
        let table = wb.active_sheet().tables()[0].id;
        wb.set_table_totals_visible(table, true, Default::default()).unwrap();
        // Display slot 5 is canonical row 3 (West, 30), not row 5 (West, 20).
        let (hidden, history, count) = prepare(&wb, id, &Default::default(), 5, 5, true).unwrap().unwrap();
        assert_eq!(count, 1);
        assert_eq!(hidden.active_sheet().manual_hidden_rows(), [3].into());
        let sheet = hidden.active_sheet();
        let view = sheet.build_saved_table_view(30).unwrap().unwrap();
        let rows = view.rows();
        assert_eq!(rows.visible_rows().iter().copied().filter(|r| (3..=6).contains(r)).collect::<Vec<_>>(), [4, 6]);
        assert_eq!(view.focus_record(3).unwrap().view_row, 6);
        assert_eq!(view_safe_selection_rows(sheet, rows, ((4, 2), (6, 2))).unwrap(), [5, 6]);
        let targets = view_safe_paste_targets(sheet, rows, (4, 2), 2, 1).unwrap();
        assert_eq!(targets.iter().map(|p| p.0).collect::<Vec<_>>(), [5, 6]);
        assert!(view_safe_paste_targets(sheet, rows, (4, 2), 3, 1).is_err());
        let writes: Vec<_> = targets.iter().map(|p| TableCellWrite::value(p.0, p.1, "99".into())).collect();
        let pasted = prepare_table_writes(&hidden, 0, &writes).unwrap();
        assert_eq!(pasted.active_sheet().get_raw(3, 2), "30");
        assert_eq!(pasted.active_sheet().get_raw(4, 2), "10");
        assert!(prepare_table_writes(&hidden, 0, &[TableCellWrite::value(3, 2, "99".into())]).is_err());
        let (clipboard, cut) = crate::table_cut::plan_cut(sheet, rows, ((4, 2), (6, 2))).unwrap();
        assert_eq!(clipboard.source_rows, [5, 6]);
        assert_eq!(clipboard.raw_tsv, "20\n40");
        let cut = prepare_table_writes(&hidden, 0, &cut).unwrap();
        assert_eq!(cut.active_sheet().get_raw(3, 2), "30");
        assert_eq!(cut.active_sheet().get_raw(4, 2), "10");
        let fill = crate::table_fill::plan_direction(sheet, rows, ((4, 2), (6, 2)), true).unwrap();
        assert_eq!(fill.iter().map(|w| w.row).collect::<Vec<_>>(), [6]);
        assert_eq!(fill[0].value.as_deref(), Some("20"));
        let append = crate::table_bulk_append::plan_bulk_append(sheet, rows, (4, 2), 3, 1).unwrap().unwrap();
        assert_eq!(append.count, 1);
        assert_eq!(append.targets.iter().map(|p| p.0).collect::<Vec<_>>(), [5, 6, 7]);
        let writes: Vec<_> = append.targets.iter().map(|p| TableCellWrite::value(p.0, p.1, "99".into())).collect();
        let (appended, append_history) = crate::table_bulk_append::prepare_append_writes(&hidden, table, append.count, &writes).unwrap();
        assert_eq!(appended.active_sheet().manual_hidden_rows(), [3].into());
        assert_eq!(appended.active_sheet().get_raw(3, 2), "30");
        assert_eq!(appended.active_sheet().get_raw(4, 2), "10");
        assert_eq!(appended.table(table).unwrap().1.totals_row(), Some(8));
        let undone = append_history.replay(&appended, true).unwrap();
        assert_eq!(undone.table(table).unwrap().1.totals_row(), Some(7));
        assert_eq!(undone.active_sheet().manual_hidden_rows(), [3].into());
        let steps = crate::table_structure::selected_row_steps(rows, 3, 6, true).unwrap();
        let (deleted, deletion) = hidden.prepare_guarded_structure(0, steps.clone()).unwrap();
        assert_eq!(deleted.active_sheet().get_raw(3, 2), "30");
        assert_eq!(deleted.active_sheet().get_raw(4, 2), "10");
        assert_eq!(deleted.active_sheet().manual_hidden_rows(), [3].into());
        assert_eq!(history.after.shifted(sheet, &steps).unwrap().hidden_rows, [3].into());
        assert_eq!(deletion.candidate(&deleted, true).unwrap().active_sheet().manual_hidden_rows(), [3].into());
        let mut stack = History::new();
        stack.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableStructureChanged {
            sheet_index: 0, history: Box::new(history), description: "Hide sorted record".into(),
        }, None);
        let preview = stack.build_workbook_before(1, Some(&wb), 100, 10_000).unwrap();
        assert!(!preview.view_state.per_sheet[0].table_rows.as_ref().unwrap().is_data_row_visible(3));
        assert!(!preview.view_state.per_sheet[0].table_rows.as_ref().unwrap().is_data_row_visible(4));
        let preview = stack.build_workbook_before(0, Some(&wb), 100, 10_000).unwrap();
        assert!(preview.view_state.per_sheet[0].table_rows.as_ref().unwrap().is_data_row_visible(3));
        assert!(!preview.view_state.per_sheet[0].table_rows.as_ref().unwrap().is_data_row_visible(4));
    }

    #[test]
    fn worksheet_hide_targets_canonical_visible_rows_and_unhide_keeps_filter_mask() {
        let mut rows = RowView::new(8);
        rows.apply_sort(vec![0, 4, 2, 1, 3, 5, 6, 7]);
        rows.apply_filter(vec![true, true, false, true, true, true, true, true]);
        let manual = [4].into();
        assert_eq!(selected_row_visibility(&rows, &manual, 1, 4, true).unwrap(), [1, 3]);
        assert_eq!(selected_row_visibility(&rows, &manual, 1, 1, false).unwrap(), [4]);
        assert!(selected_row_visibility(&rows, &manual, 2, 2, true).unwrap().is_empty());
        let manual = [1, 2, 3, 4].into();
        assert_eq!(selected_row_visibility(&rows, &manual, 1, 4, false).unwrap(), [4, 2, 1, 3]);
        assert!(!rows.is_data_row_visible(2), "unhide must not replace the worksheet filter mask");
        assert!(selected_row_visibility(&rows, &manual, 7, 8, true).is_err());
        assert!(selected_row_visibility(&rows, &manual, 5, 2, false).is_err());
    }

    #[test]
    fn worksheet_visibility_keeps_sort_history_and_other_sheet_table_criteria() {
        let mut wb = fixture(true);
        let table_spec = wb.active_sheet().table_view_spec().cloned();
        let other = wb.add_sheet_named("Plain").unwrap();
        wb.sheet_mut(other).unwrap().rows = 12;
        for (row, value) in ["Amount", "40", "10", "30", "20"].iter().enumerate() {
            wb.set_cell_value_tracked(other, row, 0, value);
        }
        wb.set_cell_value_tracked(other, 0, 2, "=SUBTOTAL(109,A2:A5)");
        let id = wb.sheet(other).unwrap().id;
        let (wb, _) = wb.prepare_table_row_visibility(id, [4].into()).unwrap();
        let mut layout = StructureLayout::default();
        layout.hidden_rows.insert(4);
        let mut rows = RowView::new(12);
        rows.apply_sort(vec![0, 2, 4, 3, 1, 5, 6, 7, 8, 9, 10, 11]);
        rows.apply_filter((0..12).map(|row| row != 3).collect());
        let (hidden, history, count) = prepare_in_view(&wb, id, &layout, &rows, 1, 4, true).unwrap().unwrap();
        assert_eq!(count, 2);
        assert_eq!(history.after.hidden_rows, [1, 2, 4].into());
        assert_eq!(hidden.sheet(other).unwrap().get_display(0, 2), "30");
        assert_eq!(hidden.sheet(0).unwrap().table_view_spec(), table_spec.as_ref());
        assert_eq!(history.commit.changed_cell_count(), 0);
        let undone = history.commit.candidate(&hidden, true).unwrap();
        assert_eq!(undone.sheet(other).unwrap().manual_hidden_rows(), [4].into());
        assert_eq!(undone.sheet(other).unwrap().get_display(0, 2), "80");
        let redone = history.commit.candidate(&undone, false).unwrap();
        assert_eq!(redone.sheet(other).unwrap().manual_hidden_rows(), [1, 2, 4].into());
        let (shown, unhide, count) = prepare_in_view(&hidden, id, &history.after, &rows, 1, 4, false).unwrap().unwrap();
        assert_eq!(count, 3);
        assert!(shown.sheet(other).unwrap().manual_hidden_rows().is_empty());
        assert!(!rows.is_data_row_visible(3));
        assert_eq!(shown.sheet(0).unwrap().table_view_spec(), table_spec.as_ref());
        let mut stack = History::new();
        stack.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::SortApplied {
            sheet_index: other,
            previous_row_order: (0..12).collect(),
            new_row_order: rows.row_order().to_vec(),
            previous_sort_state: None,
            new_sort_state: (0, true),
        }, None);
        for h in [history, unhide] {
            stack.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableStructureChanged {
                sheet_index: other, history: Box::new(h), description: "Worksheet visibility".into(),
            }, None);
        }
        for (end, hidden) in [(1, vec![4]), (2, vec![1, 2, 4]), (3, vec![])] {
            let preview = stack.build_workbook_before(end, Some(&wb), 100, 10_000).unwrap();
            let view = &preview.view_state.per_sheet[other];
            assert_eq!(view.row_order.as_deref(), Some(rows.row_order()));
            assert_eq!(view.sort, Some((0, true)));
            assert_eq!(view.structure_layout.as_ref().unwrap().hidden_rows, hidden.into_iter().collect());
            assert_eq!(preview.workbook.sheet(0).unwrap().table_view_spec(), table_spec.as_ref());
        }
    }

}
