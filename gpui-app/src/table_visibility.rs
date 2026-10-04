//! Manual visibility and totals metadata are one guarded history operation.
use crate::{app::Spreadsheet, history::UndoAction, table_structure::{StructureLayout, TableStructureHistory}};
use gpui::Context;
use visigrid_engine::{sheet::SheetId, workbook::Workbook};

fn prepare(
    wb: &Workbook, id: SheetId, before: &StructureLayout, start: usize, end: usize, hidden: bool,
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
            return Err("Clear this Table's sorting and filters before changing its manual row visibility.".into());
        }
    }
    let view = sheet.build_saved_table_view(sheet.rows)?;
    let mut after = before.clone();
    let mut count = 0;
    for slot in start..=end {
        let row = view.as_ref().map_or(slot, |view| view.rows().view_to_data(slot));
        // Hiding skips filtered-out records. Unhide removes only manual flags;
        // it never changes the saved filter criteria or reveals filtered rows.
        if hidden && view.as_ref().is_some_and(|view| !view.rows().is_data_row_visible(row)) { continue; }
        if if hidden { after.hidden_rows.insert(row) } else { after.hidden_rows.remove(&row) } { count += 1; }
    }
    if count == 0 { return Ok(None); }
    if let Some(table) = sheet.table_view_spec().filter(|spec| spec.has_criteria())
        .and_then(|spec| sheet.tables().iter().find(|t| t.id == spec.table)) {
        if crate::table_filter_ui::desktop_layout_error(
            table, Some(&after.heights), Some(&after.hidden_rows), sheet.frozen_panes.0,
        ).is_some() { return Err("Clear this Table's sorting and filters before manually hiding its records.".into()); }
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
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered()) {
            self.status_message = Some("Clear worksheet sorting and filters before changing manual row visibility with Tables.".into());
            cx.notify();
            return;
        }
        let ((start, _), (end, _)) = self.selection_range();
        let id = self.cached_sheet_id();
        let before = self.structure_layout(id);
        let result = prepare(self.wb(cx), id, &before, start, end, hidden).and_then(|result| {
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
                self.history.record_action_with_provenance(UndoAction::TableStructureChanged {
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
            stack.record_action_with_provenance(UndoAction::TableStructureChanged {
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
    fn criteria_keep_body_layout_guard_but_allow_safe_rows_and_other_sheets() {
        let mut wb = fixture(true);
        let id = wb.active_sheet_id();
        let table = wb.active_sheet().tables()[0].id;
        wb.set_table_totals_visible(table, true, Default::default()).unwrap();
        let before = wb.saved_tables();
        assert!(prepare(&wb, id, &Default::default(), 3, 6, true).is_err());
        assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), serde_json::to_value(before).unwrap());
        let mut invalid_layout = StructureLayout::default();
        invalid_layout.hidden_rows.insert(4);
        assert!(prepare(&wb, id, &invalid_layout, 3, 6, false).unwrap_err().contains("Clear this Table"));
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
}
