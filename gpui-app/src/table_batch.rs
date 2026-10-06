//! Atomic automation publication and sparse multi-sheet undo with Table views.
use crate::{
    app::Spreadsheet,
    history::{MutationSource, UndoAction},
};
use gpui::Context;
use visigrid_engine::workbook::{GuardedStructureCommit, Workbook};

pub(crate) fn remap_sheet_view(
    view: &mut crate::workbook_view::WorkbookViewState,
    before: &[visigrid_engine::sheet::SheetId],
    after: &Workbook,
) {
    let index = before.get(view.active_sheet).and_then(|id| after.sheet_index_by_id(*id));
    if index.is_none() {
        let zoom = view.zoom_level;
        *view = Default::default();
        view.zoom_level = zoom;
    }
    view.active_sheet = index.unwrap_or(after.active_sheet_index());
    let (rows, cols) = after.sheets()[view.active_sheet].frozen_panes;
    view.frozen_rows = rows;
    view.frozen_cols = cols;
    view.scroll_row = view.scroll_row.max(rows);
    view.scroll_col = view.scroll_col.max(cols);
}

/// A rectangular Lua selection must not silently include filtered-out records.
pub(crate) fn canonical_script_selection(
    rows: &visigrid_engine::filter::RowView,
    start: (usize, usize),
    end: (usize, usize),
) -> Result<(usize, usize, usize, usize), String> {
    let mut selected: Vec<_> = rows
        .visible_rows()
        .iter()
        .copied()
        .filter(|r| *r >= start.0.min(end.0) && *r <= start.0.max(end.0))
        .map(|r| rows.view_to_data(r))
        .collect();
    selected.sort_unstable();
    let first = *selected
        .first()
        .ok_or("The script selection has no visible rows.")?;
    let last = *selected.last().unwrap();
    if last - first + 1 != selected.len() {
        return Err("This filtered selection is not a worksheet rectangle. Select one cell before running the script, and use explicit worksheet addresses.".into());
    }
    Ok((first, start.1.min(end.1), last, start.1.max(end.1)))
}

impl Spreadsheet {
    pub(crate) fn publish_table_batch(
        &mut self,
        candidate: Workbook,
        commit: GuardedStructureCommit,
        description: String,
        source: MutationSource,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        if self.cloud_live_enabled() {
            return Err("This live session supports sequenced cell values and formulas only.".into());
        }
        self.validate_saved_view_layout(&candidate)?;
        let changed = !commit.is_empty();
        let sheet_index = candidate.active_sheet_index();
        self.install_table_batch(&candidate, cx);
        if changed {
            self.history
                .record_named_range_action(UndoAction::TableBatchChanged {
                    sheet_index,
                    commit: Box::new(commit),
                    description,
                });
            self.history.retag_last_source(source);
            self.is_modified = true;
            self.cached_title = None;
        }
        Ok(())
    }

    pub(crate) fn install_table_batch(&mut self, candidate: &Workbook, cx: &mut Context<Self>) {
        let before: Vec<_> = self.wb(cx).sheets().iter().map(|s| s.id).collect();
        let changed_sheets = before != candidate.sheets().iter().map(|s| s.id).collect::<Vec<_>>();
        let same_active = self.wb(cx).active_sheet_id() == candidate.active_sheet_id();
        let (row, col) = self.active_view_state().selected;
        let data_row = self.row_view.view_to_data(row);
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(candidate));
        if changed_sheets {
            if !same_active {
                self.update_cached_sheet_id(cx);
                self.row_view = visigrid_engine::filter::RowView::new(crate::app::NUM_ROWS);
                self.filter_state = Default::default();
                self.table_view_installed = false;
            }
            remap_sheet_view(&mut self.view_state, &before, candidate);
            if let Some(pane) = &mut self.split_pane {
                remap_sheet_view(&mut pane.view_state, &before, candidate);
            }
            let active = self.active_view_state_mut();
            active.active_sheet = candidate.active_sheet_index();
            if !same_active {
                active.select_cell(0, 0);
                active.scroll_row = 0;
                active.scroll_col = 0;
            }
            let (rows, cols) = candidate.active_sheet().frozen_panes;
            active.frozen_rows = rows;
            active.frozen_cols = cols;
            active.scroll_row = active.scroll_row.max(rows);
            active.scroll_col = active.scroll_col.max(cols);
            self.history_highlight_range = None;
            self.formula_ref_cell = None;
            self.formula_ref_end = None;
        }
        self.sync_table_view(cx);
        let data_row = if same_active { data_row } else { 0 };
        let col = if same_active { col } else { 0 };
        let row = self.row_view.data_to_view(data_row).unwrap_or_else(|| {
            self.row_view
                .visible_rows()
                .iter()
                .copied()
                .min_by_key(|r| r.abs_diff(row))
                .unwrap_or(0)
        });
        self.active_view_state_mut().select_cell(row, col);
        self.active_view_state_mut().additional_selections.clear();
        self.table_filter_dropdown = None;
        self.table_edit_target = None;
        self.clipboard_visual_range = None;
        self.bump_cells_rev();
        self.bump_cf_rules_rev();
        cx.notify();
    }

    /// Publish a candidate with whole-workbook undo, for changes too large
    /// for the sparse guarded history (a big recipe refresh). Same install
    /// as a guarded batch; one undo step either way.
    pub(crate) fn publish_workbook_snapshot(
        &mut self,
        candidate: Workbook,
        description: String,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.validate_saved_view_layout(&candidate)?;
        let before = self.wb(cx).clone();
        let before_row_view = self.row_view.clone();
        self.install_table_batch(&candidate, cx);
        let after = self.wb(cx).clone();
        self.history.record_named_range_action(UndoAction::WorkbookSnapshot {
            commit: Box::new(crate::history::WorkbookSnapshotCommit::new(description, before, after)),
            before_row_view,
            after_row_view: self.row_view.clone(),
        });
        self.history.retag_last_source(MutationSource::Human);
        self.is_modified = true;
        self.cached_title = None;
        Ok(())
    }

    pub(crate) fn replay_table_batch(
        &mut self,
        commit: &GuardedStructureCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = commit.candidate(self.wb(cx), undo).and_then(|candidate| {
            self.validate_saved_view_layout(&candidate)?;
            Ok(candidate)
        });
        match result {
            Ok(candidate) => {
                self.install_table_batch(&candidate, cx);
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
        cell::{CellComment, CellStyle},
        sheet::{Sheet, SheetId},
    };
    use visigrid_protocol::Op;
    use visigrid_session_host::{ApplyOpsError, ApplyOpsRequest};

    fn request(ops: Vec<Op>) -> ApplyOpsRequest {
        ApplyOpsRequest {
            request_id: "batch-test".into(),
            batch_name: "Automation".into(),
            atomic: true,
            expected_revision: None,
            ops,
            client: Some("test-agent".into()),
        }
    }
    fn set(sheet: usize, row: usize, col: usize, value: &str) -> Op {
        Op::SetCellValue {
            sheet,
            row,
            col,
            value: value.into(),
        }
    }
    fn multi_sheet() -> Workbook {
        let mut wb = fixture(true);
        assert!(wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Controls")));
        wb
    }
    #[test]
    fn canonical_hidden_writes_and_multisheet_batch_replay_as_one_transaction() {
        let base = multi_sheet();
        let mut wb = base.clone();
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![
                set(0, 4, 2, "55"),
                set(1, 0, 0, "12"),
                set(0, 4, 2, "56"),
            ]),
        );
        assert!(out.response.error.is_none(), "{:?}", out.response.error);
        assert_eq!(out.response.applied, 3);
        assert_eq!(wb.revision(), base.revision() + 1);
        assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "56");
        assert_eq!(wb.sheet(0).unwrap().get_raw(5, 2), "20");
        assert!(!wb
            .sheet(0)
            .unwrap()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .rows()
            .is_data_row_visible(4));
        let commit = out.guarded_commit.unwrap();
        commit.replay(&mut wb, true).unwrap();
        assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "10");
        assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "");
        commit.replay(&mut wb, false).unwrap();
        assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "12");
        let mut history = History::new();
        history.record_named_range_action(UndoAction::TableBatchChanged {
            sheet_index: 0,
            commit: Box::new(commit),
            description: "Automation".into(),
        });
        let preview = history
            .build_workbook_before(1, Some(&base), 100, 10_000)
            .unwrap();
        assert_eq!(preview.workbook.sheet(0).unwrap().get_raw(4, 2), "56");
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
    }
    #[test]
    fn calculated_column_override_metadata_replays_with_the_batch() {
        let mut wb = fixture(true);
        let id = wb.active_sheet().tables()[0].id;
        wb.set_calculated_column(id, 3, 3, "=C4*2", true).unwrap();
        let out = visigrid_session_host::apply_ops(&mut wb, &request(vec![set(0, 4, 3, "999")]));
        assert!(out.response.error.is_none(), "{:?}", out.response.error);
        assert!(wb.active_sheet().is_calculated_exception(4, 3));
        let commit = out.guarded_commit.unwrap();
        assert!(!commit.is_empty());
        commit.replay(&mut wb, true).unwrap();
        assert!(!wb.active_sheet().is_calculated_exception(4, 3));
        assert_eq!(wb.active_sheet().get_raw(4, 3), "=C5*2");
        commit.replay(&mut wb, false).unwrap();
        assert!(wb.active_sheet().is_calculated_exception(4, 3));
        assert_eq!(wb.active_sheet().get_raw(4, 3), "999");
    }
    #[test]
    fn middle_header_failure_rejects_every_write_and_notification() {
        let mut wb = multi_sheet();
        let revision = wb.revision();
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![
                set(1, 0, 0, "12"),
                set(0, 2, 1, "Bad header"),
                set(0, 3, 2, "99"),
            ]),
        );
        assert!(matches!(out.response.error,Some(ApplyOpsError::OpFailed(ref e)) if e.op_index==1));
        assert_eq!(out.response.applied, 0);
        assert_eq!(wb.revision(), revision);
        assert!(out.guarded_commit.is_none());
        assert!(out.changed_cells.is_empty());
        assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "");
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "30");
    }
    #[test]
    fn late_cross_sheet_spill_rolls_back_values_formats_and_revision() {
        let mut wb = multi_sheet();
        wb.set_cell_value_tracked(1, 0, 0, "0");
        wb.set_cell_value_tracked(0, 0, 0, "=IF(Controls!A1=0,\"\",SEQUENCE(6))");
        let revision = wb.revision();
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![
                Op::SetStyle {
                    sheet: 0,
                    start_row: 3,
                    start_col: 2,
                    end_row: 3,
                    end_col: 2,
                    bold: Some(true),
                    italic: None,
                    underline: None,
                },
                set(1, 0, 0, "1"),
            ]),
        );
        assert!(out.response.error.is_some());
        assert_eq!(out.response.applied, 0);
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "0");
        assert!(!wb.sheet(0).unwrap().get_format(3, 2).bold);
    }
    #[test]
    fn mixed_clear_and_format_history_restores_exact_cell_metadata() {
        let mut wb = fixture(true);
        wb.active_sheet_mut().set_comment(
            5,
            3,
            Some(CellComment {
                text: "keep".into(),
                author: String::new(),
            }),
        );
        wb.active_sheet_mut().set_cell_style(5, 3, CellStyle::Input);
        let before = wb.active_sheet().get_cell(5, 3);
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![
                Op::ClearCell {
                    sheet: 0,
                    row: 5,
                    col: 3,
                },
                Op::SetStyle {
                    sheet: 0,
                    start_row: 2,
                    start_col: 1,
                    end_row: 2,
                    end_col: 1,
                    bold: Some(true),
                    italic: None,
                    underline: None,
                },
            ]),
        );
        assert!(out.response.error.is_none(), "{:?}", out.response.error);
        assert_eq!(wb.active_sheet().get_raw(5, 3), "");
        let commit = out.guarded_commit.unwrap();
        commit.replay(&mut wb, true).unwrap();
        let after = wb.active_sheet().get_cell(5, 3);
        assert_eq!(after.value.raw_display(), before.value.raw_display());
        assert_eq!(after.format, before.format);
        assert_eq!(after.comment(), before.comment());
        commit.replay(&mut wb, false).unwrap();
        assert!(wb.active_sheet().get_format(2, 1).bold);
    }
    #[test]
    fn formatting_only_batch_advances_revision_and_replays() {
        let mut wb = fixture(true);
        let revision = wb.revision();
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![Op::SetNumberFormat {
                sheet: 0,
                start_row: 4,
                start_col: 2,
                end_row: 4,
                end_col: 2,
                format: "currency:2".into(),
            }]),
        );
        assert!(out.response.error.is_none());
        assert_eq!(wb.revision(), revision + 1);
        out.guarded_commit.unwrap().replay(&mut wb, true).unwrap();
        assert_eq!(wb.active_sheet().get_format(4, 2), Default::default());
    }
    #[test]
    fn unsafe_adjacent_format_and_stale_revision_are_non_mutating() {
        let mut wb = fixture(true);
        let revision = wb.revision();
        let mut req = request(vec![set(0, 3, 2, "99")]);
        req.expected_revision = Some(revision - 1);
        assert!(matches!(
            visigrid_session_host::apply_ops(&mut wb, &req)
                .response
                .error,
            Some(ApplyOpsError::RevisionMismatch { .. })
        ));
        let out = visigrid_session_host::apply_ops(
            &mut wb,
            &request(vec![
                set(0, 3, 2, "99"),
                Op::SetStyle {
                    sheet: 0,
                    start_row: 3,
                    start_col: 4,
                    end_row: 3,
                    end_col: 4,
                    bold: Some(true),
                    italic: None,
                    underline: None,
                },
            ]),
        );
        assert!(out.response.error.is_some());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.active_sheet().get_raw(3, 2), "30");
    }
    #[test]
    fn script_selection_maps_records_and_rejects_rectangles_with_hidden_holes() {
        let wb = fixture(true);
        let view = wb
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        let slot = view.rows().data_to_view(5).unwrap();
        assert_eq!(
            canonical_script_selection(view.rows(), (slot, 2), (slot, 2)).unwrap(),
            (5, 2, 5, 2)
        );
        assert!(canonical_script_selection(view.rows(), (3, 1), (6, 3)).is_err());
    }
    #[test]
    fn lua_batch_updates_hidden_records_and_preserves_atomic_header_refusal() {
        use crate::scripting::{LuaRuntime, SheetSnapshot};
        let mut wb = fixture(true);
        let runtime = LuaRuntime::new().unwrap();
        let result = runtime.eval_with_sheet(
            "sheet:set('C5',55); sheet:set('C6',21)",
            Box::new(SheetSnapshot::from_sheet(wb.active_sheet())),
        );
        assert!(result.error.is_none(), "{:?}", result.error);
        crate::views::lua_console::apply_captured_lua_ops(&mut wb, 0, &result.ops).unwrap();
        assert_eq!(wb.active_sheet().get_raw(4, 2), "55");
        assert_eq!(wb.active_sheet().get_raw(5, 2), "21");
        let result = runtime.eval_with_sheet(
            "sheet:set('C5',88); sheet:set('B3','Header'); sheet:set('C6',99)",
            Box::new(SheetSnapshot::from_sheet(wb.active_sheet())),
        );
        let revision = wb.revision();
        assert!(
            crate::views::lua_console::apply_captured_lua_ops(&mut wb, 0, &result.ops).is_err()
        );
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.active_sheet().get_raw(4, 2), "55");
    }
    #[test]
    fn lua_late_spill_and_format_only_revision_are_guarded() {
        use crate::scripting::{LuaCellValue, LuaOp};
        let mut wb = multi_sheet();
        wb.set_cell_value_tracked(1, 0, 0, "0");
        wb.set_cell_value_tracked(0, 0, 0, "=IF(Controls!A1=0,\"\",SEQUENCE(6))");
        let ops = [LuaOp::SetValue {
            row: 0,
            col: 0,
            value: LuaCellValue::Number(1.0),
        }];
        let revision = wb.revision();
        assert!(crate::views::lua_console::apply_captured_lua_ops(&mut wb, 1, &ops).is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "0");
        let ops = [LuaOp::SetCellStyle {
            r1: 3,
            c1: 2,
            r2: 3,
            c2: 2,
            style: CellStyle::Input.to_int() as u8,
        }];
        crate::views::lua_console::apply_captured_lua_ops(&mut wb, 0, &ops).unwrap();
        assert!(wb.revision() > revision);
    }
}
