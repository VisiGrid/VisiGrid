//! Desktop mapping and layout guards for atomic structural edits with Table views.
use crate::{
    app::{Spreadsheet, NUM_COLS, NUM_ROWS},
    history::UndoAction,
};
use gpui::Context;
use std::collections::{BTreeSet, HashMap};
use visigrid_engine::{
    filter::RowView,
    sheet::{Sheet, SheetId},
    structural::Axis,
    workbook::{shift_structure_index, GuardedStructureCommit, StructureStep, Workbook},
};

#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct StructureLayout {
    pub heights: HashMap<usize, f32>,
    pub widths: HashMap<usize, f32>,
    pub hidden_rows: BTreeSet<usize>,
    pub hidden_cols: BTreeSet<usize>,
}
impl StructureLayout {
    pub(crate) fn shifted(&self, sheet: &Sheet, steps: &[StructureStep]) -> Result<Self, String> {
        let mut next = self.clone();
        for &step in steps {
            let limit = if step.axis == Axis::Row {
                sheet.rows
            } else {
                sheet.cols
            };
            if step.count == 0
                || step
                    .at
                    .checked_add(step.count)
                    .is_none_or(|end| end > limit)
            {
                return Err("Structural edit exceeds the worksheet boundary.".into());
            }
            let (sizes, hidden) = if step.axis == Axis::Row {
                (&mut next.heights, &mut next.hidden_rows)
            } else {
                (&mut next.widths, &mut next.hidden_cols)
            };
            if !step.delete
                && sizes
                    .keys()
                    .chain(hidden.iter())
                    .any(|i| *i >= limit - step.count)
            {
                return Err(
                    "The insertion would push row or column layout off the worksheet.".into(),
                );
            }
            *sizes = sizes
                .iter()
                .filter_map(|(&i, &v)| shift_structure_index(i, step, limit).map(|j| (j, v)))
                .collect();
            *hidden = hidden
                .iter()
                .filter_map(|&i| shift_structure_index(i, step, limit))
                .collect();
        }
        Ok(next)
    }
}
/// Sparse layout only: ordinary row history must not capture every shifted cell.
#[derive(Clone, Debug)]
pub(crate) struct RowLayoutHistory {
    pub sheet: SheetId,
    pub before: StructureLayout,
    pub after: StructureLayout,
}
impl RowLayoutHistory {
    pub(crate) fn capture(mut before: StructureLayout, sheet: &Sheet, step: StructureStep) -> Result<Self, String> {
        before.hidden_rows = sheet.manual_hidden_rows();
        let after = before.shifted(sheet, &[step])?;
        Ok(Self { sheet: sheet.id, before, after })
    }

    pub(crate) fn apply(&self, wb: &mut Workbook, sheet_index: usize, undo: bool) -> Result<&StructureLayout, String> {
        let layout = if undo { &self.before } else { &self.after };
        let sheet = wb.sheet_mut(sheet_index).ok_or("Sheet no longer exists")?;
        if sheet.id != self.sheet { return Err("Row history belongs to another sheet".into()); }
        let visibility_changed = sheet.manual_hidden_rows() != layout.hidden_rows;
        if visibility_changed {
            sheet.set_manual_hidden_rows(layout.hidden_rows.clone())?;
        }
        // Plain row undo restores cells directly. Rebuild dependencies before
        // recalculating SUBTOTAL and readers on other sheets.
        if undo && visibility_changed {
            wb.rebuild_dep_graph();
            wb.recompute_full_ordered();
        }
        Ok(layout)
    }
}
#[derive(Clone, Debug)]
pub(crate) struct TableStructureHistory {
    pub commit: GuardedStructureCommit,
    /// Desktop-only frozen panes normalized into a reviewed commit's source.
    pub source_frozen: Option<(usize, usize)>,
    pub(crate) before: StructureLayout,
    pub(crate) after: StructureLayout,
}
impl TableStructureHistory {
    pub(crate) fn estimated_history_bytes(&self) -> usize {
        self.commit.estimated_history_bytes().saturating_add(visigrid_engine::history_size::estimated_debug_bytes(&(&self.before, &self.after)))
    }
}
/// Resolve selected visible slots once. Deletions are coalesced and performed
/// bottom-up, so neither hidden records nor newly shifted records are deleted.
pub(crate) fn selected_row_steps(
    rows: &RowView,
    start: usize,
    end: usize,
    delete: bool,
) -> Result<Vec<StructureStep>, String> {
    if start > end || end >= rows.row_count() {
        return Err("The row selection is outside the worksheet.".into());
    }
    let selected: Vec<_> = rows
        .visible_rows()
        .iter()
        .copied()
        .filter(|r| *r >= start && *r <= end)
        .map(|r| rows.view_to_data(r))
        .collect();
    if selected.is_empty() || selected.len() > 100_000 {
        return Err("Select between 1 and 100,000 visible rows.".into());
    }
    if !delete {
        return Ok(vec![StructureStep {
            axis: Axis::Row,
            at: selected[0],
            count: selected.len(),
            delete: false,
        }]);
    }
    let mut data = selected;
    data.sort_unstable();
    data.dedup();
    let mut spans: Vec<StructureStep> = Vec::new();
    for row in data {
        if let Some(last) = spans.last_mut().filter(|s| s.at + s.count == row) {
            last.count += 1;
        } else {
            spans.push(StructureStep {
                axis: Axis::Row,
                at: row,
                count: 1,
                delete: true,
            });
        }
    }
    spans.reverse();
    Ok(spans)
}
impl Spreadsheet {
    pub(crate) fn structure_layout(&self, id: SheetId) -> StructureLayout {
        StructureLayout {
            heights: self.row_heights.get(&id).cloned().unwrap_or_default(),
            widths: self.col_widths.get(&id).cloned().unwrap_or_default(),
            hidden_rows: self.hidden_rows.get(&id).cloned().unwrap_or_default(),
            hidden_cols: self.hidden_cols.get(&id).cloned().unwrap_or_default(),
        }
    }
    pub(crate) fn install_structure_layout(&mut self, id: SheetId, layout: &StructureLayout) {
        self.row_heights.insert(id, layout.heights.clone());
        self.col_widths.insert(id, layout.widths.clone());
        self.hidden_rows.insert(id, layout.hidden_rows.clone());
        self.hidden_cols.insert(id, layout.hidden_cols.clone());
    }
    pub(crate) fn replay_row_layout(&mut self, history: Option<&RowLayoutHistory>, sheet_index: usize, undo: bool, cx: &mut Context<Self>) {
        let Some(history) = history else { return; };
        let result = self.workbook.update(cx, |wb, _| history.apply(wb, sheet_index, undo).map(|_| ()));
        if let Err(error) = result {
            self.status_message = Some(error);
            return;
        }
        if let Some(sheet) = self.wb(cx).sheet(sheet_index) {
            self.install_structure_layout(sheet.id, if undo { &history.before } else { &history.after });
        }
    }
    pub(crate) fn validate_structure_layout(
        &self,
        wb: &Workbook,
        id: SheetId,
        layout: &StructureLayout,
    ) -> Result<(), String> {
        for sheet in wb.sheets() {
            if let Some(table) = sheet
                .table_view_spec()
                .filter(|v| v.has_criteria())
                .and_then(|s| sheet.tables().iter().find(|t| t.id == s.table))
            {
                let (heights, hidden) = if sheet.id == id {
                    (Some(&layout.heights), Some(&layout.hidden_rows))
                } else {
                    (
                        self.row_heights.get(&sheet.id),
                        self.hidden_rows.get(&sheet.id),
                    )
                };
                if let Some(error) = crate::table_filter_ui::desktop_layout_error(
                    table,
                    heights,
                    hidden,
                    sheet.frozen_panes.0,
                ) {
                    return Err(error);
                }
            }
        }
        Ok(())
    }
    pub(crate) fn table_structure_selection(&mut self, delete: bool, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        if self.mode.is_editing() || !self.view_state.additional_selections.is_empty() {
            self.status_message =
                Some("Finish editing and select one row or column range first.".into());
            cx.notify();
            return;
        }
        self.sync_table_view(cx);
        let ((r0, c0), (r1, c1)) = self.selection_range();
        let steps = if self.is_row_selection() {
            selected_row_steps(&self.row_view, r0, r1, delete)
        } else if self.is_col_selection() {
            Ok(vec![StructureStep {
                axis: Axis::Col,
                at: c0,
                count: c1 - c0 + 1,
                delete,
            }])
        } else {
            Err("Select entire rows (Shift+Space) or columns (Ctrl+Space) first.".into())
        };
        match steps {
            Ok(steps) => {
                self.apply_table_structure(steps, cx);
            }
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
            }
        }
    }
    pub(crate) fn apply_table_structure(
        &mut self,
        steps: Vec<StructureStep>,
        cx: &mut Context<Self>,
    ) -> bool {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return false;
        }
        self.sync_table_view(cx);
        if self.mode.is_editing() || !self.view_state.additional_selections.is_empty() {
            self.status_message =
                Some("Finish editing and select one row or column range first.".into());
            cx.notify();
            return false;
        }
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            self.status_message = Some(
                "Clear this sheet's worksheet sorting and filters before structural edits.".into(),
            );
            cx.notify();
            return false;
        }
        let index = self.sheet_index(cx);
        let id = self.sheet(cx).id;
        let before = self.structure_layout(id);
        let result = (|| {
            let mut base = self.wb(cx).clone();
            base.sheet_mut(index).unwrap().frozen_panes =
                (self.view_state.frozen_rows, self.view_state.frozen_cols);
            self.validate_structure_layout(&base, id, &before)?;
            let after = before.shifted(base.sheet(index).unwrap(), &steps)?;
            let (candidate, commit) = base.prepare_guarded_structure(index, steps)?;
            self.validate_structure_layout(&candidate, id, &after)?;
            Ok::<_, String>((
                candidate,
                TableStructureHistory {
                    source_frozen: None,
                    commit,
                    before,
                    after,
                },
            ))
        })();
        let (candidate, history) = match result {
            Ok(v) => v,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return false;
            }
        };
        let first = history.commit.steps[0];
        let count: usize = history.commit.steps.iter().map(|s| s.count).sum();
        let description = format!(
            "{} {} visible {}(s)",
            if first.delete { "Delete" } else { "Insert" },
            count,
            if first.axis == Axis::Row {
                "row"
            } else {
                "column"
            }
        );
        use crate::repeat::RepeatAction;
        self.set_repeat(match (first.axis, first.delete) {
            (Axis::Row, false) => RepeatAction::InsertRows(count),
            (Axis::Row, true) => RepeatAction::DeleteRows(count),
            (Axis::Col, false) => RepeatAction::InsertCols(count),
            (Axis::Col, true) => RepeatAction::DeleteCols(count),
        });
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.install_structure_layout(id, &history.after);
        self.finish_table_structure(index, first.axis, count, cx);
        self.history
            .record_named_range_action(UndoAction::TableStructureChanged {
                sheet_index: index,
                history: Box::new(history),
                description: description.clone(),
            });
        self.status_message = Some(description);
        self.is_modified = true;
        cx.notify();
        true
    }
    fn finish_table_structure(
        &mut self,
        index: usize,
        axis: Axis,
        count: usize,
        cx: &mut Context<Self>,
    ) {
        if self.sheet_index(cx) != index {
            self.activate_sheet(index, cx);
        }
        self.view_state.frozen_rows = self.sheet(cx).frozen_panes.0;
        self.view_state.frozen_cols = self.sheet(cx).frozen_panes.1;
        self.sync_table_view(cx);
        if !self.table_view_installed {
            self.row_view = RowView::new(NUM_ROWS.min(self.sheet(cx).rows));
        }
        let old = self.view_state.selected;
        let row = self
            .row_view
            .visible_rows()
            .iter()
            .copied()
            .min_by_key(|r| r.abs_diff(old.0))
            .unwrap_or(0);
        if axis == Axis::Row {
            let first = self.row_view.visible_index_of(row).unwrap_or(0);
            let last = self
                .row_view
                .visible_rows()
                .get(first + count.saturating_sub(1))
                .copied()
                .unwrap_or_else(|| self.row_view.visible_rows().last().copied().unwrap_or(row));
            self.view_state.select_cell(row, 0);
            self.view_state.selection_end = Some((last, NUM_COLS - 1));
        } else {
            let col = old.1.min(NUM_COLS - 1);
            self.view_state.select_cell(0, col);
            self.view_state.selection_end = Some((
                NUM_ROWS - 1,
                (col + count.saturating_sub(1)).min(NUM_COLS - 1),
            ));
        }

        self.table_filter_dropdown = None;
        self.table_edit_target = None;
        self.table_fill_revision = None;
        self.clipboard_visual_range = None;
        self.bump_cells_rev();
        self.ensure_visible(cx);
        cx.notify();
    }
    pub(crate) fn replay_table_structure(
        &mut self,
        history: &TableStructureHistory,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let id = history.commit.sheet;
        let result = (|| {
            let index = self
                .wb(cx)
                .sheet_index_by_id(id)
                .ok_or("The history sheet no longer exists.")?;
            let (expected, target) = if undo {
                (&history.after, &history.before)
            } else {
                (&history.before, &history.after)
            };
            if &self.structure_layout(id) != expected {
                return Err("Row or column layout changed. Undo/redo was not applied.".into());
            }
            let candidate = history.commit.candidate(self.wb(cx), undo)?;
            self.validate_structure_layout(&candidate, id, target)?;
            Ok::<_, String>((index, candidate, target.clone()))
        })();
        match result {
            Ok((index, candidate, layout)) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.install_structure_layout(id, &layout);
                if let Some(first) = history.commit.steps.first() {
                    self.finish_table_structure(
                        index, first.axis, history.commit.steps.iter().map(|s| s.count).sum(), cx,
                    );
                } else {
                    // Visibility changes share guarded workbook/layout history,
                    // but do not insert/delete rows or replace the selection.
                    if self.sheet_index(cx) != index { self.activate_sheet(index, cx); }
                    self.sync_table_view(cx);
                    self.table_filter_dropdown = None;
                    self.bump_cells_rev();
                    self.ensure_visible(cx);
                    cx.notify();
                }
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
    use crate::history::History;
    use crate::table_cell_history::TableCellsCommit;
    use crate::table_edit::{prepare_table_writes, tests::fixture, TableCellWrite};
    use visigrid_engine::{
        cell::{Cell, CellComment, CellValue},
        sheet::MergedRegion,
    };
    fn step(axis: Axis, at: usize, count: usize, delete: bool) -> StructureStep {
        StructureStep {
            axis,
            at,
            count,
            delete,
        }
    }
    fn row_view(wb: &Workbook) -> RowView {
        wb.active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .rows()
            .clone()
    }
    fn record_values(wb: &Workbook) -> Vec<String> {
        let s = wb.active_sheet();
        let t = &s.tables()[0];
        let view = row_view(wb);
        view.visible_rows()
            .iter()
            .copied()
            .filter(|r| *r > t.range.start_row && *r <= t.range.end_row)
            .map(|r| s.get_raw(view.view_to_data(r), t.range.start_col + 1))
            .collect()
    }
    #[test]
    fn visible_delete_is_atomic_skips_hidden_records_and_undo_restores_exact_cells() {
        let mut before = fixture(true);
        before.active_sheet_mut().set_comment(
            5,
            2,
            Some(CellComment {
                text: "Keep".into(),
                author: "R".into(),
            }),
        );
        let rows = row_view(&before);
        let steps = selected_row_steps(&rows, 3, 6, true).unwrap();
        assert_eq!(
            steps.iter().map(|s| (s.at, s.count)).collect::<Vec<_>>(),
            vec![(5, 2), (3, 1)]
        );
        let (mut after, commit) = before.prepare_guarded_structure(0, steps).unwrap();
        assert_eq!(after.active_sheet().tables()[0].range.end_row, 3);
        assert_eq!(after.active_sheet().get_raw(3, 1), "East");
        assert!(record_values(&after).is_empty());
        assert_eq!(after.active_sheet().get_display(0, 1), "10");
        assert_eq!(before.active_sheet().get_raw(5, 2), "20");
        commit.replay(&mut after, true).unwrap();
        assert_eq!(record_values(&after), vec!["20", "30", "40"]);
        assert_eq!(
            after.active_sheet().comment(5, 2),
            before.active_sheet().comment(5, 2)
        );
        commit.replay(&mut after, false).unwrap();
        assert!(record_values(&after).is_empty());
    }
    #[test]
    fn sorted_insert_counts_visible_rows_and_keeps_calculated_rules_and_overrides() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        before.set_cell_value_tracked(0, 5, 3, "999");
        let steps = selected_row_steps(&row_view(&before), 3, 6, false).unwrap();
        assert_eq!((steps[0].at, steps[0].count), (5, 3));
        let (mut after, c) = before.prepare_guarded_structure(0, steps).unwrap();
        assert_eq!(after.active_sheet().tables()[0].range.end_row, 9);
        for r in 5..8 {
            assert_eq!(after.active_sheet().get_display(r, 3), "0");
        }
        assert_eq!(after.active_sheet().get_raw(8, 3), "999");
        assert_eq!(record_values(&after), vec!["20", "30", "40"]);
        c.replay(&mut after, true).unwrap();
        assert_eq!(after.active_sheet().get_raw(5, 3), "999");
        c.replay(&mut after, false).unwrap();
        assert_eq!(after.active_sheet().get_raw(8, 3), "999");
    }
    #[test]
    fn column_insert_and_delete_preserve_ids_criteria_and_permanent_reference_rewrites() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 10, 0, "=SUM(Sales[Result])");
        let original = before.active_sheet().tables()[0].clone();
        let (mut inserted, c) = before
            .prepare_guarded_structure(0, vec![step(Axis::Col, 2, 1, false)])
            .unwrap();
        let new_id = inserted.active_sheet().tables()[0].columns[1].id;
        assert_eq!(inserted.active_sheet().get_raw(2, 2), "Column1");
        assert_eq!(
            inserted.active_sheet().tables()[0].columns[2].id,
            original.columns[1].id
        );
        assert_eq!(
            inserted.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        c.replay(&mut inserted, true).unwrap();
        assert_eq!(
            inserted.active_sheet().tables()[0].columns,
            original.columns
        );
        c.replay(&mut inserted, false).unwrap();
        assert_eq!(inserted.active_sheet().tables()[0].columns[1].id, new_id);
        let (mut deleted, d) = before
            .prepare_guarded_structure(0, vec![step(Axis::Col, 3, 1, true)])
            .unwrap();
        assert!(deleted.active_sheet().get_raw(10, 0).contains("#REF!"));
        d.replay(&mut deleted, true).unwrap();
        assert_eq!(deleted.active_sheet().get_raw(10, 0), "=SUM(Sales[Result])");
        assert_eq!(deleted.active_sheet().tables()[0].columns, original.columns);
        d.replay(&mut deleted, false).unwrap();
        for col in [1, 2] {
            assert!(before
                .prepare_guarded_structure(0, vec![step(Axis::Col, col, 1, true)])
                .is_err());
        }
    }
    #[test]
    fn failed_late_step_header_or_grid_edge_preserves_original_workbook() {
        let mut b = fixture(true);
        let rev = b.revision();
        assert!(b
            .prepare_guarded_structure(
                0,
                vec![step(Axis::Row, 6, 1, true), step(Axis::Row, 2, 1, true)]
            )
            .is_err());
        assert_eq!(b.revision(), rev);
        assert_eq!(b.active_sheet().get_raw(6, 2), "40");
        b.active_sheet_mut().set_comment(
            29,
            0,
            Some(CellComment {
                text: "edge".into(),
                author: String::new(),
            }),
        );
        assert!(b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 0, 1, false)])
            .is_err());
        assert!(b
            .prepare_guarded_structure(0, vec![step(Axis::Col, 1, 3, true)])
            .is_err());
        assert!(b
            .prepare_guarded_structure(0, vec![step(Axis::Row, usize::MAX, 1, true)])
            .is_err());
    }
    #[test]
    fn layouts_shift_and_restore_without_losing_hidden_rows_or_custom_sizes() {
        let b = fixture(true);
        let mut layout = StructureLayout::default();
        layout.heights.insert(9, 42.0);
        layout.hidden_rows.insert(10);
        layout.widths.insert(6, 150.0);
        layout.hidden_cols.insert(7);
        let shifted = layout
            .shifted(b.active_sheet(), &[step(Axis::Row, 0, 2, false)])
            .unwrap();
        assert_eq!(shifted.heights.get(&11), Some(&42.0));
        assert!(shifted.hidden_rows.contains(&12));
        let shifted = shifted
            .shifted(b.active_sheet(), &[step(Axis::Row, 0, 2, true)])
            .unwrap();
        assert_eq!(shifted, layout);
        assert!(layout
            .shifted(b.active_sheet(), &[step(Axis::Col, 0, 1, false)])
            .is_err());
        assert!(layout
            .shifted(b.active_sheet(), &[step(Axis::Row, 0, usize::MAX, false)])
            .is_err());
    }
    #[test]
    fn ordinary_sheet_hidden_rows_shift_with_layout_and_restore_deleted_hides() {
        let mut before = Workbook::new();
        before.active_sheet_mut().set_manual_hidden_rows([2, 8].into()).unwrap();
        let layout = StructureLayout { heights: [(2, 40.0)].into(), ..Default::default() };
        for delete in [false, true] {
            let history = RowLayoutHistory::capture(layout.clone(), before.active_sheet(), step(Axis::Row, 2, 1, delete)).unwrap();
            let mut candidate = before.clone();
            candidate.structural_edit(0, Axis::Row, 2, 1, delete).unwrap();
            assert_eq!(candidate.active_sheet().manual_hidden_rows(), history.after.hidden_rows);
            assert_eq!(history.after.hidden_rows, if delete { [7].into() } else { [3, 9].into() });
            if delete { candidate.active_sheet_mut().insert_rows(2, 1); }
            else { candidate.active_sheet_mut().delete_rows(2, 1); }
            history.apply(&mut candidate, 0, true).unwrap();
            assert_eq!(candidate.active_sheet().manual_hidden_rows(), history.before.hidden_rows);
            assert_eq!(history.before.heights.get(&2), Some(&40.0));
            candidate.structural_edit(0, Axis::Row, 2, 1, delete).unwrap();
            history.apply(&mut candidate, 0, false).unwrap();
            assert_eq!(candidate.active_sheet().manual_hidden_rows(), history.after.hidden_rows);
        }
    }
    #[test]
    fn structural_history_does_not_break_earlier_cell_undo_after_column_ids_advance() {
        let before = fixture(true);
        let writes = [TableCellWrite::value(8, 0, "note".into())];
        let edited = prepare_table_writes(&before, 0, &writes).unwrap();
        let cells =
            TableCellsCommit::capture(before.active_sheet(), edited.active_sheet(), [(8, 0)]);
        let (mut inserted, structure) = edited
            .prepare_guarded_structure(0, vec![step(Axis::Col, 2, 1, false)])
            .unwrap();
        structure.replay(&mut inserted, true).unwrap();
        cells.replay(&mut inserted, true).unwrap();
        assert_eq!(inserted.active_sheet().get_raw(8, 0), "");
        cells.replay(&mut inserted, false).unwrap();
        structure.replay(&mut inserted, false).unwrap();
    }
    #[test]
    fn deleted_literal_cells_restore_type_style_frozen_formula_and_merge_topology() {
        let mut b = fixture(true);
        let mut cell = Cell::default();
        cell.value = CellValue::Text("=literal".into());
        cell.set_style_id(Some(9));
        cell.set_frozen_formula(Some("=REMOTE()".into()));
        b.restore_cell_tracked(0, 9, 0, Some(cell)).unwrap();
        b.active_sheet_mut()
            .add_merge(MergedRegion {
                start: (9, 0),
                end: (10, 1),
            })
            .unwrap();
        b.active_sheet_mut()
            .row_formats
            .insert(9, Default::default());
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 9, 2, true)])
            .unwrap();
        assert!(a.active_sheet().merged_regions.is_empty());
        c.replay(&mut a, true).unwrap();
        assert_eq!(a.active_sheet().get_raw(9, 0), "=literal");
        assert!(!a.active_sheet().get_cell(9, 0).value().is_formula());
        assert_eq!(a.active_sheet().get_cell(9, 0).style_id(), Some(9));
        assert_eq!(
            a.active_sheet().get_cell(9, 0).frozen_formula(),
            Some("=REMOTE()")
        );
        assert_eq!(
            a.active_sheet().merged_regions,
            b.active_sheet().merged_regions
        );
        assert_eq!(a.active_sheet().row_formats, b.active_sheet().row_formats);
    }
    #[test]
    fn other_sheet_structure_rewrites_dependents_and_unsafe_new_spills_are_atomic() {
        let mut b = fixture(true);
        assert!(b.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Controls")));
        b.set_cell_value_tracked(1, 0, 0, "30");
        b.set_cell_value_tracked(0, 3, 2, "=Controls!A1");
        let (mut a, c) = b
            .prepare_guarded_structure(1, vec![step(Axis::Row, 0, 1, false)])
            .unwrap();
        assert_eq!(a.sheet(0).unwrap().get_raw(3, 2), "=Controls!A2");
        c.replay(&mut a, true).unwrap();
        assert_eq!(a.sheet(0).unwrap().get_raw(3, 2), "=Controls!A1");
        // Deleting one of the counted outside rows grows a formerly safe spill.
        b.set_cell_value_tracked(0, 0, 0, "=SEQUENCE((3-ROWS(F10:F12))*5+1)");
        assert!(b.active_sheet().get_spill_parent(5, 0).is_none());
        let rev = b.revision();
        assert!(b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 10, 1, true)])
            .is_err());
        assert_eq!(b.revision(), rev);
    }
    #[test]
    fn stale_structural_undo_refuses_new_cells_and_changed_criteria() {
        let b = fixture(true);
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 0, 1, false)])
            .unwrap();
        a.set_cell_value_tracked(0, 20, 0, "new");
        let rev = a.revision();
        assert!(c.replay(&mut a, true).is_err());
        assert_eq!(a.revision(), rev);
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Col, 3, 1, true)])
            .unwrap();
        let mut spec = a.active_sheet().table_view_spec().unwrap().clone();
        spec.filters.clear();
        a.set_table_view_spec(a.active_sheet_id(), Some(spec))
            .unwrap();
        assert!(c.replay(&mut a, true).is_err());
    }
    #[test]
    fn rewind_recovers_earlier_layout_and_preserves_unsupported_visibility_gate() {
        let b = fixture(true);
        let mut layout = StructureLayout::default();
        layout.widths.insert(5, 120.0);
        layout.hidden_cols.insert(6);
        let steps = vec![step(Axis::Col, 3, 1, false)];
        let (_, commit) = b.prepare_guarded_structure(0, steps.clone()).unwrap();
        let after = layout.shifted(b.active_sheet(), &steps).unwrap();
        let mut h = History::new();
        h.record_named_range_action(UndoAction::ColumnWidthSet {
            sheet_id: b.active_sheet_id(),
            col: 5,
            old: None,
            new: Some(120.0),
        });
        h.record_named_range_action(UndoAction::ColVisibilityChanged {
            sheet_id: b.active_sheet_id(),
            cols: vec![6],
            hidden: true,
        });
        h.record_named_range_action(UndoAction::TableStructureChanged {
            sheet_index: 0,
            history: Box::new(TableStructureHistory {
                source_frozen: None,
                commit,
                before: layout,
                after,
            }),
            description: "Insert column".into(),
        });
        h.record_named_range_action(UndoAction::ColVisibilityChanged {
            sheet_id: b.active_sheet_id(),
            cols: vec![7],
            hidden: false,
        });
        let preview = h.build_workbook_before(0, Some(&b), 100, 10_000).unwrap();
        assert_eq!(
            preview.view_state.per_sheet[0].structure_layout,
            Some(StructureLayout::default())
        );
        assert!(matches!(
            h.build_workbook_before(4, Some(&b), 100, 10_000),
            Err(crate::history::PreviewBuildError::UnsupportedAction(
                crate::history::UndoActionKind::ColVisibilityChanged
            ))
        ));
    }
    #[test]
    fn row_and_column_moves_before_table_preserve_freeze_and_rewind_layout() {
        let mut b = fixture(true);
        b.active_sheet_mut().frozen_panes = (2, 1);
        let steps = vec![step(Axis::Row, 0, 1, false)];
        let (mut a, c) = b.prepare_guarded_structure(0, steps.clone()).unwrap();
        assert_eq!(a.active_sheet().frozen_panes, (3, 1));
        assert_eq!(a.active_sheet().tables()[0].range.start_row, 3);
        c.replay(&mut a, true).unwrap();
        assert_eq!(a.active_sheet().frozen_panes, (2, 1));
        let mut before = StructureLayout::default();
        before.heights.insert(9, 40.0);
        let after = before.shifted(b.active_sheet(), &steps).unwrap();
        let mut history = History::new();
        history.record_named_range_action(UndoAction::TableStructureChanged {
            sheet_index: 0,
            history: Box::new(TableStructureHistory {
                source_frozen: None,
                commit: c,
                before: before.clone(),
                after: after.clone(),
            }),
            description: "Insert row".into(),
        });
        let preview = history
            .build_workbook_before(1, Some(&b), 100, 10_000)
            .unwrap();
        assert_eq!(
            preview.view_state.per_sheet[0].structure_layout,
            Some(after)
        );
        assert_eq!(
            preview.workbook.active_sheet().tables()[0].range.start_row,
            3
        );
        let original = history
            .build_workbook_before(0, Some(&b), 100, 10_000)
            .unwrap();
        assert_eq!(
            original.view_state.per_sheet[0].structure_layout,
            Some(before)
        );
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
    }
}

#[cfg(test)]
mod metadata_tests {
    use super::*;
    use crate::table_edit::tests::fixture;
    use visigrid_engine::validation::{CellRange, ListSource, ValidationRule, ValidationType};
    fn step(axis: Axis, at: usize, count: usize, delete: bool) -> StructureStep {
        StructureStep {
            axis,
            at,
            count,
            delete,
        }
    }
    #[test]
    fn deleting_every_record_preserves_dormant_criteria_and_undo_reactivates_them() {
        let b = fixture(false);
        let spec = b.active_sheet().table_view_spec().cloned();
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 3, 4, true)])
            .unwrap();
        assert_eq!(a.active_sheet().tables()[0].range.data_rows(), 0);
        assert_eq!(a.active_sheet().table_view_spec(), spec.as_ref());
        assert!(a
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .is_none());
        c.replay(&mut a, true).unwrap();
        assert!(a
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .is_some());
        c.replay(&mut a, false).unwrap();
        assert!(a
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .is_none());
    }
    #[test]
    fn named_ranges_validations_and_print_ranges_roundtrip_through_clipping() {
        let mut b = fixture(true);
        b.define_name_for_range("Notes", 0, 9, 0, 12, 1).unwrap();
        b.active_sheet_mut().validations.set(
            CellRange {
                start_row: 9,
                start_col: 0,
                end_row: 12,
                end_col: 0,
            },
            ValidationRule::new(ValidationType::List(ListSource::Inline(vec![
                "yes".into(),
                "no".into(),
            ]))),
        );
        b.active_sheet_mut().print_setup.area = Some(visigrid_engine::print_setup::PrintArea {
            start_row: 9,
            start_col: 0,
            end_row: 12,
            end_col: 1,
        });
        let print = b.active_sheet().print_setup.clone();
        let names = serde_json::to_value(b.named_ranges()).unwrap();
        let rules =
            serde_json::to_value(b.active_sheet().validations.iter().collect::<Vec<_>>()).unwrap();
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 9, 2, true)])
            .unwrap();
        assert_ne!(serde_json::to_value(a.named_ranges()).unwrap(), names);
        assert_ne!(a.active_sheet().print_setup, print);
        assert_ne!(
            serde_json::to_value(a.active_sheet().validations.iter().collect::<Vec<_>>()).unwrap(),
            rules
        );
        c.replay(&mut a, true).unwrap();
        assert_eq!(serde_json::to_value(a.named_ranges()).unwrap(), names);
        assert_eq!(a.active_sheet().print_setup, print);
        assert_eq!(
            serde_json::to_value(a.active_sheet().validations.iter().collect::<Vec<_>>()).unwrap(),
            rules
        );
        c.replay(&mut a, false).unwrap();
    }
    #[test]
    fn a_new_cross_sheet_reference_or_named_range_refuses_structural_undo() {
        let mut b = fixture(true);
        assert!(b.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Other")));
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Row, 0, 1, false)])
            .unwrap();
        a.set_cell_value_tracked(1, 0, 0, "=Sheet1!C5");
        let rev = a.revision();
        assert!(c.replay(&mut a, true).is_err());
        assert_eq!(a.revision(), rev);
        let (mut a, c) = b
            .prepare_guarded_structure(0, vec![step(Axis::Col, 3, 1, false)])
            .unwrap();
        a.define_name_for_cell("NewName", 0, 10, 0).unwrap();
        assert!(c.replay(&mut a, true).is_err());
    }
    #[test]
    fn pivot_outputs_move_whole_and_reject_partial_structural_cuts() {
        use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotField, PivotValueField};
        let mut b = fixture(true);
        let id = b.active_sheet().tables()[0].id;
        let source = b.table_pivot_source(id).unwrap();
        let field = &b.table(id).unwrap().1.columns[1];
        let definition = PivotDefinition {
            rows: vec![],
            column: None,
            values: vec![PivotValueField {
                field: PivotField {
                    column_id: Some(field.id),
                    offset: 1,
                    header: field.name.clone(),
                },
                aggregation: Aggregation::Sum,
                number_format: None,
            }],
        };
        let (pivot, index) = b.create_pivot(source, definition).unwrap();
        let region = b
            .sheet(index)
            .unwrap()
            .pivots
            .iter()
            .find(|p| p.id == pivot)
            .unwrap()
            .region()
            .unwrap();
        let (mut a, c) = b
            .prepare_guarded_structure(index, vec![step(Axis::Row, region.0, 1, false)])
            .unwrap();
        assert_eq!(
            a.sheet(index).unwrap().pivots[0].region().unwrap().0,
            region.0 + 1
        );
        c.replay(&mut a, true).unwrap();
        assert_eq!(a.sheet(index).unwrap().pivots[0].region().unwrap(), region);
        assert!(b
            .prepare_guarded_structure(index, vec![step(Axis::Row, region.0, 1, true)])
            .is_err());
    }
    #[test]
    fn totals_columns_keep_filtered_records_and_history_rewind() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_table_total(id, 2, visigrid_engine::table::TableTotal {
            function: Some("sum".into()), ..Default::default()
        }).unwrap();
        let spec = base.active_sheet().table_view_spec().cloned();
        let (mut after, commit) = base.prepare_guarded_structure(0,
            vec![step(Axis::Col, 2, 1, false)]).unwrap();
        assert_eq!(after.active_sheet().get_display(7, 3), "90");
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        assert_eq!(after.active_sheet().get_raw(4, 3), "10");
        let mut history = crate::history::History::new();
        history.record_action_with_provenance(crate::history::UndoAction::TableStructureChanged {
            sheet_index: 0, description: "Insert totals column".into(),
            history: Box::new(TableStructureHistory { commit: commit.clone(), source_frozen: None,
                before: StructureLayout::default(), after: StructureLayout::default() }),
        }, None);
        let preview = history.build_workbook_before(1, Some(&base), 100, 10_000).unwrap();
        assert_eq!(preview.workbook.active_sheet().get_display(7, 3), "90");
        commit.replay(&mut after, true).unwrap();
        assert_eq!(after.active_sheet().get_display(7, 2), "90");
        commit.replay(&mut after, false).unwrap();
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        assert!(after.prepare_guarded_structure(0, vec![step(Axis::Col, 3, 1, true)]).is_err());
        let (mut removed, deletion) = after.prepare_guarded_structure(0,
            vec![step(Axis::Col, 4, 1, true)]).unwrap();
        assert_eq!(removed.active_sheet().tables()[0].totals.as_ref().unwrap().columns.len(), 3);
        deletion.replay(&mut removed, true).unwrap();
        assert_eq!(removed.active_sheet().get_display(7, 4), "180");
    }

    #[test]
    fn totals_row_edits_preserve_filtered_records_and_rewind() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_calculated_column(id, 3, 3, "=[@Amount]*2", true).unwrap();
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        let spec = base.active_sheet().table_view_spec().cloned();
        let rows = base.active_sheet().build_saved_table_view(base.active_sheet().rows).unwrap().unwrap();
        let steps = selected_row_steps(rows.rows(), 3, 6, true).unwrap();
        let (mut after, commit) = base.prepare_guarded_structure(0, steps).unwrap();
        assert_eq!(after.active_sheet().tables()[0].totals_row(), Some(4));
        assert_eq!(after.active_sheet().get_raw(3, 1), "East");
        assert_eq!(after.active_sheet().get_display(4, 3), "0");
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        let mut history = crate::history::History::new();
        history.record_named_range_action(UndoAction::TableStructureChanged {
            sheet_index: 0, description: "Delete visible records".into(),
            history: Box::new(TableStructureHistory { commit: commit.clone(), source_frozen: None,
                before: StructureLayout::default(), after: StructureLayout::default() }),
        });
        let preview = history.build_workbook_before(1, Some(&base), 100, 10_000).unwrap();
        assert_eq!(preview.workbook.active_sheet().get_display(4, 3), "0");
        commit.replay(&mut after, true).unwrap();
        assert_eq!(after.active_sheet().get_display(7, 3), "180");
        commit.replay(&mut after, false).unwrap();
        assert_eq!(after.active_sheet().get_display(3, 3), "20");
        let (added, _) = base.prepare_guarded_structure(0, vec![step(Axis::Row, 7, 1, false)]).unwrap();
        assert_eq!(added.active_sheet().tables()[0].totals_row(), Some(8));
        assert_eq!(added.active_sheet().get_raw(7, 3), "=[@[Amount]]*2");
        assert_eq!(added.active_sheet().get_display(8, 3), "180");
    }

    #[test]
    fn totals_row_candidates_refuse_spills_without_changing_live_state() {
        let mut base = fixture(true);
        let id = base.active_sheet().tables()[0].id;
        base.set_table_totals_visible(id, true, Default::default()).unwrap();
        base.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IF(ROWS(Sales[Amount])=4,1,5))");
        let revision = base.revision();
        assert!(base.prepare_guarded_structure(0, vec![step(Axis::Row, 7, 1, false)]).is_err());
        assert_eq!(base.revision(), revision);
        assert_eq!(base.active_sheet().tables()[0].totals_row(), Some(7));
        assert_eq!(base.active_sheet().get_display(7, 3), "180");
    }

}

#[cfg(test)]
mod row_visibility_recalc_tests {
    use super::*;
    use visigrid_engine::RecalcClock;

    #[test]
    fn plain_insert_undo_does_not_recalculate_when_hidden_flags_already_match() {
        let mut wb = Workbook::new();
        wb.set_recalc_clock(Some(RecalcClock { now_ms: Some(0), utc_offset_seconds: Some(0), seed: None }));
        wb.active_sheet_mut().set_manual_hidden_rows([5].into()).unwrap();
        wb.set_cell_value_tracked(0, 0, 0, "=NOW()");
        let history = RowLayoutHistory::capture(StructureLayout::default(), wb.active_sheet(),
            StructureStep { axis: Axis::Row, at: 1, count: 1, delete: false }).unwrap();
        wb.structural_edit(0, Axis::Row, 1, 1, false).unwrap();
        let value = wb.active_sheet().get_display(0, 0);
        wb.active_sheet_mut().delete_rows(1, 1);
        wb.set_recalc_clock(Some(RecalcClock { now_ms: Some(86_400_000), utc_offset_seconds: Some(0), seed: None }));
        history.apply(&mut wb, 0, true).unwrap();
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [5].into());
        assert_eq!(wb.active_sheet().get_display(0, 0), value, "visibility-only replay need not evaluate NOW again");
        wb.recompute_full_ordered();
        assert_ne!(wb.active_sheet().get_display(0, 0), value, "the clock change makes a full recalc observable");
    }
}
