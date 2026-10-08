//! Grid structural operations
//!
//! Contains:
//! - Insert rows/columns
//! - Delete rows/columns
//! - Hide/unhide rows/columns
//! - Row height and column width management during insert/delete

use gpui::*;
use crate::app::{Spreadsheet, NUM_ROWS, NUM_COLS};
use crate::repeat::RepeatAction;
use visigrid_engine::structural::Axis;

/// Where an insert lands for `selection`, as `(at, count)` in view coordinates.
///
/// Whole-row and whole-column selections keep their meaning: the insert spans
/// the selection. Any other selection inserts at its top row or left column,
/// one line per selected row or column, as Excel's Insert Sheet Rows and
/// Insert Sheet Columns do. Asking for rows with whole columns selected (or
/// columns with whole rows) has no sensible span, so one line is inserted at
/// the active cell.
pub(crate) fn insert_target(
    axis: Axis,
    ((min_row, min_col), (max_row, max_col)): ((usize, usize), (usize, usize)),
    active: (usize, usize),
) -> (usize, usize) {
    let whole_rows = min_col == 0 && max_col == NUM_COLS - 1;
    let whole_cols = min_row == 0 && max_row == NUM_ROWS - 1;
    match axis {
        Axis::Row if whole_cols && !whole_rows => (active.0, 1),
        Axis::Row => (min_row, max_row - min_row + 1),
        Axis::Col if whole_rows && !whole_cols => (active.1, 1),
        Axis::Col => (min_col, max_col - min_col + 1),
    }
}

impl Spreadsheet {
    // =========================================================================
    // Row/Column insert/delete operations (Ctrl+= / Ctrl+-)
    // =========================================================================

    /// Insert rows or columns based on current selection (Ctrl+=)
    pub fn insert_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.table_structure_selection(false, cx);
            return;
        }
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            // Insert rows above selection — convert view→data when sorted
            let ((min_view, _), (max_view, _)) = self.selection_range();
            let count = max_view - min_view + 1;
            let data_row = self.view_to_data(min_view, cx);
            self.insert_rows(data_row, count, cx);
        } else if self.is_col_selection() {
            // Insert columns left of selection
            let ((_, min_col), (_, max_col)) = self.selection_range();
            let count = max_col - min_col + 1;
            self.insert_cols(min_col, count, cx);
        } else {
            // v1: No dialog, just show status message
            self.status_message = Some("Select entire row (Shift+Space) or column (Ctrl+Space) first".to_string());
            cx.notify();
        }
    }

    /// Insert rows for the current selection: above whole-row selections as
    /// before, and otherwise above the active cell, one row per selected row.
    pub fn insert_rows_at_selection(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".to_string());
            cx.notify();
            return;
        }
        let (at_view, count) = insert_target(Axis::Row, self.selection_range(), self.view_state.selected);
        let data_row = self.view_to_data(at_view, cx);
        self.insert_rows(data_row, count, cx);
    }

    /// Insert columns for the current selection: left of whole-column
    /// selections as before, and otherwise left of the active cell, one
    /// column per selected column.
    pub fn insert_cols_at_selection(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".to_string());
            cx.notify();
            return;
        }
        let (at_col, count) = insert_target(Axis::Col, self.selection_range(), self.view_state.selected);
        self.insert_cols(at_col, count, cx);
    }

    /// Delete rows or columns based on current selection (Ctrl+-)
    pub fn delete_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.table_structure_selection(true, cx);
            return;
        }
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Delete not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            let ((min_view, _), (max_view, _)) = self.selection_range();

            if self.row_view.is_sorted() {
                // Sorted: view rows may map to non-contiguous data rows.
                // Convert each view row to data row, sort descending, delete bottom-up.
                let mut data_rows: Vec<usize> = (min_view..=max_view)
                    .map(|vr| self.view_to_data(vr, cx))
                    .collect();
                data_rows.sort_unstable();
                data_rows.dedup();
                // Delete from bottom-up so earlier indices stay valid
                for &data_row in data_rows.iter().rev() {
                    self.delete_rows(data_row, 1, cx);
                }
            } else {
                // Not sorted: view == data, contiguous range
                let count = max_view - min_view + 1;
                self.delete_rows(min_view, count, cx);
            }
        } else if self.is_col_selection() {
            // Delete selected columns
            let ((_, min_col), (_, max_col)) = self.selection_range();
            let count = max_col - min_col + 1;
            self.delete_cols(min_col, count, cx);
        } else {
            // v1: No dialog, just show status message
            self.status_message = Some("Select entire row (Shift+Space) or column (Ctrl+Space) first".to_string());
            cx.notify();
        }
    }

    /// Insert rows at position with undo support
    pub(crate) fn insert_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) || self.wb(cx).tables().any(|(_, t)| t.totals.is_some()) {
            self.apply_table_structure(
                vec![visigrid_engine::workbook::StructureStep {
                    axis: visigrid_engine::structural::Axis::Row,
                    at: at_row,
                    count,
                    delete: false,
                }],
                cx,
            );
            return;
        }
        self.set_repeat(RepeatAction::InsertRows(count));
        let sheet_index = self.sheet_index(cx);
        if !self.sheet(cx).tables().is_empty() && (self.row_view.is_sorted() || self.row_view.is_filtered()) {
            self.status_message = Some("Clear sorting and filters before changing Table rows.".into()); cx.notify(); return;
        }
        let table_rows = match self.wb(cx).prepare_table_row_history(sheet_index, at_row, count, false) {
            Ok(history) => history,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let print_setup_before = self.sheet(cx).print_setup.clone();
        let row_layout = match crate::table_structure::RowLayoutHistory::capture(
            self.structure_layout(self.sheet(cx).id), self.sheet(cx),
            visigrid_engine::workbook::StructureStep {
                axis: visigrid_engine::structural::Axis::Row, at: at_row, count, delete: false,
            },
        ) {
            Ok(layout) => Some(Box::new(layout)),
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };


        // Perform the insert through the engine's structural entry point so
        // formulas, validations, and named ranges follow the moved cells.
        let rewrites = match self.workbook.update(cx, |wb, _| {
            if let Some(history) = &table_rows { wb.apply_table_row_history(history, false) }
            else { wb.structural_edit(sheet_index, visigrid_engine::structural::Axis::Row, at_row, count, false) }
        }) {
            Ok(r) => r,
            Err(msg) => {
                self.status_message = Some(msg);
                cx.notify();
                return;
            }
        };

        // Update row_view to track new data rows
        for i in 0..count {
            self.row_view.insert_row(at_row + i);
        }

        // Shift row heights down (from bottom to avoid overwriting)
        let sheet_heights = self.sheet_row_heights_mut();
        let heights_to_shift: Vec<_> = sheet_heights
            .iter()
            .filter(|(r, _)| **r >= at_row)
            .map(|(r, h)| (*r, *h))
            .collect();
        for (r, _) in &heights_to_shift {
            sheet_heights.remove(r);
        }
        for (r, h) in heights_to_shift {
            let new_row = r + count;
            if new_row < NUM_ROWS {
                sheet_heights.insert(new_row, h);
            }
        }

        if let Some(layout) = &row_layout {
            self.install_structure_layout(self.sheet(cx).id, &layout.after);
        }

        // Record undo entry
        self.record_named_range_action(cx, crate::history::UndoAction::RowsInserted {
            sheet_index,
            table_rows,
            row_layout,
            at_row,
            count,
            print_setup_before,
            formula_rewrites: rewrites.iter().map(|(si, r, c, old, _)| (*si, *r, *c, old.clone())).collect(),
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} row(s)", count));
        cx.notify();
    }

    /// Delete rows at position with undo support
    pub(crate) fn delete_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) || self.wb(cx).tables().any(|(_, t)| t.totals.is_some()) {
            self.apply_table_structure(
                vec![visigrid_engine::workbook::StructureStep {
                    axis: visigrid_engine::structural::Axis::Row,
                    at: at_row,
                    count,
                    delete: true,
                }],
                cx,
            );
            return;
        }
        self.set_repeat(RepeatAction::DeleteRows(count));
        let sheet_index = self.sheet_index(cx);
        if !self.sheet(cx).tables().is_empty() && (self.row_view.is_sorted() || self.row_view.is_filtered()) {
            self.status_message = Some("Clear sorting and filters before changing Table rows.".into()); cx.notify(); return;
        }
        let table_rows = match self.wb(cx).prepare_table_row_history(sheet_index, at_row, count, true) {
            Ok(history) => history,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let print_setup_before = self.sheet(cx).print_setup.clone();
        let row_layout = match crate::table_structure::RowLayoutHistory::capture(
            self.structure_layout(self.sheet(cx).id), self.sheet(cx),
            visigrid_engine::workbook::StructureStep {
                axis: visigrid_engine::structural::Axis::Row, at: at_row, count, delete: true,
            },
        ) {
            Ok(layout) => Some(Box::new(layout)),
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };


        // Capture cells to be deleted for undo
        // Only cells that exist can be deleted, so ask the sparse store rather
        // than walking the whole band: at grid size that walk is 16.7M lookups
        // for a one-row delete.
        let sheet = self.sheet(cx);
        let deleted_cells = sheet.occupied_cells_in_rows(at_row, count);
        let deleted_comments = sheet.comments().filter(|((r, _), _)| (at_row..at_row + count).contains(r)).map(|((r, c), v)| (r, c, v.clone())).collect();

        // Capture row heights for deleted rows (per-sheet)
        let sheet_heights = self.sheet_row_heights_mut();
        let deleted_row_heights: Vec<_> = sheet_heights
            .iter()
            .filter(|(r, _)| **r >= at_row && **r < at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();

        let rewrites = match self.workbook.update(cx, |wb, _| {
            if let Some(history) = &table_rows { wb.apply_table_row_history(history, false) }
            else { wb.structural_edit(sheet_index, visigrid_engine::structural::Axis::Row, at_row, count, true) }
        }) {
            Ok(r) => r,
            Err(msg) => {
                self.status_message = Some(msg);
                cx.notify();
                return;
            }
        };

        let sheet_heights = self.sheet_row_heights_mut();
        // Remove heights for deleted rows and shift remaining up
        let heights_to_shift: Vec<_> = sheet_heights
            .iter()
            .filter(|(r, _)| **r >= at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();
        // Remove all affected heights
        sheet_heights.retain(|r, _| *r < at_row);
        // Re-insert shifted heights
        for (r, h) in heights_to_shift {
            sheet_heights.insert(r - count, h);
        }

        // Update row_view to remove deleted data rows (bottom-up to keep indices stable)
        for i in (0..count).rev() {
            self.row_view.delete_row(at_row + i);
        }

        if let Some(layout) = &row_layout {
            self.install_structure_layout(self.sheet(cx).id, &layout.after);
        }

        // Record undo entry
        self.record_named_range_action(cx, crate::history::UndoAction::RowsDeleted {
            sheet_index,
            table_rows,
            row_layout,
            at_row,
            count,
            deleted_cells,
            deleted_comments,
            deleted_row_heights,
            print_setup_before,
            formula_rewrites: rewrites.iter().map(|(si, r, c, old, _)| (*si, *r, *c, old.clone())).collect(),
        });

        // Maintain full-row selection at the same position (Excel behavior):
        // after deleting rows 3-5, the selection highlights rows 3-5 (now shifted-up data)
        let sel_row = at_row.min(NUM_ROWS - 1);
        self.view_state.selected = (sel_row, 0);
        self.view_state.selection_end = Some(((sel_row + count - 1).min(NUM_ROWS - 1), NUM_COLS - 1));
        self.view_state.additional_selections.clear();

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} row(s)", count));
        cx.notify();
    }

    /// Insert columns at position with undo support
    pub(crate) fn insert_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) || self.wb(cx).tables().any(|(_, t)| t.totals.is_some()) {
            self.apply_table_structure(
                vec![visigrid_engine::workbook::StructureStep {
                    axis: visigrid_engine::structural::Axis::Col,
                    at: at_col,
                    count,
                    delete: false,
                }],
                cx,
            );
            return;
        }
        self.set_repeat(RepeatAction::InsertCols(count));
        let sheet_index = self.sheet_index(cx);
        let table_columns = match self.wb(cx).prepare_table_column_history(sheet_index, at_col, count, false) {
            Ok(history) => history,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let print_setup_before = self.sheet(cx).print_setup.clone();

        // Perform the insert
        let rewrites = match self.workbook.update(cx, |wb, _| {
            if let Some(history) = &table_columns { wb.apply_table_column_history(history, false) }
            else { wb.structural_edit(sheet_index, visigrid_engine::structural::Axis::Col, at_col, count, false) }
        }) {
            Ok(r) => r,
            Err(msg) => {
                self.status_message = Some(msg);
                cx.notify();
                return;
            }
        };

        // Shift column widths right (from right to avoid overwriting) - per-sheet
        let sheet_widths = self.sheet_col_widths_mut();
        let widths_to_shift: Vec<_> = sheet_widths
            .iter()
            .filter(|(c, _)| **c >= at_col)
            .map(|(c, w)| (*c, *w))
            .collect();
        for (c, _) in &widths_to_shift {
            sheet_widths.remove(c);
        }
        for (c, w) in widths_to_shift {
            let new_col = c + count;
            if new_col < NUM_COLS {
                sheet_widths.insert(new_col, w);
            }
        }

        // Record undo entry
        self.record_named_range_action(cx, crate::history::UndoAction::ColsInserted {
            sheet_index,
            table_columns,
            at_col,
            count,
            print_setup_before,
            formula_rewrites: rewrites.iter().map(|(si, r, c, old, _)| (*si, *r, *c, old.clone())).collect(),
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} column(s)", count));
        cx.notify();
    }

    /// Delete columns at position with undo support
    pub(crate) fn delete_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) || self.wb(cx).tables().any(|(_, t)| t.totals.is_some()) {
            self.apply_table_structure(
                vec![visigrid_engine::workbook::StructureStep {
                    axis: visigrid_engine::structural::Axis::Col,
                    at: at_col,
                    count,
                    delete: true,
                }],
                cx,
            );
            return;
        }
        self.set_repeat(RepeatAction::DeleteCols(count));
        let sheet_index = self.sheet_index(cx);
        let table_columns = match self.wb(cx).prepare_table_column_history(sheet_index, at_col, count, true) {
            Ok(history) => history,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let print_setup_before = self.sheet(cx).print_setup.clone();

        // Capture cells to be deleted for undo
        // See delete_rows: sparse lookup, not a full-column walk.
        let sheet = self.sheet(cx);
        let deleted_cells = sheet.occupied_cells_in_cols(at_col, count);
        let deleted_comments = sheet.comments().filter(|((_, c), _)| (at_col..at_col + count).contains(c)).map(|((r, c), v)| (r, c, v.clone())).collect();

        // Capture column widths for deleted columns (per-sheet)
        let sheet_widths = self.sheet_col_widths_mut();
        let deleted_col_widths: Vec<_> = sheet_widths
            .iter()
            .filter(|(c, _)| **c >= at_col && **c < at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();

        // Perform the delete
        let rewrites = match self.workbook.update(cx, |wb, _| {
            if let Some(history) = &table_columns { wb.apply_table_column_history(history, false) }
            else { wb.structural_edit(sheet_index, visigrid_engine::structural::Axis::Col, at_col, count, true) }
        }) {
            Ok(r) => r,
            Err(msg) => {
                self.status_message = Some(msg);
                cx.notify();
                return;
            }
        };

        let sheet_widths = self.sheet_col_widths_mut();
        // Remove widths for deleted columns and shift remaining left
        let widths_to_shift: Vec<_> = sheet_widths
            .iter()
            .filter(|(c, _)| **c >= at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();
        // Remove all affected widths
        sheet_widths.retain(|c, _| *c < at_col);
        // Re-insert shifted widths
        for (c, w) in widths_to_shift {
            sheet_widths.insert(c - count, w);
        }

        // Record undo entry
        self.record_named_range_action(cx, crate::history::UndoAction::ColsDeleted {
            sheet_index,
            table_columns,
            at_col,
            count,
            deleted_cells,
            deleted_comments,
            deleted_col_widths,
            print_setup_before,
            formula_rewrites: rewrites.iter().map(|(si, r, c, old, _)| (*si, *r, *c, old.clone())).collect(),
        });

        // Maintain full-column selection at the same position (Excel behavior):
        // after deleting cols C-E, the selection highlights cols C-E (now shifted-left data)
        let sel_col = at_col.min(NUM_COLS - 1);
        self.view_state.selected = (0, sel_col);
        self.view_state.selection_end = Some((NUM_ROWS - 1, (sel_col + count - 1).min(NUM_COLS - 1)));
        self.view_state.additional_selections.clear();

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} column(s)", count));
        cx.notify();
    }

    // =========================================================================
    // Hide/Unhide rows and columns (Ctrl+9/0, Ctrl+Shift+9/0)
    // =========================================================================

    /// Hide selected rows (Ctrl+9)
    pub(crate) fn hide_rows(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) { return; }
        if self.mode.is_editing() { return; }
        if self.wb(cx).sheets().iter().any(|s| s.tables().iter().any(|t| t.totals.is_some()))
            || crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.change_table_row_visibility(true, cx);
            return;
        }

        let ((min_row, _), (max_row, _)) = self.selection_range();
        let empty = Default::default();
        let manual = self.display_hidden_rows().unwrap_or(&empty);
        let rows = match crate::table_visibility::selected_row_visibility(
            &self.row_view, manual, min_row, max_row, true,
        ) {
            Ok(rows) => rows,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };

        if rows.is_empty() { return; }

        let sheet_id = self.cached_sheet_id();
        let set = self.hidden_rows.entry(sheet_id).or_default();
        for &r in &rows {
            set.insert(r);
        }

        if !self.sync_manual_row_visibility(sheet_id, cx) { return; }
        self.record_action_with_provenance(cx,
            crate::history::UndoAction::RowVisibilityChanged {
                sheet_id,
                rows: rows.clone(),
                hidden: true,
            },
            None,
        );
        self.is_modified = true;
        self.status_message = Some(format!("Hidden {} row(s)", rows.len()));
        cx.notify();
    }

    /// Unhide rows adjacent to selection (Ctrl+Shift+9)
    ///
    /// Excel behavior: select rows spanning the hidden range, then unhide.
    /// E.g., if rows 5-8 are hidden, select rows 4-9 and press Ctrl+Shift+9.
    pub(crate) fn unhide_rows(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) { return; }
        if self.mode.is_editing() { return; }
        if self.wb(cx).sheets().iter().any(|s| s.tables().iter().any(|t| t.totals.is_some()))
            || crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.change_table_row_visibility(false, cx);
            return;
        }

        let ((min_row, _), (max_row, _)) = self.selection_range();
        let sheet_id = self.cached_sheet_id();
        let empty = Default::default();
        let manual = self.display_hidden_rows().unwrap_or(&empty);
        let rows = match crate::table_visibility::selected_row_visibility(
            &self.row_view, manual, min_row, max_row, false,
        ) {
            Ok(rows) => rows,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };

        if rows.is_empty() {
            self.status_message = Some("No hidden rows in selection".to_string());
            cx.notify();
            return;
        }

        let set = self.hidden_rows.entry(sheet_id).or_default();
        for &r in &rows {
            set.remove(&r);
        }

        if !self.sync_manual_row_visibility(sheet_id, cx) { return; }
        self.record_action_with_provenance(cx,
            crate::history::UndoAction::RowVisibilityChanged {
                sheet_id,
                rows: rows.clone(),
                hidden: false,
            },
            None,
        );
        self.is_modified = true;
        self.status_message = Some(format!("Unhidden {} row(s)", rows.len()));
        cx.notify();
    }

    /// Hide selected columns (Ctrl+0)
    pub(crate) fn hide_cols(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        if self.mode.is_editing() { return; }

        let ((_, min_col), (_, max_col)) = self.selection_range();
        let cols: Vec<usize> = (min_col..=max_col)
            .filter(|c| !self.is_col_hidden(*c))
            .collect();

        if cols.is_empty() { return; }

        let sheet_id = self.cached_sheet_id();
        let set = self.hidden_cols.entry(sheet_id).or_default();
        for &c in &cols {
            set.insert(c);
        }

        self.record_action_with_provenance(cx,
            crate::history::UndoAction::ColVisibilityChanged {
                sheet_id,
                cols: cols.clone(),
                hidden: true,
            },
            None,
        );
        self.is_modified = true;
        self.status_message = Some(format!("Hidden {} column(s)", cols.len()));
        cx.notify();
    }

    /// Unhide columns adjacent to selection (Ctrl+Shift+0)
    pub(crate) fn unhide_cols(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        if self.mode.is_editing() { return; }

        let ((_, min_col), (_, max_col)) = self.selection_range();
        let sheet_id = self.cached_sheet_id();
        let cols: Vec<usize> = (min_col..=max_col)
            .filter(|c| self.is_col_hidden(*c))
            .collect();

        if cols.is_empty() {
            self.status_message = Some("No hidden columns in selection".to_string());
            cx.notify();
            return;
        }

        let set = self.hidden_cols.entry(sheet_id).or_default();
        for &c in &cols {
            set.remove(&c);
        }

        self.record_action_with_provenance(cx,
            crate::history::UndoAction::ColVisibilityChanged {
                sheet_id,
                cols: cols.clone(),
                hidden: false,
            },
            None,
        );
        self.is_modified = true;
        self.status_message = Some(format!("Unhidden {} column(s)", cols.len()));
        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    use super::{insert_target, Axis, NUM_COLS, NUM_ROWS};

    fn whole_rows(from: usize, to: usize) -> ((usize, usize), (usize, usize)) {
        ((from, 0), (to, NUM_COLS - 1))
    }
    fn whole_cols(from: usize, to: usize) -> ((usize, usize), (usize, usize)) {
        ((0, from), (NUM_ROWS - 1, to))
    }

    #[test]
    fn a_single_cell_inserts_one_line_at_the_cell() {
        let sel = ((6, 3), (6, 3));
        assert_eq!(insert_target(Axis::Row, sel, (6, 3)), (6, 1));
        assert_eq!(insert_target(Axis::Col, sel, (6, 3)), (3, 1));
    }

    #[test]
    fn a_cell_range_inserts_one_line_per_selected_row_or_column() {
        // D4:E6: three rows, two columns, active cell anywhere inside.
        let sel = ((3, 3), (5, 4));
        assert_eq!(insert_target(Axis::Row, sel, (5, 4)), (3, 3));
        assert_eq!(insert_target(Axis::Col, sel, (5, 4)), (3, 2));
    }

    #[test]
    fn whole_row_and_column_selections_keep_their_existing_span() {
        assert_eq!(insert_target(Axis::Row, whole_rows(4, 5), (4, 0)), (4, 2));
        assert_eq!(insert_target(Axis::Col, whole_cols(2, 3), (0, 2)), (2, 2));
    }

    #[test]
    fn the_other_axis_of_a_whole_line_selection_inserts_one_line_at_the_active_cell() {
        // Whole columns selected, rows requested: one row at the active cell.
        assert_eq!(insert_target(Axis::Row, whole_cols(2, 3), (7, 2)), (7, 1));
        // Whole rows selected, columns requested: one column at the active cell.
        assert_eq!(insert_target(Axis::Col, whole_rows(4, 5), (4, 9)), (9, 1));
    }

    #[test]
    fn select_all_spans_the_grid_so_the_engine_can_refuse_it() {
        let all = ((0, 0), (NUM_ROWS - 1, NUM_COLS - 1));
        assert_eq!(insert_target(Axis::Row, all, (0, 0)), (0, NUM_ROWS));
        assert_eq!(insert_target(Axis::Col, all, (0, 0)), (0, NUM_COLS));
    }
}
