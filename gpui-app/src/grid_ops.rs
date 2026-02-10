//! Grid structural operations
//!
//! Contains:
//! - Insert rows/columns
//! - Delete rows/columns
//! - Hide/unhide rows/columns
//! - Row height and column width management during insert/delete

use gpui::*;
use crate::app::{Spreadsheet, NUM_ROWS, NUM_COLS};

impl Spreadsheet {
    // =========================================================================
    // Row/Column insert/delete operations (Ctrl+= / Ctrl+-)
    // =========================================================================

    /// Insert rows or columns based on current selection (Ctrl+=)
    pub fn insert_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            // Insert rows above selection
            let ((min_row, _), (max_row, _)) = self.selection_range();
            let count = max_row - min_row + 1;
            self.insert_rows(min_row, count, cx);
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

    /// Delete rows or columns based on current selection (Ctrl+-)
    pub fn delete_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Delete not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            // Delete selected rows
            let ((min_row, _), (max_row, _)) = self.selection_range();
            let count = max_row - min_row + 1;
            self.delete_rows(min_row, count, cx);
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
        let sheet_index = self.sheet_index(cx);

        // Perform the insert
        self.sheet_mut(sheet_index, cx, |sheet| {
            sheet.insert_rows(at_row, count);
        });

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

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::RowsInserted {
            sheet_index,
            at_row,
            count,
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} row(s)", count));
        cx.notify();
    }

    /// Delete rows at position with undo support
    pub(crate) fn delete_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.sheet_index(cx);

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        let sheet = self.sheet(cx);
        for row in at_row..at_row + count {
            for col in 0..NUM_COLS {
                let raw = sheet.get_raw(row, col);
                let format = sheet.get_format(row, col);
                // Only store non-empty cells
                if !raw.is_empty() || format != Default::default() {
                    deleted_cells.push((row, col, raw, format));
                }
            }
        }

        // Capture row heights for deleted rows (per-sheet)
        let sheet_heights = self.sheet_row_heights_mut();
        let deleted_row_heights: Vec<_> = sheet_heights
            .iter()
            .filter(|(r, _)| **r >= at_row && **r < at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();

        // Remove heights for deleted rows and shift remaining up
        let heights_to_shift: Vec<_> = sheet_heights
            .iter()
            .filter(|(r, _)| **r >= at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();
        // Remove all affected heights
        for r in at_row..NUM_ROWS {
            sheet_heights.remove(&r);
        }
        // Re-insert shifted heights
        for (r, h) in heights_to_shift {
            sheet_heights.insert(r - count, h);
        }

        // Perform the delete
        self.sheet_mut(sheet_index, cx, |sheet| {
            sheet.delete_rows(at_row, count);
        });

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::RowsDeleted {
            sheet_index,
            at_row,
            count,
            deleted_cells,
            deleted_row_heights,
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
        let sheet_index = self.sheet_index(cx);

        // Perform the insert
        self.sheet_mut(sheet_index, cx, |sheet| {
            sheet.insert_cols(at_col, count);
        });

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
        self.history.record_named_range_action(crate::history::UndoAction::ColsInserted {
            sheet_index,
            at_col,
            count,
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} column(s)", count));
        cx.notify();
    }

    /// Delete columns at position with undo support
    pub(crate) fn delete_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.sheet_index(cx);

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        let sheet = self.sheet(cx);
        for col in at_col..at_col + count {
            for row in 0..NUM_ROWS {
                let raw = sheet.get_raw(row, col);
                let format = sheet.get_format(row, col);
                // Only store non-empty cells
                if !raw.is_empty() || format != Default::default() {
                    deleted_cells.push((row, col, raw, format));
                }
            }
        }

        // Capture column widths for deleted columns (per-sheet)
        let sheet_widths = self.sheet_col_widths_mut();
        let deleted_col_widths: Vec<_> = sheet_widths
            .iter()
            .filter(|(c, _)| **c >= at_col && **c < at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();

        // Remove widths for deleted columns and shift remaining left
        let widths_to_shift: Vec<_> = sheet_widths
            .iter()
            .filter(|(c, _)| **c >= at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();
        // Remove all affected widths
        for c in at_col..NUM_COLS {
            sheet_widths.remove(&c);
        }
        // Re-insert shifted widths
        for (c, w) in widths_to_shift {
            sheet_widths.insert(c - count, w);
        }

        // Perform the delete
        self.sheet_mut(sheet_index, cx, |sheet| {
            sheet.delete_cols(at_col, count);
        });

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::ColsDeleted {
            sheet_index,
            at_col,
            count,
            deleted_cells,
            deleted_col_widths,
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
        if self.block_if_previewing(cx) { return; }
        if self.mode.is_editing() { return; }

        let ((min_row, _), (max_row, _)) = self.selection_range();
        let rows: Vec<usize> = (min_row..=max_row)
            .filter(|r| !self.is_row_hidden(*r))
            .collect();

        if rows.is_empty() { return; }

        let sheet_id = self.cached_sheet_id();
        let set = self.hidden_rows.entry(sheet_id).or_default();
        for &r in &rows {
            set.insert(r);
        }

        self.history.record_action_with_provenance(
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
        if self.block_if_previewing(cx) { return; }
        if self.mode.is_editing() { return; }

        let ((min_row, _), (max_row, _)) = self.selection_range();
        let sheet_id = self.cached_sheet_id();
        let rows: Vec<usize> = (min_row..=max_row)
            .filter(|r| self.is_row_hidden(*r))
            .collect();

        if rows.is_empty() {
            self.status_message = Some("No hidden rows in selection".to_string());
            cx.notify();
            return;
        }

        let set = self.hidden_rows.entry(sheet_id).or_default();
        for &r in &rows {
            set.remove(&r);
        }

        self.history.record_action_with_provenance(
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

        self.history.record_action_with_provenance(
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

        self.history.record_action_with_provenance(
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
