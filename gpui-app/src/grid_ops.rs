//! Grid structural operations
//!
//! Contains:
//! - Insert rows/columns
//! - Delete rows/columns
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
    fn insert_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Perform the insert
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.insert_rows(at_row, count);
        }

        // Shift row heights down (from bottom to avoid overwriting)
        let heights_to_shift: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row)
            .map(|(r, h)| (*r, *h))
            .collect();
        for (r, _) in &heights_to_shift {
            self.row_heights.remove(r);
        }
        for (r, h) in heights_to_shift {
            let new_row = r + count;
            if new_row < NUM_ROWS {
                self.row_heights.insert(new_row, h);
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
    fn delete_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        if let Some(sheet) = self.workbook.sheet(sheet_index) {
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
        }

        // Capture row heights for deleted rows
        let deleted_row_heights: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row && **r < at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();

        // Remove heights for deleted rows and shift remaining up
        let heights_to_shift: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();
        // Remove all affected heights
        for r in at_row..NUM_ROWS {
            self.row_heights.remove(&r);
        }
        // Re-insert shifted heights
        for (r, h) in heights_to_shift {
            self.row_heights.insert(r - count, h);
        }

        // Perform the delete
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.delete_rows(at_row, count);
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::RowsDeleted {
            sheet_index,
            at_row,
            count,
            deleted_cells,
            deleted_row_heights,
        });

        // Move selection up if needed
        if self.view_state.selected.0 >= at_row + count {
            self.view_state.selected.0 -= count;
        } else if self.view_state.selected.0 >= at_row {
            self.view_state.selected.0 = at_row.saturating_sub(1);
        }
        self.view_state.selection_end = None;

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} row(s)", count));
        cx.notify();
    }

    /// Insert columns at position with undo support
    fn insert_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Perform the insert
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.insert_cols(at_col, count);
        }

        // Shift column widths right (from right to avoid overwriting)
        let widths_to_shift: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col)
            .map(|(c, w)| (*c, *w))
            .collect();
        for (c, _) in &widths_to_shift {
            self.col_widths.remove(c);
        }
        for (c, w) in widths_to_shift {
            let new_col = c + count;
            if new_col < NUM_COLS {
                self.col_widths.insert(new_col, w);
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
    fn delete_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        if let Some(sheet) = self.workbook.sheet(sheet_index) {
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
        }

        // Capture column widths for deleted columns
        let deleted_col_widths: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col && **c < at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();

        // Remove widths for deleted columns and shift remaining left
        let widths_to_shift: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();
        // Remove all affected widths
        for c in at_col..NUM_COLS {
            self.col_widths.remove(&c);
        }
        // Re-insert shifted widths
        for (c, w) in widths_to_shift {
            self.col_widths.insert(c - count, w);
        }

        // Perform the delete
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.delete_cols(at_col, count);
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::ColsDeleted {
            sheet_index,
            at_col,
            count,
            deleted_cells,
            deleted_col_widths,
        });

        // Move selection left if needed
        if self.view_state.selected.1 >= at_col + count {
            self.view_state.selected.1 -= count;
        } else if self.view_state.selected.1 >= at_col {
            self.view_state.selected.1 = at_col.saturating_sub(1);
        }
        self.view_state.selection_end = None;

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} column(s)", count));
        cx.notify();
    }
}
