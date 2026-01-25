//! Navigation and selection operations
//!
//! Contains:
//! - Cell movement (arrow keys, page up/down)
//! - Jump navigation (Ctrl+Arrow)
//! - Selection extension (Shift+Arrow, Shift+Ctrl+Arrow)
//! - Drag selection (mouse)
//! - Row/column header selection
//! - Scrolling
//! - Selection helpers

use gpui::*;
use crate::app::{Spreadsheet, NUM_ROWS, NUM_COLS};

impl Spreadsheet {
    // =========================================================================
    // Cell Movement
    // =========================================================================

    pub fn move_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selected;
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.view_state.selected = (new_row, new_col);
        self.view_state.selection_end = None;  // Clear range selection
        self.view_state.additional_selections.clear();  // Clear discontiguous selections

        self.ensure_visible(cx);
    }

    pub fn extend_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.view_state.selection_end = Some((new_row, new_col));

        self.ensure_visible(cx);
    }

    pub fn page_up(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        let visible_rows = self.visible_rows() as i32;
        self.move_selection(-visible_rows, 0, cx);
    }

    pub fn page_down(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        let visible_rows = self.visible_rows() as i32;
        self.move_selection(visible_rows, 0, cx);
    }

    // =========================================================================
    // Jump Navigation (Ctrl+Arrow)
    // =========================================================================

    /// Find the data boundary in a direction (used by Ctrl+Arrow and Ctrl+Shift+Arrow)
    pub(crate) fn find_data_boundary(&self, start_row: usize, start_col: usize, dr: i32, dc: i32) -> (usize, usize) {
        let mut row = start_row;
        let mut col = start_col;
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        (row, col)
    }

    /// Jump to edge of data region or sheet boundary (Excel-style Ctrl+Arrow)
    pub fn jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (mut row, mut col) = self.view_state.selected;
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.ensure_visible(cx);
    }

    /// Extend selection to edge of data region (Excel-style Ctrl+Shift+Arrow)
    pub fn extend_jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Start from current selection end (or selected if no selection)
        let (mut row, mut col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        // Extend selection to this point (don't move selected, just selection_end)
        self.view_state.selection_end = Some((row, col));
        self.ensure_visible(cx);
    }

    // =========================================================================
    // Visibility / Scrolling
    // =========================================================================

    pub fn ensure_visible(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // When freeze panes are active, calculate scrollable region
        let scrollable_visible_rows = visible_rows.saturating_sub(self.view_state.frozen_rows);
        let scrollable_visible_cols = visible_cols.saturating_sub(self.view_state.frozen_cols);

        // Vertical scroll - frozen rows are always visible, only scroll for rows in scrollable region
        if row < self.view_state.frozen_rows {
            // Row is in frozen region - always visible, but ensure scroll_row is valid
            self.view_state.scroll_row = self.view_state.scroll_row.max(self.view_state.frozen_rows);
        } else if row < self.view_state.scroll_row {
            self.view_state.scroll_row = row;
        } else if scrollable_visible_rows > 0 && row >= self.view_state.scroll_row + scrollable_visible_rows {
            self.view_state.scroll_row = row - scrollable_visible_rows + 1;
        }

        // Horizontal scroll - frozen cols are always visible, only scroll for cols in scrollable region
        if col < self.view_state.frozen_cols {
            // Col is in frozen region - always visible, but ensure scroll_col is valid
            self.view_state.scroll_col = self.view_state.scroll_col.max(self.view_state.frozen_cols);
        } else if col < self.view_state.scroll_col {
            self.view_state.scroll_col = col;
        } else if scrollable_visible_cols > 0 && col >= self.view_state.scroll_col + scrollable_visible_cols {
            self.view_state.scroll_col = col - scrollable_visible_cols + 1;
        }

        // Ensure scroll positions don't go below freeze bounds
        self.view_state.scroll_row = self.view_state.scroll_row.max(self.view_state.frozen_rows);
        self.view_state.scroll_col = self.view_state.scroll_col.max(self.view_state.frozen_cols);

        cx.notify();
    }

    pub fn scroll(&mut self, delta_rows: i32, delta_cols: i32, cx: &mut Context<Self>) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // When freeze panes are active, scrollable region starts after frozen rows/cols
        let min_scroll_row = self.view_state.frozen_rows;
        let min_scroll_col = self.view_state.frozen_cols;

        let new_row = (self.view_state.scroll_row as i32 + delta_rows)
            .max(min_scroll_row as i32)
            .min((NUM_ROWS.saturating_sub(visible_rows)) as i32) as usize;
        let new_col = (self.view_state.scroll_col as i32 + delta_cols)
            .max(min_scroll_col as i32)
            .min((NUM_COLS.saturating_sub(visible_cols)) as i32) as usize;

        if new_row != self.view_state.scroll_row || new_col != self.view_state.scroll_col {
            self.view_state.scroll_row = new_row;
            self.view_state.scroll_col = new_col;
            cx.notify();
        }
    }

    // =========================================================================
    // Cell Selection
    // =========================================================================

    pub fn select_cell(&mut self, row: usize, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            self.view_state.selection_end = Some((row, col));
        } else {
            self.view_state.selected = (row, col);
            self.view_state.selection_end = None;
            self.view_state.additional_selections.clear();  // Clear Ctrl+Click selections
            // Clear trace path when selection changes (unless inspector is pinned)
            if self.inspector_pinned.is_none() && self.inspector_trace_path.is_some() {
                self.inspector_trace_path = None;
                self.inspector_trace_incomplete = false;
            }
        }
        cx.notify();
    }

    /// Ctrl+Click to add/toggle cell in selection (discontiguous selection)
    pub fn ctrl_click_cell(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        // Save current selection to additional_selections
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        // Start new selection at clicked cell
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        cx.notify();
    }

    pub fn select_all(&mut self, cx: &mut Context<Self>) {
        self.view_state.selected = (0, 0);
        self.view_state.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        self.view_state.additional_selections.clear();  // Clear discontiguous selections
        cx.notify();
    }

    // =========================================================================
    // Drag Selection
    // =========================================================================

    /// Start drag selection - called on mouse_down
    pub fn start_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.view_state.additional_selections.clear();  // Clear Ctrl+Click selections on new drag
        cx.notify();
    }

    /// Start drag selection with Ctrl held (add to existing selections)
    pub fn start_ctrl_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        // Save current selection to additional_selections
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        // Start new selection at clicked cell
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        cx.notify();
    }

    /// Continue drag selection - called on mouse_move while dragging
    pub fn continue_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_selection {
            return;
        }
        // Only update if the cell changed to avoid unnecessary redraws
        if self.view_state.selection_end != Some((row, col)) {
            self.view_state.selection_end = Some((row, col));
            cx.notify();
        }
    }

    /// End drag selection - called on mouse_up
    pub fn end_drag_selection(&mut self, cx: &mut Context<Self>) {
        if self.dragging_selection {
            self.dragging_selection = false;
            cx.notify();
        }
    }

    // =========================================================================
    // Selection Helpers
    // =========================================================================

    pub fn selection_range(&self) -> ((usize, usize), (usize, usize)) {
        let start = self.view_state.selected;
        let end = self.view_state.selection_end.unwrap_or(start);
        let min_row = start.0.min(end.0);
        let max_row = start.0.max(end.0);
        let min_col = start.1.min(end.1);
        let max_col = start.1.max(end.1);
        ((min_row, min_col), (max_row, max_col))
    }

    /// Clamp selection to valid bounds after operations that might invalidate it.
    /// Preserves column where possible (user mental model), clamps row to valid range.
    pub fn clamp_selection(&mut self) {
        // Clamp selected cell
        self.view_state.selected.0 = self.view_state.selected.0.min(NUM_ROWS - 1);
        self.view_state.selected.1 = self.view_state.selected.1.min(NUM_COLS - 1);

        // Clamp selection_end if present
        if let Some(ref mut end) = self.view_state.selection_end {
            end.0 = end.0.min(NUM_ROWS - 1);
            end.1 = end.1.min(NUM_COLS - 1);
        }

        // Clamp additional selections
        for (start, end) in &mut self.view_state.additional_selections {
            start.0 = start.0.min(NUM_ROWS - 1);
            start.1 = start.1.min(NUM_COLS - 1);
            if let Some(ref mut e) = end {
                e.0 = e.0.min(NUM_ROWS - 1);
                e.1 = e.1.min(NUM_COLS - 1);
            }
        }
    }

    /// Returns true if more than one cell is selected.
    /// This includes range selections and Ctrl+Click additional selections.
    pub fn is_multi_selection(&self) -> bool {
        // Check if primary selection is a range (more than one cell)
        if let Some(end) = self.view_state.selection_end {
            if end != self.view_state.selected {
                return true;
            }
        }
        // Check if there are additional Ctrl+Click selections
        if !self.view_state.additional_selections.is_empty() {
            return true;
        }
        false
    }

    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        // Check active selection
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            return true;
        }
        // Check additional selections (Ctrl+Click ranges)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                return true;
            }
        }
        false
    }

    /// Get all selection ranges (for operations that apply to all selected cells)
    pub fn all_selection_ranges(&self) -> Vec<((usize, usize), (usize, usize))> {
        let mut ranges = Vec::new();
        // Add active selection
        ranges.push(self.selection_range());
        // Add additional selections
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            ranges.push(((min_row, min_col), (max_row, max_col)));
        }
        ranges
    }

    // =========================================================================
    // Row/Column Header Selection
    // =========================================================================

    /// Check if the active selection spans all columns (row selection)
    pub fn is_row_selection(&self) -> bool {
        let ((_, min_col), (_, max_col)) = self.selection_range();
        min_col == 0 && max_col == NUM_COLS - 1
    }

    /// Check if the active selection spans all rows (column selection)
    pub fn is_col_selection(&self) -> bool {
        let ((min_row, _), (max_row, _)) = self.selection_range();
        min_row == 0 && max_row == NUM_ROWS - 1
    }

    /// Check if row header should be highlighted (checks all selections)
    pub fn is_row_header_selected(&self, row: usize) -> bool {
        for ((min_row, _), (max_row, _)) in self.all_selection_ranges() {
            if row >= min_row && row <= max_row {
                return true;
            }
        }
        false
    }

    /// Check if column header should be highlighted (checks all selections)
    pub fn is_col_header_selected(&self, col: usize) -> bool {
        for ((_, min_col), (_, max_col)) in self.all_selection_ranges() {
            if col >= min_col && col <= max_col {
                return true;
            }
        }
        false
    }

    /// Select entire row. If extend=true, extends from current anchor row.
    pub fn select_row(&mut self, row: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            // Extend from the current anchor (self.view_state.selected.0 before this call)
            let anchor_row = self.view_state.selected.0;
            self.view_state.selected = (anchor_row.min(row), 0);
            self.view_state.selection_end = Some((anchor_row.max(row), NUM_COLS - 1));
        } else {
            self.view_state.selected = (row, 0);
            self.view_state.selection_end = Some((row, NUM_COLS - 1));
            self.view_state.additional_selections.clear();
        }
        cx.notify();
    }

    /// Select entire column. If extend=true, extends from current anchor col.
    pub fn select_col(&mut self, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            let anchor_col = self.view_state.selected.1;
            self.view_state.selected = (0, anchor_col.min(col));
            self.view_state.selection_end = Some((NUM_ROWS - 1, anchor_col.max(col)));
        } else {
            self.view_state.selected = (0, col);
            self.view_state.selection_end = Some((NUM_ROWS - 1, col));
            self.view_state.additional_selections.clear();
        }
        cx.notify();
    }

    // =========================================================================
    // Row/Column Header Drag
    // =========================================================================

    /// Start row header drag - stores stable anchor
    pub fn start_row_header_drag(&mut self, row: usize, cx: &mut Context<Self>) {
        self.dragging_row_header = true;
        self.dragging_col_header = false;
        self.dragging_selection = false;
        self.row_header_anchor = Some(row);
        self.select_row(row, false, cx);
    }

    /// Continue row header drag - uses stored anchor
    pub fn continue_row_header_drag(&mut self, row: usize, cx: &mut Context<Self>) {
        if !self.dragging_row_header { return; }
        let anchor = self.row_header_anchor.unwrap_or(row);
        let min_r = anchor.min(row);
        let max_r = anchor.max(row);
        self.view_state.selected = (min_r, 0);
        self.view_state.selection_end = Some((max_r, NUM_COLS - 1));
        cx.notify();
    }

    /// End row header drag
    pub fn end_row_header_drag(&mut self, _cx: &mut Context<Self>) {
        self.dragging_row_header = false;
        self.row_header_anchor = None;
    }

    /// Start column header drag - stores stable anchor
    pub fn start_col_header_drag(&mut self, col: usize, cx: &mut Context<Self>) {
        self.dragging_col_header = true;
        self.dragging_row_header = false;
        self.dragging_selection = false;
        self.col_header_anchor = Some(col);
        self.select_col(col, false, cx);
    }

    /// Continue column header drag - uses stored anchor
    pub fn continue_col_header_drag(&mut self, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_col_header { return; }
        let anchor = self.col_header_anchor.unwrap_or(col);
        let min_c = anchor.min(col);
        let max_c = anchor.max(col);
        self.view_state.selected = (0, min_c);
        self.view_state.selection_end = Some((NUM_ROWS - 1, max_c));
        cx.notify();
    }

    /// End column header drag
    pub fn end_col_header_drag(&mut self, _cx: &mut Context<Self>) {
        self.dragging_col_header = false;
        self.col_header_anchor = None;
    }

    /// Ctrl+click on row header - add row to additional selections
    pub fn ctrl_click_row(&mut self, row: usize, cx: &mut Context<Self>) {
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        self.select_row(row, false, cx);
    }

    /// Ctrl+click on column header - add column to additional selections
    pub fn ctrl_click_col(&mut self, col: usize, cx: &mut Context<Self>) {
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        self.select_col(col, false, cx);
    }
}
