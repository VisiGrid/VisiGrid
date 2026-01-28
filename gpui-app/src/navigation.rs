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
//!
//! ## Filter-Aware Navigation
//!
//! When filtering is active, navigation operates in VIEW space:
//! - Arrow keys skip hidden rows (move through visible_rows)
//! - Ctrl+Arrow stops at visible boundaries
//! - Page Up/Down move by visible rows
//!
//! The `visible_rows` cache in RowView provides O(1) access to the next/previous
//! visible row via binary search.

use gpui::{*};
use crate::app::{Spreadsheet, NUM_ROWS, NUM_COLS};

impl Spreadsheet {
    // =========================================================================
    // Filter-Aware Navigation Helpers
    // =========================================================================

    /// Find the next visible row in a direction.
    /// Returns the new view row, or the current row if at boundary.
    ///
    /// - `current_row`: Current view row
    /// - `delta`: Direction (+1 for down, -1 for up)
    pub(crate) fn next_visible_row(&self, current_row: usize, delta: i32) -> usize {
        let visible = self.row_view.visible_rows();

        // If not filtered, simple arithmetic
        if !self.row_view.is_filtered() {
            return (current_row as i32 + delta).max(0).min(NUM_ROWS as i32 - 1) as usize;
        }

        // Find current position in visible_rows
        // Use binary search since visible_rows is sorted
        let current_idx = match visible.binary_search(&current_row) {
            Ok(idx) => idx,
            Err(idx) => {
                // Current row is hidden - find nearest visible
                if delta > 0 {
                    // Moving down: use the row at insertion point (or last)
                    idx.min(visible.len().saturating_sub(1))
                } else {
                    // Moving up: use the row before insertion point (or first)
                    idx.saturating_sub(1)
                }
            }
        };

        // Move by delta steps in visible_rows
        let new_idx = if delta > 0 {
            (current_idx + delta as usize).min(visible.len().saturating_sub(1))
        } else {
            current_idx.saturating_sub((-delta) as usize)
        };

        visible.get(new_idx).copied().unwrap_or(current_row)
    }

    /// Check if a view row is visible (not hidden by filter)
    fn is_row_visible(&self, view_row: usize) -> bool {
        self.row_view.is_view_row_visible(view_row)
    }

    // =========================================================================
    // Cell Movement
    // =========================================================================

    pub fn move_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Close validation dropdown when selection changes
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::SelectionChanged,
            cx,
        );

        let (row, col) = self.active_view_state().selected;

        // For vertical movement, use filter-aware navigation
        let new_row = if dr != 0 {
            self.next_visible_row(row, dr)
        } else {
            row
        };

        // For horizontal movement, simple arithmetic (columns aren't filtered)
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

        let view_state = self.active_view_state_mut();
        view_state.selected = (new_row, new_col);
        view_state.selection_end = None;  // Clear range selection
        view_state.additional_selections.clear();  // Clear discontiguous selections

        self.nav_perf.mark_state_updated();
        self.ensure_visible(cx);
    }

    pub fn extend_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let view_state = self.active_view_state();
        let (row, col) = view_state.selection_end.unwrap_or(view_state.selected);

        // For vertical movement, use filter-aware navigation
        let new_row = if dr != 0 {
            self.next_visible_row(row, dr)
        } else {
            row
        };

        // For horizontal movement, simple arithmetic
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

        self.active_view_state_mut().selection_end = Some((new_row, new_col));

        self.ensure_visible(cx);
    }

    pub fn page_up(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        // Move by screen-visible rows, but through filter-visible rows
        let page_size = self.visible_rows() as i32;
        self.move_selection(-page_size, 0, cx);
    }

    pub fn page_down(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        // Move by screen-visible rows, but through filter-visible rows
        let page_size = self.visible_rows() as i32;
        self.move_selection(page_size, 0, cx);
    }

    // =========================================================================
    // Jump Navigation (Ctrl+Arrow)
    // =========================================================================

    /// Find the data boundary in a direction (used by Ctrl+Arrow and Ctrl+Shift+Arrow)
    /// When filtering is active, only considers visible rows.
    pub(crate) fn find_data_boundary(&self, start_row: usize, start_col: usize, dr: i32, dc: i32, cx: &App) -> (usize, usize) {
        let mut row = start_row;
        let mut col = start_col;

        // Get cell value, converting view row to data row
        let get_cell_value = |view_row: usize, c: usize| -> bool {
            let data_row = self.row_view.view_to_data(view_row);
            self.sheet(cx).get_cell(data_row, c).value.raw_display().is_empty()
        };

        let current_empty = get_cell_value(row, col);

        // For vertical movement with filtering, use visible rows
        if dr != 0 && self.row_view.is_filtered() {
            let visible = self.row_view.visible_rows();

            // Find current position in visible_rows
            let current_idx = match visible.binary_search(&row) {
                Ok(idx) => idx,
                Err(idx) => idx.min(visible.len().saturating_sub(1)),
            };

            // Peek at next visible row
            let peek_idx = if dr > 0 {
                (current_idx + 1).min(visible.len().saturating_sub(1))
            } else {
                current_idx.saturating_sub(1)
            };
            let peek_row = visible.get(peek_idx).copied().unwrap_or(row);
            let next_empty = if peek_row == row {
                true // At edge
            } else {
                get_cell_value(peek_row, col)
            };

            let looking_for_nonempty = current_empty || next_empty;

            // Scan through visible rows only
            let mut idx = current_idx;
            loop {
                let next_idx = if dr > 0 {
                    idx + 1
                } else {
                    if idx == 0 { break; }
                    idx - 1
                };

                if next_idx >= visible.len() {
                    break;
                }

                let next_row = visible[next_idx];
                let cell_empty = get_cell_value(next_row, col);

                if looking_for_nonempty {
                    row = next_row;
                    idx = next_idx;
                    if !cell_empty {
                        break;
                    }
                } else {
                    if cell_empty {
                        break;
                    }
                    row = next_row;
                    idx = next_idx;
                }
            }

            return (row, col);
        }

        // Horizontal movement or no filtering - original logic
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            get_cell_value(peek_row, peek_col)
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

            // Skip hidden rows when moving vertically
            if dr != 0 && !self.is_row_visible(next_row) {
                row = next_row;
                continue;
            }

            let cell_empty = get_cell_value(next_row, next_col);

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
    /// When filtering is active, stops at visible row boundaries.
    pub fn jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selected;
        let (new_row, new_col) = self.find_data_boundary(row, col, dr, dc, cx);

        self.view_state.selected = (new_row, new_col);
        self.view_state.selection_end = None;
        self.ensure_visible(cx);
    }

    /// Extend selection to edge of data region (Excel-style Ctrl+Shift+Arrow)
    /// When filtering is active, stops at visible row boundaries.
    pub fn extend_jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Start from current selection end (or selected if no selection)
        let (row, col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let (new_row, new_col) = self.find_data_boundary(row, col, dr, dc, cx);

        // Extend selection to this point (don't move selected, just selection_end)
        self.view_state.selection_end = Some((new_row, new_col));
        self.ensure_visible(cx);
    }

    // =========================================================================
    // Visibility / Scrolling
    // =========================================================================

    /// Mark that the scroll position needs to be adjusted to keep the
    /// selection visible. The actual adjustment is deferred to the start
    /// of the next render pass, which naturally coalesces multiple
    /// navigation events within a single frame into one scroll update.
    pub fn ensure_visible(&mut self, cx: &mut Context<Self>) {
        self.nav_scroll_dirty = true;
        cx.notify();
    }

    /// Perform the actual scroll adjustment to keep the active cell visible.
    /// Called once at the start of `Render::render()` to flush any deferred
    /// scroll updates — this is the coalescing point.
    pub(crate) fn flush_nav_scroll(&mut self) {
        if !self.nav_scroll_dirty {
            return;
        }
        self.nav_scroll_dirty = false;

        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        let view_state = self.active_view_state_mut();
        let (row, col) = view_state.selection_end.unwrap_or(view_state.selected);

        // When freeze panes are active, calculate scrollable region
        let scrollable_visible_rows = visible_rows.saturating_sub(view_state.frozen_rows);
        let scrollable_visible_cols = visible_cols.saturating_sub(view_state.frozen_cols);

        // Vertical scroll - frozen rows are always visible, only scroll for rows in scrollable region
        if row < view_state.frozen_rows {
            // Row is in frozen region - always visible, but ensure scroll_row is valid
            view_state.scroll_row = view_state.scroll_row.max(view_state.frozen_rows);
        } else if row < view_state.scroll_row {
            view_state.scroll_row = row;
        } else if scrollable_visible_rows > 0 && row >= view_state.scroll_row + scrollable_visible_rows {
            view_state.scroll_row = row - scrollable_visible_rows + 1;
        }

        // Horizontal scroll - frozen cols are always visible, only scroll for cols in scrollable region
        if col < view_state.frozen_cols {
            // Col is in frozen region - always visible, but ensure scroll_col is valid
            view_state.scroll_col = view_state.scroll_col.max(view_state.frozen_cols);
        } else if col < view_state.scroll_col {
            view_state.scroll_col = col;
        } else if scrollable_visible_cols > 0 && col >= view_state.scroll_col + scrollable_visible_cols {
            view_state.scroll_col = col - scrollable_visible_cols + 1;
        }

        // Ensure scroll positions don't go below freeze bounds
        view_state.scroll_row = view_state.scroll_row.max(view_state.frozen_rows);
        view_state.scroll_col = view_state.scroll_col.max(view_state.frozen_cols);
    }

    pub fn scroll(&mut self, delta_rows: i32, delta_cols: i32, cx: &mut Context<Self>) {
        // Close validation dropdown on scroll
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::Scroll,
            cx,
        );

        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        let view_state = self.active_view_state_mut();
        // When freeze panes are active, scrollable region starts after frozen rows/cols
        let min_scroll_row = view_state.frozen_rows;
        let min_scroll_col = view_state.frozen_cols;

        let new_row = (view_state.scroll_row as i32 + delta_rows)
            .max(min_scroll_row as i32)
            .min((NUM_ROWS.saturating_sub(visible_rows)) as i32) as usize;
        let new_col = (view_state.scroll_col as i32 + delta_cols)
            .max(min_scroll_col as i32)
            .min((NUM_COLS.saturating_sub(visible_cols)) as i32) as usize;

        if new_row != view_state.scroll_row || new_col != view_state.scroll_col {
            view_state.scroll_row = new_row;
            view_state.scroll_col = new_col;
            cx.notify();
        }
    }

    // =========================================================================
    // Cell Selection
    // =========================================================================

    pub fn select_cell(&mut self, row: usize, col: usize, extend: bool, cx: &mut Context<Self>) {
        // Mouse click breaks the tab-chain
        self.tab_chain_origin_col = None;
        if extend {
            self.active_view_state_mut().selection_end = Some((row, col));
        } else {
            {
                let view_state = self.active_view_state_mut();
                view_state.selected = (row, col);
                view_state.selection_end = None;
                view_state.additional_selections.clear();  // Clear Ctrl+Click selections
            }
            // Clear trace path when selection changes (unless inspector is pinned)
            if self.inspector_pinned.is_none() && self.inspector_trace_path.is_some() {
                self.inspector_trace_path = None;
                self.inspector_trace_incomplete = false;
            }
        }
        // Update dependency trace if enabled
        self.recompute_trace_if_needed(cx);
        cx.notify();
    }

    /// Ctrl+Click to add/toggle cell in selection (discontiguous selection)
    pub fn ctrl_click_cell(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        let view_state = self.active_view_state_mut();
        // Save current selection to additional_selections
        view_state.additional_selections.push((view_state.selected, view_state.selection_end));
        // Start new selection at clicked cell
        view_state.selected = (row, col);
        view_state.selection_end = None;
        cx.notify();
    }

    pub fn select_all(&mut self, cx: &mut Context<Self>) {
        let view_state = self.active_view_state_mut();
        view_state.selected = (0, 0);
        view_state.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        view_state.additional_selections.clear();  // Clear discontiguous selections
        cx.notify();
    }

    // =========================================================================
    // Drag Selection
    // =========================================================================

    /// Start drag selection - called on mouse_down
    pub fn start_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        let view_state = self.active_view_state_mut();
        view_state.selected = (row, col);
        view_state.selection_end = None;
        view_state.additional_selections.clear();  // Clear Ctrl+Click selections on new drag
        cx.notify();
    }

    /// Start drag selection with Ctrl held (add to existing selections)
    pub fn start_ctrl_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        let view_state = self.active_view_state_mut();
        // Save current selection to additional_selections
        view_state.additional_selections.push((view_state.selected, view_state.selection_end));
        // Start new selection at clicked cell
        view_state.selected = (row, col);
        view_state.selection_end = None;
        cx.notify();
    }

    /// Continue drag selection - called on mouse_move while dragging
    pub fn continue_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_selection {
            return;
        }
        // Only update if the cell changed to avoid unnecessary redraws
        let view_state = self.active_view_state_mut();
        if view_state.selection_end != Some((row, col)) {
            view_state.selection_end = Some((row, col));
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
        let view_state = self.active_view_state();
        let start = view_state.selected;
        let end = view_state.selection_end.unwrap_or(start);
        let min_row = start.0.min(end.0);
        let max_row = start.0.max(end.0);
        let min_col = start.1.min(end.1);
        let max_col = start.1.max(end.1);
        ((min_row, min_col), (max_row, max_col))
    }

    /// Clamp selection to valid bounds after operations that might invalidate it.
    /// Preserves column where possible (user mental model), clamps row to valid range.
    pub fn clamp_selection(&mut self) {
        let view_state = self.active_view_state_mut();
        // Clamp selected cell
        view_state.selected.0 = view_state.selected.0.min(NUM_ROWS - 1);
        view_state.selected.1 = view_state.selected.1.min(NUM_COLS - 1);

        // Clamp selection_end if present
        if let Some(ref mut end) = view_state.selection_end {
            end.0 = end.0.min(NUM_ROWS - 1);
            end.1 = end.1.min(NUM_COLS - 1);
        }

        // Clamp additional selections
        for (start, end) in &mut view_state.additional_selections {
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
        let view_state = self.active_view_state();
        // Check if primary selection is a range (more than one cell)
        if let Some(end) = view_state.selection_end {
            if end != view_state.selected {
                return true;
            }
        }
        // Check if there are additional Ctrl+Click selections
        if !view_state.additional_selections.is_empty() {
            return true;
        }
        false
    }

    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        let view_state = self.active_view_state();
        // Check active selection
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            return true;
        }
        // Check additional selections (Ctrl+Click ranges)
        for (start, end) in &view_state.additional_selections {
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
