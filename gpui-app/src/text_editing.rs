//! Text editing operations for Spreadsheet.
//!
//! This module contains methods for:
//! - Formula reference manipulation (inserting cell references during formula entry)
//! - Edit cursor movement and selection (character, word, home/end navigation)
//! - Formula bar hit-testing and caching
//! - Reference cycling (F4 to toggle $A$1 styles)

use gpui::*;
use crate::app::Spreadsheet;

/// Maximum rows in the spreadsheet
const NUM_ROWS: usize = 1_000_000;
/// Maximum columns in the spreadsheet
const NUM_COLS: usize = 16_384;

impl Spreadsheet {
    // ========================================================================
    // Formula Mode Reference Selection
    // ========================================================================

    /// Move formula reference with arrow keys (inserts or updates reference)
    pub fn formula_move_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when entering ref navigation (not editing text anymore)
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        let (new_row, new_col) = if let Some((row, col)) = self.formula_ref_cell {
            // Move existing reference
            let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
            (new_row, new_col)
        } else {
            // Start new reference from the selected cell (editing cell)
            let (sel_row, sel_col) = self.view_state.selected;
            let new_row = (sel_row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let new_col = (sel_col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
            (new_row, new_col)
        };

        // Update the reference
        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((new_row, new_col));
        self.formula_ref_end = None;  // Reset range when moving without shift

        // Insert or update the reference in the formula
        self.update_formula_reference(is_new);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Extend formula reference to range with Shift+arrow
    pub fn formula_extend_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when entering ref navigation
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        // Need an existing reference to extend
        let (anchor_row, anchor_col) = match self.formula_ref_cell {
            Some(cell) => cell,
            None => {
                // If no reference yet, start one first
                self.formula_move_ref(dr, dc, cx);
                return;
            }
        };

        // Get current end or use anchor as start
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((anchor_row, anchor_col));

        // Extend from the end position
        let new_row = (end_row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (end_col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

        self.formula_ref_end = Some((new_row, new_col));

        // Update the reference in the formula (not new, updating existing)
        self.update_formula_reference(false);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Insert formula reference on mouse click
    pub fn formula_click_ref(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when inserting ref via click
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((row, col));
        self.formula_ref_end = None;

        self.update_formula_reference(is_new);
        cx.notify();
    }

    /// Extend formula reference to range on Shift+click
    pub fn formula_shift_click_ref(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Need an existing reference to extend
        if self.formula_ref_cell.is_none() {
            // No reference yet, just insert single cell
            self.formula_click_ref(row, col, cx);
            return;
        }

        self.formula_ref_end = Some((row, col));
        self.update_formula_reference(false);
        cx.notify();
    }

    /// Extend formula reference to data boundary with Ctrl+Shift+arrow
    pub fn formula_extend_jump_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when entering ref navigation
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        // Need an existing reference to extend (or start one)
        let (anchor_row, anchor_col) = match self.formula_ref_cell {
            Some(cell) => cell,
            None => {
                // If no reference yet, start one first with a jump
                self.formula_jump_ref(dr, dc, cx);
                return;
            }
        };

        // Get current end or use anchor as start
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((anchor_row, anchor_col));

        // Jump to data boundary from end position
        let (new_row, new_col) = self.find_data_boundary(end_row, end_col, dr, dc);

        self.formula_ref_end = Some((new_row, new_col));
        self.update_formula_reference(false);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Move formula reference by jumping to data boundary (Ctrl+arrow in formula mode)
    pub fn formula_jump_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when entering ref navigation
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        let (start_row, start_col) = if let Some((row, col)) = self.formula_ref_cell {
            (row, col)
        } else {
            self.view_state.selected
        };

        let (new_row, new_col) = self.find_data_boundary(start_row, start_col, dr, dc);

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((new_row, new_col));
        self.formula_ref_end = None;

        self.update_formula_reference(is_new);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Start formula range drag selection
    pub fn formula_start_drag(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Close autocomplete when starting drag ref selection
        self.autocomplete_visible = false;
        self.autocomplete_suppressed = true;

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((row, col));
        self.formula_ref_end = None;
        self.dragging_selection = true;  // Reuse the drag flag

        self.update_formula_reference(is_new);
        cx.notify();
    }

    /// Continue formula range drag selection
    pub fn formula_continue_drag(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() || !self.dragging_selection {
            return;
        }

        if self.formula_ref_cell.is_none() {
            return;
        }

        // Only update if the cell changed
        if self.formula_ref_end != Some((row, col)) {
            self.formula_ref_end = Some((row, col));
            self.update_formula_reference(false);
            cx.notify();
        }
    }

    /// Update the formula string with the current reference
    fn update_formula_reference(&mut self, is_new: bool) {
        let Some((ref_row, ref_col)) = self.formula_ref_cell else {
            return;
        };

        // Build the reference string
        let ref_text = if let Some((end_row, end_col)) = self.formula_ref_end {
            Self::make_range_ref((ref_row, ref_col), (end_row, end_col))
        } else {
            Self::make_cell_ref(ref_row, ref_col)
        };

        if is_new {
            // Insert new reference at cursor (byte-indexed)
            let byte_idx = self.edit_cursor.min(self.edit_value.len());
            self.formula_ref_start_cursor = self.edit_cursor;
            self.edit_value.insert_str(byte_idx, &ref_text);
            self.edit_cursor += ref_text.len();  // Byte length
        } else {
            // Replace existing reference (from formula_ref_start_cursor to edit_cursor)
            // Both are already byte offsets
            let start_byte = self.formula_ref_start_cursor.min(self.edit_value.len());
            let end_byte = self.edit_cursor.min(self.edit_value.len());

            self.edit_value.replace_range(start_byte..end_byte, &ref_text);
            self.edit_cursor = start_byte + ref_text.len();  // Byte length
        }
        self.edit_scroll_dirty = true;
    }

    /// Ensure a cell is visible (scroll if necessary)
    pub(crate) fn ensure_cell_visible(&mut self, row: usize, col: usize) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // Adjust scroll to keep cell visible
        if row < self.view_state.scroll_row {
            self.view_state.scroll_row = row;
        } else if row >= self.view_state.scroll_row + visible_rows {
            self.view_state.scroll_row = row.saturating_sub(visible_rows - 1);
        }

        if col < self.view_state.scroll_col {
            self.view_state.scroll_col = col;
        } else if col >= self.view_state.scroll_col + visible_cols {
            self.view_state.scroll_col = col.saturating_sub(visible_cols - 1);
        }
    }

    // ========================================================================
    // Edit Cursor Movement (byte-indexed)
    // ========================================================================

    pub fn move_edit_cursor_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            self.edit_cursor = self.prev_char_boundary(self.edit_cursor);
            self.edit_selection_anchor = None;  // Clear selection
            self.edit_scroll_dirty = true;
            self.reset_caret_activity();
            cx.notify();
        }
    }

    pub fn move_edit_cursor_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                self.edit_cursor = self.next_char_boundary(self.edit_cursor);
                self.edit_selection_anchor = None;  // Clear selection
                self.edit_scroll_dirty = true;
                self.reset_caret_activity();
                cx.notify();
            }
        }
    }

    pub fn move_edit_cursor_home(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            self.edit_cursor = 0;
            self.edit_selection_anchor = None;  // Clear selection
            self.edit_scroll_dirty = true;
            self.reset_caret_activity();
            cx.notify();
        }
    }

    pub fn move_edit_cursor_end(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                self.edit_cursor = len;  // Byte offset at end
                self.edit_selection_anchor = None;  // Clear selection
                self.edit_scroll_dirty = true;
                self.reset_caret_activity();
                cx.notify();
            }
        }
    }

    // ========================================================================
    // Edit Cursor Selection (Shift+Arrow) - byte-indexed
    // ========================================================================

    pub fn select_edit_cursor_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = self.prev_char_boundary(self.edit_cursor);
            self.edit_scroll_dirty = true;
            cx.notify();
        }
    }

    pub fn select_edit_cursor_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                if self.edit_selection_anchor.is_none() {
                    self.edit_selection_anchor = Some(self.edit_cursor);
                }
                self.edit_cursor = self.next_char_boundary(self.edit_cursor);
                self.edit_scroll_dirty = true;
                cx.notify();
            }
        }
    }

    pub fn select_edit_cursor_home(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = 0;
            self.edit_scroll_dirty = true;
            cx.notify();
        }
    }

    pub fn select_edit_cursor_end(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                if self.edit_selection_anchor.is_none() {
                    self.edit_selection_anchor = Some(self.edit_cursor);
                }
                self.edit_cursor = len;  // Byte offset at end
                self.edit_scroll_dirty = true;
                cx.notify();
            }
        }
    }

    // ========================================================================
    // Byte-safe Cursor Navigation Helpers
    // ========================================================================
    //
    // All cursor positions are byte offsets into the UTF-8 buffer.
    //
    // NOTE: These helpers step by UTF-8 char boundaries, NOT grapheme clusters.
    // This means:
    // - é (e + combining accent U+0301) takes 2 cursor steps
    // - Some emoji (👨‍👩‍👧) are multiple codepoints and take multiple steps
    // This is acceptable for v1 - grapheme segmentation adds complexity.

    /// Find the previous char boundary (move cursor left by one character)
    pub(crate) fn prev_char_boundary(&self, byte_idx: usize) -> usize {
        if byte_idx == 0 {
            return 0;
        }
        let text = &self.edit_value;
        let mut idx = byte_idx - 1;
        while idx > 0 && !text.is_char_boundary(idx) {
            idx -= 1;
        }
        idx
    }

    /// Find the next char boundary (move cursor right by one character)
    pub(crate) fn next_char_boundary(&self, byte_idx: usize) -> usize {
        let text = &self.edit_value;
        let len = text.len();
        if byte_idx >= len {
            return len;
        }
        let mut idx = byte_idx + 1;
        while idx < len && !text.is_char_boundary(idx) {
            idx += 1;
        }
        idx
    }

    /// Get the char at a byte position (for word boundary detection)
    fn char_at_byte(&self, byte_idx: usize) -> Option<char> {
        self.edit_value[byte_idx..].chars().next()
    }

    // ========================================================================
    // Word Navigation Helpers (byte-indexed)
    // ========================================================================

    fn find_word_boundary_left(&self, from_byte: usize) -> usize {
        if from_byte == 0 {
            return 0;
        }
        let mut pos = self.prev_char_boundary(from_byte);

        // Skip whitespace/punctuation going left
        while pos > 0 {
            if let Some(c) = self.char_at_byte(pos) {
                if c.is_alphanumeric() {
                    break;
                }
            }
            pos = self.prev_char_boundary(pos);
        }

        // Skip word characters going left
        while pos > 0 {
            let prev = self.prev_char_boundary(pos);
            if let Some(c) = self.char_at_byte(prev) {
                if !c.is_alphanumeric() {
                    break;
                }
            }
            pos = prev;
        }
        pos
    }

    fn find_word_boundary_right(&self, from_byte: usize) -> usize {
        let text = &self.edit_value;
        let len = text.len();
        if from_byte >= len {
            return len;
        }
        let mut pos = from_byte;

        // Skip current word characters going right
        while pos < len {
            if let Some(c) = self.char_at_byte(pos) {
                if !c.is_alphanumeric() {
                    break;
                }
            }
            pos = self.next_char_boundary(pos);
        }

        // Skip whitespace/punctuation going right
        while pos < len {
            if let Some(c) = self.char_at_byte(pos) {
                if c.is_alphanumeric() {
                    break;
                }
            }
            pos = self.next_char_boundary(pos);
        }
        pos
    }

    // Ctrl+Arrow word navigation
    pub fn move_edit_cursor_word_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_cursor = self.find_word_boundary_left(self.edit_cursor);
            self.edit_selection_anchor = None;
            self.edit_scroll_dirty = true;
            self.reset_caret_activity();
            cx.notify();
        }
    }

    pub fn move_edit_cursor_word_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_cursor = self.find_word_boundary_right(self.edit_cursor);
            self.edit_selection_anchor = None;
            self.edit_scroll_dirty = true;
            self.reset_caret_activity();
            cx.notify();
        }
    }

    // Ctrl+Shift+Arrow word selection
    pub fn select_edit_cursor_word_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = self.find_word_boundary_left(self.edit_cursor);
            self.edit_scroll_dirty = true;
            cx.notify();
        }
    }

    pub fn select_edit_cursor_word_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = self.find_word_boundary_right(self.edit_cursor);
            self.edit_scroll_dirty = true;
            cx.notify();
        }
    }

    // ========================================================================
    // Edit Selection Range
    // ========================================================================

    /// Get current selection range (start, end) as byte offsets, or None.
    /// Returns normalized (min, max) range. Both endpoints are guaranteed to be
    /// valid char boundaries if created via movement helpers.
    pub fn edit_selection_range(&self) -> Option<(usize, usize)> {
        self.edit_selection_anchor.map(|anchor| {
            let start = anchor.min(self.edit_cursor);
            let end = anchor.max(self.edit_cursor);
            // Debug assert: verify both are valid char boundaries
            debug_assert!(
                self.edit_value.is_char_boundary(start) && self.edit_value.is_char_boundary(end),
                "Selection range ({}, {}) contains invalid char boundary in {:?}",
                start, end, self.edit_value
            );
            (start, end)
        })
    }

    /// Call after any caret/text change to update scroll if needed.
    /// Only does work if edit_scroll_dirty is set.
    pub fn update_edit_scroll(&mut self, window: &Window) {
        if !self.edit_scroll_dirty || !self.mode.is_editing() {
            return;
        }
        self.edit_scroll_dirty = false;

        let (_, col) = self.view_state.selected;
        let col_width = self.metrics.col_width(self.col_width(col));
        self.ensure_caret_visible(window, col_width);
    }

    /// Update edit_scroll_x to ensure the caret is visible within the cell.
    /// Only adjusts scroll when caret would go out of view - otherwise preserves position.
    /// This gives smooth "only when necessary" scrolling like Excel.
    fn ensure_caret_visible(&mut self, window: &Window, col_width: f32) {
        let text = &self.edit_value;
        let total_bytes = text.len();
        let cursor_byte = self.edit_cursor.min(total_bytes);

        // Measurements
        let padding = 4.0; // px_1 padding on each side
        let inner_w = col_width - (padding * 2.0);

        // Early exit: empty text or text fits in cell
        if total_bytes == 0 {
            self.edit_scroll_x = 0.0;
            return;
        }

        // Shape text to get accurate measurements
        let shape_text: SharedString = text.clone().into();
        let shape_len = shape_text.len();

        let shaped = window.text_system().shape_line(
            shape_text,
            px(self.metrics.font_size),
            &[TextRun {
                len: shape_len,
                font: Font::default(),
                color: Hsla::default(),
                background_color: None,
                underline: None,
                strikethrough: None,
            }],
            None,
        );

        let caret_x: f32 = shaped.x_for_index(cursor_byte).into();
        let text_w: f32 = shaped.x_for_index(total_bytes).into();

        // Text fits in cell - no scrolling needed
        if text_w <= inner_w {
            self.edit_scroll_x = 0.0;
            return;
        }

        // Current visual caret position (relative to cell inner area)
        let margin = 10.0; // Keep caret this far from edges
        let visual_caret = caret_x + self.edit_scroll_x;

        // Only adjust scroll if caret would be outside visible region
        if visual_caret < margin {
            // Caret off left edge - scroll right (make scroll_x less negative)
            self.edit_scroll_x = margin - caret_x;
        } else if visual_caret > inner_w - margin {
            // Caret off right edge - scroll left (make scroll_x more negative)
            self.edit_scroll_x = (inner_w - margin) - caret_x;
        }
        // else: caret is visible, don't change scroll

        // Clamp scroll to valid range
        let min_scroll = inner_w - text_w;
        self.edit_scroll_x = self.edit_scroll_x.min(0.0).max(min_scroll);

        // Debug asserts to catch sign mistakes
        debug_assert!(
            self.edit_scroll_x <= 0.01,
            "edit_scroll_x {} should be <= 0",
            self.edit_scroll_x
        );
        debug_assert!(
            self.edit_scroll_x >= min_scroll - 0.01,
            "edit_scroll_x {} below min_scroll {}",
            self.edit_scroll_x,
            min_scroll
        );
    }

    // =========================================================================
    // Formula Bar Hit-Testing and Caching
    // =========================================================================

    /// Rebuild the formula bar hit-testing cache (char boundaries + x positions).
    /// Call this when edit_value changes.
    pub fn rebuild_formula_bar_cache(&mut self, window: &Window) {
        let text = &self.edit_value;

        // Build char boundaries (byte offsets)
        let mut boundaries = vec![0];
        let mut byte_idx = 0;
        for c in text.chars() {
            byte_idx += c.len_utf8();
            boundaries.push(byte_idx);
        }

        // Shape once to get x positions (use same font as formula bar render)
        const FORMULA_BAR_FONT_SIZE: f32 = 14.0;
        let text_len = text.len();

        if text_len == 0 {
            self.formula_bar_char_boundaries = boundaries;
            self.formula_bar_boundary_xs = vec![0.0];
            self.formula_bar_text_width = 0.0;
            self.formula_bar_cache_dirty = false;
            return;
        }

        let shape_text: SharedString = text.clone().into();
        let shaped = window.text_system().shape_line(
            shape_text,
            px(FORMULA_BAR_FONT_SIZE),
            &[TextRun {
                len: text_len,
                font: Font::default(),
                color: Hsla::default(),
                background_color: None,
                underline: None,
                strikethrough: None,
            }],
            None,
        );

        // Cache x positions for each boundary
        let boundary_xs: Vec<f32> = boundaries
            .iter()
            .map(|&idx| shaped.x_for_index(idx).into())
            .collect();

        // Debug assert: xs should be monotonic
        debug_assert!(
            boundary_xs.windows(2).all(|w| w[0] <= w[1] + 0.01),
            "boundary_xs not monotonic - text shaping issue"
        );

        self.formula_bar_text_width = *boundary_xs.last().unwrap_or(&0.0);
        self.formula_bar_char_boundaries = boundaries;
        self.formula_bar_boundary_xs = boundary_xs;
        self.formula_bar_cache_dirty = false;
    }

    /// Rebuild formula bar cache if dirty.
    /// Call this before hit-testing (mouse clicks).
    pub fn maybe_rebuild_formula_bar_cache(&mut self, window: &Window) {
        if self.formula_bar_cache_dirty {
            self.rebuild_formula_bar_cache(window);
        }
    }

    /// Convert mouse x position to byte index in edit_value.
    /// Uses cached boundary positions for fast hit-testing.
    pub fn byte_index_for_x(&self, x: f32) -> usize {
        let boundaries = &self.formula_bar_char_boundaries;
        let xs = &self.formula_bar_boundary_xs;

        if boundaries.is_empty() || xs.is_empty() {
            return 0;
        }

        // Find first boundary whose x >= click_x using partition_point
        let i = xs.partition_point(|&bx| bx < x);

        let right_idx = i.min(boundaries.len() - 1);
        let left_idx = i.saturating_sub(1);

        let right = boundaries[right_idx];
        let left = boundaries[left_idx];

        let xr = xs[right_idx];
        let xl = xs[left_idx];

        // Return whichever boundary is closer (sticky correct)
        if (x - xl).abs() <= (xr - x).abs() {
            left
        } else {
            right
        }
    }

    /// Find word boundaries around a byte position in edit_value.
    /// Returns (word_start, word_end) as byte indices.
    pub fn word_bounds_at(&self, byte_pos: usize) -> (usize, usize) {
        let text = &self.edit_value;
        if text.is_empty() {
            return (0, 0);
        }

        let byte_pos = byte_pos.min(text.len());

        // Find start of word (scan backwards)
        let mut start = byte_pos;
        for (i, c) in text[..byte_pos].char_indices().rev() {
            if !c.is_alphanumeric() && c != '_' {
                start = i + c.len_utf8();
                break;
            }
            start = i;
        }

        // Find end of word (scan forwards)
        let mut end = byte_pos;
        for (i, c) in text[byte_pos..].char_indices() {
            if !c.is_alphanumeric() && c != '_' {
                end = byte_pos + i;
                break;
            }
            end = byte_pos + i + c.len_utf8();
        }

        (start, end)
    }

    /// Find the start of the previous word from byte position.
    pub fn prev_word_start(&self, byte_pos: usize) -> usize {
        let text = &self.edit_value;
        if text.is_empty() || byte_pos == 0 {
            return 0;
        }

        let byte_pos = byte_pos.min(text.len());

        // Skip whitespace/punctuation backwards
        let mut pos = byte_pos;
        for (i, c) in text[..byte_pos].char_indices().rev() {
            if c.is_alphanumeric() || c == '_' {
                pos = i + c.len_utf8();
                break;
            }
            pos = i;
        }

        // Now find start of word
        for (i, c) in text[..pos].char_indices().rev() {
            if !c.is_alphanumeric() && c != '_' {
                return i + c.len_utf8();
            }
        }
        0
    }

    /// Find the end of the next word from byte position.
    pub fn next_word_end(&self, byte_pos: usize) -> usize {
        let text = &self.edit_value;
        let len = text.len();
        if text.is_empty() || byte_pos >= len {
            return len;
        }

        // Skip whitespace/punctuation forwards
        let mut pos = byte_pos;
        for (i, c) in text[byte_pos..].char_indices() {
            if c.is_alphanumeric() || c == '_' {
                pos = byte_pos + i;
                break;
            }
            pos = byte_pos + i + c.len_utf8();
        }

        // Now find end of word
        for (i, c) in text[pos..].char_indices() {
            if !c.is_alphanumeric() && c != '_' {
                return pos + i;
            }
        }
        len
    }

    /// Ensure formula bar caret is visible by adjusting formula_bar_scroll_x.
    pub fn ensure_formula_bar_caret_visible(&mut self, window: &Window) {
        // Rebuild cache if dirty
        if self.formula_bar_cache_dirty {
            self.rebuild_formula_bar_cache(window);
        }

        let text = &self.edit_value;
        if text.is_empty() {
            self.formula_bar_scroll_x = 0.0;
            return;
        }

        // Formula bar visible width (approximate - full width minus cell ref and fx button)
        // This will be refined when we have actual layout info
        let visible_width = 400.0; // Conservative estimate
        let margin = 10.0;

        // Find caret x position from cache
        let cursor_byte = self.edit_cursor.min(text.len());
        let boundaries = &self.formula_bar_char_boundaries;
        let xs = &self.formula_bar_boundary_xs;

        // Find boundary index for cursor
        let boundary_idx = boundaries
            .iter()
            .position(|&b| b >= cursor_byte)
            .unwrap_or(boundaries.len().saturating_sub(1));

        let caret_x = xs.get(boundary_idx).copied().unwrap_or(0.0);
        let text_width = self.formula_bar_text_width;

        // Text fits - no scrolling needed
        if text_width <= visible_width {
            self.formula_bar_scroll_x = 0.0;
            return;
        }

        // Current visual caret position
        let visual_caret = caret_x + self.formula_bar_scroll_x;

        // Adjust scroll if caret outside visible region
        if visual_caret < margin {
            self.formula_bar_scroll_x = margin - caret_x;
        } else if visual_caret > visible_width - margin {
            self.formula_bar_scroll_x = (visible_width - margin) - caret_x;
        }

        // Clamp scroll to valid range
        let min_scroll = visible_width - text_width;
        self.formula_bar_scroll_x = self.formula_bar_scroll_x.min(0.0).max(min_scroll);
    }

    // =========================================================================
    // Select All and Reference Cycling
    // =========================================================================

    /// Select all text in edit mode (byte-indexed)
    pub fn select_all_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_selection_anchor = Some(0);
            self.edit_cursor = self.edit_value.len();  // Byte offset at end
            self.edit_scroll_dirty = true;
            cx.notify();
        }
    }

    /// F4: Cycle cell reference at cursor through A1 → $A$1 → A$1 → $A1 → A1
    pub fn cycle_reference(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            return;
        }

        // Cell reference pattern: optional $ + column letters + optional $ + row numbers
        let re = regex::Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

        // edit_cursor is already a byte offset
        let cursor_byte = self.edit_cursor.min(self.edit_value.len());

        // Find reference at or near cursor
        let mut best_match: Option<(usize, usize, regex::Captures)> = None;

        for caps in re.captures_iter(&self.edit_value) {
            let m = caps.get(0).unwrap();
            let start = m.start();
            let end = m.end();

            // Check if cursor is within or immediately after this reference
            if cursor_byte >= start && cursor_byte <= end {
                best_match = Some((start, end, caps));
                break;
            }
            // Also check if cursor is just before the reference (user may have cursor at start)
            if cursor_byte == start {
                best_match = Some((start, end, caps));
                break;
            }
        }

        // If no direct match, find the nearest reference before cursor
        if best_match.is_none() {
            for caps in re.captures_iter(&self.edit_value) {
                let m = caps.get(0).unwrap();
                let start = m.start();
                let end = m.end();

                if end <= cursor_byte {
                    best_match = Some((start, end, caps));
                }
            }
        }

        if let Some((start, end, caps)) = best_match {
            let col_dollar = caps.get(1).map(|m| m.as_str()).unwrap_or("");
            let col_letters = caps.get(2).map(|m| m.as_str()).unwrap_or("");
            let row_dollar = caps.get(3).map(|m| m.as_str()).unwrap_or("");
            let row_numbers = caps.get(4).map(|m| m.as_str()).unwrap_or("");

            // Determine current state and cycle to next
            // State 0: A1 (relative, relative)
            // State 1: $A$1 (absolute, absolute)
            // State 2: A$1 (relative col, absolute row)
            // State 3: $A1 (absolute col, relative row)
            let current_state = match (col_dollar.is_empty(), row_dollar.is_empty()) {
                (true, true) => 0,    // A1
                (false, false) => 1,  // $A$1
                (true, false) => 2,   // A$1
                (false, true) => 3,   // $A1
            };

            let next_state = (current_state + 1) % 4;

            let new_ref = match next_state {
                0 => format!("{}{}", col_letters, row_numbers),           // A1
                1 => format!("${}${}", col_letters, row_numbers),         // $A$1
                2 => format!("{}${}", col_letters, row_numbers),          // A$1
                3 => format!("${}{}", col_letters, row_numbers),          // $A1
                _ => unreachable!(),
            };

            // Replace the reference in edit_value (all byte-indexed)
            let old_ref_bytes = end - start;
            self.edit_value.replace_range(start..end, &new_ref);
            let new_ref_bytes = new_ref.len();

            // Adjust cursor if it was after or within the replaced region
            if self.edit_cursor > start {
                // Cursor was within or after the reference
                if self.edit_cursor <= start + old_ref_bytes {
                    // Cursor was within reference - move to end of new reference
                    self.edit_cursor = start + new_ref_bytes;
                } else {
                    // Cursor was after reference - adjust by length difference
                    let diff = new_ref_bytes as i32 - old_ref_bytes as i32;
                    self.edit_cursor = (self.edit_cursor as i32 + diff) as usize;
                }
            }

            cx.notify();
        }
    }
}
