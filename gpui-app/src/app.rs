use gpui::*;
use std::collections::HashMap;
use std::path::PathBuf;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::CellLookup;

use crate::history::{History, CellChange};
use crate::mode::Mode;
use crate::views;

// Grid configuration
pub const NUM_ROWS: usize = 65536;
pub const NUM_COLS: usize = 256;
pub const CELL_WIDTH: f32 = 80.0;
pub const CELL_HEIGHT: f32 = 24.0;
pub const HEADER_WIDTH: f32 = 50.0;
pub const MENU_BAR_HEIGHT: f32 = 28.0;
pub const FORMULA_BAR_HEIGHT: f32 = 28.0;
pub const COLUMN_HEADER_HEIGHT: f32 = 24.0;
pub const STATUS_BAR_HEIGHT: f32 = 24.0;

pub struct Spreadsheet {
    // Core data
    pub workbook: Workbook,
    pub history: History,

    // Selection
    pub selected: (usize, usize),                              // Anchor of active selection
    pub selection_end: Option<(usize, usize)>,                 // End of active range selection
    pub additional_selections: Vec<((usize, usize), Option<(usize, usize)>)>,  // Ctrl+Click ranges

    // Viewport
    pub scroll_row: usize,
    pub scroll_col: usize,

    // Mode & editing
    pub mode: Mode,
    pub edit_value: String,
    pub edit_original: String,
    pub goto_input: String,
    pub find_input: String,
    pub find_results: Vec<(usize, usize)>,
    pub find_index: usize,

    // Command palette
    pub palette_query: String,
    pub palette_selected: usize,

    // Clipboard
    pub clipboard: Option<String>,

    // File state
    pub current_file: Option<PathBuf>,
    pub is_modified: bool,

    // UI state
    pub focus_handle: FocusHandle,
    pub status_message: Option<String>,
    pub window_size: Size<Pixels>,

    // Column/row sizing
    pub col_widths: HashMap<usize, f32>,   // Custom column widths (default: CELL_WIDTH)
    pub row_heights: HashMap<usize, f32>,  // Custom row heights (default: CELL_HEIGHT)

    // Resize drag state
    pub resizing_col: Option<usize>,       // Column being resized (by right edge)
    pub resizing_row: Option<usize>,       // Row being resized (by bottom edge)
    pub resize_start_pos: f32,             // Mouse position at drag start
    pub resize_start_size: f32,            // Original size at drag start

    // Menu bar state (Excel 2003 style dropdown menus)
    pub open_menu: Option<crate::mode::Menu>,

    // Sheet tab state
    pub renaming_sheet: Option<usize>,     // Index of sheet being renamed
    pub sheet_rename_input: String,        // Current rename input value
    pub sheet_context_menu: Option<usize>, // Index of sheet with open context menu

    // Font picker state
    pub available_fonts: Vec<String>,      // System fonts
    pub font_picker_query: String,         // Filter query
    pub font_picker_selected: usize,       // Selected item index

    // Drag selection state
    pub dragging_selection: bool,          // Currently dragging to select cells
}

impl Spreadsheet {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let workbook = Workbook::new();

        let focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);
        let window_size = window.viewport_size();

        Self {
            workbook,
            history: History::new(),
            selected: (0, 0),
            selection_end: None,
            additional_selections: Vec::new(),
            scroll_row: 0,
            scroll_col: 0,
            mode: Mode::Navigation,
            edit_value: String::new(),
            edit_original: String::new(),
            goto_input: String::new(),
            find_input: String::new(),
            find_results: Vec::new(),
            find_index: 0,
            palette_query: String::new(),
            palette_selected: 0,
            clipboard: None,
            current_file: None,
            is_modified: false,
            focus_handle,
            status_message: None,
            window_size,
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            resizing_col: None,
            resizing_row: None,
            resize_start_pos: 0.0,
            resize_start_size: 0.0,
            open_menu: None,
            renaming_sheet: None,
            sheet_rename_input: String::new(),
            sheet_context_menu: None,
            available_fonts: Self::enumerate_fonts(),
            font_picker_query: String::new(),
            font_picker_selected: 0,
            dragging_selection: false,
        }
    }

    /// Enumerate available system fonts
    fn enumerate_fonts() -> Vec<String> {
        // Fonts commonly installed on Linux systems
        // TODO: Could use fontconfig to enumerate dynamically
        vec![
            "Adwaita Mono".to_string(),
            "Adwaita Sans".to_string(),
            "CaskaydiaMono Nerd Font".to_string(),
            "iA Writer Mono S".to_string(),
            "iA Writer Duo S".to_string(),
            "iA Writer Quattro S".to_string(),
            "Liberation Mono".to_string(),
            "Liberation Sans".to_string(),
            "Liberation Serif".to_string(),
            "Nimbus Mono PS".to_string(),
            "Nimbus Sans".to_string(),
            "Nimbus Roman".to_string(),
            "Noto Sans Mono".to_string(),
        ]
    }

    // Menu methods
    pub fn toggle_menu(&mut self, menu: crate::mode::Menu, cx: &mut Context<Self>) {
        if self.open_menu == Some(menu) {
            self.open_menu = None;
        } else {
            self.open_menu = Some(menu);
        }
        cx.notify();
    }

    pub fn close_menu(&mut self, cx: &mut Context<Self>) {
        if self.open_menu.is_some() {
            self.open_menu = None;
            cx.notify();
        }
    }

    // Sheet access convenience methods
    /// Get a reference to the active sheet
    pub fn sheet(&self) -> &Sheet {
        self.workbook.active_sheet()
    }

    /// Get a mutable reference to the active sheet
    pub fn sheet_mut(&mut self) -> &mut Sheet {
        self.workbook.active_sheet_mut()
    }

    // Sheet navigation methods
    /// Move to the next sheet
    pub fn next_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.next_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.prev_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Switch to a specific sheet by index
    pub fn goto_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        if self.workbook.set_active_sheet(index) {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Add a new sheet and switch to it
    pub fn add_sheet(&mut self, cx: &mut Context<Self>) {
        let new_index = self.workbook.add_sheet();
        self.workbook.set_active_sheet(new_index);
        self.clear_selection_state();
        self.is_modified = true;
        cx.notify();
    }

    /// Clear selection state when switching sheets
    fn clear_selection_state(&mut self) {
        self.selected = (0, 0);
        self.selection_end = None;
        self.scroll_row = 0;
        self.scroll_col = 0;
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
    }

    // Sheet rename methods
    /// Start renaming a sheet (double-click on tab)
    pub fn start_sheet_rename(&mut self, index: usize, cx: &mut Context<Self>) {
        if let Some(name) = self.workbook.sheet_names().get(index) {
            self.renaming_sheet = Some(index);
            self.sheet_rename_input = name.to_string();
            self.sheet_context_menu = None;
            cx.notify();
        }
    }

    /// Confirm the sheet rename
    pub fn confirm_sheet_rename(&mut self, cx: &mut Context<Self>) {
        if let Some(index) = self.renaming_sheet {
            let new_name = self.sheet_rename_input.trim();
            if !new_name.is_empty() {
                self.workbook.rename_sheet(index, new_name);
                self.is_modified = true;
            }
            self.renaming_sheet = None;
            self.sheet_rename_input.clear();
            cx.notify();
        }
    }

    /// Cancel the sheet rename
    pub fn cancel_sheet_rename(&mut self, cx: &mut Context<Self>) {
        self.renaming_sheet = None;
        self.sheet_rename_input.clear();
        cx.notify();
    }

    /// Handle input for sheet rename
    pub fn sheet_rename_input_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.push(c);
            cx.notify();
        }
    }

    /// Handle backspace for sheet rename
    pub fn sheet_rename_backspace(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.pop();
            cx.notify();
        }
    }

    // Sheet context menu methods
    /// Show context menu for a sheet tab
    pub fn show_sheet_context_menu(&mut self, index: usize, cx: &mut Context<Self>) {
        self.sheet_context_menu = Some(index);
        self.renaming_sheet = None;
        cx.notify();
    }

    /// Hide sheet context menu
    pub fn hide_sheet_context_menu(&mut self, cx: &mut Context<Self>) {
        self.sheet_context_menu = None;
        cx.notify();
    }

    /// Delete a sheet
    pub fn delete_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        if self.workbook.delete_sheet(index) {
            self.is_modified = true;
            self.sheet_context_menu = None;
            cx.notify();
        } else {
            self.status_message = Some("Cannot delete the last sheet".to_string());
            self.sheet_context_menu = None;
            cx.notify();
        }
    }

    /// Get width for a column (custom or default)
    pub fn col_width(&self, col: usize) -> f32 {
        *self.col_widths.get(&col).unwrap_or(&CELL_WIDTH)
    }

    /// Get height for a row (custom or default)
    pub fn row_height(&self, row: usize) -> f32 {
        *self.row_heights.get(&row).unwrap_or(&CELL_HEIGHT)
    }

    /// Set column width
    pub fn set_col_width(&mut self, col: usize, width: f32) {
        let width = width.max(20.0).min(500.0); // Clamp between 20-500px
        if (width - CELL_WIDTH).abs() < 1.0 {
            self.col_widths.remove(&col); // Remove if close to default
        } else {
            self.col_widths.insert(col, width);
        }
    }

    /// Set row height
    pub fn set_row_height(&mut self, row: usize, height: f32) {
        let height = height.max(12.0).min(200.0); // Clamp between 12-200px
        if (height - CELL_HEIGHT).abs() < 1.0 {
            self.row_heights.remove(&row); // Remove if close to default
        } else {
            self.row_heights.insert(row, height);
        }
    }

    /// Get the X position of a column's left edge (relative to start of grid, after row header)
    pub fn col_x_offset(&self, target_col: usize) -> f32 {
        let mut x = 0.0;
        for col in self.scroll_col..target_col {
            x += self.col_width(col);
        }
        x
    }

    /// Get the Y position of a row's top edge (relative to start of grid, after column header)
    pub fn row_y_offset(&self, target_row: usize) -> f32 {
        let mut y = 0.0;
        for row in self.scroll_row..target_row {
            y += self.row_height(row);
        }
        y
    }

    /// Auto-fit column width to content
    pub fn auto_fit_col_width(&mut self, col: usize, cx: &mut Context<Self>) {
        let mut max_width: f32 = 40.0; // Minimum width

        // Check all rows for content in this column
        for row in 0..NUM_ROWS {
            let text = self.sheet().get_text(row, col);
            if !text.is_empty() {
                // Estimate width: ~7px per character + padding
                let estimated_width = text.len() as f32 * 7.5 + 16.0;
                max_width = max_width.max(estimated_width);
            }
        }

        self.set_col_width(col, max_width);
        cx.notify();
    }

    /// Auto-fit row height to content (for multi-line text in future)
    pub fn auto_fit_row_height(&mut self, row: usize, cx: &mut Context<Self>) {
        // For now, just reset to default since we don't support multi-line
        self.row_heights.remove(&row);
        cx.notify();
    }

    /// Calculate visible rows based on window height
    pub fn visible_rows(&self) -> usize {
        let height: f32 = self.window_size.height.into();
        let available_height = height
            - MENU_BAR_HEIGHT
            - FORMULA_BAR_HEIGHT
            - COLUMN_HEADER_HEIGHT
            - STATUS_BAR_HEIGHT;
        let rows = (available_height / CELL_HEIGHT).floor() as usize;
        rows.max(1).min(NUM_ROWS)
    }

    /// Calculate visible columns based on window width
    pub fn visible_cols(&self) -> usize {
        let width: f32 = self.window_size.width.into();
        let available_width = width - HEADER_WIDTH;
        let cols = (available_width / CELL_WIDTH).floor() as usize;
        cols.max(1).min(NUM_COLS)
    }

    /// Update window size (called on resize)
    pub fn update_window_size(&mut self, size: Size<Pixels>, cx: &mut Context<Self>) {
        self.window_size = size;
        cx.notify();
    }

    // Column letter (A, B, ..., Z, AA, AB, ...)
    pub fn col_letter(col: usize) -> String {
        let mut result = String::new();
        let mut c = col;
        loop {
            result.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 { break; }
            c = c / 26 - 1;
        }
        result
    }

    // Cell reference (A1, B2, etc.)
    pub fn cell_ref(&self) -> String {
        format!("{}{}", Self::col_letter(self.selected.1), self.selected.0 + 1)
    }

    // Navigation
    pub fn move_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.selected = (new_row, new_col);
        self.selection_end = None;  // Clear range selection
        self.additional_selections.clear();  // Clear discontiguous selections

        self.ensure_visible(cx);
    }

    pub fn extend_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selection_end.unwrap_or(self.selected);
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.selection_end = Some((new_row, new_col));

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

    /// Jump to edge of data region or sheet boundary (Excel-style Ctrl+Arrow)
    pub fn jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (mut row, mut col) = self.selected;
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

        self.selected = (row, col);
        self.selection_end = None;
        self.ensure_visible(cx);
    }

    /// Extend selection to edge of data region (Excel-style Ctrl+Shift+Arrow)
    pub fn extend_jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Start from current selection end (or selected if no selection)
        let (mut row, mut col) = self.selection_end.unwrap_or(self.selected);
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
        self.selection_end = Some((row, col));
        self.ensure_visible(cx);
    }

    fn ensure_visible(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.selection_end.unwrap_or(self.selected);
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // Vertical scroll
        if row < self.scroll_row {
            self.scroll_row = row;
        } else if row >= self.scroll_row + visible_rows {
            self.scroll_row = row - visible_rows + 1;
        }

        // Horizontal scroll
        if col < self.scroll_col {
            self.scroll_col = col;
        } else if col >= self.scroll_col + visible_cols {
            self.scroll_col = col - visible_cols + 1;
        }

        cx.notify();
    }

    pub fn select_cell(&mut self, row: usize, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            self.selection_end = Some((row, col));
        } else {
            self.selected = (row, col);
            self.selection_end = None;
            self.additional_selections.clear();  // Clear Ctrl+Click selections
        }
        cx.notify();
    }

    /// Ctrl+Click to add/toggle cell in selection (discontiguous selection)
    pub fn ctrl_click_cell(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        // Save current selection to additional_selections
        self.additional_selections.push((self.selected, self.selection_end));
        // Start new selection at clicked cell
        self.selected = (row, col);
        self.selection_end = None;
        cx.notify();
    }

    /// Start drag selection - called on mouse_down
    pub fn start_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        self.selected = (row, col);
        self.selection_end = None;
        self.additional_selections.clear();  // Clear Ctrl+Click selections on new drag
        cx.notify();
    }

    /// Start drag selection with Ctrl held (add to existing selections)
    pub fn start_ctrl_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        // Save current selection to additional_selections
        self.additional_selections.push((self.selected, self.selection_end));
        // Start new selection at clicked cell
        self.selected = (row, col);
        self.selection_end = None;
        cx.notify();
    }

    /// Continue drag selection - called on mouse_move while dragging
    pub fn continue_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_selection {
            return;
        }
        // Only update if the cell changed to avoid unnecessary redraws
        if self.selection_end != Some((row, col)) {
            self.selection_end = Some((row, col));
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

    pub fn select_all(&mut self, cx: &mut Context<Self>) {
        self.selected = (0, 0);
        self.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        self.additional_selections.clear();  // Clear discontiguous selections
        cx.notify();
    }

    // Scrolling
    pub fn scroll(&mut self, delta_rows: i32, delta_cols: i32, cx: &mut Context<Self>) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();
        let new_row = (self.scroll_row as i32 + delta_rows)
            .max(0)
            .min((NUM_ROWS.saturating_sub(visible_rows)) as i32) as usize;
        let new_col = (self.scroll_col as i32 + delta_cols)
            .max(0)
            .min((NUM_COLS.saturating_sub(visible_cols)) as i32) as usize;

        if new_row != self.scroll_row || new_col != self.scroll_col {
            self.scroll_row = new_row;
            self.scroll_col = new_col;
            cx.notify();
        }
    }

    // Editing
    pub fn start_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;
        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = self.edit_original.clone();
        self.mode = Mode::Edit;
        cx.notify();
    }

    pub fn start_edit_clear(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;
        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = String::new();
        self.mode = Mode::Edit;
        cx.notify();
    }

    pub fn confirm_edit(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(1, 0, cx);  // Enter moves down
    }

    pub fn confirm_edit_and_move_right(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, 1, cx);  // Tab moves right
    }

    pub fn confirm_edit_and_move_left(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, -1, cx);  // Shift+Tab moves left
    }

    /// Ctrl+Enter: Confirm edit and apply to ALL selected cells (multi-edit)
    pub fn confirm_edit_in_place(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // If not editing, start editing
            self.start_edit(cx);
            return;
        }

        let new_value = self.edit_value.clone();
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        let mut changes = Vec::new();

        // Apply to all cells in selection
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(changes);
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.is_modified = true;

        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);
        if cell_count > 1 {
            self.status_message = Some(format!("Applied to {} cells", cell_count));
        }
        cx.notify();
    }

    fn confirm_edit_and_move(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            self.start_edit(cx);
            return;
        }

        let (row, col) = self.selected;
        let old_value = self.edit_original.clone();
        let new_value = self.edit_value.clone();

        self.history.record_change(row, col, old_value, new_value.clone());
        self.sheet_mut().set_value(row, col, &new_value);
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.is_modified = true;

        // Move after confirming
        self.move_selection(dr, dc, cx);
    }

    pub fn cancel_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.mode = Mode::Navigation;
            self.edit_value.clear();
            cx.notify();
        }
    }

    pub fn backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_value.pop();
            cx.notify();
        }
    }

    pub fn delete_char(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && !self.edit_value.is_empty() {
            self.edit_value.remove(0);
            cx.notify();
        }
    }

    pub fn insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_value.push(c);
            cx.notify();
        } else {
            // Start editing with this character
            let (row, col) = self.selected;
            self.edit_original = self.sheet().get_raw(row, col);
            self.edit_value = c.to_string();
            self.mode = Mode::Edit;
            cx.notify();
        }
    }

    // Clipboard
    pub fn copy(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Build tab-separated values for clipboard
        let mut text = String::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if col > min_col {
                    text.push('\t');
                }
                text.push_str(&self.sheet().get_display(row, col));
            }
            if row < max_row {
                text.push('\n');
            }
        }

        self.clipboard = Some(text.clone());
        cx.write_to_clipboard(ClipboardItem::new_string(text));
        self.status_message = Some("Copied to clipboard".to_string());
        cx.notify();
    }

    pub fn cut(&mut self, cx: &mut Context<Self>) {
        self.copy(cx);

        // Clear the selected cells and record history
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let mut changes = Vec::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().set_value(row, col, "");
            }
        }
        self.history.record_batch(changes);
        self.is_modified = true;
        self.status_message = Some("Cut to clipboard".to_string());
        cx.notify();
    }

    pub fn paste(&mut self, cx: &mut Context<Self>) {
        // If editing, paste into the edit buffer instead
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.clipboard.clone()
        };

        if let Some(text) = text {
            let (start_row, start_col) = self.selected;
            let mut changes = Vec::new();

            // Parse tab-separated values
            for (row_offset, line) in text.lines().enumerate() {
                for (col_offset, value) in line.split('\t').enumerate() {
                    let row = start_row + row_offset;
                    let col = start_col + col_offset;
                    if row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet().get_raw(row, col);
                        let new_value = value.to_string();
                        if old_value != new_value {
                            changes.push(CellChange {
                                row, col, old_value, new_value,
                            });
                        }
                        self.sheet_mut().set_value(row, col, value);
                    }
                }
            }

            self.history.record_batch(changes);
            self.is_modified = true;
            self.status_message = Some("Pasted from clipboard".to_string());
            cx.notify();
        }
    }

    /// Paste clipboard text into the edit buffer (when in editing mode)
    pub fn paste_into_edit(&mut self, cx: &mut Context<Self>) {
        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.clipboard.clone()
        };

        if let Some(text) = text {
            // Only take first line if multi-line, and trim whitespace
            let text = text.lines().next().unwrap_or("").trim();
            if !text.is_empty() {
                // Append to edit value (no cursor position tracking yet)
                self.edit_value.push_str(text);
                self.status_message = Some(format!("Pasted: {}", text));
                cx.notify();
            }
        }
    }

    pub fn delete_selection(&mut self, cx: &mut Context<Self>) {
        let mut changes = Vec::new();

        // Delete from all selection ranges (including discontiguous Ctrl+Click selections)
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            // Only get cells that actually have data (efficient for large selections)
            let cells_to_delete = self.sheet().cells_in_range(min_row, max_row, min_col, max_col);

            for (row, col) in cells_to_delete {
                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().clear_cell(row, col);
            }
        }

        if !changes.is_empty() {
            self.history.record_batch(changes);
            self.is_modified = true;
        }
        cx.notify();
    }

    // Undo/Redo
    pub fn undo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.undo() {
            for change in entry.changes {
                self.sheet_mut().set_value(change.row, change.col, &change.old_value);
            }
            self.is_modified = true;
            self.status_message = Some("Undo".to_string());
            cx.notify();
        }
    }

    pub fn redo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.redo() {
            for change in entry.changes {
                self.sheet_mut().set_value(change.row, change.col, &change.new_value);
            }
            self.is_modified = true;
            self.status_message = Some("Redo".to_string());
            cx.notify();
        }
    }

    // Selection helpers
    pub fn selection_range(&self) -> ((usize, usize), (usize, usize)) {
        let start = self.selected;
        let end = self.selection_end.unwrap_or(start);
        let min_row = start.0.min(end.0);
        let max_row = start.0.max(end.0);
        let min_col = start.1.min(end.1);
        let max_col = start.1.max(end.1);
        ((min_row, min_col), (max_row, max_col))
    }

    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        // Check active selection
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            return true;
        }
        // Check additional selections (Ctrl+Click ranges)
        for (start, end) in &self.additional_selections {
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
        for (start, end) in &self.additional_selections {
            let end = end.unwrap_or(*start);
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            ranges.push(((min_row, min_col), (max_row, max_col)));
        }
        ranges
    }

    // Formatting (applies to all discontiguous selection ranges)
    pub fn toggle_bold(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_bold(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_italic(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_italic(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_underline(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_underline(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    // Go To cell dialog
    pub fn show_goto(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::GoTo;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn hide_goto(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn confirm_goto(&mut self, cx: &mut Context<Self>) {
        if let Some((row, col)) = Self::parse_cell_ref(&self.goto_input) {
            if row < NUM_ROWS && col < NUM_COLS {
                self.selected = (row, col);
                self.selection_end = None;
                self.ensure_visible(cx);
                self.status_message = Some(format!("Jumped to {}", self.cell_ref()));
            } else {
                self.status_message = Some("Cell reference out of range".to_string());
            }
        } else {
            self.status_message = Some("Invalid cell reference".to_string());
        }
        self.mode = Mode::Navigation;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn goto_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == Mode::GoTo {
            self.goto_input.push(c.to_ascii_uppercase());
            cx.notify();
        }
    }

    pub fn goto_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::GoTo {
            self.goto_input.pop();
            cx.notify();
        }
    }

    /// Parse cell reference like "A1", "B25", "AA100"
    fn parse_cell_ref(input: &str) -> Option<(usize, usize)> {
        let input = input.trim().to_uppercase();
        if input.is_empty() {
            return None;
        }

        // Find where letters end and numbers begin
        let letter_end = input.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 || letter_end == input.len() {
            return None;
        }

        let letters = &input[..letter_end];
        let numbers = &input[letter_end..];

        // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
        let col = letters.chars().fold(0usize, |acc, c| {
            acc * 26 + (c as usize - 'A' as usize + 1)
        }) - 1;

        // Parse row (1-based to 0-based)
        let row = numbers.parse::<usize>().ok()?.checked_sub(1)?;

        Some((row, col))
    }

    // Find in cells
    pub fn show_find(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Find;
        self.find_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        cx.notify();
    }

    pub fn hide_find(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn find_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == Mode::Find {
            self.find_input.push(c);
            self.perform_find(cx);
        }
    }

    pub fn find_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::Find {
            self.find_input.pop();
            self.perform_find(cx);
        }
    }

    fn perform_find(&mut self, cx: &mut Context<Self>) {
        self.find_results.clear();
        self.find_index = 0;

        if self.find_input.is_empty() {
            self.status_message = None;
            cx.notify();
            return;
        }

        let query = self.find_input.to_lowercase();

        // Search through all populated cells
        let cell_positions: Vec<_> = self.sheet().cells_iter()
            .map(|(&pos, _)| pos)
            .collect();

        for (row, col) in cell_positions {
            let display = self.sheet().get_display(row, col);
            if display.to_lowercase().contains(&query) {
                self.find_results.push((row, col));
            }
        }

        // Sort results by row, then column
        self.find_results.sort();

        if !self.find_results.is_empty() {
            self.jump_to_find_result(cx);
            self.status_message = Some(format!(
                "Found {} match{}",
                self.find_results.len(),
                if self.find_results.len() == 1 { "" } else { "es" }
            ));
        } else {
            self.status_message = Some("No matches found".to_string());
        }
        cx.notify();
    }

    pub fn find_next(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        self.find_index = (self.find_index + 1) % self.find_results.len();
        self.jump_to_find_result(cx);
    }

    pub fn find_prev(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        if self.find_index == 0 {
            self.find_index = self.find_results.len() - 1;
        } else {
            self.find_index -= 1;
        }
        self.jump_to_find_result(cx);
    }

    fn jump_to_find_result(&mut self, cx: &mut Context<Self>) {
        if let Some(&(row, col)) = self.find_results.get(self.find_index) {
            self.selected = (row, col);
            self.selection_end = None;
            self.ensure_visible(cx);
            self.status_message = Some(format!(
                "Match {} of {}",
                self.find_index + 1,
                self.find_results.len()
            ));
        }
    }

    // Command Palette
    pub fn toggle_palette(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::Command {
            self.hide_palette(cx);
        } else {
            self.show_palette(cx);
        }
    }

    pub fn show_palette(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Command;
        self.palette_query.clear();
        self.palette_selected = 0;
        cx.notify();
    }

    pub fn hide_palette(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.palette_query.clear();
        self.palette_selected = 0;
        cx.notify();
    }

    pub fn palette_up(&mut self, cx: &mut Context<Self>) {
        if self.palette_selected > 0 {
            self.palette_selected -= 1;
            cx.notify();
        }
    }

    pub fn palette_down(&mut self, cx: &mut Context<Self>) {
        use crate::views::command_palette::filter_commands;
        let count = filter_commands(&self.palette_query).len();
        if self.palette_selected + 1 < count {
            self.palette_selected += 1;
            cx.notify();
        }
    }

    pub fn palette_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.palette_query.push(c);
        self.palette_selected = 0;  // Reset selection on filter change
        cx.notify();
    }

    pub fn palette_backspace(&mut self, cx: &mut Context<Self>) {
        self.palette_query.pop();
        self.palette_selected = 0;  // Reset selection on filter change
        cx.notify();
    }

    pub fn palette_execute(&mut self, cx: &mut Context<Self>) {
        use crate::views::command_palette::filter_commands;
        let filtered = filter_commands(&self.palette_query);
        if let Some(cmd) = filtered.get(self.palette_selected) {
            let action = cmd.action;
            self.hide_palette(cx);
            action(self, cx);
        } else {
            self.hide_palette(cx);
        }
    }

    // Font Picker
    pub fn show_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::FontPicker;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn hide_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_up(&mut self, cx: &mut Context<Self>) {
        if self.font_picker_selected > 0 {
            self.font_picker_selected -= 1;
            cx.notify();
        }
    }

    pub fn font_picker_down(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if self.font_picker_selected + 1 < filtered.len() {
            self.font_picker_selected += 1;
            cx.notify();
        }
    }

    pub fn font_picker_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.font_picker_query.push(c);
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_backspace(&mut self, cx: &mut Context<Self>) {
        self.font_picker_query.pop();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_execute(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if let Some(font_name) = filtered.get(self.font_picker_selected) {
            let font = font_name.clone();
            self.apply_font_to_selection(&font, cx);
        }
        self.hide_font_picker(cx);
    }

    /// Filter available fonts by query
    pub fn filter_fonts(&self) -> Vec<String> {
        if self.font_picker_query.is_empty() {
            return self.available_fonts.clone();
        }
        let query_lower = self.font_picker_query.to_lowercase();
        self.available_fonts
            .iter()
            .filter(|f| f.to_lowercase().contains(&query_lower))
            .cloned()
            .collect()
    }

    /// Apply font to all cells in current selection
    pub fn apply_font_to_selection(&mut self, font_name: &str, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let font = if font_name.is_empty() { None } else { Some(font_name.to_string()) };

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                self.sheet_mut().set_font_family(row, col, font.clone());
            }
        }

        self.is_modified = true;
        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);
        self.status_message = Some(format!("Applied font '{}' to {} cell(s)", font_name, cell_count));
        cx.notify();
    }

    /// Clear font from selection (reset to default)
    pub fn clear_font_from_selection(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                self.sheet_mut().set_font_family(row, col, None);
            }
        }

        self.is_modified = true;
        self.status_message = Some("Cleared font from selection".to_string());
        cx.notify();
    }

    // Fill operations (Phase 0 essentials)
    pub fn fill_down(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Need at least 2 rows selected
        if max_row <= min_row {
            self.status_message = Some("Select at least 2 rows to fill down".into());
            cx.notify();
            return;
        }

        let mut changes = Vec::new();

        // For each column in selection
        for col in min_col..=max_col {
            // Get the source value/formula from the first row
            let source = self.sheet().get_raw(min_row, col);

            // Fill down to all other rows
            for row in (min_row + 1)..=max_row {
                let old_value = self.sheet().get_raw(row, col);
                let new_value = if source.starts_with('=') {
                    // Adjust relative references for formulas
                    self.adjust_formula_refs(&source, row as i32 - min_row as i32, 0)
                } else {
                    source.clone()
                };

                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(changes);
        self.is_modified = true;

        self.status_message = Some("Filled down".into());
        cx.notify();
    }

    pub fn fill_right(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Need at least 2 columns selected
        if max_col <= min_col {
            self.status_message = Some("Select at least 2 columns to fill right".into());
            cx.notify();
            return;
        }

        let mut changes = Vec::new();

        // For each row in selection
        for row in min_row..=max_row {
            // Get the source value/formula from the first column
            let source = self.sheet().get_raw(row, min_col);

            // Fill right to all other columns
            for col in (min_col + 1)..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                let new_value = if source.starts_with('=') {
                    // Adjust relative references for formulas
                    self.adjust_formula_refs(&source, 0, col as i32 - min_col as i32)
                } else {
                    source.clone()
                };

                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(changes);
        self.is_modified = true;
        self.status_message = Some("Filled right".into());
        cx.notify();
    }

    /// Adjust cell references in a formula by delta rows and cols
    /// Handles relative (A1), absolute ($A$1), and mixed ($A1, A$1) references
    fn adjust_formula_refs(&self, formula: &str, delta_row: i32, delta_col: i32) -> String {
        use regex::Regex;

        // Match cell references: optional $ before col, col letters, optional $ before row, row numbers
        let re = Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

        re.replace_all(formula, |caps: &regex::Captures| {
            let col_absolute = &caps[1] == "$";
            let col_letters = &caps[2];
            let row_absolute = &caps[3] == "$";
            let row_num: i32 = caps[4].parse().unwrap_or(1);

            // Parse column
            let col = col_letters.to_uppercase().chars().fold(0i32, |acc, c| {
                acc * 26 + (c as i32 - 'A' as i32 + 1)
            }) - 1;

            // Apply deltas if not absolute
            let new_col = if col_absolute { col } else { col + delta_col };
            let new_row = if row_absolute { row_num } else { row_num + delta_row };

            // Bounds check
            if new_col < 0 || new_row < 1 {
                return format!("#REF!");
            }

            // Convert column back to letters
            let col_str = Self::col_letter(new_col as usize);

            format!(
                "{}{}{}{}",
                if col_absolute { "$" } else { "" },
                col_str,
                if row_absolute { "$" } else { "" },
                new_row
            )
        })
        .to_string()
    }
}

impl Render for Spreadsheet {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        // Update window size if changed (handles resize)
        let current_size = window.viewport_size();
        if self.window_size != current_size {
            self.window_size = current_size;
        }
        views::render_spreadsheet(self, cx)
    }
}

#[cfg(test)]
mod tests {
    use regex::Regex;
    use visigrid_engine::sheet::Sheet;

    /// Test-only version of adjust_formula_refs (mirrors Spreadsheet::adjust_formula_refs)
    fn adjust_formula_refs(formula: &str, delta_row: i32, delta_col: i32) -> String {
        let re = Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

        re.replace_all(formula, |caps: &regex::Captures| {
            let col_absolute = &caps[1] == "$";
            let col_letters = &caps[2];
            let row_absolute = &caps[3] == "$";
            let row_num: i32 = caps[4].parse().unwrap_or(1);

            let col = col_letters.to_uppercase().chars().fold(0i32, |acc, c| {
                acc * 26 + (c as i32 - 'A' as i32 + 1)
            }) - 1;

            let new_col = if col_absolute { col } else { col + delta_col };
            let new_row = if row_absolute { row_num } else { row_num + delta_row };

            if new_col < 0 || new_row < 1 {
                return "#REF!".to_string();
            }

            let col_str = col_to_letter(new_col as usize);

            format!(
                "{}{}{}{}",
                if col_absolute { "$" } else { "" },
                col_str,
                if row_absolute { "$" } else { "" },
                new_row
            )
        }).to_string()
    }

    fn col_to_letter(col: usize) -> String {
        let mut s = String::new();
        let mut n = col;
        loop {
            s.insert(0, (b'A' + (n % 26) as u8) as char);
            if n < 26 { break; }
            n = n / 26 - 1;
        }
        s
    }

    /// Simulate fill_down at the Sheet level (no gpui required)
    fn fill_down_on_sheet(sheet: &mut Sheet, min_row: usize, max_row: usize, col: usize) {
        let source = sheet.get_raw(min_row, col);
        for row in (min_row + 1)..=max_row {
            let new_value = if source.starts_with('=') {
                adjust_formula_refs(&source, row as i32 - min_row as i32, 0)
            } else {
                source.clone()
            };
            sheet.set_value(row, col, &new_value);
        }
    }

    // =========================================================================
    // REGRESSION TEST: Mixed references (the bug we just fixed)
    // =========================================================================

    #[test]
    fn test_fill_down_mixed_references_formulas() {
        // Test that adjust_formula_refs correctly handles all 4 reference types
        let formula = "=A1 + $A$1 + A$1 + $A1";

        // Fill down by 1 row
        assert_eq!(
            adjust_formula_refs(formula, 1, 0),
            "=A2 + $A$1 + A$1 + $A2",
            "Row 2: A1->A2 (relative), $A$1->$A$1 (absolute), A$1->A$1 (row absolute), $A1->$A2 (col absolute)"
        );

        // Fill down by 2 rows
        assert_eq!(
            adjust_formula_refs(formula, 2, 0),
            "=A3 + $A$1 + A$1 + $A3"
        );

        // Fill down by 3 rows
        assert_eq!(
            adjust_formula_refs(formula, 3, 0),
            "=A4 + $A$1 + A$1 + $A4"
        );
    }

    #[test]
    fn test_fill_down_mixed_references_end_to_end() {
        // End-to-end test: seed values, fill, verify formulas AND computed values
        let mut sheet = Sheet::new(100, 100);

        // Seed A1:A4 with distinct values
        sheet.set_value(0, 0, "10"); // A1 = 10
        sheet.set_value(1, 0, "1");  // A2 = 1
        sheet.set_value(2, 0, "2");  // A3 = 2
        sheet.set_value(3, 0, "3");  // A4 = 3

        // Set B1 formula: =A1 + $A$1 + A$1 + $A1
        sheet.set_value(0, 1, "=A1 + $A$1 + A$1 + $A1");

        // Verify B1 value before fill
        assert_eq!(sheet.get_display(0, 1), "40", "B1 should be 10+10+10+10=40");

        // Simulate fill_down from B1 to B4
        fill_down_on_sheet(&mut sheet, 0, 3, 1);

        // Assert formulas are correct
        assert_eq!(sheet.get_raw(0, 1), "=A1 + $A$1 + A$1 + $A1", "B1 formula unchanged");
        assert_eq!(sheet.get_raw(1, 1), "=A2 + $A$1 + A$1 + $A2", "B2 formula adjusted");
        assert_eq!(sheet.get_raw(2, 1), "=A3 + $A$1 + A$1 + $A3", "B3 formula adjusted");
        assert_eq!(sheet.get_raw(3, 1), "=A4 + $A$1 + A$1 + $A4", "B4 formula adjusted");

        // Assert computed values are correct
        // B1: A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
        // B2: A2(1) + $A$1(10) + A$1(10) + $A2(1) = 22
        // B3: A3(2) + $A$1(10) + A$1(10) + $A3(2) = 24
        // B4: A4(3) + $A$1(10) + A$1(10) + $A4(3) = 26
        assert_eq!(sheet.get_display(0, 1), "40", "B1 value");
        assert_eq!(sheet.get_display(1, 1), "22", "B2 value: 1+10+10+1");
        assert_eq!(sheet.get_display(2, 1), "24", "B3 value: 2+10+10+2");
        assert_eq!(sheet.get_display(3, 1), "26", "B4 value: 3+10+10+3");
    }

    // =========================================================================
    // EDGE CASE: Ranges in formulas (SUM, etc.)
    // =========================================================================

    #[test]
    fn test_fill_down_with_ranges_formulas() {
        // =SUM(A1:A3) + $A$1 should become =SUM(A2:A4) + $A$1
        let formula = "=SUM(A1:A3) + $A$1";

        assert_eq!(
            adjust_formula_refs(formula, 1, 0),
            "=SUM(A2:A4) + $A$1",
            "Range A1:A3 should become A2:A4, absolute $A$1 stays"
        );

        assert_eq!(
            adjust_formula_refs(formula, 2, 0),
            "=SUM(A3:A5) + $A$1"
        );
    }

    #[test]
    fn test_fill_down_with_ranges_end_to_end() {
        let mut sheet = Sheet::new(100, 100);

        // Seed values
        sheet.set_value(0, 0, "10"); // A1 = 10
        sheet.set_value(1, 0, "20"); // A2 = 20
        sheet.set_value(2, 0, "30"); // A3 = 30
        sheet.set_value(3, 0, "40"); // A4 = 40
        sheet.set_value(4, 0, "50"); // A5 = 50

        // B1 = SUM(A1:A3) + $A$1 = (10+20+30) + 10 = 70
        sheet.set_value(0, 1, "=SUM(A1:A3) + $A$1");
        assert_eq!(sheet.get_display(0, 1), "70", "B1: SUM(10,20,30)+10");

        // Fill down B1:B3
        fill_down_on_sheet(&mut sheet, 0, 2, 1);

        // Check formulas
        assert_eq!(sheet.get_raw(0, 1), "=SUM(A1:A3) + $A$1");
        assert_eq!(sheet.get_raw(1, 1), "=SUM(A2:A4) + $A$1");
        assert_eq!(sheet.get_raw(2, 1), "=SUM(A3:A5) + $A$1");

        // Check values
        // B1: SUM(A1:A3) + $A$1 = (10+20+30) + 10 = 70
        // B2: SUM(A2:A4) + $A$1 = (20+30+40) + 10 = 100
        // B3: SUM(A3:A5) + $A$1 = (30+40+50) + 10 = 130
        assert_eq!(sheet.get_display(0, 1), "70", "B1 value");
        assert_eq!(sheet.get_display(1, 1), "100", "B2 value: SUM(20,30,40)+10");
        assert_eq!(sheet.get_display(2, 1), "130", "B3 value: SUM(30,40,50)+10");
    }

    // =========================================================================
    // EDGE CASE: Multi-letter columns (AA, AB, etc.)
    // =========================================================================

    #[test]
    fn test_fill_down_multi_letter_columns_formulas() {
        // =AA1 + $B$1 + C$2 + $D3
        // AA1 -> AA2 (both relative)
        // $B$1 -> $B$1 (both absolute)
        // C$2 -> C$2 (row absolute)
        // $D3 -> $D4 (col absolute, row relative)
        let formula = "=AA1 + $B$1 + C$2 + $D3";

        assert_eq!(
            adjust_formula_refs(formula, 1, 0),
            "=AA2 + $B$1 + C$2 + $D4",
            "Multi-letter columns with mixed refs"
        );

        assert_eq!(
            adjust_formula_refs(formula, 2, 0),
            "=AA3 + $B$1 + C$2 + $D5"
        );
    }

    #[test]
    fn test_fill_down_multi_letter_columns_end_to_end() {
        let mut sheet = Sheet::new(100, 100);

        // AA is column 26 (0-indexed), B is 1, C is 2, D is 3
        // Seed values
        sheet.set_value(0, 26, "100"); // AA1 = 100
        sheet.set_value(1, 26, "200"); // AA2 = 200
        sheet.set_value(2, 26, "300"); // AA3 = 300

        sheet.set_value(0, 1, "10");   // B1 = 10

        sheet.set_value(1, 2, "5");    // C2 = 5

        sheet.set_value(2, 3, "1");    // D3 = 1
        sheet.set_value(3, 3, "2");    // D4 = 2
        sheet.set_value(4, 3, "3");    // D5 = 3

        // AB1 = AA1 + $B$1 + C$2 + $D3 = 100 + 10 + 5 + 1 = 116
        sheet.set_value(0, 27, "=AA1 + $B$1 + C$2 + $D3"); // AB1
        assert_eq!(sheet.get_display(0, 27), "116", "AB1: 100+10+5+1");

        // Fill down AB1:AB3
        fill_down_on_sheet(&mut sheet, 0, 2, 27);

        // Check formulas
        assert_eq!(sheet.get_raw(0, 27), "=AA1 + $B$1 + C$2 + $D3");
        assert_eq!(sheet.get_raw(1, 27), "=AA2 + $B$1 + C$2 + $D4");
        assert_eq!(sheet.get_raw(2, 27), "=AA3 + $B$1 + C$2 + $D5");

        // Check values
        // AB1: AA1(100) + $B$1(10) + C$2(5) + $D3(1) = 116
        // AB2: AA2(200) + $B$1(10) + C$2(5) + $D4(2) = 217
        // AB3: AA3(300) + $B$1(10) + C$2(5) + $D5(3) = 318
        assert_eq!(sheet.get_display(0, 27), "116", "AB1 value");
        assert_eq!(sheet.get_display(1, 27), "217", "AB2 value: 200+10+5+2");
        assert_eq!(sheet.get_display(2, 27), "318", "AB3 value: 300+10+5+3");
    }

    // =========================================================================
    // EDGE CASE: Fill right (column adjustment)
    // =========================================================================

    #[test]
    fn test_fill_right_formulas() {
        let formula = "=A1 + $A$1 + A$1 + $A1";

        // Fill right by 1 column
        // A1 -> B1 (col relative)
        // $A$1 -> $A$1 (both absolute)
        // A$1 -> B$1 (col relative, row absolute)
        // $A1 -> $A1 (col absolute)
        assert_eq!(
            adjust_formula_refs(formula, 0, 1),
            "=B1 + $A$1 + B$1 + $A1",
            "Fill right: relative cols shift, absolute cols stay"
        );
    }

    /// Simulate fill_right at the Sheet level (no gpui required)
    fn fill_right_on_sheet(sheet: &mut Sheet, row: usize, min_col: usize, max_col: usize) {
        let source = sheet.get_raw(row, min_col);
        for col in (min_col + 1)..=max_col {
            let new_value = if source.starts_with('=') {
                adjust_formula_refs(&source, 0, col as i32 - min_col as i32)
            } else {
                source.clone()
            };
            sheet.set_value(row, col, &new_value);
        }
    }

    #[test]
    fn test_fill_right_mixed_references_end_to_end() {
        // End-to-end test for fill right with mixed references
        let mut sheet = Sheet::new(100, 100);

        // Seed row 1 with distinct values: A1=10, B1=1, C1=2, D1=3
        sheet.set_value(0, 0, "10"); // A1 = 10
        sheet.set_value(0, 1, "1");  // B1 = 1
        sheet.set_value(0, 2, "2");  // C1 = 2
        sheet.set_value(0, 3, "3");  // D1 = 3

        // Set A2 formula: =A1 + $A$1 + A$1 + $A1
        // When filling right:
        // - A1 shifts column (relative col)
        // - $A$1 stays (both absolute)
        // - A$1 shifts column (relative col, absolute row)
        // - $A1 stays (absolute col, relative row)
        sheet.set_value(1, 0, "=A1 + $A$1 + A$1 + $A1");

        // Verify A2 value before fill
        // A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
        assert_eq!(sheet.get_display(1, 0), "40", "A2 should be 40");

        // Fill right A2:D2
        fill_right_on_sheet(&mut sheet, 1, 0, 3);

        // Check formulas
        assert_eq!(sheet.get_raw(1, 0), "=A1 + $A$1 + A$1 + $A1", "A2 formula unchanged");
        assert_eq!(sheet.get_raw(1, 1), "=B1 + $A$1 + B$1 + $A1", "B2 formula adjusted");
        assert_eq!(sheet.get_raw(1, 2), "=C1 + $A$1 + C$1 + $A1", "C2 formula adjusted");
        assert_eq!(sheet.get_raw(1, 3), "=D1 + $A$1 + D$1 + $A1", "D2 formula adjusted");

        // Check computed values
        // A2: A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
        // B2: B1(1) + $A$1(10) + B$1(1) + $A1(10) = 22
        // C2: C1(2) + $A$1(10) + C$1(2) + $A1(10) = 24
        // D2: D1(3) + $A$1(10) + D$1(3) + $A1(10) = 26
        assert_eq!(sheet.get_display(1, 0), "40", "A2 value");
        assert_eq!(sheet.get_display(1, 1), "22", "B2 value: 1+10+1+10");
        assert_eq!(sheet.get_display(1, 2), "24", "C2 value: 2+10+2+10");
        assert_eq!(sheet.get_display(1, 3), "26", "D2 value: 3+10+3+10");
    }

    // =========================================================================
    // EDGE CASE: Multi-edit with single undo
    // =========================================================================

    #[test]
    fn test_multi_edit_applies_once_and_single_undo() {
        use crate::history::{History, CellChange};

        let mut sheet = Sheet::new(100, 100);
        let mut history = History::new();

        // Seed initial values: A1=1, A2=2, A3=3, B1=10, B2=20, B3=30
        sheet.set_value(0, 0, "1");  // A1
        sheet.set_value(1, 0, "2");  // A2
        sheet.set_value(2, 0, "3");  // A3
        sheet.set_value(0, 1, "10"); // B1
        sheet.set_value(1, 1, "20"); // B2
        sheet.set_value(2, 1, "30"); // B3

        // Simulate multi-edit: set "=A1*2" to selection A1:B3 (6 cells)
        let new_value = "=A1*2";
        let selection = [(0, 0), (0, 1), (1, 0), (1, 1), (2, 0), (2, 1)];

        let mut changes = Vec::new();
        for (row, col) in selection.iter() {
            let old_value = sheet.get_raw(*row, *col);
            if old_value != new_value {
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: new_value.to_string(),
                });
            }
            sheet.set_value(*row, *col, new_value);
        }

        // Record as single batch (this is what multi-edit does)
        history.record_batch(changes);

        // Verify all 6 cells have the formula
        for (row, col) in selection.iter() {
            assert_eq!(
                sheet.get_raw(*row, *col), "=A1*2",
                "Cell ({}, {}) should have formula =A1*2", row, col
            );
        }

        // Verify computed values (all reference A1 which is now =A1*2, causing circular ref)
        // Actually A1 = =A1*2 is circular, so let's verify at least B1 computes
        // B1 = =A1*2 where A1 = =A1*2 (circular)
        // The key test is that single undo reverts ALL cells

        // Single undo should revert ALL 6 cells
        let entry = history.undo().expect("Should have undo entry");
        assert_eq!(entry.changes.len(), 6, "Undo entry should contain all 6 changes");

        // Apply undo to sheet
        for change in entry.changes.iter() {
            sheet.set_value(change.row, change.col, &change.old_value);
        }

        // Verify original values are restored
        assert_eq!(sheet.get_raw(0, 0), "1", "A1 restored to 1");
        assert_eq!(sheet.get_raw(1, 0), "2", "A2 restored to 2");
        assert_eq!(sheet.get_raw(2, 0), "3", "A3 restored to 3");
        assert_eq!(sheet.get_raw(0, 1), "10", "B1 restored to 10");
        assert_eq!(sheet.get_raw(1, 1), "20", "B2 restored to 20");
        assert_eq!(sheet.get_raw(2, 1), "30", "B3 restored to 30");

        // Verify redo works and contains all 6 changes
        let redo_entry = history.redo().expect("Should have redo entry");
        assert_eq!(redo_entry.changes.len(), 6, "Redo entry should contain all 6 changes");
    }

    // =========================================================================
    // EDGE CASE: Boundary conditions
    // =========================================================================

    #[test]
    fn test_fill_down_ref_error() {
        // Filling up from row 1 should produce #REF!
        let formula = "=A1";
        assert_eq!(
            adjust_formula_refs(formula, -1, 0),
            "=#REF!",
            "Row 0 (A0) doesn't exist, should be #REF!"
        );
    }

    #[test]
    fn test_fill_left_ref_error() {
        // Filling left from column A should produce #REF!
        let formula = "=A1";
        assert_eq!(
            adjust_formula_refs(formula, 0, -1),
            "=#REF!",
            "Column before A doesn't exist, should be #REF!"
        );
    }
}
