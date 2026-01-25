//! Keyboard hints integration
//!
//! Contains Spreadsheet methods for Vimium-style jump navigation:
//! - Hint mode entry/exit
//! - Key handling in hint mode
//! - Vim-style navigation keys
//!
//! Core hint generation and resolution logic is in hints.rs.

use gpui::*;
use crate::app::{Spreadsheet, NUM_COLS};
use crate::mode::Mode;

impl Spreadsheet {
    // ========================================================================
    // Keyboard hints (Vimium-style jump navigation)
    // ========================================================================

    /// Enter keyboard hint mode - show jump labels on visible cells.
    pub fn enter_hint_mode(&mut self, cx: &mut Context<Self>) {
        self.enter_hint_mode_with_labels(true, cx);
    }

    /// Enter hint/command mode with optional labels.
    ///
    /// - `show_labels`: if true, generate cell labels (full hint mode)
    /// - `show_labels`: if false, command-only mode (for vim gg without labels)
    pub fn enter_hint_mode_with_labels(&mut self, show_labels: bool, cx: &mut Context<Self>) {
        self.hint_state.buffer.clear();

        if show_labels {
            // Full hint mode: generate labels for visible cells
            let visible_rows = self.visible_rows();
            let visible_cols = self.visible_cols();
            self.hint_state.labels = crate::hints::generate_hints(
                self.view_state.scroll_row,
                self.view_state.scroll_col,
                visible_rows,
                visible_cols,
            );
            self.hint_state.viewport = (self.view_state.scroll_row, self.view_state.scroll_col, visible_rows, visible_cols);
            self.status_message = Some("Hint: type letters to jump".into());
        } else {
            // Command-only mode: no labels, just waiting for g-commands (gg, etc.)
            self.hint_state.labels.clear();
            self.hint_state.viewport = (0, 0, 0, 0);
            self.status_message = Some("g-".into());
        }

        self.mode = Mode::Hint;
        cx.notify();
    }

    /// Exit keyboard hint mode without jumping.
    pub fn exit_hint_mode(&mut self, cx: &mut Context<Self>) {
        self.hint_state.clear();
        self.mode = Mode::Navigation;
        self.status_message = None;
        cx.notify();
    }

    /// Handle a key press in hint mode.
    /// Returns true if the key was consumed.
    ///
    /// Uses the resolver architecture from hints.rs:
    /// 1. Exact command match (gg → GotoTop)
    /// 2. Cell label resolution (a, ab, zz)
    /// 3. No match → exit
    pub fn apply_hint_key(&mut self, key: &str, cx: &mut Context<Self>) -> bool {
        use crate::hints::{resolve_hint_buffer, HintResolution, HintExitReason};

        match key {
            "escape" => {
                self.hint_state.last_exit_reason = Some(HintExitReason::Cancelled);
                self.exit_hint_mode(cx);
                true
            }
            "backspace" => {
                self.hint_state.buffer.pop();
                self.update_hint_status(cx);
                true
            }
            _ if key.len() == 1 && key.chars().next().map(|c| c.is_ascii_lowercase()).unwrap_or(false) => {
                self.hint_state.buffer.push_str(key);

                // Resolve the buffer through the phase system
                match resolve_hint_buffer(&self.hint_state) {
                    HintResolution::Command(cmd) => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::Command);
                        self.execute_g_command(cmd, cx);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::Jump(row, col) => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::LabelJump);
                        self.view_state.selected = (row, col);
                        self.view_state.selection_end = None;
                        self.view_state.additional_selections.clear();
                        self.ensure_cell_visible(row, col);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::NoMatch => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::NoMatch);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::Pending => {
                        self.update_hint_status(cx);
                    }
                }
                true
            }
            _ => false, // Unhandled key
        }
    }

    /// Execute a g-prefixed command.
    fn execute_g_command(&mut self, cmd: crate::hints::GCommand, cx: &mut Context<Self>) {
        use crate::hints::GCommand;

        match cmd {
            GCommand::GotoTop => {
                // gg - Go to A1
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.additional_selections.clear();
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                self.status_message = Some("Jumped to A1".into());
                cx.notify();
            }
            // Future commands go here
        }
    }

    /// Update status bar with current hint state.
    fn update_hint_status(&mut self, cx: &mut Context<Self>) {
        let matches = self.hint_state.matching_labels();
        let buffer = &self.hint_state.buffer;

        if buffer.is_empty() {
            self.status_message = Some("Hint: type letters to jump".into());
        } else if matches.is_empty() {
            self.status_message = Some(format!("Hint: {} (no matches)", buffer));
        } else if matches.len() == 1 {
            // This shouldn't happen (we auto-jump on unique match), but handle it
            self.status_message = Some(format!("Hint: {} → jumping", buffer));
        } else {
            self.status_message = Some(format!("Hint: {} ({} matches)", buffer, matches.len()));
        }
        cx.notify();
    }

    /// Check if hints are enabled in settings.
    pub fn keyboard_hints_enabled(&self, cx: &Context<Self>) -> bool {
        use crate::settings::user_settings;
        user_settings(cx)
            .navigation
            .keyboard_hints
            .as_value()
            .copied()
            .unwrap_or(false)
    }

    /// Check if vim mode is enabled in settings.
    pub fn vim_mode_enabled(&self, cx: &Context<Self>) -> bool {
        use crate::settings::user_settings;
        user_settings(cx)
            .navigation
            .vim_mode
            .as_value()
            .copied()
            .unwrap_or(false)
    }

    /// Handle vim-style navigation keys.
    /// Returns true if the key was consumed.
    pub fn apply_vim_key(&mut self, key: &str, cx: &mut Context<Self>) -> bool {
        match key {
            "h" => {
                self.move_selection(0, -1, cx);
                true
            }
            "j" => {
                self.move_selection(1, 0, cx);
                true
            }
            "k" => {
                self.move_selection(-1, 0, cx);
                true
            }
            "l" => {
                self.move_selection(0, 1, cx);
                true
            }
            "i" => {
                // Enter edit mode (like F2 - edit without replacing)
                self.start_edit(cx);
                true
            }
            "0" => {
                // Move to first column
                self.view_state.selected = (self.view_state.selected.0, 0);
                self.view_state.selection_end = None;
                self.ensure_cell_visible(self.view_state.selected.0, 0);
                cx.notify();
                true
            }
            "$" => {
                // Move to last column with data in current row (or last visible)
                let row = self.view_state.selected.0;
                let last_col = self.find_last_data_col_in_row(row);
                self.view_state.selected = (row, last_col);
                self.view_state.selection_end = None;
                self.ensure_cell_visible(row, last_col);
                cx.notify();
                true
            }
            "w" => {
                // Forward jump: Ctrl+Right equivalent
                self.jump_selection(0, 1, cx);
                true
            }
            "b" => {
                // Back jump: Ctrl+Left equivalent
                self.jump_selection(0, -1, cx);
                true
            }
            _ => false,
        }
    }

    /// Find the last column with data in a given row.
    fn find_last_data_col_in_row(&self, row: usize) -> usize {
        let sheet = self.workbook.active_sheet();
        for col in (0..NUM_COLS).rev() {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                return col;
            }
        }
        0 // Default to first column if row is empty
    }
}
