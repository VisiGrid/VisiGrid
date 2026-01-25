//! Modal dialogs and panels
//!
//! Contains show/hide and input handling for:
//! - Go To cell dialog (Ctrl+G)
//! - Preferences panel
//! - Theme picker
//! - About dialog
//! - License dialog
//! - Import/Export report dialogs
//! - Inspector panel

use gpui::*;
use crate::app::Spreadsheet;
use crate::mode::Mode;
use crate::settings::{update_user_settings, Setting};
use crate::theme::{Theme, builtin_themes};

/// Maximum rows in the spreadsheet
const NUM_ROWS: usize = 1_000_000;
/// Maximum columns in the spreadsheet
const NUM_COLS: usize = 16_384;

impl Spreadsheet {
    // =========================================================================
    // Go To cell dialog
    // =========================================================================

    pub fn show_goto(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
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
                self.view_state.selected = (row, col);
                self.view_state.selection_end = None;
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

    // =========================================================================
    // Preferences panel
    // =========================================================================

    pub fn show_preferences(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::Preferences;
        cx.notify();
    }

    pub fn hide_preferences(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn theme_picker_up(&mut self, cx: &mut Context<Self>) {
        if self.theme_picker_selected > 0 {
            self.theme_picker_selected -= 1;
            self.update_theme_preview(cx);
        }
    }

    pub fn theme_picker_down(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if self.theme_picker_selected + 1 < filtered.len() {
            self.theme_picker_selected += 1;
            self.update_theme_preview(cx);
        }
    }

    pub fn theme_picker_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.theme_picker_query.push(c);
        self.theme_picker_selected = 0;
        self.update_theme_preview(cx);
    }

    pub fn theme_picker_backspace(&mut self, cx: &mut Context<Self>) {
        self.theme_picker_query.pop();
        self.theme_picker_selected = 0;
        self.update_theme_preview(cx);
    }

    pub fn theme_picker_execute(&mut self, cx: &mut Context<Self>) {
        self.apply_theme_at_index(self.theme_picker_selected, cx);
    }

    pub fn apply_theme_at_index(&mut self, index: usize, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if let Some(theme) = filtered.get(index) {
            self.theme = theme.clone();
            self.status_message = Some(format!("Applied theme: {}", theme.meta.name));
            // Persist theme selection to global store
            let theme_id = theme.meta.id.to_string();
            update_user_settings(cx, |settings| {
                settings.appearance.theme_id = Setting::Value(theme_id);
            });
        }
        self.theme_preview = None;
        self.mode = Mode::Navigation;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        cx.notify();
    }

    /// Filter available themes by query
    pub fn filter_themes(&self) -> Vec<Theme> {
        let themes = builtin_themes();
        if self.theme_picker_query.is_empty() {
            return themes;
        }
        let query_lower = self.theme_picker_query.to_lowercase();
        themes
            .into_iter()
            .filter(|t| t.meta.name.to_lowercase().contains(&query_lower))
            .collect()
    }

    /// Update theme preview based on current selection
    fn update_theme_preview(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if let Some(theme) = filtered.get(self.theme_picker_selected) {
            self.theme_preview = Some(theme.clone());
        } else {
            self.theme_preview = None;
        }
        cx.notify();
    }

    // =========================================================================
    // About dialog
    // =========================================================================

    pub fn show_about(&mut self, cx: &mut Context<Self>) {
        // Close console if open (About dialog needs focus)
        self.lua_console.visible = false;
        self.mode = Mode::About;
        cx.notify();
    }

    pub fn hide_about(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    // =========================================================================
    // License dialog
    // =========================================================================

    pub fn show_license(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.license_input.clear();
        self.license_error = None;
        self.mode = Mode::License;
        cx.notify();
    }

    pub fn hide_license(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.license_input.clear();
        self.license_error = None;
        cx.notify();
    }

    pub fn license_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.license_input.push(c);
        self.license_error = None;
        cx.notify();
    }

    pub fn license_backspace(&mut self, cx: &mut Context<Self>) {
        self.license_input.pop();
        self.license_error = None;
        cx.notify();
    }

    pub fn apply_license(&mut self, cx: &mut Context<Self>) {
        use crate::views::license_dialog::user_friendly_error;

        match visigrid_license::load_license(&self.license_input) {
            Ok(validation) => {
                if validation.valid {
                    self.status_message = Some(format!(
                        "License activated: {}",
                        visigrid_license::license_summary()
                    ));
                    self.hide_license(cx);
                } else {
                    // Convert technical error to user-friendly message
                    let raw_error = validation.error.as_deref().unwrap_or("Unknown error");
                    self.license_error = Some(user_friendly_error(raw_error));
                    cx.notify();
                }
            }
            Err(e) => {
                // Convert technical error to user-friendly message
                self.license_error = Some(user_friendly_error(&e));
                cx.notify();
            }
        }
    }

    pub fn clear_license(&mut self, cx: &mut Context<Self>) {
        visigrid_license::clear_license();
        self.status_message = Some("License removed".to_string());
        self.hide_license(cx);
    }

    // =========================================================================
    // Import/Export report dialogs
    // =========================================================================

    pub fn show_import_report(&mut self, cx: &mut Context<Self>) {
        if self.import_result.is_some() {
            self.lua_console.visible = false;
            self.mode = Mode::ImportReport;
            cx.notify();
        }
    }

    pub fn hide_import_report(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn show_export_report(&mut self, cx: &mut Context<Self>) {
        if self.export_result.is_some() {
            self.lua_console.visible = false;
            self.mode = Mode::ExportReport;
            cx.notify();
        }
    }

    pub fn hide_export_report(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Dismiss the import overlay (ESC during background import)
    /// Does NOT cancel the import - just hides the overlay
    pub fn dismiss_import_overlay(&mut self, cx: &mut Context<Self>) {
        self.import_overlay_visible = false;
        cx.notify();
    }

    // =========================================================================
    // Inspector panel
    // =========================================================================

    pub fn toggle_inspector_pin(&mut self, cx: &mut Context<Self>) {
        if self.inspector_pinned.is_some() {
            // Unpin: follow selection again
            self.inspector_pinned = None;
        } else {
            // Pin: lock to current selection
            self.inspector_pinned = Some(self.view_state.selected);
        }
        cx.notify();
    }

    /// Set a path trace from clicked input/output to the inspected cell.
    /// `from` is the clicked cell, `to` is the inspected cell.
    /// `forward` is true when tracing from input toward inspected cell.
    pub fn set_trace_path(&mut self, from_row: usize, from_col: usize, to_row: usize, to_col: usize, forward: bool, cx: &mut Context<Self>) {
        use visigrid_engine::cell_id::CellId;

        let sheet_id = self.sheet().id;
        let from = CellId::new(sheet_id, from_row, from_col);
        let to = CellId::new(sheet_id, to_row, to_col);

        let result = self.workbook.find_path(from, to, forward);

        if result.path.is_empty() {
            // No path found - clear any existing trace
            self.inspector_trace_path = None;
            self.inspector_trace_incomplete = result.truncated;
            if result.truncated {
                self.status_message = Some("Trace too large — refine by starting closer".to_string());
            }
        } else {
            self.inspector_trace_path = Some(result.path);
            self.inspector_trace_incomplete = result.has_dynamic_refs || result.truncated;
        }
        cx.notify();
    }

    /// Clear the current trace path.
    pub fn clear_trace_path(&mut self, cx: &mut Context<Self>) {
        if self.inspector_trace_path.is_some() {
            self.inspector_trace_path = None;
            self.inspector_trace_incomplete = false;
            cx.notify();
        }
    }
}
