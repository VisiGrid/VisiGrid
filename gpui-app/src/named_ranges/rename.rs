//! Rename Symbol (Ctrl+Shift+R) and Edit Description functionality

use gpui::{*};
use visigrid_engine::named_range::is_valid_name;
use crate::app::Spreadsheet;
use visigrid_engine::workbook::NamedRangeEdit;
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Rename Symbol (Ctrl+Shift+R)
    // =========================================================================

    /// Show the rename symbol dialog
    /// If `name` is provided, pre-fill with that named range
    pub fn show_rename_symbol(&mut self, name: Option<&str>, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        // Get list of named ranges
        let named_ranges = self.wb(cx).list_named_ranges();
        if named_ranges.is_empty() {
            self.status_message = Some("No named ranges defined".to_string());
            cx.notify();
            return;
        }

        // If name provided, use it; otherwise try to detect from current cell
        let original = if let Some(n) = name {
            n.to_string()
        } else {
            // Try to find a named range in the current cell's formula
            let (view_row, col) = self.view_state.selected;
            let row = self.row_view.view_to_data(view_row);
            let cell = self.sheet(cx).get_cell(row, col);
            let formula_text = self.get_formula_source(cell.value());
            if let Some(formula) = formula_text {
                // Look for named range references in the formula
                self.find_named_range_in_formula(&formula, cx)
            } else {
                None
            }.unwrap_or_else(|| {
                // No named range in current cell - use first available
                named_ranges.first().map(|nr| nr.name.clone()).unwrap_or_default()
            })
        };

        if original.is_empty() {
            self.status_message = Some("No named range to rename".to_string());
            cx.notify();
            return;
        }

        self.name_draft_error = None;
        self.name_draft = self.wb(cx).get_named_range(&original).cloned().map(|range| super::plan::NameDraft { revision: self.wb(cx).revision(), range });
        self.mode = Mode::RenameSymbol;
        self.rename_original_name = original.clone();
        self.rename_new_name = original;
        self.rename_select_all = true;  // First keystroke replaces entire name
        self.rename_validation_error = None;
        self.update_rename_affected_cells(cx);
        cx.notify();
    }

    /// Hide the rename symbol dialog
    pub fn hide_rename_symbol(&mut self, cx: &mut Context<Self>) {
        self.name_draft_error = None;
        self.name_draft = None;
        self.mode = Mode::Navigation;
        self.rename_original_name.clear();
        self.rename_new_name.clear();
        self.rename_select_all = false;
        self.rename_affected_cells.clear();
        self.rename_validation_error = None;
        cx.notify();
    }

    /// Insert a character into the new name
    pub fn rename_symbol_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        // If select_all is active, clear and start fresh
        if self.rename_select_all {
            self.rename_new_name.clear();
            self.rename_select_all = false;
        }
        self.rename_new_name.push(c);
        self.validate_rename_name(cx);
        cx.notify();
    }

    /// Delete the last character from the new name
    pub fn rename_symbol_backspace(&mut self, cx: &mut Context<Self>) {
        // Backspace also clears select_all mode but keeps existing text
        self.rename_select_all = false;
        self.rename_new_name.pop();
        self.validate_rename_name(cx);
        cx.notify();
    }

    /// Validate the current new name
    fn validate_rename_name(&mut self, cx: &App) {
        if self.rename_new_name.is_empty() {
            self.rename_validation_error = Some("Name cannot be empty".to_string());
            return;
        }

        // Check if it's the same as original (case-insensitive comparison for validity)
        if self.rename_new_name.to_lowercase() == self.rename_original_name.to_lowercase() {
            self.rename_validation_error = None;
            return;
        }

        // Check if name is valid
        if let Err(e) = is_valid_name(&self.rename_new_name) {
            self.rename_validation_error = Some(e);
            return;
        }

        // Check if name already exists
        if self.wb(cx).get_named_range(&self.rename_new_name).is_some() {
            self.rename_validation_error = Some(format!("'{}' already exists", self.rename_new_name));
            return;
        }

        self.rename_validation_error = None;
    }

    /// Update the list of affected cells (formulas using the named range)
    fn update_rename_affected_cells(&mut self, cx: &App) {
        self.rename_affected_cells = self.wb(cx).named_range_usages(&self.rename_original_name);
    }

    /// Check if a formula references a named range (case-insensitive)
    pub(crate) fn formula_references_name(&self, formula: &str, name_upper: &str) -> bool {
        visigrid_engine::formula::names::references_name(formula, name_upper)
    }

    /// Find a named range identifier in a formula string
    fn find_named_range_in_formula(&self, formula: &str, cx: &App) -> Option<String> {
        let named_ranges = self.wb(cx).list_named_ranges();
        let formula_upper = formula.to_uppercase();

        for nr in &named_ranges {
            let name_upper = nr.name.to_uppercase();
            if self.formula_references_name(&formula_upper, &name_upper) {
                return Some(nr.name.clone());
            }
        }
        None
    }

    /// Apply the rename operation
    pub fn confirm_rename_symbol(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        if let Err(error) = self.named_range_draft(cx) {
            self.rename_validation_error = Some(error); cx.notify(); return;
        }
        // Validate first
        self.validate_rename_name(cx);
        if self.rename_validation_error.is_some() {
            return;
        }

        let old_name = self.rename_original_name.clone();
        let new_name = self.rename_new_name.trim().to_string();

        // An exact no-op closes without touching history. Case changes are edits.
        if old_name == new_name {
            self.hide_rename_symbol(cx);
            return;
        }

        // Hide rename dialog and show impact preview
        self.mode = Mode::Navigation;  // Temporarily exit rename mode
        self.show_impact_preview_for_rename(&old_name, &new_name, cx);
    }

    /// Internal method to apply a rename (called from impact preview)
    pub(crate) fn apply_rename_internal(&mut self, old_name: &str, new_name: &str, cx: &mut Context<Self>) -> bool {
        let before = match self.named_range_draft(cx) {
            Ok(before) if before.name.eq_ignore_ascii_case(old_name) => before,
            _ => { let error = "The named range changed. Reopen the preview and try again.".to_string(); self.name_draft_error = Some(error.clone()); self.status_message = Some(error); cx.notify(); return false; }
        };
        if !self.apply_named_range_edit(NamedRangeEdit::Rename { before, name: new_name.into() }, format!("Rename named range: {old_name} to {new_name}"), cx) { return false; }
        self.log_refactor("Renamed named range", &format!("{} → {}", old_name, new_name), Some("Workbook references updated, including hidden records and Table rules"));
        self.rename_original_name.clear();
        self.rename_new_name.clear();
        self.rename_affected_cells.clear();
        true
    }

    // =========================================================================
    // Edit Description
    // =========================================================================

    /// Show the edit description modal for a named range
    pub fn show_edit_description(&mut self, name: &str, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.name_draft_error = None;
        self.name_draft = self.wb(cx).get_named_range(name).cloned().map(|range| super::plan::NameDraft { revision: self.wb(cx).revision(), range });
        // Get the current description
        let current_description = self.wb(cx).get_named_range(name)
            .and_then(|nr| nr.description.clone());

        self.edit_description_name = name.to_string();
        self.edit_description_value = current_description.clone().unwrap_or_default();
        self.edit_description_original = current_description;
        self.mode = Mode::EditDescription;
        cx.notify();
    }

    /// Hide the edit description modal without saving
    pub fn hide_edit_description(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.name_draft_error = None;
        self.name_draft = None;
        self.edit_description_name.clear();
        self.edit_description_value.clear();
        self.edit_description_original = None;
        cx.notify();
    }

    /// Insert a character into the description
    pub fn edit_description_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.edit_description_value.push(c);
        cx.notify();
    }

    /// Delete the last character from the description
    pub fn edit_description_backspace(&mut self, cx: &mut Context<Self>) {
        self.edit_description_value.pop();
        cx.notify();
    }

    /// Apply the edited description and record undo
    pub fn apply_edit_description(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        let before = match self.named_range_draft(cx) {
            Ok(before) => before,
            Err(error) => { self.name_draft_error = Some(error.clone()); self.status_message = Some(error); cx.notify(); return; }
        };
        let name = before.name.clone();
        let description = (!self.edit_description_value.is_empty()).then(|| self.edit_description_value.clone());
        let changed = before.description != description;
        if self.apply_named_range_edit(NamedRangeEdit::Description { before, description }, format!("Edit description: {name}"), cx) {
            if changed {
                let detail = format!("{name}: {}", if self.edit_description_value.is_empty() { "(cleared)" } else { &self.edit_description_value });
                self.log_refactor("Edited description", &detail, None);
            }
            self.status_message = Some(format!("Updated description for '{name}'"));
            self.hide_edit_description(cx);
        }
    }
}
