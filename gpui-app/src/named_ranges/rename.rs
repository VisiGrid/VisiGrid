//! Rename Symbol (Ctrl+Shift+R) and Edit Description functionality

use gpui::*;
use visigrid_engine::named_range::is_valid_name;
use crate::app::Spreadsheet;
use crate::history::{CellChange, UndoAction};
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
        let named_ranges = self.workbook.list_named_ranges();
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
            let sheet = self.workbook.active_sheet();
            let (row, col) = self.view_state.selected;
            let cell = sheet.get_cell(row, col);
            let formula_text = self.get_formula_source(&cell.value);
            if let Some(formula) = formula_text {
                // Look for named range references in the formula
                self.find_named_range_in_formula(&formula)
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

        self.mode = Mode::RenameSymbol;
        self.rename_original_name = original.clone();
        self.rename_new_name = original;
        self.rename_select_all = true;  // First keystroke replaces entire name
        self.rename_validation_error = None;
        self.update_rename_affected_cells();
        cx.notify();
    }

    /// Hide the rename symbol dialog
    pub fn hide_rename_symbol(&mut self, cx: &mut Context<Self>) {
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
        self.validate_rename_name();
        cx.notify();
    }

    /// Delete the last character from the new name
    pub fn rename_symbol_backspace(&mut self, cx: &mut Context<Self>) {
        // Backspace also clears select_all mode but keeps existing text
        self.rename_select_all = false;
        self.rename_new_name.pop();
        self.validate_rename_name();
        cx.notify();
    }

    /// Validate the current new name
    fn validate_rename_name(&mut self) {
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
        if self.workbook.get_named_range(&self.rename_new_name).is_some() {
            self.rename_validation_error = Some(format!("'{}' already exists", self.rename_new_name));
            return;
        }

        self.rename_validation_error = None;
    }

    /// Update the list of affected cells (formulas using the named range)
    fn update_rename_affected_cells(&mut self) {
        self.rename_affected_cells.clear();

        let name_upper = self.rename_original_name.to_uppercase();
        let sheet = self.workbook.active_sheet();

        // Scan all cells for formulas that reference this named range
        for (&(row, col), cell) in sheet.cells_iter() {
            if let Some(formula) = self.get_formula_source(&cell.value) {
                if self.formula_references_name(&formula, &name_upper) {
                    self.rename_affected_cells.push((row, col));
                }
            }
        }
    }

    /// Check if a formula references a named range (case-insensitive)
    pub(crate) fn formula_references_name(&self, formula: &str, name_upper: &str) -> bool {
        // Simple check: look for the name as a word boundary
        // A proper implementation would parse the formula and check the AST
        let formula_upper = formula.to_uppercase();

        // Check for word boundaries using simple logic
        let name_len = name_upper.len();
        for (i, _) in formula_upper.match_indices(name_upper) {
            // Check if it's a word boundary (not part of a larger identifier)
            let before_ok = i == 0 || {
                let c = formula_upper.chars().nth(i - 1).unwrap_or(' ');
                !c.is_alphanumeric() && c != '_'
            };
            let after_ok = i + name_len >= formula_upper.len() || {
                let c = formula_upper.chars().nth(i + name_len).unwrap_or(' ');
                !c.is_alphanumeric() && c != '_'
            };
            if before_ok && after_ok {
                return true;
            }
        }
        false
    }

    /// Find a named range identifier in a formula string
    fn find_named_range_in_formula(&self, formula: &str) -> Option<String> {
        let named_ranges = self.workbook.list_named_ranges();
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
        // Validate first
        self.validate_rename_name();
        if self.rename_validation_error.is_some() {
            return;
        }

        let old_name = self.rename_original_name.clone();
        let new_name = self.rename_new_name.clone();

        // If names are the same (case-insensitive), just close
        if old_name.to_lowercase() == new_name.to_lowercase() {
            self.hide_rename_symbol(cx);
            return;
        }

        // Hide rename dialog and show impact preview
        self.mode = Mode::Navigation;  // Temporarily exit rename mode
        self.show_impact_preview_for_rename(&old_name, &new_name, cx);
    }

    /// Internal method to apply a rename (called from impact preview)
    pub(crate) fn apply_rename_internal(&mut self, old_name: &str, new_name: &str, cx: &mut Context<Self>) {
        // Collect all formula changes for undo
        let mut changes: Vec<CellChange> = Vec::new();
        let sheet_index = self.workbook.active_sheet_index();
        let old_name_upper = old_name.to_uppercase();

        // Find affected cells
        let affected_cells: Vec<(usize, usize)> = self.sheet().cells_iter()
            .filter_map(|((row, col), cell)| {
                let raw = cell.value.raw_display();
                if raw.starts_with('=') {
                    let formula_upper = raw.to_uppercase();
                    let contains_name = formula_upper
                        .split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.')
                        .any(|word| word == old_name_upper);
                    if contains_name {
                        return Some((*row, *col));
                    }
                }
                None
            })
            .collect();

        // Update formulas in all affected cells
        {
            let sheet = self.workbook.active_sheet();
            for &(row, col) in &affected_cells {
                let cell = sheet.get_cell(row, col);
                if let Some(formula) = self.get_formula_source(&cell.value) {
                    let new_formula = self.replace_name_in_formula(&formula, &old_name_upper, new_name);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value: formula,
                        new_value: new_formula,
                    });
                }
            }
        }

        // Apply the formula changes
        {
            let sheet = self.workbook.active_sheet_mut();
            for change in &changes {
                sheet.set_value(change.row, change.col, &change.new_value);
            }
        }

        // Rename the named range itself
        if let Err(e) = self.workbook.rename_named_range(old_name, new_name) {
            self.status_message = Some(format!("Failed to rename: {}", e));
            cx.notify();
            return;
        }

        // Record undo action
        if !changes.is_empty() {
            self.history.record_batch(sheet_index, changes.clone());
        }

        self.is_modified = true;
        self.bump_cells_rev();

        // Log the rename
        let formula_count = changes.len();
        let impact = if formula_count > 0 {
            Some(format!("{} formula{} updated", formula_count, if formula_count == 1 { "" } else { "s" }))
        } else {
            None
        };
        self.log_refactor(
            "Renamed named range",
            &format!("{} → {}", old_name, new_name),
            impact.as_deref(),
        );

        // Clear rename state
        self.rename_original_name.clear();
        self.rename_new_name.clear();
        self.rename_affected_cells.clear();
        cx.notify();
    }

    /// Replace a named range in a formula with a new name
    /// Handles case-insensitive matching while preserving surrounding text
    fn replace_name_in_formula(&self, formula: &str, old_name_upper: &str, new_name: &str) -> String {
        let mut result = String::with_capacity(formula.len());
        let formula_chars: Vec<char> = formula.chars().collect();
        let old_name_len = old_name_upper.len();
        let mut i = 0;

        while i < formula_chars.len() {
            // Try to match old name at this position
            let remaining: String = formula_chars[i..].iter().collect();
            let remaining_upper = remaining.to_uppercase();

            if remaining_upper.starts_with(old_name_upper) {
                // Check word boundaries
                let before_ok = i == 0 || {
                    let c = formula_chars[i - 1];
                    !c.is_alphanumeric() && c != '_'
                };
                let after_ok = i + old_name_len >= formula_chars.len() || {
                    let c = formula_chars[i + old_name_len];
                    !c.is_alphanumeric() && c != '_'
                };

                if before_ok && after_ok {
                    // Found a match - replace it
                    result.push_str(new_name);
                    i += old_name_len;
                    continue;
                }
            }

            result.push(formula_chars[i]);
            i += 1;
        }

        result
    }

    // =========================================================================
    // Edit Description
    // =========================================================================

    /// Show the edit description modal for a named range
    pub fn show_edit_description(&mut self, name: &str, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        // Get the current description
        let current_description = self.workbook.get_named_range(name)
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
        let name = self.edit_description_name.clone();
        let old_description = self.edit_description_original.clone();
        let new_description = if self.edit_description_value.is_empty() {
            None
        } else {
            Some(self.edit_description_value.clone())
        };

        // Only record if there's a change
        if old_description != new_description {
            // Apply the change
            let _ = self.workbook.named_ranges_mut().set_description(&name, new_description.clone());

            // Record for undo
            self.history.record_named_range_action(UndoAction::NamedRangeDescriptionChanged {
                name: name.clone(),
                old_description,
                new_description: new_description.clone(),
            });

            self.is_modified = true;

            // Log the edit
            let detail = match &new_description {
                Some(desc) => format!("{}: \"{}\"", name, desc),
                None => format!("{}: (cleared)", name),
            };
            self.log_refactor("Edited description", &detail, None);

            self.status_message = Some(format!("Updated description for '{}'", name));
        }

        // Close the modal
        self.hide_edit_description(cx);
    }
}
