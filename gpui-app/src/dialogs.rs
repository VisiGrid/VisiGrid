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

use gpui::{*};
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
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;
        self.tab_chain_origin_col = None;  // Dialog breaks tab chain
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
        // Close validation dropdown when jumping to a cell
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::SelectionChanged,
            cx,
        );

        if let Some((row, col)) = Self::parse_cell_ref(&self.goto_input) {
            if row < NUM_ROWS && col < NUM_COLS {
                // If target is inside a merge, redirect to origin and select full merge
                let (target_row, target_col, merge_end) =
                    if let Some(merge) = self.sheet(cx).get_merge(row, col) {
                        (merge.start.0, merge.start.1, Some(merge.end))
                    } else {
                        (row, col, None)
                    };

                self.view_state.selected = (target_row, target_col);
                if let Some(end) = merge_end {
                    if end != (target_row, target_col) {
                        self.view_state.selection_end = Some(end);
                    } else {
                        self.view_state.selection_end = None;
                    }
                } else {
                    self.view_state.selection_end = None;
                }
                self.view_state.additional_selections.clear();
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
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
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
    // Paste Special dialog
    // =========================================================================

    pub fn show_paste_special(&mut self, cx: &mut Context<Self>) {
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;

        // Initialize with last selected mode (session memory)
        self.paste_special_dialog.selected = self.last_paste_special_mode;
        self.mode = Mode::PasteSpecial;
        cx.notify();
    }

    pub fn hide_paste_special(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Move selection up in the Paste Special dialog
    pub fn paste_special_up(&mut self, cx: &mut Context<Self>) {
        use crate::app::PasteType;
        let types = PasteType::all();
        let current_idx = types.iter().position(|t| *t == self.paste_special_dialog.selected).unwrap_or(0);
        if current_idx > 0 {
            self.paste_special_dialog.selected = types[current_idx - 1];
            cx.notify();
        }
    }

    /// Move selection down in the Paste Special dialog
    pub fn paste_special_down(&mut self, cx: &mut Context<Self>) {
        use crate::app::PasteType;
        let types = PasteType::all();
        let current_idx = types.iter().position(|t| *t == self.paste_special_dialog.selected).unwrap_or(0);
        if current_idx < types.len() - 1 {
            self.paste_special_dialog.selected = types[current_idx + 1];
            cx.notify();
        }
    }

    /// Execute the selected paste type and close the dialog
    pub fn apply_paste_special(&mut self, cx: &mut Context<Self>) {
        use crate::app::PasteType;

        // Remember selection for next time (session memory)
        self.last_paste_special_mode = self.paste_special_dialog.selected;

        // Close dialog first
        self.mode = Mode::Navigation;

        // Execute the selected paste operation
        match self.paste_special_dialog.selected {
            PasteType::All => self.paste(cx),
            PasteType::Values => self.paste_values(cx),
            PasteType::Formulas => self.paste_formulas(cx),
            PasteType::Formats => self.paste_formats(cx),
        }

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

        let sheet_id = self.sheet(cx).id;
        let from = CellId::new(sheet_id, from_row, from_col);
        let to = CellId::new(sheet_id, to_row, to_col);

        let result = self.wb(cx).find_path(from, to, forward);

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

    // =========================================================================
    // Data Validation dialog (Phase 4)
    // =========================================================================

    pub fn show_validation_dialog(&mut self, cx: &mut Context<Self>) {
        use visigrid_engine::validation::CellRange;

        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;

        // Reset dialog state
        self.validation_dialog.reset();

        // Capture the target range (current selection)
        let (row, col) = self.view_state.selected;
        let range = if let Some((end_row, end_col)) = self.view_state.selection_end {
            let (min_row, max_row) = if row <= end_row { (row, end_row) } else { (end_row, row) };
            let (min_col, max_col) = if col <= end_col { (col, end_col) } else { (end_col, col) };
            CellRange::new(min_row, min_col, max_row, max_col)
        } else {
            CellRange::single(row, col)
        };
        self.validation_dialog.target_range = Some(range.clone());

        // Load existing validation if present (check first cell of selection)
        // Clone the rule to release the borrow before modifying validation_dialog
        let existing_rule = self.sheet(cx).validations.get(row, col).cloned();
        if let Some(rule) = existing_rule {
            self.validation_dialog.load_from_rule(&rule);
        }

        self.mode = Mode::ValidationDialog;
        cx.notify();
    }

    pub fn hide_validation_dialog(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.validation_dialog.reset();
        cx.notify();
    }

    pub fn apply_validation_dialog(&mut self, cx: &mut Context<Self>) {
        use crate::app::{ValidationTypeOption, NumericOperatorOption};
        use crate::history::UndoAction;
        use visigrid_engine::validation::{ValidationRule, ValidationType, ListSource, NumericConstraint, ComparisonOperator, CellRange};

        // Copy all needed values upfront to avoid borrow issues
        let validation_type = self.validation_dialog.validation_type;
        let list_source_str = self.validation_dialog.list_source.clone();
        let show_dropdown = self.validation_dialog.show_dropdown;
        let numeric_operator = self.validation_dialog.numeric_operator;
        let value1_str = self.validation_dialog.value1.clone();
        let value2_str = self.validation_dialog.value2.clone();
        let ignore_blank = self.validation_dialog.ignore_blank;
        let target_range = self.validation_dialog.target_range.clone();

        // Any Value = Clear validation (not a real rule)
        // This matches Excel semantics: "Any value" means "no validation"
        if validation_type == ValidationTypeOption::AnyValue {
            if let Some(range) = target_range {
                let sheet_index = self.sheet_index(cx);

                // Capture rules to be cleared for undo
                let cleared_rules: Vec<(CellRange, ValidationRule)> = self.sheet(cx)
                    .validations
                    .iter()
                    .filter(|(r, _)| r.overlaps(&range))
                    .map(|(r, v)| (*r, v.clone()))
                    .collect();

                // Clear the rules
                self.active_sheet_mut(cx, |s| s.validations.clear_range(&range));
                self.bump_cells_rev();

                // Clear invalid markers in the range
                for row in range.start_row..=range.end_row {
                    for col in range.start_col..=range.end_col {
                        self.invalid_cells.remove(&(row, col));
                        self.validation_failures.retain(|&(r, c)| r != row || c != col);
                    }
                }

                // Record history (only if something was actually cleared)
                if !cleared_rules.is_empty() {
                    self.history.record_action_with_provenance(
                        UndoAction::ValidationCleared {
                            sheet_index,
                            range: range.clone(),
                            cleared_rules,
                        },
                        None,
                    );
                }

                let cell_count = range.cell_count();
                self.status_message = Some(format!(
                    "Cleared validation from {} cell{}",
                    cell_count,
                    if cell_count == 1 { "" } else { "s" }
                ));

                self.is_modified = true;
            }

            self.hide_validation_dialog(cx);
            return;
        }

        // Build the validation rule based on current state
        let rule = match validation_type {
            ValidationTypeOption::AnyValue => unreachable!(), // Handled above
            ValidationTypeOption::List => {
                // Parse list source
                let source = list_source_str.trim();
                if source.is_empty() {
                    self.validation_dialog.error = Some("List source is required".to_string());
                    cx.notify();
                    return;
                }

                // Determine source type: range ref, named range, or inline list
                let list_source = if source.contains('!') || source.chars().next().map(|c| c.is_ascii_alphabetic()).unwrap_or(false) && source.contains(':') {
                    // Looks like a range reference (e.g., "A1:A10" or "Sheet1!A1:A10")
                    ListSource::Range(source.to_string())
                } else if source.chars().next().map(|c| c.is_ascii_alphabetic() || c == '_').unwrap_or(false) && !source.contains(',') {
                    // Looks like a named range (starts with letter/underscore, no commas)
                    ListSource::NamedRange(source.to_string())
                } else {
                    // Inline comma-separated list
                    let items: Vec<String> = source.split(',').map(|s| s.trim().to_string()).collect();
                    if items.is_empty() || items.iter().all(|s| s.is_empty()) {
                        self.validation_dialog.error = Some("At least one list item is required".to_string());
                        cx.notify();
                        return;
                    }
                    ListSource::Inline(items)
                };

                ValidationRule::new(ValidationType::List(list_source))
                    .with_show_dropdown(show_dropdown)
            }
            ValidationTypeOption::WholeNumber | ValidationTypeOption::Decimal => {
                // Parse numeric constraint
                let operator = match numeric_operator {
                    NumericOperatorOption::Between => ComparisonOperator::Between,
                    NumericOperatorOption::NotBetween => ComparisonOperator::NotBetween,
                    NumericOperatorOption::EqualTo => ComparisonOperator::EqualTo,
                    NumericOperatorOption::NotEqualTo => ComparisonOperator::NotEqualTo,
                    NumericOperatorOption::GreaterThan => ComparisonOperator::GreaterThan,
                    NumericOperatorOption::LessThan => ComparisonOperator::LessThan,
                    NumericOperatorOption::GreaterThanOrEqual => ComparisonOperator::GreaterThanOrEqual,
                    NumericOperatorOption::LessThanOrEqual => ComparisonOperator::LessThanOrEqual,
                };

                // Parse value1
                let v1_str = value1_str.trim();
                if v1_str.is_empty() {
                    self.validation_dialog.error = Some("Value is required".to_string());
                    cx.notify();
                    return;
                }
                let value1 = Self::parse_constraint_value(v1_str);

                // Parse value2 for between operators
                let value2 = if numeric_operator.needs_two_values() {
                    let v2_str = value2_str.trim();
                    if v2_str.is_empty() {
                        self.validation_dialog.error = Some("Maximum value is required".to_string());
                        cx.notify();
                        return;
                    }
                    Some(Self::parse_constraint_value(v2_str))
                } else {
                    None
                };

                let constraint = NumericConstraint { operator, value1, value2 };

                if validation_type == ValidationTypeOption::WholeNumber {
                    ValidationRule::new(ValidationType::WholeNumber(constraint))
                } else {
                    ValidationRule::new(ValidationType::Decimal(constraint))
                }
            }
        };

        // Apply common options
        let rule = rule.with_ignore_blank(ignore_blank);

        // Apply to target range
        if let Some(range) = target_range {
            let sheet_index = self.sheet_index(cx);

            // Replace-in-range semantics: capture overlapping rules before clearing
            let previous_rules: Vec<(CellRange, ValidationRule)> = self.sheet(cx)
                .validations
                .iter()
                .filter(|(r, _)| r.overlaps(&range))
                .map(|(r, v)| (*r, v.clone()))
                .collect();

            // Clear overlapping rules, then set new rule
            self.active_sheet_mut(cx, |s| s.validations.clear_range(&range));
            self.active_sheet_mut(cx, |s| s.validations.set(range.clone(), rule.clone()));
            self.bump_cells_rev();

            // Record history for undo
            self.history.record_action_with_provenance(
                UndoAction::ValidationSet {
                    sheet_index,
                    range: range.clone(),
                    previous_rules,
                    new_rule: rule.clone(),
                },
                None,
            );

            // Validate cells in range and mark invalid ones
            let invalid_count = self.validate_and_mark_range(&range, cx);

            // Build compact rule summary for status
            let rule_summary = Self::compact_rule_summary(
                validation_type,
                &list_source_str,
                numeric_operator,
                &value1_str,
                &value2_str,
            );

            let cell_count = range.cell_count();
            if invalid_count > 0 {
                self.status_message = Some(format!(
                    "Applied {} to {} cell{}. {} invalid — press F8 to jump",
                    rule_summary,
                    cell_count,
                    if cell_count == 1 { "" } else { "s" },
                    invalid_count
                ));
            } else {
                self.status_message = Some(format!(
                    "Applied {} to {} cell{}",
                    rule_summary,
                    cell_count,
                    if cell_count == 1 { "" } else { "s" }
                ));
            }

            self.is_modified = true;
        }

        self.hide_validation_dialog(cx);
    }

    /// Generate a compact human-readable summary of a validation rule.
    /// Examples: "List(Open,Closed)", "Whole Number (Between 1 and 100)", "Decimal (>= 0)"
    fn compact_rule_summary(
        validation_type: crate::app::ValidationTypeOption,
        list_source: &str,
        numeric_op: crate::app::NumericOperatorOption,
        value1: &str,
        value2: &str,
    ) -> String {
        use crate::app::{ValidationTypeOption, NumericOperatorOption};

        match validation_type {
            ValidationTypeOption::AnyValue => "Any Value".to_string(),
            ValidationTypeOption::List => {
                // Truncate long list sources
                let source = list_source.trim();
                if source.len() > 30 {
                    format!("List({}...)", &source[..27])
                } else {
                    format!("List({})", source)
                }
            }
            ValidationTypeOption::WholeNumber | ValidationTypeOption::Decimal => {
                let type_name = if validation_type == ValidationTypeOption::WholeNumber {
                    "Whole Number"
                } else {
                    "Decimal"
                };

                let constraint = match numeric_op {
                    NumericOperatorOption::Between => format!("Between {} and {}", value1, value2),
                    NumericOperatorOption::NotBetween => format!("Not between {} and {}", value1, value2),
                    NumericOperatorOption::EqualTo => format!("= {}", value1),
                    NumericOperatorOption::NotEqualTo => format!("≠ {}", value1),
                    NumericOperatorOption::GreaterThan => format!("> {}", value1),
                    NumericOperatorOption::LessThan => format!("< {}", value1),
                    NumericOperatorOption::GreaterThanOrEqual => format!(">= {}", value1),
                    NumericOperatorOption::LessThanOrEqual => format!("<= {}", value1),
                };

                format!("{} ({})", type_name, constraint)
            }
        }
    }

    /// Validate cells in a range and mark invalid ones with circles.
    /// Returns the count of invalid cells.
    fn validate_and_mark_range(&mut self, range: &visigrid_engine::validation::CellRange, cx: &App) -> usize {
        use visigrid_engine::validation::ValidationResult;
        use visigrid_engine::workbook::Workbook;

        let mut invalid_count = 0;
        let sheet_index = self.sheet_index(cx);

        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                let display_value = self.sheet(cx).get_display(row, col);
                // Skip empty cells if ignore_blank is true (handled by validation)
                if display_value.is_empty() {
                    // Clear any existing invalid marker
                    self.invalid_cells.remove(&(row, col));
                    continue;
                }
                let result = self.wb(cx).validate_cell_input(sheet_index, row, col, &display_value);
                match result {
                    ValidationResult::Invalid { reason, .. } => {
                        let failure_reason = Workbook::classify_failure_reason(&reason);
                        self.invalid_cells.insert((row, col), failure_reason);
                        // Also add to navigation list if not already present
                        if !self.validation_failures.contains(&(row, col)) {
                            self.validation_failures.push((row, col));
                        }
                        invalid_count += 1;
                    }
                    ValidationResult::Valid => {
                        // Clear any existing invalid marker
                        self.invalid_cells.remove(&(row, col));
                        self.validation_failures.retain(|&(r, c)| r != row || c != col);
                    }
                }
            }
        }

        // Re-sort failures in row-major order
        self.validation_failures.sort_by_key(|&(r, c)| (r, c));

        invalid_count
    }

    pub fn clear_validation_dialog(&mut self, cx: &mut Context<Self>) {
        use crate::history::UndoAction;
        use visigrid_engine::validation::CellRange;

        // Clone target range to avoid borrow issues
        let target_range = self.validation_dialog.target_range.clone();

        // Clear validation from target range
        if let Some(range) = target_range {
            let sheet_index = self.sheet_index(cx);

            // Capture rules to be cleared for undo
            let cleared_rules: Vec<(CellRange, visigrid_engine::validation::ValidationRule)> = self.sheet(cx)
                .validations
                .iter()
                .filter(|(r, _)| r.overlaps(&range))
                .map(|(r, v)| (*r, v.clone()))
                .collect();

            // Clear the rules
            self.active_sheet_mut(cx, |s| s.validations.clear_range(&range));
            self.bump_cells_rev();

            // Clear invalid markers in the range
            for row in range.start_row..=range.end_row {
                for col in range.start_col..=range.end_col {
                    self.invalid_cells.remove(&(row, col));
                    self.validation_failures.retain(|&(r, c)| r != row || c != col);
                }
            }

            // Record history for undo (only if something was actually cleared)
            if !cleared_rules.is_empty() {
                self.history.record_action_with_provenance(
                    UndoAction::ValidationCleared {
                        sheet_index,
                        range: range.clone(),
                        cleared_rules,
                    },
                    None,
                );
            }

            let cell_count = range.cell_count();
            self.status_message = Some(format!(
                "Cleared validation from {} cell{}",
                cell_count,
                if cell_count == 1 { "" } else { "s" }
            ));

            self.is_modified = true;
        }

        self.hide_validation_dialog(cx);
    }

    // ========================================================================
    // Validation Exclusions
    // ========================================================================

    /// Exclude the current selection from validation.
    /// Cells in excluded ranges are not validated regardless of any rules.
    pub fn exclude_from_validation(&mut self, cx: &mut Context<Self>) {
        use crate::history::UndoAction;
        use visigrid_engine::validation::CellRange;

        let (row, col) = self.view_state.selected;
        let (end_row, end_col) = self.view_state.selection_end.unwrap_or((row, col));

        let range = CellRange::new(
            row.min(end_row),
            col.min(end_col),
            row.max(end_row),
            col.max(end_col),
        );

        let sheet_index = self.sheet_index(cx);

        // Add the exclusion
        self.active_sheet_mut(cx, |s| s.validations.exclude(range));
        self.bump_cells_rev();

        // Clear invalid markers in the excluded range
        for r in range.start_row..=range.end_row {
            for c in range.start_col..=range.end_col {
                self.invalid_cells.remove(&(r, c));
                self.validation_failures.retain(|&(rr, cc)| rr != r || cc != c);
            }
        }

        // Record history for undo
        self.history.record_action_with_provenance(
            UndoAction::ValidationExcluded {
                sheet_index,
                range: range.clone(),
            },
            None,
        );

        let cell_count = range.cell_count();
        self.status_message = Some(format!(
            "Excluded {} cell{} from validation",
            cell_count,
            if cell_count == 1 { "" } else { "s" }
        ));

        self.is_modified = true;
        cx.notify();
    }

    /// Clear validation exclusions in the current selection.
    pub fn clear_validation_exclusions(&mut self, cx: &mut Context<Self>) {
        use crate::history::UndoAction;
        use visigrid_engine::validation::CellRange;

        let (row, col) = self.view_state.selected;
        let (end_row, end_col) = self.view_state.selection_end.unwrap_or((row, col));

        let range = CellRange::new(
            row.min(end_row),
            col.min(end_col),
            row.max(end_row),
            col.max(end_col),
        );

        let sheet_index = self.sheet_index(cx);

        // Capture exclusions that will be cleared (for undo)
        let cleared_exclusions: Vec<CellRange> = self.sheet(cx)
            .validations
            .exclusions_in_range(&range);

        if cleared_exclusions.is_empty() {
            self.status_message = Some("No exclusions to clear in selection".to_string());
            cx.notify();
            return;
        }

        // Clear the exclusions
        self.active_sheet_mut(cx, |s| s.validations.clear_exclusions_in_range(&range));
        self.bump_cells_rev();

        // Revalidate the range (cells may now be validated again)
        let invalid_count = self.validate_and_mark_range(&range, cx);

        // Record history for undo
        self.history.record_action_with_provenance(
            UndoAction::ValidationExclusionCleared {
                sheet_index,
                range: range.clone(),
                cleared_exclusions,
            },
            None,
        );

        let cell_count = range.cell_count();
        if invalid_count > 0 {
            self.status_message = Some(format!(
                "Cleared exclusions from {} cell{}. {} now invalid — press F8 to jump",
                cell_count,
                if cell_count == 1 { "" } else { "s" },
                invalid_count
            ));
        } else {
            self.status_message = Some(format!(
                "Cleared exclusions from {} cell{}",
                cell_count,
                if cell_count == 1 { "" } else { "s" }
            ));
        }

        self.is_modified = true;
        cx.notify();
    }

    /// Parse a constraint value string (number, cell ref, or formula)
    fn parse_constraint_value(s: &str) -> visigrid_engine::validation::ConstraintValue {
        use visigrid_engine::validation::ConstraintValue;

        // Try parsing as number first
        if let Ok(n) = s.parse::<f64>() {
            return ConstraintValue::Number(n);
        }

        // Check if it's a cell reference (starts with letter, contains only alphanumeric)
        if s.chars().next().map(|c| c.is_ascii_alphabetic()).unwrap_or(false)
            && s.chars().all(|c| c.is_ascii_alphanumeric() || c == '$' || c == '!' || c == ':')
        {
            return ConstraintValue::CellRef(s.to_string());
        }

        // Treat as formula
        ConstraintValue::Formula(s.to_string())
    }

    // Validation dialog input handling
    pub fn validation_dialog_type_char(&mut self, c: char, cx: &mut Context<Self>) {
        use crate::app::ValidationDialogFocus;

        match self.validation_dialog.focus {
            ValidationDialogFocus::Source => {
                self.validation_dialog.list_source.push(c);
                self.validation_dialog.error = None;
            }
            ValidationDialogFocus::Value1 => {
                self.validation_dialog.value1.push(c);
                self.validation_dialog.error = None;
            }
            ValidationDialogFocus::Value2 => {
                self.validation_dialog.value2.push(c);
                self.validation_dialog.error = None;
            }
            _ => {}
        }
        cx.notify();
    }

    pub fn validation_dialog_backspace(&mut self, cx: &mut Context<Self>) {
        use crate::app::ValidationDialogFocus;

        match self.validation_dialog.focus {
            ValidationDialogFocus::Source => {
                self.validation_dialog.list_source.pop();
                self.validation_dialog.error = None;
            }
            ValidationDialogFocus::Value1 => {
                self.validation_dialog.value1.pop();
                self.validation_dialog.error = None;
            }
            ValidationDialogFocus::Value2 => {
                self.validation_dialog.value2.pop();
                self.validation_dialog.error = None;
            }
            _ => {}
        }
        cx.notify();
    }

    pub fn validation_dialog_tab(&mut self, shift: bool, cx: &mut Context<Self>) {
        use crate::app::{ValidationDialogFocus, ValidationTypeOption};

        // Close any open dropdowns
        self.validation_dialog.type_dropdown_open = false;
        self.validation_dialog.operator_dropdown_open = false;

        // Cycle through focusable fields based on validation type
        let fields: Vec<ValidationDialogFocus> = match self.validation_dialog.validation_type {
            ValidationTypeOption::AnyValue => {
                vec![ValidationDialogFocus::TypeDropdown]
            }
            ValidationTypeOption::List => {
                vec![ValidationDialogFocus::TypeDropdown, ValidationDialogFocus::Source]
            }
            ValidationTypeOption::WholeNumber | ValidationTypeOption::Decimal => {
                if self.validation_dialog.numeric_operator.needs_two_values() {
                    vec![
                        ValidationDialogFocus::TypeDropdown,
                        ValidationDialogFocus::OperatorDropdown,
                        ValidationDialogFocus::Value1,
                        ValidationDialogFocus::Value2,
                    ]
                } else {
                    vec![
                        ValidationDialogFocus::TypeDropdown,
                        ValidationDialogFocus::OperatorDropdown,
                        ValidationDialogFocus::Value1,
                    ]
                }
            }
        };

        let current_idx = fields.iter().position(|f| *f == self.validation_dialog.focus).unwrap_or(0);
        let next_idx = if shift {
            if current_idx == 0 { fields.len() - 1 } else { current_idx - 1 }
        } else {
            (current_idx + 1) % fields.len()
        };

        self.validation_dialog.focus = fields[next_idx];
        cx.notify();
    }

    // =========================================================================
    // AI Settings dialog
    // =========================================================================

    pub fn show_ai_settings(&mut self, cx: &mut Context<Self>) {
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;

        // Load current settings
        self.ai_settings.load_from_config();

        self.mode = Mode::AISettings;
        cx.notify();
    }

    pub fn hide_ai_settings(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.ai_settings.reset();
        cx.notify();
    }

    pub fn apply_ai_settings(&mut self, cx: &mut Context<Self>) {
        if let Err(e) = self.ai_settings.save_to_config() {
            self.ai_settings.error = Some(format!("Failed to save: {}", e));
            cx.notify();
            return;
        }

        self.status_message = Some(format!(
            "AI settings updated: {}",
            if self.ai_settings.provider == crate::app::AIProviderOption::None {
                "AI disabled".to_string()
            } else {
                format!("{} ({})", self.ai_settings.provider.label(), self.ai_settings.effective_model())
            }
        ));
        self.hide_ai_settings(cx);
    }

    pub fn ai_settings_set_key(&mut self, cx: &mut Context<Self>) {
        use visigrid_config::ai::set_api_key;

        let key = self.ai_settings.key_input.trim();
        if key.is_empty() {
            self.ai_settings.error = Some("API key cannot be empty".to_string());
            cx.notify();
            return;
        }

        let provider_name = match self.ai_settings.provider {
            crate::app::AIProviderOption::OpenAI => "openai",
            crate::app::AIProviderOption::Anthropic => "anthropic",
            crate::app::AIProviderOption::Gemini => "gemini",
            crate::app::AIProviderOption::Grok => "grok",
            _ => {
                self.ai_settings.error = Some("Select a provider first".to_string());
                cx.notify();
                return;
            }
        };

        match set_api_key(provider_name, key) {
            Ok(()) => {
                self.ai_settings.key_present = true;
                self.ai_settings.key_source = "keychain".to_string();
                // Cache key for session (workaround for keychain timing)
                self.ai_session_key = Some(key.to_string());
                self.ai_settings.key_input.clear();
                self.ai_settings.error = None;
                // Set session flag to work around keychain timing issues
                self.ai_key_validated_this_session = true;
                self.status_message = Some(format!("{} API key saved to keychain", self.ai_settings.provider.label()));
            }
            Err(e) => {
                self.ai_settings.error = Some(e);
            }
        }
        cx.notify();
    }

    pub fn ai_settings_clear_key(&mut self, cx: &mut Context<Self>) {
        use visigrid_config::ai::delete_api_key;

        let provider_name = match self.ai_settings.provider {
            crate::app::AIProviderOption::OpenAI => "openai",
            crate::app::AIProviderOption::Anthropic => "anthropic",
            crate::app::AIProviderOption::Gemini => "gemini",
            crate::app::AIProviderOption::Grok => "grok",
            _ => return,
        };

        match delete_api_key(provider_name) {
            Ok(()) => {
                self.ai_settings.key_present = false;
                self.ai_settings.key_source = "none".to_string();
                // Clear session flag and cached key since key was removed
                self.ai_key_validated_this_session = false;
                self.ai_session_key = None;
                self.status_message = Some(format!("{} API key removed", self.ai_settings.provider.label()));
            }
            Err(e) => {
                self.ai_settings.error = Some(e);
            }
        }
        cx.notify();
    }

    pub fn ai_settings_type_char(&mut self, c: char, cx: &mut Context<Self>) {
        use crate::app::AISettingsFocus;

        match self.ai_settings.focus {
            AISettingsFocus::Model => {
                self.ai_settings.model.push(c);
                self.ai_settings.error = None;
            }
            AISettingsFocus::Endpoint => {
                self.ai_settings.endpoint.push(c);
                self.ai_settings.error = None;
            }
            AISettingsFocus::KeyInput => {
                self.ai_settings.key_input.push(c);
                self.ai_settings.error = None;
            }
            _ => {}
        }
        cx.notify();
    }

    pub fn ai_settings_backspace(&mut self, cx: &mut Context<Self>) {
        use crate::app::AISettingsFocus;

        match self.ai_settings.focus {
            AISettingsFocus::Model => {
                self.ai_settings.model.pop();
            }
            AISettingsFocus::Endpoint => {
                self.ai_settings.endpoint.pop();
            }
            AISettingsFocus::KeyInput => {
                self.ai_settings.key_input.pop();
            }
            _ => {}
        }
        cx.notify();
    }

    pub fn ai_settings_paste(&mut self, cx: &mut Context<Self>) {
        use crate::app::AISettingsFocus;

        if let Some(item) = cx.read_from_clipboard() {
            if let Some(text) = item.text() {
                // Clean the pasted text (remove newlines, trim)
                let clean_text: String = text.chars()
                    .filter(|c| !c.is_control())
                    .collect();

                match self.ai_settings.focus {
                    AISettingsFocus::Model => {
                        self.ai_settings.model.push_str(&clean_text);
                    }
                    AISettingsFocus::Endpoint => {
                        self.ai_settings.endpoint.push_str(&clean_text);
                    }
                    AISettingsFocus::KeyInput => {
                        self.ai_settings.key_input.push_str(&clean_text);
                    }
                    _ => {}
                }
                cx.notify();
            }
        }
    }

    pub fn ai_settings_tab(&mut self, shift: bool, cx: &mut Context<Self>) {
        use crate::app::{AISettingsFocus, AIProviderOption};

        // Close dropdown
        self.ai_settings.provider_dropdown_open = false;

        // Build list of focusable fields based on provider
        let mut fields = vec![AISettingsFocus::Provider];

        match self.ai_settings.provider {
            AIProviderOption::None => {
                // No additional fields for disabled
            }
            AIProviderOption::Local => {
                fields.push(AISettingsFocus::Model);
                fields.push(AISettingsFocus::Endpoint);
            }
            AIProviderOption::OpenAI | AIProviderOption::Anthropic | AIProviderOption::Gemini | AIProviderOption::Grok => {
                fields.push(AISettingsFocus::Model);
                fields.push(AISettingsFocus::KeyInput);
            }
        }

        let current_idx = fields.iter().position(|f| *f == self.ai_settings.focus).unwrap_or(0);
        let next_idx = if shift {
            if current_idx == 0 { fields.len() - 1 } else { current_idx - 1 }
        } else {
            (current_idx + 1) % fields.len()
        };

        self.ai_settings.focus = fields[next_idx];
        cx.notify();
    }

    /// Validate the current AI configuration.
    /// This checks credentials and basic reachability, NOT feature functionality.
    pub fn ai_settings_test_connection(&mut self, cx: &mut Context<Self>) {
        use crate::app::AITestStatus;
        use visigrid_config::ai::{ResolvedAIConfig, ValidationResult, AIConfigStatus};

        // Check if the provider needs a key but dialog shows no key
        let provider_needs_key = self.ai_settings.needs_api_key();
        if provider_needs_key && !self.ai_settings.key_present {
            self.ai_settings.test_status = AITestStatus::Error("No API key configured".to_string());
            cx.notify();
            return;
        }

        // If we have a key (as indicated by dialog state), skip the key check
        // and just validate the rest of the config. This handles the case where
        // the key was just set in this session and keychain reads might not
        // reflect it immediately.
        if provider_needs_key && self.ai_settings.key_present {
            // Key was set in this session - trust the dialog state
            self.ai_settings.test_status = AITestStatus::Testing;
            cx.notify();

            // For cloud providers, we can't actually test without making an API call
            // Just report that config looks valid
            let provider_name = self.ai_settings.provider.label();
            self.ai_settings.test_status = AITestStatus::Success(
                format!("API key present ({}) - {} configured", self.ai_settings.key_source, provider_name)
            );
            cx.notify();
            return;
        }

        // Build a temporary resolved config from current dialog state
        // (This validates what the user has configured, not what's saved yet)
        let mut settings = visigrid_config::settings::Settings::load();
        settings.ai.provider = self.ai_settings.provider.to_config();
        settings.ai.model = self.ai_settings.model.clone();
        settings.ai.endpoint = if self.ai_settings.endpoint.is_empty() {
            None
        } else {
            Some(self.ai_settings.endpoint.clone())
        };
        settings.ai.privacy_mode = self.ai_settings.privacy_mode;
        settings.ai.allow_proposals = self.ai_settings.allow_proposals;

        let config = ResolvedAIConfig::from_settings(&settings.ai);

        // Update status to validating
        self.ai_settings.test_status = AITestStatus::Testing;
        cx.notify();

        // Run validation (note: this is synchronous for now)
        let result = config.validate_config();

        self.ai_settings.test_status = match result {
            ValidationResult::Valid(msg) => AITestStatus::Success(msg),
            ValidationResult::Invalid(msg) => {
                // Special case: if key check fails but dialog shows key is set,
                // there might be a keychain read timing issue - trust the dialog
                if msg.contains("API key") && self.ai_settings.key_present {
                    AITestStatus::Success(format!("API key present ({}) - configured", self.ai_settings.key_source))
                } else {
                    AITestStatus::Error(msg)
                }
            },
            ValidationResult::Skipped(msg) => AITestStatus::Success(msg),
        };
        cx.notify();
    }

    // =========================================================================
    // Ask AI dialog
    // =========================================================================

    pub fn show_ask_ai(&mut self, cx: &mut Context<Self>) {
        use visigrid_config::ai::{ResolvedAIConfig, AIConfigStatus};

        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;

        // Check AI is configured
        let config = ResolvedAIConfig::load();
        match config.status {
            AIConfigStatus::Disabled => {
                self.status_message = Some("AI is disabled. Enable in Preferences → AI.".to_string());
                cx.notify();
                return;
            }
            AIConfigStatus::MissingKey => {
                // Check session flag - keychain may have timing issues after recent save
                if !self.ai_key_validated_this_session {
                    self.status_message = Some("API key not configured. Set in Preferences → AI.".to_string());
                    cx.notify();
                    return;
                }
                // Key was validated this session, proceed despite keychain timing issue
            }
            AIConfigStatus::NotImplemented => {
                self.status_message = Some(format!(
                    "{} provider not yet implemented.",
                    config.provider.name()
                ));
                cx.notify();
                return;
            }
            AIConfigStatus::Error => {
                self.status_message = Some(config.blocking_reason.unwrap_or_else(|| "AI configuration error".to_string()));
                cx.notify();
                return;
            }
            AIConfigStatus::Ready => {
                // Continue to show dialog
            }
        }

        // Reset dialog state and set verb
        self.ask_ai.verb = crate::app::AiVerb::InsertFormula;
        self.ask_ai.reset();

        // Set mode first, then update context
        self.mode = Mode::AiDialog;

        // Build context from current selection
        self.ask_ai_update_context(cx);
    }

    pub fn show_analyze(&mut self, cx: &mut Context<Self>) {
        use visigrid_config::ai::{ResolvedAIConfig, AIConfigStatus};

        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;

        // Check AI is configured
        let config = ResolvedAIConfig::load();
        match config.status {
            AIConfigStatus::Disabled => {
                self.status_message = Some("AI is disabled. Enable in Preferences → AI.".to_string());
                cx.notify();
                return;
            }
            AIConfigStatus::MissingKey => {
                if !self.ai_key_validated_this_session {
                    self.status_message = Some("API key not configured. Set in Preferences → AI.".to_string());
                    cx.notify();
                    return;
                }
            }
            AIConfigStatus::NotImplemented => {
                self.status_message = Some(format!(
                    "{} provider not yet implemented.",
                    config.provider.name()
                ));
                cx.notify();
                return;
            }
            AIConfigStatus::Error => {
                self.status_message = Some(config.blocking_reason.unwrap_or_else(|| "AI configuration error".to_string()));
                cx.notify();
                return;
            }
            AIConfigStatus::Ready => {
                // Continue to show dialog
            }
        }

        // Check analyze capability
        if !config.provider.capabilities().analyze {
            self.status_message = Some(format!(
                "{} provider does not support Analyze.",
                config.provider.name()
            ));
            cx.notify();
            return;
        }

        // Reset dialog state and set verb
        self.ask_ai.verb = crate::app::AiVerb::Analyze;
        self.ask_ai.reset();

        // Set mode first, then update context
        self.mode = Mode::AiDialog;

        // Build context from current selection
        self.ask_ai_update_context(cx);
    }

    pub fn hide_ask_ai(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.ask_ai.reset();
        cx.notify();
    }

    /// Update context based on mode and refresh the display
    pub fn ask_ai_update_context(&mut self, cx: &mut Context<Self>) {
        use crate::ai::{build_ai_context, find_current_region, find_used_range, range_ref};
        use crate::app::AskAIContextMode;

        let config = visigrid_config::ai::ResolvedAIConfig::load();
        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let sheet_name = sheet.name.clone();

        // Determine range based on context mode
        let (start_row, start_col, end_row, end_col) = match self.ask_ai.context_mode {
            AskAIContextMode::CurrentSelection => {
                let (sr, sc) = self.view_state.selected;
                let (er, ec) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
                (sr.min(er), sc.min(ec), sr.max(er), sc.max(ec))
            }
            AskAIContextMode::CurrentRegion => {
                let (row, col) = self.view_state.selected;
                find_current_region(sheet, row, col)
            }
            AskAIContextMode::EntireUsedRange => {
                find_used_range(sheet)
            }
        };

        // Store selected range
        self.ask_ai.selected_range = Some((start_row, start_col, end_row, end_col));

        // Build context to get summary and warnings
        let context = build_ai_context(
            sheet,
            &sheet_name,
            start_row,
            start_col,
            end_row,
            end_col,
            config.privacy_mode,
        );

        self.ask_ai.context_summary = format!(
            "{}!{}",
            sheet_name,
            range_ref(start_row, start_col, end_row, end_col)
        );
        self.ask_ai.context_top_left = Some(context.top_left);
        self.ask_ai.warnings = context.warnings.iter().map(|w| w.message()).collect();

        // Close context selector
        self.ask_ai.context_selector_open = false;
        cx.notify();
    }

    /// Set context mode and update
    pub fn ask_ai_set_context_mode(&mut self, mode: crate::app::AskAIContextMode, cx: &mut Context<Self>) {
        self.ask_ai.context_mode = mode;
        self.ask_ai.context_selector_open = false; // Close dropdown after selection
        self.ask_ai_update_context(cx);
    }

    /// Toggle context selector visibility
    pub fn ask_ai_toggle_context_selector(&mut self, cx: &mut Context<Self>) {
        self.ask_ai.context_selector_open = !self.ask_ai.context_selector_open;
        cx.notify();
    }

    /// Toggle sent panel visibility
    pub fn ask_ai_toggle_sent_panel(&mut self, cx: &mut Context<Self>) {
        self.ask_ai.sent_panel_expanded = !self.ask_ai.sent_panel_expanded;
        cx.notify();
    }

    /// Execute the AI query (branches on verb: InsertFormula or Analyze)
    pub fn ask_ai_submit(&mut self, cx: &mut Context<Self>) {
        use crate::ai::{ask_ai, analyze, build_ai_context};
        use crate::app::{AiVerb, AskAIStatus, AskAISentContext, AskAITruncation};

        // Prevent duplicate requests
        if self.ask_ai.is_loading() {
            return;
        }

        if self.ask_ai.question.trim().is_empty() {
            self.ask_ai.error = Some("Please enter a question".to_string());
            cx.notify();
            return;
        }

        let mut config = visigrid_config::ai::ResolvedAIConfig::load();

        // Use cached session key if keychain hasn't caught up yet
        if config.api_key.is_none() && self.ai_session_key.is_some() {
            config.api_key = self.ai_session_key.clone();
        }

        // Get range from stored selection or current view state
        let (start_row, start_col, end_row, end_col) = self.ask_ai.selected_range
            .unwrap_or_else(|| {
                let (sr, sc) = self.view_state.selected;
                let (er, ec) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
                (sr.min(er), sc.min(ec), sr.max(er), sc.max(ec))
            });

        let sheet_name = self.workbook.read(cx)
            .active_sheet()
            .name
            .clone();

        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();

        let context = build_ai_context(
            sheet,
            &sheet_name,
            start_row,
            start_col,
            end_row,
            end_col,
            config.privacy_mode,
        );

        // Clear previous response (keep question and context)
        self.ask_ai.clear_response();

        // Generate request ID
        let request_id = format!("{:x}", rand::random::<u64>());
        self.ask_ai.request_id = Some(request_id.clone());

        // Set loading state
        self.ask_ai.status = AskAIStatus::Loading;
        cx.notify();

        // Build sent context info for transparency
        let original_rows = (end_row.saturating_sub(start_row)) + 1;
        let original_cols = (end_col.saturating_sub(start_col)) + 1;
        let truncation = match (context.row_count < original_rows, context.col_count < original_cols) {
            (true, true) => AskAITruncation::Both,
            (true, false) => AskAITruncation::Rows,
            (false, true) => AskAITruncation::Cols,
            (false, false) => AskAITruncation::None,
        };

        let sent_context = AskAISentContext {
            provider: config.provider.name().to_string(),
            model: config.model.clone(),
            privacy_mode: config.privacy_mode,
            range_display: context.actual_range.clone(),
            rows_sent: context.row_count,
            cols_sent: context.col_count,
            total_cells: context.row_count * context.col_count,
            headers_included: context.headers.is_some(),
            truncation,
        };
        self.ask_ai.sent_context = Some(sent_context);

        // Clone for thread
        let question = self.ask_ai.question.clone();
        let verb = self.ask_ai.verb;

        // Branch on verb: InsertFormula vs Analyze
        match verb {
            AiVerb::InsertFormula => {
                let result = std::thread::spawn(move || {
                    ask_ai(&config, &question, &context)
                }).join();

                match result {
                    Ok(Ok(response)) => {
                        self.ask_ai.raw_response = response.raw_response.clone();
                        self.ask_ai.explanation = Some(response.explanation);
                        self.ask_ai.formula = response.formula.as_ref().map(|f| f.trim().to_string());
                        self.ask_ai.warnings.extend(response.warnings);

                        // Validate formula if present
                        if let Some(ref formula) = self.ask_ai.formula {
                            match self.validate_ai_formula(formula, cx) {
                                Ok(()) => {
                                    self.ask_ai.formula_valid = true;
                                    self.ask_ai.formula_error = None;
                                }
                                Err(e) => {
                                    self.ask_ai.formula_valid = false;
                                    self.ask_ai.formula_error = Some(e);
                                }
                            }
                        }

                        self.ask_ai.status = AskAIStatus::Success;
                    }
                    Ok(Err(e)) => {
                        let error_msg = self.format_ai_error(&e);
                        self.ask_ai.status = AskAIStatus::Error(error_msg.clone());
                        self.ask_ai.error = Some(error_msg);
                    }
                    Err(_) => {
                        self.ask_ai.status = AskAIStatus::Error("Request failed".to_string());
                        self.ask_ai.error = Some("AI request failed unexpectedly. Please retry.".to_string());
                    }
                }
            }
            AiVerb::Analyze => {
                let result = std::thread::spawn(move || {
                    analyze(&config, &question, &context)
                }).join();

                match result {
                    Ok(Ok(response)) => {
                        self.ask_ai.raw_response = response.raw_response.clone();
                        self.ask_ai.response_text = Some(response.analysis);
                        self.ask_ai.warnings.extend(response.warnings);
                        // No formula, no validation — read-only contract
                        self.ask_ai.status = AskAIStatus::Success;
                    }
                    Ok(Err(e)) => {
                        let error_msg = self.format_ai_error(&e);
                        self.ask_ai.status = AskAIStatus::Error(error_msg.clone());
                        self.ask_ai.error = Some(error_msg);
                    }
                    Err(_) => {
                        self.ask_ai.status = AskAIStatus::Error("Request failed".to_string());
                        self.ask_ai.error = Some("AI request failed unexpectedly. Please retry.".to_string());
                    }
                }
            }
        }

        cx.notify();
    }

    /// Format AI errors for user display
    fn format_ai_error(&self, error: &crate::ai::AskError) -> String {
        use crate::ai::AskError;
        match error {
            AskError::NotConfigured(_) => "AI not configured. Open AI Settings to set up.".to_string(),
            AskError::NotImplemented(provider) => format!("{} is not yet supported.", provider),
            AskError::MissingKey => "API key missing. Check AI Settings.".to_string(),
            AskError::NetworkError(e) => {
                if e.contains("timeout") || e.contains("timed out") {
                    "Request timed out. Please retry.".to_string()
                } else if e.contains("connection") {
                    "Network unavailable. Check your connection.".to_string()
                } else {
                    format!("Network error: {}", e)
                }
            }
            AskError::ApiError { status, message } => {
                if *status == 401 || *status == 403 {
                    "Authentication failed. Check your API key.".to_string()
                } else if *status == 429 {
                    "Rate limited. Please wait and retry.".to_string()
                } else {
                    format!("API error ({}): {}", status, message)
                }
            }
            AskError::ParseError(e) => format!("Failed to parse response: {}", e),
            AskError::InvalidResponse(e) => format!("Invalid response: {}", e),
        }
    }

    /// Validate that a formula is safe to insert
    fn validate_ai_formula(&self, formula: &str, _cx: &Context<Self>) -> Result<(), String> {
        use visigrid_engine::formula::parser::parse;

        // Canonicalize: trim whitespace
        let formula = formula.trim();

        // Must start with =
        if !formula.starts_with('=') {
            return Err("Formula must start with '='".to_string());
        }

        // Try to parse (skip the leading =)
        let formula_body = &formula[1..];
        match parse(formula_body) {
            Ok(_) => Ok(()),
            Err(e) => Err(format!("Parse error: {}", e)),
        }
    }

    /// Insert the AI-proposed formula into the active cell
    pub fn ask_ai_insert_formula(&mut self, cx: &mut Context<Self>) {
        use crate::ai::cell_ref;
        use crate::history::{MutationSource, AiMutationMeta};

        if !self.ask_ai.can_insert() {
            return;
        }

        let formula = match &self.ask_ai.formula {
            Some(f) => f.clone(),
            None => return,
        };

        // Get active cell
        let (row, col) = self.view_state.selected;

        // Get old value for history
        let old_value = self.workbook.read(cx)
            .active_sheet()
            .get_raw(row, col);

        // Build AI provenance metadata from sent context
        let source = if let Some(sent) = &self.ask_ai.sent_context {
            MutationSource::Ai(AiMutationMeta {
                provider: sent.provider.clone(),
                model: sent.model.clone(),
                privacy_mode: sent.privacy_mode,
                request_id: self.ask_ai.request_id.clone(),
                context_mode: match self.ask_ai.context_mode {
                    crate::app::AskAIContextMode::CurrentSelection => "selection".to_string(),
                    crate::app::AskAIContextMode::CurrentRegion => "region".to_string(),
                    crate::app::AskAIContextMode::EntireUsedRange => "used_range".to_string(),
                },
                truncation: sent.truncation.label().to_string(),
            })
        } else {
            // Fallback if no sent context (shouldn't happen in normal flow)
            MutationSource::Human
        };

        // Record change in history with AI source
        let sheet_idx = self.sheet_index(cx);
        self.history.record_change_with_source(sheet_idx, row, col, old_value, formula.clone(), source);

        // Set the cell value using the standard helper
        self.set_cell_value(row, col, &formula, cx);

        // Show confirmation (keep dialog open)
        let cell_addr = cell_ref(row, col);
        self.ask_ai.last_insertion = Some(format!("Inserted into {}", cell_addr));
        self.ask_ai.inserted = true;

        // Update status bar too
        self.status_message = Some(format!("Formula inserted at {}", cell_addr));

        cx.notify();
    }

    /// Retry the last AI request
    pub fn ask_ai_retry(&mut self, cx: &mut Context<Self>) {
        if !self.ask_ai.can_retry() {
            return;
        }
        self.ask_ai_submit(cx);
    }

    /// Focus the question input for refinement
    pub fn ask_ai_refine(&mut self, cx: &mut Context<Self>) {
        // Clear previous response but keep question
        self.ask_ai.clear_response();
        // The UI will show the question input focused
        cx.notify();
    }

    /// Copy diagnostic details for error reporting
    pub fn ask_ai_copy_details(&mut self, cx: &mut Context<Self>) {
        let mut details = String::new();

        let verb_label = match self.ask_ai.verb {
            crate::app::AiVerb::InsertFormula => "Insert Formula",
            crate::app::AiVerb::Analyze => "Analyze",
        };
        details.push_str(&format!("=== {} AI Diagnostic Details ===\n\n", verb_label));

        let (contract_id, contract_label, write_scope) = match self.ask_ai.verb {
            crate::app::AiVerb::InsertFormula => (
                crate::ai::INSERT_FORMULA_CONTRACT, "Single-cell write", "Active cell",
            ),
            crate::app::AiVerb::Analyze => (
                crate::ai::ANALYZE_CONTRACT, "Read-only", "None",
            ),
        };
        details.push_str(&format!("Contract: {} ({})\n", contract_id, contract_label));
        details.push_str(&format!("Write scope: {}\n", write_scope));

        if let Some(sent) = &self.ask_ai.sent_context {
            details.push_str(&format!("Provider: {}\n", sent.provider));
            details.push_str(&format!("Model: {}\n", sent.model));
            details.push_str(&format!("Privacy mode: {}\n", sent.privacy_mode));
            details.push_str(&format!("Range: {}\n", sent.range_display));
            details.push_str(&format!("Cells sent: {} ({} rows x {} cols)\n",
                sent.total_cells, sent.rows_sent, sent.cols_sent));
            details.push_str(&format!("Truncation: {}\n", sent.truncation.label()));
            details.push('\n');
        }

        if let Some(err) = &self.ask_ai.error {
            details.push_str(&format!("Error: {}\n\n", err));
        }

        if let Some(raw) = &self.ask_ai.raw_response {
            details.push_str("Raw response:\n");
            details.push_str(raw);
            details.push('\n');
        }

        cx.write_to_clipboard(gpui::ClipboardItem::new_string(details));
        self.status_message = Some("Diagnostic details copied".to_string());
        cx.notify();
    }

    /// Type a character in the Ask AI question field
    pub fn ask_ai_type_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.ask_ai.question.push(c);
        self.ask_ai.error = None;
        cx.notify();
    }

    /// Backspace in the Ask AI question field
    pub fn ask_ai_backspace(&mut self, cx: &mut Context<Self>) {
        self.ask_ai.question.pop();
        cx.notify();
    }

    /// Paste into the Ask AI question field
    pub fn ask_ai_paste(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = cx.read_from_clipboard() {
            if let Some(text) = item.text() {
                // Filter out control characters but keep newlines for multi-line questions
                let clean_text: String = text.chars()
                    .filter(|c| !c.is_control() || *c == '\n')
                    .collect();
                self.ask_ai.question.push_str(&clean_text);
                self.ask_ai.error = None;
                cx.notify();
            }
        }
    }

    // =========================================================================
    // Explain Differences dialog (Phase 3)
    // =========================================================================

    /// Show the Explain Differences dialog for changes since a history entry
    pub fn show_explain_diff(&mut self, entry_id: u64, cx: &mut Context<Self>) {
        use crate::diff::build_diff_since;

        // Build the diff report
        if let Some(report) = build_diff_since(&self.history, entry_id) {
            self.diff_report = Some(report);
            self.diff_ai_only_filter = false;
            self.history_context_menu_entry_id = None;
            self.mode = Mode::ExplainDiff;
            cx.notify();
        }
    }

    /// Close the Explain Differences dialog
    pub fn close_explain_diff(&mut self, cx: &mut Context<Self>) {
        self.diff_report = None;
        self.diff_ai_only_filter = false;
        self.diff_selected_entry = None;
        self.diff_ai_summary = None;
        self.diff_ai_summary_loading = false;
        self.diff_ai_summary_error = None;
        self.diff_entry_explanations.clear();
        self.diff_explaining_entry = None;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Toggle the AI-only filter in Explain Differences
    pub fn toggle_diff_ai_filter(&mut self, cx: &mut Context<Self>) {
        self.diff_ai_only_filter = !self.diff_ai_only_filter;
        cx.notify();
    }

    /// Jump to a cell from the diff report (keeps dialog open for audit mode)
    pub fn diff_jump_to_cell(&mut self, sheet_index: usize, row: usize, col: usize, cx: &mut Context<Self>) {
        // Switch to sheet if needed
        if sheet_index != self.sheet_index(cx) {
            self.wb_mut(cx, |wb| wb.set_active_sheet(sheet_index));
            self.update_cached_sheet_id(cx);
        }
        // Select the cell
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        // Track selected entry for highlighting
        self.diff_selected_entry = Some((sheet_index, row, col));
        // Scroll to make visible
        self.ensure_cell_visible(row, col);
        cx.notify();
    }

    /// Jump to a cell and close the dialog (Enter key behavior)
    pub fn diff_jump_and_close(&mut self, cx: &mut Context<Self>) {
        if let Some((sheet_index, row, col)) = self.diff_selected_entry {
            // Switch to sheet if needed
            if sheet_index != self.sheet_index(cx) {
                self.wb_mut(cx, |wb| wb.set_active_sheet(sheet_index));
                self.update_cached_sheet_id(cx);
            }
            // Select the cell
            self.view_state.selected = (row, col);
            self.view_state.selection_end = None;
            // Scroll to make visible
            self.ensure_cell_visible(row, col);
        }
        // Close the dialog
        self.close_explain_diff(cx);
    }

    /// Copy diff report as text to clipboard
    pub fn copy_diff_report(&mut self, cx: &mut Context<Self>) {
        let report = match &self.diff_report {
            Some(r) => r,
            None => return,
        };

        // Get sheet names
        let sheet_names: Vec<String> = self.wb(cx).sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        let mut text = String::new();

        // Header
        text.push_str(&format!("Changes since: \"{}\"\n", report.since_entry_label));
        text.push_str(&format!("Actions spanned: {}\n\n", report.entries_spanned));

        // Summary
        text.push_str("Summary:\n");
        text.push_str(&format!("  {} value changes\n", report.value_changes.len()));
        text.push_str(&format!("  {} formula changes\n", report.formula_changes.len()));
        text.push_str(&format!("  {} structural changes\n", report.structural_changes.len()));
        text.push_str(&format!("  {} named range changes\n", report.named_range_changes.len()));
        text.push_str(&format!("  {} validation changes\n", report.validation_changes.len()));
        text.push_str(&format!("  {} format changes\n\n", report.format_change_count));

        // Value changes
        if !report.value_changes.is_empty() {
            text.push_str("Values:\n");
            for entry in &report.value_changes {
                let addr = entry.full_address(&sheet_names);
                let ai_tag = if entry.ai_touched { " [AI]" } else { "" };
                let old_val = if entry.old_value.is_empty() { "(empty)" } else { &entry.old_value };
                let new_val = if entry.new_value.is_empty() { "(empty)" } else { &entry.new_value };
                text.push_str(&format!("  {} : {} → {}{}\n", addr, old_val, new_val, ai_tag));
            }
            text.push('\n');
        }

        // Formula changes
        if !report.formula_changes.is_empty() {
            text.push_str("Formulas:\n");
            for entry in &report.formula_changes {
                let addr = entry.full_address(&sheet_names);
                let ai_tag = if entry.ai_touched { " [AI]" } else { "" };
                let old_val = if entry.old_value.is_empty() { "(empty)" } else { &entry.old_value };
                let new_val = if entry.new_value.is_empty() { "(empty)" } else { &entry.new_value };
                text.push_str(&format!("  {} : {} → {}{}\n", addr, old_val, new_val, ai_tag));
            }
            text.push('\n');
        }

        // Structural changes
        if !report.structural_changes.is_empty() {
            text.push_str("Structural:\n");
            for change in &report.structural_changes {
                text.push_str(&format!("  {}\n", change.description()));
            }
            text.push('\n');
        }

        // Named range changes
        if !report.named_range_changes.is_empty() {
            text.push_str("Named Ranges:\n");
            for change in &report.named_range_changes {
                text.push_str(&format!("  {}\n", change.description()));
            }
            text.push('\n');
        }

        // Validation changes
        if !report.validation_changes.is_empty() {
            text.push_str("Validation:\n");
            for change in &report.validation_changes {
                text.push_str(&format!("  {}\n", change.description()));
            }
            text.push('\n');
        }

        // Copy to clipboard
        cx.write_to_clipboard(gpui::ClipboardItem::new_string(text));
    }

    /// Generate AI summary of the diff report
    pub fn generate_diff_summary(&mut self, cx: &mut Context<Self>) {
        use visigrid_config::ai::ResolvedAIConfig;

        // Load AI config
        let config = ResolvedAIConfig::load();

        // Check if analyze capability is available (diff explain is read-only)
        if !config.provider.capabilities().analyze {
            self.diff_ai_summary_error = Some("AI provider does not support summaries".to_string());
            cx.notify();
            return;
        }

        let report = match &self.diff_report {
            Some(r) => r.clone(),
            None => return,
        };

        // Get sheet names
        let sheet_names: Vec<String> = self.wb(cx).sheet_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        // Build prompt for AI
        let prompt = build_diff_summary_prompt(&report, &sheet_names);

        // Set loading state
        self.diff_ai_summary_loading = true;
        self.diff_ai_summary_error = None;
        cx.notify();

        // Spawn background task
        cx.spawn({
            async move |this, cx| {
                let result = std::thread::spawn(move || {
                    call_diff_summary_ai(&config, &prompt)
                }).join().unwrap_or_else(|_| Err("AI thread panicked".to_string()));

                let _ = this.update(cx, |this, cx| {
                    this.diff_ai_summary_loading = false;
                    match result {
                        Ok(summary) => {
                            this.diff_ai_summary = Some(summary);
                            this.diff_ai_summary_error = None;
                        }
                        Err(e) => {
                            this.diff_ai_summary_error = Some(e);
                        }
                    }
                    cx.notify();
                });
            }
        }).detach();
    }

    /// Copy the AI summary to clipboard
    pub fn copy_diff_summary(&mut self, cx: &mut Context<Self>) {
        if let Some(summary) = &self.diff_ai_summary {
            cx.write_to_clipboard(gpui::ClipboardItem::new_string(summary.clone()));
        }
    }

    /// Explain a single diff entry (AI-powered)
    pub fn explain_diff_entry(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        old_value: String,
        new_value: String,
        ai_touched: bool,
        ai_source: Option<String>,
        cx: &mut Context<Self>,
    ) {
        use visigrid_config::ai::ResolvedAIConfig;

        let key = (sheet_index, row, col);

        // Check cache first
        if self.diff_entry_explanations.contains_key(&key) {
            return; // Already explained
        }

        // Load AI config
        let config = ResolvedAIConfig::load();

        // Check if analyze capability is available (diff explain is read-only)
        if !config.provider.capabilities().analyze {
            return;
        }

        // Get sheet name
        let sheet_name = self.wb(cx).sheet_names()
            .get(sheet_index)
            .map(|s| s.to_string())
            .unwrap_or_else(|| format!("Sheet{}", sheet_index + 1));

        // Build cell address
        let col_letter = if col < 26 {
            ((b'A' + col as u8) as char).to_string()
        } else {
            let first = (b'A' + (col / 26 - 1) as u8) as char;
            let second = (b'A' + (col % 26) as u8) as char;
            format!("{}{}", first, second)
        };
        let cell_addr = format!("{}!{}{}", sheet_name, col_letter, row + 1);

        // Build prompt
        let prompt = build_entry_explanation_prompt(
            &cell_addr,
            &old_value,
            &new_value,
            ai_touched,
            ai_source.as_deref(),
        );

        // Set loading state
        self.diff_explaining_entry = Some(key);
        cx.notify();

        // Spawn background task
        cx.spawn({
            async move |this, cx| {
                let result = std::thread::spawn(move || {
                    call_entry_explanation_ai(&config, &prompt)
                }).join().unwrap_or_else(|_| Err("AI thread panicked".to_string()));

                let _ = this.update(cx, |this, cx| {
                    this.diff_explaining_entry = None;
                    if let Ok(explanation) = result {
                        this.diff_entry_explanations.insert(key, explanation);
                    }
                    cx.notify();
                });
            }
        }).detach();
    }

    /// Show context menu for a history entry
    pub fn show_history_context_menu(&mut self, entry_id: u64, cx: &mut Context<Self>) {
        self.history_context_menu_entry_id = Some(entry_id);
        cx.notify();
    }

    /// Hide history context menu
    pub fn hide_history_context_menu(&mut self, cx: &mut Context<Self>) {
        self.history_context_menu_entry_id = None;
        cx.notify();
    }
}

// ============================================================================
// AI Diff Summary helpers
// ============================================================================

/// Build prompt for AI diff summary
fn build_diff_summary_prompt(report: &crate::diff::DiffReport, sheet_names: &[String]) -> String {
    let mut prompt = String::new();

    prompt.push_str("Summarize the following spreadsheet changes in 4-8 sentences. ");
    prompt.push_str("Focus on what was changed and why it might matter. ");
    prompt.push_str("Do not suggest any edits or formulas. Just describe what happened.\n\n");

    prompt.push_str(&format!("Changes since: \"{}\"\n", report.since_entry_label));
    prompt.push_str(&format!("Actions: {}\n\n", report.entries_spanned));

    // Summary stats
    prompt.push_str("Summary:\n");
    prompt.push_str(&format!("- {} value changes\n", report.value_changes.len()));
    prompt.push_str(&format!("- {} formula changes\n", report.formula_changes.len()));
    prompt.push_str(&format!("- {} structural changes\n", report.structural_changes.len()));
    prompt.push_str(&format!("- {} named range changes\n", report.named_range_changes.len()));
    prompt.push_str(&format!("- {} validation changes\n", report.validation_changes.len()));
    prompt.push_str(&format!("- {} format changes\n\n", report.format_change_count));

    // Sample of changes (limit to 20 to avoid token overflow)
    let mut change_count = 0;
    const MAX_CHANGES: usize = 20;

    if !report.value_changes.is_empty() && change_count < MAX_CHANGES {
        prompt.push_str("Value changes:\n");
        for entry in report.value_changes.iter().take(MAX_CHANGES - change_count) {
            let addr = entry.full_address(sheet_names);
            let ai_tag = if entry.ai_touched { " [AI]" } else { "" };
            // Truncate long values
            let old_val = truncate_value(&entry.old_value, 30);
            let new_val = truncate_value(&entry.new_value, 30);
            prompt.push_str(&format!("  {} : {} → {}{}\n", addr, old_val, new_val, ai_tag));
            change_count += 1;
        }
        prompt.push('\n');
    }

    if !report.formula_changes.is_empty() && change_count < MAX_CHANGES {
        prompt.push_str("Formula changes:\n");
        for entry in report.formula_changes.iter().take(MAX_CHANGES - change_count) {
            let addr = entry.full_address(sheet_names);
            let ai_tag = if entry.ai_touched { " [AI]" } else { "" };
            let old_val = truncate_value(&entry.old_value, 50);
            let new_val = truncate_value(&entry.new_value, 50);
            prompt.push_str(&format!("  {} : {} → {}{}\n", addr, old_val, new_val, ai_tag));
            change_count += 1;
        }
        prompt.push('\n');
    }

    if !report.structural_changes.is_empty() {
        prompt.push_str("Structural changes:\n");
        for change in report.structural_changes.iter().take(5) {
            prompt.push_str(&format!("  {}\n", change.description()));
        }
        prompt.push('\n');
    }

    if !report.named_range_changes.is_empty() {
        prompt.push_str("Named range changes:\n");
        for change in report.named_range_changes.iter().take(5) {
            prompt.push_str(&format!("  {}\n", change.description()));
        }
        prompt.push('\n');
    }

    prompt.push_str("\nProvide a concise summary (4-8 sentences). No suggestions, just description.");

    prompt
}

/// Truncate value for prompt (avoid token overflow)
fn truncate_value(val: &str, max_len: usize) -> String {
    if val.is_empty() {
        "(empty)".to_string()
    } else if val.len() <= max_len {
        val.to_string()
    } else {
        format!("{}...", &val[..max_len])
    }
}

/// Call AI API for diff summary
fn call_diff_summary_ai(config: &visigrid_config::ai::ResolvedAIConfig, prompt: &str) -> Result<String, String> {
    use visigrid_config::settings::AIProvider;

    let api_url = config.provider.chat_completions_url()
        .ok_or_else(|| format!("{} not implemented for summaries", config.provider.name()))?;

    let api_key = config.api_key.as_ref().ok_or("API key not configured")?;

    // Build request (simpler than Ask AI - just want plain text response)
    let request = serde_json::json!({
        "model": config.model,
        "messages": [
            {
                "role": "system",
                "content": "You are a spreadsheet assistant. Summarize changes concisely. No suggestions or formulas - just describe what changed."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.3,
        "max_tokens": 500
    });

    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .map_err(|e| e.to_string())?;

    let response = client
        .post(api_url)
        .header("Authorization", format!("Bearer {}", api_key))
        .header("Content-Type", "application/json")
        .json(&request)
        .send()
        .map_err(|e| format!("Network error: {}", e))?;

    let status = response.status();
    if !status.is_success() {
        let error_text = response.text().unwrap_or_default();
        return Err(format!("API error ({}): {}", status.as_u16(), error_text));
    }

    let body: serde_json::Value = response.json().map_err(|e| format!("Parse error: {}", e))?;

    // Extract content from response
    let content = body["choices"][0]["message"]["content"]
        .as_str()
        .ok_or("Invalid response format")?
        .trim()
        .to_string();

    Ok(content)
}

// ============================================================================
// AI Entry Explanation helpers
// ============================================================================

/// Build prompt for single entry explanation
fn build_entry_explanation_prompt(
    cell_addr: &str,
    old_value: &str,
    new_value: &str,
    ai_touched: bool,
    ai_source: Option<&str>,
) -> String {
    let mut prompt = String::new();

    prompt.push_str("Explain this spreadsheet cell change in 2-4 sentences. ");
    prompt.push_str("Just describe what happened. No suggestions, no 'you should', no edits.\n\n");

    prompt.push_str(&format!("Cell: {}\n", cell_addr));

    // Truncate values to avoid token overflow
    let old_display = if old_value.is_empty() {
        "(empty)".to_string()
    } else if old_value.len() > 100 {
        format!("{}...", &old_value[..100])
    } else {
        old_value.to_string()
    };

    let new_display = if new_value.is_empty() {
        "(empty)".to_string()
    } else if new_value.len() > 100 {
        format!("{}...", &new_value[..100])
    } else {
        new_value.to_string()
    };

    prompt.push_str(&format!("Before: {}\n", old_display));
    prompt.push_str(&format!("After: {}\n", new_display));

    // Source info
    if ai_touched {
        if let Some(source) = ai_source {
            prompt.push_str(&format!("Source: AI ({})\n", source));
        } else {
            prompt.push_str("Source: AI\n");
        }
    } else {
        prompt.push_str("Source: Manual edit\n");
    }

    prompt.push_str("\nExplain concisely what this change does. 2-4 sentences max.");

    prompt
}

/// Call AI API for single entry explanation
fn call_entry_explanation_ai(
    config: &visigrid_config::ai::ResolvedAIConfig,
    prompt: &str,
) -> Result<String, String> {
    use visigrid_config::settings::AIProvider;

    let api_url = config.provider.chat_completions_url()
        .ok_or_else(|| format!("{} not implemented for explanations", config.provider.name()))?;

    let api_key = config.api_key.as_ref().ok_or("API key not configured")?;

    let request = serde_json::json!({
        "model": config.model,
        "messages": [
            {
                "role": "system",
                "content": "You are a spreadsheet assistant. Explain cell changes concisely. No suggestions or edits - just describe what happened."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.3,
        "max_tokens": 200
    });

    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .map_err(|e| e.to_string())?;

    let response = client
        .post(api_url)
        .header("Authorization", format!("Bearer {}", api_key))
        .header("Content-Type", "application/json")
        .json(&request)
        .send()
        .map_err(|e| format!("Network error: {}", e))?;

    let status = response.status();
    if !status.is_success() {
        let error_text = response.text().unwrap_or_default();
        return Err(format!("API error ({}): {}", status.as_u16(), error_text));
    }

    let body: serde_json::Value = response.json().map_err(|e| format!("Parse error: {}", e))?;

    let content = body["choices"][0]["message"]["content"]
        .as_str()
        .ok_or("Invalid response format")?
        .trim()
        .to_string();

    Ok(content)
}
