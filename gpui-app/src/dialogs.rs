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
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
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
        // Close validation dropdown when jumping to a cell
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::SelectionChanged,
            cx,
        );

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
        let existing_rule = self.sheet().validations.get(row, col).cloned();
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
                let sheet_index = self.sheet_index();

                // Capture rules to be cleared for undo
                let cleared_rules: Vec<(CellRange, ValidationRule)> = self.sheet()
                    .validations
                    .iter()
                    .filter(|(r, _)| r.overlaps(&range))
                    .map(|(r, v)| (*r, v.clone()))
                    .collect();

                // Clear the rules
                self.sheet_mut().validations.clear_range(&range);
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
            let sheet_index = self.sheet_index();

            // Replace-in-range semantics: capture overlapping rules before clearing
            let previous_rules: Vec<(CellRange, ValidationRule)> = self.sheet()
                .validations
                .iter()
                .filter(|(r, _)| r.overlaps(&range))
                .map(|(r, v)| (*r, v.clone()))
                .collect();

            // Clear overlapping rules, then set new rule
            self.sheet_mut().validations.clear_range(&range);
            self.sheet_mut().validations.set(range.clone(), rule.clone());
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
            let invalid_count = self.validate_and_mark_range(&range);

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
    fn validate_and_mark_range(&mut self, range: &visigrid_engine::validation::CellRange) -> usize {
        use visigrid_engine::validation::ValidationResult;
        use visigrid_engine::workbook::Workbook;

        let mut invalid_count = 0;
        let sheet_index = self.sheet_index();

        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                let display_value = self.sheet().get_display(row, col);
                // Skip empty cells if ignore_blank is true (handled by validation)
                if display_value.is_empty() {
                    // Clear any existing invalid marker
                    self.invalid_cells.remove(&(row, col));
                    continue;
                }
                let result = self.workbook.validate_cell_input(sheet_index, row, col, &display_value);
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
            let sheet_index = self.sheet_index();

            // Capture rules to be cleared for undo
            let cleared_rules: Vec<(CellRange, visigrid_engine::validation::ValidationRule)> = self.sheet()
                .validations
                .iter()
                .filter(|(r, _)| r.overlaps(&range))
                .map(|(r, v)| (*r, v.clone()))
                .collect();

            // Clear the rules
            self.sheet_mut().validations.clear_range(&range);
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

        let sheet_index = self.sheet_index();

        // Add the exclusion
        self.sheet_mut().validations.exclude(range);
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

        let sheet_index = self.sheet_index();

        // Capture exclusions that will be cleared (for undo)
        let cleared_exclusions: Vec<CellRange> = self.sheet()
            .validations
            .exclusions_in_range(&range);

        if cleared_exclusions.is_empty() {
            self.status_message = Some("No exclusions to clear in selection".to_string());
            cx.notify();
            return;
        }

        // Clear the exclusions
        self.sheet_mut().validations.clear_exclusions_in_range(&range);
        self.bump_cells_rev();

        // Revalidate the range (cells may now be validated again)
        let invalid_count = self.validate_and_mark_range(&range);

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
}
