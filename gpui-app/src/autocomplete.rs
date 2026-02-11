//! Formula autocomplete and signature help
//!
//! This module contains autocomplete suggestions, signature help,
//! and formula error detection for the formula editor.

use gpui::*;

use crate::app::Spreadsheet;
use crate::formula_context;
use crate::mode::Mode;

/// An autocomplete entry — either a built-in function or a custom (user-defined) one.
#[derive(Debug, Clone)]
pub enum AutocompleteEntry {
    BuiltIn(&'static formula_context::FunctionInfo),
    Custom { name: String },
}

impl AutocompleteEntry {
    pub fn name(&self) -> &str {
        match self {
            AutocompleteEntry::BuiltIn(f) => f.name,
            AutocompleteEntry::Custom { name } => name,
        }
    }

    pub fn is_custom(&self) -> bool {
        matches!(self, AutocompleteEntry::Custom { .. })
    }
}

/// Signature help context for rendering
pub struct SignatureHelpInfo {
    pub function: &'static formula_context::FunctionInfo,
    pub current_arg: usize,
}

/// Error info for the error banner
pub struct FormulaErrorInfo {
    pub message: String,
}

impl Spreadsheet {
    // ========================================================================
    // Formula Autocomplete
    // ========================================================================

    /// Get filtered autocomplete suggestions based on current edit value.
    /// Returns both built-in and custom function entries.
    pub fn autocomplete_suggestions(&self) -> Vec<AutocompleteEntry> {
        // Only show autocomplete for formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            return Vec::new();
        }

        let ctx = formula_context::analyze(&self.edit_value, self.edit_cursor);

        // Check mode and identifier length
        let prefix = match ctx.mode {
            formula_context::FormulaEditMode::Start
            | formula_context::FormulaEditMode::Operator
            | formula_context::FormulaEditMode::ArgList => {
                Some("")
            }
            formula_context::FormulaEditMode::Identifier => {
                if let Some(ref id_text) = ctx.identifier_text {
                    if id_text.len() >= 2 {
                        Some(id_text.as_str())
                    } else {
                        None
                    }
                } else {
                    None
                }
            }
            _ => None,
        };

        let Some(prefix) = prefix else {
            return Vec::new();
        };

        // Built-in functions
        let mut entries: Vec<AutocompleteEntry> = formula_context::get_functions_by_prefix(prefix)
            .into_iter()
            .map(AutocompleteEntry::BuiltIn)
            .collect();

        // Custom functions from registry
        let upper = prefix.to_ascii_uppercase();
        for name in self.custom_fn_registry.functions.keys() {
            if name.starts_with(&upper) {
                entries.push(AutocompleteEntry::Custom { name: name.clone() });
            }
        }

        // Sort all entries by name for stable ordering
        entries.sort_by(|a, b| a.name().cmp(b.name()));

        entries
    }

    /// Update autocomplete state based on current context
    pub fn update_autocomplete(&mut self, cx: &mut Context<Self>) {
        // Only in formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            self.autocomplete_visible = false;
            return;
        }

        // Don't reopen autocomplete if suppressed (user is navigating refs)
        if self.autocomplete_suppressed {
            self.autocomplete_visible = false;
            return;
        }

        let ctx = formula_context::analyze(&self.edit_value, self.edit_cursor);
        let suggestions = self.autocomplete_suggestions();

        if suggestions.is_empty() {
            self.autocomplete_visible = false;
            self.autocomplete_selected = 0;
        } else {
            self.autocomplete_visible = true;
            self.autocomplete_replace_range = ctx.replace_range.clone();
            // Clamp selected index
            if self.autocomplete_selected >= suggestions.len() {
                self.autocomplete_selected = 0;
            }
        }
        cx.notify();
    }

    /// Move autocomplete selection up
    pub fn autocomplete_up(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }
        let suggestions = self.autocomplete_suggestions();
        if suggestions.is_empty() {
            return;
        }
        if self.autocomplete_selected == 0 {
            self.autocomplete_selected = suggestions.len().saturating_sub(1);
        } else {
            self.autocomplete_selected -= 1;
        }
        cx.notify();
    }

    /// Move autocomplete selection down
    pub fn autocomplete_down(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }
        let suggestions = self.autocomplete_suggestions();
        if suggestions.is_empty() {
            return;
        }
        self.autocomplete_selected = (self.autocomplete_selected + 1) % suggestions.len();
        cx.notify();
    }

    /// Accept the selected autocomplete suggestion
    pub fn autocomplete_accept(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }

        let suggestions = self.autocomplete_suggestions();
        if suggestions.is_empty() || self.autocomplete_selected >= suggestions.len() {
            self.autocomplete_visible = false;
            return;
        }

        let entry = &suggestions[self.autocomplete_selected];
        let func_name = entry.name();

        // Build replacement text: function name + opening paren
        let replacement = format!("{}(", func_name);

        // Replace the identifier at replace_range
        let range = self.autocomplete_replace_range.clone();

        // Convert char positions to byte positions
        // Note: when position is at or past the end, use string length (for insertion at end)
        let char_count = self.edit_value.chars().count();
        let start_byte = if range.start >= char_count {
            self.edit_value.len()
        } else {
            self.edit_value.char_indices()
                .nth(range.start)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len())
        };
        let end_byte = if range.end >= char_count {
            self.edit_value.len()
        } else {
            self.edit_value.char_indices()
                .nth(range.end)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len())
        };

        self.edit_value.replace_range(start_byte..end_byte, &replacement);
        self.edit_cursor = range.start + replacement.chars().count();

        // Close autocomplete
        self.autocomplete_visible = false;
        self.autocomplete_selected = 0;

        // Enter formula mode if not already
        if !self.mode.is_formula() {
            self.mode = Mode::Formula;
        }

        cx.notify();
    }

    /// Dismiss autocomplete without accepting
    pub fn autocomplete_dismiss(&mut self, cx: &mut Context<Self>) {
        if self.autocomplete_visible {
            self.autocomplete_visible = false;
            self.autocomplete_selected = 0;
            cx.notify();
        }
    }

    // ========================================================================
    // Formula Signature Help
    // ========================================================================

    /// Get signature help info if cursor is inside a function call
    pub fn signature_help(&self) -> Option<SignatureHelpInfo> {
        // Only show for formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            return None;
        }

        // Don't show when navigating on a cross-sheet for ref picking
        if self.formula_ref_sheet.is_some() {
            return None;
        }

        // Don't show signature help when autocomplete is visible
        if self.autocomplete_visible {
            return None;
        }

        let ctx = formula_context::analyze(&self.edit_value, self.edit_cursor);

        // Only show in ArgList mode
        if !matches!(ctx.mode, formula_context::FormulaEditMode::ArgList) {
            return None;
        }

        // Get the current function
        ctx.current_function.map(|func| {
            SignatureHelpInfo {
                function: func,
                current_arg: ctx.current_arg_index.unwrap_or(0),
            }
        })
    }

    /// Get formula error to display (only Hard errors)
    pub fn formula_error(&self) -> Option<FormulaErrorInfo> {
        use formula_context::{check_errors, DiagnosticKind};

        // Only check for formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            return None;
        }

        // While editing, only show Hard errors (unknown function, invalid token)
        // Transient errors (missing paren, trailing operator) are hidden - we'll auto-fix on confirm
        let custom_names: Vec<&str> = self.custom_fn_registry.functions.keys()
            .map(|s| s.as_str()).collect();
        check_errors(&self.edit_value, self.edit_cursor, &custom_names)
            .filter(|diag| matches!(diag.kind, DiagnosticKind::Hard))
            .map(|diag| FormulaErrorInfo {
                message: diag.message,
            })
    }
}
