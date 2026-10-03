//! Formula autocomplete and signature help
//!
//! This module contains autocomplete suggestions, signature help,
//! and formula error detection for the formula editor.

use gpui::*;

use crate::app::Spreadsheet;
use crate::formula_context;
use crate::mode::Mode;

/// An autocomplete entry for a function, Table, column or section selector.
#[derive(Debug, Clone)]
pub enum AutocompleteEntry {
    BuiltIn(&'static formula_context::FunctionInfo),
    Custom { name: String },
    Table(crate::table_formula_editor::Suggestion),
}

impl AutocompleteEntry {
    pub fn name(&self) -> &str {
        match self {
            AutocompleteEntry::BuiltIn(f) => f.name,
            AutocompleteEntry::Custom { name } => name,
            AutocompleteEntry::Table(s) => &s.label,
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
    /// Returns function and structured-reference entries for the current caret context.
    pub fn autocomplete_suggestions(&self, cx: &App) -> Vec<AutocompleteEntry> {
        // Only show autocomplete for formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            return Vec::new();
        }

        let cursor = self.edit_value[..self.edit_cursor.min(self.edit_value.len())].chars().count();
        let ctx = formula_context::analyze(&self.edit_value, cursor);

        let (home, cell) = self.table_formula_context(cx);
        let table_entries = crate::table_formula_editor::suggestions(&self.edit_value, self.edit_cursor, self.wb(cx), home, cell);
        if ctx.token_at_cursor.as_ref().is_some_and(|t| t.token_type == formula_context::TokenType::StructuredRef) {
            return table_entries.into_iter().map(AutocompleteEntry::Table).collect();
        }

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
            return table_entries.into_iter().map(AutocompleteEntry::Table).collect();
        };

        // Built-in functions
        let mut entries: Vec<AutocompleteEntry> = formula_context::get_functions_by_prefix(prefix)
            .into_iter()
            .map(AutocompleteEntry::BuiltIn)
            .collect();

        // Custom functions from registry
        let upper = prefix.to_ascii_uppercase();
        for name in crate::scripting::lua_formulas::function_names() {
            if name.starts_with(&upper) {
                entries.push(AutocompleteEntry::Custom { name });
            }
        }

        entries.extend(table_entries.into_iter().map(AutocompleteEntry::Table));

        // Sort all entries by name for stable ordering
        entries.sort_by(|a, b| a.name().cmp(b.name()));

        entries
    }

    /// Update autocomplete state based on current context
    pub fn update_autocomplete(&mut self, cx: &mut Context<Self>) {
        // Only in formula mode
        if !self.mode.is_formula() && !self.edit_value.starts_with('=') {
            self.autocomplete_visible = false;
            self.update_formula_refs(cx);
            return;
        }

        // Don't reopen autocomplete if suppressed (user is navigating refs)
        if self.autocomplete_suppressed {
            self.autocomplete_visible = false;
            self.update_formula_refs(cx);
            return;
        }

        let cursor = self.edit_value[..self.edit_cursor.min(self.edit_value.len())].chars().count();
        let ctx = formula_context::analyze(&self.edit_value, cursor);
        let suggestions = self.autocomplete_suggestions(cx);

        if suggestions.is_empty() {
            self.autocomplete_visible = false;
            self.autocomplete_selected = 0;
        } else {
            self.autocomplete_visible = true;
            self.autocomplete_replace_range = formula_context::char_index_to_byte_offset(&self.edit_value, ctx.replace_range.start)..formula_context::char_index_to_byte_offset(&self.edit_value, ctx.replace_range.end);
            // Clamp selected index
            if self.autocomplete_selected >= suggestions.len() {
                self.autocomplete_selected = 0;
            }
        }
        self.update_formula_refs(cx);
        cx.notify();
    }

    /// Keep an already-open list aligned with caret movement without reopening it.
    pub(crate) fn refresh_autocomplete_at_caret(&mut self, cx: &mut Context<Self>) {
        if self.autocomplete_visible {
            if self.edit_selection_anchor.is_some_and(|anchor| anchor != self.edit_cursor) {
                self.autocomplete_dismiss(cx);
            } else {
                self.autocomplete_selected = 0;
                self.update_autocomplete(cx);
            }
        }
    }

    /// Move autocomplete selection up
    pub fn autocomplete_up(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }
        let suggestions = self.autocomplete_suggestions(cx);
        if suggestions.is_empty() {
            return;
        }
        if self.autocomplete_selected == 0 {
            self.autocomplete_selected = suggestions.len().saturating_sub(1);
        } else {
            self.autocomplete_selected -= 1;
        }
        self.update_formula_refs(cx);
        cx.notify();
    }

    /// Move autocomplete selection down
    pub fn autocomplete_down(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }
        let suggestions = self.autocomplete_suggestions(cx);
        if suggestions.is_empty() {
            return;
        }
        self.autocomplete_selected = (self.autocomplete_selected + 1) % suggestions.len();
        self.update_formula_refs(cx);
        cx.notify();
    }

    /// Accept the selected autocomplete suggestion
    pub fn autocomplete_accept(&mut self, cx: &mut Context<Self>) {
        if !self.autocomplete_visible {
            return;
        }

        let suggestions = self.autocomplete_suggestions(cx);
        if suggestions.is_empty() || self.autocomplete_selected >= suggestions.len() {
            self.autocomplete_visible = false;
            self.update_formula_refs(cx);
            return;
        }

        let entry = &suggestions[self.autocomplete_selected];
        let (replacement, range, opens_columns) = match entry {
            AutocompleteEntry::Table(s) => (s.replacement.clone(), s.range.clone(), s.opens_columns),
            _ => (format!("{}(",entry.name()),self.autocomplete_replace_range.clone(),false),
        };
        self.edit_value.replace_range(range.clone(), &replacement);
        self.edit_cursor = range.start + replacement.len();
        self.clear_edit_marks();
        self.edit_scroll_dirty = true;
        self.formula_bar_cache_dirty = true;
        self.reset_caret_activity();
        self.clear_formula_nav_override();
        self.update_formula_nav_mode();

        // Close autocomplete
        self.autocomplete_visible = false;
        self.autocomplete_selected = 0;

        // Enter formula mode if not already
        if !self.mode.is_formula() {
            self.mode = Mode::Formula;
        }

        self.autocomplete_suppressed = !opens_columns;
        self.update_formula_refs(cx);
        if opens_columns { self.update_autocomplete(cx); }
        cx.notify();
    }

    /// Dismiss autocomplete without accepting
    pub fn autocomplete_dismiss(&mut self, cx: &mut Context<Self>) {
        if self.autocomplete_visible {
            self.autocomplete_visible = false;
            self.autocomplete_suppressed = true;
            self.autocomplete_selected = 0;
            self.update_formula_refs(cx);
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

        let cursor = self.edit_value[..self.edit_cursor.min(self.edit_value.len())].chars().count();
        let ctx = formula_context::analyze(&self.edit_value, cursor);

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
        let custom_names_owned = crate::scripting::lua_formulas::function_names();
        let custom_names: Vec<&str> = custom_names_owned.iter().map(|s| s.as_str()).collect();
        check_errors(&self.edit_value, self.edit_cursor, &custom_names)
            .filter(|diag| matches!(diag.kind, DiagnosticKind::Hard))
            .map(|diag| FormulaErrorInfo {
                message: diag.message,
            })
    }
}
