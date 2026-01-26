//! Find and Replace functionality for Spreadsheet.
//!
//! This module contains:
//! - MatchKind and MatchHit types for search results
//! - Find/Replace dialog control methods
//! - Search and replace operations with formula-aware matching

use gpui::*;
use crate::app::Spreadsheet;
use crate::history::CellChange;

/// The kind of cell content for a find match
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MatchKind {
    /// Raw text cell - can find and replace
    Text,
    /// Formula cell - can find and replace (token-aware)
    Formula,
}

/// A single match hit from find operation
#[derive(Clone, Debug)]
pub struct MatchHit {
    /// Sheet index (for future cross-sheet support)
    pub sheet: usize,
    /// Row index
    pub row: usize,
    /// Column index
    pub col: usize,
    /// What kind of cell this is
    pub kind: MatchKind,
    /// Byte offset of match start in the raw string
    pub start: usize,
    /// Byte offset of match end in the raw string
    pub end: usize,
}

impl Spreadsheet {
    // =========================================================================
    // Find and Replace
    // =========================================================================

    /// Show Find dialog (Ctrl+F)
    /// If already in Find mode, collapses to Find-only (hides Replace row)
    pub fn show_find(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            // Already open: collapse to Find-only mode, preserve inputs
            self.find_replace_mode = false;
            self.find_focus_replace = false;
            cx.notify();
            return;
        }

        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );

        // Fresh open: clear state
        self.lua_console.visible = false;
        self.mode = Mode::Find;
        self.find_input.clear();
        self.replace_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        self.find_replace_mode = false;
        self.find_focus_replace = false;
        cx.notify();
    }

    /// Show Find and Replace dialog (Ctrl+H)
    /// If already in Find mode, expands to show Replace row
    pub fn show_find_replace(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            // Already open: expand to Replace mode, preserve inputs
            self.find_replace_mode = true;
            // Focus Replace field if Find field has content, else stay on Find
            if !self.find_input.is_empty() {
                self.find_focus_replace = true;
            }
            cx.notify();
            return;
        }

        // Fresh open: clear state
        self.lua_console.visible = false;
        self.mode = Mode::Find;
        self.find_input.clear();
        self.replace_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        self.find_replace_mode = true;
        self.find_focus_replace = false;
        cx.notify();
    }

    pub fn hide_find(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Toggle focus between find and replace input fields
    pub fn find_toggle_focus(&mut self, cx: &mut Context<Self>) {
        if self.find_replace_mode {
            self.find_focus_replace = !self.find_focus_replace;
            cx.notify();
        }
    }

    pub fn find_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            if self.find_focus_replace {
                self.replace_input.push(c);
            } else {
                self.find_input.push(c);
                self.perform_find(cx);
            }
            cx.notify();
        }
    }

    pub fn find_backspace(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            if self.find_focus_replace {
                self.replace_input.pop();
            } else {
                self.find_input.pop();
                self.perform_find(cx);
            }
            cx.notify();
        }
    }

    /// Check if a string looks like a cell reference (A1, $A$1, Sheet1!A1, etc.)
    fn is_ref_like(s: &str) -> bool {
        let s = s.trim();
        if s.is_empty() {
            return false;
        }
        // Check for cell reference patterns: A1, $A1, A$1, $A$1, AA1, Sheet!A1
        // Simple heuristic: starts with optional $ or letter, contains letters followed by digits
        let s_upper = s.to_uppercase();
        let chars: Vec<char> = s_upper.chars().collect();

        // Skip sheet prefix (e.g., "Sheet1!")
        let start_idx = if let Some(pos) = s_upper.find('!') {
            pos + 1
        } else {
            0
        };

        if start_idx >= chars.len() {
            return false;
        }

        // After sheet prefix, check for ref pattern: [$]?[A-Z]+[$]?[0-9]+
        let mut i = start_idx;

        // Skip leading $
        if i < chars.len() && chars[i] == '$' {
            i += 1;
        }

        // Must have at least one letter
        let letter_start = i;
        while i < chars.len() && chars[i].is_ascii_alphabetic() {
            i += 1;
        }
        if i == letter_start {
            return false;
        }

        // Skip optional $ before row number
        if i < chars.len() && chars[i] == '$' {
            i += 1;
        }

        // Must have at least one digit
        let digit_start = i;
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }
        if i == digit_start {
            return false;
        }

        // Allow range suffix (:A1)
        if i < chars.len() && chars[i] == ':' {
            return true;  // It's a range reference
        }

        // Should be at end or have non-alphanumeric next
        i == chars.len() || !chars[i].is_alphanumeric()
    }

    /// Perform search and populate find_results with MatchHit entries.
    /// Only searches in Text and Formula cells (not numbers/dates).
    pub(crate) fn perform_find(&mut self, cx: &mut Context<Self>) {
        use visigrid_engine::cell::CellValue;

        self.find_results.clear();
        self.find_index = 0;

        if self.find_input.is_empty() {
            self.status_message = None;
            cx.notify();
            return;
        }

        let query = self.find_input.to_lowercase();
        let sheet_idx = self.workbook.active_sheet_index();

        // Collect cell data to search
        let cells_to_search: Vec<_> = self.sheet().cells_iter()
            .filter_map(|(&(row, col), cell)| {
                match &cell.value {
                    CellValue::Text(text) => Some((row, col, MatchKind::Text, text.clone())),
                    CellValue::Formula { source, .. } => Some((row, col, MatchKind::Formula, source.clone())),
                    _ => None,  // Skip Empty, Number - they're not replaceable
                }
            })
            .collect();

        // Find all matches
        for (row, col, kind, raw_text) in cells_to_search {
            let raw_lower = raw_text.to_lowercase();

            // Find all occurrences within this cell
            let mut search_start = 0;
            while let Some(rel_pos) = raw_lower[search_start..].find(&query) {
                let start = search_start + rel_pos;
                let end = start + query.len();

                // For formulas, skip matches inside string literals
                if kind == MatchKind::Formula && Self::is_inside_string_literal(&raw_text, start) {
                    search_start = end;
                    continue;
                }

                self.find_results.push(MatchHit {
                    sheet: sheet_idx,
                    row,
                    col,
                    kind,
                    start,
                    end,
                });

                search_start = end;
            }
        }

        // Sort results by row, then column, then offset
        self.find_results.sort_by(|a, b| {
            a.row.cmp(&b.row)
                .then(a.col.cmp(&b.col))
                .then(a.start.cmp(&b.start))
        });

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

    /// Check if a position is inside a string literal in a formula
    fn is_inside_string_literal(formula: &str, pos: usize) -> bool {
        let bytes = formula.as_bytes();
        let mut in_string = false;
        let mut i = 0;

        while i < pos && i < bytes.len() {
            if bytes[i] == b'"' {
                in_string = !in_string;
            }
            i += 1;
        }

        in_string
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
        if let Some(hit) = self.find_results.get(self.find_index) {
            self.view_state.selected = (hit.row, hit.col);
            self.view_state.selection_end = None;
            self.ensure_visible(cx);
            self.status_message = Some(format!(
                "Match {} of {}",
                self.find_index + 1,
                self.find_results.len()
            ));
        }
    }

    /// Replace the current match and move to next
    /// In Find-only mode, this just does FindNext
    pub fn replace_next(&mut self, cx: &mut Context<Self>) {
        // In Find-only mode, Enter does FindNext
        if !self.find_replace_mode {
            self.find_next(cx);
            return;
        }

        if self.find_results.is_empty() {
            return;
        }

        let hit = match self.find_results.get(self.find_index) {
            Some(h) => h.clone(),
            None => return,
        };

        // Get the raw value
        let raw_value = self.sheet().get_raw(hit.row, hit.col);

        // Perform the replacement
        let new_value = if hit.kind == MatchKind::Formula && Self::is_ref_like(&self.find_input) {
            // Token-aware replacement for ref-like patterns
            self.replace_in_formula_token_aware(&raw_value, &self.find_input, &self.replace_input)
        } else {
            // Simple case-insensitive replacement
            self.replace_case_insensitive(&raw_value, hit.start, hit.end, &self.replace_input)
        };

        // Record undo and apply
        let sheet_index = self.sheet_index();
        self.history.record_change(sheet_index, hit.row, hit.col, raw_value, new_value.clone());
        self.sheet_mut().set_value(hit.row, hit.col, &new_value);
        cx.notify();

        // Recompute find results (offsets have changed)
        self.perform_find(cx);

        // Try to stay at similar position or advance
        if self.find_index >= self.find_results.len() && !self.find_results.is_empty() {
            self.find_index = 0;
        }

        if !self.find_results.is_empty() {
            self.jump_to_find_result(cx);
        }
    }

    /// Replace all matches at once
    pub fn replace_all(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() || !self.find_replace_mode {
            return;
        }

        // Take a snapshot of matches before mutation
        let hits: Vec<MatchHit> = self.find_results.clone();
        let total = hits.len();

        // Group hits by cell (row, col) for batch replacement
        let mut cells_to_replace: std::collections::HashMap<(usize, usize), Vec<MatchHit>> =
            std::collections::HashMap::new();

        for hit in hits {
            cells_to_replace
                .entry((hit.row, hit.col))
                .or_default()
                .push(hit);
        }

        // Collect all changes for batch undo
        let mut changes: Vec<CellChange> = Vec::new();
        let mut replaced_count = 0;

        for ((row, col), mut cell_hits) in cells_to_replace {
            // Sort hits by start position descending (replace from end to preserve offsets)
            cell_hits.sort_by(|a, b| b.start.cmp(&a.start));

            let raw_value = self.sheet().get_raw(row, col);
            let mut new_value = raw_value.clone();

            // Apply replacements in reverse order
            for hit in cell_hits {
                let kind = hit.kind;
                if kind == MatchKind::Formula && Self::is_ref_like(&self.find_input) {
                    // For ref-like patterns in formulas, use token-aware replacement
                    new_value = self.replace_in_formula_token_aware(
                        &new_value,
                        &self.find_input,
                        &self.replace_input,
                    );
                    replaced_count += 1;
                    break;  // Token-aware replaces all at once
                } else {
                    // Simple replacement at specific offset
                    new_value = self.replace_case_insensitive(
                        &new_value,
                        hit.start,
                        hit.end,
                        &self.replace_input,
                    );
                    replaced_count += 1;
                }
            }

            // Record change for undo
            changes.push(CellChange {
                row,
                col,
                old_value: raw_value,
                new_value: new_value.clone(),
            });

            self.sheet_mut().set_value(row, col, &new_value);
        }

        // Record all changes as single batch undo
        let sheet_index = self.sheet_index();
        self.history.record_batch(sheet_index, changes);

        // Clear results and show status
        self.find_results.clear();
        self.find_index = 0;
        self.status_message = Some(format!("Replaced {} of {} matches", replaced_count, total));
        cx.notify();
    }

    /// Case-insensitive replacement at specific byte offsets
    fn replace_case_insensitive(&self, text: &str, start: usize, end: usize, replacement: &str) -> String {
        let mut result = String::with_capacity(text.len() + replacement.len());
        result.push_str(&text[..start]);
        result.push_str(replacement);
        result.push_str(&text[end..]);
        result
    }

    /// Token-aware replacement in formula for cell references
    /// This preserves references that partially match (e.g., A1 vs A10)
    fn replace_in_formula_token_aware(&self, formula: &str, find: &str, replace: &str) -> String {
        let find_upper = find.to_uppercase();
        let mut result = String::with_capacity(formula.len());
        let chars: Vec<char> = formula.chars().collect();
        let mut i = 0;
        let mut in_string = false;

        while i < chars.len() {
            // Track string literals
            if chars[i] == '"' {
                in_string = !in_string;
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // Don't replace inside strings
            if in_string {
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // Check for match at current position
            let remaining: String = chars[i..].iter().collect();
            let remaining_upper = remaining.to_uppercase();

            if remaining_upper.starts_with(&find_upper) {
                // Check word boundaries
                let before_ok = i == 0 || !chars[i - 1].is_alphanumeric();
                let after_idx = i + find.len();
                let after_ok = after_idx >= chars.len() || !chars[after_idx].is_alphanumeric();

                if before_ok && after_ok {
                    // Replace with same case as replacement input
                    result.push_str(replace);
                    i += find.len();
                    continue;
                }
            }

            result.push(chars[i]);
            i += 1;
        }

        result
    }
}
