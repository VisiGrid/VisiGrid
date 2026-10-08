//! Extract Named Range - extract range literals from formulas into named ranges

use gpui::{*};
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::history::MutationSource;
use super::extract_plan::ExtractionDraft;
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Extract Named Range methods
    // =========================================================================

    /// Show the extract named range modal
    pub fn show_extract_named_range(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) { return; }
        self.sync_table_view(cx);
        let draft = match ExtractionDraft::capture(self.wb(cx), &self.row_view, self.view_state.selected) {
            Ok(draft) => draft,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let range_literal = draft.literal.clone();
        let affected_cells = draft.cells();
        let occurrence_count = draft.occurrences;
        self.extract_draft = Some(draft);

        // Generate a suggested name (Range_1, Range_2, etc.)
        let suggested_name = self.generate_unique_range_name(cx);

        self.extract_range_literal = range_literal;
        self.extract_name = suggested_name;
        self.extract_description = String::new();
        self.extract_affected_cells = affected_cells;
        self.extract_occurrence_count = occurrence_count;
        self.extract_validation_error = None;
        self.extract_select_all = true;  // Type to replace the suggested name
        self.extract_focus = CreateNameFocus::Name;
        self.mode = Mode::ExtractNamedRange;
        cx.notify();
    }

    /// Generate a unique name like Range_1, Range_2, etc.
    fn generate_unique_range_name(&self, cx: &App) -> String {
        let mut i = 1;
        loop {
            let name = format!("Range_{}", i);
            if self.wb(cx).get_named_range(&name).is_none() && self.wb(cx).table_by_name(&name).is_none() {
                return name;
            }
            i += 1;
            if i > 1000 {
                // Fallback to avoid infinite loop
                return format!("ExtractedRange_{}", std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_secs())
                    .unwrap_or(0));
            }
        }
    }

    /// Hide the extract named range modal
    pub fn hide_extract_named_range(&mut self, cx: &mut Context<Self>) {
        self.extract_draft = None;
        self.extract_range_literal.clear();
        self.extract_name.clear();
        self.extract_description.clear();
        self.extract_affected_cells.clear();
        self.extract_occurrence_count = 0;
        self.extract_validation_error = None;
        self.extract_select_all = false;
        self.extract_focus = CreateNameFocus::default();
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Tab between fields in extract dialog
    pub fn extract_tab(&mut self, cx: &mut Context<Self>) {
        self.extract_focus = match self.extract_focus {
            CreateNameFocus::Name => CreateNameFocus::Description,
            CreateNameFocus::Description => CreateNameFocus::Name,
        };
        cx.notify();
    }

    /// Validate the extract name
    fn validate_extract_name(&mut self, cx: &App) {
        if self.extract_name.is_empty() {
            self.extract_validation_error = Some("Name cannot be empty".to_string());
            return;
        }

        // Check first character is letter or underscore
        let first_char = self.extract_name.chars().next().unwrap();
        if !first_char.is_ascii_alphabetic() && first_char != '_' {
            self.extract_validation_error = Some("Name must start with a letter or underscore".to_string());
            return;
        }

        // Check all characters are valid
        for c in self.extract_name.chars() {
            if !c.is_alphanumeric() && c != '_' && c != '.' {
                self.extract_validation_error = Some("Name can only contain letters, numbers, underscore, and dot".to_string());
                return;
            }
        }

        // Check for reserved names/cell references
        let name_upper = self.extract_name.to_uppercase();
        if self.is_reserved_name(&name_upper) {
            self.extract_validation_error = Some("This name is reserved or looks like a cell reference".to_string());
            return;
        }

        // Check for existing named range
        if self.wb(cx).get_named_range(&self.extract_name).is_some() {
            self.extract_validation_error = Some("A named range with this name already exists".to_string());
            return;
        }

        self.extract_validation_error = None;
    }

    /// Check if a name is reserved (cell reference, function name, etc.)
    fn is_reserved_name(&self, name: &str) -> bool {
        // Check if it looks like a cell reference
        let chars: Vec<char> = name.chars().collect();
        if !chars.is_empty() && chars[0].is_ascii_alphabetic() {
            let mut i = 0;
            // Skip letters
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }
            // If remaining are all digits, it looks like a cell ref
            if i < chars.len() && chars[i..].iter().all(|c| c.is_ascii_digit()) {
                return true;
            }
        }

        // Check against known function names (simplified list)
        let reserved = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "AND", "OR", "NOT",
                       "TRUE", "FALSE", "PI", "E", "ABS", "SQRT", "ROUND", "INT", "MOD",
                       "POWER", "LOG", "LN", "EXP", "SIN", "COS", "TAN"];
        reserved.contains(&name)
    }

    /// Insert a character into the extract name
    pub fn extract_name_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.extract_select_all {
            self.extract_name.clear();
            self.extract_select_all = false;
        }
        self.extract_name.push(c);
        self.validate_extract_name(cx);
        cx.notify();
    }

    /// Backspace in extract name
    pub fn extract_name_backspace(&mut self, cx: &mut Context<Self>) {
        self.extract_select_all = false;
        self.extract_name.pop();
        self.validate_extract_name(cx);
        cx.notify();
    }

    /// Insert a character into the extract description
    pub fn extract_description_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.extract_description.push(c);
        cx.notify();
    }

    /// Backspace in extract description
    pub fn extract_description_backspace(&mut self, cx: &mut Context<Self>) {
        self.extract_description.pop();
        cx.notify();
    }

    /// Publish the captured name and formula changes as one atomic history entry.
    pub fn confirm_extract_named_range(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) { return; }
        self.validate_extract_name(cx);
        if self.extract_validation_error.is_some() { cx.notify(); return; }
        let name = self.extract_name.clone();
        let description = (!self.extract_description.is_empty()).then(|| self.extract_description.clone());
        let result = self.extract_draft.as_ref()
            .ok_or_else(|| "Reopen extraction and try again.".to_string())
            .and_then(|draft| draft.prepare(self.wb(cx), &self.row_view, &name, description))
            .and_then(|(candidate, commit)| self.publish_table_batch(candidate, commit, format!("Extract '{name}'"), MutationSource::Human, cx));
        if let Err(error) = result {
            self.extract_validation_error = Some(error);
            cx.notify(); return;
        }
        self.refactor_log.push(
            crate::views::refactor_log::RefactorLogEntry::new(
                "Extracted to Named Range", format!("{} = {}", name, self.extract_range_literal),
            ).with_impact(format!("Replaced {} occurrences in {} visible formulas", self.extract_occurrence_count, self.extract_affected_cells.len()))
        );
        self.status_message = Some(format!("Extracted '{name}' (Ctrl+Shift+R to rename)"));
        self.hide_extract_named_range(cx);
    }
}
