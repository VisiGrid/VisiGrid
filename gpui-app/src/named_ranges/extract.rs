//! Extract Named Range - extract range literals from formulas into named ranges

use gpui::*;
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::history::UndoAction;
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Extract Named Range methods
    // =========================================================================

    /// Show the extract named range modal
    pub fn show_extract_named_range(&mut self, cx: &mut Context<Self>) {
        // Get the current cell's formula
        let (row, col) = self.view_state.selected;
        let cell = self.sheet().get_cell(row, col);
        let formula_opt = self.get_formula_source(&cell.value);

        let formula = match formula_opt {
            Some(f) => f,
            None => {
                self.status_message = Some("Place the cursor inside a formula containing a range.".to_string());
                cx.notify();
                return;
            }
        };

        // Detect range literals in the formula
        let range_literal = match self.detect_range_literal(&formula) {
            Some(r) => r,
            None => {
                self.status_message = Some("No range literal found in formula.".to_string());
                cx.notify();
                return;
            }
        };

        // Check if this range is already a named range
        if self.workbook.get_named_range(&range_literal).is_some() {
            self.status_message = Some(format!("'{}' is already a named range.", range_literal));
            cx.notify();
            return;
        }

        // Find all cells containing this range literal
        let (affected_cells, occurrence_count) = self.find_cells_with_range(&range_literal);

        // Generate a suggested name (Range_1, Range_2, etc.)
        let suggested_name = self.generate_unique_range_name();

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
    fn generate_unique_range_name(&self) -> String {
        let mut i = 1;
        loop {
            let name = format!("Range_{}", i);
            if self.workbook.get_named_range(&name).is_none() {
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

    /// Detect a range literal in a formula (e.g., A1:B10, $A$1:$B$10)
    fn detect_range_literal(&self, formula: &str) -> Option<String> {
        // Simple regex-like pattern matching for range literals
        // Matches: A1:B10, $A$1:$B$10, A1, $A$1, etc.
        let chars: Vec<char> = formula.chars().collect();
        let mut i = 0;

        while i < chars.len() {
            // Look for start of a cell reference
            if let Some(range) = self.try_parse_range_at(&chars, i) {
                // Skip named ranges (already defined)
                if self.workbook.get_named_range(&range).is_none() {
                    // Make sure it's actually a range (contains :) or a single cell
                    return Some(range);
                }
            }
            i += 1;
        }
        None
    }

    /// Try to parse a range starting at position i
    fn try_parse_range_at(&self, chars: &[char], start: usize) -> Option<String> {
        let mut i = start;

        // Skip $ if present
        if i < chars.len() && chars[i] == '$' {
            i += 1;
        }

        // Need at least one letter
        if i >= chars.len() || !chars[i].is_ascii_alphabetic() {
            return None;
        }

        // Collect column letters
        while i < chars.len() && chars[i].is_ascii_alphabetic() {
            i += 1;
        }

        // Skip $ if present before row
        if i < chars.len() && chars[i] == '$' {
            i += 1;
        }

        // Need at least one digit
        if i >= chars.len() || !chars[i].is_ascii_digit() {
            return None;
        }

        // Collect row digits
        while i < chars.len() && chars[i].is_ascii_digit() {
            i += 1;
        }

        // Check for range separator (:)
        if i < chars.len() && chars[i] == ':' {
            i += 1;

            // Parse second cell reference
            // Skip $ if present
            if i < chars.len() && chars[i] == '$' {
                i += 1;
            }

            // Need at least one letter
            if i >= chars.len() || !chars[i].is_ascii_alphabetic() {
                return None;
            }

            // Collect column letters
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }

            // Skip $ if present before row
            if i < chars.len() && chars[i] == '$' {
                i += 1;
            }

            // Need at least one digit
            if i >= chars.len() || !chars[i].is_ascii_digit() {
                return None;
            }

            // Collect row digits
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
        }

        // Make sure next char is not alphanumeric (word boundary)
        if i < chars.len() && (chars[i].is_alphanumeric() || chars[i] == '_') {
            return None;
        }

        // Make sure previous char is not alphanumeric (word boundary)
        if start > 0 && (chars[start - 1].is_alphanumeric() || chars[start - 1] == '_') {
            return None;
        }

        Some(chars[start..i].iter().collect())
    }

    /// Find all cells containing a specific range literal and count occurrences
    fn find_cells_with_range(&self, range_literal: &str) -> (Vec<(usize, usize)>, usize) {
        let range_upper = range_literal.to_uppercase();
        let mut cells = Vec::new();
        let mut total_count = 0;

        for ((row, col), cell) in self.sheet().cells_iter() {
            let raw = cell.value.raw_display();
            if !raw.starts_with('=') {
                continue;
            }

            let formula_upper = raw.to_uppercase();
            let count = self.count_range_occurrences(&formula_upper, &range_upper);
            if count > 0 {
                cells.push((*row, *col));
                total_count += count;
            }
        }

        (cells, total_count)
    }

    /// Count how many times a range appears in a formula
    fn count_range_occurrences(&self, formula: &str, range: &str) -> usize {
        let mut count = 0;
        let chars: Vec<char> = formula.chars().collect();
        let range_chars: Vec<char> = range.chars().collect();
        let range_len = range_chars.len();

        let mut i = 0;
        while i + range_len <= chars.len() {
            // Check for match
            let slice: String = chars[i..i + range_len].iter().collect();
            if slice == range {
                // Verify word boundaries
                let before_ok = i == 0 || (!chars[i - 1].is_alphanumeric() && chars[i - 1] != '_' && chars[i - 1] != '$');
                let after_ok = i + range_len >= chars.len() || (!chars[i + range_len].is_alphanumeric() && chars[i + range_len] != '_');
                if before_ok && after_ok {
                    count += 1;
                    i += range_len;
                    continue;
                }
            }
            i += 1;
        }
        count
    }

    /// Hide the extract named range modal
    pub fn hide_extract_named_range(&mut self, cx: &mut Context<Self>) {
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
    fn validate_extract_name(&mut self) {
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
        if self.workbook.get_named_range(&self.extract_name).is_some() {
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
        self.validate_extract_name();
        cx.notify();
    }

    /// Backspace in extract name
    pub fn extract_name_backspace(&mut self, cx: &mut Context<Self>) {
        self.extract_select_all = false;
        self.extract_name.pop();
        self.validate_extract_name();
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

    /// Confirm extraction - create named range and replace in formulas
    pub fn confirm_extract_named_range(&mut self, cx: &mut Context<Self>) {
        // Validate
        if self.extract_name.is_empty() {
            self.extract_validation_error = Some("Name cannot be empty".to_string());
            cx.notify();
            return;
        }
        self.validate_extract_name();
        if self.extract_validation_error.is_some() {
            cx.notify();
            return;
        }

        let range_literal = self.extract_range_literal.clone();
        let name = self.extract_name.clone();
        let description = if self.extract_description.is_empty() {
            None
        } else {
            Some(self.extract_description.clone())
        };
        let affected_cells = std::mem::take(&mut self.extract_affected_cells);
        let occurrence_count = self.extract_occurrence_count;

        // 1. Parse the range literal and create the named range
        // Handle absolute references by removing $ signs
        let clean_range = range_literal.replace('$', "");
        let parts: Vec<&str> = clean_range.split(':').collect();

        let sheet = self.workbook.active_sheet_index();
        let result: Result<(), String> = if parts.len() == 2 {
            // Range reference like A1:B10
            if let (Some(start), Some(end)) = (
                Self::parse_cell_ref(parts[0]),
                Self::parse_cell_ref(parts[1]),
            ) {
                self.workbook.define_name_for_range(&name, sheet, start.0, start.1, end.0, end.1)
            } else {
                Err("Invalid cell reference".to_string())
            }
        } else {
            // Single cell reference like A1
            if let Some((row, col)) = Self::parse_cell_ref(&clean_range) {
                self.workbook.define_name_for_cell(&name, sheet, row, col)
            } else {
                Err("Invalid cell reference".to_string())
            }
        };

        if let Err(e) = result {
            self.extract_validation_error = Some(format!("Failed to create named range: {:?}", e));
            cx.notify();
            return;
        }

        // Add description if provided
        if let Some(desc) = description {
            if let Some(nr) = self.workbook.named_ranges_mut().get(&name).cloned() {
                let mut updated = nr;
                updated.description = Some(desc);
                let _ = self.workbook.named_ranges_mut().set(updated);
            }
        }

        // 2. Replace range literal with name in all affected cells
        let mut cell_changes = Vec::new();
        for (row, col) in &affected_cells {
            let cell = self.sheet().get_cell(*row, *col);
            let old_value = cell.value.raw_display();
            if old_value.starts_with('=') {
                let new_value = self.replace_range_in_formula(&old_value, &range_literal, &name);
                if new_value != old_value {
                    // Apply the change
                    self.sheet_mut().set_value(*row, *col, &new_value);
                    cell_changes.push(crate::history::CellChange {
                        row: *row,
                        col: *col,
                        old_value,
                        new_value,
                    });
                }
            }
        }

        // 3. Record undo action (group)
        let mut actions = vec![
            UndoAction::NamedRangeCreated { name: name.clone() },
        ];
        if !cell_changes.is_empty() {
            actions.push(UndoAction::Values {
                sheet_index: 0,
                changes: cell_changes,
            });
        }
        self.history.record_named_range_action(UndoAction::Group {
            actions,
            description: format!("Extract '{}'", name),
        });

        // 4. Add to refactor log
        let impact_msg = format!("Replaced {} occurrences in {} cells", occurrence_count, affected_cells.len());
        self.refactor_log.push(
            crate::views::refactor_log::RefactorLogEntry::new(
                "Extracted to Named Range",
                format!("{} = {}", name, range_literal),
            ).with_impact(impact_msg)
        );

        // 5. Invalidate caches and show status
        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Extracted '{}' (Ctrl+Shift+R to rename)", name));

        // 6. Hide modal
        self.hide_extract_named_range(cx);
    }

    /// Replace all occurrences of a range literal with a name in a formula.
    /// This is token-aware: it won't replace inside string literals.
    fn replace_range_in_formula(&self, formula: &str, range_literal: &str, name: &str) -> String {
        let range_upper = range_literal.to_uppercase();
        let mut result = String::new();
        let chars: Vec<char> = formula.chars().collect();
        let range_len = range_upper.len();

        let mut i = 0;
        let mut in_string = false;

        while i < chars.len() {
            // Track string literal state (toggle on each unescaped quote)
            if chars[i] == '"' {
                // Check for escaped quote (doubled quote in Excel formulas)
                if in_string && i + 1 < chars.len() && chars[i + 1] == '"' {
                    result.push(chars[i]);
                    result.push(chars[i + 1]);
                    i += 2;
                    continue;
                }
                in_string = !in_string;
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // If inside a string, just copy the character
            if in_string {
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // Check for range match (only outside strings)
            if i + range_len <= chars.len() {
                let slice: String = chars[i..i + range_len].iter().collect::<String>().to_uppercase();
                if slice == range_upper {
                    // Verify word boundaries
                    let before_ok = i == 0 || (!chars[i - 1].is_alphanumeric() && chars[i - 1] != '_' && chars[i - 1] != '$');
                    let after_ok = i + range_len >= chars.len() || (!chars[i + range_len].is_alphanumeric() && chars[i + range_len] != '_');
                    if before_ok && after_ok {
                        result.push_str(name);
                        i += range_len;
                        continue;
                    }
                }
            }
            result.push(chars[i]);
            i += 1;
        }
        result
    }
}
