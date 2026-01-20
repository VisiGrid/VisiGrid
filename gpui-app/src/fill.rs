//! Fill operations for spreadsheet
//!
//! This module contains Fill Down, Fill Right, and AutoSum functionality.

use gpui::*;
use regex::Regex;

use crate::app::Spreadsheet;
use crate::history::CellChange;
use crate::mode::Mode;

impl Spreadsheet {
    // Fill operations

    /// Fill down: copy the first row's values/formulas to remaining rows in selection
    pub fn fill_down(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Need at least 2 rows selected
        if max_row <= min_row {
            self.status_message = Some("Select at least 2 rows to fill down".into());
            cx.notify();
            return;
        }

        let mut changes = Vec::new();

        // For each column in selection
        for col in min_col..=max_col {
            // Get the source value/formula from the first row
            let source = self.sheet().get_raw(min_row, col);

            // Fill down to all other rows
            for row in (min_row + 1)..=max_row {
                let old_value = self.sheet().get_raw(row, col);
                let new_value = if source.starts_with('=') {
                    // Adjust relative references for formulas
                    self.adjust_formula_refs(&source, row as i32 - min_row as i32, 0)
                } else {
                    source.clone()
                };

                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(self.sheet_index(), changes);
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;

        self.status_message = Some("Filled down".into());
        cx.notify();
    }

    /// Fill right: copy the first column's values/formulas to remaining columns in selection
    pub fn fill_right(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Need at least 2 columns selected
        if max_col <= min_col {
            self.status_message = Some("Select at least 2 columns to fill right".into());
            cx.notify();
            return;
        }

        let mut changes = Vec::new();

        // For each row in selection
        for row in min_row..=max_row {
            // Get the source value/formula from the first column
            let source = self.sheet().get_raw(row, min_col);

            // Fill right to all other columns
            for col in (min_col + 1)..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                let new_value = if source.starts_with('=') {
                    // Adjust relative references for formulas
                    self.adjust_formula_refs(&source, 0, col as i32 - min_col as i32)
                } else {
                    source.clone()
                };

                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(self.sheet_index(), changes);
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        self.status_message = Some("Filled right".into());
        cx.notify();
    }

    /// AutoSum: Insert =SUM() with detected range (Alt+=)
    /// Looks above and left for contiguous numeric cells, prefers above if longer
    pub fn autosum(&mut self, cx: &mut Context<Self>) {
        // Don't trigger if already editing
        if self.mode.is_editing() {
            return;
        }

        let (row, col) = self.selected;

        // Find contiguous numeric cells above
        let above_range = self.find_numeric_range_above(row, col);

        // Find contiguous numeric cells to the left
        let left_range = self.find_numeric_range_left(row, col);

        // Choose the longer range, preferring above if equal
        // Also track the detected range for highlighting
        let (formula, detected_range) = match (above_range, left_range) {
            (Some((start_row, end_row)), Some((start_col, end_col))) => {
                let above_len = end_row - start_row + 1;
                let left_len = end_col - start_col + 1;
                if above_len >= left_len {
                    // Use above range
                    let start_ref = self.cell_ref_at(start_row, col);
                    let end_ref = self.cell_ref_at(end_row, col);
                    (format!("=SUM({}:{})", start_ref, end_ref),
                     Some(((start_row, col), Some((end_row, col)))))
                } else {
                    // Use left range
                    let start_ref = self.cell_ref_at(row, start_col);
                    let end_ref = self.cell_ref_at(row, end_col);
                    (format!("=SUM({}:{})", start_ref, end_ref),
                     Some(((row, start_col), Some((row, end_col)))))
                }
            }
            (Some((start_row, end_row)), None) => {
                // Only above range
                let start_ref = self.cell_ref_at(start_row, col);
                let end_ref = self.cell_ref_at(end_row, col);
                (format!("=SUM({}:{})", start_ref, end_ref),
                 Some(((start_row, col), Some((end_row, col)))))
            }
            (None, Some((start_col, end_col))) => {
                // Only left range
                let start_ref = self.cell_ref_at(row, start_col);
                let end_ref = self.cell_ref_at(row, end_col);
                (format!("=SUM({}:{})", start_ref, end_ref),
                 Some(((row, start_col), Some((row, end_col)))))
            }
            (None, None) => {
                // No range found, just insert empty SUM
                ("=SUM()".to_string(), None)
            }
        };

        // Set highlighted refs for the detected range
        if let Some(range) = detected_range {
            self.formula_highlighted_refs = vec![range];
        } else {
            self.formula_highlighted_refs.clear();
        }

        // Enter edit mode with the formula
        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = formula;
        self.edit_cursor = self.edit_value.chars().count(); // Cursor at end
        self.mode = Mode::Formula;
        self.update_autocomplete(cx);
        cx.notify();
    }

    /// Find contiguous numeric cells above the given cell
    /// Returns (start_row, end_row) if found, None otherwise
    fn find_numeric_range_above(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        if row == 0 {
            return None;
        }

        let mut end_row = row - 1;
        let mut start_row = end_row;

        // Check if the cell above is numeric
        if !self.is_cell_numeric(end_row, col) {
            return None;
        }

        // Walk upward finding contiguous numeric cells
        while start_row > 0 && self.is_cell_numeric(start_row - 1, col) {
            start_row -= 1;
        }

        // Need at least 2 cells
        if end_row - start_row + 1 >= 2 {
            Some((start_row, end_row))
        } else {
            // Single cell - still include it
            Some((start_row, end_row))
        }
    }

    /// Find contiguous numeric cells to the left of the given cell
    /// Returns (start_col, end_col) if found, None otherwise
    fn find_numeric_range_left(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        if col == 0 {
            return None;
        }

        let mut end_col = col - 1;
        let mut start_col = end_col;

        // Check if the cell to the left is numeric
        if !self.is_cell_numeric(row, end_col) {
            return None;
        }

        // Walk leftward finding contiguous numeric cells
        while start_col > 0 && self.is_cell_numeric(row, start_col - 1) {
            start_col -= 1;
        }

        // Need at least 2 cells
        if end_col - start_col + 1 >= 2 {
            Some((start_col, end_col))
        } else {
            // Single cell - still include it
            Some((start_col, end_col))
        }
    }

    /// Check if a cell contains a numeric value (not empty, not text, not error)
    fn is_cell_numeric(&self, row: usize, col: usize) -> bool {
        let raw = self.sheet().get_raw(row, col);
        if raw.is_empty() {
            return false;
        }

        // If it's a formula, check the result
        if raw.starts_with('=') {
            let display = self.sheet().get_display(row, col);
            // Check if display is a number
            display.parse::<f64>().is_ok()
        } else {
            // Check if raw value is a number
            raw.parse::<f64>().is_ok()
        }
    }

    /// Get cell reference string at given row, col (e.g., "A1", "B5")
    pub fn cell_ref_at(&self, row: usize, col: usize) -> String {
        let col_letter = Self::col_to_letter(col);
        format!("{}{}", col_letter, row + 1)
    }

    /// Adjust cell references in a formula by delta rows and cols
    /// Handles relative (A1), absolute ($A$1), and mixed ($A1, A$1) references
    fn adjust_formula_refs(&self, formula: &str, delta_row: i32, delta_col: i32) -> String {
        // Match cell references: optional $ before col, col letters, optional $ before row, row numbers
        let re = Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

        re.replace_all(formula, |caps: &regex::Captures| {
            let col_absolute = &caps[1] == "$";
            let col_letters = &caps[2];
            let row_absolute = &caps[3] == "$";
            let row_num: i32 = caps[4].parse().unwrap_or(1);

            // Parse column
            let col = col_letters.to_uppercase().chars().fold(0i32, |acc, c| {
                acc * 26 + (c as i32 - 'A' as i32 + 1)
            }) - 1;

            // Apply deltas if not absolute
            let new_col = if col_absolute { col } else { col + delta_col };
            let new_row = if row_absolute { row_num } else { row_num + delta_row };

            // Bounds check
            if new_col < 0 || new_row < 1 {
                return format!("#REF!");
            }

            // Convert column back to letters
            let col_str = Self::col_letter(new_col as usize);

            format!(
                "{}{}{}{}",
                if col_absolute { "$" } else { "" },
                col_str,
                if row_absolute { "$" } else { "" },
                new_row
            )
        })
        .to_string()
    }
}
