//! Fill and transform operations for spreadsheet
//!
//! This module contains Fill Down, Fill Right, AutoSum, Fill Handle, and transform operations.

use gpui::*;
use regex::Regex;

use crate::app::{Spreadsheet, FillDrag, FillAxis};
use crate::history::CellChange;
use crate::mode::Mode;
use visigrid_engine::provenance::{MutationOp, FillDirection, FillMode};

/// Hit area size for fill handle (logical pixels, unscaled by zoom)
pub const FILL_HANDLE_HIT_SIZE: f32 = 14.0;

/// Visual size for fill handle (logical pixels, unscaled by zoom)
/// Excel uses approximately 6x6 pixels
pub const FILL_HANDLE_VISUAL_SIZE: f32 = 6.0;

/// Border width for fill handle (gives it the Excel-style white outline)
pub const FILL_HANDLE_BORDER: f32 = 1.0;

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
                self.set_cell_value(row, col, &new_value);
            }
        }

        if !changes.is_empty() {
            let provenance = MutationOp::Fill {
                sheet: self.sheet().id,
                src_start_row: min_row,
                src_start_col: min_col,
                src_end_row: min_row,
                src_end_col: max_col,
                dst_start_row: min_row + 1,
                dst_start_col: min_col,
                dst_end_row: max_row,
                dst_end_col: max_col,
                direction: FillDirection::Down,
                mode: FillMode::Both,
            }.to_provenance(&self.sheet().name);

            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        // Validate filled range and report failures
        let failures = self.workbook.validate_range(
            self.sheet_index(), min_row + 1, min_col, max_row, max_col
        );
        let total_cells = (max_row - min_row) * (max_col - min_col + 1);
        if failures.count > 0 {
            self.store_validation_failures(&failures);
            self.status_message = Some(format!(
                "Filled down (Validation: {} of {} cells failed) — Press F8 to jump",
                failures.count, total_cells
            ));
        } else {
            self.status_message = Some("Filled down".into());
        }
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
                self.set_cell_value(row, col, &new_value);
            }
        }

        if !changes.is_empty() {
            let provenance = MutationOp::Fill {
                sheet: self.sheet().id,
                src_start_row: min_row,
                src_start_col: min_col,
                src_end_row: max_row,
                src_end_col: min_col,
                dst_start_row: min_row,
                dst_start_col: min_col + 1,
                dst_end_row: max_row,
                dst_end_col: max_col,
                direction: FillDirection::Right,
                mode: FillMode::Both,
            }.to_provenance(&self.sheet().name);

            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        // Validate filled range and report failures
        let failures = self.workbook.validate_range(
            self.sheet_index(), min_row, min_col + 1, max_row, max_col
        );
        let total_cells = (max_row - min_row + 1) * (max_col - min_col);
        if failures.count > 0 {
            self.store_validation_failures(&failures);
            self.status_message = Some(format!(
                "Filled right (Validation: {} of {} cells failed) — Press F8 to jump",
                failures.count, total_cells
            ));
        } else {
            self.status_message = Some("Filled right".into());
        }
        cx.notify();
    }

    /// AutoSum: Insert =SUM() with detected range (Alt+=)
    /// Looks above and left for contiguous numeric cells, prefers above if longer
    pub fn autosum(&mut self, cx: &mut Context<Self>) {
        // Don't trigger if already editing
        if self.mode.is_editing() {
            return;
        }

        let (row, col) = self.view_state.selected;

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
        if let Some((start, end)) = detected_range {
            use crate::app::{RefKey, FormulaRef};
            let (r1, c1) = start;
            let key = if let Some((r2, c2)) = end {
                RefKey::Range { r1, c1, r2, c2 }
            } else {
                RefKey::Cell { row: r1, col: c1 }
            };
            // For AutoSum, the range text spans the whole formula argument
            // e.g., "=SUM(A1:A5)" - the text range would be 5..10
            let text_start = 5; // After "=SUM("
            let text_end = formula.len() - 1; // Before ")"
            self.formula_highlighted_refs = vec![FormulaRef {
                key,
                start,
                end,
                color_index: 0,
                text_range: text_start..text_end,
            }];
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

        let end_row = row - 1;
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

        let end_col = col - 1;
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
    /// Used by fill operations and multi-edit
    pub fn adjust_formula_refs(&self, formula: &str, delta_row: i32, delta_col: i32) -> String {
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

    // Transform operations

    /// Trim whitespace from all cells in the selection
    pub fn trim_whitespace(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        let mut changes = Vec::new();
        let mut trimmed_count = 0;

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);

                // Skip empty cells and formulas
                if old_value.is_empty() || old_value.starts_with('=') {
                    continue;
                }

                let new_value = old_value.trim().to_string();

                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                    self.set_cell_value(row, col, &new_value);
                    trimmed_count += 1;
                }
            }
        }

        if !changes.is_empty() {
            self.history.record_batch(self.sheet_index(), changes);
            self.bump_cells_rev();
            self.is_modified = true;
        }

        let msg = if trimmed_count == 0 {
            "No whitespace to trim".to_string()
        } else if trimmed_count == 1 {
            "Trimmed 1 cell".to_string()
        } else {
            format!("Trimmed {} cells", trimmed_count)
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    // Selection operations

    /// Select all blank (empty) cells within the current selection region
    pub fn select_blanks(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Find all blank cells in the region
        let mut blank_cells: Vec<(usize, usize)> = Vec::new();

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let value = self.sheet().get_raw(row, col);
                if value.is_empty() {
                    blank_cells.push((row, col));
                }
            }
        }

        if blank_cells.is_empty() {
            self.status_message = Some("No blank cells in selection".to_string());
            cx.notify();
            return;
        }

        // Set the first blank as the primary selection
        let first = blank_cells.remove(0);
        self.view_state.selected = first;
        self.view_state.selection_end = None;

        // Add remaining blanks as additional selections (single-cell each)
        self.view_state.additional_selections.clear();
        for cell in blank_cells {
            self.view_state.additional_selections.push((cell, None));
        }

        let count = 1 + self.view_state.additional_selections.len();
        let msg = if count == 1 {
            "Selected 1 blank cell".to_string()
        } else {
            format!("Selected {} blank cells", count)
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    // ========================================================================
    // Fill Handle (drag corner to fill cells)
    // ========================================================================

    /// Check if fill handle drag is currently active
    pub fn is_fill_dragging(&self) -> bool {
        matches!(self.fill_drag, FillDrag::Dragging { .. })
    }

    /// Start fill handle drag from the active cell
    pub fn start_fill_drag(&mut self, cx: &mut Context<Self>) {
        // Only allow from single cell (v1 limitation)
        if self.view_state.selection_end.is_some() || !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Fill handle works from single cell".into());
            cx.notify();
            return;
        }

        let anchor = self.view_state.selected;
        self.fill_drag = FillDrag::Dragging {
            anchor,
            current: anchor,
            axis: None,
        };
        cx.notify();
    }

    /// Continue fill handle drag - update current position and axis lock
    pub fn continue_fill_drag(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if let FillDrag::Dragging { anchor, current, axis } = self.fill_drag {
            // Skip if position unchanged
            if (row, col) == current {
                return;
            }

            let new_axis = if axis.is_some() {
                // Axis already locked - keep it
                axis
            } else {
                // Determine axis from movement direction
                let delta_row = (row as i32 - anchor.0 as i32).abs();
                let delta_col = (col as i32 - anchor.1 as i32).abs();

                if delta_row >= 1 || delta_col >= 1 {
                    // Lock axis based on primary direction
                    if delta_row > delta_col {
                        Some(FillAxis::Row)
                    } else if delta_col > delta_row {
                        Some(FillAxis::Col)
                    } else {
                        // Equal movement - prefer row (down) as default
                        Some(FillAxis::Row)
                    }
                } else {
                    None
                }
            };

            // Constrain current position to locked axis
            let constrained = match new_axis {
                Some(FillAxis::Row) => (row, anchor.1),  // Lock to anchor column
                Some(FillAxis::Col) => (anchor.0, col),  // Lock to anchor row
                None => (row, col),
            };

            self.fill_drag = FillDrag::Dragging {
                anchor,
                current: constrained,
                axis: new_axis,
            };
            cx.notify();
        }
    }

    /// End fill handle drag - execute the fill operation
    pub fn end_fill_drag(&mut self, cx: &mut Context<Self>) {
        if let FillDrag::Dragging { anchor, current, axis } = self.fill_drag {
            self.fill_drag = FillDrag::None;

            // No-op if current == anchor
            if current == anchor {
                cx.notify();
                return;
            }

            // Execute fill based on axis
            match axis {
                Some(FillAxis::Row) => {
                    self.execute_fill_handle_vertical(anchor, current, cx);
                }
                Some(FillAxis::Col) => {
                    self.execute_fill_handle_horizontal(anchor, current, cx);
                }
                None => {
                    // No axis determined - shouldn't happen, but no-op
                    cx.notify();
                }
            }
        }
    }

    /// Cancel fill handle drag without executing
    pub fn cancel_fill_drag(&mut self, cx: &mut Context<Self>) {
        if self.is_fill_dragging() {
            self.fill_drag = FillDrag::None;
            cx.notify();
        }
    }

    /// Execute vertical fill (fill down or up)
    fn execute_fill_handle_vertical(
        &mut self,
        anchor: (usize, usize),
        end: (usize, usize),
        cx: &mut Context<Self>,
    ) {
        let (anchor_row, col) = anchor;
        let end_row = end.0;

        if anchor_row == end_row {
            return;
        }

        let source = self.sheet().get_raw(anchor_row, col);
        let mut changes = Vec::new();

        // Determine fill direction and range (excluding anchor)
        let fill_range: Vec<usize> = if end_row > anchor_row {
            // Fill down
            ((anchor_row + 1)..=end_row).collect()
        } else {
            // Fill up
            (end_row..anchor_row).collect()
        };

        for row in &fill_range {
            let delta_row = *row as i32 - anchor_row as i32;
            let old_value = self.sheet().get_raw(*row, col);
            let new_value = if source.starts_with('=') {
                self.adjust_formula_refs(&source, delta_row, 0)
            } else {
                source.clone()
            };

            if old_value != new_value {
                changes.push(CellChange {
                    row: *row,
                    col,
                    old_value,
                    new_value: new_value.clone(),
                });
            }
            self.set_cell_value(*row, col, &new_value);
        }

        let count = fill_range.len();
        if !changes.is_empty() {
            let direction = if end_row > anchor_row { FillDirection::Down } else { FillDirection::Up };
            let (dst_start, dst_end) = if end_row > anchor_row {
                (anchor_row + 1, end_row)
            } else {
                (end_row, anchor_row - 1)
            };

            let provenance = MutationOp::Fill {
                sheet: self.sheet().id,
                src_start_row: anchor_row,
                src_start_col: col,
                src_end_row: anchor_row,
                src_end_col: col,
                dst_start_row: dst_start,
                dst_start_col: col,
                dst_end_row: dst_end,
                dst_end_col: col,
                direction,
                mode: FillMode::Both,
            }.to_provenance(&self.sheet().name);

            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        // Update selection to include filled range
        self.view_state.selection_end = Some(end);

        self.status_message = Some(format!("Filled {} cell{}", count, if count == 1 { "" } else { "s" }));
        cx.notify();
    }

    /// Execute horizontal fill (fill right or left)
    fn execute_fill_handle_horizontal(
        &mut self,
        anchor: (usize, usize),
        end: (usize, usize),
        cx: &mut Context<Self>,
    ) {
        let (row, anchor_col) = anchor;
        let end_col = end.1;

        if anchor_col == end_col {
            return;
        }

        let source = self.sheet().get_raw(row, anchor_col);
        let mut changes = Vec::new();

        // Determine fill direction and range (excluding anchor)
        let fill_range: Vec<usize> = if end_col > anchor_col {
            // Fill right
            ((anchor_col + 1)..=end_col).collect()
        } else {
            // Fill left
            (end_col..anchor_col).collect()
        };

        for col in &fill_range {
            let delta_col = *col as i32 - anchor_col as i32;
            let old_value = self.sheet().get_raw(row, *col);
            let new_value = if source.starts_with('=') {
                self.adjust_formula_refs(&source, 0, delta_col)
            } else {
                source.clone()
            };

            if old_value != new_value {
                changes.push(CellChange {
                    row,
                    col: *col,
                    old_value,
                    new_value: new_value.clone(),
                });
            }
            self.set_cell_value(row, *col, &new_value);
        }

        let count = fill_range.len();
        if !changes.is_empty() {
            let direction = if end_col > anchor_col { FillDirection::Right } else { FillDirection::Left };
            let (dst_start, dst_end) = if end_col > anchor_col {
                (anchor_col + 1, end_col)
            } else {
                (end_col, anchor_col - 1)
            };

            let provenance = MutationOp::Fill {
                sheet: self.sheet().id,
                src_start_row: row,
                src_start_col: anchor_col,
                src_end_row: row,
                src_end_col: anchor_col,
                dst_start_row: row,
                dst_start_col: dst_start,
                dst_end_row: row,
                dst_end_col: dst_end,
                direction,
                mode: FillMode::Both,
            }.to_provenance(&self.sheet().name);

            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        // Update selection to include filled range
        self.view_state.selection_end = Some(end);

        self.status_message = Some(format!("Filled {} cell{}", count, if count == 1 { "" } else { "s" }));
        cx.notify();
    }

    /// Get the fill preview target range (for rendering overlay)
    /// Returns None if not dragging or anchor == current
    /// Returns (min_row, min_col, max_row, max_col) excluding the anchor cell
    pub fn fill_drag_target_range(&self) -> Option<(usize, usize, usize, usize)> {
        if let FillDrag::Dragging { anchor, current, axis } = self.fill_drag {
            if current == anchor || axis.is_none() {
                return None;
            }

            match axis {
                Some(FillAxis::Row) => {
                    let (anchor_row, col) = anchor;
                    let end_row = current.0;
                    if anchor_row == end_row {
                        return None;
                    }
                    let min_row = anchor_row.min(end_row);
                    let max_row = anchor_row.max(end_row);
                    Some((min_row, col, max_row, col))
                }
                Some(FillAxis::Col) => {
                    let (row, anchor_col) = anchor;
                    let end_col = current.1;
                    if anchor_col == end_col {
                        return None;
                    }
                    let min_col = anchor_col.min(end_col);
                    let max_col = anchor_col.max(end_col);
                    Some((row, min_col, row, max_col))
                }
                None => None,
            }
        } else {
            None
        }
    }

    /// Check if a cell is in the fill preview range (excluding anchor)
    pub fn is_fill_preview_cell(&self, row: usize, col: usize) -> bool {
        if let FillDrag::Dragging { anchor, .. } = self.fill_drag {
            // Never highlight anchor
            if (row, col) == anchor {
                return false;
            }

            if let Some((min_row, min_col, max_row, max_col)) = self.fill_drag_target_range() {
                return row >= min_row && row <= max_row && col >= min_col && col <= max_col;
            }
        }
        false
    }
}
