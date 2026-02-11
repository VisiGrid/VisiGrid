//! Formula Reference Helpers
//!
//! Contains:
//! - Cell reference types (RefKey, FormulaRef)
//! - Reference formatting (col_to_letter, make_cell_ref, make_range_ref)
//! - Active ref target detection and borders
//! - Formula ref highlighting and color assignment
//! - Cell reference parsing

use std::collections::HashMap;
use crate::app::Spreadsheet;
use crate::formula_context::{tokenize_for_highlight, TokenType, char_index_to_byte_offset};

// ============================================================================
// Types
// ============================================================================

/// Stable key for formula reference deduplication - same ref gets same color
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum RefKey {
    Cell { row: usize, col: usize },
    Range { r1: usize, c1: usize, r2: usize, c2: usize },  // normalized min/max
}

impl RefKey {
    /// Construct a RefKey, collapsing single-cell ranges to Cell variant.
    /// Coordinates are normalized (min/max) for ranges.
    pub fn new(r1: usize, c1: usize, r2: usize, c2: usize) -> Self {
        if r1 == r2 && c1 == c2 {
            RefKey::Cell { row: r1, col: c1 }
        } else {
            RefKey::Range { r1, c1, r2, c2 }
        }
    }
}

/// Formula reference with color assignment and text position
#[derive(Clone, Debug)]
pub struct FormulaRef {
    pub key: RefKey,
    pub start: (usize, usize),                // top-left of range
    pub end: Option<(usize, usize)>,          // bottom-right (None for single cell)
    pub color_index: usize,                   // 0-7 rotating
    pub text_byte_range: std::ops::Range<usize>,  // byte range in formula text
}

/// Color palette for formula references (Excel-like)
pub const REF_COLORS: [u32; 8] = [
    0x4472C4,  // 0: Blue
    0xED7D31,  // 1: Orange
    0x9B59B6,  // 2: Purple
    0x70AD47,  // 3: Green
    0x00B0F0,  // 4: Cyan
    0xFFC000,  // 5: Yellow
    0xFF6B9D,  // 6: Pink
    0x00B294,  // 7: Teal
];

// ============================================================================
// Helper Functions
// ============================================================================

impl Spreadsheet {
    /// Convert column index to Excel-style letter(s): 0 -> A, 25 -> Z, 26 -> AA
    pub fn col_to_letter(col: usize) -> String {
        let mut result = String::new();
        let mut n = col;
        loop {
            result.insert(0, (b'A' + (n % 26) as u8) as char);
            if n < 26 {
                break;
            }
            n = n / 26 - 1;
        }
        result
    }

    /// Convert (row, col) to cell reference string: (0, 0) -> "A1"
    pub fn make_cell_ref(row: usize, col: usize) -> String {
        format!("{}{}", Self::col_to_letter(col), row + 1)
    }

    /// Convert range to reference string: ((0, 0), (2, 3)) -> "A1:D3"
    pub fn make_range_ref(start: (usize, usize), end: (usize, usize)) -> String {
        let (r1, c1) = (start.0.min(end.0), start.1.min(end.1));
        let (r2, c2) = (start.0.max(end.0), start.1.max(end.1));
        if r1 == r2 && c1 == c2 {
            Self::make_cell_ref(r1, c1)
        } else {
            format!("{}:{}", Self::make_cell_ref(r1, c1), Self::make_cell_ref(r2, c2))
        }
    }

    // ========================================================================
    // Active Ref Target (formula mode navigation)
    // ========================================================================

    /// Check if a cell is the active reference navigation target (formula_ref_cell).
    /// This is distinct from parsed formula refs - it's where arrow keys are pointing RIGHT NOW.
    /// Used for rendering a bright "target" indicator during formula reference navigation.
    pub fn is_active_ref_target(&self, row: usize, col: usize) -> bool {
        if !self.mode.is_formula() {
            return false;
        }

        if let Some(rect) = self.ref_target_rect() {
            crate::ref_target::contains(&rect, row, col)
        } else {
            false
        }
    }

    /// Get the normalized rectangle for the current ref target, if any.
    fn ref_target_rect(&self) -> Option<crate::ref_target::Rect> {
        let (ref_row, ref_col) = self.formula_ref_cell?;
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((ref_row, ref_col));
        Some(crate::ref_target::normalize_rect((ref_row, ref_col), (end_row, end_col)))
    }

    /// Get the border edges to draw for the active ref target (like selection_borders but for ref target)
    pub fn ref_target_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        if !self.mode.is_formula() {
            return (false, false, false, false);
        }

        let Some(rect) = self.ref_target_rect() else {
            return (false, false, false, false);
        };

        let edges = crate::ref_target::borders(&rect, row, col);
        (edges.top, edges.right, edges.bottom, edges.left)
    }

    // ========================================================================
    // Formula Ref Highlighting
    // ========================================================================

    /// Check if a cell is within any formula reference (for highlighting)
    /// This includes both the live pointing reference AND parsed refs from existing formulas
    pub fn is_formula_ref(&self, row: usize, col: usize) -> bool {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return false;
        }

        // Check the live pointing reference first
        if let Some((ref_row, ref_col)) = self.formula_ref_cell {
            if let Some((end_row, end_col)) = self.formula_ref_end {
                // Range reference - check if cell is within the range
                let (min_row, max_row) = (ref_row.min(end_row), ref_row.max(end_row));
                let (min_col, max_col) = (ref_col.min(end_col), ref_col.max(end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    return true;
                }
            } else {
                // Single cell reference
                if row == ref_row && col == ref_col {
                    return true;
                }
            }
        }

        // Check the highlighted refs from parsed formula
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                // Range - check if cell is within
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    return true;
                }
            } else {
                // Single cell
                if row == fref.start.0 && col == fref.start.1 {
                    return true;
                }
            }
        }

        false
    }

    /// Get the color index for a formula reference at this cell (for multi-color highlighting).
    /// Returns the earliest ref's color (by text position) to avoid muddy overlap.
    /// Returns None if cell is not a formula ref.
    pub fn formula_ref_color(&self, row: usize, col: usize) -> Option<usize> {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return None;
        }

        // Check the highlighted refs (already sorted by text position, so first match = earliest)
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                // Range
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    return Some(fref.color_index);
                }
            } else {
                // Single cell
                if row == fref.start.0 && col == fref.start.1 {
                    return Some(fref.color_index);
                }
            }
        }

        None
    }

    /// Get which borders should be drawn for a formula ref cell
    /// Returns (top, right, bottom, left) - true means draw that border
    pub fn formula_ref_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return (false, false, false, false);
        }

        let mut top = false;
        let mut right = false;
        let mut bottom = false;
        let mut left = false;

        // Check the live pointing reference
        if let Some((ref_row, ref_col)) = self.formula_ref_cell {
            if let Some((end_row, end_col)) = self.formula_ref_end {
                let (min_row, max_row) = (ref_row.min(end_row), ref_row.max(end_row));
                let (min_col, max_col) = (ref_col.min(end_col), ref_col.max(end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    if row == min_row { top = true; }
                    if row == max_row { bottom = true; }
                    if col == min_col { left = true; }
                    if col == max_col { right = true; }
                }
            } else {
                // Single cell - all borders
                if row == ref_row && col == ref_col {
                    top = true; right = true; bottom = true; left = true;
                }
            }
        }

        // Check the highlighted refs from parsed formula
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    if row == fref.start.0 { top = true; }
                    if row == end_row { bottom = true; }
                    if col == fref.start.1 { left = true; }
                    if col == end_col { right = true; }
                }
            } else {
                // Single cell - all borders
                if row == fref.start.0 && col == fref.start.1 {
                    top = true; right = true; bottom = true; left = true;
                }
            }
        }

        (top, right, bottom, left)
    }

    // ========================================================================
    // Cell Reference Parsing
    // ========================================================================

    /// Parse cell reference like "A1", "B25", "AA100"
    pub(crate) fn parse_cell_ref(input: &str) -> Option<(usize, usize)> {
        let input = input.trim().to_uppercase();
        if input.is_empty() {
            return None;
        }

        // Find where letters end and numbers begin
        let letter_end = input.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 || letter_end == input.len() {
            return None;
        }

        let letters = &input[..letter_end];
        let numbers = &input[letter_end..];

        // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
        let col = letters.chars().fold(0usize, |acc, c| {
            acc * 26 + (c as usize - 'A' as usize + 1)
        }) - 1;

        // Parse row (1-based to 0-based)
        let row = numbers.parse::<usize>().ok()?.checked_sub(1)?;

        Some((row, col))
    }

    /// Parse all cell references from a formula WITHOUT color assignment.
    /// Returns raw parsed refs sorted by text position, for use with assign_ref_colors().
    pub(crate) fn parse_formula_refs_raw(formula: &str) -> Vec<(RefKey, (usize, usize), Option<(usize, usize)>, std::ops::Range<usize>)> {
        if !formula.starts_with('=') && !formula.starts_with('+') {
            return Vec::new();
        }

        let tokens = tokenize_for_highlight(formula);
        // Collect raw refs with text ranges: (RefKey, start, end, text_byte_range)
        let mut parsed_refs: Vec<(RefKey, (usize, usize), Option<(usize, usize)>, std::ops::Range<usize>)> = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let (range, token_type) = &tokens[i];

            if *token_type == TokenType::CellRef {
                // Skip cross-sheet references (preceded by Bang token, e.g., Sheet1!A1)
                // These reference cells on another sheet and should not be highlighted
                // on the current sheet.
                let is_cross_sheet = {
                    let mut j = i;
                    let mut found_bang = false;
                    while j > 0 {
                        j -= 1;
                        let (_, prev_type) = &tokens[j];
                        if *prev_type == TokenType::Whitespace {
                            continue; // skip whitespace
                        }
                        found_bang = *prev_type == TokenType::Bang;
                        break;
                    }
                    found_bang
                };

                if is_cross_sheet {
                    // Skip this CellRef and any following :CellRef range
                    if i + 2 < tokens.len() {
                        let (_, next_type) = &tokens[i + 1];
                        let (_, next_next_type) = &tokens[i + 2];
                        if *next_type == TokenType::Colon && *next_next_type == TokenType::CellRef {
                            i += 3; // Skip the whole cross-sheet range
                            continue;
                        }
                    }
                    i += 1;
                    continue;
                }

                // Convert char indices to byte indices for safe slicing
                let byte_start = char_index_to_byte_offset(formula, range.start);
                let byte_end = char_index_to_byte_offset(formula, range.end);
                let cell_text = &formula[byte_start..byte_end];
                // Strip any $ signs for absolute references
                let cell_text_clean: String = cell_text.chars().filter(|c| *c != '$').collect();

                if let Some(start_cell) = Self::parse_cell_ref(&cell_text_clean) {
                    // Check if next tokens form a range (: followed by CellRef)
                    if i + 2 < tokens.len() {
                        let (_, next_type) = &tokens[i + 1];
                        let (range2, next_next_type) = &tokens[i + 2];

                        if *next_type == TokenType::Colon && *next_next_type == TokenType::CellRef {
                            // Convert char indices to byte indices for safe slicing
                            let byte_start2 = char_index_to_byte_offset(formula, range2.start);
                            let byte_end2 = char_index_to_byte_offset(formula, range2.end);
                            let end_text = &formula[byte_start2..byte_end2];
                            let end_text_clean: String = end_text.chars().filter(|c| *c != '$').collect();

                            if let Some(end_cell) = Self::parse_cell_ref(&end_text_clean) {
                                // Normalize range to min/max for stable RefKey
                                let r1 = start_cell.0.min(end_cell.0);
                                let c1 = start_cell.1.min(end_cell.1);
                                let r2 = start_cell.0.max(end_cell.0);
                                let c2 = start_cell.1.max(end_cell.1);
                                let key = RefKey::new(r1, c1, r2, c2);  // collapses A1:A1 to Cell
                                let text_byte_range = range.start..range2.end;
                                // For collapsed single-cell ranges, end should be None
                                let end = if r1 == r2 && c1 == c2 { None } else { Some((r2, c2)) };
                                parsed_refs.push((key, (r1, c1), end, text_byte_range));
                                i += 3;  // Skip the whole range
                                continue;
                            }
                        }
                    }
                    // Single cell reference
                    let key = RefKey::Cell { row: start_cell.0, col: start_cell.1 };
                    parsed_refs.push((key, start_cell, None, range.clone()));
                }
            }
            i += 1;
        }

        // Sort by text position (left-to-right in formula) for deterministic color assignment
        parsed_refs.sort_by_key(|(_, _, _, text_byte_range)| text_byte_range.start);
        parsed_refs
    }

    /// Assign colors to parsed refs using a persistent color map.
    /// Refs that were previously seen keep their colors; new refs get the next available color.
    /// Also performs garbage collection - removes refs from map that are no longer in the formula.
    pub(crate) fn assign_ref_colors(
        parsed: Vec<(RefKey, (usize, usize), Option<(usize, usize)>, std::ops::Range<usize>)>,
        color_map: &mut HashMap<RefKey, usize>,
        next_color: &mut usize,
    ) -> Vec<FormulaRef> {
        use std::collections::HashSet;

        // Garbage collect: remove keys no longer present in formula
        let present_keys: HashSet<RefKey> = parsed.iter().map(|(key, _, _, _)| key.clone()).collect();
        color_map.retain(|k, _| present_keys.contains(k));

        // Assign colors: existing refs keep their colors, new refs get next available
        parsed.into_iter().map(|(key, start, end, text_byte_range)| {
            let color_index = *color_map.entry(key.clone()).or_insert_with(|| {
                let c = *next_color;
                *next_color = (*next_color + 1) % 8;
                c
            });
            FormulaRef { key, start, end, color_index, text_byte_range }
        }).collect()
    }

    /// Parse and assign colors using persistent map from self.
    /// This is the main entry point for formula editing.
    pub(crate) fn update_formula_refs(&mut self) {
        let parsed = Self::parse_formula_refs_raw(&self.edit_value);
        self.formula_highlighted_refs = Self::assign_ref_colors(
            parsed,
            &mut self.formula_ref_color_map,
            &mut self.formula_ref_next_color,
        );
    }

    /// Parse and assign colors with a fresh (non-persistent) color map.
    /// Used for non-editing contexts (e.g., formula bar cache, read-only display).
    pub(crate) fn parse_formula_refs(formula: &str) -> Vec<FormulaRef> {
        let parsed = Self::parse_formula_refs_raw(formula);
        let mut color_map: HashMap<RefKey, usize> = HashMap::new();
        let mut next_color = 0;
        Self::assign_ref_colors(parsed, &mut color_map, &mut next_color)
    }

    /// Clear the persistent color map (call when edit session ends).
    pub(crate) fn clear_formula_ref_colors(&mut self) {
        self.formula_ref_color_map.clear();
        self.formula_ref_next_color = 0;
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn duplicate_token_stability() {
        // =A1 + A1 → two entries, same key, same color
        let refs = Spreadsheet::parse_formula_refs("=A1 + A1");
        assert_eq!(refs.len(), 2, "Both occurrences should be kept");
        assert_eq!(refs[0].key, refs[1].key, "Same cell = same key");
        assert_eq!(refs[0].color_index, refs[1].color_index, "Same key = same color");
    }

    #[test]
    fn order_defines_color() {
        // =B1 + A1 → B1 gets 0 (first in text), A1 gets 1
        let refs = Spreadsheet::parse_formula_refs("=B1 + A1");
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0].color_index, 0, "B1 first in text = color 0");
        assert_eq!(refs[1].color_index, 1, "A1 second in text = color 1");
    }

    #[test]
    fn range_normalization() {
        // =B2:A1 and =A1:B2 produce same key (normalized to min/max)
        let refs1 = Spreadsheet::parse_formula_refs("=B2:A1");
        let refs2 = Spreadsheet::parse_formula_refs("=A1:B2");
        assert_eq!(refs1[0].key, refs2[0].key, "Ranges should normalize to same key");
    }

    #[test]
    fn single_cell_range_collapse() {
        // =A1:A1 collapses to Cell key, same as =A1
        let refs1 = Spreadsheet::parse_formula_refs("=A1");
        let refs2 = Spreadsheet::parse_formula_refs("=A1:A1");
        assert_eq!(refs1[0].key, refs2[0].key, "A1:A1 should collapse to A1");
        assert!(matches!(refs1[0].key, RefKey::Cell { .. }), "=A1 should be Cell");
        assert!(matches!(refs2[0].key, RefKey::Cell { .. }), "=A1:A1 should collapse to Cell");
    }

    #[test]
    fn persistent_map_no_jump() {
        // Simulate incremental typing - colors must not jump
        let mut color_map: HashMap<RefKey, usize> = HashMap::new();
        let mut next_color = 0usize;

        // Type "=A1" → A1 gets color 0
        let parsed1 = Spreadsheet::parse_formula_refs_raw("=A1");
        let refs1 = Spreadsheet::assign_ref_colors(parsed1, &mut color_map, &mut next_color);
        assert_eq!(refs1[0].color_index, 0, "A1 should get color 0");

        // Type "=A1+B1" → A1 stays 0, B1 gets 1
        let parsed2 = Spreadsheet::parse_formula_refs_raw("=A1+B1");
        let refs2 = Spreadsheet::assign_ref_colors(parsed2, &mut color_map, &mut next_color);
        assert_eq!(refs2[0].color_index, 0, "A1 should stay 0");
        assert_eq!(refs2[1].color_index, 1, "B1 should get 1");

        // Type "=A1+B1+C1" → A1 stays 0, B1 stays 1, C1 gets 2
        let parsed3 = Spreadsheet::parse_formula_refs_raw("=A1+B1+C1");
        let refs3 = Spreadsheet::assign_ref_colors(parsed3, &mut color_map, &mut next_color);
        assert_eq!(refs3[0].color_index, 0, "A1 should stay 0");
        assert_eq!(refs3[1].color_index, 1, "B1 should stay 1");
        assert_eq!(refs3[2].color_index, 2, "C1 should get 2");
    }
}
