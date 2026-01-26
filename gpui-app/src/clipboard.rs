//! Clipboard operations for Spreadsheet.
//!
//! This module contains:
//! - InternalClipboard struct for tracking copied cell data
//! - Copy, cut, paste operations
//! - Paste Values (computed values only, no formulas)
//! - Delete selection

use gpui::*;
use visigrid_engine::formula::eval::Value;
use visigrid_engine::provenance::{MutationOp, PasteMode, ClearMode};

use crate::app::Spreadsheet;
use crate::history::CellChange;

/// Maximum rows in the spreadsheet
const NUM_ROWS: usize = 1_000_000;
/// Maximum columns in the spreadsheet
const NUM_COLS: usize = 16_384;

/// Internal clipboard for tracking copied cell data.
/// Stores both raw formulas (for normal paste) and typed values (for paste values).
#[derive(Debug, Clone)]
pub struct InternalClipboard {
    /// Tab-separated raw values (formulas/text) for normal paste + system clipboard
    pub raw_tsv: String,
    /// Typed computed values for Paste Values (2D grid aligned to copied rectangle)
    pub values: Vec<Vec<Value>>,
    /// Top-left cell position of the copied region (for reference adjustment)
    pub source: (usize, usize),
    /// Unique ID written to clipboard metadata for reliable internal detection.
    /// On paste, we check if clipboard metadata contains this ID to distinguish
    /// internal copies from external clipboard content (even if text matches).
    pub id: u128,
}

impl Spreadsheet {
    // Clipboard
    pub fn copy(&mut self, cx: &mut Context<Self>) {
        // If editing, copy selected text (or all if no selection)
        // This is text-only copy, not cell copy - no internal clipboard needed
        if self.mode.is_editing() {
            let text = if let Some((start_byte, end_byte)) = self.edit_selection_range() {
                // Byte-indexed selection
                let start = start_byte.min(self.edit_value.len());
                let end = end_byte.min(self.edit_value.len());
                self.edit_value[start..end].to_string()
            } else {
                self.edit_value.clone()
            };
            self.internal_clipboard = None;  // Text copy, not cell copy
            cx.write_to_clipboard(ClipboardItem::new_string(text));
            self.status_message = Some("Copied to clipboard".to_string());
            cx.notify();
            return;
        }

        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Build tab-separated raw values (formulas) for system clipboard and normal paste
        let mut raw_tsv = String::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if col > min_col {
                    raw_tsv.push('\t');
                }
                raw_tsv.push_str(&self.sheet().get_raw(row, col));
            }
            if row < max_row {
                raw_tsv.push('\n');
            }
        }

        // Build 2D grid of typed computed values for Paste Values
        let mut values = Vec::new();
        for row in min_row..=max_row {
            let mut row_values = Vec::new();
            for col in min_col..=max_col {
                row_values.push(self.sheet().get_computed_value(row, col));
            }
            values.push(row_values);
        }

        // Generate unique nonce for clipboard matching
        let id: u128 = rand::random();

        self.internal_clipboard = Some(InternalClipboard {
            raw_tsv: raw_tsv.clone(),
            values,
            source: (min_row, min_col),
            id,
        });
        // Write clipboard with metadata ID for reliable internal detection
        let id_json = format!("\"{}\"", id);
        cx.write_to_clipboard(ClipboardItem::new_string_with_json_metadata(raw_tsv, id_json));
        self.status_message = Some("Copied to clipboard".to_string());
        cx.notify();
    }

    pub fn cut(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        self.copy(cx);

        // Clear the selected cells and record history
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let mut changes = Vec::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().set_value(row, col, "");
            }
        }
        self.history.record_batch(self.sheet_index(), changes);
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        self.status_message = Some("Cut to clipboard".to_string());
        cx.notify();
    }

    pub fn paste(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // If editing, paste into the edit buffer instead
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        // Read clipboard item to get both text and metadata
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        // Check if clipboard matches our internal clipboard via metadata ID (primary)
        // Fall back to normalized string comparison only if metadata absent (legacy)
        let is_internal = self.internal_clipboard.as_ref().map_or(false, |ic| {
            let expected_id = format!("\"{}\"", ic.id);
            if let Some(ref m) = metadata {
                m == &expected_id
            } else {
                // Legacy fallback: normalized string compare when metadata missing
                system_text.as_ref().map_or(false, |st| {
                    Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
                })
            }
        });

        // Get the text to paste (prefer system clipboard for interop)
        let text = system_text.or_else(|| self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone()));

        if let Some(text) = text {
            let (start_row, start_col) = self.view_state.selected;
            let mut changes = Vec::new();

            // Calculate delta from source if this is an internal paste
            let (delta_row, delta_col) = if is_internal {
                if let Some(ic) = &self.internal_clipboard {
                    let (src_row, src_col) = ic.source;
                    (start_row as i32 - src_row as i32, start_col as i32 - src_col as i32)
                } else {
                    (0, 0)
                }
            } else {
                (0, 0)  // External clipboard - no adjustment
            };

            // Parse tab-separated values and build values grid for provenance
            let mut values_grid: Vec<Vec<String>> = Vec::new();
            let mut end_row = start_row;
            let mut end_col = start_col;
            for (row_offset, line) in text.lines().enumerate() {
                let mut row_values: Vec<String> = Vec::new();
                for (col_offset, value) in line.split('\t').enumerate() {
                    let row = start_row + row_offset;
                    let col = start_col + col_offset;
                    if row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet().get_raw(row, col);

                        // Adjust formula references if this is a formula and we have internal source
                        let new_value = if value.starts_with('=') && is_internal {
                            self.adjust_formula_refs(value, delta_row, delta_col)
                        } else {
                            value.to_string()
                        };

                        row_values.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.sheet_mut().set_value(row, col, &new_value);

                        // Track paste bounds
                        end_row = end_row.max(row);
                        end_col = end_col.max(col);
                    }
                }
                if !row_values.is_empty() {
                    values_grid.push(row_values);
                }
            }

            // Record with provenance (only if changes were made)
            if !changes.is_empty() {
                let provenance = MutationOp::Paste {
                    sheet: self.sheet().id,
                    dst_row: start_row,
                    dst_col: start_col,
                    values: values_grid,
                    mode: PasteMode::Both,  // Regular paste includes formulas
                }.to_provenance(&self.sheet().name);

                self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
                self.bump_cells_rev();  // Invalidate cell search cache
                self.is_modified = true;
            }

            // Validate pasted range and report failures
            let failures = self.workbook.validate_range(
                self.sheet_index(), start_row, start_col, end_row, end_col
            );
            let total_cells = (end_row - start_row + 1) * (end_col - start_col + 1);
            if failures.count > 0 {
                self.store_validation_failures(&failures);
                self.status_message = Some(format!(
                    "Pasted from clipboard (Validation: {} of {} cells failed) — Press F8 to jump",
                    failures.count, total_cells
                ));
            } else {
                self.status_message = Some("Pasted from clipboard".to_string());
            }

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();

            cx.notify();
        }
    }

    /// Normalize clipboard text for comparison (handles line ending differences)
    pub(crate) fn normalize_clipboard_text(text: &str) -> String {
        text.replace("\r\n", "\n").trim_end().to_string()
    }

    /// Paste clipboard text into the edit buffer (when in editing mode)
    pub fn paste_into_edit(&mut self, cx: &mut Context<Self>) {
        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone())
        };

        if let Some(text) = text {
            // Only take first line if multi-line, and trim whitespace
            let text = text.lines().next().unwrap_or("").trim();
            if !text.is_empty() {
                // Insert at cursor byte position
                let byte_pos = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.insert_str(byte_pos, text);
                self.edit_cursor = byte_pos + text.len();  // Advance by byte length

                // Update autocomplete for formulas
                self.update_autocomplete(cx);

                self.edit_scroll_dirty = true;
                self.status_message = Some(format!("Pasted: {}", text));
                cx.notify();
            }
        }
    }

    /// Paste Values: paste computed values only (no formulas).
    /// Uses typed values from internal clipboard, or parses external clipboard with leading-zero guard.
    pub fn paste_values(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // If editing, paste canonical text into edit buffer (top-left cell only)
        if self.mode.is_editing() {
            self.paste_values_into_edit(cx);
            return;
        }

        // Read clipboard item to get text
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));

        // For Paste Values, prefer internal clipboard values if they exist and text matches.
        // This avoids depending on metadata (which doesn't round-trip on Windows).
        // The internal clipboard stores computed values, which is exactly what we want.
        let use_internal_values = self.internal_clipboard.as_ref().map_or(false, |ic| {
            // Use internal values if we have them AND either:
            // 1. System clipboard matches our raw_tsv (we copied it)
            // 2. System clipboard is empty/unavailable (use what we have)
            system_text.as_ref().map_or(true, |st| {
                Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
            })
        });

        let (start_row, start_col) = self.view_state.selected;
        let mut changes = Vec::new();
        let mut values_grid: Vec<Vec<String>> = Vec::new();
        let mut end_row = start_row;
        let mut end_col = start_col;

        if use_internal_values {
            // Use typed values from internal clipboard (clone to avoid borrow issues)
            let values = self.internal_clipboard.as_ref().map(|ic| ic.values.clone());
            if let Some(values) = values {
                for (row_offset, row_values) in values.iter().enumerate() {
                    let mut grid_row: Vec<String> = Vec::new();
                    for (col_offset, value) in row_values.iter().enumerate() {
                        let row = start_row + row_offset;
                        let col = start_col + col_offset;
                        if row < NUM_ROWS && col < NUM_COLS {
                            let old_value = self.sheet().get_raw(row, col);
                            let new_value = Self::value_to_canonical_string(value);

                            grid_row.push(new_value.clone());

                            if old_value != new_value {
                                changes.push(CellChange {
                                    row, col, old_value, new_value: new_value.clone(),
                                });
                            }
                            self.sheet_mut().set_value(row, col, &new_value);

                            end_row = end_row.max(row);
                            end_col = end_col.max(col);
                        }
                    }
                    if !grid_row.is_empty() {
                        values_grid.push(grid_row);
                    }
                }
            }
        } else if let Some(text) = system_text {
            // Parse external clipboard with leading-zero guard
            for (row_offset, line) in text.lines().enumerate() {
                let mut grid_row: Vec<String> = Vec::new();
                for (col_offset, cell_text) in line.split('\t').enumerate() {
                    let row = start_row + row_offset;
                    let col = start_col + col_offset;
                    if row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet().get_raw(row, col);
                        let parsed_value = Self::parse_external_value(cell_text);
                        let new_value = Self::value_to_canonical_string(&parsed_value);

                        grid_row.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.sheet_mut().set_value(row, col, &new_value);

                        end_row = end_row.max(row);
                        end_col = end_col.max(col);
                    }
                }
                if !grid_row.is_empty() {
                    values_grid.push(grid_row);
                }
            }
        }

        if !changes.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet().id,
                dst_row: start_row,
                dst_col: start_col,
                values: values_grid,
                mode: PasteMode::Values,
            }.to_provenance(&self.sheet().name);

            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        // Validate pasted range and report failures
        let failures = self.workbook.validate_range(
            self.sheet_index(), start_row, start_col, end_row, end_col
        );
        let total_cells = (end_row - start_row + 1) * (end_col - start_col + 1);
        if failures.count > 0 {
            self.store_validation_failures(&failures);
            self.status_message = Some(format!(
                "Pasted values (Validation: {} of {} cells failed) — Press F8 to jump",
                failures.count, total_cells
            ));
        } else {
            self.status_message = Some("Pasted values".to_string());
        }
        cx.notify();
    }

    /// Convert a typed Value to its canonical string representation for cell storage.
    /// Guarantees: no scientific notation, deterministic output, -0.0 normalized to 0.
    pub(crate) fn value_to_canonical_string(value: &Value) -> String {
        match value {
            Value::Empty => String::new(),
            Value::Number(n) => {
                // Handle non-finite values explicitly
                if !n.is_finite() {
                    if n.is_nan() { return "NaN".to_string(); }
                    return if *n > 0.0 { "INF".to_string() } else { "-INF".to_string() };
                }

                // Normalize -0.0 to 0.0
                let n0 = if *n == 0.0 { 0.0 } else { *n };

                // Integer fast path: no decimal point needed
                if n0.fract() == 0.0 && n0.abs() < 9e15 {
                    format!("{:.0}", n0)
                } else {
                    // Fixed precision (15 decimals), trim trailing zeros, no scientific notation
                    let mut s = format!("{:.15}", n0);
                    while s.contains('.') && s.ends_with('0') { s.pop(); }
                    if s.ends_with('.') { s.pop(); }
                    s
                }
            }
            Value::Text(s) => s.clone(),
            Value::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            Value::Error(e) => e.clone(),
        }
    }

    /// Parse external clipboard text into a typed Value with leading-zero preservation.
    pub(crate) fn parse_external_value(text: &str) -> Value {
        let trimmed = text.trim();

        if trimmed.is_empty() {
            return Value::Empty;
        }

        // Check for formula prefix - treat as literal text (strip the =)
        if trimmed.starts_with('=') {
            return Value::Text(trimmed.to_string());
        }

        // Check for leading zeros that should be preserved as text
        // e.g., "007", "00123" - but not "0" or "0.5"
        if trimmed.starts_with('0') && trimmed.len() > 1 {
            let second_char = trimmed.chars().nth(1).unwrap();
            if second_char.is_ascii_digit() {
                // Starts with 0 followed by digit -> preserve as text
                return Value::Text(trimmed.to_string());
            }
        }

        // Check for boolean
        let upper = trimmed.to_uppercase();
        if upper == "TRUE" {
            return Value::Boolean(true);
        }
        if upper == "FALSE" {
            return Value::Boolean(false);
        }

        // Try to parse as number
        if let Ok(n) = trimmed.parse::<f64>() {
            return Value::Number(n);
        }

        // Default to text
        Value::Text(trimmed.to_string())
    }

    /// Paste values into edit buffer: use canonical text of top-left value only.
    fn paste_values_into_edit(&mut self, cx: &mut Context<Self>) {
        // Read clipboard item to get text
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));

        // For Paste Values, prefer internal clipboard values if they exist and text matches.
        // This avoids depending on metadata (which doesn't round-trip on Windows).
        let use_internal_values = self.internal_clipboard.as_ref().map_or(false, |ic| {
            system_text.as_ref().map_or(true, |st| {
                Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
            })
        });

        let text = if use_internal_values {
            // Get top-left value from internal clipboard
            self.internal_clipboard.as_ref().and_then(|ic| {
                ic.values.first().and_then(|row| row.first()).map(|v| Self::value_to_canonical_string(v))
            })
        } else {
            // Parse top-left cell from external clipboard
            system_text.map(|text| {
                let first_cell = text.lines().next().unwrap_or("")
                    .split('\t').next().unwrap_or("");
                let value = Self::parse_external_value(first_cell);
                Self::value_to_canonical_string(&value)
            })
        };

        if let Some(text) = text {
            if !text.is_empty() {
                // Insert at cursor byte position
                let byte_pos = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.insert_str(byte_pos, &text);
                self.edit_cursor = byte_pos + text.len();  // Advance by byte length

                self.update_autocomplete(cx);
                self.edit_scroll_dirty = true;
                self.status_message = Some(format!("Pasted value: {}", text));
                cx.notify();
            }
        }
    }

    pub fn delete_selection(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        let mut changes = Vec::new();
        let mut skipped_spill_receivers = false;

        // Delete from all selection ranges (including discontiguous Ctrl+Click selections)
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            // Only get cells that actually have data (efficient for large selections)
            let cells_to_delete = self.sheet().cells_in_range(min_row, max_row, min_col, max_col);

            for (row, col) in cells_to_delete {
                // Skip spill receivers - only the parent formula can be deleted
                if self.sheet().is_spill_receiver(row, col) {
                    skipped_spill_receivers = true;
                    continue;
                }

                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().clear_cell(row, col);
            }
        }

        let had_changes = !changes.is_empty();
        if had_changes {
            // Only attach provenance for single contiguous selection
            let provenance = if self.view_state.additional_selections.is_empty() {
                let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
                Some(MutationOp::Clear {
                    sheet: self.sheet().id,
                    start_row: min_row,
                    start_col: min_col,
                    end_row: max_row,
                    end_col: max_col,
                    mode: ClearMode::All,
                }.to_provenance(&self.sheet().name))
            } else {
                None  // Discontiguous selection - no provenance
            };
            self.history.record_batch_with_provenance(self.sheet_index(), changes, provenance);
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;
        }

        if skipped_spill_receivers && !had_changes {
            self.status_message = Some("Cannot delete spill range. Delete the parent formula instead.".to_string());
        }

        cx.notify();
    }
}
