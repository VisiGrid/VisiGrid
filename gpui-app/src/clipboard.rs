//! Clipboard operations for Spreadsheet.
//!
//! This module contains:
//! - InternalClipboard struct for tracking copied cell data
//! - Copy, cut, paste operations
//! - Paste Values (computed values only, no formulas)
//! - Delete selection

use gpui::*;
use visigrid_engine::cell::CellFormat;
use visigrid_engine::formula::eval::Value;
use visigrid_engine::provenance::{MutationOp, PasteMode, ClearMode};
use visigrid_engine::sheet::MergedRegion;

use crate::app::Spreadsheet;
use crate::history::{CellChange, CellFormatPatch, FormatActionKind, UndoAction};

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
    /// Cell formats for Paste Formats (2D grid with same dimensions as values)
    /// Every position gets a CellFormat, even if default (rectangular, not sparse).
    pub formats: Vec<Vec<CellFormat>>,
    /// Top-left cell position of the copied region (for reference adjustment)
    pub source: (usize, usize),
    /// Unique ID written to clipboard metadata for reliable internal detection.
    /// On paste, we check if clipboard metadata contains this ID to distinguish
    /// internal copies from external clipboard content (even if text matches).
    pub id: u128,
    /// Merged regions from the copied area, stored with coordinates
    /// relative to the clipboard's top-left (0,0).
    /// Empty when copied from a filtered view.
    pub merges: Vec<MergedRegion>,
    /// When this clipboard entry was created (for time-bounded Wayland fallback)
    pub created_at: std::time::Instant,
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
        let is_filtered = self.row_view.is_filtered();

        // Normalize selection to include full merge regions (only when not filtered)
        let (min_row, min_col, max_row, max_col) = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut nr = min_row;
            let mut nc = min_col;
            let mut xr = max_row;
            let mut xc = max_col;
            // Expand to include any intersecting merges (one pass suffices since merges don't overlap)
            for merge in &sheet.merged_regions {
                let intersects = merge.end.0 >= nr && merge.start.0 <= xr
                              && merge.end.1 >= nc && merge.start.1 <= xc;
                if intersects {
                    nr = nr.min(merge.start.0);
                    nc = nc.min(merge.start.1);
                    xr = xr.max(merge.end.0);
                    xc = xc.max(merge.end.1);
                }
            }
            (nr, nc, xr, xc)
        } else {
            (min_row, min_col, max_row, max_col)
        };

        // Build tab-separated raw values (formulas) for system clipboard and normal paste
        // When filtered, only include visible rows
        let mut raw_tsv = String::new();
        let mut values = Vec::new();
        let mut formats = Vec::new();
        let mut first_row = true;
        let mut source_row = min_row; // Track first visible row for source

        for view_row in min_row..=max_row {
            // Skip hidden rows when filtered
            if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                continue;
            }

            // Convert view row to data row for sheet access
            let data_row = self.row_view.view_to_data(view_row);

            if first_row {
                source_row = view_row;
                first_row = false;
            } else {
                raw_tsv.push('\n');
            }

            let mut row_values = Vec::new();
            let mut row_formats = Vec::new();
            for col in min_col..=max_col {
                if col > min_col {
                    raw_tsv.push('\t');
                }
                raw_tsv.push_str(&self.sheet(cx).get_raw(data_row, col));
                row_values.push(self.sheet(cx).get_computed_value(data_row, col));
                // Capture format for every cell position (rectangular, not sparse)
                row_formats.push(self.sheet(cx).get_format(data_row, col).clone());
            }
            values.push(row_values);
            formats.push(row_formats);
        }

        // Capture merge metadata (only when not filtered)
        let merges = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut relative_merges = Vec::new();
            for merge in &sheet.merged_regions {
                // After normalization, all intersecting merges are fully contained
                let contained = merge.start.0 >= min_row && merge.end.0 <= max_row
                              && merge.start.1 >= min_col && merge.end.1 <= max_col;
                if contained {
                    relative_merges.push(MergedRegion::new(
                        merge.start.0 - min_row,
                        merge.start.1 - min_col,
                        merge.end.0 - min_row,
                        merge.end.1 - min_col,
                    ));
                }
            }
            relative_merges
        } else {
            Vec::new()
        };

        // Generate unique nonce for clipboard matching
        let id: u128 = rand::random();

        self.internal_clipboard = Some(InternalClipboard {
            raw_tsv: raw_tsv.clone(),
            values,
            formats,
            source: (source_row, min_col),
            id,
            merges,
            created_at: std::time::Instant::now(),
        });
        // Write clipboard with metadata ID for reliable internal detection
        let id_json = format!("\"{}\"", id);
        cx.write_to_clipboard(ClipboardItem::new_string_with_json_metadata(raw_tsv, id_json));

        // Set visual range for dashed border overlay
        self.clipboard_visual_range = Some((min_row, min_col, max_row, max_col));

        if is_filtered {
            self.status_message = Some("Copied visible rows to clipboard".to_string());
        } else {
            self.status_message = Some("Copied to clipboard".to_string());
        }
        cx.notify();
    }

    pub fn cut(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        self.copy(cx);

        // Clear the selected cells and record history (visible rows only when filtered)
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let is_filtered = self.row_view.is_filtered();

        // Normalize selection to include full merge regions (same expansion as copy)
        let (min_row, min_col, max_row, max_col) = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut nr = min_row;
            let mut nc = min_col;
            let mut xr = max_row;
            let mut xc = max_col;
            for merge in &sheet.merged_regions {
                let intersects = merge.end.0 >= nr && merge.start.0 <= xr
                              && merge.end.1 >= nc && merge.start.1 <= xc;
                if intersects {
                    nr = nr.min(merge.start.0);
                    nc = nc.min(merge.start.1);
                    xr = xr.max(merge.end.0);
                    xc = xc.max(merge.end.1);
                }
            }
            (nr, nc, xr, xc)
        } else {
            (min_row, min_col, max_row, max_col)
        };

        let mut changes = Vec::new();

        self.wb_mut(cx, |wb| wb.begin_batch());
        for view_row in min_row..=max_row {
            // Skip hidden rows when filtered
            if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                continue;
            }

            // Convert view row to data row for sheet access
            let data_row = self.row_view.view_to_data(view_row);

            for col in min_col..=max_col {
                let old_value = self.sheet(cx).get_raw(data_row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row: data_row, col, old_value, new_value: String::new(),
                    });
                }
                self.set_cell_value(data_row, col, "", cx);
            }
        }
        self.end_batch_and_broadcast(cx);

        // Remove merges fully within the cut selection (move semantics, not filtered)
        let mut removed_any = false;
        let mut merges_before = Vec::new();
        let mut merges_after = Vec::new();

        if !is_filtered {
            merges_before = self.sheet(cx).merged_regions.clone();
            let origins_to_remove: Vec<(usize, usize)> = {
                let sheet = self.sheet(cx);
                sheet.merged_regions.iter()
                    .filter(|merge| {
                        merge.start.0 >= min_row && merge.end.0 <= max_row
                        && merge.start.1 >= min_col && merge.end.1 <= max_col
                    })
                    .map(|merge| merge.start)
                    .collect()
            };
            removed_any = !origins_to_remove.is_empty();
            for origin in origins_to_remove {
                self.active_sheet_mut(cx, |s| { let _ = s.remove_merge(origin); });
            }
            merges_after = self.sheet(cx).merged_regions.clone();
        }

        // Record history: Group if merges were removed, otherwise simple batch
        if removed_any {
            let sheet_index = self.sheet_index(cx);
            let merge_action = UndoAction::SetMerges {
                sheet_index,
                before: merges_before,
                after: merges_after,
                cleared_values: vec![],
                description: "Cut: remove source merges".to_string(),
            };
            let values_action = UndoAction::Values { sheet_index, changes };
            // Order matters: redo applies forward, undo applies reverse;
            // values must precede merges on redo.
            self.history.record_action_with_provenance(
                UndoAction::Group {
                    actions: vec![values_action, merge_action],
                    description: "Cut".to_string(),
                },
                None,
            );
        } else {
            self.history.record_batch(self.sheet_index(cx), changes);
        }

        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;

        if is_filtered {
            // Check if there were merges in the cut range that we didn't remove
            let has_merges_in_range = self.sheet(cx).merged_regions.iter().any(|m| {
                m.end.0 >= min_row && m.start.0 <= max_row
                && m.end.1 >= min_col && m.start.1 <= max_col
            });
            if has_merges_in_range {
                self.status_message = Some("Cut visible rows (merged regions not moved in filtered view)".to_string());
            } else {
                self.status_message = Some("Cut visible rows to clipboard".to_string());
            }
        } else {
            self.status_message = Some("Cut to clipboard".to_string());
        }
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

        // Determine if this is an internal paste (with formula adjustment) or external
        let is_internal = Self::is_internal_paste(
            self.internal_clipboard.as_ref(),
            system_text.as_deref(),
            metadata.as_deref(),
        );

        #[cfg(debug_assertions)]
        {
            let text_match = system_text.as_deref().map_or(false, |st| {
                self.internal_clipboard.as_ref().map_or(false, |ic| {
                    Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
                })
            });
            eprintln!("[paste] is_internal={}, metadata={:?}, text_match={}", is_internal, metadata.is_some(), text_match);
        }

        // Get the text to paste (prefer system clipboard for interop)
        let text = system_text.or_else(|| self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone()));

        if let Some(text) = text {
            let (start_row, start_col) = self.view_state.selected;
            let is_filtered = self.row_view.is_filtered();
            let mut changes = Vec::new();

            // Check if clipboard is a single cell (1 line, no tabs)
            let lines: Vec<&str> = text.lines().collect();
            let is_single_cell = lines.len() == 1 && !lines[0].contains('\t');

            // If single cell and multi-selection, broadcast to all selected cells
            if is_single_cell && self.is_multi_selection() {
                let single_value = lines[0].to_string();
                let primary_cell = self.view_state.selected;
                let primary_data_row = self.row_view.view_to_data(primary_cell.0);

                // Source cell position for formula rebasing (delta = target - source)
                let (src_data_row, src_col) = if is_internal {
                    if let Some(ic) = &self.internal_clipboard {
                        (self.row_view.view_to_data(ic.source.0) as i32, ic.source.1 as i32)
                    } else {
                        (primary_data_row as i32, primary_cell.1 as i32)
                    }
                } else {
                    (primary_data_row as i32, primary_cell.1 as i32)
                };

                // Collect all target cells (view_row, col) -> (data_row, col)
                let mut target_cells: Vec<(usize, usize)> = Vec::new();

                // Primary selection rectangle (filter to visible rows)
                let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
                for view_row in min_row..=max_row {
                    if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                        continue;
                    }
                    let data_row = self.row_view.view_to_data(view_row);
                    for col in min_col..=max_col {
                        target_cells.push((data_row, col));
                    }
                }

                // Additional selections (Ctrl+Click) - filter to visible rows
                for (sel_start, sel_end) in &self.view_state.additional_selections {
                    let end = sel_end.unwrap_or(*sel_start);
                    let min_r = sel_start.0.min(end.0);
                    let max_r = sel_start.0.max(end.0);
                    let min_c = sel_start.1.min(end.1);
                    let max_c = sel_start.1.max(end.1);
                    for view_row in min_r..=max_r {
                        if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                            continue;
                        }
                        let data_row = self.row_view.view_to_data(view_row);
                        for col in min_c..=max_c {
                            if !target_cells.contains(&(data_row, col)) {
                                target_cells.push((data_row, col));
                            }
                        }
                    }
                }

                let is_formula = single_value.starts_with('=');
                let mut values_grid: Vec<Vec<String>> = Vec::new();

                self.wb_mut(cx, |wb| wb.begin_batch());
                for (data_row, col) in &target_cells {
                    let old_value = self.sheet(cx).get_raw(*data_row, *col);

                    // For formulas, shift relative references based on delta from source cell
                    let new_value = if is_formula && is_internal {
                        let delta_row = *data_row as i32 - src_data_row;
                        let delta_col = *col as i32 - src_col;
                        self.adjust_formula_refs(&single_value, delta_row, delta_col)
                    } else {
                        single_value.clone()
                    };

                    if old_value != new_value {
                        changes.push(CellChange {
                            row: *data_row, col: *col, old_value, new_value: new_value.clone(),
                        });
                    }
                    self.set_cell_value(*data_row, *col, &new_value, cx);
                }
                self.end_batch_and_broadcast(cx);

                // Build values grid for provenance
                if !target_cells.is_empty() {
                    values_grid.push(vec![single_value.clone()]);
                }

                // Record with provenance
                if !changes.is_empty() {
                    let data_start_row = self.row_view.view_to_data(start_row);
                    let provenance = MutationOp::Paste {
                        sheet: self.sheet(cx).id,
                        dst_row: data_start_row,
                        dst_col: start_col,
                        values: values_grid,
                        mode: PasteMode::Both,
                    }.to_provenance(&self.sheet(cx).name);

                    self.history.record_batch_with_provenance(self.sheet_index(cx), changes, Some(provenance));
                    self.bump_cells_rev();
                    self.is_modified = true;
                }

                self.clipboard_visual_range = None;
                self.status_message = Some(format!("Pasted to {} cells", target_cells.len()));
                self.maybe_smoke_recalc(cx);
                cx.notify();
                return;
            }

            // Standard paste (multi-cell clipboard or single cell to single selection)
            // When filtered, paste to consecutive visible rows
            let data_start_row = self.row_view.view_to_data(start_row);

            // Compute paste rectangle for split-merge guard and merge recreation
            let paste_rows = lines.len();
            let paste_cols = lines.iter().map(|l| l.split('\t').count()).max().unwrap_or(1);
            let paste_max_row = (data_start_row + paste_rows).saturating_sub(1);
            let paste_max_col = (start_col + paste_cols).saturating_sub(1);

            // Block if paste would split a merged region
            if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, paste_max_row, paste_max_col, cx) {
                self.status_message = Some(format!(
                    "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                    Self::col_to_letter(mc), mr + 1,
                ));
                cx.notify();
                return;
            }

            // Calculate delta from source if this is an internal paste
            let (delta_row, delta_col) = if is_internal {
                if let Some(ic) = &self.internal_clipboard {
                    let (src_row, src_col) = ic.source;
                    let src_data_row = self.row_view.view_to_data(src_row);
                    (data_start_row as i32 - src_data_row as i32, start_col as i32 - src_col as i32)
                } else {
                    (0, 0)
                }
            } else {
                (0, 0)  // External clipboard - no adjustment
            };

            // For filtered paste: find the starting visible index
            let visible_start_idx = if is_filtered {
                self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
            } else {
                None
            };

            // Parse tab-separated values and build values grid for provenance
            let mut values_grid: Vec<Vec<String>> = Vec::new();
            let mut end_data_row = data_start_row;
            let mut end_col = start_col;

            self.wb_mut(cx, |wb| wb.begin_batch());
            for (row_offset, line) in text.lines().enumerate() {
                // Determine target view row for this clipboard row
                let (_target_view_row, target_data_row) = if is_filtered {
                    if let Some(start_idx) = visible_start_idx {
                        // Get the nth visible row from the starting position
                        if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                            let data_row = self.row_view.view_to_data(view_row);
                            (view_row, data_row)
                        } else {
                            // No more visible rows - skip this line
                            continue;
                        }
                    } else {
                        // Start row not visible - skip
                        continue;
                    }
                } else {
                    // No filtering - direct mapping
                    let view_row = start_row + row_offset;
                    if view_row >= NUM_ROWS {
                        continue;
                    }
                    (view_row, view_row)
                };

                let mut row_values: Vec<String> = Vec::new();
                for (col_offset, value) in line.split('\t').enumerate() {
                    let col = start_col + col_offset;
                    if target_data_row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet(cx).get_raw(target_data_row, col);

                        // Adjust formula references using constant delta from source to destination
                        let new_value = if value.starts_with('=') && is_internal {
                            self.adjust_formula_refs(value, delta_row, delta_col)
                        } else {
                            value.to_string()
                        };

                        row_values.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row: target_data_row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.set_cell_value(target_data_row, col, &new_value, cx);

                        // Track paste bounds (in data coordinates)
                        end_data_row = end_data_row.max(target_data_row);
                        end_col = end_col.max(col);
                    }
                }
                if !row_values.is_empty() {
                    values_grid.push(row_values);
                }
            }

            // Recreate clipboard merges at destination (only for internal paste, not filtered)
            let clipboard_merges = if is_internal {
                self.internal_clipboard.as_ref()
                    .map(|ic| ic.merges.clone())
                    .unwrap_or_default()
            } else {
                Vec::new()
            };

            let mut merge_action = None;
            if !clipboard_merges.is_empty() && !is_filtered {
                let merges_before = self.sheet(cx).merged_regions.clone();

                // Remove existing merges fully within the paste rectangle
                let origins_to_remove: Vec<(usize, usize)> = {
                    let sheet = self.sheet(cx);
                    sheet.merged_regions.iter()
                        .filter(|existing| {
                            existing.start.0 >= data_start_row
                            && existing.end.0 <= paste_max_row
                            && existing.start.1 >= start_col
                            && existing.end.1 <= paste_max_col
                        })
                        .map(|existing| existing.start)
                        .collect()
                };
                for origin in origins_to_remove {
                    self.active_sheet_mut(cx, |s| { let _ = s.remove_merge(origin); });
                }

                // Add clipboard merges at destination offsets
                let mut cleared_values: Vec<(usize, usize, String)> = Vec::new();
                for rel_merge in &clipboard_merges {
                    let dest = MergedRegion::new(
                        data_start_row + rel_merge.start.0,
                        start_col + rel_merge.start.1,
                        data_start_row + rel_merge.end.0,
                        start_col + rel_merge.end.1,
                    );
                    // Clear non-origin cells (same semantics as merge_cells)
                    for r in dest.start.0..=dest.end.0 {
                        for c in dest.start.1..=dest.end.1 {
                            if (r, c) == dest.start { continue; }
                            let raw = self.sheet(cx).get_raw(r, c);
                            if !raw.is_empty() {
                                cleared_values.push((r, c, raw));
                                self.set_cell_value(r, c, "", cx);
                            }
                        }
                    }
                    self.active_sheet_mut(cx, |s| { let _ = s.add_merge(dest); });
                }

                let merges_after = self.sheet(cx).merged_regions.clone();

                merge_action = Some(UndoAction::SetMerges {
                    sheet_index: self.sheet_index(cx),
                    before: merges_before,
                    after: merges_after,
                    cleared_values,
                    description: "Paste: recreate merges".to_string(),
                });
            }
            self.end_batch_and_broadcast(cx);

            // Record with provenance (only if changes or merge changes were made)
            if !changes.is_empty() || merge_action.is_some() {
                let provenance = MutationOp::Paste {
                    sheet: self.sheet(cx).id,
                    dst_row: data_start_row,
                    dst_col: start_col,
                    values: values_grid,
                    mode: PasteMode::Both,
                }.to_provenance(&self.sheet(cx).name);

                if let Some(merge_act) = merge_action {
                    // Order matters: redo applies forward, undo applies reverse;
                    // values must precede merges on redo.
                    let sheet_index = self.sheet_index(cx);
                    let values_action = UndoAction::Values { sheet_index, changes };
                    self.history.record_action_with_provenance(
                        UndoAction::Group {
                            actions: vec![values_action, merge_act],
                            description: "Paste".to_string(),
                        },
                        Some(provenance),
                    );
                } else if !changes.is_empty() {
                    self.history.record_batch_with_provenance(self.sheet_index(cx), changes, Some(provenance));
                }
                self.bump_cells_rev();
                self.is_modified = true;
            }

            // Validate pasted range and report failures (using data coordinates)
            let failures = self.wb(cx).validate_range(
                self.sheet_index(cx), data_start_row, start_col, end_data_row, end_col
            );
            let total_cells = (end_data_row - data_start_row + 1) * (end_col - start_col + 1);
            if failures.count > 0 {
                self.store_validation_failures(&failures);
                self.status_message = Some(format!(
                    "Pasted from clipboard (Validation: {} of {} cells failed) — Press F8 to jump",
                    failures.count, total_cells
                ));
            } else {
                self.status_message = Some("Pasted from clipboard".to_string());
            }

            // Clear copy border overlay — clipboard consumed
            self.clipboard_visual_range = None;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc(cx);

            cx.notify();
        }
    }

    /// Normalize clipboard text for comparison (handles line ending differences)
    pub(crate) fn normalize_clipboard_text(text: &str) -> String {
        // Normalize line endings and trim whitespace from both ends
        // Some clipboard managers add leading/trailing whitespace or transform line endings
        text.replace("\r\n", "\n").replace('\r', "\n").trim().to_string()
    }

    /// Determine if a paste operation should use internal clipboard data with formula adjustment.
    ///
    /// Returns true (internal paste) when:
    /// 1. Clipboard metadata matches internal clipboard ID (reliable cross-platform)
    /// 2. System clipboard text matches internal clipboard text (fallback when metadata unavailable
    ///    or when metadata doesn't match but text does — Wayland may garble metadata)
    /// 3. System clipboard is unavailable AND copy happened recently (< 2s, Wayland failure mode)
    ///
    /// Returns false (external paste) when:
    /// - No internal clipboard exists
    /// - System clipboard has different content (user copied from external source)
    ///
    /// This function is public for testing the Wayland clipboard-unavailable scenario.
    pub fn is_internal_paste(
        internal_clipboard: Option<&InternalClipboard>,
        system_text: Option<&str>,
        metadata: Option<&str>,
    ) -> bool {
        let Some(ic) = internal_clipboard else {
            return false;
        };

        let expected_id = format!("\"{}\"", ic.id);

        // Metadata match: definitive yes
        if let Some(m) = metadata {
            if m == expected_id {
                return true;
            }
            // Metadata exists but doesn't match — likely external.
            // Still check text as defensive fallback (metadata may be garbled on Wayland).
            if let Some(st) = system_text {
                return Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv);
            }
            return false;
        }

        // No metadata: fall back to text comparison
        if let Some(st) = system_text {
            return Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv);
        }

        // System clipboard completely unavailable (Wayland failure mode).
        // Only assume internal if the copy happened recently (< 2s).
        ic.created_at.elapsed() < std::time::Duration::from_secs(2)
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
    /// When filtered, pastes to consecutive visible rows only.
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
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        // Block if paste would split a merged region
        {
            let (paste_rows, paste_cols) = if use_internal_values {
                self.internal_clipboard.as_ref()
                    .map(|ic| (ic.values.len(), ic.values.first().map_or(0, |r| r.len())))
                    .unwrap_or((0, 0))
            } else {
                let text = system_text.as_deref().unwrap_or("");
                let lines: Vec<&str> = text.lines().collect();
                (lines.len(), lines.iter().map(|l| l.split('\t').count()).max().unwrap_or(1))
            };
            if paste_rows > 0 && paste_cols > 0 {
                let dest_max_row = (data_start_row + paste_rows).saturating_sub(1);
                let dest_max_col = (start_col + paste_cols).saturating_sub(1);
                if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, dest_max_row, dest_max_col, cx) {
                    self.status_message = Some(format!(
                        "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                        Self::col_to_letter(mc), mr + 1,
                    ));
                    cx.notify();
                    return;
                }
            }
        }

        let mut changes = Vec::new();
        let mut values_grid: Vec<Vec<String>> = Vec::new();
        let mut end_data_row = data_start_row;
        let mut end_col = start_col;

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        /// Helper to get the nth target row (returns data_row)
        fn get_target_data_row(
            row_view: &visigrid_engine::filter::RowView,
            is_filtered: bool,
            visible_start_idx: Option<usize>,
            start_row: usize,
            row_offset: usize,
        ) -> Option<usize> {
            if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = row_view.nth_visible(start_idx + row_offset) {
                        return Some(row_view.view_to_data(view_row));
                    }
                }
                None
            } else {
                let view_row = start_row + row_offset;
                if view_row < NUM_ROWS {
                    Some(view_row)
                } else {
                    None
                }
            }
        }

        self.wb_mut(cx, |wb| wb.begin_batch());
        if use_internal_values {
            // Use typed values from internal clipboard (clone to avoid borrow issues)
            let values = self.internal_clipboard.as_ref().map(|ic| ic.values.clone());
            if let Some(values) = values {
                for (row_offset, row_values) in values.iter().enumerate() {
                    let Some(target_data_row) = get_target_data_row(
                        &self.row_view, is_filtered, visible_start_idx, start_row, row_offset
                    ) else {
                        continue;
                    };

                    let mut grid_row: Vec<String> = Vec::new();
                    for (col_offset, value) in row_values.iter().enumerate() {
                        let col = start_col + col_offset;
                        if target_data_row < NUM_ROWS && col < NUM_COLS {
                            let old_value = self.sheet(cx).get_raw(target_data_row, col);
                            let new_value = Self::value_to_canonical_string(value);

                            grid_row.push(new_value.clone());

                            if old_value != new_value {
                                changes.push(CellChange {
                                    row: target_data_row, col, old_value, new_value: new_value.clone(),
                                });
                            }
                            self.set_cell_value(target_data_row, col, &new_value, cx);

                            end_data_row = end_data_row.max(target_data_row);
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
                let Some(target_data_row) = get_target_data_row(
                    &self.row_view, is_filtered, visible_start_idx, start_row, row_offset
                ) else {
                    continue;
                };

                let mut grid_row: Vec<String> = Vec::new();
                for (col_offset, cell_text) in line.split('\t').enumerate() {
                    let col = start_col + col_offset;
                    if target_data_row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet(cx).get_raw(target_data_row, col);
                        let parsed_value = Self::parse_external_value(cell_text);
                        let new_value = Self::value_to_canonical_string(&parsed_value);

                        grid_row.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row: target_data_row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.set_cell_value(target_data_row, col, &new_value, cx);

                        end_data_row = end_data_row.max(target_data_row);
                        end_col = end_col.max(col);
                    }
                }
                if !grid_row.is_empty() {
                    values_grid.push(grid_row);
                }
            }
        }
        self.end_batch_and_broadcast(cx);

        if !changes.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: values_grid,
                mode: PasteMode::Values,
            }.to_provenance(&self.sheet(cx).name);

            self.history.record_batch_with_provenance(self.sheet_index(cx), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc(cx);
        }

        // Validate pasted range and report failures (using data coordinates)
        let failures = self.wb(cx).validate_range(
            self.sheet_index(cx), data_start_row, start_col, end_data_row, end_col
        );
        let total_cells = (end_data_row - data_start_row + 1) * (end_col - start_col + 1);
        if failures.count > 0 {
            self.store_validation_failures(&failures);
            self.status_message = Some(format!(
                "Pasted values (Validation: {} of {} cells failed) — Press F8 to jump",
                failures.count, total_cells
            ));
        } else {
            self.status_message = Some("Pasted values".to_string());
        }
        self.clipboard_visual_range = None;
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

        // Try to parse as formatted number (commas, currency, parens)
        if let Some(n) = visigrid_engine::cell::try_parse_number(trimmed) {
            return Value::Number(n);
        }

        // Try to parse as plain number
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

    /// Paste Formulas: paste raw formulas with reference adjustment.
    /// - Internal clipboard: uses raw_tsv (formulas) with reference adjustment
    /// - External clipboard: falls back to normal paste() (no way to distinguish formula vs text)
    pub fn paste_formulas(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // If editing, paste into edit buffer
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        // Check if we have an internal clipboard with matching ID
        let clipboard_item = cx.read_from_clipboard();
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        let is_internal = self.internal_clipboard.as_ref().map_or(false, |ic| {
            let expected_id = format!("\"{}\"", ic.id);
            metadata.as_ref().map_or(false, |m| m == &expected_id)
        });

        if !is_internal {
            // External clipboard - fall back to normal paste
            // (No way to reliably distinguish "formula" vs "text starting with =" from external)
            self.paste(cx);
            return;
        }

        // Internal paste - use raw_tsv with reference adjustment
        let (start_row, start_col) = self.view_state.selected;
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        // Block if paste would split a merged region
        {
            let raw_tsv = self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.as_str()).unwrap_or("");
            let lines: Vec<&str> = raw_tsv.lines().collect();
            let paste_rows = lines.len();
            let paste_cols = lines.iter().map(|l| l.split('\t').count()).max().unwrap_or(1);
            if paste_rows > 0 && paste_cols > 0 {
                let dest_max_row = (data_start_row + paste_rows).saturating_sub(1);
                let dest_max_col = (start_col + paste_cols).saturating_sub(1);
                if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, dest_max_row, dest_max_col, cx) {
                    self.status_message = Some(format!(
                        "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                        Self::col_to_letter(mc), mr + 1,
                    ));
                    cx.notify();
                    return;
                }
            }
        }

        let mut changes = Vec::new();
        let mut values_grid: Vec<Vec<String>> = Vec::new();
        let mut end_data_row = data_start_row;
        let mut end_col = start_col;

        // Get source position and raw_tsv from internal clipboard
        let (src_row, src_col) = self.internal_clipboard.as_ref().map(|ic| ic.source).unwrap_or((0, 0));
        let raw_tsv = self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone()).unwrap_or_default();
        let src_data_row = self.row_view.view_to_data(src_row);
        let (delta_row, delta_col) = (data_start_row as i32 - src_data_row as i32, start_col as i32 - src_col as i32);

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        self.wb_mut(cx, |wb| wb.begin_batch());
        for (row_offset, line) in raw_tsv.lines().enumerate() {
            // Determine target view row for this clipboard row
            let target_data_row = if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                        self.row_view.view_to_data(view_row)
                    } else {
                        continue;
                    }
                } else {
                    continue;
                }
            } else {
                let view_row = start_row + row_offset;
                if view_row >= NUM_ROWS { continue; }
                view_row
            };

            let mut row_values: Vec<String> = Vec::new();
            for (col_offset, value) in line.split('\t').enumerate() {
                let col = start_col + col_offset;
                if target_data_row < NUM_ROWS && col < NUM_COLS {
                    let old_value = self.sheet(cx).get_raw(target_data_row, col);

                    // Adjust formula references using constant delta from source to destination
                    let new_value = if value.starts_with('=') {
                        self.adjust_formula_refs(value, delta_row, delta_col)
                    } else {
                        value.to_string()
                    };

                    row_values.push(new_value.clone());

                    if old_value != new_value {
                        changes.push(CellChange {
                            row: target_data_row, col, old_value, new_value: new_value.clone(),
                        });
                    }
                    self.set_cell_value(target_data_row, col, &new_value, cx);

                    end_data_row = end_data_row.max(target_data_row);
                    end_col = end_col.max(col);
                }
            }
            if !row_values.is_empty() {
                values_grid.push(row_values);
            }
        }
        self.end_batch_and_broadcast(cx);

        // Record with provenance (PasteMode::Formulas)
        if !changes.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: values_grid,
                mode: PasteMode::Formulas,
            }.to_provenance(&self.sheet(cx).name);

            self.history.record_batch_with_provenance(self.sheet_index(cx), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;
        }

        self.clipboard_visual_range = None;
        self.status_message = Some("Pasted formulas".to_string());
        self.maybe_smoke_recalc(cx);
        cx.notify();
    }

    /// Paste Formats: paste cell formatting only (no values).
    /// - Internal clipboard only: applies formats from copied range
    /// - External clipboard: no-op with status message (no format data available)
    pub fn paste_formats(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // Paste Formats doesn't make sense in edit mode
        if self.mode.is_editing() {
            self.status_message = Some("Exit edit mode to paste formats".to_string());
            cx.notify();
            return;
        }

        // Check if we have an internal clipboard with matching ID
        let clipboard_item = cx.read_from_clipboard();
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        let is_internal = self.internal_clipboard.as_ref().map_or(false, |ic| {
            let expected_id = format!("\"{}\"", ic.id);
            metadata.as_ref().map_or(false, |m| m == &expected_id)
        });

        if !is_internal {
            // External clipboard - no format data available
            self.status_message = Some("Paste Formats requires VisiGrid clipboard".to_string());
            cx.notify();
            return;
        }

        // Get formats from internal clipboard
        let formats = match &self.internal_clipboard {
            Some(ic) if !ic.formats.is_empty() => ic.formats.clone(),
            _ => {
                self.status_message = Some("No formats in clipboard".to_string());
                cx.notify();
                return;
            }
        };

        let (start_row, start_col) = self.view_state.selected;
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        let mut format_patches = Vec::new();

        for (row_offset, row_formats) in formats.iter().enumerate() {
            // Determine target data row
            let target_data_row = if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                        self.row_view.view_to_data(view_row)
                    } else {
                        continue;
                    }
                } else {
                    continue;
                }
            } else {
                let view_row = start_row + row_offset;
                if view_row >= NUM_ROWS { continue; }
                view_row
            };

            for (col_offset, format) in row_formats.iter().enumerate() {
                let col = start_col + col_offset;
                if target_data_row < NUM_ROWS && col < NUM_COLS {
                    // Get old format for history
                    let old_format = self.sheet(cx).get_format(target_data_row, col).clone();

                    // Apply format (full replace)
                    self.active_sheet_mut(cx, |s| {
                        s.set_format(target_data_row, col, format.clone());
                    });

                    // Track change for history
                    if old_format != *format {
                        format_patches.push(CellFormatPatch {
                            row: target_data_row,
                            col,
                            before: old_format,
                            after: format.clone(),
                        });
                    }
                }
            }
        }

        // Record format changes in history with provenance
        if !format_patches.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: vec![], // No value changes
                mode: PasteMode::Formats,
            }.to_provenance(&self.sheet(cx).name);

            self.history.record_format_with_provenance(
                self.sheet_index(cx),
                format_patches,
                FormatActionKind::PasteFormats,
                "Paste Formats".to_string(),
                Some(provenance),
            );
            self.is_modified = true;
        }

        self.clipboard_visual_range = None;
        let rows = formats.len();
        let cols = formats.first().map(|r| r.len()).unwrap_or(0);
        self.status_message = Some(format!("Pasted formats to {}x{} range", rows, cols));
        cx.notify();
    }

    pub fn delete_selection(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        let mut changes = Vec::new();
        let mut skipped_spill_receivers = false;
        let is_filtered = self.row_view.is_filtered();

        // Delete from all selection ranges (including discontiguous Ctrl+Click selections)
        // When filtered, only delete from visible rows
        self.wb_mut(cx, |wb| wb.begin_batch());
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for view_row in min_row..=max_row {
                // Skip hidden rows when filtered
                if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                    continue;
                }

                // Convert view row to data row for sheet access
                let data_row = self.row_view.view_to_data(view_row);

                for col in min_col..=max_col {
                    // Skip spill receivers - only the parent formula can be deleted
                    if self.sheet(cx).is_spill_receiver(data_row, col) {
                        skipped_spill_receivers = true;
                        continue;
                    }

                    let old_value = self.sheet(cx).get_raw(data_row, col);
                    if !old_value.is_empty() {
                        changes.push(CellChange {
                            row: data_row, col, old_value, new_value: String::new(),
                        });
                    }
                    self.clear_cell_value(data_row, col, cx);
                }
            }
        }
        self.end_batch_and_broadcast(cx);

        let had_changes = !changes.is_empty();
        if had_changes {
            // Only attach provenance for single contiguous selection
            let provenance = if self.view_state.additional_selections.is_empty() {
                let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
                // Use data rows for provenance
                let data_min_row = self.row_view.view_to_data(min_row);
                let data_max_row = self.row_view.view_to_data(max_row);
                Some(MutationOp::Clear {
                    sheet: self.sheet(cx).id,
                    start_row: data_min_row,
                    start_col: min_col,
                    end_row: data_max_row,
                    end_col: max_col,
                    mode: ClearMode::All,
                }.to_provenance(&self.sheet(cx).name))
            } else {
                None  // Discontiguous selection - no provenance
            };
            self.history.record_batch_with_provenance(self.sheet_index(cx), changes, provenance);
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;
        }

        if had_changes {
            self.clipboard_visual_range = None;
        }

        if skipped_spill_receivers && !had_changes {
            self.status_message = Some("Cannot delete spill range. Delete the parent formula instead.".to_string());
        }

        cx.notify();
    }

    /// Check if pasting into (min_row..=max_row, min_col..=max_col) would split any merge.
    /// Returns Some(merge_origin) for the first offending merge, or None if safe.
    fn paste_would_split_merge(
        &self, min_row: usize, min_col: usize,
        max_row: usize, max_col: usize, cx: &App,
    ) -> Option<(usize, usize)> {
        let sheet = self.sheet(cx);
        if sheet.merged_regions.is_empty() { return None; }

        for merge in &sheet.merged_regions {
            let intersects = merge.end.0 >= min_row && merge.start.0 <= max_row
                          && merge.end.1 >= min_col && merge.start.1 <= max_col;
            if !intersects { continue; }

            let contained = merge.start.0 >= min_row && merge.end.0 <= max_row
                          && merge.start.1 >= min_col && merge.end.1 <= max_col;
            if !contained {
                return Some(merge.start);
            }
        }
        None
    }
}
