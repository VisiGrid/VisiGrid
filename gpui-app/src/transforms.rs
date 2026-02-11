//! Cell transform operations for spreadsheet
//!
//! This module provides a generalized transform infrastructure. Each transform
//! is a pure function from old cell text → new cell text. The `apply_transform`
//! method handles selection iteration, undo batching, history, and status messages.
//!
//! ## Cell Policy
//!
//! Each transform declares which cells it operates on via [`CellPolicy`]:
//! - `ValuesOnly` — skip formula cells (text transforms)
//! - `FormulasOnly` — only operate on formula cells (future ref transforms)
//! - `Both` — operate on all non-empty cells
//!
//! ## Dry-run Preview (Pro)
//!
//! `apply_transform_preview()` computes the diff without mutating state, returning
//! a `TransformPreview` with before/after rows. The preview dialog can then call
//! `apply_transform_commit()` to apply pre-computed changes without recomputation.

use gpui::*;

use crate::app::Spreadsheet;
use crate::history::CellChange;

/// Maximum number of diff rows rendered in the preview dialog.
/// All changes are still applied — this only caps the visual display.
pub const MAX_PREVIEW_DISPLAY_ROWS: usize = 200;

// ============================================================================
// CellPolicy — which cells a transform operates on
// ============================================================================

/// Controls which cells a transform is allowed to modify.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellPolicy {
    /// Skip formula cells (`=...`), only transform plain values.
    ValuesOnly,
    /// Only transform formula cells, skip plain values.
    FormulasOnly,
    /// Transform all non-empty cells.
    Both,
}

// ============================================================================
// TransformOp — the set of available transforms
// ============================================================================

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TransformOp {
    TrimWhitespace,
    Uppercase,
    Lowercase,
    TitleCase,
    SentenceCase,
}

impl TransformOp {
    /// Human-readable name for status messages and history provenance.
    pub fn label(&self) -> &'static str {
        match self {
            Self::TrimWhitespace => "Trim Whitespace",
            Self::Uppercase => "UPPERCASE",
            Self::Lowercase => "lowercase",
            Self::TitleCase => "Title Case",
            Self::SentenceCase => "Sentence Case",
        }
    }

    /// Which cells this transform operates on.
    pub fn policy(&self) -> CellPolicy {
        match self {
            Self::TrimWhitespace
            | Self::Uppercase
            | Self::Lowercase
            | Self::TitleCase
            | Self::SentenceCase => CellPolicy::ValuesOnly,
        }
    }

    /// Apply the transform to a single cell value.
    /// Returns `None` if the value is unchanged (no-op).
    pub fn transform(&self, value: &str) -> Option<String> {
        let result = match self {
            Self::TrimWhitespace => value.trim().to_string(),
            Self::Uppercase => value.to_uppercase(),
            Self::Lowercase => value.to_lowercase(),
            Self::TitleCase => title_case(value),
            Self::SentenceCase => sentence_case(value),
        };

        if result == value {
            None
        } else {
            Some(result)
        }
    }
}

// ============================================================================
// TransformPreview — dry-run result for diff dialog
// ============================================================================

/// A single row in the transform diff preview.
#[derive(Clone, Debug)]
pub struct TransformDiffRow {
    pub row: usize,
    pub col: usize,
    pub before: String,
    pub after: String,
}

/// Result of a dry-run transform: pre-computed changes ready to apply.
#[derive(Clone, Debug)]
pub struct TransformPreview {
    /// The operation that generated this preview.
    pub op: TransformOp,
    /// Sheet index at time of preview (for commit).
    pub sheet_index: usize,
    /// All cell changes (before/after), in iteration order.
    pub diffs: Vec<TransformDiffRow>,
}

impl TransformPreview {
    /// Number of cells that will be modified.
    pub fn affected_count(&self) -> usize {
        self.diffs.len()
    }
}

// ============================================================================
// Pure transform functions
// ============================================================================

/// Capitalize first letter of each whitespace-delimited word, lowercase the rest.
pub(crate) fn title_case(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut prev_whitespace = true;

    for ch in s.chars() {
        if ch.is_whitespace() {
            result.push(ch);
            prev_whitespace = true;
        } else if prev_whitespace {
            for upper in ch.to_uppercase() {
                result.push(upper);
            }
            prev_whitespace = false;
        } else {
            for lower in ch.to_lowercase() {
                result.push(lower);
            }
            prev_whitespace = false;
        }
    }

    result
}

/// Capitalize first letter after sentence-ending punctuation (`.` `!` `?`) or at
/// string start. Lowercases all other alphabetic characters.
pub(crate) fn sentence_case(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut capitalize_next = true;

    for ch in s.chars() {
        if ch == '.' || ch == '!' || ch == '?' {
            result.push(ch);
            capitalize_next = true;
        } else if capitalize_next && ch.is_alphabetic() {
            for upper in ch.to_uppercase() {
                result.push(upper);
            }
            capitalize_next = false;
        } else {
            for lower in ch.to_lowercase() {
                result.push(lower);
            }
        }
    }

    result
}

// ============================================================================
// A1 helper (local, avoids coupling to engine internals)
// ============================================================================

fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 { break; }
        n = n / 26 - 1;
    }
    result
}

fn cell_ref(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

// ============================================================================
// apply_transform — immediate application (Free path)
// ============================================================================

impl Spreadsheet {
    /// Apply a transform to all cells in the current selection (immediate, no preview).
    ///
    /// Follows the existing batch mutation pattern:
    /// block guards → begin_batch → iterate → CellChange → end_batch → record_batch → status.
    pub fn apply_transform(
        &mut self,
        op: TransformOp,
        cx: &mut Context<Self>,
    ) {
        if self.block_if_merged("transform", cx) { return; }
        if self.block_if_previewing(cx) { return; }

        let policy = op.policy();
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        let mut changes = Vec::new();
        let mut transformed_count = 0;

        self.workbook.update(cx, |wb, _| wb.begin_batch());

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet(cx).get_raw(row, col);

                if !should_process(&old_value, policy) {
                    continue;
                }

                if let Some(new_value) = op.transform(&old_value) {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                    self.set_cell_value(row, col, &new_value, cx);
                    transformed_count += 1;
                }
            }
        }

        self.end_batch_and_broadcast(cx);

        if !changes.is_empty() {
            let scope = format!(
                "{}, {} cell{}",
                "Selection",
                transformed_count,
                if transformed_count == 1 { "" } else { "s" },
            );
            let provenance = visigrid_engine::provenance::Provenance {
                op: visigrid_engine::provenance::MutationOp::MultiEdit {
                    sheet: self.sheet(cx).id,
                    cells: changes.iter().map(|c| (c.row, c.col)).collect(),
                    value: format!("Transform: {}", op.label()),
                },
                label: format!("Transform: {}", op.label()),
                scope,
                lua: format!("-- Transform: {} ({})", op.label(), transformed_count),
            };
            self.history.record_batch_with_provenance(
                self.sheet_index(cx),
                changes,
                Some(provenance),
            );
            self.bump_cells_rev();
            self.is_modified = true;
        }

        let label = op.label();
        let msg = if transformed_count == 0 {
            format!("No cells to transform ({})", label)
        } else if transformed_count == 1 {
            format!("Transformed 1 cell ({})", label)
        } else {
            format!("Transformed {} cells ({})", transformed_count, label)
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    // ========================================================================
    // Dry-run preview (Pro path)
    // ========================================================================

    /// Compute a transform preview without mutating state.
    ///
    /// Returns `None` if no cells would change (or if blocked by guards).
    pub fn apply_transform_preview(
        &mut self,
        op: TransformOp,
        cx: &mut Context<Self>,
    ) -> Option<TransformPreview> {
        if self.block_if_merged("transform", cx) { return None; }
        if self.block_if_previewing(cx) { return None; }

        let policy = op.policy();
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        let mut diffs = Vec::new();

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet(cx).get_raw(row, col);

                if !should_process(&old_value, policy) {
                    continue;
                }

                if let Some(new_value) = op.transform(&old_value) {
                    diffs.push(TransformDiffRow {
                        row,
                        col,
                        before: old_value,
                        after: new_value,
                    });
                }
            }
        }

        if diffs.is_empty() {
            self.status_message = Some(format!("No cells to transform ({})", op.label()));
            cx.notify();
            return None;
        }

        Some(TransformPreview {
            op,
            sheet_index: self.sheet_index(cx),
            diffs,
        })
    }

    /// Apply a pre-computed transform preview (no recomputation).
    ///
    /// Uses the stored before/after values directly.
    pub fn apply_transform_commit(
        &mut self,
        preview: &TransformPreview,
        cx: &mut Context<Self>,
    ) {
        let mut changes = Vec::new();

        self.workbook.update(cx, |wb, _| wb.begin_batch());

        for diff in &preview.diffs {
            changes.push(CellChange {
                row: diff.row,
                col: diff.col,
                old_value: diff.before.clone(),
                new_value: diff.after.clone(),
            });
            self.set_cell_value(diff.row, diff.col, &diff.after, cx);
        }

        self.end_batch_and_broadcast(cx);

        let count = changes.len();
        if !changes.is_empty() {
            let scope = format!(
                "Selection, {} cell{}",
                count,
                if count == 1 { "" } else { "s" },
            );
            let provenance = visigrid_engine::provenance::Provenance {
                op: visigrid_engine::provenance::MutationOp::MultiEdit {
                    sheet: self.sheet(cx).id,
                    cells: changes.iter().map(|c| (c.row, c.col)).collect(),
                    value: format!("Transform: {}", preview.op.label()),
                },
                label: format!("Transform: {}", preview.op.label()),
                scope,
                lua: format!("-- Transform: {} ({})", preview.op.label(), count),
            };
            self.history.record_batch_with_provenance(
                preview.sheet_index,
                changes,
                Some(provenance),
            );
            self.bump_cells_rev();
            self.is_modified = true;
        }

        let label = preview.op.label();
        let msg = if count == 1 {
            format!("Transformed 1 cell ({})", label)
        } else {
            format!("Transformed {} cells ({})", count, label)
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    // ========================================================================
    // Pro-gated entry point: preview if Pro, immediate if Free
    // ========================================================================

    /// Apply a transform with Pro gating: if Pro is active, show diff preview;
    /// if Free, apply immediately.
    pub fn apply_transform_pro(
        &mut self,
        op: TransformOp,
        cx: &mut Context<Self>,
    ) {
        let is_pro = visigrid_license::is_feature_enabled("transforms");

        if is_pro {
            // Dry-run: compute preview, store it, switch to preview mode
            if let Some(preview) = self.apply_transform_preview(op, cx) {
                self.transform_preview = Some(preview);
                self.mode = crate::mode::Mode::TransformPreview;
                cx.notify();
            }
            // If None, status_message was already set by apply_transform_preview
        } else {
            // Free: apply immediately
            self.apply_transform(op, cx);
        }
    }

    /// Cancel the transform preview dialog without applying.
    pub fn cancel_transform_preview(&mut self, cx: &mut Context<Self>) {
        self.transform_preview = None;
        self.mode = crate::mode::Mode::Navigation;
        self.status_message = Some("Transform cancelled".to_string());
        cx.notify();
    }

    /// Confirm and apply the stored transform preview.
    pub fn confirm_transform_preview(&mut self, cx: &mut Context<Self>) {
        let preview = match self.transform_preview.take() {
            Some(p) => p,
            None => return,
        };
        self.mode = crate::mode::Mode::Navigation;
        self.apply_transform_commit(&preview, cx);
    }
}

// ============================================================================
// Cell policy filter (shared between apply and preview)
// ============================================================================

/// Returns true if the cell should be processed based on policy.
fn should_process(value: &str, policy: CellPolicy) -> bool {
    if value.is_empty() {
        return false;
    }
    let is_formula = value.starts_with('=');
    match policy {
        CellPolicy::ValuesOnly if is_formula => false,
        CellPolicy::FormulasOnly if !is_formula => false,
        _ => true,
    }
}
