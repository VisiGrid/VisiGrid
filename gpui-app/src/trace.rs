//! Dependency tracing for formula auditing.
//!
//! ## Overview
//!
//! When Trace Mode is enabled, selecting a cell highlights its dependency graph:
//! - **Source cell**: The selected cell (strong highlight)
//! - **Precedents**: Cells that feed into this cell (inputs)
//! - **Dependents**: Cells that depend on this cell (outputs)
//!
//! This works across split panes, making it visually clear how data flows
//! through the spreadsheet.
//!
//! ## Performance
//!
//! - TraceCache is recomputed only when selection or workbook structure changes
//! - Rendering uses O(1) HashSet membership checks
//! - Large traces (>10,000 cells) are capped with a status message
//!
//! ## Architecture
//!
//! - `TraceCache` holds the computed trace (source + precedents + dependents)
//! - `recompute_trace_if_needed()` is called on selection change
//! - `render_cell` checks membership for highlighting
//! - Both panes see the same trace (it's derived from shared workbook)

use gpui::*;
use std::collections::HashSet;
use visigrid_engine::cell_id::CellId;
use visigrid_engine::sheet::SheetId;

use crate::app::Spreadsheet;

/// Maximum number of cells to include in trace before capping.
/// Prevents pathological cases from killing performance.
pub const MAX_TRACE_CELLS: usize = 10_000;

/// Cached dependency trace for the currently selected cell.
#[derive(Clone, Debug)]
pub struct TraceCache {
    /// The source cell being traced (active selection)
    pub source: CellId,

    /// Cells that the source depends on (inputs)
    pub precedents: HashSet<CellId>,

    /// Cells that depend on the source (outputs)
    pub dependents: HashSet<CellId>,

    /// True if the trace was capped due to size limits
    pub is_capped: bool,

    /// Current index when cycling through precedents (-1 = not cycling)
    pub precedent_index: isize,

    /// Current index when cycling through dependents (-1 = not cycling)
    pub dependent_index: isize,

    /// Count of precedent cells that have validation failures
    pub invalid_precedent_count: usize,
}

impl TraceCache {
    /// Create a new trace cache for a source cell.
    pub fn new(source: CellId) -> Self {
        Self {
            source,
            precedents: HashSet::new(),
            dependents: HashSet::new(),
            is_capped: false,
            precedent_index: -1,
            dependent_index: -1,
            invalid_precedent_count: 0,
        }
    }

    /// Get precedents as a sorted Vec for stable cycling order.
    /// Sorted by (sheet, row, col) for predictable navigation.
    pub fn precedents_sorted(&self) -> Vec<CellId> {
        let mut v: Vec<_> = self.precedents.iter().cloned().collect();
        v.sort_by_key(|c| (c.sheet.raw(), c.row, c.col));
        v
    }

    /// Get dependents as a sorted Vec for stable cycling order.
    pub fn dependents_sorted(&self) -> Vec<CellId> {
        let mut v: Vec<_> = self.dependents.iter().cloned().collect();
        v.sort_by_key(|c| (c.sheet.raw(), c.row, c.col));
        v
    }

    /// Check if a cell is the trace source.
    pub fn is_source(&self, sheet: SheetId, row: usize, col: usize) -> bool {
        self.source.sheet == sheet && self.source.row == row && self.source.col == col
    }

    /// Check if a cell is a precedent (input to source).
    pub fn is_precedent(&self, sheet: SheetId, row: usize, col: usize) -> bool {
        self.precedents.contains(&CellId::new(sheet, row, col))
    }

    /// Check if a cell is a dependent (output from source).
    pub fn is_dependent(&self, sheet: SheetId, row: usize, col: usize) -> bool {
        self.dependents.contains(&CellId::new(sheet, row, col))
    }

    /// Total number of highlighted cells (source + precedents + dependents).
    pub fn total_cells(&self) -> usize {
        1 + self.precedents.len() + self.dependents.len()
    }
}

/// Result of trace classification for a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TraceRole {
    /// Not part of the current trace
    None,
    /// The source cell being traced
    Source,
    /// A precedent (input to source)
    Precedent,
    /// A dependent (output from source)
    Dependent,
}

impl TraceRole {
    /// Check if this cell has any trace role.
    pub fn is_traced(&self) -> bool {
        !matches!(self, TraceRole::None)
    }
}

// =============================================================================
// Spreadsheet Integration
// =============================================================================

impl Spreadsheet {
    /// Toggle trace mode on/off.
    pub fn toggle_trace(&mut self, cx: &mut Context<Self>) {
        self.trace_enabled = !self.trace_enabled;

        if self.trace_enabled {
            self.recompute_trace(cx);
            self.status_message = Some("Trace mode enabled".to_string());
        } else {
            self.trace_cache = None;
            self.status_message = Some("Trace mode disabled".to_string());
        }

        cx.notify();
    }

    /// Recompute trace cache for current selection.
    /// Called when selection changes (if trace is enabled).
    pub fn recompute_trace(&mut self, cx: &App) {
        if !self.trace_enabled {
            self.trace_cache = None;
            return;
        }

        // Get current selection from active pane (single cell only for now)
        let view_state = self.active_view_state();
        if view_state.selection_end.is_some() {
            // Range selection - disable trace
            self.trace_cache = None;
            return;
        }

        let (row, col) = view_state.selected;
        let sheet_id = self.sheet(cx).id;
        let source = CellId::new(sheet_id, row, col);

        let mut cache = TraceCache::new(source);

        // Collect precedents and dependents from workbook
        let wb = self.wb(cx);
        let mut total = 0;

        // Direct precedents
        for pred in wb.get_precedents(source.sheet, source.row, source.col) {
            if total >= MAX_TRACE_CELLS {
                cache.is_capped = true;
                break;
            }
            cache.precedents.insert(pred);
            total += 1;
        }

        // Direct dependents
        for dep in wb.get_dependents(source.sheet, source.row, source.col) {
            if total >= MAX_TRACE_CELLS {
                cache.is_capped = true;
                break;
            }
            cache.dependents.insert(dep);
            total += 1;
        }

        // Count invalid precedents (inputs with validation failures)
        // Note: validation_failures only tracks current sheet, so we check same-sheet precedents
        let current_sheet_id = self.sheet(cx).id;
        cache.invalid_precedent_count = cache
            .precedents
            .iter()
            .filter(|p| {
                p.sheet == current_sheet_id
                    && self.validation_failures.contains(&(p.row, p.col))
            })
            .count();

        self.trace_cache = Some(cache);
    }

    /// Recompute trace if enabled and selection changed.
    /// Call this after selection changes.
    pub fn recompute_trace_if_needed(&mut self, cx: &App) {
        if !self.trace_enabled {
            return;
        }

        // Check if selection matches cached source
        let view_state = self.active_view_state();
        if view_state.selection_end.is_some() {
            // Range selection - clear trace
            self.trace_cache = None;
            return;
        }

        let (row, col) = view_state.selected;
        let sheet_id = self.sheet(cx).id;

        // Check if cache is still valid
        if let Some(ref cache) = self.trace_cache {
            if cache.source.sheet == sheet_id
                && cache.source.row == row
                && cache.source.col == col
            {
                // Cache is still valid
                return;
            }
        }

        // Recompute
        self.recompute_trace(cx);
    }

    /// Invalidate trace cache due to workbook edit.
    /// Call this when cells are edited (dependencies may have changed).
    pub fn invalidate_trace_if_needed(&mut self, cx: &App) {
        if !self.trace_enabled {
            return;
        }
        // Clear cache - will be recomputed on next selection or access
        self.trace_cache = None;
        // Recompute immediately to keep UI in sync
        self.recompute_trace(cx);
    }

    /// Get the trace role for a cell (for rendering).
    pub fn trace_role(&self, sheet_id: SheetId, row: usize, col: usize) -> TraceRole {
        let Some(ref cache) = self.trace_cache else {
            return TraceRole::None;
        };

        if cache.is_source(sheet_id, row, col) {
            TraceRole::Source
        } else if cache.is_precedent(sheet_id, row, col) {
            TraceRole::Precedent
        } else if cache.is_dependent(sheet_id, row, col) {
            TraceRole::Dependent
        } else {
            TraceRole::None
        }
    }

    /// Get trace summary for status bar.
    pub fn trace_summary(&self) -> Option<String> {
        let cache = self.trace_cache.as_ref()?;

        let source_name = format!(
            "{}{}",
            Spreadsheet::col_letter(cache.source.col),
            cache.source.row + 1
        );

        // Platform-specific shortcut hints
        #[cfg(target_os = "macos")]
        let shortcuts = "⌥[ ⌥] | Back: ⌥↩";
        #[cfg(not(target_os = "macos"))]
        let shortcuts = "Ctrl+[ ] | Back: F5";

        let mut summary = format!(
            "Trace: {} | {} prec | {} dep | {}",
            source_name,
            cache.precedents.len(),
            cache.dependents.len(),
            shortcuts
        );

        if cache.is_capped {
            summary.push_str(" (capped)");
        }

        // Warn if any inputs have marked validation failures
        // Note: "marked" because validation only flags cells after explicit actions
        // (Circle Invalid Data, paste, fill) - not omniscient auto-revalidation
        if cache.invalid_precedent_count > 0 {
            summary.push_str(&format!(" | ⚠ {} marked invalid (F8)", cache.invalid_precedent_count));
        }

        Some(summary)
    }

    /// Cycle to next precedent cell (Alt+[).
    /// If reverse is true, cycles backwards (Alt+Shift+[).
    /// Auto-enables trace mode if not already enabled.
    pub fn cycle_trace_precedent(&mut self, reverse: bool, cx: &mut Context<Self>) {
        // Auto-enable trace if not already
        if !self.trace_enabled {
            self.trace_enabled = true;
            self.recompute_trace(cx);
        }

        // Extract data from cache to avoid borrow conflicts
        let (target, index, total) = {
            let Some(cache) = self.trace_cache.as_mut() else {
                self.status_message = Some("No trace available".to_string());
                cx.notify();
                return;
            };

            let precedents = cache.precedents_sorted();
            if precedents.is_empty() {
                self.status_message = Some("No precedents".to_string());
                cx.notify();
                return;
            }

            // Update index
            let len = precedents.len() as isize;
            if reverse {
                cache.precedent_index = if cache.precedent_index <= 0 {
                    len - 1
                } else {
                    cache.precedent_index - 1
                };
            } else {
                cache.precedent_index = (cache.precedent_index + 1) % len;
            }

            let target = precedents[cache.precedent_index as usize];
            let index = cache.precedent_index as usize + 1;
            let total = precedents.len();
            (target, index, total)
        };

        // Navigate to the target cell
        self.jump_to_trace_cell(target, "precedent", index, total, cx);
    }

    /// Cycle to next dependent cell (Alt+]).
    /// If reverse is true, cycles backwards (Alt+Shift+]).
    /// Auto-enables trace mode if not already enabled.
    pub fn cycle_trace_dependent(&mut self, reverse: bool, cx: &mut Context<Self>) {
        // Auto-enable trace if not already
        if !self.trace_enabled {
            self.trace_enabled = true;
            self.recompute_trace(cx);
        }

        // Extract data from cache to avoid borrow conflicts
        let (target, index, total) = {
            let Some(cache) = self.trace_cache.as_mut() else {
                self.status_message = Some("No trace available".to_string());
                cx.notify();
                return;
            };

            let dependents = cache.dependents_sorted();
            if dependents.is_empty() {
                self.status_message = Some("No dependents".to_string());
                cx.notify();
                return;
            }

            // Update index
            let len = dependents.len() as isize;
            if reverse {
                cache.dependent_index = if cache.dependent_index <= 0 {
                    len - 1
                } else {
                    cache.dependent_index - 1
                };
            } else {
                cache.dependent_index = (cache.dependent_index + 1) % len;
            }

            let target = dependents[cache.dependent_index as usize];
            let index = cache.dependent_index as usize + 1;
            let total = dependents.len();
            (target, index, total)
        };

        // Navigate to the target cell
        self.jump_to_trace_cell(target, "dependent", index, total, cx);
    }

    /// Jump to a traced cell, handling cross-sheet navigation.
    fn jump_to_trace_cell(
        &mut self,
        target: CellId,
        role: &str,
        index: usize,
        total: usize,
        cx: &mut Context<Self>,
    ) {
        let current_sheet = self.sheet(cx).id;

        // Switch sheet if needed
        if target.sheet != current_sheet {
            // Find sheet index
            let wb = self.wb(cx);
            if let Some(sheet_idx) = wb.sheets().iter().position(|s| s.id == target.sheet) {
                self.active_view_state_mut().active_sheet = sheet_idx;
            }
        }

        // Navigate to cell (without extending selection)
        let view_state = self.active_view_state_mut();
        view_state.selected = (target.row, target.col);
        view_state.selection_end = None;
        view_state.additional_selections.clear();

        // Ensure cell is visible
        self.ensure_cell_visible(target.row, target.col);

        // Status message
        let cell_name = format!(
            "{}{}",
            Spreadsheet::col_letter(target.col),
            target.row + 1
        );
        self.status_message = Some(format!("{} {}/{}: {}", role, index, total, cell_name));

        cx.notify();
    }

    /// Return to trace source cell (F5 on Windows, Alt+Enter on macOS).
    /// Only works when trace mode is enabled.
    pub fn return_to_trace_source(&mut self, cx: &mut Context<Self>) {
        if !self.trace_enabled {
            self.status_message = Some("Trace mode not enabled (Alt+T)".to_string());
            cx.notify();
            return;
        }

        let Some(cache) = self.trace_cache.as_ref() else {
            self.status_message = Some("No trace source".to_string());
            cx.notify();
            return;
        };

        let source = cache.source;
        let current_sheet = self.sheet(cx).id;

        // Switch sheet if needed
        if source.sheet != current_sheet {
            let wb = self.wb(cx);
            if let Some(sheet_idx) = wb.sheets().iter().position(|s| s.id == source.sheet) {
                self.active_view_state_mut().active_sheet = sheet_idx;
            }
        }

        // Navigate to source cell
        let view_state = self.active_view_state_mut();
        view_state.selected = (source.row, source.col);
        view_state.selection_end = None;
        view_state.additional_selections.clear();

        // Ensure cell is visible
        self.ensure_cell_visible(source.row, source.col);

        // Reset cycle indices
        if let Some(cache) = self.trace_cache.as_mut() {
            cache.precedent_index = -1;
            cache.dependent_index = -1;
        }

        let cell_name = format!(
            "{}{}",
            Spreadsheet::col_letter(source.col),
            source.row + 1
        );
        self.status_message = Some(format!("Back to source: {}", cell_name));

        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    use super::{TraceCache, TraceRole, MAX_TRACE_CELLS};
    use visigrid_engine::cell_id::CellId;
    use visigrid_engine::sheet::SheetId;

    fn cell(sheet: u64, row: usize, col: usize) -> CellId {
        CellId::new(SheetId::from_raw(sheet), row, col)
    }

    #[test]
    fn test_trace_cache_membership() {
        let source = cell(1, 5, 5);
        let mut cache = TraceCache::new(source);

        // Add some precedents and dependents
        cache.precedents.insert(cell(1, 0, 0));
        cache.precedents.insert(cell(1, 0, 1));
        cache.dependents.insert(cell(1, 10, 0));
        cache.dependents.insert(cell(2, 0, 0)); // Cross-sheet

        // Source checks
        assert!(cache.is_source(SheetId::from_raw(1), 5, 5));
        assert!(!cache.is_source(SheetId::from_raw(1), 5, 6));
        assert!(!cache.is_source(SheetId::from_raw(2), 5, 5));

        // Precedent checks
        assert!(cache.is_precedent(SheetId::from_raw(1), 0, 0));
        assert!(cache.is_precedent(SheetId::from_raw(1), 0, 1));
        assert!(!cache.is_precedent(SheetId::from_raw(1), 0, 2));

        // Dependent checks
        assert!(cache.is_dependent(SheetId::from_raw(1), 10, 0));
        assert!(cache.is_dependent(SheetId::from_raw(2), 0, 0));
        assert!(!cache.is_dependent(SheetId::from_raw(1), 0, 0));

        // Total count
        assert_eq!(cache.total_cells(), 5); // 1 source + 2 prec + 2 dep
    }

    #[test]
    fn test_trace_role() {
        assert!(!TraceRole::None.is_traced());
        assert!(TraceRole::Source.is_traced());
        assert!(TraceRole::Precedent.is_traced());
        assert!(TraceRole::Dependent.is_traced());
    }
}
