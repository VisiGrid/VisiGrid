//! Recalculation types and reporting.
//!
//! This module defines the types used for ordered formula recomputation
//! and cycle detection.

use crate::cell_id::CellId;
use rustc_hash::FxHashMap;
use std::time::SystemTime;

/// Per-cell recalculation metadata.
///
/// Tracks when and how a cell was last recomputed.
/// Used by the Inspector to show explainability info.
#[derive(Debug, Clone)]
pub struct CellRecalcInfo {
    /// Dependency depth of this cell.
    /// A formula with no formula dependencies has depth 1.
    /// A formula depending on another formula has depth = max(precedent depths) + 1.
    /// Value cells (non-formulas) have depth 0.
    pub depth: usize,

    /// Position in the evaluation order (0-indexed).
    /// Lower numbers are evaluated before higher numbers.
    pub eval_order: usize,

    /// When this cell was last recomputed.
    pub recompute_time: SystemTime,

    /// Whether this cell has dynamic/unknown dependencies (INDIRECT, OFFSET).
    pub has_unknown_deps: bool,
}

impl CellRecalcInfo {
    /// Create a new cell recalc info.
    pub fn new(depth: usize, eval_order: usize, has_unknown_deps: bool) -> Self {
        Self {
            depth,
            eval_order,
            recompute_time: SystemTime::now(),
            has_unknown_deps,
        }
    }
}

/// Report from a full ordered recompute operation.
///
/// This is the backbone for Phase 1.5 logging and Phase 2 status bar.
#[derive(Debug, Clone, Default)]
pub struct RecalcReport {
    /// Time taken for full recompute in milliseconds.
    pub duration_ms: u64,

    /// Number of formula cells that were recomputed.
    pub cells_recomputed: usize,

    /// Maximum dependency depth encountered.
    /// A formula with no dependencies has depth 1.
    /// A formula depending on another formula has depth = max(precedent depths) + 1.
    pub max_depth: usize,

    /// True if cycles were detected during recompute.
    /// Cycle cells are marked with #CYCLE! error.
    pub had_cycles: bool,

    /// Number of cells with unknown dependencies (INDIRECT/OFFSET)
    /// that were conservatively recomputed.
    pub unknown_deps_recomputed: usize,

    /// Errors encountered during recompute (truncated to first 100).
    pub errors: Vec<RecalcError>,

    /// Per-cell recalc metadata for Inspector explainability.
    /// Maps CellId -> CellRecalcInfo for all cells that were recomputed.
    pub cell_info: FxHashMap<CellId, CellRecalcInfo>,

    /// Number of SCCs that were iteratively resolved (0 if iteration disabled).
    pub scc_count: usize,

    /// Maximum iterations performed across all SCCs (0 if no iteration).
    pub iterations_performed: u32,

    /// True if all SCCs converged within tolerance.
    pub converged: bool,

    /// Number of cells participating in circular reference cycles.
    /// This is the graph truth — nonzero whenever cycles exist, regardless of
    /// whether iteration has resolved them. Drive "resolved" messaging from
    /// `converged && iterative_enabled` in the UI, not from this being zero.
    pub cycle_cells: usize,

    /// Phase timing: microseconds spent clearing caches (invalidation).
    pub phase_invalidation_us: u64,
    /// Phase timing: microseconds spent in topological sort.
    pub phase_topo_sort_us: u64,
    /// Phase timing: microseconds spent evaluating formulas.
    pub phase_eval_us: u64,
    /// Phase timing: aggregate microseconds in Lua custom function calls.
    /// Subset of phase_eval_us. Set by the caller, not the engine —
    /// GUI-only for v1 (CLI/headless won't report Lua timing yet).
    pub phase_lua_total_us: u64,
}

impl RecalcReport {
    /// Create a new empty report.
    pub fn new() -> Self {
        Self::default()
    }

    /// Format as a concise one-line summary for logging.
    pub fn summary(&self) -> String {
        if self.scc_count > 0 {
            format!(
                "{} cells in {}ms, depth={}, cycles={}, unknown={}, sccs={}, iters={}, converged={}",
                self.cells_recomputed,
                self.duration_ms,
                self.max_depth,
                self.had_cycles,
                self.unknown_deps_recomputed,
                self.scc_count,
                self.iterations_performed,
                self.converged,
            )
        } else {
            format!(
                "{} cells in {}ms, depth={}, cycles={}, unknown={}",
                self.cells_recomputed,
                self.duration_ms,
                self.max_depth,
                self.had_cycles,
                self.unknown_deps_recomputed
            )
        }
    }

    /// Format as a one-line log entry for smoke mode.
    ///
    /// Format: `[recalc/full] 14ms  628 cells  depth=7  unknown=3  cycles=0  errors=0`
    pub fn log_line(&self) -> String {
        if self.scc_count > 0 {
            format!(
                "[recalc/full] {:>4}ms  {} cells  depth={}  unknown={}  cycles={}  errors={}  sccs={}  iters={}  converged={}",
                self.duration_ms,
                self.cells_recomputed,
                self.max_depth,
                self.unknown_deps_recomputed,
                if self.had_cycles { 1 } else { 0 },
                self.errors.len(),
                self.scc_count,
                self.iterations_performed,
                self.converged,
            )
        } else {
            format!(
                "[recalc/full] {:>4}ms  {} cells  depth={}  unknown={}  cycles={}  errors={}",
                self.duration_ms,
                self.cells_recomputed,
                self.max_depth,
                self.unknown_deps_recomputed,
                if self.had_cycles { 1 } else { 0 },
                self.errors.len()
            )
        }
    }

    /// Get recalc info for a specific cell.
    pub fn get_cell_info(&self, cell: &CellId) -> Option<&CellRecalcInfo> {
        self.cell_info.get(cell)
    }

    /// Get cells that were evaluated immediately before/after a given cell.
    /// Returns (prev_cell, next_cell) where either may be None if at boundary.
    pub fn get_adjacent_cells(&self, cell: &CellId) -> (Option<CellId>, Option<CellId>) {
        let Some(info) = self.cell_info.get(cell) else {
            return (None, None);
        };
        let order = info.eval_order;

        let mut prev: Option<CellId> = None;
        let mut next: Option<CellId> = None;

        for (c, i) in &self.cell_info {
            if i.eval_order + 1 == order {
                prev = Some(*c);
            } else if i.eval_order == order + 1 {
                next = Some(*c);
            }
        }

        (prev, next)
    }

    /// Get all cells sorted by evaluation order.
    pub fn cells_by_eval_order(&self) -> Vec<(CellId, &CellRecalcInfo)> {
        let mut cells: Vec<_> = self.cell_info.iter().map(|(c, i)| (*c, i)).collect();
        cells.sort_by_key(|(_, i)| i.eval_order);
        cells
    }
}

/// Heuristic hotspot suspect from the dependency graph.
///
/// Used by the Performance Profiler to identify cells likely contributing
/// to slow recalculations.
#[derive(Debug, Clone)]
pub struct HotspotEntry {
    pub cell: CellId,
    pub fan_in: usize,
    pub fan_out: usize,
    pub depth: usize,
    pub has_unknown_deps: bool,
    pub score: f64,
}

impl RecalcReport {
    /// Compute heuristic hotspot suspects from cells that were actually recomputed.
    ///
    /// Only analyzes cells present in `self.cell_info` (the recomputed set),
    /// NOT the entire graph. This keeps cost proportional to the recalc, not
    /// the workbook size.
    ///
    /// Scoring: fan_out * 3.0 + fan_in * 1.5 + depth * 1.0 + (unknown ? 10.0 : 0.0)
    /// Returns top N sorted by score descending.
    pub fn hotspot_analysis(&self, dep_graph: &crate::dep_graph::DepGraph, top_n: usize) -> Vec<HotspotEntry> {
        let mut entries: Vec<HotspotEntry> = self.cell_info.iter().map(|(cell_id, info)| {
            let fan_in = dep_graph.precedent_count(*cell_id);
            let fan_out = dep_graph.dependent_count(*cell_id);
            let depth = info.depth;
            let has_unknown_deps = info.has_unknown_deps;
            let score = fan_out as f64 * 3.0
                + fan_in as f64 * 1.5
                + depth as f64 * 1.0
                + if has_unknown_deps { 10.0 } else { 0.0 };
            HotspotEntry { cell: *cell_id, fan_in, fan_out, depth, has_unknown_deps, score }
        }).collect();
        entries.sort_by(|a, b| b.score.partial_cmp(&a.score).unwrap_or(std::cmp::Ordering::Equal));
        entries.truncate(top_n);
        entries
    }
}

/// An error that occurred during recomputation of a specific cell.
#[derive(Debug, Clone)]
pub struct RecalcError {
    /// The cell where the error occurred.
    pub cell: CellId,

    /// Description of the error.
    pub error: String,
}

impl RecalcError {
    /// Create a new recalc error.
    pub fn new(cell: CellId, error: impl Into<String>) -> Self {
        Self {
            cell,
            error: error.into(),
        }
    }
}

/// Report when cycle detection finds a circular reference.
#[derive(Debug, Clone)]
pub struct CycleReport {
    /// Cells participating in the cycle.
    /// May be a subset for large cycles.
    pub cells: Vec<CellId>,

    /// Human-readable description of the cycle.
    pub message: String,
}

impl CycleReport {
    /// Create a new cycle report.
    pub fn new(cells: Vec<CellId>, message: impl Into<String>) -> Self {
        Self {
            cells,
            message: message.into(),
        }
    }

    /// Create a cycle report for a self-referencing cell.
    pub fn self_reference(cell: CellId) -> Self {
        Self {
            cells: vec![cell],
            message: format!("Cell {} references itself", cell),
        }
    }

    /// Create a cycle report for a multi-cell cycle.
    pub fn cycle(cells: Vec<CellId>) -> Self {
        let cell_list: Vec<String> = cells.iter().map(|c| c.to_string()).collect();
        let message = if cells.len() <= 5 {
            format!("Circular reference: {}", cell_list.join(" → "))
        } else {
            format!(
                "Circular reference involving {} cells: {} → ... → {}",
                cells.len(),
                cell_list[0],
                cell_list.last().unwrap()
            )
        };
        Self { cells, message }
    }
}

impl std::fmt::Display for CycleReport {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for CycleReport {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::SheetId;

    fn cell(sheet: u64, row: usize, col: usize) -> CellId {
        CellId::new(SheetId::from_raw(sheet), row, col)
    }

    #[test]
    fn test_recalc_report_default() {
        let report = RecalcReport::default();
        assert_eq!(report.duration_ms, 0);
        assert_eq!(report.cells_recomputed, 0);
        assert_eq!(report.max_depth, 0);
        assert!(!report.had_cycles);
        assert_eq!(report.unknown_deps_recomputed, 0);
        assert!(report.errors.is_empty());
        assert_eq!(report.scc_count, 0);
        assert_eq!(report.iterations_performed, 0);
        assert!(!report.converged);
    }

    #[test]
    fn test_recalc_report_summary() {
        let report = RecalcReport {
            duration_ms: 42,
            cells_recomputed: 100,
            max_depth: 5,
            had_cycles: false,
            unknown_deps_recomputed: 2,
            errors: vec![],
            cell_info: Default::default(),
            ..Default::default()
        };
        assert_eq!(
            report.summary(),
            "100 cells in 42ms, depth=5, cycles=false, unknown=2"
        );
    }

    #[test]
    fn test_recalc_report_log_line() {
        let report = RecalcReport {
            duration_ms: 14,
            cells_recomputed: 628,
            max_depth: 7,
            had_cycles: false,
            unknown_deps_recomputed: 3,
            errors: vec![],
            cell_info: Default::default(),
            ..Default::default()
        };
        assert_eq!(
            report.log_line(),
            "[recalc/full]   14ms  628 cells  depth=7  unknown=3  cycles=0  errors=0"
        );
    }

    #[test]
    fn test_recalc_report_log_line_with_cycles() {
        let report = RecalcReport {
            duration_ms: 5,
            cells_recomputed: 10,
            max_depth: 2,
            had_cycles: true,
            unknown_deps_recomputed: 0,
            errors: vec![RecalcError::new(cell(1, 0, 0), "test error")],
            cell_info: Default::default(),
            ..Default::default()
        };
        assert_eq!(
            report.log_line(),
            "[recalc/full]    5ms  10 cells  depth=2  unknown=0  cycles=1  errors=1"
        );
    }

    #[test]
    fn test_cycle_report_self_reference() {
        let a1 = cell(1, 0, 0);
        let report = CycleReport::self_reference(a1);
        assert_eq!(report.cells.len(), 1);
        assert!(report.message.contains("references itself"));
    }

    #[test]
    fn test_cycle_report_small_cycle() {
        let cells = vec![cell(1, 0, 0), cell(1, 0, 1), cell(1, 0, 2)];
        let report = CycleReport::cycle(cells);
        assert!(report.message.contains("→"));
        assert!(!report.message.contains("..."));
    }

    #[test]
    fn test_cycle_report_large_cycle() {
        let cells: Vec<CellId> = (0..10).map(|i| cell(1, i, 0)).collect();
        let report = CycleReport::cycle(cells);
        assert!(report.message.contains("..."));
        assert!(report.message.contains("10 cells"));
    }

    #[test]
    fn test_cycle_report_display() {
        let report = CycleReport::new(vec![cell(1, 0, 0)], "Test error");
        assert_eq!(format!("{}", report), "Test error");
    }
}
