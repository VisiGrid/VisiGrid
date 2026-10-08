//! Runtime references join the ordinary range index, so edits invalidate their
//! readers without scanning every formula. A target change repeats calculation
//! in the updated dependency order before publishing the result.
use super::{Recalculated, Workbook};
use crate::{
    cell_id::CellId,
    formula::{
        eval::{EvalArg, EvalResult},
        parser::bind_expr,
    },
    recalc::{RecalcError, RecalcReport},
};
use rustc_hash::FxHashSet;
use std::sync::Arc;

const MAX_REFERENCE_PASSES: usize = 16;

impl Workbook {
    fn apply_dynamic_references(&mut self) -> Vec<CellId> {
        let pending = self.pending_dynamic_refs.take();
        let mut changed = Vec::new();
        for (cell, targets) in pending {
            let Some(ast) = self
                .sheet_by_id(cell.sheet)
                .and_then(|s| s.get_cell_opt(cell.row, cell.col))
                .and_then(|c| c.value().formula_ast())
            else {
                continue;
            };
            let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));
            let (refs, mut ranges) =
                self.formula_dependencies(&bound, cell.sheet, cell.row, cell.col);
            ranges.extend(targets);
            let key = |r: &crate::dep_graph::RangeRef| {
                (
                    r.sheet.raw(),
                    r.start_row,
                    r.start_col,
                    r.end_row,
                    r.end_col,
                )
            };
            ranges.sort_unstable_by_key(key);
            ranges.dedup();
            let mut old_ranges = self.dep_graph.precedent_ranges(cell).to_vec();
            old_ranges.sort_unstable_by_key(key);
            old_ranges.dedup();
            let old_refs: FxHashSet<_> = self.dep_graph.precedents(cell).collect();
            if old_refs == refs && old_ranges == ranges {
                continue;
            }
            let graph = Arc::make_mut(&mut self.dep_graph);
            graph.replace_edges(cell, refs);
            graph.set_ranges(cell, ranges);
            graph.register_leaf_formula(cell);
            changed.push(cell);
        }
        changed.sort_unstable_by_key(|c| (c.sheet.raw(), c.row, c.col));
        changed
    }

    pub(super) fn recompute_full_ordered_inner(
        &mut self,
        handler: Option<&dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>>,
    ) -> RecalcReport {
        let _recalc_scope = crate::custom_fns::recalc_scope();
        let _clock = crate::timing::ClockGuard::install(self.recalc_clock);
        let start = crate::timing::Instant::now();
        let mut work = (0, 0, 0, 0, 0);
        for pass in 0..MAX_REFERENCE_PASSES {
            let mut report = self.recompute_full_ordered_pass(handler);
            work.0 += report.cells_recomputed;
            work.1 += report.unknown_deps_recomputed;
            work.2 += report.phase_invalidation_us;
            work.3 += report.phase_topo_sort_us;
            work.4 += report.phase_eval_us;
            let rebound = self.apply_dynamic_references();
            // A former dynamic cycle can disappear when its selector changes.
            // Rebuild static edges once so the old runtime target cannot keep
            // preventing evaluation of the new selector indefinitely.
            if pass == 0
                && report.had_cycles
                && self.dep_graph.formula_cells().any(|cell| {
                    self.sheet_by_id(cell.sheet)
                        .and_then(|s| s.get_cell_opt(cell.row, cell.col))
                        .and_then(|c| c.value().formula_ast())
                        .is_some_and(crate::formula::analyze::has_dynamic_deps)
                })
            {
                self.rebuild_dep_graph();
                continue;
            }
            if rebound.is_empty() || pass + 1 == MAX_REFERENCE_PASSES {
                if !rebound.is_empty() {
                    report
                        .errors
                        .extend(rebound.into_iter().take(100).map(|cell| {
                            RecalcError::new(
                        cell, "dynamic references not settled after 16 passes; values may be stale")
                        }));
                }
                report.cells_recomputed = work.0;
                report.unknown_deps_recomputed = work.1;
                report.phase_invalidation_us = work.2;
                report.phase_topo_sort_us = work.3;
                report.phase_eval_us = work.4;
                report.duration_ms = start.elapsed().as_millis() as u64;
                return report;
            }
        }
        unreachable!()
    }

    pub(super) fn recalc_dirty_set(&mut self, changed: &[CellId]) -> Recalculated {
        let _recalc_scope = crate::custom_fns::recalc_scope();
        let _clock = crate::timing::ClockGuard::install(self.recalc_clock);
        #[cfg(test)]
        self.recalc_count.set(self.recalc_count.get() + 1);
        let first = self.recalc_dirty_pass(changed);
        let mut changed = self.apply_dynamic_references();
        if changed.is_empty() {
            return first;
        }
        let Recalculated::Cells(mut cells) = first else {
            return Recalculated::All;
        };
        let mut seen: FxHashSet<_> = cells.iter().copied().collect();
        for pass in 1..MAX_REFERENCE_PASSES {
            match self.recalc_dirty_pass(&changed) {
                Recalculated::All => return Recalculated::All,
                Recalculated::Cells(delta) => {
                    for cell in delta {
                        if seen.insert(cell) {
                            cells.push(cell);
                        }
                    }
                }
            }
            changed = self.apply_dynamic_references();
            if changed.is_empty() {
                break;
            }
            if pass + 1 == MAX_REFERENCE_PASSES {
                self.incremental_errors
                    .extend(changed.iter().take(100).map(|&cell| {
                        RecalcError::new(
                            cell,
                            "dynamic references not settled after 16 passes; values may be stale",
                        )
                    }));
            }
        }
        Recalculated::Cells(cells)
    }
}
