//! Dependency graph for formula cells.
//!
//! Tracks precedents (cells a formula depends on) and dependents (cells that
//! depend on a given cell) for efficient queries and future recomputation.
//!
//! # Edge Direction
//!
//! ```text
//! A → B  means  "B depends on A"  (A is a precedent of B)
//! ```
//!
//! This makes "what breaks if I change X?" trivial: follow outgoing edges.

use rustc_hash::{FxHashMap, FxHashSet};

use crate::cell_id::CellId;
use crate::recalc::CycleReport;

/// Persistent dependency graph for formula cells.
///
/// Maintains bidirectional adjacency for O(1) lookups:
/// - `preds[B]` = cells that B depends on (precedents)
/// - `succs[A]` = cells that depend on A (dependents)
///
/// # Invariants
///
/// 1. **Bidirectional consistency:** If A ∈ preds[B] then B ∈ succs[A], and vice versa.
/// 2. **No dangling entries:** Empty sets are removed, not stored.
/// 3. **No duplicate edges:** Set semantics enforced by FxHashSet.
/// 4. **Atomic updates:** `replace_edges` is the only mutator that touches both maps.
#[derive(Default, Debug, Clone)]
pub struct DepGraph {
    /// Precedents: for each formula cell B, the cells A it depends on.
    /// B -> {A1, A2, ...}
    preds: FxHashMap<CellId, FxHashSet<CellId>>,

    /// Dependents: for each referenced cell A, the formula cells B that depend on it.
    /// A -> {B1, B2, ...}
    succs: FxHashMap<CellId, FxHashSet<CellId>>,
}

impl DepGraph {
    /// Create an empty dependency graph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the cells this formula cell depends on (precedents).
    ///
    /// These are the incoming edges to the cell.
    pub fn precedents(&self, cell: CellId) -> impl Iterator<Item = CellId> + '_ {
        self.preds
            .get(&cell)
            .into_iter()
            .flat_map(|s| s.iter().copied())
    }

    /// Returns the cells that depend on this cell (dependents).
    ///
    /// These are the outgoing edges from the cell.
    pub fn dependents(&self, cell: CellId) -> impl Iterator<Item = CellId> + '_ {
        self.succs
            .get(&cell)
            .into_iter()
            .flat_map(|s| s.iter().copied())
    }

    /// Returns true if this cell has formula dependencies tracked in the graph.
    pub fn is_formula_cell(&self, cell: CellId) -> bool {
        self.preds.contains_key(&cell)
    }

    /// Returns the number of formula cells (cells with precedents) in the graph.
    pub fn formula_cell_count(&self) -> usize {
        self.preds.len()
    }

    /// Returns the number of cells that are referenced by at least one formula.
    pub fn referenced_cell_count(&self) -> usize {
        self.succs.len()
    }

    /// Replace all edges for a formula cell atomically.
    ///
    /// This is the primary mutation API. It:
    /// 1. Removes the cell from all its old precedents' successor sets
    /// 2. Clears the cell's precedent set
    /// 3. Adds the cell to all new precedents' successor sets
    /// 4. Sets the cell's new precedent set
    ///
    /// Pass an empty set to clear all edges for this cell.
    pub fn replace_edges(&mut self, formula_cell: CellId, new_preds: FxHashSet<CellId>) {
        // Step 1: Remove old edges
        if let Some(old_preds) = self.preds.remove(&formula_cell) {
            for pred in old_preds {
                if let Some(deps) = self.succs.get_mut(&pred) {
                    deps.remove(&formula_cell);
                    // Clean up empty entries (invariant: no dangling)
                    if deps.is_empty() {
                        self.succs.remove(&pred);
                    }
                }
            }
        }

        // Step 2: If no new precedents, we're done (cell is not a formula or has no refs)
        if new_preds.is_empty() {
            return;
        }

        // Step 3: Add new edges
        for pred in &new_preds {
            self.succs.entry(*pred).or_default().insert(formula_cell);
        }

        // Step 4: Store new precedents
        self.preds.insert(formula_cell, new_preds);
    }

    /// Clear all edges for a cell (formula removed or cell deleted).
    ///
    /// Convenience wrapper around `replace_edges` with an empty set.
    pub fn clear_cell(&mut self, cell: CellId) {
        self.replace_edges(cell, FxHashSet::default());
    }

    /// Remove all edges involving cells from a specific sheet.
    ///
    /// Called when a sheet is deleted.
    pub fn remove_sheet(&mut self, sheet: crate::sheet::SheetId) {
        // Collect cells to remove (can't mutate while iterating)
        let cells_to_remove: Vec<CellId> = self
            .preds
            .keys()
            .filter(|c| c.sheet == sheet)
            .copied()
            .collect();

        // Clear each formula cell from this sheet
        for cell in cells_to_remove {
            self.clear_cell(cell);
        }

        // Also remove any cells from this sheet that are only in succs
        // (cells that are referenced but don't have formulas)
        let referenced_to_remove: Vec<CellId> = self
            .succs
            .keys()
            .filter(|c| c.sheet == sheet)
            .copied()
            .collect();

        for cell in referenced_to_remove {
            if let Some(dependents) = self.succs.remove(&cell) {
                // Remove this cell from the preds of all its dependents
                for dep in dependents {
                    if let Some(preds) = self.preds.get_mut(&dep) {
                        preds.remove(&cell);
                        // Clean up empty preds (invariant: no empty sets stored)
                        if preds.is_empty() {
                            self.preds.remove(&dep);
                        }
                    }
                }
            }
        }
    }

    /// Apply a coordinate mapping to all cells in the graph.
    ///
    /// Used for row/column insert/delete operations. The mapping function
    /// returns `Some(new_id)` if the cell moves, or `None` if it's deleted.
    ///
    /// This rebuilds the graph with remapped coordinates.
    pub fn apply_mapping<F>(&mut self, map: F)
    where
        F: Fn(CellId) -> Option<CellId>,
    {
        // Build new maps with remapped IDs
        let mut new_preds: FxHashMap<CellId, FxHashSet<CellId>> = FxHashMap::default();
        let mut new_succs: FxHashMap<CellId, FxHashSet<CellId>> = FxHashMap::default();

        for (formula_cell, preds) in &self.preds {
            // Map the formula cell
            let Some(new_formula_cell) = map(*formula_cell) else {
                continue; // Formula cell was deleted
            };

            // Map all precedents, keeping only those that survive
            let mapped_preds: FxHashSet<CellId> = preds
                .iter()
                .filter_map(|p| map(*p))
                .collect();

            if mapped_preds.is_empty() {
                continue; // All precedents were deleted
            }

            // Add to new maps
            for pred in &mapped_preds {
                new_succs.entry(*pred).or_default().insert(new_formula_cell);
            }
            new_preds.insert(new_formula_cell, mapped_preds);
        }

        self.preds = new_preds;
        self.succs = new_succs;
    }

    // =========================================================================
    // Cycle Membership (Tarjan's SCC)
    // =========================================================================

    /// Find all cells that are members of true cycles (SCC size > 1 or self-loop).
    ///
    /// Uses Tarjan's algorithm. Only considers edges between formula cells.
    /// Iterates nodes in sorted order (by CellId) for deterministic output.
    ///
    /// Edge direction: walks `preds` (depends-on edges) — from cell X, follow
    /// `preds[X]` to find cells X references. This is the natural cycle direction.
    pub fn find_cycle_members(&self) -> FxHashSet<CellId> {
        let formula_cells: FxHashSet<CellId> = self.preds.keys().copied().collect();
        if formula_cells.is_empty() {
            return FxHashSet::default();
        }

        // Sorted iteration order for determinism
        let mut sorted_cells: Vec<CellId> = formula_cells.iter().copied().collect();
        sorted_cells.sort_by(|a, b| {
            a.sheet.raw().cmp(&b.sheet.raw())
                .then(a.row.cmp(&b.row))
                .then(a.col.cmp(&b.col))
        });

        // Tarjan's state
        let mut index_counter: u32 = 0;
        let mut stack: Vec<CellId> = Vec::new();
        let mut on_stack: FxHashSet<CellId> = FxHashSet::default();
        let mut indices: FxHashMap<CellId, u32> = FxHashMap::default();
        let mut lowlinks: FxHashMap<CellId, u32> = FxHashMap::default();
        let mut result: FxHashSet<CellId> = FxHashSet::default();

        // Helper: collect sorted neighbours (preds that are formula cells)
        let sorted_neighbours = |cell: CellId| -> Vec<CellId> {
            let mut neighbours: Vec<CellId> = self.preds
                .get(&cell)
                .into_iter()
                .flat_map(|s| s.iter().copied())
                .filter(|c| formula_cells.contains(c))
                .collect();
            neighbours.sort_by(|a, b| {
                a.sheet.raw().cmp(&b.sheet.raw())
                    .then(a.row.cmp(&b.row))
                    .then(a.col.cmp(&b.col))
            });
            neighbours
        };

        // Iterative Tarjan's to avoid stack overflow on deep graphs.
        struct DfsFrame {
            cell: CellId,
            neighbours: Vec<CellId>,
            next_idx: usize,
        }

        for &root in &sorted_cells {
            if indices.contains_key(&root) {
                continue;
            }

            let mut dfs_stack: Vec<DfsFrame> = Vec::new();

            // Start visiting root
            let idx = index_counter;
            index_counter += 1;
            indices.insert(root, idx);
            lowlinks.insert(root, idx);
            stack.push(root);
            on_stack.insert(root);

            dfs_stack.push(DfsFrame {
                cell: root,
                neighbours: sorted_neighbours(root),
                next_idx: 0,
            });

            while let Some(frame) = dfs_stack.last_mut() {
                if frame.next_idx < frame.neighbours.len() {
                    let w = frame.neighbours[frame.next_idx];
                    frame.next_idx += 1;

                    if !indices.contains_key(&w) {
                        // Recurse into w
                        let w_idx = index_counter;
                        index_counter += 1;
                        indices.insert(w, w_idx);
                        lowlinks.insert(w, w_idx);
                        stack.push(w);
                        on_stack.insert(w);

                        dfs_stack.push(DfsFrame {
                            cell: w,
                            neighbours: sorted_neighbours(w),
                            next_idx: 0,
                        });
                    } else if on_stack.contains(&w) {
                        let w_idx = indices[&w];
                        let v_low = lowlinks.get_mut(&frame.cell).unwrap();
                        if w_idx < *v_low {
                            *v_low = w_idx;
                        }
                    }
                } else {
                    // All neighbours explored — pop and propagate lowlink
                    let finished = dfs_stack.pop().unwrap();
                    let v = finished.cell;
                    let v_low = lowlinks[&v];
                    let v_idx = indices[&v];

                    // Propagate lowlink to parent
                    if let Some(parent) = dfs_stack.last() {
                        let parent_low = lowlinks.get_mut(&parent.cell).unwrap();
                        if v_low < *parent_low {
                            *parent_low = v_low;
                        }
                    }

                    // SCC root check
                    if v_low == v_idx {
                        // Pop SCC from stack
                        let mut scc = Vec::new();
                        loop {
                            let w = stack.pop().unwrap();
                            on_stack.remove(&w);
                            scc.push(w);
                            if w == v {
                                break;
                            }
                        }

                        // Include SCC if size > 1, or size == 1 with self-loop
                        if scc.len() > 1 {
                            result.extend(scc);
                        } else if scc.len() == 1 {
                            let cell = scc[0];
                            if self.preds.get(&cell).map_or(false, |p| p.contains(&cell)) {
                                result.insert(cell);
                            }
                        }
                    }
                }
            }
        }

        result
    }

    /// Find all non-trivial SCCs (cycle groups), returned as separate groups.
    ///
    /// Each inner Vec is one SCC (size > 1, or size == 1 with self-loop).
    /// Uses the same iterative Tarjan's algorithm as `find_cycle_members`.
    /// SCCs within each group are sorted by (sheet, row, col) for determinism.
    pub fn find_cycle_sccs(&self) -> Vec<Vec<CellId>> {
        let formula_cells: FxHashSet<CellId> = self.preds.keys().copied().collect();
        if formula_cells.is_empty() {
            return Vec::new();
        }

        let mut sorted_cells: Vec<CellId> = formula_cells.iter().copied().collect();
        sorted_cells.sort_by(|a, b| {
            a.sheet.raw().cmp(&b.sheet.raw())
                .then(a.row.cmp(&b.row))
                .then(a.col.cmp(&b.col))
        });

        let mut index_counter: u32 = 0;
        let mut stack: Vec<CellId> = Vec::new();
        let mut on_stack: FxHashSet<CellId> = FxHashSet::default();
        let mut indices: FxHashMap<CellId, u32> = FxHashMap::default();
        let mut lowlinks: FxHashMap<CellId, u32> = FxHashMap::default();
        let mut sccs: Vec<Vec<CellId>> = Vec::new();

        let sorted_neighbours = |cell: CellId| -> Vec<CellId> {
            let mut neighbours: Vec<CellId> = self.preds
                .get(&cell)
                .into_iter()
                .flat_map(|s| s.iter().copied())
                .filter(|c| formula_cells.contains(c))
                .collect();
            neighbours.sort_by(|a, b| {
                a.sheet.raw().cmp(&b.sheet.raw())
                    .then(a.row.cmp(&b.row))
                    .then(a.col.cmp(&b.col))
            });
            neighbours
        };

        struct DfsFrame {
            cell: CellId,
            neighbours: Vec<CellId>,
            next_idx: usize,
        }

        for &root in &sorted_cells {
            if indices.contains_key(&root) {
                continue;
            }

            let mut dfs_stack: Vec<DfsFrame> = Vec::new();

            let idx = index_counter;
            index_counter += 1;
            indices.insert(root, idx);
            lowlinks.insert(root, idx);
            stack.push(root);
            on_stack.insert(root);

            dfs_stack.push(DfsFrame {
                cell: root,
                neighbours: sorted_neighbours(root),
                next_idx: 0,
            });

            while let Some(frame) = dfs_stack.last_mut() {
                if frame.next_idx < frame.neighbours.len() {
                    let w = frame.neighbours[frame.next_idx];
                    frame.next_idx += 1;

                    if !indices.contains_key(&w) {
                        let w_idx = index_counter;
                        index_counter += 1;
                        indices.insert(w, w_idx);
                        lowlinks.insert(w, w_idx);
                        stack.push(w);
                        on_stack.insert(w);

                        dfs_stack.push(DfsFrame {
                            cell: w,
                            neighbours: sorted_neighbours(w),
                            next_idx: 0,
                        });
                    } else if on_stack.contains(&w) {
                        let w_idx = indices[&w];
                        let v_low = lowlinks.get_mut(&frame.cell).unwrap();
                        if w_idx < *v_low {
                            *v_low = w_idx;
                        }
                    }
                } else {
                    let finished = dfs_stack.pop().unwrap();
                    let v = finished.cell;
                    let v_low = lowlinks[&v];
                    let v_idx = indices[&v];

                    if let Some(parent) = dfs_stack.last() {
                        let parent_low = lowlinks.get_mut(&parent.cell).unwrap();
                        if v_low < *parent_low {
                            *parent_low = v_low;
                        }
                    }

                    if v_low == v_idx {
                        let mut scc = Vec::new();
                        loop {
                            let w = stack.pop().unwrap();
                            on_stack.remove(&w);
                            scc.push(w);
                            if w == v {
                                break;
                            }
                        }

                        let is_cycle = if scc.len() > 1 {
                            true
                        } else {
                            // size == 1: only a cycle if self-loop
                            let cell = scc[0];
                            self.preds.get(&cell).map_or(false, |p| p.contains(&cell))
                        };

                        if is_cycle {
                            scc.sort_by(|a, b| {
                                a.sheet.raw().cmp(&b.sheet.raw())
                                    .then(a.row.cmp(&b.row))
                                    .then(a.col.cmp(&b.col))
                            });
                            sccs.push(scc);
                        }
                    }
                }
            }
        }

        sccs
    }

    // =========================================================================
    // Topological Ordering + Cycle Detection (Phase 1.2)
    // =========================================================================

    /// Returns all formula cells in the graph.
    ///
    /// A cell is a "formula cell" if it has precedents (appears in preds keys).
    pub fn formula_cells(&self) -> impl Iterator<Item = CellId> + '_ {
        self.preds.keys().copied()
    }

    /// Compute topological order of all formula cells.
    ///
    /// Returns cells in dependency order: precedents before dependents.
    /// Uses Kahn's algorithm with stable ordering for determinism.
    ///
    /// # Returns
    ///
    /// - `Ok(order)` - Valid topological order
    /// - `Err(CycleReport)` - Graph contains cycles
    ///
    /// # Algorithm
    ///
    /// Only considers edges between formula cells. Value-only cells (cells with
    /// no formula) are not included in the ordering since they don't need
    /// recomputation.
    pub fn topo_order_all_formulas(&self) -> Result<Vec<CellId>, CycleReport> {
        // Collect all formula cells
        let formula_cells: FxHashSet<CellId> = self.preds.keys().copied().collect();

        if formula_cells.is_empty() {
            return Ok(Vec::new());
        }

        // Compute in-degree for each formula cell
        // Only count edges from precedents that are ALSO formula cells
        let mut in_degree: FxHashMap<CellId, usize> = FxHashMap::default();

        for &cell in &formula_cells {
            let count = self
                .preds
                .get(&cell)
                .map(|preds| preds.iter().filter(|p| formula_cells.contains(p)).count())
                .unwrap_or(0);
            in_degree.insert(cell, count);
        }

        // Initialize queue with zero in-degree cells
        // Sort for deterministic order
        let mut queue: Vec<CellId> = in_degree
            .iter()
            .filter(|(_, &deg)| deg == 0)
            .map(|(&cell, _)| cell)
            .collect();
        // Sort in DESCENDING order so smallest is at end (popped first)
        queue.sort_by(|a, b| {
            b.sheet
                .raw()
                .cmp(&a.sheet.raw())
                .then(b.row.cmp(&a.row))
                .then(b.col.cmp(&a.col))
        });

        let mut result = Vec::with_capacity(formula_cells.len());

        while let Some(cell) = queue.pop() {
            result.push(cell);

            // For each dependent that is a formula cell
            if let Some(deps) = self.succs.get(&cell) {
                let mut new_zero_degree = Vec::new();

                for &dep in deps {
                    if formula_cells.contains(&dep) {
                        if let Some(deg) = in_degree.get_mut(&dep) {
                            *deg = deg.saturating_sub(1);
                            if *deg == 0 {
                                new_zero_degree.push(dep);
                            }
                        }
                    }
                }

                // Sort new zero-degree cells for deterministic order
                new_zero_degree.sort_by(|a, b| {
                    a.sheet
                        .raw()
                        .cmp(&b.sheet.raw())
                        .then(a.row.cmp(&b.row))
                        .then(a.col.cmp(&b.col))
                });
                // Add in reverse order so smallest is popped first
                for cell in new_zero_degree.into_iter().rev() {
                    queue.push(cell);
                }
            }
        }

        // If not all cells are in result, we have a cycle
        if result.len() < formula_cells.len() {
            // Find cells involved in cycle
            let cycle_cells: Vec<CellId> = formula_cells
                .iter()
                .filter(|c| !result.contains(c))
                .copied()
                .collect();
            return Err(CycleReport::cycle(cycle_cells));
        }

        Ok(result)
    }

    /// Check if adding edges from `cell` to `new_preds` would create a cycle.
    ///
    /// Does not modify the graph. Returns `Some(CycleReport)` if a cycle would
    /// be introduced, `None` otherwise.
    ///
    /// # Algorithm
    ///
    /// A cycle is created if any of `new_preds` can reach `cell` by following
    /// dependent edges. We do a DFS from `cell` following successors and check
    /// if we can reach any of `new_preds`.
    pub fn would_create_cycle(&self, cell: CellId, new_preds: &[CellId]) -> Option<CycleReport> {
        // Self-reference check
        if new_preds.contains(&cell) {
            return Some(CycleReport::self_reference(cell));
        }

        // DFS from cell following dependents to see if we reach any new_pred
        let new_preds_set: FxHashSet<CellId> = new_preds.iter().copied().collect();
        let mut visited = FxHashSet::default();
        let mut stack = vec![cell];

        while let Some(current) = stack.pop() {
            if !visited.insert(current) {
                continue;
            }

            if let Some(deps) = self.succs.get(&current) {
                for &dep in deps {
                    if new_preds_set.contains(&dep) {
                        // Found a path from cell to one of its would-be precedents
                        // This means new_pred -> ... -> cell -> new_pred (cycle!)
                        return Some(CycleReport::cycle(vec![dep, cell]));
                    }
                    stack.push(dep);
                }
            }
        }

        None
    }

    /// Check all invariants. Panics if any are violated.
    ///
    /// Only available in test builds.
    #[cfg(test)]
    pub fn assert_consistent(&self) {
        // Invariant 1: Bidirectional consistency (preds → succs)
        for (formula_cell, preds) in &self.preds {
            for pred in preds {
                assert!(
                    self.succs.get(pred).map_or(false, |s| s.contains(formula_cell)),
                    "Missing succ edge: {:?} should have {:?} in dependents",
                    pred,
                    formula_cell
                );
            }
        }

        // Invariant 1: Bidirectional consistency (succs → preds)
        for (cell, dependents) in &self.succs {
            for dep in dependents {
                assert!(
                    self.preds.get(dep).map_or(false, |s| s.contains(cell)),
                    "Missing pred edge: {:?} should have {:?} in precedents",
                    dep,
                    cell
                );
            }
        }

        // Invariant 2: No empty sets stored
        for (cell, preds) in &self.preds {
            assert!(
                !preds.is_empty(),
                "Empty preds set stored for {:?}",
                cell
            );
        }
        for (cell, succs) in &self.succs {
            assert!(
                !succs.is_empty(),
                "Empty succs set stored for {:?}",
                cell
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::SheetId;

    fn cell(sheet: u64, row: usize, col: usize) -> CellId {
        CellId::new(SheetId::from_raw(sheet), row, col)
    }

    fn set(cells: &[CellId]) -> FxHashSet<CellId> {
        cells.iter().copied().collect()
    }

    #[test]
    fn test_empty_graph() {
        let graph = DepGraph::new();

        assert_eq!(graph.formula_cell_count(), 0);
        assert_eq!(graph.referenced_cell_count(), 0);
        assert!(!graph.is_formula_cell(cell(1, 0, 0)));
        assert_eq!(graph.precedents(cell(1, 0, 0)).count(), 0);
        assert_eq!(graph.dependents(cell(1, 0, 0)).count(), 0);

        graph.assert_consistent();
    }

    #[test]
    fn test_single_edge() {
        // B1 = A1
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(b1, set(&[a1]));
        graph.assert_consistent();

        // B1 depends on A1
        assert!(graph.is_formula_cell(b1));
        assert!(!graph.is_formula_cell(a1));

        let preds: Vec<_> = graph.precedents(b1).collect();
        assert_eq!(preds, vec![a1]);

        let deps: Vec<_> = graph.dependents(a1).collect();
        assert_eq!(deps, vec![b1]);

        assert_eq!(graph.formula_cell_count(), 1);
        assert_eq!(graph.referenced_cell_count(), 1);
    }

    #[test]
    fn test_multiple_precedents() {
        // C1 = A1 + B1
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);
        let c1 = cell(1, 0, 2);

        graph.replace_edges(c1, set(&[a1, b1]));
        graph.assert_consistent();

        let mut preds: Vec<_> = graph.precedents(c1).collect();
        preds.sort_by_key(|c| c.col);
        assert_eq!(preds, vec![a1, b1]);

        assert_eq!(graph.dependents(a1).collect::<Vec<_>>(), vec![c1]);
        assert_eq!(graph.dependents(b1).collect::<Vec<_>>(), vec![c1]);
    }

    #[test]
    fn test_multiple_dependents() {
        // B1 = A1, C1 = A1
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);
        let c1 = cell(1, 0, 2);

        graph.replace_edges(b1, set(&[a1]));
        graph.replace_edges(c1, set(&[a1]));
        graph.assert_consistent();

        let mut deps: Vec<_> = graph.dependents(a1).collect();
        deps.sort_by_key(|c| c.col);
        assert_eq!(deps, vec![b1, c1]);

        assert_eq!(graph.formula_cell_count(), 2);
        assert_eq!(graph.referenced_cell_count(), 1);
    }

    #[test]
    fn test_rewiring() {
        // B1 = A1, then change to B1 = A2
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let a2 = cell(1, 1, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(b1, set(&[a1]));
        graph.assert_consistent();

        assert_eq!(graph.precedents(b1).collect::<Vec<_>>(), vec![a1]);
        assert_eq!(graph.dependents(a1).collect::<Vec<_>>(), vec![b1]);

        // Rewire: B1 now depends on A2 instead
        graph.replace_edges(b1, set(&[a2]));
        graph.assert_consistent();

        assert_eq!(graph.precedents(b1).collect::<Vec<_>>(), vec![a2]);
        assert_eq!(graph.dependents(a2).collect::<Vec<_>>(), vec![b1]);

        // A1 should have no dependents now
        assert_eq!(graph.dependents(a1).count(), 0);
        // And A1 should not be in succs at all (sparse)
        assert!(!graph.succs.contains_key(&a1));
    }

    #[test]
    fn test_unwiring() {
        // B1 = A1, then clear B1
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(b1, set(&[a1]));
        graph.assert_consistent();

        graph.clear_cell(b1);
        graph.assert_consistent();

        assert!(!graph.is_formula_cell(b1));
        assert_eq!(graph.precedents(b1).count(), 0);
        assert_eq!(graph.dependents(a1).count(), 0);
        assert_eq!(graph.formula_cell_count(), 0);
        assert_eq!(graph.referenced_cell_count(), 0);
    }

    #[test]
    fn test_cross_sheet_edge() {
        // Sheet2!A1 = Sheet1!B1
        let mut graph = DepGraph::new();
        let sheet1_b1 = cell(1, 0, 1);
        let sheet2_a1 = cell(2, 0, 0);

        graph.replace_edges(sheet2_a1, set(&[sheet1_b1]));
        graph.assert_consistent();

        assert!(graph.is_formula_cell(sheet2_a1));
        assert_eq!(graph.precedents(sheet2_a1).collect::<Vec<_>>(), vec![sheet1_b1]);
        assert_eq!(graph.dependents(sheet1_b1).collect::<Vec<_>>(), vec![sheet2_a1]);
    }

    #[test]
    fn test_diamond_dependency() {
        //     A1
        //    /  \
        //   B1   C1
        //    \  /
        //     D1
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);
        let c1 = cell(1, 0, 2);
        let d1 = cell(1, 0, 3);

        graph.replace_edges(b1, set(&[a1]));
        graph.replace_edges(c1, set(&[a1]));
        graph.replace_edges(d1, set(&[b1, c1]));
        graph.assert_consistent();

        // D1 depends on B1 and C1
        let mut d1_preds: Vec<_> = graph.precedents(d1).collect();
        d1_preds.sort_by_key(|c| c.col);
        assert_eq!(d1_preds, vec![b1, c1]);

        // A1 has B1 and C1 as dependents
        let mut a1_deps: Vec<_> = graph.dependents(a1).collect();
        a1_deps.sort_by_key(|c| c.col);
        assert_eq!(a1_deps, vec![b1, c1]);

        assert_eq!(graph.formula_cell_count(), 3); // B1, C1, D1
        assert_eq!(graph.referenced_cell_count(), 3); // A1, B1, C1
    }

    #[test]
    fn test_self_reference() {
        // A1 = A1 + 1 (cycle, but graph allows it - cycle detection is Phase 1.2)
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);

        graph.replace_edges(a1, set(&[a1]));
        graph.assert_consistent();

        assert!(graph.is_formula_cell(a1));
        assert_eq!(graph.precedents(a1).collect::<Vec<_>>(), vec![a1]);
        assert_eq!(graph.dependents(a1).collect::<Vec<_>>(), vec![a1]);
    }

    #[test]
    fn test_remove_sheet() {
        // Sheet1: B1 = A1
        // Sheet2: A1 = Sheet1!B1
        let mut graph = DepGraph::new();
        let s1_a1 = cell(1, 0, 0);
        let s1_b1 = cell(1, 0, 1);
        let s2_a1 = cell(2, 0, 0);

        graph.replace_edges(s1_b1, set(&[s1_a1]));
        graph.replace_edges(s2_a1, set(&[s1_b1]));
        graph.assert_consistent();

        assert_eq!(graph.formula_cell_count(), 2);

        // Delete Sheet1
        graph.remove_sheet(SheetId::from_raw(1));
        graph.assert_consistent();

        // Sheet1 cells should be gone
        assert!(!graph.is_formula_cell(s1_b1));
        assert_eq!(graph.dependents(s1_a1).count(), 0);

        // Sheet2!A1's precedent was on Sheet1, so it now has no precedents
        // and is removed from the graph (empty preds cleaned up)
        assert!(!graph.is_formula_cell(s2_a1));
        assert_eq!(graph.formula_cell_count(), 0);
        assert_eq!(graph.referenced_cell_count(), 0);
    }

    #[test]
    fn test_apply_mapping_shift_rows() {
        // B1 = A1, B2 = A2
        // Insert row at 1, so row 0 stays, row 1+ shifts down
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let a2 = cell(1, 1, 0);
        let b1 = cell(1, 0, 1);
        let b2 = cell(1, 1, 1);

        graph.replace_edges(b1, set(&[a1]));
        graph.replace_edges(b2, set(&[a2]));
        graph.assert_consistent();

        // Insert row at index 1: rows >= 1 shift by +1
        graph.apply_mapping(|c| {
            if c.sheet.raw() != 1 {
                return Some(c);
            }
            if c.row >= 1 {
                Some(CellId::new(c.sheet, c.row + 1, c.col))
            } else {
                Some(c)
            }
        });
        graph.assert_consistent();

        // B1 = A1 should be unchanged
        assert!(graph.is_formula_cell(b1));
        assert_eq!(graph.precedents(b1).collect::<Vec<_>>(), vec![a1]);

        // B2 = A2 should now be B3 = A3
        let a3 = cell(1, 2, 0);
        let b3 = cell(1, 2, 1);
        assert!(!graph.is_formula_cell(b2)); // Old position gone
        assert!(graph.is_formula_cell(b3));  // New position exists
        assert_eq!(graph.precedents(b3).collect::<Vec<_>>(), vec![a3]);
    }

    #[test]
    fn test_apply_mapping_delete_row() {
        // B1 = A1, B2 = A2
        // Delete row 0
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let a2 = cell(1, 1, 0);
        let b1 = cell(1, 0, 1);
        let b2 = cell(1, 1, 1);

        graph.replace_edges(b1, set(&[a1]));
        graph.replace_edges(b2, set(&[a2]));
        graph.assert_consistent();

        // Delete row 0: row 0 → None, rows > 0 shift by -1
        graph.apply_mapping(|c| {
            if c.sheet.raw() != 1 {
                return Some(c);
            }
            if c.row == 0 {
                None // Deleted
            } else {
                Some(CellId::new(c.sheet, c.row - 1, c.col))
            }
        });
        graph.assert_consistent();

        // Original B1 (row 0) was deleted, original B2 (row 1) shifted to row 0
        // So position (0,1) now contains the shifted B2 formula
        // Original b2 position (row 1) should be empty
        assert!(!graph.is_formula_cell(b2)); // Old row 1 position is gone

        // The shifted formula is now at row 0
        let new_a1 = cell(1, 0, 0); // What was A2 is now at row 0
        let new_b1 = cell(1, 0, 1); // What was B2 is now at row 0
        assert!(graph.is_formula_cell(new_b1));
        assert_eq!(graph.precedents(new_b1).collect::<Vec<_>>(), vec![new_a1]);

        // Only one formula cell should exist now
        assert_eq!(graph.formula_cell_count(), 1);
    }

    // =========================================================================
    // Topo Order + Cycle Detection Tests (Phase 1.2)
    // =========================================================================

    #[test]
    fn test_topo_empty_graph() {
        let graph = DepGraph::new();
        let order = graph.topo_order_all_formulas().unwrap();
        assert!(order.is_empty());
    }

    #[test]
    fn test_topo_single_formula() {
        // B1 = A1 (A1 is a value cell, B1 is formula)
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(b1, set(&[a1]));

        let order = graph.topo_order_all_formulas().unwrap();
        assert_eq!(order, vec![b1]); // Only formula cell
    }

    #[test]
    fn test_topo_chain() {
        // A → B → C → D (chain of formulas, A is value)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);
        let d = cell(1, 0, 3);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));
        graph.replace_edges(d, set(&[c]));

        let order = graph.topo_order_all_formulas().unwrap();
        assert_eq!(order, vec![b, c, d]);
    }

    #[test]
    fn test_topo_diamond() {
        // A → B, A → C, B → D, C → D
        //     A (value)
        //    / \
        //   B   C  (formulas)
        //    \ /
        //     D    (formula)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);
        let d = cell(1, 0, 3);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[a]));
        graph.replace_edges(d, set(&[b, c]));

        let order = graph.topo_order_all_formulas().unwrap();

        // B and C can be in either order, but both must come before D
        assert!(order.len() == 3);
        let d_pos = order.iter().position(|&x| x == d).unwrap();
        let b_pos = order.iter().position(|&x| x == b).unwrap();
        let c_pos = order.iter().position(|&x| x == c).unwrap();
        assert!(b_pos < d_pos);
        assert!(c_pos < d_pos);
    }

    #[test]
    fn test_topo_wide_fanout() {
        // A → {B1, B2, B3, B4, B5} (A is value)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);

        for col in 1..=5 {
            let b = cell(1, 0, col);
            graph.replace_edges(b, set(&[a]));
        }

        let order = graph.topo_order_all_formulas().unwrap();
        assert_eq!(order.len(), 5);
    }

    #[test]
    fn test_topo_cross_sheet() {
        // Sheet1!A1 → Sheet2!B1 → Sheet1!C1
        let mut graph = DepGraph::new();
        let s1_a1 = cell(1, 0, 0); // value
        let s2_b1 = cell(2, 0, 1); // formula
        let s1_c1 = cell(1, 0, 2); // formula

        graph.replace_edges(s2_b1, set(&[s1_a1]));
        graph.replace_edges(s1_c1, set(&[s2_b1]));

        let order = graph.topo_order_all_formulas().unwrap();
        assert_eq!(order, vec![s2_b1, s1_c1]);
    }

    #[test]
    fn test_topo_stable_order() {
        // Multiple independent formulas should have deterministic order
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);
        let b2 = cell(1, 0, 2);
        let b3 = cell(1, 0, 3);

        graph.replace_edges(b3, set(&[a]));
        graph.replace_edges(b1, set(&[a]));
        graph.replace_edges(b2, set(&[a]));

        // Run multiple times, should always get same order
        let order1 = graph.topo_order_all_formulas().unwrap();
        let order2 = graph.topo_order_all_formulas().unwrap();
        let order3 = graph.topo_order_all_formulas().unwrap();

        assert_eq!(order1, order2);
        assert_eq!(order2, order3);

        // Order should be by (sheet, row, col)
        assert_eq!(order1, vec![b1, b2, b3]);
    }

    #[test]
    fn test_cycle_self_reference() {
        // A1 = A1 (self reference)
        let graph = DepGraph::new();
        let a1 = cell(1, 0, 0);

        let result = graph.would_create_cycle(a1, &[a1]);
        assert!(result.is_some());

        let cycle = result.unwrap();
        assert!(cycle.message.contains("references itself"));
    }

    #[test]
    fn test_cycle_two_cell() {
        // A1 = B1, then B1 = A1 (creates cycle)
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(a1, set(&[b1]));

        // Now trying to make B1 depend on A1 should detect cycle
        let result = graph.would_create_cycle(b1, &[a1]);
        assert!(result.is_some());
    }

    #[test]
    fn test_cycle_indirect() {
        // A → B → C, then C → A (creates cycle)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));

        // Trying to make A depend on C should detect cycle
        let result = graph.would_create_cycle(a, &[c]);
        assert!(result.is_some());
    }

    #[test]
    fn test_cycle_detection_in_topo() {
        // Create a graph with an existing cycle
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);

        // Force a cycle by directly manipulating (simulating corrupted file)
        graph.replace_edges(a, set(&[b]));
        graph.replace_edges(b, set(&[a]));

        let result = graph.topo_order_all_formulas();
        assert!(result.is_err());

        let cycle = result.unwrap_err();
        assert!(!cycle.cells.is_empty());
    }

    #[test]
    fn test_no_cycle_valid_graph() {
        // A → B → C (valid, no cycle)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));

        // Adding D → C should be fine
        let d = cell(1, 0, 3);
        let result = graph.would_create_cycle(d, &[c]);
        assert!(result.is_none());
    }

    // =========================================================================
    // Tarjan's SCC (find_cycle_members) Tests
    // =========================================================================

    #[test]
    fn test_cycle_members_two_node_cycle() {
        // A1 = B1, B1 = A1 → both flagged
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);

        graph.replace_edges(a1, set(&[b1]));
        graph.replace_edges(b1, set(&[a1]));

        let members = graph.find_cycle_members();
        assert!(members.contains(&a1));
        assert!(members.contains(&b1));
        assert_eq!(members.len(), 2);
    }

    #[test]
    fn test_cycle_members_self_loop() {
        // A1 = A1 → flagged
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);

        graph.replace_edges(a1, set(&[a1]));

        let members = graph.find_cycle_members();
        assert!(members.contains(&a1));
        assert_eq!(members.len(), 1);
    }

    #[test]
    fn test_cycle_members_downstream_excluded() {
        // A1 = B1, B1 = A1 (cycle), C1 depends on A1 (downstream)
        // C1 should NOT be in cycle members
        let mut graph = DepGraph::new();
        let a1 = cell(1, 0, 0);
        let b1 = cell(1, 0, 1);
        let c1 = cell(1, 0, 2);

        graph.replace_edges(a1, set(&[b1]));
        graph.replace_edges(b1, set(&[a1]));
        graph.replace_edges(c1, set(&[a1]));

        let members = graph.find_cycle_members();
        assert!(members.contains(&a1));
        assert!(members.contains(&b1));
        assert!(!members.contains(&c1), "Downstream cell C1 should NOT be in cycle members");
        assert_eq!(members.len(), 2);
    }

    #[test]
    fn test_cycle_members_no_cycles() {
        // A → B → C (no cycles)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));

        let members = graph.find_cycle_members();
        assert!(members.is_empty());
    }

    #[test]
    fn test_cycle_members_three_node_cycle() {
        // A → B → C → A
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);

        graph.replace_edges(a, set(&[c]));
        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));

        let members = graph.find_cycle_members();
        assert_eq!(members.len(), 3);
        assert!(members.contains(&a));
        assert!(members.contains(&b));
        assert!(members.contains(&c));
    }

    #[test]
    fn test_cycle_members_stability() {
        // Run find_cycle_members twice on same graph → same set
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);

        graph.replace_edges(a, set(&[b]));
        graph.replace_edges(b, set(&[a]));

        let members1 = graph.find_cycle_members();
        let members2 = graph.find_cycle_members();
        assert_eq!(members1, members2);
    }

    #[test]
    fn test_cycle_members_empty_graph() {
        let graph = DepGraph::new();
        let members = graph.find_cycle_members();
        assert!(members.is_empty());
    }

    #[test]
    fn test_cycle_members_mixed_cycle_and_acyclic() {
        // A ↔ B (cycle), C → D (acyclic chain), E → A (downstream of cycle)
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);
        let d = cell(1, 0, 3);
        let e = cell(1, 0, 4);

        graph.replace_edges(a, set(&[b]));
        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(d, set(&[c]));
        graph.replace_edges(e, set(&[a]));

        let members = graph.find_cycle_members();
        assert_eq!(members.len(), 2);
        assert!(members.contains(&a));
        assert!(members.contains(&b));
        assert!(!members.contains(&c)); // c is a value cell (not in preds)
        assert!(!members.contains(&d)); // acyclic
        assert!(!members.contains(&e)); // downstream
    }

    #[test]
    fn test_topo_depth_order() {
        // Verify that cells are ordered by dependency depth
        // A (value) → B → C → D → E
        let mut graph = DepGraph::new();
        let a = cell(1, 0, 0);
        let b = cell(1, 0, 1);
        let c = cell(1, 0, 2);
        let d = cell(1, 0, 3);
        let e = cell(1, 0, 4);

        graph.replace_edges(b, set(&[a]));
        graph.replace_edges(c, set(&[b]));
        graph.replace_edges(d, set(&[c]));
        graph.replace_edges(e, set(&[d]));

        let order = graph.topo_order_all_formulas().unwrap();

        // Each cell must come before its dependents
        for i in 0..order.len() {
            for j in (i + 1)..order.len() {
                // Cell at i should not depend on cell at j
                let cell_i = order[i];
                let cell_j = order[j];
                assert!(
                    !graph.preds.get(&cell_i).map_or(false, |p| p.contains(&cell_j)),
                    "{:?} at position {} depends on {:?} at position {}",
                    cell_i, i, cell_j, j
                );
            }
        }
    }
}
