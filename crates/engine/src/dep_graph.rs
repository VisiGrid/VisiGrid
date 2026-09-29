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
//!
//! # Ranges (#29)
//!
//! A range reference is not expanded into one edge per cell. That made a
//! running total (`=SUM($A$1:A1)` filled down) cost n²/2 edges: 8,000 rows
//! built a 2.5 GiB graph in 32 s. Instead:
//!
//! - **Ordering** only needs edges between formulas, so the graph keeps a
//!   concrete edge from each *formula* inside a range to the formula reading
//!   it. A running total over values has none.
//! - **"What reads this cell?"** is answered by [`RangeIndex`], an interval
//!   index over every range subscription, so dirty propagation still reaches
//!   a formula when a value inside its range changes.

use std::collections::BTreeSet;

use rustc_hash::{FxHashMap, FxHashSet};

use crate::cell_id::CellId;
use crate::recalc::CycleReport;
use crate::sheet::SheetId;

/// A rectangular reference, inclusive on both ends. Whole-row and
/// whole-column references are rectangles spanning the grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RangeRef {
    pub sheet: SheetId,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

impl RangeRef {
    pub fn contains(&self, cell: CellId) -> bool {
        cell.sheet == self.sheet
            && (self.start_row..=self.end_row).contains(&cell.row)
            && (self.start_col..=self.end_col).contains(&cell.col)
    }

    /// From a whole-row or whole-column reference.
    pub fn from_whole(range: &crate::formula::whole_range::WholeRangeRef) -> RangeRef {
        use crate::formula::parser::RangeAxis;
        match range.axis {
            RangeAxis::Column => RangeRef {
                sheet: range.sheet,
                start_row: 0,
                end_row: crate::sheet::NUM_ROWS - 1,
                start_col: range.start,
                end_col: range.end,
            },
            RangeAxis::Row => RangeRef {
                sheet: range.sheet,
                start_row: range.start,
                end_row: range.end,
                start_col: 0,
                end_col: crate::sheet::NUM_COLS - 1,
            },
        }
    }

    fn height(&self) -> usize {
        self.end_row - self.start_row + 1
    }

    fn width(&self) -> usize {
        self.end_col - self.start_col + 1
    }
}

/// Positions per tree axis: 2^20 covers both 1,048,576 rows and 16,384 columns.
const TREE_BITS: u32 = 20;

/// Canonical segment-tree nodes covering `[lo, hi]` in a tree over
/// `[0, 2^TREE_BITS)`, as heap indices (root 1). At most ~2 per level.
fn cover(lo: usize, hi: usize, out: &mut Vec<u32>) {
    fn go(node: u32, nlo: usize, nhi: usize, lo: usize, hi: usize, out: &mut Vec<u32>) {
        if hi < nlo || nhi < lo {
            return;
        }
        if lo <= nlo && nhi <= hi {
            out.push(node);
            return;
        }
        let mid = nlo + (nhi - nlo) / 2;
        go(node * 2, nlo, mid, lo, hi, out);
        go(node * 2 + 1, mid + 1, nhi, lo, hi, out);
    }
    go(1, 0, (1usize << TREE_BITS) - 1, lo, hi, out);
}

/// The nodes on the root-to-leaf path of position `p`.
fn path(p: usize) -> impl Iterator<Item = u32> {
    (0..=TREE_BITS).map(move |depth| ((1usize << depth) | (p >> (TREE_BITS - depth))) as u32)
}

/// One sparse segment tree: node -> formulas whose interval covers it.
type Tree = FxHashMap<u32, FxHashSet<CellId>>;

/// Which formulas read which rectangles, queryable by cell in
/// O(log rows + answers).
///
/// A rectangle is stored along its shorter side: a tall range in one
/// row-interval tree per column it spans, a wide one (a whole-row
/// reference) in one column-interval tree per row. A point query checks
/// both trees that could hold it.
#[derive(Default, Debug, Clone)]
pub struct RangeIndex {
    /// (sheet, column) -> tree over rows
    by_col: FxHashMap<(SheetId, u32), Tree>,
    /// (sheet, row) -> tree over columns
    by_row: FxHashMap<(SheetId, u32), Tree>,
}

impl RangeIndex {
    fn tall(range: &RangeRef) -> bool {
        range.width() <= range.height()
    }

    fn insert(&mut self, formula: CellId, range: &RangeRef) {
        self.visit(range, |trees, key, nodes| {
            let tree = trees.entry(key).or_default();
            for node in nodes {
                tree.entry(*node).or_default().insert(formula);
            }
        });
    }

    fn remove(&mut self, formula: CellId, range: &RangeRef) {
        self.visit(range, |trees, key, nodes| {
            if let Some(tree) = trees.get_mut(&key) {
                for node in nodes {
                    if let Some(set) = tree.get_mut(node) {
                        set.remove(&formula);
                        if set.is_empty() {
                            tree.remove(node);
                        }
                    }
                }
                if tree.is_empty() {
                    trees.remove(&key);
                }
            }
        });
    }

    /// Call `f` once per tree the range is stored in, with its covering nodes.
    fn visit(&mut self, range: &RangeRef, mut f: impl FnMut(&mut FxHashMap<(SheetId, u32), Tree>, (SheetId, u32), &[u32])) {
        let mut nodes = Vec::new();
        if Self::tall(range) {
            cover(range.start_row, range.end_row, &mut nodes);
            for col in range.start_col..=range.end_col {
                f(&mut self.by_col, (range.sheet, col as u32), &nodes);
            }
        } else {
            cover(range.start_col, range.end_col, &mut nodes);
            for row in range.start_row..=range.end_row {
                f(&mut self.by_row, (range.sheet, row as u32), &nodes);
            }
        }
    }

    /// Formulas with a range containing `cell`.
    fn readers(&self, cell: CellId, mut f: impl FnMut(CellId)) {
        if let Some(tree) = self.by_col.get(&(cell.sheet, cell.col as u32)) {
            for node in path(cell.row) {
                if let Some(set) = tree.get(&node) {
                    set.iter().copied().for_each(&mut f);
                }
            }
        }
        if let Some(tree) = self.by_row.get(&(cell.sheet, cell.row as u32)) {
            for node in path(cell.col) {
                if let Some(set) = tree.get(&node) {
                    set.iter().copied().for_each(&mut f);
                }
            }
        }
    }

    fn remove_sheet(&mut self, sheet: SheetId) {
        self.by_col.retain(|(s, _), _| *s != sheet);
        self.by_row.retain(|(s, _), _| *s != sheet);
    }
}

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
/// 4. **Atomic updates:** edge replacement and range materialization update both maps.
/// 5. **Formulas inside ranges are edges:** if formula X lies in a range of
///    formula F, then X ∈ preds[F]. Values inside ranges are not edges; they
///    are found through `ranges`.
#[derive(Default, Debug, Clone)]
pub struct DepGraph {
    /// Precedents: for each formula cell B, the cells A it depends on.
    /// B -> {A1, A2, ...}
    preds: FxHashMap<CellId, FxHashSet<CellId>>,

    /// Dependents: for each referenced cell A, the formula cells B that depend on it.
    /// A -> {B1, B2, ...}
    succs: FxHashMap<CellId, FxHashSet<CellId>>,

    /// Each formula's range references, as registered in `ranges`.
    range_refs: FxHashMap<CellId, Vec<RangeRef>>,

    /// Interval index over every formula's ranges.
    ranges: RangeIndex,

    /// Formula cells by sheet and column, to find the formulas inside a
    /// range without walking it.
    formula_rows: FxHashMap<SheetId, FxHashMap<u32, BTreeSet<u32>>>,
}

impl DepGraph {
    /// Create an empty dependency graph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the cells this formula cell depends on for ordering: single
    /// references, and the formulas inside its ranges. Values inside ranges
    /// are not included; see [`DepGraph::precedent_ranges`].
    ///
    /// These are the incoming edges to the cell.
    pub fn precedents(&self, cell: CellId) -> impl Iterator<Item = CellId> + '_ {
        self.preds
            .get(&cell)
            .into_iter()
            .flat_map(|s| s.iter().copied())
    }

    /// The ranges this formula reads.
    pub fn precedent_ranges(&self, cell: CellId) -> &[RangeRef] {
        self.range_refs.get(&cell).map_or(&[], |v| v.as_slice())
    }

    /// Returns the cells that depend on this cell (dependents): formulas
    /// that reference it directly or through a range.
    ///
    /// These are the outgoing edges from the cell.
    pub fn dependents(&self, cell: CellId) -> impl Iterator<Item = CellId> + '_ {
        let concrete = self.succs.get(&cell);
        let mut through_ranges = FxHashSet::default();
        self.ranges.readers(cell, |formula| {
            if !concrete.is_some_and(|s| s.contains(&formula)) {
                through_ranges.insert(formula);
            }
        });
        concrete.into_iter().flat_map(|s| s.iter().copied()).chain(through_ranges)
    }

    /// Number of cells a formula reads: its single references plus its ranges.
    ///
    /// Counts the cells each range covers, as the expanded edges once did,
    /// so the profiler still ranks a wide SUM as the heavy formula it is.
    /// A whole column counts its full height, not its occupied cells.
    pub fn precedent_count(&self, cell: CellId) -> usize {
        let edges = self.preds.get(&cell).map_or(0, |s| s.len());
        self.precedent_ranges(cell).iter().fold(edges, |n, r| {
            n.saturating_add((r.end_row - r.start_row + 1).saturating_mul(r.end_col - r.start_col + 1))
        })
    }

    /// Number of dependents, including range readers.
    pub fn dependent_count(&self, cell: CellId) -> usize {
        self.dependents(cell).count()
    }

    /// Register a formula cell that has no cell references (e.g., `=1/0`, `=PI()`).
    ///
    /// These "leaf" formulas still need to appear in the topo order so that
    /// `recompute_full_ordered` evaluates them after clearing the cache.
    pub fn register_leaf_formula(&mut self, cell: CellId) {
        self.preds.entry(cell).or_default();
        self.index_formula(cell);
    }

    /// Returns true if this cell has formula dependencies tracked in the graph.
    pub fn is_formula_cell(&self, cell: CellId) -> bool {
        self.preds.contains_key(&cell)
    }

    /// Whether `clear_cell` would change anything: the cell has edges or
    /// ranges of its own. Lets a value edit leave a graph shared with a
    /// workbook clone uncopied.
    pub fn has_own_deps(&self, cell: CellId) -> bool {
        self.preds.contains_key(&cell) || self.range_refs.contains_key(&cell)
    }

    /// Returns the number of formula cells (cells with precedents) in the graph.
    pub fn formula_cell_count(&self) -> usize {
        self.preds.len()
    }

    /// Returns the number of cells that are referenced by at least one formula.
    pub fn referenced_cell_count(&self) -> usize {
        self.succs.len()
    }

    fn index_formula(&mut self, cell: CellId) {
        self.formula_rows
            .entry(cell.sheet)
            .or_default()
            .entry(cell.col as u32)
            .or_default()
            .insert(cell.row as u32);
    }

    fn unindex_formula(&mut self, cell: CellId) {
        if let Some(cols) = self.formula_rows.get_mut(&cell.sheet) {
            if let Some(rows) = cols.get_mut(&(cell.col as u32)) {
                rows.remove(&(cell.row as u32));
                if rows.is_empty() {
                    cols.remove(&(cell.col as u32));
                }
            }
            if cols.is_empty() {
                self.formula_rows.remove(&cell.sheet);
            }
        }
    }

    /// Formula cells inside `range`.
    fn formulas_in(&self, range: &RangeRef) -> Vec<CellId> {
        let mut out = Vec::new();
        let Some(cols) = self.formula_rows.get(&range.sheet) else { return out };
        let mut visit = |col: u32, rows: &BTreeSet<u32>| {
            for &row in rows.range(range.start_row as u32..=range.end_row as u32) {
                out.push(CellId::new(range.sheet, row as usize, col as usize));
            }
        };
        if range.width() <= cols.len() {
            for col in range.start_col..=range.end_col {
                if let Some(rows) = cols.get(&(col as u32)) {
                    visit(col as u32, rows);
                }
            }
        } else {
            for (&col, rows) in cols {
                if (range.start_col..=range.end_col).contains(&(col as usize)) {
                    visit(col, rows);
                }
            }
        }
        out
    }

    fn add_edge(&mut self, pred: CellId, formula: CellId) {
        self.preds.entry(formula).or_default().insert(pred);
        self.succs.entry(pred).or_default().insert(formula);
    }

    /// Replace all edges for a formula cell atomically.
    ///
    /// This is the primary mutation API. It:
    /// 1. Removes the cell from all its old precedents' successor sets
    /// 2. Drops its range subscriptions
    /// 3. Adds the cell to all new precedents' successor sets
    /// 4. Sets the cell's new precedent set
    ///
    /// Pass an empty set to clear all edges for this cell. Ranges are set
    /// separately, with [`DepGraph::set_ranges`], after this.
    pub fn replace_edges(&mut self, formula_cell: CellId, new_preds: FxHashSet<CellId>) {
        if let Some(old) = self.range_refs.remove(&formula_cell) {
            for range in &old {
                self.ranges.remove(formula_cell, range);
            }
        }
        // Step 1: Remove old edges
        let was_formula = self.preds.contains_key(&formula_cell);
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
            if was_formula {
                self.unindex_formula(formula_cell);
            }
            return;
        }

        // Step 3: Add new edges
        for pred in &new_preds {
            self.succs.entry(*pred).or_default().insert(formula_cell);
        }

        // Step 4: Store new precedents
        self.preds.insert(formula_cell, new_preds);
        self.index_formula(formula_cell);
    }

    /// Register a formula's range references, after `replace_edges`: index
    /// them, and add an ordering edge from every formula already inside
    /// them. Formulas that arrive inside the range later are linked by
    /// [`DepGraph::track_range_cell`].
    pub fn set_ranges(&mut self, formula: CellId, ranges: Vec<RangeRef>) {
        if ranges.is_empty() {
            return;
        }
        debug_assert!(
            ranges.iter().all(|r| r.start_row <= r.end_row && r.start_col <= r.end_col),
            "ranges reach the graph normalized: {ranges:?}"
        );
        // A formula with ranges is a formula, even with no single references;
        // registering it first also catches a range that contains itself.
        self.preds.entry(formula).or_default();
        self.index_formula(formula);
        for range in &ranges {
            self.ranges.insert(formula, range);
            for inner in self.formulas_in(range) {
                self.add_edge(inner, formula);
            }
        }
        self.range_refs.entry(formula).or_default().extend(ranges);
    }

    /// Link a formula cell to every formula whose ranges contain it, so the
    /// readers are ordered after it. Call when a cell becomes a formula (and
    /// for new spill receivers, which is harmless: only formulas get edges).
    pub fn track_range_cell(&mut self, cell: CellId) {
        if !self.is_formula_cell(cell) {
            return;
        }
        let mut readers = Vec::new();
        self.ranges.readers(cell, |f| readers.push(f));
        for reader in readers {
            self.add_edge(cell, reader);
        }
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
    pub fn remove_sheet(&mut self, sheet: SheetId) {
        // Formulas elsewhere keep their range lists; the entries on this
        // sheet's index go with it.
        for ranges in self.range_refs.values_mut() {
            ranges.retain(|r| r.sheet != sheet);
        }
        self.range_refs.retain(|_, r| !r.is_empty());
        self.ranges.remove_sheet(sheet);
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
        self.formula_rows.remove(&sheet);

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
                        if preds.is_empty() && !self.range_refs.contains_key(&dep) {
                            self.preds.remove(&dep);
                            self.unindex_formula(dep);
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
    /// This rebuilds the graph with remapped coordinates. A range keeps its
    /// surviving extent: its first and last surviving rows and columns.
    pub fn apply_mapping<F>(&mut self, map: F)
    where
        F: Fn(CellId) -> Option<CellId>,
    {
        let old_preds = std::mem::take(&mut self.preds);
        let old_ranges = std::mem::take(&mut self.range_refs);
        *self = DepGraph::default();

        for (formula_cell, preds) in &old_preds {
            // Map the formula cell
            let Some(new_formula_cell) = map(*formula_cell) else {
                continue; // Formula cell was deleted
            };
            let ranges = old_ranges.get(formula_cell);
            // Map all precedents, keeping only those that survive. Formulas
            // inside ranges come back through set_ranges below.
            let mapped_preds: FxHashSet<CellId> = preds
                .iter()
                .filter(|p| !ranges.is_some_and(|rs| rs.iter().any(|r| r.contains(**p))))
                .filter_map(|p| map(*p))
                .collect();

            if mapped_preds.is_empty() && ranges.is_none() {
                continue; // All precedents were deleted
            }
            if mapped_preds.is_empty() {
                self.register_leaf_formula(new_formula_cell);
            } else {
                self.replace_edges(new_formula_cell, mapped_preds);
            }
        }

        for (formula, ranges) in old_ranges {
            let Some(formula) = map(formula) else { continue };
            let mapped: Vec<RangeRef> = ranges.iter().filter_map(|r| map_range(r, &map)).collect();
            self.set_ranges(formula, mapped);
        }
        let formulas: Vec<CellId> = self.preds.keys().copied().collect();
        for formula in formulas {
            self.track_range_cell(formula);
        }
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

                    if let std::collections::hash_map::Entry::Vacant(e) = indices.entry(w) {
                        // Recurse into w
                        let w_idx = index_counter;
                        index_counter += 1;
                        e.insert(w_idx);
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
                            if self.preds.get(&cell).is_some_and(|p| p.contains(&cell)) {
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

                    if let std::collections::hash_map::Entry::Vacant(e) = indices.entry(w) {
                        let w_idx = index_counter;
                        index_counter += 1;
                        e.insert(w_idx);
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
                            self.preds.get(&cell).is_some_and(|p| p.contains(&cell))
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

    /// Topologically order a subset of the formula cells.
    ///
    /// Same algorithm and the same deterministic tie-breaking as
    /// [`Self::topo_order_all_formulas`], but the work is proportional to
    /// `cells` and the edges incident to it rather than to the whole workbook.
    ///
    /// This exists for incremental recalculation. The caller's dirty set is a
    /// forward closure — every transitive dependent of the cells that changed —
    /// so it is closed under `dependents`, and a cell inside it can only be
    /// ordered after the cells inside it that it reads. Precedents *outside*
    /// the set are clean: their cached values are still correct, so they carry
    /// no edge here and contribute nothing to the ordering.
    ///
    /// Ordering the whole workbook to place one changed cell is what made a
    /// single-cell edit cost 78 ms on 200,000 formulas — linear in the
    /// workbook, with a dirty set of one.
    ///
    /// Cells in `cells` that are not formula cells are ignored, matching
    /// `topo_order_all_formulas`, whose domain is `preds.keys()`.
    ///
    /// Returns `Err` only when the cycle is *within* `cells`. A cycle
    /// elsewhere in the workbook cannot affect this ordering and is not this
    /// function's business.
    pub fn topo_order_subset(
        &self,
        cells: &FxHashSet<CellId>,
    ) -> Result<Vec<CellId>, CycleReport> {
        let members: FxHashSet<CellId> =
            cells.iter().copied().filter(|c| self.preds.contains_key(c)).collect();

        if members.is_empty() {
            return Ok(Vec::new());
        }

        // In-degree counts only precedents that are themselves in the subset.
        let mut in_degree: FxHashMap<CellId, usize> = FxHashMap::default();
        for &cell in &members {
            let count = self
                .preds
                .get(&cell)
                .map(|preds| preds.iter().filter(|p| members.contains(p)).count())
                .unwrap_or(0);
            in_degree.insert(cell, count);
        }

        let mut queue: Vec<CellId> = in_degree
            .iter()
            .filter(|(_, &deg)| deg == 0)
            .map(|(&cell, _)| cell)
            .collect();
        // Descending, so the smallest is at the end and pops first.
        queue.sort_by(|a, b| {
            b.sheet
                .raw()
                .cmp(&a.sheet.raw())
                .then(b.row.cmp(&a.row))
                .then(b.col.cmp(&a.col))
        });

        let mut result = Vec::with_capacity(members.len());

        while let Some(cell) = queue.pop() {
            result.push(cell);

            if let Some(deps) = self.succs.get(&cell) {
                let mut new_zero_degree = Vec::new();

                for &dep in deps {
                    if members.contains(&dep) {
                        if let Some(deg) = in_degree.get_mut(&dep) {
                            *deg = deg.saturating_sub(1);
                            if *deg == 0 {
                                new_zero_degree.push(dep);
                            }
                        }
                    }
                }

                new_zero_degree.sort_by(|a, b| {
                    a.sheet
                        .raw()
                        .cmp(&b.sheet.raw())
                        .then(a.row.cmp(&b.row))
                        .then(a.col.cmp(&b.col))
                });
                for cell in new_zero_degree.into_iter().rev() {
                    queue.push(cell);
                }
            }
        }

        if result.len() < members.len() {
            let placed: FxHashSet<CellId> = result.iter().copied().collect();
            let cycle_cells: Vec<CellId> =
                members.iter().filter(|c| !placed.contains(c)).copied().collect();
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

            for dep in self.dependents(current) {
                if new_preds_set.contains(&dep) {
                    return Some(CycleReport::cycle(vec![dep, cell]));
                }
                stack.push(dep);
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
                    self.succs.get(pred).is_some_and(|s| s.contains(formula_cell)),
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
                    self.preds.get(dep).is_some_and(|s| s.contains(cell)),
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

    // ---- #29: ranges as ranges ----

    fn rect(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> RangeRef {
        RangeRef { sheet: SheetId(1), start_row, start_col, end_row, end_col }
    }

    fn to_set(it: impl Iterator<Item = CellId>) -> FxHashSet<CellId> {
        it.collect()
    }

    #[test]
    fn cover_is_exact_and_disjoint() {
        for (lo, hi) in [(0, 0), (0, 5), (3, 17), (1000, 1_048_575), (0, 1_048_575), (511, 512)] {
            let mut nodes = Vec::new();
            cover(lo, hi, &mut nodes);
            // Every position in [lo, hi] is covered exactly once, and nothing outside.
            // Probes stay inside the tree's domain, [0, 2^20): rows and
            // columns never reach it.
            for p in [lo, hi, (lo + hi) / 2, lo.saturating_sub(1), hi + 1].into_iter().filter(|p| *p < 1 << TREE_BITS) {
                let hits = path(p).filter(|n| nodes.contains(n)).count();
                let inside = (lo..=hi).contains(&p);
                assert_eq!(hits, usize::from(inside), "position {p} in [{lo}, {hi}]");
            }
            assert!(nodes.len() <= 2 * (TREE_BITS as usize + 1), "{} nodes", nodes.len());
        }
    }

    #[test]
    fn a_running_total_has_no_edges_to_values() {
        let mut g = DepGraph::new();
        for r in 0..1000 {
            let f = cell(1, r, 1);
            g.register_leaf_formula(f);
            g.set_ranges(f, vec![rect(0, 0, r, 0)]);
            g.track_range_cell(f);
        }
        assert_eq!(g.referenced_cell_count(), 0, "no per-cell edges for values");
        // ...yet every reader of A1 is found, and only readers of A500 for A500.
        assert_eq!(g.dependents(cell(1, 0, 0)).count(), 1000);
        assert_eq!(g.dependents(cell(1, 499, 0)).count(), 501);
    }

    #[test]
    fn formulas_inside_a_range_are_ordering_edges_both_ways_round() {
        // Reader registered first, formula inside added later.
        let mut g = DepGraph::new();
        let reader = cell(1, 10, 1);
        g.register_leaf_formula(reader);
        g.set_ranges(reader, vec![rect(0, 0, 9, 0)]);
        let inner = cell(1, 4, 0);
        g.register_leaf_formula(inner);
        g.track_range_cell(inner);
        assert!(to_set(g.precedents(reader)).contains(&inner));
        assert_eq!(g.topo_order_all_formulas().unwrap().iter().position(|c| *c == inner).unwrap() <
            g.topo_order_all_formulas().unwrap().iter().position(|c| *c == reader).unwrap(), true);

        // Formula inside registered first, reader later.
        let mut g = DepGraph::new();
        g.register_leaf_formula(inner);
        g.register_leaf_formula(reader);
        g.set_ranges(reader, vec![rect(0, 0, 9, 0)]);
        assert!(to_set(g.precedents(reader)).contains(&inner));
    }

    #[test]
    fn a_range_containing_its_own_formula_is_a_cycle() {
        let mut g = DepGraph::new();
        let f = cell(1, 5, 0);
        g.set_ranges(f, vec![rect(0, 0, 9, 0)]);
        assert!(g.find_cycle_members().contains(&f));
        assert!(g.topo_order_all_formulas().is_err());
    }

    #[test]
    fn wide_ranges_live_in_row_trees_and_are_found() {
        let mut g = DepGraph::new();
        let f = cell(1, 50, 0);
        // A whole row: 1 tall, 16,384 wide, so it is stored per row.
        let whole_row = rect(2, 0, 2, crate::sheet::NUM_COLS - 1);
        g.register_leaf_formula(f);
        g.set_ranges(f, vec![whole_row]);
        assert!(g.ranges.by_col.is_empty(), "not stored in 16,384 column trees");
        assert_eq!(to_set(g.dependents(cell(1, 2, 9_000))), to_set([f].into_iter()));
        assert_eq!(g.dependents(cell(1, 3, 0)).count(), 0);
    }

    #[test]
    fn re_editing_a_formula_drops_its_old_ranges() {
        let mut g = DepGraph::new();
        let f = cell(1, 20, 2);
        g.register_leaf_formula(f);
        g.set_ranges(f, vec![rect(0, 0, 9, 0)]);
        // Re-edit: replace_edges clears ranges; new ones are set after.
        g.replace_edges(f, FxHashSet::default());
        g.register_leaf_formula(f);
        g.set_ranges(f, vec![rect(0, 1, 9, 1)]);
        assert_eq!(g.dependents(cell(1, 5, 0)).count(), 0, "old range forgotten");
        assert_eq!(g.dependents(cell(1, 5, 1)).count(), 1);
        g.clear_cell(f);
        assert_eq!(g.dependents(cell(1, 5, 1)).count(), 0, "cleared formula reads nothing");
        assert!(g.ranges.by_col.is_empty() && g.ranges.by_row.is_empty(), "index emptied");
        assert!(g.formula_rows.is_empty(), "formula index emptied");
    }

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
    fn topo_order_subset_ignores_precedents_outside_the_set() {
        // C depends on B depends on A. Only B and C are dirty; A is clean and
        // keeps its cached value, so it must not appear in the order and must
        // not hold B back — B has no in-subset precedent and goes first.
        let mut graph = DepGraph::new();
        let sheet = crate::sheet::SheetId(0);
        let a = CellId::new(sheet, 0, 0);
        let b = CellId::new(sheet, 1, 0);
        let c = CellId::new(sheet, 2, 0);
        graph.replace_edges(b, [a].into_iter().collect());
        graph.replace_edges(c, [b].into_iter().collect());

        let dirty: FxHashSet<CellId> = [b, c].into_iter().collect();
        let order = graph.topo_order_subset(&dirty).unwrap();

        assert_eq!(order, vec![b, c]);
    }

    #[test]
    fn topo_order_subset_reports_only_a_cycle_inside_the_set() {
        // X <-> Y is a live cycle; Z is an unrelated acyclic formula. Ordering
        // Z alone must succeed — a cycle elsewhere in the workbook cannot
        // affect it, and treating it as unorderable is what used to drag a
        // full recompute behind every unrelated edit.
        let mut graph = DepGraph::new();
        let sheet = crate::sheet::SheetId(0);
        let x = CellId::new(sheet, 0, 0);
        let y = CellId::new(sheet, 1, 0);
        let src = CellId::new(sheet, 5, 0);
        let z = CellId::new(sheet, 2, 0);
        graph.replace_edges(x, [y].into_iter().collect());
        graph.replace_edges(y, [x].into_iter().collect());
        graph.replace_edges(z, [src].into_iter().collect());

        assert!(graph.topo_order_all_formulas().is_err(), "the workbook does contain a cycle");

        let unrelated: FxHashSet<CellId> = [z].into_iter().collect();
        assert_eq!(graph.topo_order_subset(&unrelated).unwrap(), vec![z]);

        let inside: FxHashSet<CellId> = [x, y].into_iter().collect();
        assert!(graph.topo_order_subset(&inside).is_err());
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
                    !graph.preds.get(&cell_i).is_some_and(|p| p.contains(&cell_j)),
                    "{:?} at position {} depends on {:?} at position {}",
                    cell_i, i, cell_j, j
                );
            }
        }
    }
}

/// A range through a structural mapping: its first and last surviving rows
/// and columns. `None` when none survive.
fn map_range<F>(range: &RangeRef, map: &F) -> Option<RangeRef>
where
    F: Fn(CellId) -> Option<CellId>,
{
    let row_of = |r: usize| (range.start_col..=range.end_col).find_map(|c| map(CellId::new(range.sheet, r, c))).map(|c| c.row);
    let col_of = |c: usize| (range.start_row..=range.end_row).find_map(|r| map(CellId::new(range.sheet, r, c))).map(|c| c.col);
    // Probe one line per axis rather than every cell: structural mappings
    // move whole rows and columns.
    let row_probe = |r: usize| map(CellId::new(range.sheet, r, range.start_col)).or_else(|| map(CellId::new(range.sheet, r, range.end_col))).map(|c| c.row).or_else(|| row_of(r));
    let col_probe = |c: usize| map(CellId::new(range.sheet, range.start_row, c)).or_else(|| map(CellId::new(range.sheet, range.end_row, c))).map(|c| c.col).or_else(|| col_of(c));
    let start_row = (range.start_row..=range.end_row).find_map(&row_probe)?;
    let end_row = (range.start_row..=range.end_row).rev().find_map(&row_probe)?;
    let start_col = (range.start_col..=range.end_col).find_map(&col_probe)?;
    let end_col = (range.start_col..=range.end_col).rev().find_map(&col_probe)?;
    let sheet = map(CellId::new(range.sheet, range.start_row, range.start_col)).map_or(range.sheet, |c| c.sheet);
    Some(RangeRef { sheet, start_row, start_col, end_row, end_col })
}
