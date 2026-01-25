use rustc_hash::{FxHashMap, FxHashSet};
use serde::{Deserialize, Serialize};
use crate::cell::CellValue;
use crate::cell_id::CellId;
use crate::dep_graph::DepGraph;
use crate::sheet::{Sheet, SheetId, normalize_sheet_name, is_valid_sheet_name};
use crate::named_range::{NamedRange, NamedRangeStore};
use crate::formula::eval::{CellLookup, NamedRangeResolution, Value};
use crate::formula::parser::bind_expr;
use crate::formula::refs::extract_cell_ids;

/// Impact analysis for a cell change (Phase 3.5a).
///
/// Describes what would happen if the cell's value changed.
#[derive(Debug, Clone, Default)]
pub struct ImpactInfo {
    /// Number of cells that would be affected (transitive dependents).
    pub affected_cells: usize,
    /// Maximum depth in the dependency chain.
    pub max_depth: usize,
    /// True if any cell in the chain has unknown/dynamic dependencies.
    pub has_unknown_in_chain: bool,
    /// True if impact cannot be bounded (due to dynamic refs).
    pub is_unbounded: bool,
}

/// Result of a path trace operation (Phase 3.5b).
#[derive(Debug, Clone, Default)]
pub struct PathTraceResult {
    /// The path from source to target (inclusive).
    /// Empty if no path exists.
    pub path: Vec<CellId>,
    /// True if any cell in the path has dynamic refs (INDIRECT/OFFSET).
    pub has_dynamic_refs: bool,
    /// True if the search was truncated due to caps.
    pub truncated: bool,
}

/// A workbook containing multiple sheets
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    sheets: Vec<Sheet>,
    active_sheet: usize,
    /// Next ID to assign to a new sheet. Monotonically increasing, never reused.
    #[serde(default = "default_next_sheet_id")]
    next_sheet_id: u64,
    #[serde(default)]
    named_ranges: NamedRangeStore,

    /// Dependency graph for formula cells.
    /// Rebuilt on load, updated incrementally on cell changes.
    #[serde(skip)]
    dep_graph: DepGraph,
}

fn default_next_sheet_id() -> u64 {
    1
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    /// Create a new workbook with one default sheet
    pub fn new() -> Self {
        let sheet = Sheet::new(SheetId(1), 65536, 256);
        Self {
            sheets: vec![sheet],
            active_sheet: 0,
            next_sheet_id: 2, // Next ID will be 2
            named_ranges: NamedRangeStore::new(),
            dep_graph: DepGraph::new(),
        }
    }

    /// Generate a new unique SheetId (monotonically increasing, never reused)
    fn generate_sheet_id(&mut self) -> SheetId {
        let id = SheetId(self.next_sheet_id);
        self.next_sheet_id += 1;
        id
    }

    /// Get the next_sheet_id value (for persistence)
    pub fn next_sheet_id(&self) -> u64 {
        self.next_sheet_id
    }

    /// Set the next_sheet_id value (for loading from persistence)
    pub fn set_next_sheet_id(&mut self, id: u64) {
        self.next_sheet_id = id;
    }

    /// Get the number of sheets
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Get the active sheet index
    pub fn active_sheet_index(&self) -> usize {
        self.active_sheet
    }

    /// Set the active sheet by index
    pub fn set_active_sheet(&mut self, index: usize) -> bool {
        if index < self.sheets.len() {
            self.active_sheet = index;
            true
        } else {
            false
        }
    }

    /// Get a reference to the active sheet
    pub fn active_sheet(&self) -> &Sheet {
        &self.sheets[self.active_sheet]
    }

    /// Get a mutable reference to the active sheet
    pub fn active_sheet_mut(&mut self) -> &mut Sheet {
        &mut self.sheets[self.active_sheet]
    }

    /// Get a reference to a sheet by index
    pub fn sheet(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    /// Get a mutable reference to a sheet by index
    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Sheet> {
        self.sheets.get_mut(index)
    }

    /// Get all sheet names
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    /// Add a new sheet and return its index
    pub fn add_sheet(&mut self) -> usize {
        let sheet_num = self.sheets.len() + 1;
        let mut new_name = format!("Sheet{}", sheet_num);

        // Ensure unique name (case-insensitive)
        while self.sheet_name_exists(&new_name) {
            let num: usize = new_name.strip_prefix("Sheet")
                .and_then(|n| n.parse().ok())
                .unwrap_or(sheet_num);
            new_name = format!("Sheet{}", num + 1);
        }

        let id = self.generate_sheet_id();
        let sheet = Sheet::new_with_name(id, 65536, 256, &new_name);
        self.sheets.push(sheet);
        self.sheets.len() - 1
    }

    /// Add a new sheet with a specific name
    /// Returns None if name is invalid or already exists
    pub fn add_sheet_named(&mut self, name: &str) -> Option<usize> {
        if !is_valid_sheet_name(name) {
            return None;
        }
        if self.sheet_name_exists(name) {
            return None;
        }
        let id = self.generate_sheet_id();
        let sheet = Sheet::new_with_name(id, 65536, 256, name);
        self.sheets.push(sheet);
        Some(self.sheets.len() - 1)
    }

    /// Check if a sheet name already exists (case-insensitive)
    pub fn sheet_name_exists(&self, name: &str) -> bool {
        let key = normalize_sheet_name(name);
        self.sheets.iter().any(|s| s.name_key == key)
    }

    /// Check if a sheet name is available for a given sheet (for rename)
    /// Returns true if the name is not used by any other sheet
    pub fn is_name_available(&self, name: &str, exclude_id: SheetId) -> bool {
        let key = normalize_sheet_name(name);
        !self.sheets.iter().any(|s| s.id != exclude_id && s.name_key == key)
    }

    /// Delete a sheet by index
    /// Returns false if it's the last sheet (can't delete)
    pub fn delete_sheet(&mut self, index: usize) -> bool {
        if self.sheets.len() <= 1 || index >= self.sheets.len() {
            return false;
        }

        self.sheets.remove(index);

        // Adjust active sheet if needed
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        } else if self.active_sheet > index {
            self.active_sheet -= 1;
        }

        true
    }

    /// Rename a sheet by index
    /// Returns false if:
    /// - Index is invalid
    /// - Name is invalid (empty after trim)
    /// - Name is already used by another sheet (case-insensitive)
    pub fn rename_sheet(&mut self, index: usize, new_name: &str) -> bool {
        if !is_valid_sheet_name(new_name) {
            return false;
        }
        if let Some(sheet) = self.sheets.get(index) {
            let sheet_id = sheet.id;
            if !self.is_name_available(new_name, sheet_id) {
                return false;
            }
            // Now safe to mutate
            if let Some(sheet) = self.sheets.get_mut(index) {
                sheet.set_name(new_name);
                return true;
            }
        }
        false
    }

    // =========================================================================
    // Sheet ID-based Access
    // =========================================================================

    /// Get a sheet's index by its ID
    pub fn idx_for_sheet_id(&self, id: SheetId) -> Option<usize> {
        self.sheets.iter().position(|s| s.id == id)
    }

    /// Get the SheetId at a given index
    pub fn sheet_id_at_idx(&self, idx: usize) -> Option<SheetId> {
        self.sheets.get(idx).map(|s| s.id)
    }

    /// Get a reference to a sheet by its ID
    pub fn sheet_by_id(&self, id: SheetId) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.id == id)
    }

    /// Get a mutable reference to a sheet by its ID
    pub fn sheet_by_id_mut(&mut self, id: SheetId) -> Option<&mut Sheet> {
        self.sheets.iter_mut().find(|s| s.id == id)
    }

    /// Find a sheet by name (case-insensitive)
    pub fn sheet_by_name(&self, name: &str) -> Option<&Sheet> {
        let key = normalize_sheet_name(name);
        self.sheets.iter().find(|s| s.name_key == key)
    }

    /// Get the SheetId for a sheet by name (case-insensitive)
    pub fn sheet_id_by_name(&self, name: &str) -> Option<SheetId> {
        self.sheet_by_name(name).map(|s| s.id)
    }

    /// Get the active sheet's ID
    pub fn active_sheet_id(&self) -> SheetId {
        self.sheets[self.active_sheet].id
    }

    /// Move to the next sheet
    pub fn next_sheet(&mut self) -> bool {
        if self.active_sheet + 1 < self.sheets.len() {
            self.active_sheet += 1;
            true
        } else {
            false
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self) -> bool {
        if self.active_sheet > 0 {
            self.active_sheet -= 1;
            true
        } else {
            false
        }
    }

    /// Get all sheets (for serialization)
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Create a workbook from sheets (for deserialization)
    /// Note: next_sheet_id should be set separately via set_next_sheet_id if loading from file
    /// Call `rebuild_dep_graph()` after loading to populate the dependency graph.
    pub fn from_sheets(sheets: Vec<Sheet>, active: usize) -> Self {
        let active_sheet = active.min(sheets.len().saturating_sub(1));
        // Calculate next_sheet_id as max existing id + 1
        let max_id = sheets.iter().map(|s| s.id.raw()).max().unwrap_or(0);
        Self {
            sheets,
            active_sheet,
            next_sheet_id: max_id + 1,
            named_ranges: NamedRangeStore::new(),
            dep_graph: DepGraph::new(),
        }
    }

    /// Create a workbook from sheets with explicit next_sheet_id (for full deserialization)
    /// Call `rebuild_dep_graph()` after loading to populate the dependency graph.
    pub fn from_sheets_with_meta(sheets: Vec<Sheet>, active: usize, next_sheet_id: u64) -> Self {
        let active_sheet = active.min(sheets.len().saturating_sub(1));
        Self {
            sheets,
            active_sheet,
            next_sheet_id,
            named_ranges: NamedRangeStore::new(),
            dep_graph: DepGraph::new(),
        }
    }

    // =========================================================================
    // Named Range Management
    // =========================================================================

    /// Get a reference to the named range store
    pub fn named_ranges(&self) -> &NamedRangeStore {
        &self.named_ranges
    }

    /// Get a mutable reference to the named range store
    pub fn named_ranges_mut(&mut self) -> &mut NamedRangeStore {
        &mut self.named_ranges
    }

    /// Define a named range for a single cell (convenience method)
    pub fn define_name_for_cell(
        &mut self,
        name: &str,
        sheet: usize,
        row: usize,
        col: usize,
    ) -> Result<(), String> {
        let range = NamedRange::cell(name, sheet, row, col);
        self.named_ranges.set(range)
    }

    /// Define a named range for a cell range (convenience method)
    pub fn define_name_for_range(
        &mut self,
        name: &str,
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<(), String> {
        let range = NamedRange::range(name, sheet, start_row, start_col, end_row, end_col);
        self.named_ranges.set(range)
    }

    /// Get a named range by name (case-insensitive)
    pub fn get_named_range(&self, name: &str) -> Option<&NamedRange> {
        self.named_ranges.get(name)
    }

    /// Rename a named range
    pub fn rename_named_range(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        self.named_ranges.rename(old_name, new_name)
    }

    /// Delete a named range
    pub fn delete_named_range(&mut self, name: &str) -> bool {
        self.named_ranges.remove(name).is_some()
    }

    /// Find all named ranges that reference a specific cell
    pub fn named_ranges_for_cell(&self, sheet: usize, row: usize, col: usize) -> Vec<&NamedRange> {
        self.named_ranges.find_by_cell(sheet, row, col)
    }

    /// List all named ranges
    pub fn list_named_ranges(&self) -> Vec<&NamedRange> {
        self.named_ranges.list()
    }

    // =========================================================================
    // Dependency Graph
    // =========================================================================

    /// Get a reference to the dependency graph.
    pub fn dep_graph(&self) -> &DepGraph {
        &self.dep_graph
    }

    /// Rebuild the dependency graph from scratch.
    ///
    /// Call this after loading a workbook to populate the graph.
    /// Iterates all formula cells and extracts their references.
    pub fn rebuild_dep_graph(&mut self) {
        self.dep_graph = DepGraph::new();

        // Iterate all sheets and cells
        for sheet in &self.sheets {
            let sheet_id = sheet.id;

            for ((row, col), cell) in sheet.cells_iter() {
                if let CellValue::Formula { ast: Some(ast), .. } = &cell.value {
                    // Bind the AST with cross-sheet resolution
                    let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));

                    // Extract cell references
                    let refs = extract_cell_ids(
                        &bound,
                        sheet_id,
                        &self.named_ranges,
                        |idx| self.sheet_id_at_idx(idx),
                    );

                    if !refs.is_empty() {
                        let formula_cell = CellId::new(sheet_id, *row, *col);
                        let preds: FxHashSet<CellId> = refs.into_iter().collect();
                        self.dep_graph.replace_edges(formula_cell, preds);
                    }
                }
            }
        }
    }

    /// Update the dependency graph for a specific cell.
    ///
    /// Call this after setting a cell value (formula or otherwise).
    /// If the cell has a formula, extracts and updates its dependencies.
    /// If the cell is not a formula (or empty), clears any existing dependencies.
    pub fn update_cell_deps(&mut self, sheet_id: SheetId, row: usize, col: usize) {
        let cell_id = CellId::new(sheet_id, row, col);

        // Get the cell's formula AST if it exists (clone to avoid borrow issues)
        let ast = self.sheet_by_id(sheet_id)
            .and_then(|sheet| sheet.get_cell(row, col).value.formula_ast().cloned());

        if let Some(ast) = ast {
            // Bind and extract references
            let bound = bind_expr(&ast, |name| self.sheet_id_by_name(name));
            let refs = extract_cell_ids(
                &bound,
                sheet_id,
                &self.named_ranges,
                |idx| self.sheet_id_at_idx(idx),
            );

            let preds: FxHashSet<CellId> = refs.into_iter().collect();
            self.dep_graph.replace_edges(cell_id, preds);
        } else {
            // Not a formula, clear any existing edges
            self.dep_graph.clear_cell(cell_id);
        }
    }

    /// Clear dependencies for a cell (e.g., when the cell is deleted or cleared).
    pub fn clear_cell_deps(&mut self, sheet_id: SheetId, row: usize, col: usize) {
        let cell_id = CellId::new(sheet_id, row, col);
        self.dep_graph.clear_cell(cell_id);
    }

    /// Get the precedents (cells this formula depends on) for a cell.
    pub fn get_precedents(&self, sheet_id: SheetId, row: usize, col: usize) -> Vec<CellId> {
        let cell_id = CellId::new(sheet_id, row, col);
        self.dep_graph.precedents(cell_id).collect()
    }

    /// Get the dependents (cells that depend on this cell) for a cell.
    pub fn get_dependents(&self, sheet_id: SheetId, row: usize, col: usize) -> Vec<CellId> {
        let cell_id = CellId::new(sheet_id, row, col);
        self.dep_graph.dependents(cell_id).collect()
    }

    // =========================================================================
    // Impact Analysis (Phase 3.5)
    // =========================================================================

    /// Compute the impact of changing a cell.
    ///
    /// Returns the blast radius (number of cells affected), maximum depth in
    /// the dependency chain, and whether any cell in the chain has unknown deps.
    pub fn compute_impact(&self, sheet_id: SheetId, row: usize, col: usize) -> ImpactInfo {
        use crate::formula::analyze::has_dynamic_deps;

        let cell_id = CellId::new(sheet_id, row, col);
        let mut visited = FxHashSet::default();
        let mut queue = vec![cell_id];
        let mut has_unknown_in_chain = false;

        // Check if the source cell itself has unknown deps
        if let Some(sheet) = self.sheet_by_id(sheet_id) {
            if let Some(cell) = sheet.cells.get(&(row, col)) {
                if let Some(ast) = cell.value.formula_ast() {
                    if has_dynamic_deps(ast) {
                        has_unknown_in_chain = true;
                    }
                }
            }
        }

        // BFS to find all transitive dependents
        let mut depth = 0;
        while !queue.is_empty() {
            let level_size = queue.len();
            for _ in 0..level_size {
                let current = queue.remove(0);
                if visited.contains(&current) {
                    continue;
                }
                visited.insert(current);

                // Check for unknown deps in this cell
                if !has_unknown_in_chain {
                    if let Some(sheet) = self.sheet_by_id(current.sheet) {
                        if let Some(cell) = sheet.cells.get(&(current.row, current.col)) {
                            if let Some(ast) = cell.value.formula_ast() {
                                if has_dynamic_deps(ast) {
                                    has_unknown_in_chain = true;
                                }
                            }
                        }
                    }
                }

                // Add dependents to queue
                for dep in self.dep_graph.dependents(current) {
                    if !visited.contains(&dep) {
                        queue.push(dep);
                    }
                }
            }
            if !queue.is_empty() {
                depth += 1;
            }
        }

        // Subtract 1 from visited count (don't count the source cell itself)
        let affected_count = visited.len().saturating_sub(1);

        ImpactInfo {
            affected_cells: affected_count,
            max_depth: depth,
            has_unknown_in_chain,
            is_unbounded: has_unknown_in_chain,
        }
    }

    /// Check if any upstream cell (precedent chain) is in a cycle.
    ///
    /// Returns true if this cell's value cannot be trusted because it depends
    /// on a circular reference somewhere in its precedent chain.
    pub fn has_cycle_in_upstream(&self, sheet_id: SheetId, row: usize, col: usize) -> bool {
        let cell_id = CellId::new(sheet_id, row, col);
        let mut visited = FxHashSet::default();
        let mut queue = vec![cell_id];

        // BFS through precedents
        while let Some(current) = queue.pop() {
            if visited.contains(&current) {
                continue;
            }
            visited.insert(current);

            // Check if this cell has a cycle error
            if let Some(sheet) = self.sheet_by_id(current.sheet) {
                if let Some(cell) = sheet.cells.get(&(current.row, current.col)) {
                    if cell.value.is_cycle_error() {
                        return true;
                    }
                }
            }

            // Add precedents to queue
            for prec in self.dep_graph.precedents(current) {
                if !visited.contains(&prec) {
                    queue.push(prec);
                }
            }
        }

        false
    }

    // =========================================================================
    // Path Trace (Phase 3.5b)
    // =========================================================================

    /// Find the shortest path between two cells in the dependency graph.
    ///
    /// Direction determines traversal:
    /// - `forward=true`: traverse dependents (from input toward outputs)
    /// - `forward=false`: traverse precedents (from output toward inputs)
    ///
    /// Returns the shortest path with deterministic ordering (ties broken by
    /// SheetId, row, col). Includes caps to prevent UI stalls.
    pub fn find_path(&self, from: CellId, to: CellId, forward: bool) -> PathTraceResult {
        use crate::formula::analyze::has_dynamic_deps;

        const MAX_VISITED: usize = 50_000;
        const MAX_DEPTH: usize = 500;

        let mut visited: FxHashSet<CellId> = FxHashSet::default();
        let mut prev: FxHashMap<CellId, CellId> = FxHashMap::default();
        let mut queue = std::collections::VecDeque::new();
        let mut has_dynamic_refs = false;
        let mut truncated = false;

        // Check if source has dynamic refs
        if let Some(sheet) = self.sheet_by_id(from.sheet) {
            if let Some(cell) = sheet.cells.get(&(from.row, from.col)) {
                if let Some(ast) = cell.value.formula_ast() {
                    if has_dynamic_deps(ast) {
                        has_dynamic_refs = true;
                    }
                }
            }
        }

        queue.push_back((from, 0usize));
        visited.insert(from);

        while let Some((current, current_depth)) = queue.pop_front() {
            // Check caps
            if visited.len() > MAX_VISITED || current_depth > MAX_DEPTH {
                truncated = true;
                break;
            }

            // Found target?
            if current == to {
                // Reconstruct path
                let mut path = vec![to];
                let mut node = to;
                while let Some(&p) = prev.get(&node) {
                    path.push(p);
                    node = p;
                }
                path.reverse();
                return PathTraceResult {
                    path,
                    has_dynamic_refs,
                    truncated: false,
                };
            }

            // Get neighbors based on direction
            let neighbors: Vec<CellId> = if forward {
                self.dep_graph.dependents(current).collect()
            } else {
                self.dep_graph.precedents(current).collect()
            };

            // Sort for determinism: (sheet, row, col)
            let mut sorted_neighbors = neighbors;
            sorted_neighbors.sort_by(|a, b| {
                (a.sheet.0, a.row, a.col).cmp(&(b.sheet.0, b.row, b.col))
            });

            for neighbor in sorted_neighbors {
                if visited.contains(&neighbor) {
                    continue;
                }

                visited.insert(neighbor);
                prev.insert(neighbor, current);

                // Check for dynamic refs
                if !has_dynamic_refs {
                    if let Some(sheet) = self.sheet_by_id(neighbor.sheet) {
                        if let Some(cell) = sheet.cells.get(&(neighbor.row, neighbor.col)) {
                            if let Some(ast) = cell.value.formula_ast() {
                                if has_dynamic_deps(ast) {
                                    has_dynamic_refs = true;
                                }
                            }
                        }
                    }
                }

                queue.push_back((neighbor, current_depth + 1));
            }
        }

        // No path found or truncated
        PathTraceResult {
            path: vec![],
            has_dynamic_refs,
            truncated,
        }
    }

    // =========================================================================
    // Ordered Recompute (Phase 1.2)
    // =========================================================================

    /// Perform a full ordered recompute of all formulas.
    ///
    /// Evaluates formulas in topological order (precedents before dependents)
    /// and returns a report with metrics.
    ///
    /// # Cycle Handling
    ///
    /// If cycles exist in the graph (e.g., from loading a legacy file), cycle
    /// cells are marked with #CYCLE! error and excluded from ordered recompute.
    ///
    /// # Unknown Dependencies
    ///
    /// Formulas with INDIRECT/OFFSET are evaluated after all known-deps formulas
    /// since their dependencies cannot be determined statically.
    pub fn recompute_full_ordered(&mut self) -> crate::recalc::RecalcReport {
        use crate::formula::analyze::has_dynamic_deps;
        use crate::recalc::{CellRecalcInfo, RecalcError, RecalcReport};
        use rustc_hash::FxHashMap;
        use std::time::Instant;

        let start = Instant::now();
        let mut report = RecalcReport::new();

        // Get topo order (or detect cycles)
        let (order, cycle_cells) = match self.dep_graph.topo_order_all_formulas() {
            Ok(order) => (order, Vec::new()),
            Err(cycle) => {
                report.had_cycles = true;
                // Mark cycle cells - we'll evaluate non-cycle cells only
                let cycle_cells = cycle.cells.clone();

                // Get partial order excluding cycle cells
                let all_formula_cells: Vec<CellId> = self.dep_graph.formula_cells().collect();
                let non_cycle: Vec<CellId> = all_formula_cells
                    .into_iter()
                    .filter(|c| !cycle_cells.contains(c))
                    .collect();
                (non_cycle, cycle_cells)
            }
        };

        // Mark cycle cells with #CYCLE! error
        for cell_id in &cycle_cells {
            if let Some(sheet) = self.sheet_by_id_mut(cell_id.sheet) {
                sheet.set_cycle_error(cell_id.row, cell_id.col);
            }
        }

        // Separate known-deps and unknown-deps formulas
        let mut known_deps_order = Vec::new();
        let mut unknown_deps_cells = Vec::new();

        for cell_id in order {
            if let Some(sheet) = self.sheet_by_id(cell_id.sheet) {
                if let Some(cell) = sheet.cells.get(&(cell_id.row, cell_id.col)) {
                    if let Some(ast) = cell.value.formula_ast() {
                        if has_dynamic_deps(ast) {
                            unknown_deps_cells.push(cell_id);
                        } else {
                            known_deps_order.push(cell_id);
                        }
                    } else {
                        known_deps_order.push(cell_id);
                    }
                }
            }
        }

        // Compute depths during evaluation
        // depth(value_cell) = 0, depth(formula_cell) = 1 + max(depth(precedents))
        let mut depths: FxHashMap<CellId, usize> = FxHashMap::default();

        // Track evaluation order for explainability
        let mut eval_order: usize = 0;

        // Evaluate known-deps formulas in topo order
        for cell_id in &known_deps_order {
            // Compute depth
            let mut max_pred_depth = 0;
            for pred in self.dep_graph.precedents(*cell_id) {
                let pred_depth = depths.get(&pred).copied().unwrap_or(0);
                max_pred_depth = max_pred_depth.max(pred_depth);
            }
            let cell_depth = max_pred_depth + 1;
            depths.insert(*cell_id, cell_depth);
            report.max_depth = report.max_depth.max(cell_depth);

            // Evaluate the formula
            if let Err(e) = self.evaluate_cell(*cell_id) {
                if report.errors.len() < 100 {
                    report.errors.push(RecalcError::new(*cell_id, e));
                }
            }

            // Track per-cell recalc info for Inspector explainability
            report.cell_info.insert(
                *cell_id,
                CellRecalcInfo::new(cell_depth, eval_order, false),
            );
            eval_order += 1;

            report.cells_recomputed += 1;
        }

        // Evaluate unknown-deps formulas (conservative: always recompute)
        // Sort by CellId for determinism
        unknown_deps_cells.sort_by(|a, b| {
            a.sheet.raw().cmp(&b.sheet.raw())
                .then(a.row.cmp(&b.row))
                .then(a.col.cmp(&b.col))
        });

        for cell_id in &unknown_deps_cells {
            // Unknown deps get depth = max_known_depth + 1 (after all known)
            let cell_depth = report.max_depth + 1;
            depths.insert(*cell_id, cell_depth);

            if let Err(e) = self.evaluate_cell(*cell_id) {
                if report.errors.len() < 100 {
                    report.errors.push(RecalcError::new(*cell_id, e));
                }
            }

            // Track per-cell recalc info (mark as having unknown deps)
            report.cell_info.insert(
                *cell_id,
                CellRecalcInfo::new(cell_depth, eval_order, true),
            );
            eval_order += 1;

            report.cells_recomputed += 1;
            report.unknown_deps_recomputed += 1;
        }

        // Update max_depth if we had unknown deps
        if !unknown_deps_cells.is_empty() {
            report.max_depth += 1;
        }

        report.duration_ms = start.elapsed().as_millis() as u64;
        report
    }

    /// Evaluate a single cell's formula and return the result.
    ///
    /// This forces evaluation by reading the cell value through the workbook lookup.
    fn evaluate_cell(&self, cell_id: CellId) -> Result<(), String> {
        use crate::formula::eval::evaluate;
        use crate::formula::parser::bind_expr;

        let sheet = self.sheet_by_id(cell_id.sheet)
            .ok_or_else(|| format!("Sheet not found: {:?}", cell_id.sheet))?;

        let cell = sheet.cells.get(&(cell_id.row, cell_id.col))
            .ok_or_else(|| format!("Cell not found: {:?}", cell_id))?;

        if let Some(ast) = cell.value.formula_ast() {
            let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));
            let lookup = WorkbookLookup::with_cell_context(self, cell_id.sheet, cell_id.row, cell_id.col);
            let result = evaluate(&bound, &lookup);

            // Check for error result
            if let crate::formula::eval::EvalResult::Error(e) = result {
                return Err(e);
            }
        }

        Ok(())
    }

    /// Check if setting a formula at the given cell would create a cycle.
    ///
    /// Returns `Err(CycleReport)` if the formula would introduce a circular reference.
    /// The formula should NOT be applied if this returns an error.
    pub fn check_formula_cycle(
        &self,
        sheet_id: SheetId,
        row: usize,
        col: usize,
        formula: &str,
    ) -> Result<(), crate::recalc::CycleReport> {
        use crate::formula::parser::{parse, bind_expr};
        use crate::formula::refs::extract_cell_ids;

        // Parse and bind
        let parsed = parse(formula).map_err(|e| {
            crate::recalc::CycleReport::new(vec![], format!("Parse error: {}", e))
        })?;
        let bound = bind_expr(&parsed, |name| self.sheet_id_by_name(name));

        // Extract new precedents
        let new_preds = extract_cell_ids(
            &bound,
            sheet_id,
            &self.named_ranges,
            |idx| self.sheet_id_at_idx(idx),
        );

        // Check for cycle
        let cell_id = CellId::new(sheet_id, row, col);
        if let Some(cycle) = self.dep_graph.would_create_cycle(cell_id, &new_preds) {
            return Err(cycle);
        }

        Ok(())
    }

    // =========================================================================
    // Cross-Sheet Cell Evaluation
    // =========================================================================

    /// Get the display value of a cell, with full cross-sheet reference support.
    /// This should be used instead of sheet.get_display() when cross-sheet refs are possible.
    pub fn get_cell_display(&self, sheet_idx: usize, row: usize, col: usize) -> String {
        use crate::cell::CellValue;
        use crate::formula::parser::bind_expr;
        use crate::formula::eval::{evaluate, EvalResult};

        let sheet = match self.sheets.get(sheet_idx) {
            Some(s) => s,
            None => return String::new(),
        };

        let cell = sheet.get_cell(row, col);

        match &cell.value {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                CellValue::format_number(*n, &cell.format.number_format)
            }
            CellValue::Formula { ast: Some(ast), .. } => {
                // Bind with workbook context for cross-sheet refs
                let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));
                let sheet_id = sheet.id;
                let lookup = WorkbookLookup::with_cell_context(self, sheet_id, row, col);
                let result = evaluate(&bound, &lookup);

                match result {
                    EvalResult::Number(n) => CellValue::format_number(n, &cell.format.number_format),
                    EvalResult::Text(s) => s,
                    EvalResult::Boolean(b) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                    EvalResult::Error(e) => e,
                    EvalResult::Array(arr) => arr.top_left().to_text(),
                }
            }
            CellValue::Formula { ast: None, .. } => "#ERR".to_string(),
        }
    }

    /// Get the text value of a cell, with full cross-sheet reference support.
    pub fn get_cell_text(&self, sheet_idx: usize, row: usize, col: usize) -> String {
        use crate::cell::CellValue;
        use crate::formula::parser::bind_expr;
        use crate::formula::eval::evaluate;

        let sheet = match self.sheets.get(sheet_idx) {
            Some(s) => s,
            None => return String::new(),
        };

        let cell = sheet.get_cell(row, col);

        match &cell.value {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::Formula { ast: Some(ast), .. } => {
                let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));
                let sheet_id = sheet.id;
                let lookup = WorkbookLookup::with_cell_context(self, sheet_id, row, col);
                evaluate(&bound, &lookup).to_text()
            }
            CellValue::Formula { ast: None, .. } => String::new(),
        }
    }
}

// =============================================================================
// WorkbookLookup - CellLookup implementation with cross-sheet support
// =============================================================================

/// A CellLookup implementation that supports cross-sheet references.
///
/// This wraps a Workbook reference and provides data access for formula evaluation.
/// Same-sheet lookups (get_value, get_text) access the current sheet.
/// Cross-sheet lookups (get_value_sheet, get_text_sheet) access any sheet by ID.
pub struct WorkbookLookup<'a> {
    workbook: &'a Workbook,
    current_sheet_id: SheetId,
    current_cell: Option<(usize, usize)>,
}

impl<'a> WorkbookLookup<'a> {
    /// Create a new WorkbookLookup for the given sheet
    pub fn new(workbook: &'a Workbook, current_sheet_id: SheetId) -> Self {
        Self {
            workbook,
            current_sheet_id,
            current_cell: None,
        }
    }

    /// Create a new WorkbookLookup with cell context (for ROW()/COLUMN() without args)
    pub fn with_cell_context(workbook: &'a Workbook, current_sheet_id: SheetId, row: usize, col: usize) -> Self {
        Self {
            workbook,
            current_sheet_id,
            current_cell: Some((row, col)),
        }
    }

    /// Get the current sheet
    fn current_sheet(&self) -> Option<&Sheet> {
        self.workbook.sheet_by_id(self.current_sheet_id)
    }
}

impl<'a> CellLookup for WorkbookLookup<'a> {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        self.current_sheet()
            .map(|sheet| {
                // Use the sheet's get_text to handle formulas, then parse
                let text = sheet.get_text(row, col);
                text.parse().unwrap_or(0.0)
            })
            .unwrap_or(0.0)
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        self.current_sheet()
            .map(|sheet| sheet.get_text(row, col))
            .unwrap_or_default()
    }

    fn get_value_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Value {
        match self.workbook.sheet_by_id(sheet_id) {
            Some(sheet) => {
                let text = sheet.get_text(row, col);
                if text.is_empty() {
                    Value::Empty
                } else if text.starts_with('#') {
                    // Error value (e.g., #REF!, #VALUE!)
                    Value::Error(text)
                } else if let Ok(n) = text.parse::<f64>() {
                    Value::Number(n)
                } else if text.eq_ignore_ascii_case("TRUE") {
                    Value::Boolean(true)
                } else if text.eq_ignore_ascii_case("FALSE") {
                    Value::Boolean(false)
                } else {
                    Value::Text(text)
                }
            }
            None => Value::Error("#REF!".to_string()),
        }
    }

    fn get_text_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> String {
        match self.workbook.sheet_by_id(sheet_id) {
            Some(sheet) => sheet.get_text(row, col),
            None => "#REF!".to_string(),
        }
    }

    fn resolve_named_range(&self, name: &str) -> Option<NamedRangeResolution> {
        use crate::named_range::NamedRangeTarget;
        self.workbook.named_ranges.get(name).map(|nr| {
            match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => {
                    NamedRangeResolution::Cell { row: *row, col: *col }
                }
                NamedRangeTarget::Range { start_row, start_col, end_row, end_col, .. } => {
                    NamedRangeResolution::Range {
                        start_row: *start_row,
                        start_col: *start_col,
                        end_row: *end_row,
                        end_col: *end_col,
                    }
                }
            }
        })
    }

    fn current_cell(&self) -> Option<(usize, usize)> {
        self.current_cell
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.active_sheet_index(), 0);
        assert_eq!(wb.active_sheet().name, "Sheet1");
    }

    #[test]
    fn test_add_sheet() {
        let mut wb = Workbook::new();
        let idx = wb.add_sheet();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_count(), 2);
        assert_eq!(wb.sheet(1).unwrap().name, "Sheet2");
    }

    #[test]
    fn test_navigation() {
        let mut wb = Workbook::new();
        wb.add_sheet();
        wb.add_sheet();

        assert_eq!(wb.active_sheet_index(), 0);
        assert!(wb.next_sheet());
        assert_eq!(wb.active_sheet_index(), 1);
        assert!(wb.next_sheet());
        assert_eq!(wb.active_sheet_index(), 2);
        assert!(!wb.next_sheet()); // Can't go further

        assert!(wb.prev_sheet());
        assert_eq!(wb.active_sheet_index(), 1);
    }

    #[test]
    fn test_delete_sheet() {
        let mut wb = Workbook::new();
        wb.add_sheet();
        wb.add_sheet();
        wb.set_active_sheet(2);

        assert!(wb.delete_sheet(1));
        assert_eq!(wb.sheet_count(), 2);
        assert_eq!(wb.active_sheet_index(), 1); // Adjusted

        // Can't delete last sheet
        assert!(wb.delete_sheet(0));
        assert!(!wb.delete_sheet(0)); // Last one, can't delete
    }

    // =========================================================================
    // Cross-Sheet Reference Acceptance Tests
    // =========================================================================

    use crate::formula::parser::{parse, bind_expr, format_expr};
    use crate::formula::eval::evaluate;

    #[test]
    fn test_cross_sheet_cell_ref() {
        // =Sheet2!A1 returns value from sheet2
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        // Put value in Sheet2!A1
        wb.sheet_mut(1).unwrap().set_value(0, 0, "42");

        // Parse and bind formula
        let parsed = parse("=Sheet2!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        // Evaluate with WorkbookLookup
        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);

        assert_eq!(result.to_text(), "42");
    }

    #[test]
    fn test_cross_sheet_quoted_name() {
        // ='My Sheet'!A1 works
        let mut wb = Workbook::new();
        wb.add_sheet_named("My Sheet").unwrap();

        // Put value in 'My Sheet'!A1
        wb.sheet_mut(1).unwrap().set_value(0, 0, "hello");

        let parsed = parse("='My Sheet'!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);

        assert_eq!(result.to_text(), "hello");
    }

    #[test]
    fn test_cross_sheet_escaped_quote() {
        // ='Bob''s Sheet'!A1 works
        let mut wb = Workbook::new();
        wb.add_sheet_named("Bob's Sheet").unwrap();

        wb.sheet_mut(1).unwrap().set_value(0, 0, "test");

        let parsed = parse("='Bob''s Sheet'!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);

        assert_eq!(result.to_text(), "test");
    }

    #[test]
    fn test_cross_sheet_range_sum() {
        // =SUM(Sheet2!A1:A3) works
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        // Put values in Sheet2!A1:A3
        wb.sheet_mut(1).unwrap().set_value(0, 0, "10");
        wb.sheet_mut(1).unwrap().set_value(1, 0, "20");
        wb.sheet_mut(1).unwrap().set_value(2, 0, "30");

        let parsed = parse("=SUM(Sheet2!A1:A3)").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);

        assert_eq!(result.to_text(), "60");
    }

    #[test]
    fn test_cross_sheet_rename_updates_print() {
        // Rename sheet2 → formulas print new name
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        let parsed = parse("=Sheet2!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        // Formula should print "Sheet2"
        let formula_str = format_expr(&bound, |id| wb.sheet_by_id(id).map(|s| s.name.clone()));
        assert_eq!(formula_str, "=Sheet2!A1");

        // Rename Sheet2 to "Data"
        wb.rename_sheet(1, "Data");

        // Formula should now print "Data"
        let formula_str = format_expr(&bound, |id| wb.sheet_by_id(id).map(|s| s.name.clone()));
        assert_eq!(formula_str, "=Data!A1");
    }

    #[test]
    fn test_cross_sheet_insert_preserves_refs() {
        // Insert/reorder sheets → references still correct
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2
        wb.sheet_mut(1).unwrap().set_value(0, 0, "original");

        // Get Sheet2's ID before any changes
        let sheet2_id = wb.sheet_id_at_idx(1).unwrap();

        let parsed = parse("=Sheet2!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        // Add another sheet (Sheet3) - this doesn't change indices but tests stability
        wb.add_sheet();

        // Reference should still work (SheetId is stable)
        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        assert_eq!(result.to_text(), "original");

        // Sheet2 still has the same ID
        assert_eq!(wb.sheet_id_at_idx(1).unwrap(), sheet2_id);
    }

    #[test]
    fn test_cross_sheet_delete_becomes_ref_error() {
        // Delete referenced sheet → formula evaluates #REF!
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2
        wb.add_sheet(); // Sheet3

        // Store formula referencing Sheet2
        let parsed = parse("=Sheet2!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        // Verify it works before deletion
        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        assert!(!result.to_text().contains("#REF"));

        // Delete Sheet2
        wb.delete_sheet(1);

        // Now evaluation should return #REF! because sheet ID no longer exists
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        assert!(result.to_text().contains("#REF"), "Expected #REF! but got: {}", result.to_text());
    }

    #[test]
    fn test_cross_sheet_unknown_sheet_ref_error() {
        // Reference to unknown sheet → #REF! at bind time
        let wb = Workbook::new();

        let parsed = parse("=NonExistent!A1").unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));

        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);

        assert!(result.to_text().contains("#REF"), "Expected #REF! but got: {}", result.to_text());
    }

    #[test]
    fn test_cross_sheet_rename_eval_twice_still_correct() {
        // Regression test: rename sheet → evaluate formula twice → still correct
        // This guards the "bind every eval" contract (no accidental caching)
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2
        wb.sheet_mut(1).unwrap().set_value(0, 0, "42");

        // Parse formula (stores name, not ID)
        let parsed = parse("=Sheet2!A1").unwrap();

        // Evaluate before rename
        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        assert_eq!(result.to_text(), "42");

        // Rename Sheet2 to Data
        wb.rename_sheet(1, "Data");

        // Re-bind and evaluate - should still work (formula stores "Sheet2" name)
        // but binding now fails because "Sheet2" doesn't exist
        let bound = bind_expr(&parsed, |name| wb.sheet_id_by_name(name));
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        // After rename, "Sheet2" no longer exists, so this becomes #REF!
        assert!(result.to_text().contains("#REF"), "Expected #REF! after rename, got: {}", result.to_text());

        // But if we parse with the new name, it works
        let parsed_new = parse("=Data!A1").unwrap();
        let bound = bind_expr(&parsed_new, |name| wb.sheet_id_by_name(name));
        let lookup = WorkbookLookup::new(&wb, sheet1_id);
        let result = evaluate(&bound, &lookup);
        assert_eq!(result.to_text(), "42");
    }

    // =========================================================================
    // Dependency Graph Tests
    // =========================================================================

    #[test]
    fn test_dep_graph_simple_formula() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set up data cells
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10"); // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "20"); // B1

        // Set a formula that references A1 and B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1+B1"); // C1

        // Update deps for the formula cell
        wb.update_cell_deps(sheet_id, 0, 2);

        // Check precedents of C1
        let preds = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds.len(), 2);

        let a1 = CellId::new(sheet_id, 0, 0);
        let b1 = CellId::new(sheet_id, 0, 1);
        assert!(preds.contains(&a1));
        assert!(preds.contains(&b1));

        // Check dependents of A1
        let deps = wb.get_dependents(sheet_id, 0, 0);
        assert_eq!(deps.len(), 1);
        let c1 = CellId::new(sheet_id, 0, 2);
        assert!(deps.contains(&c1));
    }

    #[test]
    fn test_dep_graph_update_formula() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set initial formula =A1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1"); // C1
        wb.update_cell_deps(sheet_id, 0, 2);

        let preds = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds.len(), 1);
        let a1 = CellId::new(sheet_id, 0, 0);
        assert!(preds.contains(&a1));

        // Update formula to =B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1"); // C1
        wb.update_cell_deps(sheet_id, 0, 2);

        // Now C1 should depend on B1, not A1
        let preds = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds.len(), 1);
        let b1 = CellId::new(sheet_id, 0, 1);
        assert!(preds.contains(&b1));

        // A1 should have no dependents now
        let deps = wb.get_dependents(sheet_id, 0, 0);
        assert_eq!(deps.len(), 0);
    }

    #[test]
    fn test_dep_graph_clear_deps() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set formula
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1"); // C1
        wb.update_cell_deps(sheet_id, 0, 2);

        // Verify deps exist
        let preds = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds.len(), 1);

        // Clear deps
        wb.clear_cell_deps(sheet_id, 0, 2);

        // Deps should be gone
        let preds = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds.len(), 0);

        // A1 should have no dependents
        let deps = wb.get_dependents(sheet_id, 0, 0);
        assert_eq!(deps.len(), 0);
    }

    #[test]
    fn test_dep_graph_range_expansion() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set formula with range
        wb.sheet_mut(0).unwrap().set_value(5, 0, "=SUM(A1:A5)"); // A6
        wb.update_cell_deps(sheet_id, 5, 0);

        // Should have 5 precedents (A1:A5)
        let preds = wb.get_precedents(sheet_id, 5, 0);
        assert_eq!(preds.len(), 5);

        // Each cell A1-A5 should have A6 as dependent
        for row in 0..5 {
            let deps = wb.get_dependents(sheet_id, row, 0);
            assert_eq!(deps.len(), 1);
            let a6 = CellId::new(sheet_id, 5, 0);
            assert!(deps.contains(&a6));
        }
    }

    #[test]
    fn test_dep_graph_cross_sheet_refs() {
        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        let sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let sheet2_id = wb.sheet_id_at_idx(1).unwrap();

        // Set value on Sheet2
        wb.sheet_mut(1).unwrap().set_value(0, 0, "100"); // Sheet2!A1

        // Set formula on Sheet1 referencing Sheet2
        wb.sheet_mut(0).unwrap().set_value(0, 0, "=Sheet2!A1"); // Sheet1!A1
        wb.update_cell_deps(sheet1_id, 0, 0);

        // Check precedents - should reference Sheet2!A1
        let preds = wb.get_precedents(sheet1_id, 0, 0);
        assert_eq!(preds.len(), 1);
        let sheet2_a1 = CellId::new(sheet2_id, 0, 0);
        assert!(preds.contains(&sheet2_a1));

        // Check dependents of Sheet2!A1 - should be Sheet1!A1
        let deps = wb.get_dependents(sheet2_id, 0, 0);
        assert_eq!(deps.len(), 1);
        let sheet1_a1 = CellId::new(sheet1_id, 0, 0);
        assert!(deps.contains(&sheet1_a1));
    }

    #[test]
    fn test_dep_graph_rebuild() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set up some formulas without updating deps
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");       // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1*2");    // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1+A1");   // C1

        // Initially no deps tracked
        assert_eq!(wb.get_precedents(sheet_id, 0, 1).len(), 0);
        assert_eq!(wb.get_precedents(sheet_id, 0, 2).len(), 0);

        // Rebuild the graph
        wb.rebuild_dep_graph();

        // Now B1 should have A1 as precedent
        let preds_b1 = wb.get_precedents(sheet_id, 0, 1);
        assert_eq!(preds_b1.len(), 1);
        let a1 = CellId::new(sheet_id, 0, 0);
        assert!(preds_b1.contains(&a1));

        // C1 should have A1 and B1 as precedents
        let preds_c1 = wb.get_precedents(sheet_id, 0, 2);
        assert_eq!(preds_c1.len(), 2);
        let b1 = CellId::new(sheet_id, 0, 1);
        assert!(preds_c1.contains(&a1));
        assert!(preds_c1.contains(&b1));

        // A1 should have B1 and C1 as dependents
        let deps_a1 = wb.get_dependents(sheet_id, 0, 0);
        assert_eq!(deps_a1.len(), 2);
        let c1 = CellId::new(sheet_id, 0, 2);
        assert!(deps_a1.contains(&b1));
        assert!(deps_a1.contains(&c1));
    }

    #[test]
    fn test_dep_graph_clear_on_non_formula() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set formula
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1+B1"); // C1
        wb.update_cell_deps(sheet_id, 0, 2);

        assert_eq!(wb.get_precedents(sheet_id, 0, 2).len(), 2);

        // Replace with a plain value
        wb.sheet_mut(0).unwrap().set_value(0, 2, "42"); // C1
        wb.update_cell_deps(sheet_id, 0, 2);

        // Deps should be cleared
        assert_eq!(wb.get_precedents(sheet_id, 0, 2).len(), 0);
        assert_eq!(wb.get_dependents(sheet_id, 0, 0).len(), 0);
        assert_eq!(wb.get_dependents(sheet_id, 0, 1).len(), 0);
    }

    // =========================================================================
    // Ordered Recompute Tests (Phase 1.2)
    // =========================================================================

    #[test]
    fn test_recompute_depth_chain() {
        // A1=1, B1=A1, C1=B1 → depth should be 2 (B1 depth 1, C1 depth 2)
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1 = 1 (value)
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1");    // B1 = A1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1");    // C1 = B1

        wb.update_cell_deps(sheet_id, 0, 1);
        wb.update_cell_deps(sheet_id, 0, 2);

        let report = wb.recompute_full_ordered();

        assert_eq!(report.cells_recomputed, 2); // B1 and C1
        assert_eq!(report.max_depth, 2);        // B1=1, C1=2
        assert!(!report.had_cycles);
        assert_eq!(report.unknown_deps_recomputed, 0);
    }

    #[test]
    fn test_recompute_depth_diamond() {
        // A1=1, B1=A1, C1=A1, D1=B1+C1 → max_depth should be 2
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");          // A1 = 1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1");        // B1 = A1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1");        // C1 = A1
        wb.sheet_mut(0).unwrap().set_value(0, 3, "=B1+C1");     // D1 = B1+C1

        wb.update_cell_deps(sheet_id, 0, 1);
        wb.update_cell_deps(sheet_id, 0, 2);
        wb.update_cell_deps(sheet_id, 0, 3);

        let report = wb.recompute_full_ordered();

        assert_eq!(report.cells_recomputed, 3); // B1, C1, D1
        assert_eq!(report.max_depth, 2);        // B1=1, C1=1, D1=2
        assert!(!report.had_cycles);
    }

    #[test]
    fn test_recompute_report_metrics() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Set up a value cell
        wb.sheet_mut(0).unwrap().set_value(0, 0, "1"); // A1 = 1

        // Set up formulas that reference A1
        for i in 1..=10 {
            wb.sheet_mut(0).unwrap().set_value(i, 0, "=A1");
            wb.update_cell_deps(sheet_id, i, 0);
        }

        let report = wb.recompute_full_ordered();

        assert_eq!(report.cells_recomputed, 10);
        assert_eq!(report.max_depth, 1); // All depth 1 (direct ref to value)
        assert!(!report.had_cycles);
    }

    #[test]
    fn test_recompute_unknown_deps_indirect() {
        // Formula with INDIRECT should be marked as unknown deps
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "A2");             // A1 = "A2" (text)
        wb.sheet_mut(0).unwrap().set_value(1, 0, "42");             // A2 = 42
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=INDIRECT(A1)");  // B1 = INDIRECT(A1)

        wb.update_cell_deps(sheet_id, 0, 1);

        let report = wb.recompute_full_ordered();

        assert_eq!(report.cells_recomputed, 1);
        assert_eq!(report.unknown_deps_recomputed, 1);
    }

    #[test]
    fn test_recompute_cycle_on_load() {
        // Simulate loading a workbook with cycles by directly manipulating the graph
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // Create cells with formulas
        wb.sheet_mut(0).unwrap().set_value(0, 0, "=B1"); // A1 = B1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1"); // B1 = A1

        // Update deps to create the cycle
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        // Recompute should detect cycle and not panic
        let report = wb.recompute_full_ordered();

        assert!(report.had_cycles);
        // Cycle cells should be marked with #CYCLE!
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "#CYCLE!");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "#CYCLE!");
    }

    #[test]
    fn test_recompute_cross_sheet() {
        let mut wb = Workbook::new();
        wb.add_sheet();

        let _sheet1_id = wb.sheet_id_at_idx(0).unwrap();
        let sheet2_id = wb.sheet_id_at_idx(1).unwrap();

        // Sheet1!A1 = 10
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");
        // Sheet2!A1 = Sheet1!A1
        wb.sheet_mut(1).unwrap().set_value(0, 0, "=Sheet1!A1");
        wb.update_cell_deps(sheet2_id, 0, 0);

        let report = wb.recompute_full_ordered();

        assert_eq!(report.cells_recomputed, 1);
        assert_eq!(report.max_depth, 1);
        assert!(!report.had_cycles);
    }

    #[test]
    fn test_check_formula_cycle_self_reference() {
        let wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = A1 should be detected as cycle
        let result = wb.check_formula_cycle(sheet_id, 0, 0, "=A1");
        assert!(result.is_err());
    }

    #[test]
    fn test_check_formula_cycle_indirect() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = B1
        wb.sheet_mut(0).unwrap().set_value(0, 0, "=B1");
        wb.update_cell_deps(sheet_id, 0, 0);

        // B1 = A1 should be detected as cycle
        let result = wb.check_formula_cycle(sheet_id, 0, 1, "=A1");
        assert!(result.is_err());
    }

    #[test]
    fn test_check_formula_cycle_valid() {
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = 10
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");

        // B1 = A1 should be valid
        let result = wb.check_formula_cycle(sheet_id, 0, 1, "=A1");
        assert!(result.is_ok());
    }
}
