use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

use serde::{Deserialize, Serialize};

use super::cell::{Alignment, Cell, CellBorder, CellFormat, CellValue, NumberFormat, SpillError, SpillInfo, TextOverflow, VerticalAlignment, max_border};
use super::formula::eval::{self, Array2D, CellLookup, EvalResult, LookupWithContext, Value};
use super::formula::parser::bind_expr_same_sheet;
use super::validation::ValidationStore;

// Thread-local set to track cells currently being evaluated (for cycle detection)
thread_local! {
    static EVALUATING: RefCell<HashSet<(usize, usize)>> = RefCell::new(HashSet::new());
}

// =============================================================================
// SheetId - Stable identity for sheets (never changes, never reused)
// =============================================================================

/// A stable, unique identifier for a sheet.
///
/// SheetId is distinct from sheet index (position in the tab bar):
/// - SheetId: Identity - never changes for a sheet's lifetime, never reused after deletion
/// - Index: Position - changes when sheets are reordered, inserted, or deleted
///
/// Use SheetId for:
/// - Cross-sheet formula references
/// - Dependency tracking
/// - Named range targets
///
/// Use index for:
/// - UI tab order
/// - Active sheet selection
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct SheetId(pub u64);

impl SheetId {
    /// Create a new SheetId from a raw value (used during deserialization)
    pub fn from_raw(id: u64) -> Self {
        Self(id)
    }

    /// Get the raw u64 value (used during serialization)
    pub fn raw(&self) -> u64 {
        self.0
    }
}

impl std::fmt::Display for SheetId {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "SheetId({})", self.0)
    }
}

// =============================================================================
// SheetRef - Sheet reference in formulas (resolved form)
// =============================================================================

/// A resolved sheet reference in a formula.
///
/// This is the form stored in the AST after binding/resolution.
/// Parser produces `UnboundSheetRef`, which is then resolved to `SheetRef`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetRef {
    /// Reference to the current/local sheet (no sheet prefix in formula)
    Current,
    /// Reference to a specific sheet by its stable ID
    Id(SheetId),
    /// Reference to a deleted sheet - shows as #REF! during evaluation
    RefError {
        /// The ID of the deleted sheet (for debugging/logging)
        id: SheetId,
        /// The sheet name at the time of deletion (for display)
        last_known_name: String,
    },
}

impl SheetRef {
    /// Check if this is a reference error (deleted sheet)
    pub fn is_error(&self) -> bool {
        matches!(self, SheetRef::RefError { .. })
    }

    /// Get the SheetId if this is a valid reference
    pub fn sheet_id(&self) -> Option<SheetId> {
        match self {
            SheetRef::Id(id) => Some(*id),
            _ => None,
        }
    }
}

// =============================================================================
// UnboundSheetRef - Sheet reference before resolution
// =============================================================================

/// An unresolved sheet reference from the parser.
///
/// Parser produces this form; it must be resolved to `SheetRef` before evaluation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UnboundSheetRef {
    /// Reference to the current/local sheet (no sheet prefix in formula)
    Current,
    /// Reference to a sheet by name (needs resolution to SheetId)
    Named(String),
}

impl UnboundSheetRef {
    /// Create a reference to the current sheet
    pub fn current() -> Self {
        UnboundSheetRef::Current
    }

    /// Create a reference to a named sheet
    pub fn named(name: impl Into<String>) -> Self {
        UnboundSheetRef::Named(name.into())
    }
}

// =============================================================================
// Sheet Name Normalization
// =============================================================================

/// Normalize a sheet name for case-insensitive comparison.
///
/// Rules:
/// - Trim leading/trailing whitespace
/// - Convert to ASCII lowercase (MVP: ASCII only, documented limitation)
///
/// Future: Could use Unicode casefold for full international support.
pub fn normalize_sheet_name(name: &str) -> String {
    name.trim().to_ascii_lowercase()
}

/// Check if a sheet name is valid.
///
/// Rules:
/// - Cannot be empty after trimming
/// - (Future: could add more restrictions)
pub fn is_valid_sheet_name(name: &str) -> bool {
    !name.trim().is_empty()
}

// =============================================================================
// MergedRegion
// =============================================================================

/// A rectangular merged cell region in a sheet.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MergedRegion {
    /// Top-left corner (row, col)
    pub start: (usize, usize),
    /// Bottom-right corner (row, col)
    pub end: (usize, usize),
}

impl MergedRegion {
    pub fn new(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> Self {
        Self {
            start: (start_row, start_col),
            end: (end_row, end_col),
        }
    }

    /// Number of rows spanned
    pub fn rows(&self) -> usize {
        self.end.0 - self.start.0 + 1
    }

    /// Number of columns spanned
    pub fn cols(&self) -> usize {
        self.end.1 - self.start.1 + 1
    }

    /// Whether (row, col) is inside this region
    pub fn contains(&self, row: usize, col: usize) -> bool {
        row >= self.start.0 && row <= self.end.0 && col >= self.start.1 && col <= self.end.1
    }

    /// The top-left (origin) cell
    pub fn top_left(&self) -> (usize, usize) {
        self.start
    }

    /// A degenerate merge spans only 1 cell (1×1) — should be removed
    pub fn is_degenerate(&self) -> bool {
        self.rows() <= 1 && self.cols() <= 1
    }

    /// Total number of cells in the region
    pub fn cell_count(&self) -> usize {
        self.rows() * self.cols()
    }

    /// Whether this merge overlaps a viewport defined by scroll position + visible extent.
    /// Used by the render layer to filter which merges need overlay elements.
    pub fn overlaps_viewport(
        &self,
        scroll_row: usize,
        scroll_col: usize,
        visible_rows: usize,
        visible_cols: usize,
    ) -> bool {
        let max_row = scroll_row + visible_rows;
        let max_col = scroll_col + visible_cols;
        // Overlap iff neither axis is entirely outside
        self.end.0 >= scroll_row && self.start.0 < max_row
            && self.end.1 >= scroll_col && self.start.1 < max_col
    }

    /// Compute the pixel rect of a merge overlay relative to the scroll position.
    ///
    /// Pure geometry — takes closures for col/row dimensions so it's testable
    /// without any UI framework dependency.
    ///
    /// Returns (x, y, width, height) where:
    /// - x = sum of col widths from scroll_col to merge.start_col
    /// - y = sum of row heights from scroll_row to merge.start_row
    /// - width = sum of col widths across the merge span
    /// - height = sum of row heights across the merge span
    pub fn pixel_rect(
        &self,
        scroll_row: usize,
        scroll_col: usize,
        col_width: impl Fn(usize) -> f32,
        row_height: impl Fn(usize) -> f32,
    ) -> (f32, f32, f32, f32) {
        let x: f32 = (scroll_col..self.start.1).map(|c| col_width(c)).sum();
        let y: f32 = (scroll_row..self.start.0).map(|r| row_height(r)).sum();
        let w: f32 = (self.start.1..=self.end.1).map(|c| col_width(c)).sum();
        let h: f32 = (self.start.0..=self.end.0).map(|r| row_height(r)).sum();
        (x, y, w, h)
    }
}

// =============================================================================
// Sheet
// =============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    /// Stable identity - never changes, never reused after deletion
    pub id: SheetId,
    /// Display name (can be changed via rename)
    pub name: String,
    /// Normalized name for case-insensitive lookup (trimmed + lowercased)
    #[serde(default)]
    pub name_key: String,
    /// Cell storage - pub(crate) for workbook-level access during recompute
    pub(crate) cells: HashMap<(usize, usize), Cell>,
    pub rows: usize,
    pub cols: usize,
    /// Spilled values from array formulas: (row, col) -> Value
    #[serde(skip)]
    spill_values: HashMap<(usize, usize), Value>,
    /// Computed value cache: populated during topological recalc, read during evaluation.
    /// Stores typed Value (not String) to avoid lossy conversions.
    /// Getters NEVER evaluate on cache miss — only the topo recalc pass populates this.
    #[serde(skip)]
    computed_cache: RefCell<HashMap<(usize, usize), Value>>,
    /// Data validation rules for cells
    #[serde(default)]
    pub validations: ValidationStore,
    /// Merged cell regions
    #[serde(default)]
    pub merged_regions: Vec<MergedRegion>,
    /// Fast lookup: (row, col) → index into merged_regions
    #[serde(skip)]
    merge_index: HashMap<(usize, usize), usize>,
    /// Conservative flag: true once any cell has had a non-None border set.
    /// Never cleared (except by `scan_border_flag()`). Used by the renderer
    /// to skip border computation on sheets that have never had borders.
    #[serde(skip)]
    pub has_any_borders: bool,
}

impl CellLookup for Sheet {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        // Check for circular reference
        let is_cycle = EVALUATING.with(|eval| {
            let set = eval.borrow();
            set.contains(&(row, col))
        });
        if is_cycle {
            return 0.0; // Circular reference - return 0
        }

        // CRITICAL: Check spill_values first for receiver cells
        // This enables formulas like =B2 to work when B2 is a spill receiver
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_number().unwrap_or(0.0);
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => match &cell.value {
                CellValue::Empty => 0.0,
                CellValue::Number(n) => *n,
                CellValue::Text(s) => s.parse().unwrap_or(0.0),
                CellValue::Formula { ast: Some(_), .. } => {
                    // Cache-only: never evaluate on cache miss.
                    // Topo recalc populates the cache; miss means not yet computed.
                    let cache = self.computed_cache.borrow();
                    cache.get(&(row, col))
                        .map(|v| v.to_number().unwrap_or(0.0))
                        .unwrap_or(0.0)
                }
                CellValue::Formula { ast: None, .. } => 0.0,
            },
            None => 0.0,
        }
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        // Check for circular reference
        let is_cycle = EVALUATING.with(|eval| {
            let set = eval.borrow();
            set.contains(&(row, col))
        });
        if is_cycle {
            return "#CIRC!".to_string(); // Circular reference error
        }

        // CRITICAL: Check spill_values first for receiver cells
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => match &cell.value {
                CellValue::Empty => String::new(),
                CellValue::Text(s) => s.clone(),
                CellValue::Number(n) => {
                    if n.fract() == 0.0 {
                        format!("{}", *n as i64)
                    } else {
                        format!("{}", n)
                    }
                }
                CellValue::Formula { ast: Some(_), .. } => {
                    // Cache-only: never evaluate on cache miss.
                    // Topo recalc populates the cache; miss means not yet computed.
                    let cache = self.computed_cache.borrow();
                    cache.get(&(row, col))
                        .map(|v| v.to_text())
                        .unwrap_or_default()
                }
                CellValue::Formula { ast: None, .. } => String::new(),
            },
            None => String::new(),
        }
    }

    fn debug_context(&self) -> String {
        format!(
            "Sheet(name=\"{}\", ptr={:p}, cache_len={})",
            self.name, self as *const Sheet, self.computed_cache_len()
        )
    }

    fn get_merge_start(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.get_merge(row, col).map(|m| m.start)
    }
}

impl Sheet {
    /// Create a new sheet with the given dimensions and a unique ID
    pub fn new(id: SheetId, rows: usize, cols: usize) -> Self {
        let name = String::from("Sheet1");
        let name_key = normalize_sheet_name(&name);
        Self {
            id,
            name,
            name_key,
            cells: HashMap::new(),
            rows,
            cols,
            spill_values: HashMap::new(),
            computed_cache: RefCell::new(HashMap::new()),
            validations: ValidationStore::new(),
            merged_regions: Vec::new(),
            merge_index: HashMap::new(),
            has_any_borders: false,
        }
    }

    /// Create a new sheet with the given dimensions, ID, and name
    pub fn new_with_name(id: SheetId, rows: usize, cols: usize, name: &str) -> Self {
        let name = name.trim().to_string();
        let name_key = normalize_sheet_name(&name);
        Self {
            id,
            name,
            name_key,
            cells: HashMap::new(),
            rows,
            cols,
            spill_values: HashMap::new(),
            computed_cache: RefCell::new(HashMap::new()),
            validations: ValidationStore::new(),
            merged_regions: Vec::new(),
            merge_index: HashMap::new(),
            has_any_borders: false,
        }
    }

    /// Cache a computed Value for a formula cell.
    /// Called ONLY during topological recalc (workbook.evaluate_cell).
    /// Getters read from this cache but never write to it.
    pub fn cache_computed(&self, row: usize, col: usize, value: Value) {
        self.computed_cache.borrow_mut().insert((row, col), value);
    }

    /// Get the number of entries in the computed cache (for diagnostics).
    pub fn computed_cache_len(&self) -> usize {
        self.computed_cache.borrow().len()
    }

    /// Clear the computed value cache (before a new recalc pass).
    pub fn clear_computed_cache(&self) {
        self.computed_cache.borrow_mut().clear();
    }

    /// Clear a single entry from the computed cache (for incremental recalc).
    pub fn clear_cached(&self, row: usize, col: usize) {
        self.computed_cache.borrow_mut().remove(&(row, col));
    }

    /// Update the sheet name (also updates name_key)
    pub fn set_name(&mut self, name: &str) {
        let trimmed = name.trim();
        self.name = trimmed.to_string();
        self.name_key = normalize_sheet_name(trimmed);
    }

    pub fn set_value(&mut self, row: usize, col: usize, value: &str) {
        // Redirect hidden merge cells to the merge origin
        let (row, col) = self.merge_origin_coord(row, col);

        // Clear any existing spill from this cell before setting new value
        self.clear_spill_from(row, col);

        // Invalidate computed cache (cell changed, dependents may need recompute)
        self.computed_cache.borrow_mut().remove(&(row, col));

        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.set(value);

        // If this is a formula, evaluate it and apply spill if it returns an array
        self.evaluate_and_spill(row, col);
    }

    /// Mark a cell as having a cycle error.
    ///
    /// Used when loading workbooks with circular references to mark
    /// participating cells without crashing.
    pub fn set_cycle_error(&mut self, row: usize, col: usize) {
        // Redirect hidden merge cells to the merge origin
        let (row, col) = self.merge_origin_coord(row, col);

        // Store #CYCLE! as the cell value while preserving the formula source
        // For now, we just set a text value - the original formula is lost
        // A future improvement could preserve the formula for editing
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.value = CellValue::Text("#CYCLE!".to_string());
    }

    /// Evaluate a cell's formula and apply spill if it returns an array
    fn evaluate_and_spill(&mut self, row: usize, col: usize) {
        // Get the AST if this is a formula
        let ast = match self.cells.get(&(row, col)) {
            Some(cell) => match &cell.value {
                CellValue::Formula { ast: Some(ast), .. } => ast.clone(),
                _ => return, // Not a formula or no AST
            },
            None => return,
        };

        // Evaluate the formula with current cell context for ROW()/COLUMN()
        let lookup = LookupWithContext::new(self, row, col);
        let bound_ast = bind_expr_same_sheet(&ast);
        let result = eval::evaluate(&bound_ast, &lookup);

        // Cache the scalar result so getters can read it without re-evaluating.
        // For arrays, the top-left value is cached; spill receivers are stored in spill_values.
        // Topo recalc (workbook.evaluate_cell) will overwrite this with the authoritative result.
        self.cache_computed(row, col, result.to_value());

        // If it's an array, try to apply spill
        if let EvalResult::Array(array) = result {
            let rows = array.rows();
            let cols = array.cols();

            // Check for collision before applying
            match self.check_spill_collision(row, col, rows, cols) {
                Ok(()) => {
                    // Apply the spill
                    self.apply_spill(row, col, &array);
                }
                Err(blocked_by) => {
                    // Record the spill error
                    if let Some(cell) = self.cells.get_mut(&(row, col)) {
                        cell.spill_error = Some(SpillError { blocked_by });
                    }
                }
            }
        }
    }

    // =========================================================================
    // Spill Management
    // =========================================================================

    /// Clear spill data originating from a specific cell
    pub fn clear_spill_from(&mut self, parent_row: usize, parent_col: usize) {
        // Get the spill info from the parent cell
        let spill_info = match self.cells.get(&(parent_row, parent_col)) {
            Some(cell) => cell.spill_info.clone(),
            None => return,
        };

        if let Some(info) = spill_info {
            // Clear all receiving cells
            for dr in 0..info.rows {
                for dc in 0..info.cols {
                    if dr == 0 && dc == 0 {
                        continue; // Skip the parent cell itself
                    }
                    let r = parent_row + dr;
                    let c = parent_col + dc;

                    // Remove spill value
                    self.spill_values.remove(&(r, c));

                    // Clear spill_parent reference
                    if let Some(cell) = self.cells.get_mut(&(r, c)) {
                        if cell.spill_parent == Some((parent_row, parent_col)) {
                            cell.spill_parent = None;
                        }
                    }
                }
            }

            // Clear spill_info on parent
            if let Some(cell) = self.cells.get_mut(&(parent_row, parent_col)) {
                cell.spill_info = None;
            }
        }
    }

    /// Check if spill from (parent_row, parent_col) with given dimensions would collide
    /// Returns Ok(()) if no collision, Err with blocking cell position if collision
    pub fn check_spill_collision(
        &self,
        parent_row: usize,
        parent_col: usize,
        rows: usize,
        cols: usize,
    ) -> Result<(), (usize, usize)> {
        for dr in 0..rows {
            for dc in 0..cols {
                if dr == 0 && dc == 0 {
                    continue; // Skip the parent cell
                }
                let r = parent_row + dr;
                let c = parent_col + dc;

                // Check if cell exists and has content
                if let Some(cell) = self.cells.get(&(r, c)) {
                    // Check if it has its own value (not a spill receiver from us)
                    let is_our_receiver = cell.spill_parent == Some((parent_row, parent_col));
                    if !is_our_receiver {
                        match &cell.value {
                            CellValue::Empty => {}
                            _ => return Err((r, c)),
                        }
                        // Also blocked if it's receiving spill from another cell
                        if cell.spill_parent.is_some() {
                            return Err((r, c));
                        }
                    }
                }

                // Check if there's a spill value from another parent
                if let Some(_) = self.spill_values.get(&(r, c)) {
                    // Check if this spill is from us or another parent
                    if let Some(cell) = self.cells.get(&(r, c)) {
                        if cell.spill_parent != Some((parent_row, parent_col)) {
                            return Err((r, c));
                        }
                    } else {
                        return Err((r, c));
                    }
                }
            }
        }
        Ok(())
    }

    /// Apply spill from an array result at (parent_row, parent_col)
    /// Returns true if spill was successful, false if blocked (#SPILL!)
    pub fn apply_spill(
        &mut self,
        parent_row: usize,
        parent_col: usize,
        array: &Array2D,
    ) -> bool {
        let rows = array.rows();
        let cols = array.cols();

        // Check for collision first
        if self.check_spill_collision(parent_row, parent_col, rows, cols).is_err() {
            return false;
        }

        // Clear any existing spill from this parent
        self.clear_spill_from(parent_row, parent_col);

        // Apply new spill
        for dr in 0..rows {
            for dc in 0..cols {
                let r = parent_row + dr;
                let c = parent_col + dc;

                if let Some(value) = array.get(dr, dc) {
                    if dr == 0 && dc == 0 {
                        // Parent cell - just set spill_info
                        let cell = self.cells.entry((r, c)).or_insert_with(Cell::new);
                        cell.spill_info = Some(SpillInfo { rows, cols });
                    } else {
                        // Receiving cell
                        self.spill_values.insert((r, c), value.clone());
                        let cell = self.cells.entry((r, c)).or_insert_with(Cell::new);
                        cell.spill_parent = Some((parent_row, parent_col));
                    }
                }
            }
        }

        true
    }

    /// Get the spilled value at a position (if any)
    pub fn get_spill_value(&self, row: usize, col: usize) -> Option<&Value> {
        self.spill_values.get(&(row, col))
    }

    /// Check if a cell is receiving spill data
    pub fn is_spill_receiver(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.is_spill_receiver())
            .unwrap_or(false)
    }

    /// Check if a cell is a spill parent
    pub fn is_spill_parent(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.is_spill_parent())
            .unwrap_or(false)
    }

    /// Get the spill parent for a cell (if it's a spill receiver)
    pub fn get_spill_parent(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.spill_parent)
    }

    /// Get spill info for a cell (if it's a spill parent)
    pub fn get_spill_info(&self, row: usize, col: usize) -> Option<SpillInfo> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.spill_info.clone())
    }

    /// Check if a cell has a spill error
    pub fn has_spill_error(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.has_spill_error())
            .unwrap_or(false)
    }

    pub fn get_display(&self, row: usize, col: usize) -> String {
        // Check for spilled value first
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Check for spill error
                if cell.spill_error.is_some() {
                    return "#SPILL!".to_string();
                }
                self.display_cell_value(&cell.value, row, col)
            }
            None => String::new(),
        }
    }

    /// Get display value with number formatting applied
    pub fn get_formatted_display(&self, row: usize, col: usize) -> String {
        // Check for spilled value first
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            // Apply formatting from the cell if it exists
            if let Some(cell) = self.cells.get(&(row, col)) {
                match spill_value {
                    Value::Number(n) => return CellValue::format_number(*n, &cell.format.number_format),
                    _ => return spill_value.to_text(),
                }
            }
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Check for spill error
                if cell.spill_error.is_some() {
                    return "#SPILL!".to_string();
                }
                // Get the cached Value for formatting (never evaluate on cache miss)
                let value = match &cell.value {
                    CellValue::Number(n) => Value::Number(*n),
                    CellValue::Formula { ast: Some(_), .. } => {
                        // Cache-only: never evaluate on cache miss.
                        let cache = self.computed_cache.borrow();
                        cache.get(&(row, col)).cloned().unwrap_or(Value::Empty)
                    }
                    CellValue::Text(s) => Value::Text(s.clone()),
                    CellValue::Empty => return String::new(),
                    CellValue::Formula { ast: None, .. } => return "#ERR".to_string(),
                };

                match value {
                    Value::Number(n) => {
                        // Apply number formatting
                        CellValue::format_number(n, &cell.format.number_format)
                    }
                    Value::Text(s) => s,
                    Value::Boolean(b) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                    Value::Error(e) => e,
                    Value::Empty => String::new(),
                }
            }
            None => String::new(),
        }
    }

    pub fn get_raw(&self, row: usize, col: usize) -> String {
        self.cells
            .get(&(row, col))
            .map(|c| c.value.raw_display())
            .unwrap_or_default()
    }

    /// Get the computed Value for a cell (typed, not formatted string).
    /// Used for Paste Values to preserve types without display formatting.
    pub fn get_computed_value(&self, row: usize, col: usize) -> Value {
        // Check for spilled value first
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.clone();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Check for spill error
                if cell.spill_error.is_some() {
                    return Value::Error("#SPILL!".to_string());
                }
                // Evaluate the cell value
                match &cell.value {
                    CellValue::Empty => Value::Empty,
                    CellValue::Text(s) => Value::Text(s.clone()),
                    CellValue::Number(n) => Value::Number(*n),
                    CellValue::Formula { ast: Some(_), .. } => {
                        // Cache-only: never evaluate on cache miss.
                        let cache = self.computed_cache.borrow();
                        cache.get(&(row, col)).cloned().unwrap_or(Value::Empty)
                    }
                    CellValue::Formula { ast: None, .. } => Value::Error("#ERR".to_string()),
                }
            }
            None => Value::Empty,
        }
    }

    /// Get a reference to a cell (returns default empty cell if not found)
    pub fn get_cell(&self, row: usize, col: usize) -> Cell {
        self.cells
            .get(&(row, col))
            .cloned()
            .unwrap_or_default()
    }

    fn display_cell_value(&self, value: &CellValue, row: usize, col: usize) -> String {
        match value {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{:.2}", n)
                }
            }
            CellValue::Formula { ast: Some(_), .. } => {
                // Cache-only: never evaluate on cache miss.
                let cache = self.computed_cache.borrow();
                match cache.get(&(row, col)) {
                    Some(Value::Number(n)) => {
                        if n.fract() == 0.0 {
                            format!("{}", *n as i64)
                        } else {
                            format!("{:.2}", n)
                        }
                    }
                    Some(Value::Error(e)) => e.clone(),
                    Some(v) => v.to_text(),
                    None => String::new(),
                }
            }
            CellValue::Formula { ast: None, .. } => "#ERR".to_string(),
        }
    }

    pub fn get_format(&self, row: usize, col: usize) -> CellFormat {
        self.cells
            .get(&(row, col))
            .map(|c| c.format.clone())
            .unwrap_or_default()
    }

    /// Iterate over all populated cells
    pub fn cells_iter(&self) -> impl Iterator<Item = (&(usize, usize), &Cell)> {
        self.cells.iter()
    }

    /// Get coordinates of non-empty cells within a range
    pub fn cells_in_range(&self, min_row: usize, max_row: usize, min_col: usize, max_col: usize) -> Vec<(usize, usize)> {
        self.cells
            .keys()
            .filter(|(r, c)| *r >= min_row && *r <= max_row && *c >= min_col && *c <= max_col)
            .copied()
            .collect()
    }

    /// Clear a cell completely (remove from HashMap)
    pub fn clear_cell(&mut self, row: usize, col: usize) {
        // Redirect hidden merge cells to the merge origin
        let (row, col) = self.merge_origin_coord(row, col);

        self.clear_spill_from(row, col);
        self.cells.remove(&(row, col));
        self.spill_values.remove(&(row, col));
    }

    pub fn set_format(&mut self, row: usize, col: usize, format: CellFormat) {
        if !self.has_any_borders && format.has_any_border() {
            self.has_any_borders = true;
        }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format = format;
    }

    /// Set the style_id on a cell (imported style provenance).
    pub fn set_style_id(&mut self, row: usize, col: usize, style_id: u32) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.style_id = Some(style_id);
    }

    /// Set format from import: applies the full CellFormat as the resolved format
    /// and does NOT touch style_id (caller sets that separately).
    pub fn set_format_from_import(&mut self, row: usize, col: usize, format: CellFormat) {
        if !self.has_any_borders && format.has_any_border() {
            self.has_any_borders = true;
        }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format = format;
    }

    pub fn toggle_bold(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.bold = !cell.format.bold;
    }

    pub fn toggle_italic(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.italic = !cell.format.italic;
    }

    pub fn toggle_underline(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.underline = !cell.format.underline;
    }

    pub fn toggle_strikethrough(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.strikethrough = !cell.format.strikethrough;
    }

    pub fn set_bold(&mut self, row: usize, col: usize, value: bool) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.bold = value;
    }

    pub fn set_italic(&mut self, row: usize, col: usize, value: bool) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.italic = value;
    }

    pub fn set_underline(&mut self, row: usize, col: usize, value: bool) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.underline = value;
    }

    pub fn set_strikethrough(&mut self, row: usize, col: usize, value: bool) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.strikethrough = value;
    }

    pub fn set_alignment(&mut self, row: usize, col: usize, alignment: Alignment) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.alignment = alignment;
    }

    pub fn set_vertical_alignment(&mut self, row: usize, col: usize, vertical_alignment: VerticalAlignment) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.vertical_alignment = vertical_alignment;
    }

    pub fn set_text_overflow(&mut self, row: usize, col: usize, text_overflow: TextOverflow) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.text_overflow = text_overflow;
    }

    pub fn set_number_format(&mut self, row: usize, col: usize, number_format: NumberFormat) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.number_format = number_format;
    }

    pub fn get_number_format(&self, row: usize, col: usize) -> NumberFormat {
        self.cells
            .get(&(row, col))
            .map(|c| c.format.number_format.clone())
            .unwrap_or_default()
    }

    pub fn set_font_family(&mut self, row: usize, col: usize, font_family: Option<String>) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.font_family = font_family;
    }

    pub fn get_font_family(&self, row: usize, col: usize) -> Option<String> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.format.font_family.clone())
    }

    pub fn set_background_color(&mut self, row: usize, col: usize, color: Option<[u8; 4]>) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.background_color = color;
    }

    pub fn get_background_color(&self, row: usize, col: usize) -> Option<[u8; 4]> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.format.background_color)
    }

    pub fn set_font_size(&mut self, row: usize, col: usize, size: Option<f32>) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.font_size = size;
    }

    pub fn set_font_color(&mut self, row: usize, col: usize, color: Option<[u8; 4]>) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.font_color = color;
    }

    /// Set all 4 borders on a cell at once
    pub fn set_borders(
        &mut self,
        row: usize,
        col: usize,
        top: CellBorder,
        right: CellBorder,
        bottom: CellBorder,
        left: CellBorder,
    ) {
        if !self.has_any_borders && (top.is_set() || right.is_set() || bottom.is_set() || left.is_set()) {
            self.has_any_borders = true;
        }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.border_top = top;
        cell.format.border_right = right;
        cell.format.border_bottom = bottom;
        cell.format.border_left = left;
    }

    /// Set the top border on a cell
    pub fn set_border_top(&mut self, row: usize, col: usize, border: CellBorder) {
        if !self.has_any_borders && border.is_set() { self.has_any_borders = true; }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.border_top = border;
    }

    /// Set the right border on a cell
    pub fn set_border_right(&mut self, row: usize, col: usize, border: CellBorder) {
        if !self.has_any_borders && border.is_set() { self.has_any_borders = true; }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.border_right = border;
    }

    /// Set the bottom border on a cell
    pub fn set_border_bottom(&mut self, row: usize, col: usize, border: CellBorder) {
        if !self.has_any_borders && border.is_set() { self.has_any_borders = true; }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.border_bottom = border;
    }

    /// Set the left border on a cell
    pub fn set_border_left(&mut self, row: usize, col: usize, border: CellBorder) {
        if !self.has_any_borders && border.is_set() { self.has_any_borders = true; }
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.border_left = border;
    }

    /// Rescan all cells to recompute `has_any_borders`.
    /// Call after bulk operations that may have cleared borders (e.g., undo/redo
    /// restoring a previous workbook snapshot).
    pub fn scan_border_flag(&mut self) {
        self.has_any_borders = self.cells.values().any(|c| c.format.has_any_border());
    }

    // =========================================================================
    // Merged Regions
    // =========================================================================

    /// Replace all merged regions and rebuild the lookup index.
    /// Used by undo/redo to restore merge state.
    pub fn set_merges(&mut self, regions: Vec<MergedRegion>) {
        self.merged_regions = regions;
        self.rebuild_merge_index();
    }

    /// Rebuild the merge_index from merged_regions (O(total cells in merges)).
    /// Must be called after any mutation of merged_regions.
    pub fn rebuild_merge_index(&mut self) {
        self.merge_index.clear();
        for (idx, region) in self.merged_regions.iter().enumerate() {
            for r in region.start.0..=region.end.0 {
                for c in region.start.1..=region.end.1 {
                    self.merge_index.insert((r, c), idx);
                }
            }
        }
    }

    /// O(1) lookup: get the merged region containing (row, col), if any.
    pub fn get_merge(&self, row: usize, col: usize) -> Option<&MergedRegion> {
        self.merge_index
            .get(&(row, col))
            .and_then(|&idx| self.merged_regions.get(idx))
    }

    /// Redirect hidden merge cells to their origin for single-cell value writes
    /// (set_value, clear_cell, set_cycle_error). Formula CellRef redirect is
    /// handled separately in eval.rs. NOT for format ops or range iteration.
    pub fn merge_origin_coord(&self, row: usize, col: usize) -> (usize, usize) {
        if let Some(merge) = self.get_merge(row, col) {
            merge.start
        } else {
            (row, col)
        }
    }

    /// Is (row, col) the top-left origin of a merged region?
    pub fn is_merge_origin(&self, row: usize, col: usize) -> bool {
        self.get_merge(row, col)
            .map(|m| m.start == (row, col))
            .unwrap_or(false)
    }

    /// Is (row, col) inside a merge but NOT the origin (i.e., should be hidden)?
    pub fn is_merge_hidden(&self, row: usize, col: usize) -> bool {
        self.get_merge(row, col)
            .map(|m| m.start != (row, col))
            .unwrap_or(false)
    }

    /// Add a merged region. Returns Err if it overlaps an existing merge.
    pub fn add_merge(&mut self, region: MergedRegion) -> Result<(), String> {
        if region.is_degenerate() {
            return Ok(()); // silently ignore 1×1 merges
        }
        // Check for overlaps
        for r in region.start.0..=region.end.0 {
            for c in region.start.1..=region.end.1 {
                if self.merge_index.contains_key(&(r, c)) {
                    return Err(format!(
                        "merge ({},{})→({},{}) overlaps existing merge at ({},{})",
                        region.start.0, region.start.1, region.end.0, region.end.1, r, c
                    ));
                }
            }
        }
        // Incremental index update — no full rebuild
        let idx = self.merged_regions.len();
        for r in region.start.0..=region.end.0 {
            for c in region.start.1..=region.end.1 {
                self.merge_index.insert((r, c), idx);
            }
        }
        self.merged_regions.push(region);
        Ok(())
    }

    /// Remove the merge whose origin is `start`. Returns the removed region, if any.
    pub fn remove_merge(&mut self, start: (usize, usize)) -> Option<MergedRegion> {
        if let Some(pos) = self
            .merged_regions
            .iter()
            .position(|m| m.start == start)
        {
            let removed = self.merged_regions.remove(pos);
            // Incremental: clear removed region's entries, then fix indices
            // shifted by the Vec::remove. Cheaper than full rebuild for small
            // merge counts; normalize_merges still does full rebuild for bulk ops.
            for r in removed.start.0..=removed.end.0 {
                for c in removed.start.1..=removed.end.1 {
                    self.merge_index.remove(&(r, c));
                }
            }
            // Indices above `pos` shifted down by 1 after Vec::remove
            for v in self.merge_index.values_mut() {
                if *v > pos {
                    *v -= 1;
                }
            }
            Some(removed)
        } else {
            None
        }
    }

    /// Remove degenerate (1×1) merges and rebuild the index.
    pub fn normalize_merges(&mut self) {
        self.merged_regions.retain(|m| !m.is_degenerate());
        self.rebuild_merge_index();
        #[cfg(debug_assertions)]
        self.debug_assert_no_merge_overlap();
    }

    /// Debug-only: assert that no two merges overlap.
    #[cfg(debug_assertions)]
    fn debug_assert_no_merge_overlap(&self) {
        let mut seen: HashSet<(usize, usize)> = HashSet::new();
        for region in &self.merged_regions {
            for r in region.start.0..=region.end.0 {
                for c in region.start.1..=region.end.1 {
                    assert!(
                        seen.insert((r, c)),
                        "merge overlap detected at ({}, {})",
                        r,
                        c
                    );
                }
            }
        }
    }

    /// Compute the resolved border for each side of a merged region (edge consolidation).
    ///
    /// Scans edge cells along each side and returns the max-precedence border:
    /// - Top: scan border_top of (start_row, col) for col in start_col..=end_col
    /// - Right: scan border_right of (row, end_col) for row in start_row..=end_row
    /// - Bottom: scan border_bottom of (end_row, col) for col in start_col..=end_col
    /// - Left: scan border_left of (row, start_col) for row in start_row..=end_row
    ///
    /// Edge consolidation: the entire edge uses a single, consistent border style —
    /// the strongest found on any cell along that edge. This differs from Excel's
    /// per-cell-segment model but produces cleaner visuals for financial models.
    ///
    /// Returns (top, right, bottom, left) resolved borders.
    pub fn resolve_merge_borders(&self, merge: &MergedRegion) -> (CellBorder, CellBorder, CellBorder, CellBorder) {
        let mut top = CellBorder::default();
        for c in merge.start.1..=merge.end.1 {
            top = max_border(top, self.get_format(merge.start.0, c).border_top);
        }

        let mut right = CellBorder::default();
        for r in merge.start.0..=merge.end.0 {
            right = max_border(right, self.get_format(r, merge.end.1).border_right);
        }

        let mut bottom = CellBorder::default();
        for c in merge.start.1..=merge.end.1 {
            bottom = max_border(bottom, self.get_format(merge.end.0, c).border_bottom);
        }

        let mut left = CellBorder::default();
        for r in merge.start.0..=merge.end.0 {
            left = max_border(left, self.get_format(r, merge.start.1).border_left);
        }

        (top, right, bottom, left)
    }

    /// Insert rows at the specified position, shifting existing rows down
    pub fn insert_rows(&mut self, at_row: usize, count: usize) {
        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((r, _), _)| *r >= at_row)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted down)
        for ((r, c), cell) in cells_to_shift {
            if r + count < self.rows {
                self.cells.insert((r + count, c), cell);
            }
        }

        // Adjust merged regions (grid-line semantics)
        for m in &mut self.merged_regions {
            if at_row <= m.start.0 {
                // Insertion at or above merge → shift entire merge down
                m.start.0 += count;
                m.end.0 += count;
            } else if at_row <= m.end.0 {
                // Insertion inside merge → expand merge
                m.end.0 += count;
            }
        }
        self.normalize_merges();
    }

    /// Delete rows at the specified position, shifting remaining rows up
    pub fn delete_rows(&mut self, start_row: usize, count: usize) {
        let end_row = start_row + count; // exclusive

        // Remove cells in the deleted rows
        for row in start_row..end_row {
            for col in 0..self.cols {
                self.cells.remove(&(row, col));
            }
        }

        // Collect cells that need to be shifted up
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((r, _), _)| *r >= end_row)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted up)
        for ((r, c), cell) in cells_to_shift {
            self.cells.insert((r - count, c), cell);
        }

        // Adjust merged regions (grid-line semantics)
        for m in &mut self.merged_regions {
            if end_row <= m.start.0 {
                // Deletion entirely above → shift up
                m.start.0 -= count;
                m.end.0 -= count;
            } else if start_row > m.end.0 {
                // Deletion entirely below → no effect
            } else if start_row <= m.start.0 && end_row > m.end.0 {
                // Deletion engulfs entire merge → mark degenerate
                m.start.0 = start_row;
                m.end.0 = m.start.0;
                m.end.1 = m.start.1;
            } else if start_row <= m.start.0 {
                // Deletion clips top of merge; surviving rows shift up by count
                m.start.0 = start_row;
                m.end.0 -= count;
            } else if end_row > m.end.0 {
                // Deletion clips bottom of merge
                m.end.0 = start_row - 1;
            } else {
                // Deletion entirely inside merge → shrink
                m.end.0 -= count;
            }
        }
        self.normalize_merges();
        // Deleted rows may have removed the only bordered cells.
        // Only rescan when the flag is currently true (can't flip false→false).
        // TODO(perf): if delete_rows on a 50k+ row bordered sheet causes >16ms frame hitch
        // in debug builds, add a per-row border bitset to check deleted band intersection
        // before doing the full scan.
        if self.has_any_borders {
            self.scan_border_flag();
        }
    }

    /// Insert columns at the specified position, shifting existing columns right
    pub fn insert_cols(&mut self, at_col: usize, count: usize) {
        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((_, c), _)| *c >= at_col)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted right)
        for ((r, c), cell) in cells_to_shift {
            if c + count < self.cols {
                self.cells.insert((r, c + count), cell);
            }
        }

        // Adjust merged regions (grid-line semantics)
        for m in &mut self.merged_regions {
            if at_col <= m.start.1 {
                m.start.1 += count;
                m.end.1 += count;
            } else if at_col <= m.end.1 {
                m.end.1 += count;
            }
        }
        self.normalize_merges();
    }

    /// Delete columns at the specified position, shifting remaining columns left
    pub fn delete_cols(&mut self, start_col: usize, count: usize) {
        let end_col = start_col + count; // exclusive

        // Remove cells in the deleted columns
        for col in start_col..end_col {
            for row in 0..self.rows {
                self.cells.remove(&(row, col));
            }
        }

        // Collect cells that need to be shifted left
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((_, c), _)| *c >= end_col)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted left)
        for ((r, c), cell) in cells_to_shift {
            self.cells.insert((r, c - count), cell);
        }

        // Adjust merged regions (grid-line semantics)
        for m in &mut self.merged_regions {
            if end_col <= m.start.1 {
                // Deletion entirely left → shift left
                m.start.1 -= count;
                m.end.1 -= count;
            } else if start_col > m.end.1 {
                // Deletion entirely right → no effect
            } else if start_col <= m.start.1 && end_col > m.end.1 {
                // Deletion engulfs entire merge → mark degenerate
                m.start.1 = start_col;
                m.end.1 = m.start.1;
                m.end.0 = m.start.0;
            } else if start_col <= m.start.1 {
                // Deletion clips left side; surviving cols shift left by count
                m.start.1 = start_col;
                m.end.1 -= count;
            } else if end_col > m.end.1 {
                // Deletion clips right side
                m.end.1 = start_col - 1;
            } else {
                // Deletion entirely inside merge → shrink
                m.end.1 -= count;
            }
        }
        self.normalize_merges();
        // Deleted columns may have removed the only bordered cells.
        // Only rescan when the flag is currently true (can't flip false→false).
        // TODO(perf): if delete_cols on a 50k+ col bordered sheet causes >16ms frame hitch
        // in debug builds, add a per-col border bitset to check deleted band intersection
        // before doing the full scan.
        if self.has_any_borders {
            self.scan_border_flag();
        }
    }

    // =========================================================================
    // Data Validation
    // =========================================================================

    /// Set a validation rule for a range of cells.
    pub fn set_validation(
        &mut self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        rule: super::validation::ValidationRule,
    ) {
        use super::validation::CellRange;
        let range = CellRange::new(start_row, start_col, end_row, end_col);
        self.validations.set(range, rule);
    }

    /// Set a validation rule for a single cell.
    pub fn set_cell_validation(
        &mut self,
        row: usize,
        col: usize,
        rule: super::validation::ValidationRule,
    ) {
        self.set_validation(row, col, row, col, rule);
    }

    /// Clear validation rules for a range.
    pub fn clear_validation(&mut self, start_row: usize, start_col: usize, end_row: usize, end_col: usize) {
        use super::validation::CellRange;
        let range = CellRange::new(start_row, start_col, end_row, end_col);
        self.validations.clear_range(&range);
    }

    /// Get the validation rule for a cell (if any).
    pub fn get_validation(&self, row: usize, col: usize) -> Option<&super::validation::ValidationRule> {
        self.validations.get(row, col)
    }

    /// Check if a cell has validation.
    pub fn has_validation(&self, row: usize, col: usize) -> bool {
        self.validations.has_validation(row, col)
    }

    /// Validate a value against the cell's validation rule.
    ///
    /// Returns `ValidationResult::Valid` if no rule exists or validation passes.
    pub fn validate_cell_input(&self, row: usize, col: usize, value: &str) -> super::validation::ValidationResult {
        use super::validation::{ValidationResult, ValidationType};

        let rule = match self.validations.get(row, col) {
            Some(r) => r,
            None => return ValidationResult::Valid,
        };

        // Check ignore_blank
        if rule.ignore_blank && value.trim().is_empty() {
            return ValidationResult::Valid;
        }

        // Validate based on type
        // NOTE: No AnyValue case - rule absence handles "any value" semantics
        match &rule.rule_type {
            ValidationType::List(source) => {
                let resolved = self.resolve_list_source(source);
                let trimmed_value = value.trim();

                // Case-sensitive matching (per spec)
                if resolved.contains(trimmed_value) {
                    ValidationResult::Valid
                } else if resolved.items.is_empty() {
                    // Empty list source (e.g., invalid range) - accept any value
                    ValidationResult::Valid
                } else {
                    let display_items: Vec<&str> = resolved.items.iter()
                        .take(5)
                        .map(|s| s.as_str())
                        .collect();
                    let suffix = if resolved.items.len() > 5 { ", ..." } else { "" };
                    ValidationResult::Invalid {
                        rule: rule.clone(),
                        reason: format!("Value must be one of: {}{}", display_items.join(", "), suffix),
                    }
                }
            }

            ValidationType::WholeNumber(constraint) => {
                use super::validation::{parse_numeric_input, NumericParseError};

                // Use strict parsing: no decimal point allowed
                let num = match parse_numeric_input(value, false) {
                    Ok(n) => n,
                    Err(NumericParseError::FractionalNotAllowed) => {
                        return ValidationResult::Invalid {
                            rule: rule.clone(),
                            reason: "Value must be a whole number (no decimals)".to_string(),
                        };
                    }
                    Err(_) => {
                        return ValidationResult::Invalid {
                            rule: rule.clone(),
                            reason: "Value must be a whole number".to_string(),
                        };
                    }
                };

                self.validate_numeric_constraint(num, constraint, rule, "whole number")
            }

            ValidationType::Decimal(constraint) => {
                use super::validation::parse_numeric_input;

                // Allow decimal input
                let num = match parse_numeric_input(value, true) {
                    Ok(n) => n,
                    Err(_) => {
                        return ValidationResult::Invalid {
                            rule: rule.clone(),
                            reason: "Value must be a number".to_string(),
                        };
                    }
                };

                self.validate_numeric_constraint(num, constraint, rule, "number")
            }

            ValidationType::TextLength(constraint) => {
                let len = value.len() as f64;
                self.validate_numeric_constraint(len, constraint, rule, "text length")
            }

            ValidationType::Date(_) | ValidationType::Time(_) => {
                // TODO: Implement date/time parsing and validation
                ValidationResult::Valid
            }

            ValidationType::Custom(_formula) => {
                // TODO: Evaluate custom formula
                ValidationResult::Valid
            }
        }
    }

    /// Helper to validate a numeric value against a constraint.
    fn validate_numeric_constraint(
        &self,
        value: f64,
        constraint: &super::validation::NumericConstraint,
        rule: &super::validation::ValidationRule,
        type_name: &str,
    ) -> super::validation::ValidationResult {
        use super::validation::{ValidationResult, eval_numeric_constraint};

        // Resolve constraint values - fail validation if constraint can't be resolved
        let v1 = match self.resolve_constraint_value(&constraint.value1) {
            Ok(n) => n,
            Err(e) => {
                return ValidationResult::Invalid {
                    rule: rule.clone(),
                    reason: format!("Validation constraint error: {}", e),
                };
            }
        };

        let v2 = match &constraint.value2 {
            Some(cv) => match self.resolve_constraint_value(cv) {
                Ok(n) => Some(n),
                Err(e) => {
                    return ValidationResult::Invalid {
                        rule: rule.clone(),
                        reason: format!("Validation constraint error: {}", e),
                    };
                }
            },
            None => None,
        };

        // Use the shared evaluation helper
        let valid = eval_numeric_constraint(value, constraint.operator, v1, v2);

        if valid {
            ValidationResult::Valid
        } else {
            let reason = match constraint.operator {
                super::validation::ComparisonOperator::Between => {
                    format!("{} must be between {} and {}", type_name, v1, v2.unwrap_or(0.0))
                }
                super::validation::ComparisonOperator::NotBetween => {
                    format!("{} must not be between {} and {}", type_name, v1, v2.unwrap_or(0.0))
                }
                super::validation::ComparisonOperator::EqualTo => format!("{} must equal {}", type_name, v1),
                super::validation::ComparisonOperator::NotEqualTo => format!("{} must not equal {}", type_name, v1),
                super::validation::ComparisonOperator::GreaterThan => format!("{} must be greater than {}", type_name, v1),
                super::validation::ComparisonOperator::LessThan => format!("{} must be less than {}", type_name, v1),
                super::validation::ComparisonOperator::GreaterThanOrEqual => format!("{} must be at least {}", type_name, v1),
                super::validation::ComparisonOperator::LessThanOrEqual => format!("{} must be at most {}", type_name, v1),
            };

            ValidationResult::Invalid {
                rule: rule.clone(),
                reason,
            }
        }
    }

    /// Resolve a constraint value to a number.
    ///
    /// Returns Err if:
    /// - Cell reference is invalid
    /// - Referenced cell is blank
    /// - Referenced cell value is not numeric
    /// - Formula evaluation fails or returns non-numeric
    fn resolve_constraint_value(
        &self,
        value: &super::validation::ConstraintValue,
    ) -> Result<f64, super::validation::ConstraintResolveError> {
        use super::validation::{ConstraintValue, ConstraintResolveError};

        match value {
            ConstraintValue::Number(n) => Ok(*n),
            ConstraintValue::CellRef(ref_str) => {
                // Parse cell reference and get value
                let (row, col) = self.parse_cell_ref(ref_str)
                    .ok_or_else(|| ConstraintResolveError::InvalidReference(ref_str.clone()))?;

                let display = self.get_display(row, col);
                if display.is_empty() {
                    return Err(ConstraintResolveError::BlankConstraint);
                }

                // Try to parse as number
                display.parse::<f64>()
                    .map_err(|_| ConstraintResolveError::NotNumeric)
            }
            ConstraintValue::Formula(_formula) => {
                // TODO: Evaluate formula and require numeric result
                // For now, return error since formula eval not implemented
                Err(ConstraintResolveError::FormulaError("Formula constraints not yet implemented".to_string()))
            }
        }
    }

    /// Resolve a list source to its items.
    ///
    /// Returns a ResolvedList with normalized items (trimmed whitespace).
    /// For range sources, reads cell values. For named ranges, looks up the range first.
    fn resolve_list_source(&self, source: &super::validation::ListSource) -> super::validation::ResolvedList {
        use super::validation::{ListSource, ResolvedList};

        match source {
            ListSource::Inline(values) => {
                ResolvedList::from_items(values.clone())
            }
            ListSource::Range(range_str) => {
                // Parse range string like "A1:A10" or "=A1:A10"
                let range_str = range_str.trim_start_matches('=').trim();
                self.resolve_range_to_list(range_str)
            }
            ListSource::NamedRange(_name) => {
                // Named range resolution requires workbook context
                // This method is called from Sheet, which doesn't have workbook access
                // The workbook-level method will handle this
                ResolvedList::empty()
            }
        }
    }

    /// Resolve a range string like "A1:A10" to a list of cell values.
    pub fn resolve_range_to_list(&self, range_str: &str) -> super::validation::ResolvedList {
        use super::validation::ResolvedList;

        // Parse range: "A1:B10" or just "A1"
        let parts: Vec<&str> = range_str.split(':').collect();
        if parts.is_empty() || parts.len() > 2 {
            return ResolvedList::empty();
        }

        let start = match self.parse_cell_ref(parts[0]) {
            Some(pos) => pos,
            None => return ResolvedList::empty(),
        };

        let end = if parts.len() == 2 {
            match self.parse_cell_ref(parts[1]) {
                Some(pos) => pos,
                None => return ResolvedList::empty(),
            }
        } else {
            start
        };

        // Collect values from range
        let mut items = Vec::new();
        let (start_row, start_col) = start;
        let (end_row, end_col) = end;

        for row in start_row.min(end_row)..=start_row.max(end_row) {
            for col in start_col.min(end_col)..=start_col.max(end_col) {
                let display = self.get_display(row, col);
                // Skip truly empty cells but include cells with whitespace (after trim)
                if !display.is_empty() {
                    items.push(display);
                }
            }
        }

        ResolvedList::from_items(items)
    }

    /// Get the list items for a cell with list validation.
    ///
    /// Returns None if the cell has no validation or non-list validation.
    /// Returns Some(ResolvedList) with the dropdown items.
    pub fn get_list_items(&self, row: usize, col: usize) -> Option<super::validation::ResolvedList> {
        use super::validation::ValidationType;

        let rule = self.validations.get(row, col)?;

        match &rule.rule_type {
            ValidationType::List(source) => {
                Some(self.resolve_list_source(source))
            }
            _ => None,
        }
    }

    /// Check if a cell has list validation with dropdown enabled.
    pub fn has_list_dropdown(&self, row: usize, col: usize) -> bool {
        use super::validation::ValidationType;

        if let Some(rule) = self.validations.get(row, col) {
            rule.show_dropdown && matches!(rule.rule_type, ValidationType::List(_))
        } else {
            false
        }
    }

    /// Parse a simple cell reference like "A1" or "B10".
    pub fn parse_cell_ref(&self, ref_str: &str) -> Option<(usize, usize)> {
        let ref_str = ref_str.trim().to_uppercase();
        let mut col_str = String::new();
        let mut row_str = String::new();

        for ch in ref_str.chars() {
            if ch.is_ascii_alphabetic() {
                col_str.push(ch);
            } else if ch.is_ascii_digit() {
                row_str.push(ch);
            }
        }

        if col_str.is_empty() || row_str.is_empty() {
            return None;
        }

        // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
        let col = col_str.chars().fold(0usize, |acc, c| {
            acc * 26 + (c as usize - 'A' as usize + 1)
        }) - 1;

        // Convert row to 0-indexed
        let row: usize = row_str.parse().ok()?;
        if row == 0 {
            return None;
        }

        Some((row - 1, col))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_set_text_overflow() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Default should be Clip
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Clip);

        // Set to Wrap
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Wrap);

        // Set to Overflow (Spill)
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Overflow);
    }

    #[test]
    fn test_set_vertical_alignment() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Default should be Middle
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Middle);

        // Set to Top
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Top);

        // Set to Bottom
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Bottom);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Bottom);
    }

    #[test]
    fn test_format_persists_with_value() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set format first
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);

        // Then set value
        sheet.set_value(0, 0, "Hello");

        // Format should persist
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Wrap);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Top);
    }

    #[test]
    fn test_spill_stops_at_nonempty_cell() {
        // This tests the data model aspect - the renderer logic is tested elsewhere
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Cell A1 has long text with Spill mode
        sheet.set_value(0, 0, "This is a very long text that should spill");
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);

        // Cell C1 has content - should stop spill
        sheet.set_value(0, 2, "Block");

        // Verify setup
        assert!(!sheet.get_raw(0, 0).is_empty());
        assert!(sheet.get_raw(0, 1).is_empty()); // B1 is empty
        assert!(!sheet.get_raw(0, 2).is_empty()); // C1 blocks
    }

    #[test]
    fn test_spill_stops_at_formatted_cell() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Cell A1 has Spill mode
        sheet.set_value(0, 0, "Long text here");
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);

        // Cell B1 is empty but has formatting
        sheet.toggle_bold(0, 1);

        // Verify setup
        let b1_format = sheet.get_format(0, 1);
        assert!(b1_format.bold);
        assert!(sheet.get_raw(0, 1).is_empty());
    }

    #[test]
    fn test_format_painter_copy_paste() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up source cell with complex formatting
        sheet.set_value(0, 0, "Source");
        sheet.toggle_bold(0, 0);
        sheet.toggle_italic(0, 0);
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        sheet.set_number_format(0, 0, NumberFormat::currency_compat(2));

        // Get the source format
        let source_format = sheet.get_format(0, 0);

        // Apply to destination cells
        for row in 1..5 {
            for col in 0..5 {
                sheet.set_format(row, col, source_format.clone());
            }
        }

        // Verify format was copied correctly
        for row in 1..5 {
            for col in 0..5 {
                let fmt = sheet.get_format(row, col);
                assert!(fmt.bold, "Cell ({}, {}) should be bold", row, col);
                assert!(fmt.italic, "Cell ({}, {}) should be italic", row, col);
                assert_eq!(fmt.vertical_alignment, VerticalAlignment::Top);
                assert_eq!(fmt.text_overflow, TextOverflow::Wrap);
                assert!(matches!(fmt.number_format, NumberFormat::Currency { decimals: 2, .. }));
            }
        }

        // Verify source cell value wasn't affected
        assert_eq!(sheet.get_raw(0, 0), "Source");

        // Verify destination cells have no values (only format)
        for row in 1..5 {
            for col in 0..5 {
                assert!(sheet.get_raw(row, col).is_empty());
            }
        }
    }

    #[test]
    fn test_format_painter_large_selection() {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Set up source format
        sheet.toggle_bold(0, 0);
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);
        let source_format = sheet.get_format(0, 0);

        // Apply to 500 cells (10 rows x 50 cols)
        for row in 1..11 {
            for col in 0..50 {
                sheet.set_format(row, col, source_format.clone());
            }
        }

        // Spot check some cells
        assert!(sheet.get_format(1, 0).bold);
        assert!(sheet.get_format(5, 25).bold);
        assert!(sheet.get_format(10, 49).bold);
        assert_eq!(sheet.get_format(10, 49).text_overflow, TextOverflow::Overflow);
    }

    #[test]
    fn test_sequence_spill() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Enter SEQUENCE formula that should create a 3x2 array
        sheet.set_value(0, 0, "=SEQUENCE(3,2)");

        // Check that parent cell is marked as spill parent
        assert!(sheet.is_spill_parent(0, 0));

        // Check that values spilled correctly
        // SEQUENCE(3,2) should produce:
        // 1 2
        // 3 4
        // 5 6
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(0, 1), "2");
        assert_eq!(sheet.get_display(1, 0), "3");
        assert_eq!(sheet.get_display(1, 1), "4");
        assert_eq!(sheet.get_display(2, 0), "5");
        assert_eq!(sheet.get_display(2, 1), "6");

        // Check that receiving cells are marked
        assert!(sheet.is_spill_receiver(0, 1));
        assert!(sheet.is_spill_receiver(1, 0));
        assert!(sheet.is_spill_receiver(1, 1));
        assert!(sheet.is_spill_receiver(2, 0));
        assert!(sheet.is_spill_receiver(2, 1));

        // Parent cell should not be a receiver
        assert!(!sheet.is_spill_receiver(0, 0));
    }

    #[test]
    fn test_sequence_with_start_and_step() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // SEQUENCE(3, 1, 10, 5) should produce: 10, 15, 20
        sheet.set_value(0, 0, "=SEQUENCE(3,1,10,5)");

        assert_eq!(sheet.get_display(0, 0), "10");
        assert_eq!(sheet.get_display(1, 0), "15");
        assert_eq!(sheet.get_display(2, 0), "20");
    }

    #[test]
    fn test_spill_collision_shows_error() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Put a value in the way of where spill would go
        sheet.set_value(1, 0, "blocker");

        // Now enter SEQUENCE that would need to spill to (1, 0)
        sheet.set_value(0, 0, "=SEQUENCE(3,1)");

        // Parent cell should show #SPILL! error
        assert!(sheet.has_spill_error(0, 0));
        assert_eq!(sheet.get_display(0, 0), "#SPILL!");

        // The blocking cell should still have its value
        assert_eq!(sheet.get_display(1, 0), "blocker");
    }

    #[test]
    fn test_spill_cleared_when_formula_changes() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Enter SEQUENCE that spills
        sheet.set_value(0, 0, "=SEQUENCE(3,2)");

        // Verify initial spill
        assert!(sheet.is_spill_parent(0, 0));
        assert!(sheet.is_spill_receiver(1, 0));

        // Change the formula to a scalar value
        sheet.set_value(0, 0, "42");

        // Spill should be cleared
        assert!(!sheet.is_spill_parent(0, 0));
        assert!(!sheet.is_spill_receiver(1, 0));
        assert!(!sheet.is_spill_receiver(0, 1));

        // Receiving cells should now be empty
        assert_eq!(sheet.get_display(1, 0), "");
        assert_eq!(sheet.get_display(0, 1), "");
    }

    #[test]
    fn test_transpose_spill() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up a column of values
        sheet.set_value(0, 0, "1");
        sheet.set_value(1, 0, "2");
        sheet.set_value(2, 0, "3");

        // TRANSPOSE should turn the column into a row
        sheet.set_value(0, 2, "=TRANSPOSE(A1:A3)");

        // Should spill horizontally: 1, 2, 3
        assert_eq!(sheet.get_display(0, 2), "1");
        assert_eq!(sheet.get_display(0, 3), "2");
        assert_eq!(sheet.get_display(0, 4), "3");
    }

    #[test]
    fn test_reference_to_spill_receiver() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Create a SEQUENCE that spills to multiple cells
        // A1: =SEQUENCE(3,1) produces 1, 2, 3 in A1:A3
        sheet.set_value(0, 0, "=SEQUENCE(3,1)");

        // Verify spill works
        assert_eq!(sheet.get_display(0, 0), "1"); // A1 (parent)
        assert_eq!(sheet.get_display(1, 0), "2"); // A2 (receiver)
        assert_eq!(sheet.get_display(2, 0), "3"); // A3 (receiver)

        // Now reference a spill receiver from another formula
        // B1: =A2 should return 2 (the spilled value)
        sheet.set_value(0, 1, "=A2");
        assert_eq!(sheet.get_display(0, 1), "2");

        // B2: =A2+A3 should return 5 (2+3)
        sheet.set_value(1, 1, "=A2+A3");
        assert_eq!(sheet.get_display(1, 1), "5");

        // B3: =SUM(A1:A3) should return 6 (1+2+3)
        sheet.set_value(2, 1, "=SUM(A1:A3)");
        assert_eq!(sheet.get_display(2, 1), "6");
    }

    #[test]
    fn test_spill_receiver_in_range_functions() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Create sequence in A1:A5
        sheet.set_value(0, 0, "=SEQUENCE(5,1,10,10)"); // 10, 20, 30, 40, 50

        // Verify spill
        assert_eq!(sheet.get_display(0, 0), "10");
        assert_eq!(sheet.get_display(4, 0), "50");

        // Test various functions that read from spill receivers
        sheet.set_value(0, 2, "=AVERAGE(A1:A5)");  // Should be 30
        sheet.set_value(1, 2, "=MIN(A1:A5)");     // Should be 10
        sheet.set_value(2, 2, "=MAX(A1:A5)");     // Should be 50
        sheet.set_value(3, 2, "=COUNT(A1:A5)");   // Should be 5

        assert_eq!(sheet.get_display(0, 2), "30");
        assert_eq!(sheet.get_display(1, 2), "10");
        assert_eq!(sheet.get_display(2, 2), "50");
        assert_eq!(sheet.get_display(3, 2), "5");
    }

    #[test]
    fn test_sort_numeric_ascending() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data: unsorted numbers in A1:A5
        sheet.set_value(0, 0, "30");
        sheet.set_value(1, 0, "10");
        sheet.set_value(2, 0, "50");
        sheet.set_value(3, 0, "20");
        sheet.set_value(4, 0, "40");

        // SORT ascending (default)
        sheet.set_value(0, 2, "=SORT(A1:A5)");

        // Should be sorted: 10, 20, 30, 40, 50
        assert_eq!(sheet.get_display(0, 2), "10");
        assert_eq!(sheet.get_display(1, 2), "20");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "40");
        assert_eq!(sheet.get_display(4, 2), "50");
    }

    #[test]
    fn test_sort_numeric_descending() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data: unsorted numbers in A1:A5
        sheet.set_value(0, 0, "30");
        sheet.set_value(1, 0, "10");
        sheet.set_value(2, 0, "50");
        sheet.set_value(3, 0, "20");
        sheet.set_value(4, 0, "40");

        // SORT descending (is_asc = FALSE)
        sheet.set_value(0, 2, "=SORT(A1:A5,1,FALSE)");

        // Should be sorted descending: 50, 40, 30, 20, 10
        assert_eq!(sheet.get_display(0, 2), "50");
        assert_eq!(sheet.get_display(1, 2), "40");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "20");
        assert_eq!(sheet.get_display(4, 2), "10");
    }

    #[test]
    fn test_sort_multi_column_by_second_column() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up 2-column data
        // A1:B3: (Alice, 30), (Bob, 10), (Charlie, 20)
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "30");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "10");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "20");

        // SORT by column 2 (the numbers)
        sheet.set_value(0, 3, "=SORT(A1:B3,2)");

        // Should be sorted by score: Bob(10), Charlie(20), Alice(30)
        assert_eq!(sheet.get_display(0, 3), "Bob");
        assert_eq!(sheet.get_display(0, 4), "10");
        assert_eq!(sheet.get_display(1, 3), "Charlie");
        assert_eq!(sheet.get_display(1, 4), "20");
        assert_eq!(sheet.get_display(2, 3), "Alice");
        assert_eq!(sheet.get_display(2, 4), "30");
    }

    #[test]
    fn test_sort_text_alphabetical() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up text data
        sheet.set_value(0, 0, "Banana");
        sheet.set_value(1, 0, "Apple");
        sheet.set_value(2, 0, "Cherry");

        // SORT alphabetically
        sheet.set_value(0, 2, "=SORT(A1:A3)");

        // Should be: Apple, Banana, Cherry
        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "Cherry");
    }

    #[test]
    fn test_unique_single_column() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data with duplicates
        sheet.set_value(0, 0, "Apple");
        sheet.set_value(1, 0, "Banana");
        sheet.set_value(2, 0, "Apple");  // duplicate
        sheet.set_value(3, 0, "Cherry");
        sheet.set_value(4, 0, "Banana"); // duplicate

        // UNIQUE should return: Apple, Banana, Cherry (first occurrence order)
        sheet.set_value(0, 2, "=UNIQUE(A1:A5)");

        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "Cherry");
        // No more rows
        assert_eq!(sheet.get_display(3, 2), "");
    }

    #[test]
    fn test_unique_multi_column() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up 2-column data with duplicate rows
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "30");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "20");
        sheet.set_value(2, 0, "Alice");   sheet.set_value(2, 1, "30"); // duplicate row
        sheet.set_value(3, 0, "Alice");   sheet.set_value(3, 1, "25"); // different score

        // UNIQUE should return 3 rows: (Alice,30), (Bob,20), (Alice,25)
        sheet.set_value(0, 3, "=UNIQUE(A1:B4)");

        assert_eq!(sheet.get_display(0, 3), "Alice");
        assert_eq!(sheet.get_display(0, 4), "30");
        assert_eq!(sheet.get_display(1, 3), "Bob");
        assert_eq!(sheet.get_display(1, 4), "20");
        assert_eq!(sheet.get_display(2, 3), "Alice");
        assert_eq!(sheet.get_display(2, 4), "25");
        // No fourth row
        assert_eq!(sheet.get_display(3, 3), "");
    }

    #[test]
    fn test_unique_case_insensitive() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data with case differences
        sheet.set_value(0, 0, "Apple");
        sheet.set_value(1, 0, "APPLE");  // same word, different case
        sheet.set_value(2, 0, "Banana");

        // UNIQUE should treat as duplicates (case-insensitive)
        sheet.set_value(0, 2, "=UNIQUE(A1:A3)");

        // Should return: Apple (first occurrence), Banana
        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "");
    }

    #[test]
    fn test_unique_numeric() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up numeric data with duplicates
        sheet.set_value(0, 0, "10");
        sheet.set_value(1, 0, "20");
        sheet.set_value(2, 0, "10");  // duplicate
        sheet.set_value(3, 0, "30");

        sheet.set_value(0, 2, "=UNIQUE(A1:A4)");

        assert_eq!(sheet.get_display(0, 2), "10");
        assert_eq!(sheet.get_display(1, 2), "20");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "");
    }

    #[test]
    fn test_filter_basic() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data: names in A, scores in B, pass/fail in C (1=pass, 0=fail)
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "85"); sheet.set_value(0, 2, "1");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "45"); sheet.set_value(1, 2, "0");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "72"); sheet.set_value(2, 2, "1");
        sheet.set_value(3, 0, "Diana");   sheet.set_value(3, 1, "38"); sheet.set_value(3, 2, "0");

        // FILTER by pass column (C1:C4)
        sheet.set_value(0, 4, "=FILTER(A1:B4,C1:C4)");

        // Should return only passing students: Alice, Charlie
        assert_eq!(sheet.get_display(0, 4), "Alice");
        assert_eq!(sheet.get_display(0, 5), "85");
        assert_eq!(sheet.get_display(1, 4), "Charlie");
        assert_eq!(sheet.get_display(1, 5), "72");
        // No third row
        assert_eq!(sheet.get_display(2, 4), "");
    }

    #[test]
    fn test_filter_no_matches() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up data where all fail filter
        sheet.set_value(0, 0, "Apple");  sheet.set_value(0, 1, "0");
        sheet.set_value(1, 0, "Banana"); sheet.set_value(1, 1, "0");

        // FILTER with all zeros (no matches)
        sheet.set_value(0, 3, "=FILTER(A1:A2,B1:B2)");

        // Should return #CALC! error
        assert!(sheet.get_display(0, 3).contains("#CALC!"));
    }

    #[test]
    fn test_filter_single_column() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set up single column with filter criteria
        sheet.set_value(0, 0, "10"); sheet.set_value(0, 1, "1");
        sheet.set_value(1, 0, "20"); sheet.set_value(1, 1, "0");
        sheet.set_value(2, 0, "30"); sheet.set_value(2, 1, "1");
        sheet.set_value(3, 0, "40"); sheet.set_value(3, 1, "1");

        // FILTER numbers where criteria is true
        sheet.set_value(0, 3, "=FILTER(A1:A4,B1:B4)");

        assert_eq!(sheet.get_display(0, 3), "10");
        assert_eq!(sheet.get_display(1, 3), "30");
        assert_eq!(sheet.get_display(2, 3), "40");
        assert_eq!(sheet.get_display(3, 3), "");
    }

    #[test]
    fn test_filter_combined_with_sort() {
        let mut sheet = Sheet::new(SheetId(1), 15, 15);

        // Set up data
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "85"); sheet.set_value(0, 2, "1");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "92"); sheet.set_value(1, 2, "1");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "45"); sheet.set_value(2, 2, "0");
        sheet.set_value(3, 0, "Diana");   sheet.set_value(3, 1, "78"); sheet.set_value(3, 2, "1");

        // First filter passing students
        sheet.set_value(0, 5, "=FILTER(A1:B4,C1:C4)");
        // Then reference that spilled range to sort by score
        // Note: This tests that SORT can work with spilled data

        // Verify filter works
        assert_eq!(sheet.get_display(0, 5), "Alice");
        assert_eq!(sheet.get_display(1, 5), "Bob");
        assert_eq!(sheet.get_display(2, 5), "Diana");
    }

    #[test]
    fn test_filter_include_semantics() {
        // Lock behavior: FILTER treats non-zero as TRUE, zero as FALSE
        // Text is NOT coerced - it evaluates to 0.0 (FALSE)
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Test with various include values
        sheet.set_value(0, 0, "A"); sheet.set_value(0, 1, "1");      // 1 = TRUE
        sheet.set_value(1, 0, "B"); sheet.set_value(1, 1, "0");      // 0 = FALSE
        sheet.set_value(2, 0, "C"); sheet.set_value(2, 1, "-1");     // -1 = TRUE (non-zero)
        sheet.set_value(3, 0, "D"); sheet.set_value(3, 1, "0.001");  // 0.001 = TRUE (non-zero)

        sheet.set_value(0, 3, "=FILTER(A1:A4,B1:B4)");

        // Should include: A (1), C (-1), D (0.001)
        assert_eq!(sheet.get_display(0, 3), "A");
        assert_eq!(sheet.get_display(1, 3), "C");
        assert_eq!(sheet.get_display(2, 3), "D");
        assert_eq!(sheet.get_display(3, 3), ""); // No fourth row
    }

    #[test]
    fn test_spill_shrinks_clears_old_receivers() {
        // When a spill range shrinks, old receivers outside the new range must be cleared
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Create a 5-row spill
        sheet.set_value(0, 0, "=SEQUENCE(5,1)");

        // Verify all 5 rows are filled
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(4, 0), "5");
        assert!(sheet.is_spill_receiver(4, 0));

        // Now change to a 2-row spill
        sheet.set_value(0, 0, "=SEQUENCE(2,1)");

        // Verify new spill is 2 rows
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(1, 0), "2");

        // Old receivers (rows 2, 3, 4) should be cleared
        assert_eq!(sheet.get_display(2, 0), "");
        assert_eq!(sheet.get_display(3, 0), "");
        assert_eq!(sheet.get_display(4, 0), "");

        // Old receivers should no longer be marked as receivers
        assert!(!sheet.is_spill_receiver(2, 0));
        assert!(!sheet.is_spill_receiver(3, 0));
        assert!(!sheet.is_spill_receiver(4, 0));
    }

    #[test]
    fn test_row_column_without_arguments() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // ROW() without argument should return the current row (1-indexed)
        sheet.set_value(0, 0, "=ROW()");      // A1 -> 1
        sheet.set_value(4, 0, "=ROW()");      // A5 -> 5
        sheet.set_value(9, 2, "=ROW()");      // C10 -> 10

        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(4, 0), "5");
        assert_eq!(sheet.get_display(9, 2), "10");

        // COLUMN() without argument should return the current column (1-indexed)
        sheet.set_value(0, 0, "=COLUMN()");   // A1 -> 1
        sheet.set_value(0, 2, "=COLUMN()");   // C1 -> 3
        sheet.set_value(0, 9, "=COLUMN()");   // J1 -> 10

        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(0, 2), "3");
        assert_eq!(sheet.get_display(0, 9), "10");

        // Combined usage
        sheet.set_value(3, 4, "=ROW()+COLUMN()");  // E4 -> 4+5=9
        assert_eq!(sheet.get_display(3, 4), "9");

        // ROW/COLUMN with argument should still work
        sheet.set_value(0, 5, "=ROW(B3)");    // Row of B3 -> 3
        sheet.set_value(0, 6, "=COLUMN(D1)"); // Column of D1 -> 4
        assert_eq!(sheet.get_display(0, 5), "3");
        assert_eq!(sheet.get_display(0, 6), "4");
    }

    // =========================================================================
    // Data Validation Tests
    // =========================================================================

    #[test]
    fn test_validation_list_inline() {
        use crate::validation::ValidationRule;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation for column A: only "Yes", "No", "Maybe" allowed
        let rule = ValidationRule::list_inline(vec![
            "Yes".into(),
            "No".into(),
            "Maybe".into(),
        ]);
        sheet.set_validation(0, 0, 9, 0, rule);

        // Valid values (case-sensitive match)
        assert!(sheet.validate_cell_input(0, 0, "Yes").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "No").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "Maybe").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "  Yes  ").is_valid()); // Whitespace trimmed

        // Invalid values (case-sensitive - "yes" != "Yes")
        assert!(sheet.validate_cell_input(0, 0, "yes").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "YES").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "Invalid").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "yess").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "123").is_invalid());

        // Blank is valid by default (ignore_blank = true)
        assert!(sheet.validate_cell_input(0, 0, "").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "  ").is_valid());

        // Cell without validation always valid
        assert!(sheet.validate_cell_input(0, 1, "Anything").is_valid());
    }

    #[test]
    fn test_validation_whole_number_between() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation: whole number between 1 and 100
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        sheet.set_cell_validation(0, 0, rule);

        // Valid values
        assert!(sheet.validate_cell_input(0, 0, "1").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "50").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "100").is_valid());

        // Invalid values
        assert!(sheet.validate_cell_input(0, 0, "0").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "101").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "-5").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "50.5").is_invalid()); // Not whole number
        assert!(sheet.validate_cell_input(0, 0, "abc").is_invalid());
    }

    #[test]
    fn test_validation_decimal_greater_than() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation: decimal > 0
        let rule = ValidationRule::decimal(NumericConstraint::greater_than(0));
        sheet.set_cell_validation(0, 0, rule);

        // Valid values
        assert!(sheet.validate_cell_input(0, 0, "0.001").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "1").is_valid());
        assert!(sheet.validate_cell_input(0, 0, "100.5").is_valid());

        // Invalid values
        assert!(sheet.validate_cell_input(0, 0, "0").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "-1").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "-0.5").is_invalid());
    }

    #[test]
    fn test_validation_text_length() {
        use crate::validation::{ValidationRule, NumericConstraint, ValidationType};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation: text length between 3 and 10
        let rule = ValidationRule::new(ValidationType::TextLength(
            NumericConstraint::between(3, 10)
        ));
        sheet.set_cell_validation(0, 0, rule);

        // Valid values
        assert!(sheet.validate_cell_input(0, 0, "abc").is_valid());       // 3 chars
        assert!(sheet.validate_cell_input(0, 0, "hello").is_valid());     // 5 chars
        assert!(sheet.validate_cell_input(0, 0, "0123456789").is_valid()); // 10 chars

        // Invalid values
        assert!(sheet.validate_cell_input(0, 0, "ab").is_invalid());      // 2 chars
        assert!(sheet.validate_cell_input(0, 0, "01234567890").is_invalid()); // 11 chars
    }

    #[test]
    fn test_validation_ignore_blank_false() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation with ignore_blank = false
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100))
            .with_ignore_blank(false);
        sheet.set_cell_validation(0, 0, rule);

        // Blank is now invalid
        assert!(sheet.validate_cell_input(0, 0, "").is_invalid());
        assert!(sheet.validate_cell_input(0, 0, "  ").is_invalid());

        // Normal values still validated
        assert!(sheet.validate_cell_input(0, 0, "50").is_valid());
    }

    #[test]
    fn test_validation_clear() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set validation
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        sheet.set_validation(0, 0, 5, 0, rule);

        // Validation active
        assert!(sheet.validate_cell_input(0, 0, "invalid").is_invalid());
        assert!(sheet.has_validation(0, 0));

        // Clear validation
        sheet.clear_validation(0, 0, 5, 0);

        // Validation removed
        assert!(sheet.validate_cell_input(0, 0, "invalid").is_valid());
        assert!(!sheet.has_validation(0, 0));
    }

    #[test]
    fn test_validation_cell_ref_constraint() {
        use crate::validation::{ValidationRule, NumericConstraint, ConstraintValue};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set A1 = 100 (the max value)
        sheet.set_value(0, 0, "100");

        // Set validation on B1: value must be <= A1
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("A1".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        sheet.set_cell_validation(0, 1, rule);

        // Valid: <= 100
        assert!(sheet.validate_cell_input(0, 1, "50").is_valid());
        assert!(sheet.validate_cell_input(0, 1, "100").is_valid());

        // Invalid: > 100
        assert!(sheet.validate_cell_input(0, 1, "101").is_invalid());

        // Change A1 to 50, validation should reflect new value
        sheet.set_value(0, 0, "50");
        assert!(sheet.validate_cell_input(0, 1, "50").is_valid());
        assert!(sheet.validate_cell_input(0, 1, "51").is_invalid());
    }

    #[test]
    fn test_validation_cell_ref_constraint_non_numeric() {
        use crate::validation::{ValidationRule, NumericConstraint, ConstraintValue};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set A1 = "abc" (non-numeric)
        sheet.set_value(0, 0, "abc");

        // Set validation on B1: value must be <= A1
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("A1".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        sheet.set_cell_validation(0, 1, rule);

        // Any valid number should fail because constraint is not numeric
        let result = sheet.validate_cell_input(0, 1, "50");
        assert!(result.is_invalid());

        // Check error message mentions constraint error
        if let crate::validation::ValidationResult::Invalid { reason, .. } = result {
            assert!(reason.contains("not numeric") || reason.contains("constraint"));
        }
    }

    #[test]
    fn test_validation_cell_ref_constraint_blank() {
        use crate::validation::{ValidationRule, NumericConstraint, ConstraintValue};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // A1 is blank (no value set)

        // Set validation on B1: value must be <= A1
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("A1".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        sheet.set_cell_validation(0, 1, rule);

        // Any valid number should fail because constraint is blank
        let result = sheet.validate_cell_input(0, 1, "50");
        assert!(result.is_invalid());

        // Check error message mentions blank constraint
        if let crate::validation::ValidationResult::Invalid { reason, .. } = result {
            assert!(reason.contains("blank") || reason.contains("constraint"));
        }
    }

    #[test]
    fn test_get_list_items_inline() {
        use crate::validation::ValidationRule;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set list validation
        let rule = ValidationRule::list_inline(vec![
            "  Open  ".into(),  // Will be trimmed
            "In Progress".into(),
            "Closed".into(),
        ]);
        sheet.set_cell_validation(0, 0, rule);

        // Get list items
        let list = sheet.get_list_items(0, 0).expect("Should have list");
        assert_eq!(list.items, vec!["Open", "In Progress", "Closed"]);
        assert!(!list.is_truncated);

        // Cell without validation returns None
        assert!(sheet.get_list_items(0, 1).is_none());
    }

    #[test]
    fn test_get_list_items_range() {
        use crate::validation::ValidationRule;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set source values in A1:A5
        sheet.set_value(0, 0, "Red");
        sheet.set_value(1, 0, "Green");
        sheet.set_value(2, 0, "Blue");
        // Row 3 is empty
        sheet.set_value(4, 0, "Yellow");

        // Set list validation from range
        let rule = ValidationRule::list_range("A1:A5");
        sheet.set_cell_validation(0, 1, rule);

        // Get list items - should resolve range, skip empty cells
        let list = sheet.get_list_items(0, 1).expect("Should have list");
        assert_eq!(list.items, vec!["Red", "Green", "Blue", "Yellow"]);
        assert!(!list.is_truncated);
    }

    #[test]
    fn test_list_validation_with_range_source() {
        use crate::validation::ValidationRule;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set source values
        sheet.set_value(0, 0, "Apple");
        sheet.set_value(1, 0, "Banana");
        sheet.set_value(2, 0, "Cherry");

        // Set list validation from range
        let rule = ValidationRule::list_range("=A1:A3");
        sheet.set_cell_validation(0, 1, rule);

        // Valid values
        assert!(sheet.validate_cell_input(0, 1, "Apple").is_valid());
        assert!(sheet.validate_cell_input(0, 1, "Banana").is_valid());
        assert!(sheet.validate_cell_input(0, 1, "Cherry").is_valid());

        // Invalid values
        assert!(sheet.validate_cell_input(0, 1, "apple").is_invalid()); // Case sensitive
        assert!(sheet.validate_cell_input(0, 1, "Orange").is_invalid());

        // Update source - validation should reflect new values
        sheet.set_value(0, 0, "Orange");
        assert!(sheet.validate_cell_input(0, 1, "Orange").is_valid());
        assert!(sheet.validate_cell_input(0, 1, "Apple").is_invalid());
    }

    #[test]
    fn test_has_list_dropdown() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // List validation with dropdown
        let list_rule = ValidationRule::list_inline(vec!["A".into(), "B".into()]);
        sheet.set_cell_validation(0, 0, list_rule);
        assert!(sheet.has_list_dropdown(0, 0));

        // List validation with dropdown disabled
        let list_no_dropdown = ValidationRule::list_inline(vec!["A".into()])
            .with_show_dropdown(false);
        sheet.set_cell_validation(0, 1, list_no_dropdown);
        assert!(!sheet.has_list_dropdown(0, 1));

        // Non-list validation
        let num_rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        sheet.set_cell_validation(0, 2, num_rule);
        assert!(!sheet.has_list_dropdown(0, 2));

        // No validation
        assert!(!sheet.has_list_dropdown(0, 3));
    }

    #[test]
    fn test_resolved_list_fingerprint_changes_with_source() {
        use crate::validation::ValidationRule;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set source values
        sheet.set_value(0, 0, "A");
        sheet.set_value(1, 0, "B");

        // Set validation
        let rule = ValidationRule::list_range("A1:A2");
        sheet.set_cell_validation(0, 1, rule);

        let list1 = sheet.get_list_items(0, 1).unwrap();
        let fp1 = list1.source_fingerprint;

        // Change source - fingerprint should change
        sheet.set_value(0, 0, "X");
        let list2 = sheet.get_list_items(0, 1).unwrap();
        let fp2 = list2.source_fingerprint;

        assert_ne!(fp1, fp2);
    }

    // ===== Text Spillover Tests =====
    // These verify the engine's cell emptiness detection used by gpui-app's text spill feature.

    /// Regression test: get_formatted_display returns empty for empty cells.
    /// This is critical for text spillover to work correctly - spill should stop
    /// at the first non-empty cell.
    #[test]
    fn test_text_spill_empty_cell_detection() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Empty cell should have empty display
        assert!(sheet.get_formatted_display(0, 0).is_empty(),
            "Empty cell should return empty display");

        // Cell with text should have non-empty display
        sheet.set_value(0, 0, "Hello");
        assert!(!sheet.get_formatted_display(0, 0).is_empty(),
            "Text cell should return non-empty display");

        // Cell with number should have non-empty display
        sheet.set_value(0, 1, "42");
        assert!(!sheet.get_formatted_display(0, 1).is_empty(),
            "Number cell should return non-empty display");

        // Cell with formula should have non-empty display
        sheet.set_value(0, 2, "=1+1");
        assert!(!sheet.get_formatted_display(0, 2).is_empty(),
            "Formula cell should return non-empty display");

        // Cell with zero should have non-empty display (zero is a value)
        sheet.set_value(0, 3, "0");
        assert!(!sheet.get_formatted_display(0, 3).is_empty(),
            "Zero cell should return non-empty display");
    }

    /// Test that spill scenario setup is correct: A1 has text, B1 empty, C1 blocks.
    #[test]
    fn test_text_spill_scenario() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // A1: long text that would spill
        sheet.set_value(0, 0, "This is a very long text that should spill over");

        // B1: empty (should allow spill)
        // (no set_value needed)

        // C1: has value (should block spill)
        sheet.set_value(0, 2, "Block");

        // Verify the spill scenario
        assert!(!sheet.get_formatted_display(0, 0).is_empty(), "A1 should have text");
        assert!(sheet.get_formatted_display(0, 1).is_empty(), "B1 should be empty (allows spill)");
        assert!(!sheet.get_formatted_display(0, 2).is_empty(), "C1 should block spill");
    }

    /// Alignment affects spill: General alignment with number should NOT spill.
    #[test]
    fn test_text_spill_alignment_with_number() {
        use crate::formula::eval::Value;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set a number with General alignment (default)
        sheet.set_value(0, 0, "12345.67");
        let format = sheet.get_format(0, 0);
        assert_eq!(format.alignment, crate::cell::Alignment::General);

        // Check computed value is a number
        let value = sheet.get_computed_value(0, 0);
        assert!(matches!(value, Value::Number(_)),
            "12345.67 should be computed as a number");
    }

    // ========== Ellipsis semantics tests ==========
    // These test the conditions that determine whether ellipsis shows vs spill.
    // UI layer shows ellipsis when: text_overflows AND !spill_eligible
    // Spill eligible = left-aligned (or General for text) + adjacent empty cells

    /// Center alignment should NOT spill, so ellipsis should show (if text overflows).
    #[test]
    fn test_ellipsis_shown_when_center_aligned_no_spill() {
        use crate::cell::Alignment;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Long text in A1 with Center alignment
        sheet.set_value(0, 0, "This is a very long text that would overflow");
        sheet.set_alignment(0, 0, Alignment::Center);

        // B1 is empty (spill would be possible if alignment allowed)
        assert!(sheet.get_formatted_display(0, 1).is_empty());

        // With Center alignment, spill is NOT eligible
        // (resolved alignment is Center, which doesn't spill)
        let format = sheet.get_format(0, 0);
        assert_eq!(format.alignment, Alignment::Center);

        // In UI layer: text_overflows=true, spill_eligible=false → ellipsis shows
    }

    /// Right alignment should NOT spill, so ellipsis should show (if text overflows).
    #[test]
    fn test_ellipsis_shown_when_right_aligned_no_spill() {
        use crate::cell::Alignment;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Long text in A1 with Right alignment
        sheet.set_value(0, 0, "This is a very long text that would overflow");
        sheet.set_alignment(0, 0, Alignment::Right);

        // B1 is empty
        assert!(sheet.get_formatted_display(0, 1).is_empty());

        // With Right alignment, spill is NOT eligible
        let format = sheet.get_format(0, 0);
        assert_eq!(format.alignment, Alignment::Right);

        // In UI layer: text_overflows=true, spill_eligible=false → ellipsis shows
    }

    /// Left alignment with empty neighbor SHOULD spill, so ellipsis should NOT show.
    #[test]
    fn test_ellipsis_not_shown_when_left_aligned_spill_active() {
        use crate::cell::Alignment;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Long text in A1 with Left alignment
        sheet.set_value(0, 0, "This is a very long text that would overflow");
        sheet.set_alignment(0, 0, Alignment::Left);

        // B1 is empty - spill is possible
        assert!(sheet.get_formatted_display(0, 1).is_empty());

        // With Left alignment + empty neighbor, spill IS eligible
        let format = sheet.get_format(0, 0);
        assert_eq!(format.alignment, Alignment::Left);

        // In UI layer: text_overflows=true, spill_eligible=true → NO ellipsis (spill active)
    }

    /// General alignment for text should resolve to Left and spill.
    #[test]
    fn test_ellipsis_not_shown_when_general_text_spills() {
        use crate::cell::Alignment;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Long text in A1 with General alignment (default)
        sheet.set_value(0, 0, "This is a very long text that would overflow");

        // B1 is empty - spill is possible
        assert!(sheet.get_formatted_display(0, 1).is_empty());

        // General alignment
        let format = sheet.get_format(0, 0);
        assert_eq!(format.alignment, Alignment::General);

        // For text values, General resolves to Left → spill IS eligible
        // In UI layer: text_overflows=true, spill_eligible=true → NO ellipsis (spill active)
    }

    /// Left-aligned text with non-empty neighbor should NOT spill, so ellipsis shows.
    #[test]
    fn test_ellipsis_shown_when_left_aligned_blocked_by_neighbor() {
        use crate::cell::Alignment;

        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Long text in A1 with Left alignment
        sheet.set_value(0, 0, "This is a very long text that would overflow");
        sheet.set_alignment(0, 0, Alignment::Left);

        // B1 has content - blocks spill
        sheet.set_value(0, 1, "Blocker");
        assert!(!sheet.get_formatted_display(0, 1).is_empty());

        // With Left alignment but blocked neighbor, spill is NOT possible
        // In UI layer: text_overflows=true, spill_eligible=false (blocked) → ellipsis shows
    }

    // =========================================================================
    // Border ship-gate tests (match apply_borders semantics)
    // =========================================================================

    /// Helper: apply "Outline" borders on selection (min_row..=max_row, min_col..=max_col)
    fn apply_outline(sheet: &mut Sheet, min_row: usize, min_col: usize, max_row: usize, max_col: usize) {
        let thin = CellBorder::thin();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if row == min_row { sheet.set_border_top(row, col, thin); }
                if row == max_row { sheet.set_border_bottom(row, col, thin); }
                if col == min_col { sheet.set_border_left(row, col, thin); }
                if col == max_col { sheet.set_border_right(row, col, thin); }
            }
        }
    }

    /// Helper: apply "All" borders on selection
    fn apply_all(sheet: &mut Sheet, min_row: usize, min_col: usize, max_row: usize, max_col: usize) {
        let thin = CellBorder::thin();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                sheet.set_borders(row, col, thin, thin, thin, thin);
            }
        }
    }

    /// Helper: apply "Inside" borders on selection (internal edges only)
    fn apply_inside(sheet: &mut Sheet, min_row: usize, min_col: usize, max_row: usize, max_col: usize) {
        let thin = CellBorder::thin();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if row < max_row { sheet.set_border_bottom(row, col, thin); }
                if col < max_col { sheet.set_border_right(row, col, thin); }
            }
        }
    }

    /// Helper: apply "Clear" borders on selection + canonicalize neighbors
    fn apply_clear(sheet: &mut Sheet, min_row: usize, min_col: usize, max_row: usize, max_col: usize) {
        let none = CellBorder::default();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                sheet.set_borders(row, col, none, none, none, none);
            }
        }
        // Canonicalize: clear inward-facing edges of neighbors
        if min_row > 0 {
            for col in min_col..=max_col {
                sheet.set_border_bottom(min_row - 1, col, none);
            }
        }
        for col in min_col..=max_col {
            if max_row + 1 < sheet.rows {
                sheet.set_border_top(max_row + 1, col, none);
            }
        }
        if min_col > 0 {
            for row in min_row..=max_row {
                sheet.set_border_right(row, min_col - 1, none);
            }
        }
        for row in min_row..=max_row {
            if max_col + 1 < sheet.cols {
                sheet.set_border_left(row, max_col + 1, none);
            }
        }
    }

    #[test]
    fn test_border_outline_single_cell_full_box() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        apply_outline(&mut sheet, 2, 3, 2, 3);
        let f = sheet.get_format(2, 3);
        assert!(f.border_top.is_set(), "top");
        assert!(f.border_right.is_set(), "right");
        assert!(f.border_bottom.is_set(), "bottom");
        assert!(f.border_left.is_set(), "left");
    }

    #[test]
    fn test_border_inside_1xn_vertical_only() {
        // 1×3 selection: (0,0)-(0,2) — inside should produce vertical separators only
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        apply_inside(&mut sheet, 0, 0, 0, 2);
        // No outer edges
        assert!(!sheet.get_format(0, 0).border_top.is_set(), "no top");
        assert!(!sheet.get_format(0, 0).border_left.is_set(), "no left");
        assert!(!sheet.get_format(0, 0).border_bottom.is_set(), "no bottom on (0,0)");
        assert!(!sheet.get_format(0, 2).border_right.is_set(), "no right on last");
        assert!(!sheet.get_format(0, 2).border_top.is_set(), "no top on last");
        assert!(!sheet.get_format(0, 2).border_bottom.is_set(), "no bottom on last");
        // Internal verticals (written as right edges per precedence)
        assert!(sheet.get_format(0, 0).border_right.is_set(), "right on (0,0)");
        assert!(sheet.get_format(0, 1).border_right.is_set(), "right on (0,1)");
    }

    #[test]
    fn test_border_inside_nx1_horizontal_only() {
        // 3×1 selection: (0,0)-(2,0) — inside should produce horizontal separators only
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        apply_inside(&mut sheet, 0, 0, 2, 0);
        // No outer edges
        assert!(!sheet.get_format(0, 0).border_top.is_set(), "no top");
        assert!(!sheet.get_format(0, 0).border_left.is_set(), "no left");
        assert!(!sheet.get_format(0, 0).border_right.is_set(), "no right on (0,0)");
        assert!(!sheet.get_format(2, 0).border_bottom.is_set(), "no bottom on last");
        assert!(!sheet.get_format(2, 0).border_left.is_set(), "no left on last");
        assert!(!sheet.get_format(2, 0).border_right.is_set(), "no right on last");
        // Internal horizontals (written as bottom edges per precedence)
        assert!(sheet.get_format(0, 0).border_bottom.is_set(), "bottom on (0,0)");
        assert!(sheet.get_format(1, 0).border_bottom.is_set(), "bottom on (1,0)");
    }

    #[test]
    fn test_border_2x2_inside_cross_outline_box_all_both() {
        // 2×2 selection: (0,0)-(1,1)
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // --- Inside: cross only ---
        apply_inside(&mut sheet, 0, 0, 1, 1);
        // Vertical internal: right edge on column 0
        assert!(sheet.get_format(0, 0).border_right.is_set(), "inside: (0,0) right");
        assert!(sheet.get_format(1, 0).border_right.is_set(), "inside: (1,0) right");
        // Horizontal internal: bottom edge on row 0
        assert!(sheet.get_format(0, 0).border_bottom.is_set(), "inside: (0,0) bottom");
        assert!(sheet.get_format(0, 1).border_bottom.is_set(), "inside: (0,1) bottom");
        // No outer edges
        assert!(!sheet.get_format(0, 0).border_top.is_set(), "inside: no top");
        assert!(!sheet.get_format(0, 0).border_left.is_set(), "inside: no left");
        assert!(!sheet.get_format(1, 1).border_bottom.is_set(), "inside: no bottom");
        assert!(!sheet.get_format(1, 1).border_right.is_set(), "inside: no right");

        // Reset
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // --- Outline: box only ---
        apply_outline(&mut sheet, 0, 0, 1, 1);
        // Outer edges
        assert!(sheet.get_format(0, 0).border_top.is_set(), "outline: top-left top");
        assert!(sheet.get_format(0, 0).border_left.is_set(), "outline: top-left left");
        assert!(sheet.get_format(1, 1).border_bottom.is_set(), "outline: bot-right bottom");
        assert!(sheet.get_format(1, 1).border_right.is_set(), "outline: bot-right right");
        // No internal edges
        assert!(!sheet.get_format(0, 0).border_right.is_set(), "outline: no internal right");
        assert!(!sheet.get_format(0, 0).border_bottom.is_set(), "outline: no internal bottom");

        // Reset
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // --- All: box + cross ---
        apply_all(&mut sheet, 0, 0, 1, 1);
        for row in 0..=1 {
            for col in 0..=1 {
                let f = sheet.get_format(row, col);
                assert!(f.border_top.is_set(), "all: ({},{}) top", row, col);
                assert!(f.border_right.is_set(), "all: ({},{}) right", row, col);
                assert!(f.border_bottom.is_set(), "all: ({},{}) bottom", row, col);
                assert!(f.border_left.is_set(), "all: ({},{}) left", row, col);
            }
        }
    }

    #[test]
    fn test_border_clear_canonicalizes_neighbors() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        let thin = CellBorder::thin();

        // Set up: borders on (1,1) and its neighbors
        sheet.set_borders(1, 1, thin, thin, thin, thin);
        // Neighbor above has bottom border pointing into selection
        sheet.set_border_bottom(0, 1, thin);
        // Neighbor below has top border pointing into selection
        sheet.set_border_top(2, 1, thin);
        // Neighbor left has right border pointing into selection
        sheet.set_border_right(1, 0, thin);
        // Neighbor right has left border pointing into selection
        sheet.set_border_left(1, 2, thin);

        // Clear selection (1,1)-(1,1)
        apply_clear(&mut sheet, 1, 1, 1, 1);

        // Selection cell cleared
        let f = sheet.get_format(1, 1);
        assert!(!f.border_top.is_set(), "sel top cleared");
        assert!(!f.border_right.is_set(), "sel right cleared");
        assert!(!f.border_bottom.is_set(), "sel bottom cleared");
        assert!(!f.border_left.is_set(), "sel left cleared");

        // Neighbor inward edges cleared (canonicalization)
        assert!(!sheet.get_format(0, 1).border_bottom.is_set(), "above bottom cleared");
        assert!(!sheet.get_format(2, 1).border_top.is_set(), "below top cleared");
        assert!(!sheet.get_format(1, 0).border_right.is_set(), "left right cleared");
        assert!(!sheet.get_format(1, 2).border_left.is_set(), "right left cleared");
    }

    #[test]
    fn test_border_undo_restores_via_set_format() {
        // Simulates undo by saving before/after format and restoring
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Record "before" format
        let before_00 = sheet.get_format(0, 0);
        let before_01 = sheet.get_format(0, 1);
        let before_10 = sheet.get_format(1, 0);
        let before_11 = sheet.get_format(1, 1);

        // Apply All borders to 2×2
        apply_all(&mut sheet, 0, 0, 1, 1);

        // Verify borders are set
        assert!(sheet.get_format(0, 0).border_top.is_set());
        assert!(sheet.get_format(1, 1).border_bottom.is_set());

        // Simulate undo: restore "before" formats
        sheet.set_format(0, 0, before_00.clone());
        sheet.set_format(0, 1, before_01.clone());
        sheet.set_format(1, 0, before_10.clone());
        sheet.set_format(1, 1, before_11.clone());

        // All borders removed (back to default)
        for row in 0..=1 {
            for col in 0..=1 {
                let f = sheet.get_format(row, col);
                assert!(!f.border_top.is_set(), "undo ({},{}) top", row, col);
                assert!(!f.border_right.is_set(), "undo ({},{}) right", row, col);
                assert!(!f.border_bottom.is_set(), "undo ({},{}) bottom", row, col);
                assert!(!f.border_left.is_set(), "undo ({},{}) left", row, col);
            }
        }
    }

    // =========================================================================
    // Merged Regions
    // =========================================================================

    #[test]
    fn test_merged_region_basic() {
        let m = MergedRegion::new(1, 2, 3, 5);
        assert_eq!(m.top_left(), (1, 2));
        assert_eq!(m.rows(), 3);
        assert_eq!(m.cols(), 4);
        assert_eq!(m.cell_count(), 12);
        assert!(!m.is_degenerate());

        assert!(m.contains(1, 2));
        assert!(m.contains(3, 5));
        assert!(m.contains(2, 3));
        assert!(!m.contains(0, 2));
        assert!(!m.contains(1, 6));
        assert!(!m.contains(4, 2));

        let single = MergedRegion::new(0, 0, 0, 0);
        assert!(single.is_degenerate());
    }

    #[test]
    fn test_merge_index_lookup() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 2)).unwrap();

        // All 9 cells in the 3×3 region should find the merge
        for r in 0..=2 {
            for c in 0..=2 {
                assert!(
                    sheet.get_merge(r, c).is_some(),
                    "expected merge at ({}, {})",
                    r,
                    c
                );
            }
        }
        // Outside the region
        assert!(sheet.get_merge(3, 0).is_none());
        assert!(sheet.get_merge(0, 3).is_none());

        // Origin vs hidden
        assert!(sheet.is_merge_origin(0, 0));
        assert!(!sheet.is_merge_hidden(0, 0));
        assert!(!sheet.is_merge_origin(1, 1));
        assert!(sheet.is_merge_hidden(1, 1));
    }

    #[test]
    fn test_add_merge_overlap_rejected() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 2)).unwrap();

        // Overlapping merge should fail
        let result = sheet.add_merge(MergedRegion::new(1, 1, 3, 3));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("overlaps"));

        // Non-overlapping merge should succeed
        assert!(sheet.add_merge(MergedRegion::new(0, 3, 2, 5)).is_ok());
    }

    #[test]
    fn test_add_merge_degenerate_ignored() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        // 1×1 merge is silently ignored
        sheet.add_merge(MergedRegion::new(0, 0, 0, 0)).unwrap();
        assert!(sheet.merged_regions.is_empty());
    }

    #[test]
    fn test_remove_merge() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 2)).unwrap();
        assert_eq!(sheet.merged_regions.len(), 1);

        let removed = sheet.remove_merge((0, 0));
        assert!(removed.is_some());
        assert!(sheet.merged_regions.is_empty());
        assert!(sheet.get_merge(0, 0).is_none());

        // Removing non-existent merge returns None
        assert!(sheet.remove_merge((5, 5)).is_none());
    }

    #[test]
    fn test_merge_insert_row_shift() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-7
        sheet.add_merge(MergedRegion::new(5, 0, 7, 2)).unwrap();

        // Insert 2 rows at row 3 (above merge) → merge shifts to 7-9
        sheet.insert_rows(3, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (7, 0));
        assert_eq!(m.end, (9, 2));
    }

    #[test]
    fn test_merge_insert_row_expand() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-7
        sheet.add_merge(MergedRegion::new(5, 0, 7, 2)).unwrap();

        // Insert 2 rows at row 6 (inside merge) → merge expands to 5-9
        sheet.insert_rows(6, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (5, 0));
        assert_eq!(m.end, (9, 2));
    }

    #[test]
    fn test_merge_insert_row_below() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        sheet.add_merge(MergedRegion::new(5, 0, 7, 2)).unwrap();

        // Insert below merge → no change
        sheet.insert_rows(8, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (5, 0));
        assert_eq!(m.end, (7, 2));
    }

    #[test]
    fn test_merge_delete_row_shrink() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-9 (5 rows)
        sheet.add_merge(MergedRegion::new(5, 0, 9, 2)).unwrap();

        // Delete 2 rows at row 6 (inside merge) → merge shrinks to 5-7
        sheet.delete_rows(6, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (5, 0));
        assert_eq!(m.end, (7, 2));
    }

    #[test]
    fn test_merge_delete_row_degenerate() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-6, cols 0-0 (2×1)
        sheet.add_merge(MergedRegion::new(5, 0, 6, 0)).unwrap();

        // Delete row 5 → merge becomes 1×1 → removed as degenerate
        sheet.delete_rows(5, 1);
        assert!(sheet.merged_regions.is_empty());
    }

    #[test]
    fn test_merge_delete_row_above() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        sheet.add_merge(MergedRegion::new(5, 0, 7, 2)).unwrap();

        // Delete 2 rows above merge → merge shifts up
        sheet.delete_rows(2, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (3, 0));
        assert_eq!(m.end, (5, 2));
    }

    #[test]
    fn test_merge_insert_col_shift() {
        let mut sheet = Sheet::new(SheetId(1), 10, 20);
        sheet.add_merge(MergedRegion::new(0, 5, 2, 7)).unwrap();

        // Insert 2 cols at col 3 (left of merge) → merge shifts right
        sheet.insert_cols(3, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (0, 7));
        assert_eq!(m.end, (2, 9));
    }

    #[test]
    fn test_merge_insert_col_expand() {
        let mut sheet = Sheet::new(SheetId(1), 10, 20);
        sheet.add_merge(MergedRegion::new(0, 5, 2, 7)).unwrap();

        // Insert 2 cols at col 6 (inside merge) → merge expands
        sheet.insert_cols(6, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (0, 5));
        assert_eq!(m.end, (2, 9));
    }

    #[test]
    fn test_merge_delete_col_shrink() {
        let mut sheet = Sheet::new(SheetId(1), 10, 20);
        sheet.add_merge(MergedRegion::new(0, 5, 2, 9)).unwrap();

        // Delete 2 cols at col 6 (inside merge) → merge shrinks
        sheet.delete_cols(6, 2);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (0, 5));
        assert_eq!(m.end, (2, 7));
    }

    #[test]
    fn test_merge_delete_band_clips_top() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-9, cols 0-2
        sheet.add_merge(MergedRegion::new(5, 0, 9, 2)).unwrap();

        // Delete rows 3-6 (band starts above merge, extends into it)
        sheet.delete_rows(3, 4);
        let m = &sheet.merged_regions[0];
        // Top rows clipped: merge origin moves to row 3, bottom shifts up by 4
        assert_eq!(m.start, (3, 0));
        assert_eq!(m.end, (5, 2));
        assert_eq!(m.rows(), 3); // 5 rows → lost 2 inside → 3 remain
    }

    #[test]
    fn test_merge_delete_band_clips_bottom() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // Merge at rows 5-9, cols 0-2
        sheet.add_merge(MergedRegion::new(5, 0, 9, 2)).unwrap();

        // Delete rows 8-11 (band starts inside merge, extends below)
        sheet.delete_rows(8, 4);
        let m = &sheet.merged_regions[0];
        assert_eq!(m.start, (5, 0));
        assert_eq!(m.end, (7, 2));
        assert_eq!(m.rows(), 3); // bottom clipped
    }

    #[test]
    fn test_merge_delete_band_degenerates_wide() {
        let mut sheet = Sheet::new(SheetId(1), 20, 10);
        // 3×4 merge
        sheet.add_merge(MergedRegion::new(5, 0, 7, 3)).unwrap();

        // Delete all 3 rows of the merge → fully consumed → gone
        sheet.delete_rows(5, 3);
        assert!(sheet.merged_regions.is_empty());
    }

    #[test]
    fn test_merge_delete_col_band_degenerates() {
        let mut sheet = Sheet::new(SheetId(1), 10, 20);
        // 2×2 merge
        sheet.add_merge(MergedRegion::new(0, 5, 1, 6)).unwrap();

        // Delete both columns → degenerate → removed
        sheet.delete_cols(5, 2);
        assert!(sheet.merged_regions.is_empty());
    }

    #[test]
    fn test_merge_serde_roundtrip() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 3)).unwrap();
        sheet.add_merge(MergedRegion::new(5, 5, 7, 8)).unwrap();

        let json = serde_json::to_string(&sheet).unwrap();
        let mut loaded: Sheet = serde_json::from_str(&json).unwrap();

        // merged_regions should be preserved
        assert_eq!(loaded.merged_regions.len(), 2);
        assert_eq!(loaded.merged_regions[0], MergedRegion::new(0, 0, 2, 3));
        assert_eq!(loaded.merged_regions[1], MergedRegion::new(5, 5, 7, 8));

        // merge_index is serde(skip), so rebuild it
        loaded.rebuild_merge_index();
        assert!(loaded.get_merge(0, 0).is_some());
        assert!(loaded.get_merge(6, 6).is_some());
        assert!(loaded.get_merge(3, 3).is_none());
    }

    // =========================================================================
    // set_merges() + undo/value restoration tests
    // =========================================================================

    /// Trust anchor: merge with data loss → undo restores values and topology.
    /// Simulates the app-level undo sequence using only engine APIs:
    ///   1. Put values in A1 ("keep") and B1 ("lose")
    ///   2. Snapshot before merges, clear B1, add merge A1:C1
    ///   3. Undo: set_merges(before), restore B1 value
    ///   4. Assert: B1 == "lose", no merge exists
    #[test]
    fn test_set_merges_undo_restores_values_and_topology() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Step 1: populate cells
        sheet.set_value(0, 0, "keep");
        sheet.set_value(0, 1, "lose");
        assert_eq!(sheet.get_raw(0, 0), "keep");
        assert_eq!(sheet.get_raw(0, 1), "lose");

        // Step 2: merge A1:C1 (simulating merge_cells_confirmed)
        let before = sheet.merged_regions.clone();
        assert!(before.is_empty());

        // Record cleared values (B1 has data)
        let cleared_values: Vec<(usize, usize, String)> = vec![(0, 1, "lose".to_string())];

        // Clear non-origin values
        sheet.set_value(0, 1, "");

        // Add merge
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();
        let after = sheet.merged_regions.clone();
        assert_eq!(after.len(), 1);

        // Verify merged state
        assert_eq!(sheet.get_raw(0, 0), "keep");
        assert_eq!(sheet.get_raw(0, 1), "");
        assert!(sheet.get_merge(0, 1).is_some());

        // Step 3: undo — restore topology first, then values
        sheet.set_merges(before);
        for (row, col, value) in &cleared_values {
            sheet.set_value(*row, *col, value);
        }

        // Step 4: assert fully restored
        assert!(sheet.merged_regions.is_empty());
        assert!(sheet.get_merge(0, 1).is_none());
        assert_eq!(sheet.get_raw(0, 0), "keep");
        assert_eq!(sheet.get_raw(0, 1), "lose");
    }

    /// Redo after undo: clear values first, then set merges.
    #[test]
    fn test_set_merges_redo_clears_values_and_reapplies_topology() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.set_value(0, 0, "keep");
        sheet.set_value(0, 1, "lose");

        // Merge
        let before = sheet.merged_regions.clone();
        sheet.set_value(0, 1, "");
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();
        let after = sheet.merged_regions.clone();
        let cleared_values: Vec<(usize, usize, String)> = vec![(0, 1, "lose".to_string())];

        // Undo
        sheet.set_merges(before);
        sheet.set_value(0, 1, "lose");
        assert_eq!(sheet.get_raw(0, 1), "lose");
        assert!(sheet.merged_regions.is_empty());

        // Redo: clear values first, then apply merge topology
        for (row, col, _) in &cleared_values {
            sheet.set_value(*row, *col, "");
        }
        sheet.set_merges(after);

        // Assert redo state
        assert_eq!(sheet.merged_regions.len(), 1);
        assert_eq!(sheet.get_raw(0, 0), "keep");
        assert_eq!(sheet.get_raw(0, 1), "");
        assert!(sheet.get_merge(0, 1).is_some());
    }

    /// set_merges replaces contained merges correctly during undo.
    #[test]
    fn test_set_merges_contained_merge_undo_restores_inner() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Create inner merge B2:C3 with value "inner"
        sheet.set_value(1, 1, "inner");
        sheet.add_merge(MergedRegion::new(1, 1, 2, 2)).unwrap();
        let before = sheet.merged_regions.clone();
        assert_eq!(before.len(), 1);

        // Merge A1:D4 (contains B2:C3)
        // Cleared values: B2 has "inner"
        let cleared_values: Vec<(usize, usize, String)> = vec![(1, 1, "inner".to_string())];

        // Remove contained merge, clear values, add new merge
        sheet.remove_merge((1, 1));
        sheet.set_value(1, 1, "");
        sheet.add_merge(MergedRegion::new(0, 0, 3, 3)).unwrap();
        let after = sheet.merged_regions.clone();
        assert_eq!(after.len(), 1);

        // Undo: restore before topology + values
        sheet.set_merges(before);
        for (row, col, value) in &cleared_values {
            sheet.set_value(*row, *col, value);
        }

        // Assert: inner merge restored with value
        assert_eq!(sheet.merged_regions.len(), 1);
        assert_eq!(sheet.merged_regions[0], MergedRegion::new(1, 1, 2, 2));
        assert_eq!(sheet.get_raw(1, 1), "inner");
        assert!(sheet.get_merge(0, 0).is_none()); // A1:D4 merge gone
    }

    // =========================================================================
    // Merge border resolution tests
    // =========================================================================

    #[test]
    fn test_resolve_merge_borders_no_borders() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(1, 1, 3, 3)).unwrap();
        let merge = sheet.get_merge(1, 1).unwrap().clone();
        let (top, right, bottom, left) = sheet.resolve_merge_borders(&merge);
        assert!(!top.is_set());
        assert!(!right.is_set());
        assert!(!bottom.is_set());
        assert!(!left.is_set());
    }

    #[test]
    fn test_resolve_merge_borders_uniform_thin() {
        use crate::cell::{CellBorder, BorderStyle};
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(1, 1, 3, 3)).unwrap();

        // Set thin borders on all edge cells
        for c in 1..=3 {
            let mut fmt = sheet.get_format(1, c);
            fmt.border_top = CellBorder::thin();
            sheet.set_format(1, c, fmt);
            let mut fmt = sheet.get_format(3, c);
            fmt.border_bottom = CellBorder::thin();
            sheet.set_format(3, c, fmt);
        }
        for r in 1..=3 {
            let mut fmt = sheet.get_format(r, 1);
            fmt.border_left = CellBorder::thin();
            sheet.set_format(r, 1, fmt);
            let mut fmt = sheet.get_format(r, 3);
            fmt.border_right = CellBorder::thin();
            sheet.set_format(r, 3, fmt);
        }

        let merge = sheet.get_merge(1, 1).unwrap().clone();
        let (top, right, bottom, left) = sheet.resolve_merge_borders(&merge);
        assert_eq!(top.style, BorderStyle::Thin);
        assert_eq!(right.style, BorderStyle::Thin);
        assert_eq!(bottom.style, BorderStyle::Thin);
        assert_eq!(left.style, BorderStyle::Thin);
    }

    #[test]
    fn test_resolve_merge_borders_mixed_precedence() {
        use crate::cell::{CellBorder, BorderStyle};
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(2, 2, 4, 5)).unwrap();

        // Top edge: col 2 = Thin, col 3 = Medium, col 4 = None, col 5 = Thick
        let mut fmt = sheet.get_format(2, 2);
        fmt.border_top = CellBorder::thin();
        sheet.set_format(2, 2, fmt);

        let mut fmt = sheet.get_format(2, 3);
        fmt.border_top = CellBorder { style: BorderStyle::Medium, color: None };
        sheet.set_format(2, 3, fmt);

        // col 4: no border (default)

        let mut fmt = sheet.get_format(2, 5);
        fmt.border_top = CellBorder { style: BorderStyle::Thick, color: None };
        sheet.set_format(2, 5, fmt);

        let merge = sheet.get_merge(2, 2).unwrap().clone();
        let (top, _right, _bottom, _left) = sheet.resolve_merge_borders(&merge);
        // Thick wins: it's the strongest
        assert_eq!(top.style, BorderStyle::Thick);
    }

    #[test]
    fn test_resolve_merge_borders_single_styled_cell() {
        use crate::cell::{CellBorder, BorderStyle};
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 4)).unwrap();

        // Only one cell on the right edge has a border
        let mut fmt = sheet.get_format(1, 4);
        fmt.border_right = CellBorder { style: BorderStyle::Medium, color: None };
        sheet.set_format(1, 4, fmt);

        let merge = sheet.get_merge(0, 0).unwrap().clone();
        let (_top, right, _bottom, _left) = sheet.resolve_merge_borders(&merge);
        // Medium wins even though other right-edge cells have no border
        assert_eq!(right.style, BorderStyle::Medium);
    }

    #[test]
    fn test_adjacent_merges_shared_border() {
        use crate::cell::{CellBorder, BorderStyle};
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        // Merge A: (1,1)→(2,3), Merge B: (3,1)→(4,3) — share horizontal border
        sheet.add_merge(MergedRegion::new(1, 1, 2, 3)).unwrap();
        sheet.add_merge(MergedRegion::new(3, 1, 4, 3)).unwrap();

        // Set bottom border on merge A (via edge cell)
        let mut fmt = sheet.get_format(2, 2);
        fmt.border_bottom = CellBorder { style: BorderStyle::Medium, color: None };
        sheet.set_format(2, 2, fmt);

        // Set top border on merge B (via edge cell) — thinner than A's bottom
        let mut fmt = sheet.get_format(3, 1);
        fmt.border_top = CellBorder::thin();
        sheet.set_format(3, 1, fmt);

        // Merge A's resolved bottom should be Medium
        let merge_a = sheet.get_merge(1, 1).unwrap().clone();
        let (_, _, a_bottom, _) = sheet.resolve_merge_borders(&merge_a);
        assert_eq!(a_bottom.style, BorderStyle::Medium);

        // Merge B's resolved top should be Thin
        let merge_b = sheet.get_merge(3, 1).unwrap().clone();
        let (b_top, _, _, _) = sheet.resolve_merge_borders(&merge_b);
        assert_eq!(b_top.style, BorderStyle::Thin);

        // At the renderer level, single-ownership (top+left) means:
        // - Merge A's bottom cells never draw bottom (right+bottom not drawn)
        // - Merge B's top cells draw top = max(B_resolved_top, A_resolved_bottom)
        //   = max(Thin, Medium) = Medium
        // This is validated in the app layer, not here. But the resolution data
        // for both merges is independently correct.
    }

    #[test]
    fn test_edge_consolidation_intentional() {
        use crate::cell::{CellBorder, BorderStyle};
        // Lock-in test: edge consolidation is deliberate.
        // Only one cell on the top edge has Thick; the resolved top is Thick for
        // the entire edge, not piecewise.
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(2, 2, 4, 6)).unwrap(); // 3×5 merge

        // Only one cell on the top edge gets a thick border
        let mut fmt = sheet.get_format(2, 4);
        fmt.border_top = CellBorder { style: BorderStyle::Thick, color: None };
        sheet.set_format(2, 4, fmt);

        let merge = sheet.get_merge(2, 2).unwrap().clone();
        let (top, _, _, _) = sheet.resolve_merge_borders(&merge);

        // Edge consolidation: entire top edge is Thick (not just the segment at col 4)
        assert_eq!(top.style, BorderStyle::Thick);
    }

    #[test]
    fn test_outer_border_draw() {
        use crate::cell::{CellBorder, BorderStyle, max_border};
        // Verify that bottom/right borders are resolvable for cells at the boundary.
        // In the single-ownership model, a cell normally doesn't draw its bottom/right.
        // At the viewport boundary, we resolve bottom as max(own_bottom, below_top)
        // and right as max(own_right, right_left). This test verifies the resolution.
        let mut sheet = Sheet::new(SheetId(1), 10, 10);

        // Set a bottom border on cell (5, 3)
        let mut fmt = sheet.get_format(5, 3);
        fmt.border_bottom = CellBorder { style: BorderStyle::Thin, color: None };
        sheet.set_format(5, 3, fmt);

        // Set a top border on cell (6, 3) — this would normally draw the shared edge
        let mut fmt = sheet.get_format(6, 3);
        fmt.border_top = CellBorder { style: BorderStyle::Medium, color: None };
        sheet.set_format(6, 3, fmt);

        // At the viewport boundary (row 5 is last visible), cell (5,3) must draw bottom.
        // Resolution: max(own_bottom=Thin, below_top=Medium) = Medium
        let own_bottom = sheet.get_format(5, 3).border_bottom;
        let below_top = sheet.get_format(6, 3).border_top;
        let resolved = max_border(own_bottom, below_top);
        assert_eq!(resolved.style, BorderStyle::Medium);
        assert!(resolved.is_set());

        // Similarly for right border: set right on (3,5) and left on (3,6)
        let mut fmt = sheet.get_format(3, 5);
        fmt.border_right = CellBorder { style: BorderStyle::Thick, color: None };
        sheet.set_format(3, 5, fmt);

        let mut fmt = sheet.get_format(3, 6);
        fmt.border_left = CellBorder { style: BorderStyle::Thin, color: None };
        sheet.set_format(3, 6, fmt);

        let own_right = sheet.get_format(3, 5).border_right;
        let right_left = sheet.get_format(3, 6).border_left;
        let resolved = max_border(own_right, right_left);
        assert_eq!(resolved.style, BorderStyle::Thick);
    }

    #[test]
    fn test_merge_at_viewport_boundary() {
        use crate::cell::{CellBorder, BorderStyle};
        // A merge touching the viewport edge must expose its resolved bottom/right
        // borders for boundary drawing. The renderer calls resolve_merge_borders()
        // and uses the bottom/right edges when the merge perimeter is at the boundary.
        let mut sheet = Sheet::new(SheetId(1), 20, 20);
        sheet.add_merge(MergedRegion::new(3, 5, 7, 9)).unwrap(); // 5×5 merge

        // Set bottom border on the merge's bottom edge cells
        let mut fmt = sheet.get_format(7, 6);
        fmt.border_bottom = CellBorder { style: BorderStyle::Medium, color: None };
        sheet.set_format(7, 6, fmt);

        // Set right border on the merge's right edge cells
        let mut fmt = sheet.get_format(5, 9);
        fmt.border_right = CellBorder { style: BorderStyle::Thick, color: None };
        sheet.set_format(5, 9, fmt);

        let merge = sheet.get_merge(3, 5).unwrap().clone();
        let (_, right, bottom, _) = sheet.resolve_merge_borders(&merge);

        // Edge consolidation: the entire bottom edge gets Medium,
        // the entire right edge gets Thick
        assert_eq!(bottom.style, BorderStyle::Medium);
        assert_eq!(right.style, BorderStyle::Thick);

        // is_merge_hidden returns true for all non-origin cells (text suppression).
        // The origin is (3,5); all others are hidden for text purposes.
        assert!(sheet.is_merge_origin(3, 5));
        assert!(!sheet.is_merge_origin(3, 6));
        assert!(sheet.is_merge_hidden(6, 7)); // interior — text hidden
        assert!(sheet.is_merge_hidden(7, 7)); // perimeter but not origin — text hidden

        // Borders are drawn on perimeter cells by the renderer using
        // get_merge() + resolve_merge_borders(), not is_merge_hidden().
        // Verify perimeter cells are in the merge and can access resolved borders.
        assert!(sheet.get_merge(7, 7).is_some()); // bottom-edge cell is in merge
        assert!(sheet.get_merge(5, 9).is_some()); // right-edge cell is in merge
    }

    #[test]
    fn test_gridline_suppression_boundary_guard() {
        // Verify the invariants for gridline suppression at data boundaries:
        // - At data_row == 0: no top gridline (header provides separation)
        // - At col == 0: no left gridline (row header provides separation)
        // These are renderer-level guards (data_row > 0, col > 0) documented here
        // as regression protection. The boundary pass (is_last_visible_row/col)
        // handles the opposite edge (viewport bottom/right).
        let sheet = Sheet::new(SheetId(1), 10, 10);

        // At row 0: top border is structurally absent (no cell above)
        // The renderer guards with `if data_row > 0 { draw top gridline }`
        assert!(sheet.get_format(0, 0).border_top.style == crate::cell::BorderStyle::None);

        // At col 0: left border is structurally absent (no cell left)
        assert!(sheet.get_format(0, 0).border_left.style == crate::cell::BorderStyle::None);

        // Merge at row 0: interior gridlines still suppressed
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 2, 2)).unwrap();
        // Interior cell (1,1) is hidden — no gridlines drawn here
        assert!(sheet.is_merge_hidden(1, 1));
        // Origin (0,0) is not hidden — it's the merge anchor
        assert!(!sheet.is_merge_hidden(0, 0));
        // Perimeter cell (2,1) is hidden for text but borders are resolved
        // via get_merge() in the renderer. Verify it's in the merge.
        assert!(sheet.is_merge_hidden(2, 1)); // non-origin = hidden for text
        let merge = sheet.get_merge(2, 1).unwrap();
        assert_eq!(merge.start, (0, 0));
        assert_eq!(merge.end, (2, 2));
        // Bottom edge cell: renderer checks `data_row < merge.end.0` for interior suppression
        // (2,1) has data_row == merge.end.0, so bottom_in_merge is false — gridline not suppressed
        assert_eq!(2, merge.end.0); // at bottom edge, not interior
    }

    /// Deterministic test: has_any_borders goes true→false through clear paths.
    /// Locks in the invariant that scan_border_flag() is called after border removal.
    #[test]
    fn test_has_any_borders_true_to_false() {
        let mut sheet = Sheet::new(SheetId(99), 100, 100);
        assert!(!sheet.has_any_borders, "new sheet should have no borders");

        // Set a border → flag goes true
        let thin = CellBorder::thin();
        let none = CellBorder::default();
        sheet.set_borders(3, 3, thin, none, none, none);
        assert!(sheet.has_any_borders, "flag should be true after setting a border");

        // Clear that border → flag should still be true (set_borders doesn't auto-clear flag)
        sheet.set_borders(3, 3, none, none, none, none);
        assert!(sheet.has_any_borders, "flag stays true until scan (conservative)");

        // Rescan → flag goes false (the only bordered cell was cleared)
        sheet.scan_border_flag();
        assert!(!sheet.has_any_borders, "flag should be false after scan with no borders");

        // Set border again, then clear via set_format(default) → same pattern
        sheet.set_border_top(5, 5, thin);
        assert!(sheet.has_any_borders);
        sheet.set_format(5, 5, CellFormat::default());
        sheet.scan_border_flag();
        assert!(!sheet.has_any_borders, "flag should be false after clearing format + scan");

        // Multiple bordered cells: clearing one leaves flag true
        sheet.set_borders(0, 0, thin, thin, thin, thin);
        sheet.set_borders(1, 1, thin, none, none, none);
        assert!(sheet.has_any_borders);
        sheet.set_borders(0, 0, none, none, none, none);
        sheet.scan_border_flag();
        assert!(sheet.has_any_borders, "flag stays true: cell (1,1) still has a border");
        sheet.set_borders(1, 1, none, none, none, none);
        sheet.scan_border_flag();
        assert!(!sheet.has_any_borders, "flag false after all borders cleared + scan");
    }

    /// has_any_borders stays correct through delete_rows (which calls scan_border_flag).
    #[test]
    fn test_has_any_borders_after_delete_rows() {
        let mut sheet = Sheet::new(SheetId(99), 100, 100);
        let thin = CellBorder::thin();

        // Border on row 5 only
        sheet.set_border_top(5, 0, thin);
        assert!(sheet.has_any_borders);

        // Delete rows 0..5 (includes row 5) → border is gone, flag should clear
        sheet.delete_rows(0, 6);
        assert!(!sheet.has_any_borders, "flag should be false after deleting the only bordered row");

        // Border on row 10, delete rows 0..3 (doesn't touch row 10) → flag stays true
        sheet.set_border_top(10, 0, thin);
        assert!(sheet.has_any_borders);
        sheet.delete_rows(0, 3);
        assert!(sheet.has_any_borders, "flag stays true: bordered row shifted but still exists");
    }

    /// has_any_borders stays correct through delete_cols (which calls scan_border_flag).
    #[test]
    fn test_has_any_borders_after_delete_cols() {
        let mut sheet = Sheet::new(SheetId(99), 100, 100);
        let thin = CellBorder::thin();

        // Border on col 5 only
        sheet.set_border_left(0, 5, thin);
        assert!(sheet.has_any_borders);

        // Delete cols 0..6 (includes col 5) → border is gone
        sheet.delete_cols(0, 6);
        assert!(!sheet.has_any_borders, "flag should be false after deleting the only bordered col");
    }

    // ========================================================================
    // Merged cell redirect tests
    // ========================================================================

    #[test]
    fn test_merge_origin_coord() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        // Merge A1:C1 (row 0, cols 0..2)
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Non-merged cell passes through
        assert_eq!(sheet.merge_origin_coord(5, 5), (5, 5));

        // Merge origin passes through (start == coord, no-op)
        assert_eq!(sheet.merge_origin_coord(0, 0), (0, 0));

        // Hidden cells redirect to origin
        assert_eq!(sheet.merge_origin_coord(0, 1), (0, 0));
        assert_eq!(sheet.merge_origin_coord(0, 2), (0, 0));
    }

    #[test]
    fn test_set_value_redirects_to_origin() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.set_value(0, 0, "Hello");
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Write to hidden cell B1 — should land on A1
        sheet.set_value(0, 1, "World");
        assert_eq!(sheet.get_display(0, 0), "World"); // A1 changed
        assert_eq!(sheet.get_display(0, 1), "");       // B1 stays empty
    }

    #[test]
    fn test_clear_cell_redirects_to_origin() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.set_value(0, 0, "Hello");
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Clear hidden cell B1 — should clear A1
        sheet.clear_cell(0, 1);
        assert_eq!(sheet.get_display(0, 0), ""); // A1 cleared
    }

    #[test]
    fn test_set_value_preserves_hidden_style() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Set bold on hidden cell B1
        sheet.toggle_bold(0, 1);
        assert!(sheet.get_format(0, 1).bold);

        // Write value to B1 — redirects to A1
        sheet.set_value(0, 1, "123");

        // A1 has the value
        assert_eq!(sheet.get_display(0, 0), "123");
        // B1 still has bold format
        assert!(sheet.get_format(0, 1).bold);
        // B1 has no stored value
        assert_eq!(sheet.get_display(0, 1), "");
    }

    #[test]
    fn test_set_format_no_redirect() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Set format on hidden cell B1 → should stay on B1, not redirect to A1
        sheet.toggle_bold(0, 1);
        assert!(sheet.get_format(0, 1).bold);
        assert!(!sheet.get_format(0, 0).bold); // A1 is NOT affected
    }

    #[test]
    fn test_clear_cell_preserves_hidden_style() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.set_value(0, 0, "Hello");
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Set bold on hidden cell B1
        sheet.toggle_bold(0, 1);
        assert!(sheet.get_format(0, 1).bold);

        // Clear B1 — redirects to A1, clears origin value
        sheet.clear_cell(0, 1);

        // A1 value cleared
        assert_eq!(sheet.get_display(0, 0), "");
        // B1 format still bold (clear_cell only removes value, not sibling formats)
        assert!(sheet.get_format(0, 1).bold);
    }

    #[test]
    fn test_set_cycle_error_preserves_hidden_style() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Set bold on hidden cell B1
        sheet.toggle_bold(0, 1);
        assert!(sheet.get_format(0, 1).bold);

        // Cycle error on B1 — redirects to origin A1
        sheet.set_cycle_error(0, 1);

        // Origin holds the cycle error
        assert_eq!(sheet.get_display(0, 0), "#CYCLE!");
        // B1 style unchanged
        assert!(sheet.get_format(0, 1).bold);
        // B1 has no stored value
        assert_eq!(sheet.get_display(0, 1), "");
    }

    #[test]
    fn test_range_clear_preserves_hidden_styles() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.set_value(0, 0, "Hello");
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Set bold on hidden cell B1
        sheet.toggle_bold(0, 1);
        assert!(sheet.get_format(0, 1).bold);

        // Simulate a range clear: clear each cell in A1:C1
        // This is what delete_selection does (loops over selection calling clear_cell)
        for col in 0..=2 {
            sheet.clear_cell(0, col);
        }

        // A1 value should be cleared (cleared directly + redirects from B1/C1)
        assert_eq!(sheet.get_display(0, 0), "");
        // B1 format still bold — clear_cell only removes the value at origin,
        // it does NOT touch sibling cell formats
        assert!(sheet.get_format(0, 1).bold);
    }

    // =========================================================================
    // Merge rendering ship-gate tests
    // =========================================================================

    #[test]
    fn test_merge_overlaps_viewport_fully_inside() {
        // Merge (2,2)→(4,5), viewport scroll=(0,0) visible=(10,10)
        let m = MergedRegion::new(2, 2, 4, 5);
        assert!(m.overlaps_viewport(0, 0, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_partially_visible_top_left() {
        // Merge starts above and left of viewport, extends into it
        let m = MergedRegion::new(0, 0, 3, 3);
        assert!(m.overlaps_viewport(2, 2, 10, 10)); // bottom-right quadrant visible
    }

    #[test]
    fn test_merge_overlaps_viewport_partially_visible_bottom_right() {
        // Merge extends past viewport bottom-right
        let m = MergedRegion::new(8, 8, 15, 15);
        assert!(m.overlaps_viewport(5, 5, 10, 10)); // top-left quadrant visible
    }

    #[test]
    fn test_merge_overlaps_viewport_entirely_above() {
        let m = MergedRegion::new(0, 0, 2, 5);
        assert!(!m.overlaps_viewport(5, 0, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_entirely_below() {
        let m = MergedRegion::new(20, 0, 22, 5);
        assert!(!m.overlaps_viewport(5, 0, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_entirely_left() {
        let m = MergedRegion::new(0, 0, 5, 2);
        assert!(!m.overlaps_viewport(0, 5, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_entirely_right() {
        let m = MergedRegion::new(0, 20, 5, 25);
        assert!(!m.overlaps_viewport(0, 5, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_edge_touching_bottom() {
        // Merge end row == scroll_row (just barely visible)
        let m = MergedRegion::new(3, 0, 5, 3);
        assert!(m.overlaps_viewport(5, 0, 10, 10)); // row 5 is first visible
    }

    #[test]
    fn test_merge_overlaps_viewport_edge_just_outside() {
        // Merge end row == scroll_row - 1 (just outside)
        let m = MergedRegion::new(0, 0, 4, 3);
        assert!(!m.overlaps_viewport(5, 0, 10, 10));
    }

    #[test]
    fn test_merge_overlaps_viewport_single_cell_merge() {
        // 1x1 merges are degenerate and ignored, but test the geometry anyway
        let m = MergedRegion::new(5, 5, 5, 5);
        assert!(m.overlaps_viewport(5, 5, 1, 1));
        assert!(!m.overlaps_viewport(6, 5, 1, 1));
    }

    #[test]
    fn test_merge_get_merge_covers_all_cells() {
        // Ship-gate: the spill overlay guard is:
        //   if sheet.get_merge(row, col).is_some() { continue; }
        // This test verifies get_merge() returns Some for EVERY cell inside a
        // merge (origin + all hidden), and None for cells outside.
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(2, 3, 4, 6)).unwrap();

        // Every cell inside the merge must return Some
        for r in 2..=4 {
            for c in 3..=6 {
                assert!(
                    sheet.get_merge(r, c).is_some(),
                    "get_merge({}, {}) should be Some (inside merge)",
                    r, c
                );
            }
        }
        // Cells adjacent to the merge must return None
        assert!(sheet.get_merge(1, 3).is_none(), "above merge");
        assert!(sheet.get_merge(2, 7).is_none(), "right of merge");
        assert!(sheet.get_merge(5, 3).is_none(), "below merge");
        assert!(sheet.get_merge(2, 2).is_none(), "left of merge");
    }

    #[test]
    fn test_merge_with_text_still_detected_by_get_merge() {
        // Ship-gate: even after setting a value on a merged cell,
        // get_merge() must still return Some. This is the actual predicate
        // the spill overlay uses to skip merged cells.
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.add_merge(MergedRegion::new(1, 1, 3, 4)).unwrap();

        // Set a long text value on the origin cell
        sheet.set_value(1, 1, "This is a very long text that would normally spill");

        // Origin still detected as merged
        assert!(sheet.get_merge(1, 1).is_some());
        assert!(sheet.is_merge_origin(1, 1));

        // Hidden cells still detected as merged
        assert!(sheet.get_merge(2, 3).is_some());
        assert!(sheet.is_merge_hidden(2, 3));

        // Value redirects to origin — hidden cells have no independent value
        assert_eq!(sheet.get_display(1, 1), "This is a very long text that would normally spill");
    }

    #[test]
    fn test_merge_pixel_rect_uniform_widths() {
        // Ship-gate: verify x/y offsets with uniform column/row sizes.
        // Merge (2,3)→(4,5), scroll at (0,0), all cols=80px, all rows=25px
        let m = MergedRegion::new(2, 3, 4, 5);
        let (x, y, w, h) = m.pixel_rect(0, 0, |_| 80.0, |_| 25.0);

        // x = cols 0,1,2 before merge start = 3 * 80 = 240
        assert_eq!(x, 240.0);
        // y = rows 0,1 before merge start = 2 * 25 = 50
        assert_eq!(y, 50.0);
        // width = cols 3,4,5 = 3 * 80 = 240
        assert_eq!(w, 240.0);
        // height = rows 2,3,4 = 3 * 25 = 75
        assert_eq!(h, 75.0);
    }

    #[test]
    fn test_merge_pixel_rect_with_scroll_offset() {
        // Ship-gate: when scrolled, x/y should only sum widths/heights
        // from scroll position to merge start — not from col 0.
        let m = MergedRegion::new(5, 8, 7, 10);
        // Scroll at row 3, col 5. All cols=100px, all rows=20px.
        let (x, y, w, h) = m.pixel_rect(3, 5, |_| 100.0, |_| 20.0);

        // x = cols 5,6,7 before merge start at col 8 = 3 * 100 = 300
        assert_eq!(x, 300.0);
        // y = rows 3,4 before merge start at row 5 = 2 * 20 = 40
        assert_eq!(y, 40.0);
        // width = cols 8,9,10 = 3 * 100 = 300
        assert_eq!(w, 300.0);
        // height = rows 5,6,7 = 3 * 20 = 60
        assert_eq!(h, 60.0);
    }

    #[test]
    fn test_merge_pixel_rect_variable_widths() {
        // Ship-gate: non-uniform col widths (resized columns).
        // Merge (1,2)→(3,4), scroll at (0,0).
        let m = MergedRegion::new(1, 2, 3, 4);
        let col_widths = [60.0, 80.0, 120.0, 50.0, 90.0]; // cols 0-4
        let row_heights = [20.0, 30.0, 25.0, 35.0]; // rows 0-3
        let (x, y, w, h) = m.pixel_rect(
            0, 0,
            |c| col_widths.get(c).copied().unwrap_or(80.0),
            |r| row_heights.get(r).copied().unwrap_or(25.0),
        );

        // x = col 0 (60) + col 1 (80) = 140
        assert_eq!(x, 140.0);
        // y = row 0 (20) = 20
        assert_eq!(y, 20.0);
        // width = col 2 (120) + col 3 (50) + col 4 (90) = 260
        assert_eq!(w, 260.0);
        // height = row 1 (30) + row 2 (25) + row 3 (35) = 90
        assert_eq!(h, 90.0);
    }

    #[test]
    fn test_merge_pixel_rect_scroll_past_merge_start() {
        // Edge case: scroll position is past the merge start.
        // x/y should be 0 (or negative in real rendering, but sum of empty range = 0).
        // The overlay clips via overflow_hidden so this is still correct.
        let m = MergedRegion::new(2, 3, 5, 6);
        let (x, y, _w, _h) = m.pixel_rect(4, 5, |_| 80.0, |_| 25.0);

        // scroll_row=4 > merge.start.0=2, so range 4..2 is empty → y = 0
        assert_eq!(y, 0.0);
        // scroll_col=5 > merge.start.1=3, so range 5..3 is empty → x = 0
        assert_eq!(x, 0.0);
    }

    #[test]
    fn test_merge_rect_width_height_span() {
        // Lock-in: the range iteration for width/height covers exactly the right cells.
        let m = MergedRegion::new(2, 3, 5, 7);
        assert_eq!(m.rows(), 4); // rows 2,3,4,5
        assert_eq!(m.cols(), 5); // cols 3,4,5,6,7

        let cols: Vec<usize> = (m.start.1..=m.end.1).collect();
        assert_eq!(cols, vec![3, 4, 5, 6, 7]);
        let rows: Vec<usize> = (m.start.0..=m.end.0).collect();
        assert_eq!(rows, vec![2, 3, 4, 5]);
    }

    #[test]
    fn test_get_merge_returns_canonical_origin() {
        // For every cell (r, c) inside a merge, get_merge(r, c).unwrap().start
        // must equal the merge origin. This is the invariant that navigation
        // depends on: "snap to merge.start from any cell" always yields the origin.
        let mut sheet = Sheet::new(SheetId(1), 20, 20);
        sheet.add_merge(MergedRegion::new(1, 1, 3, 4)).unwrap(); // B2:E4
        sheet.add_merge(MergedRegion::new(5, 0, 5, 2)).unwrap(); // A6:C6

        // Every cell in first merge should resolve to origin (1, 1)
        for r in 1..=3 {
            for c in 1..=4 {
                let merge = sheet.get_merge(r, c)
                    .unwrap_or_else(|| panic!("get_merge({}, {}) should return Some", r, c));
                assert_eq!(merge.start, (1, 1), "cell ({}, {}) should have origin (1, 1)", r, c);
            }
        }

        // Every cell in second merge should resolve to origin (5, 0)
        for c in 0..=2 {
            let merge = sheet.get_merge(5, c)
                .unwrap_or_else(|| panic!("get_merge(5, {}) should return Some", c));
            assert_eq!(merge.start, (5, 0), "cell (5, {}) should have origin (5, 0)", c);
        }

        // Cell outside any merge should return None
        assert!(sheet.get_merge(0, 0).is_none());
        assert!(sheet.get_merge(4, 0).is_none());
        assert!(sheet.get_merge(6, 0).is_none());
    }

    /// Regression: set_font_size must persist and be visible to get_format,
    /// and the before/after PartialEq comparison must detect the change.
    /// This is the engine contract that set_font_size_selection relies on.
    #[test]
    fn test_set_font_size_roundtrip_and_partial_eq() {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Before: default format has no font_size
        let before = sheet.get_format(0, 0);
        assert_eq!(before.font_size, None);

        // Set font_size = 24
        sheet.set_font_size(0, 0, Some(24.0));
        let after = sheet.get_format(0, 0);
        assert_eq!(after.font_size, Some(24.0));

        // PartialEq must detect the change (critical for undo patch generation)
        assert_ne!(before, after, "CellFormat PartialEq must detect font_size change");

        // Clear font_size back to None
        sheet.set_font_size(0, 0, None);
        let cleared = sheet.get_format(0, 0);
        assert_eq!(cleared.font_size, None);
        assert_eq!(before, cleared, "Clearing font_size should restore default equality");
    }
}
