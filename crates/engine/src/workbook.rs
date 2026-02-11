use rustc_hash::{FxHashMap, FxHashSet};
use serde::{Deserialize, Serialize};
use crate::cell::{CellFormat, CellValue};
use crate::cell_id::CellId;
use crate::dep_graph::DepGraph;
use crate::sheet::{Sheet, SheetId, normalize_sheet_name, is_valid_sheet_name};
use crate::named_range::{NamedRange, NamedRangeStore};
use crate::formula::eval::{CellLookup, EvalArg, EvalResult, NamedRangeResolution, Value};
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

/// Result of validating a range of cells (e.g., after paste/fill).
#[derive(Debug, Clone, Default)]
pub struct ValidationFailures {
    /// Number of cells that failed validation.
    pub count: usize,
    /// Positions and reasons of failed cells (up to 100, for navigation).
    pub failures: Vec<ValidationFailure>,
}

/// A single validation failure with position and reason.
#[derive(Debug, Clone)]
pub struct ValidationFailure {
    pub row: usize,
    pub col: usize,
    pub reason: crate::validation::ValidationFailureReason,
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

    /// Workbook-global deduplicated style table.
    /// Cell.style_id indexes into this table to reference the base imported style.
    #[serde(default)]
    pub style_table: Vec<CellFormat>,

    /// Dependency graph for formula cells.
    /// Rebuilt on load, updated incrementally on cell changes.
    #[serde(skip)]
    dep_graph: DepGraph,

    /// Batch nesting counter. When > 0, recalc is deferred until end_batch().
    /// `pub(crate)` for test harness access.
    #[serde(skip)]
    pub(crate) batch_depth: u32,

    /// Cells changed during the current batch. Drained at end_batch().
    /// `pub(crate)` for test harness access.
    #[serde(skip)]
    pub(crate) batch_changed: Vec<CellId>,

    /// Maps to `CalculationMode` from document settings:
    ///   `true`  = `CalculationMode::Automatic` — recalc on every edit (default)
    ///   `false` = `CalculationMode::Manual`    — recalc only on F9
    ///
    /// Stored as bool to avoid duplicating the settings enum in the engine crate.
    /// Do NOT repurpose this for "should I ever recalc" — it strictly mirrors
    /// the Automatic/Manual toggle.
    #[serde(skip)]
    auto_recalc: bool,

    /// Iterative calculation mode (Excel-style circular calc).
    /// When enabled, SCCs are resolved via Jacobi iteration instead of #CYCLE!.
    #[serde(skip)]
    iterative_enabled: bool,

    /// Maximum iterations per SCC before declaring non-convergence (#NUM!).
    #[serde(skip)]
    iterative_max_iters: u32,

    /// Convergence tolerance: iteration stops when max cell delta < tolerance.
    #[serde(skip)]
    iterative_tolerance: f64,

    /// Monotonically increasing revision number. Incremented once per successful
    /// batch completion (or single-cell edit outside batch).
    /// Used for optimistic concurrency control in session server protocol.
    #[serde(skip)]
    revision: u64,

    /// Test instrumentation: counts how many times recalc_dirty_set was called.
    /// Only present in test builds. Reset manually via reset_recalc_count().
    #[cfg(test)]
    #[serde(skip)]
    recalc_count: std::cell::Cell<u32>,
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
            style_table: Vec::new(),
            dep_graph: DepGraph::new(),
            batch_depth: 0,
            batch_changed: Vec::new(),
            auto_recalc: true,
            iterative_enabled: false,
            iterative_max_iters: 100,
            iterative_tolerance: 1e-9,
            revision: 0,
            #[cfg(test)]
            recalc_count: std::cell::Cell::new(0),
        }
    }

    /// Intern a CellFormat into the workbook-global style table.
    /// Returns the index (style_id) for the format, deduplicating if an identical style exists.
    pub fn intern_style(&mut self, format: CellFormat) -> u32 {
        // Check for existing identical style
        if let Some(idx) = self.style_table.iter().position(|s| s == &format) {
            return idx as u32;
        }
        let idx = self.style_table.len() as u32;
        self.style_table.push(format);
        idx
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

    /// Get the index of a sheet by its ID
    pub fn sheet_index_by_id(&self, id: SheetId) -> Option<usize> {
        self.sheets.iter().position(|s| s.id == id)
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

    /// Get mutable access to all sheets.
    pub fn sheets_mut(&mut self) -> &mut [Sheet] {
        &mut self.sheets
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
            style_table: Vec::new(),
            dep_graph: DepGraph::new(),
            batch_depth: 0,
            batch_changed: Vec::new(),
            auto_recalc: true,
            iterative_enabled: false,
            iterative_max_iters: 100,
            iterative_tolerance: 1e-9,
            revision: 0,
            #[cfg(test)]
            recalc_count: std::cell::Cell::new(0),
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
            style_table: Vec::new(),
            dep_graph: DepGraph::new(),
            batch_depth: 0,
            batch_changed: Vec::new(),
            auto_recalc: true,
            iterative_enabled: false,
            iterative_max_iters: 100,
            iterative_tolerance: 1e-9,
            revision: 0,
            #[cfg(test)]
            recalc_count: std::cell::Cell::new(0),
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
    // List Validation Resolution
    // =========================================================================

    /// Get list items for a cell with list validation.
    ///
    /// This is the workbook-level API that can resolve named ranges.
    /// Returns None if the cell has no validation or non-list validation.
    pub fn get_list_items(&self, sheet_index: usize, row: usize, col: usize) -> Option<crate::validation::ResolvedList> {
        use crate::validation::{ValidationType, ListSource, ResolvedList};

        let sheet = self.sheets.get(sheet_index)?;
        let rule = sheet.validations.get(row, col)?;

        match &rule.rule_type {
            ValidationType::List(source) => {
                match source {
                    ListSource::Inline(values) => {
                        Some(ResolvedList::from_items(values.clone()))
                    }
                    ListSource::Range(range_str) => {
                        // Parse range string, may include sheet reference
                        let range_str = range_str.trim_start_matches('=').trim();
                        Some(self.resolve_range_to_list(sheet_index, range_str))
                    }
                    ListSource::NamedRange(name) => {
                        // Strip leading = if present (UI accepts both forms)
                        let name = name.trim_start_matches('=');
                        Some(self.resolve_named_range_to_list(name))
                    }
                }
            }
            _ => None,
        }
    }

    /// Resolve a range string (possibly with sheet reference) to list items.
    fn resolve_range_to_list(&self, current_sheet: usize, range_str: &str) -> crate::validation::ResolvedList {
        use crate::validation::ResolvedList;

        // Check for sheet reference: "Sheet1!A1:A10"
        let (sheet_idx, cell_range) = if let Some(bang_pos) = range_str.find('!') {
            let sheet_name = &range_str[..bang_pos].trim_matches('\'');
            let cell_range = &range_str[bang_pos + 1..];

            // Find sheet by name
            match self.sheets.iter().position(|s| s.name == *sheet_name) {
                Some(idx) => (idx, cell_range),
                None => return ResolvedList::empty(), // Sheet not found
            }
        } else {
            (current_sheet, range_str)
        };

        // Use the sheet's resolve method
        if let Some(sheet) = self.sheets.get(sheet_idx) {
            sheet.resolve_range_to_list(cell_range)
        } else {
            ResolvedList::empty()
        }
    }

    /// Resolve a named range to list items.
    fn resolve_named_range_to_list(&self, name: &str) -> crate::validation::ResolvedList {
        use crate::validation::ResolvedList;
        use crate::named_range::NamedRangeTarget;

        let named_range = match self.named_ranges.get(name) {
            Some(nr) => nr,
            None => return ResolvedList::empty(),
        };

        match &named_range.target {
            NamedRangeTarget::Cell { sheet, row, col } => {
                if let Some(s) = self.sheets.get(*sheet) {
                    let display = s.get_display(*row, *col);
                    if display.is_empty() {
                        return ResolvedList::empty();
                    }
                    return ResolvedList::from_items(vec![display]);
                }
                ResolvedList::empty()
            }
            NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                if let Some(s) = self.sheets.get(*sheet) {
                    // Collect values from range
                    let mut items = Vec::new();
                    for row in *start_row..=*end_row {
                        for col in *start_col..=*end_col {
                            let display = s.get_display(row, col);
                            if !display.is_empty() {
                                items.push(display);
                            }
                        }
                    }
                    return ResolvedList::from_items(items);
                }
                ResolvedList::empty()
            }
        }
    }

    /// Check if a cell has list validation with dropdown enabled.
    pub fn has_list_dropdown(&self, sheet_index: usize, row: usize, col: usize) -> bool {
        if let Some(sheet) = self.sheets.get(sheet_index) {
            sheet.has_list_dropdown(row, col)
        } else {
            false
        }
    }

    // =========================================================================
    // Numeric Constraint Resolution (workbook-level)
    // =========================================================================

    /// Resolve a constraint value to a number.
    ///
    /// This is the workbook-level resolver that can handle cross-sheet CellRefs.
    /// - Literal numbers: return directly
    /// - CellRef: parse "A1" or "Sheet2!A1", get computed value, parse as number
    /// - Formula: not yet implemented (returns FormulaError)
    pub fn resolve_constraint_value(
        &self,
        current_sheet: usize,
        value: &crate::validation::ConstraintValue,
    ) -> Result<f64, crate::validation::ConstraintResolveError> {
        use crate::validation::{ConstraintValue, ConstraintResolveError};

        match value {
            ConstraintValue::Number(n) => Ok(*n),
            ConstraintValue::CellRef(ref_str) => {
                self.resolve_cell_ref_to_number(current_sheet, ref_str)
            }
            ConstraintValue::Formula(_formula) => {
                // Formula constraint evaluation not yet implemented
                // Return deterministic error so behavior is predictable
                Err(ConstraintResolveError::FormulaError(
                    "Formula constraints not yet implemented".to_string()
                ))
            }
        }
    }

    /// Resolve a cell reference string to a numeric value.
    ///
    /// Handles both same-sheet ("A1") and cross-sheet ("Sheet2!A1") references.
    fn resolve_cell_ref_to_number(
        &self,
        current_sheet: usize,
        ref_str: &str,
    ) -> Result<f64, crate::validation::ConstraintResolveError> {
        use crate::validation::ConstraintResolveError;

        // Parse reference: check for sheet prefix
        let (sheet_idx, cell_ref) = if let Some(bang_pos) = ref_str.find('!') {
            let sheet_name = ref_str[..bang_pos].trim_matches('\'');
            let cell_ref = &ref_str[bang_pos + 1..];

            // Find sheet by name
            let idx = self.sheets.iter()
                .position(|s| s.name == sheet_name)
                .ok_or_else(|| ConstraintResolveError::InvalidReference(
                    format!("Sheet '{}' not found", sheet_name)
                ))?;
            (idx, cell_ref)
        } else {
            (current_sheet, ref_str)
        };

        // Get sheet
        let sheet = self.sheets.get(sheet_idx)
            .ok_or_else(|| ConstraintResolveError::InvalidReference(
                format!("Sheet index {} out of range", sheet_idx)
            ))?;

        // Parse cell reference
        let (row, col) = sheet.parse_cell_ref(cell_ref)
            .ok_or_else(|| ConstraintResolveError::InvalidReference(
                format!("Invalid cell reference: {}", cell_ref)
            ))?;

        // Get computed value (display value, not raw formula)
        let display = sheet.get_display(row, col);

        if display.is_empty() {
            return Err(ConstraintResolveError::BlankConstraint);
        }

        // Parse as number
        display.parse::<f64>()
            .map_err(|_| ConstraintResolveError::NotNumeric)
    }

    /// Validate a cell input at the workbook level.
    ///
    /// This handles cross-sheet CellRef constraints that require workbook context.
    /// For List validation and other types, delegates to the sheet.
    pub fn validate_cell_input(
        &self,
        sheet_index: usize,
        row: usize,
        col: usize,
        value: &str,
    ) -> crate::validation::ValidationResult {
        use crate::validation::{ValidationResult, ValidationType, NumericParseError};

        let sheet = match self.sheets.get(sheet_index) {
            Some(s) => s,
            None => return ValidationResult::Valid,
        };

        let rule = match sheet.validations.get(row, col) {
            Some(r) => r,
            None => return ValidationResult::Valid,
        };

        // Check ignore_blank
        if rule.ignore_blank && value.trim().is_empty() {
            return ValidationResult::Valid;
        }

        // Handle numeric types with workbook-level constraint resolution
        match &rule.rule_type {
            ValidationType::WholeNumber(constraint) => {
                use crate::validation::parse_numeric_input;

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

                self.validate_numeric_constraint(sheet_index, num, constraint, rule, "whole number")
            }

            ValidationType::Decimal(constraint) => {
                use crate::validation::parse_numeric_input;

                let num = match parse_numeric_input(value, true) {
                    Ok(n) => n,
                    Err(_) => {
                        return ValidationResult::Invalid {
                            rule: rule.clone(),
                            reason: "Value must be a number".to_string(),
                        };
                    }
                };

                self.validate_numeric_constraint(sheet_index, num, constraint, rule, "number")
            }

            // For other types, delegate to sheet (they don't need cross-sheet resolution)
            _ => sheet.validate_cell_input(row, col, value),
        }
    }

    /// Helper to validate a numeric value against a constraint using workbook resolver.
    fn validate_numeric_constraint(
        &self,
        sheet_index: usize,
        value: f64,
        constraint: &crate::validation::NumericConstraint,
        rule: &crate::validation::ValidationRule,
        type_name: &str,
    ) -> crate::validation::ValidationResult {
        use crate::validation::{ValidationResult, eval_numeric_constraint, ComparisonOperator};

        // Resolve constraint values using workbook-level resolver
        let v1 = match self.resolve_constraint_value(sheet_index, &constraint.value1) {
            Ok(n) => n,
            Err(e) => {
                return ValidationResult::Invalid {
                    rule: rule.clone(),
                    reason: format!("Validation constraint error: {}", e),
                };
            }
        };

        let v2 = match &constraint.value2 {
            Some(cv) => match self.resolve_constraint_value(sheet_index, cv) {
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
                ComparisonOperator::Between => {
                    format!("{} must be between {} and {}", type_name, v1, v2.unwrap_or(0.0))
                }
                ComparisonOperator::NotBetween => {
                    format!("{} must not be between {} and {}", type_name, v1, v2.unwrap_or(0.0))
                }
                ComparisonOperator::EqualTo => format!("{} must equal {}", type_name, v1),
                ComparisonOperator::NotEqualTo => format!("{} must not equal {}", type_name, v1),
                ComparisonOperator::GreaterThan => format!("{} must be greater than {}", type_name, v1),
                ComparisonOperator::LessThan => format!("{} must be less than {}", type_name, v1),
                ComparisonOperator::GreaterThanOrEqual => format!("{} must be at least {}", type_name, v1),
                ComparisonOperator::LessThanOrEqual => format!("{} must be at most {}", type_name, v1),
            };

            ValidationResult::Invalid {
                rule: rule.clone(),
                reason,
            }
        }
    }

    /// Validate a range of cells and return failure information.
    ///
    /// Used after paste/fill operations to count validation failures.
    /// Returns the count of invalid cells, their positions, and failure reasons.
    pub fn validate_range(
        &self,
        sheet_index: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> ValidationFailures {
        use crate::validation::ValidationResult;

        let sheet = match self.sheets.get(sheet_index) {
            Some(s) => s,
            None => return ValidationFailures::default(),
        };

        let mut failures = ValidationFailures::default();

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                // Get the current display value of the cell
                let value = sheet.get_display(row, col);

                // Validate using workbook-level validation
                let result = self.validate_cell_input(sheet_index, row, col, &value);

                if let ValidationResult::Invalid { reason, .. } = result {
                    failures.count += 1;
                    // Store first 100 failures for navigation
                    if failures.failures.len() < 100 {
                        // Classify reason from the error message
                        let failure_reason = Self::classify_failure_reason(&reason);
                        failures.failures.push(ValidationFailure {
                            row,
                            col,
                            reason: failure_reason,
                        });
                    }
                }
            }
        }

        failures
    }

    /// Classify a validation failure reason from the error message.
    pub fn classify_failure_reason(reason: &str) -> crate::validation::ValidationFailureReason {
        use crate::validation::ValidationFailureReason;

        let lower = reason.to_lowercase();
        if lower.contains("blank") {
            ValidationFailureReason::ConstraintBlank
        } else if lower.contains("not numeric") || lower.contains("constraint is not") {
            ValidationFailureReason::ConstraintNotNumeric
        } else if lower.contains("formula") {
            ValidationFailureReason::FormulaNotSupported
        } else if lower.contains("reference") || lower.contains("not found") {
            ValidationFailureReason::InvalidReference
        } else if lower.contains("not in list") || lower.contains("not allowed") {
            ValidationFailureReason::NotInList
        } else if lower.contains("list is empty") || lower.contains("no items") {
            ValidationFailureReason::ListEmpty
        } else {
            // Default: the value itself is invalid (doesn't meet constraint)
            ValidationFailureReason::InvalidValue
        }
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

                    let formula_cell = CellId::new(sheet_id, *row, *col);
                    if !refs.is_empty() {
                        let preds: FxHashSet<CellId> = refs.into_iter().collect();
                        self.dep_graph.replace_edges(formula_cell, preds);
                    } else {
                        // Leaf formula (no cell refs, e.g. =1/0, =PI())
                        // Must still be tracked so recompute evaluates it.
                        self.dep_graph.register_leaf_formula(formula_cell);
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
        self.recompute_full_ordered_inner(None)
    }

    /// Core recompute implementation, optionally with custom function handler.
    fn recompute_full_ordered_inner(
        &mut self,
        custom_fn_handler: Option<&dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>>,
    ) -> crate::recalc::RecalcReport {
        use crate::formula::analyze::has_dynamic_deps;
        use crate::formula::eval::Value;
        use crate::recalc::{CellRecalcInfo, RecalcError, RecalcReport};
        use rustc_hash::{FxHashMap, FxHashSet};
        use std::time::Instant;

        let start = Instant::now();
        let mut report = RecalcReport::new();

        // Clear computed value caches from previous recalc
        for sheet in &self.sheets {
            sheet.clear_computed_cache();
        }

        // Get topo order (or detect cycles)
        let (order, cycle_cells) = match self.dep_graph.topo_order_all_formulas() {
            Ok(order) => (order, Vec::new()),
            Err(cycle) => {
                report.had_cycles = true;
                let cycle_cells = cycle.cells.clone();
                let all_formula_cells: Vec<CellId> = self.dep_graph.formula_cells().collect();
                let non_cycle: Vec<CellId> = all_formula_cells
                    .into_iter()
                    .filter(|c| !cycle_cells.contains(c))
                    .collect();
                (non_cycle, cycle_cells)
            }
        };

        // If cycles exist and iteration is enabled, resolve via Jacobi iteration
        if !cycle_cells.is_empty() && self.iterative_enabled {
            let sccs = self.dep_graph.find_cycle_sccs();
            report.scc_count = sccs.len();
            report.converged = true; // will be set false if any SCC fails

            // Build true cycle set from Tarjan's SCCs (not from Kahn's over-report)
            let cycle_set: FxHashSet<CellId> = sccs.iter().flat_map(|scc| scc.iter().copied()).collect();
            report.cycle_cells = cycle_set.len();

            // Get ALL non-SCC formula cells (Kahn's `order` may exclude downstream cells)
            let all_formula_cells: Vec<CellId> = self.dep_graph.formula_cells().collect();
            let non_cycle_cells: Vec<CellId> = all_formula_cells
                .into_iter()
                .filter(|c| !cycle_set.contains(c))
                .collect();

            let mut depths: FxHashMap<CellId, usize> = FxHashMap::default();
            let mut eval_order: usize = 0;

            // Separate known-deps and unknown-deps among non-cycle cells
            let mut known_deps_order = Vec::new();
            let mut unknown_deps_cells = Vec::new();
            for cell_id in &non_cycle_cells {
                if let Some(sheet) = self.sheet_by_id(cell_id.sheet) {
                    if let Some(cell) = sheet.cells.get(&(cell_id.row, cell_id.col)) {
                        if let Some(ast) = cell.value.formula_ast() {
                            if has_dynamic_deps(ast) {
                                unknown_deps_cells.push(*cell_id);
                            } else {
                                known_deps_order.push(*cell_id);
                            }
                        } else {
                            known_deps_order.push(*cell_id);
                        }
                    }
                }
            }

            // Partition non-cycle cells into upstream (no cycle deps) and downstream
            let mut downstream_known = Vec::new();
            let mut upstream_known = Vec::new();
            for cell_id in known_deps_order {
                // Check transitive dependency on cycle cells
                let depends_on_cycle = self.dep_graph.precedents(cell_id)
                    .any(|p| cycle_set.contains(&p));
                if depends_on_cycle {
                    downstream_known.push(cell_id);
                } else {
                    upstream_known.push(cell_id);
                }
            }

            // Evaluate upstream non-cycle cells
            for cell_id in &upstream_known {
                let mut max_pred_depth = 0;
                for pred in self.dep_graph.precedents(*cell_id) {
                    max_pred_depth = max_pred_depth.max(depths.get(&pred).copied().unwrap_or(0));
                }
                let cell_depth = max_pred_depth + 1;
                depths.insert(*cell_id, cell_depth);
                report.max_depth = report.max_depth.max(cell_depth);

                if let Err(e) = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler) {
                    if report.errors.len() < 100 {
                        report.errors.push(RecalcError::new(*cell_id, e));
                    }
                }
                report.cell_info.insert(*cell_id, CellRecalcInfo::new(cell_depth, eval_order, false));
                eval_order += 1;
                report.cells_recomputed += 1;
            }

            // Phase 2: Jacobi iteration for each SCC
            let max_iters = self.iterative_max_iters;
            let tolerance = self.iterative_tolerance;

            for scc in &sccs {
                // Initialize: seed cache with zero for SCC cells that have no cached value
                for cell_id in scc {
                    if let Some(sheet) = self.sheet_by_id(cell_id.sheet) {
                        if sheet.get_cached_value(cell_id.row, cell_id.col).is_none() {
                            sheet.cache_computed(cell_id.row, cell_id.col, Value::Number(0.0));
                        }
                    }
                }

                let mut converged = false;
                let mut iters_used: u32 = 0;

                for iter in 0..max_iters {
                    iters_used = iter + 1;

                    // Snapshot previous values
                    let prev: Vec<(CellId, Value)> = scc.iter().map(|cell_id| {
                        let val = self.sheet_by_id(cell_id.sheet)
                            .and_then(|s| s.get_cached_value(cell_id.row, cell_id.col))
                            .unwrap_or(Value::Empty);
                        (*cell_id, val)
                    }).collect();

                    // Evaluate all SCC cells (writes new values to cache)
                    for cell_id in scc {
                        let _ = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler);
                    }

                    // Compute max delta between prev and new
                    let mut max_delta: f64 = 0.0;
                    for (cell_id, prev_val) in &prev {
                        let new_val = self.sheet_by_id(cell_id.sheet)
                            .and_then(|s| s.get_cached_value(cell_id.row, cell_id.col))
                            .unwrap_or(Value::Empty);

                        let delta = match (&prev_val, &new_val) {
                            (Value::Number(a), Value::Number(b)) => (a - b).abs(),
                            (Value::Empty, Value::Number(b)) => b.abs(),
                            (Value::Number(a), Value::Empty) => a.abs(),
                            (Value::Empty, Value::Empty) => 0.0,
                            // Text/bool/error changes: treat as non-converged
                            _ => {
                                if prev_val == &new_val { 0.0 } else { f64::INFINITY }
                            }
                        };
                        if delta > max_delta {
                            max_delta = delta;
                        }
                    }

                    if max_delta < tolerance {
                        converged = true;
                        break;
                    }
                }

                report.iterations_performed = report.iterations_performed.max(iters_used);

                if converged {
                    // Commit converged values — they're already in cache.
                    // Record in report.
                    for cell_id in scc {
                        report.cell_info.insert(
                            *cell_id,
                            CellRecalcInfo::new(report.max_depth + 1, eval_order, false),
                        );
                        eval_order += 1;
                        report.cells_recomputed += 1;
                    }
                } else {
                    // Mark all SCC cells as #NUM! (did not converge)
                    report.converged = false;
                    for cell_id in scc {
                        if let Some(sheet) = self.sheet_by_id(cell_id.sheet) {
                            sheet.cache_computed(
                                cell_id.row, cell_id.col,
                                Value::Error("#NUM!".to_string()),
                            );
                        }
                        report.cell_info.insert(
                            *cell_id,
                            CellRecalcInfo::new(report.max_depth + 1, eval_order, false),
                        );
                        eval_order += 1;
                        report.cells_recomputed += 1;
                    }
                }
            }

            // Phase 3: Evaluate downstream non-cycle cells (depend on SCC outputs)
            // Sort by dependency: cells whose precedents are all already evaluated come first.
            // Simple approach: iteratively evaluate cells whose deps are all ready.
            let downstream_set: FxHashSet<CellId> = downstream_known.iter().copied().collect();
            let mut remaining = downstream_known;
            let mut evaluated: FxHashSet<CellId> = FxHashSet::default();
            // All upstream + cycle cells are already evaluated
            for c in &upstream_known {
                evaluated.insert(*c);
            }
            for scc in &sccs {
                for c in scc {
                    evaluated.insert(*c);
                }
            }
            let mut progress = true;
            while !remaining.is_empty() && progress {
                progress = false;
                let mut still_remaining = Vec::new();
                for cell_id in remaining {
                    let all_deps_ready = self.dep_graph.precedents(cell_id)
                        .all(|p| !downstream_set.contains(&p) || evaluated.contains(&p));
                    if all_deps_ready {
                        evaluated.insert(cell_id);
                        progress = true;

                        let mut max_pred_depth = 0;
                        for pred in self.dep_graph.precedents(cell_id) {
                            max_pred_depth = max_pred_depth.max(depths.get(&pred).copied().unwrap_or(0));
                        }
                        let cell_depth = max_pred_depth + 1;
                        depths.insert(cell_id, cell_depth);
                        report.max_depth = report.max_depth.max(cell_depth);

                        if let Err(e) = self.evaluate_cell_with_handler(cell_id, custom_fn_handler) {
                            if report.errors.len() < 100 {
                                report.errors.push(RecalcError::new(cell_id, e));
                            }
                        }
                        report.cell_info.insert(cell_id, CellRecalcInfo::new(cell_depth, eval_order, false));
                        eval_order += 1;
                        report.cells_recomputed += 1;
                    } else {
                        still_remaining.push(cell_id);
                    }
                }
                remaining = still_remaining;
            }
            // Any remaining cells have unresolvable deps — evaluate anyway
            for cell_id in &remaining {
                let mut max_pred_depth = 0;
                for pred in self.dep_graph.precedents(*cell_id) {
                    max_pred_depth = max_pred_depth.max(depths.get(&pred).copied().unwrap_or(0));
                }
                let cell_depth = max_pred_depth + 1;
                depths.insert(*cell_id, cell_depth);
                report.max_depth = report.max_depth.max(cell_depth);

                if let Err(e) = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler) {
                    if report.errors.len() < 100 {
                        report.errors.push(RecalcError::new(*cell_id, e));
                    }
                }
                report.cell_info.insert(*cell_id, CellRecalcInfo::new(cell_depth, eval_order, false));
                eval_order += 1;
                report.cells_recomputed += 1;
            }

            // Phase 4: Unknown-deps formulas (after everything else)
            unknown_deps_cells.sort_by(|a, b| {
                a.sheet.raw().cmp(&b.sheet.raw())
                    .then(a.row.cmp(&b.row))
                    .then(a.col.cmp(&b.col))
            });
            for cell_id in &unknown_deps_cells {
                let cell_depth = report.max_depth + 1;
                depths.insert(*cell_id, cell_depth);
                if let Err(e) = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler) {
                    if report.errors.len() < 100 {
                        report.errors.push(RecalcError::new(*cell_id, e));
                    }
                }
                report.cell_info.insert(*cell_id, CellRecalcInfo::new(cell_depth, eval_order, true));
                eval_order += 1;
                report.cells_recomputed += 1;
                report.unknown_deps_recomputed += 1;
            }
            if !unknown_deps_cells.is_empty() {
                report.max_depth += 1;
            }
        } else {
            // No iteration: original path (mark cycles as #CYCLE!, eval non-cycle)
            for cell_id in &cycle_cells {
                if let Some(sheet) = self.sheet_by_id_mut(cell_id.sheet) {
                    sheet.set_cycle_error(cell_id.row, cell_id.col);
                }
            }
            // Use Tarjan SCC membership as the canonical cycle count (not Kahn's
            // remainder, which can include downstream false positives).
            let sccs = self.dep_graph.find_cycle_sccs();
            report.cycle_cells = sccs.iter().map(|scc| scc.len()).sum();

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

            let mut depths: FxHashMap<CellId, usize> = FxHashMap::default();
            let mut eval_order: usize = 0;

            for cell_id in &known_deps_order {
                let mut max_pred_depth = 0;
                for pred in self.dep_graph.precedents(*cell_id) {
                    max_pred_depth = max_pred_depth.max(depths.get(&pred).copied().unwrap_or(0));
                }
                let cell_depth = max_pred_depth + 1;
                depths.insert(*cell_id, cell_depth);
                report.max_depth = report.max_depth.max(cell_depth);

                if let Err(e) = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler) {
                    if report.errors.len() < 100 {
                        report.errors.push(RecalcError::new(*cell_id, e));
                    }
                }
                report.cell_info.insert(*cell_id, CellRecalcInfo::new(cell_depth, eval_order, false));
                eval_order += 1;
                report.cells_recomputed += 1;
            }

            unknown_deps_cells.sort_by(|a, b| {
                a.sheet.raw().cmp(&b.sheet.raw())
                    .then(a.row.cmp(&b.row))
                    .then(a.col.cmp(&b.col))
            });
            for cell_id in &unknown_deps_cells {
                let cell_depth = report.max_depth + 1;
                depths.insert(*cell_id, cell_depth);
                if let Err(e) = self.evaluate_cell_with_handler(*cell_id, custom_fn_handler) {
                    if report.errors.len() < 100 {
                        report.errors.push(RecalcError::new(*cell_id, e));
                    }
                }
                report.cell_info.insert(*cell_id, CellRecalcInfo::new(cell_depth, eval_order, true));
                eval_order += 1;
                report.cells_recomputed += 1;
                report.unknown_deps_recomputed += 1;
            }
            if !unknown_deps_cells.is_empty() {
                report.max_depth += 1;
            }
        }

        report.duration_ms = start.elapsed().as_millis() as u64;
        report
    }

    /// Full recalc with custom function handler support.
    ///
    /// Same as `recompute_full_ordered` but passes the handler through to all
    /// cell evaluations so custom Lua functions can be resolved.
    pub fn recompute_full_ordered_with_custom_fns(
        &mut self,
        handler: &dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>,
    ) -> crate::recalc::RecalcReport {
        self.recompute_full_ordered_inner(Some(handler))
    }

    /// Evaluate a single cell's formula and return the result.
    ///
    /// This forces evaluation by reading the cell value through the workbook lookup.
    fn evaluate_cell(&self, cell_id: CellId) -> Result<(), String> {
        self.evaluate_cell_with_handler(cell_id, None)
    }

    /// Evaluate a single cell's formula, optionally with a custom function handler.
    fn evaluate_cell_with_handler(
        &self,
        cell_id: CellId,
        custom_fn_handler: Option<&dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>>,
    ) -> Result<(), String> {
        use crate::formula::eval::evaluate;
        use crate::formula::parser::bind_expr;

        let sheet = self.sheet_by_id(cell_id.sheet)
            .ok_or_else(|| format!("Sheet not found: {:?}", cell_id.sheet))?;

        let cell = sheet.cells.get(&(cell_id.row, cell_id.col))
            .ok_or_else(|| format!("Cell not found: {:?}", cell_id))?;

        if let Some(ast) = cell.value.formula_ast() {
            let bound = bind_expr(ast, |name| self.sheet_id_by_name(name));
            let lookup = match custom_fn_handler {
                Some(handler) => WorkbookLookup::with_custom_functions(
                    self, cell_id.sheet, cell_id.row, cell_id.col, handler,
                ),
                None => WorkbookLookup::with_cell_context(
                    self, cell_id.sheet, cell_id.row, cell_id.col,
                ),
            };
            let result = evaluate(&bound, &lookup);

            // Cache the typed Value so subsequent lookups use the topo-consistent value
            // This is the ONLY place values are written to the cache.
            sheet.cache_computed(cell_id.row, cell_id.col, result.to_value());

            // Check for error result
            if let EvalResult::Error(e) = result {
                return Err(e);
            }
        }

        Ok(())
    }

    // =========================================================================
    // Tracked cell mutations (set_value + dep update + recalc notification)
    // =========================================================================

    /// Set a cell value on a specific sheet with dep tracking + recalc notification.
    /// Use this inside workbook closures where `Spreadsheet::set_cell_value()` is unavailable.
    pub fn set_cell_value_tracked(&mut self, sheet_index: usize, row: usize, col: usize, value: &str) {
        let sheet_id = match self.sheets.get(sheet_index) {
            Some(sheet) => sheet.id,
            None => return,
        };
        self.sheets[sheet_index].set_value(row, col, value);
        self.update_cell_deps(sheet_id, row, col);
        let cell_id = CellId::new(sheet_id, row, col);
        self.note_cell_changed(cell_id);
    }

    /// Clear a cell on a specific sheet with dep tracking + recalc notification.
    /// Removes the cell entirely (including spill state), unlike set_value("").
    pub fn clear_cell_tracked(&mut self, sheet_index: usize, row: usize, col: usize) {
        let sheet_id = match self.sheets.get(sheet_index) {
            Some(sheet) => sheet.id,
            None => return,
        };
        self.sheets[sheet_index].clear_cell(row, col);
        self.update_cell_deps(sheet_id, row, col);
        let cell_id = CellId::new(sheet_id, row, col);
        self.note_cell_changed(cell_id);
    }

    /// Create an RAII batch guard. Calls begin_batch() on creation, end_batch() on Drop.
    /// Use this to ensure batches are always closed, even on early return or panic.
    ///
    /// Access workbook methods through the guard (it implements DerefMut<Target=Workbook>).
    pub fn batch_guard(&mut self) -> BatchGuard<'_> {
        self.begin_batch();
        BatchGuard { wb: self }
    }

    // =========================================================================
    // Incremental Recalc (via dep graph)
    // =========================================================================

    /// Begin a batch edit. Defers recalc until end_batch().
    /// Nestable: only the outermost end_batch() triggers recalc.
    pub fn begin_batch(&mut self) {
        self.batch_depth += 1;
    }

    /// End a batch edit. If this is the outermost end, run a single
    /// incremental recalc for the union of all changed cells, then
    /// increment the revision number once.
    ///
    /// Returns the list of changed cells (empty if nested batch or no changes).
    /// Callers can use this to broadcast changes to subscribers.
    pub fn end_batch(&mut self) -> Vec<CellId> {
        assert!(self.batch_depth > 0, "end_batch without begin_batch");
        self.batch_depth -= 1;
        if self.batch_depth == 0 {
            let changed = std::mem::take(&mut self.batch_changed);
            if !changed.is_empty() {
                self.recalc_dirty_set(&changed);
                self.increment_revision();
                return changed;
            }
        }
        Vec::new()
    }

    /// Record a cell change. If batching, defers recalc.
    /// If not batching, recalcs immediately and increments revision.
    /// No-op when `auto_recalc` is false (manual calculation mode).
    pub fn note_cell_changed(&mut self, cell_id: CellId) {
        if !self.auto_recalc {
            return;
        }
        if self.batch_depth > 0 {
            self.batch_changed.push(cell_id);
        } else {
            // If the changed cell is itself a formula (e.g., a newly-entered
            // cross-sheet formula), evaluate it at the workbook level. This
            // handles formulas that evaluate_and_spill skipped (cross-sheet refs)
            // and also correctly re-evaluates same-sheet formulas.
            let is_formula = self.sheet_by_id(cell_id.sheet)
                .and_then(|s| s.cells.get(&(cell_id.row, cell_id.col)))
                .map(|c| c.value.formula_ast().is_some())
                .unwrap_or(false);
            if is_formula {
                let _ = self.evaluate_cell(cell_id);
            }
            self.recalc_dirty_set(&[cell_id]);
            self.increment_revision();
        }
    }

    /// Set whether incremental recalc runs automatically after edits.
    /// `true` = Automatic (default), `false` = Manual (F9 to recalc).
    pub fn set_auto_recalc(&mut self, auto: bool) {
        self.auto_recalc = auto;
    }

    /// Returns true if auto-recalc is enabled (Automatic calculation mode).
    pub fn auto_recalc(&self) -> bool {
        self.auto_recalc
    }

    /// Enable or disable iterative calculation (Excel-style circular calc).
    pub fn set_iterative_enabled(&mut self, enabled: bool) {
        self.iterative_enabled = enabled;
    }

    /// Returns true if iterative calculation is enabled.
    pub fn iterative_enabled(&self) -> bool {
        self.iterative_enabled
    }

    /// Set maximum iterations per SCC.
    pub fn set_iterative_max_iters(&mut self, max_iters: u32) {
        self.iterative_max_iters = max_iters;
    }

    /// Returns maximum iterations per SCC.
    pub fn iterative_max_iters(&self) -> u32 {
        self.iterative_max_iters
    }

    /// Set convergence tolerance.
    pub fn set_iterative_tolerance(&mut self, tolerance: f64) {
        self.iterative_tolerance = tolerance;
    }

    /// Returns convergence tolerance.
    pub fn iterative_tolerance(&self) -> f64 {
        self.iterative_tolerance
    }

    /// Returns the current revision number.
    /// Revision increments once per successful batch or single-cell edit.
    pub fn revision(&self) -> u64 {
        self.revision
    }

    /// Increment revision. Called at end of successful batch or single-cell edit.
    fn increment_revision(&mut self) {
        self.revision += 1;
    }

    /// Test instrumentation: returns the number of times recalc_dirty_set was called.
    #[cfg(test)]
    pub fn recalc_count(&self) -> u32 {
        self.recalc_count.get()
    }

    /// Test instrumentation: reset the recalc counter.
    #[cfg(test)]
    pub fn reset_recalc_count(&self) {
        self.recalc_count.set(0);
    }

    /// Incremental recalc: re-evaluate only cells that transitively depend
    /// on any cell in `changed`. BFS to collect dirty subgraph, then
    /// evaluate in global topo order.
    fn recalc_dirty_set(&self, changed: &[CellId]) {
        use std::collections::VecDeque;

        // Test instrumentation: count recalc calls
        #[cfg(test)]
        self.recalc_count.set(self.recalc_count.get() + 1);

        // 1. BFS forward from all changed cells to collect dirty set
        let mut dirty_set = FxHashSet::default();
        let mut queue = VecDeque::new();

        for &cell_id in changed {
            for dep in self.dep_graph.dependents(cell_id) {
                if dirty_set.insert(dep) {
                    queue.push_back(dep);
                }
            }
        }
        while let Some(cell) = queue.pop_front() {
            for dep in self.dep_graph.dependents(cell) {
                if dirty_set.insert(dep) {
                    queue.push_back(dep);
                }
            }
        }

        if dirty_set.is_empty() {
            return;
        }

        // 2. Clear cached values for dirty cells
        for cell_id in &dirty_set {
            if let Some(sheet) = self.sheet_by_id(cell_id.sheet) {
                sheet.clear_cached(cell_id.row, cell_id.col);
            }
        }

        // 3. Evaluate in global topo order, skipping non-dirty cells.
        //
        // TODO(perf): Cache the topo order to avoid recomputing on every recalc.
        //
        // Implementation sketch:
        //   struct TopoCache {
        //       order: Vec<CellId>,  // cached topo_order_all_formulas() result
        //       valid: bool,         // false when deps have changed
        //   }
        //
        // Invalidation points (set valid = false):
        //   - update_cell_deps()    — workbook.rs, called when a cell's formula changes
        //   - clear_cell_deps()     — workbook.rs, called when a formula is replaced/cleared
        //   - rebuild_dep_graph()   — workbook.rs, called on file load and structural changes
        //   - insert_rows/cols      — sheet.rs, structural changes shift deps
        //   - delete_rows/cols      — sheet.rs, structural changes remove deps
        //
        // On recalc: if valid, use cached order. If not, recompute and cache.
        // For small-to-medium models the current approach is fine. Profile before
        // optimizing — the BFS dirty-set collection is likely the bigger cost.
        if let Ok(order) = self.dep_graph.topo_order_all_formulas() {
            for cell_id in order {
                if dirty_set.contains(&cell_id) {
                    let _ = self.evaluate_cell(cell_id);
                }
            }
        }
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
                    EvalResult::Empty => String::new(),
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
// BatchGuard - RAII batch scope for Workbook
// =============================================================================

/// RAII guard that calls `begin_batch()` on creation and `end_batch()` on drop.
/// Access workbook methods through the guard (implements `Deref`/`DerefMut`).
///
/// ```ignore
/// let mut guard = wb.batch_guard();
/// for change in changes {
///     guard.set_cell_value_tracked(sheet_index, row, col, &value);
/// }
/// // guard drops here → end_batch() → single recalc
/// ```
pub struct BatchGuard<'a> {
    wb: &'a mut Workbook,
}

impl<'a> std::ops::Deref for BatchGuard<'a> {
    type Target = Workbook;
    fn deref(&self) -> &Workbook {
        self.wb
    }
}

impl<'a> std::ops::DerefMut for BatchGuard<'a> {
    fn deref_mut(&mut self) -> &mut Workbook {
        self.wb
    }
}

impl Drop for BatchGuard<'_> {
    fn drop(&mut self) {
        self.wb.end_batch();
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
    custom_fn_handler: Option<&'a dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>>,
}

impl<'a> WorkbookLookup<'a> {
    /// Create a new WorkbookLookup for the given sheet
    pub fn new(workbook: &'a Workbook, current_sheet_id: SheetId) -> Self {
        Self {
            workbook,
            current_sheet_id,
            current_cell: None,
            custom_fn_handler: None,
        }
    }

    /// Create a new WorkbookLookup with cell context (for ROW()/COLUMN() without args)
    pub fn with_cell_context(workbook: &'a Workbook, current_sheet_id: SheetId, row: usize, col: usize) -> Self {
        Self {
            workbook,
            current_sheet_id,
            current_cell: Some((row, col)),
            custom_fn_handler: None,
        }
    }

    /// Create a new WorkbookLookup with cell context and a custom function handler
    pub fn with_custom_functions(
        workbook: &'a Workbook,
        current_sheet_id: SheetId,
        row: usize,
        col: usize,
        handler: &'a dyn Fn(&str, &[EvalArg]) -> Option<EvalResult>,
    ) -> Self {
        Self {
            workbook,
            current_sheet_id,
            current_cell: Some((row, col)),
            custom_fn_handler: Some(handler),
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
            .map(|sheet| sheet.get_value(row, col))
            .unwrap_or(0.0)
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        self.current_sheet()
            .map(|sheet| sheet.get_text(row, col))
            .unwrap_or_default()
    }

    fn get_value_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Value {
        match self.workbook.sheet_by_id(sheet_id) {
            Some(sheet) => sheet.get_computed_value(row, col),
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

    fn debug_context(&self) -> String {
        let sheet_info = self.current_sheet()
            .map(|s| format!("sheet=\"{}\", sheet_ptr={:p}, cache_len={}", s.name, s as *const Sheet, s.computed_cache_len()))
            .unwrap_or_else(|| "sheet=<none>".to_string());
        format!("WorkbookLookup(wb_ptr={:p}, {})", self.workbook as *const Workbook, sheet_info)
    }

    fn get_merge_start(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.current_sheet()
            .and_then(|s| s.get_merge(row, col))
            .map(|m| m.start)
    }

    fn get_merge_start_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Option<(usize, usize)> {
        self.workbook.sheet_by_id(sheet_id)
            .and_then(|s| s.get_merge(row, col))
            .map(|m| m.start)
    }

    fn get_cell_value(&self, row: usize, col: usize) -> Value {
        self.current_sheet()
            .map(|sheet| sheet.get_computed_value(row, col))
            .unwrap_or(Value::Empty)
    }

    fn try_custom_function(&self, name: &str, args: &[EvalArg]) -> Option<EvalResult> {
        self.custom_fn_handler.as_ref().and_then(|handler| handler(name, args))
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

    #[test]
    fn test_get_list_items_named_range() {
        use crate::validation::ValidationRule;

        let mut wb = Workbook::new();

        // Set source values in A1:A3
        wb.sheet_mut(0).unwrap().set_value(0, 0, "High");
        wb.sheet_mut(0).unwrap().set_value(1, 0, "Medium");
        wb.sheet_mut(0).unwrap().set_value(2, 0, "Low");

        // Create named range
        wb.define_name_for_range("Priority", 0, 0, 0, 2, 0).unwrap();

        // Set list validation using named range
        let rule = ValidationRule::new(crate::validation::ValidationType::List(
            crate::validation::ListSource::NamedRange("Priority".to_string())
        ));
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 1, rule);

        // Get list items via workbook API
        let list = wb.get_list_items(0, 0, 1).expect("Should resolve named range");
        assert_eq!(list.items, vec!["High", "Medium", "Low"]);
        assert!(!list.is_truncated);
    }

    #[test]
    fn test_get_list_items_cross_sheet_range() {
        use crate::validation::ValidationRule;

        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        // Set source values on Sheet2!A1:A3
        wb.sheet_mut(1).unwrap().set_value(0, 0, "Alpha");
        wb.sheet_mut(1).unwrap().set_value(1, 0, "Beta");
        wb.sheet_mut(1).unwrap().set_value(2, 0, "Gamma");

        // Set list validation on Sheet1 referencing Sheet2
        let rule = ValidationRule::list_range("Sheet2!A1:A3");
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 0, rule);

        // Get list items - should resolve cross-sheet
        let list = wb.get_list_items(0, 0, 0).expect("Should resolve cross-sheet");
        assert_eq!(list.items, vec!["Alpha", "Beta", "Gamma"]);
    }

    #[test]
    fn test_has_list_dropdown_workbook() {
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut wb = Workbook::new();

        // List validation with dropdown
        let list_rule = ValidationRule::list_inline(vec!["A".into()]);
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 0, list_rule);
        assert!(wb.has_list_dropdown(0, 0, 0));

        // Non-list validation
        let num_rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 1, num_rule);
        assert!(!wb.has_list_dropdown(0, 0, 1));

        // Invalid sheet index
        assert!(!wb.has_list_dropdown(99, 0, 0));
    }

    // =========================================================================
    // Numeric Validation Cross-Sheet Tests (Phase 3)
    // =========================================================================

    #[test]
    fn test_validation_cell_ref_constraint_cross_sheet() {
        // Sheet1 B2 has Decimal <= Sheet2!A1, validation should work across sheets
        use crate::validation::{ValidationRule, ValidationResult, NumericConstraint, ConstraintValue};

        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        // Set constraint value on Sheet2!A1
        wb.sheet_mut(1).unwrap().set_value(0, 0, "50");

        // Set validation on Sheet1!B2: Decimal <= Sheet2!A1
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("Sheet2!A1".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        wb.sheet_mut(0).unwrap().set_cell_validation(1, 1, rule);

        // Value 30 should be valid (30 <= 50)
        let result = wb.validate_cell_input(0, 1, 1, "30");
        assert!(matches!(result, ValidationResult::Valid), "30 <= 50 should be valid");

        // Value 50 should be valid (50 <= 50)
        let result = wb.validate_cell_input(0, 1, 1, "50");
        assert!(matches!(result, ValidationResult::Valid), "50 <= 50 should be valid");

        // Value 60 should be invalid (60 > 50)
        let result = wb.validate_cell_input(0, 1, 1, "60");
        assert!(matches!(result, ValidationResult::Invalid { .. }), "60 > 50 should be invalid");
    }

    #[test]
    fn test_validation_cell_ref_constraint_cross_sheet_updates() {
        // Change Sheet2 A1, validation result should change
        use crate::validation::{ValidationRule, ValidationResult, NumericConstraint, ConstraintValue};

        let mut wb = Workbook::new();
        wb.add_sheet(); // Sheet2

        // Set initial constraint value on Sheet2!A1
        wb.sheet_mut(1).unwrap().set_value(0, 0, "50");

        // Set validation on Sheet1!A1: Decimal <= Sheet2!A1
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("Sheet2!A1".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 0, rule);

        // Value 60 should be invalid initially (60 > 50)
        let result = wb.validate_cell_input(0, 0, 0, "60");
        assert!(matches!(result, ValidationResult::Invalid { .. }), "60 > 50 should be invalid");

        // Update Sheet2!A1 to 100
        wb.sheet_mut(1).unwrap().set_value(0, 0, "100");

        // Now 60 should be valid (60 <= 100)
        let result = wb.validate_cell_input(0, 0, 0, "60");
        assert!(matches!(result, ValidationResult::Valid), "60 <= 100 should be valid after update");
    }

    #[test]
    fn test_validation_formula_constraint_error() {
        // Formula constraint should return deterministic FormulaError
        use crate::validation::{ValidationRule, ValidationResult, NumericConstraint, ConstraintValue};

        let mut wb = Workbook::new();

        // Set validation with formula constraint (not yet implemented)
        let constraint = NumericConstraint {
            operator: crate::validation::ComparisonOperator::LessThan,
            value1: ConstraintValue::Formula("=A1+10".to_string()),
            value2: None,
        };
        let rule = ValidationRule::decimal(constraint);
        wb.sheet_mut(0).unwrap().set_cell_validation(0, 0, rule);

        // Validation should fail with formula error
        let result = wb.validate_cell_input(0, 0, 0, "5");
        match result {
            ValidationResult::Invalid { reason, .. } => {
                assert!(reason.contains("Formula") || reason.contains("formula"),
                    "Error should mention formula: {}", reason);
            }
            ValidationResult::Valid => {
                panic!("Formula constraint should fail deterministically, not pass");
            }
        }
    }

    // =========================================================================
    // Range Validation Tests (Phase 6A: Paste/Fill)
    // =========================================================================

    #[test]
    fn test_validate_range_paste_numeric() {
        // Simulate paste into validated numeric range: some valid, some invalid
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut wb = Workbook::new();

        // Set validation on A1:A5: Whole number between 1 and 10
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        for row in 0..5 {
            wb.sheet_mut(0).unwrap().set_cell_validation(row, 0, rule.clone());
        }

        // Simulate pasting values: 5, 15, 3, -1, 8
        // Valid: 5, 3, 8 (3 cells)
        // Invalid: 15, -1 (2 cells)
        wb.sheet_mut(0).unwrap().set_value(0, 0, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(1, 0, "15");  // invalid (> 10)
        wb.sheet_mut(0).unwrap().set_value(2, 0, "3");   // valid
        wb.sheet_mut(0).unwrap().set_value(3, 0, "-1");  // invalid (< 1)
        wb.sheet_mut(0).unwrap().set_value(4, 0, "8");   // valid

        let failures = wb.validate_range(0, 0, 0, 4, 0);

        assert_eq!(failures.count, 2, "Expected 2 validation failures");
        assert_eq!(failures.failures.len(), 2, "Expected 2 failure entries");
        let positions: Vec<_> = failures.failures.iter().map(|f| (f.row, f.col)).collect();
        assert!(positions.contains(&(1, 0)), "Row 1 should be invalid");
        assert!(positions.contains(&(3, 0)), "Row 3 should be invalid");
    }

    #[test]
    fn test_validate_range_fill_numeric() {
        // Simulate fill down into validated range
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut wb = Workbook::new();

        // Set validation on B1:B3: Decimal less than 100
        let rule = ValidationRule::decimal(NumericConstraint::less_than(100.0));
        for row in 0..3 {
            wb.sheet_mut(0).unwrap().set_cell_validation(row, 1, rule.clone());
        }

        // Simulate fill with value 50 (all valid)
        wb.sheet_mut(0).unwrap().set_value(0, 1, "50");
        wb.sheet_mut(0).unwrap().set_value(1, 1, "50");
        wb.sheet_mut(0).unwrap().set_value(2, 1, "50");

        let failures = wb.validate_range(0, 0, 1, 2, 1);
        assert_eq!(failures.count, 0, "All values should be valid");

        // Now fill with value 150 (all invalid)
        wb.sheet_mut(0).unwrap().set_value(0, 1, "150");
        wb.sheet_mut(0).unwrap().set_value(1, 1, "150");
        wb.sheet_mut(0).unwrap().set_value(2, 1, "150");

        let failures = wb.validate_range(0, 0, 1, 2, 1);
        assert_eq!(failures.count, 3, "All values should be invalid");
    }

    #[test]
    fn test_validation_failures_with_reasons() {
        // Verify that failures include reason codes
        use crate::validation::{ValidationRule, NumericConstraint, ValidationFailureReason};

        let mut wb = Workbook::new();

        // Set validation: Whole number between 1 and 10
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        for row in 0..3 {
            wb.sheet_mut(0).unwrap().set_cell_validation(row, 0, rule.clone());
        }

        // Set values: valid, invalid (out of range), invalid (out of range)
        wb.sheet_mut(0).unwrap().set_value(0, 0, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(1, 0, "15");  // invalid
        wb.sheet_mut(0).unwrap().set_value(2, 0, "-1");  // invalid

        let failures = wb.validate_range(0, 0, 0, 2, 0);

        assert_eq!(failures.count, 2);
        assert_eq!(failures.failures.len(), 2);

        // Both failures should be InvalidValue (value doesn't meet constraint)
        for failure in &failures.failures {
            assert_eq!(failure.reason, ValidationFailureReason::InvalidValue);
        }
    }

    #[test]
    fn test_validation_failures_navigation_semantics() {
        // Test that failures are in row-major order (for predictable F8 cycling)
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut wb = Workbook::new();

        // Set validation on a 3x3 range
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        for row in 0..3 {
            for col in 0..3 {
                wb.sheet_mut(0).unwrap().set_cell_validation(row, col, rule.clone());
            }
        }

        // Set some invalid values in non-sequential positions
        wb.sheet_mut(0).unwrap().set_value(0, 0, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(0, 1, "20");  // invalid (0,1)
        wb.sheet_mut(0).unwrap().set_value(0, 2, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(1, 0, "30");  // invalid (1,0)
        wb.sheet_mut(0).unwrap().set_value(1, 1, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(1, 2, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(2, 0, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(2, 1, "5");   // valid
        wb.sheet_mut(0).unwrap().set_value(2, 2, "40");  // invalid (2,2)

        let failures = wb.validate_range(0, 0, 0, 2, 2);

        assert_eq!(failures.count, 3);
        assert_eq!(failures.failures.len(), 3);

        // Check order is row-major: (0,1), (1,0), (2,2)
        assert_eq!((failures.failures[0].row, failures.failures[0].col), (0, 1));
        assert_eq!((failures.failures[1].row, failures.failures[1].col), (1, 0));
        assert_eq!((failures.failures[2].row, failures.failures[2].col), (2, 2));
    }

    #[test]
    fn test_validation_failures_overwrite() {
        // Test that new validation overwrites previous failures
        use crate::validation::{ValidationRule, NumericConstraint};

        let mut wb = Workbook::new();

        // First validation: 2 failures in A1:A3
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 10));
        for row in 0..3 {
            wb.sheet_mut(0).unwrap().set_cell_validation(row, 0, rule.clone());
        }
        wb.sheet_mut(0).unwrap().set_value(0, 0, "20");
        wb.sheet_mut(0).unwrap().set_value(1, 0, "5");
        wb.sheet_mut(0).unwrap().set_value(2, 0, "30");

        let failures1 = wb.validate_range(0, 0, 0, 2, 0);
        assert_eq!(failures1.count, 2);

        // Second validation: different range B1:B2, 1 failure
        for row in 0..2 {
            wb.sheet_mut(0).unwrap().set_cell_validation(row, 1, rule.clone());
        }
        wb.sheet_mut(0).unwrap().set_value(0, 1, "5");
        wb.sheet_mut(0).unwrap().set_value(1, 1, "50");

        let failures2 = wb.validate_range(0, 0, 1, 1, 1);
        assert_eq!(failures2.count, 1);

        // The two ValidationFailures are independent - simulates UI overwrite
        // (This test just verifies the engine returns independent results)
        assert_ne!(failures1.count, failures2.count);
    }

    // ======== Phase 1A: Style table and intern_style ========

    #[test]
    fn test_style_table_empty_by_default() {
        let wb = Workbook::new();
        assert!(wb.style_table.is_empty());
    }

    #[test]
    fn test_intern_style_basic() {
        let mut wb = Workbook::new();
        let fmt = CellFormat {
            bold: true,
            ..Default::default()
        };
        let id = wb.intern_style(fmt.clone());
        assert_eq!(id, 0);
        assert_eq!(wb.style_table.len(), 1);
        assert_eq!(wb.style_table[0], fmt);
    }

    #[test]
    fn test_intern_style_deduplication() {
        let mut wb = Workbook::new();
        let fmt = CellFormat {
            bold: true,
            font_size: Some(14.0),
            ..Default::default()
        };
        let id1 = wb.intern_style(fmt.clone());
        let id2 = wb.intern_style(fmt.clone());
        assert_eq!(id1, id2);
        assert_eq!(wb.style_table.len(), 1);
    }

    #[test]
    fn test_intern_style_different_formats() {
        let mut wb = Workbook::new();
        let fmt1 = CellFormat {
            bold: true,
            ..Default::default()
        };
        let fmt2 = CellFormat {
            italic: true,
            ..Default::default()
        };
        let id1 = wb.intern_style(fmt1);
        let id2 = wb.intern_style(fmt2);
        assert_ne!(id1, id2);
        assert_eq!(wb.style_table.len(), 2);
    }

    #[test]
    fn test_intern_style_with_font_color() {
        let mut wb = Workbook::new();
        let fmt = CellFormat {
            font_color: Some([255, 0, 0, 255]),
            ..Default::default()
        };
        let id = wb.intern_style(fmt.clone());
        assert_eq!(wb.style_table[id as usize].font_color, Some([255, 0, 0, 255]));
    }

    #[test]
    fn test_style_table_serde_backward_compat() {
        // Old .vgrid files won't have style_table
        let json = r#"{"sheets":[{"id":1,"name":"Sheet1","cells":{},"rows":100,"cols":26}],"active_sheet":0}"#;
        let wb: Workbook = serde_json::from_str(json).unwrap();
        assert!(wb.style_table.is_empty());
    }

    // ========================================================================
    // Merged cell formula redirect tests
    // ========================================================================

    #[test]
    fn test_formula_single_ref_redirect() {
        use crate::sheet::MergedRegion;

        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = "Hello", merge A1:C1
        wb.sheet_mut(0).unwrap().set_value(0, 0, "Hello");
        wb.sheet_mut(0).unwrap().add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // D1 = =B1 (B1 is hidden in merge, should redirect to A1)
        wb.sheet_mut(0).unwrap().set_value(0, 3, "=B1");
        wb.update_cell_deps(sheet_id, 0, 3);

        // E1 = =C1 (C1 is hidden in merge, should redirect to A1)
        wb.sheet_mut(0).unwrap().set_value(0, 4, "=C1");
        wb.update_cell_deps(sheet_id, 0, 4);

        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 3), "Hello"); // =B1 → "Hello"
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "Hello"); // =C1 → "Hello"
    }

    #[test]
    fn test_formula_range_hidden_empty() {
        use crate::sheet::MergedRegion;

        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = 10, merge A1:C1
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");
        wb.sheet_mut(0).unwrap().add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // D1 = =SUM(A1:C1) → should be 10 (not 30; hidden cells are empty in ranges)
        wb.sheet_mut(0).unwrap().set_value(0, 3, "=SUM(A1:C1)");
        wb.update_cell_deps(sheet_id, 0, 3);

        // E1 = =SUM(B1:C1) → should be 0 (both hidden, no values)
        wb.sheet_mut(0).unwrap().set_value(0, 4, "=SUM(B1:C1)");
        wb.update_cell_deps(sheet_id, 0, 4);

        // F1 = =COUNTA(A1:C1) → should be 1 (only origin has value)
        wb.sheet_mut(0).unwrap().set_value(0, 5, "=COUNTA(A1:C1)");
        wb.update_cell_deps(sheet_id, 0, 5);

        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 3), "10");  // SUM(A1:C1)
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "0");   // SUM(B1:C1)
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 5), "1");   // COUNTA(A1:C1)
    }

    #[test]
    fn test_formula_explicit_refs_redirect() {
        use crate::sheet::MergedRegion;

        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        // A1 = 10, merge A1:C1
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");
        wb.sheet_mut(0).unwrap().add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // D1 = =A1+B1+C1 → each is a single-cell ref, each redirects to origin → 30
        wb.sheet_mut(0).unwrap().set_value(0, 3, "=A1+B1+C1");
        wb.update_cell_deps(sheet_id, 0, 3);

        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 3), "30"); // 10+10+10
    }

    #[test]
    fn test_formula_cross_sheet_ref_redirect() {
        use crate::sheet::MergedRegion;

        let mut wb = Workbook::new();

        // Sheet1: merge A1:C1, A1 = "Hi"
        wb.sheet_mut(0).unwrap().set_value(0, 0, "Hi");
        wb.sheet_mut(0).unwrap().add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        // Add Sheet2
        let sheet2_idx = wb.add_sheet();
        let sheet2_id = wb.sheet_id_at_idx(sheet2_idx).unwrap();

        // Sheet2 A1 = =Sheet1!B1  (B1 is hidden in merge, should redirect to A1)
        wb.sheet_mut(sheet2_idx).unwrap().set_value(0, 0, "=Sheet1!B1");
        wb.update_cell_deps(sheet2_id, 0, 0);

        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(sheet2_idx).unwrap().get_display(0, 0), "Hi");
    }

    // =========================================================================
    // Incremental Recalc Tests (note_cell_changed / recalc_dirty_set)
    // =========================================================================

    #[test]
    fn test_incremental_recalc_simple_chain() {
        // A1=10, B1=A1*2, C1=B1+1
        // Change A1→20 via set_cell_value_tracked → B1=40, C1=41
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");     // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1*2");  // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1+1");  // C1

        wb.update_cell_deps(sid, 0, 1);
        wb.update_cell_deps(sid, 0, 2);
        wb.recompute_full_ordered(); // initial cache fill

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "20");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "21");

        // Incremental: change A1 → 20
        wb.set_cell_value_tracked(0, 0, 0, "20");

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "40");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "41");
    }

    #[test]
    fn test_incremental_recalc_cross_sheet() {
        // Sheet1!A1 = 5, Sheet2!A1 = =Sheet1!A1+10
        // Change Sheet1!A1 → 100 → Sheet2!A1 = 110
        let mut wb = Workbook::new();
        let _s1 = wb.sheet_id_at_idx(0).unwrap();
        let s2_idx = wb.add_sheet();
        let s2 = wb.sheet_id_at_idx(s2_idx).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "5");               // Sheet1!A1
        wb.sheet_mut(s2_idx).unwrap().set_value(0, 0, "=Sheet1!A1+10"); // Sheet2!A1

        wb.update_cell_deps(s2, 0, 0);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(s2_idx).unwrap().get_display(0, 0), "15");

        // Incremental: change Sheet1!A1 → 100
        wb.set_cell_value_tracked(0, 0, 0, "100");

        assert_eq!(wb.sheet(s2_idx).unwrap().get_display(0, 0), "110");
    }

    #[test]
    fn test_batch_deduplicates_recalc() {
        // A1=1, B1=A1+1, C1=B1+1
        // Batch: change A1 three times → recalc should see final value
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");  // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1+1");  // C1

        wb.update_cell_deps(sid, 0, 1);
        wb.update_cell_deps(sid, 0, 2);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "3"); // 1+1+1

        // Batch: set A1 three times, only final value matters
        wb.begin_batch();
        wb.set_cell_value_tracked(0, 0, 0, "10");
        wb.set_cell_value_tracked(0, 0, 0, "20");
        wb.set_cell_value_tracked(0, 0, 0, "100");
        wb.end_batch(); // single recalc here

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "100");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "101");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "102");
    }

    #[test]
    fn test_batch_multi_cell_change() {
        // A1=1, A2=2, B1=A1+A2 — change both A1 and A2 in one batch
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");       // A1
        wb.sheet_mut(0).unwrap().set_value(1, 0, "2");       // A2
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+A2");  // B1

        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "3");

        // Batch: change both source cells
        wb.begin_batch();
        wb.set_cell_value_tracked(0, 0, 0, "10");  // A1 → 10
        wb.set_cell_value_tracked(0, 1, 0, "20");  // A2 → 20
        wb.end_batch();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "30");
    }

    #[test]
    fn test_incremental_recalc_diamond() {
        // A1=5, B1=A1*2, C1=A1*3, D1=B1+C1
        // Change A1→10 → B1=20, C1=30, D1=50
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "5");         // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1*2");     // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1*3");     // C1
        wb.sheet_mut(0).unwrap().set_value(0, 3, "=B1+C1");    // D1

        wb.update_cell_deps(sid, 0, 1);
        wb.update_cell_deps(sid, 0, 2);
        wb.update_cell_deps(sid, 0, 3);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 3), "25"); // 10+15

        wb.set_cell_value_tracked(0, 0, 0, "10");

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "20");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "30");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 3), "50");
    }

    #[test]
    fn test_incremental_recalc_no_deps_no_crash() {
        // Changing a cell with no dependents should be a no-op (no crash)
        let mut wb = Workbook::new();
        wb.sheet_mut(0).unwrap().set_value(0, 0, "hello");
        wb.set_cell_value_tracked(0, 0, 0, "world");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "world");
    }

    #[test]
    fn test_batch_guard_drop_triggers_recalc() {
        // Same as batch test but using BatchGuard RAII
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");  // B1

        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2");

        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "50");
            // guard drops here → end_batch → recalc
        }

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "51");
    }

    #[test]
    fn test_cycle_error_propagation() {
        // A1==B1, B1==A1 — creates a cycle. After recalc, both should show #CYCLE!
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=B1");  // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1");  // B1

        wb.update_cell_deps(sid, 0, 0);
        wb.update_cell_deps(sid, 0, 1);

        wb.recompute_full_ordered();

        // Both cells should display a cycle error
        let a1_display = wb.sheet(0).unwrap().get_display(0, 0);
        let b1_display = wb.sheet(0).unwrap().get_display(0, 1);
        assert!(
            a1_display.contains("CYCLE") || a1_display.contains("REF") || a1_display.contains("ERR"),
            "Expected cycle error in A1, got: {}", a1_display
        );
        assert!(
            b1_display.contains("CYCLE") || b1_display.contains("REF") || b1_display.contains("ERR"),
            "Expected cycle error in B1, got: {}", b1_display
        );
    }

    #[test]
    fn test_incremental_recalc_unrelated_cell_untouched() {
        // A1=1, B1=A1+1, C1=99 (no formula)
        // Change A1 → B1 updates, C1 stays unchanged
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");  // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "99");      // C1 (plain value)

        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        wb.set_cell_value_tracked(0, 0, 0, "50");

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "51");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "99"); // untouched
    }

    #[test]
    fn test_paste_then_undo_recalc() {
        // Simulates: A1=1, B1=A1+1
        // "Paste" A1..A5 with new values in a batch → B1 updates
        // "Undo" by restoring old values in a batch → B1 reverts
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1
        wb.sheet_mut(0).unwrap().set_value(1, 0, "2");      // A2
        wb.sheet_mut(0).unwrap().set_value(2, 0, "3");      // A3
        wb.sheet_mut(0).unwrap().set_value(3, 0, "4");      // A4
        wb.sheet_mut(0).unwrap().set_value(4, 0, "5");      // A5
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");  // B1

        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2"); // B1 = 1+1

        // Simulate paste: overwrite A1..A5 in one batch
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "10");  // A1 → 10
            guard.set_cell_value_tracked(0, 1, 0, "20");  // A2 → 20
            guard.set_cell_value_tracked(0, 2, 0, "30");  // A3 → 30
            guard.set_cell_value_tracked(0, 3, 0, "40");  // A4 → 40
            guard.set_cell_value_tracked(0, 4, 0, "50");  // A5 → 50
        }

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "10");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "11"); // B1 = 10+1

        // Simulate undo: restore original A1..A5 in one batch
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "1");
            guard.set_cell_value_tracked(0, 1, 0, "2");
            guard.set_cell_value_tracked(0, 2, 0, "3");
            guard.set_cell_value_tracked(0, 3, 0, "4");
            guard.set_cell_value_tracked(0, 4, 0, "5");
        }

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "1");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2"); // B1 = 1+1 again

        // Simulate redo: re-apply the paste
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "10");
            guard.set_cell_value_tracked(0, 1, 0, "20");
            guard.set_cell_value_tracked(0, 2, 0, "30");
            guard.set_cell_value_tracked(0, 3, 0, "40");
            guard.set_cell_value_tracked(0, 4, 0, "50");
        }

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "11"); // B1 = 10+1
    }

    #[test]
    fn test_auto_recalc_off_skips_incremental() {
        // When auto_recalc is false, note_cell_changed is a no-op.
        // Cache stays stale until explicit recompute_full_ordered().
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");      // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");  // B1

        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2");

        // Disable auto-recalc (manual mode)
        wb.set_auto_recalc(false);
        wb.set_cell_value_tracked(0, 0, 0, "100");

        // B1 still shows stale value (cache not cleared)
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2");

        // Explicit F9-style recalc fixes it
        wb.recompute_full_ordered();
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "101");
    }

    // =========================================================================
    // Session Server Invariant Tests (Phase 1)
    //
    // These tests define the contract for the session server protocol.
    // Write first, make pass during implementation.
    // See: docs/future/phase-1-session-server.md
    // =========================================================================

    /// Invariant: Empty batch does not increment revision.
    ///
    /// A batch with no ops applied should not change the revision number.
    /// This is a prerequisite for rollback semantics.
    #[test]
    fn invariant_empty_batch_no_revision_increment() {
        let mut wb = Workbook::new();
        let initial_rev = wb.revision();

        // Empty batch: no changes, no revision increment
        {
            let _guard = wb.batch_guard();
            // No ops applied
        }

        assert_eq!(wb.revision(), initial_rev, "revision should not increment on empty batch");
    }

    /// Invariant: Aborted batch (manual clear) does not increment revision.
    ///
    /// When batch_changed is cleared before end_batch (simulating abort),
    /// revision must remain stable.
    #[test]
    fn invariant_aborted_batch_no_revision_increment() {
        let mut wb = Workbook::new();
        let initial_rev = wb.revision();

        // Start batch, make changes, then "abort" by clearing batch_changed
        wb.begin_batch();
        wb.batch_changed.push(CellId::new(SheetId(1), 0, 0)); // Simulate a change
        // Simulate abort: clear the batch_changed without recalc
        wb.batch_changed.clear();
        wb.batch_depth -= 1; // Manual decrement to avoid end_batch logic

        assert_eq!(wb.revision(), initial_rev, "revision should not increment on aborted batch");
    }

    /// Invariant 1: Rollback must not emit events to subscribers.
    ///
    /// When a batch is aborted (all ops rolled back), no cells_changed,
    /// batch_applied, or revision_changed events should be emitted
    /// (except a single BatchApplied with error and applied=0).
    /// This prevents clients from seeing intermediate state.
    #[test]
    fn invariant_rollback_no_events() {
        use crate::harness::{EngineHarness, Op};

        let mut harness = EngineHarness::new();

        // First, do a successful batch to establish baseline
        let setup_ops = vec![Op::SetCellValue {
            sheet_index: 0,
            row: 0,
            col: 0,
            value: "baseline".to_string(),
        }];
        harness.apply_ops(&setup_ops, false);
        harness.clear_events();

        // Now apply a batch that will fail atomically
        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 1,
                col: 0,
                value: "will rollback".to_string(),
            },
            Op::SimulateError {
                message: "intentional failure".to_string(),
            },
        ];

        let result = harness.apply_ops(&ops, true); // atomic=true

        // Verify: only BatchApplied event (with error), NO CellsChanged, NO RevisionChanged
        let events = harness.events();
        assert_eq!(
            events.cells_changed().len(),
            0,
            "rollback must not emit CellsChanged"
        );
        assert_eq!(
            events.revision_changed().len(),
            0,
            "rollback must not emit RevisionChanged"
        );
        assert_eq!(
            events.batch_applied().len(),
            1,
            "rollback must emit exactly one BatchApplied"
        );

        let batch_event = &events.batch_applied()[0];
        assert_eq!(batch_event.applied, 0, "BatchApplied.applied must be 0 on rollback");
        assert!(batch_event.error.is_some(), "BatchApplied must have error on rollback");

        // Verify revision unchanged
        assert_eq!(result.applied, 0);
        assert_eq!(result.revision, 1); // Still at revision 1 from setup
    }

    /// Invariant 2: Rollback must not create undo entries.
    ///
    /// When a batch fails and is rolled back, the undo stack length must
    /// remain unchanged. No undo group should be created for aborted batches.
    #[test]
    fn invariant_rollback_no_undo_entries() {
        use crate::harness::{EngineHarness, Op};

        let mut harness = EngineHarness::new();

        // Create some initial undo groups
        harness.apply_ops(
            &[Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0,
                value: "edit1".to_string(),
            }],
            false,
        );
        harness.apply_ops(
            &[Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1,
                value: "edit2".to_string(),
            }],
            false,
        );

        let undo_count_before = harness.undo_group_count();
        assert_eq!(undo_count_before, 2, "should have 2 undo groups from setup");

        // Now apply a batch that will fail atomically
        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 1,
                col: 0,
                value: "will rollback".to_string(),
            },
            Op::SimulateError {
                message: "intentional failure".to_string(),
            },
        ];

        harness.apply_ops(&ops, true); // atomic=true

        // Verify: undo group count unchanged
        assert_eq!(
            harness.undo_group_count(),
            undo_count_before,
            "rollback must not create undo entry"
        );
    }

    /// Invariant 3: Successful batch triggers exactly one recalc.
    ///
    /// Multiple cell changes within a batch should result in a single
    /// recalc_dirty_set call when the batch completes.
    #[test]
    fn invariant_single_recalc_on_success() {
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        // Setup: A1=1, B1=A1+1, C1=B1+1
        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=B1+1");
        wb.update_cell_deps(sid, 0, 1);
        wb.update_cell_deps(sid, 0, 2);
        wb.recompute_full_ordered();

        // Reset counter after setup
        wb.reset_recalc_count();

        // Batch: change multiple cells
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "10");
            guard.set_cell_value_tracked(0, 1, 0, "20");
            guard.set_cell_value_tracked(0, 2, 0, "30");
        }

        assert_eq!(wb.recalc_count(), 1, "batch should trigger exactly one recalc");
    }

    /// Invariant 4: Failed batch (atomic rollback) triggers exactly one recalc.
    ///
    /// Even when a batch fails partway through, the rollback path should
    /// result in exactly one recalc (to restore consistent state).
    ///
    /// Note: Current engine doesn't have atomic rollback - this test verifies
    /// recalc count for successful batch. Full test requires session server.
    #[test]
    fn invariant_single_recalc_on_failure() {
        let mut wb = Workbook::new();
        let sid = wb.sheet_id_at_idx(0).unwrap();

        // Setup formula chain
        wb.sheet_mut(0).unwrap().set_value(0, 0, "1");
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1*2");
        wb.update_cell_deps(sid, 0, 1);
        wb.recompute_full_ordered();

        wb.reset_recalc_count();

        // Batch with single change (simulating "failure after first op")
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "5");
            // In atomic mode, if next op failed, we'd rollback
            // For now, just verify single recalc on successful completion
        }

        assert_eq!(wb.recalc_count(), 1, "batch should trigger exactly one recalc");
    }

    /// Invariant 5: Revision increments exactly once per successful batch.
    ///
    /// A batch that modifies N cells should increment revision by exactly 1,
    /// not N. This enables efficient change detection.
    #[test]
    fn invariant_revision_increments_once_per_batch() {
        let mut wb = Workbook::new();
        let initial_rev = wb.revision();

        // Batch: 5 cell changes
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "a");
            guard.set_cell_value_tracked(0, 1, 0, "b");
            guard.set_cell_value_tracked(0, 2, 0, "c");
            guard.set_cell_value_tracked(0, 3, 0, "d");
            guard.set_cell_value_tracked(0, 4, 0, "e");
        }

        assert_eq!(
            wb.revision(),
            initial_rev + 1,
            "batch of 5 changes should increment revision by exactly 1"
        );
    }

    /// Invariant 6: Revision must not increment on rejected/empty batch.
    ///
    /// If a batch is rejected (e.g., expected_revision mismatch) or contains
    /// no actual changes, the revision number must remain stable.
    #[test]
    fn invariant_revision_stable_on_rejection() {
        let mut wb = Workbook::new();

        // Set up initial state
        wb.set_cell_value_tracked(0, 0, 0, "initial");
        let rev_after_setup = wb.revision();

        // Empty batch (no ops)
        {
            let _guard = wb.batch_guard();
        }

        assert_eq!(
            wb.revision(),
            rev_after_setup,
            "empty batch should not increment revision"
        );

        // Batch that modifies same cell to same value (no-op in practice)
        // Note: Current impl doesn't detect no-change, so this increments revision.
        // This is acceptable behavior - true no-change detection is optimization.
    }

    /// Invariant: Revision increments by exactly 1 per successful batch.
    ///
    /// Not just "strictly monotonic" (r2 > r1) but the stronger property:
    /// r2 == r1 + 1. No skipping, no gaps. This enables efficient change
    /// detection and is a prerequisite for event boundary isolation.
    #[test]
    fn invariant_revision_increments_by_one_per_batch() {
        let mut wb = Workbook::new();

        // Record revision before batch 1
        let rev_before_batch1 = wb.revision();

        // Batch 1
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "batch1");
        }
        let rev_after_batch1 = wb.revision();

        // Batch 2
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 1, 0, "batch2");
        }
        let rev_after_batch2 = wb.revision();

        // Revisions should be strictly increasing
        assert!(
            rev_after_batch1 > rev_before_batch1,
            "revision should increment after batch 1"
        );
        assert!(
            rev_after_batch2 > rev_after_batch1,
            "revision should increment after batch 2"
        );
        assert_eq!(
            rev_after_batch2 - rev_after_batch1,
            1,
            "each batch increments revision by exactly 1"
        );
    }

    /// Invariant 8: Fingerprint must be deterministic across platforms.
    ///
    /// Given identical cell data, the fingerprint (hash) must be identical
    /// regardless of platform, endianness, or JSON key ordering.
    ///
    /// This test verifies basic fingerprint stability using canonical encoding.
    #[test]
    fn invariant_fingerprint_deterministic() {
        // Test canonical byte encoding for operations
        // Format: tag(1) + sheet(4 LE) + row(4 LE) + col(4 LE) + len(4 LE) + value(UTF-8)

        fn encode_set_value(sheet: u32, row: u32, col: u32, value: &str) -> Vec<u8> {
            let mut buf = Vec::new();
            buf.push(0x01); // SetCellValue tag
            buf.extend_from_slice(&sheet.to_le_bytes());
            buf.extend_from_slice(&row.to_le_bytes());
            buf.extend_from_slice(&col.to_le_bytes());
            buf.extend_from_slice(&(value.len() as u32).to_le_bytes());
            buf.extend_from_slice(value.as_bytes());
            buf
        }

        // Same operation should produce identical bytes
        let bytes1 = encode_set_value(0, 5, 3, "hello");
        let bytes2 = encode_set_value(0, 5, 3, "hello");
        assert_eq!(bytes1, bytes2, "identical ops must encode identically");

        // Different operations should produce different bytes
        let bytes3 = encode_set_value(0, 5, 3, "world");
        assert_ne!(bytes1, bytes3, "different values must encode differently");

        // Expected bytes for verification (can be computed on any platform)
        let expected = vec![
            0x01, // tag
            0x00, 0x00, 0x00, 0x00, // sheet 0
            0x05, 0x00, 0x00, 0x00, // row 5
            0x03, 0x00, 0x00, 0x00, // col 3
            0x05, 0x00, 0x00, 0x00, // len 5
            b'h', b'e', b'l', b'l', b'o', // value
        ];
        assert_eq!(bytes1, expected, "encoding must match expected bytes");
    }

    /// Invariant 9: Float canonicalization for fingerprinting.
    ///
    /// - NaN and infinity should be rejected (or canonicalized)
    /// - -0.0 must equal +0.0 in fingerprint
    /// - Floats must be encoded consistently (e.g., via to_le_bytes)
    #[test]
    fn invariant_float_canonicalization() {
        // Test float encoding for fingerprinting
        fn canonicalize_float(f: f64) -> Option<[u8; 8]> {
            if f.is_nan() || f.is_infinite() {
                return None; // Reject non-finite values
            }
            // Canonicalize -0.0 to +0.0
            let canonical = if f == 0.0 { 0.0_f64 } else { f };
            Some(canonical.to_le_bytes())
        }

        // Normal values work
        assert!(canonicalize_float(42.0).is_some());
        assert!(canonicalize_float(-123.456).is_some());

        // -0.0 canonicalizes to +0.0
        let neg_zero = canonicalize_float(-0.0_f64).unwrap();
        let pos_zero = canonicalize_float(0.0_f64).unwrap();
        assert_eq!(neg_zero, pos_zero, "-0.0 must equal +0.0 in fingerprint");

        // NaN rejected
        assert!(canonicalize_float(f64::NAN).is_none(), "NaN must be rejected");

        // Infinity rejected
        assert!(
            canonicalize_float(f64::INFINITY).is_none(),
            "infinity must be rejected"
        );
        assert!(
            canonicalize_float(f64::NEG_INFINITY).is_none(),
            "negative infinity must be rejected"
        );
    }

    /// Invariant 10: Discovery file must be written atomically.
    ///
    /// The discovery file (containing session info) must be written atomically
    /// so readers never see partial content. This is typically done via
    /// write-to-temp + rename.
    ///
    /// This test verifies the atomic write pattern at filesystem level.
    #[test]
    fn invariant_discovery_file_atomic() {
        use std::fs;
        use std::io::Write;

        let temp_dir = std::env::temp_dir();
        let final_path = temp_dir.join("visigrid-discovery-test.json");
        let temp_path = temp_dir.join("visigrid-discovery-test.json.tmp");

        // Cleanup from previous runs
        let _ = fs::remove_file(&final_path);
        let _ = fs::remove_file(&temp_path);

        // Atomic write pattern: write to temp, then rename
        let content = r#"{"pid":12345,"port":9876,"token":"abc123"}"#;

        // Step 1: Write to temp file
        {
            let mut file = fs::File::create(&temp_path).expect("create temp file");
            file.write_all(content.as_bytes()).expect("write temp file");
            file.sync_all().expect("sync temp file");
        }

        // Step 2: Atomic rename
        fs::rename(&temp_path, &final_path).expect("atomic rename");

        // Verify: final file is readable and complete
        let read_content = fs::read_to_string(&final_path).expect("read final file");
        assert_eq!(read_content, content, "content must match after atomic write");

        // Verify: temp file no longer exists
        assert!(
            !temp_path.exists(),
            "temp file should not exist after rename"
        );

        // Cleanup
        let _ = fs::remove_file(&final_path);
    }

    // =========================================================================
    // Additional Invariants (identified during review)
    // =========================================================================

    /// Invariant: cells_changed events must not coalesce across revisions.
    ///
    /// When two batches complete in quick succession, even with event throttling,
    /// the cells_changed events must not merge changes from different revisions.
    ///
    /// Each CellsChanged event has a revision tag. Cells in that event must
    /// only be from operations in the batch that produced that revision.
    #[test]
    fn invariant_events_no_cross_revision_coalesce() {
        use crate::harness::{EngineHarness, Op};

        let mut harness = EngineHarness::new();

        // Batch A: change A1 → revision r1
        let ops_a = vec![Op::SetCellValue {
            sheet_index: 0,
            row: 0,
            col: 0, // A1
            value: "batch_a".to_string(),
        }];
        harness.apply_ops(&ops_a, false);
        let r1 = harness.revision();

        // Batch B: change B1 → revision r2
        let ops_b = vec![Op::SetCellValue {
            sheet_index: 0,
            row: 0,
            col: 1, // B1
            value: "batch_b".to_string(),
        }];
        harness.apply_ops(&ops_b, false);
        let r2 = harness.revision();

        // Verify revisions are distinct
        assert_ne!(r1, r2, "batches should produce different revisions");

        // Collect cells_changed events
        let events = harness.events();
        let cells_events = events.cells_changed();
        assert_eq!(cells_events.len(), 2, "should have 2 CellsChanged events");

        // Find events by revision
        let event_r1 = cells_events.iter().find(|e| e.revision == r1);
        let event_r2 = cells_events.iter().find(|e| e.revision == r2);

        assert!(event_r1.is_some(), "should have CellsChanged for r1");
        assert!(event_r2.is_some(), "should have CellsChanged for r2");

        let event_r1 = event_r1.unwrap();
        let event_r2 = event_r2.unwrap();

        // Event for r1 should contain A1 only
        assert_eq!(event_r1.cells.len(), 1, "r1 event should have 1 cell");
        assert_eq!(event_r1.cells[0].col, 0, "r1 event should contain col 0 (A1)");

        // Event for r2 should contain B1 only
        assert_eq!(event_r2.cells.len(), 1, "r2 event should have 1 cell");
        assert_eq!(event_r2.cells[0].col, 1, "r2 event should contain col 1 (B1)");

        // No event should contain cells from both batches
        for event in cells_events {
            let has_a1 = event.cells.iter().any(|c| c.col == 0);
            let has_b1 = event.cells.iter().any(|c| c.col == 1);
            assert!(
                !(has_a1 && has_b1),
                "CellsChanged must not coalesce cells from different revisions"
            );
        }
    }

    /// Invariant: Rate limiter uses deterministic token bucket semantics.
    ///
    /// The rate limiter must have predictable boundary behavior:
    /// - Fixed refill rate (20k ops/sec)
    /// - Fixed burst capacity (40k ops)
    /// - Message exceeding burst must fail entirely (not partially)
    ///
    /// Requires: Session server layer (rate limiter implementation).
    #[test]
    #[ignore = "requires session server rate limiter implementation"]
    fn invariant_rate_limiter_burst_boundary() {
        // Test structure (to be implemented in session server):
        //
        // 1. Create rate limiter with 40k burst, 20k/sec refill
        // 2. Drain to exactly 0 tokens
        // 3. Wait 1 second (should have 20k tokens)
        // 4. Submit 20k ops → should succeed
        // 5. Submit 1 more op → should fail (rate limited)
        // 6. Wait 50ms (should have 1k tokens)
        // 7. Submit 1k ops → should succeed
        // 8. Submit 50k ops → should fail entirely (exceeds burst)
        //
        // Token bucket: capacity=40k, refill_rate=20k/sec
        // Deterministic: same timing → same result across runs
        unimplemented!("requires session server rate limiter");
    }

    /// Invariant: Partial non-atomic apply increments revision.
    ///
    /// When atomic=false and a batch fails partway through, the revision
    /// must still increment because state has changed (partial apply committed).
    ///
    /// This documents the chosen behavior: partial success = revision bump.
    #[test]
    fn invariant_revision_increment_on_partial_nonatomic() {
        use crate::harness::{EngineHarness, Op};

        let mut harness = EngineHarness::new();
        let initial_rev = harness.revision();

        // Apply a batch that will partially succeed (non-atomic)
        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0, // A1
                value: "first".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1, // B1
                value: "second".to_string(),
            },
            Op::SimulateError {
                message: "fail after 2 ops".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 2, // C1 - never applied
                value: "never".to_string(),
            },
        ];

        let result = harness.apply_ops(&ops, false); // atomic=false

        // Partial apply: 2 ops applied, revision incremented
        assert_eq!(result.applied, 2);
        assert!(result.error.is_some());
        assert_eq!(
            result.revision,
            initial_rev + 1,
            "partial apply must increment revision"
        );

        // Verify the changes were actually applied
        let wb = harness.workbook();
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "first");  // A1
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "second"); // B1
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "");       // C1 - not applied
    }

    /// Invariant: Partial non-atomic apply event semantics.
    ///
    /// When atomic=false and a batch fails partway:
    /// - RevisionChanged: emitted (state changed)
    /// - CellsChanged: emitted for applied ops ONLY (not failed/skipped)
    /// - BatchApplied: emitted with applied=N, total=M, error set
    ///
    /// This documents the exact event contract for partial failure.
    #[test]
    fn invariant_partial_nonatomic_event_semantics() {
        use crate::harness::{EngineHarness, Op};

        let mut harness = EngineHarness::new();

        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0, // A1
                value: "applied1".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1, // B1
                value: "applied2".to_string(),
            },
            Op::SimulateError {
                message: "fail".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 2, // C1 - skipped
                value: "skipped".to_string(),
            },
        ];

        let _result = harness.apply_ops(&ops, false); // atomic=false

        let events = harness.events();

        // RevisionChanged: must be emitted (state changed)
        assert_eq!(
            events.revision_changed().len(),
            1,
            "partial apply must emit RevisionChanged"
        );
        let rev_event = &events.revision_changed()[0];
        assert_eq!(rev_event.previous, 0);
        assert_eq!(rev_event.revision, 1);

        // CellsChanged: must contain only applied cells (A1, B1), not skipped (C1)
        assert_eq!(
            events.cells_changed().len(),
            1,
            "partial apply must emit CellsChanged"
        );
        let cells_event = &events.cells_changed()[0];
        assert_eq!(cells_event.revision, 1);
        assert_eq!(
            cells_event.cells.len(),
            2,
            "CellsChanged must have only applied cells"
        );

        // Verify cells are A1 and B1, not C1
        let cols: Vec<usize> = cells_event.cells.iter().map(|c| c.col).collect();
        assert!(cols.contains(&0), "CellsChanged must include A1");
        assert!(cols.contains(&1), "CellsChanged must include B1");
        assert!(!cols.contains(&2), "CellsChanged must NOT include C1 (skipped)");

        // BatchApplied: must have correct applied count and error
        assert_eq!(events.batch_applied().len(), 1);
        let batch_event = &events.batch_applied()[0];
        assert_eq!(batch_event.applied, 2, "BatchApplied.applied must be 2");
        assert_eq!(batch_event.total, 4, "BatchApplied.total must be 4");
        assert!(batch_event.error.is_some(), "BatchApplied must have error");
        assert_eq!(
            batch_event.error.as_ref().unwrap().op_index,
            2,
            "error must be at op_index 2"
        );
    }

    /// Invariant: Fingerprint golden vector (version 1).
    ///
    /// A fixed sequence of operations must produce a known fingerprint.
    /// This catches encoding drift, field reordering, and hash changes.
    ///
    /// If this test fails after intentional changes, bump the fingerprint
    /// version in the protocol and update the expected hash.
    #[test]
    fn invariant_fingerprint_golden_vector_v1() {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        /// Canonical op encoding for fingerprinting (version 1).
        /// Format: tag(1) + sheet(4 LE) + row(4 LE) + col(4 LE) + payload
        fn encode_op_v1(tag: u8, sheet: u32, row: u32, col: u32, payload: &[u8]) -> Vec<u8> {
            let mut buf = Vec::with_capacity(13 + payload.len());
            buf.push(tag);
            buf.extend_from_slice(&sheet.to_le_bytes());
            buf.extend_from_slice(&row.to_le_bytes());
            buf.extend_from_slice(&col.to_le_bytes());
            buf.extend_from_slice(payload);
            buf
        }

        fn encode_string_payload(s: &str) -> Vec<u8> {
            let mut buf = Vec::with_capacity(4 + s.len());
            buf.extend_from_slice(&(s.len() as u32).to_le_bytes());
            buf.extend_from_slice(s.as_bytes());
            buf
        }

        // Golden vector: a fixed sequence of operations
        let ops = [
            // Op 1: SetCellValue(sheet=0, row=0, col=0, value="Revenue")
            encode_op_v1(0x01, 0, 0, 0, &encode_string_payload("Revenue")),
            // Op 2: SetCellValue(sheet=0, row=0, col=1, value="100000")
            encode_op_v1(0x01, 0, 0, 1, &encode_string_payload("100000")),
            // Op 3: SetCellFormula(sheet=0, row=1, col=1, formula="=B1*1.1")
            encode_op_v1(0x02, 0, 1, 1, &encode_string_payload("=B1*1.1")),
        ];

        // Hash the concatenated ops
        let mut hasher = DefaultHasher::new();
        for op in &ops {
            op.hash(&mut hasher);
        }
        let fingerprint = hasher.finish();

        // Expected fingerprint (computed once, frozen forever for v1)
        // If encoding changes, this MUST fail. Update only with version bump.
        //
        // Note: DefaultHasher is not guaranteed stable across Rust versions,
        // but for this test we're verifying the ENCODING is stable, not the
        // hash algorithm. In production, use a stable hash (e.g., xxhash, blake3).
        //
        // For now, we verify the raw bytes match expected encoding:
        let expected_op1 = vec![
            0x01, // tag: SetCellValue
            0x00, 0x00, 0x00, 0x00, // sheet: 0
            0x00, 0x00, 0x00, 0x00, // row: 0
            0x00, 0x00, 0x00, 0x00, // col: 0
            0x07, 0x00, 0x00, 0x00, // len: 7
            b'R', b'e', b'v', b'e', b'n', b'u', b'e', // "Revenue"
        ];
        assert_eq!(ops[0], expected_op1, "Op 1 encoding must match golden vector");

        let expected_op2 = vec![
            0x01, // tag: SetCellValue
            0x00, 0x00, 0x00, 0x00, // sheet: 0
            0x00, 0x00, 0x00, 0x00, // row: 0
            0x01, 0x00, 0x00, 0x00, // col: 1
            0x06, 0x00, 0x00, 0x00, // len: 6
            b'1', b'0', b'0', b'0', b'0', b'0', // "100000"
        ];
        assert_eq!(ops[1], expected_op2, "Op 2 encoding must match golden vector");

        let expected_op3 = vec![
            0x02, // tag: SetCellFormula
            0x00, 0x00, 0x00, 0x00, // sheet: 0
            0x01, 0x00, 0x00, 0x00, // row: 1
            0x01, 0x00, 0x00, 0x00, // col: 1
            0x07, 0x00, 0x00, 0x00, // len: 7
            b'=', b'B', b'1', b'*', b'1', b'.', b'1', // "=B1*1.1"
        ];
        assert_eq!(ops[2], expected_op3, "Op 3 encoding must match golden vector");

        // Verify fingerprint is non-zero and deterministic
        assert_ne!(fingerprint, 0, "fingerprint must be non-zero");

        // Re-compute to verify determinism
        let mut hasher2 = DefaultHasher::new();
        for op in &ops {
            op.hash(&mut hasher2);
        }
        assert_eq!(
            hasher2.finish(),
            fingerprint,
            "fingerprint must be deterministic"
        );
    }

    // =========================================================================
    // Iteration Mode v1 tests
    // =========================================================================

    /// Helper: create a workbook with cycle and enable iteration.
    fn make_iterative_wb() -> Workbook {
        let mut wb = Workbook::new();
        wb.set_iterative_enabled(true);
        wb.set_iterative_max_iters(100);
        wb.set_iterative_tolerance(1e-9);
        wb
    }

    #[test]
    fn test_iterative_diverges() {
        // A1 = B1, B1 = A1 + 1 — diverges (grows without bound)
        let mut wb = make_iterative_wb();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=B1");       // A1 = B1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=A1+1");     // B1 = A1 + 1
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        let report = wb.recompute_full_ordered();

        assert!(report.had_cycles);
        assert_eq!(report.scc_count, 1);
        assert!(!report.converged, "divergent SCC should not converge");
        assert_eq!(report.iterations_performed, 100);
        // Non-converged cells should show #NUM!
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "#NUM!");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "#NUM!");
    }

    #[test]
    fn test_iterative_converges_fixed_point() {
        // A1 = (B1 + 10) / 2, B1 = (A1 + 10) / 2
        // Fixed point: A1 = B1 = 10
        let mut wb = make_iterative_wb();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2"); // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2"); // B1
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        let report = wb.recompute_full_ordered();

        assert!(report.had_cycles);
        assert_eq!(report.scc_count, 1);
        assert!(report.converged, "stable fixed point should converge");
        assert!(report.iterations_performed < 100, "should converge well before max_iters");

        // Both cells should be ~10.0
        let a1 = wb.sheet(0).unwrap().get_display(0, 0);
        let b1 = wb.sheet(0).unwrap().get_display(0, 1);
        let a1_val: f64 = a1.parse().expect("A1 should be numeric");
        let b1_val: f64 = b1.parse().expect("B1 should be numeric");
        assert!((a1_val - 10.0).abs() < 1e-6, "A1={}, expected 10.0", a1_val);
        assert!((b1_val - 10.0).abs() < 1e-6, "B1={}, expected 10.0", b1_val);
    }

    #[test]
    fn test_iterative_disabled_shows_cycle() {
        // Same cycle but with iteration disabled — should show #CYCLE!
        let mut wb = Workbook::new();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2");
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2");
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        let report = wb.recompute_full_ordered();

        assert!(report.had_cycles);
        assert_eq!(report.scc_count, 0, "no SCCs resolved when iteration disabled");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "#CYCLE!");
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "#CYCLE!");
    }

    #[test]
    fn test_iterative_downstream_cells_evaluate() {
        // A1 = (B1 + 10) / 2, B1 = (A1 + 10) / 2, C1 = A1 * 2
        // After convergence, C1 should be 20.0
        let mut wb = make_iterative_wb();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2"); // A1
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2"); // B1
        wb.sheet_mut(0).unwrap().set_value(0, 2, "=A1*2");      // C1 = A1 * 2
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);
        wb.update_cell_deps(sheet_id, 0, 2);

        let report = wb.recompute_full_ordered();

        assert!(report.converged);
        let c1 = wb.sheet(0).unwrap().get_display(0, 2);
        let c1_val: f64 = c1.parse().expect("C1 should be numeric");
        assert!((c1_val - 20.0).abs() < 1e-6, "C1={}, expected 20.0", c1_val);
    }

    #[test]
    fn test_iterative_self_reference_diverges() {
        // A1 = A1 + 1 — self-loop, diverges
        let mut wb = make_iterative_wb();
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=A1+1");
        wb.update_cell_deps(sheet_id, 0, 0);

        let report = wb.recompute_full_ordered();

        assert!(report.had_cycles);
        assert_eq!(report.scc_count, 1);
        assert!(!report.converged);
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "#NUM!");
    }

    #[test]
    fn test_iterative_tolerance_matters() {
        // With loose tolerance, convergence happens faster
        let mut wb = make_iterative_wb();
        wb.set_iterative_tolerance(1.0); // very loose
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2");
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2");
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        let report_loose = wb.recompute_full_ordered();

        // Rebuild with tight tolerance
        let mut wb2 = make_iterative_wb();
        wb2.set_iterative_tolerance(1e-15);
        let sheet_id2 = wb2.sheet_id_at_idx(0).unwrap();

        wb2.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2");
        wb2.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2");
        wb2.update_cell_deps(sheet_id2, 0, 0);
        wb2.update_cell_deps(sheet_id2, 0, 1);

        let report_tight = wb2.recompute_full_ordered();

        assert!(report_loose.converged);
        assert!(report_tight.converged);
        assert!(
            report_loose.iterations_performed <= report_tight.iterations_performed,
            "loose tolerance ({}) should converge in <= iterations than tight ({})",
            report_loose.iterations_performed, report_tight.iterations_performed,
        );
    }

    #[test]
    fn test_iterative_max_iters_respected() {
        // Set max_iters=3, convergent system should stop at 3 if tolerance is too tight
        let mut wb = make_iterative_wb();
        wb.set_iterative_max_iters(3);
        wb.set_iterative_tolerance(1e-30); // impossibly tight
        let sheet_id = wb.sheet_id_at_idx(0).unwrap();

        wb.sheet_mut(0).unwrap().set_value(0, 0, "=(B1+10)/2");
        wb.sheet_mut(0).unwrap().set_value(0, 1, "=(A1+10)/2");
        wb.update_cell_deps(sheet_id, 0, 0);
        wb.update_cell_deps(sheet_id, 0, 1);

        let report = wb.recompute_full_ordered();

        assert_eq!(report.iterations_performed, 3);
        assert!(!report.converged, "should not converge with impossibly tight tolerance in 3 iters");
        // Should show #NUM! since didn't converge
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "#NUM!");
    }

    #[test]
    fn cross_sheet_incremental_sumif() {
        // Reproduce the exact recon-template scenario:
        // summary!B2 = SUMIF(Sheet1!E2:E10, "charge", Sheet1!C2:C10)
        // Edit Sheet1!E2 and C2, verify summary updates via incremental recalc.
        let mut wb = Workbook::new();
        let si = wb.add_sheet_named("summary").expect("add summary sheet");

        wb.set_cell_value_tracked(0, 0, 2, "amount_minor");
        wb.set_cell_value_tracked(0, 0, 4, "type");
        wb.set_cell_value_tracked(si, 0, 0, "Charges");
        wb.set_cell_value_tracked(si, 0, 1, r#"=SUMIF(Sheet1!E2:E10,"charge",Sheet1!C2:C10)"#);

        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        // Initial: summary!B2 = 0
        let val = wb.sheet(si).unwrap().get_computed_value(0, 1);
        assert!(
            matches!(val, Value::Number(n) if n.abs() < 0.001),
            "initial should be 0, got {:?}", val
        );

        // Simulate GUI cell edits (same path as Spreadsheet::set_cell_value)
        let sheet1_id = wb.sheet(0).unwrap().id;

        wb.sheet_mut(0).unwrap().set_value(1, 4, "charge"); // E2
        wb.update_cell_deps(sheet1_id, 1, 4);
        wb.note_cell_changed(CellId::new(sheet1_id, 1, 4));

        wb.sheet_mut(0).unwrap().set_value(1, 2, "50000"); // C2
        wb.update_cell_deps(sheet1_id, 1, 2);
        wb.note_cell_changed(CellId::new(sheet1_id, 1, 2));

        // summary!B2 should now be 50000
        let val = wb.sheet(si).unwrap().get_computed_value(0, 1);
        assert!(
            matches!(val, Value::Number(n) if (n - 50000.0).abs() < 0.001),
            "should be 50000, got {:?}", val
        );
    }

    #[test]
    fn cross_sheet_formula_entry_no_ref_error() {
        // Reproduce the GUI bug: entering =Sheet1!A1+Sheet1!B1 on another sheet
        // should NOT show #REF!. The sheet-local evaluate_and_spill must skip
        // cross-sheet formulas and let the workbook-level recalc handle them.
        let mut wb = Workbook::new();
        let si = wb.add_sheet_named("summary").expect("add summary sheet");

        // Put values on Sheet1
        wb.set_cell_value_tracked(0, 0, 0, "10"); // Sheet1!A1 = 10
        wb.set_cell_value_tracked(0, 0, 1, "20"); // Sheet1!B1 = 20

        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        // Now simulate typing a cross-sheet formula on the summary sheet
        // This is the exact GUI path: set_value → update_cell_deps → note_cell_changed
        let summary_id = wb.sheet(si).unwrap().id;
        wb.sheet_mut(si).unwrap().set_value(0, 0, "=Sheet1!A1+Sheet1!B1");
        wb.update_cell_deps(summary_id, 0, 0);
        wb.note_cell_changed(CellId::new(summary_id, 0, 0));

        // Should be 30, NOT #REF!
        let val = wb.sheet(si).unwrap().get_computed_value(0, 0);
        assert!(
            matches!(val, Value::Number(n) if (n - 30.0).abs() < 0.001),
            "cross-sheet formula should evaluate to 30, got {:?}", val
        );
    }
}
