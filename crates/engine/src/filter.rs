//! AutoFilter and Sort - Row View Layer
//!
//! This module provides the view layer that maps between:
//! - View space (what the user sees, affected by sort/filter)
//! - Data space (canonical storage, row 0..N-1)
//!
//! Key invariants:
//! - UI code uses view space
//! - Storage and formulas use data space
//! - Conversion happens at the boundary only
//! - visible_mask is indexed by DATA row (not view row)
//! - All lookups are O(1)

use std::collections::{HashMap, HashSet};
use ordered_float::OrderedFloat;
use serde::{Deserialize, Serialize};

use crate::formula::eval::Value;

// =============================================================================
// RowView: The core view layer mapping
// =============================================================================

/// Row view layer: maps between view space and data space
#[derive(Debug, Clone)]
pub struct RowView {
    /// Maps view_row index -> data_row index
    /// Identity by default: [0, 1, 2, ..., N-1]
    /// After sort: permuted to reflect sort order
    row_order: Vec<usize>,

    /// Inverse map: data_row -> view_row (O(1) lookup)
    /// Rebuilt whenever row_order changes
    data_to_view_map: Vec<usize>,

    /// Visibility mask indexed by DATA row (not view row)
    /// true = visible, false = hidden by filter
    /// Sorting doesn't affect this — it stays aligned with canonical data
    visible_mask: Vec<bool>,

    /// Cached list of visible VIEW row indices
    /// Rebuilt when filters change OR sort changes
    /// Used for fast iteration in rendering/navigation
    visible_rows: Vec<usize>,
}

impl Default for RowView {
    fn default() -> Self {
        Self::new(0)
    }
}

impl RowView {
    /// Initialize identity mapping for N rows
    pub fn new(row_count: usize) -> Self {
        Self {
            row_order: (0..row_count).collect(),
            data_to_view_map: (0..row_count).collect(),
            visible_mask: vec![true; row_count],
            visible_rows: (0..row_count).collect(),
        }
    }

    /// Total number of rows (data rows)
    pub fn row_count(&self) -> usize {
        self.row_order.len()
    }

    /// Number of visible rows
    pub fn visible_count(&self) -> usize {
        self.visible_rows.len()
    }

    /// Map view row to data row - O(1)
    pub fn view_to_data(&self, view_row: usize) -> usize {
        self.row_order[view_row]
    }

    /// Map data row to view row - O(1)
    /// Returns None if the data row is hidden by a filter
    pub fn data_to_view(&self, data_row: usize) -> Option<usize> {
        if self.is_data_row_visible(data_row) {
            Some(self.data_to_view_map[data_row])
        } else {
            None
        }
    }

    /// Map data row to view row unconditionally - O(1)
    /// Returns the view position even if hidden (needed for undo/sort)
    pub fn data_to_view_unchecked(&self, data_row: usize) -> usize {
        self.data_to_view_map[data_row]
    }

    /// Check if a data row is visible - O(1)
    pub fn is_data_row_visible(&self, data_row: usize) -> bool {
        data_row < self.visible_mask.len() && self.visible_mask[data_row]
    }

    /// Check if a view row is visible - O(1)
    pub fn is_view_row_visible(&self, view_row: usize) -> bool {
        if view_row >= self.row_order.len() {
            return false;
        }
        let data_row = self.view_to_data(view_row);
        self.visible_mask[data_row]
    }

    /// Get visible view rows (for rendering/navigation)
    pub fn visible_rows(&self) -> &[usize] {
        &self.visible_rows
    }

    /// Get the nth visible view row (for indexed access)
    pub fn nth_visible(&self, n: usize) -> Option<usize> {
        self.visible_rows.get(n).copied()
    }

    /// Find the index of a view row in visible_rows
    pub fn visible_index_of(&self, view_row: usize) -> Option<usize> {
        self.visible_rows.iter().position(|&vr| vr == view_row)
    }

    /// Is any filtering active?
    pub fn is_filtered(&self) -> bool {
        self.visible_count() < self.row_count()
    }

    /// Is the view sorted (non-identity order)?
    pub fn is_sorted(&self) -> bool {
        self.row_order.iter().enumerate().any(|(i, &d)| i != d)
    }

    // -------------------------------------------------------------------------
    // Internal rebuilders
    // -------------------------------------------------------------------------

    /// Rebuild inverse map after sort
    fn rebuild_inverse_map(&mut self) {
        // Ensure map is correct size
        if self.data_to_view_map.len() != self.row_order.len() {
            self.data_to_view_map.resize(self.row_order.len(), 0);
        }
        for (view_row, &data_row) in self.row_order.iter().enumerate() {
            if data_row < self.data_to_view_map.len() {
                self.data_to_view_map[data_row] = view_row;
            }
        }
    }

    /// Rebuild visible_rows cache (visible view rows in order)
    fn rebuild_visible_cache(&mut self) {
        self.visible_rows = self.row_order
            .iter()
            .enumerate()
            .filter_map(|(view_row, &data_row)| {
                if data_row < self.visible_mask.len() && self.visible_mask[data_row] {
                    Some(view_row)
                } else {
                    None
                }
            })
            .collect();
    }

    // -------------------------------------------------------------------------
    // Mutators
    // -------------------------------------------------------------------------

    /// Apply a sort permutation (stable sort required)
    /// The permutation maps new_view_row -> data_row
    pub fn apply_sort(&mut self, permutation: Vec<usize>) {
        self.row_order = permutation;
        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }

    /// Reset to identity order (unsorted)
    pub fn clear_sort(&mut self) {
        self.row_order = (0..self.row_order.len()).collect();
        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }

    /// Apply filter visibility (mask indexed by data row)
    pub fn apply_filter(&mut self, visible_mask: Vec<bool>) {
        self.visible_mask = visible_mask;
        self.rebuild_visible_cache();
    }

    /// Clear all filters (all rows visible)
    pub fn clear_filter(&mut self) {
        self.visible_mask = vec![true; self.row_order.len()];
        self.rebuild_visible_cache();
    }

    /// Handle row insertion at data_row index
    ///
    /// Behavior depends on sort state:
    /// - No sort active: insert at view position = data_row (preserves canonical order)
    /// - Sort active: append to end (user must re-sort to integrate)
    pub fn insert_row(&mut self, data_row: usize) {
        let sort_active = self.is_sorted();

        // Shift all existing data_row references >= inserted index
        for data_row_ref in self.row_order.iter_mut() {
            if *data_row_ref >= data_row {
                *data_row_ref += 1;
            }
        }

        // Insert new row (visible by default)
        if sort_active {
            // Append to end when sorted — user re-sorts to integrate
            self.row_order.push(data_row);
        } else {
            // Insert at canonical position when unsorted
            let insert_view_pos = data_row.min(self.row_order.len());
            self.row_order.insert(insert_view_pos, data_row);
        }

        // Grow inverse map and visibility mask
        self.data_to_view_map.push(0); // Placeholder, rebuilt below
        self.visible_mask.push(true);

        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }

    /// Handle row deletion at data_row index
    pub fn delete_row(&mut self, data_row: usize) {
        if data_row >= self.row_order.len() {
            return;
        }

        let view_row = self.data_to_view_map[data_row];
        self.row_order.remove(view_row);

        if data_row < self.visible_mask.len() {
            self.visible_mask.remove(data_row);
        }
        if data_row < self.data_to_view_map.len() {
            self.data_to_view_map.remove(data_row);
        }

        // Shift all data_row references > deleted index
        for data_row_ref in self.row_order.iter_mut() {
            if *data_row_ref > data_row {
                *data_row_ref -= 1;
            }
        }

        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }

    /// Resize to match a new row count (e.g., after sheet resize)
    pub fn resize(&mut self, new_row_count: usize) {
        let old_count = self.row_order.len();

        if new_row_count > old_count {
            // Add new rows at the end
            for i in old_count..new_row_count {
                self.row_order.push(i);
                self.visible_mask.push(true);
            }
        } else if new_row_count < old_count {
            // Remove rows from the end
            self.row_order.retain(|&d| d < new_row_count);
            self.visible_mask.truncate(new_row_count);
        }

        self.data_to_view_map.resize(new_row_count, 0);
        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }

    /// Get current row_order for undo storage
    pub fn row_order(&self) -> &[usize] {
        &self.row_order
    }

    /// Get current visible_mask for undo storage
    pub fn visible_mask(&self) -> &[bool] {
        &self.visible_mask
    }

    /// Restore from undo state (row_order + visible_mask)
    pub fn restore(&mut self, row_order: Vec<usize>, visible_mask: Vec<bool>) {
        self.row_order = row_order;
        self.visible_mask = visible_mask;
        self.data_to_view_map.resize(self.row_order.len(), 0);
        self.rebuild_inverse_map();
        self.rebuild_visible_cache();
    }
}

// =============================================================================
// FilterKey: Typed key for filter comparison
// =============================================================================

/// Error kinds for filter keys (compact enum, stable equality)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
pub enum ErrorKind {
    Ref,
    Value,
    Div0,
    Name,
    Null,
    Num,
    Na,
    Spill,
    Other,
}

impl ErrorKind {
    /// Parse from error string
    pub fn from_str(s: &str) -> Self {
        let upper = s.to_uppercase();
        if upper.contains("REF") {
            ErrorKind::Ref
        } else if upper.contains("VALUE") {
            ErrorKind::Value
        } else if upper.contains("DIV") {
            ErrorKind::Div0
        } else if upper.contains("NAME") {
            ErrorKind::Name
        } else if upper.contains("NULL") {
            ErrorKind::Null
        } else if upper.contains("NUM") {
            ErrorKind::Num
        } else if upper.contains("N/A") || upper.contains("NA") {
            ErrorKind::Na
        } else if upper.contains("SPILL") {
            ErrorKind::Spill
        } else {
            ErrorKind::Other
        }
    }
}

/// Typed key for filter comparison
/// Derived from cell's computed Value, not display formatting
/// Stores RAW values; use normalized() for comparison
#[derive(Debug, Clone)]
pub enum FilterKey {
    /// Empty cell
    Blank,
    /// Error value (#REF!, #VALUE!, etc.)
    Error(ErrorKind),
    /// Boolean
    Bool(bool),
    /// Numeric (includes dates as serial numbers)
    Number(OrderedFloat<f64>),
    /// Text (stores RAW value; normalize for comparison)
    Text(String),
}

impl FilterKey {
    /// Create from a cell's computed Value
    pub fn from_value(value: &Value) -> Self {
        match value {
            Value::Empty => FilterKey::Blank,
            Value::Number(n) => FilterKey::Number(OrderedFloat(*n)),
            Value::Text(s) => FilterKey::Text(s.clone()),
            Value::Boolean(b) => FilterKey::Bool(*b),
            Value::Error(e) => FilterKey::Error(ErrorKind::from_str(e)),
        }
    }

    /// Normalized form for equality/hashing (used in HashSet)
    pub fn normalized(&self) -> NormalizedFilterKey {
        match self {
            FilterKey::Blank => NormalizedFilterKey::Blank,
            FilterKey::Error(e) => NormalizedFilterKey::Error(*e),
            FilterKey::Bool(b) => NormalizedFilterKey::Bool(*b),
            FilterKey::Number(n) => NormalizedFilterKey::Number(*n),
            FilterKey::Text(s) => NormalizedFilterKey::Text(s.trim().to_lowercase()),
        }
    }

    /// Display string for dropdown (raw text, formatted numbers)
    pub fn display_string(&self) -> String {
        match self {
            FilterKey::Blank => "(Blanks)".to_string(),
            FilterKey::Error(e) => format!("#{:?}", e).to_uppercase(),
            FilterKey::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            FilterKey::Number(n) => {
                // Simple display - could be enhanced with format
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", n.0 as i64)
                } else {
                    format!("{}", n.0)
                }
            }
            FilterKey::Text(s) => s.clone(),
        }
    }
}

/// Normalized key for comparison/hashing (used in filter HashSet)
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum NormalizedFilterKey {
    Blank,
    Error(ErrorKind),
    Bool(bool),
    Number(OrderedFloat<f64>),
    /// Already normalized: trimmed + lowercase (ASCII-ish)
    Text(String),
}

// =============================================================================
// SortKey: For stable, deterministic sorting
// =============================================================================

/// Key for sorting rows (includes tie-breaker for stability)
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct SortKey {
    /// Type rank: Numbers(0) < Text(1) < Bool(2) < Error(3) < Blank(4)
    pub type_rank: u8,
    /// Normalized value for comparison
    pub value: NormalizedFilterKey,
    /// Original VIEW row index at the moment sort is applied
    /// This is the tie-breaker: "stable relative to what user currently sees"
    pub original_view_index: usize,
}

impl SortKey {
    /// Create from filter key and current view row index
    pub fn from_filter_key(key: &FilterKey, view_row: usize) -> Self {
        let type_rank = match key {
            FilterKey::Number(_) => 0,
            FilterKey::Text(_) => 1,
            FilterKey::Bool(_) => 2,
            FilterKey::Error(_) => 3,
            FilterKey::Blank => 4,
        };
        Self {
            type_rank,
            value: key.normalized(),
            original_view_index: view_row,
        }
    }
}

/// Sort direction
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SortDirection {
    Ascending,
    Descending,
}

/// Current sort state
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SortState {
    pub column: usize,
    pub direction: SortDirection,
}

/// Undo item for sort operations
#[derive(Debug, Clone)]
pub struct SortUndoItem {
    /// Previous row_order before this sort
    pub previous_row_order: Vec<usize>,
    /// Previous sort state (None if unsorted)
    pub previous_sort_state: Option<SortState>,
}

// =============================================================================
// Sorting Logic
// =============================================================================

/// Sort rows by column value within the filter range.
///
/// This is the **one** function UI calls. Engine owns the permutation.
///
/// # Arguments
/// - `row_view`: Current row view state
/// - `filter_state`: Filter configuration (provides sort range)
/// - `value_at`: Callback to get computed value for (data_row, col)
/// - `col`: Column to sort by
/// - `direction`: Ascending or Descending
///
/// # Returns
/// - New row_order permutation to apply via `row_view.apply_sort()`
/// - Undo item with previous state
///
/// # Invariants
/// - Header row (min_row of filter_range) never moves
/// - Rows outside filter_range stay in place
/// - visible_mask is preserved (sorting doesn't unhide rows)
/// - Stable sort: equal keys preserve relative view order
pub fn sort_by_column<F>(
    row_view: &RowView,
    filter_state: &FilterState,
    value_at: F,
    col: usize,
    direction: SortDirection,
) -> (Vec<usize>, SortUndoItem)
where
    F: Fn(usize, usize) -> Value, // (data_row, col) -> Value
{
    // Capture previous state for undo
    let undo = SortUndoItem {
        previous_row_order: row_view.row_order().to_vec(),
        previous_sort_state: filter_state.sort.clone(),
    };

    // Get data range (excludes header row)
    let Some((data_start, _, data_end, _)) = filter_state.data_range() else {
        // No filter range = no sort (return identity)
        return (row_view.row_order().to_vec(), undo);
    };

    // Current row_order
    let current_order = row_view.row_order();
    let row_count = current_order.len();

    // Find the view row indices that map to data rows in sort range
    // Build: Vec<(SortKey, view_row)> for stable sorting
    let mut sortable: Vec<(SortKey, usize)> = Vec::new();

    for (view_row, &data_row) in current_order.iter().enumerate() {
        if data_row >= data_start && data_row <= data_end {
            // This view_row's data is in the sortable range
            let value = value_at(data_row, col);
            let filter_key = FilterKey::from_value(&value);
            let sort_key = SortKey::from_filter_key(&filter_key, view_row);
            sortable.push((sort_key, view_row));
        }
    }

    // Stable sort the sortable portion
    sortable.sort_by(|a, b| a.0.cmp(&b.0));

    // Reverse for descending
    if direction == SortDirection::Descending {
        sortable.reverse();
    }

    // Build new row_order:
    // - Header rows (before data_start) stay in place
    // - Sortable rows get reordered
    // - Trailing rows (after data_end) stay in place
    let mut new_order = vec![0usize; row_count];
    let mut sortable_iter = sortable.iter();

    for view_row in 0..row_count {
        let data_row = current_order[view_row];

        if data_row >= data_start && data_row <= data_end {
            // This position gets the next sorted data_row
            if let Some((_, orig_view_row)) = sortable_iter.next() {
                new_order[view_row] = current_order[*orig_view_row];
            }
        } else {
            // Outside sort range: preserve
            new_order[view_row] = data_row;
        }
    }

    (new_order, undo)
}

// =============================================================================
// FilterState: Per-sheet filter configuration
// =============================================================================

/// Entry in unique values cache
#[derive(Debug, Clone)]
pub struct UniqueValueEntry {
    /// Normalized key for comparison/filtering
    pub key: NormalizedFilterKey,
    /// Display string (from first-seen raw value for Text)
    pub display: String,
    /// Count of rows with this value (for UI and quick select-all)
    pub count: usize,
}

/// Text filter mode
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub enum TextFilterMode {
    Contains,
    NotContains,
    StartsWith,
    EndsWith,
    Equals,
    NotEquals,
}

/// Text filter predicate
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextFilter {
    pub mode: TextFilterMode,
    pub value: String,
    pub case_sensitive: bool,
}

impl TextFilter {
    pub fn matches(&self, text: &str) -> bool {
        let (haystack, needle) = if self.case_sensitive {
            (text.to_string(), self.value.clone())
        } else {
            (text.to_lowercase(), self.value.to_lowercase())
        };

        match self.mode {
            TextFilterMode::Contains => haystack.contains(&needle),
            TextFilterMode::NotContains => !haystack.contains(&needle),
            TextFilterMode::StartsWith => haystack.starts_with(&needle),
            TextFilterMode::EndsWith => haystack.ends_with(&needle),
            TextFilterMode::Equals => haystack == needle,
            TextFilterMode::NotEquals => haystack != needle,
        }
    }
}

/// Per-column filter criteria
#[derive(Debug, Clone, Default)]
pub struct ColumnFilter {
    /// Selected normalized filter keys to INCLUDE (None = all pass)
    /// Uses NormalizedFilterKey for consistent comparison
    pub selected: Option<HashSet<NormalizedFilterKey>>,

    /// Optional text predicate (AND with selected)
    pub text_filter: Option<TextFilter>,
}

impl ColumnFilter {
    /// Check if a value passes this filter
    pub fn passes(&self, key: &FilterKey) -> bool {
        let normalized = key.normalized();

        // Check value selection
        if let Some(selected) = &self.selected {
            if !selected.contains(&normalized) {
                return false;
            }
        }

        // Check text filter (AND with value selection)
        // RULE: Non-text values FAIL text filters
        if let Some(text_filter) = &self.text_filter {
            match key {
                FilterKey::Text(s) => {
                    if !text_filter.matches(s) {
                        return false;
                    }
                }
                // Non-text values fail text filters
                _ => return false,
            }
        }

        true
    }

    /// Is this filter active (has any criteria)?
    pub fn is_active(&self) -> bool {
        self.selected.is_some() || self.text_filter.is_some()
    }
}

/// Filter state for a sheet
#[derive(Debug, Clone, Default)]
pub struct FilterState {
    /// The range that has AutoFilter enabled (min_row, min_col, max_row, max_col)
    /// None = no filter active
    /// Header row is ALWAYS min_row (derived, not stored separately)
    pub filter_range: Option<(usize, usize, usize, usize)>,

    /// Per-column filter criteria, keyed by column index
    pub column_filters: HashMap<usize, ColumnFilter>,

    /// Current sort state
    pub sort: Option<SortState>,

    /// Cached unique values per column (invalidate on edit/insert/delete/paste)
    pub unique_values_cache: HashMap<usize, Vec<UniqueValueEntry>>,
}

impl FilterState {
    /// Header row (always min_row of filter_range)
    pub fn header_row(&self) -> Option<usize> {
        self.filter_range.map(|(min_r, _, _, _)| min_r)
    }

    /// Build unique values cache for a column
    ///
    /// Scans data rows (not header) and builds:
    /// - Unique values with counts (frequency)
    /// - Sorted by count descending (most frequent first)
    /// - Limited to max_values entries
    ///
    /// Returns cached values if available; call invalidate_column() to force rebuild.
    pub fn build_unique_values<F>(
        &mut self,
        col: usize,
        value_at: F,
        max_values: usize,
    ) -> &[UniqueValueEntry]
    where
        F: Fn(usize, usize) -> Value, // (data_row, col) -> Value
    {
        // Return cached if available
        if self.unique_values_cache.contains_key(&col) {
            return &self.unique_values_cache[&col];
        }

        // Get data range (excludes header)
        let Some((data_start, _, data_end, _)) = self.data_range() else {
            self.unique_values_cache.insert(col, Vec::new());
            return &self.unique_values_cache[&col];
        };

        // Collect values with counts
        // Use NormalizedFilterKey for grouping, keep first raw display for each
        let mut counts: HashMap<NormalizedFilterKey, (String, usize)> = HashMap::new();

        for data_row in data_start..=data_end {
            let value = value_at(data_row, col);
            let filter_key = FilterKey::from_value(&value);
            let normalized = filter_key.normalized();
            let display = filter_key.display_string();

            counts
                .entry(normalized)
                .and_modify(|(_, count)| *count += 1)
                .or_insert((display, 1));
        }

        // Convert to entries and sort by frequency (descending)
        let mut entries: Vec<UniqueValueEntry> = counts
            .into_iter()
            .map(|(key, (display, count))| UniqueValueEntry { key, display, count })
            .collect();

        entries.sort_by(|a, b| b.count.cmp(&a.count));

        // Limit to max_values
        entries.truncate(max_values);

        self.unique_values_cache.insert(col, entries);
        &self.unique_values_cache[&col]
    }

    /// Get cached unique values for a column (returns None if not cached)
    pub fn get_unique_values(&self, col: usize) -> Option<&[UniqueValueEntry]> {
        self.unique_values_cache.get(&col).map(|v| v.as_slice())
    }

    /// Build unique values cache from pre-collected values
    ///
    /// Variant of build_unique_values that takes pre-computed values
    /// to avoid borrow conflicts in the caller.
    pub fn build_unique_values_from_vec(
        &mut self,
        col: usize,
        values: &[(usize, Value)],
        max_values: usize,
    ) {
        // Return if cached
        if self.unique_values_cache.contains_key(&col) {
            return;
        }

        // Collect values with counts
        let mut counts: HashMap<NormalizedFilterKey, (String, usize)> = HashMap::new();

        for (_data_row, value) in values {
            let filter_key = FilterKey::from_value(value);
            let normalized = filter_key.normalized();
            let display = filter_key.display_string();

            counts
                .entry(normalized)
                .and_modify(|(_, count)| *count += 1)
                .or_insert((display, 1));
        }

        // Convert to entries and sort by frequency (descending)
        let mut entries: Vec<UniqueValueEntry> = counts
            .into_iter()
            .map(|(key, (display, count))| UniqueValueEntry { key, display, count })
            .collect();

        entries.sort_by(|a, b| b.count.cmp(&a.count));

        // Limit to max_values
        entries.truncate(max_values);

        self.unique_values_cache.insert(col, entries);
    }

    /// Data range for sort/filter (excludes header row)
    pub fn data_range(&self) -> Option<(usize, usize, usize, usize)> {
        self.filter_range.map(|(min_r, min_c, max_r, max_c)| {
            (min_r + 1, min_c, max_r, max_c)
        })
    }

    /// Check if a column is within filter range
    pub fn contains_column(&self, col: usize) -> bool {
        self.filter_range.map_or(false, |(_, min_c, _, max_c)| {
            col >= min_c && col <= max_c
        })
    }

    /// Check if a data row is within filter data range
    pub fn contains_data_row(&self, row: usize) -> bool {
        self.data_range().map_or(false, |(min_r, _, max_r, _)| {
            row >= min_r && row <= max_r
        })
    }

    /// Is AutoFilter enabled?
    pub fn is_enabled(&self) -> bool {
        self.filter_range.is_some()
    }

    /// Is any column actively filtered?
    pub fn has_active_filter(&self) -> bool {
        self.column_filters.values().any(|f| f.is_active())
    }

    /// Invalidate cache for a column (call on cell edit)
    pub fn invalidate_column(&mut self, col: usize) {
        self.unique_values_cache.remove(&col);
    }

    /// Invalidate all caches (call on row/column insert/delete, paste, structural changes)
    pub fn invalidate_all_caches(&mut self) {
        self.unique_values_cache.clear();
    }

    /// Clear filter for a specific column
    pub fn clear_column_filter(&mut self, col: usize) {
        self.column_filters.remove(&col);
    }

    /// Clear all column filters (keeps filter range and sort)
    pub fn clear_all_filters(&mut self) {
        self.column_filters.clear();
        self.unique_values_cache.clear();
    }

    /// Disable AutoFilter entirely
    pub fn disable(&mut self) {
        self.filter_range = None;
        self.column_filters.clear();
        self.sort = None;
        self.unique_values_cache.clear();
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_row_view_identity() {
        let view = RowView::new(5);
        assert_eq!(view.row_count(), 5);
        assert_eq!(view.visible_count(), 5);

        // Identity mapping
        for i in 0..5 {
            assert_eq!(view.view_to_data(i), i);
            assert_eq!(view.data_to_view(i), Some(i));
            assert!(view.is_data_row_visible(i));
        }

        assert!(!view.is_sorted());
        assert!(!view.is_filtered());
    }

    #[test]
    fn test_row_view_sort() {
        let mut view = RowView::new(5);

        // Sort: reverse order
        view.apply_sort(vec![4, 3, 2, 1, 0]);

        assert!(view.is_sorted());
        assert_eq!(view.view_to_data(0), 4);
        assert_eq!(view.view_to_data(4), 0);
        assert_eq!(view.data_to_view(0), Some(4));
        assert_eq!(view.data_to_view(4), Some(0));

        // Clear sort
        view.clear_sort();
        assert!(!view.is_sorted());
        assert_eq!(view.view_to_data(0), 0);
    }

    #[test]
    fn test_row_view_filter() {
        let mut view = RowView::new(5);

        // Hide rows 1 and 3
        view.apply_filter(vec![true, false, true, false, true]);

        assert!(view.is_filtered());
        assert_eq!(view.visible_count(), 3);
        assert_eq!(view.visible_rows(), &[0, 2, 4]);

        assert!(view.is_data_row_visible(0));
        assert!(!view.is_data_row_visible(1));
        assert!(view.is_data_row_visible(2));
        assert!(!view.is_data_row_visible(3));
        assert!(view.is_data_row_visible(4));

        // Clear filter
        view.clear_filter();
        assert!(!view.is_filtered());
        assert_eq!(view.visible_count(), 5);
    }

    #[test]
    fn test_row_view_sort_and_filter() {
        let mut view = RowView::new(5);

        // Sort reverse
        view.apply_sort(vec![4, 3, 2, 1, 0]);

        // Filter: hide data rows 1 and 3
        // visible_mask is indexed by DATA row
        view.apply_filter(vec![true, false, true, false, true]);

        assert!(view.is_sorted());
        assert!(view.is_filtered());
        assert_eq!(view.visible_count(), 3);

        // View order is [4, 3, 2, 1, 0] (data rows)
        // Visible data rows are 0, 2, 4
        // So visible view rows are: view 0 (data 4), view 2 (data 2), view 4 (data 0)
        assert_eq!(view.visible_rows(), &[0, 2, 4]);
    }

    #[test]
    fn test_row_view_roundtrip() {
        let view = RowView::new(100);

        // view->data->view roundtrip should hold (for visible rows)
        for v in 0..100 {
            let d = view.view_to_data(v);
            let v2 = view.data_to_view(d).expect("all rows visible");
            assert_eq!(v, v2);
        }
    }

    #[test]
    fn test_filter_key_normalization() {
        let k1 = FilterKey::Text("  Apple  ".to_string());
        let k2 = FilterKey::Text("apple".to_string());
        let k3 = FilterKey::Text("APPLE".to_string());

        assert_eq!(k1.normalized(), k2.normalized());
        assert_eq!(k2.normalized(), k3.normalized());
    }

    #[test]
    fn test_column_filter_passes() {
        let mut filter = ColumnFilter::default();

        // No criteria = all pass
        assert!(filter.passes(&FilterKey::Text("anything".to_string())));
        assert!(filter.passes(&FilterKey::Number(OrderedFloat(42.0))));
        assert!(filter.passes(&FilterKey::Blank));

        // Value selection
        let mut selected = HashSet::new();
        selected.insert(NormalizedFilterKey::Text("apple".to_string()));
        filter.selected = Some(selected);

        assert!(filter.passes(&FilterKey::Text("Apple".to_string())));
        assert!(filter.passes(&FilterKey::Text("  apple  ".to_string())));
        assert!(!filter.passes(&FilterKey::Text("banana".to_string())));
        assert!(!filter.passes(&FilterKey::Number(OrderedFloat(42.0))));
    }

    #[test]
    fn test_text_filter_non_text_fails() {
        let mut filter = ColumnFilter::default();
        filter.text_filter = Some(TextFilter {
            mode: TextFilterMode::Contains,
            value: "foo".to_string(),
            case_sensitive: false,
        });

        // Text with match passes
        assert!(filter.passes(&FilterKey::Text("foobar".to_string())));

        // Text without match fails
        assert!(!filter.passes(&FilterKey::Text("bar".to_string())));

        // Non-text always fails text filter
        assert!(!filter.passes(&FilterKey::Number(OrderedFloat(42.0))));
        assert!(!filter.passes(&FilterKey::Blank));
        assert!(!filter.passes(&FilterKey::Bool(true)));
    }

    #[test]
    fn test_sort_key_ordering() {
        // Numbers come before text
        let num = SortKey::from_filter_key(&FilterKey::Number(OrderedFloat(1.0)), 0);
        let text = SortKey::from_filter_key(&FilterKey::Text("a".to_string()), 1);
        assert!(num < text);

        // Text comes before blanks
        let blank = SortKey::from_filter_key(&FilterKey::Blank, 2);
        assert!(text < blank);

        // Stable sort: same value, earlier view row wins
        let num1 = SortKey::from_filter_key(&FilterKey::Number(OrderedFloat(1.0)), 0);
        let num2 = SortKey::from_filter_key(&FilterKey::Number(OrderedFloat(1.0)), 1);
        assert!(num1 < num2);
    }

    // =========================================================================
    // Phase 2 Tests: sort_by_column
    // =========================================================================

    /// Helper: create test data for sorting
    /// Returns (row_view, filter_state, values_map)
    fn setup_sort_test() -> (RowView, FilterState, HashMap<(usize, usize), Value>) {
        let row_view = RowView::new(10);
        let mut filter_state = FilterState::default();
        // Filter range: rows 0-9, col 0-2
        // Header row is 0, data rows are 1-9
        filter_state.filter_range = Some((0, 0, 9, 2));

        let mut values: HashMap<(usize, usize), Value> = HashMap::new();
        // Column 0: Numbers (for numeric sort test)
        // Row 0 = header, rows 1-9 = data
        values.insert((0, 0), Value::Text("Numbers".to_string())); // header
        values.insert((1, 0), Value::Number(30.0));
        values.insert((2, 0), Value::Number(10.0));
        values.insert((3, 0), Value::Number(50.0));
        values.insert((4, 0), Value::Number(20.0));
        values.insert((5, 0), Value::Number(40.0));
        values.insert((6, 0), Value::Number(10.0)); // duplicate for stability test
        values.insert((7, 0), Value::Number(60.0));
        values.insert((8, 0), Value::Number(10.0)); // another duplicate
        values.insert((9, 0), Value::Number(25.0));

        // Column 1: Text (for text sort test)
        values.insert((0, 1), Value::Text("Names".to_string())); // header
        values.insert((1, 1), Value::Text("Charlie".to_string()));
        values.insert((2, 1), Value::Text("alice".to_string()));   // lowercase
        values.insert((3, 1), Value::Text("BOB".to_string()));     // uppercase
        values.insert((4, 1), Value::Text("Alice".to_string()));   // mixed case
        values.insert((5, 1), Value::Text("bob".to_string()));     // lowercase
        values.insert((6, 1), Value::Text("ALICE".to_string()));   // uppercase
        values.insert((7, 1), Value::Text("David".to_string()));
        values.insert((8, 1), Value::Text("  alice  ".to_string())); // with whitespace
        values.insert((9, 1), Value::Text("Bob".to_string()));

        (row_view, filter_state, values)
    }

    #[test]
    fn test_sort_numbers_ascending() {
        let (row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        let (new_order, _undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0, // sort by column 0 (numbers)
            SortDirection::Ascending,
        );

        // Header row 0 should still be at view position 0
        assert_eq!(new_order[0], 0, "Header row must stay at position 0");

        // Check sorted data: 10, 10, 10, 20, 25, 30, 40, 50, 60
        // Data rows: 2=10, 6=10, 8=10, 4=20, 9=25, 1=30, 5=40, 3=50, 7=60
        let sorted_data: Vec<f64> = new_order[1..].iter()
            .map(|&d| match values.get(&(d, 0)) {
                Some(Value::Number(n)) => *n,
                _ => f64::MAX,
            })
            .collect();

        assert_eq!(sorted_data, vec![10.0, 10.0, 10.0, 20.0, 25.0, 30.0, 40.0, 50.0, 60.0]);
    }

    #[test]
    fn test_sort_numbers_descending() {
        let (row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        let (new_order, _undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0,
            SortDirection::Descending,
        );

        // Header row 0 should still be at view position 0
        assert_eq!(new_order[0], 0, "Header row must stay at position 0");

        // Check sorted data: 60, 50, 40, 30, 25, 20, 10, 10, 10
        let sorted_data: Vec<f64> = new_order[1..].iter()
            .map(|&d| match values.get(&(d, 0)) {
                Some(Value::Number(n)) => *n,
                _ => f64::MIN,
            })
            .collect();

        assert_eq!(sorted_data, vec![60.0, 50.0, 40.0, 30.0, 25.0, 20.0, 10.0, 10.0, 10.0]);
    }

    #[test]
    fn test_sort_text_case_insensitive() {
        let (row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        let (new_order, _undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            1, // sort by column 1 (text)
            SortDirection::Ascending,
        );

        // Header stays at position 0
        assert_eq!(new_order[0], 0);

        // All "alice" variants (2, 4, 6, 8) should be grouped together
        // All "bob" variants (3, 5, 9) should be grouped together
        // Charlie (1) and David (7) are unique
        let sorted_names: Vec<String> = new_order[1..].iter()
            .map(|&d| match values.get(&(d, 1)) {
                Some(Value::Text(s)) => s.trim().to_lowercase(),
                _ => String::new(),
            })
            .collect();

        // Verify alphabetical order (case-insensitive)
        for i in 1..sorted_names.len() {
            assert!(sorted_names[i-1] <= sorted_names[i],
                "Text sort not alphabetical: {:?} > {:?}",
                sorted_names[i-1], sorted_names[i]);
        }
    }

    #[test]
    fn test_sort_stability() {
        let (row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        let (new_order, _undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0, // sort by numbers
            SortDirection::Ascending,
        );

        // Find the three rows with value 10 (data rows 2, 6, 8)
        // In original order: 2 appears before 6 appears before 8
        // After stable sort, they should maintain this relative order
        let ten_rows: Vec<usize> = new_order.iter()
            .filter(|&&d| {
                matches!(values.get(&(d, 0)), Some(Value::Number(n)) if *n == 10.0)
            })
            .copied()
            .collect();

        assert_eq!(ten_rows.len(), 3);
        // Data rows 2, 6, 8 should appear in that order (original view order preserved)
        assert_eq!(ten_rows, vec![2, 6, 8], "Stable sort: equal values preserve original order");
    }

    #[test]
    fn test_sort_header_row_unchanged() {
        let (row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        // Sort ascending
        let (new_order_asc, _) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0,
            SortDirection::Ascending,
        );

        // Sort descending
        let (new_order_desc, _) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0,
            SortDirection::Descending,
        );

        // Header row (data row 0) must stay at view position 0 in both cases
        assert_eq!(new_order_asc[0], 0, "Header must not move (ascending)");
        assert_eq!(new_order_desc[0], 0, "Header must not move (descending)");
    }

    #[test]
    fn test_sort_undo_restores_exact_order() {
        let (mut row_view, filter_state, values) = setup_sort_test();

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        // Capture original order
        let original_order = row_view.row_order().to_vec();

        // Sort
        let (new_order, undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0,
            SortDirection::Ascending,
        );

        // Apply sort
        row_view.apply_sort(new_order);
        assert!(row_view.is_sorted());

        // Verify undo captured previous order
        assert_eq!(undo.previous_row_order, original_order);

        // Apply undo (restore previous order)
        row_view.apply_sort(undo.previous_row_order);

        // Verify exact restoration
        assert_eq!(row_view.row_order(), &original_order[..]);
        assert!(!row_view.is_sorted(), "Should be back to identity");
    }

    #[test]
    fn test_sort_with_hidden_rows() {
        let (mut row_view, filter_state, values) = setup_sort_test();

        // Hide some data rows (via visible_mask, indexed by DATA row)
        // Hide data rows 2 and 5
        let mut mask = vec![true; 10];
        mask[2] = false; // hide data row 2
        mask[5] = false; // hide data row 5
        row_view.apply_filter(mask.clone());

        assert!(row_view.is_filtered());
        assert_eq!(row_view.visible_count(), 8);

        let value_at = |data_row: usize, col: usize| -> Value {
            values.get(&(data_row, col)).cloned().unwrap_or(Value::Empty)
        };

        // Sort
        let (new_order, _undo) = sort_by_column(
            &row_view,
            &filter_state,
            value_at,
            0,
            SortDirection::Ascending,
        );

        // Apply sort
        row_view.apply_sort(new_order);

        // visible_mask must be preserved (still indexed by DATA row)
        assert!(!row_view.is_data_row_visible(2), "Data row 2 should still be hidden");
        assert!(!row_view.is_data_row_visible(5), "Data row 5 should still be hidden");

        // visible_rows should still have 8 entries
        assert_eq!(row_view.visible_count(), 8);

        // Header should still be at position 0
        assert_eq!(row_view.view_to_data(0), 0, "Header row must stay at position 0");
    }

    // =========================================================================
    // Phase 3 Tests: Filtering (unique values, apply_filter)
    // =========================================================================

    #[test]
    fn test_build_unique_values() {
        let mut filter_state = FilterState::default();
        filter_state.filter_range = Some((0, 0, 5, 0));

        let values = vec![
            (1, Value::Text("Apple".to_string())),
            (2, Value::Text("apple".to_string())),
            (3, Value::Text("Banana".to_string())),
            (4, Value::Text("Apple".to_string())),
            (5, Value::Text("cherry".to_string())),
        ];

        filter_state.build_unique_values_from_vec(0, &values, 500);

        let unique = filter_state.get_unique_values(0).unwrap();
        assert_eq!(unique.len(), 3, "Should have 3 unique values (case-insensitive)");

        // Should be sorted by frequency (descending)
        // Apple (3) > Banana (1) = cherry (1)
        assert_eq!(unique[0].count, 3);
    }

    #[test]
    fn test_filter_state_column_filter() {
        let mut filter_state = FilterState::default();
        filter_state.filter_range = Some((0, 0, 10, 2));

        // Initially no filters
        assert!(!filter_state.has_active_filter());

        // Add a filter
        let mut selected = HashSet::new();
        selected.insert(NormalizedFilterKey::Text("apple".to_string()));
        filter_state.column_filters.insert(0, ColumnFilter {
            selected: Some(selected),
            text_filter: None,
        });

        assert!(filter_state.has_active_filter());
        assert!(filter_state.column_filters.get(&0).unwrap().is_active());

        // Test passes
        let apple = FilterKey::Text("Apple".to_string());
        let banana = FilterKey::Text("Banana".to_string());
        let cf = filter_state.column_filters.get(&0).unwrap();

        assert!(cf.passes(&apple), "Apple should pass (case-insensitive)");
        assert!(!cf.passes(&banana), "Banana should not pass");
    }

    #[test]
    fn test_apply_filter_visibility() {
        let mut row_view = RowView::new(6);

        // Hide rows 2 and 4
        let mask = vec![true, true, false, true, false, true];
        row_view.apply_filter(mask);

        assert!(row_view.is_filtered());
        assert_eq!(row_view.visible_count(), 4);
        assert_eq!(row_view.visible_rows(), &[0, 1, 3, 5]);

        // Check individual visibility
        assert!(row_view.is_data_row_visible(0));
        assert!(row_view.is_data_row_visible(1));
        assert!(!row_view.is_data_row_visible(2));
        assert!(row_view.is_data_row_visible(3));
        assert!(!row_view.is_data_row_visible(4));
        assert!(row_view.is_data_row_visible(5));
    }

    #[test]
    fn test_clear_filter_restores_visibility() {
        let mut row_view = RowView::new(5);

        // Apply filter
        row_view.apply_filter(vec![true, false, true, false, true]);
        assert_eq!(row_view.visible_count(), 3);

        // Clear filter
        row_view.clear_filter();
        assert_eq!(row_view.visible_count(), 5);
        assert!(!row_view.is_filtered());
    }

    #[test]
    fn test_filter_with_sort() {
        let mut row_view = RowView::new(5);

        // Apply sort (reverse order)
        row_view.apply_sort(vec![4, 3, 2, 1, 0]);

        // Apply filter (hide data rows 1 and 3)
        row_view.apply_filter(vec![true, false, true, false, true]);

        // View order: [4, 3, 2, 1, 0] (data rows)
        // Visible data rows: 0, 2, 4
        // Visible view rows: view 0 (data 4), view 2 (data 2), view 4 (data 0)
        assert_eq!(row_view.visible_count(), 3);
        assert_eq!(row_view.visible_rows(), &[0, 2, 4]);

        // Verify view_to_data for visible rows
        assert_eq!(row_view.view_to_data(0), 4);
        assert_eq!(row_view.view_to_data(2), 2);
        assert_eq!(row_view.view_to_data(4), 0);
    }

    #[test]
    fn test_unique_values_max_limit() {
        let mut filter_state = FilterState::default();
        filter_state.filter_range = Some((0, 0, 100, 0));

        // Create 100 unique values
        let values: Vec<(usize, Value)> = (1..=100)
            .map(|i| (i, Value::Text(format!("Value{}", i))))
            .collect();

        // Limit to 10
        filter_state.build_unique_values_from_vec(0, &values, 10);

        let unique = filter_state.get_unique_values(0).unwrap();
        assert_eq!(unique.len(), 10, "Should be limited to 10 values");
    }

    #[test]
    fn test_invalidate_cache() {
        let mut filter_state = FilterState::default();
        filter_state.filter_range = Some((0, 0, 5, 0));

        let values = vec![
            (1, Value::Text("Test".to_string())),
        ];

        filter_state.build_unique_values_from_vec(0, &values, 500);
        assert!(filter_state.get_unique_values(0).is_some());

        // Invalidate
        filter_state.invalidate_column(0);
        assert!(filter_state.get_unique_values(0).is_none());
    }
}
