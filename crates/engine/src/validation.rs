//! Data Validation for cells
//!
//! Constrains what users can enter into cells: dropdown lists, number ranges,
//! date limits, and custom formula rules.
//!
//! ## Case Sensitivity
//!
//! - **List validation matching**: Case-sensitive. "Yes" != "yes".
//! - **Dropdown filter search**: Case-insensitive (UI layer handles this).
//!
//! This is explicit and consistent. Users who want case-insensitive matching
//! should normalize their list items.

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};

/// Maximum number of items in a resolved list. Prevents UI freeze on huge ranges.
pub const MAX_LIST_ITEMS: usize = 10_000;

// ============================================================================
// Core Types
// ============================================================================

/// A validation rule that constrains cell input.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ValidationRule {
    /// The type of validation to apply.
    pub rule_type: ValidationType,
    /// If true, empty/blank values are always valid.
    pub ignore_blank: bool,
    /// For List type: show dropdown arrow in cell. Ignored for other types.
    pub show_dropdown: bool,
    /// Optional message shown when cell is selected.
    pub input_message: Option<InputMessage>,
    /// Optional error alert shown when validation fails.
    pub error_alert: Option<ErrorAlert>,
}

impl ValidationRule {
    /// Create a new validation rule with the given type.
    pub fn new(rule_type: ValidationType) -> Self {
        // Default show_dropdown to true for List types
        let show_dropdown = matches!(rule_type, ValidationType::List(_));
        Self {
            rule_type,
            ignore_blank: true,
            show_dropdown,
            input_message: None,
            error_alert: None,
        }
    }

    /// Set ignore_blank option.
    pub fn with_ignore_blank(mut self, ignore: bool) -> Self {
        self.ignore_blank = ignore;
        self
    }

    /// Set show_dropdown option (for List validation).
    pub fn with_show_dropdown(mut self, show: bool) -> Self {
        self.show_dropdown = show;
        self
    }

    /// Set input message.
    pub fn with_input_message(mut self, message: InputMessage) -> Self {
        self.input_message = Some(message);
        self
    }

    /// Set error alert.
    pub fn with_error_alert(mut self, alert: ErrorAlert) -> Self {
        self.error_alert = Some(alert);
        self
    }

    /// Create a list validation rule from inline values.
    pub fn list_inline(values: Vec<String>) -> Self {
        Self::new(ValidationType::List(ListSource::Inline(values)))
    }

    /// Create a list validation rule from a range reference.
    pub fn list_range(range_ref: impl Into<String>) -> Self {
        Self::new(ValidationType::List(ListSource::Range(range_ref.into())))
    }

    /// Create a whole number validation rule.
    pub fn whole_number(constraint: NumericConstraint) -> Self {
        Self::new(ValidationType::WholeNumber(constraint))
    }

    /// Create a decimal validation rule.
    pub fn decimal(constraint: NumericConstraint) -> Self {
        Self::new(ValidationType::Decimal(constraint))
    }

    /// Create a custom formula validation rule.
    pub fn custom(formula: impl Into<String>) -> Self {
        Self::new(ValidationType::Custom(formula.into()))
    }
}

/// The type of validation to apply.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ValidationType {
    /// No validation (accept any value).
    AnyValue,
    /// Restrict to integers within bounds.
    WholeNumber(NumericConstraint),
    /// Restrict to decimals within bounds.
    Decimal(NumericConstraint),
    /// Restrict to a list of allowed values.
    List(ListSource),
    /// Restrict to dates within bounds.
    Date(NumericConstraint),
    /// Restrict to times within bounds.
    Time(NumericConstraint),
    /// Restrict text to character count bounds.
    TextLength(NumericConstraint),
    /// Custom formula that must return TRUE.
    Custom(String),
}

/// Numeric constraint for validation (used by WholeNumber, Decimal, Date, Time, TextLength).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct NumericConstraint {
    /// The comparison operator.
    pub operator: ComparisonOperator,
    /// First value (required for all operators).
    pub value1: ConstraintValue,
    /// Second value (required for Between/NotBetween).
    pub value2: Option<ConstraintValue>,
}

impl NumericConstraint {
    /// Create a "between" constraint.
    pub fn between(min: impl Into<ConstraintValue>, max: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::Between,
            value1: min.into(),
            value2: Some(max.into()),
        }
    }

    /// Create a "not between" constraint.
    pub fn not_between(min: impl Into<ConstraintValue>, max: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::NotBetween,
            value1: min.into(),
            value2: Some(max.into()),
        }
    }

    /// Create an "equal to" constraint.
    pub fn equal_to(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::EqualTo,
            value1: value.into(),
            value2: None,
        }
    }

    /// Create a "not equal to" constraint.
    pub fn not_equal_to(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::NotEqualTo,
            value1: value.into(),
            value2: None,
        }
    }

    /// Create a "greater than" constraint.
    pub fn greater_than(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::GreaterThan,
            value1: value.into(),
            value2: None,
        }
    }

    /// Create a "less than" constraint.
    pub fn less_than(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::LessThan,
            value1: value.into(),
            value2: None,
        }
    }

    /// Create a "greater than or equal" constraint.
    pub fn greater_than_or_equal(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::GreaterThanOrEqual,
            value1: value.into(),
            value2: None,
        }
    }

    /// Create a "less than or equal" constraint.
    pub fn less_than_or_equal(value: impl Into<ConstraintValue>) -> Self {
        Self {
            operator: ComparisonOperator::LessThanOrEqual,
            value1: value.into(),
            value2: None,
        }
    }
}

/// Comparison operator for numeric constraints.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ComparisonOperator {
    Between,
    NotBetween,
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

/// A value used in a constraint (number, cell reference, or formula).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ConstraintValue {
    /// A literal number.
    Number(f64),
    /// A cell reference (e.g., "A1" or "Sheet2!B5").
    CellRef(String),
    /// A formula (e.g., "=TODAY()").
    Formula(String),
}

impl From<f64> for ConstraintValue {
    fn from(n: f64) -> Self {
        ConstraintValue::Number(n)
    }
}

impl From<i32> for ConstraintValue {
    fn from(n: i32) -> Self {
        ConstraintValue::Number(n as f64)
    }
}

impl From<i64> for ConstraintValue {
    fn from(n: i64) -> Self {
        ConstraintValue::Number(n as f64)
    }
}

/// Source of values for list validation.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ListSource {
    /// Inline list of allowed values.
    Inline(Vec<String>),
    /// Range reference (e.g., "=A1:A10" or "=Sheet2!B1:B20").
    Range(String),
    /// Named range (e.g., "StatusOptions").
    NamedRange(String),
}

// ============================================================================
// Numeric Validation Helpers
// ============================================================================

/// Error when parsing numeric input for validation.
#[derive(Debug, Clone, PartialEq)]
pub enum NumericParseError {
    /// Input is empty (after trimming whitespace).
    Empty,
    /// Input contains invalid characters or format.
    InvalidFormat,
    /// Input has a fractional part but WholeNumber validation requires integer.
    FractionalNotAllowed,
}

impl std::fmt::Display for NumericParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            NumericParseError::Empty => write!(f, "Value is empty"),
            NumericParseError::InvalidFormat => write!(f, "Value is not a valid number"),
            NumericParseError::FractionalNotAllowed => write!(f, "Whole number required (no decimals)"),
        }
    }
}

/// Error when resolving a constraint value.
#[derive(Debug, Clone, PartialEq)]
pub enum ConstraintResolveError {
    /// Referenced cell is blank.
    BlankConstraint,
    /// Referenced cell or formula result is not numeric.
    NotNumeric,
    /// Cell reference could not be resolved.
    InvalidReference(String),
    /// Formula evaluation failed.
    FormulaError(String),
}

impl std::fmt::Display for ConstraintResolveError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ConstraintResolveError::BlankConstraint => write!(f, "Constraint is blank"),
            ConstraintResolveError::NotNumeric => write!(f, "Constraint is not numeric"),
            ConstraintResolveError::InvalidReference(r) => write!(f, "Invalid reference: {}", r),
            ConstraintResolveError::FormulaError(e) => write!(f, "Formula error: {}", e),
        }
    }
}

/// Reason why a cell failed validation.
///
/// Used for reporting after paste/fill operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValidationFailureReason {
    /// Input value doesn't match the allowed type/range.
    InvalidValue,
    /// Constraint references a blank cell.
    ConstraintBlank,
    /// Constraint reference is not numeric.
    ConstraintNotNumeric,
    /// Constraint reference could not be resolved.
    InvalidReference,
    /// Formula constraint not supported.
    FormulaNotSupported,
    /// List is empty (no valid options).
    ListEmpty,
    /// Value not in list.
    NotInList,
}

/// Parse user input as a number for validation.
///
/// # Rules
/// - Whitespace is trimmed
/// - Leading `+` is allowed
/// - Decimal point allowed only if `allow_decimal` is true
/// - For WholeNumber: rejects any fractional input (including `3.0`, `3.`)
///
/// # Examples
/// ```
/// use visigrid_engine::validation::parse_numeric_input;
///
/// // Decimal allows fractional
/// assert!(parse_numeric_input("3.14", true).is_ok());
/// assert!(parse_numeric_input(".5", true).is_ok());
///
/// // WholeNumber rejects fractional
/// assert!(parse_numeric_input("3.14", false).is_err());
/// assert!(parse_numeric_input("3.0", false).is_err());
/// assert!(parse_numeric_input("3", false).is_ok());
/// ```
pub fn parse_numeric_input(value: &str, allow_decimal: bool) -> Result<f64, NumericParseError> {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return Err(NumericParseError::Empty);
    }

    // Strip leading + (allowed for positive numbers)
    let normalized = if trimmed.starts_with('+') {
        &trimmed[1..]
    } else {
        trimmed
    };

    if normalized.is_empty() {
        return Err(NumericParseError::InvalidFormat);
    }

    // Check for decimal point
    let has_decimal = normalized.contains('.');

    if !allow_decimal && has_decimal {
        // WholeNumber validation: reject any input with decimal point
        // This includes "3.0" and "3." - strict integer requirement
        return Err(NumericParseError::FractionalNotAllowed);
    }

    // Parse the number
    normalized.parse::<f64>().map_err(|_| NumericParseError::InvalidFormat)
}

/// Evaluate a numeric constraint.
///
/// # Arguments
/// - `x`: The value being validated
/// - `operator`: The comparison operator
/// - `a`: First constraint value (min for Between, the value for single-value operators)
/// - `b`: Second constraint value (max for Between/NotBetween, None for others)
///
/// # Between Inclusivity
/// - `Between(a, b)`: returns true if `a <= x <= b` (inclusive)
/// - `NotBetween(a, b)`: returns true if `x < a || x > b`
pub fn eval_numeric_constraint(
    x: f64,
    operator: ComparisonOperator,
    a: f64,
    b: Option<f64>,
) -> bool {
    match operator {
        ComparisonOperator::Between => {
            let max = b.unwrap_or(a);
            x >= a && x <= max
        }
        ComparisonOperator::NotBetween => {
            let max = b.unwrap_or(a);
            x < a || x > max
        }
        ComparisonOperator::EqualTo => (x - a).abs() < f64::EPSILON,
        ComparisonOperator::NotEqualTo => (x - a).abs() >= f64::EPSILON,
        ComparisonOperator::GreaterThan => x > a,
        ComparisonOperator::LessThan => x < a,
        ComparisonOperator::GreaterThanOrEqual => x >= a,
        ComparisonOperator::LessThanOrEqual => x <= a,
    }
}

// ============================================================================
// Messages
// ============================================================================

/// Input message shown when a validated cell is selected.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct InputMessage {
    /// Whether to show the message.
    pub show: bool,
    /// Title of the message.
    pub title: String,
    /// Body of the message.
    pub message: String,
}

impl InputMessage {
    /// Create a new input message.
    pub fn new(title: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            show: true,
            title: title.into(),
            message: message.into(),
        }
    }
}

/// Error alert shown when validation fails.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ErrorAlert {
    /// Whether to show the alert.
    pub show: bool,
    /// Style of the alert (Stop, Warning, Information).
    pub style: ErrorStyle,
    /// Title of the alert.
    pub title: String,
    /// Body of the alert.
    pub message: String,
}

impl ErrorAlert {
    /// Create a new error alert with Stop style.
    pub fn stop(title: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            show: true,
            style: ErrorStyle::Stop,
            title: title.into(),
            message: message.into(),
        }
    }

    /// Create a new error alert with Warning style.
    pub fn warning(title: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            show: true,
            style: ErrorStyle::Warning,
            title: title.into(),
            message: message.into(),
        }
    }

    /// Create a new error alert with Information style.
    pub fn info(title: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            show: true,
            style: ErrorStyle::Information,
            title: title.into(),
            message: message.into(),
        }
    }
}

impl Default for ErrorAlert {
    fn default() -> Self {
        Self {
            show: true,
            style: ErrorStyle::Stop,
            title: "Invalid Entry".to_string(),
            message: "The value you entered is not valid.".to_string(),
        }
    }
}

/// Style of error alert.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum ErrorStyle {
    /// Reject input, user must fix or cancel.
    #[default]
    Stop,
    /// Allow override with confirmation.
    Warning,
    /// Show message, allow input anyway.
    Information,
}

// ============================================================================
// Validation Result
// ============================================================================

/// Result of validating a cell input.
#[derive(Debug, Clone, PartialEq)]
pub enum ValidationResult {
    /// Input is valid.
    Valid,
    /// Input is invalid.
    Invalid {
        /// The rule that was violated.
        rule: ValidationRule,
        /// Human-readable description of why validation failed.
        reason: String,
    },
}

impl ValidationResult {
    /// Returns true if the result is valid.
    pub fn is_valid(&self) -> bool {
        matches!(self, ValidationResult::Valid)
    }

    /// Returns true if the result is invalid.
    pub fn is_invalid(&self) -> bool {
        matches!(self, ValidationResult::Invalid { .. })
    }
}

// ============================================================================
// Resolved List (for dropdown UI)
// ============================================================================

/// A resolved list of dropdown items ready for UI display.
///
/// Returned by `Sheet::get_list_items()` for cells with list validation.
#[derive(Debug, Clone, PartialEq)]
pub struct ResolvedList {
    /// The list items (trimmed, possibly truncated).
    pub items: Vec<String>,
    /// True if the list was truncated due to exceeding MAX_LIST_ITEMS.
    pub is_truncated: bool,
    /// Fingerprint of the source data. Changes when source cells change.
    /// Used to detect stale dropdowns.
    pub source_fingerprint: u64,
}

impl ResolvedList {
    /// Create a new resolved list from items.
    ///
    /// Normalizes items (trim whitespace) and truncates if needed.
    pub fn from_items(raw_items: Vec<String>) -> Self {
        let mut items: Vec<String> = raw_items
            .into_iter()
            .map(|s| s.trim().to_string())
            .collect();

        let is_truncated = items.len() > MAX_LIST_ITEMS;
        if is_truncated {
            items.truncate(MAX_LIST_ITEMS);
        }

        let source_fingerprint = Self::compute_fingerprint(&items);

        Self {
            items,
            is_truncated,
            source_fingerprint,
        }
    }

    /// Create an empty resolved list.
    pub fn empty() -> Self {
        Self {
            items: Vec::new(),
            is_truncated: false,
            source_fingerprint: 0,
        }
    }

    /// Compute a fingerprint for the list items.
    fn compute_fingerprint(items: &[String]) -> u64 {
        let mut hasher = DefaultHasher::new();
        items.hash(&mut hasher);
        hasher.finish()
    }

    /// Check if a value matches any item in the list (case-sensitive).
    pub fn contains(&self, value: &str) -> bool {
        let trimmed = value.trim();
        self.items.iter().any(|item| item == trimmed)
    }

    /// Check if a value matches any item in the list (case-insensitive).
    /// Used for dropdown filter searching.
    pub fn contains_case_insensitive(&self, value: &str) -> bool {
        let trimmed = value.trim();
        self.items.iter().any(|item| item.eq_ignore_ascii_case(trimmed))
    }

    /// Filter items by a search string (case-insensitive).
    /// Returns items that contain the search string as a substring.
    pub fn filter(&self, search: &str) -> Vec<&str> {
        let search_lower = search.to_lowercase();
        self.items
            .iter()
            .filter(|item| item.to_lowercase().contains(&search_lower))
            .map(|s| s.as_str())
            .collect()
    }
}

// ============================================================================
// Cell Range (for validation storage)
// ============================================================================

/// A rectangular range of cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRange {
    /// Start row (0-indexed).
    pub start_row: usize,
    /// Start column (0-indexed).
    pub start_col: usize,
    /// End row (inclusive, 0-indexed).
    pub end_row: usize,
    /// End column (inclusive, 0-indexed).
    pub end_col: usize,
}

impl CellRange {
    /// Create a new cell range.
    pub fn new(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> Self {
        Self {
            start_row: start_row.min(end_row),
            start_col: start_col.min(end_col),
            end_row: start_row.max(end_row),
            end_col: start_col.max(end_col),
        }
    }

    /// Create a range for a single cell.
    pub fn single(row: usize, col: usize) -> Self {
        Self::new(row, col, row, col)
    }

    /// Check if this range contains the given cell.
    pub fn contains(&self, row: usize, col: usize) -> bool {
        row >= self.start_row && row <= self.end_row
            && col >= self.start_col && col <= self.end_col
    }

    /// Check if this range overlaps with another range.
    pub fn overlaps(&self, other: &CellRange) -> bool {
        !(self.end_row < other.start_row
            || self.start_row > other.end_row
            || self.end_col < other.start_col
            || self.start_col > other.end_col)
    }

    /// Number of cells in this range.
    pub fn cell_count(&self) -> usize {
        (self.end_row - self.start_row + 1) * (self.end_col - self.start_col + 1)
    }
}

impl PartialOrd for CellRange {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for CellRange {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        (self.start_row, self.start_col, self.end_row, self.end_col)
            .cmp(&(other.start_row, other.start_col, other.end_row, other.end_col))
    }
}

// ============================================================================
// Validation Store (per-sheet storage)
// ============================================================================

/// Storage for validation rules in a sheet.
///
/// Uses a BTreeMap for deterministic ordering. When looking up a validation rule
/// for a cell, we find the first rule whose range contains the cell.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ValidationStore {
    /// Map from cell range to validation rule.
    rules: BTreeMap<CellRange, ValidationRule>,
}

impl ValidationStore {
    /// Create a new empty validation store.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set a validation rule for a range.
    ///
    /// If a rule already exists for this exact range, it is replaced.
    pub fn set(&mut self, range: CellRange, rule: ValidationRule) {
        self.rules.insert(range, rule);
    }

    /// Remove the validation rule for an exact range.
    ///
    /// Returns the removed rule if it existed.
    pub fn remove(&mut self, range: &CellRange) -> Option<ValidationRule> {
        self.rules.remove(range)
    }

    /// Clear all validation rules that overlap with the given range.
    pub fn clear_range(&mut self, range: &CellRange) {
        self.rules.retain(|r, _| !r.overlaps(range));
    }

    /// Get the validation rule that applies to a cell.
    ///
    /// If multiple rules cover the cell, returns the first one (by range order).
    pub fn get(&self, row: usize, col: usize) -> Option<&ValidationRule> {
        for (range, rule) in &self.rules {
            if range.contains(row, col) {
                return Some(rule);
            }
        }
        None
    }

    /// Check if any validation rule applies to a cell.
    pub fn has_validation(&self, row: usize, col: usize) -> bool {
        self.get(row, col).is_some()
    }

    /// Iterate over all (range, rule) pairs.
    pub fn iter(&self) -> impl Iterator<Item = (&CellRange, &ValidationRule)> {
        self.rules.iter()
    }

    /// Number of validation rules.
    pub fn len(&self) -> usize {
        self.rules.len()
    }

    /// Check if there are no validation rules.
    pub fn is_empty(&self) -> bool {
        self.rules.is_empty()
    }

    /// Clear all validation rules.
    pub fn clear(&mut self) {
        self.rules.clear();
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_range_contains() {
        let range = CellRange::new(1, 1, 3, 3);
        assert!(range.contains(1, 1));
        assert!(range.contains(2, 2));
        assert!(range.contains(3, 3));
        assert!(!range.contains(0, 0));
        assert!(!range.contains(4, 4));
        assert!(!range.contains(1, 0));
    }

    #[test]
    fn test_cell_range_overlaps() {
        let range1 = CellRange::new(1, 1, 3, 3);
        let range2 = CellRange::new(2, 2, 4, 4);
        let range3 = CellRange::new(5, 5, 6, 6);

        assert!(range1.overlaps(&range2));
        assert!(range2.overlaps(&range1));
        assert!(!range1.overlaps(&range3));
        assert!(!range3.overlaps(&range1));
    }

    #[test]
    fn test_cell_range_single() {
        let range = CellRange::single(5, 3);
        assert_eq!(range.start_row, 5);
        assert_eq!(range.start_col, 3);
        assert_eq!(range.end_row, 5);
        assert_eq!(range.end_col, 3);
        assert_eq!(range.cell_count(), 1);
    }

    #[test]
    fn test_validation_store_set_get() {
        let mut store = ValidationStore::new();
        let range = CellRange::new(0, 0, 9, 0);
        let rule = ValidationRule::list_inline(vec!["Yes".into(), "No".into()]);

        store.set(range, rule.clone());

        assert!(store.get(0, 0).is_some());
        assert!(store.get(5, 0).is_some());
        assert!(store.get(9, 0).is_some());
        assert!(store.get(10, 0).is_none());
        assert!(store.get(0, 1).is_none());
    }

    #[test]
    fn test_validation_store_clear_range() {
        let mut store = ValidationStore::new();
        store.set(CellRange::new(0, 0, 5, 5), ValidationRule::new(ValidationType::AnyValue));
        store.set(CellRange::new(10, 10, 15, 15), ValidationRule::new(ValidationType::AnyValue));

        assert_eq!(store.len(), 2);

        // Clear range that overlaps with first rule
        store.clear_range(&CellRange::new(3, 3, 7, 7));

        assert_eq!(store.len(), 1);
        assert!(store.get(0, 0).is_none());
        assert!(store.get(10, 10).is_some());
    }

    #[test]
    fn test_numeric_constraint_builders() {
        let between = NumericConstraint::between(1, 100);
        assert_eq!(between.operator, ComparisonOperator::Between);
        assert_eq!(between.value1, ConstraintValue::Number(1.0));
        assert_eq!(between.value2, Some(ConstraintValue::Number(100.0)));

        let gt = NumericConstraint::greater_than(0);
        assert_eq!(gt.operator, ComparisonOperator::GreaterThan);
        assert_eq!(gt.value1, ConstraintValue::Number(0.0));
        assert!(gt.value2.is_none());
    }

    #[test]
    fn test_validation_rule_builders() {
        let list_rule = ValidationRule::list_inline(vec!["A".into(), "B".into(), "C".into()]);
        assert!(matches!(list_rule.rule_type, ValidationType::List(ListSource::Inline(_))));
        assert!(list_rule.ignore_blank);

        let num_rule = ValidationRule::whole_number(NumericConstraint::between(1, 10))
            .with_ignore_blank(false)
            .with_error_alert(ErrorAlert::stop("Error", "Must be 1-10"));
        assert!(!num_rule.ignore_blank);
        assert!(num_rule.error_alert.is_some());
    }

    #[test]
    fn test_validation_result() {
        let valid = ValidationResult::Valid;
        assert!(valid.is_valid());
        assert!(!valid.is_invalid());

        let invalid = ValidationResult::Invalid {
            rule: ValidationRule::new(ValidationType::AnyValue),
            reason: "test".into(),
        };
        assert!(!invalid.is_valid());
        assert!(invalid.is_invalid());
    }

    #[test]
    fn test_error_alert_builders() {
        let stop = ErrorAlert::stop("Title", "Message");
        assert_eq!(stop.style, ErrorStyle::Stop);

        let warning = ErrorAlert::warning("Title", "Message");
        assert_eq!(warning.style, ErrorStyle::Warning);

        let info = ErrorAlert::info("Title", "Message");
        assert_eq!(info.style, ErrorStyle::Information);
    }

    #[test]
    fn test_serialization() {
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100))
            .with_error_alert(ErrorAlert::stop("Invalid", "Enter 1-100"));

        let json = serde_json::to_string(&rule).unwrap();
        let parsed: ValidationRule = serde_json::from_str(&json).unwrap();

        assert_eq!(rule, parsed);
    }

    // ========================================================================
    // ResolvedList tests
    // ========================================================================

    #[test]
    fn test_resolved_list_from_items() {
        let items = vec![
            "  Yes  ".to_string(),
            "No".to_string(),
            "  Maybe ".to_string(),
        ];
        let resolved = ResolvedList::from_items(items);

        assert_eq!(resolved.items, vec!["Yes", "No", "Maybe"]);
        assert!(!resolved.is_truncated);
        assert!(resolved.source_fingerprint != 0);
    }

    #[test]
    fn test_resolved_list_contains_case_sensitive() {
        let resolved = ResolvedList::from_items(vec![
            "Yes".to_string(),
            "No".to_string(),
        ]);

        // Case-sensitive matching
        assert!(resolved.contains("Yes"));
        assert!(resolved.contains("No"));
        assert!(!resolved.contains("yes")); // Different case = no match
        assert!(!resolved.contains("YES"));
        assert!(!resolved.contains("Maybe"));

        // Whitespace is trimmed
        assert!(resolved.contains("  Yes  "));
    }

    #[test]
    fn test_resolved_list_contains_case_insensitive() {
        let resolved = ResolvedList::from_items(vec![
            "Yes".to_string(),
            "No".to_string(),
        ]);

        // Case-insensitive matching
        assert!(resolved.contains_case_insensitive("Yes"));
        assert!(resolved.contains_case_insensitive("yes"));
        assert!(resolved.contains_case_insensitive("YES"));
        assert!(resolved.contains_case_insensitive("No"));
        assert!(!resolved.contains_case_insensitive("Maybe"));
    }

    #[test]
    fn test_resolved_list_filter() {
        let resolved = ResolvedList::from_items(vec![
            "Open".to_string(),
            "In Progress".to_string(),
            "Closed".to_string(),
            "On Hold".to_string(),
        ]);

        // Case-insensitive substring filtering
        let filtered = resolved.filter("o");
        assert_eq!(filtered, vec!["Open", "In Progress", "Closed", "On Hold"]);

        let filtered = resolved.filter("pro");
        assert_eq!(filtered, vec!["In Progress"]);

        let filtered = resolved.filter("CL");
        assert_eq!(filtered, vec!["Closed"]);

        let filtered = resolved.filter("xyz");
        assert!(filtered.is_empty());
    }

    #[test]
    fn test_resolved_list_truncation() {
        // Create a list larger than MAX_LIST_ITEMS
        let large_items: Vec<String> = (0..MAX_LIST_ITEMS + 100)
            .map(|i| format!("Item{}", i))
            .collect();

        let resolved = ResolvedList::from_items(large_items);

        assert_eq!(resolved.items.len(), MAX_LIST_ITEMS);
        assert!(resolved.is_truncated);
    }

    #[test]
    fn test_resolved_list_fingerprint_changes() {
        let list1 = ResolvedList::from_items(vec!["A".to_string(), "B".to_string()]);
        let list2 = ResolvedList::from_items(vec!["A".to_string(), "C".to_string()]);
        let list3 = ResolvedList::from_items(vec!["A".to_string(), "B".to_string()]);

        // Different content = different fingerprint
        assert_ne!(list1.source_fingerprint, list2.source_fingerprint);

        // Same content = same fingerprint
        assert_eq!(list1.source_fingerprint, list3.source_fingerprint);
    }

    #[test]
    fn test_resolved_list_empty() {
        let resolved = ResolvedList::empty();

        assert!(resolved.items.is_empty());
        assert!(!resolved.is_truncated);
        assert_eq!(resolved.source_fingerprint, 0);
    }

    // ========================================================================
    // Numeric Validation Tests
    // ========================================================================

    #[test]
    fn test_parse_numeric_input_basic() {
        // Basic parsing
        assert_eq!(parse_numeric_input("5", true).unwrap(), 5.0);
        assert_eq!(parse_numeric_input("  5  ", true).unwrap(), 5.0);
        assert_eq!(parse_numeric_input("+5", true).unwrap(), 5.0);
        assert_eq!(parse_numeric_input("-5", true).unwrap(), -5.0);

        // Decimal input
        assert_eq!(parse_numeric_input("3.14", true).unwrap(), 3.14);
        assert_eq!(parse_numeric_input(".5", true).unwrap(), 0.5);
        assert_eq!(parse_numeric_input("-.5", true).unwrap(), -0.5);

        // Empty/invalid
        assert!(parse_numeric_input("", true).is_err());
        assert!(parse_numeric_input("  ", true).is_err());
        assert!(parse_numeric_input("abc", true).is_err());
        assert!(parse_numeric_input("+", true).is_err());
    }

    #[test]
    fn test_parse_numeric_input_whole_number_strictness() {
        // WholeNumber allows integers
        assert_eq!(parse_numeric_input("3", false).unwrap(), 3.0);
        assert_eq!(parse_numeric_input("-10", false).unwrap(), -10.0);
        assert_eq!(parse_numeric_input("+7", false).unwrap(), 7.0);

        // WholeNumber rejects ANY decimal point (strict)
        assert_eq!(
            parse_numeric_input("3.5", false),
            Err(NumericParseError::FractionalNotAllowed)
        );
        assert_eq!(
            parse_numeric_input("3.0", false),
            Err(NumericParseError::FractionalNotAllowed)
        );
        assert_eq!(
            parse_numeric_input("3.", false),
            Err(NumericParseError::FractionalNotAllowed)
        );
        assert_eq!(
            parse_numeric_input(".0", false),
            Err(NumericParseError::FractionalNotAllowed)
        );
    }

    #[test]
    fn test_eval_numeric_constraint_between_inclusive() {
        // Between is inclusive on both ends: a <= x <= b
        assert!(eval_numeric_constraint(1.0, ComparisonOperator::Between, 1.0, Some(100.0)));
        assert!(eval_numeric_constraint(50.0, ComparisonOperator::Between, 1.0, Some(100.0)));
        assert!(eval_numeric_constraint(100.0, ComparisonOperator::Between, 1.0, Some(100.0)));

        // Outside bounds
        assert!(!eval_numeric_constraint(0.0, ComparisonOperator::Between, 1.0, Some(100.0)));
        assert!(!eval_numeric_constraint(101.0, ComparisonOperator::Between, 1.0, Some(100.0)));
    }

    #[test]
    fn test_eval_numeric_constraint_not_between() {
        // NotBetween: x < a || x > b
        assert!(eval_numeric_constraint(0.0, ComparisonOperator::NotBetween, 1.0, Some(100.0)));
        assert!(eval_numeric_constraint(101.0, ComparisonOperator::NotBetween, 1.0, Some(100.0)));

        // Inside bounds = fail
        assert!(!eval_numeric_constraint(1.0, ComparisonOperator::NotBetween, 1.0, Some(100.0)));
        assert!(!eval_numeric_constraint(50.0, ComparisonOperator::NotBetween, 1.0, Some(100.0)));
        assert!(!eval_numeric_constraint(100.0, ComparisonOperator::NotBetween, 1.0, Some(100.0)));
    }

    #[test]
    fn test_eval_numeric_constraint_greater_than() {
        assert!(eval_numeric_constraint(1.0, ComparisonOperator::GreaterThan, 0.0, None));
        assert!(!eval_numeric_constraint(0.0, ComparisonOperator::GreaterThan, 0.0, None));
        assert!(!eval_numeric_constraint(-1.0, ComparisonOperator::GreaterThan, 0.0, None));
    }

    #[test]
    fn test_eval_numeric_constraint_less_than_or_equal() {
        assert!(eval_numeric_constraint(0.5, ComparisonOperator::LessThanOrEqual, 1.0, None));
        assert!(eval_numeric_constraint(1.0, ComparisonOperator::LessThanOrEqual, 1.0, None)); // boundary
        assert!(!eval_numeric_constraint(1.1, ComparisonOperator::LessThanOrEqual, 1.0, None));
    }

    #[test]
    fn test_eval_numeric_constraint_decimal_boundary() {
        // Decimal validation: >=0
        assert!(eval_numeric_constraint(0.0, ComparisonOperator::GreaterThanOrEqual, 0.0, None));
        assert!(eval_numeric_constraint(0.5, ComparisonOperator::GreaterThanOrEqual, 0.0, None));
        assert!(!eval_numeric_constraint(-0.1, ComparisonOperator::GreaterThanOrEqual, 0.0, None));
    }

    #[test]
    fn test_eval_numeric_constraint_equal() {
        assert!(eval_numeric_constraint(42.0, ComparisonOperator::EqualTo, 42.0, None));
        assert!(!eval_numeric_constraint(42.1, ComparisonOperator::EqualTo, 42.0, None));

        assert!(eval_numeric_constraint(42.0, ComparisonOperator::NotEqualTo, 0.0, None));
        assert!(!eval_numeric_constraint(0.0, ComparisonOperator::NotEqualTo, 0.0, None));
    }
}
