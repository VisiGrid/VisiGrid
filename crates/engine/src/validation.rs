//! Data Validation for cells
//!
//! Constrains what users can enter into cells: dropdown lists, number ranges,
//! date limits, and custom formula rules.

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

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
    /// Optional message shown when cell is selected.
    pub input_message: Option<InputMessage>,
    /// Optional error alert shown when validation fails.
    pub error_alert: Option<ErrorAlert>,
}

impl ValidationRule {
    /// Create a new validation rule with the given type.
    pub fn new(rule_type: ValidationType) -> Self {
        Self {
            rule_type,
            ignore_blank: true,
            input_message: None,
            error_alert: None,
        }
    }

    /// Set ignore_blank option.
    pub fn with_ignore_blank(mut self, ignore: bool) -> Self {
        self.ignore_blank = ignore;
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
}
