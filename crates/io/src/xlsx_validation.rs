//! XLSX Data Validation import/export support
//!
//! This module handles mapping between VisiGrid's validation model and Excel's
//! dataValidation XML format. Both import and export share these mappings to
//! ensure consistency.
//!
//! ## Scope (Phase 5A)
//! - List validation (inline, range, named range)
//! - WholeNumber validation (all operators)
//! - Decimal validation (all operators)
//!
//! ## Key gotchas
//! - Excel's `showDropDown="1"` means HIDE dropdown (inverted from VisiGrid)
//! - Excel's `allowBlank="1"` maps to VisiGrid's `ignore_blank: true`

use rust_xlsxwriter::{DataValidation, DataValidationErrorStyle, DataValidationRule, Formula};
use visigrid_engine::validation::{
    ComparisonOperator, ConstraintValue, ErrorStyle, ListSource, NumericConstraint,
    ValidationRule, ValidationType,
};

// ============================================================================
// Export: VisiGrid -> Excel
// ============================================================================

/// Convert a VisiGrid ValidationRule to rust_xlsxwriter DataValidation.
///
/// Returns None if the validation type is not supported for export (Date, Time,
/// TextLength, Custom are deferred to Phase 5B).
pub fn rule_to_xlsx(rule: &ValidationRule) -> Option<DataValidation> {
    let mut dv = match &rule.rule_type {
        ValidationType::List(source) => list_to_xlsx(source)?,
        ValidationType::WholeNumber(constraint) => whole_number_to_xlsx(constraint)?,
        ValidationType::Decimal(constraint) => decimal_to_xlsx(constraint)?,
        // Phase 5B: Date, Time, TextLength, Custom
        _ => return None,
    };

    // Common options
    dv = dv.ignore_blank(rule.ignore_blank);

    // CRITICAL: Excel inverts this!
    // Excel: showDropDown="1" means HIDE the dropdown
    // VisiGrid: show_dropdown=true means SHOW the dropdown
    // rust_xlsxwriter: show_dropdown(true) sets Excel's showDropDown="1" (hide)
    // So we pass the OPPOSITE of our value
    if matches!(rule.rule_type, ValidationType::List(_)) {
        if !rule.show_dropdown {
            // VisiGrid wants to hide -> tell rust_xlsxwriter to "show" (which sets Excel's hide flag)
            dv = dv.show_dropdown(true);
        }
    }

    // Input message (if present)
    if let Some(msg) = &rule.input_message {
        if msg.show {
            dv = dv.set_input_title(&msg.title).ok()?;
            dv = dv.set_input_message(&msg.message).ok()?;
        }
    }

    // Error alert (if present)
    if let Some(alert) = &rule.error_alert {
        if alert.show {
            dv = dv.set_error_title(&alert.title).ok()?;
            dv = dv.set_error_message(&alert.message).ok()?;
            dv = match alert.style {
                ErrorStyle::Stop => dv.set_error_style(DataValidationErrorStyle::Stop),
                ErrorStyle::Warning => dv.set_error_style(DataValidationErrorStyle::Warning),
                ErrorStyle::Information => dv.set_error_style(DataValidationErrorStyle::Information),
            };
        }
    }

    Some(dv)
}

/// Convert List validation to Excel DataValidation
fn list_to_xlsx(source: &ListSource) -> Option<DataValidation> {
    let dv = DataValidation::new();

    Some(match source {
        ListSource::Inline(items) => {
            // rust_xlsxwriter handles quoting/escaping
            let refs: Vec<&str> = items.iter().map(|s| s.as_str()).collect();
            dv.allow_list_strings(&refs).ok()?
        }
        ListSource::Range(range_ref) => {
            // Range reference like "=Sheet1!$A$1:$A$10" or "=$A$1:$A$10"
            let formula = if range_ref.starts_with('=') {
                Formula::new(&range_ref[1..]) // Strip leading =
            } else {
                Formula::new(range_ref)
            };
            dv.allow_list_formula(formula)
        }
        ListSource::NamedRange(name) => {
            // Named range like "StatusOptions"
            dv.allow_list_formula(Formula::new(name))
        }
    })
}

/// Convert WholeNumber validation to Excel DataValidation
///
/// Uses either allow_whole_number (for literal numbers) or
/// allow_whole_number_formula (for cell references/formulas).
fn whole_number_to_xlsx(constraint: &NumericConstraint) -> Option<DataValidation> {
    let dv = DataValidation::new();

    // Check if any constraint value is a cell reference or formula
    let uses_formula = is_formula_constraint(&constraint.value1)
        || constraint.value2.as_ref().map_or(false, is_formula_constraint);

    if uses_formula {
        // Use formula-based validation
        let rule = operator_to_xlsx_formula_rule(
            &constraint.operator,
            &constraint.value1,
            constraint.value2.as_ref(),
        );
        Some(dv.allow_whole_number_formula(rule))
    } else {
        // Use numeric validation
        let rule = operator_to_xlsx_i32_rule(
            &constraint.operator,
            &constraint.value1,
            constraint.value2.as_ref(),
        );
        Some(dv.allow_whole_number(rule))
    }
}

/// Convert Decimal validation to Excel DataValidation
fn decimal_to_xlsx(constraint: &NumericConstraint) -> Option<DataValidation> {
    let dv = DataValidation::new();

    // Check if any constraint value is a cell reference or formula
    let uses_formula = is_formula_constraint(&constraint.value1)
        || constraint.value2.as_ref().map_or(false, is_formula_constraint);

    if uses_formula {
        // Use formula-based validation
        let rule = operator_to_xlsx_formula_rule(
            &constraint.operator,
            &constraint.value1,
            constraint.value2.as_ref(),
        );
        Some(dv.allow_decimal_number_formula(rule))
    } else {
        // Use numeric validation
        let rule = operator_to_xlsx_f64_rule(
            &constraint.operator,
            &constraint.value1,
            constraint.value2.as_ref(),
        );
        Some(dv.allow_decimal_number(rule))
    }
}

/// Check if a constraint value is a cell reference or formula (not a literal number)
fn is_formula_constraint(value: &ConstraintValue) -> bool {
    matches!(value, ConstraintValue::CellRef(_) | ConstraintValue::Formula(_))
}

/// Convert operator + values to DataValidationRule<i32> for whole numbers
fn operator_to_xlsx_i32_rule(
    op: &ComparisonOperator,
    value1: &ConstraintValue,
    value2: Option<&ConstraintValue>,
) -> DataValidationRule<i32> {
    let v1 = constraint_value_to_i32(value1);
    let v2 = value2.map(constraint_value_to_i32).unwrap_or(v1);

    match op {
        ComparisonOperator::Between => DataValidationRule::Between(v1, v2),
        ComparisonOperator::NotBetween => DataValidationRule::NotBetween(v1, v2),
        ComparisonOperator::EqualTo => DataValidationRule::EqualTo(v1),
        ComparisonOperator::NotEqualTo => DataValidationRule::NotEqualTo(v1),
        ComparisonOperator::GreaterThan => DataValidationRule::GreaterThan(v1),
        ComparisonOperator::LessThan => DataValidationRule::LessThan(v1),
        ComparisonOperator::GreaterThanOrEqual => DataValidationRule::GreaterThanOrEqualTo(v1),
        ComparisonOperator::LessThanOrEqual => DataValidationRule::LessThanOrEqualTo(v1),
    }
}

/// Convert operator + values to DataValidationRule<f64> for decimals
fn operator_to_xlsx_f64_rule(
    op: &ComparisonOperator,
    value1: &ConstraintValue,
    value2: Option<&ConstraintValue>,
) -> DataValidationRule<f64> {
    let v1 = constraint_value_to_f64(value1);
    let v2 = value2.map(constraint_value_to_f64).unwrap_or(v1);

    match op {
        ComparisonOperator::Between => DataValidationRule::Between(v1, v2),
        ComparisonOperator::NotBetween => DataValidationRule::NotBetween(v1, v2),
        ComparisonOperator::EqualTo => DataValidationRule::EqualTo(v1),
        ComparisonOperator::NotEqualTo => DataValidationRule::NotEqualTo(v1),
        ComparisonOperator::GreaterThan => DataValidationRule::GreaterThan(v1),
        ComparisonOperator::LessThan => DataValidationRule::LessThan(v1),
        ComparisonOperator::GreaterThanOrEqual => DataValidationRule::GreaterThanOrEqualTo(v1),
        ComparisonOperator::LessThanOrEqual => DataValidationRule::LessThanOrEqualTo(v1),
    }
}

/// Convert operator + values to DataValidationRule<Formula> for cell refs/formulas
fn operator_to_xlsx_formula_rule(
    op: &ComparisonOperator,
    value1: &ConstraintValue,
    value2: Option<&ConstraintValue>,
) -> DataValidationRule<Formula> {
    let v1 = constraint_value_to_formula(value1);
    let v2 = value2.map(constraint_value_to_formula).unwrap_or_else(|| v1.clone());

    match op {
        ComparisonOperator::Between => DataValidationRule::Between(v1, v2),
        ComparisonOperator::NotBetween => DataValidationRule::NotBetween(v1, v2),
        ComparisonOperator::EqualTo => DataValidationRule::EqualTo(v1),
        ComparisonOperator::NotEqualTo => DataValidationRule::NotEqualTo(v1),
        ComparisonOperator::GreaterThan => DataValidationRule::GreaterThan(v1),
        ComparisonOperator::LessThan => DataValidationRule::LessThan(v1),
        ComparisonOperator::GreaterThanOrEqual => DataValidationRule::GreaterThanOrEqualTo(v1),
        ComparisonOperator::LessThanOrEqual => DataValidationRule::LessThanOrEqualTo(v1),
    }
}

/// Convert ConstraintValue to i32 (for whole number validation)
fn constraint_value_to_i32(value: &ConstraintValue) -> i32 {
    match value {
        ConstraintValue::Number(n) => *n as i32,
        // Cell refs and formulas shouldn't reach here (we check is_formula_constraint first)
        _ => 0,
    }
}

/// Convert ConstraintValue to f64 (for decimal validation)
fn constraint_value_to_f64(value: &ConstraintValue) -> f64 {
    match value {
        ConstraintValue::Number(n) => *n,
        // Cell refs and formulas shouldn't reach here
        _ => 0.0,
    }
}

/// Convert ConstraintValue to Formula (for cell ref/formula validation)
fn constraint_value_to_formula(value: &ConstraintValue) -> Formula {
    match value {
        ConstraintValue::Number(n) => {
            // Format integers without decimal point
            if n.fract() == 0.0 && n.abs() < 1e15 {
                Formula::new(&format!("{}", *n as i64))
            } else {
                Formula::new(&format!("{}", n))
            }
        }
        ConstraintValue::CellRef(cell_ref) => {
            // Cell reference like "A1" or "Sheet2!B5"
            // Formula::new expects no leading =
            let cleaned = cell_ref.strip_prefix('=').unwrap_or(cell_ref);
            Formula::new(cleaned)
        }
        ConstraintValue::Formula(formula) => {
            // Formula like "=TODAY()" or "=MAX(A1:A10)"
            let cleaned = formula.strip_prefix('=').unwrap_or(formula);
            Formula::new(cleaned)
        }
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_list_inline_export() {
        let rule = ValidationRule::list_inline(vec!["Yes".into(), "No".into(), "Maybe".into()]);
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "List inline should export");
    }

    #[test]
    fn test_list_range_export() {
        let rule = ValidationRule::list_range("=$A$1:$A$10");
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "List range should export");
    }

    #[test]
    fn test_list_named_range_export() {
        let rule = ValidationRule::new(ValidationType::List(ListSource::NamedRange("StatusOptions".into())));
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "List named range should export");
    }

    #[test]
    fn test_whole_number_between_export() {
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "WholeNumber between should export");
    }

    #[test]
    fn test_whole_number_with_cell_ref() {
        let constraint = NumericConstraint {
            operator: ComparisonOperator::LessThanOrEqual,
            value1: ConstraintValue::CellRef("=D1".into()),
            value2: None,
        };
        let rule = ValidationRule::whole_number(constraint);
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "WholeNumber with cell ref should export");
    }

    #[test]
    fn test_decimal_greater_than_export() {
        let rule = ValidationRule::decimal(NumericConstraint::greater_than(0.0));
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some(), "Decimal greater_than should export");
    }

    #[test]
    fn test_all_operators() {
        let operators = [
            ComparisonOperator::Between,
            ComparisonOperator::NotBetween,
            ComparisonOperator::EqualTo,
            ComparisonOperator::NotEqualTo,
            ComparisonOperator::GreaterThan,
            ComparisonOperator::LessThan,
            ComparisonOperator::GreaterThanOrEqual,
            ComparisonOperator::LessThanOrEqual,
        ];

        for op in operators {
            let constraint = NumericConstraint {
                operator: op.clone(),
                value1: ConstraintValue::Number(1.0),
                value2: if matches!(op, ComparisonOperator::Between | ComparisonOperator::NotBetween) {
                    Some(ConstraintValue::Number(100.0))
                } else {
                    None
                },
            };
            let rule = ValidationRule::whole_number(constraint);
            let dv = rule_to_xlsx(&rule);
            assert!(dv.is_some(), "Operator {:?} should export", op);
        }
    }

    #[test]
    fn test_ignore_blank_mapping() {
        let mut rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        rule.ignore_blank = true;
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some());

        rule.ignore_blank = false;
        let dv = rule_to_xlsx(&rule);
        assert!(dv.is_some());
    }

    #[test]
    fn test_unsupported_types_return_none() {
        // Date, Time, TextLength, Custom should return None for now
        let rule = ValidationRule::new(ValidationType::Date(NumericConstraint::between(0, 100)));
        assert!(rule_to_xlsx(&rule).is_none(), "Date should not export yet");

        let rule = ValidationRule::new(ValidationType::Time(NumericConstraint::between(0, 1)));
        assert!(rule_to_xlsx(&rule).is_none(), "Time should not export yet");

        let rule = ValidationRule::new(ValidationType::TextLength(NumericConstraint::between(1, 100)));
        assert!(rule_to_xlsx(&rule).is_none(), "TextLength should not export yet");

        let rule = ValidationRule::custom("=A1>0");
        assert!(rule_to_xlsx(&rule).is_none(), "Custom should not export yet");
    }
}
