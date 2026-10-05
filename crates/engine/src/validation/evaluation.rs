//! Shared validation policy. Values are typed; display strings are only used
//! for list matching, whose sources also use displayed labels.
use super::*;
use crate::cell::CellValue;
use crate::formula::eval::{EvalResult, Value};

pub(crate) fn invalid(rule: &ValidationRule, reason: impl Into<String>) -> ValidationResult {
    ValidationResult::Invalid {
        rule: rule.clone(),
        reason: reason.into(),
    }
}

/// Keep the existing strict whole/decimal input grammar. Formula input is
/// evaluated on a private candidate before the result reaches validate_rule.
pub(crate) fn input_value(rule: &ValidationRule, input: &str) -> Result<Value, ValidationResult> {
    if input.trim_start().starts_with('=') {
        return Ok(Value::Empty);
    }
    let decimal = match rule.rule_type {
        ValidationType::WholeNumber(_) => Some(false),
        ValidationType::Decimal(_) => Some(true),
        _ => None,
    };
    if let Some(decimal) = decimal {
        return parse_numeric_input(input, decimal)
            .map(Value::Number)
            .map_err(|error| {
                invalid(
                    rule,
                    if decimal {
                        "Value must be a number"
                    } else if error == NumericParseError::FractionalNotAllowed {
                        "Value must be a whole number (no decimals)"
                    } else {
                        "Value must be a whole number"
                    },
                )
            });
    }
    // Text length and list matching historically check the supplied text.
    if matches!(
        rule.rule_type,
        ValidationType::TextLength(_) | ValidationType::List(_)
    ) {
        return Ok(Value::Text(input.into()));
    }
    Ok(match CellValue::from_input(input) {
        CellValue::Empty => Value::Empty,
        CellValue::Number(n) => Value::Number(n),
        CellValue::Text(text) => Value::Text(text),
        CellValue::Formula { .. } => unreachable!("formula input handled above"),
    })
}

/// A constraint must return one finite number. Do not turn blank into zero,
/// parse a displayed percentage, coerce text/booleans or take an array's first
/// element. A singleton array is a scalar result.
pub(crate) fn constraint_number(result: EvalResult) -> Result<f64, ConstraintResolveError> {
    if result.dimensions() != (1, 1) {
        return Err(ConstraintResolveError::FormulaError(
            "Expected one numeric value, not an array".into(),
        ));
    }
    match result.to_value() {
        Value::Number(n) if n.is_finite() => Ok(n),
        Value::Empty => Err(ConstraintResolveError::BlankConstraint),
        Value::Error(error) => Err(ConstraintResolveError::FormulaError(error)),
        _ => Err(ConstraintResolveError::NotNumeric),
    }
}

pub(crate) fn resolve_constraint(
    value: &ConstraintValue,
    formula: &impl Fn(&str) -> EvalResult,
) -> Result<f64, ConstraintResolveError> {
    match value {
        ConstraintValue::Number(n) => constraint_number(EvalResult::Number(*n)),
        ConstraintValue::CellRef(source) => {
            let text = format!("={}", source.trim().trim_start_matches('='));
            if !matches!(
                crate::formula::parser::parse(&text),
                Ok(crate::formula::parser::Expr::CellRef { .. })
            ) {
                return Err(ConstraintResolveError::InvalidReference(source.clone()));
            }
            let result = formula(source);
            if matches!(&result, EvalResult::Error(error) if error.starts_with("#REF!")) {
                return Err(ConstraintResolveError::InvalidReference(source.clone()));
            }
            constraint_number(result)
        }
        ConstraintValue::Formula(source) => constraint_number(formula(source)),
    }
}

pub(crate) fn validate_rule(
    rule: &ValidationRule,
    value: &Value,
    list_text: &str,
    formula: impl Fn(&str) -> EvalResult,
    list: impl Fn(&ListSource) -> ResolvedList,
) -> ValidationResult {
    if rule.ignore_blank
        && (matches!(value, Value::Empty) || matches!(value, Value::Text(s) if s.trim().is_empty()))
    {
        return ValidationResult::Valid;
    }
    let (constraint, number, type_name) = match &rule.rule_type {
        ValidationType::Custom(source) => {
            let result = formula(source);
            if result.dimensions() != (1, 1) {
                return invalid(
                    rule,
                    "Validation formula error: expected one logical value, not an array",
                );
            }
            if matches!(result.to_value(), Value::Number(n) if !n.is_finite()) {
                return invalid(rule, "Validation formula error: non-finite result");
            }
            return match result.to_bool() {
                Ok(true) => ValidationResult::Valid,
                Ok(false) => invalid(rule, "Value does not satisfy the custom rule"),
                Err(error) => invalid(rule, format!("Validation formula error: {error}")),
            };
        }
        ValidationType::List(source) => {
            let list = list(source);
            return if list.items.is_empty() || list.contains(list_text.trim()) {
                ValidationResult::Valid
            } else {
                let preview: Vec<_> = list.items.iter().take(5).map(String::as_str).collect();
                invalid(
                    rule,
                    format!(
                        "Value must be one of: {}{}",
                        preview.join(", "),
                        if list.items.len() > 5 { ", ..." } else { "" }
                    ),
                )
            };
        }
        ValidationType::TextLength(constraint) => {
            if let Value::Error(error) = value {
                return invalid(rule, format!("Cell error: {error}"));
            }
            (
                constraint,
                value.to_text().chars().count() as f64,
                "text length",
            )
        }
        numeric => {
            let (constraint, type_name) = match numeric {
                ValidationType::WholeNumber(c) => (c, "whole number"),
                ValidationType::Decimal(c) => (c, "number"),
                ValidationType::Date(c) => (c, "date serial"),
                ValidationType::Time(c) => (c, "time serial"),
                _ => unreachable!(),
            };
            let number = match value {
                Value::Number(n) if n.is_finite() => *n,
                _ => return invalid(rule, format!("Value must be a {type_name}")),
            };
            if matches!(numeric, ValidationType::WholeNumber(_)) && number.fract() != 0.0 {
                return invalid(rule, "Value must be a whole number (no decimals)");
            }
            (constraint, number, type_name)
        }
    };
    let first = match resolve_constraint(&constraint.value1, &formula) {
        Ok(n) => n,
        Err(error) => return invalid(rule, format!("Validation constraint error: {error}")),
    };
    let second = match constraint
        .value2
        .as_ref()
        .map(|v| resolve_constraint(v, &formula))
        .transpose()
    {
        Ok(n) => n,
        Err(error) => return invalid(rule, format!("Validation constraint error: {error}")),
    };
    if eval_numeric_constraint(number, constraint.operator, first, second) {
        return ValidationResult::Valid;
    }
    let reason = match constraint.operator {
        ComparisonOperator::Between => format!(
            "{type_name} must be between {first} and {}",
            second.unwrap_or(first)
        ),
        ComparisonOperator::NotBetween => format!(
            "{type_name} must not be between {first} and {}",
            second.unwrap_or(first)
        ),
        ComparisonOperator::EqualTo => format!("{type_name} must equal {first}"),
        ComparisonOperator::NotEqualTo => format!("{type_name} must not equal {first}"),
        ComparisonOperator::GreaterThan => format!("{type_name} must be greater than {first}"),
        ComparisonOperator::LessThan => format!("{type_name} must be less than {first}"),
        ComparisonOperator::GreaterThanOrEqual => format!("{type_name} must be at least {first}"),
        ComparisonOperator::LessThanOrEqual => format!("{type_name} must be at most {first}"),
    };
    invalid(rule, reason)
}
