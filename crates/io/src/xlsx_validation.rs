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
// Import: Excel -> VisiGrid
// ============================================================================

use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;
use visigrid_engine::validation::{CellRange, ErrorAlert, InputMessage};
use zip::ZipArchive;

/// A validation rule parsed from XLSX, ready to be added to a sheet
#[derive(Debug, Clone)]
pub struct ImportedValidation {
    pub range: CellRange,
    pub rule: ValidationRule,
}

/// Parse all validation rules from an XLSX file for a specific sheet.
///
/// Returns a list of (range, rule) pairs that can be added to the sheet's ValidationStore.
pub fn parse_sheet_validations(
    xlsx_path: &Path,
    sheet_name: &str,
) -> Result<Vec<ImportedValidation>, String> {
    let file = std::fs::File::open(xlsx_path)
        .map_err(|e| format!("Failed to open XLSX file: {}", e))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|e| format!("Failed to read XLSX as ZIP: {}", e))?;

    // Step 1: Find the worksheet XML path for this sheet name
    let xml_path = find_worksheet_xml_path(&mut archive, sheet_name)?;

    // Step 2: Read and parse the worksheet XML
    let xml_content = read_zip_file(&mut archive, &xml_path)?;

    // Step 3: Parse <dataValidation> elements
    parse_validations_from_xml(&xml_content)
}

/// Find the worksheet XML path for a given sheet name.
///
/// This requires parsing:
/// 1. xl/workbook.xml to find the sheet's rId
/// 2. xl/_rels/workbook.xml.rels to map rId to the actual XML path
fn find_worksheet_xml_path<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    sheet_name: &str,
) -> Result<String, String> {
    // Parse workbook.xml to find sheet rId
    let workbook_xml = read_zip_file(archive, "xl/workbook.xml")?;
    let rid = find_sheet_rid(&workbook_xml, sheet_name)?;

    // Parse workbook.xml.rels to find the target path
    let rels_xml = read_zip_file(archive, "xl/_rels/workbook.xml.rels")?;
    let target = find_relationship_target(&rels_xml, &rid)?;

    // Target is relative to xl/, so prepend it
    Ok(format!("xl/{}", target))
}

/// Find the rId for a sheet name in workbook.xml
fn find_sheet_rid(workbook_xml: &str, sheet_name: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut rid = None;

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            name = Some(
                                String::from_utf8_lossy(&attr.value).to_string()
                            );
                        }
                        b"r:id" => {
                            rid = Some(
                                String::from_utf8_lossy(&attr.value).to_string()
                            );
                        }
                        _ => {}
                    }
                }

                if name.as_deref() == Some(sheet_name) {
                    if let Some(r) = rid {
                        return Ok(r);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Err(format!("Sheet '{}' not found in workbook.xml", sheet_name))
}

/// Find the target path for a relationship ID in workbook.xml.rels
fn find_relationship_target(rels_xml: &str, rid: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => {
                            id = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        b"Target" => {
                            target = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }

                if id.as_deref() == Some(rid) {
                    if let Some(t) = target {
                        return Ok(t);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Err(format!("Relationship '{}' not found", rid))
}

/// Read a file from the ZIP archive
fn read_zip_file<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String, String> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| format!("File '{}' not found in XLSX: {}", path, e))?;

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| format!("Failed to read '{}': {}", path, e))?;

    Ok(content)
}

/// Parse <dataValidation> elements from worksheet XML
fn parse_validations_from_xml(xml: &str) -> Result<Vec<ImportedValidation>, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut validations = Vec::new();
    let mut buf = Vec::new();
    let mut in_data_validation = false;
    let mut current_attrs: HashMap<String, String> = HashMap::new();
    let mut formula1: Option<String> = None;
    let mut formula2: Option<String> = None;
    let mut in_formula1 = false;
    let mut in_formula2 = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.name().as_ref() == b"dataValidation" => {
                in_data_validation = true;
                current_attrs.clear();
                formula1 = None;
                formula2 = None;

                // Collect attributes
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let value = String::from_utf8_lossy(&attr.value).to_string();
                    current_attrs.insert(key, value);
                }
            }
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"dataValidation" => {
                // Self-closing <dataValidation /> - collect attrs and process
                current_attrs.clear();
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let value = String::from_utf8_lossy(&attr.value).to_string();
                    current_attrs.insert(key, value);
                }

                // Process this validation (no formulas since self-closing)
                if let Some(sqref) = current_attrs.get("sqref") {
                    if let Some(imported) = parse_single_validation(&current_attrs, None, None) {
                        for range in parse_sqref(sqref) {
                            validations.push(ImportedValidation {
                                range,
                                rule: imported.clone(),
                            });
                        }
                    }
                }
            }
            Ok(Event::Start(ref e)) if in_data_validation && e.name().as_ref() == b"formula1" => {
                in_formula1 = true;
            }
            Ok(Event::Start(ref e)) if in_data_validation && e.name().as_ref() == b"formula2" => {
                in_formula2 = true;
            }
            Ok(Event::Text(ref e)) if in_formula1 => {
                formula1 = Some(e.unescape().unwrap_or_default().to_string());
            }
            Ok(Event::Text(ref e)) if in_formula2 => {
                formula2 = Some(e.unescape().unwrap_or_default().to_string());
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"formula1" => {
                in_formula1 = false;
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"formula2" => {
                in_formula2 = false;
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"dataValidation" => {
                in_data_validation = false;

                // Process the collected validation
                if let Some(sqref) = current_attrs.get("sqref") {
                    if let Some(imported) =
                        parse_single_validation(&current_attrs, formula1.as_deref(), formula2.as_deref())
                    {
                        for range in parse_sqref(sqref) {
                            validations.push(ImportedValidation {
                                range,
                                rule: imported.clone(),
                            });
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(validations)
}

/// Parse a single <dataValidation> element into a ValidationRule
fn parse_single_validation(
    attrs: &HashMap<String, String>,
    formula1: Option<&str>,
    formula2: Option<&str>,
) -> Option<ValidationRule> {
    let validation_type = attrs.get("type").map(|s| s.as_str()).unwrap_or("none");

    let rule_type = match validation_type {
        "list" => {
            let source = parse_list_source(formula1?)?;
            ValidationType::List(source)
        }
        "whole" => {
            let constraint = parse_numeric_constraint(attrs, formula1, formula2)?;
            ValidationType::WholeNumber(constraint)
        }
        "decimal" => {
            let constraint = parse_numeric_constraint(attrs, formula1, formula2)?;
            ValidationType::Decimal(constraint)
        }
        // Phase 5B types - skip for now
        "date" | "time" | "textLength" | "custom" => return None,
        // "none" or unknown - skip
        _ => return None,
    };

    let mut rule = ValidationRule::new(rule_type);

    // allowBlank: "1" = true, "0" or absent = false
    rule.ignore_blank = attrs.get("allowBlank").map(|v| v == "1").unwrap_or(false);

    // showDropDown: INVERTED! "1" = hide dropdown, "0" or absent = show
    // VisiGrid: show_dropdown=true means show
    rule.show_dropdown = attrs.get("showDropDown").map(|v| v != "1").unwrap_or(true);

    // Input message
    let show_input = attrs.get("showInputMessage").map(|v| v == "1").unwrap_or(false);
    if show_input {
        let title = attrs.get("promptTitle").cloned().unwrap_or_default();
        let message = attrs.get("prompt").cloned().unwrap_or_default();
        if !title.is_empty() || !message.is_empty() {
            rule.input_message = Some(InputMessage {
                show: true,
                title,
                message,
            });
        }
    }

    // Error alert
    let show_error = attrs.get("showErrorMessage").map(|v| v == "1").unwrap_or(false);
    if show_error {
        let title = attrs.get("errorTitle").cloned().unwrap_or_default();
        let message = attrs.get("error").cloned().unwrap_or_default();
        let style = match attrs.get("errorStyle").map(|s| s.as_str()) {
            Some("warning") => ErrorStyle::Warning,
            Some("information") => ErrorStyle::Information,
            _ => ErrorStyle::Stop, // Default
        };
        if !title.is_empty() || !message.is_empty() {
            rule.error_alert = Some(ErrorAlert {
                show: true,
                style,
                title,
                message,
            });
        }
    }

    Some(rule)
}

/// Parse list source from formula1
fn parse_list_source(formula1: &str) -> Option<ListSource> {
    let formula1 = formula1.trim();

    if formula1.is_empty() {
        return None;
    }

    // Inline list: starts and ends with quotes, comma-separated
    // e.g., "Yes,No,Maybe" or "\"Yes\",\"No\""
    if formula1.starts_with('"') && formula1.ends_with('"') {
        let inner = &formula1[1..formula1.len() - 1];
        let items: Vec<String> = inner.split(',').map(|s| s.trim().to_string()).collect();
        return Some(ListSource::Inline(items));
    }

    // Range reference: contains $ or : or !
    // e.g., $A$1:$A$10, Sheet2!$B$1:$B$20
    if formula1.contains('$') || formula1.contains(':') || formula1.contains('!') {
        // Prepend = for VisiGrid's Range format
        return Some(ListSource::Range(format!("={}", formula1)));
    }

    // Named range: simple identifier
    // e.g., StatusOptions
    Some(ListSource::NamedRange(formula1.to_string()))
}

/// Parse numeric constraint from attributes and formulas
fn parse_numeric_constraint(
    attrs: &HashMap<String, String>,
    formula1: Option<&str>,
    formula2: Option<&str>,
) -> Option<NumericConstraint> {
    let operator = parse_operator(attrs.get("operator").map(|s| s.as_str()))?;
    let value1 = parse_constraint_value(formula1?)?;
    let value2 = if matches!(operator, ComparisonOperator::Between | ComparisonOperator::NotBetween) {
        Some(parse_constraint_value(formula2?)?)
    } else {
        None
    };

    Some(NumericConstraint {
        operator,
        value1,
        value2,
    })
}

/// Parse comparison operator from Excel attribute
fn parse_operator(op: Option<&str>) -> Option<ComparisonOperator> {
    Some(match op.unwrap_or("between") {
        "between" => ComparisonOperator::Between,
        "notBetween" => ComparisonOperator::NotBetween,
        "equal" => ComparisonOperator::EqualTo,
        "notEqual" => ComparisonOperator::NotEqualTo,
        "greaterThan" => ComparisonOperator::GreaterThan,
        "lessThan" => ComparisonOperator::LessThan,
        "greaterThanOrEqual" => ComparisonOperator::GreaterThanOrEqual,
        "lessThanOrEqual" => ComparisonOperator::LessThanOrEqual,
        _ => return None,
    })
}

/// Parse a constraint value (number, cell reference, or formula)
fn parse_constraint_value(value: &str) -> Option<ConstraintValue> {
    let value = value.trim();

    if value.is_empty() {
        return None;
    }

    // Try to parse as number first
    if let Ok(n) = value.parse::<f64>() {
        return Some(ConstraintValue::Number(n));
    }

    // Cell reference: starts with letter or $, contains no functions
    // e.g., A1, $A$1, Sheet2!B5
    if is_cell_reference(value) {
        return Some(ConstraintValue::CellRef(format!("={}", value)));
    }

    // Otherwise treat as formula
    Some(ConstraintValue::Formula(format!("={}", value)))
}

/// Check if a string looks like a cell reference (not a formula)
fn is_cell_reference(s: &str) -> bool {
    // Cell refs: A1, $A$1, Sheet1!A1, 'Sheet Name'!$A$1
    // NOT formulas: TODAY(), MAX(A1:A10), A1+B1

    // If it contains parentheses or operators, it's likely a formula
    if s.contains('(') || s.contains('+') || s.contains('-') || s.contains('*') || s.contains('/') {
        return false;
    }

    // Simple heuristic: cell refs match [Sheet!][$]Col[$]Row pattern
    // This is a simplified check - just verify it's not obviously a formula
    true
}

/// Parse Excel sqref into CellRange(s)
///
/// sqref can be:
/// - Single cell: "A1" -> CellRange(0,0,0,0)
/// - Single range: "A1:B10" -> CellRange(0,0,9,1)
/// - Multiple: "A1:A10 C1:C10" -> two ranges
fn parse_sqref(sqref: &str) -> Vec<CellRange> {
    sqref
        .split_whitespace()
        .filter_map(|part| parse_single_range(part))
        .collect()
}

/// Parse a single range reference like "A1" or "A1:B10"
fn parse_single_range(range_str: &str) -> Option<CellRange> {
    let range_str = range_str.trim();

    if let Some((start, end)) = range_str.split_once(':') {
        // Range: A1:B10
        let (start_row, start_col) = parse_cell_ref(start)?;
        let (end_row, end_col) = parse_cell_ref(end)?;
        Some(CellRange::new(start_row, start_col, end_row, end_col))
    } else {
        // Single cell: A1 -> A1:A1
        let (row, col) = parse_cell_ref(range_str)?;
        Some(CellRange::new(row, col, row, col))
    }
}

/// Parse a cell reference like "A1" or "$A$1" into (row, col)
fn parse_cell_ref(cell_ref: &str) -> Option<(usize, usize)> {
    let cell_ref = cell_ref.replace('$', ""); // Strip $ signs
    let cell_ref = cell_ref.trim();

    // Find where letters end and numbers begin
    let mut col_end = 0;
    for (i, c) in cell_ref.chars().enumerate() {
        if c.is_ascii_digit() {
            col_end = i;
            break;
        }
    }

    if col_end == 0 {
        return None;
    }

    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    let col = col_from_letters(col_str)?;
    let row: usize = row_str.parse().ok()?;

    // Excel rows are 1-indexed, VisiGrid is 0-indexed
    Some((row.saturating_sub(1), col))
}

/// Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, ...)
fn col_from_letters(letters: &str) -> Option<usize> {
    let mut col = 0usize;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    Some(col.saturating_sub(1))
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

    // ========================================================================
    // Import Tests
    // ========================================================================

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("$A$1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("AB10"), Some((9, 27)));
    }

    #[test]
    fn test_col_from_letters() {
        assert_eq!(col_from_letters("A"), Some(0));
        assert_eq!(col_from_letters("B"), Some(1));
        assert_eq!(col_from_letters("Z"), Some(25));
        assert_eq!(col_from_letters("AA"), Some(26));
        assert_eq!(col_from_letters("AB"), Some(27));
        assert_eq!(col_from_letters("AZ"), Some(51));
        assert_eq!(col_from_letters("BA"), Some(52));
    }

    #[test]
    fn test_parse_sqref_single_cell() {
        let ranges = parse_sqref("A1");
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0], CellRange::new(0, 0, 0, 0));
    }

    #[test]
    fn test_parse_sqref_range() {
        let ranges = parse_sqref("A1:B10");
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0], CellRange::new(0, 0, 9, 1));
    }

    #[test]
    fn test_parse_sqref_multiple() {
        let ranges = parse_sqref("A1:A10 C1:C10");
        assert_eq!(ranges.len(), 2);
        assert_eq!(ranges[0], CellRange::new(0, 0, 9, 0));
        assert_eq!(ranges[1], CellRange::new(0, 2, 9, 2));
    }

    #[test]
    fn test_parse_list_source_inline() {
        let source = parse_list_source("\"Yes,No,Maybe\"");
        assert!(matches!(source, Some(ListSource::Inline(items)) if items == vec!["Yes", "No", "Maybe"]));
    }

    #[test]
    fn test_parse_list_source_range() {
        let source = parse_list_source("$A$1:$A$10");
        assert!(matches!(source, Some(ListSource::Range(r)) if r == "=$A$1:$A$10"));

        let source = parse_list_source("Sheet2!$B$1:$B$20");
        assert!(matches!(source, Some(ListSource::Range(r)) if r == "=Sheet2!$B$1:$B$20"));
    }

    #[test]
    fn test_parse_list_source_named_range() {
        let source = parse_list_source("StatusOptions");
        assert!(matches!(source, Some(ListSource::NamedRange(n)) if n == "StatusOptions"));
    }

    #[test]
    fn test_parse_operator() {
        assert_eq!(parse_operator(Some("between")), Some(ComparisonOperator::Between));
        assert_eq!(parse_operator(Some("notBetween")), Some(ComparisonOperator::NotBetween));
        assert_eq!(parse_operator(Some("equal")), Some(ComparisonOperator::EqualTo));
        assert_eq!(parse_operator(Some("notEqual")), Some(ComparisonOperator::NotEqualTo));
        assert_eq!(parse_operator(Some("greaterThan")), Some(ComparisonOperator::GreaterThan));
        assert_eq!(parse_operator(Some("lessThan")), Some(ComparisonOperator::LessThan));
        assert_eq!(parse_operator(Some("greaterThanOrEqual")), Some(ComparisonOperator::GreaterThanOrEqual));
        assert_eq!(parse_operator(Some("lessThanOrEqual")), Some(ComparisonOperator::LessThanOrEqual));
    }

    #[test]
    fn test_parse_constraint_value_number() {
        let val = parse_constraint_value("42");
        assert!(matches!(val, Some(ConstraintValue::Number(n)) if (n - 42.0).abs() < 0.001));

        let val = parse_constraint_value("3.14");
        assert!(matches!(val, Some(ConstraintValue::Number(n)) if (n - 3.14).abs() < 0.001));
    }

    #[test]
    fn test_parse_constraint_value_cell_ref() {
        let val = parse_constraint_value("$A$1");
        assert!(matches!(val, Some(ConstraintValue::CellRef(r)) if r == "=$A$1"));
    }

    #[test]
    fn test_parse_validations_from_xml_list() {
        let xml = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations count="1">
                    <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="B2:B100">
                        <formula1>"Open,In Progress,Closed"</formula1>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml).unwrap();
        assert_eq!(validations.len(), 1);

        let v = &validations[0];
        assert_eq!(v.range, CellRange::new(1, 1, 99, 1));
        assert!(v.rule.ignore_blank);
        assert!(v.rule.show_dropdown); // showDropDown="0" in Excel means show

        match &v.rule.rule_type {
            ValidationType::List(ListSource::Inline(items)) => {
                assert_eq!(items, &vec!["Open", "In Progress", "Closed"]);
            }
            _ => panic!("Expected inline list"),
        }
    }

    #[test]
    fn test_parse_validations_from_xml_whole_number() {
        let xml = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations count="1">
                    <dataValidation type="whole" operator="between" allowBlank="0" sqref="C2:C50">
                        <formula1>1</formula1>
                        <formula2>100</formula2>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml).unwrap();
        assert_eq!(validations.len(), 1);

        let v = &validations[0];
        assert_eq!(v.range, CellRange::new(1, 2, 49, 2));
        assert!(!v.rule.ignore_blank);

        match &v.rule.rule_type {
            ValidationType::WholeNumber(c) => {
                assert!(matches!(c.operator, ComparisonOperator::Between));
                assert!(matches!(c.value1, ConstraintValue::Number(n) if (n - 1.0).abs() < 0.001));
                assert!(matches!(c.value2, Some(ConstraintValue::Number(n)) if (n - 100.0).abs() < 0.001));
            }
            _ => panic!("Expected whole number"),
        }
    }

    #[test]
    fn test_parse_validations_from_xml_with_messages() {
        let xml = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations count="1">
                    <dataValidation type="decimal" operator="greaterThan"
                        allowBlank="1"
                        showInputMessage="1" promptTitle="Enter Value" prompt="Must be positive"
                        showErrorMessage="1" errorStyle="warning" errorTitle="Warning" error="Value should be positive"
                        sqref="D2:D10">
                        <formula1>0</formula1>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml).unwrap();
        assert_eq!(validations.len(), 1);

        let v = &validations[0];
        assert!(v.rule.ignore_blank);

        // Check input message
        let msg = v.rule.input_message.as_ref().unwrap();
        assert!(msg.show);
        assert_eq!(msg.title, "Enter Value");
        assert_eq!(msg.message, "Must be positive");

        // Check error alert
        let alert = v.rule.error_alert.as_ref().unwrap();
        assert!(alert.show);
        assert!(matches!(alert.style, ErrorStyle::Warning));
        assert_eq!(alert.title, "Warning");
        assert_eq!(alert.message, "Value should be positive");
    }

    #[test]
    fn test_parse_validations_multiple_ranges() {
        let xml = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations count="1">
                    <dataValidation type="list" sqref="A1:A10 C1:C10">
                        <formula1>"Yes,No"</formula1>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml).unwrap();
        assert_eq!(validations.len(), 2);
        assert_eq!(validations[0].range, CellRange::new(0, 0, 9, 0));
        assert_eq!(validations[1].range, CellRange::new(0, 2, 9, 2));
    }

    #[test]
    fn test_show_dropdown_inversion() {
        // Excel showDropDown="1" means HIDE dropdown
        // VisiGrid show_dropdown=true means SHOW dropdown
        let xml_hide = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations>
                    <dataValidation type="list" showDropDown="1" sqref="A1">
                        <formula1>"Yes,No"</formula1>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml_hide).unwrap();
        assert!(!validations[0].rule.show_dropdown, "showDropDown='1' should map to show_dropdown=false");

        let xml_show = r#"<?xml version="1.0"?>
            <worksheet>
                <dataValidations>
                    <dataValidation type="list" showDropDown="0" sqref="A1">
                        <formula1>"Yes,No"</formula1>
                    </dataValidation>
                </dataValidations>
            </worksheet>"#;

        let validations = parse_validations_from_xml(xml_show).unwrap();
        assert!(validations[0].rule.show_dropdown, "showDropDown='0' should map to show_dropdown=true");
    }
}
