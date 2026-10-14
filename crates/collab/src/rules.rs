//! A sheet's conditional formats and data validation as operation payloads.
//!
//! `SetCondFormats` carries the engine's `CondFormatStore` in its serde form,
//! the same JSON as visigrid-json's `cond_formats`. `SetValidation` carries
//! visigrid-json's list form (`{"validations": [{range, rule}],
//! "exclusions": [range]}`): the store's own serde form is a range-keyed map,
//! which JSON cannot hold. Both are checked here before any replica applies
//! them, so every replica builds the same store or none does.

use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::BTreeSet;
use visigrid_engine::cond_format::CondFormatStore;
use visigrid_engine::sheet::{Sheet, NUM_COLS, NUM_ROWS};
use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore};

/// The most conditional-format rules one sheet may carry.
pub const MAX_COND_RULES: usize = 10_000;
/// The most validation ranges (rules plus exclusions) one sheet may carry,
/// as visigrid-json allows.
pub const MAX_VALIDATION_RANGES: usize = 100_000;

#[derive(Serialize, Deserialize)]
struct Entry {
    range: CellRange,
    rule: ValidationRule,
}

#[derive(Serialize, Deserialize, Default)]
#[serde(deny_unknown_fields)]
struct ValidationList {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    validations: Vec<Entry>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    exclusions: Vec<CellRange>,
}

/// `sheet`'s conditional formats as a `SetCondFormats` payload.
pub fn cond_formats_of(sheet: &Sheet) -> Value {
    serde_json::to_value(&sheet.cond_formats).expect("conditional formats serialize")
}

/// `sheet`'s data validation as a `SetValidation` payload.
pub fn validation_of(sheet: &Sheet) -> Value {
    let list = ValidationList {
        validations: sheet.validations.iter().map(|(range, rule)| Entry { range: *range, rule: rule.clone() }).collect(),
        exclusions: sheet.validations.exclusions_iter().copied().collect(),
    };
    serde_json::to_value(&list).expect("validation serializes")
}

/// The store a `SetCondFormats` payload describes, predicates parsed.
pub fn parse_cond_formats(store: &Value) -> Result<CondFormatStore, String> {
    let mut parsed: CondFormatStore = serde_json::from_value(store.clone()).map_err(|e| format!("invalid conditional formats: {e}"))?;
    if parsed.len() > MAX_COND_RULES {
        return Err(format!("a sheet may carry at most {MAX_COND_RULES} conditional-format rules"));
    }
    for rule in parsed.iter() {
        if rule.ranges.is_empty() {
            return Err("a conditional-format rule needs a range".into());
        }
        for r in &rule.ranges {
            in_grid(r)?;
        }
    }
    parsed.reparse_all();
    Ok(parsed)
}

/// The store a `SetValidation` payload describes.
pub fn parse_validation(rules: &Value) -> Result<ValidationStore, String> {
    let list: ValidationList = serde_json::from_value(rules.clone()).map_err(|e| format!("invalid validation: {e}"))?;
    if list.validations.len().saturating_add(list.exclusions.len()) > MAX_VALIDATION_RANGES {
        return Err(format!("a sheet may carry at most {MAX_VALIDATION_RANGES} validation ranges"));
    }
    let mut store = ValidationStore::new();
    let mut seen = BTreeSet::new();
    for e in &list.validations {
        in_grid(&e.range)?;
        if e.rule.reference_origin.is_some_and(|(r, c)| r >= NUM_ROWS || c >= NUM_COLS) {
            return Err("a validation reference origin is outside the sheet".into());
        }
        if !seen.insert(e.range) {
            return Err("validation names the same range twice".into());
        }
        store.set(e.range, e.rule.clone());
    }
    for range in &list.exclusions {
        in_grid(range)?;
        store.exclude(*range);
    }
    Ok(store)
}

/// A sheet's rules for the rules panel, by stable sheet key:
/// `{cond_formats, validation}` as their op payloads.
pub fn rules_view(wb: &visigrid_engine::workbook::Workbook, sheet: u64) -> Option<Value> {
    let idx = wb.idx_for_sheet_id(visigrid_engine::sheet::SheetId(sheet))?;
    let s = &wb.sheets()[idx];
    Some(serde_json::json!({ "cond_formats": cond_formats_of(s), "validation": validation_of(s) }))
}

/// The validation on one cell, for its dropdown and input message: the rule
/// (as stored, its references resolved for this cell), whether it offers a
/// dropdown, and for a list its items (`items`, `truncated`, `source_error`).
/// `None` when the cell has no rule.
pub fn validation_at(wb: &visigrid_engine::workbook::Workbook, sheet: u64, row: usize, col: usize) -> Option<Value> {
    let idx = wb.idx_for_sheet_id(visigrid_engine::sheet::SheetId(sheet))?;
    let rule = wb.sheets()[idx].validations.get(row, col)?.at(row, col).into_owned();
    let mut out = serde_json::json!({ "rule": rule, "dropdown": wb.has_list_dropdown(idx, row, col) });
    if let Some(list) = wb.get_list_items(idx, row, col) {
        out["items"] = serde_json::json!(list.items);
        out["truncated"] = serde_json::json!(list.is_truncated);
        if let Some(e) = list.source_error {
            out["source_error"] = serde_json::json!(e);
        }
    }
    Some(out)
}

/// Whether typing `input` into a cell passes its validation:
/// `{"valid": true}`, or `{"valid": false, "reason", "alert"}` where `alert`
/// is the rule's error alert (`style` Stop rejects, Warning and Information
/// let the user keep the entry).
pub fn check_input(wb: &visigrid_engine::workbook::Workbook, sheet: u64, row: usize, col: usize, input: &str) -> Value {
    let Some(idx) = wb.idx_for_sheet_id(visigrid_engine::sheet::SheetId(sheet)) else {
        return serde_json::json!({ "valid": true });
    };
    match wb.validate_cell_input(idx, row, col, input) {
        visigrid_engine::validation::ValidationResult::Valid => serde_json::json!({ "valid": true }),
        visigrid_engine::validation::ValidationResult::Invalid { rule, reason } => {
            serde_json::json!({ "valid": false, "reason": reason, "alert": rule.error_alert })
        }
    }
}

fn in_grid(r: &CellRange) -> Result<(), String> {
    if r.start_row > r.end_row || r.start_col > r.end_col || r.end_row >= NUM_ROWS || r.end_col >= NUM_COLS {
        return Err("a rule range is inverted or outside the sheet".into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::cell::CellStyle;
    use visigrid_engine::cond_format::CondStyle;
    use visigrid_engine::sheet::SheetId;

    #[test]
    fn a_sheets_rules_round_trip_through_their_payloads() {
        let mut s = Sheet::new(SheetId(1), 100, 10);
        s.cond_formats.add(vec![CellRange::new(0, 0, 9, 0)], "=A1>5", CondStyle::Named(CellStyle::Error));
        s.validations.set(CellRange::new(0, 1, 9, 1), ValidationRule::list_inline(vec!["Yes".into(), "No".into()]));
        s.validations.exclude(CellRange::new(3, 1, 3, 1));
        let cf = parse_cond_formats(&cond_formats_of(&s)).unwrap();
        assert_eq!(serde_json::to_value(&cf).unwrap(), cond_formats_of(&s));
        assert!(cf.iter().next().unwrap().parse_error().is_none(), "the predicate is parsed");
        let v = parse_validation(&validation_of(&s)).unwrap();
        assert_eq!(v, s.validations);
    }

    #[test]
    fn an_empty_sheets_payloads_are_empty() {
        let s = Sheet::new(SheetId(1), 10, 10);
        assert!(parse_cond_formats(&cond_formats_of(&s)).unwrap().is_empty());
        assert_eq!(validation_of(&s), serde_json::json!({}));
        assert!(parse_validation(&serde_json::json!({})).unwrap().is_empty());
    }

    #[test]
    fn malformed_payloads_are_refused() {
        assert!(parse_cond_formats(&serde_json::json!({"rules": "no"})).is_err());
        let range = serde_json::json!({"start_row": 0, "start_col": 0, "end_row": 0, "end_col": 0});
        let rule = serde_json::to_value(ValidationRule::list_inline(vec!["A".into()])).unwrap();
        let twice = serde_json::json!({"validations": [{"range": range, "rule": rule}, {"range": range, "rule": rule}]});
        assert!(parse_validation(&twice).is_err(), "the same range twice");
        let outside = serde_json::json!({"start_row": 0, "start_col": 0, "end_row": NUM_ROWS, "end_col": 0});
        assert!(parse_validation(&serde_json::json!({"exclusions": [outside]})).is_err(), "outside the sheet");
        assert!(parse_validation(&serde_json::json!({"other": 1})).is_err(), "an unknown field");
    }
}
