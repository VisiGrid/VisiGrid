//! Client-side verification: the real engine, compiled to WASM.
//!
//! The web editor computes live results with Univer's formula engine; truth
//! is the Rust engine. This crate exposes a single `recompute` entry point:
//! feed it raw cells (formulas and literals, exactly as a user typed them),
//! it rebuilds the workbook, recomputes dependency-ordered, and returns the
//! engine's result for every formula cell. The JS side diffs those against
//! Univer's displayed values and surfaces divergences.
//!
//! Input shape (JsValue):  [{ name?, cells: [{ row, col, raw }] }]
//! Output shape (JsValue): { engine_version, results: [{ sheet, row, col,
//!                           value: number|string|bool|null, error?, display }] }
//!
//! Only formula cells produce results — literals need no verification.

use serde::{Deserialize, Serialize};
use visigrid_engine::formula::eval::Value;
use visigrid_engine::workbook::Workbook;
use wasm_bindgen::prelude::*;

#[derive(Deserialize, Clone)]
struct InSheet {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    cells: Vec<InCell>,
}

#[derive(Deserialize, Clone)]
struct InCell {
    row: usize,
    col: usize,
    raw: String,
}

#[derive(Serialize)]
struct OutResult {
    sheet: usize,
    row: usize,
    col: usize,
    /// Engine result as a JSON-friendly value; null when the cell evaluated
    /// to empty or to an error (see `error`).
    value: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
    /// The formatted display string, for human-facing divergence messages.
    display: String,
}

#[derive(Serialize)]
struct Output {
    engine_version: String,
    results: Vec<OutResult>,
}

#[wasm_bindgen]
pub fn engine_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

/// Build a workbook from raw input sheets and recompute it (shared by every
/// export). Cells are written directly onto sheets, so the dependency graph
/// is rebuilt before the ordered recompute (io::json::import_any pattern).
fn build_workbook(sheets: &[InSheet]) -> Workbook {
    let mut wb = Workbook::new();

    // Workbook::new() pre-creates one sheet; grow to match, then name them.
    for i in 1..sheets.len() {
        // Univer enforces unique sheet names, so the named add normally
        // succeeds; fall back to the auto-named sheet on collision/invalid
        // (cross-sheet refs to that name then surface as divergences, which
        // is honest — better than guessing).
        let name = sheets[i].name.clone().unwrap_or_default();
        if name.is_empty() || wb.add_sheet_named(&name).is_none() {
            wb.add_sheet();
        }
    }
    if let Some(first_name) = sheets.first().and_then(|s| s.name.clone()) {
        if !first_name.is_empty() {
            wb.sheets_mut()[0].name = first_name;
        }
    }

    for (i, sheet_in) in sheets.iter().enumerate() {
        let sheet = &mut wb.sheets_mut()[i];
        for cell in &sheet_in.cells {
            if cell.row >= 65536 || cell.col >= 256 {
                continue; // canonical grid bounds; mapper clamps the same way
            }
            sheet.set_value(cell.row, cell.col, &cell.raw);
        }
    }

    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

#[wasm_bindgen]
pub fn recompute(input: JsValue) -> Result<JsValue, JsValue> {
    console_error_panic_hook::set_once();
    let sheets: Vec<InSheet> =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;

    let wb = build_workbook(&sheets);

    let mut results = Vec::new();
    for (i, sheet_in) in sheets.iter().enumerate() {
        let sheet = &wb.sheets()[i];
        for cell in &sheet_in.cells {
            if !cell.raw.starts_with('=') || cell.row >= 65536 || cell.col >= 256 {
                continue;
            }
            let (value, error) = match sheet.get_computed_value(cell.row, cell.col) {
                Value::Number(n) => (serde_json::Number::from_f64(n).map(serde_json::Value::Number), None),
                Value::Text(t) => (Some(serde_json::Value::String(t)), None),
                Value::Boolean(b) => (Some(serde_json::Value::Bool(b)), None),
                Value::Error(e) => (None, Some(e)),
                Value::Empty => (None, None),
            };
            results.push(OutResult {
                sheet: i,
                row: cell.row,
                col: cell.col,
                value,
                error,
                display: sheet.get_display(cell.row, cell.col),
            });
        }
    }

    let output = Output {
        engine_version: engine_version(),
        results,
    };
    serde_wasm_bindgen::to_value(&output).map_err(|e| JsValue::from_str(&e.to_string()))
}

// ---------------------------------------------------------------------------
// Conditional formatting + data validation evaluation (Tier-1, own engine)
// ---------------------------------------------------------------------------

#[derive(Deserialize)]
struct ExtrasSheet {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    cells: Vec<InCell>,
    /// serde form of engine CondFormatStore (rules reparse on load)
    #[serde(default)]
    cond_formats: Option<serde_json::Value>,
    /// JSON-friendly list projection of the ValidationStore: its native
    /// serde form is a CellRange-keyed map, which JSON cannot represent
    /// ("key must be a string") — the vg-json schema must use this list
    /// form too.
    #[serde(default)]
    validations: Vec<ValidationEntry>,
}

#[derive(Deserialize, Clone)]
struct ValidationEntry {
    range: visigrid_engine::validation::CellRange,
    rule: visigrid_engine::validation::ValidationRule,
}

#[derive(Serialize)]
struct CondHit {
    sheet: usize,
    row: usize,
    col: usize,
    /// Engine CellFormatOverride as its serde value (fg/bg/bold/… deltas)
    style: serde_json::Value,
}

#[derive(Serialize)]
struct Violation {
    sheet: usize,
    row: usize,
    col: usize,
    reason: String,
}

#[derive(Serialize)]
struct ExtrasOutput {
    engine_version: String,
    cond: Vec<CondHit>,
    violations: Vec<Violation>,
}

/// Evaluate conditional-formatting rules and data-validation rules through
/// the real engine. Input mirrors `recompute` plus optional per-sheet
/// `cond_formats` / `validations` stores in their engine serde forms (the
/// same shapes visigrid-json will carry once the schema fields land).
/// CF is evaluated at each provided cell; validation at each literal cell.
#[wasm_bindgen]
pub fn evaluate_sheet_extras(input: JsValue) -> Result<JsValue, JsValue> {
    console_error_panic_hook::set_once();
    let extras: Vec<ExtrasSheet> =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;
    let output = evaluate_extras_core(extras);
    serde_wasm_bindgen::to_value(&output).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn evaluate_extras_core(extras: Vec<ExtrasSheet>) -> ExtrasOutput {
    let base: Vec<InSheet> = extras
        .iter()
        .map(|s| InSheet { name: s.name.clone(), cells: s.cells.clone() })
        .collect();
    let mut wb = build_workbook(&base);

    // Install the stores, reparsing CF predicates (parse state is serde-skipped).
    for (i, sheet_in) in extras.iter().enumerate() {
        let sheet = &mut wb.sheets_mut()[i];
        if let Some(cf) = &sheet_in.cond_formats {
            if let Ok(mut store) =
                serde_json::from_value::<visigrid_engine::cond_format::CondFormatStore>(cf.clone())
            {
                store.reparse_all();
                sheet.cond_formats = store;
            }
        }
        for entry in &sheet_in.validations {
            sheet.validations.set(entry.range.clone(), entry.rule.clone());
        }
    }

    let mut cond = Vec::new();
    let mut violations = Vec::new();
    for (i, sheet_in) in extras.iter().enumerate() {
        let sheet = &wb.sheets()[i];
        for cell in &sheet_in.cells {
            if cell.row >= 65536 || cell.col >= 256 {
                continue;
            }
            if let Some(over) = sheet.cond_formats.override_for_cell(cell.row, cell.col, sheet) {
                if let Ok(style) = serde_json::to_value(&over) {
                    cond.push(CondHit { sheet: i, row: cell.row, col: cell.col, style });
                }
            }
            if !cell.raw.starts_with('=') {
                if let visigrid_engine::validation::ValidationResult::Invalid { reason, .. } =
                    sheet.validate_cell_input(cell.row, cell.col, &cell.raw)
                {
                    violations.push(Violation { sheet: i, row: cell.row, col: cell.col, reason });
                }
            }
        }
    }

    ExtrasOutput { engine_version: engine_version(), cond, violations }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Host-side test of the core path (wasm-bindgen types work natively too,
    // but we test through the plain structs to keep it toolchain-independent).
    #[test]
    fn recompute_core_path() {
        let mut wb = Workbook::new();
        wb.sheets_mut()[0].set_value(0, 1, "42");
        wb.sheets_mut()[0].set_value(1, 0, "=B1*2");
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        match wb.sheets()[0].get_computed_value(1, 0) {
            Value::Number(n) => assert_eq!(n, 84.0),
            other => panic!("expected 84, got {:?}", other),
        }
    }

    #[test]
    fn cond_format_and_validation_evaluate() {
        use visigrid_engine::cell::CellFormatOverride;
        use visigrid_engine::cond_format::{CondFormatStore, CondStyle};
        use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore};

        // Build the stores with engine APIs, serialize them — exactly what
        // the web will send once vg-json carries the schema fields.
        let mut cf = CondFormatStore::new();
        cf.add(
            vec![CellRange::new(0, 0, 10, 0)],
            "=A1>10",
            CondStyle::Inline(CellFormatOverride {
                background_color: Some(Some([255, 0, 0, 255])),
                ..Default::default()
            }),
        );

        let validations = vec![ValidationEntry {
            range: CellRange::new(0, 1, 10, 1),
            rule: ValidationRule::list_inline(vec!["Yes".into(), "No".into()]),
        }];

        let extras = vec![ExtrasSheet {
            name: Some("S".into()),
            cells: vec![
                InCell { row: 0, col: 0, raw: "42".into() },   // CF matches (>10)
                InCell { row: 1, col: 0, raw: "3".into() },    // CF no match
                InCell { row: 0, col: 1, raw: "Maybe".into() }, // invalid vs list
                InCell { row: 1, col: 1, raw: "Yes".into() },  // valid
            ],
            cond_formats: Some(serde_json::to_value(&cf).unwrap()),
            validations,
        }];

        let out = evaluate_extras_core(extras);
        assert_eq!(out.cond.len(), 1, "exactly the >10 cell gets the style");
        assert_eq!((out.cond[0].row, out.cond[0].col), (0, 0));
        assert_eq!(out.violations.len(), 1, "exactly the off-list cell violates");
        assert_eq!((out.violations[0].row, out.violations[0].col), (0, 1));
        assert!(out.violations[0].reason.len() > 0);
    }

    #[test]
    fn cross_sheet_formula() {
        let mut wb = Workbook::new();
        assert!(wb.add_sheet_named("Data").is_some());
        wb.sheets_mut()[1].set_value(0, 0, "7");
        wb.sheets_mut()[0].set_value(0, 0, "=Data!A1+1");
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        match wb.sheets()[0].get_computed_value(0, 0) {
            Value::Number(n) => assert_eq!(n, 8.0),
            other => panic!("expected 8, got {:?}", other),
        }
    }
}
