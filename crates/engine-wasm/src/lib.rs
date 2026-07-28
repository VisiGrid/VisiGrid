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

#[derive(Deserialize)]
struct InSheet {
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    cells: Vec<InCell>,
}

#[derive(Deserialize)]
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

#[wasm_bindgen]
pub fn recompute(input: JsValue) -> Result<JsValue, JsValue> {
    let sheets: Vec<InSheet> =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;

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

    // Cells were written directly onto sheets: rebuild the dependency graph
    // before the ordered recompute (same pattern as io::json::import_any).
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

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
