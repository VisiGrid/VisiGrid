//! Import recipes in the browser (the custom grid's engine worker): the
//! desktop's recipe runner (`visigrid_io::recipe`) over files the page read,
//! for step-by-step previews, and its output as collaboration operations so
//! an import is one shared, undoable change.
//!
//! Files cross as one byte buffer plus their names and sizes, in order (the
//! first is the source; an Append-folder recipe reads them all).

use std::path::PathBuf;

use serde_json::{json, Value};
use visigrid_collab::op::{CellContent, CollabOp, FormatProps, Rect};
use visigrid_collab::undo::number_format_code;
use visigrid_io::recipe::{run, Landed, Recipe, RecipeOutput, RunResult, Snapshot};
use wasm_bindgen::prelude::*;

/// Rows a preview carries.
const PREVIEW_ROWS: usize = 200;

fn snapshot(names: &[String], sizes: &[u32], data: &[u8]) -> Result<Snapshot, String> {
    if names.is_empty() || names.len() != sizes.len() {
        return Err("files and sizes do not match".into());
    }
    let total: usize = sizes.iter().map(|&n| n as usize).sum();
    if total != data.len() {
        return Err("file sizes do not add up to the data".into());
    }
    let mut at = 0;
    let mut snaps = names.iter().zip(sizes).map(|(name, &n)| {
        let bytes = data[at..at + n as usize].to_vec();
        at += n as usize;
        // The same snapshot the native reader makes (merged recipes, which
        // the browser doesn't read, are left empty)
        Snapshot::from_bytes(&PathBuf::from(name), bytes)
    });
    let mut first = snaps.next().expect("one file");
    first.more = snaps.collect();
    if !first.more.is_empty() {
        first.hash = std::iter::once(first.hash.clone()).chain(first.more.iter().map(|s| s.hash.clone())).collect();
    }
    Ok(first)
}

fn recipe_from(recipe: &Value, upto: Option<usize>) -> Result<Recipe, String> {
    let mut r: Recipe = serde_json::from_value(recipe.clone()).map_err(|e| format!("invalid recipe: {e}"))?;
    if let Some(n) = upto {
        r.steps.truncate(n);
    }
    Ok(r)
}

fn kind_name(output: &RecipeOutput, col: usize) -> &'static str {
    use visigrid_io::recipe::ValueKind;
    match output.columns[col].kind {
        ValueKind::Number => "number",
        ValueKind::DateTime => "datetime",
        ValueKind::Time => "time",
        ValueKind::Plain => match output.columns[col].rule {
            visigrid_io::csv_import::ColumnRule::Number => "number",
            visigrid_io::csv_import::ColumnRule::Date(_) => "date",
            visigrid_io::csv_import::ColumnRule::Text => "text",
            _ => "auto",
        },
    }
}

fn preview(result: &RunResult) -> Value {
    let out = &result.output;
    json!({
        "columns": out.columns.iter().enumerate().map(|(i, c)| json!({"name": c.name, "kind": kind_name(out, i)})).collect::<Vec<_>>(),
        "rows": out.rows.iter().take(PREVIEW_ROWS).collect::<Vec<_>>(),
        "total_rows": out.rows.len(),
        "report": serde_json::to_value(&result.report).unwrap_or(Value::Null),
    })
}

/// Run `recipe` (its first `upto` steps, or all) over the files: the
/// output's columns, its first rows, and the run report (per-step rows in
/// and out, notes, failures).
pub(crate) fn run_preview(recipe: &Value, names: &[String], sizes: &[u32], data: &[u8], upto: Option<usize>) -> Result<Value, String> {
    let r = recipe_from(recipe, upto)?;
    let snap = snapshot(names, sizes, data)?;
    Ok(preview(&run(&r, &snap)))
}

/// The whole output as operations at (`row`, `col`) of `sheet`: the header
/// and values in one ReplaceRange (numbers and dates as typed input, text as
/// text), the header bold, and each date or time column's format. Applied
/// as one envelope, an import is one undo step.
pub(crate) fn import_ops(recipe: &Value, names: &[String], sizes: &[u32], data: &[u8], sheet: u64, row: usize, col: usize) -> Result<Value, String> {
    let r = recipe_from(recipe, None)?;
    let snap = snapshot(names, sizes, data)?;
    let result = run(&r, &snap);
    if !result.report.ok {
        return Err(result.report.failures.join("; "));
    }
    let out = &result.output;
    let width = out.columns.len();
    if width == 0 {
        return Err("the recipe produced no columns".into());
    }
    let mut values = Vec::with_capacity(out.rows.len() + 1);
    values.push(out.columns.iter().map(|c| CellContent::Text(c.name.clone())).collect::<Vec<_>>());
    let mut formats: Vec<Option<String>> = vec![None; width];
    for line in &out.rows {
        let mut cells = Vec::with_capacity(width);
        for c in 0..width {
            let v = line.get(c).map(String::as_str).unwrap_or("");
            let (landed, format) = out.landed(c, v);
            if let Some(f) = format.as_ref().and_then(number_format_code) {
                formats[c] = Some(f);
            }
            cells.push(match landed {
                Landed::Empty => CellContent::Clear,
                Landed::Input(s) => CellContent::Value(s),
                Landed::Text(s) => CellContent::Text(s),
            });
        }
        values.push(cells);
    }
    let rows = values.len();
    let mut ops = vec![
        CollabOp::ReplaceRange { sheet, row, col, values },
        CollabOp::SetFormat { sheet, rect: Rect { r0: row, c0: col, r1: row, c1: col + width - 1 }, props: FormatProps::bold(true) },
    ];
    if rows > 1 {
        for (c, f) in formats.into_iter().enumerate() {
            if let Some(code) = f {
                ops.push(CollabOp::SetFormat {
                    sheet,
                    rect: Rect { r0: row + 1, c0: col + c, r1: row + rows - 1, c1: col + c },
                    props: FormatProps { number_format: Some(Some(code)), ..Default::default() },
                });
            }
        }
    }
    Ok(json!({
        "ops": visigrid_collab::op::ops_to_json(&ops),
        "rows": rows,
        "cols": width,
        "report": serde_json::to_value(&result.report).unwrap_or(Value::Null),
    }))
}

fn to_js(v: &Value) -> Result<JsValue, JsValue> {
    v.serialize(&serde_wasm_bindgen::Serializer::json_compatible()).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn from_js(v: JsValue) -> Result<Value, JsValue> {
    serde_wasm_bindgen::from_value(v).map_err(|e| JsValue::from_str(&e.to_string()))
}

use serde::Serialize as _;

/// Preview a recipe (see `run_preview`). `upto`: how many steps to apply.
#[wasm_bindgen]
pub fn recipe_preview(recipe: JsValue, names: Vec<String>, sizes: Vec<u32>, data: &[u8], upto: Option<usize>) -> Result<JsValue, JsValue> {
    to_js(&run_preview(&from_js(recipe)?, &names, &sizes, data, upto).map_err(|e| JsValue::from_str(&e))?)
}

/// The import as collaboration operations (see `import_ops`).
#[wasm_bindgen]
pub fn recipe_import_ops(recipe: JsValue, names: Vec<String>, sizes: Vec<u32>, data: &[u8], sheet: f64, row: usize, col: usize) -> Result<JsValue, JsValue> {
    to_js(&import_ops(&from_js(recipe)?, &names, &sizes, data, sheet as u64, row, col).map_err(|e| JsValue::from_str(&e))?)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn csv() -> (Vec<String>, Vec<u32>, Vec<u8>) {
        let text = "Region,Amount,Code,When\nWest,10,007,2026-01-05\nEast,5,012,2026-02-01\nWest,7,007,2026-03-09\n";
        (vec!["sales.csv".into()], vec![text.len() as u32], text.as_bytes().to_vec())
    }

    fn recipe() -> Value {
        json!({"version": 1, "source": {"kind": "csv", "path": "sales.csv"},
               "step": [{"op": "types", "columns": {"Code": "text", "When": "date"}},
                        {"op": "group", "by": ["Region"], "totals": [{"fn": "sum", "column": "Amount", "as": "Total"}, {"fn": "first", "column": "Code", "as": "Code"}, {"fn": "first", "column": "When", "as": "When"}]},
                        {"op": "sort", "by": [{"column": "Total", "descending": true}]}]})
    }

    #[test]
    fn previews_each_step_and_imports_typed_values_as_one_change() {
        let (n, s, d) = csv();
        let p = run_preview(&recipe(), &n, &s, &d, Some(0)).unwrap();
        assert_eq!(p["total_rows"], json!(3));
        let p = run_preview(&recipe(), &n, &s, &d, None).unwrap();
        assert_eq!(p["rows"], json!([["West", "17", "007", "2026-01-05"], ["East", "5", "012", "2026-02-01"]]), "{p}");
        assert_eq!(p["report"]["steps"].as_array().unwrap().len(), 3);

        let out = import_ops(&recipe(), &n, &s, &d, 1, 0, 0).unwrap();
        let ops: Vec<CollabOp> = serde_json::from_value(out["ops"].clone()).unwrap();
        let CollabOp::ReplaceRange { values, .. } = &ops[0] else { panic!("{ops:?}") };
        assert_eq!(values[0][0], CellContent::Text("Region".into()));
        assert_eq!(values[1][1], CellContent::Value("17".into()));
        assert_eq!(values[1][2], CellContent::Text("007".into()), "text stays text");
        assert!(matches!(&values[1][3], CellContent::Value(_)), "a date lands as its serial");
        assert!(ops.iter().any(|o| matches!(o, CollabOp::SetFormat { props, .. } if props.number_format == Some(Some("yyyy-mm-dd".into())))));

        // Applied, "007" is text and the date is a date.
        let mut wb = visigrid_engine::workbook::Workbook::new();
        let key = wb.sheets()[0].id.0;
        let ops: Vec<CollabOp> = serde_json::from_value(import_ops(&recipe(), &n, &s, &d, key, 0, 0).unwrap()["ops"].clone()).unwrap();
        visigrid_collab::apply::apply_ops(&mut wb, &ops);
        let sh = &wb.sheets()[0];
        assert_eq!(sh.get_formatted_display(1, 2), "007");
        assert_eq!(sh.get_formatted_display(1, 3), "2026-01-05");
        assert_eq!(sh.get_formatted_display(1, 1), "17");
        assert!(sh.get_format(0, 0).bold);
    }

    #[test]
    fn parquet_reads_from_bytes() {
        let data = include_bytes!("../../cli/tests/fixtures/parquet_orders.parquet");
        let r = json!({"version": 1, "source": {"kind": "parquet", "path": "orders.parquet"}});
        let p = run_preview(&r, &["orders.parquet".into()], &[data.len() as u32], data, None).unwrap();
        assert!(p["report"]["ok"].as_bool().unwrap(), "{p}");
        assert!(p["total_rows"].as_u64().unwrap() > 0 && !p["columns"].as_array().unwrap().is_empty(), "{p}");
    }

    #[test]
    fn a_failed_run_is_an_error() {
        let (n, s, d) = csv();
        let bad = json!({"version": 1, "source": {"kind": "csv", "path": "sales.csv"}, "step": [{"op": "sort", "by": [{"column": "Nope"}]}]});
        assert!(import_ops(&bad, &n, &s, &d, 1, 0, 0).is_err());
    }
}
