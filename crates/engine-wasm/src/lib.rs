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

#[cfg(feature = "collab")]
mod collab;
mod formula_help;
// The precedent/dependent inspector serves the collaboration replica only.
#[cfg(feature = "collab")]
mod inspect;
#[cfg(feature = "recipes")]
mod recipes;
mod session;
#[cfg(feature = "collab")]
pub use collab::CollabClient;
pub use session::Session;

#[derive(Deserialize, Clone)]
pub(crate) struct InSheet {
    #[serde(default)]
    pub(crate) name: Option<String>,
    #[serde(default)]
    pub(crate) cells: Vec<InCell>,
}

#[derive(Deserialize, Clone)]
pub(crate) struct InCell {
    pub(crate) row: usize,
    pub(crate) col: usize,
    pub(crate) raw: String,
}

#[derive(Serialize, Debug)]
pub(crate) struct OutResult {
    pub(crate) sheet: usize,
    pub(crate) row: usize,
    pub(crate) col: usize,
    /// Engine result as a JSON-friendly value; null when the cell evaluated
    /// to empty or to an error (see `error`).
    pub(crate) value: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) error: Option<String>,
    /// The formatted display string, for human-facing divergence messages.
    pub(crate) display: String,
}

#[derive(Serialize)]
struct Output {
    engine_version: String,
    engine_commit: String,
    results: Vec<OutResult>,
}

#[wasm_bindgen]
pub fn engine_version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

/// The commit this bundle was built from, suffixed "-modified" if the tree was
/// dirty, and empty if it could not be determined.
///
/// The version alone cannot tell a build made at a release tag from one made
/// afterwards — they report the same number — so anything vendoring this
/// artifact and publishing a version beside it needs this to check the two
/// actually correspond.
#[wasm_bindgen]
pub fn engine_commit() -> String {
    env!("VISIGRID_ENGINE_COMMIT").to_string()
}

/// Bytes of linear memory the engine has grown to (WASM pages never shrink,
/// so this is the high-water mark), for the large-sheet measurements.
#[wasm_bindgen]
pub fn wasm_memory_bytes() -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        (core::arch::wasm32::memory_size(0) * 65536) as f64
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        0.0
    }
}

/// Every built-in function name, sorted — the engine's own FUNCTION_NAMES.
///
/// Callers that state how many functions the engine has (the website says it
/// on several pages) read it from here, so the number follows the engine
/// instead of being retyped each release. It was 96+, 100+, 123 and 132 on
/// different pages at different times, each written by hand.
#[wasm_bindgen]
pub fn function_names() -> Result<JsValue, JsValue> {
    serde_wasm_bindgen::to_value(&visigrid_engine::formula::functions::list_functions())
        .map_err(|e| JsValue::from_str(&e.to_string()))
}

/// Whether the engine can compute each formula: the functions it would not
/// recognise, and any syntax it cannot parse. Takes `["=SUM(A1:A3)", ...]`
/// (formulas with their leading `=`; anything else is a literal and always
/// supported) and returns `[{ supported, unknown_functions, parse_error? }]`
/// in the same order.
///
/// This is the per-workbook gate for making the engine authoritative: a
/// function-name list alone cannot tell `f(3)` inside
/// `LET(f, LAMBDA(x, x*2), f(3))` from an unknown function `F`, and cannot
/// see syntax the parser rejects. The analysis is the engine's own.
#[wasm_bindgen]
pub fn check_formulas(input: JsValue) -> Result<JsValue, JsValue> {
    let formulas: Vec<String> =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&check_formulas_core(&formulas)).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[derive(Serialize, Debug, PartialEq)]
pub(crate) struct FormulaSupport {
    pub(crate) supported: bool,
    pub(crate) unknown_functions: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) parse_error: Option<String>,
}

pub(crate) fn check_formulas_core(formulas: &[String]) -> Vec<FormulaSupport> {
    use visigrid_engine::formula::{analyze, parser};
    formulas
        .iter()
        .map(|raw| {
            if !raw.starts_with('=') {
                return FormulaSupport { supported: true, unknown_functions: Vec::new(), parse_error: None };
            }
            match parser::parse(raw) {
                Err(e) => FormulaSupport { supported: false, unknown_functions: Vec::new(), parse_error: Some(e) },
                Ok(expr) => {
                    let mut counts = std::collections::HashMap::new();
                    analyze::tally_unknown_functions(&expr, &mut counts);
                    let mut unknown: Vec<String> = counts.into_keys().collect();
                    unknown.sort();
                    FormulaSupport { supported: unknown.is_empty(), unknown_functions: unknown, parse_error: None }
                }
            }
        })
        .collect()
}

/// Build a workbook from raw input sheets and recompute it (shared by every
/// export). Cells are written directly onto sheets, so the dependency graph
/// is rebuilt before the ordered recompute (io::json::import_any pattern).
pub(crate) fn build_workbook(sheets: &[InSheet]) -> Workbook {
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

        // Widen the sheet to cover the incoming data before writing it. Cells
        // are stored sparsely, so a write past the recorded dimensions lands
        // either way — but `rows`/`cols` bound whole-column references like
        // SUM(A:A). Left at the 65536x256 default, such a formula would total
        // only part of the data and the shortfall would be reported as an
        // engine divergence, blaming the engine for a setup mistake.
        let (needed_rows, needed_cols) = sheet_in
            .cells
            .iter()
            .fold((0, 0), |(r, c), cell| (r.max(cell.row + 1), c.max(cell.col + 1)));
        sheet.rows = sheet.rows.max(needed_rows);
        sheet.cols = sheet.cols.max(needed_cols);

        for cell in &sheet_in.cells {
            // A leading apostrophe forces text, as it does in every
            // spreadsheet, and is not part of the value.
            //
            // Two reasons. It was already wrong without it: '102 arrived as the
            // four-character string "'102", so any consumer of this API reading
            // a user's typed apostrophe got the marker back as data. And these
            // cells are plain strings with no type channel, so there was no way
            // to express "text that looks like a number" — which is exactly the
            // case the lookup functions now distinguish, and therefore exactly
            // the case a caller needs to be able to construct.
            //
            // Deliberately scoped to this entry point, which carries what a
            // person typed. visigrid-json states types directly and needs no
            // marker; a bare apostrophe in a CSV field is ambiguous and is left
            // alone rather than guessed at.
            if let Some(forced_text) = cell.raw.strip_prefix('\'') {
                sheet.set_text(cell.row, cell.col, forced_text);
            } else {
                // Deferred: the ordered recompute below evaluates each formula
                // once with its dependencies present, and places spills
                // afterwards. Inserting eagerly would evaluate every formula
                // here as well, and spill against a half-built sheet for the
                // recompute to then undo.
                sheet.set_value_deferred(cell.row, cell.col, &cell.raw);
            }
        }
    }

    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

/// Project one evaluated cell into the wire shape.
///
/// Shared by `recompute` and by `Session`, deliberately: the verify chip and a
/// live session must describe the same cell the same way, or a divergence
/// report and a repaint would disagree about what the engine said.
pub(crate) fn out_result(
    sheet_idx: usize,
    sheet: &visigrid_engine::sheet::Sheet,
    row: usize,
    col: usize,
) -> OutResult {
    let (value, error) = match sheet.get_computed_value(row, col) {
        Value::Number(n) => (serde_json::Number::from_f64(n).map(serde_json::Value::Number), None),
        Value::Text(t) => (Some(serde_json::Value::String(t)), None),
        Value::Boolean(b) => (Some(serde_json::Value::Bool(b)), None),
        Value::Error(e) => (None, Some(e)),
        Value::Empty => (None, None),
    };
    OutResult { sheet: sheet_idx, row, col, value, error, display: sheet.get_display(row, col) }
}

#[wasm_bindgen]
pub fn recompute(input: JsValue) -> Result<JsValue, JsValue> {
    console_error_panic_hook::set_once();
    let sheets: Vec<InSheet> =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;

    let output = recompute_core(&sheets);
    serde_wasm_bindgen::to_value(&output).map_err(|e| JsValue::from_str(&e.to_string()))
}

/// The verification itself, free of `JsValue` so its coverage can be asserted
/// in an ordinary test. Every formula cell handed in must come back out: the
/// chip in the web editor reports "engine-verified" against this list, so a
/// cell skipped here is a cell the badge silently vouches for without having
/// checked it.
fn recompute_core(sheets: &[InSheet]) -> Output {
    let wb = build_workbook(sheets);

    let mut results = Vec::new();
    for (i, sheet_in) in sheets.iter().enumerate() {
        let sheet = &wb.sheets()[i];
        for cell in &sheet_in.cells {
            if !cell.raw.starts_with('=') {
                continue;
            }
            results.push(out_result(i, sheet, cell.row, cell.col));
        }
    }

    Output {
        engine_version: engine_version(),
        engine_commit: engine_commit(),
        results,
    }
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

/// JS-facing style delta. A typed struct on purpose: serde_wasm_bindgen
/// serializes structs as plain JS objects but serde_json::Value maps as JS
/// `Map`s, which read as empty from JS property access.
#[derive(Serialize, Default)]
struct CondStyleOut {
    #[serde(skip_serializing_if = "Option::is_none")]
    background_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    bold: Option<bool>,
}

fn hex_of(rgba: [u8; 4]) -> String {
    format!("#{:02X}{:02X}{:02X}", rgba[0], rgba[1], rgba[2])
}

#[derive(Serialize)]
struct CondHit {
    sheet: usize,
    row: usize,
    col: usize,
    /// Engine CellFormatOverride, flattened to JS-friendly values
    style: CondStyleOut,
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
/// CF and validation are evaluated at each provided cell using computed values.
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
            sheet.validations.set(entry.range, entry.rule.clone());
        }
    }

    let mut cond = Vec::new();
    let mut violations = Vec::new();
    for (i, sheet_in) in extras.iter().enumerate() {
        let sheet = &wb.sheets()[i];
        for cell in &sheet_in.cells {
            if let Some(over) = sheet.cond_formats.override_for_cell(cell.row, cell.col, sheet) {
                let style = CondStyleOut {
                    background_color: over.background_color.flatten().map(hex_of),
                    font_color: over.font_color.flatten().map(hex_of),
                    bold: over.bold,
                };
                cond.push(CondHit { sheet: i, row: cell.row, col: cell.col, style });
            }
            if let visigrid_engine::validation::ValidationResult::Invalid { reason, .. } =
                wb.validate_cell(i, cell.row, cell.col)
            {
                violations.push(Violation { sheet: i, row: cell.row, col: cell.col, reason });
            }
        }
    }

    ExtrasOutput { engine_version: engine_version(), cond, violations }
}

// ---------------------------------------------------------------------------
// Sort + filter (Tier-1) — the engine owns the ordering and matching rules
// ---------------------------------------------------------------------------

#[derive(Deserialize)]
struct SortInput {
    cells: Vec<InCell>,
    /// Inclusive data range to sort (header row excluded by the caller).
    start_row: usize,
    end_row: usize,
    /// Column whose values order the rows.
    col: usize,
    #[serde(default)]
    descending: bool,
}

#[derive(Serialize)]
struct SortOutput {
    /// data_row order after sorting, parallel to start_row..=end_row.
    order: Vec<usize>,
}

/// Sort rows by a column using the engine's exact ordering semantics:
/// Numbers < Text < Bool < Error < Blank, normalized comparison within a
/// type, and a STABLE tie-break on current position. Returns the row
/// permutation; the caller moves the data (the web grid has no row view).
#[wasm_bindgen]
pub fn sort_rows(input: JsValue) -> Result<JsValue, JsValue> {
    console_error_panic_hook::set_once();
    let req: SortInput =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;

    let sheets = vec![InSheet { name: None, cells: req.cells }];
    let wb = build_workbook(&sheets);
    let sheet = &wb.sheets()[0];

    let mut keyed: Vec<(visigrid_engine::filter::SortKey, usize)> = (req.start_row..=req.end_row)
        .enumerate()
        .map(|(offset, data_row)| {
            let value = sheet.get_computed_value(data_row, req.col);
            let filter_key = visigrid_engine::filter::FilterKey::from_value(&value);
            (
                visigrid_engine::filter::SortKey::from_filter_key(&filter_key, offset),
                data_row,
            )
        })
        .collect();

    keyed.sort_by(|a, b| a.0.cmp(&b.0));
    if req.descending {
        keyed.reverse();
    }

    let output = SortOutput { order: keyed.into_iter().map(|(_, row)| row).collect() };
    serde_wasm_bindgen::to_value(&output).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[derive(Deserialize)]
struct FilterInput {
    cells: Vec<InCell>,
    start_row: usize,
    end_row: usize,
    col: usize,
    /// Display strings to KEEP (engine-normalized matching).
    keep: Vec<String>,
}

#[derive(Serialize)]
struct FilterOutput {
    /// data_rows to hide.
    hide: Vec<usize>,
    /// Every distinct display value in the column, engine-formatted.
    values: Vec<String>,
}

/// Which rows survive a value filter, using the engine's FilterKey
/// normalization (so "1.0" and "1" match, blanks group, errors group).
#[wasm_bindgen]
pub fn filter_rows(input: JsValue) -> Result<JsValue, JsValue> {
    console_error_panic_hook::set_once();
    let req: FilterInput =
        serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;

    let sheets = vec![InSheet { name: None, cells: req.cells }];
    let wb = build_workbook(&sheets);
    let sheet = &wb.sheets()[0];

    let keep: std::collections::HashSet<String> = req.keep.into_iter().collect();
    let mut hide = Vec::new();
    let mut values: Vec<String> = Vec::new();

    for data_row in req.start_row..=req.end_row {
        let value = sheet.get_computed_value(data_row, req.col);
        let display = visigrid_engine::filter::FilterKey::from_value(&value).display_string();
        if !values.contains(&display) {
            values.push(display.clone());
        }
        if !keep.is_empty() && !keep.contains(&display) {
            hide.push(data_row);
        }
    }

    let output = FilterOutput { hide, values };
    serde_wasm_bindgen::to_value(&output).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    // Host-side test of the core path (wasm-bindgen types work natively too,
    // but we test through the plain structs to keep it toolchain-independent).
    #[test]
    /// A leading apostrophe forces text and is not part of the value.
    ///
    /// This is the only way a caller can express "text that looks like a
    /// number" through this API — the cells are plain strings — and it is the
    /// case the lookup functions distinguish, so it has to be constructible.
    /// It also fixes '102 arriving as a four-character string.
    #[test]
    fn a_leading_apostrophe_forces_text_and_is_dropped() {
        use visigrid_engine::cell::CellValue;

        let sheets = vec![InSheet {
            name: None,
            cells: vec![
                InCell { row: 0, col: 0, raw: "'102".into() },
                InCell { row: 1, col: 0, raw: "102".into() },
                InCell { row: 2, col: 0, raw: "'hello".into() },
            ],
        }];
        let wb = build_workbook(&sheets);
        let sheet = &wb.sheets()[0];

        // Text, and the marker is gone.
        assert!(
            matches!(&sheet.get_cell(0, 0).value, CellValue::Text(t) if t == "102"),
            "'102 should be the text 102, got {:?}",
            sheet.get_cell(0, 0).value
        );
        // Without the marker it is still a number.
        assert!(matches!(sheet.get_cell(1, 0).value, CellValue::Number(n) if n == 102.0));
        // The marker works on ordinary text too, and is still dropped.
        assert!(matches!(&sheet.get_cell(2, 0).value, CellValue::Text(t) if t == "hello"));
    }

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

    // The web editor shows an "engine-verified" shield built from this list.
    // It used to skip every cell past column 256 and row 65536 — Excel 2003's
    // limits, which this engine does not share — so on a wide sheet the shield
    // appeared having checked none of the columns in question. Nothing asserted
    // what the badge actually covered, which is how it survived.
    #[test]
    fn a_formula_past_column_256_is_verified_not_skipped() {
        let sheets = vec![InSheet {
            name: Some("Wide".into()),
            cells: vec![
                InCell { row: 0, col: 0, raw: "21".into() },
                InCell { row: 0, col: 300, raw: "=A1*2".into() },
            ],
        }];

        let out = recompute_core(&sheets);

        let checked = out
            .results
            .iter()
            .find(|r| r.col == 300)
            .expect("a formula at column 300 must be verified, not silently skipped");
        assert_eq!(checked.value, Some(serde_json::json!(42.0)));
    }

    #[test]
    fn a_formula_past_row_65536_is_verified_not_skipped() {
        let sheets = vec![InSheet {
            name: Some("Tall".into()),
            cells: vec![
                InCell { row: 0, col: 0, raw: "21".into() },
                InCell { row: 70_000, col: 0, raw: "=A1*2".into() },
            ],
        }];

        let out = recompute_core(&sheets);

        let checked = out
            .results
            .iter()
            .find(|r| r.row == 70_000)
            .expect("a formula at row 70000 must be verified, not silently skipped");
        assert_eq!(checked.value, Some(serde_json::json!(42.0)));
    }

    // Writes land sparsely whatever the recorded dimensions say, so this is
    // not about storage: `rows`/`cols` bound whole-column references, and a
    // sheet left at the default would evaluate SUM(A:A) over part of its own
    // data and the gap would surface as an engine divergence.
    #[test]
    fn the_sheet_is_widened_to_cover_its_data() {
        let wb = build_workbook(&[InSheet {
            name: None,
            cells: vec![InCell { row: 100_000, col: 300, raw: "1".into() }],
        }]);

        assert!(wb.sheets()[0].cols > 300, "columns must cover the data");
        assert!(wb.sheets()[0].rows > 100_000, "rows must cover the data");
    }

    #[test]
    fn cond_format_and_validation_evaluate() {
        use visigrid_engine::cell::CellFormatOverride;
        use visigrid_engine::cond_format::{CondFormatStore, CondStyle};
        use visigrid_engine::validation::{CellRange, ValidationRule};

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
        assert_eq!(
            out.cond[0].style.background_color.as_deref(),
            Some("#FF0000"),
            "style survives as a JS-friendly hex string"
        );
        assert_eq!(out.violations.len(), 1, "exactly the off-list cell violates");
        assert_eq!((out.violations[0].row, out.violations[0].col), (0, 1));
        assert!(!out.violations[0].reason.is_empty());
    }

    #[test]
    fn validation_extras_check_formula_results_and_cross_sheet_bounds() {
        use visigrid_engine::validation::{CellRange, ValidationRule};
        let out = evaluate_extras_core(vec![
            ExtrasSheet {
                name: Some("Data".into()),
                cells: vec![
                    InCell { row: 0, col: 0, raw: "=2+3".into() },
                    InCell { row: 1, col: 0, raw: "=3+4".into() },
                ],
                cond_formats: None,
                validations: vec![ValidationEntry {
                    range: CellRange::new(0, 0, 1, 0),
                    rule: visigrid_engine::validation::ValidationRule {
                        reference_origin: Some((0, 0)),
                        ..ValidationRule::custom("=A1<Limits!$A$1")
                    },
                }],
            },
            ExtrasSheet {
                name: Some("Limits".into()),
                cells: vec![InCell { row: 0, col: 0, raw: "6".into() }],
                cond_formats: None,
                validations: vec![],
            },
        ]);
        assert_eq!(out.violations.len(), 1);
        assert_eq!((out.violations[0].sheet, out.violations[0].row, out.violations[0].col), (0, 1, 0));
    }

    #[test]
    fn sort_uses_engine_type_ranking() {
        // Numbers before text before blanks, regardless of input order.
        let cells = vec![
            InCell { row: 0, col: 0, raw: "banana".into() },
            InCell { row: 1, col: 0, raw: "10".into() },
            InCell { row: 2, col: 0, raw: "apple".into() },
            InCell { row: 3, col: 0, raw: "2".into() },
        ];
        let sheets = vec![InSheet { name: None, cells: cells.clone() }];
        let wb = build_workbook(&sheets);
        let sheet = &wb.sheets()[0];
        let mut keyed: Vec<(visigrid_engine::filter::SortKey, usize)> = (0..=3)
            .map(|r| {
                let v = sheet.get_computed_value(r, 0);
                let k = visigrid_engine::filter::FilterKey::from_value(&v);
                (visigrid_engine::filter::SortKey::from_filter_key(&k, r), r)
            })
            .collect();
        keyed.sort_by(|a, b| a.0.cmp(&b.0));
        let order: Vec<usize> = keyed.into_iter().map(|(_, r)| r).collect();
        // 2 (row 3), 10 (row 1), apple (row 2), banana (row 0)
        assert_eq!(order, vec![3, 1, 2, 0]);
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

#[cfg(test)]
mod stamp_tests {
    /// The stamp must identify this build, not merely be present.
    #[test]
    fn the_commit_stamp_is_a_real_commit_or_empty() {
        let stamp = super::engine_commit();
        let base = stamp.trim_end_matches("-modified");
        assert!(
            stamp.is_empty() || (base.len() == 40 && base.chars().all(|c| c.is_ascii_hexdigit())),
            "stamp should be a 40-char sha, optionally -modified, or empty; got {stamp:?}"
        );
        println!("STAMP={stamp}");
    }
}

#[cfg(test)]
mod check_formulas_tests {
    use super::*;

    fn check(f: &str) -> FormulaSupport {
        check_formulas_core(&[f.to_string()]).remove(0)
    }

    #[test]
    fn known_functions_let_names_and_constants_are_supported() {
        for f in [
            "=SUM(A1:A3)",
            "=LARGE(A1:A9,2)+MINIFS(B1:B9,C1:C9,\">1\")",
            "=LET(f,LAMBDA(x,x*2),f(3))",
            "=LAMBDA(x,x+1)(2)",
            "=SUM({1,2;3,4})",
            "=TEXTSPLIT(A1,\",\")",
            "=BYROW(A1:A3,LAMBDA(r,SUM(r)))+MAP(A1:A3,LAMBDA(x,x*2))",
            "plain text",
            "42",
        ] {
            assert!(check(f).supported, "{f}: {:?}", check(f));
        }
    }

    #[test]
    fn unknown_functions_are_named() {
        let got = check("=DGET(A1:A3,1,B1:B2)+CUBEVALUE(1,2)");
        assert!(!got.supported);
        assert_eq!(got.unknown_functions, vec!["CUBEVALUE".to_string(), "DGET".to_string()]);
        assert_eq!(got.parse_error, None);
    }

    #[test]
    fn unparseable_formulas_are_unsupported_with_the_reason() {
        let got = check("=SUM(1,");
        assert!(!got.supported);
        assert!(got.parse_error.is_some());
    }
}

