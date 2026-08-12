// JSON export

use std::path::Path;
use std::fs::File;
use std::io::BufWriter;

use visigrid_engine::sheet::Sheet;

/// Export sheet as JSON array of arrays
/// Each row is an array of cell values (strings)
pub fn export(sheet: &Sheet, path: &Path) -> Result<(), String> {
    let file = File::create(path).map_err(|e| e.to_string())?;
    let writer = BufWriter::new(file);

    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut last_non_empty_row = 0;

    for row in 0..sheet.rows {
        let mut record: Vec<String> = Vec::new();
        let mut last_non_empty_col = 0;

        for col in 0..sheet.cols {
            let value = sheet.get_display(row, col);
            if !value.is_empty() {
                last_non_empty_col = col + 1;
                last_non_empty_row = row + 1;
            }
            record.push(value);
        }

        // Trim trailing empty cells
        record.truncate(last_non_empty_col);
        rows.push(record);
    }

    // Trim trailing empty rows
    rows.truncate(last_non_empty_row);

    serde_json::to_writer_pretty(writer, &rows).map_err(|e| e.to_string())?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::tempdir;
    use visigrid_engine::sheet::SheetId;

    #[test]
    fn test_json_export() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.json");

        let mut sheet = Sheet::new(SheetId(1), 100, 10);
        sheet.set_value(0, 0, "Name");
        sheet.set_value(0, 1, "Value");
        sheet.set_value(1, 0, "Alice");
        sheet.set_value(1, 1, "42");
        sheet.set_value(2, 0, "Bob");
        sheet.set_value(2, 1, "17");

        export(&sheet, &path).unwrap();

        let content = fs::read_to_string(&path).unwrap();
        let parsed: Vec<Vec<String>> = serde_json::from_str(&content).unwrap();

        assert_eq!(parsed.len(), 3);
        assert_eq!(parsed[0], vec!["Name", "Value"]);
        assert_eq!(parsed[1], vec!["Alice", "42"]);
        assert_eq!(parsed[2], vec!["Bob", "17"]);
    }
}

// ============================================================================
// visigrid-json v1 — full-fidelity JSON interchange
// ============================================================================
//
// A stable, versioned schema carrying values, formulas, formats, and merges,
// so external tools (the web app, VisiAPI, scripts) can round-trip sheets
// through the engine without parsing xlsx or the native SQLite format.
//
// Contract: fields may be ADDED in later versions; existing fields keep
// their meaning. Consumers must ignore unknown fields. `version` bumps only
// on breaking changes.
//
// Single-sheet form (version 1):
// {
//   "format": "visigrid-json",
//   "version": 1,
//   "name": "Sheet1",
//   "cells": [
//     {"row":0, "col":0, "value":"Item", "fmt":{"bold":true, "bg":"#FFEB3B"}},
//     {"row":1, "col":2, "formula":"=A2*B2", "value":85}
//   ],
//   "merges": [{"start_row":0,"start_col":0,"end_row":0,"end_col":2}],
//   "col_widths": {"0": 120.0},          // added 2026-07-28 (additive)
//   "hidden_rows": [4, 5], "hidden_cols": [2],  // added 2026-08-05 (additive)
//   "row_heights": {"3": 40.0},
//   "frozen_rows": 1,
//   "frozen_cols": 0,
//   "cond_formats": { ...engine CondFormatStore serde form... },   // added 2026-07-29
//   "validations": [ {"range": {...}, "rule": {...}} ],            // list form, NOT a map
//   "filter": {"range": [0,0,99,3], "columns": [{"col":1, "filter": {...}}], "sort": {...}},
//   "charts": [ ...opaque; preserved, not interpreted... ]
// }
//
// Workbook form (version 2) — canonical storage for multi-sheet documents
// (e.g. the web app's R2 blobs). Old consumers reject it loudly via the
// version gate rather than silently reading one sheet:
// {
//   "format": "visigrid-json",
//   "version": 2,
//   "active_sheet": 0,
//   "sheets": [ { ...same per-sheet fields as the v1 body... } ]
// }
//
// Layout fields (col_widths/row_heights/frozen_*) are presentation state —
// the engine does not model them; they travel as a SheetLayout side-car so
// canonical storage never strips layout (the GUI and web mapper own them).
//
// Formula cells carry both the formula and the last computed value, so
// consumers without an engine still see data. On import, formulas are
// recomputed; the stored value is a fallback only.

use std::collections::{BTreeMap, BTreeSet};

use serde::{Deserialize, Serialize};
use visigrid_engine::cell::{Alignment, CellStyle, CellValue, VerticalAlignment};
use visigrid_engine::sheet::MergedRegion;

pub const FULL_JSON_FORMAT: &str = "visigrid-json";
pub const FULL_JSON_VERSION: u32 = 1;
/// Version written for workbook-form (multi-sheet) documents.
pub const FULL_JSON_WORKBOOK_VERSION: u32 = 2;

/// Per-sheet presentation state that lives outside the engine (the GUI and
/// the web mapper own it). BTreeMap for deterministic serialization.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SheetLayout {
    pub col_widths: BTreeMap<usize, f32>,
    pub row_heights: BTreeMap<usize, f32>,
    pub frozen_rows: usize,
    pub frozen_cols: usize,
    /// Rows the user collapsed out of view. BTreeSet for deterministic
    /// serialization, like the maps above.
    pub hidden_rows: BTreeSet<usize>,
    /// Columns the user collapsed out of view.
    ///
    /// Worth carrying even though it looks cosmetic: a hidden column is
    /// usually a scratch column full of intermediates, so losing the flag
    /// doesn't just change the look — it dumps working notes into the middle
    /// of the sheet.
    pub hidden_cols: BTreeSet<usize>,
    /// AutoFilter/sort state (engine-backed on the web side).
    pub filter: Option<FilterSpec>,
    /// Opaque per-sheet charts payload: crates/io doesn't model charts, but
    /// preserves them so a recalc round-trip never strips web-authored charts.
    ///
    /// ONE FIELD IS NOT OPAQUE, and it is a load-bearing cross-repo
    /// convention: each chart entry may carry `data_range`, an **A1 notation
    /// string** ("B2:D10", "Sheet1!A1:A9"). `SheetLayout::shift_for_structural`
    /// rewrites it on row/column edits so a chart never silently plots the
    /// wrong rows. It is produced by the web app's chart UI and consumed here.
    ///
    /// Changing its shape (e.g. to `{sheet, start, end}` for multi-sheet
    /// charts) silently disables that adjustment — charts would keep pointing
    /// at pre-edit rows with no error anywhere. Any such change must update
    /// `shift_a1_range` in this file in the same breath.
    pub charts: Option<serde_json::Value>,
}

impl SheetLayout {
    /// Move presentation state to follow a structural edit on this sheet.
    ///
    /// The engine adjusts cells, formulas, validations, and named ranges;
    /// this side-car holds what the engine doesn't model, and it has to move
    /// too or widths, frozen panes, filters, and charts end up describing
    /// the wrong rows.
    ///
    /// Charts are otherwise opaque here, but by convention each entry may
    /// carry a `data_range` A1 string — the one field worth understanding,
    /// since a chart pointing at shifted rows silently plots the wrong data.
    /// A range wholly consumed by a delete becomes `#REF!`, which the web app
    /// can surface as a broken chart rather than a plausible wrong one.
    pub fn shift_for_structural(&mut self, at: usize, count: usize, delete: bool, is_row: bool) {
        use visigrid_engine::structural::shift_span;

        let shift_keys = |m: &BTreeMap<usize, f32>| -> BTreeMap<usize, f32> {
            m.iter()
                .filter_map(|(k, v)| shift_span(*k, *k, at, count, delete).map(|(nk, _)| (nk, *v)))
                .collect()
        };
        let shift_set = |set: &BTreeSet<usize>| -> BTreeSet<usize> {
            set.iter()
                .filter_map(|k| shift_span(*k, *k, at, count, delete).map(|(nk, _)| nk))
                .collect()
        };
        if is_row {
            self.row_heights = shift_keys(&self.row_heights);
            self.hidden_rows = shift_set(&self.hidden_rows);
            if at < self.frozen_rows {
                self.frozen_rows = if delete {
                    self.frozen_rows.saturating_sub(count.min(self.frozen_rows - at))
                } else {
                    self.frozen_rows + count
                };
            }
        } else {
            self.col_widths = shift_keys(&self.col_widths);
            self.hidden_cols = shift_set(&self.hidden_cols);
            if at < self.frozen_cols {
                self.frozen_cols = if delete {
                    self.frozen_cols.saturating_sub(count.min(self.frozen_cols - at))
                } else {
                    self.frozen_cols + count
                };
            }
        }

        if let Some(f) = &mut self.filter {
            let (s, e) = if is_row { (f.range.0, f.range.2) } else { (f.range.1, f.range.3) };
            match shift_span(s, e, at, count, delete) {
                Some((ns, ne)) => {
                    if is_row {
                        f.range.0 = ns;
                        f.range.2 = ne;
                    } else {
                        f.range.1 = ns;
                        f.range.3 = ne;
                        f.columns.retain_mut(|c| match shift_span(c.col, c.col, at, count, delete) {
                            Some((nc, _)) => {
                                c.col = nc;
                                true
                            }
                            None => false,
                        });
                        if let Some(sort) = &mut f.sort {
                            match shift_span(sort.column, sort.column, at, count, delete) {
                                Some((nc, _)) => sort.column = nc,
                                None => f.sort = None,
                            }
                        }
                    }
                }
                None => self.filter = None, // the filtered region was deleted
            }
        }

        if let Some(charts) = &mut self.charts {
            if let Some(list) = charts.as_array_mut() {
                for chart in list.iter_mut() {
                    let Some(range) = chart.get("data_range").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    let adjusted = shift_a1_range(range, at, count, delete, is_row);
                    if let Some(obj) = chart.as_object_mut() {
                        obj.insert("data_range".into(), serde_json::Value::String(adjusted));
                    }
                }
            }
        }
    }

    pub fn is_empty(&self) -> bool {
        self.hidden_rows.is_empty()
            && self.hidden_cols.is_empty()
            && self.col_widths.is_empty()
            && self.row_heights.is_empty()
            && self.frozen_rows == 0
            && self.frozen_cols == 0
            && self.filter.is_none()
            && self.charts.is_none()
    }
}

fn is_zero(n: &usize) -> bool {
    *n == 0
}

/// One data-validation rule: JSON-friendly list projection of the engine's
/// ValidationStore (whose native serde form is a range-keyed map JSON can't
/// represent). Same shape as engine-wasm's evaluate_sheet_extras input.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationSpec {
    pub range: visigrid_engine::validation::CellRange,
    pub rule: visigrid_engine::validation::ValidationRule,
}

/// AutoFilter/sort state projection (GUI/web presentation state — the
/// engine's FilterState also carries runtime caches, which never serialize).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FilterSpec {
    /// (min_row, min_col, max_row, max_col); header row = min_row.
    pub range: (usize, usize, usize, usize),
    /// List form (not a map): JSON object keys are strings and serde(flatten)
    /// can't round-trip integer-keyed maps.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<ColumnFilterSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort: Option<visigrid_engine::filter::SortState>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ColumnFilterSpec {
    pub col: usize,
    pub filter: visigrid_engine::filter::ColumnFilter,
}

/// Adjust an A1 range string ("B2:D10", "Sheet1!A1:A9") for a structural
/// edit, reusing the engine's formula-reference transform so chart ranges
/// follow exactly the same rules as formula ranges.
fn shift_a1_range(range: &str, at: usize, count: usize, delete: bool, is_row: bool) -> String {
    use visigrid_engine::structural::{adjust_formula_text, Axis, StructuralEdit};

    // The transform works on formulas; a range is one wrapped in '='.
    // Sheet-qualified ranges keep their prefix, and because the edit names
    // the same sheet the reference resolves either way.
    let sheet_name = range
        .split_once('!')
        .map(|(s, _)| s.trim_matches('\'').to_string())
        .unwrap_or_default();
    let edit = StructuralEdit {
        sheet_name: sheet_name.clone(),
        axis: if is_row { Axis::Row } else { Axis::Col },
        at,
        count,
        delete,
    };
    match adjust_formula_text(&format!("={}", range), &edit, &sheet_name) {
        Some(adjusted) => adjusted.trim_start_matches('=').to_string(),
        None => range.to_string(),
    }
}

fn keys_to_string(m: &BTreeMap<usize, f32>) -> BTreeMap<String, f32> {
    m.iter().map(|(k, v)| (k.to_string(), *v)).collect()
}

fn keys_to_usize(m: &BTreeMap<String, f32>) -> BTreeMap<usize, f32> {
    m.iter()
        .filter_map(|(k, v)| k.parse::<usize>().ok().map(|k| (k, *v)))
        .collect()
}

#[derive(Serialize, Deserialize)]
struct FullDoc {
    format: String,
    version: u32,
    /// v1 single-sheet body, flattened at the top level for compatibility.
    #[serde(flatten)]
    body: SheetBody,
    /// v2 workbook form: when non-empty, `body` is unused.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    sheets: Vec<SheetBody>,
    /// Present whenever `sheets` is (v2 workbook form), including when it is
    /// 0 — a consumer should never have to infer whether an absent field means
    /// "sheet 0" or "unknown". Omitted in v1, where it has no meaning.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    active_sheet: Option<usize>,
}

/// The per-sheet payload, shared between the v1 top-level body and each
/// entry of the v2 `sheets` array.
#[derive(Serialize, Deserialize, Default)]
struct SheetBody {
    #[serde(default, skip_serializing_if = "String::is_empty")]
    name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    cells: Vec<FullCell>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    merges: Vec<MergeSpec>,
    // String keys: JSON object keys are strings, and serde(flatten) cannot
    // round-trip integer-keyed maps (it buffers through string-keyed content).
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    col_widths: BTreeMap<String, f32>,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    row_heights: BTreeMap<String, f32>,
    // Sorted index arrays rather than string-keyed maps: there is no value to
    // carry, only membership.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    hidden_rows: Vec<usize>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    hidden_cols: Vec<usize>,
    /// Sheet tab colour as "#RRGGBB".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    tab_color: Option<String>,
    #[serde(default, skip_serializing_if = "is_zero")]
    frozen_rows: usize,
    #[serde(default, skip_serializing_if = "is_zero")]
    frozen_cols: usize,
    /// Engine CondFormatStore in its serde form; predicates reparse on load.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    cond_formats: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    validations: Vec<ValidationSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    filter: Option<FilterSpec>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    charts: Option<serde_json::Value>,
}

#[derive(Serialize, Deserialize)]
struct FullCell {
    row: usize,
    col: usize,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    value: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    fmt: Option<FullFormat>,
    /// Set on cells a formula spilled into, naming the cell it came from.
    ///
    /// These carry a value for readers without an engine, which otherwise see
    /// a one-cell answer where the sheet shows a range. The marker is what
    /// makes that safe to write down: on the way back in these are skipped and
    /// regenerated, because a spill cannot be placed into cells that are
    /// already occupied — writing them as ordinary cells turns every spilling
    /// workbook into #SPILL! on its own round trip.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    spill_from: Option<[usize; 2]>,
}

#[derive(Serialize, Deserialize, Default)]
struct FullFormat {
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    underline: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    strikethrough: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    fg: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    bg: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    valign: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    size: Option<f32>,
    /// Engine NumberFormat as its serde value (VisiGrid-specific; consumers
    /// may pass it through opaquely)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    number_format: Option<serde_json::Value>,
    /// Text overflow behavior: "wrap" | "overflow" (absent = clip, the default)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    overflow: Option<String>,
    /// Per-edge borders (absent edges have no border). Added 2026-07-28 (additive).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    borders: Option<FullBorders>,
}

#[derive(Serialize, Deserialize, Default)]
struct FullBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    top: Option<FullBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    right: Option<FullBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    bottom: Option<FullBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    left: Option<FullBorder>,
}

#[derive(Serialize, Deserialize)]
struct FullBorder {
    /// "thin" | "medium" | "thick"
    style: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    color: Option<String>,
}

fn border_out(b: visigrid_engine::cell::CellBorder) -> Option<FullBorder> {
    use visigrid_engine::cell::BorderStyle;
    let style = match b.style {
        BorderStyle::None => return None,
        BorderStyle::Thin => "thin",
        BorderStyle::Medium => "medium",
        BorderStyle::Thick => "thick",
    };
    Some(FullBorder { style: style.to_string(), color: b.color.map(hex) })
}

fn border_in(b: &Option<FullBorder>) -> visigrid_engine::cell::CellBorder {
    use visigrid_engine::cell::{BorderStyle, CellBorder};
    match b {
        None => CellBorder::default(),
        Some(fb) => CellBorder {
            style: match fb.style.as_str() {
                "medium" => BorderStyle::Medium,
                "thick" => BorderStyle::Thick,
                "thin" => BorderStyle::Thin,
                _ => BorderStyle::None,
            },
            color: fb.color.as_deref().and_then(parse_hex),
        },
    }
}

#[derive(Serialize, Deserialize)]
struct MergeSpec {
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
}

fn hex(rgba: [u8; 4]) -> String {
    format!("#{:02X}{:02X}{:02X}", rgba[0], rgba[1], rgba[2])
}

fn parse_hex(s: &str) -> Option<[u8; 4]> {
    let h = s.trim_start_matches('#');
    if h.len() != 6 || !h.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    Some([
        u8::from_str_radix(&h[0..2], 16).ok()?,
        u8::from_str_radix(&h[2..4], 16).ok()?,
        u8::from_str_radix(&h[4..6], 16).ok()?,
        255,
    ])
}

/// Export a sheet as visigrid-json v1 (no layout side-car).
pub fn export_full(sheet: &Sheet) -> Result<String, String> {
    export_full_with_layout(sheet, &SheetLayout::default())
}

/// Export a sheet as visigrid-json v1 with presentation state.
pub fn export_full_with_layout(sheet: &Sheet, layout: &SheetLayout) -> Result<String, String> {
    let doc = FullDoc {
        format: FULL_JSON_FORMAT.to_string(),
        version: FULL_JSON_VERSION,
        body: sheet_body(sheet, layout),
        sheets: Vec::new(),
        active_sheet: None,
    };
    serde_json::to_string_pretty(&doc).map_err(|e| e.to_string())
}

/// Export a whole workbook as visigrid-json v2 (workbook form).
/// `layouts` is per-sheet, parallel to `wb.sheets()`; missing entries mean
/// no presentation state.
pub fn export_workbook(
    wb: &visigrid_engine::workbook::Workbook,
    layouts: &[SheetLayout],
    active_sheet: usize,
) -> Result<String, String> {
    let default_layout = SheetLayout::default();
    let sheets: Vec<SheetBody> = wb
        .sheets()
        .iter()
        .enumerate()
        .map(|(i, s)| sheet_body(s, layouts.get(i).unwrap_or(&default_layout)))
        .collect();
    let doc = FullDoc {
        format: FULL_JSON_FORMAT.to_string(),
        version: FULL_JSON_WORKBOOK_VERSION,
        body: SheetBody::default(),
        active_sheet: Some(active_sheet.min(sheets.len().saturating_sub(1))),
        sheets,
    };
    serde_json::to_string_pretty(&doc).map_err(|e| e.to_string())
}

fn sheet_body(sheet: &Sheet, layout: &SheetLayout) -> SheetBody {
    let mut cells: Vec<FullCell> = Vec::new();

    let mut coords: Vec<(usize, usize)> = sheet.cells_iter().map(|(&rc, _)| rc).collect();
    coords.sort_unstable();

    for (row, col) in coords {
        let raw = sheet.get_raw(row, col);
        let format = sheet.get_format(row, col);
        let has_format = !format.is_default();
        if raw.is_empty() && !has_format {
            continue;
        }

        let (value, formula) = if raw.starts_with('=') {
            let computed = match sheet.get_computed_value(row, col) {
                visigrid_engine::formula::eval::Value::Number(n) => {
                    serde_json::Number::from_f64(n).map(serde_json::Value::Number)
                }
                visigrid_engine::formula::eval::Value::Text(t) => Some(serde_json::Value::String(t)),
                visigrid_engine::formula::eval::Value::Boolean(b) => Some(serde_json::Value::Bool(b)),
                visigrid_engine::formula::eval::Value::Error(e) => Some(serde_json::Value::String(e)),
                visigrid_engine::formula::eval::Value::Empty => None,
            };
            (computed, Some(raw))
        } else if raw.is_empty() {
            (None, None)
        } else {
            match &sheet.get_cell(row, col).value {
                CellValue::Number(n) => (
                    serde_json::Number::from_f64(*n).map(serde_json::Value::Number),
                    None,
                ),
                _ => (Some(serde_json::Value::String(raw)), None),
            }
        };

        let fmt = if has_format {
            Some(FullFormat {
                bold: format.bold,
                italic: format.italic,
                underline: format.underline,
                strikethrough: format.strikethrough,
                fg: format.font_color.map(hex),
                bg: format.background_color.map(hex),
                style: match format.cell_style {
                    CellStyle::None => None,
                    s => Some(format!("{:?}", s).to_lowercase()),
                },
                align: match format.alignment {
                    Alignment::General => None,
                    Alignment::Left => Some("left".into()),
                    Alignment::Center => Some("center".into()),
                    Alignment::Right => Some("right".into()),
                    Alignment::CenterAcrossSelection => Some("center_across".into()),
                },
                valign: match format.vertical_alignment {
                    VerticalAlignment::Middle => None,
                    VerticalAlignment::Top => Some("top".into()),
                    VerticalAlignment::Bottom => Some("bottom".into()),
                },
                font: format.font_family.clone(),
                size: format.font_size,
                overflow: match format.text_overflow {
                    visigrid_engine::cell::TextOverflow::Clip => None,
                    visigrid_engine::cell::TextOverflow::Wrap => Some("wrap".into()),
                    visigrid_engine::cell::TextOverflow::Overflow => Some("overflow".into()),
                },
                borders: if format.has_any_border() {
                    Some(FullBorders {
                        top: border_out(format.border_top),
                        right: border_out(format.border_right),
                        bottom: border_out(format.border_bottom),
                        left: border_out(format.border_left),
                    })
                } else {
                    None
                },
                number_format: serde_json::to_value(&format.number_format).ok().filter(|v| {
                    // omit the default number format
                    serde_json::to_value(visigrid_engine::cell::NumberFormat::default())
                        .map(|d| *v != d)
                        .unwrap_or(true)
                }),
            })
        } else {
            None
        };

        cells.push(FullCell { row, col, value, formula, fmt, spill_from: None });
    }

    // Cells a formula spilled into. They hold no Cell of their own, so the loop
    // above never sees them.
    for row in 0..sheet.rows {
        for col in 0..sheet.cols {
            if !sheet.is_spill_receiver(row, col) {
                continue;
            }
            let Some(parent) = sheet.get_spill_parent(row, col) else {
                continue;
            };
            use visigrid_engine::formula::eval::Value as EvalValue;
            let value = match sheet.get_spill_value(row, col) {
                Some(EvalValue::Number(n)) => Some(serde_json::json!(n)),
                Some(EvalValue::Text(t)) => Some(serde_json::json!(t)),
                Some(EvalValue::Boolean(b)) => Some(serde_json::json!(b)),
                Some(EvalValue::Error(e)) => Some(serde_json::json!(e)),
                Some(EvalValue::Empty) | None => None,
            };
            cells.push(FullCell {
                row,
                col,
                value,
                formula: None,
                fmt: None,
                spill_from: Some([parent.0, parent.1]),
            });
        }
    }

    let merges = sheet
        .merged_regions
        .iter()
        .map(|m| MergeSpec {
            start_row: m.start.0,
            start_col: m.start.1,
            end_row: m.end.0,
            end_col: m.end.1,
        })
        .collect();

    SheetBody {
        name: sheet.name.clone(),
        tab_color: sheet.tab_color.map(|[r, g, b, _]| format!("#{:02X}{:02X}{:02X}", r, g, b)),
        cells,
        merges,
        col_widths: keys_to_string(&layout.col_widths),
        row_heights: keys_to_string(&layout.row_heights),
        hidden_rows: layout.hidden_rows.iter().copied().collect(),
        hidden_cols: layout.hidden_cols.iter().copied().collect(),
        frozen_rows: layout.frozen_rows,
        frozen_cols: layout.frozen_cols,
        cond_formats: if sheet.cond_formats.is_empty() {
            None
        } else {
            serde_json::to_value(&sheet.cond_formats).ok()
        },
        validations: sheet
            .validations
            .iter()
            .map(|(range, rule)| ValidationSpec { range: range.clone(), rule: rule.clone() })
            .collect(),
        filter: layout.filter.clone(),
        charts: layout.charts.clone(),
    }
}

/// Cheap check: is this a workbook-form (version 2) document?
/// Used by callers that want to preserve the input's form on re-export.
pub fn is_workbook_form(content: &str) -> bool {
    serde_json::from_str::<serde_json::Value>(content)
        .ok()
        .and_then(|v| v.get("sheets").map(|s| s.is_array() && !s.as_array().unwrap().is_empty()))
        .unwrap_or(false)
}

/// Import visigrid-json into a Sheet (formulas recomputed). Accepts both
/// forms; workbook-form documents yield the active sheet.
pub fn import_full(content: &str) -> Result<Sheet, String> {
    import_full_with_layout(content).map(|(sheet, _)| sheet)
}

/// Import visigrid-json into a Sheet plus its presentation side-car.
pub fn import_full_with_layout(content: &str) -> Result<(Sheet, SheetLayout), String> {
    let (wb, mut layouts, active) = import_any(content)?;
    let sheet = wb.sheets()[active].clone();
    let layout = if active < layouts.len() { layouts.swap_remove(active) } else { SheetLayout::default() };
    Ok((sheet, layout))
}

/// Import either form of visigrid-json as a recomputed Workbook plus
/// per-sheet layout side-cars. Single-sheet documents become one-sheet
/// workbooks. Returns (workbook, layouts, active_sheet_index).
pub fn import_any(
    content: &str,
) -> Result<(visigrid_engine::workbook::Workbook, Vec<SheetLayout>, usize), String> {
    use visigrid_engine::sheet::SheetId;
    use visigrid_engine::workbook::Workbook;

    let doc: FullDoc = serde_json::from_str(content).map_err(|e| format!("invalid visigrid-json: {}", e))?;
    if doc.format != FULL_JSON_FORMAT {
        return Err(format!("not a visigrid-json document (format: {:?})", doc.format));
    }
    if doc.version > FULL_JSON_WORKBOOK_VERSION {
        return Err(format!(
            "visigrid-json version {} is newer than supported ({})",
            doc.version, FULL_JSON_WORKBOOK_VERSION
        ));
    }

    let bodies: Vec<&SheetBody> = if doc.sheets.is_empty() {
        vec![&doc.body]
    } else {
        doc.sheets.iter().collect()
    };

    let mut sheets = Vec::with_capacity(bodies.len());
    let mut layouts = Vec::with_capacity(bodies.len());
    for (i, body) in bodies.iter().enumerate() {
        let (sheet, layout) = apply_body(body, SheetId(i as u64 + 1), i)?;
        sheets.push(sheet);
        layouts.push(layout);
    }

    let active = doc.active_sheet.unwrap_or(0).min(sheets.len() - 1);
    // Recompute formulas (stored values are only a fallback for engine-less consumers)
    let mut wb = Workbook::from_sheets(sheets, active);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok((wb, layouts, active))
}

fn apply_body(body: &SheetBody, id: visigrid_engine::sheet::SheetId, index: usize) -> Result<(Sheet, SheetLayout), String> {
    let mut sheet = Sheet::new(id, 65536, 256);
    // Quoted values that read as numbers. Ours are deliberate — the writer only
    // quotes text — but a foreign document may have quoted a number by accident,
    // and it now stays text. Said out loud so that is discoverable rather than
    // something to deduce from a total that came out wrong.
    let mut quoted_numerics = 0usize;
    if !body.name.is_empty() {
        sheet.set_name(&body.name);
        sheet.tab_color = body.tab_color.as_deref().and_then(parse_hex_rgb);
    } else if index > 0 {
        sheet.set_name(&format!("Sheet{}", index + 1));
    }

    for cell in &body.cells {
        // Spill receivers are written for readers without an engine. Loading
        // them would occupy the range the spill needs and turn it into #SPILL!,
        // so they are skipped and the recompute puts them back.
        if cell.spill_from.is_some() {
            continue;
        }

        // Content: formula wins; else typed value
        if let Some(f) = &cell.formula {
            // Deferred: import_any runs an ordered recompute afterwards, which
            // evaluates once with every dependency present and places spills
            // against a finished sheet.
            sheet.set_value_deferred(cell.row, cell.col, f);
        } else if let Some(v) = &cell.value {
            match v {
                // A quoted value is text, and stays text even when it reads
                // like a number. The document already told us the type; this
                // used to flatten it to a string and hand it to set_value,
                // which inferred the type over again — so "007" came back as
                // 7 with nothing to say a zip code had become an integer.
                //
                // Writers only quote what was text, so our own round trips are
                // exact. A foreign document quoting a number gets text, which
                // is visible and fixable; the previous behaviour was neither.
                serde_json::Value::String(s) => {
                    if !s.is_empty() && s.parse::<f64>().is_ok() {
                        quoted_numerics += 1;
                    }
                    sheet.set_text(cell.row, cell.col, s);
                }
                serde_json::Value::Number(n) => {
                    sheet.set_value_deferred(cell.row, cell.col, &n.to_string());
                }
                serde_json::Value::Bool(b) => {
                    sheet.set_value_deferred(cell.row, cell.col, if *b { "TRUE" } else { "FALSE" });
                }
                other => {
                    sheet.set_value_deferred(cell.row, cell.col, &other.to_string());
                }
            }
        }

        if let Some(f) = &cell.fmt {
            let mut format = sheet.get_format(cell.row, cell.col);
            format.bold = f.bold;
            format.italic = f.italic;
            format.underline = f.underline;
            format.strikethrough = f.strikethrough;
            format.font_color = f.fg.as_deref().and_then(parse_hex);
            format.background_color = f.bg.as_deref().and_then(parse_hex);
            if let Some(style) = &f.style {
                format.cell_style = match style.as_str() {
                    "error" => CellStyle::Error,
                    "warning" => CellStyle::Warning,
                    "success" => CellStyle::Success,
                    "input" => CellStyle::Input,
                    "total" => CellStyle::Total,
                    "note" => CellStyle::Note,
                    _ => CellStyle::None,
                };
            }
            if let Some(a) = &f.align {
                format.alignment = match a.as_str() {
                    "left" => Alignment::Left,
                    "center" => Alignment::Center,
                    "right" => Alignment::Right,
                    "center_across" => Alignment::CenterAcrossSelection,
                    _ => Alignment::General,
                };
            }
            if let Some(v) = &f.valign {
                format.vertical_alignment = match v.as_str() {
                    "top" => VerticalAlignment::Top,
                    "bottom" => VerticalAlignment::Bottom,
                    _ => VerticalAlignment::Middle,
                };
            }
            format.font_family = f.font.clone();
            format.font_size = f.size;
            if let Some(o) = &f.overflow {
                format.text_overflow = match o.as_str() {
                    "wrap" => visigrid_engine::cell::TextOverflow::Wrap,
                    "overflow" => visigrid_engine::cell::TextOverflow::Overflow,
                    _ => visigrid_engine::cell::TextOverflow::Clip,
                };
            }
            if let Some(b) = &f.borders {
                format.border_top = border_in(&b.top);
                format.border_right = border_in(&b.right);
                format.border_bottom = border_in(&b.bottom);
                format.border_left = border_in(&b.left);
            }
            if let Some(nf) = &f.number_format {
                if let Ok(parsed) = serde_json::from_value(nf.clone()) {
                    format.number_format = parsed;
                }
            }
            sheet.set_format(cell.row, cell.col, format);
        }
    }

    for m in &body.merges {
        let _ = sheet.add_merge(MergedRegion::new(m.start_row, m.start_col, m.end_row, m.end_col));
    }

    if let Some(cf) = &body.cond_formats {
        // Loud on malformed stores: silently dropping rules would be data
        // loss. A structurally alien store fails the whole import.
        let mut store = serde_json::from_value::<visigrid_engine::cond_format::CondFormatStore>(cf.clone())
            .map_err(|e| format!("sheet {:?}: invalid cond_formats: {}", body.name, e))?;
        store.reparse_all();
        sheet.cond_formats = store;
    }
    for v in &body.validations {
        sheet.validations.set(v.range.clone(), v.rule.clone());
    }

    let layout = SheetLayout {
        col_widths: keys_to_usize(&body.col_widths),
        row_heights: keys_to_usize(&body.row_heights),
        hidden_rows: body.hidden_rows.iter().copied().collect(),
        hidden_cols: body.hidden_cols.iter().copied().collect(),
        frozen_rows: body.frozen_rows,
        frozen_cols: body.frozen_cols,
        filter: body.filter.clone(),
        charts: body.charts.clone(),
    };

    if quoted_numerics > 0 {
        eprintln!(
            "[JSON import] {} quoted value(s) in {:?} read as numbers and were kept as text \
             (a quoted value is text in visigrid-json)",
            quoted_numerics,
            sheet.name
        );
    }

    Ok((sheet, layout))
}

#[cfg(test)]
mod full_json_tests {
    use super::*;
    use visigrid_engine::sheet::SheetId;

    /// Render a sheet's top-left corner as the user would see it.
    fn render(sheet: &Sheet, rows: usize, cols: usize) -> String {
        (0..rows)
            .map(|r| {
                (0..cols)
                    .map(|c| sheet.get_display(r, c))
                    .collect::<Vec<_>>()
                    .join("|")
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn import_cells(cells: &[&str]) -> Sheet {
        let doc = format!(
            r#"{{"format":"visigrid-json","version":2,"active_sheet":0,"sheets":[{{"name":"S","cells":[{}]}}]}}"#,
            cells.join(",")
        );
        import_full(&doc).unwrap()
    }

    /// A reader without an engine sees the whole spilled range, and the file
    /// still survives being read back.
    ///
    /// These two pull against each other, which is the only interesting thing
    /// here. Writing the spilled cells down is what the format promises — its
    /// stored values exist for consumers that cannot recompute. But a spill
    /// cannot be placed into cells that are already occupied, so writing them
    /// as ordinary cells makes the document refuse its own spill with #SPILL!
    /// the moment it is reopened. The marker is what lets both be true.
    #[test]
    fn spilled_values_are_written_down_and_do_not_block_the_reload() {
        let mut sheet = Sheet::new(SheetId(1), 100, 10);
        sheet.set_value(0, 0, "=SEQUENCE(3,1,10,5)");

        let exported = export_full(&sheet).unwrap();
        let doc: serde_json::Value = serde_json::from_str(&exported).unwrap();
        let cells = doc["cells"].as_array().unwrap();

        let spilled: Vec<_> = cells.iter().filter(|c| c.get("spill_from").is_some()).collect();
        assert_eq!(spilled.len(), 2, "both cells the array spilled into should be written");
        assert_eq!(spilled[0]["value"], 15.0);
        assert_eq!(spilled[1]["value"], 20.0);
        assert_eq!(spilled[0]["spill_from"], serde_json::json!([0, 0]));

        // Reading it back must reproduce the sheet, not a refusal.
        let restored = import_full(&exported).unwrap();
        assert_eq!(
            (restored.get_display(0, 0), restored.get_display(1, 0), restored.get_display(2, 0)),
            ("10".to_string(), "15".to_string(), "20".to_string()),
            "reopening the file must not turn its own spill into #SPILL!"
        );

        // And again, since the second write is the one that would carry any
        // damage forward.
        let twice = import_full(&export_full(&restored).unwrap()).unwrap();
        assert_eq!(twice.get_display(2, 0), "20");
    }

    /// A document means the same thing whichever order its cells are listed in.
    ///
    /// Spills are applied as each cell is inserted, so whether a spill lands
    /// before or after the cell it would overwrite decides the outcome — and
    /// which one a user gets depends on the order their writer happened to
    /// emit. Listed one way a spill silently destroys an occupied cell; listed
    /// the other it correctly refuses with #SPILL! and the value survives.
    ///
    /// Deliberately not a characterization test. Pinning today's output would
    /// pin the bug; order independence is checkable without knowing which of
    /// the two answers is right.
    #[test]
    fn spilling_does_not_depend_on_the_order_cells_are_listed() {
        let corpus: Vec<(&str, Vec<&str>)> = vec![
            (
                "spill onto an occupied cell",
                vec![
                    r#"{"row":0,"col":0,"formula":"=SEQUENCE(3,1,10,5)"}"#,
                    r#"{"row":1,"col":0,"value":"keep me"}"#,
                ],
            ),
            (
                "spill reading cells declared later",
                vec![
                    r#"{"row":0,"col":2,"formula":"=SORT(A1:A3)"}"#,
                    r#"{"row":0,"col":0,"value":30}"#,
                    r#"{"row":1,"col":0,"value":10}"#,
                    r#"{"row":2,"col":0,"value":20}"#,
                ],
            ),
            (
                "two spills whose ranges overlap",
                vec![
                    r#"{"row":0,"col":0,"formula":"=SEQUENCE(3,1,1,1)"}"#,
                    r#"{"row":1,"col":0,"formula":"=SEQUENCE(3,1,100,1)"}"#,
                ],
            ),
            (
                "a spill feeding a non-array formula",
                vec![
                    r#"{"row":0,"col":0,"formula":"=SEQUENCE(3,1,10,5)"}"#,
                    r#"{"row":0,"col":2,"formula":"=SUM(A1:A3)"}"#,
                ],
            ),
        ];

        for (label, cells) in corpus {
            let forward = import_cells(&cells);
            let reversed: Vec<&str> = cells.iter().rev().copied().collect();
            let backward = import_cells(&reversed);

            assert_eq!(
                render(&forward, 4, 4),
                render(&backward, 4, 4),
                "{label}: the same document read in a different cell order produced \
                 different results"
            );
        }

        // Agreement is not correctness: two runs agreeing on a wrong answer
        // satisfies the property completely. An earlier version of this fix
        // passed everything above while silently dropping the #SPILL! — both
        // orders reported the array's first value as though nothing had gone
        // wrong. So at least one case states outright what the answer should be.
        let blocked = import_cells(&[
            r#"{"row":0,"col":0,"formula":"=SEQUENCE(3,1,10,5)"}"#,
            r#"{"row":1,"col":0,"value":"keep me"}"#,
        ]);
        assert_eq!(
            blocked.get_display(0, 0),
            "#SPILL!",
            "a spill that cannot fit must say so"
        );
        assert_eq!(
            blocked.get_display(1, 0),
            "keep me",
            "and must not overwrite what was in its way"
        );

        // The unobstructed case must still spill, or the assertion above could
        // be satisfied by never spilling at all.
        let clear = import_cells(&[r#"{"row":0,"col":0,"formula":"=SEQUENCE(3,1,10,5)"}"#]);
        assert_eq!(
            (clear.get_display(0, 0), clear.get_display(1, 0), clear.get_display(2, 0)),
            ("10".to_string(), "15".to_string(), "20".to_string()),
            "an unobstructed spill must still land"
        );
    }

    // A text cell and a numeric cell holding the same digits are different
    // rows, and every lookup function agrees which is which.
    //
    // This is the other half of the strict rule. The key's type was respected
    // first, which made a text key stop matching a numeric cell — but cells
    // were still read as text and re-typed by parsing, so a *numeric* key
    // happily matched a text cell. The rule only held for text written into a
    // formula, which is not a rule anyone could state.
    //
    // The test lives here rather than in the engine because a genuinely
    // text-typed numeric cell can only arrive through a file: writing one into
    // a cell gets it re-read as a number.
    #[test]
    fn a_text_cell_and_a_numeric_cell_are_told_apart_by_every_lookup() {
        let doc = r#"{"format":"visigrid-json","version":2,"active_sheet":0,"sheets":[{"name":"S","cells":[
            {"row":0,"col":0,"value":"102"},{"row":0,"col":1,"value":"text-row"},
            {"row":1,"col":0,"value":102},{"row":1,"col":1,"value":"number-row"},
            {"row":3,"col":3,"formula":"=VLOOKUP(102,A1:B2,2,FALSE)"},
            {"row":4,"col":3,"formula":"=VLOOKUP(\"102\",A1:B2,2,FALSE)"},
            {"row":5,"col":3,"formula":"=MATCH(102,A1:A2,0)"},
            {"row":6,"col":3,"formula":"=MATCH(\"102\",A1:A2,0)"},
            {"row":7,"col":3,"formula":"=XLOOKUP(102,A1:A2,B1:B2,\"MISS\")"},
            {"row":8,"col":3,"formula":"=XLOOKUP(\"102\",A1:A2,B1:B2,\"MISS\")"}
        ]}]}"#;
        let sheet = import_full(doc).unwrap();

        // A numeric key skips the text row and finds the number.
        assert_eq!(sheet.get_display(3, 3), "number-row");
        assert_eq!(sheet.get_display(5, 3), "2");
        assert_eq!(sheet.get_display(7, 3), "number-row");

        // A text key does the opposite. Both directions, not just the one.
        assert_eq!(sheet.get_display(4, 3), "text-row");
        assert_eq!(sheet.get_display(6, 3), "1");
        assert_eq!(sheet.get_display(8, 3), "text-row");
    }

    // A quoted value is text, and survives being read back as text.
    //
    // The reader used to match on the JSON type and then flatten it to a string
    // for set_value, which inferred the type over again — so a zip code written
    // as "007" returned as the number 7. Every server-side recalc goes through
    // this path, which is why the xlsx reader's fix was undone the first time
    // anyone saved.
    #[test]
    fn quoted_values_stay_text_through_a_round_trip() {
        use visigrid_engine::cell::CellValue;

        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_text(0, 0, "007");
        sheet.set_text(1, 0, "0123456789");
        sheet.set_value(2, 0, "42"); // a real number, must stay one
        sheet.set_text(3, 0, "label"); // ordinary text, unaffected

        let restored = import_full(&export_full(&sheet).unwrap()).unwrap();

        let text_at = |row: usize, expected: &str| {
            let value = &restored.get_cell(row, 0).value;
            assert!(
                matches!(value, CellValue::Text(s) if s == expected),
                "row {row} should still be the text {expected:?}, got {value:?}"
            );
        };
        text_at(0, "007");
        text_at(1, "0123456789");
        text_at(3, "label");

        let number = &restored.get_cell(2, 0).value;
        assert!(
            matches!(number, CellValue::Number(n) if *n == 42.0),
            "a real number must stay a number, got {number:?}"
        );
    }

    #[test]
    fn full_json_roundtrip() {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_name("Model");
        sheet.set_value(0, 0, "Revenue");
        sheet.set_value(1, 0, "100");
        sheet.set_value(1, 1, "=A2*2");
        let mut f = sheet.get_format(0, 0);
        f.bold = true;
        f.background_color = Some([255, 235, 59, 255]);
        sheet.set_format(0, 0, f);
        let _ = sheet.add_merge(MergedRegion::new(3, 0, 3, 2));

        let json = export_full(&sheet).unwrap();
        assert!(json.contains("\"visigrid-json\""));
        assert!(json.contains("=A2*2"));
        assert!(json.contains("#FFEB3B"));

        let restored = import_full(&json).unwrap();
        assert_eq!(restored.name, "Model");
        assert_eq!(restored.get_raw(0, 0), "Revenue");
        assert_eq!(restored.get_raw(1, 1), "=A2*2");
        assert_eq!(restored.get_display(1, 1), "200", "formula recomputed");
        assert!(restored.get_format(0, 0).bold);
        assert_eq!(restored.get_format(0, 0).background_color, Some([255, 235, 59, 255]));
        assert_eq!(restored.merged_regions.len(), 1);
    }

    #[test]
    fn layout_side_car_roundtrip() {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_value(0, 0, "x");
        let mut layout = SheetLayout::default();
        layout.col_widths.insert(0, 120.0);
        layout.row_heights.insert(3, 40.0);
        layout.frozen_rows = 1;

        let json = export_full_with_layout(&sheet, &layout).unwrap();
        assert!(json.contains("\"col_widths\""));
        let (_, restored) = import_full_with_layout(&json).unwrap();
        assert_eq!(restored, layout);

        // Layout-less docs (all pre-2026-07-28 blobs) parse with empty layout
        let (_, empty) = import_full_with_layout(&export_full(&sheet).unwrap()).unwrap();
        assert!(empty.is_empty());
    }

    #[test]
    fn workbook_form_roundtrip_with_cross_sheet_formula() {
        use visigrid_engine::workbook::Workbook;

        let mut a = Sheet::new(SheetId(1), 100, 100);
        a.set_name("Data");
        a.set_value(0, 0, "21");
        let mut b = Sheet::new(SheetId(2), 100, 100);
        b.set_name("Summary");
        b.set_value(0, 0, "=Data!A1*2");

        let mut wb = Workbook::from_sheets(vec![a, b], 1);
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        let mut layout_b = SheetLayout::default();
        layout_b.col_widths.insert(0, 200.0);
        let json = export_workbook(&wb, &[SheetLayout::default(), layout_b.clone()], 1).unwrap();
        assert!(json.contains("\"version\": 2"));
        assert!(json.contains("\"sheets\""));

        let (restored, layouts, active) = import_any(&json).unwrap();
        assert_eq!(restored.sheets().len(), 2);
        assert_eq!(active, 1);
        assert_eq!(restored.sheets()[0].name, "Data");
        assert_eq!(restored.sheets()[1].get_display(0, 0), "42", "cross-sheet formula recomputed");
        assert_eq!(layouts[1], layout_b);

        // import_full on a workbook doc yields the active sheet
        let active_sheet = import_full(&json).unwrap();
        assert_eq!(active_sheet.name, "Summary");
    }

    #[test]
    fn borders_and_wrap_roundtrip() {
        use visigrid_engine::cell::{BorderStyle, CellBorder, TextOverflow};

        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_value(0, 0, "boxed");
        let mut f = sheet.get_format(0, 0);
        f.text_overflow = TextOverflow::Wrap;
        f.border_top = CellBorder { style: BorderStyle::Thick, color: Some([255, 0, 0, 255]) };
        f.border_bottom = CellBorder { style: BorderStyle::Thin, color: None };
        sheet.set_format(0, 0, f);

        let json = export_full(&sheet).unwrap();
        assert!(json.contains("\"wrap\"") && json.contains("\"thick\""));

        let restored = import_full(&json).unwrap();
        let rf = restored.get_format(0, 0);
        assert_eq!(rf.text_overflow, TextOverflow::Wrap);
        assert_eq!(rf.border_top.style, BorderStyle::Thick);
        assert_eq!(rf.border_top.color, Some([255, 0, 0, 255]));
        assert_eq!(rf.border_bottom.style, BorderStyle::Thin);
        assert_eq!(rf.border_left.style, BorderStyle::None);
    }

    #[test]
    fn layout_follows_structural_edits() {
        use visigrid_engine::filter::{ColumnFilter, SortDirection, SortState};

        let mut layout = SheetLayout {
            col_widths: [(0, 100.0), (3, 150.0)].into_iter().collect(),
            row_heights: [(1, 30.0), (5, 40.0)].into_iter().collect(),
            hidden_rows: Default::default(),
            hidden_cols: Default::default(),
            frozen_rows: 2,
            frozen_cols: 1,
            filter: Some(FilterSpec {
                range: (4, 0, 20, 3),
                columns: vec![ColumnFilterSpec { col: 2, filter: ColumnFilter::default() }],
                sort: Some(SortState { column: 2, direction: SortDirection::Ascending }),
            }),
            charts: Some(serde_json::json!([
                {"id": "c1", "chart_type": "bar", "data_range": "A5:B20"},
                {"id": "c2", "chart_type": "pie", "data_range": "Sheet1!D1:D4"},
            ])),
        };

        // Insert 2 rows at row index 1 (inside the frozen band, above everything).
        layout.shift_for_structural(1, 2, false, true);
        assert_eq!(layout.row_heights.keys().copied().collect::<Vec<_>>(), vec![3, 7]);
        assert_eq!(layout.frozen_rows, 4, "frozen band grew with the insert");
        assert_eq!(layout.col_widths.keys().copied().collect::<Vec<_>>(), vec![0, 3], "columns untouched");
        let f = layout.filter.as_ref().unwrap();
        assert_eq!((f.range.0, f.range.2), (6, 22), "filter range moved down");
        // Chart ranges follow — a chart pointing at shifted rows would
        // otherwise plot the wrong data with no visible error.
        let charts = layout.charts.as_ref().unwrap().as_array().unwrap();
        assert_eq!(charts[0]["data_range"], "A7:B22", "range below the insert shifts");
        // D1:D4 spans rows 0-3, so an insert at row index 1 lands INSIDE it:
        // grid-line semantics expand rather than shift, matching Excel and
        // the formula-range rule.
        assert_eq!(charts[1]["data_range"], "Sheet1!D1:D6", "range containing the insert expands");

        // Deleting the columns a filter/sort targets drops them cleanly.
        let mut l2 = SheetLayout {
            filter: Some(FilterSpec {
                range: (0, 0, 9, 5),
                columns: vec![ColumnFilterSpec { col: 2, filter: ColumnFilter::default() }],
                sort: Some(SortState { column: 2, direction: SortDirection::Ascending }),
            }),
            ..SheetLayout::default()
        };
        l2.shift_for_structural(2, 1, true, false);
        let f2 = l2.filter.as_ref().unwrap();
        assert!(f2.columns.is_empty(), "filter on a deleted column is dropped");
        assert!(f2.sort.is_none(), "sort on a deleted column is dropped");

        // A chart whose whole range is deleted becomes #REF!, not a plausible
        // wrong range.
        let mut l3 = SheetLayout {
            charts: Some(serde_json::json!([{"id": "c", "data_range": "A2:A4"}])),
            ..SheetLayout::default()
        };
        l3.shift_for_structural(0, 10, true, true);
        assert_eq!(l3.charts.unwrap().as_array().unwrap()[0]["data_range"], "#REF!");
    }

    #[test]
    fn tier1_extras_roundtrip() {
        use visigrid_engine::cell::CellStyle;
        use visigrid_engine::cond_format::CondStyle;
        use visigrid_engine::filter::{ColumnFilter, SortDirection, SortState};
        use visigrid_engine::validation::{
            CellRange, ListSource, ValidationResult, ValidationRule, ValidationType,
        };
        use visigrid_engine::workbook::Workbook;

        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_name("Data");
        sheet.set_value(0, 0, "150");
        // CF: values > 100 get the error style
        sheet.cond_formats.add(
            vec![CellRange { start_row: 0, start_col: 0, end_row: 9, end_col: 0 }],
            "=A1>100",
            CondStyle::Named(CellStyle::Error),
        );
        // Validation: B column restricted to a list
        sheet.validations.set(
            CellRange { start_row: 0, start_col: 1, end_row: 9, end_col: 1 },
            ValidationRule::new(ValidationType::List(ListSource::Inline(vec![
                "yes".into(),
                "no".into(),
            ]))),
        );
        let mut wb = Workbook::from_sheets(vec![sheet], 0);
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        let layout = SheetLayout {
            filter: Some(FilterSpec {
                range: (0, 0, 9, 3),
                columns: vec![ColumnFilterSpec { col: 1, filter: ColumnFilter::default() }],
                sort: Some(SortState { column: 0, direction: SortDirection::Descending }),
            }),
            charts: Some(serde_json::json!([{"kind": "bar", "web_only": true}])),
            ..SheetLayout::default()
        };

        let json = export_workbook(&wb, &[layout.clone()], 0).unwrap();
        assert!(json.contains("cond_formats") && json.contains("validations"));
        assert!(json.contains("\"filter\"") && json.contains("web_only"));

        let (restored, layouts, _) = import_any(&json).unwrap();
        let rsheet = &restored.sheets()[0];
        // CF survived AND predicates reparsed (rule actually evaluates)
        assert!(rsheet.cond_formats.override_for_cell(0, 0, rsheet).is_some(),
            "reparsed CF rule must match A1=150");
        assert!(rsheet.cond_formats.override_for_cell(1, 0, rsheet).is_none());
        // Validation survived and enforces
        assert!(matches!(
            rsheet.validate_cell_input(0, 1, "maybe"),
            ValidationResult::Invalid { .. }
        ));
        // Filter + charts side-car round-tripped exactly
        assert_eq!(layouts[0].filter, layout.filter);
        assert_eq!(layouts[0].charts, layout.charts);
    }

    #[test]
    fn malformed_cond_formats_fail_loudly() {
        let doc = r#"{"format":"visigrid-json","version":1,"cond_formats":{"rules":"not-a-list"}}"#;
        assert!(import_full(doc).is_err());
    }

    #[test]
    fn rejects_versions_beyond_workbook() {
        assert!(import_full("{\"format\":\"visigrid-json\",\"version\":3}").is_err());
    }

    #[test]
    fn rejects_foreign_documents() {
        assert!(import_full("[[1,2],[3,4]]").is_err());
        assert!(import_full("{\"format\":\"other\",\"version\":1}").is_err());
        assert!(import_full("{\"format\":\"visigrid-json\",\"version\":99}").is_err());
    }
}

#[cfg(test)]
mod hidden_layout_tests {
    use super::*;

    fn layout_with_hidden() -> SheetLayout {
        SheetLayout {
            hidden_rows: [2, 5].into_iter().collect(),
            hidden_cols: [1].into_iter().collect(),
            ..Default::default()
        }
    }

    /// Hidden indices are positions, so a structural edit has to move them —
    /// otherwise inserting a row above a hidden one hides the wrong row.
    #[test]
    fn hidden_rows_shift_on_insert() {
        let mut l = layout_with_hidden();
        l.shift_for_structural(0, 1, false, true);
        assert_eq!(l.hidden_rows.iter().copied().collect::<Vec<_>>(), vec![3, 6]);
        assert_eq!(l.hidden_cols.iter().copied().collect::<Vec<_>>(), vec![1], "columns untouched");
    }

    #[test]
    fn deleting_a_hidden_row_drops_it() {
        let mut l = layout_with_hidden();
        l.shift_for_structural(2, 1, true, true);
        assert_eq!(l.hidden_rows.iter().copied().collect::<Vec<_>>(), vec![4]);
    }

    #[test]
    fn hidden_cols_shift_on_column_insert() {
        let mut l = layout_with_hidden();
        l.shift_for_structural(0, 2, false, false);
        assert_eq!(l.hidden_cols.iter().copied().collect::<Vec<_>>(), vec![3]);
        assert_eq!(l.hidden_rows.iter().copied().collect::<Vec<_>>(), vec![2, 5], "rows untouched");
    }

    /// A layout that only carries hidden state is not "empty" — treating it as
    /// empty would drop the side-car and silently unhide everything.
    #[test]
    fn hidden_only_layout_is_not_empty() {
        assert!(!layout_with_hidden().is_empty());
        assert!(SheetLayout::default().is_empty());
    }
}

/// Parse "#RRGGBB" (or bare "RRGGBB") into RGBA. Alpha is always opaque —
/// visigrid-json writes tab colours without one.
fn parse_hex_rgb(s: &str) -> Option<[u8; 4]> {
    let h = s.trim_start_matches('#');
    if h.len() != 6 {
        return None;
    }
    Some([
        u8::from_str_radix(&h[0..2], 16).ok()?,
        u8::from_str_radix(&h[2..4], 16).ok()?,
        u8::from_str_radix(&h[4..6], 16).ok()?,
        255,
    ])
}
