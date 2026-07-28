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
//   "row_heights": {"3": 40.0},
//   "frozen_rows": 1,
//   "frozen_cols": 0
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

use std::collections::BTreeMap;

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
}

impl SheetLayout {
    pub fn is_empty(&self) -> bool {
        self.col_widths.is_empty()
            && self.row_heights.is_empty()
            && self.frozen_rows == 0
            && self.frozen_cols == 0
    }
}

fn is_zero(n: &usize) -> bool {
    *n == 0
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
    #[serde(default, skip_serializing_if = "is_zero")]
    active_sheet: usize,
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
    #[serde(default, skip_serializing_if = "is_zero")]
    frozen_rows: usize,
    #[serde(default, skip_serializing_if = "is_zero")]
    frozen_cols: usize,
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
        active_sheet: 0,
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
        active_sheet: active_sheet.min(sheets.len().saturating_sub(1)),
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

        cells.push(FullCell { row, col, value, formula, fmt });
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
        cells,
        merges,
        col_widths: keys_to_string(&layout.col_widths),
        row_heights: keys_to_string(&layout.row_heights),
        frozen_rows: layout.frozen_rows,
        frozen_cols: layout.frozen_cols,
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
        let (sheet, layout) = apply_body(body, SheetId(i as u64 + 1), i);
        sheets.push(sheet);
        layouts.push(layout);
    }

    let active = doc.active_sheet.min(sheets.len() - 1);
    // Recompute formulas (stored values are only a fallback for engine-less consumers)
    let mut wb = Workbook::from_sheets(sheets, active);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok((wb, layouts, active))
}

fn apply_body(body: &SheetBody, id: visigrid_engine::sheet::SheetId, index: usize) -> (Sheet, SheetLayout) {
    let mut sheet = Sheet::new(id, 65536, 256);
    if !body.name.is_empty() {
        sheet.set_name(&body.name);
    } else if index > 0 {
        sheet.set_name(&format!("Sheet{}", index + 1));
    }

    for cell in &body.cells {
        // Content: formula wins; else typed value
        if let Some(f) = &cell.formula {
            sheet.set_value(cell.row, cell.col, f);
        } else if let Some(v) = &cell.value {
            let text = match v {
                serde_json::Value::String(s) => s.clone(),
                serde_json::Value::Number(n) => n.to_string(),
                serde_json::Value::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
                other => other.to_string(),
            };
            sheet.set_value(cell.row, cell.col, &text);
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

    let layout = SheetLayout {
        col_widths: keys_to_usize(&body.col_widths),
        row_heights: keys_to_usize(&body.row_heights),
        frozen_rows: body.frozen_rows,
        frozen_cols: body.frozen_cols,
    };
    (sheet, layout)
}

#[cfg(test)]
mod full_json_tests {
    use super::*;
    use visigrid_engine::sheet::SheetId;

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
