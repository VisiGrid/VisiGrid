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
// {
//   "format": "visigrid-json",
//   "version": 1,
//   "name": "Sheet1",
//   "cells": [
//     {"row":0, "col":0, "value":"Item", "fmt":{"bold":true, "bg":"#FFEB3B"}},
//     {"row":1, "col":2, "formula":"=A2*B2", "value":85}
//   ],
//   "merges": [{"start_row":0,"start_col":0,"end_row":0,"end_col":2}]
// }
//
// Formula cells carry both the formula and the last computed value, so
// consumers without an engine still see data. On import, formulas are
// recomputed; the stored value is a fallback only.

use serde::{Deserialize, Serialize};
use visigrid_engine::cell::{Alignment, CellStyle, CellValue, VerticalAlignment};
use visigrid_engine::sheet::MergedRegion;

pub const FULL_JSON_FORMAT: &str = "visigrid-json";
pub const FULL_JSON_VERSION: u32 = 1;

#[derive(Serialize, Deserialize)]
struct FullDoc {
    format: String,
    version: u32,
    #[serde(default)]
    name: String,
    #[serde(default)]
    cells: Vec<FullCell>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    merges: Vec<MergeSpec>,
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

/// Export a sheet as visigrid-json v1.
pub fn export_full(sheet: &Sheet) -> Result<String, String> {
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

    let doc = FullDoc {
        format: FULL_JSON_FORMAT.to_string(),
        version: FULL_JSON_VERSION,
        name: sheet.name.clone(),
        cells,
        merges,
    };
    serde_json::to_string_pretty(&doc).map_err(|e| e.to_string())
}

/// Import visigrid-json into a Sheet (formulas recomputed).
pub fn import_full(content: &str) -> Result<Sheet, String> {
    use visigrid_engine::sheet::SheetId;
    use visigrid_engine::workbook::Workbook;

    let doc: FullDoc = serde_json::from_str(content).map_err(|e| format!("invalid visigrid-json: {}", e))?;
    if doc.format != FULL_JSON_FORMAT {
        return Err(format!("not a visigrid-json document (format: {:?})", doc.format));
    }
    if doc.version > FULL_JSON_VERSION {
        return Err(format!(
            "visigrid-json version {} is newer than supported ({})",
            doc.version, FULL_JSON_VERSION
        ));
    }

    let mut sheet = Sheet::new(SheetId(1), 65536, 256);
    if !doc.name.is_empty() {
        sheet.set_name(&doc.name);
    }

    for cell in &doc.cells {
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

    for m in &doc.merges {
        let _ = sheet.add_merge(MergedRegion::new(m.start_row, m.start_col, m.end_row, m.end_col));
    }

    // Recompute formulas (stored values are only a fallback for engine-less consumers)
    let mut wb = Workbook::from_sheets(vec![sheet], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok(wb.sheets()[0].clone())
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
    fn rejects_foreign_documents() {
        assert!(import_full("[[1,2],[3,4]]").is_err());
        assert!(import_full("{\"format\":\"other\",\"version\":1}").is_err());
        assert!(import_full("{\"format\":\"visigrid-json\",\"version\":99}").is_err());
    }
}
