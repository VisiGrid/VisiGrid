//! What the grid draws, read from a workbook: the sheet list, a sheet's
//! layout, and the cells of a rectangle. The browser replica
//! (`CollabClient.viewport` and friends) and the collaboration host's
//! `first_screen` share these, so a first screen made on the server draws
//! exactly as the full grid does.

use std::collections::HashMap;

use serde_json::{json, Value};
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;

use crate::op::SheetKey;

/// The sheets, in tab order: `[{key, name, index}]`.
pub fn sheets(wb: &Workbook) -> Value {
    Value::Array(
        wb.sheets().iter().enumerate().map(|(index, s)| json!({"key": s.id.0, "name": s.name, "index": index})).collect(),
    )
}

/// A sheet's layout for drawing: column widths and row heights that
/// differ from the defaults, frozen and hidden rows and columns, and the
/// extent of its data (`rows`, `cols`: one past the last used row and
/// column).
pub fn layout(wb: &Workbook, sheet: SheetKey) -> Option<Value> {
    let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
    let s = &wb.sheets()[idx];
    let (max_row, max_col) = s.data_extent();
    let used = !s.get_raw(max_row, max_col).is_empty() || max_row > 0 || max_col > 0;
    let l = &s.layout;
    let merges: Vec<Value> = s
        .merged_regions
        .iter()
        .map(|m| json!({"r0": m.start.0, "c0": m.start.1, "r1": m.end.0, "c1": m.end.1}))
        .collect();
    Some(json!({
        "col_widths": l.col_widths,
        "row_heights": l.row_heights,
        "frozen_rows": l.frozen_rows,
        "frozen_cols": l.frozen_cols,
        "hidden_rows": l.hidden_rows,
        "hidden_cols": l.hidden_cols,
        "merges": merges,
        "rows": if used { max_row + 1 } else { 0 },
        "cols": if used { max_col + 1 } else { 0 },
    }))
}

/// What the grid draws for the cells of a rectangle (inclusive), read
/// straight from the engine: the display string (number formats
/// applied), the computed value's kind (`n` number, `t` text, `b`
/// boolean, `e` error: general alignment depends on it) and an index
/// into `formats` (`-1`: the default format), and `num`, a number's value
/// (null otherwise) for selection statistics. Only cells with content
/// or a format are listed, as parallel arrays. `row_formats` and
/// `col_formats` carry whole-row and whole-column formats in range.
pub fn viewport(wb: &Workbook, sheet: SheetKey, r0: usize, c0: usize, r1: usize, c1: usize) -> Option<Value> {
    use visigrid_engine::cell::CellFormat;
    use visigrid_engine::formula::eval::Value as V;
    let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
    let s = &wb.sheets()[idx];
    let default = CellFormat::default();
    let mut formats: Vec<Value> = Vec::new();
    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut format_index = |f: &CellFormat| -> i64 {
        if *f == default {
            return -1;
        }
        let props = serde_json::to_value(crate::undo::props_of(f, None)).expect("props serialize");
        let key = props.to_string();
        let next = formats.len();
        let at = *seen.entry(key).or_insert(next);
        if at == next {
            formats.push(props);
        }
        at as i64
    };
    let mut coords = s.cells_in_range(r0, r1, c0, c1);
    coords.extend(s.spill_receiver_coords().filter(|&(r, c)| r >= r0 && r <= r1 && c >= c0 && c <= c1));
    coords.sort_unstable();
    coords.dedup();
    let (mut rows, mut cols, mut text, mut kind, mut fmt, mut num, mut raw) =
        (Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new());
    for (r, c) in coords {
        let shown = s.get_formatted_display(r, c);
        let f = format_index(&s.get_format(r, c));
        if shown.is_empty() && f < 0 {
            continue;
        }
        let (k, n) = match s.get_computed_value(r, c) {
            V::Number(n) => ("n", Some(n)),
            V::Boolean(_) => ("b", None),
            V::Error(_) => ("e", None),
            _ => ("t", None),
        };
        rows.push(r);
        cols.push(c);
        text.push(shown);
        kind.push(k);
        fmt.push(f);
        num.push(n.filter(|n| n.is_finite()));
        raw.push(s.get_raw(r, c));
    }
    let mut row_formats = serde_json::Map::new();
    for (r, f) in &s.row_formats {
        if *r >= r0 && *r <= r1 {
            let i = format_index(f);
            if i >= 0 {
                row_formats.insert(r.to_string(), json!(i));
            }
        }
    }
    let mut col_formats = serde_json::Map::new();
    for (c, f) in &s.col_formats {
        if *c >= c0 && *c <= c1 {
            let i = format_index(f);
            if i >= 0 {
                col_formats.insert(c.to_string(), json!(i));
            }
        }
    }
    Some(json!({
        "rows": rows, "cols": cols, "text": text, "kind": kind, "fmt": fmt, "num": num, "raw": raw,
        "formats": formats, "row_formats": row_formats, "col_formats": col_formats,
    }))
}

/// Rows and columns a first screen covers from A1: enough for a large
/// window at 100% (a 4K screen shows about 70 rows and 30 columns), and
/// whole tiles of the web grid's cache (64 × 32), so the page can answer
/// those tiles from it exactly.
pub const FIRST_SCREEN_ROWS: usize = 128;
pub const FIRST_SCREEN_COLS: usize = 32;

/// What the grid shows the moment a workbook opens, before the engine has
/// loaded in the browser: the sheet list, and for the first tab (the one
/// the grid opens on) its layout and the cells of its top-left
/// `rows` × `cols`, scrolled to A1. Same shapes as `sheets`, `layout` and
/// `viewport`.
pub fn first_screen(wb: &Workbook, rows: usize, cols: usize) -> Option<Value> {
    let sheet = wb.sheets().first()?.id.0;
    Some(json!({
        "sheets": sheets(wb),
        "sheet": sheet,
        "layout": layout(wb, sheet)?,
        "viewport": viewport(wb, sheet, 0, 0, rows.max(1) - 1, cols.max(1) - 1)?,
        "rows": rows,
        "cols": cols,
    }))
}
