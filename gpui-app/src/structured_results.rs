//! Parse and render structured JSON results from vgrid CLI commands.
//!
//! Bridges terminal output back into the spreadsheet: parse JSON from
//! `vgrid --json` commands and create new sheets with the data.

use visigrid_engine::workbook::Workbook;

/// A structured result extracted from terminal output.
pub enum StructuredResult {
    /// `vgrid diff --json` output: contract_version + summary + results
    Diff {
        /// Raw JSON value containing the full diff output
        raw: serde_json::Value,
    },
    /// `vgrid peek --json` output: columns + rows
    Peek {
        columns: Vec<String>,
        rows: Vec<Vec<serde_json::Value>>,
    },
    /// `vgrid calc --json` output: scalar or 2D array
    Calc {
        value: serde_json::Value,
    },
}

impl StructuredResult {
    /// Human-readable description for the affordance bar.
    pub fn description(&self) -> String {
        match self {
            Self::Diff { raw } => {
                let count = raw.get("results")
                    .and_then(|r| r.as_array())
                    .map(|a| a.len())
                    .unwrap_or(0);
                format!("Diff result \u{00b7} {} rows", count)
            }
            Self::Peek { rows, .. } => {
                format!("Peek result \u{00b7} {} rows", rows.len())
            }
            Self::Calc { value } => {
                if let Some(arr) = value.as_array() {
                    let rows = arr.len();
                    let cols = arr.first().and_then(|r| r.as_array()).map(|a| a.len()).unwrap_or(0);
                    format!("Calc result \u{00b7} {}x{}", rows, cols)
                } else {
                    "Calc result \u{00b7} scalar".to_string()
                }
            }
        }
    }
}

/// Attempt to parse a text block as a structured vgrid result.
///
/// Strategy:
/// 1. Try `serde_json::from_str` on the full text block first
/// 2. If that fails, scan line-by-line for `{` or `[` at line start
/// 3. Validate shape strictly and classify
pub fn parse(text: &str) -> Option<StructuredResult> {
    // Strategy 1: try parsing the whole text as JSON
    if let Some(result) = try_parse_value(text) {
        return Some(result);
    }

    // Strategy 2: scan for JSON start on a line boundary
    for line in text.lines() {
        let trimmed = line.trim_start();
        if trimmed.starts_with('{') || trimmed.starts_with('[') {
            // Try parsing from this line to end of text
            let offset = line.as_ptr() as usize - text.as_ptr() as usize;
            let candidate = &text[offset..];
            if let Some(result) = try_parse_value(candidate) {
                return Some(result);
            }
        }
    }

    None
}

/// Try to parse a string as JSON and classify it.
fn try_parse_value(text: &str) -> Option<StructuredResult> {
    let val: serde_json::Value = serde_json::from_str(text.trim()).ok()?;
    classify(val)
}

/// Classify a parsed JSON value into a StructuredResult variant.
fn classify(val: serde_json::Value) -> Option<StructuredResult> {
    match &val {
        serde_json::Value::Object(obj) => {
            // Diff: has contract_version field
            if obj.contains_key("contract_version") {
                return Some(StructuredResult::Diff { raw: val });
            }
            // Peek: has columns array + rows array-of-arrays
            if let (Some(cols), Some(rows)) = (obj.get("columns"), obj.get("rows")) {
                if let (Some(cols_arr), Some(rows_arr)) = (cols.as_array(), rows.as_array()) {
                    let columns: Vec<String> = cols_arr.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect();
                    if columns.len() != cols_arr.len() {
                        return None; // Non-string column names
                    }
                    let rows: Vec<Vec<serde_json::Value>> = rows_arr.iter()
                        .filter_map(|r| r.as_array().cloned())
                        .collect();
                    if rows.len() != rows_arr.len() {
                        return None; // Non-array row
                    }
                    return Some(StructuredResult::Peek { columns, rows });
                }
            }
            // Single scalar wrapped in {"value": ...}
            if let Some(inner) = obj.get("value") {
                return Some(StructuredResult::Calc { value: inner.clone() });
            }
            None
        }
        serde_json::Value::Array(arr) => {
            // 2D array → Calc (spill)
            if arr.first().map_or(false, |r| r.is_array()) {
                return Some(StructuredResult::Calc { value: val });
            }
            None
        }
        // Bare scalar → Calc
        serde_json::Value::Number(_) | serde_json::Value::Bool(_) | serde_json::Value::String(_) => {
            Some(StructuredResult::Calc { value: val })
        }
        _ => None,
    }
}

/// Metadata about a rendered result (for the Run Log).
pub struct ResultMeta {
    pub sheet_idx: usize,
    pub sheet_name: String,
    pub result_type: &'static str,
    pub row_count: usize,
    pub col_count: usize,
}

/// Render a structured result into a new sheet in the workbook.
/// Returns metadata including the sheet index and name.
pub fn render_to_sheet(result: &StructuredResult, wb: &mut Workbook) -> ResultMeta {
    match result {
        StructuredResult::Diff { raw } => render_diff(raw, wb),
        StructuredResult::Peek { columns, rows } => render_peek(columns, rows, wb),
        StructuredResult::Calc { value } => render_calc(value, wb),
    }
}

/// Name of the Run Log index sheet.
const RUN_LOG_SHEET: &str = "Run Log";
/// Run Log header columns.
const RUN_LOG_HEADERS: &[&str] = &["Timestamp", "Type", "Rows", "Cols", "Sheet", "Source", "Command"];
/// Maximum data rows per Run Log sheet before rotating.
const RUN_LOG_MAX_ROWS: usize = 1000;

/// Append a row to the Run Log sheet, creating it if needed.
/// Rotates to "Run Log (2)", "(3)", etc. when a sheet exceeds 1000 data rows.
pub fn append_run_log(
    wb: &mut Workbook,
    meta: &ResultMeta,
    source_file: Option<&str>,
    command: Option<&str>,
) {
    // Find the latest Run Log sheet (prefer highest-numbered)
    let log_idx = find_or_create_run_log(wb);

    // Find next empty row
    let prev_active = wb.active_sheet_index();
    let _ = wb.set_active_sheet(log_idx);
    let sheet = wb.active_sheet_mut();
    let mut next_row = 1;
    while !sheet.get_display(next_row, 0).is_empty() {
        next_row += 1;
        if next_row > 10_000 { break; } // safety cap
    }

    // If this sheet is full, rotate to a new one
    let (log_idx, next_row) = if next_row > RUN_LOG_MAX_ROWS {
        let _ = wb.set_active_sheet(prev_active);
        let new_idx = create_run_log_sheet(wb);
        let _ = wb.set_active_sheet(new_idx);
        (new_idx, 1)
    } else {
        (log_idx, next_row)
    };

    // Timestamp: ISO 8601 compact
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let timestamp = format_epoch_timestamp(now);

    let sheet = wb.active_sheet_mut();
    sheet.set_value(next_row, 0, &timestamp);
    sheet.set_value(next_row, 1, meta.result_type);
    sheet.set_value(next_row, 2, &meta.row_count.to_string());
    sheet.set_value(next_row, 3, &meta.col_count.to_string());
    sheet.set_value(next_row, 4, &meta.sheet_name);
    sheet.set_value(next_row, 5, source_file.unwrap_or(""));
    sheet.set_value(next_row, 6, command.unwrap_or(""));

    // Restore active sheet (use log_idx to suppress unused warning)
    let _ = log_idx;
    let _ = wb.set_active_sheet(prev_active);
}

/// Find the latest Run Log sheet, or create the first one.
fn find_or_create_run_log(wb: &mut Workbook) -> usize {
    // Find all sheets named "Run Log" or "Run Log (N)"
    let mut best_idx = None;
    for i in 0..wb.sheet_count() {
        if let Some(s) = wb.sheet(i) {
            if s.name == RUN_LOG_SHEET || s.name.starts_with("Run Log (") {
                best_idx = Some(i);
            }
        }
    }
    best_idx.unwrap_or_else(|| create_run_log_sheet(wb))
}

/// Create a new Run Log sheet with headers, using dedup naming.
fn create_run_log_sheet(wb: &mut Workbook) -> usize {
    let name = unique_sheet_name(wb, RUN_LOG_SHEET);
    let idx = wb.add_sheet_named(&name).unwrap_or_else(|| wb.add_sheet());
    let prev_active = wb.active_sheet_index();
    let _ = wb.set_active_sheet(idx);
    let sheet = wb.active_sheet_mut();
    for (c, header) in RUN_LOG_HEADERS.iter().enumerate() {
        sheet.set_value(0, c, header);
    }
    let _ = wb.set_active_sheet(prev_active);
    idx
}

/// Format epoch seconds as a human-readable timestamp (UTC, compact).
fn format_epoch_timestamp(secs: u64) -> String {
    // Simple UTC formatting without chrono dependency
    let days = secs / 86400;
    let time_of_day = secs % 86400;
    let hours = time_of_day / 3600;
    let minutes = (time_of_day % 3600) / 60;
    let seconds = time_of_day % 60;

    // Days since epoch to Y-M-D (simplified, good enough for logging)
    let (year, month, day) = days_to_ymd(days);
    format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hours, minutes, seconds)
}

/// Convert days since 1970-01-01 to (year, month, day).
fn days_to_ymd(mut days: u64) -> (u64, u64, u64) {
    let mut year = 1970u64;
    loop {
        let days_in_year = if is_leap(year) { 366 } else { 365 };
        if days < days_in_year { break; }
        days -= days_in_year;
        year += 1;
    }
    let month_days: &[u64] = if is_leap(year) {
        &[31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        &[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };
    let mut month = 1u64;
    for &md in month_days {
        if days < md { break; }
        days -= md;
        month += 1;
    }
    (year, month, days + 1)
}

fn is_leap(y: u64) -> bool {
    y % 4 == 0 && (y % 100 != 0 || y % 400 == 0)
}

/// Find a unique sheet name by appending (2), (3), etc.
pub(crate) fn unique_sheet_name(wb: &Workbook, base: &str) -> String {
    if !wb.sheet_name_exists(base) {
        return base.to_string();
    }
    for i in 2..=100 {
        let candidate = format!("{} ({})", base, i);
        if !wb.sheet_name_exists(&candidate) {
            return candidate;
        }
    }
    // Fallback: timestamp-based
    format!("{} ({})", base, std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs())
}

fn render_diff(raw: &serde_json::Value, wb: &mut Workbook) -> ResultMeta {
    let name = unique_sheet_name(wb, "Diff Results");
    let sheet_idx = wb.add_sheet_named(&name).unwrap_or_else(|| wb.add_sheet());
    let _ = wb.set_active_sheet(sheet_idx);

    let results = raw.get("results").and_then(|r| r.as_array());
    let Some(results) = results else {
        // Just put the summary if no results array
        wb.active_sheet_mut().set_value(0, 0, "No diff results found");
        return ResultMeta {
            sheet_idx,
            sheet_name: name,
            result_type: "diff",
            row_count: 0,
            col_count: 1,
        };
    };

    // Collect all column names from the diffs
    let mut col_names: Vec<String> = Vec::new();
    let mut col_set = std::collections::HashSet::new();
    for row in results {
        if let Some(diffs) = row.get("diffs").and_then(|d| d.as_array()) {
            for diff in diffs {
                if let Some(col) = diff.get("column").and_then(|c| c.as_str()) {
                    if col_set.insert(col.to_string()) {
                        col_names.push(col.to_string());
                    }
                }
            }
        }
    }

    // Header row: Status | Key | col1_left | col1_right | col1_delta | ...
    let sheet = wb.active_sheet_mut();
    sheet.set_value(0, 0, "Status");
    sheet.set_value(0, 1, "Key");
    let mut c = 2;
    for col_name in &col_names {
        sheet.set_value(0, c, &format!("{} (left)", col_name));
        sheet.set_value(0, c + 1, &format!("{} (right)", col_name));
        sheet.set_value(0, c + 2, &format!("{} (delta)", col_name));
        c += 3;
    }

    // Data rows
    for (r, row) in results.iter().enumerate() {
        let sheet = wb.active_sheet_mut();
        let data_row = r + 1;
        let status = row.get("status").and_then(|s| s.as_str()).unwrap_or("");
        let key = row.get("key").and_then(|k| k.as_str())
            .or_else(|| row.get("key").map(|k| k.to_string()).as_deref().map(|_| ""))
            .unwrap_or("");
        sheet.set_value(data_row, 0, status);
        // Key might be non-string
        let key_str = match row.get("key") {
            Some(serde_json::Value::String(s)) => s.clone(),
            Some(v) => v.to_string(),
            None => String::new(),
        };
        sheet.set_value(data_row, 1, &key_str);

        // Build a map of column -> diff for this row
        let diffs: std::collections::HashMap<&str, &serde_json::Value> = row.get("diffs")
            .and_then(|d| d.as_array())
            .map(|arr| {
                arr.iter()
                    .filter_map(|d| d.get("column").and_then(|c| c.as_str()).map(|c| (c, d)))
                    .collect()
            })
            .unwrap_or_default();

        let mut c = 2;
        for col_name in &col_names {
            if let Some(diff) = diffs.get(col_name.as_str()) {
                let left = json_display(diff.get("left"));
                let right = json_display(diff.get("right"));
                let delta = json_display(diff.get("delta"));
                sheet.set_value(data_row, c, &left);
                sheet.set_value(data_row, c + 1, &right);
                sheet.set_value(data_row, c + 2, &delta);
            }
            c += 3;
        }
    }

    let total_cols = 2 + col_names.len() * 3;
    ResultMeta {
        sheet_idx,
        sheet_name: name,
        result_type: "diff",
        row_count: results.len(),
        col_count: total_cols,
    }
}

fn render_peek(columns: &[String], rows: &[Vec<serde_json::Value>], wb: &mut Workbook) -> ResultMeta {
    let name = unique_sheet_name(wb, "Peek Results");
    let sheet_idx = wb.add_sheet_named(&name).unwrap_or_else(|| wb.add_sheet());
    let _ = wb.set_active_sheet(sheet_idx);

    let sheet = wb.active_sheet_mut();
    // Header row
    for (c, col_name) in columns.iter().enumerate() {
        sheet.set_value(0, c, col_name);
    }
    // Data rows
    for (r, row) in rows.iter().enumerate() {
        for (c, val) in row.iter().enumerate() {
            sheet.set_value(r + 1, c, &json_cell_value(val));
        }
    }

    ResultMeta {
        sheet_idx,
        sheet_name: name,
        result_type: "peek",
        row_count: rows.len(),
        col_count: columns.len(),
    }
}

fn render_calc(value: &serde_json::Value, wb: &mut Workbook) -> ResultMeta {
    let name = unique_sheet_name(wb, "Calc Results");
    let sheet_idx = wb.add_sheet_named(&name).unwrap_or_else(|| wb.add_sheet());
    let _ = wb.set_active_sheet(sheet_idx);

    let sheet = wb.active_sheet_mut();
    let (row_count, col_count) = match value {
        serde_json::Value::Array(rows) => {
            let mut max_cols = 0;
            for (r, row) in rows.iter().enumerate() {
                if let Some(cells) = row.as_array() {
                    for (c, val) in cells.iter().enumerate() {
                        sheet.set_value(r, c, &json_cell_value(val));
                    }
                    max_cols = max_cols.max(cells.len());
                } else {
                    // 1D array: put each element in a column
                    sheet.set_value(r, 0, &json_cell_value(row));
                    max_cols = max_cols.max(1);
                }
            }
            (rows.len(), max_cols)
        }
        _ => {
            // Scalar: single cell A1
            sheet.set_value(0, 0, &json_cell_value(value));
            (1, 1)
        }
    };

    ResultMeta {
        sheet_idx,
        sheet_name: name,
        result_type: "calc",
        row_count,
        col_count,
    }
}

/// Public accessor for json_cell_value (used by AI explain prompt building).
pub fn json_cell_value_pub(val: &serde_json::Value) -> String {
    json_cell_value(val)
}

/// Convert a JSON value to a cell-appropriate string.
fn json_cell_value(val: &serde_json::Value) -> String {
    match val {
        serde_json::Value::Null => String::new(),
        serde_json::Value::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::String(s) => s.clone(),
        other => other.to_string(),
    }
}

/// Display a JSON value for diff cells (handles None / missing values).
fn json_display(val: Option<&serde_json::Value>) -> String {
    match val {
        None => String::new(),
        Some(v) => json_cell_value(v),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_peek_json() {
        let json = r#"{"columns":["id","name"],"rows":[[1,"Alice"],[2,"Bob"]]}"#;
        let result = parse(json).expect("should parse");
        match result {
            StructuredResult::Peek { columns, rows } => {
                assert_eq!(columns, vec!["id", "name"]);
                assert_eq!(rows.len(), 2);
            }
            _ => panic!("expected Peek"),
        }
    }

    #[test]
    fn test_parse_diff_json() {
        let json = r#"{"contract_version":1,"summary":{},"results":[]}"#;
        let result = parse(json).expect("should parse");
        assert!(matches!(result, StructuredResult::Diff { .. }));
    }

    #[test]
    fn test_parse_calc_scalar() {
        let json = "42";
        let result = parse(json).expect("should parse");
        match result {
            StructuredResult::Calc { value } => {
                assert_eq!(value.as_i64(), Some(42));
            }
            _ => panic!("expected Calc"),
        }
    }

    #[test]
    fn test_parse_calc_array() {
        let json = "[[1,2],[3,4]]";
        let result = parse(json).expect("should parse");
        assert!(matches!(result, StructuredResult::Calc { .. }));
    }

    #[test]
    fn test_parse_with_prefix_noise() {
        let text = "Loading data...\nProcessing...\n{\"columns\":[\"a\"],\"rows\":[[1]]}";
        let result = parse(text).expect("should parse");
        assert!(matches!(result, StructuredResult::Peek { .. }));
    }

    #[test]
    fn test_parse_garbage_returns_none() {
        assert!(parse("not json at all").is_none());
        assert!(parse("").is_none());
    }

    #[test]
    fn test_description() {
        let peek = StructuredResult::Peek {
            columns: vec!["a".into()],
            rows: vec![vec![serde_json::Value::from(1)]],
        };
        assert_eq!(peek.description(), "Peek result \u{00b7} 1 rows");
    }
}
