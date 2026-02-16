//! Structured diff results view.
//!
//! Parses `vgrid diff --out json` output and populates a "Diff Results" sheet
//! with summary + detail rows. Keeps diff logic out of app.rs.

use std::path::{Path, PathBuf};

use gpui::Context;
use serde_json::Value;

use crate::app::Spreadsheet;

/// Maximum detail rows before truncation.
const MAX_DETAIL_ROWS: usize = 50_000;

/// Row colors (RGBA)
const COLOR_DIFF: [u8; 4] = [255, 243, 205, 255];      // Light amber
const COLOR_ONLY: [u8; 4] = [230, 230, 230, 255];       // Light gray (only_left / only_right)

// ============================================================================
// Parsed types
// ============================================================================

pub struct ParsedDiff {
    pub summary: ParsedSummary,
    pub details: Vec<ParsedDetailRow>,
    pub total_detail_count: usize,
    pub invocation: Option<String>,
}

pub struct ParsedSummary {
    pub left_rows: String,
    pub right_rows: String,
    pub matched: String,
    pub only_left: String,
    pub only_right: String,
    pub diffs: String,
    pub key: String,
}

pub struct ParsedDetailRow {
    pub status: String,
    pub key: String,
    pub column: String,
    pub left_value: String,
    pub right_value: String,
    pub delta: String,
}

// ============================================================================
// Parsing
// ============================================================================

pub fn parse_diff_json(bytes: &[u8]) -> Result<ParsedDiff, String> {
    let root: Value = serde_json::from_slice(bytes)
        .map_err(|e| format!("Invalid JSON: {}", e))?;

    // Enforce contract_version == 1
    let version = root.get("contract_version")
        .and_then(|v| v.as_u64());
    match version {
        Some(1) => {}
        Some(n) => return Err(format!("Unsupported diff contract version {}. Expected 1.", n)),
        None => return Err("Unsupported diff contract version (missing). Expected 1.".to_string()),
    }

    // Parse summary
    let summary_val = root.get("summary").unwrap_or(&Value::Null);
    let summary = ParsedSummary {
        left_rows: json_field_string(summary_val, "left_rows"),
        right_rows: json_field_string(summary_val, "right_rows"),
        matched: json_field_string(summary_val, "matched"),
        only_left: json_field_string(summary_val, "only_left"),
        only_right: json_field_string(summary_val, "only_right"),
        diffs: json_field_string(summary_val, "diff"),
        key: json_field_string(summary_val, "key"),
    };

    // Parse results → detail rows
    let results = root.get("results")
        .and_then(|v| v.as_array());

    let mut details = Vec::new();
    let mut total_detail_count = 0usize;

    if let Some(results) = results {
        for entry in results {
            let status = entry.get("status")
                .and_then(|v| v.as_str())
                .unwrap_or("");
            let key = entry.get("key")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .to_string();

            match status {
                "matched" => continue, // Exclude matched rows
                "diff" => {
                    let diffs = entry.get("diffs")
                        .and_then(|v| v.as_array());
                    if let Some(diffs) = diffs {
                        for d in diffs {
                            total_detail_count += 1;
                            if details.len() < MAX_DETAIL_ROWS {
                                details.push(ParsedDetailRow {
                                    status: "DIFF".to_string(),
                                    key: key.clone(),
                                    column: json_str(d, "column"),
                                    left_value: json_str(d, "left"),
                                    right_value: json_str(d, "right"),
                                    delta: json_str(d, "delta"),
                                });
                            }
                        }
                    }
                }
                "only_left" | "only_right" => {
                    total_detail_count += 1;
                    if details.len() < MAX_DETAIL_ROWS {
                        details.push(ParsedDetailRow {
                            status: status.to_uppercase().replace(' ', "_"),
                            key,
                            column: String::new(),
                            left_value: String::new(),
                            right_value: String::new(),
                            delta: String::new(),
                        });
                    }
                }
                _ => {
                    // Unknown status — include as-is
                    total_detail_count += 1;
                    if details.len() < MAX_DETAIL_ROWS {
                        details.push(ParsedDetailRow {
                            status: status.to_uppercase(),
                            key,
                            column: String::new(),
                            left_value: String::new(),
                            right_value: String::new(),
                            delta: String::new(),
                        });
                    }
                }
            }
        }
    }

    let invocation = root.get("invocation").and_then(|v| v.as_str()).map(String::from);

    Ok(ParsedDiff { summary, details, total_detail_count, invocation })
}

fn json_field_string(obj: &Value, key: &str) -> String {
    match obj.get(key) {
        Some(Value::String(s)) => s.clone(),
        Some(Value::Number(n)) => n.to_string(),
        Some(Value::Null) | None => String::new(),
        Some(v) => v.to_string(),
    }
}

fn json_str(obj: &Value, key: &str) -> String {
    obj.get(key)
        .and_then(|v| v.as_str())
        .unwrap_or("")
        .to_string()
}

// ============================================================================
// Stem sanitization
// ============================================================================

/// Sanitize a filename stem for use in diff output filenames.
/// Lowercase, collapse non-alphanumeric runs to `_`, trim leading/trailing `_`, cap at 32 chars.
pub fn sanitize_stem(name: &str) -> String {
    let stem = Path::new(name)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(name);

    let mut result = String::new();
    let mut last_was_sep = true; // Trim leading separators

    for ch in stem.chars() {
        if ch.is_alphanumeric() {
            result.push(ch.to_ascii_lowercase());
            last_was_sep = false;
        } else if !last_was_sep {
            result.push('_');
            last_was_sep = true;
        }
    }

    // Trim trailing underscore
    while result.ends_with('_') {
        result.pop();
    }

    // Cap at 32 chars (avoid breaking mid-underscore)
    if result.len() > 32 {
        result.truncate(32);
        while result.ends_with('_') {
            result.pop();
        }
    }

    if result.is_empty() {
        "file".to_string()
    } else {
        result
    }
}

// ============================================================================
// File discovery
// ============================================================================

/// Find the newest `diff-*.json` file in the workspace directory.
/// Sorts by (mtime desc, filename desc) for deterministic results.
pub fn find_latest_diff_file(workspace: &Path) -> Option<PathBuf> {
    let entries = std::fs::read_dir(workspace).ok()?;

    let mut candidates: Vec<(std::time::SystemTime, PathBuf)> = Vec::new();

    for entry in entries.flatten() {
        let name = entry.file_name();
        let name_str = name.to_string_lossy();
        if name_str.starts_with("diff-") && name_str.ends_with(".json") {
            if let Ok(meta) = entry.metadata() {
                if meta.is_file() {
                    let mtime = meta.modified().unwrap_or(std::time::UNIX_EPOCH);
                    candidates.push((mtime, entry.path()));
                }
            }
        }
    }

    if candidates.is_empty() {
        return None;
    }

    // Sort by mtime desc, then filename desc (for determinism)
    candidates.sort_by(|a, b| {
        b.0.cmp(&a.0).then_with(|| b.1.cmp(&a.1))
    });

    Some(candidates[0].1.clone())
}

// ============================================================================
// Provenance extraction
// ============================================================================

/// Extract human-readable source file names from a diff filename like
/// `diff-stripe_export_feb_vs_bank_statement.json` → ("stripe_export_feb", "bank_statement").
/// Falls back to the full filename if the pattern doesn't match.
fn extract_sources_from_filename(diff_path: &Path) -> (String, String) {
    let stem = diff_path.file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("diff");

    // Strip "diff-" prefix
    let body = stem.strip_prefix("diff-").unwrap_or(stem);

    if let Some(idx) = body.find("_vs_") {
        let left = &body[..idx];
        let right = &body[idx + 4..];
        (left.to_string(), right.to_string())
    } else {
        (body.to_string(), String::new())
    }
}

/// Format the current local time as "YYYY-MM-DD HH:MM".
fn now_timestamp() -> String {
    // Use std::time to get seconds since epoch, then do basic date math.
    // This avoids adding a chrono dependency.
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);

    // Apply local timezone offset via libc (POSIX only, but we target Linux/macOS)
    #[cfg(unix)]
    {
        let time: libc::time_t = secs as libc::time_t;
        let mut tm: libc::tm = unsafe { std::mem::zeroed() };
        unsafe { libc::localtime_r(&time, &mut tm); }
        let year = tm.tm_year + 1900;
        let mon = tm.tm_mon + 1;
        let day = tm.tm_mday;
        let hour = tm.tm_hour;
        let min = tm.tm_min;
        return format!("{:04}-{:02}-{:02} {:02}:{:02}", year, mon, day, hour, min);
    }

    #[cfg(not(unix))]
    {
        // Fallback: UTC
        let days = secs / 86400;
        let time_of_day = secs % 86400;
        let hour = time_of_day / 3600;
        let min = (time_of_day % 3600) / 60;

        // Approximate date from days since epoch (1970-01-01)
        // Good enough for display purposes
        let (year, month, day) = days_to_date(days);
        format!("{:04}-{:02}-{:02} {:02}:{:02} UTC", year, month, day, hour, min)
    }
}

#[cfg(not(unix))]
fn days_to_date(days: i64) -> (i64, i64, i64) {
    // Civil date from days since 1970-01-01 (Howard Hinnant's algorithm)
    let z = days + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    (y, m, d)
}

// ============================================================================
// Sheet population
// ============================================================================

// Provenance header layout
pub const ROW_TITLE: usize = 0;        // "Diff Report"
pub const ROW_SOURCE: usize = 1;       // "Generated from" | "left vs right"
pub const ROW_WORKSPACE: usize = 2;    // "Workspace" | "/path/..."
pub const ROW_TIMESTAMP: usize = 3;    // "Generated at" | "2026-02-15 14:30"
pub const ROW_COMMAND: usize = 4;      // "Command" | "vgrid diff ..."
pub const ROW_DIFF_FILE: usize = 5;    // "Diff file" | "./diff-foo_vs_bar.json"

/// Number of header/provenance rows before the summary section.
const PROVENANCE_ROWS: usize = 6;
/// Summary starts after provenance + 1 blank row.
pub const SUMMARY_START: usize = PROVENANCE_ROWS + 1; // row 7

/// Create (or replace) a "Diff Results" sheet and populate it with parsed diff data.
pub fn populate_diff_sheet(
    app: &mut Spreadsheet,
    parsed: ParsedDiff,
    diff_path: &Path,
    workspace: &Path,
    cx: &mut Context<Spreadsheet>,
) {
    let sheet_name = "Diff Results";

    // Delete existing "Diff Results" sheet if present
    let existing_idx = {
        let wb = app.workbook.read(cx);
        let names = wb.sheet_names();
        names.iter().position(|n| n.eq_ignore_ascii_case(sheet_name))
    };

    if let Some(idx) = existing_idx {
        let count = app.workbook.read(cx).sheet_count();
        if count > 1 {
            app.wb_mut(cx, |wb| { wb.delete_sheet(idx); });
        }
    }

    // Add new sheet
    let sheet_idx = app.wb_mut(cx, |wb| {
        wb.add_sheet_named(sheet_name)
    });

    let sheet_idx = match sheet_idx {
        Some(idx) => idx,
        None => {
            let mut fallback_idx = None;
            for n in 2..=10 {
                let name = format!("{} ({})", sheet_name, n);
                if let Some(idx) = app.wb_mut(cx, |wb| wb.add_sheet_named(&name)) {
                    fallback_idx = Some(idx);
                    break;
                }
            }
            match fallback_idx {
                Some(idx) => idx,
                None => {
                    app.status_message = Some("Could not create Diff Results sheet.".to_string());
                    cx.notify();
                    return;
                }
            }
        }
    };

    // Extract provenance info
    let (left_source, right_source) = extract_sources_from_filename(diff_path);
    let source_line = if right_source.is_empty() {
        left_source.clone()
    } else {
        format!("{} vs {}", left_source, right_source)
    };
    let workspace_display = workspace.display().to_string();
    let diff_file_display = diff_path.file_name()
        .and_then(|n| n.to_str())
        .map(|n| format!("./{}", n))
        .unwrap_or_else(|| diff_path.display().to_string());
    let timestamp = now_timestamp();

    // Layout constants
    let summary_start = SUMMARY_START; // row 5
    let detail_header_row = summary_start + 10; // row 15
    let detail_data_start = detail_header_row + 1; // row 16

    // Populate cells via batch guard
    app.workbook.update(cx, |wb, _| {
        let mut guard = wb.batch_guard();

        // -- Provenance header (rows 0-4) --
        guard.set_cell_value_tracked(sheet_idx, ROW_TITLE, 0, "Diff Report");
        if let Some(sheet) = guard.sheet_mut(sheet_idx) {
            sheet.toggle_bold(ROW_TITLE, 0);
        }

        guard.set_cell_value_tracked(sheet_idx, ROW_SOURCE, 0, "Generated from");
        guard.set_cell_value_tracked(sheet_idx, ROW_SOURCE, 1, &source_line);
        if let Some(sheet) = guard.sheet_mut(sheet_idx) {
            sheet.toggle_bold(ROW_SOURCE, 0);
        }

        guard.set_cell_value_tracked(sheet_idx, ROW_WORKSPACE, 0, "Workspace");
        guard.set_cell_value_tracked(sheet_idx, ROW_WORKSPACE, 1, &workspace_display);
        if let Some(sheet) = guard.sheet_mut(sheet_idx) {
            sheet.toggle_bold(ROW_WORKSPACE, 0);
        }

        guard.set_cell_value_tracked(sheet_idx, ROW_TIMESTAMP, 0, "Generated at");
        guard.set_cell_value_tracked(sheet_idx, ROW_TIMESTAMP, 1, &timestamp);
        if let Some(sheet) = guard.sheet_mut(sheet_idx) {
            sheet.toggle_bold(ROW_TIMESTAMP, 0);
        }

        if let Some(ref cmd) = parsed.invocation {
            guard.set_cell_value_tracked(sheet_idx, ROW_COMMAND, 0, "Command");
            guard.set_cell_value_tracked(sheet_idx, ROW_COMMAND, 1, cmd);
            if let Some(sheet) = guard.sheet_mut(sheet_idx) {
                sheet.toggle_bold(ROW_COMMAND, 0);
            }
        }

        guard.set_cell_value_tracked(sheet_idx, ROW_DIFF_FILE, 0, "Diff file");
        guard.set_cell_value_tracked(sheet_idx, ROW_DIFF_FILE, 1, &diff_file_display);
        if let Some(sheet) = guard.sheet_mut(sheet_idx) {
            sheet.toggle_bold(ROW_DIFF_FILE, 0);
        }

        // -- Summary section --
        let labels = [
            (summary_start, "Left rows", &parsed.summary.left_rows),
            (summary_start + 1, "Right rows", &parsed.summary.right_rows),
            (summary_start + 2, "Matched", &parsed.summary.matched),
            (summary_start + 3, "Only left", &parsed.summary.only_left),
            (summary_start + 4, "Only right", &parsed.summary.only_right),
            (summary_start + 5, "Differences", &parsed.summary.diffs),
            (summary_start + 6, "Key column", &parsed.summary.key),
        ];

        for (row, label, value) in &labels {
            guard.set_cell_value_tracked(sheet_idx, *row, 0, label);
            guard.set_cell_value_tracked(sheet_idx, *row, 1, value);
            if let Some(sheet) = guard.sheet_mut(sheet_idx) {
                sheet.toggle_bold(*row, 0);
            }
        }

        // -- Detail header --
        let headers = ["Status", "Key", "Column", "Left Value", "Right Value", "Delta"];
        for (col, header) in headers.iter().enumerate() {
            guard.set_cell_value_tracked(sheet_idx, detail_header_row, col, header);
            if let Some(sheet) = guard.sheet_mut(sheet_idx) {
                sheet.toggle_bold(detail_header_row, col);
            }
        }

        // -- Detail rows with color coding --
        for (i, row_data) in parsed.details.iter().enumerate() {
            let row = detail_data_start + i;
            guard.set_cell_value_tracked(sheet_idx, row, 0, &row_data.status);
            guard.set_cell_value_tracked(sheet_idx, row, 1, &row_data.key);
            guard.set_cell_value_tracked(sheet_idx, row, 2, &row_data.column);
            guard.set_cell_value_tracked(sheet_idx, row, 3, &row_data.left_value);
            guard.set_cell_value_tracked(sheet_idx, row, 4, &row_data.right_value);
            guard.set_cell_value_tracked(sheet_idx, row, 5, &row_data.delta);

            // Color the entire row based on status
            let color = match row_data.status.as_str() {
                "DIFF" => Some(COLOR_DIFF),
                "ONLY_LEFT" | "ONLY_RIGHT" => Some(COLOR_ONLY),
                _ => None,
            };
            if let Some(color) = color {
                if let Some(sheet) = guard.sheet_mut(sheet_idx) {
                    for col in 0..6 {
                        sheet.set_background_color(row, col, Some(color));
                    }
                }
            }
        }
    });

    // Switch to the new sheet
    app.wb_mut(cx, |wb| { wb.set_active_sheet(sheet_idx); });
    app.update_cached_sheet_id(cx);
    // Reset selection to A1 for the new sheet
    app.view_state.selected = (0, 0);
    app.view_state.selection_end = None;
    app.view_state.scroll_row = 0;
    app.view_state.scroll_col = 0;

    // Status message
    let truncation_note = if parsed.total_detail_count > MAX_DETAIL_ROWS {
        format!(" (showing first {} of {} rows)", MAX_DETAIL_ROWS, parsed.total_detail_count)
    } else {
        String::new()
    };
    app.status_message = Some(format!(
        "Diff Results: {} detail rows loaded{}",
        parsed.details.len(),
        truncation_note,
    ));

    app.is_modified = true;
    cx.notify();
}
