//! `vgrid fill` — load CSV data into a .sheet template.
//!
//! Strict CSV parsing for financial data: integers, exact 2-decimal
//! amounts, and text. No auto-detection of dates, booleans, or
//! percentages. Rejects currency symbols, commas in numbers, and
//! formula injection (`=` prefix).

use std::path::Path;

use crate::sheet_ops::parse_cell_ref;
use crate::CliError;

// ── Fill target parsing ─────────────────────────────────────────────

/// Parsed `sheet!cell` target reference.
#[derive(Debug, PartialEq)]
pub struct FillTarget {
    pub sheet_name: Option<String>,
    pub row: usize,
    pub col: usize,
}

/// Parse a fill target like `tx!A1` or `B5`.
pub fn parse_fill_target(s: &str) -> Result<FillTarget, CliError> {
    if let Some(bang) = s.find('!') {
        let sheet = &s[..bang];
        let cell = &s[bang + 1..];
        if sheet.is_empty() {
            return Err(CliError::args(format!(
                "empty sheet name in target {:?}",
                s
            )));
        }
        let (row, col) = parse_cell_ref(cell).ok_or_else(|| {
            CliError::args(format!("invalid cell reference {:?} in target {:?}", cell, s))
        })?;
        Ok(FillTarget {
            sheet_name: Some(sheet.to_string()),
            row,
            col,
        })
    } else {
        let (row, col) = parse_cell_ref(s).ok_or_else(|| {
            CliError::args(format!("invalid cell reference {:?}", s))
        })?;
        Ok(FillTarget {
            sheet_name: None,
            row,
            col,
        })
    }
}

// ── Strict CSV value parsing ────────────────────────────────────────

/// A strictly-parsed CSV value.
#[derive(Debug, PartialEq)]
pub enum StrictValue {
    Empty,
    Integer(f64),
    Decimal(f64),
    Text(String),
}

/// Check if a string looks like a number with a currency symbol.
fn has_currency_symbol(s: &str) -> bool {
    s.contains('$') || s.contains('€') || s.contains('£')
        || s.starts_with("USD") || s.starts_with("EUR") || s.starts_with("GBP")
}

/// Check if a numeric-looking string contains commas (e.g., "1,250.00").
fn has_numeric_comma(s: &str) -> bool {
    let trimmed = s.trim_start_matches('-');
    // Only flag commas that look like thousands separators in numbers
    trimmed.contains(',')
        && trimmed
            .chars()
            .all(|c| c.is_ascii_digit() || c == ',' || c == '.' || c == '-')
}

/// Parse a single CSV field with strict financial rules.
///
/// - Integer: `-?\d+` → `StrictValue::Integer`
/// - Strict decimal: `-?\d+\.\d{2}` → `StrictValue::Decimal`
/// - Reject currency symbols in numeric-looking fields
/// - Reject commas in numeric-looking fields
/// - Reject wrong decimal places (`1250.5`, `1250.500`)
/// - Everything else → `StrictValue::Text` (no auto-detect)
/// - Values starting with `=` → error (formula injection)
pub fn parse_strict_value(s: &str, row_num: usize, col_num: usize) -> Result<StrictValue, CliError> {
    let trimmed = s.trim();

    if trimmed.is_empty() {
        return Ok(StrictValue::Empty);
    }

    // Reject formula injection
    if trimmed.starts_with('=') {
        return Err(CliError::parse(format!(
            "formula injection rejected at row {} col {}: {:?}",
            row_num + 1,
            col_num + 1,
            trimmed
        )));
    }

    // Check for currency symbols in numeric-looking fields
    if has_currency_symbol(trimmed) {
        // Only error if it looks numeric (has digits)
        if trimmed.chars().any(|c| c.is_ascii_digit()) {
            return Err(CliError::parse(format!(
                "currency symbol in numeric field at row {} col {}: {:?} (strip in adapter)",
                row_num + 1,
                col_num + 1,
                trimmed
            )));
        }
    }

    // Check for commas in numeric-looking fields
    if has_numeric_comma(trimmed) {
        return Err(CliError::parse(format!(
            "comma in numeric field at row {} col {}: {:?} (use plain digits)",
            row_num + 1,
            col_num + 1,
            trimmed
        )));
    }

    // Try integer: -?\d+
    let is_integer = {
        let t = trimmed.strip_prefix('-').unwrap_or(trimmed);
        !t.is_empty() && t.chars().all(|c| c.is_ascii_digit())
    };
    if is_integer {
        match trimmed.parse::<f64>() {
            Ok(n) => return Ok(StrictValue::Integer(n)),
            Err(_) => return Ok(StrictValue::Text(trimmed.to_string())),
        }
    }

    // Try strict decimal: -?\d+\.\d{2}
    if let Some(dot_pos) = trimmed.rfind('.') {
        let before = &trimmed[..dot_pos];
        let after = &trimmed[dot_pos + 1..];

        let before_ok = {
            let b = before.strip_prefix('-').unwrap_or(before);
            !b.is_empty() && b.chars().all(|c| c.is_ascii_digit())
        };
        let after_ok = !after.is_empty() && after.chars().all(|c| c.is_ascii_digit());

        if before_ok && after_ok {
            // It's a numeric decimal — enforce exactly 2 decimal places
            if after.len() != 2 {
                return Err(CliError::parse(format!(
                    "wrong decimal places at row {} col {}: {:?} (expected exactly 2)",
                    row_num + 1,
                    col_num + 1,
                    trimmed
                )));
            }
            match trimmed.parse::<f64>() {
                Ok(n) => return Ok(StrictValue::Decimal(n)),
                Err(_) => return Ok(StrictValue::Text(trimmed.to_string())),
            }
        }
    }

    // Everything else is text
    Ok(StrictValue::Text(trimmed.to_string()))
}

// ── Strict CSV parsing ──────────────────────────────────────────────

/// A row of strictly-parsed values, plus line context for error reporting.
pub struct StrictRow {
    pub values: Vec<StrictValue>,
}

/// Parse an entire CSV file with strict financial rules.
///
/// Uses `visigrid_io::csv::read_file_as_utf8` for encoding handling,
/// then applies strict value parsing.
pub fn parse_strict_csv(
    path: &Path,
    delimiter: char,
    has_headers: bool,
) -> Result<Vec<StrictRow>, CliError> {
    let content = visigrid_io::csv::read_file_as_utf8(path)
        .map_err(|e| CliError::io(format!("{}: {}", path.display(), e)))?;

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter as u8)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    let mut rows = Vec::new();

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| {
            CliError::parse(format!("CSV parse error at row {}: {}", row_idx + 1, e))
        })?;

        if row_idx == 0 && has_headers {
            // Skip header row — template already has its own headers
            continue;
        }

        let values: Result<Vec<StrictValue>, CliError> = record
            .iter()
            .enumerate()
            .map(|(col_idx, field)| parse_strict_value(field, row_idx, col_idx))
            .collect();

        rows.push(StrictRow { values: values? });
    }

    Ok(rows)
}

// ── Fill execution ──────────────────────────────────────────────────

/// Convert a StrictValue to the string representation for `set_cell_value_tracked`.
fn strict_value_to_string(v: &StrictValue) -> Option<String> {
    match v {
        StrictValue::Empty => None,
        StrictValue::Integer(n) => Some(format!("{}", *n as i64)),
        StrictValue::Decimal(n) => Some(format!("{:.2}", n)),
        StrictValue::Text(s) => Some(s.clone()),
    }
}

/// Execute `vgrid fill`.
pub fn cmd_fill(
    template: std::path::PathBuf,
    csv_path: std::path::PathBuf,
    target: String,
    headers: bool,
    clear: bool,
    out: std::path::PathBuf,
    delimiter: char,
    json: bool,
) -> Result<(), CliError> {
    use visigrid_io::native;

    // 1. Parse target
    let fill_target = parse_fill_target(&target)?;

    // 2. Load workbook
    if !template.exists() {
        return Err(CliError::io(format!(
            "template not found: {}",
            template.display()
        )));
    }
    let mut workbook = native::load_workbook(&template)
        .map_err(|e| CliError::io(format!("failed to load template: {}", e)))?;

    // 3. Rebuild dependency graph
    workbook.rebuild_dep_graph();

    // 4. Resolve sheet index
    let sheet_idx = if let Some(ref name) = fill_target.sheet_name {
        let sheet_id = workbook.sheet_id_by_name(name).ok_or_else(|| {
            CliError::args(format!("sheet {:?} not found in workbook", name))
        })?;
        workbook.idx_for_sheet_id(sheet_id).ok_or_else(|| {
            CliError::args(format!("sheet {:?} not found in workbook", name))
        })?
    } else {
        0
    };

    // 5. Parse CSV (fail-fast before modifying workbook)
    if !csv_path.exists() {
        return Err(CliError::io(format!(
            "CSV file not found: {}",
            csv_path.display()
        )));
    }
    let csv_rows = parse_strict_csv(&csv_path, delimiter, headers)?;

    if csv_rows.is_empty() {
        return Err(CliError::parse("CSV file is empty"));
    }

    // 6. Begin batch (defer recalc)
    workbook.begin_batch();

    // 7. Clear sheet if requested
    if clear {
        clear_sheet(&mut workbook, sheet_idx);
    }

    // 8. Fill cells
    let base_row = fill_target.row;
    let base_col = fill_target.col;
    let mut cells_set: usize = 0;

    for (row_offset, csv_row) in csv_rows.iter().enumerate() {
        for (col_offset, value) in csv_row.values.iter().enumerate() {
            if let Some(s) = strict_value_to_string(value) {
                workbook.set_cell_value_tracked(
                    sheet_idx,
                    base_row + row_offset,
                    base_col + col_offset,
                    &s,
                );
                cells_set += 1;
            }
        }
    }

    // 9. End batch (incremental recalc)
    workbook.end_batch();

    // 10. Full recalc (belt-and-suspenders)
    workbook.recompute_full_ordered();

    // 11. Compute fingerprint
    let fingerprint = native::compute_semantic_fingerprint(&workbook);

    // 12. Save workbook (atomic: write .tmp then rename)
    let tmp_path = out.with_extension("sheet.tmp");
    native::save_workbook(&workbook, &tmp_path)
        .map_err(|e| CliError::io(format!("failed to save: {}", e)))?;
    std::fs::rename(&tmp_path, &out)
        .map_err(|e| CliError::io(format!("failed to rename tmp to output: {}", e)))?;

    // 13. Print result
    let row_count = csv_rows.len();
    let col_count = csv_rows.iter().map(|r| r.values.len()).max().unwrap_or(0);

    if json {
        let result = serde_json::json!({
            "status": "ok",
            "cells_set": cells_set,
            "rows": row_count,
            "cols": col_count,
            "fingerprint": fingerprint,
            "output": out.display().to_string(),
        });
        println!("{}", serde_json::to_string(&result).unwrap());
    } else {
        eprintln!("Filled {} cells ({} rows x {} cols)", cells_set, row_count, col_count);
        eprintln!("Fingerprint: {}", fingerprint);
        eprintln!("Output: {}", out.display());
    }

    Ok(())
}

/// Clear all cells on a specific sheet.
fn clear_sheet(workbook: &mut visigrid_engine::workbook::Workbook, sheet_idx: usize) {
    let positions: Vec<(usize, usize)> = if let Some(sheet) = workbook.sheet(sheet_idx) {
        sheet.cells_iter().map(|(&(r, c), _)| (r, c)).collect()
    } else {
        return;
    };
    for (row, col) in positions {
        workbook.clear_cell_tracked(sheet_idx, row, col);
    }
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    // Target parsing tests

    #[test]
    fn test_target_parsing_sheet_cell() {
        let t = parse_fill_target("tx!A1").unwrap();
        assert_eq!(
            t,
            FillTarget {
                sheet_name: Some("tx".to_string()),
                row: 0,
                col: 0,
            }
        );
    }

    #[test]
    fn test_target_parsing_cell_only() {
        let t = parse_fill_target("B5").unwrap();
        assert_eq!(
            t,
            FillTarget {
                sheet_name: None,
                row: 4,
                col: 1,
            }
        );
    }

    #[test]
    fn test_target_parsing_empty_sheet() {
        let err = parse_fill_target("!A1");
        assert!(err.is_err());
    }

    // Strict value parsing tests

    #[test]
    fn test_strict_integer() {
        assert_eq!(parse_strict_value("42", 0, 0).unwrap(), StrictValue::Integer(42.0));
        assert_eq!(parse_strict_value("-7", 0, 0).unwrap(), StrictValue::Integer(-7.0));
        assert_eq!(parse_strict_value("0", 0, 0).unwrap(), StrictValue::Integer(0.0));
    }

    #[test]
    fn test_strict_decimal() {
        assert_eq!(parse_strict_value("1250.50", 0, 0).unwrap(), StrictValue::Decimal(1250.50));
        assert_eq!(parse_strict_value("-100.00", 0, 0).unwrap(), StrictValue::Decimal(-100.00));
    }

    #[test]
    fn test_reject_wrong_decimal_places() {
        assert!(parse_strict_value("1250.5", 0, 0).is_err());
        assert!(parse_strict_value("1250.500", 0, 0).is_err());
        assert!(parse_strict_value("100.1", 0, 0).is_err());
    }

    #[test]
    fn test_reject_currency_symbol() {
        assert!(parse_strict_value("$100.00", 0, 0).is_err());
        assert!(parse_strict_value("€50.00", 0, 0).is_err());
        assert!(parse_strict_value("£30.00", 0, 0).is_err());
    }

    #[test]
    fn test_reject_comma_in_number() {
        assert!(parse_strict_value("1,250.00", 0, 0).is_err());
        assert!(parse_strict_value("1,000", 0, 0).is_err());
    }

    #[test]
    fn test_text_passthrough() {
        assert_eq!(
            parse_strict_value("hello", 0, 0).unwrap(),
            StrictValue::Text("hello".to_string())
        );
        assert_eq!(
            parse_strict_value("2026-01-01", 0, 0).unwrap(),
            StrictValue::Text("2026-01-01".to_string())
        );
    }

    #[test]
    fn test_reject_formula_injection() {
        assert!(parse_strict_value("=SUM(A1:A10)", 0, 0).is_err());
        assert!(parse_strict_value("=1+1", 0, 0).is_err());
    }

    #[test]
    fn test_empty_value() {
        assert_eq!(parse_strict_value("", 0, 0).unwrap(), StrictValue::Empty);
        assert_eq!(parse_strict_value("  ", 0, 0).unwrap(), StrictValue::Empty);
    }
}
