use std::path::Path;

use crate::util;

pub struct PeekData {
    /// Row-major cell data (already display-ready strings)
    pub rows: Vec<Vec<String>>,
    /// Typed JSON cells when the source carries types (Parquet). Built only
    /// for --json so interactive previews don't keep a second copy of values.
    pub json_rows: Option<Vec<Vec<serde_json::Value>>>,
    /// Raw cell content (formulas as "=...", values as-is). Native workbooks only.
    pub raw: Option<Vec<Vec<String>>>,
    pub num_rows: usize,
    pub num_cols: usize,
    /// Pre-computed column widths (display columns, clamped to [3, 40])
    pub col_widths: Vec<usize>,
    /// Column names: schema, first-row values (--headers), or generated A,B,C...
    pub col_names: Vec<String>,
    /// Whether explicit headers are present (schema or a consumed header row)
    pub has_headers: bool,
    /// 1-based file row number of the first data row (2 if headers consumed, else 1)
    pub first_data_file_row: usize,
    /// Total data row count in file (if known, even when truncated by --max-rows)
    pub total_rows: Option<usize>,
    /// Detected delimiter, or 0 for structured files
    pub delimiter: u8,
}

impl PeekData {
    /// File row number for data row at index `i` (1-based, accounts for header row).
    pub fn file_row(&self, i: usize) -> usize {
        self.first_data_file_row + i
    }

    /// Total data rows (actual count if truncated, else loaded count).
    pub fn total_data_rows(&self) -> usize {
        self.total_rows.unwrap_or(self.num_rows)
    }

    /// Compute column widths by scanning up to `scan_rows` data rows (0 = all).
    /// Always includes the header names in the scan.
    pub(crate) fn compute_widths(col_names: &[String], rows: &[Vec<String>], num_cols: usize, scan_rows: usize) -> Vec<usize> {
        let scan_limit = if scan_rows == 0 { rows.len() } else { scan_rows.min(rows.len()) };
        (0..num_cols)
            .map(|c| {
                let header_w = col_names.get(c).map(|s| util::display_width(s)).unwrap_or(0);
                let max_cell = rows[..scan_limit]
                    .iter()
                    .map(|row| row.get(c).map(|s| util::display_width(s)).unwrap_or(0))
                    .max()
                    .unwrap_or(0);
                header_w.max(max_cell).clamp(3, 40)
            })
            .collect()
    }
}

/// Load a CSV or TSV file into PeekData.
///
/// `delimiter` is b',' for CSV or b'\t' for TSV.
/// `headers`: `Some(true)` makes the first row column names, `Some(false)`
/// keeps it as data, `None` decides with [`looks_like_header`].
/// `max_rows`: cap on data rows loaded (0 = unlimited).
/// `width_scan_rows`: how many rows to scan for column width (0 = all loaded rows).
pub fn load_csv(
    path: &Path,
    delimiter: u8,
    headers: Option<bool>,
    max_rows: usize,
    width_scan_rows: usize,
) -> Result<PeekData, String> {
    let content = visigrid_io::csv::read_file_as_utf8(path)
        .map_err(|e| format!("failed to open {}: {}", path.display(), e))?;
    let mut rdr = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    let mut all_rows: Vec<Vec<String>> = Vec::new();
    let mut max_cols: usize = 0;
    let mut total_count: usize = 0;
    let cap = if max_rows == 0 { usize::MAX } else { max_rows };
    // Read one extra row: it is the header row, or (when there is none) the
    // sign that the preview is truncated.
    let row_limit = cap.saturating_add(1);
    let mut capped = false;

    for result in rdr.records() {
        let record = result.map_err(|e| format!("CSV parse error: {}", e))?;
        total_count += 1;
        if all_rows.len() < row_limit {
            let row: Vec<String> = record.iter().map(|s| s.to_string()).collect();
            if row.len() > max_cols {
                max_cols = row.len();
            }
            all_rows.push(row);
        } else {
            capped = true;
        }
    }

    let has_headers = headers.unwrap_or_else(|| looks_like_header(&all_rows));
    if !has_headers && all_rows.len() > cap {
        all_rows.truncate(cap);
        capped = true;
    }

    // Extract header row if requested
    let col_names: Vec<String>;
    let first_data_file_row: usize;
    let data_rows;
    if has_headers && !all_rows.is_empty() {
        let header_row = all_rows.remove(0);
        col_names = (0..max_cols)
            .map(|i| {
                header_row
                    .get(i)
                    .filter(|s| !s.is_empty())
                    .cloned()
                    .unwrap_or_else(|| util::col_to_letter(i))
            })
            .collect();
        first_data_file_row = 2; // header was file row 1, data starts at 2
        data_rows = all_rows;
    } else {
        col_names = (0..max_cols).map(util::col_to_letter).collect();
        first_data_file_row = 1;
        data_rows = all_rows;
    }

    // Pad short rows so all have max_cols entries
    let mut rows = data_rows;
    for row in &mut rows {
        row.resize(max_cols, String::new());
    }

    let num_rows = rows.len();
    let num_cols = max_cols;

    let col_widths = PeekData::compute_widths(&col_names, &rows, num_cols, width_scan_rows);

    let total_rows = if capped {
        Some(total_count - if has_headers { 1 } else { 0 })
    } else {
        None
    };

    Ok(PeekData {
        rows,
        json_rows: None,
        raw: None,
        num_rows,
        num_cols,
        col_widths,
        col_names,
        has_headers,
        first_data_file_row,
        total_rows,
        delimiter,
    })
}

/// Does the first row name the columns? Biased toward yes, as most delimited
/// files have a header (VisiData assumes one). The first row is data when it
/// looks like data: an empty or numeric/date-like cell, two cells with the same
/// text, or a value that recurs further down its own column.
pub fn looks_like_header(rows: &[Vec<String>]) -> bool {
    let Some((first, rest)) = rows.split_first() else { return false };
    if rest.is_empty() || first.is_empty() {
        return false;
    }
    let mut seen = std::collections::HashSet::new();
    for cell in first {
        let t = cell.trim();
        if t.is_empty() || !seen.insert(t.to_lowercase()) {
            return false;
        }
        let lead = t.trim_start_matches(['-', '+', '$', '(', '€', '£']);
        if lead.starts_with(|c: char| c.is_ascii_digit()) || matches!(t.to_lowercase().as_str(), "true" | "false") {
            return false;
        }
    }
    let sample = &rest[..rest.len().min(200)];
    !first.iter().enumerate().any(|(c, cell)| {
        let t = cell.trim();
        sample.iter().any(|row| row.get(c).is_some_and(|v| v.trim().eq_ignore_ascii_case(t)))
    })
}

/// A sheet within a workbook, extracted as display-ready data.
pub struct SheetData {
    pub name: String,
    pub data: PeekData,
}

/// Bounded Parquet preview. Schema names are always headers; record 1 is the
/// first data row, because the file has no physical header record to skip.
pub fn load_parquet(
    path: &Path,
    max_rows: usize,
    force: bool,
    width_scan_rows: usize,
    json: bool,
) -> Result<PeekData, String> {
    let limit = if max_rows == 0 && !force {
        crate::PEEK_FORCE_CAP + 1
    } else if max_rows == 0 {
        usize::MAX
    } else {
        max_rows
    };
    let imported = visigrid_io::parquet::import_with_limits(
        path,
        limit,
        if force { None } else { Some(PEEK_CELL_CAP) },
    )?;
    if max_rows == 0 && !force && imported.total_rows > crate::PEEK_FORCE_CAP as u64 {
        return Err("Parquet file has >200k rows; use --max-rows to preview fewer rows or --force to override".into());
    }
    if imported.cols_loaded < imported.total_cols {
        return Err(format!(
            "Parquet file has {} columns; peek supports at most {} (select fewer columns first)",
            imported.total_cols, imported.cols_loaded,
        ));
    }
    let total = usize::try_from(imported.total_rows)
        .map_err(|_| "Parquet row count exceeds this platform's capacity")?;
    let num_rows = imported.rows_loaded;
    let num_cols = imported.cols_loaded;
    let col_names: Vec<String> = (0..num_cols)
        .map(|c| imported.sheet.get_display(0, c))
        .collect();
    let rows: Vec<Vec<String>> = (1..=num_rows)
        .map(|r| (0..num_cols).map(|c| imported.sheet.get_interchange_display(r, c)).collect())
        .collect();
    let json_rows = if json {
        Some((1..=num_rows).map(|r| {
            (0..num_cols).map(|c| {
                if imported.sheet.get_cell_opt(r, c).is_none_or(|cell| cell.value().is_empty()) {
                    serde_json::Value::Null
                } else {
                    let value = crate::convert::cell_json_value(&imported.sheet, r, c);
                    // Temporal values are numbers in the engine, but their
                    // interchange representation is an ISO date/time string.
                    if value.is_number() {
                        crate::convert::string_to_json_value(&rows[r - 1][c])
                    } else {
                        value
                    }
                }
            }).collect()
        }).collect())
    } else {
        None
    };
    let col_widths = PeekData::compute_widths(&col_names, &rows, num_cols, width_scan_rows);
    Ok(PeekData {
        rows,
        json_rows,
        raw: None,
        num_rows,
        num_cols,
        col_widths,
        col_names,
        has_headers: true,
        first_data_file_row: 1,
        total_rows: (num_rows < total).then_some(total),
        delimiter: 0,
    })
}

/// Load a native workbook file and return a PeekData for each sheet.
///
/// Loads the workbook, rebuilds the dependency graph, recomputes all formulas,
/// then extracts evaluated cell values as display strings.
/// `max_rows`: cap on data rows loaded per sheet (0 = unlimited).
pub fn load_sheet(
    path: &Path,
    max_rows: usize,
    width_scan_rows: usize,
    force: bool,
) -> Result<Vec<SheetData>, String> {
    let mut workbook = visigrid_io::native::load_workbook(path)
        .map_err(|e| format!("failed to load {}: {}", path.display(), e))?;

    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    let sheet_count = workbook.sheet_count();
    if sheet_count == 0 {
        return Ok(vec![]);
    }

    let cap = if max_rows == 0 { usize::MAX } else { max_rows };
    let mut sheets = Vec::with_capacity(sheet_count);

    for idx in 0..sheet_count {
        let sheet = workbook.sheet(idx).unwrap();
        let name = sheet.name.clone();

        // Find the bounding box of non-empty cells
        let mut max_row: usize = 0;
        let mut max_col: usize = 0;
        let mut has_cells = false;
        for ((r, c), _) in sheet.cells_iter() {
            has_cells = true;
            if r > max_row {
                max_row = r;
            }
            if c > max_col {
                max_col = c;
            }
        }
        // Also check spill values (array formula results)
        // They may extend beyond the cell map

        if !has_cells {
            // Empty sheet — still include it with zero rows/cols
            sheets.push(SheetData {
                name,
                data: PeekData {
                    rows: vec![],
                    json_rows: None,
                    raw: Some(vec![]),
                    num_rows: 0,
                    num_cols: 0,
                    col_widths: vec![],
                    col_names: vec![],
                    has_headers: false,
                    first_data_file_row: 1,
                    total_rows: None,
                    delimiter: 0,
                },
            });
            continue;
        }

        let total_rows_in_sheet = max_row + 1;
        let num_cols = max_col + 1;
        let effective_rows = total_rows_in_sheet.min(cap);

        check_preview_size(&name, effective_rows, num_cols, force)?;

        // Extract evaluated values and raw formulas into row-major grids
        let mut rows: Vec<Vec<String>> = Vec::with_capacity(effective_rows);
        let mut raw_rows: Vec<Vec<String>> = Vec::with_capacity(effective_rows);
        for r in 0..effective_rows {
            let mut row = Vec::with_capacity(num_cols);
            let mut raw_row = Vec::with_capacity(num_cols);
            for c in 0..num_cols {
                row.push(sheet.get_display(r, c));
                raw_row.push(sheet.get_raw(r, c));
            }
            rows.push(row);
            raw_rows.push(raw_row);
        }

        let num_rows = rows.len();

        // Generate column names (A, B, C, ...)
        let col_names: Vec<String> = (0..num_cols).map(util::col_to_letter).collect();

        let col_widths = PeekData::compute_widths(&col_names, &rows, num_cols, width_scan_rows);

        let total_rows = if effective_rows < total_rows_in_sheet {
            Some(total_rows_in_sheet)
        } else {
            None
        };

        sheets.push(SheetData {
            name,
            data: PeekData {
                rows,
                json_rows: None,
                raw: Some(raw_rows),
                num_rows,
                num_cols,
                col_widths,
                col_names,
                has_headers: false,
                first_data_file_row: 1,
                total_rows,
                delimiter: 0,
            },
        });
    }

    Ok(sheets)
}

/// Maximum cell count (rows * cols) for a single sheet before requiring --force.
/// Prevents catastrophic allocation from xlsx/ods "used range" formatting artifacts.
const PEEK_CELL_CAP: usize = 10_000_000;

fn check_preview_size(name: &str, rows: usize, cols: usize, force: bool) -> Result<(), String> {
    if !force && rows.saturating_mul(cols) > PEEK_CELL_CAP {
        return Err(format!(
            "sheet '{}' preview has {} rows x {} cols; exceeds {} cells\n\
             hint: reduce --max-rows or use --force to override",
            name, rows, cols, PEEK_CELL_CAP,
        ));
    }
    Ok(())
}

/// Load an Excel/ODS workbook file and return a PeekData for each sheet.
///
/// Uses `visigrid_io::xlsx::import()` (calamine) to read stored cell values.
/// By default does NOT recompute formulas — shows calamine's cached values (fast).
/// When `recompute` is true, rebuilds the dependency graph and recomputes all formulas.
/// `max_rows`: cap on data rows per sheet (0 = unlimited).
/// `force`: if true, skip the cell-count guard for oversized bounding boxes.
pub fn load_workbook_peek(
    path: &Path,
    max_rows: usize,
    width_scan_rows: usize,
    recompute: bool,
    force: bool,
) -> Result<Vec<SheetData>, String> {
    let (mut workbook, _result) = visigrid_io::xlsx::import(path)
        .map_err(|e| format!("failed to import {}: {}", path.display(), e))?;

    if recompute {
        workbook.rebuild_dep_graph();
        workbook.recompute_full_ordered();
    }

    let sheet_count = workbook.sheet_count();
    if sheet_count == 0 {
        return Ok(vec![]);
    }

    let cap = if max_rows == 0 { usize::MAX } else { max_rows };
    let mut sheets = Vec::with_capacity(sheet_count);

    for idx in 0..sheet_count {
        let sheet = workbook.sheet(idx).unwrap();
        let name = sheet.name.clone();

        // Find the bounding box based on non-empty cell values only.
        // XLSX/ODS can have formatting-only cells that inflate the "used range"
        // far beyond actual data. Filtering by non-empty values prevents
        // catastrophic allocation from template artifacts.
        let mut max_row: usize = 0;
        let mut max_col: usize = 0;
        let mut has_cells = false;
        for ((r, c), cell) in sheet.cells_iter() {
            if cell.value().raw_display().is_empty() {
                continue;
            }
            has_cells = true;
            if r > max_row {
                max_row = r;
            }
            if c > max_col {
                max_col = c;
            }
        }

        if !has_cells {
            sheets.push(SheetData {
                name,
                data: PeekData {
                    rows: vec![],
                    json_rows: None,
                    raw: None,
                    num_rows: 0,
                    num_cols: 0,
                    col_widths: vec![],
                    col_names: vec![],
                    has_headers: false,
                    first_data_file_row: 1,
                    total_rows: None,
                    delimiter: 0,
                },
            });
            continue;
        }

        let total_rows_in_sheet = max_row + 1;
        let num_cols = max_col + 1;

        let effective_rows = total_rows_in_sheet.min(cap);
        check_preview_size(&name, effective_rows, num_cols, force)?;

        // Extract display values into row-major grid
        let mut rows: Vec<Vec<String>> = Vec::with_capacity(effective_rows);
        for r in 0..effective_rows {
            let mut row = Vec::with_capacity(num_cols);
            for c in 0..num_cols {
                row.push(sheet.get_display(r, c));
            }
            rows.push(row);
        }

        let num_rows = rows.len();
        let col_names: Vec<String> = (0..num_cols).map(util::col_to_letter).collect();
        let col_widths = PeekData::compute_widths(&col_names, &rows, num_cols, width_scan_rows);

        let total_rows = if effective_rows < total_rows_in_sheet {
            Some(total_rows_in_sheet)
        } else {
            None
        };

        sheets.push(SheetData {
            name,
            data: PeekData {
                rows,
                json_rows: None,
                raw: None,
                num_rows,
                num_cols,
                col_widths,
                col_names,
                has_headers: false,
                first_data_file_row: 1,
                total_rows,
                delimiter: 0,
            },
        });
    }

    Ok(sheets)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn write_csv(content: &str) -> tempfile::NamedTempFile {
        let mut f = tempfile::NamedTempFile::new().unwrap();
        f.write_all(content.as_bytes()).unwrap();
        f.flush().unwrap();
        f
    }

    #[test]
    fn preview_cell_budget_can_be_overridden() {
        assert!(check_preview_size("wide", 1001, 10000, false).is_err());
        assert!(check_preview_size("wide", 1, 10000, false).is_ok());
        assert!(check_preview_size("wide", 1001, 10000, true).is_ok());
    }

    #[test]
    fn ragged_rows_padded() {
        let f = write_csv("a,b,c\n1,2\n3\n");
        let data = load_csv(f.path(), b',', Some(false), 0, 0).unwrap();
        assert_eq!(data.num_cols, 3);
        assert_eq!(data.num_rows, 3);
        // Short rows should be padded with empty strings
        assert_eq!(data.rows[1], vec!["1", "2", ""]);
        assert_eq!(data.rows[2], vec!["3", "", ""]);
    }

    #[test]
    fn headers_consumed_file_row_mapping() {
        let f = write_csv("Name,Value\nAlice,100\nBob,200\n");
        let data = load_csv(f.path(), b',', Some(true), 0, 0).unwrap();
        assert!(data.has_headers);
        assert_eq!(data.num_rows, 2); // header consumed, 2 data rows
        assert_eq!(data.col_names, vec!["Name", "Value"]);
        assert_eq!(data.first_data_file_row, 2);
        // Data row 0 is file row 2, data row 1 is file row 3
        assert_eq!(data.file_row(0), 2);
        assert_eq!(data.file_row(1), 3);
    }

    #[test]
    fn no_headers_file_row_mapping() {
        let f = write_csv("1,2\n3,4\n");
        let data = load_csv(f.path(), b',', Some(false), 0, 0).unwrap();
        assert!(!data.has_headers);
        assert_eq!(data.first_data_file_row, 1);
        assert_eq!(data.file_row(0), 1);
        assert_eq!(data.col_names, vec!["A", "B"]);
    }

    #[test]
    fn max_rows_cap() {
        let mut csv = String::from("h1,h2\n");
        for i in 0..100 {
            csv.push_str(&format!("{},{}\n", i, i * 10));
        }
        let f = write_csv(&csv);
        let data = load_csv(f.path(), b',', Some(true), 10, 0).unwrap();
        assert_eq!(data.num_rows, 10);
        assert!(data.total_rows.is_some());
        assert_eq!(data.total_data_rows(), 100);
    }

    #[test]
    fn tsv_delimiter() {
        let f = write_csv("a\tb\tc\n1\t2\t3\n");
        let data = load_csv(f.path(), b'\t', Some(false), 0, 0).unwrap();
        assert_eq!(data.num_cols, 3);
        assert_eq!(data.rows[0], vec!["a", "b", "c"]);
        assert_eq!(data.delimiter, b'\t');
    }

    #[test]
    fn width_scan_rows_limits_scan() {
        // First 2 rows have short values, row 3 has a long value
        let f = write_csv("a,b\n1,2\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx,y\n");
        // Scan only 1 row — should not see the long value
        let data_limited = load_csv(f.path(), b',', Some(false), 0, 1).unwrap();
        // Scan all — should see the long value (clamped to 40)
        let data_all = load_csv(f.path(), b',', Some(false), 0, 0).unwrap();
        assert!(data_limited.col_widths[0] < data_all.col_widths[0]);
        assert_eq!(data_all.col_widths[0], 40); // clamped
    }

    #[test]
    fn empty_csv() {
        let f = write_csv("");
        let data = load_csv(f.path(), b',', Some(false), 0, 0).unwrap();
        assert_eq!(data.num_rows, 0);
        assert_eq!(data.num_cols, 0);
    }

    #[test]
    fn header_detection() {
        let rows = |v: &[&[&str]]| v.iter().map(|r| r.iter().map(|s| s.to_string()).collect()).collect::<Vec<Vec<String>>>();
        // Typical header over mixed data, and over all-text data.
        assert!(looks_like_header(&rows(&[&["Region", "Amount"], &["West", "10"], &["East", "5"]])));
        assert!(looks_like_header(&rows(&[&["Name", "City"], &["Ann", "Paris"], &["Bo", "Rome"]])));
        // First row that looks like data.
        assert!(!looks_like_header(&rows(&[&["West", "10"], &["East", "5"]])), "numeric cell");
        assert!(!looks_like_header(&rows(&[&["2026-01-02", "x"], &["2026-01-03", "y"]])), "date cell");
        assert!(!looks_like_header(&rows(&[&["West", ""], &["East", "a"]])), "empty cell");
        assert!(!looks_like_header(&rows(&[&["West", "West"], &["East", "a"]])), "duplicate cells");
        assert!(!looks_like_header(&rows(&[&["West", "Ann"], &["East", "Bo"], &["West", "Cy"]])), "value recurs below");
        assert!(!looks_like_header(&rows(&[&["only", "row"]])), "a lone row is data");
        assert!(!looks_like_header(&rows(&[&["TRUE", "x"], &["FALSE", "y"]])), "boolean cell");
    }

    #[test]
    fn detected_headers_keep_the_row_cap_exact() {
        let f = write_csv("Region,Amount\nWest,10\nEast,5\nNorth,7\n");
        let d = load_csv(f.path(), b',', None, 2, 0).unwrap();
        assert!(d.has_headers);
        assert_eq!(d.col_names, vec!["Region", "Amount"]);
        assert_eq!((d.num_rows, d.total_rows), (2, Some(3)));

        let f = write_csv("West,10\nEast,5\nNorth,7\n");
        let d = load_csv(f.path(), b',', None, 2, 0).unwrap();
        assert!(!d.has_headers);
        assert_eq!((d.num_rows, d.total_rows), (2, Some(3)));
        let d = load_csv(f.path(), b',', None, 0, 0).unwrap();
        assert_eq!((d.num_rows, d.total_rows), (3, None));
    }
}
