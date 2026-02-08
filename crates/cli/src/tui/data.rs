use std::path::Path;

use crate::util;

pub struct PeekData {
    /// Row-major cell data (already display-ready strings)
    pub rows: Vec<Vec<String>>,
    pub num_rows: usize,
    pub num_cols: usize,
    /// Pre-computed column widths (display columns, clamped to [3, 40])
    pub col_widths: Vec<usize>,
    /// Column header names: first-row values (--headers) or generated A,B,C...
    pub col_names: Vec<String>,
    /// Whether first data row was consumed as headers
    pub has_headers: bool,
    /// 1-based file row number of the first data row (2 if headers consumed, else 1)
    pub first_data_file_row: usize,
    /// Total data row count in file (if known, even when truncated by --max-rows)
    pub total_rows: Option<usize>,
    /// Detected delimiter
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
    fn compute_widths(col_names: &[String], rows: &[Vec<String>], num_cols: usize, scan_rows: usize) -> Vec<usize> {
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
/// `has_headers`: if true, first row becomes column names instead of data.
/// `max_rows`: cap on data rows loaded (0 = unlimited).
/// `width_scan_rows`: how many rows to scan for column width (0 = all loaded rows).
pub fn load_csv(
    path: &Path,
    delimiter: u8,
    has_headers: bool,
    max_rows: usize,
    width_scan_rows: usize,
) -> Result<PeekData, String> {
    let mut rdr = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_path(path)
        .map_err(|e| format!("failed to open {}: {}", path.display(), e))?;

    let mut all_rows: Vec<Vec<String>> = Vec::new();
    let mut max_cols: usize = 0;
    let mut total_count: usize = 0;
    let cap = if max_rows == 0 { usize::MAX } else { max_rows };
    // If using headers, we need one extra row for the header row itself
    let row_limit = if has_headers { cap.saturating_add(1) } else { cap };
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
        col_names = (0..max_cols).map(|i| util::col_to_letter(i)).collect();
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
    fn ragged_rows_padded() {
        let f = write_csv("a,b,c\n1,2\n3\n");
        let data = load_csv(f.path(), b',', false, 0, 0).unwrap();
        assert_eq!(data.num_cols, 3);
        assert_eq!(data.num_rows, 3);
        // Short rows should be padded with empty strings
        assert_eq!(data.rows[1], vec!["1", "2", ""]);
        assert_eq!(data.rows[2], vec!["3", "", ""]);
    }

    #[test]
    fn headers_consumed_file_row_mapping() {
        let f = write_csv("Name,Value\nAlice,100\nBob,200\n");
        let data = load_csv(f.path(), b',', true, 0, 0).unwrap();
        assert_eq!(data.has_headers, true);
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
        let data = load_csv(f.path(), b',', false, 0, 0).unwrap();
        assert_eq!(data.has_headers, false);
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
        let data = load_csv(f.path(), b',', true, 10, 0).unwrap();
        assert_eq!(data.num_rows, 10);
        assert!(data.total_rows.is_some());
        assert_eq!(data.total_data_rows(), 100);
    }

    #[test]
    fn tsv_delimiter() {
        let f = write_csv("a\tb\tc\n1\t2\t3\n");
        let data = load_csv(f.path(), b'\t', false, 0, 0).unwrap();
        assert_eq!(data.num_cols, 3);
        assert_eq!(data.rows[0], vec!["a", "b", "c"]);
        assert_eq!(data.delimiter, b'\t');
    }

    #[test]
    fn width_scan_rows_limits_scan() {
        // First 2 rows have short values, row 3 has a long value
        let f = write_csv("a,b\n1,2\nxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx,y\n");
        // Scan only 1 row — should not see the long value
        let data_limited = load_csv(f.path(), b',', false, 0, 1).unwrap();
        // Scan all — should see the long value (clamped to 40)
        let data_all = load_csv(f.path(), b',', false, 0, 0).unwrap();
        assert!(data_limited.col_widths[0] < data_all.col_widths[0]);
        assert_eq!(data_all.col_widths[0], 40); // clamped
    }

    #[test]
    fn empty_csv() {
        let f = write_csv("");
        let data = load_csv(f.path(), b',', false, 0, 0).unwrap();
        assert_eq!(data.num_rows, 0);
        assert_eq!(data.num_cols, 0);
    }
}
