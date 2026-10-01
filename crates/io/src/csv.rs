// CSV/TSV import/export

use std::path::Path;
use std::io::Read;

use visigrid_engine::sheet::{Sheet, SheetId};

pub fn import(path: &Path) -> Result<Sheet, String> {
    import_report(path, None, CsvOptions::default()).map(|r| r.sheet)
}

pub fn import_tsv(path: &Path) -> Result<Sheet, String> {
    import_report(path, Some(b'\t'), CsvOptions::default()).map(|r| r.sheet)
}

pub fn import_with_delimiter(path: &Path, delimiter: u8) -> Result<Sheet, String> {
    import_report(path, Some(delimiter), CsvOptions::default()).map(|r| r.sheet)
}

/// How a CSV is read.
#[derive(Debug, Clone, Copy, Default)]
pub struct CsvOptions {
    /// Run fields that start with `=` as formulas. Off by default: a CSV is
    /// data, and a formula in one is the classic injection vector (OWASP; a
    /// 2025 USENIX WOOT study found every spreadsheet tested would exfiltrate
    /// through one). `vgrid fill` already refuses them outright.
    pub evaluate_formulas: bool,
}

/// A CSV import, with what was kept as text and what did not fit.
#[derive(Debug)]
pub struct CsvImport {
    pub sheet: Sheet,
    /// Values kept as text because reading them as numbers would change them:
    /// leading zeros (007), more than 15 digits, or ID-like E-notation (12E4).
    pub kept_as_text: usize,
    /// Fields starting with `=` kept as text (see `CsvOptions::evaluate_formulas`).
    pub formulas_as_text: usize,
    /// Header labels (first row) of columns where values were kept as text.
    pub text_columns: Vec<String>,
    /// Records in the file, and records placed on the sheet.
    pub rows_in_file: usize,
    pub rows_loaded: usize,
}

impl CsvImport {
    /// One line for the status bar, or None when nothing needs saying.
    pub fn message(&self) -> Option<String> {
        let mut parts = Vec::new();
        if self.rows_loaded < self.rows_in_file {
            parts.push(format!(
                "Loaded the first {} of {} rows — the sheet holds {}; the rest were not imported",
                self.rows_loaded,
                self.rows_in_file,
                visigrid_engine::sheet::NUM_ROWS
            ));
        }
        if self.kept_as_text > 0 {
            let cols = match self.text_columns.len() {
                0 => String::new(),
                n if n <= 3 => format!(" ({})", self.text_columns.join(", ")),
                n => format!(" ({} and {} more)", self.text_columns[..3].join(", "), n - 3),
            };
            parts.push(format!(
                "Kept {} value{} as text so leading zeros and long IDs stay exact{}",
                self.kept_as_text,
                if self.kept_as_text == 1 { "" } else { "s" },
                cols
            ));
        }
        if self.formulas_as_text > 0 {
            parts.push(format!(
                "Left {} formula{} as text",
                self.formulas_as_text,
                if self.formulas_as_text == 1 { "" } else { "s" }
            ));
        }
        (!parts.is_empty()).then(|| parts.join(". ") + ".")
    }
}

/// Read a CSV, sniffing the delimiter when none is given.
pub fn import_report(path: &Path, delimiter: Option<u8>, options: CsvOptions) -> Result<CsvImport, String> {
    let content = read_file_as_utf8(path)?;
    let delimiter = delimiter.unwrap_or_else(|| sniff_delimiter(&content));
    import_from_string_with(&content, delimiter, options)
}

/// Why a field is stored as text instead of being typed.
#[derive(Debug, PartialEq)]
enum KeepText {
    /// Reading it as a number would change it.
    Exact,
    /// It starts with `=` and formulas are off.
    Formula,
}

/// Decide whether a CSV field must stay text. Everything else is typed as if it
/// had been entered by hand (numbers, 5%, $1,234, (500)) — and nothing becomes a
/// date, so gene names, part numbers and 1-2 stay as written.
fn keep_as_text(field: &str, options: CsvOptions) -> Option<KeepText> {
    let t = field.trim();
    if t.starts_with('=') {
        return (!options.evaluate_formulas).then_some(KeepText::Formula);
    }
    let digits = t.strip_prefix(['+', '-']).unwrap_or(t);
    if !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()) {
        // 007, 00123: a number would drop the zeros (ZIP codes, SKUs, accounts).
        // More than 15 digits: past what a double holds exactly (cards, barcodes).
        if (digits.len() > 1 && digits.starts_with('0')) || digits.len() > 15 {
            return Some(KeepText::Exact);
        }
        return None;
    }
    // 12E4, 2310009E13: digits, E, digits, nothing else. Programs write real
    // scientific notation with a sign or a point (1e-05, 1.5E+07); this shape
    // is an ID, and reading it as 120000 destroys it.
    if let Some((m, e)) = t.split_once(['e', 'E']) {
        let all_digits = |x: &str| !x.is_empty() && x.bytes().all(|b| b.is_ascii_digit());
        if all_digits(m) && all_digits(e) {
            return Some(KeepText::Exact);
        }
    }
    None
}

/// Parse delimited text into a 2D grid of strings.
///
/// Uses `sniff_delimiter` to auto-detect the delimiter (comma, semicolon, pipe, or tab),
/// then parses with the `csv` crate (handles quoted fields). Returns a `Vec<Vec<String>>` grid.
/// Useful for pasting CSV-formatted text from external sources.
pub fn parse_delimited_text(text: &str) -> Vec<Vec<String>> {
    let delimiter = sniff_delimiter(text);
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(text.as_bytes());

    let mut grid = Vec::new();
    for result in reader.records() {
        let record = match result {
            Ok(r) => r,
            Err(_) => continue,
        };
        let row: Vec<String> = record.iter().map(|f| f.to_string()).collect();
        grid.push(row);
    }
    grid
}

/// Detect the most likely field delimiter by checking consistency across the first few lines.
///
/// For each candidate (tab, semicolon, comma, pipe), count fields per line. The delimiter
/// that produces the most consistent field count (>1 field) wins.
pub fn sniff_delimiter(content: &str) -> u8 {
    let candidates: &[u8] = b"\t;,|";
    let sample_lines: Vec<&str> = content.lines().take(10).collect();

    if sample_lines.is_empty() {
        return b',';
    }

    let mut best = b',';
    let mut best_score = 0u64;

    for &delim in candidates {
        let counts: Vec<usize> = sample_lines
            .iter()
            .map(|line| {
                csv::ReaderBuilder::new()
                    .delimiter(delim)
                    .has_headers(false)
                    .flexible(true)
                    .from_reader(line.as_bytes())
                    .records()
                    .next()
                    .and_then(|r| r.ok())
                    .map(|r| r.len())
                    .unwrap_or(1)
            })
            .collect();

        // Must produce >1 field on the first line to be viable
        if counts.first().copied().unwrap_or(0) <= 1 {
            continue;
        }

        // Score: (number of lines with same field count as line 1) * field_count
        // Higher field count breaks ties — more columns = more likely real delimiter
        let target = counts[0];
        let consistent = counts.iter().filter(|&&c| c == target).count() as u64;
        let score = consistent * target as u64;

        if score > best_score {
            best_score = score;
            best = delim;
        }
    }

    best
}

/// Read file and convert to UTF-8 if needed (handles Windows-1252, Latin-1, etc.)
pub fn read_file_as_utf8(path: &Path) -> Result<String, String> {
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).map_err(|e| e.to_string())?;

    // Try UTF-8 first; on failure, recover the buffer from the error
    match String::from_utf8(bytes) {
        Ok(s) => Ok(s),
        Err(e) => {
            let bytes = e.into_bytes();
            // Fall back to Windows-1252 (common for Excel-exported CSVs)
            let (decoded, _, _) = encoding_rs::WINDOWS_1252.decode(&bytes);
            Ok(decoded.into_owned())
        }
    }
}

#[cfg(test)]
fn import_from_string(content: &str, delimiter: u8) -> Result<Sheet, String> {
    import_from_string_with(content, delimiter, CsvOptions::default()).map(|r| r.sheet)
}

fn import_from_string_with(content: &str, delimiter: u8, options: CsvOptions) -> Result<CsvImport, String> {
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    // Start with reasonable defaults, will track actual extent
    let mut sheet = Sheet::new(SheetId(1), visigrid_engine::sheet::NUM_ROWS, visigrid_engine::sheet::NUM_COLS);
    sheet.reserve_cells(estimate_cells(content, delimiter));
    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let (mut kept_as_text, mut formulas_as_text) = (0usize, 0usize);
    let mut header: Vec<String> = Vec::new();
    let mut text_cols: Vec<usize> = Vec::new();
    let mut rows_in_file = 0usize;

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| e.to_string())?;
        rows_in_file = row_idx + 1;
        // A sheet holds NUM_ROWS rows. Rows past it are counted, not stored —
        // the message says so, rather than the grid holding rows it cannot show.
        if row_idx >= visigrid_engine::sheet::NUM_ROWS {
            continue;
        }
        if row_idx == 0 {
            header = record.iter().map(|f| f.trim().to_string()).collect();
        }
        for (col_idx, field) in record.iter().enumerate() {
            if field.is_empty() {
                continue;
            }
            match keep_as_text(field, options) {
                Some(reason) => {
                    sheet.set_text(row_idx, col_idx, field);
                    match reason {
                        KeepText::Formula => formulas_as_text += 1,
                        KeepText::Exact => {
                            kept_as_text += 1;
                            if row_idx > 0 && !text_cols.contains(&col_idx) {
                                text_cols.push(col_idx);
                            }
                        }
                    }
                }
                None => sheet.set_value_deferred(row_idx, col_idx, field),
            }
            max_col = max_col.max(col_idx);
        }
        max_row = row_idx;
    }

    // Update sheet dimensions to actual data extent (for export efficiency)
    sheet.rows = (max_row + 1).max(1000);
    sheet.cols = (max_col + 1).max(26);

    // Evaluate in dependency order rather than in the order the rows happened
    // to be read. Cells were inserted without evaluating, and previously each
    // was evaluated on arrival: a formula naming a row further down the file
    // was computed against a cell that did not exist yet, came out as zero, and
    // nothing ever revisited it — so anything referring to that formula was
    // wrong too, including formulas pointing safely backwards.
    let mut wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

    let text_columns = text_cols
        .iter()
        .map(|&c| header.get(c).filter(|h| !h.is_empty()).cloned().unwrap_or_else(|| col_label(c)))
        .collect();
    Ok(CsvImport {
        sheet: wb.into_sheets().swap_remove(0),
        kept_as_text,
        formulas_as_text,
        text_columns,
        rows_in_file,
        rows_loaded: rows_in_file.min(visigrid_engine::sheet::NUM_ROWS),
    })
}

/// Column letters for a 0-based index (0 -> A, 26 -> AA).
fn col_label(mut c: usize) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    s
}

/// Rough count of the non-empty cells in `content`, for sizing the sheet before
/// the import fills it: non-empty fields per record in a sample from the start,
/// times the number of lines. Capped at one cell per two bytes (a value plus a
/// delimiter), so a sparse file can't reserve more than it could hold.
fn estimate_cells(content: &str, delimiter: u8) -> usize {
    const SAMPLE: usize = 1000;
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());
    let (mut records, mut filled) = (0usize, 0usize);
    for record in reader.records().take(SAMPLE).flatten() {
        records += 1;
        filled += record.iter().filter(|f| !f.is_empty()).count();
    }
    if records == 0 {
        return 0;
    }
    let lines = content.as_bytes().iter().filter(|&&b| b == b'\n').count().max(records);
    (filled * lines / records).min(content.len() / 2)
}

pub fn export(sheet: &Sheet, path: &Path) -> Result<(), String> {
    export_with_delimiter(sheet, path, b',')
}

pub fn export_tsv(sheet: &Sheet, path: &Path) -> Result<(), String> {
    export_with_delimiter(sheet, path, b'\t')
}

fn export_with_delimiter(sheet: &Sheet, path: &Path, delimiter: u8) -> Result<(), String> {
    // Rows may be variable width because merge-hidden cells are forced empty
    // and trailing empties are omitted, so different rows can have different field counts.
    let mut writer = csv::WriterBuilder::new()
        .delimiter(delimiter)
        .flexible(true)
        .from_path(path)
        .map_err(|e| e.to_string())?;

    // Bounded by the data, not the grid: a sheet is 1,048,576 x 16,384, so
    // walking it to find 20 rows of CSV is 17 billion lookups.
    let (last_row, last_col) = sheet.data_extent();
    for row in 0..=last_row {
        let mut record: Vec<String> = Vec::new();
        let mut last_non_empty = 0;

        for col in 0..=last_col {
            let value = if sheet.is_merge_hidden(row, col) {
                String::new()
            } else {
                // Dates as ISO 8601, not serials; see get_interchange_display.
                sheet.get_interchange_display(row, col)
            };
            if !value.is_empty() {
                last_non_empty = col + 1;
            }
            record.push(value);
        }

        // Only write rows that have data
        if last_non_empty > 0 {
            record.truncate(last_non_empty);
            writer.write_record(&record).map_err(|e| e.to_string())?;
        }
    }

    writer.flush().map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {

    #[test]
    fn estimate_counts_filled_fields_per_line() {
        let content = "1,a,\n2,b,x\n3,c,\n4,d,y\n";
        // 6 of 12 sampled fields... per record: 2,3,2,3 = 10 over 4 records
        assert_eq!(super::estimate_cells(content, b','), 10);
        assert_eq!(super::estimate_cells("", b','), 0);
    }

    #[test]
    fn estimate_never_exceeds_what_the_bytes_could_hold() {
        // Dense first rows, then a long empty tail: the per-record rate from
        // the sample would claim ~3 cells for every line of the tail.
        let mut content = String::from("a,b,c\n").repeat(10);
        content.push_str(&"\n".repeat(100_000));
        let estimate = super::estimate_cells(&content, b',');
        assert!(estimate <= content.len() / 2, "{estimate} > {}", content.len() / 2);
    }

    /// A formula may name a row further down the file.
    ///
    /// Cells used to be evaluated as they were read, so a reference pointing
    /// forward was computed against a cell that had not been reached yet and
    /// came out as zero — and stayed zero, because nothing recomputed. The
    /// third row here is what made it worse than a single wrong cell: it points
    /// safely backwards and was still wrong, having inherited the first row's
    /// bad value.
    #[test]
    fn formulas_referring_to_later_rows_are_still_evaluated() {
        let options = CsvOptions { evaluate_formulas: true };
        let sheet = import_from_string_with("=A2*2\n5\n=A1+1\n", b',', options).unwrap().sheet;

        assert_eq!(sheet.get_display(0, 0), "10", "a forward reference must resolve");
        assert_eq!(sheet.get_display(1, 0), "5");
        assert_eq!(
            sheet.get_display(2, 0),
            "11",
            "and a backward reference must not inherit a wrong one"
        );
    }
    use super::*;
    use std::fs;
    use tempfile::tempdir;

    use visigrid_engine::sheet::MergedRegion;

    /// A date-formatted cell used to export as its serial (46266.58), which no
    /// CSV reader outside a spreadsheet understands. Parquet timestamps made
    /// that visible; any date column had it.
    #[test]
    fn test_csv_export_writes_dates_as_iso_8601() {
        use visigrid_engine::cell::NumberFormat;

        let dir = tempdir().unwrap();
        let path = dir.path().join("dates.csv");

        let mut sheet = Sheet::new(SheetId(1), 2, 3);
        sheet.set_value(0, 0, "placed_at");
        sheet.set_value(0, 1, "amount");
        sheet.set_value(1, 0, "46266.58472222222");
        sheet.set_number_format(1, 0, NumberFormat::DateTime);
        sheet.set_value(1, 1, "1234.5");
        sheet.set_number_format(1, 1, NumberFormat::Currency { decimals: 2, thousands: true, negative: Default::default(), symbol: None });

        export(&sheet, &path).unwrap();
        let content = fs::read_to_string(&path).unwrap();
        // Only the date changes; the currency cell exports exactly as before.
        let amount = sheet.get_display(1, 1);
        assert_eq!(content, format!("placed_at,amount\n2026-09-01 14:02:00,{}\n", amount));
        assert!(!amount.contains('$'), "currency must stay machine-readable: {}", amount);
    }

    #[test]
    fn test_csv_export_merged_cells_no_leak() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("merged.csv");

        // Small sheet to avoid dimension noise
        let mut sheet = Sheet::new(SheetId(1), 3, 4);
        sheet.set_value(0, 0, "Header");
        sheet.set_value(0, 1, "LEAK1"); // will become hidden
        sheet.set_value(0, 2, "LEAK2"); // will become hidden
        sheet.set_value(1, 0, "A");
        sheet.set_value(1, 1, "B");
        sheet.set_value(1, 2, "C");

        // Merge A1:C1 — B1/C1 become hidden but still hold residual data
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        export(&sheet, &path).unwrap();

        let content = fs::read_to_string(&path).unwrap();

        // Residual data must not appear anywhere in the output
        assert!(!content.contains("LEAK1"), "hidden merge cell B1 leaked into CSV");
        assert!(!content.contains("LEAK2"), "hidden merge cell C1 leaked into CSV");

        // Parse back with csv reader to verify structure
        let mut reader = csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .from_reader(content.as_bytes());
        let records: Vec<csv::StringRecord> = reader.records().map(|r| r.unwrap()).collect();

        // Row 0: origin value present, hidden cells empty
        assert_eq!(records[0].get(0), Some("Header"));
        let b1 = records[0].get(1).unwrap_or("");
        let c1 = records[0].get(2).unwrap_or("");
        assert!(b1.is_empty(), "B1 should be empty, got: {b1}");
        assert!(c1.is_empty(), "C1 should be empty, got: {c1}");

        // Row 1: normal cells unaffected
        assert_eq!(records[1].get(0), Some("A"));
        assert_eq!(records[1].get(1), Some("B"));
        assert_eq!(records[1].get(2), Some("C"));
    }

    #[test]
    fn test_sniff_semicolon_delimiter() {
        let content = "Name;Age;City\nAlice;30;Paris\nBob;25;London\n";
        assert_eq!(sniff_delimiter(content), b';');
    }

    #[test]
    fn test_sniff_comma_delimiter() {
        let content = "Name,Age,City\nAlice,30,Paris\nBob,25,London\n";
        assert_eq!(sniff_delimiter(content), b',');
    }

    #[test]
    fn test_sniff_tab_delimiter() {
        let content = "Name\tAge\tCity\nAlice\t30\tParis\nBob\t25\tLondon\n";
        assert_eq!(sniff_delimiter(content), b'\t');
    }

    #[test]
    fn test_sniff_pipe_delimiter() {
        let content = "Name|Age|City\nAlice|30|Paris\nBob|25|London\n";
        assert_eq!(sniff_delimiter(content), b'|');
    }

    #[test]
    fn test_parse_delimited_text_csv() {
        let text = "Name,Age,City\nAlice,30,Paris\nBob,25,London\n";
        let grid = parse_delimited_text(text);
        assert_eq!(grid.len(), 3);
        assert_eq!(grid[0], vec!["Name", "Age", "City"]);
        assert_eq!(grid[1], vec!["Alice", "30", "Paris"]);
        assert_eq!(grid[2], vec!["Bob", "25", "London"]);
    }

    #[test]
    fn test_parse_delimited_text_quoted_fields() {
        let text = "Name,Address,City\n\"Doe, Jane\",\"123 Main St\",Paris\n";
        let grid = parse_delimited_text(text);
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[1][0], "Doe, Jane");
        assert_eq!(grid[1][1], "123 Main St");
    }

    #[test]
    fn test_parse_delimited_text_single_column() {
        // Plain text lines with no delimiters — should return single-column rows
        let text = "hello\nworld\n";
        let grid = parse_delimited_text(text);
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[0].len(), 1);
        assert_eq!(grid[0][0], "hello");
    }

    #[test]
    fn test_parse_delimited_text_pipe() {
        let text = "A|B|C\n1|2|3\n";
        let grid = parse_delimited_text(text);
        assert_eq!(grid[0], vec!["A", "B", "C"]);
        assert_eq!(grid[1], vec!["1", "2", "3"]);
    }

    #[test]
    fn test_sniff_semicolon_with_commas_in_values() {
        // Semicolon delimiter but commas appear inside quoted fields
        let content = "Name;Address;City\n\"Doe, Jane\";\"123 Main St, Apt 4\";Paris\nBob;\"456 Elm\";London\n";
        assert_eq!(sniff_delimiter(content), b';');
    }

    #[test]
    fn test_semicolon_csv_import() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.csv");
        fs::write(&path, "Name;Age;City\nAlice;30;Paris\nBob;25;London\n").unwrap();

        let sheet = import(&path).unwrap();
        assert_eq!(sheet.get_display(0, 0), "Name");
        assert_eq!(sheet.get_display(0, 1), "Age");
        assert_eq!(sheet.get_display(0, 2), "City");
        assert_eq!(sheet.get_display(1, 0), "Alice");
        assert_eq!(sheet.get_display(1, 1), "30");
        assert_eq!(sheet.get_display(1, 2), "Paris");
    }

    #[test]
    fn test_tsv_roundtrip() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.tsv");

        // Create a sheet with some data
        let mut sheet = Sheet::new(SheetId(1), 100, 10);
        sheet.set_value(0, 0, "Name");
        sheet.set_value(0, 1, "Value");
        sheet.set_value(1, 0, "Alice");
        sheet.set_value(1, 1, "42");
        sheet.set_value(2, 0, "Bob");
        sheet.set_value(2, 1, "17");

        // Export to TSV
        export_tsv(&sheet, &path).unwrap();

        // Verify the file contains tabs
        let content = fs::read_to_string(&path).unwrap();
        assert!(content.contains('\t'), "TSV should contain tab characters");
        assert!(!content.contains(','), "TSV should not contain commas as delimiters");

        // Import back
        let imported = import_tsv(&path).unwrap();
        assert_eq!(imported.get_display(0, 0), "Name");
        assert_eq!(imported.get_display(0, 1), "Value");
        assert_eq!(imported.get_display(1, 0), "Alice");
        assert_eq!(imported.get_display(1, 1), "42");
        assert_eq!(imported.get_display(2, 0), "Bob");
        assert_eq!(imported.get_display(2, 1), "17");
    }

    // --- Safe defaults (#65 and the leading-zeros question) ---

    fn text(sheet: &Sheet, r: usize, c: usize) -> (String, bool) {
        let v = sheet.get_computed_value(r, c);
        (sheet.get_display(r, c), matches!(v, visigrid_engine::formula::eval::Value::Text(_)))
    }

    #[test]
    fn leading_zeros_and_long_ids_stay_text() {
        let csv = "zip,sku,card,gene_id,amount\n00501,007,4111111111111111,2310009E13,1.5\n02134,0123,12345678901234567890,12E4,0\n";
        let r = import_from_string_with(csv, b',', CsvOptions::default()).unwrap();
        assert_eq!(text(&r.sheet, 1, 0), ("00501".into(), true));
        assert_eq!(text(&r.sheet, 1, 1), ("007".into(), true));
        assert_eq!(text(&r.sheet, 1, 2), ("4111111111111111".into(), true), "16 digits: past a double's exact range");
        assert_eq!(text(&r.sheet, 1, 3), ("2310009E13".into(), true));
        assert_eq!(text(&r.sheet, 2, 2), ("12345678901234567890".into(), true));
        assert_eq!(text(&r.sheet, 2, 3), ("12E4".into(), true));
        // Real numbers are still numbers.
        assert_eq!(text(&r.sheet, 1, 4).1, false);
        assert_eq!(text(&r.sheet, 2, 4), ("0".into(), false));
        assert_eq!(r.kept_as_text, 8);
        assert_eq!(r.text_columns, ["zip", "sku", "card", "gene_id"]);
    }

    #[test]
    fn ordinary_values_are_still_typed() {
        let csv = "a\n42\n-7\n+3\n0.5\n1.5e-07\n2.5E+3\n5%\n$1,234\n(500)\n123456789012345\n";
        let r = import_from_string_with(csv, b',', CsvOptions::default()).unwrap();
        for row in 1..=10 {
            assert!(!text(&r.sheet, row, 0).1, "row {row} ({}) should be a number", r.sheet.get_display(row, 0));
        }
        assert_eq!(r.kept_as_text, 0);
        assert_eq!(r.message(), None, "nothing to say");
    }

    #[test]
    fn nothing_becomes_a_date() {
        let csv = "g\nMARCH1\nSEPT2\n1-2\n2024-01-05\n1/5/2024\n";
        let r = import_from_string_with(csv, b',', CsvOptions::default()).unwrap();
        for (row, want) in [(1, "MARCH1"), (2, "SEPT2"), (3, "1-2"), (4, "2024-01-05"), (5, "1/5/2024")] {
            assert_eq!(text(&r.sheet, row, 0), (want.to_string(), true));
        }
    }

    #[test]
    fn formulas_stay_text_unless_asked_for() {
        let csv = "a,b\n2,=A2*10\n";
        let r = import_from_string_with(csv, b',', CsvOptions::default()).unwrap();
        assert_eq!(text(&r.sheet, 1, 1), ("=A2*10".into(), true));
        assert_eq!(r.formulas_as_text, 1);
        assert!(r.message().unwrap().contains("Left 1 formula as text"));
        let r = import_from_string_with(csv, b',', CsvOptions { evaluate_formulas: true }).unwrap();
        assert_eq!(r.sheet.get_display(1, 1), "20");
        assert_eq!(r.formulas_as_text, 0);
    }

    #[test]
    fn rows_past_the_grid_are_reported_not_kept() {
        let rows = visigrid_engine::sheet::NUM_ROWS + 3;
        let mut csv = String::with_capacity(rows * 2);
        for _ in 0..rows {
            csv.push_str("1\n");
        }
        let r = import_from_string_with(&csv, b',', CsvOptions::default()).unwrap();
        assert_eq!(r.rows_in_file, rows);
        assert_eq!(r.rows_loaded, visigrid_engine::sheet::NUM_ROWS);
        assert!(r.message().unwrap().contains("first 1048576 of 1048579 rows"));
    }

    #[test]
    fn round_trip_keeps_ids_and_full_precision() {
        // #65: big integers printed as 9223372036854775807 and General numbers
        // exported at two decimals.
        let csv = "id,val\n12345678901234567890,1.234\n2310009E13,0.1234567\n007,0.5\n";
        let r = import_from_string_with(csv, b',', CsvOptions::default()).unwrap();
        let dir = tempdir().unwrap();
        let out = dir.path().join("out.csv");
        export(&r.sheet, &out).unwrap();
        assert_eq!(fs::read_to_string(&out).unwrap(), csv);
    }

    #[test]
    fn interchange_numbers_round_trip() {
        use visigrid_engine::cell::interchange_number;
        for n in [0.0, 1.234, 0.1234567, -2.5, 1e15, 9007199254740993.0, 1.2345678901234567e19, 1e300, 1.5e-7, 0.1 + 0.2] {
            let s = interchange_number(n);
            assert_eq!(s.parse::<f64>().unwrap(), n, "{n} -> {s}");
        }
        assert_eq!(interchange_number(1.234), "1.234");
        assert_eq!(interchange_number(42.0), "42");
    }
}
