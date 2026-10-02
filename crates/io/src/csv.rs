// CSV/TSV import/export

use std::path::Path;

use visigrid_engine::sheet::Sheet;

pub use crate::csv_import::{
    count_records, import_report, import_text, preview, preview_bytes, read_head, ColumnDecision, ColumnRule, CsvImport, CsvOptions, DateOrder, Encoding, Resolved,
};

pub fn import(path: &Path) -> Result<Sheet, String> {
    import_report(path, &CsvOptions::default()).map(|r| r.sheet)
}

pub fn import_tsv(path: &Path) -> Result<Sheet, String> {
    import_with_delimiter(path, b'\t')
}

pub fn import_with_delimiter(path: &Path, delimiter: u8) -> Result<Sheet, String> {
    import_report(path, &CsvOptions { delimiter: Some(delimiter), ..Default::default() }).map(|r| r.sheet)
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
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    Ok(crate::csv_import::decode(&bytes, None).0)
}

#[cfg(test)]
fn import_from_string(content: &str, delimiter: u8) -> Result<Sheet, String> {
    import_from_string_with(content, delimiter, CsvOptions::default()).map(|r| r.sheet)
}

#[cfg(test)]
fn import_from_string_with(content: &str, delimiter: u8, options: CsvOptions) -> Result<CsvImport, String> {
    let options = CsvOptions { delimiter: Some(delimiter), ..options };
    crate::csv_import::import_str(content, crate::csv_import::Encoding::Utf8, &options, usize::MAX)
}

/// Rough count of the non-empty cells in `content`, for sizing the sheet before
/// the import fills it: non-empty fields per record in a sample from the start,
/// times the number of lines. Capped at one cell per two bytes (a value plus a
/// delimiter), so a sparse file can't reserve more than it could hold.
pub(crate) fn estimate_cells(content: &str, delimiter: u8) -> usize {
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
        let options = CsvOptions { evaluate_formulas: true, no_header: true, ..Default::default() };
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
    use visigrid_engine::sheet::SheetId;
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
        let r = import_from_string_with(csv, b',', CsvOptions { evaluate_formulas: true, ..Default::default() }).unwrap();
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
