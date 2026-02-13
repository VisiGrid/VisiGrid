// CSV/TSV import/export

use std::path::Path;
use std::io::Read;

use visigrid_engine::sheet::{Sheet, SheetId};

pub fn import(path: &Path) -> Result<Sheet, String> {
    let content = read_file_as_utf8(path)?;
    let delimiter = sniff_delimiter(&content);
    import_from_string(&content, delimiter)
}

pub fn import_tsv(path: &Path) -> Result<Sheet, String> {
    let content = read_file_as_utf8(path)?;
    import_from_string(&content, b'\t')
}

pub fn import_with_delimiter(path: &Path, delimiter: u8) -> Result<Sheet, String> {
    let content = read_file_as_utf8(path)?;
    import_from_string(&content, delimiter)
}

/// Detect the most likely field delimiter by checking consistency across the first few lines.
///
/// For each candidate (tab, semicolon, comma, pipe), count fields per line. The delimiter
/// that produces the most consistent field count (>1 field) wins.
fn sniff_delimiter(content: &str) -> u8 {
    let candidates: &[u8] = &[b'\t', b';', b',', b'|'];
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

fn import_from_string(content: &str, delimiter: u8) -> Result<Sheet, String> {
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    // Start with reasonable defaults, will track actual extent
    let mut sheet = Sheet::new(SheetId(1), 65536, 256);
    let mut max_row = 0usize;
    let mut max_col = 0usize;

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| e.to_string())?;
        for (col_idx, field) in record.iter().enumerate() {
            if !field.is_empty() {
                sheet.set_value(row_idx, col_idx, field);
                max_col = max_col.max(col_idx);
            }
        }
        max_row = row_idx;
    }

    // Update sheet dimensions to actual data extent (for export efficiency)
    sheet.rows = (max_row + 1).max(1000);
    sheet.cols = (max_col + 1).max(26);

    Ok(sheet)
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

    for row in 0..sheet.rows {
        let mut record: Vec<String> = Vec::new();
        let mut last_non_empty = 0;

        for col in 0..sheet.cols {
            let value = if sheet.is_merge_hidden(row, col) {
                String::new()
            } else {
                sheet.get_display(row, col)
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
    use super::*;
    use std::fs;
    use tempfile::tempdir;

    use visigrid_engine::sheet::MergedRegion;

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
}
