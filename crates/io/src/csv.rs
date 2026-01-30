// CSV/TSV import/export

use std::path::Path;
use std::io::Read;

use visigrid_engine::sheet::{Sheet, SheetId};

pub fn import(path: &Path) -> Result<Sheet, String> {
    import_with_delimiter(path, b',')
}

pub fn import_tsv(path: &Path) -> Result<Sheet, String> {
    import_with_delimiter(path, b'\t')
}

/// Read file and convert to UTF-8 if needed (handles Windows-1252, Latin-1, etc.)
fn read_file_as_utf8(path: &Path) -> Result<String, String> {
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).map_err(|e| e.to_string())?;

    // Try UTF-8 first
    if let Ok(s) = String::from_utf8(bytes.clone()) {
        return Ok(s);
    }

    // Fall back to Windows-1252 (common for Excel-exported CSVs)
    // This handles most non-UTF-8 files from Windows/Excel
    let (decoded, _, _) = encoding_rs::WINDOWS_1252.decode(&bytes);
    Ok(decoded.into_owned())
}

fn import_with_delimiter(path: &Path, delimiter: u8) -> Result<Sheet, String> {
    // Read file with encoding conversion
    let content = read_file_as_utf8(path)?;

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
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
