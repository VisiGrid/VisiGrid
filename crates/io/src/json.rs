// JSON export

use std::path::Path;
use std::fs::File;
use std::io::BufWriter;

use visigrid_engine::sheet::{Sheet, SheetId};

/// Export sheet as JSON array of arrays
/// Each row is an array of cell values (strings)
pub fn export(sheet: &Sheet, path: &Path) -> Result<(), String> {
    let file = File::create(path).map_err(|e| e.to_string())?;
    let writer = BufWriter::new(file);

    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut last_non_empty_row = 0;

    for row in 0..sheet.rows {
        let mut record: Vec<String> = Vec::new();
        let mut last_non_empty_col = 0;

        for col in 0..sheet.cols {
            let value = sheet.get_display(row, col);
            if !value.is_empty() {
                last_non_empty_col = col + 1;
                last_non_empty_row = row + 1;
            }
            record.push(value);
        }

        // Trim trailing empty cells
        record.truncate(last_non_empty_col);
        rows.push(record);
    }

    // Trim trailing empty rows
    rows.truncate(last_non_empty_row);

    serde_json::to_writer_pretty(writer, &rows).map_err(|e| e.to_string())?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::tempdir;

    #[test]
    fn test_json_export() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.json");

        let mut sheet = Sheet::new(SheetId(1), 100, 10);
        sheet.set_value(0, 0, "Name");
        sheet.set_value(0, 1, "Value");
        sheet.set_value(1, 0, "Alice");
        sheet.set_value(1, 1, "42");
        sheet.set_value(2, 0, "Bob");
        sheet.set_value(2, 1, "17");

        export(&sheet, &path).unwrap();

        let content = fs::read_to_string(&path).unwrap();
        let parsed: Vec<Vec<String>> = serde_json::from_str(&content).unwrap();

        assert_eq!(parsed.len(), 3);
        assert_eq!(parsed[0], vec!["Name", "Value"]);
        assert_eq!(parsed[1], vec!["Alice", "42"]);
        assert_eq!(parsed[2], vec!["Bob", "17"]);
    }
}
