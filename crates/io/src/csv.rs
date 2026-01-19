// CSV import/export

use std::path::Path;

use visigrid_engine::sheet::Sheet;

pub fn import(path: &Path) -> Result<Sheet, String> {
    let mut reader = csv::Reader::from_path(path).map_err(|e| e.to_string())?;

    let mut sheet = Sheet::new(1000, 26);

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| e.to_string())?;
        for (col_idx, field) in record.iter().enumerate() {
            if !field.is_empty() {
                sheet.set_value(row_idx, col_idx, field);
            }
        }
    }

    Ok(sheet)
}

pub fn export(sheet: &Sheet, path: &Path) -> Result<(), String> {
    let mut writer = csv::Writer::from_path(path).map_err(|e| e.to_string())?;

    for row in 0..sheet.rows {
        let mut record: Vec<String> = Vec::new();
        let mut last_non_empty = 0;

        for col in 0..sheet.cols {
            let value = sheet.get_display(row, col);
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
