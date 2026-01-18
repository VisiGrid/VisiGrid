// Native .sheet format using SQLite

use std::path::Path;

use rusqlite::{Connection, params};

use crate::engine::cell::{Alignment, CellFormat, NumberFormat};
use crate::engine::sheet::Sheet;

const SCHEMA: &str = r#"
CREATE TABLE IF NOT EXISTS cells (
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    value_type INTEGER NOT NULL,  -- 0=empty, 1=number, 2=text, 3=formula
    value_num REAL,
    value_text TEXT,
    fmt_bold INTEGER DEFAULT 0,
    fmt_italic INTEGER DEFAULT 0,
    fmt_underline INTEGER DEFAULT 0,
    fmt_alignment INTEGER DEFAULT 0,     -- 0=left, 1=center, 2=right
    fmt_number_type INTEGER DEFAULT 0,   -- 0=general, 1=number, 2=currency, 3=percent
    fmt_decimals INTEGER DEFAULT 2,      -- decimal places
    PRIMARY KEY (row, col)
);

CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"#;

// Value type constants
const TYPE_EMPTY: i32 = 0;
const TYPE_NUMBER: i32 = 1;
const TYPE_TEXT: i32 = 2;
const TYPE_FORMULA: i32 = 3;

pub fn save(sheet: &Sheet, path: &Path) -> Result<(), String> {
    // Delete existing file if present (SQLite will create fresh)
    if path.exists() {
        std::fs::remove_file(path).map_err(|e| e.to_string())?;
    }

    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Create schema
    conn.execute_batch(SCHEMA).map_err(|e| e.to_string())?;

    // Save metadata
    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?1, ?2)",
        params!["sheet_name", &sheet.name],
    ).map_err(|e| e.to_string())?;

    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?1, ?2)",
        params!["rows", sheet.rows.to_string()],
    ).map_err(|e| e.to_string())?;

    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?1, ?2)",
        params!["cols", sheet.cols.to_string()],
    ).map_err(|e| e.to_string())?;

    // Save cells using a transaction for performance
    conn.execute("BEGIN TRANSACTION", []).map_err(|e| e.to_string())?;

    {
        let mut stmt = conn.prepare(
            "INSERT INTO cells (row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11)"
        ).map_err(|e| e.to_string())?;

        for row in 0..sheet.rows {
            for col in 0..sheet.cols {
                let raw = sheet.get_raw(row, col);
                let format = sheet.get_format(row, col);

                // Skip cells with no value and default formatting
                let has_formatting = format.bold || format.italic || format.underline
                    || format.alignment != Alignment::Left
                    || !matches!(format.number_format, NumberFormat::General);
                if raw.is_empty() && !has_formatting {
                    continue;
                }

                // Determine value type and store appropriately
                let (value_type, value_num, value_text): (i32, Option<f64>, Option<&str>) =
                    if raw.is_empty() {
                        (TYPE_EMPTY, None, None)
                    } else if raw.starts_with('=') {
                        (TYPE_FORMULA, None, Some(&raw))
                    } else if let Ok(num) = raw.parse::<f64>() {
                        (TYPE_NUMBER, Some(num), None)
                    } else {
                        (TYPE_TEXT, None, Some(&raw))
                    };

                // Convert alignment to integer
                let alignment_int = match format.alignment {
                    Alignment::Left => 0,
                    Alignment::Center => 1,
                    Alignment::Right => 2,
                };

                // Convert number format to integer + decimals
                let (number_type, decimals) = match format.number_format {
                    NumberFormat::General => (0, 2),
                    NumberFormat::Number { decimals } => (1, decimals as i32),
                    NumberFormat::Currency { decimals } => (2, decimals as i32),
                    NumberFormat::Percent { decimals } => (3, decimals as i32),
                };

                stmt.execute(params![
                    row as i64,
                    col as i64,
                    value_type,
                    value_num,
                    value_text,
                    format.bold as i32,
                    format.italic as i32,
                    format.underline as i32,
                    alignment_int,
                    number_type,
                    decimals
                ]).map_err(|e| e.to_string())?;
            }
        }
    }

    conn.execute("COMMIT", []).map_err(|e| e.to_string())?;

    Ok(())
}

pub fn load(path: &Path) -> Result<Sheet, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Load metadata
    let sheet_name: String = conn
        .query_row("SELECT value FROM meta WHERE key = 'sheet_name'", [], |row| row.get(0))
        .unwrap_or_else(|_| "Sheet1".to_string());

    let rows: usize = conn
        .query_row("SELECT value FROM meta WHERE key = 'rows'", [], |row| {
            let s: String = row.get(0)?;
            Ok(s.parse().unwrap_or(1000))
        })
        .unwrap_or(1000);

    let cols: usize = conn
        .query_row("SELECT value FROM meta WHERE key = 'cols'", [], |row| {
            let s: String = row.get(0)?;
            Ok(s.parse().unwrap_or(26))
        })
        .unwrap_or(26);

    let mut sheet = Sheet::new(rows, cols);
    sheet.name = sheet_name;

    // Load cells - check if new format columns exist for backward compatibility
    let has_new_columns = conn
        .prepare("SELECT fmt_alignment FROM cells LIMIT 1")
        .is_ok();

    let query = if has_new_columns {
        "SELECT row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals FROM cells"
    } else {
        "SELECT row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline FROM cells"
    };

    let mut stmt = conn.prepare(query).map_err(|e| e.to_string())?;

    let cell_iter = stmt
        .query_map([], |row| {
            let r: i64 = row.get(0)?;
            let c: i64 = row.get(1)?;
            let value_type: i32 = row.get(2)?;
            let value_num: Option<f64> = row.get(3)?;
            let value_text: Option<String> = row.get(4)?;
            let fmt_bold: i32 = row.get(5).unwrap_or(0);
            let fmt_italic: i32 = row.get(6).unwrap_or(0);
            let fmt_underline: i32 = row.get(7).unwrap_or(0);
            let fmt_alignment: i32 = row.get(8).unwrap_or(0);
            let fmt_number_type: i32 = row.get(9).unwrap_or(0);
            let fmt_decimals: i32 = row.get(10).unwrap_or(2);
            Ok((r as usize, c as usize, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals))
        })
        .map_err(|e| e.to_string())?;

    for cell_result in cell_iter {
        let (row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals) =
            cell_result.map_err(|e| e.to_string())?;

        let value = match value_type {
            TYPE_NUMBER => {
                if let Some(n) = value_num {
                    if n.fract() == 0.0 {
                        format!("{}", n as i64)
                    } else {
                        format!("{}", n)
                    }
                } else {
                    String::new()
                }
            }
            TYPE_TEXT | TYPE_FORMULA => value_text.unwrap_or_default(),
            _ => String::new(),
        };

        if !value.is_empty() {
            sheet.set_value(row, col, &value);
        }

        // Apply formatting
        let alignment = match fmt_alignment {
            1 => Alignment::Center,
            2 => Alignment::Right,
            _ => Alignment::Left,
        };
        let number_format = match fmt_number_type {
            1 => NumberFormat::Number { decimals: fmt_decimals as u8 },
            2 => NumberFormat::Currency { decimals: fmt_decimals as u8 },
            3 => NumberFormat::Percent { decimals: fmt_decimals as u8 },
            _ => NumberFormat::General,
        };
        let format = CellFormat {
            bold: fmt_bold != 0,
            italic: fmt_italic != 0,
            underline: fmt_underline != 0,
            alignment,
            number_format,
        };
        let has_formatting = format.bold || format.italic || format.underline
            || format.alignment != Alignment::Left
            || !matches!(format.number_format, NumberFormat::General);
        if has_formatting {
            sheet.set_format(row, col, format);
        }
    }

    Ok(sheet)
}
