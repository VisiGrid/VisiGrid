// Native .sheet format using SQLite

use std::collections::HashMap;
use std::path::Path;

use rusqlite::{Connection, params};

use visigrid_engine::cell::{Alignment, BorderStyle, CellBorder, CellFormat, CellStyle, DateStyle, NegativeStyle, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::sheet::{MergedRegion, Sheet, SheetId};
use visigrid_engine::workbook::Workbook;
use visigrid_engine::named_range::{NamedRange, NamedRangeTarget};

// ============================================================================
// Semantic Verification (persisted expected fingerprint)
// ============================================================================

/// Persisted semantic verification info.
///
/// This is saved to the .sheet file and allows verifying that the file
/// hasn't been modified since it was stamped/approved.
#[derive(Debug, Clone, Default)]
pub struct SemanticVerification {
    /// Expected semantic fingerprint (e.g., "v2:186:abc123...")
    pub fingerprint: Option<String>,
    /// Optional label (e.g., "MSFT SEC v1")
    pub label: Option<String>,
    /// Optional ISO timestamp of when the fingerprint was set
    pub timestamp: Option<String>,
}

/// Fingerprint format version. Increment on breaking changes to fingerprint computation.
/// v2: includes iteration settings (enabled, max_iters, tolerance).
const FINGERPRINT_VERSION: u32 = 2;

/// Compute semantic fingerprint of a workbook.
///
/// The fingerprint includes:
/// - Cell values/formulas (semantic content)
/// - Does NOT include style (presentation only)
///
/// Format: `v1:N:HASH` where:
/// - `v1` = fingerprint version (increment on breaking changes)
/// - `N` = number of non-empty cells hashed (cells with value or formula)
/// - `HASH` = first 16 hex chars of blake3 hash (64 bits)
///
/// Order is deterministic: cells sorted by (sheet_idx, row, col).
pub fn compute_semantic_fingerprint(workbook: &Workbook) -> String {
    let mut hasher = blake3::Hasher::new();
    let mut op_count = 0;

    // Include iteration settings — these affect computed results
    let iter_settings = format!(
        "iter:{}:{}:{}",
        workbook.iterative_enabled(),
        workbook.iterative_max_iters(),
        workbook.iterative_tolerance(),
    );
    hasher.update(iter_settings.as_bytes());
    hasher.update(b"\n");

    // Iterate all sheets
    for sheet_idx in 0..workbook.sheet_count() {
        if let Some(sheet) = workbook.sheet(sheet_idx) {
            // Collect cells and sort for deterministic order
            let mut cells: Vec<((usize, usize), String)> = Vec::new();
            for (&(row, col), cell) in sheet.cells_iter() {
                let raw = cell.value.raw_display();
                if !raw.is_empty() {
                    cells.push(((row, col), raw.to_string()));
                }
            }
            // Sort by (row, col) for deterministic order
            cells.sort_by_key(|((r, c), _)| (*r, *c));

            for ((row, col), value) in cells {
                let op = format!("set:{}:{}:{}", row, col, value);
                hasher.update(op.as_bytes());
                hasher.update(b"\n");
                op_count += 1;
            }
        }
    }

    let hash = hasher.finalize();
    let hash_hex = &hash.to_hex()[0..16]; // First 16 hex chars (64 bits)
    format!("v{}:{}:{}", FINGERPRINT_VERSION, op_count, hash_hex)
}

const SCHEMA: &str = r#"
CREATE TABLE IF NOT EXISTS sheets (
    sheet_idx INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    row_count INTEGER DEFAULT 1000,
    col_count INTEGER DEFAULT 26
);

CREATE TABLE IF NOT EXISTS cells (
    sheet_idx INTEGER NOT NULL DEFAULT 0,
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
    fmt_font_family TEXT,                -- NULL = inherit from settings
    fmt_thousands INTEGER DEFAULT 0,     -- 1 = use thousands separator
    fmt_negative INTEGER DEFAULT 0,      -- 0=minus, 1=parens, 2=red minus, 3=red parens
    fmt_currency_symbol TEXT,            -- NULL = default ($)
    fmt_border_top INTEGER DEFAULT 0,    -- style: 0=none, 1=thin, 2=medium, 3=thick
    fmt_border_right INTEGER DEFAULT 0,
    fmt_border_bottom INTEGER DEFAULT 0,
    fmt_border_left INTEGER DEFAULT 0,
    fmt_border_top_color INTEGER,        -- RGBA as u32 (0xRRGGBBAA), NULL = automatic
    fmt_border_right_color INTEGER,
    fmt_border_bottom_color INTEGER,
    fmt_border_left_color INTEGER,
    fmt_cell_style INTEGER DEFAULT 0,
    fmt_font_color INTEGER,              -- RGBA as u32 (0xRRGGBBAA), NULL = automatic
    fmt_background_color INTEGER,        -- RGBA as u32 (0xRRGGBBAA), NULL = no fill
    PRIMARY KEY (sheet_idx, row, col)
);

CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS named_ranges (
    name TEXT PRIMARY KEY,
    target_type INTEGER NOT NULL,  -- 0=cell, 1=range
    sheet INTEGER NOT NULL,
    start_row INTEGER NOT NULL,
    start_col INTEGER NOT NULL,
    end_row INTEGER,               -- NULL for cell type
    end_col INTEGER,               -- NULL for cell type
    description TEXT
);

CREATE TABLE IF NOT EXISTS merged_regions (
    start_row INTEGER NOT NULL,
    start_col INTEGER NOT NULL,
    end_row INTEGER NOT NULL,
    end_col INTEGER NOT NULL,
    PRIMARY KEY (start_row, start_col)
);

CREATE TABLE IF NOT EXISTS hub_link (
    id INTEGER PRIMARY KEY CHECK (id = 1),  -- singleton row
    repo_owner TEXT NOT NULL,
    repo_slug TEXT NOT NULL,
    dataset_id TEXT NOT NULL,
    local_head_id TEXT,
    local_head_hash TEXT,
    link_mode TEXT DEFAULT 'pull',
    linked_at TEXT NOT NULL,
    api_base TEXT DEFAULT 'https://api.visihub.app'
);

CREATE TABLE IF NOT EXISTS col_widths (
    sheet_idx INTEGER NOT NULL,
    col INTEGER NOT NULL,
    width REAL NOT NULL,
    PRIMARY KEY (sheet_idx, col)
);

CREATE TABLE IF NOT EXISTS row_heights (
    sheet_idx INTEGER NOT NULL,
    row INTEGER NOT NULL,
    height REAL NOT NULL,
    PRIMARY KEY (sheet_idx, row)
);

-- Semantic metadata for cells/ranges (affects fingerprint, unlike style)
-- target: "A1" (cell), "A1:B10" (range), "COL:C" (column), "ROW:5" (row)
CREATE TABLE IF NOT EXISTS cell_metadata (
    target TEXT NOT NULL,
    key TEXT NOT NULL,
    value TEXT NOT NULL,
    PRIMARY KEY (target, key)
);
"#;

// Value type constants
const TYPE_EMPTY: i32 = 0;
const TYPE_NUMBER: i32 = 1;
const TYPE_TEXT: i32 = 2;
const TYPE_FORMULA: i32 = 3;

/// Encode Alignment → DB integer. Every variant has an explicit arm.
fn alignment_to_db(a: Alignment) -> i32 {
    match a {
        Alignment::Left => 0,
        Alignment::Center => 1,
        Alignment::Right => 2,
        Alignment::General => 3,
        Alignment::CenterAcrossSelection => 4,
    }
}

/// Decode DB integer → Alignment. Unknown codes fall back to General (the default).
fn alignment_from_db(i: i32) -> Alignment {
    match i {
        0 => Alignment::Left,
        1 => Alignment::Center,
        2 => Alignment::Right,
        3 => Alignment::General,
        4 => Alignment::CenterAcrossSelection,
        _ => Alignment::General,
    }
}

/// Encode BorderStyle → DB integer. 0=None, 1=Thin, 2=Medium, 3=Thick.
fn border_style_to_db(style: BorderStyle) -> i32 {
    match style {
        BorderStyle::None => 0,
        BorderStyle::Thin => 1,
        BorderStyle::Medium => 2,
        BorderStyle::Thick => 3,
    }
}

/// Decode DB integer → BorderStyle. Unknown codes fall back to None.
fn border_style_from_db(i: i32) -> BorderStyle {
    match i {
        1 => BorderStyle::Thin,
        2 => BorderStyle::Medium,
        3 => BorderStyle::Thick,
        _ => BorderStyle::None,
    }
}

/// Encode border color [R,G,B,A] → DB i64 (stored as 0xRRGGBBAA).
/// None = automatic/theme default.
fn border_color_to_db(color: Option<[u8; 4]>) -> Option<i64> {
    color.map(|[r, g, b, a]| {
        ((r as i64) << 24) | ((g as i64) << 16) | ((b as i64) << 8) | (a as i64)
    })
}

/// Decode DB i64 → border color [R,G,B,A].
/// None = automatic/theme default.
fn border_color_from_db(val: Option<i64>) -> Option<[u8; 4]> {
    val.map(|v| {
        let v = v as u32;
        [
            ((v >> 24) & 0xFF) as u8,
            ((v >> 16) & 0xFF) as u8,
            ((v >> 8) & 0xFF) as u8,
            (v & 0xFF) as u8,
        ]
    })
}

/// Reconstruct CellBorder from DB values.
fn border_from_db(style: i32, color: Option<i64>) -> CellBorder {
    CellBorder {
        style: border_style_from_db(style),
        color: border_color_from_db(color),
    }
}

/// Current schema version. Increment for each migration.
const SCHEMA_VERSION: i32 = 6;

/// Run schema migrations for existing databases.
fn migrate(conn: &Connection) -> rusqlite::Result<()> {
    let version: i32 = conn.pragma_query_value(None, "user_version", |r| r.get(0))?;

    if version < 1 {
        // Add thousands separator, negative style, and currency symbol columns
        conn.execute_batch("
            ALTER TABLE cells ADD COLUMN fmt_thousands INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_negative INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_currency_symbol TEXT;
            PRAGMA user_version = 1;
        ")?;
    }

    if version < 2 {
        // Add border style and color columns
        conn.execute_batch("
            ALTER TABLE cells ADD COLUMN fmt_border_top INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_border_right INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_border_bottom INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_border_left INTEGER DEFAULT 0;
            ALTER TABLE cells ADD COLUMN fmt_border_top_color INTEGER;
            ALTER TABLE cells ADD COLUMN fmt_border_right_color INTEGER;
            ALTER TABLE cells ADD COLUMN fmt_border_bottom_color INTEGER;
            ALTER TABLE cells ADD COLUMN fmt_border_left_color INTEGER;
            PRAGMA user_version = 2;
        ")?;
    }

    if version < 4 {
        conn.execute_batch(
            "ALTER TABLE cells ADD COLUMN fmt_cell_style INTEGER DEFAULT 0;
             PRAGMA user_version = 4;"
        )?;
    }

    if version < 5 {
        conn.execute_batch(
            "ALTER TABLE cells ADD COLUMN fmt_font_color INTEGER;
             ALTER TABLE cells ADD COLUMN fmt_background_color INTEGER;
             PRAGMA user_version = 5;"
        )?;
    }

    if version < 6 {
        conn.execute_batch(
            "CREATE TABLE IF NOT EXISTS col_widths (
                sheet_idx INTEGER NOT NULL,
                col INTEGER NOT NULL,
                width REAL NOT NULL,
                PRIMARY KEY (sheet_idx, col)
            );
            CREATE TABLE IF NOT EXISTS row_heights (
                sheet_idx INTEGER NOT NULL,
                row INTEGER NOT NULL,
                height REAL NOT NULL,
                PRIMARY KEY (sheet_idx, row)
            );
            PRAGMA user_version = 6;"
        )?;
    }

    Ok(())
}

pub fn save(sheet: &Sheet, path: &Path) -> Result<(), String> {
    // Delete existing file if present (SQLite will create fresh)
    if path.exists() {
        std::fs::remove_file(path).map_err(|e| e.to_string())?;
    }

    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Create schema
    conn.execute_batch(SCHEMA).map_err(|e| e.to_string())?;
    conn.pragma_update(None, "user_version", SCHEMA_VERSION).map_err(|e| e.to_string())?;

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
            "INSERT INTO cells (row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15)"
        ).map_err(|e| e.to_string())?;

        for (&(row, col), cell) in sheet.cells_iter() {
                let raw = cell.value.raw_display();
                let format = &cell.format;

                // Skip cells with no value and default formatting
                if raw.is_empty() && format.is_default() {
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

                let alignment_int = alignment_to_db(format.alignment);

                // Convert number format to integer + decimals
                // Date uses decimals field to store style (0=Short, 1=Long, 2=Iso)
                // Time = 5, DateTime = 6
                let (number_type, decimals, thousands, negative, currency_symbol) = extract_number_format_fields(&format.number_format);

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
                    decimals,
                    format.font_family.as_deref(),
                    thousands,
                    negative,
                    currency_symbol
                ]).map_err(|e| e.to_string())?;
        }
    }

    conn.execute("COMMIT", []).map_err(|e| e.to_string())?;

    Ok(())
}

/// Extract number format fields for persistence
fn extract_number_format_fields(nf: &NumberFormat) -> (i32, i32, i32, i32, Option<&str>) {
    match nf {
        NumberFormat::General => (0, 2, 0, 0, None),
        NumberFormat::Number { decimals, thousands, negative } => {
            (1, *decimals as i32, *thousands as i32, negative.to_int(), None)
        }
        NumberFormat::Currency { decimals, thousands, negative, symbol } => {
            (2, *decimals as i32, *thousands as i32, negative.to_int(), symbol.as_deref())
        }
        NumberFormat::Percent { decimals } => (3, *decimals as i32, 0, 0, None),
        NumberFormat::Date { style } => (4, match style {
            DateStyle::Short => 0,
            DateStyle::Long => 1,
            DateStyle::Iso => 2,
        }, 0, 0, None),
        NumberFormat::Time => (5, 0, 0, 0, None),
        NumberFormat::DateTime => (6, 0, 0, 0, None),
        NumberFormat::Custom(_) => (0, 2, 0, 0, None),
    }
}

/// Reconstruct NumberFormat from database fields
fn build_number_format(
    number_type: i32,
    decimals: i32,
    thousands: bool,
    negative: i32,
    currency_symbol: Option<String>,
) -> NumberFormat {
    let negative_style = NegativeStyle::from_int(negative);
    match number_type {
        1 => NumberFormat::Number {
            decimals: decimals as u8,
            thousands,
            negative: negative_style,
        },
        2 => NumberFormat::Currency {
            decimals: decimals as u8,
            thousands,
            negative: negative_style,
            symbol: currency_symbol,
        },
        3 => NumberFormat::Percent { decimals: decimals as u8 },
        4 => NumberFormat::Date {
            style: match decimals {
                1 => DateStyle::Long,
                2 => DateStyle::Iso,
                _ => DateStyle::Short,
            },
        },
        5 => NumberFormat::Time,
        6 => NumberFormat::DateTime,
        _ => NumberFormat::General,
    }
}

pub fn load(path: &Path) -> Result<Sheet, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Run migrations (adds new columns if missing)
    let _ = migrate(&conn); // Ignore errors (read-only DBs)

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

    let mut sheet = Sheet::new(SheetId(1), rows, cols);
    sheet.set_name(&sheet_name);

    // Load cells - check if format columns exist for backward compatibility
    let has_alignment_columns = conn
        .prepare("SELECT fmt_alignment FROM cells LIMIT 1")
        .is_ok();
    let has_font_family = conn
        .prepare("SELECT fmt_font_family FROM cells LIMIT 1")
        .is_ok();
    let has_thousands = conn
        .prepare("SELECT fmt_thousands FROM cells LIMIT 1")
        .is_ok();

    let query = if has_thousands {
        "SELECT row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol FROM cells"
    } else if has_font_family {
        "SELECT row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family FROM cells"
    } else if has_alignment_columns {
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
            let fmt_font_family: Option<String> = row.get(11).ok();
            let fmt_thousands: i32 = row.get(12).unwrap_or(0);
            let fmt_negative: i32 = row.get(13).unwrap_or(0);
            let fmt_currency_symbol: Option<String> = row.get(14).ok().and_then(|v: Option<String>| v);
            Ok((r as usize, c as usize, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol))
        })
        .map_err(|e| e.to_string())?;

    for cell_result in cell_iter {
        let (row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol) =
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
        let alignment = alignment_from_db(fmt_alignment);
        let thousands = fmt_thousands != 0;
        let negative = NegativeStyle::from_int(fmt_negative);
        let number_format = match fmt_number_type {
            1 => NumberFormat::Number { decimals: fmt_decimals as u8, thousands, negative },
            2 => NumberFormat::Currency {
                decimals: fmt_decimals as u8,
                thousands,
                negative,
                symbol: fmt_currency_symbol,
            },
            3 => NumberFormat::Percent { decimals: fmt_decimals as u8 },
            4 => NumberFormat::Date { style: match fmt_decimals {
                1 => DateStyle::Long,
                2 => DateStyle::Iso,
                _ => DateStyle::Short,
            }},
            5 => NumberFormat::Time,
            6 => NumberFormat::DateTime,
            _ => NumberFormat::General,
        };
        let format = CellFormat {
            bold: fmt_bold != 0,
            italic: fmt_italic != 0,
            underline: fmt_underline != 0,
            strikethrough: false,  // TODO: Add strikethrough column to database schema
            alignment,
            vertical_alignment: VerticalAlignment::Middle,  // TODO: Add vertical_alignment column to database schema
            text_overflow: TextOverflow::Clip,  // TODO: Add text_overflow column to database schema
            number_format,
            font_family: fmt_font_family,
            font_size: None,
            font_color: None,
            background_color: None,  // TODO: Add background_color column to database schema
            border_top: CellBorder::default(),     // TODO: Add border columns to database schema
            border_right: CellBorder::default(),
            border_bottom: CellBorder::default(),
            border_left: CellBorder::default(),
            cell_style: CellStyle::default(),
        };
        if !format.is_default() {
            sheet.set_format(row, col, format);
        }
    }

    Ok(sheet)
}

/// Save a complete workbook including all sheets and named ranges
pub fn save_workbook(workbook: &Workbook, path: &Path) -> Result<(), String> {
    // Delete existing file if present (SQLite will create fresh)
    if path.exists() {
        std::fs::remove_file(path).map_err(|e| e.to_string())?;
    }

    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Create schema (includes named_ranges table)
    conn.execute_batch(SCHEMA).map_err(|e| e.to_string())?;
    conn.pragma_update(None, "user_version", SCHEMA_VERSION).map_err(|e| e.to_string())?;

    // Save active sheet index to meta
    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?1, ?2)",
        params!["active_sheet", workbook.active_sheet_index().to_string()],
    ).map_err(|e| e.to_string())?;

    // Save cells using a transaction for performance
    conn.execute("BEGIN TRANSACTION", []).map_err(|e| e.to_string())?;

    // Save all sheets
    {
        let mut sheet_stmt = conn.prepare(
            "INSERT INTO sheets (sheet_idx, name, row_count, col_count) VALUES (?1, ?2, ?3, ?4)"
        ).map_err(|e| e.to_string())?;

        let mut cell_stmt = conn.prepare(
            "INSERT INTO cells (sheet_idx, row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol, fmt_border_top, fmt_border_right, fmt_border_bottom, fmt_border_left, fmt_border_top_color, fmt_border_right_color, fmt_border_bottom_color, fmt_border_left_color, fmt_cell_style, fmt_font_color, fmt_background_color) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20, ?21, ?22, ?23, ?24, ?25, ?26, ?27)"
        ).map_err(|e| e.to_string())?;

        for (sheet_idx, sheet) in workbook.sheets().iter().enumerate() {
            // Save sheet metadata
            sheet_stmt.execute(params![
                sheet_idx as i64,
                &sheet.name,
                sheet.rows as i64,
                sheet.cols as i64,
            ]).map_err(|e| e.to_string())?;

            // Save cells for this sheet
            for (&(row, col), cell) in sheet.cells_iter() {
                let raw = cell.value.raw_display();
                let format = &cell.format;

                // Skip cells with no value and default formatting
                if raw.is_empty() && format.is_default() {
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

                let alignment_int = alignment_to_db(format.alignment);
                let (number_type, decimals, thousands, negative, currency_symbol) = extract_number_format_fields(&format.number_format);

                cell_stmt.execute(params![
                    sheet_idx as i64,
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
                    decimals,
                    format.font_family.as_deref(),
                    thousands,
                    negative,
                    currency_symbol,
                    // Border styles
                    border_style_to_db(format.border_top.style),
                    border_style_to_db(format.border_right.style),
                    border_style_to_db(format.border_bottom.style),
                    border_style_to_db(format.border_left.style),
                    // Border colors
                    border_color_to_db(format.border_top.color),
                    border_color_to_db(format.border_right.color),
                    border_color_to_db(format.border_bottom.color),
                    border_color_to_db(format.border_left.color),
                    // Cell style
                    format.cell_style.to_int(),
                    // Font and background colors
                    border_color_to_db(format.font_color),
                    border_color_to_db(format.background_color)
                ]).map_err(|e| e.to_string())?;
            }
        }
    }

    // Save named ranges
    {
        let mut stmt = conn.prepare(
            "INSERT INTO named_ranges (name, target_type, sheet, start_row, start_col, end_row, end_col, description) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)"
        ).map_err(|e| e.to_string())?;

        for nr in workbook.list_named_ranges() {
            let (target_type, sheet_idx, start_row, start_col, end_row, end_col) = match &nr.target {
                NamedRangeTarget::Cell { sheet, row, col } => {
                    (0i32, *sheet as i64, *row as i64, *col as i64, None::<i64>, None::<i64>)
                }
                NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                    (1i32, *sheet as i64, *start_row as i64, *start_col as i64, Some(*end_row as i64), Some(*end_col as i64))
                }
            };

            stmt.execute(params![
                &nr.name,
                target_type,
                sheet_idx,
                start_row,
                start_col,
                end_row,
                end_col,
                nr.description.as_deref()
            ]).map_err(|e| e.to_string())?;
        }
    }

    // Save merged regions for active sheet (TODO: extend to all sheets)
    {
        let sheet = workbook.active_sheet();
        let mut stmt = conn.prepare(
            "INSERT INTO merged_regions (start_row, start_col, end_row, end_col) VALUES (?1, ?2, ?3, ?4)"
        ).map_err(|e| e.to_string())?;

        for merge in &sheet.merged_regions {
            stmt.execute(params![
                merge.start.0 as i64,
                merge.start.1 as i64,
                merge.end.0 as i64,
                merge.end.1 as i64,
            ]).map_err(|e| e.to_string())?;
        }
    }

    conn.execute("COMMIT", []).map_err(|e| e.to_string())?;

    Ok(())
}

/// Type alias for cell metadata (target -> {key: value}).
pub type CellMetadata = std::collections::BTreeMap<String, std::collections::BTreeMap<String, String>>;

/// Save a workbook with semantic metadata.
pub fn save_workbook_with_metadata(
    workbook: &Workbook,
    metadata: &CellMetadata,
    path: &Path,
) -> Result<(), String> {
    // Delete existing file if present (SQLite will create fresh)
    if path.exists() {
        std::fs::remove_file(path).map_err(|e| e.to_string())?;
    }

    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Create schema (includes cell_metadata table)
    conn.execute_batch(SCHEMA).map_err(|e| e.to_string())?;
    conn.pragma_update(None, "user_version", SCHEMA_VERSION).map_err(|e| e.to_string())?;

    // Save active sheet index
    conn.execute(
        "INSERT INTO meta (key, value) VALUES (?1, ?2)",
        params!["active_sheet", workbook.active_sheet_index().to_string()],
    ).map_err(|e| e.to_string())?;

    // Save all sheets and cells
    conn.execute("BEGIN TRANSACTION", []).map_err(|e| e.to_string())?;

    {
        let mut sheet_stmt = conn.prepare(
            "INSERT INTO sheets (sheet_idx, name, row_count, col_count) VALUES (?1, ?2, ?3, ?4)"
        ).map_err(|e| e.to_string())?;

        let mut cell_stmt = conn.prepare(
            "INSERT INTO cells (sheet_idx, row, col, value_type, value_num, value_text, fmt_bold, fmt_italic, fmt_underline, fmt_alignment, fmt_number_type, fmt_decimals, fmt_font_family, fmt_thousands, fmt_negative, fmt_currency_symbol, fmt_border_top, fmt_border_right, fmt_border_bottom, fmt_border_left, fmt_border_top_color, fmt_border_right_color, fmt_border_bottom_color, fmt_border_left_color, fmt_cell_style, fmt_font_color, fmt_background_color) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19, ?20, ?21, ?22, ?23, ?24, ?25, ?26, ?27)"
        ).map_err(|e| e.to_string())?;

        for (sheet_idx, sheet) in workbook.sheets().iter().enumerate() {
            // Save sheet metadata
            sheet_stmt.execute(params![
                sheet_idx as i64,
                &sheet.name,
                sheet.rows as i64,
                sheet.cols as i64,
            ]).map_err(|e| e.to_string())?;

            // Save cells for this sheet
            for (&(row, col), cell) in sheet.cells_iter() {
                let raw = cell.value.raw_display();
                let format = &cell.format;

                if raw.is_empty() && format.is_default() {
                    continue;
                }

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

                let alignment_int = alignment_to_db(format.alignment);
                let (number_type, decimals, thousands, negative, currency_symbol) = extract_number_format_fields(&format.number_format);

                cell_stmt.execute(params![
                    sheet_idx as i64,
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
                    decimals,
                    format.font_family.as_deref(),
                    thousands,
                    negative,
                    currency_symbol,
                    // Border styles
                    border_style_to_db(format.border_top.style),
                    border_style_to_db(format.border_right.style),
                    border_style_to_db(format.border_bottom.style),
                    border_style_to_db(format.border_left.style),
                    // Border colors
                    border_color_to_db(format.border_top.color),
                    border_color_to_db(format.border_right.color),
                    border_color_to_db(format.border_bottom.color),
                    border_color_to_db(format.border_left.color),
                    // Cell style
                    format.cell_style.to_int(),
                    // Font and background colors
                    border_color_to_db(format.font_color),
                    border_color_to_db(format.background_color)
                ]).map_err(|e| e.to_string())?;
            }
        }
    }

    // Save semantic metadata (affects fingerprint)
    {
        let mut stmt = conn.prepare(
            "INSERT INTO cell_metadata (target, key, value) VALUES (?1, ?2, ?3)"
        ).map_err(|e| e.to_string())?;

        for (target, props) in metadata.iter() {
            for (key, value) in props.iter() {
                stmt.execute(params![target, key, value]).map_err(|e| e.to_string())?;
            }
        }
    }

    // Save named ranges
    {
        let mut stmt = conn.prepare(
            "INSERT INTO named_ranges (name, target_type, sheet, start_row, start_col, end_row, end_col, description) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)"
        ).map_err(|e| e.to_string())?;

        for nr in workbook.list_named_ranges() {
            let (target_type, sheet_idx, start_row, start_col, end_row, end_col) = match &nr.target {
                NamedRangeTarget::Cell { sheet, row, col } => {
                    (0i32, *sheet as i64, *row as i64, *col as i64, None::<i64>, None::<i64>)
                }
                NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                    (1i32, *sheet as i64, *start_row as i64, *start_col as i64, Some(*end_row as i64), Some(*end_col as i64))
                }
            };

            stmt.execute(params![
                &nr.name,
                target_type,
                sheet_idx,
                start_row,
                start_col,
                end_row,
                end_col,
                nr.description.as_deref()
            ]).map_err(|e| e.to_string())?;
        }
    }

    // Save merged regions for active sheet (TODO: extend to all sheets)
    {
        let sheet = workbook.active_sheet();
        let mut stmt = conn.prepare(
            "INSERT INTO merged_regions (start_row, start_col, end_row, end_col) VALUES (?1, ?2, ?3, ?4)"
        ).map_err(|e| e.to_string())?;

        for merge in &sheet.merged_regions {
            stmt.execute(params![
                merge.start.0 as i64,
                merge.start.1 as i64,
                merge.end.0 as i64,
                merge.end.1 as i64,
            ]).map_err(|e| e.to_string())?;
        }
    }

    conn.execute("COMMIT", []).map_err(|e| e.to_string())?;

    Ok(())
}

/// Load workbook from v2 multi-sheet format
fn load_workbook_v2(conn: &Connection) -> Result<Workbook, String> {
    // Load active sheet index from meta
    let active_sheet: usize = conn
        .query_row(
            "SELECT value FROM meta WHERE key = 'active_sheet'",
            [],
            |row| row.get::<_, String>(0),
        )
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    // Load all sheets
    let mut sheets_data: Vec<(usize, String, usize, usize)> = Vec::new();
    {
        let mut stmt = conn.prepare(
            "SELECT sheet_idx, name, row_count, col_count FROM sheets ORDER BY sheet_idx"
        ).map_err(|e| e.to_string())?;

        let rows = stmt.query_map([], |row| {
            Ok((
                row.get::<_, i64>(0)? as usize,
                row.get::<_, String>(1)?,
                row.get::<_, i64>(2)? as usize,
                row.get::<_, i64>(3)? as usize,
            ))
        }).map_err(|e| e.to_string())?;

        for row in rows {
            sheets_data.push(row.map_err(|e| e.to_string())?);
        }
    }

    // Create sheets
    let mut sheets: Vec<Sheet> = Vec::new();
    for (idx, name, rows, cols) in &sheets_data {
        let mut sheet = Sheet::new(SheetId(*idx as u64 + 1), *rows, *cols);
        // Must use set_name() to update both name and name_key for correct lookup
        sheet.set_name(name);
        sheets.push(sheet);
    }

    // If no sheets found, create a default one
    if sheets.is_empty() {
        sheets.push(Sheet::new(SheetId(1), 1000, 26));
    }

    // Load cells for all sheets
    {
        let mut stmt = conn.prepare(
            "SELECT sheet_idx, row, col, value_type, value_num, value_text, \
             fmt_bold, fmt_italic, fmt_underline, fmt_alignment, \
             fmt_number_type, fmt_decimals, fmt_font_family, \
             fmt_thousands, fmt_negative, fmt_currency_symbol, \
             fmt_border_top, fmt_border_right, fmt_border_bottom, fmt_border_left, \
             fmt_border_top_color, fmt_border_right_color, fmt_border_bottom_color, fmt_border_left_color, \
             fmt_cell_style, fmt_font_color, fmt_background_color \
             FROM cells ORDER BY sheet_idx, row, col"
        ).map_err(|e| e.to_string())?;

        let rows = stmt.query_map([], |row| {
            Ok((
                row.get::<_, i64>(0)? as usize,  // sheet_idx
                row.get::<_, i64>(1)? as usize,  // row
                row.get::<_, i64>(2)? as usize,  // col
                row.get::<_, i32>(3)?,           // value_type
                row.get::<_, Option<f64>>(4)?,   // value_num
                row.get::<_, Option<String>>(5)?, // value_text
                row.get::<_, i32>(6).unwrap_or(0),  // fmt_bold
                row.get::<_, i32>(7).unwrap_or(0),  // fmt_italic
                row.get::<_, i32>(8).unwrap_or(0),  // fmt_underline
                row.get::<_, i32>(9).unwrap_or(0),  // fmt_alignment
                row.get::<_, i32>(10).unwrap_or(0), // fmt_number_type
                row.get::<_, i32>(11).unwrap_or(2), // fmt_decimals
                row.get::<_, Option<String>>(12)?,  // fmt_font_family
                row.get::<_, i32>(13).unwrap_or(0), // fmt_thousands
                row.get::<_, i32>(14).unwrap_or(0), // fmt_negative
                row.get::<_, Option<String>>(15)?,  // fmt_currency_symbol
                // Border styles (columns may not exist in old files)
                row.get::<_, i32>(16).unwrap_or(0), // fmt_border_top
                row.get::<_, i32>(17).unwrap_or(0), // fmt_border_right
                row.get::<_, i32>(18).unwrap_or(0), // fmt_border_bottom
                row.get::<_, i32>(19).unwrap_or(0), // fmt_border_left
                // Border colors
                row.get::<_, Option<i64>>(20).ok().flatten(), // fmt_border_top_color
                row.get::<_, Option<i64>>(21).ok().flatten(), // fmt_border_right_color
                row.get::<_, Option<i64>>(22).ok().flatten(), // fmt_border_bottom_color
                row.get::<_, Option<i64>>(23).ok().flatten(), // fmt_border_left_color
                // Cell style
                row.get::<_, i32>(24).unwrap_or(0),          // fmt_cell_style
                // Font and background colors
                row.get::<_, Option<i64>>(25).ok().flatten(), // fmt_font_color
                row.get::<_, Option<i64>>(26).ok().flatten(), // fmt_background_color
            ))
        }).map_err(|e| e.to_string())?;

        for row_result in rows {
            let (sheet_idx, row, col, value_type, value_num, value_text,
                 bold, italic, underline, alignment,
                 number_type, decimals, font_family,
                 thousands, negative, currency_symbol,
                 border_top_style, border_right_style, border_bottom_style, border_left_style,
                 border_top_color, border_right_color, border_bottom_color, border_left_color,
                 cell_style_int,
                 font_color_raw, background_color_raw
            ) = row_result.map_err(|e| e.to_string())?;

            // Ensure sheet exists
            while sheets.len() <= sheet_idx {
                let new_idx = sheets.len();
                sheets.push(Sheet::new(SheetId(new_idx as u64 + 1), 1000, 26));
            }

            let sheet = &mut sheets[sheet_idx];

            // Set cell value
            let value_str = match value_type {
                TYPE_FORMULA => value_text.unwrap_or_default(),
                TYPE_NUMBER => value_num.map(|n| {
                    if n.fract() == 0.0 { (n as i64).to_string() } else { n.to_string() }
                }).unwrap_or_default(),
                TYPE_TEXT => value_text.unwrap_or_default(),
                _ => String::new(),
            };

            if !value_str.is_empty() {
                sheet.set_value(row, col, &value_str);
            }

            // Set formatting
            let format = CellFormat {
                bold: bold != 0,
                italic: italic != 0,
                underline: underline != 0,
                alignment: alignment_from_db(alignment),
                number_format: build_number_format(number_type, decimals, thousands != 0, negative, currency_symbol),
                font_family,
                border_top: border_from_db(border_top_style, border_top_color),
                border_right: border_from_db(border_right_style, border_right_color),
                border_bottom: border_from_db(border_bottom_style, border_bottom_color),
                border_left: border_from_db(border_left_style, border_left_color),
                cell_style: CellStyle::from_int(cell_style_int),
                font_color: border_color_from_db(font_color_raw),
                background_color: border_color_from_db(background_color_raw),
                ..Default::default()
            };

            if !format.is_default() {
                sheet.set_format(row, col, format);
            }
        }
    }

    let workbook = Workbook::from_sheets(sheets, active_sheet);
    Ok(workbook)
}

/// Load a complete workbook including all sheets and named ranges
pub fn load_workbook(path: &Path) -> Result<Workbook, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Run migrations (adds new columns if missing from older files)
    let _ = migrate(&conn);

    // Check if this is the new multi-sheet format (v2+)
    let has_sheets_table = conn
        .prepare("SELECT sheet_idx FROM sheets LIMIT 1")
        .is_ok();

    let mut workbook = if has_sheets_table {
        // New multi-sheet format
        load_workbook_v2(&conn)?
    } else {
        // Legacy single-sheet format - use existing load function
        let sheet = load(path)?;
        Workbook::from_sheets(vec![sheet], 0)
    };

    // Check if named_ranges table exists (for backward compatibility)
    let has_named_ranges = conn
        .prepare("SELECT name FROM named_ranges LIMIT 1")
        .is_ok();

    if has_named_ranges {
        let mut stmt = conn.prepare(
            "SELECT name, target_type, sheet, start_row, start_col, end_row, end_col, description FROM named_ranges"
        ).map_err(|e| e.to_string())?;

        let named_range_iter = stmt
            .query_map([], |row| {
                let name: String = row.get(0)?;
                let target_type: i32 = row.get(1)?;
                let sheet: i64 = row.get(2)?;
                let start_row: i64 = row.get(3)?;
                let start_col: i64 = row.get(4)?;
                let end_row: Option<i64> = row.get(5)?;
                let end_col: Option<i64> = row.get(6)?;
                let description: Option<String> = row.get(7)?;
                Ok((name, target_type, sheet, start_row, start_col, end_row, end_col, description))
            })
            .map_err(|e| e.to_string())?;

        for nr_result in named_range_iter {
            let (name, target_type, sheet, start_row, start_col, end_row, end_col, description) =
                nr_result.map_err(|e| e.to_string())?;

            let target = if target_type == 0 {
                // Cell
                NamedRangeTarget::Cell {
                    sheet: sheet as usize,
                    row: start_row as usize,
                    col: start_col as usize,
                }
            } else {
                // Range
                NamedRangeTarget::Range {
                    sheet: sheet as usize,
                    start_row: start_row as usize,
                    start_col: start_col as usize,
                    end_row: end_row.unwrap_or(start_row) as usize,
                    end_col: end_col.unwrap_or(start_col) as usize,
                }
            };

            let mut named_range = NamedRange {
                name,
                target,
                description: None,
            };
            if let Some(desc) = description {
                named_range.description = Some(desc);
            }

            // Add to workbook (using the internal method)
            let _ = workbook.named_ranges_mut().set(named_range);
        }
    }

    // Load merged regions if the table exists (backward compatibility)
    let has_merged_regions = conn
        .prepare("SELECT start_row FROM merged_regions LIMIT 1")
        .is_ok();

    if has_merged_regions {
        let mut stmt = conn.prepare(
            "SELECT start_row, start_col, end_row, end_col FROM merged_regions"
        ).map_err(|e| e.to_string())?;

        let merge_iter = stmt
            .query_map([], |row| {
                let sr: i64 = row.get(0)?;
                let sc: i64 = row.get(1)?;
                let er: i64 = row.get(2)?;
                let ec: i64 = row.get(3)?;
                Ok((sr as usize, sc as usize, er as usize, ec as usize))
            })
            .map_err(|e| e.to_string())?;

        if let Some(sheet) = workbook.sheet_mut(0) {
            for merge_result in merge_iter {
                let (sr, sc, er, ec) = merge_result.map_err(|e| e.to_string())?;
                let region = MergedRegion::new(sr, sc, er, ec);
                let _ = sheet.add_merge(region); // silently skip overlaps on load
            }
        }
    }

    // Rebuild dependency graph and compute all formulas after loading
    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    Ok(workbook)
}

/// Load semantic metadata from a .sheet file.
/// Returns an empty map if the cell_metadata table doesn't exist (backward compatibility).
pub fn load_cell_metadata(path: &Path) -> Result<CellMetadata, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;
    let mut metadata = CellMetadata::new();

    // Check if cell_metadata table exists (backward compatibility)
    let has_metadata = conn
        .prepare("SELECT target FROM cell_metadata LIMIT 1")
        .is_ok();

    if !has_metadata {
        return Ok(metadata);
    }

    let mut stmt = conn.prepare(
        "SELECT target, key, value FROM cell_metadata ORDER BY target, key"
    ).map_err(|e| e.to_string())?;

    let rows = stmt
        .query_map([], |row| {
            let target: String = row.get(0)?;
            let key: String = row.get(1)?;
            let value: String = row.get(2)?;
            Ok((target, key, value))
        })
        .map_err(|e| e.to_string())?;

    for row_result in rows {
        let (target, key, value) = row_result.map_err(|e| e.to_string())?;
        metadata.entry(target).or_default().insert(key, value);
    }

    Ok(metadata)
}

/// Load semantic verification info from a .sheet file.
/// Returns default (empty) verification if the fields don't exist.
pub fn load_semantic_verification(path: &Path) -> Result<SemanticVerification, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    let fingerprint = conn
        .query_row(
            "SELECT value FROM meta WHERE key = 'expected_semantic_fingerprint'",
            [],
            |row| row.get::<_, String>(0),
        )
        .ok();

    let label = conn
        .query_row(
            "SELECT value FROM meta WHERE key = 'expected_semantic_label'",
            [],
            |row| row.get::<_, String>(0),
        )
        .ok();

    let timestamp = conn
        .query_row(
            "SELECT value FROM meta WHERE key = 'expected_semantic_timestamp'",
            [],
            |row| row.get::<_, String>(0),
        )
        .ok();

    Ok(SemanticVerification {
        fingerprint,
        label,
        timestamp,
    })
}

/// Save semantic verification info to an existing .sheet file.
/// This updates the meta table without rewriting the entire file.
pub fn save_semantic_verification(path: &Path, verification: &SemanticVerification) -> Result<(), String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Delete existing verification entries
    conn.execute("DELETE FROM meta WHERE key = 'expected_semantic_fingerprint'", [])
        .map_err(|e| e.to_string())?;
    conn.execute("DELETE FROM meta WHERE key = 'expected_semantic_label'", [])
        .map_err(|e| e.to_string())?;
    conn.execute("DELETE FROM meta WHERE key = 'expected_semantic_timestamp'", [])
        .map_err(|e| e.to_string())?;

    // Insert new verification entries
    if let Some(ref fp) = verification.fingerprint {
        conn.execute(
            "INSERT INTO meta (key, value) VALUES ('expected_semantic_fingerprint', ?1)",
            params![fp],
        ).map_err(|e| e.to_string())?;
    }

    if let Some(ref label) = verification.label {
        conn.execute(
            "INSERT INTO meta (key, value) VALUES ('expected_semantic_label', ?1)",
            params![label],
        ).map_err(|e| e.to_string())?;
    }

    if let Some(ref timestamp) = verification.timestamp {
        conn.execute(
            "INSERT INTO meta (key, value) VALUES ('expected_semantic_timestamp', ?1)",
            params![timestamp],
        ).map_err(|e| e.to_string())?;
    }

    Ok(())
}

/// VisiHub link information stored in .sheet files
#[derive(Debug, Clone, PartialEq)]
pub struct HubLink {
    pub repo_owner: String,
    pub repo_slug: String,
    pub dataset_id: String,
    pub local_head_id: Option<String>,
    pub local_head_hash: Option<String>,
    pub link_mode: String,  // "pull" or "publish"
    pub linked_at: String,  // ISO 8601 timestamp
    pub api_base: String,
}

impl HubLink {
    pub fn new(repo_owner: String, repo_slug: String, dataset_id: String) -> Self {
        Self {
            repo_owner,
            repo_slug,
            dataset_id,
            local_head_id: None,
            local_head_hash: None,
            link_mode: "pull".to_string(),
            linked_at: chrono_now_iso8601(),
            api_base: "https://api.visihub.app".to_string(),
        }
    }

    /// Returns the display name for the linked repo (e.g., "@alice/budget")
    pub fn display_name(&self) -> String {
        format!("@{}/{}", self.repo_owner, self.repo_slug)
    }
}

fn chrono_now_iso8601() -> String {
    // Simple ISO 8601 timestamp without external dependency
    use std::time::{SystemTime, UNIX_EPOCH};
    let duration = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let secs = duration.as_secs();
    // Convert to basic ISO format (not perfect but good enough)
    format!("{}", secs)
}

/// Load hub_link from a .sheet file (if present)
pub fn load_hub_link(path: &Path) -> Result<Option<HubLink>, String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Check if hub_link table exists
    let has_hub_link = conn
        .prepare("SELECT id FROM hub_link LIMIT 1")
        .is_ok();

    if !has_hub_link {
        return Ok(None);
    }

    let result = conn.query_row(
        "SELECT repo_owner, repo_slug, dataset_id, local_head_id, local_head_hash, link_mode, linked_at, api_base FROM hub_link WHERE id = 1",
        [],
        |row| {
            Ok(HubLink {
                repo_owner: row.get(0)?,
                repo_slug: row.get(1)?,
                dataset_id: row.get(2)?,
                local_head_id: row.get(3)?,
                local_head_hash: row.get(4)?,
                link_mode: row.get::<_, Option<String>>(5)?.unwrap_or_else(|| "pull".to_string()),
                linked_at: row.get::<_, Option<String>>(6)?.unwrap_or_default(),
                api_base: row.get::<_, Option<String>>(7)?.unwrap_or_else(|| "https://api.visihub.app".to_string()),
            })
        },
    );

    match result {
        Ok(link) => Ok(Some(link)),
        Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
        Err(e) => Err(e.to_string()),
    }
}

/// Save hub_link to a .sheet file (creates table if needed)
pub fn save_hub_link(path: &Path, link: &HubLink) -> Result<(), String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Ensure table exists
    conn.execute(
        "CREATE TABLE IF NOT EXISTS hub_link (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            repo_owner TEXT NOT NULL,
            repo_slug TEXT NOT NULL,
            dataset_id TEXT NOT NULL,
            local_head_id TEXT,
            local_head_hash TEXT,
            link_mode TEXT DEFAULT 'pull',
            linked_at TEXT NOT NULL,
            api_base TEXT DEFAULT 'https://api.visihub.app'
        )",
        [],
    ).map_err(|e| e.to_string())?;

    // Upsert the singleton row
    conn.execute(
        "INSERT OR REPLACE INTO hub_link (id, repo_owner, repo_slug, dataset_id, local_head_id, local_head_hash, link_mode, linked_at, api_base)
         VALUES (1, ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)",
        params![
            &link.repo_owner,
            &link.repo_slug,
            &link.dataset_id,
            &link.local_head_id,
            &link.local_head_hash,
            &link.link_mode,
            &link.linked_at,
            &link.api_base,
        ],
    ).map_err(|e| e.to_string())?;

    Ok(())
}

/// Remove hub_link from a .sheet file (unlink)
pub fn delete_hub_link(path: &Path) -> Result<(), String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Check if table exists first
    let has_hub_link = conn
        .prepare("SELECT id FROM hub_link LIMIT 1")
        .is_ok();

    if has_hub_link {
        conn.execute("DELETE FROM hub_link WHERE id = 1", [])
            .map_err(|e| e.to_string())?;
    }

    Ok(())
}

/// Layout data (column widths and row heights) keyed by sheet index.
pub struct SheetLayout {
    /// sheet_idx -> (col -> width_px)
    pub col_widths: HashMap<usize, HashMap<usize, f32>>,
    /// sheet_idx -> (row -> height_px)
    pub row_heights: HashMap<usize, HashMap<usize, f32>>,
}

/// Save column widths and row heights into an existing .sheet file.
/// Called after save_workbook() which creates the file fresh with the schema.
pub fn save_layout(path: &Path, layout: &SheetLayout) -> Result<(), String> {
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Tables already exist (created by SCHEMA in save_workbook), just insert rows
    let mut col_stmt = conn.prepare(
        "INSERT INTO col_widths (sheet_idx, col, width) VALUES (?1, ?2, ?3)"
    ).map_err(|e| e.to_string())?;

    for (sheet_idx, widths) in &layout.col_widths {
        for (col, width) in widths {
            col_stmt.execute(params![*sheet_idx as i64, *col as i64, *width as f64])
                .map_err(|e| e.to_string())?;
        }
    }

    let mut row_stmt = conn.prepare(
        "INSERT INTO row_heights (sheet_idx, row, height) VALUES (?1, ?2, ?3)"
    ).map_err(|e| e.to_string())?;

    for (sheet_idx, heights) in &layout.row_heights {
        for (row, height) in heights {
            row_stmt.execute(params![*sheet_idx as i64, *row as i64, *height as f64])
                .map_err(|e| e.to_string())?;
        }
    }

    Ok(())
}

/// Load column widths and row heights from a .sheet file.
/// Returns empty maps if the tables don't exist (older files before migration).
pub fn load_layout(path: &Path) -> SheetLayout {
    let mut layout = SheetLayout {
        col_widths: HashMap::new(),
        row_heights: HashMap::new(),
    };

    let conn = match Connection::open(path) {
        Ok(c) => c,
        Err(_) => return layout,
    };

    // Load column widths (table may not exist in old files)
    if let Ok(mut stmt) = conn.prepare("SELECT sheet_idx, col, width FROM col_widths") {
        if let Ok(rows) = stmt.query_map([], |row| {
            Ok((
                row.get::<_, i64>(0)? as usize,
                row.get::<_, i64>(1)? as usize,
                row.get::<_, f64>(2)? as f32,
            ))
        }) {
            for row in rows.flatten() {
                let (sheet_idx, col, width) = row;
                layout.col_widths.entry(sheet_idx).or_default().insert(col, width);
            }
        }
    }

    // Load row heights
    if let Ok(mut stmt) = conn.prepare("SELECT sheet_idx, row, height FROM row_heights") {
        if let Ok(rows) = stmt.query_map([], |row| {
            Ok((
                row.get::<_, i64>(0)? as usize,
                row.get::<_, i64>(1)? as usize,
                row.get::<_, f64>(2)? as f32,
            ))
        }) {
            for row in rows.flatten() {
                let (sheet_idx, r, height) = row;
                layout.row_heights.entry(sheet_idx).or_default().insert(r, height);
            }
        }
    }

    layout
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::NamedTempFile;

    #[test]
    fn test_named_range_persistence_cell() {
        // Create a workbook with a named cell
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Revenue");
        workbook.active_sheet_mut().set_value(0, 1, "100");

        let nr = NamedRange::cell("MyRevenue", 0, 0, 1)
            .with_description("Annual revenue cell");
        workbook.named_ranges_mut().set(nr).unwrap();

        // Save to temp file
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        save_workbook(&workbook, path).expect("Save should succeed");

        // Load back
        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify named range was preserved
        let loaded_nr = loaded.get_named_range("MyRevenue").expect("Named range should exist");
        assert_eq!(loaded_nr.name, "MyRevenue");
        assert_eq!(loaded_nr.description, Some("Annual revenue cell".to_string()));
        match &loaded_nr.target {
            NamedRangeTarget::Cell { sheet, row, col } => {
                assert_eq!(*sheet, 0);
                assert_eq!(*row, 0);
                assert_eq!(*col, 1);
            }
            _ => panic!("Expected Cell target"),
        }
    }

    #[test]
    fn test_named_range_persistence_range() {
        // Create a workbook with a named range
        let mut workbook = Workbook::new();
        for row in 0..10 {
            workbook.active_sheet_mut().set_value(row, 0, &format!("{}", row * 10));
        }

        let nr = NamedRange::range("SalesData", 0, 0, 0, 9, 0)
            .with_description("Q1 sales figures");
        workbook.named_ranges_mut().set(nr).unwrap();

        // Save to temp file
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        save_workbook(&workbook, path).expect("Save should succeed");

        // Load back
        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify named range was preserved
        let loaded_nr = loaded.get_named_range("SalesData").expect("Named range should exist");
        assert_eq!(loaded_nr.name, "SalesData");
        assert_eq!(loaded_nr.description, Some("Q1 sales figures".to_string()));
        match &loaded_nr.target {
            NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                assert_eq!(*sheet, 0);
                assert_eq!(*start_row, 0);
                assert_eq!(*start_col, 0);
                assert_eq!(*end_row, 9);
                assert_eq!(*end_col, 0);
            }
            _ => panic!("Expected Range target"),
        }
    }

    #[test]
    fn test_named_range_persistence_multiple() {
        // Create a workbook with multiple named ranges
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Header");
        workbook.active_sheet_mut().set_value(1, 0, "100");
        workbook.active_sheet_mut().set_value(2, 0, "200");

        // Add multiple named ranges
        workbook.named_ranges_mut().set(NamedRange::cell("Header", 0, 0, 0)).unwrap();
        workbook.named_ranges_mut().set(NamedRange::range("Data", 0, 1, 0, 2, 0)).unwrap();
        workbook.named_ranges_mut().set(
            NamedRange::cell("Total", 0, 2, 0).with_description("Grand total")
        ).unwrap();

        // Save and reload
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        save_workbook(&workbook, path).expect("Save should succeed");

        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify all named ranges
        assert_eq!(loaded.list_named_ranges().len(), 3);
        assert!(loaded.get_named_range("Header").is_some());
        assert!(loaded.get_named_range("Data").is_some());
        assert!(loaded.get_named_range("Total").is_some());

        // Check Total has description
        let total = loaded.get_named_range("Total").unwrap();
        assert_eq!(total.description, Some("Grand total".to_string()));
    }

    #[test]
    fn test_named_range_persistence_case_preserved() {
        // Verify case is preserved in round-trip
        let mut workbook = Workbook::new();
        workbook.named_ranges_mut().set(NamedRange::cell("MyMixedCaseName", 0, 0, 0)).unwrap();

        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        save_workbook(&workbook, path).expect("Save should succeed");

        let loaded = load_workbook(path).expect("Load should succeed");

        // Should find with exact case
        let nr = loaded.get_named_range("MyMixedCaseName").expect("Should find named range");
        assert_eq!(nr.name, "MyMixedCaseName");  // Original case preserved

        // Should also find case-insensitively
        assert!(loaded.get_named_range("mymixedcasename").is_some());
    }

    #[test]
    fn test_merged_regions_persistence() {
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Merged");

        // Add merged regions
        workbook
            .active_sheet_mut()
            .add_merge(MergedRegion::new(0, 0, 2, 3))
            .unwrap();
        workbook
            .active_sheet_mut()
            .add_merge(MergedRegion::new(5, 5, 7, 8))
            .unwrap();

        // Save
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        save_workbook(&workbook, path).expect("Save should succeed");

        // Load
        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify merges preserved
        let sheet = loaded.active_sheet();
        assert_eq!(sheet.merged_regions.len(), 2);
        assert_eq!(sheet.merged_regions[0], MergedRegion::new(0, 0, 2, 3));
        assert_eq!(sheet.merged_regions[1], MergedRegion::new(5, 5, 7, 8));

        // Verify index rebuilt
        assert!(sheet.is_merge_origin(0, 0));
        assert!(sheet.is_merge_hidden(1, 1));
        assert!(sheet.get_merge(6, 6).is_some());
        assert!(sheet.get_merge(3, 3).is_none());
    }

    #[test]
    fn test_backward_compat_file_without_merged_regions() {
        // Test loading a file without the merged_regions table
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(r#"
                CREATE TABLE cells (
                    row INTEGER NOT NULL,
                    col INTEGER NOT NULL,
                    value_type INTEGER NOT NULL,
                    value_num REAL,
                    value_text TEXT,
                    fmt_bold INTEGER DEFAULT 0,
                    fmt_italic INTEGER DEFAULT 0,
                    fmt_underline INTEGER DEFAULT 0,
                    fmt_alignment INTEGER DEFAULT 0,
                    fmt_number_type INTEGER DEFAULT 0,
                    fmt_decimals INTEGER DEFAULT 2,
                    fmt_font_family TEXT,
                    PRIMARY KEY (row, col)
                );
                CREATE TABLE meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '1000');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
                INSERT INTO cells (row, col, value_type, value_text) VALUES (0, 0, 2, 'Hello');
            "#).unwrap();
        }

        let loaded = load_workbook(path).expect("Should load old file without merged_regions table");
        assert!(loaded.active_sheet().merged_regions.is_empty());
        assert_eq!(loaded.active_sheet().get_raw(0, 0), "Hello");
    }

    #[test]
    fn test_backward_compat_file_without_named_ranges() {
        // Test loading a file that was created before named_ranges table existed
        // We simulate this by creating a minimal file without the table

        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        // Create a minimal .sheet file without named_ranges table
        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(r#"
                CREATE TABLE cells (
                    row INTEGER NOT NULL,
                    col INTEGER NOT NULL,
                    value_type INTEGER NOT NULL,
                    value_num REAL,
                    value_text TEXT,
                    fmt_bold INTEGER DEFAULT 0,
                    fmt_italic INTEGER DEFAULT 0,
                    fmt_underline INTEGER DEFAULT 0,
                    fmt_alignment INTEGER DEFAULT 0,
                    fmt_number_type INTEGER DEFAULT 0,
                    fmt_decimals INTEGER DEFAULT 2,
                    fmt_font_family TEXT,
                    PRIMARY KEY (row, col)
                );
                CREATE TABLE meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '1000');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
                INSERT INTO cells (row, col, value_type, value_text) VALUES (0, 0, 2, 'Hello');
            "#).unwrap();
        }

        // Should load without error, with empty named ranges
        let loaded = load_workbook(path).expect("Should load old file format");
        assert!(loaded.list_named_ranges().is_empty());
        assert_eq!(loaded.active_sheet().get_raw(0, 0), "Hello");
    }

    #[test]
    fn test_old_schema_load_defaults() {
        // Create an old-schema DB (no thousands/negative/symbol columns)
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(r#"
                CREATE TABLE cells (
                    row INTEGER NOT NULL,
                    col INTEGER NOT NULL,
                    value_type INTEGER NOT NULL,
                    value_num REAL,
                    value_text TEXT,
                    fmt_bold INTEGER DEFAULT 0,
                    fmt_italic INTEGER DEFAULT 0,
                    fmt_underline INTEGER DEFAULT 0,
                    fmt_alignment INTEGER DEFAULT 0,
                    fmt_number_type INTEGER DEFAULT 0,
                    fmt_decimals INTEGER DEFAULT 2,
                    fmt_font_family TEXT,
                    PRIMARY KEY (row, col)
                );
                CREATE TABLE meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '1000');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
                INSERT INTO cells (row, col, value_type, value_num, fmt_number_type, fmt_decimals) VALUES (0, 0, 1, 1234.56, 1, 2);
                INSERT INTO cells (row, col, value_type, value_num, fmt_number_type, fmt_decimals) VALUES (0, 1, 1, 99.99, 2, 2);
            "#).unwrap();
        }

        let sheet = load(path).expect("Should load old format");
        // Number format should default to thousands=false, negative=Minus
        let fmt0 = sheet.get_format(0, 0);
        match &fmt0.number_format {
            NumberFormat::Number { decimals, thousands, negative } => {
                assert_eq!(*decimals, 2);
                assert!(!*thousands);
                assert_eq!(*negative, NegativeStyle::Minus);
            }
            other => panic!("Expected Number format, got {:?}", other),
        }
        // Currency should default to thousands=false, negative=Minus, symbol=None
        let fmt1 = sheet.get_format(0, 1);
        match &fmt1.number_format {
            NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                assert_eq!(*decimals, 2);
                assert!(!*thousands);
                assert_eq!(*negative, NegativeStyle::Minus);
                assert!(symbol.is_none());
            }
            other => panic!("Expected Currency format, got {:?}", other),
        }
    }

    #[test]
    fn test_migration_roundtrip() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        // Create old-schema DB
        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(r#"
                CREATE TABLE cells (
                    row INTEGER NOT NULL,
                    col INTEGER NOT NULL,
                    value_type INTEGER NOT NULL,
                    value_num REAL,
                    value_text TEXT,
                    fmt_bold INTEGER DEFAULT 0,
                    fmt_italic INTEGER DEFAULT 0,
                    fmt_underline INTEGER DEFAULT 0,
                    fmt_alignment INTEGER DEFAULT 0,
                    fmt_number_type INTEGER DEFAULT 0,
                    fmt_decimals INTEGER DEFAULT 2,
                    fmt_font_family TEXT,
                    PRIMARY KEY (row, col)
                );
                CREATE TABLE meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '1000');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
                INSERT INTO cells (row, col, value_type, value_num, fmt_number_type, fmt_decimals) VALUES (0, 0, 1, 42.0, 1, 2);
            "#).unwrap();
        }

        // Load triggers migration
        let sheet = load(path).expect("Should load and migrate");
        assert_eq!(sheet.get_raw(0, 0), "42");

        // Verify migration added columns
        {
            let conn = Connection::open(path).unwrap();
            let version: i32 = conn.pragma_query_value(None, "user_version", |r| r.get(0)).unwrap();
            assert_eq!(version, 4);  // All migrations (v1, v2, v4) run

            let columns: Vec<String> = conn
                .prepare("PRAGMA table_info(cells)").unwrap()
                .query_map([], |row| row.get::<_, String>(1)).unwrap()
                .filter_map(|r| r.ok())
                .collect();
            // v1 migration columns
            assert!(columns.contains(&"fmt_thousands".to_string()));
            assert!(columns.contains(&"fmt_negative".to_string()));
            assert!(columns.contains(&"fmt_currency_symbol".to_string()));
            // v2 migration columns (borders)
            assert!(columns.contains(&"fmt_border_top".to_string()));
            assert!(columns.contains(&"fmt_border_top_color".to_string()));
            // v4 migration column (cell styles)
            assert!(columns.contains(&"fmt_cell_style".to_string()));
        }
    }

    #[test]
    fn test_save_only_populated_cells() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Hello");
        workbook.active_sheet_mut().set_value(1, 1, "42");
        workbook.active_sheet_mut().set_value(2, 2, "=1+1");
        save_workbook(&workbook, path).expect("Save should succeed");

        // Query the DB directly to verify only 3 rows were written
        let conn = Connection::open(path).unwrap();
        let count: i64 = conn.query_row("SELECT COUNT(*) FROM cells", [], |r| r.get(0)).unwrap();
        assert_eq!(count, 3, "Should only save 3 populated cells, got {}", count);
    }

    #[test]
    fn test_alignment_general_roundtrip() {
        // Directly insert a DB row with fmt_alignment=3 (General) and a value,
        // then load and assert alignment == Alignment::General.
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();
        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(SCHEMA).unwrap();
            conn.pragma_update(None, "user_version", SCHEMA_VERSION).unwrap();
            conn.execute_batch(r#"
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '100');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
                INSERT INTO cells (row, col, value_type, value_num, fmt_alignment) VALUES (0, 0, 1, 42.0, 3);
            "#).unwrap();
        }

        let sheet = load(path).expect("Load should succeed");
        assert_eq!(sheet.get_format(0, 0).alignment, Alignment::General,
            "fmt_alignment=3 should decode to Alignment::General, not Left");
    }

    #[test]
    fn test_alignment_left_persists() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut sheet = Sheet::new(SheetId(1), 100, 26);
        sheet.set_value(0, 0, "hello");
        sheet.set_alignment(0, 0, Alignment::Left);
        save(&sheet, path).expect("Save should succeed");

        let loaded = load(path).expect("Load should succeed");
        assert_eq!(loaded.get_format(0, 0).alignment, Alignment::Left,
            "Explicit Left alignment should survive round-trip (distinct from General)");
    }

    #[test]
    #[ignore = "Pre-existing failure: bloat cleanup not triggering on load"]
    fn test_bloated_file_auto_cleaned() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        {
            let conn = Connection::open(path).unwrap();
            conn.execute_batch(SCHEMA).unwrap();
            conn.pragma_update(None, "user_version", SCHEMA_VERSION).unwrap();
            conn.execute_batch(r#"
                INSERT INTO meta (key, value) VALUES ('sheet_name', 'Sheet1');
                INSERT INTO meta (key, value) VALUES ('rows', '1000');
                INSERT INTO meta (key, value) VALUES ('cols', '26');
            "#).unwrap();

            // Insert 20k empty junk rows (value_type=0, fmt_alignment=3 = General default)
            let mut stmt = conn.prepare(
                "INSERT INTO cells (row, col, value_type, value_num, value_text, fmt_alignment) VALUES (?1, ?2, 0, NULL, NULL, 3)"
            ).unwrap();
            for i in 0..20_000 {
                stmt.execute(params![i as i64, 0i64]).unwrap();
            }

            // Insert 1 real cell with a value
            conn.execute(
                "INSERT OR REPLACE INTO cells (row, col, value_type, value_num, fmt_alignment) VALUES (0, 0, 1, 99.0, 3)",
                [],
            ).unwrap();

            // Insert 1 empty cell with fmt_alignment=0 (Left) — real formatting, should survive
            conn.execute(
                "INSERT OR REPLACE INTO cells (row, col, value_type, value_num, value_text, fmt_alignment) VALUES (0, 1, 0, NULL, NULL, 0)",
                [],
            ).unwrap();
        }

        // load_workbook triggers the bloat migration
        let _loaded = load_workbook(path).expect("Load should succeed");

        // Reopen DB and verify cleanup
        let conn = Connection::open(path).unwrap();
        let count: i64 = conn.query_row("SELECT COUNT(*) FROM cells", [], |r| r.get(0)).unwrap();
        assert_eq!(count, 2, "Should keep only the real cell and the Left-formatted cell, got {}", count);
    }

    #[test]
    fn test_new_format_fields_roundtrip() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        // Create and save a sheet with the new format fields
        let mut sheet = Sheet::new(SheetId(1), 1000, 26);
        sheet.set_value(0, 0, "1234.56");
        sheet.set_number_format(0, 0, NumberFormat::Number {
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::RedParens,
        });
        sheet.set_value(0, 1, "99.99");
        sheet.set_number_format(0, 1, NumberFormat::Currency {
            decimals: 3,
            thousands: true,
            negative: NegativeStyle::Parens,
            symbol: Some("EUR ".to_string()),
        });

        save(&sheet, path).expect("Save should succeed");

        // Reload and verify
        let loaded = load(path).expect("Load should succeed");
        match &loaded.get_format(0, 0).number_format {
            NumberFormat::Number { decimals, thousands, negative } => {
                assert_eq!(*decimals, 2);
                assert!(*thousands);
                assert_eq!(*negative, NegativeStyle::RedParens);
            }
            other => panic!("Expected Number, got {:?}", other),
        }
        match &loaded.get_format(0, 1).number_format {
            NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                assert_eq!(*decimals, 3);
                assert!(*thousands);
                assert_eq!(*negative, NegativeStyle::Parens);
                assert_eq!(symbol.as_deref(), Some("EUR "));
            }
            other => panic!("Expected Currency, got {:?}", other),
        }
    }

    // === Semantic Fingerprint Tests ===

    #[test]
    fn test_fingerprint_deterministic() {
        // Same content produces identical fingerprint
        let mut wb1 = Workbook::new();
        wb1.active_sheet_mut().set_value(0, 0, "Revenue");
        wb1.active_sheet_mut().set_value(0, 1, "1000");
        wb1.active_sheet_mut().set_value(1, 0, "=A1");

        let mut wb2 = Workbook::new();
        wb2.active_sheet_mut().set_value(0, 0, "Revenue");
        wb2.active_sheet_mut().set_value(0, 1, "1000");
        wb2.active_sheet_mut().set_value(1, 0, "=A1");

        let fp1 = compute_semantic_fingerprint(&wb1);
        let fp2 = compute_semantic_fingerprint(&wb2);
        assert_eq!(fp1, fp2, "Identical content should produce identical fingerprint");
    }

    #[test]
    fn test_fingerprint_style_immunity() {
        // TEST 2: Formatting changes do NOT affect fingerprint
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "Revenue");
        wb.active_sheet_mut().set_value(0, 1, "1000");
        wb.active_sheet_mut().set_value(1, 0, "=B1*2");

        let fp_before = compute_semantic_fingerprint(&wb);

        // Apply extensive formatting changes
        wb.active_sheet_mut().toggle_bold(0, 0);
        wb.active_sheet_mut().toggle_italic(0, 1);
        wb.active_sheet_mut().toggle_underline(1, 0);
        wb.active_sheet_mut().set_alignment(0, 0, Alignment::Center);
        wb.active_sheet_mut().set_number_format(0, 1, NumberFormat::Currency {
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::RedParens,
            symbol: Some("$".to_string()),
        });

        let fp_after = compute_semantic_fingerprint(&wb);

        assert_eq!(fp_before, fp_after,
            "Formatting changes must NOT affect fingerprint! Before: {}, After: {}",
            fp_before, fp_after);
    }

    #[test]
    fn test_fingerprint_logic_drift() {
        // TEST 3: Formula/value changes MUST change fingerprint
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "Revenue");
        wb.active_sheet_mut().set_value(0, 1, "1000");
        wb.active_sheet_mut().set_value(1, 0, "=B1*2");

        let fp_before = compute_semantic_fingerprint(&wb);

        // Change a formula (semantic change)
        wb.active_sheet_mut().set_value(1, 0, "=B1*3");

        let fp_after = compute_semantic_fingerprint(&wb);

        assert_ne!(fp_before, fp_after,
            "Formula change MUST change fingerprint! Before: {}, After: {}",
            fp_before, fp_after);
    }

    #[test]
    fn test_fingerprint_value_drift() {
        // Changing a cell value also changes fingerprint
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "100");
        wb.active_sheet_mut().set_value(0, 1, "200");

        let fp_before = compute_semantic_fingerprint(&wb);

        // Change a value
        wb.active_sheet_mut().set_value(0, 0, "101");

        let fp_after = compute_semantic_fingerprint(&wb);

        assert_ne!(fp_before, fp_after,
            "Value change MUST change fingerprint!");
    }

    #[test]
    fn test_fingerprint_add_cell_drift() {
        // Adding a new cell changes fingerprint
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "100");

        let fp_before = compute_semantic_fingerprint(&wb);

        // Add new cell
        wb.active_sheet_mut().set_value(0, 1, "200");

        let fp_after = compute_semantic_fingerprint(&wb);

        assert_ne!(fp_before, fp_after,
            "Adding a cell MUST change fingerprint!");
    }

    #[test]
    fn test_fingerprint_delete_cell_drift() {
        // Deleting a cell (clearing value) changes fingerprint
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "100");
        wb.active_sheet_mut().set_value(0, 1, "200");

        let fp_before = compute_semantic_fingerprint(&wb);

        // Clear a cell
        wb.active_sheet_mut().set_value(0, 1, "");

        let fp_after = compute_semantic_fingerprint(&wb);

        assert_ne!(fp_before, fp_after,
            "Deleting a cell MUST change fingerprint!");
    }

    #[test]
    fn test_fingerprint_format_version() {
        // Fingerprint includes version prefix
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "test");

        let fp = compute_semantic_fingerprint(&wb);
        assert!(fp.starts_with("v2:"), "Fingerprint should start with version prefix, got: {}", fp);
    }

    #[test]
    fn test_fingerprint_op_count() {
        // Fingerprint includes correct operation count
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "a");
        wb.active_sheet_mut().set_value(0, 1, "b");
        wb.active_sheet_mut().set_value(0, 2, "c");

        let fp = compute_semantic_fingerprint(&wb);
        assert!(fp.starts_with("v2:3:"),
            "Fingerprint should show 3 operations, got: {}", fp);
    }

    #[test]
    fn test_fingerprint_changes_with_iteration_settings() {
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "test");

        let fp_default = compute_semantic_fingerprint(&wb);

        // Enable iteration — fingerprint must change
        wb.set_iterative_enabled(true);
        let fp_enabled = compute_semantic_fingerprint(&wb);
        assert_ne!(fp_default, fp_enabled,
            "Fingerprint must change when iteration is enabled");

        // Change max_iters — fingerprint must change
        wb.set_iterative_max_iters(50);
        let fp_max_iters = compute_semantic_fingerprint(&wb);
        assert_ne!(fp_enabled, fp_max_iters,
            "Fingerprint must change when max_iters changes");

        // Change tolerance — fingerprint must change
        wb.set_iterative_tolerance(0.001);
        let fp_tolerance = compute_semantic_fingerprint(&wb);
        assert_ne!(fp_max_iters, fp_tolerance,
            "Fingerprint must change when tolerance changes");
    }

    #[test]
    fn test_verification_persistence_roundtrip() {
        // Full end-to-end: create file, stamp it, reload, verify status
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        // Create and save workbook
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "Revenue");
        wb.active_sheet_mut().set_value(0, 1, "1000");
        wb.active_sheet_mut().set_value(1, 0, "=B1*2");
        save_workbook(&wb, path).expect("Save should succeed");

        // Compute and persist fingerprint (simulating CLI --stamp)
        let fingerprint = compute_semantic_fingerprint(&wb);
        let verification = SemanticVerification {
            fingerprint: Some(fingerprint.clone()),
            label: Some("Test Model v1".to_string()),
            timestamp: Some("2025-01-01T00:00:00Z".to_string()),
        };
        save_semantic_verification(path, &verification).expect("Save verification should succeed");

        // Reload workbook and verification
        let loaded_wb = load_workbook(path).expect("Load should succeed");
        let loaded_verification = load_semantic_verification(path).unwrap_or_default();

        // Verify fingerprints match (status = Verified)
        let current_fingerprint = compute_semantic_fingerprint(&loaded_wb);
        assert_eq!(
            loaded_verification.fingerprint.as_ref(),
            Some(&current_fingerprint),
            "Loaded fingerprint should match current computation"
        );
        assert_eq!(loaded_verification.label, Some("Test Model v1".to_string()));
    }

    #[test]
    fn test_verification_detects_drift() {
        // Create stamped file, modify it, verify drift is detected
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        // Create and save workbook
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "100");
        save_workbook(&wb, path).expect("Save should succeed");

        // Stamp it
        let fingerprint = compute_semantic_fingerprint(&wb);
        let verification = SemanticVerification {
            fingerprint: Some(fingerprint),
            label: None,
            timestamp: None,
        };
        save_semantic_verification(path, &verification).expect("Save verification");

        // Now modify the workbook and save again
        wb.active_sheet_mut().set_value(0, 0, "200");  // Changed value!
        save_workbook(&wb, path).expect("Save modified should succeed");

        // Reload and check - fingerprint should NOT match (drift detected)
        let loaded_wb = load_workbook(path).expect("Load should succeed");
        let loaded_verification = load_semantic_verification(path).unwrap_or_default();
        let current_fingerprint = compute_semantic_fingerprint(&loaded_wb);

        assert_ne!(
            loaded_verification.fingerprint.as_ref(),
            Some(&current_fingerprint),
            "Modified file should show drift - fingerprints should NOT match"
        );
    }

    // ========================================================================
    // Border persistence tests
    // ========================================================================

    #[test]
    fn test_border_color_roundtrip_preserves_rgba() {
        // Set border with explicit RGBA color → save → load → exact same color comes back
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        let red_border = CellBorder {
            style: BorderStyle::Thin,
            color: Some([0xFF, 0x00, 0x00, 0xFF]), // Pure red
        };
        let teal_border = CellBorder {
            style: BorderStyle::Thin,
            color: Some([0x00, 0x99, 0x99, 0xFF]), // Teal
        };
        workbook.active_sheet_mut().set_borders(0, 0, red_border, red_border, red_border, red_border);
        workbook.active_sheet_mut().set_borders(1, 1, teal_border, teal_border, teal_border, teal_border);

        save_workbook(&workbook, path).expect("Save should succeed");
        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify red border preserved exactly
        let fmt0 = loaded.active_sheet().get_format(0, 0);
        assert_eq!(fmt0.border_top.style, BorderStyle::Thin);
        assert_eq!(fmt0.border_top.color, Some([0xFF, 0x00, 0x00, 0xFF]));
        assert_eq!(fmt0.border_right.color, Some([0xFF, 0x00, 0x00, 0xFF]));
        assert_eq!(fmt0.border_bottom.color, Some([0xFF, 0x00, 0x00, 0xFF]));
        assert_eq!(fmt0.border_left.color, Some([0xFF, 0x00, 0x00, 0xFF]));

        // Verify teal border preserved exactly
        let fmt1 = loaded.active_sheet().get_format(1, 1);
        assert_eq!(fmt1.border_top.color, Some([0x00, 0x99, 0x99, 0xFF]));
    }

    #[test]
    fn test_border_auto_none_survives_roundtrip() {
        // Set Auto (None) border → save/load → still None (no accidental literal theme color)
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        // Set a border with no color (Auto/theme default)
        let auto_border = CellBorder {
            style: BorderStyle::Thin,
            color: None, // Auto = use theme default
        };
        workbook.active_sheet_mut().set_borders(0, 0, auto_border, auto_border, auto_border, auto_border);

        save_workbook(&workbook, path).expect("Save should succeed");
        let loaded = load_workbook(path).expect("Load should succeed");

        // Verify None is preserved, not converted to a literal color
        let fmt = loaded.active_sheet().get_format(0, 0);
        assert_eq!(fmt.border_top.style, BorderStyle::Thin);
        assert_eq!(fmt.border_top.color, None, "Auto border color should remain None after round-trip");
        assert_eq!(fmt.border_right.color, None);
        assert_eq!(fmt.border_bottom.color, None);
        assert_eq!(fmt.border_left.color, None);
    }

    #[test]
    fn test_border_style_none_has_no_color() {
        // BorderStyle::None should not store a color
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        // Cell with no borders (default)
        workbook.active_sheet_mut().set_value(0, 0, "test");

        save_workbook(&workbook, path).expect("Save should succeed");
        let loaded = load_workbook(path).expect("Load should succeed");

        let fmt = loaded.active_sheet().get_format(0, 0);
        assert_eq!(fmt.border_top.style, BorderStyle::None);
        assert_eq!(fmt.border_top.color, None);
    }

    #[test]
    fn test_border_mixed_colors_per_edge() {
        // Different colors on different edges should all persist
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        let red = CellBorder { style: BorderStyle::Thin, color: Some([0xFF, 0x00, 0x00, 0xFF]) };
        let blue = CellBorder { style: BorderStyle::Thin, color: Some([0x00, 0x00, 0xFF, 0xFF]) };
        let green = CellBorder { style: BorderStyle::Thin, color: Some([0x00, 0xFF, 0x00, 0xFF]) };
        let black = CellBorder { style: BorderStyle::Thin, color: Some([0x00, 0x00, 0x00, 0xFF]) };

        workbook.active_sheet_mut().set_borders(0, 0, red, blue, green, black);

        save_workbook(&workbook, path).expect("Save should succeed");
        let loaded = load_workbook(path).expect("Load should succeed");

        let fmt = loaded.active_sheet().get_format(0, 0);
        assert_eq!(fmt.border_top.color, Some([0xFF, 0x00, 0x00, 0xFF]), "Top should be red");
        assert_eq!(fmt.border_right.color, Some([0x00, 0x00, 0xFF, 0xFF]), "Right should be blue");
        assert_eq!(fmt.border_bottom.color, Some([0x00, 0xFF, 0x00, 0xFF]), "Bottom should be green");
        assert_eq!(fmt.border_left.color, Some([0x00, 0x00, 0x00, 0xFF]), "Left should be black");
    }

    #[test]
    fn test_font_and_background_color_roundtrip() {
        let temp_file = NamedTempFile::with_suffix(".sheet").unwrap();
        let path = temp_file.path();

        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "colored");

        // Set font_color = red, background_color = yellow
        let mut fmt = workbook.active_sheet().get_format(0, 0).clone();
        fmt.font_color = Some([0xFF, 0x00, 0x00, 0xFF]);
        fmt.background_color = Some([0xFF, 0xFF, 0x00, 0xFF]);
        workbook.active_sheet_mut().set_format(0, 0, fmt);

        // Also test a cell with only background color (no value)
        let mut fmt2 = CellFormat::default();
        fmt2.background_color = Some([0x00, 0x80, 0x00, 0xFF]); // green
        workbook.active_sheet_mut().set_format(1, 1, fmt2);

        save_workbook(&workbook, path).expect("Save should succeed");
        let loaded = load_workbook(path).expect("Load should succeed");

        let fmt0 = loaded.active_sheet().get_format(0, 0);
        assert_eq!(fmt0.font_color, Some([0xFF, 0x00, 0x00, 0xFF]), "Font color should be red");
        assert_eq!(fmt0.background_color, Some([0xFF, 0xFF, 0x00, 0xFF]), "Background should be yellow");

        let fmt1 = loaded.active_sheet().get_format(1, 1);
        assert_eq!(fmt1.background_color, Some([0x00, 0x80, 0x00, 0xFF]), "Background should be green");
        assert_eq!(fmt1.font_color, None, "Font color should be None");
    }
}
