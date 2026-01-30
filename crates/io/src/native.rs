// Native .sheet format using SQLite

use std::path::Path;

use rusqlite::{Connection, params};

use visigrid_engine::cell::{Alignment, CellBorder, CellFormat, DateStyle, NegativeStyle, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::sheet::{MergedRegion, Sheet, SheetId};
use visigrid_engine::workbook::Workbook;
use visigrid_engine::named_range::{NamedRange, NamedRangeTarget};

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
    fmt_font_family TEXT,                -- NULL = inherit from settings
    fmt_thousands INTEGER DEFAULT 0,     -- 1 = use thousands separator
    fmt_negative INTEGER DEFAULT 0,      -- 0=minus, 1=parens, 2=red minus, 3=red parens
    fmt_currency_symbol TEXT,            -- NULL = default ($)
    PRIMARY KEY (row, col)
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
"#;

// Value type constants
const TYPE_EMPTY: i32 = 0;
const TYPE_NUMBER: i32 = 1;
const TYPE_TEXT: i32 = 2;
const TYPE_FORMULA: i32 = 3;

/// Current schema version. Increment for each migration.
const SCHEMA_VERSION: i32 = 1;

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

    // Future migrations: if version < 2 { ... PRAGMA user_version = 2; }

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

        for row in 0..sheet.rows {
            for col in 0..sheet.cols {
                let raw = sheet.get_raw(row, col);
                let format = sheet.get_format(row, col);

                // Skip cells with no value and default formatting
                let has_formatting = format.bold || format.italic || format.underline || format.strikethrough
                    || format.alignment != Alignment::Left
                    || !matches!(format.number_format, NumberFormat::General)
                    || format.font_family.is_some();
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
                    Alignment::General => 3,
                    Alignment::Left => 0,
                    Alignment::Center => 1,
                    Alignment::Right => 2,
                    Alignment::CenterAcrossSelection => 1, // Stored as Center in native format
                };

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
        let alignment = match fmt_alignment {
            1 => Alignment::Center,
            2 => Alignment::Right,
            _ => Alignment::Left,
        };
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
        };
        let has_formatting = format.bold || format.italic || format.underline || format.strikethrough
            || format.alignment != Alignment::Left
            || !matches!(format.number_format, NumberFormat::General)
            || format.font_family.is_some();
        if has_formatting {
            sheet.set_format(row, col, format);
        }
    }

    Ok(sheet)
}

/// Save a complete workbook including named ranges
pub fn save_workbook(workbook: &Workbook, path: &Path) -> Result<(), String> {
    // Delete existing file if present (SQLite will create fresh)
    if path.exists() {
        std::fs::remove_file(path).map_err(|e| e.to_string())?;
    }

    let conn = Connection::open(path).map_err(|e| e.to_string())?;

    // Create schema (includes named_ranges table)
    conn.execute_batch(SCHEMA).map_err(|e| e.to_string())?;
    conn.pragma_update(None, "user_version", SCHEMA_VERSION).map_err(|e| e.to_string())?;

    // Save the active sheet (for now, single-sheet support)
    let sheet = workbook.active_sheet();

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

        for row in 0..sheet.rows {
            for col in 0..sheet.cols {
                let raw = sheet.get_raw(row, col);
                let format = sheet.get_format(row, col);

                // Skip cells with no value and default formatting
                let has_formatting = format.bold || format.italic || format.underline || format.strikethrough
                    || format.alignment != Alignment::Left
                    || !matches!(format.number_format, NumberFormat::General)
                    || format.font_family.is_some();
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
                    Alignment::General => 3,
                    Alignment::Left => 0,
                    Alignment::Center => 1,
                    Alignment::Right => 2,
                    Alignment::CenterAcrossSelection => 1, // Stored as Center in native format
                };

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

    // Save merged regions
    {
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

/// Load a complete workbook including named ranges
pub fn load_workbook(path: &Path) -> Result<Workbook, String> {
    // First load the sheet using existing logic
    let sheet = load(path)?;

    // Create workbook with the loaded sheet
    let mut workbook = Workbook::from_sheets(vec![sheet], 0);

    // Now load named ranges if the table exists
    let conn = Connection::open(path).map_err(|e| e.to_string())?;

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

    // Rebuild dependency graph after loading all data
    workbook.rebuild_dep_graph();

    Ok(workbook)
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
            assert_eq!(version, 1);

            let columns: Vec<String> = conn
                .prepare("PRAGMA table_info(cells)").unwrap()
                .query_map([], |row| row.get::<_, String>(1)).unwrap()
                .filter_map(|r| r.ok())
                .collect();
            assert!(columns.contains(&"fmt_thousands".to_string()));
            assert!(columns.contains(&"fmt_negative".to_string()));
            assert!(columns.contains(&"fmt_currency_symbol".to_string()));
        }
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
}
