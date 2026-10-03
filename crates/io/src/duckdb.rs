//! Local DuckDB tables, opened read-only. Parquet is the typed interchange
//! boundary so all clients use the same spreadsheet conversion policy.
use std::ffi::{CStr, CString};
use std::path::Path;
use std::ptr;

use libduckdb_sys as ffi;
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;

use crate::{parquet, parquet_export};

pub const IMPORT_CELL_LIMIT: usize = 10_000_000;

fn literal(value: &str) -> String {
    format!("'{}'", value.replace('\'', "''"))
}
fn identifier(value: &str) -> String {
    format!("\"{}\"", value.replace('"', "\"\""))
}
fn path_string(path: &Path) -> Result<&str, String> {
    path.to_str()
        .ok_or_else(|| "DuckDB requires a UTF-8 file path".into())
}
fn cstring(value: &str) -> Result<CString, String> {
    CString::new(value).map_err(|_| "DuckDB input contains a NUL byte".into())
}

// Every opaque C handle has one owner. Results are destroyed on failed queries
// too; varchar/error allocations are copied and freed before returning to Rust.
struct Config(ffi::duckdb_config);
impl Drop for Config {
    fn drop(&mut self) {
        unsafe { ffi::duckdb_destroy_config(&mut self.0) }
    }
}
struct Connection {
    db: ffi::duckdb_database,
    conn: ffi::duckdb_connection,
}
impl Drop for Connection {
    fn drop(&mut self) {
        unsafe {
            if !self.conn.is_null() {
                ffi::duckdb_disconnect(&mut self.conn);
            }
            if !self.db.is_null() {
                ffi::duckdb_close(&mut self.db);
            }
        }
    }
}
struct Query(ffi::duckdb_result);
impl Drop for Query {
    fn drop(&mut self) {
        unsafe { ffi::duckdb_destroy_result(&mut self.0) }
    }
}
impl Query {
    fn rows(&mut self) -> usize {
        unsafe { ffi::duckdb_row_count(&mut self.0) as usize }
    }
    fn text(&mut self, column: usize, row: usize) -> Result<String, String> {
        unsafe {
            let p = ffi::duckdb_value_varchar(&mut self.0, column as u64, row as u64);
            if p.is_null() {
                return Err("Unexpected NULL in DuckDB metadata".into());
            }
            let value = CStr::from_ptr(p).to_string_lossy().into_owned();
            ffi::duckdb_free(p.cast());
            Ok(value)
        }
    }
}
impl Connection {
    fn open(path: &Path, read_only: bool, temp: &Path) -> Result<Self, String> {
        let path = cstring(path_string(path)?)?;
        let mut config = Config(ptr::null_mut());
        let mut db = Self {
            db: ptr::null_mut(),
            conn: ptr::null_mut(),
        };
        unsafe {
            if ffi::duckdb_create_config(&mut config.0) != ffi::duckdb_state_DuckDBSuccess {
                return Err("Could not configure DuckDB".into());
            }
            for (name, value) in [
                (
                    "access_mode",
                    if read_only { "READ_ONLY" } else { "READ_WRITE" },
                ),
                ("autoload_known_extensions", "false"),
                ("autoinstall_known_extensions", "false"),
                ("allow_persistent_secrets", "false"),
                ("threads", "2"),
                ("memory_limit", "256MB"),
                ("temp_directory", ""),
            ] {
                if ffi::duckdb_set_config(
                    config.0,
                    cstring(name)?.as_ptr(),
                    cstring(value)?.as_ptr(),
                ) != ffi::duckdb_state_DuckDBSuccess
                {
                    return Err(format!("Could not set DuckDB option {name}"));
                }
            }
            let mut error = ptr::null_mut();
            let state = ffi::duckdb_open_ext(path.as_ptr(), &mut db.db, config.0, &mut error);
            let detail = if error.is_null() {
                String::new()
            } else {
                let text = CStr::from_ptr(error).to_string_lossy().into_owned();
                ffi::duckdb_free(error.cast());
                text
            };
            if state != ffi::duckdb_state_DuckDBSuccess {
                return Err(format!("Cannot open DuckDB file: {detail}"));
            }
            if ffi::duckdb_connect(db.db, &mut db.conn) != ffi::duckdb_state_DuckDBSuccess {
                return Err("Could not connect to DuckDB".into());
            }
        }
        // These runtime options must be applied in this order. The allowlist
        // cannot be installed after external access has been disabled.
        db.query(&format!(
            "SET allowed_directories=[{}]",
            literal(path_string(temp)?)
        ))?;
        db.query("SET enable_external_access=false")?;
        db.query("SET lock_configuration=true")?;
        Ok(db)
    }
    fn query(&self, sql: &str) -> Result<Query, String> {
        let sql = cstring(sql)?;
        let mut result = Query(unsafe { std::mem::zeroed() });
        unsafe {
            if ffi::duckdb_query(self.conn, sql.as_ptr(), &mut result.0)
                != ffi::duckdb_state_DuckDBSuccess
            {
                let error = ffi::duckdb_result_error(&mut result.0);
                return Err(if error.is_null() {
                    "DuckDB query failed".into()
                } else {
                    CStr::from_ptr(error).to_string_lossy().into_owned()
                });
            }
        }
        Ok(result)
    }
}

#[derive(Debug, Clone)]
pub struct Table {
    pub name: String,
    catalog: String,
    schema: String,
    table: String,
}
impl Table {
    fn qualified(&self) -> String {
        format!(
            "{}.{}.{}",
            identifier(&self.catalog),
            identifier(&self.schema),
            identifier(&self.table)
        )
    }
}

#[derive(Debug, Clone)]
pub struct ColumnInfo {
    pub name: String,
    pub data_type: String,
}

/// Metadata only: describing a table does not materialize its records.
#[derive(Debug, Clone)]
pub struct TableInfo {
    pub index: usize,
    pub name: String,
    pub row_count: u64,
    pub columns: Vec<ColumnInfo>,
}

impl TableInfo {
    pub fn cell_count(&self) -> u64 {
        self.row_count
            .saturating_add(1)
            .saturating_mul(self.columns.len() as u64)
    }

    pub fn import_error(&self) -> Option<String> {
        if self.row_count >= parquet::MAX_ROWS as u64 {
            Some(format!("This table has {} rows; a sheet holds {} data rows. Filter or split it in DuckDB before importing.", self.row_count, parquet::MAX_ROWS - 1))
        } else if self.columns.len() > parquet::MAX_COLS {
            Some(format!("This table has {} columns; maximum is {}. Select fewer columns in DuckDB before importing.", self.columns.len(), parquet::MAX_COLS))
        } else if self.cell_count() > IMPORT_CELL_LIMIT as u64 {
            Some("This table exceeds the import cell budget (10 million cells). Filter or split it in DuckDB before importing.".into())
        } else {
            None
        }
    }
}

/// One connection/snapshot per import. The temporary directory outlives it.
pub struct Database {
    connection: Connection,
    temp: tempfile::TempDir,
    tables: Vec<Table>,
}
impl Database {
    pub fn open(path: &Path) -> Result<Self, String> {
        if !path.is_file() {
            return Err(format!("DuckDB file does not exist: {}", path.display()));
        }
        // DuckDB can detect other storage backends from magic bytes. Reject
        // those before opening, rather than letting a renamed SQLite/Parquet
        // file request a storage extension during database initialization.
        use std::io::Read;
        let mut header = [0u8; 12];
        std::fs::File::open(path)
            .and_then(|mut file| file.read_exact(&mut header))
            .map_err(|e| format!("Cannot open DuckDB file: {e}"))?;
        if &header[8..12] != b"DUCK" {
            return Err("Cannot open DuckDB file: not a DuckDB database".into());
        }
        let temp = tempfile::tempdir().map_err(|e| e.to_string())?;
        let connection = Connection::open(path, true, temp.path())?;
        connection.query("BEGIN TRANSACTION")?;
        // Fully qualify built-ins so stored macros cannot shadow metadata calls.
        let mut result = connection.query("SELECT database_name, schema_name, table_name FROM system.main.duckdb_tables() WHERE NOT internal AND database_name=system.main.current_database() ORDER BY schema_name, table_name")?;
        let mut tables = Vec::new();
        for row in 0..result.rows() {
            let catalog = result.text(0, row)?;
            let schema = result.text(1, row)?;
            let table = result.text(2, row)?;
            // Quote unusual names to distinguish e.g. schema a.b from table b.c.
            let display = |s: &str| {
                if s.chars().all(|c| c.is_alphanumeric() || c == '_') {
                    s.to_owned()
                } else {
                    identifier(s)
                }
            };
            let name = format!("{}.{}", display(&schema), display(&table));
            tables.push(Table {
                name,
                catalog,
                schema,
                table,
            });
        }
        if tables.is_empty() {
            return Err("DuckDB database has no base tables (views are not imported)".into());
        }
        Ok(Self {
            connection,
            temp,
            tables,
        })
    }
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    pub fn describe_table(&self, index: usize) -> Result<TableInfo, String> {
        let table = self
            .tables
            .get(index)
            .ok_or("DuckDB table index out of range")?;
        let mut count = self.connection.query(&format!(
            "SELECT system.main.count(*) FROM {}",
            table.qualified()
        ))?;
        let row_count = count
            .text(0, 0)?
            .parse()
            .map_err(|_| "Invalid DuckDB row count")?;
        let mut result = self.connection.query(&format!(
            "SELECT column_name, data_type FROM system.main.duckdb_columns() WHERE database_name={} AND schema_name={} AND table_name={} ORDER BY column_index",
            literal(&table.catalog), literal(&table.schema), literal(&table.table)
        ))?;
        let mut columns = Vec::with_capacity(result.rows());
        for row in 0..result.rows() {
            columns.push(ColumnInfo {
                name: result.text(0, row)?,
                data_type: result.text(1, row)?,
            });
        }
        Ok(TableInfo {
            index,
            name: table.name.clone(),
            row_count,
            columns,
        })
    }

    /// Import exactly the chosen tables from this connection's snapshot.
    /// Validate the complete selection before allocating any worksheets.
    pub fn import_tables(&self, indices: &[usize]) -> Result<Workbook, String> {
        if indices.is_empty() {
            return Err("Choose at least one table".into());
        }
        let mut seen = std::collections::HashSet::new();
        let mut cells = 0u64;
        for &index in indices {
            if !seen.insert(index) {
                return Err("A table can only be imported once".into());
            }
            let info = self.describe_table(index)?;
            if let Some(error) = info.import_error() {
                return Err(format!("{}: {error}", info.name));
            }
            cells = cells.saturating_add(info.cell_count());
        }
        if cells > IMPORT_CELL_LIMIT as u64 {
            return Err(
                "Selection exceeds the import cell budget (10 million cells). Select fewer tables."
                    .into(),
            );
        }
        let mut sheets = Vec::with_capacity(indices.len());
        let mut budget = IMPORT_CELL_LIMIT;
        for &index in indices {
            let imported = self.read_table(index, parquet::MAX_ROWS - 1, Some(budget), true)?;
            budget -= (imported.rows_loaded + 1) * imported.cols_loaded;
            sheets.push(imported.sheet);
        }
        Ok(Workbook::from_sheets(sheets, 0))
    }
    /// Accept a zero-based tab index, a displayed schema.table name, or an
    /// unqualified table name when it is unique across schemas.
    pub fn resolve(&self, selection: &str) -> Result<usize, String> {
        if let Ok(index) = selection.parse::<usize>() {
            if index < self.tables.len() {
                return Ok(index);
            }
        }
        let exact: Vec<_> = self
            .tables
            .iter()
            .enumerate()
            .filter(|(_, t)| t.name.eq_ignore_ascii_case(selection))
            .collect();
        if exact.len() == 1 {
            return Ok(exact[0].0);
        }
        let matches: Vec<_> = self
            .tables
            .iter()
            .enumerate()
            .filter(|(_, t)| t.table.eq_ignore_ascii_case(selection))
            .collect();
        if matches.len() == 1 {
            return Ok(matches[0].0);
        }
        Err(format!(
            "Unknown or ambiguous DuckDB table {selection:?}; available: {}",
            self.tables
                .iter()
                .map(|t| t.name.as_str())
                .collect::<Vec<_>>()
                .join(", ")
        ))
    }
    pub fn read_table(
        &self,
        index: usize,
        max_rows: usize,
        max_cells: Option<usize>,
        require_whole: bool,
    ) -> Result<parquet::ParquetImport, String> {
        let table = self
            .tables
            .get(index)
            .ok_or("DuckDB table index out of range")?;
        let mut count = self.connection.query(&format!(
            "SELECT system.main.count(*) FROM {}",
            table.qualified()
        ))?;
        let total_rows: u64 = count
            .text(0, 0)?
            .parse()
            .map_err(|_| "Invalid DuckDB row count")?;
        let mut columns = self.connection.query(&format!("SELECT column_name FROM system.main.duckdb_columns() WHERE database_name={} AND schema_name={} AND table_name={} ORDER BY column_index", literal(&table.catalog), literal(&table.schema), literal(&table.table)))?;
        let total_cols = columns.rows();
        if total_cols > parquet::MAX_COLS {
            return Err(format!(
                "{} has {total_cols} columns; maximum is {}",
                table.name,
                parquet::MAX_COLS
            ));
        }
        let rows = total_rows.min(max_rows.min(parquet::MAX_ROWS - 1) as u64) as usize;
        if require_whole && rows as u64 != total_rows {
            return Err(format!("{} has {total_rows} rows; a sheet holds {} data rows. Use peek for a bounded preview.", table.name, parquet::MAX_ROWS - 1));
        }
        if max_cells.is_some_and(|cap| (rows + 1).saturating_mul(total_cols) > cap) {
            return Err(format!(
                "{} exceeds the import cell budget; select a table or preview fewer rows",
                table.name
            ));
        }
        let file = self.temp.path().join(format!("table-{index}.parquet"));
        // No SQL text is accepted from the caller. Identifiers are quoted;
        // DuckDB evaluates table values, including generated columns.
        self.connection.query(&format!(
            "COPY (SELECT * FROM {} LIMIT {rows}) TO {} (FORMAT PARQUET)",
            table.qualified(),
            literal(path_string(&file)?)
        ))?;
        let mut imported = parquet::import_with_limits(&file, rows, max_cells)?;
        imported.total_rows = total_rows;
        imported.sheet.id = SheetId(index as u64 + 1);
        imported.sheet.set_name(&table.name);
        std::fs::remove_file(file).map_err(|e| e.to_string())?;
        Ok(imported)
    }
}

/// Materialize entire tables, refusing to silently drop rows or columns.
pub fn import(path: &Path, selection: Option<&str>) -> Result<Workbook, String> {
    let db = Database::open(path)?;
    let indices = match selection {
        Some(name) => vec![db.resolve(name)?],
        None => (0..db.tables().len()).collect(),
    };
    db.import_tables(&indices)
}

/// Export one worksheet as main.data in a NEW database. Never replace an
/// existing database, including in a race with another writer.
pub fn export(plan: &parquet_export::Plan<'_>, path: &Path) -> Result<(), String> {
    if let Some(error) = plan.validation_error() {
        return Err(error);
    }
    if path.exists() {
        return Err("DuckDB export requires a new file; destination already exists".into());
    }
    let parent = path
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let temp = tempfile::tempdir_in(parent).map_err(|e| e.to_string())?;
    // Absolute paths are needed for DuckDB's external-access allowlist.
    let dir = temp.path().canonicalize().map_err(|e| e.to_string())?;
    let parquet = dir.join("data.parquet");
    plan.write_path(&parquet)?;
    let database = dir.join("export.duckdb");
    {
        let db = Connection::open(&database, false, &dir)?;
        db.query(&format!(
            "CREATE TABLE main.data AS SELECT * FROM system.main.read_parquet({})",
            literal(path_string(&parquet)?)
        ))?;
        // DuckDB otherwise auto-renames columns that differ only in case.
        let mut names = db.query("SELECT column_name FROM system.main.duckdb_columns() WHERE table_name='data' ORDER BY column_index")?;
        for (index, column) in plan.report().columns.iter().enumerate() {
            if names.text(0, index)? != column.name {
                return Err("DuckDB column names must be unique ignoring case; rename conflicting columns before exporting".into());
            }
        }
        db.query("CHECKPOINT")?;
    }
    let mut staged = tempfile::NamedTempFile::new_in(parent).map_err(|e| e.to_string())?;
    std::io::copy(
        &mut std::fs::File::open(database).map_err(|e| e.to_string())?,
        staged.as_file_mut(),
    )
    .map_err(|e| e.to_string())?;
    staged.as_file().sync_all().map_err(|e| e.to_string())?;
    staged
        .persist_noclobber(path)
        .map_err(|e| format!("Could not create DuckDB destination: {e}"))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn fixture() -> (tempfile::TempDir, std::path::PathBuf) {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("source'quoted.duckdb");
        {
            let db = Connection::open(&path, false, dir.path()).unwrap();
            db.query("CREATE TABLE orders AS SELECT 9007199254740993::BIGINT AS id, '007'::VARCHAR AS code, '=1+1' AS literal UNION ALL SELECT NULL, NULL, NULL").unwrap();
            db.query("CREATE MACRO count(x) AS 999").unwrap();
            db.query("CREATE MACRO duckdb_tables() AS TABLE SELECT 'shadowed' AS table_name")
                .unwrap();
            db.query("CREATE SCHEMA archive").unwrap();
            db.query("CREATE TABLE archive.orders (empty INTEGER)")
                .unwrap();
            db.query("CREATE TABLE \"odd'table\" (\"a\"\"b\" INTEGER)")
                .unwrap();
            db.query("CREATE VIEW ignored AS SELECT * FROM orders")
                .unwrap();
        }
        (dir, path)
    }
    #[test]
    fn metadata_and_selection_share_the_snapshot() {
        let (_dir, path) = fixture();
        let before = std::fs::read(&path).unwrap();
        let db = Database::open(&path).unwrap();
        let orders = db.resolve("main.orders").unwrap();
        let empty = db.resolve("archive.orders").unwrap();
        let info = db.describe_table(orders).unwrap();
        assert_eq!(info.row_count, 2);
        assert_eq!(info.columns[0].name, "id");
        assert_eq!(info.columns[0].data_type, "BIGINT");
        assert_eq!(info.cell_count(), 9);
        assert!(info.import_error().is_none());
        // Metadata reads do not create the intermediate Parquet files.
        assert_eq!(std::fs::read_dir(db.temp.path()).unwrap().count(), 0);
        let workbook = db.import_tables(&[orders, empty]).unwrap();
        assert_eq!(workbook.sheet_count(), 2);
        assert_eq!(workbook.sheet(0).unwrap().name, "main.orders");
        assert_eq!(
            workbook.sheet(0).unwrap().get_display(1, 0),
            "9007199254740993"
        );
        assert_eq!(workbook.sheet(1).unwrap().get_display(0, 0), "empty");
        assert!(db.import_tables(&[]).is_err());
        assert!(db.import_tables(&[orders, orders]).is_err());
        assert!(db.import_tables(&[usize::MAX]).is_err());
        drop(db);
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }

    #[test]
    fn oversized_tables_remain_previewable_and_selection_is_validated_first() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("large.duckdb");
        {
            let db = Connection::open(&path, false, dir.path()).unwrap();
            db.query("CREATE TABLE oversized AS SELECT range AS n FROM range(1048576)")
                .unwrap();
            db.query(
                "CREATE TABLE wide_a AS SELECT 1 a, 2 b, 3 c, 4 d, 5 e, 6 f FROM range(900000)",
            )
            .unwrap();
            db.query("CREATE TABLE wide_b AS SELECT * FROM wide_a")
                .unwrap();
        }
        let db = Database::open(&path).unwrap();
        let large = db.resolve("oversized").unwrap();
        let a = db.resolve("wide_a").unwrap();
        let b = db.resolve("wide_b").unwrap();
        assert!(db
            .describe_table(large)
            .unwrap()
            .import_error()
            .unwrap()
            .contains("data rows"));
        let preview = db.read_table(large, 8, Some(100), false).unwrap();
        assert_eq!(preview.rows_loaded, 8);
        assert_eq!(preview.total_rows, 1_048_576);
        assert!(db.import_tables(&[large]).is_err());
        assert!(db.describe_table(a).unwrap().import_error().is_none());
        assert!(db
            .import_tables(&[a, b])
            .err()
            .unwrap()
            .contains("10 million"));
        assert_eq!(std::fs::read_dir(db.temp.path()).unwrap().count(), 0);
    }

    #[test]
    fn tables_preview_and_source_unchanged() {
        let (_dir, path) = fixture();
        let before = std::fs::read(&path).unwrap();
        let db = Database::open(&path).unwrap();
        assert_eq!(db.tables().len(), 3);
        assert!(db.resolve("orders").is_err());
        let index = db.resolve("main.orders").unwrap();
        let table = db.read_table(index, 1, Some(100), false).unwrap();
        assert_eq!(table.total_rows, 2);
        assert!(table.truncated());
        assert_eq!(table.sheet.get_display(1, 0), "9007199254740993");
        assert_eq!(table.sheet.get_display(1, 1), "007");
        assert_eq!(table.sheet.get_display(1, 2), "=1+1");
        assert!(db.read_table(index, 1, None, true).is_err());
        assert!(db.read_table(index, 2, Some(2), true).is_err());
        let odd = db
            .read_table(db.resolve("odd'table").unwrap(), 0, None, true)
            .unwrap();
        assert_eq!(odd.sheet.get_display(0, 0), "a\"b");
        assert!(db
            .connection
            .query("CREATE TABLE forbidden (n INTEGER)")
            .is_err());
        drop(db);
        assert_eq!(std::fs::read(&path).unwrap(), before);
        assert_eq!(import(&path, None).unwrap().sheet_count(), 3);
        assert_eq!(import(&path, Some("main.orders")).unwrap().sheet_count(), 1);
    }
    #[test]
    fn external_access_disabled_and_missing_file_not_created() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("missing.duckdb");
        assert!(Database::open(&path).is_err());
        assert!(!path.exists());
        let db = Connection::open(&path, false, dir.path()).unwrap();
        assert!(db.query("SELECT * FROM read_csv('/etc/passwd')").is_err());
        assert!(db.query("SET enable_external_access=true").is_err());
        db.query("CREATE VIEW only_view AS SELECT 1 AS n").unwrap();
        drop(db);
        assert!(Database::open(&path)
            .err()
            .unwrap()
            .contains("no base tables"));
        std::fs::write(&path, b"not a database").unwrap();
        assert!(Database::open(&path)
            .err()
            .unwrap()
            .contains("Cannot open DuckDB file"));
    }
    #[test]
    fn export_new_database_and_no_overwrite() {
        use parquet_export::Column;
        let (_dir, path) = fixture();
        let wb = import(&path, Some("main.orders")).unwrap();
        let sheet = &wb.sheets()[0];
        let columns = (0..3)
            .map(|index| Column {
                index,
                name: sheet.get_display(0, index),
                as_text: false,
            })
            .collect();
        let plan = parquet_export::analyze(sheet, vec![1, 2], columns).unwrap();
        let out = path.with_file_name("out.duckdb");
        export(&plan, &out).unwrap();
        let imported = import(&out, None).unwrap();
        assert_eq!(imported.sheets()[0].get_display(1, 0), "9007199254740993");
        let bytes = std::fs::read(&out).unwrap();
        assert!(export(&plan, &out).is_err());
        assert_eq!(std::fs::read(&out).unwrap(), bytes);
    }
}
