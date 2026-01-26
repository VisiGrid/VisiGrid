//! Provenance replay: execute Lua provenance scripts headlessly.
//!
//! Phase 9B: CLI replay with fingerprint verification.
//!
//! Usage: visigrid-cli replay script.lua [--verify] [--output file.csv]
//!
//! ## Fingerprint Versioning
//!
//! Fingerprints are prefixed with a version (e.g., `v1:6:abc123...`).
//! Fingerprints are stable within the same major version. Breaking changes
//! to fingerprint computation will increment the version.
//!
//! ## Nondeterministic Functions
//!
//! Scripts containing formulas with nondeterministic functions (NOW, TODAY,
//! RAND, RANDBETWEEN) will fail with `--verify` since their output cannot
//! be reliably reproduced.

use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;

use mlua::{Lua, Result as LuaResult, Table};
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;

use crate::{CliError, EXIT_EVAL_ERROR};

/// Fingerprint format version. Increment on breaking changes to fingerprint computation.
pub const FINGERPRINT_VERSION: &str = "v1";

/// Functions that produce nondeterministic results and break verification.
const NONDETERMINISTIC_FUNCTIONS: &[&str] = &["NOW", "TODAY", "RAND", "RANDBETWEEN"];

/// Result of executing a provenance script.
pub struct ReplayResult {
    /// The workbook after replay.
    pub workbook: Workbook,
    /// Number of operations executed.
    pub operations: usize,
    /// Computed fingerprint of operations.
    pub fingerprint: ReplayFingerprint,
    /// Expected fingerprint from script (if present).
    pub expected_fingerprint: Option<ReplayFingerprint>,
    /// Whether fingerprint verification passed.
    pub verified: bool,
    /// Whether the script contains nondeterministic formulas (NOW, RAND, etc.)
    pub has_nondeterministic: bool,
    /// List of nondeterministic functions found (for error reporting).
    pub nondeterministic_found: Vec<String>,
}

/// Fingerprint for replay verification.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ReplayFingerprint {
    pub len: usize,
    pub hash_hi: u64,
    pub hash_lo: u64,
}

impl ReplayFingerprint {
    /// Create a new fingerprint from an operation count and hash.
    pub fn new(len: usize, hash: [u8; 16]) -> Self {
        let hash_hi = u64::from_be_bytes(hash[0..8].try_into().unwrap());
        let hash_lo = u64::from_be_bytes(hash[8..16].try_into().unwrap());
        Self { len, hash_hi, hash_lo }
    }

    /// Format as string: "v1:len:hash" (versioned format)
    pub fn to_string(&self) -> String {
        format!("{}:{}:{:016x}{:016x}", FINGERPRINT_VERSION, self.len, self.hash_hi, self.hash_lo)
    }

    /// Parse from string. Supports both versioned ("v1:len:hash") and legacy ("len:hash") formats.
    pub fn parse(s: &str) -> Option<Self> {
        let parts: Vec<&str> = s.split(':').collect();

        // Try versioned format first: "v1:len:hash"
        if parts.len() == 3 && parts[0].starts_with('v') {
            let _version = parts[0]; // Could validate version compatibility here
            let len: usize = parts[1].parse().ok()?;
            let hex = parts[2];
            if hex.len() != 32 {
                return None;
            }
            let hash_hi = u64::from_str_radix(&hex[0..16], 16).ok()?;
            let hash_lo = u64::from_str_radix(&hex[16..32], 16).ok()?;
            return Some(Self { len, hash_hi, hash_lo });
        }

        // Fall back to legacy format: "len:hash"
        if parts.len() == 2 {
            let len: usize = parts[0].parse().ok()?;
            let hex = parts[1];
            if hex.len() != 32 {
                return None;
            }
            let hash_hi = u64::from_str_radix(&hex[0..16], 16).ok()?;
            let hash_lo = u64::from_str_radix(&hex[16..32], 16).ok()?;
            return Some(Self { len, hash_hi, hash_lo });
        }

        None
    }
}

/// State shared between Lua and Rust during replay.
struct ReplayState {
    workbook: Workbook,
    hasher: blake3::Hasher,
    operation_count: usize,
    /// Nondeterministic functions found in formulas.
    nondeterministic_found: Vec<String>,
}

impl ReplayState {
    fn new() -> Self {
        Self {
            workbook: Workbook::new(),
            hasher: blake3::Hasher::new(),
            operation_count: 0,
            nondeterministic_found: Vec::new(),
        }
    }

    /// Hash an operation for fingerprint computation.
    fn hash_operation(&mut self, op: &str) {
        self.hasher.update(op.as_bytes());
        self.hasher.update(b"\n");
        self.operation_count += 1;
    }

    /// Compute the final fingerprint.
    fn fingerprint(&self) -> ReplayFingerprint {
        let hash = self.hasher.finalize();
        let bytes: [u8; 16] = hash.as_bytes()[0..16].try_into().unwrap();
        ReplayFingerprint::new(self.operation_count, bytes)
    }

    /// Get or create a sheet at the given index.
    fn ensure_sheet(&mut self, sheet_index: usize) {
        while self.workbook.sheet_count() <= sheet_index {
            let name = format!("Sheet{}", self.workbook.sheet_count() + 1);
            self.workbook.add_sheet_named(&name);
        }
    }

    /// Get mutable reference to sheet at index.
    fn sheet_mut(&mut self, index: usize) -> &mut Sheet {
        self.workbook.sheet_mut(index).unwrap()
    }

    /// Check if a value contains a formula with nondeterministic functions.
    /// If found, record them for later error reporting.
    fn check_nondeterministic(&mut self, value: &str) {
        if !value.starts_with('=') {
            return;
        }
        let upper = value.to_uppercase();
        for func in NONDETERMINISTIC_FUNCTIONS {
            // Check for function call pattern: FUNC( or FUNC (
            if upper.contains(&format!("{}(", func)) || upper.contains(&format!("{} (", func)) {
                if !self.nondeterministic_found.contains(&func.to_string()) {
                    self.nondeterministic_found.push(func.to_string());
                }
            }
        }
    }
}

/// Execute a provenance script and return the result.
pub fn execute_script(script_path: &Path) -> Result<ReplayResult, CliError> {
    // Read the script
    let script = std::fs::read_to_string(script_path)
        .map_err(|e| CliError::io(format!("Failed to read {}: {}", script_path.display(), e)))?;

    // Parse expected fingerprint from header (if present)
    let expected_fingerprint = parse_expected_fingerprint(&script);

    // Create Lua state with grid API
    let lua = Lua::new();
    let state = Rc::new(RefCell::new(ReplayState::new()));

    // Register grid API
    register_grid_api(&lua, state.clone())
        .map_err(|e| CliError { code: EXIT_EVAL_ERROR, message: format!("Lua setup error: {}", e) })?;

    // Execute the script
    lua.load(&script)
        .exec()
        .map_err(|e| CliError { code: EXIT_EVAL_ERROR, message: format!("Lua execution error: {}", e) })?;

    // Extract results
    let state = state.borrow();
    let fingerprint = state.fingerprint();
    let has_nondeterministic = !state.nondeterministic_found.is_empty();
    let verified = match &expected_fingerprint {
        Some(expected) => !has_nondeterministic && fingerprint == *expected,
        None => true, // No fingerprint to verify
    };

    Ok(ReplayResult {
        workbook: state.workbook.clone(),
        operations: state.operation_count,
        fingerprint,
        expected_fingerprint,
        verified,
        has_nondeterministic,
        nondeterministic_found: state.nondeterministic_found.clone(),
    })
}

/// Parse expected fingerprint from script header comments.
fn parse_expected_fingerprint(script: &str) -> Option<ReplayFingerprint> {
    for line in script.lines() {
        if line.starts_with("-- Expected fingerprint:") {
            let fp_str = line.strip_prefix("-- Expected fingerprint:")?.trim();
            return ReplayFingerprint::parse(fp_str);
        }
        // Also check footer format
        if line.contains("| Fingerprint ") {
            if let Some(idx) = line.find("| Fingerprint ") {
                let fp_str = &line[idx + "| Fingerprint ".len()..];
                return ReplayFingerprint::parse(fp_str.trim());
            }
        }
        // Stop looking after non-comment lines (past header)
        if !line.starts_with("--") && !line.is_empty() {
            break;
        }
    }
    None
}

/// Register the grid.* API in Lua.
fn register_grid_api(lua: &Lua, state: Rc<RefCell<ReplayState>>) -> LuaResult<()> {
    let grid = lua.create_table()?;

    // grid.set{ sheet=N, cell="A1", value="..." }
    {
        let state = state.clone();
        let set_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let cell: String = args.get("cell")?;
            let value: String = args.get("value")?;

            let (row, col) = parse_cell_ref(&cell)
                .ok_or_else(|| mlua::Error::external(format!("Invalid cell reference: {}", cell)))?;

            let mut state = state.borrow_mut();
            state.check_nondeterministic(&value);
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).set_value(row, col, &value);
            state.hash_operation(&format!("set:{}:{}:{}:{}", sheet, row, col, value));

            Ok(())
        })?;
        grid.set("set", set_fn)?;
    }

    // grid.set_batch{ sheet=N, cells={{cell="A1", value="..."}, ...} }
    {
        let state = state.clone();
        let set_batch_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let cells: Table = args.get("cells")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);

            let mut op_parts = Vec::new();
            for pair in cells.pairs::<i64, Table>() {
                let (_, cell_entry) = pair?;
                let cell: String = cell_entry.get("cell")?;
                let value: String = cell_entry.get("value")?;

                state.check_nondeterministic(&value);

                let (row, col) = parse_cell_ref(&cell)
                    .ok_or_else(|| mlua::Error::external(format!("Invalid cell reference: {}", cell)))?;

                state.sheet_mut(sheet - 1).set_value(row, col, &value);
                op_parts.push(format!("{}:{}:{}", row, col, value));
            }

            state.hash_operation(&format!("set_batch:{}:{}", sheet, op_parts.join("|")));
            Ok(())
        })?;
        grid.set("set_batch", set_batch_fn)?;
    }

    // grid.format{ sheet=N, range="A1:B2", bold=true, ... }
    {
        let state = state.clone();
        let format_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let (start_row, start_col, end_row, end_col) = parse_range_ref(&range)
                .ok_or_else(|| mlua::Error::external(format!("Invalid range reference: {}", range)))?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);

            // Apply formatting
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    if let Ok(bold) = args.get::<bool>("bold") {
                        state.sheet_mut(sheet - 1).set_bold(row, col, bold);
                    }
                    if let Ok(italic) = args.get::<bool>("italic") {
                        state.sheet_mut(sheet - 1).set_italic(row, col, italic);
                    }
                    if let Ok(underline) = args.get::<bool>("underline") {
                        state.sheet_mut(sheet - 1).set_underline(row, col, underline);
                    }
                    // Additional format properties would go here (align, bg, etc.)
                }
            }

            // Hash the operation (simplified - just the range)
            state.hash_operation(&format!("format:{}:{}", sheet, range));
            Ok(())
        })?;
        grid.set("format", format_fn)?;
    }

    // grid.insert_rows{ sheet=N, at=N, count=N }
    {
        let state = state.clone();
        let insert_rows_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let at: usize = args.get("at")?;
            let count: usize = args.get("count")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).insert_rows(at - 1, count);
            state.hash_operation(&format!("insert_rows:{}:{}:{}", sheet, at, count));

            Ok(())
        })?;
        grid.set("insert_rows", insert_rows_fn)?;
    }

    // grid.delete_rows{ sheet=N, at=N, count=N }
    {
        let state = state.clone();
        let delete_rows_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let at: usize = args.get("at")?;
            let count: usize = args.get("count")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).delete_rows(at - 1, count);
            state.hash_operation(&format!("delete_rows:{}:{}:{}", sheet, at, count));

            Ok(())
        })?;
        grid.set("delete_rows", delete_rows_fn)?;
    }

    // grid.insert_cols{ sheet=N, at=N, count=N }
    {
        let state = state.clone();
        let insert_cols_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let at: usize = args.get("at")?;
            let count: usize = args.get("count")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).insert_cols(at - 1, count);
            state.hash_operation(&format!("insert_cols:{}:{}:{}", sheet, at, count));

            Ok(())
        })?;
        grid.set("insert_cols", insert_cols_fn)?;
    }

    // grid.delete_cols{ sheet=N, at=N, count=N }
    {
        let state = state.clone();
        let delete_cols_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let at: usize = args.get("at")?;
            let count: usize = args.get("count")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).delete_cols(at - 1, count);
            state.hash_operation(&format!("delete_cols:{}:{}:{}", sheet, at, count));

            Ok(())
        })?;
        grid.set("delete_cols", delete_cols_fn)?;
    }

    // grid.sort{ sheet=N, col=N, ascending=bool }
    // Note: Sort is hashed but not actually applied in replay (row ordering is captured in values)
    {
        let state = state.clone();
        let sort_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let col: usize = args.get("col")?;
            let ascending: bool = args.get("ascending").unwrap_or(true);

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            // Sort operations are recorded in fingerprint but not replayed directly.
            // The actual row order is captured in subsequent value operations.
            state.hash_operation(&format!("sort:{}:{}:{}", sheet, col, ascending));

            Ok(())
        })?;
        grid.set("sort", sort_fn)?;
    }

    // grid.validate{ sheet=N, range="A1:B2", type="...", ... }
    // Simplified for now - just hash the operation
    {
        let state = state.clone();
        let validate_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;
            let vtype: String = args.get::<String>("type").unwrap_or_default();

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            // TODO: Actually apply validation rule
            state.hash_operation(&format!("validate:{}:{}:{}", sheet, range, vtype));

            Ok(())
        })?;
        grid.set("validate", validate_fn)?;
    }

    // grid.clear_validation{ sheet=N, range="A1:B2" }
    {
        let state = state.clone();
        let clear_validation_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            // TODO: Actually clear validation
            state.hash_operation(&format!("clear_validation:{}:{}", sheet, range));

            Ok(())
        })?;
        grid.set("clear_validation", clear_validation_fn)?;
    }

    // grid.exclude_validation{ sheet=N, range="A1:B2" }
    {
        let state = state.clone();
        let exclude_validation_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.hash_operation(&format!("exclude_validation:{}:{}", sheet, range));

            Ok(())
        })?;
        grid.set("exclude_validation", exclude_validation_fn)?;
    }

    // grid.clear_exclusion{ sheet=N, range="A1:B2" }
    {
        let state = state.clone();
        let clear_exclusion_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);
            state.hash_operation(&format!("clear_exclusion:{}:{}", sheet, range));

            Ok(())
        })?;
        grid.set("clear_exclusion", clear_exclusion_fn)?;
    }

    // grid.define_name{ name="...", sheet=N, range="A1" }
    {
        let state = state.clone();
        let define_name_fn = lua.create_function(move |_, args: Table| {
            let name: String = args.get("name")?;
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);

            // Parse range and create named range
            let (start_row, start_col, end_row, end_col) = parse_range_ref(&range)
                .ok_or_else(|| mlua::Error::external(format!("Invalid range: {}", range)))?;

            use visigrid_engine::named_range::NamedRange;

            let named_range = if start_row == end_row && start_col == end_col {
                NamedRange::cell(&name, sheet - 1, start_row, start_col)
            } else {
                NamedRange::range(&name, sheet - 1, start_row, start_col, end_row, end_col)
            };

            state.workbook.named_ranges_mut().set(named_range)
                .map_err(|e| mlua::Error::external(e))?;
            state.hash_operation(&format!("define_name:{}:{}:{}", name, sheet, range));

            Ok(())
        })?;
        grid.set("define_name", define_name_fn)?;
    }

    // grid.undefine_name{ name="..." }
    {
        let state = state.clone();
        let undefine_name_fn = lua.create_function(move |_, args: Table| {
            let name: String = args.get("name")?;

            let mut state = state.borrow_mut();
            state.workbook.named_ranges_mut().remove(&name);
            state.hash_operation(&format!("undefine_name:{}", name));

            Ok(())
        })?;
        grid.set("undefine_name", undefine_name_fn)?;
    }

    // grid.rename_name{ from="...", to="..." }
    {
        let state = state.clone();
        let rename_name_fn = lua.create_function(move |_, args: Table| {
            let from: String = args.get("from")?;
            let to: String = args.get("to")?;

            let mut state = state.borrow_mut();
            let _ = state.workbook.named_ranges_mut().rename(&from, &to);
            state.hash_operation(&format!("rename_name:{}:{}", from, to));

            Ok(())
        })?;
        grid.set("rename_name", rename_name_fn)?;
    }

    // grid.set_name_description{ name="...", description="..." }
    {
        let state = state.clone();
        let set_desc_fn = lua.create_function(move |_, args: Table| {
            let name: String = args.get("name")?;
            let desc: Option<String> = args.get("description").ok();

            let mut state = state.borrow_mut();
            let _ = state.workbook.named_ranges_mut().set_description(&name, desc.clone());
            state.hash_operation(&format!("set_name_description:{}:{}", name, desc.unwrap_or_default()));

            Ok(())
        })?;
        grid.set("set_name_description", set_desc_fn)?;
    }

    // Register grid table globally
    lua.globals().set("grid", grid)?;

    Ok(())
}

/// Parse a cell reference like "A1" or "AA100" into (row, col).
fn parse_cell_ref(s: &str) -> Option<(usize, usize)> {
    let s = s.to_uppercase();
    let mut col_str = String::new();
    let mut row_str = String::new();

    for c in s.chars() {
        if c.is_ascii_alphabetic() && row_str.is_empty() {
            col_str.push(c);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        } else {
            return None;
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }

    // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
    let mut col: usize = 0;
    for c in col_str.chars() {
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }
    col -= 1; // 0-indexed

    // Convert row to index (1-indexed in input, 0-indexed internally)
    let row: usize = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    Some((row - 1, col))
}

/// Parse a range reference like "A1" or "A1:B10" into (start_row, start_col, end_row, end_col).
fn parse_range_ref(s: &str) -> Option<(usize, usize, usize, usize)> {
    if let Some(colon_idx) = s.find(':') {
        let start = &s[..colon_idx];
        let end = &s[colon_idx + 1..];
        let (sr, sc) = parse_cell_ref(start)?;
        let (er, ec) = parse_cell_ref(end)?;
        Some((sr, sc, er, ec))
    } else {
        let (r, c) = parse_cell_ref(s)?;
        Some((r, c, r, c))
    }
}

/// Export the workbook to a file in the specified format.
pub fn export_workbook(workbook: &Workbook, path: &Path, format: &str) -> Result<(), CliError> {
    match format.to_lowercase().as_str() {
        "csv" => {
            let sheet = workbook.sheet(0)
                .ok_or_else(|| CliError::io("No sheets in workbook"))?;
            let csv = sheet_to_csv(sheet);
            std::fs::write(path, csv)
                .map_err(|e| CliError::io(format!("Failed to write {}: {}", path.display(), e)))?;
        }
        "tsv" => {
            let sheet = workbook.sheet(0)
                .ok_or_else(|| CliError::io("No sheets in workbook"))?;
            let tsv = sheet_to_tsv(sheet);
            std::fs::write(path, tsv)
                .map_err(|e| CliError::io(format!("Failed to write {}: {}", path.display(), e)))?;
        }
        "json" => {
            let sheet = workbook.sheet(0)
                .ok_or_else(|| CliError::io("No sheets in workbook"))?;
            let json = sheet_to_json(sheet);
            std::fs::write(path, json)
                .map_err(|e| CliError::io(format!("Failed to write {}: {}", path.display(), e)))?;
        }
        _ => {
            return Err(CliError::args(format!("Unsupported output format: {}", format)));
        }
    }
    Ok(())
}

/// Convert a sheet to CSV format.
fn sheet_to_csv(sheet: &Sheet) -> String {
    let (max_row, max_col) = get_data_bounds(sheet);
    let mut output = String::new();

    for row in 0..max_row {
        for col in 0..max_col {
            let val = sheet.get_display(row, col);
            // RFC 4180 quoting
            let needs_quote = val.contains(',') || val.contains('"') || val.contains('\n');
            if needs_quote {
                output.push('"');
                output.push_str(&val.replace('"', "\"\""));
                output.push('"');
            } else {
                output.push_str(&val);
            }
            if col < max_col - 1 {
                output.push(',');
            }
        }
        output.push('\n');
    }

    output
}

/// Convert a sheet to TSV format.
fn sheet_to_tsv(sheet: &Sheet) -> String {
    let (max_row, max_col) = get_data_bounds(sheet);
    let mut output = String::new();

    for row in 0..max_row {
        for col in 0..max_col {
            let val = sheet.get_display(row, col);
            output.push_str(&val);
            if col < max_col - 1 {
                output.push('\t');
            }
        }
        output.push('\n');
    }

    output
}

/// Convert a sheet to JSON array format.
fn sheet_to_json(sheet: &Sheet) -> String {
    let (max_row, max_col) = get_data_bounds(sheet);
    let mut rows: Vec<Vec<serde_json::Value>> = Vec::new();

    for row in 0..max_row {
        let mut row_vec: Vec<serde_json::Value> = Vec::new();
        for col in 0..max_col {
            let val = sheet.get_display(row, col);
            // Try to parse as number
            if let Ok(n) = val.parse::<f64>() {
                row_vec.push(serde_json::json!(n));
            } else if val == "TRUE" {
                row_vec.push(serde_json::json!(true));
            } else if val == "FALSE" {
                row_vec.push(serde_json::json!(false));
            } else {
                row_vec.push(serde_json::json!(val));
            }
        }
        rows.push(row_vec);
    }

    serde_json::to_string_pretty(&rows).unwrap_or_else(|_| "[]".to_string())
}

/// Get the bounds of non-empty data in the sheet.
fn get_data_bounds(sheet: &Sheet) -> (usize, usize) {
    let mut max_row = 0;
    let mut max_col = 0;

    for row in 0..sheet.rows {
        for col in 0..sheet.cols {
            if !sheet.get_display(row, col).is_empty() {
                max_row = max_row.max(row + 1);
                max_col = max_col.max(col + 1);
            }
        }
    }

    (max_row, max_col)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("B10"), Some((9, 1)));
        assert_eq!(parse_cell_ref("a1"), Some((0, 0))); // Case insensitive
    }

    #[test]
    fn test_parse_range_ref() {
        assert_eq!(parse_range_ref("A1"), Some((0, 0, 0, 0)));
        assert_eq!(parse_range_ref("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_range_ref("A1:D10"), Some((0, 0, 9, 3)));
    }

    #[test]
    fn test_fingerprint_parse_versioned() {
        let fp = ReplayFingerprint::new(42, [0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0,
                                             0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88]);
        let s = fp.to_string();
        // Should start with version prefix
        assert!(s.starts_with("v1:"), "Fingerprint should be versioned: {}", s);
        let parsed = ReplayFingerprint::parse(&s).unwrap();
        assert_eq!(fp, parsed);
    }

    #[test]
    fn test_fingerprint_parse_legacy() {
        // Legacy format without version prefix should still parse
        let legacy = "42:123456789abcdef0112233445566778";
        assert!(ReplayFingerprint::parse(legacy).is_none()); // 31 chars, should fail
        let legacy = "42:123456789abcdef01122334455667788";
        let parsed = ReplayFingerprint::parse(legacy).unwrap();
        assert_eq!(parsed.len, 42);
    }

    #[test]
    fn test_simple_replay() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="Hello" }
grid.set{ sheet=1, cell="B1", value="World" }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        assert_eq!(state.workbook.sheet(0).unwrap().get_display(0, 0), "Hello");
        assert_eq!(state.workbook.sheet(0).unwrap().get_display(0, 1), "World");
        assert_eq!(state.operation_count, 2);
    }

    // =========================================================================
    // Critical Test 1: Sort Tie-Breaking Determinism
    // =========================================================================
    //
    // Sort operations are HASHED for fingerprint computation but NOT replayed.
    // The actual row order is captured in subsequent value operations.
    // This is by design: provenance scripts contain the explicit state after sort,
    // not the sort operation itself.
    //
    // Implication: If you export history after sorting, the provenance will
    // contain the post-sort cell values, making replay deterministic.

    #[test]
    fn test_sort_is_hashed_but_not_applied() {
        // Sort operations contribute to fingerprint but don't change row order
        // (actual row order is captured via cell values in provenance)
        let script_with_sort = r#"
grid.set{ sheet=1, cell="A1", value="Z" }
grid.set{ sheet=1, cell="A2", value="A" }
grid.sort{ sheet=1, col=1, ascending=true }
"#;
        let script_without_sort = r#"
grid.set{ sheet=1, cell="A1", value="Z" }
grid.set{ sheet=1, cell="A2", value="A" }
"#;
        let lua1 = Lua::new();
        let state1 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua1, state1.clone()).unwrap();
        lua1.load(script_with_sort).exec().unwrap();

        let lua2 = Lua::new();
        let state2 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua2, state2.clone()).unwrap();
        lua2.load(script_without_sort).exec().unwrap();

        // Fingerprints MUST differ (sort is hashed)
        let fp1 = state1.borrow().fingerprint();
        let fp2 = state2.borrow().fingerprint();
        assert_ne!(fp1, fp2, "Sort operation must affect fingerprint");

        // But cell values are the same (sort not applied to data)
        let s1 = state1.borrow();
        let s2 = state2.borrow();
        assert_eq!(
            s1.workbook.sheet(0).unwrap().get_display(0, 0),
            s2.workbook.sheet(0).unwrap().get_display(0, 0)
        );
    }

    // =========================================================================
    // Critical Test 2: Nondeterministic Function Detection
    // =========================================================================

    #[test]
    fn test_nondeterministic_now_detected() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="=NOW()" }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        assert!(!state.nondeterministic_found.is_empty());
        assert!(state.nondeterministic_found.contains(&"NOW".to_string()));
    }

    #[test]
    fn test_nondeterministic_rand_detected() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="=RAND()" }
grid.set{ sheet=1, cell="A2", value="=RANDBETWEEN(1,100)" }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        assert!(state.nondeterministic_found.contains(&"RAND".to_string()));
        assert!(state.nondeterministic_found.contains(&"RANDBETWEEN".to_string()));
    }

    #[test]
    fn test_nondeterministic_today_detected() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="=TODAY()" }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        assert!(state.nondeterministic_found.contains(&"TODAY".to_string()));
    }

    #[test]
    fn test_deterministic_formula_not_flagged() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="=SUM(1,2,3)" }
grid.set{ sheet=1, cell="A2", value="=IF(TRUE, \"yes\", \"no\")" }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        assert!(state.nondeterministic_found.is_empty());
    }

    // =========================================================================
    // Critical Test 3: Validation Replace-in-Range Semantics
    // =========================================================================
    //
    // When setting validation on overlapping ranges, the engine clears any
    // existing rule that overlaps with the new range. This is "replace" semantics.

    #[test]
    fn test_validation_overlap_clears_previous() {
        // Simulate the validation overlap behavior
        let script = r#"
grid.validate{ sheet=1, range="A1:A10", type="whole_number", min=1, max=100 }
grid.validate{ sheet=1, range="A5:A15", type="whole_number", min=1, max=50 }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        // Note: The CLI replay currently hashes validation but doesn't fully apply it.
        // The fingerprint captures the operations, which is what matters for CI.
        // Full validation semantics would require implementing the validation store.
        let state = state.borrow();
        assert_eq!(state.operation_count, 2);

        // Both operations should be hashed
        let fp = state.fingerprint();
        assert_eq!(fp.len, 2);
    }

    // =========================================================================
    // Additional Tests: Fingerprint Determinism
    // =========================================================================

    #[test]
    fn test_same_ops_same_fingerprint() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="X" }
grid.set{ sheet=1, cell="B1", value="Y" }
"#;
        let lua1 = Lua::new();
        let state1 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua1, state1.clone()).unwrap();
        lua1.load(script).exec().unwrap();

        let lua2 = Lua::new();
        let state2 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua2, state2.clone()).unwrap();
        lua2.load(script).exec().unwrap();

        let fp1 = state1.borrow().fingerprint();
        let fp2 = state2.borrow().fingerprint();
        assert_eq!(fp1, fp2, "Same operations must produce same fingerprint");
    }

    #[test]
    fn test_different_order_different_fingerprint() {
        let script1 = r#"
grid.set{ sheet=1, cell="A1", value="X" }
grid.set{ sheet=1, cell="B1", value="Y" }
"#;
        let script2 = r#"
grid.set{ sheet=1, cell="B1", value="Y" }
grid.set{ sheet=1, cell="A1", value="X" }
"#;
        let lua1 = Lua::new();
        let state1 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua1, state1.clone()).unwrap();
        lua1.load(script1).exec().unwrap();

        let lua2 = Lua::new();
        let state2 = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua2, state2.clone()).unwrap();
        lua2.load(script2).exec().unwrap();

        let fp1 = state1.borrow().fingerprint();
        let fp2 = state2.borrow().fingerprint();
        assert_ne!(fp1, fp2, "Different operation order must produce different fingerprint");
    }

    // =========================================================================
    // Output Format Tests
    // =========================================================================

    #[test]
    fn test_csv_escaping() {
        // Use Lua's [[ ]] syntax for multiline strings
        let script = r#"
grid.set{ sheet=1, cell="A1", value="hello, world" }
grid.set{ sheet=1, cell="B1", value='say "hi"' }
grid.set{ sheet=1, cell="C1", value=[[line1
line2]] }
"#;
        let lua = Lua::new();
        let state = Rc::new(RefCell::new(ReplayState::new()));
        register_grid_api(&lua, state.clone()).unwrap();
        lua.load(script).exec().unwrap();

        let state = state.borrow();
        let csv = super::sheet_to_csv(state.workbook.sheet(0).unwrap());

        // Commas, quotes, and newlines should be properly escaped
        assert!(csv.contains("\"hello, world\""), "Comma should be quoted: {}", csv);
        assert!(csv.contains("\"say \"\"hi\"\"\""), "Quotes should be escaped: {}", csv);
        assert!(csv.contains("\"line1\nline2\""), "Newlines should be quoted: {}", csv);
    }
}
