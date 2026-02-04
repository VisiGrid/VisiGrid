//! Phase 2A: Agent-ready sheet operations
//!
//! Headless closed loop for LLM agents:
//! - `sheet apply --lua` — build sheet from Lua script (replacement semantics)
//! - `sheet inspect` — read cells/ranges/workbook metadata
//! - `sheet verify` — verify fingerprint
//!
//! The Lua API provides an agent-friendly shim over the existing grid.* runtime.

use std::cell::RefCell;
use std::path::Path;
use std::rc::Rc;

use mlua::{Lua, Result as LuaResult, Table, Value as LuaValue};
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;

use crate::replay::ReplayFingerprint;
use crate::{CliError, EXIT_EVAL_ERROR};

/// Result of executing a sheet build script.
#[derive(Debug)]
pub struct SheetApplyResult {
    /// The workbook after build.
    pub workbook: Workbook,
    /// Number of operations (set/clear/meta — fingerprint-affecting).
    pub semantic_ops: usize,
    /// Number of style operations (not in fingerprint).
    pub style_ops: usize,
    /// Computed fingerprint (semantic ops only).
    pub fingerprint: ReplayFingerprint,
    /// Whether verification passed (if --verify was used).
    pub verified: Option<bool>,
    /// Cells changed count.
    pub cells_changed: usize,
}

/// State shared between Lua and Rust during sheet build.
struct BuildState {
    workbook: Workbook,
    /// Hasher for semantic operations (set, clear, meta).
    semantic_hasher: blake3::Hasher,
    /// Count of semantic operations.
    semantic_ops: usize,
    /// Count of style operations (not hashed).
    style_ops: usize,
    /// Track which cells were touched.
    cells_touched: std::collections::HashSet<(usize, usize)>,
}

impl BuildState {
    fn new() -> Self {
        Self {
            workbook: Workbook::new(),
            semantic_hasher: blake3::Hasher::new(),
            semantic_ops: 0,
            style_ops: 0,
            cells_touched: std::collections::HashSet::new(),
        }
    }

    /// Hash a semantic operation (affects fingerprint).
    fn hash_semantic(&mut self, op: &str) {
        self.semantic_hasher.update(op.as_bytes());
        self.semantic_hasher.update(b"\n");
        self.semantic_ops += 1;
    }

    /// Compute the final fingerprint.
    fn fingerprint(&self) -> ReplayFingerprint {
        let hash = self.semantic_hasher.finalize();
        let bytes: [u8; 16] = hash.as_bytes()[0..16].try_into().unwrap();
        ReplayFingerprint::new(self.semantic_ops, bytes)
    }

    /// Get or create a sheet at the given index (0-indexed).
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
}

/// Execute a Lua build script and return the result.
///
/// The Lua script uses agent-friendly API:
/// - `set(cell, value)` — set cell value or formula
/// - `clear(cell)` — clear cell
/// - `meta(target, table)` — set semantic metadata (affects fingerprint)
/// - `style(target, table)` — set presentation style (excluded from fingerprint)
pub fn execute_build_script(script_path: &Path, verify_fp: Option<&str>) -> Result<SheetApplyResult, CliError> {
    let script = std::fs::read_to_string(script_path)
        .map_err(|e| CliError::io(format!("Failed to read {}: {}", script_path.display(), e)))?;

    let lua = Lua::new();
    let state = Rc::new(RefCell::new(BuildState::new()));

    // Register agent-friendly API
    register_agent_api(&lua, state.clone())
        .map_err(|e| CliError { code: EXIT_EVAL_ERROR, message: format!("Lua setup error: {}", e), hint: None })?;

    // Execute the script
    lua.load(&script)
        .exec()
        .map_err(|e| {
            let msg = format!("Lua error: {}", e);
            let hint = if msg.contains("attempt to call a nil value") {
                Some("Available functions: set(cell, value), clear(cell), meta(target, table), style(target, table)".to_string())
            } else {
                None
            };
            CliError { code: EXIT_EVAL_ERROR, message: msg, hint }
        })?;

    // Extract results and run recalc
    let mut state = state.borrow_mut();

    // CRITICAL: Build dependency graph and recalculate all formulas
    // This ensures formula cells have computed values, not just the formula text
    state.workbook.rebuild_dep_graph();
    let _recalc_report = state.workbook.recompute_full_ordered();

    // Compute fingerprint from resulting workbook (same as file fingerprint)
    // This ensures apply fingerprint == file fingerprint regardless of Lua op order
    let fingerprint = compute_sheet_fingerprint(&state.workbook);

    // Verify if requested
    let verified = verify_fp.map(|expected| {
        if let Some(expected_fp) = ReplayFingerprint::parse(expected) {
            fingerprint == expected_fp
        } else {
            false
        }
    });

    Ok(SheetApplyResult {
        workbook: state.workbook.clone(),
        semantic_ops: state.semantic_ops,
        style_ops: state.style_ops,
        fingerprint,
        verified,
        cells_changed: state.cells_touched.len(),
    })
}

/// Register the agent-friendly Lua API.
fn register_agent_api(lua: &Lua, state: Rc<RefCell<BuildState>>) -> LuaResult<()> {
    let globals = lua.globals();

    // set(cell, value) — set cell value or formula
    // Examples: set("A1", "Hello"), set("B2", "=SUM(A:A)"), set("C3", 42)
    {
        let state = state.clone();
        let set_fn = lua.create_function(move |_, (cell, value): (String, LuaValue)| {
            let (row, col) = parse_cell_ref(&cell)
                .ok_or_else(|| mlua::Error::external(format!("Invalid cell reference: {}", cell)))?;

            let value_str = lua_value_to_string(&value);

            let mut state = state.borrow_mut();
            state.ensure_sheet(0);
            state.sheet_mut(0).set_value(row, col, &value_str);
            state.cells_touched.insert((row, col));
            state.hash_semantic(&format!("set:{}:{}:{}", row, col, value_str));

            Ok(())
        })?;
        globals.set("set", set_fn)?;
    }

    // clear(cell) — clear cell value
    {
        let state = state.clone();
        let clear_fn = lua.create_function(move |_, cell: String| {
            let (row, col) = parse_cell_ref(&cell)
                .ok_or_else(|| mlua::Error::external(format!("Invalid cell reference: {}", cell)))?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(0);
            state.sheet_mut(0).set_value(row, col, "");
            state.cells_touched.insert((row, col));
            state.hash_semantic(&format!("clear:{}:{}", row, col));

            Ok(())
        })?;
        globals.set("clear", clear_fn)?;
    }

    // meta(target, table) — set semantic metadata
    // NOTE: Currently meta() is NOT stored in the file, so it's excluded from fingerprint
    // to maintain consistency between apply fingerprint and file fingerprint.
    // When meta storage is implemented, this should be changed to affect fingerprint.
    // Examples: meta("A1", { role = "header" }), meta("A1:D1", { type = "input" })
    {
        let state = state.clone();
        let meta_fn = lua.create_function(move |_, (target, _props): (String, Table)| {
            // Parse target (cell or range) for validation
            let _ = parse_target(&target)
                .ok_or_else(|| mlua::Error::external(format!("Invalid target: {}", target)))?;

            let mut state = state.borrow_mut();
            // Meta operations currently do NOT affect fingerprint (not stored in file)
            // This ensures apply fingerprint == file fingerprint
            state.style_ops += 1;  // Count as style-like op for reporting

            // TODO: Actually store metadata in workbook and include in fingerprint

            Ok(())
        })?;
        globals.set("meta", meta_fn)?;
    }

    // style(target, table) — set presentation style (excluded from fingerprint)
    // Examples: style("A1", { bold = true }), style("A1:D1", { bg = "navy", fg = "white" })
    {
        let state = state.clone();
        let style_fn = lua.create_function(move |_, (target, props): (String, Table)| {
            let parsed = parse_target(&target)
                .ok_or_else(|| mlua::Error::external(format!("Invalid target: {}", target)))?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(0);

            // Apply formatting to each cell in target
            let (start_row, start_col, end_row, end_col) = parsed;
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    // Apply each style property
                    if let Ok(bold) = props.get::<bool>("bold") {
                        state.sheet_mut(0).set_bold(row, col, bold);
                    }
                    if let Ok(italic) = props.get::<bool>("italic") {
                        state.sheet_mut(0).set_italic(row, col, italic);
                    }
                    if let Ok(underline) = props.get::<bool>("underline") {
                        state.sheet_mut(0).set_underline(row, col, underline);
                    }
                    // Note: bg, fg, border would need CellFormat extensions
                }
            }

            // Style operations do NOT affect fingerprint
            state.style_ops += 1;

            Ok(())
        })?;
        globals.set("style", style_fn)?;
    }

    // Also expose grid.* API for compatibility with existing scripts
    register_grid_api_compat(lua, state)?;

    Ok(())
}

/// Register grid.* API for compatibility with existing replay scripts.
fn register_grid_api_compat(lua: &Lua, state: Rc<RefCell<BuildState>>) -> LuaResult<()> {
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
            state.ensure_sheet(sheet - 1);
            state.sheet_mut(sheet - 1).set_value(row, col, &value);
            state.cells_touched.insert((row, col));
            state.hash_semantic(&format!("set:{}:{}:{}:{}", sheet, row, col, value));

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

                let (row, col) = parse_cell_ref(&cell)
                    .ok_or_else(|| mlua::Error::external(format!("Invalid cell reference: {}", cell)))?;

                state.sheet_mut(sheet - 1).set_value(row, col, &value);
                state.cells_touched.insert((row, col));
                op_parts.push(format!("{}:{}:{}", row, col, value));
            }

            state.hash_semantic(&format!("set_batch:{}:{}", sheet, op_parts.join("|")));
            Ok(())
        })?;
        grid.set("set_batch", set_batch_fn)?;
    }

    // grid.format{ sheet=N, range="A1:B2", kind="bold", bold=true, ... }
    // Style operations — do NOT affect fingerprint
    {
        let state = state.clone();
        let format_fn = lua.create_function(move |_, args: Table| {
            let sheet: usize = args.get("sheet")?;
            let range: String = args.get("range")?;

            let (start_row, start_col, end_row, end_col) = parse_target(&range)
                .ok_or_else(|| mlua::Error::external(format!("Invalid range reference: {}", range)))?;

            let mut state = state.borrow_mut();
            state.ensure_sheet(sheet - 1);

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
                }
            }

            // grid.format is style — excluded from fingerprint
            state.style_ops += 1;
            Ok(())
        })?;
        grid.set("format", format_fn)?;
    }

    lua.globals().set("grid", grid)?;
    Ok(())
}

/// Convert Lua value to string for storage.
fn lua_value_to_string(value: &LuaValue) -> String {
    match value {
        LuaValue::Nil => String::new(),
        LuaValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        LuaValue::Integer(n) => n.to_string(),
        LuaValue::Number(n) => {
            // Avoid ".0" suffix for whole numbers
            if n.fract() == 0.0 {
                (*n as i64).to_string()
            } else {
                n.to_string()
            }
        }
        LuaValue::String(s) => s.to_str().map(|s| s.to_string()).unwrap_or_default(),
        _ => format!("{:?}", value),
    }
}

/// Parse a cell reference like "A1" or "AA100" into (row, col).
pub fn parse_cell_ref(s: &str) -> Option<(usize, usize)> {
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

/// Parse a target (cell or range) into (start_row, start_col, end_row, end_col).
fn parse_target(s: &str) -> Option<(usize, usize, usize, usize)> {
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

/// Inspect a cell from a .sheet file.
#[derive(Debug, serde::Serialize)]
pub struct CellInspectResult {
    pub cell: String,
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    pub value_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<CellFormatInfo>,
}

#[derive(Debug, serde::Serialize)]
pub struct CellFormatInfo {
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
}

/// Inspect a range from a .sheet file.
#[derive(Debug, serde::Serialize)]
pub struct RangeInspectResult {
    pub range: String,
    pub cells: Vec<CellInspectResult>,
}

/// Inspect workbook metadata from a .sheet file.
#[derive(Debug, serde::Serialize)]
pub struct WorkbookInspectResult {
    pub fingerprint: String,
    pub sheet_count: usize,
    pub cell_count: usize,
}

/// Compute fingerprint of a .sheet file by rebuilding from cell data.
///
/// This computes the fingerprint that would result from building the sheet
/// via Lua `set()` calls in row-major order.
pub fn compute_sheet_fingerprint(workbook: &Workbook) -> ReplayFingerprint {
    let mut hasher = blake3::Hasher::new();
    let mut op_count = 0;

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
    let bytes: [u8; 16] = hash.as_bytes()[0..16].try_into().unwrap();
    ReplayFingerprint::new(op_count, bytes)
}

/// Format a cell reference from (row, col).
pub fn format_cell_ref(row: usize, col: usize) -> String {
    let mut col_str = String::new();
    let mut c = col;
    loop {
        col_str.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    format!("{}{}", col_str, row + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("a1"), Some((0, 0))); // Case insensitive
    }

    #[test]
    fn test_parse_target_cell() {
        assert_eq!(parse_target("A1"), Some((0, 0, 0, 0)));
    }

    #[test]
    fn test_parse_target_range() {
        assert_eq!(parse_target("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_target("A1:D10"), Some((0, 0, 9, 3)));
    }

    #[test]
    fn test_format_cell_ref() {
        assert_eq!(format_cell_ref(0, 0), "A1");
        assert_eq!(format_cell_ref(0, 25), "Z1");
        assert_eq!(format_cell_ref(0, 26), "AA1");
        assert_eq!(format_cell_ref(9, 3), "D10");
    }

    #[test]
    fn test_simple_build() {
        let script = r#"
set("A1", "Hello")
set("B1", "World")
set("C1", 42)
"#;
        // Write to temp file
        let temp_dir = std::env::temp_dir();
        let script_path = temp_dir.join("test_build.lua");
        std::fs::write(&script_path, script).unwrap();

        let result = execute_build_script(&script_path, None).unwrap();
        assert_eq!(result.semantic_ops, 3);
        assert_eq!(result.cells_changed, 3);

        let sheet = result.workbook.sheet(0).unwrap();
        assert_eq!(sheet.get_display(0, 0), "Hello");
        assert_eq!(sheet.get_display(0, 1), "World");
        assert_eq!(sheet.get_display(0, 2), "42");

        std::fs::remove_file(&script_path).ok();
    }

    #[test]
    fn test_style_excluded_from_fingerprint() {
        let script_without_style = r#"
set("A1", "Hello")
"#;
        let script_with_style = r#"
set("A1", "Hello")
style("A1", { bold = true })
"#;
        let temp_dir = std::env::temp_dir();

        let path1 = temp_dir.join("test_nostyle.lua");
        std::fs::write(&path1, script_without_style).unwrap();
        let result1 = execute_build_script(&path1, None).unwrap();

        let path2 = temp_dir.join("test_withstyle.lua");
        std::fs::write(&path2, script_with_style).unwrap();
        let result2 = execute_build_script(&path2, None).unwrap();

        // Fingerprints should be the same (style excluded)
        assert_eq!(result1.fingerprint, result2.fingerprint);
        // But style_ops should differ
        assert_eq!(result1.style_ops, 0);
        assert_eq!(result2.style_ops, 1);

        std::fs::remove_file(&path1).ok();
        std::fs::remove_file(&path2).ok();
    }

    #[test]
    fn test_meta_included_in_fingerprint() {
        let script_without_meta = r#"
set("A1", "Hello")
"#;
        let script_with_meta = r#"
set("A1", "Hello")
meta("A1", { role = "header" })
"#;
        let temp_dir = std::env::temp_dir();

        let path1 = temp_dir.join("test_nometa.lua");
        std::fs::write(&path1, script_without_meta).unwrap();
        let result1 = execute_build_script(&path1, None).unwrap();

        let path2 = temp_dir.join("test_withmeta.lua");
        std::fs::write(&path2, script_with_meta).unwrap();
        let result2 = execute_build_script(&path2, None).unwrap();

        // Fingerprints should differ (meta included)
        assert_ne!(result1.fingerprint, result2.fingerprint);

        std::fs::remove_file(&path1).ok();
        std::fs::remove_file(&path2).ok();
    }

    #[test]
    fn test_grid_compat_api() {
        let script = r#"
grid.set{ sheet=1, cell="A1", value="Hello" }
grid.format{ sheet=1, range="A1", bold=true }
"#;
        let temp_dir = std::env::temp_dir();
        let path = temp_dir.join("test_grid_compat.lua");
        std::fs::write(&path, script).unwrap();

        let result = execute_build_script(&path, None).unwrap();
        assert_eq!(result.semantic_ops, 1); // Only grid.set counted
        assert_eq!(result.style_ops, 1);    // grid.format is style

        let sheet = result.workbook.sheet(0).unwrap();
        assert_eq!(sheet.get_display(0, 0), "Hello");
        assert!(sheet.get_format(0, 0).bold);

        std::fs::remove_file(&path).ok();
    }
}
