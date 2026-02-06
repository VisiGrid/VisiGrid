//! Lua sheet userdata registration.
//!
//! This module provides the `sheet` global that Lua scripts use to interact
//! with the workbook. All operations go through `LuaOpSink` - Lua NEVER
//! touches workbook state directly.
//!
//! # API (all 1-indexed for Lua convention)
//!
//! ## Read/Write by row,col
//! - `sheet:get_value(row, col)` → value or nil
//! - `sheet:set_value(row, col, val_or_nil)` → sets cell value
//! - `sheet:get_formula(row, col)` → formula string or nil
//! - `sheet:set_formula(row, col, formula)` → sets cell formula
//!
//! ## Read/Write by A1 notation
//! - `sheet:get("A1")` → value at A1 notation (shorthand)
//! - `sheet:set("A1", val_or_nil)` → sets value at A1 notation (shorthand)
//! - `sheet:get_a1("A1")` → same as get()
//! - `sheet:set_a1("A1", val_or_nil)` → same as set()
//!
//! ## Sheet info
//! - `sheet:rows()` → number of rows with data
//! - `sheet:cols()` → number of columns with data
//! - `sheet:selection()` → {start_row, start_col, end_row, end_col, range}
//!
//! ## Transactions
//! - `sheet:begin()` → start transaction (noop, for clarity)
//! - `sheet:rollback()` → discard pending changes, returns count discarded
//! - `sheet:commit()` → commit changes (noop, auto-committed at script end)
//!
//! ## Range operations (bulk read/write)
//! - `sheet:range("A1:C5")` → Range object
//! - `range:values()` → 2D table of values
//! - `range:set_values(table)` → write 2D table to range
//! - `range:rows()` → number of rows in range
//! - `range:cols()` → number of columns in range
//! - `range:address()` → "A1:C5" string

use mlua::{Lua, Result as LuaResult, UserData, UserDataMethods, Value};
use std::cell::RefCell;
use std::rc::Rc;

use super::ops::{parse_a1, parse_range, format_a1, CellKey, LuaCellValue, LuaOp, PendingCell, SheetReader};

// ============================================================================
// Limits
// ============================================================================

/// Maximum number of operations per script execution
pub const MAX_OPS: usize = 1_000_000;

/// Maximum lines of print output per script execution
pub const MAX_OUTPUT_LINES: usize = 5_000;

// ============================================================================
// Dynamic Op Sink (type-erased for Lua)
// ============================================================================

/// Type-erased operation sink for Lua userdata.
///
/// This wraps the generic parts of LuaOpSink in a non-generic struct
/// that can be used as mlua UserData.
pub struct DynOpSink {
    /// Ordered list of operations (journal)
    ops: Vec<LuaOp>,
    /// Shadow map for read-after-write
    pending: std::collections::HashMap<CellKey, PendingCell>,
    /// Boxed reader for workbook access
    reader: Box<dyn SheetReader>,
    /// Whether ops limit was exceeded
    ops_limit_exceeded: bool,
    /// Current selection (start_row, start_col, end_row, end_col) - 0-indexed
    selection: (usize, usize, usize, usize),
}

impl DynOpSink {
    /// Create a new sink with a boxed reader (default selection at A1)
    pub fn new(reader: Box<dyn SheetReader>) -> Self {
        Self {
            ops: Vec::new(),
            pending: std::collections::HashMap::new(),
            reader,
            ops_limit_exceeded: false,
            selection: (0, 0, 0, 0),  // A1:A1
        }
    }

    /// Create a new sink with a boxed reader and selection info
    pub fn with_selection(
        reader: Box<dyn SheetReader>,
        selection: (usize, usize, usize, usize),
    ) -> Self {
        Self {
            ops: Vec::new(),
            pending: std::collections::HashMap::new(),
            reader,
            ops_limit_exceeded: false,
            selection,
        }
    }

    /// Check if the ops limit was exceeded
    pub fn ops_limit_exceeded(&self) -> bool {
        self.ops_limit_exceeded
    }

    /// Get the number of unique cells modified
    pub fn mutations(&self) -> usize {
        self.pending.len()
    }

    /// Take the ops out, leaving an empty vec
    pub fn take_ops(&mut self) -> Vec<LuaOp> {
        std::mem::take(&mut self.ops)
    }

    // ========================================================================
    // Read operations (check pending first, then fall back to reader)
    // ========================================================================

    /// Get value at (row, col) - 1-indexed
    fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
        // Convert 1-indexed to 0-indexed for internal use
        if row == 0 || col == 0 {
            return LuaCellValue::Error("Row and column must be >= 1".to_string());
        }
        let row = row - 1;
        let col = col - 1;

        let key = CellKey::from((row, col));

        // Check pending first (read-after-write)
        if let Some(pending) = self.pending.get(&key) {
            return pending.read_value();
        }

        // Fall back to workbook
        self.reader.get_value(row, col)
    }

    /// Get formula at (row, col) - 1-indexed
    fn get_formula(&self, row: usize, col: usize) -> Option<String> {
        if row == 0 || col == 0 {
            return None;
        }
        let row = row - 1;
        let col = col - 1;

        let key = CellKey::from((row, col));

        // Check pending first
        if let Some(pending) = self.pending.get(&key) {
            return pending.formula().map(|s| s.to_string());
        }

        // Fall back to workbook
        self.reader.get_formula(row, col)
    }

    /// Get row count
    fn rows(&self) -> usize {
        self.reader.rows()
    }

    /// Get column count
    fn cols(&self) -> usize {
        self.reader.cols()
    }

    /// Get current selection (start_row, start_col, end_row, end_col) - 0-indexed
    fn selection(&self) -> (usize, usize, usize, usize) {
        self.selection
    }

    // ========================================================================
    // Transaction control
    // ========================================================================

    /// Begin a transaction (noop - all ops are already batched)
    fn begin(&self) {
        // No-op: ops are already batched by default
        // This method exists for API clarity and forward compatibility
    }

    /// Rollback all pending changes, returning the count of discarded ops
    fn rollback(&mut self) -> usize {
        let count = self.ops.len();
        self.ops.clear();
        self.pending.clear();
        count
    }

    /// Commit pending changes (noop - auto-committed at script end)
    fn commit(&self) {
        // No-op: changes are automatically committed when script finishes
        // This method exists for API clarity
    }

    // ========================================================================
    // Write operations (update pending + append to journal)
    // ========================================================================

    /// Check if we can add another op (returns error message if not)
    fn check_ops_limit(&mut self) -> Result<(), String> {
        if self.ops.len() >= MAX_OPS {
            self.ops_limit_exceeded = true;
            Err(format!("operation limit exceeded ({} ops)", MAX_OPS))
        } else {
            Ok(())
        }
    }

    /// Set value at (row, col) - 1-indexed
    fn set_value(&mut self, row: usize, col: usize, value: LuaCellValue) -> Result<(), String> {
        self.check_ops_limit()?;

        if row == 0 || col == 0 {
            return Err("row and column must be >= 1".to_string());
        }
        let row0 = row - 1;
        let col0 = col - 1;

        let key = CellKey::from((row0, col0));

        // Update shadow map
        self.pending.insert(key, PendingCell::Value(value.clone()));

        // Append to journal (store 0-indexed internally)
        self.ops.push(LuaOp::SetValue {
            row: row0 as u32,
            col: col0 as u32,
            value,
        });

        Ok(())
    }

    /// Push a SetCellStyle range op.
    /// Range coords are 1-indexed (Lua convention); converted to 0-indexed internally.
    fn push_style_op(&mut self, r1: usize, c1: usize, r2: usize, c2: usize, style: u8) -> Result<(), String> {
        self.check_ops_limit()?;
        if r1 == 0 || c1 == 0 || r2 == 0 || c2 == 0 {
            return Err("row and column must be >= 1".to_string());
        }
        self.ops.push(LuaOp::SetCellStyle {
            r1: (r1 - 1) as u32,
            c1: (c1 - 1) as u32,
            r2: (r2 - 1) as u32,
            c2: (c2 - 1) as u32,
            style,
        });
        Ok(())
    }

    /// Set formula at (row, col) - 1-indexed
    fn set_formula(&mut self, row: usize, col: usize, formula: String) -> Result<(), String> {
        self.check_ops_limit()?;

        if row == 0 || col == 0 {
            return Err("row and column must be >= 1".to_string());
        }
        let row0 = row - 1;
        let col0 = col - 1;

        let key = CellKey::from((row0, col0));

        // Update shadow map
        self.pending
            .insert(key, PendingCell::Formula(formula.clone()));

        // Append to journal
        self.ops.push(LuaOp::SetFormula {
            row: row0 as u32,
            col: col0 as u32,
            formula,
        });

        Ok(())
    }
}

// ============================================================================
// Sheet UserData (the Lua-facing type)
// ============================================================================

/// The `sheet` userdata that Lua scripts interact with.
///
/// This is a thin wrapper around `Rc<RefCell<DynOpSink>>` that implements
/// mlua's UserData trait.
#[derive(Clone)]
pub struct SheetUserData {
    sink: Rc<RefCell<DynOpSink>>,
}

impl SheetUserData {
    /// Create a new sheet userdata wrapping a sink
    pub fn new(sink: Rc<RefCell<DynOpSink>>) -> Self {
        Self { sink }
    }
}

impl UserData for SheetUserData {
    fn add_methods<M: UserDataMethods<Self>>(methods: &mut M) {
        // ====================================================================
        // get_value(row, col) -> value or nil
        // ====================================================================
        methods.add_method("get_value", |lua, this, (row, col): (usize, usize)| {
            let sink = this.sink.borrow();
            cell_value_to_lua(lua, sink.get_value(row, col))
        });

        // ====================================================================
        // set_value(row, col, val_or_nil)
        // ====================================================================
        methods.add_method("set_value", |_, this, (row, col, val): (usize, usize, Value)| {
            let value = lua_to_cell_value(val);
            this.sink.borrow_mut().set_value(row, col, value)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        // ====================================================================
        // get_formula(row, col) -> formula string or nil
        // ====================================================================
        methods.add_method("get_formula", |_, this, (row, col): (usize, usize)| {
            let sink = this.sink.borrow();
            Ok(sink.get_formula(row, col))
        });

        // ====================================================================
        // set_formula(row, col, formula)
        // ====================================================================
        methods.add_method("set_formula", |_, this, (row, col, formula): (usize, usize, String)| {
            this.sink.borrow_mut().set_formula(row, col, formula)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        // ====================================================================
        // get_a1("A1") -> value at A1 notation
        // ====================================================================
        methods.add_method("get_a1", |lua, this, a1: String| {
            if let Some((row, col)) = parse_a1(&a1) {
                let sink = this.sink.borrow();
                cell_value_to_lua(lua, sink.get_value(row, col))
            } else {
                Ok(Value::Nil)
            }
        });

        // ====================================================================
        // set_a1("A1", val_or_nil)
        // ====================================================================
        methods.add_method("set_a1", |_, this, (a1, val): (String, Value)| {
            if let Some((row, col)) = parse_a1(&a1) {
                let value = lua_to_cell_value(val);
                this.sink.borrow_mut().set_value(row, col, value)
                    .map_err(|e| mlua::Error::RuntimeError(e))?;
            }
            Ok(())
        });

        // ====================================================================
        // get("A1") -> shorthand alias for get_a1
        // ====================================================================
        methods.add_method("get", |lua, this, a1: String| {
            if let Some((row, col)) = parse_a1(&a1) {
                let sink = this.sink.borrow();
                cell_value_to_lua(lua, sink.get_value(row, col))
            } else {
                Ok(Value::Nil)
            }
        });

        // ====================================================================
        // set("A1", val_or_nil) -> shorthand alias for set_a1
        // ====================================================================
        methods.add_method("set", |_, this, (a1, val): (String, Value)| {
            if let Some((row, col)) = parse_a1(&a1) {
                let value = lua_to_cell_value(val);
                this.sink.borrow_mut().set_value(row, col, value)
                    .map_err(|e| mlua::Error::RuntimeError(e))?;
            }
            Ok(())
        });

        // ====================================================================
        // rows() -> number of rows with data
        // ====================================================================
        methods.add_method("rows", |_, this, ()| {
            Ok(this.sink.borrow().rows())
        });

        // ====================================================================
        // cols() -> number of columns with data
        // ====================================================================
        methods.add_method("cols", |_, this, ()| {
            Ok(this.sink.borrow().cols())
        });

        // ====================================================================
        // selection() -> {start_row, start_col, end_row, end_col} (1-indexed)
        // ====================================================================
        methods.add_method("selection", |lua, this, ()| {
            let sink = this.sink.borrow();
            let sel = sink.selection();

            // Return as Lua table with named fields (1-indexed for Lua)
            let table = lua.create_table()?;
            table.set("start_row", sel.0 + 1)?;
            table.set("start_col", sel.1 + 1)?;
            table.set("end_row", sel.2 + 1)?;
            table.set("end_col", sel.3 + 1)?;

            // Also include A1 notation for convenience
            let start_a1 = format_a1(sel.0 + 1, sel.1 + 1);
            let end_a1 = format_a1(sel.2 + 1, sel.3 + 1);
            if start_a1 == end_a1 {
                table.set("range", start_a1)?;
            } else {
                table.set("range", format!("{}:{}", start_a1, end_a1))?;
            }

            Ok(Value::Table(table))
        });

        // ====================================================================
        // begin() -> start transaction (noop, for API clarity)
        // ====================================================================
        methods.add_method("begin", |_, this, ()| {
            this.sink.borrow().begin();
            Ok(())
        });

        // ====================================================================
        // rollback() -> discard all pending changes, returns count
        // ====================================================================
        methods.add_method("rollback", |_, this, ()| {
            let count = this.sink.borrow_mut().rollback();
            Ok(count)
        });

        // ====================================================================
        // commit() -> commit changes (noop, auto-committed at script end)
        // ====================================================================
        methods.add_method("commit", |_, this, ()| {
            this.sink.borrow().commit();
            Ok(())
        });

        // ====================================================================
        // style("A1:C5", "Error") -> set cell style on a range
        // ====================================================================
        methods.add_method("style", |_, this, (range_str, style_val): (String, Value)| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            let style = parse_style_arg(style_val)?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, style)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        // ====================================================================
        // Sugar aliases: sheet:error("A1:C5"), sheet:warning("A1:C5"), etc.
        // ====================================================================
        methods.add_method("error", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 1)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("warning", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 2)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("success", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 3)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("input", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 4)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("total", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 5)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("note", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 6)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        methods.add_method("clear_style", |_, this, range_str: String| {
            let ((r1, c1), (r2, c2)) = parse_range(&range_str)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))?;
            this.sink.borrow_mut().push_style_op(r1, c1, r2, c2, 0)
                .map_err(|e| mlua::Error::RuntimeError(e))
        });

        // ====================================================================
        // range("A1:C5") -> Range object for bulk operations
        // ====================================================================
        methods.add_method("range", |_, this, range_str: String| {
            if let Some(((start_row, start_col), (end_row, end_col))) = parse_range(&range_str) {
                Ok(RangeUserData {
                    sink: this.sink.clone(),
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                })
            } else {
                Err(mlua::Error::RuntimeError(format!("Invalid range: '{}'", range_str)))
            }
        });
    }
}

// ============================================================================
// Range UserData (for bulk operations)
// ============================================================================

/// A range object that supports bulk read/write via values()/set_values().
///
/// Created via `sheet:range("A1:C5")`.
#[derive(Clone)]
pub struct RangeUserData {
    sink: Rc<RefCell<DynOpSink>>,
    start_row: usize,  // 1-indexed
    start_col: usize,  // 1-indexed
    end_row: usize,    // 1-indexed
    end_col: usize,    // 1-indexed
}

impl UserData for RangeUserData {
    fn add_methods<M: UserDataMethods<Self>>(methods: &mut M) {
        // ====================================================================
        // values() -> 2D table of values
        // ====================================================================
        methods.add_method("values", |lua, this, ()| {
            let sink = this.sink.borrow();
            let outer = lua.create_table()?;

            for row in this.start_row..=this.end_row {
                let inner = lua.create_table()?;
                for col in this.start_col..=this.end_col {
                    let value = sink.get_value(row, col);
                    let lua_value = cell_value_to_lua(lua, value)?;
                    // Use 1-indexed within the range (Lua convention)
                    inner.set(col - this.start_col + 1, lua_value)?;
                }
                outer.set(row - this.start_row + 1, inner)?;
            }

            Ok(Value::Table(outer))
        });

        // ====================================================================
        // set_values(table) -> write 2D table to range
        // ====================================================================
        methods.add_method("set_values", |_, this, table: mlua::Table| {
            let mut sink = this.sink.borrow_mut();

            // Iterate over outer table (rows)
            for row_idx in 1..=(this.end_row - this.start_row + 1) {
                if let Ok(row_table) = table.get::<mlua::Table>(row_idx) {
                    // Iterate over inner table (columns)
                    for col_idx in 1..=(this.end_col - this.start_col + 1) {
                        if let Ok(value) = row_table.get::<Value>(col_idx) {
                            let cell_value = lua_to_cell_value(value);
                            let actual_row = this.start_row + row_idx - 1;
                            let actual_col = this.start_col + col_idx - 1;
                            sink.set_value(actual_row, actual_col, cell_value)
                                .map_err(|e| mlua::Error::RuntimeError(e))?;
                        }
                    }
                }
            }

            Ok(())
        });

        // ====================================================================
        // rows() -> number of rows in range
        // ====================================================================
        methods.add_method("rows", |_, this, ()| {
            Ok(this.end_row - this.start_row + 1)
        });

        // ====================================================================
        // cols() -> number of columns in range
        // ====================================================================
        methods.add_method("cols", |_, this, ()| {
            Ok(this.end_col - this.start_col + 1)
        });

        // ====================================================================
        // address() -> "A1:C5" string representation
        // ====================================================================
        methods.add_method("address", |_, this, ()| {
            let start = format_a1(this.start_row, this.start_col);
            let end = format_a1(this.end_row, this.end_col);
            if start == end {
                Ok(start)
            } else {
                Ok(format!("{}:{}", start, end))
            }
        });
    }
}

// ============================================================================
// Type conversion helpers
// ============================================================================

/// Parse a Lua value as a cell style identifier (string name or integer constant).
/// Returns the u8 style code (0-6).
fn parse_style_arg(val: Value) -> Result<u8, mlua::Error> {
    match val {
        Value::Integer(i) => {
            let i = i as i32;
            if !(0..=6).contains(&i) {
                return Err(mlua::Error::RuntimeError(format!("Unknown style id: {}. Valid: 0-6", i)));
            }
            Ok(i as u8)
        }
        Value::Number(n) => {
            let i = n as i32;
            if !(0..=6).contains(&i) {
                return Err(mlua::Error::RuntimeError(format!("Unknown style id: {}. Valid: 0-6", i)));
            }
            Ok(i as u8)
        }
        Value::String(s) => {
            let s = s.to_str().map_err(|_| mlua::Error::RuntimeError("Invalid UTF-8 in style name".into()))?;
            match s.to_lowercase().as_str() {
                "error" => Ok(1),
                "warning" | "warn" => Ok(2),
                "success" | "ok" => Ok(3),
                "input" => Ok(4),
                "total" | "totals" => Ok(5),
                "note" => Ok(6),
                "default" | "none" | "clear" => Ok(0),
                other => Err(mlua::Error::RuntimeError(format!(
                    "Unknown style: '{}'. Valid: error, warning, success, input, total, note, default", other
                ))),
            }
        }
        _ => Err(mlua::Error::RuntimeError("style must be a string or integer".into())),
    }
}

/// Convert LuaCellValue to mlua Value (requires Lua context for strings)
fn cell_value_to_lua(lua: &Lua, value: LuaCellValue) -> LuaResult<Value> {
    match value {
        LuaCellValue::Nil => Ok(Value::Nil),
        LuaCellValue::Number(n) => Ok(Value::Number(n)),
        LuaCellValue::String(s) => {
            let lua_str = lua.create_string(&s)?;
            Ok(Value::String(lua_str))
        }
        LuaCellValue::Bool(b) => Ok(Value::Boolean(b)),
        LuaCellValue::Error(e) => {
            // Return error string so Lua can see what went wrong
            let lua_str = lua.create_string(&format!("#ERROR: {}", e))?;
            Ok(Value::String(lua_str))
        }
    }
}

/// Convert mlua Value to LuaCellValue
fn lua_to_cell_value(value: Value) -> LuaCellValue {
    match value {
        Value::Nil => LuaCellValue::Nil,
        Value::Boolean(b) => LuaCellValue::Bool(b),
        Value::Integer(i) => LuaCellValue::Number(i as f64),
        Value::Number(n) => LuaCellValue::Number(n),
        Value::String(s) => {
            match s.to_str() {
                Ok(str) => LuaCellValue::String(str.to_string()),
                Err(_) => LuaCellValue::Error("Invalid UTF-8".to_string()),
            }
        }
        _ => LuaCellValue::Error("Unsupported type".to_string()),
    }
}

// ============================================================================
// Registration helper
// ============================================================================

/// Register the `sheet` global in a Lua instance.
///
/// Returns the sink wrapped in Rc<RefCell<>> so the caller can extract ops after eval.
pub fn register_sheet_global(lua: &Lua, reader: Box<dyn SheetReader>) -> LuaResult<Rc<RefCell<DynOpSink>>> {
    let sink = Rc::new(RefCell::new(DynOpSink::new(reader)));
    let userdata = SheetUserData::new(sink.clone());

    lua.globals().set("sheet", userdata)?;
    register_styles_table(lua)?;

    Ok(sink)
}

/// Register the `sheet` global with selection info.
///
/// Selection is (start_row, start_col, end_row, end_col) in 0-indexed coords.
pub fn register_sheet_global_with_selection(
    lua: &Lua,
    reader: Box<dyn SheetReader>,
    selection: (usize, usize, usize, usize),
) -> LuaResult<Rc<RefCell<DynOpSink>>> {
    let sink = Rc::new(RefCell::new(DynOpSink::with_selection(reader, selection)));
    let userdata = SheetUserData::new(sink.clone());

    lua.globals().set("sheet", userdata)?;
    register_styles_table(lua)?;

    Ok(sink)
}

/// Register the `styles` constant table as a Lua global.
fn register_styles_table(lua: &Lua) -> LuaResult<()> {
    let styles = lua.create_table()?;
    styles.set("Default", 0)?;
    styles.set("Error", 1)?;
    styles.set("Warning", 2)?;
    styles.set("Success", 3)?;
    styles.set("Input", 4)?;
    styles.set("Total", 5)?;
    styles.set("Note", 6)?;
    lua.globals().set("styles", styles)?;
    Ok(())
}

// ============================================================================
// Workbook Adapter (for actual VisiGrid sheets)
// ============================================================================

use visigrid_engine::sheet::{Sheet, SheetId};
use visigrid_engine::cell::CellValue;

/// Adapter that wraps a snapshot of sheet data for Lua read access.
///
/// This takes a copy of the relevant cell data to avoid borrowing issues
/// during Lua execution.
pub struct SheetSnapshot {
    /// Copied cell values (0-indexed)
    values: std::collections::HashMap<(usize, usize), LuaCellValue>,
    /// Copied formulas (0-indexed)
    formulas: std::collections::HashMap<(usize, usize), String>,
    /// Sheet dimensions
    rows: usize,
    cols: usize,
}

impl SheetSnapshot {
    /// Create a snapshot from a Sheet reference.
    ///
    /// Uses sparse iteration - only copies cells that actually exist.
    /// O(populated cells), not O(rows * cols).
    pub fn from_sheet(sheet: &Sheet) -> Self {
        let mut values = std::collections::HashMap::new();
        let mut formulas = std::collections::HashMap::new();

        // Sparse iteration - only populated cells
        for (&(row, col), cell) in sheet.cells_iter() {
            let raw = cell.value.raw_display();
            if !raw.is_empty() {
                // Convert CellValue to LuaCellValue
                let lua_value = cell_value_to_lua_cell_value(&cell.value, sheet, row, col);
                values.insert((row, col), lua_value);

                // Check if it's a formula
                if raw.starts_with('=') {
                    formulas.insert((row, col), raw);
                }
            }
        }

        Self {
            values,
            formulas,
            rows: sheet.rows,
            cols: sheet.cols,
        }
    }
}

impl SheetReader for SheetSnapshot {
    fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
        self.values.get(&(row, col)).cloned().unwrap_or(LuaCellValue::Nil)
    }

    fn get_formula(&self, row: usize, col: usize) -> Option<String> {
        self.formulas.get(&(row, col)).cloned()
    }

    fn rows(&self) -> usize {
        self.rows
    }

    fn cols(&self) -> usize {
        self.cols
    }
}

/// Convert engine CellValue to LuaCellValue
fn cell_value_to_lua_cell_value(value: &CellValue, sheet: &Sheet, row: usize, col: usize) -> LuaCellValue {
    match value {
        CellValue::Empty => LuaCellValue::Nil,
        CellValue::Number(n) => LuaCellValue::Number(*n),
        CellValue::Text(s) => LuaCellValue::String(s.clone()),
        CellValue::Formula { .. } => {
            // For formulas, return the evaluated display value
            let display = sheet.get_display(row, col);
            // Try to parse as number first
            if let Ok(n) = display.parse::<f64>() {
                LuaCellValue::Number(n)
            } else if display.starts_with('#') {
                // Error value
                LuaCellValue::Error(display)
            } else if display == "TRUE" || display == "FALSE" {
                LuaCellValue::Bool(display == "TRUE")
            } else {
                LuaCellValue::String(display)
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    // Mock reader for tests
    struct MockReader {
        data: HashMap<(usize, usize), LuaCellValue>,
        formulas: HashMap<(usize, usize), String>,
    }

    impl MockReader {
        fn new() -> Self {
            Self {
                data: HashMap::new(),
                formulas: HashMap::new(),
            }
        }

        fn with_value(mut self, row: usize, col: usize, val: LuaCellValue) -> Self {
            self.data.insert((row, col), val);
            self
        }
    }

    impl SheetReader for MockReader {
        fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
            self.data.get(&(row, col)).cloned().unwrap_or(LuaCellValue::Nil)
        }

        fn get_formula(&self, row: usize, col: usize) -> Option<String> {
            self.formulas.get(&(row, col)).cloned()
        }

        fn rows(&self) -> usize {
            100
        }

        fn cols(&self) -> usize {
            26
        }
    }

    #[test]
    fn test_dyn_sink_1_indexed() {
        let reader = MockReader::new().with_value(0, 0, LuaCellValue::Number(42.0));
        let mut sink = DynOpSink::new(Box::new(reader));

        // Lua uses 1-indexed, so (1,1) maps to internal (0,0)
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(42.0));

        // Invalid coords return error
        assert!(matches!(sink.get_value(0, 1), LuaCellValue::Error(_)));
    }

    #[test]
    fn test_dyn_sink_set_read() {
        let reader = MockReader::new();
        let mut sink = DynOpSink::new(Box::new(reader));

        sink.set_value(1, 1, LuaCellValue::Number(100.0));
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(100.0));

        // Check ops were recorded (0-indexed internally)
        let ops = sink.take_ops();
        assert_eq!(ops.len(), 1);
        match &ops[0] {
            LuaOp::SetValue { row, col, value } => {
                assert_eq!(*row, 0);
                assert_eq!(*col, 0);
                assert_eq!(*value, LuaCellValue::Number(100.0));
            }
            _ => panic!("Expected SetValue"),
        }
    }

    #[test]
    fn test_lua_to_cell_value() {
        assert!(matches!(lua_to_cell_value(Value::Nil), LuaCellValue::Nil));
        assert!(matches!(lua_to_cell_value(Value::Boolean(true)), LuaCellValue::Bool(true)));
        assert!(matches!(lua_to_cell_value(Value::Number(3.14)), LuaCellValue::Number(n) if (n - 3.14).abs() < 0.001));
        assert!(matches!(lua_to_cell_value(Value::Integer(42)), LuaCellValue::Number(n) if n == 42.0));
    }

    // ========================================================================
    // Performance tests for SheetSnapshot
    // ========================================================================

    #[test]
    fn test_snapshot_sparse_performance() {
        // Create a large sheet (1000 rows x 100 cols = 100k grid)
        // but with only 10k populated cells (10% density)
        let mut sheet = Sheet::new(SheetId(1), 1000, 100);

        // Populate 10k cells - use deterministic pattern with no collisions
        // row = i / 100, col = i % 100 gives unique (row, col) for i in 0..10k
        for i in 0..10_000usize {
            let row = i / 100;
            let col = i % 100;
            sheet.set_value(row, col, &format!("{}", i));
        }

        // Time the snapshot creation
        let start = std::time::Instant::now();
        let snapshot = SheetSnapshot::from_sheet(&sheet);
        let elapsed = start.elapsed();

        // Verify correct number of cells copied
        assert_eq!(snapshot.values.len(), 10_000);

        // Should be fast - under 50ms for 10k cells
        assert!(
            elapsed.as_millis() < 50,
            "Snapshot took {}ms for 10k cells - too slow!",
            elapsed.as_millis()
        );

        println!(
            "SheetSnapshot: 10k cells in {:?} ({:.2} µs/cell)",
            elapsed,
            elapsed.as_micros() as f64 / 10_000.0
        );
    }

    #[test]
    fn test_snapshot_read_performance() {
        // Create sheet with 10k cells
        let mut sheet = Sheet::new(SheetId(1), 1000, 100);
        for i in 0..10_000usize {
            let row = i / 100;
            let col = i % 100;
            sheet.set_value(row, col, &format!("{}", i));
        }

        let snapshot = SheetSnapshot::from_sheet(&sheet);

        // Time 10k reads (mix of hits and misses)
        let start = std::time::Instant::now();
        for i in 0..10_000usize {
            let row = i / 100;
            let col = i % 100;
            let _ = snapshot.get_value(row, col);
        }
        let elapsed = start.elapsed();

        // Should be fast - HashMap lookup is O(1)
        assert!(
            elapsed.as_millis() < 10,
            "10k reads took {}ms - too slow!",
            elapsed.as_millis()
        );

        println!(
            "SheetSnapshot reads: 10k in {:?} ({:.2} ns/read)",
            elapsed,
            elapsed.as_nanos() as f64 / 10_000.0
        );
    }

    #[test]
    fn test_snapshot_empty_sheet_fast() {
        // Large dimensions but empty - should be instant
        let sheet = Sheet::new(SheetId(1), 10_000, 1_000);  // 10M grid, 0 cells

        let start = std::time::Instant::now();
        let snapshot = SheetSnapshot::from_sheet(&sheet);
        let elapsed = start.elapsed();

        assert_eq!(snapshot.values.len(), 0);

        // Should be essentially instant (under 1ms)
        assert!(
            elapsed.as_micros() < 1000,
            "Empty sheet snapshot took {:?} - sparse iteration broken!",
            elapsed
        );
    }
}
