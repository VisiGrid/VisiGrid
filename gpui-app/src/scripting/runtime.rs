//! Lua runtime for VisiGrid scripting.
//!
//! # Architecture Notes
//!
//! This runtime is intentionally isolated from workbook state. Sheet access is
//! through a command sink interface (`LuaOpSink`), not direct references.
//!
//! The runtime captures print() output and returns it along with evaluation results.
//! When evaluating with sheet access, operations are collected and returned for
//! batch application to the workbook.

use mlua::{Lua, MultiValue, Result as LuaResult, Value, HookTriggers, VmState};
use std::cell::RefCell;
use std::rc::Rc;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicI64, Ordering};
use std::time::{Duration, Instant};

use super::ops::{LuaOp, SheetReader};
use super::sheet_api::{DynOpSink, SheetUserData, MAX_OUTPUT_LINES};

/// Maximum number of Lua instructions per script execution.
/// 100 million instructions is enough for any reasonable spreadsheet script.
pub const INSTRUCTION_LIMIT: i64 = 100_000_000;

/// How often to check the instruction budget (every N instructions).
/// 10,000 means we call the hook ~10,000 times for a full budget.
pub const INSTRUCTION_HOOK_INTERVAL: u32 = 10_000;

/// Cancel token for script execution.
/// Set to true to signal the script should stop.
pub type CancelToken = Arc<AtomicBool>;

/// Default wall-clock timeout for script execution (30 seconds).
/// This catches pathological code patterns that burn instructions slowly.
pub const DEFAULT_TIMEOUT: Duration = Duration::from_secs(30);

/// Result of evaluating a Lua chunk
#[derive(Debug, Clone)]
pub struct LuaEvalResult {
    /// Lines printed via print()
    pub output: Vec<String>,
    /// String representation of returned value (if any)
    pub returned: Option<String>,
    /// Error message (if evaluation failed)
    pub error: Option<String>,
    /// Operations queued by the script (only when eval_with_sheet is used)
    pub ops: Vec<LuaOp>,
    /// Number of unique cells modified
    pub mutations: usize,
    /// Whether output was truncated due to limit
    pub output_truncated: bool,
    /// Whether execution was stopped due to instruction limit
    pub instruction_limit_exceeded: bool,
    /// Whether execution was cancelled via cancel token
    pub cancelled: bool,
    /// Whether execution was stopped due to wall-clock timeout
    pub timed_out: bool,
}

impl LuaEvalResult {
    fn success(output: Vec<String>, returned: Option<String>, truncated: bool) -> Self {
        Self {
            output,
            returned,
            error: None,
            ops: Vec::new(),
            mutations: 0,
            output_truncated: truncated,
            instruction_limit_exceeded: false,
            cancelled: false,
            timed_out: false,
        }
    }

    fn success_with_ops(output: Vec<String>, returned: Option<String>, ops: Vec<LuaOp>, mutations: usize, truncated: bool) -> Self {
        Self {
            output,
            returned,
            error: None,
            ops,
            mutations,
            output_truncated: truncated,
            instruction_limit_exceeded: false,
            cancelled: false,
            timed_out: false,
        }
    }

    fn error(output: Vec<String>, error: String, truncated: bool) -> Self {
        Self {
            output,
            returned: None,
            error: Some(error),
            ops: Vec::new(),
            mutations: 0,
            output_truncated: truncated,
            instruction_limit_exceeded: false,
            cancelled: false,
            timed_out: false,
        }
    }

    fn error_with_limit(output: Vec<String>, error: String, truncated: bool) -> Self {
        Self {
            output,
            returned: None,
            error: Some(error),
            ops: Vec::new(),
            mutations: 0,
            output_truncated: truncated,
            instruction_limit_exceeded: true,
            cancelled: false,
            timed_out: false,
        }
    }

    fn error_cancelled(output: Vec<String>, truncated: bool) -> Self {
        Self {
            output,
            returned: None,
            error: Some("execution cancelled".to_string()),
            ops: Vec::new(),
            mutations: 0,
            output_truncated: truncated,
            instruction_limit_exceeded: false,
            cancelled: true,
            timed_out: false,
        }
    }

    fn error_timed_out(output: Vec<String>, truncated: bool) -> Self {
        Self {
            output,
            returned: None,
            error: Some(format!("execution timeout ({}s limit)", DEFAULT_TIMEOUT.as_secs())),
            ops: Vec::new(),
            mutations: 0,
            output_truncated: truncated,
            instruction_limit_exceeded: false,
            cancelled: false,
            timed_out: true,
        }
    }

    /// Returns true if the script made any mutations
    pub fn has_mutations(&self) -> bool {
        self.mutations > 0
    }
}

/// Output buffer state (shared between print() and eval)
struct OutputState {
    lines: Vec<String>,
    truncated: bool,
}

impl OutputState {
    fn new() -> Self {
        Self {
            lines: Vec::new(),
            truncated: false,
        }
    }

    fn clear(&mut self) {
        self.lines.clear();
        self.truncated = false;
    }

    fn push(&mut self, line: String) {
        if self.lines.len() < MAX_OUTPUT_LINES {
            self.lines.push(line);
        } else if !self.truncated {
            self.truncated = true;
            // Don't push anything more - we've hit the limit
        }
    }
}

/// The Lua runtime for VisiGrid.
///
/// Owns the mlua::Lua instance and provides evaluation with output capture.
pub struct LuaRuntime {
    lua: Lua,
    /// Captured output from print() calls during evaluation
    output_state: Rc<RefCell<OutputState>>,
}

impl LuaRuntime {
    /// Create a new Lua runtime with sandboxed globals.
    pub fn new() -> LuaResult<Self> {
        let lua = Lua::new();

        // Create output state that print() will write to
        let output_state = Rc::new(RefCell::new(OutputState::new()));

        // Override print() to capture output (with cap)
        {
            let state = output_state.clone();
            let print_fn = lua.create_function(move |_, args: MultiValue| {
                let parts: Vec<String> = args
                    .into_iter()
                    .map(|v| lua_value_to_string(&v))
                    .collect();
                let line = parts.join("\t");
                state.borrow_mut().push(line);
                Ok(())
            })?;
            lua.globals().set("print", print_fn)?;
        }

        // Sandbox: remove dangerous globals
        // We keep: basic, string, table, math, utf8
        // We remove: os, io, debug, package, require, loadfile, dofile
        let globals = lua.globals();
        globals.set("os", Value::Nil)?;
        globals.set("io", Value::Nil)?;
        globals.set("debug", Value::Nil)?;
        globals.set("package", Value::Nil)?;
        globals.set("require", Value::Nil)?;
        globals.set("loadfile", Value::Nil)?;
        globals.set("dofile", Value::Nil)?;

        // Also remove load() which can execute arbitrary bytecode
        globals.set("load", Value::Nil)?;

        Ok(Self { lua, output_state })
    }

    /// Evaluate a Lua chunk with REPL-style behavior (no sheet access).
    ///
    /// If the input looks like an expression (not a statement), it's wrapped
    /// in `return (...)` so the value is returned. This makes the REPL feel
    /// natural: typing `1 + 1` shows `2`, not nothing.
    ///
    /// For sheet access, use `eval_with_sheet` instead.
    pub fn eval(&self, input: &str) -> LuaEvalResult {
        self.eval_internal(input, None, None)
    }

    /// Evaluate a Lua chunk with sheet access.
    ///
    /// The `sheet` global is registered, allowing the script to read/write cells.
    /// Operations are queued and returned in the result for batch application.
    ///
    /// # Arguments
    ///
    /// * `input` - The Lua code to evaluate
    /// * `reader` - A boxed SheetReader for workbook access
    ///
    /// # Returns
    ///
    /// A `LuaEvalResult` with `ops` populated with the queued operations.
    pub fn eval_with_sheet(&self, input: &str, reader: Box<dyn SheetReader>) -> LuaEvalResult {
        // Create sink and wrap it
        let sink = Rc::new(RefCell::new(DynOpSink::new(reader)));

        self.eval_internal(input, Some(sink), None)
    }

    /// Evaluate a Lua chunk with sheet access and selection info.
    ///
    /// Same as `eval_with_sheet`, but passes the current selection to Lua.
    ///
    /// # Arguments
    ///
    /// * `input` - The Lua code to evaluate
    /// * `reader` - A boxed SheetReader for workbook access
    /// * `selection` - Current selection as (start_row, start_col, end_row, end_col), 0-indexed
    ///
    /// # Returns
    ///
    /// A `LuaEvalResult` with the selection accessible via `sheet:selection()`.
    pub fn eval_with_sheet_and_selection(
        &self,
        input: &str,
        reader: Box<dyn SheetReader>,
        selection: (usize, usize, usize, usize),
    ) -> LuaEvalResult {
        let sink = Rc::new(RefCell::new(DynOpSink::with_selection(reader, selection)));
        self.eval_internal(input, Some(sink), None)
    }

    /// Evaluate a Lua chunk with sheet access and cancellation support.
    ///
    /// Same as `eval_with_sheet`, but accepts a cancel token that can be used
    /// to stop execution early. Set the token to `true` from another thread
    /// (e.g., UI thread) to signal cancellation.
    ///
    /// # Arguments
    ///
    /// * `input` - The Lua code to evaluate
    /// * `reader` - A boxed SheetReader for workbook access
    /// * `cancel` - Optional cancel token to signal early termination
    ///
    /// # Returns
    ///
    /// A `LuaEvalResult` with `cancelled: true` if execution was stopped.
    pub fn eval_with_sheet_cancellable(
        &self,
        input: &str,
        reader: Box<dyn SheetReader>,
        cancel: Option<CancelToken>,
    ) -> LuaEvalResult {
        let sink = Rc::new(RefCell::new(DynOpSink::new(reader)));
        self.eval_internal(input, Some(sink), cancel)
    }

    /// Internal eval implementation
    fn eval_internal(
        &self,
        input: &str,
        sink: Option<Rc<RefCell<DynOpSink>>>,
        cancel: Option<CancelToken>,
    ) -> LuaEvalResult {
        // Clear output state
        self.output_state.borrow_mut().clear();

        let trimmed = input.trim();
        if trimmed.is_empty() {
            return LuaEvalResult::success(vec![], None, false);
        }

        // Register sheet global if sink provided
        if let Some(ref sink) = sink {
            let userdata = SheetUserData::new(sink.clone());
            if let Err(e) = self.lua.globals().set("sheet", userdata) {
                return LuaEvalResult::error(vec![], format!("Failed to register sheet: {}", e), false);
            }
        }

        // Try expression-first: if it parses as `return (input)`, use that
        let (code, is_expression) = self.prepare_code(trimmed);

        // Set up instruction limit hook (also checks cancel flag and timeout)
        let start_time = Instant::now();
        let budget = Arc::new(AtomicI64::new(INSTRUCTION_LIMIT));
        let budget_clone = budget.clone();
        let cancel_clone = cancel.clone();
        let was_cancelled = Arc::new(AtomicBool::new(false));
        let was_cancelled_clone = was_cancelled.clone();
        let was_timed_out = Arc::new(AtomicBool::new(false));
        let was_timed_out_clone = was_timed_out.clone();

        self.lua.set_hook(
            HookTriggers::new().every_nth_instruction(INSTRUCTION_HOOK_INTERVAL),
            move |_lua, _debug| {
                // Check cancel flag first
                if let Some(ref cancel) = cancel_clone {
                    if cancel.load(Ordering::Relaxed) {
                        was_cancelled_clone.store(true, Ordering::Relaxed);
                        return Err(mlua::Error::RuntimeError("execution cancelled".to_string()));
                    }
                }

                // Check wall-clock timeout
                if start_time.elapsed() > DEFAULT_TIMEOUT {
                    was_timed_out_clone.store(true, Ordering::Relaxed);
                    return Err(mlua::Error::RuntimeError(
                        format!("execution timeout ({}s limit)", DEFAULT_TIMEOUT.as_secs())
                    ));
                }

                // Check instruction budget
                let remaining = budget_clone.fetch_sub(INSTRUCTION_HOOK_INTERVAL as i64, Ordering::Relaxed);
                if remaining <= 0 {
                    Err(mlua::Error::RuntimeError(
                        format!("instruction limit exceeded ({} instructions)", INSTRUCTION_LIMIT)
                    ))
                } else {
                    Ok(VmState::Continue)
                }
            },
        );

        // Execute
        let result = self.lua.load(&code).eval::<MultiValue>();

        // Remove the instruction hook
        self.lua.remove_hook();

        // Remove sheet global after execution (cleanup)
        if sink.is_some() {
            let _ = self.lua.globals().set("sheet", Value::Nil);
        }

        // Collect output and truncation status
        let state = self.output_state.borrow();
        let mut output = state.lines.clone();
        let truncated = state.truncated;
        drop(state);

        // Add truncation notice if needed
        if truncated {
            output.push(format!("... output truncated ({} line limit)", MAX_OUTPUT_LINES));
        }

        // Extract ops from sink
        let (ops, mutations) = if let Some(sink) = sink {
            let mut borrowed = sink.borrow_mut();
            let mutations = borrowed.mutations();
            let ops = borrowed.take_ops();
            (ops, mutations)
        } else {
            (Vec::new(), 0)
        };

        // Check if instruction limit, cancellation, or timeout was hit
        let instruction_limit_hit = budget.load(Ordering::Relaxed) <= 0;
        let cancelled = was_cancelled.load(Ordering::Relaxed);
        let timed_out = was_timed_out.load(Ordering::Relaxed);

        match result {
            Ok(values) => {
                // Format returned values
                let returned = if values.is_empty() {
                    None
                } else if is_expression || !values.iter().all(|v| matches!(v, Value::Nil)) {
                    // Show non-nil returns, or any return if it was an expression
                    let parts: Vec<String> =
                        values.iter().map(|v| lua_value_to_string(v)).collect();
                    let joined = parts.join(", ");
                    if joined == "nil" && !is_expression {
                        None
                    } else {
                        Some(joined)
                    }
                } else {
                    None
                };
                LuaEvalResult::success_with_ops(output, returned, ops, mutations, truncated)
            }
            Err(e) => {
                // On error, check if it was instruction limit, cancellation, or timeout
                let error_msg = format_lua_error(&e);
                let mut result = if cancelled {
                    LuaEvalResult::error_cancelled(output, truncated)
                } else if timed_out {
                    LuaEvalResult::error_timed_out(output, truncated)
                } else if instruction_limit_hit {
                    LuaEvalResult::error_with_limit(output, error_msg, truncated)
                } else {
                    LuaEvalResult::error(output, error_msg, truncated)
                };
                result.ops = ops;
                result.mutations = mutations;
                result
            }
        }
    }

    /// Prepare code for execution.
    ///
    /// Returns (code_to_execute, was_expression).
    fn prepare_code(&self, input: &str) -> (String, bool) {
        // First, check if it's a valid expression by trying to parse `return (input)`
        let as_expr = format!("return ({})", input);
        if self.lua.load(&as_expr).into_function().is_ok() {
            // It's a valid expression
            return (as_expr, true);
        }

        // Otherwise, execute as-is (statement or multi-statement)
        (input.to_string(), false)
    }

    /// Get a reference to the Lua instance (for M2 extension).
    ///
    /// This should only be used to register the `sheet` userdata, not to
    /// give direct workbook access.
    #[allow(dead_code)]
    pub(crate) fn lua(&self) -> &Lua {
        &self.lua
    }
}

impl Default for LuaRuntime {
    fn default() -> Self {
        Self::new().expect("Failed to create Lua runtime")
    }
}

/// Convert a Lua value to a display string.
fn lua_value_to_string(value: &Value) -> String {
    match value {
        Value::Nil => "nil".to_string(),
        Value::Boolean(b) => b.to_string(),
        Value::Integer(i) => i.to_string(),
        Value::Number(n) => {
            // Format nicely: no trailing zeros for integers
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            }
        }
        Value::String(s) => s.to_str().map(|s| s.to_string()).unwrap_or_else(|_| "<invalid utf8>".to_string()),
        Value::Table(_) => "table".to_string(),
        Value::Function(_) => "function".to_string(),
        Value::Thread(_) => "thread".to_string(),
        Value::UserData(_) => "userdata".to_string(),
        Value::LightUserData(_) => "lightuserdata".to_string(),
        Value::Error(e) => format!("error: {}", e),
        _ => "<unknown>".to_string(),
    }
}

/// Format a Lua error for display.
fn format_lua_error(error: &mlua::Error) -> String {
    match error {
        mlua::Error::SyntaxError { message, .. } => {
            // Strip the "[string \"...\"]:1: " prefix if present
            if let Some(idx) = message.find("]: ") {
                message[idx + 3..].to_string()
            } else {
                message.clone()
            }
        }
        mlua::Error::RuntimeError(msg) => msg.clone(),
        mlua::Error::CallbackError { cause, .. } => format_lua_error(cause),
        _ => error.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_basic_expression() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("1 + 1");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("2".to_string()));
    }

    #[test]
    fn test_string_expression() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("'hello' .. ' ' .. 'world'");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("hello world".to_string()));
    }

    #[test]
    fn test_print_capture() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("print('hello', 'world')");
        assert!(result.error.is_none());
        assert_eq!(result.output, vec!["hello\tworld"]);
    }

    #[test]
    fn test_multiple_prints() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("print('one'); print('two'); print('three')");
        assert!(result.error.is_none());
        assert_eq!(result.output, vec!["one", "two", "three"]);
    }

    #[test]
    fn test_statement_no_return() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("local x = 42");
        assert!(result.error.is_none());
        assert!(result.returned.is_none() || result.returned == Some("nil".to_string()));
    }

    #[test]
    fn test_for_loop() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("for i = 1, 3 do print(i) end");
        assert!(result.error.is_none());
        assert_eq!(result.output, vec!["1", "2", "3"]);
    }

    #[test]
    fn test_syntax_error() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("if then");
        assert!(result.error.is_some());
        // Error should be present and readable
        let err = result.error.unwrap();
        assert!(!err.is_empty());
    }

    #[test]
    fn test_runtime_error() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("error('oops')");
        assert!(result.error.is_some());
        assert!(result.error.unwrap().contains("oops"));
    }

    #[test]
    fn test_sandbox_no_os() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("os.execute('ls')");
        assert!(result.error.is_some());
    }

    #[test]
    fn test_sandbox_no_io() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("io.open('/etc/passwd')");
        assert!(result.error.is_some());
    }

    #[test]
    fn test_sandbox_no_require() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("require('os')");
        assert!(result.error.is_some());
    }

    #[test]
    fn test_sandbox_no_load() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("load('return 1')()");
        assert!(result.error.is_some());
    }

    #[test]
    fn test_table_expression() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("{1, 2, 3}");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("table".to_string()));
    }

    #[test]
    fn test_function_expression() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("function() end");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("function".to_string()));
    }

    #[test]
    fn test_multiline_code() {
        let rt = LuaRuntime::new().unwrap();
        let code = r#"
            local sum = 0
            for i = 1, 10 do
                sum = sum + i
            end
            return sum
        "#;
        let result = rt.eval(code);
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("55".to_string()));
    }

    #[test]
    fn test_math_library_available() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("math.floor(3.7)");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("3".to_string()));
    }

    #[test]
    fn test_string_library_available() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("string.upper('hello')");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("HELLO".to_string()));
    }

    #[test]
    fn test_table_library_available() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("table.concat({'a', 'b', 'c'}, ',')");
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("a,b,c".to_string()));
    }

    #[test]
    fn test_empty_input() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("");
        assert!(result.error.is_none());
        assert!(result.returned.is_none());
        assert!(result.output.is_empty());
    }

    #[test]
    fn test_whitespace_input() {
        let rt = LuaRuntime::new().unwrap();
        let result = rt.eval("   \n\t  ");
        assert!(result.error.is_none());
        assert!(result.returned.is_none());
    }

    // ========================================================================
    // M2 Tests: eval_with_sheet
    // ========================================================================

    use super::super::ops::LuaCellValue;
    use std::collections::HashMap;

    /// Mock sheet reader for testing
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

        fn with_formula(mut self, row: usize, col: usize, formula: &str) -> Self {
            self.formulas.insert((row, col), formula.to_string());
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

        fn rows(&self) -> usize { 100 }
        fn cols(&self) -> usize { 26 }
    }

    #[test]
    fn test_sheet_read_value() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_value(0, 0, LuaCellValue::Number(42.0));  // Internal 0-indexed

        let result = rt.eval_with_sheet("sheet:get_value(1, 1)", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("42".to_string()));
        assert_eq!(result.ops.len(), 0);  // No writes
    }

    #[test]
    fn test_sheet_write_value() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let result = rt.eval_with_sheet("sheet:set_value(1, 1, 100)", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.ops.len(), 1);
        assert_eq!(result.mutations, 1);

        // Check the op
        match &result.ops[0] {
            LuaOp::SetValue { row, col, value } => {
                assert_eq!(*row, 0);  // 0-indexed internally
                assert_eq!(*col, 0);
                assert_eq!(*value, LuaCellValue::Number(100.0));
            }
            _ => panic!("Expected SetValue"),
        }
    }

    #[test]
    fn test_sheet_read_after_write() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Write then read - should see the pending value
        let code = r#"
            sheet:set_value(1, 1, 42)
            return sheet:get_value(1, 1)
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("42".to_string()));
        assert_eq!(result.mutations, 1);
    }

    #[test]
    fn test_sheet_a1_notation() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_value(0, 0, LuaCellValue::Number(99.0));

        let result = rt.eval_with_sheet("sheet:get_a1('A1')", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("99".to_string()));
    }

    #[test]
    fn test_sheet_set_a1() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            sheet:set_a1('B2', 'hello')
            return sheet:get_a1('B2')
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("hello".to_string()));

        // B2 = row 2, col 2 (1-indexed) = row 1, col 1 (0-indexed)
        match &result.ops[0] {
            LuaOp::SetValue { row, col, .. } => {
                assert_eq!(*row, 1);  // 0-indexed
                assert_eq!(*col, 1);
            }
            _ => panic!("Expected SetValue"),
        }
    }

    #[test]
    fn test_sheet_set_formula() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let result = rt.eval_with_sheet("sheet:set_formula(1, 1, '=A2*2')", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);

        match &result.ops[0] {
            LuaOp::SetFormula { formula, .. } => {
                assert_eq!(formula, "=A2*2");
            }
            _ => panic!("Expected SetFormula"),
        }
    }

    #[test]
    fn test_sheet_get_formula() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_formula(0, 0, "=SUM(B1:B10)");

        let result = rt.eval_with_sheet("sheet:get_formula(1, 1)", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("=SUM(B1:B10)".to_string()));
    }

    #[test]
    fn test_sheet_rows_cols() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let result = rt.eval_with_sheet("return sheet:rows(), sheet:cols()", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("100, 26".to_string()));
    }

    #[test]
    fn test_sheet_multiple_writes() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            sheet:set_value(1, 1, 10)
            sheet:set_value(1, 2, 20)
            sheet:set_value(2, 1, 30)
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.ops.len(), 3);
        assert_eq!(result.mutations, 3);  // 3 unique cells
    }

    #[test]
    fn test_sheet_overwrite_same_cell() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            sheet:set_value(1, 1, 10)
            sheet:set_value(1, 1, 20)
            return sheet:get_value(1, 1)
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("20".to_string()));  // Latest value
        assert_eq!(result.ops.len(), 2);  // Both ops recorded
        assert_eq!(result.mutations, 1);  // Only 1 unique cell
    }

    #[test]
    fn test_sheet_nil_clears_value() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_value(0, 0, LuaCellValue::Number(42.0));

        // Set value to nil to clear it
        let result = rt.eval_with_sheet("sheet:set_value(1, 1, nil)", Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.ops.len(), 1);

        // Check the op contains Nil (clearing the cell)
        match &result.ops[0] {
            LuaOp::SetValue { value, .. } => {
                assert_eq!(*value, LuaCellValue::Nil);
            }
            _ => panic!("Expected SetValue"),
        }
    }

    #[test]
    fn test_sheet_global_cleaned_up() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // First eval with sheet
        let _ = rt.eval_with_sheet("sheet:set_value(1, 1, 42)", Box::new(reader));

        // Second eval WITHOUT sheet - should error because sheet is nil
        let result = rt.eval("sheet:get_value(1, 1)");
        assert!(result.error.is_some(), "Expected error - sheet should be nil");
    }

    #[test]
    fn test_sheet_for_loop_fill() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Fill column A with 1-10
        let code = r#"
            for i = 1, 10 do
                sheet:set_value(i, 1, i * 10)
            end
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.ops.len(), 10);
        assert_eq!(result.mutations, 10);

        // Verify values
        match &result.ops[0] {
            LuaOp::SetValue { row, col, value } => {
                assert_eq!(*row, 0);  // First row (0-indexed)
                assert_eq!(*col, 0);
                assert_eq!(*value, LuaCellValue::Number(10.0));
            }
            _ => panic!("Expected SetValue"),
        }
        match &result.ops[9] {
            LuaOp::SetValue { row, value, .. } => {
                assert_eq!(*row, 9);  // 10th row (0-indexed)
                assert_eq!(*value, LuaCellValue::Number(100.0));
            }
            _ => panic!("Expected SetValue"),
        }
    }

    #[test]
    fn test_sheet_selection_single_cell() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Selection at B3 (row 2, col 1 in 0-indexed) = single cell
        let selection = (2, 1, 2, 1);  // 0-indexed

        let code = r#"
            local sel = sheet:selection()
            return sel.start_row, sel.start_col, sel.end_row, sel.end_col, sel.range
        "#;
        let result = rt.eval_with_sheet_and_selection(code, Box::new(reader), selection);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        // Lua returns 1-indexed values: row 3, col 2 = B3
        assert_eq!(result.returned, Some("3, 2, 3, 2, B3".to_string()));
    }

    #[test]
    fn test_sheet_selection_range() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Selection A1:C5 (row 0, col 0 to row 4, col 2 in 0-indexed)
        let selection = (0, 0, 4, 2);

        let code = r#"
            local sel = sheet:selection()
            return sel.start_row, sel.start_col, sel.end_row, sel.end_col, sel.range
        "#;
        let result = rt.eval_with_sheet_and_selection(code, Box::new(reader), selection);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        // Lua returns 1-indexed: A1:C5
        assert_eq!(result.returned, Some("1, 1, 5, 3, A1:C5".to_string()));
    }

    #[test]
    fn test_sheet_get_set_shorthand() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Test get/set shorthand with A1 notation
        let code = r#"
            sheet:set("A1", 42)
            sheet:set("B2", "hello")
            return sheet:get("A1"), sheet:get("B2")
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("42, hello".to_string()));
        assert_eq!(result.ops.len(), 2);
    }

    #[test]
    fn test_sheet_rollback_discards_changes() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Write some values, then rollback
        let code = r#"
            sheet:set("A1", 1)
            sheet:set("A2", 2)
            sheet:set("A3", 3)
            local discarded = sheet:rollback()
            return discarded
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("3".to_string()));  // 3 ops discarded
        assert_eq!(result.ops.len(), 0);  // No ops after rollback
        assert_eq!(result.mutations, 0);  // No mutations after rollback
    }

    #[test]
    fn test_sheet_begin_commit_noop() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // begin/commit should work but not affect behavior
        let code = r#"
            sheet:begin()
            sheet:set("A1", 42)
            sheet:commit()
            return sheet:get("A1")
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("42".to_string()));
        assert_eq!(result.ops.len(), 1);  // Op still recorded
    }

    #[test]
    fn test_sheet_rollback_then_continue() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Rollback then continue writing
        let code = r#"
            sheet:set("A1", 1)
            sheet:set("A2", 2)
            sheet:rollback()  -- discard first 2
            sheet:set("A3", 3)
            sheet:set("A4", 4)
            return sheet:get("A3"), sheet:get("A4")
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("3, 4".to_string()));
        assert_eq!(result.ops.len(), 2);  // Only ops after rollback
    }

    // ========================================================================
    // Range API Tests
    // ========================================================================

    #[test]
    fn test_range_values_read() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_value(0, 0, LuaCellValue::Number(1.0))  // A1
            .with_value(0, 1, LuaCellValue::Number(2.0))  // B1
            .with_value(1, 0, LuaCellValue::Number(3.0))  // A2
            .with_value(1, 1, LuaCellValue::Number(4.0)); // B2

        let code = r#"
            local r = sheet:range("A1:B2")
            local v = r:values()
            return v[1][1], v[1][2], v[2][1], v[2][2]
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("1, 2, 3, 4".to_string()));
    }

    #[test]
    fn test_range_set_values_write() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            local r = sheet:range("A1:B2")
            r:set_values({{10, 20}, {30, 40}})
            return sheet:get("A1"), sheet:get("B1"), sheet:get("A2"), sheet:get("B2")
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("10, 20, 30, 40".to_string()));
        assert_eq!(result.ops.len(), 4);
    }

    #[test]
    fn test_range_info_methods() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            local r = sheet:range("A1:C5")
            return r:rows(), r:cols(), r:address()
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("5, 3, A1:C5".to_string()));
    }

    #[test]
    fn test_range_single_cell() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new()
            .with_value(0, 0, LuaCellValue::Number(42.0));

        let code = r#"
            local r = sheet:range("A1")
            local v = r:values()
            return v[1][1], r:rows(), r:cols(), r:address()
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("42, 1, 1, A1".to_string()));
    }

    #[test]
    fn test_range_invalid_error() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            local r = sheet:range("invalid")
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));
        assert!(result.error.is_some());
        assert!(result.error.unwrap().contains("Invalid range"));
    }

    // ========================================================================
    // M4 Hardening Tests
    // ========================================================================

    #[test]
    fn test_infinite_loop_stopped_by_instruction_limit() {
        let rt = LuaRuntime::new().unwrap();

        // An infinite loop that does nothing but spin
        let result = rt.eval("while true do end");

        // Should error with instruction limit exceeded
        assert!(result.error.is_some(), "Expected error for infinite loop");
        assert!(result.instruction_limit_exceeded, "Expected instruction_limit_exceeded flag");
        assert!(result.error.unwrap().contains("instruction limit exceeded"));
    }

    #[test]
    fn test_output_cap_truncates() {
        let rt = LuaRuntime::new().unwrap();

        // Print way more than the output cap
        let code = r#"
            for i = 1, 10000 do
                print("line " .. i)
            end
        "#;
        let result = rt.eval(code);

        // Should succeed but truncate output
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert!(result.output_truncated, "Expected output to be truncated");

        // Output should be capped at MAX_OUTPUT_LINES plus truncation notice
        assert!(result.output.len() <= super::super::sheet_api::MAX_OUTPUT_LINES + 1);

        // Last line should be truncation notice
        let last_line = result.output.last().unwrap();
        assert!(last_line.contains("truncated"), "Expected truncation notice, got: {}", last_line);
    }

    #[test]
    fn test_ops_cap_triggers() {
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Try to write way more ops than allowed
        // Note: We can't actually hit 1M ops in a reasonable test time,
        // so we test the mechanism with a smaller loop and verify
        // that the cap error message is correct when triggered
        let code = r#"
            for i = 1, 10000 do
                for j = 1, 26 do
                    sheet:set_value(i, j, i * j)
                end
            end
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));

        // This should either complete (if under limit) or error with ops cap
        // With 260k ops it should complete
        if result.error.is_some() {
            let err = result.error.as_ref().unwrap();
            // If it errors, it should be because of ops or instruction limit
            assert!(
                err.contains("operation limit exceeded") || err.contains("instruction limit exceeded"),
                "Unexpected error: {}", err
            );
        } else {
            // Should have many ops
            assert!(result.ops.len() > 100000, "Expected many ops, got {}", result.ops.len());
        }
    }

    #[test]
    fn test_cancel_flag_stops_execution() {
        use std::sync::atomic::AtomicBool;

        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        // Create a cancel token that's already set
        let cancel = Arc::new(AtomicBool::new(true));

        // Try to run code with an already-cancelled token
        let result = rt.eval_with_sheet_cancellable(
            "for i = 1, 1000000 do sheet:set_value(1, 1, i) end",
            Box::new(reader),
            Some(cancel),
        );

        // Should be cancelled
        assert!(result.cancelled, "Expected execution to be cancelled");
        assert!(result.error.is_some());
        assert!(result.error.unwrap().contains("cancelled"));
    }

    #[test]
    fn test_normal_script_completes_successfully() {
        // Verify that hardening doesn't break normal scripts
        let rt = LuaRuntime::new().unwrap();
        let reader = MockReader::new();

        let code = r#"
            -- A normal script that does reasonable work
            local sum = 0
            for i = 1, 100 do
                for j = 1, 26 do
                    sheet:set_value(i, j, i + j)
                    sum = sum + 1
                end
            end
            return sum
        "#;
        let result = rt.eval_with_sheet(code, Box::new(reader));

        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("2600".to_string()));
        assert_eq!(result.mutations, 2600);
        assert!(!result.output_truncated);
        assert!(!result.instruction_limit_exceeded);
        assert!(!result.cancelled);
        assert!(!result.timed_out);
    }
}
