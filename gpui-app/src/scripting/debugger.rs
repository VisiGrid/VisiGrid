//! Lua debugger thread scaffold (Phase 1).
//!
//! Spawns a background thread with its own Lua VM, runs a script, and sends
//! results back via `mpsc` channels. No debug hook yet — that's Phase 2.
//!
//! # Architecture
//!
//! - `spawn_debug_session()` creates channels + a named thread
//! - The thread creates its own `Lua` VM (independent of `LuaRuntime`)
//! - Sheet globals, custom functions, and sandbox are set up identically to `LuaRuntime`
//! - `DebuggerIntrospection` captures `debug.getlocal/getupvalue/getinfo` before sandbox
//! - Results are sent back as `DebugEvent` messages

use std::cell::RefCell;
use std::rc::Rc;
use std::sync::atomic::{AtomicBool, AtomicI64, AtomicU64, Ordering};
use std::sync::{mpsc, Arc};
use std::time::{Duration, Instant};

use mlua::{Function as LuaFunction, HookTriggers, Lua, MultiValue, Value, VmState};

use super::custom_functions::load_custom_functions;
use super::ops::LuaOp;
use super::runtime::{
    format_lua_error, lua_value_to_string, prepare_code, LuaEvalResult, DEFAULT_TIMEOUT,
    INSTRUCTION_HOOK_INTERVAL, INSTRUCTION_LIMIT,
};
use super::sheet_api::{register_sheet_global_with_selection, SheetSnapshot, MAX_OUTPUT_LINES};

// =============================================================================
// Identity
// =============================================================================

/// Unique identifier for a debug session.
pub type SessionId = u64;

/// Monotonically increasing session ID counter.
static NEXT_SESSION_ID: AtomicU64 = AtomicU64::new(1);

// =============================================================================
// Commands (UI → debug thread)
// =============================================================================

/// A command sent from the UI to the debug thread.
#[derive(Debug)]
pub struct DebugCommand {
    pub session_id: SessionId,
    pub action: DebugAction,
}

/// The action payload of a debug command.
#[derive(Debug)]
pub enum DebugAction {
    Continue,
    StepIn,
    StepOver,
    StepOut,
    Stop,
    AddBreakpoint { line: usize },
    RemoveBreakpoint { line: usize },
    ExpandVariable { path: Vec<VarPathSegment> },
}

/// A segment in a variable expansion path.
#[derive(Debug, Clone)]
pub enum VarPathSegment {
    Local(usize),
    Upvalue(usize),
    KeyString(String),
    KeyInt(i64),
    KeyOther(String),
}

// =============================================================================
// Events (debug thread → UI)
// =============================================================================

/// An event sent from the debug thread to the UI.
#[derive(Debug)]
pub struct DebugEvent {
    pub session_id: SessionId,
    pub payload: DebugEventPayload,
}

/// The payload of a debug event.
#[derive(Debug)]
pub enum DebugEventPayload {
    /// Script execution paused (Phase 2+).
    Paused(DebugSnapshot),
    /// A chunk of print() output.
    OutputChunk(String),
    /// Script finished (success or script-level error).
    Completed(LuaEvalResult),
    /// Catastrophic failure (VM creation failed, sandbox failed, etc.).
    Error(String),
    /// Response to ExpandVariable (Phase 3+).
    VariableExpanded {
        path: Vec<VarPathSegment>,
        children: Vec<Variable>,
    },
}

// =============================================================================
// Snapshot types (stubs for Phase 1, populated in Phase 2/3)
// =============================================================================

/// Snapshot of debugger state at a pause point.
#[derive(Debug, Clone)]
pub struct DebugSnapshot {
    pub line: usize,
    pub call_stack: Vec<StackFrame>,
    pub locals: Vec<Variable>,
    pub upvalues: Vec<Variable>,
    pub reason: PauseReason,
}

/// A single stack frame.
#[derive(Debug, Clone)]
pub struct StackFrame {
    pub function_name: Option<String>,
    pub source: Option<String>,
    pub line: usize,
    pub is_tail_call: bool,
}

/// A variable visible at a pause point.
#[derive(Debug, Clone)]
pub struct Variable {
    pub name: String,
    pub value: String,
    pub expandable: bool,
}

/// Why execution paused.
#[derive(Debug, Clone)]
pub enum PauseReason {
    Breakpoint,
    StepIn,
    StepOver,
    StepOut,
    Entry,
}

// =============================================================================
// Session
// =============================================================================

/// Handle returned by `spawn_debug_session`.
pub struct DebugSession {
    pub id: SessionId,
    pub state: DebugSessionState,
    pub cmd_tx: mpsc::Sender<DebugCommand>,
    pub event_rx: mpsc::Receiver<DebugEvent>,
}

/// Current state of a debug session.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DebugSessionState {
    Starting,
    Running,
    Paused,
    Completed,
}

// =============================================================================
// Configuration
// =============================================================================

/// Configuration for launching a debug session.
pub struct DebugConfig {
    /// The Lua source code to execute.
    pub code: String,
    /// Snapshot of the sheet data (concrete, `Send`-safe).
    pub snapshot: SheetSnapshot,
    /// Current selection as (start_row, start_col, end_row, end_col), 0-indexed.
    pub selection: (usize, usize, usize, usize),
}

// =============================================================================
// Private: DebuggerIntrospection
// =============================================================================

/// Captured debug introspection functions from the `debug` library.
///
/// These are captured before sandboxing removes the `debug` global.
/// Unused in Phase 1 — will be used by the debug hook in Phase 2.
#[allow(dead_code)]
struct DebuggerIntrospection {
    getlocal: LuaFunction,
    getupvalue: LuaFunction,
    getinfo: LuaFunction,
}

// =============================================================================
// Private: OutputBuffer
// =============================================================================

/// Coalesces print() output into time/size-bounded chunks for streaming.
///
/// Flushes every 16ms or 4KB, whichever comes first.
struct OutputBuffer {
    buf: String,
    last_flush: Instant,
    session_id: SessionId,
    event_tx: mpsc::Sender<DebugEvent>,
}

const OUTPUT_FLUSH_INTERVAL: Duration = Duration::from_millis(16);
const OUTPUT_FLUSH_SIZE: usize = 4096;

impl OutputBuffer {
    fn new(session_id: SessionId, event_tx: mpsc::Sender<DebugEvent>) -> Self {
        Self {
            buf: String::new(),
            last_flush: Instant::now(),
            session_id,
            event_tx,
        }
    }

    fn push_line(&mut self, line: &str) {
        if !self.buf.is_empty() {
            self.buf.push('\n');
        }
        self.buf.push_str(line);

        if self.buf.len() >= OUTPUT_FLUSH_SIZE
            || self.last_flush.elapsed() >= OUTPUT_FLUSH_INTERVAL
        {
            self.flush();
        }
    }

    fn flush(&mut self) {
        if !self.buf.is_empty() {
            let chunk = std::mem::take(&mut self.buf);
            let _ = self.event_tx.send(DebugEvent {
                session_id: self.session_id,
                payload: DebugEventPayload::OutputChunk(chunk),
            });
            self.last_flush = Instant::now();
        }
    }
}

// =============================================================================
// Public API
// =============================================================================

/// Spawn a new debug session on a background thread.
///
/// Returns a `DebugSession` handle with channels for communication.
/// The thread creates its own Lua VM, sets up the environment identically to
/// `LuaRuntime`, runs the script, and sends `Completed`/`Error` back.
pub fn spawn_debug_session(config: DebugConfig) -> DebugSession {
    let session_id = NEXT_SESSION_ID.fetch_add(1, Ordering::Relaxed);

    let (cmd_tx, cmd_rx) = mpsc::channel::<DebugCommand>();
    let (event_tx, event_rx) = mpsc::channel::<DebugEvent>();

    let thread_name = format!("debug-session-{}", session_id);
    std::thread::Builder::new()
        .name(thread_name)
        .spawn(move || {
            debug_thread_main(session_id, config, cmd_rx, event_tx);
        })
        .expect("Failed to spawn debug thread");

    DebugSession {
        id: session_id,
        state: DebugSessionState::Starting,
        cmd_tx,
        event_rx,
    }
}

// =============================================================================
// Thread entry point
// =============================================================================

fn debug_thread_main(
    session_id: SessionId,
    config: DebugConfig,
    _cmd_rx: mpsc::Receiver<DebugCommand>,
    event_tx: mpsc::Sender<DebugEvent>,
) {
    // 1. Create Lua VM
    let lua = Lua::new();

    // 2. Capture debug introspection before sandbox
    let _introspection = capture_debug_introspection(&lua);

    // 3. Sandbox: remove dangerous globals
    if let Err(e) = sandbox_lua(&lua) {
        let _ = event_tx.send(DebugEvent {
            session_id,
            payload: DebugEventPayload::Error(format!("Failed to sandbox Lua: {}", e)),
        });
        return;
    }

    // 4. Override print() with OutputBuffer-based coalescing + line collection
    let output_buffer = Rc::new(RefCell::new(OutputBuffer::new(session_id, event_tx.clone())));
    let output_lines: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let output_truncated = Rc::new(RefCell::new(false));

    {
        let ob = output_buffer.clone();
        let lines = output_lines.clone();
        let truncated = output_truncated.clone();
        let print_fn = match lua.create_function(move |_, args: MultiValue| {
            let parts: Vec<String> = args
                .into_iter()
                .map(|v| lua_value_to_string(&v))
                .collect();
            let line = parts.join("\t");

            // Collect into line buffer (for LuaEvalResult.output)
            {
                let mut lines = lines.borrow_mut();
                if lines.len() < MAX_OUTPUT_LINES {
                    lines.push(line.clone());
                } else if !*truncated.borrow() {
                    *truncated.borrow_mut() = true;
                }
            }

            // Stream via OutputBuffer
            ob.borrow_mut().push_line(&line);

            Ok(())
        }) {
            Ok(f) => f,
            Err(e) => {
                let _ = event_tx.send(DebugEvent {
                    session_id,
                    payload: DebugEventPayload::Error(format!(
                        "Failed to create print function: {}",
                        e
                    )),
                });
                return;
            }
        };

        if let Err(e) = lua.globals().set("print", print_fn) {
            let _ = event_tx.send(DebugEvent {
                session_id,
                payload: DebugEventPayload::Error(format!(
                    "Failed to register print: {}",
                    e
                )),
            });
            return;
        }
    }

    // 5. Register sheet global with selection
    let sink = match register_sheet_global_with_selection(
        &lua,
        Box::new(config.snapshot),
        config.selection,
    ) {
        Ok(s) => s,
        Err(e) => {
            let _ = event_tx.send(DebugEvent {
                session_id,
                payload: DebugEventPayload::Error(format!(
                    "Failed to register sheet global: {}",
                    e
                )),
            });
            return;
        }
    };

    // 6. Load custom functions (best-effort, don't fail on missing file)
    let _ = load_custom_functions(&lua);

    // 7. Prepare code (expression wrapping)
    let trimmed = config.code.trim();
    if trimmed.is_empty() {
        let result = LuaEvalResult {
            output: vec![],
            returned: None,
            error: None,
            ops: Vec::new(),
            mutations: 0,
            output_truncated: false,
            instruction_limit_exceeded: false,
            cancelled: false,
            timed_out: false,
        };
        let _ = event_tx.send(DebugEvent {
            session_id,
            payload: DebugEventPayload::Completed(result),
        });
        return;
    }

    let (code, is_expression) = prepare_code(&lua, trimmed);

    // 8. Set instruction hook (budget + timeout, no debug hook in Phase 1)
    let start_time = Instant::now();
    let budget = Arc::new(AtomicI64::new(INSTRUCTION_LIMIT));
    let budget_clone = budget.clone();
    let was_timed_out = Arc::new(AtomicBool::new(false));
    let was_timed_out_clone = was_timed_out.clone();

    lua.set_hook(
        HookTriggers::new().every_nth_instruction(INSTRUCTION_HOOK_INTERVAL),
        move |_lua, _debug| {
            if start_time.elapsed() > DEFAULT_TIMEOUT {
                was_timed_out_clone.store(true, Ordering::Relaxed);
                return Err(mlua::Error::RuntimeError(format!(
                    "execution timeout ({}s limit)",
                    DEFAULT_TIMEOUT.as_secs()
                )));
            }

            let remaining =
                budget_clone.fetch_sub(INSTRUCTION_HOOK_INTERVAL as i64, Ordering::Relaxed);
            if remaining <= 0 {
                Err(mlua::Error::RuntimeError(format!(
                    "instruction limit exceeded ({} instructions)",
                    INSTRUCTION_LIMIT
                )))
            } else {
                Ok(VmState::Continue)
            }
        },
    );

    // 9. Execute
    let result = lua.load(&code).eval::<MultiValue>();

    // 10. Remove hook, flush OutputBuffer
    lua.remove_hook();
    output_buffer.borrow_mut().flush();

    // 11. Collect output
    let mut output = output_lines.borrow().clone();
    let truncated = *output_truncated.borrow();
    if truncated {
        output.push(format!(
            "... output truncated ({} line limit)",
            MAX_OUTPUT_LINES
        ));
    }

    // 12. Collect ops from sink
    let (ops, mutations) = {
        let mut borrowed = sink.borrow_mut();
        let mutations = borrowed.mutations();
        let ops = borrowed.take_ops();
        (ops, mutations)
    };

    let instruction_limit_hit = budget.load(Ordering::Relaxed) <= 0;
    let timed_out = was_timed_out.load(Ordering::Relaxed);

    // 13. Build LuaEvalResult, send Completed
    let eval_result = match result {
        Ok(values) => {
            let returned = format_return_values(&values, is_expression);
            LuaEvalResult {
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
        Err(e) => {
            let error_msg = format_lua_error(&e);
            if timed_out {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some(format!(
                        "execution timeout ({}s limit)",
                        DEFAULT_TIMEOUT.as_secs()
                    )),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: false,
                    cancelled: false,
                    timed_out: true,
                }
            } else if instruction_limit_hit {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some(error_msg),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: true,
                    cancelled: false,
                    timed_out: false,
                }
            } else {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some(error_msg),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: false,
                    cancelled: false,
                    timed_out: false,
                }
            }
        }
    };

    let _ = event_tx.send(DebugEvent {
        session_id,
        payload: DebugEventPayload::Completed(eval_result),
    });
    // Thread exits
}

// =============================================================================
// Helpers
// =============================================================================

/// Capture `debug.getlocal`, `debug.getupvalue`, `debug.getinfo` before sandbox.
fn capture_debug_introspection(lua: &Lua) -> Option<DebuggerIntrospection> {
    let debug_table: mlua::Table = lua.globals().get("debug").ok()?;
    let getlocal: LuaFunction = debug_table.get("getlocal").ok()?;
    let getupvalue: LuaFunction = debug_table.get("getupvalue").ok()?;
    let getinfo: LuaFunction = debug_table.get("getinfo").ok()?;
    Some(DebuggerIntrospection {
        getlocal,
        getupvalue,
        getinfo,
    })
}

/// Sandbox the Lua VM by removing dangerous globals.
fn sandbox_lua(lua: &Lua) -> Result<(), mlua::Error> {
    let globals = lua.globals();
    globals.set("os", Value::Nil)?;
    globals.set("io", Value::Nil)?;
    globals.set("debug", Value::Nil)?;
    globals.set("package", Value::Nil)?;
    globals.set("require", Value::Nil)?;
    globals.set("loadfile", Value::Nil)?;
    globals.set("dofile", Value::Nil)?;
    globals.set("load", Value::Nil)?;
    Ok(())
}

/// Format return values from Lua execution.
fn format_return_values(values: &MultiValue, is_expression: bool) -> Option<String> {
    if values.is_empty() {
        return None;
    }
    if is_expression || !values.iter().all(|v| matches!(v, Value::Nil)) {
        let parts: Vec<String> = values.iter().map(|v| lua_value_to_string(v)).collect();
        let joined = parts.join(", ");
        if joined == "nil" && !is_expression {
            None
        } else {
            Some(joined)
        }
    } else {
        None
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use std::time::Duration;

    use super::super::ops::{LuaCellValue, SheetReader};

    /// Mock reader for tests
    struct MockReader {
        data: HashMap<(usize, usize), LuaCellValue>,
    }

    impl MockReader {
        fn new() -> Self {
            Self {
                data: HashMap::new(),
            }
        }
    }

    impl SheetReader for MockReader {
        fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
            self.data
                .get(&(row, col))
                .cloned()
                .unwrap_or(LuaCellValue::Nil)
        }
        fn get_formula(&self, _row: usize, _col: usize) -> Option<String> {
            None
        }
        fn rows(&self) -> usize {
            100
        }
        fn cols(&self) -> usize {
            26
        }
    }

    fn mock_snapshot() -> SheetSnapshot {
        SheetSnapshot::from_mock(HashMap::new(), HashMap::new(), 100, 26)
    }

    fn recv_completed(session: &DebugSession) -> LuaEvalResult {
        let timeout = Duration::from_secs(10);
        loop {
            match session.event_rx.recv_timeout(timeout) {
                Ok(event) => {
                    assert_eq!(event.session_id, session.id);
                    match event.payload {
                        DebugEventPayload::Completed(result) => return result,
                        DebugEventPayload::OutputChunk(_) => continue,
                        DebugEventPayload::Error(e) => panic!("Unexpected error event: {}", e),
                        _ => continue,
                    }
                }
                Err(e) => panic!("Timed out waiting for Completed event: {}", e),
            }
        }
    }

    #[test]
    fn test_basic_script_completes() {
        let session = spawn_debug_session(DebugConfig {
            code: "return 1 + 1".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let result = recv_completed(&session);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("2".to_string()));
    }

    #[test]
    fn test_print_output() {
        let session = spawn_debug_session(DebugConfig {
            code: "print('hello')".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let result = recv_completed(&session);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert!(
            result.output.contains(&"hello".to_string()),
            "Output should contain 'hello', got: {:?}",
            result.output
        );
    }

    #[test]
    fn test_sheet_ops_collected() {
        let session = spawn_debug_session(DebugConfig {
            code: "sheet:set('A1', 42)".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let result = recv_completed(&session);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert!(!result.ops.is_empty(), "Expected ops from sheet:set");
        match &result.ops[0] {
            LuaOp::SetValue { row, col, value } => {
                assert_eq!(*row, 0);
                assert_eq!(*col, 0);
                assert_eq!(*value, LuaCellValue::Number(42.0));
            }
            other => panic!("Expected SetValue, got: {:?}", other),
        }
    }

    #[test]
    fn test_error_handling() {
        let session = spawn_debug_session(DebugConfig {
            code: "if then".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let result = recv_completed(&session);
        assert!(result.error.is_some(), "Expected syntax error");
    }

    #[test]
    fn test_session_id_increments() {
        let s1 = spawn_debug_session(DebugConfig {
            code: "return 1".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let s2 = spawn_debug_session(DebugConfig {
            code: "return 2".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        assert_ne!(s1.id, s2.id);
        assert!(s2.id > s1.id);

        // Drain both
        recv_completed(&s1);
        recv_completed(&s2);
    }

    #[test]
    fn test_instruction_limit() {
        let session = spawn_debug_session(DebugConfig {
            code: "while true do end".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let result = recv_completed(&session);
        assert!(result.error.is_some(), "Expected error for infinite loop");
        assert!(
            result.instruction_limit_exceeded,
            "Expected instruction_limit_exceeded flag"
        );
    }

    #[test]
    fn test_expression_wrapping() {
        // Expression: should return a value
        let s1 = spawn_debug_session(DebugConfig {
            code: "1 + 1".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let r1 = recv_completed(&s1);
        assert!(r1.error.is_none(), "Error: {:?}", r1.error);
        assert_eq!(r1.returned, Some("2".to_string()));

        // Statement: should return None
        let s2 = spawn_debug_session(DebugConfig {
            code: "local x = 1".to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
        });
        let r2 = recv_completed(&s2);
        assert!(r2.error.is_none(), "Error: {:?}", r2.error);
        assert!(
            r2.returned.is_none() || r2.returned == Some("nil".to_string()),
            "Statement should not return a value, got: {:?}",
            r2.returned
        );
    }
}
