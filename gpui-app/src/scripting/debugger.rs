//! Lua debugger with hook-based pause/resume and variable inspection (Phase 3).
//!
//! Spawns a background thread with its own Lua VM, installs a debug hook for
//! breakpoints and stepping, and communicates via `mpsc` channels.
//!
//! # Architecture
//!
//! - `spawn_debug_session()` creates channels + a named thread
//! - The thread creates its own `Lua` VM (independent of `LuaRuntime`)
//! - A debug hook fires on every line, call, return, and nth instruction
//! - `HookState` tracks mode, breakpoints, stack depth, and budget
//! - When paused, the thread blocks on `cmd_rx.recv()` — no spinning
//! - `Arc<AtomicBool>` cancel flag enables stop from UI without breakpoint
//! - Variables are collected lazily: top frame on pause, other frames on request
//! - Table expansion is one level deep, capped at 50 children

use std::cell::RefCell;
use std::collections::HashSet;
use std::rc::Rc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{mpsc, Arc};
use std::time::{Duration, Instant};

use mlua::{
    DebugEvent as HookEvent, Function as LuaFunction, HookTriggers, Lua, MultiValue, Value,
    VmState,
};

use super::custom_functions::load_custom_functions;
use super::ops::LuaOp;
use super::runtime::{
    format_lua_error, lua_value_to_string, prepare_code, LuaEvalResult, DEFAULT_TIMEOUT,
    INSTRUCTION_HOOK_INTERVAL, INSTRUCTION_LIMIT,
};
use super::sheet_api::{register_sheet_global_with_selection, SheetSnapshot, MAX_OUTPUT_LINES};

// =============================================================================
// Constants
// =============================================================================

/// Source name for the console chunk (set via `Chunk::set_name`).
pub const CONSOLE_SOURCE: &str = "@console";

/// Maximum locals collected per frame.
const MAX_LOCALS: usize = 200;

/// Maximum upvalues collected per frame.
const MAX_UPVALUES: usize = 200;

/// Maximum display length for a variable value string.
const MAX_VALUE_LEN: usize = 200;

/// Maximum children returned from table expansion.
const MAX_EXPAND_CHILDREN: usize = 50;

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
    AddBreakpoint { source: String, line: usize },
    RemoveBreakpoint { source: String, line: usize },
    /// Request variables for a specific call stack frame.
    RequestFrameVars { frame_index: usize },
    /// Expand a variable (table) at the given path within a frame.
    ExpandVariable {
        frame_index: usize,
        path: Vec<VarPathSegment>,
    },
}

/// A segment in a variable expansion path.
#[derive(Debug, Clone)]
pub enum VarPathSegment {
    /// Local variable at position N in the locals Vec (0-indexed).
    Local(usize),
    /// Upvalue at position N in the upvalues Vec (0-indexed).
    Upvalue(usize),
    /// Table key (string).
    KeyString(String),
    /// Table key (integer).
    KeyInt(i64),
    /// Table key (other, displayed as string).
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
    /// Script execution paused at a breakpoint or step.
    /// Includes locals/upvalues for the top frame.
    Paused(DebugSnapshot),
    /// A chunk of print() output.
    OutputChunk(String),
    /// Script finished (success or script-level error).
    Completed(LuaEvalResult),
    /// Catastrophic failure (VM creation failed, sandbox failed, etc.).
    Error(String),
    /// Variables for a requested frame (response to RequestFrameVars).
    FrameVars {
        frame_index: usize,
        locals: Vec<Variable>,
        upvalues: Vec<Variable>,
    },
    /// Expanded table children (response to ExpandVariable).
    VariableExpanded {
        frame_index: usize,
        path: Vec<VarPathSegment>,
        children: Vec<Variable>,
    },
}

// =============================================================================
// Snapshot types
// =============================================================================

/// Snapshot of debugger state at a pause point.
#[derive(Debug, Clone)]
pub struct DebugSnapshot {
    pub line: usize,
    pub call_stack: Vec<StackFrame>,
    /// Locals for the top frame (frame_index 0).
    pub locals: Vec<Variable>,
    /// Upvalues for the top frame (frame_index 0).
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
    /// The Lua stack level for `inspect_stack` / variable collection.
    pub lua_stack_level: usize,
}

/// A variable visible at a pause point.
#[derive(Debug, Clone)]
pub struct Variable {
    pub name: String,
    pub value: String,
    pub expandable: bool,
}

/// Why execution paused.
#[derive(Debug, Clone, PartialEq, Eq)]
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
    pub cancel: Arc<AtomicBool>,
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
    /// Initial breakpoints as (source_name, line) pairs.
    pub breakpoints: HashSet<(String, usize)>,
}

// =============================================================================
// Private: HookIntrospection
// =============================================================================

/// Raw debug library functions captured before sandboxing.
///
/// Called from Rust via `Function::call()` during hook pauses.
/// The caller must add the appropriate stack-level offset.
struct HookIntrospection {
    /// Raw `debug.getlocal(level, index) -> name, value`
    getlocal: LuaFunction,
    /// Raw `debug.getinfo(level, what) -> info_table`
    getinfo: LuaFunction,
    /// Raw `debug.getupvalue(func, index) -> name, value`
    getupvalue: LuaFunction,
}

// =============================================================================
// Private: OutputBuffer
// =============================================================================

/// Coalesces print() output into time/size-bounded chunks for streaming.
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
// Private: TimeoutAccounting
// =============================================================================

/// Tracks execution time excluding paused intervals.
struct TimeoutAccounting {
    started_at: Instant,
    total_paused: Duration,
    paused_at: Option<Instant>,
}

impl TimeoutAccounting {
    fn new() -> Self {
        Self {
            started_at: Instant::now(),
            total_paused: Duration::ZERO,
            paused_at: None,
        }
    }

    fn elapsed_execution_time(&self) -> Duration {
        self.started_at.elapsed() - self.total_paused
    }

    fn pause(&mut self) {
        self.paused_at = Some(Instant::now());
    }

    fn resume(&mut self) {
        if let Some(t) = self.paused_at.take() {
            self.total_paused += t.elapsed();
        }
    }
}

// =============================================================================
// Private: DebugMode
// =============================================================================

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DebugMode {
    Running,
    StepIn,
    StepOver,
    StepOut,
}

// =============================================================================
// Private: HookState
// =============================================================================

/// All mutable state accessed by the debug hook.
///
/// Wrapped in `Rc<RefCell<>>` because `set_hook` takes `Fn`, not `FnMut`.
/// This is safe: everything runs on the single debug thread.
struct HookState {
    session_id: SessionId,
    cancel: Arc<AtomicBool>,
    budget: i64,
    timeout: TimeoutAccounting,
    mode: DebugMode,
    breakpoints: HashSet<(String, usize)>,
    stack_depth: usize,
    step_target_depth: usize,
    cmd_rx: mpsc::Receiver<DebugCommand>,
    event_tx: mpsc::Sender<DebugEvent>,
    output_buffer: Rc<RefCell<OutputBuffer>>,
    introspection: Option<HookIntrospection>,
    /// Lua getlocal indices for the last sent locals (maps Vec position → lua 1-index).
    last_locals_lua_indices: Vec<usize>,
    /// Lua getupvalue indices for the last sent upvalues.
    last_upvalues_lua_indices: Vec<usize>,
    /// Frame level for the last sent variables.
    last_vars_frame_level: usize,
    // Flags set during execution
    timed_out: bool,
    instruction_limit_exceeded: bool,
}

impl HookState {
    fn on_hook(&mut self, lua: &Lua, debug: &mlua::Debug) -> Result<VmState, mlua::Error> {
        // Fast path: cancel check (runs on every event type)
        if self.cancel.load(Ordering::Relaxed) {
            return Err(mlua::Error::RuntimeError("execution stopped".into()));
        }

        match debug.event() {
            HookEvent::Count => self.on_count(),
            HookEvent::Call => {
                self.stack_depth += 1;
                Ok(VmState::Continue)
            }
            HookEvent::TailCall => Ok(VmState::Continue),
            HookEvent::Ret => {
                self.stack_depth = self.stack_depth.saturating_sub(1);
                Ok(VmState::Continue)
            }
            HookEvent::Line => self.on_line(lua, debug),
            _ => Ok(VmState::Continue),
        }
    }

    fn on_count(&mut self) -> Result<VmState, mlua::Error> {
        // Budget
        self.budget -= INSTRUCTION_HOOK_INTERVAL as i64;
        if self.budget <= 0 {
            self.instruction_limit_exceeded = true;
            return Err(mlua::Error::RuntimeError(format!(
                "instruction limit exceeded ({} instructions)",
                INSTRUCTION_LIMIT
            )));
        }

        // Timeout (execution time only, excludes paused time)
        if self.timeout.elapsed_execution_time() > DEFAULT_TIMEOUT {
            self.timed_out = true;
            return Err(mlua::Error::RuntimeError(format!(
                "execution timeout ({}s limit)",
                DEFAULT_TIMEOUT.as_secs()
            )));
        }

        // Drain pending commands (non-blocking)
        self.drain_commands()
    }

    fn on_line(&mut self, lua: &Lua, debug: &mlua::Debug) -> Result<VmState, mlua::Error> {
        match self.mode {
            DebugMode::Running => {
                if self.breakpoints.is_empty() {
                    return Ok(VmState::Continue);
                }
                let line = debug.curr_line() as usize;
                let source = get_source_name(debug);
                if self.breakpoints.contains(&(source, line)) {
                    self.pause_and_wait(lua, debug, PauseReason::Breakpoint)
                } else {
                    Ok(VmState::Continue)
                }
            }
            DebugMode::StepIn => self.pause_and_wait(lua, debug, PauseReason::StepIn),
            DebugMode::StepOver => {
                if self.stack_depth <= self.step_target_depth {
                    self.pause_and_wait(lua, debug, PauseReason::StepOver)
                } else {
                    Ok(VmState::Continue)
                }
            }
            DebugMode::StepOut => {
                if self.stack_depth < self.step_target_depth {
                    self.pause_and_wait(lua, debug, PauseReason::StepOut)
                } else {
                    Ok(VmState::Continue)
                }
            }
        }
    }

    fn pause_and_wait(
        &mut self,
        lua: &Lua,
        debug: &mlua::Debug,
        reason: PauseReason,
    ) -> Result<VmState, mlua::Error> {
        let line = debug.curr_line().max(0) as usize;
        let call_stack = collect_call_stack(lua);

        // Collect locals/upvalues for top frame (frame 0)
        let (locals, local_indices) = collect_locals(&self.introspection, 0);
        let (upvalues, upvalue_indices) = collect_upvalues(&self.introspection, 0);

        // Store index maps for ExpandVariable
        self.last_locals_lua_indices = local_indices;
        self.last_upvalues_lua_indices = upvalue_indices;
        self.last_vars_frame_level = 0;

        // Flush output before Paused event (ordering guarantee)
        self.output_buffer.borrow_mut().flush();

        let snapshot = DebugSnapshot {
            line,
            call_stack,
            locals,
            upvalues,
            reason,
        };

        let _ = self.event_tx.send(DebugEvent {
            session_id: self.session_id,
            payload: DebugEventPayload::Paused(snapshot),
        });

        // Pause timeout accounting
        self.timeout.pause();
        let result = self.wait_for_resume(lua);
        self.timeout.resume();

        result
    }

    fn wait_for_resume(&mut self, lua: &Lua) -> Result<VmState, mlua::Error> {
        loop {
            match self.cmd_rx.recv() {
                Ok(cmd) => {
                    if cmd.session_id != self.session_id {
                        continue;
                    }
                    match cmd.action {
                        DebugAction::Continue => {
                            self.mode = DebugMode::Running;
                            return Ok(VmState::Continue);
                        }
                        DebugAction::StepIn => {
                            self.mode = DebugMode::StepIn;
                            return Ok(VmState::Continue);
                        }
                        DebugAction::StepOver => {
                            self.mode = DebugMode::StepOver;
                            self.step_target_depth = self.stack_depth;
                            return Ok(VmState::Continue);
                        }
                        DebugAction::StepOut => {
                            self.mode = DebugMode::StepOut;
                            self.step_target_depth = self.stack_depth;
                            return Ok(VmState::Continue);
                        }
                        DebugAction::Stop => {
                            self.cancel.store(true, Ordering::Relaxed);
                            return Err(mlua::Error::RuntimeError(
                                "execution stopped".into(),
                            ));
                        }
                        DebugAction::AddBreakpoint { source, line } => {
                            self.breakpoints.insert((source, line));
                        }
                        DebugAction::RemoveBreakpoint { source, line } => {
                            self.breakpoints.remove(&(source, line));
                        }
                        DebugAction::RequestFrameVars { frame_index } => {
                            self.handle_request_frame_vars(frame_index);
                        }
                        DebugAction::ExpandVariable { frame_index, path } => {
                            self.handle_expand_variable(lua, frame_index, path);
                        }
                    }
                }
                Err(_) => {
                    self.cancel.store(true, Ordering::Relaxed);
                    return Err(mlua::Error::RuntimeError(
                        "debug session closed".into(),
                    ));
                }
            }
        }
    }

    fn handle_request_frame_vars(&mut self, frame_index: usize) {
        let (locals, local_indices) = collect_locals(&self.introspection, frame_index);
        let (upvalues, upvalue_indices) = collect_upvalues(&self.introspection, frame_index);

        self.last_locals_lua_indices = local_indices;
        self.last_upvalues_lua_indices = upvalue_indices;
        self.last_vars_frame_level = frame_index;

        let _ = self.event_tx.send(DebugEvent {
            session_id: self.session_id,
            payload: DebugEventPayload::FrameVars {
                frame_index,
                locals,
                upvalues,
            },
        });
    }

    fn handle_expand_variable(
        &mut self,
        lua: &Lua,
        frame_index: usize,
        path: Vec<VarPathSegment>,
    ) {
        // If the frame doesn't match last collected, re-collect index maps
        if frame_index != self.last_vars_frame_level {
            let (_, local_indices) = collect_locals(&self.introspection, frame_index);
            let (_, upvalue_indices) = collect_upvalues(&self.introspection, frame_index);
            self.last_locals_lua_indices = local_indices;
            self.last_upvalues_lua_indices = upvalue_indices;
            self.last_vars_frame_level = frame_index;
        }

        let children = navigate_and_expand(
            &self.introspection,
            lua,
            frame_index,
            &path,
            &self.last_locals_lua_indices,
            &self.last_upvalues_lua_indices,
        );

        let _ = self.event_tx.send(DebugEvent {
            session_id: self.session_id,
            payload: DebugEventPayload::VariableExpanded {
                frame_index,
                path,
                children,
            },
        });
    }

    fn drain_commands(&mut self) -> Result<VmState, mlua::Error> {
        loop {
            match self.cmd_rx.try_recv() {
                Ok(cmd) => {
                    if cmd.session_id != self.session_id {
                        continue;
                    }
                    match cmd.action {
                        DebugAction::Stop => {
                            self.cancel.store(true, Ordering::Relaxed);
                            return Err(mlua::Error::RuntimeError(
                                "execution stopped".into(),
                            ));
                        }
                        DebugAction::AddBreakpoint { source, line } => {
                            self.breakpoints.insert((source, line));
                        }
                        DebugAction::RemoveBreakpoint { source, line } => {
                            self.breakpoints.remove(&(source, line));
                        }
                        _ => {} // Continue/Step/RequestFrameVars/Expand not valid while running
                    }
                }
                Err(mpsc::TryRecvError::Empty) => return Ok(VmState::Continue),
                Err(mpsc::TryRecvError::Disconnected) => {
                    self.cancel.store(true, Ordering::Relaxed);
                    return Err(mlua::Error::RuntimeError(
                        "debug session closed".into(),
                    ));
                }
            }
        }
    }
}

// =============================================================================
// Public API
// =============================================================================

/// Spawn a new debug session on a background thread.
pub fn spawn_debug_session(config: DebugConfig) -> DebugSession {
    let session_id = NEXT_SESSION_ID.fetch_add(1, Ordering::Relaxed);

    let (cmd_tx, cmd_rx) = mpsc::channel::<DebugCommand>();
    let (event_tx, event_rx) = mpsc::channel::<DebugEvent>();
    let cancel = Arc::new(AtomicBool::new(false));
    let cancel_clone = cancel.clone();

    std::thread::Builder::new()
        .name(format!("debug-session-{}", session_id))
        .spawn(move || {
            debug_thread_main(session_id, config, cmd_rx, event_tx, cancel_clone);
        })
        .expect("Failed to spawn debug thread");

    DebugSession {
        id: session_id,
        state: DebugSessionState::Starting,
        cmd_tx,
        event_rx,
        cancel,
    }
}

// =============================================================================
// Thread entry point
// =============================================================================

fn debug_thread_main(
    session_id: SessionId,
    config: DebugConfig,
    cmd_rx: mpsc::Receiver<DebugCommand>,
    event_tx: mpsc::Sender<DebugEvent>,
    cancel: Arc<AtomicBool>,
) {
    // 1. Create Lua VM with debug library (needed for variable inspection).
    //    SAFETY: We sandbox the VM immediately after capturing debug functions,
    //    removing the debug global and other dangerous modules.
    let lua = unsafe { Lua::unsafe_new() };

    // 2. Capture raw debug functions (before sandbox removes debug global)
    let introspection = create_hook_introspection(&lua);

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

            {
                let mut lines = lines.borrow_mut();
                if lines.len() < MAX_OUTPUT_LINES {
                    lines.push(line.clone());
                } else if !*truncated.borrow() {
                    *truncated.borrow_mut() = true;
                }
            }

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
                payload: DebugEventPayload::Error(format!("Failed to register print: {}", e)),
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

    // 6. Load custom functions (best-effort)
    let _ = load_custom_functions(&lua);

    // 7. Empty code fast path
    let trimmed = config.code.trim();
    if trimmed.is_empty() {
        let _ = event_tx.send(DebugEvent {
            session_id,
            payload: DebugEventPayload::Completed(LuaEvalResult {
                output: vec![],
                returned: None,
                error: None,
                ops: Vec::new(),
                mutations: 0,
                output_truncated: false,
                instruction_limit_exceeded: false,
                cancelled: false,
                timed_out: false,
                cells_read: 0,
            }),
        });
        return;
    }

    // 8. Prepare code (expression wrapping)
    let (code, is_expression) = prepare_code(&lua, trimmed);

    // 9. Create HookState
    let hook_state = Rc::new(RefCell::new(HookState {
        session_id,
        cancel: cancel.clone(),
        budget: INSTRUCTION_LIMIT,
        timeout: TimeoutAccounting::new(),
        mode: DebugMode::Running,
        breakpoints: config.breakpoints,
        stack_depth: 0,
        step_target_depth: 0,
        cmd_rx,
        event_tx: event_tx.clone(),
        output_buffer: output_buffer.clone(),
        introspection,
        last_locals_lua_indices: Vec::new(),
        last_upvalues_lua_indices: Vec::new(),
        last_vars_frame_level: 0,
        timed_out: false,
        instruction_limit_exceeded: false,
    }));

    // 10. Install hook (every line + calls + returns + nth instruction)
    {
        let hook_ref = hook_state.clone();
        lua.set_hook(
            HookTriggers::new()
                .every_line()
                .on_calls()
                .on_returns()
                .every_nth_instruction(INSTRUCTION_HOOK_INTERVAL),
            move |lua, debug| hook_ref.borrow_mut().on_hook(lua, &debug),
        );
    }

    // 11. Execute with source name set for breakpoint matching
    let result = lua
        .load(&code)
        .set_name(CONSOLE_SOURCE)
        .eval::<MultiValue>();

    // 12. Remove hook, flush OutputBuffer
    lua.remove_hook();
    output_buffer.borrow_mut().flush();

    // 13. Collect output
    let mut output = output_lines.borrow().clone();
    let truncated = *output_truncated.borrow();
    if truncated {
        output.push(format!(
            "... output truncated ({} line limit)",
            MAX_OUTPUT_LINES
        ));
    }

    // 14. Collect ops from sink
    let (ops, mutations) = {
        let mut borrowed = sink.borrow_mut();
        let mutations = borrowed.mutations();
        let ops = borrowed.take_ops();
        (ops, mutations)
    };

    // 15. Read flags from hook state
    let hs = hook_state.borrow();
    let timed_out = hs.timed_out;
    let instruction_limit_exceeded = hs.instruction_limit_exceeded;
    drop(hs);
    let cancelled = cancel.load(Ordering::Relaxed);

    // 16. Build LuaEvalResult, send Completed
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
                cells_read: 0,
            }
        }
        Err(e) => {
            if cancelled {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some("execution stopped".into()),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: false,
                    cancelled: true,
                    timed_out: false,
                    cells_read: 0,
                }
            } else if timed_out {
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
                    cells_read: 0,
                }
            } else if instruction_limit_exceeded {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some(format_lua_error(&e)),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: true,
                    cancelled: false,
                    timed_out: false,
                    cells_read: 0,
                }
            } else {
                LuaEvalResult {
                    output,
                    returned: None,
                    error: Some(format_lua_error(&e)),
                    ops,
                    mutations,
                    output_truncated: truncated,
                    instruction_limit_exceeded: false,
                    cancelled: false,
                    timed_out: false,
                    cells_read: 0,
                }
            }
        }
    };

    let _ = event_tx.send(DebugEvent {
        session_id,
        payload: DebugEventPayload::Completed(eval_result),
    });
}

// =============================================================================
// Helpers: Introspection
// =============================================================================

/// Capture raw `debug.getlocal`, `debug.getinfo`, `debug.getupvalue` before
/// sandbox removes the `debug` global.
fn create_hook_introspection(lua: &Lua) -> Option<HookIntrospection> {
    let debug_table: mlua::Table = lua.globals().get("debug").ok()?;
    let getlocal: LuaFunction = debug_table.get("getlocal").ok()?;
    let getinfo: LuaFunction = debug_table.get("getinfo").ok()?;
    let getupvalue: LuaFunction = debug_table.get("getupvalue").ok()?;
    Some(HookIntrospection {
        getlocal,
        getinfo,
        getupvalue,
    })
}

// =============================================================================
// Helpers: Variable Collection
// =============================================================================

/// Stack-level offset when calling `debug.getlocal`/`debug.getinfo` from Rust
/// via mlua's `Function::call()` during a hook. The C function's own CallInfo
/// is level 0, so user frame N is at level N + RUST_CALL_OFFSET.
const RUST_CALL_OFFSET: usize = 1;

/// Collect local variables for a frame. Returns (variables, lua_getlocal_indices).
fn collect_locals(
    introspection: &Option<HookIntrospection>,
    frame_level: usize,
) -> (Vec<Variable>, Vec<usize>) {
    let Some(intro) = introspection else {
        return (Vec::new(), Vec::new());
    };

    let lua_level = frame_level + RUST_CALL_OFFSET;
    let mut vars = Vec::new();
    let mut indices = Vec::new();

    for lua_idx in 1..=(MAX_LOCALS + 100) {
        // Over-scan to account for skipped internal vars
        if vars.len() >= MAX_LOCALS {
            break;
        }

        let result: MultiValue = match intro.getlocal.call((lua_level, lua_idx)) {
            Ok(r) => r,
            Err(_) => break,
        };

        let result_vec = result.into_vec();
        if result_vec.len() < 2 {
            break; // No more locals
        }

        let name = match &result_vec[0] {
            Value::String(s) => match s.to_str() {
                Ok(s) => s.to_string(),
                Err(_) => break,
            },
            _ => break,
        };

        // Skip internal variables (temporaries named "(for ...)" etc.)
        if name.starts_with('(') {
            continue;
        }

        let (display, expandable) = format_value_for_display(&result_vec[1]);
        vars.push(Variable {
            name,
            value: display,
            expandable,
        });
        indices.push(lua_idx);
    }

    (vars, indices)
}

/// Collect upvalues for a frame. Returns (variables, lua_getupvalue_indices).
fn collect_upvalues(
    introspection: &Option<HookIntrospection>,
    frame_level: usize,
) -> (Vec<Variable>, Vec<usize>) {
    let Some(intro) = introspection else {
        return (Vec::new(), Vec::new());
    };

    // Get the function at this frame via debug.getinfo
    let lua_level = frame_level + RUST_CALL_OFFSET;
    let func = match get_frame_function(intro, lua_level) {
        Some(f) => f,
        None => return (Vec::new(), Vec::new()),
    };

    let mut vars = Vec::new();
    let mut indices = Vec::new();

    for lua_idx in 1..=MAX_UPVALUES {
        let result: MultiValue = match intro.getupvalue.call((func.clone(), lua_idx)) {
            Ok(r) => r,
            Err(_) => break,
        };

        let result_vec = result.into_vec();
        if result_vec.len() < 2 {
            break; // No more upvalues
        }

        let name = match &result_vec[0] {
            Value::String(s) => match s.to_str() {
                Ok(s) => s.to_string(),
                Err(_) => break,
            },
            _ => break,
        };

        let (display, expandable) = format_value_for_display(&result_vec[1]);
        vars.push(Variable {
            name,
            value: display,
            expandable,
        });
        indices.push(lua_idx);
    }

    (vars, indices)
}

/// Get the function object at a given Lua stack level via `debug.getinfo`.
fn get_frame_function(intro: &HookIntrospection, lua_level: usize) -> Option<LuaFunction> {
    let info: mlua::Table = intro.getinfo.call((lua_level, "f")).ok()?;
    info.get::<LuaFunction>("func").ok()
}

/// Navigate a path to a table and return its children.
fn navigate_and_expand(
    introspection: &Option<HookIntrospection>,
    lua: &Lua,
    frame_level: usize,
    path: &[VarPathSegment],
    local_indices: &[usize],
    upvalue_indices: &[usize],
) -> Vec<Variable> {
    let Some(intro) = introspection else {
        return Vec::new();
    };

    if path.is_empty() {
        return Vec::new();
    }

    let lua_level = frame_level + RUST_CALL_OFFSET;

    // Get the root value
    let root_value = match &path[0] {
        VarPathSegment::Local(pos) => {
            let Some(&lua_idx) = local_indices.get(*pos) else {
                return Vec::new();
            };
            get_local_value(&intro.getlocal, lua_level, lua_idx)
        }
        VarPathSegment::Upvalue(pos) => {
            let Some(&lua_idx) = upvalue_indices.get(*pos) else {
                return Vec::new();
            };
            let Some(func) = get_frame_function(intro, lua_level) else {
                return Vec::new();
            };
            get_upvalue_value(&intro.getupvalue, &func, lua_idx)
        }
        _ => return Vec::new(),
    };

    let Some(mut current) = root_value else {
        return Vec::new();
    };

    // Navigate remaining path segments through tables
    for segment in &path[1..] {
        match (&current, segment) {
            (Value::Table(t), VarPathSegment::KeyString(s)) => {
                current = t.get::<Value>(s.as_str()).unwrap_or(Value::Nil);
            }
            (Value::Table(t), VarPathSegment::KeyInt(i)) => {
                current = t.get::<Value>(*i).unwrap_or(Value::Nil);
            }
            _ => return Vec::new(),
        }
    }

    // Expand the final value
    match current {
        Value::Table(table) => expand_table(lua, &table),
        _ => Vec::new(),
    }
}

/// Get a local variable's value via `debug.getlocal(lua_level, lua_idx)`.
fn get_local_value(getlocal: &LuaFunction, lua_level: usize, lua_idx: usize) -> Option<Value> {
    let result: MultiValue = getlocal.call((lua_level, lua_idx)).ok()?;
    let mut result_vec = result.into_vec();
    if result_vec.len() < 2 {
        return None;
    }
    Some(result_vec.remove(1)) // index 1 = value (index 0 = name)
}

/// Get an upvalue's value via `debug.getupvalue(func, lua_idx)`.
fn get_upvalue_value(getupvalue: &LuaFunction, func: &LuaFunction, lua_idx: usize) -> Option<Value> {
    let result: MultiValue = getupvalue.call((func.clone(), lua_idx)).ok()?;
    let mut result_vec = result.into_vec();
    if result_vec.len() < 2 {
        return None;
    }
    Some(result_vec.remove(1))
}

/// Expand a table into its children (capped at MAX_EXPAND_CHILDREN).
fn expand_table(_lua: &Lua, table: &mlua::Table) -> Vec<Variable> {
    let mut children = Vec::new();
    let Ok(pairs) = table.pairs::<Value, Value>().collect::<Result<Vec<_>, _>>() else {
        return children;
    };
    for (key, value) in pairs {
        if children.len() >= MAX_EXPAND_CHILDREN {
            break;
        }
        let name = format_key_for_display(&key);
        let (display, expandable) = format_value_for_display(&value);
        children.push(Variable {
            name,
            value: display,
            expandable,
        });
    }
    children
}

// =============================================================================
// Helpers: Display Formatting
// =============================================================================

/// Format a Lua value for display in the variables panel.
/// Returns (display_string, is_expandable).
fn format_value_for_display(value: &Value) -> (String, bool) {
    let expandable = matches!(value, Value::Table(_));
    let display = match value {
        Value::Nil => "nil".to_string(),
        Value::Boolean(b) => b.to_string(),
        Value::Integer(i) => i.to_string(),
        Value::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            }
        }
        Value::String(s) => {
            match s.to_str() {
                Ok(s) => {
                    if s.len() > MAX_VALUE_LEN {
                        format!("\"{}...\"", &s[..MAX_VALUE_LEN])
                    } else {
                        format!("\"{}\"", s)
                    }
                }
                Err(_) => "\"<invalid utf8>\"".to_string(),
            }
        }
        Value::Table(_) => "{...}".to_string(),
        Value::Function(_) => "function".to_string(),
        Value::Thread(_) => "thread".to_string(),
        Value::UserData(_) => "userdata".to_string(),
        Value::LightUserData(_) => "lightuserdata".to_string(),
        Value::Error(e) => format!("error: {}", e),
        _ => "<unknown>".to_string(),
    };
    (display, expandable)
}

/// Format a table key for display.
fn format_key_for_display(key: &Value) -> String {
    match key {
        Value::String(s) => match s.to_str() {
            Ok(s) => s.to_string(),
            Err(_) => "?".to_string(),
        },
        Value::Integer(i) => format!("[{}]", i),
        Value::Number(n) => format!("[{}]", n),
        Value::Boolean(b) => format!("[{}]", b),
        _ => "[?]".to_string(),
    }
}

// =============================================================================
// Helpers: General
// =============================================================================

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

/// Extract the source name from a debug info structure.
fn get_source_name(debug: &mlua::Debug) -> String {
    let src = debug.source();
    match &src.source {
        Some(s) => s.to_string(),
        None => "?".to_string(),
    }
}

/// Walk the Lua call stack and collect frames (capped at 64).
fn collect_call_stack(lua: &Lua) -> Vec<StackFrame> {
    let mut frames = Vec::new();
    for level in 0..64usize {
        match lua.inspect_stack(level) {
            Some(debug) => {
                let names = debug.names();
                let source = debug.source();
                frames.push(StackFrame {
                    function_name: names.name.as_ref().map(|s| s.to_string()),
                    source: source.source.as_ref().map(|s| s.to_string()),
                    line: debug.curr_line().max(0) as usize,
                    is_tail_call: debug.is_tail_call(),
                    lua_stack_level: level,
                });
            }
            None => break,
        }
    }
    frames
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use std::time::Duration;

    use super::super::ops::LuaCellValue;

    fn mock_snapshot() -> SheetSnapshot {
        SheetSnapshot::from_mock(HashMap::new(), HashMap::new(), 100, 26)
    }

    fn simple_config(code: &str) -> DebugConfig {
        DebugConfig {
            code: code.to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
            breakpoints: HashSet::new(),
        }
    }

    fn config_with_breakpoints(code: &str, bps: Vec<(&str, usize)>) -> DebugConfig {
        DebugConfig {
            code: code.to_string(),
            snapshot: mock_snapshot(),
            selection: (0, 0, 0, 0),
            breakpoints: bps
                .into_iter()
                .map(|(s, l)| (s.to_string(), l))
                .collect(),
        }
    }

    fn send_cmd(session: &DebugSession, action: DebugAction) {
        session
            .cmd_tx
            .send(DebugCommand {
                session_id: session.id,
                action,
            })
            .unwrap();
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

    fn recv_paused(session: &DebugSession) -> DebugSnapshot {
        let timeout = Duration::from_secs(5);
        loop {
            match session.event_rx.recv_timeout(timeout) {
                Ok(event) => {
                    assert_eq!(event.session_id, session.id);
                    match event.payload {
                        DebugEventPayload::Paused(snap) => return snap,
                        DebugEventPayload::OutputChunk(_) => continue,
                        DebugEventPayload::Completed(r) => {
                            panic!("Got Completed instead of Paused: {:?}", r)
                        }
                        DebugEventPayload::Error(e) => panic!("Unexpected error: {}", e),
                        _ => continue,
                    }
                }
                Err(e) => panic!("Timed out waiting for Paused event: {}", e),
            }
        }
    }

    fn recv_frame_vars(session: &DebugSession) -> (usize, Vec<Variable>, Vec<Variable>) {
        let timeout = Duration::from_secs(5);
        loop {
            match session.event_rx.recv_timeout(timeout) {
                Ok(event) => {
                    assert_eq!(event.session_id, session.id);
                    match event.payload {
                        DebugEventPayload::FrameVars {
                            frame_index,
                            locals,
                            upvalues,
                        } => return (frame_index, locals, upvalues),
                        DebugEventPayload::OutputChunk(_) => continue,
                        other => panic!("Expected FrameVars, got: {:?}", other),
                    }
                }
                Err(e) => panic!("Timed out waiting for FrameVars: {}", e),
            }
        }
    }

    fn recv_expanded(session: &DebugSession) -> Vec<Variable> {
        let timeout = Duration::from_secs(5);
        loop {
            match session.event_rx.recv_timeout(timeout) {
                Ok(event) => {
                    assert_eq!(event.session_id, session.id);
                    match event.payload {
                        DebugEventPayload::VariableExpanded { children, .. } => return children,
                        DebugEventPayload::OutputChunk(_) => continue,
                        other => panic!("Expected VariableExpanded, got: {:?}", other),
                    }
                }
                Err(e) => panic!("Timed out waiting for VariableExpanded: {}", e),
            }
        }
    }

    // =====================================================================
    // Phase 1 tests
    // =====================================================================

    #[test]
    fn test_basic_script_completes() {
        let session = spawn_debug_session(simple_config("return 1 + 1"));
        let result = recv_completed(&session);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("2".to_string()));
    }

    #[test]
    fn test_print_output() {
        let session = spawn_debug_session(simple_config("print('hello')"));
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
        let session = spawn_debug_session(simple_config("sheet:set('A1', 42)"));
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
        let session = spawn_debug_session(simple_config("if then"));
        let result = recv_completed(&session);
        assert!(result.error.is_some(), "Expected syntax error");
    }

    #[test]
    fn test_session_id_increments() {
        let s1 = spawn_debug_session(simple_config("return 1"));
        let s2 = spawn_debug_session(simple_config("return 2"));
        assert_ne!(s1.id, s2.id);
        assert!(s2.id > s1.id);
        recv_completed(&s1);
        recv_completed(&s2);
    }

    #[test]
    fn test_instruction_limit() {
        let session = spawn_debug_session(simple_config("while true do end"));
        let result = recv_completed(&session);
        assert!(result.error.is_some(), "Expected error for infinite loop");
        assert!(
            result.instruction_limit_exceeded,
            "Expected instruction_limit_exceeded flag"
        );
    }

    #[test]
    fn test_expression_wrapping() {
        let s1 = spawn_debug_session(simple_config("1 + 1"));
        let r1 = recv_completed(&s1);
        assert!(r1.error.is_none(), "Error: {:?}", r1.error);
        assert_eq!(r1.returned, Some("2".to_string()));

        let s2 = spawn_debug_session(simple_config("local x = 1"));
        let r2 = recv_completed(&s2);
        assert!(r2.error.is_none(), "Error: {:?}", r2.error);
        assert!(
            r2.returned.is_none() || r2.returned == Some("nil".to_string()),
            "Statement should not return a value, got: {:?}",
            r2.returned
        );
    }

    // =====================================================================
    // Phase 2 tests: Breakpoints + stepping + cleanup
    // =====================================================================

    #[test]
    fn test_breakpoint_pauses() {
        let code = "local x = 1\nlocal y = 2\nreturn x + y";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 2)],
        ));
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 2);
        assert_eq!(snap.reason, PauseReason::Breakpoint);

        send_cmd(&session, DebugAction::Continue);
        let result = recv_completed(&session);
        assert!(result.error.is_none(), "Error: {:?}", result.error);
        assert_eq!(result.returned, Some("3".to_string()));
    }

    #[test]
    fn test_step_in() {
        let code = "local x = 1\nlocal y = 2\nreturn x + y";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 1)],
        ));
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 1);

        send_cmd(&session, DebugAction::StepIn);
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 2);
        assert_eq!(snap.reason, PauseReason::StepIn);

        send_cmd(&session, DebugAction::StepIn);
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 3);

        send_cmd(&session, DebugAction::Continue);
        let result = recv_completed(&session);
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("3".to_string()));
    }

    #[test]
    fn test_step_over() {
        let code = "\
local function add(a, b)
  return a + b
end
local x = add(1, 2)
local y = x + 1
return y";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 4)],
        ));
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 4);

        send_cmd(&session, DebugAction::StepOver);
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 5);
        assert_eq!(snap.reason, PauseReason::StepOver);

        send_cmd(&session, DebugAction::Continue);
        let result = recv_completed(&session);
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("4".to_string()));
    }

    #[test]
    fn test_step_out() {
        let code = "\
local function add(a, b)
  local sum = a + b
  return sum
end
local x = add(1, 2)
return x";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 2)],
        ));
        let snap = recv_paused(&session);
        assert_eq!(snap.line, 2);

        send_cmd(&session, DebugAction::StepOut);
        let snap = recv_paused(&session);
        assert_eq!(snap.reason, PauseReason::StepOut);
        assert_eq!(snap.line, 6);

        send_cmd(&session, DebugAction::Continue);
        let result = recv_completed(&session);
        assert!(result.error.is_none());
        assert_eq!(result.returned, Some("3".to_string()));
    }

    #[test]
    fn test_stop_while_paused() {
        let code = "local x = 1\nlocal y = 2\nreturn x + y";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 1)],
        ));
        recv_paused(&session);

        send_cmd(&session, DebugAction::Stop);
        let result = recv_completed(&session);
        assert!(result.cancelled);

        assert!(session.event_rx.recv_timeout(Duration::from_millis(100)).is_err());
    }

    #[test]
    fn test_stop_while_running() {
        let session = spawn_debug_session(simple_config(
            "local s = 0\nfor i = 1, 1000000000 do s = s + i end\nreturn s",
        ));
        std::thread::sleep(Duration::from_millis(50));
        session.cancel.store(true, Ordering::Relaxed);

        let result = recv_completed(&session);
        assert!(result.cancelled);
    }

    #[test]
    fn test_restart_after_stop() {
        let session_a = spawn_debug_session(config_with_breakpoints(
            "local x = 1\nreturn x",
            vec![(CONSOLE_SOURCE, 1)],
        ));
        recv_paused(&session_a);
        send_cmd(&session_a, DebugAction::Stop);
        let result_a = recv_completed(&session_a);
        assert!(result_a.cancelled);

        let session_b = spawn_debug_session(simple_config("return 42"));
        assert_ne!(session_a.id, session_b.id);
        let result_b = recv_completed(&session_b);
        assert!(result_b.error.is_none());
        assert_eq!(result_b.returned, Some("42".to_string()));

        assert!(session_a.event_rx.recv_timeout(Duration::from_millis(100)).is_err());
    }

    // =====================================================================
    // Phase 3 tests: Variables
    // =====================================================================

    #[test]
    fn test_locals_collected_on_pause() {
        let code = "local x = 42\nlocal y = 'hello'\nreturn x";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 3)],
        ));

        let snap = recv_paused(&session);
        assert_eq!(snap.line, 3);

        // By line 3, both x and y should be in scope
        let names: Vec<&str> = snap.locals.iter().map(|v| v.name.as_str()).collect();
        assert!(names.contains(&"x"), "Expected local 'x', got: {:?}", names);
        assert!(names.contains(&"y"), "Expected local 'y', got: {:?}", names);

        // Check values
        let x = snap.locals.iter().find(|v| v.name == "x").unwrap();
        assert_eq!(x.value, "42");
        assert!(!x.expandable);

        let y = snap.locals.iter().find(|v| v.name == "y").unwrap();
        assert_eq!(y.value, "\"hello\"");
        assert!(!y.expandable);

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_tables_flagged_expandable() {
        let code = "local t = {1, 2, 3}\nlocal n = 42\nreturn n";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 3)],
        ));

        let snap = recv_paused(&session);
        let t = snap.locals.iter().find(|v| v.name == "t").unwrap();
        assert!(t.expandable, "Table should be expandable");
        assert_eq!(t.value, "{...}");

        let n = snap.locals.iter().find(|v| v.name == "n").unwrap();
        assert!(!n.expandable, "Number should not be expandable");

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_upvalues_collected() {
        let code = "\
local captured = 99
local function foo()
  return captured
end
foo()
return captured";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 3)],
        ));

        let snap = recv_paused(&session);
        assert_eq!(snap.line, 3);

        // 'captured' should appear as an upvalue of foo
        let names: Vec<&str> = snap.upvalues.iter().map(|v| v.name.as_str()).collect();
        assert!(
            names.contains(&"captured"),
            "Expected upvalue 'captured', got: {:?}",
            names
        );

        let captured = snap.upvalues.iter().find(|v| v.name == "captured").unwrap();
        assert_eq!(captured.value, "99");

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_value_truncation() {
        let code = format!(
            "local s = string.rep('x', 500)\nreturn s"
        );
        let session = spawn_debug_session(config_with_breakpoints(
            &code,
            vec![(CONSOLE_SOURCE, 2)],
        ));

        let snap = recv_paused(&session);
        let s = snap.locals.iter().find(|v| v.name == "s").unwrap();
        // Value should be truncated at MAX_VALUE_LEN + quotes + "..."
        assert!(
            s.value.len() <= MAX_VALUE_LEN + 10,
            "Value too long: {} chars",
            s.value.len()
        );
        assert!(s.value.ends_with("...\""), "Expected truncation: {}", s.value);

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_request_frame_vars_non_top_frame() {
        // Function with locals, breakpoint inside it
        let code = "\
local outer = 'outer_val'
local function foo(a)
  local inner = a * 2
  return inner
end
local result = foo(21)
return result";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 4)],
        ));

        let snap = recv_paused(&session);
        assert_eq!(snap.line, 4);

        // Top frame (foo) should have 'a' and 'inner'
        let top_names: Vec<&str> = snap.locals.iter().map(|v| v.name.as_str()).collect();
        assert!(top_names.contains(&"a"), "Expected 'a' in top frame, got: {:?}", top_names);
        assert!(top_names.contains(&"inner"), "Expected 'inner' in top frame, got: {:?}", top_names);

        // Request frame 1 (the caller = main chunk)
        send_cmd(
            &session,
            DebugAction::RequestFrameVars { frame_index: 1 },
        );
        let (frame_idx, locals, _upvalues) = recv_frame_vars(&session);
        assert_eq!(frame_idx, 1);

        let caller_names: Vec<&str> = locals.iter().map(|v| v.name.as_str()).collect();
        assert!(
            caller_names.contains(&"outer"),
            "Expected 'outer' in caller frame, got: {:?}",
            caller_names
        );
        assert!(
            caller_names.contains(&"foo"),
            "Expected 'foo' in caller frame, got: {:?}",
            caller_names
        );

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_expand_variable() {
        let code = "local t = {a = 1, b = 'two', c = {nested = true}}\nreturn t";
        let session = spawn_debug_session(config_with_breakpoints(
            code,
            vec![(CONSOLE_SOURCE, 2)],
        ));

        let snap = recv_paused(&session);
        let t_pos = snap.locals.iter().position(|v| v.name == "t").unwrap();
        assert!(snap.locals[t_pos].expandable);

        // Expand the table
        send_cmd(
            &session,
            DebugAction::ExpandVariable {
                frame_index: 0,
                path: vec![VarPathSegment::Local(t_pos)],
            },
        );

        let children = recv_expanded(&session);
        assert!(!children.is_empty(), "Table should have children");

        let names: Vec<&str> = children.iter().map(|v| v.name.as_str()).collect();
        assert!(names.contains(&"a"), "Expected key 'a', got: {:?}", names);
        assert!(names.contains(&"b"), "Expected key 'b', got: {:?}", names);
        assert!(names.contains(&"c"), "Expected key 'c', got: {:?}", names);

        // 'c' should be expandable (nested table)
        let c = children.iter().find(|v| v.name == "c").unwrap();
        assert!(c.expandable, "Nested table should be expandable");
        assert_eq!(c.value, "{...}");

        // 'a' should not be expandable
        let a = children.iter().find(|v| v.name == "a").unwrap();
        assert!(!a.expandable);
        assert_eq!(a.value, "1");

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }

    #[test]
    fn test_locals_bounded_at_max() {
        // Create code with many locals (Lua 5.4 max is 200 per function)
        let mut code = String::new();
        for i in 0..200 {
            code.push_str(&format!("local v{} = {}\n", i, i));
        }
        code.push_str("return v0");

        let session = spawn_debug_session(config_with_breakpoints(
            &code,
            vec![(CONSOLE_SOURCE, 201)],
        ));

        let snap = recv_paused(&session);
        assert!(
            snap.locals.len() <= MAX_LOCALS,
            "Locals should be capped at {}, got {}",
            MAX_LOCALS,
            snap.locals.len()
        );

        send_cmd(&session, DebugAction::Continue);
        recv_completed(&session);
    }
}
