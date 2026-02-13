//! Lua scripting subsystem for VisiGrid.
//!
//! # Architecture
//!
//! The scripting system follows a strict separation:
//!
//! 1. **LuaRuntime** owns the mlua::Lua instance and handles evaluation
//! 2. **ConsoleState** manages REPL panel UI state (input, output, history)
//! 3. **LuaOp** defines operations that scripts can queue
//! 4. **LuaOpSink** is the only interface Lua has to the workbook
//!
//! # Critical Rule
//!
//! **Lua never touches workbook state directly.**
//!
//! Scripts have access to a command sink that queues operations. The sink
//! consults a pending map for reads-after-writes, then falls back to a
//! workbook snapshot. After the chunk finishes, ops are applied as a batch
//! with a single undo entry.
//!
//! # Safety Guarantees
//!
//! - **Sandboxed**: No OS/file/network access
//! - **Limited**: 1M ops, 5K output lines, 100M instructions, 30s timeout
//! - **Single undo**: All changes from one script = one Ctrl+Z

mod console_state;
pub mod custom_functions;
pub mod debugger;
pub mod examples;
pub mod lua_tokenizer;
mod ops;
mod runtime;
pub mod script_state;
mod sheet_api;
pub mod text_buffer;

pub use console_state::{ConsoleState, ConsoleTab, ActiveDebugSession, OutputEntry, OutputKind, VIEW_LEN, DEFAULT_CONSOLE_HEIGHT, MIN_CONSOLE_HEIGHT, MAX_CONSOLE_HEIGHT, DEBUG_OUTPUT_CAP};
pub use script_state::ScriptState;
pub use text_buffer::TextBuffer;
pub use ops::{CellKey, LuaCellValue, LuaOp, LuaOpSink, PendingCell, SheetReader, parse_a1, format_a1};
pub use runtime::{LuaEvalResult, LuaRuntime, CancelToken, INSTRUCTION_LIMIT, INSTRUCTION_HOOK_INTERVAL, DEFAULT_TIMEOUT};
pub use sheet_api::{DynOpSink, SheetUserData, SheetSnapshot, register_sheet_global, MAX_OPS, MAX_OUTPUT_LINES};
pub use custom_functions::{CustomFunctionRegistry, CustomFunction, MemoCache};
pub use lua_tokenizer::LuaTokenType;
pub use debugger::{
    SessionId, DebugCommand, DebugAction, DebugEvent, DebugEventPayload,
    DebugSnapshot, DebugSession, DebugSessionState, DebugConfig,
    Variable, StackFrame, PauseReason, VarPathSegment,
    spawn_debug_session, CONSOLE_SOURCE,
};
