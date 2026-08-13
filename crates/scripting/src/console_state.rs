//! Console panel state for the Lua REPL.
//!
//! Manages UI state separately from the Lua runtime itself.
//!
//! ## Virtual Scroll
//!
//! The console supports virtual scrolling for large outputs. Instead of rendering
//! all lines, it shows a window of `VIEW_LEN` lines at a time. This prevents
//! performance issues with very long outputs (scripts can print thousands of lines).

use std::collections::{HashMap, HashSet};
use std::io::{BufRead, Write};
use std::ops::Range;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Arc};

use super::debugger::{
    DebugAction, DebugCommand, DebugEvent, DebugSession, DebugSessionState, DebugSnapshot,
    SessionId, Variable, VarPathSegment,
};
use super::lua_tokenizer::LuaTokenType;
pub use super::text_buffer::TextBuffer;

/// Default number of visible lines in the console output
pub const VIEW_LEN: usize = 200;

/// Kind of output entry in the console log
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OutputKind {
    /// User input (echoed)
    Input,
    /// Returned value from expression
    Result,
    /// Output from print() calls
    Print,
    /// Error message
    Error,
    /// System message (e.g., "100 cells modified")
    System,
    /// Execution metadata (ops/cells/time)
    Stats,
}

/// A single entry in the console output log
#[derive(Debug, Clone)]
pub struct OutputEntry {
    pub kind: OutputKind,
    pub text: String,
    /// Group id for visual grouping (0 = ungrouped)
    pub group_id: u32,
}

impl OutputEntry {
    pub fn input(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Input, text: text.into(), group_id: 0 }
    }

    pub fn result(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Result, text: text.into(), group_id: 0 }
    }

    pub fn print(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Print, text: text.into(), group_id: 0 }
    }

    pub fn error(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Error, text: text.into(), group_id: 0 }
    }

    pub fn system(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::System, text: text.into(), group_id: 0 }
    }

    pub fn stats(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Stats, text: text.into(), group_id: 0 }
    }
}

/// Which tab is active in the console panel
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConsoleTab {
    Run,
    Debug,
}

/// An active debug session, owned by ConsoleState.
pub struct ActiveDebugSession {
    pub id: SessionId,
    pub cmd_tx: mpsc::Sender<DebugCommand>,
    pub event_rx: mpsc::Receiver<DebugEvent>,
    pub cancel: Arc<AtomicBool>,
    pub state: DebugSessionState,
    pub start_sheet_index: usize,
}

// Manual Debug impl because mpsc types don't implement Debug
impl std::fmt::Debug for ActiveDebugSession {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ActiveDebugSession")
            .field("id", &self.id)
            .field("state", &self.state)
            .field("start_sheet_index", &self.start_sheet_index)
            .finish_non_exhaustive()
    }
}

/// State for the Lua console panel
#[derive(Debug)]
pub struct ConsoleState {
    /// Whether the console panel is visible
    pub visible: bool,

    /// REPL input buffer (text, cursor, scroll, tokens)
    pub input_buffer: TextBuffer,

    /// Output log (scrollable history of inputs, outputs, errors)
    pub output: Vec<OutputEntry>,

    /// Start index for virtual scroll view (0 = at top)
    pub view_start: usize,

    /// Whether view is pinned to bottom (auto-scroll on new output)
    pub view_pinned_to_bottom: bool,

    /// Command history (for up/down navigation)
    pub history: Vec<String>,

    /// Current position in history (-1 = not browsing, 0 = most recent)
    pub history_index: Option<usize>,

    /// Saved input when browsing history (to restore if user cancels)
    pub saved_input: Option<String>,

    /// Whether a script is currently executing
    pub executing: bool,

    /// Whether this is the first time the console has been opened (for welcome hint)
    pub first_open: bool,

    /// Panel height in pixels (resizable)
    pub height: f32,

    /// Whether the panel is maximized
    pub is_maximized: bool,

    /// Height to restore when un-maximizing
    pub restore_height: f32,

    /// Whether currently resizing the panel
    pub resizing: bool,

    /// Mouse Y position at resize start
    pub resize_start_y: f32,

    /// Panel height at resize start
    pub resize_start_height: f32,

    // ========================================================================
    // Debug session state
    // ========================================================================

    /// Which tab is active (Run or Debug)
    pub active_tab: ConsoleTab,

    /// Breakpoints as (source_name, line) pairs
    pub breakpoints: HashSet<(String, usize)>,

    /// Current debug session (if any)
    pub debug_session: Option<ActiveDebugSession>,

    /// Debug tab output log (cleared on session start, persists after stop)
    pub debug_output: Vec<OutputEntry>,

    /// Last pause snapshot from the debugger
    pub debug_snapshot: Option<DebugSnapshot>,

    /// Currently selected call stack frame index (0 = top frame)
    pub selected_frame: usize,

    /// Cache of (locals, upvalues) per frame index — populated by FrameVars events
    pub frame_vars_cache: HashMap<usize, (Vec<Variable>, Vec<Variable>)>,

    /// Expanded variable children, keyed by "F{frame}:{path}" string
    pub expanded_vars: HashMap<String, Vec<Variable>>,

    /// Source pane scroll offset (line index of first visible line, 0-indexed)
    pub debug_source_scroll: usize,

    /// Last known visible line count in source pane (updated by render)
    pub debug_source_viewport_lines: usize,

    // ========================================================================
    // Output grouping
    // ========================================================================

    /// Next group id to assign (monotonically increasing)
    pub next_group_id: u32,

    /// Current group id (set by begin_group, used by push_output; 0 = ungrouped)
    pub current_group_id: u32,
}

/// Ring buffer cap for debug_output (prevent unbounded memory growth)
pub const DEBUG_OUTPUT_CAP: usize = 10_000;

/// Default and minimum console height
pub const DEFAULT_CONSOLE_HEIGHT: f32 = 250.0;
pub const MIN_CONSOLE_HEIGHT: f32 = 100.0;
pub const MAX_CONSOLE_HEIGHT: f32 = 600.0;

// ============================================================================
// Console History Persistence (JSONL)
// ============================================================================

/// Maximum number of history entries to persist.
const MAX_HISTORY_ENTRIES: usize = 1000;

/// Path to the console history file.
fn history_file_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("console_history")
}

/// Load console history from the JSONL file.
///
/// Each line is a JSON object: `{"ts":"...","input":"..."}`.
/// Returns last `MAX_HISTORY_ENTRIES` entries (input field only).
pub fn load_history() -> Vec<String> {
    let path = history_file_path();
    let file = match std::fs::File::open(&path) {
        Ok(f) => f,
        Err(_) => return Vec::new(),
    };

    let reader = std::io::BufReader::new(file);
    let mut entries = Vec::new();

    for line in reader.lines().flatten() {
        if let Ok(obj) = serde_json::from_str::<serde_json::Value>(&line) {
            if let Some(input) = obj.get("input").and_then(|v| v.as_str()) {
                entries.push(input.to_string());
            }
        }
    }

    // Keep only the last MAX_HISTORY_ENTRIES
    if entries.len() > MAX_HISTORY_ENTRIES {
        entries = entries.split_off(entries.len() - MAX_HISTORY_ENTRIES);
    }

    entries
}

/// Append a single history entry to the JSONL file (crash-safe — written immediately).
pub fn append_history_entry(entry: &str) {
    let path = history_file_path();

    // Ensure parent directory exists
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }

    let ts = chrono::Utc::now().to_rfc3339();
    let obj = serde_json::json!({
        "ts": ts,
        "input": entry,
    });

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
    {
        let _ = writeln!(file, "{}", obj);
    }
}

impl Default for ConsoleState {
    fn default() -> Self {
        Self::new()
    }
}

impl ConsoleState {
    pub fn new() -> Self {
        let history = load_history();
        Self {
            visible: false,
            input_buffer: TextBuffer::new(),
            output: Vec::new(),
            view_start: 0,
            view_pinned_to_bottom: true,
            history,
            history_index: None,
            saved_input: None,
            executing: false,
            first_open: true,
            height: DEFAULT_CONSOLE_HEIGHT,
            is_maximized: false,
            restore_height: DEFAULT_CONSOLE_HEIGHT,
            resizing: false,
            resize_start_y: 0.0,
            resize_start_height: 0.0,
            active_tab: ConsoleTab::Run,
            breakpoints: HashSet::new(),
            debug_session: None,
            debug_output: Vec::new(),
            debug_snapshot: None,
            selected_frame: 0,
            frame_vars_cache: HashMap::new(),
            expanded_vars: HashMap::new(),
            debug_source_scroll: 0,
            debug_source_viewport_lines: 20,
            next_group_id: 0,
            current_group_id: 0,
        }
    }

    /// Toggle console visibility
    pub fn toggle(&mut self) {
        if self.visible {
            self.visible = false;
        } else {
            self.show();
        }
    }

    /// Show the console
    pub fn show(&mut self) {
        self.visible = true;

        // On first open, insert welcome hint
        if self.first_open {
            self.first_open = false;
            self.input_buffer.set_text("examples".to_string());
        }
    }

    /// Hide the console
    pub fn hide(&mut self) {
        self.visible = false;
    }

    /// Toggle maximize/restore
    pub fn toggle_maximize(&mut self, effective_max: f32) {
        if self.is_maximized {
            self.height = self.restore_height;
            self.is_maximized = false;
        } else {
            self.restore_height = self.height;
            self.height = effective_max;
            self.is_maximized = true;
        }
    }

    /// Set height from drag resize. Exits maximize mode.
    pub fn set_height_from_drag(&mut self, new_height: f32) {
        self.height = new_height.max(MIN_CONSOLE_HEIGHT).min(MAX_CONSOLE_HEIGHT);
        self.is_maximized = false;
    }

    /// Clear the output log
    pub fn clear_output(&mut self) {
        self.output.clear();
        self.view_start = 0;
        self.view_pinned_to_bottom = true;
        self.current_group_id = 0;
    }

    /// Add an entry to the output log, tagged with the current group id.
    pub fn push_output(&mut self, entry: OutputEntry) {
        let mut entry = entry;
        entry.group_id = self.current_group_id;
        self.output.push(entry);

        if self.view_pinned_to_bottom {
            self.scroll_to_end();
        }
    }

    /// Push an entry that is never grouped, regardless of current group state.
    pub fn push_output_ungrouped(&mut self, entry: OutputEntry) {
        let mut entry = entry;
        entry.group_id = 0;
        self.output.push(entry);

        if self.view_pinned_to_bottom {
            self.scroll_to_end();
        }
    }

    // ========================================================================
    // Output Grouping
    // ========================================================================

    /// Start a new output group. Returns the new group_id.
    pub fn begin_group(&mut self) -> u32 {
        self.next_group_id = self.next_group_id.wrapping_add(1);
        if self.next_group_id == 0 {
            self.next_group_id = 1;
        }
        self.current_group_id = self.next_group_id;
        self.current_group_id
    }

    /// End the current group. Resets to ungrouped (0).
    pub fn end_group(&mut self) {
        self.current_group_id = 0;
    }

    /// Get the group_id of the entry at `index`, or 0 if out of bounds.
    pub fn group_id_at(&self, index: usize) -> u32 {
        self.output.get(index).map(|e| e.group_id).unwrap_or(0)
    }

    // ========================================================================
    // Virtual Scroll
    // ========================================================================

    /// Get the visible output entries (virtual scroll window)
    pub fn visible_output(&self) -> &[OutputEntry] {
        if self.output.is_empty() {
            return &[];
        }
        let start = self.view_start.min(self.output.len().saturating_sub(1));
        let end = (start + VIEW_LEN).min(self.output.len());
        &self.output[start..end]
    }

    /// Check if we can scroll up
    pub fn can_scroll_up(&self) -> bool {
        self.view_start > 0
    }

    /// Check if we can scroll down
    pub fn can_scroll_down(&self) -> bool {
        self.view_start + VIEW_LEN < self.output.len()
    }

    /// Scroll up one page
    pub fn scroll_page_up(&mut self) {
        if self.view_start > 0 {
            self.view_start = self.view_start.saturating_sub(VIEW_LEN);
            self.view_pinned_to_bottom = false;
        }
    }

    /// Scroll down one page
    pub fn scroll_page_down(&mut self) {
        if self.can_scroll_down() {
            self.view_start = (self.view_start + VIEW_LEN).min(
                self.output.len().saturating_sub(VIEW_LEN)
            );
            // Check if we've reached the end
            if self.view_start + VIEW_LEN >= self.output.len() {
                self.view_pinned_to_bottom = true;
            }
        }
    }

    /// Scroll to the beginning of output
    pub fn scroll_to_start(&mut self) {
        self.view_start = 0;
        self.view_pinned_to_bottom = false;
    }

    /// Scroll to the end of output (pin to bottom)
    pub fn scroll_to_end(&mut self) {
        self.view_start = self.output.len().saturating_sub(VIEW_LEN);
        self.view_pinned_to_bottom = true;
    }

    /// Get scroll position info for UI (e.g., "lines 100-200 of 500")
    pub fn scroll_info(&self) -> Option<String> {
        if self.output.len() <= VIEW_LEN {
            return None; // No scrolling needed
        }
        let start = self.view_start + 1;
        let end = (self.view_start + VIEW_LEN).min(self.output.len());
        Some(format!("{}-{} of {}", start, end, self.output.len()))
    }

    // ========================================================================
    // Debug Session Lifecycle
    // ========================================================================

    /// Start a new debug session, stopping any existing one first.
    pub fn start_debug_session(&mut self, session: DebugSession, start_sheet_index: usize) {
        self.stop_debug_session();

        self.debug_session = Some(ActiveDebugSession {
            id: session.id,
            cmd_tx: session.cmd_tx,
            event_rx: session.event_rx,
            cancel: session.cancel,
            state: session.state,
            start_sheet_index,
        });
        self.active_tab = ConsoleTab::Debug;
        self.debug_output.clear();
        self.debug_snapshot = None;
        self.selected_frame = 0;
        self.frame_vars_cache.clear();
        self.expanded_vars.clear();
        self.debug_source_scroll = 0;
    }

    /// Stop the current debug session (idempotent).
    /// Sets cancel flag and drops the session (closing cmd_tx).
    pub fn stop_debug_session(&mut self) {
        if let Some(session) = self.debug_session.take() {
            session.cancel.store(true, Ordering::Relaxed);
            // Dropping session closes cmd_tx, which unblocks recv() in the debug thread
        }
        self.debug_snapshot = None;
        self.expanded_vars.clear();
        self.frame_vars_cache.clear();
        self.selected_frame = 0;
    }

    /// Send a debug action to the active session (no-op if no session).
    pub fn send_debug_action(&self, action: DebugAction) {
        if let Some(ref session) = self.debug_session {
            let _ = session.cmd_tx.send(DebugCommand {
                session_id: session.id,
                action,
            });
        }
    }

    /// Optimistically set session state to Running after sending a step/continue command.
    pub fn set_debug_running(&mut self) {
        if let Some(ref mut session) = self.debug_session {
            session.state = DebugSessionState::Running;
        }
    }

    /// Called when debugger pauses. Resets UI state and ensures paused line is visible.
    /// `paused_line_0` is 0-indexed (caller normalizes from 1-indexed Lua line).
    pub fn on_debug_paused(&mut self, paused_line_0: usize, total_lines: usize) {
        self.selected_frame = 0;
        self.frame_vars_cache.clear();
        self.expanded_vars.clear();

        // Ensure paused line visible with 3-line margin
        let visible = self.debug_source_viewport_lines;
        let scroll = &mut self.debug_source_scroll;
        if paused_line_0 < *scroll + 3 {
            *scroll = paused_line_0.saturating_sub(3);
        } else if paused_line_0 > *scroll + visible.saturating_sub(3) {
            *scroll = paused_line_0.saturating_sub(visible.saturating_sub(3));
        }
        *scroll = (*scroll).min(total_lines.saturating_sub(1));
    }

    /// Select a call stack frame for variable inspection.
    pub fn select_frame(&mut self, frame_index: usize) {
        self.selected_frame = frame_index;
        self.expanded_vars.clear();
        if frame_index != 0 && !self.frame_vars_cache.contains_key(&frame_index) {
            self.send_debug_action(DebugAction::RequestFrameVars { frame_index });
        }
    }

    /// Build an expansion key for the variable cache, including frame index.
    pub fn var_expansion_key(frame_index: usize, path: &[VarPathSegment]) -> String {
        let mut key = format!("F{}:", frame_index);
        for seg in path {
            match seg {
                VarPathSegment::Local(i) => key.push_str(&format!("L{}", i)),
                VarPathSegment::Upvalue(i) => key.push_str(&format!("U{}", i)),
                VarPathSegment::KeyString(s) => key.push_str(&format!(".{}", s)),
                VarPathSegment::KeyInt(i) => key.push_str(&format!("[{}]", i)),
                VarPathSegment::KeyOther(s) => key.push_str(&format!("?{}", s)),
            }
        }
        key
    }

    /// Adjust scroll so the cursor line is visible within max visible lines.
    pub fn ensure_input_cursor_visible(&mut self, max_visible_lines: usize) {
        self.input_buffer.ensure_cursor_visible(max_visible_lines);
    }

    /// Return cached tokens for the current input.
    pub fn tokens(&mut self) -> &[(Range<usize>, LuaTokenType)] {
        self.input_buffer.tokens()
    }

    /// Get current input, consuming it and adding to history.
    ///
    /// Persists to JSONL history file immediately (crash-safe).
    pub fn consume_input(&mut self) -> String {
        let input = self.input_buffer.consume();

        // Add to history if non-empty and different from last entry
        if !input.trim().is_empty() {
            if self.history.last().map(|s| s.as_str()) != Some(&input) {
                self.history.push(input.clone());
                // Persist immediately (crash-safe)
                append_history_entry(&input);
            }
        }

        // Reset history browsing
        self.history_index = None;
        self.saved_input = None;

        input
    }

    /// Navigate to previous history entry (up arrow)
    pub fn history_prev(&mut self) {
        if self.history.is_empty() {
            return;
        }

        match self.history_index {
            None => {
                // Start browsing - save current input
                self.saved_input = Some(self.input_buffer.text.clone());
                self.history_index = Some(self.history.len() - 1);
                self.input_buffer.set_text(self.history[self.history.len() - 1].clone());
            }
            Some(idx) if idx > 0 => {
                // Go further back
                self.history_index = Some(idx - 1);
                self.input_buffer.set_text(self.history[idx - 1].clone());
            }
            Some(_) => {
                // Already at oldest entry
            }
        }
        self.ensure_input_cursor_visible(12);
    }

    /// Navigate to next history entry (down arrow)
    pub fn history_next(&mut self) {
        match self.history_index {
            None => {
                // Not browsing history
            }
            Some(idx) => {
                if idx + 1 < self.history.len() {
                    // Go forward in history
                    self.history_index = Some(idx + 1);
                    self.input_buffer.set_text(self.history[idx + 1].clone());
                } else {
                    // Return to saved input
                    self.history_index = None;
                    if let Some(saved) = self.saved_input.take() {
                        self.input_buffer.set_text(saved);
                    }
                }
            }
        }
        self.ensure_input_cursor_visible(12);
    }

    /// Insert text at cursor (delegates to input_buffer).
    pub fn insert(&mut self, text: &str) {
        self.input_buffer.insert(text);
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Delete character before cursor (delegates to input_buffer).
    pub fn backspace(&mut self) {
        self.input_buffer.backspace();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Delete character at cursor (delegates to input_buffer).
    pub fn delete(&mut self) {
        self.input_buffer.delete();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor left (delegates to input_buffer).
    pub fn cursor_left(&mut self) {
        self.input_buffer.cursor_left();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor right (delegates to input_buffer).
    pub fn cursor_right(&mut self) {
        self.input_buffer.cursor_right();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor to start of current line.
    pub fn cursor_home(&mut self) {
        self.input_buffer.cursor_home();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor to end of current line.
    pub fn cursor_end(&mut self) {
        self.input_buffer.cursor_end();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor to start of buffer.
    pub fn cursor_buffer_home(&mut self) {
        self.input_buffer.cursor_buffer_home();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// Move cursor to end of buffer.
    pub fn cursor_buffer_end(&mut self) {
        self.input_buffer.cursor_buffer_end();
        self.input_buffer.ensure_cursor_visible(12);
    }

    /// RAII guard: begin a group that ends when the guard is dropped.
    /// Restores the previous group_id (usually 0) even on early return or panic.
    ///
    /// Note: this borrows `&mut ConsoleState` for the guard's lifetime, so it
    /// cannot be used in contexts that also need `&mut Spreadsheet`.  For those,
    /// use the wrapper-function pattern (begin in outer, body can early-return,
    /// end in outer after body returns).
    pub fn begin_group_scoped(&mut self) -> OutputGroupGuard<'_> {
        let prev = self.current_group_id;
        self.begin_group();
        OutputGroupGuard { console: self, prev_gid: prev }
    }
}

/// RAII guard that restores `current_group_id` when dropped.
pub struct OutputGroupGuard<'a> {
    pub console: &'a mut ConsoleState,
    prev_gid: u32,
}

impl Drop for OutputGroupGuard<'_> {
    fn drop(&mut self) {
        self.console.current_group_id = self.prev_gid;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_history_navigation() {
        let mut state = ConsoleState::new();
        state.history.clear(); // Clear any persisted history for deterministic test

        // Add some history
        state.input_buffer.set_text("first".to_string());
        state.consume_input();
        state.input_buffer.set_text("second".to_string());
        state.consume_input();
        state.input_buffer.set_text("third".to_string());
        state.consume_input();

        assert_eq!(state.history, vec!["first", "second", "third"]);

        // Type something new
        state.input_buffer.set_text("current".to_string());

        // Go back in history
        state.history_prev();
        assert_eq!(state.input_buffer.text, "third");
        assert_eq!(state.history_index, Some(2));

        state.history_prev();
        assert_eq!(state.input_buffer.text, "second");
        assert_eq!(state.history_index, Some(1));

        // Go forward
        state.history_next();
        assert_eq!(state.input_buffer.text, "third");
        assert_eq!(state.history_index, Some(2));

        // Go past end returns to saved input
        state.history_next();
        assert_eq!(state.input_buffer.text, "current");
        assert_eq!(state.history_index, None);
    }

    #[test]
    fn test_duplicate_history_prevention() {
        let mut state = ConsoleState::new();
        state.history.clear(); // Clear any persisted history for deterministic test

        state.input_buffer.set_text("same".to_string());
        state.consume_input();
        state.input_buffer.set_text("same".to_string());
        state.consume_input();
        state.input_buffer.set_text("same".to_string());
        state.consume_input();

        // Should only have one entry
        assert_eq!(state.history.len(), 1);
    }

    #[test]
    fn test_cursor_movement() {
        let mut state = ConsoleState::new();
        state.input_buffer.set_text("hello".to_string());

        state.cursor_left();
        assert_eq!(state.input_buffer.cursor, 4);

        state.cursor_home();
        assert_eq!(state.input_buffer.cursor, 0);

        state.cursor_end();
        assert_eq!(state.input_buffer.cursor, 5);
    }
}
