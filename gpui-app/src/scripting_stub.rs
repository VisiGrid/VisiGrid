//! Stub scripting module for Free edition.
//!
//! Provides the same types as the real scripting module but shows
//! an upgrade prompt when the user tries to use Lua features.

/// Default number of visible lines in the console output
pub const VIEW_LEN: usize = 200;
pub const DEFAULT_CONSOLE_HEIGHT: f32 = 200.0;
pub const MIN_CONSOLE_HEIGHT: f32 = 100.0;
pub const MAX_CONSOLE_HEIGHT: f32 = 600.0;

/// Kind of output entry in the console log
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OutputKind {
    Input,
    Result,
    Print,
    Error,
    System,
}

/// A single entry in the console output log
#[derive(Debug, Clone)]
pub struct OutputEntry {
    pub kind: OutputKind,
    pub text: String,
}

impl OutputEntry {
    pub fn input(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Input, text: text.into() }
    }
    pub fn result(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Result, text: text.into() }
    }
    pub fn print(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Print, text: text.into() }
    }
    pub fn error(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::Error, text: text.into() }
    }
    pub fn system(text: impl Into<String>) -> Self {
        Self { kind: OutputKind::System, text: text.into() }
    }
}

/// Console state for Free edition - shows upgrade prompt
#[derive(Debug)]
pub struct ConsoleState {
    pub visible: bool,
    pub input: String,
    pub cursor: usize,
    pub output: Vec<OutputEntry>,
    pub view_start: usize,
    pub view_pinned_to_bottom: bool,
    pub history: Vec<String>,
    pub history_index: Option<usize>,
    pub saved_input: Option<String>,
    pub executing: bool,
    pub first_open: bool,
    pub height: f32,
    pub resizing: bool,
    pub resize_start_y: f32,
    pub resize_start_height: f32,
}

impl Default for ConsoleState {
    fn default() -> Self {
        Self::new()
    }
}

impl ConsoleState {
    pub fn new() -> Self {
        Self {
            visible: false,
            input: String::new(),
            cursor: 0,
            output: Vec::new(),
            view_start: 0,
            view_pinned_to_bottom: true,
            history: Vec::new(),
            history_index: None,
            saved_input: None,
            executing: false,
            first_open: true,
            height: DEFAULT_CONSOLE_HEIGHT,
            resizing: false,
            resize_start_y: 0.0,
            resize_start_height: 0.0,
        }
    }

    /// Toggle console - shows upgrade message in Free edition
    pub fn toggle(&mut self) {
        if self.visible {
            self.visible = false;
        } else {
            self.show();
        }
    }

    /// Show console - displays upgrade message
    pub fn show(&mut self) {
        self.visible = true;
        if self.first_open {
            self.first_open = false;
            self.output.clear();
            self.output.push(OutputEntry::system("Lua Scripting is a Pro feature"));
            self.output.push(OutputEntry::system(""));
            self.output.push(OutputEntry::system("Upgrade to VisiGrid Pro for:"));
            self.output.push(OutputEntry::print("  - Lua scripting & automation"));
            self.output.push(OutputEntry::print("  - Large-file mode (million+ rows)"));
            self.output.push(OutputEntry::print("  - Advanced transforms"));
            self.output.push(OutputEntry::print("  - Inspector panel"));
            self.output.push(OutputEntry::print("  - Plugin runtime"));
            self.output.push(OutputEntry::system(""));
            self.output.push(OutputEntry::system("Visit visigrid.com/pro to upgrade"));
        }
    }

    pub fn hide(&mut self) {
        self.visible = false;
    }

    pub fn clear_output(&mut self) {
        self.output.clear();
        self.view_start = 0;
        self.view_pinned_to_bottom = true;
    }

    pub fn push_output(&mut self, entry: OutputEntry) {
        self.output.push(entry);
        if self.view_pinned_to_bottom {
            self.scroll_to_end();
        }
    }

    pub fn visible_output(&self) -> &[OutputEntry] {
        if self.output.is_empty() {
            return &[];
        }
        let start = self.view_start.min(self.output.len().saturating_sub(1));
        let end = (start + VIEW_LEN).min(self.output.len());
        &self.output[start..end]
    }

    pub fn can_scroll_up(&self) -> bool { self.view_start > 0 }
    pub fn can_scroll_down(&self) -> bool { self.view_start + VIEW_LEN < self.output.len() }

    pub fn scroll_page_up(&mut self) {
        if self.view_start > 0 {
            self.view_start = self.view_start.saturating_sub(VIEW_LEN);
            self.view_pinned_to_bottom = false;
        }
    }

    pub fn scroll_page_down(&mut self) {
        if self.can_scroll_down() {
            self.view_start = (self.view_start + VIEW_LEN).min(
                self.output.len().saturating_sub(VIEW_LEN)
            );
            if self.view_start + VIEW_LEN >= self.output.len() {
                self.view_pinned_to_bottom = true;
            }
        }
    }

    pub fn scroll_to_start(&mut self) {
        self.view_start = 0;
        self.view_pinned_to_bottom = false;
    }

    pub fn scroll_to_end(&mut self) {
        self.view_start = self.output.len().saturating_sub(VIEW_LEN);
        self.view_pinned_to_bottom = true;
    }

    pub fn scroll_info(&self) -> Option<String> {
        if self.output.len() <= VIEW_LEN {
            return None;
        }
        let start = self.view_start + 1;
        let end = (self.view_start + VIEW_LEN).min(self.output.len());
        Some(format!("{}-{} of {}", start, end, self.output.len()))
    }

    pub fn consume_input(&mut self) -> String {
        let input = std::mem::take(&mut self.input);
        self.cursor = 0;
        self.history_index = None;
        self.saved_input = None;
        input
    }

    pub fn history_prev(&mut self) {}
    pub fn history_next(&mut self) {}

    pub fn insert(&mut self, _text: &str) {
        // No-op in Free edition
    }

    pub fn backspace(&mut self) {}
    pub fn delete(&mut self) {}
    pub fn cursor_left(&mut self) {}
    pub fn cursor_right(&mut self) {}
    pub fn cursor_home(&mut self) {}
    pub fn cursor_end(&mut self) {}
}

/// Stub Lua runtime - does nothing
#[derive(Debug, Default)]
pub struct LuaRuntime;

/// Stub eval result
pub struct LuaEvalResult {
    pub output: Vec<String>,
    pub returned: Option<String>,
    pub error: Option<String>,
    pub ops: Vec<LuaOp>,
}

impl LuaEvalResult {
    pub fn has_mutations(&self) -> bool { false }
}

impl LuaRuntime {
    pub fn eval_with_sheet_and_selection(
        &mut self,
        _code: &str,
        _snapshot: Box<dyn SheetReader>,
        _selection: (usize, usize, usize, usize),
    ) -> LuaEvalResult {
        LuaEvalResult {
            output: vec!["Lua scripting requires VisiGrid Pro".to_string()],
            returned: None,
            error: Some("Upgrade to Pro for Lua scripting".to_string()),
            ops: vec![],
        }
    }
}

// Stub types needed for compilation
pub struct CancelToken;
pub const INSTRUCTION_LIMIT: u64 = 0;
pub const INSTRUCTION_HOOK_INTERVAL: u32 = 0;
pub const DEFAULT_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(0);
pub const MAX_OPS: usize = 0;
pub const MAX_OUTPUT_LINES: usize = 0;

pub type CellKey = (usize, usize);

#[derive(Debug, Clone)]
pub enum LuaCellValue {
    Nil,
    Number(f64),
    Text(String),
    Bool(bool),
}

#[derive(Debug, Clone)]
pub enum LuaOp {
    SetCell { key: CellKey, value: LuaCellValue },
}

pub trait SheetReader: Send {
    fn get(&self, row: usize, col: usize) -> LuaCellValue;
}

pub struct SheetSnapshot;

impl SheetSnapshot {
    pub fn from_sheet(_sheet: &visigrid_engine::sheet::Sheet) -> Self {
        Self
    }
}

impl SheetReader for SheetSnapshot {
    fn get(&self, _row: usize, _col: usize) -> LuaCellValue {
        LuaCellValue::Nil
    }
}

// Stub custom functions types
pub struct CustomFunctionRegistry {
    pub functions: std::collections::HashMap<String, CustomFunction>,
    pub warnings: Vec<String>,
}

pub struct CustomFunction {
    pub name: String,
}

pub struct MemoCache;

impl CustomFunctionRegistry {
    pub fn empty() -> Self {
        Self {
            functions: std::collections::HashMap::new(),
            warnings: Vec::new(),
        }
    }
}

impl MemoCache {
    pub fn new() -> Self { Self }
}

// Stub examples module
pub mod examples {
    pub struct LuaExample {
        pub name: &'static str,
        pub description: &'static str,
        pub code: &'static str,
    }

    pub const EXAMPLES: &[LuaExample] = &[];

    pub fn get_example(_index: usize) -> Option<&'static LuaExample> { None }
    pub fn find_example(_name: &str) -> Option<&'static LuaExample> { None }
}
