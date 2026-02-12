//! Console panel state for the Lua REPL.
//!
//! Manages UI state separately from the Lua runtime itself.
//!
//! ## Virtual Scroll
//!
//! The console supports virtual scrolling for large outputs. Instead of rendering
//! all lines, it shows a window of `VIEW_LEN` lines at a time. This prevents
//! performance issues with very long outputs (scripts can print thousands of lines).

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
}

/// A single entry in the console output log
#[derive(Debug, Clone)]
pub struct OutputEntry {
    pub kind: OutputKind,
    pub text: String,
}

impl OutputEntry {
    pub fn input(text: impl Into<String>) -> Self {
        Self {
            kind: OutputKind::Input,
            text: text.into(),
        }
    }

    pub fn result(text: impl Into<String>) -> Self {
        Self {
            kind: OutputKind::Result,
            text: text.into(),
        }
    }

    pub fn print(text: impl Into<String>) -> Self {
        Self {
            kind: OutputKind::Print,
            text: text.into(),
        }
    }

    pub fn error(text: impl Into<String>) -> Self {
        Self {
            kind: OutputKind::Error,
            text: text.into(),
        }
    }

    pub fn system(text: impl Into<String>) -> Self {
        Self {
            kind: OutputKind::System,
            text: text.into(),
        }
    }
}

/// State for the Lua console panel
#[derive(Debug)]
pub struct ConsoleState {
    /// Whether the console panel is visible
    pub visible: bool,

    /// Current input text
    pub input: String,

    /// Cursor position in input (byte offset)
    pub cursor: usize,

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
}

/// Default and minimum console height
pub const DEFAULT_CONSOLE_HEIGHT: f32 = 250.0;
pub const MIN_CONSOLE_HEIGHT: f32 = 100.0;
pub const MAX_CONSOLE_HEIGHT: f32 = 600.0;

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
            is_maximized: false,
            restore_height: DEFAULT_CONSOLE_HEIGHT,
            resizing: false,
            resize_start_y: 0.0,
            resize_start_height: 0.0,
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
            self.input = "examples".to_string();
            self.cursor = self.input.len();
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
    }

    /// Add an entry to the output log
    pub fn push_output(&mut self, entry: OutputEntry) {
        self.output.push(entry);

        // If pinned to bottom, auto-scroll to show new content
        if self.view_pinned_to_bottom {
            self.scroll_to_end();
        }
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

    /// Get current input, consuming it and adding to history
    pub fn consume_input(&mut self) -> String {
        let input = std::mem::take(&mut self.input);
        self.cursor = 0;

        // Add to history if non-empty and different from last entry
        if !input.trim().is_empty() {
            if self.history.last().map(|s| s.as_str()) != Some(&input) {
                self.history.push(input.clone());
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
                self.saved_input = Some(self.input.clone());
                self.history_index = Some(self.history.len() - 1);
                self.input = self.history[self.history.len() - 1].clone();
            }
            Some(idx) if idx > 0 => {
                // Go further back
                self.history_index = Some(idx - 1);
                self.input = self.history[idx - 1].clone();
            }
            Some(_) => {
                // Already at oldest entry
            }
        }
        self.cursor = self.input.len();
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
                    self.input = self.history[idx + 1].clone();
                } else {
                    // Return to saved input
                    self.history_index = None;
                    if let Some(saved) = self.saved_input.take() {
                        self.input = saved;
                    }
                }
            }
        }
        self.cursor = self.input.len();
    }

    /// Insert text at cursor
    pub fn insert(&mut self, text: &str) {
        self.input.insert_str(self.cursor, text);
        self.cursor += text.len();
    }

    /// Delete character before cursor (backspace)
    pub fn backspace(&mut self) {
        if self.cursor > 0 {
            // Find the previous character boundary
            let prev = self.input[..self.cursor]
                .char_indices()
                .last()
                .map(|(i, _)| i)
                .unwrap_or(0);
            self.input.replace_range(prev..self.cursor, "");
            self.cursor = prev;
        }
    }

    /// Delete character at cursor (delete)
    pub fn delete(&mut self) {
        if self.cursor < self.input.len() {
            // Find the next character boundary
            let next = self.input[self.cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.cursor + i)
                .unwrap_or(self.input.len());
            self.input.replace_range(self.cursor..next, "");
        }
    }

    /// Move cursor left
    pub fn cursor_left(&mut self) {
        if self.cursor > 0 {
            self.cursor = self.input[..self.cursor]
                .char_indices()
                .last()
                .map(|(i, _)| i)
                .unwrap_or(0);
        }
    }

    /// Move cursor right
    pub fn cursor_right(&mut self) {
        if self.cursor < self.input.len() {
            self.cursor = self.input[self.cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.cursor + i)
                .unwrap_or(self.input.len());
        }
    }

    /// Move cursor to start
    pub fn cursor_home(&mut self) {
        self.cursor = 0;
    }

    /// Move cursor to end
    pub fn cursor_end(&mut self) {
        self.cursor = self.input.len();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_history_navigation() {
        let mut state = ConsoleState::new();

        // Add some history
        state.input = "first".to_string();
        state.consume_input();
        state.input = "second".to_string();
        state.consume_input();
        state.input = "third".to_string();
        state.consume_input();

        assert_eq!(state.history, vec!["first", "second", "third"]);

        // Type something new
        state.input = "current".to_string();

        // Go back in history
        state.history_prev();
        assert_eq!(state.input, "third");
        assert_eq!(state.history_index, Some(2));

        state.history_prev();
        assert_eq!(state.input, "second");
        assert_eq!(state.history_index, Some(1));

        // Go forward
        state.history_next();
        assert_eq!(state.input, "third");
        assert_eq!(state.history_index, Some(2));

        // Go past end returns to saved input
        state.history_next();
        assert_eq!(state.input, "current");
        assert_eq!(state.history_index, None);
    }

    #[test]
    fn test_duplicate_history_prevention() {
        let mut state = ConsoleState::new();

        state.input = "same".to_string();
        state.consume_input();
        state.input = "same".to_string();
        state.consume_input();
        state.input = "same".to_string();
        state.consume_input();

        // Should only have one entry
        assert_eq!(state.history.len(), 1);
    }

    #[test]
    fn test_cursor_movement() {
        let mut state = ConsoleState::new();
        state.input = "hello".to_string();
        state.cursor = 5;

        state.cursor_left();
        assert_eq!(state.cursor, 4);

        state.cursor_home();
        assert_eq!(state.cursor, 0);

        state.cursor_end();
        assert_eq!(state.cursor, 5);
    }
}
