//! Script View state — persists with the document.

use super::text_buffer::TextBuffer;

/// State for the full-panel script editor.
#[derive(Debug)]
pub struct ScriptState {
    /// Whether the Script View is open (replaces grid slot).
    pub open: bool,
    /// The script editor buffer (independent from REPL input).
    pub buffer: TextBuffer,
    /// Whether the Script View was opened from the console Expand button.
    /// Used to determine focus return target on close.
    pub opened_from_console: bool,
}

impl ScriptState {
    pub fn new() -> Self {
        Self {
            open: false,
            buffer: TextBuffer::new(),
            opened_from_console: false,
        }
    }
}

impl Default for ScriptState {
    fn default() -> Self {
        Self::new()
    }
}
