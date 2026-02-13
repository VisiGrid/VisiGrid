//! Reusable text editing buffer for code editors.
//!
//! Used by both the REPL input and the Script View editor.
//! Provides cursor movement, text manipulation, and cached tokenization.

use std::ops::Range;

use super::lua_tokenizer::{tokenize_lua, LuaTokenType};

/// A reusable text editing buffer with cursor, scroll, and cached tokenization.
///
/// Caches both tokens and line byte offsets, keyed by a text snapshot.
/// Both caches are refreshed together on first access after a text change.
#[derive(Debug)]
pub struct TextBuffer {
    pub text: String,
    pub cursor: usize,
    pub scroll_offset: usize,
    /// Sticky column for up/down navigation (prevents cursor drift through short lines).
    preferred_col: Option<usize>,
    // Caches — keyed by text snapshot, refreshed together via ensure_caches().
    cached_snapshot: String,
    cached_tokens: Vec<(Range<usize>, LuaTokenType)>,
    /// (line_start_byte, line_end_byte) for each line. Used by cursor_up/down and rendering.
    cached_line_offsets: Vec<(usize, usize)>,
}

/// Snap a byte position to the nearest char boundary at or before it.
/// Used by cursor_up/down where a byte-offset column may land inside a multi-byte char.
fn snap_to_char_boundary(text: &str, pos: usize) -> usize {
    let mut p = pos.min(text.len());
    while p > 0 && !text.is_char_boundary(p) {
        p -= 1;
    }
    p
}

/// Normalize \r\n and lone \r to \n. Returns a new String only if \r is present.
fn normalize_newlines(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}

/// Normalize \r\n and lone \r to \n. Avoids allocation if no \r is present.
fn normalize_newlines_owned(text: String) -> String {
    if text.as_bytes().contains(&b'\r') {
        normalize_newlines(&text)
    } else {
        text
    }
}

impl TextBuffer {
    pub fn new() -> Self {
        Self {
            text: String::new(),
            cursor: 0,
            scroll_offset: 0,
            preferred_col: None,
            cached_snapshot: String::new(),
            cached_tokens: Vec::new(),
            cached_line_offsets: vec![(0, 0)],
        }
    }

    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    pub fn len(&self) -> usize {
        self.text.len()
    }

    /// Clear all state (text, cursor, scroll).
    pub fn clear(&mut self) {
        self.text.clear();
        self.cursor = 0;
        self.scroll_offset = 0;
        self.preferred_col = None;
    }

    /// Replace contents, cursor to end. Normalizes CRLF/CR to LF.
    pub fn set_text(&mut self, text: String) {
        let text = normalize_newlines_owned(text);
        self.cursor = text.len();
        self.text = text;
        self.scroll_offset = 0;
        self.preferred_col = None;
    }

    /// Take text, reset cursor/scroll (for REPL execute).
    pub fn consume(&mut self) -> String {
        let text = std::mem::take(&mut self.text);
        self.cursor = 0;
        self.scroll_offset = 0;
        self.preferred_col = None;
        text
    }

    /// Insert text at cursor. Normalizes CRLF/CR to LF.
    pub fn insert(&mut self, text: &str) {
        if text.as_bytes().contains(&b'\r') {
            let normalized = normalize_newlines(text);
            self.text.insert_str(self.cursor, &normalized);
            self.cursor += normalized.len();
        } else {
            self.text.insert_str(self.cursor, text);
            self.cursor += text.len();
        }
        self.preferred_col = None;
    }

    /// Delete character before cursor (backspace).
    pub fn backspace(&mut self) {
        if self.cursor > 0 {
            let prev = self.text[..self.cursor]
                .char_indices()
                .last()
                .map(|(i, _)| i)
                .unwrap_or(0);
            self.text.replace_range(prev..self.cursor, "");
            self.cursor = prev;
            self.preferred_col = None;
        }
    }

    /// Delete character at cursor (delete key).
    pub fn delete(&mut self) {
        if self.cursor < self.text.len() {
            let next = self.text[self.cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.cursor + i)
                .unwrap_or(self.text.len());
            self.text.replace_range(self.cursor..next, "");
            self.preferred_col = None;
        }
    }

    /// Move cursor left one character.
    pub fn cursor_left(&mut self) {
        if self.cursor > 0 {
            self.cursor = self.text[..self.cursor]
                .char_indices()
                .last()
                .map(|(i, _)| i)
                .unwrap_or(0);
            self.preferred_col = None;
        }
    }

    /// Move cursor right one character.
    pub fn cursor_right(&mut self) {
        if self.cursor < self.text.len() {
            self.cursor = self.text[self.cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.cursor + i)
                .unwrap_or(self.text.len());
            self.preferred_col = None;
        }
    }

    /// Move cursor to start of current line.
    pub fn cursor_home(&mut self) {
        let line_start = self.text[..self.cursor]
            .rfind('\n')
            .map(|i| i + 1)
            .unwrap_or(0);
        self.cursor = line_start;
        self.preferred_col = None;
    }

    /// Move cursor to end of current line.
    pub fn cursor_end(&mut self) {
        let line_end = self.text[self.cursor..]
            .find('\n')
            .map(|i| self.cursor + i)
            .unwrap_or(self.text.len());
        self.cursor = line_end;
        self.preferred_col = None;
    }

    /// Move cursor to start of entire buffer.
    pub fn cursor_buffer_home(&mut self) {
        self.cursor = 0;
        self.preferred_col = None;
    }

    /// Move cursor to end of entire buffer.
    pub fn cursor_buffer_end(&mut self) {
        self.cursor = self.text.len();
        self.preferred_col = None;
    }

    /// Move cursor up one line, preserving column via preferred_col.
    pub fn cursor_up(&mut self) {
        let (current_line, current_col) = self.cursor_line_col();
        if current_line == 0 {
            return;
        }

        let target_col = self.preferred_col.unwrap_or(current_col);
        if self.preferred_col.is_none() {
            self.preferred_col = Some(current_col);
        }

        self.ensure_caches();
        let (prev_start, prev_end) = self.cached_line_offsets[current_line - 1];
        let prev_line_len = prev_end - prev_start;
        let pos = prev_start + target_col.min(prev_line_len);
        self.cursor = snap_to_char_boundary(&self.text, pos);
    }

    /// Move cursor down one line, preserving column via preferred_col.
    pub fn cursor_down(&mut self) {
        let (current_line, current_col) = self.cursor_line_col();

        self.ensure_caches();
        if current_line + 1 >= self.cached_line_offsets.len() {
            return;
        }

        let target_col = self.preferred_col.unwrap_or(current_col);
        if self.preferred_col.is_none() {
            self.preferred_col = Some(current_col);
        }

        let (next_start, next_end) = self.cached_line_offsets[current_line + 1];
        let next_line_len = next_end - next_start;
        let pos = next_start + target_col.min(next_line_len);
        self.cursor = snap_to_char_boundary(&self.text, pos);
    }

    /// Adjust scroll_offset so the cursor line is visible within max_visible_lines.
    pub fn ensure_cursor_visible(&mut self, max_visible_lines: usize) {
        let cursor_line = self.text[..self.cursor].matches('\n').count();
        if cursor_line < self.scroll_offset {
            self.scroll_offset = cursor_line;
        } else if cursor_line >= self.scroll_offset + max_visible_lines {
            self.scroll_offset = cursor_line + 1 - max_visible_lines;
        }
    }

    /// Refresh all caches (tokens + line offsets) if text has changed since last call.
    fn ensure_caches(&mut self) {
        if self.text == self.cached_snapshot {
            return;
        }
        self.cached_snapshot = self.text.clone();
        self.cached_tokens = tokenize_lua(&self.text);

        // Compute line offsets
        let newline_count = self.text.as_bytes().iter().filter(|&&b| b == b'\n').count();
        self.cached_line_offsets.clear();
        self.cached_line_offsets.reserve(newline_count + 1);
        let mut start = 0;
        for line in self.text.split('\n') {
            let end = start + line.len();
            self.cached_line_offsets.push((start, end));
            start = end + 1;
        }
        if self.cached_line_offsets.is_empty() {
            self.cached_line_offsets.push((0, 0));
        }
    }

    /// Return cached tokens, refreshing caches if text changed.
    pub fn tokens(&mut self) -> &[(Range<usize>, LuaTokenType)] {
        self.ensure_caches();
        &self.cached_tokens
    }

    /// Return cached line offsets (start, end byte pairs), refreshing caches if text changed.
    pub fn line_offsets(&mut self) -> &[(usize, usize)] {
        self.ensure_caches();
        &self.cached_line_offsets
    }

    /// Read-only access to cached tokens (for rendering after ensure_caches has been primed).
    pub fn cached_tokens(&self) -> &[(Range<usize>, LuaTokenType)] {
        &self.cached_tokens
    }

    /// Read-only access to cached line offsets (for rendering after ensure_caches has been primed).
    pub fn cached_line_offsets(&self) -> &[(usize, usize)] {
        &self.cached_line_offsets
    }

    /// Compute (line_number, column) of the cursor. Both 0-indexed.
    fn cursor_line_col(&self) -> (usize, usize) {
        let before = &self.text[..self.cursor];
        let line = before.matches('\n').count();
        let line_start = before.rfind('\n').map(|i| i + 1).unwrap_or(0);
        let col = self.cursor - line_start;
        (line, col)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_insert_and_cursor() {
        let mut buf = TextBuffer::new();
        buf.insert("hello");
        assert_eq!(buf.text, "hello");
        assert_eq!(buf.cursor, 5);

        buf.insert(" world");
        assert_eq!(buf.text, "hello world");
        assert_eq!(buf.cursor, 11);
    }

    #[test]
    fn test_backspace() {
        let mut buf = TextBuffer::new();
        buf.insert("abc");
        buf.backspace();
        assert_eq!(buf.text, "ab");
        assert_eq!(buf.cursor, 2);

        // Backspace at position 0 is a no-op
        buf.cursor = 0;
        buf.backspace();
        assert_eq!(buf.text, "ab");
    }

    #[test]
    fn test_delete() {
        let mut buf = TextBuffer::new();
        buf.insert("abc");
        buf.cursor = 1;
        buf.delete();
        assert_eq!(buf.text, "ac");
        assert_eq!(buf.cursor, 1);

        // Delete at end is a no-op
        buf.cursor = buf.text.len();
        buf.delete();
        assert_eq!(buf.text, "ac");
    }

    #[test]
    fn test_cursor_left_right() {
        let mut buf = TextBuffer::new();
        buf.insert("abc");
        assert_eq!(buf.cursor, 3);

        buf.cursor_left();
        assert_eq!(buf.cursor, 2);

        buf.cursor_left();
        buf.cursor_left();
        assert_eq!(buf.cursor, 0);

        // Left at 0 is no-op
        buf.cursor_left();
        assert_eq!(buf.cursor, 0);

        buf.cursor_right();
        assert_eq!(buf.cursor, 1);
    }

    #[test]
    fn test_cursor_home_end() {
        let mut buf = TextBuffer::new();
        buf.insert("line1\nline2\nline3");
        // Cursor is at end of "line3"
        buf.cursor_home();
        assert_eq!(buf.cursor, 12); // start of "line3"

        buf.cursor_end();
        assert_eq!(buf.cursor, 17); // end of "line3"

        // Move to middle of line2
        buf.cursor = 8; // in "line2"
        buf.cursor_home();
        assert_eq!(buf.cursor, 6); // start of "line2"

        buf.cursor_end();
        assert_eq!(buf.cursor, 11); // end of "line2"
    }

    #[test]
    fn test_cursor_buffer_home_end() {
        let mut buf = TextBuffer::new();
        buf.insert("line1\nline2");
        buf.cursor = 8;

        buf.cursor_buffer_home();
        assert_eq!(buf.cursor, 0);

        buf.cursor_buffer_end();
        assert_eq!(buf.cursor, 11);
    }

    #[test]
    fn test_cursor_up_down() {
        let mut buf = TextBuffer::new();
        buf.insert("hello\nworld\nfoo");
        // After insert, cursor is at end = 15 (col 3 on "foo")
        assert_eq!(buf.cursor, 15);

        buf.cursor_up();
        // Should be on "world" at col 3
        assert_eq!(buf.cursor, 9); // "hello\n" (6) + "wor" (3) = 9

        buf.cursor_up();
        // Should be on "hello" at col 3
        assert_eq!(buf.cursor, 3); // "hel" = 3

        buf.cursor_down();
        // Back to "world" at col 3
        assert_eq!(buf.cursor, 9);

        buf.cursor_down();
        // Back to "foo" at col 3 (end of "foo")
        assert_eq!(buf.cursor, 15); // "hello\nworld\n" (12) + "foo" (3) = 15
    }

    #[test]
    fn test_preferred_col_through_short_line() {
        let mut buf = TextBuffer::new();
        buf.insert("longline\na\nlongline");
        // After insert, cursor is at end = 19, col 8 on last "longline"
        assert_eq!(buf.cursor, 19);

        buf.cursor_up(); // to "a" (line 1), col clamped to 1 (len of "a")
        assert_eq!(buf.cursor, 10); // "longline\n" (9) + "a" (1) = 10

        // preferred_col should be 8 (from original position)
        buf.cursor_up(); // to first "longline", col should be 8 (preferred)
        assert_eq!(buf.cursor, 8);
    }

    #[test]
    fn test_clear_and_set_text() {
        let mut buf = TextBuffer::new();
        buf.insert("hello");
        buf.clear();
        assert!(buf.is_empty());
        assert_eq!(buf.cursor, 0);

        buf.set_text("new content".to_string());
        assert_eq!(buf.text, "new content");
        assert_eq!(buf.cursor, 11); // cursor at end
    }

    #[test]
    fn test_consume() {
        let mut buf = TextBuffer::new();
        buf.insert("script");
        let text = buf.consume();
        assert_eq!(text, "script");
        assert!(buf.is_empty());
        assert_eq!(buf.cursor, 0);
    }

    #[test]
    fn test_ensure_cursor_visible() {
        let mut buf = TextBuffer::new();
        buf.insert("a\nb\nc\nd\ne\nf\ng\nh\ni\nj");
        buf.cursor = buf.text.len(); // line 9
        buf.ensure_cursor_visible(3);
        assert_eq!(buf.scroll_offset, 7); // lines 7,8,9 visible
    }

    /// Verify line offset cache invariants after a sequence of mixed edits.
    ///
    /// After every mutation, cached_line_offsets must:
    /// 1. Cover the entire string (first starts at 0, last ends at text.len())
    /// 2. Be contiguous (each start == previous end + 1, accounting for '\n')
    /// 3. Each (start, end) slice must be on valid char boundaries
    /// 4. No line contains '\n'
    #[test]
    fn test_line_offsets_invariants_after_edits() {
        fn check_invariants(buf: &mut TextBuffer, label: &str) {
            let offsets = buf.line_offsets().to_vec();
            let text = &buf.text;

            assert!(!offsets.is_empty(), "{label}: offsets must never be empty");

            // First line starts at 0
            assert_eq!(offsets[0].0, 0, "{label}: first line must start at 0");

            // Last line ends at text.len()
            assert_eq!(
                offsets.last().unwrap().1,
                text.len(),
                "{label}: last line must end at text.len() ({})",
                text.len()
            );

            for (i, &(start, end)) in offsets.iter().enumerate() {
                // Valid range
                assert!(start <= end, "{label}: line {i} start ({start}) > end ({end})");
                assert!(end <= text.len(), "{label}: line {i} end ({end}) > text.len() ({})", text.len());

                // On char boundaries
                assert!(text.is_char_boundary(start), "{label}: line {i} start ({start}) not on char boundary");
                assert!(text.is_char_boundary(end), "{label}: line {i} end ({end}) not on char boundary");

                // No newline within the line
                let slice = &text[start..end];
                assert!(!slice.contains('\n'), "{label}: line {i} contains newline: {slice:?}");

                // Contiguous: this line's start == prev line's end + 1 (the '\n')
                if i > 0 {
                    let prev_end = offsets[i - 1].1;
                    assert_eq!(
                        start,
                        prev_end + 1,
                        "{label}: line {i} start ({start}) != prev end ({prev_end}) + 1"
                    );
                    // The byte between prev_end and start must be '\n'
                    assert_eq!(
                        text.as_bytes()[prev_end], b'\n',
                        "{label}: byte at {prev_end} between lines {} and {i} is not newline",
                        i - 1
                    );
                }
            }
        }

        let mut buf = TextBuffer::new();
        check_invariants(&mut buf, "empty");

        buf.insert("hello");
        check_invariants(&mut buf, "after insert 'hello'");

        buf.insert("\nworld");
        check_invariants(&mut buf, "after insert newline+world");

        buf.insert("\n\n\n");
        check_invariants(&mut buf, "after insert 3 newlines");

        buf.backspace();
        check_invariants(&mut buf, "after backspace");

        buf.cursor = 0;
        buf.insert("first\n");
        check_invariants(&mut buf, "after insert at start");

        buf.cursor = 5;
        buf.delete();
        check_invariants(&mut buf, "after delete at position 5");

        buf.set_text("a\nb\nc".to_string());
        check_invariants(&mut buf, "after set_text");

        buf.clear();
        check_invariants(&mut buf, "after clear");

        // Simulate rapid edits like typing a multiline script
        buf.insert("for i = 1, 100 do\n  print(i)\nend\n");
        check_invariants(&mut buf, "after multiline insert");

        buf.cursor = 18; // start of "  print(i)"
        buf.insert("  x = i * 2\n");
        check_invariants(&mut buf, "after mid-buffer insert");

        // Simulate PageUp: many cursor_up calls (should not corrupt offsets)
        for _ in 0..10 {
            buf.cursor_up();
        }
        check_invariants(&mut buf, "after 10x cursor_up");
    }

    #[test]
    fn test_crlf_normalization() {
        let mut buf = TextBuffer::new();

        // CRLF paste → normalized to LF
        buf.insert("line1\r\nline2\r\nline3");
        assert_eq!(buf.text, "line1\nline2\nline3");
        assert_eq!(buf.cursor, 17);

        // Lone CR → normalized to LF
        buf.clear();
        buf.insert("a\rb\rc");
        assert_eq!(buf.text, "a\nb\nc");

        // Mixed: CRLF + lone CR + LF
        buf.clear();
        buf.insert("a\r\nb\rc\nd");
        assert_eq!(buf.text, "a\nb\nc\nd");

        // set_text also normalizes
        buf.set_text("x\r\ny\rz".to_string());
        assert_eq!(buf.text, "x\ny\nz");
        assert_eq!(buf.cursor, 5); // cursor at end of normalized text

        // No \r → no allocation, text unchanged
        buf.clear();
        buf.insert("clean\ntext");
        assert_eq!(buf.text, "clean\ntext");
    }

    /// Fuzz-lite: 500 random operations, check invariants after each.
    #[test]
    fn test_fuzz_lite_line_offsets() {
        // Simple deterministic PRNG (xorshift32)
        struct Rng(u32);
        impl Rng {
            fn next(&mut self) -> u32 {
                self.0 ^= self.0 << 13;
                self.0 ^= self.0 >> 17;
                self.0 ^= self.0 << 5;
                self.0
            }
            fn range(&mut self, max: u32) -> u32 {
                self.next() % max
            }
        }

        fn check(buf: &mut TextBuffer, step: usize) {
            let offsets = buf.line_offsets().to_vec();
            let text = &buf.text;

            assert!(!offsets.is_empty(), "step {step}: empty offsets");
            assert_eq!(offsets[0].0, 0, "step {step}: first start != 0");
            assert_eq!(offsets.last().unwrap().1, text.len(), "step {step}: last end != text.len()");

            for (i, &(start, end)) in offsets.iter().enumerate() {
                assert!(start <= end && end <= text.len(), "step {step}: line {i} bounds");
                assert!(text.is_char_boundary(start), "step {step}: line {i} start boundary");
                assert!(text.is_char_boundary(end), "step {step}: line {i} end boundary");
                assert!(!text[start..end].contains('\n'), "step {step}: line {i} contains newline");
                if i > 0 {
                    assert_eq!(start, offsets[i - 1].1 + 1, "step {step}: line {i} not contiguous");
                }
            }

            // Cursor must be valid
            assert!(buf.cursor <= text.len(), "step {step}: cursor out of bounds");
            assert!(text.is_char_boundary(buf.cursor), "step {step}: cursor not on char boundary");
        }

        let mut rng = Rng(0xDEAD_BEEF);
        let mut buf = TextBuffer::new();

        let snippets = ["a", "hello", "\n", "  ", "for i=1,10 do\n  print(i)\nend", "\r\n", "\r", "\u{00e9}", "\u{1f600}", "x\r\ny"];

        for step in 0..500 {
            match rng.range(8) {
                0 => {
                    // Insert random snippet at cursor
                    let idx = rng.range(snippets.len() as u32) as usize;
                    buf.insert(snippets[idx]);
                }
                1 => {
                    // Backspace
                    buf.backspace();
                }
                2 => {
                    // Delete
                    buf.delete();
                }
                3 => {
                    // Cursor left
                    buf.cursor_left();
                }
                4 => {
                    // Cursor right
                    buf.cursor_right();
                }
                5 => {
                    // Cursor up
                    buf.cursor_up();
                }
                6 => {
                    // Cursor down
                    buf.cursor_down();
                }
                7 => {
                    // Move cursor to random valid position
                    if !buf.text.is_empty() {
                        let mut pos = rng.range(buf.text.len() as u32 + 1) as usize;
                        // Walk back to char boundary (slicing panics on non-boundary)
                        while pos > 0 && !buf.text.is_char_boundary(pos) {
                            pos -= 1;
                        }
                        buf.cursor = pos;
                    }
                }
                _ => unreachable!(),
            }
            check(&mut buf, step);
        }
    }
}
