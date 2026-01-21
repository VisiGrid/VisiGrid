//! Keyboard hints system (Vimium-style jump navigation)
//!
//! Provides fast cell navigation by showing letter labels on visible cells.
//! Press 'g' to enter hint mode, type the hint letters to jump directly.
//!
//! # Input Resolution
//!
//! The hint buffer is resolved in phases (most specific first):
//!
//! ```text
//! HintInputBuffer
//! → Phase 1: Exact command match (gg, g$, g0, etc.)
//! → Phase 2: Prefix command match (reserved for future)
//! → Phase 3: Cell label resolution (a, ab, zz, etc.)
//! → Phase 4: Fallback / no match
//! ```
//!
//! This architecture keeps the grammar extensible without committing to
//! specific semantics too early. Commands are checked first; if no command
//! matches, the buffer is treated as a cell label.

// ============================================================================
// Input Resolution
// ============================================================================

/// How hint mode was exited (for analytics/debugging)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HintExitReason {
    /// User pressed Escape
    Cancelled,
    /// Resolved to a g-command (gg, g$, etc.)
    Command,
    /// Resolved to a cell label jump
    LabelJump,
    /// No matches found, auto-exited
    NoMatch,
}

/// G-prefixed commands (resolved before cell labels)
///
/// These are exact-match commands. The buffer must match exactly.
/// Keep this list minimal - let demand pull complexity.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GCommand {
    /// gg - Go to cell A1 (top of sheet)
    GotoTop,
    // Future: G$ - end of row, G0 - start of row, etc.
    // Add only when users explicitly ask.
}

impl GCommand {
    /// Try to match a buffer to an exact command.
    ///
    /// Returns Some(command) if buffer exactly matches a command string.
    /// Returns None if no match (buffer should be treated as cell label).
    pub fn from_buffer(buffer: &str) -> Option<GCommand> {
        // Phase 1: Exact command match
        // No heuristics. Explicit string match only.
        match buffer {
            "g" => Some(GCommand::GotoTop), // gg (first g entered hint mode, second g is "g" in buffer)
            _ => None,
        }
    }

    /// Check if buffer is a prefix of any command (for future use).
    ///
    /// Returns true if the buffer could potentially become a command
    /// with more characters. Used to avoid premature label resolution.
    #[allow(dead_code)]
    pub fn is_command_prefix(buffer: &str) -> bool {
        // Phase 2: Prefix command match (reserved)
        // Currently only "g" could become "gg", but we resolve "g" immediately.
        // This is here for future commands like "gt" (next tab) where "g" alone
        // shouldn't resolve yet.
        matches!(buffer, "g")
    }
}

/// Result of resolving the hint buffer
#[derive(Debug, Clone)]
pub enum HintResolution {
    /// Waiting for more input
    Pending,
    /// Execute a g-command
    Command(GCommand),
    /// Jump to a cell (row, col)
    Jump(usize, usize),
    /// No matches - should exit
    NoMatch,
}

/// Resolve the current hint buffer against commands and labels.
///
/// Resolution order (most specific first):
/// 1. Exact command match → execute command
/// 2. Unique label match → jump to cell
/// 3. Multiple label matches → keep waiting
/// 4. No matches → exit hint mode
pub fn resolve_hint_buffer(state: &HintState) -> HintResolution {
    let buffer = &state.buffer;

    // Phase 1: Exact command match
    if let Some(cmd) = GCommand::from_buffer(buffer) {
        return HintResolution::Command(cmd);
    }

    // Phase 2: Reserved for prefix command matching (future)
    // Currently we resolve commands immediately, but this phase would
    // allow for commands that share prefixes with cell labels.

    // Phase 3: Cell label resolution
    if let Some(hint) = state.unique_match() {
        return HintResolution::Jump(hint.row, hint.col);
    }

    // Phase 4: Check if we should keep waiting or give up
    if state.no_matches() {
        return HintResolution::NoMatch;
    }

    // Still have multiple matches - wait for more input
    HintResolution::Pending
}

// ============================================================================
// Core Types
// ============================================================================

/// A hint label for a single cell
#[derive(Debug, Clone)]
pub struct HintLabel {
    pub row: usize,
    pub col: usize,
    pub label: String,
}

/// State for the hint system
#[derive(Debug, Clone, Default)]
pub struct HintState {
    /// Typed hint characters so far
    pub buffer: String,
    /// Generated hint labels for visible cells
    pub labels: Vec<HintLabel>,
    /// Viewport when hints were generated (scroll_row, scroll_col, visible_rows, visible_cols)
    pub viewport: (usize, usize, usize, usize),
    /// How the last hint session exited (for analytics)
    pub last_exit_reason: Option<HintExitReason>,
}

impl HintState {
    /// Clear all hint state
    pub fn clear(&mut self) {
        self.buffer.clear();
        self.labels.clear();
        self.viewport = (0, 0, 0, 0);
    }

    /// Get labels that match the current buffer prefix
    pub fn matching_labels(&self) -> Vec<&HintLabel> {
        if self.buffer.is_empty() {
            self.labels.iter().collect()
        } else {
            self.labels
                .iter()
                .filter(|h| h.label.starts_with(&self.buffer))
                .collect()
        }
    }

    /// Check if there's exactly one matching label (time to jump)
    pub fn unique_match(&self) -> Option<&HintLabel> {
        let matches = self.matching_labels();
        if matches.len() == 1 {
            Some(matches[0])
        } else {
            None
        }
    }

    /// Check if there are no matches
    pub fn no_matches(&self) -> bool {
        !self.buffer.is_empty() && self.matching_labels().is_empty()
    }
}

/// Generate hint labels for a given count of cells.
///
/// Labels are generated in sequence:
/// - 0..25 → a..z (single letters)
/// - 26..(26 + 26*26) → aa..zz (two letters)
///
/// This gives us 26 + 676 = 702 unique labels, enough for most visible grids.
pub fn generate_label(index: usize) -> String {
    const ALPHABET: &[u8] = b"abcdefghijklmnopqrstuvwxyz";
    let len = ALPHABET.len();

    if index < len {
        // Single letter: a-z
        String::from(ALPHABET[index] as char)
    } else {
        // Two letters: aa-zz
        let adjusted = index - len;
        let first = adjusted / len;
        let second = adjusted % len;
        if first < len {
            let mut s = String::with_capacity(2);
            s.push(ALPHABET[first] as char);
            s.push(ALPHABET[second] as char);
            s
        } else {
            // Overflow - shouldn't happen with typical grid sizes
            format!("{}", index)
        }
    }
}

/// Generate hint labels for visible cells in row-major order.
///
/// # Arguments
/// * `scroll_row` - First visible row
/// * `scroll_col` - First visible column
/// * `visible_rows` - Number of visible rows
/// * `visible_cols` - Number of visible columns
///
/// # Returns
/// A vector of HintLabels for each visible cell position.
pub fn generate_hints(
    scroll_row: usize,
    scroll_col: usize,
    visible_rows: usize,
    visible_cols: usize,
) -> Vec<HintLabel> {
    let count = visible_rows * visible_cols;
    let mut labels = Vec::with_capacity(count);

    for i in 0..count {
        let local_row = i / visible_cols;
        let local_col = i % visible_cols;
        let row = scroll_row + local_row;
        let col = scroll_col + local_col;
        let label = generate_label(i);
        labels.push(HintLabel { row, col, label });
    }

    labels
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_single_letter_labels() {
        assert_eq!(generate_label(0), "a");
        assert_eq!(generate_label(1), "b");
        assert_eq!(generate_label(25), "z");
    }

    #[test]
    fn test_two_letter_labels() {
        assert_eq!(generate_label(26), "aa");
        assert_eq!(generate_label(27), "ab");
        assert_eq!(generate_label(51), "az");
        assert_eq!(generate_label(52), "ba");
        assert_eq!(generate_label(26 + 26 * 26 - 1), "zz");
    }

    #[test]
    fn test_generate_hints_small_grid() {
        let hints = generate_hints(0, 0, 3, 3);
        assert_eq!(hints.len(), 9);

        // Check positions (row-major)
        assert_eq!((hints[0].row, hints[0].col), (0, 0));
        assert_eq!((hints[1].row, hints[1].col), (0, 1));
        assert_eq!((hints[2].row, hints[2].col), (0, 2));
        assert_eq!((hints[3].row, hints[3].col), (1, 0));

        // Check labels
        assert_eq!(hints[0].label, "a");
        assert_eq!(hints[8].label, "i");
    }

    #[test]
    fn test_generate_hints_with_scroll() {
        let hints = generate_hints(10, 5, 2, 2);
        assert_eq!(hints.len(), 4);

        assert_eq!((hints[0].row, hints[0].col), (10, 5));
        assert_eq!((hints[3].row, hints[3].col), (11, 6));
    }

    #[test]
    fn test_hint_state_matching() {
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 3, 3);

        // No buffer - all match
        assert_eq!(state.matching_labels().len(), 9);

        // With 'a' - only "a" matches
        state.buffer = "a".to_string();
        let matches = state.matching_labels();
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].label, "a");
        assert!(state.unique_match().is_some());
    }

    #[test]
    fn test_hint_state_no_matches() {
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 3, 3);
        state.buffer = "xyz".to_string();

        assert!(state.no_matches());
        assert!(state.unique_match().is_none());
    }

    // ========================================================================
    // Resolver tests
    // ========================================================================

    #[test]
    fn test_gg_command_resolves() {
        // "g" in buffer (after pressing g to enter hint mode) = gg command
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 10, 10);
        state.buffer = "g".to_string();

        let result = resolve_hint_buffer(&state);
        assert!(matches!(result, HintResolution::Command(GCommand::GotoTop)));
    }

    #[test]
    fn test_command_takes_precedence_over_label() {
        // Even though "g" could match cell label index 6, command wins
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 10, 10);
        // Label at index 6 is "g"
        assert_eq!(state.labels[6].label, "g");

        state.buffer = "g".to_string();
        let result = resolve_hint_buffer(&state);

        // Command should win over label
        assert!(matches!(result, HintResolution::Command(GCommand::GotoTop)));
    }

    #[test]
    fn test_label_resolves_when_unique() {
        // Small grid where "b" is unique (no "ba", "bb", etc.)
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 2, 2); // Only a, b, c, d
        state.buffer = "b".to_string();

        let result = resolve_hint_buffer(&state);
        assert!(matches!(result, HintResolution::Jump(0, 1))); // b is at (0, 1)
    }

    #[test]
    fn test_pending_with_multiple_matches() {
        // Larger grid where "a" prefix matches multiple labels (a, aa, ab, etc.)
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 10, 10); // 100 cells = a-z + aa-bv
        state.buffer = "a".to_string();

        // "a" matches "a" and "aa", "ab", "ac", etc. - so pending
        let result = resolve_hint_buffer(&state);
        assert!(matches!(result, HintResolution::Pending));
    }

    #[test]
    fn test_no_match_resolves() {
        let mut state = HintState::default();
        state.labels = generate_hints(0, 0, 3, 3); // Only a-i
        state.buffer = "xyz".to_string();

        let result = resolve_hint_buffer(&state);
        assert!(matches!(result, HintResolution::NoMatch));
    }
}
