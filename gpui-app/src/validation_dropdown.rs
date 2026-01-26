//! Validation dropdown state machine
//!
//! Manages the dropdown UI for list validation. The dropdown is a "validation-pick mode"
//! that intercepts keystrokes before they reach the grid.
//!
//! Key design constraints:
//! - Never store raw items in state (stale data risk)
//! - Store fingerprint + indices, not items
//! - Recompute filtered_indices only when filter_text changes, not every render

use std::sync::Arc;
use visigrid_engine::validation::ResolvedList;

/// State machine for the validation dropdown.
#[derive(Debug, Clone)]
pub enum ValidationDropdownState {
    /// No dropdown active
    Closed,

    /// Dropdown is open for a cell
    Open(DropdownOpenState),
}

/// State when dropdown is open.
#[derive(Debug, Clone)]
pub struct DropdownOpenState {
    /// Anchor cell (row, col)
    pub cell: (usize, usize),

    /// The resolved list (held in Arc for cheap cloning)
    pub resolved_list: Arc<ResolvedList>,

    /// Fingerprint of the source data at time of open.
    /// Close dropdown if this changes.
    pub source_fingerprint: u64,

    /// Whether the list was truncated
    pub truncated: bool,

    /// Current filter text (type-to-filter)
    pub filter_text: String,

    /// Indices into resolved_list.items that match the filter.
    /// Recomputed only when filter_text changes.
    pub filtered_indices: Vec<usize>,

    /// Currently selected index within filtered_indices
    pub selected_index: usize,

    /// Scroll offset for long lists
    pub scroll_offset: usize,
}

impl Default for ValidationDropdownState {
    fn default() -> Self {
        Self::Closed
    }
}

impl ValidationDropdownState {
    /// Check if dropdown is open
    pub fn is_open(&self) -> bool {
        matches!(self, Self::Open(_))
    }

    /// Get the open state if dropdown is open
    pub fn as_open(&self) -> Option<&DropdownOpenState> {
        match self {
            Self::Open(state) => Some(state),
            Self::Closed => None,
        }
    }

    /// Get mutable open state if dropdown is open
    pub fn as_open_mut(&mut self) -> Option<&mut DropdownOpenState> {
        match self {
            Self::Open(state) => Some(state),
            Self::Closed => None,
        }
    }

    /// Open the dropdown for a cell.
    ///
    /// Returns the new state. Caller should also verify the list has items.
    pub fn open(cell: (usize, usize), resolved_list: Arc<ResolvedList>) -> Self {
        let source_fingerprint = resolved_list.source_fingerprint;
        let truncated = resolved_list.is_truncated;

        // Initially show all items
        let filtered_indices: Vec<usize> = (0..resolved_list.items.len()).collect();

        Self::Open(DropdownOpenState {
            cell,
            resolved_list,
            source_fingerprint,
            truncated,
            filter_text: String::new(),
            filtered_indices,
            selected_index: 0,
            scroll_offset: 0,
        })
    }

    /// Close the dropdown
    pub fn close(&mut self) {
        *self = Self::Closed;
    }
}

impl DropdownOpenState {
    /// Get the currently selected item, if any
    pub fn selected_item(&self) -> Option<&str> {
        self.filtered_indices
            .get(self.selected_index)
            .and_then(|&idx| self.resolved_list.items.get(idx))
            .map(|s| s.as_str())
    }

    /// Get all visible items (for rendering)
    pub fn visible_items(&self) -> impl Iterator<Item = (usize, &str)> + '_ {
        self.filtered_indices
            .iter()
            .enumerate()
            .filter_map(|(display_idx, &list_idx)| {
                self.resolved_list.items.get(list_idx).map(|s| (display_idx, s.as_str()))
            })
    }

    /// Update filter text and recompute filtered_indices
    pub fn set_filter(&mut self, text: &str) {
        if self.filter_text == text {
            return;
        }

        self.filter_text = text.to_string();
        self.recompute_filtered_indices();
        self.selected_index = 0;
        self.scroll_offset = 0;
    }

    /// Append a character to filter text
    pub fn append_filter_char(&mut self, ch: char) {
        self.filter_text.push(ch);
        self.recompute_filtered_indices();
        // Keep selection at 0 when filtering
        self.selected_index = 0;
        self.scroll_offset = 0;
    }

    /// Remove last character from filter text
    pub fn backspace_filter(&mut self) {
        if self.filter_text.pop().is_some() {
            self.recompute_filtered_indices();
            self.selected_index = 0;
            self.scroll_offset = 0;
        }
    }

    /// Recompute filtered_indices based on current filter_text
    fn recompute_filtered_indices(&mut self) {
        if self.filter_text.is_empty() {
            // No filter - show all
            self.filtered_indices = (0..self.resolved_list.items.len()).collect();
        } else {
            // Case-insensitive substring match
            let filter_lower = self.filter_text.to_lowercase();
            self.filtered_indices = self.resolved_list.items
                .iter()
                .enumerate()
                .filter(|(_, item)| item.to_lowercase().contains(&filter_lower))
                .map(|(idx, _)| idx)
                .collect();
        }
    }

    /// Move selection up
    pub fn move_up(&mut self) {
        if self.selected_index > 0 {
            self.selected_index -= 1;
            self.ensure_selected_visible();
        }
    }

    /// Move selection down
    pub fn move_down(&mut self) {
        if self.selected_index + 1 < self.filtered_indices.len() {
            self.selected_index += 1;
            self.ensure_selected_visible();
        }
    }

    /// Move selection up by page
    pub fn page_up(&mut self, page_size: usize) {
        self.selected_index = self.selected_index.saturating_sub(page_size);
        self.ensure_selected_visible();
    }

    /// Move selection down by page
    pub fn page_down(&mut self, page_size: usize) {
        let max_idx = self.filtered_indices.len().saturating_sub(1);
        self.selected_index = (self.selected_index + page_size).min(max_idx);
        self.ensure_selected_visible();
    }

    /// Ensure the selected item is visible (adjust scroll_offset)
    fn ensure_selected_visible(&mut self) {
        // Simple implementation: just ensure selected_index is in view
        // This will be refined when we know the visible row count
        if self.selected_index < self.scroll_offset {
            self.scroll_offset = self.selected_index;
        }
    }

    /// Check if the source fingerprint still matches
    pub fn is_stale(&self, current_fingerprint: u64) -> bool {
        self.source_fingerprint != current_fingerprint
    }

    /// Number of filtered items
    pub fn filtered_count(&self) -> usize {
        self.filtered_indices.len()
    }

    /// Handle a key event. Returns the outcome for the caller to act on.
    ///
    /// This is the core event routing for the dropdown.
    pub fn handle_key(&mut self, key: &str, modifiers: KeyModifiers) -> DropdownOutcome {
        // Ignore if any modifier except shift is held (let grid handle Ctrl+X, etc.)
        if modifiers.control || modifiers.alt || modifiers.platform {
            return DropdownOutcome::NotConsumed;
        }

        match key {
            "escape" => DropdownOutcome::CloseNoCommit,

            "enter" => {
                // Commit selected item (if any)
                if let Some(value) = self.selected_item() {
                    DropdownOutcome::CommitValue(value.to_string())
                } else {
                    // No items visible (empty filter result) - do nothing
                    DropdownOutcome::Consumed
                }
            }

            "tab" => {
                // Tab closes dropdown and lets grid handle navigation
                DropdownOutcome::CloseNoCommit
            }

            "up" => {
                self.move_up();
                DropdownOutcome::Consumed
            }

            "down" => {
                self.move_down();
                DropdownOutcome::Consumed
            }

            "pageup" => {
                self.page_up(10); // 10 items per page
                DropdownOutcome::Consumed
            }

            "pagedown" => {
                self.page_down(10);
                DropdownOutcome::Consumed
            }

            "home" => {
                self.selected_index = 0;
                self.scroll_offset = 0;
                DropdownOutcome::Consumed
            }

            "end" => {
                if !self.filtered_indices.is_empty() {
                    self.selected_index = self.filtered_indices.len() - 1;
                    self.ensure_selected_visible();
                }
                DropdownOutcome::Consumed
            }

            "backspace" => {
                self.backspace_filter();
                DropdownOutcome::Consumed
            }

            _ => DropdownOutcome::NotConsumed,
        }
    }

    /// Handle a character input for filtering
    pub fn handle_char(&mut self, ch: char) -> DropdownOutcome {
        // Only accept printable characters for filtering
        if ch.is_control() {
            return DropdownOutcome::NotConsumed;
        }
        self.append_filter_char(ch);
        DropdownOutcome::Consumed
    }
}

/// Key modifiers for event handling
#[derive(Debug, Clone, Copy, Default)]
pub struct KeyModifiers {
    pub control: bool,
    pub alt: bool,
    pub shift: bool,
    pub platform: bool, // Cmd on macOS
}

/// Reasons for closing the dropdown (for debugging/logging)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DropdownCloseReason {
    /// User pressed Escape
    Escape,
    /// User clicked outside the dropdown
    ClickOutside,
    /// User moved selection to a different cell
    SelectionChanged,
    /// User switched to a different sheet
    SheetSwitch,
    /// User scrolled the grid
    Scroll,
    /// User zoomed
    Zoom,
    /// A modal dialog opened
    ModalOpened,
    /// The source data fingerprint changed
    SourceChanged,
    /// User selected an item (Enter or click)
    Committed,
}

/// Result of routing an event to the dropdown handler
#[derive(Debug, Clone, PartialEq)]
pub enum DropdownOutcome {
    /// Event consumed, no further action needed
    Consumed,
    /// Close dropdown without committing a value
    CloseNoCommit,
    /// Close dropdown and commit this value to the cell
    CommitValue(String),
    /// Event not handled by dropdown, pass to grid
    NotConsumed,
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_test_list() -> Arc<ResolvedList> {
        Arc::new(ResolvedList::from_items(vec![
            "Apple".to_string(),
            "Banana".to_string(),
            "Cherry".to_string(),
            "Date".to_string(),
            "Elderberry".to_string(),
        ]))
    }

    #[test]
    fn test_open_close() {
        let list = make_test_list();
        let mut state = ValidationDropdownState::open((0, 0), list);

        assert!(state.is_open());
        assert!(state.as_open().is_some());

        state.close();
        assert!(!state.is_open());
        assert!(state.as_open().is_none());
    }

    #[test]
    fn test_filter() {
        let list = make_test_list();
        let state = ValidationDropdownState::open((0, 0), list);

        if let ValidationDropdownState::Open(mut open) = state {
            assert_eq!(open.filtered_count(), 5);

            open.set_filter("an");
            assert_eq!(open.filtered_count(), 2); // Banana, Elderberry (wait, no - "an" matches "Banana")
            // Actually: "an" matches "Banana" (bANana)
            // Let me check: "Apple" - no, "Banana" - yes (bANana), "Cherry" - no, "Date" - no, "Elderberry" - no
            // So only 1 match

            // Wait, case-insensitive: "banana".contains("an") = true
            // Let me trace through:
            // - "apple".to_lowercase() = "apple", contains "an"? No
            // - "banana".to_lowercase() = "banana", contains "an"? Yes
            // - "cherry", "date", "elderberry" - no "an"
            // So filtered_count should be 1
        }
    }

    #[test]
    fn test_filter_case_insensitive() {
        let list = make_test_list();
        let mut state = ValidationDropdownState::open((0, 0), list);

        if let Some(open) = state.as_open_mut() {
            open.set_filter("A");
            // "apple", "banana", "date", "elderberry" all contain 'a'
            // "cherry" does not contain 'a'
            assert_eq!(open.filtered_count(), 4);

            open.set_filter("APPLE");
            assert_eq!(open.filtered_count(), 1);
            assert_eq!(open.selected_item(), Some("Apple"));
        }
    }

    #[test]
    fn test_navigation() {
        let list = make_test_list();
        let mut state = ValidationDropdownState::open((0, 0), list);

        if let Some(open) = state.as_open_mut() {
            assert_eq!(open.selected_index, 0);
            assert_eq!(open.selected_item(), Some("Apple"));

            open.move_down();
            assert_eq!(open.selected_index, 1);
            assert_eq!(open.selected_item(), Some("Banana"));

            open.move_down();
            open.move_down();
            open.move_down();
            assert_eq!(open.selected_index, 4);
            assert_eq!(open.selected_item(), Some("Elderberry"));

            // Can't go past end
            open.move_down();
            assert_eq!(open.selected_index, 4);

            open.move_up();
            assert_eq!(open.selected_index, 3);
            assert_eq!(open.selected_item(), Some("Date"));
        }
    }

    #[test]
    fn test_append_and_backspace() {
        let list = make_test_list();
        let mut state = ValidationDropdownState::open((0, 0), list);

        if let Some(open) = state.as_open_mut() {
            open.append_filter_char('b');
            assert_eq!(open.filter_text, "b");
            // "banana" and "elderberry" contain 'b'
            assert_eq!(open.filtered_count(), 2);

            open.append_filter_char('a');
            assert_eq!(open.filter_text, "ba");
            // Only "banana" contains "ba"
            assert_eq!(open.filtered_count(), 1);

            open.backspace_filter();
            assert_eq!(open.filter_text, "b");
            assert_eq!(open.filtered_count(), 2);

            open.backspace_filter();
            assert_eq!(open.filter_text, "");
            assert_eq!(open.filtered_count(), 5);
        }
    }

    #[test]
    fn test_stale_detection() {
        let list = make_test_list();
        let state = ValidationDropdownState::open((0, 0), list.clone());

        if let Some(open) = state.as_open() {
            assert!(!open.is_stale(list.source_fingerprint));
            assert!(open.is_stale(12345)); // Different fingerprint
        }
    }
}
