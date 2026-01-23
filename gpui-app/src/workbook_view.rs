//! WorkbookView: the binding between a workbook and its view state.
//!
//! This module implements the core abstraction that enables multiple views
//! of the same workbook (tabs, splits, linked scrolling).
//!
//! ## Architecture
//!
//! ```text
//! Workbook ── shared document data (sheets, cells, formulas, named ranges)
//! WorkbookViewState ── per-view state (scroll, selection, zoom, active sheet)
//! WorkbookView ── combines workbook + view state
//! ```
//!
//! A single Workbook can have multiple WorkbookViews. Each view has independent:
//! - Scroll position
//! - Selection (active cell, selection range, additional selections)
//! - Zoom level
//! - Active sheet
//! - Freeze pane configuration
//!
//! Changes to the Workbook (cell edits, formula changes) are visible in all views.
//! View state changes (scrolling, selecting) only affect that specific view.

use visigrid_engine::workbook::Workbook;

/// View-specific state that can differ between views of the same workbook.
///
/// This is the "per-(pane, workbook)" state. When you split a pane, each side
/// gets its own `WorkbookViewState` while sharing the same `Workbook`.
#[derive(Clone, Debug)]
pub struct WorkbookViewState {
    // === Selection ===
    /// Active cell (anchor of the current selection)
    pub selected: (usize, usize),
    /// End of active range selection (None = single cell selected)
    pub selection_end: Option<(usize, usize)>,
    /// Additional selections from Ctrl+Click (each is anchor + optional end)
    pub additional_selections: Vec<((usize, usize), Option<(usize, usize)>)>,

    // === Viewport ===
    /// First visible row (0-indexed)
    pub scroll_row: usize,
    /// First visible column (0-indexed)
    pub scroll_col: usize,

    // === Freeze Panes ===
    /// Number of rows frozen at top (0 = none)
    pub frozen_rows: usize,
    /// Number of columns frozen at left (0 = none)
    pub frozen_cols: usize,

    // === Zoom ===
    /// Zoom level (1.0 = 100%)
    pub zoom_level: f32,

    // === Active Sheet ===
    /// Index of the active sheet in the workbook
    pub active_sheet: usize,
}

impl Default for WorkbookViewState {
    fn default() -> Self {
        Self {
            selected: (0, 0),
            selection_end: None,
            additional_selections: Vec::new(),
            scroll_row: 0,
            scroll_col: 0,
            frozen_rows: 0,
            frozen_cols: 0,
            zoom_level: 1.0,
            active_sheet: 0,
        }
    }
}

impl WorkbookViewState {
    /// Create a new view state with default values
    pub fn new() -> Self {
        Self::default()
    }

    /// Clone this view state for a split operation.
    /// The new view starts with the same scroll/selection as the original.
    pub fn clone_for_split(&self) -> Self {
        self.clone()
    }

    /// Reset selection to a single cell
    pub fn select_cell(&mut self, row: usize, col: usize) {
        self.selected = (row, col);
        self.selection_end = None;
        self.additional_selections.clear();
    }

    /// Get the active cell position
    pub fn active_cell(&self) -> (usize, usize) {
        self.selected
    }

    /// Check if a cell is within the current selection
    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        // Check primary selection
        if self.is_in_primary_selection(row, col) {
            return true;
        }
        // Check additional selections
        for (anchor, end) in &self.additional_selections {
            if Self::is_in_range(row, col, *anchor, *end) {
                return true;
            }
        }
        false
    }

    /// Check if a cell is in the primary selection
    fn is_in_primary_selection(&self, row: usize, col: usize) -> bool {
        Self::is_in_range(row, col, self.selected, self.selection_end)
    }

    /// Check if a cell is within a selection range
    fn is_in_range(
        row: usize,
        col: usize,
        anchor: (usize, usize),
        end: Option<(usize, usize)>,
    ) -> bool {
        let (r1, c1) = anchor;
        let (r2, c2) = end.unwrap_or(anchor);
        let (min_r, max_r) = (r1.min(r2), r1.max(r2));
        let (min_c, max_c) = (c1.min(c2), c1.max(c2));
        row >= min_r && row <= max_r && col >= min_c && col <= max_c
    }

    /// Ensure the active cell is visible by adjusting scroll position
    pub fn ensure_visible(&mut self, visible_rows: usize, visible_cols: usize) {
        let (row, col) = self.selected;

        // Account for frozen panes
        let effective_scroll_row = self.scroll_row + self.frozen_rows;
        let effective_scroll_col = self.scroll_col + self.frozen_cols;

        // Rows
        if row < effective_scroll_row {
            self.scroll_row = row.saturating_sub(self.frozen_rows);
        } else if row >= effective_scroll_row + visible_rows.saturating_sub(self.frozen_rows) {
            self.scroll_row = row.saturating_sub(visible_rows.saturating_sub(1)).saturating_sub(self.frozen_rows);
        }

        // Columns
        if col < effective_scroll_col {
            self.scroll_col = col.saturating_sub(self.frozen_cols);
        } else if col >= effective_scroll_col + visible_cols.saturating_sub(self.frozen_cols) {
            self.scroll_col = col.saturating_sub(visible_cols.saturating_sub(1)).saturating_sub(self.frozen_cols);
        }
    }
}

/// A view of a workbook with independent view state.
///
/// Multiple `WorkbookView`s can reference the same `Workbook` while maintaining
/// separate scroll positions, selections, and zoom levels.
///
/// In the current Phase 1 implementation, the workbook is owned directly.
/// Future phases will use `Entity<Workbook>` for shared ownership across views.
pub struct WorkbookView {
    /// The workbook data (owned for now, Entity<Workbook> in future)
    pub workbook: Workbook,
    /// This view's independent state
    pub state: WorkbookViewState,
}

impl WorkbookView {
    /// Create a new view for a workbook
    pub fn new(workbook: Workbook) -> Self {
        Self {
            workbook,
            state: WorkbookViewState::default(),
        }
    }

    /// Create a new view with specific initial state
    pub fn with_state(workbook: Workbook, state: WorkbookViewState) -> Self {
        Self { workbook, state }
    }

    /// Get the active sheet from the workbook
    pub fn active_sheet(&self) -> Option<&visigrid_engine::sheet::Sheet> {
        self.workbook.sheet(self.state.active_sheet)
    }

    /// Get the active sheet mutably from the workbook
    pub fn active_sheet_mut(&mut self) -> Option<&mut visigrid_engine::sheet::Sheet> {
        self.workbook.sheet_mut(self.state.active_sheet)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_view_state_default() {
        let state = WorkbookViewState::default();
        assert_eq!(state.selected, (0, 0));
        assert_eq!(state.selection_end, None);
        assert!(state.additional_selections.is_empty());
        assert_eq!(state.scroll_row, 0);
        assert_eq!(state.scroll_col, 0);
        assert_eq!(state.zoom_level, 1.0);
    }

    #[test]
    fn test_select_cell() {
        let mut state = WorkbookViewState::default();
        state.selection_end = Some((5, 5));
        state.additional_selections.push(((10, 10), None));

        state.select_cell(3, 4);

        assert_eq!(state.selected, (3, 4));
        assert_eq!(state.selection_end, None);
        assert!(state.additional_selections.is_empty());
    }

    #[test]
    fn test_is_selected_single_cell() {
        let mut state = WorkbookViewState::default();
        state.selected = (5, 5);

        assert!(state.is_selected(5, 5));
        assert!(!state.is_selected(5, 6));
        assert!(!state.is_selected(6, 5));
    }

    #[test]
    fn test_is_selected_range() {
        let mut state = WorkbookViewState::default();
        state.selected = (2, 2);
        state.selection_end = Some((4, 4));

        assert!(state.is_selected(2, 2));
        assert!(state.is_selected(3, 3));
        assert!(state.is_selected(4, 4));
        assert!(state.is_selected(2, 4));
        assert!(!state.is_selected(1, 1));
        assert!(!state.is_selected(5, 5));
    }

    #[test]
    fn test_is_selected_additional_selections() {
        let mut state = WorkbookViewState::default();
        state.selected = (0, 0);
        state.additional_selections.push(((10, 10), Some((12, 12))));

        assert!(state.is_selected(0, 0));
        assert!(state.is_selected(11, 11));
        assert!(!state.is_selected(5, 5));
    }

    #[test]
    fn test_clone_for_split() {
        let mut state = WorkbookViewState::default();
        state.selected = (5, 5);
        state.scroll_row = 100;
        state.zoom_level = 1.5;

        let cloned = state.clone_for_split();

        assert_eq!(cloned.selected, (5, 5));
        assert_eq!(cloned.scroll_row, 100);
        assert_eq!(cloned.zoom_level, 1.5);
    }
}
