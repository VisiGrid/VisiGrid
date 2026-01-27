//! Split View functionality for side-by-side panes.
//!
//! ## Overview
//!
//! Split view allows viewing the same workbook in two side-by-side panes.
//! This is useful for comparing distant parts of a sheet, editing while
//! viewing results, or tracing formula dependencies.
//!
//! ## Architecture Invariants
//!
//! **These invariants MUST be maintained. Violating them breaks split view.**
//!
//! 1. **Workbook data is SHARED via `Entity<Workbook>`**
//!    - Both panes see the same cells, formulas, and formatting
//!    - Edits in one pane are immediately visible in the other
//!    - Undo/redo operates on the shared workbook
//!
//! 2. **View state is NEVER shared**
//!    - Each pane has its own `WorkbookViewState` (scroll, selection, zoom)
//!    - Left pane: `self.view_state`
//!    - Right pane: `self.split_pane.view_state`
//!
//! 3. **All mutations route through `active_view_state_mut()`**
//!    - Selection changes → `active_view_state_mut().selected = ...`
//!    - Scroll changes → `active_view_state_mut().scroll_row = ...`
//!    - NEVER write directly to `self.view_state` when split is active
//!
//! 4. **Rendering is pane-aware, logic is pane-agnostic**
//!    - `render_grid(app, window, cx, Some(SplitSide::Left))` for left pane
//!    - `render_grid(app, window, cx, Some(SplitSide::Right))` for right pane
//!    - Business logic doesn't know which pane is active; it just uses
//!      `active_view_state_mut()` and the routing is handled automatically
//!
//! 5. **Click activates pane before any other action**
//!    - Mouse handlers call `activate_pane(pane_side, cx)` first
//!    - This ensures subsequent mutations go to the clicked pane
//!
//! ## Keybindings
//!
//! - `Ctrl+\` - Split View (creates right pane or toggles active pane)
//! - `Ctrl+Shift+\` - Close Split (keeps active pane's state)
//! - `Ctrl+]` - Focus Other Pane
//!
//! ## Current Limitations (intentional for MVP)
//!
//! - Fixed 50/50 split (no draggable divider)
//! - Horizontal split only (no vertical)
//! - Maximum 2 panes (no nested splits)
//! - Single formula bar (shows active pane's cell)

use gpui::*;
use crate::app::Spreadsheet;
use crate::workbook_view::WorkbookViewState;

/// Which side of a split is active
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum SplitSide {
    #[default]
    Left,
    Right,
}

impl SplitSide {
    pub fn opposite(self) -> Self {
        match self {
            SplitSide::Left => SplitSide::Right,
            SplitSide::Right => SplitSide::Left,
        }
    }
}

/// State for the secondary (right) pane in a split view
#[derive(Clone)]
pub struct SplitPane {
    /// View state for the right pane (selection, scroll, etc.)
    pub view_state: WorkbookViewState,
    /// Focus handle for the right pane
    pub focus_handle: FocusHandle,
}

impl SplitPane {
    /// Create a new split pane, cloning state from the original
    pub fn new(original_state: &WorkbookViewState, cx: &mut App) -> Self {
        Self {
            view_state: original_state.clone_for_split(),
            focus_handle: cx.focus_handle(),
        }
    }
}

impl Spreadsheet {
    // =========================================================================
    // Split View Operations
    // =========================================================================

    /// Check if split view is active
    pub fn is_split(&self) -> bool {
        self.split_pane.is_some()
    }

    /// Get the view state for the currently active pane
    pub fn active_view_state(&self) -> &WorkbookViewState {
        match self.split_active_side {
            SplitSide::Left => &self.view_state,
            SplitSide::Right => self.split_pane.as_ref()
                .map(|p| &p.view_state)
                .unwrap_or(&self.view_state),
        }
    }

    /// Get mutable view state for the currently active pane
    pub fn active_view_state_mut(&mut self) -> &mut WorkbookViewState {
        match self.split_active_side {
            SplitSide::Left => &mut self.view_state,
            SplitSide::Right => self.split_pane.as_mut()
                .map(|p| &mut p.view_state)
                .unwrap_or(&mut self.view_state),
        }
    }

    /// Get the focus handle for the currently active pane
    pub fn active_focus_handle(&self) -> &FocusHandle {
        match self.split_active_side {
            SplitSide::Left => &self.focus_handle,
            SplitSide::Right => self.split_pane.as_ref()
                .map(|p| &p.focus_handle)
                .unwrap_or(&self.focus_handle),
        }
    }

    /// Split the view (Ctrl+\)
    /// Creates a new right pane with cloned state from the current view
    pub fn split_right(&mut self, cx: &mut Context<Self>) {
        if self.split_pane.is_some() {
            // Already split - just switch to the other pane
            self.split_active_side = self.split_active_side.opposite();
            cx.notify();
            return;
        }

        // Create new right pane with cloned state
        self.split_pane = Some(SplitPane::new(&self.view_state, cx));
        self.split_active_side = SplitSide::Right;
        self.status_message = Some("Split view enabled".to_string());
        cx.notify();
    }

    /// Close the split view, keeping the active pane's state
    pub fn close_split(&mut self, cx: &mut Context<Self>) {
        if self.split_pane.is_none() {
            return;
        }

        // If right pane is active, adopt its state
        if self.split_active_side == SplitSide::Right {
            if let Some(pane) = &self.split_pane {
                self.view_state = pane.view_state.clone();
            }
        }

        self.split_pane = None;
        self.split_active_side = SplitSide::Left;
        self.status_message = Some("Split view closed".to_string());
        cx.notify();
    }

    /// Focus the other pane (cycles between panes when split)
    pub fn focus_other_pane(&mut self, cx: &mut Context<Self>) {
        if self.split_pane.is_none() {
            return;
        }

        self.split_active_side = self.split_active_side.opposite();
        cx.notify();
    }

    /// Activate a specific pane (called when clicking in a pane).
    /// If pane_side is None or we're not split, this is a no-op.
    pub fn activate_pane(&mut self, pane_side: Option<SplitSide>, cx: &mut Context<Self>) {
        if self.split_pane.is_none() {
            return; // Not split, nothing to do
        }
        if let Some(side) = pane_side {
            if self.split_active_side != side {
                self.split_active_side = side;
                cx.notify();
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::SplitSide;

    #[test]
    fn test_split_side_opposite() {
        assert_eq!(SplitSide::Left.opposite(), SplitSide::Right);
        assert_eq!(SplitSide::Right.opposite(), SplitSide::Left);
    }
}
