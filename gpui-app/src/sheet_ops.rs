//! Sheet operations - freeze panes, navigation, rename, and document identity
//!
//! Contains:
//! - Freeze Panes (freeze/unfreeze rows and columns)
//! - Sheet access convenience methods
//! - Document identity and title bar
//! - Sheet navigation (next/prev/goto/add sheet)
//! - Sheet rename (start/confirm/cancel)
//! - Sheet context menu (show/hide/delete)
//! - Session persistence

use gpui::*;
use visigrid_engine::sheet::Sheet;
use crate::app::{Spreadsheet, display_filename, ext_lower, is_native_ext, DocumentMeta, DocumentSource};
use crate::mode::Mode;
use crate::session::SessionManager;

impl Spreadsheet {
    // =========================================================================
    // Freeze Panes
    // =========================================================================

    /// Freeze the top row (row 0)
    pub fn freeze_top_row(&mut self, cx: &mut Context<Self>) {
        self.view_state.frozen_rows = 1;
        self.view_state.frozen_cols = 0;
        self.clamp_scroll_to_freeze(cx);
        self.status_message = Some("Frozen top row".to_string());
        cx.notify();
    }

    /// Freeze the first column (column A)
    pub fn freeze_first_column(&mut self, cx: &mut Context<Self>) {
        self.view_state.frozen_rows = 0;
        self.view_state.frozen_cols = 1;
        self.clamp_scroll_to_freeze(cx);
        self.status_message = Some("Frozen first column".to_string());
        cx.notify();
    }

    /// Freeze panes at the current selection
    /// Freezes all rows above and all columns to the left of the active cell
    pub fn freeze_panes(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selected;
        if row == 0 && col == 0 {
            // Nothing to freeze - show message
            self.status_message = Some("Select a cell to freeze rows above and columns to the left".to_string());
            cx.notify();
            return;
        }
        self.view_state.frozen_rows = row;
        self.view_state.frozen_cols = col;
        self.clamp_scroll_to_freeze(cx);
        let msg = match (row, col) {
            (0, c) => format!("Frozen {} column{}", c, if c == 1 { "" } else { "s" }),
            (r, 0) => format!("Frozen {} row{}", r, if r == 1 { "" } else { "s" }),
            (r, c) => format!("Frozen {} row{} and {} column{}", r, if r == 1 { "" } else { "s" }, c, if c == 1 { "" } else { "s" }),
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    /// Remove all freeze panes
    pub fn unfreeze_panes(&mut self, cx: &mut Context<Self>) {
        if self.view_state.frozen_rows == 0 && self.view_state.frozen_cols == 0 {
            self.status_message = Some("No frozen panes to unfreeze".to_string());
            cx.notify();
            return;
        }
        self.view_state.frozen_rows = 0;
        self.view_state.frozen_cols = 0;
        self.status_message = Some("Unfrozen all panes".to_string());
        cx.notify();
    }

    /// Clamp scroll position to ensure it doesn't overlap with frozen regions
    fn clamp_scroll_to_freeze(&mut self, _cx: &mut Context<Self>) {
        // When freeze panes are active, scrollable region starts after frozen rows/cols
        // Ensure scroll position doesn't show frozen rows/cols in the scrollable area
        if self.view_state.frozen_rows > 0 && self.view_state.scroll_row < self.view_state.frozen_rows {
            self.view_state.scroll_row = self.view_state.frozen_rows;
        }
        if self.view_state.frozen_cols > 0 && self.view_state.scroll_col < self.view_state.frozen_cols {
            self.view_state.scroll_col = self.view_state.frozen_cols;
        }
    }

    /// Check if freeze panes are active
    pub fn has_frozen_panes(&self) -> bool {
        self.view_state.frozen_rows > 0 || self.view_state.frozen_cols > 0
    }

    /// Save document settings to sidecar if document has a path
    pub(crate) fn save_doc_settings_if_needed(&self) {
        if let Some(ref path) = self.current_file {
            // Best-effort save - don't block on errors
            let _ = crate::settings::save_doc_settings(path, &self.doc_settings);
        }
    }

    // =========================================================================
    // Session persistence
    // =========================================================================

    /// Update the global session with this window's current state.
    /// Called on significant state changes (file open/save, panel toggles).
    pub fn update_session(&self, window: &Window, cx: &mut Context<Self>) {
        let snapshot = self.snapshot(window);
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Update session using cached window bounds (for use without Window access).
    /// Useful from file_ops or other places where Window isn't available.
    pub fn update_session_cached(&self, cx: &mut Context<Self>) {
        let snapshot = self.snapshot_cached();
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Internal: update session with a snapshot
    fn update_session_with_snapshot(&self, snapshot: crate::session::WindowSession, cx: &mut Context<Self>) {
        cx.update_global::<SessionManager, _>(|mgr, _| {
            // Find and update this window's entry, or add a new one
            // For now, we use the file path as the key (simple single-window case)
            let session = mgr.session_mut();

            // Find existing window by file path, or add new
            let idx = session.windows.iter().position(|w| w.file == snapshot.file);

            if let Some(idx) = idx {
                session.windows[idx] = snapshot;
            } else {
                session.windows.push(snapshot);
            }
        });
    }

    /// Save session immediately (for quit/close).
    /// This saves the session to disk synchronously.
    pub fn save_session_now(&self, window: &Window, cx: &mut Context<Self>) {
        self.update_session(window, cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    /// Save session using cached window bounds (for use without Window access).
    pub fn save_session_cached(&self, cx: &mut Context<Self>) {
        self.update_session_cached(cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    // =========================================================================
    // Sheet access convenience methods
    // =========================================================================

    /// Get a reference to the active sheet
    pub fn sheet(&self) -> &Sheet {
        self.workbook.active_sheet()
    }

    /// Get a mutable reference to the active sheet
    pub fn sheet_mut(&mut self) -> &mut Sheet {
        self.workbook.active_sheet_mut()
    }

    /// Get the active sheet index (for undo history)
    pub fn sheet_index(&self) -> usize {
        self.workbook.active_sheet_index()
    }

    /// Set a cell value and update the dependency graph.
    /// This is the preferred way to set cell values - it ensures the dep graph stays in sync.
    pub fn set_cell_value(&mut self, row: usize, col: usize, value: &str) {
        let sheet_id = self.workbook.active_sheet_id();
        self.workbook.active_sheet_mut().set_value(row, col, value);
        self.workbook.update_cell_deps(sheet_id, row, col);
    }

    // =========================================================================
    // Document identity and title bar
    // =========================================================================

    /// Returns true if document has unsaved changes.
    /// Computed from history state, not tracked manually.
    pub fn is_dirty(&self) -> bool {
        self.history.is_dirty()
    }

    /// Update window title if it changed (debounced).
    /// This is the ONLY way titles should update.
    pub fn update_title_if_needed(&mut self, window: &mut Window) {
        let title = self.document_meta.title_string(self.is_dirty());
        if self.cached_title.as_deref() != Some(&title) {
            window.set_window_title(&title);
            self.cached_title = Some(title);
        }
    }

    /// Invalidate the title cache (forces update on next update_title_if_needed call).
    /// Use this when title-affecting state changes but you don't have window access.
    /// Request a title refresh on the next UI pass.
    /// Use when title-affecting state changes but you don't have window access.
    pub fn request_title_refresh(&mut self, cx: &mut Context<Self>) {
        self.cached_title = None;
        self.pending_title_refresh = true;
        cx.notify();
    }

    /// Finalize document state after loading a file
    pub fn finalize_load(&mut self, path: &std::path::Path) {
        let ext = ext_lower(path);
        let filename = display_filename(path);
        let is_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

        // Determine source and saved state based on file type
        let (source, is_saved) = if is_native {
            // Native formats - no provenance, considered "saved"
            (None, true)
        } else {
            // Import formats - show provenance, not "saved" until Save As
            (Some(DocumentSource::Imported { filename: filename.clone() }), false)
        };

        self.document_meta = DocumentMeta {
            display_name: filename,
            is_saved,
            is_read_only: false,
            source,
            path: Some(path.to_path_buf()),
        };

        // Keep current_file in sync (legacy)
        self.current_file = Some(path.to_path_buf());

        // CRITICAL: Set save point AFTER load completes
        // This ensures the document starts "clean" (not dirty)
        self.history.mark_saved();
    }

    /// Finalize document state after saving
    pub fn finalize_save(&mut self, path: &std::path::Path) {
        let ext = ext_lower(path);
        let becomes_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

        self.document_meta.display_name = display_filename(path);
        self.document_meta.path = Some(path.to_path_buf());
        self.history.mark_saved();

        if becomes_native {
            // Saving to native format clears import provenance
            self.document_meta.source = None;
            self.document_meta.is_saved = true;
        }
        // Note: Exporting to CSV/JSON does NOT clear provenance or mark as saved

        // Keep legacy fields in sync
        self.current_file = Some(path.to_path_buf());
        self.is_modified = false;
    }

    // =========================================================================
    // Sheet navigation methods
    // =========================================================================

    /// Move to the next sheet
    pub fn next_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.next_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.prev_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Switch to a specific sheet by index
    pub fn goto_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        // Close validation dropdown when switching sheets
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::SheetSwitch,
            cx,
        );

        // Commit any pending edit before switching sheets
        self.commit_pending_edit();
        if self.workbook.set_active_sheet(index) {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Add a new sheet and switch to it
    pub fn add_sheet(&mut self, cx: &mut Context<Self>) {
        let new_index = self.workbook.add_sheet();
        self.workbook.set_active_sheet(new_index);
        self.clear_selection_state();
        self.is_modified = true;
        cx.notify();
    }

    /// Clear selection state when switching sheets
    fn clear_selection_state(&mut self) {
        self.view_state.selected = (0, 0);
        self.view_state.selection_end = None;
        self.view_state.scroll_row = 0;
        self.view_state.scroll_col = 0;
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
    }

    // =========================================================================
    // Sheet rename methods
    // =========================================================================

    /// Start renaming a sheet (double-click on tab)
    pub fn start_sheet_rename(&mut self, index: usize, cx: &mut Context<Self>) {
        if let Some(name) = self.workbook.sheet_names().get(index) {
            self.renaming_sheet = Some(index);
            self.sheet_rename_input = name.to_string();
            self.sheet_context_menu = None;
            cx.notify();
        }
    }

    /// Confirm the sheet rename
    pub fn confirm_sheet_rename(&mut self, cx: &mut Context<Self>) {
        if let Some(index) = self.renaming_sheet {
            let new_name = self.sheet_rename_input.trim();
            if !new_name.is_empty() {
                self.workbook.rename_sheet(index, new_name);
                self.is_modified = true;
            }
            self.renaming_sheet = None;
            self.sheet_rename_input.clear();
            self.request_title_refresh(cx);
        }
    }

    /// Cancel the sheet rename
    pub fn cancel_sheet_rename(&mut self, cx: &mut Context<Self>) {
        self.renaming_sheet = None;
        self.sheet_rename_input.clear();
        cx.notify();
    }

    /// Handle input for sheet rename
    pub fn sheet_rename_input_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.push(c);
            cx.notify();
        }
    }

    /// Handle backspace for sheet rename
    pub fn sheet_rename_backspace(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.pop();
            cx.notify();
        }
    }

    // =========================================================================
    // Sheet context menu methods
    // =========================================================================

    /// Show context menu for a sheet tab
    pub fn show_sheet_context_menu(&mut self, index: usize, cx: &mut Context<Self>) {
        self.sheet_context_menu = Some(index);
        self.renaming_sheet = None;
        cx.notify();
    }

    /// Hide sheet context menu
    pub fn hide_sheet_context_menu(&mut self, cx: &mut Context<Self>) {
        self.sheet_context_menu = None;
        cx.notify();
    }

    /// Delete a sheet
    pub fn delete_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        if self.workbook.delete_sheet(index) {
            self.is_modified = true;
            self.sheet_context_menu = None;
            self.request_title_refresh(cx);
        } else {
            self.status_message = Some("Cannot delete the last sheet".to_string());
            self.sheet_context_menu = None;
            cx.notify();
        }
    }
}
