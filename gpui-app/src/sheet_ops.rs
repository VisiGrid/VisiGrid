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
use visigrid_engine::workbook::Workbook;
use crate::app::{Spreadsheet, display_filename, ext_lower, is_native_ext, DocumentMeta, DocumentSource};
use crate::mode::Mode;
use crate::session::SessionManager;
use crate::window_registry::{WindowInfo, WindowRegistry};

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
    pub fn update_session(&mut self, window: &Window, cx: &mut Context<Self>) {
        let snapshot = self.snapshot(window, cx);
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Update session using cached window bounds (for use without Window access).
    /// Useful from file_ops or other places where Window isn't available.
    pub fn update_session_cached(&mut self, cx: &mut Context<Self>) {
        let snapshot = self.snapshot_cached(cx);
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Internal: update session with a snapshot.
    /// Self-heals if session_window_id was never assigned: allocates one on the spot.
    fn update_session_with_snapshot(&mut self, mut snapshot: crate::session::WindowSession, cx: &mut Context<Self>) {
        debug_assert!(
            self.session_window_id != crate::app::WINDOW_ID_UNSET,
            "update_session called before session_window_id was assigned"
        );

        // Self-heal: allocate an ID if one was never assigned
        if snapshot.window_id == crate::app::WINDOW_ID_UNSET {
            let id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
            self.session_window_id = id;
            snapshot.window_id = id;
        }

        cx.update_global::<SessionManager, _>(|mgr, _| {
            let session = mgr.session_mut();

            // Match by window_id (stable within a session)
            let idx = session.windows.iter().position(|w| w.window_id == snapshot.window_id);

            if let Some(idx) = idx {
                session.windows[idx] = snapshot;
            } else {
                session.windows.push(snapshot);
            }
        });
    }

    /// Save session immediately (for quit/close).
    /// This saves the session to disk synchronously.
    pub fn save_session_now(&mut self, window: &Window, cx: &mut Context<Self>) {
        self.update_session(window, cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    /// Save session using cached window bounds (for use without Window access).
    pub fn save_session_cached(&mut self, cx: &mut Context<Self>) {
        self.update_session_cached(cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    // =========================================================================
    // Sheet access convenience methods
    // =========================================================================

    /// Get a reference to the active sheet (preview-aware).
    /// Returns the snapshot's sheet during preview, live sheet otherwise.
    /// Pass &**cx from Context, or &app directly.
    pub fn sheet<'a>(&'a self, cx: &'a App) -> &'a Sheet {
        self.display_workbook(cx).active_sheet()
    }

    /// Get the active sheet index (for undo history)
    /// Pass &**cx from Context, or &app directly.
    pub fn sheet_index(&self, cx: &App) -> usize {
        self.wb(cx).active_sheet_index()
    }

    /// Get the role for a cell from metadata (for role-based auto-styling)
    pub fn get_cell_role(&self, row: usize, col: usize) -> Option<crate::role_styles::Role> {
        crate::role_styles::get_cell_role(&self.cell_metadata, row, col)
    }

    /// Get the style for a cell's role
    pub fn get_cell_role_style(&self, row: usize, col: usize) -> Option<&crate::role_styles::RoleStyle> {
        self.get_cell_role(row, col)
            .and_then(|role| self.role_style_map.get(role))
    }

    /// Set a cell value and update the dependency graph.
    /// This is the preferred way to set cell values - it ensures the dep graph stays in sync.
    pub fn set_cell_value(&mut self, row: usize, col: usize, value: &str, cx: &mut Context<Self>) {
        self.workbook.update(cx, |wb, _| {
            let sheet_id = wb.active_sheet_id();
            wb.active_sheet_mut().set_value(row, col, value);
            wb.update_cell_deps(sheet_id, row, col);
            let cell_id = visigrid_engine::cell_id::CellId::new(sheet_id, row, col);
            wb.note_cell_changed(cell_id);
        });
        cx.notify(); // Ensure view re-renders with updated cross-sheet values
    }

    /// Clear a cell value on the active sheet and update the dependency graph + recalc.
    /// This is the preferred way to clear cells - it ensures the dep graph stays in sync.
    pub fn clear_cell_value(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.workbook.update(cx, |wb, _| {
            let sheet_id = wb.active_sheet_id();
            wb.active_sheet_mut().clear_cell(row, col);
            wb.update_cell_deps(sheet_id, row, col);
            let cell_id = visigrid_engine::cell_id::CellId::new(sheet_id, row, col);
            wb.note_cell_changed(cell_id);
        });
        cx.notify(); // Ensure view re-renders with updated cross-sheet values
    }

    /// Execute a mutation on the workbook's active sheet.
    /// Use this instead of sheet_mut() for proper Entity semantics.
    pub fn with_active_sheet_mut<R>(
        &mut self,
        cx: &mut Context<Self>,
        f: impl FnOnce(&mut Sheet) -> R,
    ) -> R {
        self.workbook.update(cx, |wb, _| {
            f(wb.active_sheet_mut())
        })
    }

    /// Execute a mutation on a specific sheet by index.
    pub fn with_sheet_mut<R>(
        &mut self,
        sheet_idx: usize,
        cx: &mut Context<Self>,
        f: impl FnOnce(&mut Sheet) -> R,
    ) -> Option<R> {
        self.workbook.update(cx, |wb, _| {
            wb.sheet_mut(sheet_idx).map(f)
        })
    }

    /// Read from the workbook (returns reference valid during borrow).
    /// For more complex reads, use self.wb(cx) directly.
    pub fn with_workbook<R>(&self, cx: &App, f: impl FnOnce(&Workbook) -> R) -> R {
        f(self.wb(cx))
    }

    /// Update the workbook (for mutations that need full workbook access).
    pub fn update_workbook<R>(
        &mut self,
        cx: &mut Context<Self>,
        f: impl FnOnce(&mut Workbook, &mut Context<Workbook>) -> R,
    ) -> R {
        self.workbook.update(cx, f)
    }

    // =========================================================================
    // Shorthand Workbook Accessors (minimize boilerplate)
    // =========================================================================

    /// Shorthand for read-only workbook access: `self.wb(cx).method()`
    /// Pass &*cx from Context, or &app directly. Context auto-derefs to App.
    #[inline]
    pub fn wb<'a>(&'a self, cx: &'a App) -> &'a Workbook {
        self.workbook.read(cx)
    }

    /// Shorthand for workbook mutation: `self.wb_mut(cx, |wb| wb.method())`
    #[inline]
    pub fn wb_mut<R>(&mut self, cx: &mut Context<Self>, f: impl FnOnce(&mut Workbook) -> R) -> R {
        self.workbook.update(cx, |wb, _| f(wb))
    }

    /// Shorthand for sheet mutation by index: `self.sheet_mut(idx, cx, |s| s.method())`
    #[inline]
    pub fn sheet_mut<R>(&mut self, sheet_idx: usize, cx: &mut Context<Self>, f: impl FnOnce(&mut Sheet) -> R) -> Option<R> {
        self.workbook.update(cx, |wb, _| wb.sheet_mut(sheet_idx).map(f))
    }

    /// Shorthand for active sheet mutation: `self.active_sheet_mut(cx, |s| s.method())`
    #[inline]
    pub fn active_sheet_mut<R>(&mut self, cx: &mut Context<Self>, f: impl FnOnce(&mut Sheet) -> R) -> R {
        self.workbook.update(cx, |wb, _| f(wb.active_sheet_mut()))
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
    /// Also updates the global window registry for the window switcher.
    pub fn update_title_if_needed(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let title = self.document_meta.title_string(self.is_dirty());
        if self.cached_title.as_deref() != Some(&title) {
            window.set_window_title(&title);
            self.cached_title = Some(title.clone());

            // Update the window registry for the window switcher
            self.sync_window_registry(cx);
        }
    }

    /// Sync this window's info to the global registry
    pub fn sync_window_registry(&self, cx: &mut Context<Self>) {
        let display_title = self.document_meta.display_name.clone();

        cx.update_global::<WindowRegistry, _>(|registry, _| {
            registry.update(
                self.window_handle,
                display_title,
                self.is_dirty(),
                self.current_file.clone(),
            );
        });
    }

    /// Register this window with the global registry (call once on creation)
    pub fn register_with_window_registry(&self, cx: &mut Context<Self>) {
        let display_title = self.document_meta.display_name.clone();

        cx.update_global::<WindowRegistry, _>(|registry, _| {
            registry.register(WindowInfo::new(
                self.window_handle,
                display_title,
                self.is_dirty(),
                self.current_file.clone(),
            ));
        });
    }

    /// Remove this window from session state.
    /// Matches by window_id (stable within a session).
    /// Self-heals if ID was never assigned (assigns one, but nothing to remove).
    pub fn remove_from_session(&mut self, cx: &mut Context<Self>) {
        debug_assert!(
            self.session_window_id != crate::app::WINDOW_ID_UNSET,
            "remove_from_session called before session_window_id was assigned"
        );

        // Self-heal: if ID was never assigned, this window was never tracked in session
        if self.session_window_id == crate::app::WINDOW_ID_UNSET {
            self.session_window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
            return; // Nothing in session to remove
        }

        let window_id = self.session_window_id;
        cx.update_global::<SessionManager, _>(|mgr, _| {
            let session = mgr.session_mut();
            if let Some(idx) = session.windows.iter().position(|w| w.window_id == window_id) {
                session.windows.remove(idx);
            }
        });
    }

    /// Prepare this window for closing: remove from session, unregister from registry, persist.
    /// Call before `window.remove_window()`.
    pub fn prepare_close(&mut self, cx: &mut Context<Self>) {
        self.remove_from_session(cx);
        self.unregister_from_window_registry(cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    /// Unregister this window from the global registry (call on close)
    pub fn unregister_from_window_registry(&self, cx: &mut Context<Self>) {
        cx.update_global::<WindowRegistry, _>(|registry, _| {
            registry.unregister(self.window_handle);
        });
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
        if self.wb_mut(cx, |wb| wb.next_sheet()) {
            self.update_cached_sheet_id(cx);
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self, cx: &mut Context<Self>) {
        if self.wb_mut(cx, |wb| wb.prev_sheet()) {
            self.update_cached_sheet_id(cx);
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Switch to a specific sheet by index
    pub fn goto_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        // In formula mode, switch sheets for cross-sheet reference picking
        // without committing the formula or clearing edit state.
        if self.mode.is_formula() {
            if self.wb_mut(cx, |wb| wb.set_active_sheet(index)) {
                self.update_cached_sheet_id(cx);
                // Reset scroll/selection on the target sheet but stay in formula mode
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                // Track which sheet the ref target is on
                let home = self.formula_home_sheet.unwrap_or(0);
                if index != home {
                    self.formula_ref_sheet = Some(index);
                    // Cache the sheet name for reference text insertion
                    let name = self.wb(cx).sheet_names().get(index)
                        .map(|s| s.to_string()).unwrap_or_default();
                    self.formula_cross_sheet_name = Some(name);
                } else {
                    self.formula_ref_sheet = None;
                    self.formula_cross_sheet_name = None;
                }
                // Clear any in-progress ref when switching sheets
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
                // Hide autocomplete on cross-sheet navigation
                self.autocomplete_visible = false;
                self.autocomplete_suppressed = true;
                cx.notify();
            }
            return;
        }

        // Close validation dropdown when switching sheets
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::SheetSwitch,
            cx,
        );

        // Commit any pending edit before switching sheets
        self.commit_pending_edit(cx);
        if self.wb_mut(cx, |wb| wb.set_active_sheet(index)) {
            self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
            self.debug_assert_sheet_cache_sync(cx);  // Catch desync immediately at switch point
            self.clear_selection_state();
            // Clear history highlight unless it's for the new sheet
            if let Some((sheet_idx, _, _, _, _)) = self.history_highlight_range {
                if sheet_idx != index {
                    self.history_highlight_range = None;
                }
            }
            cx.notify();
        }
    }

    /// Add a new sheet and switch to it
    pub fn add_sheet(&mut self, cx: &mut Context<Self>) {
        let new_index = self.wb_mut(cx, |wb| wb.add_sheet());
        self.wb_mut(cx, |wb| wb.set_active_sheet(new_index));
        self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
        self.debug_assert_sheet_cache_sync(cx);  // Catch desync immediately
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
        self.tab_chain_origin_col = None;  // Sheet switch breaks tab chain
    }

    // =========================================================================
    // Sheet rename methods
    // =========================================================================

    /// Start renaming a sheet (double-click on tab or context menu)
    pub fn start_sheet_rename(&mut self, index: usize, cx: &mut Context<Self>) {
        if let Some(name) = self.wb(cx).sheet_names().get(index).map(|s| s.to_string()) {
            self.renaming_sheet = Some(index);
            self.sheet_rename_input = name;
            self.sheet_rename_cursor = self.sheet_rename_input.len();
            self.sheet_rename_select_all = true;  // Select all on start
            self.sheet_context_menu = None;
            self.start_caret_blink(cx);
            cx.notify();
        }
    }

    /// Confirm the sheet rename with validation.
    /// Rejects: empty names, duplicates, too long. Trims whitespace.
    pub fn confirm_sheet_rename(&mut self, cx: &mut Context<Self>) {
        if let Some(index) = self.renaming_sheet {
            let new_name = self.sheet_rename_input.trim();

            // Helper to reset rename state
            let reset_state = |this: &mut Self| {
                this.renaming_sheet = None;
                this.sheet_rename_input.clear();
                this.sheet_rename_cursor = 0;
                this.sheet_rename_select_all = false;
                this.stop_caret_blink();
            };

            // Validation: reject empty names
            if new_name.is_empty() {
                self.status_message = Some("Sheet name cannot be empty".to_string());
                reset_state(self);
                cx.notify();
                return;
            }

            // Validation: reject too long names (Excel uses 31 chars max)
            if new_name.chars().count() > 31 {
                self.status_message = Some("Sheet name cannot exceed 31 characters".to_string());
                reset_state(self);
                cx.notify();
                return;
            }

            // Validation: reject duplicates (case-insensitive)
            let is_duplicate = self.wb(cx).sheet_names()
                .iter()
                .enumerate()
                .any(|(i, name)| i != index && name.eq_ignore_ascii_case(new_name));

            if is_duplicate {
                self.status_message = Some(format!("Sheet '{}' already exists", new_name));
                reset_state(self);
                cx.notify();
                return;
            }

            // Apply the rename
            let new_name_owned = new_name.to_string();
            self.wb_mut(cx, |wb| wb.rename_sheet(index, &new_name_owned));
            self.is_modified = true;

            reset_state(self);
            self.request_title_refresh(cx);
        }
    }

    /// Cancel the sheet rename
    pub fn cancel_sheet_rename(&mut self, cx: &mut Context<Self>) {
        self.renaming_sheet = None;
        self.sheet_rename_input.clear();
        self.sheet_rename_cursor = 0;
        self.sheet_rename_select_all = false;
        self.stop_caret_blink();
        cx.notify();
    }

    /// Handle input for sheet rename
    pub fn sheet_rename_input_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            // If select-all is active, replace all text
            if self.sheet_rename_select_all {
                self.sheet_rename_input.clear();
                self.sheet_rename_cursor = 0;
                self.sheet_rename_select_all = false;
            }
            // Insert at cursor position
            let byte_idx = self.sheet_rename_cursor.min(self.sheet_rename_input.len());
            self.sheet_rename_input.insert(byte_idx, c);
            self.sheet_rename_cursor = byte_idx + c.len_utf8();
            self.reset_caret_activity();
            cx.notify();
        }
    }

    /// Handle backspace for sheet rename
    pub fn sheet_rename_backspace(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            // If select-all is active, clear all text
            if self.sheet_rename_select_all {
                self.sheet_rename_input.clear();
                self.sheet_rename_cursor = 0;
                self.sheet_rename_select_all = false;
                cx.notify();
                return;
            }
            // Delete char before cursor
            if self.sheet_rename_cursor > 0 {
                let prev_byte = self.sheet_rename_prev_char_boundary();
                self.sheet_rename_input.remove(prev_byte);
                self.sheet_rename_cursor = prev_byte;
                cx.notify();
            }
        }
    }

    /// Handle delete key for sheet rename
    pub fn sheet_rename_delete(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            // If select-all is active, clear all text
            if self.sheet_rename_select_all {
                self.sheet_rename_input.clear();
                self.sheet_rename_cursor = 0;
                self.sheet_rename_select_all = false;
                cx.notify();
                return;
            }
            // Delete char at cursor
            if self.sheet_rename_cursor < self.sheet_rename_input.len() {
                self.sheet_rename_input.remove(self.sheet_rename_cursor);
                cx.notify();
            }
        }
    }

    /// Move cursor left in sheet rename
    pub fn sheet_rename_cursor_left(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_select_all = false;
            if self.sheet_rename_cursor > 0 {
                self.sheet_rename_cursor = self.sheet_rename_prev_char_boundary();
                cx.notify();
            }
        }
    }

    /// Move cursor right in sheet rename
    pub fn sheet_rename_cursor_right(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_select_all = false;
            if self.sheet_rename_cursor < self.sheet_rename_input.len() {
                self.sheet_rename_cursor = self.sheet_rename_next_char_boundary();
                cx.notify();
            }
        }
    }

    /// Move cursor to start in sheet rename
    pub fn sheet_rename_cursor_home(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_select_all = false;
            self.sheet_rename_cursor = 0;
            cx.notify();
        }
    }

    /// Move cursor to end in sheet rename
    pub fn sheet_rename_cursor_end(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_select_all = false;
            self.sheet_rename_cursor = self.sheet_rename_input.len();
            cx.notify();
        }
    }

    /// Get previous char boundary for sheet rename cursor
    fn sheet_rename_prev_char_boundary(&self) -> usize {
        let s = &self.sheet_rename_input;
        let mut idx = self.sheet_rename_cursor.saturating_sub(1);
        while idx > 0 && !s.is_char_boundary(idx) {
            idx -= 1;
        }
        idx
    }

    /// Get next char boundary for sheet rename cursor
    fn sheet_rename_next_char_boundary(&self) -> usize {
        let s = &self.sheet_rename_input;
        let mut idx = self.sheet_rename_cursor + 1;
        while idx < s.len() && !s.is_char_boundary(idx) {
            idx += 1;
        }
        idx.min(s.len())
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

    /// Show right-click context menu for cells or headers
    pub fn show_context_menu(
        &mut self,
        kind: crate::app::ContextMenuKind,
        position: gpui::Point<gpui::Pixels>,
        cx: &mut Context<Self>,
    ) {
        // Commit edit if in edit mode (save value, stay in place — don't move cursor)
        if self.mode.is_editing() {
            self.commit_pending_edit(cx);
        }
        // Cancel format painter if active
        if self.mode == crate::mode::Mode::FormatPainter {
            self.cancel_format_painter(cx);
        }
        self.context_menu = Some(crate::app::ContextMenuState { kind, position });
        cx.notify();
    }

    /// Hide right-click context menu
    pub fn hide_context_menu(&mut self, cx: &mut Context<Self>) {
        self.context_menu = None;
        cx.notify();
    }

    /// Delete a sheet
    pub fn delete_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        if self.wb_mut(cx, |wb| wb.delete_sheet(index)) {
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

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod sheet_rename_tests {
    /// Test UTF-8 char boundary navigation
    #[test]
    fn utf8_prev_char_boundary() {
        // Test with ASCII
        let s = "Sheet1";
        assert_eq!(prev_boundary(s, 6), 5); // '1' -> 't'
        assert_eq!(prev_boundary(s, 1), 0); // 'h' -> 'S'
        assert_eq!(prev_boundary(s, 0), 0); // at start, stay at 0

        // Test with multi-byte UTF-8 (Chinese characters are 3 bytes each)
        let s = "中文";
        assert_eq!(s.len(), 6); // 2 chars × 3 bytes
        assert_eq!(prev_boundary(s, 6), 3); // end -> after first char
        assert_eq!(prev_boundary(s, 3), 0); // after first char -> start

        // Test with emoji (4 bytes)
        let s = "A👩B";
        assert_eq!(s.len(), 6); // 1 + 4 + 1
        assert_eq!(prev_boundary(s, 6), 5); // after B -> before B
        assert_eq!(prev_boundary(s, 5), 1); // before B -> after A (skip emoji)
        assert_eq!(prev_boundary(s, 1), 0); // after A -> start
    }

    #[test]
    fn utf8_next_char_boundary() {
        // Test with ASCII
        let s = "Sheet1";
        assert_eq!(next_boundary(s, 0), 1); // 'S' -> 'h'
        assert_eq!(next_boundary(s, 5), 6); // 't' -> end

        // Test with multi-byte UTF-8
        let s = "中文";
        assert_eq!(next_boundary(s, 0), 3); // start -> after first char
        assert_eq!(next_boundary(s, 3), 6); // after first -> end

        // Test with emoji
        let s = "A👩B";
        assert_eq!(next_boundary(s, 0), 1); // start -> after A
        assert_eq!(next_boundary(s, 1), 5); // after A -> after emoji
        assert_eq!(next_boundary(s, 5), 6); // after emoji -> end
    }

    #[test]
    fn select_all_replacement() {
        // Simulate: text = "Sheet1", select_all = true, type 'A'
        // Expected: text = "A", cursor = 1, select_all = false
        let mut text = String::from("Sheet1");
        let mut cursor = text.len();
        let mut select_all = true;

        // Simulate input char 'A' with select_all active
        if select_all {
            text.clear();
            cursor = 0;
            select_all = false;
        }
        let c = 'A';
        let byte_idx = cursor.min(text.len());
        text.insert(byte_idx, c);
        cursor = byte_idx + c.len_utf8();

        assert_eq!(text, "A");
        assert_eq!(cursor, 1);
        assert!(!select_all);
    }

    #[test]
    fn backspace_with_select_all_clears() {
        // Simulate: text = "Sheet1", select_all = true, backspace
        // Expected: text = "", cursor = 0, select_all = false
        let mut text = String::from("Sheet1");
        let mut cursor = text.len();
        let mut select_all = true;

        // Simulate backspace with select_all active
        if select_all {
            text.clear();
            cursor = 0;
            select_all = false;
        }

        assert_eq!(text, "");
        assert_eq!(cursor, 0);
        assert!(!select_all);
    }

    #[test]
    fn name_validation_trims() {
        // Test trimming
        let input = "  Sheet1  ";
        let trimmed = input.trim();
        assert_eq!(trimmed, "Sheet1");
    }

    #[test]
    fn name_validation_rejects_too_long() {
        // Names > 31 chars are rejected (not capped)
        let long_name = "A".repeat(50);
        let char_count = long_name.chars().count();
        assert!(char_count > 31, "Test requires > 31 chars");
        // Policy: reject, don't cap
    }

    #[test]
    fn duplicate_check_case_insensitive() {
        // "Sheet1" and "SHEET1" should be considered duplicates
        let existing = "Sheet1";
        let attempt = "SHEET1";
        assert!(existing.eq_ignore_ascii_case(attempt));
    }

    // Helper: simulate prev_char_boundary logic
    fn prev_boundary(s: &str, cursor: usize) -> usize {
        let mut idx = cursor.saturating_sub(1);
        while idx > 0 && !s.is_char_boundary(idx) {
            idx -= 1;
        }
        idx
    }

    // Helper: simulate next_char_boundary logic
    fn next_boundary(s: &str, cursor: usize) -> usize {
        let mut idx = cursor + 1;
        while idx < s.len() && !s.is_char_boundary(idx) {
            idx += 1;
        }
        idx.min(s.len())
    }
}
