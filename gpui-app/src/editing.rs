//! Cell editing and caret blinking functionality
//!
//! Handles:
//! - Starting/confirming edits
//! - Multi-edit (applying edits to multiple cells)
//! - Text manipulation (backspace, delete, insert)
//! - Caret blinking animation
//! - Verified mode (deterministic recalc)

use gpui::*;
use crate::app::{Spreadsheet, EditorSurface, is_smoke_recalc_enabled};
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Edit Mode Management
    // =========================================================================

    /// Enter formula mode with clean state. Called from every path that
    /// transitions into Formula mode (initial char, start_edit, recompute).
    fn enter_formula_mode(&mut self) {
        self.mode = Mode::Formula;
        self.formula_nav_mode = crate::mode::FormulaNavMode::Point;
        self.formula_nav_manual_override = None;
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
    }

    /// Reset formula/edit transient state. Called on every exit from edit mode.
    fn reset_edit_state(&mut self) {
        self.formula_nav_mode = crate::mode::FormulaNavMode::Point;
        self.formula_nav_manual_override = None;
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
    }

    /// Recompute edit mode based on current edit buffer content.
    /// Call this after every edit buffer change to keep mode in sync.
    /// - If edit buffer starts with '=' or '+', mode is Formula (ref-pick enabled)
    /// - Otherwise, mode is Edit (plain text)
    pub(crate) fn recompute_edit_mode(&mut self) {
        if !self.mode.is_editing() {
            return; // Only relevant when editing
        }

        let is_formula = self.edit_value.starts_with('=') || self.edit_value.starts_with('+');
        let should_be_formula = is_formula;
        let currently_formula = self.mode.is_formula();

        if should_be_formula && !currently_formula {
            // Transition Edit -> Formula
            self.enter_formula_mode();
        } else if !should_be_formula && currently_formula {
            // Transition Formula -> Edit (user deleted the '=')
            self.mode = Mode::Edit;
            // Clear formula-specific state
            self.formula_ref_cell = None;
            self.formula_ref_end = None;
            self.formula_highlighted_refs.clear();
        }
    }

    // =========================================================================
    // Editing
    // =========================================================================

    pub fn start_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Block editing during preview mode
        if self.block_if_previewing(cx) { return; }

        // Clear copy/cut border overlay when entering edit mode
        self.clipboard_visual_range = None;

        let (mut row, mut col) = self.view_state.selected;

        // If on a merge-hidden cell, redirect selection to the merge origin
        let merge_redirect = self.sheet(cx).get_merge(row, col)
            .filter(|m| m.start != (row, col))
            .map(|m| (m.start, m.end));
        if let Some((origin, end)) = merge_redirect {
            row = origin.0;
            col = origin.1;
            self.view_state.selected = (row, col);
            self.view_state.selection_end = Some(end);
        }

        // Block editing spill receivers - show message and redirect to parent
        if let Some((parent_row, parent_col)) = self.sheet(cx).get_spill_parent(row, col) {
            let parent_ref = self.cell_ref_at(parent_row, parent_col);
            self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
            cx.notify();
            return;
        }

        self.edit_original = self.sheet(cx).get_raw(row, col);
        self.edit_value = self.edit_original.clone();
        self.edit_cursor = self.edit_value.len();  // Cursor at end (byte offset)
        self.edit_scroll_x = 0.0;
        self.edit_scroll_dirty = true;  // Trigger scroll update to show caret
        self.formula_bar_cache_dirty = true;  // Rebuild hit-test cache
        self.formula_bar_scroll_x = 0.0;
        self.active_editor = EditorSurface::Cell;  // Default to cell editor
        self.edit_selection_anchor = None;

        // Debug assert: cursor must be valid
        debug_assert!(
            self.edit_cursor <= self.edit_value.len(),
            "edit_cursor {} exceeds edit_value.len() {}",
            self.edit_cursor, self.edit_value.len()
        );

        // Set mode based on content: Formula if starts with '=' or '+', else Edit
        let is_formula = self.edit_value.starts_with('=') || self.edit_value.starts_with('+');
        if is_formula {
            self.enter_formula_mode();
        } else {
            self.mode = Mode::Edit;
        }

        // Parse and highlight formula references if editing a formula
        // Clear color map for fresh edit session
        self.clear_formula_ref_colors();
        if is_formula {
            self.update_formula_refs();
            // F2 on existing formula: start in Caret mode (user wants to edit text)
            self.formula_nav_mode = crate::mode::FormulaNavMode::Caret;
        } else {
            self.formula_highlighted_refs.clear();
        }

        self.start_caret_blink(cx);
        cx.notify();
    }

    pub fn start_edit_clear(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Block editing during preview mode
        if self.block_if_previewing(cx) { return; }

        // Clear copy/cut border overlay when entering edit mode
        self.clipboard_visual_range = None;

        let (mut row, mut col) = self.view_state.selected;

        // If on a merge-hidden cell, redirect selection to the merge origin
        let merge_redirect = self.sheet(cx).get_merge(row, col)
            .filter(|m| m.start != (row, col))
            .map(|m| (m.start, m.end));
        if let Some((origin, end)) = merge_redirect {
            row = origin.0;
            col = origin.1;
            self.view_state.selected = (row, col);
            self.view_state.selection_end = Some(end);
        }

        // Block editing spill receivers - show message and redirect to parent
        if let Some((parent_row, parent_col)) = self.sheet(cx).get_spill_parent(row, col) {
            let parent_ref = self.cell_ref_at(parent_row, parent_col);
            self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
            cx.notify();
            return;
        }

        self.edit_original = self.sheet(cx).get_raw(row, col);
        self.edit_value = String::new();
        self.edit_cursor = 0;
        self.edit_scroll_x = 0.0;
        self.edit_scroll_dirty = true;  // Trigger scroll update
        self.formula_bar_cache_dirty = true;  // Rebuild hit-test cache
        self.formula_bar_scroll_x = 0.0;
        self.active_editor = EditorSurface::Cell;  // Default to cell editor
        self.edit_selection_anchor = None;
        // Clear formula state - fresh edit session with empty buffer
        self.clear_formula_ref_colors();
        self.formula_highlighted_refs.clear();
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
        self.mode = Mode::Edit;
        self.start_caret_blink(cx);
        cx.notify();
    }

    /// Commit edit and move down (Enter, or Down arrow in Edit mode)
    ///
    /// Multi-edit: If multiple cells selected, applies to all (the "wow" moment).
    /// Single cell: commits and moves down.
    ///
    /// # Commit-on-Arrow Policy (Excel-like fast data entry)
    ///
    /// In Mode::Edit (non-formula): Arrow keys commit the edit and move selection.
    /// In Mode::Formula: Arrow keys do ref-picking (Option A), NOT commit.
    pub fn confirm_edit(&mut self, cx: &mut Context<Self>) {
        // Multi-edit: If multiple cells selected, apply to all (the "wow" moment)
        if self.is_multi_selection() {
            self.confirm_edit_in_place(cx);
        } else {
            self.confirm_edit_and_move(1, 0, cx);  // Enter moves down
        }
    }

    /// Commit edit and move up (Shift+Enter, or Up arrow in Edit mode)
    ///
    /// See `confirm_edit` for commit-on-arrow policy.
    pub fn confirm_edit_up(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(-1, 0, cx);  // Shift+Enter moves up
    }

    /// Commit any pending edit without moving the cursor.
    /// Call this before file operations (Save, Export) to ensure unsaved edits are captured.
    pub fn commit_pending_edit(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            return;
        }

        let (row, col) = self.view_state.selected;
        let old_value = self.edit_original.clone();

        // Convert leading + to = for formulas (Excel compatibility)
        let mut new_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if new_value.starts_with('=') {
            let open_count = new_value.chars().filter(|&c| c == '(').count();
            let close_count = new_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    new_value.push(')');
                }
            }
        }

        self.history.record_change(self.sheet_index(cx), row, col, old_value, new_value.clone());
        self.set_cell_value(row, col, &new_value, cx);  // Use helper that updates dep graph
        self.mode = Mode::Navigation;
        self.reset_edit_state();
        self.edit_value.clear();
        self.edit_original.clear();
        self.bump_cells_rev();
        self.is_modified = true;
        // Clear formula state
        self.formula_ref_start_cursor = 0;
        self.formula_highlighted_refs.clear();
        // Non-Enter commit breaks tab chain (Save/Export path)
        self.tab_chain_origin_col = None;

        // Smoke mode: trigger full ordered recompute for dogfooding
        self.maybe_smoke_recalc(cx);

        // Auto-clear invalid marker if cell is now valid (Phase 6C)
        if self.invalid_cells.contains_key(&(row, col)) {
            use visigrid_engine::validation::ValidationResult;
            let display_value = self.sheet(cx).get_display(row, col);
            let result = self.wb(cx).validate_cell_input(self.sheet_index(cx), row, col, &display_value);
            if matches!(result, ValidationResult::Valid) {
                self.clear_cell_invalid(row, col);
            }
        }
    }

    /// Run full ordered recompute if enabled (smoke mode or verified mode).
    ///
    /// - Smoke mode (VISIGRID_RECALC=full): Logs to file for dogfooding
    /// - Verified mode: Updates last_recalc_report for status bar display
    pub(crate) fn maybe_smoke_recalc(&mut self, cx: &mut Context<Self>) {
        let smoke_enabled = is_smoke_recalc_enabled();

        // Skip if neither mode is active or we're already in a recalc
        if (!smoke_enabled && !self.verified_mode) || self.in_smoke_recalc {
            return;
        }

        self.in_smoke_recalc = true;
        let report = self.wb_mut(cx, |wb| wb.recompute_full_ordered());

        // Store report for verified mode status bar
        if self.verified_mode {
            self.last_recalc_report = Some(report.clone());
        }

        // Smoke mode logging
        if smoke_enabled {
            let log_line = report.log_line();

            // On Linux/macOS: print to stderr (visible in terminal)
            #[cfg(not(target_os = "windows"))]
            eprintln!("{}", log_line);

            // On all platforms: also write to file (Windows GUI apps don't have stderr)
            use std::io::Write;
            let log_path = dirs::home_dir()
                .map(|p| p.join("smoke.log"))
                .unwrap_or_else(|| std::path::PathBuf::from("smoke.log"));
            if let Ok(mut f) = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(&log_path)
            {
                let _ = writeln!(f, "{}", log_line);
            }
        }

        self.in_smoke_recalc = false;
    }

    /// Toggle verified mode on/off.
    pub fn toggle_verified_mode(&mut self, cx: &mut Context<Self>) {
        self.verified_mode = !self.verified_mode;
        if self.verified_mode {
            // Run initial recalc when enabling
            self.in_smoke_recalc = true;
            let report = self.wb_mut(cx, |wb| wb.recompute_full_ordered());
            self.last_recalc_report = Some(report);
            self.in_smoke_recalc = false;
            self.status_message = Some("Verified mode enabled".to_string());
        } else {
            self.last_recalc_report = None;
            self.status_message = Some("Verified mode disabled".to_string());
        }
        cx.notify();
    }

    // =========================================================================
    // Semantic Approval (Fingerprint Boundary)
    // =========================================================================

    /// Approve the current semantic state.
    ///
    /// This captures the current fingerprint as the "known-good" state.
    /// Future changes to formulas, values, or metadata will invalidate approval.
    /// Formatting changes (bold, colors, column widths) do NOT affect approval.
    ///
    /// If already approved and drifted, shows a confirmation dialog first.
    pub fn approve_model(&mut self, note: Option<String>, cx: &mut Context<Self>) {
        // If drifted from a previous approval, show confirmation
        if self.approval_status() == crate::app::ApprovalStatus::Drifted {
            self.approval_confirm_visible = true;
            cx.notify();
            return;
        }

        self.approve_model_confirmed(note, cx);
    }

    /// Approve without confirmation (called directly or after confirmation).
    pub fn approve_model_confirmed(&mut self, note: Option<String>, cx: &mut Context<Self>) {
        let fingerprint = self.history.fingerprint();
        self.approved_fingerprint = Some(fingerprint);
        self.approval_timestamp = Some(std::time::Instant::now());
        self.approval_note = note;
        self.approval_confirm_visible = false;

        self.status_message = Some("Model approved".to_string());
        cx.notify();
    }

    /// Cancel the approval confirmation dialog.
    pub fn cancel_approval_confirm(&mut self, cx: &mut Context<Self>) {
        self.approval_confirm_visible = false;
        cx.notify();
    }

    /// Clear the approved state.
    pub fn clear_approval(&mut self, cx: &mut Context<Self>) {
        self.approved_fingerprint = None;
        self.approval_timestamp = None;
        self.approval_note = None;
        self.approval_confirm_visible = false;

        self.status_message = Some("Approval cleared".to_string());
        cx.notify();
    }

    /// Get the current approval status.
    pub fn approval_status(&self) -> crate::app::ApprovalStatus {
        match &self.approved_fingerprint {
            None => crate::app::ApprovalStatus::NotApproved,
            Some(approved) => {
                let current = self.history.fingerprint();
                if current == *approved {
                    crate::app::ApprovalStatus::Approved
                } else {
                    crate::app::ApprovalStatus::Drifted
                }
            }
        }
    }

    /// Check if the current state matches the approved fingerprint.
    pub fn is_approved(&self) -> bool {
        self.approval_status() == crate::app::ApprovalStatus::Approved
    }

    /// Get a display string for the approval status.
    pub fn approval_display(&self) -> &'static str {
        match self.approval_status() {
            crate::app::ApprovalStatus::NotApproved => "",
            crate::app::ApprovalStatus::Approved => "Approved ✓",
            crate::app::ApprovalStatus::Drifted => "Drifted ⚠",
        }
    }

    /// Get a summary of what changed since approval (for the drift dialog).
    /// Returns a list of change descriptions.
    pub fn approval_drift_summary(&self) -> Vec<String> {
        // For now, return a simple message. Future: integrate with history diff.
        vec!["Logic has changed since the model was approved.".to_string()]
    }

    /// Force full recalculation of all formulas (F9 - Excel muscle memory).
    ///
    /// In Excel, F9 recalculates all formulas in all open workbooks.
    /// VisiGrid uses automatic recalc, so this is primarily useful for:
    /// - Refreshing volatile functions (NOW, TODAY, RAND, etc.)
    /// - Forcing recalc after external data changes
    /// - Verifying formula results match expectations
    pub fn recalculate(&mut self, cx: &mut Context<Self>) {
        self.in_smoke_recalc = true;
        let report = self.wb_mut(cx, |wb| wb.recompute_full_ordered());
        self.in_smoke_recalc = false;

        // Build informative status message
        let cells = report.cells_recomputed;
        let ms = report.duration_ms;
        let verified_suffix = if self.verified_mode { " · Verified" } else { "" };
        let msg = if cells == 0 {
            format!("Recalculated · no formulas{}", verified_suffix)
        } else if ms == 0 {
            format!("Recalculated · {} cells · <1 ms{}", cells, verified_suffix)
        } else {
            format!("Recalculated · {} cells · {} ms{}", cells, ms, verified_suffix)
        };
        self.status_message = Some(msg);

        if self.verified_mode {
            self.last_recalc_report = Some(report);
        }
        cx.notify();
    }

    /// Commit edit and move right (Right arrow in Edit mode)
    ///
    /// # Commit-on-Arrow Policy (Excel-like fast data entry)
    ///
    /// In Mode::Edit (non-formula): Arrow keys commit the edit and move selection.
    /// This enables fast grid data entry without pressing Enter after each cell.
    ///
    /// In Mode::Formula: Arrow keys do ref-picking (Option A), NOT commit.
    /// Exit formula mode with Enter (confirm) or Escape (cancel).
    pub fn confirm_edit_and_move_right(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, 1, cx);
    }

    /// Commit edit and move left (Left arrow in Edit mode)
    ///
    /// See `confirm_edit_and_move_right` for commit-on-arrow policy.
    pub fn confirm_edit_and_move_left(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, -1, cx);
    }

    /// Tab commit-and-move-right: records the tab-chain origin column,
    /// then commits the current edit and moves right.
    ///
    /// The origin column is used by `confirm_edit_enter` to return the
    /// cursor to the starting column when Enter is pressed (Excel behavior).
    pub fn confirm_edit_and_tab_right(&mut self, cx: &mut Context<Self>) {
        if self.tab_chain_origin_col.is_none() {
            self.tab_chain_origin_col = Some(self.view_state.selected.1);
        }
        self.confirm_edit_and_move(0, 1, cx);
    }

    /// Shift+Tab commit-and-move-left: records the tab-chain origin column,
    /// then commits the current edit and moves left.
    pub fn confirm_edit_and_tab_left(&mut self, cx: &mut Context<Self>) {
        if self.tab_chain_origin_col.is_none() {
            self.tab_chain_origin_col = Some(self.view_state.selected.1);
        }
        self.confirm_edit_and_move(0, -1, cx);
    }

    /// Enter key: confirm edit and move down, with tab-chain return.
    ///
    /// If the user tabbed across cells entering data, Enter returns the cursor
    /// to the origin column on the next row instead of staying in the current
    /// column. This matches Excel's Tab-chain return behavior.
    pub fn confirm_edit_enter(&mut self, cx: &mut Context<Self>) {
        if self.is_multi_selection() && self.mode.is_editing() {
            self.tab_chain_origin_col = None;
            self.confirm_edit_in_place(cx);
            return;
        }

        if let Some(origin_col) = self.tab_chain_origin_col.take() {
            // Commit the edit if currently editing
            self.commit_current_edit(cx);
            // Move to next row at the origin column
            let (row, _) = self.active_view_state().selected;
            let new_row = self.next_visible_row(row, 1);
            self.close_validation_dropdown(
                crate::validation_dropdown::DropdownCloseReason::SelectionChanged,
                cx,
            );
            // Snap to merge origin if landing on a merged cell
            let (final_row, final_col) = if let Some(merge) = self.sheet(cx).get_merge(new_row, origin_col) {
                merge.start
            } else {
                (new_row, origin_col)
            };
            let view_state = self.active_view_state_mut();
            view_state.selected = (final_row, final_col);
            view_state.selection_end = None;
            view_state.additional_selections.clear();
            self.ensure_visible(cx);
        } else {
            self.confirm_edit(cx);
        }
    }

    /// Shift+Enter key: confirm edit and move up, with tab-chain return.
    pub fn confirm_edit_up_enter(&mut self, cx: &mut Context<Self>) {
        if let Some(origin_col) = self.tab_chain_origin_col.take() {
            self.commit_current_edit(cx);
            let (row, _) = self.active_view_state().selected;
            let new_row = self.next_visible_row(row, -1);
            self.close_validation_dropdown(
                crate::validation_dropdown::DropdownCloseReason::SelectionChanged,
                cx,
            );
            // Snap to merge origin if landing on a merged cell
            let (final_row, final_col) = if let Some(merge) = self.sheet(cx).get_merge(new_row, origin_col) {
                merge.start
            } else {
                (new_row, origin_col)
            };
            let view_state = self.active_view_state_mut();
            view_state.selected = (final_row, final_col);
            view_state.selection_end = None;
            view_state.additional_selections.clear();
            self.ensure_visible(cx);
        } else {
            self.confirm_edit_up(cx);
        }
    }

    /// Ctrl+Enter: Multi-edit commit / Fill selection / Open link
    ///
    /// Behavior (Excel muscle memory):
    /// - If editing: apply edit to ALL selected cells with formula shifting
    /// - If navigation + multi-selection: fill selection from primary cell
    /// - If navigation + single cell + link: open link
    /// - If navigation + single cell + no link: start editing
    ///
    /// Multi-edit semantics:
    /// - Applies edited value to all cells in primary selection AND additional_selections
    /// - For formulas: shifts relative references for each target cell
    ///   (e.g., =A1 typed at B2, applied to C3, becomes =B2)
    /// - Absolute references ($A$1) are preserved unchanged
    /// - One undo step for all changes
    pub fn confirm_edit_in_place(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // Navigation mode: fill selection or open link
            if self.is_multi_selection() {
                // Multi-selection: fill from primary cell (Excel Ctrl+Enter)
                self.fill_selection_from_primary(cx);
                return;
            }
            // Single cell: try to open link, else start editing
            if self.try_open_link(cx) {
                return;
            }
            self.start_edit(cx);
            return;
        }

        // Convert leading + to = for formulas (Excel compatibility)
        let mut base_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if base_value.starts_with('=') {
            let open_count = base_value.chars().filter(|&c| c == '(').count();
            let close_count = base_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    base_value.push(')');
                }
            }
        }

        let is_formula = base_value.starts_with('=');
        let primary_cell = self.view_state.selected;  // Base cell for formula reference shifting

        // Collect all target cells from primary selection and additional_selections
        let mut target_cells: Vec<(usize, usize)> = Vec::new();

        // Primary selection rectangle
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                target_cells.push((row, col));
            }
        }

        // Additional selections (Ctrl+Click)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_r = start.0.min(end.0);
            let max_r = start.0.max(end.0);
            let min_c = start.1.min(end.1);
            let max_c = start.1.max(end.1);
            for row in min_r..=max_r {
                for col in min_c..=max_c {
                    // Avoid duplicates (primary selection might overlap)
                    if !target_cells.contains(&(row, col)) {
                        target_cells.push((row, col));
                    }
                }
            }
        }

        let mut changes = Vec::new();

        // Apply to all target cells
        for (row, col) in &target_cells {
            // Skip spill receivers
            if self.sheet(cx).get_spill_parent(*row, *col).is_some() {
                continue;
            }

            let old_value = self.sheet(cx).get_raw(*row, *col);

            // For formulas, shift relative references based on delta from primary cell
            let new_value = if is_formula {
                let delta_row = *row as i32 - primary_cell.0 as i32;
                let delta_col = *col as i32 - primary_cell.1 as i32;
                self.adjust_formula_refs(&base_value, delta_row, delta_col)
            } else {
                base_value.clone()
            };

            if new_value != old_value {
                changes.push(crate::history::CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value,
                });
            }
        }

        // Apply all changes (batched to defer recalc until all cells set)
        let sheet_index = self.sheet_index(cx);
        self.workbook.update(cx, |wb, _| wb.begin_batch());
        for change in &changes {
            self.set_cell_value(change.row, change.col, &change.new_value, cx);
        }
        self.end_batch_and_broadcast(cx);

        // Record batch for undo
        let had_changes = !changes.is_empty();
        if had_changes {
            self.history.record_batch(sheet_index, changes);
        }

        // Exit edit mode
        self.mode = Mode::Navigation;
        self.reset_edit_state();
        self.edit_value.clear();
        self.edit_original.clear();
        self.formula_highlighted_refs.clear();
        self.clear_formula_ref_colors();
        self.autocomplete_visible = false;
        self.bump_cells_rev();
        self.is_modified = true;
        self.maybe_smoke_recalc(cx);

        // Invalidate trace cache (dependencies may have changed)
        if had_changes {
            self.invalidate_trace_if_needed(cx);
        }

        cx.notify();
    }

    pub fn cancel_edit(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.reset_edit_state();
        self.edit_value.clear();
        self.edit_original.clear();
        self.formula_highlighted_refs.clear();
        self.clear_formula_ref_colors();
        self.autocomplete_visible = false;
        self.tab_chain_origin_col = None;  // Escape breaks tab chain
        self.stop_caret_blink();
        cx.notify();
    }

    // =========================================================================
    // Caret Blinking
    // =========================================================================

    /// Start the caret blink timer. Called when entering edit mode.
    pub fn start_caret_blink(&mut self, cx: &mut Context<Self>) {
        use std::time::Duration;

        self.caret_visible = true;
        self.caret_last_activity = std::time::Instant::now();

        // Cancel any existing blink task
        self.caret_blink_task = None;

        // Spawn repeating blink task
        let task = cx.spawn(async move |this, cx| {
            let blink_interval = Duration::from_millis(530);
            let idle_delay = Duration::from_millis(500);

            loop {
                // Wait for blink interval
                smol::Timer::after(blink_interval).await;

                // Update caret visibility
                let should_continue = this.update(cx, |this, cx| {
                    // Don't blink if not editing (cell or sheet rename)
                    let is_sheet_renaming = this.renaming_sheet.is_some() && !this.sheet_rename_select_all;
                    if !this.mode.is_editing() && !is_sheet_renaming {
                        return false;
                    }

                    // Don't blink if there's a text selection (cell edit only)
                    if this.edit_selection_anchor.is_some() {
                        this.caret_visible = true;
                        cx.notify();
                        return true;
                    }

                    // Don't blink during active typing (wait for idle)
                    if this.caret_last_activity.elapsed() < idle_delay {
                        this.caret_visible = true;
                        cx.notify();
                        return true;
                    }

                    // Toggle visibility
                    this.caret_visible = !this.caret_visible;
                    cx.notify();
                    true
                });

                match should_continue {
                    Ok(true) => continue,
                    _ => break,
                }
            }
        });

        self.caret_blink_task = Some(task);
    }

    /// Stop the caret blink timer. Called when leaving edit mode.
    pub fn stop_caret_blink(&mut self) {
        self.caret_blink_task = None;
        self.caret_visible = true;
    }

    /// Reset caret activity timestamp. Called on text edits and cursor moves.
    /// Keeps caret visible and resets the idle timer.
    pub fn reset_caret_activity(&mut self) {
        self.caret_visible = true;
        self.caret_last_activity = std::time::Instant::now();
    }

    // =========================================================================
    // Text Manipulation
    // =========================================================================

    /// Delete selected text and return true if there was a selection
    fn delete_edit_selection(&mut self) -> bool {
        if let Some((start_byte, end_byte)) = self.edit_selection_range() {
            // start_byte and end_byte are already byte offsets
            let start_byte = start_byte.min(self.edit_value.len());
            let end_byte = end_byte.min(self.edit_value.len());
            self.edit_value.replace_range(start_byte..end_byte, "");
            self.edit_cursor = start_byte;
            self.edit_selection_anchor = None;
            true
        } else {
            false
        }
    }

    pub fn backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // Text edit: clear ref_target and suppression so autocomplete can reopen
            if self.mode.is_formula() {
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            }
            self.autocomplete_suppressed = false;
            self.reset_caret_activity();

            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Recompute mode (deleting '=' at start should exit formula mode)
                self.recompute_edit_mode();
                // Update highlighted refs for formulas
                if self.mode.is_formula() {
                    self.update_formula_refs();
                    self.clear_formula_nav_override();
                    self.update_formula_nav_mode();
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
                return;
            }
            // Otherwise delete char before cursor (byte-indexed)
            if self.edit_cursor > 0 {
                let prev_byte = self.prev_char_boundary(self.edit_cursor);
                let curr_byte = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.replace_range(prev_byte..curr_byte, "");
                self.edit_cursor = prev_byte;
                // Recompute mode (deleting '=' at start should exit formula mode)
                self.recompute_edit_mode();
                // Update highlighted refs for formulas
                if self.mode.is_formula() {
                    self.update_formula_refs();
                    self.clear_formula_nav_override();
                    self.update_formula_nav_mode();
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
            }
        }
    }

    pub fn delete_char(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // Text edit: clear ref_target and suppression so autocomplete can reopen
            if self.mode.is_formula() {
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            }
            self.autocomplete_suppressed = false;
            self.reset_caret_activity();

            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Recompute mode (deleting '=' at start should exit formula mode)
                self.recompute_edit_mode();
                // Update highlighted refs for formulas
                if self.mode.is_formula() {
                    self.update_formula_refs();
                    self.clear_formula_nav_override();
                    self.update_formula_nav_mode();
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
                return;
            }
            // Otherwise delete char at cursor (byte-indexed)
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                let curr_byte = self.edit_cursor;
                let next_byte = self.next_char_boundary(curr_byte);
                self.edit_value.replace_range(curr_byte..next_byte, "");
                // Cursor stays at same byte position (deleted forward)
                // Recompute mode (deleting '=' at start should exit formula mode)
                self.recompute_edit_mode();
                // Update highlighted refs for formulas
                if self.mode.is_formula() {
                    self.update_formula_refs();
                    self.clear_formula_nav_override();
                    self.update_formula_nav_mode();
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
            }
        }
    }

    pub fn insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // In Formula mode, typing an operator finalizes the current reference
            if self.mode.is_formula() && self.formula_ref_cell.is_some() {
                if Self::is_formula_operator(c) {
                    self.finalize_formula_reference();
                } else {
                    // Non-operator character: clear ref_target since we're typing, not navigating
                    self.formula_ref_cell = None;
                    self.formula_ref_end = None;
                }
            }

            // Delete selection if any (replaces selected text)
            self.delete_edit_selection();

            // Insert at cursor byte position
            let byte_idx = self.edit_cursor.min(self.edit_value.len());
            self.edit_value.insert(byte_idx, c);
            self.edit_cursor = byte_idx + c.len_utf8();  // Advance by byte length of char

            // Recompute mode: typing '=' at start should transition to Formula mode
            self.recompute_edit_mode();

            // Update highlighted refs for formulas
            if self.mode.is_formula() {
                self.update_formula_refs();
                // Buffer mutation clears F2 override, then auto-switch based on caret
                self.clear_formula_nav_override();
                self.update_formula_nav_mode();
            }

            // Text edit: clear suppression so autocomplete can reopen
            self.autocomplete_suppressed = false;

            // Reset caret blink (keep visible while typing)
            self.reset_caret_activity();

            // Mark scroll/cache dirty
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;

            // Update autocomplete for formulas
            self.update_autocomplete(cx);
        } else {
            // Start editing with this character — clear copy/cut border overlay
            self.clipboard_visual_range = None;

            let (row, col) = self.view_state.selected;

            // Block editing spill receivers
            if let Some((parent_row, parent_col)) = self.sheet(cx).get_spill_parent(row, col) {
                let parent_ref = self.cell_ref_at(parent_row, parent_col);
                self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
                cx.notify();
                return;
            }

            self.edit_original = self.sheet(cx).get_raw(row, col);
            self.edit_value = c.to_string();
            self.edit_cursor = c.len_utf8();  // Byte offset after first char

            // Enter Formula mode if starting with = or +
            if c == '=' || c == '+' {
                self.enter_formula_mode();
            } else {
                self.mode = Mode::Edit;
            }

            // Mark scroll/cache dirty
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;
            self.formula_bar_scroll_x = 0.0;
            self.active_editor = EditorSurface::Cell;

            // Start caret blinking
            self.start_caret_blink(cx);

            // Update autocomplete for formulas
            self.update_autocomplete(cx);
        }
        cx.notify();
    }

    /// Check if character is a formula operator that finalizes a reference
    fn is_formula_operator(c: char) -> bool {
        matches!(c, '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | ',' | '(' | ')' | ':' | ';')
    }

    /// Finalize the current formula reference (clear the active reference state)
    fn finalize_formula_reference(&mut self) {
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
    }

    // =========================================================================
    // Edit Movement and Link Opening
    // =========================================================================

    /// Commit the current edit without moving the cursor or changing selection.
    /// Returns true if an edit was actually committed (was in editing mode).
    /// Used by `confirm_edit_and_move` and `confirm_edit_enter`.
    fn commit_current_edit(&mut self, cx: &mut Context<Self>) -> bool {
        if !self.mode.is_editing() {
            return false;
        }

        // Clear copy/cut border overlay on edit commit
        self.clipboard_visual_range = None;

        let (row, col) = self.view_state.selected;
        let old_value = self.edit_original.clone();

        // Convert leading + to = for formulas (Excel compatibility)
        let mut new_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if new_value.starts_with('=') {
            let open_count = new_value.chars().filter(|&c| c == '(').count();
            let close_count = new_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    new_value.push(')');
                }
            }
        }

        // Capture raw edit value before clearing for percent auto-format check
        let raw_edit = self.edit_value.clone();

        self.history.record_change(self.sheet_index(cx), row, col, old_value, new_value.clone());
        self.set_cell_value(row, col, &new_value, cx);

        // Auto-apply Percent format when user typed "X%" and cell format is General
        if raw_edit.trim().ends_with('%') {
            let current_fmt = self.sheet(cx).get_format(row, col).number_format.clone();
            if matches!(current_fmt, visigrid_engine::cell::NumberFormat::General) {
                let fmt = visigrid_engine::cell::NumberFormat::Percent { decimals: 0 };
                self.with_active_sheet_mut(cx, |s| s.set_number_format(row, col, fmt));
            }
        }

        self.mode = Mode::Navigation;
        self.reset_edit_state();
        self.edit_value.clear();
        self.edit_original.clear();
        self.bump_cells_rev();
        self.is_modified = true;
        // Clear formula reference state
        self.formula_ref_start_cursor = 0;
        // Clear formula highlighting state
        self.formula_highlighted_refs.clear();
        self.clear_formula_ref_colors();
        // Close autocomplete
        self.autocomplete_visible = false;
        // Reset editor surface
        self.active_editor = EditorSurface::Cell;
        // Stop caret blinking
        self.stop_caret_blink();

        // Smoke mode: trigger full ordered recompute for dogfooding
        self.maybe_smoke_recalc(cx);

        cx.notify();

        true
    }

    fn confirm_edit_and_move(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // Not editing - just move (Excel behavior)
            self.move_selection(dr, dc, cx);
            return;
        }

        self.commit_current_edit(cx);

        // Move after confirming
        self.move_selection(dr, dc, cx);
    }

    /// Fill selection from primary cell (Ctrl+Enter in navigation mode with multi-selection)
    ///
    /// Excel muscle memory: select range, type in first cell, Ctrl+Enter fills all.
    /// This is the navigation-mode equivalent - fills from existing primary cell content.
    fn fill_selection_from_primary(&mut self, cx: &mut Context<Self>) {
        use crate::history::CellChange;
        use visigrid_engine::provenance::MutationOp;

        if self.block_if_merged("fill selection", cx) { return; }

        let primary_cell = self.view_state.selected;
        let base_value = self.sheet(cx).get_raw(primary_cell.0, primary_cell.1);

        let is_formula = base_value.starts_with('=');

        // Collect all target cells (excluding primary cell itself)
        let mut target_cells: Vec<(usize, usize)> = Vec::new();

        // Primary selection rectangle
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if (row, col) != primary_cell {
                    target_cells.push((row, col));
                }
            }
        }

        // Additional selections (Ctrl+Click)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_r = start.0.min(end.0);
            let max_r = start.0.max(end.0);
            let min_c = start.1.min(end.1);
            let max_c = start.1.max(end.1);
            for row in min_r..=max_r {
                for col in min_c..=max_c {
                    if (row, col) != primary_cell && !target_cells.contains(&(row, col)) {
                        target_cells.push((row, col));
                    }
                }
            }
        }

        if target_cells.is_empty() {
            return;
        }

        let mut changes = Vec::new();
        let mut filled_count = 0;
        let mut skipped_spill = 0;

        self.wb_mut(cx, |wb| wb.begin_batch());
        for (row, col) in &target_cells {
            // Skip spill receivers
            if self.sheet(cx).get_spill_parent(*row, *col).is_some() {
                skipped_spill += 1;
                continue;
            }

            let old_value = self.sheet(cx).get_raw(*row, *col);

            // For formulas, shift relative references based on delta from primary cell
            let new_value = if is_formula {
                let delta_row = *row as i32 - primary_cell.0 as i32;
                let delta_col = *col as i32 - primary_cell.1 as i32;
                self.adjust_formula_refs(&base_value, delta_row, delta_col)
            } else {
                base_value.clone()
            };

            if old_value != new_value {
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: new_value.clone(),
                });
            }
            self.set_cell_value(*row, *col, &new_value, cx);
            filled_count += 1;
        }
        self.end_batch_and_broadcast(cx);

        let sheet_id = self.sheet(cx).id;
        let sheet_name = self.sheet(cx).name.clone();
        if !changes.is_empty() {
            let provenance = MutationOp::MultiEdit {
                sheet: sheet_id,
                cells: target_cells.clone(),
                value: base_value.clone(),
            }.to_provenance(&sheet_name);
            self.history.record_batch_with_provenance(self.sheet_index(cx), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc(cx);
        }

        self.view_state.additional_selections.clear();

        // Status message with optional spill skip note
        let status = if skipped_spill > 0 {
            format!("Filled {} cells (skipped {} spill)", filled_count, skipped_spill)
        } else {
            format!("Filled {} cells", filled_count)
        };
        self.status_message = Some(status);
        cx.notify();
    }

    /// Try to open a detected link in the current cell.
    /// Returns true if a link was found and opened, false otherwise.
    pub fn try_open_link(&mut self, cx: &mut Context<Self>) -> bool {
        use crate::links;

        // Guard: only open links from single-cell selection
        if self.is_multi_selection() {
            return false;
        }

        // Guard: debounce - ignore if already opening a link
        if self.link_open_in_flight {
            return false;
        }

        let (row, col) = self.view_state.selected;
        let cell_value = self.sheet(cx).get_display(row, col);

        if let Some(target) = links::detect_link(&cell_value) {
            let open_string = target.open_string();
            let target_desc = match &target {
                links::LinkTarget::Url(_) => "Opening URL...",
                links::LinkTarget::Email(_) => "Opening email...",
                links::LinkTarget::Path(_) => "Opening file...",
            };

            // Mark as in-flight
            self.link_open_in_flight = true;

            // Open asynchronously to avoid blocking the UI
            cx.spawn(async move |this, cx| {
                let result = open::that(&open_string);

                let _ = this.update(cx, |this, cx| {
                    this.link_open_in_flight = false;
                    this.status_message = Some(match result {
                        Ok(()) => format!("Opened: {}", open_string),
                        Err(e) => format!("Couldn't open link: {}", e),
                    });
                    cx.notify();
                });
            }).detach();

            self.status_message = Some(target_desc.to_string());
            cx.notify();
            true
        } else {
            false
        }
    }

    /// Detect link in current cell (for status bar hint)
    pub fn detected_link(&self, cx: &App) -> Option<crate::links::LinkTarget> {
        let (row, col) = self.view_state.selected;
        let cell_value = self.sheet(cx).get_display(row, col);
        crate::links::detect_link(&cell_value)
    }
}
