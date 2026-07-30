//! Soft-rewind: preview sessions (scrubbing history without committing),
//! rewind planning, safety checks, and the confirm/apply flow.
//! State types live in rewind_state.rs; this is the behavior.
//! Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::app::{Spreadsheet, NUM_ROWS};
use crate::history::HistoryFingerprint;
use crate::rewind_state::*;

impl Spreadsheet {
    /// Get multi-edit preview for a cell during editing.
    /// Returns the value that will be applied to this cell when edit is confirmed.
    /// Returns None if not in multi-edit mode or if this is the active cell.
    pub fn multi_edit_preview(&self, row: usize, col: usize) -> Option<String> {
        // Only in editing mode with multi-selection
        if !self.mode.is_editing() || !self.is_multi_selection() {
            return None;
        }
        // Skip the active cell (it shows the real edit_value)
        if (row, col) == self.view_state.selected {
            return None;
        }
        // Only for selected cells
        if !self.is_selected(row, col) {
            return None;
        }

        // Compute delta from primary cell
        let delta_row = row as i32 - self.view_state.selected.0 as i32;
        let delta_col = col as i32 - self.view_state.selected.1 as i32;

        // If it's a formula, adjust references
        if self.edit_value.starts_with('=') {
            Some(self.adjust_formula_refs(&self.edit_value, delta_row, delta_col))
        } else {
            // Plain text: same value for all cells
            Some(self.edit_value.clone())
        }
    }
    /// Check if we're currently in preview mode
    pub fn is_previewing(&self) -> bool {
        matches!(self.rewind_preview, RewindPreviewState::On(_))
    }
    /// Get the current preview session, if any
    pub fn preview_session(&self) -> Option<&RewindPreviewSession> {
        match &self.rewind_preview {
            RewindPreviewState::On(session) => Some(session),
            RewindPreviewState::Off => None,
        }
    }
    /// Block a command if in preview mode.
    /// Returns true if blocked (command should return early).
    /// Sets status message with consistent preview warning.
    pub fn block_if_previewing(&mut self, cx: &mut Context<Self>) -> bool {
        if self.is_previewing() {
            self.status_message = Some(PREVIEW_BLOCK_MSG.to_string());
            cx.notify();
            true
        } else {
            false
        }
    }
    /// Enter preview mode for the currently selected history entry
    pub fn enter_preview(&mut self, cx: &mut Context<Self>) -> Result<(), String> {
        // Must have a history highlight to preview
        let (sheet_idx, start_row, start_col, end_row, end_col) = match self.history_highlight_range {
            Some(range) => range,
            None => return Err("No history entry selected".to_string()),
        };

        // Must have a selected history entry
        let entry_id = match self.selected_history_id {
            Some(id) => id,
            None => return Err("No history entry selected".to_string()),
        };

        // Find the history index for this entry
        let history_index = match self.history.global_index_for_id(entry_id) {
            Some(idx) => idx,
            None => return Err("History entry not found".to_string()),
        };

        // Get entry info for the session
        let entry = match self.history.entry_at(history_index) {
            Some(e) => e,
            None => return Err("Invalid history index".to_string()),
        };
        let action_summary = entry.action.summary().unwrap_or_else(|| entry.action.label());

        // Build the preview workbook and view state (state BEFORE this action)
        let build_result = self.history.build_workbook_before(
            history_index,
            &self.base_workbook,
            MAX_PREVIEW_REPLAY,
            MAX_PREVIEW_BUILD_MS,
        ).map_err(|e| match e {
            crate::history::PreviewBuildError::InvalidIndex => "Invalid history index".to_string(),
            crate::history::PreviewBuildError::TooManyActions(n) => {
                format!("Preview unavailable — history too large to replay (limit: {} actions)", n)
            }
            crate::history::PreviewBuildError::Timeout => {
                format!("Preview unavailable — replay timed out ({}ms)", MAX_PREVIEW_BUILD_MS)
            }
            crate::history::PreviewBuildError::UnsupportedAction(kind) => {
                format!("Preview unavailable — history contains unsupported action: {}", kind.display_name())
            }
            crate::history::PreviewBuildError::InvariantViolation(msg) => {
                format!("Preview aborted — data integrity error: {}", msg)
            }
        })?;

        // Capture current focus for restoration
        let live_focus = PreviewFocus {
            sheet_index: self.sheet_index(cx),
            selected: self.view_state.selected,
            selection_end: self.view_state.selection_end,
            scroll_row: self.view_state.scroll_row,
            scroll_col: self.view_state.scroll_col,
        };

        // Create the preview session
        let session = RewindPreviewSession {
            entry_id,
            target_global_index: history_index,
            action_summary: action_summary.clone(),
            snapshot: build_result.workbook,
            view_state: build_result.view_state,
            live_focus,
            history_fingerprint: self.history.fingerprint(),
            replay_count: build_result.replay_count,
            build_ms: build_result.build_ms,
            quality: PreviewQuality::Ok,
        };

        self.rewind_preview = RewindPreviewState::On(session);

        // Navigate to the affected area in preview
        // Switch to the sheet where the action occurred
        self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(sheet_idx); });
        self.view_state.selected = (start_row, start_col);
        self.view_state.selection_end = if start_row != end_row || start_col != end_col {
            Some((end_row, end_col))
        } else {
            None
        };

        // Ensure the selection is visible
        self.ensure_visible(cx);

        self.status_message = Some(format!("Preview: Before \"{}\" — Release Space to return", action_summary));
        cx.notify();
        Ok(())
    }
    /// Exit preview mode, restoring live state
    pub fn exit_preview(&mut self, cx: &mut Context<Self>) {
        if let RewindPreviewState::On(session) = std::mem::take(&mut self.rewind_preview) {
            // Restore live focus (Option A: peek behavior)
            self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(session.live_focus.sheet_index); });
            self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
            self.debug_assert_sheet_cache_sync(cx);  // Catch desync at preview exit
            self.view_state.selected = session.live_focus.selected;
            self.view_state.selection_end = session.live_focus.selection_end;
            self.view_state.scroll_row = session.live_focus.scroll_row;
            self.view_state.scroll_col = session.live_focus.scroll_col;

            self.status_message = Some("Returned to current state".to_string());
            cx.notify();
        }
    }
    /// Scrub the preview timeline: navigate to adjacent history entry while holding Space.
    /// direction: -1 for older (up), +1 goes to newer (down)
    pub fn scrub_preview(&mut self, direction: i32, cx: &mut Context<Self>) {
        let current_id = match self.selected_history_id {
            Some(id) => id,
            None => return,
        };

        // Find current position in global history
        let current_idx = match self.history.global_index_for_id(current_id) {
            Some(idx) => idx,
            None => return,
        };

        // Compute new index (direction: -1 goes to older = lower index, +1 goes to newer = higher index)
        let history_len = self.history.undo_count();
        let new_idx = if direction < 0 {
            current_idx.saturating_sub(1)
        } else {
            (current_idx + 1).min(history_len.saturating_sub(1))
        };

        // Don't update if at boundary
        if new_idx == current_idx {
            return;
        }

        // Get the new entry and compute its display info
        let new_entry = match self.history.entry_at(new_idx) {
            Some(e) => e,
            None => return,
        };
        let new_id = new_entry.id;
        let action_summary = new_entry.action.summary()
            .unwrap_or_else(|| new_entry.action.label());

        // Compute highlight range from action details
        let new_highlight = {
            let display_entries = self.history.display_entries();
            display_entries.iter()
                .find(|e| e.id == new_id)
                .and_then(|e| e.sheet_index.and_then(|si| e.affected_range.map(|(sr, sc, er, ec)| (si, sr, sc, er, ec))))
        };

        // Store the current live focus, fingerprint, and quality (preserve across scrubs)
        let (live_focus, history_fingerprint, original_quality) = if let RewindPreviewState::On(ref session) = self.rewind_preview {
            (session.live_focus.clone(), session.history_fingerprint, session.quality.clone())
        } else {
            return; // Not actually previewing
        };

        // Update selection
        self.selected_history_id = Some(new_id);
        self.history_highlight_range = new_highlight;

        // Exit current preview temporarily
        self.rewind_preview = RewindPreviewState::Off;

        // Re-enter preview with new entry
        match self.history.build_workbook_before(
            new_idx,
            &self.base_workbook,
            MAX_PREVIEW_REPLAY,
            MAX_PREVIEW_BUILD_MS,
        ) {
            Ok(build_result) => {
                let session = RewindPreviewSession {
                    entry_id: new_id,
                    target_global_index: new_idx,
                    action_summary: action_summary.clone(),
                    snapshot: build_result.workbook,
                    view_state: build_result.view_state,
                    live_focus,
                    history_fingerprint,  // Preserved from original preview
                    replay_count: build_result.replay_count,
                    build_ms: build_result.build_ms,
                    quality: original_quality,  // Preserve quality from original entry
                };

                self.rewind_preview = RewindPreviewState::On(session);

                // Navigate to the affected area
                if let Some((sheet_idx, start_row, start_col, end_row, end_col)) = new_highlight {
                    self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(sheet_idx); });
                    self.view_state.selected = (start_row, start_col);
                    self.view_state.selection_end = if start_row != end_row || start_col != end_col {
                        Some((end_row, end_col))
                    } else {
                        None
                    };
                    self.ensure_visible(cx);
                }

                self.status_message = Some(format!(
                    "Preview: Before \"{}\" [{}/{}] — ↑↓ to scrub, release Space to return",
                    action_summary, new_idx + 1, history_len
                ));
            }
            Err(e) => {
                // Preview build failed - show error and restore live focus
                self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(live_focus.sheet_index); });
                self.view_state.selected = live_focus.selected;
                self.view_state.selection_end = live_focus.selection_end;
                self.view_state.scroll_row = live_focus.scroll_row;
                self.view_state.scroll_col = live_focus.scroll_col;

                self.status_message = Some(format!("Preview failed: {:?}", e));
            }
        }
        cx.notify();
    }
    /// Build a rewind plan from the current preview session.
    /// Returns None if not previewing or preview is invalid.
    pub fn build_rewind_plan(&self) -> Option<RewindPlan> {
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return None,
        };

        // The truncate point is the target entry index
        // We keep entries [0..target_index), discard [target_index..]
        let truncate_at = session.target_global_index;
        let discarded_count = self.history.undo_count().saturating_sub(truncate_at);

        // Generate timestamp now (will be close to commit time)
        let timestamp_utc = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs().to_string())
            .unwrap_or_else(|_| "0".to_string());

        // Build the audit action with full provenance
        let audit_action = crate::history::UndoAction::Rewind {
            target_entry_id: session.entry_id,
            target_index: session.target_global_index,
            target_action_summary: session.action_summary.clone(),
            discarded_count,
            old_history_len: self.history.undo_count(),
            new_history_len: truncate_at + 1, // After truncate + audit entry
            timestamp_utc,
            preview_replay_count: session.replay_count,
            preview_build_ms: session.build_ms,
        };

        Some(RewindPlan {
            new_workbook: session.snapshot.clone(),
            new_view_state: session.view_state.clone(),
            truncate_at,
            audit_action,
            discarded_count,
            focus: session.live_focus.clone(),
        })
    }
    /// Apply a rewind plan atomically. This is a destructive operation.
    /// Returns Err if the history has changed since the plan was built.
    pub fn apply_rewind_plan(&mut self, plan: RewindPlan, cx: &mut Context<Self>) -> Result<(), String> {
        // Validate history fingerprint hasn't changed
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return Err("No preview active".to_string()),
        };

        let current_fingerprint = self.history.fingerprint();
        if current_fingerprint != session.history_fingerprint {
            return Err(format!(
                "History changed during preview. Expected {:?}, got {:?}. Please re-enter preview to try again.",
                session.history_fingerprint, current_fingerprint
            ));
        }

        // Extract audit entry details before consuming plan
        let (target_entry_id, target_index, action_summary, preview_replay_count, preview_build_ms) = match &plan.audit_action {
            crate::history::UndoAction::Rewind {
                target_entry_id,
                target_index,
                target_action_summary,
                preview_replay_count,
                preview_build_ms,
                ..
            } => (*target_entry_id, *target_index, target_action_summary.clone(), *preview_replay_count, *preview_build_ms),
            _ => return Err("Invalid audit action in plan".to_string()),
        };

        // === ATOMIC COMMIT: Do not fail after this point ===

        // 1. Replace the workbook content
        self.workbook.update(cx, |wb, _| {
            *wb = plan.new_workbook;
        });
        self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
        self.debug_assert_sheet_cache_sync(cx);  // Catch desync at rewind
        // Update base_workbook to match (this is now the canonical state)
        self.base_workbook = self.wb(cx).clone();

        // 2. Apply view state from the plan (row ordering per sheet)
        // Reset row_view to identity for the current sheet
        self.row_view = visigrid_engine::filter::RowView::new(NUM_ROWS);

        // If the preview view state has sort info for current sheet, re-apply it
        let active_idx = self.sheet_index(cx);
        if let Some(sheet_view) = plan.new_view_state.per_sheet.get(active_idx) {
            if let Some(ref row_order) = sheet_view.row_order {
                // Apply the stored row order
                self.row_view.apply_sort(row_order.clone());
            }
        }

        // 3. Truncate history and append audit entry
        self.history.truncate_and_append_rewind(
            plan.truncate_at,
            target_entry_id,
            target_index,
            action_summary.clone(),
            preview_replay_count,
            preview_build_ms,
        );

        // 4. Reset preview state
        self.rewind_preview = RewindPreviewState::Off;

        // 5. Clear history selection/highlight (we're now at end of history)
        self.selected_history_id = None;
        self.history_highlight_range = None;

        // 6. Keep current position in grid (don't restore pre-preview focus)
        // User is looking at the rewound state; changing view would be jarring

        // 7. Mark document as modified
        self.is_modified = true;

        // 8. Status message
        let discarded = plan.discarded_count;
        self.status_message = Some(format!(
            "Rewound to before \"{}\" — {} action{} discarded",
            action_summary,
            discarded,
            if discarded == 1 { "" } else { "s" }
        ));

        cx.notify();
        Ok(())
    }
    /// Check if a rewind is safe (history hasn't changed during preview).
    /// Returns (is_safe, discarded_count, target_summary).
    pub fn rewind_safety_check(&self) -> Option<(bool, usize, String)> {
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return None,
        };

        let current_fingerprint = self.history.fingerprint();
        let is_safe = current_fingerprint == session.history_fingerprint;
        let discarded = self.history.undo_count().saturating_sub(session.target_global_index);

        Some((is_safe, discarded, session.action_summary.clone()))
    }
    /// Show the rewind confirmation dialog (requires preview to be active).
    /// This builds the plan and presents the destructive warning.
    pub fn show_rewind_confirm(&mut self, cx: &mut Context<Self>) {
        // Must be previewing
        if !self.is_previewing() {
            self.status_message = Some("Not in preview mode".to_string());
            cx.notify();
            return;
        }

        // Build the plan
        let plan = match self.build_rewind_plan() {
            Some(p) => p,
            None => {
                self.status_message = Some("Cannot build rewind plan".to_string());
                cx.notify();
                return;
            }
        };

        // Check safety (fingerprint)
        let (is_safe, discard_count, target_summary) = match self.rewind_safety_check() {
            Some(s) => s,
            None => {
                self.status_message = Some("Cannot verify rewind safety".to_string());
                cx.notify();
                return;
            }
        };

        if !is_safe {
            self.status_message = Some("History changed during preview — please re-enter preview".to_string());
            cx.notify();
            return;
        }

        // Check preview quality - block degraded previews from hard rewind
        if let RewindPreviewState::On(ref session) = self.rewind_preview {
            if let PreviewQuality::Degraded(reason) = &session.quality {
                self.status_message = Some(format!("Rewind unavailable — preview was incomplete: {}", reason));
                cx.notify();
                return;
            }
        }

        // Extract additional context from preview session
        let (entry_id, replay_count, build_ms, fingerprint, sheet_name, location) =
            if let RewindPreviewState::On(ref session) = self.rewind_preview {
                // Get sheet name and location from the history entry
                let entry = self.history.entry_at(session.target_global_index);
                let (sheet_name, location) = if let Some(e) = entry {
                    let display = crate::history::History::to_display_entry(e, true);
                    let sheet = display.sheet_index.and_then(|i| {
                        self.wb(cx).sheet(i).map(|s| s.name.clone())
                    });
                    (sheet, display.location)
                } else {
                    (None, None)
                };

                (
                    session.entry_id,
                    session.replay_count,
                    session.build_ms,
                    session.history_fingerprint,
                    sheet_name,
                    location,
                )
            } else {
                (0, 0, 0, HistoryFingerprint::default(), None, None)
            };

        // Show the confirmation dialog with full context
        self.rewind_confirm.show(
            discard_count,
            target_summary,
            sheet_name,
            location,
            entry_id,
            replay_count,
            build_ms,
            fingerprint,
            plan,
        );
        cx.notify();
    }
    /// Confirm and execute the rewind (called from dialog Confirm button).
    pub fn confirm_rewind(&mut self, cx: &mut Context<Self>) {
        // Take the plan from dialog state
        let plan = match self.rewind_confirm.plan.take() {
            Some(p) => p,
            None => {
                self.status_message = Some("No rewind plan available".to_string());
                self.rewind_confirm.hide();
                cx.notify();
                return;
            }
        };

        // Capture audit data before consuming plan
        let audit_data = RewindAuditData {
            target_entry_id: self.rewind_confirm.target_entry_id,
            target_summary: self.rewind_confirm.target_summary.clone(),
            discarded_count: plan.discarded_count,
            replay_count: self.rewind_confirm.replay_count,
            build_ms: self.rewind_confirm.build_ms,
            fingerprint: self.rewind_confirm.fingerprint,
        };

        // Hide dialog first
        self.rewind_confirm.hide();

        // Apply the rewind
        match self.apply_rewind_plan(plan, cx) {
            Ok(()) => {
                // Success - show banner with full audit data
                self.rewind_success.show(audit_data);
            }
            Err(e) => {
                self.status_message = Some(format!("Rewind failed: {}", e));
            }
        }
        cx.notify();
    }
    /// Cancel the rewind confirmation dialog.
    pub fn cancel_rewind(&mut self, cx: &mut Context<Self>) {
        self.rewind_confirm.hide();
        cx.notify();
    }
    /// Dismiss the rewind success banner.
    pub fn dismiss_rewind_banner(&mut self, cx: &mut Context<Self>) {
        self.rewind_success.hide();
        cx.notify();
    }
    /// Copy rewind audit details to clipboard.
    pub fn copy_rewind_details(&mut self, cx: &mut Context<Self>) {
        let details = self.rewind_success.audit_details.clone();
        cx.write_to_clipboard(gpui::ClipboardItem::new_string(details));
        self.status_message = Some("Rewind details copied to clipboard".to_string());
        cx.notify();
    }

}
