//! Soft-rewind: preview sessions (scrubbing history without committing),
//! rewind planning, safety checks, and the confirm/apply flow.
//! State types live in rewind_state.rs; this is the behavior.
//! Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::app::{Spreadsheet, NUM_ROWS};
use crate::history::HistoryFingerprint;
use crate::rewind_state::*;

pub(crate) fn preview_rows(view: Option<&PreviewSheetView>) -> visigrid_engine::filter::RowView {
    if let Some(rows) = view.and_then(|v| v.table_rows.as_ref()) { return rows.clone(); }
    let mut rows = visigrid_engine::filter::RowView::new(NUM_ROWS);
    if let Some(order) = view.and_then(|v| v.row_order.as_ref()) { rows.apply_sort(order.clone()); }
    rows
}

pub(crate) fn preview_selection(
    rows: &visigrid_engine::filter::RowView,
    start: (usize, usize), end: (usize, usize),
    table: Option<visigrid_engine::table::TableRange>,
) -> ((usize, usize), (usize, usize)) {
    let mut visible = (start.0.min(end.0)..=start.0.max(end.0)).filter_map(|r| rows.data_to_view(r));
    if let Some(first) = visible.next() {
        let (lo, hi) = visible.fold((first,first), |(lo,hi), r| (lo.min(r),hi.max(r)));
        ((lo,start.1),(hi,end.1))
    } else {
        let row = table.map(|t| t.start_row).unwrap_or(0);
        ((row,start.1),(row,start.1))
    }
}

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
    /// Block a command while a read-only preview is active.
    /// Returns true if blocked (command should return early).
    /// Sets status message with consistent preview warning.
    pub fn block_if_previewing(&mut self, cx: &mut Context<Self>) -> bool {
        if self.cloud_live_enabled() {
            self.status_message = Some("This live session supports cell values and formulas only".into());
            cx.notify(); return true;
        }
        self.block_if_previewing_only(cx) || self.block_table_view_edit(cx)
    }
    pub(crate) fn block_if_previewing_cell(&mut self, cx: &mut Context<Self>) -> bool {
        self.block_live_read_only(cx) || self.block_if_previewing_only(cx) || self.block_table_view_edit(cx)
    }

    pub(crate) fn block_if_previewing_only(&mut self, cx: &mut Context<Self>) -> bool {
        if self.block_read_only_recovery(cx) { return true; }
        if self.review_mode.is_some() {
            self.status_message =
                Some("Apply or dismiss Review Mode before editing the workbook.".to_string());
            cx.notify();
            return true;
        }
        if self.is_previewing() {
            self.status_message = Some(PREVIEW_BLOCK_MSG.to_string());
            cx.notify();
            true
        } else {
            false
        }
    }
    fn validate_preview_layout(&self, workbook: &visigrid_engine::workbook::Workbook, views: &PreviewViewState) -> Result<(), String> {
        for (sheet, view) in workbook.sheets().iter().zip(&views.per_sheet) {
            if view.table_rows.is_none() { continue; }
            let table = sheet.table_view_spec().and_then(|s| sheet.tables().iter().find(|t| t.id == s.table))
                .ok_or("The preview Table no longer exists.")?;
            if let Some(error) = crate::table_filter_ui::desktop_layout_error(
                table,
                view.structure_layout
                    .as_ref()
                    .map(|l| &l.heights)
                    .or_else(|| self.row_heights.get(&sheet.id)),
                view.structure_layout
                    .as_ref()
                    .map(|l| &l.hidden_rows)
                    .or_else(|| self.hidden_rows.get(&sheet.id)),
                sheet.frozen_panes.0,
            ) {
                return Err(error);
            }
        }
        Ok(())
    }

    pub(crate) fn install_preview_rows(&mut self) {
        let RewindPreviewState::On(session) = &self.rewind_preview else { return; };
        let view = session.view_state.per_sheet.get(session.snapshot.active_sheet_index());
        self.row_view = preview_rows(view);
        self.table_view_installed = view.is_some_and(|v| v.table_rows.is_some());
    }

    fn navigate_preview(&mut self, sheet: usize, range: Option<(usize, usize, usize, usize, usize)>, cx: &mut Context<Self>) {
        self.activate_sheet(sheet, cx);
        let table_range = self.sheet(cx).table_view_spec()
            .and_then(|s| self.sheet(cx).tables().iter().find(|t| t.id == s.table)).map(|t| t.range);
        let (r0, c0, r1, c1) = range.map(|(_, a, b, c, d)| (a,b,c,d))
            .or_else(|| table_range.map(|r| (r.start_row,r.start_col,r.start_row,r.end_col)))
            .unwrap_or((0,0,0,0));
        let (start, end) = preview_selection(&self.row_view, (r0,c0), (r1,c1), table_range);
        self.view_state.select_cell(start.0, start.1);
        if end != start { self.view_state.selection_end = Some(end); }
        self.ensure_visible(cx);
    }

    /// Enter preview mode for the currently selected history entry
    pub fn enter_preview(&mut self, cx: &mut Context<Self>) -> Result<(), String> {
        self.wb(cx).ensure_writable()?;
        if self.review_mode.is_some() {
            self.block_if_previewing(cx);
            return Err("Review Mode is active".to_string());
        }
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
        let sheet_idx = self.history.display_entries().iter().find(|e| e.id == entry_id)
            .and_then(|e| e.sheet_index).unwrap_or(self.sheet_index(cx));

        // Build the preview workbook and view state (state BEFORE this action).
        // The fallback clone is only needed when no rewind baseline was captured;
        // borrowing it beside the mutable history call would overlap fields.
        let fallback = if self.history.has_rewind_base() { None } else { self.base_workbook.clone() };
        let build_result = self.history.build_workbook_before(
            history_index,
            fallback.as_ref(),
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
            crate::history::PreviewBuildError::NoBaseSnapshot => {
                "Preview unavailable — no starting snapshot to replay from".to_string()
            }
        })?;

        self.validate_preview_layout(&build_result.workbook, &build_result.view_state)?;

        // Capture current focus for restoration
        let live_focus = PreviewFocus {
            additional_selections: self.view_state.additional_selections.clone(),
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
            live_rows: self.row_view.clone(),
            live_table_view_installed: self.table_view_installed,
            live_revision: self.wb(cx).revision(),
            history_fingerprint: self.history.fingerprint(),
            replay_count: build_result.replay_count,
            build_ms: build_result.build_ms,
            quality: PreviewQuality::Ok,
        };

        self.rewind_preview = RewindPreviewState::On(session);

        self.navigate_preview(sheet_idx, self.history_highlight_range, cx);

        self.status_message = Some(format!("Preview: Before \"{}\" — Release Space to return", action_summary));
        cx.notify();
        Ok(())
    }
    /// Exit preview mode, restoring live state
    pub fn exit_preview(&mut self, cx: &mut Context<Self>) {
        if let RewindPreviewState::On(session) = std::mem::take(&mut self.rewind_preview) {
            self.activate_sheet(session.live_focus.sheet_index, cx);
            self.row_view = session.live_rows;
            self.table_view_installed = session.live_table_view_installed;
            self.view_state.selected = session.live_focus.selected;
            self.view_state.selection_end = session.live_focus.selection_end;
            self.view_state.additional_selections = session.live_focus.additional_selections;
            self.view_state.scroll_row = session.live_focus.scroll_row;
            self.view_state.scroll_col = session.live_focus.scroll_col;
            if self.wb(cx).revision() == session.live_revision {
                self.table_view_sync_key = Some((self.sheet(cx).id, self.wb(cx).revision()));
            } else {
                self.table_view_sync_key = None;
                self.sync_table_view(cx);
            }

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

        // Build first: a refused scrub must not replace the current snapshot
        // or disturb the live projection saved when Space was first pressed.
        let fallback = if self.history.has_rewind_base() { None } else { self.base_workbook.clone() };
        match self.history.build_workbook_before(
            new_idx, fallback.as_ref(), MAX_PREVIEW_REPLAY, MAX_PREVIEW_BUILD_MS,
        ).and_then(|build| {
            self.validate_preview_layout(&build.workbook, &build.view_state)
                .map_err(crate::history::PreviewBuildError::InvariantViolation)?;
            Ok(build)
        }) {
            Ok(build) => {
                let sheet_idx = self.history.display_entries().iter().find(|e| e.id == new_id)
                    .and_then(|e| e.sheet_index).unwrap_or(self.sheet_index(cx));
                let RewindPreviewState::On(session) = &mut self.rewind_preview else { return; };
                session.entry_id = new_id;
                session.target_global_index = new_idx;
                session.action_summary = action_summary.clone();
                session.snapshot = build.workbook;
                session.view_state = build.view_state;
                session.replay_count = build.replay_count;
                session.build_ms = build.build_ms;
                self.selected_history_id = Some(new_id);
                self.history_highlight_range = new_highlight;
                self.navigate_preview(sheet_idx, new_highlight, cx);
                self.status_message = Some(format!(
                    "Preview: Before \"{}\" [{}/{}] — ↑↓ to scrub, release Space to return",
                    action_summary, new_idx + 1, history_len
                ));
            }
            Err(e) => self.status_message = Some(format!("Preview unchanged: {:?}", e)),
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
        self.wb(cx).ensure_writable()?;
        // Validate history fingerprint hasn't changed
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return Err("No preview active".to_string()),
        };

        let current_fingerprint = self.history.fingerprint();
        if current_fingerprint != session.history_fingerprint || self.wb(cx).revision() != session.live_revision {
            return Err("The workbook or history changed during preview. Re-enter preview to try again.".into());
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

        if target_entry_id != session.entry_id || plan.truncate_at != session.target_global_index {
            return Err("The preview target changed. Re-enter preview to try again.".into());
        }
        let active_projection = plan.new_workbook.active_sheet().build_saved_table_view(NUM_ROWS.min(plan.new_workbook.active_sheet().rows))?;
        let mut filters = active_projection.as_ref().map(|v| v.filters().clone()).unwrap_or_default();
        if active_projection.is_none() {
            filters.sort = plan.new_view_state.per_sheet.get(plan.new_workbook.active_sheet_index())
                .and_then(|v| v.sort).map(|(column, ascending)| visigrid_engine::filter::SortState {
                    column, direction: if ascending { visigrid_engine::filter::SortDirection::Ascending } else { visigrid_engine::filter::SortDirection::Descending },
                });
        }
        if session.quality != PreviewQuality::Ok {
            return Err("Cannot rewind an incomplete preview.".into());
        }
        self.validate_preview_layout(&plan.new_workbook, &plan.new_view_state)?;

        // === ATOMIC COMMIT: Do not fail after this point ===

        // 1. Replace the workbook content
        self.rewind_preview = RewindPreviewState::Off;
        self.workbook.update(cx, |wb, _| {
            wb.restore_snapshot_monotonic(&plan.new_workbook)
        });
        for (sheet, view) in plan
            .new_workbook
            .sheets()
            .iter()
            .zip(&plan.new_view_state.per_sheet)
        {
            if let Some(layout) = &view.structure_layout {
                self.install_structure_layout(sheet.id, layout);
            }
        }
        self.view_state.frozen_rows = plan.new_workbook.active_sheet().frozen_panes.0;
        self.view_state.frozen_cols = plan.new_workbook.active_sheet().frozen_panes.1;
        self.update_cached_sheet_id(cx); // Keep per-sheet sizing cache in sync
        self.debug_assert_sheet_cache_sync(cx); // Catch desync at rewind
        // Retained actions still replay from the original base. Replacing it
        // with the target snapshot would replay those edits twice next time.
        let active_idx = self.sheet_index(cx);
        let view = plan.new_view_state.per_sheet.get(active_idx);
        self.row_view = preview_rows(view);
        self.filter_state = filters;
        self.table_view_installed = view.is_some_and(|v| v.table_rows.is_some());
        self.table_view_sync_key = Some((self.sheet(cx).id, self.wb(cx).revision()));
        self.bump_cells_rev();
        self.clipboard_visual_range = None;

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
    pub fn rewind_safety_check(&self, cx: &App) -> Option<(bool, usize, String)> {
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return None,
        };

        let current_fingerprint = self.history.fingerprint();
        let is_safe = current_fingerprint == session.history_fingerprint && self.wb(cx).revision() == session.live_revision;
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
        let (is_safe, discard_count, target_summary) = match self.rewind_safety_check(cx) {
            Some(s) => s,
            None => {
                self.status_message = Some("Cannot verify rewind safety".to_string());
                cx.notify();
                return;
            }
        };

        if !is_safe {
            self.status_message = Some("Workbook or history changed during preview — please re-enter preview".to_string());
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
                self.exit_preview(cx);
                self.status_message = Some(format!("Rewind failed: {}", e));
            }
        }
        cx.notify();
    }
    /// Cancel the rewind confirmation dialog.
    pub fn cancel_rewind(&mut self, cx: &mut Context<Self>) {
        self.rewind_confirm.hide();
        self.exit_preview(cx);
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
