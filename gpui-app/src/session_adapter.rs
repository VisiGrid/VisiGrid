//! The GUI's side of the session protocol: pumping bridge requests,
//! adapting session-host's handlers to GUI concerns (undo history with
//! agent attribution, view notification, the pairing dialog), and the
//! Problems-panel scan.
//!
//! The protocol logic itself lives in the visigrid-session-host crate; this
//! is only the host adapter. Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::app::{PairingPrompt, Problem, Spreadsheet};

impl Spreadsheet {
    /// Create a bridge handle for the session server.
    /// The handle can be cloned and passed to the TCP server.
    pub fn session_bridge_handle(&self) -> crate::session_server::SessionBridgeHandle {
        crate::session_server::SessionBridgeHandle::new(self.session_request_tx.clone())
    }
    /// Drain pending session requests and process them.
    /// Called at the start of each render cycle.
    pub(crate) fn drain_session_requests(&mut self, cx: &mut Context<Self>) {
        use crate::session_server::{SessionRequest, SubscribeResponse, UnsubscribeResponse};

        // Non-blocking drain: process all pending requests
        while let Ok(request) = self.session_request_rx.try_recv() {
            match request {
                SessionRequest::ApplyOps { req, reply } => {
                    // Apply ops through the canonical mutation path
                    let response = self.handle_session_apply_ops(&req, cx);
                    let _ = reply.send(response);
                }
                SessionRequest::Inspect { req, reply } => {
                    let response = self.handle_session_inspect(&req, cx);
                    let _ = reply.send(response);
                }
                SessionRequest::Subscribe { req, reply } => {
                    // TODO: Implement subscription tracking
                    let _ = reply.send(SubscribeResponse {
                        topics: req.topics,
                        current_revision: self.workbook.read(cx).revision(),
                    });
                }
                SessionRequest::Structure { op, client, reply } => {
                    let outcome = self.handle_session_structure(&op, client, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::History { redo, steps, client, reply } => {
                    let outcome = self.handle_session_history(redo, steps, client, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::Save { reply, .. } => {
                    // The GUI owns its save flow (prompts, cloud sync); the
                    // protocol save op is for headless hosts.
                    let _ = reply.send(crate::session_server::SaveOutcome {
                        path: None,
                        revision: self.workbook.read(cx).revision(),
                        error: Some((
                            "save_unsupported".to_string(),
                            "this session is a GUI window — save from the app (Ctrl+S)".to_string(),
                        )),
                    });
                }
                SessionRequest::Pair { client_name, reply } => {
                    if self.pairing_prompt.is_some() {
                        // One dialog at a time; the server also gates this,
                        // but a race between two servers' requests lands here.
                        let _ = reply.send(false);
                    } else {
                        self.pairing_prompt = Some(PairingPrompt { client_name, reply: Some(reply) });
                        cx.notify();
                    }
                }
                SessionRequest::Unsubscribe { req, reply } => {
                    // TODO: Implement unsubscription
                    let _ = reply.send(UnsubscribeResponse {
                        topics: req.topics,
                    });
                }
            }
        }
    }
    /// Scan all sheets for cells whose computed value is an error.
    /// Sparse iteration over occupied cells only — cheap even on big books.
    /// Capped at 200 problems; the panel reports truncation.
    pub fn collect_problems(&self, cx: &Context<Self>) -> (Vec<Problem>, bool) {
        use visigrid_engine::formula::eval::Value;
        const CAP: usize = 200;

        let wb = self.workbook.read(cx);
        let mut problems = Vec::new();
        let mut truncated = false;
        'outer: for (sheet_idx, sheet) in wb.sheets().iter().enumerate() {
            let mut coords: Vec<(usize, usize)> =
                sheet.cells_iter().map(|(&rc, _)| rc).collect();
            coords.sort_unstable();
            for (row, col) in coords {
                // Computed errors surface as Value::Error; cycle-marked cells
                // store the literal text "#CYCLE!" instead (see set_cycle_error).
                let error = match sheet.get_computed_value(row, col) {
                    Value::Error(e) => Some(e),
                    Value::Text(t) if t == "#CYCLE!" => Some(t),
                    _ => None,
                };
                if let Some(error) = error {
                    if problems.len() >= CAP {
                        truncated = true;
                        break 'outer;
                    }
                    problems.push(Problem {
                        sheet_idx,
                        sheet_name: sheet.name.clone(),
                        row,
                        col,
                        error,
                        formula: sheet.get_raw(row, col),
                    });
                }
            }
        }
        (problems, truncated)
    }
    /// Switch to a sheet (if needed) and move the cursor to a cell,
    /// scrolling it into view. Used by the Problems panel's click-to-jump.
    pub fn reveal_cell(&mut self, sheet_idx: usize, row: usize, col: usize, cx: &mut Context<Self>) {
        let current_idx = self.wb(cx).active_sheet_index();
        if sheet_idx != current_idx && sheet_idx < self.wb(cx).sheets().len() {
            self.wb_mut(cx, |wb| wb.set_active_sheet(sheet_idx));
            self.update_cached_sheet_id(cx);
            self.active_view_state_mut().active_sheet = sheet_idx;
        }
        let view_state = self.active_view_state_mut();
        view_state.selected = (row, col);
        view_state.selection_end = None;
        view_state.additional_selections.clear();
        self.ensure_cell_visible(row, col);
        cx.notify();
    }
    /// Resolve the pending pairing dialog. `approve` = user clicked Allow.
    /// The TCP thread persists the credential and replies to the client;
    /// if it already timed out, the send fails silently — dialog just closes.
    pub fn respond_pairing(&mut self, approve: bool, cx: &mut Context<Self>) {
        if let Some(mut prompt) = self.pairing_prompt.take() {
            if let Some(reply) = prompt.reply.take() {
                let _ = reply.send(approve);
            }
            self.status_message = Some(if approve {
                format!("Paired \"{}\" — it can now control this workbook (revoke: vgrid pair --revoke)", prompt.client_name)
            } else {
                format!("Denied pairing request from \"{}\"", prompt.client_name)
            });
            cx.notify();
        }
    }
    /// Handle an apply_ops request: delegate to session-host, then record
    /// undo history from the outcome and broadcast to subscribers.
    fn handle_session_apply_ops(
        &mut self,
        req: &crate::session_server::ApplyOpsRequest,
        cx: &mut Context<Self>,
    ) -> crate::session_server::ApplyOpsResponse {
        use crate::history::{CellChange, CellFormatPatch, FormatActionKind, MutationSource};

        let source = match &req.client {
            Some(client) => MutationSource::Agent { client: client.clone() },
            None => MutationSource::Human,
        };

        let outcome = self
            .workbook
            .update(cx, |wb, _| visigrid_session_host::apply_ops(wb, req));

        if outcome.response.error.is_none() && outcome.response.applied > 0 {
            for (sheet_idx, changes) in &outcome.value_changes {
                if !changes.is_empty() {
                    self.history.record_batch_from(
                        *sheet_idx,
                        changes
                            .iter()
                            .map(|c| CellChange {
                                row: c.row,
                                col: c.col,
                                old_value: c.old_value.clone(),
                                new_value: c.new_value.clone(),
                            })
                            .collect(),
                        source.clone(),
                    );
                }
            }
            for (sheet_idx, patches) in &outcome.format_patches {
                if !patches.is_empty() {
                    self.history.record_format_from(
                        *sheet_idx,
                        patches
                            .iter()
                            .map(|p| CellFormatPatch {
                                row: p.row,
                                col: p.col,
                                before: p.before.clone(),
                                after: p.after.clone(),
                            })
                            .collect(),
                        FormatActionKind::PasteFormats,
                        req.batch_name.clone(),
                        source.clone(),
                    );
                }
            }

            self.is_modified = true;
            self.cached_title = None;

            if !outcome.changed_cells.is_empty() {
                self.session_server
                    .broadcast_cells(outcome.response.current_revision, outcome.changed_cells);
            }
        }

        outcome.response
    }
    /// Handle a structural edit from a session client.
    ///
    /// Routes through the GUI's own row/column methods rather than the
    /// engine so per-sheet view state (row view, row heights) and the undo
    /// entry are recorded exactly as for a human edit — then re-tags that
    /// entry with the agent's identity so the undo guard can tell them apart.
    /// Row/column ops only apply to the ACTIVE sheet: the view state they
    /// maintain is per-active-sheet.
    fn handle_session_structure(
        &mut self,
        op: &visigrid_protocol::StructureOp,
        client: Option<String>,
        cx: &mut Context<Self>,
    ) -> crate::session_server::StructureOutcome {
        use visigrid_protocol::StructureOp;
        use crate::history::MutationSource;

        let mut out = crate::session_server::StructureOutcome {
            revision: self.workbook.read(cx).revision(),
            sheet_count: self.workbook.read(cx).sheets().len(),
            active_sheet: self.workbook.read(cx).active_sheet_index(),
            ..Default::default()
        };

        // Shared validation (bounds, counts, name clashes).
        {
            let wb = self.workbook.read(cx);
            if let Some((code, message, suggestion)) =
                visigrid_session_host::validate_structure_op(op, wb)
            {
                let msg = match suggestion {
                    Some(s) => format!("{} — {}", message, s),
                    None => message,
                };
                out.error = Some((code.to_string(), msg));
                return out;
            }
        }

        // Row/column ops are active-sheet only in a GUI window.
        let active = self.workbook.read(cx).active_sheet_index();
        let target = visigrid_session_host::structure_target_sheet(op, active);
        let row_col_op = !matches!(op, StructureOp::AddSheet { .. } | StructureOp::RenameSheet { .. });
        if row_col_op && target != active {
            out.error = Some((
                "invalid_op".to_string(),
                format!(
                    "row and column edits apply to the active sheet ({}) in a GUI window; sheet {} is not active",
                    self.workbook.read(cx).sheets()[active].name, target
                ),
            ));
            return out;
        }

        // Row/column ops route through the GUI's own methods (so view state
        // and undo stay right), and those methods record the F4 repeat slot.
        // An agent's insert must not become what the user's F4 repeats.
        self.suppress_repeat_capture = true;
        let description = match op {
            StructureOp::InsertRows { at, count, .. } => {
                self.insert_rows(*at, *count, cx);
                format!("Inserted {} row(s) at row {}", count, at + 1)
            }
            StructureOp::DeleteRows { at, count, .. } => {
                self.delete_rows(*at, *count, cx);
                format!("Deleted {} row(s) at row {}", count, at + 1)
            }
            StructureOp::InsertCols { at, count, .. } => {
                self.insert_cols(*at, *count, cx);
                format!("Inserted {} column(s) at column {}", count, at + 1)
            }
            StructureOp::DeleteCols { at, count, .. } => {
                self.delete_cols(*at, *count, cx);
                format!("Deleted {} column(s) at column {}", count, at + 1)
            }
            StructureOp::AddSheet { name } => {
                let idx = self.wb_mut(cx, |wb| match name {
                    Some(n) => wb.add_sheet_named(n.trim()).unwrap_or_else(|| wb.add_sheet()),
                    None => wb.add_sheet(),
                });
                let sheet_name = self.workbook.read(cx).sheets()[idx].name.clone();
                self.is_modified = true;
                cx.notify();
                format!("Added sheet \"{}\"", sheet_name)
            }
            StructureOp::RenameSheet { name, .. } => {
                let old = self.workbook.read(cx).sheets()[target].name.clone();
                let new_name = name.trim().to_string();
                self.wb_mut(cx, |wb| { wb.rename_sheet(target, &new_name); });
                self.is_modified = true;
                cx.notify();
                format!("Renamed sheet \"{}\" to \"{}\"", old, new_name)
            }
        };

        self.suppress_repeat_capture = false;

        // Attribute the undo entry the GUI method just recorded (row/col ops
        // record one; sheet ops record none, matching the GUI's own behavior).
        if row_col_op {
            if let Some(client) = client {
                self.history.retag_last_source(MutationSource::Agent { client });
            }
        }

        out.description = description;
        out.revision = self.workbook.read(cx).revision();
        out.sheet_count = self.workbook.read(cx).sheets().len();
        out.active_sheet = self.workbook.read(cx).active_sheet_index();
        out
    }
    /// Handle an undo/redo request from a session client.
    ///
    /// SAFETY RULE: an agent may only revert entries it (or another session
    /// client) authored. If the next entry on the stack is a human edit, we
    /// refuse — "the agent can undo its own mistakes, never your work".
    /// Redo is unrestricted: it only re-applies what was just undone.
    fn handle_session_history(
        &mut self,
        redo: bool,
        steps: u32,
        client: Option<String>,
        cx: &mut Context<Self>,
    ) -> crate::session_server::HistoryOutcome {
        use crate::history::MutationSource;

        let mut out = crate::session_server::HistoryOutcome {
            revision: self.workbook.read(cx).revision(),
            can_undo: self.history.can_undo(),
            can_redo: self.history.can_redo(),
            ..Default::default()
        };

        for _ in 0..steps {
            if redo {
                if !self.history.can_redo() {
                    break;
                }
                self.redo(cx);
                out.applied += 1;
            } else {
                if !self.history.can_undo() {
                    break;
                }
                // Human-edit guard: stop before reverting the user's work.
                if matches!(self.history.peek_undo_source(), Some(MutationSource::Human)) {
                    if out.applied == 0 {
                        let who = client.as_deref().unwrap_or("this client");
                        out.error = Some((
                            "history_blocked".to_string(),
                            format!(
                                "the next undo step is a change the user made ({}); {} may only undo its own edits",
                                self.history.peek_undo_description().unwrap_or_else(|| "manual edit".into()),
                                who
                            ),
                        ));
                    }
                    break;
                }
                if let Some(desc) = self.history.peek_undo_description() {
                    out.descriptions.push(desc);
                }
                self.undo(cx);
                out.applied += 1;
            }
        }

        out.revision = self.workbook.read(cx).revision();
        out.can_undo = self.history.can_undo();
        out.can_redo = self.history.can_redo();
        out
    }
    /// Handle an inspect request: delegate to session-host.
    fn handle_session_inspect(
        &self,
        req: &crate::session_server::InspectRequest,
        cx: &Context<Self>,
    ) -> crate::session_server::InspectResponse {
        visigrid_session_host::inspect(
            self.workbook.read(cx),
            req,
            &self.document_meta.display_name,
        )
    }
    /// Start the session server with the given mode.
    ///
    /// If `token_override` is provided (e.g. from VISIGRID_SESSION_TOKEN env var),
    /// uses that token instead of generating a fresh one. This allows test harnesses
    /// to know the token in advance.
    pub fn start_session_server(
        &mut self,
        mode: crate::session_server::ServerMode,
        token_override: Option<String>,
        cx: &mut Context<Self>,
    ) -> std::io::Result<()> {
        // Waker: the bridge pings this channel after enqueueing a request and
        // this task drains immediately. Without it, requests are only drained
        // at the start of a render frame — an unfocused window renders no
        // frames, so the bridge would sit dead until the user interacts.
        let (waker_tx, waker_rx) = smol::channel::unbounded::<()>();
        cx.spawn(async move |this, cx| {
            while waker_rx.recv().await.is_ok() {
                while waker_rx.try_recv().is_ok() {} // coalesce bursts
                let alive = this.update(cx, |app, cx| {
                    app.drain_session_requests(cx);
                    cx.notify();
                });
                if alive.is_err() {
                    break; // entity dropped
                }
            }
        })
        .detach();

        let bridge = crate::session_server::SessionBridgeHandle::new_with_waker(
            self.session_request_tx.clone(),
            waker_tx,
        );
        let workbook_path = self.current_file.clone();
        let workbook_title = self.document_meta.display_name.clone();

        self.session_server.start(crate::session_server::SessionServerConfig {
            mode,
            workbook_path,
            workbook_title,
            bridge: Some(bridge),
            token_override,
            ..Default::default()
        })
    }
    /// Get structured READY info for CI output.
    pub fn session_server_ready_info(&self) -> Option<(String, u16, std::path::PathBuf)> {
        self.session_server.ready_info()
    }
    /// Stop the session server.
    pub fn stop_session_server(&mut self) {
        self.session_server.stop();
    }
}
