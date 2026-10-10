//! The GUI's side of the session protocol: pumping bridge requests,
//! adapting session-host's handlers to GUI concerns (undo history with
//! agent attribution, view notification, the pairing dialog), and the
//! Problems-panel scan.
//!
//! The protocol logic itself lives in the visigrid-session-host crate; this
//! is only the host adapter. Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;
use serde_json::{json, Value};

use crate::app::{PairingPrompt, Problem, Spreadsheet};

pub(crate) fn plan_under_review_error() -> (String, String) {
    (
        "plan_under_review".to_string(),
        "this workbook has a plan under review; wait for the user to apply or dismiss it"
            .to_string(),
    )
}

fn review_blocked_apply_response(
    req: &crate::session_server::ApplyOpsRequest,
    current_revision: u64,
) -> crate::session_server::ApplyOpsResponse {
    let (code, message) = plan_under_review_error();
    crate::session_server::ApplyOpsResponse {
        applied: 0,
        total: req.ops.len(),
        current_revision,
        error: Some(crate::session_server::ApplyOpsError::OpFailed(
            visigrid_protocol::OpError {
                code,
                message,
                op_index: 0,
                suggestion: Some("Retry after the plan is applied or dismissed".to_string()),
            },
        )),
        warnings: Vec::new(),
    }
}

fn mutation_blocked_apply_response(
    req: &crate::session_server::ApplyOpsRequest,
    revision: u64,
) -> crate::session_server::ApplyOpsResponse {
    let mut response = review_blocked_apply_response(req, revision);
    response.error = Some(crate::session_server::ApplyOpsError::OpFailed(
        visigrid_protocol::OpError {
            code: "mutation_blocked".into(),
            message: "Finish cell editing or leave the read-only preview before applying a batch."
                .into(),
            op_index: 0,
            suggestion: None,
        },
    ));
    response
}

struct SessionPlanGuard {
    mode: crate::mode::Mode,
    entity_id: gpui::EntityId,
    cells_rev: u64,
    context: crate::scripting::ExecutionContextGenerationKey,
    started: std::time::Instant,
}

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
                // A recipe refresh replies when its background run is done
                SessionRequest::Structure {
                    op: visigrid_protocol::StructureOp::RefreshRecipeTable { table },
                    client,
                    reply,
                } => self.start_agent_refresh(table, client, reply, cx),
                SessionRequest::Structure { op, client, reply } => {
                    let outcome = self.handle_session_structure(&op, client, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::History {
                    redo,
                    steps,
                    client,
                    reply,
                } => {
                    let outcome = self.handle_session_history(redo, steps, client, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::CreatePlan { req, client, reply } => {
                    self.start_session_create_plan(req, client, reply, cx);
                }
                SessionRequest::GetPlan { req, reply } => {
                    let outcome = self.handle_session_get_plan(&req.plan_id, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::ListPlanChanges { req, reply } => {
                    let outcome = self.handle_session_list_plan_changes(&req, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::ApplyPlan { req, reply } => {
                    let outcome = self.handle_session_apply_plan(&req, cx);
                    let _ = reply.send(outcome);
                }
                SessionRequest::DismissPlan { req, client, reply } => {
                    let outcome = self.handle_session_dismiss_plan(&req, &client, cx);
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

    fn prepare_session_plan(
        &mut self,
        req: visigrid_protocol::CreatePlanMessage,
        client: String,
        cx: &mut Context<Self>,
    ) -> Result<crate::session_plan::PlanJob, crate::session_server::PlanBridgeOutcome> {
        use crate::plan_manager::{McpPlanRecord, McpPlanState};
        let producer_source = match &req.producer {
            visigrid_protocol::PlanProducerPayload::LuaScript { source } => source,
        };
        let canonical = json!({
            "expected_revision": req.expected_revision,
            "sheet": req.sheet,
            "title": req.title,
            "description": req.description,
            "producer": req.producer,
            "verification": req.verification,
        });
        let request_hash = blake3::hash(canonical.to_string().as_bytes())
            .to_hex()
            .to_string();

        match self
            .mcp_plans
            .existing_for_key(&client, &req.idempotency_key, &request_hash)
        {
            Ok(Some(record)) => {
                let plan_id = record.plan_id.clone();
                return Err(self.handle_session_get_plan(&plan_id, cx));
            }
            Err(()) => {
                return Err(plan_error(
                    "idempotency_conflict",
                    "this idempotency key was already used with a different plan payload",
                    false,
                ));
            }
            Ok(None) => {}
        }

        if let Some(active) = self.mcp_plans.active_record() {
            return Err(plan_error(
                "plan_in_progress",
                format!("plan {} is already active in this workbook", active.plan_id),
                true,
            ));
        }
        if self.review_mode.is_some() {
            return Err(plan_error(
                "plan_in_progress",
                "another proposal is already open in Review Mode",
                true,
            ));
        }
        if self.mode.is_editing() || self.is_previewing() {
            return Err(plan_error(
                "review_unavailable",
                "Finish cell editing and return to the live workbook before creating a plan",
                true,
            ));
        }
        if self.import_in_progress || self.hub_activity.is_some() {
            return Err(plan_error(
                "review_unavailable",
                "wait for the active import or Hub operation before creating a plan",
                true,
            ));
        }
        if req.idempotency_key.is_empty() || req.idempotency_key.len() > 128 {
            return Err(plan_error(
                "bad_request",
                "idempotency_key must be 1–128 bytes",
                false,
            ));
        }
        if req.title.trim().is_empty() || req.title.len() > 120 {
            return Err(plan_error(
                "bad_request",
                "title must be 1–120 bytes",
                false,
            ));
        }
        if req
            .description
            .as_ref()
            .is_some_and(|value| value.len() > 1000)
        {
            return Err(plan_error(
                "bad_request",
                "description must be at most 1000 bytes",
                false,
            ));
        }
        if producer_source.is_empty() {
            return Err(plan_error("bad_request", "script must not be empty", false));
        }
        if producer_source.len() > 262_144 {
            return Err(plan_error(
                "script_too_large",
                "script exceeds the 256 KiB limit",
                false,
            ));
        }
        if req.verification.len() > visigrid_engine::operation_plan::MAX_VERIFICATION_DEFINITIONS {
            return Err(plan_error(
                "verification_limit",
                "at most 8 verification definitions are allowed",
                false,
            ));
        }

        if self.recipe_run_in_progress
            || self.recovery_warning.is_some()
            || self.cloud_live_enabled()
        {
            return Err(plan_error(
                "review_unavailable",
                "Wait for the recipe or leave live/read-only mode before creating a plan",
                true,
            ));
        }
        let workbook = self.wb(cx);
        if workbook.revision() != req.expected_revision {
            return Err(plan_error(
                "revision_mismatch",
                format!(
                    "expected revision {}, but the workbook is at revision {}",
                    req.expected_revision,
                    workbook.revision()
                ),
                true,
            ));
        }
        if req
            .sheet
            .is_some_and(|index| index != workbook.active_sheet_index())
        {
            return Err(plan_error(
                "single_sheet_active_only",
                "Review Mode currently targets the active sheet",
                true,
            ));
        }
        let verification = plan_verification_definitions(&req.verification)
            .map_err(|message| plan_error("plan_invalid", message, false))?;
        let record = McpPlanRecord {
            plan_id: format!("pv_{}", uuid::Uuid::new_v4().simple()),
            owner: client,
            source_revision: req.expected_revision,
            request_hash,
            state: McpPlanState::Preparing,
            invalid_message: None,
            terminal_result: None,
        };
        let job = crate::session_plan::PlanJob {
            source_layout: self.structure_layout(workbook.active_sheet_id()),
            source_frozen: (self.view_state.frozen_rows, self.view_state.frozen_cols),
            context: crate::scripting::execution_context_generation_key(workbook),
            workbook: workbook.clone(),
            layout: self.table_review_layout(workbook.active_sheet_id()),
            session_id: self.session_window_id,
            blocks_delete: !self.table_view_installed
                && (self.row_view.is_sorted() || self.filter_state.is_enabled()),
            verification,
            request: req,
            record,
        };
        self.mcp_plans
            .insert(job.record.clone(), job.request.idempotency_key.clone());
        Ok(job)
    }

    fn start_session_create_plan(
        &mut self,
        req: visigrid_protocol::CreatePlanMessage,
        client: String,
        reply: crate::session_server::bridge::oneshot::Sender<
            crate::session_server::PlanBridgeOutcome,
        >,
        cx: &mut Context<Self>,
    ) {
        let job = match self.prepare_session_plan(req, client, cx) {
            Ok(job) => job,
            Err(outcome) => {
                let _ = reply.send(outcome);
                return;
            }
        };
        let guard = SessionPlanGuard {
            mode: self.mode,
            entity_id: self.workbook.entity_id(),
            cells_rev: self.cells_rev,
            context: crate::scripting::execution_context_generation_key(&job.workbook),
            started: std::time::Instant::now(),
        };
        self.status_message = Some(format!(
            "Preparing {} for {}…",
            job.request.title.trim(),
            job.record.owner
        ));
        cx.notify();
        cx.spawn(async move |this, cx| {
            let (job, result) = cx
                .background_executor()
                .spawn(async move {
                    let result = job.build();
                    (job, result)
                })
                .await;
            let mut result = Some(result);
            let _ = this.update(cx, |this, cx| {
                let outcome = this.complete_session_plan(&job, result.take().unwrap(), &guard, cx);
                let _ = reply.send(outcome);
            });
            // Release snapshots off the UI thread even if the window closed
            // and the update closure never ran.
            cx.background_executor()
                .spawn(async move {
                    drop((job, result));
                })
                .detach();
        })
        .detach();
    }

    fn complete_session_plan(
        &mut self,
        job: &crate::session_plan::PlanJob,
        result: Result<crate::session_plan::BuiltPlan, String>,
        guard: &SessionPlanGuard,
        cx: &mut Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        let unchanged = self.workbook.entity_id() == guard.entity_id
            && self.session_window_id == job.session_id
            && self.cells_rev == guard.cells_rev
            && crate::scripting::execution_context_generation_key(self.wb(cx)) == guard.context
            && visigrid_engine::operation_plan::shares_plan_source(&job.workbook, self.wb(cx))
            && self.table_review_layout(job.workbook.active_sheet_id()) == job.layout;
        let available = self.mode == guard.mode
            && self.review_mode.is_none()
            && !self.mode.is_editing()
            && !self.is_previewing()
            && !self.import_in_progress
            && self.hub_activity.is_none()
            && !self.recipe_run_in_progress
            && self.recovery_warning.is_none()
            && !self.cloud_live_enabled()
            && job.blocks_delete
                == (!self.table_view_installed
                    && (self.row_view.is_sorted() || self.filter_state.is_enabled()));
        let rejection = if guard.started.elapsed() >= job.request.host_timeout() {
            Some("plan_timeout: preparation exceeded its deadline; nothing was changed".to_string())
        } else if !unchanged || !available {
            Some(
                "plan_stale: workbook or window changed during preparation; re-preview required"
                    .to_string(),
            )
        } else {
            None
        };
        let result = if let Some(error) = rejection {
            cx.background_executor()
                .spawn(async move {
                    drop(result);
                })
                .detach();
            Err(error)
        } else {
            result
        };
        self.finish_session_plan(job, result, cx)
    }

    fn finish_session_plan(
        &mut self,
        job: &crate::session_plan::PlanJob,
        result: Result<crate::session_plan::BuiltPlan, String>,
        cx: &mut Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        use crate::plan_manager::McpPlanState;
        let id = &job.record.plan_id;
        // A dismissed preparation must never reopen Review Mode.
        if self
            .mcp_plans
            .record(id)
            .is_none_or(|record| record.state != McpPlanState::Preparing)
        {
            cx.background_executor()
                .spawn(async move {
                    drop(result);
                })
                .detach();
            return self.handle_session_get_plan(id, cx);
        }
        let mut record = job.record.clone();
        match result {
            Ok(built) => {
                let count = built.preview.cells_written;
                self.review_mode = Some(built.review);
                self.terminal.pending_result = Some(
                    crate::terminal::state::PendingResult::LuaPreview(built.preview),
                );
                self.focus_first_review_change(cx);
                record.state = McpPlanState::Ready;
                self.status_message = Some(format!(
                    "{} proposed {count} change(s). Review them before applying.",
                    record.owner
                ));
            }
            Err(message) => {
                self.status_message = Some(format!("Plan not opened: {message}"));
                record.state = McpPlanState::Invalid;
                record.invalid_message = Some(message);
            }
        }
        self.mcp_plans
            .insert(record, job.request.idempotency_key.clone());
        cx.notify();
        self.handle_session_get_plan(id, cx)
    }

    fn handle_session_get_plan(
        &self,
        plan_id: &str,
        cx: &Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        use crate::plan_manager::McpPlanState;
        let Some(record) = self.mcp_plans.record(plan_id) else {
            return plan_error("plan_not_found", "plan is unknown or expired", false);
        };
        match record.state {
            McpPlanState::Preparing => crate::session_server::PlanBridgeOutcome::success(json!({
                "plan_id": plan_id, "state": "preparing", "workbook_changed": false,
            })),
            McpPlanState::Invalid => crate::session_server::PlanBridgeOutcome::success(json!({
                "plan_id": plan_id,
                "state": "invalid",
                "source": { "revision": record.source_revision },
                "approval": { "required": true, "approved_by_gui": false },
                "problems": [{
                    "code": "plan_invalid",
                    "severity": "blocking",
                    "message": record.invalid_message.as_deref().unwrap_or("plan materialization failed"),
                }],
            })),
            McpPlanState::Applied | McpPlanState::Dismissed => {
                crate::session_server::PlanBridgeOutcome::success(
                    record.terminal_result.clone().unwrap_or_else(|| {
                        json!({
                            "plan_id": plan_id,
                            "state": record.state.as_str(),
                        })
                    }),
                )
            }
            McpPlanState::Ready => {
                let Some(prepared) = self.pending_prepared_plan(plan_id) else {
                    return plan_error(
                        "plan_not_found",
                        "prepared plan is no longer available",
                        false,
                    );
                };
                crate::session_server::PlanBridgeOutcome::success(self.plan_snapshot(prepared, cx))
            }
        }
    }

    fn handle_session_list_plan_changes(
        &self,
        req: &visigrid_protocol::ListPlanChangesMessage,
        _cx: &Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        if !(1..=500).contains(&req.limit) {
            return plan_error("page_limit", "limit must be between 1 and 500", false);
        }
        let Some(record) = self.mcp_plans.record(&req.plan_id) else {
            return plan_error("plan_not_found", "plan is unknown or expired", false);
        };
        if record.state != crate::plan_manager::McpPlanState::Ready {
            return plan_error(
                "plan_invalid",
                format!("changes are unavailable while plan is {}", record.state.as_str()),
                false,
            );
        }
        let Some(prepared) = self.pending_prepared_plan(&req.plan_id) else {
            return plan_error("plan_not_found", "prepared plan is no longer available", false);
        };
        let plan = prepared.plan();
        let filter_hash = blake3::hash(
            json!({"group": req.group, "kind": req.kind})
                .to_string()
                .as_bytes(),
        )
        .to_hex()
        .to_string();
        let offset = match req.cursor.as_deref() {
            None => 0,
            Some(cursor) => match decode_plan_cursor(
                cursor,
                &req.plan_id,
                &plan.plan_hash,
                &filter_hash,
            ) {
                Some(offset) => offset,
                None => {
                    return plan_error(
                        "invalid_cursor",
                        "cursor does not match this plan or filter",
                        false,
                    )
                }
            },
        };
        let mut indexed: Vec<_> = plan
            .changes
            .iter()
            .enumerate()
            .filter(|(_, change)| change_matches_filter(change, req))
            .collect();
        let display_positions: std::collections::HashMap<_, _> = plan
            .row_lineage
            .iter()
            .map(|row| (row.id.0, row.display_position))
            .collect();
        indexed.sort_by_key(|(index, change)| {
            (
                display_positions
                    .get(&change.row_id.0)
                    .copied()
                    .unwrap_or(usize::MAX),
                change
                    .after_coordinate
                    .or(change.before_coordinate)
                    .map(|cell| cell.col)
                    .unwrap_or(usize::MAX),
                change.cause == visigrid_engine::operation_plan::ChangeCause::Recalculated,
                *index,
            )
        });
        if offset > indexed.len() {
            return plan_error("invalid_cursor", "cursor offset is beyond the result set", false);
        }
        let end = (offset + req.limit).min(indexed.len());
        let changes: Vec<_> = indexed[offset..end]
            .iter()
            .map(|(index, change)| plan_change_json(*index, change))
            .collect();
        let next_cursor = (end < indexed.len())
            .then(|| encode_plan_cursor(&req.plan_id, &plan.plan_hash, &filter_hash, end));
        crate::session_server::PlanBridgeOutcome::success(json!({
            "plan_id": req.plan_id,
            "plan_hash": plan.plan_hash,
            "changes": changes,
            "next_cursor": next_cursor,
            "remaining": indexed.len().saturating_sub(end),
        }))
    }

    fn handle_session_apply_plan(
        &self,
        req: &visigrid_protocol::ApplyPlanMessage,
        cx: &Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        use crate::plan_manager::McpPlanState;
        if self.cloud_live_enabled() {
            return plan_error("go_collaboration_required", "Local plans cannot replace a live workbook outside its sequencer.", false);
        }
        let Some(record) = self.mcp_plans.record(&req.plan_id) else {
            return plan_error("plan_not_found", "plan is unknown or expired", false);
        };
        if req.expected_revision != record.source_revision {
            return plan_error(
                "revision_mismatch",
                format!(
                    "expected_revision must be the plan source revision {}",
                    record.source_revision
                ),
                true,
            );
        }
        match record.state {
            McpPlanState::Applied => {
                let mut result = record.terminal_result.clone().unwrap_or_else(|| json!({}));
                result["already_applied"] = json!(true);
                crate::session_server::PlanBridgeOutcome::success(result)
            }
            McpPlanState::Preparing => {
                plan_error("plan_in_progress", "the plan is still being prepared", true)
            }
            McpPlanState::Dismissed => plan_error("plan_not_found", "plan was dismissed", false),
            McpPlanState::Invalid => plan_error(
                "plan_invalid",
                record.invalid_message.as_deref().unwrap_or("plan is invalid"),
                false,
            ),
            McpPlanState::Ready => {
                let Some(prepared) = self.pending_prepared_plan(&req.plan_id) else {
                    return plan_error(
                        "plan_not_found",
                        "prepared plan is no longer available",
                        false,
                    );
                };
                let status = self.review_apply_status(prepared, cx);
                if status.is_some_and(|status| status.stale) {
                    plan_error(
                        "plan_stale",
                        "the workbook or calculation context changed after preview",
                        true,
                    )
                } else {
                    plan_error(
                        "approval_required",
                        "review the proposal in VisiGrid and click Apply or Dismiss",
                        true,
                    )
                }
            }
        }
    }

    fn handle_session_dismiss_plan(
        &mut self,
        req: &visigrid_protocol::DismissPlanMessage,
        client: &str,
        cx: &mut Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        use crate::plan_manager::McpPlanState;
        let Some(record) = self.mcp_plans.record(&req.plan_id) else {
            return plan_error("plan_not_found", "plan is unknown or expired", false);
        };
        if record.owner != client {
            return plan_error(
                "plan_owner_mismatch",
                "only the proposing client may dismiss this plan",
                false,
            );
        }
        match record.state {
            McpPlanState::Applied => {
                return plan_error("plan_already_applied", "the plan was already applied", false);
            }
            McpPlanState::Dismissed => {
                let mut result = record.terminal_result.clone().unwrap_or_else(|| json!({}));
                result["already_dismissed"] = json!(true);
                return crate::session_server::PlanBridgeOutcome::success(result);
            }
            McpPlanState::Preparing | McpPlanState::Ready | McpPlanState::Invalid => {}
        }
        let result = json!({
            "plan_id": req.plan_id,
            "state": "dismissed",
            "workbook_revision": self.workbook.read(cx).revision(),
            "workbook_changed": false,
            "already_dismissed": false,
        });
        self.mcp_plans.mark_dismissed(&req.plan_id, result.clone());
        if self
            .review_mode
            .as_ref()
            .is_some_and(|state| state.plan_id.0 == req.plan_id)
        {
            self.review_mode = None;
            self.terminal.pending_result = None;
        }
        self.status_message =
            Some("Dismissed the proposed changes without modifying the workbook.".into());
        cx.notify();
        crate::session_server::PlanBridgeOutcome::success(result)
    }

    fn pending_prepared_plan(
        &self,
        plan_id: &str,
    ) -> Option<&visigrid_engine::operation_plan::PreparedOperationPlan> {
        let crate::terminal::state::PendingResult::LuaPreview(preview) =
            self.terminal.pending_result.as_ref()?
        else {
            return None;
        };
        preview
            .prepared_plan
            .as_ref()
            .filter(|prepared| prepared.plan().id.0 == plan_id)
    }

    fn plan_snapshot(
        &self,
        prepared: &visigrid_engine::operation_plan::PreparedOperationPlan,
        cx: &Context<Self>,
    ) -> Value {
        let plan = prepared.plan();
        let sheet = self
            .workbook
            .read(cx)
            .sheet_index_by_id(plan.source_sheet_id)
            .unwrap_or(0);
        let stale = self
            .review_apply_status(prepared, cx)
            .is_some_and(|status| status.stale);
        let session_id = self
            .session_server
            .ready_info()
            .map(|(session_id, _, _)| session_id)
            .unwrap_or_else(|| plan.workbook_session_id.clone());
        json!({
            "plan_id": plan.id.0,
            "state": if stale { "stale" } else { "ready" },
            "producer": {
                "client": plan.producer.name,
                "kind": plan.producer.kind,
                "script_hash": plan.producer.source_hash,
            },
            "title": plan.title,
            "description": plan.description,
            "source": {
                "session_id": session_id,
                "sheet": sheet,
                "sheet_id": plan.source_sheet_id.raw(),
                "revision": plan.source_revision,
                "fingerprint": plan.source_fingerprint,
            },
            "execution_context": {
                "engine_semantics": plan.execution_context.engine_version,
                "functions_lua_generation": plan.execution_context.functions_generation,
                "functions_lua_hash": plan.execution_context.functions_source_hash,
                "fingerprint": plan.execution_context.hash(),
            },
            "plan_hash": plan.plan_hash,
            "preview_fingerprint": plan.preview_fingerprint,
            "determinism": plan.determinism,
            "summary": {
                "cells_changed": plan.summary.cells_changed,
                "cells_cleared": plan.summary.cells_cleared,
                "formulas_changed": plan.summary.formulas_changed,
                "formatting_changes": plan.summary.formatting_changes,
                "rows_deleted": plan.summary.rows_deleted,
                "recalculated_cells": plan.summary.recalculated_cells,
                "total_changes": plan.summary.total_changes(),
            },
            "groups": plan.groups,
            "verification": plan.verification,
            "problems": plan.problems,
            "changes": { "total": plan.changes.len(), "page_size_max": 500 },
            "approval": { "required": true, "approved_by_gui": false },
            "review": { "gui_visible": self.review_mode.is_some(), "apply_requires_gui": true },
            "result": Value::Null,
        })
    }

    /// Scan all sheets for cells whose computed value is an error.
    /// Sparse iteration over occupied cells only — cheap even on big books.
    /// Capped at 200 problems; the panel reports truncation.
    pub fn collect_problems(&self, cx: &Context<Self>) -> (Vec<Problem>, bool) {
        use visigrid_engine::formula::eval::Value;
        const CAP: usize = 200;

        let mut problems = self.pivot_problems(cx);
        let wb = self.workbook.read(cx);
        let mut truncated = false;
        'outer: for (sheet_idx, sheet) in wb.sheets().iter().enumerate() {
            let mut coords: Vec<(usize, usize)> = sheet.cells_iter().map(|(rc, _)| rc).collect();
            coords.sort_unstable();
            for (row, col) in coords {
                // Computed errors surface as Value::Error; cycle members carry
                // the "#CYCLE!" marker as their computed text (see
                // set_cycle_error), and so do files saved before #95.
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
    pub fn reveal_cell(
        &mut self,
        sheet_idx: usize,
        row: usize,
        col: usize,
        cx: &mut Context<Self>,
    ) {
        let current_idx = self.wb(cx).active_sheet_index();
        if sheet_idx != current_idx && sheet_idx < self.wb(cx).sheets().len() {
            if !self.activate_sheet(sheet_idx, cx) {
                return;
            }
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

        if self.cloud_live_enabled() {
            let mut response = mutation_blocked_apply_response(req, self.wb(cx).revision());
            response.error = Some(crate::session_server::ApplyOpsError::OpFailed(visigrid_protocol::OpError {
                code: "go_collaboration_required".into(),
                message: "The live workbook accepts sequenced Go operations; local session batches are disabled.".into(),
                op_index: 0, suggestion: Some("Use the shared workbook's Go collaboration connection.".into()),
            }));
            return response;
        }
        if self.review_mode.is_some() {
            return review_blocked_apply_response(req, self.workbook.read(cx).revision());
        }
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            if self.block_if_previewing_only(cx) || self.mode.is_editing() {
                return mutation_blocked_apply_response(req, self.wb(cx).revision());
            }
            let mut candidate = self.wb(cx).clone();
            let mut outcome = visigrid_session_host::apply_ops(&mut candidate, req);
            if outcome.response.error.is_some() || outcome.response.applied == 0 {
                return outcome.response;
            }
            let source = req
                .client
                .clone()
                .map(|client| MutationSource::Agent { client })
                .unwrap_or(MutationSource::Human);
            let commit = outcome
                .guarded_commit
                .take()
                .expect("Table batch captures guarded history");
            if let Err(message) = self.publish_table_batch(
                candidate,
                commit,
                if req.batch_name.is_empty() {
                    "Session batch".into()
                } else {
                    req.batch_name.clone()
                },
                source,
                cx,
            ) {
                let mut response = mutation_blocked_apply_response(req, self.wb(cx).revision());
                response.error = Some(crate::session_server::ApplyOpsError::OpFailed(
                    visigrid_protocol::OpError {
                        code: "table_view_unsafe".into(),
                        message,
                        op_index: req.ops.len().saturating_sub(1),
                        suggestion: None,
                    },
                ));
                return response;
            }
            outcome.response.current_revision = self.wb(cx).revision();
            if !outcome.changed_cells.is_empty() {
                self.session_server
                    .broadcast_cells(outcome.response.current_revision, outcome.changed_cells);
            }
            return outcome.response;
        }

        let source = match &req.client {
            Some(client) => MutationSource::Agent {
                client: client.clone(),
            },
            None => MutationSource::Human,
        };

        let outcome = self
            .workbook
            .update(cx, |wb, _| visigrid_session_host::apply_ops(wb, req));

        if outcome.response.error.is_none() && outcome.response.applied > 0 {
            for (sheet_idx, changes) in &outcome.value_changes {
                if !changes.is_empty() {
                    self.record_batch_from(cx,
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
                    self.record_format_from(cx,
                        *sheet_idx,
                        patches
                            .iter()
                            .map(|p| CellFormatPatch { remove_cell_on_undo: false,
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
    pub(crate) fn handle_session_structure(
        &mut self,
        op: &visigrid_protocol::StructureOp,
        client: Option<String>,
        cx: &mut Context<Self>,
    ) -> crate::session_server::StructureOutcome {
        use crate::history::MutationSource;
        use visigrid_protocol::StructureOp;

        let mut out = crate::session_server::StructureOutcome {
            revision: self.workbook.read(cx).revision(),
            sheet_count: self.workbook.read(cx).sheets().len(),
            active_sheet: self.workbook.read(cx).active_sheet_index(),
            ..Default::default()
        };


        if crate::table_filter_ui::has_table_criteria(self.wb(cx))
            && !matches!(
                op,
                StructureOp::InsertRows { .. }
                    | StructureOp::DeleteRows { .. }
                    | StructureOp::InsertCols { .. }
                    | StructureOp::DeleteCols { .. }
                    | StructureOp::RenameSheet { .. }
                    | StructureOp::AddSheet { .. }
            )
        {
            out.error = Some((
                "table_view_active".into(),
                crate::table_filter_ui::TABLE_VIEW_EDIT_MESSAGE.into(),
            ));
            return out;
        }
        if self.review_mode.is_some() {
            out.error = Some(plan_under_review_error());
            return out;
        }
        if matches!(op, StructureOp::RenameSheet { .. } | StructureOp::AddSheet { .. })
            && ((self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) || self.mode.is_editing())
        {
            out.error = Some(("invalid_op".into(), "Finish editing or reviewing before changing sheets.".into()));
            return out;
        }

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

        if self.cloud_live_enabled() {
            out.error = Some(("invalid_op".into(), "Live structural edits must use the sequencer; they are not enabled in this slice.".into()));
            return out;
        }

        // Row/column ops are active-sheet only in a GUI window.
        let active = self.workbook.read(cx).active_sheet_index();
        let target = visigrid_session_host::structure_target_sheet(op, active);
        let row_col_op = !matches!(
            op,
            StructureOp::AddSheet { .. }
                | StructureOp::RenameSheet { .. }
                | StructureOp::CreatePivot { .. }
                | StructureOp::RefreshPivot { .. }
        );
        let pivot_op = matches!(
            op,
            StructureOp::CreatePivot { .. } | StructureOp::RefreshPivot { .. }
        );
        if pivot_op && self.pivot_panel.is_some() {
            out.error = Some((
                "invalid_op".to_string(),
                "the pivot field list is open in the window — wait for the user to close it"
                    .to_string(),
            ));
            return out;
        }
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
        let history_before = self.history.canonical_entries().last().map(|entry| entry.id);
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
                let result = self.wb(cx).prepare_sheet_add(name.as_deref()).and_then(|(candidate, commit)| {
                    let description = format!("Added sheet \"{}\"", candidate.sheets().last().unwrap().name);
                    let source = client.clone().map(|client| MutationSource::Agent { client }).unwrap_or(MutationSource::Human);
                    self.publish_table_batch(candidate, commit, description.clone(), source, cx)?;
                    Ok(description)
                });
                match result {
                    Ok(description) => description,
                    Err(message) => {
                        self.suppress_repeat_capture = false;
                        out.error = Some(("invalid_op".into(), message));
                        return out;
                    }
                }
            }
            StructureOp::RenameSheet { name, .. } => {
                let old = self.workbook.read(cx).sheets()[target].name.clone();
                let new_name = name.trim().to_string();
                let id = self.workbook.read(cx).sheets()[target].id;
                let description = format!("Renamed sheet \"{}\" to \"{}\"", old, new_name);
                let result = self.wb(cx).prepare_sheet_rename(id, &old, &new_name)
                    .and_then(|(candidate, commit)| {
                        if commit.is_empty() { return Ok(()); }
                        let source = client.clone()
                            .map(|client| MutationSource::Agent { client })
                            .unwrap_or(MutationSource::Human);
                        self.publish_table_batch(candidate, commit, description.clone(), source, cx)
                    });
                if let Err(message) = result {
                    self.suppress_repeat_capture = false;
                    out.error = Some(("invalid_op".into(), message));
                    return out;
                }
                description
            }
            StructureOp::CreatePivot { .. } => {
                let resolved =
                    visigrid_session_host::resolve_create_pivot(op, self.workbook.read(cx));
                let result = match resolved {
                    Ok((source, definition)) => self.session_create_pivot(source, definition, cx),
                    Err((_, msg)) => Err(msg),
                };
                match result {
                    Ok(desc) => {
                        if let Some(client) = client.clone() {
                            self
                                .retag_last_source(cx, MutationSource::Agent { client });
                        }
                        desc
                    }
                    Err(msg) => {
                        self.suppress_repeat_capture = false;
                        out.error = Some(("invalid_op".to_string(), msg));
                        return out;
                    }
                }
            }
            // Answered by start_agent_refresh, in the background; never reaches here
            StructureOp::RefreshRecipeTable { .. } => {
                self.suppress_repeat_capture = false;
                out.error = Some(("invalid_op".to_string(), "recipe refresh was not handled".to_string()));
                return out;
            }
            StructureOp::RefreshPivot { pivot } => {
                let ids = visigrid_session_host::resolve_refresh_pivots(
                    pivot.as_deref(),
                    self.workbook.read(cx),
                );
                let mut done = Vec::new();
                let mut failure = None;
                match ids {
                    Ok(ids) => {
                        for id in ids {
                            match self.session_refresh_pivot(id, cx) {
                                Ok(d) => {
                                    if let Some(client) = client.clone() {
                                        self
                                            .retag_last_source(cx, MutationSource::Agent { client });
                                    }
                                    done.push(d);
                                }
                                Err(msg) => {
                                    failure = Some(msg);
                                    break;
                                }
                            }
                        }
                    }
                    Err((_, msg)) => failure = Some(msg),
                }
                if let Some(msg) = failure {
                    self.suppress_repeat_capture = false;
                    let msg = if done.is_empty() {
                        msg
                    } else {
                        format!("{} (already refreshed: {})", msg, done.join(", "))
                    };
                    out.error = Some(("invalid_op".to_string(), msg));
                    out.revision = self.workbook.read(cx).revision();
                    return out;
                }
                format!("Refreshed {}", done.join(", "))
            }
        };

        self.suppress_repeat_capture = false;

        if row_col_op && self.history.canonical_entries().last().map(|entry| entry.id) == history_before {
            out.error = Some((
                "invalid_op".into(),
                self.status_message
                    .clone()
                    .unwrap_or_else(|| "Structural edit was not applied.".into()),
            ));
            return out;
        }

        // Attribute the undo entry the GUI method just recorded (row/col ops
        // record one; sheet rename records its agent source at publication).
        if row_col_op {
            if let Some(client) = client {
                self
                    .retag_last_source(cx, MutationSource::Agent { client });
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

        if self.review_mode.is_some() {
            out.error = Some(plan_under_review_error());
            return out;
        }

        if self.block_if_previewing_only(cx) {
            out.error = Some((
                "history_blocked".into(),
                self.status_message.clone().unwrap_or_default(),
            ));
            return out;
        }
        for _ in 0..steps {
            let before = self.history.undo_count();
            if redo {
                if !self.history.can_redo() {
                    break;
                }
                self.redo(cx);
                if self.history.undo_count() == before {
                    out.error = Some((
                        "history_blocked".into(),
                        self.status_message.clone().unwrap_or_default(),
                    ));
                    break;
                }
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
                let description = self.history.peek_undo_description();
                self.undo(cx);
                if self.history.undo_count() == before {
                    out.error = Some((
                        "history_blocked".into(),
                        self.status_message.clone().unwrap_or_default(),
                    ));
                    break;
                }
                if let Some(description) = description {
                    out.descriptions.push(description);
                }
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
            review_capable: true,
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

fn plan_error(
    code: impl Into<String>,
    message: impl Into<String>,
    retryable: bool,
) -> crate::session_server::PlanBridgeOutcome {
    crate::session_server::PlanBridgeOutcome::error(code, message, retryable)
}

fn plan_verification_definitions(
    definitions: &[visigrid_protocol::PlanVerificationDefinition],
) -> Result<Vec<visigrid_engine::operation_plan::VerificationDefinition>, String> {
    use visigrid_engine::operation_plan::{
        CellCoordinate, CellRange, GroupId, VerificationDefinition,
    };
    definitions
        .iter()
        .map(|definition| match definition {
            visigrid_protocol::PlanVerificationDefinition::NoNewFormulaErrors { id, label } => {
                Ok(VerificationDefinition::NoNewFormulaErrors {
                    id: id.clone(),
                    label: label.clone(),
                })
            }
            visigrid_protocol::PlanVerificationDefinition::GrossMinusGroupEqualsPreview {
                id,
                label,
                source_range,
                amount_column,
                excluded_group,
                tolerance,
                currency,
            } => {
                let parse_cell = |cell: &str| {
                    crate::scripting::parse_a1(cell)
                        .map(|(row, col)| CellCoordinate {
                            row: row - 1,
                            col: col - 1,
                        })
                        .ok_or_else(|| format!("invalid A1 reference: {cell}"))
                };
                let (start, end) = match source_range.split_once(':') {
                    Some((start, end)) => (parse_cell(start)?, parse_cell(end)?),
                    None => {
                        let cell = parse_cell(source_range)?;
                        (cell, cell)
                    }
                };
                if start.row > end.row || start.col > end.col {
                    return Err(format!("source_range starts after it ends: {source_range}"));
                }
                let amount = parse_cell(&format!("{}1", amount_column.trim()))?.col;
                Ok(VerificationDefinition::RetainedTotal {
                    id: id.clone(),
                    label: label.clone(),
                    source_range: CellRange { start, end },
                    amount_column: amount,
                    excluded_group: GroupId(excluded_group.clone()),
                    tolerance: *tolerance,
                    currency: currency.clone().unwrap_or_else(|| "XXX".into()),
                })
            }
        })
        .collect()
}

fn change_matches_filter(
    change: &visigrid_engine::operation_plan::MaterializedChange,
    req: &visigrid_protocol::ListPlanChangesMessage,
) -> bool {
    use visigrid_engine::operation_plan::{ChangeCause, ChangeKind};
    if let Some(group) = req.group.as_deref() {
        if !change.group_id.as_ref().is_some_and(|id| id.0 == group) {
            return false;
        }
    }
    match req.kind {
        None => true,
        Some(visigrid_protocol::PlanChangeKindFilter::Recalculated) => {
            change.cause == ChangeCause::Recalculated
        }
        Some(visigrid_protocol::PlanChangeKindFilter::Value) => {
            change.cause == ChangeCause::Direct && change.kind == ChangeKind::Value
        }
        Some(visigrid_protocol::PlanChangeKindFilter::Formula) => {
            change.cause == ChangeCause::Direct && change.kind == ChangeKind::Formula
        }
        Some(visigrid_protocol::PlanChangeKindFilter::Clear) => {
            change.cause == ChangeCause::Direct && change.kind == ChangeKind::Cleared
        }
        Some(visigrid_protocol::PlanChangeKindFilter::RowDelete) => {
            change.cause == ChangeCause::Direct && change.kind == ChangeKind::RowDeleted
        }
    }
}

fn plan_change_json(
    index: usize,
    change: &visigrid_engine::operation_plan::MaterializedChange,
) -> Value {
    use visigrid_engine::operation_plan::{ChangeCause, ChangeKind};
    let coordinate =
        |cell: Option<visigrid_engine::operation_plan::CellCoordinate>,
         snapshot: &visigrid_engine::operation_plan::CellSnapshot| {
            cell.map(|cell| {
                json!({
                    "cell": crate::scripting::format_a1(cell.row + 1, cell.col + 1),
                    "raw": snapshot.raw,
                    "display": snapshot.display,
                })
            })
        };
    json!({
        "change_id": format!("ch_{index}"),
        "kind": match change.kind {
            ChangeKind::Value => "value",
            ChangeKind::Formula => "formula",
            ChangeKind::Cleared => "clear",
            ChangeKind::Format => "format",
            ChangeKind::RowDeleted => "row_delete",
        },
        "cause": match change.cause {
            ChangeCause::Direct => "direct",
            ChangeCause::Recalculated => "recalculated",
        },
        "row_id": format!("row_{}", change.row_id.0),
        "before": coordinate(change.before_coordinate, &change.before),
        "after": coordinate(change.after_coordinate, &change.after),
        "before_row": change.before_coordinate.map(|cell| cell.row),
        "after_row": change.after_coordinate.map(|cell| cell.row),
        "group": change.group_id.as_ref().map(|id| id.0.as_str()),
        "reason": change.reason.as_ref().map(|reason| reason.0.as_str()),
        "sources": change.sources,
    })
}

fn encode_plan_cursor(plan_id: &str, plan_hash: &str, filter_hash: &str, offset: usize) -> String {
    let body = format!("{plan_id}:{plan_hash}:{filter_hash}:{offset}");
    let signature = blake3::hash(body.as_bytes()).to_hex();
    format!("{offset}.{}", &signature[..16])
}

fn decode_plan_cursor(
    cursor: &str,
    plan_id: &str,
    plan_hash: &str,
    filter_hash: &str,
) -> Option<usize> {
    let (offset, signature) = cursor.split_once('.')?;
    let offset = offset.parse::<usize>().ok()?;
    let expected = encode_plan_cursor(plan_id, plan_hash, filter_hash, offset);
    (expected.split_once('.')?.1 == signature).then_some(offset)
}

#[cfg(test)]
mod review_block_tests {
    use super::{
        decode_plan_cursor, encode_plan_cursor, plan_verification_definitions,
        review_blocked_apply_response,
    };

    #[test]
    fn session_apply_rejection_is_explicit_and_non_mutating() {
        let request = crate::session_server::ApplyOpsRequest {
            request_id: "request-1".into(),
            batch_name: "Agent edit".into(),
            atomic: true,
            expected_revision: Some(7),
            ops: vec![visigrid_protocol::Op::SetCellValue {
                sheet: 0,
                row: 0,
                col: 0,
                value: "blocked".into(),
            }],
            client: Some("Test agent".into()),
        };
        let response = review_blocked_apply_response(&request, 7);
        assert_eq!(response.applied, 0);
        assert_eq!(response.current_revision, 7);
        assert!(matches!(
            response.error,
            Some(crate::session_server::ApplyOpsError::OpFailed(
                visigrid_protocol::OpError { ref code, .. }
            )) if code == "plan_under_review"
        ));
    }

    #[test]
    fn blocked_batch_response_is_explicit_and_non_mutating() {
        let request = crate::session_server::ApplyOpsRequest {
            request_id: "request-1".into(),
            batch_name: "Agent edit".into(),
            atomic: true,
            expected_revision: Some(7),
            ops: vec![visigrid_protocol::Op::SetCellValue {
                sheet: 0,
                row: 0,
                col: 0,
                value: "blocked".into(),
            }],
            client: Some("Test agent".into()),
        };
        let response = super::mutation_blocked_apply_response(&request, 7);
        assert_eq!(response.applied, 0);
        assert_eq!(response.total, 1);
        assert_eq!(response.current_revision, 7);
        assert!(matches!(
            response.error,
            Some(crate::session_server::ApplyOpsError::OpFailed(
                visigrid_protocol::OpError { ref code, .. }
            )) if code == "mutation_blocked"
        ));
    }

    #[test]
    fn plan_cursor_is_bound_to_plan_hash_and_filter() {
        let cursor = encode_plan_cursor("pv_one", "hash-one", "filter-one", 100);
        assert_eq!(
            decode_plan_cursor(&cursor, "pv_one", "hash-one", "filter-one"),
            Some(100)
        );
        assert_eq!(
            decode_plan_cursor(&cursor, "pv_two", "hash-one", "filter-one"),
            None
        );
        assert_eq!(
            decode_plan_cursor(&cursor, "pv_one", "hash-two", "filter-one"),
            None
        );
        assert_eq!(
            decode_plan_cursor(&cursor, "pv_one", "hash-one", "filter-two"),
            None
        );
    }

    #[test]
    fn mcp_verification_definitions_convert_a1_coordinates() {
        let definitions = plan_verification_definitions(&[
            visigrid_protocol::PlanVerificationDefinition::GrossMinusGroupEqualsPreview {
                id: "retained".into(),
                label: Some("Retained total".into()),
                source_range: "A2:C6".into(),
                amount_column: "C".into(),
                excluded_group: "duplicates".into(),
                tolerance: 0.01,
                currency: Some("USD".into()),
            },
        ])
        .unwrap();
        let visigrid_engine::operation_plan::VerificationDefinition::RetainedTotal {
            source_range,
            amount_column,
            ..
        } = &definitions[0]
        else {
            panic!("expected retained-total definition");
        };
        assert_eq!(source_range.start.row, 1);
        assert_eq!(source_range.end.row, 5);
        assert_eq!(*amount_column, 2);
    }

    #[test]
    fn plan_change_json_uses_one_based_a1_addresses() {
        use visigrid_engine::operation_plan::{
            CellCoordinate, CellSnapshot, ChangeCause, ChangeKind, MaterializedChange, ReviewRowId,
        };

        let change = MaterializedChange {
            sheet_id: visigrid_engine::sheet::SheetId(1),
            row_id: ReviewRowId(2),
            before_coordinate: Some(CellCoordinate { row: 1, col: 0 }),
            after_coordinate: Some(CellCoordinate { row: 1, col: 0 }),
            before: CellSnapshot {
                raw: "grace".into(),
                display: "grace".into(),
                format: visigrid_engine::cell::CellFormat::default(),
            },
            after: CellSnapshot {
                raw: "Grace".into(),
                display: "Grace".into(),
                format: visigrid_engine::cell::CellFormat::default(),
            },
            kind: ChangeKind::Value,
            cause: ChangeCause::Direct,
            group_id: None,
            reason: None,
            sources: Vec::new(),
        };
        let json = super::plan_change_json(0, &change);
        assert_eq!(json["before"]["cell"], "A2");
        assert_eq!(json["after"]["cell"], "A2");
    }
}

#[cfg(test)]
mod background_plan_tests {
    use super::{SessionPlanGuard, Spreadsheet};
    use crate::plan_manager::McpPlanState;
    use gpui::{AppContext, BorrowAppContext};

    fn request(revision: u64, key: &str) -> visigrid_protocol::CreatePlanMessage {
        visigrid_protocol::CreatePlanMessage {
            id: "request".into(),
            idempotency_key: key.into(),
            expected_revision: revision,
            sheet: Some(0),
            title: "Update amount".into(),
            description: None,
            producer: visigrid_protocol::PlanProducerPayload::LuaScript {
                source: "set('A2', 42)".into(),
            },
            verification: vec![],
        }
    }

    fn init(cx: &mut gpui::TestAppContext) {
        cx.update(|cx| {
            crate::settings::init_settings_store(cx);
            crate::load_embedded_fonts(cx);
            cx.set_global(crate::session::SessionManager::new());
            cx.set_global(crate::window_registry::WindowRegistry::new());
        });
    }

    #[gpui::test]
    fn background_plan_returns_before_review_and_never_writes_live_cells(
        cx: &mut gpui::TestAppContext,
    ) {
        init(cx);
        let view = cx.add_window(Spreadsheet::new);
        let (tx, rx) = crate::session_server::bridge::oneshot::channel();
        view.update(cx, |app, _, cx| {
            app.wb_mut(cx, |wb| {
                wb.set_cell_value_tracked(0, 0, 0, "Amount");
                wb.set_cell_value_tracked(0, 1, 0, "10");
                wb.set_cell_value_tracked(0, 2, 0, "20");
                let table = wb
                    .create_table(
                        wb.active_sheet_id(),
                        visigrid_engine::table::TableRange {
                            start_row: 0,
                            start_col: 0,
                            end_row: 2,
                            end_col: 0,
                        },
                        "Sales",
                    )
                    .unwrap()
                    .table_id();
                wb.set_table_source(
                    table,
                    Some(visigrid_engine::table::TableSource {
                        recipe: "orders.recipe.toml".into(),
                        refreshed: None,
                    }),
                )
                .unwrap();
                wb.set_cell_value_tracked(0, 0, 2, "=SUM(Sales[Amount])");
            });
            let req = request(app.wb(cx).revision(), "one");
            app.start_session_create_plan(req.clone(), "test".into(), tx, cx);
            assert!(
                app.review_mode.is_none(),
                "creation must yield to the UI before building"
            );
            assert_eq!(
                app.mcp_plans.active_record().unwrap().state,
                McpPlanState::Preparing
            );
            let retry = app
                .prepare_session_plan(req, "test".into(), cx)
                .err()
                .unwrap();
            assert_eq!(retry.value.unwrap()["state"], "preparing");
            assert_eq!(app.wb(cx).active_sheet().get_raw(1, 0), "10");
        })
        .unwrap();
        cx.run_until_parked();
        let result = rx
            .recv_within(std::time::Duration::from_secs(1))
            .unwrap()
            .unwrap();
        assert_eq!(result.value.unwrap()["state"], "ready");
        view.update(cx, |app, _, cx| {
            assert!(app.review_mode.is_some());
            let id = app.mcp_plans.active_plan_id().unwrap();
            let prepared = app.pending_prepared_plan(id).unwrap();
            assert_eq!(
                prepared.preview_workbook().active_sheet().get_display(0, 2),
                "62"
            );
            assert_eq!(app.wb(cx).active_sheet().get_display(0, 2), "30");
            assert_eq!(app.wb(cx).active_sheet().get_raw(1, 0), "10");
        })
        .unwrap();
    }

    #[gpui::test]
    fn background_plan_rechecks_source_window_deadline_and_dismissal(
        cx: &mut gpui::TestAppContext,
    ) {
        init(cx);
        let view = cx.add_window(Spreadsheet::new);
        for case in [
            "cell",
            "replace",
            "edit",
            "layout",
            "timeout",
            "dismiss",
            "context",
            "import",
            "other_review",
            "dialog",
        ] {
            view.update(cx, |app, _, cx| {
                app.mode = crate::mode::Mode::Navigation;
                let req = request(app.wb(cx).revision(), case);
                let job = app
                    .prepare_session_plan(req, "test".into(), cx)
                    .ok()
                    .unwrap();
                let mut guard = SessionPlanGuard {
                    mode: app.mode,
                    entity_id: app.workbook.entity_id(),
                    cells_rev: app.cells_rev,
                    context: crate::scripting::execution_context_generation_key(app.wb(cx)),
                    started: std::time::Instant::now(),
                };
                // Build the exact production job on a separate OS thread.
                let (job, built) = std::thread::spawn(move || {
                    let built = job.build();
                    (job, built)
                })
                .join()
                .unwrap();
                assert!(built.is_ok(), "{case}: {:?}", built.as_ref().err());
                match case {
                    "cell" => {
                        app.wb_mut(cx, |wb| wb.set_cell_value_tracked(0, 5, 0, "human edit"));
                    }
                    "import" => {
                        app.import_in_progress = true;
                    }
                    "other_review" => {
                        let prepared = built
                            .as_ref()
                            .unwrap()
                            .preview
                            .prepared_plan
                            .as_ref()
                            .unwrap();
                        app.review_mode = Some(crate::review_mode::ReviewModeState::from_plan(
                            prepared.plan(),
                        ));
                    }
                    "replace" => {
                        app.workbook = cx.new(|_| job.workbook.clone());
                    }
                    "edit" => {
                        app.mode = crate::mode::Mode::Edit;
                    }
                    "dialog" => {
                        app.mode = crate::mode::Mode::GoTo;
                    }
                    "layout" => {
                        app.view_state.frozen_rows += 1;
                    }
                    "timeout" => {
                        guard.started -= job.request.host_timeout();
                    }
                    "context" => {
                        app.wb_mut(cx, |wb| wb.set_auto_recalc(!wb.auto_recalc()));
                    }
                    "dismiss" => {
                        app.mcp_plans.mark_dismissed(
                            &job.record.plan_id,
                            serde_json::json!({"state": "dismissed"}),
                        );
                    }
                    _ => unreachable!(),
                }
                let result = app.complete_session_plan(&job, built, &guard, cx);
                let expected = if case == "dismiss" {
                    "dismissed"
                } else {
                    "invalid"
                };
                assert_eq!(result.value.unwrap()["state"], expected, "{case}");
                assert_eq!(app.review_mode.is_some(), case == "other_review", "{case}");
                app.review_mode = None;
                app.import_in_progress = false;
                assert!(app.mcp_plans.active_record().is_none(), "{case}");
                assert!(app.wb(cx).active_sheet().get_raw(1, 0).is_empty());
            })
            .unwrap();
        }
    }

    #[gpui::test]
    fn failed_background_plan_releases_reservation(cx: &mut gpui::TestAppContext) {
        init(cx);
        let view = cx.add_window(Spreadsheet::new);
        let (tx, rx) = crate::session_server::bridge::oneshot::channel();
        view.update(cx, |app, _, cx| {
            let mut req = request(app.wb(cx).revision(), "error");
            req.producer = visigrid_protocol::PlanProducerPayload::LuaScript {
                source: "set('A2', 42); error('stop')".into(),
            };
            app.start_session_create_plan(req, "test".into(), tx, cx);
        })
        .unwrap();
        cx.run_until_parked();
        assert_eq!(
            rx.recv_within(std::time::Duration::from_secs(1))
                .unwrap()
                .unwrap()
                .value
                .unwrap()["state"],
            "invalid"
        );
        view.update(cx, |app, _, cx| {
            assert!(app.mcp_plans.active_record().is_none());
            assert!(app.review_mode.is_none());
            assert!(app.wb(cx).active_sheet().get_raw(1, 0).is_empty());
        })
        .unwrap();
    }
    #[gpui::test]
    #[ignore = "manual large-Table timing; run with --ignored --nocapture"]
    fn background_plan_large_table_timing(cx: &mut gpui::TestAppContext) {
        init(cx);
        let view = cx.add_window(Spreadsheet::new);
        view.update(cx, |app, _, cx| {
            let rows = 300_000;
            app.wb_mut(cx, |wb| {
                wb.begin_batch();
                wb.set_cell_value_tracked(0, 0, 0, "Amount");
                for row in 1..=rows { wb.set_cell_value_tracked(0, row, 0, "10"); }
                wb.end_batch();
                let table = wb.create_table(wb.active_sheet_id(), visigrid_engine::table::TableRange {
                    start_row: 0, start_col: 0, end_row: rows, end_col: 0,
                }, "Sales").unwrap().table_id();
                wb.set_table_source(table, Some(visigrid_engine::table::TableSource {
                    recipe: "large.recipe.toml".into(), refreshed: None,
                })).unwrap();
            });
            let req = request(app.wb(cx).revision(), "timing");
            let started = std::time::Instant::now();
            let job = app.prepare_session_plan(req, "test".into(), cx).ok().unwrap();
            let capture = started.elapsed();
            let guard = SessionPlanGuard {
                mode: app.mode,
                entity_id: app.workbook.entity_id(), cells_rev: app.cells_rev,
                context: crate::scripting::execution_context_generation_key(app.wb(cx)),
                started: std::time::Instant::now(),
            };
            let (job, built, background) = std::thread::spawn(move || {
                let started = std::time::Instant::now();
                let built = job.build(); (job, built, started.elapsed())
            }).join().unwrap();
            assert!(built.is_ok(), "{:?}", built.as_ref().err());
            let started = std::time::Instant::now();
            let result = app.complete_session_plan(&job, built, &guard, cx);
            let publish = started.elapsed();
            assert_eq!(result.value.unwrap()["state"], "ready");
            assert_eq!(app.wb(cx).active_sheet().get_raw(1, 0), "10");
            eprintln!("300k-row recipe Table: capture={capture:?}, background={background:?}, publish={publish:?}");
        }).unwrap();
    }
}
