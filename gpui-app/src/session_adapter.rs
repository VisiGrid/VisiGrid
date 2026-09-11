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

fn plan_under_review_error() -> (String, String) {
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
                    let outcome = self.handle_session_create_plan(req, client, cx);
                    let _ = reply.send(outcome);
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

    fn handle_session_create_plan(
        &mut self,
        req: visigrid_protocol::CreatePlanMessage,
        client: String,
        cx: &mut Context<Self>,
    ) -> crate::session_server::PlanBridgeOutcome {
        use crate::plan_manager::{McpPlanRecord, McpPlanState};
        use crate::terminal::state::{LuaPreviewData, PendingResult};
        use visigrid_engine::operation_plan::{PlanId, PlanProducer};

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
                return self.handle_session_get_plan(&plan_id, cx);
            }
            Err(()) => {
                return plan_error(
                    "idempotency_conflict",
                    "this idempotency key was already used with a different plan payload",
                    false,
                );
            }
            Ok(None) => {}
        }

        if let Some(active) = self.mcp_plans.active_record() {
            return plan_error(
                "plan_in_progress",
                format!("plan {} is already active in this workbook", active.plan_id),
                true,
            );
        }
        if self.review_mode.is_some() {
            return plan_error(
                "plan_in_progress",
                "another proposal is already open in Review Mode",
                true,
            );
        }
        if self.import_in_progress || self.hub_activity.is_some() {
            return plan_error(
                "review_unavailable",
                "wait for the active import or Hub operation before creating a plan",
                true,
            );
        }
        if req.idempotency_key.is_empty() || req.idempotency_key.len() > 128 {
            return plan_error("bad_request", "idempotency_key must be 1–128 bytes", false);
        }
        if req.title.trim().is_empty() || req.title.len() > 120 {
            return plan_error("bad_request", "title must be 1–120 bytes", false);
        }
        if req.description.as_ref().is_some_and(|value| value.len() > 1000) {
            return plan_error(
                "bad_request",
                "description must be at most 1000 bytes",
                false,
            );
        }
        if producer_source.is_empty() {
            return plan_error("bad_request", "script must not be empty", false);
        }
        if producer_source.len() > 262_144 {
            return plan_error("script_too_large", "script exceeds the 256 KiB limit", false);
        }
        if req.verification.len() > visigrid_engine::operation_plan::MAX_VERIFICATION_DEFINITIONS {
            return plan_error(
                "verification_limit",
                "at most 8 verification definitions are allowed",
                false,
            );
        }

        let (active_sheet, source_fingerprint, snapshot) = {
            let workbook = self.workbook.read(cx);
            let actual_revision = workbook.revision();
            if actual_revision != req.expected_revision {
                return plan_error(
                    "revision_mismatch",
                    format!(
                        "expected revision {}, but the workbook is at revision {}",
                        req.expected_revision, actual_revision
                    ),
                    true,
                );
            }
            let active_sheet = workbook.active_sheet_index();
            if req.sheet.is_some_and(|sheet| sheet != active_sheet) {
                return plan_error(
                    "single_sheet_active_only",
                    format!("Review Mode currently targets active sheet {active_sheet}"),
                    true,
                );
            }
            (
                active_sheet,
                crate::app::sheet_fingerprint(workbook.active_sheet()),
                crate::scripting::SheetSnapshot::from_sheet(workbook.active_sheet()),
            )
        };
        let verification = match plan_verification_definitions(&req.verification) {
            Ok(value) => value,
            Err(message) => return plan_error("plan_invalid", message, false),
        };
        let plan_id = format!("pv_{}", uuid::Uuid::new_v4().simple());
        let script_hash = blake3::hash(producer_source.as_bytes()).to_hex().to_string();
        let source_sheet_index = active_sheet;

        let result = self
            .lua_runtime
            .eval_with_sheet(producer_source, Box::new(snapshot));
        let invalid = |message: String| McpPlanRecord {
            plan_id: plan_id.clone(),
            owner: client.clone(),
            source_revision: req.expected_revision,
            request_hash: request_hash.clone(),
            state: McpPlanState::Invalid,
            invalid_message: Some(message),
            terminal_result: None,
        };
        if let Some(message) = result.error {
            self.mcp_plans.insert(invalid(message), req.idempotency_key);
            return self.handle_session_get_plan(&plan_id, cx);
        }
        if result
            .ops
            .iter()
            .any(|op| matches!(op, crate::scripting::LuaOp::DeleteRows { .. }))
            && (self.row_view.is_sorted() || self.filter_state.is_enabled())
        {
            self.mcp_plans.insert(
                invalid(
                    "unsupported_view_state: clear the active sort/filter before reviewing row deletion"
                        .into(),
                ),
                req.idempotency_key,
            );
            return self.handle_session_get_plan(&plan_id, cx);
        }

        let prepared = crate::ai_actions::prepare_lua_operation_plan_with_metadata(
            self.workbook.read(cx),
            self.session_window_id,
            PlanId(plan_id.clone()),
            PlanProducer {
                kind: "mcp_lua".into(),
                name: client.clone(),
                source_path: None,
                source_hash: Some(script_hash.clone()),
            },
            req.title.trim().to_string(),
            req.description.clone(),
            &script_hash,
            &result.ops,
            verification,
        );
        let prepared = match prepared.and_then(crate::ai_actions::require_visible_plan_changes) {
            Ok(plan) => plan,
            Err(message) => {
                self.mcp_plans.insert(invalid(message), req.idempotency_key);
                return self.handle_session_get_plan(&plan_id, cx);
            }
        };

        let cells_written = prepared.plan().summary.total_changes();
        let cells_overwritten = self
            .workbook
            .read(cx)
            .sheet(source_sheet_index)
            .map(|sheet| crate::app::count_lua_overwrites(&result.ops, sheet))
            .unwrap_or(0);
        self.review_mode = Some(crate::review_mode::ReviewModeState::from_prepared(
            &prepared,
            self.workbook.read(cx),
        ));
        self.terminal.pending_result = Some(PendingResult::LuaPreview(LuaPreviewData {
            script_path: std::path::PathBuf::from(format!("mcp/{plan_id}.lua")),
            script_hash,
            ops: result.ops,
            prepared_plan: Some(prepared),
            cells_written,
            cells_overwritten,
            source_sheet_index,
            source_fingerprint,
            output: result.output,
            error: None,
        }));
        self.mcp_plans.insert(
            McpPlanRecord {
                plan_id: plan_id.clone(),
                owner: client.clone(),
                source_revision: req.expected_revision,
                request_hash,
                state: McpPlanState::Ready,
                invalid_message: None,
                terminal_result: None,
            },
            req.idempotency_key,
        );
        self.status_message = Some(format!(
            "{client} proposed {cells_written} change(s). Review them before applying."
        ));
        cx.notify();
        self.handle_session_get_plan(&plan_id, cx)
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
            McpPlanState::Ready | McpPlanState::Invalid => {}
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

        let wb = self.workbook.read(cx);
        let mut problems = Vec::new();
        let mut truncated = false;
        'outer: for (sheet_idx, sheet) in wb.sheets().iter().enumerate() {
            let mut coords: Vec<(usize, usize)> = sheet.cells_iter().map(|(&rc, _)| rc).collect();
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

        if self.review_mode.is_some() {
            return review_blocked_apply_response(req, self.workbook.read(cx).revision());
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
        use crate::history::MutationSource;
        use visigrid_protocol::StructureOp;

        let mut out = crate::session_server::StructureOutcome {
            revision: self.workbook.read(cx).revision(),
            sheet_count: self.workbook.read(cx).sheets().len(),
            active_sheet: self.workbook.read(cx).active_sheet_index(),
            ..Default::default()
        };

        if self.review_mode.is_some() {
            out.error = Some(plan_under_review_error());
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

        // Row/column ops are active-sheet only in a GUI window.
        let active = self.workbook.read(cx).active_sheet_index();
        let target = visigrid_session_host::structure_target_sheet(op, active);
        let row_col_op = !matches!(
            op,
            StructureOp::AddSheet { .. } | StructureOp::RenameSheet { .. }
        );
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
                    Some(n) => wb
                        .add_sheet_named(n.trim())
                        .unwrap_or_else(|| wb.add_sheet()),
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
                self.wb_mut(cx, |wb| {
                    wb.rename_sheet(target, &new_name);
                });
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
                self.history
                    .retag_last_source(MutationSource::Agent { client });
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
