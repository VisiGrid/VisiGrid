//! AI and Lua actions on Spreadsheet: generating agent context files,
//! Explain-with-AI, and capturing/previewing/applying AI-authored Lua.
//! Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::ai_cli::{detect_ai_cli, project_ai_context_template, system_ai_context, ALL_AI_CLIS};
use crate::settings::{update_user_settings, user_settings, TipId};
use crate::app::{count_lua_overwrites, extract_last_lua_block, save_ai_lua_script, sheet_fingerprint, Spreadsheet};

pub(crate) fn lua_preview_source_matches(
    workbook: &visigrid_engine::workbook::Workbook,
    source_sheet_index: usize,
    source_fingerprint: u64,
) -> bool {
    workbook
        .sheet(source_sheet_index)
        .map(sheet_fingerprint)
        == Some(source_fingerprint)
}

pub(crate) fn prepare_lua_operation_plan_with_metadata(
    workbook: &visigrid_engine::workbook::Workbook,
    session_window_id: u64,
    plan_id: visigrid_engine::operation_plan::PlanId,
    producer: visigrid_engine::operation_plan::PlanProducer,
    title: String,
    description: Option<String>,
    _script_hash: &str,
    ops: &[crate::scripting::LuaOp],
    mut verification: Vec<visigrid_engine::operation_plan::VerificationDefinition>,
) -> Result<visigrid_engine::operation_plan::PreparedOperationPlan, String> {
    use visigrid_engine::operation_plan::{OperationPlanRequest, PreparedOperationPlan};

    let parts = crate::scripting::lua_journal_to_plan(ops)?;
    verification.extend(parts.verification);
    let planned_ops = parts.operations;
    let context = crate::scripting::execution_context_fingerprint(workbook, &planned_ops);
    PreparedOperationPlan::materialize(workbook, OperationPlanRequest {
        id: plan_id,
        workbook_session_id: session_window_id.to_string(),
        source_sheet_id: workbook.active_sheet_id(),
        expected_revision: workbook.revision(),
        execution_context: context,
        producer,
        title,
        description,
        operations: planned_ops,
        groups: parts.groups,
        verification,
    }).map_err(|error| error.to_string())
}

fn prepare_lua_operation_plan(
    workbook: &visigrid_engine::workbook::Workbook,
    session_window_id: u64,
    script_path: &std::path::Path,
    script_hash: &str,
    ops: &[crate::scripting::LuaOp],
    verification: Vec<visigrid_engine::operation_plan::VerificationDefinition>,
) -> Result<visigrid_engine::operation_plan::PreparedOperationPlan, String> {
    use visigrid_engine::operation_plan::{PlanId, PlanProducer};
    prepare_lua_operation_plan_with_metadata(
        workbook,
        session_window_id,
        PlanId(format!("pv_{}", uuid::Uuid::new_v4().simple())),
        PlanProducer {
            kind: "lua".into(),
            name: "Lua Console".into(),
            source_path: Some(script_path.to_string_lossy().into_owned()),
            source_hash: Some(script_hash.to_string()),
        },
        "Review Lua changes".into(),
        None,
        script_hash,
        ops,
        verification,
    )
}

fn default_lua_verification() -> Vec<visigrid_engine::operation_plan::VerificationDefinition> {
    vec![
        visigrid_engine::operation_plan::VerificationDefinition::NoNewFormulaErrors {
            id: "no_new_errors".into(),
            label: Some("No new formula errors".into()),
        },
    ]
}

pub(crate) fn require_visible_plan_changes(
    plan: visigrid_engine::operation_plan::PreparedOperationPlan,
) -> Result<visigrid_engine::operation_plan::PreparedOperationPlan, String> {
    if plan.plan().changes.is_empty() {
        Err("empty_plan: the script produced no visible workbook changes".into())
    } else {
        Ok(plan)
    }
}

#[cfg(test)]
mod review_plan_tests {
    use super::{default_lua_verification, prepare_lua_operation_plan, require_visible_plan_changes};
    use visigrid_engine::operation_plan::{
        ChangeKind, DeterminismClass, ProblemSeverity, VerificationEvidence,
        VerificationStatus,
    };
    use visigrid_engine::workbook::Workbook;

    fn transaction_fixture_workbook() -> Workbook {
        let mut workbook = Workbook::new();
        for (row, values) in [
            (0, ["Transaction", "Vendor", "Amount"]),
            (1, ["tx-001", "Amazon.com", "100"]),
            (2, ["tx-002", "AMZN", "200"]),
            (3, ["tx-001", "Amazon.com", "100"]),
            (4, ["", "", ""]),
            (5, ["tx-003", "Acme", "50"]),
            (6, ["", "Total", "=SUM(C2:C4)"]),
        ] {
            for (col, value) in values.into_iter().enumerate() {
                if !value.is_empty() {
                    workbook.set_cell_value_tracked(0, row, col, value);
                }
            }
        }
        workbook
    }

    fn prepare_fixture(
        workbook: &Workbook,
        path: &str,
        source: &str,
    ) -> visigrid_engine::operation_plan::PreparedOperationPlan {
        let runtime = crate::scripting::LuaRuntime::new().unwrap();
        let snapshot = crate::scripting::SheetSnapshot::from_sheet(workbook.active_sheet());
        let result = runtime.eval_with_sheet(source, Box::new(snapshot));
        assert!(result.error.is_none(), "fixture error: {:?}", result.error);

        prepare_lua_operation_plan(
            workbook,
            42,
            std::path::Path::new(path),
            "fixture-hash",
            &result.ops,
            default_lua_verification(),
        )
        .unwrap()
    }

    #[test]
    fn desktop_lua_adapter_evaluates_fixture_requested_verification() {
        let mut workbook = transaction_fixture_workbook();
        let prepared = prepare_fixture(
            &workbook,
            "fixtures/review_mode/transaction_cleanup.lua",
            include_str!("../../fixtures/review_mode/transaction_cleanup.lua"),
        );
        let retained = prepared
            .plan()
            .verification
            .iter()
            .find(|result| result.id == "retained_payments")
            .expect("fixture assertion should reach the desktop adapter");
        assert_eq!(retained.status, VerificationStatus::Passed);
        assert!(matches!(
            &retained.evidence,
            VerificationEvidence::RetainedTotal { expected, actual, currency, .. }
                if expected == "350" && actual == "350" && currency == "USD"
        ));

        let mut review = crate::review_mode::ReviewModeState::from_prepared(&prepared, &workbook);
        assert_eq!(review.review_item_count(), review.navigation_len(false));
        let first_focused = review.first_navigable_cell().unwrap();
        assert_eq!(review.focused_cell(), Some(first_focused));
        assert_eq!(review.review_position(first_focused.0, first_focused.1), Some(0));
        let original_halo = review.halo_generation();
        review.focus_cell(5, 1);
        assert_eq!(review.focused_cell(), Some((5, 1)));
        assert_ne!(review.halo_generation(), original_halo);
        assert!(!review.collapsed);
        review.toggle_collapsed();
        assert!(review.collapsed);
        review.toggle_collapsed();
        assert!(!review.collapsed);
        review.begin_card_drag((120.0, 240.0), (100.0, 100.0), 120.0);
        assert!(review.card_is_dragging());
        review.drag_card(
            (900.0, 900.0),
            120.0,
            (500.0, 400.0),
            (360.0, 310.0),
        );
        assert_eq!(review.card_position(), Some((132.0, 82.0)));
        review.end_card_drag();
        assert!(!review.card_is_dragging());
        let vendor_change = review.change_index_at_source(1, 1).unwrap();
        let vendor_change = &prepared.plan().changes[vendor_change];
        assert_eq!(vendor_change.kind, ChangeKind::Value);
        review.set_endpoint(crate::review_mode::ReviewEndpoint::Before);
        let (before_sheet, before_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 1)
            .unwrap();
        assert_eq!(
            before_sheet.get_formatted_display(before_row, 1),
            "Amazon.com"
        );
        review.set_endpoint(crate::review_mode::ReviewEndpoint::After);
        let (after_sheet, after_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 1)
            .unwrap();
        assert_eq!(
            after_sheet.get_formatted_display(after_row, 1),
            "Amazon"
        );
        let (after_sheet, shifted_formula_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 6)
            .unwrap();
        assert_eq!(shifted_formula_row, 4);
        assert_eq!(after_sheet.get_raw(shifted_formula_row, 2), "=SUM(C2:C4)");
        let (deleted_source, deleted_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 3)
            .unwrap();
        assert_eq!(deleted_row, 3);
        assert_eq!(deleted_source.get_formatted_display(deleted_row, 0), "tx-001");
        assert!(review.is_deleted_source_row(3));
        assert!(review.is_deleted_source_row(4));
        assert!(review
            .overview_buckets()
            .iter()
            .any(|bucket| bucket.change_count > 0));
        assert!(review.overview_buckets()[review.bucket_for_source_row(3)].has_deleted_row);
        let first = *review.navigable_cells().first().unwrap();
        let last = *review.navigable_cells().last().unwrap();
        assert_eq!(review.adjacent_source_change(first.0, first.1, false), None);
        assert_eq!(review.adjacent_source_change(last.0, last.1, true), None);
        assert!(review.adjacent_source_group(0, 0, true).is_some());
        let visible_after_hidden_vendor = review
            .adjacent_visible_source_change(0, 0, true, false, |row| row != 1)
            .unwrap();
        assert_ne!(visible_after_hidden_vendor.0, 1);
        assert!(review
            .adjacent_visible_source_change(0, 0, true, false, |_| false)
            .is_none());

        let execution_context = crate::scripting::execution_context_fingerprint(
            &workbook,
            &prepared.plan().operations,
        );
        let ready = review.apply_status(&prepared, &workbook);
        assert!(ready.can_apply());
        assert_eq!(review.apply_status(&prepared, &workbook), ready);
        let mut changed_context = execution_context.clone();
        changed_context.timezone.push_str("-changed");
        let stale = crate::review_mode::ReviewApplyStatus::evaluate(
            &prepared,
            &workbook,
            &changed_context,
        );
        assert!(stale.stale);
        assert!(!stale.can_apply());
        assert_eq!(stale.disabled_reason(), Some("Plan is stale · re-preview required"));

        let unchanged_revision = workbook.revision();
        let original_tolerance = workbook.iterative_tolerance();
        workbook.set_iterative_tolerance(0.01);
        assert_eq!(workbook.revision(), unchanged_revision);
        assert!(review.apply_status(&prepared, &workbook).stale);
        workbook.set_iterative_tolerance(original_tolerance);
        assert_eq!(workbook.revision(), unchanged_revision);
        assert!(
            review.apply_status(&prepared, &workbook).stale,
            "staleness must remain latched after execution context is restored"
        );

        workbook.set_cell_value_tracked(0, 1, 1, "live drift");
        assert!(review.apply_status(&prepared, &workbook).stale);
        review.set_endpoint(crate::review_mode::ReviewEndpoint::Before);
        let (before_sheet, before_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 1)
            .unwrap();
        assert_eq!(before_sheet.get_formatted_display(before_row, 1), "Amazon.com");
        review.set_endpoint(crate::review_mode::ReviewEndpoint::After);
        let (after_sheet, after_row) = review
            .endpoint_sheet_row(&prepared, prepared.plan().source_sheet_id, 1)
            .unwrap();
        assert_eq!(after_sheet.get_formatted_display(after_row, 1), "Amazon");
    }

    #[test]
    fn review_apply_status_disables_conditional_plans_without_override() {
        let workbook = Workbook::new();
        let runtime = crate::scripting::LuaRuntime::new().unwrap();
        let snapshot = crate::scripting::SheetSnapshot::from_sheet(workbook.active_sheet());
        let result = runtime.eval_with_sheet(
            "sheet:set_formula(2, 1, '=TODAY()')",
            Box::new(snapshot),
        );
        assert!(result.error.is_none(), "fixture error: {:?}", result.error);
        let prepared = prepare_lua_operation_plan(
            &workbook,
            42,
            std::path::Path::new("conditional.lua"),
            "conditional-hash",
            &result.ops,
            default_lua_verification(),
        )
        .unwrap();
        assert_eq!(prepared.plan().determinism, DeterminismClass::Conditional);
        let context = crate::scripting::execution_context_fingerprint(
            &workbook,
            &prepared.plan().operations,
        );
        let status = crate::review_mode::ReviewApplyStatus::evaluate(
            &prepared,
            &workbook,
            &context,
        );
        assert!(!status.can_apply());
        assert_eq!(status.determinism_label(), "Conditional");
        assert_eq!(
            status.disabled_reason(),
            Some("Conditional plan · explicit override is not available yet")
        );
    }

    #[test]
    fn dogfood_failure_fixture_blocks_apply_with_expected_evidence() {
        let workbook = transaction_fixture_workbook();
        let prepared = prepare_fixture(
            &workbook,
            "fixtures/review_mode/transaction_cleanup_verification_failure.lua",
            include_str!(
                "../../fixtures/review_mode/transaction_cleanup_verification_failure.lua"
            ),
        );
        let retained = prepared
            .plan()
            .verification
            .iter()
            .find(|result| result.id == "retained_payments")
            .unwrap();
        assert_eq!(retained.status, VerificationStatus::Failed);
        assert!(matches!(
            &retained.evidence,
            VerificationEvidence::RetainedTotal { expected, actual, currency, .. }
                if expected == "350" && actual == "300" && currency == "USD"
        ));
        assert!(prepared
            .plan()
            .problems
            .iter()
            .any(|problem| problem.severity == ProblemSeverity::Blocking));
    }

    #[test]
    fn filtered_navigation_fixture_remains_value_only() {
        let workbook = transaction_fixture_workbook();
        let prepared = prepare_fixture(
            &workbook,
            "fixtures/review_mode/filtered_value_changes.lua",
            include_str!("../../fixtures/review_mode/filtered_value_changes.lua"),
        );
        assert_eq!(prepared.plan().changes.len(), 2);
        assert!(prepared
            .plan()
            .changes
            .iter()
            .all(|change| change.kind == ChangeKind::Value));
        assert_eq!(
            prepared
                .plan()
                .changes
                .iter()
                .map(|change| change.before_coordinate.unwrap().row)
                .collect::<Vec<_>>(),
            vec![1, 5]
        );
    }

    #[test]
    fn large_dogfood_fixture_spans_the_thirty_thousand_row_sheet() {
        let workbook = transaction_fixture_workbook();
        let prepared = prepare_fixture(
            &workbook,
            "fixtures/review_mode/large_sparse_review.lua",
            include_str!("../../fixtures/review_mode/large_sparse_review.lua"),
        );
        assert_eq!(prepared.plan().changes.len(), 4);
        assert_eq!(
            prepared
                .plan()
                .changes
                .iter()
                .map(|change| change.before_coordinate.unwrap().row)
                .collect::<Vec<_>>(),
            vec![1, 10_001, 20_001, 29_999]
        );
    }

    #[test]
    fn overview_marks_new_formula_errors() {
        let workbook = Workbook::new();
        let runtime = crate::scripting::LuaRuntime::new().unwrap();
        let snapshot = crate::scripting::SheetSnapshot::from_sheet(workbook.active_sheet());
        let result = runtime.eval_with_sheet(
            "sheet:set_formula(2, 1, '=1/0')",
            Box::new(snapshot),
        );
        assert!(result.error.is_none(), "fixture error: {:?}", result.error);
        let prepared = prepare_lua_operation_plan(
            &workbook,
            42,
            std::path::Path::new("error.lua"),
            "error-hash",
            &result.ops,
            default_lua_verification(),
        )
        .unwrap();
        let review = crate::review_mode::ReviewModeState::from_plan(prepared.plan());
        let bucket = review.bucket_for_source_row(1);
        assert!(review.overview_buckets()[bucket].has_new_formula_error);
    }

    #[test]
    fn no_op_script_cannot_enter_review_mode() {
        let workbook = Workbook::new();
        let runtime = crate::scripting::LuaRuntime::new().unwrap();
        let snapshot = crate::scripting::SheetSnapshot::from_sheet(workbook.active_sheet());
        let result = runtime.eval_with_sheet("sheet:set('A1', '')", Box::new(snapshot));
        assert!(result.error.is_none());
        let prepared = prepare_lua_operation_plan(
            &workbook,
            42,
            std::path::Path::new("no-op.lua"),
            "no-op-hash",
            &result.ops,
            default_lua_verification(),
        )
        .unwrap();
        let error = require_visible_plan_changes(prepared).unwrap_err();
        assert!(error.starts_with("empty_plan:"));
    }
}

fn plan_row_state_after_apply(
    row_view: &visigrid_engine::filter::RowView,
    row_heights: &std::collections::HashMap<usize, f32>,
    operations: &[visigrid_engine::operation_plan::PlannedOperation],
) -> (visigrid_engine::filter::RowView, std::collections::HashMap<usize, f32>) {
    let mut after_view = row_view.clone();
    let mut after_heights = row_heights.clone();
    for planned in operations {
        let visigrid_engine::operation_plan::PlannedOp::DeleteRows { at, count } = &planned.operation else {
            continue;
        };
        for row in (0..*count).rev() {
            after_view.delete_row(at + row);
        }
        let shifted: Vec<_> = after_heights.iter()
            .filter(|(row, _)| **row >= at + count)
            .map(|(row, height)| (*row, *height))
            .collect();
        after_heights.retain(|row, _| *row < *at);
        for (row, height) in shifted {
            after_heights.insert(row - count, height);
        }
    }
    (after_view, after_heights)
}

impl Spreadsheet {
    /// Generate AI context files for all supported CLIs.
    ///
    /// Creates two layers:
    /// - System: `~/.config/visigrid/ai/{CLAUDE.md, AGENTS.md, GEMINI.md}` (always)
    /// - Project: `<workbook_dir>/.visigrid/{CLAUDE.md, AGENTS.md, GEMINI.md}` (if saved)
    /// - Root stubs: `<workbook_dir>/{CLAUDE.md, AGENTS.md, GEMINI.md}` (if saved + writable)
    pub fn generate_ai_context_files(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use std::fs;

        let mut created = Vec::new();
        let mut skipped = 0u32;
        let mut project_error: Option<String> = None;

        // 1. System-level files (always created, workbook-independent)
        let system_dir = Self::system_ai_config_dir();
        if let Err(e) = fs::create_dir_all(&system_dir) {
            self.status_message = Some(format!("Failed to create {}: {}", system_dir.display(), e));
            cx.notify();
            return;
        }
        let system_content = system_ai_context();
        for cli in ALL_AI_CLIS {
            let path = system_dir.join(cli.context_filename());
            if !path.exists() {
                if let Err(e) = fs::write(&path, system_content) {
                    log::warn!("Failed to write {}: {}", path.display(), e);
                } else {
                    created.push(format!("~/.config/visigrid/ai/{}", cli.context_filename()));
                }
            } else {
                skipped += 1;
            }
        }

        // 2. Project-level files (only if workbook is saved)
        if let Some(file_path) = &self.current_file.clone() {
            if let Some(workbook_dir) = file_path.parent() {
                let project_dir = workbook_dir.join(".visigrid");
                match fs::create_dir_all(&project_dir) {
                    Err(e) => {
                        crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::Error { category: "permission_denied" });
                        project_error = Some(format!(
                            "Cannot write project context in {}: {}. Use Save As to move the workbook.",
                            workbook_dir.display(), e
                        ));
                    }
                    Ok(()) => {
                        let wb_name = self.document_meta.display_name.clone();
                        let project_content = project_ai_context_template(&wb_name);
                        for cli in ALL_AI_CLIS {
                            let path = project_dir.join(cli.context_filename());
                            if !path.exists() {
                                if let Err(e) = fs::write(&path, &project_content) {
                                    log::warn!("Failed to write {}: {}", path.display(), e);
                                    if project_error.is_none() {
                                        project_error = Some(format!(
                                            "Cannot write to {}: {} (permission denied?)",
                                            project_dir.display(), e
                                        ));
                                    }
                                } else {
                                    created.push(format!(".visigrid/{}", cli.context_filename()));
                                }
                            } else {
                                skipped += 1;
                            }
                        }

                        // 3. Root stubs — tiny pointers so CLIs discover context at repo root
                        for cli in ALL_AI_CLIS {
                            let root_path = workbook_dir.join(cli.context_filename());
                            if !root_path.exists() {
                                let stub = format!(
                                    "<!-- This is a stub generated by VisiGrid. Edit .visigrid/{f} instead. -->\n\
                                     <!-- {name} reads this file for project context. -->\n\
                                     <!-- Do not put secrets here — this file may be committed to version control. -->\n\n\
                                     See [.visigrid/{f}](.visigrid/{f}) for project instructions.\n",
                                    f = cli.context_filename(),
                                    name = cli.display_name(),
                                );
                                if let Err(e) = fs::write(&root_path, &stub) {
                                    log::warn!("Failed to write root stub {}: {}", root_path.display(), e);
                                } else {
                                    created.push(cli.context_filename().to_string());
                                }
                            } else {
                                skipped += 1;
                            }
                        }
                    }
                }
            }
        } else {
            // Unsaved workbook: system files created, project files skipped with clear message
            project_error = Some(
                "Save the workbook first to create project-level context files.".into()
            );
        }

        // Build status message
        let mut msg = if !created.is_empty() {
            format!(
                "Created {}: {}",
                if created.len() == 1 { "1 file".to_string() } else { format!("{} files", created.len()) },
                created.join(", ")
            )
        } else if skipped > 0 {
            "AI context files already exist.".into()
        } else {
            String::new()
        };

        if let Some(err) = project_error {
            if !msg.is_empty() {
                msg.push_str(". ");
            }
            msg.push_str(&err);
        }

        if msg.is_empty() {
            msg = "No files created.".into();
        }

        self.status_message = Some(msg);
        cx.notify();
    }
    /// Open the project-level AI context folder in the system file manager.
    /// Falls back to the system-level folder if no workbook is saved.
    /// If the file manager can't open, copies path to clipboard as fallback.
    pub fn open_ai_context_folder(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use std::fs;

        let folder = if let Some(file_path) = &self.current_file {
            if let Some(workbook_dir) = file_path.parent() {
                let project_dir = workbook_dir.join(".visigrid");
                if !project_dir.exists() {
                    let _ = fs::create_dir_all(&project_dir);
                }
                project_dir
            } else {
                Self::system_ai_config_dir()
            }
        } else {
            Self::system_ai_config_dir()
        };

        // Use platform file manager
        #[cfg(target_os = "linux")]
        let cmd = "xdg-open";
        #[cfg(target_os = "macos")]
        let cmd = "open";
        #[cfg(target_os = "windows")]
        let cmd = "explorer";

        match std::process::Command::new(cmd).arg(&folder).spawn() {
            Ok(_) => {
                self.status_message = Some(format!("Opened {}", folder.display()));
            }
            Err(_) => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::Error { category: "open_folder_failed" });
                // Fallback: copy path to clipboard
                let path_str = folder.display().to_string();
                cx.write_to_clipboard(gpui::ClipboardItem::new_string(path_str.clone()));
                self.status_message = Some(format!(
                    "Could not open file manager. Path copied to clipboard: {}", path_str
                ));
            }
        }
        cx.notify();
    }
    fn system_ai_config_dir() -> std::path::PathBuf {
        dirs::config_dir()
            .unwrap_or_else(|| std::path::PathBuf::from("."))
            .join("visigrid")
            .join("ai")
    }
    /// Open the AI metrics JSON file in the system file manager.
    /// Falls back to copying the path to clipboard if the file manager fails.
    pub fn open_ai_metrics(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        // Flush any buffered metrics first so the file is up to date
        crate::ai_metrics::flush();

        let path = crate::ai_metrics::metrics_path();
        if !path.exists() {
            self.status_message = Some("No AI metrics recorded yet.".into());
            cx.notify();
            return;
        }

        #[cfg(target_os = "linux")]
        let cmd = "xdg-open";
        #[cfg(target_os = "macos")]
        let cmd = "open";
        #[cfg(target_os = "windows")]
        let cmd = "explorer";

        match std::process::Command::new(cmd).arg(&path).spawn() {
            Ok(_) => {
                self.status_message = Some(format!("Opened {}", path.display()));
            }
            Err(_) => {
                let path_str = path.display().to_string();
                cx.write_to_clipboard(gpui::ClipboardItem::new_string(path_str.clone()));
                self.status_message = Some(format!(
                    "Could not open file. Path copied to clipboard: {}", path_str
                ));
            }
        }
        cx.notify();
    }
    /// Orchestration: launch AI CLI (if needed), paste full context, and append an analysis prompt.
    ///
    /// This is the "instant value" command — one keystroke from data to insight.
    pub fn ai_explain_selection(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let pref = crate::settings::user_settings(cx).terminal.preferred_ai_cli;

        // Check if AI CLI is available before doing anything
        if detect_ai_cli(pref).is_none() {
            crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::Error { category: "no_cli" });
            self.status_message = Some(
                "No AI CLI found. Install one: npm i -g @anthropic-ai/claude-code | @openai/codex | @google/gemini-cli".into()
            );
            cx.notify();
            return;
        }

        // Check if we have a non-empty selection worth explaining
        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let ((top, left), (bottom, right)) = self.selection_range();
        let has_data = (top..=bottom).any(|r| {
            (left..=right).any(|c| !sheet.get_display(r, c).is_empty())
        });
        drop(wb);

        if !has_data {
            self.status_message = Some("Select cells with data first.".into());
            cx.notify();
            return;
        }

        crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::ExplainSelection);

        // Ensure AI is running in terminal (launches if not already)
        let ai_already_running = self.terminal.visible && self.terminal.term.is_some();
        if !ai_already_running {
            // Snapshot the output epoch before launch so we can detect new output
            let epoch_before = self.terminal.output_epoch();
            self.launch_ai_terminal(window, cx);
            // Wait for the CLI to produce output (detected via epoch change).
            // O(1) check — no grid lock. Timeout after 3s.
            let entity = cx.entity().downgrade();
            cx.spawn(async move |_, cx| {
                let start = std::time::Instant::now();
                let timeout = std::time::Duration::from_secs(3);
                let poll_interval = std::time::Duration::from_millis(100);

                loop {
                    smol::Timer::after(poll_interval).await;

                    let ready = entity.update(cx, |this: &mut Self, _cx| {
                        this.terminal.output_epoch() != epoch_before
                    }).unwrap_or(false);

                    if ready || start.elapsed() >= timeout {
                        break;
                    }
                }

                let _ = entity.update(cx, |this: &mut Self, cx| {
                    this.paste_full_context_with_prompt(cx);
                });
            }).detach();
        } else {
            // Terminal already running — paste immediately
            self.paste_full_context_with_prompt(cx);
        }
    }
    /// Write the AI context file (`.visigrid/CLAUDE.md`) with system context,
    /// sheet metadata, and optional extra instructions (e.g., Lua API contract).
    ///
    /// Returns the (sheet_name, headers, selection TSV) for the caller to paste
    /// into the terminal as visible user context.
    pub(crate) fn write_ai_context_and_collect_data(
        &mut self,
        cx: &mut Context<Self>,
        extra_instructions: &str,
    ) -> Option<(String, Vec<String>, String, usize, usize, usize, usize)> {
        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let sheet_name = sheet.name.clone();
        let (used_start_row, used_start_col, used_end_row, used_end_col) =
            crate::ai::find_used_range(sheet);

        // Detect headers
        let mut header_row = used_start_row;
        for r in used_start_row..(used_start_row + 20).min(used_end_row + 1) {
            let row_vals: Vec<String> = (used_start_col..=used_end_col)
                .map(|c| sheet.get_display(r, c))
                .collect();
            if crate::ai::looks_like_header_row(&row_vals) {
                header_row = r;
                break;
            }
        }
        let headers: Vec<String> = (used_start_col..=used_end_col)
            .map(|c| sheet.get_display(header_row, c))
            .filter(|s| !s.is_empty())
            .collect();

        // Selection TSV (capped)
        let ((top_row, left_col), (bottom_row, right_col)) = self.selection_range();
        let sel_rows = bottom_row - top_row + 1;
        let sel_cols = right_col - left_col + 1;
        let capped_rows = sel_rows.min(Self::PASTE_MAX_ROWS);
        let capped_cols = sel_cols.min(Self::PASTE_MAX_COLS);

        let mut lines = Vec::new();
        let mut total_chars = 0usize;
        for r in top_row..top_row + capped_rows {
            let mut cells = Vec::new();
            for c in left_col..left_col + capped_cols {
                cells.push(sheet.get_display(r, c));
            }
            let line = cells.join("\t");
            total_chars += line.len() + 1;
            if total_chars > Self::PASTE_MAX_CHARS {
                break;
            }
            lines.push(line);
        }
        drop(wb);

        let tsv = lines.join("\n");

        // Write .visigrid/CLAUDE.md with full context (instructions stay out of terminal)
        let workspace = self.terminal.workspace_root.clone()
            .or_else(|| self.current_file.as_ref().and_then(|p| p.parent().map(|d| d.to_path_buf())));
        if let Some(ref ws) = workspace {
            let visigrid_dir = ws.join(".visigrid");
            let _ = std::fs::create_dir_all(&visigrid_dir);
            let context_path = visigrid_dir.join("CLAUDE.md");

            let mut context = String::new();
            context.push_str(system_ai_context());
            context.push('\n');

            // Sheet metadata
            context.push_str("## Current Session\n\n");
            if let Some(p) = &self.current_file {
                context.push_str(&format!("- **File**: {}\n", p.display()));
            }
            context.push_str(&format!(
                "- **Sheet**: \"{}\"  Used range: {}x{}\n",
                sheet_name, used_end_row + 1, used_end_col + 1,
            ));
            if !headers.is_empty() {
                context.push_str(&format!(
                    "- **Headers** (row {}): {}\n",
                    header_row + 1, headers.join(", ")
                ));
            }
            context.push_str(&format!(
                "- **Selection**: {}x{} starting at {}\n",
                sel_rows, sel_cols, crate::ai::cell_ref(top_row, left_col),
            ));

            // Extra instructions (Lua contract, analysis hints, etc.)
            if !extra_instructions.is_empty() {
                context.push('\n');
                context.push_str(extra_instructions);
            }

            if let Err(e) = std::fs::write(&context_path, &context) {
                log::warn!("Failed to write AI context file {}: {}", context_path.display(), e);
            }
        }

        Some((sheet_name, headers, tsv, sel_rows, sel_cols, used_end_row, used_end_col))
    }
    /// Open the pending result as a sheet AND send it to the AI CLI for explanation.
    pub fn explain_structured_result(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let pref = crate::settings::user_settings(cx).terminal.preferred_ai_cli;
        if detect_ai_cli(pref).is_none() {
            self.status_message = Some(
                "No AI CLI found. Install one: npm i -g @anthropic-ai/claude-code | @openai/codex | @google/gemini-cli".into()
            );
            cx.notify();
            return;
        }

        // Build explanation prompt from the pending result before consuming it.
        // Hard cap: 8000 chars to avoid blowing up the AI CLI.
        const EXPLAIN_MAX_CHARS: usize = 8000;
        const EXPLAIN_MAX_ROWS: usize = 20;

        let result = match self.terminal.pending_result {
            Some(crate::terminal::state::PendingResult::Structured(ref r)) => r,
            _ => return,
        };
        let description = result.description();
        let prompt = match result {
            crate::structured_results::StructuredResult::Diff { raw } => {
                let summary = raw.get("summary")
                    .map(|s| serde_json::to_string_pretty(s).unwrap_or_default())
                    .unwrap_or_default();
                let results = raw.get("results").and_then(|r| r.as_array());
                let n_results = results.map(|a| a.len()).unwrap_or(0);
                // Include first N diff rows as compact JSON
                let mut diff_sample = String::new();
                if let Some(rows) = results {
                    for row in rows.iter().take(EXPLAIN_MAX_ROWS) {
                        let line = serde_json::to_string(row).unwrap_or_default();
                        if diff_sample.len() + line.len() > EXPLAIN_MAX_CHARS / 2 { break; }
                        diff_sample.push_str(&line);
                        diff_sample.push('\n');
                    }
                    if n_results > EXPLAIN_MAX_ROWS {
                        diff_sample.push_str(&format!("# ... {} more changes\n", n_results - EXPLAIN_MAX_ROWS));
                    }
                }
                format!(
                    "# VisiGrid structured result: {}\n\
                     # Type: diff\n\
                     # Summary: {}\n\
                     # {} change(s) detected\n\n\
                     {}\n\
                     Explain this diff result. What changed and why might it matter? \
                     Are there anomalies or patterns worth investigating?\n",
                    description, summary.trim(), n_results, diff_sample
                )
            }
            crate::structured_results::StructuredResult::Peek { columns, rows } => {
                // Include column names + first N rows, capped by chars
                let mut tsv = columns.join("\t");
                tsv.push('\n');
                for row in rows.iter().take(EXPLAIN_MAX_ROWS) {
                    let cells: Vec<String> = row.iter()
                        .map(|v| crate::structured_results::json_cell_value_pub(v))
                        .collect();
                    let line = cells.join("\t");
                    if tsv.len() + line.len() > EXPLAIN_MAX_CHARS / 2 { break; }
                    tsv.push_str(&line);
                    tsv.push('\n');
                }
                if rows.len() > EXPLAIN_MAX_ROWS {
                    tsv.push_str(&format!("# ... {} more rows\n", rows.len() - EXPLAIN_MAX_ROWS));
                }
                format!(
                    "# VisiGrid structured result: {}\n\
                     # Type: peek ({} cols x {} rows)\n\n\
                     {}\n\
                     Summarize what this data shows. Detect anomalies, outliers, or patterns. \
                     Suggest formulas or next steps.\n",
                    description, columns.len(), rows.len(), tsv
                )
            }
            crate::structured_results::StructuredResult::Calc { value } => {
                let mut display = serde_json::to_string_pretty(value).unwrap_or_default();
                if display.len() > EXPLAIN_MAX_CHARS / 2 {
                    display.truncate(EXPLAIN_MAX_CHARS / 2);
                    display.push_str("\n# ... truncated");
                }
                format!(
                    "# VisiGrid structured result: {}\n\
                     # Type: calc\n\
                     # Value:\n{}\n\n\
                     Explain this calculation result. Is the value expected? \
                     What does it mean in context?\n",
                    description, display
                )
            }
        };

        // Open the result as a sheet
        self.open_structured_result_inner(cx);

        // Launch AI if not already running, then paste prompt
        let ai_already_running = self.terminal.visible && self.terminal.term.is_some();
        if !ai_already_running {
            let epoch_before = self.terminal.output_epoch();
            self.launch_ai_terminal(window, cx);
            let entity = cx.entity().downgrade();
            cx.spawn(async move |_, cx| {
                let start = std::time::Instant::now();
                let timeout = std::time::Duration::from_secs(3);
                let poll_interval = std::time::Duration::from_millis(100);
                loop {
                    smol::Timer::after(poll_interval).await;
                    let ready = entity.update(cx, |this: &mut Self, _cx| {
                        this.terminal.output_epoch() != epoch_before
                    }).unwrap_or(false);
                    if ready || start.elapsed() >= timeout { break; }
                }
                let _ = entity.update(cx, |this: &mut Self, cx| {
                    this.write_to_pty_bracketed(&prompt, cx);
                    this.status_message = Some("Sent result to AI for explanation.".into());
                    cx.notify();
                });
            }).detach();
        } else {
            self.write_to_pty_bracketed(&prompt, cx);
            self.status_message = Some("Sent result to AI for explanation.".into());
            cx.notify();
        }
    }
    /// Orchestrator: launch AI CLI, paste full VisiGrid context, and instruct
    /// the AI to output a Lua model in a fenced ```lua block.
    ///
    /// After the AI responds, the user runs "AI: Capture Last Lua Block" to preview.
    pub fn build_model_with_lua(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        // Launch AI CLI (reuse existing logic — opens terminal, starts CLI)
        let pref = crate::settings::user_settings(cx).terminal.preferred_ai_cli;
        if detect_ai_cli(pref).is_none() {
            self.status_message = Some(
                "No AI CLI found. Install one: npm i -g @anthropic-ai/claude-code | @openai/codex | @google/gemini-cli".into()
            );
            cx.notify();
            return;
        }

        // Snapshot epoch before launch so we can detect when the CLI produces output.
        let epoch_before = self.terminal.output_epoch();
        let ai_already_running = self.terminal.visible && self.terminal.term.is_some();

        self.launch_ai_terminal(window, cx);

        if ai_already_running {
            // Terminal already running — paste context + contract immediately
            self.paste_lua_build_context(cx);
        } else {
            // Wait for output epoch bump (CLI boot), then paste. Timeout 3s.
            let entity = cx.entity().downgrade();
            cx.spawn(async move |_, cx| {
                let start = std::time::Instant::now();
                let timeout = std::time::Duration::from_secs(3);
                let poll_interval = std::time::Duration::from_millis(100);

                loop {
                    smol::Timer::after(poll_interval).await;
                    let ready = entity.update(cx, |this: &mut Self, _cx| {
                        this.terminal.output_epoch() != epoch_before
                    }).unwrap_or(false);
                    if ready || start.elapsed() >= timeout {
                        break;
                    }
                }

                let _ = entity.update(cx, |this: &mut Self, cx| {
                    this.paste_lua_build_context(cx);
                });
            }).detach();
        }

        crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::LaunchAi {
            cli: "build_model_lua",
        });
        self.status_message = Some(
            "Building Lua model… AI will output a ```lua block. Then run \"AI: Capture Last Lua Block\".".into()
        );
        cx.notify();
    }
    /// Paste VisiGrid context + Lua contract into the terminal.
    fn paste_lua_build_context(&mut self, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            return;
        }

        // Lua API contract + rules — written to .visigrid/CLAUDE.md, not pasted
        let lua_instructions = "\
## Task: Build Lua Model

Output ONLY a single ```lua fenced code block. No other text or code.

### VisiGrid Lua Sheet API (1-indexed)

```
sheet:set_value(row, col, value)
sheet:set_formula(row, col, \"=...\")
sheet:clear(row, col) or sheet:clear(\"A1:C3\")
sheet:delete_rows(at, count)
sheet:review({ group=\"id\", title=\"...\", reason=\"...\", sources={\"A1:C2\"} })
sheet:verify({ id=\"retained\", kind=\"gross_minus_group_equals_preview\", source_range=\"A2:C100\", amount_column=\"C\", excluded_group=\"duplicates\", tolerance=0.01, currency=\"USD\" })
sheet:get_value(row, col)
sheet:rows()
sheet:cols()
```

### Rules

- Read inputs from the current sheet. Write outputs to new rows/columns.
- Do NOT delete or overwrite existing data.
- No `os`, `io`, file, or network operations.
- The user will preview your code before applying it.
- Use sheet:review(...) before related mutations when a plain-language reason
  and source references would help the user review them.
- Row numbers always refer to the source sheet. A write below deleted rows is
  shifted with those deletions in the materialized preview.
- sheet:verify(...) requests an engine-evaluated assertion; never claim or
  calculate its pass/fail result in Lua.
";

        let Some((sheet_name, headers, tsv, sel_rows, sel_cols, ..)) =
            self.write_ai_context_and_collect_data(cx, lua_instructions)
        else {
            return;
        };

        // Paste a visible prompt (seen by the user AND sent to the AI CLI).
        // The Lua API contract and full instructions are in .visigrid/CLAUDE.md
        // which the AI CLI auto-discovers.
        let mut block = String::new();
        block.push_str(&format!(
            "# VisiGrid: Build Lua Model\n# Sheet: \"{}\" | Selection: {}x{}",
            sheet_name, sel_rows, sel_cols,
        ));
        if !headers.is_empty() {
            block.push_str(&format!(" | Headers: {}", headers.join(", ")));
        }
        block.push_str("\n# Lua API docs and rules are in .visigrid/CLAUDE.md\n");
        if !tsv.trim().is_empty() {
            block.push_str(&tsv);
            block.push('\n');
        }
        block.push_str("\nBuild a Lua model for this data. Output a single ```lua block.\n");

        // Write directly — not bracketed paste. The AI CLI is the foreground
        // process and doesn't interpret bracketed paste escapes (they'd show
        // as literal "200~" / "201~").
        self.terminal.write_to_pty(block.as_bytes());
    }
    /// Capture the last ` ```lua ` block from terminal scrollback, preview it, and
    /// store as a PendingResult::LuaPreview for the affordance bar.
    pub fn capture_ai_lua(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use crate::terminal::state::{LuaPreviewData, PendingResult};
        use crate::scripting::SheetSnapshot;

        if self.block_review_entry_for_workbook_transition(cx) { return; }

        // Guard: terminal must be running
        let Some(ref term_arc) = self.terminal.term else {
            self.status_message = Some("No terminal session.".into());
            cx.notify();
            return;
        };

        // Extract recent terminal text
        let text = crate::terminal::extract::extract_recent_text(term_arc, 2000);

        // Find last ```lua block
        let Some(code) = extract_last_lua_block(&text) else {
            // Check if last.lua exists for a helpful hint
            let has_last = self.terminal.workspace_root.as_ref()
                .map(|r| r.join("ai").join("generated").join("last.lua").exists())
                .unwrap_or(false);
            let hint = if has_last {
                " Try \"AI: Preview last.lua\" to re-preview the previous script."
            } else {
                ""
            };
            self.status_message = Some(format!(
                "No ```lua block found. Make sure the AI outputs a ```lua fenced block.{}",
                hint
            ));
            cx.notify();
            return;
        };

        // Save script to disk
        let script_path = save_ai_lua_script(&code, &self.terminal.workspace_root);

        // Compute blake3 hash (first 16 bytes hex = 32 hex chars)
        let hash = blake3::hash(code.as_bytes());
        let script_hash = hash.to_hex()[..32].to_string();

        if self.mode.is_editing() {
            self.cancel_edit(cx);
        }
        // Snapshot current sheet for preview
        let sheet = self.sheet(cx);
        let source_sheet_index = self.sheet_index(cx);
        let source_fingerprint = sheet_fingerprint(sheet);
        let snapshot = SheetSnapshot::from_sheet(sheet);

        // Run Lua in sandboxed preview mode (no mutation)
        let result = self.lua_runtime.eval_with_sheet(&code, Box::new(snapshot));

        // Count overwrites
        let cells_overwritten = self.workbook.read(cx).sheet(source_sheet_index)
            .map(|s| count_lua_overwrites(&result.ops, s))
            .unwrap_or(0);

        // Handle replace-pending: if there's already a pending result, toast
        if self.terminal.pending_result.is_some() {
            self.status_message = Some("Previous pending result replaced.".into());
        }

        let mut preview_error = result.error;
        let prepared_plan = if preview_error.is_none() {
            if result.ops.iter().any(|op| matches!(op, crate::scripting::LuaOp::DeleteRows { .. }))
                && (self.row_view.is_sorted() || self.filter_state.is_enabled())
            {
                preview_error = Some(
                    "unsupported_view_state: clear the active sort/filter before reviewing row deletion".into(),
                );
                None
            } else {
                match prepare_lua_operation_plan(
                    self.workbook.read(cx),
                    self.session_window_id,
                    &script_path,
                    &script_hash,
                    &result.ops,
                    default_lua_verification(),
                ) {
                    Ok(plan) => match require_visible_plan_changes(plan) {
                        Ok(plan) => Some(plan),
                        Err(error) => {
                            preview_error = Some(error);
                            None
                        }
                    },
                    Err(error) => {
                        preview_error = Some(error);
                        None
                    }
                }
            }
        } else {
            None
        };
        let cells_written = prepared_plan.as_ref()
            .map(|prepared| prepared.plan().summary.total_changes())
            .unwrap_or(result.mutations);
        self.review_mode = prepared_plan
            .as_ref()
            .map(|prepared| {
                crate::review_mode::ReviewModeState::from_prepared(
                    prepared,
                    self.workbook.read(cx),
                )
            });

        // Store preview
        self.terminal.pending_result = Some(PendingResult::LuaPreview(LuaPreviewData {
            script_path,
            script_hash,
            cells_written,
            cells_overwritten,
            ops: result.ops,
            prepared_plan,
            source_sheet_index,
            source_fingerprint,
            output: result.output,
            error: preview_error,
        }));

        if self.review_mode.is_some() {
            self.focus_first_review_change(cx);
        }

        cx.notify();
    }
    /// Re-preview last.lua from disk — recovers a dismissed preview or re-runs
    /// the most recent AI-generated script.
    pub fn preview_last_lua(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use crate::terminal::state::{LuaPreviewData, PendingResult};
        use crate::scripting::SheetSnapshot;

        if self.block_review_entry_for_workbook_transition(cx) { return; }

        let last_path = self.terminal.workspace_root.as_ref()
            .map(|r| r.join("ai").join("generated").join("last.lua"));

        let Some(path) = last_path else {
            self.status_message = Some("No workspace root — open or save a file first.".into());
            cx.notify();
            return;
        };

        let code = match std::fs::read_to_string(&path) {
            Ok(c) if !c.trim().is_empty() => c,
            _ => {
                self.status_message = Some("No last.lua found. Run \"AI: Capture Last Lua Block\" first.".into());
                cx.notify();
                return;
            }
        };

        // Same preview flow as capture_ai_lua but from file
        let hash = blake3::hash(code.as_bytes());
        let script_hash = hash.to_hex()[..32].to_string();

        if self.mode.is_editing() {
            self.cancel_edit(cx);
        }
        let sheet = self.sheet(cx);
        let source_sheet_index = self.sheet_index(cx);
        let source_fingerprint = sheet_fingerprint(sheet);
        let snapshot = SheetSnapshot::from_sheet(sheet);

        let result = self.lua_runtime.eval_with_sheet(&code, Box::new(snapshot));

        let cells_overwritten = self.workbook.read(cx).sheet(source_sheet_index)
            .map(|s| count_lua_overwrites(&result.ops, s))
            .unwrap_or(0);

        if self.terminal.pending_result.is_some() {
            self.status_message = Some("Previous pending result replaced.".into());
        }

        let mut preview_error = result.error;
        let prepared_plan = if preview_error.is_none() {
            if result.ops.iter().any(|op| matches!(op, crate::scripting::LuaOp::DeleteRows { .. }))
                && (self.row_view.is_sorted() || self.filter_state.is_enabled())
            {
                preview_error = Some(
                    "unsupported_view_state: clear the active sort/filter before reviewing row deletion".into(),
                );
                None
            } else {
                match prepare_lua_operation_plan(
                    self.workbook.read(cx),
                    self.session_window_id,
                    &path,
                    &script_hash,
                    &result.ops,
                    default_lua_verification(),
                ) {
                    Ok(plan) => match require_visible_plan_changes(plan) {
                        Ok(plan) => Some(plan),
                        Err(error) => {
                            preview_error = Some(error);
                            None
                        }
                    },
                    Err(error) => {
                        preview_error = Some(error);
                        None
                    }
                }
            }
        } else {
            None
        };
        let cells_written = prepared_plan.as_ref()
            .map(|prepared| prepared.plan().summary.total_changes())
            .unwrap_or(result.mutations);
        self.review_mode = prepared_plan
            .as_ref()
            .map(|prepared| {
                crate::review_mode::ReviewModeState::from_prepared(
                    prepared,
                    self.workbook.read(cx),
                )
            });

        self.terminal.pending_result = Some(PendingResult::LuaPreview(LuaPreviewData {
            script_path: path,
            script_hash,
            cells_written,
            cells_overwritten,
            ops: result.ops,
            prepared_plan,
            source_sheet_index,
            source_fingerprint,
            output: result.output,
            error: preview_error,
        }));

        if self.review_mode.is_some() {
            self.focus_first_review_change(cx);
        }

        cx.notify();
    }
    /// Apply the pending Lua preview to a new sheet.
    pub fn apply_lua_to_new_sheet(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use crate::terminal::state::PendingResult;

        let preview = match self.terminal.pending_result.take() {
            Some(PendingResult::LuaPreview(data)) => data,
            other => {
                self.terminal.pending_result = other;
                return;
            }
        };

        // Refuse if preview had error
        if preview.error.is_some() {
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            self.status_message = Some("Cannot apply: Lua script had errors.".into());
            cx.notify();
            return;
        }

        let Some(prepared) = preview.prepared_plan.as_ref() else {
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            self.status_message = Some("Cannot copy: complete materialized preview is unavailable.".into());
            cx.notify();
            return;
        };
        let Some(preview_sheet) = prepared
            .preview_workbook()
            .sheet_by_id(prepared.plan().source_sheet_id)
            .cloned()
        else {
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            self.status_message = Some("Cannot copy: preview source sheet is unavailable.".into());
            cx.notify();
            return;
        };

        let hash_prefix = if preview.script_hash.len() >= 8 {
            &preview.script_hash[..8]
        } else {
            &preview.script_hash
        };
        let base_name = format!("AI Result - {}", hash_prefix);
        let before_workbook = self.wb(cx).clone();
        let before_row_view = self.row_view.clone();

        // Copy the complete materialized preview, not just changed operations.
        // This remains honest even when the original source has gone stale.
        let (sheet_name, sheet_idx) = self.workbook.update(cx, |wb, _| {
            let name = crate::structured_results::unique_sheet_name(wb, &base_name);
            let idx = wb
                .add_sheet_clone_named(&preview_sheet, &name)
                .expect("unique preview sheet name must be accepted");
            (name, idx)
        });

        self.review_mode = None;
        let activated = self.activate_sheet(sheet_idx, cx);
        debug_assert!(activated);
        self.row_view = visigrid_engine::filter::RowView::new(crate::app::NUM_ROWS);
        self.clear_selection_state();

        let after_workbook = self.wb(cx).clone();
        self.history.record_action_with_provenance(
            crate::history::UndoAction::WorkbookSnapshot {
                commit: Box::new(crate::history::WorkbookSnapshotCommit::new(
                    format!("Copy reviewed result to '{}'", sheet_name),
                    before_workbook,
                    after_workbook,
                )),
                before_row_view,
                after_row_view: self.row_view.clone(),
            },
            None,
        );
        self.is_modified = true;

        self.status_message = Some(format!(
            "Applied AI Lua to new sheet '{}'.", sheet_name
        ));
        cx.notify();
    }
    /// Apply the pending Lua preview to the current (source) sheet.
    pub fn apply_lua_to_current_sheet(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use crate::terminal::state::PendingResult;

        let preview = match self.terminal.pending_result.take() {
            Some(PendingResult::LuaPreview(data)) => data,
            other => {
                self.terminal.pending_result = other;
                return;
            }
        };

        // Refuse if preview had error
        if preview.error.is_some() {
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            self.status_message = Some("Cannot apply: Lua script had errors.".into());
            cx.notify();
            return;
        }

        let Some(prepared) = preview.prepared_plan.as_ref() else {
            self.status_message = Some("Cannot apply: preview did not produce a valid operation plan.".into());
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            cx.notify();
            return;
        };
        if prepared.plan().operations.iter().any(|operation| {
            matches!(operation.operation, visigrid_engine::operation_plan::PlannedOp::DeleteRows { .. })
        }) && (self.row_view.is_sorted() || self.filter_state.is_enabled()) {
            self.status_message = Some(
                "Cannot apply row deletion while the sheet is sorted or filtered.".into(),
            );
            self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
            cx.notify();
            return;
        }

        let context = crate::scripting::execution_context_fingerprint(
            self.workbook.read(cx),
            &prepared.plan().operations,
        );
        let commit = match prepared.verify_candidate(self.workbook.read(cx), &context) {
            Ok(commit) => commit,
            Err(error) => {
                self.status_message = Some(format!(
                    "Cannot apply: {error}. Re-preview the script first."
                ));
                self.terminal.pending_result = Some(PendingResult::LuaPreview(preview));
                cx.notify();
                return;
            }
        };
        let passed_verifications = commit
            .verification
            .iter()
            .filter(|result| {
                result.status
                    == visigrid_engine::operation_plan::VerificationStatus::Passed
            })
            .count();
        let verification_count = commit.verification.len();
        let applied_plan_id = prepared.plan().id.0.clone();
        let applied_plan_hash = prepared.plan().plan_hash.clone();
        let applied_changes = prepared.plan().changes.len();

        let sheet_id = prepared.plan().source_sheet_id;
        let sheet_idx = self.workbook.read(cx)
            .sheet_index_by_id(sheet_id)
            .expect("verified plan source sheet still exists");
        let before_row_view = self.row_view.clone();
        let before_row_heights = self.row_heights.get(&sheet_id).cloned().unwrap_or_default();
        let (after_row_view, after_row_heights) = plan_row_state_after_apply(
            &before_row_view,
            &before_row_heights,
            &prepared.plan().operations,
        );
        self.workbook.update(cx, |workbook, _| *workbook = commit.applied.clone());
        self.row_view = after_row_view.clone();
        self.row_heights.insert(sheet_id, after_row_heights.clone());

        let sheet_name = self.workbook.read(cx)
            .sheet_names()
            .get(sheet_idx)
            .map(|s| s.to_string())
            .unwrap_or_default();
        self.history.record_action_with_provenance(
            crate::history::UndoAction::PlanCommit {
                commit: Box::new(commit),
                sheet_id,
                before_row_view,
                after_row_view,
                before_row_heights,
                after_row_heights,
            },
            None,
        );
        self.bump_cells_rev();
        self.is_modified = true;
        self.review_mode = None;

        if self.mcp_plans.record(&applied_plan_id).is_some() {
            let result_fingerprint =
                visigrid_engine::operation_plan::workbook_fingerprint(self.workbook.read(cx));
            self.mcp_plans.mark_applied(
                &applied_plan_id,
                serde_json::json!({
                    "plan_id": applied_plan_id,
                    "state": "applied",
                    "applied_revision": self.workbook.read(cx).revision(),
                    "applied_changes": applied_changes,
                    "plan_hash": applied_plan_hash,
                    "result_fingerprint": result_fingerprint,
                    "undo_available": true,
                    "already_applied": false,
                    "warnings": [],
                }),
            );
        }

        self.status_message = Some(if verification_count == 0 {
            format!("Applied AI Lua to current sheet '{}'.", sheet_name)
        } else {
            format!(
                "Applied AI Lua to current sheet '{}'. Final verification: {passed_verifications}/{verification_count} passed.",
                sheet_name
            )
        });
        cx.notify();
    }
    /// One-time hint: "AI: Explain Selection auto-pastes everything."
    /// Appended to the first paste status message, then dismissed.
    pub(crate) fn maybe_append_ai_paste_tip(&self, msg: &mut String, cx: &mut Context<Self>) {
        if !user_settings(cx).is_tip_dismissed(TipId::AiPasteShortcut) {
            msg.push_str(" Tip: \"AI: Explain Selection\" auto-pastes everything.");
            update_user_settings(cx, |s| s.dismiss_tip(TipId::AiPasteShortcut));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::lua_preview_source_matches;
    use crate::app::sheet_fingerprint;

    #[test]
    fn lua_preview_source_must_still_exist_and_match() {
        let mut workbook = visigrid_engine::workbook::Workbook::new();
        let fingerprint = sheet_fingerprint(workbook.active_sheet());
        assert!(lua_preview_source_matches(&workbook, 0, fingerprint));

        workbook.set_cell_value_tracked(0, 0, 0, "changed");
        assert!(!lua_preview_source_matches(&workbook, 0, fingerprint));
        assert!(!lua_preview_source_matches(&workbook, 99, fingerprint));
    }
}
