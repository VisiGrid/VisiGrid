//! Publication and presentation of reviewed plans with saved Table criteria.
use crate::{app::Spreadsheet, history::UndoAction, table_structure::TableStructureHistory};
use gpui::Context;
use visigrid_engine::{
    operation_plan::{PlanCommit, PlannedOp, PlannedOperation, PreparedOperationPlan},
    structural::Axis,
    workbook::StructureStep,
};

pub(crate) fn review_steps(ops: &[PlannedOperation]) -> Vec<StructureStep> {
    ops.iter()
        .filter_map(|op| match op.operation {
            PlannedOp::DeleteRows { at, count } => Some(StructureStep {
                axis: Axis::Row,
                at,
                count,
                delete: true,
            }),
            _ => None,
        })
        .collect()
}

fn frozen_rows_after(mut boundary: usize, steps: &[StructureStep]) -> usize {
    for step in steps {
        if step.at < boundary {
            boundary -= step.count.min(boundary - step.at);
        }
    }
    boundary
}

impl Spreadsheet {
    pub(crate) fn validate_table_review(
        &self,
        prepared: &PreparedOperationPlan,
    ) -> Result<(), String> {
        if !prepared.source_workbook().has_table_criteria() {
            return Ok(());
        }
        let id = prepared.plan().source_sheet_id;
        let sheet = prepared
            .source_workbook()
            .sheet_by_id(id)
            .ok_or("Review source sheet no longer exists.")?;
        let steps = review_steps(&prepared.plan().operations);
        let layout = self.structure_layout(id).shifted(sheet, &steps)?;
        let mut candidate = prepared.preview_workbook().clone();
        if !steps.is_empty() && id == self.cached_sheet_id() {
            candidate.sheet_by_id_mut(id).unwrap().frozen_panes = (
                frozen_rows_after(self.view_state.frozen_rows, &steps),
                self.view_state.frozen_cols,
            );
        }
        self.validate_structure_layout(&candidate, id, &layout)
    }

    pub(crate) fn publish_table_review(
        &mut self,
        plan: &PlanCommit,
        sheet_id: visigrid_engine::sheet::SheetId,
        operations: &[PlannedOperation],
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        let steps = review_steps(operations);
        let before = self.structure_layout(sheet_id);
        let source = self
            .wb(cx)
            .sheet_by_id(sheet_id)
            .ok_or("Review source sheet no longer exists.")?;
        let after = before.shifted(source, &steps)?;
        // Frozen panes are owned by the desktop view until save. Normalize
        // that existing state into both history endpoints without publishing
        // an intermediate workbook or changing the reviewed cell result.
        let source_frozen = (!steps.is_empty() && sheet_id == self.cached_sheet_id())
            .then_some((self.view_state.frozen_rows, self.view_state.frozen_cols));
        let mut source = self.wb(cx).clone();
        let mut candidate = plan.applied.clone();
        if !steps.is_empty() && sheet_id == self.cached_sheet_id() {
            source.sheet_by_id_mut(sheet_id).unwrap().frozen_panes =
                (self.view_state.frozen_rows, self.view_state.frozen_cols);
            candidate.sheet_by_id_mut(sheet_id).unwrap().frozen_panes = (
                frozen_rows_after(self.view_state.frozen_rows, &steps),
                self.view_state.frozen_cols,
            );
        }
        self.validate_structure_layout(&candidate, sheet_id, &after)?;
        let mut commit = source.capture_guarded_batch(&candidate)?;
        commit.sheet = sheet_id;
        commit.steps = steps;
        let sheet_index = self.wb(cx).sheet_index_by_id(sheet_id).unwrap();
        let description = format!("Apply reviewed Lua plan {}", plan.plan_id.0);
        let structural = !commit.steps.is_empty();
        let action = if !structural {
            UndoAction::TableBatchChanged {
                sheet_index,
                commit: Box::new(commit),
                description,
            }
        } else {
            UndoAction::TableStructureChanged {
                sheet_index,
                history: Box::new(TableStructureHistory {
                    source_frozen,
                    commit,
                    before,
                    after: after.clone(),
                }),
                description,
            }
        };
        // All validation precedes publication. Review is a frozen source-row view.
        self.review_mode = None;
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.install_structure_layout(sheet_id, &after);
        if structural {
            self.view_state.frozen_rows = self.sheet(cx).frozen_panes.0;
            self.view_state.frozen_cols = self.sheet(cx).frozen_panes.1;
        }
        self.table_view_sync_key = None;
        self.sync_table_view(cx);
        self.view_state.selection_end = None;
        self.view_state.additional_selections.clear();
        self.clipboard_visual_range = None;
        self.history.record_action_with_provenance(action, None);
        self.cached_title = None;
        self.bump_cells_rev();
        self.bump_cf_rules_rev();
        self.is_modified = true;
        cx.notify();
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        history::History,
        review_mode::{ReviewEndpoint, ReviewModeState},
        table_edit::tests::fixture,
    };
    use visigrid_engine::operation_plan::{PlanId, PlanProducer};

    fn prepare(wb: &visigrid_engine::workbook::Workbook, code: &str) -> PreparedOperationPlan {
        let runtime = crate::scripting::LuaRuntime::new().unwrap();
        let result = runtime.eval_with_sheet(
            code,
            Box::new(crate::scripting::SheetSnapshot::from_sheet(
                wb.active_sheet(),
            )),
        );
        assert!(result.error.is_none(), "{:?}", result.error);
        crate::ai_actions::prepare_lua_operation_plan_with_metadata(
            wb,
            42,
            PlanId("filtered-review".into()),
            PlanProducer {
                kind: "test".into(),
                name: "test".into(),
                source_path: None,
                source_hash: None,
            },
            "Filtered review".into(),
            None,
            "test",
            &result.ops,
            vec![],
        )
        .unwrap()
    }

    #[test]
    fn reviewed_deletions_shift_or_shrink_the_frozen_boundary() {
        let step = |at, count| StructureStep {
            axis: Axis::Row,
            at,
            count,
            delete: true,
        };
        assert_eq!(frozen_rows_after(3, &[step(0, 1)]), 2);
        assert_eq!(frozen_rows_after(3, &[step(2, 4)]), 2);
        assert_eq!(frozen_rows_after(3, &[step(3, 1)]), 3);
        assert_eq!(frozen_rows_after(5, &[step(3, 1), step(0, 1)]), 3);
    }

    #[test]
    fn hidden_row_deletion_review_endpoints_and_sparse_rewind_preserve_layout() {
        let wb = fixture(true);
        let prepared = prepare(&wb, "sheet:set('C6',25); sheet:delete_rows(5,1)");
        let mut review = ReviewModeState::from_prepared(&prepared, &wb);
        assert!(review.is_deleted_source_row(4));
        assert!(review.navigable_cells().iter().any(|(r, _)| *r == 4));
        review.set_endpoint(ReviewEndpoint::After);
        let (sheet, row) = review
            .endpoint_sheet_row(&prepared, wb.active_sheet_id(), 4)
            .unwrap();
        assert_eq!(row, 4);
        assert_eq!(sheet.get_raw(row, 1), "East"); // Tombstone remains the source record.
        let (sheet, row) = review
            .endpoint_sheet_row(&prepared, wb.active_sheet_id(), 5)
            .unwrap();
        assert_eq!(row, 4);
        assert_eq!(sheet.get_raw(row, 2), "25");
        let context =
            crate::scripting::execution_context_fingerprint(&wb, &prepared.plan().operations);
        let applied = prepared.verify_candidate(&wb, &context).unwrap().applied;
        let mut commit = wb.capture_guarded_batch(&applied).unwrap();
        commit.steps = review_steps(&prepared.plan().operations);
        let mut before = crate::table_structure::StructureLayout::default();
        before.heights.insert(9, 40.0);
        before.hidden_rows.insert(20);
        let after = before.shifted(wb.active_sheet(), &commit.steps).unwrap();
        assert_eq!(after.heights.get(&8), Some(&40.0));
        assert!(after.hidden_rows.contains(&19));
        let mut live = applied.clone();
        commit.replay(&mut live, true).unwrap();
        assert_eq!(live.active_sheet().get_raw(4, 1), "East");
        commit.replay(&mut live, false).unwrap();
        assert_eq!(live.active_sheet().get_raw(4, 2), "25");
        let mut history = History::new();
        history.record_named_range_action(UndoAction::TableStructureChanged {
            sheet_index: 0,
            history: Box::new(TableStructureHistory {
                source_frozen: None,
                commit,
                before,
                after: after.clone(),
            }),
            description: "Reviewed row deletion".into(),
        });
        let rewind = history
            .build_workbook_before(1, Some(&wb), 100, 10_000)
            .unwrap();
        assert_eq!(rewind.workbook.active_sheet().get_raw(4, 2), "25");
        assert_eq!(
            rewind.view_state.per_sheet[0].structure_layout.as_ref(),
            Some(&after)
        );
        assert!(rewind.view_state.per_sheet[0].table_rows.is_some());
    }

    #[test]
    fn reviewed_frozen_panes_rewind_from_the_desktop_source_state() {
        let wb = fixture(true);
        let prepared = prepare(&wb, "sheet:delete_rows(1,1)");
        let context =
            crate::scripting::execution_context_fingerprint(&wb, &prepared.plan().operations);
        let mut applied = prepared.verify_candidate(&wb, &context).unwrap().applied;
        let mut source = wb.clone();
        source.active_sheet_mut().frozen_panes = (3, 1);
        applied.active_sheet_mut().frozen_panes = (2, 1);
        let mut commit = source.capture_guarded_batch(&applied).unwrap();
        commit.steps = review_steps(&prepared.plan().operations);
        let mut live = applied.clone();
        commit.replay(&mut live, true).unwrap();
        assert_eq!(live.active_sheet().frozen_panes, (3, 1));
        commit.replay(&mut live, false).unwrap();
        assert_eq!(live.active_sheet().frozen_panes, (2, 1));
        let mut history = History::new();
        history.record_named_range_action(UndoAction::TableStructureChanged {
            sheet_index: 0,
            history: Box::new(TableStructureHistory {
                source_frozen: Some((3, 1)),
                commit,
                before: Default::default(),
                after: Default::default(),
            }),
            description: "Reviewed rows".into(),
        });
        let before = history
            .build_workbook_before(0, Some(&wb), 100, 10_000)
            .unwrap();
        assert_eq!(before.workbook.active_sheet().frozen_panes, (3, 1));
        let after = history
            .build_workbook_before(1, Some(&wb), 100, 10_000)
            .unwrap();
        assert_eq!(after.workbook.active_sheet().frozen_panes, (2, 1));
    }

    #[test]
    fn hidden_cell_plan_keeps_saved_criteria_and_replays_as_a_batch() {
        let wb = fixture(true);
        let prepared = prepare(&wb, "sheet:set('C5',55)");
        assert!(review_steps(&prepared.plan().operations).is_empty());
        assert_eq!(
            prepared.preview_workbook().active_sheet().table_view_spec(),
            wb.active_sheet().table_view_spec()
        );
        let context =
            crate::scripting::execution_context_fingerprint(&wb, &prepared.plan().operations);
        let applied = prepared.verify_candidate(&wb, &context).unwrap().applied;
        let commit = wb.capture_guarded_batch(&applied).unwrap();
        let mut history = History::new();
        history.record_named_range_action(UndoAction::TableBatchChanged {
            sheet_index: 0,
            commit: Box::new(commit),
            description: "Reviewed cells".into(),
        });
        let rewind = history
            .build_workbook_before(1, Some(&wb), 100, 10_000)
            .unwrap();
        assert_eq!(rewind.workbook.active_sheet().get_raw(4, 2), "55");
        assert!(rewind.view_state.per_sheet[0].table_rows.is_some());
    }
}
