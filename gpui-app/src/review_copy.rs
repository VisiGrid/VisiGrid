//! One-sheet history for copying a complete reviewed result.
use crate::{app::Spreadsheet, history::UndoAction, table_structure::StructureLayout};
use gpui::Context;
use visigrid_engine::{
    operation_plan::PreparedOperationPlan,
    sheet::{Sheet, SheetId},
    workbook::Workbook,
};

#[derive(Clone, Debug)]
pub(crate) struct ReviewCopyHistory {
    pub sheet: Sheet,
    pub index: usize,
    pub layout: StructureLayout,
    before_active: SheetId,
    before: String,
    after: String,
}
// Fingerprint authored state, not volatile computed caches or spill receivers.
// Stale replay must refuse changed comments, metadata and formulas too.
fn normalize_signature(value: &mut serde_json::Value) {
    match value {
        serde_json::Value::Object(map) => {
            // Allocators stay monotonic across undo; membership is in columns.
            map.remove("next_column_id");
            for (key, value) in map {
                if key == "selected" {
                    if let Some(values) = value.as_array_mut() {
                        values.sort_by_key(|v| v.to_string());
                    }
                }
                normalize_signature(value);
            }
        }
        serde_json::Value::Array(values) => {
            for value in values {
                normalize_signature(value);
            }
        }
        _ => {}
    }
}

fn sheet_signature(sheet: &Sheet) -> Result<serde_json::Value, String> {
    use visigrid_engine::cell::ValueRef;
    let mut cells: Vec<_> = sheet
        .cells_iter()
        .filter(|(_, c)| {
            c.spill_parent().is_none()
                && !(matches!(c.value(), ValueRef::Empty)
                    && c.format().is_default()
                    && c.comment().is_none()
                    && c.style_id().is_none()
                    && c.frozen_formula().is_none())
        })
        .map(|((r, c), cell)| {
            let value = match cell.value() {
                ValueRef::Empty => serde_json::json!(["empty"]),
                ValueRef::Text(s) => serde_json::json!(["text", s]),
                ValueRef::Number(n) => serde_json::json!(["number", n.to_bits()]),
                ValueRef::Formula { source, .. } => serde_json::json!(["formula", source]),
            };
            (
                r,
                c,
                value,
                cell.format().clone(),
                cell.comment().cloned(),
                cell.style_id(),
                cell.frozen_formula().map(str::to_string),
            )
        })
        .collect();
    cells.sort_by_key(|v| (v.0, v.1));
    let mut value = serde_json::json!({"id":sheet.id,"name":sheet.name,"size":[sheet.rows,sheet.cols],"cells":cells,
        "tables":sheet.tables(),"view":sheet.table_view_spec(),"pivots":sheet.pivots,"merges":sheet.merged_regions,
        "validation":sheet.validations,"conditional":sheet.cond_formats,"frozen":sheet.frozen_panes,
        "print":sheet.print_setup,"tab":sheet.tab_color,"rows":sheet.row_formats,"cols":sheet.col_formats});
    normalize_signature(&mut value);
    Ok(value)
}
fn signature(wb: &Workbook) -> String {
    let mut names = wb.named_ranges().list();
    names.sort_by(|a, b| a.name.cmp(&b.name));
    let value = serde_json::json!({"sheets":wb.sheets().iter().map(sheet_signature).collect::<Result<Vec<_>,_>>().expect("sheet state serializes"),"names":names,"styles":wb.style_table});
    blake3::hash(value.to_string().as_bytes())
        .to_hex()
        .to_string()
}

fn validate_views(wb: &Workbook) -> Result<(), String> {
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(())
}

pub(crate) fn prepare_copy(
    wb: &Workbook,
    prepared: &PreparedOperationPlan,
    layout: &StructureLayout,
    frozen: (usize, usize),
    base_name: &str,
) -> Result<(Workbook, ReviewCopyHistory), String> {
    wb.ensure_writable()?;
    validate_views(wb)?;
    let id = prepared.plan().source_sheet_id;
    let source = prepared
        .source_workbook()
        .sheet_by_id(id)
        .ok_or("The reviewed source sheet is unavailable.")?;
    let steps = crate::table_review::review_steps(&prepared.plan().operations);
    let layout = layout.shifted(source, &steps)?;
    let name = crate::structured_results::unique_sheet_name(wb, base_name);
    let (mut candidate, index) = wb.prepare_sheet_copy(prepared.preview_workbook(), id, &name)?;
    let mut frozen_rows = frozen.0;
    for step in steps {
        if step.at < frozen_rows {
            frozen_rows -= step.count.min(frozen_rows - step.at);
        }
    }
    candidate.sheet_mut(index).unwrap().frozen_panes = (frozen_rows, frozen.1);
    validate_views(&candidate)?;
    let history = ReviewCopyHistory {
        sheet: candidate.sheet(index).unwrap().clone(),
        index,
        layout,
        before_active: wb.active_sheet_id(),
        before: signature(wb),
        after: signature(&candidate),
    };
    Ok((candidate, history))
}

impl ReviewCopyHistory {
    pub(crate) fn replay(&self, wb: &Workbook, undo: bool) -> Result<Workbook, String> {
        wb.ensure_writable()?;
        let expected = if undo { &self.after } else { &self.before };
        if &signature(wb) != expected {
            return Err(
                "The workbook changed since this result was copied. Undo those changes first."
                    .into(),
            );
        }
        let mut candidate = wb.clone();
        if undo {
            let sheet = wb
                .sheet(self.index)
                .ok_or("The copied sheet is no longer in its original position.")?;
            if sheet_signature(sheet)? != sheet_signature(&self.sheet)? {
                return Err(
                    "The copied sheet changed. Undo its edits before removing the copy.".into(),
                );
            }
            candidate.take_sheet(self.index).ok_or(
                "The copied sheet cannot be removed while referenced by another Table or pivot.",
            )?;
        } else if !candidate.restore_sheet(self.index, self.sheet.clone()) {
            return Err(
                "The copied sheet or one of its Table identities is already in use.".into(),
            );
        }
        candidate.rebuild_dep_graph();
        candidate.recompute_full_ordered();
        validate_views(&candidate)?;
        if signature(&candidate) != *if undo { &self.before } else { &self.after } {
            return Err("Copy history would no longer reproduce the expected workbook.".into());
        }
        Ok(candidate)
    }
}

impl Spreadsheet {
    pub(crate) fn publish_review_copy(
        &mut self,
        candidate: Workbook,
        history: ReviewCopyHistory,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.validate_structure_layout(&candidate, history.sheet.id, &history.layout)?;
        self.review_mode = None;
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.install_structure_layout(history.sheet.id, &history.layout);
        self.finish_review_copy_view(history.sheet.id, cx);
        self.history.record_action_with_provenance(
            UndoAction::ReviewCopy {
                history: Box::new(history),
            },
            None,
        );
        self.is_modified = true;
        self.bump_cells_rev();
        self.bump_cf_rules_rev();
        self.cached_title = None;
        Ok(())
    }
    fn finish_review_copy_view(&mut self, id: SheetId, cx: &mut Context<Self>) {
        let index = self.wb(cx).sheet_index_by_id(id).unwrap_or(0);
        self.activate_sheet(index, cx);
        self.row_view =
            visigrid_engine::filter::RowView::new(crate::app::NUM_ROWS.min(self.sheet(cx).rows));
        self.view_state.frozen_rows = self.sheet(cx).frozen_panes.0;
        self.view_state.frozen_cols = self.sheet(cx).frozen_panes.1;
        self.clear_selection_state();
        self.table_filter_dropdown = None;
        self.table_view_sync_key = None;
        self.sync_table_view(cx);
        self.clipboard_visual_range = None;
    }
    pub(crate) fn replay_review_copy(
        &mut self,
        history: &ReviewCopyHistory,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = (|| {
            if undo && self.structure_layout(history.sheet.id) != history.layout {
                return Err("The copied sheet layout changed. Undo those changes first.".into());
            }
            let candidate = history.replay(self.wb(cx), undo)?;
            self.validate_structure_layout(&candidate, history.sheet.id, &history.layout)?;
            Ok::<_, String>(candidate)
        })();
        match result {
            Ok(candidate) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                if undo {
                    self.row_heights.remove(&history.sheet.id);
                    self.col_widths.remove(&history.sheet.id);
                    self.hidden_rows.remove(&history.sheet.id);
                    self.hidden_cols.remove(&history.sheet.id);
                } else {
                    self.install_structure_layout(history.sheet.id, &history.layout);
                }
                self.finish_review_copy_view(
                    if undo {
                        history.before_active
                    } else {
                        history.sheet.id
                    },
                    cx,
                );
                self.is_modified = true;
                self.bump_cells_rev();
                self.bump_cf_rules_rev();
                cx.notify();
                true
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                false
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{history::History, table_edit::tests::fixture};
    use visigrid_engine::{
        operation_plan::{PlanId, PlanProducer},
        table::TableRange,
    };
    fn plan(wb: &Workbook, code: &str) -> PreparedOperationPlan {
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
            PlanId("copy-test".into()),
            PlanProducer {
                kind: "test".into(),
                name: "test".into(),
                source_path: None,
                source_hash: None,
            },
            "Copy review".into(),
            None,
            "test",
            &result.ops,
            vec![],
        )
        .unwrap()
    }
    fn copy(wb: &Workbook, p: &PreparedOperationPlan) -> (Workbook, ReviewCopyHistory) {
        prepare_copy(
            wb,
            p,
            &StructureLayout::default(),
            (0, 0),
            "AI Result - test",
        )
        .unwrap()
    }
    #[test]
    fn copy_preserves_source_and_includes_hidden_records_with_fresh_table_identity() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (after, h) = copy(&before, &prepared);
        assert_eq!(signature(&before), signature(prepared.source_workbook()));
        assert_eq!(after.active_sheet().get_raw(5, 2), "20");
        let result = after.sheet(h.index).unwrap();
        assert_eq!(result.get_raw(5, 2), "25");
        assert_eq!(result.get_raw(4, 1), "East");
        assert_eq!(result.tables()[0].name, "Sales_Copy");
        assert_ne!(result.tables()[0].id, before.active_sheet().tables()[0].id);
        assert_eq!(
            result.table_view_spec().unwrap().table,
            result.tables()[0].id
        );
        assert_eq!(
            result.table_view_spec().unwrap().filters,
            before.active_sheet().table_view_spec().unwrap().filters
        );
        assert!(result
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .rows()
            .data_to_view(4)
            .is_none());
        let undone = h.replay(&after, true).unwrap();
        assert_eq!(signature(&undone), signature(&before));
        let redone = h.replay(&undone, false).unwrap();
        assert_eq!(signature(&redone), signature(&after));
    }
    #[test]
    fn local_table_formulas_rules_and_overrides_follow_copy() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=Sales[@Amount]*2", true)
            .unwrap();
        before.set_cell_value_tracked(0, 4, 3, "999");
        before.set_cell_value_tracked(0, 0, 4, "=SUM(Sales[Amount])");
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (after, h) = copy(&before, &prepared);
        let sheet = after.sheet(h.index).unwrap();
        assert!(sheet.get_raw(3, 3).contains("Sales_Copy"));
        assert!(sheet.tables()[0].columns[2]
            .formula
            .as_ref()
            .unwrap()
            .contains("Sales_Copy"));
        assert_eq!(sheet.get_raw(4, 3), "999");
        assert!(sheet.is_calculated_exception(4, 3));
        assert_eq!(sheet.get_display(0, 4), "105");
        assert_eq!(after.active_sheet().get_display(0, 4), "100");
    }
    #[test]
    fn repeated_copies_choose_unique_sheet_and_table_names() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (first, _) = copy(&before, &prepared);
        let (second, h) = copy(&first, &prepared);
        assert_eq!(second.sheet_count(), 3);
        assert_ne!(second.sheet(1).unwrap().name, h.sheet.name);
        assert_eq!(h.sheet.tables()[0].name, "Sales_Copy2");
        assert_ne!(
            second.sheet(1).unwrap().tables()[0].id,
            h.sheet.tables()[0].id
        );
    }
    #[test]
    fn stale_source_copies_frozen_review_instead_of_current_cells() {
        let mut live = fixture(true);
        let prepared = plan(&live, "sheet:set('C6',25)");
        live.set_cell_value_tracked(0, 3, 2, "777");
        let (after, h) = copy(&live, &prepared);
        assert_eq!(after.active_sheet().get_raw(3, 2), "777");
        assert_eq!(h.sheet.get_raw(3, 2), "30");
        assert_eq!(h.sheet.get_raw(5, 2), "25");
        assert_eq!(
            signature(&h.replay(&after, true).unwrap()),
            signature(&live)
        );
    }
    #[test]
    fn reviewed_row_deletion_moves_copy_layout_and_keeps_original_records() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:delete_rows(5,1)");
        let mut layout = StructureLayout::default();
        layout.heights.insert(12, 35.0);
        layout.hidden_rows.insert(15);
        layout.widths.insert(2, 145.0);
        let (after, h) = prepare_copy(&before, &prepared, &layout, (3, 1), "Result").unwrap();
        assert_eq!(h.sheet.tables()[0].range.end_row, 5);
        assert_eq!(after.active_sheet().get_raw(4, 1), "East");
        assert_eq!(h.layout.heights.get(&11), Some(&35.0));
        assert!(h.layout.hidden_rows.contains(&14));
        assert_eq!(h.layout.widths, layout.widths);
        assert_eq!(h.sheet.frozen_panes, (3, 1));
        assert!(h.replay(&after, true).is_ok());
    }
    #[test]
    fn new_sheet_name_binding_that_invalidates_existing_view_refuses_atomically() {
        let mut live = fixture(true);
        live.set_cell_value_tracked(0, 0, 0, "=IF(IFERROR('Result'!C6,0)>0,SEQUENCE(5),0)");
        let prepared = plan(&live, "sheet:set('C6',25)");
        let sig = signature(&live);
        assert!(prepare_copy(
            &live,
            &prepared,
            &StructureLayout::default(),
            (0, 0),
            "Result"
        )
        .is_err());
        assert_eq!(signature(&live), sig);
        assert_eq!(live.sheet_count(), 1);
    }
    #[test]
    fn undo_refuses_changed_copy_comments_or_external_dependents_without_mutation() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (mut after, h) = copy(&before, &prepared);
        after.sheet_mut(h.index).unwrap().set_comment(
            0,
            0,
            Some(visigrid_engine::cell::CellComment {
                text: "keep".into(),
                author: "me".into(),
            }),
        );
        let sig = signature(&after);
        assert!(h.replay(&after, true).is_err());
        assert_eq!(signature(&after), sig);
        let (mut after, h) = copy(&before, &prepared);
        after.set_cell_value_tracked(0, 0, 5, "=SUM(Sales_Copy[Amount])");
        assert!(h.replay(&after, true).is_err());
    }
    #[test]
    fn rewind_and_native_full_save_keep_both_views_and_copied_layout() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (after, h) = copy(&before, &prepared);
        let index = h.index;
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::ReviewCopy {
                history: Box::new(h),
            },
            None,
        );
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(signature(&preview.workbook), signature(&after));
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
        assert!(preview.view_state.per_sheet[index].table_rows.is_some());
        let original = history
            .build_workbook_before(0, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(original.workbook.sheet_count(), 1);
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("review-copy.sheet");
        visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
            .unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(
            loaded.sheets()[index].tables(),
            after.sheets()[index].tables()
        );
        assert_eq!(
            loaded.sheets()[index].table_view_spec(),
            after.sheets()[index].table_view_spec()
        );
    }
    #[test]
    fn plain_reviewed_sheet_copies_while_another_sheet_has_criteria() {
        let mut before = fixture(true);
        let other = before.add_sheet_named("Plain").unwrap();
        before.set_active_sheet(other);
        before.set_cell_value_tracked(other, 0, 0, "source");
        let prepared = plan(&before, "sheet:set('B1',42)");
        let (after, h) = copy(&before, &prepared);
        assert!(h.sheet.tables().is_empty());
        assert_eq!(h.sheet.get_raw(0, 0), "source");
        assert_eq!(h.sheet.get_raw(0, 1), "42");
        assert_eq!(
            after.sheet(0).unwrap().table_view_spec(),
            before.sheet(0).unwrap().table_view_spec()
        );
        assert!(h.replay(&after, true).is_ok());
    }
    #[test]
    fn volatile_formulas_do_not_make_copy_history_stale() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 0, 5, "=RAND()");
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (after, h) = copy(&before, &prepared);
        let undone = h.replay(&after, true).unwrap();
        assert!(h.replay(&undone, false).is_ok());
    }
    #[test]
    fn recovery_and_style_changes_refuse_without_adding_sheet() {
        let mut before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        before.active_sheet_mut().read_only_reason = Some("Damaged tables".into());
        assert!(prepare_copy(
            &before,
            &prepared,
            &StructureLayout::default(),
            (0, 0),
            "Result"
        )
        .is_err());
        before.active_sheet_mut().read_only_reason = None;
        before.style_table.push(Default::default());
        assert!(prepare_copy(
            &before,
            &prepared,
            &StructureLayout::default(),
            (0, 0),
            "Result"
        )
        .unwrap_err()
        .contains("styles changed"));
        assert_eq!(before.sheet_count(), 1);
    }
    #[test]
    fn copied_table_growth_uses_remapped_calculated_rules() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=Sales[@Amount]*2", true)
            .unwrap();
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (mut after, h) = copy(&before, &prepared);
        let id = h.sheet.tables()[0].id;
        after
            .append_table_rows(id, 1, &[(7, 1, "West".into()), (7, 2, "7".into())])
            .unwrap();
        assert_eq!(after.sheet(h.index).unwrap().get_display(7, 3), "14");
        assert_eq!(
            after.active_sheet().tables()[0].range,
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 6,
                end_col: 3
            }
        );
    }
    #[test]
    fn undo_after_undone_resize_ignores_monotonic_column_allocator() {
        let before = fixture(true);
        let prepared = plan(&before, "sheet:set('C6',25)");
        let (mut after, h) = copy(&before, &prepared);
        let table = after.sheet(h.index).unwrap().tables()[0].clone();
        let change = after
            .resize_table(
                table.id,
                TableRange {
                    end_col: 4,
                    ..table.range
                },
            )
            .unwrap();
        after.apply_table_commit(&change, true).unwrap();
        assert!(h.replay(&after, true).is_ok());
    }
    #[test]
    fn rewind_copy_after_guarded_row_insertion_preserves_both_sheet_layouts() {
        use crate::table_structure::TableStructureHistory;
        use visigrid_engine::{structural::Axis, workbook::StructureStep};
        let base = fixture(true);
        let mut layout = StructureLayout::default();
        layout.heights.insert(10, 35.0);
        let steps = vec![StructureStep {
            axis: Axis::Row,
            at: 0,
            count: 1,
            delete: false,
        }];
        let shifted = layout.shifted(base.active_sheet(), &steps).unwrap();
        let (before, commit) = base.prepare_guarded_structure(0, steps).unwrap();
        let prepared = plan(&before, "sheet:set('C7',25)");
        let (after, copy) = prepare_copy(&before, &prepared, &shifted, (0, 0), "Result").unwrap();
        let index = copy.index;
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableStructureChanged {
                sheet_index: 0,
                history: Box::new(TableStructureHistory {
                    commit,
                    source_frozen: None,
                    before: layout,
                    after: shifted.clone(),
                }),
                description: "Insert row".into(),
            },
            None,
        );
        history.record_action_with_provenance(
            UndoAction::ReviewCopy {
                history: Box::new(copy),
            },
            None,
        );
        let preview = history
            .build_workbook_before(2, Some(&base), 100, 10_000)
            .unwrap();
        assert_eq!(signature(&preview.workbook), signature(&after));
        assert_eq!(
            preview.view_state.per_sheet[0].structure_layout,
            Some(shifted.clone())
        );
        assert_eq!(
            preview.view_state.per_sheet[index].structure_layout,
            Some(shifted)
        );
    }
}
