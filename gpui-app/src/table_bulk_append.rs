//! Resolve existing visible records before adding canonical records at the bottom.
use crate::{
    app::Spreadsheet,
    history::UndoAction,
    table_append::TableAppendHistory,
    table_cell_history::TableCellsCommit,
    table_edit::{prepare_table_writes_with_new_rows, validate_view_safe_targets, TableCellWrite},
};
use gpui::Context;
use visigrid_engine::{
    filter::RowView,
    sheet::Sheet,
    table::{TableId, TableRange},
    workbook::Workbook,
};

pub(crate) struct BulkAppendPlan {
    pub id: TableId,
    pub count: usize,
    pub targets: Vec<(usize, usize, usize, usize)>,
}

/// None retains normal paste/broadcast behavior. Formats never request growth.
pub(crate) fn plan_bulk_append(
    sheet: &Sheet,
    rows: &RowView,
    start: (usize, usize),
    height: usize,
    width: usize,
) -> Result<Option<BulkAppendPlan>, String> {
    if height == 0 || width == 0 || height.saturating_mul(width) > 100_000 {
        return Err("Paste between 1 and 100,000 cells at a time.".into());
    }
    let end_col = start
        .1
        .checked_add(width)
        .filter(|c| *c <= sheet.cols)
        .ok_or("The paste extends beyond the worksheet's columns.")?;
    let data_row = rows.view_to_data(start.0);
    let Some(table) = sheet.tables().iter().find(|t| {
        data_row > t.range.start_row
            && data_row <= t.range.end_row.saturating_add(1)
            && start.1 <= t.range.end_col
            && end_col > t.range.start_col
    }) else {
        return Ok(None);
    };
    if table.totals_row() == Some(data_row) {
        return Err("The totals row is protected. Use Add row or paste from an existing body record to add records above it.".into());
    }
    if !rows.is_view_row_visible(start.0) {
        return Err("Select a visible cell before pasting.".into());
    }
    let mut records: Vec<_> = if data_row <= table.range.end_row {
        rows.visible_rows()
            .iter()
            .copied()
            .skip(rows.visible_rows().partition_point(|r| *r < start.0))
            .take_while(|r| *r <= table.range.end_row)
            .take(height)
            .map(|r| rows.view_to_data(r))
            .collect()
    } else {
        Vec::new()
    };
    let count = height - records.len();
    if count == 0 {
        return Ok(None);
    }
    if start.1 < table.range.start_col || end_col > table.range.end_col + 1 {
        return Err("The paste crosses the Table's side boundary. Resize the Table first.".into());
    }
    let end = table
        .range
        .end_row
        .checked_add(count)
        .filter(|r| *r < sheet.rows)
        .ok_or("The append extends beyond the worksheet's rows.")?;
    if count.saturating_mul(table.range.end_col - table.range.start_col + 1) > 100_000 {
        return Err(
            "Append at most 100,000 Table cells at a time, including calculated columns.".into(),
        );
    }
    records.extend(table.range.end_row + 1..=end);
    let targets = records
        .into_iter()
        .enumerate()
        .flat_map(|(ri, r)| (0..width).map(move |ci| (r, start.1 + ci, ri, ci)))
        .collect();
    Ok(Some(BulkAppendPlan {
        id: table.id,
        count,
        targets,
    }))
}

pub(crate) fn prepare_append_writes(
    wb: &Workbook,
    id: TableId,
    count: usize,
    writes: &[TableCellWrite],
) -> Result<(Workbook, TableAppendHistory), String> {
    wb.ensure_writable()?;
    let (sheet_id, table) = wb.table(id).ok_or("The Table no longer exists.")?;
    let index = wb.sheet_index_by_id(sheet_id).unwrap();
    let end_row = table
        .range
        .end_row
        .checked_add(count)
        .ok_or("Append exceeds the worksheet.")?;
    let region = TableRange {
        start_row: table.range.end_row + 1,
        end_row,
        ..table.range
    };
    region.validate(wb.sheet(index).unwrap().rows, wb.sheet(index).unwrap().cols)?;
    if count == 0
        || writes.is_empty()
        || writes.len() > 100_000
        || count.saturating_mul(region.end_col - region.start_col + 1) > 100_000
    {
        return Err("Append between 1 and 100,000 Table cells at a time.".into());
    }
    let expanded = TableRange {
        end_row,
        ..table.range
    };
    let mut seen = std::collections::HashSet::new();
    if writes.iter().any(|w| {
        w.row <= expanded.start_row
            || !expanded.contains(w.row, w.col)
            || !seen.insert((w.row, w.col))
    }) {
        return Err("Paste targets must be unique cells within the expanded Table body.".into());
    }
    let existing: Vec<_> = writes
        .iter()
        .filter(|w| w.row < region.start_row)
        .map(|w| (w.row, w.col))
        .collect();
    validate_view_safe_targets(wb, index, &existing, false)?;
    let view = wb.sheet(index).unwrap().table_view_spec().cloned();
    let mut expanded_book = wb.clone();
    let commit = expanded_book.append_table_rows(id, count, &[])?;
    // This intermediate state is also replayed when undoing supplied overrides.
    for sheet in expanded_book.sheets() {
        sheet.build_saved_table_view(sheet.rows)?;
    }
    let candidate =
        prepare_table_writes_with_new_rows(&expanded_book, index, writes, Some(region))?;
    let patch = TableCellsCommit::capture(
        expanded_book.sheet(index).unwrap(),
        candidate.sheet(index).unwrap(),
        writes.iter().map(|w| (w.row, w.col)),
    );
    Ok((
        candidate,
        TableAppendHistory::with_appended_writes(commit, patch, view),
    ))
}

impl Spreadsheet {
    pub(crate) fn paste_and_append_table(
        &mut self,
        plan: BulkAppendPlan,
        writes: Vec<TableCellWrite>,
        cx: &mut Context<Self>,
    ) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            self.status_message = Some(
                "Clear this sheet's worksheet sorting and filters before appending Table rows."
                    .into(),
            );
            cx.notify();
            return;
        }
        let result = self
            .validate_saved_view_layout(self.wb(cx))
            .and_then(|_| prepare_append_writes(self.wb(cx), plan.id, plan.count, &writes))
            .and_then(|(candidate, history)| {
                self.validate_saved_view_layout(&candidate)?;
                Ok((candidate, history))
            });
        match result {
            Ok((candidate, history)) => {
                let index = candidate
                    .sheet_index_by_id(history.table.sheet_id())
                    .unwrap();
                let range = history.table.after_table().unwrap().range;
                let first_new = range.end_row + 1 - plan.count;
                let view = candidate
                    .sheet(index)
                    .unwrap()
                    .build_saved_table_view(
                        crate::app::NUM_ROWS.min(candidate.sheet(index).unwrap().rows),
                    )
                    .unwrap();
                let hidden = view.as_ref().map_or(0, |v| {
                    (first_new..=range.end_row)
                        .filter(|r| !v.rows().is_data_row_visible(*r))
                        .count()
                });
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                let first = (writes[0].row, writes[0].col);
                let focus = self.row_view.data_to_view(first.0).unwrap_or_else(|| {
                    let visible = self.row_view.visible_rows();
                    let end = visible.partition_point(|r| *r <= range.end_row);
                    end.checked_sub(1)
                        .map(|i| visible[i])
                        .filter(|r| *r > range.start_row)
                        .unwrap_or(range.start_row)
                });
                self.view_state.select_cell(focus, first.1);
                self.view_state.additional_selections.clear();
                self.ensure_visible(cx);
                let description = format!("Paste and append {} Table row(s)", plan.count);
                self.history.record_action_with_provenance(
                    UndoAction::TableAppend {
                        sheet_index: index,
                        history: Box::new(history),
                        description: description.clone(),
                    },
                    None,
                );
                self.bump_cells_rev();
                self.is_modified = true;
                self.clipboard_visual_range = None;
                self.status_message = Some(if hidden > 0 {
                    format!("{description}; {hidden} new row(s) hidden by the current view.")
                } else {
                    description
                });
                cx.notify();
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        clipboard::{table_paste_writes, TablePasteKind},
        history::History,
        table_edit::tests::fixture,
    };
    use visigrid_engine::{
        cell::{CellComment, NumberFormat},
        sheet::SheetId,
    };

    fn book() -> (Workbook, TableId) {
        let mut wb = fixture(true);
        let id = wb.active_sheet().tables()[0].id;
        wb.set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        wb.set_cell_value_tracked(0, 4, 3, "999");
        wb.set_cell_value_tracked(0, 10, 1, "Notes stay below");
        (wb, id)
    }

    fn plan(wb: &Workbook, start: (usize, usize), h: usize, w: usize) -> BulkAppendPlan {
        let view = wb.active_sheet().build_saved_table_view(30).unwrap();
        let rows = view
            .map(|v| v.rows().clone())
            .unwrap_or_else(|| RowView::new(30));
        plan_bulk_append(wb.active_sheet(), &rows, start, h, w)
            .unwrap()
            .unwrap()
    }

    #[test]
    fn bulk_append_moves_hidden_footer_without_writing_existing_hidden_records() {
        let (mut before, id) = book();
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        let (before, _) = before.prepare_table_row_visibility(before.active_sheet_id(), [7, 9].into()).unwrap();
        let append = plan(&before, (6, 1), 3, 2);
        assert_eq!(append.count, 2);
        let writes = table_paste_writes(
            &grid(&[&["West", "5"], &["West", "7"], &["West", "11"]]),
            None, TablePasteKind::Contents, append.targets,
        );
        let (after, history) = prepare_append_writes(&before, id, 2, &writes).unwrap();
        assert_eq!(after.table(id).unwrap().1.totals_row(), Some(9));
        assert_eq!(after.active_sheet().manual_hidden_rows(), [7, 9].into());
        assert_eq!(after.active_sheet().get_display(7, 3), "14");
        assert_eq!(after.active_sheet().get_display(8, 3), "22");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        let view = after.active_sheet().build_saved_table_view(30).unwrap().unwrap();
        assert!(view.rows().data_to_view(7).is_none());
        assert!(view.rows().data_to_view(8).is_some());
        assert!(prepare_append_writes(&before, id, 2, &[TableCellWrite::value(4, 2, "99".into())]).is_err());
        let undone = history.replay(&after, true).unwrap();
        assert_eq!(undone.active_sheet().get_raw(6, 2), before.active_sheet().get_raw(6, 2));
        let redone = history.replay(&undone, false).unwrap();
        assert_eq!(redone.active_sheet().manual_hidden_rows(), [7, 9].into());
        assert_eq!(redone.active_sheet().get_display(9, 3), after.active_sheet().get_display(9, 3));
    }

    fn grid(rows: &[&[&str]]) -> Vec<Vec<String>> {
        rows.iter()
            .map(|r| r.iter().map(|s| s.to_string()).collect())
            .collect()
    }

    #[test]
    fn linked_footer_and_overflow_paste_replay_in_one_history_entry() {
        let (mut before, id) = book();
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        before.set_cell_value_tracked(0, 0, 0, "=D8");
        let plan = plan(&before, (4, 1), 5, 2);
        let values = grid(&[ &["West", "11"], &["West", "12"], &["West", "13"], &["West", "14"], &["East", "15"] ]);
        let writes = table_paste_writes(&values, None, TablePasteKind::Contents, plan.targets);
        let (after, entry) = prepare_append_writes(&before, id, plan.count, &writes).unwrap();
        assert_eq!(after.active_sheet().get_raw(0, 0), "=D10");
        assert_eq!(after.active_sheet().get_display(0, 0), "100");
        let undone = entry.replay(&after, true).unwrap();
        assert_eq!(undone.active_sheet().get_raw(0, 0), "=D8");
        assert_eq!(undone.active_sheet().get_raw(4, 3), "999");
        let redone = entry.replay(&undone, false).unwrap();
        assert_eq!(redone.active_sheet().get_display(0, 0), "100");
    }

    #[test]
    fn dynamic_footer_readers_survive_filtered_overflow_paste_and_replay() {
        let (mut before, id) = book();
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        before.set_cell_value_tracked(0, 0, 0, "=OFFSET(D8,0,0)");
        before.set_cell_value_tracked(0, 1, 0, "=INDIRECT(\"D8\")");
        let plan = plan(&before, (4, 1), 5, 2);
        let values = grid(&[ &["West", "11"], &["West", "12"], &["West", "13"], &["West", "14"], &["East", "15"] ]);
        let writes = table_paste_writes(&values, None, TablePasteKind::Contents, plan.targets);
        let (after, entry) = prepare_append_writes(&before, id, plan.count, &writes).unwrap();
        assert_eq!(after.active_sheet().get_raw(0, 0), "=OFFSET(D10, 0, 0)");
        assert_eq!(after.active_sheet().get_display(0, 0), "100");
        assert_eq!(after.active_sheet().get_display(1, 0), "28");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        let undone = entry.replay(&after, true).unwrap();
        for row in [0, 1] {
            assert_eq!(undone.active_sheet().get_raw(row, 0), before.active_sheet().get_raw(row, 0));
            assert_eq!(undone.active_sheet().get_display(row, 0), before.active_sheet().get_display(row, 0));
        }
        let redone = entry.replay(&undone, false).unwrap();
        assert_eq!(redone.active_sheet().get_display(0, 0), "100");
        assert_eq!(redone.active_sheet().get_display(1, 0), "28");
        assert_eq!(redone.active_sheet().table_view_spec(), before.active_sheet().table_view_spec());
    }

    #[test]
    fn footer_rules_follow_filtered_overflow_paste_and_replay() {
        use visigrid_engine::{cond_format::CondStyle, validation::{CellRange, ValidationRule}};
        let (mut before, id) = book();
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        before.active_sheet_mut().validations.set(CellRange::single(7, 3), ValidationRule::custom("=$D$8>0"));
        before.active_sheet_mut().cond_formats.add(vec![CellRange::single(7, 3)], "=$D$8>0", CondStyle::Inline(Default::default()));
        let plan = plan(&before, (4, 1), 5, 2);
        let values = grid(&[ &["West", "11"], &["West", "12"], &["West", "13"], &["West", "14"], &["East", "15"] ]);
        let writes = table_paste_writes(&values, None, TablePasteKind::Contents, plan.targets);
        let (after, entry) = prepare_append_writes(&before, id, plan.count, &writes).unwrap();
        let undone = entry.replay(&after, true).unwrap();
        let redone = entry.replay(&undone, false).unwrap();
        for (wb, row) in [(&after, 9), (&undone, 7), (&redone, 9)] {
            assert!(wb.active_sheet().validations.has_validation(row, 3));
            assert_eq!(wb.active_sheet().cond_formats.iter().next().unwrap().predicate_at(row, 3), Some(format!("=$D${}>0", row + 1)));
            assert_eq!(wb.active_sheet().get_raw(4, 3), "999");
            assert_eq!(wb.active_sheet().table_view_spec(), before.active_sheet().table_view_spec());
        }
        assert_eq!(after.active_sheet().get_display(9, 3), "100");
        assert!(!after.active_sheet().validations.has_validation(7, 3));
    }

    #[test]
    fn filtered_overflow_paste_moves_footer_and_replays_all_values() {
        let (mut before, id) = book();
        before.set_table_totals_visible(id, true, Default::default()).unwrap();
        let footer = before.active_sheet().get_raw(7, 3);
        let plan = plan(&before, (4, 1), 5, 2);
        let values = grid(&[ &["West", "11"], &["West", "12"], &["West", "13"], &["West", "14"], &["East", "15"] ]);
        let writes = table_paste_writes(&values, None, TablePasteKind::Contents, plan.targets);
        let (after, entry) = prepare_append_writes(&before, id, plan.count, &writes).unwrap();
        assert_eq!(after.table(id).unwrap().1.totals_row(), Some(9));
        assert_eq!(after.active_sheet().get_raw(9, 3), footer);
        assert_eq!(after.active_sheet().get_display(9, 3), "100");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().get_raw(10, 1), "Notes stay below");
        let restored = entry.replay(&after, true).unwrap();
        for row in 2..=10 { for col in 1..=3 {
            assert_eq!(restored.active_sheet().get_raw(row, col), before.active_sheet().get_raw(row, col));
        } }
        let redone = entry.replay(&restored, false).unwrap();
        assert_eq!(redone.active_sheet().get_display(9, 3), "100");
        let rows = before.active_sheet().build_saved_table_view(30).unwrap().unwrap();
        assert!(plan_bulk_append(before.active_sheet(), rows.rows(), (7, 1), 1, 2).is_err());
    }

    #[test]
    fn overflow_maps_visible_records_then_appends_without_touching_hidden_values() {
        let (before, id) = book();
        let plan = plan(&before, (4, 1), 5, 2);
        assert_eq!(plan.count, 2);
        assert_eq!(
            plan.targets
                .iter()
                .step_by(2)
                .map(|t| t.0)
                .collect::<Vec<_>>(),
            [5, 3, 6, 7, 8]
        );
        let grid = grid(&[
            &["West", "11"],
            &["West", "12"],
            &["West", "13"],
            &["West", "14"],
            &["East", "15"],
        ]);
        let writes = table_paste_writes(&grid, None, TablePasteKind::Contents, plan.targets);
        let (after, history) = prepare_append_writes(&before, id, plan.count, &writes).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 8);
        assert_eq!(after.active_sheet().get_raw(4, 2), "10");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(after.active_sheet().get_display(7, 3), "28");
        assert_eq!(after.active_sheet().get_display(8, 3), "30");
        assert_eq!(after.active_sheet().get_raw(10, 1), "Notes stay below");
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.rows().is_data_row_visible(7));
        assert!(!view.rows().is_data_row_visible(8));
        let undo = history.replay(&after, true).unwrap();
        assert_eq!(undo.table(id).unwrap().1.range.end_row, 6);
        for r in 3..=8 {
            for c in 1..=3 {
                assert_eq!(
                    undo.active_sheet().get_raw(r, c),
                    before.active_sheet().get_raw(r, c)
                );
            }
        }
        let redo = history.replay(&undo, false).unwrap();
        assert_eq!(redo.active_sheet().get_display(7, 3), "28");
        assert_eq!(redo.active_sheet().get_raw(8, 1), "East");
    }

    #[test]
    fn below_table_paste_appends_all_rows_even_when_they_are_filtered_out() {
        let (before, id) = book();
        let plan = plan(&before, (7, 1), 2, 2);
        assert_eq!(plan.count, 2);
        let writes = table_paste_writes(
            &grid(&[&["East", "5"], &["East", "7"]]),
            None,
            TablePasteKind::Values,
            plan.targets,
        );
        let (after, history) = prepare_append_writes(&before, id, 2, &writes).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(!view.rows().is_data_row_visible(7));
        assert!(!view.rows().is_data_row_visible(8));
        assert_eq!(after.active_sheet().get_display(8, 3), "14");
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .table(id)
                .unwrap()
                .1
                .range
                .end_row,
            6
        );
    }

    #[test]
    fn explicit_blanks_override_calculated_formulas_and_replay_with_comments_formats_and_literal_text(
    ) {
        let (before, id) = book();
        let mut writes = vec![
            TableCellWrite::value(7, 1, "=not a formula".into()),
            TableCellWrite::value(7, 3, String::new()),
        ];
        writes[0].literal_text = true;
        writes[0].comment = Some(Some(CellComment {
            text: "Copied note".into(),
            author: "QA".into(),
        }));
        let mut format = before.active_sheet().get_format(7, 1);
        format.number_format = NumberFormat::Custom("@".into());
        writes[0].format = Some(format.clone());
        let (after, history) = prepare_append_writes(&before, id, 1, &writes).unwrap();
        assert_eq!(after.active_sheet().get_display(7, 1), "=not a formula");
        assert_eq!(after.active_sheet().get_raw(7, 3), "");
        assert_eq!(after.active_sheet().get_format(7, 1), format);
        let undo = history.replay(&after, true).unwrap();
        assert!(undo.active_sheet().comment(7, 1).is_none());
        let redo = history.replay(&undo, false).unwrap();
        assert_eq!(
            redo.active_sheet().comment(7, 1).unwrap().text,
            "Copied note"
        );
        assert_eq!(redo.active_sheet().get_raw(7, 3), "");
        assert_eq!(redo.active_sheet().get_format(7, 1), format);
    }

    #[test]
    fn source_formulas_are_rebased_to_each_canonical_destination() {
        let (before, id) = book();
        let plan = plan(&before, (6, 3), 3, 1);
        assert_eq!(
            plan.targets.iter().map(|t| t.0).collect::<Vec<_>>(),
            [6, 7, 8]
        );
        let internal = crate::clipboard::InternalClipboard {
            raw_tsv: "=C9*3\n=C4*3\n=C6*3".into(),
            raw_cells: grid(&[&["=C9*3"], &["=C4*3"], &["=C6*3"]]),
            values: vec![vec![visigrid_engine::formula::eval::Value::Number(0.0)]; 3],
            formats: vec![vec![Default::default()]; 3],
            comments: vec![vec![None]; 3],
            source: (8, 3),
            source_rows: vec![8, 3, 5],
            source_formulas: vec![vec![true]; 3],
            id: 1,
            merges: vec![],
            created_at: std::time::Instant::now(),
        };
        let writes = table_paste_writes(
            &internal.raw_cells,
            Some(&internal),
            TablePasteKind::Formulas,
            plan.targets,
        );
        let (after, history) = prepare_append_writes(&before, id, 2, &writes).unwrap();
        assert_eq!(after.active_sheet().get_display(6, 3), "120");
        assert_eq!(after.active_sheet().get_raw(7, 3), "=C8*3");
        assert_eq!(after.active_sheet().get_raw(8, 3), "=C9*3");
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .active_sheet()
                .get_raw(6, 3),
            "=[@Amount]*2"
        );
    }

    #[test]
    fn no_growth_when_paste_fits_and_side_crossing_or_limits_refuse() {
        let (wb, _) = book();
        let view = wb
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        let rows = view.rows();
        assert!(plan_bulk_append(wb.active_sheet(), rows, (4, 1), 3, 2)
            .unwrap()
            .is_none());
        assert!(plan_bulk_append(wb.active_sheet(), rows, (7, 0), 2, 3).is_err());
        assert!(plan_bulk_append(wb.active_sheet(), rows, (6, 3), 3, 2).is_err());
        assert!(plan_bulk_append(wb.active_sheet(), rows, (7, 1), 24, 1).is_err());
        assert!(plan_bulk_append(wb.active_sheet(), rows, (7, 1), 100_001, 1).is_err());
        assert!(plan_bulk_append(wb.active_sheet(), rows, (9, 1), 1, 1)
            .unwrap()
            .is_none());
    }

    #[test]
    fn collision_refuses_entire_paste_including_existing_visible_targets() {
        for beside in [false, true] {
            let (mut before, id) = book();
            before.set_cell_value_tracked(0, 8, if beside { 5 } else { 2 }, "Keep me");
            let writes = vec![
                TableCellWrite::value(6, 2, "777".into()),
                TableCellWrite::value(8, 2, "55".into()),
            ];
            let rev = before.revision();
            assert!(prepare_append_writes(&before, id, 2, &writes).is_err());
            assert_eq!(before.revision(), rev);
            assert_eq!(before.active_sheet().get_raw(6, 2), "40");
            assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
        }
    }

    #[test]
    fn hidden_existing_target_and_duplicate_targets_are_refused() {
        let (before, id) = book();
        assert!(
            prepare_append_writes(&before, id, 1, &[TableCellWrite::value(4, 2, "99".into())])
                .is_err()
        );
        let w = TableCellWrite::value(7, 2, "99".into());
        assert!(prepare_append_writes(&before, id, 1, &[w.clone(), w]).is_err());
    }

    #[test]
    fn append_paste_rewinds_and_survives_native_save() {
        let (before, id) = book();
        let writes = vec![
            TableCellWrite::value(7, 1, "West".into()),
            TableCellWrite::value(7, 2, "12".into()),
        ];
        let (after, commit) = prepare_append_writes(&before, id, 1, &writes).unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableAppend {
                sheet_index: 0,
                history: Box::new(commit),
                description: "Paste and append".into(),
            },
            None,
        );
        assert_eq!(history.undo_count(), 1);
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(preview.workbook.active_sheet().get_display(7, 3), "24");
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bulk.sheet");
        visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
            .unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(loaded.table(id).unwrap().1.range.end_row, 7);
        assert_eq!(loaded.active_sheet().get_display(7, 3), "24");
        assert_eq!(
            loaded.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
    }

    #[test]
    fn changed_view_and_stale_paste_cells_block_replay() {
        let (before, id) = book();
        let (mut after, commit) =
            prepare_append_writes(&before, id, 1, &[TableCellWrite::value(7, 2, "12".into())])
                .unwrap();
        after.set_cell_value_tracked(0, 7, 2, "99");
        assert!(commit.replay(&after, true).is_err());
        assert_eq!(after.table(id).unwrap().1.range.end_row, 7);
        after.set_table_view_spec(SheetId(7), None).unwrap();
        assert!(commit
            .replay(&after, true)
            .unwrap_err()
            .contains("view changed"));
    }
    #[test]
    fn paste_recalculation_cannot_spill_beside_a_filtered_table() {
        let (mut before, id) = book();
        before.set_cell_value_tracked(0, 0, 0, "=IF(SUM(Sales[Amount])>100,SEQUENCE(5),0)");
        before.active_sheet().build_saved_table_view(30).unwrap();
        let writes = vec![
            TableCellWrite::value(7, 1, "West".into()),
            TableCellWrite::value(7, 2, "100".into()),
        ];
        assert!(prepare_append_writes(&before, id, 1, &writes).is_err());
        assert_eq!(before.table(id).unwrap().1.range.end_row, 6);
        assert_eq!(before.active_sheet().get_display(0, 0), "0");
    }

    #[test]
    fn header_only_table_accepts_bulk_paste_below_header() {
        let (mut before, id) = book();
        for r in 3..=6 {
            for c in 1..=3 {
                before.clear_cell_tracked(0, r, c);
            }
        }
        let mut range = before.table(id).unwrap().1.range;
        range.end_row = range.start_row;
        before.resize_table(id, range).unwrap();
        let plan = plan(&before, (3, 1), 2, 2);
        assert_eq!(plan.count, 2);
        let writes = table_paste_writes(
            &grid(&[&["West", "5"], &["East", "7"]]),
            None,
            TablePasteKind::Contents,
            plan.targets,
        );
        let (after, history) = prepare_append_writes(&before, id, 2, &writes).unwrap();
        assert_eq!(after.active_sheet().get_display(3, 3), "10");
        assert_eq!(after.active_sheet().get_display(4, 3), "14");
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .table(id)
                .unwrap()
                .1
                .range
                .data_rows(),
            0
        );
    }
    #[test]
    fn other_table_growth_preserves_the_active_projection_and_fit_pastes() {
        let (mut before, _) = book();
        before.set_cell_value_tracked(0, 12, 1, "Value");
        let id = before
            .create_table(
                SheetId(7),
                TableRange {
                    start_row: 12,
                    end_row: 14,
                    start_col: 1,
                    end_col: 1,
                },
                "Other",
            )
            .unwrap()
            .table_id();
        let view = before
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        // A fitting paste on this unprojected Table retains ordinary cross-column behavior.
        assert!(
            plan_bulk_append(before.active_sheet(), view.rows(), (13, 1), 1, 2)
                .unwrap()
                .is_none()
        );
        let plan = plan_bulk_append(before.active_sheet(), view.rows(), (14, 1), 3, 1)
            .unwrap()
            .unwrap();
        assert_eq!(plan.count, 2);
        let writes = table_paste_writes(
            &grid(&[&["1"], &["2"], &["3"]]),
            None,
            TablePasteKind::Values,
            plan.targets,
        );
        let (after, history) = prepare_append_writes(&before, id, 2, &writes).unwrap();
        assert_eq!(after.table(id).unwrap().1.range.end_row, 16);
        assert_eq!(after.active_sheet().get_raw(16, 1), "3");
        assert_eq!(after.active_sheet().get_raw(4, 3), "999");
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        assert_eq!(
            history
                .replay(&after, true)
                .unwrap()
                .table(id)
                .unwrap()
                .1
                .range
                .end_row,
            14
        );
    }
}
