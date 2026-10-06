//! Atomic cell writes through a Table projection. Plan against the old view,
//! recalculate a candidate, and publish only if every saved view is still safe.
use crate::{
    app::{Spreadsheet, NUM_ROWS},
    history::UndoAction,
    table_cell_history::TableCellsCommit,
};
use gpui::*;
use visigrid_engine::{
    cell::{CellComment, CellFormat},
    workbook::Workbook,
};

/// Point-picked references address canonical cells. A range retains ordinary
/// A1 semantics: it spans its canonical endpoints, including hidden records.
pub(crate) fn table_formula_reference(
    rows: &visigrid_engine::filter::RowView,
    start: (usize, usize),
    end: Option<(usize, usize)>,
) -> String {
    let start = (rows.view_to_data(start.0), start.1);
    match end {
        Some((row, col)) => Spreadsheet::make_range_ref(start, (rows.view_to_data(row), col)),
        None => Spreadsheet::make_cell_ref(start.0, start.1),
    }
}

/// Keep a filtered-out edit near its old on-screen position, within the
/// remaining records. Fall back to the header only when no records remain.
fn focus_after_edit(
    rows: &visigrid_engine::filter::RowView,
    range: visigrid_engine::table::TableRange,
    old_slot: usize,
    record: usize,
) -> usize {
    rows.data_to_view(record).unwrap_or_else(|| {
        rows.visible_rows()
            .iter()
            .copied()
            .filter(|r| *r > range.start_row && *r <= range.end_row)
            .min_by_key(|r| (r.abs_diff(old_slot), *r < old_slot))
            .unwrap_or(range.start_row)
    })
}

#[derive(Clone, Debug)]
pub(crate) struct TableCellWrite {
    pub row: usize,
    pub col: usize,
    pub value: Option<String>,
    pub literal_text: bool,
    pub format: Option<CellFormat>,
    pub comment: Option<Option<CellComment>>,
}

impl TableCellWrite {
    pub(crate) fn value(row: usize, col: usize, value: String) -> Self {
        Self {
            row,
            col,
            value: Some(value),
            literal_text: false,
            format: None,
            comment: None,
        }
    }
}

/// Validate canonical cells in a workbook with saved Table views. History may
/// revisit hidden records because an earlier edit can itself hide its target.
pub(crate) fn validate_view_safe_targets(
    wb: &Workbook,
    sheet_index: usize,
    targets: &[(usize, usize)],
    allow_hidden: bool,
) -> Result<(), String> {
    wb.ensure_writable()?;
    let sheet = wb.sheet(sheet_index).ok_or("The sheet no longer exists.")?;
    let view = sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows))?;
    for &(row, col) in targets {
        if row >= sheet.rows || col >= sheet.cols {
            return Err("The edit extends beyond the worksheet. Nothing was changed.".into());
        }
        if let Some(view) = &view {
            if !allow_hidden && !view.rows().is_data_row_visible(row) {
                return Err("Cannot write to a hidden Table record.".into());
            }
        }
        if sheet.is_pivot_owned(row, col) {
            return Err("Cannot edit PivotTable output. Edit its source data instead.".into());
        }
        if sheet.get_spill_parent(row, col).is_some() {
            return Err("Cannot edit a spill receiver. Edit the source formula instead.".into());
        }
        if sheet
            .get_merge(row, col)
            .is_some_and(|m| m.start != (row, col))
        {
            return Err("The edit includes covered merged cells. Edit the merged cell's top-left cell or unmerge first.".into());
        }
        if let Some(error) = sheet.table_value_write_error(row, col) {
            return Err(error);
        }
    }
    if let Some(view) = view {
        view.validate_mutation_ranges(
            sheet,
            &targets
                .iter()
                .map(|&(row, col)| visigrid_engine::validation::CellRange {
                    start_row: row,
                    end_row: row,
                    start_col: col,
                    end_col: col,
                })
                .collect::<Vec<_>>(),
        )?;
    }
    Ok(())
}

/// A paste beginning in a Table body stays in that body. Other destinations
/// use visible worksheet rows, then share the same all-target preflight.
pub(crate) fn view_safe_paste_targets(
    sheet: &visigrid_engine::sheet::Sheet,
    rows: &visigrid_engine::filter::RowView,
    start: (usize, usize),
    height: usize,
    width: usize,
) -> Result<Vec<(usize, usize, usize, usize)>, String> {
    if height == 0 || width == 0 || height.saturating_mul(width) > 100_000 {
        return Err("Paste between 1 and 100,000 cells at a time.".into());
    }
    if start
        .1
        .checked_add(width)
        .is_none_or(|end| end > sheet.cols)
    {
        return Err("The paste extends beyond the worksheet's columns.".into());
    }
    let view = sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows))?;
    let in_body = view.as_ref().filter(|v| {
        let r = v.range();
        start.0 > r.start_row && start.0 <= r.end_row
    });
    let data_rows = if let Some(view) = in_body {
        let range = view.range();
        if start.1 < range.start_col || start.1 + width > range.end_col + 1 {
            return Err("The paste exceeds the Table's columns. Nothing was pasted.".into());
        }
        view.visible_body_rows(start.0, height)?
    } else {
        let first = rows
            .visible_index_of(start.0)
            .ok_or("Select a visible cell before pasting.")?;
        let data: Vec<_> = rows
            .visible_rows()
            .iter()
            .skip(first)
            .take(height)
            .map(|&r| rows.view_to_data(r))
            .collect();
        if data.len() != height || data.iter().any(|&r| r >= sheet.rows) {
            return Err("The paste extends beyond the worksheet's rows.".into());
        }
        data
    };
    Ok(data_rows
        .into_iter()
        .enumerate()
        .flat_map(|(ri, r)| (0..width).map(move |ci| (r, start.1 + ci, ri, ci)))
        .collect())
}

/// Resolve a cut/fill rectangle once in display order. These operations do not
/// transfer merges or cross the boundary between a projected Table and notes.
/// Sources receive the same protection as destinations, even when not written.
pub(crate) fn view_safe_selection_rows(
    sheet: &visigrid_engine::sheet::Sheet,
    rows: &visigrid_engine::filter::RowView,
    rect: ((usize, usize), (usize, usize)),
) -> Result<Vec<usize>, String> {
    let ((r0, c0), (r1, c1)) = rect;
    if r0 > r1 || c0 > c1 || r1 >= rows.row_count() || c1 >= sheet.cols {
        return Err("The selection extends beyond the worksheet.".into());
    }
    if let Some(view) = sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows))? {
        let range = view.range();
        if r0 <= range.end_row
            && r1 > range.start_row
            && (r0 <= range.start_row
                || r1 > range.end_row
                || c0 < range.start_col
                || c1 > range.end_col)
        {
            return Err(
                "Cut and fill cannot cross the Table body boundary or include cells beside it."
                    .into(),
            );
        }
    }
    let data: Vec<_> = rows
        .visible_rows()
        .iter()
        .copied()
        .filter(|r| *r >= r0 && *r <= r1)
        .map(|r| rows.view_to_data(r))
        .collect();
    if data.is_empty() {
        return Err("Select at least one visible cell.".into());
    }
    if data.len().saturating_mul(c1 - c0 + 1) > 100_000 {
        return Err("Cut or fill at most 100,000 visible cells at once.".into());
    }
    for &row in &data {
        if row >= sheet.rows {
            return Err("The selection extends beyond the worksheet.".into());
        }
        for col in c0..=c1 {
            if sheet.get_merge(row, col).is_some() {
                return Err("Unmerge source and destination cells before cutting or filling through a Table view.".into());
            }
            if sheet.is_pivot_owned(row, col) {
                return Err("Cut and fill cannot include PivotTable output.".into());
            }
            if sheet.get_spill_parent(row, col).is_some() {
                return Err("Cut and fill cannot include spill receivers.".into());
            }
            if let Some(error) = sheet.table_value_write_error(row, col) {
                return Err(error);
            }
        }
    }
    Ok(data)
}

pub(crate) fn prepare_table_writes(
    wb: &Workbook,
    sheet_index: usize,
    writes: &[TableCellWrite],
) -> Result<Workbook, String> {
    prepare_table_writes_with_new_rows(wb, sheet_index, writes, None)
}

/// Only the newly allocated append stripe may receive writes while hidden.
/// Existing records still require visibility in the pre-paste projection.
pub(crate) fn prepare_table_writes_with_new_rows(
    wb: &Workbook,
    sheet_index: usize,
    writes: &[TableCellWrite],
    new_rows: Option<visigrid_engine::table::TableRange>,
) -> Result<Workbook, String> {
    let targets: Vec<_> = writes.iter().map(|w| (w.row, w.col)).collect();
    validate_view_safe_targets(wb, sheet_index, &targets, new_rows.is_some())?;
    if let Some(new_rows) = new_rows {
        let existing: Vec<_> = targets.iter().copied()
            .filter(|&(r, c)| !new_rows.contains(r, c)).collect();
        validate_view_safe_targets(wb, sheet_index, &existing, false)?;
    }
    let mut candidate = wb.clone();
    {
        let mut batch = candidate.batch_guard();
        for write in writes {
            if let Some(value) = &write.value {
                if write.literal_text {
                    batch.set_cell_text_exact_tracked(sheet_index, write.row, write.col, value);
                } else {
                    batch.set_cell_value_tracked(sheet_index, write.row, write.col, value);
                }
            }
            let sheet = batch.sheet_mut(sheet_index).unwrap();
            if let Some(format) = &write.format {
                sheet.set_format(write.row, write.col, format.clone());
            }
            if let Some(comment) = &write.comment {
                sheet.set_comment(write.row, write.col, comment.clone());
            }
        }
    }
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!("The edit could not be recalculated: {error:?}"));
    }
    // A new spill must not silently consume another target from this batch.
    validate_view_safe_targets(
        &candidate,
        sheet_index,
        &writes.iter().map(|w| (w.row, w.col)).collect::<Vec<_>>(),
        true,
    )?;
    // Recalculation can change a spill or a formula-backed key on another sheet.
    for sheet in candidate.sheets() {
        sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows))?;
    }
    Ok(candidate)
}

impl Spreadsheet {
    pub(crate) fn table_cell_edit_guard(&mut self, cx: &mut Context<Self>) -> bool {
        self.table_edit_target = None;
        if self.block_live_read_only(cx) { return true; }
        if self.cloud_live_enabled() && crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.status_message = Some("Live Table editing is not enabled in this cell-editing slice".into());
            cx.notify(); return true;
        }
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return true;
        }
        if !crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            return false;
        }
        self.sync_table_view(cx);
        if let Some(table) = self
            .sheet(cx)
            .table_view_spec()
            .and_then(|v| self.sheet(cx).tables().iter().find(|t| t.id == v.table))
        {
            if let Err(error) = self.table_layout_check(table) {
                self.status_message = Some(error);
                cx.notify();
                return true;
            }
        }
        let (view_row, col) = self.view_state.selected;
        let row = self.row_view.view_to_data(view_row);
        let (row, col) = self
            .sheet(cx)
            .get_merge(row, col)
            .map_or((row, col), |m| m.start);
        let result = self.validate_saved_view_layout(self.wb(cx)).and_then(|_| {
            validate_view_safe_targets(self.wb(cx), self.sheet_index(cx), &[(row, col)], false)
        });
        match result {
            Ok(_) => {
                self.view_state.selected = (self.row_view.data_to_view(row).unwrap_or(row), col);
                self.table_edit_target =
                    Some((self.sheet_index(cx), row, col, self.wb(cx).revision()));
                false
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                true
            }
        }
    }

    pub(crate) fn apply_table_cell_writes(
        &mut self,
        writes: Vec<TableCellWrite>,
        description: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return false;
        }
        if writes.is_empty() {
            return false;
        }
        if let Some(table) = self
            .sheet(cx)
            .table_view_spec()
            .and_then(|v| self.sheet(cx).tables().iter().find(|t| t.id == v.table))
        {
            if let Err(error) = self.table_layout_check(table) {
                self.status_message = Some(error);
                cx.notify();
                return false;
            }
        }
        let candidate = match prepare_table_writes(self.wb(cx), self.sheet_index(cx), &writes)
            .and_then(|candidate| {
                self.validate_saved_view_layout(&candidate)?;
                Ok(candidate)
            }) {
            Ok(candidate) => candidate,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return false;
            }
        };
        let commit = TableCellsCommit::capture(
            self.sheet(cx),
            candidate.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        if commit.patches.is_empty() {
            return true;
        }
        let old_focus = self.view_state.selected;
        let record = self.row_view.view_to_data(old_focus.0);
        let range = self
            .sheet(cx)
            .tables()
            .iter()
            .find(|t| Some(t.id) == self.sheet(cx).table_view_spec().map(|s| s.table))
            .map(|table| table.range);
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.sync_table_view(cx);
        let focus = if let Some(range) = range {
            focus_after_edit(&self.row_view, range, old_focus.0, record)
        } else {
            self.row_view.data_to_view(record).unwrap_or(old_focus.0)
        };
        self.view_state.select_cell(focus, old_focus.1);
        self.ensure_visible(cx);
        self.history.record_action_with_provenance(
            UndoAction::TableCellsChanged {
                sheet_index: self.sheet_index(cx),
                commit: Box::new(commit),
                description: description.into(),
            },
            None,
        );
        self.bump_cells_rev();
        self.is_modified = true;
        self.clipboard_visual_range = None;
        cx.notify();
        true
    }

    pub(crate) fn replay_table_cells(
        &mut self,
        commit: &TableCellsCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let Some(index) = self.wb(cx).sheet_index_by_id(commit.sheet) else {
            self.status_message = Some("The history sheet no longer exists.".into());
            cx.notify();
            return false;
        };
        let sheet = self.wb(cx).sheet(index).unwrap();
        if let Some(spec) = sheet.table_view_spec().filter(|v| v.has_criteria()) {
            if let Some(table) = sheet.tables().iter().find(|t| t.id == spec.table) {
                if let Some(error) = crate::table_filter_ui::desktop_layout_error(
                    table,
                    self.row_heights.get(&sheet.id),
                    self.hidden_rows.get(&sheet.id),
                    sheet.frozen_panes.0,
                ) {
                    self.status_message = Some(error);
                    cx.notify();
                    return false;
                }
            }
        }
        let result = self
            .validate_saved_view_layout(self.wb(cx))
            .and_then(|_| self.workbook.update(cx, |wb, _| commit.replay(wb, undo)));
        if let Err(error) = result {
            self.status_message = Some(error);
            cx.notify();
            return false;
        }
        if self.sheet_index(cx) != index {
            self.activate_sheet(index, cx);
        }
        self.sync_table_view(cx);
        self.bump_cells_rev();
        cx.notify();
        true
    }

    /// Canonical targets for a single-cell broadcast or Delete. Skip hidden
    /// slots and validate the complete selection before changing any record.
    pub(crate) fn table_selection_targets(&self, cx: &App) -> Result<Vec<(usize, usize)>, String> {
        let mut targets = std::collections::BTreeSet::new();
        for ((r1, c1), (r2, c2)) in self.all_selection_ranges() {
            for row in r1..=r2 {
                if !self.row_view.is_view_row_visible(row) {
                    continue;
                }
                for col in c1..=c2 {
                    if targets.len() >= 100_000 {
                        return Err("Select at most 100,000 cells for one edit.".into());
                    }
                    targets.insert((self.row_view.view_to_data(row), col));
                }
            }
        }
        let _ = cx;
        Ok(targets.into_iter().collect())
    }

    pub(crate) fn delete_table_selection(&mut self, cx: &mut Context<Self>) {
        match self.table_selection_targets(cx) {
            Ok(targets) => {
                self.apply_table_cell_writes(
                    targets
                        .into_iter()
                        .map(|(r, c)| TableCellWrite::value(r, c, String::new()))
                        .collect(),
                    "Clear cells",
                    cx,
                );
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
            }
        }
    }
}

#[cfg(test)]
pub(crate) mod tests {
    use super::{prepare_table_writes, TableCellWrite};
    use crate::table_cell_history::TableCellsCommit;
    use visigrid_engine::{
        filter::{ColumnFilter, FilterKey, SortDirection},
        formula::eval::Value,
        sheet::{Sheet, SheetId},
        table::TableRange,
        table_view::{TableFilter, TableSort, TableViewSpec},
        workbook::Workbook,
    };

    pub(crate) fn fixture(filtered: bool) -> Workbook {
        let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 30, 8)], 0);
        for (r, values) in [
            ["Group", "Amount", "Result"],
            ["West", "30", "=C4*2"],
            ["East", "10", "=C5*2"],
            ["West", "20", "=C6*2"],
            ["West", "40", "=C7*2"],
        ]
        .iter()
        .enumerate()
        {
            for (c, v) in values.iter().enumerate() {
                wb.set_cell_value_tracked(0, r + 2, c + 1, v);
            }
        }
        wb.set_cell_value_tracked(0, 0, 1, "=SUM(C4:C7)");
        let id = wb
            .create_table(
                SheetId(7),
                TableRange {
                    start_row: 2,
                    end_row: 6,
                    start_col: 1,
                    end_col: 3,
                },
                "Sales",
            )
            .unwrap()
            .table_id();
        let table = wb.table(id).unwrap().1;
        let mut spec = TableViewSpec::new(id);
        spec.sort = Some(TableSort {
            column: table.columns[1].id,
            direction: SortDirection::Ascending,
        });
        if filtered {
            spec.filters.push(TableFilter {
                column: table.columns[0].id,
                criteria: ColumnFilter {
                    selected: Some([FilterKey::Text("West".into()).normalized()].into()),
                    text_filter: None,
                },
            });
        }
        wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        wb
    }

    #[test]
    fn point_picking_uses_canonical_addresses_after_sort_and_filter() {
        let wb = fixture(true);
        let view = wb
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            super::table_formula_reference(view.rows(), (4, 2), None),
            "C6"
        );
        assert_eq!(
            super::table_formula_reference(view.rows(), (4, 2), Some((5, 3))),
            "C4:D6"
        );
    }

    #[test]
    fn mapped_batch_skips_hidden_records_recalculates_and_undoes_atomically() {
        let before = fixture(true);
        let view = before
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        let rows = view.visible_body_rows(4, 3).unwrap();
        assert_eq!(rows, vec![5, 3, 6]);
        let writes: Vec<_> = rows
            .iter()
            .zip([60, 70, 80])
            .map(|(&row, n)| TableCellWrite::value(row, 2, n.to_string()))
            .collect();
        let after = prepare_table_writes(&before, 0, &writes).unwrap();
        assert_eq!(after.active_sheet().get_raw(4, 2), "10");
        assert_eq!(
            after.active_sheet().get_computed_value(0, 1),
            Value::Number(220.0)
        );
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Number(120.0)
        );
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        assert_eq!(commit.patches.len(), 3);
        let mut wb = after;
        commit.replay(&mut wb, true).unwrap();
        assert_eq!(wb.active_sheet().get_raw(5, 2), "20");
        assert_eq!(
            wb.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        commit.replay(&mut wb, false).unwrap();
        assert_eq!(wb.active_sheet().get_raw(5, 2), "60");
    }

    #[test]
    fn invalid_batch_and_late_spill_leave_source_unchanged() {
        let wb = fixture(true);
        for invalid in [
            TableCellWrite::value(4, 2, "99".into()),
            TableCellWrite::value(2, 2, "Header".into()),
            TableCellWrite::value(5, 4, "Adjacent".into()),
            TableCellWrite::value(30, 2, "Out of bounds".into()),
        ] {
            assert!(prepare_table_writes(
                &wb,
                0,
                &[TableCellWrite::value(3, 2, "99".into()), invalid]
            )
            .is_err());
            assert_eq!(wb.active_sheet().get_raw(3, 2), "30");
        }
        // A Table edit makes an array outside the Table grow into adjacent
        // body rows. That must reject the whole candidate after recalc.
        let mut wb = fixture(false);
        wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(C4-29)");
        assert!(prepare_table_writes(&wb, 0, &[TableCellWrite::value(3, 2, "35".into())]).is_err());
        assert_eq!(wb.active_sheet().get_raw(3, 2), "30");
        let view = wb
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.visible_body_rows(6, 2).is_err());
    }

    #[test]
    fn edited_key_moves_or_disappears_without_changing_hidden_data() {
        let wb = fixture(true);
        let after =
            prepare_table_writes(&wb, 0, &[TableCellWrite::value(5, 1, "East".into())]).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.focus_record(5).unwrap().record_hidden);
        assert_eq!(super::focus_after_edit(view.rows(), view.range(), 4, 5), 5);
        // With no visible records, leave focus on the Table's header.
        let empty = prepare_table_writes(
            &wb,
            0,
            &[
                TableCellWrite::value(3, 1, "East".into()),
                TableCellWrite::value(5, 1, "East".into()),
                TableCellWrite::value(6, 1, "East".into()),
            ],
        )
        .unwrap();
        let empty_view = empty
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(
            super::focus_after_edit(empty_view.rows(), empty_view.range(), 4, 5),
            2
        );
        assert_eq!(after.active_sheet().get_raw(4, 1), "East");
        let after =
            prepare_table_writes(&wb, 0, &[TableCellWrite::value(5, 2, "90".into())]).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert_eq!(view.visible_body_rows(4, 3).unwrap(), vec![3, 6, 5]);
    }

    #[test]
    fn calculated_column_edit_is_one_record_override_and_text_stays_literal() {
        let mut wb = fixture(true);
        let id = wb.active_sheet().tables()[0].id;
        wb.set_calculated_column(id, 3, 3, "=C4*2", true).unwrap();
        let mut write = TableCellWrite::value(5, 3, "=1+1".into());
        write.literal_text = true;
        let after = prepare_table_writes(&wb, 0, &[write]).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Text("=1+1".into())
        );
        assert_eq!(after.active_sheet().get_raw(4, 3), "=C5*2");
        assert!(after.active_sheet().is_calculated_exception(5, 3));
        assert!(!after.active_sheet().is_calculated_exception(4, 3));
    }
    #[test]
    fn sparse_history_guards_whole_batch_and_preserves_unrelated_cells() {
        let before = fixture(true);
        let writes = vec![
            TableCellWrite::value(3, 2, "70".into()),
            TableCellWrite::value(6, 2, "80".into()),
        ];
        let after = prepare_table_writes(&before, 0, &writes).unwrap();
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            [(3, 2), (6, 2), (3, 2)],
        );
        assert_eq!(commit.patches.len(), 2);
        let mut wb = after.clone();
        wb.set_cell_value_tracked(0, 6, 2, "999");
        let rev = wb.revision();
        assert!(commit.replay(&mut wb, true).is_err());
        assert_eq!(wb.revision(), rev);
        assert_eq!(wb.active_sheet().get_raw(3, 2), "70");
        wb = after;
        wb.set_cell_value_tracked(0, 20, 7, "unrelated");
        commit.replay(&mut wb, true).unwrap();
        assert_eq!(wb.active_sheet().get_raw(20, 7), "unrelated");
        assert_eq!(
            wb.active_sheet().get_computed_value(0, 1),
            Value::Number(100.0)
        );
        commit.replay(&mut wb, false).unwrap();
        assert_eq!(
            wb.active_sheet().get_computed_value(0, 1),
            Value::Number(180.0)
        );
        wb.set_table_view_spec(SheetId(7), None).unwrap();
        let rev = wb.revision();
        assert!(commit.replay(&mut wb, true).is_err());
        assert_eq!(wb.revision(), rev);
    }

    #[test]
    fn sparse_history_restores_hidden_record_and_literal_cell_image() {
        let mut before = fixture(true);
        before.set_cell_text_tracked(0, 5, 3, "=1+1");
        let mut format = before.active_sheet().get_format(5, 3).clone();
        format.bold = true;
        before
            .sheet_mut(0)
            .unwrap()
            .set_format(5, 3, format.clone());
        let writes = vec![
            TableCellWrite::value(5, 1, "East".into()),
            TableCellWrite::value(5, 3, "=C6+1".into()),
        ];
        let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            [(5, 1), (5, 3)],
        );
        assert!(
            after
                .active_sheet()
                .build_saved_table_view(30)
                .unwrap()
                .unwrap()
                .focus_record(5)
                .unwrap()
                .record_hidden
        );
        commit.replay(&mut after, true).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Text("=1+1".into())
        );
        assert_eq!(after.active_sheet().get_format(5, 3), format);
        commit.replay(&mut after, false).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Number(21.0)
        );
        let noop =
            TableCellsCommit::capture(before.active_sheet(), before.active_sheet(), [(5, 1)]);
        assert!(noop.patches.is_empty());
    }

    #[test]
    fn sparse_history_postflight_failure_is_atomic() {
        let before = fixture(false);
        let mut after =
            prepare_table_writes(&before, 0, &[TableCellWrite::value(3, 2, "29".into())]).unwrap();
        let commit =
            TableCellsCommit::capture(before.active_sheet(), after.active_sheet(), [(3, 2)]);
        // Safe now; undo would grow this spill into adjacent Table body rows.
        after.set_cell_value_tracked(0, 0, 0, "=SEQUENCE((C4-29)*5+1)");
        let rev = after.revision();
        assert!(commit.replay(&mut after, true).is_err());
        assert_eq!(after.revision(), rev);
        assert_eq!(after.active_sheet().get_raw(3, 2), "29");
    }
}
