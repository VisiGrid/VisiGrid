//! Table creation outside active projections, with atomic header insertion.
use crate::{app::Spreadsheet, history::UndoAction, table_structure::StructureLayout};
use gpui::Context;
use visigrid_engine::{
    filter::RowView,
    sheet::SheetId,
    structural::Axis,
    table::TableRange,
    workbook::{StructureStep, TableCommit, Workbook},
};

#[derive(Clone, Debug, serde::Serialize)]
pub(crate) struct HeaderLayout {
    pub before: StructureLayout,
    pub after: StructureLayout,
    pub frozen_before: (usize, usize),
    pub frozen_after: (usize, usize),
}

pub(crate) fn is_creation(commit: &TableCommit) -> bool {
    commit.before_table().is_none() && commit.after_table().is_some()
}

fn validate_views(wb: &Workbook) -> Result<(), String> {
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(())
}

/// Outside an active Table's body, view slots are worksheet rows. Reject a
/// crossing selection instead of turning sorted endpoints into a new rectangle.
pub(crate) fn creation_selection(
    wb: &Workbook,
    sheet: SheetId,
    range: TableRange,
) -> Result<TableRange, String> {
    let sheet = wb.sheet_by_id(sheet).ok_or("The sheet no longer exists.")?;
    range.validate(sheet.rows, sheet.cols)?;
    if let Some(table) = sheet
        .table_view_spec()
        .filter(|s| s.has_criteria())
        .and_then(|spec| sheet.tables().iter().find(|t| t.id == spec.table))
    {
        if table.range.data_rows() > 0
            && range.start_row <= table.range.end_row
            && range.end_row > table.range.start_row
        {
            return Err("Select cells above or below the sorted/filtered Table, or on another sheet, to create a new Table.".into());
        }
    }
    Ok(range)
}

fn finish_candidate(mut wb: Workbook) -> Result<Workbook, String> {
    if let Some(error) = wb.take_incremental_errors().first() {
        return Err(format!("Could not recalculate the new Table: {error:?}"));
    }
    validate_views(&wb)?;
    Ok(wb)
}

fn prepare_creation(
    wb: &Workbook,
    sheet_id: SheetId,
    range: TableRange,
    name: &str,
    has_headers: bool,
    layout: &StructureLayout,
) -> Result<(Workbook, TableCommit, Option<Box<HeaderLayout>>), String> {
    wb.ensure_writable()?;
    validate_views(wb)?;
    creation_selection(wb, sheet_id, range)?;
    let sheet = wb
        .sheet_by_id(sheet_id)
        .ok_or("The sheet no longer exists.")?;
    let header_layout = if has_headers {
        None
    } else {
        let after = layout.shifted(
            sheet,
            &[StructureStep {
                axis: Axis::Row,
                at: range.start_row,
                count: 1,
                delete: false,
            }],
        )?;
        let frozen_before = sheet.frozen_panes;
        let frozen_after = (
            if range.start_row < frozen_before.0 {
                (frozen_before.0 + 1).min(sheet.rows)
            } else {
                frozen_before.0
            },
            frozen_before.1,
        );
        Some(Box::new(HeaderLayout {
            before: layout.clone(),
            after,
            frozen_before,
            frozen_after,
        }))
    };
    let mut candidate = wb.clone();
    let commit = if has_headers {
        candidate.create_table(sheet_id, range, name)?
    } else {
        candidate.create_table_without_headers(sheet_id, range, name)?
    };
    if let Some(layout) = &header_layout {
        candidate.sheet_by_id_mut(sheet_id).unwrap().frozen_panes = layout.frozen_after;
    }
    Ok((finish_candidate(candidate)?, commit, header_layout))
}

pub(crate) fn prepare_creation_replay(
    wb: &Workbook,
    commit: &TableCommit,
    layout: Option<&HeaderLayout>,
    undo: bool,
) -> Result<Workbook, String> {
    wb.ensure_writable()?;
    validate_views(wb)?;
    if !is_creation(commit) {
        return Err("This history entry is not Table creation.".into());
    }
    let mut candidate = wb.clone();
    if let Some(layout) = layout {
        let expected = if undo {
            layout.frozen_after
        } else {
            layout.frozen_before
        };
        if wb
            .sheet_by_id(commit.sheet_id())
            .ok_or("The history sheet no longer exists.")?
            .frozen_panes
            != expected
        {
            return Err(
                "Frozen panes changed since header insertion. Undo those changes first.".into(),
            );
        }
    }
    candidate.apply_table_commit(commit, undo)?;
    if let Some(layout) = layout {
        candidate
            .sheet_by_id_mut(commit.sheet_id())
            .unwrap()
            .frozen_panes = if undo {
            layout.frozen_before
        } else {
            layout.frozen_after
        };
    }
    finish_candidate(candidate)
}

impl Spreadsheet {
    pub(crate) fn submit_table_creation(
        &mut self,
        sheet_id: SheetId,
        range: TableRange,
        name: &str,
        has_headers: bool,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.sync_table_view(cx);
        if sheet_id != self.sheet(cx).id {
            return Err("Return to the source sheet before creating this Table.".into());
        }
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            return Err(
                "Clear this sheet's worksheet sorting and filters before creating a Table.".into(),
            );
        }
        let before = self.structure_layout(sheet_id);
        let mut base = self.wb(cx).clone();
        base.sheet_by_id_mut(sheet_id).unwrap().frozen_panes =
            (self.view_state.frozen_rows, self.view_state.frozen_cols);
        self.validate_structure_layout(&base, sheet_id, &before)?;
        let (candidate, commit, header_layout) =
            prepare_creation(&base, sheet_id, range, name, has_headers, &before)?;
        let layout = header_layout.as_ref().map(|l| &l.after).unwrap_or(&before);
        self.validate_structure_layout(&candidate, sheet_id, layout)?;
        let index = candidate.sheet_index_by_id(sheet_id).unwrap();
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        if let Some(layout) = &header_layout {
            self.install_structure_layout(sheet_id, &layout.after);
        }
        self.finish_creation_view(&commit, cx);
        self.history.record_action_with_provenance(
            UndoAction::TableCommit {
                sheet_index: index,
                commit: Box::new(commit),
                header_layout,
                description: format!("Create Table: {name}"),
            },
            None,
        );
        self.is_modified = true;
        self.bump_cells_rev();
        self.status_message = Some(format!("Create Table: {name}"));
        Ok(())
    }

    fn finish_creation_view(&mut self, commit: &TableCommit, cx: &mut Context<Self>) {
        let index = self.wb(cx).sheet_index_by_id(commit.sheet_id()).unwrap();
        if self.sheet_index(cx) != index {
            self.activate_sheet(index, cx);
        }
        self.view_state.frozen_rows = self.sheet(cx).frozen_panes.0;
        self.view_state.frozen_cols = self.sheet(cx).frozen_panes.1;
        self.table_filter_dropdown = None;
        self.sync_table_view(cx);
        if !self.table_view_installed {
            self.row_view = RowView::new(crate::app::NUM_ROWS.min(self.sheet(cx).rows));
        }
        let range = commit.after_table().unwrap().range;
        self.view_state
            .select_cell(range.start_row, range.start_col);
        self.view_state.additional_selections.clear();
        self.ensure_visible(cx);
    }

    pub(crate) fn replay_table_creation(
        &mut self,
        commit: &TableCommit,
        header_layout: Option<&HeaderLayout>,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = (|| {
            let current = self.structure_layout(commit.sheet_id());
            if let Some(layout) = header_layout {
                let expected = if undo { &layout.after } else { &layout.before };
                if &current != expected {
                    return Err("Row or column layout changed since header insertion. Undo those changes first.".into());
                }
            }
            let candidate = prepare_creation_replay(self.wb(cx), commit, header_layout, undo)?;
            let target = header_layout
                .map(|l| if undo { &l.before } else { &l.after })
                .unwrap_or(&current);
            self.validate_structure_layout(&candidate, commit.sheet_id(), target)?;
            Ok::<_, String>((candidate, target.clone()))
        })();
        match result {
            Ok((candidate, layout)) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.install_structure_layout(commit.sheet_id(), &layout);
                self.finish_creation_view(commit, cx);
                self.bump_cells_rev();
                self.is_modified = true;
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
    use visigrid_engine::{cell::CellComment, table_view::TableViewSpec};

    fn range(r0: usize, c0: usize, r1: usize, c1: usize) -> TableRange {
        TableRange {
            start_row: r0,
            start_col: c0,
            end_row: r1,
            end_col: c1,
        }
    }
    fn below() -> TableRange {
        range(10, 1, 12, 2)
    }
    fn source() -> Workbook {
        let mut wb = fixture(true);
        for (row, col, value) in [
            (10, 1, "Item"),
            (10, 2, "Cost"),
            (11, 1, "Tea"),
            (11, 2, "12"),
            (12, 1, "Coffee"),
            (12, 2, "25"),
        ] {
            wb.set_cell_value_tracked(0, row, col, value);
        }
        wb
    }
    fn create(wb: &Workbook, headers: bool) -> (Workbook, TableCommit, Option<Box<HeaderLayout>>) {
        prepare_creation(
            wb,
            wb.active_sheet_id(),
            below(),
            "Stock",
            headers,
            &StructureLayout::default(),
        )
        .unwrap()
    }
    fn action(commit: TableCommit, header_layout: Option<Box<HeaderLayout>>) -> UndoAction {
        UndoAction::TableCommit {
            sheet_index: 0,
            commit: Box::new(commit),
            header_layout,
            description: "Create Table: Stock".into(),
        }
    }

    #[test]
    fn creates_below_filtered_table_and_preserves_hidden_records_and_criteria() {
        let before = source();
        let spec = before.active_sheet().table_view_spec().cloned();
        let (after, commit, layout) = create(&before, true);
        assert!(is_creation(&commit));
        assert!(layout.is_none());
        assert_eq!(after.active_sheet().tables().len(), 2);
        assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
        assert_eq!(after.active_sheet().get_raw(4, 1), "East");
        assert_eq!(after.active_sheet().get_raw(12, 2), "25");
        let undone = prepare_creation_replay(&after, &commit, None, true).unwrap();
        assert_eq!(
            undone.active_sheet().tables(),
            before.active_sheet().tables()
        );
        assert_eq!(undone.active_sheet().get_raw(10, 1), "Item");
        let redone = prepare_creation_replay(&undone, &commit, None, false).unwrap();
        assert_eq!(
            redone.active_sheet().tables(),
            after.active_sheet().tables()
        );
    }

    #[test]
    fn creates_on_other_sheet_and_binds_existing_structured_formulas() {
        let mut before = fixture(true);
        let other = before.add_sheet_named("New data").unwrap();
        let sheet_id = before.sheet(other).unwrap().id;
        before.set_cell_value_tracked(other, 0, 0, "Price");
        before.set_cell_value_tracked(other, 1, 0, "17");
        before.set_cell_value_tracked(0, 0, 4, "=SUM(Stock[Price])");
        let (after, commit, _) = prepare_creation(
            &before,
            sheet_id,
            range(0, 0, 1, 0),
            "Stock",
            true,
            &StructureLayout::default(),
        )
        .unwrap();
        assert_eq!(after.active_sheet().get_display(0, 4), "17");
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        let undone = prepare_creation_replay(&after, &commit, None, true).unwrap();
        assert_eq!(
            undone.active_sheet().get_raw(0, 4),
            before.active_sheet().get_raw(0, 4)
        );
        assert_eq!(
            undone.active_sheet().get_display(0, 4),
            before.active_sheet().get_display(0, 4)
        );
    }

    #[test]
    fn selection_crossing_sorted_or_hidden_rows_refuses_instead_of_mapping_endpoints() {
        let before = fixture(true);
        for r in [range(1, 5, 5, 6), range(3, 5, 4, 6), range(4, 5, 8, 6)] {
            assert!(creation_selection(&before, before.active_sheet_id(), r).is_err());
            assert!(prepare_creation(
                &before,
                before.active_sheet_id(),
                r,
                "Stock",
                true,
                &StructureLayout::default()
            )
            .is_err());
        }
        assert_eq!(
            creation_selection(&before, before.active_sheet_id(), below()).unwrap(),
            below()
        );
        assert!(creation_selection(&before, before.active_sheet_id(), range(0, 0, 1, 0)).is_ok());
    }

    #[test]
    fn headerless_insertion_above_filtered_table_moves_records_layout_and_freeze_boundary() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 0, 4, "10");
        before.set_cell_value_tracked(0, 0, 5, "20");
        before.active_sheet_mut().frozen_panes = (3, 1);
        let mut layout = StructureLayout::default();
        layout.heights.insert(8, 35.0);
        layout.hidden_rows.insert(9);
        layout.widths.insert(5, 140.0);
        let (after, commit, header_layout) = prepare_creation(
            &before,
            before.active_sheet_id(),
            range(0, 4, 0, 5),
            "Stock",
            false,
            &layout,
        )
        .unwrap();
        let saved = header_layout.as_ref().unwrap();
        assert_eq!(saved.after.heights.get(&9), Some(&35.0));
        assert!(saved.after.hidden_rows.contains(&10));
        assert_eq!(saved.after.widths, layout.widths);
        assert_eq!(after.active_sheet().frozen_panes, (4, 1));
        assert_eq!(after.active_sheet().tables()[0].range.start_row, 3);
        assert_eq!(after.active_sheet().get_raw(5, 1), "East");
        assert_eq!(after.active_sheet().get_raw(1, 4), "10");
        assert_eq!(after.active_sheet().get_raw(0, 4), "Column1");
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        let undone =
            prepare_creation_replay(&after, &commit, header_layout.as_deref(), true).unwrap();
        assert_eq!(
            undone.active_sheet().tables(),
            before.active_sheet().tables()
        );
        assert_eq!(undone.active_sheet().frozen_panes, (3, 1));
        assert_eq!(
            undone.active_sheet().get_raw(0, 1),
            before.active_sheet().get_raw(0, 1)
        );
        let redone =
            prepare_creation_replay(&undone, &commit, header_layout.as_deref(), false).unwrap();
        assert_eq!(
            redone.active_sheet().tables(),
            after.active_sheet().tables()
        );
    }

    #[test]
    fn headerless_creation_preserves_all_source_rows_and_neighboring_comments() {
        let mut before = source();
        before.set_cell_value_tracked(0, 12, 5, "neighbor");
        before.active_sheet_mut().set_comment(
            12,
            5,
            Some(CellComment {
                text: "keep".into(),
                author: "me".into(),
            }),
        );
        let (after, commit, layout) = create(&before, false);
        assert_eq!(commit.after_table().unwrap().range, range(10, 1, 13, 2));
        assert_eq!(after.active_sheet().get_raw(11, 1), "Item");
        assert_eq!(after.active_sheet().get_raw(13, 5), "neighbor");
        assert_eq!(after.active_sheet().comment(13, 5).unwrap().text, "keep");
        let undone = prepare_creation_replay(&after, &commit, layout.as_deref(), true).unwrap();
        assert_eq!(undone.active_sheet().get_raw(12, 5), "neighbor");
        assert_eq!(
            undone.active_sheet().comment(12, 5),
            before.active_sheet().comment(12, 5)
        );
    }

    #[test]
    fn refused_creation_does_not_consume_ids_or_change_headers() {
        let before = source();
        let rev = before.revision();
        for (r, name) in [
            (below(), "Sales"),
            (range(2, 1, 4, 2), "Stock"),
            (range(29, 1, 30, 2), "Stock"),
        ] {
            assert!(prepare_creation(
                &before,
                before.active_sheet_id(),
                r,
                name,
                true,
                &StructureLayout::default()
            )
            .is_err());
        }
        assert_eq!(before.revision(), rev);
        assert_eq!(before.active_sheet().get_raw(10, 1), "Item");
        let (_, commit, _) = create(&before, true);
        let (_, again, _) = create(&before, true);
        assert_eq!(commit.table_id(), again.table_id());
    }

    #[test]
    fn headerless_creation_refuses_layout_at_grid_edge() {
        let before = source();
        let mut layout = StructureLayout::default();
        layout.heights.insert(29, 30.0);
        assert!(prepare_creation(
            &before,
            before.active_sheet_id(),
            below(),
            "Stock",
            false,
            &layout
        )
        .is_err());
        layout.heights.clear();
        layout.hidden_rows.insert(29);
        assert!(prepare_creation(
            &before,
            before.active_sheet_id(),
            below(),
            "Stock",
            false,
            &layout
        )
        .is_err());
    }

    #[test]
    fn binding_new_table_cannot_invalidate_an_existing_view_via_spill() {
        let mut before = source();
        before.set_cell_value_tracked(0, 0, 0, "=IF(IFERROR(SUM(Stock[Cost]),0)>0,SEQUENCE(5),0)");
        validate_views(&before).unwrap();
        let rev = before.revision();
        assert!(prepare_creation(
            &before,
            before.active_sheet_id(),
            below(),
            "Stock",
            true,
            &StructureLayout::default()
        )
        .is_err());
        assert_eq!(before.revision(), rev);
        assert_eq!(before.active_sheet().get_display(0, 0), "0");
    }

    #[test]
    fn replay_refuses_stale_headers_new_criteria_and_unsafe_dependents() {
        let before = source();
        let (mut after, commit, _) = create(&before, true);
        after.set_cell_value_tracked(0, 0, 0, "=IF(IFERROR(SUM(Stock[Cost]),0)=0,SEQUENCE(5),0)");
        assert!(prepare_creation_replay(&after, &commit, None, true).is_err());
        after.set_cell_value_tracked(0, 0, 0, "");
        after.rename_table(commit.table_id(), "Renamed").unwrap();
        assert!(prepare_creation_replay(&after, &commit, None, true).is_err());
        let (mut after, commit, _) = create(&before, true);
        // Even a buttons-only saved view on the newly created Table must be cleared before removal.
        after
            .set_table_view_spec(before.active_sheet_id(), None)
            .unwrap();
        let mut spec = TableViewSpec::new(commit.table_id());
        spec.show_filter_buttons = false;
        after
            .set_table_view_spec(before.active_sheet_id(), Some(spec))
            .unwrap();
        assert!(prepare_creation_replay(&after, &commit, None, true).is_err());
    }

    #[test]
    fn headerless_replay_checks_inserted_row_and_frozen_panes() {
        let before = source();
        let (mut after, commit, layout) = create(&before, false);
        after.active_sheet_mut().frozen_panes = (1, 0);
        assert!(prepare_creation_replay(&after, &commit, layout.as_deref(), true).is_err());
        after.active_sheet_mut().frozen_panes = (0, 0);
        after.set_cell_value_tracked(0, 10, 5, "new header-row content");
        assert!(prepare_creation_replay(&after, &commit, layout.as_deref(), true).is_err());
    }

    #[test]
    fn rewind_restores_header_insertion_layout_before_and_after_creation() {
        let before = source();
        let mut layout = StructureLayout::default();
        layout.heights.insert(12, 35.0);
        layout.hidden_rows.insert(15);
        let (after, commit, header_layout) = prepare_creation(
            &before,
            before.active_sheet_id(),
            below(),
            "Stock",
            false,
            &layout,
        )
        .unwrap();
        let shifted = header_layout.as_ref().unwrap().after.clone();
        let mut history = History::new();
        history.record_action_with_provenance(action(commit, header_layout), None);
        assert_eq!(history.undo_count(), 1);
        let preview = history
            .build_workbook_before(1, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(
            preview.workbook.active_sheet().tables(),
            after.active_sheet().tables()
        );
        assert_eq!(
            preview.view_state.per_sheet[0].structure_layout,
            Some(shifted)
        );
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
        let original = history
            .build_workbook_before(0, Some(&before), 100, 10_000)
            .unwrap();
        assert_eq!(
            original.view_state.per_sheet[0].structure_layout,
            Some(layout)
        );
        assert_eq!(
            original.workbook.active_sheet().tables(),
            before.active_sheet().tables()
        );
    }

    #[test]
    fn native_roundtrip_retains_both_tables_and_original_criteria() {
        for headers in [true, false] {
            let before = source();
            let (after, _, _) = create(&before, headers);
            let dir = tempfile::tempdir().unwrap();
            let path = dir.path().join("creation.sheet");
            visigrid_io::native::save_workbook_full(&after, &Default::default(), &[], &[], &path)
                .unwrap();
            let loaded = visigrid_io::native::load_workbook(&path).unwrap();
            assert_eq!(
                loaded.active_sheet().tables(),
                after.active_sheet().tables()
            );
            assert_eq!(
                loaded.active_sheet().table_view_spec(),
                before.active_sheet().table_view_spec()
            );
        }
    }

    #[test]
    fn recovery_blocks_both_creation_modes() {
        let mut before = source();
        before.active_sheet_mut().read_only_reason = Some("Corrupt metadata".into());
        for headers in [true, false] {
            assert!(prepare_creation(
                &before,
                before.active_sheet_id(),
                below(),
                "Stock",
                headers,
                &StructureLayout::default()
            )
            .unwrap_err()
            .contains("Read-only"));
        }
    }
    #[test]
    fn rewind_across_freeze_then_header_insertion_restores_each_boundary() {
        let mut base = fixture(true);
        base.set_cell_value_tracked(0, 0, 4, "10");
        let mut before = base.clone();
        before.active_sheet_mut().frozen_panes = (3, 0);
        let (after, commit, layout) = prepare_creation(
            &before,
            before.active_sheet_id(),
            range(0, 4, 0, 4),
            "Stock",
            false,
            &StructureLayout::default(),
        )
        .unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::FreezePanesChanged {
                sheet_id: before.active_sheet_id(),
                old_frozen_rows: 0,
                old_frozen_cols: 0,
                new_frozen_rows: 3,
                new_frozen_cols: 0,
            },
            None,
        );
        history.record_action_with_provenance(action(commit, layout), None);
        for (at, frozen) in [(0, 0), (1, 3), (2, 4)] {
            let preview = history
                .build_workbook_before(at, Some(&base), 100, 10_000)
                .unwrap();
            assert_eq!(preview.workbook.active_sheet().frozen_panes, (frozen, 0));
            if at == 2 {
                assert_eq!(
                    preview.workbook.active_sheet().tables(),
                    after.active_sheet().tables()
                );
            }
        }
    }
    #[test]
    fn dormant_header_only_and_button_only_specs_do_not_block_adjacent_creation() {
        for empty_body in [true, false] {
            let mut before = fixture(true);
            let table = before.active_sheet().tables()[0].clone();
            if empty_body {
                before
                    .resize_table(
                        table.id,
                        TableRange {
                            end_row: table.range.start_row,
                            ..table.range
                        },
                    )
                    .unwrap();
            } else {
                let mut spec = TableViewSpec::new(table.id);
                spec.show_filter_buttons = false;
                before
                    .set_table_view_spec(before.active_sheet_id(), Some(spec))
                    .unwrap();
            }
            let spec = before.active_sheet().table_view_spec().cloned();
            let (after, _, _) = prepare_creation(
                &before,
                before.active_sheet_id(),
                range(2, 5, 4, 6),
                "Stock",
                true,
                &StructureLayout::default(),
            )
            .unwrap();
            assert_eq!(after.active_sheet().table_view_spec(), spec.as_ref());
            assert_eq!(after.active_sheet().tables().len(), 2);
        }
    }
}
