//! Edits outside a projected Table must remain atomic across recalculation.
use crate::{
    clipboard::{table_paste_writes, TablePasteKind},
    history::{History, UndoAction},
    table_cell_history::TableCellsCommit,
    table_edit::{prepare_table_writes, tests::fixture, view_safe_paste_targets, TableCellWrite},
};
use visigrid_engine::{
    cell::{CellComment, CellFormat},
    filter::RowView,
    sheet::{MergedRegion, Sheet, SheetId},
    workbook::Workbook,
};

fn with_controls() -> Workbook {
    let mut wb = fixture(true);
    assert!(wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Controls")));
    wb.set_cell_value_tracked(1, 0, 0, "30");
    wb.set_cell_value_tracked(0, 3, 2, "=Controls!A1");
    wb.set_cell_value_tracked(0, 3, 1, "=IF(Controls!A1>0,\"West\",\"East\")");
    wb
}

fn records(wb: &Workbook) -> Vec<usize> {
    let sheet = wb.sheet(0).unwrap();
    let view = sheet.build_saved_table_view(30).unwrap().unwrap();
    view.rows()
        .visible_rows()
        .iter()
        .copied()
        .filter(|r| *r > 2 && *r <= 6)
        .map(|r| view.rows().view_to_data(r))
        .collect()
}

fn commit(
    before: &Workbook,
    after: &Workbook,
    index: usize,
    writes: &[TableCellWrite],
) -> TableCellsCommit {
    TableCellsCommit::capture(
        before.sheet(index).unwrap(),
        after.sheet(index).unwrap(),
        writes.iter().map(|w| (w.row, w.col)),
    )
}

#[test]
fn titles_notes_totals_and_formats_are_undoable_without_changing_records() {
    let before = fixture(true);
    let mut note = TableCellWrite::value(7, 1, "Notes below the Table".into());
    note.format = Some(CellFormat {
        bold: true,
        ..Default::default()
    });
    note.comment = Some(Some(CellComment {
        text: "Review".into(),
        author: "Tester".into(),
    }));
    let writes = vec![
        TableCellWrite::value(0, 0, "Sales report".into()),
        note,
        TableCellWrite::value(8, 2, "=SUM(Sales[Amount])".into()),
    ];
    let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_display(8, 2), "100");
    assert_eq!(records(&after), records(&before));
    assert_eq!(after.active_sheet().get_raw(4, 1), "East");
    assert_eq!(
        after.active_sheet().tables(),
        before.active_sheet().tables()
    );
    let history = commit(&before, &after, 0, &writes);
    assert_eq!(history.patches.len(), 3);
    history.replay(&mut after, true).unwrap();
    assert!(after.active_sheet().get_cell_opt(7, 1).is_none());
    history.replay(&mut after, false).unwrap();
    assert!(after.active_sheet().get_format(7, 1).bold);
    assert_eq!(after.active_sheet().comment(7, 1).unwrap().text, "Review");
    assert_eq!(records(&after), records(&before));
}

#[test]
fn another_sheet_can_change_filter_membership_and_sort_order_with_sparse_history() {
    let before = with_controls();
    assert_eq!(records(&before), vec![5, 3, 6]);
    let writes = vec![TableCellWrite::value(0, 0, "5".into())];
    let mut after = prepare_table_writes(&before, 1, &writes).unwrap();
    assert_eq!(records(&after), vec![3, 5, 6]);
    assert_eq!(after.sheet(0).unwrap().get_display(0, 1), "75");
    let history = commit(&before, &after, 1, &writes);
    assert_eq!(history.patches.len(), 1);
    history.replay(&mut after, true).unwrap();
    assert_eq!(records(&after), vec![5, 3, 6]);
    history.replay(&mut after, false).unwrap();
    assert_eq!(records(&after), vec![3, 5, 6]);
    let hidden =
        prepare_table_writes(&after, 1, &[TableCellWrite::value(0, 0, "-1".into())]).unwrap();
    assert_eq!(records(&hidden), vec![5, 6]);
    assert_eq!(
        hidden.sheet(0).unwrap().table_view_spec(),
        before.sheet(0).unwrap().table_view_spec()
    );
}

#[test]
fn external_precedent_that_would_break_a_saved_layout_is_not_published() {
    let mut before = with_controls();
    before.set_cell_value_tracked(1, 0, 0, "1");
    before.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(Controls!A1)");
    assert!(before.sheet(0).unwrap().build_saved_table_view(30).is_ok());
    let revision = before.revision();
    assert!(prepare_table_writes(
        &before,
        1,
        &[
            TableCellWrite::value(0, 0, "5".into()),
            TableCellWrite::value(2, 0, "must not apply".into())
        ]
    )
    .is_err());
    assert_eq!(before.revision(), revision);
    assert_eq!(before.sheet(1).unwrap().get_raw(0, 0), "1");
    assert_eq!(before.sheet(1).unwrap().get_raw(2, 0), "");
    assert_eq!(before.sheet(0).unwrap().get_display(0, 0), "1");
}

#[test]
fn undo_of_an_outside_edit_also_rejects_a_new_unsafe_dependency() {
    let mut before = with_controls();
    before.set_cell_value_tracked(1, 0, 0, "5");
    let writes = vec![TableCellWrite::value(0, 0, "1".into())];
    let mut after = prepare_table_writes(&before, 1, &writes).unwrap();
    let history = commit(&before, &after, 1, &writes);
    after.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(Controls!A1)");
    let revision = after.revision();
    assert!(history.replay(&mut after, true).is_err());
    assert_eq!(after.revision(), revision);
    assert_eq!(after.sheet(1).unwrap().get_raw(0, 0), "1");
    after.set_cell_value_tracked(0, 0, 0, "");
    history.replay(&mut after, true).unwrap();
    assert_eq!(after.sheet(1).unwrap().get_raw(0, 0), "5");
}

#[test]
fn every_target_is_checked_before_a_mixed_clear_or_paste() {
    let before = fixture(true);
    for bad in [(3, 0), (4, 2), (2, 1), (30, 0), (0, 8)] {
        let writes = vec![
            TableCellWrite::value(0, 0, "allowed".into()),
            TableCellWrite::value(bad.0, bad.1, "".into()),
        ];
        assert!(
            prepare_table_writes(&before, 0, &writes).is_err(),
            "{bad:?}"
        );
        assert_eq!(before.active_sheet().get_raw(0, 0), "");
    }
    // Ordinary clearing on an unrelated sheet needs no Table definition.
    let before = with_controls();
    let writes = vec![TableCellWrite::value(0, 0, String::new())];
    let mut after = prepare_table_writes(&before, 1, &writes).unwrap();
    let history = commit(&before, &after, 1, &writes);
    assert_eq!(records(&after), vec![5, 6]);
    history.replay(&mut after, true).unwrap();
    assert_eq!(records(&after), records(&before));
}

#[test]
fn ordinary_paste_uses_visible_canonical_rows_and_body_paste_stays_bounded() {
    let before = fixture(true);
    let view = before
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    let grid = vec![
        vec!["First".into(), "001".into()],
        vec!["Second".into(), "=SUM(Sales[Amount])".into()],
    ];
    let targets =
        view_safe_paste_targets(before.active_sheet(), view.rows(), (8, 1), 2, 2).unwrap();
    assert_eq!(
        targets,
        vec![(8, 1, 0, 0), (8, 2, 0, 1), (9, 1, 1, 0), (9, 2, 1, 1)]
    );
    let writes = table_paste_writes(&grid, None, TablePasteKind::Contents, targets);
    let after = prepare_table_writes(&before, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_display(9, 2), "100");
    assert_eq!(records(&after), records(&before));
    assert!(view_safe_paste_targets(before.active_sheet(), view.rows(), (4, 1), 4, 1).is_err());
    assert!(view_safe_paste_targets(before.active_sheet(), view.rows(), (29, 1), 2, 1).is_err());
    let targets =
        view_safe_paste_targets(before.active_sheet(), view.rows(), (1, 1), 2, 1).unwrap();
    let writes = table_paste_writes(&grid, None, TablePasteKind::Contents, targets);
    assert!(prepare_table_writes(&before, 0, &writes).is_err());
    // Another sheet can still have an ordinary worksheet sort/filter projection.
    let other = with_controls();
    let mut rows = RowView::new(30);
    let mut order: Vec<_> = (0..30).collect();
    order.swap(8, 10);
    let mut visible = vec![true; 30];
    visible[9] = false;
    rows.restore(order, visible);
    assert_eq!(
        view_safe_paste_targets(other.sheet(1).unwrap(), &rows, (8, 1), 2, 1).unwrap(),
        vec![(10, 1, 0, 0), (8, 1, 1, 0)]
    );
}

#[test]
fn merged_origins_are_editable_but_receivers_and_pivot_output_are_not() {
    let mut before = with_controls();
    before
        .sheet_mut(1)
        .unwrap()
        .add_merge(MergedRegion {
            start: (8, 0),
            end: (8, 2),
        })
        .unwrap();
    assert!(prepare_table_writes(
        &before,
        1,
        &[TableCellWrite::value(8, 0, "Merged title".into())]
    )
    .is_ok());
    assert!(
        prepare_table_writes(&before, 1, &[TableCellWrite::value(8, 1, "covered".into())]).is_err()
    );
    before.set_cell_value_tracked(1, 10, 0, "=SEQUENCE(2)");
    assert!(prepare_table_writes(
        &before,
        1,
        &[TableCellWrite::value(11, 0, "receiver".into())]
    )
    .is_err());
    let id = before.sheet(0).unwrap().tables()[0].id;
    let source = before.table_pivot_source(id).unwrap();
    use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotField, PivotValueField};
    let field = &before.table(id).unwrap().1.columns[1];
    let definition = PivotDefinition {
        rows: vec![],
        column: None,
        values: vec![PivotValueField {
            field: PivotField {
                column_id: Some(field.id),
                offset: 1,
                header: field.name.clone(),
            },
            aggregation: Aggregation::Sum,
            number_format: None,
        }],
    };
    let (pivot, index) = before.create_pivot(source, definition).unwrap();
    let region = before
        .sheet(index)
        .unwrap()
        .pivots
        .iter()
        .find(|p| p.id == pivot)
        .unwrap()
        .region()
        .unwrap();
    assert!(prepare_table_writes(
        &before,
        index,
        &[TableCellWrite::value(region.0, region.1, "output".into())]
    )
    .is_err());
}

#[test]
fn stale_outside_history_cannot_overwrite_a_changed_target() {
    let before = with_controls();
    let writes = vec![TableCellWrite::value(0, 0, "5".into())];
    let mut after = prepare_table_writes(&before, 1, &writes).unwrap();
    let history = commit(&before, &after, 1, &writes);
    after.set_cell_value_tracked(1, 0, 0, "1000");
    let revision = after.revision();
    assert!(history.replay(&mut after, true).is_err());
    assert_eq!(after.revision(), revision);
    assert_eq!(after.sheet(1).unwrap().get_raw(0, 0), "1000");
}

#[test]
fn rewind_rebuilds_all_sheet_views_after_an_outside_edit() {
    let mut base = with_controls();
    base.set_active_sheet(1);
    let writes = vec![TableCellWrite::value(0, 0, "-1".into())];
    let after = prepare_table_writes(&base, 1, &writes).unwrap();
    let mut history = History::new();
    history.record_action_with_provenance(
        UndoAction::TableCellsChanged {
            sheet_index: 1,
            commit: Box::new(commit(&base, &after, 1, &writes)),
            description: "Edit control".into(),
        },
        None,
    );
    let result = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap();
    assert_eq!(records(&result.workbook), vec![5, 6]);
    assert!(result.view_state.per_sheet[0].table_rows.is_some());
    assert!(result.view_state.per_sheet[1].table_rows.is_none());
    let before = history
        .build_workbook_before(0, Some(&base), 100, 10_000)
        .unwrap();
    assert_eq!(records(&before.workbook), records(&base));
}

#[test]
fn a_new_spill_cannot_silently_consume_another_paste_target() {
    let before = with_controls();
    let revision = before.revision();
    let writes = vec![
        TableCellWrite::value(8, 0, "=SEQUENCE(2)".into()),
        TableCellWrite::value(9, 0, String::new()),
    ];
    assert!(prepare_table_writes(&before, 1, &writes).is_err());
    assert_eq!(before.revision(), revision);
    assert!(before.sheet(1).unwrap().get_cell_opt(8, 0).is_none());
}

#[test]
fn replay_compares_schema_by_identity_after_an_intervening_rename_is_undone() {
    use visigrid_engine::table::TableRange;
    let mut before = fixture(true);
    let id = before.active_sheet().tables()[0].id;
    before
        .create_table(
            SheetId(7),
            TableRange {
                start_row: 15,
                end_row: 16,
                start_col: 1,
                end_col: 2,
            },
            "OtherTable",
        )
        .unwrap();
    let writes = vec![TableCellWrite::value(10, 0, "note".into())];
    let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
    let history = commit(&before, &after, 0, &writes);
    let rename = after.rename_table(id, "Renamed").unwrap();
    after.apply_table_commit(&rename, true).unwrap();
    history.replay(&mut after, true).unwrap();
    assert_eq!(after.active_sheet().get_raw(10, 0), "");
    history.replay(&mut after, false).unwrap();
    assert_eq!(after.active_sheet().get_raw(10, 0), "note");
}
