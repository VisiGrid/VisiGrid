//! Guarded cut/fill outside Tables, including other-sheet dependencies.
use crate::{
    clipboard::{table_paste_writes, TablePasteKind},
    history::{History, UndoAction},
    table_cell_history::TableCellsCommit,
    table_cut::plan_cut,
    table_edit::{prepare_table_writes, tests::fixture, view_safe_selection_rows, TableCellWrite},
    table_fill::{plan_broadcast, plan_direction, plan_handle},
};
use visigrid_engine::{
    cell::{CellComment, CellFormat},
    filter::RowView,
    sheet::{MergedRegion, Sheet, SheetId},
    workbook::Workbook,
};
fn controls() -> Workbook {
    let mut wb = fixture(true);
    assert!(wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Controls")));
    wb.set_cell_value_tracked(1, 0, 0, "30");
    wb.set_cell_value_tracked(0, 3, 2, "=Controls!A1");
    wb.set_cell_value_tracked(0, 3, 1, "=IF(Controls!A1>0,\"West\",\"East\")");
    wb
}
fn rows(wb: &Workbook, index: usize) -> RowView {
    let s = wb.sheet(index).unwrap();
    s.build_saved_table_view(s.rows)
        .unwrap()
        .map(|v| v.rows().clone())
        .unwrap_or_else(|| RowView::new(s.rows))
}
fn records(wb: &Workbook) -> Vec<usize> {
    let v = rows(wb, 0);
    v.visible_rows()
        .iter()
        .copied()
        .filter(|r| (3..=6).contains(r))
        .map(|r| v.view_to_data(r))
        .collect()
}
fn commit(b: &Workbook, a: &Workbook, i: usize, w: &[TableCellWrite]) -> TableCellsCommit {
    TableCellsCommit::capture(
        b.sheet(i).unwrap(),
        a.sheet(i).unwrap(),
        w.iter().map(|w| (w.row, w.col)),
    )
}

#[test]
fn notes_cut_paste_preserves_exact_text_metadata_and_separate_history() {
    let mut b = fixture(true);
    b.set_cell_text_tracked(0, 8, 1, "001\t=literal\nnotes");
    b.set_cell_value_tracked(0, 9, 1, "=SUM(Sales[Amount])+$C$4");
    b.active_sheet_mut().set_format(
        8,
        1,
        CellFormat {
            bold: true,
            ..Default::default()
        },
    );
    b.active_sheet_mut().set_comment(
        9,
        1,
        Some(CellComment {
            text: "Review".into(),
            author: "R".into(),
        }),
    );
    let (ic, w) = plan_cut(b.active_sheet(), &rows(&b, 0), ((8, 1), (9, 1))).unwrap();
    assert_eq!(ic.source_rows, vec![8, 9]);
    assert_eq!(ic.raw_cells[0][0], "001\t=literal\nnotes");
    let cut = prepare_table_writes(&b, 0, &w).unwrap();
    let cut_commit = commit(&b, &cut, 0, &w);
    assert!(cut.active_sheet().get_format(8, 1).bold);
    assert!(cut.active_sheet().comment(9, 1).is_none());
    assert_eq!(records(&cut), records(&b));
    let paste = table_paste_writes(
        &ic.raw_cells,
        Some(&ic),
        TablePasteKind::All,
        vec![(12, 2, 0, 0), (13, 2, 1, 0)],
    );
    let mut a = prepare_table_writes(&cut, 0, &paste).unwrap();
    assert_eq!(a.active_sheet().get_raw(12, 2), ic.raw_cells[0][0]);
    assert!(!a.active_sheet().get_cell(12, 2).value().is_formula());
    assert_eq!(a.active_sheet().get_display(13, 2), "130");
    assert!(a.active_sheet().get_format(12, 2).bold);
    assert_eq!(
        a.active_sheet().comment(13, 2),
        b.active_sheet().comment(9, 1)
    );
    commit(&cut, &a, 0, &paste).replay(&mut a, true).unwrap();
    assert_eq!(a.active_sheet().get_raw(8, 1), "");
    cut_commit.replay(&mut a, true).unwrap();
    assert_eq!(a.active_sheet().get_raw(8, 1), ic.raw_cells[0][0]);
    cut_commit.replay(&mut a, false).unwrap();
    assert_eq!(a.active_sheet().get_raw(9, 1), "");
}

#[test]
fn other_sheet_cut_refreshes_membership_and_replays_without_a_local_table() {
    let b = controls();
    let (ic, w) = plan_cut(b.sheet(1).unwrap(), &rows(&b, 1), ((0, 0), (0, 0))).unwrap();
    assert_eq!(ic.raw_tsv, "30");
    let mut a = prepare_table_writes(&b, 1, &w).unwrap();
    assert_eq!(records(&a), vec![5, 6]);
    assert_eq!(a.sheet(0).unwrap().get_display(0, 1), "70");
    let h = commit(&b, &a, 1, &w);
    h.replay(&mut a, true).unwrap();
    assert_eq!(records(&a), vec![5, 3, 6]);
    h.replay(&mut a, false).unwrap();
    assert_eq!(records(&a), vec![5, 6]);
}

#[test]
fn unrelated_worksheet_projection_maps_cut_and_fill_to_canonical_rows() {
    let mut b = controls();
    b.set_cell_value_tracked(1, 10, 1, "=A11+$A$1");
    b.set_cell_value_tracked(1, 9, 1, "Hidden");
    let mut v = RowView::new(30);
    let mut order: Vec<_> = (0..30).collect();
    order.swap(8, 10);
    let mut visible = vec![true; 30];
    visible[9] = false;
    v.restore(order, visible);
    let s = b.sheet(1).unwrap();
    let (ic, w) = plan_cut(s, &v, ((8, 1), (10, 1))).unwrap();
    assert_eq!(ic.source_rows, vec![10, 8]);
    let a = prepare_table_writes(&b, 1, &w).unwrap();
    assert_eq!(a.sheet(1).unwrap().get_raw(9, 1), "Hidden");
    let w = plan_direction(s, &v, ((8, 1), (10, 1)), true).unwrap();
    assert_eq!(w.len(), 1);
    assert_eq!((w[0].row, w[0].value.as_deref()), (8, Some("=A9+$A$1")));
    let single = plan_direction(s, &v, ((10, 1), (10, 1)), true).unwrap();
    assert_eq!(single[0].value, w[0].value);
}

#[test]
fn outside_direction_and_broadcast_preserve_formula_offsets_and_target_metadata() {
    let mut b = fixture(true);
    b.set_cell_value_tracked(0, 8, 2, "=$B9+C$1");
    b.active_sheet_mut().set_format(
        10,
        2,
        CellFormat {
            italic: true,
            ..Default::default()
        },
    );
    b.active_sheet_mut().set_comment(
        10,
        2,
        Some(CellComment {
            text: "Keep".into(),
            author: String::new(),
        }),
    );
    let s = b.active_sheet();
    let v = rows(&b, 0);
    let w = plan_direction(s, &v, ((8, 2), (10, 2)), true).unwrap();
    assert_eq!(
        w.iter()
            .map(|w| w.value.as_deref().unwrap())
            .collect::<Vec<_>>(),
        vec!["=$B10+C$1", "=$B11+C$1"]
    );
    let mut a = prepare_table_writes(&b, 0, &w).unwrap();
    assert!(a.active_sheet().get_format(10, 2).italic);
    assert_eq!(a.active_sheet().comment(10, 2), s.comment(10, 2));
    let h = commit(&b, &a, 0, &w);
    h.replay(&mut a, true).unwrap();
    h.replay(&mut a, false).unwrap();
    assert_eq!(records(&a), records(&b));
    assert_eq!(a.active_sheet().tables(), s.tables());
    let w = plan_direction(s, &v, ((8, 3), (8, 3)), false).unwrap();
    assert_eq!(w[0].value.as_deref(), Some("=$B9+D$1"));
    let w = plan_broadcast(
        s,
        &v,
        (8, 2),
        &[((8, 2), (9, 2)), ((12, 3), (12, 3))],
        Some("+SUM($B9,C$1"),
    )
    .unwrap();
    assert_eq!(w.len(), 3);
    assert_eq!(w[2].value.as_deref(), Some("=SUM($B13,D$1)"));
}

#[test]
fn other_sheet_fill_refreshes_filters_and_history_rewind() {
    let mut b = controls();
    b.set_cell_value_tracked(1, 0, 1, "-1");
    b.set_active_sheet(1);
    let w = plan_broadcast(
        b.sheet(1).unwrap(),
        &rows(&b, 1),
        (0, 1),
        &[((0, 0), (0, 1))],
        None,
    )
    .unwrap();
    let mut a = prepare_table_writes(&b, 1, &w).unwrap();
    assert_eq!(records(&a), vec![5, 6]);
    assert_eq!(a.sheet(0).unwrap().get_display(0, 1), "69");
    let h = commit(&b, &a, 1, &w);
    h.replay(&mut a, true).unwrap();
    assert_eq!(records(&a), vec![5, 3, 6]);
    h.replay(&mut a, false).unwrap();
    let mut history = History::new();
    history.record_action_with_provenance(
        UndoAction::TableCellsChanged {
            sheet_index: 1,
            commit: Box::new(h),
            description: "Fill controls".into(),
        },
        None,
    );
    let result = history
        .build_workbook_before(1, Some(&b), 100, 10_000)
        .unwrap();
    assert_eq!(records(&result.workbook), vec![5, 6]);
    assert!(result.view_state.per_sheet[0].table_rows.is_some());
    assert!(result.view_state.per_sheet[1].table_rows.is_none());
}

#[test]
fn outside_handle_supports_four_directions_and_ctrl_copy() {
    let mut b = controls();
    for (r, c, value) in [(8, 2, "10"), (9, 2, "20"), (8, 3, "20")] {
        b.set_cell_value_tracked(1, r, c, value);
    }
    let s = b.sheet(1).unwrap();
    let v = rows(&b, 1);
    for (source, end, vertical, expected) in [
        (((8, 2), (9, 2)), (11, 2), true, vec!["30", "40"]),
        (((8, 2), (9, 2)), (6, 2), true, vec!["0", "-10"]),
        (((8, 2), (8, 3)), (8, 5), false, vec!["30", "40"]),
        (((8, 2), (8, 3)), (8, 0), false, vec!["0", "-10"]),
    ] {
        let w = plan_handle(s, &v, source, end, vertical, false).unwrap();
        assert_eq!(
            w.iter()
                .map(|w| w.value.as_deref().unwrap())
                .collect::<Vec<_>>(),
            expected
        );
        assert!(prepare_table_writes(&b, 1, &w).is_ok());
    }
    let w = plan_handle(s, &v, ((8, 2), (9, 2)), (11, 2), true, true).unwrap();
    assert_eq!(
        w.iter()
            .map(|w| w.value.as_deref().unwrap())
            .collect::<Vec<_>>(),
        vec!["10", "20"]
    );
    assert!(plan_handle(s, &v, ((8, 2), (9, 2)), (30, 2), true, false).is_err());
}

#[test]
fn literal_fill_stays_text_and_does_not_change_table_rules() {
    let mut b = fixture(true);
    b.set_cell_text_tracked(0, 8, 1, "=literal");
    let w = plan_direction(b.active_sheet(), &rows(&b, 0), ((8, 1), (8, 3)), false).unwrap();
    let a = prepare_table_writes(&b, 0, &w).unwrap();
    assert_eq!(a.active_sheet().get_raw(8, 3), "=literal");
    assert!(!a.active_sheet().get_cell(8, 3).value().is_formula());
    assert_eq!(a.active_sheet().tables(), b.active_sheet().tables());
}

#[test]
fn protected_sources_targets_and_body_boundary_crossings_are_refused() {
    let b = fixture(true);
    let s = b.active_sheet();
    let v = rows(&b, 0);
    for rect in [
        ((2, 1), (2, 1)),
        ((6, 1), (7, 1)),
        ((3, 0), (6, 0)),
        ((29, 1), (30, 1)),
        ((9, 2), (8, 2)),
    ] {
        assert!(plan_cut(s, &v, rect).is_err());
        assert!(plan_direction(s, &v, rect, true).is_err());
    }
    assert!(plan_direction(s, &v, ((7, 2), (7, 2)), true).is_err());
    assert!(plan_direction(s, &v, ((0, 1), (0, 1)), true).is_err());
    assert!(plan_direction(s, &v, ((8, 0), (8, 0)), false).is_err());
    assert!(plan_broadcast(
        s,
        &v,
        (8, 2),
        &[((8, 2), (10, 2)), ((2, 1), (2, 1))],
        Some("9")
    )
    .is_err());
    assert!(plan_handle(s, &v, ((8, 2), (9, 2)), (5, 2), true, false).is_err());
    let mut b = controls();
    b.sheet_mut(1)
        .unwrap()
        .add_merge(MergedRegion {
            start: (8, 1),
            end: (8, 2),
        })
        .unwrap();
    let s = b.sheet(1).unwrap();
    let v = rows(&b, 1);
    assert!(plan_cut(s, &v, ((8, 1), (8, 1))).is_err());
    assert!(plan_direction(s, &v, ((9, 1), (9, 1)), true).is_err());
    assert!(plan_handle(s, &v, ((8, 1), (8, 1)), (10, 1), true, false).is_err());
    let mut b = controls();
    b.set_cell_value_tracked(1, 8, 1, "=SEQUENCE(2)");
    assert!(plan_direction(b.sheet(1).unwrap(), &rows(&b, 1), ((10, 1), (10, 1)), true).is_err());
}

#[test]
fn oversized_and_empty_selections_are_refused() {
    let s = Sheet::new_with_name(SheetId(123), 100_002, 2, "Large");
    let mut v = RowView::new(s.rows);
    assert!(plan_cut(&s, &v, ((0, 0), (100_000, 0))).is_err());
    assert!(plan_handle(&s, &v, ((0, 0), (0, 0)), (100_000, 0), true, false).is_err());
    assert!(plan_broadcast(
        &s,
        &v,
        (0, 0),
        &[((0, 0), (59_999, 0)), ((60_000, 0), (100_000, 0))],
        None
    )
    .is_err());
    v.apply_filter(vec![false; s.rows]);
    assert!(view_safe_selection_rows(&s, &v, ((0, 0), (1, 0))).is_err());
}

#[test]
fn unsafe_cross_sheet_cut_and_fill_publish_nothing() {
    let mut b = controls();
    b.set_cell_value_tracked(1, 0, 1, "0");
    b.set_cell_value_tracked(0, 0, 0, "=IF(Controls!A1=30,\"\",SEQUENCE(5))");
    let rev = b.revision();
    let v = rows(&b, 1);
    let s = b.sheet(1).unwrap();
    let (_, w) = plan_cut(s, &v, ((0, 0), (0, 1))).unwrap();
    assert!(prepare_table_writes(&b, 1, &w).is_err());
    let w = plan_broadcast(s, &v, (0, 1), &[((0, 0), (0, 1))], Some("0")).unwrap();
    assert!(prepare_table_writes(&b, 1, &w).is_err());
    assert_eq!(b.revision(), rev);
    assert_eq!(b.sheet(1).unwrap().get_raw(0, 0), "30");
    assert_eq!(records(&b), vec![5, 3, 6]);
}

#[test]
fn broadcast_can_mix_safe_outside_cells_and_visible_table_records_atomically() {
    let b = fixture(true);
    let w = plan_broadcast(
        b.active_sheet(),
        &rows(&b, 0),
        (8, 1),
        &[((8, 1), (8, 1)), ((3, 1), (6, 1))],
        Some("East"),
    )
    .unwrap();
    assert_eq!(w.len(), 4);
    let mut a = prepare_table_writes(&b, 0, &w).unwrap();
    assert!(records(&a).is_empty());
    assert_eq!(a.active_sheet().get_raw(8, 1), "East");
    assert_eq!(a.active_sheet().get_raw(4, 1), "East");
    let h = commit(&b, &a, 0, &w);
    h.replay(&mut a, true).unwrap();
    assert_eq!(records(&a), records(&b));
    assert_eq!(a.active_sheet().get_raw(8, 1), "");
}

#[test]
fn pivot_output_cannot_be_used_as_a_fill_source_or_cut_destination() {
    use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotField, PivotValueField};
    let mut b = controls();
    let id = b.sheet(0).unwrap().tables()[0].id;
    let source = b.table_pivot_source(id).unwrap();
    let field = &b.table(id).unwrap().1.columns[1];
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
    let (pivot, index) = b.create_pivot(source, definition).unwrap();
    let s = b.sheet(index).unwrap();
    let region = s
        .pivots
        .iter()
        .find(|p| p.id == pivot)
        .unwrap()
        .region()
        .unwrap();
    let v = rows(&b, index);
    let cell = (region.0, region.1);
    assert!(plan_cut(s, &v, (cell, cell)).is_err());
    assert!(plan_broadcast(s, &v, cell, &[((20, 0), (20, 1))], None).is_err());
}
