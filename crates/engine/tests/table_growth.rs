use visigrid_engine::{
    cell::CellComment,
    sheet::{MergedRegion, Sheet, SheetId},
    table::{TableId, TableRange},
    workbook::Workbook,
};

fn range(sr: usize, sc: usize, er: usize, ec: usize) -> TableRange {
    TableRange {
        start_row: sr,
        start_col: sc,
        end_row: er,
        end_col: ec,
    }
}

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 2000, 20, "Data")], 0);
    for (col, text) in ["Qty", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, col, text);
    }
    wb.set_cell_value_tracked(0, 1, 0, "2");
    wb.set_cell_value_tracked(0, 1, 1, "=[@Qty]*10");
    let id = wb
        .create_table(SheetId(1), range(0, 0, 1, 1), "Sales")
        .unwrap()
        .table_id();
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[Amount])");
    (wb, id)
}

#[test]
fn append_is_one_revision_and_replays_values_bounds_and_dependencies() {
    let (mut wb, id) = book();
    let revision = wb.revision();
    let columns = wb.table(id).unwrap().1.columns.clone();
    let commit = wb
        .append_table_rows(id, 1, &[(2, 0, "3".into()), (2, 1, "=[@Qty]*10".into())])
        .unwrap();
    assert_eq!(wb.revision(), revision + 1);
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "50");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 1);
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 0), "");
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 1), "");
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "20");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.columns, columns);
    wb.set_cell_value_tracked(0, 2, 0, "4");
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "60");
}

#[test]
fn thousand_row_paste_includes_blank_records_and_restores_overwritten_body() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 1500, 3, "adjacent");
    let writes: Vec<_> = (1..=1001)
        .flat_map(|r| {
            [
                (r, 0, if r == 1001 { String::new() } else { "1".into() }),
                (
                    r,
                    1,
                    if r == 1001 {
                        String::new()
                    } else {
                        "10".into()
                    },
                ),
            ]
        })
        .collect();
    let commit = wb.append_table_rows(id, 1000, &writes).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 1001);
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "10000");
    assert_eq!(wb.sheet(0).unwrap().get_raw(1500, 3), "adjacent");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 1), "=[@Qty]*10");
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "20");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 1001);
}

#[test]
fn append_intent_is_distinct_from_ordinary_writes() {
    let (mut wb, id) = book();
    assert_eq!(
        wb.table_append_target(SheetId(1), range(2, 0, 2, 0))
            .unwrap(),
        Some(id)
    );
    assert_eq!(
        wb.table_append_target(SheetId(1), range(1, 0, 5, 1))
            .unwrap(),
        Some(id)
    );
    assert_eq!(
        wb.table_append_target(SheetId(1), range(3, 0, 4, 1))
            .unwrap(),
        None
    );
    assert_eq!(
        wb.table_append_target(SheetId(1), range(2, 2, 3, 2))
            .unwrap(),
        None
    );
    assert!(wb
        .table_append_target(SheetId(1), range(1, 0, 3, 2))
        .unwrap_err()
        .contains("Resize"));
    wb.set_cell_value_tracked(0, 2, 0, "7");
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 1);
    assert!(wb.append_table_rows(id, 1, &[]).is_err());
    wb.resize_table(id, range(0, 0, 2, 1)).unwrap();
    wb.clear_cell_tracked(0, 2, 0);
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 2);
}

#[test]
fn collision_refusals_are_atomic_even_when_paste_overwrites_existing_body() {
    for collision in 0..5 {
        let (mut wb, id) = book();
        match collision {
            0 => {
                wb.set_cell_value_tracked(0, 2, 1, "neighbor");
            }
            1 => {
                wb.sheet_mut(0)
                    .unwrap()
                    .add_merge(MergedRegion {
                        start: (2, 0),
                        end: (2, 1),
                    })
                    .unwrap();
            }
            2 => {
                wb.create_table(SheetId(1), range(2, 0, 3, 1), "Other")
                    .unwrap();
            }
            3 => {
                wb.set_cell_value_tracked(0, 2, 1, "=SEQUENCE(2)");
            }
            _ => {
                wb.sheet_mut(0).unwrap().set_comment(
                    2,
                    1,
                    Some(CellComment {
                        text: "note".into(),
                        author: "Author".into(),
                    }),
                );
            }
        }
        let revision = wb.revision();
        assert!(wb
            .append_table_rows(id, 2, &[(1, 0, "99".into()), (2, 0, "4".into())])
            .is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.table(id).unwrap().1.range.end_row, 1);
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 0), "2");
        assert_eq!(
            wb.sheet(0).unwrap().get_raw(2, 0),
            if collision == 2 { "Column1" } else { "" }
        );
    }
}

#[test]
fn stale_append_replay_never_releases_or_absorbs_new_values() {
    let (mut wb, id) = book();
    let commit = wb.append_table_rows(id, 1, &[(2, 0, "3".into())]).unwrap();
    wb.set_cell_value_tracked(0, 2, 1, "new");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    wb.clear_cell_tracked(0, 2, 1);
    wb.set_cell_value_tracked(0, 2, 0, "changed");
    assert!(wb.apply_table_commit(&commit, true).is_err());
    wb.set_cell_value_tracked(0, 2, 0, "3");
    wb.apply_table_commit(&commit, true).unwrap();
    wb.set_cell_value_tracked(0, 2, 1, "new");
    assert!(wb.apply_table_commit(&commit, false).is_err());
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 1);
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 1), "new");
}

#[test]
fn header_only_add_row_is_sparse_and_grid_edge_refuses_atomically() {
    let (mut wb, id) = book();
    wb.resize_table(id, range(0, 0, 0, 1)).unwrap();
    wb.clear_cell_tracked(0, 1, 0);
    wb.clear_cell_tracked(0, 1, 1);
    let before = wb.sheet(0).unwrap().cells_iter().count();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().cells_iter().count(), before);
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 1);
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 0);
    let revision = wb.revision();
    assert!(wb.append_table_rows(id, 2000, &[]).is_err());
    assert!(wb.append_table_rows(id, usize::MAX, &[]).is_err());
    assert_eq!(wb.revision(), revision);
}

#[test]
fn row_history_restores_last_body_rows_with_stable_identity() {
    let (mut wb, id) = book();
    let original = wb.table(id).unwrap().1.clone();
    let history = wb
        .prepare_table_row_history(0, 1, 1, true)
        .unwrap()
        .unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 0);
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "0");
    wb.apply_table_row_history(&history, true).unwrap();
    // Cell payloads belong to the ordinary row history, restored after bounds.
    wb.set_cell_value_tracked(0, 1, 0, "2");
    wb.set_cell_value_tracked(0, 1, 1, "=[@Qty]*10");
    assert_eq!(wb.table(id).unwrap().1, &original);
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "20");
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 0);
}

#[test]
fn row_history_moves_a1_and_structured_formulas_and_refuses_stale_schema() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(1, 1, 0, "=Data!B2");
    let history = wb
        .prepare_table_row_history(0, 1, 2, false)
        .unwrap()
        .unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 3);
    assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "=[@[Qty]]*10");
    assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), "=Data!B4");
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "20");
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), "=Data!B2");
    wb.rename_table(id, "Orders").unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_row_history(&history, false).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 1);
}
