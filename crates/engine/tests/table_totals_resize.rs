use visigrid_engine::{
    cell::CellComment,
    sheet::{Sheet, SheetId},
    table::{TableId, TableRange, TableTotal},
    workbook::Workbook,
};

fn book(visible: bool) -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 40, 10, "Data")], 0);
    for (col, value) in ["Qty", "Amount"].iter().enumerate() { wb.set_cell_value_tracked(0, 1, col, value); }
    for row in 2..=3 { wb.set_cell_value_tracked(0, row, 0, "2"); }
    let id = wb.create_table(SheetId(1), TableRange { start_row: 1, end_row: 3, start_col: 0, end_col: 1 }, "Sales").unwrap().table_id();
    wb.set_calculated_column(id, 1, 2, "=[@Qty]*10", true).unwrap();
    wb.set_cell_value_tracked(0, 3, 1, "777");
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
    (wb, id)
}

#[test]
fn widening_totals_keeps_values_rules_settings_and_latent_local_formulas() {
    for visible in [true, false] {
        let (mut wb, id) = book(visible);
        wb.set_cell_value_tracked(0, 1, 2, "Qty");
        wb.set_cell_value_tracked(0, 2, 2, "=[@Qty]*3");
        wb.set_cell_value_tracked(0, 3, 2, "9");
        let original = wb.table(id).unwrap().1.clone();
        let revision = wb.revision();
        let commit = wb.resize_table(id, TableRange { end_col: 2, ..original.range }).unwrap();
        assert_eq!(wb.revision(), revision + 1);
        let expanded = wb.table(id).unwrap().1;
        assert_eq!(expanded.columns[0].id, original.columns[0].id);
        assert_eq!(expanded.columns[2].name, "Qty2");
        assert_eq!(expanded.totals.as_ref().unwrap().columns.len(), 3);
        assert_eq!(expanded.totals.as_ref().unwrap().columns[2], Default::default());
        assert_eq!(expanded.totals.as_ref().unwrap().columns[1], original.totals.as_ref().unwrap().columns[1]);
        let allocated = expanded.columns[2].id;
        assert_eq!(wb.sheet(0).unwrap().get_display(2, 2), "6");
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "777");
        if visible { assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "797"); }
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns, original.columns);
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "Qty");
        assert!(wb.sheet(0).unwrap().get_display(2, 2).starts_with('#'));
        wb.apply_table_commit(&commit, false).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns[2].id, allocated);
        assert_eq!(wb.sheet(0).unwrap().get_display(2, 2), "6");
        wb.apply_table_commit(&commit, true).unwrap();
        wb.resize_table(id, TableRange { end_col: 2, ..original.range }).unwrap();
        assert!(wb.table(id).unwrap().1.columns[2].id.0 > allocated.0);
    }
}

#[test]
fn shrinking_rewrites_retained_and_foreign_totals_and_preserves_released_cells() {
    for visible in [true, false] {
        let (mut wb, id) = book(visible);
        if !visible { wb.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        wb.set_table_total(id, 0, TableTotal { function: Some("custom".into()), formula: Some("=SUM([Amount])".into()), label: None }).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        let other = wb.add_sheet_named("Other").unwrap();
        wb.set_cell_value_tracked(other, 0, 0, "Value");
        wb.set_cell_value_tracked(other, 1, 0, "1");
        let other_id = wb.create_table(wb.sheet(other).unwrap().id, TableRange { start_row: 0, end_row: 1, start_col: 0, end_col: 0 }, "OtherTable").unwrap().table_id();
        wb.set_table_totals_visible(other_id, true, Default::default()).unwrap();
        wb.set_table_total(other_id, 0, TableTotal { function: Some("custom".into()), formula: Some("=SUM(Sales[Amount])+SUM([Value])".into()), label: None }).unwrap();
        if !visible { wb.set_table_totals_visible(other_id, false, Default::default()).unwrap(); }
        let range = wb.table(id).unwrap().1.range;
        let commit = wb.resize_table(id, TableRange { end_col: 0, ..range }).unwrap();
        assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns.len(), 1);
        assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_deref(), Some("=SUM(#REF!)"));
        assert!(wb.table(other_id).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_ref().unwrap().contains("SUM(#REF!)"));
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "777");
        if visible {
            assert_eq!(wb.sheet(0).unwrap().get_raw(4, 0), "=SUM(#REF!)");
            assert!(wb.sheet(0).unwrap().get_raw(4, 1).contains("#REF!"));
            assert!(wb.sheet(other).unwrap().get_raw(2, 0).contains("SUM(#REF!)"));
        }
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_deref(), Some("=SUM([Amount])"));
        assert_eq!(wb.table(other_id).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_deref(), Some("=SUM(Sales[Amount])+SUM([Value])"));
        if visible { assert_eq!(wb.sheet(other).unwrap().get_display(2, 0), "798"); }
        wb.apply_table_commit(&commit, false).unwrap();
    }
}

#[test]
fn combined_resize_moves_only_surviving_footer_columns_and_is_atomic() {
    let (mut wb, id) = book(true);
    wb.sheet_mut(0).unwrap().set_comment(4, 1, Some(CellComment { text: "Released footer".into(), author: "QA".into() }));
    let original = wb.table(id).unwrap().1.range;
    let commit = wb.resize_table(id, TableRange { end_row: 6, end_col: 0, ..original }).unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(7));
    assert_eq!(wb.sheet(0).unwrap().get_raw(7, 0), "Total");
    assert_eq!(wb.sheet(0).unwrap().get_cell(4, 1).comment().map(|c| c.text.as_str()), Some("Released footer"));
    assert!(wb.sheet(0).unwrap().get_raw(4, 1).contains("#REF!"));
    assert_eq!(wb.sheet(0).unwrap().get_raw(7, 1), "");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "797");
    wb.set_cell_value_tracked(0, 1, 2, "Extra");
    wb.set_cell_value_tracked(0, 4, 2, "Existing body data");
    let commit = wb.resize_table(id, TableRange { end_row: 6, end_col: 2, ..original }).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "Existing body data");
    assert_eq!(wb.sheet(0).unwrap().get_display(7, 1), "797");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "797");
    wb.set_cell_value_tracked(0, 7, 2, "Protect destination");
    let revision = wb.revision();
    assert!(wb.resize_table(id, TableRange { end_row: 6, end_col: 2, ..original }).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.range, original);
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "797");
}

#[test]
fn new_footer_ownership_and_stale_replay_are_checked_before_writes() {
    let (mut wb, id) = book(true);
    let range = wb.table(id).unwrap().1.range;
    wb.set_cell_value_tracked(0, 4, 2, "Note");
    let revision = wb.revision();
    assert!(wb.resize_table(id, TableRange { end_col: 2, ..range }).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "Note");
    wb.clear_cell_tracked(0, 4, 2);
    let commit = wb.resize_table(id, TableRange { end_col: 2, ..range }).unwrap();
    wb.set_cell_value_tracked(0, 10, 5, "Later edit");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.columns.len(), 3);
}

#[test]
fn resizing_undo_composes_with_totals_and_calculated_rule_edits() {
    let (mut wb, id) = book(true);
    let range = wb.table(id).unwrap().1.range;
    let resize = wb.resize_table(id, TableRange { end_col: 2, ..range }).unwrap();
    let total = wb.set_table_total(id, 2, TableTotal { function: Some("sum".into()), ..Default::default() }).unwrap();
    wb.apply_table_commit(&total, true).unwrap();
    let rule = wb.set_calculated_column(id, 2, 2, "=[@Qty]*3", true).unwrap();
    wb.apply_table_commit(&rule, true).unwrap();
    wb.apply_table_commit(&resize, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range);
    assert!(wb.sheet(0).unwrap().get_cell_opt(2, 2).is_none());
    assert!(wb.sheet(0).unwrap().get_cell_opt(4, 2).is_none());
    wb.apply_table_commit(&resize, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.columns.len(), 3);
}
