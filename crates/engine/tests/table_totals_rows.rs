use visigrid_engine::{
    cell::CellComment,
    sheet::{Sheet, SheetId},
    structural::Axis,
    table::{TableId, TableRange, TableTotal},
    workbook::{StructureStep, Workbook},
};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 40, 8, "Data")], 0);
    for (col, text) in ["Qty", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 2, col, text);
    }
    for row in 3..=5 { wb.set_cell_value_tracked(0, row, 0, "2"); }
    let id = wb.create_table(SheetId(1), TableRange { start_row: 2, end_row: 5, start_col: 0, end_col: 1 }, "Sales").unwrap().table_id();
    wb.set_calculated_column(id, 1, 3, "=A4*10", true).unwrap();
    wb.set_cell_value_tracked(0, 4, 1, "999");
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let other = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=Data!B7");
    wb.set_cell_value_tracked(other, 1, 0, "=SUM(Sales[#Totals])");
    (wb, id)
}

#[test]
fn inserts_at_footer_fill_body_move_comments_and_rewrite_fixed_links() {
    let (mut wb, id) = book();
    wb.sheet_mut(0).unwrap().set_comment(6, 1, Some(CellComment { text: "Footer".into(), author: "QA".into() }));
    wb.sheet_mut(0).unwrap().toggle_bold(6, 1);
    let columns = wb.table(id).unwrap().1.columns.clone();
    let history = wb.prepare_table_row_history(0, 6, 2, false).unwrap().unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 7);
    assert_eq!(wb.sheet(0).unwrap().get_raw(6, 1), "=A7*10");
    assert_eq!(wb.sheet(0).unwrap().get_raw(7, 1), "=A8*10");
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 1), "1039");
    assert_eq!(wb.sheet(0).unwrap().get_cell(8, 1).comment().map(|c| c.text.as_str()), Some("Footer"));
    assert!(wb.sheet(0).unwrap().get_format(8, 1).bold);
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=Data!B9");
    assert_eq!(wb.sheet(1).unwrap().get_display(1, 0), "1039");
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.columns, columns);
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 5);
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=Data!B7");
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 1), "1039");
}

#[test]
fn row_deletion_preserves_hidden_positions_custom_settings_and_restores_footers() {
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_total(id, 1, TableTotal {
            function: Some("custom".into()), formula: Some("=SUBTOTAL(109,[Amount])+A4".into()), label: None,
        }).unwrap();
        let mut catalog = wb.saved_tables();
        catalog.sheets[0].tables[0].totals.as_mut().unwrap().hidden_rows = [4, 12].into();
        wb.restore_tables(catalog).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, [4, 12].into()).unwrap(); }
        let original = wb.table(id).unwrap().1.clone();
        let history = wb.prepare_table_row_history(0, 3, 1, true).unwrap().unwrap();
        wb.apply_table_row_history(&history, false).unwrap();
        let totals = wb.table(id).unwrap().1.totals.as_ref().unwrap();
        assert_eq!(totals.hidden_rows, [3, 11].into());
        assert!(totals.columns[1].formula.as_ref().unwrap().contains("#REF!"));
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "999");
        wb.apply_table_row_history(&history, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.totals, original.totals);
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "=A4*10");
    }
    let (mut wb, id) = book();
    let history = wb.prepare_table_row_history(0, 6, 1, true).unwrap().unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), None);
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().shown, Some(false));
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[1].function.as_deref(), Some("sum"));
    assert!(wb.sheet(1).unwrap().get_raw(0, 0).contains("#REF!"));
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(6, 1), "1039");
    wb.apply_table_row_history(&history, false).unwrap();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(6, 1), "1039");
}

#[test]
fn row_moves_preserve_dormant_templates_empty_body_and_boundaries() {
    let (mut wb, id) = book();
    wb.set_table_total(id, 1, TableTotal { function: Some("custom".into()), formula: Some("=SUM([Amount])+A4".into()), label: None }).unwrap();
    wb.set_table_totals_visible(id, false, Default::default()).unwrap();
    let history = wb.prepare_table_row_history(0, 0, 2, false).unwrap().unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.start_row, 4);
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[1].formula.as_deref(), Some("=SUM([Amount])+A6"));
    wb.apply_table_row_history(&history, true).unwrap();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(6, 1), "1041");
    let (mut emptied, delete) = wb.prepare_guarded_structure(0, vec![StructureStep { axis: Axis::Row, at: 3, count: 3, delete: true }]).unwrap();
    assert_eq!(emptied.table(id).unwrap().1.range.data_rows(), 0);
    assert_eq!(emptied.table(id).unwrap().1.totals_row(), Some(3));
    emptied.structural_edit(0, Axis::Row, 3, 1, false).unwrap();
    assert_eq!(emptied.table(id).unwrap().1.range.data_rows(), 1);
    assert_eq!(emptied.sheet(0).unwrap().get_display(3, 1), "0");
    assert!(delete.candidate(&emptied, true).is_err());
    let revision = wb.revision();
    assert!(wb.structural_edit(0, Axis::Row, 2, 1, true).is_err());
    assert!(wb.structural_edit(0, Axis::Row, 0, 34, false).is_err());
    assert_eq!(wb.revision(), revision);
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0].totals.as_mut().unwrap().hidden_rows.insert(39);
    wb.restore_tables(catalog).unwrap();
    let revision = wb.revision();
    assert!(wb.structural_edit(0, Axis::Row, 10, 1, false).is_err());
    assert_eq!(wb.revision(), revision);
}


#[test]
fn headerless_creation_moves_neighboring_totals_and_undoes_as_one_transaction() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 0, 4, "First");
    wb.set_cell_value_tracked(0, 1, 4, "Second");
    wb.set_cell_value_tracked(1, 2, 0, "=ROWS(NewData[Column1])");
    let commit = wb.create_table_without_headers(SheetId(1), TableRange {
        start_row: 0, end_row: 1, start_col: 4, end_col: 4,
    }, "NewData").unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(7));
    assert_eq!(wb.sheet(1).unwrap().get_display(2, 0), "2");
    assert_eq!(wb.sheet(0).unwrap().get_display(7, 1), "1039");
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 4), "First");
    wb.apply_table_commit(&commit, true).unwrap();
    assert!(wb.table_by_name("NewData").is_none());
    assert!(wb.sheet(1).unwrap().get_display(2, 0).starts_with("#"));
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(6));
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 4), "First");
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=Data!B7");
    wb.apply_table_commit(&commit, false).unwrap();
    assert!(wb.table_by_name("NewData").is_some());
    assert_eq!(wb.sheet(1).unwrap().get_display(2, 0), "2");
    assert_eq!(wb.sheet(0).unwrap().get_display(7, 1), "1039");
}


#[test]
fn structural_replay_invalidates_indirect_footer_dependents_on_other_sheets() {
    let (mut wb, _) = book();
    wb.clear_cell_tracked(1, 0, 0); // Leave only the unchanged structured formula.
    let history = wb.prepare_table_row_history(0, 4, 1, true).unwrap().unwrap();
    for undo in [false, true, false] {
        let generation = wb.sheet(1).unwrap().edit_generation();
        wb.apply_table_row_history(&history, undo).unwrap();
        assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), "=SUM(Sales[#Totals])");
        assert_eq!(wb.sheet(1).unwrap().get_display(1, 0), if undo { "1039" } else { "40" });
        assert!(wb.sheet(1).unwrap().edit_generation() > generation);
    }
}
