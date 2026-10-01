use visigrid_engine::{
    cell::CellValue,
    named_range::NamedRange,
    sheet::{MergedRegion, Sheet, SheetId},
    structural::Axis,
    table::TableRange,
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
fn book() -> Workbook {
    Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 100, 20, "Data")], 0)
}

#[test]
fn table_creation_normalizes_only_headers_and_undo_preserves_types() {
    let mut wb = book();
    for (c, value) in ["Amount", "Amount", "Amount2", "42", "=2+3"]
        .iter()
        .enumerate()
    {
        wb.set_cell_value_tracked(0, 0, c, value);
    }
    wb.set_cell_value_tracked(0, 1, 0, "17");
    let commit = wb
        .create_table(SheetId(1), range(0, 0, 99, 4), "Sales")
        .unwrap();
    assert_eq!(commit.header_cell_count(), 3);
    let table = wb.table(commit.table_id()).unwrap().1;
    assert_eq!(
        table
            .columns
            .iter()
            .map(|c| c.name.as_str())
            .collect::<Vec<_>>(),
        ["Amount", "Amount3", "Amount2", "42", "5"]
    );
    assert_eq!(wb.active_sheet().get_raw(1, 0), "17");
    assert!(matches!(
        wb.active_sheet().get_cell(0, 3).value,
        CellValue::Text(_)
    ));
    wb.apply_table_commit(&commit, true).unwrap();
    assert!(wb.tables().next().is_none());
    assert!(matches!(
        wb.active_sheet().get_cell(0, 3).value,
        CellValue::Number(42.0)
    ));
    assert_eq!(wb.active_sheet().get_raw(0, 4), "=2+3");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.table(commit.table_id()).unwrap().1.range.data_rows(), 99);
}

#[test]
fn table_ids_and_column_ids_are_not_reused_after_undo_or_shrink() {
    let mut wb = book();
    let create = wb
        .create_table(SheetId(1), range(0, 0, 0, 0), "Sales")
        .unwrap();
    let id = create.table_id();
    let grow = wb.resize_table(id, range(0, 0, 0, 1)).unwrap();
    let old_column = wb.table(id).unwrap().1.columns[1].id;
    wb.apply_table_commit(&grow, true).unwrap();
    wb.apply_table_commit(&create, true).unwrap();
    wb.apply_table_commit(&create, false).unwrap();
    wb.resize_table(id, range(0, 0, 0, 1)).unwrap();
    assert!(wb.table(id).unwrap().1.columns[1].id.0 > old_column.0);
    let allocated = wb.table(id).unwrap().1.columns[1].id;
    wb.resize_table(id, range(0, 0, 0, 0)).unwrap();
    assert_eq!(wb.active_sheet().get_raw(0, 1), "Column2");
    wb.resize_table(id, range(0, 0, 0, 1)).unwrap();
    assert!(wb.table(id).unwrap().1.columns[1].id.0 > allocated.0);
    wb.remove_table(id).unwrap();
    let new = wb
        .create_table(SheetId(1), range(0, 0, 0, 1), "Sales")
        .unwrap();
    assert!(new.table_id().0 > id.0);
}

#[test]
fn table_schema_renames_are_atomic_and_share_named_range_namespace() {
    let mut wb = book();
    wb.named_ranges_mut()
        .set(NamedRange::cell("Reserved", 0, 5, 0))
        .unwrap();
    assert!(wb
        .create_table(SheetId(1), range(0, 0, 1, 1), "reserved")
        .is_err());
    let id = wb
        .create_table(SheetId(1), range(0, 0, 1, 1), "Sales")
        .unwrap()
        .table_id();
    assert!(wb
        .named_ranges_mut()
        .set(NamedRange::cell("SALES", 0, 5, 0))
        .is_err());
    assert!(wb
        .rename_table_columns(id, &["Same".into(), "same".into()])
        .is_err());
    let before = wb.table(id).unwrap().1.clone();
    let rename = wb
        .rename_table_columns(id, &["Column2".into(), "Column1".into()])
        .unwrap();
    assert_eq!(wb.table(id).unwrap().1.columns[0].id, before.columns[0].id);
    wb.apply_table_commit(&rename, true).unwrap();
    wb.rename_table(id, "Revenue").unwrap();
    wb.named_ranges_mut()
        .set(NamedRange::cell("Sales", 0, 5, 0))
        .unwrap();
    assert!(wb.rename_table(id, "Sales").is_err());
}

#[test]
fn table_headers_refuse_ordinary_writes_but_body_remains_editable() {
    let mut wb = book();
    wb.create_table(SheetId(1), range(0, 0, 2, 0), "Sales")
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "Broken");
    wb.clear_cell_tracked(0, 0, 0);
    wb.active_sheet_mut()
        .freeze_cell(0, 0, CellValue::Number(1.0), "=1".into());
    assert_eq!(wb.active_sheet().get_raw(0, 0), "Column1");
    wb.set_cell_value_tracked(0, 1, 0, "=2+3");
    assert_eq!(wb.active_sheet().get_display(1, 0), "5");
}

#[test]
fn table_region_excludes_merges_other_tables_and_spills_in_both_directions() {
    let mut wb = book();
    wb.active_sheet_mut()
        .add_merge(MergedRegion {
            start: (3, 0),
            end: (3, 1),
        })
        .unwrap();
    assert!(wb
        .create_table(SheetId(1), range(0, 0, 4, 1), "Sales")
        .is_err());
    wb.active_sheet_mut().set_merges(vec![]);
    let id = wb
        .create_table(SheetId(1), range(0, 0, 4, 1), "Sales")
        .unwrap()
        .table_id();
    assert!(wb
        .create_table(SheetId(1), range(4, 1, 5, 2), "Other")
        .is_err());
    assert!(wb
        .active_sheet_mut()
        .add_merge(MergedRegion {
            start: (3, 0),
            end: (3, 1)
        })
        .is_err());
    wb.active_sheet_mut().set_merges(vec![MergedRegion {
        start: (3, 0),
        end: (3, 1),
    }]);
    assert!(wb.active_sheet().merged_regions.is_empty());
    wb.set_cell_value_tracked(0, 1, 0, "=SEQUENCE(2)");
    assert!(wb.active_sheet().get_display(1, 0).starts_with("#SPILL!"));
    wb.remove_table(id).unwrap();
    wb.set_cell_value_tracked(0, 1, 0, "=SEQUENCE(2)");
    assert!(wb
        .create_table(SheetId(1), range(0, 0, 4, 1), "Sales")
        .is_err());
}

#[test]
fn table_structural_edits_move_and_resize_rows_but_protect_schema() {
    let mut wb = book();
    let id = wb
        .create_table(SheetId(1), range(2, 2, 5, 3), "Sales")
        .unwrap()
        .table_id();
    wb.structural_edit(0, Axis::Row, 0, 2, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range(4, 2, 7, 3));
    wb.structural_edit(0, Axis::Row, 6, 2, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range(4, 2, 9, 3));
    wb.structural_edit(0, Axis::Row, 5, 5, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range(4, 2, 4, 3));
    assert!(wb.structural_edit(0, Axis::Row, 4, 1, true).is_err());
    wb.structural_edit(0, Axis::Col, 3, 1, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.columns.len(), 3);
    wb.structural_edit(0, Axis::Col, 0, 1, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range(4, 3, 4, 5));
    assert_eq!(wb.active_sheet().get_raw(4, 3), "Column1");
}

#[test]
fn table_stale_commit_fails_without_overwriting_new_header_edits() {
    let mut wb = book();
    let create = wb
        .create_table(SheetId(1), range(0, 0, 1, 0), "Sales")
        .unwrap();
    wb.apply_table_commit(&create, true).unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "New edit");
    assert!(wb.apply_table_commit(&create, false).is_err());
    assert_eq!(wb.active_sheet().get_raw(0, 0), "New edit");
    assert!(wb.tables().next().is_none());
}

#[test]
fn table_catalog_invalid_restore_is_atomic_and_sheet_removal_releases_names() {
    let mut wb = book();
    let id = wb
        .create_table(SheetId(1), range(0, 0, 1, 0), "Sales")
        .unwrap()
        .table_id();
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0].columns[0].name = "Wrong".into();
    assert!(wb.restore_tables(catalog).is_err());
    assert_eq!(wb.table(id).unwrap().1.columns[0].name, "Column1");
    wb.add_sheet_named("Other").unwrap();
    let sheet = wb.take_sheet(0).unwrap();
    wb.named_ranges_mut()
        .set(NamedRange::cell("Sales", 0, 5, 0))
        .unwrap();
    assert!(!wb.restore_sheet(0, sheet.clone()));
    wb.named_ranges_mut().remove("Sales");
    assert!(wb.restore_sheet(0, sheet));
    assert_eq!(wb.table_by_name("sales").unwrap().1.id, id);
    assert!(wb
        .add_sheet_clone_named(&wb.sheet(0).unwrap().clone(), "Copy")
        .is_none());
}

#[test]
fn table_undo_remove_checks_unchanged_headers_before_restoring_schema() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "Amount");
    let create = wb
        .create_table(SheetId(1), range(0, 0, 1, 0), "Sales")
        .unwrap();
    assert_eq!(create.header_cell_count(), 0);
    let remove = wb.remove_table(create.table_id()).unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "Changed");
    assert!(wb.apply_table_commit(&remove, true).is_err());
    assert!(wb.tables().next().is_none());
    wb.set_cell_value_tracked(0, 0, 0, "Amount");
    wb.apply_table_commit(&remove, true).unwrap();
    wb.apply_table_commit(&create, true).unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "Changed again");
    assert!(wb.apply_table_commit(&create, false).is_err());
}
