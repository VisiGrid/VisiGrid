use visigrid_engine::{
    sheet::{Sheet, SheetId},
    table::TableRange,
    workbook::Workbook,
};
use visigrid_io::{json, native};

fn table_book() -> Workbook {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(7), 100, 20, "Data")], 0);
    wb.set_cell_value_tracked(0, 0, 0, "42");
    wb.set_cell_value_tracked(0, 1, 0, "17");
    wb.create_table(
        SheetId(7),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 4,
            end_col: 1,
        },
        "Sales",
    )
    .unwrap();
    wb
}
fn check(wb: &Workbook) {
    let (_, t) = wb.table_by_name("sales").unwrap();
    assert_eq!(t.columns[0].name, "42");
    assert_eq!(t.columns[1].name, "Column2");
    assert_eq!(t.range.data_rows(), 4);
    assert_eq!(wb.active_sheet().get_raw(1, 0), "17");
    assert!(wb.active_sheet().table_value_write_error(0, 0).is_some());
}

#[test]
fn table_json_roundtrip_workbook_and_single_sheet() {
    let wb = table_book();
    for content in [
        json::export_workbook(&wb, &[], 0).unwrap(),
        json::export_full(wb.active_sheet()).unwrap(),
    ] {
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(&content).unwrap()["version"],
            3
        );
        let (loaded, _, _) = json::import_any(&content).unwrap();
        check(&loaded);
        assert_eq!(
            serde_json::to_value(wb.saved_tables()).unwrap(),
            serde_json::to_value(loaded.saved_tables()).unwrap()
        );
    }
}

#[test]
fn table_native_roundtrip_all_save_entrypoints() {
    let wb = table_book();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tables.sheet");
    native::save_workbook(&wb, &path).unwrap();
    check(&native::load_workbook(&path).unwrap());
    native::save(wb.active_sheet(), &path).unwrap();
    check(&native::load_workbook(&path).unwrap());
    assert_eq!(native::load(&path).unwrap().tables().len(), 1);
    native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap();
    check(&native::load_workbook(&path).unwrap());
    native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap();
    check(&native::load_workbook(&path).unwrap());
}

#[test]
fn table_deleted_identity_survives_json_and_native_reopen() {
    let mut wb = table_book();
    let id = wb.table_by_name("Sales").unwrap().1.id;
    wb.remove_table(id).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tables.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let content = json::export_workbook(&wb, &[], 0).unwrap();
    for mut loaded in [
        native::load_workbook(&path).unwrap(),
        json::import_any(&content).unwrap().0,
    ] {
        assert!(loaded.tables().next().is_none());
        let sid = loaded.active_sheet().id;
        let commit = loaded
            .create_table(
                sid,
                TableRange {
                    start_row: 0,
                    start_col: 0,
                    end_row: 0,
                    end_col: 0,
                },
                "Sales",
            )
            .unwrap();
        assert!(commit.table_id().0 > id.0);
    }
}

#[test]
fn table_invalid_metadata_is_rejected_instead_of_silently_dropped() {
    let wb = table_book();
    let content = json::export_workbook(&wb, &[], 0).unwrap();
    let mut doc: serde_json::Value = serde_json::from_str(&content).unwrap();
    doc["table_catalog"]["sheets"][0]["tables"][0]["columns"][0]["name"] = "Wrong".into();
    assert!(json::import_any(&doc.to_string())
        .unwrap_err()
        .contains("header"));
    doc.as_object_mut().unwrap().remove("table_catalog");
    assert!(json::import_any(&doc.to_string())
        .unwrap_err()
        .contains("missing"));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tables.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let conn = rusqlite::Connection::open(&path).unwrap();
    conn.execute("UPDATE meta SET value = 'invalid' WHERE key = 'tables'", [])
        .unwrap();
    assert!(native::load_workbook(&path)
        .unwrap_err()
        .contains("Tables metadata"));
}

#[test]
fn table_free_json_retains_legacy_versions() {
    let wb = Workbook::new();
    let single: serde_json::Value =
        serde_json::from_str(&json::export_full(wb.active_sheet()).unwrap()).unwrap();
    let multi: serde_json::Value =
        serde_json::from_str(&json::export_workbook(&wb, &[], 0).unwrap()).unwrap();
    assert_eq!(single["version"], 1);
    assert_eq!(multi["version"], 2);
    assert!(multi.get("table_catalog").is_none());
}

#[test]
fn table_catalog_remaps_sheet_ids_and_retains_column_allocators() {
    let mut wb = table_book();
    let index = wb.add_sheet_named("Second").unwrap();
    let sid = wb.sheet(index).unwrap().id;
    let id = wb
        .create_table(
            sid,
            TableRange {
                start_row: 3,
                start_col: 4,
                end_row: 3,
                end_col: 4,
            },
            "Items",
        )
        .unwrap()
        .table_id();
    let grow = wb
        .resize_table(
            id,
            TableRange {
                start_row: 3,
                start_col: 4,
                end_row: 3,
                end_col: 5,
            },
        )
        .unwrap();
    let old_column = wb.table(id).unwrap().1.columns[1].id;
    wb.apply_table_commit(&grow, true).unwrap();
    wb.set_active_sheet(index);
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("tables.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let content = json::export_workbook(&wb, &[], index).unwrap();
    for mut loaded in [
        native::load_workbook(&path).unwrap(),
        json::import_any(&content).unwrap().0,
    ] {
        assert_eq!(loaded.active_sheet_index(), index);
        assert_eq!(loaded.table(id).unwrap().0, loaded.sheet(index).unwrap().id);
        loaded
            .resize_table(
                id,
                TableRange {
                    start_row: 3,
                    start_col: 4,
                    end_row: 3,
                    end_col: 5,
                },
            )
            .unwrap();
        assert!(loaded.table(id).unwrap().1.columns[1].id.0 > old_column.0);
        assert_eq!(loaded.tables().count(), 2);
    }
}

#[test]
fn structured_formulas_survive_rename_save_reopen_and_recalculate() {
    let mut wb = table_book();
    let id = wb.table_by_name("Sales").unwrap().1.id;
    wb.set_cell_value_tracked(0, 1, 1, "=[@[42]]*2");
    let index = wb.add_sheet_named("Report").unwrap();
    wb.set_cell_value_tracked(index, 0, 0, "=SUM(Sales[Column2])");
    wb.rename_table_columns(id, &["Qty".into(), "Net [Amount]".into()])
        .unwrap();
    wb.rename_table(id, "Orders").unwrap();
    assert_eq!(wb.sheet(index).unwrap().get_display(0, 0), "34");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("structured.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let json = json::export_workbook(&wb, &[], 0).unwrap();
    for mut loaded in [
        native::load_workbook(&path).unwrap(),
        json::import_any(&json).unwrap().0,
    ] {
        assert_eq!(loaded.sheet(index).unwrap().get_display(0, 0), "34");
        assert_eq!(
            loaded.sheet(index).unwrap().get_raw(0, 0),
            wb.sheet(index).unwrap().get_raw(0, 0)
        );
        loaded.set_cell_value_tracked(0, 1, 0, "25");
        assert_eq!(loaded.sheet(index).unwrap().get_display(0, 0), "50");
        let mut extent = loaded.table(id).unwrap().1.range;
        extent.end_row = 5;
        loaded.set_cell_value_tracked(0, 5, 1, "9");
        loaded.resize_table(id, extent).unwrap();
        assert_eq!(loaded.sheet(index).unwrap().get_display(0, 0), "59");
    }
}

#[test]
fn table_shape_is_part_of_semantic_fingerprint() {
    let mut wb = table_book();
    let id = wb.table_by_name("Sales").unwrap().1.id;
    let before = native::compute_semantic_fingerprint(&wb);
    assert!(before.starts_with("v3:"));
    let mut extent = wb.table(id).unwrap().1.range;
    extent.end_row += 1;
    let resize = wb.resize_table(id, extent).unwrap();
    assert_ne!(native::compute_semantic_fingerprint(&wb), before);
    wb.apply_table_commit(&resize, true).unwrap();
    assert_eq!(native::compute_semantic_fingerprint(&wb), before);
    assert!(native::compute_semantic_fingerprint(&Workbook::new()).starts_with("v2:"));
}

#[test]
fn appended_records_and_structured_dependencies_survive_native_and_json() {
    let mut wb = table_book();
    let id = wb.table_by_name("Sales").unwrap().1.id;
    let columns = wb.table(id).unwrap().1.columns.clone();
    wb.set_cell_value_tracked(0, 0, 4, "=SUM(Sales[Column2])");
    wb.append_table_rows(id, 2, &[(5, 0, "3".into()), (5, 1, "=[@[42]]*10".into())])
        .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("appended.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let content = json::export_workbook(&wb, &[], 0).unwrap();
    for mut loaded in [
        native::load_workbook(&path).unwrap(),
        json::import_any(&content).unwrap().0,
    ] {
        assert_eq!(loaded.table(id).unwrap().1.range.end_row, 6);
        assert_eq!(loaded.table(id).unwrap().1.columns, columns);
        assert_eq!(loaded.active_sheet().get_raw(6, 0), "");
        assert_eq!(loaded.active_sheet().get_display(0, 4), "30");
        loaded.set_cell_value_tracked(0, 5, 0, "4");
        assert_eq!(loaded.active_sheet().get_display(0, 4), "40");
    }
}
