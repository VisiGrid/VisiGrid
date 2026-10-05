use std::io::{Cursor, Read};
use visigrid_engine::{
    filter::SortDirection,
    table::TableRange,
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};
use visigrid_io::{json, native, xlsx};

fn book() -> Workbook {
    let mut wb = Workbook::new();
    for (row, value) in ["10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    wb.set_cell_value_tracked(0, 0, 2, "=SUBTOTAL(109,A1:A3)");
    wb.prepare_table_row_visibility(wb.active_sheet_id(), [1].into())
        .unwrap()
        .0
}
fn check(wb: &Workbook) {
    assert_eq!(wb.active_sheet().manual_hidden_rows(), [1].into());
    assert_eq!(wb.active_sheet().get_display(0, 2), "40");
}

#[test]
fn moved_hidden_totals_and_row_flags_survive_native_json_and_stored_excel() {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 0,
            },
            "Data",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Descending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    let (mut wb, _) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [4, 5, 6].into())
        .unwrap();
    wb.append_table_rows(id, 1, &[(4, 0, "7".into())]).unwrap();
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..3 {
        let path = dir.path().join(if mode == 2 {
            "hidden.xlsx"
        } else {
            "hidden.sheet"
        });
        let mut loaded = match mode {
            0 => {
                native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap();
                native::load_workbook(&path).unwrap()
            }
            1 => {
                json::import_any(&json::export_workbook(&wb, &[], 0).unwrap())
                    .unwrap()
                    .0
            }
            _ => {
                xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
                xlsx::import(&path).unwrap().0
            }
        };
        let id = loaded.active_sheet().tables()[0].id;
        assert_eq!(
            loaded.active_sheet().manual_hidden_rows(),
            [4, 5, 6].into(),
            "mode {mode}"
        );
        assert_eq!(loaded.table(id).unwrap().1.totals_row(), Some(5));
        assert_eq!(loaded.active_sheet().get_raw(4, 0), "7");
        assert_eq!(loaded.active_sheet().get_display(5, 0), "60");
        assert!(loaded
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .rows()
            .data_to_view(4)
            .is_none());
        let next = loaded
            .append_table_rows(id, 1, &[(5, 0, "11".into())])
            .unwrap();
        assert_eq!(loaded.active_sheet().get_display(6, 0), "60");
        loaded.apply_table_commit(&next, true).unwrap();
        assert_eq!(loaded.active_sheet().get_display(5, 0), "60");
        loaded.apply_table_commit(&next, false).unwrap();
        assert_eq!(loaded.active_sheet().manual_hidden_rows(), [4, 5, 6].into());
        let (shown, _) = loaded
            .prepare_table_row_visibility(loaded.active_sheet_id(), Default::default())
            .unwrap();
        assert_eq!(shown.active_sheet().get_display(6, 0), "78");
    }
}

#[test]
fn every_native_workbook_writer_and_full_json_preserve_manual_hides_without_totals() {
    let wb = book();
    let fingerprint = native::compute_semantic_fingerprint(&wb);
    assert!(fingerprint.starts_with("v5:"));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("hidden.sheet");
    for mode in 0..3 {
        match mode {
            0 => native::save_workbook(&wb, &path).unwrap(),
            1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap(),
            _ => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap(),
        }
        let layout = native::SheetLayout {
            col_widths: Default::default(),
            row_heights: Default::default(),
            hidden_rows: [(0, [1].into())].into(),
            hidden_cols: Default::default(),
        };
        native::save_layout(&path, &layout).unwrap();
        let loaded = native::load_workbook(&path).unwrap();
        check(&loaded);
        assert_eq!(native::compute_semantic_fingerprint(&loaded), fingerprint);
        assert!(native::load_layout(&path).hidden_rows[&0].contains(&1));
    }
    native::save(wb.active_sheet(), &path).unwrap();
    let loaded = native::load(&path).unwrap();
    assert_eq!(loaded.manual_hidden_rows(), [1].into());
    assert_eq!(loaded.get_display(0, 2), "40");
    for content in [
        json::export_workbook(&wb, &[], 0).unwrap(),
        json::export_full(wb.active_sheet()).unwrap(),
    ] {
        let (loaded, layouts, _) = json::import_any(&content).unwrap();
        check(&loaded);
        assert_eq!(layouts[0].hidden_rows, [1].into());
    }
    let (shown, _) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), Default::default())
        .unwrap();
    assert!(native::compute_semantic_fingerprint(&shown).starts_with("v2:"));
}

#[test]
fn xlsx_manual_hides_import_independently_and_survive_headless_export() {
    let mut excel = rust_xlsxwriter::Workbook::new();
    let sheet = excel.add_worksheet();
    for (row, value) in [10.0, 20.0, 30.0].iter().enumerate() {
        sheet.write_number(row as u32, 0, *value).unwrap();
    }
    sheet.set_row_hidden(1).unwrap();
    sheet
        .write_formula(
            0,
            2,
            rust_xlsxwriter::Formula::new("=SUBTOTAL(109,A1:A3)").set_result("40"),
        )
        .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("independent.xlsx");
    excel.save(&path).unwrap();
    let (loaded, _) = xlsx::import(&path).unwrap();
    check(&loaded);
    let (bytes, _) =
        xlsx::export_to_buffer_with_order(&loaded, None, xlsx::ExportOrder::Stored).unwrap();
    let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    let row = xml
        .split("<row ")
        .find(|r| r.starts_with("r=\"2\""))
        .unwrap()
        .split('>')
        .next()
        .unwrap();
    assert!(row.contains("hidden=\"1\""), "{row}");
    std::fs::write(&path, bytes).unwrap();
    check(&xlsx::import(&path).unwrap().0);
}

#[test]
fn sorted_export_cannot_silently_move_records_away_from_manual_flags() {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "30", "10", "20"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let table = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 0,
            },
            "Data",
        )
        .unwrap()
        .table_id();
    let mut spec = TableViewSpec::new(table);
    spec.sort = Some(TableSort {
        column: wb.table(table).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    let (wb, _) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [1].into())
        .unwrap();
    assert!(
        xlsx::export_to_buffer_with_order(&wb, None, xlsx::ExportOrder::Sorted)
            .unwrap_err()
            .contains("stored-order")
    );
    let (bytes, _) =
        xlsx::export_to_buffer_with_order(&wb, None, xlsx::ExportOrder::Stored).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("stored.xlsx");
    std::fs::write(&path, bytes).unwrap();
    let (loaded, _) = xlsx::import(&path).unwrap();
    assert_eq!(loaded.active_sheet().manual_hidden_rows(), [1].into());
    assert_eq!(loaded.active_sheet().get_raw(1, 0), "30");
}

#[test]
fn legacy_native_and_json_layout_fields_hydrate_calculation_state() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("legacy-layout.sheet");
    let hidden = book();
    let (visible, _) = hidden
        .prepare_table_row_visibility(hidden.active_sheet_id(), Default::default())
        .unwrap();
    native::save_workbook(&visible, &path).unwrap();
    let mut layout = native::SheetLayout {
        col_widths: Default::default(),
        row_heights: Default::default(),
        hidden_rows: Default::default(),
        hidden_cols: Default::default(),
    };
    layout.hidden_rows.insert(0, [1].into());
    native::save_layout(&path, &layout).unwrap();
    check(&native::load_workbook(&path).unwrap());
    let text = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Sheet1","hidden_rows":[1],"cells":[{"row":0,"col":0,"value":10},{"row":1,"col":0,"value":20},{"row":2,"col":0,"value":30},{"row":0,"col":2,"formula":"=SUBTOTAL(109,A1:A3)"}]}]}"#;
    check(&json::import_any(text).unwrap().0);
}

#[test]
fn single_sheet_native_load_keeps_formulas_forward_spills_and_unknown_function_caches() {
    use visigrid_engine::formula::eval::Value;
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 3, "=SEQUENCE(A3/10)");
    wb.set_cell_value_tracked(0, 0, 5, "=CUSTOM_VALUE()");
    wb.active_sheet().cache_computed(0, 5, Value::Number(123.0));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("single.sheet");
    native::save(wb.active_sheet(), &path).unwrap();
    let sheet = native::load(&path).unwrap();
    assert_eq!(sheet.get_raw(0, 2), "=SUBTOTAL(109,A1:A3)");
    assert_eq!(sheet.get_display(0, 2), "40");
    assert_eq!(sheet.get_display(2, 3), "3");
    assert_eq!(sheet.get_raw(0, 5), "=CUSTOM_VALUE()");
    assert_eq!(sheet.get_display(0, 5), "123");
}
