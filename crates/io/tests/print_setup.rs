use rusqlite::Connection;
use tempfile::tempdir;
use visigrid_engine::{
    print_setup::*,
    sheet::{Sheet, SheetId},
    workbook::Workbook,
};
use visigrid_io::{json, native};

fn setup() -> PrintSetup {
    PrintSetup {
        paper: PrintPaper::Legal,
        landscape: true,
        scale: PrintScale::Actual,
        gridlines: true,
        page_numbers: true,
        area: Some(PrintArea {
            start_row: 2,
            start_col: 1,
            end_row: 80,
            end_col: 8,
        }),
        repeat_rows: Some(PrintRows { start: 2, end: 4 }),
    }
}

#[test]
fn every_native_writer_and_full_json_preserve_sheet_setup_without_changing_fingerprint() {
    let dir = tempdir().unwrap();
    let mut wb = Workbook::new();
    wb.active_sheet_mut().set_value(0, 0, "report");
    let hash = native::compute_semantic_fingerprint(&wb);
    let revision = wb.revision();
    wb.set_print_setup(wb.active_sheet_id(), setup()).unwrap();
    assert!(wb.revision() > revision);
    assert_eq!(native::compute_semantic_fingerprint(&wb), hash);
    wb.add_sheet();
    for mode in 0..3 {
        let path = dir.path().join(format!("writer-{mode}.sheet"));
        match mode {
            0 => native::save_workbook(&wb, &path).unwrap(),
            1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap(),
            _ => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap(),
        }
        let mut loaded = native::load_workbook(&path).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().print_setup, setup());
        assert!(loaded.sheet(1).unwrap().print_setup.is_default());
        loaded.sheet_mut(0).unwrap().set_value(0, 1, "CLI edit");
        native::save_workbook_full(&loaded, &Default::default(), &[], &[], &path).unwrap();
        assert_eq!(
            native::load_workbook(&path)
                .unwrap()
                .sheet(0)
                .unwrap()
                .print_setup,
            setup()
        );
    }
    let raw = json::export_workbook(&wb, &[], 0).unwrap();
    let (loaded, _, _) = json::import_any(&raw).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().print_setup, setup());
    let raw = json::export_full(wb.sheet(0).unwrap()).unwrap();
    assert_eq!(json::import_full(&raw).unwrap().print_setup, setup());
    let path = dir.path().join("single.sheet");
    native::save(wb.sheet(0).unwrap(), &path).unwrap();
    assert_eq!(native::load(&path).unwrap().print_setup, setup());
}

#[test]
fn old_files_default_and_unknown_or_corrupt_setup_is_not_silently_discarded() {
    let dir = tempdir().unwrap();
    let path = dir.path().join("old.sheet");
    native::save_workbook(&Workbook::new(), &path).unwrap();
    let conn = Connection::open(&path).unwrap();
    conn.execute_batch("DROP TABLE sheet_print_setup; PRAGMA user_version = 9;")
        .unwrap();
    assert!(native::load_workbook(&path)
        .unwrap()
        .active_sheet()
        .print_setup
        .is_default());
    assert_eq!(native::sheet_schema_version(&path).unwrap(), 11);
    conn.execute("INSERT INTO sheet_print_setup VALUES (0,99,'{}')", [])
        .unwrap();
    assert!(native::load_workbook(&path)
        .unwrap_err()
        .contains("Unsupported print setup"));
    conn.execute(
        "UPDATE sheet_print_setup SET version=1, settings='broken'",
        [],
    )
    .unwrap();
    assert!(native::load_workbook(&path)
        .unwrap_err()
        .contains("Invalid print setup"));
    let raw = r#"{"format":"visigrid-json","version":1,"name":"Old","cells":[]}"#;
    assert!(json::import_full(raw).unwrap().print_setup.is_default());
}

#[test]
fn setup_tracks_sheet_identity_clones_and_structural_edits() {
    let mut sheet = Sheet::new(SheetId(40), 100, 20);
    sheet.print_setup = setup();
    let mut wb = Workbook::from_sheets(vec![sheet.clone()], 0);
    let copy = wb.add_sheet_clone_named(&sheet, "Copy").unwrap();
    assert_eq!(wb.sheet(copy).unwrap().print_setup, setup());
    wb.structural_edit(copy, visigrid_engine::structural::Axis::Row, 0, 3, false)
        .unwrap();
    assert_eq!(
        wb.sheet(copy).unwrap().print_setup.repeat_rows,
        Some(PrintRows { start: 5, end: 7 })
    );
    assert_eq!(wb.sheet(0).unwrap().print_setup, setup());
    let moved = wb.take_sheet(copy).unwrap();
    let id = moved.id;
    assert!(wb.restore_sheet(0, moved));
    assert_eq!(
        wb.sheet_by_id(id)
            .unwrap()
            .print_setup
            .area
            .unwrap()
            .start_row,
        5
    );
    wb.structural_edit(0, visigrid_engine::structural::Axis::Row, 5, 3, true)
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().print_setup.repeat_rows, None);
    let dir = tempdir().unwrap();
    let path = dir.path().join("reordered.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let loaded = native::load_workbook(&path).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().name, "Copy");
    assert_eq!(
        loaded.sheet(0).unwrap().print_setup,
        wb.sheet(0).unwrap().print_setup
    );
    assert_eq!(loaded.sheet(1).unwrap().print_setup, setup());
}
