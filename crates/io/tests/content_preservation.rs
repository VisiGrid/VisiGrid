use visigrid_io::json::{export_workbook, import_any};

// Unknown content must remain attached to its original document, sheet, cell
// and format. Merely accepting JSON on input does not establish compatibility.
#[test]
fn unknown_fields_survive_canonical_round_trip() {
    let source = include_str!("fixtures/unknown-workbook-fields.json");
    let before: serde_json::Value = serde_json::from_str(source).unwrap();
    let (workbook, layouts, active) = import_any(source).unwrap();
    assert!(workbook.read_only_reason().is_some());
    let after: serde_json::Value =
        serde_json::from_str(&export_workbook(&workbook, &layouts, active).unwrap()).unwrap();
    for pointer in [
        "/future_document_feature",
        "/sheets/0/future_sheet_feature",
        "/sheets/0/cells/0/future_cell_feature",
        "/sheets/0/cells/1/fmt/future_format_feature",
    ] {
        assert_eq!(
            after.pointer(pointer),
            before.pointer(pointer),
            "lost {pointer}"
        );
    }
}

#[test]
fn unsupported_content_survives_native_storage_without_becoming_editable() {
    let source = include_str!("fixtures/unknown-workbook-fields.json");
    let (workbook, _, _) = import_any(source).unwrap();
    let file = tempfile::NamedTempFile::new().unwrap();
    visigrid_io::native::save_workbook(&workbook, file.path()).unwrap();
    // The source envelope supplements a normal preview, never an empty DB.
    let conn = rusqlite::Connection::open(file.path()).unwrap();
    let sheets: i64 = conn
        .query_row("SELECT count(*) FROM sheets", [], |r| r.get(0))
        .unwrap();
    let cells: i64 = conn
        .query_row("SELECT count(*) FROM cells", [], |r| r.get(0))
        .unwrap();
    assert_eq!(sheets, 2);
    assert!(cells > 0);
    drop(conn);

    let restored = visigrid_io::native::load_workbook(file.path()).unwrap();
    assert!(restored.read_only_reason().is_some());
    let output = export_workbook(&restored, &[], 0).unwrap();
    assert_eq!(
        serde_json::from_str::<serde_json::Value>(&output).unwrap(),
        serde_json::from_str::<serde_json::Value>(source).unwrap()
    );
}

#[test]
fn clearing_the_warning_does_not_allow_a_lossy_rewrite() {
    let source = include_str!("fixtures/unknown-workbook-fields.json");
    let (mut workbook, layouts, active) = import_any(source).unwrap();
    workbook.sheet_mut(0).unwrap().read_only_reason = None;
    assert!(export_workbook(&workbook, &layouts, active).is_err());
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), b"keep existing contents").unwrap();
    assert!(visigrid_io::native::save_workbook(&workbook, file.path()).is_err());
    assert_eq!(
        std::fs::read(file.path()).unwrap(),
        b"keep existing contents"
    );
}

#[test]
fn recognizable_future_version_keeps_original_without_recalculation() {
    let source = r#"{"format":"visigrid-json","version":99,"sheets":[{"name":"Future","cells":[{"row":0,"col":0,"formula":"=1+1","value":123}]}]}"#;
    let (workbook, layouts, active) = import_any(source).unwrap();
    assert!(workbook.read_only_reason().is_some());
    assert_eq!(
        export_workbook(&workbook, &layouts, active).unwrap(),
        source
    );
}
