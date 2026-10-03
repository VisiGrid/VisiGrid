use visigrid_engine::{table::TableRange, workbook::Workbook};
use visigrid_io::{json, native, xlsx};

fn fixture() -> Workbook {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "Amount");
    wb.set_cell_value_tracked(0, 1, 0, "42");
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0,
        },
        "Sales",
    )
    .unwrap();
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(Sales[Amount])");
    wb
}

#[test]
fn native_future_and_corrupt_tables_offer_non_destructive_read_only_recovery() {
    for (raw, expected) in [
        (r#"{"version":99,"unknown_shape":true}"#, "Upgrade VisiGrid"),
        ("{damaged", "corrupt"),
        (r#"{"version":2,"next_table_id":0,"sheets":[]}"#, "corrupt"),
    ] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("source.sheet");
        native::save_workbook(&fixture(), &path).unwrap();
        {
            let db = rusqlite::Connection::open(&path).unwrap();
            db.execute("UPDATE meta SET value=?1 WHERE key='tables'", [raw])
                .unwrap();
        }
        let before = std::fs::read(&path).unwrap();
        assert!(native::load_workbook(&path).unwrap_err().contains(expected));
        let (wb, issue) = native::load_workbook_for_recovery(&path).unwrap();
        assert!(issue.unwrap().to_string().contains(expected));
        assert!(wb.read_only_reason().unwrap().contains(expected));
        assert_eq!(wb.active_sheet().get_raw(1, 0), "42");
        assert_eq!(wb.active_sheet().get_raw(0, 2), "=SUM(Sales[Amount])");
        assert_eq!(wb.active_sheet().get_display(0, 2), "42");
        assert!(native::save_workbook(&wb, &path).is_err());
        let copy = dir.path().join("save-as.sheet");
        assert!(native::save_workbook(&wb, &copy).is_err());
        assert!(native::save_workbook_with_metadata(&wb, &Default::default(), &copy).is_err());
        assert!(native::save_workbook_full(&wb, &Default::default(), &[], &[], &copy).is_err());
        assert!(!copy.exists());
        assert!(native::save(wb.active_sheet(), &copy).is_err());
        assert!(json::export_workbook(&wb, &[], 0).is_err());
        assert!(json::export_full(wb.active_sheet()).is_err());
        assert!(xlsx::export_to_buffer_with_stored_fallback(&wb, None).is_err());
        for order in [xlsx::ExportOrder::Sorted, xlsx::ExportOrder::Stored] {
            assert!(xlsx::table_export_warnings_with_order(&wb, None, order).is_err());
            assert!(xlsx::export_with_order(&wb, &copy, None, order).is_err());
            assert!(xlsx::export_to_buffer_with_order(&wb, None, order).is_err());
            assert!(!copy.exists());
        }
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }
}

#[test]
fn json_recovery_distinguishes_future_format_and_corrupt_metadata() {
    for (catalog, expected) in [
        (
            serde_json::json!({"version":99,"future":[]}),
            "Upgrade VisiGrid",
        ),
        (serde_json::json!({"version":2,"sheets":"bad"}), "corrupt"),
    ] {
        let mut doc: serde_json::Value =
            serde_json::from_str(&json::export_workbook(&fixture(), &[], 0).unwrap()).unwrap();
        doc["table_catalog"] = catalog;
        let raw = doc.to_string();
        assert!(json::import_any(&raw).unwrap_err().contains(expected));
        let (wb, _, _) = json::import_any_for_recovery(&raw).unwrap();
        assert!(wb.read_only_reason().unwrap().contains(expected));
        assert_eq!(wb.active_sheet().get_display(0, 2), "42");
        assert!(json::export_workbook(&wb, &[], 0).is_err());
    }
}

#[test]
fn recovery_requires_readable_cells_and_leaves_valid_workbooks_editable() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("source.sheet");
    native::save_workbook(&fixture(), &path).unwrap();
    let (wb, issue) = native::load_workbook_for_recovery(&path).unwrap();
    assert!(issue.is_none());
    assert!(wb.read_only_reason().is_none());
    assert_eq!(wb.active_sheet().get_display(0, 2), "42");
    native::save_workbook(&wb, &dir.path().join("valid-copy.sheet")).unwrap();
    {
        let db = rusqlite::Connection::open(&path).unwrap();
        db.execute(
            "UPDATE meta SET value='{\"version\":99}' WHERE key='tables'",
            [],
        )
        .unwrap();
        db.execute("DROP TABLE cells", []).unwrap();
    }
    let before = std::fs::read(&path).unwrap();
    assert!(native::load_workbook_for_recovery(&path).is_err());
    assert_eq!(std::fs::read(&path).unwrap(), before);
}
