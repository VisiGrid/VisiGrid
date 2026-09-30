use rusqlite::Connection;
use tempfile::tempdir;
use visigrid_engine::{sheet::MergedRegion, workbook::Workbook};
use visigrid_io::native;

fn workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.add_sheet_named("Report").unwrap();
    for i in 0..2 {
        let sheet = wb.sheet_mut(i).unwrap();
        sheet.set_value(0, 0, &format!("Title {i}"));
        sheet
            .add_merge(MergedRegion::new(0, 0, i + 1, i + 2))
            .unwrap();
    }
    wb.set_active_sheet(1);
    wb
}

#[test]
fn every_native_writer_keeps_each_sheets_merges() {
    let wb = workbook();
    let dir = tempdir().unwrap();
    for mode in 0..3 {
        let path = dir.path().join(format!("merges-{mode}.sheet"));
        match mode {
            0 => native::save_workbook(&wb, &path).unwrap(),
            1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap(),
            _ => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap(),
        }
        let loaded = native::load_workbook(&path).unwrap();
        for i in 0..2 {
            assert_eq!(
                loaded.sheet(i).unwrap().merged_regions,
                wb.sheet(i).unwrap().merged_regions
            );
            assert_eq!(loaded.sheet(i).unwrap().get_raw(0, 0), format!("Title {i}"));
        }
        // A subsequent save must not bring old legacy merge rows back.
        let mut cleared = loaded;
        cleared.sheet_mut(1).unwrap().merged_regions.clear();
        native::save_workbook(&cleared, &path).unwrap();
        let reopened = native::load_workbook(&path).unwrap();
        assert!(reopened.sheet(1).unwrap().merged_regions.is_empty());
        assert_eq!(
            reopened.sheet(0).unwrap().merged_regions,
            wb.sheet(0).unwrap().merged_regions
        );
    }
    let path = dir.path().join("single.sheet");
    native::save(wb.sheet(1).unwrap(), &path).unwrap();
    assert_eq!(
        native::load(&path).unwrap().merged_regions,
        wb.sheet(1).unwrap().merged_regions
    );
    // Legacy v1 has no sheets table; its loader must not add the merges twice.
    Connection::open(&path)
        .unwrap()
        .execute_batch("DROP TABLE sheets;")
        .unwrap();
    assert_eq!(
        native::load_workbook(&path)
            .unwrap()
            .active_sheet()
            .merged_regions,
        wb.sheet(1).unwrap().merged_regions
    );
}

#[test]
fn legacy_active_sheet_merges_migrate_to_their_original_sheet() {
    let wb = workbook();
    let dir = tempdir().unwrap();
    let path = dir.path().join("legacy.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let conn = Connection::open(&path).unwrap();
    conn.execute_batch("DROP TABLE sheet_merged_regions; PRAGMA user_version=10;")
        .unwrap();
    drop(conn);
    let loaded = native::load_workbook(&path).unwrap();
    // Pre-v11 stored only the active sheet. Missing other-sheet merges cannot be recovered.
    assert!(loaded.sheet(0).unwrap().merged_regions.is_empty());
    assert_eq!(
        loaded.sheet(1).unwrap().merged_regions,
        wb.sheet(1).unwrap().merged_regions
    );
    assert_eq!(native::sheet_schema_version(&path).unwrap(), 11);
}
