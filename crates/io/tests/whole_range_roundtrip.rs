use visigrid_engine::workbook::Workbook;

#[test]
fn whole_ranges_survive_xlsx_and_json_roundtrips() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "2");
    wb.set_cell_value_tracked(0, 1, 0, "3");
    wb.set_cell_value_tracked(0, 2, 2, "=SUM($A:$A)");
    wb.set_cell_value_tracked(0, 3, 2, "=SUM(1:2)");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("whole-ranges.xlsx");
    visigrid_io::xlsx::export(&wb, &path, None).unwrap();
    let (mut imported, _) = visigrid_io::xlsx::import(&path).unwrap();
    imported.rebuild_dep_graph();
    imported.recompute_full_ordered();
    for row in [2, 3] {
        assert_eq!(imported.active_sheet().get_display(row, 2), "5");
        assert_eq!(
            imported.active_sheet().get_raw(row, 2),
            wb.active_sheet().get_raw(row, 2)
        );
    }
    let json = visigrid_io::json::export_full(imported.active_sheet()).unwrap();
    let sheet = visigrid_io::json::import_full(&json).unwrap();
    let mut restored = Workbook::from_sheets(vec![sheet], 0);
    restored.rebuild_dep_graph();
    restored.recompute_full_ordered();
    restored.set_cell_value_tracked(0, 100, 0, "7");
    assert_eq!(restored.active_sheet().get_display(2, 2), "12");
    assert_eq!(restored.active_sheet().get_display(3, 2), "5");
}
