#[test]
fn imported_cycles_keep_formula_sources_and_can_be_repaired() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("cycles.xlsx");
    let mut source = rust_xlsxwriter::Workbook::new();
    let sheet = source.add_worksheet();
    sheet.write_formula(0, 0, "=B1+1").unwrap();
    sheet.write_formula(0, 1, "=A1+1").unwrap();
    source.save(&path).unwrap();
    let (mut wb, report) = visigrid_io::xlsx::import(&path).unwrap();
    assert_eq!(report.recalc_circular, 2);
    assert_eq!(report.recalc_errors, 0);
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 0), "=B1+1");
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 1), "=A1+1");
    assert!(report
        .recalc_error_examples
        .iter()
        .all(|e| e.kind == "circular" && e.formula.is_some()));
    wb.set_cell_value_tracked(0, 0, 1, "1");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "2");
}
