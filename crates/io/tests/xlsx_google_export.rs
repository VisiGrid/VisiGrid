//! A Google Sheets .xlsx export freezes Sheets-only formulas as
//! `IFERROR(__xludf.DUMMYFUNCTION("<formula>"),<last value>)`. Import restores
//! them, so they recalculate instead of showing an error or an old value.
use rust_xlsxwriter::Formula;
use visigrid_io::xlsx;

fn google_export(path: &std::path::Path) {
    let mut wb = rust_xlsxwriter::Workbook::new();
    let ws = wb.add_worksheet();
    for (r, (name, amount)) in [("rent", 1200), ("food", 300), ("fun", 80)]
        .into_iter()
        .enumerate()
    {
        ws.write_string(r as u32, 0, name).unwrap();
        ws.write_number(r as u32, 1, amount).unwrap();
    }
    // FILTER's result spilled down D1:D2: Google writes the formula in D1 and
    // a COMPUTED_VALUE placeholder in D2, each with its cached value.
    ws.write_formula(
        0,
        3,
        Formula::new(r#"=IFERROR(__xludf.DUMMYFUNCTION("FILTER(A1:A3,B1:B3>100)"),"rent")"#)
            .set_result("rent"),
    )
    .unwrap();
    ws.write_formula(
        1,
        3,
        Formula::new(r#"=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),"food")"#)
            .set_result("food"),
    )
    .unwrap();
    ws.write_formula(
        0,
        4,
        Formula::new(r#"=IFERROR(__xludf.DUMMYFUNCTION("REGEXMATCH(A1,""^r"")"),TRUE)"#)
            .set_result("TRUE"),
    )
    .unwrap();
    // QUERY isn't supported yet: its result spilled down F1:F2.
    ws.write_formula(
        0,
        5,
        Formula::new(r#"=IFERROR(__xludf.DUMMYFUNCTION("QUERY(A1:B3,""select A"")"),"rent")"#)
            .set_result("rent"),
    )
    .unwrap();
    ws.write_formula(
        1,
        5,
        Formula::new(r#"=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),"food")"#)
            .set_result("food"),
    )
    .unwrap();
    // An ordinary formula is untouched.
    ws.write_formula(0, 6, Formula::new("=SUM(B1:B3)").set_result("1580"))
        .unwrap();
    wb.save(path).unwrap();
}

#[test]
fn google_sheets_formulas_come_back_live() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("google.xlsx");
    google_export(&path);

    let (mut wb, report) = xlsx::import(&path).unwrap();
    assert_eq!(report.google_formulas_restored, 2);
    let sheet = wb.active_sheet();
    assert_eq!(sheet.get_raw(0, 3), "=FILTER(A1:A3,B1:B3>100)");
    assert_eq!(sheet.get_display(0, 3), "rent");
    assert_eq!(sheet.get_display(1, 3), "food");
    assert!(
        sheet.is_spill_receiver(1, 3),
        "the placeholder became part of the live spill"
    );
    assert_eq!(sheet.get_raw(0, 4), r#"=REGEXMATCH(A1,"^r")"#);
    assert_eq!(sheet.get_display(0, 4), "TRUE");
    assert_eq!(sheet.get_display(0, 6), "1580");

    // Not supported yet: Google's last values stay (no wall of errors), the
    // cell's note names the formula, and the report lists it.
    assert_eq!(sheet.get_display(0, 5), "rent");
    assert_eq!(sheet.get_display(1, 5), "food");
    let note = &sheet.comment(0, 5).expect("a note on the frozen cell").text;
    assert!(note.contains(r#"=QUERY(A1:B3,"select A")"#), "{note}");
    assert_eq!(
        report.google_formulas_frozen,
        vec![(0, 0, 5, r#"=QUERY(A1:B3,"select A")"#.to_string())]
    );
    assert!(
        report.unsupported_functions.contains_key("QUERY"),
        "{:?}",
        report.unsupported_functions
    );

    // Excel would keep showing "rent" and "food"; here the formulas follow the data.
    wb.set_cell_value_tracked(0, 2, 1, "500");
    let sheet = wb.active_sheet();
    assert_eq!(sheet.get_display(2, 3), "fun");
    wb.set_cell_value_tracked(0, 0, 0, "mortgage");
    assert_eq!(wb.active_sheet().get_display(0, 3), "mortgage");
    assert_eq!(wb.active_sheet().get_display(0, 4), "FALSE");
}
