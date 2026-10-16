//! A Google placeholder in B2 with one array formula above it (B1) and
//! another beside it (A2): the export doesn't say which array filled it. When
//! one of the two is live and the other frozen, import decides after
//! recalculating, by whether the live array actually spills there.
use rust_xlsxwriter::Formula;
use visigrid_io::xlsx;

const PLACEHOLDER: &str = r#"=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),"b2")"#;

fn export(path: &std::path::Path, b1: &str, a2: &str) {
    let mut wb = rust_xlsxwriter::Workbook::new();
    let ws = wb.add_worksheet();
    ws.write_formula(0, 1, Formula::new(b1).set_result("b1"))
        .unwrap();
    ws.write_formula(1, 0, Formula::new(a2).set_result("a2"))
        .unwrap();
    ws.write_formula(1, 1, Formula::new(PLACEHOLDER).set_result("b2"))
        .unwrap();
    wb.save(path).unwrap();
}

#[test]
fn a_live_array_beside_it_spills_into_the_corner() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("corner.xlsx");
    // B1: unsupported, one cell. A2: SPLIT, live, spills A2:B2.
    export(
        &path,
        r#"=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(""GOOG"")"),"b1")"#,
        r#"=IFERROR(__xludf.DUMMYFUNCTION("SPLIT(""x,y"","","")"),"x")"#,
    );
    let (wb, _) = xlsx::import(&path).unwrap();
    let sheet = wb.active_sheet();
    assert_eq!(sheet.get_display(1, 0), "x");
    assert_eq!(sheet.get_display(1, 1), "y", "SPLIT spilled into B2");
    assert!(sheet.is_spill_receiver(1, 1));
    assert_eq!(sheet.get_display(0, 1), "b1");
}

#[test]
fn a_frozen_array_above_it_gets_its_value_back() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("corner.xlsx");
    // B1: unsupported, spilled B1:B2. A2: REGEXMATCH, live, one cell.
    export(
        &path,
        r#"=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(""GOOG"",""all"")"),"b1")"#,
        r#"=IFERROR(__xludf.DUMMYFUNCTION("REGEXMATCH(""a"",""a"")"),TRUE)"#,
    );
    let (wb, _) = xlsx::import(&path).unwrap();
    let sheet = wb.active_sheet();
    assert_eq!(sheet.get_display(1, 0), "TRUE");
    assert_eq!(sheet.get_display(1, 1), "b2", "Google's value is back");
    assert!(!sheet.is_spill_receiver(1, 1));
}
