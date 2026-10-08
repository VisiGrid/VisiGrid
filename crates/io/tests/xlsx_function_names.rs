//! #89: Excel needs `_xlfn.` on newer functions however they were typed, and
//! `_xlpm.` on LET and LAMBDA names, or it shows #NAME? / offers a repair.
use std::io::Read;
use visigrid_engine::workbook::Workbook;
use visigrid_io::xlsx;

fn sheet_xml(wb: &Workbook) -> String {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("names.xlsx");
    xlsx::export(wb, &file, None).unwrap();
    let mut zip = zip::ZipArchive::new(std::fs::File::open(&file).unwrap()).unwrap();
    let mut out = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut out).unwrap();
    out
}

#[test]
fn newer_functions_and_let_lambda_names_are_written_as_excel_stores_them() {
    let mut wb = Workbook::new();
    for (r, value) in ["a", "b", "c"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 7, value);
        wb.set_cell_value_tracked(0, r, 8, &(r + 1).to_string());
    }
    wb.set_cell_value_tracked(0, 0, 0, "=xlookup(\"c\",H1:H3,I1:I3)");
    wb.set_cell_value_tracked(0, 1, 0, "=LET(x,5,y,7,x*y)");
    wb.set_cell_value_tracked(0, 2, 0, "=LAMBDA(a,b,a+b)(2,3)");
    wb.set_cell_value_tracked(0, 3, 0, "=let(f,lambda(n,n*2),f(3))");
    wb.set_cell_value_tracked(0, 4, 0, "=SUM(A1:A3)");
    assert_eq!(wb.active_sheet().get_display(1, 0), "35");

    let xml = sheet_xml(&wb);
    // Dynamic functions are written as single-cell array formulas.
    for expected in [
        ">_xlfn.XLOOKUP(\"c\",H1:H3,I1:I3)</f>",
        "<f>_xlfn.LET(_xlpm.x,5,_xlpm.y,7,_xlpm.x*_xlpm.y)</f>",
        ">_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(2,3)</f>",
        ">_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2),_xlpm.f(3))</f>",
        "<f>SUM(A1:A3)</f>",
    ] {
        assert!(xml.contains(expected), "missing {expected} in {xml}");
    }
}


#[test]
fn prefixed_formulas_import_and_calculate() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=LET(x,5,y,7,x*y)");
    for (r, (key, value)) in [("1", "10"), ("2", "20")].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 7, key);
        wb.set_cell_value_tracked(0, r, 8, value);
    }
    wb.set_cell_value_tracked(0, 1, 0, "=xlookup(2,H1:H2,I1:I2)");
    wb.set_cell_value_tracked(0, 2, 0, "=let(f,lambda(n,n*2),f(3))");
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("names.xlsx");
    xlsx::export(&wb, &file, None).unwrap();
    let (back, _) = xlsx::import(&file).unwrap();
    let sheet = back.active_sheet();
    assert_eq!(sheet.get_raw(0, 0), "=LET(x,5,y,7,x*y)");
    assert_eq!(sheet.get_display(0, 0), "35");
    assert_eq!(sheet.get_raw(1, 0), "=XLOOKUP(2,H1:H2,I1:I2)");
    assert_eq!(sheet.get_display(1, 0), "20");
    assert_eq!(sheet.get_raw(2, 0), "=LET(f,LAMBDA(n,n*2),f(3))");
    assert_eq!(sheet.get_display(2, 0), "6");
}
