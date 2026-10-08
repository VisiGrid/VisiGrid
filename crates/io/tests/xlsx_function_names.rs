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
fn calculated_column_and_totals_formulas_prefix_future_functions() {
    use visigrid_engine::table::{TableRange, TableTotal};
    let mut wb = Workbook::new();
    for (col, header) in ["Key", "Value", "Found"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, col, header);
    }
    wb.set_cell_value_tracked(0, 1, 0, "a");
    wb.set_cell_value_tracked(0, 1, 1, "10");
    wb.set_cell_value_tracked(0, 2, 0, "b");
    wb.set_cell_value_tracked(0, 2, 1, "20");
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange { start_row: 0, start_col: 0, end_row: 2, end_col: 2 },
            "Items",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(id, 2, 1, "=XLOOKUP([@Key],A2:A3,B2:B3)+LET(n,1,n)", true)
        .unwrap();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(
        id,
        2,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=ROWS(FILTER(A2:A3,A2:A3=A2))+ROWS(SORT(A2:A3))".into()),
            label: None,
        },
    )
    .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("table.xlsx");
    xlsx::export_with_order(&wb, &file, None, xlsx::ExportOrder::Stored).unwrap();
    let mut zip = zip::ZipArchive::new(std::fs::File::open(&file).unwrap()).unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut zip.by_name("xl/tables/table1.xml").unwrap(), &mut xml).unwrap();
    assert!(
        xml.contains("<calculatedColumnFormula>_xlfn.XLOOKUP([[#This Row],[Key]],A2:A3,B2:B3)+_xlfn.LET(_xlpm.n,1,_xlpm.n)</calculatedColumnFormula>"),
        "{xml}"
    );
    assert!(
        xml.contains("<totalsRowFormula>ROWS(_xlfn._xlws.FILTER(A2:A3,A2:A3=A2))+ROWS(_xlfn._xlws.SORT(A2:A3))</totalsRowFormula>"),
        "{xml}"
    );
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
