//! xlsx import reads cells sparsely. calamine's dense `Range` over a sheet
//! touching XFD2 and A1048576 asked for 2^34 cells (512 GiB) and aborted the
//! process; IronCalc's INDIRECT.xlsx test workbook has exactly that shape.

use visigrid_engine::formula::eval::Value;
use visigrid_io::xlsx;

#[test]
fn opposite_corner_cells_import_without_a_dense_allocation() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("corners.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let sheet = excel.add_worksheet();
    sheet.write_number(1, 16_383, 7.0).unwrap(); // XFD2
    sheet.write_number(1_048_575, 0, 1.0).unwrap(); // A1048576
    sheet.write_formula(1_048_575, 1, "=A1048576+XFD2").unwrap(); // B1048576
    excel.save(&path).unwrap();

    let (wb, result) = xlsx::import(&path).unwrap();
    let sheet = wb.sheet(0).unwrap();
    assert_eq!(sheet.get_computed_value(1, 16_383), Value::Number(7.0));
    assert_eq!(sheet.get_computed_value(1_048_575, 0), Value::Number(1.0));
    assert_eq!(sheet.get_raw(1_048_575, 1), "=A1048576+XFD2");
    assert_eq!(sheet.get_computed_value(1_048_575, 1), Value::Number(8.0));
    assert_eq!(result.formulas_imported, 1);
    assert!(!result.truncated);
}

#[test]
fn cells_keep_their_absolute_positions_when_the_sheet_starts_away_from_a1() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("offset.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let sheet = excel.add_worksheet();
    sheet.write_string(2, 1, "header").unwrap(); // B3
    sheet.write_number(3, 1, 5.0).unwrap(); // B4
    sheet.write_formula(4, 3, "=B4*2").unwrap(); // D5
    excel.save(&path).unwrap();

    let (wb, result) = xlsx::import(&path).unwrap();
    let sheet = wb.sheet(0).unwrap();
    assert_eq!(sheet.get_display(2, 1), "header");
    assert_eq!(sheet.get_computed_value(3, 1), Value::Number(5.0));
    assert_eq!(sheet.get_computed_value(4, 3), Value::Number(10.0));
    assert_eq!(result.cells_imported, 3);
}
