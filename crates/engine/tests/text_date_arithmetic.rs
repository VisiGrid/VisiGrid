//! A date typed as text still works in arithmetic, including with a two-digit
//! year, and text that is not a number or date is #VALUE!, not a sentence.
//! From a loan schedule where =IF(K10="",0,K10-I10) showed
//! "Cannot convert '10/8/26' to number" in the cell.

use visigrid_engine::workbook::Workbook;

fn display(wb: &Workbook, row: usize, col: usize) -> String {
    wb.active_sheet().get_display(row, col)
}

#[test]
fn two_digit_year_text_subtracts_like_a_date() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "10/8/26");
    wb.set_cell_value_tracked(0, 0, 1, "=DATE(2026,9,1)");
    wb.set_cell_value_tracked(0, 0, 2, "=IF(A1=\"\",0,A1-B1)");
    assert_eq!(display(&wb, 0, 2), "37");
}

#[test]
fn non_numeric_text_in_arithmetic_is_value_error() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "abc");
    wb.set_cell_value_tracked(0, 0, 1, "=A1-1");
    wb.set_cell_value_tracked(0, 0, 2, "=1*A1");
    assert_eq!(display(&wb, 0, 1), "#VALUE!");
    assert_eq!(display(&wb, 0, 2), "#VALUE!");
}
