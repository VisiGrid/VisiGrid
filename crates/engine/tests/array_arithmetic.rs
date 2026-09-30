//! Operators over ranges: =B2:B5>15 is an array, and the functions that take
//! arrays — SUMPRODUCT, FILTER, SUM, COUNTA, ROWS, INDEX — consume it.
//! Until this existed every such formula was "#VALUE! Array arithmetic not
//! supported", which ruled out FILTER conditions and SUMPRODUCT(--(cond)).

use visigrid_engine::workbook::Workbook;

/// A small table: name, qty, tag, floor.
fn table() -> Workbook {
    let mut wb = Workbook::new();
    let rows = [
        ["name", "qty", "tag", "floor"],
        ["Ana", "10", "x", "5"],
        ["Bo", "30", "y", "40"],
        ["Cy", "20", "x", "10"],
        ["Di", "25", "z", "25"],
    ];
    for (r, row) in rows.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            wb.set_cell_value_tracked(0, r, c, v);
        }
    }
    wb
}

fn show(wb: &Workbook, r: usize, c: usize) -> String {
    wb.active_sheet().get_display(r, c)
}

fn formula(wb: &mut Workbook, r: usize, c: usize, f: &str) -> String {
    wb.set_cell_value_tracked(0, r, c, f);
    show(wb, r, c)
}

#[test]
fn conditional_counts_and_sums() {
    let mut wb = table();
    assert_eq!(formula(&mut wb, 10, 0, "=SUMPRODUCT(--(B2:B5>15))"), "3");
    assert_eq!(formula(&mut wb, 11, 0, "=SUMPRODUCT((C2:C5=\"x\")*B2:B5)"), "30");
    assert_eq!(formula(&mut wb, 12, 0, "=SUMPRODUCT(--(B2:B5>D2:D5))"), "2");
    assert_eq!(formula(&mut wb, 13, 0, "=SUMPRODUCT(B2:B5,--(C2:C5=\"x\"))"), "30");
    assert_eq!(formula(&mut wb, 14, 0, "=SUM(B2:B5*2)"), "170");
    assert_eq!(formula(&mut wb, 15, 0, "=SUM(--(C2:C5=\"x\"))"), "2");
    assert_eq!(formula(&mut wb, 16, 0, "=MAX(B2:B5-D2:D5)"), "10");
}

#[test]
fn filter_takes_a_condition() {
    let mut wb = table();
    formula(&mut wb, 0, 6, "=FILTER(A2:A5,B2:B5>15)");
    let got: Vec<String> = (0..3).map(|r| show(&wb, r, 6)).collect();
    assert_eq!(got, ["Bo", "Cy", "Di"]);
    // AND is *, OR is +
    formula(&mut wb, 0, 7, "=FILTER(A2:A5,(B2:B5>15)*(C2:C5=\"x\"))");
    assert_eq!(show(&wb, 0, 7), "Cy");
    formula(&mut wb, 0, 8, "=FILTER(A2:B5,(C2:C5=\"x\")+(C2:C5=\"z\"))");
    let got: Vec<String> = (0..3).map(|r| format!("{} {}", show(&wb, r, 8), show(&wb, r, 9))).collect();
    assert_eq!(got, ["Ana 10", "Cy 20", "Di 25"]);
    assert_eq!(formula(&mut wb, 10, 0, "=FILTER(A2:A5,B2:B5>99,\"none\")"), "none");
    assert_eq!(formula(&mut wb, 11, 0, "=SUM(FILTER(B2:B5,C2:C5=\"x\"))"), "30");
}

#[test]
fn functions_that_read_a_filter_result() {
    let mut wb = table();
    assert_eq!(formula(&mut wb, 10, 0, "=COUNTA(FILTER(A2:A5,B2:B5>15))"), "3");
    assert_eq!(formula(&mut wb, 11, 0, "=ROWS(FILTER(A2:A5,B2:B5>15))"), "3");
    assert_eq!(formula(&mut wb, 12, 0, "=COLUMNS(FILTER(A2:B5,B2:B5>15))"), "2");
    assert_eq!(formula(&mut wb, 13, 0, "=INDEX(FILTER(A2:A5,B2:B5>15),2)"), "Cy");
    assert_eq!(formula(&mut wb, 14, 0, "=INDEX(FILTER(A2:B5,B2:B5>15),3,2)"), "25");
    assert!(formula(&mut wb, 15, 0, "=INDEX(FILTER(A2:A5,B2:B5>15),9)").starts_with("#REF"));
}

#[test]
fn an_operator_over_a_range_spills() {
    let mut wb = table();
    formula(&mut wb, 0, 6, "=B2:B5>15");
    let got: Vec<String> = (0..4).map(|r| show(&wb, r, 6)).collect();
    assert_eq!(got, ["FALSE", "TRUE", "TRUE", "TRUE"]);
    formula(&mut wb, 0, 7, "=A2:A5&\"!\"");
    assert_eq!(show(&wb, 3, 7), "Di!");
}

#[test]
fn shapes_broadcast_and_mismatches_are_na() {
    let mut wb = table();
    // A column against a row fills a grid.
    formula(&mut wb, 0, 6, "=B2:B3*B2:C2");
    assert_eq!(show(&wb, 1, 6), "300"); // 30 * 10
    // Arrays of different lengths leave #N/A where they don't overlap.
    assert_eq!(formula(&mut wb, 10, 0, "=SUM(B2:B3*B2:B5)"), "#N/A");
}

#[test]
fn a_function_that_is_not_array_aware_errors_rather_than_guessing() {
    // Reading a multi-cell array as one value used to take its top-left cell.
    let mut wb = table();
    let got = formula(&mut wb, 10, 0, "=ABS(B2:B5*2)");
    assert!(got.starts_with("#"), "expected an error, got {got}");
    let got = formula(&mut wb, 11, 0, "=IF(B2:B5>15,1,0)");
    assert!(got.starts_with("#"), "expected an error, got {got}");
}

#[test]
fn array_formulas_recalculate_when_a_cell_in_the_range_changes() {
    let mut wb = table();
    formula(&mut wb, 10, 0, "=SUMPRODUCT(--(B2:B5>15))");
    formula(&mut wb, 11, 0, "=COUNTA(FILTER(A2:A5,B2:B5>15))");
    wb.set_cell_value_tracked(0, 1, 1, "99"); // Ana: 10 -> 99
    assert_eq!(show(&wb, 10, 0), "4");
    assert_eq!(show(&wb, 11, 0), "4");
}
