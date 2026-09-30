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
fn value_functions_spill_and_others_still_refuse() {
    // ABS and IF apply per element (#43)...
    let mut wb = table();
    formula(&mut wb, 0, 6, "=ABS(D2:D5-B2:B5)");
    let got: Vec<String> = (0..4).map(|r| show(&wb, r, 6)).collect();
    assert_eq!(got, ["5", "10", "10", "0"]);
    formula(&mut wb, 0, 7, "=IF(B2:B5>15,\"big\",\"small\")");
    let got: Vec<String> = (0..4).map(|r| show(&wb, r, 7)).collect();
    assert_eq!(got, ["small", "big", "big", "big"]);
    // ...but a function outside the per-element table still refuses an array
    // rather than silently reading its first cell.
    let got = formula(&mut wb, 10, 0, "=DATE(B2:B5,1,1)");
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

#[test]
fn sort_unique_and_transpose_transform_a_computed_array() {
    // Each used to return a computed argument untouched: SORT(UNIQUE(x)) came
    // back unsorted and UNIQUE(FILTER(...)) kept its duplicates, silently.
    let mut wb = Workbook::new();
    for (r, v) in ["West", "East", "West", "North", "East"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, v);
    }
    formula(&mut wb, 0, 2, "=SORT(UNIQUE(A1:A5))");
    let got: Vec<String> = (0..3).map(|r| show(&wb, r, 2)).collect();
    assert_eq!(got, ["East", "North", "West"]);
    assert_eq!(show(&wb, 3, 2), "", "three distinct values, no fourth");

    formula(&mut wb, 0, 3, "=UNIQUE(FILTER(A1:A5,A1:A5<>\"North\"))");
    let got: Vec<String> = (0..2).map(|r| show(&wb, r, 3)).collect();
    assert_eq!(got, ["West", "East"]);
    assert_eq!(show(&wb, 2, 3), "");

    formula(&mut wb, 10, 0, "=TRANSPOSE(SORT(UNIQUE(A1:A5),1,-1))");
    let got: Vec<String> = (0..3).map(|c| show(&wb, 10, c)).collect();
    assert_eq!(got, ["West", "North", "East"]);

    assert_eq!(formula(&mut wb, 12, 0, "=COUNTA(UNIQUE(A1:A5))"), "3");
}

#[test]
fn scalar_functions_over_ranges_43() {
    // The silent wrong answers from #43.
    let mut wb = table();
    formula(&mut wb, 0, 6, "=LEN(A2:A5)");
    let got: Vec<String> = (0..4).map(|r| show(&wb, r, 6)).collect();
    assert_eq!(got, ["3", "2", "2", "2"]);
    assert_eq!(formula(&mut wb, 10, 0, "=SUMPRODUCT(LEN(A2:A5))"), "9");
    assert_eq!(formula(&mut wb, 11, 0, "=SUMPRODUCT(--ISNUMBER(B2:B5))"), "4");
    assert_eq!(formula(&mut wb, 12, 0, "=SUMPRODUCT(--ISTEXT(A1:D1))"), "4");
    assert_eq!(formula(&mut wb, 13, 0, "=SUMPRODUCT(--(LEFT(A2:A5,1)=\"C\"))"), "1");
    assert_eq!(formula(&mut wb, 14, 0, "=SUM(ROUND(B2:B5/3,0))"), "28");
    assert_eq!(formula(&mut wb, 15, 0, "=SUMPRODUCT(--REGEXTEST(A2:A5,\"^[AB]\"))"), "2");
}

#[test]
fn errors_pass_through_value_functions() {
    // LEN of an error used to measure the error's text.
    let mut wb = table();
    wb.set_cell_value_tracked(0, 5, 0, "=1/0");
    assert_eq!(formula(&mut wb, 10, 0, "=LEN(A6)"), "#DIV/0!");
    assert_eq!(formula(&mut wb, 11, 0, "=UPPER(A6)"), "#DIV/0!");
    // ...while the functions that inspect errors still see them.
    assert_eq!(formula(&mut wb, 12, 0, "=ISERROR(A6)"), "TRUE");
    assert_eq!(formula(&mut wb, 13, 0, "=ISNUMBER(A6)"), "FALSE");
    assert_eq!(formula(&mut wb, 14, 0, "=IFERROR(A6,\"none\")"), "none");
    // Per element, too.
    formula(&mut wb, 0, 6, "=ISERROR(B2:B5/(D2:D5-D2:D5))");
    assert_eq!(show(&wb, 0, 6), "TRUE");
    formula(&mut wb, 0, 7, "=IFERROR(B2:B5/(B2:B5-10),\"skip\")");
    let got: Vec<String> = (0..2).map(|r| show(&wb, r, 7)).collect();
    assert_eq!(got, ["skip", "1.5"]);
}

#[test]
fn lifted_if_only_fails_on_the_branch_it_takes() {
    // Rows where qty is 30 would divide by zero in the TRUE branch; they take
    // the FALSE branch, so the error never surfaces.
    let mut wb = table();
    formula(&mut wb, 0, 6, "=IF(B2:B5<>30,100/(B2:B5-30),\"n/a\")");
    let got: Vec<String> = (0..4).map(|r| show(&wb, r, 6)).collect();
    assert_eq!(got, ["-5", "n/a", "-10", "-20"]);
}
