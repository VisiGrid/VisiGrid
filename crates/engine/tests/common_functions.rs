// Functions added so the engine can be the web app's formula authority:
// LARGE, SMALL, RANK, RANK.EQ, MINIFS, MAXIFS, LOOKUP, MODE, MODE.SNGL,
// PERCENTILE, PERCENTILE.INC, QUARTILE, QUARTILE.INC, WEEKNUM, XOR, NA,
// HYPERLINK, SORTBY, TEXTSPLIT, LET, LAMBDA.
//
// Expected values are Excel's, including its error cases.

use visigrid_engine::workbook::Workbook;

fn cell(a1: &str) -> (usize, usize) {
    let split = a1.find(|c: char| c.is_ascii_digit()).unwrap();
    let (letters, digits) = a1.split_at(split);
    let col = letters.bytes().fold(0, |n, b| n * 26 + (b - b'A' + 1) as usize) - 1;
    (digits.parse::<usize>().unwrap() - 1, col)
}

/// A workbook holding `cells` (A1 notation, raw input as typed).
fn book(cells: &[(&str, &str)]) -> Workbook {
    let mut wb = Workbook::new();
    for (at, raw) in cells {
        let (r, c) = cell(at);
        wb.set_cell_value_tracked(0, r, c, raw);
    }
    wb
}

fn show(wb: &Workbook, at: &str) -> String {
    let (r, c) = cell(at);
    wb.active_sheet().get_display(r, c)
}

/// Evaluate `formula` in Z1 of a workbook holding `cells`.
fn eval(cells: &[(&str, &str)], formula: &str) -> String {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    show(&wb, "Z1")
}

/// Evaluate `formula` in Z1 and return its computed number.
fn num(cells: &[(&str, &str)], formula: &str) -> f64 {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    match wb.active_sheet().get_computed_value(0, 25) {
        visigrid_engine::formula::eval::Value::Number(n) => n,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

/// Evaluate an array formula in T1 and read the spilled block, row by row.
fn spill(cells: &[(&str, &str)], formula: &str, rows: usize, cols: usize) -> Vec<Vec<String>> {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 19, formula);
    (0..rows)
        .map(|r| (0..cols).map(|c| wb.active_sheet().get_display(r, 19 + c)).collect())
        .collect()
}

const NUMS: &[(&str, &str)] = &[("A1", "3"), ("A2", "5"), ("A3", "3"), ("A4", "5"), ("A5", "4"), ("A6", "text"), ("A7", "")];

// --- LARGE / SMALL ------------------------------------------------------------

#[test]
fn large_and_small_pick_by_rank_and_skip_text() {
    assert_eq!(eval(NUMS, "=LARGE(A1:A7,1)"), "5");
    assert_eq!(eval(NUMS, "=LARGE(A1:A7,3)"), "4");
    assert_eq!(eval(NUMS, "=SMALL(A1:A7,1)"), "3");
    assert_eq!(eval(NUMS, "=SMALL(A1:A7,5)"), "5");
    // A fractional k rounds up, as in Excel.
    assert_eq!(eval(NUMS, "=SMALL(A1:A7,1.2)"), "3");
    assert_eq!(eval(NUMS, "=LARGE(A1:A7,2.5)"), "4");
}

#[test]
fn large_and_small_errors() {
    assert_eq!(eval(NUMS, "=LARGE(A1:A7,0)"), "#NUM!");
    assert_eq!(eval(NUMS, "=SMALL(A1:A7,6)"), "#NUM!");
    assert_eq!(eval(&[("A1", "")], "=LARGE(A1:A3,1)"), "#NUM!");
    assert_eq!(eval(&[("A1", "1"), ("A2", "=1/0")], "=LARGE(A1:A2,1)"), "#DIV/0!");
}

#[test]
fn large_with_an_array_of_k_spills() {
    let got = spill(NUMS, "=LARGE(A1:A7,SEQUENCE(1,3))", 1, 3);
    assert_eq!(got, vec![vec!["5", "5", "4"]]);
}

// --- RANK / RANK.EQ -----------------------------------------------------------

#[test]
fn rank_descending_ascending_and_ties() {
    let data = &[("A1", "7"), ("A2", "3.5"), ("A3", "3.5"), ("A4", "1"), ("A5", "2")];
    assert_eq!(eval(data, "=RANK(A2,A1:A5)"), "2"); // Excel docs: RANK(A3,A2:A6,1)=3
    assert_eq!(eval(data, "=RANK(7,A1:A5)"), "1");
    assert_eq!(eval(data, "=RANK(A4,A1:A5,1)"), "1");
    assert_eq!(eval(data, "=RANK.EQ(3.5,A1:A5,1)"), "3");
    assert_eq!(eval(data, "=RANK.EQ(2,A1:A5)"), "4");
}

#[test]
fn rank_of_a_missing_number_is_na() {
    assert_eq!(eval(&[("A1", "1"), ("A2", "2")], "=RANK(5,A1:A2)"), "#N/A");
    assert_eq!(eval(&[("A1", "1")], "=RANK(\"x\",A1:A1)"), "#VALUE!");
}

// --- MINIFS / MAXIFS ----------------------------------------------------------

const SALES: &[(&str, &str)] = &[
    ("A1", "89"), ("B1", "1"), ("C1", "a"),
    ("A2", "93"), ("B2", "2"), ("C2", "b"),
    ("A3", "96"), ("B3", "2"), ("C3", "a"),
    ("A4", "85"), ("B4", "3"), ("C4", "b"),
    ("A5", "91"), ("B5", "1"), ("C5", "b"),
    ("A6", "88"), ("B6", "1"), ("C6", "a"),
];

#[test]
fn minifs_and_maxifs_follow_sumifs_criteria() {
    // Excel docs: MINIFS(A2:A7,B2:B7,1) = 88
    assert_eq!(eval(SALES, "=MINIFS(A1:A6,B1:B6,1)"), "88");
    assert_eq!(eval(SALES, "=MAXIFS(A1:A6,B1:B6,1)"), "91");
    assert_eq!(eval(SALES, "=MAXIFS(A1:A6,B1:B6,\">1\",C1:C6,\"a\")"), "96");
    assert_eq!(eval(SALES, "=MINIFS(A1:A6,C1:C6,\"b\")"), "85");
}

#[test]
fn minifs_with_no_match_is_zero_and_shape_mismatch_is_value() {
    assert_eq!(eval(SALES, "=MAXIFS(A1:A6,B1:B6,9)"), "0");
    assert_eq!(eval(SALES, "=MINIFS(A1:A6,B1:B5,1)"), "#VALUE!");
    assert_eq!(eval(&[("A1", "=1/0"), ("B1", "1")], "=MAXIFS(A1:A1,B1:B1,1)"), "#DIV/0!");
}

// --- LOOKUP -------------------------------------------------------------------

const BANDS: &[(&str, &str)] = &[
    ("A1", "4.14"), ("B1", "red"),
    ("A2", "4.19"), ("B2", "orange"),
    ("A3", "5.17"), ("B3", "yellow"),
    ("A4", "5.77"), ("B4", "green"),
    ("A5", "6.39"), ("B5", "blue"),
];

#[test]
fn lookup_vector_form_finds_the_largest_value_not_above() {
    // Excel docs examples.
    assert_eq!(eval(BANDS, "=LOOKUP(4.19,A1:A5,B1:B5)"), "orange");
    assert_eq!(eval(BANDS, "=LOOKUP(5.75,A1:A5,B1:B5)"), "yellow");
    assert_eq!(eval(BANDS, "=LOOKUP(7.66,A1:A5,B1:B5)"), "blue");
    assert_eq!(eval(BANDS, "=LOOKUP(0,A1:A5,B1:B5)"), "#N/A");
}

#[test]
fn lookup_array_form_uses_first_and_last_row_or_column() {
    // Tall: search the first column, return the last.
    assert_eq!(eval(BANDS, "=LOOKUP(5.2,A1:B5)"), "yellow");
    // Wide: search the first row, return the last.
    let wide = &[("A1", "1"), ("B1", "2"), ("C1", "3"), ("A2", "x"), ("B2", "y"), ("C2", "z")];
    assert_eq!(eval(wide, "=LOOKUP(2.5,A1:C2)"), "y");
    // Text compares case-insensitively.
    let words = &[("A1", "apple"), ("A2", "banana"), ("A3", "cherry"), ("B1", "1"), ("B2", "2"), ("B3", "3")];
    assert_eq!(eval(words, "=LOOKUP(\"BZ\",A1:A3,B1:B3)"), "2");
}

#[test]
fn lookup_last_match_idiom() {
    // LOOKUP(2, 1/(cond), range): the last row meeting cond.
    let data = &[("A1", "x"), ("A2", "y"), ("A3", "x"), ("A4", "z"), ("B1", "10"), ("B2", "20"), ("B3", "30"), ("B4", "40")];
    assert_eq!(eval(data, "=LOOKUP(2,1/(A1:A4=\"x\"),B1:B4)"), "30");
}

// --- MODE ---------------------------------------------------------------------

#[test]
fn mode_returns_the_most_frequent_first_on_ties() {
    let data = &[("A1", "5.6"), ("A2", "4"), ("A3", "4"), ("A4", "3"), ("A5", "2"), ("A6", "4")];
    assert_eq!(eval(data, "=MODE(A1:A6)"), "4"); // Excel docs
    assert_eq!(eval(&[("A1", "1"), ("A2", "2"), ("A3", "2"), ("A4", "1")], "=MODE.SNGL(A1:A4)"), "1");
    assert_eq!(eval(&[("A1", "1"), ("A2", "2")], "=MODE(A1:A2)"), "#N/A");
    assert_eq!(eval(&[], "=MODE(1,2,3,3)"), "3");
}

// --- PERCENTILE / QUARTILE ----------------------------------------------------

const QDATA: &[(&str, &str)] = &[("A1", "1"), ("A2", "2"), ("A3", "4"), ("A4", "7"), ("A5", "8"), ("A6", "9"), ("A7", "10"), ("A8", "12")];

#[test]
fn percentile_interpolates_inclusively() {
    assert_eq!(num(&[("A1", "1"), ("A2", "3"), ("A3", "2"), ("A4", "4")], "=PERCENTILE(A1:A4,0.3)"), 1.9); // Excel docs
    assert_eq!(eval(QDATA, "=PERCENTILE.INC(A1:A8,0)"), "1");
    assert_eq!(eval(QDATA, "=PERCENTILE.INC(A1:A8,1)"), "12");
    assert_eq!(eval(QDATA, "=PERCENTILE(A1:A8,1.1)"), "#NUM!");
    assert_eq!(eval(&[("A1", "x")], "=PERCENTILE(A1:A1,0.5)"), "#NUM!");
}

#[test]
fn quartile_is_percentile_by_quarters() {
    assert_eq!(num(QDATA, "=QUARTILE(A1:A8,1)"), 3.5); // Excel docs
    assert_eq!(num(QDATA, "=QUARTILE.INC(A1:A8,2)"), 7.5);
    assert_eq!(eval(QDATA, "=QUARTILE(A1:A8,4)"), "12");
    assert_eq!(num(QDATA, "=QUARTILE(A1:A8,2.9)"), 7.5); // truncated
    assert_eq!(eval(QDATA, "=QUARTILE(A1:A8,5)"), "#NUM!");
}

// --- WEEKNUM ------------------------------------------------------------------

#[test]
fn weeknum_systems_and_types() {
    // Excel docs: 2012-03-09 is week 10 (Sunday start) and 11 (Monday start).
    assert_eq!(eval(&[("A1", "=DATE(2012,3,9)")], "=WEEKNUM(A1)"), "10");
    assert_eq!(eval(&[("A1", "=DATE(2012,3,9)")], "=WEEKNUM(A1,2)"), "11");
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,1,1))"), "1");
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,1,7))"), "2"); // Sunday starts week 2
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,1,7),2)"), "1"); // Monday system
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,12,31))"), "53");
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,1,6),16)"), "2"); // Saturday start
}

#[test]
fn weeknum_iso_crosses_years() {
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2021,1,1),21)"), "53"); // belongs to 2020's week 53
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,12,30),21)"), "1"); // 2025's week 1
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2026,6,15),21)"), "25");
    assert_eq!(eval(&[], "=WEEKNUM(DATE(2024,1,1),3)"), "#NUM!");
    assert_eq!(eval(&[], "=WEEKNUM(-1)"), "#NUM!");
}

// --- XOR / NA -----------------------------------------------------------------

#[test]
fn xor_counts_trues_and_skips_text_in_references() {
    assert_eq!(eval(&[], "=XOR(3>0,2<9)"), "FALSE"); // Excel docs
    assert_eq!(eval(&[], "=XOR(3>12,4>6)"), "FALSE");
    assert_eq!(eval(&[], "=XOR(TRUE,FALSE,FALSE)"), "TRUE");
    assert_eq!(eval(&[("A1", "TRUE"), ("A2", "words"), ("A3", ""), ("A4", "1")], "=XOR(A1:A4)"), "FALSE");
    assert_eq!(eval(&[("A1", "TRUE"), ("A2", "words")], "=XOR(A1:A2)"), "TRUE");
}

#[test]
fn xor_errors() {
    assert_eq!(eval(&[("A1", "words")], "=XOR(A1:A1)"), "#VALUE!");
    assert_eq!(eval(&[], "=XOR(\"maybe\")"), "#VALUE!");
    assert_eq!(eval(&[("A1", "=1/0")], "=XOR(TRUE,A1)"), "#DIV/0!");
}

#[test]
fn na_is_the_na_error() {
    assert_eq!(eval(&[], "=NA()"), "#N/A");
    assert_eq!(eval(&[], "=ISNA(NA())"), "TRUE");
    assert_eq!(eval(&[], "=IFNA(NA(),\"fallback\")"), "fallback");
}

// --- HYPERLINK ----------------------------------------------------------------

#[test]
fn hyperlink_shows_the_friendly_name_or_the_link() {
    assert_eq!(eval(&[], "=HYPERLINK(\"https://visigrid.app\",\"VisiGrid\")"), "VisiGrid");
    assert_eq!(eval(&[], "=HYPERLINK(\"https://visigrid.app\")"), "https://visigrid.app");
    assert_eq!(eval(&[("A1", "42")], "=HYPERLINK(\"#Sheet1!A1\",A1)"), "42");
    assert_eq!(eval(&[], "=HYPERLINK(1/0)"), "#DIV/0!");
}

// --- SORTBY -------------------------------------------------------------------

#[test]
fn sortby_sorts_rows_by_keys_and_spills() {
    let data = &[("A1", "Tom"), ("A2", "Fred"), ("A3", "Amy"), ("A4", "Sal"), ("B1", "52"), ("B2", "65"), ("B3", "22"), ("B4", "73")];
    let got = spill(data, "=SORTBY(A1:A4,B1:B4)", 4, 1);
    assert_eq!(got, vec![vec!["Amy"], vec!["Tom"], vec!["Fred"], vec!["Sal"]]);
    let got = spill(data, "=SORTBY(A1:B4,B1:B4,-1)", 2, 2);
    assert_eq!(got, vec![vec!["Sal", "73"], vec!["Fred", "65"]]);
}

#[test]
fn sortby_multiple_keys_and_stability() {
    let data = &[
        ("A1", "a"), ("B1", "x"), ("C1", "2"),
        ("A2", "b"), ("B2", "y"), ("C2", "1"),
        ("A3", "c"), ("B3", "x"), ("C3", "1"),
        ("A4", "d"), ("B4", "y"), ("C4", "1"),
    ];
    let got = spill(data, "=SORTBY(A1:A4,B1:B4,1,C1:C4,-1)", 4, 1);
    assert_eq!(got, vec![vec!["a"], vec!["c"], vec!["b"], vec!["d"]]);
}

#[test]
fn sortby_columns_and_shape_errors() {
    let data = &[("A1", "c"), ("B1", "a"), ("C1", "b"), ("A2", "3"), ("B2", "1"), ("C2", "2")];
    let got = spill(data, "=SORTBY(A1:C1,A2:C2)", 1, 3);
    assert_eq!(got, vec![vec!["a", "b", "c"]]);
    assert_eq!(eval(data, "=SORTBY(A1:C1,A2:B2)"), "#VALUE!");
    assert_eq!(eval(data, "=SORTBY(A1:C2,A2:C2,2)"), "#VALUE!");
}

// --- TEXTSPLIT ----------------------------------------------------------------

#[test]
fn textsplit_columns_rows_and_padding() {
    assert_eq!(spill(&[], "=TEXTSPLIT(\"a,b,c\",\",\")", 1, 3), vec![vec!["a", "b", "c"]]);
    assert_eq!(
        spill(&[], "=TEXTSPLIT(\"1,2;3\",\",\",\";\")", 2, 2),
        vec![vec!["1", "2"], vec!["3", "#N/A"]]
    );
    assert_eq!(
        spill(&[], "=TEXTSPLIT(\"1,2;3\",\",\",\";\",FALSE,0,\"-\")", 2, 2),
        vec![vec!["1", "2"], vec!["3", "-"]]
    );
    assert_eq!(spill(&[], "=TEXTSPLIT(\"a;b\",,\";\")", 2, 1), vec![vec!["a"], vec!["b"]]);
}

#[test]
fn textsplit_multiple_delimiters_ignore_empty_and_case() {
    assert_eq!(spill(&[], "=TEXTSPLIT(\"a-b_c\",{\"-\",\"_\"})", 1, 3), vec![vec!["a", "b", "c"]]);
    assert_eq!(spill(&[], "=TEXTSPLIT(\"a,,b\",\",\",,TRUE)", 1, 2), vec![vec!["a", "b"]]);
    assert_eq!(spill(&[], "=TEXTSPLIT(\"a,,b\",\",\")", 1, 3), vec![vec!["a", "", "b"]]);
    assert_eq!(spill(&[], "=TEXTSPLIT(\"oneXtwoxthree\",\"x\",,,1)", 1, 3), vec![vec!["one", "two", "three"]]);
    assert_eq!(eval(&[], "=TEXTSPLIT(\"abc\",\",\")"), "abc");
    assert_eq!(eval(&[], "=TEXTSPLIT(\"abc\",\"\")"), "#VALUE!");
}

// --- LET ----------------------------------------------------------------------

#[test]
fn let_names_values_and_references() {
    assert_eq!(eval(&[], "=LET(x,1,x+1)"), "2"); // Excel docs
    assert_eq!(eval(&[], "=LET(x,1,y,x*10,x+y)"), "11");
    assert_eq!(eval(&[("A1", "2"), ("A2", "3"), ("A3", "4")], "=LET(r,A1:A3,SUM(r)*2)"), "18");
    assert_eq!(eval(&[("A1", "a"), ("A2", "b"), ("A3", "a")], "=LET(r,A1:A3,COUNTIF(r,\"a\"))"), "2");
    assert_eq!(eval(&[], "=LET(total,SEQUENCE(4),SUM(total))"), "10");
    assert_eq!(eval(&[], "=LET(s,\"Hello\",s&\" world\")"), "Hello world");
}

#[test]
fn let_scoping_and_shadowing() {
    assert_eq!(eval(&[], "=LET(x,1,LET(x,2,x)+x)"), "3");
    assert_eq!(eval(&[], "=LET(x,5,y,LET(x,x+1,x*2),x+y)"), "17");
    // A LET name does not hide a built-in function of the same name.
    assert_eq!(eval(&[], "=LET(sum,5,SUM(sum,1))"), "6");
}

#[test]
fn let_errors_and_spills() {
    assert_eq!(eval(&[], "=LET(x,1/0,IFERROR(x,\"caught\"))"), "caught");
    assert!(eval(&[], "=LET(x,1,x,2)").starts_with("#VALUE!"), "no calculation after the pairs");
    assert!(eval(&[], "=LET(x,1)").starts_with("#VALUE!"));
    assert!(eval(&[], "=LET(1,1,2)").starts_with("#VALUE!"));
    assert_eq!(spill(&[], "=LET(n,3,SEQUENCE(n))", 3, 1), vec![vec!["1"], vec!["2"], vec!["3"]]);
}

#[test]
fn let_evaluates_each_value_once() {
    // RAND inside a LET value is evaluated once, so both uses agree.
    assert_eq!(eval(&[], "=LET(r,RAND(),r=r)"), "TRUE");
}

// --- LAMBDA -------------------------------------------------------------------

#[test]
fn lambda_called_in_place() {
    assert_eq!(eval(&[], "=LAMBDA(x,x+1)(2)"), "3");
    assert_eq!(eval(&[], "=LAMBDA(a,b,a*b)(6,7)"), "42");
    assert_eq!(eval(&[], "=LAMBDA(x,LAMBDA(y,x+y))(1)(2)"), "3");
    assert_eq!(eval(&[], "=LAMBDA(42)()"), "42");
}

#[test]
fn lambda_through_a_let_name() {
    assert_eq!(eval(&[], "=LET(f,LAMBDA(x,x*2),f(3))"), "6");
    assert_eq!(eval(&[], "=LET(f,LAMBDA(x,y,x-y),f(10,4)+f(1,1))"), "6");
    // Closes over names defined before it.
    assert_eq!(eval(&[], "=LET(k,100,f,LAMBDA(x,x+k),k,1,f(1)+k)"), "102");
    // A parameter shadows an outer name.
    assert_eq!(eval(&[], "=LET(x,1,f,LAMBDA(x,x*10),f(5)+x)"), "51");
    assert_eq!(eval(&[("A1", "4"), ("A2", "6")], "=LET(avg,LAMBDA(r,SUM(r)/COUNT(r)),avg(A1:A2))"), "5");
}

#[test]
fn lambda_errors() {
    assert_eq!(eval(&[], "=LAMBDA(x,x+1)"), "#CALC!");
    assert_eq!(eval(&[], "=LAMBDA(x,x+1)(1,2)"), "#VALUE!");
    assert_eq!(eval(&[], "=LET(f,LAMBDA(x,x),f(1,2))"), "#VALUE!");
}

#[test]
fn lambda_and_let_formulas_round_trip_through_a_structural_edit() {
    // Inserting a row above rewrites stored formulas; LET names and an
    // in-place LAMBDA call must print back as calls, not internals.
    let mut wb = book(&[("A2", "5"), ("B2", "=LET(x,A2,x*2)"), ("C2", "=LAMBDA(v,v+1)(A2)")]);
    wb.structural_edit(0, visigrid_engine::structural::Axis::Row, 0, 1, false).unwrap();
    assert_eq!(show(&wb, "B3"), "10");
    assert_eq!(show(&wb, "C3"), "6");
    let raw = |at: &str| {
        let (r, c) = cell(at);
        wb.active_sheet().get_raw(r, c)
    };
    assert_eq!(raw("B3"), "=LET(X, A3, X*2)");
    assert_eq!(raw("C3"), "=LAMBDA(V, V+1)(A3)");
}

#[test]
fn let_and_lambda_follow_their_inputs_on_edit() {
    let mut wb = book(&[("A1", "2"), ("B1", "=LET(f,LAMBDA(x,x*x),f(A1)+1)")]);
    assert_eq!(show(&wb, "B1"), "5");
    wb.set_cell_value_tracked(0, 0, 0, "3");
    assert_eq!(show(&wb, "B1"), "10");
}


// --- Array constants -----------------------------------------------------------

#[test]
fn array_constants_evaluate_spill_and_feed_functions() {
    assert_eq!(eval(&[], "=SUM({1,2,3})"), "6");
    assert_eq!(eval(&[], "=SUM({1,2;3,4})"), "10");
    assert_eq!(spill(&[], "={1,2;3,4}", 2, 2), vec![vec!["1", "2"], vec!["3", "4"]]);
    assert_eq!(spill(&[], "={\"a\",TRUE,-1.5}", 1, 3), vec![vec!["a", "TRUE", "-1.5"]]);
    assert_eq!(spill(&[], "={1,2}*10", 1, 2), vec![vec!["10", "20"]]);
    assert_eq!(eval(&[], "=LARGE({4,9,2},2)"), "4");
}

#[test]
fn array_constants_print_back_and_reject_bad_shapes() {
    let mut wb = book(&[("A2", "1"), ("B2", "=SUM({1,2;3,4})+A2")]);
    wb.structural_edit(0, visigrid_engine::structural::Axis::Row, 0, 1, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(2, 1), "=SUM({1,2;3,4})+A3");
    assert_eq!(show(&wb, "B3"), "11");
    assert!(eval(&[], "={1,2;3}").starts_with('#'), "ragged rows are an error");
    assert!(eval(&[], "={A1,2}").starts_with('#'), "references are not constants");
}
