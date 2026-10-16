// Google Sheets functions, second batch: SORTN, FLATTEN, ARRAY_CONSTRAIN,
// JOIN, ISBETWEEN, ISEMAIL, ISURL, EPOCHTODATE, the TO_* conversions and the
// operator functions (ADD, MINUS, MULTIPLY, DIVIDE, EQ, NE, GT, GTE, LT, LTE,
// UMINUS, UPLUS, UNARY_PERCENT).
//
// Expected values are Google Sheets', from its function reference, including
// the error cases.

use visigrid_engine::formula::eval::Value;
use visigrid_engine::workbook::Workbook;

fn cell(a1: &str) -> (usize, usize) {
    let split = a1.find(|c: char| c.is_ascii_digit()).unwrap();
    let (letters, digits) = a1.split_at(split);
    let col = letters
        .bytes()
        .fold(0, |n, b| n * 26 + (b - b'A' + 1) as usize)
        - 1;
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

/// Evaluate `formula` in Z1 of a workbook holding `cells`, as computed.
fn value(cells: &[(&str, &str)], formula: &str) -> Value {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    wb.active_sheet().get_computed_value(0, 25)
}

fn num(cells: &[(&str, &str)], formula: &str) -> f64 {
    match value(cells, formula) {
        Value::Number(n) => n,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

fn text(cells: &[(&str, &str)], formula: &str) -> String {
    match value(cells, formula) {
        Value::Text(t) => t,
        other => panic!("{formula}: expected text, got {other:?}"),
    }
}

fn boolean(cells: &[(&str, &str)], formula: &str) -> bool {
    match value(cells, formula) {
        Value::Boolean(b) => b,
        other => panic!("{formula}: expected TRUE or FALSE, got {other:?}"),
    }
}

fn is_error(cells: &[(&str, &str)], formula: &str) -> bool {
    matches!(value(cells, formula), Value::Error(_))
}

/// Evaluate an array formula in T1 and read the spilled block, row by row.
fn spill(cells: &[(&str, &str)], formula: &str, rows: usize, cols: usize) -> Vec<Vec<String>> {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 19, formula);
    (0..rows)
        .map(|r| {
            (0..cols)
                .map(|c| wb.active_sheet().get_display(r, 19 + c))
                .collect()
        })
        .collect()
}

fn close(a: f64, b: f64) -> bool {
    (a - b).abs() < 1e-9
}

const SCORES: &[(&str, &str)] = &[
    ("A1", "Ann"),
    ("B1", "90"),
    ("A2", "Bob"),
    ("B2", "80"),
    ("A3", "Cat"),
    ("B3", "90"),
    ("A4", "Dan"),
    ("B4", "70"),
    ("A5", "Eve"),
    ("B5", "80"),
    ("A6", "Fay"),
    ("B6", "60"),
];

fn names(block: Vec<Vec<String>>) -> Vec<String> {
    block.into_iter().map(|row| row[0].clone()).collect()
}

// ---------------------------------------------------------------- SORTN

#[test]
fn sortn_ties_modes() {
    // Sorted by score, descending (stable): Ann 90, Cat 90, Bob 80, Eve 80, Dan 70, Fay 60.
    let sortn = |formula: &str, rows: usize| names(spill(SCORES, formula, rows, 2));
    assert_eq!(
        sortn("=SORTN(A1:B6,3,0,2,FALSE)", 4),
        ["Ann", "Cat", "Bob", ""]
    );
    // 1: plus the rows tied with the third (Eve ties Bob at 80).
    assert_eq!(
        sortn("=SORTN(A1:B6,3,1,2,FALSE)", 5),
        ["Ann", "Cat", "Bob", "Eve", ""]
    );
    // 2: duplicates (by the sort column) dropped first.
    assert_eq!(
        sortn("=SORTN(A1:B6,3,2,2,FALSE)", 4),
        ["Ann", "Bob", "Dan", ""]
    );
    // 3: the first two distinct scores, with every row holding them.
    assert_eq!(
        sortn("=SORTN(A1:B6,2,3,2,FALSE)", 5),
        ["Ann", "Cat", "Bob", "Eve", ""]
    );
}

#[test]
fn sortn_defaults_and_sort_columns() {
    // n defaults to 1, and with no sort column rows sort by every column ascending.
    assert_eq!(
        spill(SCORES, "=SORTN(A1:B6)", 2, 2),
        [["Ann", "90"], ["", ""]]
    );
    // A sort column may be a range beside the data, as tall as it.
    assert_eq!(
        names(spill(SCORES, "=SORTN(A1:A6,2,0,B1:B6,TRUE)", 2, 1)),
        ["Fay", "Dan"]
    );
    // Several sort columns: score descending, then name descending.
    assert_eq!(
        names(spill(SCORES, "=SORTN(A1:B6,2,0,2,FALSE,1,FALSE)", 2, 2)),
        ["Cat", "Ann"]
    );
    // n larger than the data returns every row.
    assert_eq!(
        names(spill(SCORES, "=SORTN(A1:B6,10,0,2,TRUE)", 7, 2)),
        ["Fay", "Dan", "Bob", "Eve", "Ann", "Cat", ""]
    );
}

#[test]
fn sortn_errors() {
    assert!(is_error(SCORES, "=SORTN(A1:B6,-1)"));
    assert!(is_error(SCORES, "=SORTN(A1:B6,2,4)"));
    assert!(is_error(SCORES, "=SORTN(A1:B6,2,0,3,TRUE)"));
    assert!(is_error(SCORES, "=SORTN(A1:B6,2,0,B1:B3,TRUE)"));
    assert!(is_error(SCORES, "=SORTN(A1:B6,0)"));
}

// ---------------------------------------------------------------- FLATTEN

#[test]
fn flatten_goes_row_by_row_through_each_argument() {
    let cells = [("A1", "x"), ("C1", "y")];
    assert_eq!(
        spill(&cells, "=FLATTEN({1,2;3,4},A1:C1)", 8, 1),
        [["1"], ["2"], ["3"], ["4"], ["x"], [""], ["y"], [""]]
    );
    assert_eq!(spill(&[], "=FLATTEN(5)", 2, 1), [["5"], [""]]);
}

// ---------------------------------------------------------------- ARRAY_CONSTRAIN

#[test]
fn array_constrain_keeps_the_top_left_corner() {
    let cells = [
        ("A1", "1"),
        ("B1", "2"),
        ("C1", "3"),
        ("A2", "4"),
        ("B2", "5"),
        ("C2", "6"),
        ("A3", "7"),
        ("B3", "8"),
        ("C3", "9"),
    ];
    assert_eq!(
        spill(&cells, "=ARRAY_CONSTRAIN(A1:C3,2,2)", 3, 3),
        [["1", "2", ""], ["4", "5", ""], ["", "", ""]]
    );
    // Asking for more than there is returns all of it, unpadded.
    assert_eq!(
        spill(&cells, "=ARRAY_CONSTRAIN(A1:B2,5,5)", 3, 3),
        [["1", "2", ""], ["4", "5", ""], ["", "", ""]]
    );
    assert!(is_error(&cells, "=ARRAY_CONSTRAIN(A1:B2,0,1)"));
    assert!(is_error(&cells, "=ARRAY_CONSTRAIN(A1:B2,1)"));
}

// ---------------------------------------------------------------- JOIN

#[test]
fn join_keeps_blanks_and_takes_values_or_lines() {
    let cells = [("A1", "a"), ("A3", "c"), ("B1", "1"), ("B2", "2")];
    assert_eq!(text(&cells, "=JOIN(\"-\",A1:A3)"), "a--c");
    assert_eq!(text(&cells, "=JOIN(\", \",1,2,\"x\")"), "1, 2, x");
    assert_eq!(text(&cells, "=JOIN(\",\",{1,2,3},B1:B2)"), "1,2,3,1,2");
    assert!(is_error(&cells, "=JOIN(\",\",A1:B3)"));
    assert!(is_error(&cells, "=JOIN(\",\",NA())"));
    assert!(is_error(&cells, "=JOIN(\",\")"));
}

// ---------------------------------------------------------------- ISBETWEEN

#[test]
fn isbetween_bounds() {
    assert!(boolean(&[], "=ISBETWEEN(5,1,10)"));
    assert!(boolean(&[], "=ISBETWEEN(10,1,10)"));
    assert!(!boolean(&[], "=ISBETWEEN(10,1,10,TRUE,FALSE)"));
    assert!(!boolean(&[], "=ISBETWEEN(1,1,10,FALSE)"));
    assert!(!boolean(&[], "=ISBETWEEN(11,1,10)"));
    assert!(boolean(&[], "=ISBETWEEN(\"b\",\"a\",\"c\")"));
    assert!(boolean(
        &[],
        "=ISBETWEEN(DATE(2026,5,1),DATE(2026,1,1),DATE(2026,12,31))"
    ));
    let cells = [
        ("A1", "1"),
        ("A2", "2"),
        ("A3", "3"),
        ("A4", "4"),
        ("A5", "5"),
    ];
    assert_eq!(num(&cells, "=SUMPRODUCT(--ISBETWEEN(A1:A5,2,4))"), 3.0);
    assert!(is_error(&[], "=ISBETWEEN(NA(),1,2)"));
    assert!(is_error(&[], "=ISBETWEEN(1,2)"));
}

// ---------------------------------------------------------------- ISEMAIL / ISURL

#[test]
fn isemail_and_isurl() {
    assert!(boolean(&[], "=ISEMAIL(\"info@example.com\")"));
    assert!(!boolean(&[], "=ISEMAIL(\"example.com\")"));
    assert!(!boolean(&[], "=ISEMAIL(42)"));
    assert!(!boolean(&[], "=ISEMAIL(NA())"));
    assert!(boolean(&[], "=ISURL(\"http://www.google.com\")"));
    assert!(boolean(&[], "=ISURL(\"www.google.com\")"));
    assert!(!boolean(&[], "=ISURL(\"google\")"));
    assert!(!boolean(&[], "=ISURL(NA())"));
    let cells = [("A1", "a@b.co"), ("A2", "nope"), ("A3", "c@d.org")];
    assert_eq!(num(&cells, "=SUMPRODUCT(--ISEMAIL(A1:A3))"), 2.0);
}

// ---------------------------------------------------------------- EPOCHTODATE

#[test]
fn epochtodate_units() {
    // Google's example: 2022-06-22 14:05:10 UTC.
    let expected = num(&[], "=DATE(2022,6,22)+TIME(14,5,10)");
    assert!(close(num(&[], "=EPOCHTODATE(1655906710)"), expected));
    assert!(close(num(&[], "=EPOCHTODATE(1655906710,1)"), expected));
    assert!(close(num(&[], "=EPOCHTODATE(1655906710000,2)"), expected));
    assert!(close(
        num(&[], "=EPOCHTODATE(1655906710000000,3)"),
        expected
    ));
    assert_eq!(num(&[], "=EPOCHTODATE(0)"), 25569.0);
    // Precision stops at the millisecond.
    assert!(close(
        num(&[], "=EPOCHTODATE(1.0009)"),
        25569.0 + 1000.0 / 86_400_000.0
    ));
    assert!(is_error(&[], "=EPOCHTODATE(100,4)"));
    assert!(is_error(&[], "=EPOCHTODATE(-1)"));
    assert!(is_error(&[], "=EPOCHTODATE(\"x\")"));
}

// ---------------------------------------------------------------- TO_*

#[test]
fn to_conversions_keep_the_value() {
    // A number stays a number (Sheets changes only its display format).
    assert_eq!(num(&[], "=TO_PERCENT(0.5)"), 0.5);
    assert_eq!(num(&[], "=TO_DOLLARS(1.5)"), 1.5);
    assert_eq!(num(&[], "=TO_DATE(44927)"), 44927.0);
    assert_eq!(num(&[], "=TO_PURE_NUMBER(0.25)"), 0.25);
    // Anything else comes back unchanged.
    assert_eq!(text(&[], "=TO_DATE(\"abc\")"), "abc");
    assert_eq!(text(&[], "=TO_PERCENT(\"abc\")"), "abc");
    assert_eq!(text(&[], "=TO_TEXT(24)"), "24");
    assert_eq!(text(&[], "=TO_TEXT(0.5)"), "0.5");
    assert_eq!(text(&[], "=TO_TEXT(\"x\")"), "x");
    assert!(is_error(&[], "=TO_TEXT(NA())"));
    let cells = [("A1", "1"), ("A2", "2")];
    assert_eq!(spill(&cells, "=TO_TEXT(A1:A2)", 2, 1), [["1"], ["2"]]);
}

// ---------------------------------------------------------------- operators

#[test]
fn operator_functions_match_their_operators() {
    assert_eq!(num(&[], "=ADD(2,3)"), 5.0);
    assert_eq!(num(&[], "=MINUS(5,2)"), 3.0);
    assert_eq!(num(&[], "=MULTIPLY(2,3)"), 6.0);
    assert_eq!(num(&[], "=DIVIDE(6,3)"), 2.0);
    assert_eq!(num(&[], "=ADD(\"2\",1)"), 3.0);
    assert!(boolean(&[], "=EQ(\"A\",\"a\")"));
    assert!(boolean(&[], "=NE(1,2)"));
    assert!(boolean(&[], "=GT(2,1)"));
    assert!(boolean(&[], "=GTE(2,2)"));
    assert!(boolean(&[], "=LT(1,2)"));
    assert!(!boolean(&[], "=LTE(3,2)"));
    assert_eq!(num(&[], "=UMINUS(4)"), -4.0);
    assert_eq!(num(&[], "=UNARY_PERCENT(50)"), 0.5);
    assert_eq!(text(&[], "=UPLUS(\"x\")"), "x");
    assert_eq!(num(&[], "=UPLUS(7)"), 7.0);
    let cells = [("A1", "1"), ("A2", "2"), ("A3", "3")];
    assert_eq!(spill(&cells, "=ADD(A1:A3,1)", 3, 1), [["2"], ["3"], ["4"]]);
}

#[test]
fn operator_function_errors() {
    assert!(is_error(&[], "=DIVIDE(1,0)"));
    assert!(is_error(&[], "=ADD(\"a\",1)"));
    assert!(is_error(&[], "=UMINUS(\"x\")"));
    assert!(is_error(&[], "=ADD(NA(),1)"));
    assert!(is_error(&[], "=EQ(1)"));
    assert!(is_error(&[], "=ADD(1,2,3)"));
}
