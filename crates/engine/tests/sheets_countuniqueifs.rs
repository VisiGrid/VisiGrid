// COUNTUNIQUEIFS(range, criteria_range1, criteria1, ...): Google Sheets'
// distinct count under COUNTIFS-style criteria. Expected values are Google
// Sheets', from its function reference.

use visigrid_engine::formula::eval::Value;
use visigrid_engine::workbook::Workbook;

/// Region in A, rep in B, amount in C.
const SALES: &[(&str, &str)] = &[
    ("A1", "West"), ("B1", "Ann"), ("C1", "100"),
    ("A2", "West"), ("B2", "Bob"), ("C2", "50"),
    ("A3", "East"), ("B3", "Ann"), ("C3", "70"),
    ("A4", "west"), ("B4", "ann"), ("C4", "20"),
    ("A5", "West"), ("B5", ""), ("C5", "90"),
    ("A6", "East"), ("B6", "Cy"), ("C6", "10"),
];

fn value(cells: &[(&str, &str)], formula: &str) -> Value {
    let mut wb = Workbook::new();
    for (at, raw) in cells {
        let col = (at.as_bytes()[0] - b'A') as usize;
        let row: usize = at[1..].parse().unwrap();
        wb.set_cell_value_tracked(0, row - 1, col, raw);
    }
    wb.set_cell_value_tracked(0, 0, 25, formula);
    wb.active_sheet().get_computed_value(0, 25)
}

fn num(cells: &[(&str, &str)], formula: &str) -> f64 {
    match value(cells, formula) {
        Value::Number(n) => n,
        other => panic!("{formula}: expected a number, got {other:?}"),
    }
}

#[test]
fn counts_distinct_values_in_matching_rows() {
    // West rows (any case): Ann, Bob, ann, blank. Ann and ann are one value.
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(B1:B6,A1:A6,"West")"#), 2.0);
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(B1:B6,A1:A6,"East")"#), 2.0);
    // Every criterion must hold.
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(B1:B6,A1:A6,"West",C1:C6,">=50")"#), 2.0);
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(B1:B6,A1:A6,"West",C1:C6,">60")"#), 1.0);
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(A1:A6,B1:B6,"A*")"#), 2.0);
    assert_eq!(num(SALES, r#"=COUNTUNIQUEIFS(B1:B6,A1:A6,"North")"#), 0.0);
}

#[test]
fn numbers_and_text_count_apart() {
    let cells = [("A1", "1"), ("A2", "1"), ("A3", "2"), ("A4", "TRUE"), ("B1", "x"), ("B2", "x"), ("B3", "x"), ("B4", "x")];
    // 1, 2 and TRUE.
    assert_eq!(num(&cells, r#"=COUNTUNIQUEIFS(A1:A4,B1:B4,"x")"#), 3.0);
}

#[test]
fn bad_arguments_are_errors() {
    for f in [
        r#"=COUNTUNIQUEIFS(B1:B6,A1:A5,"West")"#,
        "=COUNTUNIQUEIFS(B1:B6,A1:A6)",
        "=COUNTUNIQUEIFS(B1:B6)",
    ] {
        assert!(matches!(value(SALES, f), Value::Error(_)), "{f}");
    }
    // An error in a matched cell is the result.
    let cells = [("A1", "=1/0"), ("B1", "x")];
    assert!(matches!(value(&cells, r#"=COUNTUNIQUEIFS(A1:A1,B1:B1,"x")"#), Value::Error(_)));
}
