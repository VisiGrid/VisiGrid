// QUERY (Google Sheets): the Google Visualization query language over a range
// or array. Expected results follow Google's QUERY and query language
// references.

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

fn book(cells: &[(&str, &str)]) -> Workbook {
    let mut wb = Workbook::new();
    for (at, raw) in cells {
        let (r, c) = cell(at);
        wb.set_cell_value_tracked(0, r, c, raw);
    }
    wb
}

/// The formula in T1 and the `rows` x `cols` block it spills, as displayed.
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

fn value(cells: &[(&str, &str)], formula: &str) -> Value {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 19, formula);
    wb.active_sheet().get_computed_value(0, 19)
}

fn error(cells: &[(&str, &str)], formula: &str) -> String {
    match value(cells, formula) {
        Value::Error(e) => e,
        other => panic!("{formula}: expected an error, got {other:?}"),
    }
}

fn rows(spilled: &[&[&str]]) -> Vec<Vec<String>> {
    spilled
        .iter()
        .map(|r| r.iter().map(|s| s.to_string()).collect())
        .collect()
}

/// Name, Region, Amount, Date, with a header row.
const SALES: &[(&str, &str)] = &[
    ("A1", "Name"),
    ("B1", "Region"),
    ("C1", "Amount"),
    ("D1", "Date"),
    ("A2", "Ann"),
    ("B2", "East"),
    ("C2", "100"),
    ("D2", "=DATE(2026,1,5)"),
    ("A3", "Bob"),
    ("B3", "West"),
    ("C3", "250"),
    ("D3", "=DATE(2026,2,10)"),
    ("A4", "Cara"),
    ("B4", "East"),
    ("C4", "50"),
    ("D4", "=DATE(2026,3,15)"),
    ("A5", "Dan"),
    ("B5", "West"),
    ("C5", "75"),
    ("D5", "=DATE(2026,1,20)"),
    ("A6", "Eve"),
    ("B6", "North"),
    ("C6", "300"),
    ("D6", "=DATE(2026,2,28)"),
];

#[test]
fn select_where_order_by_with_the_header_row() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select A, C where C > 90 order by C desc")"#,
            5,
            2
        ),
        rows(&[
            &["Name", "Amount"],
            &["Eve", "300"],
            &["Bob", "250"],
            &["Ann", "100"],
            &["", ""]
        ])
    );
    // No select: every column. Whole columns work too.
    assert_eq!(
        spill(SALES, r#"=QUERY(A:C,"where B = 'North'")"#, 2, 3),
        rows(&[&["Name", "Region", "Amount"], &["Eve", "North", "300"]])
    );
}

#[test]
fn group_by_with_aggregates_sorted_by_the_group_or_an_aggregate() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select B, sum(C) group by B order by sum(C) desc")"#,
            4,
            2
        ),
        rows(&[
            &["Region", "sum Amount"],
            &["West", "325"],
            &["North", "300"],
            &["East", "150"]
        ])
    );
    // Without order by, groups come out in group-by order.
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select B, count(A), avg(C), max(A) group by B")"#,
            4,
            4
        ),
        rows(&[
            &["Region", "count Name", "avg Amount", "max Name"],
            &["East", "2", "75", "Cara"],
            &["North", "1", "300", "Eve"],
            &["West", "2", "162.5", "Dan"],
        ])
    );
    // Aggregates with no group by: one row over everything.
    assert_eq!(
        spill(SALES, r#"=QUERY(A1:D6,"select sum(C), min(C)")"#, 2, 2),
        rows(&[&["sum Amount", "min Amount"], &["775", "50"]])
    );
}

#[test]
fn label_renames_columns() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select B, sum(C) group by B label B 'Area', sum(C) 'Total'")"#,
            2,
            2
        ),
        rows(&[&["Area", "Total"], &["East", "150"]])
    );
}

#[test]
fn limit_and_offset_apply_after_ordering() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select A order by C limit 2 offset 1")"#,
            4,
            1
        ),
        rows(&[&["Name"], &["Dan"], &["Ann"], &[""]])
    );
}

#[test]
fn where_with_date_literals() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select A where D >= date '2026-02-01' and D < date '2026-03-01'")"#,
            3,
            1
        ),
        rows(&[&["Name"], &["Bob"], &["Eve"]])
    );
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select A, month(D) where year(D) = 2026 and month(D) = 3")"#,
            2,
            1
        ),
        rows(&[&["Name"], &["Cara"]])
    );
}

#[test]
fn string_operators_are_case_sensitive() {
    let names = |q: &str, n: usize| -> Vec<String> {
        spill(
            SALES,
            &format!("=QUERY(A1:D6,\"select A where {q}\")"),
            n + 1,
            1,
        )
        .into_iter()
        .skip(1)
        .map(|r| r[0].clone())
        .collect()
    };
    assert_eq!(names("A contains 'a'", 2), ["Cara", "Dan"]);
    assert_eq!(names("A starts with 'C'", 1), ["Cara"]);
    assert_eq!(names("A ends with 'e'", 1), ["Eve"]);
    assert_eq!(names("A matches '[AE].*'", 2), ["Ann", "Eve"]);
    assert_eq!(names("A like 'B_b'", 1), ["Bob"]);
    assert_eq!(names("not (B = 'East' or B = 'West')", 1), ["Eve"]);
}

#[test]
fn arithmetic_columns() {
    assert_eq!(
        spill(
            SALES,
            r#"=QUERY(A1:D6,"select A, C * 2 - 10 where A = 'Ann'")"#,
            2,
            2
        ),
        rows(&[&["Name", "C*2-10"], &["Ann", "190"]])
    );
}

#[test]
fn pivot_makes_a_column_per_value() {
    assert_eq!(
        spill(SALES, r#"=QUERY(A1:D6,"select sum(C) pivot B")"#, 2, 3),
        rows(&[&["East", "North", "West"], &["150", "300", "325"]])
    );
    // With group by: a row per group; an empty pivot cell stays empty.
    let wb = spill(
        SALES,
        r#"=QUERY(A1:D6,"select A, sum(C) where C > 90 group by A pivot B")"#,
        4,
        4,
    );
    assert_eq!(wb[0], ["Name", "East", "North", "West"]);
    assert_eq!(wb[1], ["Ann", "100", "", ""]);
    assert_eq!(wb[2], ["Bob", "", "", "250"]);
    assert_eq!(wb[3], ["Eve", "", "300", ""]);
}

#[test]
fn headers_are_guessed_or_given() {
    // No header row: numbers throughout, so none is guessed.
    let plain = &[("F1", "3"), ("F2", "1"), ("F3", "2")];
    assert_eq!(
        spill(plain, r#"=QUERY(F1:F3,"select F order by F")"#, 3, 1),
        rows(&[&["1"], &["2"], &["3"]])
    );
    // All text: none guessed (as in Sheets); the argument says there is one.
    let words = &[("F1", "Fruit"), ("F2", "pear"), ("F3", "apple")];
    assert_eq!(
        spill(words, r#"=QUERY(F1:F3,"select F order by F")"#, 3, 1),
        rows(&[&["apple"], &["Fruit"], &["pear"]])
    );
    assert_eq!(
        spill(words, r#"=QUERY(F1:F3,"select F order by F",1)"#, 3, 1),
        rows(&[&["Fruit"], &["apple"], &["pear"]])
    );
}

#[test]
fn a_columns_minority_type_reads_as_null() {
    let mixed = &[("G1", "Val"), ("G2", "5"), ("G3", "x"), ("G4", "7")];
    assert_eq!(
        spill(
            mixed,
            r#"=QUERY(G1:G4,"select G where G is not null")"#,
            4,
            1
        ),
        rows(&[&["Val"], &["5"], &["7"], &[""]])
    );
    assert_eq!(
        spill(mixed, r#"=QUERY(G1:G4,"select count(G)")"#, 2, 1),
        rows(&[&["count Val"], &["2"]])
    );
}

#[test]
fn arrays_name_columns_col1_col2() {
    assert_eq!(
        spill(
            &[],
            r#"=QUERY({1,"a";2,"b";3,"c"},"select Col2 where Col1 >= 2")"#,
            2,
            1
        ),
        rows(&[&["b"], &["c"]])
    );
}

#[test]
fn empty_results_and_errors() {
    // Header row only when the data has one.
    assert_eq!(
        spill(SALES, r#"=QUERY(A1:D6,"select A where C > 1000")"#, 2, 1),
        rows(&[&["Name"], &[""]])
    );
    let plain = &[("F1", "3"), ("F2", "1")];
    assert!(error(plain, r#"=QUERY(F1:F2,"select F where F > 10")"#).starts_with("#N/A"));

    for bad in [
        "selec A",
        "select Z",
        "select A where",
        "select A format A 'x'",
        "select A, sum(C) group by B",
        "where sum(C) > 1",
        "order by A select A",
        "select A where A = 'unclosed",
    ] {
        let e = error(SALES, &format!("=QUERY(A1:D6,\"{bad}\")"));
        assert!(e.starts_with("#VALUE!"), "{bad}: {e}");
    }
}
