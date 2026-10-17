// SUMIF/COUNTIF-family criteria matched against a cell's real type.
//
// Expected values are Excel's, taken from Excel-authored workbooks (IronCalc's
// calc_tests COUNTIF, COUNTIFS, SUMIF_AVERAGE_IF and SUMIFS). The cell side
// used to be read as display text and parsed back, so the text "23" passed
// for a number, a blank passed for 0 and TRUE passed for the text "TRUE".
//
// Criteria are not lookups: a numeric-looking criterion still compares
// numerically, and 23 / "=23" still match the text "23", as in Excel. That
// is the deliberate difference from VLOOKUP's "types must agree".

use visigrid_engine::workbook::Workbook;

fn cell(a1: &str) -> (usize, usize) {
    let split = a1.find(|c: char| c.is_ascii_digit()).unwrap();
    let (letters, digits) = a1.split_at(split);
    let col = letters.bytes().fold(0, |n, b| n * 26 + (b - b'A' + 1) as usize) - 1;
    (digits.parse::<usize>().unwrap() - 1, col)
}

/// A workbook holding `cells`. A leading `'` writes the rest as literal text
/// ("'23" is the text 23); anything else is entered as typed.
fn book(cells: &[(&str, &str)]) -> Workbook {
    let mut wb = Workbook::new();
    for (at, raw) in cells {
        let (r, c) = cell(at);
        match raw.strip_prefix('\'') {
            Some(text) => wb.set_cell_text_tracked(0, r, c, text),
            None => wb.set_cell_value_tracked(0, r, c, raw),
        };
    }
    wb
}

/// Evaluate `formula` in Z1 of a workbook holding `cells`, as displayed.
fn eval(cells: &[(&str, &str)], formula: &str) -> String {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    wb.active_sheet().get_display(0, 25)
}

/// COUNTIF.xlsx Rows!A2:A8 / G2:G9: a number, the same digits as text,
/// logicals, a word.
const MIXED: &[(&str, &str)] = &[
    ("A1", "23"), ("A2", "'23"), ("A3", "TRUE"), ("A4", "1"),
    ("A5", "-23"), ("A6", "hey"), ("A7", "FALSE"), ("A8", "123"),
];

#[test]
fn equality_matches_number_and_numeric_text_but_not_equal_excludes_only_the_number() {
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A7,23)"), "2");
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A7,\"=23\")"), "2");
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A7,\"23\")"), "2");
    // "<>23" still counts the text "23" (Excel 6 of 7).
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A7,\"<>23\")"), "6");
    assert_eq!(eval(MIXED, "=COUNTIFS(A1:A7,\"<>23\")"), "6");
}

#[test]
fn ordering_with_a_number_sees_only_number_cells() {
    // Numbers: 23, 1, -23, 123. Text "23" and the logicals never qualify.
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A8,\"<=23\")"), "3");
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A8,\">=23\")"), "2");
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A8,\">1\")"), "2");
    assert_eq!(eval(MIXED, "=COUNTIF(A1:A8,\"<0\")"), "1");
}

#[test]
fn blanks_never_satisfy_a_numeric_criterion() {
    // SUMIF_AVERAGE_IF.xlsx E5: the blank in B10 is not "<5".
    let cells = &[("B1", "3"), ("B2", "4"), ("B3", "5"), ("B4", "6"), ("B5", "-3"), ("B6", "12")];
    assert_eq!(eval(cells, "=COUNTIF(B1:B9,\"<5\")"), "3");
    assert_eq!(eval(cells, "=SUMIF(B1:B9,\"<5\")"), "4");
    assert_eq!(eval(cells, "=AVERAGEIF(B1:B9,\"<5\")*3"), "4");
    assert_eq!(eval(cells, "=COUNTIF(B1:B9,0)"), "0");
    assert_eq!(eval(cells, "=COUNTIFS(B1:B9,\">-4\",B1:B9,\"<89\")"), "6");
}

#[test]
fn a_criterion_read_from_an_empty_cell_is_zero() {
    // COUNTIF.xlsx AI2/AA6, COUNTIFS.xlsx L10: an empty criteria cell matches
    // zeros, not blanks.
    let cells = &[("A1", "TRUE"), ("A2", "1"), ("A3", "0"), ("A5", "FALSE")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A8,C1)"), "1");
    assert_eq!(eval(cells, "=COUNTIFS(A1:A8,C1,A1:A8,C2)"), "1");
    let no_zero = &[("A1", "1"), ("A2", "7"), ("A3", "8")];
    assert_eq!(eval(no_zero, "=COUNTIFS(A1:A8,C1,A1:A8,C2)"), "0");
    assert_eq!(eval(no_zero, "=AVERAGEIF(A1:A8,C1)"), "#DIV/0!");
}

#[test]
fn empty_and_not_empty_criteria() {
    // COUNTIF.xlsx AA5/AB5: "" matches blanks and empty text; "<>" counts
    // every cell that is not blank, empty text included.
    let cells = &[("A1", "12"), ("A2", "=\"\""), ("A4", "0"), ("A5", "x")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A6,\"\")"), "3");
    assert_eq!(eval(cells, "=COUNTIF(A1:A6,\"<>\")"), "4");
}

#[test]
fn logical_criteria_match_only_logical_cells() {
    // COUNTIF.xlsx AG/AH: TRUE and "true" match logical cells, never 1 or 0.
    // A typed TRUE is stored as text here, so it counts as the logical; a
    // formula's TRUE is a real logical.
    let cells = &[
        ("A1", "TRUE"), ("A2", "1"), ("A3", "0"), ("A4", "FALSE"),
        ("A6", "=1=1"), ("A8", "TRUE"), ("A9", "TRUE"), ("A10", "FALSE"),
    ];
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,TRUE)"), "4");
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,\"true\")"), "4");
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,\"FaLSe\")"), "2");
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,\"<>TRUE\")"), "9");
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,1)"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A13,0)"), "1");
}

#[test]
fn text_criteria_order_text_cells() {
    // COUNTIF.xlsx V4/V5 and COUNTIFS.xlsx C13: "<brown", ">brown", ">c".
    let names = &[
        ("A1", "Steven"), ("A2", "Albert"), ("A3", "Bernanrd"), ("A4", "Steven"),
        ("A5", "Mr. Brown"), ("A6", "Mrs. Brown"), ("A7", "Albert"), ("A8", "=NA()"),
        ("A10", "'#DIV/0!"), ("A11", "=1/0"), ("A12", "5"),
    ];
    assert_eq!(eval(names, "=COUNTIF(A1:A14,\"<brown\")"), "4");
    assert_eq!(eval(names, "=COUNTIF(A1:A14,\">brown\")"), "4");
    assert_eq!(eval(names, "=COUNTIFS(A1:A14,\">c\")"), "4");
    assert_eq!(eval(names, "=COUNTIF(A1:A14,\"<>*brown\")"), "12");
    assert_eq!(eval(names, "=COUNTIF(A1:A14,\"=*brown\")"), "2");
}

#[test]
fn wildcards_apply_to_text_cells_in_every_criteria_function() {
    // COUNTIFS.xlsx C2:D12 and SUMIF_AVERAGE_IF.xlsx J6: wildcards work in
    // COUNTIFS as in COUNTIF, and "12*" matches the text "1213", never the
    // number 1213.
    let colors = &[
        ("A1", "red"), ("A2", "BLUE"), ("A3", "yellow"), ("A4", "dark blue"),
        ("A5", "light blue"), ("A6", "light yellow"), ("A7", "dark yellow"),
        ("A8", "brown"), ("A9", "white"), ("A10", "Mr. White"), ("A11", "brown"),
        ("A12", "blue"), ("A13", "blue"), ("A14", "white dark"), ("A15", "red-ish"),
        ("A16", "blue-ish"), ("A17", "magenta"), ("A18", "magenta"), ("A19", "yellow"),
    ];
    for (criterion, expected) in [
        ("*blue", "5"), ("*o*", "6"), ("dark*", "2"), ("*-ish", "2"),
        ("*ll*", "4"), ("???-ish", "1"), ("~*", "0"),
    ] {
        assert_eq!(eval(colors, &format!("=COUNTIFS(A1:A19,\"{criterion}\")")), expected, "COUNTIFS {criterion}");
        assert_eq!(eval(colors, &format!("=COUNTIF(A1:A19,\"{criterion}\")")), expected, "COUNTIF {criterion}");
    }
    let ids = &[("A1", "'1213"), ("A2", "1213"), ("B1", "3"), ("B2", "4")];
    assert_eq!(eval(ids, "=AVERAGEIF(A1:A2,\"12*\",B1:B2)"), "3");
}

#[test]
fn error_criteria_match_error_cells_only() {
    // COUNTIF.xlsx AA2/AA4: NA() and the text "#N/A" both mean the error.
    let cells = &[
        ("A1", "=NA()"), ("A2", "12"), ("A3", "=1/0"), ("A4", "'#N/A"), ("A5", "0"),
    ];
    assert_eq!(eval(cells, "=COUNTIF(A1:A8,NA())"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A8,\"#N/A\")"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\"<>#N/A\")"), "4");
    assert_eq!(eval(cells, "=COUNTIF(A1:A8,1/0)"), "1");
}

#[test]
fn matched_errors_propagate_and_text_is_not_summed() {
    // SUMIFS.xlsx: sum_range cells that are text, logical or blank are
    // skipped; an error in a matched cell is the result.
    let cells = &[
        ("A1", "13"), ("A2", "17"), ("A3", "-23"), ("A4", "5"), ("A6", "=NA()"),
        ("A7", "'123"), ("A8", "TRUE"), ("A9", "=1/0"),
        ("B1", "red"), ("B2", "BLUE"), ("B3", "yellow"), ("B4", "dark blue"),
        ("B5", "dark yellow"), ("B6", "brown"), ("B7", "blue"), ("B8", "blue"),
        ("B9", "white"),
    ];
    assert_eq!(eval(cells, "=SUMIF(B1:B9,\"blue\",A1:A9)"), "17");
    assert_eq!(eval(cells, "=SUMIFS(A1:A9,B1:B9,\"blue\")"), "17");
    assert_eq!(eval(cells, "=AVERAGEIFS(A1:A9,B1:B9,\"blue\")"), "17");
    // "dark yellow" has a blank amount: averaged over the one number.
    assert_eq!(eval(cells, "=AVERAGEIF(B1:B9,\"dark*\",A1:A9)"), "5");
    assert_eq!(eval(cells, "=AVERAGEIFS(A1:A9,B1:B9,\"dark*\")"), "5");
    assert_eq!(eval(cells, "=SUMIF(B1:B9,\"*o*\",A1:A9)"), "#N/A");
    assert_eq!(eval(cells, "=SUMIFS(A1:A9,B1:B9,\"brown\")"), "#N/A");
    assert_eq!(eval(cells, "=AVERAGEIF(B1:B9,\"white\",A1:A9)"), "#DIV/0!");
    assert_eq!(eval(cells, "=MAXIFS(A1:A9,B1:B9,\">c\")"), "#DIV/0!");
    // SUMIF_AVERAGE_IF.xlsx F6: the criterion NA() matches the #N/A cell,
    // whose value is then the sum.
    assert_eq!(eval(&[("A1", "3"), ("A2", "=NA()")], "=SUMIF(A1:A2,NA())"), "#N/A");
    assert_eq!(eval(&[("A1", "3"), ("A2", "=NA()")], "=AVERAGEIF(A1:A2,NA())"), "#N/A");
    // An unmatched error is ignored.
    assert_eq!(eval(&[("A1", "3"), ("A2", "=NA()")], "=SUMIF(A1:A2,\">0\")"), "3");
}

#[test]
fn infinity_and_nan_spellings_are_text_not_numbers() {
    // arithmetic.xlsx Sheet2 D1/E1: "inf" and "1e400" are not numbers.
    let cells = &[("A1", "'1e400"), ("A2", "'inf"), ("A3", "'nan"), ("A4", "a"), ("A5", "23")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A7,\"inf\")"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A7,\"<1e400\")"), "0");
    assert_eq!(eval(cells, "=COUNTIF(A1:A7,\"nan\")"), "1");
}

#[test]
fn equals_alone_matches_only_truly_blank_cells() {
    // "=" is blank only; "" also takes empty text from a formula.
    let cells = &[("A1", "12"), ("A2", "=\"\""), ("A4", "0")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\"=\")"), "2");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\"\")"), "3");
}

#[test]
fn date_criteria_compare_serials_and_ignore_date_text() {
    let cells = &[
        ("A1", "=DATE(2024,1,15)"), ("A2", "=DATE(2023,6,1)"),
        ("A3", "'2024-01-15"), ("A4", "'2023-06-01"), ("A5", "'1/2/2024"),
    ];
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\">\"&DATE(2024,1,1))"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\">1/1/2024\")"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\"<=2024-01-01\")"), "1");
    // Text spelling a date never satisfies a date ordering.
    assert_eq!(eval(&cells[2..], "=COUNTIF(A3:A5,\">1/1/2024\")"), "0");
}

#[test]
fn criteria_and_sum_ranges_on_another_sheet() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet_named("Data").unwrap();
    for (r, (k, v)) in [("a", "10"), ("b", "20"), ("a", "'5"), ("a", "7")].iter().enumerate() {
        wb.set_cell_value_tracked(data, r, 0, k);
        match v.strip_prefix('\'') {
            Some(t) => wb.set_cell_text_tracked(data, r, 1, t),
            None => wb.set_cell_value_tracked(data, r, 1, v),
        };
    }
    wb.set_cell_value_tracked(0, 0, 0, "=SUMIF(Data!A1:A4,\"a\",Data!B1:B4)");
    wb.set_cell_value_tracked(0, 0, 1, "=COUNTIFS(Data!A1:A4,\"a\",Data!B1:B4,\">6\")");
    wb.set_cell_value_tracked(0, 0, 2, "=AVERAGEIFS(Data!B1:B4,Data!A1:A4,\"a\")*2");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 0), "17");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "2");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "17");
}

#[test]
fn whole_column_with_a_blank_criterion_matches_zeros() {
    let cells = &[("A1", "0"), ("A2", "5"), ("A3", "0"), ("B1", "10"), ("B2", "1"), ("B3", "7")];
    assert_eq!(eval(cells, "=COUNTIF(A:A,D1)"), "2");
    assert_eq!(eval(cells, "=SUMIF(A:A,D1,B:B)"), "17");
}

#[test]
fn numbers_stored_as_text_in_the_sum_range_are_skipped() {
    // Excel: SUMIF/AVERAGEIF add only numbers; the text "5" is not one.
    let cells = &[("A1", "x"), ("A2", "x"), ("B1", "'5"), ("B2", "3")];
    assert_eq!(eval(cells, "=SUMIF(A1:A2,\"x\",B1:B2)"), "3");
    assert_eq!(eval(cells, "=AVERAGEIF(A1:A2,\"x\",B1:B2)"), "3");
}

#[test]
fn logical_ordering_includes_typed_logicals() {
    // FALSE < TRUE; a typed TRUE (stored as text) orders like the formula's.
    let cells = &[("A1", "TRUE"), ("A2", "FALSE"), ("A3", "=1=1"), ("A4", "1"), ("A5", "'yes")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\">FALSE\")"), "2");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\"<TRUE\")"), "1");
    assert_eq!(eval(cells, "=COUNTIF(A1:A5,\">=FALSE\")"), "3");
}

#[test]
fn a_criterion_without_an_operator_is_read_as_written() {
    // " 5" is text, not 5; after an operator spaces are dropped.
    let cells = &[("A1", "5"), ("A2", "6")];
    assert_eq!(eval(cells, "=COUNTIF(A1:A2,\" 5\")"), "0");
    assert_eq!(eval(cells, "=COUNTIF(A1:A2,\">= 5\")"), "2");
}
