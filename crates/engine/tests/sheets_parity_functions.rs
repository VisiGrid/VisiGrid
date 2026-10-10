// Everyday functions added for Google Sheets parity (phase P2): RANK.AVG,
// CORREL, COUNTUNIQUE, SPLIT, CHAR, CODE, CLEAN, FIXED, DOLLAR, ISOWEEKNUM,
// YEARFRAC, MAP, REDUCE, SCAN, BYROW, BYCOL, ARRAYFORMULA.
//
// Expected values are Excel's (Google Sheets' for its own functions: SPLIT,
// COUNTUNIQUE, ARRAYFORMULA), including their error cases.

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

/// Evaluate `formula` in Z1 of a workbook holding `cells`, as displayed.
fn eval(cells: &[(&str, &str)], formula: &str) -> String {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    wb.active_sheet().get_display(0, 25)
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

/// Evaluate `formula` in Z1 and return its computed text.
fn text(cells: &[(&str, &str)], formula: &str) -> String {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    match wb.active_sheet().get_computed_value(0, 25) {
        visigrid_engine::formula::eval::Value::Text(t) => t,
        other => panic!("{formula}: expected text, got {other:?}"),
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

/// Whether `formula` evaluates to an error value of any kind. Wrong argument
/// counts report a message rather than an Excel error code, as elsewhere.
fn is_error(cells: &[(&str, &str)], formula: &str) -> bool {
    let mut wb = book(cells);
    wb.set_cell_value_tracked(0, 0, 25, formula);
    matches!(wb.active_sheet().get_computed_value(0, 25), visigrid_engine::formula::eval::Value::Error(_))
}

fn close(a: f64, b: f64) -> bool {
    (a - b).abs() < 1e-8
}

const NUMS: &[(&str, &str)] = &[("A1", "3"), ("A2", "5"), ("A3", "3"), ("A4", "5"), ("A5", "5"), ("A6", "text"), ("A7", "")];

// --- RANK.AVG -----------------------------------------------------------------

#[test]
fn rank_avg_averages_tied_ranks() {
    // 5, 5, 5 hold ranks 1-3 from the top, so each is 2; 3, 3 hold 4-5.
    assert_eq!(num(NUMS, "=RANK.AVG(5,A1:A7)"), 2.0);
    assert_eq!(num(NUMS, "=RANK.AVG(3,A1:A7)"), 4.5);
    assert_eq!(num(NUMS, "=RANK.AVG(3,A1:A7,1)"), 1.5);
    // Excel's documentation example: 94 in {89,88,92,101,94,97,95} is 4th.
    let doc = &[("B1", "89"), ("B2", "88"), ("B3", "92"), ("B4", "101"), ("B5", "94"), ("B6", "97"), ("B7", "95")];
    assert_eq!(num(doc, "=RANK.AVG(94,B1:B7)"), 4.0);
    // RANK.EQ is untouched: ties share the top rank.
    assert_eq!(num(NUMS, "=RANK.EQ(3,A1:A7)"), 4.0);
}

#[test]
fn rank_avg_errors() {
    assert_eq!(eval(NUMS, "=RANK.AVG(4,A1:A7)"), "#N/A");
    assert_eq!(eval(NUMS, "=RANK.AVG(\"x\",A1:A7)"), "#VALUE!");
    assert!(is_error(NUMS, "=RANK.AVG(3)"));
}

// --- CORREL -------------------------------------------------------------------

#[test]
fn correl_is_pearson_and_skips_non_numbers() {
    // Excel's documentation example: 0.997054486.
    let doc = &[
        ("A1", "3"), ("A2", "2"), ("A3", "4"), ("A4", "5"), ("A5", "6"),
        ("B1", "9"), ("B2", "7"), ("B3", "12"), ("B4", "15"), ("B5", "17"),
    ];
    assert!(close(num(doc, "=CORREL(A1:A5,B1:B5)"), 0.997_054_485_744_178_5));
    assert!(close(num(&[], "=CORREL({1,2,3},{3,2,1})"), -1.0));
    // The text in A3's pair drops that pair entirely.
    let gap = &[("A1", "1"), ("A2", "2"), ("A3", "x"), ("A4", "3"), ("B1", "2"), ("B2", "4"), ("B3", "100"), ("B4", "6")];
    assert!(close(num(gap, "=CORREL(A1:A4,B1:B4)"), 1.0));
}

#[test]
fn correl_errors() {
    assert_eq!(eval(&[], "=CORREL({1,2,3},{1,2})"), "#N/A");
    assert_eq!(eval(&[], "=CORREL({1,1,1},{1,2,3})"), "#DIV/0!");
    assert_eq!(eval(&[], "=CORREL({1},{2})"), "#DIV/0!");
    assert_eq!(eval(&[("A1", "1"), ("A2", "=1/0"), ("B1", "1"), ("B2", "2")], "=CORREL(A1:A2,B1:B2)"), "#DIV/0!");
    assert!(is_error(&[], "=CORREL({1,2})"));
}

// --- COUNTUNIQUE --------------------------------------------------------------

#[test]
fn countunique_counts_distinct_values_across_arguments() {
    assert_eq!(num(NUMS, "=COUNTUNIQUE(A1:A7)"), 3.0); // 3, 5, "text"; the blank is not a value
    assert_eq!(num(NUMS, "=COUNTUNIQUE(A1:A7,7,\"TEXT\")"), 4.0); // "TEXT" is "text"
    assert_eq!(num(&[], "=COUNTUNIQUE(1,2,2,\"1\")"), 3.0); // the number 1 and the text "1" differ
    assert_eq!(num(&[], "=COUNTUNIQUE({1,2;2,3},TRUE,1)"), 4.0); // TRUE is not 1
    assert_eq!(num(&[("A1", "")], "=COUNTUNIQUE(A1:A3)"), 0.0);
    assert_eq!(num(&[("A1", "=\"\""), ("A2", "x")], "=COUNTUNIQUE(A1:A2)"), 1.0);
}

#[test]
fn countunique_errors() {
    assert_eq!(eval(&[("A1", "1"), ("A2", "=1/0")], "=COUNTUNIQUE(A1:A2)"), "#DIV/0!");
    assert!(is_error(&[], "=COUNTUNIQUE()"));
}

// --- SPLIT --------------------------------------------------------------------

#[test]
fn split_spills_a_row_and_splits_by_each_character() {
    assert_eq!(spill(&[], "=SPLIT(\"a,b,c\",\",\")", 1, 3), vec![vec!["a", "b", "c"]]);
    // Each character of the delimiter splits on its own, and empty pieces go.
    assert_eq!(spill(&[], "=SPLIT(\"a-b/c--d\",\"-/\")", 1, 4), vec![vec!["a", "b", "c", "d"]]);
    // split_by_each FALSE: only the whole delimiter splits.
    assert_eq!(spill(&[], "=SPLIT(\"a-/b-c\",\"-/\",FALSE)", 1, 2), vec![vec!["a", "b-c"]]);
    // remove_empty_text FALSE keeps the empty piece between two commas.
    assert_eq!(spill(&[], "=SPLIT(\"a,,b\",\",\",TRUE,FALSE)", 1, 3), vec![vec!["a", "", "b"]]);
    // Omitted flags keep their defaults.
    assert_eq!(spill(&[], "=SPLIT(\"a,,b\",\",\",,)", 1, 2), vec![vec!["a", "b"]]);
    // One piece is a plain value, not an array.
    assert_eq!(text(&[], "=SPLIT(\"abc\",\",\")"), "abc");
}

#[test]
fn split_pieces_that_read_as_numbers_are_numbers() {
    assert_eq!(num(&[], "=SUM(SPLIT(\"1,2,3.5\",\",\"))"), 6.5);
    // As in Sheets, "007" is the number 7.
    assert_eq!(num(&[], "=INDEX(SPLIT(\"007,x\",\",\"),1,1)"), 7.0);
    assert_eq!(text(&[], "=INDEX(SPLIT(\"007,x\",\",\"),1,2)"), "x");
}

#[test]
fn split_errors_and_edges() {
    assert!(eval(&[], "=SPLIT(\"\",\",\")").starts_with("#VALUE!"));
    assert!(eval(&[], "=SPLIT(\"a,b\",\"\")").starts_with("#VALUE!"));
    assert_eq!(eval(&[], "=SPLIT(1/0,\",\")"), "#DIV/0!");
    // Nothing but delimiters leaves a blank.
    assert_eq!(eval(&[], "=SPLIT(\",,\",\",\")"), "");
    assert!(is_error(&[], "=SPLIT(\"a\")"));
}

// --- CHAR, CODE, CLEAN ----------------------------------------------------------

#[test]
fn char_and_code_are_inverses() {
    assert_eq!(text(&[], "=CHAR(65)"), "A");
    assert_eq!(text(&[], "=CHAR(65.9)"), "A");
    assert_eq!(text(&[], "=CHAR(10)"), "\n");
    // 128-159 are Windows-1252, as in Excel.
    assert_eq!(text(&[], "=CHAR(128)"), "€");
    assert_eq!(text(&[], "=CHAR(150)"), "–");
    assert_eq!(text(&[], "=CHAR(233)"), "é");
    // Above 255 a Unicode code point, as in Google Sheets.
    assert_eq!(text(&[], "=CHAR(9731)"), "☃");
    assert_eq!(num(&[], "=CODE(\"A\")"), 65.0);
    assert_eq!(num(&[], "=CODE(\"Apple\")"), 65.0);
    assert_eq!(num(&[], "=CODE(\"€\")"), 128.0);
    assert_eq!(num(&[], "=CODE(\"é\")"), 233.0);
    assert_eq!(num(&[], "=CODE(CHAR(9731))"), 9731.0);
    assert_eq!(num(&[], "=CODE(1)"), 49.0);
    // Over a range they go element by element.
    assert_eq!(spill(&[], "=CHAR({72,105})", 1, 2), vec![vec!["H", "i"]]);
}

#[test]
fn char_and_code_errors() {
    assert_eq!(eval(&[], "=CHAR(0)"), "#VALUE!");
    assert_eq!(eval(&[], "=CHAR(-1)"), "#VALUE!");
    assert_eq!(eval(&[], "=CHAR(55296)"), "#VALUE!"); // a surrogate is not a character
    assert_eq!(eval(&[], "=CHAR(1114112)"), "#VALUE!");
    assert_eq!(eval(&[], "=CHAR(\"x\")"), "#VALUE!");
    assert_eq!(eval(&[], "=CODE(\"\")"), "#VALUE!");
    assert_eq!(eval(&[], "=CODE(NA())"), "#N/A");
}

#[test]
fn clean_removes_control_characters_only() {
    assert_eq!(text(&[], "=CLEAN(\"a\"&CHAR(10)&\"b\"&CHAR(9)&\"c\"&CHAR(7))"), "abc");
    // Spaces and other printable characters stay; CLEAN is not TRIM.
    assert_eq!(text(&[], "=CLEAN(\"  a é  \")"), "  a é  ");
    assert_eq!(text(&[], "=CLEAN(12)"), "12");
    assert_eq!(eval(&[], "=CLEAN(1/0)"), "#DIV/0!");
}

// --- FIXED, DOLLAR ----------------------------------------------------------------

#[test]
fn fixed_rounds_and_groups_thousands() {
    // Excel's documentation examples.
    assert_eq!(text(&[], "=FIXED(1234.567,1)"), "1,234.6");
    assert_eq!(text(&[], "=FIXED(1234.567,-1)"), "1,230");
    assert_eq!(text(&[], "=FIXED(-1234.567,-1,TRUE)"), "-1230");
    assert_eq!(text(&[], "=FIXED(44.332)"), "44.33");
    assert_eq!(text(&[], "=FIXED(1234567.891)"), "1,234,567.89");
    assert_eq!(text(&[], "=FIXED(2.5,0)"), "3");
    assert_eq!(text(&[], "=FIXED(1.005,2)"), "1.01");
    assert_eq!(text(&[], "=FIXED(-0.001,2)"), "0.00");
    assert_eq!(text(&[], "=FIXED(999.996,2)"), "1,000.00");
    assert_eq!(text(&[], "=FIXED(\"12.5\",1)"), "12.5");
}

#[test]
fn dollar_writes_currency_with_negatives_in_parentheses() {
    // Excel's documentation examples.
    assert_eq!(text(&[], "=DOLLAR(1234.567,2)"), "$1,234.57");
    assert_eq!(text(&[], "=DOLLAR(1234.567,-2)"), "$1,200");
    assert_eq!(text(&[], "=DOLLAR(-1234.567,-2)"), "($1,200)");
    assert_eq!(text(&[], "=DOLLAR(-0.123,4)"), "($0.1230)");
    assert_eq!(text(&[], "=DOLLAR(99.888)"), "$99.89");
    assert_eq!(text(&[], "=DOLLAR(0)"), "$0.00");
}

#[test]
fn fixed_and_dollar_errors() {
    assert_eq!(eval(&[], "=FIXED(1,128)"), "#VALUE!");
    assert_eq!(eval(&[], "=DOLLAR(1,128)"), "#VALUE!");
    assert_eq!(eval(&[], "=FIXED(\"abc\")"), "#VALUE!");
    assert_eq!(eval(&[], "=DOLLAR(NA())"), "#N/A");
    assert!(is_error(&[], "=DOLLAR(1,2,TRUE)"), "DOLLAR has no no_commas argument");
    assert!(is_error(&[], "=FIXED()"));
}

// --- ISOWEEKNUM -----------------------------------------------------------------

#[test]
fn isoweeknum_is_weeknum_type_21() {
    // Excel's documentation example: 2012-03-09 is in ISO week 10.
    assert_eq!(num(&[], "=ISOWEEKNUM(DATE(2012,3,9))"), 10.0);
    // 2021-01-01 (a Friday) belongs to week 53 of 2020; 2024-12-30 to week 1 of 2025.
    assert_eq!(num(&[], "=ISOWEEKNUM(DATE(2021,1,1))"), 53.0);
    assert_eq!(num(&[], "=ISOWEEKNUM(DATE(2024,12,30))"), 1.0);
    assert_eq!(num(&[], "=ISOWEEKNUM(\"2026-10-09\")"), 41.0);
    for date in ["DATE(2026,1,1)", "DATE(2027,1,3)", "DATE(2015,12,31)"] {
        let iso = num(&[], &format!("=ISOWEEKNUM({date})"));
        assert_eq!(iso, num(&[], &format!("=WEEKNUM({date},21)")), "{date}");
    }
}

#[test]
fn isoweeknum_errors() {
    assert_eq!(eval(&[], "=ISOWEEKNUM(-1)"), "#NUM!");
    assert_eq!(eval(&[], "=ISOWEEKNUM(\"soon\")"), "#VALUE!");
    assert!(is_error(&[], "=ISOWEEKNUM()"));
}

// --- YEARFRAC -------------------------------------------------------------------

#[test]
fn yearfrac_all_bases() {
    // Excel's documentation example, 2012-01-01 to 2012-07-30, in each basis.
    let f = |basis: &str| num(&[], &format!("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30){basis})"));
    assert!(close(f(""), 209.0 / 360.0)); // 0.58055556
    assert!(close(f(",0"), 209.0 / 360.0));
    assert!(close(f(",1"), 211.0 / 366.0)); // 0.57650273
    assert!(close(f(",2"), 211.0 / 360.0));
    assert!(close(f(",3"), 211.0 / 365.0)); // 0.57808219
    assert!(close(f(",4"), 209.0 / 360.0));
    // The order of the dates does not matter, and times are ignored.
    assert!(close(num(&[], "=YEARFRAC(DATE(2012,7,30)+0.75,DATE(2012,1,1),3)"), 211.0 / 365.0));
}

#[test]
fn yearfrac_30_360_month_ends() {
    // US: February's last day counts as the 30th when both ends are on it...
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),0)"), 1.0));
    // ...and a start on it moves to the 30th either way.
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,2,28),DATE(2011,3,31),0)"), 30.0 / 360.0));
    // A 31st end counts as the 30th only when the start is the 30th or 31st.
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,1,30),DATE(2011,3,31),0)"), 60.0 / 360.0));
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,1,15),DATE(2011,3,31),0)"), 76.0 / 360.0));
    // European: every 31st is the 30th; February is not special.
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),4)"), 361.0 / 360.0));
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,1,15),DATE(2011,3,31),4)"), 75.0 / 360.0));
}

#[test]
fn yearfrac_actual_actual() {
    // Within a year that holds February 29 the year is 366 days...
    assert!(close(num(&[], "=YEARFRAC(DATE(2011,6,1),DATE(2012,3,1),1)"), 274.0 / 366.0));
    // ...otherwise 365.
    assert!(close(num(&[], "=YEARFRAC(DATE(2010,6,1),DATE(2011,3,1),1)"), 273.0 / 365.0));
    // Across years, the average length of the years touched.
    assert!(close(num(&[], "=YEARFRAC(DATE(2010,1,1),DATE(2012,7,1),1)"), 912.0 / (1096.0 / 3.0)));
    assert_eq!(num(&[], "=YEARFRAC(DATE(2012,5,5),DATE(2012,5,5),1)"), 0.0);
}

#[test]
fn yearfrac_errors() {
    assert_eq!(eval(&[], "=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),5)"), "#NUM!");
    assert_eq!(eval(&[], "=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),-1)"), "#NUM!");
    assert_eq!(eval(&[], "=YEARFRAC(-1,DATE(2012,7,30))"), "#NUM!");
    assert_eq!(eval(&[], "=YEARFRAC(\"x\",DATE(2012,7,30))"), "#VALUE!");
    assert!(is_error(&[], "=YEARFRAC(DATE(2012,1,1))"));
}

// --- MAP, REDUCE, SCAN, BYROW, BYCOL ---------------------------------------------------

const GRID: &[(&str, &str)] = &[("A1", "1"), ("B1", "2"), ("C1", "3"), ("A2", "4"), ("B2", "5"), ("C2", "6")];

#[test]
fn map_applies_a_lambda_to_each_element() {
    assert_eq!(spill(GRID, "=MAP(A1:C2,LAMBDA(x,x*10))", 2, 3), vec![vec!["10", "20", "30"], vec!["40", "50", "60"]]);
    // One parameter per array.
    assert_eq!(spill(GRID, "=MAP(A1:C1,A2:C2,LAMBDA(a,b,a+b))", 1, 3), vec![vec!["5", "7", "9"]]);
    // A single value pairs with every element; a missing element is #N/A.
    assert_eq!(spill(&[], "=MAP({1,2,3},10,LAMBDA(a,b,a*b))", 1, 3), vec![vec!["10", "20", "30"]]);
    assert_eq!(spill(&[], "=MAP({1,2,3},{1,2},LAMBDA(a,b,a+b))", 1, 3), vec![vec!["2", "4", "#N/A"]]);
    // Through a LET name, and with a captured LET value.
    assert_eq!(spill(GRID, "=LET(k,100,f,LAMBDA(x,x+k),MAP(A1:A2,f))", 2, 1), vec![vec!["101"], vec!["104"]]);
    // Text, and an error element reaches the LAMBDA rather than ending MAP.
    assert_eq!(spill(&[], "=MAP({\"a\",\"b\"},LAMBDA(s,UPPER(s)))", 1, 2), vec![vec!["A", "B"]]);
    assert_eq!(
        spill(&[("A1", "1"), ("A2", "=1/0")], "=MAP(A1:A2,LAMBDA(x,IFERROR(x,\"bad\")))", 2, 1),
        vec![vec!["1"], vec!["bad"]]
    );
    // A blank cell reads as blank inside the LAMBDA, and a blank answer is 0.
    assert_eq!(spill(&[("A1", "1")], "=MAP(A1:A2,LAMBDA(x,ISBLANK(x)))", 2, 1), vec![vec!["FALSE"], vec!["TRUE"]]);
    assert_eq!(spill(&[("A1", "1")], "=MAP(A1:A2,LAMBDA(x,x))", 2, 1), vec![vec!["1"], vec!["0"]]);
    // A built-in function's name stands for a LAMBDA (Excel's eta form).
    assert_eq!(spill(&[], "=MAP({-1,2},ABS)", 1, 2), vec![vec!["1", "2"]]);
    // A single element gives a single value.
    assert_eq!(eval(&[], "=MAP(4,LAMBDA(x,SQRT(x)))"), "2");
}

#[test]
fn map_errors() {
    assert!(eval(GRID, "=MAP(A1:A2,5)").starts_with("#VALUE!"));
    assert!(eval(GRID, "=MAP(A1:A2,LAMBDA(a,b,a+b))").starts_with("#VALUE!"));
    assert!(is_error(GRID, "=MAP(LAMBDA(x,x))"));
    // A LAMBDA that answers with an array cannot fill one cell.
    assert!(eval(GRID, "=MAP(A1:A2,LAMBDA(x,SEQUENCE(2)))").starts_with("#CALC!"));
    assert_eq!(eval(GRID, "=MAP(1/0,LAMBDA(x,x))"), "#DIV/0!");
}

#[test]
fn reduce_folds_and_scan_keeps_each_step() {
    assert_eq!(num(GRID, "=REDUCE(0,A1:C2,LAMBDA(acc,x,acc+x))"), 21.0);
    assert_eq!(num(GRID, "=REDUCE(1,A1:C2,LAMBDA(acc,x,acc*x))"), 720.0);
    // Initial value omitted: the accumulator starts blank.
    assert_eq!(num(GRID, "=REDUCE(,A1:C2,LAMBDA(acc,x,acc+x))"), 21.0);
    // Excel's documentation: sum of squares of the even values.
    assert_eq!(num(GRID, "=REDUCE(0,A1:C2,LAMBDA(a,b,IF(MOD(b,2)=0,a+b^2,a)))"), 56.0);
    assert_eq!(text(&[], "=REDUCE(\"\",{\"a\",\"b\",\"c\"},LAMBDA(acc,s,acc&s))"), "abc");
    // The accumulator may become an array.
    assert_eq!(spill(&[], "=REDUCE({0,0},{1,2},LAMBDA(acc,x,acc+x))", 1, 2), vec![vec!["3", "3"]]);
    // SCAN returns every step, shaped like the input.
    assert_eq!(spill(GRID, "=SCAN(0,A1:C2,LAMBDA(acc,x,acc+x))", 2, 3), vec![vec!["1", "3", "6"], vec!["10", "15", "21"]]);
    assert_eq!(spill(&[], "=SCAN(\"\",{\"a\",\"b\",\"c\"},LAMBDA(acc,s,acc&s))", 1, 3), vec![vec!["a", "ab", "abc"]]);
}

#[test]
fn reduce_and_scan_errors() {
    assert!(eval(GRID, "=REDUCE(0,A1:A2,LAMBDA(x,x))").starts_with("#VALUE!"));
    assert!(is_error(GRID, "=REDUCE(0,A1:A2)"));
    assert!(eval(GRID, "=SCAN(0,A1:A2,\"f\")").starts_with("#VALUE!"));
    // An error step carries on through the accumulator, as in Excel.
    assert_eq!(eval(&[("A1", "1"), ("A2", "=1/0")], "=REDUCE(0,A1:A2,LAMBDA(a,x,a+x))"), "#DIV/0!");
    // SCAN steps must be single values.
    assert!(eval(&[], "=SCAN({0,0},{1,2},LAMBDA(a,x,a+x))").starts_with("#CALC!"));
}

#[test]
fn byrow_and_bycol_reduce_each_row_or_column() {
    assert_eq!(spill(GRID, "=BYROW(A1:C2,LAMBDA(r,SUM(r)))", 2, 1), vec![vec!["6"], vec!["15"]]);
    assert_eq!(spill(GRID, "=BYCOL(A1:C2,LAMBDA(c,MAX(c)))", 1, 3), vec![vec!["4", "5", "6"]]);
    assert_eq!(spill(GRID, "=BYROW(A1:C2,SUM)", 2, 1), vec![vec!["6"], vec!["15"]]);
    assert_eq!(spill(GRID, "=BYROW(A1:C2,LAMBDA(r,COUNT(r)&\" items\"))", 2, 1), vec![vec!["3 items"], vec!["3 items"]]);
    // A one-column array passes each row as a single value.
    assert_eq!(spill(GRID, "=BYROW(A1:A2,LAMBDA(r,r*2))", 2, 1), vec![vec!["2"], vec!["8"]]);
}

#[test]
fn byrow_and_bycol_errors() {
    assert!(eval(GRID, "=BYROW(A1:C2,LAMBDA(r,r*2))").starts_with("#CALC!"));
    assert!(eval(GRID, "=BYCOL(A1:C2,LAMBDA(a,b,a))").starts_with("#VALUE!"));
    assert!(eval(GRID, "=BYROW(A1:C2,A1)").starts_with("#VALUE!"));
    assert!(is_error(GRID, "=BYROW(A1:C2)"));
}

#[test]
fn lambda_helpers_recalculate_when_their_inputs_change() {
    let mut wb = book(GRID);
    wb.set_cell_value_tracked(0, 0, 19, "=BYROW(A1:C2,LAMBDA(r,SUM(r)))");
    assert_eq!(wb.active_sheet().get_display(0, 19), "6");
    wb.set_cell_value_tracked(0, 0, 0, "10");
    assert_eq!(wb.active_sheet().get_display(0, 19), "15");
}

// --- ARRAYFORMULA ----------------------------------------------------------------

#[test]
fn arrayformula_passes_its_argument_through() {
    let cells = &[("A1", "1"), ("A2", "2"), ("A3", "3"), ("B1", "10"), ("B2", "20"), ("B3", "30")];
    assert_eq!(spill(cells, "=ARRAYFORMULA(A1:A3*B1:B3)", 3, 1), vec![vec!["10"], vec!["40"], vec!["90"]]);
    // A bare range spills as an array.
    assert_eq!(spill(cells, "=ARRAYFORMULA(A1:A3)", 3, 1), vec![vec!["1"], vec!["2"], vec!["3"]]);
    assert_eq!(spill(cells, "=ARRAYFORMULA(IF(A1:A3>1,\"big\",\"small\"))", 3, 1), vec![vec!["small"], vec!["big"], vec!["big"]]);
    // Reduced to one value it is that value, exactly as without the wrapper.
    assert_eq!(num(cells, "=ARRAYFORMULA(SUM(A1:A3*B1:B3))"), 140.0);
    assert_eq!(num(cells, "=ARRAYFORMULA(5)"), 5.0);
}

#[test]
fn arrayformula_errors() {
    assert_eq!(eval(&[], "=ARRAYFORMULA(1/0)"), "#DIV/0!");
    assert!(is_error(&[], "=ARRAYFORMULA(1,2)"));
    assert!(is_error(&[], "=ARRAYFORMULA()"));
}

// --- Registration ------------------------------------------------------------------

#[test]
fn every_new_function_is_known() {
    for name in [
        "RANK.AVG", "CORREL", "COUNTUNIQUE", "SPLIT", "CHAR", "CODE", "CLEAN", "FIXED", "DOLLAR",
        "ISOWEEKNUM", "YEARFRAC", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "ARRAYFORMULA",
    ] {
        assert!(visigrid_engine::formula::functions::is_known_function(name), "{name}");
        assert!(visigrid_engine::formula::help::get_function(name).is_some(), "{name} has no help entry");
    }
}

