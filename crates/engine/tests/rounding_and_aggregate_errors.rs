use visigrid_engine::{formula::eval::Value, workbook::Workbook};

fn eval(formula: &str) -> Value {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, formula);
    wb.active_sheet().get_computed_value(0, 0)
}

#[test]
fn mround_rounds_multiples_and_decimal_money() {
    for (formula, expected) in [
        ("=MROUND(10,3)", 9.0),
        ("=MROUND(-10,-3)", -9.0),
        ("=MROUND(1.3,0.2)", 1.4),
        ("=MROUND(10.5,3)", 12.0),
        ("=MROUND(-10.5,-3)", -12.0),
        ("=MROUND(1.005,0.01)", 1.01),
        ("=MROUND(-1.005,-0.01)", -1.01),
        ("=MROUND(1.2345*89.2,0.01)", 110.12),
        ("=MROUND(0,-3)", 0.0),
        ("=MROUND(-10,0)", 0.0),
        ("=MROUND(1e300,1e-300)", 1e300),
    ] {
        assert_eq!(eval(formula), Value::Number(expected), "{formula}");
    }
}

#[test]
fn mround_validates_arguments_and_preserves_errors() {
    for (formula, expected) in [
        ("=MROUND(5,-2)", "#NUM!"),
        ("=MROUND(-5,2)", "#NUM!"),
        ("=MROUND(1/0,0.01)", "#DIV/0!"),
        ("=MROUND(1,NA())", "#N/A"),
    ] {
        assert_eq!(eval(formula), Value::Error(expected.into()), "{formula}");
    }
    for formula in [
        "=MROUND()",
        "=MROUND(1)",
        "=MROUND(1,2,3)",
        "=MROUND(\"x\",1)",
    ] {
        assert!(matches!(eval(formula), Value::Error(_)), "{formula}");
    }
}

#[test]
fn mround_lifts_over_arrays() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=MROUND(SEQUENCE(3,1,1.005,1),0.01)");
    for (r, expected) in [1.01, 2.01, 3.01].into_iter().enumerate() {
        assert_eq!(
            wb.active_sheet().get_computed_value(r, 0),
            Value::Number(expected)
        );
    }
}

#[test]
fn aggregates_propagate_errors_and_recover_after_edits() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "4785");
    wb.set_cell_value_tracked(0, 1, 0, "=NOT_SUPPORTED_YET(1)");
    let error = wb.active_sheet().get_computed_value(1, 0);
    assert!(matches!(error, Value::Error(_)));
    for (r, function) in [
        "SUM", "AVERAGE", "MIN", "MAX", "PRODUCT", "MEDIAN", "STDEV.S", "VAR.S",
    ]
    .iter()
    .enumerate()
    {
        wb.set_cell_value_tracked(0, r, 2, &format!("={function}(A1:A2)"));
        assert_eq!(
            wb.active_sheet().get_computed_value(r, 2),
            error,
            "{function}"
        );
    }
    wb.set_cell_value_tracked(0, 0, 3, "=IFERROR(SUM(A1:A2),\"Check fees\")");
    assert_eq!(
        wb.active_sheet().get_computed_value(0, 3),
        Value::Text("Check fees".into())
    );
    wb.set_cell_value_tracked(0, 1, 0, "=MROUND(1.2345*89.2,0.01)");
    assert_eq!(
        wb.active_sheet().get_computed_value(0, 2),
        Value::Number(4895.12)
    );
    assert_eq!(
        wb.active_sheet().get_computed_value(0, 3),
        Value::Number(4895.12)
    );
}

#[test]
fn sum_checks_named_cross_sheet_whole_column_and_array_ranges() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "12");
    wb.set_cell_value_tracked(0, 9999, 0, "=1/0");
    wb.define_name_for_range("Fees", 0, 0, 0, 9999, 0).unwrap();
    wb.define_name_for_cell("BadFee", 0, 9999, 0).unwrap();
    wb.add_sheet_named("Summary").unwrap();
    for (r, formula) in [
        "=SUM(Sheet1!A1:A10000)",
        "=SUM(Sheet1!A:A)",
        "=SUM(Fees)",
        "=SUM(BadFee)",
        "=SUM(Sheet1!A10000)",
        "=SUM(OFFSET(Sheet1!A9999,0,0,2,1))",
    ]
    .iter()
    .enumerate()
    {
        wb.set_cell_value_tracked(1, r, 0, formula);
        assert_eq!(
            wb.sheet(1).unwrap().get_computed_value(r, 0),
            Value::Error("#DIV/0!".into()),
            "{formula}"
        );
    }
}

#[test]
fn count_ignores_range_errors_and_sum_ignores_error_looking_text() {
    let mut wb = Workbook::new();
    for (r, raw) in ["10", "=1/0", "=\"#N/A\"", "text", ""].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, raw);
    }
    for (r, formula, expected) in [
        (0, "=COUNT(A1:A5)", 1.0),
        (1, "=COUNTA(A1:A5)", 4.0),
        (2, "=COUNT(OFFSET(A1,0,0,5,1))", 1.0),
        (3, "=SUM(A3:A5)", 0.0),
    ] {
        wb.set_cell_value_tracked(0, r, 2, formula);
        assert_eq!(
            wb.active_sheet().get_computed_value(r, 2),
            Value::Number(expected),
            "{formula}"
        );
    }
}

#[test]
fn sum_propagates_spilled_errors_in_row_major_order() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=1/(SEQUENCE(3)-2)");
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A2:A3)");
    assert_eq!(
        wb.active_sheet().get_computed_value(0, 2),
        Value::Error("#DIV/0!".into())
    );
    wb.set_cell_value_tracked(0, 0, 1, "=NA()");
    wb.set_cell_value_tracked(0, 1, 2, "=SUM(A1:B3)");
    assert_eq!(
        wb.active_sheet().get_computed_value(1, 2),
        Value::Error("#N/A".into())
    );
}

#[test]
fn sum_propagates_formula_parse_errors() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "10");
    wb.set_cell_value_tracked(0, 1, 0, "=SUM(");
    let error = wb.active_sheet().get_computed_value(1, 0);
    assert!(matches!(error, Value::Error(_)));
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A1:A2)");
    assert_eq!(wb.active_sheet().get_computed_value(0, 2), error);
    wb.set_cell_value_tracked(0, 1, 2, "=COUNT(A1:A2)");
    assert_eq!(
        wb.active_sheet().get_computed_value(1, 2),
        Value::Number(1.0)
    );
}
