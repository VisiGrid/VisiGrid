use visigrid_engine::{
    formula::{
        eval::{
            evaluate, EvalResult, LookupWithContext, LookupWithNamedRanges, NamedRangeResolution,
        },
        parser::{bind_expr_same_sheet, parse},
    },
    sheet::{Sheet, SheetId},
    workbook::Workbook,
};

#[test]
fn nested_dynamic_cells_recalculate_cross_sheet_values_and_spills() {
    let mut wb = Workbook::new();
    let other = wb.add_sheet_named("Options! O'Brien").unwrap();
    for (r, value) in ["3", "1", "2"].iter().enumerate() {
        wb.set_cell_value_tracked(other, r, 0, value);
    }
    wb.define_name_for_range("Values", other, 0, 0, 2, 0)
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "999");
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(OFFSET(INDIRECT(\"Values\"),0,0))");
    wb.set_cell_value_tracked(0, 1, 2, "=SUM(INDIRECT(\"'Options! O''Brien'!A1:A3\"))");
    wb.set_cell_value_tracked(0, 0, 4, "=SORT(INDIRECT(\"Values\"))");
    for r in 0..2 {
        assert_eq!(wb.sheet(0).unwrap().get_display(r, 2), "6");
    }
    for (r, expected) in ["1", "2", "3"].iter().enumerate() {
        assert_eq!(wb.sheet(0).unwrap().get_display(r, 4), *expected);
    }
    wb.set_cell_value_tracked(other, 0, 0, "4");
    for r in 0..2 {
        assert_eq!(wb.sheet(0).unwrap().get_display(r, 2), "7");
    }
    assert_eq!(wb.sheet(0).unwrap().get_display(2, 4), "4");
}

#[test]
fn reference_producers_compose_without_losing_reference_geometry() {
    let mut wb = Workbook::new();
    for (r, value) in ["10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, value);
    }
    for (r, formula) in [
        "=SUM(OFFSET(OFFSET(A1,1,0),0,0,2))",
        "=SUM(OFFSET(IF(TRUE,A2:A3,B1),0,0))",
        "=SUM(OFFSET(CHOOSE(2,B1,A2:A3),0,0))",
        "=SUM(OFFSET(INDEX(A1:B3,0,1),1,0,2))",
    ]
    .iter()
    .enumerate()
    {
        wb.set_cell_value_tracked(0, r, 3, formula);
        assert_eq!(wb.sheet(0).unwrap().get_display(r, 3), "50", "{formula}");
    }
}

#[test]
fn standalone_and_legacy_name_contexts_use_the_shared_resolver() {
    let mut sheet = Sheet::new(SheetId(1), 20, 10);
    sheet.set_value(0, 0, "10");
    sheet.set_value(1, 0, "20");
    let named = LookupWithNamedRanges::new(&sheet, |name| {
        name.eq_ignore_ascii_case("Numbers")
            .then_some(NamedRangeResolution::Range {
                start_row: 0,
                start_col: 0,
                end_row: 1,
                end_col: 0,
            })
    });
    let context = LookupWithContext::new(&named, 0, 0);
    let eval = |text| evaluate(&bind_expr_same_sheet(&parse(text).unwrap()), &context);
    assert_eq!(
        eval("=SUM(OFFSET(INDIRECT(\"Numbers\"),ROW()-1,0))"),
        EvalResult::Number(30.0)
    );
    for formula in [
        "=INDIRECT(\"A1\",FALSE)",
        "=INDIRECT(\"Missing!A1\")",
        "=INDIRECT(\"1+1\")",
    ] {
        assert!(matches!(eval(formula), EvalResult::Error(_)), "{formula}");
    }
}

#[test]
fn retargeting_updates_dependencies_and_downstream_values_in_one_batch() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "B1");
    wb.set_cell_value_tracked(0, 0, 1, "10");
    wb.set_cell_value_tracked(0, 0, 2, "=F1*3");
    wb.set_cell_value_tracked(0, 0, 3, "=INDIRECT(A1)");
    wb.set_cell_value_tracked(0, 0, 4, "=D1*2");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "20");
    wb.begin_batch();
    wb.set_cell_value_tracked(0, 0, 0, "C1");
    wb.set_cell_value_tracked(0, 0, 5, "7");
    let outcome = wb.end_batch_outcome();
    assert!(outcome.errors.is_empty());
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "42");
    for col in [2, 3, 4] {
        assert!(outcome
            .recalculated
            .cells()
            .unwrap()
            .iter()
            .any(|cell| cell.col == col));
    }
    wb.set_cell_value_tracked(0, 0, 5, "8");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "48");
    let unrelated = wb.set_cell_value_tracked(0, 0, 1, "999");
    assert!(
        unrelated.cells().unwrap().is_empty(),
        "retired target must lose its subscription"
    );
}

#[test]
fn full_recalculation_orders_dynamic_producers_before_their_readers() {
    let mut wb = Workbook::new();
    for (col, formula) in [
        (0, "=INDIRECT(\"D1\")"),
        (1, "=A1*2"),
        (2, "7"),
        (3, "=C1*3"),
    ] {
        wb.sheet_mut(0).unwrap().set_value(0, col, formula);
    }
    wb.rebuild_dep_graph();
    let report = wb.recompute_full_ordered();
    assert!(!report.had_cycles);
    assert!(report.errors.is_empty(), "{:?}", report.errors);
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "42");
    wb.set_cell_value_tracked(0, 0, 2, "8");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "48");
}

#[test]
fn inactive_indirect_branches_do_not_create_cycles_and_dynamic_cycles_can_clear() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "0");
    wb.set_cell_value_tracked(0, 0, 1, "=IF(A1,INDIRECT(\"B1\"),1)");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "1");
    wb.set_cell_value_tracked(0, 0, 0, "1");
    assert!(wb
        .sheet(0)
        .unwrap()
        .get_display(0, 1)
        .starts_with("#CYCLE!"));
    wb.set_cell_value_tracked(0, 0, 0, "0");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "1");
    wb.set_cell_value_tracked(0, 1, 0, "B2");
    wb.set_cell_value_tracked(0, 1, 2, "9");
    wb.set_cell_value_tracked(0, 1, 1, "=INDIRECT(A2)");
    assert!(wb
        .sheet(0)
        .unwrap()
        .get_display(1, 1)
        .starts_with("#CYCLE!"));
    wb.set_cell_value_tracked(0, 1, 0, "C2");
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 1), "9");
}

#[test]
fn dynamic_readers_keep_dependency_order_alongside_iterative_cycles() {
    let mut wb = Workbook::new();
    wb.set_auto_recalc(false);
    wb.set_iterative_enabled(true);
    wb.set_cell_value_tracked(0, 0, 0, "=INDIRECT(\"D1\")");
    wb.set_cell_value_tracked(0, 0, 1, "=A1*2");
    wb.set_cell_value_tracked(0, 0, 2, "=(C1+1)/2");
    wb.set_cell_value_tracked(0, 0, 3, "7");
    wb.set_cell_value_tracked(0, 1, 0, "=INDIRECT(\"C1\")");
    wb.set_cell_value_tracked(0, 1, 1, "=A2*2");
    let report = wb.recompute_full_ordered();
    assert!(report.converged, "{:?}", report.errors);
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "14");
    let value = wb
        .sheet(0)
        .unwrap()
        .get_computed_value(1, 1)
        .to_number()
        .unwrap();
    assert!((value - 2.0).abs() < 1e-8, "{value}");
    assert!(wb.sheet(0).unwrap().get_raw(0, 2).starts_with('='));
}
