use std::cell::Cell;
use visigrid_engine::cell_id::CellId;
use visigrid_engine::formula::eval::{evaluate, CellLookup, EvalResult};
use visigrid_engine::formula::parser::{bind_expr_same_sheet, format_parsed_expr, parse};
use visigrid_engine::sheet::{SheetRef, NUM_COLS, NUM_ROWS};
use visigrid_engine::structural::Axis;
use visigrid_engine::workbook::Workbook;

#[test]
fn parse_format_and_reject_invalid_whole_ranges() {
    for formula in [
        "=SUM(A:A)",
        "=SUM($A:B)",
        "=SUM(A:$B)",
        "=SUM($1:$3)",
        "=SUM(3:1)",
        "=SUM('Data Sheet'!$A:$B)",
        "=SUM(Sheet2!1:3)",
        "=SUM(XFD:XFD)",
        "=SUM(1048576:1048576)",
    ] {
        assert_eq!(format_parsed_expr(&parse(formula).unwrap()), formula);
    }
    for formula in [
        "=SUM(A:1)",
        "=SUM(1:A)",
        "=SUM(0:1)",
        "=SUM(1.5:2)",
        "=SUM(1.0:2)",
        "=SUM(XFE:XFE)",
        "=SUM(1:1048577)",
        "=SUM($A)",
        "=SUM($0:$1)",
        "=SUM(A:)",
    ] {
        assert!(parse(formula).is_err(), "must reject {formula}");
    }
    assert!(parse("=SUM(NamedRange,1,TRUE)").is_ok());
}

#[test]
fn cell_formulas_support_columns_rows_and_reversed_ranges() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "2");
    wb.set_cell_value_tracked(0, 1, 0, "3");
    wb.set_cell_value_tracked(0, 0, 1, "7");
    for (r, formula, expected) in [
        (3, "=SUM(A:A)", "5"),
        (4, "=COUNT(A:A)", "2"),
        (5, "=SUM(A:B)", "12"),
        (6, "=SUM(1:1)", "9"),
        (7, "=SUM(B:A)", "12"),
        (8, "=SUM(2:1)", "12"),
        (9, "=SUM($A:$A)+SUM($2:$2)", "8"),
        (10, "=IF(TRUE,SUM(A:A),0)", "5"),
        (11, "=SUMIF(A:A,\">2\",B:B)", "0"),
    ] {
        wb.set_cell_value_tracked(0, r, 3, formula);
        assert_eq!(wb.active_sheet().get_display(r, 3), expected, "{formula}");
    }
}

#[test]
fn future_values_and_formulas_recalculate_in_dependency_order() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A:A)");
    wb.set_cell_value_tracked(0, 0, 3, "=C1*2");
    wb.set_cell_value_tracked(0, 100, 0, "4");
    assert_eq!(wb.active_sheet().get_display(0, 3), "8");
    wb.set_cell_value_tracked(0, 200, 1, "3");
    wb.set_cell_value_tracked(0, 200, 0, "=B201*2");
    assert_eq!(wb.active_sheet().get_display(0, 3), "20");
    wb.set_cell_value_tracked(0, 200, 1, "5");
    assert_eq!(wb.active_sheet().get_display(0, 3), "28");
    wb.clear_cell_tracked(0, 200, 0);
    assert_eq!(wb.active_sheet().get_display(0, 3), "8");
    wb.set_cell_value_tracked(0, 200, 0, "7");
    assert_eq!(wb.active_sheet().get_display(0, 3), "22");
    let id = wb.active_sheet().id;
    assert!(wb
        .get_dependents(id, NUM_ROWS - 1, 0)
        .contains(&CellId::new(id, 0, 2)));
    assert!(
        wb.dep_graph().referenced_cell_count() < 10,
        "no grid-sized dependency expansion"
    );
}

#[test]
fn future_columns_update_whole_rows() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 2, 0, "=SUM(1:1)");
    wb.set_cell_value_tracked(0, 0, NUM_COLS - 1, "11");
    assert_eq!(wb.active_sheet().get_display(2, 0), "11");
    wb.clear_cell_tracked(0, 0, NUM_COLS - 1);
    assert_eq!(wb.active_sheet().get_display(2, 0), "0");
}

#[test]
fn cross_sheet_ranges_rebuild_and_track_new_data() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet_named("Data Sheet").unwrap();
    wb.set_cell_value_tracked(data, 0, 0, "2");
    wb.set_cell_value_tracked(0, 0, 0, "=SUM('Data Sheet'!A:A)");
    wb.set_cell_value_tracked(0, 1, 0, "=SUM('Data Sheet'!1:1)");
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    assert_eq!(wb.active_sheet().get_display(0, 0), "2");
    wb.set_cell_value_tracked(data, 100, 0, "5");
    wb.set_cell_value_tracked(data, 0, 30, "9");
    assert_eq!(wb.active_sheet().get_display(0, 0), "7");
    assert_eq!(wb.active_sheet().get_display(1, 0), "11");
    wb.set_cell_value_tracked(0, 2, 0, "=SUM(Missing!A:A)");
    assert!(wb.active_sheet().get_display(2, 0).starts_with("#REF!"));
}

#[test]
fn subscriptions_are_removed_when_formula_changes() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A:A)");
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(B:B)");
    let id = wb.active_sheet().id;
    assert!(wb.get_dependents(id, 100, 0).is_empty());
    assert_eq!(wb.get_dependents(id, 100, 1), vec![CellId::new(id, 0, 2)]);
    wb.clear_cell_tracked(0, 0, 2);
    assert!(wb.get_dependents(id, 100, 1).is_empty());
}

#[test]
fn self_reference_and_new_cycles_are_detected() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=SUM(A:A)");
    assert!(wb.active_sheet().get_display(0, 0).starts_with('#'));
    wb.clear_cell_tracked(0, 0, 0);
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A:A)");
    wb.set_cell_value_tracked(0, 20, 0, "=C1");
    assert!(wb.active_sheet().get_display(0, 2).starts_with('#'));
}

#[test]
fn structural_edits_keep_the_open_axis() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "5");
    wb.set_cell_value_tracked(0, 3, 3, "=SUM($A:$B)");
    wb.structural_edit(0, Axis::Row, 0, 1, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(4, 3), "=SUM($A:$B)");
    wb.structural_edit(0, Axis::Col, 0, 1, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(4, 4), "=SUM($B:$C)");
    assert_eq!(wb.active_sheet().get_display(4, 4), "5");
    wb.structural_edit(0, Axis::Col, 1, 2, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(4, 2), "=SUM(#REF!)");
}

#[test]
fn whole_columns_read_only_the_data_extent() {
    struct ThreeRows {
        reads: Cell<usize>,
    }
    impl CellLookup for ThreeRows {
        fn data_bounds(&self, _: &SheetRef) -> (usize, usize) {
            (3, 1)
        }
        fn get_value(&self, row: usize, col: usize) -> f64 {
            assert!(row < 3 && col == 0);
            self.reads.set(self.reads.get() + 1);
            (row + 1) as f64
        }
        fn get_text(&self, row: usize, col: usize) -> String {
            self.get_value(row, col).to_string()
        }
    }
    let lookup = ThreeRows {
        reads: Cell::new(0),
    };
    for formula in ["=SUM(A:A)", "=SUM(A1:A3)"] {
        lookup.reads.set(0);
        let bound = bind_expr_same_sheet(&parse(formula).unwrap());
        assert_eq!(evaluate(&bound, &lookup).to_text(), "6");
        assert_eq!(lookup.reads.get(), 3);
    }
    let empty = Workbook::new();
    let bound = bind_expr_same_sheet(&parse("=SUM(A:A)").unwrap());
    assert!(matches!(
        evaluate(&bound, empty.active_sheet()),
        EvalResult::Number(0.0)
    ));
}

#[test]
fn copy_fill_preserves_absolute_axes_strings_and_sheet_names() {
    use visigrid_engine::formula::parser::adjust_formula_refs;
    for (formula, expected) in [
        ("=SUM(A:A,$B:C,1:1,$2:3)", "=SUM(C:C,$B:E,2:2,$2:4)"),
        (
            "=SUM(Sheet1!A:A,'Data 2'!1:1)",
            "=SUM(Sheet1!C:C,'Data 2'!2:2)",
        ),
        (
            "=IF(A1=\"A:A\",(SUM(A:A)+2)*3,0)",
            "=IF(C2=\"A:A\",(SUM(C:C)+2)*3,0)",
        ),
        ("=SUM($A:$A,$1:$1)", "=SUM($A:$A,$1:$1)"),
        ("=LOG10(A1)", "=LOG10(C2)"),
    ] {
        assert_eq!(adjust_formula_refs(formula, 1, 2), expected);
    }
    assert_eq!(adjust_formula_refs("=SUM(A:A)", 0, -1), "=SUM(#REF!)");
    assert_eq!(adjust_formula_refs("=SUM(1:1)", -1, 0), "=SUM(#REF!)");
    assert_eq!(adjust_formula_refs("=SUM(XFD:XFD)", 0, 1), "=SUM(#REF!)");
}

#[test]
fn cycle_preview_includes_empty_cells_in_open_ranges() {
    let mut wb = Workbook::new();
    let id = wb.active_sheet().id;
    assert!(wb.check_formula_cycle(id, 10, 0, "=SUM(A:A)").is_err());
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(A:A)");
    assert!(wb.check_formula_cycle(id, 100, 0, "=C1").is_err());
    assert!(wb.check_formula_cycle(id, 100, 1, "=C1").is_ok());
}

#[test]
fn spill_growth_and_shrink_update_whole_columns() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 2, "3");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(C1)");
    wb.set_cell_value_tracked(0, 0, 3, "=SUM(A:A)");
    assert_eq!(wb.active_sheet().get_display(0, 3), "6");
    wb.set_cell_value_tracked(0, 0, 2, "5");
    assert_eq!(wb.active_sheet().get_display(0, 3), "15");
    wb.set_cell_value_tracked(0, 0, 2, "2");
    assert_eq!(wb.active_sheet().get_display(0, 3), "3");
}

#[test]
#[ignore = "timing measurement; run explicitly with --ignored --nocapture"]
fn whole_range_three_row_timing() {
    use std::hint::black_box;
    use std::time::Instant;
    let mut wb = Workbook::new();
    for row in 0..3 {
        wb.set_cell_value_tracked(0, row, 0, "2");
    }
    for formula in ["=SUM(A1:A3)", "=SUM(A:A)"] {
        let ast = bind_expr_same_sheet(&parse(formula).unwrap());
        let start = Instant::now();
        for _ in 0..100_000 {
            black_box(evaluate(black_box(&ast), wb.active_sheet()));
        }
        eprintln!("{formula}: {:?} per evaluation", start.elapsed() / 100_000);
    }
}

#[test]
fn graph_mapping_preserves_open_axis_when_first_row_is_deleted() {
    use std::collections::HashSet;
    use visigrid_engine::dep_graph::DepGraph;
    use visigrid_engine::formula::parser::RangeAxis;
    use visigrid_engine::formula::whole_range::WholeRangeRef;
    use visigrid_engine::sheet::SheetId;
    let sheet = SheetId(1);
    let mut graph = DepGraph::new();
    let formula = CellId::new(sheet, 5, 2);
    graph.register_leaf_formula(formula);
    graph.set_whole_ranges(
        formula,
        vec![WholeRangeRef {
            sheet,
            axis: RangeAxis::Column,
            start: 0,
            end: 1,
        }],
    );
    graph.apply_mapping(|cell| {
        cell.row
            .checked_sub(1)
            .map(|row| CellId::new(cell.sheet, row, cell.col))
    });
    let moved = CellId::new(sheet, 4, 2);
    assert!(graph.is_formula_cell(moved));
    assert_eq!(
        graph
            .dependents(CellId::new(sheet, 100, 0))
            .collect::<HashSet<_>>(),
        HashSet::from([moved])
    );
}
