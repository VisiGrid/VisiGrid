use visigrid_engine::{
    formula::eval::{EvalArg, EvalResult, Value},
    named_range::NamedRange,
    workbook::Workbook,
};

fn book() -> Workbook {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "10");
    wb.set_cell_value_tracked(0, 1, 0, "20");
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "999");
    wb.set_cell_value_tracked(other, 1, 0, "888");
    wb.define_name_for_cell("TargetCell", 0, 0, 0).unwrap();
    wb.define_name_for_range("TargetSpan", 0, 0, 0, 1, 0)
        .unwrap();
    wb
}

#[test]
fn names_preserve_target_sheet_in_scalars_aggregates_arrays_and_lookup_arguments() {
    let mut wb = book();
    for (row, source) in [
        "=TargetCell",
        "=SUM(TargetSpan)",
        "=SUM(TargetSpan*2)",
        "=INDEX(TargetSpan,2)",
        "=SUMIF(TargetSpan,\">10\")",
        "=IF(TRUE,TargetCell,0)",
        "=SUBTOTAL(9,TargetSpan)",
        "=ROW(TargetCell)",
        "=SUM(OFFSET(TargetCell,0,0,2,1))",
    ]
    .iter()
    .enumerate()
    {
        wb.set_cell_value_tracked(1, row, 2, source);
    }
    for (row, value) in ["10", "30", "60", "20", "20", "10", "30", "1", "30"]
        .iter()
        .enumerate()
    {
        assert_eq!(wb.sheet(1).unwrap().get_display(row, 2), *value);
    }
    wb.set_cell_value_tracked(0, 0, 0, "15");
    for (row, value) in ["15", "35", "70", "20", "35", "15", "35", "1", "35"]
        .iter()
        .enumerate()
    {
        assert_eq!(wb.sheet(1).unwrap().get_display(row, 2), *value);
    }
    wb.named_ranges_mut()
        .set(NamedRange::cell("MissingSheet", 99, 0, 0))
        .unwrap();
    wb.set_cell_value_tracked(1, 9, 2, "=MissingSheet");
    assert!(wb.sheet(1).unwrap().get_display(9, 2).starts_with("#REF!"));
}

#[test]
fn named_ranges_keep_sheet_identity_through_custom_function_context() {
    let mut wb = book();
    wb.set_cell_value_tracked(1, 0, 2, "=MYTOTAL(TargetSpan)");
    let handler = |name: &str, args: &[EvalArg]| {
        if name != "MYTOTAL" {
            return None;
        }
        let [EvalArg::Range { values, num_cells }] = args else {
            panic!("expected a range")
        };
        assert_eq!(*num_cells, 2);
        Some(EvalResult::Number(
            values
                .iter()
                .map(|v| match v {
                    Value::Number(n) => *n,
                    _ => panic!("expected a number"),
                })
                .sum(),
        ))
    };
    wb.recompute_full_ordered_with_custom_fns(&handler);
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 2), "30");
}
