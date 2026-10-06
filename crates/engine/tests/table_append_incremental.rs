use visigrid_engine::{
    table::{TableId, TableRange},
    workbook::Workbook,
};

fn linked() -> (Workbook, TableId, usize) {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 0,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 2, "=A4");
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "unrelated");
    (wb, id, other)
}

#[test]
fn static_append_history_preserves_unrelated_edits_but_refuses_new_footer_readers() {
    let (mut wb, id, other) = linked();
    let commit = wb.append_table_rows(id, 1, &[(3, 0, "5".into())]).unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "keep this edit");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "keep this edit");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "30");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 2), "35");
    wb.set_cell_value_tracked(other, 1, 0, "=Sheet1!A5");
    let before = wb.revision();
    assert!(wb
        .apply_table_commit(&commit, true)
        .unwrap_err()
        .contains("New footer references"));
    assert_eq!(wb.revision(), before);
}

#[test]
fn dirty_append_matches_full_calculation_with_symbolic_cross_sheet_and_dynamic_readers() {
    for dynamic in [false, true] {
        let (mut wb, id, other) = linked();
        wb.set_cell_value_tracked(other, 1, 0, "=SUM(Sales[Amount])");
        wb.set_cell_value_tracked(other, 2, 0, "=A2*2");
        wb.set_cell_value_tracked(
            other,
            3,
            0,
            if dynamic {
                "=INDIRECT(\"Sheet1!A4\")"
            } else {
                "=Sheet1!A4"
            },
        );
        wb.append_table_rows(id, 2, &[(3, 0, "5".into()), (4, 0, "=A2*3".into())])
            .unwrap();
        let mut full = wb.clone();
        full.rebuild_dep_graph();
        full.recompute_full_ordered();
        for (index, sheet) in wb.sheets().iter().enumerate() {
            for ((row, col), _) in sheet.cells_iter() {
                assert_eq!(
                    sheet.get_display(row, col),
                    full.sheet(index).unwrap().get_display(row, col),
                    "{index}:{row}:{col}"
                );
            }
        }
        assert_eq!(wb.sheet(other).unwrap().get_display(2, 0), "130");
    }
}

#[test]
fn append_does_not_evaluate_unrelated_table_or_name_readers() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    use visigrid_engine::{
        custom_fns,
        formula::eval::{EvalArg, EvalResult},
    };
    static CALLS: AtomicUsize = AtomicUsize::new(0);
    fn handler(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
        if name != "APPENDPROBE" {
            return None;
        }
        CALLS.fetch_add(1, Ordering::SeqCst);
        Some(EvalResult::Number(42.0))
    }
    struct Reset;
    impl Drop for Reset {
        fn drop(&mut self) {
            custom_fns::set_default_custom_fn_handler(None);
        }
    }
    custom_fns::set_default_custom_fn_handler(Some(handler));
    let _reset = Reset;
    let (mut wb, id, other) = linked();
    wb.set_cell_value_tracked(other, 4, 0, "Amount");
    wb.set_cell_value_tracked(other, 5, 0, "8");
    wb.create_table(
        wb.sheets()[other].id,
        TableRange {
            start_row: 4,
            start_col: 0,
            end_row: 5,
            end_col: 0,
        },
        "Unrelated",
    )
    .unwrap();
    wb.define_name_for_cell("UnrelatedName", other, 5, 0)
        .unwrap();
    wb.set_cell_value_tracked(other, 1, 0, "=APPENDPROBE(SUM(Unrelated[Amount]))");
    wb.set_cell_value_tracked(other, 2, 0, "=APPENDPROBE(UnrelatedName)");
    assert!(CALLS.load(Ordering::SeqCst) > 0);
    CALLS.store(0, Ordering::SeqCst);
    wb.append_table_rows(id, 1, &[(3, 0, "5".into())]).unwrap();
    assert_eq!(CALLS.load(Ordering::SeqCst), 0);
    assert_eq!(wb.sheet(other).unwrap().get_display(1, 0), "42");
    wb.set_cell_value_tracked(other, 5, 0, "9");
    assert!(
        CALLS.load(Ordering::SeqCst) > 0,
        "unrelated dependencies must remain live"
    );
}
