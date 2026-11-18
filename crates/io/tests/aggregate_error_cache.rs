use visigrid_engine::{formula::eval::Value, workbook::Workbook};
use visigrid_io::{json, native};

#[test]
fn loading_does_not_restore_partial_totals_or_turn_errors_into_text() {
    let dir = tempfile::tempdir().unwrap();
    for stale_total in [false, true] {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "4785");
        wb.set_cell_value_tracked(0, 1, 0, "=MISSING_FEE(1)");
        wb.set_cell_value_tracked(0, 2, 0, "=SUM(A1:A2)");
        let error = wb.active_sheet().get_computed_value(1, 0);
        assert!(matches!(error, Value::Error(_)));
        if stale_total {
            wb.active_sheet()
                .cache_computed(2, 0, Value::Number(4785.0));
        }
        for single_sheet in [false, true] {
            let path = dir.path().join("fees.sheet");
            if single_sheet {
                native::save(wb.active_sheet(), &path).unwrap();
            } else {
                native::save_workbook(&wb, &path).unwrap();
            }
            let loaded = native::load_workbook(&path).unwrap();
            assert_eq!(loaded.active_sheet().get_computed_value(1, 0), error);
            assert_eq!(loaded.active_sheet().get_computed_value(2, 0), error);
            assert!(loaded.active_sheet().kept_uncomputable.is_empty());
        }
        let text = json::export_workbook(&wb, &[], 0).unwrap();
        let (loaded, _, _) = json::import_any(&text).unwrap();
        assert_eq!(loaded.active_sheet().get_computed_value(1, 0), error);
        assert_eq!(loaded.active_sheet().get_computed_value(2, 0), error);
    }
}

#[test]
fn successful_custom_function_caches_are_still_preserved() {
    let mut wb = Workbook::new();
    for (row, value) in [Value::Number(42.0), Value::Text("Approved".into())]
        .into_iter()
        .enumerate()
    {
        wb.set_cell_value_tracked(0, row, 0, "=CUSTOM_FEE(1)");
        wb.active_sheet().cache_computed(row, 0, value);
    }
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("custom.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let native = native::load_workbook(&path).unwrap();
    let (json, _, _) = json::import_any(&json::export_workbook(&wb, &[], 0).unwrap()).unwrap();
    for loaded in [native, json] {
        assert_eq!(
            loaded.active_sheet().get_computed_value(0, 0),
            Value::Number(42.0)
        );
        assert_eq!(
            loaded.active_sheet().get_computed_value(1, 0),
            Value::Text("Approved".into())
        );
        assert_eq!(loaded.active_sheet().kept_uncomputable.len(), 2);
    }
}
