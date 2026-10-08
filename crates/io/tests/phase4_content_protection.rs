use visigrid_engine::{
    cell::CellComment,
    filter::SortDirection,
    named_range::NamedRange,
    table::TableRange,
    table_view::{TableSort, TableViewSpec},
    validation::{CellRange, NumericConstraint, ValidationRule},
    workbook::Workbook,
};
use visigrid_io::{content_protection::first_loss, json, native};

fn phase4(names: bool, saved_view: bool) -> Workbook {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let sid = wb.active_sheet_id();
    let id = wb
        .create_table(
            sid,
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 0,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, [2].into()).unwrap();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Descending,
    });
    view.show_filter_buttons = false;
    wb.set_table_view_spec(sid, Some(view.clone())).unwrap();
    if saved_view {
        wb.save_named_table_view(id, "Descending", view).unwrap();
    }
    let mut rule = ValidationRule::decimal(NumericConstraint::between(0.0, 100.0));
    rule.reference_origin = Some((1, 0));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(1, 0, 3, 0), rule);
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(2, 0));
    wb.active_sheet_mut().frozen_panes = (1, 1);
    wb.active_sheet_mut().set_comment(
        6,
        2,
        Some(CellComment {
            text: "Preserve the empty cell comment".into(),
            author: "Review".into(),
        }),
    );
    wb.set_cell_value_tracked(0, 11, 3, "obstruction");
    wb.set_cell_value_tracked(0, 10, 3, "=SEQUENCE(2)");
    wb.set_cell_value_tracked(0, 10, 5, "=FILTER(A2:A4,A2:A4>100)");
    assert_eq!(wb.active_sheet().get_display(10, 3), "#SPILL!");
    assert_eq!(wb.active_sheet().get_display(10, 5), "#CALC! No matches");
    if names {
        wb.named_ranges_mut()
            .set(NamedRange::cell("Rate", 0, 1, 0))
            .unwrap();
    }
    wb
}

#[test]
fn phase4_json_v4_v5_and_catalog_v5_v6_reopen_editable_without_content_loss() {
    for names in [false, true] {
        for saved_view in [false, true] {
            let wb = phase4(names, saved_view);
            let source = json::export_workbook(&wb, &[], 0).unwrap();
            let before: serde_json::Value = serde_json::from_str(&source).unwrap();
            assert_eq!(before["version"], if names { 5 } else { 4 });
            assert_eq!(
                before["table_catalog"]["version"],
                if saved_view { 6 } else { 5 }
            );
            let (mut loaded, layouts, active) = json::import_any(&source).unwrap();
            assert!(
                loaded.read_only_reason().is_none(),
                "{:?}",
                loaded.read_only_reason()
            );
            assert_eq!(loaded.active_sheet().get_display(10, 3), "#SPILL!");
            assert_eq!(loaded.active_sheet().get_display(10, 5), "#CALC! No matches");
            assert_eq!(loaded.active_sheet().manual_hidden_rows(), [2].into());
            assert_eq!(loaded.active_sheet().frozen_panes, (1, 1));
            assert_eq!(
                loaded.active_sheet().comment(6, 2),
                wb.active_sheet().comment(6, 2)
            );
            let after =
                serde_json::from_str(&json::export_workbook(&loaded, &layouts, active).unwrap())
                    .unwrap();
            assert_eq!(first_loss(&before, &after), None);
            loaded.set_cell_value_tracked(0, 8, 2, "editable");
            assert_eq!(loaded.active_sheet().get_raw(8, 2), "editable");
        }
    }
}

#[test]
fn future_phase4_field_stays_read_only_and_native_source_precedes_upgrade_marker() {
    for saved_view in [false, true] {
        let mut source: serde_json::Value = serde_json::from_str(
            &json::export_workbook(&phase4(true, saved_view), &[], 0).unwrap(),
        )
        .unwrap();
        source["sheets"][0]["future_phase4_field"] = serde_json::json!({"enabled": true});
        let source = serde_json::to_string(&source).unwrap();
        let (loaded, layouts, active) = json::import_any(&source).unwrap();
        assert!(loaded
            .read_only_reason()
            .unwrap()
            .contains("future_phase4_field"));
        assert_eq!(
            json::export_workbook(&loaded, &layouts, active).unwrap(),
            source
        );
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("protected.sheet");
        native::save_workbook(&loaded, &path).unwrap();
        let conn = rusqlite::Connection::open(&path).unwrap();
        let marker: String = conn
            .query_row("SELECT value FROM meta WHERE key='tables'", [], |r| {
                r.get(0)
            })
            .unwrap();
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(&marker).unwrap()["version"].as_u64(),
            Some(u64::MAX)
        );
        drop(conn);
        let restored = native::load_workbook(&path).unwrap();
        assert!(restored.read_only_reason().is_some());
        assert_eq!(json::export_workbook(&restored, &[], 0).unwrap(), source);
        let (recovery, issue) = native::load_workbook_for_recovery(&path).unwrap();
        assert!(
            issue.is_none(),
            "the retained source must be loaded before the sentinel catalog"
        );
        assert_eq!(json::export_workbook(&recovery, &[], 0).unwrap(), source);
    }
}

#[test]
fn protected_phase4_fingerprints_cover_names_and_engine_freeze_panes() {
    let mut source: serde_json::Value =
        serde_json::from_str(&json::export_workbook(&phase4(true, true), &[], 0).unwrap()).unwrap();
    source["future_field"] = serde_json::json!(true);
    let (wb, _, _) = json::import_any(&source.to_string()).unwrap();
    let mut changed = wb.clone();
    changed
        .named_ranges_mut()
        .set(NamedRange::cell("Rate", 0, 3, 0))
        .unwrap();
    assert!(json::export_workbook(&changed, &[], 0).is_err());
    let mut changed = wb;
    changed.active_sheet_mut().frozen_panes = (2, 1);
    assert!(json::export_workbook(&changed, &[], 0).is_err());
}
