use visigrid_engine::{
    filter::{ColumnFilter, FilterKey, SortDirection},
    formula::eval::Value,
    sheet::{Sheet, SheetId},
    table::TableRange,
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::Workbook,
};
use visigrid_io::{json, native};

fn fixture() -> (Workbook, TableViewSpec) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(77), 30, 8, "Data")], 0);
    for (row, value) in ["Amount", "20", "10", "5"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row + 2, 1, value);
    }
    let id = wb
        .create_table(
            SheetId(77),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 5,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_cell_value_tracked(0, 0, 0, "=SUM(Sales[Amount])");
    let column = wb.table(id).unwrap().1.columns[0].id;
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column,
        direction: SortDirection::Descending,
    });
    spec.filters.push(TableFilter {
        column,
        criteria: ColumnFilter {
            selected: Some(
                [
                    FilterKey::from_value(&Value::Number(10.0)).normalized(),
                    FilterKey::from_value(&Value::Number(5.0)).normalized(),
                ]
                .into_iter()
                .collect(),
            ),
            text_filter: None,
        },
    });
    spec.show_filter_buttons = false;
    (wb, spec)
}

fn check(wb: &Workbook, spec: &TableViewSpec) {
    let sheet = wb.active_sheet();
    assert_eq!(sheet.table_view_spec(), Some(spec));
    assert_eq!(sheet.get_display(0, 0), "35");
    assert_eq!(sheet.get_raw(3, 1), "20");
    let view = sheet.build_saved_table_view(30).unwrap().unwrap();
    assert_eq!(view.visible_body_rows(4, 2).unwrap(), [4, 5]);
    assert!(!view.rows().is_data_row_visible(3));
    assert!(view.rows().is_data_row_visible(2));
}

#[test]
fn every_native_save_path_roundtrips_criteria_and_rebuilds_projection() {
    let (mut wb, spec) = fixture();
    let fingerprint = native::compute_semantic_fingerprint(&wb);
    wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
        .unwrap();
    assert_eq!(native::compute_semantic_fingerprint(&wb), fingerprint);
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..4 {
        let path = dir.path().join(format!("view-{mode}.sheet"));
        match mode {
            0 => native::save_workbook(&wb, &path).unwrap(),
            1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap(),
            2 => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap(),
            _ => native::save(wb.active_sheet(), &path).unwrap(),
        }
        let loaded = native::load_workbook(&path).unwrap();
        check(&loaded, &spec);
        assert_eq!(native::compute_semantic_fingerprint(&loaded), fingerprint);
        let db = rusqlite::Connection::open(&path).unwrap();
        let raw: String = db
            .query_row("SELECT value FROM meta WHERE key='tables'", [], |r| {
                r.get(0)
            })
            .unwrap();
        let catalog: serde_json::Value = serde_json::from_str(&raw).unwrap();
        assert_eq!(catalog["version"], 3);
        let view = &catalog["sheets"][0]["view"];
        assert_eq!(view["table"], spec.table.0);
        assert!(view.get("row_order").is_none());
        assert!(view.get("visible_mask").is_none());
    }
}

#[test]
fn json_single_and_multi_sheet_roundtrips_owner_after_sheet_id_remapping() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
        .unwrap();
    let single = json::export_full(wb.active_sheet()).unwrap();
    wb.add_sheet_named("Other").unwrap();
    let second = wb.sheet(1).unwrap().id;
    let id = wb
        .create_table(
            second,
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 5,
                end_col: 1,
            },
            "OtherRecords",
        )
        .unwrap()
        .table_id();
    let other = TableViewSpec::new(id);
    wb.set_table_view_spec(second, Some(other.clone())).unwrap();
    let multi = json::export_workbook(&wb, &[], 0).unwrap();
    for raw in [single, multi.clone()] {
        let doc: serde_json::Value = serde_json::from_str(&raw).unwrap();
        assert_eq!(doc["version"], 3);
        assert_eq!(doc["table_catalog"]["version"], 3);
        let loaded = json::import_any(&raw).unwrap().0;
        check(&loaded, &spec);
        assert_ne!(loaded.active_sheet().id, SheetId(77));
    }
    let loaded = json::import_any(&multi).unwrap().0;
    assert!(loaded.sheet(1).unwrap().table_view_spec().is_none());
}

#[test]
fn clearing_last_criterion_omits_empty_views_on_every_save_path() {
    for calculated in [false, true] {
        for clear_sort_first in [false, true] {
            let (mut wb, mut spec) = fixture();
            spec.show_filter_buttons = true;
            if calculated {
                wb.set_calculated_column(spec.table, 1, 3, "=ROW()", true)
                    .unwrap();
            }
            wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
                .unwrap();
            if clear_sort_first {
                spec.clear_sort();
            } else {
                spec.clear_filters();
            }
            wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
                .unwrap();
            assert_eq!(wb.saved_tables().version, 3, "remaining criterion needs v3");
            spec.clear_sort();
            spec.clear_filters();
            // Clear sort/filters leave Some(empty), unlike the Clear view action.
            wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
                .unwrap();
            assert!(wb.active_sheet().table_view_spec().is_some());
            let expected_version = if calculated { 2 } else { 1 };
            assert_eq!(wb.saved_tables().version, expected_version);
            assert!(wb.saved_tables().sheets[0].view.is_none());
            let check_cleared = |loaded: Workbook| {
                assert!(loaded.read_only_reason().is_none());
                assert!(loaded.active_sheet().table_view_spec().is_none());
                assert_eq!(loaded.saved_tables().version, expected_version);
                assert_eq!(
                    loaded.table(spec.table).unwrap().1.columns[0]
                        .formula
                        .is_some(),
                    calculated
                );
            };
            for raw in [
                json::export_full(wb.active_sheet()).unwrap(),
                json::export_workbook(&wb, &[], 0).unwrap(),
            ] {
                let doc: serde_json::Value = serde_json::from_str(&raw).unwrap();
                assert_eq!(doc["table_catalog"]["version"], expected_version);
                check_cleared(json::import_any(&raw).unwrap().0);
            }
            let dir = tempfile::tempdir().unwrap();
            for mode in 0..4 {
                let path = dir.path().join(format!("cleared-{mode}.sheet"));
                match mode {
                    0 => native::save_workbook(&wb, &path).unwrap(),
                    1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path)
                        .unwrap(),
                    2 => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path)
                        .unwrap(),
                    _ => native::save(wb.active_sheet(), &path).unwrap(),
                }
                check_cleared(native::load_workbook(&path).unwrap());
            }
        }
    }
}

#[test]
fn hidden_filter_buttons_without_criteria_still_require_version_three() {
    let (mut wb, mut spec) = fixture();
    spec.clear_sort();
    spec.clear_filters();
    wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
        .unwrap();
    assert_eq!(wb.saved_tables().version, 3);
    let raw = json::export_workbook(&wb, &[], 0).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("hidden-buttons.sheet");
    native::save_workbook(&wb, &path).unwrap();
    for loaded in [
        json::import_any(&raw).unwrap().0,
        native::load_workbook(&path).unwrap(),
    ] {
        assert_eq!(loaded.active_sheet().table_view_spec(), Some(&spec));
    }
}

#[test]
fn damaged_current_view_and_future_catalog_use_distinct_read_only_recovery() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(77), Some(spec)).unwrap();
    let original: serde_json::Value =
        serde_json::from_str(&json::export_workbook(&wb, &[], 0).unwrap()).unwrap();
    for case in 0..4 {
        let mut doc = original.clone();
        let expected = if case == 0 {
            "Upgrade VisiGrid"
        } else {
            "corrupt"
        };
        match case {
            0 => doc["table_catalog"]["version"] = 6.into(),
            1 => doc["table_catalog"]["sheets"][0]["view"]["sort"]["column"] = 999.into(),
            2 => doc["table_catalog"]["version"] = 2.into(),
            _ => doc["table_catalog"]["sheets"][0]["view"]["unknown_criterion"] = true.into(),
        }
        let raw = doc.to_string();
        assert!(json::import_any(&raw).unwrap_err().contains(expected));
        let recovered = json::import_any_for_recovery(&raw).unwrap().0;
        assert!(recovered.read_only_reason().unwrap().contains(expected));
        assert_eq!(recovered.active_sheet().get_display(0, 0), "35");
        assert!(json::export_workbook(&recovered, &[], 0).is_err());
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("view.sheet");
        native::save_workbook(&wb, &path).unwrap();
        {
            let db = rusqlite::Connection::open(&path).unwrap();
            db.execute(
                "UPDATE meta SET value=?1 WHERE key='tables'",
                [doc["table_catalog"].to_string()],
            )
            .unwrap();
        }
        let before = std::fs::read(&path).unwrap();
        assert!(native::load_workbook(&path).unwrap_err().contains(expected));
        let (recovered, issue) = native::load_workbook_for_recovery(&path).unwrap();
        assert!(issue.unwrap().to_string().contains(expected));
        assert_eq!(recovered.active_sheet().get_display(0, 0), "35");
        assert!(native::save_workbook(&recovered, &path).is_err());
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }
}

#[test]
fn json_refuses_two_view_owners_on_one_sheet_without_discarding_either() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(77), Some(spec)).unwrap();
    let layout = json::SheetLayout {
        filter: Some(json::FilterSpec {
            range: (2, 1, 5, 1),
            columns: vec![],
            sort: None,
        }),
        ..Default::default()
    };
    assert!(json::export_full_with_layout(wb.active_sheet(), &layout).is_err());
    assert!(json::export_workbook(&wb, &[layout], 0).is_err());
    let mut doc: serde_json::Value =
        serde_json::from_str(&json::export_workbook(&wb, &[], 0).unwrap()).unwrap();
    doc["sheets"][0]["filter"] = serde_json::json!({"range":[2,1,5,1]});
    assert!(json::import_any(&doc.to_string())
        .unwrap_err()
        .contains("both"));
    assert!(json::import_any_for_recovery(&doc.to_string())
        .unwrap()
        .0
        .read_only_reason()
        .is_some());
}

#[test]
fn saved_intent_survives_suspended_layout_and_rebuilds_from_new_values() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
        .unwrap();
    wb.set_cell_value_tracked(0, 3, 1, "10");
    wb.set_cell_value_tracked(0, 4, 1, "99");
    wb.set_cell_value_tracked(0, 3, 6, "Adjacent content");
    let raw = json::export_workbook(&wb, &[], 0).unwrap();
    let mut loaded = json::import_any(&raw).unwrap().0;
    assert_eq!(loaded.active_sheet().table_view_spec(), Some(&spec));
    assert!(loaded.active_sheet().build_saved_table_view(30).is_err());
    loaded.clear_cell_tracked(0, 3, 6);
    let view = loaded
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    assert_eq!(&view.rows().row_order()[3..=5], &[4, 3, 5]);
    assert!(!view.rows().is_data_row_visible(4));
    assert!(view.rows().is_data_row_visible(3));
    assert_eq!(loaded.active_sheet().get_display(0, 0), "114");
}

#[test]
fn typed_selections_and_text_predicates_survive_both_formats() {
    use visigrid_engine::filter::{TextFilter, TextFilterMode};
    let (mut wb, mut spec) = fixture();
    wb.active_sheet_mut().set_text(4, 1, "10");
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    spec.filters[0].criteria = ColumnFilter {
        selected: Some(
            [
                Value::Number(10.0),
                Value::Text("10".into()),
                Value::Boolean(true),
                Value::Empty,
                Value::Error("#DIV/0!".into()),
            ]
            .iter()
            .map(|v| FilterKey::from_value(v).normalized())
            .collect(),
        ),
        text_filter: Some(TextFilter {
            mode: TextFilterMode::NotContains,
            value: "Ignore".into(),
            case_sensitive: true,
        }),
    };
    wb.set_table_view_spec(SheetId(77), Some(spec.clone()))
        .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("typed.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let raw = json::export_workbook(&wb, &[], 0).unwrap();
    for loaded in [
        native::load_workbook(&path).unwrap(),
        json::import_any(&raw).unwrap().0,
    ] {
        assert_eq!(loaded.active_sheet().table_view_spec(), Some(&spec));
        let view = loaded
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.rows().is_data_row_visible(4));
        assert!(!view.rows().is_data_row_visible(3));
        assert!(!view.rows().is_data_row_visible(5));
    }
}
