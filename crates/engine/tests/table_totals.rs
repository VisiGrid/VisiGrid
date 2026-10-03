use visigrid_engine::{
    filter::{ColumnFilter, NormalizedFilterKey},
    table::{TableRange, TableTotal, TableTotals},
    table_view::{TableFilter, TableViewSpec},
    workbook::Workbook,
};

fn book() -> Workbook {
    let mut wb = Workbook::new();
    for (row, data) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in data.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    wb.set_cell_value_tracked(0, 4, 0, "Total");
    wb.set_cell_value_tracked(0, 4, 1, "=SUBTOTAL(109,[Amount])");
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 3,
            end_col: 1,
        },
        "Sales",
    )
    .unwrap();
    let mut catalog = wb.saved_tables();
    catalog.version = 5;
    catalog.sheets[0].tables[0].totals = Some(TableTotals {
        visible: true,
        shown: Some(true),
        hidden_rows: Default::default(),
        columns: vec![
            TableTotal {
                label: Some("Total".into()),
                ..Default::default()
            },
            TableTotal {
                function: Some("sum".into()),
                ..Default::default()
            },
        ],
    });
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

#[test]
fn subtotal_function_numbers_and_hidden_rows_remain_distinct() {
    let mut wb = book();
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0]
        .totals
        .as_mut()
        .unwrap()
        .hidden_rows
        .insert(2);
    wb.restore_tables(catalog).unwrap();
    for (function, expected) in [
        (1, "20"),
        (2, "3"),
        (3, "3"),
        (4, "30"),
        (5, "10"),
        (6, "6000"),
        (7, "10"),
        (9, "60"),
        (10, "100"),
        (101, "20"),
        (102, "2"),
        (103, "2"),
        (104, "30"),
        (105, "10"),
        (106, "300"),
        (109, "40"),
        (110, "200"),
    ] {
        wb.set_cell_value_tracked(0, 0, 4, &format!("=SUBTOTAL({function},Sales[Amount])"));
        assert_eq!(
            wb.sheet(0).unwrap().get_display(0, 4),
            expected,
            "code {function}"
        );
    }
    wb.set_cell_value_tracked(0, 0, 4, "=SUBTOTAL(12,Sales[Amount])");
    assert!(wb
        .sheet(0)
        .unwrap()
        .get_display(0, 4)
        .starts_with("#VALUE!"));
}

#[test]
fn filter_changes_and_cross_sheet_precedents_recalculate_totals_and_dependents() {
    let mut wb = book();
    let other = wb.add_sheet_named("Control").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "East");
    wb.set_cell_value_tracked(0, 2, 0, "=Control!A1");
    wb.set_cell_value_tracked(other, 0, 1, "=SUM(Sales[[#Totals],[Amount]])");
    let (sid, table) = wb.tables().next().unwrap();
    let mut spec = TableViewSpec::new(table.id);
    spec.filters.push(TableFilter {
        column: table.columns[0].id,
        criteria: ColumnFilter {
            selected: Some([NormalizedFilterKey::Text("west".into())].into()),
            text_filter: None,
        },
    });
    let generation = wb.sheet(other).unwrap().edit_generation();
    let commit = wb.set_table_view_spec(sid, Some(spec)).unwrap();
    assert!(
        wb.sheet(other).unwrap().edit_generation() > generation,
        "changed subtotal dependents must invalidate pivot sources"
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 1), "40");
    wb.set_cell_value_tracked(other, 0, 0, "West");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 1), "60");
    wb.set_cell_value_tracked(other, 0, 0, "East");
    wb.apply_table_view_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_view_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
}

#[test]
fn structured_totals_and_all_sections_include_footer_only_where_requested() {
    let mut wb = book();
    for (source, expected) in [
        ("=SUM(Sales[Amount])", "60"),
        ("=SUM(Sales[[#All],[Amount]])", "120"),
        ("=SUM(Sales[[#Totals],[Amount]])", "60"),
    ] {
        wb.set_cell_value_tracked(0, 0, 4, source);
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), expected);
    }
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0].totals.as_mut().unwrap().visible = false;
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    assert!(wb.sheet(0).unwrap().get_display(0, 4).starts_with("#REF!"));
}

#[test]
fn nested_subtotals_are_ignored_and_numeric_looking_text_stays_text() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 2, 1, "=SUBTOTAL(9,G1:G1)");
    wb.set_cell_text_tracked(0, 3, 1, "30");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "10");
    wb.set_cell_value_tracked(0, 0, 4, "=SUBTOTAL(103,Sales[Amount])");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "2");
}

#[test]
fn totals_ownership_and_conversion_preserve_formulas_and_history() {
    let mut wb = book();
    let id = wb.tables().next().unwrap().1.id;
    let before = wb.saved_tables();
    assert!(wb.rename_table(id, "Renamed").is_err());
    assert!(wb.append_table_rows(id, 1, &[]).is_err());
    assert!(wb
        .prepare_sheet_copy(&wb, wb.active_sheet_id(), "Copy")
        .is_err());
    wb.set_cell_value_tracked(0, 4, 1, "999");
    wb.sheet_mut(0).unwrap().set_value(4, 1, "999");
    wb.sheet_mut(0).unwrap().clear_cell(4, 1);
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert!(wb
        .structural_edit(0, visigrid_engine::structural::Axis::Row, 2, 1, false)
        .is_err());
    assert_eq!(
        serde_json::to_value(wb.saved_tables()).unwrap(),
        serde_json::to_value(before).unwrap()
    );
    let commit = wb.remove_table(id).unwrap();
    assert!(wb.sheet(0).unwrap().get_raw(4, 1).contains("$B$2:$B$4"));
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(
        wb.sheet(0).unwrap().get_raw(4, 1),
        "=SUBTOTAL(109,[Amount])"
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
}

#[test]
fn totals_catalog_rejects_bad_versions_and_bounds_without_mutation() {
    let mut wb = book();
    let original = wb.saved_tables();
    let mut old = original.clone();
    old.version = 4;
    assert!(wb.restore_tables(old).is_err());
    let mut bad = original.clone();
    bad.sheets[0].tables[0]
        .totals
        .as_mut()
        .unwrap()
        .columns
        .pop();
    assert!(wb.restore_tables(bad).is_err());
    let mut bad = original.clone();
    bad.sheets[0].tables[0].range.end_row = visigrid_engine::sheet::NUM_ROWS - 1;
    assert!(wb.restore_tables(bad).is_err());
    assert_eq!(
        serde_json::to_value(wb.saved_tables()).unwrap(),
        serde_json::to_value(original).unwrap()
    );
}

#[test]
fn cross_sheet_structural_rewrite_of_a_footer_refuses_before_any_mutation() {
    let mut wb = book();
    let other = wb.add_sheet_named("Control").unwrap();
    wb.set_cell_value_tracked(other, 1, 0, "25");
    let mut catalog = wb.saved_tables();
    let mut cells_only = catalog.clone();
    cells_only.sheets[0].tables.clear();
    wb.restore_tables(cells_only).unwrap();
    wb.set_cell_value_tracked(0, 4, 1, "=Control!A2");
    let total = &mut catalog.sheets[0].tables[0].totals.as_mut().unwrap().columns[1];
    total.function = Some("custom".into());
    total.formula = Some("=Control!A2".into());
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    assert!(wb
        .structural_edit(other, visigrid_engine::structural::Axis::Row, 0, 1, false)
        .is_err());
    assert_eq!(wb.sheet(other).unwrap().get_raw(1, 0), "25");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "=Control!A2");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "25");
}
