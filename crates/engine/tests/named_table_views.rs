use visigrid_engine::{
    filter::{ColumnFilter, SortDirection, TextFilter, TextFilterMode},
    sheet::{Sheet, SheetId},
    table::{TableId, TableRange, TableTotal},
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::Workbook,
};

fn fixture() -> (Workbook, TableId, TableViewSpec) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 50, 20)], 0);
    for (r, row) in [
        ["Region", "Amount"],
        ["West", "30"],
        ["East", "10"],
        ["West", "20"],
    ]
    .iter()
    .enumerate()
    {
        for (c, value) in row.iter().enumerate() {
            wb.set_cell_value_tracked(0, r, c, value);
        }
    }
    let id = wb
        .create_table(
            SheetId(7),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let columns = &wb.table(id).unwrap().1.columns;
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: columns[1].id,
        direction: SortDirection::Ascending,
    });
    spec.filters.push(TableFilter {
        column: columns[0].id,
        criteria: ColumnFilter {
            selected: None,
            text_filter: Some(TextFilter {
                mode: TextFilterMode::Equals,
                value: "West".into(),
                case_sensitive: false,
            }),
        },
    });
    (wb, id, spec)
}

#[test]
fn presets_apply_current_values_keep_manual_hides_and_restore_totals() {
    let (mut wb, id, spec) = fixture();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        id,
        1,
        TableTotal {
            function: Some("sum".into()),
            ..Default::default()
        },
    )
    .unwrap();
    wb.sheet_mut(0)
        .unwrap()
        .set_manual_hidden_rows([1].into())
        .unwrap();
    wb.recompute_full_ordered();
    assert_eq!(wb.active_sheet().get_display(4, 1), "30");
    let saved = wb.save_named_table_view(id, "West", spec.clone()).unwrap();
    assert!(saved.is_saved_view_change());
    assert!(wb.active_sheet().table_view_spec().is_none());
    assert_eq!(wb.saved_tables().version, 6);
    let applied = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(4, 1), "20");
    let view = wb
        .active_sheet()
        .build_saved_table_view(50)
        .unwrap()
        .unwrap();
    assert!(!view.rows().is_data_row_visible(1));
    assert!(!view.rows().is_data_row_visible(2));
    assert!(view.rows().is_data_row_visible(3));
    wb.apply_table_view_commit(&applied, true).unwrap();
    assert_eq!(wb.active_sheet().get_display(4, 1), "30");
    // A preset stores criteria, never a stale record permutation or cell value.
    wb.set_cell_value_tracked(0, 3, 1, "5");
    wb.apply_table_view_commit(&applied, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(4, 1), "5");
    wb.apply_table_commit(&saved, true).unwrap();
    assert!(wb.table(id).unwrap().1.saved_views.is_empty());
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    wb.apply_table_commit(&saved, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.saved_views.len(), 1);
}

#[test]
fn crud_replays_independently_of_active_criteria_and_rejects_stale_history() {
    let (mut wb, id, spec) = fixture();
    wb.save_named_table_view(id, "  West  ", spec.clone())
        .unwrap();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    let rename = wb
        .rename_named_table_view(id, "West", "Western region")
        .unwrap();
    let update = wb
        .update_named_table_view(id, "Western region", TableViewSpec::new(id))
        .unwrap();
    assert!(wb.apply_table_commit(&rename, true).is_err());
    let delete = wb.delete_named_table_view(id, "Western region").unwrap();
    for commit in [&delete, &update, &rename] {
        wb.apply_table_commit(commit, true).unwrap();
    }
    assert_eq!(wb.table(id).unwrap().1.saved_views[0].name, "West");
    assert_eq!(wb.table(id).unwrap().1.saved_views[0].view, spec);
    for commit in [&rename, &update, &delete] {
        wb.apply_table_commit(commit, false).unwrap();
    }
    assert!(wb.table(id).unwrap().1.saved_views.is_empty());
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
}

#[test]
fn names_limits_binding_validation_and_version_errors_are_atomic() {
    let (mut wb, id, spec) = fixture();
    wb.save_named_table_view(id, "West", spec.clone()).unwrap();
    let before = serde_json::to_string(&wb.saved_tables()).unwrap();
    for name in ["", "west", "A\nB", &"x".repeat(81)] {
        assert!(wb.save_named_table_view(id, name, spec.clone()).is_err());
    }
    let mut wrong = spec.clone();
    wrong.table = TableId(999);
    assert!(wb.save_named_table_view(id, "Wrong", wrong).is_err());
    let mut wrong = spec.clone();
    wrong.sort.as_mut().unwrap().column.0 = 999;
    assert!(wb
        .save_named_table_view(id, "Missing column", wrong)
        .is_err());
    let mut wrong = spec.clone();
    wrong.filters.push(wrong.filters[0].clone());
    assert!(wb
        .save_named_table_view(id, "Duplicate filter", wrong)
        .is_err());
    let mut catalog = wb.saved_tables();
    catalog.version = 5;
    assert!(wb
        .restore_tables(catalog)
        .unwrap_err()
        .contains("version 6"));
    assert_eq!(serde_json::to_string(&wb.saved_tables()).unwrap(), before);
    for i in 1..64 {
        wb.save_named_table_view(id, &format!("View {i}"), spec.clone())
            .unwrap();
    }
    assert!(wb.save_named_table_view(id, "65th", spec).is_err());
    assert_eq!(wb.table(id).unwrap().1.saved_views.len(), 64);
}

#[test]
fn renames_insertions_copies_and_conversion_keep_bindings_and_guard_deletion() {
    let (mut wb, id, spec) = fixture();
    wb.save_named_table_view(id, "West", spec.clone()).unwrap();
    wb.rename_table(id, "Orders").unwrap();
    wb.rename_table_columns(id, &["Area".into(), "Revenue".into()])
        .unwrap();
    assert_eq!(wb.table(id).unwrap().1.saved_views[0].view, spec);
    assert!(wb
        .prepare_table_column_history(0, 1, 1, true)
        .unwrap_err()
        .contains("Saved view"));
    let mut range = wb.table(id).unwrap().1.range;
    range.end_col = 0;
    assert!(wb
        .resize_table(id, range)
        .unwrap_err()
        .contains("Saved view"));
    let insert = wb
        .prepare_table_column_history(0, 1, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&insert, false).unwrap();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    assert_eq!(
        wb.active_sheet()
            .build_saved_table_view(50)
            .unwrap()
            .unwrap()
            .rows()
            .view_to_data(1),
        2
    );
    let (copy, index) = wb.prepare_sheet_copy(&wb, SheetId(7), "Copy").unwrap();
    let copied = &copy.sheet(index).unwrap().tables()[0];
    assert_ne!(copied.id, id);
    assert_eq!(copied.saved_views[0].view.table, copied.id);
    assert_eq!(copied.saved_views[0].view.sort, spec.sort);
    copy.sheet(index)
        .unwrap()
        .build_saved_table_view(50)
        .unwrap()
        .unwrap();
    wb.set_table_view_spec(SheetId(7), None).unwrap();
    let convert = wb.remove_table(id).unwrap();
    wb.apply_table_commit(&convert, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.saved_views[0].view, spec);
    wb.delete_named_table_view(id, "West").unwrap();
    assert!(wb.saved_tables().version < 6);
}
