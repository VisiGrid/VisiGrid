use visigrid_engine::{
    filter::{ColumnFilter, FilterKey, SortDirection},
    formula::eval::Value,
    sheet::{Sheet, SheetId},
    structural::Axis,
    table::{TableColumnId, TableId, TableRange},
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::Workbook,
};

fn fixture() -> (Workbook, TableViewSpec) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 30, 8)], 0);
    for (row, values) in [["Name", "Amount"], ["b", "20"], ["a", "10"], ["c", "5"]]
        .iter()
        .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, row + 2, col + 1, value);
        }
    }
    let id = wb
        .create_table(
            SheetId(7),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 5,
                end_col: 2,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_cell_value_tracked(0, 0, 0, "=SUM(Sales[Amount])");
    let column = wb.table(id).unwrap().1.columns[1].id;
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column,
        direction: SortDirection::Ascending,
    });
    spec.filters.push(TableFilter {
        column,
        criteria: ColumnFilter {
            selected: Some(
                [
                    FilterKey::from_value(&Value::Number(10.0)).normalized(),
                    FilterKey::from_value(&Value::Number(20.0)).normalized(),
                ]
                .into_iter()
                .collect(),
            ),
            text_filter: None,
        },
    });
    (wb, spec)
}

#[test]
fn criteria_changes_and_history_only_touch_presentation() {
    let (mut wb, spec) = fixture();
    let revision = wb.revision();
    let generation = wb.active_sheet().edit_generation();
    let before: Vec<_> = wb
        .active_sheet()
        .cells_iter()
        .map(|(pos, cell)| (pos, cell.raw_display()))
        .collect();
    let commit = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    assert_eq!(commit.sheet_id(), SheetId(7));
    assert!(commit.before().is_none());
    assert_eq!(commit.after(), Some(&spec));
    assert_eq!(wb.revision(), revision + 1);
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    let view = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    assert_eq!(&view.rows().row_order()[3..=5], &[5, 4, 3]);
    assert!(!view.rows().is_data_row_visible(5));
    wb.apply_table_view_commit(&commit, true).unwrap();
    assert!(wb.active_sheet().table_view_spec().is_none());
    assert!(wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .is_none());
    wb.apply_table_view_commit(&commit, false).unwrap();
    assert_eq!(wb.revision(), revision + 3);
    assert_eq!(wb.active_sheet().edit_generation(), generation);
    assert_eq!(wb.active_sheet().get_display(0, 0), "35");
    assert_eq!(
        before,
        wb.active_sheet()
            .cells_iter()
            .map(|(pos, cell)| (pos, cell.raw_display()))
            .collect::<Vec<_>>()
    );
    let noop = wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
    assert!(noop.is_noop());
    assert_eq!(wb.revision(), revision + 3);
}

#[test]
fn hide_buttons_clear_sort_and_clear_filters_undo_independently() {
    let (mut wb, mut spec) = fixture();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    spec.show_filter_buttons = false;
    let hide = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    spec.clear_sort();
    let sort = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    assert!(!wb
        .active_sheet()
        .table_view_spec()
        .unwrap()
        .filters
        .is_empty());
    spec.clear_filters();
    let filter = wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
    wb.apply_table_view_commit(&filter, true).unwrap();
    assert!(!wb
        .active_sheet()
        .table_view_spec()
        .unwrap()
        .filters
        .is_empty());
    assert!(wb.active_sheet().table_view_spec().unwrap().sort.is_none());
    wb.apply_table_view_commit(&sort, true).unwrap();
    assert!(wb.active_sheet().table_view_spec().unwrap().sort.is_some());
    assert!(
        !wb.active_sheet()
            .table_view_spec()
            .unwrap()
            .show_filter_buttons
    );
    wb.apply_table_view_commit(&hide, true).unwrap();
    assert!(
        wb.active_sheet()
            .table_view_spec()
            .unwrap()
            .show_filter_buttons
    );
}

#[test]
fn stale_history_and_failed_activation_preserve_state_and_revision() {
    let (mut wb, mut spec) = fixture();
    let initial = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    spec.show_filter_buttons = false;
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    let revision = wb.revision();
    assert!(wb
        .apply_table_view_commit(&initial, true)
        .unwrap_err()
        .contains("changed"));
    let mut bad = spec.clone();
    bad.sort.as_mut().unwrap().column = TableColumnId(999);
    assert!(wb.set_table_view_spec(SheetId(7), Some(bad)).is_err());
    assert!(wb.set_table_view_spec(SheetId(999), None).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    let clear = wb.set_table_view_spec(SheetId(7), None).unwrap();
    wb.set_cell_value_tracked(0, 3, 6, "Adjacent note");
    let revision = wb.revision();
    assert!(wb.apply_table_view_commit(&clear, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert!(wb.active_sheet().table_view_spec().is_none());
}

#[test]
fn switching_owner_requires_clear_and_history_restores_each_owner() {
    let (mut wb, first) = fixture();
    let second = wb
        .create_table(
            SheetId(7),
            TableRange {
                start_row: 10,
                start_col: 1,
                end_row: 12,
                end_col: 2,
            },
            "Orders",
        )
        .unwrap()
        .table_id();
    let second = TableViewSpec::new(second);
    wb.set_table_view_spec(SheetId(7), Some(first.clone()))
        .unwrap();
    assert!(wb
        .set_table_view_spec(SheetId(7), Some(second.clone()))
        .unwrap_err()
        .contains("Clear"));
    let clear = wb.set_table_view_spec(SheetId(7), None).unwrap();
    let switch = wb.set_table_view_spec(SheetId(7), Some(second)).unwrap();
    wb.apply_table_view_commit(&switch, true).unwrap();
    wb.apply_table_view_commit(&clear, true).unwrap();
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&first));
}

#[test]
fn bindings_survive_rename_and_move_but_referenced_field_removal_is_atomic() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    wb.rename_table(spec.table, "Orders").unwrap();
    wb.rename_table_columns(spec.table, &["Region".into(), "Revenue".into()])
        .unwrap();
    wb.structural_edit(0, Axis::Col, 2, 1, false).unwrap();
    wb.structural_edit(0, Axis::Row, 1, 1, false).unwrap();
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    assert_eq!(
        wb.active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .filters()
            .sort
            .as_ref()
            .unwrap()
            .column,
        3
    );
    let before = serde_json::to_value(wb.saved_tables()).unwrap();
    let revision = wb.revision();
    assert!(wb.structural_edit(0, Axis::Col, 3, 1, true).is_err());
    assert!(wb.remove_table(spec.table).is_err());
    let mut range = wb.table(spec.table).unwrap().1.range;
    range.end_col = 2;
    assert!(wb.resize_table(spec.table, range).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
    wb.set_table_view_spec(SheetId(7), None).unwrap();
    wb.structural_edit(0, Axis::Col, 3, 1, true).unwrap();
}

#[test]
fn column_history_cannot_remove_a_field_with_a_later_view_binding() {
    let (mut wb, mut spec) = fixture();
    let history = wb
        .prepare_table_column_history(0, 2, 1, false)
        .unwrap()
        .unwrap();
    wb.structural_edit(0, Axis::Col, 2, 1, false).unwrap();
    spec.clear_filters();
    spec.sort.as_mut().unwrap().column = wb.table(spec.table).unwrap().1.columns[1].id;
    wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_column_history(&history, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(history_table(&wb)).unwrap().1.columns.len(), 3);
}

fn history_table(wb: &Workbook) -> TableId {
    wb.table_by_name("Sales").unwrap().1.id
}

#[test]
fn empty_bodies_and_late_neighbors_suspend_projection_without_losing_intent() {
    let (mut wb, spec) = fixture();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    wb.structural_edit(0, Axis::Row, 3, 3, true).unwrap();
    assert!(wb.validate_table_view_specs().is_ok());
    assert!(wb.active_sheet().build_saved_table_view(30).is_err());
    wb.restore_tables(wb.saved_tables()).unwrap();
    wb.append_table_rows(spec.table, 1, &[(3, 2, "10".into())])
        .unwrap();
    assert!(wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .is_some());
    wb.set_cell_value_tracked(0, 3, 6, "Late neighbor");
    assert!(wb.validate_table_view_specs().is_ok());
    assert!(wb.active_sheet().build_saved_table_view(30).is_err());
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    wb.set_table_view_spec(SheetId(7), None).unwrap();
}

#[test]
fn catalog_restore_is_atomic_strict_and_versioned() {
    let (mut wb, spec) = fixture();
    assert_eq!(wb.saved_tables().version, 1);
    wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
    assert_eq!(wb.saved_tables().version, 3);
    let valid = wb.saved_tables();
    let before = serde_json::to_value(&valid).unwrap();
    for case in 0..6 {
        let mut invalid = valid.clone();
        match case {
            0 => invalid.version = 2,
            1 => invalid.sheets[0].view.as_mut().unwrap().table = TableId(999),
            2 => {
                invalid.sheets[0]
                    .view
                    .as_mut()
                    .unwrap()
                    .sort
                    .as_mut()
                    .unwrap()
                    .column = TableColumnId(999)
            }
            3 => {
                let filter = invalid.sheets[0].view.as_ref().unwrap().filters[0].clone();
                invalid.sheets[0]
                    .view
                    .as_mut()
                    .unwrap()
                    .filters
                    .push(filter);
            }
            4 => invalid.sheets.push(invalid.sheets[0].clone()),
            _ => invalid.sheets[0].tables[0].range.start_col = usize::MAX,
        }
        assert!(wb.restore_tables(invalid).is_err());
        assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
    }
    let mut cleared = valid;
    cleared.sheets[0].view = None;
    cleared.version = 1;
    wb.restore_tables(cleared).unwrap();
    assert!(wb.active_sheet().table_view_spec().is_none());
    assert_eq!(wb.saved_tables().version, 1);
}

#[test]
fn recovery_is_read_only_for_view_changes_and_history_too() {
    let (mut wb, spec) = fixture();
    let commit = wb
        .set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    let revision = wb.revision();
    wb.active_sheet_mut().read_only_reason = Some("future Tables".into());
    assert!(wb.set_table_view_spec(SheetId(7), None).is_err());
    assert!(wb.apply_table_view_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
}

#[test]
fn nonfinite_filter_numbers_are_rejected_before_state_changes() {
    let (mut wb, mut spec) = fixture();
    spec.filters[0].criteria.selected = Some(
        [FilterKey::from_value(&Value::Number(f64::NAN)).normalized()]
            .into_iter()
            .collect(),
    );
    let revision = wb.revision();
    assert!(wb
        .set_table_view_spec(SheetId(7), Some(spec))
        .unwrap_err()
        .contains("finite"));
    assert_eq!(wb.revision(), revision);
    assert!(wb.active_sheet().table_view_spec().is_none());
}
