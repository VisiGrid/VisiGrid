use visigrid_engine::{
    filter::{ColumnFilter, FilterKey, SortDirection},
    formula::eval::Value,
    structural::Axis,
    table::TableRange,
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::{StructureStep, Workbook},
};

fn book() -> Workbook {
    let mut wb = Workbook::new();
    for (row, value) in ["10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    wb.set_cell_value_tracked(0, 0, 2, "=SUBTOTAL(109,A1:A3)");
    wb.set_cell_value_tracked(0, 1, 2, "=SUBTOTAL(9,A1:A3)");
    wb
}

#[test]
fn manual_hides_without_totals_recalculate_and_replay_with_stale_protection() {
    let mut before = book();
    let other = before.add_sheet_named("Summary").unwrap();
    before.set_cell_value_tracked(other, 0, 0, "=Sheet1!C1*2");
    let id = before.sheet(0).unwrap().id;
    let generation = before.sheet(other).unwrap().edit_generation();
    let (mut after, commit) = before.prepare_table_row_visibility(id, [1].into()).unwrap();
    assert!(!commit.is_empty());
    assert_eq!(commit.changed_cell_count(), 0);
    assert_eq!(after.sheet(0).unwrap().manual_hidden_rows(), [1].into());
    assert_eq!(after.sheet(0).unwrap().get_display(0, 2), "40");
    assert_eq!(after.sheet(0).unwrap().get_display(1, 2), "60");
    assert_eq!(after.sheet(other).unwrap().get_display(0, 0), "80");
    assert!(after.sheet(other).unwrap().edit_generation() > generation);
    assert_eq!(before.sheet(0).unwrap().get_display(0, 2), "60");
    commit.replay(&mut after, true).unwrap();
    assert!(after.sheet(0).unwrap().manual_hidden_rows().is_empty());
    assert_eq!(after.sheet(other).unwrap().get_display(0, 0), "120");
    commit.replay(&mut after, false).unwrap();
    after
        .sheet_mut(0)
        .unwrap()
        .set_manual_hidden_rows([0, 1].into())
        .unwrap();
    assert!(commit.replay(&mut after, true).is_err());
}

#[test]
fn projection_composes_manual_and_filter_visibility_without_changing_sort_or_criteria() {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "30", "10", "20"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 0,
            },
            "Data",
        )
        .unwrap()
        .table_id();
    let column = wb.table(id).unwrap().1.columns[0].id;
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column,
        direction: SortDirection::Ascending,
    });
    spec.filters.push(TableFilter {
        column,
        criteria: ColumnFilter {
            selected: Some(
                [10.0, 30.0]
                    .map(|v| FilterKey::from_value(&Value::Number(v)).normalized())
                    .into(),
            ),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec.clone()))
        .unwrap();
    let (hidden, commit) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [1, 8].into())
        .unwrap();
    assert!(hidden.active_sheet().tables()[0].totals.is_none());
    let view = hidden
        .active_sheet()
        .build_saved_table_view(20)
        .unwrap()
        .unwrap();
    assert_eq!(view.spec(), &spec);
    assert_eq!(view.visible_body_rows(1, 1).unwrap(), vec![2]);
    assert!(view.visible_body_rows(1, 2).is_err());
    assert_eq!(view.rows().data_to_view_unchecked(1), 3);
    assert!(!view.rows().is_data_row_visible(1));
    assert!(!view.rows().is_data_row_visible(3));
    assert!(!view.rows().is_data_row_visible(8));
    assert!(commit
        .candidate(&hidden, true)
        .unwrap()
        .active_sheet()
        .manual_hidden_rows()
        .is_empty());
    let (shown, _) = hidden
        .prepare_table_row_visibility(hidden.active_sheet_id(), Default::default())
        .unwrap();
    let view = shown
        .active_sheet()
        .build_saved_table_view(20)
        .unwrap()
        .unwrap();
    assert_eq!(view.visible_body_rows(1, 2).unwrap(), vec![2, 1]);
    assert!(!view.rows().is_data_row_visible(3));
}

#[test]
fn row_structure_shifts_manual_flags_and_undo_restores_deleted_flags() {
    let before = book();
    let (before, _) = before
        .prepare_table_row_visibility(before.active_sheet_id(), [1, 10].into())
        .unwrap();
    let (inserted, insertion) = before
        .prepare_guarded_structure(
            0,
            vec![StructureStep {
                axis: Axis::Row,
                at: 1,
                count: 2,
                delete: false,
            }],
        )
        .unwrap();
    assert_eq!(inserted.active_sheet().manual_hidden_rows(), [3, 12].into());
    assert_eq!(
        insertion
            .candidate(&inserted, true)
            .unwrap()
            .active_sheet()
            .manual_hidden_rows(),
        [1, 10].into()
    );
    let (deleted, deletion) = before
        .prepare_guarded_structure(
            0,
            vec![StructureStep {
                axis: Axis::Row,
                at: 1,
                count: 1,
                delete: true,
            }],
        )
        .unwrap();
    assert_eq!(deleted.active_sheet().manual_hidden_rows(), [9].into());
    assert_eq!(
        deletion
            .candidate(&deleted, true)
            .unwrap()
            .active_sheet()
            .manual_hidden_rows(),
        [1, 10].into()
    );
}

#[test]
fn unchanged_visibility_keeps_caches_and_invalid_visibility_is_atomic() {
    let mut wb = book();
    let cached = wb.active_sheet().get_computed_value(0, 2);
    wb.active_sheet_mut()
        .set_manual_hidden_rows(Default::default())
        .unwrap();
    assert_eq!(wb.active_sheet().get_computed_value(0, 2), cached);
    assert!(wb
        .active_sheet_mut()
        .set_manual_hidden_rows([visigrid_engine::sheet::NUM_ROWS].into())
        .is_err());
    assert!(wb.active_sheet().manual_hidden_rows().is_empty());
    wb.active_sheet_mut()
        .set_manual_hidden_rows([visigrid_engine::sheet::NUM_ROWS - 1].into())
        .unwrap();
    assert!(wb
        .prepare_guarded_structure(
            0,
            vec![StructureStep {
                axis: Axis::Row,
                at: 0,
                count: 1,
                delete: false
            }]
        )
        .is_err());
}
