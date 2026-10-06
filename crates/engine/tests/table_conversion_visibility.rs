use visigrid_engine::{
    filter::SortDirection,
    table::{TableId, TableRange},
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};

fn book(legacy: bool) -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20", "30"].iter().enumerate() {
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
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    if legacy {
        let mut catalog = wb.saved_tables();
        catalog.sheets[0].tables[0]
            .totals
            .as_mut()
            .unwrap()
            .hidden_rows = [2, 4].into();
        wb.restore_tables(catalog).unwrap();
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
    } else {
        wb = wb
            .prepare_table_row_visibility(wb.active_sheet_id(), [2, 4].into())
            .unwrap()
            .0;
    }
    (wb, id)
}

#[test]
fn conversion_preserves_legacy_and_shared_visibility_with_exact_undo() {
    for legacy in [false, true] {
        let (mut wb, id) = book(legacy);
        wb.set_cell_value_tracked(0, 0, 3, "=SUBTOTAL(9,Sales[Amount])");
        wb.set_cell_value_tracked(0, 1, 3, "=SUM(Sales[[#Totals],[Amount]])");
        assert_eq!(wb.active_sheet().get_display(1, 3), "40");
        let original = serde_json::to_value(wb.saved_tables()).unwrap();
        let formula = wb.active_sheet().get_raw(4, 0);
        let commit = wb.remove_table(id).unwrap();
        assert!(commit.is_conversion());
        assert!(wb.table(id).is_none());
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4].into());
        assert_eq!(wb.active_sheet().get_display(4, 0), "40");
        assert!(wb.active_sheet().get_raw(4, 0).contains("$A$2:$A$4"));
        assert_eq!(wb.active_sheet().get_display(0, 3), "60");
        assert_eq!(wb.active_sheet().get_display(1, 3), "40");
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), original);
        assert_eq!(wb.active_sheet().get_raw(4, 0), formula);
        assert_eq!(wb.active_sheet().get_display(4, 0), "40");
        wb.apply_table_commit(&commit, false).unwrap();
        let (shown, _) = wb
            .prepare_table_row_visibility(wb.active_sheet_id(), Default::default())
            .unwrap();
        assert_eq!(shown.active_sheet().get_display(4, 0), "60");
    }
}

#[test]
fn dormant_totals_and_unrelated_table_metadata_survive_conversion_history() {
    let (mut wb, id) = book(true);
    wb.set_table_totals_visible(id, false, [2, 4].into())
        .unwrap();
    for (row, value) in ["Other", "5", "7"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row + 7, 0, value);
    }
    let other = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 7,
                end_row: 9,
                start_col: 0,
                end_col: 0,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(other, true, [2, 4].into())
        .unwrap();
    let other_before = wb.table(other).unwrap().1.clone();
    let commit = wb.remove_table(id).unwrap();
    assert_eq!(wb.table(other).unwrap().1, &other_before);
    assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4].into());
    assert_eq!(wb.active_sheet().get_display(10, 0), "12");
    wb.apply_table_commit(&commit, true).unwrap();
    assert!(!wb.table(id).unwrap().1.totals.as_ref().unwrap().visible);
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.table(other).unwrap().1, &other_before);
}

#[test]
fn changed_visibility_and_authored_cells_refuse_stale_conversion_replay() {
    let (mut wb, id) = book(true);
    let commit = wb.remove_table(id).unwrap();
    let (mut changed, visibility) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [4].into())
        .unwrap();
    let revision = changed.revision();
    assert!(changed.apply_table_commit(&commit, true).is_err());
    assert_eq!(changed.revision(), revision);
    visibility.replay(&mut changed, true).unwrap();
    changed.apply_table_commit(&commit, true).unwrap();
    changed.set_cell_value_tracked(0, 1, 0, "99");
    assert!(changed.apply_table_commit(&commit, false).is_err());
    assert!(changed.table(id).is_some());
    assert_eq!(changed.active_sheet().get_raw(1, 0), "99");
}

#[test]
fn another_sheet_view_is_preserved_and_invalid_dynamic_recalculation_is_atomic() {
    for unsafe_formula in [false, true] {
        let (mut wb, id) = book(true);
        let other = wb.add_sheet_named("Sorted").unwrap();
        wb.set_cell_value_tracked(other, 2, 1, "Value");
        wb.set_cell_value_tracked(other, 3, 1, "2");
        wb.set_cell_value_tracked(other, 4, 1, "1");
        let sheet_id = wb.sheet(other).unwrap().id;
        let sorted = wb
            .create_table(
                sheet_id,
                TableRange {
                    start_row: 2,
                    end_row: 4,
                    start_col: 1,
                    end_col: 1,
                },
                "SortedData",
            )
            .unwrap()
            .table_id();
        let mut spec = TableViewSpec::new(sorted);
        spec.sort = Some(TableSort {
            column: wb.table(sorted).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        wb.set_table_view_spec(sheet_id, Some(spec.clone()))
            .unwrap();
        if unsafe_formula {
            wb.set_cell_value_tracked(
                other,
                0,
                0,
                "=IFERROR(SUM(INDIRECT(\"Sales[Amount]\")),SEQUENCE(5))",
            );
            wb.set_auto_recalc(false);
        }
        let revision = wb.revision();
        let result = wb.remove_table(id);
        if unsafe_formula {
            assert!(result.is_err());
            assert_eq!(wb.revision(), revision);
            assert!(wb.table(id).is_some());
            assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "60");
        } else {
            let commit = result.unwrap();
            wb.apply_table_commit(&commit, true).unwrap();
            wb.apply_table_commit(&commit, false).unwrap();
        }
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4].into());
        assert_eq!(wb.sheet(other).unwrap().table_view_spec(), Some(&spec));
    }
}

#[test]
fn own_saved_view_and_recovery_still_refuse_before_removing_table() {
    let (mut wb, id) = book(true);
    let spec = TableViewSpec::new(id);
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    let revision = wb.revision();
    assert!(wb.remove_table(id).is_err());
    assert_eq!(wb.revision(), revision);
    assert!(wb.table(id).is_some());
    wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(wb.remove_table(id).is_err());
    assert!(wb.table(id).is_some());
}

#[test]
fn frozen_formula_sources_convert_without_changing_their_cached_text() {
    let (mut wb, id) = book(true);
    wb.set_cell_text_tracked(0, 0, 3, "Captured total");
    let mut cell = wb.active_sheet().get_cell(0, 3);
    cell.set_frozen_formula(Some("=SUM(Sales[Amount])".into()));
    wb.restore_cell_tracked(0, 0, 3, Some(cell)).unwrap();
    let commit = wb.remove_table(id).unwrap();
    assert_eq!(wb.active_sheet().get_display(0, 3), "Captured total");
    assert!(wb
        .active_sheet()
        .get_cell(0, 3)
        .frozen_formula()
        .unwrap()
        .contains("$A$2:$A$4"));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(
        wb.active_sheet().get_cell(0, 3).frozen_formula(),
        Some("=SUM(Sales[Amount])")
    );
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(0, 3), "Captured total");
}

#[test]
fn referenced_rule_metadata_converts_and_undo_restores_its_binding() {
    use visigrid_engine::{
        cond_format::CondStyle,
        validation::{CellRange, ValidationRule},
    };
    for conditional in [false, true] {
        let (mut wb, id) = book(true);
        if conditional {
            wb.active_sheet_mut().cond_formats.add(
                vec![CellRange::new(0, 0, 8, 0)],
                "=[@Amount]>0",
                CondStyle::Inline(Default::default()),
            );
        } else {
            wb.active_sheet_mut().validations.set(
                CellRange::single(8, 0),
                ValidationRule::list_range("Sales[Amount]"),
            );
        }
        let before = serde_json::to_value((
            &wb.active_sheet().cond_formats,
            wb.active_sheet().validations.iter().collect::<Vec<_>>(),
        ))
        .unwrap();
        let commit = wb.remove_table(id).unwrap();
        assert!(wb.table(id).is_none());
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4].into());
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(
            serde_json::to_value((
                &wb.active_sheet().cond_formats,
                wb.active_sheet().validations.iter().collect::<Vec<_>>()
            ))
            .unwrap(),
            before
        );
        wb.apply_table_commit(&commit, false).unwrap();
    }
}

#[test]
fn conversion_invalidates_cross_sheet_dynamic_formula_source_generations() {
    let (mut wb, id) = book(true);
    let other = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=IFERROR(SUM(INDIRECT(\"Sales[Amount]\")),0)");
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "60");
    let generation = wb.sheet(other).unwrap().edit_generation();
    wb.remove_table(id).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "0");
    assert!(wb.sheet(other).unwrap().edit_generation() > generation);
}

#[test]
fn an_unrelated_existing_cycle_does_not_block_conversion_or_its_history() {
    let (mut wb, id) = book(false);
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=A1");
    let commit = wb.remove_table(id).unwrap();
    assert!(wb.table(id).is_none());
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "#CYCLE!");
    wb.apply_table_commit(&commit, true).unwrap();
    assert!(wb.table(id).is_some());
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "=A1");
    wb.apply_table_commit(&commit, false).unwrap();
    assert!(wb.table(id).is_none());
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "#CYCLE!");
}
