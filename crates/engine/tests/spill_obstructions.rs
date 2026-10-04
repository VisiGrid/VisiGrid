use visigrid_engine::workbook::Workbook;

#[test]
fn clearing_a_blocker_retries_the_spill_and_cross_sheet_dependents() {
    let mut wb = Workbook::new();
    wb.add_sheet();
    wb.set_cell_value_tracked(0, 1, 0, "occupied");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(3)");
    wb.set_cell_value_tracked(1, 0, 0, "=Sheet1!A3*10");
    wb.recompute_full_ordered();
    assert!(wb.sheet(0).unwrap().has_spill_error(0, 0));
    wb.clear_cell_tracked(0, 1, 0);
    assert_eq!(wb.sheet(0).unwrap().get_display(2, 0), "3");
    assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "30");
    assert!(!wb.sheet(0).unwrap().has_spill_error(0, 0));
}

#[test]
fn multiple_obstructions_rebind_and_batch_clearing_retries_once() {
    for batched in [false, true] {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 1, 0, "one");
        wb.set_cell_value_tracked(0, 2, 0, "two");
        wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(3)");
        if batched {
            wb.begin_batch();
        }
        wb.clear_cell_tracked(0, 1, 0);
        if !batched {
            assert_eq!(
                wb.active_sheet()
                    .get_cell(0, 0)
                    .spill_error()
                    .unwrap()
                    .blocked_by,
                (2, 0)
            );
        }
        wb.clear_cell_tracked(0, 2, 0);
        if batched {
            wb.end_batch();
        }
        assert_eq!(wb.active_sheet().get_display(2, 0), "3");
        assert!(wb.take_incremental_errors().is_empty());
    }
}

#[test]
fn occupancy_dependency_is_not_a_formula_cycle_and_replacing_parent_removes_it() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 1, "=A1");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(1,2)");
    assert!(wb.dep_graph().find_cycle_members().is_empty());
    assert!(wb.active_sheet().has_spill_error(0, 0));
    wb.clear_cell_tracked(0, 0, 1);
    assert_eq!(wb.active_sheet().get_display(0, 1), "2");

    wb.set_cell_value_tracked(0, 4, 0, "block");
    wb.set_cell_value_tracked(0, 3, 0, "=SEQUENCE(2)");
    wb.set_cell_value_tracked(0, 3, 0, "replacement");
    wb.clear_cell_tracked(0, 4, 0);
    assert_eq!(wb.active_sheet().get_raw(3, 0), "replacement");
    assert!(!wb.active_sheet().is_spill_receiver(4, 0));
}

#[test]
fn shrinking_another_spill_retries_a_parent_blocked_by_a_receiver() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 3, "3");
    wb.set_cell_value_tracked(0, 1, 0, "=SEQUENCE(1,D1)");
    wb.set_cell_value_tracked(0, 0, 2, "=SEQUENCE(2)");
    assert!(wb.active_sheet().has_spill_error(0, 2));
    wb.set_cell_value_tracked(0, 0, 3, "1");
    assert_eq!(wb.active_sheet().get_display(1, 2), "2");
    assert!(!wb.active_sheet().has_spill_error(0, 2));
    assert!(wb.take_incremental_errors().is_empty());
}

#[test]
fn restoring_an_obstruction_retires_receivers_and_retries_after_redo() {
    for batched in [false, true] {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 1, 0, "occupied");
        wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(3)");
        wb.set_cell_value_tracked(0, 0, 2, "=A3*10");
        wb.clear_cell_tracked(0, 1, 0);
        assert_eq!(wb.active_sheet().get_display(0, 2), "30");
        if batched {
            wb.begin_batch();
        }
        // Ordinary value undo restores the previously authored obstruction.
        wb.set_cell_value_tracked(0, 1, 0, "occupied");
        if batched {
            wb.end_batch();
        }
        assert!(wb.active_sheet().has_spill_error(0, 0));
        assert_eq!(wb.active_sheet().get_display(1, 0), "occupied");
        assert_eq!(wb.active_sheet().get_display(2, 0), "");
        assert_eq!(wb.active_sheet().get_display(0, 2), "0");
        assert!(!wb.active_sheet().is_spill_receiver(1, 0));
        wb.clear_cell_tracked(0, 1, 0);
        assert_eq!(wb.active_sheet().get_display(0, 2), "30");
        assert!(wb.take_incremental_errors().is_empty());
    }
}

#[test]
fn deleting_an_array_retries_formulas_blocked_by_its_former_receivers() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 1, 0, "=SEQUENCE(1,3)");
    wb.set_cell_value_tracked(0, 0, 2, "=SEQUENCE(2)");
    wb.set_cell_value_tracked(0, 0, 4, "=B2*10");
    assert!(wb.active_sheet().has_spill_error(0, 2));
    wb.clear_cell_tracked(0, 1, 0);
    assert_eq!(wb.active_sheet().get_display(1, 2), "2");
    assert_eq!(wb.active_sheet().get_display(0, 4), "0");
    assert!(!wb.active_sheet().has_spill_error(0, 2));
    assert!(wb.take_incremental_errors().is_empty());
}
