//! #95: closing a reference cycle used to replace every member's formula
//! with the literal text "#CYCLE!". Members must keep their formulas, show
//! #CYCLE! only as their computed value, and recompute once the cycle breaks.

use visigrid_engine::structural::Axis;
use visigrid_engine::workbook::Workbook;

fn raw(wb: &Workbook, sheet: usize, row: usize, col: usize) -> String {
    wb.sheets()[sheet].get_raw(row, col)
}

fn shown(wb: &Workbook, sheet: usize, row: usize, col: usize) -> String {
    wb.sheets()[sheet].get_display(row, col)
}

/// The simulator's repro (PR #94), now a regular test.
#[test]
fn closing_a_cycle_keeps_the_formula_text() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1+1");
    wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
    assert_eq!(raw(&wb, 0, 0, 0), "=B1+1", "A1's formula must survive the cycle");
    assert_eq!(raw(&wb, 0, 0, 1), "=A1+1", "B1's formula must survive the cycle");
    assert_eq!(shown(&wb, 0, 0, 0), "#CYCLE!");
    assert_eq!(shown(&wb, 0, 0, 1), "#CYCLE!");
    assert!(wb.sheets()[0].is_cycle_error(0, 0));

    // A later full recompute (a structural edit forces one) keeps them too.
    wb.structural_edit(0, Axis::Row, 5, 1, false).unwrap();
    assert_eq!(raw(&wb, 0, 0, 0), "=B1+1", "and a later full recompute");
    assert_eq!(raw(&wb, 0, 0, 1), "=A1+1");
    assert_eq!(shown(&wb, 0, 0, 0), "#CYCLE!");
}

#[test]
fn breaking_the_cycle_recomputes_the_other_members() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1+1");
    wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
    // Editing one member breaks the cycle; the other member's formula is
    // still there and computes from the new value.
    wb.set_cell_value_tracked(0, 0, 0, "5");
    assert_eq!(shown(&wb, 0, 0, 0), "5");
    assert_eq!(raw(&wb, 0, 0, 1), "=A1+1");
    assert_eq!(shown(&wb, 0, 0, 1), "6");
    assert!(!wb.sheets()[0].is_cycle_error(0, 1));
}

#[test]
fn a_three_member_cycle_keeps_all_formulas_and_recovers() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1");
    wb.set_cell_value_tracked(0, 0, 1, "=C1");
    wb.set_cell_value_tracked(0, 0, 2, "=A1*2");
    for (col, f) in [(0, "=B1"), (1, "=C1"), (2, "=A1*2")] {
        assert_eq!(raw(&wb, 0, 0, col), f);
        assert_eq!(shown(&wb, 0, 0, col), "#CYCLE!");
    }
    wb.set_cell_value_tracked(0, 0, 2, "7");
    assert_eq!(shown(&wb, 0, 0, 1), "7");
    assert_eq!(shown(&wb, 0, 0, 0), "7");
    assert_eq!(raw(&wb, 0, 0, 0), "=B1");
}

#[test]
fn a_cross_sheet_cycle_keeps_both_formulas() {
    let mut wb = Workbook::new();
    let other = wb.add_sheet_named("Other").expect("second sheet");
    wb.set_cell_value_tracked(0, 0, 0, "=Other!A1+1");
    wb.set_cell_value_tracked(other, 0, 0, "=Sheet1!A1+1");
    assert_eq!(raw(&wb, 0, 0, 0), "=Other!A1+1");
    assert_eq!(raw(&wb, other, 0, 0), "=Sheet1!A1+1");
    assert_eq!(shown(&wb, 0, 0, 0), "#CYCLE!");
    assert_eq!(shown(&wb, other, 0, 0), "#CYCLE!");
    wb.set_cell_value_tracked(other, 0, 0, "1");
    assert_eq!(shown(&wb, 0, 0, 0), "2");
}

#[test]
fn a_cell_downstream_of_a_cycle_reports_it_upstream() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1+1");
    wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
    wb.set_cell_value_tracked(0, 0, 2, "=A1*10");
    let id = wb.sheet_id_at_idx(0).unwrap();
    assert!(wb.has_cycle_in_upstream(id, 0, 2), "C1 depends on a cycle");
    assert!(!wb.sheets()[0].is_cycle_error(0, 2), "but is not a member");
    wb.recompute_full_ordered();
    assert!(wb.sheets()[0].is_cycle_error(0, 0));
    assert!(wb.sheets()[0].is_cycle_error(0, 1));
    assert!(!wb.sheets()[0].is_cycle_error(0, 2));
    assert!(wb.has_cycle_in_upstream(id, 0, 2));
}
