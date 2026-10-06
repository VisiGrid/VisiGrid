//! Engine behaviour the simulator found that breaks convergence. Not fixed
//! here (engine internals are out of this crate's scope); each test pins a
//! minimal repro and is ignored so CI stays green until the engine changes.

use visigrid_engine::structural::Axis;
use visigrid_engine::workbook::Workbook;

/// Closing a reference cycle forces a full recompute, and the full
/// recompute marks every cycle member with `Sheet::set_cycle_error`, which
/// replaces the formula *text* with the literal text "#CYCLE!". The user's
/// formulas are gone (the cycle can never be repaired by editing one cell),
/// and which formulas are lost depends on the order edits are applied in,
/// so replicas diverge. Expected: keep the formula, show #CYCLE! as its
/// computed value (as for other errors).
#[test]
#[ignore = "engine bug: closing a cycle replaces the members' formula text with \"#CYCLE!\""]
fn closing_a_cycle_keeps_the_formula_text() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1+1");
    wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
    assert_eq!(
        wb.sheets()[0].get_raw(0, 0),
        "=B1+1",
        "A1's formula must survive the cycle"
    );
    assert_eq!(
        wb.sheets()[0].get_raw(0, 1),
        "=A1+1",
        "B1's formula must survive the cycle"
    );
    wb.structural_edit(0, Axis::Row, 5, 1, false).unwrap();
    assert_eq!(
        wb.sheets()[0].get_raw(0, 0),
        "=B1+1",
        "and a later full recompute"
    );
}
