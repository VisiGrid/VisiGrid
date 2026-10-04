//! #95: a workbook containing a reference cycle must save and reload with
//! every member's formula intact, in every format we write.

use visigrid_engine::workbook::Workbook;
use visigrid_io::{json, native, xlsx};

fn cyclic() -> Workbook {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=B1+1");
    wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
    wb.set_cell_value_tracked(0, 1, 0, "4");
    wb
}

fn assert_cycle_intact(wb: &Workbook, label: &str) {
    let sheet = &wb.sheets()[0];
    assert_eq!(sheet.get_raw(0, 0), "=B1+1", "{label}: A1 formula");
    assert_eq!(sheet.get_raw(0, 1), "=A1+1", "{label}: B1 formula");
    assert_eq!(sheet.get_display(0, 0), "#CYCLE!", "{label}: A1 shows the cycle");
    assert_eq!(sheet.get_raw(1, 0), "4", "{label}: unrelated cell");
}

#[test]
fn native_save_and_load_keep_cycle_formulas() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("cycle.sheet");
    let wb = cyclic();
    assert_cycle_intact(&wb, "before save");
    native::save_workbook(&wb, &path).unwrap();
    let loaded = native::load_workbook(&path).unwrap();
    assert_cycle_intact(&loaded, "after native load");
}

#[test]
fn json_export_keeps_cycle_formulas() {
    let wb = cyclic();
    let text = json::export_workbook(&wb, &[Default::default()], 0).unwrap();
    assert!(text.contains("=B1+1") && text.contains("=A1+1"), "formulas written: {text}");
    let (loaded, _, _) = json::import_any_for_recovery(&text).unwrap();
    assert_eq!(loaded.sheets()[0].get_raw(0, 0), "=B1+1");
    assert_eq!(loaded.sheets()[0].get_raw(0, 1), "=A1+1");
}

#[test]
fn xlsx_round_trip_keeps_cycle_formulas() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("cycle.xlsx");
    let wb = cyclic();
    xlsx::export(&wb, &path, None).unwrap();
    let (loaded, _) = xlsx::import(&path).unwrap();
    assert_eq!(loaded.sheets()[0].get_raw(0, 0), "=B1+1");
    assert_eq!(loaded.sheets()[0].get_raw(0, 1), "=A1+1");
}
