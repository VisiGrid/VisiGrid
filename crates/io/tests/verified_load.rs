//! `import_any_verified` skips only the final content-protection diff, so for
//! content `import_any` loads complete it must build the identical workbook.
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::{self, SheetLayout};

fn document() -> String {
    let mut wb = Workbook::new();
    for (row, value) in ["10", "20", "=A1+A2", "=SUM(A1:A3)", "text"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    wb.add_sheet_named("Second").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=Sheet1!A4*2");
    let layouts = vec![SheetLayout::default(); wb.sheet_count()];
    json::export_workbook(&wb, &layouts, 0).unwrap()
}

fn reexport(loaded: &(Workbook, Vec<SheetLayout>, usize)) -> String {
    json::export_workbook(&loaded.0, &loaded.1, loaded.2).unwrap()
}

#[test]
fn verified_load_builds_the_same_workbook_as_a_full_load() {
    let doc = document();
    let full = json::import_any(&doc).unwrap();
    assert!(json::loaded_complete(&full.0));
    let verified = json::import_any_verified(&doc).unwrap();
    assert!(json::loaded_complete(&verified.0));
    assert_eq!(reexport(&verified), reexport(&full));
    assert_eq!(verified.2, full.2);
}

#[test]
fn a_lossy_document_is_not_complete_so_it_is_never_verified() {
    let mut doc: serde_json::Value = serde_json::from_str(&document()).unwrap();
    doc["future_feature"] = serde_json::json!([1, 2, 3]);
    let full = json::import_any(&doc.to_string()).unwrap();
    assert!(!json::loaded_complete(&full.0), "content protection opens it read-only");
}

#[test]
fn other_read_only_paths_still_run_on_a_verified_load() {
    let mut doc: serde_json::Value = serde_json::from_str(&document()).unwrap();
    doc["version"] = serde_json::json!(999);
    let verified = json::import_any_verified(&doc.to_string()).unwrap();
    assert!(!json::loaded_complete(&verified.0), "a newer format still opens read-only");
}

fn reasons(wb: &Workbook) -> Vec<Option<String>> {
    wb.sheets().iter().map(|s| s.read_only_reason.clone()).collect()
}

/// Draw first, check after: `protect_loaded` on a verified load must end
/// exactly where `import_any` does, complete or not.
#[test]
fn checking_a_verified_load_later_ends_where_a_full_load_does() {
    let complete = document();
    let mut lossy: serde_json::Value = serde_json::from_str(&complete).unwrap();
    lossy["future_feature"] = serde_json::json!([1, 2, 3]);
    // A stored formula value the recompute disagrees with: read-only keeps it.
    lossy["sheets"][0]["cells"].as_array_mut().unwrap().iter_mut()
        .find(|c| c["formula"] == "=A1+A2").unwrap()["value"] = serde_json::json!(999);
    for (doc, expect_complete) in [(complete, true), (lossy.to_string(), false)] {
        let full = json::import_any(&doc).unwrap();
        let mut later = json::import_any_verified(&doc).unwrap();
        assert!(json::loaded_complete(&later.0), "nothing is checked yet");
        let (wb, layouts, active) = (&mut later.0, &later.1, later.2);
        assert_eq!(json::protect_loaded(wb, layouts, active, &doc).unwrap(), expect_complete);
        assert_eq!(json::loaded_complete(&later.0), expect_complete);
        assert_eq!(reasons(&later.0), reasons(&full.0));
        assert_eq!(reexport(&later), reexport(&full));
        assert_eq!(later.0.sheet(0).unwrap().get_display(2, 0), full.0.sheet(0).unwrap().get_display(2, 0));
    }
}

#[test]
fn checking_a_workbook_that_already_opened_read_only_changes_nothing() {
    let mut doc: serde_json::Value = serde_json::from_str(&document()).unwrap();
    doc["version"] = serde_json::json!(999);
    let doc = doc.to_string();
    let mut loaded = json::import_any_verified(&doc).unwrap();
    let before = reasons(&loaded.0);
    let (wb, layouts, active) = (&mut loaded.0, &loaded.1, loaded.2);
    assert!(!json::protect_loaded(wb, layouts, active, &doc).unwrap());
    assert_eq!(reasons(&loaded.0), before);
}
