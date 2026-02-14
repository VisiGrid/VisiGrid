// Build the recon-template.sheet for abuse testing and the demo pipeline.
// Run with: cargo test -p visigrid-cli --test build_recon_template -- --ignored --nocapture
//
// Creates a 2-sheet workbook:
//   tx:      headers in row 1, data target is tx!A1 (with --headers, data lands at A2+)
//   summary: SUMIF formulas, variance check at B7

use visigrid_engine::workbook::Workbook;
use std::path::Path;

fn build_template(out_path: &Path) {
    let mut wb = Workbook::new();

    // Rename default Sheet1 → tx
    let renamed = wb.rename_sheet(0, "tx");
    assert!(renamed, "rename_sheet(0, 'tx') failed");

    // tx (index 0) = transaction data sheet
    wb.set_cell_value_tracked(0, 0, 0, "effective_date"); // A1
    wb.set_cell_value_tracked(0, 0, 1, "posted_date");    // B1
    wb.set_cell_value_tracked(0, 0, 2, "amount_minor");   // C1
    wb.set_cell_value_tracked(0, 0, 3, "currency");       // D1
    wb.set_cell_value_tracked(0, 0, 4, "type");           // E1
    wb.set_cell_value_tracked(0, 0, 5, "source");         // F1
    wb.set_cell_value_tracked(0, 0, 6, "source_id");      // G1
    wb.set_cell_value_tracked(0, 0, 7, "group_id");       // H1
    wb.set_cell_value_tracked(0, 0, 8, "description");    // I1

    // summary (index 1) = invariant formulas
    let si = wb.add_sheet_named("summary").expect("add summary sheet");

    wb.set_cell_value_tracked(si, 0, 0, "Category");
    wb.set_cell_value_tracked(si, 0, 1, "Total (minor units)");

    wb.set_cell_value_tracked(si, 1, 0, "Charges");
    wb.set_cell_value_tracked(si, 2, 0, "Payouts");
    wb.set_cell_value_tracked(si, 3, 0, "Fees");
    wb.set_cell_value_tracked(si, 4, 0, "Refunds");
    wb.set_cell_value_tracked(si, 5, 0, "Adjustments");
    wb.set_cell_value_tracked(si, 6, 0, "Variance");

    // SUMIF: sum amount_minor (C) where type (E) matches
    wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(tx!E2:E1000,\"charge\",tx!C2:C1000)");
    wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(tx!E2:E1000,\"payout\",tx!C2:C1000)");
    wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(tx!E2:E1000,\"fee\",tx!C2:C1000)");
    wb.set_cell_value_tracked(si, 4, 1, "=SUMIF(tx!E2:E1000,\"refund\",tx!C2:C1000)");
    wb.set_cell_value_tracked(si, 5, 1, "=SUMIF(tx!E2:E1000,\"adjustment\",tx!C2:C1000)");

    // Variance = sum all categories (should be 0 for balanced books)
    wb.set_cell_value_tracked(si, 6, 1, "=B2+B3+B4+B5+B6");

    // For test #13: intentional error cell at B8 (=1/0)
    wb.set_cell_value_tracked(si, 7, 0, "Error test");
    wb.set_cell_value_tracked(si, 7, 1, "=1/0");

    // For test #14: string cell at B9
    wb.set_cell_value_tracked(si, 8, 0, "String test");
    wb.set_cell_value_tracked(si, 8, 1, "not a number");

    // For test #12: intentionally blank cell at B10
    wb.set_cell_value_tracked(si, 9, 0, "Blank test");
    // B10 left empty

    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

    let parent = out_path.parent().unwrap();
    std::fs::create_dir_all(parent).expect("create output dir");
    visigrid_io::native::save_workbook(&wb, out_path).expect("save template");

    let fp = visigrid_io::native::compute_semantic_fingerprint(&wb);
    eprintln!("Template built: {}", out_path.display());
    eprintln!("Fingerprint:    {}", fp);
    eprintln!("Sheets:         {}", wb.sheet_count());

    // Verify summary formulas computed to 0 (no data yet)
    let summary = wb.sheet(si).unwrap();
    let var = summary.get_computed_value(6, 1);
    eprintln!("Variance (B7):  {:?}", var);
}

#[test]
#[ignore] // Run explicitly to build the template
fn build_recon_template() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    // Build into tests/abuse/templates (used by abuse_tests)
    build_template(&manifest_dir.join("tests/abuse/templates/recon-template.sheet"));
    // Build into ../../demo/templates (used by demo scripts — relative to app/crates/cli)
    build_template(&manifest_dir.join("../../demo/templates/recon-template.sheet"));
}
