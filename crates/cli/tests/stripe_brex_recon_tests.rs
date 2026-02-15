// Integration tests for the Stripe ↔ Brex reconciliation template.
// Run with: cargo test -p visigrid-cli --test stripe_brex_recon_tests -- --nocapture
//
// Tests fill the 3-sheet template with test CSV data and verify computed values.

use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Once;
use visigrid_engine::formula::eval::Value;
use visigrid_engine::workbook::Workbook;
use visigrid_io::native;

static TEMPLATE_INIT: Once = Once::new();

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

fn template_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/recon/templates/stripe-brex-recon.sheet")
}

fn csv_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/recon/csv/{}", name))
}

const CANONICAL_HEADERS: [&str; 9] = [
    "effective_date", "posted_date", "amount_minor", "currency",
    "type", "source", "source_id", "group_id", "description",
];

fn ensure_template() {
    TEMPLATE_INIT.call_once(|| {
        let tmpl = template_path();

        let mut wb = Workbook::new();

        // ── Sheet 0: stripe ──
        wb.rename_sheet(0, "stripe");
        for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
            wb.set_cell_value_tracked(0, 0, col, hdr);
        }
        wb.set_cell_value_tracked(0, 0, 9, "payout_abs");
        wb.set_cell_value_tracked(0, 0, 10, "brex_match");
        wb.set_cell_value_tracked(0, 0, 11, "match_status");

        for r in 1..=1000 {
            let row1 = r + 1;
            wb.set_cell_value_tracked(0, r, 9,
                &format!("=IF(E{row1}=\"payout\",-C{row1},\"\")"));
            wb.set_cell_value_tracked(0, r, 10,
                &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},brex!J$2:J$1001,brex!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
            wb.set_cell_value_tracked(0, r, 11,
                &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
        }

        // ── Sheet 1: brex ──
        let bi = wb.add_sheet_named("brex").unwrap();
        for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
            wb.set_cell_value_tracked(bi, 0, col, hdr);
        }
        wb.set_cell_value_tracked(bi, 0, 9, "deposit_amt");
        wb.set_cell_value_tracked(bi, 0, 10, "stripe_match");
        wb.set_cell_value_tracked(bi, 0, 11, "match_status");

        for r in 1..=1000 {
            let row1 = r + 1;
            wb.set_cell_value_tracked(bi, r, 9,
                &format!("=IF(E{row1}=\"deposit\",C{row1},\"\")"));
            wb.set_cell_value_tracked(bi, r, 10,
                &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},stripe!J$2:J$1001,stripe!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
            wb.set_cell_value_tracked(bi, r, 11,
                &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
        }

        // ── Sheet 2: summary ──
        let si = wb.add_sheet_named("summary").unwrap();

        wb.set_cell_value_tracked(si, 0, 0, "STRIPE BALANCE");
        wb.set_cell_value_tracked(si, 0, 1, "Amount");
        wb.set_cell_value_tracked(si, 0, 2, "Status");

        wb.set_cell_value_tracked(si, 1, 0, "Charges");
        wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(stripe!E$2:E$1001,\"charge\",stripe!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 2, 0, "Fees");
        wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(stripe!E$2:E$1001,\"fee\",stripe!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 3, 0, "Refunds");
        wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(stripe!E$2:E$1001,\"refund\",stripe!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 4, 0, "Payouts");
        wb.set_cell_value_tracked(si, 4, 1, "=SUMIF(stripe!E$2:E$1001,\"payout\",stripe!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 5, 0, "Adjustments");
        wb.set_cell_value_tracked(si, 5, 1, "=SUMIF(stripe!E$2:E$1001,\"adjustment\",stripe!C$2:C$1001)");

        wb.set_cell_value_tracked(si, 6, 0, "Net");
        wb.set_cell_value_tracked(si, 6, 1, "=SUM(B2:B6)");
        wb.set_cell_value_tracked(si, 6, 2, "=IF(B7=0,\"PASS\",\"FAIL\")");

        wb.set_cell_value_tracked(si, 8, 0, "PAYOUT MATCHING");
        wb.set_cell_value_tracked(si, 8, 1, "Amount");
        wb.set_cell_value_tracked(si, 8, 2, "Status");

        wb.set_cell_value_tracked(si, 9, 0, "Stripe Payouts");
        wb.set_cell_value_tracked(si, 9, 1, "=ABS(B5)");
        wb.set_cell_value_tracked(si, 10, 0, "Matched Deposits");
        wb.set_cell_value_tracked(si, 10, 1, "=SUMIF(stripe!L$2:L$1001,\"MATCHED\",stripe!J$2:J$1001)");
        wb.set_cell_value_tracked(si, 11, 0, "Unmatched Payouts");
        wb.set_cell_value_tracked(si, 11, 1, "=SUMIF(stripe!L$2:L$1001,\"UNMATCHED\",stripe!J$2:J$1001)");
        wb.set_cell_value_tracked(si, 12, 0, "Difference");
        wb.set_cell_value_tracked(si, 12, 1, "=B10-B11");
        wb.set_cell_value_tracked(si, 12, 2, "=IF(B13=0,\"PASS\",\"FAIL\")");

        wb.set_cell_value_tracked(si, 14, 0, "MATCH COUNTS");
        wb.set_cell_value_tracked(si, 14, 1, "Count");
        wb.set_cell_value_tracked(si, 15, 0, "Stripe Payouts");
        wb.set_cell_value_tracked(si, 15, 1, "=COUNTIF(stripe!E$2:E$1001,\"payout\")");
        wb.set_cell_value_tracked(si, 16, 0, "Matched");
        wb.set_cell_value_tracked(si, 16, 1, "=COUNTIF(stripe!L$2:L$1001,\"MATCHED\")");
        wb.set_cell_value_tracked(si, 17, 0, "Unmatched");
        wb.set_cell_value_tracked(si, 17, 1, "=COUNTIF(stripe!L$2:L$1001,\"UNMATCHED\")");

        wb.set_cell_value_tracked(si, 19, 0, "BREX UNMATCHED");
        wb.set_cell_value_tracked(si, 19, 1, "Count");
        wb.set_cell_value_tracked(si, 19, 2, "Amount");
        wb.set_cell_value_tracked(si, 20, 0, "Total Deposits");
        wb.set_cell_value_tracked(si, 20, 1, "=COUNTIF(brex!E$2:E$1001,\"deposit\")");
        wb.set_cell_value_tracked(si, 20, 2, "=SUMIF(brex!E$2:E$1001,\"deposit\",brex!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 21, 0, "Matched to Stripe");
        wb.set_cell_value_tracked(si, 21, 1, "=COUNTIF(brex!L$2:L$1001,\"MATCHED\")");
        wb.set_cell_value_tracked(si, 21, 2, "=SUMIF(brex!L$2:L$1001,\"MATCHED\",brex!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 22, 0, "Unmatched");
        wb.set_cell_value_tracked(si, 22, 1, "=COUNTIF(brex!L$2:L$1001,\"UNMATCHED\")");
        wb.set_cell_value_tracked(si, 22, 2, "=SUMIF(brex!L$2:L$1001,\"UNMATCHED\",brex!C$2:C$1001)");

        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        std::fs::create_dir_all(tmpl.parent().unwrap()).ok();
        native::save_workbook(&wb, &tmpl).expect("save template");
    });
}

fn load_and_read_cell(sheet_path: &Path, sheet_name: &str, row: usize, col: usize) -> Value {
    let mut wb = native::load_workbook(sheet_path).expect("load workbook");
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet_id = wb.sheet_id_by_name(sheet_name).expect("find sheet");
    let idx = wb.idx_for_sheet_id(sheet_id).expect("sheet idx");
    wb.sheet(idx).unwrap().get_computed_value(row, col)
}

/// Fill the template with stripe + brex CSVs (two-step fill).
fn fill_template(stripe_csv: &str, brex_csv: &str) -> tempfile::NamedTempFile {
    ensure_template();

    let step1 = tempfile::NamedTempFile::new().unwrap();
    let final_out = tempfile::NamedTempFile::new().unwrap();

    // Step 1: fill stripe data
    let out1 = vgrid()
        .args([
            "fill", template_path().to_str().unwrap(),
            "--csv", csv_path(stripe_csv).to_str().unwrap(),
            "--target", "stripe!A2", "--headers",
            "--out", step1.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(out1.status.success(), "stripe fill failed: {}", String::from_utf8_lossy(&out1.stderr));

    // Step 2: fill brex data
    let out2 = vgrid()
        .args([
            "fill", step1.path().to_str().unwrap(),
            "--csv", csv_path(brex_csv).to_str().unwrap(),
            "--target", "brex!A2", "--headers",
            "--out", final_out.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(out2.status.success(), "brex fill failed: {}", String::from_utf8_lossy(&out2.stderr));

    final_out
}

/// Fill the template with only stripe data (no brex fill step).
fn fill_stripe_only(stripe_csv: &str) -> tempfile::NamedTempFile {
    ensure_template();

    let out = tempfile::NamedTempFile::new().unwrap();

    let result = vgrid()
        .args([
            "fill", template_path().to_str().unwrap(),
            "--csv", csv_path(stripe_csv).to_str().unwrap(),
            "--target", "stripe!A2", "--headers",
            "--out", out.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(result.status.success(), "stripe fill failed: {}", String::from_utf8_lossy(&result.stderr));

    out
}

// ── Test 1: Balanced data, all payouts matched ──

#[test]
fn test_balanced_all_matched() {
    let out = fill_template("stripe-balanced.csv", "brex-balanced.csv");
    let p = out.path();

    // Stripe Net = 0
    let net = load_and_read_cell(p, "summary", 6, 1);
    eprintln!("summary!B7 (Net) = {:?}", net);
    assert!(matches!(net, Value::Number(n) if n == 0.0), "Net should be 0, got {:?}", net);

    // Net status = PASS
    let net_status = load_and_read_cell(p, "summary", 6, 2);
    eprintln!("summary!C7 (Net Status) = {:?}", net_status);
    assert!(matches!(net_status, Value::Text(ref s) if s == "PASS"), "got {:?}", net_status);

    // Difference = 0
    let diff = load_and_read_cell(p, "summary", 12, 1);
    eprintln!("summary!B13 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 0.0), "Difference should be 0, got {:?}", diff);

    // Difference status = PASS
    let diff_status = load_and_read_cell(p, "summary", 12, 2);
    eprintln!("summary!C13 (Difference Status) = {:?}", diff_status);
    assert!(matches!(diff_status, Value::Text(ref s) if s == "PASS"), "got {:?}", diff_status);

    // Stripe Payout count = 2
    let payout_count = load_and_read_cell(p, "summary", 15, 1);
    eprintln!("summary!B16 (Stripe Payouts count) = {:?}", payout_count);
    assert!(matches!(payout_count, Value::Number(n) if n == 2.0), "got {:?}", payout_count);

    // Matched count = 2
    let matched = load_and_read_cell(p, "summary", 16, 1);
    eprintln!("summary!B17 (Matched count) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 2.0), "got {:?}", matched);

    // Unmatched count = 0
    let unmatched = load_and_read_cell(p, "summary", 17, 1);
    eprintln!("summary!B18 (Unmatched count) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 0.0), "got {:?}", unmatched);

    // Brex unmatched count = 1 (CLIENT WIRE PAYMENT has no Stripe payout)
    let brex_unmatched = load_and_read_cell(p, "summary", 22, 1);
    eprintln!("summary!B23 (Brex Unmatched count) = {:?}", brex_unmatched);
    assert!(matches!(brex_unmatched, Value::Number(n) if n == 1.0), "got {:?}", brex_unmatched);

    eprintln!("TEST PASS: balanced_all_matched");
}

// ── Test 2: Missing deposit detected ──

#[test]
fn test_missing_deposit_detected() {
    let out = fill_template("stripe-balanced.csv", "brex-missing-one.csv");
    let p = out.path();

    // Difference = 97100 (the unmatched payout)
    let diff = load_and_read_cell(p, "summary", 12, 1);
    eprintln!("summary!B13 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 97100.0),
        "Difference should be 97100, got {:?}", diff);

    // Difference status = FAIL
    let diff_status = load_and_read_cell(p, "summary", 12, 2);
    eprintln!("summary!C13 (Difference Status) = {:?}", diff_status);
    assert!(matches!(diff_status, Value::Text(ref s) if s == "FAIL"), "got {:?}", diff_status);

    // Matched = 1
    let matched = load_and_read_cell(p, "summary", 16, 1);
    eprintln!("summary!B17 (Matched) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 1.0), "got {:?}", matched);

    // Unmatched = 1
    let unmatched = load_and_read_cell(p, "summary", 17, 1);
    eprintln!("summary!B18 (Unmatched) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 1.0), "got {:?}", unmatched);

    eprintln!("TEST PASS: missing_deposit_detected");
}

// ── Test 3: Empty data produces zero counts ──

#[test]
fn test_empty_data_zero_counts() {
    ensure_template();

    let p = template_path();

    // Net = 0 (no data)
    let net = load_and_read_cell(&p, "summary", 6, 1);
    eprintln!("summary!B7 (Net) = {:?}", net);
    assert!(matches!(net, Value::Number(n) if n == 0.0), "got {:?}", net);

    // Payout count = 0
    let count = load_and_read_cell(&p, "summary", 15, 1);
    eprintln!("summary!B16 (Stripe Payouts count) = {:?}", count);
    assert!(matches!(count, Value::Number(n) if n == 0.0), "got {:?}", count);

    // Matched = 0
    let matched = load_and_read_cell(&p, "summary", 16, 1);
    eprintln!("summary!B17 (Matched) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 0.0), "got {:?}", matched);

    eprintln!("TEST PASS: empty_data_zero_counts");
}

// ── Test 4: Stripe only, all payouts unmatched ──

#[test]
fn test_stripe_only_all_unmatched() {
    let out = fill_stripe_only("stripe-balanced.csv");
    let p = out.path();

    // Unmatched = 2 (both payouts have no Brex match)
    let unmatched = load_and_read_cell(p, "summary", 17, 1);
    eprintln!("summary!B18 (Unmatched) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 2.0), "got {:?}", unmatched);

    // Difference = total payout amounts (77750 + 97100 = 174850)
    let diff = load_and_read_cell(p, "summary", 12, 1);
    eprintln!("summary!B13 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 174850.0),
        "Difference should be 174850, got {:?}", diff);

    eprintln!("TEST PASS: stripe_only_all_unmatched");
}
