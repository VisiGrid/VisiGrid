// Build the stripe-qbo-recon.sheet template.
// Run with: cargo test -p visigrid-cli --test build_stripe_qbo_recon -- --ignored --nocapture
//
// Creates a 3-sheet workbook:
//   stripe:  headers A1-L1, matching formulas in J-L for rows 2-1001
//   qbo:     headers A1-L1, matching formulas in J-L for rows 2-1001
//   summary: aggregate checks (balance, payout matching, counts)
//
// Stripe payouts (type="payout") are matched against QBO deposits
// (type="deposit") by amount. The summary sheet tracks:
//   - Stripe internal balance (charges + fees + refunds + payouts + adjustments = 0)
//   - Payout-to-deposit matching (processor total vs bank total)
//   - Match counts and unreconciled transactions

use std::path::Path;
use visigrid_engine::workbook::Workbook;

const CANONICAL_HEADERS: [&str; 9] = [
    "effective_date", "posted_date", "amount_minor", "currency",
    "type", "source", "source_id", "group_id", "description",
];

fn build_template(out_path: &Path) {
    let mut wb = Workbook::new();

    // ── Sheet 0: stripe ──
    let renamed = wb.rename_sheet(0, "stripe");
    assert!(renamed, "rename_sheet(0, 'stripe') failed");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, col, hdr);
    }
    // Matching column headers (J1-L1)
    wb.set_cell_value_tracked(0, 0, 9, "payout_abs");
    wb.set_cell_value_tracked(0, 0, 10, "qbo_match");
    wb.set_cell_value_tracked(0, 0, 11, "match_status");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1; // 1-indexed row number for formula references
        // J: payout amount (negated to positive)
        wb.set_cell_value_tracked(0, r, 9,
            &format!("=IF(E{row1}=\"payout\",-C{row1},\"\")"));
        // K: XLOOKUP matching deposit by amount in qbo sheet
        wb.set_cell_value_tracked(0, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},qbo!J$2:J$1001,qbo!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(0, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 1: qbo ──
    let qi = wb.add_sheet_named("qbo").expect("add qbo sheet");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(qi, 0, col, hdr);
    }
    // Matching column headers (J1-L1)
    wb.set_cell_value_tracked(qi, 0, 9, "deposit_amt");
    wb.set_cell_value_tracked(qi, 0, 10, "stripe_match");
    wb.set_cell_value_tracked(qi, 0, 11, "match_status");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1;
        // J: deposit amount (only for deposits — these correspond to Stripe payouts)
        wb.set_cell_value_tracked(qi, r, 9,
            &format!("=IF(E{row1}=\"deposit\",C{row1},\"\")"));
        // K: XLOOKUP matching payout by amount in stripe sheet
        wb.set_cell_value_tracked(qi, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},stripe!J$2:J$1001,stripe!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(qi, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 2: summary ──
    let si = wb.add_sheet_named("summary").expect("add summary sheet");

    // Row 1: STRIPE BALANCE header
    wb.set_cell_value_tracked(si, 0, 0, "STRIPE BALANCE");
    wb.set_cell_value_tracked(si, 0, 1, "Amount");
    wb.set_cell_value_tracked(si, 0, 2, "Status");

    // Rows 2-6: category breakdowns (Stripe internal balance check)
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

    // Row 7: Net (should be 0 for balanced books — Stripe internal invariant)
    wb.set_cell_value_tracked(si, 6, 0, "Net");
    wb.set_cell_value_tracked(si, 6, 1, "=SUM(B2:B6)");
    wb.set_cell_value_tracked(si, 6, 2, "=IF(B7=0,\"PASS\",\"FAIL\")");

    // Row 8: blank separator

    // Row 9: PAYOUT MATCHING header
    wb.set_cell_value_tracked(si, 8, 0, "PAYOUT MATCHING");
    wb.set_cell_value_tracked(si, 8, 1, "Amount");
    wb.set_cell_value_tracked(si, 8, 2, "Status");

    // Rows 10-13: payout → deposit matching (the core reconciliation)
    wb.set_cell_value_tracked(si, 9, 0, "Stripe Payouts");
    wb.set_cell_value_tracked(si, 9, 1, "=ABS(B5)");

    wb.set_cell_value_tracked(si, 10, 0, "Matched Deposits");
    wb.set_cell_value_tracked(si, 10, 1, "=SUMIF(stripe!L$2:L$1001,\"MATCHED\",stripe!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 11, 0, "Unmatched Payouts");
    wb.set_cell_value_tracked(si, 11, 1, "=SUMIF(stripe!L$2:L$1001,\"UNMATCHED\",stripe!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 12, 0, "Difference");
    wb.set_cell_value_tracked(si, 12, 1, "=B10-B11");
    wb.set_cell_value_tracked(si, 12, 2, "=IF(B13=0,\"PASS\",\"FAIL\")");

    // Row 14: blank separator

    // Row 15: MATCH COUNTS header
    wb.set_cell_value_tracked(si, 14, 0, "MATCH COUNTS");
    wb.set_cell_value_tracked(si, 14, 1, "Count");

    wb.set_cell_value_tracked(si, 15, 0, "Stripe Payouts");
    wb.set_cell_value_tracked(si, 15, 1, "=COUNTIF(stripe!E$2:E$1001,\"payout\")");

    wb.set_cell_value_tracked(si, 16, 0, "Matched");
    wb.set_cell_value_tracked(si, 16, 1, "=COUNTIF(stripe!L$2:L$1001,\"MATCHED\")");

    wb.set_cell_value_tracked(si, 17, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 17, 1, "=COUNTIF(stripe!L$2:L$1001,\"UNMATCHED\")");

    // Row 19: blank separator

    // Row 20: QBO UNMATCHED header
    wb.set_cell_value_tracked(si, 19, 0, "QBO UNMATCHED");
    wb.set_cell_value_tracked(si, 19, 1, "Count");
    wb.set_cell_value_tracked(si, 19, 2, "Amount");

    wb.set_cell_value_tracked(si, 20, 0, "Total Deposits");
    wb.set_cell_value_tracked(si, 20, 1, "=COUNTIF(qbo!E$2:E$1001,\"deposit\")");
    wb.set_cell_value_tracked(si, 20, 2, "=SUMIF(qbo!E$2:E$1001,\"deposit\",qbo!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 21, 0, "Matched to Stripe");
    wb.set_cell_value_tracked(si, 21, 1, "=COUNTIF(qbo!L$2:L$1001,\"MATCHED\")");
    wb.set_cell_value_tracked(si, 21, 2, "=SUMIF(qbo!L$2:L$1001,\"MATCHED\",qbo!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 22, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 22, 1, "=COUNTIF(qbo!L$2:L$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 22, 2, "=SUMIF(qbo!L$2:L$1001,\"UNMATCHED\",qbo!C$2:C$1001)");

    // ── Build and save ──
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

    let parent = out_path.parent().unwrap();
    std::fs::create_dir_all(parent).expect("create output dir");
    visigrid_io::native::save_workbook(&wb, out_path).expect("save template");

    let fp = visigrid_io::native::compute_semantic_fingerprint(&wb);
    eprintln!("Template built: {}", out_path.display());
    eprintln!("Fingerprint:    {}", fp);
    eprintln!("Sheets:         {}", wb.sheet_count());
}

#[test]
#[ignore] // Run explicitly to build the template
fn build_stripe_qbo_recon_template() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    build_template(&manifest_dir.join("tests/recon/templates/stripe-qbo-recon.sheet"));
    build_template(&manifest_dir.join("../../demo/templates/stripe-qbo-recon.sheet"));
}
