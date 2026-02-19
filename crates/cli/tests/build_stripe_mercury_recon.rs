// Build the stripe-mercury-recon.sheet template.
// Run with: cargo test -p visigrid-cli --test build_stripe_mercury_recon -- --ignored --nocapture
//
// Creates a 3-sheet workbook:
//   stripe:  headers A1-L1, matching formulas in J-L for rows 2-1001
//            + rollup cols M-N, fee audit cols O-T for rows 2-1001
//   mercury: headers A1-L1, matching formulas in J-L for rows 2-1001
//   summary: aggregate checks (balance, payout matching, counts,
//            rollup integrity, fee audit, overall verdict)

use std::path::Path;
use visigrid_engine::workbook::Workbook;

// Headers as formulas so they survive `vgrid fill --clear` (which removes non-formula cells).
const CANONICAL_HEADERS: [&str; 9] = [
    "=\"effective_date\"", "=\"posted_date\"", "=\"amount_minor\"", "=\"currency\"",
    "=\"type\"", "=\"source\"", "=\"source_id\"", "=\"group_id\"", "=\"description\"",
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
    // Matching column headers (J1-L1) — formulas to survive --clear
    wb.set_cell_value_tracked(0, 0, 9, "=\"payout_abs\"");
    wb.set_cell_value_tracked(0, 0, 10, "=\"mercury_match\"");
    wb.set_cell_value_tracked(0, 0, 11, "=\"match_status\"");

    // Rollup + fee audit column headers (M1-T1)
    wb.set_cell_value_tracked(0, 0, 12, "=\"rollup_sum\"");
    wb.set_cell_value_tracked(0, 0, 13, "=\"rollup_check\"");
    wb.set_cell_value_tracked(0, 0, 14, "=\"charge_total\"");
    wb.set_cell_value_tracked(0, 0, 15, "=\"fee_total\"");
    wb.set_cell_value_tracked(0, 0, 16, "=\"charge_count\"");
    wb.set_cell_value_tracked(0, 0, 17, "=\"expected_fee\"");
    wb.set_cell_value_tracked(0, 0, 18, "=\"fee_variance\"");
    wb.set_cell_value_tracked(0, 0, 19, "=\"fee_check\"");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1; // 1-indexed row number for formula references
        // J: payout amount (negated to positive, in-period payouts only)
        wb.set_cell_value_tracked(0, r, 9,
            &format!("=IF(AND(E{row1}=\"payout\",OR(summary!B$42=\"\",A{row1}>=summary!B$42)),-C{row1},\"\")"));
        // K: XLOOKUP matching deposit by amount
        wb.set_cell_value_tracked(0, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},mercury!J$2:J$1001,mercury!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(0, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));

        // M: rollup sum — sum all amounts in same group_id
        wb.set_cell_value_tracked(0, r, 12,
            &format!("=IF(H{row1}=\"\",\"\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1}))"));
        // N: rollup check — in-period payout rows only
        wb.set_cell_value_tracked(0, r, 13,
            &format!("=IF(AND(E{row1}=\"payout\",OR(summary!B$42=\"\",A{row1}>=summary!B$42)),IF(M{row1}=0,\"OK\",\"FAIL\"),\"\")"));
        // O: charge total per payout
        wb.set_cell_value_tracked(0, r, 14,
            &format!("=IF(E{row1}=\"payout\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1},E$2:E$1001,\"charge\"),\"\")"));
        // P: fee total per payout
        wb.set_cell_value_tracked(0, r, 15,
            &format!("=IF(E{row1}=\"payout\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1},E$2:E$1001,\"fee\"),\"\")"));
        // Q: charge count per payout
        wb.set_cell_value_tracked(0, r, 16,
            &format!("=IF(E{row1}=\"payout\",COUNTIFS(H$2:H$1001,H{row1},E$2:E$1001,\"charge\"),\"\")"));
        // R: expected fee (rate × charges + per-txn × count)
        wb.set_cell_value_tracked(0, r, 17,
            &format!("=IF(E{row1}=\"payout\",-(O{row1}*summary!B$32/100+Q{row1}*summary!B$33),\"\")"));
        // S: fee variance (actual − expected)
        wb.set_cell_value_tracked(0, r, 18,
            &format!("=IF(E{row1}=\"payout\",P{row1}-R{row1},\"\")"));
        // T: fee check (within 1¢ tolerance)
        wb.set_cell_value_tracked(0, r, 19,
            &format!("=IF(E{row1}=\"payout\",IF(ABS(S{row1})<=1,\"OK\",\"REVIEW\"),\"\")"));
    }

    // ── Sheet 1: mercury ──
    let mi = wb.add_sheet_named("mercury").expect("add mercury sheet");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(mi, 0, col, hdr);
    }
    // Matching column headers (J1-L1) — formulas to survive --clear
    wb.set_cell_value_tracked(mi, 0, 9, "=\"deposit_amt\"");
    wb.set_cell_value_tracked(mi, 0, 10, "=\"stripe_match\"");
    wb.set_cell_value_tracked(mi, 0, 11, "=\"match_status\"");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1;
        // J: deposit amount (only for deposits)
        wb.set_cell_value_tracked(mi, r, 9,
            &format!("=IF(E{row1}=\"deposit\",C{row1},\"\")"));
        // K: XLOOKUP matching payout by amount
        wb.set_cell_value_tracked(mi, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},stripe!J$2:J$1001,stripe!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(mi, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 2: summary ──
    let si = wb.add_sheet_named("summary").expect("add summary sheet");

    // Row 1: STRIPE BALANCE header — all labels as formulas to survive --clear
    wb.set_cell_value_tracked(si, 0, 0, "=\"STRIPE BALANCE\"");
    wb.set_cell_value_tracked(si, 0, 1, "=\"Amount\"");
    wb.set_cell_value_tracked(si, 0, 2, "=\"Status\"");

    // Rows 2-6: category breakdowns
    wb.set_cell_value_tracked(si, 1, 0, "=\"Charges\"");
    wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(stripe!E$2:E$1001,\"charge\",stripe!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 2, 0, "=\"Fees\"");
    wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(stripe!E$2:E$1001,\"fee\",stripe!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 3, 0, "=\"Refunds\"");
    wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(stripe!E$2:E$1001,\"refund\",stripe!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 4, 0, "=\"Payouts\"");
    wb.set_cell_value_tracked(si, 4, 1, "=SUMIF(stripe!E$2:E$1001,\"payout\",stripe!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 5, 0, "=\"Adjustments\"");
    wb.set_cell_value_tracked(si, 5, 1, "=SUMIF(stripe!E$2:E$1001,\"adjustment\",stripe!C$2:C$1001)");

    // Row 7: Net (should be 0 for balanced books)
    wb.set_cell_value_tracked(si, 6, 0, "=\"Net\"");
    wb.set_cell_value_tracked(si, 6, 1, "=SUM(B2:B6)");
    wb.set_cell_value_tracked(si, 6, 2, "=IF(B7=0,\"PASS\",\"FAIL\")");

    // Row 8: blank separator

    // Row 9: PAYOUT MATCHING header
    wb.set_cell_value_tracked(si, 8, 0, "=\"PAYOUT MATCHING\"");
    wb.set_cell_value_tracked(si, 8, 1, "=\"Amount\"");
    wb.set_cell_value_tracked(si, 8, 2, "=\"Status\"");

    // Rows 10-13: payout matching
    wb.set_cell_value_tracked(si, 9, 0, "=\"Stripe Payouts\"");
    wb.set_cell_value_tracked(si, 9, 1, "=SUM(stripe!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 10, 0, "=\"Matched Deposits\"");
    wb.set_cell_value_tracked(si, 10, 1, "=SUMIF(stripe!L$2:L$1001,\"MATCHED\",stripe!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 11, 0, "=\"Unmatched Payouts\"");
    wb.set_cell_value_tracked(si, 11, 1, "=SUMIF(stripe!L$2:L$1001,\"UNMATCHED\",stripe!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 12, 0, "=\"Difference\"");
    wb.set_cell_value_tracked(si, 12, 1, "=B10-B11");
    wb.set_cell_value_tracked(si, 12, 2, "=IF(B13=0,\"PASS\",\"FAIL\")");

    // Row 14: blank separator

    // Row 15: MATCH COUNTS header
    wb.set_cell_value_tracked(si, 14, 0, "=\"MATCH COUNTS\"");
    wb.set_cell_value_tracked(si, 14, 1, "=\"Count\"");

    wb.set_cell_value_tracked(si, 15, 0, "=\"Stripe Payouts\"");
    wb.set_cell_value_tracked(si, 15, 1, "=COUNTIF(stripe!E$2:E$1001,\"payout\")");

    wb.set_cell_value_tracked(si, 16, 0, "=\"Matched\"");
    wb.set_cell_value_tracked(si, 16, 1, "=COUNTIF(stripe!L$2:L$1001,\"MATCHED\")");

    wb.set_cell_value_tracked(si, 17, 0, "=\"Unmatched\"");
    wb.set_cell_value_tracked(si, 17, 1, "=COUNTIF(stripe!L$2:L$1001,\"UNMATCHED\")");

    // Row 19: blank separator

    // Row 20: MERCURY UNMATCHED header
    wb.set_cell_value_tracked(si, 19, 0, "=\"MERCURY UNMATCHED\"");
    wb.set_cell_value_tracked(si, 19, 1, "=\"Count\"");
    wb.set_cell_value_tracked(si, 19, 2, "=\"Amount\"");

    wb.set_cell_value_tracked(si, 20, 0, "=\"Total Deposits\"");
    wb.set_cell_value_tracked(si, 20, 1, "=COUNTIF(mercury!E$2:E$1001,\"deposit\")");
    wb.set_cell_value_tracked(si, 20, 2, "=SUMIF(mercury!E$2:E$1001,\"deposit\",mercury!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 21, 0, "=\"Matched to Stripe\"");
    wb.set_cell_value_tracked(si, 21, 1, "=COUNTIF(mercury!L$2:L$1001,\"MATCHED\")");
    wb.set_cell_value_tracked(si, 21, 2, "=SUMIF(mercury!L$2:L$1001,\"MATCHED\",mercury!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 22, 0, "=\"Unmatched\"");
    wb.set_cell_value_tracked(si, 22, 1, "=COUNTIF(mercury!L$2:L$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 22, 2, "=SUMIF(mercury!L$2:L$1001,\"UNMATCHED\",mercury!C$2:C$1001)");

    // Row 25: ROLLUP INTEGRITY header (0-indexed row 24)
    wb.set_cell_value_tracked(si, 24, 0, "=\"ROLLUP INTEGRITY\"");
    wb.set_cell_value_tracked(si, 24, 1, "=\"Count\"");
    wb.set_cell_value_tracked(si, 24, 2, "=\"Status\"");

    wb.set_cell_value_tracked(si, 25, 0, "=\"Payouts Checked\"");
    wb.set_cell_value_tracked(si, 25, 1, "=COUNTIF(stripe!N$2:N$1001,\"OK\")+COUNTIF(stripe!N$2:N$1001,\"FAIL\")");

    wb.set_cell_value_tracked(si, 26, 0, "=\"Passed\"");
    wb.set_cell_value_tracked(si, 26, 1, "=COUNTIF(stripe!N$2:N$1001,\"OK\")");

    wb.set_cell_value_tracked(si, 27, 0, "=\"Failed\"");
    wb.set_cell_value_tracked(si, 27, 1, "=COUNTIF(stripe!N$2:N$1001,\"FAIL\")");

    wb.set_cell_value_tracked(si, 28, 0, "=\"Rollup Status\"");
    wb.set_cell_value_tracked(si, 28, 2, "=IF(B28=0,\"PASS\",\"FAIL\")");

    // Row 31: FEE AUDIT header (0-indexed row 30)
    wb.set_cell_value_tracked(si, 30, 0, "=\"FEE AUDIT\"");
    wb.set_cell_value_tracked(si, 30, 1, "=\"Amount\"");
    wb.set_cell_value_tracked(si, 30, 2, "=\"Status\"");

    wb.set_cell_value_tracked(si, 31, 0, "=\"Contract Rate (%)\"");
    wb.set_cell_value_tracked(si, 31, 1, "=2.90");

    wb.set_cell_value_tracked(si, 32, 0, "=\"Per-Txn Fee (¢)\"");
    wb.set_cell_value_tracked(si, 32, 1, "=30");

    wb.set_cell_value_tracked(si, 33, 0, "=\"Payouts Checked\"");
    wb.set_cell_value_tracked(si, 33, 1, "=COUNTIF(stripe!T$2:T$1001,\"OK\")+COUNTIF(stripe!T$2:T$1001,\"REVIEW\")");

    wb.set_cell_value_tracked(si, 34, 0, "=\"Within Tolerance\"");
    wb.set_cell_value_tracked(si, 34, 1, "=COUNTIF(stripe!T$2:T$1001,\"OK\")");

    wb.set_cell_value_tracked(si, 35, 0, "=\"Needs Review\"");
    wb.set_cell_value_tracked(si, 35, 1, "=COUNTIF(stripe!T$2:T$1001,\"REVIEW\")");

    wb.set_cell_value_tracked(si, 36, 0, "=\"Fee Status\"");
    wb.set_cell_value_tracked(si, 36, 2, "=IF(B36=0,\"OK\",\"REVIEW\")");

    // Row 39: OVERALL VERDICT header (0-indexed row 38)
    wb.set_cell_value_tracked(si, 38, 0, "=\"OVERALL VERDICT\"");
    wb.set_cell_value_tracked(si, 38, 2, "=\"Status\"");

    wb.set_cell_value_tracked(si, 39, 2, "=IF(AND(C13=\"PASS\",C29=\"PASS\"),\"PASS\",\"FAIL\")");

    // Row 42: Period Start (0-indexed row 41) — filled by workflow for date-window filtering
    wb.set_cell_value_tracked(si, 41, 0, "=\"Period Start\"");

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
fn build_stripe_mercury_recon_template() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    build_template(&manifest_dir.join("tests/recon/templates/stripe-mercury-recon.sheet"));
    build_template(&manifest_dir.join("../../demo/templates/stripe-mercury-recon.sheet"));
}
