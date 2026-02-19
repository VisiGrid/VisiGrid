// Build the stripe-qbo-mercury-recon.sheet template (3-way).
// Run with: cargo test -p visigrid-cli --test build_stripe_qbo_mercury_recon -- --ignored --nocapture
//
// Creates a 4-sheet workbook:
//   stripe:  headers A1-I1, matching formulas against both mercury (J-L) and qbo (M-O)
//            + rollup cols P-Q, fee audit cols R-W for rows 2-1001
//   qbo:     headers A1-I1, matching formulas against stripe (J-L) for rows 2-1001
//   mercury: headers A1-I1, matching formulas against stripe (J-L) for rows 2-1001
//   summary: aggregate checks (balance, payout matching, counts,
//            rollup integrity, fee audit, overall verdict)

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
    // Mercury matching column headers (J1-L1)
    wb.set_cell_value_tracked(0, 0, 9, "payout_abs");
    wb.set_cell_value_tracked(0, 0, 10, "mercury_match");
    wb.set_cell_value_tracked(0, 0, 11, "mercury_status");
    // QBO matching column headers (M1-O1)
    wb.set_cell_value_tracked(0, 0, 12, "qbo_match");
    wb.set_cell_value_tracked(0, 0, 13, "qbo_status");
    wb.set_cell_value_tracked(0, 0, 14, "match_status");

    // Rollup + fee audit column headers (P1-W1)
    wb.set_cell_value_tracked(0, 0, 15, "rollup_sum");
    wb.set_cell_value_tracked(0, 0, 16, "rollup_check");
    wb.set_cell_value_tracked(0, 0, 17, "charge_total");
    wb.set_cell_value_tracked(0, 0, 18, "fee_total");
    wb.set_cell_value_tracked(0, 0, 19, "charge_count");
    wb.set_cell_value_tracked(0, 0, 20, "expected_fee");
    wb.set_cell_value_tracked(0, 0, 21, "fee_variance");
    wb.set_cell_value_tracked(0, 0, 22, "fee_check");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1; // 1-indexed row number for formula references

        // J: payout amount (negated to positive, in-period payouts only)
        wb.set_cell_value_tracked(0, r, 9,
            &format!("=IF(AND(E{row1}=\"payout\",OR(summary!B$48=\"\",A{row1}>=summary!B$48)),-C{row1},\"\")"));

        // K: XLOOKUP matching mercury deposit by amount
        wb.set_cell_value_tracked(0, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},mercury!J$2:J$1001,mercury!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: mercury match status
        wb.set_cell_value_tracked(0, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));

        // M: XLOOKUP matching qbo deposit by amount
        wb.set_cell_value_tracked(0, r, 12,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},qbo!J$2:J$1001,qbo!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // N: qbo match status
        wb.set_cell_value_tracked(0, r, 13,
            &format!("=IF(J{row1}=\"\",\"\",IF(M{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(M{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));

        // O: overall 3-way match status
        wb.set_cell_value_tracked(0, r, 14,
            &format!("=IF(J{row1}=\"\",\"\",IF(AND(L{row1}=\"MATCHED\",N{row1}=\"MATCHED\"),\"MATCHED\",IF(OR(L{row1}=\"MATCHED\",N{row1}=\"MATCHED\"),\"PARTIAL\",\"UNMATCHED\")))"));

        // P: rollup sum — sum all amounts in same group_id
        wb.set_cell_value_tracked(0, r, 15,
            &format!("=IF(H{row1}=\"\",\"\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1}))"));
        // Q: rollup check — in-period payout rows only
        wb.set_cell_value_tracked(0, r, 16,
            &format!("=IF(AND(E{row1}=\"payout\",OR(summary!B$48=\"\",A{row1}>=summary!B$48)),IF(P{row1}=0,\"OK\",\"FAIL\"),\"\")"));
        // R: charge total per payout
        wb.set_cell_value_tracked(0, r, 17,
            &format!("=IF(E{row1}=\"payout\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1},E$2:E$1001,\"charge\"),\"\")"));
        // S: fee total per payout
        wb.set_cell_value_tracked(0, r, 18,
            &format!("=IF(E{row1}=\"payout\",SUMIFS(C$2:C$1001,H$2:H$1001,H{row1},E$2:E$1001,\"fee\"),\"\")"));
        // T: charge count per payout
        wb.set_cell_value_tracked(0, r, 19,
            &format!("=IF(E{row1}=\"payout\",COUNTIFS(H$2:H$1001,H{row1},E$2:E$1001,\"charge\"),\"\")"));
        // U: expected fee (rate × charges + per-txn × count)
        wb.set_cell_value_tracked(0, r, 20,
            &format!("=IF(E{row1}=\"payout\",-(R{row1}*summary!B$38/100+T{row1}*summary!B$39),\"\")"));
        // V: fee variance (actual − expected)
        wb.set_cell_value_tracked(0, r, 21,
            &format!("=IF(E{row1}=\"payout\",S{row1}-U{row1},\"\")"));
        // W: fee check (within 1¢ tolerance)
        wb.set_cell_value_tracked(0, r, 22,
            &format!("=IF(E{row1}=\"payout\",IF(ABS(V{row1})<=1,\"OK\",\"REVIEW\"),\"\")"));
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

    // Formulas for rows 2-1001
    for r in 1..=1000 {
        let row1 = r + 1;
        // J: deposit amount (only for deposits)
        wb.set_cell_value_tracked(qi, r, 9,
            &format!("=IF(E{row1}=\"deposit\",C{row1},\"\")"));
        // K: XLOOKUP matching stripe payout by amount
        wb.set_cell_value_tracked(qi, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},stripe!J$2:J$1001,stripe!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(qi, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 2: mercury ──
    let mi = wb.add_sheet_named("mercury").expect("add mercury sheet");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(mi, 0, col, hdr);
    }
    // Matching column headers (J1-L1)
    wb.set_cell_value_tracked(mi, 0, 9, "deposit_amt");
    wb.set_cell_value_tracked(mi, 0, 10, "stripe_match");
    wb.set_cell_value_tracked(mi, 0, 11, "match_status");

    // Formulas for rows 2-1001
    for r in 1..=1000 {
        let row1 = r + 1;
        // J: deposit amount (only for deposits)
        wb.set_cell_value_tracked(mi, r, 9,
            &format!("=IF(E{row1}=\"deposit\",C{row1},\"\")"));
        // K: XLOOKUP matching stripe payout by amount
        wb.set_cell_value_tracked(mi, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",IFERROR(XLOOKUP(J{row1},stripe!J$2:J$1001,stripe!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // L: status label
        wb.set_cell_value_tracked(mi, r, 11,
            &format!("=IF(J{row1}=\"\",\"\",IF(K{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(K{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 3: summary ──
    let si = wb.add_sheet_named("summary").expect("add summary sheet");

    // Row 1: STRIPE BALANCE header
    wb.set_cell_value_tracked(si, 0, 0, "STRIPE BALANCE");
    wb.set_cell_value_tracked(si, 0, 1, "Amount");
    wb.set_cell_value_tracked(si, 0, 2, "Status");

    // Rows 2-6: category breakdowns (divide by 100: amount_minor → dollars)
    wb.set_cell_value_tracked(si, 1, 0, "Charges");
    wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(stripe!E$2:E$1001,\"charge\",stripe!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 2, 0, "Fees");
    wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(stripe!E$2:E$1001,\"fee\",stripe!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 3, 0, "Refunds");
    wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(stripe!E$2:E$1001,\"refund\",stripe!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 4, 0, "Payouts");
    wb.set_cell_value_tracked(si, 4, 1, "=SUMIF(stripe!E$2:E$1001,\"payout\",stripe!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 5, 0, "Adjustments");
    wb.set_cell_value_tracked(si, 5, 1, "=SUMIF(stripe!E$2:E$1001,\"adjustment\",stripe!C$2:C$1001)/100");

    // Row 7: Net (should be 0 for balanced books)
    wb.set_cell_value_tracked(si, 6, 0, "Net");
    wb.set_cell_value_tracked(si, 6, 1, "=SUM(B2:B6)");
    wb.set_cell_value_tracked(si, 6, 2, "=IF(B7=0,\"PASS\",\"FAIL\")");

    // Row 8: blank separator

    // Row 9: 3-WAY PAYOUT MATCHING header
    wb.set_cell_value_tracked(si, 8, 0, "3-WAY PAYOUT MATCHING");
    wb.set_cell_value_tracked(si, 8, 1, "Amount");
    wb.set_cell_value_tracked(si, 8, 2, "Status");

    // Row 10: Stripe Payouts total (divide by 100: cents → dollars)
    wb.set_cell_value_tracked(si, 9, 0, "Stripe Payouts");
    wb.set_cell_value_tracked(si, 9, 1, "=SUM(stripe!J$2:J$1001)/100");

    // Row 11: Matched in Mercury
    wb.set_cell_value_tracked(si, 10, 0, "Matched in Mercury");
    wb.set_cell_value_tracked(si, 10, 1, "=SUMIF(stripe!L$2:L$1001,\"MATCHED\",stripe!J$2:J$1001)/100");

    // Row 12: Matched in QBO
    wb.set_cell_value_tracked(si, 11, 0, "Matched in QBO");
    wb.set_cell_value_tracked(si, 11, 1, "=SUMIF(stripe!N$2:N$1001,\"MATCHED\",stripe!J$2:J$1001)/100");

    // Row 13: Fully Matched (both mercury + qbo)
    wb.set_cell_value_tracked(si, 12, 0, "Fully Matched (3-way)");
    wb.set_cell_value_tracked(si, 12, 1, "=SUMIF(stripe!O$2:O$1001,\"MATCHED\",stripe!J$2:J$1001)/100");

    // Row 14: Partially Matched
    wb.set_cell_value_tracked(si, 13, 0, "Partially Matched");
    wb.set_cell_value_tracked(si, 13, 1, "=SUMIF(stripe!O$2:O$1001,\"PARTIAL\",stripe!J$2:J$1001)/100");

    // Row 15: Unmatched
    wb.set_cell_value_tracked(si, 14, 0, "Unmatched Payouts");
    wb.set_cell_value_tracked(si, 14, 1, "=SUMIF(stripe!O$2:O$1001,\"UNMATCHED\",stripe!J$2:J$1001)/100");

    // Row 16: Variance
    wb.set_cell_value_tracked(si, 15, 0, "Variance");
    wb.set_cell_value_tracked(si, 15, 1, "=B10-B13");
    wb.set_cell_value_tracked(si, 15, 2, "=IF(B16=0,\"PASS\",\"FAIL\")");

    // Row 17: blank separator

    // Row 18: MATCH COUNTS header
    wb.set_cell_value_tracked(si, 17, 0, "MATCH COUNTS");
    wb.set_cell_value_tracked(si, 17, 1, "Count");

    wb.set_cell_value_tracked(si, 18, 0, "Stripe Payouts");
    wb.set_cell_value_tracked(si, 18, 1, "=COUNTIF(stripe!E$2:E$1001,\"payout\")");

    wb.set_cell_value_tracked(si, 19, 0, "Fully Matched");
    wb.set_cell_value_tracked(si, 19, 1, "=COUNTIF(stripe!O$2:O$1001,\"MATCHED\")");

    wb.set_cell_value_tracked(si, 20, 0, "Partially Matched");
    wb.set_cell_value_tracked(si, 20, 1, "=COUNTIF(stripe!O$2:O$1001,\"PARTIAL\")");

    wb.set_cell_value_tracked(si, 21, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 21, 1, "=COUNTIF(stripe!O$2:O$1001,\"UNMATCHED\")");

    // Row 23: blank separator

    // Row 24: QBO DEPOSITS header
    wb.set_cell_value_tracked(si, 23, 0, "QBO DEPOSITS");
    wb.set_cell_value_tracked(si, 23, 1, "Count");
    wb.set_cell_value_tracked(si, 23, 2, "Amount");

    wb.set_cell_value_tracked(si, 24, 0, "Total Deposits");
    wb.set_cell_value_tracked(si, 24, 1, "=COUNTIF(qbo!E$2:E$1001,\"deposit\")");
    wb.set_cell_value_tracked(si, 24, 2, "=SUMIF(qbo!E$2:E$1001,\"deposit\",qbo!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 25, 0, "Matched to Stripe");
    wb.set_cell_value_tracked(si, 25, 1, "=COUNTIF(qbo!L$2:L$1001,\"MATCHED\")");
    wb.set_cell_value_tracked(si, 25, 2, "=SUMIF(qbo!L$2:L$1001,\"MATCHED\",qbo!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 26, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 26, 1, "=COUNTIF(qbo!L$2:L$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 26, 2, "=SUMIF(qbo!L$2:L$1001,\"UNMATCHED\",qbo!C$2:C$1001)/100");

    // Row 28: blank separator

    // Row 29: MERCURY DEPOSITS header
    wb.set_cell_value_tracked(si, 28, 0, "MERCURY DEPOSITS");
    wb.set_cell_value_tracked(si, 28, 1, "Count");
    wb.set_cell_value_tracked(si, 28, 2, "Amount");

    wb.set_cell_value_tracked(si, 29, 0, "Total Deposits");
    wb.set_cell_value_tracked(si, 29, 1, "=COUNTIF(mercury!E$2:E$1001,\"deposit\")");
    wb.set_cell_value_tracked(si, 29, 2, "=SUMIF(mercury!E$2:E$1001,\"deposit\",mercury!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 30, 0, "Matched to Stripe");
    wb.set_cell_value_tracked(si, 30, 1, "=COUNTIF(mercury!L$2:L$1001,\"MATCHED\")");
    wb.set_cell_value_tracked(si, 30, 2, "=SUMIF(mercury!L$2:L$1001,\"MATCHED\",mercury!C$2:C$1001)/100");

    wb.set_cell_value_tracked(si, 31, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 31, 1, "=COUNTIF(mercury!L$2:L$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 31, 2, "=SUMIF(mercury!L$2:L$1001,\"UNMATCHED\",mercury!C$2:C$1001)/100");

    // Row 33: blank separator

    // Row 34: ROLLUP INTEGRITY header (0-indexed row 33)
    wb.set_cell_value_tracked(si, 33, 0, "ROLLUP INTEGRITY");
    wb.set_cell_value_tracked(si, 33, 1, "Count");
    wb.set_cell_value_tracked(si, 33, 2, "Status");

    wb.set_cell_value_tracked(si, 34, 0, "Payouts Checked");
    wb.set_cell_value_tracked(si, 34, 1, "=COUNTIF(stripe!Q$2:Q$1001,\"OK\")+COUNTIF(stripe!Q$2:Q$1001,\"FAIL\")");

    wb.set_cell_value_tracked(si, 35, 0, "Passed");
    wb.set_cell_value_tracked(si, 35, 1, "=COUNTIF(stripe!Q$2:Q$1001,\"OK\")");

    wb.set_cell_value_tracked(si, 36, 0, "Failed");
    wb.set_cell_value_tracked(si, 36, 1, "=COUNTIF(stripe!Q$2:Q$1001,\"FAIL\")");
    wb.set_cell_value_tracked(si, 36, 2, "=IF(B37=0,\"PASS\",\"FAIL\")");

    // Row 38: FEE AUDIT header (0-indexed row 37)
    wb.set_cell_value_tracked(si, 37, 0, "FEE AUDIT");
    wb.set_cell_value_tracked(si, 37, 1, "Amount");
    wb.set_cell_value_tracked(si, 37, 2, "Status");

    wb.set_cell_value_tracked(si, 38, 0, "Contract Rate (%)");  // Row 39 → B39
    wb.set_cell_value_tracked(si, 38, 1, "2.90");

    wb.set_cell_value_tracked(si, 39, 0, "Per-Txn Fee (¢)");   // Row 40 → B40
    wb.set_cell_value_tracked(si, 39, 1, "30");

    wb.set_cell_value_tracked(si, 40, 0, "Payouts Checked");
    wb.set_cell_value_tracked(si, 40, 1, "=COUNTIF(stripe!W$2:W$1001,\"OK\")+COUNTIF(stripe!W$2:W$1001,\"REVIEW\")");

    wb.set_cell_value_tracked(si, 41, 0, "Within Tolerance");
    wb.set_cell_value_tracked(si, 41, 1, "=COUNTIF(stripe!W$2:W$1001,\"OK\")");

    wb.set_cell_value_tracked(si, 42, 0, "Needs Review");
    wb.set_cell_value_tracked(si, 42, 1, "=COUNTIF(stripe!W$2:W$1001,\"REVIEW\")");

    wb.set_cell_value_tracked(si, 43, 0, "Fee Status");
    wb.set_cell_value_tracked(si, 43, 2, "=IF(B43=0,\"OK\",\"REVIEW\")");

    // Row 45: blank separator

    // Row 46: OVERALL VERDICT (0-indexed row 45)
    wb.set_cell_value_tracked(si, 45, 0, "OVERALL VERDICT");
    wb.set_cell_value_tracked(si, 45, 2, "Status");

    wb.set_cell_value_tracked(si, 46, 2, "=IF(AND(C16=\"PASS\",C37=\"PASS\"),\"PASS\",\"FAIL\")");

    // Row 48: Period Start (0-indexed row 47) — filled by workflow for date-window filtering
    wb.set_cell_value_tracked(si, 47, 0, "Period Start");

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
fn build_stripe_qbo_mercury_recon_template() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    build_template(&manifest_dir.join("tests/recon/templates/stripe-qbo-mercury-recon.sheet"));
    build_template(&manifest_dir.join("../../demo/templates/stripe-qbo-mercury-recon.sheet"));
}
