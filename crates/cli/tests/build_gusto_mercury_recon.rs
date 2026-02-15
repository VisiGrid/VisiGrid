// Build the gusto-mercury-recon.sheet template.
// Run with: cargo test -p visigrid-cli --test build_gusto_mercury_recon -- --ignored --nocapture
//
// Creates a 3-sheet workbook:
//   gusto:   headers A1-M1, matching formulas in J-M for rows 2-1001
//   mercury: headers A1-M1, matching formulas in J-M for rows 2-1001
//   summary: aggregate checks (payroll totals, matching, counts)

use std::path::Path;
use visigrid_engine::workbook::Workbook;

const CANONICAL_HEADERS: [&str; 9] = [
    "effective_date", "posted_date", "amount_minor", "currency",
    "type", "source", "source_id", "group_id", "description",
];

fn build_template(out_path: &Path) {
    let mut wb = Workbook::new();

    // ── Sheet 0: gusto ──
    let renamed = wb.rename_sheet(0, "gusto");
    assert!(renamed, "rename_sheet(0, 'gusto') failed");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, col, hdr);
    }
    // Matching column headers (J1-M1)
    wb.set_cell_value_tracked(0, 0, 9, "debit_abs");
    wb.set_cell_value_tracked(0, 0, 10, "match_key");
    wb.set_cell_value_tracked(0, 0, 11, "mercury_match");
    wb.set_cell_value_tracked(0, 0, 12, "match_status");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1; // 1-indexed row number for formula references
        // J: absolute debit amount (negate negative payroll amounts)
        wb.set_cell_value_tracked(0, r, 9,
            &format!("=IF(LEFT(E{row1},8)=\"payroll_\",-C{row1},\"\")"));
        // K: composite match key (date|amount)
        wb.set_cell_value_tracked(0, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",A{row1}&\"|\"&J{row1})"));
        // L: XLOOKUP matching withdrawal by date+amount key
        wb.set_cell_value_tracked(0, r, 11,
            &format!("=IF(K{row1}=\"\",\"\",IFERROR(XLOOKUP(K{row1},mercury!K$2:K$1001,mercury!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // M: status label
        wb.set_cell_value_tracked(0, r, 12,
            &format!("=IF(K{row1}=\"\",\"\",IF(L{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(L{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 1: mercury ──
    let mi = wb.add_sheet_named("mercury").expect("add mercury sheet");

    // Canonical CSV headers (A1-I1)
    for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
        wb.set_cell_value_tracked(mi, 0, col, hdr);
    }
    // Matching column headers (J1-M1)
    wb.set_cell_value_tracked(mi, 0, 9, "withdrawal_abs");
    wb.set_cell_value_tracked(mi, 0, 10, "match_key");
    wb.set_cell_value_tracked(mi, 0, 11, "gusto_match");
    wb.set_cell_value_tracked(mi, 0, 12, "match_status");

    // Formulas for rows 2-1001 (0-indexed rows 1-1000)
    for r in 1..=1000 {
        let row1 = r + 1;
        // J: absolute withdrawal amount
        wb.set_cell_value_tracked(mi, r, 9,
            &format!("=IF(OR(E{row1}=\"withdrawal\",E{row1}=\"expense\"),-C{row1},\"\")"));
        // K: composite match key (date|amount)
        wb.set_cell_value_tracked(mi, r, 10,
            &format!("=IF(J{row1}=\"\",\"\",A{row1}&\"|\"&J{row1})"));
        // L: XLOOKUP matching debit by date+amount key
        wb.set_cell_value_tracked(mi, r, 11,
            &format!("=IF(K{row1}=\"\",\"\",IFERROR(XLOOKUP(K{row1},gusto!K$2:K$1001,gusto!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
        // M: status label
        wb.set_cell_value_tracked(mi, r, 12,
            &format!("=IF(K{row1}=\"\",\"\",IF(L{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(L{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
    }

    // ── Sheet 2: summary ──
    let si = wb.add_sheet_named("summary").expect("add summary sheet");

    // Row 1: GUSTO PAYROLL header
    wb.set_cell_value_tracked(si, 0, 0, "GUSTO PAYROLL");
    wb.set_cell_value_tracked(si, 0, 1, "Amount");
    wb.set_cell_value_tracked(si, 0, 2, "Status");

    // Rows 2-4: category breakdowns
    wb.set_cell_value_tracked(si, 1, 0, "Net Pay");
    wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_net\",gusto!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 2, 0, "Taxes");
    wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_tax\",gusto!C$2:C$1001)");

    wb.set_cell_value_tracked(si, 3, 0, "Fees/Other");
    wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_other\",gusto!C$2:C$1001)");

    // Row 5: Total Debits
    wb.set_cell_value_tracked(si, 4, 0, "Total Debits");
    wb.set_cell_value_tracked(si, 4, 1, "=SUM(B2:B4)");

    // Row 6: blank separator

    // Row 7: PAYROLL MATCHING header
    wb.set_cell_value_tracked(si, 6, 0, "PAYROLL MATCHING");
    wb.set_cell_value_tracked(si, 6, 1, "Amount");
    wb.set_cell_value_tracked(si, 6, 2, "Status");

    // Rows 8-11: payroll matching
    wb.set_cell_value_tracked(si, 7, 0, "Gusto Debits");
    wb.set_cell_value_tracked(si, 7, 1, "=ABS(B5)");

    wb.set_cell_value_tracked(si, 8, 0, "Matched Withdrawals");
    wb.set_cell_value_tracked(si, 8, 1, "=SUMIF(gusto!M$2:M$1001,\"MATCHED\",gusto!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 9, 0, "Unmatched Debits");
    wb.set_cell_value_tracked(si, 9, 1, "=SUMIF(gusto!M$2:M$1001,\"UNMATCHED\",gusto!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 10, 0, "Difference");
    wb.set_cell_value_tracked(si, 10, 1, "=B8-B9");
    wb.set_cell_value_tracked(si, 10, 2, "=IF(B11=0,\"PASS\",\"FAIL\")");

    // Row 12: blank separator

    // Row 13: MATCH COUNTS header
    wb.set_cell_value_tracked(si, 12, 0, "MATCH COUNTS");
    wb.set_cell_value_tracked(si, 12, 1, "Count");

    wb.set_cell_value_tracked(si, 13, 0, "Gusto Rows");
    wb.set_cell_value_tracked(si, 13, 1,
        "=COUNTIF(gusto!M$2:M$1001,\"MATCHED\")+COUNTIF(gusto!M$2:M$1001,\"UNMATCHED\")");

    wb.set_cell_value_tracked(si, 14, 0, "Matched");
    wb.set_cell_value_tracked(si, 14, 1, "=COUNTIF(gusto!M$2:M$1001,\"MATCHED\")");

    wb.set_cell_value_tracked(si, 15, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 15, 1, "=COUNTIF(gusto!M$2:M$1001,\"UNMATCHED\")");

    // Row 17: blank separator

    // Row 18: MERCURY UNMATCHED header
    wb.set_cell_value_tracked(si, 17, 0, "MERCURY UNMATCHED");
    wb.set_cell_value_tracked(si, 17, 1, "Count");
    wb.set_cell_value_tracked(si, 17, 2, "Amount");

    wb.set_cell_value_tracked(si, 18, 0, "Total Withdrawals");
    wb.set_cell_value_tracked(si, 18, 1,
        "=COUNTIF(mercury!M$2:M$1001,\"MATCHED\")+COUNTIF(mercury!M$2:M$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 18, 2,
        "=SUMIF(mercury!M$2:M$1001,\"MATCHED\",mercury!J$2:J$1001)+SUMIF(mercury!M$2:M$1001,\"UNMATCHED\",mercury!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 19, 0, "Matched to Gusto");
    wb.set_cell_value_tracked(si, 19, 1, "=COUNTIF(mercury!M$2:M$1001,\"MATCHED\")");
    wb.set_cell_value_tracked(si, 19, 2, "=SUMIF(mercury!M$2:M$1001,\"MATCHED\",mercury!J$2:J$1001)");

    wb.set_cell_value_tracked(si, 20, 0, "Unmatched");
    wb.set_cell_value_tracked(si, 20, 1, "=COUNTIF(mercury!M$2:M$1001,\"UNMATCHED\")");
    wb.set_cell_value_tracked(si, 20, 2, "=SUMIF(mercury!M$2:M$1001,\"UNMATCHED\",mercury!J$2:J$1001)");

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
fn build_gusto_mercury_recon_template() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    build_template(&manifest_dir.join("tests/recon/templates/gusto-mercury-recon.sheet"));
    build_template(&manifest_dir.join("../../demo/templates/gusto-mercury-recon.sheet"));
}
