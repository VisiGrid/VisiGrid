// Integration tests for the Gusto ↔ Mercury reconciliation template.
// Run with: cargo test -p visigrid-cli --test gusto_mercury_recon_tests -- --nocapture
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
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/recon/templates/gusto-mercury-recon.sheet")
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

        // ── Sheet 0: gusto ──
        wb.rename_sheet(0, "gusto");
        for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
            wb.set_cell_value_tracked(0, 0, col, hdr);
        }
        wb.set_cell_value_tracked(0, 0, 9, "debit_abs");
        wb.set_cell_value_tracked(0, 0, 10, "match_key");
        wb.set_cell_value_tracked(0, 0, 11, "mercury_match");
        wb.set_cell_value_tracked(0, 0, 12, "match_status");

        for r in 1..=1000 {
            let row1 = r + 1;
            wb.set_cell_value_tracked(0, r, 9,
                &format!("=IF(LEFT(E{row1},8)=\"payroll_\",-C{row1},\"\")"));
            wb.set_cell_value_tracked(0, r, 10,
                &format!("=IF(J{row1}=\"\",\"\",A{row1}&\"|\"&J{row1})"));
            wb.set_cell_value_tracked(0, r, 11,
                &format!("=IF(K{row1}=\"\",\"\",IFERROR(XLOOKUP(K{row1},mercury!K$2:K$1001,mercury!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
            wb.set_cell_value_tracked(0, r, 12,
                &format!("=IF(K{row1}=\"\",\"\",IF(L{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(L{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
        }

        // ── Sheet 1: mercury ──
        let mi = wb.add_sheet_named("mercury").unwrap();
        for (col, hdr) in CANONICAL_HEADERS.iter().enumerate() {
            wb.set_cell_value_tracked(mi, 0, col, hdr);
        }
        wb.set_cell_value_tracked(mi, 0, 9, "withdrawal_abs");
        wb.set_cell_value_tracked(mi, 0, 10, "match_key");
        wb.set_cell_value_tracked(mi, 0, 11, "gusto_match");
        wb.set_cell_value_tracked(mi, 0, 12, "match_status");

        for r in 1..=1000 {
            let row1 = r + 1;
            wb.set_cell_value_tracked(mi, r, 9,
                &format!("=IF(OR(E{row1}=\"withdrawal\",E{row1}=\"expense\"),-C{row1},\"\")"));
            wb.set_cell_value_tracked(mi, r, 10,
                &format!("=IF(J{row1}=\"\",\"\",A{row1}&\"|\"&J{row1})"));
            wb.set_cell_value_tracked(mi, r, 11,
                &format!("=IF(K{row1}=\"\",\"\",IFERROR(XLOOKUP(K{row1},gusto!K$2:K$1001,gusto!G$2:G$1001,\"UNMATCHED\"),\"ERROR\"))"));
            wb.set_cell_value_tracked(mi, r, 12,
                &format!("=IF(K{row1}=\"\",\"\",IF(L{row1}=\"UNMATCHED\",\"UNMATCHED\",IF(L{row1}=\"ERROR\",\"ERROR\",\"MATCHED\")))"));
        }

        // ── Sheet 2: summary ──
        let si = wb.add_sheet_named("summary").unwrap();

        wb.set_cell_value_tracked(si, 0, 0, "GUSTO PAYROLL");
        wb.set_cell_value_tracked(si, 0, 1, "Amount");
        wb.set_cell_value_tracked(si, 0, 2, "Status");

        wb.set_cell_value_tracked(si, 1, 0, "Net Pay");
        wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_net\",gusto!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 2, 0, "Taxes");
        wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_tax\",gusto!C$2:C$1001)");
        wb.set_cell_value_tracked(si, 3, 0, "Fees/Other");
        wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(gusto!E$2:E$1001,\"payroll_other\",gusto!C$2:C$1001)");

        wb.set_cell_value_tracked(si, 4, 0, "Total Debits");
        wb.set_cell_value_tracked(si, 4, 1, "=SUM(B2:B4)");

        wb.set_cell_value_tracked(si, 6, 0, "PAYROLL MATCHING");
        wb.set_cell_value_tracked(si, 6, 1, "Amount");
        wb.set_cell_value_tracked(si, 6, 2, "Status");

        wb.set_cell_value_tracked(si, 7, 0, "Gusto Debits");
        wb.set_cell_value_tracked(si, 7, 1, "=ABS(B5)");
        wb.set_cell_value_tracked(si, 8, 0, "Matched Withdrawals");
        wb.set_cell_value_tracked(si, 8, 1, "=SUMIF(gusto!M$2:M$1001,\"MATCHED\",gusto!J$2:J$1001)");
        wb.set_cell_value_tracked(si, 9, 0, "Unmatched Debits");
        wb.set_cell_value_tracked(si, 9, 1, "=SUMIF(gusto!M$2:M$1001,\"UNMATCHED\",gusto!J$2:J$1001)");
        wb.set_cell_value_tracked(si, 10, 0, "Difference");
        wb.set_cell_value_tracked(si, 10, 1, "=B8-B9");
        wb.set_cell_value_tracked(si, 10, 2, "=IF(B11=0,\"PASS\",\"FAIL\")");

        wb.set_cell_value_tracked(si, 12, 0, "MATCH COUNTS");
        wb.set_cell_value_tracked(si, 12, 1, "Count");
        wb.set_cell_value_tracked(si, 13, 0, "Gusto Rows");
        wb.set_cell_value_tracked(si, 13, 1,
            "=COUNTIF(gusto!M$2:M$1001,\"MATCHED\")+COUNTIF(gusto!M$2:M$1001,\"UNMATCHED\")");
        wb.set_cell_value_tracked(si, 14, 0, "Matched");
        wb.set_cell_value_tracked(si, 14, 1, "=COUNTIF(gusto!M$2:M$1001,\"MATCHED\")");
        wb.set_cell_value_tracked(si, 15, 0, "Unmatched");
        wb.set_cell_value_tracked(si, 15, 1, "=COUNTIF(gusto!M$2:M$1001,\"UNMATCHED\")");

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

/// Fill the template with gusto + mercury CSVs (two-step fill).
fn fill_template(gusto_csv: &str, mercury_csv: &str) -> tempfile::NamedTempFile {
    ensure_template();

    let step1 = tempfile::NamedTempFile::new().unwrap();
    let final_out = tempfile::NamedTempFile::new().unwrap();

    // Step 1: fill gusto data
    let out1 = vgrid()
        .args([
            "fill", template_path().to_str().unwrap(),
            "--csv", csv_path(gusto_csv).to_str().unwrap(),
            "--target", "gusto!A2", "--headers",
            "--out", step1.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(out1.status.success(), "gusto fill failed: {}", String::from_utf8_lossy(&out1.stderr));

    // Step 2: fill mercury data
    let out2 = vgrid()
        .args([
            "fill", step1.path().to_str().unwrap(),
            "--csv", csv_path(mercury_csv).to_str().unwrap(),
            "--target", "mercury!A2", "--headers",
            "--out", final_out.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(out2.status.success(), "mercury fill failed: {}", String::from_utf8_lossy(&out2.stderr));

    final_out
}

/// Fill the template with only gusto data (no mercury fill step).
fn fill_gusto_only(gusto_csv: &str) -> tempfile::NamedTempFile {
    ensure_template();

    let out = tempfile::NamedTempFile::new().unwrap();

    let result = vgrid()
        .args([
            "fill", template_path().to_str().unwrap(),
            "--csv", csv_path(gusto_csv).to_str().unwrap(),
            "--target", "gusto!A2", "--headers",
            "--out", out.path().to_str().unwrap(),
        ])
        .output()
        .unwrap();
    assert!(result.status.success(), "gusto fill failed: {}", String::from_utf8_lossy(&result.stderr));

    out
}

// ── Test 1: All payrolls matched, Difference=0, PASS ──

#[test]
fn test_all_payrolls_matched() {
    let out = fill_template("gusto-two-payrolls.csv", "mercury-payroll-matched.csv");
    let p = out.path();

    // Total Debits (sum of all payroll amounts)
    // Net: -152000 + -155000 = -307000
    // Tax: -28400 + -29000 = -57400
    // Other: -1500 + -1500 = -3000
    // Total: -367400
    let total = load_and_read_cell(p, "summary", 4, 1);
    eprintln!("summary!B5 (Total Debits) = {:?}", total);
    assert!(matches!(total, Value::Number(n) if n == -367400.0), "got {:?}", total);

    // Difference = 0
    let diff = load_and_read_cell(p, "summary", 10, 1);
    eprintln!("summary!B11 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 0.0), "Difference should be 0, got {:?}", diff);

    // Difference status = PASS
    let diff_status = load_and_read_cell(p, "summary", 10, 2);
    eprintln!("summary!C11 (Status) = {:?}", diff_status);
    assert!(matches!(diff_status, Value::Text(ref s) if s == "PASS"), "got {:?}", diff_status);

    // Gusto Rows = 6
    let gusto_rows = load_and_read_cell(p, "summary", 13, 1);
    eprintln!("summary!B14 (Gusto Rows) = {:?}", gusto_rows);
    assert!(matches!(gusto_rows, Value::Number(n) if n == 6.0), "got {:?}", gusto_rows);

    // Matched = 6
    let matched = load_and_read_cell(p, "summary", 14, 1);
    eprintln!("summary!B15 (Matched) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 6.0), "got {:?}", matched);

    // Unmatched = 0
    let unmatched = load_and_read_cell(p, "summary", 15, 1);
    eprintln!("summary!B16 (Unmatched) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 0.0), "got {:?}", unmatched);

    // Mercury: 1 unmatched (OFFICE SUPPLIES expense)
    let merc_unmatched = load_and_read_cell(p, "summary", 20, 1);
    eprintln!("summary!B21 (Mercury Unmatched) = {:?}", merc_unmatched);
    assert!(matches!(merc_unmatched, Value::Number(n) if n == 1.0), "got {:?}", merc_unmatched);

    eprintln!("TEST PASS: all_payrolls_matched");
}

// ── Test 2: Missing withdrawal detected ──

#[test]
fn test_missing_withdrawal_detected() {
    let out = fill_template("gusto-two-payrolls.csv", "mercury-payroll-missing.csv");
    let p = out.path();

    // Difference = 155000 (the unmatched net pay withdrawal)
    let diff = load_and_read_cell(p, "summary", 10, 1);
    eprintln!("summary!B11 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 155000.0),
        "Difference should be 155000, got {:?}", diff);

    // Difference status = FAIL
    let diff_status = load_and_read_cell(p, "summary", 10, 2);
    eprintln!("summary!C11 (Status) = {:?}", diff_status);
    assert!(matches!(diff_status, Value::Text(ref s) if s == "FAIL"), "got {:?}", diff_status);

    // Matched = 5 (all except the missing net pay)
    let matched = load_and_read_cell(p, "summary", 14, 1);
    eprintln!("summary!B15 (Matched) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 5.0), "got {:?}", matched);

    // Unmatched = 1
    let unmatched = load_and_read_cell(p, "summary", 15, 1);
    eprintln!("summary!B16 (Unmatched) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 1.0), "got {:?}", unmatched);

    eprintln!("TEST PASS: missing_withdrawal_detected");
}

// ── Test 3: Empty data produces zero counts ──

#[test]
fn test_empty_data_zero_counts() {
    ensure_template();

    // Use the template directly without filling — all data rows are empty
    let p = template_path();

    // Total Debits = 0
    let total = load_and_read_cell(&p, "summary", 4, 1);
    eprintln!("summary!B5 (Total Debits) = {:?}", total);
    assert!(matches!(total, Value::Number(n) if n == 0.0), "got {:?}", total);

    // Gusto Rows = 0
    let count = load_and_read_cell(&p, "summary", 13, 1);
    eprintln!("summary!B14 (Gusto Rows) = {:?}", count);
    assert!(matches!(count, Value::Number(n) if n == 0.0), "got {:?}", count);

    // Matched = 0
    let matched = load_and_read_cell(&p, "summary", 14, 1);
    eprintln!("summary!B15 (Matched) = {:?}", matched);
    assert!(matches!(matched, Value::Number(n) if n == 0.0), "got {:?}", matched);

    eprintln!("TEST PASS: empty_data_zero_counts");
}

// ── Test 4: Gusto only, all payrolls unmatched ──

#[test]
fn test_gusto_only_all_unmatched() {
    let out = fill_gusto_only("gusto-two-payrolls.csv");
    let p = out.path();

    // Unmatched = 6 (all 6 gusto rows have no Mercury match)
    let unmatched = load_and_read_cell(p, "summary", 15, 1);
    eprintln!("summary!B16 (Unmatched) = {:?}", unmatched);
    assert!(matches!(unmatched, Value::Number(n) if n == 6.0), "got {:?}", unmatched);

    // Difference = total debit amounts = 367400
    let diff = load_and_read_cell(p, "summary", 10, 1);
    eprintln!("summary!B11 (Difference) = {:?}", diff);
    assert!(matches!(diff, Value::Number(n) if n == 367400.0),
        "Difference should be 367400, got {:?}", diff);

    // Status = FAIL
    let diff_status = load_and_read_cell(p, "summary", 10, 2);
    eprintln!("summary!C11 (Status) = {:?}", diff_status);
    assert!(matches!(diff_status, Value::Text(ref s) if s == "FAIL"), "got {:?}", diff_status);

    eprintln!("TEST PASS: gusto_only_all_unmatched");
}

// ── Test 5: Duplicate amounts on different dates don't false-match ──

#[test]
fn test_duplicate_amounts_different_dates() {
    // Both payrolls have identical fee amounts (1500) but on different dates
    // (2026-01-15 and 2026-01-31). The date+amount composite key should
    // prevent false matches.
    let out = fill_template("gusto-two-payrolls.csv", "mercury-payroll-matched.csv");
    let p = out.path();

    // Both fee rows (1500 on Jan 15 and 1500 on Jan 31) should be MATCHED
    // because mercury has matching withdrawals on the same dates

    // Row 3 in gusto sheet (0-indexed row 2) = first payroll_other (Jan 15, 1500)
    let status_1 = load_and_read_cell(p, "gusto", 2, 12); // col M (0-indexed 12)
    eprintln!("gusto!M3 (first other match_status) = {:?}", status_1);
    assert!(matches!(status_1, Value::Text(ref s) if s == "MATCHED"),
        "first other row should be MATCHED, got {:?}", status_1);

    // Row 6 in gusto sheet (0-indexed row 5) = second payroll_other (Jan 31, 1500)
    let status_2 = load_and_read_cell(p, "gusto", 5, 12);
    eprintln!("gusto!M6 (second other match_status) = {:?}", status_2);
    assert!(matches!(status_2, Value::Text(ref s) if s == "MATCHED"),
        "second other row should be MATCHED, got {:?}", status_2);

    // Verify the match keys are different (date disambiguates)
    let key_1 = load_and_read_cell(p, "gusto", 2, 10); // col K (0-indexed 10)
    let key_2 = load_and_read_cell(p, "gusto", 5, 10);
    eprintln!("gusto!K3 = {:?}, gusto!K6 = {:?}", key_1, key_2);
    // key_1 should be "2026-01-15|1500", key_2 should be "2026-01-31|1500"
    assert!(
        key_1 != key_2,
        "match keys should differ: {:?} vs {:?}",
        key_1, key_2,
    );

    eprintln!("TEST PASS: duplicate_amounts_different_dates");
}
