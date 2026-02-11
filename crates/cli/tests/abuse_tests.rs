// Abuse tests for `vgrid fill` and cell assertions.
// Run with: cargo test -p visigrid-cli --test abuse_tests -- --nocapture
//
// These tests validate real-world failure scenarios without VisiHub auth.

use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::Value;
use visigrid_io::native;
use std::path::Path;
use std::process::Command;
use std::sync::Once;

static TEMPLATE_INIT: Once = Once::new();

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

fn template_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/abuse/templates/recon-template.sheet")
}

fn csv_path(name: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/abuse/csv/{}", name))
}

fn load_and_read_cell(sheet_path: &Path, sheet_name: &str, row: usize, col: usize) -> Value {
    let mut wb = native::load_workbook(sheet_path).expect("load workbook");
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet_id = wb.sheet_id_by_name(sheet_name).expect("find sheet");
    let idx = wb.idx_for_sheet_id(sheet_id).expect("sheet idx");
    wb.sheet(idx).unwrap().get_computed_value(row, col)
}

fn ensure_template() {
    TEMPLATE_INIT.call_once(|| {
        let tmpl = template_path();
        // Build it
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "effective_date");
        wb.set_cell_value_tracked(0, 0, 1, "posted_date");
        wb.set_cell_value_tracked(0, 0, 2, "amount_minor");
        wb.set_cell_value_tracked(0, 0, 3, "currency");
        wb.set_cell_value_tracked(0, 0, 4, "type");
        wb.set_cell_value_tracked(0, 0, 5, "source");
        wb.set_cell_value_tracked(0, 0, 6, "source_id");
        wb.set_cell_value_tracked(0, 0, 7, "group_id");
        wb.set_cell_value_tracked(0, 0, 8, "description");
        wb.set_cell_value_tracked(0, 0, 9, "amount");

        let si = wb.add_sheet_named("summary").unwrap();
        wb.set_cell_value_tracked(si, 0, 0, "Category");
        wb.set_cell_value_tracked(si, 0, 1, "Total (minor units)");
        wb.set_cell_value_tracked(si, 1, 0, "Charges");
        wb.set_cell_value_tracked(si, 2, 0, "Payouts");
        wb.set_cell_value_tracked(si, 3, 0, "Fees");
        wb.set_cell_value_tracked(si, 4, 0, "Refunds");
        wb.set_cell_value_tracked(si, 5, 0, "Adjustments");
        wb.set_cell_value_tracked(si, 6, 0, "Variance");
        wb.set_cell_value_tracked(si, 1, 1, "=SUMIF(Sheet1!E2:E1000,\"charge\",Sheet1!C2:C1000)");
        wb.set_cell_value_tracked(si, 2, 1, "=SUMIF(Sheet1!E2:E1000,\"payout\",Sheet1!C2:C1000)");
        wb.set_cell_value_tracked(si, 3, 1, "=SUMIF(Sheet1!E2:E1000,\"fee\",Sheet1!C2:C1000)");
        wb.set_cell_value_tracked(si, 4, 1, "=SUMIF(Sheet1!E2:E1000,\"refund\",Sheet1!C2:C1000)");
        wb.set_cell_value_tracked(si, 5, 1, "=SUMIF(Sheet1!E2:E1000,\"adjustment\",Sheet1!C2:C1000)");
        wb.set_cell_value_tracked(si, 6, 1, "=B2+B3+B4+B5+B6");
        wb.set_cell_value_tracked(si, 7, 0, "Error test");
        wb.set_cell_value_tracked(si, 7, 1, "=1/0");
        wb.set_cell_value_tracked(si, 8, 0, "String test");
        wb.set_cell_value_tracked(si, 8, 1, "not a number");
        wb.set_cell_value_tracked(si, 9, 0, "Blank test");

        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();

        std::fs::create_dir_all(tmpl.parent().unwrap()).ok();
        native::save_workbook(&wb, &tmpl).expect("save template");
    });
}

// ── Test #6: Currency symbol rejection ──

#[test]
fn test_06_currency_symbol_rejection() {
    ensure_template();
    let out = tempfile::NamedTempFile::new().unwrap();
    let output = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("currency-symbol.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out.path().to_str().unwrap()])
        .output().unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success(), "should fail");
    assert!(stderr.contains("currency symbol"), "error should mention currency symbol: {}", stderr);
    eprintln!("TEST #6 PASS: {}", stderr.trim());
}

// ── Test #7: Comma rejection ──

#[test]
fn test_07_comma_rejection() {
    ensure_template();
    let out = tempfile::NamedTempFile::new().unwrap();
    let output = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("comma-number.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out.path().to_str().unwrap()])
        .output().unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success());
    assert!(stderr.contains("comma in numeric field"), "error: {}", stderr);
    eprintln!("TEST #7 PASS: {}", stderr.trim());
}

// ── Test #8: Wrong decimals ──

#[test]
fn test_08_wrong_decimals() {
    ensure_template();
    let out = tempfile::NamedTempFile::new().unwrap();
    let output = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("wrong-decimals.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out.path().to_str().unwrap()])
        .output().unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success());
    assert!(stderr.contains("wrong decimal places"), "error: {}", stderr);
    eprintln!("TEST #8 PASS: {}", stderr.trim());
}

// ── Test #9: Formula injection ──

#[test]
fn test_09_formula_injection() {
    ensure_template();
    let out = tempfile::NamedTempFile::new().unwrap();
    let output = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("formula-injection.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out.path().to_str().unwrap()])
        .output().unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success());
    assert!(stderr.contains("formula injection"), "error: {}", stderr);
    eprintln!("TEST #9 PASS: {}", stderr.trim());
}

// ── Test #10: Missing column detection ──
// (Note: vgrid fill doesn't validate column names — it just fills cells.
// Missing amount_minor means the summary formulas compute wrong values.
// The right check is that the SUMIF formulas produce 0.)

#[test]
fn test_10_missing_column_still_fills() {
    ensure_template();
    let out = tempfile::NamedTempFile::new().unwrap();
    let output = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("missing-column.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out.path().to_str().unwrap()])
        .output().unwrap();
    // Fill succeeds — CSV is valid, just has fewer columns
    assert!(output.status.success(), "fill should succeed: {}", String::from_utf8_lossy(&output.stderr));

    // But the summary formulas should show 0 (no amount_minor column)
    let val = load_and_read_cell(out.path(), "summary", 1, 1); // B2 = charges
    match val {
        Value::Number(n) => assert_eq!(n, 0.0, "charges should be 0 with missing column"),
        _ => panic!("expected Number, got {:?}", val),
    }
    eprintln!("TEST #10 PASS: fill succeeds, formulas return 0 for missing column data");
}

// ── Test #11a: Stale rows WITHOUT --clear ──

#[test]
fn test_11a_stale_rows_without_clear() {
    ensure_template();
    let step1 = tempfile::NamedTempFile::new().unwrap();
    let step2 = tempfile::NamedTempFile::new().unwrap();

    // Fill with balanced.csv (4 rows)
    let out1 = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", step1.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out1.status.success(), "step1: {}", String::from_utf8_lossy(&out1.stderr));

    // Verify variance is 0 (balanced data)
    let var1 = load_and_read_cell(step1.path(), "summary", 6, 1);
    eprintln!("After balanced fill, variance = {:?}", var1);
    assert!(matches!(var1, Value::Number(n) if n == 0.0), "balanced should net to 0");

    // Fill again with short.csv (1 row) WITHOUT --clear
    let out2 = vgrid()
        .args(["fill", step1.path().to_str().unwrap(),
            "--csv", csv_path("short.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers",
            "--out", step2.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out2.status.success(), "step2: {}", String::from_utf8_lossy(&out2.stderr));

    // Stale rows: variance should NOT be 0
    let var2 = load_and_read_cell(step2.path(), "summary", 6, 1);
    eprintln!("After stale fill (no --clear), variance = {:?}", var2);
    assert!(!matches!(var2, Value::Number(n) if n == 0.0),
        "stale fill should corrupt totals, got {:?}", var2);

    eprintln!("TEST #11a PASS: stale rows detected (variance != 0 without --clear)");
}

// ── Test #11b: Clean fill WITH --clear ──

#[test]
fn test_11b_clean_fill_with_clear() {
    ensure_template();
    let step1 = tempfile::NamedTempFile::new().unwrap();
    let step2 = tempfile::NamedTempFile::new().unwrap();

    // Fill with balanced.csv
    let out1 = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", step1.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out1.status.success());

    // Fill again with short.csv WITH --clear
    let out2 = vgrid()
        .args(["fill", step1.path().to_str().unwrap(),
            "--csv", csv_path("short.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", step2.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out2.status.success());

    // With --clear: only short.csv data, variance = 200000 (just charges, no offsets)
    let var = load_and_read_cell(step2.path(), "summary", 6, 1);
    eprintln!("After clean fill (--clear), variance = {:?}", var);
    // Should be 200000 (just the one charge)
    match var {
        Value::Number(n) => {
            assert_eq!(n, 200000.0, "clean fill should only have charge data");
        }
        _ => panic!("expected Number, got {:?}", var),
    }

    eprintln!("TEST #11b PASS: --clear removes stale rows, totals correct");
}

// ── Test #12: Blank assertion cell ──

#[test]
fn test_12_blank_assertion_cell() {
    ensure_template();
    let filled = tempfile::NamedTempFile::new().unwrap();

    let out = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", filled.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out.status.success());

    // B10 is intentionally blank in the template
    let val = load_and_read_cell(filled.path(), "summary", 9, 1);
    eprintln!("summary!B10 = {:?}", val);
    assert!(matches!(val, Value::Empty), "B10 should be empty");
    eprintln!("TEST #12 PASS: blank cell correctly identified as Empty");
}

// ── Test #13: Error assertion cell ──

#[test]
fn test_13_error_assertion_cell() {
    ensure_template();
    let filled = tempfile::NamedTempFile::new().unwrap();

    let out = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", filled.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out.status.success());

    // B8 = =1/0 should be an error
    let val = load_and_read_cell(filled.path(), "summary", 7, 1);
    eprintln!("summary!B8 (=1/0) = {:?}", val);
    assert!(matches!(val, Value::Error(_)), "B8 should be Error, got {:?}", val);
    eprintln!("TEST #13 PASS: error cell correctly identified");
}

// ── Test #14: String cell assertion ──

#[test]
fn test_14_string_assertion_cell() {
    ensure_template();
    let filled = tempfile::NamedTempFile::new().unwrap();

    let out = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", filled.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out.status.success());

    // B9 = "not a number" (text)
    let val = load_and_read_cell(filled.path(), "summary", 8, 1);
    eprintln!("summary!B9 = {:?}", val);
    assert!(matches!(val, Value::Text(_)), "B9 should be Text, got {:?}", val);
    eprintln!("TEST #14 PASS: string cell correctly identified");
}

// ── Test: Balanced fill produces variance 0 ──

#[test]
fn test_balanced_fill_variance_zero() {
    ensure_template();
    let out_path = tempfile::NamedTempFile::new().unwrap();

    let out = vgrid()
        .args(["fill", template_path().to_str().unwrap(),
            "--csv", csv_path("balanced.csv").to_str().unwrap(),
            "--target", "Sheet1!A2", "--headers", "--clear",
            "--out", out_path.path().to_str().unwrap()])
        .output().unwrap();
    assert!(out.status.success(), "fill: {}", String::from_utf8_lossy(&out.stderr));

    let var = load_and_read_cell(out_path.path(), "summary", 6, 1);
    eprintln!("Balanced fill variance (B7) = {:?}", var);
    assert!(matches!(var, Value::Number(n) if n == 0.0),
        "balanced data should produce variance 0, got {:?}", var);

    // Also verify charges (B2)
    let charges = load_and_read_cell(out_path.path(), "summary", 1, 1);
    eprintln!("Charges (B2) = {:?}", charges);
    assert!(matches!(charges, Value::Number(n) if n == 150000.0),
        "charges should be 150000 (100000 + 50000), got {:?}", charges);
}
