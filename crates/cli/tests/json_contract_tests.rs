// Integration tests enforcing the --json stdout contract.
//
// These tests guarantee that stdout from --json commands is:
//   1. Valid JSON
//   2. Exactly one JSON value (no extra lines, no banners, no colors)
//   3. The correct shape for its command type
//
// Run with: cargo test -p visigrid-cli --test json_contract_tests -- --nocapture

use std::process::Command;

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

/// Assert stdout is a single, parseable JSON value with no extra lines.
fn assert_single_json(stdout: &str) -> serde_json::Value {
    let trimmed = stdout.trim();
    assert!(!trimmed.is_empty(), "stdout should not be empty");

    // Must parse as exactly one JSON value
    let val: serde_json::Value = serde_json::from_str(trimmed)
        .unwrap_or_else(|e| panic!(
            "stdout must be valid JSON.\nParse error: {}\nstdout:\n{}",
            e, trimmed
        ));

    // Verify there's no trailing garbage after the JSON value.
    // serde_json::from_str already rejects trailing content, but let's be explicit:
    // the trimmed output should round-trip cleanly.
    let re_serialized = serde_json::to_string(&val).unwrap();
    let re_parsed: serde_json::Value = serde_json::from_str(&re_serialized).unwrap();
    assert_eq!(val, re_parsed, "JSON round-trip should be stable");

    val
}

// ===========================================================================
// vgrid peek --json
// ===========================================================================

#[test]
fn peek_json_csv_produces_valid_json() {
    let csv = std::env::temp_dir().join("vgrid_peek_json_test.csv");
    std::fs::write(&csv, "id,name,amount\n1,Alice,100.50\n2,Bob,200\n3,Charlie,\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--headers", "--json"])
        .output()
        .expect("vgrid peek --json");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);

    // Shape: {"columns":[...], "rows":[[...],...]}
    let obj = val.as_object().expect("should be JSON object");
    assert!(obj.contains_key("columns"), "must have 'columns' key");
    assert!(obj.contains_key("rows"), "must have 'rows' key");

    let columns = obj["columns"].as_array().expect("columns must be array");
    assert_eq!(columns.len(), 3, "should have 3 columns");
    assert_eq!(columns[0].as_str().unwrap(), "id");
    assert_eq!(columns[1].as_str().unwrap(), "name");
    assert_eq!(columns[2].as_str().unwrap(), "amount");

    let rows = obj["rows"].as_array().expect("rows must be array");
    assert_eq!(rows.len(), 3, "should have 3 data rows");

    // Values should be JSON scalars (numbers, not strings for numeric cells)
    let first_row = rows[0].as_array().expect("row must be array");
    assert_eq!(first_row[0], serde_json::json!(1), "id should be numeric");
    assert_eq!(first_row[1], serde_json::json!("Alice"), "name should be string");
    assert_eq!(first_row[2], serde_json::json!(100.5), "amount should be numeric");

    // Empty cell
    let third_row = rows[2].as_array().expect("row must be array");
    assert_eq!(third_row[2], serde_json::json!(""), "empty cell should be empty string");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_json_no_headers_uses_generated_columns() {
    let csv = std::env::temp_dir().join("vgrid_peek_json_noheaders.csv");
    std::fs::write(&csv, "1,Alice\n2,Bob\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--json"])
        .output()
        .expect("vgrid peek --json (no headers)");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);
    let obj = val.as_object().unwrap();
    let columns = obj["columns"].as_array().unwrap();
    // Without --headers, columns should be generated (A, B, ...)
    assert_eq!(columns[0].as_str().unwrap(), "A");
    assert_eq!(columns[1].as_str().unwrap(), "B");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_json_sheet_file_produces_valid_json() {
    // Create a minimal .sheet workbook
    use visigrid_engine::workbook::Workbook;

    let mut wb = Workbook::new();
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set_value(0, 0, "Name");
    sheet.set_value(0, 1, "Score");
    sheet.set_value(1, 0, "Alice");
    sheet.set_value(1, 1, "95");
    sheet.set_value(2, 0, "Bob");
    sheet.set_value(2, 1, "87");

    let dir = tempfile::tempdir().expect("create temp dir");
    let path = dir.path().join("test.sheet");
    visigrid_io::native::save_workbook(&wb, &path).expect("save workbook");

    let output = vgrid()
        .args(["peek", path.to_str().unwrap(), "--json"])
        .output()
        .expect("vgrid peek --json .sheet");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);
    let obj = val.as_object().unwrap();
    assert!(obj.contains_key("columns"));
    assert!(obj.contains_key("rows"));
}

#[test]
fn peek_json_xlsx_produces_valid_json() {
    let xlsx = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/inspect_small.xlsx");

    let output = vgrid()
        .args(["peek", xlsx.to_str().unwrap(), "--json"])
        .output()
        .expect("vgrid peek --json xlsx");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);
    let obj = val.as_object().unwrap();
    assert!(obj.contains_key("columns"));
    assert!(obj.contains_key("rows"));
}

#[test]
fn peek_json_conflicts_with_tui() {
    let csv = std::env::temp_dir().join("vgrid_peek_json_tui.csv");
    std::fs::write(&csv, "a\n1\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--json", "--tui"])
        .output()
        .expect("vgrid peek --json --tui");

    assert!(!output.status.success(), "--json and --tui should conflict");

    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// vgrid calc --json
// ===========================================================================

#[test]
fn calc_json_scalar_produces_valid_json() {
    let csv = std::env::temp_dir().join("vgrid_calc_json_scalar.csv");
    std::fs::write(&csv, "10\n20\n30\n").unwrap();

    let output = vgrid()
        .args(["calc", "=SUM(A1:A3)", "-f", "csv", "--json"])
        .stdin(std::process::Stdio::from(std::fs::File::open(&csv).unwrap()))
        .output()
        .expect("vgrid calc --json");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);

    // Should be a numeric scalar
    assert!(val.is_number(), "SUM should produce a number, got: {}", val);
    assert_eq!(val.as_f64().unwrap(), 60.0);

    std::fs::remove_file(&csv).ok();
}

#[test]
fn calc_json_array_produces_valid_json() {
    let csv = std::env::temp_dir().join("vgrid_calc_json_array.csv");
    std::fs::write(&csv, "1,2\n3,4\n").unwrap();

    // MMULT should produce a 2D array
    let output = vgrid()
        .args(["calc", "=TRANSPOSE(A1:B2)", "-f", "csv", "--json"])
        .stdin(std::process::Stdio::from(std::fs::File::open(&csv).unwrap()))
        .output()
        .expect("vgrid calc --json (array)");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);

    // Should be a 2D array
    assert!(val.is_array(), "TRANSPOSE should produce an array, got: {}", val);
    let rows = val.as_array().unwrap();
    assert_eq!(rows.len(), 2, "TRANSPOSE of 2x2 should have 2 rows");
    assert!(rows[0].is_array(), "each row should be an array");
}

#[test]
fn calc_json_boolean_produces_valid_json() {
    let csv = std::env::temp_dir().join("vgrid_calc_json_bool.csv");
    std::fs::write(&csv, "5\n").unwrap();

    let output = vgrid()
        .args(["calc", "=A1>3", "-f", "csv", "--json"])
        .stdin(std::process::Stdio::from(std::fs::File::open(&csv).unwrap()))
        .output()
        .expect("vgrid calc --json (bool)");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);
    assert!(val.is_boolean(), "comparison should produce boolean, got: {}", val);
    assert_eq!(val.as_bool().unwrap(), true);

    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// vgrid diff --json
// ===========================================================================

#[test]
fn diff_json_produces_valid_json() {
    let left = std::env::temp_dir().join("vgrid_diff_json_left.csv");
    let right = std::env::temp_dir().join("vgrid_diff_json_right.csv");
    std::fs::write(&left, "id,amount\n1,100\n2,200\n3,300\n").unwrap();
    std::fs::write(&right, "id,amount\n1,100\n2,250\n4,400\n").unwrap();

    let output = vgrid()
        .args([
            "diff", left.to_str().unwrap(), right.to_str().unwrap(),
            "--key", "id", "--json",
        ])
        .output()
        .expect("vgrid diff --json");

    // diff may exit 1 (has diffs) — that's ok, check stdout is still JSON
    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);

    // Shape: must have contract_version
    let obj = val.as_object().expect("diff output should be JSON object");
    assert!(obj.contains_key("contract_version"), "must have contract_version");
    assert!(obj.contains_key("summary"), "must have summary");
    assert!(obj.contains_key("results"), "must have results");

    let results = obj["results"].as_array().expect("results must be array");
    // Should detect: row 2 changed, row 3 missing right, row 4 missing left
    assert!(results.len() >= 2, "should detect diffs, got {} results", results.len());

    std::fs::remove_file(&left).ok();
    std::fs::remove_file(&right).ok();
}

#[test]
fn diff_json_no_diffs_produces_valid_json() {
    let left = std::env::temp_dir().join("vgrid_diff_json_nodiff_l.csv");
    let right = std::env::temp_dir().join("vgrid_diff_json_nodiff_r.csv");
    std::fs::write(&left, "id,val\n1,a\n2,b\n").unwrap();
    std::fs::write(&right, "id,val\n1,a\n2,b\n").unwrap();

    let output = vgrid()
        .args([
            "diff", left.to_str().unwrap(), right.to_str().unwrap(),
            "--key", "id", "--json",
        ])
        .output()
        .expect("vgrid diff --json (no diffs)");

    assert!(output.status.success(), "identical files should exit 0\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let val = assert_single_json(&stdout);
    let obj = val.as_object().unwrap();
    assert!(obj.contains_key("contract_version"));

    std::fs::remove_file(&left).ok();
    std::fs::remove_file(&right).ok();
}

#[test]
fn diff_json_suppresses_stderr_banners() {
    let left = std::env::temp_dir().join("vgrid_diff_json_quiet_l.csv");
    let right = std::env::temp_dir().join("vgrid_diff_json_quiet_r.csv");
    std::fs::write(&left, "id,val\n1,a\n2,b\n").unwrap();
    std::fs::write(&right, "id,val\n1,a\n2,c\n").unwrap();

    let output = vgrid()
        .args([
            "diff", left.to_str().unwrap(), right.to_str().unwrap(),
            "--key", "id", "--json",
        ])
        .output()
        .expect("vgrid diff --json (quiet)");

    let stderr = String::from_utf8_lossy(&output.stderr);
    // --json implies --quiet: stderr should not have the ASCII summary banner
    assert!(!stderr.contains("====="), "stderr should not have summary banner with --json");

    std::fs::remove_file(&left).ok();
    std::fs::remove_file(&right).ok();
}

// ===========================================================================
// Cross-cutting: stdout must be ONLY JSON (no color codes, no banners)
// ===========================================================================

#[test]
fn peek_json_stdout_has_no_ansi_codes() {
    let csv = std::env::temp_dir().join("vgrid_peek_json_noansi.csv");
    std::fs::write(&csv, "a,b\n1,2\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--headers", "--json"])
        .output()
        .expect("vgrid peek --json");

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.contains('\x1b'), "stdout must not contain ANSI escape codes");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn calc_json_stdout_has_no_ansi_codes() {
    let csv = std::env::temp_dir().join("vgrid_calc_json_noansi.csv");
    std::fs::write(&csv, "1\n").unwrap();

    let output = vgrid()
        .args(["calc", "=A1+1", "-f", "csv", "--json"])
        .stdin(std::process::Stdio::from(std::fs::File::open(&csv).unwrap()))
        .output()
        .expect("vgrid calc --json");

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.contains('\x1b'), "stdout must not contain ANSI escape codes");

    std::fs::remove_file(&csv).ok();
}
