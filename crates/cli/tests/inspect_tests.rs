// Integration tests for `vgrid sheet inspect` multi-sheet, sparse, and NDJSON output.
// Run with: cargo test -p visigrid-cli --test inspect_tests -- --nocapture

use std::path::Path;
use std::process::Command;

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

fn template_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/abuse/templates/recon-template.sheet")
}

// ---------------------------------------------------------------------------
// --sheets: list sheets with correct count, names, dimensions
// ---------------------------------------------------------------------------

#[test]
fn inspect_sheets_lists_both_sheets() {
    let output = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheets", "--json"])
        .output()
        .expect("vgrid sheet inspect --sheets --json");

    assert!(output.status.success(), "exit code was {:?}", output.status);

    let stdout = String::from_utf8_lossy(&output.stdout);
    let entries: Vec<serde_json::Value> = serde_json::from_str(&stdout)
        .expect("valid JSON array");

    assert_eq!(entries.len(), 2, "expected 2 sheets");
    assert_eq!(entries[0]["name"], "Sheet1");
    assert_eq!(entries[1]["name"], "summary");
    assert_eq!(entries[1]["index"], 1);

    // summary sheet has known content
    let non_empty = entries[1]["non_empty_cells"].as_u64().unwrap();
    assert!(non_empty > 0, "summary sheet should have non-empty cells");
}

// ---------------------------------------------------------------------------
// --sheet 1: select sheet 1 and get known cell
// ---------------------------------------------------------------------------

#[test]
fn inspect_sheet_1_returns_known_cell() {
    let output = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheet", "1", "A1", "--json"])
        .output()
        .expect("vgrid sheet inspect --sheet 1 A1 --json");

    assert!(output.status.success());

    let stdout = String::from_utf8_lossy(&output.stdout);
    let cell: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    assert_eq!(cell["cell"], "A1");
    assert_eq!(cell["value"], "Category");
    assert_eq!(cell["value_type"], "text");
}

// ---------------------------------------------------------------------------
// --non-empty: sparse has fewer cells than dense for same range
// ---------------------------------------------------------------------------

#[test]
fn inspect_non_empty_is_sparse() {
    // Dense range A1:B20 on sheet 1 = 40 cells
    let dense = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheet", "1", "A1:B20", "--json"])
        .output()
        .expect("dense inspect");

    assert!(dense.status.success());
    let dense_json: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&dense.stdout)
    ).expect("valid JSON");
    let dense_count = dense_json["cells"].as_array().unwrap().len();

    // Sparse range A1:B20 on sheet 1 = only non-empty
    let sparse = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheet", "1", "A1:B20", "--non-empty", "--json"])
        .output()
        .expect("sparse inspect");

    assert!(sparse.status.success());
    let sparse_json: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&sparse.stdout)
    ).expect("valid JSON");
    let sparse_count = sparse_json["cells"].as_array().unwrap().len();

    assert!(sparse_count < dense_count,
        "sparse ({}) should have fewer cells than dense ({})", sparse_count, dense_count);
    assert!(sparse_count > 0, "sparse should have some cells");
}

// ---------------------------------------------------------------------------
// Golden: SparseInspectResult cell ordering is deterministic (row-major)
// ---------------------------------------------------------------------------

#[test]
fn sparse_inspect_ordering_is_deterministic() {
    let run = || -> String {
        let output = vgrid()
            .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheet", "1", "--non-empty", "--json"])
            .output()
            .expect("vgrid sheet inspect --non-empty --json");
        assert!(output.status.success());
        String::from_utf8(output.stdout).unwrap()
    };

    let a = run();
    let b = run();
    assert_eq!(a, b, "two runs of sparse inspect must produce byte-identical output");

    // Verify row-major ordering: parse cells and check (row, col) is non-decreasing
    let result: serde_json::Value = serde_json::from_str(&a).expect("valid JSON");
    let cells = result["cells"].as_array().unwrap();
    assert!(!cells.is_empty());

    let mut prev = (0usize, 0usize);
    for cell in cells {
        let cell_ref = cell["cell"].as_str().unwrap();
        // Parse cell ref to (row, col) for ordering check
        let (row, col) = parse_cell_ref(cell_ref);
        assert!((row, col) >= prev,
            "cell {} at ({},{}) is before previous ({},{})", cell_ref, row, col, prev.0, prev.1);
        prev = (row, col);
    }
}

// ---------------------------------------------------------------------------
// convert --sheet: sheet 0 != sheet 1; error on csv --sheet
// ---------------------------------------------------------------------------

#[test]
fn convert_sheet_selects_different_sheet() {
    let sheet0 = vgrid()
        .args(["convert", template_path().to_str().unwrap(), "-t", "csv", "--sheet", "0"])
        .output()
        .expect("convert --sheet 0");

    let sheet1 = vgrid()
        .args(["convert", template_path().to_str().unwrap(), "-t", "csv", "--sheet", "1"])
        .output()
        .expect("convert --sheet 1");

    assert!(sheet0.status.success());
    assert!(sheet1.status.success());

    let out0 = String::from_utf8_lossy(&sheet0.stdout);
    let out1 = String::from_utf8_lossy(&sheet1.stdout);
    assert_ne!(out0, out1, "sheet 0 and sheet 1 should produce different CSV output");
}

#[test]
fn convert_sheet_errors_on_csv() {
    // Create a temp CSV to test with
    let tmp = std::env::temp_dir().join("test_sheet_flag.csv");
    std::fs::write(&tmp, "a,b\n1,2\n").unwrap();

    let output = vgrid()
        .args(["convert", tmp.to_str().unwrap(), "-t", "json", "--sheet", "1"])
        .output()
        .expect("convert csv --sheet");

    assert!(!output.status.success(), "should fail with --sheet on csv");

    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("not supported for single-sheet"),
        "error should mention single-sheet formats, got: {}", stderr);

    std::fs::remove_file(&tmp).ok();
}

// ---------------------------------------------------------------------------
// --ndjson: produces valid NDJSON (one object per line)
// ---------------------------------------------------------------------------

#[test]
fn ndjson_produces_one_object_per_line() {
    let output = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--sheet", "1", "--non-empty", "--ndjson"])
        .output()
        .expect("vgrid sheet inspect --ndjson");

    assert!(output.status.success());

    let stdout = String::from_utf8(output.stdout).unwrap();
    let lines: Vec<&str> = stdout.trim().lines().collect();
    assert!(!lines.is_empty(), "should have at least one line");

    for (i, line) in lines.iter().enumerate() {
        let parsed: Result<serde_json::Value, _> = serde_json::from_str(line);
        assert!(parsed.is_ok(), "line {} is not valid JSON: {:?}", i, line);
        let obj = parsed.unwrap();
        assert!(obj["cell"].is_string(), "line {} should have cell field", i);
    }
}

// ===========================================================================
// CSV Tests
// ===========================================================================

fn csv_fixture(name: &str, content: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("vgrid_test_{}.csv", name));
    std::fs::write(&path, content).unwrap();
    path
}

fn tsv_fixture(name: &str, content: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("vgrid_test_{}.tsv", name));
    std::fs::write(&path, content).unwrap();
    path
}

fn xlsx_fixture_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/inspect_small.xlsx")
}

#[test]
fn inspect_csv_single_cell() {
    let csv = csv_fixture("single_cell", "Name,Age\nAlice,30\nBob,25\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "A1", "--json"])
        .output()
        .expect("csv single cell");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(cell["cell"], "A1");
    assert_eq!(cell["value"], "Name");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_csv_sheets_single_entry() {
    let csv = csv_fixture("sheets_single", "a,b\n1,2\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--sheets", "--json"])
        .output()
        .expect("csv --sheets");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let entries: Vec<serde_json::Value> = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON array");
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0]["index"], 0);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_csv_headers_column_name() {
    let csv = csv_fixture("headers", "Name,Age,City\nAlice,30,Paris\nBob,25,London\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "B2", "--headers", "--json"])
        .output()
        .expect("csv --headers B2");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(cell["cell"], "B2");
    assert_eq!(cell["value"], "30");
    assert_eq!(cell["column_name"], "Age");
    // Not a header row itself
    assert!(cell["header"].is_null(), "B2 should not be marked as header");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_csv_non_empty_ndjson_headers() {
    let csv = csv_fixture("ndjson_headers", "Name,Age\nAlice,30\nBob,25\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--headers", "--non-empty", "--ndjson"])
        .output()
        .expect("csv --headers --non-empty --ndjson");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let stdout = String::from_utf8(output.stdout).unwrap();
    let lines: Vec<&str> = stdout.trim().lines().collect();
    assert!(!lines.is_empty());

    // Every line should have column_name
    for line in &lines {
        let obj: serde_json::Value = serde_json::from_str(line).expect("valid NDJSON line");
        assert!(obj["column_name"].is_string(),
            "expected column_name in NDJSON line: {}", line);
    }
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_csv_rejects_sheet_flag() {
    let csv = csv_fixture("rejects_sheet", "a,b\n1,2\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--sheet", "0"])
        .output()
        .expect("csv --sheet 0");

    assert!(!output.status.success(), "should fail with --sheet on csv");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("not valid for CSV/TSV"),
        "error should mention CSV/TSV, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// TSV Tests
// ===========================================================================

#[test]
fn inspect_tsv_non_empty() {
    let tsv = tsv_fixture("basic", "Name\tAge\nAlice\t30\nBob\t25\n");
    let output = vgrid()
        .args(["sheet", "inspect", tsv.to_str().unwrap(), "--non-empty", "--json"])
        .output()
        .expect("tsv --non-empty");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let cells = result["cells"].as_array().unwrap();
    assert_eq!(cells.len(), 6, "3 rows x 2 cols = 6 non-empty cells");
    std::fs::remove_file(&tsv).ok();
}

// ===========================================================================
// XLSX Tests
// ===========================================================================

#[test]
fn inspect_xlsx_sheets_lists_real_names() {
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(), "--sheets", "--json"])
        .output()
        .expect("xlsx --sheets");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let entries: Vec<serde_json::Value> = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON array");
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0]["name"], "Invoices");
}

#[test]
fn inspect_xlsx_sheet_selection() {
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(),
               "--sheet", "Invoices", "A1:C3", "--json"])
        .output()
        .expect("xlsx --sheet Invoices A1:C3");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let cells = result["cells"].as_array().unwrap();
    assert_eq!(cells.len(), 9, "3x3 = 9 cells");

    // Check A1 is "Amount"
    assert_eq!(cells[0]["cell"], "A1");
    assert_eq!(cells[0]["value"], "Amount");
}

#[test]
fn inspect_xlsx_invalid_sheet_lists_available() {
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(),
               "--sheet", "Bogus", "A1", "--json"])
        .output()
        .expect("xlsx --sheet Bogus");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("Invoices"),
        "error should list available sheet 'Invoices', got: {}", stderr);
}

#[test]
fn inspect_xlsx_values_first() {
    // C2 should show cached value 110, not recalculated, with formula =A2+B2
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(),
               "--sheet", "Invoices", "C2", "--json"])
        .output()
        .expect("xlsx C2 values first");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(cell["cell"], "C2");
    // Cached value should be "110" (as number or string representation)
    let value = cell["value"].as_str().unwrap();
    assert!(value == "110" || value == "110.0",
        "expected cached value 110, got: {}", value);
    // Formula should be present
    let formula = cell["formula"].as_str().unwrap();
    assert!(formula.contains("A2+B2"),
        "expected formula containing A2+B2, got: {}", formula);
}

// ===========================================================================
// Backward Compat Tests (critical)
// ===========================================================================

#[test]
fn inspect_sheet_workbook_json_byte_identical() {
    let output = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "--workbook", "--json"])
        .output()
        .expect("sheet --workbook --json");

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).unwrap();
    let result: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    // Must have fingerprint as non-null string
    assert!(result["fingerprint"].is_string(), "fingerprint must be a string");
    let fp = result["fingerprint"].as_str().unwrap();
    assert!(fp.starts_with("v"), "fingerprint should start with version prefix, got: {}", fp);

    // Must NOT have format, path, or import_notes
    assert!(result.get("format").is_none() || result["format"].is_null(),
        "native .sheet should not have 'format' field");
    assert!(result.get("path").is_none() || result["path"].is_null(),
        "native .sheet should not have 'path' field");
    assert!(result.get("import_notes").is_none() || result["import_notes"].is_null(),
        "native .sheet should not have 'import_notes' field");
}

#[test]
fn inspect_sheet_cell_json_unchanged() {
    let output = vgrid()
        .args(["sheet", "inspect", template_path().to_str().unwrap(), "A1", "--json"])
        .output()
        .expect("sheet A1 --json");

    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).unwrap();
    let cell: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    // Must NOT have header or column_name
    assert!(cell.get("header").is_none() || cell["header"].is_null(),
        "native .sheet without --headers should not have 'header' field");
    assert!(cell.get("column_name").is_none() || cell["column_name"].is_null(),
        "native .sheet without --headers should not have 'column_name' field");
}

// ===========================================================================
// Format Override Test
// ===========================================================================

#[test]
fn inspect_format_override() {
    let path = std::env::temp_dir().join("vgrid_test_override.txt");
    std::fs::write(&path, "Name,Age\nAlice,30\n").unwrap();

    let output = vgrid()
        .args(["sheet", "inspect", path.to_str().unwrap(), "A1", "--format", "csv", "--json"])
        .output()
        .expect("--format csv on .txt");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(cell["value"], "Name");
    std::fs::remove_file(&path).ok();
}

// ===========================================================================
// --calc Tests
// ===========================================================================

#[test]
fn inspect_calc_csv_sum() {
    let csv = csv_fixture("calc_sum", "Amount\n100\n200\n300\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--calc", "SUM(A:A)"])
        .output()
        .expect("calc sum");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["format"], "csv");
    assert_eq!(result["sheet"], "Sheet1");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results.len(), 1);
    assert_eq!(results[0]["expr"], "SUM(A:A)");
    assert_eq!(results[0]["value"], "600");
    assert_eq!(results[0]["value_type"], "number");
    assert!(results[0].get("error").is_none() || results[0]["error"].is_null());
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_csv_multiple() {
    let csv = csv_fixture("calc_multi", "Val\n10\n20\n30\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--calc", "SUM(A:A)", "--calc", "AVERAGE(A:A)"])
        .output()
        .expect("calc multiple");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results.len(), 2);
    assert_eq!(results[0]["expr"], "SUM(A:A)");
    assert_eq!(results[1]["expr"], "AVERAGE(A:A)");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_csv_headers_skips_header() {
    let csv = csv_fixture("calc_headers", "Amount\n100\n200\n300\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(A:A)"])
        .output()
        .expect("calc with headers");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    // With --headers, SUM(A:A) should skip "Amount" header → 100+200+300 = 600
    assert_eq!(results[0]["value"], "600");
    assert_eq!(results[0]["value_type"], "number");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_error_exits_nonzero() {
    let csv = csv_fixture("calc_err", "Val\n1\n2\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--calc", "1/0"])
        .output()
        .expect("calc 1/0");

    assert!(!output.status.success(), "should exit nonzero for error");
    // stdout should still be valid JSON
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("stdout should be valid JSON even on error");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value_type"], "error");
    assert!(results[0]["error"].is_string());
    // stderr should be empty (clean stdout)
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.is_empty(), "stderr should be empty, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_parse_error() {
    let csv = csv_fixture("calc_parse", "Val\n1\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(), "--calc", "SUM("])
        .output()
        .expect("calc parse error");

    assert!(!output.status.success(), "should exit nonzero for parse error");
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("stdout should be valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value_type"], "error");
    assert!(results[0]["error"].is_string());
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_xlsx() {
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(),
               "--sheet", "Invoices", "--calc", "SUM(A:A)"])
        .output()
        .expect("calc xlsx");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["format"], "xlsx");
    assert_eq!(result["sheet"], "Invoices");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value_type"], "number");
}

#[test]
fn inspect_calc_conflicts_with_workbook() {
    let csv = csv_fixture("calc_wb", "a\n1\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--calc", "SUM(A:A)", "--workbook"])
        .output()
        .expect("calc + workbook");

    assert!(!output.status.success(), "should fail with --calc + --workbook");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("--calc cannot be used with --workbook"),
        "expected conflict error, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_conflicts_with_sheets() {
    let csv = csv_fixture("calc_sheets", "a\n1\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--calc", "SUM(A:A)", "--sheets"])
        .output()
        .expect("calc + sheets");

    assert!(!output.status.success(), "should fail with --calc + --sheets");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("--calc cannot be used with --sheets"),
        "expected conflict error, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_ndjson_conflict() {
    let csv = csv_fixture("calc_ndjson", "a\n1\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--calc", "SUM(A:A)", "--ndjson"])
        .output()
        .expect("calc + ndjson");

    assert!(!output.status.success(), "should fail with --calc + --ndjson");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("--calc cannot be used with --ndjson"),
        "expected conflict error, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_empty_sheet_with_headers() {
    // CSV with only a header row — no data rows
    let csv = csv_fixture("calc_empty", "Amount\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(A:A)"])
        .output()
        .expect("calc empty with headers");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    // Should return 0 or empty, not crash
    let value = results[0]["value"].as_str().unwrap();
    assert!(value == "0" || value.is_empty(),
        "expected 0 or empty for empty sheet, got: {}", value);
    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// Semantic Calc Tests (header-name references)
// ===========================================================================

#[test]
fn inspect_calc_header_name_sum() {
    let csv = csv_fixture("calc_hdr_sum", "Amount,Tax\n100,10\n200,20\n300,30\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(Amount)"])
        .output()
        .expect("calc header name sum");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["expr"], "SUM(Amount)");
    assert_eq!(results[0]["value"], "600");
    assert_eq!(results[0]["value_type"], "number");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_case_insensitive() {
    let csv = csv_fixture("calc_hdr_case", "Amount\n100\n200\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(amount)"])
        .output()
        .expect("calc header case-insensitive");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value"], "300");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_bracket_syntax() {
    let csv = csv_fixture("calc_hdr_bracket", "WO Number,Amount\n1001,100\n1002,200\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM([WO Number])"])
        .output()
        .expect("calc bracket syntax");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value"], "2003");
    assert_eq!(results[0]["value_type"], "number");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_xlsx_total() {
    // XLSX fixture: Amount(100,200) Tax(10,20) Total(110,220)
    // SUM(Total) = 110 + 220 = 330
    let output = vgrid()
        .args(["sheet", "inspect", xlsx_fixture_path().to_str().unwrap(),
               "--sheet", "Invoices", "--headers", "--calc", "SUM(Total)"])
        .output()
        .expect("calc xlsx header Total");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value"], "330");
    assert_eq!(results[0]["value_type"], "number");
}

#[test]
fn inspect_calc_header_nonexistent() {
    let csv = csv_fixture("calc_hdr_noexist", "Amount\n100\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(DoesNotExist)"])
        .output()
        .expect("calc nonexistent header");

    // Should fail — DoesNotExist is not a header and not a valid formula token
    assert!(!output.status.success(), "should exit nonzero for unknown identifier");
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("stdout should be valid JSON even on error");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value_type"], "error");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_mixed_with_column_ref() {
    // Mix header names and column refs in the same invocation
    let csv = csv_fixture("calc_hdr_mixed", "Amount,Tax\n100,10\n200,20\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(Amount)", "--calc", "SUM(B:B)"])
        .output()
        .expect("calc mixed header + column ref");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results.len(), 2);
    assert_eq!(results[0]["value"], "300"); // SUM(Amount) → A2:A3
    assert_eq!(results[1]["value"], "30");  // SUM(B:B) → B2:B3
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_no_resolve_without_headers_flag() {
    // Without --headers, "Amount" should NOT be resolved — should error
    let csv = csv_fixture("calc_hdr_noflag", "Amount\n100\n200\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--calc", "SUM(Amount)"])
        .output()
        .expect("calc without --headers");

    assert!(!output.status.success(), "should fail without --headers");
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("stdout should be valid JSON even on error");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value_type"], "error");
    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// Adversarial header names + ambiguity
// ===========================================================================

#[test]
fn inspect_calc_header_named_if() {
    // Header "IF" collides with formula function name.
    // IF(1>0,1,0) must still work as function (not rewritten to column ref).
    // SUM([IF]) must resolve header "IF" via bracket syntax.
    let csv = csv_fixture("calc_hdr_if", "IF,Amount\n10,100\n20,200\n");

    // IF(1>0,1,0) → IF is followed by '(' → treated as function, not header
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "IF(1>0, 1, 0)"])
        .output()
        .expect("calc IF as function");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value"], "1");
    assert_eq!(results[0]["value_type"], "number");

    // SUM([IF]) → bracket syntax resolves header "IF" to column A
    let output2 = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM([IF])"])
        .output()
        .expect("calc bracket IF header");

    assert!(output2.status.success(), "stderr: {}", String::from_utf8_lossy(&output2.stderr));
    let result2: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output2.stdout)
    ).expect("valid JSON");
    let results2 = result2["results"].as_array().unwrap();
    assert_eq!(results2[0]["value"], "30"); // 10+20
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_header_named_sum() {
    // Header "SUM" collides with formula function name.
    // SUM(SUM) should work: outer SUM is function (followed by paren),
    // inner SUM is bare identifier matching header.
    let csv = csv_fixture("calc_hdr_sum_name", "SUM,Other\n100,1\n200,2\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(SUM)"])
        .output()
        .expect("calc header named SUM");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    let results = result["results"].as_array().unwrap();
    assert_eq!(results[0]["value"], "300"); // 100+200
    std::fs::remove_file(&csv).ok();
}

#[test]
fn inspect_calc_ambiguous_header_errors() {
    // Two headers that normalize to the same key: "Amount" and "AMOUNT"
    let csv = csv_fixture("calc_hdr_ambig", "Amount,AMOUNT\n100,200\n");
    let output = vgrid()
        .args(["sheet", "inspect", csv.to_str().unwrap(),
               "--headers", "--calc", "SUM(Amount)"])
        .output()
        .expect("calc ambiguous header");

    assert!(!output.status.success(), "should fail on ambiguous headers");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("ambiguous header"),
        "expected ambiguous header error, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

// ===========================================================================
// Import Tests
// ===========================================================================

fn sheet_output(name: &str) -> std::path::PathBuf {
    std::env::temp_dir().join(format!("vgrid_import_test_{}.sheet", name))
}

#[test]
fn import_csv_creates_valid_sheet() {
    let csv = csv_fixture("import_basic", "Amount,Tax\n100,10\n200,20\n300,30\n");
    let out = sheet_output("basic");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--json"])
        .output()
        .expect("import csv");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    assert_eq!(result["format"], "csv");
    assert_eq!(result["sheet"], "Sheet1");
    assert!(result["rows"].as_u64().unwrap() > 0);
    assert!(result["cols"].as_u64().unwrap() > 0);
    assert!(result["cells"].as_u64().unwrap() > 0);
    assert!(result["fingerprint"].as_str().unwrap().starts_with("v"));
    assert!(result["output"].as_str().is_some());

    // Verify produced file can be inspected
    let inspect = vgrid()
        .args(["sheet", "inspect", out.to_str().unwrap(), "A1", "--json"])
        .output()
        .expect("inspect produced sheet");
    assert!(inspect.status.success());
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&inspect.stdout)
    ).expect("valid JSON");
    assert_eq!(cell["value"], "Amount");

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_csv_fingerprint_stable() {
    let csv = csv_fixture("import_fp_stable", "A,B\n1,2\n3,4\n");
    let out1 = sheet_output("fp1");
    let out2 = sheet_output("fp2");

    let r1 = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out1.to_str().unwrap(), "--json"])
        .output().expect("import 1");
    let r2 = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out2.to_str().unwrap(), "--json"])
        .output().expect("import 2");

    assert!(r1.status.success());
    assert!(r2.status.success());

    let j1: serde_json::Value = serde_json::from_str(&String::from_utf8_lossy(&r1.stdout)).unwrap();
    let j2: serde_json::Value = serde_json::from_str(&String::from_utf8_lossy(&r2.stdout)).unwrap();
    assert_eq!(j1["fingerprint"], j2["fingerprint"], "same CSV imported twice must yield same fingerprint");

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out1).ok();
    std::fs::remove_file(&out2).ok();
}

#[test]
fn import_xlsx_sheet_selects_correct_sheet() {
    let out = sheet_output("xlsx_sheet_select");
    let output = vgrid()
        .args(["sheet", "import", xlsx_fixture_path().to_str().unwrap(),
               out.to_str().unwrap(), "--sheet", "0", "--json"])
        .output()
        .expect("import xlsx --sheet 0");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    assert_eq!(result["sheet"], "Invoices");
    assert_eq!(result["format"], "xlsx");

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_xlsx_formulas_values_default() {
    let out = sheet_output("xlsx_fv");
    let output = vgrid()
        .args(["sheet", "import", xlsx_fixture_path().to_str().unwrap(),
               out.to_str().unwrap(), "--formulas", "values", "--json"])
        .output()
        .expect("import xlsx --formulas values");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    let formulas = &result["formulas"];
    assert_eq!(formulas["policy"], "values");
    assert!(formulas["captured"].as_u64().unwrap() > 0, "should count formula strings from source");

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_xlsx_formulas_keep() {
    let out = sheet_output("xlsx_fk");
    let output = vgrid()
        .args(["sheet", "import", xlsx_fixture_path().to_str().unwrap(),
               out.to_str().unwrap(), "--formulas", "keep", "--json"])
        .output()
        .expect("import xlsx --formulas keep");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    let formulas = &result["formulas"];
    assert_eq!(formulas["policy"], "keep");
    assert!(formulas["captured"].as_u64().unwrap() > 0, "should capture formula strings");

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_xlsx_formulas_recalc() {
    let out = sheet_output("xlsx_fr");
    let output = vgrid()
        .args(["sheet", "import", xlsx_fixture_path().to_str().unwrap(),
               out.to_str().unwrap(), "--formulas", "recalc", "--json"])
        .output()
        .expect("import xlsx --formulas recalc");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    let formulas = &result["formulas"];
    assert_eq!(formulas["policy"], "recalc");
    assert!(formulas["kept"].as_u64().unwrap() > 0, "should have kept formulas from recalc");

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_stamp_then_verify_passes() {
    let csv = csv_fixture("import_stamp", "X\n1\n2\n");
    let out = sheet_output("stamped");
    let import = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--stamp", "test", "--json"])
        .output()
        .expect("import with stamp");
    assert!(import.status.success(), "stderr: {}", String::from_utf8_lossy(&import.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&import.stdout)
    ).expect("valid JSON");
    assert_eq!(result["stamped"], true);

    // Verify the stamped file
    let verify = vgrid()
        .args(["sheet", "verify", out.to_str().unwrap()])
        .output()
        .expect("verify stamped");
    assert!(verify.status.success(), "verify should pass for stamped file. stderr: {}",
        String::from_utf8_lossy(&verify.stderr));

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_verify_mismatch_exits_1() {
    let csv = csv_fixture("import_verify_fail", "A\n1\n");
    let out = sheet_output("verify_fail");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(),
               "--verify", "v2:999:0000000000000000000000000000000000000000000000000000000000000000",
               "--json"])
        .output()
        .expect("import with verify mismatch");

    assert!(!output.status.success(), "should exit nonzero on mismatch");
    assert_eq!(output.status.code(), Some(1));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("stdout should be valid JSON");
    assert_eq!(result["ok"], false);
    assert_eq!(result["error"], "fingerprint_mismatch");

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_json_output_has_expected_fields() {
    let csv = csv_fixture("import_fields", "Name,Age\nAlice,30\n");
    let out = sheet_output("fields");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--json"])
        .output()
        .expect("import json fields");

    assert!(output.status.success());
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert!(result["ok"].as_bool().is_some());
    assert!(result["source"].as_str().is_some());
    assert!(result["format"].as_str().is_some());
    assert!(result["sheet"].as_str().is_some());
    assert!(result["rows"].as_u64().is_some());
    assert!(result["cols"].as_u64().is_some());
    assert!(result["cells"].as_u64().is_some());
    assert!(result["fingerprint"].as_str().is_some());

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_sheet_flag_on_csv_errors() {
    let csv = csv_fixture("import_sheet_csv", "a,b\n1,2\n");
    let out = sheet_output("sheet_csv");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--sheet", "0"])
        .output()
        .expect("import csv --sheet");

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("XLSX"), "error should mention XLSX, got: {}", stderr);

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_formulas_keep_on_csv_errors() {
    let csv = csv_fixture("import_fk_csv", "a,b\n1,2\n");
    let out = sheet_output("fk_csv");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--formulas", "keep"])
        .output()
        .expect("import csv --formulas keep");

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("XLSX"), "error should mention XLSX, got: {}", stderr);

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_empty_csv() {
    let csv = csv_fixture("import_empty", "");
    let out = sheet_output("empty");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(), "--json"])
        .output()
        .expect("import empty csv");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    assert_eq!(result["cells"], 0);

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_tsv() {
    let tsv = tsv_fixture("import_basic", "Name\tAge\nAlice\t30\n");
    let out = sheet_output("tsv_basic");
    let output = vgrid()
        .args(["sheet", "import", tsv.to_str().unwrap(), out.to_str().unwrap(), "--json"])
        .output()
        .expect("import tsv");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    assert_eq!(result["format"], "tsv");

    std::fs::remove_file(&tsv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_sheet_to_sheet_rejected() {
    let out = sheet_output("sheet_reject");
    let output = vgrid()
        .args(["sheet", "import", template_path().to_str().unwrap(), out.to_str().unwrap()])
        .output()
        .expect("import .sheet");

    assert!(!output.status.success());
    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("already .sheet"), "error should mention .sheet, got: {}", stderr);

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_values_policy_does_not_recalc() {
    // Import xlsx with --formulas values, inspect a formula cell: should show cached value
    let out = sheet_output("values_no_recalc");
    let import = vgrid()
        .args(["sheet", "import", xlsx_fixture_path().to_str().unwrap(),
               out.to_str().unwrap(), "--formulas", "values", "--json"])
        .output()
        .expect("import xlsx values");
    assert!(import.status.success(), "stderr: {}", String::from_utf8_lossy(&import.stderr));

    // Inspect C2 — in inspect_small.xlsx, C2 has formula =A2+B2 with cached value 110
    let inspect = vgrid()
        .args(["sheet", "inspect", out.to_str().unwrap(), "C2", "--json"])
        .output()
        .expect("inspect C2");
    assert!(inspect.status.success());
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&inspect.stdout)
    ).expect("valid JSON");
    let value = cell["value"].as_str().unwrap();
    assert!(value == "110" || value == "110.0",
        "expected cached value 110, got: {}", value);
    // Should NOT have a formula (values mode strips formulas)
    assert!(cell.get("formula").is_none() || cell["formula"].is_null(),
        "values mode should not store formulas in cells");

    std::fs::remove_file(&out).ok();
}

#[test]
fn import_nulls_error_does_not_dense_fill_outside_bounds() {
    // CSV with data in A1:B3 only
    let csv = csv_fixture("import_nulls", "X,Y\n1,2\n3,4\n");
    let out = sheet_output("nulls_bounds");
    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(),
               "--nulls", "error", "--json"])
        .output()
        .expect("import with nulls error");
    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    // Cell outside bounds (C1) should still be empty, not #NULL!
    let inspect = vgrid()
        .args(["sheet", "inspect", out.to_str().unwrap(), "C1", "--json"])
        .output()
        .expect("inspect C1");
    assert!(inspect.status.success());
    let cell: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&inspect.stdout)
    ).expect("valid JSON");
    let val = cell["value"].as_str().unwrap_or("");
    assert!(val.is_empty() || val != "#NULL!",
        "cell C1 outside data bounds should not be #NULL!, got: {:?}", val);

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&out).ok();
}

#[test]
fn import_dry_run_no_file() {
    let csv = csv_fixture("import_dryrun", "A\n1\n2\n");
    let out = sheet_output("dryrun");
    // Make sure output doesn't exist
    std::fs::remove_file(&out).ok();

    let output = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), out.to_str().unwrap(),
               "--dry-run", "--json"])
        .output()
        .expect("import dry run");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let result: serde_json::Value = serde_json::from_str(
        &String::from_utf8_lossy(&output.stdout)
    ).expect("valid JSON");
    assert_eq!(result["ok"], true);
    assert_eq!(result["dry_run"], true);
    assert!(!out.exists(), "output file should NOT exist after dry-run");

    std::fs::remove_file(&csv).ok();
}

// ---------------------------------------------------------------------------
// ===========================================================================
// Hub Publish Tests (local validation, no network)
// ===========================================================================

#[test]
fn hub_publish_rejects_non_sheet() {
    let csv = csv_fixture("hub_non_sheet", "a,b\n1,2\n");
    let output = vgrid()
        .args(["hub", "publish", csv.to_str().unwrap(), "--repo", "x/y"])
        .output()
        .expect("hub publish non-sheet");

    assert!(!output.status.success(), "should fail for non-.sheet file");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains(".sheet") || stderr.contains("sheet"),
        "error should mention .sheet, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn hub_publish_rejects_bad_repo() {
    // Create a temp .sheet file (doesn't need to be valid — repo check is first after file check)
    let tmp = std::env::temp_dir().join("hub_bad_repo_test.sheet");
    std::fs::write(&tmp, b"not a real sheet but has .sheet extension").unwrap();

    let output = vgrid()
        .args(["hub", "publish", tmp.to_str().unwrap(), "--repo", "noslash"])
        .output()
        .expect("hub publish bad repo");

    assert!(!output.status.success(), "should fail for bad repo format");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("owner/slug") || stderr.contains("Invalid repo"),
        "error should mention repo format, got: {}", stderr);
    std::fs::remove_file(&tmp).ok();
}

#[test]
fn hub_publish_rejects_missing_file() {
    let output = vgrid()
        .args(["hub", "publish", "/tmp/nonexistent_9999.sheet", "--repo", "x/y"])
        .output()
        .expect("hub publish missing file");

    assert!(!output.status.success(), "should fail for missing file");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("not found") || stderr.contains("not found") || stderr.contains("No such file") || stderr.contains("File not found"),
        "error should mention file not found, got: {}", stderr);
}

#[test]
fn hub_publish_dry_run_schema() {
    // Import a CSV to .sheet first
    let csv = csv_fixture("hub_dry_run_src", "Amount,Tax\n100,10\n200,20\n");
    let sheet = std::env::temp_dir().join("hub_dry_run_test.sheet");

    let import_out = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), sheet.to_str().unwrap(), "--headers"])
        .output()
        .expect("import for dry-run test");
    assert!(import_out.status.success(), "import failed: {}",
        String::from_utf8_lossy(&import_out.stderr));

    let output = vgrid()
        .args(["hub", "publish", sheet.to_str().unwrap(), "--repo", "x/y", "--dry-run", "--json"])
        .output()
        .expect("hub publish dry-run");

    assert!(output.status.success(), "dry-run should succeed without auth, stderr: {}",
        String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout)
        .expect("dry-run should produce valid JSON");

    // Schema contract: all required fields present
    assert_eq!(json["schema_version"], 1, "schema_version must be 1");
    assert_eq!(json["dry_run"], true);
    assert_eq!(json["repo"], "x/y");
    assert!(json["fingerprint"].is_string(), "should have fingerprint");
    let fp = json["fingerprint"].as_str().unwrap();
    assert!(fp.starts_with("v"), "fingerprint should start with version prefix, got: {}", fp);

    // Stamp semantics: unstamped file
    assert_eq!(json["stamped"], false, "unstamped import should have stamped=false");
    assert_eq!(json["stamp_matches"], false, "unstamped import should have stamp_matches=false");

    // Optional fields present
    assert!(json["message"].is_string());
    assert!(json["source_metadata"].is_object());
    assert!(json["checks_attached"].is_boolean());
    assert!(json["notes_attached"].is_boolean());

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&sheet).ok();
}

#[test]
fn hub_publish_dry_run_stamped_file() {
    // Import with --stamp to get a stamped .sheet
    let csv = csv_fixture("hub_stamp_src", "Amount,Tax\n100,10\n200,20\n");
    let sheet = std::env::temp_dir().join("hub_stamp_dry_run.sheet");

    let import_out = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), sheet.to_str().unwrap(),
               "--headers", "--stamp", "Q4 Filing"])
        .output()
        .expect("stamped import");
    assert!(import_out.status.success(), "import failed: {}",
        String::from_utf8_lossy(&import_out.stderr));

    let output = vgrid()
        .args(["hub", "publish", sheet.to_str().unwrap(), "--repo", "x/y", "--dry-run", "--json"])
        .output()
        .expect("hub publish dry-run stamped");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    assert_eq!(json["schema_version"], 1);
    assert_eq!(json["stamped"], true, "stamped file should have stamped=true");
    assert_eq!(json["stamp_matches"], true, "freshly stamped file should match");

    // trust_pipeline.stamp should be present in source_metadata
    let sm = &json["source_metadata"];
    assert_eq!(sm["type"], "trust_pipeline");
    assert!(sm["trust_pipeline"]["stamp"]["expected_fingerprint"].is_string());

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&sheet).ok();
}

#[test]
fn hub_publish_checks_size_guard() {
    // Import a CSV to .sheet
    let csv = csv_fixture("hub_checks_guard_src", "a\n1\n");
    let sheet = std::env::temp_dir().join("hub_checks_guard.sheet");
    let checks_path = std::env::temp_dir().join("hub_oversized_checks.json");

    let import_out = vgrid()
        .args(["sheet", "import", csv.to_str().unwrap(), sheet.to_str().unwrap(), "--headers"])
        .output()
        .expect("import");
    assert!(import_out.status.success());

    // Create an oversized checks file (> 256 KB)
    let big_json = format!("{{\"data\": \"{}\"}}", "x".repeat(300 * 1024));
    std::fs::write(&checks_path, &big_json).unwrap();

    let output = vgrid()
        .args(["hub", "publish", sheet.to_str().unwrap(), "--repo", "x/y",
               "--checks", checks_path.to_str().unwrap(), "--dry-run", "--json"])
        .output()
        .expect("hub publish oversized checks");

    assert!(!output.status.success(), "should fail for oversized checks");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("too large"), "error should mention size, got: {}", stderr);

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&sheet).ok();
    std::fs::remove_file(&checks_path).ok();
}

// ===========================================================================
// Pipeline Publish Tests (local validation + dry-run, no network)
// ===========================================================================

#[test]
fn pipeline_publish_rejects_sheet_source() {
    let tmp = std::env::temp_dir().join("pipeline_reject_sheet.sheet");
    std::fs::write(&tmp, b"placeholder").unwrap();

    let output = vgrid()
        .args(["pipeline", "publish", tmp.to_str().unwrap(), "--repo", "x/y"])
        .output()
        .expect("pipeline publish .sheet source");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("already .sheet") || stderr.contains("hub publish"),
        "error should suggest hub publish, got: {}", stderr);
    std::fs::remove_file(&tmp).ok();
}

#[test]
fn pipeline_publish_rejects_bad_repo() {
    let csv = csv_fixture("pipeline_bad_repo", "a\n1\n");
    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(), "--repo", "noslash"])
        .output()
        .expect("pipeline publish bad repo");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("owner/slug") || stderr.contains("Invalid repo"),
        "error should mention repo format, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn pipeline_publish_dry_run_csv() {
    let csv = csv_fixture("pipeline_dry_csv", "Amount,Tax\n100,10\n200,20\n");
    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(),
               "--repo", "x/y", "--headers", "--dry-run", "--json"])
        .output()
        .expect("pipeline publish dry-run CSV");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    assert_eq!(json["schema_version"], 1);
    assert_eq!(json["dry_run"], true);
    assert_eq!(json["repo"], "x/y");
    assert_eq!(json["format"], "csv");
    assert!(json["fingerprint"].is_string());
    assert!(json["rows"].as_u64().unwrap() > 0);
    assert!(json["cols"].as_u64().unwrap() > 0);
    assert_eq!(json["stamped"], false);
    assert_eq!(json["stamp_matches"], false);
    assert_eq!(json["checks_attached"], false);
    assert!(json["source_metadata"]["type"].as_str() == Some("trust_pipeline"));

    std::fs::remove_file(&csv).ok();
}

#[test]
fn pipeline_publish_dry_run_with_stamp() {
    let csv = csv_fixture("pipeline_dry_stamp", "Amount,Tax\n100,10\n200,20\n");
    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(),
               "--repo", "x/y", "--headers", "--stamp", "Q4 Close",
               "--dry-run", "--json"])
        .output()
        .expect("pipeline publish dry-run stamped");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    assert_eq!(json["schema_version"], 1);
    assert_eq!(json["stamped"], true);
    assert_eq!(json["stamp_matches"], true);

    let sm = &json["source_metadata"];
    assert!(sm["trust_pipeline"]["stamp"]["expected_fingerprint"].is_string());
    assert_eq!(sm["trust_pipeline"]["stamp"]["label"], "Q4 Close");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn pipeline_publish_dry_run_with_checks_calc() {
    let csv = csv_fixture("pipeline_dry_checks", "Amount,Tax\n100,10\n200,20\n");
    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(),
               "--repo", "x/y", "--headers",
               "--checks-calc", "SUM(Amount)", "--checks-calc", "SUM(Tax)",
               "--dry-run", "--json"])
        .output()
        .expect("pipeline publish dry-run with checks-calc");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");

    assert_eq!(json["schema_version"], 1);
    assert_eq!(json["checks_attached"], true);

    // Checks should contain the calc results
    let checks = &json["source_metadata"]["trust_pipeline"]["checks"];
    assert!(checks.is_object(), "checks should be an object");
    let results = checks["results"].as_array().expect("checks should have results array");
    assert_eq!(results.len(), 2);
    assert_eq!(results[0]["expr"], "SUM(Amount)");
    assert_eq!(results[0]["value"], "300");
    assert_eq!(results[1]["expr"], "SUM(Tax)");
    assert_eq!(results[1]["value"], "30");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn pipeline_publish_dry_run_saves_sheet_with_out() {
    let csv = csv_fixture("pipeline_out_flag", "a,b\n1,2\n");
    let sheet_out = std::env::temp_dir().join("pipeline_out_test.sheet");

    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(),
               "--repo", "x/y", "--headers",
               "--out", sheet_out.to_str().unwrap(),
               "--dry-run", "--json"])
        .output()
        .expect("pipeline publish dry-run with --out");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    // .sheet file should exist because --out was given
    assert!(sheet_out.exists(), ".sheet file should be saved with --out");

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    assert_eq!(json["sheet_path"], sheet_out.display().to_string());

    std::fs::remove_file(&csv).ok();
    std::fs::remove_file(&sheet_out).ok();
}

#[test]
fn pipeline_publish_dry_run_no_out_cleans_temp() {
    let csv = csv_fixture("pipeline_no_out", "a,b\n1,2\n");

    let output = vgrid()
        .args(["pipeline", "publish", csv.to_str().unwrap(),
               "--repo", "x/y", "--headers",
               "--dry-run", "--json"])
        .output()
        .expect("pipeline publish dry-run without --out");

    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    let json: serde_json::Value = serde_json::from_str(&stdout).expect("valid JSON");
    // sheet_path should be null when temp
    assert!(json["sheet_path"].is_null(), "temp path should not be exposed");

    std::fs::remove_file(&csv).ok();
}

// Helpers
// ---------------------------------------------------------------------------

fn parse_cell_ref(s: &str) -> (usize, usize) {
    let s = s.to_uppercase();
    let mut col_str = String::new();
    let mut row_str = String::new();
    for c in s.chars() {
        if c.is_ascii_alphabetic() && row_str.is_empty() {
            col_str.push(c);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        }
    }
    let mut col: usize = 0;
    for c in col_str.chars() {
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }
    col -= 1;
    let row: usize = row_str.parse().unwrap();
    (row - 1, col)
}
