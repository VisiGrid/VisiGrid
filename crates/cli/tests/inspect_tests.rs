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

// ---------------------------------------------------------------------------
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
