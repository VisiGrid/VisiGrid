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
