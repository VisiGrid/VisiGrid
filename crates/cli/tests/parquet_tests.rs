// Integration tests for Parquet through `vgrid convert` and `vgrid sheet inspect`.
// Run with: cargo test -p visigrid-cli --test parquet_tests
//
// Fixtures (tests/fixtures/, written with pyarrow):
//   parquet_orders.parquet      order_id string ("007"…), placed_at timestamp[us]
//                               (one with .123 s), ship_date date32,
//                               amount decimal(10,2), status string; 3 rows
//   parquet_70000_rows.parquet  one int32 column `n`, every value 1; 70,000
//                               rows, more than a sheet holds

use std::path::PathBuf;
use std::process::{Command, Output};

fn vgrid(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .current_dir(env!("CARGO_MANIFEST_DIR"))
        .args(args)
        .output()
        .expect("run vgrid")
}

fn fixture(name: &str) -> String {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
        .to_string_lossy()
        .into_owned()
}

fn stdout(output: &Output) -> String {
    String::from_utf8_lossy(&output.stdout).into_owned()
}

fn stderr(output: &Output) -> String {
    String::from_utf8_lossy(&output.stderr).into_owned()
}

// ---------------------------------------------------------------------------
// Oversized files: an aggregate over the rows that fit is a wrong answer
// ---------------------------------------------------------------------------

/// `--calc` used to total the first 65,535 records of a 70,000-row file and
/// exit 0 with no warning.
#[test]
fn inspect_calc_refuses_a_file_bigger_than_a_sheet() {
    let out = vgrid(&["sheet", "inspect", &fixture("parquet_70000_rows.parquet"), "--headers", "--calc", "SUM(n)"]);

    assert_eq!(out.status.code(), Some(4), "stdout: {}\nstderr: {}", stdout(&out), stderr(&out));
    assert!(stdout(&out).is_empty(), "no result may be printed: {}", stdout(&out));
    let err = stderr(&out);
    assert!(err.contains("first 65535 of 70000 rows"), "stderr: {}", err);
    assert!(err.contains("--calc"), "the hint should say why --calc refused: {}", err);
}

/// Looking at part of an oversized file is fine, but not silently: the note
/// goes to stderr so JSON on stdout stays parseable.
#[test]
fn inspect_range_of_an_oversized_file_notes_the_truncation_on_stderr() {
    let out = vgrid(&["sheet", "inspect", &fixture("parquet_70000_rows.parquet"), "A1:A3", "--json"]);

    assert!(out.status.success(), "stderr: {}", stderr(&out));
    let parsed: serde_json::Value = serde_json::from_str(&stdout(&out)).expect("stdout stays valid JSON");
    assert_eq!(parsed["cells"][1]["value"], "1");
    assert!(stderr(&out).contains("first 65535 of 70000 rows"), "stderr: {}", stderr(&out));
}

#[test]
fn convert_still_refuses_a_file_bigger_than_a_sheet() {
    let out = vgrid(&["convert", &fixture("parquet_70000_rows.parquet"), "-t", "csv"]);
    assert_eq!(out.status.code(), Some(4), "stderr: {}", stderr(&out));
}

// ---------------------------------------------------------------------------
// Text that looks numeric stays text
// ---------------------------------------------------------------------------

/// JSON used to be typed by whether a cell's display text parsed as a number,
/// so the string "007" became the number 7.
#[test]
fn convert_json_keeps_numeric_looking_text_as_strings() {
    let out = vgrid(&["convert", &fixture("parquet_orders.parquet"), "-t", "json", "--headers"]);
    assert!(out.status.success(), "stderr: {}", stderr(&out));

    let rows: serde_json::Value = serde_json::from_str(&stdout(&out)).unwrap();
    assert_eq!(rows[0]["order_id"], serde_json::json!("007"));
    assert_eq!(rows[0]["amount"], serde_json::json!(1234.5), "numbers stay numbers");
    assert_eq!(rows[0]["status"], serde_json::json!("paid"));
}

#[test]
fn inspect_reports_numeric_looking_text_as_text() {
    let out = vgrid(&["sheet", "inspect", &fixture("parquet_orders.parquet"), "A2:D2", "--json"]);
    assert!(out.status.success(), "stderr: {}", stderr(&out));

    let parsed: serde_json::Value = serde_json::from_str(&stdout(&out)).unwrap();
    let cells = parsed["cells"].as_array().unwrap();
    assert_eq!(cells[0]["value"], "007");
    assert_eq!(cells[0]["value_type"], "text");
    assert_eq!(cells[3]["value_type"], "number");
}

/// CSV import already stores numbers and booleans as such, so typing JSON by
/// what the cell holds must not change CSV-to-JSON output.
#[test]
fn convert_csv_to_json_types_are_unchanged() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("t.csv");
    std::fs::write(&csv, "id,amount,flag,name\n42,3.5,TRUE,Ann\n").unwrap();

    let out = vgrid(&["convert", csv.to_str().unwrap(), "-t", "json", "--headers"]);
    assert!(out.status.success(), "stderr: {}", stderr(&out));

    let rows: serde_json::Value = serde_json::from_str(&stdout(&out)).unwrap();
    assert_eq!(rows[0]["id"], serde_json::json!(42));
    assert_eq!(rows[0]["amount"], serde_json::json!(3.5));
    assert_eq!(rows[0]["flag"], serde_json::json!(true));
    assert_eq!(rows[0]["name"], serde_json::json!("Ann"));
}

// ---------------------------------------------------------------------------
// Dates in CSV are ISO 8601, not serials
// ---------------------------------------------------------------------------

#[test]
fn convert_csv_writes_timestamps_and_dates_as_iso_8601() {
    let out = vgrid(&["convert", &fixture("parquet_orders.parquet"), "-t", "csv"]);
    assert!(out.status.success(), "stderr: {}", stderr(&out));

    let csv = stdout(&out);
    let lines: Vec<&str> = csv.lines().collect();
    assert_eq!(lines[0], "order_id,placed_at,ship_date,amount,status");
    assert_eq!(lines[1], "007,2026-09-01 14:02:00,2026-09-03,1234.50,paid");
    assert_eq!(lines[2], "008,2026-09-01 14:07:30.123,2026-09-04,89,pending");
}
