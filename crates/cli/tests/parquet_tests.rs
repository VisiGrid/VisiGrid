// Integration tests for Parquet through `vgrid convert` and `vgrid sheet inspect`.
// Run with: cargo test -p visigrid-cli --test parquet_tests
//
// Fixtures (tests/fixtures/, written with pyarrow):
//   parquet_orders.parquet      order_id string ("007"…), placed_at timestamp[us]
//                               (one with .123 s), ship_date date32,
//                               amount decimal(10,2), status string; 3 rows
//   parquet_70000_rows.parquet  one int32 column `n`, every value 1; 70,000
//                               rows — more than a sheet held before 0.35.0,
//                               and an ordinary file since

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

#[test]
fn peek_parquet_json_preserves_ids_and_reports_preview_limit() {
    let out = vgrid(&["peek", &fixture("parquet_orders.parquet"), "--json", "--max-rows", "1"]);
    assert!(out.status.success(), "{}", stderr(&out));
    let data: serde_json::Value = serde_json::from_str(&stdout(&out)).unwrap();
    assert_eq!(data["columns"][0], "order_id");
    assert_eq!(data["rows"].as_array().unwrap().len(), 1);
    assert_eq!(data["rows"][0][0], "007");
    assert_eq!(data["rows"][0][1], "2026-09-01 14:02:00");
    assert_eq!(data["rows"][0][2], "2026-09-03");
    assert_eq!(data["rows"][0][3], 1234.5);
    assert!(stderr(&out).contains("1 of 3 rows"));
}

#[test]
fn peek_parquet_plain_and_shape_show_total_rows() {
    let path = fixture("parquet_70000_rows.parquet");
    let plain = vgrid(&["peek", &path, "--plain", "--max-rows", "2"]);
    assert!(plain.status.success(), "{}", stderr(&plain));
    assert!(stdout(&plain).contains("showing 2 of 70000 rows"));
    let shape = vgrid(&["peek", &path, "--shape"]);
    assert!(shape.status.success(), "{}", stderr(&shape));
    assert!(stdout(&shape).contains("rows:       70000"));
    assert!(stdout(&shape).contains("loaded:     5000"));
    assert!(stdout(&shape).contains("headers:    yes"));
}

#[test]
fn peek_parquet_all_rows_and_non_tty_fallback() {
    let out = vgrid(&["peek", &fixture("parquet_orders.parquet"), "--max-rows", "0"]);
    assert!(out.status.success(), "{}", stderr(&out));
    assert!(stdout(&out).contains("007"));
    assert!(stdout(&out).contains("2026-09-01 14:02:00"));
    assert!(stdout(&out).contains("2026-09-03"));
    assert!(!stdout(&out).contains("showing"));
    assert!(!stdout(&out).contains('\x1b'));
}

#[test]
fn peek_parquet_rejects_inapplicable_flags() {
    for extra in [vec!["--no-headers"], vec!["--delimiter", "tab"], vec!["--sheet", "0"], vec!["--recompute"]] {
        let path = fixture("parquet_orders.parquet");
        let mut args = vec!["peek", &path, "--json"];
        args.extend(extra);
        let out = vgrid(&args);
        assert!(!out.status.success());
        assert!(stderr(&out).contains("do not apply"), "{}", stderr(&out));
    }
}

// ---------------------------------------------------------------------------
// Files a sheet used to be too small for
//
// Until 0.35.0 a sheet held 65,536 rows, so this 70,000-row file was refused
// by `convert` and by `--calc`, and `inspect` showed a truncated prefix with a
// warning. The grid is now 1,048,576 rows and the file is unremarkable. The
// refusal itself still matters for files past the new limit; that path is
// covered in visigrid-io (parquet.rs builds a MAX_ROWS + 10 file synthetically,
// which is cheaper than carrying a million-row fixture here).
// ---------------------------------------------------------------------------

/// The whole point of the raise: an aggregate now covers every record instead
/// of refusing, and the answer is the true one — 70,000 values of 1.
#[test]
fn inspect_calc_totals_every_row_of_a_70000_row_file() {
    let out = vgrid(&["sheet", "inspect", &fixture("parquet_70000_rows.parquet"), "--headers", "--calc", "SUM(n)"]);

    assert_eq!(out.status.code(), Some(0), "stdout: {}\nstderr: {}", stdout(&out), stderr(&out));
    assert!(stdout(&out).contains("70000"), "SUM over every row: {}", stdout(&out));
}

/// Nothing is truncated, so nothing is warned about, and stdout stays JSON.
#[test]
fn inspect_range_of_a_70000_row_file_reports_no_truncation() {
    let out = vgrid(&["sheet", "inspect", &fixture("parquet_70000_rows.parquet"), "A1:A3", "--json"]);

    assert!(out.status.success(), "stderr: {}", stderr(&out));
    let parsed: serde_json::Value = serde_json::from_str(&stdout(&out)).expect("stdout stays valid JSON");
    assert_eq!(parsed["cells"][1]["value"], "1");
    assert!(!stderr(&out).contains("of 70000 rows"), "no truncation note: {}", stderr(&out));
}

#[test]
fn convert_takes_a_70000_row_file() {
    let out = vgrid(&["convert", &fixture("parquet_70000_rows.parquet"), "-t", "csv"]);
    assert!(out.status.success(), "stderr: {}", stderr(&out));
    // Header plus every record, none dropped.
    assert_eq!(stdout(&out).lines().count(), 70_001, "stderr: {}", stderr(&out));
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
