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
    // Unformatted numbers are written in full, not padded to the on-screen two
    // decimals: 1234.5, and 1.234 stays 1.234 rather than becoming 1.23 (#65).
    assert_eq!(lines[1], "007,2026-09-01 14:02:00,2026-09-03,1234.5,paid");
    assert_eq!(lines[2], "008,2026-09-01 14:07:30.123,2026-09-04,89,pending");
}

fn export_input(dir: &std::path::Path) -> std::path::PathBuf {
    let input = dir.join("typed.json");
    std::fs::write(&input, serde_json::json!({
        "format": "visigrid-json", "version": 2, "sheets": [{ "name": "Data", "cells": [
            {"row":0,"col":0,"value":"ID"}, {"row":0,"col":1,"value":"Amount"},
            {"row":1,"col":0,"value":"007"}, {"row":1,"col":1,"value":1.123456789012345},
            {"row":2,"col":0,"value":"008"}, {"row":2,"col":1,"value":2}
        ]}]
    }).to_string()).unwrap();
    input
}

#[test]
fn export_parquet_keeps_typed_values_and_filters_without_display_rounding() {
    use visigrid_engine::formula::eval::Value;
    let dir = tempfile::tempdir().unwrap();
    let input = export_input(dir.path());
    let output = dir.path().join("out.parquet");
    let out = vgrid(&["convert", input.to_str().unwrap(), "-f", "json-full", "-t", "parquet", "--headers", "--select", "Amount,ID", "--where", "Amount<2", "-o", output.to_str().unwrap()]);
    assert!(out.status.success(), "{}", stderr(&out));
    let imported = visigrid_io::parquet::import(&output).unwrap();
    assert_eq!(imported.total_rows, 1);
    assert_eq!(imported.sheet.get_computed_value(1,0), Value::Number(1.123456789012345));
    assert_eq!(imported.sheet.get_computed_value(1,1), Value::Text("007".into()));
}

#[test]
fn export_parquet_plan_is_json_and_stdout_is_binary_only() {
    let dir = tempfile::tempdir().unwrap();
    let input = export_input(dir.path());
    let args = ["convert", input.to_str().unwrap(), "-f", "json-full", "-t", "parquet", "--headers"];
    let mut plan_args = args.to_vec(); plan_args.push("--parquet-plan");
    let out = vgrid(&plan_args);
    assert!(out.status.success(), "{}", stderr(&out));
    let plan: serde_json::Value = serde_json::from_slice(&out.stdout).unwrap();
    assert_eq!(plan["ready"], true);
    assert_eq!(plan["row_count"], 2);
    assert_eq!(plan["columns"][0]["data_type"], "string");
    assert_eq!(plan["columns"][1]["data_type"], "double");
    let out = vgrid(&args);
    assert!(out.status.success(), "{}", stderr(&out));
    assert!(out.stdout.starts_with(b"PAR1"));
    assert!(out.stdout.ends_with(b"PAR1"));
}

#[test]
fn export_mixed_column_requires_explicit_text_option_and_preserves_existing_file() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("mixed.csv");
    std::fs::write(&input, "Value\n1.123456789\nN/A\n").unwrap();
    let output = dir.path().join("out.parquet");
    std::fs::write(&output, "previous contents").unwrap();
    let args = ["convert", input.to_str().unwrap(), "-t", "parquet", "--headers", "-o", output.to_str().unwrap()];
    let out = vgrid(&args);
    assert!(!out.status.success());
    assert!(stderr(&out).contains("A3"), "{}", stderr(&out));
    assert_eq!(std::fs::read_to_string(&output).unwrap(), "previous contents");
    let mut text_args = args.to_vec(); text_args.extend(["--text-column", "Value"]);
    let out = vgrid(&text_args);
    assert!(out.status.success(), "{}", stderr(&out));
    let imported = visigrid_io::parquet::import(&output).unwrap();
    assert_eq!(imported.sheet.get_raw(1,0), "1.123456789");
}

#[test]
fn export_rejects_duplicate_headers_and_parquet_options_on_other_formats() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("headers.csv");
    std::fs::write(&input, "ID,id\n1,2\n").unwrap();
    let out = vgrid(&["convert", input.to_str().unwrap(), "-t", "parquet", "--headers"]);
    assert!(!out.status.success());
    assert!(stderr(&out).contains("duplicate header"));
    let out = vgrid(&["convert", input.to_str().unwrap(), "-t", "csv", "--parquet-plan"]);
    assert!(!out.status.success());
    assert!(stderr(&out).contains("require -t parquet"));
}

#[test]
fn export_json_full_stdin_selects_the_requested_worksheet() {
    use std::io::Write;
    use std::process::Stdio;
    let doc = serde_json::json!({"format":"visigrid-json", "version":2, "sheets":[
        {"name":"First", "cells":[{"row":0,"col":0,"value":1}]},
        {"name":"Second", "cells":[{"row":0,"col":0,"value":"007"}]}
    ]});
    let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .args(["convert", "-f", "json-full", "-t", "parquet", "--sheet", "Second", "--parquet-plan"])
        .stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::piped()).spawn().unwrap();
    child.stdin.take().unwrap().write_all(doc.to_string().as_bytes()).unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(out.status.success(), "{}", stderr(&out));
    let plan: serde_json::Value = serde_json::from_slice(&out.stdout).unwrap();
    assert_eq!(plan["row_count"], 1);
    assert_eq!(plan["columns"][0]["data_type"], "string");
}
