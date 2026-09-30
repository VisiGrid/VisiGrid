// Integration tests for `vgrid pivot` on files.
// Run with: cargo test -p visigrid-cli --test pivot_tests
//
// The --session path is covered by the session-host handler tests; a manual
// smoke test against a headless host:
//   vgrid serve data.sheet            # note the session id
//   vgrid pivot --session <id> --rows Region --values sum:Amount
//   vgrid pivot --session <id> --refresh

use std::path::Path;
use std::process::Command;

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

fn sales_csv() -> tempfile::NamedTempFile {
    let f = tempfile::Builder::new().suffix(".csv").tempfile().unwrap();
    std::fs::write(
        f.path(),
        "Region,Rep,Month,Amount,Order\n\
         West,Ann,Jan,1200.50,A1\n\
         East,Bo,Jan,800,A2\n\
         West,Cy,Feb,300,A3\n\
         East,Bo,Feb,4500.25,A4\n",
    )
    .unwrap();
    f
}

fn run(args: &[&str]) -> (bool, String, String) {
    let out = vgrid().args(args).output().expect("run vgrid");
    (
        out.status.success(),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

#[test]
fn csv_output_is_unformatted_with_grand_total() {
    let f = sales_csv();
    let path = f.path().to_str().unwrap();
    let (ok, stdout, stderr) = run(&["pivot", path, "--rows", "region", "--values", "sum:Amount,count:Order", "--csv"]);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        "Region,Sum of Amount,Count of Order\nEast,5300.25,2\nWest,1500.5,2\nGrand Total,6800.75,4\n"
    );
}

#[test]
fn column_field_table_and_json() {
    let f = sales_csv();
    let path = f.path().to_str().unwrap();
    let (ok, stdout, _) = run(&["pivot", path, "-r", "Region", "-c", "Month", "-v", "Amount"]);
    assert!(ok);
    assert!(stdout.contains("Feb"), "{stdout}");
    assert!(stdout.contains("5,300.25"), "table output is formatted: {stdout}");

    let (ok, stdout, _) = run(&["pivot", path, "-r", "Region", "-v", "distinct:Rep", "--json"]);
    assert!(ok);
    let doc: serde_json::Value = serde_json::from_str(&stdout).unwrap();
    assert_eq!(doc["headers"][0][1], "Distinct Count of Rep");
    assert_eq!(doc["rows"][0], serde_json::json!(["East", 1]));
    assert_eq!(doc["rows"][2], serde_json::json!(["Grand Total", 3]));
    assert_eq!(doc["source_rows"], 4);
}

#[test]
fn parquet_fixture() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/parquet_orders.parquet");
    let (ok, stdout, stderr) = run(&["pivot", path.to_str().unwrap(), "-r", "status", "-v", "sum:amount", "--csv"]);
    assert!(ok, "{stderr}");
    assert!(stdout.starts_with("status,Sum of amount\n"), "{stdout}");
    assert!(stdout.contains("Grand Total,"));
}

#[test]
fn refusals_name_the_problem() {
    let f = sales_csv();
    let path = f.path().to_str().unwrap();
    let (ok, _, stderr) = run(&["pivot", path, "-r", "Mnth", "-v", "Amount"]);
    assert!(!ok);
    assert!(stderr.contains("no column headed \"Mnth\"") && stderr.contains("Region, Rep, Month"), "{stderr}");

    let (ok, _, stderr) = run(&["pivot", path, "-r", "Region", "-v", "median:Amount"]);
    assert!(!ok);
    assert!(stderr.contains("unknown aggregation \"median\""), "{stderr}");

    let (ok, _, stderr) = run(&["pivot", "--rows", "Region"]);
    assert!(!ok);
    assert!(stderr.contains("give a file"), "{stderr}");
}
