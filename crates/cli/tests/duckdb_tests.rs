use std::process::{Command, Output};

fn run(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .args(args)
        .output()
        .unwrap()
}
fn success(output: Output) -> String {
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8(output.stdout).unwrap()
}

#[test]
fn convert_preview_and_import_database() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.csv");
    std::fs::write(&input, "Amount,Status\n10,paid\n20,pending\n30,paid\n").unwrap();
    let db = dir.path().join("new.duckdb");
    let input = input.to_str().unwrap();
    let path = db.to_str().unwrap();
    success(run(&[
        "convert",
        input,
        "-t",
        "duckdb",
        "--headers",
        "-o",
        path,
    ]));
    let bytes = std::fs::read(&db).unwrap();
    let peek = run(&["peek", path, "--sheet", "data", "--json", "--max-rows", "1"]);
    assert!(String::from_utf8_lossy(&peek.stderr).contains("1 of 3 rows"));
    let json: serde_json::Value = serde_json::from_str(&success(peek)).unwrap();
    assert_eq!(json["columns"], serde_json::json!(["Amount", "Status"]));
    assert_eq!(json["rows"], serde_json::json!([[10, "paid"]]));
    let shape = success(run(&["peek", path, "--shape", "--max-rows", "1"]));
    assert!(shape.contains("3 rows x 2 cols"), "{shape}");
    assert!(shape.contains("loaded: 1 rows"), "{shape}");
    let csv = success(run(&[
        "convert",
        path,
        "--sheet",
        "main.data",
        "-t",
        "csv",
        "--headers",
        "--where",
        "Status=paid",
    ]));
    assert_eq!(csv.lines().count(), 3);
    assert!(csv.contains("30,paid"));
    let sheet = dir.path().join("workbook.sheet");
    success(run(&[
        "convert",
        path,
        "-t",
        "sheet",
        "-o",
        sheet.to_str().unwrap(),
    ]));
    let inspect = success(run(&["sheet", "inspect", path, "--sheets", "--json"]));
    assert!(inspect.contains("main.data"));
    assert_eq!(std::fs::read(&db).unwrap(), bytes);
    let overwrite = run(&["convert", input, "-t", "duckdb", "--headers", "-o", path]);
    assert!(!overwrite.status.success());
    assert!(String::from_utf8_lossy(&overwrite.stderr).contains("already exists"));
    assert_eq!(std::fs::read(&db).unwrap(), bytes);
}

#[test]
fn mixed_columns_require_explicit_text_policy() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("mixed.csv");
    std::fs::write(&input, "Amount\n10\nN/A\n").unwrap();
    let db = dir.path().join("mixed.duckdb");
    let input = input.to_str().unwrap();
    let path = db.to_str().unwrap();
    let plan = success(run(&[
        "convert",
        input,
        "-t",
        "duckdb",
        "--headers",
        "--export-plan",
    ]));
    assert_eq!(
        serde_json::from_str::<serde_json::Value>(&plan).unwrap()["ready"],
        false
    );
    assert!(
        !run(&["convert", input, "-t", "duckdb", "--headers", "-o", path])
            .status
            .success()
    );
    assert!(!db.exists());
    success(run(&[
        "convert",
        input,
        "-t",
        "duckdb",
        "--headers",
        "--text-column",
        "Amount",
        "-o",
        path,
    ]));
    let json: serde_json::Value =
        serde_json::from_str(&success(run(&["peek", path, "--json"]))).unwrap();
    assert_eq!(json["rows"], serde_json::json!([["10"], ["N/A"]]));
    for flag in ["--no-headers", "--recompute"] {
        let out = run(&["peek", path, flag]);
        assert!(!out.status.success());
        assert!(String::from_utf8_lossy(&out.stderr).contains("do not apply"));
    }
    let unknown = run(&["convert", path, "--sheet", "missing", "-t", "csv"]);
    assert!(!unknown.status.success());
    assert!(String::from_utf8_lossy(&unknown.stderr).contains("available: main.data"));
}
