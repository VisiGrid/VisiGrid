//! Headless XLSX export must remain usable for sorted 0.44 workbooks.
use std::process::Command;
use visigrid_engine::{
    filter::SortDirection,
    table::TableRange,
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};
use visigrid_io::{native, xlsx};

fn sorted_book(lookup: bool) -> Workbook {
    let mut wb = Workbook::new();
    for (row, values) in [["Key", "Amount"], ["2", "20"], ["1", "10"]]
        .iter()
        .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    let sid = wb.active_sheet_id();
    let id = wb
        .create_table(
            sid,
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(sid, Some(view)).unwrap();
    let other = wb.add_sheet_named("Lookup").unwrap();
    if lookup {
        wb.set_cell_value_tracked(other, 0, 0, "=VLOOKUP(1,Sheet1!A2:B3,2,FALSE)");
    }
    assert_eq!(wb.saved_tables().version, 3, "0.44-compatible catalog");
    wb
}

#[test]
fn convert_sorted_native_with_vlookup_falls_back_for_file_and_binary_stdout_even_when_quiet() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    let wb = sorted_book(true);
    native::save_workbook(&wb, &input).unwrap();
    let original = std::fs::read(&input).unwrap();
    assert!(xlsx::export_to_buffer(&wb, None)
        .unwrap_err()
        .contains("VLOOKUP"));
    for file in [true, false] {
        let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
        cmd.arg("convert")
            .arg(&input)
            .args(["-t", "xlsx", "--quiet"]);
        if file {
            cmd.arg("-o").arg(&output);
        }
        let result = cmd.output().unwrap();
        let stderr = String::from_utf8_lossy(&result.stderr);
        assert!(result.status.success(), "{stderr}");
        assert!(
            stderr.contains("warning:") && stderr.contains("VLOOKUP"),
            "{stderr}"
        );
        assert!(
            stderr.contains("Exported in stored order instead") && stderr.contains("Reapply"),
            "{stderr}"
        );
        if file {
            assert!(result.stdout.is_empty());
        } else {
            assert!(
                result.stdout.starts_with(b"PK"),
                "stdout is only XLSX bytes"
            );
            std::fs::write(&output, result.stdout).unwrap();
        }
        let (loaded, _) = xlsx::import(&output).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "2");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 0), "1");
        assert!(loaded
            .sheet(0)
            .unwrap()
            .table_view_spec()
            .unwrap()
            .sort
            .is_some());
        assert_eq!(
            loaded.sheet(1).unwrap().get_raw(0, 0),
            wb.sheet(1).unwrap().get_raw(0, 0)
        );
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "10");
        assert_eq!(std::fs::read(&input).unwrap(), original);
    }
}

#[test]
fn convert_still_materializes_supported_sorting() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    native::save_workbook(&sorted_book(false), &input).unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "xlsx", "-o"])
        .arg(&output)
        .output()
        .unwrap();
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(result.status.success(), "{stderr}");
    assert!(!stderr.contains("stored order instead"), "{stderr}");
    let (loaded, _) = xlsx::import(&output).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "1");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 0), "2");
}

#[test]
fn convert_does_not_fallback_past_schema_refusal_or_touch_destination() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    let mut wb = sorted_book(true);
    let id = wb.tables().next().unwrap().1.id;
    wb.rename_table(id, &"T".repeat(256)).unwrap();
    native::save_workbook(&wb, &input).unwrap();
    std::fs::write(&output, b"keep destination").unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "xlsx", "-o"])
        .arg(&output)
        .output()
        .unwrap();
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("255-character"));
    assert!(result.stdout.is_empty());
    assert_eq!(std::fs::read(&output).unwrap(), b"keep destination");
}

const INPUT_SECRET: &str = "VISIGRID_XLSX_PASSWORD";
const OUTPUT_SECRET: &str = "VISIGRID_XLSX_OUTPUT_PASSWORD";

fn convert(path: &std::path::Path, format: &str) -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.arg("convert")
        .arg(path)
        .args(["-t", format, "--quiet"])
        .env_remove(INPUT_SECRET)
        .env_remove(OUTPUT_SECRET);
    cmd
}

#[test]
fn protected_source_keeps_changes_or_removes_its_password_only_as_requested() {
    use visigrid_io::xlsx_crypto::{decrypt_ooxml, CryptoError};
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../io/tests/fixtures/encrypted_msoffcrypto_agile.xlsx");
    let run = |args: &[&str], output_password: Option<&str>| {
        let mut cmd = convert(&path, "xlsx");
        cmd.env(INPUT_SECRET, "hunter2").args(args);
        if let Some(pw) = output_password {
            cmd.env(OUTPUT_SECRET, pw);
        }
        let out = cmd.output().unwrap();
        let stderr = String::from_utf8(out.stderr).unwrap();
        assert!(out.status.success(), "{stderr}");
        assert!(!stderr.contains("hunter2") && !stderr.contains("replacement-secret"));
        (out.stdout, stderr)
    };
    let (original, notice) = run(&[], None);
    assert!(decrypt_ooxml(&original, "hunter2")
        .unwrap()
        .starts_with(b"PK"));
    assert!(notice.contains("encrypted with the source workbook password"));

    // Taking just one sheet uses the other export path.
    let (changed, notice) = run(&["--sheet", "0"], Some("replacement-secret"));
    assert!(decrypt_ooxml(&changed, "replacement-secret")
        .unwrap()
        .starts_with(b"PK"));
    assert_eq!(
        decrypt_ooxml(&changed, "hunter2").unwrap_err(),
        CryptoError::IncorrectPassword
    );
    assert!(notice.contains("supplied output password"));

    let (plain, warning) = run(&["--no-encrypt"], Some("replacement-secret"));
    assert!(plain.starts_with(b"PK"));
    assert!(warning.contains("warning: writing an unencrypted XLSX copy"));
    let csv = convert(&path, "csv")
        .env(INPUT_SECRET, "hunter2")
        .output()
        .unwrap();
    assert!(csv.status.success() && !csv.stdout.is_empty());
    assert!(!String::from_utf8_lossy(&csv.stderr).contains("unencrypted XLSX"));
}

#[test]
fn plain_sources_encrypt_only_on_request() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("source.csv");
    std::fs::write(&csv, "Name,Value\nTest,123\n").unwrap();
    let plain = convert(&csv, "xlsx")
        .env(OUTPUT_SECRET, "not-requested")
        .output()
        .unwrap();
    assert!(plain.status.success() && plain.stdout.starts_with(b"PK"));

    let xlsx = dir.path().join("plain.xlsx");
    std::fs::write(&xlsx, plain.stdout).unwrap();
    let ordinary = convert(&xlsx, "xlsx")
        .env(INPUT_SECRET, "unused")
        .output()
        .unwrap();
    assert!(ordinary.status.success() && ordinary.stdout.starts_with(b"PK"));

    let encrypted = dir.path().join("encrypted.xlsx");
    let output = convert(&xlsx, "xlsx")
        .arg("--encrypt")
        .arg("-o")
        .arg(&encrypted)
        .env(OUTPUT_SECRET, "test-secret")
        .output()
        .unwrap();
    assert!(output.status.success() && output.stdout.is_empty());
    assert!(!String::from_utf8_lossy(&output.stderr).contains("test-secret"));
    let back = convert(&encrypted, "csv")
        .env(INPUT_SECRET, "test-secret")
        .output()
        .unwrap();
    assert!(back.status.success());
    assert_eq!(
        String::from_utf8(back.stdout).unwrap(),
        "Name,Value\nTest,123\n"
    );
}

#[test]
fn invalid_encryption_requests_fail_before_replacing_output() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("input.csv");
    let dest = dir.path().join("previous.xlsx");
    std::fs::write(&csv, "value\n1\n").unwrap();
    std::fs::write(&dest, "previous").unwrap();
    for (format, flags, pw, diagnostic) in [
        ("csv", vec!["--encrypt"], None, "--encrypt requires XLSX"),
        (
            "xlsx",
            vec!["--encrypt", "--no-encrypt"],
            None,
            "--no-encrypt",
        ),
        ("xlsx", vec!["--encrypt"], None, OUTPUT_SECRET),
        ("xlsx", vec!["--encrypt"], Some(""), "1 to 255 UTF-16"),
    ] {
        let mut cmd = convert(&csv, format);
        cmd.args(flags).arg("-o").arg(&dest);
        if let Some(password) = pw {
            cmd.env(OUTPUT_SECRET, password);
        }
        let result = cmd.output().unwrap();
        let err = String::from_utf8_lossy(&result.stderr);
        assert!(!result.status.success() && result.stdout.is_empty());
        assert!(err.contains(diagnostic), "{err}");
        assert_eq!(std::fs::read(&dest).unwrap(), b"previous");
    }
}
