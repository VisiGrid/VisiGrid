// Integration tests for `vgrid peek` across formats and safety caps.
// Run with: cargo test -p visigrid-cli --test peek_tests -- --nocapture
//
// Manual smoke test (cannot be automated — requires a real TTY):
//   vgrid peek tests/fixtures/inspect_small.xlsx --headers
//   Verify: TUI launches, tab switching works, q exits cleanly, terminal state restored.

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
// .sheet peek --shape works
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_shape() {
    let output = vgrid()
        .args(["peek", template_path().to_str().unwrap(), "--shape"])
        .output()
        .expect("vgrid peek --shape");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("format:"), "should print format line");
    assert!(stdout.contains("sheets:"), "should print sheets count");
}

// ---------------------------------------------------------------------------
// .sheet peek --plain works
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_plain() {
    let output = vgrid()
        .args(["peek", template_path().to_str().unwrap(), "--plain"])
        .output()
        .expect("vgrid peek --plain");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    // Should have at least some output
    assert!(!stdout.is_empty(), "plain output should not be empty");
}

// ---------------------------------------------------------------------------
// .sheet peek --max-rows truncates
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_max_rows_truncates() {
    let output = vgrid()
        .args(["peek", template_path().to_str().unwrap(), "--max-rows", "2", "--shape"])
        .output()
        .expect("vgrid peek --max-rows 2 --shape");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));
}

// ---------------------------------------------------------------------------
// .sheet peek with --max-rows 0 on a small file succeeds (under cap)
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_max_rows_zero_small_file() {
    let output = vgrid()
        .args(["peek", template_path().to_str().unwrap(), "--max-rows", "0", "--shape"])
        .output()
        .expect("vgrid peek --max-rows 0 --shape");

    // Small template should be well under 200k rows, so this should succeed
    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));
}

// ---------------------------------------------------------------------------
// .sheet peek --max-rows 0 on a large workbook errors without --force
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_safety_cap_rejects_huge_workbook() {
    // Create a .sheet workbook with >200k rows
    use visigrid_engine::workbook::Workbook;

    let mut wb = Workbook::new();
    let sheet = wb.sheet_mut(0).unwrap();

    // We need > 200_000 rows. To keep the test fast, just set a single cell
    // far beyond row 200_000 — the bbox will extend to that row.
    sheet.set_value(200_001, 0, "sentinel");

    let dir = tempfile::tempdir().expect("create temp dir");
    let path = dir.path().join("huge.sheet");
    visigrid_io::native::save_workbook(&wb, &path).expect("save workbook");

    // Without --force, --max-rows 0 should error
    let output = vgrid()
        .args(["peek", path.to_str().unwrap(), "--max-rows", "0", "--plain"])
        .output()
        .expect("vgrid peek huge.sheet --max-rows 0");

    assert!(!output.status.success(), "should fail without --force");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains(">200k rows") || stderr.contains("200k"),
        "error should mention row cap, got: {}", stderr);
}

// ---------------------------------------------------------------------------
// .sheet peek --max-rows 0 --force succeeds on large workbook
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_safety_cap_force_overrides() {
    use visigrid_engine::workbook::Workbook;

    let mut wb = Workbook::new();
    let sheet = wb.sheet_mut(0).unwrap();
    // Sparse sheet: only 2 cells but bbox extends to row 200_001
    sheet.set_value(0, 0, "start");
    sheet.set_value(200_001, 0, "end");

    let dir = tempfile::tempdir().expect("create temp dir");
    let path = dir.path().join("huge.sheet");
    visigrid_io::native::save_workbook(&wb, &path).expect("save workbook");

    // With --force, should succeed
    let output = vgrid()
        .args(["peek", path.to_str().unwrap(), "--max-rows", "0", "--force", "--shape"])
        .output()
        .expect("vgrid peek huge.sheet --max-rows 0 --force --shape");

    assert!(output.status.success(), "should succeed with --force\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));
}

// ---------------------------------------------------------------------------
// .sheet peek with explicit --max-rows avoids cap
// ---------------------------------------------------------------------------

#[test]
fn peek_sheet_explicit_max_rows_avoids_cap() {
    use visigrid_engine::workbook::Workbook;

    let mut wb = Workbook::new();
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set_value(0, 0, "start");
    sheet.set_value(200_001, 0, "end");

    let dir = tempfile::tempdir().expect("create temp dir");
    let path = dir.path().join("huge.sheet");
    visigrid_io::native::save_workbook(&wb, &path).expect("save workbook");

    // Explicit --max-rows 100 should work (no cap check for non-zero)
    let output = vgrid()
        .args(["peek", path.to_str().unwrap(), "--max-rows", "100", "--shape"])
        .output()
        .expect("vgrid peek huge.sheet --max-rows 100 --shape");

    assert!(output.status.success(), "should succeed with explicit --max-rows\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));
}

// ---------------------------------------------------------------------------
// Non-TTY auto-fallback: peek exits 0 with table output when stdout is piped
// ---------------------------------------------------------------------------

#[test]
fn peek_csv_non_tty_auto_fallback() {
    let csv = std::env::temp_dir().join("vgrid_peek_tty_test.csv");
    std::fs::write(&csv, "Name,Age\nAlice,30\nBob,25\n").unwrap();

    // Command::output() captures stdout → not a TTY → should auto-fallback
    let output = vgrid()
        .args(["peek", csv.to_str().unwrap()])
        .output()
        .expect("vgrid peek csv (non-TTY)");

    assert!(output.status.success(), "peek should exit 0 in non-TTY mode\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.is_empty(), "should produce table output");
    // Should contain data rows, not raw-mode error
    assert!(!stdout.contains("raw mode"), "should not mention raw mode errors");
    assert!(stdout.contains("Alice") || stdout.contains("Name"),
        "output should contain data from CSV");

    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_sheet_non_tty_auto_fallback() {
    // .sheet file with stdout captured → should auto-fallback to plain
    let output = vgrid()
        .args(["peek", template_path().to_str().unwrap()])
        .output()
        .expect("vgrid peek sheet (non-TTY)");

    assert!(output.status.success(), "peek should exit 0 in non-TTY mode\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.is_empty(), "should produce table output");
    assert!(!stdout.contains("raw mode"), "should not mention raw mode errors");
}

#[test]
fn peek_no_tui_flag() {
    let csv = std::env::temp_dir().join("vgrid_peek_no_tui.csv");
    std::fs::write(&csv, "X,Y\n1,2\n3,4\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--no-tui"])
        .output()
        .expect("vgrid peek --no-tui");

    assert!(output.status.success(), "exit code: {:?}\nstderr: {}",
        output.status, String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.is_empty(), "should produce output");
    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_tui_flag_errors_when_not_tty() {
    let csv = std::env::temp_dir().join("vgrid_peek_force_tui.csv");
    std::fs::write(&csv, "X,Y\n1,2\n").unwrap();

    // stdout captured → not a TTY → --tui should error
    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--tui"])
        .output()
        .expect("vgrid peek --tui (non-TTY)");

    assert!(!output.status.success(), "--tui should fail when not a TTY");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("interactive terminal") || stderr.contains("TTY"),
        "error should mention TTY requirement, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_tui_conflicts_with_plain() {
    let csv = std::env::temp_dir().join("vgrid_peek_tui_plain.csv");
    std::fs::write(&csv, "X,Y\n1,2\n").unwrap();

    let output = vgrid()
        .args(["peek", csv.to_str().unwrap(), "--tui", "--plain"])
        .output()
        .expect("vgrid peek --tui --plain");

    assert!(!output.status.success(), "--tui --plain should conflict");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("--tui") || stderr.contains("cannot be used with"),
        "error should mention conflict, got: {}", stderr);
    std::fs::remove_file(&csv).ok();
}

#[test]
fn peek_xlsx_non_tty_auto_fallback() {
    let xlsx = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/inspect_small.xlsx");

    let output = vgrid()
        .args(["peek", xlsx.to_str().unwrap()])
        .output()
        .expect("vgrid peek xlsx (non-TTY)");

    assert!(output.status.success(), "peek xlsx should exit 0 in non-TTY\nstderr: {}",
        String::from_utf8_lossy(&output.stderr));

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.is_empty(), "should produce table output");
    assert!(stdout.contains("Amount") || stdout.contains("100"),
        "output should contain data from XLSX");
}

// ---------------------------------------------------------------------------
// Unsupported extension gives helpful error
// ---------------------------------------------------------------------------

#[test]
fn peek_unsupported_extension_error() {
    let dir = tempfile::tempdir().expect("create temp dir");
    let path = dir.path().join("data.parquet");
    std::fs::write(&path, b"dummy").expect("write dummy file");

    let output = vgrid()
        .args(["peek", path.to_str().unwrap()])
        .output()
        .expect("vgrid peek data.parquet");

    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("xlsx"), "error should mention xlsx as supported, got: {}", stderr);
    assert!(stderr.contains("ods"), "error should mention ods as supported, got: {}", stderr);
}
