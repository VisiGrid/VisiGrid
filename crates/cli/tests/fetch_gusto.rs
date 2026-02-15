// Integration tests for `vgrid fetch gusto`.
// Run with: cargo test -p visigrid-cli --test fetch_gusto

use std::process::Command;

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    cmd
}

#[test]
fn missing_credentials_exits_50() {
    let output = vgrid()
        .args(["fetch", "gusto", "--from", "2026-01-01", "--to", "2026-01-31", "--quiet"])
        .output()
        .expect("failed to run vgrid");

    assert_eq!(
        output.status.code(),
        Some(50),
        "expected exit 50, got {:?}\nstderr: {}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr),
    );
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("missing Gusto credentials"),
        "stderr: {}",
        stderr,
    );
}

#[test]
fn invalid_date_range_exits_2() {
    let output = vgrid()
        .args([
            "fetch", "gusto",
            "--from", "2026-01-31",
            "--to", "2026-01-01",
            "--access-token", "gp_test_fake",
            "--company-uuid", "abc-123",
            "--quiet",
        ])
        .output()
        .expect("failed to run vgrid");

    assert_eq!(
        output.status.code(),
        Some(2),
        "expected exit 2, got {:?}\nstderr: {}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr),
    );
}

#[test]
fn bad_date_format_exits_2() {
    let output = vgrid()
        .args([
            "fetch", "gusto",
            "--from", "not-a-date",
            "--to", "2026-01-31",
            "--access-token", "gp_test_fake",
            "--company-uuid", "abc-123",
            "--quiet",
        ])
        .output()
        .expect("failed to run vgrid");

    assert_eq!(
        output.status.code(),
        Some(2),
        "expected exit 2, got {:?}\nstderr: {}",
        output.status.code(),
        String::from_utf8_lossy(&output.stderr),
    );
}
