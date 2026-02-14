// Integration tests for `vgrid fetch stripe`.
// Run with: cargo test -p visigrid-cli --test fetch_stripe

use std::process::Command;

fn vgrid() -> Command {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.current_dir(env!("CARGO_MANIFEST_DIR"));
    // Clear env to avoid leaking a real key into tests
    cmd.env_remove("STRIPE_API_KEY");
    cmd
}

#[test]
fn missing_api_key_exits_50() {
    let output = vgrid()
        .args(["fetch", "stripe", "--from", "2026-01-01", "--to", "2026-01-31", "--quiet"])
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
        stderr.contains("missing Stripe API key"),
        "stderr: {}",
        stderr,
    );
}

#[test]
fn invalid_date_range_exits_2() {
    let output = vgrid()
        .args([
            "fetch", "stripe",
            "--from", "2026-01-31",
            "--to", "2026-01-01",
            "--api-key", "sk_test_fake",
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
            "fetch", "stripe",
            "--from", "not-a-date",
            "--to", "2026-01-31",
            "--api-key", "sk_test_fake",
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
