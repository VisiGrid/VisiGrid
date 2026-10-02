// Integration tests for CSV through `vgrid convert`.
// Run with: cargo test -p visigrid-cli --test csv_convert_tests

use std::path::PathBuf;
use std::process::{Command, Output};

fn vgrid(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .current_dir(env!("CARGO_MANIFEST_DIR"))
        .args(args)
        .output()
        .expect("run vgrid")
}

fn temp_csv(name: &str, contents: &str) -> PathBuf {
    let path = PathBuf::from(env!("CARGO_TARGET_TMPDIR")).join(name);
    std::fs::write(&path, contents).expect("write csv");
    path
}

/// `--delimiter` sets the output delimiter; a file's own delimiter is still
/// sniffed. Reading the comma file with `;` would put each row in one cell.
#[test]
fn convert_delimiter_changes_the_output_not_how_a_file_is_read() {
    let input = temp_csv("comma_to_semicolon.csv", "name,qty\nwidgets,12\ngadgets,7\n");
    let output = vgrid(&["convert", input.to_str().unwrap(), "-t", "csv", "--delimiter", ";"]);
    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    assert_eq!(String::from_utf8_lossy(&output.stdout), "name;qty\nwidgets;12\ngadgets;7\n");
}

/// Piped input has nothing to sniff from a file name, so `--delimiter`
/// still says how to read it, as before.
#[test]
fn convert_delimiter_still_reads_piped_input() {
    use std::io::Write;
    let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .args(["convert", "-f", "csv", "-t", "json", "--delimiter", ";"])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .expect("run vgrid");
    child.stdin.take().unwrap().write_all(b"name;qty\nwidgets;12\n").unwrap();
    let output = child.wait_with_output().unwrap();
    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let json = String::from_utf8_lossy(&output.stdout);
    assert!(json.contains("widgets") && json.contains("12"), "got {json}");
    assert!(!json.contains("widgets;12"), "piped row was not split on ';': {json}");
}
