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

fn vgrid_stdin(args: &[&str], input: &[u8]) -> Output {
    use std::io::Write;
    let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .current_dir(env!("CARGO_MANIFEST_DIR"))
        .args(args)
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .expect("run vgrid");
    child.stdin.take().unwrap().write_all(input).unwrap();
    child.wait_with_output().unwrap()
}

/// Piped CSV honours --encoding, as a file does. It used to be read as a
/// UTF-8 string, so a Windows-1252 file failed outright.
#[test]
fn piped_csv_honours_encoding() {
    let latin1 = b"name;qty\nCaf\xe9;007\n";
    let output = vgrid_stdin(&["convert", "-f", "csv", "--delimiter", ";", "--encoding", "windows-1252", "-t", "csv"], latin1);
    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    // --delimiter also sets the output delimiter
    assert_eq!(String::from_utf8_lossy(&output.stdout), "name;qty\nCafé;007\n");
}

/// The command the desktop import dialog prints for a chosen delimiter:
/// piped input with the column flags, written to a .sheet. Reading the
/// .sheet back shows the flags were applied.
#[test]
fn dialog_command_line_shape_imports_as_the_dialog_does() {
    let sheet = PathBuf::from(env!("CARGO_TARGET_TMPDIR")).join("dialog_shape.sheet");
    let _ = std::fs::remove_file(&sheet);
    let input = b"zip;sku;amount\n00501;007;1,5\n";
    let output = vgrid_stdin(
        &["convert", "-f", "csv", "--delimiter", ";", "--number", "sku", "--decimal-comma", "-t", "sheet", "-o", sheet.to_str().unwrap()],
        input,
    );
    assert!(output.status.success(), "stderr: {}", String::from_utf8_lossy(&output.stderr));
    let back = vgrid(&["convert", sheet.to_str().unwrap(), "-t", "csv"]);
    assert!(back.status.success(), "stderr: {}", String::from_utf8_lossy(&back.stderr));
    assert_eq!(String::from_utf8_lossy(&back.stdout), "zip,sku,amount\n00501,7,1.5\n");
}
