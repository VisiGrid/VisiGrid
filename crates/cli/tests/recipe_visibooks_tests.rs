// `vgrid recipe run` on a VisiBooks source, against a local stand-in for the
// VisiBooks API: the request it makes, the key binding, and the table it
// writes. Run with: cargo test -p visigrid-cli --test recipe_visibooks_tests

use std::io::{BufRead, BufReader, Write};
use std::net::TcpListener;
use std::process::Command;
use std::sync::mpsc;

/// Serve `body` to every request on a local port; send each request's
/// first line and Authorization header back on the channel.
fn stand_in(body: &'static str) -> (String, mpsc::Receiver<(String, String)>) {
    let listener = TcpListener::bind("127.0.0.1:0").unwrap();
    let origin = format!("http://127.0.0.1:{}", listener.local_addr().unwrap().port());
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        for stream in listener.incoming() {
            let Ok(mut stream) = stream else { continue };
            let mut reader = BufReader::new(stream.try_clone().unwrap());
            let mut first = String::new();
            reader.read_line(&mut first).unwrap();
            let mut auth = String::new();
            loop {
                let mut line = String::new();
                if reader.read_line(&mut line).unwrap() == 0 || line == "\r\n" {
                    break;
                }
                if line.to_lowercase().starts_with("authorization:") {
                    auth = line[14..].trim().to_string();
                }
            }
            let _ = tx.send((first.trim().to_string(), auth));
            let _ = write!(stream, "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{body}", body.len());
        }
    });
    (origin, rx)
}

const TRIAL_BALANCE: &str = r#"{"basis":"cash","as_of":"2026-02-28","lines":[
  {"account_code":"1010","account_name":"Bank","account_type":"asset","debit_minor":123456,"credit_minor":0},
  {"account_code":"4010","account_name":"Rental Income","account_type":"revenue","debit_minor":0,"credit_minor":123456}]}"#;

fn recipe(dir: &std::path::Path, origin: &str) -> std::path::PathBuf {
    let path = dir.join("tb.recipe.toml");
    std::fs::write(
        &path,
        format!(
            "version = 1\n[source]\nkind = \"visibooks\"\nserver = \"{origin}\"\nentity = \"42\"\nreport = \"trial_balance\"\nas_of = \"2026-02-28\"\nbasis = \"cash\"\n[[step]]\nop = \"filter\"\ncolumn = \"Type\"\nis = \"=\"\nvalue = \"revenue\"\n"
        ),
    )
    .unwrap();
    path
}

fn vgrid(dir: &std::path::Path, args: &[&str], env: &[(&str, &str)]) -> std::process::Output {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
    cmd.args(args).current_dir(dir).env_remove("VISIBOOKS_API_KEY").env_remove("VISIBOOKS_API_SERVER");
    // No real keychain: a throwaway home and config
    cmd.env("HOME", dir).env("XDG_CONFIG_HOME", dir.join("xdg")).env("DBUS_SESSION_BUS_ADDRESS", "unix:path=/nonexistent");
    for (k, v) in env {
        cmd.env(k, v);
    }
    cmd.output().unwrap()
}

#[test]
fn a_visibooks_recipe_reads_the_report_with_the_bound_key() {
    let dir = tempfile::tempdir().unwrap();
    let (origin, requests) = stand_in(TRIAL_BALANCE);
    let path = recipe(dir.path(), &origin);
    let out = dir.path().join("tb.csv");
    let o = vgrid(
        dir.path(),
        &["recipe", "run", path.to_str().unwrap(), "-o", out.to_str().unwrap()],
        &[("VISIBOOKS_API_KEY", "vb_test_key"), ("VISIBOOKS_API_SERVER", &origin)],
    );
    assert!(o.status.success(), "stderr: {}", String::from_utf8_lossy(&o.stderr));
    let (line, auth) = requests.recv_timeout(std::time::Duration::from_secs(5)).unwrap();
    assert_eq!(line, "GET /api/v1/books/agent/trial_balance?entity_id=42&as_of=2026-02-28&basis=cash HTTP/1.1");
    assert_eq!(auth, "Bearer vb_test_key");
    // Exact amounts, account codes as text, the filter step applied
    assert_eq!(
        std::fs::read_to_string(&out).unwrap().lines().collect::<Vec<_>>(),
        ["Account,Name,Type,Debit,Credit,Balance", "4010,Rental Income,revenue,0,1234.56,-1234.56"]
    );
}

#[test]
fn a_key_for_another_server_is_never_sent() {
    let dir = tempfile::tempdir().unwrap();
    let (origin, requests) = stand_in(TRIAL_BALANCE);
    let path = recipe(dir.path(), &origin);
    // The environment key belongs to VisiAPI, not to the server this recipe names
    let o = vgrid(dir.path(), &["recipe", "run", path.to_str().unwrap()], &[("VISIBOOKS_API_KEY", "vb_test_key")]);
    assert!(!o.status.success());
    let stderr = String::from_utf8_lossy(&o.stderr);
    assert!(stderr.contains("no VisiBooks API key for"), "{stderr}");
    assert!(requests.recv_timeout(std::time::Duration::from_millis(300)).is_err(), "nothing was requested");
}

#[test]
fn a_visibooks_source_cant_be_swapped_for_a_file() {
    let dir = tempfile::tempdir().unwrap();
    let path = recipe(dir.path(), "https://api.visiapi.com");
    std::fs::write(dir.path().join("x.csv"), "a\n1\n").unwrap();
    let o = vgrid(dir.path(), &["recipe", "run", path.to_str().unwrap(), "--source", "x.csv"], &[]);
    assert!(!o.status.success());
    assert!(String::from_utf8_lossy(&o.stderr).contains("can't be replaced by a file"));
}
