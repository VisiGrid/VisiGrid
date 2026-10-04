//! The MCP recipe tools, through a real `vgrid mcp` process: an agent may
//! only run a recipe whose source the user approved in VisiGrid.
#![cfg(unix)]

use std::io::{BufRead, BufReader, Write};
use std::path::Path;
use std::process::{Command, Stdio};

/// Run one MCP exchange (initialize, then each call) and return the replies.
fn mcp(config_home: &Path, calls: &[serde_json::Value]) -> Vec<serde_json::Value> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("mcp")
        // The approval store lives in the user's config folder: give this
        // server its own (XDG on Linux, HOME on macOS)
        .env("XDG_CONFIG_HOME", config_home.join("xdg"))
        .env("HOME", config_home)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut out = BufReader::new(child.stdout.take().unwrap());
    let mut send = |v: serde_json::Value| writeln!(stdin, "{v}").unwrap();
    send(serde_json::json!({"jsonrpc":"2.0","id":0,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"t","version":"0"}}}));
    let mut line = String::new();
    out.read_line(&mut line).unwrap();
    let mut replies = Vec::new();
    for (i, call) in calls.iter().enumerate() {
        send(serde_json::json!({"jsonrpc":"2.0","id":i + 1,"method":"tools/call","params":call}));
        line.clear();
        out.read_line(&mut line).unwrap();
        replies.push(serde_json::from_str(&line).unwrap());
    }
    drop(stdin);
    let _ = child.wait();
    replies
}

fn text(reply: &serde_json::Value) -> String {
    reply["result"]["content"][0]["text"].as_str().unwrap_or_default().to_string()
}

fn approve(config_home: &Path, recipe_path: &Path) {
    let recipe = visigrid_io::recipe::Recipe::load(recipe_path).unwrap();
    let key = visigrid_io::recipe_trust::approval_key(recipe_path, &recipe);
    for dir in [config_home.join("xdg").join("visigrid"), config_home.join("Library/Application Support/visigrid")] {
        std::fs::create_dir_all(&dir).unwrap();
        std::fs::write(dir.join("recipe_sources.json"), serde_json::json!([key]).to_string()).unwrap();
    }
}

#[test]
fn run_recipe_needs_the_users_approval_then_runs_and_writes() {
    let dir = tempfile::tempdir().unwrap();
    let home = tempfile::tempdir().unwrap();
    std::fs::write(dir.path().join("orders.csv"), "ID,Amount\n001,10\n002,0\n003,7\n").unwrap();
    let recipe = dir.path().join("orders.recipe.toml");
    std::fs::write(
        &recipe,
        "version = 1\n[source]\nkind = \"csv\"\npath = \"orders.csv\"\n[[step]]\nop = \"types\"\ncolumns = { ID = \"text\", Amount = \"number\" }\n[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"0\"\n",
    )
    .unwrap();
    let out = dir.path().join("clean.csv");
    let call = serde_json::json!({"name": "run_recipe", "arguments": {"recipe": recipe, "output": out}});

    // Not approved: refused, nothing read or written
    let r = mcp(home.path(), &[call.clone()]);
    assert_eq!(r[0]["result"]["isError"], true, "{}", r[0]);
    assert!(text(&r[0]).contains("hasn't approved"), "{}", text(&r[0]));
    assert!(!out.exists());

    // Approved in VisiGrid: runs, previews, writes
    approve(home.path(), &recipe);
    let r = mcp(home.path(), &[call]);
    let body: serde_json::Value = serde_json::from_str(&text(&r[0])).unwrap();
    assert_eq!(body["ok"], true, "{body}");
    assert_eq!(body["columns"], serde_json::json!(["ID", "Amount"]));
    assert_eq!(body["preview"], serde_json::json!([["001", "10"], ["003", "7"]]));
    assert_eq!(std::fs::read_to_string(&out).unwrap().lines().collect::<Vec<_>>(), ["ID,Amount", "001,10", "003,7"]);
}

#[test]
fn refresh_table_is_offered() {
    let home = tempfile::tempdir().unwrap();
    let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("mcp")
        .env("HOME", home.path())
        .env("XDG_CONFIG_HOME", home.path().join("xdg"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .unwrap();
    let mut stdin = child.stdin.take().unwrap();
    let mut out = BufReader::new(child.stdout.take().unwrap());
    writeln!(stdin, r#"{{"jsonrpc":"2.0","id":1,"method":"tools/list"}}"#).unwrap();
    let mut line = String::new();
    out.read_line(&mut line).unwrap();
    drop(stdin);
    let _ = child.wait();
    let tools: serde_json::Value = serde_json::from_str(&line).unwrap();
    let names: Vec<&str> = tools["result"]["tools"].as_array().unwrap().iter().filter_map(|t| t["name"].as_str()).collect();
    assert!(names.contains(&"refresh_table") && names.contains(&"run_recipe"), "{names:?}");
}

#[test]
fn run_recipe_rejects_unknown_arguments() {
    let home = tempfile::tempdir().unwrap();
    let r = mcp(home.path(), &[serde_json::json!({"name": "run_recipe", "arguments": {"recipe": "x.recipe.toml", "preview": 2}})]);
    assert!(text(&r[0]).contains("unknown argument: preview"), "{}", r[0]);
}
