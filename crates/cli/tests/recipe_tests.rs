// `vgrid recipe run`: the automation contract.
// Run with: cargo test -p visigrid-cli --test recipe_tests

use std::path::{Path, PathBuf};
use std::process::{Command, Output};

fn vgrid(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_vgrid")).args(args).output().expect("run vgrid")
}

/// A fresh folder per test, so tests can run in parallel.
fn dir(name: &str) -> PathBuf {
    let d = PathBuf::from(env!("CARGO_TARGET_TMPDIR")).join("recipe_tests").join(name);
    let _ = std::fs::remove_dir_all(&d);
    std::fs::create_dir_all(&d).unwrap();
    d
}

const RECIPE: &str = r#"version = 1

[source]
kind = "csv"
path = "export.csv"
header_row = 2
columns = ["Order ID", "Amount", "Date"]

[[step]]
op = "types"
columns = { "Order ID" = "text", Amount = "number", Date = "date:ymd" }

[[step]]
op = "filter"
column = "Amount"
is = ">"
value = "0"
"#;

const SEPTEMBER: &str = "Acme export\nOrder ID,Amount,Date\n00042,120.50,2026-09-01\n00043,0,2026-09-02\n";
const OCTOBER: &str = "Acme export\nOrder ID,Amount,Date\n00051,80,2026-10-01\n00052,15,2026-10-02\n";

fn setup(name: &str) -> PathBuf {
    let d = dir(name);
    std::fs::write(d.join("orders.recipe.toml"), RECIPE).unwrap();
    std::fs::write(d.join("export.csv"), SEPTEMBER).unwrap();
    d
}

fn s(p: &Path) -> &str {
    p.to_str().unwrap()
}

#[test]
fn a_clean_run_writes_the_result_and_reports_steps() {
    let d = setup("clean");
    let out = d.join("orders.csv");
    let o = vgrid(&["recipe", "run", s(&d.join("orders.recipe.toml")), "-o", s(&out)]);
    assert!(o.status.success(), "stderr: {}", String::from_utf8_lossy(&o.stderr));
    // Leading zeros kept, the zero-amount row filtered, the title line skipped
    assert_eq!(std::fs::read_to_string(&out).unwrap(), "Order ID,Amount,Date\n00042,120.5,2026-09-01\n");
    let stderr = String::from_utf8_lossy(&o.stderr);
    assert!(stderr.contains("Keep rows where Amount > 0"), "{stderr}");
    assert!(stderr.contains("removed 1 row"), "{stderr}");
}

#[test]
fn a_failed_run_writes_nothing_and_keeps_the_previous_output() {
    let d = setup("failed");
    let out = d.join("orders.csv");
    std::fs::write(&out, "previous result\n").unwrap();
    // Next month's export renamed a column
    std::fs::write(d.join("export.csv"), OCTOBER.replace("Order ID", "Order Number")).unwrap();
    let o = vgrid(&["recipe", "run", s(&d.join("orders.recipe.toml")), "-o", s(&out)]);
    assert_eq!(o.status.code(), Some(70));
    assert_eq!(std::fs::read_to_string(&out).unwrap(), "previous result\n");
    let stderr = String::from_utf8_lossy(&o.stderr);
    assert!(stderr.contains("Order ID not found"), "{stderr}");
    assert!(stderr.contains("possibly renamed: Order ID -> Order Number"), "{stderr}");
    // No partial file left behind
    let leftovers: Vec<_> = std::fs::read_dir(&d).unwrap().filter_map(|e| e.ok()).filter(|e| e.file_name().to_string_lossy().contains("partial")).collect();
    assert!(leftovers.is_empty());
}

#[test]
fn source_override_reads_next_months_file() {
    let d = setup("override");
    std::fs::write(d.join("october.csv"), OCTOBER).unwrap();
    let o = vgrid(&["recipe", "run", s(&d.join("orders.recipe.toml")), "--source", s(&d.join("october.csv")), "-q"]);
    assert!(o.status.success(), "stderr: {}", String::from_utf8_lossy(&o.stderr));
    assert_eq!(String::from_utf8_lossy(&o.stdout), "Order ID,Amount,Date\n00051,80,2026-10-01\n00052,15,2026-10-02\n");
    assert!(o.stderr.is_empty(), "quiet prints nothing on success");
}

#[test]
fn an_invalid_value_is_located_in_the_json_report() {
    let d = setup("report");
    std::fs::write(d.join("export.csv"), SEPTEMBER.replace("2026-09-01", "2026-09-31")).unwrap();
    let report = d.join("run.json");
    let o = vgrid(&["recipe", "run", s(&d.join("orders.recipe.toml")), "--report", s(&report)]);
    assert_eq!(o.status.code(), Some(70));
    let json: serde_json::Value = serde_json::from_str(&std::fs::read_to_string(&report).unwrap()).unwrap();
    assert_eq!(json["ok"], false);
    assert_eq!(json["error_count"], 1);
    assert_eq!(json["errors"][0]["line"], 3);
    assert_eq!(json["errors"][0]["column"], "Date");
    assert!(o.stdout.is_empty(), "nothing written to stdout on failure");
}

#[test]
fn a_broken_recipe_exits_71() {
    let d = dir("broken");
    std::fs::write(d.join("bad.toml"), "version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n[[step]]\nop = \"explode\"\n").unwrap();
    let o = vgrid(&["recipe", "run", s(&d.join("bad.toml"))]);
    assert_eq!(o.status.code(), Some(71));
}

#[test]
fn colliding_paths_are_refused_before_anything_is_written() {
    let d = setup("collide");
    let out = d.join("orders.csv");
    std::fs::write(&out, "previous result\n").unwrap();
    std::fs::write(d.join("export.csv"), SEPTEMBER.replace("Order ID", "Order Number")).unwrap();
    let recipe = d.join("orders.recipe.toml");
    // Report on top of the output (and a run that would fail)
    let o = vgrid(&["recipe", "run", s(&recipe), "-o", s(&out), "--report", s(&out)]);
    assert_eq!(o.status.code(), Some(2), "{}", String::from_utf8_lossy(&o.stderr));
    assert_eq!(std::fs::read_to_string(&out).unwrap(), "previous result\n");
    // Report on top of the source, through a relative spelling
    let o = vgrid(&["recipe", "run", s(&recipe), "--report", s(&d.join(".").join("export.csv"))]);
    assert_eq!(o.status.code(), Some(2));
    assert!(std::fs::read_to_string(d.join("export.csv")).unwrap().starts_with("Acme export"));
    // Output on top of the recipe
    let o = vgrid(&["recipe", "run", s(&recipe), "-o", s(&recipe)]);
    assert_eq!(o.status.code(), Some(2));
    assert!(std::fs::read_to_string(&recipe).unwrap().starts_with("version = 1"));
}

#[test]
fn recipe_xlsx_output_uses_shared_headless_writer() {
    let d = setup("xlsx_output");
    let path = d.join("orders.xlsx");
    let result = vgrid(&["recipe", "run", s(&d.join("orders.recipe.toml")), "-o", s(&path), "--quiet"]);
    assert!(result.status.success(), "{}", String::from_utf8_lossy(&result.stderr));
    assert!(result.stdout.is_empty());
    let (wb, _) = visigrid_io::xlsx::import(&path).unwrap();
    assert_eq!(wb.active_sheet().get_raw(1, 0), "00042");
    assert_eq!(wb.active_sheet().get_raw(1, 1), "120.5");
    assert_eq!(wb.active_sheet().get_raw(2, 0), "");
}

/// An appending recipe reads every matching file, and later steps can use
/// the Source file column it adds.
#[test]
fn recipe_run_appends_a_folder_and_groups_it() {
    let dir = tempfile::tempdir().unwrap();
    let old = std::time::SystemTime::now() - std::time::Duration::from_secs(60);
    for (name, body) in [("sales-01.csv", "Region,Amount\nWest,1\nEast,2\n"), ("sales-02.csv", "Region,Amount\nWest,3\n")] {
        let p = dir.path().join(name);
        std::fs::write(&p, body).unwrap();
        std::fs::File::options().write(true).open(&p).unwrap().set_modified(old).unwrap();
    }
    std::fs::write(
        dir.path().join("sales.recipe.toml"),
        "version = 1\n[source]\nkind = \"csv\"\npath = \"sales-*.csv\"\ncombine = true\n[[step]]\nop = \"group\"\nby = [\"Region\"]\ntotals = [{ fn = \"sum\", column = \"Amount\", as = \"Total\" }, { fn = \"distinct\", column = \"Source file\", as = \"Files\" }]\n",
    )
    .unwrap();
    let out = std::process::Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .args(["recipe", "run", "sales.recipe.toml", "-q"])
        .current_dir(dir.path())
        .output()
        .unwrap();
    assert!(out.status.success(), "{}", String::from_utf8_lossy(&out.stderr));
    let csv = String::from_utf8(out.stdout).unwrap();
    assert_eq!(csv.lines().collect::<Vec<_>>(), ["Region,Total,Files", "West,4,2", "East,2,1"]);
}

#[test]
fn recipe_run_fills_replaces_splits_and_sorts() {
    let d = dir("shape");
    // A report-style export: the region printed once per group, n/a for no
    // amount, "Last, First" names
    std::fs::write(d.join("report.csv"), "Region,Rep,Amount\nWest,\"Doe, Jane\",9\n,\"Roe, Rick\",n/a\nEast,\"Poe, Ed\",10\n").unwrap();
    std::fs::write(
        d.join("report.recipe.toml"),
        r#"version = 1
[source]
kind = "csv"
path = "report.csv"

[[step]]
op = "fill_down"
columns = ["Region"]

[[step]]
op = "replace"
columns = ["Amount"]
find = "n/a"
with = "0"

[[step]]
op = "split"
column = "Rep"
by = ", "
into = ["Last", "First"]

[[step]]
op = "sort"
by = [{ column = "Amount", descending = true }]
"#,
    )
    .unwrap();
    let o = vgrid(&["recipe", "run", s(&d.join("report.recipe.toml"))]);
    assert!(o.status.success(), "stderr: {}", String::from_utf8_lossy(&o.stderr));
    assert_eq!(
        String::from_utf8_lossy(&o.stdout).lines().collect::<Vec<_>>(),
        ["Region,Last,First,Amount", "East,Poe,Ed,10", "West,Doe,Jane,9", "West,Roe,Rick,0"]
    );
    let stderr = String::from_utf8_lossy(&o.stderr);
    assert!(stderr.contains("Fill down Region") && stderr.contains("filled 1 empty cell"), "{stderr}");
    assert!(stderr.contains("Sort by Amount (descending)"), "{stderr}");
}
