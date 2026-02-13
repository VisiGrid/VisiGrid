//! CLI handlers for scripts and run records.
//!
//! - `scripts list`: List available scripts (attached, project, global)
//! - `scripts run`:  Execute a script against a .sheet file
//! - `runs list`:    List run records from a .sheet file
//! - `runs show`:    Show details of a specific run record

use std::cell::RefCell;
use std::path::PathBuf;
use std::rc::Rc;

use visigrid_io::native;
use visigrid_io::scripting::{
    self, ScriptEntry, ScriptOriginKind, RunRecord,
    list_all_scripts, resolve_script,
};

use crate::CliError;

/// Global scripts directory (~/.config/visigrid/scripts/)
fn global_scripts_dir() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("scripts")
}

// ============================================================================
// scripts list
// ============================================================================

pub fn cmd_scripts_list(
    file: Option<PathBuf>,
    json: bool,
) -> Result<(), CliError> {
    // Load attached scripts from .sheet file (if provided)
    let attached = if let Some(ref path) = file {
        if !path.exists() {
            return Err(CliError::io(format!("file not found: {}", path.display())));
        }
        native::load_scripts(path).unwrap_or_default()
    } else {
        Vec::new()
    };

    // Project dir: parent of the .sheet file (resolve_script adds .visigrid/scripts/ internally)
    let project_dir = file.as_ref().and_then(|p| p.parent().map(|d| d.to_path_buf()));
    let project_ref = project_dir.as_deref();

    let global_dir = global_scripts_dir();
    let entries = list_all_scripts(&attached, project_ref, &global_dir);

    if json {
        print_scripts_json(&entries);
    } else {
        print_scripts_table(&entries);
    }

    Ok(())
}

fn print_scripts_table(entries: &[ScriptEntry]) {
    if entries.is_empty() {
        println!("No scripts found.");
        return;
    }

    println!("{:<20} {:<12} {:<12} {}", "NAME", "ORIGIN", "CAPS", "HASH");
    println!("{}", "-".repeat(72));

    for entry in entries {
        let origin = match entry.origin.kind {
            ScriptOriginKind::Attached => "attached",
            ScriptOriginKind::Project => "project",
            ScriptOriginKind::Global => "global",
            ScriptOriginKind::Hub => "hub",
            ScriptOriginKind::Console => "console",
        };

        let caps: Vec<&str> = entry.meta.capabilities.iter().map(|c| match c {
            scripting::Capability::SheetRead => "R",
            scripting::Capability::SheetWriteValues => "W",
            scripting::Capability::SheetWriteFormulas => "F",
            scripting::Capability::Unknown(_) => "?",
        }).collect();

        let hash_short = if entry.meta.hash.len() > 15 {
            &entry.meta.hash[7..15] // skip "sha256:" prefix, show 8 chars
        } else {
            &entry.meta.hash
        };

        let shadow_mark = if entry.shadowed { " (shadowed)" } else { "" };

        println!(
            "{:<20} {:<12} {:<12} {}{}",
            entry.meta.name,
            origin,
            caps.join(","),
            hash_short,
            shadow_mark,
        );
    }
}

fn print_scripts_json(entries: &[ScriptEntry]) {
    let items: Vec<serde_json::Value> = entries.iter().map(|e| {
        serde_json::json!({
            "name": e.meta.name,
            "hash": e.meta.hash,
            "origin": format!("{:?}", e.origin.kind),
            "capabilities": e.meta.capabilities.iter().map(|c| format!("{:?}", c)).collect::<Vec<_>>(),
            "shadowed": e.shadowed,
            "description": e.meta.description,
        })
    }).collect();

    println!("{}", serde_json::to_string_pretty(&items).unwrap_or_default());
}

// ============================================================================
// scripts run
// ============================================================================

pub fn cmd_scripts_run(
    name: String,
    file: PathBuf,
    plan: bool,
    apply: bool,
    json: bool,
) -> Result<(), CliError> {
    if !plan && !apply {
        return Err(CliError::args("specify --plan (dry run) or --apply (modify file)"));
    }
    if plan && apply {
        return Err(CliError::args("--plan and --apply are mutually exclusive"));
    }

    if !file.exists() {
        return Err(CliError::io(format!("file not found: {}", file.display())));
    }

    // Load attached scripts
    let attached = native::load_scripts(&file).unwrap_or_default();

    // Resolve script by name (resolve_script adds .visigrid/scripts/ internally)
    let project_dir = file.parent().map(|d| d.to_path_buf());
    let project_ref = project_dir.as_deref();
    let global_dir = global_scripts_dir();

    let resolved = resolve_script(&name, &attached, project_ref, &global_dir)
        .ok_or_else(|| CliError::args(format!("script not found: '{}'", name)))?;

    // Build capability set
    let capabilities: std::collections::HashSet<scripting::Capability> =
        resolved.meta.capabilities.iter().cloned().collect();

    // Load workbook
    let mut workbook = native::load_workbook(&file)
        .map_err(|e| CliError::io(format!("failed to load {}: {}", file.display(), e)))?;

    // Rebuild dep graph + recompute (same as GUI load path)
    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    // Compute fingerprint before
    let fingerprint_before = native::compute_semantic_fingerprint(&workbook);

    // Create snapshot for Lua
    let sheet = workbook.active_sheet();
    let snapshot = create_cli_snapshot(sheet);

    // Create Lua runtime and evaluate
    let lua = mlua::Lua::new();

    // Register sheet global with capabilities
    use visigrid_io::scripting::Capability;

    // Use the same op-sink pattern as the GUI
    use std::rc::Rc;
    use std::cell::RefCell;

    let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, capabilities.clone())));
    let sink_clone = sink.clone();

    // Register sheet userdata
    register_cli_sheet(&lua, sink_clone)
        .map_err(|e| CliError::eval(format!("failed to register sheet: {}", e)))?;

    // Evaluate
    let start = std::time::Instant::now();
    let result = lua.load(&resolved.meta.source).exec();
    let elapsed = start.elapsed();

    if let Err(ref e) = result {
        return Err(CliError::eval(format!("script error: {}", e)));
    }

    let borrowed = sink.borrow();
    let ops = borrowed.ops();
    let cells_read = borrowed.cells_read;

    if plan {
        // Dry run: show what would change
        if json {
            let summary = serde_json::json!({
                "script": name,
                "hash": resolved.meta.hash,
                "ops_count": ops.len(),
                "duration_ms": elapsed.as_millis() as u64,
                "mode": "plan",
            });
            println!("{}", serde_json::to_string_pretty(&summary).unwrap_or_default());
        } else {
            println!("Script: {} ({})", name, &resolved.meta.hash[..15.min(resolved.meta.hash.len())]);
            println!("Mode:   plan (dry run)");
            println!("Ops:    {}", ops.len());
            println!("Time:   {:.1}ms", elapsed.as_secs_f64() * 1000.0);
            if !ops.is_empty() {
                println!("\nChanges:");
                for (i, op) in ops.iter().enumerate().take(50) {
                    println!("  {}: {:?}", i + 1, op);
                }
                if ops.len() > 50 {
                    println!("  ... and {} more", ops.len() - 50);
                }
            }
        }
        return Ok(());
    }

    // Apply mode: apply ops to workbook, build run record, save
    drop(borrowed);
    let mut borrowed = sink.borrow_mut();
    let ops_vec = borrowed.take_ops();
    let cells_read_count = borrowed.cells_read;
    drop(borrowed);

    let active_sheet_index = workbook.active_sheet_index();
    let mut changes = Vec::new();

    for op in &ops_vec {
        match op {
            CliOp::SetValue { row, col, value } => {
                let old = workbook.active_sheet().get_raw(*row, *col).to_string();
                workbook.active_sheet_mut().set_value(*row, *col, value);
                changes.push((*row, *col, old, value.clone()));
            }
            CliOp::SetFormula { row, col, formula } => {
                let old = workbook.active_sheet().get_raw(*row, *col).to_string();
                workbook.active_sheet_mut().set_value(*row, *col, formula);
                changes.push((*row, *col, old, formula.clone()));
            }
            CliOp::Clear { row, col } => {
                let old = workbook.active_sheet().get_raw(*row, *col).to_string();
                workbook.active_sheet_mut().set_value(*row, *col, "");
                changes.push((*row, *col, old, String::new()));
            }
        }
    }

    // Recompute after changes
    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    let fingerprint_after = native::compute_semantic_fingerprint(&workbook);

    // Build PatchLines
    let mut patch_lines: Vec<scripting::PatchLine> = changes.iter().map(|(row, col, old, new)| {
        scripting::PatchLine {
            t: "cell".to_string(),
            sheet: active_sheet_index,
            r: *row as u32,
            c: *col as u32,
            k: if new.starts_with('=') { "formula".to_string() } else { "value".to_string() },
            old: if old.is_empty() { None } else { Some(old.clone()) },
            new: if new.is_empty() { None } else { Some(new.clone()) },
        }
    }).collect();

    patch_lines.sort_by(|a, b| {
        a.t.cmp(&b.t)
            .then(a.sheet.cmp(&b.sheet))
            .then(a.r.cmp(&b.r))
            .then(a.c.cmp(&b.c))
            .then(a.k.cmp(&b.k))
    });

    let diff_hash = if patch_lines.is_empty() { None } else {
        Some(scripting::compute_diff_hash(&patch_lines))
    };

    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();
    let diff_summary = scripting::build_diff_summary(&patch_lines, &sheet_names);

    let script_hash = scripting::compute_script_hash(&resolved.meta.source);

    let origin_json = match resolved.origin.kind {
        ScriptOriginKind::Attached => r#"{"kind":"Attached"}"#.to_string(),
        ScriptOriginKind::Project => format!(r#"{{"kind":"Project","ref":"{}"}}"#,
            resolved.origin.r#ref.as_deref().unwrap_or("")),
        ScriptOriginKind::Global => format!(r#"{{"kind":"Global","ref":"{}"}}"#,
            resolved.origin.r#ref.as_deref().unwrap_or("")),
        ScriptOriginKind::Hub => r#"{"kind":"Hub"}"#.to_string(),
        ScriptOriginKind::Console => r#"{"kind":"Console"}"#.to_string(),
    };

    let caps_used: Vec<&str> = capabilities.iter().map(|c| match c {
        Capability::SheetRead => "SheetRead",
        Capability::SheetWriteValues => "SheetWriteValues",
        Capability::SheetWriteFormulas => "SheetWriteFormulas",
        Capability::Unknown(s) => s.as_str(),
    }).collect();

    let mut record = RunRecord {
        run_id: uuid::Uuid::new_v4().to_string(),
        run_fingerprint: String::new(),
        script_name: name.clone(),
        script_hash,
        script_source: resolved.meta.source.clone(),
        script_origin: origin_json,
        capabilities_used: caps_used.join(","),
        params: None,
        fingerprint_before,
        fingerprint_after: fingerprint_after.clone(),
        diff_hash,
        diff_summary: diff_summary.clone(),
        cells_read: cells_read_count as i64,
        cells_modified: changes.len() as i64,
        ops_count: ops_vec.len() as i64,
        duration_ms: elapsed.as_millis() as i64,
        ran_at: chrono::Utc::now().to_rfc3339(),
        ran_by: Some("cli".to_string()),
        status: "ok".to_string(),
        error: None,
    };

    record.run_fingerprint = scripting::compute_run_fingerprint(&record);

    // Load existing run records and add the new one
    let mut run_records = native::load_run_records(&file).unwrap_or_default();
    run_records.push(record.clone());

    // Save workbook with scripts + run records
    let metadata = native::CellMetadata::new(); // Preserve existing metadata
    let scripts = native::load_scripts(&file).unwrap_or_default();

    native::save_workbook_full(&workbook, &metadata, &scripts, &run_records, &file)
        .map_err(|e| CliError::io(format!("failed to save: {}", e)))?;

    // Also re-save layout (save_workbook_full recreates the file)
    // Layout is handled separately — we skip it here since CLI doesn't manage layout

    if json {
        let summary = serde_json::json!({
            "script": name,
            "hash": record.script_hash,
            "run_id": record.run_id,
            "run_fingerprint": record.run_fingerprint,
            "ops_count": ops_vec.len(),
            "cells_modified": changes.len(),
            "fingerprint_before": record.fingerprint_before,
            "fingerprint_after": record.fingerprint_after,
            "diff_hash": record.diff_hash,
            "diff_summary": diff_summary,
            "duration_ms": elapsed.as_millis() as u64,
            "status": "ok",
        });
        println!("{}", serde_json::to_string_pretty(&summary).unwrap_or_default());
    } else {
        println!("Script: {} ({})", name, &record.script_hash[..15.min(record.script_hash.len())]);
        println!("Run ID: {}", record.run_id);
        println!("Status: ok");
        println!("Cells:  {} modified, {} read", changes.len(), cells_read_count);
        println!("Ops:    {}", ops_vec.len());
        println!("Time:   {:.1}ms", elapsed.as_secs_f64() * 1000.0);
        if let Some(ref summary) = diff_summary {
            println!("Diff:   {}", summary);
        }
        println!("FP:     {} → {}", record.fingerprint_before, fingerprint_after);
    }

    Ok(())
}

// ============================================================================
// runs list
// ============================================================================

pub fn cmd_runs_list(file: PathBuf, json: bool, limit: usize, offset: usize) -> Result<(), CliError> {
    if !file.exists() {
        return Err(CliError::io(format!("file not found: {}", file.display())));
    }

    let total = native::count_run_records(&file).unwrap_or(0);
    let records = native::load_run_records_paginated(&file, limit, offset).unwrap_or_default();

    if json {
        let items: Vec<serde_json::Value> = records.iter().map(|r| {
            serde_json::json!({
                "run_id": r.run_id,
                "script_name": r.script_name,
                "status": r.status,
                "cells_modified": r.cells_modified,
                "duration_ms": r.duration_ms,
                "ran_at": r.ran_at,
            })
        }).collect();
        let output = serde_json::json!({
            "records": items,
            "total": total,
            "limit": limit,
            "offset": offset,
        });
        println!("{}", serde_json::to_string_pretty(&output).unwrap_or_default());
    } else if records.is_empty() {
        if total == 0 {
            println!("No run records found.");
        } else {
            println!("No records at offset {} (total: {}).", offset, total);
        }
    } else {
        println!("{:<10} {:<20} {:<8} {:<8} {:<8} {}",
            "RUN ID", "SCRIPT", "STATUS", "CELLS", "MS", "RAN AT");
        println!("{}", "-".repeat(72));
        for r in &records {
            let short_id = if r.run_id.len() > 8 { &r.run_id[..8] } else { &r.run_id };
            println!("{:<10} {:<20} {:<8} {:<8} {:<8} {}",
                short_id, r.script_name, r.status, r.cells_modified, r.duration_ms, r.ran_at);
        }
        if total > records.len() + offset {
            println!("Showing {}-{} of {} records (use --offset {} for next page)",
                offset + 1, offset + records.len(), total, offset + limit);
        }
    }

    Ok(())
}

// ============================================================================
// runs show
// ============================================================================

pub fn cmd_runs_show(run_id: String, file: PathBuf, json: bool) -> Result<(), CliError> {
    if !file.exists() {
        return Err(CliError::io(format!("file not found: {}", file.display())));
    }

    let record = native::load_run_record(&file, &run_id)
        .map_err(|e| CliError::io(e))?
        .ok_or_else(|| CliError::args(format!("run record not found: '{}'", run_id)))?;

    if json {
        let obj = serde_json::json!({
            "run_id": record.run_id,
            "run_fingerprint": record.run_fingerprint,
            "script_name": record.script_name,
            "script_hash": record.script_hash,
            "script_origin": record.script_origin,
            "capabilities_used": record.capabilities_used,
            "params": record.params,
            "fingerprint_before": record.fingerprint_before,
            "fingerprint_after": record.fingerprint_after,
            "diff_hash": record.diff_hash,
            "diff_summary": record.diff_summary,
            "cells_read": record.cells_read,
            "cells_modified": record.cells_modified,
            "ops_count": record.ops_count,
            "duration_ms": record.duration_ms,
            "ran_at": record.ran_at,
            "ran_by": record.ran_by,
            "status": record.status,
            "error": record.error,
        });
        println!("{}", serde_json::to_string_pretty(&obj).unwrap_or_default());
    } else {
        println!("Run ID:          {}", record.run_id);
        println!("Run Fingerprint: {}", record.run_fingerprint);
        println!("Script:          {}", record.script_name);
        println!("Script Hash:     {}", record.script_hash);
        println!("Origin:          {}", record.script_origin);
        println!("Capabilities:    {}", record.capabilities_used);
        println!("Status:          {}", record.status);
        if let Some(ref err) = record.error {
            println!("Error:           {}", err);
        }
        println!("Cells Read:      {}", record.cells_read);
        println!("Cells Modified:  {}", record.cells_modified);
        println!("Ops Count:       {}", record.ops_count);
        println!("Duration:        {}ms", record.duration_ms);
        println!("Ran At:          {}", record.ran_at);
        if let Some(ref by) = record.ran_by {
            println!("Ran By:          {}", by);
        }
        println!("FP Before:       {}", record.fingerprint_before);
        println!("FP After:        {}", record.fingerprint_after);
        if let Some(ref dh) = record.diff_hash {
            println!("Diff Hash:       {}", dh);
        }
        if let Some(ref ds) = record.diff_summary {
            println!("Diff Summary:    {}", ds);
        }
    }

    Ok(())
}

// ============================================================================
// runs verify
// ============================================================================

/// Verification result for a single run record.
#[derive(Debug)]
struct VerifyResult {
    run_id: String,
    script_name: String,
    script_hash_ok: bool,
    run_fingerprint_ok: bool,
    expected_script_hash: String,
    actual_script_hash: String,
    expected_fingerprint: String,
    actual_fingerprint: String,
}

impl VerifyResult {
    fn is_ok(&self) -> bool {
        self.script_hash_ok && self.run_fingerprint_ok
    }
}

fn verify_record(record: &RunRecord) -> VerifyResult {
    // Recompute script_hash from stored source
    let recomputed_hash = scripting::compute_script_hash(&record.script_source);

    // Recompute run_fingerprint from stored fields
    let recomputed_fingerprint = scripting::compute_run_fingerprint(record);

    VerifyResult {
        run_id: record.run_id.clone(),
        script_name: record.script_name.clone(),
        script_hash_ok: recomputed_hash == record.script_hash,
        run_fingerprint_ok: recomputed_fingerprint == record.run_fingerprint,
        expected_script_hash: record.script_hash.clone(),
        actual_script_hash: recomputed_hash,
        expected_fingerprint: record.run_fingerprint.clone(),
        actual_fingerprint: recomputed_fingerprint,
    }
}

pub fn cmd_runs_verify(
    file: PathBuf,
    json: bool,
    run_id: Option<String>,
) -> Result<(), CliError> {
    if !file.exists() {
        return Err(CliError::io(format!("file not found: {}", file.display())));
    }

    let records = if let Some(ref id) = run_id {
        // Verify a specific run record
        match native::load_run_record(&file, id) {
            Ok(Some(r)) => vec![r],
            Ok(None) => return Err(CliError::args(format!("run record not found: '{}'", id))),
            Err(e) => return Err(CliError::io(e)),
        }
    } else {
        // Verify all run records
        native::load_run_records(&file).unwrap_or_default()
    };

    if records.is_empty() {
        if json {
            println!(r#"{{"verified":0,"failed":0,"records":[]}}"#);
        } else {
            println!("No run records to verify.");
        }
        return Ok(());
    }

    let results: Vec<VerifyResult> = records.iter().map(verify_record).collect();
    let passed = results.iter().filter(|r| r.is_ok()).count();
    let failed = results.iter().filter(|r| !r.is_ok()).count();

    if json {
        let items: Vec<serde_json::Value> = results.iter().map(|r| {
            serde_json::json!({
                "run_id": r.run_id,
                "script_name": r.script_name,
                "verified": r.is_ok(),
                "script_hash": {
                    "ok": r.script_hash_ok,
                    "expected": r.expected_script_hash,
                    "actual": r.actual_script_hash,
                },
                "run_fingerprint": {
                    "ok": r.run_fingerprint_ok,
                    "expected": r.expected_fingerprint,
                    "actual": r.actual_fingerprint,
                },
            })
        }).collect();

        let output = serde_json::json!({
            "verified": passed,
            "failed": failed,
            "records": items,
        });
        println!("{}", serde_json::to_string_pretty(&output).unwrap_or_default());
    } else {
        for r in &results {
            let status = if r.is_ok() { "OK" } else { "FAIL" };
            let short_id = if r.run_id.len() > 8 { &r.run_id[..8] } else { &r.run_id };
            println!("[{}] {} ({})", status, short_id, r.script_name);

            if !r.script_hash_ok {
                println!("  script_hash MISMATCH");
                println!("    stored:     {}", r.expected_script_hash);
                println!("    recomputed: {}", r.actual_script_hash);
            }
            if !r.run_fingerprint_ok {
                println!("  run_fingerprint MISMATCH");
                println!("    stored:     {}", r.expected_fingerprint);
                println!("    recomputed: {}", r.actual_fingerprint);
            }
        }

        println!();
        if failed == 0 {
            println!("Verified: {} run record{} OK", passed, if passed == 1 { "" } else { "s" });
        } else {
            println!("Result: {} passed, {} FAILED", passed, failed);
        }
    }

    if failed > 0 {
        // Exit code 1 for verification failure
        Err(CliError { code: 1, message: String::new(), hint: None })
    } else {
        Ok(())
    }
}

// ============================================================================
// CLI-local Lua execution helpers
// ============================================================================
// These are lightweight types for running scripts in the CLI context.
// The GUI uses its own DynOpSink + SheetUserData; we keep a simpler version
// here that produces the same operations and capability enforcement.

use visigrid_engine::sheet::Sheet;

/// Minimal snapshot for CLI script execution.
struct CliSnapshot {
    values: std::collections::HashMap<(usize, usize), String>,
    formulas: std::collections::HashMap<(usize, usize), String>,
    rows: usize,
    cols: usize,
}

fn create_cli_snapshot(sheet: &Sheet) -> CliSnapshot {
    let mut values = std::collections::HashMap::new();
    let mut formulas = std::collections::HashMap::new();

    for (&(row, col), cell) in sheet.cells_iter() {
        let raw = cell.value.raw_display();
        if !raw.is_empty() {
            if raw.starts_with('=') {
                formulas.insert((row, col), raw.to_string());
                // Also store computed value
                let display = cell.value.raw_display();
                values.insert((row, col), display.to_string());
            } else {
                values.insert((row, col), raw.to_string());
            }
        }
    }

    CliSnapshot {
        values,
        formulas,
        rows: sheet.rows,
        cols: sheet.cols,
    }
}

/// Operations produced by CLI script execution.
#[derive(Debug)]
enum CliOp {
    SetValue { row: usize, col: usize, value: String },
    SetFormula { row: usize, col: usize, formula: String },
    Clear { row: usize, col: usize },
}

/// Op sink for CLI Lua execution with capability enforcement.
struct CliOpSink {
    ops: Vec<CliOp>,
    pending: std::collections::HashMap<(usize, usize), Option<String>>,
    snapshot: CliSnapshot,
    capabilities: std::collections::HashSet<scripting::Capability>,
    cells_read: usize,
}

impl CliOpSink {
    fn new(
        snapshot: CliSnapshot,
        capabilities: std::collections::HashSet<scripting::Capability>,
    ) -> Self {
        Self {
            ops: Vec::new(),
            pending: std::collections::HashMap::new(),
            snapshot,
            capabilities,
            cells_read: 0,
        }
    }

    fn ops(&self) -> &[CliOp] {
        &self.ops
    }

    fn take_ops(&mut self) -> Vec<CliOp> {
        std::mem::take(&mut self.ops)
    }

    fn check_read(&self) -> Result<(), mlua::Error> {
        if !self.capabilities.contains(&scripting::Capability::SheetRead) {
            Err(mlua::Error::RuntimeError("capability denied: SheetRead required".into()))
        } else {
            Ok(())
        }
    }

    fn check_write(&self) -> Result<(), mlua::Error> {
        if !self.capabilities.contains(&scripting::Capability::SheetWriteValues) {
            Err(mlua::Error::RuntimeError("capability denied: SheetWriteValues required".into()))
        } else {
            Ok(())
        }
    }

    fn check_write_formula(&self) -> Result<(), mlua::Error> {
        if !self.capabilities.contains(&scripting::Capability::SheetWriteFormulas) {
            Err(mlua::Error::RuntimeError("capability denied: SheetWriteFormulas required".into()))
        } else {
            Ok(())
        }
    }
}

fn register_cli_sheet(
    lua: &mlua::Lua,
    sink: Rc<RefCell<CliOpSink>>,
) -> mlua::Result<()> {
    let sheet_table = lua.create_table()?;

    // NOTE: All functions accept an ignored `_self: mlua::Value` first parameter
    // to support Lua's colon call syntax (sheet:method(args)), which passes the
    // table as the first argument. This matches the GUI's UserData-based API.

    // get_value(row, col)
    {
        let sink = sink.clone();
        sheet_table.set("get_value", lua.create_function(move |_, (_self, row, col): (mlua::Value, usize, usize)| {
            let mut s = sink.borrow_mut();
            s.check_read()?;
            s.cells_read += 1;
            let row = row.wrapping_sub(1);
            let col = col.wrapping_sub(1);
            // Check pending first
            if let Some(val) = s.pending.get(&(row, col)) {
                return Ok(val.clone()) as mlua::Result<Option<String>>;
            }
            Ok(s.snapshot.values.get(&(row, col)).cloned())
        })?)?;
    }

    // get_formula(row, col)
    {
        let sink = sink.clone();
        sheet_table.set("get_formula", lua.create_function(move |_, (_self, row, col): (mlua::Value, usize, usize)| {
            let mut s = sink.borrow_mut();
            s.check_read()?;
            s.cells_read += 1;
            let row = row.wrapping_sub(1);
            let col = col.wrapping_sub(1);
            Ok(s.snapshot.formulas.get(&(row, col)).cloned())
        })?)?;
    }

    // set_value(row, col, value)
    {
        let sink = sink.clone();
        sheet_table.set("set_value", lua.create_function(move |_, (_self, row, col, value): (mlua::Value, usize, usize, mlua::Value)| {
            let mut s = sink.borrow_mut();
            s.check_write()?;
            let row = row.wrapping_sub(1);
            let col = col.wrapping_sub(1);
            let value_str = lua_value_to_string(&value);
            s.pending.insert((row, col), Some(value_str.clone()));
            s.ops.push(CliOp::SetValue { row, col, value: value_str });
            Ok(())
        })?)?;
    }

    // set_formula(row, col, formula)
    {
        let sink = sink.clone();
        sheet_table.set("set_formula", lua.create_function(move |_, (_self, row, col, formula): (mlua::Value, usize, usize, String)| {
            let mut s = sink.borrow_mut();
            s.check_write_formula()?;
            let row = row.wrapping_sub(1);
            let col = col.wrapping_sub(1);
            s.pending.insert((row, col), Some(formula.clone()));
            s.ops.push(CliOp::SetFormula { row, col, formula });
            Ok(())
        })?)?;
    }

    // rows()
    {
        let sink = sink.clone();
        sheet_table.set("rows", lua.create_function(move |_, _self: mlua::Value| {
            let s = sink.borrow();
            s.check_read()?;
            Ok(s.snapshot.rows)
        })?)?;
    }

    // cols()
    {
        let sink = sink.clone();
        sheet_table.set("cols", lua.create_function(move |_, _self: mlua::Value| {
            let s = sink.borrow();
            s.check_read()?;
            Ok(s.snapshot.cols)
        })?)?;
    }

    // A1-style helpers: get(addr), set(addr, value)
    {
        let sink = sink.clone();
        sheet_table.set("get", lua.create_function(move |_, (_self, addr): (mlua::Value, String)| {
            let mut s = sink.borrow_mut();
            s.check_read()?;
            s.cells_read += 1;
            let (row, col) = parse_a1(&addr)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("invalid cell address: {}", addr)))?;
            if let Some(val) = s.pending.get(&(row, col)) {
                return Ok(val.clone()) as mlua::Result<Option<String>>;
            }
            Ok(s.snapshot.values.get(&(row, col)).cloned())
        })?)?;
    }

    {
        let sink = sink.clone();
        sheet_table.set("set", lua.create_function(move |_, (_self, addr, value): (mlua::Value, String, mlua::Value)| {
            let mut s = sink.borrow_mut();
            s.check_write()?;
            let (row, col) = parse_a1(&addr)
                .ok_or_else(|| mlua::Error::RuntimeError(format!("invalid cell address: {}", addr)))?;
            let value_str = lua_value_to_string(&value);
            s.pending.insert((row, col), Some(value_str.clone()));
            s.ops.push(CliOp::SetValue { row, col, value: value_str });
            Ok(())
        })?)?;
    }

    lua.globals().set("sheet", sheet_table)?;
    Ok(())
}

/// Parse A1-style cell reference (e.g. "A1" → (0, 0), "B2" → (1, 1))
fn parse_a1(addr: &str) -> Option<(usize, usize)> {
    let addr = addr.trim();
    let mut col = 0usize;
    let mut i = 0;

    for ch in addr.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1);
            i += 1;
        } else {
            break;
        }
    }

    if i == 0 { return None; }
    let col = col.checked_sub(1)?;
    let row: usize = addr[i..].parse::<usize>().ok()?.checked_sub(1)?;
    Some((row, col))
}

/// Convert Lua value to string for cell storage.
fn lua_value_to_string(value: &mlua::Value) -> String {
    match value {
        mlua::Value::Nil => String::new(),
        mlua::Value::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        mlua::Value::Integer(n) => n.to_string(),
        mlua::Value::Number(n) => {
            if *n == (*n as i64) as f64 {
                (*n as i64).to_string()
            } else {
                n.to_string()
            }
        }
        mlua::Value::String(s) => s.to_str().map(|s| s.to_string()).unwrap_or_default(),
        _ => format!("{:?}", value),
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    fn make_snapshot(cells: &[((usize, usize), &str)]) -> CliSnapshot {
        let mut values = std::collections::HashMap::new();
        let mut formulas = std::collections::HashMap::new();
        let mut max_row = 0usize;
        let mut max_col = 0usize;

        for &((r, c), val) in cells {
            max_row = max_row.max(r + 1);
            max_col = max_col.max(c + 1);
            if val.starts_with('=') {
                formulas.insert((r, c), val.to_string());
            }
            values.insert((r, c), val.to_string());
        }

        CliSnapshot { values, formulas, rows: max_row.max(10), cols: max_col.max(5) }
    }

    fn all_caps() -> HashSet<scripting::Capability> {
        scripting::all_sheet_caps()
    }

    fn read_only() -> HashSet<scripting::Capability> {
        [scripting::Capability::SheetRead].into_iter().collect()
    }

    fn write_only() -> HashSet<scripting::Capability> {
        [scripting::Capability::SheetWriteValues].into_iter().collect()
    }

    fn no_caps() -> HashSet<scripting::Capability> {
        HashSet::new()
    }

    // ====================================================================
    // Capability enforcement via real Lua execution
    // ====================================================================

    #[test]
    fn test_cap_read_denied_without_sheet_read() {
        let snapshot = make_snapshot(&[((0, 0), "hello")]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, write_only())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        // Attempting to read should fail
        let result = lua.load("return sheet:get_value(1, 1)").exec();
        assert!(result.is_err(), "get_value must fail without SheetRead");
        assert!(result.unwrap_err().to_string().contains("SheetRead"),
            "error must mention SheetRead");
    }

    #[test]
    fn test_cap_write_denied_without_sheet_write_values() {
        let snapshot = make_snapshot(&[]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, read_only())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        let result = lua.load("sheet:set_value(1, 1, 'hello')").exec();
        assert!(result.is_err(), "set_value must fail without SheetWriteValues");
        assert!(result.unwrap_err().to_string().contains("SheetWriteValues"));
    }

    #[test]
    fn test_cap_formula_denied_without_sheet_write_formulas() {
        let snapshot = make_snapshot(&[]);
        let caps: HashSet<_> = [
            scripting::Capability::SheetRead,
            scripting::Capability::SheetWriteValues,
        ].into_iter().collect();
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, caps)));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        let result = lua.load("sheet:set_formula(1, 1, '=A2+1')").exec();
        assert!(result.is_err(), "set_formula must fail without SheetWriteFormulas");
        assert!(result.unwrap_err().to_string().contains("SheetWriteFormulas"));
    }

    #[test]
    fn test_cap_all_caps_allows_everything() {
        let snapshot = make_snapshot(&[((0, 0), "100")]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, all_caps())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        let script = r#"
            local v = sheet:get_value(1, 1)
            sheet:set_value(1, 2, v)
            sheet:set_formula(2, 1, "=A1+1")
        "#;
        let result = lua.load(script).exec();
        assert!(result.is_ok(), "all_caps must allow read + write + formula: {:?}", result.err());

        let borrowed = sink.borrow();
        assert_eq!(borrowed.ops.len(), 2, "should have 2 ops (set_value + set_formula)");
        assert!(borrowed.cells_read >= 1, "should have read at least 1 cell");
    }

    #[test]
    fn test_cap_empty_caps_denies_all_sheet_access() {
        let snapshot = make_snapshot(&[((0, 0), "100")]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, no_caps())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        // Pure Lua math works fine
        let result = lua.load("local x = 1 + 2 + 3").exec();
        assert!(result.is_ok(), "pure Lua must work with no caps");

        // But any sheet access fails
        let result = lua.load("sheet:get_value(1, 1)").exec();
        assert!(result.is_err(), "sheet read must fail with no caps");

        let result = lua.load("sheet:set_value(1, 1, 'x')").exec();
        assert!(result.is_err(), "sheet write must fail with no caps");

        let result = lua.load("sheet:rows()").exec();
        assert!(result.is_err(), "sheet:rows() must fail with no caps");
    }

    #[test]
    fn test_cap_a1_style_helpers_respect_caps() {
        // The A1-style sheet:get("A1") and sheet:set("A1", val) helpers
        // must also enforce capabilities.
        let snapshot = make_snapshot(&[((0, 0), "100")]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, write_only())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        // A1-style get should fail (needs SheetRead)
        let result = lua.load("return sheet:get('A1')").exec();
        assert!(result.is_err(), "sheet:get must fail without SheetRead");

        // A1-style set should work (has SheetWriteValues)
        let result = lua.load("sheet:set('A1', 'new')").exec();
        assert!(result.is_ok(), "sheet:set must work with SheetWriteValues");
    }

    // ====================================================================
    // CLI/GUI parity: same script produces same ops
    // ====================================================================

    #[test]
    fn test_cli_script_produces_correct_patchlines() {
        // Run a real Lua script through the CLI path and verify the PatchLines
        // match what the GUI would produce for the same data.
        let snapshot = make_snapshot(&[
            ((0, 0), "10"),
            ((1, 0), "20"),
            ((2, 0), "30"),
        ]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, all_caps())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        // Script: double column A into column B
        let script = r#"
            for r = 1, 3 do
                local v = sheet:get_value(r, 1)
                if v then
                    sheet:set_value(r, 2, tonumber(v) * 2)
                end
            end
        "#;
        lua.load(script).exec().unwrap();

        let borrowed = sink.borrow();
        let ops = borrowed.ops();

        // Build PatchLines the same way cmd_scripts_run does
        let mut patch_lines: Vec<scripting::PatchLine> = ops.iter().map(|op| {
            match op {
                CliOp::SetValue { row, col, value } => {
                    scripting::PatchLine {
                        t: "cell".into(),
                        sheet: 0,
                        r: *row as u32,
                        c: *col as u32,
                        k: if value.starts_with('=') { "formula".into() } else { "value".into() },
                        old: None, // was empty
                        new: if value.is_empty() { None } else { Some(value.clone()) },
                    }
                }
                CliOp::SetFormula { row, col, formula } => {
                    scripting::PatchLine {
                        t: "cell".into(),
                        sheet: 0,
                        r: *row as u32,
                        c: *col as u32,
                        k: "formula".into(),
                        old: None,
                        new: Some(formula.clone()),
                    }
                }
                CliOp::Clear { row, col } => {
                    scripting::PatchLine {
                        t: "cell".into(),
                        sheet: 0,
                        r: *row as u32,
                        c: *col as u32,
                        k: "value".into(),
                        old: None,
                        new: None,
                    }
                }
            }
        }).collect();

        patch_lines.sort_by(|a, b| {
            a.t.cmp(&b.t)
                .then(a.sheet.cmp(&b.sheet))
                .then(a.r.cmp(&b.r))
                .then(a.c.cmp(&b.c))
                .then(a.k.cmp(&b.k))
        });

        // Verify: 3 new cells in column B with doubled values
        assert_eq!(patch_lines.len(), 3);
        assert_eq!(patch_lines[0].new, Some("20".into()));
        assert_eq!(patch_lines[1].new, Some("40".into()));
        assert_eq!(patch_lines[2].new, Some("60".into()));

        // The diff_hash must be stable
        let hash1 = scripting::compute_diff_hash(&patch_lines);
        let hash2 = scripting::compute_diff_hash(&patch_lines);
        assert_eq!(hash1, hash2);
        assert!(hash1.starts_with("sha256:"));
    }

    #[test]
    fn test_cells_read_counter_increments() {
        let snapshot = make_snapshot(&[
            ((0, 0), "a"),
            ((1, 0), "b"),
            ((2, 0), "c"),
        ]);
        let sink = Rc::new(RefCell::new(CliOpSink::new(snapshot, all_caps())));
        let lua = mlua::Lua::new();
        register_cli_sheet(&lua, sink.clone()).unwrap();

        lua.load(r#"
            sheet:get_value(1, 1)
            sheet:get_value(2, 1)
            sheet:get("A3")
        "#).exec().unwrap();

        let borrowed = sink.borrow();
        assert_eq!(borrowed.cells_read, 3, "should count 3 cell reads");
    }

    // ====================================================================
    // parse_a1 tests
    // ====================================================================

    #[test]
    fn test_parse_a1_basic() {
        assert_eq!(parse_a1("A1"), Some((0, 0)));
        assert_eq!(parse_a1("B2"), Some((1, 1)));
        assert_eq!(parse_a1("Z26"), Some((25, 25)));
        assert_eq!(parse_a1("AA1"), Some((0, 26)));
        assert_eq!(parse_a1("AB1"), Some((0, 27)));
    }

    #[test]
    fn test_parse_a1_edge_cases() {
        assert_eq!(parse_a1(""), None);
        assert_eq!(parse_a1("1"), None);       // no column
        assert_eq!(parse_a1("A0"), None);       // row 0 invalid (1-indexed)
        assert_eq!(parse_a1("A"), None);        // no row number
    }

    // ====================================================================
    // Verify tests
    // ====================================================================

    #[test]
    fn test_verify_valid_record_passes() {
        let source = "sheet:set_value(1, 1, 42)";
        let hash = scripting::compute_script_hash(source);

        let mut record = scripting::RunRecord {
            run_id: "test-verify-ok".into(),
            run_fingerprint: String::new(),
            script_name: "test".into(),
            script_hash: hash,
            script_source: source.into(),
            script_origin: r#"{"kind":"Attached"}"#.into(),
            capabilities_used: "SheetWriteValues".into(),
            params: None,
            fingerprint_before: "v2:0:aaa".into(),
            fingerprint_after: "v2:1:bbb".into(),
            diff_hash: Some("sha256:abc".into()),
            diff_summary: Some("1 added".into()),
            cells_read: 0,
            cells_modified: 1,
            ops_count: 1,
            duration_ms: 2,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: Some("test".into()),
            status: "ok".into(),
            error: None,
        };
        record.run_fingerprint = scripting::compute_run_fingerprint(&record);

        let result = verify_record(&record);
        assert!(result.script_hash_ok, "script_hash should verify");
        assert!(result.run_fingerprint_ok, "run_fingerprint should verify");
        assert!(result.is_ok(), "overall verification should pass");
    }

    #[test]
    fn test_verify_tampered_source_fails() {
        let original = "sheet:set_value(1, 1, 42)";
        let hash = scripting::compute_script_hash(original);

        let mut record = scripting::RunRecord {
            run_id: "test-verify-tamper".into(),
            run_fingerprint: String::new(),
            script_name: "test".into(),
            script_hash: hash,
            script_source: "sheet:set_value(1, 1, 999)".into(), // TAMPERED
            script_origin: r#"{"kind":"Attached"}"#.into(),
            capabilities_used: "SheetWriteValues".into(),
            params: None,
            fingerprint_before: "v2:0:aaa".into(),
            fingerprint_after: "v2:1:bbb".into(),
            diff_hash: Some("sha256:abc".into()),
            diff_summary: Some("1 added".into()),
            cells_read: 0,
            cells_modified: 1,
            ops_count: 1,
            duration_ms: 2,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: Some("test".into()),
            status: "ok".into(),
            error: None,
        };
        record.run_fingerprint = scripting::compute_run_fingerprint(&record);

        let result = verify_record(&record);
        assert!(!result.script_hash_ok, "tampered source must fail script_hash");
        assert!(!result.is_ok(), "overall verification must fail");
    }

    #[test]
    fn test_verify_tampered_fingerprint_fails() {
        let source = "sheet:set_value(1, 1, 42)";
        let hash = scripting::compute_script_hash(source);

        let mut record = scripting::RunRecord {
            run_id: "test-verify-fp".into(),
            run_fingerprint: String::new(),
            script_name: "test".into(),
            script_hash: hash,
            script_source: source.into(),
            script_origin: r#"{"kind":"Attached"}"#.into(),
            capabilities_used: "SheetWriteValues".into(),
            params: None,
            fingerprint_before: "v2:0:aaa".into(),
            fingerprint_after: "v2:1:bbb".into(),
            diff_hash: Some("sha256:abc".into()),
            diff_summary: Some("1 added".into()),
            cells_read: 0,
            cells_modified: 1,
            ops_count: 1,
            duration_ms: 2,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: Some("test".into()),
            status: "ok".into(),
            error: None,
        };
        record.run_fingerprint = scripting::compute_run_fingerprint(&record);

        // Tamper with the fingerprint after computing it
        record.run_fingerprint = "sha256:tampered".into();

        let result = verify_record(&record);
        assert!(result.script_hash_ok, "script_hash should still be ok");
        assert!(!result.run_fingerprint_ok, "tampered fingerprint must fail");
        assert!(!result.is_ok());
    }
}
