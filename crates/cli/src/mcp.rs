//! MCP (Model Context Protocol) stdio server — `vgrid mcp`.
//!
//! Bridges an MCP client (Claude Code, Claude Desktop, any MCP host) to a
//! running VisiGrid GUI session over the existing session protocol. The MCP
//! host spawns this process and speaks newline-delimited JSON-RPC 2.0 on
//! stdio; each tool call opens a short-lived TCP connection to the session
//! server, so a stale GUI restart never wedges the bridge.
//!
//! Auth: the session token comes from VISIGRID_SESSION_TOKEN (set it in the
//! MCP host's env config; the GUI session panel shows the token). Pairing
//! flow is planned to replace this — see planning/visigrid/features/mcp-v1.md.

use std::io::{BufRead, Write};

use serde_json::{json, Value};

use crate::session::{self, SessionClient};
use crate::{parse_cell_ref, CliError};
use visigrid_protocol::{InspectResult, Op};

/// MCP protocol revision this server implements.
const MCP_PROTOCOL_VERSION: &str = "2025-06-18";

/// Instructions surfaced to the model by MCP hosts.
const SERVER_INSTRUCTIONS: &str = "VisiGrid is a native spreadsheet running on this machine; these tools drive a live GUI window the user can see. Reads return display values plus formulas. Writes land in the user's undo history and render immediately. Batch related edits into one write_cells call; pass expected_revision (from any read) to avoid clobbering concurrent human edits, and re-read on revision_mismatch. Coordinates are A1-style; sheet is a 0-based index.";

pub fn cmd_mcp(session_id: Option<String>) -> Result<(), CliError> {
    let stdin = std::io::stdin();
    let mut stdout = std::io::stdout();

    let mut server = McpServer { session_pref: session_id };

    for line in stdin.lock().lines() {
        let line = line.map_err(|e| CliError::io(format!("stdin read failed: {}", e)))?;
        if line.trim().is_empty() {
            continue;
        }
        if let Some(response) = server.handle_line(&line) {
            let mut out = serde_json::to_string(&response)
                .map_err(|e| CliError::io(format!("response serialization failed: {}", e)))?;
            out.push('\n');
            stdout
                .write_all(out.as_bytes())
                .and_then(|_| stdout.flush())
                .map_err(|e| CliError::io(format!("stdout write failed: {}", e)))?;
        }
    }
    Ok(())
}

struct McpServer {
    /// --session prefix from the command line; resolved per call.
    session_pref: Option<String>,
}

impl McpServer {
    /// Handle one JSON-RPC line. Returns None for notifications (no reply).
    fn handle_line(&mut self, line: &str) -> Option<Value> {
        let msg: Value = match serde_json::from_str(line) {
            Ok(v) => v,
            Err(e) => {
                return Some(json!({
                    "jsonrpc": "2.0",
                    "id": null,
                    "error": { "code": -32700, "message": format!("parse error: {}", e) },
                }));
            }
        };
        let id = msg.get("id").filter(|v| !v.is_null()).cloned();
        let method = msg.get("method").and_then(|m| m.as_str()).unwrap_or("");

        let result = match method {
            "initialize" => {
                // Echo the client's requested version when present — we speak
                // the tools-only core, which is stable across revisions.
                let version = msg
                    .pointer("/params/protocolVersion")
                    .and_then(|v| v.as_str())
                    .unwrap_or(MCP_PROTOCOL_VERSION);
                Ok(json!({
                    "protocolVersion": version,
                    "capabilities": { "tools": {} },
                    "serverInfo": {
                        "name": "visigrid",
                        "version": env!("CARGO_PKG_VERSION"),
                    },
                    "instructions": SERVER_INSTRUCTIONS,
                }))
            }
            "ping" => Ok(json!({})),
            "tools/list" => Ok(json!({ "tools": tool_definitions() })),
            "tools/call" => {
                let name = msg.pointer("/params/name").and_then(|v| v.as_str()).unwrap_or("");
                let args = msg.pointer("/params/arguments").cloned().unwrap_or(json!({}));
                Ok(self.call_tool(name, &args))
            }
            _ => {
                if id.is_none() {
                    return None; // unknown notification — ignore per spec
                }
                Err((-32601, format!("method not found: {}", method)))
            }
        };

        // Notifications never get a reply, even for handled methods.
        let id = id?;
        Some(match result {
            Ok(result) => json!({ "jsonrpc": "2.0", "id": id, "result": result }),
            Err((code, message)) => json!({
                "jsonrpc": "2.0",
                "id": id,
                "error": { "code": code, "message": message },
            }),
        })
    }

    /// Dispatch a tools/call. Tool failures are returned as isError results
    /// (visible to the model for self-correction), not JSON-RPC errors.
    fn call_tool(&mut self, name: &str, args: &Value) -> Value {
        let outcome = match name {
            "list_sessions" => self.tool_list_sessions(),
            "get_workbook" => self.tool_get_workbook(args),
            "read_range" => self.tool_read_range(args),
            "write_cells" => self.tool_write_cells(args),
            "set_format" => self.tool_set_format(args),
            _ => Err(format!("unknown tool: {}", name)),
        };
        match outcome {
            Ok(text) => json!({ "content": [{ "type": "text", "text": text }], "isError": false }),
            Err(text) => json!({ "content": [{ "type": "text", "text": text }], "isError": true }),
        }
    }

    // ------------------------------------------------------------------
    // Tools
    // ------------------------------------------------------------------

    fn tool_list_sessions(&self) -> Result<String, String> {
        let sessions = session::list_sessions()
            .map_err(|e| format!("failed to list sessions: {}", e))?;
        if sessions.is_empty() {
            return Err("No running VisiGrid sessions. Ask the user to start VisiGrid — the session server starts with the GUI.".to_string());
        }
        let rows: Vec<Value> = sessions
            .iter()
            .map(|s| {
                json!({
                    "session": s.session_id.to_string(),
                    "title": s.workbook_title,
                    "path": s.workbook_path,
                    "created_at": s.created_at.to_rfc3339(),
                })
            })
            .collect();
        serde_json::to_string_pretty(&rows).map_err(|e| e.to_string())
    }

    fn tool_get_workbook(&mut self, args: &Value) -> Result<String, String> {
        let mut client = self.connect(args)?;
        let result = client.inspect_workbook().map_err(session_error_text)?;
        match result.result {
            InspectResult::Workbook(info) => serde_json::to_string_pretty(&json!({
                "revision": result.revision,
                "title": info.title,
                "sheet_count": info.sheet_count,
                "active_sheet": info.active_sheet,
            }))
            .map_err(|e| e.to_string()),
            _ => Err("unexpected inspect result".to_string()),
        }
    }

    fn tool_read_range(&mut self, args: &Value) -> Result<String, String> {
        let range = require_str(args, "range")?;
        let sheet = args.get("sheet").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
        let ((start_row, start_col), (end_row, end_col)) = parse_a1_range(range)?;

        let mut client = self.connect(args)?;
        let result = client
            .inspect_range(sheet, start_row, start_col, end_row, end_col)
            .map_err(session_error_text)?;

        let cells = match result.result {
            InspectResult::Range { cells } => cells,
            InspectResult::Cell(info) => vec![info],
            _ => return Err("unexpected inspect result".to_string()),
        };

        // Row-major display grid + a sparse formula map for formula cells.
        let num_cols = end_col - start_col + 1;
        let mut rows: Vec<Vec<String>> = Vec::new();
        let mut formulas = serde_json::Map::new();
        for (i, cell) in cells.iter().enumerate() {
            let r = i / num_cols;
            let c = i % num_cols;
            if c == 0 {
                rows.push(Vec::with_capacity(num_cols));
            }
            rows[r].push(cell.display.clone());
            if let Some(f) = &cell.formula {
                formulas.insert(cell_name(start_row + r, start_col + c), json!(f));
            }
        }

        serde_json::to_string_pretty(&json!({
            "revision": result.revision,
            "sheet": sheet,
            "range": range.to_uppercase(),
            "rows": rows,
            "formulas": formulas,
        }))
        .map_err(|e| e.to_string())
    }

    fn tool_write_cells(&mut self, args: &Value) -> Result<String, String> {
        let sheet = args.get("sheet").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
        let cells = args
            .get("cells")
            .and_then(|v| v.as_array())
            .ok_or("missing required argument: cells (array)")?;
        if cells.is_empty() {
            return Err("cells array is empty".to_string());
        }

        let mut ops = Vec::with_capacity(cells.len());
        let mut preview = Vec::with_capacity(cells.len());
        for (i, entry) in cells.iter().enumerate() {
            let cell_ref = entry
                .get("cell")
                .and_then(|v| v.as_str())
                .ok_or(format!("cells[{}]: missing 'cell' (A1-style reference)", i))?;
            let (row, col) = parse_cell_ref(cell_ref)
                .ok_or(format!("cells[{}]: invalid cell reference '{}'", i, cell_ref))?;

            let value = entry.get("value").and_then(|v| v.as_str());
            let formula = entry.get("formula").and_then(|v| v.as_str());
            let clear = entry.get("clear").and_then(|v| v.as_bool()).unwrap_or(false);

            let (op, describe) = match (value, formula, clear) {
                (Some(v), None, false) => (
                    Op::SetCellValue { sheet, row, col, value: v.to_string() },
                    format!("{} = {:?}", cell_ref.to_uppercase(), v),
                ),
                (None, Some(f), false) => {
                    let f = if f.starts_with('=') { f.to_string() } else { format!("={}", f) };
                    let describe = format!("{} = {}", cell_ref.to_uppercase(), f);
                    (Op::SetCellFormula { sheet, row, col, formula: f }, describe)
                }
                (None, None, true) => (
                    Op::ClearCell { sheet, row, col },
                    format!("{} cleared", cell_ref.to_uppercase()),
                ),
                _ => {
                    return Err(format!(
                        "cells[{}]: provide exactly one of 'value', 'formula', or 'clear': true",
                        i
                    ));
                }
            };
            ops.push(op);
            preview.push(describe);
        }

        if args.get("dry_run").and_then(|v| v.as_bool()).unwrap_or(false) {
            return serde_json::to_string_pretty(&json!({
                "dry_run": true,
                "would_apply": preview,
                "sheet": sheet,
            }))
            .map_err(|e| e.to_string());
        }

        let atomic = args.get("atomic").and_then(|v| v.as_bool()).unwrap_or(true);
        let expected_revision = args.get("expected_revision").and_then(|v| v.as_u64());

        let mut client = self.connect(args)?;
        let result = client
            .apply_ops(ops, atomic, expected_revision)
            .map_err(session_error_text)?;

        if let Some(err) = result.error {
            return Err(apply_error_text(&err, result.revision));
        }
        serde_json::to_string_pretty(&json!({
            "applied": result.applied,
            "total": result.total,
            "revision": result.revision,
        }))
        .map_err(|e| e.to_string())
    }

    fn tool_set_format(&mut self, args: &Value) -> Result<String, String> {
        let range = require_str(args, "range")?;
        let sheet = args.get("sheet").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
        let ((start_row, start_col), (end_row, end_col)) = parse_a1_range(range)?;

        let bold = args.get("bold").and_then(|v| v.as_bool());
        let italic = args.get("italic").and_then(|v| v.as_bool());
        let underline = args.get("underline").and_then(|v| v.as_bool());
        let number_format = args.get("number_format").and_then(|v| v.as_str());

        let mut ops = Vec::new();
        let mut preview = Vec::new();
        if bold.is_some() || italic.is_some() || underline.is_some() {
            ops.push(Op::SetStyle {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                bold,
                italic,
                underline,
            });
            preview.push(format!(
                "style {}: bold={:?} italic={:?} underline={:?}",
                range.to_uppercase(),
                bold,
                italic,
                underline
            ));
        }
        if let Some(nf) = number_format {
            ops.push(Op::SetNumberFormat {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                format: nf.to_string(),
            });
            preview.push(format!("number format {}: {}", range.to_uppercase(), nf));
        }
        if ops.is_empty() {
            return Err(
                "nothing to do: provide bold/italic/underline and/or number_format".to_string()
            );
        }

        if args.get("dry_run").and_then(|v| v.as_bool()).unwrap_or(false) {
            return serde_json::to_string_pretty(&json!({
                "dry_run": true,
                "would_apply": preview,
                "sheet": sheet,
            }))
            .map_err(|e| e.to_string());
        }

        let expected_revision = args.get("expected_revision").and_then(|v| v.as_u64());
        let mut client = self.connect(args)?;
        let result = client
            .apply_ops(ops, true, expected_revision)
            .map_err(session_error_text)?;

        if let Some(err) = result.error {
            return Err(apply_error_text(&err, result.revision));
        }
        serde_json::to_string_pretty(&json!({
            "applied": result.applied,
            "total": result.total,
            "revision": result.revision,
        }))
        .map_err(|e| e.to_string())
    }

    // ------------------------------------------------------------------
    // Session plumbing
    // ------------------------------------------------------------------

    /// Connect to the target session: per-call `session` argument beats the
    /// --session flag; with neither, auto-selects when exactly one is running.
    fn connect(&self, args: &Value) -> Result<SessionClient, String> {
        let pref = args
            .get("session")
            .and_then(|v| v.as_str())
            .map(str::to_string)
            .or_else(|| self.session_pref.clone());

        let sessions = session::list_sessions()
            .map_err(|e| format!("failed to list sessions: {}", e))?;
        let discovery = match &pref {
            Some(id) => session::find_session(id)
                .map_err(|e| e.to_string())?
                .ok_or(format!("session '{}' not found — use list_sessions", id))?,
            None => match sessions.len() {
                0 => return Err("No running VisiGrid sessions. Ask the user to start VisiGrid.".to_string()),
                1 => sessions.into_iter().next().unwrap(),
                n => {
                    return Err(format!(
                        "{} sessions running — pass a 'session' argument (use list_sessions to pick)",
                        n
                    ))
                }
            },
        };

        let token = std::env::var("VISIGRID_SESSION_TOKEN").map_err(|_| {
            "VISIGRID_SESSION_TOKEN is not set. Add it to this MCP server's env config; the user can copy the token from the VisiGrid session panel.".to_string()
        })?;

        SessionClient::connect(&discovery, &token).map_err(session_error_text)
    }
}

// ----------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------

fn require_str<'a>(args: &'a Value, key: &str) -> Result<&'a str, String> {
    args.get(key)
        .and_then(|v| v.as_str())
        .ok_or(format!("missing required argument: {}", key))
}

/// Parse "A1" or "A1:D10" into ((start_row, start_col), (end_row, end_col)).
fn parse_a1_range(range: &str) -> Result<((usize, usize), (usize, usize)), String> {
    let invalid = || format!("invalid range '{}' — expected A1 or A1:D10", range);
    match range.split_once(':') {
        Some((start, end)) => {
            let s = parse_cell_ref(start).ok_or_else(invalid)?;
            let e = parse_cell_ref(end).ok_or_else(invalid)?;
            if s.0 > e.0 || s.1 > e.1 {
                return Err(format!("range '{}' start is after its end", range));
            }
            Ok((s, e))
        }
        None => {
            let c = parse_cell_ref(range).ok_or_else(invalid)?;
            Ok((c, c))
        }
    }
}

/// 0-based (row, col) → A1-style name.
fn cell_name(row: usize, col: usize) -> String {
    let mut letters = String::new();
    let mut c = col + 1;
    while c > 0 {
        let rem = (c - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        c = (c - 1) / 26;
    }
    format!("{}{}", letters, row + 1)
}

fn session_error_text(e: session::SessionError) -> String {
    match e {
        session::SessionError::ServerError { code, message, retry_after_ms } => {
            let mut text = format!("{}: {}", code, message);
            if code == "revision_mismatch" {
                text.push_str(" — re-read the affected range and retry with the new revision");
            }
            if let Some(ms) = retry_after_ms {
                text.push_str(&format!(" (retry after {}ms)", ms));
            }
            text
        }
        session::SessionError::AuthFailed(msg) => format!(
            "auth_failed: {} — VISIGRID_SESSION_TOKEN may be stale; the token changes when the GUI restarts",
            msg
        ),
        other => other.to_string(),
    }
}

fn apply_error_text(err: &visigrid_protocol::OpError, revision: u64) -> String {
    let mut text = format!("{} (op {}): {}", err.code, err.op_index, err.message);
    if let Some(s) = &err.suggestion {
        text.push_str(&format!(" — {}", s));
    }
    text.push_str(&format!(" [current revision: {}]", revision));
    text
}

/// Static tool catalog. Names follow the visibooks-mcp conventions:
/// get_*/list_* reads, dry_run on every mutating verb.
fn tool_definitions() -> Value {
    json!([
        {
            "name": "list_sessions",
            "description": "List running VisiGrid sessions (open GUI windows) with their workbook titles and session IDs. Use when get_workbook reports multiple sessions or to find a specific workbook.",
            "inputSchema": { "type": "object", "properties": {}, "additionalProperties": false }
        },
        {
            "name": "get_workbook",
            "description": "Get the current workbook's title, sheet count, active sheet index, and revision. Call this first to orient; the revision feeds expected_revision on writes.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "session": { "type": "string", "description": "Session ID (prefix ok). Omit when one session is running." }
                },
                "additionalProperties": false
            }
        },
        {
            "name": "read_range",
            "description": "Read a cell or range: display values in row-major order plus a map of formula cells. Max 65,536 cells per request (one full column).",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "range": { "type": "string", "description": "A1-style cell or range, e.g. 'B2' or 'A1:D10'" },
                    "sheet": { "type": "integer", "description": "0-based sheet index (default 0)" },
                    "session": { "type": "string", "description": "Session ID (prefix ok). Omit when one session is running." }
                },
                "required": ["range"],
                "additionalProperties": false
            }
        },
        {
            "name": "write_cells",
            "description": "Write values and formulas to cells in the live workbook. Edits render immediately in the user's GUI and land in undo history. Batch related edits into one call — a batch is one undo step and one recalc. Pass expected_revision from a prior read to avoid clobbering concurrent human edits.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "cells": {
                        "type": "array",
                        "description": "Cells to write. Each entry sets exactly one of value, formula, or clear.",
                        "items": {
                            "type": "object",
                            "properties": {
                                "cell": { "type": "string", "description": "A1-style reference" },
                                "value": { "type": "string", "description": "Literal value to set" },
                                "formula": { "type": "string", "description": "Formula (leading '=' optional)" },
                                "clear": { "type": "boolean", "description": "true to clear the cell" }
                            },
                            "required": ["cell"]
                        }
                    },
                    "sheet": { "type": "integer", "description": "0-based sheet index (default 0)" },
                    "atomic": { "type": "boolean", "description": "All-or-nothing (default true)" },
                    "expected_revision": { "type": "integer", "description": "Fail with revision_mismatch if the workbook changed since this revision" },
                    "dry_run": { "type": "boolean", "description": "Preview the planned edits without applying" },
                    "session": { "type": "string", "description": "Session ID (prefix ok). Omit when one session is running." }
                },
                "required": ["cells"],
                "additionalProperties": false
            }
        },
        {
            "name": "set_format",
            "description": "Format a range: bold/italic/underline and/or a number format. Number formats: named ('general', 'number:2', 'currency:2', 'percent:1', 'date', 'time', 'datetime') or a raw Excel code like '#,##0.00'. Max 250,000 cells per call.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "range": { "type": "string", "description": "A1-style cell or range" },
                    "sheet": { "type": "integer", "description": "0-based sheet index (default 0)" },
                    "bold": { "type": "boolean" },
                    "italic": { "type": "boolean" },
                    "underline": { "type": "boolean" },
                    "number_format": { "type": "string", "description": "Named format or Excel format code" },
                    "expected_revision": { "type": "integer", "description": "Fail with revision_mismatch if the workbook changed since this revision" },
                    "dry_run": { "type": "boolean", "description": "Preview without applying" },
                    "session": { "type": "string", "description": "Session ID (prefix ok). Omit when one session is running." }
                },
                "required": ["range"],
                "additionalProperties": false
            }
        }
    ])
}

#[cfg(test)]
mod tests {
    use super::*;

    fn server() -> McpServer {
        McpServer { session_pref: None }
    }

    #[test]
    fn initialize_handshake() {
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"test","version":"0"}}}"#)
            .unwrap();
        assert_eq!(resp["id"], 1);
        assert_eq!(resp["result"]["protocolVersion"], "2025-03-26");
        assert_eq!(resp["result"]["serverInfo"]["name"], "visigrid");
        assert!(resp["result"]["capabilities"]["tools"].is_object());
    }

    #[test]
    fn initialized_notification_gets_no_reply() {
        let mut s = server();
        assert!(s
            .handle_line(r#"{"jsonrpc":"2.0","method":"notifications/initialized"}"#)
            .is_none());
    }

    #[test]
    fn tools_list_names_and_schemas() {
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":2,"method":"tools/list"}"#)
            .unwrap();
        let tools = resp["result"]["tools"].as_array().unwrap();
        let names: Vec<&str> = tools.iter().map(|t| t["name"].as_str().unwrap()).collect();
        assert_eq!(
            names,
            vec!["list_sessions", "get_workbook", "read_range", "write_cells", "set_format"]
        );
        for t in tools {
            assert!(t["description"].as_str().unwrap().len() > 20);
            assert_eq!(t["inputSchema"]["type"], "object");
        }
    }

    #[test]
    fn unknown_method_errors_unknown_notification_ignored() {
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":3,"method":"bogus/method"}"#)
            .unwrap();
        assert_eq!(resp["error"]["code"], -32601);
        assert!(s.handle_line(r#"{"jsonrpc":"2.0","method":"bogus/notify"}"#).is_none());
    }

    #[test]
    fn unknown_tool_is_iserror_result_not_rpc_error() {
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"bogus","arguments":{}}}"#)
            .unwrap();
        assert!(resp["error"].is_null());
        assert_eq!(resp["result"]["isError"], true);
    }

    #[test]
    fn write_cells_validates_before_connecting() {
        // Bad cell refs and ambiguous entries fail without a session.
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"write_cells","arguments":{"cells":[{"cell":"NOPE!","value":"1"}]}}}"#)
            .unwrap();
        assert_eq!(resp["result"]["isError"], true);
        let text = resp["result"]["content"][0]["text"].as_str().unwrap();
        assert!(text.contains("invalid cell reference"));

        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"write_cells","arguments":{"cells":[{"cell":"A1","value":"1","formula":"=B1"}]}}}"#)
            .unwrap();
        let text = resp["result"]["content"][0]["text"].as_str().unwrap();
        assert!(text.contains("exactly one of"));
    }

    #[test]
    fn dry_run_previews_without_session() {
        // dry_run must not require a running session or token.
        let mut s = server();
        let resp = s
            .handle_line(r#"{"jsonrpc":"2.0","id":7,"method":"tools/call","params":{"name":"write_cells","arguments":{"dry_run":true,"cells":[{"cell":"a1","value":"Hello"},{"cell":"B2","formula":"SUM(A1:A2)"},{"cell":"C3","clear":true}]}}}"#)
            .unwrap();
        assert_eq!(resp["result"]["isError"], false);
        let text = resp["result"]["content"][0]["text"].as_str().unwrap();
        assert!(text.contains("\"dry_run\": true"));
        assert!(text.contains(r#"A1 = \"Hello\""#));
        assert!(text.contains("B2 = =SUM(A1:A2)"), "leading '=' must be added: {}", text);
        assert!(text.contains("C3 cleared"));
    }

    #[test]
    fn a1_range_parsing() {
        assert_eq!(parse_a1_range("A1").unwrap(), ((0, 0), (0, 0)));
        assert_eq!(parse_a1_range("B2:D10").unwrap(), ((1, 1), (9, 3)));
        assert!(parse_a1_range("D10:B2").is_err());
        assert!(parse_a1_range("junk").is_err());
    }

    #[test]
    fn cell_names_round_trip() {
        assert_eq!(cell_name(0, 0), "A1");
        assert_eq!(cell_name(9, 3), "D10");
        assert_eq!(cell_name(0, 25), "Z1");
        assert_eq!(cell_name(0, 26), "AA1");
        for (row, col) in [(0, 0), (99, 51), (65535, 255)] {
            assert_eq!(parse_cell_ref(&cell_name(row, col)), Some((row, col)));
        }
    }
}
