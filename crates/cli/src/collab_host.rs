//! `vgrid collab-host`: the engine host the Go sequencer drives over stdio.
//!
//! The contract is the vault spec "VisiGrid Collaborative Workbook Model",
//! section *Phase 2 interfaces*. One JSON object per line in, one per line
//! out, strictly in order. Every response echoes the request's `id` and
//! carries `"ok": true` or `"ok": false, "error": "…"`. A malformed line,
//! an unknown command or a failed request never ends the process.
//!
//! The host is the server's replica: it holds the canonical workbook and
//! applies every sequenced operation. It does not deduplicate envelopes
//! (the sequencer owns `client_op_id`), and it never chooses sequence
//! numbers; it only checks that the ones it is given are consecutive.
//!
//! The `op` payload is opaque to the sequencer. Here it is the envelope's
//! atomic op list (`collab::CollabOp` serde form), as a JSON array; a single
//! op object is accepted on input.
//!
//! visigrid-json v2 numbers sheets 1..n on import, but operations name
//! sheets by stable id. `snapshot` therefore adds a top-level
//! `collab_sheet_ids` array to the document (ignored by every other reader),
//! and `load` / `replace_document` restore the ids from it when present.

use std::io::{BufRead, Write};
use std::panic::{catch_unwind, AssertUnwindSafe};

use serde_json::{json, Map, Value};
use visigrid_collab::apply::{apply_ops, checksum, collab_checksum, filter_unappliable};
use visigrid_collab::op::{ops_from_json, ops_to_json, CollabOp};
use visigrid_collab::transform::{transform_lists, Order};
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::SheetLayout;

/// Engine host protocol version (spec: protocol 2).
pub const PROTOCOL: u32 = 2;

/// Largest request line accepted. Documents arrive inline in `load` and
/// `replace_document`; this bounds memory per request, and is well above
/// the 4 MiB workbook limit the API enforces for imports and MCP writes.
pub const MAX_LINE_BYTES: usize = 32 * 1024 * 1024;

/// The engine commit compiled into this binary (set by the build), so the
/// sequencer can refuse clients on a different engine.
fn engine_commit() -> &'static str {
    // Stamped by build.rs with the helper the WASM engine's build uses, so a
    // browser and this host built from the same source agree exactly.
    env!("VISIGRID_ENGINE_COMMIT")
}

#[derive(Debug, Default, Clone, PartialEq)]
pub struct Clock {
    pub now: Option<String>,
    pub tz: Option<String>,
    pub seed: Option<u64>,
    pub utc_offset_seconds: Option<i64>,
}

pub struct Host {
    wb: Option<Workbook>,
    /// File layouts by stable sheet id (filters, charts; lines live in the sheets).
    layouts: std::collections::HashMap<u64, SheetLayout>,
    active_sheet: usize,
    /// The last sequence number this replica includes.
    seq: u64,
    /// Envelopes based before a `replace_document` barrier cannot be
    /// transformed across it: they are refused and their clients resync.
    barrier: u64,
    /// After a panic mid-request the workbook may be half-updated; refuse
    /// everything but `hello` and `load` until the sequencer reloads.
    poisoned: bool,
    /// Calculation context supplied by the sequencer.
    pub clock: Clock,
    /// The bands of the last `snapshot`, by key, until the next one: the
    /// sequencer fetches each with `band` (one per line, never one huge line).
    bands: std::collections::HashMap<String, Vec<u8>>,
}

impl Default for Host {
    fn default() -> Self {
        Self::new()
    }
}

type Reply = Result<Map<String, Value>, String>;

impl Host {
    pub fn new() -> Self {
        Host {
            wb: None,
            layouts: std::collections::HashMap::new(),
            active_sheet: 0,
            seq: 0,
            barrier: 0,
            poisoned: false,
            clock: Clock::default(),
            bands: std::collections::HashMap::new(),
        }
    }

    pub fn seq(&self) -> u64 {
        self.seq
    }

    pub fn workbook(&self) -> Option<&Workbook> {
        self.wb.as_ref()
    }

    /// Handle one request line and return the response line (no newline).
    pub fn handle_line(&mut self, line: &str) -> String {
        let request: Value = match serde_json::from_str(line) {
            Ok(v) => v,
            Err(e) => return error_line(Value::Null, &format!("malformed request: {e}")),
        };
        let id = request.get("id").cloned().unwrap_or(Value::Null);
        let Some(obj) = request.as_object() else {
            return error_line(id, "a request must be a JSON object");
        };
        let cmd = obj.get("cmd").and_then(Value::as_str).unwrap_or("");
        let result = catch_unwind(AssertUnwindSafe(|| self.dispatch(cmd, obj)));
        match result {
            Ok(Ok(mut fields)) => {
                fields.insert("id".into(), id);
                fields.insert("ok".into(), Value::Bool(true));
                Value::Object(fields).to_string()
            }
            Ok(Err(message)) => error_line(id, &message),
            Err(panic) => {
                if !matches!(cmd, "hello" | "snapshot") {
                    self.poisoned = true;
                }
                let what = panic
                    .downcast_ref::<&str>()
                    .map(|s| s.to_string())
                    .or_else(|| panic.downcast_ref::<String>().cloned())
                    .unwrap_or_else(|| "unknown panic".into());
                error_line(id, &format!("engine panic during {cmd}: {what}; reload required"))
            }
        }
    }

    fn dispatch(&mut self, cmd: &str, req: &Map<String, Value>) -> Reply {
        if self.poisoned && !matches!(cmd, "hello" | "load") {
            return Err("host state is unknown after an engine panic; send load".into());
        }
        match cmd {
            "hello" => Ok(fields(json!({"engine_commit": engine_commit(), "protocol": PROTOCOL}))),
            "load" => self.load(req),
            "replay" => self.replay(req),
            "submit" => self.submit(req),
            "snapshot" => self.snapshot(),
            "band" => self.band(req),
            "read_cells" => self.read_cells(req),
            "load_band" => self.load_band(req),
            "finish_load" => self.finish_load(),
            "replace_document" => self.replace_document(req),
            "set_clock" => self.set_clock(req),
            "restore_clock" => {
                let clock = visigrid_collab::clock::parse_clock(req.get("clock").ok_or("restore_clock needs clock")?)?;
                let wb = self.wb_mut()?;
                wb.set_recalc_clock(Some(clock));
                wb.recompute_full_ordered();
                Ok(Map::new())
            },
            "" => Err("missing cmd".into()),
            other => Err(format!("unknown cmd {other:?}")),
        }
    }

    fn wb(&self) -> Result<&Workbook, String> {
        self.wb.as_ref().ok_or_else(|| "no workbook loaded; send load first".to_string())
    }

    fn wb_mut(&mut self) -> Result<&mut Workbook, String> {
        self.wb.as_mut().ok_or_else(|| "no workbook loaded; send load first".to_string())
    }

    fn load(&mut self, req: &Map<String, Value>) -> Reply {
        let doc = req.get("document").ok_or("load needs document")?;
        let seq = req
            .get("seq")
            .and_then(Value::as_u64)
            .ok_or("load needs a non-negative integer seq")?;
        let (wb, layouts, active) = import_document(doc)?;
        self.layouts = wb.sheets().iter().map(|s| s.id.0).zip(layouts).collect();
        self.wb = Some(wb);
        self.active_sheet = active;
        self.seq = seq;
        self.barrier = seq;
        self.poisoned = false;
        // `complete`: loaded editable with nothing withheld, so the content
        // protection diff found no loss. The server records it per snapshot,
        // and browsers on this engine commit skip repeating the diff.
        let complete = visigrid_io::json::loaded_complete(self.wb()?);
        Ok(fields(json!({"checksum": collab_checksum(self.wb()?), "complete": complete})))
    }

    fn replay(&mut self, req: &Map<String, Value>) -> Reply {
        let entries = req.get("ops").and_then(Value::as_array).ok_or("replay needs ops")?;
        // Validate everything before applying anything: a bad batch leaves
        // the replica where it was.
        let mut parsed = Vec::with_capacity(entries.len());
        let mut expect = self.seq + 1;
        for entry in entries {
            let seq = entry.get("seq").and_then(Value::as_u64).ok_or("replay entry needs seq")?;
            if seq != expect {
                return Err(format!("replay expected seq {expect}, got {seq}"));
            }
            let ops = ops_from_json(entry.get("op").ok_or("replay entry needs op")?)?;
            let clock = entry.get("clock").map(visigrid_collab::clock::parse_clock).transpose()?;
            parsed.push((ops, clock));
            expect += 1;
        }
        let wb = self.wb_mut()?;
        for (ops, clock) in &parsed {
            if let Some(clock) = clock {
                wb.set_recalc_clock(Some(*clock));
                wb.recompute_full_ordered();
            }
            apply_ops(wb, ops);
        }
        self.seq += parsed.len() as u64;
        Ok(fields(json!({"seq": self.seq, "checksum": collab_checksum(self.wb()?)})))
    }

    fn submit(&mut self, req: &Map<String, Value>) -> Reply {
        self.wb()?;
        let env = req.get("envelope").and_then(Value::as_object).ok_or("submit needs envelope")?;
        let base = env
            .get("base_seq")
            .and_then(Value::as_u64)
            .ok_or("envelope needs a non-negative integer base_seq")?;
        let mut ops = ops_from_json(env.get("op").ok_or("envelope needs op")?)?;
        if base > self.seq {
            return Err(format!("base_seq {base} is ahead of head {}", self.seq));
        }
        let concurrent = req.get("concurrent").and_then(Value::as_array).ok_or("submit needs concurrent")?;
        if concurrent.len() as u64 != self.seq - base {
            return Err(format!(
                "concurrent must hold every op in ({base}, {}]: expected {}, got {}",
                self.seq,
                self.seq - base,
                concurrent.len()
            ));
        }
        if base < self.barrier {
            return Ok(rejected("refused", "document_replaced"));
        }
        let original_len = ops.len();
        for (k, entry) in concurrent.iter().enumerate() {
            let seq = entry.get("seq").and_then(Value::as_u64).ok_or("concurrent entry needs seq")?;
            if seq != base + 1 + k as u64 {
                return Err(format!("concurrent seq {seq} out of order (expected {})", base + 1 + k as u64));
            }
            let theirs: Vec<CollabOp> = ops_from_json(entry.get("op").ok_or("concurrent entry needs op")?)?;
            match transform_lists(&ops, &theirs, Order::Later) {
                Ok((a, _)) => ops = a,
                Err(refusal) => return Ok(rejected("refused", &refusal.reason)),
            }
        }
        if ops.is_empty() && original_len > 0 {
            return Ok(rejected(
                "dropped",
                "every operation targeted something a concurrent edit removed",
            ));
        }
        // An op that cannot apply here (missing sheet, taken name, last
        // sheet) would sequence and broadcast a no-op that still moves
        // positions in every concurrent transform. Drop it; the rest of the
        // envelope applies. The reason is the first dropped op's.
        let (kept, reasons) = filter_unappliable(self.wb()?, &ops);
        if kept.is_empty() {
            return Ok(rejected("dropped", reasons.first().copied().unwrap_or("not_applicable")));
        }
        if !reasons.is_empty() {
            ops = kept;
        }
        let wb = self.wb_mut()?;
        // Advance volatile results only once this envelope is accepted. A
        // refused operation must not change the committed calculation state.
        if wb.recalc_clock().is_some() {
            wb.recompute_full_ordered();
        }
        apply_ops(wb, &ops);
        self.seq += 1;
        Ok(fields(json!({
            "result": "op",
            "op": ops_to_json(&ops),
            "seq": self.seq,
            "checksum": collab_checksum(self.wb()?),
        })))
    }

    /// The document at the current seq. A large sheet's cells come back as
    /// bands (`visigrid_io::json::bands`): `document` is then the manifest,
    /// `bands` lists them (key, bytes, sheet, r0, r1), and `band` returns each.
    fn snapshot(&mut self) -> Reply {
        let wb = self.wb()?;
        let (document, bands) = export_banded_document(wb, &self.layouts, self.active_sheet)?;
        let list: Vec<Value> = bands
            .iter()
            .map(|b| json!({"key": b.reference.key, "bytes": b.reference.bytes, "sheet": b.sheet, "r0": b.reference.r0, "r1": b.reference.r1}))
            .collect();
        let reply = json!({"document": document, "bands": list, "seq": self.seq, "checksum": collab_checksum(wb)});
        self.bands = bands.into_iter().map(|b| (b.reference.key, b.data)).collect();
        Ok(fields(reply))
    }

    /// The cells at `cells` ([[row, col], ...], at most 10,000) of the tab at
    /// `sheet_index`, in snapshot form (value, formula, fmt; an empty cell
    /// has no entry), with the tab's stable key and name. A server write
    /// reads the cells it touches before and after applying, instead of
    /// snapshotting the whole workbook twice.
    fn read_cells(&mut self, req: &Map<String, Value>) -> Reply {
        let wb = self.wb()?;
        let index = req.get("sheet_index").and_then(Value::as_u64).ok_or("read_cells needs sheet_index")? as usize;
        let sheet = wb.sheets().get(index).ok_or_else(|| format!("no sheet at index {index}"))?;
        let list = req.get("cells").and_then(Value::as_array).ok_or("read_cells needs cells")?;
        if list.len() > 10_000 {
            return Err("read_cells reads at most 10,000 cells".into());
        }
        let mut wanted = Vec::with_capacity(list.len());
        for rc in list {
            let pair = rc.as_array().filter(|p| p.len() == 2);
            let (Some(r), Some(c)) = (pair.and_then(|p| p[0].as_u64()), pair.and_then(|p| p[1].as_u64())) else {
                return Err("read_cells cells are [row, col] pairs".into());
            };
            wanted.push((r as usize, c as usize));
        }
        let cells = visigrid_io::json::cells_at(sheet, &wanted)?;
        Ok(fields(json!({"sheet": sheet.id.0, "sheet_name": sheet.name, "cells": cells})))
    }

    /// One band of the last snapshot, base64.
    fn band(&mut self, req: &Map<String, Value>) -> Reply {
        use base64::Engine as _;
        let key = req.get("key").and_then(Value::as_str).ok_or("band needs key")?;
        let data = self.bands.get(key).ok_or_else(|| format!("no band {key} in the last snapshot"))?;
        Ok(fields(json!({"key": key, "data": base64::engine::general_purpose::STANDARD.encode(data)})))
    }

    /// Write one band (base64 `data`, checked against `key`) into the loaded
    /// manifest; `finish_load` after the last.
    fn load_band(&mut self, req: &Map<String, Value>) -> Reply {
        use base64::Engine as _;
        let data = req.get("data").and_then(Value::as_str).ok_or("load_band needs data")?;
        let key = req.get("key").and_then(Value::as_str);
        let bytes = base64::engine::general_purpose::STANDARD.decode(data).map_err(|e| format!("band data is not base64: {e}"))?;
        let (sheet, cells) = visigrid_io::json::bands::apply(self.wb_mut()?, &bytes, key)?;
        Ok(fields(json!({"sheet": sheet, "cells": cells})))
    }

    fn finish_load(&mut self) -> Reply {
        let wb = self.wb_mut()?;
        visigrid_io::json::bands::finish(wb)?;
        Ok(fields(json!({"checksum": collab_checksum(self.wb()?)})))
    }

    fn replace_document(&mut self, req: &Map<String, Value>) -> Reply {
        self.wb()?;
        let doc = req.get("document").ok_or("replace_document needs document")?;
        // Optional: when the sequencer logs the barrier it names its seq.
        let seq = match req.get("seq") {
            None | Some(Value::Null) => None,
            Some(v) => Some(v.as_u64().ok_or("seq must be a non-negative integer")?),
        };
        if let Some(s) = seq {
            if s != self.seq + 1 {
                return Err(format!("replace_document seq must be head+1 ({}), got {s}", self.seq + 1));
            }
        }
        let (wb, layouts, active) = import_document(doc)?;
        self.layouts = wb.sheets().iter().map(|s| s.id.0).zip(layouts).collect();
        self.wb = Some(wb);
        self.active_sheet = active;
        if let Some(s) = seq {
            self.seq = s;
        }
        self.barrier = self.seq;
        let complete = visigrid_io::json::loaded_complete(self.wb()?);
        Ok(fields(json!({"seq": self.seq, "checksum": collab_checksum(self.wb()?), "complete": complete})))
    }

    fn set_clock(&mut self, req: &Map<String, Value>) -> Reply {
        let now = req.get("now").and_then(Value::as_str).map(str::to_string);
        let parsed = now.as_ref().map(|n| chrono::DateTime::parse_from_rfc3339(n)
            .map_err(|e| format!("now must be RFC 3339: {e}"))).transpose()?;
        let tz = req.get("tz").and_then(Value::as_str).map(str::to_string);
        let offset = match req.get("utc_offset_seconds") {
            Some(value) => Some(value.as_i64().filter(|v| (-86400..=86400).contains(v))
                .ok_or("utc_offset_seconds must be an integer within one day")?),
            None if tz.as_deref().is_some_and(|tz| tz != "UTC") =>
                return Err("a named timezone requires utc_offset_seconds".into()),
            None => parsed.as_ref().map(|date| i64::from(date.offset().local_minus_utc())),
        };
        let seed = match req.get("seed") {
            None | Some(Value::Null) => None,
            Some(v) => Some(v.as_u64().ok_or("seed must be a non-negative integer")?),
        };
        self.clock = Clock {
            now,
            tz,
            seed,
            utc_offset_seconds: offset,
        };
        if let Some(wb) = self.wb.as_mut() {
            wb.set_recalc_clock(Some(visigrid_engine::RecalcClock {
                now_ms: parsed.as_ref().map(|date| date.timestamp_millis()),
                utc_offset_seconds: offset,
                seed,
            }));
        }
        Ok(Map::new())
    }
}

fn fields(v: Value) -> Map<String, Value> {
    match v {
        Value::Object(m) => m,
        _ => Map::new(),
    }
}

fn rejected(result: &str, reason: &str) -> Map<String, Value> {
    fields(json!({"result": result, "reason": reason}))
}

fn error_line(id: Value, message: &str) -> String {
    json!({"id": id, "ok": false, "error": message}).to_string()
}

/// visigrid-json v2 (any version the importer reads) to a workbook, with
/// stable sheet ids restored from `collab_sheet_ids` when present.
pub fn import_document(doc: &Value) -> Result<(Workbook, Vec<SheetLayout>, usize), String> {
    if !doc.is_object() {
        return Err("document must be a visigrid-json object".into());
    }
    let text = doc.to_string();
    let (mut wb, layouts, active) = visigrid_io::json::import_any(&text)?;
    // Lines (sizes, hidden, frozen) live in the engine sheets, where
    // collaboration ops change them and checksums cover them.
    for (i, l) in layouts.iter().enumerate() {
        if let Some(s) = wb.sheet_mut(i) {
            s.layout = l.line_layout();
        }
    }
    if let Some(ids) = doc.get("collab_sheet_ids") {
        let ids: Vec<u64> = serde_json::from_value(ids.clone())
            .map_err(|_| "collab_sheet_ids must be an array of non-negative integers")?;
        if ids.len() != wb.sheets().len() {
            return Err(format!(
                "collab_sheet_ids has {} ids for {} sheets",
                ids.len(),
                wb.sheets().len()
            ));
        }
        let mut seen = std::collections::HashSet::new();
        if ids.iter().any(|id| *id == 0 || !seen.insert(*id)) {
            return Err("collab_sheet_ids must be distinct and non-zero".into());
        }
        for (i, id) in ids.iter().enumerate() {
            wb.sheet_mut(i).ok_or("sheet index out of range")?.id = SheetId(*id);
        }
        let next = ids.iter().copied().max().unwrap_or(0) + 1;
        wb.set_next_sheet_id(next.max(wb.next_sheet_id()));
    }
    Ok((wb, layouts, active))
}

/// A workbook to visigrid-json v2 plus `collab_sheet_ids`.
/// `layouts` are the file layouts by stable sheet id (their filters and
/// charts); each sheet's lines come from the engine.
pub fn export_document(wb: &Workbook, layouts: &std::collections::HashMap<u64, SheetLayout>, active: usize) -> Result<Value, String> {
    let layouts: Vec<SheetLayout> = wb
        .sheets()
        .iter()
        .map(|s| {
            let mut l = layouts.get(&s.id.0).cloned().unwrap_or_default();
            l.set_line_layout(&s.layout);
            l
        })
        .collect();
    let text = visigrid_io::json::export_workbook(wb, &layouts, active.min(wb.sheets().len().saturating_sub(1)))?;
    let mut doc: Value = serde_json::from_str(&text).map_err(|e| format!("export produced invalid JSON: {e}"))?;
    let ids: Vec<u64> = wb.sheets().iter().map(|s| s.id.0).collect();
    doc.as_object_mut()
        .ok_or("export produced a non-object document")?
        .insert("collab_sheet_ids".into(), json!(ids));
    Ok(doc)
}

/// `export_document`, with large sheets' cells moved into bands.
pub fn export_banded_document(
    wb: &Workbook,
    layouts: &std::collections::HashMap<u64, SheetLayout>,
    active: usize,
) -> Result<(Value, Vec<visigrid_io::json::bands::Band>), String> {
    let layouts: Vec<SheetLayout> = wb
        .sheets()
        .iter()
        .map(|s| {
            let mut l = layouts.get(&s.id.0).cloned().unwrap_or_default();
            l.set_line_layout(&s.layout);
            l
        })
        .collect();
    let (text, bands) = visigrid_io::json::bands::export_banded(wb, &layouts, active.min(wb.sheets().len().saturating_sub(1)))?;
    let mut doc: Value = serde_json::from_str(&text).map_err(|e| format!("export produced invalid JSON: {e}"))?;
    let ids: Vec<u64> = wb.sheets().iter().map(|s| s.id.0).collect();
    doc.as_object_mut()
        .ok_or("export produced a non-object document")?
        .insert("collab_sheet_ids".into(), json!(ids));
    Ok((doc, bands))
}

/// Read one line of at most `max` bytes (newline excluded). Returns
/// `Ok(None)` at end of input and `Err(len)` for an over-long line, which is
/// consumed through its newline so the next request still parses.
pub fn read_bounded_line<R: BufRead>(r: &mut R, max: usize) -> std::io::Result<Option<Result<String, usize>>> {
    let mut buf = Vec::new();
    let mut too_long = 0usize;
    loop {
        let available = r.fill_buf()?;
        if available.is_empty() {
            if buf.is_empty() && too_long == 0 {
                return Ok(None);
            }
            break;
        }
        let (chunk, done) = match available.iter().position(|b| *b == b'\n') {
            Some(i) => (&available[..i], Some(i + 1)),
            None => (available, None),
        };
        if too_long > 0 || buf.len() + chunk.len() > max {
            too_long += chunk.len() + buf.len();
            buf.clear();
        } else {
            buf.extend_from_slice(chunk);
        }
        let used = done.unwrap_or(available.len());
        r.consume(used);
        if done.is_some() {
            break;
        }
    }
    if too_long > 0 {
        return Ok(Some(Err(too_long)));
    }
    Ok(Some(Ok(String::from_utf8_lossy(&buf).trim_end_matches('\r').to_string())))
}

/// The protocol stream. On unix the real stdout is moved aside and fd 1 is
/// pointed at stderr, so stray prints from library code cannot corrupt it.
fn protocol_writer() -> Box<dyn Write> {
    #[cfg(unix)]
    unsafe {
        use std::os::unix::io::FromRawFd;
        let fd = libc::dup(1);
        if fd >= 0 && libc::dup2(2, 1) >= 0 {
            return Box::new(std::io::BufWriter::new(std::fs::File::from_raw_fd(fd)));
        }
    }
    Box::new(std::io::BufWriter::new(std::io::stdout()))
}

pub fn run() -> std::io::Result<()> {
    let mut out = protocol_writer();
    let stdin = std::io::stdin();
    let mut input = stdin.lock();
    let mut host = Host::new();
    while let Some(line) = read_bounded_line(&mut input, MAX_LINE_BYTES)? {
        let reply = match line {
            Ok(l) if l.trim().is_empty() => continue,
            Ok(l) => host.handle_line(&l),
            Err(len) => error_line(
                Value::Null,
                &format!("request line of {len} bytes exceeds the {MAX_LINE_BYTES}-byte limit"),
            ),
        };
        out.write_all(reply.as_bytes())?;
        out.write_all(b"\n")?;
        out.flush()?;
    }
    Ok(())
}
