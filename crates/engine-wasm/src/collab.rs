//! The browser's collaboration replica: `visigrid_collab::client::Client`
//! behind the JS API of the spec's "Phase 3 interfaces" section.
//!
//! The browser never transforms anything itself. Every local edit, server
//! frame and rebuild goes through the same client code the convergence
//! simulator verified; this layer only translates JSON and reports what
//! changed (`Effects`) so the grid can repaint.
//!
//! ```text
//! const c = new CollabClient(document, seq);  // visigrid-json v2 + the seq it includes
//! let fx = c.local([{SetCell: {...}}]);       // optimistic; queued
//! const env = c.poll_send();                  // {client_op_id, base_seq, op} or null
//! fx = c.receive(frame);                      // any WebSocket v2 frame, parsed
//! ```
//!
//! Documents load and save through `visigrid_io::json` (built without its
//! native file formats), exactly as the server's engine host does, so the
//! checksum both sides compute covers the same formats and displays.

use std::collections::HashMap;

use serde_json::{json, Value};
use uuid::Uuid;
use visigrid_collab::apply::{checksum, collab_checksum, Changes};
use visigrid_collab::client::{Client, ToClient, ToServer};
use visigrid_collab::op::{ops_from_json, ops_to_json, SheetKey};
use visigrid_collab::server::Committed;
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::SheetLayout;
use wasm_bindgen::prelude::*;

use crate::out_result;

/// A document loaded for collaboration.
struct Loaded {
    wb: Workbook,
    layouts: HashMap<SheetKey, SheetLayout>,
    active: usize,
}

/// visigrid-json v2 → workbook, honouring `collab_sheet_ids` (the stable sheet
/// keys operations name), exactly as `vgrid collab-host` loads it.
fn load(document: &Value) -> Result<Loaded, String> {
    let (mut wb, layouts, active) = visigrid_io::json::import_any(&document.to_string())?;
    if let Some(ids) = document.get("collab_sheet_ids") {
        let ids: Vec<u64> = serde_json::from_value(ids.clone()).map_err(|_| "bad collab_sheet_ids")?;
        if ids.len() != wb.sheets().len() {
            return Err("collab_sheet_ids does not match the sheet count".into());
        }
        for (i, id) in ids.iter().enumerate() {
            wb.sheet_mut(i).ok_or("sheet index")?.id = SheetId(*id);
        }
        let next = ids.iter().copied().max().unwrap_or(0) + 1;
        wb.set_next_sheet_id(next.max(wb.next_sheet_id()));
    }
    for (i, l) in layouts.iter().enumerate() {
        if let Some(s) = wb.sheet_mut(i) {
            s.layout = l.line_layout();
        }
    }
    let layouts = wb.sheets().iter().map(|s| s.id.0).zip(layouts).collect();
    Ok(Loaded { wb, layouts, active })
}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = crypto, js_name = randomUUID)]
    fn crypto_random_uuid() -> String;
}

/// A fresh client_op_id. In the browser, `crypto.randomUUID()` (no Rust RNG
/// backend needed); natively (tests), unique per process run.
fn new_op_id() -> Uuid {
    #[cfg(target_arch = "wasm32")]
    {
        Uuid::parse_str(&crypto_random_uuid()).expect("crypto.randomUUID returns a UUID")
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        use std::sync::atomic::{AtomicU64, Ordering};
        static NEXT: AtomicU64 = AtomicU64::new(1);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_nanos() as u64)
            .unwrap_or(0);
        Uuid::from_u64_pair(nanos, NEXT.fetch_add(1, Ordering::Relaxed))
    }
}

fn committed_from(v: &Value) -> Result<Committed, String> {
    Ok(Committed {
        seq: v.get("seq").and_then(Value::as_u64).ok_or("op without seq")?,
        client_op_id: uuid_field(v)?,
        actor: v.get("actor").and_then(Value::as_u64).unwrap_or(0),
        ops: ops_from_json(v.get("op").ok_or("op without op")?)?,
    })
}

fn uuid_field(v: &Value) -> Result<Uuid, String> {
    v.get("client_op_id")
        .and_then(Value::as_str)
        .and_then(|s| Uuid::parse_str(s).ok())
        .ok_or_else(|| "frame without a valid client_op_id".to_string())
}

/// The engine logic behind `CollabClient`, on plain JSON values so it can be
/// tested natively. Every method mirrors the JS method of the same name.
pub(crate) struct CollabCore {
    client: Client,
    layouts: HashMap<SheetKey, SheetLayout>,
    active: usize,
    /// A welcome that pointed at a snapshot, waiting for `load_snapshot`.
    waiting_welcome: Option<Value>,
    // Per-call outcome flags, reported in the next Effects.
    rejected: Vec<Value>,
    checksum_mismatch: bool,
    need_snapshot: Option<(String, u64)>,
    reconnect: bool,
    /// Effects for a page that reads cells through `viewport()` (the custom
    /// grid in a worker): a full repaint, or a change touching more than
    /// `LEAN_CELLS` cells, reports `full` with no cell list, so a large paste
    /// or a recalculation never ships every cell across.
    lean: bool,
    /// The client keeps a confirmed copy (a live session): bands loaded into
    /// the optimistic workbook are copied there when the load finishes.
    confirmed_copy: bool,
}

const LEAN_CELLS: usize = 2_000;

impl CollabCore {
    pub(crate) fn new(document: &Value, seq: u64) -> Result<CollabCore, String> {
        let loaded = load(document)?;
        Ok(Self::from_loaded(loaded, seq, true))
    }

    /// `confirmed`: keep the server-confirmed copy a connected replica
    /// needs. A replica that never connects (the benchmark) skips it.
    fn from_loaded(Loaded { wb, layouts, active }: Loaded, seq: u64, confirmed: bool) -> CollabCore {
        let mut client = Client::new(0);
        if confirmed {
            client.confirmed = wb.clone();
        }
        client.wb = wb;
        client.last_seen = seq;
        client.record_changes();
        client.enable_undo();
        CollabCore {
            client,
            layouts,
            active,
            waiting_welcome: None,
            rejected: Vec::new(),
            checksum_mismatch: false,
            need_snapshot: None,
            reconnect: false,
            lean: false,
            confirmed_copy: confirmed,
        }
    }

    pub(crate) fn local(&mut self, ops: &Value) -> Result<Value, String> {
        let ops = ops_from_json(ops)?;
        if !ops.is_empty() {
            self.client.local(new_op_id(), ops);
        }
        Ok(self.effects())
    }

    /// Undo this user's last change (spec §Undo): Effects plus `undo`, the
    /// outcome (`applied`, `kept_others`, `reason`: "empty" or "gone").
    pub(crate) fn undo(&mut self) -> Value {
        let outcome = self.client.undo(new_op_id());
        self.with_outcome(outcome)
    }

    pub(crate) fn redo(&mut self) -> Value {
        let outcome = self.client.redo(new_op_id());
        self.with_outcome(outcome)
    }

    fn with_outcome(&mut self, outcome: visigrid_collab::undo::UndoOutcome) -> Value {
        let mut fx = self.effects();
        fx["undo"] = serde_json::to_value(outcome).expect("outcome serializes");
        fx
    }

    pub(crate) fn poll_send(&mut self) -> Option<Value> {
        match self.client.poll_send()? {
            ToServer::Submit(env) => Some(json!({
                "client_op_id": env.client_op_id,
                "base_seq": env.base_seq,
                "op": ops_to_json(&env.ops),
            })),
            ToServer::Hello { .. } => None,
        }
    }

    pub(crate) fn receive(&mut self, frame: &Value) -> Result<Value, String> {
        let kind = frame.get("type").and_then(Value::as_str).unwrap_or("");
        match kind {
            "welcome" => {
                if let Some(url) = frame.get("snapshot_url").and_then(Value::as_str) {
                    let at = frame
                        .get("snapshot_seq")
                        .and_then(Value::as_u64)
                        .ok_or("snapshot welcome without snapshot_seq")?;
                    self.need_snapshot = Some((url.to_string(), at));
                    self.waiting_welcome = Some(frame.clone());
                } else {
                    self.finish_welcome(frame)?;
                }
            }
            "op" => {
                let c = committed_from(frame)?;
                self.check_order(c.seq)?;
                if c.seq > self.client.last_seen {
                    visigrid_collab::clock::install_frame_clock(&mut self.client, frame)?;
                    self.client.receive(ToClient::Op(c));
                }
            }
            "ack" => {
                let id = uuid_field(frame)?;
                let seq = frame.get("seq").and_then(Value::as_u64).ok_or("ack without seq")?;
                if seq == self.client.last_seen + 1 {
                    visigrid_collab::clock::install_frame_clock(&mut self.client, frame)?;
                    self.client.ack(id, seq);
                } else if seq > self.client.last_seen + 1 {
                    return Err(format!("ack seq {seq} skips ops after {}", self.client.last_seen));
                }
            }
            "rejected" => {
                let id = uuid_field(frame)?;
                let reason = frame.get("reason").and_then(Value::as_str).unwrap_or("").to_string();
                match frame.get("result").and_then(Value::as_str).unwrap_or("") {
                    "dropped" => self.client.dropped(id),
                    "refused" => self.client.receive(ToClient::Refused { client_op_id: id, reason: reason.clone() }),
                    other => return Err(format!("rejected with unknown result {other:?}")),
                }
                self.rejected.push(json!({"client_op_id": id, "reason": reason}));
            }
            "checksum" => {
                let seq = frame.get("seq").and_then(Value::as_u64).ok_or("checksum without seq")?;
                let sum = frame.get("checksum").and_then(Value::as_str).ok_or("checksum without value")?;
                // Frames are FIFO and a checksum follows the op it covers, so
                // at that moment our confirmed state is exactly at `seq`.
                // An empty checksum is a workbook too large to checksum.
                if !sum.is_empty() && seq == self.client.last_seen && collab_checksum(&self.client.confirmed) != sum {
                    self.checksum_mismatch = true;
                }
            }
            "resync" => {
                // The server replaced the document (a whole-workbook save).
                // Reconnecting gets a welcome with the snapshot to adopt.
                self.client.disconnect();
                self.reconnect = true;
            }
            // Presence is the page's business; errors and pongs carry no state.
            "presence" | "presence_leave" | "pong" | "error" => {}
            other => return Err(format!("unknown frame type {other:?}")),
        }
        Ok(self.effects())
    }

    fn finish_welcome(&mut self, frame: &Value) -> Result<(), String> {
        for o in frame.get("ops").and_then(Value::as_array).into_iter().flatten() {
            let c = committed_from(o)?;
            self.check_order(c.seq)?;
            if c.seq > self.client.last_seen {
                visigrid_collab::clock::install_frame_clock(&mut self.client, o)?;
                self.client.receive(ToClient::Op(c));
            }
        }
        let head = frame.get("seq").and_then(Value::as_u64).ok_or("welcome without seq")?;
        self.client.receive(ToClient::Welcome { head });
        Ok(())
    }

    fn check_order(&self, seq: u64) -> Result<(), String> {
        if seq > self.client.last_seen + 1 {
            return Err(format!("op seq {seq} arrived after {} (gap)", self.client.last_seen));
        }
        Ok(())
    }

    pub(crate) fn load_snapshot(&mut self, document: &Value, seq: u64) -> Result<Value, String> {
        let Loaded { wb, layouts, active } = load(document)?;
        self.layouts = layouts;
        self.active = active;
        self.client.replace_document(wb, seq);
        self.need_snapshot = None;
        if let Some(frame) = self.waiting_welcome.take() {
            self.finish_welcome(&frame)?;
        }
        Ok(self.effects())
    }

    /// One band of a large sheet (see `visigrid_io::json::bands`), written
    /// unevaluated; `finish_load` evaluates once every band is in. Effects
    /// report a full repaint so a viewport reader shows the new rows.
    pub(crate) fn load_band(&mut self, data: &[u8], key: Option<&str>) -> Result<Value, String> {
        visigrid_io::json::bands::apply(&mut self.client.wb, data, key)?;
        if let Some(ch) = self.client.changes.as_mut() {
            ch.full = true;
        }
        Ok(self.effects())
    }

    /// After the last band: dependencies, one ordered recompute, and the
    /// confirmed copy brought level.
    pub(crate) fn finish_load(&mut self) -> Result<Value, String> {
        visigrid_io::json::bands::finish(&mut self.client.wb)?;
        if self.confirmed_copy {
            self.client.confirmed = self.client.wb.clone();
        }
        if let Some(ch) = self.client.changes.as_mut() {
            ch.full = true;
        }
        Ok(self.effects())
    }

    /// `load_snapshot` from a workbook assembled elsewhere (a banded
    /// snapshot loaded into a scratch client), so pending edits rebase onto
    /// the whole sheet, never a half-loaded one.
    pub(crate) fn load_snapshot_from(&mut self, other: &CollabCore, seq: u64) -> Result<Value, String> {
        self.layouts = other.layouts.clone();
        self.active = other.active;
        self.client.replace_document(other.client.wb.clone(), seq);
        self.need_snapshot = None;
        if let Some(frame) = self.waiting_welcome.take() {
            self.finish_welcome(&frame)?;
        }
        Ok(self.effects())
    }

    /// A client built from a document fetched before the welcome learns the
    /// seq that document includes. Only before it has sent or received
    /// anything.
    pub(crate) fn set_snapshot_seq(&mut self, seq: u64) -> Result<(), String> {
        if self.client.pending_count() > 0 || self.client.last_seen != 0 {
            return Err("the snapshot seq is set once, before any operation".into());
        }
        self.client.last_seen = seq;
        Ok(())
    }

    pub(crate) fn reconnect(&mut self) {
        self.client.disconnect();
        let _ = self.client.reconnect();
    }

    pub(crate) fn last_seen(&self) -> u64 {
        self.client.last_seen
    }

    pub(crate) fn pending(&self) -> usize {
        self.client.pending_count()
    }

    pub(crate) fn snapshot(&self) -> Result<Value, String> {
        let wb = &self.client.wb;
        let layouts: Vec<SheetLayout> = wb
            .sheets()
            .iter()
            .map(|s| {
                let mut l = self.layouts.get(&s.id.0).cloned().unwrap_or_default();
                l.set_line_layout(&s.layout);
                l
            })
            .collect();
        let text = visigrid_io::json::export_workbook(wb, &layouts, self.active.min(wb.sheets().len().saturating_sub(1)))?;
        let mut doc: Value = serde_json::from_str(&text).map_err(|e| e.to_string())?;
        doc["collab_sheet_ids"] = json!(wb.sheets().iter().map(|s| s.id.0).collect::<Vec<_>>());
        Ok(doc)
    }

    /// A sheet key no sheet has, in [2^40, 2^52) like the grid's own fresh
    /// keys (exact in JS numbers, and unlikely to meet another client's).
    fn fresh_key(&self, taken: &[SheetKey]) -> SheetKey {
        loop {
            let k = (1u64 << 40) + (new_op_id().as_u64_pair().1 % ((1u64 << 52) - (1u64 << 40)));
            if !taken.contains(&k) && self.client.wb.idx_for_sheet_id(SheetId(k)).is_none() {
                return k;
            }
        }
    }

    /// `name`, or "name (2)", "name (3)", … the first no sheet has (and no
    /// name in `also`).
    fn free_name(&self, name: &str, also: &[String]) -> String {
        let norm = visigrid_engine::sheet::normalize_sheet_name;
        let taken = |n: &str| {
            self.client.wb.sheets().iter().any(|s| norm(&s.name) == norm(n)) || also.iter().any(|a| norm(a) == norm(n))
        };
        if !taken(name) {
            return name.to_string();
        }
        (2..).map(|i| format!("{name} ({i})")).find(|n| !taken(n)).expect("a free name")
    }

    /// The ops that add every sheet of `document` (visigrid-json, as the
    /// server's converter returns it) after the existing tabs, as one change:
    /// `{ops, sheets: [{key, name}], dropped: [line, …]}`. `dropped` says, in
    /// plain language, what the file has that the workbook cannot take, so
    /// the page can ask before importing.
    pub(crate) fn import_ops(&self, document: &Value) -> Result<Value, String> {
        let (wb, layouts, _) = visigrid_io::json::import_any(&document.to_string())?;
        wb.ensure_writable()?;
        let mut ops = Vec::new();
        let mut sheets = Vec::new();
        let mut keys = Vec::new();
        let mut names: Vec<String> = Vec::new();
        let mut dropped = Vec::new();
        let base = self.client.wb.sheets().len();
        for (i, src) in wb.sheets().iter().enumerate() {
            let mut src = src.clone();
            if let Some(l) = layouts.get(i) {
                src.layout = l.line_layout();
            }
            let key = self.fresh_key(&keys);
            let name = self.free_name(&src.name, &names);
            if name != src.name {
                dropped.push(format!(
                    "{}: renamed \"{name}\" because this workbook already has a sheet by that name; the file's formulas that name \"{}\" will read this workbook's sheet",
                    src.name, src.name
                ));
            }
            let (sheet_ops, lost) = visigrid_collab::copy::sheet_ops(&src, key, &name, base + i);
            if let Some(l) = layouts.get(i) {
                let charts = l.charts.as_ref().and_then(|c| c.as_array()).map_or(0, |c| c.len());
                if charts > 0 {
                    dropped.push(format!("{}: {charts} chart{}", src.name, if charts == 1 { "" } else { "s" }));
                }
                if l.filter.is_some() {
                    dropped.push(format!("{}: its filter and sort settings (the rows are kept)", src.name));
                }
            }
            ops.extend(sheet_ops);
            dropped.extend(lost);
            sheets.push(json!({"key": key, "name": name}));
            keys.push(key);
            names.push(name);
        }
        let named = wb.list_named_ranges().len();
        if named > 0 {
            dropped.push(format!("{named} named range{} (formulas that use {} show #NAME?)", if named == 1 { "" } else { "s" }, if named == 1 { "it" } else { "them" }));
        }
        Ok(json!({"ops": visigrid_collab::op::ops_to_json(&ops), "sheets": sheets, "dropped": dropped}))
    }

    /// The ops that copy sheet `sheet` into a new tab right after it, named
    /// "Name (2)" (Excel's convention): `{ops, key, name, dropped}`.
    pub(crate) fn duplicate_ops(&self, sheet: SheetKey) -> Result<Value, String> {
        let wb = &self.client.wb;
        let idx = wb.idx_for_sheet_id(SheetId(sheet)).ok_or("no such sheet")?;
        let src = &wb.sheets()[idx];
        let key = self.fresh_key(&[]);
        let base = src.name.rsplit_once(" (").filter(|(_, n)| n.trim_end_matches(')').parse::<u32>().is_ok()).map_or(src.name.as_str(), |(b, _)| b);
        let name = self.free_name(base, &[]);
        let (ops, dropped) = visigrid_collab::copy::sheet_ops(src, key, &name, idx + 1);
        Ok(json!({"ops": visigrid_collab::op::ops_to_json(&ops), "key": key, "name": name, "dropped": dropped}))
    }

    pub(crate) fn checksum(&self) -> String {
        collab_checksum(&self.client.confirmed)
    }

    pub(crate) fn set_clock(&mut self, now_ms: Option<f64>, utc_offset_minutes: Option<i32>, seed: Option<f64>) {
        let clock = visigrid_engine::RecalcClock {
            now_ms: now_ms.filter(|v| v.is_finite()).map(|v| v as i64),
            utc_offset_seconds: utc_offset_minutes.map(|m| i64::from(m) * 60),
            seed: seed.filter(|v| v.is_finite() && *v >= 0.0).map(|v| v as u64),
        };
        let any = clock.now_ms.is_some() || clock.utc_offset_seconds.is_some() || clock.seed.is_some();
        let clock = any.then_some(clock);
        self.client.wb.set_recalc_clock(clock.clone());
        self.client.confirmed.set_recalc_clock(clock);
    }

    /// The optimistic display of one cell, by stable sheet key.
    pub(crate) fn display(&self, sheet: SheetKey, row: usize, col: usize) -> Option<String> {
        let wb = &self.client.wb;
        let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
        Some(wb.sheets()[idx].get_formatted_display(row, col))
    }

    /// Recalculate every formula in the optimistic workbook.
    pub(crate) fn recalc_all(&mut self) -> Value {
        self.client.wb.recompute_full_ordered();
        if let Some(ch) = self.client.changes.as_mut() {
            ch.full = true;
        }
        self.effects()
    }

    /// What the user typed into a cell (formula text with `=`, or the
    /// literal input), for an editor to start from.
    pub(crate) fn raw(&self, sheet: SheetKey, row: usize, col: usize) -> Option<String> {
        let wb = &self.client.wb;
        let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
        Some(wb.sheets()[idx].get_raw(row, col))
    }

    /// The sheets, in tab order: `[{key, name, index}]`.
    pub(crate) fn sheets(&self) -> Value {
        let wb = &self.client.wb;
        Value::Array(
            wb.sheets().iter().enumerate().map(|(index, s)| json!({"key": s.id.0, "name": s.name, "index": index})).collect(),
        )
    }

    /// A sheet's layout for drawing: column widths and row heights that
    /// differ from the defaults, frozen and hidden rows and columns, and the
    /// extent of its data (`rows`, `cols`: one past the last used row and
    /// column).
    pub(crate) fn layout(&self, sheet: SheetKey) -> Option<Value> {
        let wb = &self.client.wb;
        let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
        let s = &wb.sheets()[idx];
        let (max_row, max_col) = s.data_extent();
        let used = !s.get_raw(max_row, max_col).is_empty() || max_row > 0 || max_col > 0;
        let l = &s.layout;
        let merges: Vec<Value> = s
            .merged_regions
            .iter()
            .map(|m| json!({"r0": m.start.0, "c0": m.start.1, "r1": m.end.0, "c1": m.end.1}))
            .collect();
        Some(json!({
            "col_widths": l.col_widths,
            "row_heights": l.row_heights,
            "frozen_rows": l.frozen_rows,
            "frozen_cols": l.frozen_cols,
            "hidden_rows": l.hidden_rows,
            "hidden_cols": l.hidden_cols,
            "merges": merges,
            "rows": if used { max_row + 1 } else { 0 },
            "cols": if used { max_col + 1 } else { 0 },
        }))
    }

    /// What the grid draws for the cells of a rectangle (inclusive), read
    /// straight from the engine: the display string (number formats
    /// applied), the computed value's kind (`n` number, `t` text, `b`
    /// boolean, `e` error: general alignment depends on it) and an index
    /// into `formats` (`-1`: the default format), and `num`, a number's value
    /// (null otherwise) for selection statistics. Only cells with content
    /// or a format are listed, as parallel arrays. `row_formats` and
    /// `col_formats` carry whole-row and whole-column formats in range.
    pub(crate) fn viewport(&self, sheet: SheetKey, r0: usize, c0: usize, r1: usize, c1: usize) -> Option<Value> {
        use visigrid_engine::cell::CellFormat;
        use visigrid_engine::formula::eval::Value as V;
        let wb = &self.client.wb;
        let idx = wb.idx_for_sheet_id(SheetId(sheet))?;
        let s = &wb.sheets()[idx];
        let default = CellFormat::default();
        let mut formats: Vec<Value> = Vec::new();
        let mut seen: HashMap<String, usize> = HashMap::new();
        let mut format_index = |f: &CellFormat| -> i64 {
            if *f == default {
                return -1;
            }
            let props = serde_json::to_value(visigrid_collab::undo::props_of(f, None)).expect("props serialize");
            let key = props.to_string();
            let next = formats.len();
            let at = *seen.entry(key).or_insert(next);
            if at == next {
                formats.push(props);
            }
            at as i64
        };
        let mut coords = s.cells_in_range(r0, r1, c0, c1);
        coords.extend(s.spill_receiver_coords().filter(|&(r, c)| r >= r0 && r <= r1 && c >= c0 && c <= c1));
        coords.sort_unstable();
        coords.dedup();
        let (mut rows, mut cols, mut text, mut kind, mut fmt, mut num, mut raw) =
            (Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new());
        for (r, c) in coords {
            let shown = s.get_formatted_display(r, c);
            let f = format_index(&s.get_format(r, c));
            if shown.is_empty() && f < 0 {
                continue;
            }
            let (k, n) = match s.get_computed_value(r, c) {
                V::Number(n) => ("n", Some(n)),
                V::Boolean(_) => ("b", None),
                V::Error(_) => ("e", None),
                _ => ("t", None),
            };
            rows.push(r);
            cols.push(c);
            text.push(shown);
            kind.push(k);
            fmt.push(f);
            num.push(n.filter(|n| n.is_finite()));
            raw.push(s.get_raw(r, c));
        }
        let mut row_formats = serde_json::Map::new();
        for (r, f) in &s.row_formats {
            if *r >= r0 && *r <= r1 {
                let i = format_index(f);
                if i >= 0 {
                    row_formats.insert(r.to_string(), json!(i));
                }
            }
        }
        let mut col_formats = serde_json::Map::new();
        for (c, f) in &s.col_formats {
            if *c >= c0 && *c <= c1 {
                let i = format_index(f);
                if i >= 0 {
                    col_formats.insert(c.to_string(), json!(i));
                }
            }
        }
        Some(json!({
            "rows": rows, "cols": cols, "text": text, "kind": kind, "fmt": fmt, "num": num, "raw": raw,
            "formats": formats, "row_formats": row_formats, "col_formats": col_formats,
        }))
    }

    /// A large generated workbook for the grid benchmark: `rows` × `cols`
    /// of mixed text, integers, decimals with a number format, dates and
    /// percentages, a bold header row and some coloured cells. Values are
    /// written without recalculation (there are no formulas).
    pub(crate) fn synthetic(rows: usize, cols: usize) -> CollabCore {
        Self::synthetic_with(rows, cols, false)
    }

    /// `synthetic`, and with `formulas` every tenth column (J, T, …) is a
    /// formula over two number columns of its row, so a full recalculation
    /// has real work (100,000 x 50: 500,000 formulas).
    pub(crate) fn synthetic_with(rows: usize, cols: usize, formulas: bool) -> CollabCore {
        use visigrid_engine::cell::{CellFormat, NumberFormat};
        let mut wb = Workbook::new();
        let key = wb.sheets()[0].id.0;
        if let Some(s) = wb.sheet_mut(0) {
            let header = CellFormat { bold: true, background_color: Some([244, 242, 237, 255]), ..Default::default() };
            let money = CellFormat { number_format: NumberFormat::Custom("#,##0.00".into()), ..Default::default() };
            let pct = CellFormat { number_format: NumberFormat::Custom("0.0%".into()), ..Default::default() };
            let flagged = CellFormat { font_color: Some([180, 35, 24, 255]), italic: true, ..Default::default() };
            for c in 0..cols {
                s.set_value(0, c, &format!("Column {}", c + 1));
                s.set_format(0, c, header.clone());
            }
            for r in 1..rows {
                for c in 0..cols {
                    if formulas && c % 10 == 9 {
                        s.set_value(r, c, &format!("={}{}*2+{}{}", col_letters(c - 8), r + 1, col_letters(c - 7), r + 1));
                        s.set_format(r, c, money.clone());
                        continue;
                    }
                    match c % 5 {
                        0 => s.set_value(r, c, &format!("Item {r}-{c}")),
                        1 => s.set_value(r, c, &((r * 37 + c * 11) % 100_000).to_string()),
                        2 => {
                            s.set_value(r, c, &format!("{}.{:02}", (r * 13 + c) % 9_999, (r + c) % 100));
                            s.set_format(r, c, money.clone());
                        }
                        3 => {
                            s.set_value(r, c, &format!("0.{:03}", (r * 7 + c) % 1000));
                            s.set_format(r, c, pct.clone());
                        }
                        _ => {
                            s.set_value(r, c, if (r + c) % 9 == 0 { "TRUE" } else { "pending" });
                            if (r + c) % 9 == 0 {
                                s.set_format(r, c, flagged.clone());
                            }
                        }
                    }
                }
            }
        }
        if let Some(s) = wb.sheet_mut(0) {
            s.layout.frozen_rows = 1;
        }
        if formulas {
            wb.rebuild_dep_graph();
            wb.recompute_full_ordered();
        }
        let layouts = HashMap::from([(key, SheetLayout { frozen_rows: 1, ..Default::default() })]);
        Self::from_loaded(Loaded { wb, layouts, active: 0 }, 0, false)
    }

    /// What changed since the last call, for the page to repaint.
    fn effects(&mut self) -> Value {
        let ch: Changes = self.client.take_changes();
        let wb = &self.client.wb;
        let mut cells = Vec::new();
        let full = ch.full || (self.lean && ch.cells.len() > LEAN_CELLS);
        if full && self.lean {
            // The page re-reads what it shows.
        } else if ch.full {
            // A full repaint (join, resync, refusal rebuild) must carry the
            // engine's display for every cell, or the page shows raw values.
            // The store is sparse, so this is proportional to populated cells.
            for (idx, sheet) in wb.sheets().iter().enumerate() {
                let mut coords: Vec<(usize, usize)> = sheet
                    .cells_iter()
                    .map(|(pos, _)| pos)
                    .chain(sheet.spill_receiver_coords())
                    .collect();
                coords.sort_unstable();
                coords.dedup();
                for (row, col) in coords {
                    if sheet.get_raw(row, col).is_empty()
                        && matches!(sheet.get_computed_value(row, col), visigrid_engine::formula::eval::Value::Empty)
                    {
                        continue; // format-only cell: its style comes from the snapshot
                    }
                    cells.push(cell_json(idx, sheet, row, col));
                }
            }
        } else {
            let mut seen = std::collections::HashSet::new();
            for (key, row, col) in &ch.cells {
                if !seen.insert((*key, *row, *col)) {
                    continue;
                }
                let Some(idx) = wb.idx_for_sheet_id(SheetId(*key)) else { continue };
                cells.push(cell_json(idx, &wb.sheets()[idx], *row, *col));
            }
        }
        let sheets: Vec<Value> = if full || ch.sheets {
            wb.sheets()
                .iter()
                .enumerate()
                .map(|(index, s)| json!({"key": s.id.0, "name": s.name, "index": index}))
                .collect()
        } else {
            Vec::new()
        };
        let (snapshot_url, snapshot_seq) = match &self.need_snapshot {
            Some((url, seq)) => (json!(url), json!(seq)),
            None => (Value::Null, Value::Null),
        };
        // Format ops in apply order: the rectangle and only the properties the
        // op set (value) or cleared (null), for incremental restyling.
        let formats: Vec<Value> = ch
            .formats
            .iter()
            .map(|(sheet, rect, props)| json!({"sheet": sheet, "rect": rect, "props": props}))
            .collect();
        let out = json!({
            "full": full,
            "cells": cells,
            "formats": formats,
            "sheets": sheets,
            "structural": ops_to_json(&ch.structural),
            "rejected": std::mem::take(&mut self.rejected),
            "checksum_mismatch": self.checksum_mismatch,
            "need_snapshot": self.need_snapshot.is_some(),
            "snapshot_url": snapshot_url,
            "snapshot_seq": snapshot_seq,
            "reconnect": self.reconnect,
            "layout": ch.layout,
            "can_undo": !self.client.undo_stack.is_empty(),
            "can_redo": !self.client.redo_stack.is_empty(),
        });
        self.checksum_mismatch = false;
        self.reconnect = false;
        out
    }
}

/// Column letters: 0 → A, 26 → AA.
fn col_letters(c: usize) -> String {
    let mut s = String::new();
    let mut n = c + 1;
    while n > 0 {
        let m = (n - 1) % 26;
        s.insert(0, (b'A' + m as u8) as char);
        n = (n - 1) / 26;
    }
    s
}

fn js_err(e: String) -> JsValue {
    JsValue::from_str(&e)
}

fn from_js(v: JsValue) -> Result<Value, JsValue> {
    serde_wasm_bindgen::from_value(v).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn to_js(v: &Value) -> Result<JsValue, JsValue> {
    // Plain objects (not Maps) so the page reads `fx.cells` directly.
    let ser = serde_wasm_bindgen::Serializer::json_compatible();
    serde::Serialize::serialize(v, &ser).map_err(|e| JsValue::from_str(&e.to_string()))
}

/// A collaboration replica of one workbook. See the module docs.
#[wasm_bindgen]
pub struct CollabClient {
    core: CollabCore,
}

#[wasm_bindgen]
impl CollabClient {
    /// `document`: visigrid-json v2 (with `collab_sheet_ids`); `seq`: the
    /// last sequenced operation it includes.
    #[wasm_bindgen(constructor)]
    pub fn new(document: JsValue, seq: f64) -> Result<CollabClient, JsValue> {
        console_error_panic_hook::set_once();
        let doc = from_js(document)?;
        Ok(CollabClient { core: CollabCore::new(&doc, seq as u64).map_err(js_err)? })
    }

    /// A user edit: a CollabOp list (or one op). Applied optimistically and
    /// queued for `poll_send`.
    pub fn local(&mut self, ops: JsValue) -> Result<JsValue, JsValue> {
        let ops = from_js(ops)?;
        to_js(&self.core.local(&ops).map_err(js_err)?)
    }

    /// Undo this user's most recent change as a new op (never another
    /// user's): Effects plus `undo: {applied, kept_others, reason}`.
    pub fn undo(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.core.undo())
    }

    /// Redo the most recently undone change: Effects plus `undo`.
    pub fn redo(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.core.redo())
    }

    /// The next envelope to send as `{type: "op", envelope}`, or null.
    pub fn poll_send(&mut self) -> Result<JsValue, JsValue> {
        match self.core.poll_send() {
            Some(env) => to_js(&env),
            None => Ok(JsValue::NULL),
        }
    }

    /// Any server frame, parsed.
    pub fn receive(&mut self, frame: JsValue) -> Result<JsValue, JsValue> {
        let frame = from_js(frame)?;
        to_js(&self.core.receive(&frame).map_err(js_err)?)
    }

    /// After a welcome with `need_snapshot`: the fetched document and its seq.
    pub fn load_snapshot(&mut self, document: JsValue, seq: f64) -> Result<JsValue, JsValue> {
        let doc = from_js(document)?;
        to_js(&self.core.load_snapshot(&doc, seq as u64).map_err(js_err)?)
    }

    /// Write one band of a large sheet (bytes as stored; `key`, when given,
    /// is checked against them). Call `finish_load` after the last one.
    pub fn load_band(&mut self, data: &[u8], key: Option<String>) -> Result<JsValue, JsValue> {
        to_js(&self.core.load_band(data, key.as_deref()).map_err(js_err)?)
    }

    /// Evaluate the workbook once every band is in (full-repaint Effects).
    pub fn finish_load(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.core.finish_load().map_err(js_err)?)
    }

    /// `load_snapshot` taking the workbook of another client (a banded
    /// snapshot assembled in a scratch client).
    pub fn load_snapshot_from(&mut self, other: &CollabClient, seq: f64) -> Result<JsValue, JsValue> {
        to_js(&self.core.load_snapshot_from(&other.core, seq as u64).map_err(js_err)?)
    }

    /// The seq of the document this client was built from, when it was
    /// fetched before the welcome said (see `CollabCore::set_snapshot_seq`).
    pub fn set_snapshot_seq(&mut self, seq: f64) -> Result<(), JsValue> {
        self.core.set_snapshot_seq(seq as u64).map_err(js_err)
    }

    /// The inspector's view of a cell: its formula, direct precedents (cells
    /// and ranges as written) and dependents, and with `full` the whole
    /// chain both ways (see `crate::inspect`). Undefined for an unknown sheet.
    pub fn inspect(&self, sheet: f64, row: usize, col: usize, full: bool) -> Result<JsValue, JsValue> {
        to_js(&crate::inspect::inspect(&self.core.client.wb, sheet as u64, row, col, full).unwrap_or(Value::Null))
    }

    /// What typing `raw` into the cell would change downstream (before and
    /// after for every dependent whose value changes), without applying it.
    pub fn what_if(&self, sheet: f64, row: usize, col: usize, raw: &str) -> Result<JsValue, JsValue> {
        to_js(&crate::inspect::what_if(&self.core.client.wb, sheet as u64, row, col, raw).unwrap_or(Value::Null))
    }

    /// The socket dropped and is reconnecting: send `hello` with
    /// `last_seen()` next; pending envelopes are resent after the welcome.
    pub fn reconnect(&mut self) {
        self.core.reconnect();
    }

    pub fn last_seen(&self) -> f64 {
        self.core.last_seen() as f64
    }

    pub fn pending(&self) -> usize {
        self.core.pending()
    }

    /// The ops that import `document` (visigrid-json) as new tabs, and what
    /// will not survive: `{ops, sheets, dropped}`. Nothing is applied.
    pub fn import_ops(&self, document: JsValue) -> Result<JsValue, JsValue> {
        to_js(&self.core.import_ops(&from_js(document)?).map_err(js_err)?)
    }

    /// The ops that duplicate a sheet into the tab after it:
    /// `{ops, key, name, dropped}`. Nothing is applied.
    pub fn duplicate_ops(&self, sheet: f64) -> Result<JsValue, JsValue> {
        to_js(&self.core.duplicate_ops(sheet as SheetKey).map_err(js_err)?)
    }

    /// The optimistic document (visigrid-json v2 with `collab_sheet_ids`).
    pub fn snapshot(&self) -> Result<JsValue, JsValue> {
        to_js(&self.core.snapshot().map_err(js_err)?)
    }

    /// What the engine shows in a cell (number formats applied), or
    /// `undefined` for an unknown sheet key.
    pub fn display(&self, sheet: f64, row: usize, col: usize) -> Option<String> {
        self.core.display(sheet as SheetKey, row, col)
    }

    /// Recalculate every formula (the benchmark's "full recalculation"):
    /// Effects with `full`.
    pub fn recalc_all(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.core.recalc_all())
    }

    /// A cell's raw input (formula text or literal), or `undefined` for an
    /// unknown sheet key.
    pub fn raw(&self, sheet: f64, row: usize, col: usize) -> Option<String> {
        self.core.raw(sheet as SheetKey, row, col)
    }

    /// The sheets in tab order: `[{key, name, index}]`.
    pub fn sheets(&self) -> Result<JsValue, JsValue> {
        to_js(&self.core.sheets())
    }

    /// A sheet's layout for drawing (widths, heights, frozen, hidden, extent),
    /// or `undefined` for an unknown sheet key.
    pub fn layout(&self, sheet: f64) -> Result<JsValue, JsValue> {
        match self.core.layout(sheet as SheetKey) {
            Some(v) => to_js(&v),
            None => Ok(JsValue::UNDEFINED),
        }
    }

    /// The cells of a rectangle as the grid draws them: parallel arrays
    /// `rows`, `cols`, `text` (the engine's display), `kind`, `fmt` (an index
    /// into `formats`, `-1` default), `num` (a number's value, else null),
    /// plus `row_formats` / `col_formats`.
    pub fn viewport(&self, sheet: f64, r0: usize, c0: usize, r1: usize, c1: usize) -> Result<JsValue, JsValue> {
        match self.core.viewport(sheet as SheetKey, r0, c0, r1, c1) {
            Some(v) => to_js(&v),
            None => Ok(JsValue::UNDEFINED),
        }
    }

    /// Benchmark only: a generated `rows` × `cols` workbook (no socket, no
    /// server). See `CollabCore::synthetic`.
    pub fn synthetic(rows: usize, cols: usize) -> CollabClient {
        console_error_panic_hook::set_once();
        CollabClient { core: CollabCore::synthetic(rows, cols) }
    }

    /// `synthetic` with a formula in every tenth column (see
    /// `CollabCore::synthetic_with`).
    pub fn synthetic_formulas(rows: usize, cols: usize) -> CollabClient {
        console_error_panic_hook::set_once();
        CollabClient { core: CollabCore::synthetic_with(rows, cols, true) }
    }

    /// Lean Effects for a page that reads through `viewport()`: a full
    /// repaint or a change over 2,000 cells reports `full` and no cells.
    pub fn set_lean(&mut self, lean: bool) {
        self.core.lean = lean;
    }

    /// Fingerprint of the confirmed state, as the server's checksum frames.
    pub fn checksum(&self) -> String {
        self.core.checksum()
    }

    /// The clock and seed volatile functions read (`undefined` = machine's).
    pub fn set_clock(&mut self, now_ms: Option<f64>, utc_offset_minutes: Option<i32>, seed: Option<f64>) {
        self.core.set_clock(now_ms, utc_offset_minutes, seed);
    }
}

/// One repainted cell: raw input, computed value or error, and what the
/// desktop shows (number formats applied).
fn cell_json(idx: usize, sheet: &visigrid_engine::sheet::Sheet, row: usize, col: usize) -> Value {
    let r = out_result(idx, sheet, row, col);
    json!({
        "sheet": sheet.id.0,
        "row": row,
        "col": col,
        "raw": sheet.get_raw(row, col),
        "value": r.value,
        "error": r.error,
        "display": sheet.get_formatted_display(row, col),
    })
}

#[cfg(test)]
mod tests {
    #[test]
    fn a_banded_document_loads_band_by_band_to_the_same_workbook() {
        use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
        let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
        sheet.set_name("Sheet1");
        for r in 0..140_000 {
            sheet.set_value_deferred(r, 0, &(r % 101).to_string());
            sheet.set_value_deferred(r, 1, &format!("row {r}"));
        }
        sheet.set_value_deferred(0, 2, "=SUM(A1:A140000)");
        let mut wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet], 0);
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        let layouts = vec![SheetLayout::default()];
        let inline: Value = serde_json::from_str(&visigrid_io::json::export_workbook(&wb, &layouts, 0).unwrap()).unwrap();
        let (manifest, bands) = visigrid_io::json::bands::export_banded(&wb, &layouts, 0).unwrap();
        assert_eq!(bands.len(), 9);
        let whole = CollabCore::new(&inline, 4).unwrap();
        let mut banded = CollabCore::new(&serde_json::from_str(&manifest).unwrap(), 4).unwrap();
        banded.lean = true;
        for b in &bands {
            let fx = banded.load_band(&b.data, Some(&b.reference.key)).unwrap();
            assert_eq!(fx["full"], json!(true), "each band repaints");
        }
        // Readable mid-load (values; formulas evaluate at the end).
        assert_eq!(banded.display(1, 139_999, 1).as_deref(), Some("row 139999"));
        banded.finish_load().unwrap();
        assert_eq!(banded.display(1, 0, 2), whole.display(1, 0, 2));
        assert_eq!(checksum(&banded.client.confirmed), checksum(&whole.client.confirmed), "the confirmed copy matches the whole document");
        assert_eq!(banded.checksum(), "", "too large to checksum in collaboration");
        assert!(banded.load_band(&bands[0].data, Some(&bands[1].reference.key)).is_err());
    }

    #[test]
    fn lean_effects_report_large_changes_as_full_and_formulas_recalculate() {
        let mut c = CollabCore::synthetic_with(3000, 20, true);
        let key = c.client.wb.sheets()[0].id.0;
        // J2 = B2*2 + C2.
        let shown = c.display(key, 1, 9).unwrap();
        let b: f64 = c.display(key, 1, 1).unwrap().replace(',', "").parse().unwrap();
        let cc: f64 = c.display(key, 1, 2).unwrap().replace(',', "").parse().unwrap();
        assert_eq!(shown.replace(',', "").parse::<f64>().unwrap(), b * 2.0 + cc);
        c.lean = true;
        let paste: Vec<Value> = (0..3000).map(|r| json!({"SetCell": {"sheet": key, "sheet_name": "Sheet1", "row": r, "col": 1, "content": {"Value": "1"}}})).collect();
        let fx = c.local(&Value::Array(paste)).unwrap();
        assert_eq!((fx["full"].clone(), fx["cells"].as_array().unwrap().len()), (json!(true), 0));
        assert_eq!(c.display(key, 1, 9).unwrap().replace(',', "").parse::<f64>().unwrap(), 2.0 + cc, "the dependent recalculated");
        let fx = c.recalc_all();
        assert_eq!((fx["full"].clone(), fx["cells"].as_array().unwrap().len()), (json!(true), 0));
        let one = json!([{"SetCell": {"sheet": key, "sheet_name": "Sheet1", "row": 5, "col": 0, "content": {"Value": "x"}}}]);
        let fx = c.local(&one).unwrap();
        assert_eq!(fx["full"], json!(false), "a small edit still lists its cells");
        assert!(!fx["cells"].as_array().unwrap().is_empty());
    }

    #[test]
    #[ignore]
    fn timing_probe() {
        let t = std::time::Instant::now();
        let mut c = CollabCore::synthetic_with(100_000, 50, true);
        c.lean = true;
        eprintln!("build {:?}", t.elapsed());
        let key = c.client.wb.sheets()[0].id.0;
        let ops: Vec<Value> = (0..2000).flat_map(|r| (0..50).map(move |cc| json!({"SetCell": {"sheet": key, "sheet_name": "Sheet1", "row": r, "col": cc, "content": {"Value": format!("{}", r * 50 + cc)}}}))).collect();
        let t = std::time::Instant::now();
        let fx = c.local(&Value::Array(ops)).unwrap();
        eprintln!("paste {:?} cells {}", t.elapsed(), fx["cells"].as_array().map(|a| a.len()).unwrap_or(0));
        let t = std::time::Instant::now();
        let fx = c.recalc_all();
        eprintln!("recalc {:?} cells {}", t.elapsed(), fx["cells"].as_array().map(|a| a.len()).unwrap_or(0));
    }

    use super::*;
    use visigrid_collab::op::Envelope;
    use visigrid_collab::server::{Server, Submitted};

    const DOC: &str = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Sheet1","cells":[{"row":0,"col":0,"value":1}]}],"collab_sheet_ids":[1]}"#;

    fn doc() -> Value {
        serde_json::from_str(DOC).unwrap()
    }

    fn set(row: usize, col: usize, raw: &str) -> Value {
        let content = if raw.starts_with('=') { json!({"Formula": raw}) } else { json!({"Value": raw}) };
        json!([{"SetCell": {"sheet": 1, "sheet_name": "Sheet1", "row": row, "col": col, "content": content}}])
    }

    /// The Go room, in miniature: the collab crate's sequencer plus the
    /// protocol v2 frames the real server sends. Each client gets FIFO frames.
    struct Room {
        server: Server,
        inbox: Vec<Vec<Value>>,
    }

    impl Room {
        fn new(clients: usize) -> Room {
            let mut server = Server::new();
            server.wb = load(&doc()).unwrap().wb;
            Room { server, inbox: vec![Vec::new(); clients] }
        }

        /// Submit whatever client `i` has to send; queue the resulting frames.
        fn pump_sends(&mut self, i: usize, c: &mut CollabCore) {
            while let Some(env) = c.poll_send() {
                let env = Envelope {
                    client_op_id: Uuid::parse_str(env["client_op_id"].as_str().unwrap()).unwrap(),
                    base_seq: env["base_seq"].as_u64().unwrap(),
                    actor: i as u64 + 1,
                    ops: ops_from_json(&env["op"]).unwrap(),
                };
                match self.server.submit(&env) {
                    Submitted::Committed(c) => {
                        self.inbox[i].push(json!({"type": "ack", "client_op_id": c.client_op_id, "seq": c.seq}));
                        let op = json!({"type": "op", "seq": c.seq, "actor": c.actor,
                            "client_op_id": c.client_op_id, "op": ops_to_json(&c.ops)});
                        let sum = json!({"type": "checksum", "seq": c.seq, "checksum": checksum(&self.server.wb)});
                        for inbox in &mut self.inbox {
                            inbox.push(op.clone());
                            inbox.push(sum.clone());
                        }
                    }
                    Submitted::Refused { client_op_id, reason } => {
                        self.inbox[i].push(json!({"type": "rejected", "client_op_id": client_op_id,
                            "result": "refused", "reason": reason}));
                    }
                    Submitted::Duplicate => {}
                }
            }
        }

        /// Deliver everything queued for client `i`; return the Effects seen.
        fn deliver(&mut self, i: usize, c: &mut CollabCore) -> Vec<Value> {
            std::mem::take(&mut self.inbox[i]).iter().map(|f| c.receive(f).unwrap()).collect()
        }

        fn settle(&mut self, clients: &mut [CollabCore]) {
            for _ in 0..20 {
                let mut busy = false;
                for (i, c) in clients.iter_mut().enumerate() {
                    self.pump_sends(i, c);
                    busy |= !self.inbox[i].is_empty();
                    self.deliver(i, c);
                    busy |= c.pending() > 0;
                }
                if !busy {
                    return;
                }
            }
            panic!("did not settle");
        }
    }

    fn cell<'a>(fx: &'a Value, row: u64, col: u64) -> Option<&'a Value> {
        fx["cells"].as_array().unwrap().iter().find(|c| c["row"] == row && c["col"] == col)
    }

    #[test]
    fn a_local_format_reports_its_props_and_the_display_it_changes() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        c.local(&set(0, 0, "0.25")).unwrap();
        let op = json!([{"SetFormat": {"sheet": 1, "rect": {"r0":0,"c0":0,"r1":0,"c1":0},
            "props": {"number_format": "0%", "bold": true, "color": null}}}]);
        let fx = c.local(&op).unwrap();
        assert_eq!(fx["full"], false);
        assert_eq!(fx["formats"], json!([{"sheet": 1, "rect": {"r0":0,"c0":0,"r1":0,"c1":0},
            "props": {"bold": true, "color": null, "number_format": "0%"}}]));
        assert_eq!(cell(&fx, 0, 0).unwrap()["display"], "25%");
        // A full repaint reads formats from the snapshot, so they must survive
        // the visigrid-json export and a reload.
        let snap = c.snapshot().unwrap();
        let text = snap.to_string();
        assert!(text.contains("0%") && text.contains("bold"), "formats in the snapshot: {text}");
        let reloaded = CollabCore::new(&snap, c.last_seen()).unwrap();
        assert_eq!(
            checksum(&reloaded.client.wb),
            checksum(&c.client.wb),
            "the optimistic workbook, formats included, round-trips through the snapshot"
        );
    }

    #[test]
    fn local_edit_reports_the_written_cell_and_its_dependents_incrementally() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        let fx = c.local(&set(1, 0, "=A1*10")).unwrap();
        assert_eq!(fx["full"], false);
        let a2 = cell(&fx, 1, 0).expect("written formula reported");
        assert_eq!(a2["raw"], "=A1*10");
        assert_eq!(a2["value"], 10.0);
        assert_eq!(a2["sheet"], 1);
        // A dependent of an edit is reported too.
        let fx = c.local(&set(0, 0, "4")).unwrap();
        assert_eq!(fx["full"], false);
        assert_eq!(cell(&fx, 0, 0).unwrap()["value"], 4.0);
        assert_eq!(cell(&fx, 1, 0).unwrap()["value"], 40.0);
        assert_eq!(c.pending(), 2);
        let env = c.poll_send().unwrap();
        assert_eq!(env["base_seq"], 0);
        assert!(env["op"].is_array());
    }

    #[test]
    fn two_clients_converge_through_the_json_api_and_checksums_agree() {
        let mut room = Room::new(2);
        let mut clients = vec![CollabCore::new(&doc(), 0).unwrap(), CollabCore::new(&doc(), 0).unwrap()];
        clients[0].local(&set(0, 0, "2")).unwrap();
        clients[1].local(&set(0, 1, "=A1*3")).unwrap();
        clients[1].local(&set(1, 1, "=B1+1")).unwrap();
        room.settle(&mut clients);
        let server = checksum(&room.server.wb);
        for c in &clients {
            assert_eq!(c.checksum(), server);
            assert_eq!(c.pending(), 0);
        }
        assert_eq!(clients[0].snapshot().unwrap(), clients[1].snapshot().unwrap());
        let wb = &clients[0].client.wb;
        assert_eq!(wb.sheets()[0].get_display(0, 1), "6");
        assert_eq!(wb.sheets()[0].get_display(1, 1), "7");
    }

    #[test]
    fn a_remote_op_reports_the_cells_it_changed() {
        let mut room = Room::new(2);
        let mut a = CollabCore::new(&doc(), 0).unwrap();
        let mut b = CollabCore::new(&doc(), 0).unwrap();
        b.local(&set(1, 0, "=A1+100")).unwrap();
        room.pump_sends(1, &mut b);
        room.deliver(1, &mut b);
        room.deliver(0, &mut a);
        a.local(&set(0, 0, "5")).unwrap();
        room.pump_sends(0, &mut a);
        let fxs = room.deliver(1, &mut b);
        let fx = fxs.iter().find(|fx| !fx["cells"].as_array().unwrap().is_empty()).expect("remote op effects");
        assert_eq!(fx["full"], false);
        assert_eq!(cell(fx, 0, 0).unwrap()["value"], 5.0);
        assert_eq!(cell(fx, 1, 0).unwrap()["value"], 105.0, "dependent of the remote write");
    }

    #[test]
    fn a_refused_edit_rebuilds_with_full_effects_and_keeps_later_edits() {
        let mut room = Room::new(2);
        let mut clients = vec![CollabCore::new(&doc(), 0).unwrap(), CollabCore::new(&doc(), 0).unwrap()];
        // A writes inside the range B is about to replace atomically (V1:
        // an overlapping concurrent op refuses the atomic one).
        clients[0].local(&set(0, 0, "7")).unwrap();
        room.pump_sends(0, &mut clients[0]);
        let replace = json!([{"ReplaceRange": {"sheet": 1, "row": 0, "col": 0, "values": [[{"Value": "9"}]]}}]);
        clients[1].local(&replace).unwrap();
        clients[1].local(&set(5, 5, "keep me")).unwrap();
        let fxs = room.deliver(1, &mut clients[1]);
        assert!(fxs.iter().any(|fx| fx["full"] == true), "refusal rebuilds the optimistic state");
        room.settle(&mut clients);
        let server = checksum(&room.server.wb);
        assert!(clients.iter().all(|c| c.checksum() == server));
        let wb = &clients[1].client.wb;
        assert_eq!(wb.sheets()[0].get_display(0, 0), "7", "the refused replace is gone");
        assert_eq!(wb.sheets()[0].get_raw(5, 5), "keep me", "the later edit survives the refusal");
    }

    #[test]
    fn checksum_mismatch_resync_and_snapshot_welcome_are_reported() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        let fx = c.receive(&json!({"type": "checksum", "seq": 0, "checksum": "not-it"})).unwrap();
        assert_eq!(fx["checksum_mismatch"], true);
        let fx = c.receive(&json!({"type": "checksum", "seq": 0, "checksum": c.checksum()})).unwrap();
        assert_eq!(fx["checksum_mismatch"], false);

        let fx = c.receive(&json!({"type": "resync", "reason": "document_replaced"})).unwrap();
        assert_eq!(fx["reconnect"], true);

        c.reconnect();
        let fx = c
            .receive(&json!({"type": "welcome", "seq": 3, "actor": 1, "snapshot_url": "/api/sheets/x/snapshot",
                "snapshot_seq": 2, "ops": [{"seq": 3, "actor": 2, "client_op_id": Uuid::from_u128(9),
                "op": set(2, 2, "z")}]}))
            .unwrap();
        assert_eq!(fx["need_snapshot"], true);
        assert_eq!(fx["snapshot_seq"], 2);
        let replaced: Value = serde_json::from_str(
            r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Sheet1","cells":[{"row":0,"col":0,"value":42}]}],"collab_sheet_ids":[1]}"#,
        )
        .unwrap();
        let fx = c.load_snapshot(&replaced, 2).unwrap();
        assert_eq!(fx["full"], true);
        assert_eq!(fx["need_snapshot"], false);
        assert_eq!(c.last_seen(), 3, "the welcome's ops after the snapshot are applied");
        let wb = &c.client.wb;
        assert_eq!(wb.sheets()[0].get_display(0, 0), "42");
        assert_eq!(wb.sheets()[0].get_raw(2, 2), "z");
    }

    #[test]
    fn snapshot_round_trips_through_load() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        c.local(&set(3, 3, "=A1*2")).unwrap();
        let snap = c.snapshot().unwrap();
        assert_eq!(snap["collab_sheet_ids"], json!([1]));
        let again = CollabCore::new(&snap, 1).unwrap();
        assert_eq!(checksum(&again.client.wb), checksum(&c.client.wb));
    }

    #[test]
    fn full_effects_carry_the_display_of_every_populated_cell() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        c.local(&set(0, 0, "3.5")).unwrap(); // unformatted decimal
        c.local(&set(1, 0, "0.25")).unwrap();
        c.local(&json!([{"SetFormat": {"sheet": 1, "rect": {"r0": 1, "c0": 0, "r1": 1, "c1": 0},
            "props": {"number_format": "0.00%"}}}])).unwrap();
        c.local(&set(2, 0, "=1/0")).unwrap(); // error
        c.local(&set(3, 0, "=SEQUENCE(3)")).unwrap(); // spills into A5:A6
        c.local(&json!([{"SetFormat": {"sheet": 1, "rect": {"r0": 9, "c0": 9, "r1": 9, "c1": 9},
            "props": {"bold": true}}}])).unwrap(); // format-only, no content
        let snap = c.snapshot().unwrap();
        let mut fresh = CollabCore::new(&doc(), 0).unwrap();
        let fx = fresh.load_snapshot(&snap, 0).unwrap();
        assert_eq!(fx["full"], true);
        let c00 = cell(&fx, 0, 0).expect("unformatted decimal");
        assert_eq!(c00["display"], "3.50", "the engine's display, not the raw value");
        assert_eq!(c00["raw"], "3.5");
        assert_eq!(cell(&fx, 1, 0).expect("formatted")["display"], "25.00%");
        let err = cell(&fx, 2, 0).expect("error cell");
        assert_eq!(err["error"], "#DIV/0!");
        assert_eq!(err["display"], "#DIV/0!");
        assert_eq!(cell(&fx, 3, 0).expect("spill parent")["display"], "1");
        let receiver = cell(&fx, 5, 0).expect("spill receiver");
        assert_eq!(receiver["display"], "3");
        assert_eq!(receiver["raw"], "");
        assert!(cell(&fx, 9, 9).is_none(), "format-only cells come from the snapshot, not cells");
        assert_eq!(fx["cells"].as_array().unwrap().len(), 6);
        for c in fx["cells"].as_array().unwrap() {
            assert_eq!(c["sheet"], 1);
            assert_eq!(fresh.display(1, c["row"].as_u64().unwrap() as usize, c["col"].as_u64().unwrap() as usize).as_deref(),
                c["display"].as_str());
        }
        assert_eq!(fresh.display(99, 0, 0), None);
    }

    #[test]
    fn a_file_imports_as_new_tabs_in_one_change_and_one_undo() {
        let mut room = Room::new(2);
        let mut clients = [CollabCore::new(&doc(), 0).unwrap(), CollabCore::new(&doc(), 0).unwrap()];
        let file = json!({"format": "visigrid-json", "version": 2, "sheets": [
            {"name": "Sheet1", "cells": [{"row": 0, "col": 0, "value": "Qty", "fmt": {"bold": true}},
                {"row": 1, "col": 0, "value": 4}, {"row": 2, "col": 0, "formula": "=A2*3"}],
             "col_widths": {"0": 90.0}, "frozen_rows": 1, "hidden_cols": [3],
             "merges": [{"start_row": 5, "start_col": 0, "end_row": 5, "end_col": 1}]},
            {"name": "Notes", "cells": [{"row": 0, "col": 0, "value": "hello"}], "charts": [{"kind": "bar"}]}]});
        let plan = clients[0].import_ops(&file).unwrap();
        let names: Vec<&str> = plan["sheets"].as_array().unwrap().iter().map(|s| s["name"].as_str().unwrap()).collect();
        assert_eq!(names, ["Sheet1 (2)", "Notes"]);
        let dropped = plan["dropped"].as_array().unwrap();
        assert!(dropped.iter().any(|d| d.as_str().unwrap().contains("renamed")), "{dropped:?}");
        assert!(dropped.iter().any(|d| d.as_str().unwrap().contains("1 chart")), "{dropped:?}");
        clients[0].local(&plan["ops"]).unwrap();
        room.settle(&mut clients);
        for c in &clients {
            let wb = &c.client.wb;
            assert_eq!(wb.sheets().len(), 3);
            let s = &wb.sheets()[1];
            assert_eq!(s.name, "Sheet1 (2)");
            assert_eq!(s.get_display(2, 0), "12");
            assert!(s.get_format(0, 0).bold);
            assert_eq!(s.layout.frozen_rows, 1);
            assert_eq!(s.layout.col_widths.get(&0), Some(&90.0));
            assert!(s.layout.hidden_cols.contains(&3));
            assert_eq!(s.merged_regions.len(), 1);
            assert_eq!(wb.sheets()[2].get_raw(0, 0), "hello");
        }
        assert_eq!(checksum(&clients[0].client.wb), checksum(&room.server.wb));
        assert_eq!(clients[0].undo()["undo"]["applied"], true);
        room.settle(&mut clients);
        for c in &clients {
            assert_eq!(c.client.wb.sheets().len(), 1, "one undo removes the whole import");
        }
    }

    #[test]
    fn a_duplicated_tab_follows_its_source_and_is_named_like_excel() {
        let mut room = Room::new(2);
        let mut clients = [CollabCore::new(&doc(), 0).unwrap(), CollabCore::new(&doc(), 0).unwrap()];
        clients[0].local(&set(1, 0, "=A1+1")).unwrap();
        room.settle(&mut clients);
        let plan = clients[1].duplicate_ops(1).unwrap();
        assert_eq!(plan["name"], "Sheet1 (2)");
        clients[1].local(&plan["ops"]).unwrap();
        room.settle(&mut clients);
        let again = clients[0].duplicate_ops(plan["key"].as_u64().unwrap()).unwrap();
        assert_eq!(again["name"], "Sheet1 (3)", "a copy of a copy counts on from the base name");
        for c in &clients {
            let wb = &c.client.wb;
            assert_eq!(wb.sheets().iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Sheet1", "Sheet1 (2)"]);
            assert_eq!(wb.sheets()[1].get_display(1, 0), "2");
        }
        assert_eq!(checksum(&clients[0].client.wb), checksum(&room.server.wb));
    }

    #[test]
    fn undo_reverts_only_this_users_edit_and_reaches_the_other_browser() {
        let mut room = Room::new(2);
        let mut clients = [CollabCore::new(&doc(), 0).unwrap(), CollabCore::new(&doc(), 0).unwrap()];
        clients[0].local(&set(1, 0, "mine")).unwrap();
        clients[1].local(&set(2, 0, "theirs")).unwrap();
        room.settle(&mut clients);
        let fx = clients[0].undo();
        assert_eq!(fx["undo"]["applied"], true);
        assert_eq!(cell(&fx, 1, 0).unwrap()["raw"], "");
        assert_eq!(fx["can_redo"], true);
        room.settle(&mut clients);
        for c in &clients {
            assert_eq!(c.client.wb.sheets()[0].get_raw(1, 0), "");
            assert_eq!(c.client.wb.sheets()[0].get_raw(2, 0), "theirs");
        }
        assert_eq!(checksum(&clients[0].client.wb), checksum(&room.server.wb));
        let fx = clients[0].redo();
        assert_eq!(cell(&fx, 1, 0).unwrap()["raw"], "mine");
        assert_eq!(clients[1].undo()["undo"]["applied"], true);
        room.settle(&mut clients);
        assert_eq!(clients[1].client.wb.sheets()[0].get_raw(2, 0), "");
        assert_eq!(clients[0].undo()["undo"]["reason"], Value::Null);
        let empty = CollabCore::new(&doc(), 0).unwrap().undo();
        assert_eq!(empty["undo"]["reason"], "empty");
    }

    #[test]
    fn viewport_reports_engine_displays_kinds_and_deduplicated_formats() {
        let mut c = CollabCore::new(&doc(), 0).unwrap();
        c.local(&set(0, 1, "0.25")).unwrap();
        c.local(&set(1, 0, "=1/0")).unwrap();
        c.local(&set(1, 1, "label")).unwrap();
        let fmt = json!([{"SetFormat": {"sheet": 1, "rect": {"r0":0,"c0":0,"r1":0,"c1":1},
            "props": {"number_format": "0%", "bold": true}}}]);
        c.local(&fmt).unwrap();
        let v = c.viewport(1, 0, 0, 5, 5).unwrap();
        let at = |r: u64, col: u64| {
            let i = (0..v["rows"].as_array().unwrap().len())
                .find(|&i| v["rows"][i] == r && v["cols"][i] == col)
                .expect("cell listed");
            (v["text"][i].clone(), v["kind"][i].clone(), v["fmt"][i].clone())
        };
        let (t, k, f) = at(0, 1);
        assert_eq!((t, k), (json!("25%"), json!("n")));
        assert_eq!(v["formats"][f.as_u64().unwrap() as usize]["bold"], true);
        assert_eq!(at(0, 0).2, f, "A1 and B1 share one format entry");
        assert_eq!(v["formats"].as_array().unwrap().len(), 1);
        assert_eq!(at(1, 0).1, json!("e"));
        assert_eq!(at(1, 1), (json!("label"), json!("t"), json!(-1)));
        // Every listed display is exactly what display() reports.
        for i in 0..v["rows"].as_array().unwrap().len() {
            let (r, col) = (v["rows"][i].as_u64().unwrap() as usize, v["cols"][i].as_u64().unwrap() as usize);
            assert_eq!(v["text"][i], json!(c.display(1, r, col).unwrap()));
        }
        assert!(c.viewport(99, 0, 0, 1, 1).is_none());
        assert_eq!(c.raw(1, 1, 0).as_deref(), Some("=1/0"));
        let r = v["raw"].as_array().unwrap();
        assert!(r.iter().any(|x| x == "=1/0"), "viewport carries raw input");
        assert_eq!(c.raw(1, 0, 1).as_deref(), Some("0.25"));
    }

    #[test]
    fn synthetic_workbook_has_its_extent_layout_and_formats() {
        let c = CollabCore::synthetic(200, 10);
        let key = c.sheets()[0]["key"].as_u64().unwrap();
        let l = c.layout(key).unwrap();
        assert_eq!((l["rows"].clone(), l["cols"].clone(), l["frozen_rows"].clone()), (json!(200), json!(10), json!(1)));
        let v = c.viewport(key, 0, 0, 3, 4).unwrap();
        assert_eq!(v["rows"].as_array().unwrap().len(), 20);
        assert!(v["text"].as_array().unwrap().iter().any(|t| t.as_str().unwrap().ends_with('%')));
    }

    /// Timing probe for the grid benchmark's sheet:
    /// `cargo test --release -p visigrid-engine-wasm synthetic_timing -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn synthetic_timing() {
        let t = std::time::Instant::now();
        let c = CollabCore::synthetic(100_000, 50);
        let built = t.elapsed();
        let key = c.sheets()[0]["key"].as_u64().unwrap();
        let t = std::time::Instant::now();
        let l = c.layout(key).unwrap();
        let laid = t.elapsed();
        let t = std::time::Instant::now();
        let mut n = 0;
        for i in 0..100 {
            let r0 = i * 997 % 99_950;
            n += c.viewport(key, r0, 0, r0 + 40, 20).unwrap()["rows"].as_array().unwrap().len();
        }
        eprintln!("build {built:?}, layout {laid:?} {l}, 100 viewports (41x21) {:?}, {n} cells", t.elapsed());
    }
    #[test]
    fn protected_import_cannot_be_flattened_into_editable_collaboration_ops() {
        let core = CollabCore::new(&doc(), 0).unwrap();
        let before = core.snapshot().unwrap();
        let mut source = doc();
        source["future_feature"] = json!({"setting": "retain"});
        assert!(core.import_ops(&source).is_err());
        assert_eq!(core.snapshot().unwrap(), before);
    }

}
