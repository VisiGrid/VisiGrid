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
use visigrid_collab::apply::{checksum, Changes};
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
}

impl CollabCore {
    pub(crate) fn new(document: &Value, seq: u64) -> Result<CollabCore, String> {
        let Loaded { wb, layouts, active } = load(document)?;
        let mut client = Client::new(0);
        client.confirmed = wb.clone();
        client.wb = wb;
        client.last_seen = seq;
        client.record_changes();
        Ok(CollabCore {
            client,
            layouts,
            active,
            waiting_welcome: None,
            rejected: Vec::new(),
            checksum_mismatch: false,
            need_snapshot: None,
            reconnect: false,
        })
    }

    pub(crate) fn local(&mut self, ops: &Value) -> Result<Value, String> {
        let ops = ops_from_json(ops)?;
        if !ops.is_empty() {
            self.client.local(new_op_id(), ops);
        }
        Ok(self.effects())
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
                self.client.receive(ToClient::Op(c));
            }
            "ack" => {
                let id = uuid_field(frame)?;
                let seq = frame.get("seq").and_then(Value::as_u64).ok_or("ack without seq")?;
                if seq == self.client.last_seen + 1 {
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
                if seq == self.client.last_seen && checksum(&self.client.confirmed) != sum {
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
            self.client.receive(ToClient::Op(c));
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
            .map(|s| self.layouts.get(&s.id.0).cloned().unwrap_or_default())
            .collect();
        let text = visigrid_io::json::export_workbook(wb, &layouts, self.active.min(wb.sheets().len().saturating_sub(1)))?;
        let mut doc: Value = serde_json::from_str(&text).map_err(|e| e.to_string())?;
        doc["collab_sheet_ids"] = json!(wb.sheets().iter().map(|s| s.id.0).collect::<Vec<_>>());
        Ok(doc)
    }

    pub(crate) fn checksum(&self) -> String {
        checksum(&self.client.confirmed)
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

    /// What changed since the last call, for the page to repaint.
    fn effects(&mut self) -> Value {
        let ch: Changes = self.client.take_changes();
        let wb = &self.client.wb;
        let mut cells = Vec::new();
        if !ch.full {
            let mut seen = std::collections::HashSet::new();
            for (key, row, col) in &ch.cells {
                if !seen.insert((*key, *row, *col)) {
                    continue;
                }
                let Some(idx) = wb.idx_for_sheet_id(SheetId(*key)) else { continue };
                let sheet = &wb.sheets()[idx];
                let r = out_result(idx, sheet, *row, *col);
                cells.push(json!({
                    "sheet": key,
                    "row": row,
                    "col": col,
                    "raw": sheet.get_raw(*row, *col),
                    "value": r.value,
                    "error": r.error,
                    "display": r.display,
                }));
            }
        }
        let sheets: Vec<Value> = if ch.full || ch.sheets {
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
        let out = json!({
            "full": ch.full,
            "cells": cells,
            "sheets": sheets,
            "structural": ops_to_json(&ch.structural),
            "rejected": std::mem::take(&mut self.rejected),
            "checksum_mismatch": self.checksum_mismatch,
            "need_snapshot": self.need_snapshot.is_some(),
            "snapshot_url": snapshot_url,
            "snapshot_seq": snapshot_seq,
            "reconnect": self.reconnect,
        });
        self.checksum_mismatch = false;
        self.reconnect = false;
        out
    }
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

    /// The optimistic document (visigrid-json v2 with `collab_sheet_ids`).
    pub fn snapshot(&self) -> Result<JsValue, JsValue> {
        to_js(&self.core.snapshot().map_err(js_err)?)
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

#[cfg(test)]
mod tests {
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
}
