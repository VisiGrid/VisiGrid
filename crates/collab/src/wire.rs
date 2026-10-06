//! Go WebSocket protocol v2 adapter for the native replica. This layer only
//! translates frames; `Client` remains the sole convergence implementation.
use crate::{
    apply::collab_checksum,
    client::{Client, ToClient, ToServer},
    op::{ops_from_json, ops_to_json, CellContent, CollabOp},
    server::Committed,
};
use serde_json::{json, Value};
use uuid::Uuid;
use visigrid_engine::{sheet::SheetId, workbook::Workbook};

pub struct SnapshotRequest {
    pub url: String,
    pub seq: u64,
    pub revision: Option<u64>,
}
pub enum Received {
    Updated,
    Snapshot(SnapshotRequest),
    Reconnect,
}
pub struct WireReplica {
    pub client: Client,
    writable: bool,
    ready: bool,
    welcome: Option<Value>,
}
impl WireReplica {
    pub fn new(wb: Workbook, seq: u64, writable: bool) -> Self {
        let mut client = Client::new(0);
        client.confirmed = wb.clone();
        client.wb = wb;
        client.last_seen = seq;
        client.record_changes();
        client.disconnect();
        Self {
            client,
            writable,
            ready: false,
            welcome: None,
        }
    }
    pub fn hello(&mut self, engine_commit: &str, sequence_known: bool) -> Value {
        self.ready = false;
        let _ = self.client.reconnect();
        let mut frame = json!({"type":"hello", "protocol":2, "engine_commit":engine_commit});
        if sequence_known {
            frame["last_seq"] = json!(self.client.last_seen);
        }
        frame
    }
    pub fn disconnect(&mut self) {
        self.ready = false;
        self.client.disconnect();
    }
    pub fn set_cell(
        &mut self,
        id: Uuid,
        key: u64,
        row: usize,
        col: usize,
        raw: String,
    ) -> Result<(), String> {
        if !self.ready {
            return Err("Wait for collaboration to reconnect".into());
        }
        if !self.writable {
            return Err("This workbook is read-only".into());
        }
        self.client.wb.ensure_writable()?;
        let index = self
            .client
            .wb
            .idx_for_sheet_id(SheetId(key))
            .ok_or("Unknown sheet identity")?;
        let content = if raw.is_empty() {
            CellContent::Clear
        } else if raw.starts_with('=') {
            CellContent::Formula(raw)
        } else {
            CellContent::Value(raw)
        };
        let sheet_name = self.client.wb.sheets()[index].name.clone();
        self.client.local(
            id,
            vec![CollabOp::SetCell {
                sheet: key,
                sheet_name,
                row,
                col,
                content,
            }],
        );
        Ok(())
    }
    pub fn poll_send(&mut self) -> Option<Value> {
        if !self.ready {
            return None;
        }
        match self.client.poll_send()? {
            ToServer::Submit(e) => Some(json!({"type":"op", "envelope":{
                "client_op_id":e.client_op_id, "base_seq":e.base_seq, "op":ops_to_json(&e.ops)}})),
            ToServer::Hello { .. } => None,
        }
    }
    pub fn receive(&mut self, frame: &Value) -> Result<Received, String> {
        match frame.get("type").and_then(Value::as_str).unwrap_or("") {
            "welcome" => {
                if let Some(url) = frame.get("snapshot_url").and_then(Value::as_str) {
                    self.ready = false;
                    self.welcome = Some(frame.clone());
                    return Ok(Received::Snapshot(SnapshotRequest {
                        url: url.into(),
                        seq: number(frame, "snapshot_seq")?,
                        revision: frame.get("snapshot_revision").and_then(Value::as_u64),
                    }));
                }
                self.finish_welcome(frame)?;
            }
            "op" => self.apply_committed(frame)?,
            "ack" => {
                let seq = number(frame, "seq")?;
                if seq > self.client.last_seen {
                    self.check_order(seq)?;
                    let id = op_id(frame)?;
                    crate::clock::install_frame_clock(&mut self.client, frame)?;
                    self.client.ack(id, seq);
                    if self.client.last_seen != seq {
                        return Err("Acknowledgement does not match our pending edit".into());
                    }
                }
            }
            "rejected" => {
                let id = op_id(frame)?;
                match frame.get("result").and_then(Value::as_str) {
                    Some("dropped") => self.client.dropped(id),
                    Some("refused") => self.client.receive(ToClient::Refused {
                        client_op_id: id,
                        reason: frame
                            .get("reason")
                            .and_then(Value::as_str)
                            .unwrap_or("")
                            .into(),
                    }),
                    _ => return Err("Invalid rejection frame".into()),
                }
            }
            "checksum" => {
                let seq = number(frame, "seq")?;
                let sum = frame
                    .get("checksum")
                    .and_then(Value::as_str)
                    .ok_or("Missing checksum")?;
                if seq == self.client.last_seen
                    && !sum.is_empty()
                    && sum != collab_checksum(&self.client.confirmed)
                {
                    return Err("Workbook checksum differs from the server".into());
                }
            }
            "resync" => {
                self.disconnect();
                return Ok(Received::Reconnect);
            }
            "error" => {
                self.disconnect();
                return Err(frame
                    .get("message")
                    .and_then(Value::as_str)
                    .unwrap_or("Collaboration refused")
                    .into());
            }
            "presence" | "presence_leave" | "pong" => {}
            _ => return Err("Unknown collaboration frame".into()),
        }
        Ok(Received::Updated)
    }
    pub fn load_snapshot(&mut self, wb: Workbook, seq: u64) -> Result<(), String> {
        // Never substitute an incomplete/protected projection into a live
        // editable replica. Its original stays with the normal read-only UI.
        wb.ensure_writable()?;
        let pending = self.welcome.as_ref().ok_or("No snapshot welcome pending")?;
        if number(pending, "snapshot_seq")? != seq {
            return Err("Snapshot sequence differs from its welcome".into());
        }
        self.client.replace_document(wb, seq);
        let frame = self.welcome.take().ok_or("No snapshot welcome pending")?;
        self.finish_welcome(&frame)
    }
    fn finish_welcome(&mut self, frame: &Value) -> Result<(), String> {
        crate::clock::install_frame_clock(&mut self.client, frame)?;
        self.client.actor = number(frame, "actor")?;
        for op in frame
            .get("ops")
            .and_then(Value::as_array)
            .ok_or("Welcome missing operations")?
        {
            self.apply_committed(op)?;
        }
        let head = number(frame, "seq")?;
        if head != self.client.last_seen {
            return Err("Welcome is missing sequenced operations".into());
        }
        self.client.receive(ToClient::Welcome { head });
        self.ready = true;
        Ok(())
    }
    fn check_order(&self, seq: u64) -> Result<(), String> {
        if seq > self.client.last_seen + 1 {
            Err("Collaboration sequence gap".into())
        } else {
            Ok(())
        }
    }
    fn apply_committed(&mut self, frame: &Value) -> Result<(), String> {
        let seq = number(frame, "seq")?;
        self.check_order(seq)?;
        if seq <= self.client.last_seen {
            return Ok(());
        }
        self.client.wb.ensure_writable()?;
        let committed = Committed {
            seq,
            actor: number(frame, "actor")?,
            client_op_id: op_id(frame)?,
            ops: ops_from_json(frame.get("op").ok_or("Missing operations")?)?,
        };
        crate::clock::install_frame_clock(&mut self.client, frame)?;
        self.client.receive(ToClient::Op(committed));
        Ok(())
    }
}
fn number(frame: &Value, key: &str) -> Result<u64, String> {
    frame
        .get(key)
        .and_then(Value::as_u64)
        .ok_or_else(|| format!("Missing {key}"))
}
fn op_id(frame: &Value) -> Result<Uuid, String> {
    frame
        .get("client_op_id")
        .and_then(Value::as_str)
        .and_then(|s| Uuid::parse_str(s).ok())
        .ok_or_else(|| "Invalid operation identity".into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        op::Envelope,
        server::{Server, Submitted},
    };
    fn wb() -> Workbook {
        let mut wb = Workbook::new();
        wb.sheet_mut(0).unwrap().id = SheetId(37);
        wb.set_next_sheet_id(38);
        wb
    }
    fn live(writable: bool, actor: u64) -> WireReplica {
        let mut replica = WireReplica::new(wb(), 0, writable);
        replica.hello("engine", true);
        replica
            .receive(&json!({"type":"welcome","seq":0,"actor":actor,"ops":[]}))
            .unwrap();
        replica
    }
    fn committed(c: &Committed) -> Value {
        json!({"type":"op","seq":c.seq,"actor":c.actor,"client_op_id":c.client_op_id,"op":ops_to_json(&c.ops)})
    }
    fn sequence(server: &mut Server, replica: &mut WireReplica) -> Committed {
        let frame = replica.poll_send().unwrap();
        let e = &frame["envelope"];
        let env = Envelope {
            client_op_id: op_id(e).unwrap(),
            base_seq: number(e, "base_seq").unwrap(),
            actor: replica.client.actor,
            ops: ops_from_json(&e["op"]).unwrap(),
        };
        match server.submit(&env) {
            Submitted::Committed(c) => c,
            _ => panic!("expected commit"),
        }
    }
    #[test]
    fn concurrent_values_and_formulas_converge_by_stable_sheet_identity() {
        let (mut a, mut b) = (live(true, 1), live(true, 2));
        let mut server = Server::new();
        server.wb = wb();
        a.set_cell(Uuid::from_u128(1), 37, 0, 0, "3".into())
            .unwrap();
        b.set_cell(Uuid::from_u128(2), 37, 0, 1, "=A1*2".into())
            .unwrap();
        let first = sequence(&mut server, &mut a);
        b.receive(&committed(&first)).unwrap();
        a.receive(&json!({"type":"ack","seq":first.seq,"client_op_id":first.client_op_id}))
            .unwrap();
        let second = sequence(&mut server, &mut b);
        a.receive(&committed(&second)).unwrap();
        b.receive(&json!({"type":"ack","seq":second.seq,"client_op_id":second.client_op_id}))
            .unwrap();
        assert_eq!(collab_checksum(&a.client.wb), collab_checksum(&b.client.wb));
        assert_eq!(collab_checksum(&a.client.wb), collab_checksum(&server.wb));
        assert_eq!(a.client.wb.sheet(0).unwrap().get_display(0, 1), "6");
    }
    #[test]
    fn replay_of_a_committed_pending_edit_is_not_applied_twice() {
        let mut a = live(true, 1);
        let mut server = Server::new();
        server.wb = wb();
        a.set_cell(Uuid::from_u128(1), 37, 0, 0, "4".into())
            .unwrap();
        let c = sequence(&mut server, &mut a);
        a.disconnect();
        assert_eq!(a.hello("engine", true)["last_seq"], 0);
        a.receive(&json!({"type":"welcome","seq":1,"actor":1,"ops":[committed(&c)]}))
            .unwrap();
        assert_eq!(a.client.pending_count(), 0);
        assert!(a.poll_send().is_none());
        a.receive(&committed(&c)).unwrap();
        a.receive(&json!({"type":"ack","seq":1,"client_op_id":c.client_op_id}))
            .unwrap();
        assert_eq!(a.client.last_seen, 1);
        assert_eq!(a.client.wb.sheet(0).unwrap().get_raw(0, 0), "4");
    }
    #[test]
    fn viewers_receive_but_cannot_submit() {
        let mut viewer = live(false, 2);
        let mut editor = live(true, 1);
        let mut server = Server::new();
        server.wb = wb();
        editor
            .set_cell(Uuid::from_u128(1), 37, 0, 0, "7".into())
            .unwrap();
        viewer
            .receive(&committed(&sequence(&mut server, &mut editor)))
            .unwrap();
        assert_eq!(viewer.client.wb.sheet(0).unwrap().get_raw(0, 0), "7");
        assert!(viewer
            .set_cell(Uuid::from_u128(2), 37, 0, 0, "8".into())
            .is_err());
        assert!(viewer.poll_send().is_none());
    }
    #[test]
    fn protected_workbook_refuses_local_and_remote_mutation() {
        let mut a = live(true, 1);
        a.client.wb.sheet_mut(0).unwrap().read_only_reason = Some("unsupported content".into());
        assert!(a
            .set_cell(Uuid::from_u128(1), 37, 0, 0, "3".into())
            .is_err());
        let mut editor = live(true, 2);
        let mut server = Server::new();
        server.wb = wb();
        editor
            .set_cell(Uuid::from_u128(2), 37, 0, 0, "7".into())
            .unwrap();
        assert!(a
            .receive(&committed(&sequence(&mut server, &mut editor)))
            .is_err());
        assert_eq!(a.client.last_seen, 0);
        assert_eq!(a.client.wb.sheet(0).unwrap().get_raw(0, 0), "");
    }
    #[test]
    fn sequenced_clock_is_applied_once_and_invalid_clock_cannot_apply_cells() {
        let mut viewer = live(false, 2);
        let id = Uuid::from_u128(77);
        let mut frame = json!({"type":"op","seq":1,"actor":1,"client_op_id":id,
            "op":ops_to_json(&[CollabOp::SetCell { sheet:37,sheet_name:"Sheet1".into(),row:0,col:0,
                content:CellContent::Formula("=RAND()".into()) }]),
            "clock":{"now_ms":1791028800123_i64,"utc_offset_seconds":0,"seed":"18446744073709551615"}});
        viewer.receive(&frame).unwrap();
        let before = collab_checksum(&viewer.client.wb);
        frame["clock"]["seed"] = json!("2");
        viewer.receive(&frame).unwrap();
        assert_eq!(collab_checksum(&viewer.client.wb), before);
        assert_eq!(viewer.client.wb.recalc_clock().unwrap().seed, Some(u64::MAX));
        frame["seq"] = json!(2);
        frame["clock"]["seed"] = json!(2);
        assert!(viewer.receive(&frame).is_err());
        assert_eq!(viewer.client.last_seen, 1);
        assert_eq!(collab_checksum(&viewer.client.wb), before);
    }
    #[test]
    fn unexpected_snapshot_does_not_replace_the_current_workbook() {
        let mut a = live(true, 1);
        let mut replacement = wb();
        replacement.set_cell_value_tracked_at(0, 0, 0, "replacement");
        assert!(a.load_snapshot(replacement.clone(), 4).is_err());
        assert_eq!(a.client.last_seen, 0);
        assert_eq!(a.client.wb.sheet(0).unwrap().get_raw(0, 0), "");
        a.receive(&json!({"type":"welcome","seq":4,"actor":1,"ops":[],"snapshot_url":"/snapshot","snapshot_seq":4})).unwrap();
        assert!(a.load_snapshot(replacement, 3).is_err());
        assert_eq!(a.client.last_seen, 0);
        assert_eq!(a.client.wb.sheet(0).unwrap().get_raw(0, 0), "");
    }
    #[test]
    fn welcome_with_missing_operations_is_refused() {
        let mut a = WireReplica::new(wb(), 0, true);
        a.hello("engine", true);
        assert!(a
            .receive(&json!({"type":"welcome","seq":2,"actor":1,"ops":[]}))
            .is_err());
        assert!(a
            .set_cell(Uuid::from_u128(1), 37, 0, 0, "3".into())
            .is_err());
    }
}
