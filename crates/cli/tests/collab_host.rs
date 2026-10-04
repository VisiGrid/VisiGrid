//! Golden protocol tests for `vgrid collab-host` (vault spec "VisiGrid
//! Collaborative Workbook Model", Phase 2 interfaces). They drive the real
//! binary over stdio, one JSON line per request, exactly as the Go
//! sequencer does.
//!
//! Idempotency is the sequencer's job, not the host's: the host does not
//! deduplicate `client_op_id`. A resent envelope reaches the host only if
//! the sequencer forgot it, and would then be applied twice; the sequencer
//! answers resends from its log (`UNIQUE (sheet_id, client_op_id)`).

use std::io::{BufRead, BufReader, Write};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};

use serde_json::{json, Value};
use uuid::Uuid;
use visigrid_cli::collab_host::read_bounded_line;
use visigrid_collab::apply::checksum;
use visigrid_collab::op::{ops_to_json, Axis, CellContent, CollabOp};
use visigrid_engine::workbook::Workbook;

const BLANK: &str = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Sheet1"}]}"#;
const DATA_SHEET: u64 = (1 << 62) | 7;

struct Host {
    child: Child,
    stdin: Option<ChildStdin>,
    stdout: BufReader<ChildStdout>,
    next_id: u64,
}

impl Host {
    fn spawn() -> Host {
        let mut child = Command::new(env!("CARGO_BIN_EXE_vgrid"))
            .arg("collab-host")
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn vgrid collab-host");
        let stdin = child.stdin.take().unwrap();
        let stdout = BufReader::new(child.stdout.take().unwrap());
        Host { child, stdin: Some(stdin), stdout, next_id: 1 }
    }

    fn raw(&mut self, line: &str) -> Value {
        let stdin = self.stdin.as_mut().expect("stdin open");
        stdin.write_all(line.as_bytes()).unwrap();
        stdin.write_all(b"\n").unwrap();
        stdin.flush().unwrap();
        let mut reply = String::new();
        self.stdout.read_line(&mut reply).unwrap();
        assert!(reply.ends_with('\n'), "one response per line, got {reply:?}");
        serde_json::from_str(&reply).expect("response is JSON")
    }

    fn call(&mut self, mut req: Value) -> Value {
        let id = self.next_id;
        self.next_id += 1;
        req["id"] = json!(id);
        let reply = self.raw(&req.to_string());
        assert_eq!(reply["id"], json!(id), "id echoed: {reply}");
        reply
    }

    fn ok(&mut self, req: Value) -> Value {
        let reply = self.call(req.clone());
        assert_eq!(reply["ok"], json!(true), "{req} -> {reply}");
        reply
    }
}

impl Drop for Host {
    fn drop(&mut self) {
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

fn blank() -> Value {
    serde_json::from_str(BLANK).unwrap()
}

fn set(row: usize, col: usize, content: CellContent) -> CollabOp {
    CollabOp::SetCell { sheet: 1, sheet_name: "Sheet1".into(), row, col, content }
}

fn val(s: &str) -> CellContent {
    CellContent::Value(s.into())
}

fn formula(s: &str) -> CellContent {
    CellContent::Formula(s.into())
}

fn env(base: u64, ops: &[CollabOp]) -> Value {
    json!({"client_op_id": Uuid::new_v4(), "base_seq": base, "actor": 9, "op": ops_to_json(ops)})
}

fn submit(h: &mut Host, base: u64, ops: &[CollabOp], concurrent: &[(u64, Value)]) -> Value {
    let concurrent: Vec<Value> = concurrent.iter().map(|(s, op)| json!({"seq": s, "op": op})).collect();
    h.ok(json!({"cmd": "submit", "envelope": env(base, ops), "concurrent": concurrent}))
}

#[test]
fn every_command_over_stdio() {
    let mut h = Host::spawn();

    // hello
    let hello = h.ok(json!({"cmd": "hello"}));
    assert_eq!(hello["protocol"], json!(2));
    assert!(hello["engine_commit"].as_str().is_some_and(|s| !s.is_empty()));

    // Malformed input and errors never kill the host.
    let bad = h.raw("this is not json");
    assert_eq!(bad["ok"], json!(false));
    assert_eq!(bad["id"], Value::Null);
    assert_eq!(h.raw("[1,2]")["ok"], json!(false));
    assert_eq!(h.call(json!({"cmd": "teleport"}))["ok"], json!(false));
    assert_eq!(h.call(json!({}))["ok"], json!(false));
    let early = h.call(json!({"cmd": "submit", "envelope": env(0, &[set(0, 0, val("1"))]), "concurrent": []}));
    assert_eq!(early["ok"], json!(false), "submit before load is an error");
    assert!(early["error"].as_str().unwrap().contains("load"));

    // load: the blank document equals Workbook::new() for every replica.
    let loaded = h.ok(json!({"cmd": "load", "document": blank(), "seq": 0}));
    assert_eq!(loaded["checksum"], json!(checksum(&Workbook::new())));

    // seq 1: A1 = 1
    let a = submit(&mut h, 0, &[set(0, 0, val("1"))], &[]);
    assert_eq!(a["result"], json!("op"));
    assert_eq!(a["seq"], json!(1));
    let op1 = a["op"].clone();
    assert!(op1.is_array(), "op payload is the atomic op list");

    // seq 2: A2 = A1+1, written without seeing seq 1 (base 0).
    let b = submit(&mut h, 0, &[set(1, 0, formula("=A1+1"))], &[(1, op1.clone())]);
    assert_eq!(b["result"], json!("op"));
    let op2 = b["op"].clone();

    // seq 3: delete row 5 (index 4).
    let del = CollabOp::Structural {
        sheet: 1,
        sheet_name: "Sheet1".into(),
        axis: Axis::Row,
        at: 4,
        count: 1,
        delete: true,
    };
    let c = submit(&mut h, 2, &[del], &[(1, op1.clone()), (2, op2.clone())][2..]);
    assert_eq!(c["result"], json!("op"));
    let op3 = c["op"].clone();

    // dropped: a write into row 5 written before the delete.
    let d = submit(&mut h, 2, &[set(4, 1, val("lost"))], &[(3, op3.clone())]);
    assert_eq!(d["result"], json!("dropped"), "{d}");
    assert!(d["reason"].as_str().is_some_and(|r| !r.is_empty()));
    assert!(d.get("op").is_none(), "no seq for a dropped envelope");

    // refused: an atomic range op against a concurrent write inside it.
    let replace = CollabOp::ReplaceRange { sheet: 1, row: 0, col: 0, values: vec![vec![val("9")]] };
    let e = submit(&mut h, 0, &[replace], &[(1, op1.clone()), (2, op2.clone()), (3, op3.clone())]);
    assert_eq!(e["result"], json!("refused"), "{e}");

    // concurrent must cover (base_seq, head] exactly, in order.
    let short = h.call(json!({"cmd": "submit", "envelope": env(1, &[set(0, 1, val("x"))]),
        "concurrent": [{"seq": 2, "op": op2}]}));
    assert_eq!(short["ok"], json!(false));
    let order = h.call(json!({"cmd": "submit", "envelope": env(1, &[set(0, 1, val("x"))]),
        "concurrent": [{"seq": 3, "op": op3}, {"seq": 2, "op": op2}]}));
    assert_eq!(order["ok"], json!(false));
    let ahead = h.call(json!({"cmd": "submit", "envelope": env(9, &[set(0, 1, val("x"))]), "concurrent": []}));
    assert_eq!(ahead["ok"], json!(false));

    // seq 4: a new sheet with a stable id, and seq 5: a cell on it.
    let add = CollabOp::AddSheet { sheet: DATA_SHEET, name: "Data".into(), index: 1 };
    let f = submit(&mut h, 3, &[add], &[]);
    assert_eq!(f["seq"], json!(4));
    let op4 = f["op"].clone();
    let on_data =
        CollabOp::SetCell { sheet: DATA_SHEET, sheet_name: "Data".into(), row: 0, col: 0, content: val("42") };
    let g = submit(&mut h, 4, &[on_data], &[]);
    assert_eq!(g["seq"], json!(5));
    let op5 = g["op"].clone();
    let live = g["checksum"].clone();

    // snapshot carries the stable sheet ids.
    let snap = h.ok(json!({"cmd": "snapshot"}));
    assert_eq!(snap["checksum"], live);
    assert_eq!(snap["seq"], json!(5));
    assert_eq!(snap["document"]["collab_sheet_ids"], json!([1, DATA_SHEET]));

    // set_clock: accepted when well formed.
    h.ok(json!({"cmd": "set_clock", "now": "2026-10-03T12:00:00Z", "tz": "UTC", "seed": 7}));
    assert_eq!(h.call(json!({"cmd": "set_clock", "now": "yesterday"}))["ok"], json!(false));

    // Load + replay equals the live host.
    let mut r = Host::spawn();
    r.ok(json!({"cmd": "load", "document": blank(), "seq": 0}));
    let replayed = r.ok(json!({"cmd": "replay", "ops": [
        {"seq": 1, "op": op1}, {"seq": 2, "op": op2}, {"seq": 3, "op": op3},
        {"seq": 4, "op": op4}, {"seq": 5, "op": op5}]}));
    assert_eq!(replayed["seq"], json!(5));
    assert_eq!(replayed["checksum"], live, "load + replay must equal the live replica");
    // A gap in replay is refused without applying anything.
    let gap = r.call(json!({"cmd": "replay", "ops": [{"seq": 7, "op": []}]}));
    assert_eq!(gap["ok"], json!(false));
    assert_eq!(r.ok(json!({"cmd": "snapshot"}))["checksum"], live);

    // Load from the snapshot equals the live host, and ops on the added
    // sheet still find it by its stable id.
    let mut s = Host::spawn();
    let reloaded = s.ok(json!({"cmd": "load", "document": snap["document"], "seq": 5}));
    assert_eq!(reloaded["checksum"], live);
    let more = CollabOp::SetCell { sheet: DATA_SHEET, sheet_name: "Data".into(), row: 1, col: 0, content: formula("=A1*2") };
    let after = submit(&mut s, 5, &[more.clone()], &[]);
    let after_live = submit(&mut h, 5, &[more], &[]);
    assert_eq!(after["checksum"], after_live["checksum"]);

    // replace_document is a barrier: older bases are refused.
    let barrier = h.ok(json!({"cmd": "replace_document", "document": blank(), "seq": 7}));
    assert_eq!(barrier["seq"], json!(7));
    assert_eq!(barrier["checksum"], json!(checksum(&Workbook::new())));
    let stale = submit(&mut h, 6, &[set(0, 0, val("1"))], &[(7, json!([]))]);
    assert_eq!(stale["result"], json!("refused"));
    assert_eq!(stale["reason"], json!("document_replaced"));
    let fresh = submit(&mut h, 7, &[set(0, 0, val("1"))], &[]);
    assert_eq!(fresh["result"], json!("op"));
    let wrong = h.call(json!({"cmd": "replace_document", "document": blank(), "seq": 99}));
    assert_eq!(wrong["ok"], json!(false));
    assert_eq!(h.call(json!({"cmd": "load", "document": "nope", "seq": 0}))["ok"], json!(false));

    // End of input ends the process cleanly.
    h.stdin.take();
    let status = h.child.wait().unwrap();
    assert!(status.success());
}

#[test]
fn over_long_lines_are_consumed_and_reported() {
    let input = format!("{}\n{{\"id\":1}}\n", "x".repeat(100));
    let mut r = std::io::Cursor::new(input.into_bytes());
    match read_bounded_line(&mut r, 10).unwrap() {
        Some(Err(len)) => assert_eq!(len, 100),
        other => panic!("expected an over-long line, got {other:?}"),
    }
    assert_eq!(read_bounded_line(&mut r, 10).unwrap(), Some(Ok("{\"id\":1}".into())));
    assert_eq!(read_bounded_line(&mut r, 10).unwrap(), None);
}
