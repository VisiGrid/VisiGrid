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

#[test]
fn sequencer_clock_controls_volatile_formula_results() {
    let mut checksums = Vec::new();
    for _ in 0..2 {
        let mut host = Host::spawn();
        host.ok(json!({"cmd":"load","document":blank(),"seq":0}));
        host.ok(json!({"cmd":"set_clock","now":"2026-10-03T12:00:00.123Z",
            "tz":"America/Chicago","utc_offset_seconds":-18000,"seed":42}));
        let result = submit(&mut host, 0, &[
            set(0, 0, CellContent::Formula("=NOW()".into())),
            set(0, 1, CellContent::Formula("=RAND()".into())),
        ], &[]);
        checksums.push(result["checksum"].clone());
    }
    assert_eq!(checksums[0], checksums[1]);
    assert!(checksums[0].as_str().is_some_and(|sum| !sum.is_empty()));
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

    // An op naming a sheet that does not exist is dropped, not sequenced.
    let ghost = |sheet: u64| CollabOp::SetCell { sheet, sheet_name: "Nope".into(), row: 0, col: 0, content: val("x") };
    let before = h.ok(json!({"cmd": "snapshot"}));
    for op in [
        ghost(0),
        CollabOp::SetBold { sheet: 0, rect: visigrid_collab::op::Rect::new(0, 0, 1, 1), bold: true },
        CollabOp::Structural { sheet: 0, sheet_name: "Nope".into(), axis: Axis::Row, at: 0, count: 1, delete: false },
        CollabOp::RenameSheet { sheet: 0, name: "X".into() },
        CollabOp::DeleteSheet { sheet: 0, index: 0 },
        CollabOp::ReplaceRange { sheet: 0, row: 0, col: 0, values: vec![vec![val("1")]] },
    ] {
        let r = submit(&mut h, 8, &[op.clone()], &[]);
        assert_eq!(r["result"], json!("dropped"), "{op:?} -> {r}");
        assert_eq!(r["reason"], json!("no_such_sheet"));
    }
    let after = h.ok(json!({"cmd": "snapshot"}));
    assert_eq!(after["seq"], before["seq"], "a dropped envelope takes no seq");
    assert_eq!(after["checksum"], before["checksum"]);
    // Mixed: the real op applies, the ghost is removed from the broadcast.
    let mixed = submit(&mut h, 8, &[ghost(0), set(2, 2, val("kept"))], &[]);
    assert_eq!(mixed["result"], json!("op"));
    assert_eq!(mixed["op"].as_array().unwrap().len(), 1, "{mixed}");
    // An add or rename to a name already in use is dropped too.
    let dup = CollabOp::AddSheet { sheet: 55, name: "sheet1".into(), index: 0 };
    let r = submit(&mut h, 9, &[dup], &[]);
    assert_eq!((r["result"].clone(), r["reason"].clone()), (json!("dropped"), json!("name_taken")), "{r}");
    // A sheet added earlier in the same envelope counts as present.
    let added = CollabOp::AddSheet { sheet: 77, name: "Fresh".into(), index: 1 };
    let on_new = CollabOp::SetCell { sheet: 77, sheet_name: "Fresh".into(), row: 0, col: 0, content: val("1") };
    let both = submit(&mut h, 9, &[added, on_new], &[]);
    assert_eq!(both["op"].as_array().unwrap().len(), 2, "{both}");

    // A format op applies through the host and survives into the snapshot.
    let before = h.ok(json!({"cmd": "snapshot"}));
    let seq = before["seq"].as_u64().unwrap();
    let pct = json!([{"SetFormat": {"sheet": 1, "rect": {"r0":2,"c0":2,"r1":2,"c1":2},
        "props": {"number_format": "0.00", "h_align": "center", "color": "#00aa00"}}}]);
    let r = h.ok(json!({"cmd": "submit", "envelope": {"client_op_id": "00000000-0000-0000-0000-0000000000f0",
        "base_seq": seq, "actor": 1, "op": pct}, "concurrent": []}));
    assert_eq!(r["result"], json!("op"), "{r}");
    assert_eq!(r["op"][0]["SetFormat"]["props"]["h_align"], json!("center"));
    assert_ne!(r["checksum"], before["checksum"], "a format change changes the checksum");
    let after = h.ok(json!({"cmd": "snapshot"}));
    assert_eq!(after["checksum"], r["checksum"]);
    let text = after["document"].to_string();
    assert!(text.contains("0.00"), "the number format is in the snapshot: {text}");
    // An invalid format is refused at the wire, never applied.
    let bad = json!([{"SetFormat": {"sheet": 1, "rect": {"r0":0,"c0":0,"r1":0,"c1":0}, "props": {"color": "red"}}}]);
    let r = h.call(json!({"cmd": "submit", "envelope": {"client_op_id": "00000000-0000-0000-0000-0000000000f1",
        "base_seq": seq + 1, "actor": 1, "op": bad}, "concurrent": []}));
    assert_eq!(r["ok"], json!(false), "{r}");

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

/// The collaboration server refuses a browser whose engine_commit differs from
/// its host's, so `vgrid collab-host` and the WASM engine must report the same
/// identifier when built from the same source. Both build scripts include
/// build-support/engine_commit.rs; this pins that they agree.
#[test]
fn hello_engine_commit_matches_the_wasm_engine() {
    let mut host = Host::spawn();
    let hello = host.call(serde_json::json!({"cmd": "hello"}));
    let host_commit = hello["engine_commit"].as_str().expect("hello carries engine_commit");
    let wasm_commit = visigrid_engine_wasm::engine_commit();
    assert_eq!(host_commit, wasm_commit, "collab-host and the WASM engine must report the same engine_commit");
    // In a git checkout (every CI and Docker build) it identifies a commit.
    if std::path::Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../.git")).exists() {
        let sha = host_commit.trim_end_matches("-modified");
        assert_eq!(sha.len(), 40, "expected a full commit SHA, got {host_commit:?}");
        assert!(sha.chars().all(|c| c.is_ascii_hexdigit()));
    }
}

#[test]
fn large_sheets_snapshot_as_bands_and_load_back() {
    // 125,000 rows x 2 = 250,000 cells: over the band threshold.
    let cells: Vec<Value> = (0..125_000)
        .flat_map(|r| [json!({"row": r, "col": 0, "value": r % 89}), json!({"row": r, "col": 1, "value": format!("r{r}")})])
        .chain([json!({"row": 0, "col": 2, "formula": "=SUM(A1:A125000)"})])
        .collect();
    let doc = json!({"format": "visigrid-json", "version": 2, "sheets": [{"name": "Big", "cells": cells}], "active_sheet": 0});
    let mut a = Host::spawn();
    let loaded = a.call(json!({"cmd": "load", "document": doc, "seq": 3}));
    assert_eq!(loaded["ok"], json!(true), "{loaded}");
    let snap = a.call(json!({"cmd": "snapshot"}));
    let bands = snap["bands"].as_array().expect("bands listed").clone();
    assert_eq!(bands.len(), 8, "{bands:?}");
    assert!(snap["document"]["sheets"][0]["cells"].is_null(), "the manifest carries no cells");
    assert!(snap["document"].to_string().len() < 4096);

    let mut b = Host::spawn();
    assert_eq!(b.call(json!({"cmd": "load", "document": snap["document"], "seq": 3}))["ok"], json!(true));
    for band in &bands {
        let got = a.call(json!({"cmd": "band", "key": band["key"]}));
        let put = b.call(json!({"cmd": "load_band", "key": band["key"], "data": got["data"]}));
        assert_eq!(put["ok"], json!(true), "{put}");
    }
    assert_eq!(b.call(json!({"cmd": "finish_load"}))["ok"], json!(true));
    // Too large to checksum (empty); content-addressed bands compare instead.
    assert_eq!(snap["checksum"], json!(""));
    let again = b.call(json!({"cmd": "snapshot"}));
    assert_eq!(again["bands"], snap["bands"], "the banded load is the same workbook");
    assert_eq!(again["document"], snap["document"]);
    assert_eq!(b.call(json!({"cmd": "band", "key": "nope"}))["ok"], json!(false));
    let bad = b.call(json!({"cmd": "load_band", "key": bands[0]["key"], "data": "AAAA"}));
    assert_eq!(bad["ok"], json!(false), "a band is checked against its key");
}
