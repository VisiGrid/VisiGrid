//! Drive a real collaboration server (protocol v2 over WebSocket) with the
//! Phase 1 client replicas and check convergence.
//!
//! ```text
//! cargo run -p visigrid-collab --features remote --bin collab-sim-remote -- \
//!     --url ws://127.0.0.1:5159 --sheet <pid> --token <jwt> \
//!     --clients 4 --seed 7 --edits 200
//! ```
//!
//! The sheet must start as the blank workbook (one sheet, "Sheet1", no
//! cells): every replica here starts from `Workbook::new()` and catches up
//! from seq 0. The token is sent as `Authorization: Bearer`. A reconnecting
//! client whose last seq has been folded gets `snapshot_url`: the harness
//! downloads it, loads it (with `collab_sheet_ids`), and applies it as a
//! document replacement at `snapshot_seq`.
//!
//! After quiescence (no pending edits, no traffic for `--quiet-ms`):
//! 1. every client has the same head and an identical fingerprint;
//! 2. every `checksum{seq, checksum}` the server published equals the
//!    checksum of the committed log replayed up to that seq;
//! 3. no client saw an `error` or an unexpected `resync`.
//!
//! Edits come from the same conflict-weighted generator as the in-memory
//! simulator, seeded per client. Network timing is real, so a failure is
//! reported with its seed and per-client traces but is not exactly
//! replayable.

use std::collections::BTreeMap;
use std::net::TcpStream;
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

use rand::rngs::StdRng;
use rand::{Rng, SeedableRng};
use serde_json::{json, Value};
use tungstenite::client::IntoClientRequest;
use tungstenite::stream::MaybeTlsStream;
use tungstenite::{Message, WebSocket};
use uuid::Uuid;
use visigrid_collab::apply::{apply_ops, checksum, checksum_of, fingerprint, first_difference, Fingerprint};
use visigrid_collab::client::{Client, ToClient, ToServer};
use visigrid_collab::gen::random_ops;
use visigrid_collab::op::{ops_from_json, ops_to_json};
use visigrid_collab::server::Committed;
use visigrid_engine::workbook::Workbook;

#[derive(Clone, Debug)]
struct Args {
    url: String,
    sheet: String,
    token: String,
    clients: usize,
    seed: u64,
    edits: usize,
    max_delay_ms: u64,
    toggle_prob: f64,
    quiet_ms: u64,
    timeout_s: u64,
    engine_commit: String,
    legacy_refusal: bool,
}

fn default_engine_commit() -> String {
    std::env::var("COLLAB_ENGINE_COMMIT").unwrap_or_else(|_| {
        option_env!("VISIGRID_ENGINE_COMMIT")
            .or(option_env!("GIT_COMMIT"))
            .unwrap_or(env!("CARGO_PKG_VERSION"))
            .to_string()
    })
}

fn parse_args() -> Result<Args, String> {
    let mut a = Args {
        url: String::new(),
        sheet: String::new(),
        token: std::env::var("COLLAB_TOKEN").unwrap_or_default(),
        clients: 3,
        seed: 1,
        edits: 100,
        max_delay_ms: 15,
        toggle_prob: 0.02,
        quiet_ms: 1500,
        timeout_s: 120,
        engine_commit: default_engine_commit(),
        legacy_refusal: false,
    };
    let mut it = std::env::args().skip(1);
    while let Some(flag) = it.next() {
        let mut val = || it.next().ok_or(format!("{flag} needs a value"));
        match flag.as_str() {
            "--url" => a.url = val()?,
            "--sheet" => a.sheet = val()?,
            "--token" => a.token = val()?,
            "--clients" => a.clients = val()?.parse().map_err(|_| "bad --clients")?,
            "--seed" => a.seed = val()?.parse().map_err(|_| "bad --seed")?,
            "--edits" => a.edits = val()?.parse().map_err(|_| "bad --edits")?,
            "--max-delay-ms" => a.max_delay_ms = val()?.parse().map_err(|_| "bad --max-delay-ms")?,
            "--toggle-prob" => a.toggle_prob = val()?.parse().map_err(|_| "bad --toggle-prob")?,
            "--quiet-ms" => a.quiet_ms = val()?.parse().map_err(|_| "bad --quiet-ms")?,
            "--timeout-s" => a.timeout_s = val()?.parse().map_err(|_| "bad --timeout-s")?,
            "--engine-commit" => a.engine_commit = val()?,
            "--legacy-refusal" => a.legacy_refusal = true,
            "-h" | "--help" => {
                return Err("usage: collab-sim-remote --url ws://HOST:PORT --sheet PID --token JWT \
[--clients N] [--seed S] [--edits E] [--max-delay-ms D] [--toggle-prob P] \
[--quiet-ms Q] [--timeout-s T] [--engine-commit C]"
                    .into())
            }
            other => return Err(format!("unknown flag {other}")),
        }
    }
    if a.url.is_empty() || a.sheet.is_empty() || a.token.is_empty() {
        return Err("--url, --sheet and --token (or COLLAB_TOKEN) are required".into());
    }
    Ok(a)
}

type Socket = WebSocket<MaybeTlsStream<TcpStream>>;

fn connect(args: &Args) -> Result<Socket, String> {
    let url = format!("{}/api/sheets/{}/collab", args.url.trim_end_matches('/'), args.sheet);
    let mut req = url.as_str().into_client_request().map_err(|e| format!("bad url {url}: {e}"))?;
    req.headers_mut().insert(
        "Authorization",
        format!("Bearer {}", args.token).parse().map_err(|_| "token is not a valid header")?,
    );
    let (ws, _) = tungstenite::connect(req).map_err(|e| format!("connect {url}: {e}"))?;
    if let MaybeTlsStream::Plain(s) = ws.get_ref() {
        s.set_read_timeout(Some(Duration::from_millis(5))).ok();
        s.set_nodelay(true).ok();
    }
    Ok(ws)
}

/// What one client thread reports.
struct Outcome {
    index: usize,
    last_seen: u64,
    print: Fingerprint,
    committed: Vec<Committed>,
    checksums: BTreeMap<u64, String>,
    stats: String,
    failure: Option<String>,
    trace: Vec<String>,
}

struct Peer {
    index: usize,
    args: Args,
    rng: StdRng,
    client: Client,
    ws: Option<Socket>,
    checksums: BTreeMap<u64, String>,
    trace: Vec<String>,
    last_traffic: Instant,
}

impl Peer {
    fn log(&mut self, line: String) {
        if self.trace.len() < 5000 {
            self.trace.push(line);
        }
    }

    fn send(&mut self, v: Value) -> Result<(), String> {
        let ws = self.ws.as_mut().ok_or("not connected")?;
        ws.send(Message::text(v.to_string())).map_err(|e| format!("send: {e}"))?;
        self.last_traffic = Instant::now();
        Ok(())
    }

    fn hello(&mut self) -> Result<(), String> {
        let last = self.client.last_seen;
        let msg = json!({"type": "hello", "protocol": 2, "engine_commit": self.args.engine_commit, "last_seq": last});
        self.log(format!("hello last_seq={last}"));
        self.send(msg)
    }

    fn flush(&mut self) -> Result<(), String> {
        while let Some(m) = self.client.poll_send() {
            if let ToServer::Submit(env) = m {
                if self.args.max_delay_ms > 0 {
                    let d = self.rng.gen_range(0..=self.args.max_delay_ms);
                    std::thread::sleep(Duration::from_millis(d));
                }
                self.log(format!("send {} base={} {:?}", short(env.client_op_id), env.base_seq, env.ops));
                self.send(json!({
                    "type": "op",
                    "envelope": {
                        "client_op_id": env.client_op_id,
                        "base_seq": env.base_seq,
                        "op": ops_to_json(&env.ops),
                    }
                }))?;
            }
        }
        Ok(())
    }

    fn committed_from(v: &Value) -> Result<Committed, String> {
        Ok(Committed {
            seq: v.get("seq").and_then(Value::as_u64).ok_or("op without seq")?,
            client_op_id: v
                .get("client_op_id")
                .and_then(Value::as_str)
                .and_then(|s| Uuid::parse_str(s).ok())
                .ok_or("op without client_op_id")?,
            actor: v.get("actor").and_then(Value::as_u64).unwrap_or(0),
            ops: ops_from_json(v.get("op").ok_or("op without op")?)?,
        })
    }

    /// Download and load the snapshot a welcome points at.
    fn load_snapshot(&self, path: &str) -> Result<Workbook, String> {
        let http = self.args.url.replacen("ws://", "http://", 1);
        let base = url::Url::parse(&http).map_err(|e| format!("bad url: {e}"))?;
        let host = base.host_str().ok_or("url without host")?;
        let port = base.port().unwrap_or(80);
        let body = http_get(host, port, path, &self.args.token)?;
        let v: Value = serde_json::from_slice(&body).map_err(|e| format!("snapshot is not JSON: {e}"))?;
        let doc = match v.get("transport").and_then(Value::as_str) {
            Some("inline") => v.get("document").cloned().ok_or("inline snapshot without document")?,
            Some(other) => return Err(format!("snapshot transport {other} not supported by the harness")),
            None => v,
        };
        import_document(&doc)
    }

    fn handle(&mut self, text: &str) -> Result<(), String> {
        let v: Value = serde_json::from_str(text).map_err(|e| format!("server sent invalid JSON: {e}"))?;
        let kind = v.get("type").and_then(Value::as_str).unwrap_or("");
        match kind {
            "welcome" => {
                if let Some(path) = v.get("snapshot_url").and_then(Value::as_str) {
                    let at = v.get("snapshot_seq").and_then(Value::as_u64).ok_or("snapshot without snapshot_seq")?;
                    let wb = self.load_snapshot(path)?;
                    self.log(format!("snapshot seq={at}"));
                    self.client.replace_document(wb, at);
                }
                let ops = v.get("ops").and_then(Value::as_array).cloned().unwrap_or_default();
                for o in &ops {
                    let c = Self::committed_from(o)?;
                    self.check_order(&c)?;
                    self.client.receive(ToClient::Op(c));
                }
                let head = v.get("seq").and_then(Value::as_u64).ok_or("welcome without seq")?;
                self.log(format!("welcome seq={head} ops={}", ops.len()));
                self.client.receive(ToClient::Welcome { head });
            }
            "op" => {
                let c = Self::committed_from(&v)?;
                self.check_order(&c)?;
                self.log(format!("op seq={} from {} {:?}", c.seq, short(c.client_op_id), c.ops));
                self.client.receive(ToClient::Op(c));
            }
            "ack" => {
                let id = uuid_of(&v)?;
                let seq = v.get("seq").and_then(Value::as_u64).ok_or("ack without seq")?;
                self.log(format!("ack {} seq={seq}", short(id)));
                if seq == self.client.last_seen + 1 {
                    self.client.ack(id, seq);
                } else if seq > self.client.last_seen + 1 {
                    return Err(format!("ack seq {seq} skipped ops after {}", self.client.last_seen));
                }
            }
            "rejected" => {
                let id = uuid_of(&v)?;
                let result = v.get("result").and_then(Value::as_str).unwrap_or("");
                let reason = v.get("reason").and_then(Value::as_str).unwrap_or("").to_string();
                self.log(format!("rejected {} {result}: {reason}", short(id)));
                match result {
                    "dropped" => self.client.dropped(id),
                    "refused" => self.client.receive(ToClient::Refused { client_op_id: id, reason }),
                    other => return Err(format!("rejected with unknown result {other:?}")),
                }
            }
            "checksum" => {
                let seq = v.get("seq").and_then(Value::as_u64).ok_or("checksum without seq")?;
                let sum = v.get("checksum").and_then(Value::as_str).ok_or("checksum without value")?;
                if let Some(prev) = self.checksums.insert(seq, sum.to_string()) {
                    if prev != sum {
                        return Err(format!("server sent two checksums for seq {seq}"));
                    }
                }
            }
            "presence" | "presence_leave" | "pong" => {}
            "resync" => return Err(format!("unexpected resync: {}", v.get("reason").unwrap_or(&Value::Null))),
            "error" => return Err(format!("server error: {v}")),
            other => return Err(format!("unknown message type {other:?}: {v}")),
        }
        Ok(())
    }

    fn check_order(&self, c: &Committed) -> Result<(), String> {
        if c.seq > self.client.last_seen + 1 {
            return Err(format!("op seq {} arrived after {} (gap)", c.seq, self.client.last_seen));
        }
        Ok(())
    }

    /// Read whatever has arrived. Returns false when the socket closed.
    fn pump(&mut self) -> Result<bool, String> {
        loop {
            let Some(ws) = self.ws.as_mut() else { return Ok(false) };
            match ws.read() {
                Ok(Message::Text(t)) => {
                    self.last_traffic = Instant::now();
                    let text = t.to_string();
                    self.handle(&text)?;
                    self.flush()?;
                }
                Ok(Message::Close(_)) => return Ok(false),
                Ok(_) => {}
                Err(tungstenite::Error::Io(e))
                    if matches!(e.kind(), std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut) =>
                {
                    return Ok(true)
                }
                Err(tungstenite::Error::ConnectionClosed | tungstenite::Error::AlreadyClosed) => return Ok(false),
                Err(e) => return Err(format!("read: {e}")),
            }
        }
    }

    fn disconnect(&mut self) {
        if let Some(mut ws) = self.ws.take() {
            let _ = ws.close(None);
            let _ = ws.flush();
        }
        self.client.disconnect();
        self.log("disconnect".into());
    }

    fn reconnect(&mut self) -> Result<(), String> {
        self.ws = Some(connect(&self.args)?);
        let _ = self.client.reconnect();
        self.hello()
    }

    fn run(&mut self, edits: usize) -> Result<(), String> {
        self.reconnect()?;
        let deadline = Instant::now() + Duration::from_secs(self.args.timeout_s);
        let mut left = edits;
        while left > 0 {
            if Instant::now() > deadline {
                return Err("timed out while editing".into());
            }
            if self.ws.is_some() && !self.pump()? {
                self.disconnect();
            }
            if self.ws.is_none() && self.rng.gen_bool(0.3) {
                self.reconnect()?;
            }
            // A burst of 1-3 edits, online or offline.
            let burst = self.rng.gen_range(1..=3).min(left);
            for _ in 0..burst {
                left -= 1;
                let key = (1u64 << 62) | ((self.index as u64) << 40) | self.rng.gen_range(0..(1u64 << 40));
                let ops = random_ops(&mut self.rng, &self.client.wb, key);
                if ops.is_empty() {
                    continue;
                }
                let id = Uuid::from_u128(self.rng.gen());
                self.log(format!("local {} {:?}", short(id), ops));
                self.client.local(id, ops);
            }
            if self.ws.is_some() {
                self.flush()?;
                if self.rng.gen_bool(self.args.toggle_prob) {
                    self.disconnect();
                }
            }
            let pause = self.rng.gen_range(0..=self.args.max_delay_ms.max(1));
            std::thread::sleep(Duration::from_millis(pause));
        }
        // Quiesce: connected, nothing pending, and quiet for quiet_ms.
        if self.ws.is_none() {
            self.reconnect()?;
        }
        loop {
            if Instant::now() > deadline {
                return Err(format!(
                    "timed out waiting for quiescence (pending {}, syncing {})",
                    self.client.pending_count(),
                    self.client.syncing
                ));
            }
            if !self.pump()? {
                self.disconnect();
                self.reconnect()?;
            }
            self.flush()?;
            let quiet = self.last_traffic.elapsed() >= Duration::from_millis(self.args.quiet_ms);
            if quiet && self.client.pending_count() == 0 && !self.client.syncing {
                return Ok(());
            }
        }
    }
}

fn uuid_of(v: &Value) -> Result<Uuid, String> {
    v.get("client_op_id")
        .and_then(Value::as_str)
        .and_then(|s| Uuid::parse_str(s).ok())
        .ok_or_else(|| format!("message without client_op_id: {v}"))
}

fn short(id: Uuid) -> String {
    id.simple().to_string()[..6].to_string()
}

fn main() {
    let args = match parse_args() {
        Ok(a) => a,
        Err(e) => {
            eprintln!("{e}");
            std::process::exit(2);
        }
    };
    let per_client = args.edits.div_ceil(args.clients.max(1));
    let outcomes = Arc::new(Mutex::new(Vec::new()));
    let mut threads = Vec::new();
    for index in 0..args.clients {
        let args = args.clone();
        let outcomes = Arc::clone(&outcomes);
        threads.push(std::thread::spawn(move || {
            let mut peer = Peer {
                index,
                rng: StdRng::seed_from_u64(args.seed.wrapping_mul(1_000_003).wrapping_add(index as u64)),
                client: {
                    let mut c = Client::new(index as u64 + 1);
                    c.legacy_refusal = args.legacy_refusal;
                    c
                },
                ws: None,
                checksums: BTreeMap::new(),
                trace: Vec::new(),
                last_traffic: Instant::now(),
                args,
            };
            let failure = peer.run(per_client).err();
            peer.disconnect();
            let s = &peer.client.stats;
            let stats = format!(
                "local={} acked={} refused={} discarded={} replaced_kept={} replaced_dropped={} resyncs={} max_pending={}",
                s.local_envelopes,
                s.acked,
                s.refused_envelopes,
                s.discarded_after_refusal,
                s.kept_across_replacement,
                s.discarded_by_replacement,
                s.resyncs,
                s.max_pending
            );
            outcomes.lock().unwrap().push(Outcome {
                index,
                last_seen: peer.client.last_seen,
                print: fingerprint(&peer.client.wb),
                committed: peer.client.committed.clone(),
                checksums: peer.checksums,
                stats,
                failure,
                trace: peer.trace,
            });
        }));
    }
    for t in threads {
        let _ = t.join();
    }
    let mut outcomes = Arc::try_unwrap(outcomes).ok().unwrap().into_inner().unwrap();
    outcomes.sort_by_key(|o| o.index);
    std::process::exit(verify(&args, &outcomes));
}

fn verify(args: &Args, outcomes: &[Outcome]) -> i32 {
    let mut failures = Vec::new();
    for o in outcomes {
        if let Some(f) = &o.failure {
            failures.push(format!("client {}: {f}", o.index));
        }
    }
    let head = outcomes.iter().map(|o| o.last_seen).max().unwrap_or(0);
    let reference = outcomes.iter().find(|o| o.last_seen == head);
    // Checksums are replayed from the blank workbook, so they need a client
    // that saw the whole log (one that loaded a snapshot did not).
    let complete = outcomes.iter().find(|o| o.last_seen == head && o.committed.len() as u64 == head);
    if let Some(r) = reference {
        for o in outcomes {
            if o.last_seen != head {
                failures.push(format!("client {} stopped at seq {} (head {head})", o.index, o.last_seen));
            } else if let Some(diff) = first_difference(&r.print, &o.print) {
                failures.push(format!("client {} diverged from client {}: {diff}", o.index, r.index));
            }
        }
        // Every published server checksum must equal the committed log
        // replayed to that seq from the blank workbook.
        let mut published: BTreeMap<u64, String> = BTreeMap::new();
        for o in outcomes {
            for (seq, sum) in &o.checksums {
                if let Some(prev) = published.insert(*seq, sum.clone()) {
                    if &prev != sum {
                        failures.push(format!("clients received different checksums for seq {seq}"));
                    }
                }
            }
        }
        let mut wb = Workbook::new();
        let mut applied = 0u64;
        let replay_from = complete.unwrap_or(r);
        let mut verified = 0usize;
        for (seq, sum) in &published {
            if complete.is_none() {
                break;
            }
            if *seq > head {
                failures.push(format!("checksum for seq {seq} beyond head {head}"));
                continue;
            }
            while applied < *seq {
                let c = &replay_from.committed[applied as usize];
                apply_ops(&mut wb, &c.ops);
                applied += 1;
            }
            let ours = checksum(&wb);
            if &ours != sum {
                failures.push(format!("server checksum at seq {seq} is {sum}, replicas compute {ours}"));
            }
            verified += 1;
        }
        // The final state must match the server's last published checksum.
        if let Some((seq, sum)) = published.iter().next_back() {
            if *seq == head && &checksum_of(&r.print) != sum {
                failures.push(format!("final state differs from the server's checksum at seq {seq}"));
            }
        }
        println!(
            "seed {} clients {} head {} server checksums verified {} of {}",
            args.seed,
            outcomes.len(),
            head,
            verified,
            published.len()
        );
    }
    for o in outcomes {
        println!("client {}: last_seen={} {}", o.index, o.last_seen, o.stats);
    }
    if failures.is_empty() {
        println!("CONVERGED");
        return 0;
    }
    println!("FAILED (seed {}):", args.seed);
    for f in &failures {
        println!("  {f}");
    }
    for o in outcomes {
        println!("--- trace client {} (last 60 events)", o.index);
        for line in o.trace.iter().rev().take(60).collect::<Vec<_>>().into_iter().rev() {
            println!("  {line}");
        }
    }
    1
}

/// GET over HTTP/1.0 (the server then closes instead of chunking).
fn http_get(host: &str, port: u16, path: &str, token: &str) -> Result<Vec<u8>, String> {
    use std::io::{Read, Write};
    let mut s = TcpStream::connect((host, port)).map_err(|e| format!("snapshot connect: {e}"))?;
    s.set_read_timeout(Some(Duration::from_secs(30))).ok();
    write!(s, "GET {path} HTTP/1.0\r\nHost: {host}:{port}\r\nAuthorization: Bearer {token}\r\n\r\n")
        .map_err(|e| format!("snapshot request: {e}"))?;
    let mut raw = Vec::new();
    s.read_to_end(&mut raw).map_err(|e| format!("snapshot read: {e}"))?;
    let split = raw.windows(4).position(|w| w == b"\r\n\r\n").ok_or("snapshot: no HTTP header end")?;
    let head = String::from_utf8_lossy(&raw[..split]);
    if !head.starts_with("HTTP/1.0 200") && !head.starts_with("HTTP/1.1 200") {
        return Err(format!("snapshot: {}", head.lines().next().unwrap_or("")));
    }
    Ok(raw[split + 4..].to_vec())
}

/// visigrid-json to a workbook with stable sheet ids (`collab_sheet_ids`),
/// as the engine host loads it.
fn import_document(doc: &Value) -> Result<Workbook, String> {
    let (mut wb, _, _) = visigrid_io::json::import_any(&doc.to_string())?;
    if let Some(ids) = doc.get("collab_sheet_ids") {
        let ids: Vec<u64> = serde_json::from_value(ids.clone()).map_err(|_| "bad collab_sheet_ids")?;
        if ids.len() != wb.sheets().len() {
            return Err("collab_sheet_ids does not match the sheet count".into());
        }
        for (i, id) in ids.iter().enumerate() {
            wb.sheet_mut(i).ok_or("sheet index")?.id = visigrid_engine::sheet::SheetId(*id);
        }
        let next = ids.iter().copied().max().unwrap_or(0) + 1;
        wb.set_next_sheet_id(next.max(wb.next_sheet_id()));
    }
    Ok(wb)
}
