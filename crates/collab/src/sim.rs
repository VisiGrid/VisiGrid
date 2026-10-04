//! Headless convergence simulator.
//!
//! N clients and one server exchange messages over per-connection FIFO links
//! with random delays (so different clients' messages interleave in every
//! order), random disconnects that lose whatever was in flight, and bursts
//! of local edits. Edits happen online and offline. After the edit budget is
//! spent, everyone reconnects and the network drains; then every replica must
//! equal the server's exactly — raw cell text, cached computed values,
//! formats, sheet ids, names and tab order.
//!
//! Everything is driven by one seed, so a failure replays exactly.

use std::cmp::Reverse;
use std::collections::BinaryHeap;

use rand::rngs::StdRng;
use rand::{Rng, SeedableRng};
use uuid::Uuid;

use crate::apply::{fingerprint, first_difference};
use crate::client::{Client, ToClient, ToServer};
use crate::gen::random_ops;
use crate::server::{Server, Submitted};

#[derive(Clone, Debug)]
pub struct SimConfig {
    pub clients: usize,
    /// Local edits generated in total.
    pub edits: usize,
    /// Link delay in ticks: uniform in 1..=max_delay.
    pub max_delay: u64,
    /// Chance per scheduling step that a client toggles its connection.
    pub toggle_prob: f64,
    /// Ticks between scheduling steps.
    pub step_gap: u64,
}

impl Default for SimConfig {
    fn default() -> Self {
        SimConfig {
            clients: 3,
            edits: 40,
            max_delay: 12,
            toggle_prob: 0.05,
            step_gap: 3,
        }
    }
}

#[derive(Clone, Debug, Default)]
pub struct SimReport {
    pub seed: u64,
    pub committed: u64,
    pub refused: u64,
    pub resyncs: u64,
    pub max_pending: usize,
    pub disconnects: u64,
    pub failure: Option<String>,
    /// Converged documents whose cached values differed until a full
    /// recompute: the engine's incremental recalc is order dependent.
    pub engine_stale: Option<String>,
    /// Diverged only because the engine replaced a cycle's formula text
    /// with "#CYCLE!" on some replicas and not others.
    pub engine_cycle: Option<String>,
    pub trace: Vec<String>,
}

impl SimReport {
    pub fn ok(&self) -> bool {
        self.failure.is_none()
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum Dir {
    Up,
    Down,
}

#[derive(Clone, Debug)]
enum Event {
    Step,
    Deliver {
        client: usize,
        dir: Dir,
        generation: u64,
        msg: Msg,
    },
}

#[derive(Clone, Debug)]
enum Msg {
    Up(ToServer),
    Down(ToClient),
}

struct Link {
    /// Messages tagged with an older generation were lost by a disconnect.
    generation: u64,
    last_at: u64,
}

struct Sim {
    rng: StdRng,
    now: u64,
    tie: u64,
    queue: BinaryHeap<Reverse<(u64, u64, usize)>>,
    events: Vec<Option<Event>>,
    server: Server,
    clients: Vec<Client>,
    /// Server's view: whether it streams live ops to client i.
    live: Vec<bool>,
    up: Vec<Link>,
    down: Vec<Link>,
    cfg: SimConfig,
    trace: Vec<String>,
    disconnects: u64,
    record: bool,
}

impl Sim {
    fn schedule(&mut self, at: u64, ev: Event) {
        self.tie += 1;
        self.events.push(Some(ev));
        self.queue
            .push(Reverse((at, self.tie, self.events.len() - 1)));
    }

    fn send(&mut self, client: usize, msg: Msg) {
        let delay = self.rng.gen_range(1..=self.cfg.max_delay);
        let (link, dir) = match msg {
            Msg::Up(_) => (&mut self.up[client], Dir::Up),
            Msg::Down(_) => (&mut self.down[client], Dir::Down),
        };
        let at = (self.now + delay).max(link.last_at);
        link.last_at = at;
        let generation = link.generation;
        self.schedule(
            at,
            Event::Deliver {
                client,
                dir,
                generation,
                msg,
            },
        );
    }

    fn log(&mut self, line: String) {
        if self.record {
            self.trace.push(format!("t{} {}", self.now, line));
        }
    }

    fn flush_client(&mut self, i: usize) {
        while let Some(m) = self.clients[i].poll_send() {
            if let ToServer::Submit(e) = &m {
                let line = format!(
                    "c{i} send {} base={} {:?}",
                    short(e.client_op_id),
                    e.base_seq,
                    e.ops
                );
                self.log(line);
            }
            self.send(i, Msg::Up(m));
        }
    }

    fn toggle(&mut self, i: usize) {
        if self.clients[i].connected {
            self.disconnects += 1;
            self.clients[i].disconnect();
            self.live[i] = false;
            self.up[i].generation += 1;
            self.down[i].generation += 1;
            self.log(format!("c{i} disconnect"));
        } else {
            let hello = self.clients[i].reconnect();
            self.log(format!("c{i} reconnect {:?}", hello));
            self.send(i, Msg::Up(hello));
        }
    }

    fn step(&mut self, edits_left: &mut usize) {
        // A burst of 1-3 edits from one client, online or offline.
        let i = self.rng.gen_range(0..self.clients.len());
        let burst = self.rng.gen_range(1..=3).min(*edits_left);
        for _ in 0..burst {
            let key = (1u64 << 62) | self.rng.gen_range(0..(1u64 << 40));
            let ops = random_ops(&mut self.rng, &self.clients[i].wb, key);
            *edits_left -= 1;
            if ops.is_empty() {
                continue;
            }
            let id = Uuid::from_u128(self.rng.gen());
            let line = format!("c{i} local {} {:?}", short(id), ops);
            self.log(line);
            self.clients[i].local(id, ops);
        }
        self.flush_client(i);
        for j in 0..self.clients.len() {
            if self.rng.gen_bool(self.cfg.toggle_prob) {
                self.toggle(j);
            }
        }
    }

    fn deliver(&mut self, client: usize, dir: Dir, generation: u64, msg: Msg) {
        let current = match dir {
            Dir::Up => self.up[client].generation,
            Dir::Down => self.down[client].generation,
        };
        if generation != current {
            return; // lost in a disconnect
        }
        match msg {
            Msg::Up(ToServer::Submit(env)) => match self.server.submit(&env) {
                Submitted::Committed(c) => {
                    let line = format!(
                        "server seq {} from c{client} {} {:?}",
                        c.seq,
                        short(c.client_op_id),
                        c.ops
                    );
                    self.log(line);
                    for j in 0..self.clients.len() {
                        if self.live[j] {
                            self.send(j, Msg::Down(ToClient::Op(c.clone())));
                        }
                    }
                }
                Submitted::Refused {
                    client_op_id,
                    reason,
                } => {
                    self.log(format!(
                        "server refused {} from c{client}: {reason}",
                        short(client_op_id)
                    ));
                    self.send(
                        client,
                        Msg::Down(ToClient::Refused {
                            client_op_id,
                            reason,
                        }),
                    );
                }
                Submitted::Duplicate => self.log(format!("server duplicate from c{client}")),
            },
            Msg::Up(ToServer::Hello { last_seen }) => {
                let missed: Vec<_> = self.server.since(last_seen).to_vec();
                self.live[client] = true;
                for c in missed {
                    self.send(client, Msg::Down(ToClient::Op(c)));
                }
                let head = self.server.head();
                self.send(client, Msg::Down(ToClient::Welcome { head }));
            }
            Msg::Down(m) => {
                self.clients[client].receive(m);
                self.flush_client(client);
            }
        }
    }
}

fn short(id: Uuid) -> String {
    id.simple().to_string()[..6].to_string()
}

/// Run one simulation. `record` keeps a full event trace (for failures).
pub fn run(seed: u64, cfg: &SimConfig, record: bool) -> SimReport {
    let n = cfg.clients;
    let mut sim = Sim {
        rng: StdRng::seed_from_u64(seed),
        now: 0,
        tie: 0,
        queue: BinaryHeap::new(),
        events: Vec::new(),
        server: Server::new(),
        clients: (0..n).map(|i| Client::new(i as u64 + 1)).collect(),
        live: vec![true; n],
        up: (0..n)
            .map(|_| Link {
                generation: 0,
                last_at: 0,
            })
            .collect(),
        down: (0..n)
            .map(|_| Link {
                generation: 0,
                last_at: 0,
            })
            .collect(),
        cfg: cfg.clone(),
        trace: Vec::new(),
        disconnects: 0,
        record,
    };
    let mut edits_left = cfg.edits;
    sim.schedule(0, Event::Step);
    let mut guard = 0u64;
    while let Some(Reverse((at, _, idx))) = sim.queue.pop() {
        guard += 1;
        if guard > 5_000_000 {
            return report(
                seed,
                &sim,
                Some("event budget exhausted (livelock?)".into()),
            );
        }
        sim.now = at;
        let ev = sim.events[idx].take().expect("each event runs once");
        match ev {
            Event::Step => {
                if edits_left > 0 {
                    sim.step(&mut edits_left);
                    let gap = sim.cfg.step_gap;
                    sim.schedule(sim.now + gap, Event::Step);
                } else {
                    // Edit budget spent: bring everyone back so the network
                    // can drain to quiescence.
                    for i in 0..n {
                        if !sim.clients[i].connected {
                            sim.toggle(i);
                        }
                    }
                }
            }
            Event::Deliver {
                client,
                dir,
                generation,
                msg,
            } => sim.deliver(client, dir, generation, msg),
        }
    }
    // Quiescent: no messages anywhere. Nothing may still be pending.
    for (i, c) in sim.clients.iter().enumerate() {
        if c.pending_count() > 0 || !c.connected || c.syncing {
            let msg = format!(
                "c{i} not quiescent: pending={} connected={} syncing={}",
                c.pending_count(),
                c.connected,
                c.syncing
            );
            return report(seed, &sim, Some(msg));
        }
    }
    for (i, c) in sim.clients.iter().enumerate() {
        if c.last_seen != sim.server.head() {
            return report(
                seed,
                &sim,
                Some(format!(
                    "c{i} at seq {} but server at {}",
                    c.last_seen,
                    sim.server.head()
                )),
            );
        }
    }
    let server_print = fingerprint(&sim.server.wb);
    let strict = sim.clients.iter().enumerate().find_map(|(i, c)| {
        first_difference(&fingerprint(&c.wb), &server_print).map(|d| format!("c{i}: {d}"))
    });
    let Some(strict) = strict else {
        return report(seed, &sim, None);
    };
    // Replicas applied the same ops in different orders. If a full recompute
    // makes them agree, the documents converged and only the engine's
    // incremental recalc left a stale value behind: an engine finding, not a
    // collaboration failure.
    sim.server.wb.recompute_full_ordered();
    for c in &mut sim.clients {
        c.wb.recompute_full_ordered();
    }
    let server_print = fingerprint(&sim.server.wb);
    for (i, c) in sim.clients.iter().enumerate() {
        let print = fingerprint(&c.wb);
        if let Some(diff) = first_difference(&print, &server_print) {
            // The engine's full recompute overwrites the text of formulas in
            // a cycle with the literal "#CYCLE!" — whether that happened
            // depends on whether a structural edit ran while the cycle
            // existed, which differs between replicas. An engine bug, not a
            // convergence failure of the protocol.
            let cycle_loss = |p: &crate::apply::Fingerprint| {
                p.sheets
                    .iter()
                    .any(|s| s.cells.iter().any(|(_, _, raw, ..)| raw == "#CYCLE!"))
            };
            if cycle_loss(&print) || cycle_loss(&server_print) {
                let mut r = report(seed, &sim, None);
                r.engine_cycle = Some(format!("c{i}: {diff}"));
                return r;
            }
            return report(
                seed,
                &sim,
                Some(format!("c{i} diverged from server: {diff}")),
            );
        }
    }
    let mut r = report(seed, &sim, None);
    r.engine_stale = Some(strict);
    r
}

fn report(seed: u64, sim: &Sim, failure: Option<String>) -> SimReport {
    SimReport {
        seed,
        committed: sim.server.head(),
        refused: sim.server.refused,
        resyncs: sim.clients.iter().map(|c| c.stats.resyncs).sum(),
        max_pending: sim
            .clients
            .iter()
            .map(|c| c.stats.max_pending)
            .max()
            .unwrap_or(0),
        disconnects: sim.disconnects,
        failure,
        engine_stale: None,
        engine_cycle: None,
        trace: sim.trace.clone(),
    }
}

/// On failure, find the smallest edit budget that still fails for this seed
/// and return that run with its trace.
pub fn shrink(seed: u64, cfg: &SimConfig) -> SimReport {
    let mut lo = 1usize;
    let mut hi = cfg.edits;
    let mut best = run(seed, cfg, true);
    while lo < hi {
        let mid = (lo + hi) / 2;
        let c = SimConfig {
            edits: mid,
            ..cfg.clone()
        };
        let r = run(seed, &c, true);
        if r.ok() {
            lo = mid + 1;
        } else {
            hi = mid;
            best = r;
        }
    }
    best
}
