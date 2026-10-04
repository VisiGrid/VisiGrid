//! A client replica: the real engine, plus edits not yet acknowledged.
//!
//! Local edits apply immediately. At most one envelope is in flight; the
//! rest wait in a buffer (the ot.js / Wave pattern), so the server never has
//! to transform an op against its own sender's unacknowledged op.
//!
//! When a sequenced op from someone else arrives, every pending envelope is
//! transformed past it (they will be sequenced after it) and the op is
//! transformed past them; the transformed op is then applied on top of the
//! local state. TP1 makes that equal to "committed + their op + our ops".
//!
//! If a pending envelope is refused (a V1-serialized conflict), that
//! envelope and everything queued after it are discarded and the replica is
//! rebuilt from the committed log ("refresh"). Later queued edits are
//! discarded too because each was written on top of the refused one.

use std::collections::VecDeque;

use uuid::Uuid;
use visigrid_engine::workbook::Workbook;

use crate::apply::apply_ops;
use crate::op::{CollabOp, Envelope};
use crate::server::Committed;
use crate::transform::{transform_lists, Order};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ToServer {
    Submit(Envelope),
    /// Sent on (re)connect: "I have applied everything up to `last_seen`".
    Hello {
        last_seen: u64,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ToClient {
    Op(Committed),
    Refused {
        client_op_id: Uuid,
        reason: String,
    },
    /// End of the catch-up that answers a Hello.
    Welcome {
        head: u64,
    },
}

#[derive(Clone, Debug)]
pub struct Pending {
    pub client_op_id: Uuid,
    pub ops: Vec<CollabOp>,
}

#[derive(Clone, Debug, Default)]
pub struct ClientStats {
    pub local_envelopes: u64,
    pub acked: u64,
    pub refused_envelopes: u64,
    pub discarded_after_refusal: u64,
    pub resyncs: u64,
    pub max_pending: usize,
}

pub struct Client {
    pub actor: u64,
    pub wb: Workbook,
    pub last_seen: u64,
    /// Everything sequenced so far, in order (what a real client would get
    /// as a snapshot plus the log tail when it refreshes).
    pub committed: Vec<Committed>,
    pub inflight: Option<Pending>,
    pub buffer: VecDeque<Pending>,
    pub connected: bool,
    /// Between sending Hello and receiving Welcome: no submissions.
    pub syncing: bool,
    resend_inflight: bool,
    pub stats: ClientStats,
}

impl Client {
    pub fn new(actor: u64) -> Self {
        Client {
            actor,
            wb: Workbook::new(),
            last_seen: 0,
            committed: Vec::new(),
            inflight: None,
            buffer: VecDeque::new(),
            connected: true,
            syncing: false,
            resend_inflight: false,
            stats: ClientStats::default(),
        }
    }

    pub fn pending_count(&self) -> usize {
        self.buffer.len() + usize::from(self.inflight.is_some())
    }

    /// A local edit: applied now, queued for the server.
    pub fn local(&mut self, client_op_id: Uuid, ops: Vec<CollabOp>) {
        apply_ops(&mut self.wb, &ops);
        self.buffer.push_back(Pending { client_op_id, ops });
        self.stats.local_envelopes += 1;
        self.stats.max_pending = self.stats.max_pending.max(self.pending_count());
    }

    /// The next message to send, if any.
    pub fn poll_send(&mut self) -> Option<ToServer> {
        if !self.connected || self.syncing {
            return None;
        }
        if self.resend_inflight {
            self.resend_inflight = false;
            if let Some(p) = &self.inflight {
                return Some(ToServer::Submit(self.envelope(p)));
            }
        }
        if self.inflight.is_some() {
            return None;
        }
        let p = self.buffer.pop_front()?;
        let env = self.envelope(&p);
        self.inflight = Some(p);
        Some(ToServer::Submit(env))
    }

    fn envelope(&self, p: &Pending) -> Envelope {
        Envelope {
            client_op_id: p.client_op_id,
            base_seq: self.last_seen,
            actor: self.actor,
            ops: p.ops.clone(),
        }
    }

    pub fn disconnect(&mut self) {
        self.connected = false;
        self.syncing = false;
    }

    pub fn reconnect(&mut self) -> ToServer {
        self.connected = true;
        self.syncing = true;
        ToServer::Hello {
            last_seen: self.last_seen,
        }
    }

    pub fn receive(&mut self, msg: ToClient) {
        match msg {
            ToClient::Op(c) => self.on_op(c),
            ToClient::Refused { client_op_id, .. } => {
                // Normally already handled: the ops that caused the refusal
                // arrived first (FIFO) and our own transform refused the
                // envelope. If not, refresh now.
                if self.inflight.as_ref().map(|p| p.client_op_id) == Some(client_op_id) {
                    self.refuse_from(0);
                }
            }
            ToClient::Welcome { .. } => {
                self.syncing = false;
                if self.inflight.is_some() {
                    // The server may never have received it; resending is
                    // safe because the id makes it idempotent.
                    self.resend_inflight = true;
                }
            }
        }
    }

    fn on_op(&mut self, c: Committed) {
        if c.seq <= self.last_seen {
            return; // already have it (overlapping catch-up)
        }
        assert_eq!(c.seq, self.last_seen + 1, "ops arrive in sequence order");
        self.last_seen = c.seq;
        self.committed.push(c.clone());
        if self.inflight.as_ref().map(|p| p.client_op_id) == Some(c.client_op_id) {
            // Our own op: the acknowledgement. Its effect is already local.
            self.inflight = None;
            self.stats.acked += 1;
            return;
        }
        // Someone else's: transform our pending envelopes past it, and it
        // past them, in order.
        let mut remote = c.ops.clone();
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        let had_inflight = !entries.is_empty();
        entries.extend(self.buffer.drain(..));
        let mut refused_at = None;
        for (k, p) in entries.iter_mut().enumerate() {
            match transform_lists(&p.ops, &remote, Order::Later) {
                Ok((p2, r2)) => {
                    p.ops = p2;
                    remote = r2;
                }
                Err(_) => {
                    refused_at = Some(k);
                    break;
                }
            }
        }
        // Put the (transformed) entries back.
        let mut it = entries.into_iter();
        if had_inflight {
            self.inflight = it.next();
        }
        self.buffer.extend(it);
        match refused_at {
            None => {
                apply_ops(&mut self.wb, &remote);
            }
            Some(k) => self.refuse_from(k),
        }
    }

    /// Drop pending entry `k` (0 = in-flight when present) and everything
    /// after it, then rebuild from the committed log plus what remains.
    fn refuse_from(&mut self, k: usize) {
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        let had_inflight = !entries.is_empty();
        entries.extend(self.buffer.drain(..));
        let dropped = entries.len().saturating_sub(k);
        entries.truncate(k);
        self.stats.refused_envelopes += 1;
        self.stats.discarded_after_refusal += dropped.saturating_sub(1) as u64;
        let mut it = entries.into_iter();
        if had_inflight && k > 0 {
            self.inflight = it.next();
        }
        self.buffer.extend(it);
        self.resync();
    }

    /// Rebuild the replica: committed log, then pending edits on top.
    pub fn resync(&mut self) {
        self.stats.resyncs += 1;
        let mut wb = Workbook::new();
        for c in &self.committed {
            apply_ops(&mut wb, &c.ops);
        }
        if let Some(p) = &self.inflight {
            apply_ops(&mut wb, &p.ops);
        }
        for p in &self.buffer {
            apply_ops(&mut wb, &p.ops);
        }
        self.wb = wb;
    }
}
