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
//! The replica keeps two workbooks: `confirmed` (exactly the server's state
//! as of `last_seen`) and `wb`, the optimistic state the user sees
//! (confirmed plus every pending envelope, in order).
//!
//! If a pending envelope is refused (a V1-serialized conflict), only that
//! envelope is discarded. Envelopes queued after it were written on top of
//! it, so each is transformed past the refused envelope's *positional
//! inverse* (an insert it made is undone as a delete, and so on) before it
//! is kept; one that no longer applies is discarded with it. The optimistic
//! state is then rebuilt as confirmed + the remaining pending envelopes.

use std::collections::VecDeque;

use uuid::Uuid;
use visigrid_engine::workbook::Workbook;

use crate::apply::{apply_ops, apply_ops_tracked, filter_unappliable, Changes};
use crate::op::{CollabOp, Envelope, SheetKey};
use crate::server::Committed;
use crate::transform::{transform, transform_lists, Order, Transformed};

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
    /// Envelopes refused themselves (the conflicting edit).
    pub refused_envelopes: u64,
    /// Other pending envelopes lost because of a refusal: ones built on
    /// the refused edit that no longer apply without it.
    pub discarded_after_refusal: u64,
    /// Of `discarded_after_refusal`: lost because the removed edit's effect
    /// could not be excluded (a sheet rename or delete).
    pub discarded_no_inverse: u64,
    /// Pending envelopes dropped by a document replacement.
    pub discarded_by_replacement: u64,
    /// Pending envelopes kept across a document replacement and resent.
    pub kept_across_replacement: u64,
    pub resyncs: u64,
    pub max_pending: usize,
}

pub struct Client {
    pub actor: u64,
    /// Optimistic state: `confirmed` plus every pending envelope.
    pub wb: Workbook,
    /// The server's state as of `last_seen`.
    pub confirmed: Workbook,
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
    /// Measurement only: the pre-10/4 refusal policy, which discarded every
    /// envelope queued after a refused one. Off by default.
    pub legacy_refusal: bool,
    /// When set, every change to the optimistic state `wb` is recorded here,
    /// for a UI mirroring it (the browser client). The caller takes it with
    /// `take_changes`. Recording never changes what is applied.
    pub changes: Option<Changes>,
}

impl Client {
    pub fn new(actor: u64) -> Self {
        Client {
            actor,
            wb: Workbook::new(),
            confirmed: Workbook::new(),
            last_seen: 0,
            committed: Vec::new(),
            inflight: None,
            buffer: VecDeque::new(),
            connected: true,
            syncing: false,
            resend_inflight: false,
            stats: ClientStats::default(),
            legacy_refusal: false,
            changes: None,
        }
    }

    /// Start recording changes to the optimistic state.
    pub fn record_changes(&mut self) {
        self.changes.get_or_insert_with(Changes::default);
    }

    /// The changes recorded since the last call (empty when not recording).
    pub fn take_changes(&mut self) -> Changes {
        match self.changes.as_mut() {
            Some(ch) => std::mem::take(ch),
            None => Changes::default(),
        }
    }

    /// Apply ops to the optimistic state, recording them when asked.
    fn apply_optimistic(&mut self, ops: &[CollabOp]) {
        match self.changes.as_mut() {
            Some(ch) => {
                apply_ops_tracked(&mut self.wb, ops, ch);
            }
            None => {
                apply_ops(&mut self.wb, ops);
            }
        }
    }

    pub fn pending_count(&self) -> usize {
        self.buffer.len() + usize::from(self.inflight.is_some())
    }

    /// A local edit: applied now, queued for the server.
    pub fn local(&mut self, client_op_id: Uuid, ops: Vec<CollabOp>) {
        self.apply_optimistic(&ops);
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
                    self.refuse_inflight();
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

    /// Protocol v2 `ack{client_op_id, seq}` for our in-flight envelope. Its
    /// effect is already local: by FIFO we have applied every op sequenced
    /// before it and transformed the envelope past each, exactly as the
    /// server did, so the committed form equals our pending form.
    pub fn ack(&mut self, client_op_id: Uuid, seq: u64) {
        let ops = match &self.inflight {
            Some(p) if p.client_op_id == client_op_id => p.ops.clone(),
            _ => return, // a duplicate ack, or one for an envelope already resolved
        };
        self.on_op(Committed {
            seq,
            client_op_id,
            actor: self.actor,
            ops,
        });
    }

    /// Protocol v2 `rejected{result: "dropped"}`: the server applied nothing
    /// and assigned no sequence number. Our own transform saw the same ops
    /// and normally reduced the envelope to nothing as well; if it did not,
    /// refresh rather than keep an effect the server does not have.
    pub fn dropped(&mut self, client_op_id: Uuid) {
        match &self.inflight {
            Some(p) if p.client_op_id == client_op_id => {
                if p.ops.is_empty() {
                    self.inflight = None;
                    self.stats.acked += 1;
                } else {
                    self.refuse_inflight();
                }
            }
            _ => {}
        }
    }

    fn on_op(&mut self, c: Committed) {
        if c.seq <= self.last_seen {
            return; // already have it (overlapping catch-up)
        }
        assert_eq!(c.seq, self.last_seen + 1, "ops arrive in sequence order");
        self.last_seen = c.seq;
        apply_ops(&mut self.confirmed, &c.ops);
        self.committed.push(c.clone());
        if self.inflight.as_ref().map(|p| p.client_op_id) == Some(c.client_op_id) {
            // Our own op: the acknowledgement. Its effect is already local,
            // and the committed form equals our pending form (the server
            // transformed it past the same ops we did).
            self.inflight = None;
            self.stats.acked += 1;
            return;
        }
        // Someone else's: transform our pending envelopes past it, and it
        // past them, in order. A refused envelope is removed (with its
        // positional effect excluded from the envelopes after it) and the
        // remaining ones continue past the remote op.
        let mut remote = c.ops.clone();
        let had_inflight = self.inflight.is_some();
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        entries.extend(self.buffer.drain(..));
        let mut inflight_alive = had_inflight;
        let mut rebuild = false;
        let mut k = 0;
        while k < entries.len() {
            match transform_lists(&entries[k].ops, &remote, Order::Later) {
                Ok((p2, r2)) => {
                    entries[k].ops = p2;
                    remote = r2;
                    k += 1;
                }
                Err(_) => {
                    if k == 0 && had_inflight {
                        // The server sees the same pair in the same order and
                        // refuses it too; its `rejected` will find nothing.
                        inflight_alive = false;
                    }
                    self.remove_entry(&mut entries, k);
                    rebuild = true;
                }
            }
        }
        let mut it = entries.into_iter();
        if inflight_alive {
            self.inflight = it.next();
        }
        self.buffer.extend(it);
        if rebuild {
            self.rebuild();
        } else {
            self.apply_optimistic(&remote);
        }
    }

    /// Remove `entries[k]`, rebasing every later entry past its positional
    /// inverse. Pending envelopes reach no other replica until sent, and the
    /// optimistic state is rebuilt as confirmed + exactly the envelopes we
    /// will send, so this rebase cannot affect convergence: it only decides
    /// how much of the user's later intent survives. It is therefore
    /// permissive: V1 conflicts keep the later op unchanged, and only an op
    /// whose target no longer exists is dropped (an entry left empty goes).
    fn remove_entry(&mut self, entries: &mut Vec<Pending>, k: usize) {
        // The state the removed envelope was written against: confirmed plus
        // the envelopes before it (all in the current frame).
        let mut before = self.confirmed.clone();
        for e in &entries[..k] {
            apply_ops(&mut before, &e.ops);
        }
        if self.legacy_refusal {
            entries.remove(k);
            self.stats.refused_envelopes += 1;
            let lost = entries.len() - k;
            entries.truncate(k);
            self.stats.discarded_after_refusal += lost as u64;
            return;
        }
        let removed = entries.remove(k);
        self.stats.refused_envelopes += 1;
        let mut inverse = positional_inverse(&removed.ops, &before);
        let mut j = k;
        while j < entries.len() && !inverse.is_empty() {
            let (e2, inv2) = rebase(&entries[j].ops, &inverse);
            inverse = inv2;
            if e2.is_empty() && !entries[j].ops.is_empty() {
                entries.remove(j);
                self.stats.discarded_after_refusal += 1;
            } else {
                entries[j].ops = e2;
                j += 1;
            }
        }
    }

    /// The server refused our in-flight envelope without our own transform
    /// having seen the conflict.
    fn refuse_inflight(&mut self) {
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        if entries.is_empty() {
            return;
        }
        entries.extend(self.buffer.drain(..));
        self.remove_entry(&mut entries, 0);
        self.buffer.extend(entries);
        self.rebuild();
    }

    /// Optimistic state = confirmed + pending envelopes, in order.
    ///
    /// Every buffered op's carried sheet name is reset to the name its sheet
    /// has in the state it applies to. Ops carry names because formula
    /// rewriting is name based, and after a refused rename is removed the
    /// names in later envelopes would otherwise describe a state that never
    /// existed (the simulator found the resulting reference rewrites
    /// diverging). The in-flight envelope is left as sent.
    fn rebuild(&mut self) {
        self.stats.resyncs += 1;
        let mut wb = self.confirmed.clone();
        if let Some(p) = &self.inflight {
            apply_ops(&mut wb, &p.ops);
        }
        // An op that cannot apply where it lands (missing sheet, taken name,
        // last sheet) is removed: the rule the sequencer applies
        // (`filter_unappliable`). Kept here, it would still shift positions
        // and rewrite formulas in remote ops transformed past it.
        let mut emptied = 0u64;
        for p in self.buffer.iter_mut() {
            let had_ops = !p.ops.is_empty();
            p.ops = filter_unappliable(&wb, &p.ops).0;
            if had_ops && p.ops.is_empty() {
                emptied += 1;
            }
            for op in p.ops.iter_mut() {
                let current = wb.sheets().iter().find(|s| s.id.0 == op.sheet()).map(|s| s.name.clone());
                if let Some(name) = current {
                    *op = with_carried_name(op, op.sheet(), &name);
                }
                // Tab positions describe the state too.
                match op {
                    CollabOp::DeleteSheet { sheet, index } => {
                        if let Some(at) = wb.sheets().iter().position(|s| s.id.0 == *sheet) {
                            *index = at;
                        }
                    }
                    CollabOp::AddSheet { index, .. } => *index = (*index).min(wb.sheets().len()),
                    _ => {}
                }
                apply_ops(&mut wb, std::slice::from_ref(op));
            }
        }
        if emptied > 0 {
            self.buffer.retain(|p| !p.ops.is_empty());
            self.stats.discarded_after_refusal += emptied;
        }
        self.wb = wb;
        if let Some(ch) = self.changes.as_mut() {
            // The optimistic state was replaced wholesale.
            ch.full = true;
            ch.sheets = true;
        }
    }

    /// The server replaced the whole document (an old client's whole-file
    /// save) at `seq`. Pending envelopes made only of cell content and
    /// formatting (SetCell, SetBold, SetFormat, ReplaceRange) on sheets that still exist
    /// are kept and resent on top of the new document: they name cells by
    /// stable sheet id and position, which a replacement keeps meaning.
    /// Envelopes with structural or sheet edits are dropped: their row,
    /// column and tab positions were relative to a document that no longer
    /// exists. The in-flight envelope, if any, is resent (the server refuses
    /// it as `document_replaced` if it arrived before the barrier, and the id
    /// keeps a duplicate harmless).
    pub fn replace_document(&mut self, document: Workbook, seq: u64) {
        let present: std::collections::HashSet<SheetKey> = document.sheets().iter().map(|s| s.id.0).collect();
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        entries.extend(self.buffer.drain(..));
        for p in entries {
            let keep = p.ops.iter().all(|op| match op {
                CollabOp::SetCell { sheet, .. }
                | CollabOp::SetBold { sheet, .. }
                | CollabOp::SetFormat { sheet, .. }
                | CollabOp::ReplaceRange { sheet, .. } => present.contains(sheet),
                _ => false,
            });
            if keep {
                self.stats.kept_across_replacement += 1;
                // A fresh id: the old one may already be refused at the server.
                self.buffer.push_back(Pending {
                    client_op_id: Uuid::from_u128(rand_like(p.client_op_id, seq)),
                    ops: p.ops,
                });
            } else {
                self.stats.discarded_by_replacement += 1;
            }
        }
        self.confirmed = document;
        self.last_seen = seq;
        self.committed.clear();
        self.rebuild();
    }

    /// Rebuild the optimistic state from the confirmed copy.
    pub fn resync(&mut self) {
        self.rebuild();
    }
}

/// Ops that undo `ops`' effect on positions and sheet names, for rebasing
/// envelopes that were written after them. `before` is the state `ops` were
/// written against. Content and formatting have no such effect. A rename is
/// undone with the name it replaced; a sheet delete with an add at its old
/// tab index (the name only matters for conflict checks, which the rebase
/// does not apply).
pub fn positional_inverse(ops: &[CollabOp], before: &Workbook) -> Vec<CollabOp> {
    // Step through the ops on a scratch copy so a rename or delete inside
    // the same envelope sees the name it actually replaced.
    let mut scratch = before.clone();
    let mut inverse = Vec::new();
    for op in ops {
        let name_of = |wb: &Workbook, sheet: SheetKey| {
            wb.sheets().iter().find(|s| s.id.0 == sheet).map(|s| s.name.clone())
        };
        match op {
            CollabOp::SetCell { .. }
            | CollabOp::SetBold { .. }
            | CollabOp::SetFormat { .. }
            | CollabOp::ReplaceRange { .. } => {}
            CollabOp::Structural { sheet, sheet_name, axis, at, count, delete } => {
                inverse.push(CollabOp::Structural {
                    sheet: *sheet,
                    sheet_name: sheet_name.clone(),
                    axis: *axis,
                    at: *at,
                    count: *count,
                    delete: !*delete,
                });
            }
            CollabOp::AddSheet { sheet, index, .. } => {
                inverse.push(CollabOp::DeleteSheet { sheet: *sheet, index: *index });
            }
            CollabOp::RenameSheet { sheet, .. } => {
                if let Some(name) = name_of(&scratch, *sheet) {
                    inverse.push(CollabOp::RenameSheet { sheet: *sheet, name });
                }
            }
            CollabOp::DeleteSheet { sheet, index } => {
                inverse.push(CollabOp::AddSheet {
                    sheet: *sheet,
                    name: name_of(&scratch, *sheet).unwrap_or_default(),
                    index: *index,
                });
            }
        }
        apply_ops(&mut scratch, std::slice::from_ref(op));
    }
    // Undo in reverse order.
    inverse.reverse();
    inverse
}

/// Rebase `entry` past `inverse` and `inverse` past `entry` (the usual
/// list transform), keeping a conflicting op unchanged instead of refusing
/// it and dropping ops whose target is gone.
fn rebase(entry: &[CollabOp], inverse: &[CollabOp]) -> (Vec<CollabOp>, Vec<CollabOp>) {
    let mut inv: Vec<CollabOp> = inverse.to_vec();
    let mut out = Vec::with_capacity(entry.len());
    for e in entry {
        let mut cur = vec![e.clone()];
        let mut next_inv = Vec::with_capacity(inv.len());
        for iv in &inv {
            next_inv.extend(permissive(iv, &cur));
            cur = cur.iter().flat_map(|x| permissive(x, std::slice::from_ref(iv))).collect();
        }
        out.extend(cur);
        inv = next_inv;
    }
    (out, inv)
}

/// `a` past every op of `bs` in turn; V1 conflicts keep `a`. A rename is
/// applied as what it is to a later op, a new carried sheet name, without
/// the V1 rename conflict: a stale carried name would make formula
/// rewriting in later transforms resolve the wrong sheet.
fn permissive(a: &CollabOp, bs: &[CollabOp]) -> Vec<CollabOp> {
    let mut cur = vec![a.clone()];
    for b in bs {
        let mut next = Vec::new();
        for x in &cur {
            if let CollabOp::RenameSheet { sheet, name } = b {
                next.push(with_carried_name(x, *sheet, name));
                continue;
            }
            match transform(x, b, Order::Later) {
                Transformed::Ops(v) => next.extend(v),
                Transformed::Dropped(_) => {}
                Transformed::Refused(_) => next.push(x.clone()),
            }
        }
        cur = next;
    }
    cur
}

/// A deterministic new id derived from an old one and the barrier seq.
fn rand_like(id: Uuid, seq: u64) -> u128 {
    id.as_u128() ^ (u128::from(seq).wrapping_mul(0x9E37_79B9_7F4A_7C15_F39C_C060_5CED_C835))
}

/// `op` with the sheet name it carries for `sheet` set to `name`.
fn with_carried_name(op: &CollabOp, sheet: SheetKey, name: &str) -> CollabOp {
    let mut op = op.clone();
    match &mut op {
        CollabOp::SetCell { sheet: s, sheet_name, .. } | CollabOp::Structural { sheet: s, sheet_name, .. }
            if *s == sheet =>
        {
            *sheet_name = name.to_string();
        }
        _ => {}
    }
    op
}
