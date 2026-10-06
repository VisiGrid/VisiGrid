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
use crate::undo::{apply_recording, resolve, UndoEntry, UndoOutcome};

/// Undo and redo entries kept per user.
pub const MAX_UNDO: usize = 100;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    Edit,
    Undo,
    Redo,
}

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
    /// Per-user undo (`enable_undo`): entries in the current frame of `wb`.
    pub undo_stack: Vec<UndoEntry>,
    pub redo_stack: Vec<UndoEntry>,
    track_undo: bool,
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
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            track_undo: false,
        }
    }

    /// Record undo entries for local edits from now on.
    pub fn enable_undo(&mut self) {
        self.track_undo = true;
    }

    /// Start recording changes to the optimistic state.
    pub fn record_changes(&mut self) {
        self.changes.get_or_insert_with(Changes::default);
    }

    /// Calculation outside an operation's tracked writes can change cells
    /// anywhere in the workbook, such as NOW/RAND after a sequenced clock.
    pub fn record_calculation_refresh(&mut self) {
        if let Some(changes) = self.changes.as_mut() {
            changes.full = true;
        }
    }

    /// The changes recorded since the last call (empty when not recording).
    pub fn take_changes(&mut self) -> Changes {
        match self.changes.as_mut() {
            Some(ch) => std::mem::take(ch),
            None => Changes::default(),
        }
    }

    /// Apply ops to the optimistic state, recording them when asked, and
    /// carry the undo history past them.
    fn apply_optimistic(&mut self, ops: &[CollabOp]) {
        apply_to(&mut self.wb, &mut self.changes, ops);
        self.shift_history(ops);
    }

    /// Transform every undo and redo entry past ops just applied to `wb`.
    fn shift_history(&mut self, ops: &[CollabOp]) {
        // Only structure and sheet ops move or drop a later op; cell, format,
        // line, freeze and merge ops leave it as it is (see `transform` with
        // `Order::Later`). Skipping them keeps a large paste from costing
        // history x paste transforms.
        let movers: Vec<CollabOp> = ops.iter().filter(|op| moves_others(op)).cloned().collect();
        if movers.is_empty() {
            return;
        }
        for e in self.undo_stack.iter_mut().chain(self.redo_stack.iter_mut()) {
            e.inverse = e.inverse.iter().flat_map(|x| permissive(x, &movers)).collect();
            e.expect = e.expect.iter().flat_map(|x| permissive(x, &movers)).collect();
        }
    }

    pub fn pending_count(&self) -> usize {
        self.buffer.len() + usize::from(self.inflight.is_some())
    }

    /// A local edit: applied now, queued for the server.
    pub fn local(&mut self, client_op_id: Uuid, ops: Vec<CollabOp>) {
        self.local_kind(client_op_id, ops, Kind::Edit);
    }

    /// Undo this user's most recent change still on the stack, as a new
    /// local envelope with id `client_op_id`. Cells someone else changed
    /// since are left alone.
    pub fn undo(&mut self, client_op_id: Uuid) -> UndoOutcome {
        self.undo_redo(client_op_id, Kind::Undo)
    }

    /// Redo the most recently undone change (until a new edit clears it).
    pub fn redo(&mut self, client_op_id: Uuid) -> UndoOutcome {
        self.undo_redo(client_op_id, Kind::Redo)
    }

    fn undo_redo(&mut self, client_op_id: Uuid, kind: Kind) -> UndoOutcome {
        let stack = if kind == Kind::Undo { &mut self.undo_stack } else { &mut self.redo_stack };
        let Some(entry) = stack.pop() else {
            return UndoOutcome { reason: Some("empty"), ..Default::default() };
        };
        let (ops, kept_others) = resolve(&entry, &self.wb);
        let ops = with_current_tab_positions(&self.wb, filter_unappliable(&self.wb, &ops).0);
        if ops.is_empty() {
            return UndoOutcome { applied: false, kept_others, reason: Some("gone") };
        }
        self.local_kind(client_op_id, ops, kind);
        UndoOutcome { applied: true, kept_others, reason: None }
    }

    fn local_kind(&mut self, client_op_id: Uuid, ops: Vec<CollabOp>, kind: Kind) {
        if self.track_undo {
            let changes = &mut self.changes;
            let entry = apply_recording(&mut self.wb, client_op_id, &ops, |wb, op| apply_to(wb, changes, op));
            self.shift_history(&ops);
            // An edit that changed nothing (an editor writing the same value
            // twice) is no undo step, and keeps the redo stack.
            if !(kind == Kind::Edit && entry.is_noop()) {
                let stack = match kind {
                    Kind::Edit => {
                        self.redo_stack.clear();
                        &mut self.undo_stack
                    }
                    Kind::Undo => &mut self.redo_stack,
                    Kind::Redo => &mut self.undo_stack,
                };
                stack.push(entry);
                if stack.len() > MAX_UNDO {
                    stack.remove(0);
                }
            }
        } else {
            self.apply_optimistic(&ops);
        }
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
        // How the optimistic frame moves when envelopes are removed: each
        // removal's inverse, then the remote op (for the undo history).
        let mut delta: Vec<CollabOp> = Vec::new();
        let mut removed: Vec<Uuid> = Vec::new();
        let mut k = 0;
        while k < entries.len() {
            match transform_lists(&entries[k].ops, &remote, Order::Later) {
                Ok((p2, r2)) => {
                    if p2.is_empty() && !entries[k].ops.is_empty() {
                        // The remote op made ours moot (the same row or sheet
                        // deleted twice): there is nothing of ours to undo.
                        removed.push(entries[k].client_op_id);
                    }
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
                    removed.push(entries[k].client_op_id);
                    delta.extend(self.remove_entry(&mut entries, k));
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
            delta.extend(remote);
            self.forget(&removed);
            self.shift_history(&delta);
        } else {
            self.apply_optimistic(&remote);
            self.forget(&removed);
        }
    }

    /// Drop the undo and redo entries of envelopes that never took effect.
    fn forget(&mut self, removed: &[Uuid]) {
        self.undo_stack.retain(|e| !removed.contains(&e.origin));
        self.redo_stack.retain(|e| !removed.contains(&e.origin));
    }

    /// Remove `entries[k]`, rebasing every later entry past its positional
    /// inverse. Pending envelopes reach no other replica until sent, and the
    /// optimistic state is rebuilt as confirmed + exactly the envelopes we
    /// will send, so this rebase cannot affect convergence: it only decides
    /// how much of the user's later intent survives. It is therefore
    /// permissive: V1 conflicts keep the later op unchanged, and only an op
    /// whose target no longer exists is dropped (an entry left empty goes).
    ///
    /// Returns the removed envelope's inverse as rebased to the end of the
    /// envelopes after it (empty if it had no positional effect).
    fn remove_entry(&mut self, entries: &mut Vec<Pending>, k: usize) -> Vec<CollabOp> {
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
            return Vec::new();
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
        if j < entries.len() {
            // Stopped early: nothing left to rebase past the rest.
            return Vec::new();
        }
        inverse
    }

    /// The server refused our in-flight envelope without our own transform
    /// having seen the conflict.
    fn refuse_inflight(&mut self) {
        let mut entries: Vec<Pending> = self.inflight.take().into_iter().collect();
        if entries.is_empty() {
            return;
        }
        entries.extend(self.buffer.drain(..));
        let removed = entries[0].client_op_id;
        let delta = self.remove_entry(&mut entries, 0);
        self.buffer.extend(entries);
        self.rebuild();
        self.forget(&[removed]);
        self.shift_history(&delta);
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
        let mut moot = Vec::new();
        for p in self.buffer.iter_mut() {
            let had_ops = !p.ops.is_empty();
            p.ops = filter_unappliable(&wb, &p.ops).0;
            if had_ops && p.ops.is_empty() {
                emptied += 1;
                moot.push(p.client_op_id);
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
                    CollabOp::MoveSheet { index, .. } => *index = (*index).min(wb.sheets().len().saturating_sub(1)),
                    _ => {}
                }
                apply_ops(&mut wb, std::slice::from_ref(op));
            }
        }
        if emptied > 0 {
            self.forget(&moot);
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
        // Positions in the old document mean nothing in the new one.
        self.undo_stack.clear();
        self.redo_stack.clear();
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
            | CollabOp::ReplaceRange { .. }
            | CollabOp::SetLines { .. }
            | CollabOp::SetFreeze { .. }
            | CollabOp::Merge { .. }
            | CollabOp::Unmerge { .. } => {}
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
            CollabOp::MoveSheet { sheet, .. } => {
                if let Some(at) = scratch.sheets().iter().position(|s| s.id.0 == *sheet) {
                    inverse.push(CollabOp::MoveSheet { sheet: *sheet, index: at });
                }
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

/// Whether `op` can change another op transformed past it as the later one.
fn moves_others(op: &CollabOp) -> bool {
    matches!(
        op,
        CollabOp::Structural { .. } | CollabOp::AddSheet { .. } | CollabOp::DeleteSheet { .. } | CollabOp::RenameSheet { .. } | CollabOp::MoveSheet { .. }
    )
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

/// Sheet ops carry tab positions that transforms compute with, so they must
/// describe the state the ops apply to. An undo entry's recorded positions
/// can be stale (an add asked for a position past the end), so set each
/// from the tab order as the list applies, as `rebuild` does.
fn with_current_tab_positions(wb: &Workbook, mut ops: Vec<CollabOp>) -> Vec<CollabOp> {
    let mut tabs: Vec<SheetKey> = wb.sheets().iter().map(|s| s.id.0).collect();
    for op in ops.iter_mut() {
        match op {
            CollabOp::DeleteSheet { sheet, index } => {
                if let Some(at) = tabs.iter().position(|s| s == sheet) {
                    *index = at;
                    tabs.remove(at);
                }
            }
            CollabOp::AddSheet { sheet, index, .. } => {
                *index = (*index).min(tabs.len());
                tabs.insert(*index, *sheet);
            }
            CollabOp::MoveSheet { sheet, index } => {
                if let Some(at) = tabs.iter().position(|s| s == sheet) {
                    tabs.remove(at);
                    *index = (*index).min(tabs.len());
                    tabs.insert(*index, *sheet);
                }
            }
            _ => {}
        }
    }
    ops
}

/// Apply ops to `wb`, recording into `changes` when it is set.
fn apply_to(wb: &mut Workbook, changes: &mut Option<Changes>, ops: &[CollabOp]) {
    match changes.as_mut() {
        Some(ch) => {
            apply_ops_tracked(wb, ops, ch);
        }
        None => {
            apply_ops(wb, ops);
        }
    }
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
