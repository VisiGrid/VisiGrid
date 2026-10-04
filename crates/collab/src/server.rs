//! In-memory sequencer: one per open workbook.
//!
//! An incoming envelope names the last sequence number its sender had
//! applied (`base_seq`). The server transforms it past every op sequenced
//! since, appends it with the next sequence number, applies it to its own
//! replica, and broadcasts the transformed ops to everyone (the sender
//! treats its own op coming back as the acknowledgement).

use std::collections::BTreeMap;

use uuid::Uuid;
use visigrid_engine::workbook::Workbook;

use crate::apply::apply_ops;
use crate::op::{CollabOp, Envelope};
use crate::transform::{transform_lists, Order};

/// A sequenced envelope.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Committed {
    pub seq: u64,
    pub client_op_id: Uuid,
    pub actor: u64,
    /// Transformed ops; empty when everything in the envelope was dropped
    /// (it still takes a sequence number, so acknowledgements stay simple).
    pub ops: Vec<CollabOp>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Submitted {
    Committed(Committed),
    Refused {
        client_op_id: Uuid,
        reason: String,
    },
    /// Already seen this id (a resend). The caller does nothing: the client
    /// learns the outcome from the op stream it catches up on.
    Duplicate,
}

pub struct Server {
    pub wb: Workbook,
    pub log: Vec<Committed>,
    seen: BTreeMap<Uuid, Option<u64>>,
    pub refused: u64,
}

impl Default for Server {
    fn default() -> Self {
        Self::new()
    }
}

impl Server {
    pub fn new() -> Self {
        Server {
            wb: Workbook::new(),
            log: Vec::new(),
            seen: BTreeMap::new(),
            refused: 0,
        }
    }

    pub fn head(&self) -> u64 {
        self.log.len() as u64
    }

    /// Committed envelopes after `seq`.
    pub fn since(&self, seq: u64) -> &[Committed] {
        &self.log[(seq as usize).min(self.log.len())..]
    }

    pub fn submit(&mut self, env: &Envelope) -> Submitted {
        if self.seen.contains_key(&env.client_op_id) {
            return Submitted::Duplicate;
        }
        assert!(
            env.base_seq <= self.head(),
            "a client cannot have seen the future"
        );
        let mut ops = env.ops.clone();
        for c in &self.log[env.base_seq as usize..] {
            match transform_lists(&ops, &c.ops, Order::Later) {
                Ok((a, _)) => ops = a,
                Err(r) => {
                    self.seen.insert(env.client_op_id, None);
                    self.refused += 1;
                    return Submitted::Refused {
                        client_op_id: env.client_op_id,
                        reason: r.reason,
                    };
                }
            }
        }
        let committed = Committed {
            seq: self.head() + 1,
            client_op_id: env.client_op_id,
            actor: env.actor,
            ops,
        };
        apply_ops(&mut self.wb, &committed.ops);
        self.seen.insert(env.client_op_id, Some(committed.seq));
        self.log.push(committed.clone());
        Submitted::Committed(committed)
    }
}
