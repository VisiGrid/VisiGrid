//! Intent tests for per-user undo (spec §Undo) through real clients and the
//! in-memory sequencer: undo is a new op, reverts only the user's own change,
//! follows concurrent structural edits, and leaves collaborators' later
//! changes alone.

use uuid::Uuid;
use visigrid_collab::apply::fingerprint;
use visigrid_collab::client::{Client, ToClient, ToServer};
use visigrid_collab::op::{Axis, CellContent, CollabOp, FormatProps, Rect};
use visigrid_collab::server::{Server, Submitted};

struct Room {
    server: Server,
    clients: Vec<Client>,
    next_id: u128,
}

impl Room {
    fn new(n: usize) -> Room {
        let clients = (0..n)
            .map(|i| {
                let mut c = Client::new(i as u64 + 1);
                c.enable_undo();
                c
            })
            .collect();
        Room { server: Server::new(), clients, next_id: 1 }
    }

    fn id(&mut self) -> Uuid {
        self.next_id += 1;
        Uuid::from_u128(self.next_id)
    }

    fn key(&self) -> u64 {
        self.server.wb.sheets()[0].id.0
    }

    fn set(&self, row: usize, col: usize, text: &str) -> CollabOp {
        CollabOp::SetCell {
            sheet: self.key(),
            sheet_name: "Sheet1".into(),
            row,
            col,
            content: if text.is_empty() { CellContent::Clear } else { CellContent::Value(text.into()) },
        }
    }

    fn rows(&self, at: usize, count: usize, delete: bool) -> CollabOp {
        CollabOp::Structural { sheet: self.key(), sheet_name: "Sheet1".into(), axis: Axis::Row, at, count, delete }
    }

    fn edit(&mut self, who: usize, ops: Vec<CollabOp>) {
        let id = self.id();
        self.clients[who].local(id, ops);
    }

    fn undo(&mut self, who: usize) -> visigrid_collab::undo::UndoOutcome {
        let id = self.id();
        self.clients[who].undo(id)
    }

    fn redo(&mut self, who: usize) -> visigrid_collab::undo::UndoOutcome {
        let id = self.id();
        self.clients[who].redo(id)
    }

    /// Deliver everything until quiet; every replica must equal the server.
    fn sync(&mut self) {
        loop {
            let mut moved = false;
            for i in 0..self.clients.len() {
                while let Some(ToServer::Submit(env)) = self.clients[i].poll_send() {
                    moved = true;
                    match self.server.submit(&env) {
                        Submitted::Committed(c) => {
                            for client in self.clients.iter_mut() {
                                client.receive(ToClient::Op(c.clone()));
                            }
                        }
                        Submitted::Refused { client_op_id, reason } => {
                            self.clients[i].receive(ToClient::Refused { client_op_id, reason });
                        }
                        Submitted::Duplicate => {}
                    }
                }
            }
            if !moved {
                break;
            }
        }
        for c in &self.clients {
            assert_eq!(fingerprint(&c.wb), fingerprint(&self.server.wb), "replica {} diverged", c.actor);
        }
    }

    fn raw(&self, row: usize, col: usize) -> String {
        self.server.wb.sheets()[0].get_raw(row, col)
    }
}

#[test]
fn undo_follows_a_concurrent_row_insert() {
    let mut room = Room::new(2);
    let a3 = room.set(2, 0, "mine");
    room.edit(0, vec![a3]);
    // Before A's edit reaches B, B inserts two rows above it.
    let insert = room.rows(0, 2, false);
    room.edit(1, vec![insert]);
    room.sync();
    assert_eq!(room.raw(4, 0), "mine");
    assert!(room.undo(0).applied);
    room.sync();
    assert_eq!(room.raw(4, 0), "", "undo cleared the shifted cell, not A3");
    assert_eq!(room.raw(2, 0), "");
}

#[test]
fn undo_reverts_only_the_users_own_change() {
    let mut room = Room::new(2);
    let (a, b) = (room.set(0, 0, "a"), room.set(0, 1, "b"));
    room.edit(0, vec![a]);
    room.edit(1, vec![b]);
    room.sync();
    assert!(room.undo(1).applied);
    room.sync();
    assert_eq!((room.raw(0, 0).as_str(), room.raw(0, 1).as_str()), ("a", ""));
}

#[test]
fn undo_leaves_a_collaborators_later_write() {
    let mut room = Room::new(2);
    let first = room.set(0, 0, "first");
    room.edit(0, vec![first]);
    room.sync();
    let theirs = room.set(0, 0, "theirs");
    room.edit(1, vec![theirs]);
    room.sync();
    let out = room.undo(0);
    assert_eq!((out.applied, out.kept_others, out.reason), (false, 1, Some("gone")));
    room.sync();
    assert_eq!(room.raw(0, 0), "theirs");
}

#[test]
fn undo_after_the_row_was_deleted_is_a_noop_with_a_notice() {
    let mut room = Room::new(2);
    let a3 = room.set(2, 0, "doomed");
    room.edit(0, vec![a3]);
    room.sync();
    let delete = room.rows(2, 1, true);
    room.edit(1, vec![delete]);
    room.sync();
    let out = room.undo(0);
    assert_eq!((out.applied, out.reason), (false, Some("gone")));
    room.sync();
}

#[test]
fn redo_reapplies_and_a_new_edit_clears_it() {
    let mut room = Room::new(2);
    let (x, y) = (room.set(0, 0, "x"), room.set(1, 0, "y"));
    room.edit(0, vec![x]);
    room.edit(0, vec![y]);
    room.sync();
    assert!(room.undo(0).applied);
    assert!(room.undo(0).applied);
    room.sync();
    assert_eq!((room.raw(0, 0).as_str(), room.raw(1, 0).as_str()), ("", ""));
    assert!(room.redo(0).applied);
    room.sync();
    assert_eq!(room.raw(0, 0), "x");
    let z = room.set(5, 5, "z");
    room.edit(0, vec![z]);
    assert_eq!(room.redo(0).reason, Some("empty"));
    room.sync();
}

#[test]
fn undoing_a_row_delete_restores_its_cells_and_formats() {
    let mut room = Room::new(2);
    let key = room.key();
    let setup = vec![
        room.set(1, 0, "keep"),
        room.set(1, 1, "=A2&\"!\""),
        CollabOp::SetFormat { sheet: key, rect: Rect::new(1, 0, 1, 0), props: FormatProps::bold(true) },
    ];
    room.edit(0, setup);
    room.sync();
    let before = fingerprint(&room.server.wb);
    let delete = room.rows(1, 1, true);
    room.edit(0, vec![delete]);
    // Concurrently, B writes elsewhere.
    let elsewhere = room.set(9, 9, "b");
    room.edit(1, vec![elsewhere]);
    room.sync();
    assert!(room.undo(0).applied);
    room.sync();
    assert_eq!(room.raw(1, 0), "keep");
    assert!(room.server.wb.sheets()[0].get_format(1, 0).bold);
    // Everything but B's write is as before.
    let mut after = fingerprint(&room.server.wb);
    after.sheets[0].cells.retain(|(r, c, ..)| (*r, *c) != (9, 9));
    assert_eq!(after, before);
}

#[test]
fn format_undo_keeps_a_collaborators_property() {
    let mut room = Room::new(2);
    let key = room.key();
    let mine = CollabOp::SetFormat {
        sheet: key,
        rect: Rect::new(0, 0, 0, 0),
        props: FormatProps { bold: Some(Some(true)), color: Some(Some("#FF0000".into())), ..Default::default() },
    };
    room.edit(0, vec![mine]);
    room.sync();
    let theirs = CollabOp::SetFormat {
        sheet: key,
        rect: Rect::new(0, 0, 0, 0),
        props: FormatProps { color: Some(Some("#0000FF".into())), ..Default::default() },
    };
    room.edit(1, vec![theirs]);
    room.sync();
    let out = room.undo(0);
    assert!(out.applied);
    room.sync();
    let f = room.server.wb.sheets()[0].get_format(0, 0);
    assert!(!f.bold, "my bold is undone");
    assert_eq!(f.font_color, Some([0, 0, 255, 255]), "their colour stays");
}
