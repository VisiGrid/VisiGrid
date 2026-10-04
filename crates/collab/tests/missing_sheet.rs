//! An op naming a sheet the server does not have is dropped by the
//! sequencer and never applied anywhere; every replica still converges.

use uuid::Uuid;
use visigrid_collab::apply::{filter_missing_sheets, fingerprint};
use visigrid_collab::client::{Client, ToClient, ToServer};
use visigrid_collab::op::{CellContent, CollabOp};
use visigrid_collab::server::{Server, Submitted};
use visigrid_engine::workbook::Workbook;

fn cell(sheet: u64, row: usize, text: &str) -> CollabOp {
    CollabOp::SetCell { sheet, sheet_name: "Sheet1".into(), row, col: 0, content: CellContent::Value(text.into()) }
}

/// Deliver every message until quiet: submissions to the server, results to everyone.
fn exchange(server: &mut Server, clients: &mut [Client]) {
    loop {
        let mut moved = false;
        for i in 0..clients.len() {
            while let Some(ToServer::Submit(env)) = clients[i].poll_send() {
                moved = true;
                match server.submit(&env) {
                    Submitted::Committed(c) => {
                        for c2 in clients.iter_mut() {
                            c2.receive(ToClient::Op(c.clone()));
                        }
                    }
                    Submitted::Refused { client_op_id, reason } => {
                        clients[i].receive(ToClient::Refused { client_op_id, reason })
                    }
                    Submitted::Duplicate => {}
                }
            }
        }
        if !moved {
            return;
        }
    }
}

#[test]
fn ghost_ops_are_filtered_and_replicas_converge() {
    let mut server = Server::new();
    let mut clients = vec![Client::new(1), Client::new(2)];
    clients[0].local(Uuid::from_u128(1), vec![cell(0, 0, "ghost"), cell(1, 1, "real")]);
    clients[1].local(Uuid::from_u128(2), vec![cell(99, 0, "ghost only")]);
    clients[1].local(Uuid::from_u128(3), vec![cell(1, 2, "after")]);
    exchange(&mut server, &mut clients);

    for c in &server.log {
        assert!(c.ops.iter().all(|op| op.sheet() == 1), "nothing on a missing sheet is sequenced: {:?}", c.ops);
    }
    let reference = fingerprint(&server.wb);
    for c in &clients {
        assert_eq!(c.pending_count(), 0);
        assert_eq!(fingerprint(&c.wb), reference);
    }
    assert_eq!(server.wb.sheets()[0].get_raw(1, 0), "real");
    assert_eq!(server.wb.sheets()[0].get_raw(2, 0), "after");
}

#[test]
fn filter_follows_sheets_added_and_deleted_in_the_same_envelope() {
    let wb = Workbook::new();
    let ops = vec![
        cell(5, 0, "before add"),
        CollabOp::AddSheet { sheet: 5, name: "Five".into(), index: 1 },
        cell(5, 0, "after add"),
        CollabOp::DeleteSheet { sheet: 5, index: 1 },
        cell(5, 1, "after delete"),
    ];
    let (kept, dropped) = filter_missing_sheets(&wb, &ops);
    assert_eq!(dropped, 2);
    assert_eq!(kept, ops[1..4].to_vec());
}
