//! `Client::replace_document`: after a whole-document replacement, pending
//! envelopes of cell content and formatting are kept and resent; ones with
//! structural or sheet edits are dropped.

use uuid::Uuid;
use visigrid_collab::client::{Client, ToServer};
use visigrid_collab::op::{Axis, CellContent, CollabOp};
use visigrid_engine::workbook::Workbook;

fn cell(row: usize, text: &str) -> CollabOp {
    CollabOp::SetCell { sheet: 1, sheet_name: "Sheet1".into(), row, col: 0, content: CellContent::Value(text.into()) }
}

#[test]
fn keeps_cell_edits_and_drops_structural_ones() {
    let mut c = Client::new(1);
    c.local(Uuid::from_u128(1), vec![cell(0, "a")]);
    c.local(Uuid::from_u128(2), vec![CollabOp::Structural {
        sheet: 1,
        sheet_name: "Sheet1".into(),
        axis: Axis::Row,
        at: 0,
        count: 1,
        delete: false,
    }]);
    c.local(Uuid::from_u128(3), vec![cell(1, "b")]);
    c.local(Uuid::from_u128(4), vec![CollabOp::SetCell {
        sheet: 99,
        sheet_name: "Gone".into(),
        row: 0,
        col: 0,
        content: CellContent::Value("x".into()),
    }]);
    let mut doc = Workbook::new();
    doc.sheet_mut(0).unwrap().set_value(5, 5, "replaced");
    c.replace_document(doc, 10);

    assert_eq!(c.stats.kept_across_replacement, 2);
    assert_eq!(c.stats.discarded_by_replacement, 2);
    assert_eq!(c.last_seen, 10);
    let sheet = &c.wb.sheets()[0];
    assert_eq!(sheet.get_raw(5, 5), "replaced");
    assert_eq!(sheet.get_raw(0, 0), "a");
    assert_eq!(sheet.get_raw(1, 0), "b");
    // Resent with fresh ids, based on the barrier.
    match c.poll_send() {
        Some(ToServer::Submit(env)) => {
            assert_eq!(env.base_seq, 10);
            assert_ne!(env.client_op_id, Uuid::from_u128(1));
        }
        other => panic!("expected a resend, got {other:?}"),
    }
}
