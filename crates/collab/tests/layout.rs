//! Intent and undo tests for line layout, freeze, borders and merges.

use uuid::Uuid;
use visigrid_collab::apply::{apply_ops, fingerprint};
use visigrid_collab::op::{Axis, BorderLine, BorderSpec, CellContent, CollabOp, FormatProps, LineProps, Rect};
use visigrid_collab::transform::{transform_lists, Order};
use visigrid_collab::undo::{record, resolve};
use visigrid_engine::workbook::Workbook;

fn key(wb: &Workbook) -> u64 {
    wb.sheets()[0].id.0
}

fn lines(wb: &Workbook, axis: Axis, lo: usize, hi: usize, size: Option<f64>, hidden: Option<bool>) -> CollabOp {
    CollabOp::SetLines { sheet: key(wb), axis, lo, hi, props: LineProps { size: size.map(Some), hidden } }
}

fn rows(wb: &Workbook, at: usize, count: usize, delete: bool) -> CollabOp {
    CollabOp::Structural { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), axis: Axis::Row, at, count, delete }
}

/// Both orders of a concurrent pair end the same; returns the result.
fn converge(setup: &[CollabOp], a: &CollabOp, b: &CollabOp) -> Workbook {
    let (a2, b2) = transform_lists(std::slice::from_ref(a), std::slice::from_ref(b), Order::Later).expect("not refused");
    let mut left = Workbook::new();
    apply_ops(&mut left, setup);
    let mut right = left.clone();
    apply_ops(&mut left, std::slice::from_ref(b));
    apply_ops(&mut left, &a2);
    apply_ops(&mut right, std::slice::from_ref(a));
    apply_ops(&mut right, &b2);
    assert_eq!(fingerprint(&left), fingerprint(&right), "TP1");
    left
}

#[test]
fn concurrent_line_changes_merge_per_property() {
    let wb = Workbook::new();
    let a = lines(&wb, Axis::Col, 0, 3, Some(120.0), None);
    let b = lines(&wb, Axis::Col, 2, 5, Some(40.0), Some(true));
    let out = converge(&[], &a, &b);
    let l = &out.sheets()[0].layout;
    assert_eq!(l.col_widths.get(&0), Some(&120.0));
    assert_eq!(l.col_widths.get(&2), Some(&120.0), "the later op (a) wins the overlap's width");
    assert_eq!(l.col_widths.get(&5), Some(&40.0));
    assert!(l.hidden_cols.contains(&3) && l.hidden_cols.contains(&5), "b's hiding survives where a set no visibility");
}

#[test]
fn lines_follow_rows_inserted_and_deleted_concurrently() {
    let wb = Workbook::new();
    let out = converge(&[], &lines(&wb, Axis::Row, 2, 4, Some(60.0), None), &rows(&wb, 3, 2, false));
    let h = &out.sheets()[0].layout.row_heights;
    assert_eq!(h.keys().copied().collect::<Vec<_>>(), vec![2, 5, 6], "inserted rows keep the default height");
    let out = converge(&[], &lines(&wb, Axis::Row, 2, 4, None, Some(true)), &rows(&wb, 1, 2, true));
    assert_eq!(out.sheets()[0].layout.hidden_rows.iter().copied().collect::<Vec<_>>(), vec![1, 2]);
}

#[test]
fn freeze_follows_structural_edits_and_the_later_freeze_wins() {
    let wb = Workbook::new();
    let k = key(&wb);
    let out = converge(&[], &CollabOp::SetFreeze { sheet: k, rows: 3, cols: 1 }, &rows(&wb, 0, 2, false));
    assert_eq!(out.sheets()[0].layout.frozen_rows, 5);
    let out = converge(&[], &CollabOp::SetFreeze { sheet: k, rows: 1, cols: 0 }, &CollabOp::SetFreeze { sheet: k, rows: 4, cols: 2 });
    assert_eq!((out.sheets()[0].layout.frozen_rows, out.sheets()[0].layout.frozen_cols), (1, 0));
}

#[test]
fn borders_are_format_properties() {
    let wb = Workbook::new();
    let thick = BorderSpec { style: BorderLine::Thick, color: Some("#3355FF".into()) };
    let a = CollabOp::SetFormat { sheet: key(&wb), rect: Rect::new(0, 0, 1, 1), props: FormatProps { border_bottom: Some(Some(thick)), ..Default::default() } };
    let b = CollabOp::SetFormat { sheet: key(&wb), rect: Rect::new(1, 1, 2, 2), props: FormatProps { border_bottom: Some(None), bold: Some(Some(true)), ..Default::default() } };
    let out = converge(&[], &a, &b);
    let s = &out.sheets()[0];
    assert_eq!(s.get_format(1, 1).border_bottom.color, Some([0x33, 0x55, 0xFF, 255]), "the later op wins the overlap");
    assert!(s.get_format(1, 1).bold);
}

#[test]
fn merges_are_serialized_against_writes_inside_them_and_structural_edits() {
    let wb = Workbook::new();
    let k = key(&wb);
    let merge = CollabOp::Merge { sheet: k, rect: Rect::new(1, 1, 2, 2) };
    let write = CollabOp::SetCell { sheet: k, sheet_name: "Sheet1".into(), row: 2, col: 2, content: CellContent::Value("x".into()) };
    assert!(transform_lists(&[merge.clone()], &[write.clone()], Order::Later).is_err());
    assert!(transform_lists(&[write], &[merge.clone()], Order::Later).is_err());
    assert!(transform_lists(&[merge.clone()], &[rows(&wb, 9, 1, false)], Order::Later).is_err());
    let outside = CollabOp::SetCell { sheet: k, sheet_name: "Sheet1".into(), row: 5, col: 5, content: CellContent::Value("y".into()) };
    let out = converge(&[], &merge, &outside);
    assert_eq!(out.sheets()[0].merged_regions.len(), 1);
}

#[test]
fn every_layout_op_undoes() {
    let wb0 = Workbook::new();
    let k = key(&wb0);
    let mut wb = wb0.clone();
    apply_ops(&mut wb, &[lines(&wb0, Axis::Col, 1, 1, Some(80.0), None), CollabOp::SetFreeze { sheet: k, rows: 1, cols: 0 }, CollabOp::Merge { sheet: k, rect: Rect::new(4, 0, 4, 2) }]);
    let start = fingerprint(&wb);
    let cases: Vec<Vec<CollabOp>> = vec![
        vec![lines(&wb0, Axis::Col, 0, 3, Some(150.0), Some(true))],
        vec![lines(&wb0, Axis::Row, 0, 9, None, Some(true))],
        vec![CollabOp::SetFreeze { sheet: k, rows: 3, cols: 2 }],
        vec![CollabOp::Merge { sheet: k, rect: Rect::new(0, 0, 1, 1) }],
        vec![CollabOp::Unmerge { sheet: k, rect: Rect::new(4, 1, 4, 1) }],
        vec![CollabOp::SetFormat { sheet: k, rect: Rect::new(0, 0, 2, 2), props: FormatProps { border_top: Some(Some(BorderSpec { style: BorderLine::Thin, color: None })), ..Default::default() } }],
    ];
    for ops in cases {
        let mut w = wb.clone();
        let entry = record(Uuid::nil(), &ops, &w);
        apply_ops(&mut w, &ops);
        let (undo, _) = resolve(&entry, &w);
        apply_ops(&mut w, &undo);
        assert_eq!(fingerprint(&w), start, "undo of {ops:?}");
    }
}
