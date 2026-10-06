//! One intent test per row of the spec's pair table: the transformed op does
//! what the user meant, not just "something consistent".

use visigrid_collab::apply::{apply_ops, fingerprint};
use visigrid_collab::op::{Axis, CellContent, CollabOp, Rect};
use visigrid_collab::transform::{transform, transform_lists, Order, Transformed};
use visigrid_engine::workbook::Workbook;

fn sheet1(wb: &Workbook) -> (u64, String) {
    let s = &wb.sheets()[0];
    (s.id.0, s.name.clone())
}

fn set(wb: &Workbook, row: usize, col: usize, raw: &str) -> CollabOp {
    let (sheet, sheet_name) = sheet1(wb);
    let content = if raw.is_empty() {
        CellContent::Clear
    } else if raw.starts_with('=') {
        CellContent::Formula(raw.into())
    } else {
        CellContent::Value(raw.into())
    };
    CollabOp::SetCell {
        sheet,
        sheet_name,
        row,
        col,
        content,
    }
}

fn rows(wb: &Workbook, at: usize, count: usize, delete: bool) -> CollabOp {
    let (sheet, sheet_name) = sheet1(wb);
    CollabOp::Structural {
        sheet,
        sheet_name,
        axis: Axis::Row,
        at,
        count,
        delete,
    }
}

fn ops(t: Transformed) -> Vec<CollabOp> {
    match t {
        Transformed::Ops(v) => v,
        other => panic!("expected ops, got {other:?}"),
    }
}

/// apply(S, b ++ T(a,b,Later)) == apply(S, a ++ T(b,a,Earlier)), and return
/// the converged workbook.
fn converge(setup: &[CollabOp], a: &CollabOp, b: &CollabOp) -> Workbook {
    let (a2, b2) = transform_lists(
        std::slice::from_ref(a),
        std::slice::from_ref(b),
        Order::Later,
    )
    .expect("not refused");
    let mut left = Workbook::new();
    apply_ops(&mut left, setup);
    apply_ops(&mut left, std::slice::from_ref(b));
    apply_ops(&mut left, &a2);
    let mut right = Workbook::new();
    apply_ops(&mut right, setup);
    apply_ops(&mut right, std::slice::from_ref(a));
    apply_ops(&mut right, &b2);
    assert_eq!(
        fingerprint(&left),
        fingerprint(&right),
        "TP1 for {a:?} / {b:?}"
    );
    left
}

#[test]
fn same_cell_later_write_wins() {
    let wb = Workbook::new();
    let a = set(&wb, 0, 0, "1");
    let b = set(&wb, 0, 0, "2");
    assert_eq!(ops(transform(&a, &b, Order::Later)), vec![a.clone()]);
    assert!(matches!(
        transform(&b, &a, Order::Earlier),
        Transformed::Dropped(_)
    ));
    let out = converge(&[], &a, &b);
    assert_eq!(out.sheets()[0].get_raw(0, 0), "1");
}

#[test]
fn edit_below_an_insert_lands_in_the_shifted_cell() {
    let wb = Workbook::new();
    let edit = set(&wb, 5, 0, "x"); // A6
    let insert = rows(&wb, 2, 2, false); // insert 2 rows at row 3
    let out = converge(&[set(&wb, 5, 0, "old")], &edit, &insert);
    assert_eq!(out.sheets()[0].get_raw(7, 0), "x"); // A8
    assert_eq!(out.sheets()[0].get_raw(5, 0), "");
}

#[test]
fn write_into_a_concurrently_deleted_row_is_dropped() {
    let wb = Workbook::new();
    let edit = set(&wb, 3, 1, "lost");
    let delete = rows(&wb, 2, 3, true);
    assert!(matches!(
        transform(&edit, &delete, Order::Later),
        Transformed::Dropped(_)
    ));
    let out = converge(&[], &edit, &delete);
    assert!(fingerprint(&out).sheets[0].cells.is_empty());
}

#[test]
fn formula_written_concurrently_with_an_insert_is_rewritten_like_stored_formulas() {
    let wb = Workbook::new();
    let f = set(&wb, 0, 1, "=A5*2");
    let insert = rows(&wb, 0, 1, false);
    match &ops(transform(&f, &insert, Order::Later))[0] {
        CollabOp::SetCell {
            row,
            content: CellContent::Formula(t),
            ..
        } => {
            assert_eq!(*row, 1);
            assert_eq!(t, "=A6*2");
        }
        other => panic!("{other:?}"),
    }
    converge(&[set(&wb, 4, 0, "21")], &f, &insert);
}

#[test]
fn formula_into_deleted_rows_becomes_ref_error() {
    let wb = Workbook::new();
    let f = set(&wb, 0, 1, "=A5+1");
    let delete = rows(&wb, 3, 2, true);
    match &ops(transform(&f, &delete, Order::Later))[0] {
        CollabOp::SetCell {
            content: CellContent::Formula(t),
            ..
        } => assert!(t.contains("#REF!"), "{t}"),
        other => panic!("{other:?}"),
    }
    converge(&[], &f, &delete);
}

#[test]
fn two_inserts_at_the_same_row_order_by_sequence() {
    let wb = Workbook::new();
    let earlier = rows(&wb, 3, 1, false);
    let later = rows(&wb, 3, 2, false);
    // The later insert lands below the earlier one.
    match &ops(transform(&later, &earlier, Order::Later))[0] {
        CollabOp::Structural { at, .. } => assert_eq!(*at, 4),
        other => panic!("{other:?}"),
    }
    match &ops(transform(&earlier, &later, Order::Earlier))[0] {
        CollabOp::Structural { at, .. } => assert_eq!(*at, 3),
        other => panic!("{other:?}"),
    }
    converge(&[set(&wb, 3, 0, "kept")], &later, &earlier);
}

#[test]
fn insert_touching_a_concurrent_delete_is_serialized() {
    // Inside the deleted block, at its first line, or at the line just after
    // it: formula ranges would be rewritten differently by the two orders.
    let wb = Workbook::new();
    let delete = rows(&wb, 2, 4, true); // rows 3..6
    for k in [2, 4, 6] {
        let insert = rows(&wb, k, 1, false);
        assert!(
            matches!(
                transform(&delete, &insert, Order::Later),
                Transformed::Refused(_)
            ),
            "k={k}"
        );
        assert!(
            matches!(
                transform(&insert, &delete, Order::Later),
                Transformed::Refused(_)
            ),
            "k={k}"
        );
    }
    // Away from the block the pair commutes, ranges included.
    let setup = vec![
        set(&wb, 9, 1, "=SUM(A1:A12)"),
        set(&wb, 10, 1, "=SUM(A4:A9)"),
    ];
    converge(&setup, &delete, &rows(&wb, 8, 2, false));
    converge(&setup, &delete, &rows(&wb, 1, 2, false));
}

/// The simulator's counterexample for the old split rule: SUM(C3:E6) became
/// C4:E6 in one order and C1:E6 in the other.
#[test]
fn range_boundary_counterexample_is_refused() {
    let wb = Workbook::new();
    let delete = rows(&wb, 0, 3, true);
    let insert = rows(&wb, 3, 3, false);
    assert!(matches!(
        transform(&insert, &delete, Order::Later),
        Transformed::Refused(_)
    ));
}

#[test]
fn overlapping_deletes_delete_each_row_once() {
    let wb = Workbook::new();
    let a = rows(&wb, 2, 3, true); // 2,3,4
    let b = rows(&wb, 3, 3, true); // 3,4,5
    let setup: Vec<CollabOp> = (0..8).map(|r| set(&wb, r, 0, &format!("r{r}"))).collect();
    let out = converge(&setup, &a, &b);
    let s = &out.sheets()[0];
    assert_eq!(s.get_raw(1, 0), "r1");
    assert_eq!(s.get_raw(2, 0), "r6");
}

#[test]
fn bold_range_across_an_insert_splits_and_skips_the_new_rows() {
    let wb = Workbook::new();
    let (sheet, _) = sheet1(&wb);
    let bold = CollabOp::SetBold {
        sheet,
        rect: Rect::new(1, 0, 4, 0),
        bold: true,
    };
    let insert = rows(&wb, 3, 2, false);
    assert_eq!(ops(transform(&bold, &insert, Order::Later)).len(), 2);
    converge(&[set(&wb, 1, 0, "a"), set(&wb, 4, 0, "b")], &bold, &insert);
}

#[test]
fn rename_updates_carried_names_and_leaves_formula_text_like_the_engine() {
    let wb = Workbook::new();
    let (sheet, _) = sheet1(&wb);
    let rename = CollabOp::RenameSheet {
        sheet,
        name: "Budget".into(),
    };
    let edit = set(&wb, 0, 0, "=B1+1");
    match &ops(transform(&edit, &rename, Order::Later))[0] {
        CollabOp::SetCell {
            sheet_name,
            content: CellContent::Formula(t),
            ..
        } => {
            assert_eq!(sheet_name, "Budget");
            assert_eq!(t, "=B1+1");
        }
        other => panic!("{other:?}"),
    }
    converge(&[], &edit, &rename);
}

#[test]
fn edits_on_a_concurrently_deleted_sheet_are_dropped() {
    let mut wb = Workbook::new();
    let add = CollabOp::AddSheet {
        sheet: 1 << 62,
        name: "Extra".into(),
        index: 1,
    };
    apply_ops(&mut wb, std::slice::from_ref(&add));
    let extra = &wb.sheets()[1];
    let edit = CollabOp::SetCell {
        sheet: extra.id.0,
        sheet_name: extra.name.clone(),
        row: 0,
        col: 0,
        content: CellContent::Value("1".into()),
    };
    let delete = CollabOp::DeleteSheet {
        sheet: extra.id.0,
        index: 1,
    };
    assert!(matches!(
        transform(&edit, &delete, Order::Later),
        Transformed::Dropped(_)
    ));
    converge(&[add], &edit, &delete);
}

#[test]
fn concurrent_sheet_adds_converge_on_tab_order() {
    let a = CollabOp::AddSheet {
        sheet: (1 << 62) | 1,
        name: "S1".into(),
        index: 1,
    };
    let b = CollabOp::AddSheet {
        sheet: (1 << 62) | 2,
        name: "S2".into(),
        index: 1,
    };
    let out = converge(&[], &a, &b);
    let names: Vec<_> = out.sheets().iter().map(|s| s.name.clone()).collect();
    // The earlier add (b) keeps index 1; the later one goes after it.
    assert_eq!(names[1..], ["S2".to_string(), "S1".to_string()]);
}

#[test]
fn same_sheet_name_is_refused_for_the_later_op() {
    let a = CollabOp::AddSheet {
        sheet: (1 << 62) | 1,
        name: "Plan".into(),
        index: 1,
    };
    let b = CollabOp::AddSheet {
        sheet: (1 << 62) | 2,
        name: "plan".into(),
        index: 1,
    };
    assert!(matches!(
        transform(&a, &b, Order::Later),
        Transformed::Refused(_)
    ));
    assert_eq!(
        transform(&b, &a, Order::Earlier),
        Transformed::Ops(vec![b.clone()])
    );
}

#[test]
fn concurrent_deletes_of_different_sheets_are_serialized() {
    let a = CollabOp::DeleteSheet { sheet: 1, index: 0 };
    let b = CollabOp::DeleteSheet { sheet: 2, index: 1 };
    assert!(matches!(
        transform(&a, &b, Order::Later),
        Transformed::Refused(_)
    ));
}

#[test]
fn atomic_range_refuses_overlapping_later_ops_and_shifts_otherwise() {
    let wb = Workbook::new();
    let (sheet, _) = sheet1(&wb);
    let block = CollabOp::ReplaceRange {
        sheet,
        row: 2,
        col: 0,
        values: vec![
            vec![CellContent::Value("1".into())],
            vec![CellContent::Value("2".into())],
        ],
    };
    let inside = set(&wb, 3, 0, "x");
    assert!(matches!(
        transform(&inside, &block, Order::Later),
        Transformed::Refused(_)
    ));
    assert!(matches!(
        transform(&block, &inside, Order::Later),
        Transformed::Refused(_)
    ));
    let outside = set(&wb, 7, 0, "y");
    converge(&[], &outside, &block);
    // An insert above the block shifts it.
    let insert = rows(&wb, 0, 1, false);
    match &ops(transform(&block, &insert, Order::Later))[0] {
        CollabOp::ReplaceRange { row, .. } => assert_eq!(*row, 3),
        other => panic!("{other:?}"),
    }
    converge(&[], &block, &insert);
    // An insert inside it is refused.
    let inside_insert = rows(&wb, 3, 1, false);
    assert!(matches!(
        transform(&inside_insert, &block, Order::Later),
        Transformed::Refused(_)
    ));
}
