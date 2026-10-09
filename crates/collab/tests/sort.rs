//! Intent tests for `SortRange`: what a sort moves, how concurrent edits
//! follow their rows, what is serialized, undo, and the JSON contract the web
//! client codes against. Plus a TP1 sweep with a sort on one side of every
//! pair (the general fuzz generates sorts rarely).

use rand::rngs::StdRng;
use rand::{Rng, SeedableRng};
use uuid::Uuid;
use visigrid_collab::apply::{apply_ops, fingerprint, first_difference};
use visigrid_collab::gen::random_ops;
use visigrid_collab::op::{ops_from_json, Axis, CellContent, CollabOp, FormatProps, LineProps, Rect};
use visigrid_collab::sort::{sort_op, SortKey};
use visigrid_collab::transform::{transform, transform_lists, Order, Transformed};
use visigrid_collab::undo::{record, resolve};
use visigrid_engine::workbook::Workbook;

fn key(wb: &Workbook) -> u64 {
    wb.sheets()[0].id.0
}

fn set(wb: &Workbook, row: usize, col: usize, raw: &str) -> CollabOp {
    let content = if raw.is_empty() {
        CellContent::Clear
    } else if raw.starts_with('=') {
        CellContent::Formula(raw.into())
    } else {
        CellContent::Value(raw.into())
    };
    CollabOp::SetCell { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), row, col, content }
}

fn raw(wb: &Workbook, row: usize, col: usize) -> String {
    wb.sheets()[0].get_raw(row, col)
}

fn shown(wb: &Workbook, row: usize, col: usize) -> String {
    wb.sheets()[0].get_display(row, col)
}

/// Name | Score | Double, rows 1..=4, with a header row and a bold row.
fn scores() -> Workbook {
    let mut wb = Workbook::new();
    let rows = [["Name", "Score", ""], ["Cy", "30", "=B2*2"], ["Ana", "10", "=B3*2"], ["Bo", "20", "=B4*2"], ["Di", "", "=B5*2"]];
    let mut ops = Vec::new();
    for (r, line) in rows.iter().enumerate() {
        for (c, v) in line.iter().enumerate() {
            if !v.is_empty() {
                ops.push(set(&wb, r, c, v));
            }
        }
    }
    ops.push(CollabOp::SetFormat { sheet: key(&wb), rect: Rect::new(2, 0, 2, 0), props: FormatProps::bold(true) });
    apply_ops(&mut wb, &ops);
    wb
}

fn by_score(wb: &Workbook, ascending: bool) -> CollabOp {
    sort_op(wb, key(wb), Rect::new(1, 0, 4, 2), &[SortKey { col: 1, ascending }]).unwrap().expect("not already sorted")
}

/// TP1 for one pair against `wb`; returns the converged result.
fn converge(wb: &Workbook, a: &CollabOp, b: &CollabOp) -> Workbook {
    let (a2, b2) = transform_lists(std::slice::from_ref(a), std::slice::from_ref(b), Order::Later).expect("not refused");
    let mut left = wb.clone();
    let mut right = wb.clone();
    apply_ops(&mut left, std::slice::from_ref(b));
    apply_ops(&mut left, &a2);
    apply_ops(&mut right, std::slice::from_ref(a));
    apply_ops(&mut right, &b2);
    if let Some(d) = first_difference(&fingerprint(&left), &fingerprint(&right)) {
        panic!("TP1: {d}\n a={a:?}\n b={b:?}\n a'={a2:?}\n b'={b2:?}");
    }
    left
}

#[test]
fn moves_rows_with_their_formats_and_shifts_formulas_like_a_copy() {
    let mut wb = scores();
    let sort = by_score(&wb, true);
    // Ana 10, Bo 20, Cy 30, then Di (blank score) last.
    assert_eq!(sort, CollabOp::SortRange { sheet: key(&wb), rect: Rect::new(1, 0, 4, 2), order: vec![1, 2, 0, 3] });
    apply_ops(&mut wb, &[sort]);
    let names: Vec<String> = (1..=4).map(|r| raw(&wb, r, 0)).collect();
    assert_eq!(names, ["Ana", "Bo", "Cy", "Di"]);
    assert_eq!(raw(&wb, 0, 0), "Name", "the header row stays");
    // Each row's formula still doubles its own score.
    assert_eq!(raw(&wb, 1, 2), "=B2*2");
    assert_eq!(shown(&wb, 1, 2), "20");
    assert_eq!(shown(&wb, 3, 2), "60");
    // Ana's bold went with Ana.
    assert!(wb.sheets()[0].get_format(1, 0).bold);
    assert!(!wb.sheets()[0].get_format(2, 0).bold);
}

#[test]
fn leaves_columns_outside_and_row_heights_alone() {
    let mut wb = scores();
    let setup = [
        set(&wb, 1, 4, "beside Cy"),
        CollabOp::SetLines { sheet: key(&wb), axis: Axis::Row, lo: 1, hi: 1, props: LineProps { size: Some(Some(40.0)), hidden: None } },
    ];
    apply_ops(&mut wb, &setup);
    let sort = by_score(&wb, true);
    apply_ops(&mut wb, &[sort]);
    assert_eq!(raw(&wb, 1, 4), "beside Cy");
    assert_eq!(wb.sheets()[0].layout.row_heights.get(&1), Some(&40.0));
}

#[test]
fn a_concurrent_write_follows_its_row() {
    let wb = scores();
    let sort = by_score(&wb, true);
    // Someone types Cy's new score while the sort is in flight.
    let edit = set(&wb, 1, 1, "5");
    let out = converge(&wb, &edit, &sort);
    // The order was decided before the edit, so Cy keeps the third row.
    assert_eq!(raw(&out, 3, 0), "Cy");
    assert_eq!(raw(&out, 3, 1), "5");
    match transform(&edit, &sort, Order::Later) {
        Transformed::Ops(v) => assert_eq!(v, vec![set(&wb, 3, 1, "5")]),
        other => panic!("{other:?}"),
    }
    // Both orders, and a formula typed into a moving row.
    converge(&wb, &sort, &edit);
    converge(&wb, &set(&wb, 2, 3, "=B3+C3"), &sort);
    converge(&wb, &sort, &set(&wb, 4, 3, "=A5&B5"));
}

#[test]
fn a_concurrent_format_follows_its_rows() {
    let wb = scores();
    let sort = by_score(&wb, false);
    let red = FormatProps { color: Some(Some("#FF0000".into())), ..Default::default() };
    for rect in [Rect::new(0, 0, 2, 4), Rect::new(1, 1, 4, 1), Rect::new(3, 0, 6, 2)] {
        let f = CollabOp::SetFormat { sheet: key(&wb), rect, props: red.clone() };
        converge(&wb, &f, &sort);
        converge(&wb, &sort, &f);
    }
}

#[test]
fn serialized_against_structural_edits_overlapping_sorts_replaces_and_merges() {
    let wb = scores();
    let sort = by_score(&wb, true);
    let k = key(&wb);
    let refused = [
        CollabOp::Structural { sheet: k, sheet_name: "Sheet1".into(), axis: Axis::Row, at: 9, count: 1, delete: false },
        CollabOp::Structural { sheet: k + 1, sheet_name: "Other".into(), axis: Axis::Col, at: 0, count: 1, delete: true },
        CollabOp::SortRange { sheet: k, rect: Rect::new(3, 1, 5, 1), order: vec![2, 1, 0] },
        CollabOp::ReplaceRange { sheet: k, row: 4, col: 2, values: vec![vec![CellContent::Value("x".into())]] },
        CollabOp::Merge { sheet: k, rect: Rect::new(0, 0, 1, 1) },
    ];
    for other in refused {
        assert!(matches!(transform(&sort, &other, Order::Later), Transformed::Refused(_)), "{other:?}");
        assert!(matches!(transform(&other, &sort, Order::Later), Transformed::Refused(_)), "{other:?}");
    }
    // A sort of other rows, or another range beside it, goes through.
    let beside = CollabOp::SortRange { sheet: k, rect: Rect::new(6, 0, 7, 0), order: vec![1, 0] };
    converge(&wb, &beside, &sort);
}

#[test]
fn undo_puts_rows_back_with_later_edits_and_restores_refs() {
    let mut wb = scores();
    // Ana's formula points one row up; sorted to the top row it falls off
    // the sheet (#REF!), and undo must still bring back "=A1".
    let up = set(&wb, 2, 2, "=A1");
    apply_ops(&mut wb, &[up]);
    let start = fingerprint(&wb);
    let sort = by_score(&wb, true);
    let entry = record(Uuid::nil(), std::slice::from_ref(&sort), &wb);
    apply_ops(&mut wb, std::slice::from_ref(&sort));
    let (ops, kept) = resolve(&entry, &wb);
    assert_eq!(kept, 0);
    apply_ops(&mut wb, &ops);
    assert_eq!(fingerprint(&wb), start);
}

#[test]
fn json_is_what_the_web_client_sends() {
    let ops = ops_from_json(&serde_json::json!({ "SortRange": { "sheet": 1, "rect": { "r0": 1, "c0": 0, "r1": 3, "c1": 2 }, "order": [2, 0, 1] } })).unwrap();
    assert_eq!(ops, vec![CollabOp::SortRange { sheet: 1, rect: Rect::new(1, 0, 3, 2), order: vec![2, 0, 1] }]);
    for bad in [
        serde_json::json!({ "SortRange": { "sheet": 1, "rect": { "r0": 1, "c0": 0, "r1": 3, "c1": 2 }, "order": [0, 0, 1] } }),
        serde_json::json!({ "SortRange": { "sheet": 1, "rect": { "r0": 1, "c0": 0, "r1": 3, "c1": 2 }, "order": [0, 1] } }),
        serde_json::json!({ "SortRange": { "sheet": 1, "rect": { "r0": 3, "c0": 0, "r1": 1, "c1": 2 }, "order": [] } }),
    ] {
        assert!(ops_from_json(&bad).is_err(), "{bad}");
    }
}

/// TP1 with a sort on one side of every pair: random states from the fuzz
/// generator, a random sort of the hot region, and a random concurrent op.
#[test]
fn tp1_sorts_against_random_ops() {
    let n: u64 = std::env::var("COLLAB_FUZZ_SEEDS").ok().and_then(|s| s.parse().ok()).unwrap_or(3000);
    let mut checked = 0;
    for seed in 0..n {
        let mut rng = StdRng::seed_from_u64(seed ^ 0x5047);
        let mut wb = Workbook::new();
        // A block of data to sort (values, text, and formulas on their own
        // row or rows above), then random history on top.
        let mut fill = Vec::new();
        for r in 0..10 {
            for c in 0..5 {
                let raw = match rng.gen_range(0..10) {
                    0..=4 => rng.gen_range(-5..50).to_string(),
                    5 | 6 => ["pear", "Apple", "fig"][rng.gen_range(0..3)].to_string(),
                    7 => format!("={}{}*2", (b'A' + rng.gen_range(0..5u8)) as char, rng.gen_range(1..=r + 1)),
                    _ => continue,
                };
                fill.push(set(&wb, r, c, &raw));
            }
        }
        apply_ops(&mut wb, &fill);
        for _ in 0..rng.gen_range(2..10) {
            let k = (1 << 62) | rng.gen_range(0..1u64 << 40);
            let ops = random_ops(&mut rng, &wb, k);
            apply_ops(&mut wb, &ops);
        }
        let sheet = wb.sheets()[rng.gen_range(0..wb.sheets().len())].id.0;
        let (r0, c0) = (rng.gen_range(0..6), rng.gen_range(0..4));
        let rect = Rect::new(r0, c0, r0 + rng.gen_range(1..6), c0 + rng.gen_range(0..3));
        let keys = [SortKey { col: rng.gen_range(rect.c0..=rect.c1), ascending: rng.gen_bool(0.5) }];
        let Ok(Some(sort)) = sort_op(&wb, sheet, rect, &keys) else { continue };
        let k = (1 << 62) | rng.gen_range(0..1u64 << 40);
        let other = random_ops(&mut rng, &wb, k);
        let (a, b) = if rng.gen_bool(0.5) { (vec![sort], other) } else { (other, vec![sort]) };
        let Ok((a2, b2)) = transform_lists(&a, &b, Order::Later) else { continue };
        let mut left = wb.clone();
        apply_ops(&mut left, &b);
        apply_ops(&mut left, &a2);
        let mut right = wb.clone();
        apply_ops(&mut right, &a);
        apply_ops(&mut right, &b2);
        if let Some(d) = first_difference(&fingerprint(&left), &fingerprint(&right)) {
            left.recompute_full_ordered();
            right.recompute_full_ordered();
            // The fuzz classifies a difference a full recompute removes as the
            // engine's incremental recalc, not the transform.
            if first_difference(&fingerprint(&left), &fingerprint(&right)).is_some() {
                panic!("seed {seed}: {d}\n a={a:?}\n b={b:?}\n a'={a2:?}\n b'={b2:?}");
            }
        }
        checked += 1;
    }
    assert!(checked > n / 4, "only {checked} of {n} seeds produced a sort pair");
}
