//! Intent tests for `SetFormat` against every op it can meet, plus the JSON
//! contract the web client codes against.

use visigrid_collab::apply::{apply_ops, fingerprint};
use visigrid_collab::op::{ops_from_json, ops_to_json, Axis, CellContent, CollabOp, FormatProps, HAlign, Rect};
use visigrid_collab::transform::{transform, transform_lists, Order, Transformed};
use visigrid_engine::cell::{Alignment, NumberFormat, TextOverflow};
use visigrid_engine::workbook::Workbook;

fn key(wb: &Workbook) -> u64 {
    wb.sheets()[0].id.0
}

fn fmt(wb: &Workbook, rect: Rect, props: FormatProps) -> CollabOp {
    CollabOp::SetFormat { sheet: key(wb), rect, props }
}

fn rows(wb: &Workbook, at: usize, count: usize, delete: bool) -> CollabOp {
    CollabOp::Structural {
        sheet: key(wb),
        sheet_name: wb.sheets()[0].name.clone(),
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

/// TP1 for one pair from an empty workbook (plus `setup`); returns the result.
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
fn overlapping_formats_merge_per_property_and_the_later_wins_each_one() {
    let wb = Workbook::new();
    let red_bold = FormatProps {
        bold: Some(Some(true)),
        color: Some(Some("#FF0000".into())),
        ..Default::default()
    };
    let blue_italic = FormatProps {
        italic: Some(Some(true)),
        color: Some(Some("#0000FF".into())),
        ..Default::default()
    };
    let a = fmt(&wb, Rect::new(0, 0, 1, 1), red_bold); // sequenced first
    let b = fmt(&wb, Rect::new(1, 1, 2, 2), blue_italic); // later
    let out = converge(&[], &b, &a);
    let s = &out.sheets()[0];
    let overlap = s.get_format(1, 1);
    assert!(overlap.bold, "the earlier op's bold survives where the later op says nothing");
    assert!(overlap.italic);
    assert_eq!(overlap.font_color, Some([0, 0, 255, 255]), "the later op's colour wins");
    assert_eq!(s.get_format(0, 0).font_color, Some([255, 0, 0, 255]));
    assert!(!s.get_format(2, 2).bold);
}

#[test]
fn a_format_across_an_insert_splits_and_skips_the_new_rows() {
    let wb = Workbook::new();
    let f = fmt(&wb, Rect::new(1, 0, 4, 0), FormatProps { wrap: Some(Some(true)), ..Default::default() });
    let insert = rows(&wb, 3, 2, false);
    assert_eq!(ops(transform(&f, &insert, Order::Later)).len(), 2);
    let out = converge(&[], &f, &insert);
    let s = &out.sheets()[0];
    assert_eq!(s.get_format(3, 0).text_overflow, TextOverflow::Clip, "inserted rows start unformatted");
    assert_eq!(s.get_format(5, 0).text_overflow, TextOverflow::Wrap, "the original row 4 moved to 6 (index 5)");
}

#[test]
fn a_format_on_deleted_rows_shrinks_or_drops() {
    let wb = Workbook::new();
    let f = fmt(&wb, Rect::new(2, 0, 3, 0), FormatProps::bold(true));
    let delete = rows(&wb, 2, 2, true);
    assert!(matches!(transform(&f, &delete, Order::Later), Transformed::Dropped(_)));
    let wider = fmt(&wb, Rect::new(1, 0, 5, 0), FormatProps::bold(true));
    converge(&[], &wider, &delete);
}

#[test]
fn values_and_formats_commute() {
    let wb = Workbook::new();
    let write = CollabOp::SetCell {
        sheet: key(&wb),
        sheet_name: wb.sheets()[0].name.clone(),
        row: 0,
        col: 0,
        content: CellContent::Value("0.5".into()),
    };
    let pct = fmt(&wb, Rect::new(0, 0, 0, 0), FormatProps { number_format: Some(Some("0%".into())), ..Default::default() });
    let out = converge(&[], &write, &pct);
    assert_eq!(out.sheets()[0].get_formatted_display(0, 0), "50%");
    let clear = CollabOp::SetCell {
        sheet: key(&wb),
        sheet_name: wb.sheets()[0].name.clone(),
        row: 0,
        col: 0,
        content: CellContent::Clear,
    };
    let out = converge(&[write.clone(), pct.clone()], &clear, &fmt(&wb, Rect::new(0, 0, 0, 0), FormatProps::bold(true)));
    let f = out.sheets()[0].get_format(0, 0);
    assert!(f.bold);
    assert!(matches!(f.number_format, NumberFormat::Custom(ref c) if c == "0%"), "clear contents keeps the format");
}

#[test]
fn an_atomic_range_refuses_an_overlapping_format() {
    let wb = Workbook::new();
    let replace = CollabOp::ReplaceRange {
        sheet: key(&wb),
        row: 0,
        col: 0,
        values: vec![vec![CellContent::Value("1".into())]],
    };
    let f = fmt(&wb, Rect::new(0, 0, 1, 1), FormatProps::bold(true));
    assert!(matches!(transform(&f, &replace, Order::Later), Transformed::Refused(_)));
}

#[test]
fn null_clears_absent_keeps_and_legacy_bold_still_applies() {
    let mut wb = Workbook::new();
    let k = key(&wb);
    let set = ops_from_json(&serde_json::json!([{"SetFormat": {"sheet": k, "rect": {"r0":0,"c0":0,"r1":0,"c1":0},
        "props": {"bold": true, "h_align": "center", "color": "#00aa00", "number_format": "#,##0"}}}]))
    .unwrap();
    apply_ops(&mut wb, &set);
    let clear = ops_from_json(&serde_json::json!({"SetFormat": {"sheet": k, "rect": {"r0":0,"c0":0,"r1":0,"c1":0},
        "props": {"h_align": null, "color": null}}}))
    .unwrap();
    apply_ops(&mut wb, &clear);
    let f = wb.sheets()[0].get_format(0, 0);
    assert!(f.bold, "absent = unchanged");
    assert_eq!(f.alignment, Alignment::General, "null = cleared");
    assert_eq!(f.font_color, None);
    // Round trip keeps null vs absent.
    let json = ops_to_json(&clear);
    assert_eq!(json[0]["SetFormat"]["props"], serde_json::json!({"color": null, "h_align": null}));
    // SetBold from an older log is the same as SetFormat { bold }.
    let legacy = CollabOp::SetBold { sheet: k, rect: Rect::new(0, 0, 0, 0), bold: false };
    assert_eq!(legacy.normalized(), CollabOp::SetFormat { sheet: k, rect: Rect::new(0, 0, 0, 0), props: FormatProps::bold(false) });
    apply_ops(&mut wb, &[legacy]);
    assert!(!wb.sheets()[0].get_format(0, 0).bold);
    assert_eq!(HAlign::Center, serde_json::from_value(serde_json::json!("center")).unwrap());
}

#[test]
fn invalid_formats_are_rejected_at_the_wire() {
    let k = 1u64;
    for props in [
        serde_json::json!({"color": "red"}),
        serde_json::json!({"font_size": -3}),
        serde_json::json!({"shadow": true}),
        serde_json::json!({"h_align": "justify"}),
    ] {
        let v = serde_json::json!({"SetFormat": {"sheet": k, "rect": {"r0":0,"c0":0,"r1":0,"c1":0}, "props": props}});
        assert!(ops_from_json(&v).is_err(), "{v}");
    }
}
