//! Cell notes as operations: they move with their cell, survive what
//! happens to its content, converge, undo, and copy with their sheet.

use uuid::Uuid;
use visigrid_collab::apply::{apply_ops, collab_checksum, fingerprint};
use visigrid_collab::copy::sheet_ops;
use visigrid_collab::op::{Axis, CellContent, CollabOp, Rect};
use visigrid_collab::transform::{transform_lists, Order};
use visigrid_collab::undo::{record, resolve};
use visigrid_collab::view::{comment_at, comments, viewport};
use visigrid_engine::cell::CellComment;
use visigrid_engine::workbook::Workbook;

fn key(wb: &Workbook) -> u64 {
    wb.sheets()[0].id.0
}

fn note(wb: &Workbook, row: usize, col: usize, text: &str) -> CollabOp {
    CollabOp::SetComment { sheet: key(wb), row, col, comment: Some(CellComment { text: text.into(), author: "ana".into() }) }
}

fn unnote(wb: &Workbook, row: usize, col: usize) -> CollabOp {
    CollabOp::SetComment { sheet: key(wb), row, col, comment: None }
}

fn set(wb: &Workbook, row: usize, col: usize, content: CellContent) -> CollabOp {
    CollabOp::SetCell { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), row, col, content }
}

fn rows(wb: &Workbook, at: usize, count: usize, delete: bool) -> CollabOp {
    CollabOp::Structural { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), axis: Axis::Row, at, count, delete }
}

fn text_at(wb: &Workbook, row: usize, col: usize) -> Option<String> {
    wb.sheets()[0].comment(row, col).map(|n| n.text.clone())
}

/// Both orders of a concurrent pair, from `base`, end the same; returns the result.
fn converge(base: &Workbook, a: &CollabOp, b: &CollabOp) -> Workbook {
    let (a2, b2) = transform_lists(std::slice::from_ref(a), std::slice::from_ref(b), Order::Later).expect("not refused");
    let mut left = base.clone();
    let mut right = base.clone();
    apply_ops(&mut left, std::slice::from_ref(b));
    apply_ops(&mut left, &a2);
    apply_ops(&mut right, std::slice::from_ref(a));
    apply_ops(&mut right, &b2);
    assert_eq!(fingerprint(&left), fingerprint(&right), "TP1");
    left
}

#[test]
fn a_note_survives_clearing_and_overwriting_its_cell() {
    let mut wb = Workbook::new();
    let ops = [set(&wb, 0, 0, CellContent::Value("5".into())), note(&wb, 0, 0, "Q3 estimate")];
    apply_ops(&mut wb, &ops);
    { let ops = [set(&wb, 0, 0, CellContent::Value("6".into()))]; apply_ops(&mut wb, &ops); }
    assert_eq!(text_at(&wb, 0, 0).as_deref(), Some("Q3 estimate"));
    { let ops = [set(&wb, 0, 0, CellContent::Clear)]; apply_ops(&mut wb, &ops); }
    assert_eq!(text_at(&wb, 0, 0).as_deref(), Some("Q3 estimate"), "clear contents keeps the note, as in Excel");
    { let ops = [unnote(&wb, 0, 0)]; apply_ops(&mut wb, &ops); }
    assert_eq!(text_at(&wb, 0, 0), None);
    assert_eq!(wb.sheets()[0].cells_iter().count(), 0, "removing the last thing a cell had leaves no empty cell behind");
}

#[test]
fn a_note_follows_its_cell_through_a_concurrent_row_insert_or_delete() {
    let wb = Workbook::new();
    let out = converge(&wb, &note(&wb, 4, 1, "check"), &rows(&wb, 2, 3, false));
    assert_eq!(text_at(&out, 7, 1).as_deref(), Some("check"), "three rows inserted above: it moves down");
    let out = converge(&wb, &note(&wb, 4, 1, "check"), &rows(&wb, 3, 2, true));
    assert_eq!(fingerprint(&out).sheets[0].comments, vec![], "its row was deleted: the note goes with it");
    // And a note already in the sheet moves with a row insert like its cell.
    let mut wb = Workbook::new();
    { let ops = [note(&wb, 4, 1, "check")]; apply_ops(&mut wb, &ops); }
    { let ops = [rows(&wb, 0, 1, false)]; apply_ops(&mut wb, &ops); }
    assert_eq!(text_at(&wb, 5, 1).as_deref(), Some("check"));
}

#[test]
fn the_later_of_two_notes_on_one_cell_wins_on_every_replica() {
    let wb = Workbook::new();
    let out = converge(&wb, &note(&wb, 0, 0, "mine"), &note(&wb, 0, 0, "theirs"));
    assert_eq!(text_at(&out, 0, 0).as_deref(), Some("mine"), "a is sequenced later");
    let out = converge(&wb, &unnote(&wb, 0, 0), &note(&wb, 0, 0, "theirs"));
    assert_eq!(text_at(&out, 0, 0), None);
    let out = converge(&wb, &note(&wb, 0, 0, "a"), &note(&wb, 0, 1, "b"));
    assert_eq!((text_at(&out, 0, 0).as_deref(), text_at(&out, 0, 1).as_deref()), (Some("a"), Some("b")));
}

#[test]
fn a_note_goes_with_its_row_when_the_range_is_sorted() {
    let mut base = Workbook::new();
    let setup = [
        set(&base, 0, 0, CellContent::Value("3".into())),
        set(&base, 1, 0, CellContent::Value("1".into())),
        set(&base, 2, 0, CellContent::Value("2".into())),
        note(&base, 0, 1, "three"),
    ];
    apply_ops(&mut base, &setup);
    let sort = visigrid_collab::sort::sort_op(
        &base,
        key(&base),
        Rect::new(0, 0, 2, 1),
        &[visigrid_collab::sort::SortKey { col: 0, ascending: true }],
    )
    .unwrap()
    .expect("out of order");
    let mut wb = base.clone();
    apply_ops(&mut wb, std::slice::from_ref(&sort));
    assert_eq!(wb.sheets()[0].get_raw(2, 0), "3");
    assert_eq!(text_at(&wb, 2, 1).as_deref(), Some("three"), "the note moved with its row");
    assert_eq!(text_at(&wb, 0, 1), None);
    // A note written concurrently into a sorted row follows the row too.
    let out = converge(&base, &note(&base, 1, 1, "one"), &sort);
    assert_eq!(text_at(&out, 0, 1).as_deref(), Some("one"));
    assert_eq!(text_at(&out, 2, 1).as_deref(), Some("three"));
}

#[test]
fn undo_puts_back_a_note_and_leaves_someone_elses_change_alone() {
    let mut wb = Workbook::new();
    { let ops = [note(&wb, 0, 0, "old")]; apply_ops(&mut wb, &ops); }
    let before = fingerprint(&wb);
    let change = vec![note(&wb, 0, 0, "new"), note(&wb, 3, 2, "added")];
    let entry = record(Uuid::nil(), &change, &wb);
    apply_ops(&mut wb, &change);
    let (inverse, kept) = resolve(&entry, &wb);
    assert_eq!(kept, 0);
    let mut undone = wb.clone();
    apply_ops(&mut undone, &inverse);
    assert_eq!(fingerprint(&undone), before);
    // Someone else edits the first note before the undo: theirs stays.
    { let ops = [note(&wb, 0, 0, "someone else")]; apply_ops(&mut wb, &ops); }
    let (inverse, kept) = resolve(&entry, &wb);
    assert_eq!(kept, 1);
    apply_ops(&mut wb, &inverse);
    assert_eq!(text_at(&wb, 0, 0).as_deref(), Some("someone else"));
    assert_eq!(text_at(&wb, 3, 2), None);
}

#[test]
fn undoing_a_row_delete_brings_back_its_notes() {
    let mut wb = Workbook::new();
    { let ops = [note(&wb, 2, 0, "keep me"), set(&wb, 2, 1, CellContent::Value("9".into()))]; apply_ops(&mut wb, &ops); }
    let before = fingerprint(&wb);
    let del = vec![rows(&wb, 1, 3, true)];
    let entry = record(Uuid::nil(), &del, &wb);
    apply_ops(&mut wb, &del);
    assert_eq!(text_at(&wb, 2, 0), None);
    let (inverse, _) = resolve(&entry, &wb);
    apply_ops(&mut wb, &inverse);
    assert_eq!(fingerprint(&wb), before);
}

#[test]
fn notes_are_compared_for_convergence_and_none_keep_the_old_checksum() {
    let mut wb = Workbook::new();
    let print = serde_json::to_string(&fingerprint(&wb)).unwrap();
    assert!(!print.contains("comments"), "no notes: the same bytes as before notes were compared");
    let plain = collab_checksum(&wb);
    { let ops = [note(&wb, 0, 0, "x")]; apply_ops(&mut wb, &ops); }
    assert_ne!(collab_checksum(&wb), plain, "a replica missing a note would be caught");
}

#[test]
fn a_copied_sheet_keeps_its_notes() {
    let mut src = Workbook::new();
    { let ops = [note(&src, 0, 0, "a"), note(&src, 5, 3, "b")]; apply_ops(&mut src, &ops); }
    let (ops, dropped) = sheet_ops(&src.sheets()[0], 99, "Copy", 1);
    assert!(dropped.is_empty(), "{dropped:?}");
    let mut dst = Workbook::new();
    apply_ops(&mut dst, &ops);
    assert_eq!(comments(&dst, 99).unwrap(), comments(&src, key(&src)).unwrap());
}

#[test]
fn an_empty_or_huge_note_is_refused_before_any_replica_applies_it() {
    let wb = Workbook::new();
    assert!(note(&wb, 0, 0, "   ").validate().is_err(), "remove a note with null, not empty text");
    assert!(note(&wb, 0, 0, &"x".repeat(40_000)).validate().is_err());
    assert!(note(&wb, 0, 0, "fine").validate().is_ok());
    assert!(unnote(&wb, 0, 0).validate().is_ok());
}

#[test]
fn the_viewport_flags_noted_cells_and_the_reads_give_the_text() {
    let mut wb = Workbook::new();
    let ops = [set(&wb, 0, 0, CellContent::Value("1".into())), note(&wb, 0, 0, "on a value"), note(&wb, 4, 2, "on an empty cell")];
    apply_ops(&mut wb, &ops);
    let v = viewport(&wb, key(&wb), 0, 0, 9, 9).unwrap();
    let at = |i: usize| (v["rows"][i].as_u64().unwrap(), v["cols"][i].as_u64().unwrap());
    let flagged: Vec<(u64, u64)> = v["notes"].as_array().unwrap().iter().map(|i| at(i.as_u64().unwrap() as usize)).collect();
    assert_eq!(flagged, vec![(0, 0), (4, 2)], "an empty cell with a note is listed");
    assert_eq!(comment_at(&wb, key(&wb), 4, 2).unwrap()["text"], "on an empty cell");
    assert!(comment_at(&wb, key(&wb), 1, 1).is_none());
    let list = comments(&wb, key(&wb)).unwrap();
    assert_eq!(list.as_array().unwrap().len(), 2);
    assert_eq!(list[1]["row"], 4);
    let plain = viewport(&Workbook::new(), key(&wb), 0, 0, 9, 9).unwrap();
    assert!(plain.get("notes").is_none(), "no notes: the viewport is unchanged");
}
