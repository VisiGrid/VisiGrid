//! Conditional formats and validation as operations: concurrency, undo,
//! convergence and sheet copy.

use uuid::Uuid;
use visigrid_collab::apply::{apply_ops, collab_checksum, fingerprint};
use visigrid_collab::copy::sheet_ops;
use visigrid_collab::op::{Axis, CellContent, CollabOp};
use visigrid_collab::rules::{cond_formats_of, validation_of};
use visigrid_collab::transform::{transform, transform_lists, Order, Transformed};
use visigrid_collab::undo::{record, resolve};
use visigrid_engine::cell::CellStyle;
use visigrid_engine::cond_format::{CondFormatStore, CondStyle};
use visigrid_engine::sheet::{Sheet, SheetId};
use visigrid_engine::validation::{CellRange, ValidationRule};
use visigrid_engine::workbook::Workbook;

fn key(wb: &Workbook) -> u64 {
    wb.sheets()[0].id.0
}

/// A `SetCondFormats` that leaves `wb`'s first sheet with one rule: `predicate` over `range`.
fn cond(wb: &Workbook, range: CellRange, predicate: &str) -> CollabOp {
    let mut store = CondFormatStore::new();
    store.add(vec![range], predicate, CondStyle::Named(CellStyle::Error));
    CollabOp::SetCondFormats { sheet: key(wb), store: serde_json::to_value(&store).unwrap() }
}

/// A `SetValidation` that leaves `wb`'s first sheet with a Yes/No list over `range`.
fn dropdown(wb: &Workbook, range: CellRange) -> CollabOp {
    let mut s = Sheet::new(SheetId(1), 100, 10);
    s.validations.set(range, ValidationRule::list_inline(vec!["Yes".into(), "No".into()]));
    CollabOp::SetValidation { sheet: key(wb), rules: validation_of(&s) }
}

fn rows(wb: &Workbook, at: usize, count: usize, delete: bool) -> CollabOp {
    CollabOp::Structural { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), axis: Axis::Row, at, count, delete }
}

fn set(wb: &Workbook, row: usize, col: usize, text: &str) -> CollabOp {
    CollabOp::SetCell { sheet: key(wb), sheet_name: wb.sheets()[0].name.clone(), row, col, content: CellContent::Value(text.into()) }
}

/// Both orders of a concurrent pair end the same; returns the result.
fn converge(a: &CollabOp, b: &CollabOp) -> Workbook {
    let (a2, b2) = transform_lists(std::slice::from_ref(a), std::slice::from_ref(b), Order::Later).expect("not refused");
    let mut left = Workbook::new();
    let mut right = left.clone();
    apply_ops(&mut left, std::slice::from_ref(b));
    apply_ops(&mut left, &a2);
    apply_ops(&mut right, std::slice::from_ref(a));
    apply_ops(&mut right, &b2);
    assert_eq!(fingerprint(&left), fingerprint(&right), "TP1");
    left
}

fn refused(a: &CollabOp, b: &CollabOp) -> bool {
    matches!(transform(a, b, Order::Later), Transformed::Refused(_)) && matches!(transform(b, a, Order::Earlier), Transformed::Ops(_))
}

#[test]
fn a_rules_change_lands_on_every_replica_beside_cell_edits() {
    let wb = Workbook::new();
    let out = converge(&cond(&wb, CellRange::new(0, 0, 9, 0), "=A1>5"), &set(&wb, 2, 0, "7"));
    assert_eq!(out.sheets()[0].cond_formats.len(), 1);
    assert_eq!(out.sheets()[0].get_raw(2, 0), "7");
    let out = converge(&dropdown(&wb, CellRange::new(0, 1, 9, 1)), &cond(&wb, CellRange::new(0, 0, 0, 0), "=A1<>\"\""));
    assert!(out.sheets()[0].validations.get(3, 1).is_some(), "validation and conditional formats are separate");
    assert_eq!(out.sheets()[0].cond_formats.len(), 1);
}

#[test]
fn two_changes_to_the_same_rules_are_serialized_never_overwritten() {
    let wb = Workbook::new();
    let (a, b) = (cond(&wb, CellRange::new(0, 0, 3, 0), "=A1>1"), cond(&wb, CellRange::new(0, 1, 3, 1), "=B1>2"));
    assert!(refused(&a, &b), "the later is refused, so neither user's rule disappears unseen");
    let (a, b) = (dropdown(&wb, CellRange::new(0, 0, 3, 0)), dropdown(&wb, CellRange::new(5, 0, 6, 0)));
    assert!(refused(&a, &b));
}

#[test]
fn a_rules_change_is_serialized_against_row_edits_and_renames() {
    let wb = Workbook::new();
    let rule = cond(&wb, CellRange::new(2, 0, 5, 0), "=A3>0");
    assert!(refused(&rule, &rows(&wb, 1, 2, false)), "its ranges were written before the rows moved");
    assert!(refused(&rows(&wb, 1, 2, true), &rule));
    let rename = CollabOp::RenameSheet { sheet: key(&wb), name: "Data".into() };
    assert!(refused(&dropdown(&wb, CellRange::new(0, 0, 1, 0)), &rename));
}

#[test]
fn a_rules_change_on_a_deleted_sheet_is_dropped() {
    let mut wb = Workbook::new();
    apply_ops(&mut wb, &[CollabOp::AddSheet { sheet: 77, name: "Two".into(), index: 1 }]);
    let rule = CollabOp::SetCondFormats { sheet: 77, store: serde_json::to_value(CondFormatStore::new()).unwrap() };
    assert!(matches!(transform(&rule, &CollabOp::DeleteSheet { sheet: 77, index: 1 }, Order::Later), Transformed::Dropped(_)));
}

#[test]
fn one_undo_puts_back_the_rules_a_change_replaced() {
    let mut wb = Workbook::new();
    let setup = [cond(&wb, CellRange::new(0, 0, 3, 0), "=A1>1"), dropdown(&wb, CellRange::new(0, 1, 3, 1))];
    apply_ops(&mut wb, &setup);
    let before = fingerprint(&wb);
    let change = vec![cond(&wb, CellRange::new(5, 2, 6, 2), "=C6=0"), dropdown(&wb, CellRange::new(9, 0, 9, 0))];
    let entry = record(Uuid::nil(), &change, &wb);
    apply_ops(&mut wb, &change);
    assert_ne!(fingerprint(&wb), before);
    let (inverse, kept) = resolve(&entry, &wb);
    assert_eq!(kept, 0);
    apply_ops(&mut wb, &inverse);
    assert_eq!(fingerprint(&wb), before, "the rules are back exactly");
}

#[test]
fn rules_are_compared_for_convergence_and_absent_rules_keep_the_old_checksum() {
    let mut wb = Workbook::new();
    let print = serde_json::to_string(&fingerprint(&wb)).unwrap();
    assert!(!print.contains("cond_formats") && !print.contains("validation"), "no rules: the same bytes as before rules were compared");
    let plain = collab_checksum(&wb);
    let one = cond(&wb, CellRange::new(0, 0, 0, 0), "=A1>1");
    apply_ops(&mut wb, &[one]);
    assert_ne!(collab_checksum(&wb), plain, "a replica missing a rule would be caught");
}

#[test]
fn a_copied_sheet_keeps_its_rules() {
    let mut src = Workbook::new();
    let setup = [cond(&src, CellRange::new(0, 0, 4, 0), "=A1>3"), dropdown(&src, CellRange::new(0, 1, 4, 1))];
    apply_ops(&mut src, &setup);
    let s = &src.sheets()[0];
    let (ops, dropped) = sheet_ops(s, 99, "Copy", 1);
    assert!(dropped.is_empty(), "{dropped:?}");
    let mut dst = Workbook::new();
    apply_ops(&mut dst, &ops);
    let copy = dst.sheets().iter().find(|x| x.id.0 == 99).unwrap();
    assert_eq!(cond_formats_of(copy), cond_formats_of(s));
    assert_eq!(validation_of(copy), validation_of(s));
}

#[test]
fn a_malformed_rules_op_is_refused_before_any_replica_applies_it() {
    let wb = Workbook::new();
    let bad = CollabOp::SetValidation { sheet: key(&wb), rules: serde_json::json!({"validations": "nope"}) };
    assert!(bad.validate().is_err());
    let bad = CollabOp::SetCondFormats { sheet: key(&wb), store: serde_json::json!([1, 2]) };
    assert!(bad.validate().is_err());
    assert!(cond(&wb, CellRange::new(0, 0, 1, 1), "=A1>0").validate().is_ok());
}

#[test]
fn the_viewport_carries_what_conditional_formatting_draws() {
    use visigrid_collab::view::viewport;
    let mut wb = Workbook::new();
    let k = key(&wb);
    let setup = [set(&wb, 0, 0, "9"), set(&wb, 1, 0, "2")];
    apply_ops(&mut wb, &setup);
    let plain = viewport(&wb, k, 0, 0, 5, 3).unwrap();
    assert!(plain.get("cf").is_none(), "no rule, no cf: the viewport is unchanged");
    // Over 5 (row 0), and empty (rows 2-4) for a second rule.
    let mut store = CondFormatStore::new();
    store.add(vec![CellRange::new(0, 0, 4, 0)], "=A1>5", CondStyle::Named(CellStyle::Error));
    store.add(vec![CellRange::new(0, 0, 4, 0)], "=A1=\"\"", CondStyle::Named(CellStyle::Note));
    apply_ops(&mut wb, &[CollabOp::SetCondFormats { sheet: k, store: serde_json::to_value(&store).unwrap() }]);
    let v = viewport(&wb, k, 0, 0, 5, 3).unwrap();
    let at = |r: u64| v["rows"].as_array().unwrap().iter().zip(v["cols"].as_array().unwrap()).position(|(x, c)| x == r && c == 0);
    let cf = |r: u64| at(r).map(|i| v["cf"][i].as_i64().unwrap());
    assert!(cf(0).unwrap() >= 0, "9 > 5: styled");
    assert_eq!(cf(1), Some(-1), "2 is listed (it has content) but not styled");
    assert!(cf(3).unwrap() >= 0, "an empty cell the second rule styles is listed");
    assert_eq!(v["fmt"][at(0).unwrap()], serde_json::json!(-1), "fmt stays the cell's own format");
}

#[test]
fn a_cells_dropdown_and_input_check_come_from_its_validation() {
    use visigrid_collab::rules::{check_input, validation_at};
    let mut wb = Workbook::new();
    let k = key(&wb);
    let setup = [dropdown(&wb, CellRange::new(0, 1, 9, 1))];
    apply_ops(&mut wb, &setup);
    let v = validation_at(&wb, k, 2, 1).expect("the cell has a rule");
    assert_eq!(v["items"], serde_json::json!(["Yes", "No"]));
    assert_eq!(v["dropdown"], serde_json::json!(true));
    assert!(validation_at(&wb, k, 2, 0).is_none(), "no rule outside the range");
    assert_eq!(check_input(&wb, k, 2, 1, "Yes")["valid"], serde_json::json!(true));
    let bad = check_input(&wb, k, 2, 1, "Maybe");
    assert_eq!(bad["valid"], serde_json::json!(false), "{bad}");
    assert!(bad["reason"].as_str().is_some_and(|r| !r.is_empty()));
}
