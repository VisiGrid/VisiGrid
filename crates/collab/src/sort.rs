//! Computing a sort's row order from a replica's own state.
//!
//! A `SortRange` op carries its order rather than its keys, so a sort means
//! the same thing on every replica however concurrent edits interleave with
//! it. The writer computes that order here, from the computed values it sees.
//!
//! Order within a key, as Sheets sorts: numbers, then text (ignoring case),
//! then booleans, then errors. Blank cells go last in both directions. Rows
//! that tie on every key keep their order (the sort is stable, and so is its
//! descending form).

use std::cmp::Ordering;

use serde::{Deserialize, Serialize};
use visigrid_engine::filter::{FilterKey, NormalizedFilterKey};
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;

use crate::op::{CollabOp, Rect, SheetKey, MAX_SORT_ROWS};

/// One sort key: an absolute column (inside the sorted rectangle) and its direction.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct SortKey {
    pub col: usize,
    pub ascending: bool,
}

/// Why no sort op was produced.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SortError {
    NoSuchSheet,
    /// No key, or a key column outside the rectangle.
    BadKeys,
    TooManyRows,
    /// Merged cells inside the rectangle (Sheets refuses these too).
    Merged,
}

impl SortError {
    pub fn message(&self) -> &'static str {
        match self {
            SortError::NoSuchSheet => "That sheet no longer exists.",
            SortError::BadKeys => "Choose a column inside the range to sort by.",
            SortError::TooManyRows => "Sorting more than 30,000 rows at once isn't supported yet.",
            SortError::Merged => "Unmerge the cells in the range before sorting it.",
        }
    }
}

fn rank(k: &NormalizedFilterKey) -> u8 {
    match k {
        NormalizedFilterKey::Number(_) => 0,
        NormalizedFilterKey::Text(_) => 1,
        NormalizedFilterKey::Bool(_) => 2,
        NormalizedFilterKey::Error(_) => 3,
        NormalizedFilterKey::Blank => 4,
    }
}

fn compare(a: &NormalizedFilterKey, b: &NormalizedFilterKey, ascending: bool) -> Ordering {
    // Blanks last whichever the direction.
    match (matches!(a, NormalizedFilterKey::Blank), matches!(b, NormalizedFilterKey::Blank)) {
        (true, true) => return Ordering::Equal,
        (true, false) => return Ordering::Greater,
        (false, true) => return Ordering::Less,
        _ => {}
    }
    let o = rank(a).cmp(&rank(b)).then_with(|| a.cmp(b));
    if ascending {
        o
    } else {
        o.reverse()
    }
}

/// The `SortRange` that sorts `rect` of `sheet` by `keys` as `wb` stands.
/// `None` when the rows are already in order (nothing to send).
pub fn sort_op(wb: &Workbook, sheet: SheetKey, rect: Rect, keys: &[SortKey]) -> Result<Option<CollabOp>, SortError> {
    let idx = wb.idx_for_sheet_id(SheetId(sheet)).ok_or(SortError::NoSuchSheet)?;
    if keys.is_empty() || keys.iter().any(|k| k.col < rect.c0 || k.col > rect.c1) || rect.r0 > rect.r1 {
        return Err(SortError::BadKeys);
    }
    let h = rect.r1 - rect.r0 + 1;
    if h > MAX_SORT_ROWS {
        return Err(SortError::TooManyRows);
    }
    let s = &wb.sheets()[idx];
    if s.merged_regions.iter().any(|m| rect.intersects(&Rect::new(m.start.0, m.start.1, m.end.0, m.end.1))) {
        return Err(SortError::Merged);
    }
    let values: Vec<Vec<NormalizedFilterKey>> = (0..h)
        .map(|i| keys.iter().map(|k| FilterKey::from_value(&s.get_computed_value(rect.r0 + i, k.col)).normalized()).collect())
        .collect();
    let mut order: Vec<u32> = (0..h as u32).collect();
    order.sort_by(|&x, &y| {
        let (vx, vy) = (&values[x as usize], &values[y as usize]);
        keys.iter()
            .enumerate()
            .map(|(i, k)| compare(&vx[i], &vy[i], k.ascending))
            .find(|o| o.is_ne())
            .unwrap_or(Ordering::Equal)
    });
    if order.iter().enumerate().all(|(i, &o)| i as u32 == o) {
        return Ok(None);
    }
    Ok(Some(CollabOp::SortRange { sheet, rect, order }))
}

/// Whether row `r0` of `rect` looks like a header over `keys`' columns: as
/// Sheets guesses, text above a column whose next row is a number, or every
/// key cell text and bold. Callers offer this as the default, never force it.
pub fn looks_like_header(wb: &Workbook, sheet: SheetKey, rect: Rect, keys: &[SortKey]) -> bool {
    let Some(idx) = wb.idx_for_sheet_id(SheetId(sheet)) else { return false };
    if rect.r1 <= rect.r0 {
        return false;
    }
    let s = &wb.sheets()[idx];
    let key = |r: usize, c: usize| FilterKey::from_value(&s.get_computed_value(r, c));
    let cols: Vec<usize> = (rect.c0..=rect.c1).collect();
    let all_text_bold = keys.iter().all(|k| matches!(key(rect.r0, k.col), FilterKey::Text(_)) && s.get_format(rect.r0, k.col).bold);
    let text_over_number = cols
        .iter()
        .any(|&c| matches!(key(rect.r0, c), FilterKey::Text(_)) && matches!(key(rect.r0 + 1, c), FilterKey::Number(_)));
    all_text_bold || text_over_number
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::apply::apply_ops;
    use crate::op::CellContent;

    fn sheet_with(rows: &[&[&str]]) -> (Workbook, SheetKey) {
        let mut wb = Workbook::new();
        let key = wb.sheets()[0].id.0;
        let ops: Vec<CollabOp> = rows
            .iter()
            .enumerate()
            .flat_map(|(r, line)| {
                line.iter().enumerate().map(move |(c, v)| CollabOp::SetCell {
                    sheet: key,
                    sheet_name: "Sheet1".into(),
                    row: r,
                    col: c,
                    content: if v.is_empty() { CellContent::Clear } else { CellContent::Value(v.to_string()) },
                })
            })
            .collect();
        apply_ops(&mut wb, &ops);
        (wb, key)
    }

    fn order(op: Option<CollabOp>) -> Vec<u32> {
        match op {
            Some(CollabOp::SortRange { order, .. }) => order,
            None => vec![],
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn numbers_then_text_then_blanks_both_ways() {
        let (wb, k) = sheet_with(&[&["pear"], &["10"], &[""], &["Apple"], &["2"]]);
        let rect = Rect::new(0, 0, 4, 0);
        let up = order(sort_op(&wb, k, rect, &[SortKey { col: 0, ascending: true }]).unwrap());
        assert_eq!(up, vec![4, 1, 3, 0, 2]);
        let down = order(sort_op(&wb, k, rect, &[SortKey { col: 0, ascending: false }]).unwrap());
        assert_eq!(down, vec![0, 3, 1, 4, 2], "blank stays last");
    }

    #[test]
    fn ties_keep_their_order_and_later_keys_break_them() {
        let (wb, k) = sheet_with(&[&["b", "2"], &["a", "9"], &["b", "1"], &["a", "3"]]);
        let rect = Rect::new(0, 0, 3, 1);
        let one = order(sort_op(&wb, k, rect, &[SortKey { col: 0, ascending: false }]).unwrap());
        assert_eq!(one, vec![0, 2, 1, 3]);
        let two = order(
            sort_op(&wb, k, rect, &[SortKey { col: 0, ascending: true }, SortKey { col: 1, ascending: true }]).unwrap(),
        );
        assert_eq!(two, vec![3, 1, 2, 0]);
    }

    #[test]
    fn already_sorted_is_nothing_to_send() {
        let (wb, k) = sheet_with(&[&["1"], &["2"]]);
        assert_eq!(sort_op(&wb, k, Rect::new(0, 0, 1, 0), &[SortKey { col: 0, ascending: true }]), Ok(None));
    }

    #[test]
    fn refuses_keys_outside_and_merges() {
        let (mut wb, k) = sheet_with(&[&["1", "x"], &["2", "y"]]);
        let rect = Rect::new(0, 0, 1, 0);
        assert_eq!(sort_op(&wb, k, rect, &[SortKey { col: 1, ascending: true }]), Err(SortError::BadKeys));
        apply_ops(&mut wb, &[CollabOp::Merge { sheet: k, rect: Rect::new(0, 0, 0, 1) }]);
        assert_eq!(sort_op(&wb, k, rect, &[SortKey { col: 0, ascending: true }]), Err(SortError::Merged));
    }

    #[test]
    fn guesses_a_header_over_numbers() {
        let (wb, k) = sheet_with(&[&["Name", "Score"], &["Ana", "3"], &["Bo", "1"]]);
        let keys = [SortKey { col: 1, ascending: true }];
        assert!(looks_like_header(&wb, k, Rect::new(0, 0, 2, 1), &keys));
        assert!(!looks_like_header(&wb, k, Rect::new(1, 0, 2, 1), &keys));
    }
}
