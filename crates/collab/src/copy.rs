//! A whole sheet as operations: how a file is imported into a live workbook
//! and how a tab is duplicated, so either lands as one envelope (one change
//! and one undo step for its author) through the same sequencer, transforms
//! and undo as every other edit.
//!
//! The operation vocabulary carries contents, cell formats, row and column
//! sizes and visibility, frozen panes and merges. Anything else a sheet can
//! hold (conditional formats, validations, comments, tables, pivots, a tab
//! colour, formats on whole rows or columns beyond the cells that have them)
//! has no operation yet: [`sheet_ops`] reports it in `dropped` so the user is
//! told before it is lost.

use std::collections::BTreeMap;

use visigrid_engine::cell::CellFormat;
use visigrid_engine::sheet::Sheet;

use crate::op::{Axis, CellContent, CollabOp, LineProps, Rect, SheetKey};
use crate::undo::{content_at, per_cell_formats, props_of};

/// Contents per `ReplaceRange`: a long row is split so no single op is huge.
const RUN_CELLS: usize = 4_096;
/// Blank cells a content run may bridge rather than start a new op.
const GAP: usize = 2;

/// The ops that recreate `s` as a new sheet `key` named `name` at tab
/// position `index`, and a plain-language line for each thing they cannot
/// carry.
pub fn sheet_ops(s: &Sheet, key: SheetKey, name: &str, index: usize) -> (Vec<CollabOp>, Vec<String>) {
    let mut ops = vec![CollabOp::AddSheet { sheet: key, name: name.to_string(), index }];

    // Contents, row by row, in runs of nearby cells.
    let mut rows: BTreeMap<usize, Vec<(usize, CellContent)>> = BTreeMap::new();
    let mut formats = Vec::new();
    let default = CellFormat::default();
    for ((r, c), _) in s.cells_iter() {
        if !s.get_raw(r, c).is_empty() {
            rows.entry(r).or_default().push((c, content_at(s, r, c)));
        }
        let fmt = s.get_format(r, c);
        if fmt != default {
            formats.push((r, c, props_of(&fmt, None)));
        }
    }
    for (r, mut cells) in rows {
        cells.sort_by_key(|(c, _)| *c);
        let mut run: Vec<CellContent> = Vec::new();
        let mut start = 0;
        for (c, content) in cells {
            let end = start + run.len();
            if !run.is_empty() && (c > end + GAP || run.len() >= RUN_CELLS) {
                ops.push(CollabOp::ReplaceRange { sheet: key, row: r, col: start, values: vec![std::mem::take(&mut run)] });
            }
            if run.is_empty() {
                start = c;
            }
            while start + run.len() < c {
                run.push(CellContent::Clear);
            }
            run.push(content);
        }
        if !run.is_empty() {
            ops.push(CollabOp::ReplaceRange { sheet: key, row: r, col: start, values: vec![run] });
        }
    }
    ops.extend(per_cell_formats(key, formats));

    // Sizes and hidden lines, in runs of equal settings.
    let l = &s.layout;
    for (axis, sizes, hidden) in [(Axis::Col, &l.col_widths, &l.hidden_cols), (Axis::Row, &l.row_heights, &l.hidden_rows)] {
        let mut lines: BTreeMap<usize, LineProps> = BTreeMap::new();
        for (i, v) in sizes {
            lines.entry(*i).or_default().size = Some(Some(*v as f64));
        }
        for i in hidden {
            lines.entry(*i).or_default().hidden = Some(true);
        }
        let mut run: Option<(usize, usize, LineProps)> = None;
        for (i, props) in lines {
            match &mut run {
                Some((_, hi, p)) if *hi + 1 == i && *p == props => *hi = i,
                _ => {
                    if let Some((lo, hi, props)) = run.take() {
                        ops.push(CollabOp::SetLines { sheet: key, axis, lo, hi, props });
                    }
                    run = Some((i, i, props));
                }
            }
        }
        if let Some((lo, hi, props)) = run {
            ops.push(CollabOp::SetLines { sheet: key, axis, lo, hi, props });
        }
    }
    if l.frozen_rows > 0 || l.frozen_cols > 0 {
        ops.push(CollabOp::SetFreeze { sheet: key, rows: l.frozen_rows, cols: l.frozen_cols });
    }
    for m in &s.merged_regions {
        ops.push(CollabOp::Merge { sheet: key, rect: Rect::new(m.start.0, m.start.1, m.end.0, m.end.1) });
    }

    let mut dropped = Vec::new();
    let mut note = |n: usize, one: &str, many: &str| {
        if n > 0 {
            dropped.push(format!("{}: {}", s.name, if n == 1 { format!("1 {one}") } else { format!("{n} {many}") }));
        }
    };
    note(s.cond_formats.len(), "conditional format", "conditional formats");
    note(s.validations.len(), "data validation rule", "data validation rules");
    note(s.comments().count(), "comment", "comments");
    note(s.tables().len(), "table (its cells are kept)", "tables (their cells are kept)");
    note(s.pivots.len(), "pivot table definition (its cells are kept)", "pivot table definitions (their cells are kept)");
    note(s.row_formats.len() + s.col_formats.len(), "whole-row or whole-column format (kept on cells with data)", "whole-row or whole-column formats (kept on cells with data)");
    if s.tab_color.is_some() {
        dropped.push(format!("{}: the tab colour", s.name));
    }
    (ops, dropped)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::apply::{apply_ops, collab_checksum};
    use visigrid_engine::sheet::{MergedRegion, SheetId};
    use visigrid_engine::workbook::Workbook;

    /// A sheet with every property the ops carry.
    fn put(wb: &mut Workbook, row: usize, col: usize, raw: &str) {
        let key = wb.sheets()[0].id.0;
        let content = if raw.starts_with('=') { CellContent::Formula(raw.into()) } else if raw == "007" { CellContent::Text(raw.into()) } else { CellContent::Value(raw.into()) };
        assert!(apply_ops(wb, &[CollabOp::SetCell { sheet: key, sheet_name: "Sheet1".into(), row, col, content }]).is_empty());
    }

    fn source() -> Workbook {
        let mut wb = Workbook::new();
        for (r, c, raw) in [(0, 0, "Item"), (0, 1, "Amount"), (1, 0, "007"), (1, 1, "12.5"), (2, 1, "=B2*2"), (5, 40, "far")] {
            put(&mut wb, r, c, raw);
        }
        let s = wb.sheet_mut(0).unwrap();
        let mut f = s.get_format(0, 0);
        f.bold = true;
        f.font_color = Some([255, 0, 0, 255]);
        s.set_format(0, 0, f.clone());
        s.set_format(0, 1, f);
        s.layout.col_widths.insert(0, 120.0);
        s.layout.col_widths.insert(1, 120.0);
        s.layout.row_heights.insert(3, 30.0);
        s.layout.hidden_cols.insert(7);
        s.layout.frozen_rows = 1;
        s.add_merge(MergedRegion::new(10, 0, 11, 2)).unwrap();
        wb
    }

    #[test]
    fn a_sheet_rebuilt_from_its_ops_is_the_same_sheet() {
        let src = source();
        let (ops, dropped) = sheet_ops(&src.sheets()[0], 99, "Copy", 1);
        assert!(dropped.is_empty(), "{dropped:?}");
        let mut wb = Workbook::new();
        assert!(apply_ops(&mut wb, &ops).is_empty());
        let copy = &wb.sheets()[1];
        assert_eq!(copy.id, SheetId(99));
        assert_eq!(copy.name, "Copy");
        let orig = &src.sheets()[0];
        for (r, c) in [(0, 0), (0, 1), (1, 0), (1, 1), (2, 1), (5, 40)] {
            assert_eq!(copy.get_raw(r, c), orig.get_raw(r, c), "raw at {r},{c}");
            assert_eq!(copy.get_display(r, c), orig.get_display(r, c), "display at {r},{c}");
            assert_eq!(copy.get_format(r, c), orig.get_format(r, c), "format at {r},{c}");
        }
        // Text stays text: "007" is not the number 7.
        assert_eq!(copy.get_display(1, 0), "007");
        assert_eq!(copy.get_display(2, 1), "25");
        assert_eq!(copy.layout, orig.layout);
        assert_eq!(copy.merged_regions.len(), 1);
        // Equal widths on adjacent columns are one op.
        assert_eq!(ops.iter().filter(|o| matches!(o, CollabOp::SetLines { axis: Axis::Col, .. })).count(), 2);
    }

    #[test]
    fn duplicating_converges_like_any_other_edit() {
        let mut a = source();
        let mut b = source();
        let (ops, _) = sheet_ops(&a.sheets()[0], 7, "Sheet1 (2)", 1);
        apply_ops(&mut a, &ops);
        apply_ops(&mut b, &ops);
        assert_eq!(collab_checksum(&a), collab_checksum(&b));
        assert_eq!(a.sheets().len(), 2);
    }

    #[test]
    fn what_ops_cannot_carry_is_reported() {
        let mut wb = source();
        let s = wb.sheet_mut(0).unwrap();
        s.tab_color = Some([0, 128, 0, 255]);
        s.row_formats.insert(4, CellFormat::default());
        let (_, dropped) = sheet_ops(&wb.sheets()[0], 5, "X", 0);
        assert!(dropped.iter().any(|d| d.contains("tab colour")), "{dropped:?}");
        assert!(dropped.iter().any(|d| d.contains("whole-row")), "{dropped:?}");
    }

    #[test]
    fn long_rows_split_and_gaps_start_new_runs() {
        let mut wb = Workbook::new();
        for c in 0..(RUN_CELLS + 10) {
            put(&mut wb, 0, c, "1");
        }
        put(&mut wb, 1, 0, "a");
        put(&mut wb, 1, 2, "b"); // gap of 1: same run
        put(&mut wb, 1, 50, "c"); // far: new run
        let (ops, _) = sheet_ops(&wb.sheets()[0], 3, "S", 0);
        let runs: Vec<(usize, usize, usize)> = ops
            .iter()
            .filter_map(|o| match o {
                CollabOp::ReplaceRange { row, col, values, .. } => Some((*row, *col, values[0].len())),
                _ => None,
            })
            .collect();
        assert_eq!(runs, vec![(0, 0, RUN_CELLS), (0, RUN_CELLS, 10), (1, 0, 3), (1, 50, 1)]);
    }
}
