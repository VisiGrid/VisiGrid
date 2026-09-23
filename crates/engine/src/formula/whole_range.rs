//! Whole-row/column references retain their unbounded axis until evaluation.

use super::eval::CellLookup;
use super::parser::{BoundExpr, Expr, RangeAxis};
use crate::cell_id::CellId;
use crate::sheet::{SheetId, SheetRef};

/// Compact dependency subscription. Only existing cells need concrete graph
/// edges; this subscription also catches writes anywhere along the open axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WholeRangeRef {
    pub sheet: SheetId,
    pub axis: RangeAxis,
    pub start: usize,
    pub end: usize,
}

impl WholeRangeRef {
    pub fn contains(&self, cell: CellId) -> bool {
        let coord = match self.axis {
            RangeAxis::Row => cell.row as usize,
            RangeAxis::Column => cell.col as usize,
        };
        self.sheet == cell.sheet && self.start <= coord && coord <= self.end
    }
}

pub fn extract_whole_ranges(expr: &BoundExpr, current_sheet: SheetId) -> Vec<WholeRangeRef> {
    fn collect(expr: &BoundExpr, current_sheet: SheetId, ranges: &mut Vec<WholeRangeRef>) {
        match expr {
            Expr::WholeRange {
                sheet,
                axis,
                start,
                end,
                ..
            } => {
                let sheet = match sheet {
                    SheetRef::Current => current_sheet,
                    SheetRef::Id(id) => *id,
                    SheetRef::RefError { .. } => return,
                };
                ranges.push(WholeRangeRef {
                    sheet,
                    axis: *axis,
                    start: (*start).min(*end),
                    end: (*start).max(*end),
                });
            }
            Expr::Function { args, .. } => {
                for arg in args {
                    collect(arg, current_sheet, ranges);
                }
            }
            Expr::BinaryOp { left, right, .. } => {
                collect(left, current_sheet, ranges);
                collect(right, current_sheet, ranges);
            }
            _ => {}
        }
    }
    let mut ranges = Vec::new();
    collect(expr, current_sheet, &mut ranges);
    ranges
}

/// Resolve just this reference, leaving ordinary expressions untouched.
/// The finite axis is preserved even on an empty sheet. At least one blank
/// row/column is supplied for the existing rectangular range evaluators.
pub fn bound_for_evaluation<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> BoundExpr {
    match expr {
        Expr::WholeRange {
            sheet,
            axis,
            start,
            end,
            start_abs,
            end_abs,
        } => {
            let (rows, cols) = lookup.data_bounds(sheet);
            let (start, end) = ((*start).min(*end), (*start).max(*end));
            let (start_col, start_row, end_col, end_row) = match axis {
                RangeAxis::Column => (
                    start,
                    lookup.whole_column_start(),
                    end,
                    rows.saturating_sub(1).max(lookup.whole_column_start()),
                ),
                RangeAxis::Row => (0, start, cols.saturating_sub(1), end),
            };
            Expr::Range {
                sheet: sheet.clone(),
                start_col,
                start_row,
                end_col,
                end_row,
                start_col_abs: *axis == RangeAxis::Row || *start_abs,
                end_col_abs: *axis == RangeAxis::Row || *end_abs,
                start_row_abs: *axis == RangeAxis::Column || *start_abs,
                end_row_abs: *axis == RangeAxis::Column || *end_abs,
            }
        }
        _ => expr.clone(),
    }
}
