//! Formula reference adjustment for structural edits (row/column insert and
//! delete).
//!
//! When rows or columns are inserted or deleted, cells move — and every
//! formula that referenced the moved cells must be rewritten so it still
//! points at the same *content*. Excel and Sheets do this; VisiGrid did not
//! until 2026-07-30, which silently broke formulas on a top-five operation.
//!
//! Three semantics that are easy to get wrong, all pinned by tests below:
//!
//! 1. **Absolute references shift too.** `$` pins a reference against
//!    copy/fill, not against structural edits: inserting a row above `$A$1`
//!    gives `$A$2`. Anything else makes the reference point at different
//!    data.
//! 2. **`#REF!` is a text rewrite, not an evaluation-time error.** Deleting
//!    the target of `=A1+B1` stores `=#REF!+B1` permanently; re-inserting
//!    the row does not bring it back. Modelling it in the AST is what keeps
//!    undo, save, and visigrid-json round-trips agreeing.
//! 3. **Ranges use grid-line semantics, matching merged regions.** An insert
//!    at or before a range's first line shifts the whole range; an insert
//!    strictly inside it expands the range. (`at <= start` shifts,
//!    `at <= end` expands — the exact rule `Sheet::insert_rows` already
//!    applies to merges.)

use crate::formula::parser::{parse, Expr, ParsedExpr};
use crate::sheet::UnboundSheetRef;

/// Which axis a structural edit operates on.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    Row,
    Col,
}

/// A structural edit, as seen by reference adjustment.
#[derive(Debug, Clone)]
pub struct StructuralEdit {
    /// Name of the sheet being edited. References resolve against this:
    /// unqualified refs in formulas on this sheet, and `Sheet!`-qualified
    /// refs from anywhere.
    pub sheet_name: String,
    pub axis: Axis,
    /// First index inserted, or first index deleted.
    pub at: usize,
    pub count: usize,
    /// true = delete, false = insert.
    pub delete: bool,
}

/// How a single coordinate is affected.
enum Coord {
    Same(usize),
    Moved(usize),
    /// The coordinate's target was deleted.
    Dead,
}

fn adjust_coord(v: usize, edit: &StructuralEdit) -> Coord {
    let (at, n) = (edit.at, edit.count);
    if edit.delete {
        if v >= at + n {
            Coord::Moved(v - n)
        } else if v >= at {
            Coord::Dead
        } else {
            Coord::Same(v)
        }
    } else if v >= at {
        Coord::Moved(v + n)
    } else {
        Coord::Same(v)
    }
}

/// Adjust a range's [start, end] span. Returns None if the whole span died.
fn adjust_span(start: usize, end: usize, edit: &StructuralEdit) -> Option<(usize, usize)> {
    let (at, n) = (edit.at, edit.count);
    if edit.delete {
        let del_end = at + n; // exclusive
        if start >= at && end < del_end {
            return None; // entirely inside the deleted span
        }
        let new_start = if start >= del_end {
            start - n
        } else if start >= at {
            at
        } else {
            start
        };
        let new_end = if end >= del_end {
            end - n
        } else if end >= at {
            // End fell inside the deleted span: clamp to just before it.
            if at == 0 {
                return None;
            }
            at - 1
        } else {
            end
        };
        if new_end < new_start {
            return None;
        }
        Some((new_start, new_end))
    } else {
        // Grid-line semantics, mirroring merged regions.
        if at <= start {
            Some((start + n, end + n)) // shift the whole range
        } else if at <= end {
            Some((start, end + n)) // insert strictly inside → expand
        } else {
            Some((start, end))
        }
    }
}

/// Does this reference target the sheet being edited?
fn targets_edited_sheet(sheet: &UnboundSheetRef, edit: &StructuralEdit, formula_sheet: &str) -> bool {
    match sheet {
        UnboundSheetRef::Current => formula_sheet.eq_ignore_ascii_case(&edit.sheet_name),
        UnboundSheetRef::Named(name) => name.eq_ignore_ascii_case(&edit.sheet_name),
    }
}

/// Rewrite an expression's references for a structural edit.
/// `formula_sheet` is the name of the sheet the formula lives on, which
/// determines whether its unqualified references point at the edited sheet.
pub fn adjust_expr(expr: &ParsedExpr, edit: &StructuralEdit, formula_sheet: &str) -> ParsedExpr {
    match expr {
        Expr::CellRef { sheet, col, row, col_abs, row_abs } => {
            if !targets_edited_sheet(sheet, edit, formula_sheet) {
                return expr.clone();
            }
            let v = if edit.axis == Axis::Row { *row } else { *col };
            match adjust_coord(v, edit) {
                Coord::Same(_) => expr.clone(),
                Coord::Dead => Expr::RefError,
                Coord::Moved(nv) => {
                    let (row, col) = if edit.axis == Axis::Row { (nv, *col) } else { (*row, nv) };
                    Expr::CellRef {
                        sheet: sheet.clone(),
                        col,
                        row,
                        col_abs: *col_abs,
                        row_abs: *row_abs,
                    }
                }
            }
        }
        Expr::Range {
            sheet,
            start_col, start_row, end_col, end_row,
            start_col_abs, start_row_abs, end_col_abs, end_row_abs,
        } => {
            if !targets_edited_sheet(sheet, edit, formula_sheet) {
                return expr.clone();
            }
            let (s, e) = if edit.axis == Axis::Row {
                (*start_row, *end_row)
            } else {
                (*start_col, *end_col)
            };
            match adjust_span(s.min(e), s.max(e), edit) {
                None => Expr::RefError,
                Some((ns, ne)) => {
                    let (start_row, end_row, start_col, end_col) = if edit.axis == Axis::Row {
                        (ns, ne, *start_col, *end_col)
                    } else {
                        (*start_row, *end_row, ns, ne)
                    };
                    Expr::Range {
                        sheet: sheet.clone(),
                        start_col, start_row, end_col, end_row,
                        start_col_abs: *start_col_abs,
                        start_row_abs: *start_row_abs,
                        end_col_abs: *end_col_abs,
                        end_row_abs: *end_row_abs,
                    }
                }
            }
        }
        Expr::Function { name, args } => Expr::Function {
            name: name.clone(),
            args: args.iter().map(|a| adjust_expr(a, edit, formula_sheet)).collect(),
        },
        Expr::BinaryOp { op, left, right } => Expr::BinaryOp {
            op: *op,
            left: Box::new(adjust_expr(left, edit, formula_sheet)),
            right: Box::new(adjust_expr(right, edit, formula_sheet)),
        },
        other => other.clone(),
    }
}

/// Adjust a formula's text for a structural edit.
///
/// Returns `Some(new_text)` when the formula changed, `None` when it did not
/// (including when it is not a formula, or cannot be parsed — unparseable
/// text is left exactly as the user typed it rather than being mangled).
pub fn adjust_formula_text(
    raw: &str,
    edit: &StructuralEdit,
    formula_sheet: &str,
) -> Option<String> {
    if !raw.starts_with('=') {
        return None;
    }
    let parsed = parse(raw).ok()?;
    let adjusted = adjust_expr(&parsed, edit, formula_sheet);
    let text = crate::formula::parser::format_parsed_expr(&adjusted);
    if text == raw {
        None
    } else {
        Some(text)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn edit(axis: Axis, at: usize, count: usize, delete: bool) -> StructuralEdit {
        StructuralEdit { sheet_name: "Sheet1".into(), axis, at, count, delete }
    }

    fn adjust(raw: &str, e: &StructuralEdit) -> String {
        adjust_formula_text(raw, e, "Sheet1").unwrap_or_else(|| raw.to_string())
    }

    #[test]
    fn insert_shifts_refs_at_or_after() {
        let e = edit(Axis::Row, 0, 1, false); // insert a row at the top
        assert_eq!(adjust("=A1*2", &e), "=A2*2");
        assert_eq!(adjust("=A5+B5", &e), "=A6+B6");
        // Before the insertion point: untouched.
        let e2 = edit(Axis::Row, 10, 3, false);
        assert_eq!(adjust("=A1*2", &e2), "=A1*2");
        assert_eq!(adjust("=A11", &e2), "=A14");
    }

    #[test]
    fn absolute_refs_follow_content_on_structural_edits() {
        // $ pins against copy/fill — NOT against inserts.
        let e = edit(Axis::Row, 0, 1, false);
        assert_eq!(adjust("=$A$1", &e), "=$A$2");
        assert_eq!(adjust("=A$1", &e), "=A$2");
        let e2 = edit(Axis::Col, 0, 1, false);
        assert_eq!(adjust("=$A$1", &e2), "=$B$1");
    }

    #[test]
    fn columns_shift_independently_of_rows() {
        let e = edit(Axis::Col, 1, 2, false); // insert 2 cols at B
        assert_eq!(adjust("=A1", &e), "=A1");
        assert_eq!(adjust("=B1", &e), "=D1");
        assert_eq!(adjust("=SUM(B1:C1)", &e), "=SUM(D1:E1)");
    }

    #[test]
    fn ranges_use_grid_line_semantics() {
        // Insert AT the range's first row shifts the whole range...
        let at_start = edit(Axis::Row, 1, 1, false);
        assert_eq!(adjust("=SUM(A2:A10)", &at_start), "=SUM(A3:A11)");
        // ...strictly inside expands it.
        let inside = edit(Axis::Row, 4, 1, false);
        assert_eq!(adjust("=SUM(A2:A10)", &inside), "=SUM(A2:A11)");
        // Entirely after: untouched.
        let after = edit(Axis::Row, 20, 1, false);
        assert_eq!(adjust("=SUM(A2:A10)", &after), "=SUM(A2:A10)");
    }

    #[test]
    fn delete_shifts_back_and_kills_dead_refs() {
        let e = edit(Axis::Row, 0, 1, true); // delete row 1
        assert_eq!(adjust("=A2", &e), "=A1");
        // The deleted cell itself becomes #REF! in the TEXT.
        assert_eq!(adjust("=A1+B1", &e), "=#REF!+#REF!");
        // Mixed: one side dies, the other shifts.
        let e2 = edit(Axis::Row, 4, 1, true); // delete row 5
        assert_eq!(adjust("=A5+A6", &e2), "=#REF!+A5");
    }

    #[test]
    fn delete_clamps_and_consumes_ranges() {
        // Delete rows 4-5 out of the middle of A2:A10.
        let mid = edit(Axis::Row, 3, 2, true);
        assert_eq!(adjust("=SUM(A2:A10)", &mid), "=SUM(A2:A8)");
        // Overlapping the start: the range starts where the deletion did.
        let head = edit(Axis::Row, 0, 3, true);
        assert_eq!(adjust("=SUM(A2:A10)", &head), "=SUM(A1:A7)");
        // Swallowing the range entirely: #REF!.
        let all = edit(Axis::Row, 0, 20, true);
        assert_eq!(adjust("=SUM(A2:A10)", &all), "=SUM(#REF!)");
    }

    #[test]
    fn only_the_edited_sheet_is_rewritten() {
        let e = edit(Axis::Row, 0, 1, false);
        // A formula on another sheet, referencing it unqualified: untouched.
        assert_eq!(
            adjust_formula_text("=A1", &e, "Sheet2"),
            None,
            "unqualified ref on a different sheet must not move"
        );
        // ...but a qualified ref to the edited sheet moves from anywhere.
        assert_eq!(
            adjust_formula_text("=Sheet1!A1", &e, "Sheet2").as_deref(),
            Some("=Sheet1!A2")
        );
        // Refs to a third sheet are untouched even from the edited sheet.
        assert_eq!(adjust_formula_text("=Sheet3!A1", &e, "Sheet1"), None);
    }

    #[test]
    fn ref_error_round_trips_through_text() {
        // Once dead, a reference stays dead — re-inserting does not revive it.
        let del = edit(Axis::Row, 0, 1, true);
        let dead = adjust("=A1+1", &del);
        assert_eq!(dead, "=#REF!+1");
        let ins = edit(Axis::Row, 0, 1, false);
        assert_eq!(adjust(&dead, &ins), "=#REF!+1");
    }

    #[test]
    fn non_formulas_and_junk_are_left_alone() {
        let e = edit(Axis::Row, 0, 1, false);
        assert_eq!(adjust_formula_text("hello", &e, "Sheet1"), None);
        assert_eq!(adjust_formula_text("42", &e, "Sheet1"), None);
        // Unparseable formula text is preserved verbatim, never mangled.
        assert_eq!(adjust_formula_text("=SUM(((", &e, "Sheet1"), None);
    }
}
