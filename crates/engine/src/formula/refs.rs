//! Reference extraction from formula AST.
//!
//! Extracts all cell references from a bound expression as `CellId`s
//! for dependency graph construction.

use rustc_hash::FxHashSet;

use crate::cell_id::CellId;
use crate::named_range::NamedRangeStore;
use crate::sheet::{SheetId, SheetRef};

use super::parser::BoundExpr;

/// Extract all cell references from a bound expression.
///
/// Returns a deduplicated list of `CellId`s that the formula depends on.
///
/// # Arguments
///
/// * `expr` - The bound expression to extract references from
/// * `context_sheet` - The sheet ID where the formula resides (for `SheetRef::Current`)
/// * `named_ranges` - Store for looking up named range targets
/// * `sheet_id_at_idx` - Closure to convert sheet index to SheetId (for named ranges)
///
/// # Known Limitations
///
/// - Dynamic references (INDIRECT, OFFSET) cannot be statically analyzed.
///   These functions are evaluated at runtime, not during extraction.
/// - If a named range points to a deleted sheet or invalid index, it's skipped.
pub fn extract_cell_ids<F>(
    expr: &BoundExpr,
    context_sheet: SheetId,
    named_ranges: &NamedRangeStore,
    sheet_id_at_idx: F,
) -> Vec<CellId>
where
    F: Fn(usize) -> Option<SheetId>,
{
    let mut refs = FxHashSet::default();
    collect_refs(expr, context_sheet, named_ranges, &sheet_id_at_idx, &mut refs);
    refs.into_iter().collect()
}

/// Recursively collect cell references from an expression.
fn collect_refs<F>(
    expr: &BoundExpr,
    context_sheet: SheetId,
    named_ranges: &NamedRangeStore,
    sheet_id_at_idx: &F,
    refs: &mut FxHashSet<CellId>,
) where
    F: Fn(usize) -> Option<SheetId>,
{
    use super::parser::Expr;

    match expr {
        Expr::Number(_) | Expr::Text(_) | Expr::Boolean(_) => {
            // Literals have no dependencies
        }

        Expr::CellRef { sheet, row, col, .. } => {
            if let Some(sheet_id) = resolve_sheet_ref(sheet, context_sheet) {
                refs.insert(CellId::new(sheet_id, *row, *col));
            }
            // If SheetRef::RefError, skip (formula will error anyway)
        }

        Expr::Range {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => {
            if let Some(sheet_id) = resolve_sheet_ref(sheet, context_sheet) {
                // Expand range to individual cells
                for row in *start_row..=*end_row {
                    for col in *start_col..=*end_col {
                        refs.insert(CellId::new(sheet_id, row, col));
                    }
                }
            }
        }

        Expr::NamedRange(name) => {
            if let Some(named_range) = named_ranges.get(name) {
                expand_named_range_target(&named_range.target, sheet_id_at_idx, refs);
            }
            // If named range not found, skip (evaluator handles error)
        }

        Expr::Function { args, .. } => {
            // Recurse into function arguments
            for arg in args {
                collect_refs(arg, context_sheet, named_ranges, sheet_id_at_idx, refs);
            }
        }

        Expr::BinaryOp { left, right, .. } => {
            collect_refs(left, context_sheet, named_ranges, sheet_id_at_idx, refs);
            collect_refs(right, context_sheet, named_ranges, sheet_id_at_idx, refs);
        }
    }
}

/// Resolve a SheetRef to a SheetId.
///
/// Returns None for RefError (deleted sheet).
fn resolve_sheet_ref(sheet_ref: &SheetRef, context_sheet: SheetId) -> Option<SheetId> {
    match sheet_ref {
        SheetRef::Current => Some(context_sheet),
        SheetRef::Id(id) => Some(*id),
        SheetRef::RefError { .. } => None,
    }
}

/// Expand a named range target into CellIds.
fn expand_named_range_target<F>(
    target: &crate::named_range::NamedRangeTarget,
    sheet_id_at_idx: &F,
    refs: &mut FxHashSet<CellId>,
) where
    F: Fn(usize) -> Option<SheetId>,
{
    use crate::named_range::NamedRangeTarget;

    match target {
        NamedRangeTarget::Cell { sheet, row, col } => {
            if let Some(sheet_id) = sheet_id_at_idx(*sheet) {
                refs.insert(CellId::new(sheet_id, *row, *col));
            }
        }
        NamedRangeTarget::Range {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
        } => {
            if let Some(sheet_id) = sheet_id_at_idx(*sheet) {
                for row in *start_row..=*end_row {
                    for col in *start_col..=*end_col {
                        refs.insert(CellId::new(sheet_id, row, col));
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::parser::{parse, Expr};
    use crate::named_range::NamedRange;

    /// Helper to create a bound expression from formula text.
    /// Uses a simple binding that treats all sheet refs as current sheet.
    fn bind_simple(expr: &Expr<crate::sheet::UnboundSheetRef>) -> BoundExpr {
        match expr {
            Expr::Number(n) => Expr::Number(*n),
            Expr::Text(s) => Expr::Text(s.clone()),
            Expr::Boolean(b) => Expr::Boolean(*b),
            Expr::CellRef { sheet: _, row, col, col_abs, row_abs } => {
                Expr::CellRef {
                    sheet: SheetRef::Current,
                    row: *row,
                    col: *col,
                    col_abs: *col_abs,
                    row_abs: *row_abs,
                }
            }
            Expr::Range { sheet: _, start_row, start_col, end_row, end_col,
                          start_col_abs, start_row_abs, end_col_abs, end_row_abs } => {
                Expr::Range {
                    sheet: SheetRef::Current,
                    start_row: *start_row,
                    start_col: *start_col,
                    end_row: *end_row,
                    end_col: *end_col,
                    start_col_abs: *start_col_abs,
                    start_row_abs: *start_row_abs,
                    end_col_abs: *end_col_abs,
                    end_row_abs: *end_row_abs,
                }
            }
            Expr::NamedRange(name) => Expr::NamedRange(name.clone()),
            Expr::Function { name, args } => {
                Expr::Function {
                    name: name.clone(),
                    args: args.iter().map(|a| bind_simple(a)).collect(),
                }
            }
            Expr::BinaryOp { op, left, right } => {
                Expr::BinaryOp {
                    op: *op,
                    left: Box::new(bind_simple(left)),
                    right: Box::new(bind_simple(right)),
                }
            }
        }
    }

    fn sheet(id: u64) -> SheetId {
        SheetId::from_raw(id)
    }

    fn cell(sheet_id: u64, row: usize, col: usize) -> CellId {
        CellId::new(sheet(sheet_id), row, col)
    }

    #[test]
    fn test_same_sheet_ref() {
        // =A1
        let parsed = parse("=A1").unwrap();
        let bound = bind_simple(&parsed);
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 1);
        assert!(refs.contains(&cell(1, 0, 0)));
    }

    #[test]
    fn test_cross_sheet_ref() {
        // =Sheet2!A1 (manually construct bound expr with SheetRef::Id)
        let bound = Expr::CellRef {
            sheet: SheetRef::Id(sheet(2)),
            row: 0,
            col: 0,
            col_abs: false,
            row_abs: false,
        };
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 1);
        assert!(refs.contains(&cell(2, 0, 0)));
    }

    #[test]
    fn test_range_expansion() {
        // =A1:A3
        let parsed = parse("=A1:A3").unwrap();
        let bound = bind_simple(&parsed);
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&cell(1, 0, 0))); // A1
        assert!(refs.contains(&cell(1, 1, 0))); // A2
        assert!(refs.contains(&cell(1, 2, 0))); // A3
    }

    #[test]
    fn test_range_2d_expansion() {
        // =A1:B2
        let parsed = parse("=A1:B2").unwrap();
        let bound = bind_simple(&parsed);
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 4);
        assert!(refs.contains(&cell(1, 0, 0))); // A1
        assert!(refs.contains(&cell(1, 0, 1))); // B1
        assert!(refs.contains(&cell(1, 1, 0))); // A2
        assert!(refs.contains(&cell(1, 1, 1))); // B2
    }

    #[test]
    fn test_named_range_cell() {
        // =Revenue where Revenue = Sheet1!A1
        let bound = Expr::NamedRange("Revenue".to_string());
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("Revenue", 0, 0, 0)).unwrap(); // sheet index 0, A1

        // Map sheet index 0 -> SheetId(1)
        let refs = extract_cell_ids(&bound, sheet(1), &store, |idx| {
            if idx == 0 { Some(sheet(1)) } else { None }
        });

        assert_eq!(refs.len(), 1);
        assert!(refs.contains(&cell(1, 0, 0)));
    }

    #[test]
    fn test_named_range_range() {
        // =Revenue where Revenue = A1:A10
        let bound = Expr::NamedRange("Revenue".to_string());
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::range("Revenue", 0, 0, 0, 9, 0)).unwrap(); // A1:A10

        let refs = extract_cell_ids(&bound, sheet(1), &store, |idx| {
            if idx == 0 { Some(sheet(1)) } else { None }
        });

        assert_eq!(refs.len(), 10);
        for row in 0..10 {
            assert!(refs.contains(&cell(1, row, 0)));
        }
    }

    #[test]
    fn test_mixed_refs() {
        // =Sheet2!A1 + Revenue where Revenue = B1:B5
        let bound = Expr::BinaryOp {
            op: crate::formula::parser::Op::Add,
            left: Box::new(Expr::CellRef {
                sheet: SheetRef::Id(sheet(2)),
                row: 0,
                col: 0,
                col_abs: false,
                row_abs: false,
            }),
            right: Box::new(Expr::NamedRange("Revenue".to_string())),
        };

        let mut store = NamedRangeStore::new();
        store.set(NamedRange::range("Revenue", 0, 0, 1, 4, 1)).unwrap(); // B1:B5 on sheet 0

        let refs = extract_cell_ids(&bound, sheet(1), &store, |idx| {
            if idx == 0 { Some(sheet(1)) } else { None }
        });

        assert_eq!(refs.len(), 6); // Sheet2!A1 + B1:B5 (5 cells)
        assert!(refs.contains(&cell(2, 0, 0))); // Sheet2!A1
        for row in 0..5 {
            assert!(refs.contains(&cell(1, row, 1))); // B1:B5
        }
    }

    #[test]
    fn test_function_args() {
        // =SUM(A1, B1, C1)
        let bound = Expr::Function {
            name: "SUM".to_string(),
            args: vec![
                Expr::CellRef { sheet: SheetRef::Current, row: 0, col: 0, col_abs: false, row_abs: false },
                Expr::CellRef { sheet: SheetRef::Current, row: 0, col: 1, col_abs: false, row_abs: false },
                Expr::CellRef { sheet: SheetRef::Current, row: 0, col: 2, col_abs: false, row_abs: false },
            ],
        };
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 3);
        assert!(refs.contains(&cell(1, 0, 0))); // A1
        assert!(refs.contains(&cell(1, 0, 1))); // B1
        assert!(refs.contains(&cell(1, 0, 2))); // C1
    }

    #[test]
    fn test_duplicate_refs_deduped() {
        // =A1 + A1 + A1
        let bound = Expr::BinaryOp {
            op: crate::formula::parser::Op::Add,
            left: Box::new(Expr::BinaryOp {
                op: crate::formula::parser::Op::Add,
                left: Box::new(Expr::CellRef {
                    sheet: SheetRef::Current, row: 0, col: 0, col_abs: false, row_abs: false
                }),
                right: Box::new(Expr::CellRef {
                    sheet: SheetRef::Current, row: 0, col: 0, col_abs: false, row_abs: false
                }),
            }),
            right: Box::new(Expr::CellRef {
                sheet: SheetRef::Current, row: 0, col: 0, col_abs: false, row_abs: false
            }),
        };
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 1); // Deduped
        assert!(refs.contains(&cell(1, 0, 0)));
    }

    #[test]
    fn test_ref_error_skipped() {
        // Reference to deleted sheet
        let bound = Expr::CellRef {
            sheet: SheetRef::RefError {
                id: sheet(99),
                last_known_name: "DeletedSheet".to_string()
            },
            row: 0,
            col: 0,
            col_abs: false,
            row_abs: false,
        };
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 0); // No refs from deleted sheet
    }

    #[test]
    fn test_unknown_named_range_skipped() {
        // =UnknownName
        let bound = Expr::NamedRange("UnknownName".to_string());
        let store = NamedRangeStore::new(); // Empty store

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 0); // No refs from unknown named range
    }

    #[test]
    fn test_nested_function() {
        // =SUM(A1:A3, MAX(B1:B3))
        let bound = Expr::Function {
            name: "SUM".to_string(),
            args: vec![
                Expr::Range {
                    sheet: SheetRef::Current,
                    start_row: 0, start_col: 0,
                    end_row: 2, end_col: 0,
                    start_col_abs: false, start_row_abs: false,
                    end_col_abs: false, end_row_abs: false,
                },
                Expr::Function {
                    name: "MAX".to_string(),
                    args: vec![
                        Expr::Range {
                            sheet: SheetRef::Current,
                            start_row: 0, start_col: 1,
                            end_row: 2, end_col: 1,
                            start_col_abs: false, start_row_abs: false,
                            end_col_abs: false, end_row_abs: false,
                        },
                    ],
                },
            ],
        };
        let store = NamedRangeStore::new();

        let refs = extract_cell_ids(&bound, sheet(1), &store, |_| None);

        assert_eq!(refs.len(), 6); // A1:A3 (3) + B1:B3 (3)
    }
}
