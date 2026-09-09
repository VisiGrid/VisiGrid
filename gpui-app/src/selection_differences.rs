//! Excel's "Row differences" and "Column differences" (Go To Special).
//!
//! Row differences (Ctrl+\): within the selected region, every cell whose
//! value differs from the cell in the active cell's column on the same row.
//! Column differences (Ctrl+Shift+|): every cell whose value differs from the
//! cell in the active cell's row in the same column.
//!
//! Comparison is by computed value, not by formula text: the question a user
//! is asking is "which rows disagree", and `=A1+1` and `=B1+1` that evaluate
//! to the same number agree. Excel compares formulas by relative structure
//! instead, which answers a different question; this is the one the
//! spreadsheet's own engine can answer exactly.
//!
//! The helpers are pure over an engine `Sheet` so they can be tested without a
//! window, matching `find_current_region`.

use visigrid_engine::sheet::Sheet;

/// Cells in `[(min_row, min_col), (max_row, max_col)]` that differ from the
/// cell in `pivot_col` on their own row. The pivot column's own cells are never
/// reported. Empty when the region is a single column, since there is nothing
/// to compare against.
pub fn find_row_differences(
    sheet: &Sheet,
    (min_row, min_col): (usize, usize),
    (max_row, max_col): (usize, usize),
    pivot_col: usize,
) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    if min_col == max_col || pivot_col < min_col || pivot_col > max_col {
        return out;
    }
    for row in min_row..=max_row {
        let pivot = sheet.get_computed_value(row, pivot_col);
        for col in min_col..=max_col {
            if col == pivot_col {
                continue;
            }
            if sheet.get_computed_value(row, col) != pivot {
                out.push((row, col));
            }
        }
    }
    out
}

/// Cells in the region that differ from the cell in `pivot_row` in their own
/// column. Mirror of [`find_row_differences`].
pub fn find_column_differences(
    sheet: &Sheet,
    (min_row, min_col): (usize, usize),
    (max_row, max_col): (usize, usize),
    pivot_row: usize,
) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    if min_row == max_row || pivot_row < min_row || pivot_row > max_row {
        return out;
    }
    for col in min_col..=max_col {
        let pivot = sheet.get_computed_value(pivot_row, col);
        for row in min_row..=max_row {
            if row == pivot_row {
                continue;
            }
            if sheet.get_computed_value(row, col) != pivot {
                out.push((row, col));
            }
        }
    }
    out
}
