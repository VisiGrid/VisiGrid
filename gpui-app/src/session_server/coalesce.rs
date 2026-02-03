//! Cell range coalescing for efficient event broadcasting.
//!
//! Converts a set of changed cells into a minimal(ish) set of rectangular ranges.
//! This reduces event payload size and client redraw churn.
//!
//! Algorithm:
//! 1. Group cells by sheet, then by row
//! 2. For each row, collapse columns into contiguous runs
//! 3. Merge runs vertically when they have identical column bounds
//! 4. Apply caps to prevent pathological cases (e.g., checkerboard patterns)
//!
//! Guarantees:
//! - Coverage: output ranges cover all input cells (superset, not exact)
//! - Determinism: same input always produces same output
//! - Bounded: output size is capped regardless of input pattern

use std::collections::HashMap;
use super::protocol::{CellRange, CellRef};

/// Maximum number of ranges per event before falling back to bounding box.
const MAX_RANGES_PER_SHEET: usize = 2000;

/// Coalesce a list of changed cells into rectangular ranges.
///
/// The output ranges are guaranteed to cover all input cells, but may
/// include additional cells (superset guarantee). Output is deterministic
/// for a given input set.
///
/// # Performance
/// - Time: O(n log n) where n = number of cells (dominated by sorting)
/// - Space: O(n) for intermediate structures
/// - Output size: bounded by MAX_RANGES_PER_SHEET per sheet
pub fn coalesce_cells_to_ranges(cells: &[CellRef]) -> Vec<CellRange> {
    if cells.is_empty() {
        return Vec::new();
    }

    // Group by sheet
    let mut by_sheet: HashMap<usize, Vec<(usize, usize)>> = HashMap::new();
    for cell in cells {
        by_sheet
            .entry(cell.sheet)
            .or_default()
            .push((cell.row, cell.col));
    }

    let mut result = Vec::new();

    for (sheet, mut coords) in by_sheet {
        let ranges = coalesce_sheet_cells(sheet, &mut coords);
        result.extend(ranges);
    }

    // Sort for deterministic output
    result.sort_by_key(|r| (r.sheet, r.r1, r.c1, r.r2, r.c2));
    result
}

/// Coalesce cells within a single sheet.
fn coalesce_sheet_cells(sheet: usize, coords: &mut Vec<(usize, usize)>) -> Vec<CellRange> {
    if coords.is_empty() {
        return Vec::new();
    }

    // Sort by (row, col) for deterministic processing
    coords.sort_unstable();

    // Deduplicate (after sorting, duplicates are adjacent)
    coords.dedup();

    if coords.is_empty() {
        return Vec::new();
    }

    // Step 1: Group by row and collapse into horizontal runs
    let mut row_runs: HashMap<usize, Vec<(usize, usize)>> = HashMap::new(); // row -> [(start_col, end_col)]

    let mut current_row = coords[0].0;
    let mut run_start = coords[0].1;
    let mut run_end = coords[0].1;

    for &(row, col) in &coords[1..] {
        if row == current_row && col == run_end + 1 {
            // Extend current run
            run_end = col;
        } else {
            // Save current run
            row_runs
                .entry(current_row)
                .or_default()
                .push((run_start, run_end));

            // Start new run
            current_row = row;
            run_start = col;
            run_end = col;
        }
    }
    // Save final run
    row_runs
        .entry(current_row)
        .or_default()
        .push((run_start, run_end));

    // Step 2: Merge runs vertically
    // Key: (start_col, end_col) -> Vec<(start_row, end_row)>
    let mut vertical_runs: HashMap<(usize, usize), Vec<(usize, usize)>> = HashMap::new();

    // Process rows in sorted order for determinism
    let mut rows: Vec<_> = row_runs.keys().copied().collect();
    rows.sort_unstable();

    for row in rows {
        let runs = row_runs.get(&row).unwrap();
        for &(c1, c2) in runs {
            let key = (c1, c2);
            let entry = vertical_runs.entry(key).or_default();

            // Try to extend an existing vertical run
            if let Some(last) = entry.last_mut() {
                if last.1 + 1 == row {
                    // Extend downward
                    last.1 = row;
                    continue;
                }
            }
            // Start new vertical run
            entry.push((row, row));
        }
    }

    // Step 3: Convert to CellRange
    let mut ranges: Vec<CellRange> = Vec::new();
    for ((c1, c2), row_spans) in vertical_runs {
        for (r1, r2) in row_spans {
            ranges.push(CellRange::new(sheet, r1, c1, r2, c2));
        }
    }

    // Step 4: Apply cap - fall back to bounding box if too many ranges
    if ranges.len() > MAX_RANGES_PER_SHEET {
        // Compute bounding box
        let r1 = ranges.iter().map(|r| r.r1).min().unwrap();
        let r2 = ranges.iter().map(|r| r.r2).max().unwrap();
        let c1 = ranges.iter().map(|r| r.c1).min().unwrap();
        let c2 = ranges.iter().map(|r| r.c2).max().unwrap();
        return vec![CellRange::new(sheet, r1, c1, r2, c2)];
    }

    ranges
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(sheet: usize, row: usize, col: usize) -> CellRef {
        CellRef { sheet, row, col }
    }

    fn range(sheet: usize, r1: usize, c1: usize, r2: usize, c2: usize) -> CellRange {
        CellRange::new(sheet, r1, c1, r2, c2)
    }

    #[test]
    fn test_empty_input() {
        let result = coalesce_cells_to_ranges(&[]);
        assert!(result.is_empty());
    }

    #[test]
    fn test_single_cell() {
        let cells = vec![cell(0, 5, 3)];
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result, vec![range(0, 5, 3, 5, 3)]);
    }

    #[test]
    fn test_horizontal_run() {
        // A1:D1 (row 0, cols 0-3)
        let cells = vec![
            cell(0, 0, 0),
            cell(0, 0, 1),
            cell(0, 0, 2),
            cell(0, 0, 3),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result, vec![range(0, 0, 0, 0, 3)]);
    }

    #[test]
    fn test_vertical_run() {
        // A1:A4 (rows 0-3, col 0)
        let cells = vec![
            cell(0, 0, 0),
            cell(0, 1, 0),
            cell(0, 2, 0),
            cell(0, 3, 0),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result, vec![range(0, 0, 0, 3, 0)]);
    }

    #[test]
    fn test_rectangle_fill() {
        // A1:C3 (3x3 block)
        let mut cells = Vec::new();
        for r in 0..3 {
            for c in 0..3 {
                cells.push(cell(0, r, c));
            }
        }
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result, vec![range(0, 0, 0, 2, 2)]);
    }

    #[test]
    fn test_two_disjoint_blocks() {
        // Block 1: A1:B2, Block 2: D4:E5
        let cells = vec![
            cell(0, 0, 0), cell(0, 0, 1),
            cell(0, 1, 0), cell(0, 1, 1),
            cell(0, 3, 3), cell(0, 3, 4),
            cell(0, 4, 3), cell(0, 4, 4),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result.len(), 2);
        assert!(result.contains(&range(0, 0, 0, 1, 1)));
        assert!(result.contains(&range(0, 3, 3, 4, 4)));
    }

    #[test]
    fn test_staircase_pattern() {
        // Staircase: (0,0), (1,1), (2,2)
        let cells = vec![
            cell(0, 0, 0),
            cell(0, 1, 1),
            cell(0, 2, 2),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        // Each cell becomes its own range (no merging possible)
        assert_eq!(result.len(), 3);
        assert!(result.contains(&range(0, 0, 0, 0, 0)));
        assert!(result.contains(&range(0, 1, 1, 1, 1)));
        assert!(result.contains(&range(0, 2, 2, 2, 2)));
    }

    #[test]
    fn test_deterministic_output() {
        // Same input in different order should produce same output
        let cells1 = vec![
            cell(0, 2, 0), cell(0, 0, 0), cell(0, 1, 0),
        ];
        let cells2 = vec![
            cell(0, 0, 0), cell(0, 1, 0), cell(0, 2, 0),
        ];
        let result1 = coalesce_cells_to_ranges(&cells1);
        let result2 = coalesce_cells_to_ranges(&cells2);
        assert_eq!(result1, result2);
    }

    #[test]
    fn test_multiple_sheets() {
        let cells = vec![
            cell(0, 0, 0), cell(0, 0, 1),
            cell(1, 5, 5), cell(1, 5, 6),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        assert_eq!(result.len(), 2);
        assert!(result.contains(&range(0, 0, 0, 0, 1)));
        assert!(result.contains(&range(1, 5, 5, 5, 6)));
    }

    #[test]
    fn test_l_shape_pattern() {
        // L-shape: vertical arm A1:A3, horizontal arm A3:C3
        let cells = vec![
            cell(0, 0, 0),
            cell(0, 1, 0),
            cell(0, 2, 0), cell(0, 2, 1), cell(0, 2, 2),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        // Should produce 2 ranges: vertical part + horizontal extension
        // The exact split depends on algorithm, but coverage is guaranteed
        let total_cells: usize = result.iter().map(|r| r.cell_count()).sum();
        assert!(total_cells >= 5); // Coverage guarantee
        assert!(result.len() <= 3); // Reasonable efficiency
    }

    #[test]
    fn test_cap_fallback_to_bounding_box() {
        // Create a checkerboard pattern that would produce many ranges
        let mut cells = Vec::new();
        for r in 0..100 {
            for c in 0..100 {
                if (r + c) % 2 == 0 {
                    cells.push(cell(0, r, c));
                }
            }
        }
        let result = coalesce_cells_to_ranges(&cells);
        // Should fall back to bounding box due to cap
        assert_eq!(result.len(), 1);
        assert_eq!(result[0], range(0, 0, 0, 99, 99));
    }

    #[test]
    fn test_duplicate_cells_handled() {
        // Duplicates should not cause issues
        let cells = vec![
            cell(0, 0, 0),
            cell(0, 0, 0), // duplicate
            cell(0, 0, 1),
        ];
        let result = coalesce_cells_to_ranges(&cells);
        // Should still produce a single range
        assert_eq!(result, vec![range(0, 0, 0, 0, 1)]);
    }
}
