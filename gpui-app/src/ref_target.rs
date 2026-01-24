//! Pure functions for ref target border logic.
//!
//! This module contains the border calculation logic extracted from Spreadsheet
//! so it can be unit tested without any app scaffolding.

/// A normalized rectangle (min/max already computed).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rect {
    pub min_row: usize,
    pub max_row: usize,
    pub min_col: usize,
    pub max_col: usize,
}

/// Border edges to draw.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Edges {
    pub top: bool,
    pub right: bool,
    pub bottom: bool,
    pub left: bool,
}

impl Edges {
    pub const NONE: Edges = Edges {
        top: false,
        right: false,
        bottom: false,
        left: false,
    };

    pub const ALL: Edges = Edges {
        top: true,
        right: true,
        bottom: true,
        left: true,
    };
}

/// Normalize two corner points into a Rect with min/max computed.
/// Handles reversed ranges (where end < start).
pub fn normalize_rect(
    (r1, c1): (usize, usize),
    (r2, c2): (usize, usize),
) -> Rect {
    Rect {
        min_row: r1.min(r2),
        max_row: r1.max(r2),
        min_col: c1.min(c2),
        max_col: c1.max(c2),
    }
}

/// Check if a cell is contained within the rectangle.
pub fn contains(rect: &Rect, row: usize, col: usize) -> bool {
    row >= rect.min_row
        && row <= rect.max_row
        && col >= rect.min_col
        && col <= rect.max_col
}

/// Compute which border edges to draw for a cell within the rectangle.
/// Returns Edges::NONE if the cell is outside the rectangle.
/// Only outer edges of the rectangle are marked true.
pub fn borders(rect: &Rect, row: usize, col: usize) -> Edges {
    if !contains(rect, row, col) {
        return Edges::NONE;
    }

    Edges {
        top: row == rect.min_row,
        bottom: row == rect.max_row,
        left: col == rect.min_col,
        right: col == rect.max_col,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // =========================================================================
    // Test 1: Single cell (2,2)-(2,2) - all four edges true only on that cell
    // =========================================================================
    #[test]
    fn test_single_cell_all_edges() {
        let rect = normalize_rect((2, 2), (2, 2));

        // The single cell should have all edges
        assert_eq!(borders(&rect, 2, 2), Edges::ALL);

        // Adjacent cells should have no edges (they're outside)
        assert_eq!(borders(&rect, 1, 2), Edges::NONE);
        assert_eq!(borders(&rect, 3, 2), Edges::NONE);
        assert_eq!(borders(&rect, 2, 1), Edges::NONE);
        assert_eq!(borders(&rect, 2, 3), Edges::NONE);
    }

    // =========================================================================
    // Test 2: Horizontal range (2,2)-(2,5) - top/bottom true across,
    //         left only at c2, right only at c5
    // =========================================================================
    #[test]
    fn test_horizontal_range() {
        let rect = normalize_rect((2, 2), (2, 5));

        // Left edge cell: top, bottom, left (no right)
        assert_eq!(
            borders(&rect, 2, 2),
            Edges { top: true, right: false, bottom: true, left: true }
        );

        // Middle cells: top and bottom only
        assert_eq!(
            borders(&rect, 2, 3),
            Edges { top: true, right: false, bottom: true, left: false }
        );
        assert_eq!(
            borders(&rect, 2, 4),
            Edges { top: true, right: false, bottom: true, left: false }
        );

        // Right edge cell: top, bottom, right (no left)
        assert_eq!(
            borders(&rect, 2, 5),
            Edges { top: true, right: true, bottom: true, left: false }
        );
    }

    // =========================================================================
    // Test 3: Vertical range (2,2)-(5,2) - left/right true across,
    //         top only at r2, bottom only at r5
    // =========================================================================
    #[test]
    fn test_vertical_range() {
        let rect = normalize_rect((2, 2), (5, 2));

        // Top edge cell: top, left, right (no bottom)
        assert_eq!(
            borders(&rect, 2, 2),
            Edges { top: true, right: true, bottom: false, left: true }
        );

        // Middle cells: left and right only
        assert_eq!(
            borders(&rect, 3, 2),
            Edges { top: false, right: true, bottom: false, left: true }
        );
        assert_eq!(
            borders(&rect, 4, 2),
            Edges { top: false, right: true, bottom: false, left: true }
        );

        // Bottom edge cell: bottom, left, right (no top)
        assert_eq!(
            borders(&rect, 5, 2),
            Edges { top: false, right: true, bottom: true, left: true }
        );
    }

    // =========================================================================
    // Test 4: 2D rectangle (2,2)-(5,5) - outer edges only
    // =========================================================================
    #[test]
    fn test_2d_rectangle_outer_edges() {
        let rect = normalize_rect((2, 2), (5, 5));

        // Top-left corner: top and left
        assert_eq!(
            borders(&rect, 2, 2),
            Edges { top: true, right: false, bottom: false, left: true }
        );

        // Top-right corner: top and right
        assert_eq!(
            borders(&rect, 2, 5),
            Edges { top: true, right: true, bottom: false, left: false }
        );

        // Bottom-left corner: bottom and left
        assert_eq!(
            borders(&rect, 5, 2),
            Edges { top: false, right: false, bottom: true, left: true }
        );

        // Bottom-right corner: bottom and right
        assert_eq!(
            borders(&rect, 5, 5),
            Edges { top: false, right: true, bottom: true, left: false }
        );

        // Top edge (middle): top only
        assert_eq!(
            borders(&rect, 2, 3),
            Edges { top: true, right: false, bottom: false, left: false }
        );

        // Bottom edge (middle): bottom only
        assert_eq!(
            borders(&rect, 5, 3),
            Edges { top: false, right: false, bottom: true, left: false }
        );

        // Left edge (middle): left only
        assert_eq!(
            borders(&rect, 3, 2),
            Edges { top: false, right: false, bottom: false, left: true }
        );

        // Right edge (middle): right only
        assert_eq!(
            borders(&rect, 3, 5),
            Edges { top: false, right: true, bottom: false, left: false }
        );
    }

    // =========================================================================
    // Test 5: Reversed range (5,5)-(2,2) - identical to test 4
    // =========================================================================
    #[test]
    fn test_reversed_range_identical_to_normal() {
        let normal = normalize_rect((2, 2), (5, 5));
        let reversed = normalize_rect((5, 5), (2, 2));

        // Normalization should produce identical rects
        assert_eq!(normal, reversed);

        // All border calculations should be identical
        for row in 2..=5 {
            for col in 2..=5 {
                assert_eq!(
                    borders(&normal, row, col),
                    borders(&reversed, row, col),
                    "Mismatch at ({}, {})",
                    row,
                    col
                );
            }
        }
    }

    // =========================================================================
    // Test 6: Inside cell (3,3) for 2D rect - all edges false
    // =========================================================================
    #[test]
    fn test_inside_cell_no_edges() {
        let rect = normalize_rect((2, 2), (5, 5));

        // Interior cells should have no edges
        assert_eq!(borders(&rect, 3, 3), Edges::NONE);
        assert_eq!(borders(&rect, 3, 4), Edges::NONE);
        assert_eq!(borders(&rect, 4, 3), Edges::NONE);
        assert_eq!(borders(&rect, 4, 4), Edges::NONE);
    }

    // =========================================================================
    // Additional: contains() correctness
    // =========================================================================
    #[test]
    fn test_contains() {
        let rect = normalize_rect((2, 2), (5, 5));

        // Inside
        assert!(contains(&rect, 2, 2));
        assert!(contains(&rect, 3, 3));
        assert!(contains(&rect, 5, 5));

        // Outside
        assert!(!contains(&rect, 1, 3));
        assert!(!contains(&rect, 6, 3));
        assert!(!contains(&rect, 3, 1));
        assert!(!contains(&rect, 3, 6));
    }
}
