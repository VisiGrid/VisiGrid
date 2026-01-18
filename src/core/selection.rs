/// A rectangular range of cells, inclusive on both ends.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Range {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

impl Range {
    /// Create a new range, automatically normalizing so start <= end.
    pub fn new(r1: usize, c1: usize, r2: usize, c2: usize) -> Self {
        Self {
            start_row: r1.min(r2),
            start_col: c1.min(c2),
            end_row: r1.max(r2),
            end_col: c1.max(c2),
        }
    }

    /// Create a single-cell range.
    pub fn single(row: usize, col: usize) -> Self {
        Self {
            start_row: row,
            start_col: col,
            end_row: row,
            end_col: col,
        }
    }

    /// Check if this range contains a cell.
    pub fn contains(&self, row: usize, col: usize) -> bool {
        row >= self.start_row && row <= self.end_row &&
        col >= self.start_col && col <= self.end_col
    }

    /// Number of cells in this range.
    pub fn cell_count(&self) -> usize {
        (self.end_row - self.start_row + 1) * (self.end_col - self.start_col + 1)
    }

    /// Iterate over all cells in this range (row-major order).
    pub fn cells(&self) -> impl Iterator<Item = (usize, usize)> {
        let start_row = self.start_row;
        let end_row = self.end_row;
        let start_col = self.start_col;
        let end_col = self.end_col;

        (start_row..=end_row).flat_map(move |r| {
            (start_col..=end_col).map(move |c| (r, c))
        })
    }

    /// Check if this is a single cell.
    pub fn is_single(&self) -> bool {
        self.start_row == self.end_row && self.start_col == self.end_col
    }
}

/// The selection model: ordered list of ranges with an active cell.
#[derive(Debug, Clone)]
pub struct Selection {
    ranges: Vec<Range>,
    active_range: usize,
    anchor: (usize, usize),
}

impl Selection {
    /// Create a new selection with a single cell.
    pub fn new(row: usize, col: usize) -> Self {
        Self {
            ranges: vec![Range::single(row, col)],
            active_range: 0,
            anchor: (row, col),
        }
    }

    /// Get the active cell (top-left of active range).
    pub fn active_cell(&self) -> (usize, usize) {
        let range = &self.ranges[self.active_range];
        (range.start_row, range.start_col)
    }

    /// Get the anchor cell (for extending selections).
    pub fn anchor(&self) -> (usize, usize) {
        self.anchor
    }

    /// Get all ranges.
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }

    /// Check if a cell is selected.
    pub fn contains(&self, row: usize, col: usize) -> bool {
        self.ranges.iter().any(|r| r.contains(row, col))
    }

    /// Check if selection is a single cell.
    pub fn is_single_cell(&self) -> bool {
        self.ranges.len() == 1 && self.ranges[0].is_single()
    }

    /// Total number of selected cells.
    pub fn cell_count(&self) -> usize {
        self.ranges.iter().map(|r| r.cell_count()).sum()
    }

    /// Iterate over all selected cells.
    pub fn all_cells(&self) -> impl Iterator<Item = (usize, usize)> + '_ {
        self.ranges.iter().flat_map(|r| r.cells())
    }

    /// Set selection to a single cell (click).
    pub fn select_cell(&mut self, row: usize, col: usize) {
        self.ranges = vec![Range::single(row, col)];
        self.active_range = 0;
        self.anchor = (row, col);
    }

    /// Extend the active range from anchor to the given cell (shift+click/arrow).
    pub fn extend_to(&mut self, row: usize, col: usize) {
        let (anchor_row, anchor_col) = self.anchor;
        self.ranges[self.active_range] = Range::new(anchor_row, anchor_col, row, col);
    }

    /// Add a new range (ctrl+click).
    pub fn add_cell(&mut self, row: usize, col: usize) {
        self.ranges.push(Range::single(row, col));
        self.active_range = self.ranges.len() - 1;
        self.anchor = (row, col);
    }

    /// Add a new range from anchor to cell (ctrl+shift+click).
    pub fn add_range_to(&mut self, row: usize, col: usize) {
        let (anchor_row, anchor_col) = self.anchor;
        self.ranges.push(Range::new(anchor_row, anchor_col, row, col));
        self.active_range = self.ranges.len() - 1;
    }

    /// Move active cell by delta, collapsing to single cell.
    pub fn move_by(&mut self, d_row: isize, d_col: isize, max_row: usize, max_col: usize) {
        let (row, col) = self.active_cell();
        let new_row = (row as isize + d_row).clamp(0, max_row as isize - 1) as usize;
        let new_col = (col as isize + d_col).clamp(0, max_col as isize - 1) as usize;
        self.select_cell(new_row, new_col);
    }

    /// Extend selection by delta from current extent.
    pub fn extend_by(&mut self, d_row: isize, d_col: isize, max_row: usize, max_col: usize) {
        let range = &self.ranges[self.active_range];

        // Figure out which edge to extend based on anchor position
        let (anchor_row, anchor_col) = self.anchor;
        let current_row = if anchor_row == range.start_row { range.end_row } else { range.start_row };
        let current_col = if anchor_col == range.start_col { range.end_col } else { range.start_col };

        let new_row = (current_row as isize + d_row).clamp(0, max_row as isize - 1) as usize;
        let new_col = (current_col as isize + d_col).clamp(0, max_col as isize - 1) as usize;

        self.ranges[self.active_range] = Range::new(anchor_row, anchor_col, new_row, new_col);
    }

    /// Select entire column(s) based on current selection
    pub fn select_column(&mut self, max_row: usize) {
        let range = &self.ranges[self.active_range];
        let start_col = range.start_col;
        let end_col = range.end_col;
        self.ranges[self.active_range] = Range::new(0, start_col, max_row - 1, end_col);
        // Keep anchor at top of column
        self.anchor = (0, start_col);
    }

    /// Select entire row(s) based on current selection
    pub fn select_row(&mut self, max_col: usize) {
        let range = &self.ranges[self.active_range];
        let start_row = range.start_row;
        let end_row = range.end_row;
        self.ranges[self.active_range] = Range::new(start_row, 0, end_row, max_col - 1);
        // Keep anchor at start of row
        self.anchor = (start_row, 0);
    }

    /// Check if selection is an entire column (spans all rows)
    pub fn is_full_column(&self, max_row: usize) -> bool {
        let range = &self.ranges[self.active_range];
        range.start_row == 0 && range.end_row == max_row - 1
    }

    /// Check if selection is an entire row (spans all columns)
    pub fn is_full_row(&self, max_col: usize) -> bool {
        let range = &self.ranges[self.active_range];
        range.start_col == 0 && range.end_col == max_col - 1
    }

    /// Get the column range of the selection
    pub fn col_range(&self) -> (usize, usize) {
        let range = &self.ranges[self.active_range];
        (range.start_col, range.end_col)
    }

    /// Get the row range of the selection
    pub fn row_range(&self) -> (usize, usize) {
        let range = &self.ranges[self.active_range];
        (range.start_row, range.end_row)
    }
}

impl Default for Selection {
    fn default() -> Self {
        Self::new(0, 0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_range_single() {
        let r = Range::single(5, 3);
        assert!(r.contains(5, 3));
        assert!(!r.contains(5, 4));
        assert!(r.is_single());
        assert_eq!(r.cell_count(), 1);
    }

    #[test]
    fn test_range_multi() {
        let r = Range::new(1, 1, 3, 2);
        assert!(r.contains(1, 1));
        assert!(r.contains(2, 2));
        assert!(r.contains(3, 1));
        assert!(!r.contains(0, 0));
        assert!(!r.is_single());
        assert_eq!(r.cell_count(), 6); // 3 rows x 2 cols
    }

    #[test]
    fn test_range_normalizes() {
        let r = Range::new(5, 5, 1, 1);
        assert_eq!(r.start_row, 1);
        assert_eq!(r.start_col, 1);
        assert_eq!(r.end_row, 5);
        assert_eq!(r.end_col, 5);
    }

    #[test]
    fn test_selection_extend() {
        let mut sel = Selection::new(2, 2);
        sel.extend_to(4, 5);

        assert_eq!(sel.anchor(), (2, 2));
        assert!(sel.contains(2, 2));
        assert!(sel.contains(3, 3));
        assert!(sel.contains(4, 5));
        assert!(!sel.contains(1, 1));
    }
}
