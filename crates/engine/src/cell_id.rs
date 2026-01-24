//! Cell identity for dependency graph.
//!
//! A `CellId` uniquely identifies a cell across all sheets in a workbook.

use crate::sheet::SheetId;

/// Unique identifier for a cell in a workbook.
///
/// Combines sheet identity with row/column coordinates.
/// Used as graph nodes in the dependency graph.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellId {
    /// The sheet this cell belongs to (stable, never reused after deletion)
    pub sheet: SheetId,
    /// Row index (0-based)
    pub row: usize,
    /// Column index (0-based)
    pub col: usize,
}

impl CellId {
    /// Create a new CellId.
    #[inline]
    pub fn new(sheet: SheetId, row: usize, col: usize) -> Self {
        Self { sheet, row, col }
    }
}

impl std::fmt::Display for CellId {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Convert column to letter(s): 0=A, 1=B, ..., 25=Z, 26=AA, etc.
        let col_str = col_to_letters(self.col);
        write!(f, "Sheet{}!{}{}", self.sheet.raw(), col_str, self.row + 1)
    }
}

/// Convert 0-based column index to Excel-style letter(s).
fn col_to_letters(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_id_equality() {
        let a = CellId::new(SheetId::from_raw(1), 0, 0);
        let b = CellId::new(SheetId::from_raw(1), 0, 0);
        let c = CellId::new(SheetId::from_raw(2), 0, 0);

        assert_eq!(a, b);
        assert_ne!(a, c);
    }

    #[test]
    fn test_cell_id_hash() {
        use std::collections::HashSet;

        let mut set = HashSet::new();
        set.insert(CellId::new(SheetId::from_raw(1), 0, 0));
        set.insert(CellId::new(SheetId::from_raw(1), 0, 0)); // duplicate
        set.insert(CellId::new(SheetId::from_raw(1), 1, 0));

        assert_eq!(set.len(), 2);
    }

    #[test]
    fn test_col_to_letters() {
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(1), "B");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
        assert_eq!(col_to_letters(27), "AB");
        assert_eq!(col_to_letters(701), "ZZ");
        assert_eq!(col_to_letters(702), "AAA");
    }

    #[test]
    fn test_display() {
        let cell = CellId::new(SheetId::from_raw(1), 0, 0);
        assert_eq!(format!("{}", cell), "Sheet1!A1");

        let cell2 = CellId::new(SheetId::from_raw(2), 9, 26);
        assert_eq!(format!("{}", cell2), "Sheet2!AA10");
    }
}
