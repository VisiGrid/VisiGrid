use serde::{Deserialize, Serialize};
use crate::sheet::Sheet;
use crate::named_range::{NamedRange, NamedRangeStore};

/// A workbook containing multiple sheets
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    sheets: Vec<Sheet>,
    active_sheet: usize,
    #[serde(default)]
    named_ranges: NamedRangeStore,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    /// Create a new workbook with one default sheet
    pub fn new() -> Self {
        let mut sheet = Sheet::new(65536, 256);
        sheet.name = "Sheet1".to_string();
        Self {
            sheets: vec![sheet],
            active_sheet: 0,
            named_ranges: NamedRangeStore::new(),
        }
    }

    /// Get the number of sheets
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Get the active sheet index
    pub fn active_sheet_index(&self) -> usize {
        self.active_sheet
    }

    /// Set the active sheet by index
    pub fn set_active_sheet(&mut self, index: usize) -> bool {
        if index < self.sheets.len() {
            self.active_sheet = index;
            true
        } else {
            false
        }
    }

    /// Get a reference to the active sheet
    pub fn active_sheet(&self) -> &Sheet {
        &self.sheets[self.active_sheet]
    }

    /// Get a mutable reference to the active sheet
    pub fn active_sheet_mut(&mut self) -> &mut Sheet {
        &mut self.sheets[self.active_sheet]
    }

    /// Get a reference to a sheet by index
    pub fn sheet(&self, index: usize) -> Option<&Sheet> {
        self.sheets.get(index)
    }

    /// Get a mutable reference to a sheet by index
    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Sheet> {
        self.sheets.get_mut(index)
    }

    /// Get all sheet names
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    /// Add a new sheet and return its index
    pub fn add_sheet(&mut self) -> usize {
        let sheet_num = self.sheets.len() + 1;
        let mut new_name = format!("Sheet{}", sheet_num);

        // Ensure unique name
        while self.sheets.iter().any(|s| s.name == new_name) {
            let num: usize = new_name.strip_prefix("Sheet")
                .and_then(|n| n.parse().ok())
                .unwrap_or(sheet_num);
            new_name = format!("Sheet{}", num + 1);
        }

        let mut sheet = Sheet::new(65536, 256);
        sheet.name = new_name;
        self.sheets.push(sheet);
        self.sheets.len() - 1
    }

    /// Add a new sheet with a specific name
    pub fn add_sheet_named(&mut self, name: &str) -> usize {
        let mut sheet = Sheet::new(65536, 256);
        sheet.name = name.to_string();
        self.sheets.push(sheet);
        self.sheets.len() - 1
    }

    /// Delete a sheet by index
    /// Returns false if it's the last sheet (can't delete)
    pub fn delete_sheet(&mut self, index: usize) -> bool {
        if self.sheets.len() <= 1 || index >= self.sheets.len() {
            return false;
        }

        self.sheets.remove(index);

        // Adjust active sheet if needed
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        } else if self.active_sheet > index {
            self.active_sheet -= 1;
        }

        true
    }

    /// Rename a sheet
    pub fn rename_sheet(&mut self, index: usize, new_name: &str) -> bool {
        if let Some(sheet) = self.sheets.get_mut(index) {
            sheet.name = new_name.to_string();
            true
        } else {
            false
        }
    }

    /// Move to the next sheet
    pub fn next_sheet(&mut self) -> bool {
        if self.active_sheet + 1 < self.sheets.len() {
            self.active_sheet += 1;
            true
        } else {
            false
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self) -> bool {
        if self.active_sheet > 0 {
            self.active_sheet -= 1;
            true
        } else {
            false
        }
    }

    /// Get all sheets (for serialization)
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Create a workbook from sheets (for deserialization)
    pub fn from_sheets(sheets: Vec<Sheet>, active: usize) -> Self {
        let active_sheet = active.min(sheets.len().saturating_sub(1));
        Self {
            sheets,
            active_sheet,
            named_ranges: NamedRangeStore::new(),
        }
    }

    // =========================================================================
    // Named Range Management
    // =========================================================================

    /// Get a reference to the named range store
    pub fn named_ranges(&self) -> &NamedRangeStore {
        &self.named_ranges
    }

    /// Get a mutable reference to the named range store
    pub fn named_ranges_mut(&mut self) -> &mut NamedRangeStore {
        &mut self.named_ranges
    }

    /// Define a named range for a single cell (convenience method)
    pub fn define_name_for_cell(
        &mut self,
        name: &str,
        sheet: usize,
        row: usize,
        col: usize,
    ) -> Result<(), String> {
        let range = NamedRange::cell(name, sheet, row, col);
        self.named_ranges.set(range)
    }

    /// Define a named range for a cell range (convenience method)
    pub fn define_name_for_range(
        &mut self,
        name: &str,
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<(), String> {
        let range = NamedRange::range(name, sheet, start_row, start_col, end_row, end_col);
        self.named_ranges.set(range)
    }

    /// Get a named range by name (case-insensitive)
    pub fn get_named_range(&self, name: &str) -> Option<&NamedRange> {
        self.named_ranges.get(name)
    }

    /// Rename a named range
    pub fn rename_named_range(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        self.named_ranges.rename(old_name, new_name)
    }

    /// Delete a named range
    pub fn delete_named_range(&mut self, name: &str) -> bool {
        self.named_ranges.remove(name).is_some()
    }

    /// Find all named ranges that reference a specific cell
    pub fn named_ranges_for_cell(&self, sheet: usize, row: usize, col: usize) -> Vec<&NamedRange> {
        self.named_ranges.find_by_cell(sheet, row, col)
    }

    /// List all named ranges
    pub fn list_named_ranges(&self) -> Vec<&NamedRange> {
        self.named_ranges.list()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.sheet_count(), 1);
        assert_eq!(wb.active_sheet_index(), 0);
        assert_eq!(wb.active_sheet().name, "Sheet1");
    }

    #[test]
    fn test_add_sheet() {
        let mut wb = Workbook::new();
        let idx = wb.add_sheet();
        assert_eq!(idx, 1);
        assert_eq!(wb.sheet_count(), 2);
        assert_eq!(wb.sheet(1).unwrap().name, "Sheet2");
    }

    #[test]
    fn test_navigation() {
        let mut wb = Workbook::new();
        wb.add_sheet();
        wb.add_sheet();

        assert_eq!(wb.active_sheet_index(), 0);
        assert!(wb.next_sheet());
        assert_eq!(wb.active_sheet_index(), 1);
        assert!(wb.next_sheet());
        assert_eq!(wb.active_sheet_index(), 2);
        assert!(!wb.next_sheet()); // Can't go further

        assert!(wb.prev_sheet());
        assert_eq!(wb.active_sheet_index(), 1);
    }

    #[test]
    fn test_delete_sheet() {
        let mut wb = Workbook::new();
        wb.add_sheet();
        wb.add_sheet();
        wb.set_active_sheet(2);

        assert!(wb.delete_sheet(1));
        assert_eq!(wb.sheet_count(), 2);
        assert_eq!(wb.active_sheet_index(), 1); // Adjusted

        // Can't delete last sheet
        assert!(wb.delete_sheet(0));
        assert!(!wb.delete_sheet(0)); // Last one, can't delete
    }
}
