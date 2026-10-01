//! Editable Tables: identity and schema, separate from pivot output.
//!
//! Cells remain in Sheet's column store. This module never owns a second
//! copy of a table's body. See docs/table-semantics.md for the staged contract.

use crate::sheet::{NUM_COLS, NUM_ROWS};
use serde::{Deserialize, Serialize};
use std::collections::HashSet;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct TableId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct TableColumnId(pub u64);

/// Inclusive canonical coordinates; the first row is always the header.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableRange {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

impl TableRange {
    pub fn validate(self, rows: usize, cols: usize) -> Result<(), String> {
        if self.start_row > self.end_row || self.start_col > self.end_col {
            return Err("Table range must be a nonempty rectangle with its header first.".into());
        }
        if self.end_row >= rows.min(NUM_ROWS) || self.end_col >= cols.min(NUM_COLS) {
            return Err("Table range extends beyond the sheet.".into());
        }
        Ok(())
    }

    pub fn contains(self, row: usize, col: usize) -> bool {
        row >= self.start_row && row <= self.end_row && col >= self.start_col && col <= self.end_col
    }

    pub fn intersects(self, other: Self) -> bool {
        self.start_row <= other.end_row
            && other.start_row <= self.end_row
            && self.start_col <= other.end_col
            && other.start_col <= self.end_col
    }

    pub fn data_rows(self) -> usize {
        self.end_row - self.start_row
    }
    pub fn width(self) -> usize {
        self.end_col - self.start_col + 1
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableColumn {
    pub id: TableColumnId,
    pub name: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableStyle {
    pub banded_rows: bool,
}

impl Default for TableStyle {
    fn default() -> Self {
        Self { banded_rows: true }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DataTable {
    pub id: TableId,
    pub name: String,
    pub range: TableRange,
    pub columns: Vec<TableColumn>,
    /// Column IDs are local to this table and never reused after shrinking.
    pub next_column_id: u64,
    pub style: TableStyle,
}

impl DataTable {
    pub fn column_by_name(&self, name: &str) -> Option<&TableColumn> {
        let key = name.to_lowercase();
        self.columns.iter().find(|c| c.name.to_lowercase() == key)
    }

    pub fn validate(&self, rows: usize, cols: usize) -> Result<(), String> {
        self.range.validate(rows, cols)?;
        validate_table_name(&self.name)?;
        if self.id.0 == 0 || self.id.0 == u64::MAX || self.columns.len() != self.range.width() {
            return Err("Invalid table identity or column count.".into());
        }
        let mut names = HashSet::new();
        let mut ids = HashSet::new();
        for col in &self.columns {
            validate_column_name(&col.name)?;
            if !names.insert(col.name.to_lowercase())
                || !ids.insert(col.id)
                || col.id.0 == 0
                || col.id.0 >= self.next_column_id
            {
                return Err(
                    "Table columns must have unique names and stable, allocated IDs.".into(),
                );
            }
        }
        Ok(())
    }
}

pub fn validate_table_name(name: &str) -> Result<(), String> {
    crate::named_range::is_valid_name(name)?;
    if !name
        .chars()
        .enumerate()
        .all(|(i, c)| c == '_' || c.is_ascii_alphabetic() || (i > 0 && c.is_ascii_digit()))
    {
        return Err("Table names use letters, digits and underscores, starting with a letter or underscore.".into());
    }
    let upper = name.to_ascii_uppercase();
    let r1c1 = upper
        .strip_prefix('R')
        .and_then(|s| s.split_once('C'))
        .is_some_and(|(r, c)| {
            !r.is_empty()
                && !c.is_empty()
                && r.bytes().all(|b| b.is_ascii_digit())
                && c.bytes().all(|b| b.is_ascii_digit())
        });
    if upper == "R" || upper == "C" || r1c1 {
        return Err("Table name conflicts with a row or column reference.".into());
    }
    Ok(())
}

pub fn validate_column_name(name: &str) -> Result<(), String> {
    if name.is_empty()
        || name.trim() != name
        || name.starts_with('=')
        || name.chars().any(char::is_control)
    {
        return Err("Column names must be nonempty text without surrounding whitespace or control characters.".into());
    }
    Ok(())
}

/// Reserve original nonblank names first so generated suffixes never steal a
/// later valid header (Amount, Amount, Amount2 -> Amount, Amount3, Amount2).
pub fn normalize_headers(headers: &[String]) -> Vec<String> {
    let normalized: Vec<_> = headers
        .iter()
        .map(|s| {
            let s = s.trim().to_string();
            if validate_column_name(&s).is_ok() {
                s
            } else {
                String::new()
            }
        })
        .collect();
    let reserved: HashSet<_> = normalized
        .iter()
        .filter(|s| !s.is_empty())
        .map(|s| s.to_lowercase())
        .collect();
    let mut used = HashSet::new();
    normalized
        .iter()
        .enumerate()
        .map(|(i, original)| {
            if !original.is_empty() && used.insert(original.to_lowercase()) {
                return original.clone();
            }
            let base = if original.is_empty() {
                format!("Column{}", i + 1)
            } else {
                original.clone()
            };
            let mut candidate = base.clone();
            let mut suffix = 2;
            while reserved.contains(&candidate.to_lowercase())
                || used.contains(&candidate.to_lowercase())
            {
                candidate = format!("{base}{suffix}");
                suffix += 1;
            }
            used.insert(candidate.to_lowercase());
            candidate
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn header_normalization_reserves_valid_later_names() {
        let input = [" Amount ", "amount", "Amount2", "", "Column4", "é", "É"];
        assert_eq!(
            normalize_headers(&input.map(str::to_string)),
            ["Amount", "amount3", "Amount2", "Column42", "Column4", "é", "É2"]
        );
    }

    #[test]
    fn reject_ambiguous_table_names() {
        for name in [
            "A1",
            "R1C1",
            "r25c12",
            "R",
            "C",
            "TRUE",
            "SUM",
            "9Sales",
            "Sales Data",
            "Sales.Data",
        ] {
            assert!(validate_table_name(name).is_err(), "{name}");
        }
        for name in ["Sales", "Table1", "_Sales_2026"] {
            assert!(validate_table_name(name).is_ok(), "{name}");
        }
    }
}
