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
    /// Formula expressed at its authored origin in this column. Cell contents
    /// are authoritative: a differing formula/value (including blank) is an
    /// exception. This keeps every editing/history path consistent.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Offset from header to the authored formula's row. Retaining its origin
    /// avoids losing relative references when projecting above row 1.
    #[serde(
        default = "formula_origin_default",
        skip_serializing_if = "formula_origin_is_default"
    )]
    pub formula_origin: usize,
}

fn formula_origin_default() -> usize {
    1
}
fn formula_origin_is_default(value: &usize) -> bool {
    *value == 1
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableStyle {
    pub banded_rows: bool,
    /// Built-in OOXML style identity, retained for Excel round trips. None is
    /// Excel's explicit "no style". Native rendering still uses the app theme.
    #[serde(default = "default_excel_style")]
    pub excel_style: Option<String>,
    #[serde(default)]
    pub banded_columns: bool,
    #[serde(default)]
    pub first_column: bool,
    #[serde(default)]
    pub last_column: bool,
}

fn default_excel_style() -> Option<String> {
    Some("TableStyleMedium2".into())
}

impl TableStyle {
    pub fn is_builtin_excel_style(name: &str) -> bool {
        [
            ("TableStyleLight", 21),
            ("TableStyleMedium", 28),
            ("TableStyleDark", 11),
        ]
        .iter()
        .any(|(prefix, max)| {
            name.strip_prefix(prefix)
                .and_then(|n| n.parse::<u8>().ok().map(|v| (n, v)))
                .is_some_and(|(n, v)| v > 0 && v <= *max && n == v.to_string())
        })
    }
}

impl Default for TableStyle {
    fn default() -> Self {
        Self {
            banded_rows: true,
            excel_style: default_excel_style(),
            banded_columns: false,
            first_column: false,
            last_column: false,
        }
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
    /// The recipe this Table is loaded from, if any. Refresh re-runs it and
    /// replaces the records; a Table without one is edited by hand only.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source: Option<TableSource>,
    /// Excel totals metadata; range continues to describe header and data only.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals: Option<TableTotals>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub saved_views: Vec<crate::table_view::NamedTableView>,
}

/// Retained totals-row settings. A visible totals row is immediately below the body.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableTotals {
    pub visible: bool,
    pub shown: Option<bool>,
    #[serde(default)]
    pub hidden_rows: std::collections::BTreeSet<usize>,
    pub columns: Vec<TableTotal>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableTotal {
    pub function: Option<String>,
    pub label: Option<String>,
    pub formula: Option<String>,
}

/// Where a recipe-backed Table comes from, and its last good refresh.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TableSource {
    /// Path to the `.recipe.toml`, as the user chose it.
    pub recipe: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refreshed: Option<RefreshStamp>,
}

/// The last refresh that published: what it read and what it loaded.
/// A failed refresh never changes it.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RefreshStamp {
    /// RFC 3339, UTC.
    pub at: String,
    /// The source file read, as resolved for that run.
    pub source: String,
    /// The snapshot's content hash.
    pub snapshot: String,
    pub rows: usize,
}

impl DataTable {
    pub fn totals_row(&self) -> Option<usize> {
        self.totals
            .as_ref()
            .filter(|t| t.visible)
            .map(|_| self.range.end_row + 1)
    }

    pub fn full_range(&self) -> TableRange {
        TableRange {
            end_row: self.totals_row().unwrap_or(self.range.end_row),
            ..self.range
        }
    }

    pub fn column_by_name(&self, name: &str) -> Option<&TableColumn> {
        let key = name.to_lowercase();
        self.columns.iter().find(|c| c.name.to_lowercase() == key)
    }

    pub fn formula_at(&self, row: usize, col: usize) -> Option<String> {
        if col < self.range.start_col || col > self.range.end_col {
            return None;
        }
        let column = &self.columns[col - self.range.start_col];
        let formula = column.formula.as_ref()?;
        Some(crate::formula::parser::adjust_formula_refs(
            formula,
            row as i32 - (self.range.start_row + column.formula_origin) as i32,
            0,
        ))
    }

    pub fn validate(&self, rows: usize, cols: usize) -> Result<(), String> {
        self.range.validate(rows, cols)?;
        self.full_range().validate(rows, cols)?;
        if let Some(totals) = &self.totals {
            if totals.hidden_rows.iter().any(|r| *r >= rows.min(NUM_ROWS)) {
                return Err("Invalid hidden-row metadata for totals.".into());
            }
            if totals.columns.len() != self.columns.len() {
                return Err("Invalid totals-column count.".into());
            }
            for total in &totals.columns {
                if total.function.as_deref().is_some_and(|f| {
                    !matches!(
                        f,
                        "none"
                            | "sum"
                            | "min"
                            | "max"
                            | "average"
                            | "count"
                            | "countNums"
                            | "stdDev"
                            | "var"
                            | "custom"
                    )
                }) {
                    return Err("Unsupported totals function.".into());
                }
                if total.formula.as_ref().is_some_and(|f| {
                    !f.starts_with('=') || crate::formula::parser::parse(f).is_err()
                }) {
                    return Err("Invalid totals formula.".into());
                }
            }
        }
        validate_table_name(&self.name)?;
        if self
            .style
            .excel_style
            .as_deref()
            .is_some_and(|name| !TableStyle::is_builtin_excel_style(name))
        {
            return Err("Unsupported Excel Table style identity.".into());
        }
        if self.id.0 == 0 || self.id.0 == u64::MAX || self.columns.len() != self.range.width() {
            return Err("Invalid table identity or column count.".into());
        }
        let mut names = HashSet::new();
        let mut ids = HashSet::new();
        for col in &self.columns {
            validate_column_name(&col.name)?;
            if let Some(formula) = &col.formula {
                if col.formula_origin > NUM_ROWS
                    || !formula.starts_with('=')
                    || crate::formula::parser::parse(formula).is_err()
                {
                    return Err("Invalid calculated-column formula.".into());
                }
            }
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
        if self.saved_views.len() > crate::table_view::MAX_NAMED_TABLE_VIEWS {
            return Err("A Table can have at most 64 named views.".into());
        }
        let mut view_names = HashSet::new();
        for saved in &self.saved_views {
            crate::table_view::validate_view_name(&saved.name)?;
            if !view_names.insert(saved.name.to_lowercase()) || saved.view.table != self.id {
                return Err("Saved views must have unique names and belong to their Table.".into());
            }
            // Do not call validate_schema here: it validates this DataTable.
            saved.view.resolve(self).map_err(|e| format!(
                "Saved view '{}': {e} Update or delete this saved view first.", saved.name
            ))?;
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
