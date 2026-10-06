//! Line layout: column widths, row heights, hidden and frozen lines.
//!
//! Presentation state, but replicated: collaboration operations set it,
//! convergence checksums cover it, and structural edits move it with the
//! cells (the same rules as `visigrid_io::json::SheetLayout`, which carries
//! it in files).

use std::collections::{BTreeMap, BTreeSet};

use serde::{Deserialize, Serialize};

use crate::structural::shift_span;

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct LineLayout {
    /// Columns whose width differs from the default, in points.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub col_widths: BTreeMap<usize, f32>,
    /// Rows whose height differs from the default, in points.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub row_heights: BTreeMap<usize, f32>,
    #[serde(default, skip_serializing_if = "BTreeSet::is_empty")]
    pub hidden_rows: BTreeSet<usize>,
    #[serde(default, skip_serializing_if = "BTreeSet::is_empty")]
    pub hidden_cols: BTreeSet<usize>,
    #[serde(default)]
    pub frozen_rows: usize,
    #[serde(default)]
    pub frozen_cols: usize,
}

impl LineLayout {
    pub fn is_default(&self) -> bool {
        *self == LineLayout::default()
    }

    /// Move lines with a row or column insert or delete.
    pub fn shift_for_structural(&mut self, at: usize, count: usize, delete: bool, is_row: bool) {
        let shift_keys = |m: &BTreeMap<usize, f32>| -> BTreeMap<usize, f32> {
            m.iter().filter_map(|(k, v)| shift_span(*k, *k, at, count, delete).map(|(nk, _)| (nk, *v))).collect()
        };
        let shift_set = |set: &BTreeSet<usize>| -> BTreeSet<usize> {
            set.iter().filter_map(|k| shift_span(*k, *k, at, count, delete).map(|(nk, _)| nk)).collect()
        };
        let frozen = |n: usize| -> usize {
            if at >= n {
                n
            } else if delete {
                n.saturating_sub(count.min(n - at))
            } else {
                n + count
            }
        };
        if is_row {
            self.row_heights = shift_keys(&self.row_heights);
            self.hidden_rows = shift_set(&self.hidden_rows);
            self.frozen_rows = frozen(self.frozen_rows);
        } else {
            self.col_widths = shift_keys(&self.col_widths);
            self.hidden_cols = shift_set(&self.hidden_cols);
            self.frozen_cols = frozen(self.frozen_cols);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lines_move_with_inserts_and_deletes() {
        let mut l = LineLayout::default();
        l.row_heights.insert(5, 40.0);
        l.hidden_rows.insert(8);
        l.frozen_rows = 2;
        l.shift_for_structural(0, 3, false, true);
        assert_eq!(l.row_heights.get(&8), Some(&40.0));
        assert!(l.hidden_rows.contains(&11));
        assert_eq!(l.frozen_rows, 5);
        l.shift_for_structural(7, 3, true, true);
        assert!(l.row_heights.is_empty(), "a deleted line's size goes with it");
        assert!(l.hidden_rows.contains(&8));
        l.shift_for_structural(1, 10, true, true);
        assert_eq!(l.frozen_rows, 1);
    }
}
