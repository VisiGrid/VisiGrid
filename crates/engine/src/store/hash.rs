//! The one-hash-entry-per-cell store VisiGrid used through 0.37 (#18).
//!
//! Kept as the reference implementation: the differential tests in
//! `store/mod.rs` drive it and the column store through the same operations
//! and require the same result.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use crate::cell::{Cell, CellRef};
use crate::sheet::{NUM_COLS, NUM_ROWS};

/// Row and column as stored: `u32` pairs, where callers speak `usize`.
type CellKey = (u32, u32);

/// Coordinates to the storage key. Debug builds catch a coordinate past the
/// grid here rather than storing a cell nothing can address.
#[inline]
fn key(row: usize, col: usize) -> CellKey {
    debug_assert!(row < NUM_ROWS && col < NUM_COLS, "cell ({row}, {col}) is outside the grid");
    (row as u32, col as u32)
}

#[inline]
fn coords((row, col): CellKey) -> (usize, usize) {
    (row as usize, col as usize)
}

/// The cells of one sheet.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(transparent)]
pub(crate) struct HashStore {
    map: HashMap<CellKey, Cell>,
}

impl HashStore {
    /// Number of stored cells (with a value, a format, or metadata).
    pub fn len(&self) -> usize {
        self.map.len()
    }

    /// Make room for about `additional` more cells.
    pub fn reserve(&mut self, additional: usize) {
        self.map.reserve(additional);
    }

    pub fn get(&self, row: usize, col: usize) -> Option<CellRef<'_>> {
        self.map.get(&key(row, col)).map(CellRef::new)
    }

    /// Every stored cell, in no particular order.
    pub fn iter(&self) -> impl Iterator<Item = ((usize, usize), CellRef<'_>)> {
        self.map.iter().map(|(k, cell)| (coords(*k), CellRef::new(cell)))
    }

    /// Change a cell that exists; `None` when there is no cell there.
    pub fn update<R>(&mut self, row: usize, col: usize, f: impl FnOnce(&mut Cell) -> R) -> Option<R> {
        self.map.get_mut(&key(row, col)).map(f)
    }

    /// Change a cell, creating it from `init` first if it does not exist.
    pub fn upsert<R>(
        &mut self,
        row: usize,
        col: usize,
        init: impl FnOnce() -> Cell,
        f: impl FnOnce(&mut Cell) -> R,
    ) -> R {
        f(self.map.entry(key(row, col)).or_insert_with(init))
    }

    /// Give a cell a new format without touching its value, creating it from
    /// `init` if there is none.
    pub fn set_format(&mut self, row: usize, col: usize, init: impl FnOnce() -> Cell, format: std::sync::Arc<crate::cell::CellFormat>) {
        self.map.entry(key(row, col)).or_insert_with(init).format = format;
    }

    pub fn remove(&mut self, row: usize, col: usize) -> Option<Cell> {
        self.map.remove(&key(row, col))
    }

    /// Move cells at or below `at` down by `count` rows. Cells that would land
    /// at or past `limit` are dropped (callers refuse inserts that would push
    /// data off the grid before getting here).
    pub fn insert_rows(&mut self, at: usize, count: usize, limit: usize) {
        self.shift(|(r, c)| if r >= at { (r + count < limit).then_some((r + count, c)) } else { Some((r, c)) });
    }

    /// Delete `count` rows from `start`; cells below move up.
    pub fn delete_rows(&mut self, start: usize, count: usize) {
        let end = start + count;
        self.shift(|(r, c)| {
            if (start..end).contains(&r) {
                None
            } else if r >= end {
                Some((r - count, c))
            } else {
                Some((r, c))
            }
        });
    }

    /// Move cells at or right of `at` right by `count` columns, dropping any
    /// that would land at or past `limit`.
    pub fn insert_cols(&mut self, at: usize, count: usize, limit: usize) {
        self.shift(|(r, c)| if c >= at { (c + count < limit).then_some((r, c + count)) } else { Some((r, c)) });
    }

    /// Delete `count` columns from `start`; cells to the right move left.
    pub fn delete_cols(&mut self, start: usize, count: usize) {
        let end = start + count;
        self.shift(|(r, c)| {
            if (start..end).contains(&c) {
                None
            } else if c >= end {
                Some((r, c - count))
            } else {
                Some((r, c))
            }
        });
    }

    /// Re-key every cell through `to`, dropping those it maps to `None`.
    /// Cells that stay put are untouched.
    fn shift(&mut self, to: impl Fn((usize, usize)) -> Option<(usize, usize)>) {
        let moving: Vec<CellKey> = self
            .map
            .keys()
            .filter(|k| to(coords(**k)) != Some(coords(**k)))
            .copied()
            .collect();
        let moved: Vec<(CellKey, Cell)> = moving
            .into_iter()
            .filter_map(|k| self.map.remove(&k).map(|cell| (k, cell)))
            .collect();
        for (k, cell) in moved {
            if let Some((r, c)) = to(coords(k)) {
                self.map.insert(key(r, c), cell);
            }
        }
    }
}
