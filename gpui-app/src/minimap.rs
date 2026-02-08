//! Minimap cache and bucket computation for row-density navigator.
//!
//! The minimap shows a vertical density strip indicating where data exists
//! across the used row range. Bucket computation is lazy: recompute when
//! the sheet revision changes or minimap height changes.

use visigrid_engine::cell::CellValue;
use visigrid_engine::sheet::Sheet;

/// Cached bucket data for minimap rendering.
///
/// Each bucket maps to a vertical pixel in the minimap strip.
/// Value is 0 (empty) or 140 (has content) — binary density in Phase 1.
#[derive(Clone, Debug)]
pub struct MinimapCache {
    /// First row with content (0-indexed)
    pub used_min: usize,
    /// Last row with content (0-indexed)
    pub used_max: usize,
    /// Density buckets: len == height_px, each 0 or 140
    pub buckets: Vec<u8>,
    /// Current minimap height in logical pixels
    pub height_px: usize,
    /// cells_rev at last compute (0 = never computed)
    pub last_revision: u64,
    /// Sheet ID at last compute (invalidate on sheet switch)
    pub last_sheet_id: u64,
}

impl Default for MinimapCache {
    fn default() -> Self {
        Self {
            used_min: 0,
            used_max: 0,
            buckets: Vec::new(),
            height_px: 0,
            last_revision: 0,
            last_sheet_id: 0,
        }
    }
}

impl MinimapCache {
    /// Returns true if the cache has valid data (non-empty sheet with buckets).
    pub fn has_data(&self) -> bool {
        !self.buckets.is_empty() && self.used_max >= self.used_min
    }

    /// Total row range covered by this cache.
    pub fn row_range(&self) -> usize {
        if self.has_data() {
            self.used_max - self.used_min + 1
        } else {
            0
        }
    }
}

/// Returns true if a cell has meaningful content (value or formula, not empty).
fn cell_has_content(cell: &visigrid_engine::cell::Cell) -> bool {
    !matches!(cell.value, CellValue::Empty)
}

/// Pass 1: find the min and max row that have content.
/// Returns None if the sheet has no content.
fn find_row_extent(sheet: &Sheet) -> Option<(usize, usize)> {
    let mut min = usize::MAX;
    let mut max = 0;
    let mut found = false;
    for (&(r, _), cell) in sheet.cells_iter() {
        if cell_has_content(cell) {
            min = min.min(r);
            max = max.max(r);
            found = true;
        }
    }
    found.then_some((min, max))
}

/// Pass 2: fill bucket array from cell positions.
fn fill_buckets(sheet: &Sheet, used_min: usize, rows_per_bucket: usize, buckets: &mut [u8]) {
    for (&(r, _), cell) in sheet.cells_iter() {
        if cell_has_content(cell) {
            let b = (r - used_min) / rows_per_bucket;
            if b < buckets.len() {
                buckets[b] = 140;
            }
        }
    }
}

/// Compute density buckets for the minimap.
///
/// Two passes over populated cells (sparse HashMap iteration):
/// 1. Find row extent (min/max row with content)
/// 2. Fill buckets based on row positions
///
/// Returns a MinimapCache with the computed data.
pub fn compute_buckets(sheet: &Sheet, height_px: usize, cells_rev: u64, sheet_id: u64) -> MinimapCache {
    if height_px == 0 {
        return MinimapCache {
            last_revision: cells_rev,
            last_sheet_id: sheet_id,
            ..Default::default()
        };
    }

    let Some((used_min, used_max)) = find_row_extent(sheet) else {
        return MinimapCache {
            height_px,
            last_revision: cells_rev,
            last_sheet_id: sheet_id,
            ..Default::default()
        };
    };

    let row_range = used_max - used_min + 1;
    let rows_per_bucket = row_range.div_ceil(height_px);
    let bucket_count = row_range.div_ceil(rows_per_bucket).min(height_px);

    let mut buckets = vec![0u8; bucket_count];
    fill_buckets(sheet, used_min, rows_per_bucket, &mut buckets);

    MinimapCache {
        used_min,
        used_max,
        buckets,
        height_px,
        last_revision: cells_rev,
        last_sheet_id: sheet_id,
    }
}
