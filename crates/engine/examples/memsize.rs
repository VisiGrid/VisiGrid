//! What a cell costs in memory, and what a sheet of them costs.
//!
//! The grid is 1,048,576 x 16,384, so a large import is bounded by memory
//! rather than by the grid. Run this after touching `Cell`, `CellValue` or
//! `CellFormat`:
//!
//! ```text
//! cargo run --release -p visigrid-engine --example memsize
//! ```

use std::mem::size_of;
use visigrid_engine::cell::Cell;
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};

fn rss_mb() -> f64 {
    let status = std::fs::read_to_string("/proc/self/status").unwrap_or_default();
    status
        .lines()
        .find_map(|line| line.strip_prefix("VmRSS:"))
        .and_then(|v| v.trim().trim_end_matches(" kB").parse::<f64>().ok())
        .map(|kb| kb / 1024.0)
        .unwrap_or(0.0)
}

fn main() {
    println!("size_of::<Cell>() = {}", size_of::<Cell>());
    let base = rss_mb();
    let (rows, cols) = (200_000usize, 5usize);
    let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    for r in 0..rows {
        sheet.set_value(r, 0, "12345");
        sheet.set_value(r, 1, "item-name-here");
        sheet.set_value(r, 2, "42");
        sheet.set_value(r, 3, "3.75");
        sheet.set_value(r, 4, "note");
    }
    let used = rss_mb() - base;
    let cells = (rows * cols) as f64;
    println!(
        "{} cells: {used:.0} MB  =>  {:.0} bytes/cell",
        cells as u64,
        used * 1024.0 * 1024.0 / cells
    );
    println!("  a 1M x 10 sheet would be ~{:.1} GB", used / 1024.0 * (10_000_000.0 / cells));
    std::hint::black_box(&sheet);
}
