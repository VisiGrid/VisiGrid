//! Storage benchmark for #18: memory, lookup, structural edits, recalculation.
//! Public API only, so the same file measures any version of the engine.
//!
//! ```text
//! cargo run --release -p visigrid-engine --example storage_bench
//! ```

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;

#[path = "support/rss.rs"]
mod rss;

const ROWS: usize = 1_000_000;

/// Counts live heap bytes and the peak, so a structural edit's transient
/// memory is measured, not just its time.
struct Counting;
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let p = unsafe { System.alloc(layout) };
        if !p.is_null() {
            let now = LIVE.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            PEAK.fetch_max(now, Ordering::Relaxed);
        }
        p
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let p = unsafe { System.realloc(ptr, layout, new_size) };
        if !p.is_null() {
            if new_size >= layout.size() {
                let now = LIVE.fetch_add(new_size - layout.size(), Ordering::Relaxed) + new_size - layout.size();
                PEAK.fetch_max(now, Ordering::Relaxed);
            } else {
                LIVE.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
            }
        }
        p
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn mib(bytes: usize) -> f64 {
    bytes as f64 / (1024.0 * 1024.0)
}

/// Run `f`, reporting its time and how far the heap rose above where it
/// started while it ran.
fn timed_peak(label: &str, note: &str, f: impl FnOnce()) {
    let before = LIVE.load(Ordering::Relaxed);
    PEAK.store(before, Ordering::Relaxed);
    let t = Instant::now();
    f();
    let elapsed = ms(t);
    let peak = PEAK.load(Ordering::Relaxed) - before;
    println!("{label:<12} {elapsed:>9.1} ms   peak +{:.1} MiB over {:.1} MiB  ({note})", mib(peak), mib(before));
}

fn ms(t: Instant) -> f64 {
    t.elapsed().as_secs_f64() * 1000.0
}

/// 1M x 5, like import_mem's generated file: id, name (250k distinct),
/// quantity, price, note (4 distinct, every fifth empty).
fn build() -> Sheet {
    let notes = ["", "ok", "backorder", "returned", "priority"];
    let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    for r in 0..ROWS {
        sheet.set_value_deferred(r, 0, &format!("{}", 100_000 + r));
        sheet.set_value_deferred(r, 1, &format!("item-{:06}", r % 250_000));
        sheet.set_value_deferred(r, 2, &format!("{}", r % 97));
        sheet.set_value_deferred(r, 3, &format!("{:.2}", (r % 10_000) as f64 * 0.37));
        let note = notes[r % notes.len()];
        if !note.is_empty() {
            sheet.set_value_deferred(r, 4, note);
        }
    }
    sheet
}

fn main() {
    let base = rss::rss_mb();

    let t = Instant::now();
    let sheet = build();
    let build_ms = ms(t);
    let cells = sheet.cells_iter().count();
    let mem = rss::rss_mb() - base;
    println!("build        {build_ms:>9.1} ms   {cells} cells");
    println!("memory       {mem:>9.1} MB   {:.1} B/cell (RSS delta)", mem * 1048576.0 / cells as f64);

    // Point lookups: sequential down each column, then a scattered pattern.
    let t = Instant::now();
    let mut sum = 0.0f64;
    for c in 0..5 {
        for r in 0..ROWS {
            if let Some(cell) = sheet.get_cell_opt(r, c) {
                sum += cell.value().as_number();
            }
        }
    }
    println!("lookup seq   {:>9.1} ms   ({:.1} ns/lookup)", ms(t), ms(t) * 1e6 / (5 * ROWS) as f64);
    let t = Instant::now();
    let mut x: u64 = 88172645463325252;
    for _ in 0..(5 * ROWS) {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        let (r, c) = ((x as usize) % ROWS, (x as usize >> 32) % 5);
        if let Some(cell) = sheet.get_cell_opt(r, c) {
            sum += cell.value().as_number();
        }
    }
    println!("lookup rand  {:>9.1} ms   ({:.1} ns/lookup)", ms(t), ms(t) * 1e6 / (5 * ROWS) as f64);

    let t = Instant::now();
    let n = sheet.cells_iter().filter(|(_, c)| !c.value().is_empty()).count();
    println!("iterate      {:>9.1} ms   ({n} cells)", ms(t));

    // Structural edits at the top of the sheet: every cell below moves.
    // Measured on the sheet itself (no clone), so the peak is the edit's own.
    let mut s2 = sheet;
    timed_peak("insert row", "at row 0", || s2.insert_rows(0, 1));
    timed_peak("delete row", "at row 0", || s2.delete_rows(0, 1));
    timed_peak("insert col", "at col A", || s2.insert_cols(0, 1));
    timed_peak("delete col", "col A", || s2.delete_cols(0, 1));
    let sheet = s2;

    // A format-only edit over repeated text: 200k cells of column B (names).
    let mut s3 = sheet;
    timed_peak("bold 200k", "toggle_bold on repeated text", || {
        for r in 0..200_000 {
            s3.toggle_bold(r, 1);
        }
    });
    drop(s3);

    // Recalculation: 200k formulas over the numeric columns, then a full
    // ordered recompute, then SUM over whole columns.
    let mut calc = Sheet::new(SheetId(2), NUM_ROWS, NUM_COLS);
    for r in 0..200_000 {
        calc.set_value_deferred(r, 0, &format!("{}", r % 1000));
        calc.set_value_deferred(r, 1, &format!("=A{}*2+1", r + 1));
    }
    calc.set_value_deferred(0, 3, "=SUM(A1:A200000)");
    calc.set_value_deferred(1, 3, "=SUM(B1:B200000)");
    let mut wb = Workbook::from_sheets(vec![calc], 0);
    let t = Instant::now();
    wb.rebuild_dep_graph();
    println!("dep graph    {:>9.1} ms   (200k formulas)", ms(t));
    let t = Instant::now();
    wb.recompute_full_ordered();
    println!("recompute    {:>9.1} ms   (200k formulas + 2 SUMs)", ms(t));
    println!("check        {} / {}", wb.sheets()[0].get_display(1, 3), sum as u64 % 1000);
}
