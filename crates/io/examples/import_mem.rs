//! Where the memory of a CSV import goes: the file buffer, parsing, the cells
//! that stay, and what the allocator keeps on top (issue #18).
//!
//! ```text
//! cargo run --release -p visigrid-io --example import_mem -- data.csv
//! cargo run --release -p visigrid-io --example import_mem -- --generate 1000000 data.csv
//! ```
//!
//! Heap figures come from a counting allocator, so they are exact live bytes.
//! RSS is what the OS sees; the gap between the two is allocator overhead and
//! freed memory that was never returned.

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Write;
use std::mem::size_of;
use std::path::Path;
use std::sync::atomic::{AtomicUsize, Ordering};

use visigrid_engine::cell::{Cell, CellValue};

#[path = "../../engine/examples/support/rss.rs"]
mod rss;

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
static ALLOCS: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let p = unsafe { System.alloc(layout) };
        if !p.is_null() {
            let now = LIVE.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            PEAK.fetch_max(now, Ordering::Relaxed);
            ALLOCS.fetch_add(1, Ordering::Relaxed);
        }
        p
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        ALLOCS.fetch_sub(1, Ordering::Relaxed);
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

fn mb(bytes: usize) -> f64 {
    bytes as f64 / (1024.0 * 1024.0)
}

fn report(stage: &str) {
    println!(
        "{stage:<34} live {:>8.1} MB   peak {:>8.1} MB   allocs {:>10}   RSS {:>8.1} MB   HWM {:>8.1} MB",
        mb(LIVE.load(Ordering::Relaxed)),
        mb(PEAK.load(Ordering::Relaxed)),
        ALLOCS.load(Ordering::Relaxed),
        rss::rss_mb(),
        rss::peak_rss_mb(),
    );
}

/// Mixed columns like a typical export: id, name, quantity, price, note.
fn generate(rows: usize, path: &Path) {
    let mut out = std::io::BufWriter::new(std::fs::File::create(path).unwrap());
    let notes = ["", "ok", "backorder", "returned", "priority"];
    for r in 0..rows {
        writeln!(
            out,
            "{},item-{:06},{},{:.2},{}",
            100_000 + r,
            r % 250_000,
            r % 97,
            (r % 10_000) as f64 * 0.37,
            notes[r % notes.len()]
        )
        .unwrap();
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let path = match args.as_slice() {
        [flag, rows, path] if flag == "--generate" => {
            let p = Path::new(path).to_path_buf();
            generate(rows.parse().expect("row count"), &p);
            p
        }
        [path] => Path::new(path).to_path_buf(),
        _ => {
            eprintln!("usage: import_mem [--generate ROWS] FILE.csv");
            std::process::exit(2);
        }
    };

    println!("size_of: Cell {}  CellValue {}  (u32,u32)+Cell entry {}",
        size_of::<Cell>(), size_of::<CellValue>(), size_of::<((u32, u32), Cell)>());

    PEAK.store(LIVE.load(Ordering::Relaxed), Ordering::Relaxed);
    report("start");

    let content = visigrid_io::csv::read_file_as_utf8(&path).unwrap();
    report("file read into String");
    let file_bytes = content.len();

    let delimiter = visigrid_io::csv::sniff_delimiter(&content);
    let before_import = LIVE.load(Ordering::Relaxed);
    PEAK.store(before_import, Ordering::Relaxed);
    drop(content);
    // Re-run through the public entry point so the measured path is the real one.
    let sheet = visigrid_io::csv::import(&path).unwrap();
    let import_peak = PEAK.load(Ordering::Relaxed);
    report("after csv::import");
    let _ = delimiter;

    let cells = sheet.cells_iter().count();
    let live_sheet = LIVE.load(Ordering::Relaxed);

    // Break the sheet down: text heap vs everything else.
    let (mut text_cells, mut text_heap, mut numbers, mut formulas) = (0usize, 0usize, 0usize, 0usize);
    for (_, cell) in sheet.cells_iter() {
        match &cell.value {
            CellValue::Text(s) => { text_cells += 1; text_heap += s.capacity(); }
            CellValue::Number(_) => numbers += 1,
            CellValue::Formula { .. } => formulas += 1,
            CellValue::Empty => {}
        }
    }

    // hashbrown: buckets = next_power_of_two(ceil(n * 8 / 7)), one control byte each.
    let buckets = ((cells * 8 + 6) / 7).next_power_of_two();
    let entry = size_of::<((u32, u32), Cell)>();
    let table = buckets * entry + buckets + 16;

    println!();
    println!("file {:.1} MB, {} cells ({} text, {} number, {} formula)", mb(file_bytes), cells, text_cells, numbers, formulas);
    println!("import peak (heap) {:.1} MB = {:.0} B/cell", mb(import_peak), import_peak as f64 / cells as f64);
    println!("retained sheet     {:.1} MB = {:.0} B/cell", mb(live_sheet), live_sheet as f64 / cells as f64);
    println!("  hash table        {:.1} MB = {:.0} B/cell  ({} buckets x {} B entry, load {:.2})",
        mb(table), table as f64 / cells as f64, buckets, entry, cells as f64 / buckets as f64);
    println!("  text heap (exact) {:.1} MB = {:.0} B/cell, {:.1} B per text cell",
        mb(text_heap), text_heap as f64 / cells as f64, text_heap as f64 / text_cells.max(1) as f64);
    println!("  other             {:.1} MB", mb(live_sheet.saturating_sub(table + text_heap)));

    report("before trim");
    rss::trim_allocator();
    report("after trimming the allocator");

    // The desktop keeps a clone as base_workbook for replay.
    let before_clone = LIVE.load(Ordering::Relaxed);
    let copy = sheet.clone();
    let clone_cost = LIVE.load(Ordering::Relaxed) - before_clone;
    println!("\nsheet.clone() adds {:.1} MB ({:.0} B/cell)", mb(clone_cost), clone_cost as f64 / cells as f64);
    report("with clone");
    std::hint::black_box((&sheet, &copy));
}
