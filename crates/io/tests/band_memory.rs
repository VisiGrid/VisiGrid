//! Memory of a banded load (ignored; run with --ignored --nocapture).
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering::Relaxed};

struct Counting;
static NOW: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 {
        let n = NOW.fetch_add(l.size(), Relaxed) + l.size();
        PEAK.fetch_max(n, Relaxed);
        unsafe { System.alloc(l) }
    }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) {
        NOW.fetch_sub(l.size(), Relaxed);
        unsafe { System.dealloc(p, l) }
    }
}
#[global_allocator]
static A: Counting = Counting;

#[test]
#[ignore]
fn banded_load_memory() {
    use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
    let rows = 300_000;
    let mut s = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    for r in 0..rows {
        for c in 0..20 {
            if c == 0 { s.set_value_deferred(r, 0, &format!("Item {r}")); }
            else if c % 3 == 0 { s.set_text(r, c, &format!("t{}", (r + c) % 97)); }
            else { s.set_value_deferred(r, c, &((r * 7 + c) % 1000).to_string()); }
        }
    }
    let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![s], 0);
    let mb = |n: usize| n / 1_048_576;
    eprintln!("direct workbook: {} MB", mb(NOW.load(Relaxed)));
    let layouts = vec![visigrid_io::json::SheetLayout::default()];
    let (manifest, bands) = visigrid_io::json::bands::export_banded(&wb, &layouts, 0).unwrap();
    drop(wb);
    let base = NOW.load(Relaxed);
    PEAK.store(base, Relaxed);
    let (mut loaded, _, _) = visigrid_io::json::import_any(&manifest).unwrap();
    for b in &bands {
        visigrid_io::json::bands::apply(&mut loaded, &b.data, None).unwrap();
        eprintln!("after band: now {} MB peak {} MB", mb(NOW.load(Relaxed) - base), mb(PEAK.load(Relaxed) - base));
    }
    visigrid_io::json::bands::finish(&mut loaded).unwrap();
    eprintln!("after finish: now {} MB peak {} MB", mb(NOW.load(Relaxed) - base), mb(PEAK.load(Relaxed) - base));
    let copy = loaded.clone();
    eprintln!("with a confirmed copy: now {} MB", mb(NOW.load(Relaxed) - base));
    drop(copy);
}
