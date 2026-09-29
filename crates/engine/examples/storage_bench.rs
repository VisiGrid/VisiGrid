//! Storage benchmark for #18: memory, lookup, structural edits, recalculation.
//! Public API only, so the same file measures any version of the engine.
//!
//! ```text
//! cargo run --release -p visigrid-engine --example storage_bench
//! ```

use std::time::Instant;

use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;

#[path = "support/rss.rs"]
mod rss;

const ROWS: usize = 1_000_000;

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
    let mut s2 = sheet.clone();
    let t = Instant::now();
    s2.insert_rows(0, 1);
    println!("insert row   {:>9.1} ms   (at row 0)", ms(t));
    let t = Instant::now();
    s2.delete_rows(0, 1);
    println!("delete row   {:>9.1} ms   (at row 0)", ms(t));
    let t = Instant::now();
    s2.insert_cols(0, 1);
    println!("insert col   {:>9.1} ms   (at col A)", ms(t));
    drop(s2);

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
