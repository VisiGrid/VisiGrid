//! Loading a large banded sheet: per-band cost must stay near the cost of
//! reading the cells, not grow with checks done once per cell.
//!
//! Run with `cargo test --release -p visigrid-io --test band_speed -- --ignored --nocapture`.
use std::time::Instant;
use visigrid_engine::cell::{CellFormat, NumberFormat};
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::{bands, import_any, SheetLayout};

/// `rows` x 20: numbers, text, a formula column, and formatted cells in a
/// handful of repeating formats, as an export from a ledger or report has.
fn wide(rows: usize) -> Workbook {
    let mut s = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    s.set_name("Data");
    let bold = CellFormat { bold: true, ..Default::default() };
    let money = CellFormat {
        number_format: NumberFormat::Number { decimals: 2, thousands: true, negative: Default::default() },
        ..Default::default()
    };
    for r in 0..rows {
        for c in 0..19 {
            let v = match c % 4 {
                0 => format!("{}", r * 7 + c),
                1 => format!("{}.{:02}", r % 10_000, (r + c) % 100),
                2 => format!("item {r}-{c}"),
                _ => format!("{}", (r as f64) * 0.25),
            };
            s.set_value_deferred(r, c, &v);
        }
        s.set_value_deferred(r, 19, &format!("=A{}+B{}", r + 1, r + 1));
        if r % 10 == 0 {
            s.set_format(r, 1, money.clone());
        }
        if r % 1000 == 0 {
            s.set_format(r, 2, bold.clone());
        }
    }
    let mut wb = Workbook::from_sheets(vec![s], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

#[test]
#[ignore = "timing; run in release with --ignored --nocapture"]
fn banded_load_time_per_band() {
    let rows: usize = std::env::var("BAND_SPEED_ROWS").ok().and_then(|v| v.parse().ok()).unwrap_or(100_000);
    let wb = wide(rows);
    let (manifest, out) = bands::export_banded(&wb, &[SheetLayout::default()], 0).unwrap();
    let (mut loaded, _, _) = import_any(&manifest).unwrap();
    let mut times = Vec::new();
    for b in &out {
        let t = Instant::now();
        bands::apply(&mut loaded, &b.data, Some(&b.reference.key)).unwrap();
        times.push(t.elapsed());
    }
    let t = Instant::now();
    bands::finish(&mut loaded).unwrap();
    let finish = t.elapsed();
    let total: std::time::Duration = times.iter().sum();
    eprintln!(
        "{rows} rows x 20: {} bands, apply total {:.2?} ({:.2?}/band, first {:.2?}, last {:.2?}), finish {:.2?}",
        out.len(),
        total,
        total / out.len() as u32,
        times[0],
        times[times.len() - 1],
        finish
    );
    assert_eq!(loaded.sheets()[0].cells_iter().count(), wb.sheets()[0].cells_iter().count());
}
