//! Timings for operations whose cost could scale with the grid rather than
//! with the data, run at the bottom of a 1,048,576-row sheet.
//!
//! Ignored by default: this is a measurement, not an assertion about a
//! developer machine's speed. Run with
//! `cargo test -p visigrid-engine --release --test grid_scale_bench -- --ignored --nocapture`.

use std::time::Instant;
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;

/// A sheet holding `rows` rows of data that ends at the last row of the grid,
/// so every "walk to the edge" path has the longest possible distance to cover.
fn sheet_with_data_at_the_bottom(rows: usize) -> Sheet {
    let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    let start = NUM_ROWS - rows;
    for r in start..NUM_ROWS {
        sheet.set_value(r, 0, &format!("{r}"));
        sheet.set_value(r, 1, "text");
        sheet.set_value(r, 2, "3.5");
    }
    sheet
}

fn time<T>(label: &str, f: impl FnOnce() -> T) -> T {
    let start = Instant::now();
    let out = f();
    println!("{label}: {:?}", start.elapsed());
    out
}

#[test]
#[ignore]
fn structural_edits_at_the_bottom_of_the_grid() {
    let rows = 50_000;
    let mut sheet = time("build 50k rows at the grid floor", || {
        sheet_with_data_at_the_bottom(rows)
    });
    println!("  cells: {}", sheet.cells_iter().count());

    time("data_bounds", || sheet.data_bounds());
    time("data_extent", || sheet.data_extent());
    time("insert_rows(0, 1)", || sheet.insert_rows(0, 1));
    time("delete_rows(0, 1)", || sheet.delete_rows(0, 1));
    time("insert_cols(0, 1)", || sheet.insert_cols(0, 1));
    time("delete_cols(0, 1)", || sheet.delete_cols(0, 1));
    time("occupied_cells_in_rows(last 10)", || {
        sheet.occupied_cells_in_rows(NUM_ROWS - 10, 10)
    });
    time("occupied_cells_in_cols(col 0)", || {
        sheet.occupied_cells_in_cols(0, 1)
    });
}

#[test]
#[ignore]
fn recompute_with_formulas_at_the_bottom() {
    let mut wb = Workbook::new();
    let rows = 20_000;
    let start = NUM_ROWS - rows;
    time("write 20k formulas at the grid floor", || {
        for r in start..NUM_ROWS {
            wb.sheets_mut()[0].set_value(r, 0, "2");
            wb.sheets_mut()[0].set_value(r, 1, &format!("=A{}*2", r + 1));
        }
    });
    wb.rebuild_dep_graph();
    time("recompute_full_ordered", || wb.recompute_full_ordered());
    let last = wb.sheets()[0].get_display(NUM_ROWS - 1, 1);
    assert_eq!(last, "4", "formula at the last row computes");

    // An explicit full-column range now spans 1,048,576 rows. (The `A:A`
    // shorthand is not accepted in a cell formula — it parses in the CLI's
    // --calc only — and that was true before the grid grew.)
    wb.sheets_mut()[0].set_value(0, 3, &format!("=SUM(A1:A{NUM_ROWS})"));
    wb.rebuild_dep_graph();
    time("SUM over the full column", || wb.recompute_full_ordered());
    let total = wb.sheets()[0].get_display(0, 3);
    println!("  SUM(A1:A{NUM_ROWS}) = {total}");
    assert_eq!(total, format!("{}", rows * 2), "every populated row counted");
}
