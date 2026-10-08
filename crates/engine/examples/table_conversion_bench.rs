//! Diagnostic conversion/history timings; setup is excluded. Run each size in
//! a separate process for independent peak RSS. This does not measure the GUI.
use std::time::Instant;
use visigrid_engine::{
    sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS},
    table::TableRange,
    workbook::Workbook,
};
#[path = "support/rss.rs"]
mod rss;

fn main() {
    let rows: usize = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "10000".into())
        .parse()
        .unwrap();
    assert!((2..NUM_ROWS - 1).contains(&rows));
    let setup = Instant::now();
    let mut data = Sheet::new_with_name(SheetId(1), NUM_ROWS, NUM_COLS, "Data");
    for col in 0..8 {
        data.set_value_deferred(0, col, &format!("Value{}", col + 1));
    }
    data.set_value_deferred(0, 8, "Double");
    data.set_value_deferred(0, 9, "PlusOne");
    for row in 1..=rows {
        for col in 0..8 {
            data.set_value_deferred(row, col, &((row + col) % 100).to_string());
        }
        data.set_value_deferred(row, 8, "=[@Value1]*2");
        data.set_value_deferred(row, 9, "=[@Value2]+1");
    }
    let summary = Sheet::new_with_name(SheetId(2), NUM_ROWS, NUM_COLS, "Summary");
    let mut wb = Workbook::from_sheets(vec![data, summary], 0);
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                end_row: rows,
                start_col: 0,
                end_col: 9,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let mut catalog = wb.saved_tables();
    catalog.version = 2;
    catalog.sheets[0].tables[0].columns[8].formula = Some("=[@Value1]*2".into());
    catalog.sheets[0].tables[0].columns[9].formula = Some("=[@Value2]+1".into());
    wb.restore_tables(catalog).unwrap();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Double])");
    let total = wb.active_sheet().get_computed_value(rows + 1, 9);
    let sum = wb.sheet(1).unwrap().get_computed_value(0, 0);
    rss::trim_allocator();
    println!("rows={rows} columns=10 calculated_columns=2 stored_cells={} setup_ms={:.1} baseline_rss_mib={:.1}", wb.sheets().iter().map(|s|s.cells_iter().count()).sum::<usize>(), setup.elapsed().as_secs_f64()*1000., rss::rss_mb());
    let start = Instant::now();
    let result = wb.remove_table(id);
    println!(
        "convert_ms={:.1} rss_mib={:.1} peak_rss_mib={:.1}",
        start.elapsed().as_secs_f64() * 1000.,
        rss::rss_mb(),
        rss::peak_rss_mb()
    );
    let commit = match result {
        Ok(commit) => commit,
        Err(error) => {
            assert!(wb.table(id).is_some());
            assert_eq!(wb.active_sheet().get_raw(rows, 8), "=[@Value1]*2");
            assert_eq!(wb.active_sheet().get_computed_value(rows + 1, 9), total);
            println!("refused_atomically={error}");
            return;
        }
    };
    assert!(wb.table(id).is_none());
    assert_eq!(wb.active_sheet().get_computed_value(rows + 1, 9), total);
    assert_eq!(wb.sheet(1).unwrap().get_computed_value(0, 0), sum);
    for row in 1..=rows {
        assert_eq!(
            wb.active_sheet().get_display(row, 8),
            ((row % 100) * 2).to_string()
        );
        assert_eq!(
            wb.active_sheet().get_display(row, 9),
            (((row + 1) % 100) + 1).to_string()
        );
    }
    let start = Instant::now();
    wb.apply_table_commit(&commit, true).unwrap();
    println!("undo_ms={:.1}", start.elapsed().as_secs_f64() * 1000.);
    assert_eq!(wb.active_sheet().get_raw(rows, 8), "=[@Value1]*2");
    let start = Instant::now();
    wb.apply_table_commit(&commit, false).unwrap();
    println!("redo_ms={:.1}", start.elapsed().as_secs_f64() * 1000.);
    assert!(wb.table(id).is_none());
    assert_eq!(wb.active_sheet().get_computed_value(rows + 1, 9), total);
    assert_eq!(wb.sheet(1).unwrap().get_computed_value(0, 0), sum);
    wb.set_cell_value_tracked(0, rows, 0, "1234");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_display(rows, 8), "2468");
    drop(wb);
    rss::trim_allocator();
    println!(
        "history_only_rss_mib={:.1} peak_rss_mib={:.1}",
        rss::rss_mb(),
        rss::peak_rss_mb()
    );
    std::hint::black_box(&commit);
    drop(commit);
    rss::trim_allocator();
    println!("after_history_drop_rss_mib={:.1}", rss::rss_mb());
}
