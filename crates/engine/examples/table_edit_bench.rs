//! Reproducible component timings for guarded Table editing (no GUI rendering).
//! Run with the shipping optimization profile:
//! cargo run --release -p visigrid-engine --example table_edit_bench -- 10000 100000
//! Optional TABLE_BENCH_SAMPLES (default 3) controls measured repetitions.
//! Each repetition starts from the same workbook; setup is excluded. Results
//! include median/min/max and assertions that mutations and replay are correct.

use std::{collections::BTreeMap, hint::black_box, time::Instant};
use visigrid_engine::{
    filter::{ColumnFilter, NormalizedFilterKey, SortDirection},
    sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS},
    table::TableRange,
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::Workbook,
};

#[path = "support/rss.rs"]
mod rss;

fn fixture(rows: usize, formulas: bool) -> Workbook {
    // Full worksheet dimensions matter: projections also represent blank rows.
    let mut sheet = Sheet::new_with_name(SheetId(1), NUM_ROWS, NUM_COLS, "Data");
    for (col, header) in ["Region", "Amount", "Result"].iter().enumerate() {
        sheet.set_value_deferred(0, col, header);
    }
    for row in 1..=rows {
        sheet.set_value_deferred(row, 0, if row % 2 == 0 { "West" } else { "East" });
        sheet.set_value_deferred(row, 1, &(row % 1000).to_string());
        sheet.set_value_deferred(
            row,
            2,
            &if formulas {
                format!("=B{}*2", row + 1)
            } else {
                ((row % 1000) * 2).to_string()
            },
        );
    }
    let summary = Sheet::new_with_name(SheetId(2), NUM_ROWS, NUM_COLS, "Summary");
    let mut wb = Workbook::from_sheets(vec![sheet, summary], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                end_row: rows,
                start_col: 0,
                end_col: 2,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let table = wb.table(id).unwrap().1;
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: table.columns[1].id,
        direction: SortDirection::Descending,
    });
    spec.filters.push(TableFilter {
        column: table.columns[0].id,
        criteria: ColumnFilter {
            selected: Some(
                [NormalizedFilterKey::Text("west".into())]
                    .into_iter()
                    .collect(),
            ),
            ..Default::default()
        },
    });
    wb.set_table_view_spec(SheetId(1), Some(spec)).unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Result])");
    assert!(wb.take_incremental_errors().is_empty());
    wb
}

fn measure<T>(
    samples: &mut BTreeMap<&'static str, Vec<f64>>,
    name: &'static str,
    f: impl FnOnce() -> T,
) -> T {
    let start = Instant::now();
    let result = black_box(f());
    samples
        .entry(name)
        .or_default()
        .push(start.elapsed().as_secs_f64() * 1000.0);
    result
}

fn run(rows: usize, formulas: bool, repeats: usize) {
    let start = Instant::now();
    let wb = fixture(rows, formulas);
    rss::trim_allocator();
    println!(
        "\nrows={rows} formulas={formulas} cells={} setup_ms={:.1} baseline_rss_mib={:.1}",
        wb.sheets()
            .iter()
            .map(|s| s.cells_iter().count())
            .sum::<usize>(),
        start.elapsed().as_secs_f64() * 1000.0,
        rss::rss_mb()
    );
    let original = wb.sheet(1).unwrap().get_computed_value(0, 0);
    let original_sum: f64 = wb.sheet(1).unwrap().get_display(0, 0).parse().unwrap();
    let original_total: f64 = wb.active_sheet().get_display(rows + 1, 2).parse().unwrap();
    let mut times = BTreeMap::new();
    for _ in 0..repeats {
        measure(&mut times, "01 build sorted+filtered view", || {
            let view = wb
                .active_sheet()
                .build_saved_table_view(NUM_ROWS)
                .unwrap()
                .unwrap();
            assert!(view.rows().is_data_row_visible(2));
            assert!(!view.rows().is_data_row_visible(1));
            black_box(view);
        });
        let mut candidate = measure(&mut times, "02 clone workbook", || wb.clone());
        measure(&mut times, "03 write + incremental recalc", || {
            let mut batch = candidate.batch_guard();
            batch.set_cell_value_tracked(0, 2, 1, "12345");
        });
        assert!(candidate.take_incremental_errors().is_empty());
        assert_eq!(candidate.active_sheet().get_raw(2, 1), "12345");
        assert_eq!(
            candidate.active_sheet().get_display(2, 2),
            if formulas { "24690" } else { "4" }
        );
        let delta = if formulas { 24686.0 } else { 0.0 };
        assert_eq!(
            candidate
                .sheet(1)
                .unwrap()
                .get_display(0, 0)
                .parse::<f64>()
                .unwrap(),
            original_sum + delta
        );
        assert_eq!(
            candidate
                .active_sheet()
                .get_display(rows + 1, 2)
                .parse::<f64>()
                .unwrap(),
            original_total + delta
        );
        measure(&mut times, "04 validate all saved views", || {
            for sheet in candidate.sheets() {
                black_box(sheet.build_saved_table_view(NUM_ROWS).unwrap());
            }
        });
        // This is automation/structural batch history, not the desktop's
        // cheaper target-only TableCellsCommit used for ordinary typing.
        let commit = measure(&mut times, "05 capture guarded batch", || {
            wb.capture_guarded_batch(&candidate).unwrap()
        });
        assert_eq!(commit.changed_cell_count(), 1);
        let undone = measure(&mut times, "06 guarded batch undo candidate", || {
            commit.candidate(&candidate, true).unwrap()
        });
        assert_eq!(undone.active_sheet().get_raw(2, 1), "2");
        assert_eq!(undone.sheet(1).unwrap().get_computed_value(0, 0), original);
        let redone = measure(&mut times, "07 guarded batch redo candidate", || {
            commit.candidate(&undone, false).unwrap()
        });
        assert_eq!(redone.active_sheet().get_raw(2, 1), "12345");
        assert_eq!(
            redone.sheet(1).unwrap().get_computed_value(0, 0),
            candidate.sheet(1).unwrap().get_computed_value(0, 0)
        );
        drop((undone, redone, commit));
        measure(&mut times, "08 publish monotonic snapshot", || {
            candidate.restore_snapshot_monotonic(&wb)
        });
        assert_eq!(candidate.active_sheet().get_raw(2, 1), "2");
        drop(candidate);
        assert_eq!(wb.active_sheet().get_raw(2, 1), "2");
        assert_eq!(wb.sheet(1).unwrap().get_computed_value(0, 0), original);
    }
    println!("component                                  median_ms     min_ms     max_ms");
    for (name, mut samples) in times {
        samples.sort_by(f64::total_cmp);
        let mid = samples.len() / 2;
        let median = if samples.len() % 2 == 0 {
            (samples[mid - 1] + samples[mid]) / 2.0
        } else {
            samples[mid]
        };
        println!(
            "{name:<42} {median:>9.2} {:>10.2} {:>10.2}",
            samples[0],
            samples[samples.len() - 1]
        );
    }
    println!("process_peak_rss_mib={:.1} (cumulative, includes setup and simultaneous replay candidates)", rss::peak_rss_mb());
}

fn main() {
    let sizes: Vec<usize> = std::env::args()
        .skip(1)
        .map(|s| s.parse().expect("row counts must be integers"))
        .collect();
    let sizes = if sizes.is_empty() {
        vec![10_000, 100_000]
    } else {
        sizes
    };
    let repeats: usize = std::env::var("TABLE_BENCH_SAMPLES")
        .unwrap_or_else(|_| "3".into())
        .parse()
        .expect("TABLE_BENCH_SAMPLES must be an integer");
    assert!((1..=100).contains(&repeats), "use 1..=100 samples");
    for rows in sizes {
        assert!(
            (2..NUM_ROWS - 1).contains(&rows),
            "leave room for the header and totals row"
        );
        for formulas in [false, true] {
            run(rows, formulas, repeats);
        }
    }
}
