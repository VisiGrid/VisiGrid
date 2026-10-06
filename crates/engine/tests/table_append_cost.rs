use std::{hint::black_box, time::Instant};
use visigrid_engine::{table::TableRange, workbook::Workbook};

#[test]
#[ignore = "manual timing probe; not a wall-clock CI assertion"]
fn measure_guarded_append_and_cycle_scans() {
    for rows in [10_000, 100_000] {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "Amount");
        wb.set_cell_value_tracked(0, 1, 0, "1");
        let id = wb
            .create_table(
                wb.active_sheet_id(),
                TableRange {
                    start_row: 0,
                    start_col: 0,
                    end_row: 1,
                    end_col: 0,
                },
                "Sales",
            )
            .unwrap()
            .table_id();
        wb.set_table_totals_visible(id, true, Default::default())
            .unwrap();
        let other = wb.add_sheet_named("Large").unwrap();
        for row in 0..rows {
            wb.sheet_mut(other).unwrap().set_value_deferred(row, 0, "1");
            wb.sheet_mut(other)
                .unwrap()
                .set_value_deferred(row, 1, &format!("=A{}+1", row + 1));
        }
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        let mut append_ms = Vec::new();
        let mut cycle_ms = Vec::new();
        let mut uncached_cycle_ms = Vec::new();
        for _ in 0..3 {
            let t = Instant::now();
            black_box(wb.dep_graph().find_cycle_members());
            black_box(wb.dep_graph().find_cycle_members());
            cycle_ms.push(t.elapsed().as_secs_f64() * 1000.0);
            // Invalidate only the derived proof, leaving the graph unchanged:
            // this absent cell has no dependencies. Cloning is outside timing.
            let mut fresh_a = wb.dep_graph().clone();
            fresh_a.clear_cell(visigrid_engine::cell_id::CellId::new(
                wb.active_sheet_id(),
                1000,
                100,
            ));
            let fresh_b = fresh_a.clone();
            let t = Instant::now();
            black_box(fresh_a.find_cycle_members());
            black_box(fresh_b.find_cycle_members());
            uncached_cycle_ms.push(t.elapsed().as_secs_f64() * 1000.0);
            drop((fresh_a, fresh_b));
            let mut candidate = wb.clone();
            let t = Instant::now();
            black_box(candidate.append_table_rows(id, 1, &[]).unwrap());
            append_ms.push(t.elapsed().as_secs_f64() * 1000.0);
        }
        println!("{rows} formulas: append_ms={append_ms:?}; cached_cycle_pair_ms={cycle_ms:?}; uncached_cycle_pair_ms={uncached_cycle_ms:?}");
    }
}
