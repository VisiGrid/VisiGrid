//! Deterministic instruction counts for the engine's hot paths.
//!
//! Run by `scripts/engine-counts.sh` under Valgrind's callgrind, which counts
//! CPU instructions instead of timing them. One run is enough: the count for
//! a scenario does not depend on the machine's load, so it can gate a pull
//! request where a millisecond figure could not (see docs/engine-counts.md).
//!
//! Each scenario builds its input first, then calls exactly one `measured_*`
//! function. Callgrind's `--toggle-collect='*measured_*'` counts only inside
//! those, so building a 20,000-row workbook costs nothing in the figure for
//! recalculating it. The functions are `#[inline(never)]` so that boundary
//! survives optimisation.
//!
//! Lives in `visigrid-io` rather than `visigrid-engine` because the JSON
//! round trip needs both, and io already depends on engine. Build with
//! `--no-default-features`: the canonical JSON reader is feature-free, and
//! the default set pulls in DuckDB for nothing.
//!
//! Without Valgrind this is still an ordinary binary: `cargo run --example
//! engine_counts -- full_recalc` runs the scenario and checks its result.

use std::hint::black_box;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;
use visigrid_io::json;

const SCENARIOS: &[&str] = &["parse_formulas", "full_recalc", "incremental_edit", "json_roundtrip"];

fn main() {
    let scenario = std::env::args().nth(1).unwrap_or_default();
    match scenario.as_str() {
        "parse_formulas" => parse_formulas(),
        "full_recalc" => full_recalc(),
        "incremental_edit" => incremental_edit(),
        "json_roundtrip" => json_roundtrip(),
        "list" => {
            for s in SCENARIOS {
                println!("{s}");
            }
        }
        other => {
            eprintln!("unknown scenario {other:?}; one of: {}", SCENARIOS.join(", "));
            std::process::exit(2);
        }
    }
}

/// A workbook whose first sheet holds `rows` rows of `2` in column A and a
/// formula over it in column B. The formula shape is the cheap one (one cell
/// read); range shapes are covered by `full_recalc`'s SUM.
fn workbook_with_formulas(rows: usize) -> Workbook {
    let mut wb = Workbook::new();
    let sheet = &mut wb.sheets_mut()[0];
    for r in 0..rows {
        sheet.set_value(r, 0, "2");
        sheet.set_value(r, 1, &format!("=A{}*2", r + 1));
    }
    wb
}

// ---------------------------------------------------------------------------
// parse_formulas: writing 1,000 formula cells, which parses each one.
// ---------------------------------------------------------------------------

fn parse_formulas() {
    let mut wb = Workbook::new();
    let formulas: Vec<String> = (0..1_000)
        .map(|r| format!("=IF(A{r1}>1, SUM(A1:A{r1})/{r1}, ROUND(A{r1}*1.5, 2))", r1 = r + 1))
        .collect();
    measured_parse_formulas(&mut wb, &formulas);
    let sheet = &wb.sheets()[0];
    assert!(sheet.get_raw(999, 1).starts_with('='), "last formula stored as a formula");
}

#[inline(never)]
fn measured_parse_formulas(wb: &mut Workbook, formulas: &[String]) {
    let sheet = &mut wb.sheets_mut()[0];
    for (r, f) in formulas.iter().enumerate() {
        sheet.set_value(r, 1, black_box(f));
    }
    black_box(sheet);
}

// ---------------------------------------------------------------------------
// full_recalc: dependency graph and full ordered recompute of 20,000
// formulas plus one SUM over the column.
// ---------------------------------------------------------------------------

fn full_recalc() {
    let rows = 20_000;
    let mut wb = workbook_with_formulas(rows);
    wb.sheets_mut()[0].set_value(0, 3, &format!("=SUM(A1:A{rows})"));
    measured_full_recalc(&mut wb);
    assert_eq!(wb.sheets()[0].get_display(rows - 1, 1), "4");
    assert_eq!(wb.sheets()[0].get_display(0, 3), format!("{}", rows * 2));
}

#[inline(never)]
fn measured_full_recalc(wb: &mut Workbook) {
    wb.rebuild_dep_graph();
    black_box(wb.recompute_full_ordered());
}

// ---------------------------------------------------------------------------
// incremental_edit: one edit to a cell with 1,000 direct dependents, after
// the workbook is built and calculated. This is the keystroke path.
// ---------------------------------------------------------------------------

fn incremental_edit() {
    let mut wb = Workbook::new();
    {
        let sheet = &mut wb.sheets_mut()[0];
        sheet.set_value(0, 0, "2");
        for r in 0..1_000 {
            sheet.set_value(r, 1, "=$A$1*2");
        }
    }
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    assert_eq!(wb.sheets()[0].get_display(999, 1), "4");

    measured_incremental_edit(&mut wb);
    assert_eq!(wb.sheets()[0].get_display(999, 1), "6", "dependents recalculated after the edit");
}

#[inline(never)]
fn measured_incremental_edit(wb: &mut Workbook) {
    black_box(wb.set_cell_value_tracked(0, 0, 0, black_box("3")));
}

// ---------------------------------------------------------------------------
// json_roundtrip: export a 5,000-row sheet to canonical visigrid-json and
// read it back. This is the save and open path for native, web and server.
// ---------------------------------------------------------------------------

fn json_roundtrip() {
    let rows = 5_000;
    let mut wb = workbook_with_formulas(rows);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet = &wb.sheets()[0];
    let back = measured_json_roundtrip(sheet);
    assert_eq!(back.get_raw(rows - 1, 0), "2", "literal survives the round trip");
    assert!(back.get_raw(rows - 1, 1).starts_with('='), "formula survives the round trip");
}

#[inline(never)]
fn measured_json_roundtrip(sheet: &Sheet) -> Sheet {
    let text = json::export_full(black_box(sheet)).expect("export");
    json::import_full(black_box(&text)).expect("import")
}
