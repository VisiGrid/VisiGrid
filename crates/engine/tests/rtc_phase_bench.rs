//! RTC spec step 1: is the critical path *building* the workbook or *recomputing* it?
//!
//! Not run by default. Invoke explicitly:
//!
//!   cargo test -p visigrid-engine --release --test rtc_phase_bench -- --ignored --nocapture
//!
//! Why this exists, from planning/visigrid/realtime-collaboration-architecture.md:
//! the WASM boundary is stateless, so `recompute(input)` reconstructs the whole
//! workbook from scratch on every call. Two different native-lane projects
//! could fix the per-keystroke cost — (A) hold a resident workbook across the
//! boundary, (B) recalculate incrementally inside the engine — and a figure
//! that times the exported `recompute` end to end cannot tell them apart. The
//! browser-measurement note says so outright: the wasm API does construct +
//! dep_graph + recompute in ONE call, so a browser figure never maps to a
//! single phase, and separating them from outside by varying the
//! literal-to-formula mix rests on construction scaling with cell count, which
//! is false.
//!
//! So the separation has to happen where the phases are: here, in-process,
//! against the same engine calls `build_workbook` makes.
//!
//! ## Method
//!
//! The measurement note's rule is that absolute milliseconds are meaningful
//! for the machine and the minute they were taken on and nothing else — a
//! figure captured one day cannot legitimately be compared against one
//! captured the next, and best-of-N tightens the estimate of a contaminated
//! state rather than escaping it.
//!
//! Its answer there was to interleave the things being compared in one run.
//! That cannot be done literally here: these phases are ordered by dependency,
//! and recompute cannot run before the writes that give it something to
//! recompute. So the discipline is applied one level up:
//!
//!   - every phase is timed inside the SAME round, on the same workbook;
//!   - the decision quantity is the per-round RATIO recompute/build, and the
//!     reported figure is the median of those ratios — drift that dirties a
//!     round inflates numerator and denominator together and cancels;
//!   - fixtures are visited in a rotating order across rounds, so no size
//!     systematically inherits a bigger one's heap;
//!   - a discarded warm-up round precedes the timed ones.
//!
//! Absolute medians are printed too, because "which dominates" is not the only
//! question — 3 ms and 300 ms mean different things for a keystroke — but they
//! are machine-and-minute figures and are labelled as such.

use std::time::{Duration, Instant};
use visigrid_engine::workbook::Workbook;

/// Formula shapes, from the browser-measurement note: formula EXPENSE, not
/// cell count, turned out to be the variable that matters — at identical cell
/// counts the range shape cost roughly 13x more.
#[derive(Clone, Copy, PartialEq)]
enum Shape {
    /// One cell read per formula.
    Cheap,
    /// 200 cell reads per formula.
    Range,
}

impl Shape {
    fn label(self) -> &'static str {
        match self {
            Shape::Cheap => "cheap",
            Shape::Range => "range",
        }
    }
}

/// Column letter for a 0-based index. Only ever called with small indices —
/// 200k formula cells laid out 50k to a column-pair needs eight columns.
fn col_name(mut col: usize) -> String {
    let mut out = Vec::new();
    loop {
        out.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).unwrap()
}

/// The engine grid is 65,536 rows, so 200k formulas cannot be one column.
/// Wrap into column-pairs — literal in the left, formula in the right — which
/// keeps every formula's dependency in its own row and well inside 256 columns.
const ROWS_PER_PAIR: usize = 50_000;

/// (row, col, raw) triples, in the same form `build_workbook` receives them.
fn fixture(n: usize, shape: Shape) -> Vec<(usize, usize, String)> {
    let mut cells = Vec::with_capacity(n * 2);
    for i in 0..n {
        let pair = i / ROWS_PER_PAIR;
        let row = i % ROWS_PER_PAIR;
        let lit_col = pair * 2;
        let formula_col = lit_col + 1;

        cells.push((row, lit_col, format!("{}", i + 1)));
        let formula = match shape {
            Shape::Cheap => format!("={}{}*2", col_name(lit_col), row + 1),
            // Constant 200 reads regardless of position, as the note's fixture had.
            Shape::Range => format!("=SUM({0}1:{0}200)", col_name(lit_col)),
        };
        cells.push((row, formula_col, formula));
    }
    cells
}

struct Phases {
    /// Writing every cell onto the sheet (`set_value_deferred`), the bulk of
    /// what a resident workbook would stop repeating.
    writes: Duration,
    /// `rebuild_dep_graph` — also paid per call today, also eliminated by a
    /// resident workbook.
    dep_graph: Duration,
    /// First `recompute_full_ordered`: every formula evaluated for the first
    /// time, from deferred values.
    first_recompute: Duration,
    /// Second `recompute_full_ordered` on the already-built workbook. This is
    /// the steady-state figure — what a resident workbook would still pay per
    /// keystroke, and the only phase incremental recalc attacks.
    warm_recompute: Duration,
}

impl Phases {
    /// What a resident workbook removes: everything before the recompute.
    fn build(&self) -> Duration {
        self.writes + self.dep_graph
    }
}

fn time_phases(cells: &[(usize, usize, String)]) -> Phases {
    let mut wb = Workbook::new();

    // Mirror build_workbook: widen before writing, so whole-column references
    // cover the data rather than the 65536x256 default.
    let (needed_rows, needed_cols) = cells
        .iter()
        .fold((0, 0), |(r, c), (row, col, _)| (r.max(row + 1), c.max(col + 1)));
    {
        let sheet = &mut wb.sheets_mut()[0];
        sheet.rows = sheet.rows.max(needed_rows);
        sheet.cols = sheet.cols.max(needed_cols);
    }

    let t = Instant::now();
    {
        let sheet = &mut wb.sheets_mut()[0];
        for (row, col, raw) in cells {
            sheet.set_value_deferred(*row, *col, raw);
        }
    }
    let writes = t.elapsed();

    let t = Instant::now();
    wb.rebuild_dep_graph();
    let dep_graph = t.elapsed();

    let t = Instant::now();
    wb.recompute_full_ordered();
    let first_recompute = t.elapsed();

    let t = Instant::now();
    wb.recompute_full_ordered();
    let warm_recompute = t.elapsed();

    Phases { writes, dep_graph, first_recompute, warm_recompute }
}

fn median(mut xs: Vec<f64>) -> f64 {
    xs.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let n = xs.len();
    if n % 2 == 1 {
        xs[n / 2]
    } else {
        (xs[n / 2 - 1] + xs[n / 2]) / 2.0
    }
}

fn ms(d: Duration) -> f64 {
    d.as_secs_f64() * 1000.0
}

#[test]
#[ignore = "measurement, not an assertion — run explicitly with --ignored --nocapture"]
fn build_versus_recompute() {
    // Overridable so a pilot run can be cheap: RTC_BENCH_SIZES=10000 etc.
    let sizes: Vec<usize> = std::env::var("RTC_BENCH_SIZES")
        .ok()
        .map(|s| s.split(',').filter_map(|x| x.trim().parse().ok()).collect())
        .unwrap_or_else(|| vec![10_000, 50_000, 200_000]);
    let shapes: Vec<Shape> = match std::env::var("RTC_BENCH_SHAPES").ok().as_deref() {
        Some("cheap") => vec![Shape::Cheap],
        Some("range") => vec![Shape::Range],
        _ => vec![Shape::Cheap, Shape::Range],
    };
    let rounds: usize = std::env::var("RTC_BENCH_ROUNDS")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(5);

    let mut cases: Vec<(usize, Shape)> = Vec::new();
    for &n in &sizes {
        for &shape in &shapes {
            cases.push((n, shape));
        }
    }

    // Fixtures are built once and reused: generating them is string formatting,
    // not engine work, and timing it here would be timing `format!`.
    let fixtures: Vec<Vec<(usize, usize, String)>> =
        cases.iter().map(|(n, shape)| fixture(*n, *shape)).collect();

    println!("\nrounds: {rounds} (plus one discarded warm-up)");
    println!("engine: visigrid-engine (native, release)\n");

    // Warm-up, discarded. Allocator and caches settle here rather than inside
    // the first timed round.
    for f in &fixtures {
        let _ = time_phases(f);
    }

    let mut samples: Vec<Vec<Phases>> = (0..cases.len()).map(|_| Vec::new()).collect();
    for round in 0..rounds {
        // Rotate the visiting order so no case always follows the same
        // predecessor — the heap it inherits differs from round to round.
        for offset in 0..cases.len() {
            let i = (offset + round) % cases.len();
            samples[i].push(time_phases(&fixtures[i]));
        }
    }

    println!(
        "{:>9}  {:>6}  {:>9} {:>9} {:>9} {:>9}   {:>11}  {:>11}",
        "formulas", "shape", "writes", "depgraph", "recalc-1", "recalc-N", "build:recalc", "recalc/total"
    );
    println!("{}", "-".repeat(96));

    for (i, (n, shape)) in cases.iter().enumerate() {
        let s = &samples[i];
        let writes = median(s.iter().map(|p| ms(p.writes)).collect());
        let dep = median(s.iter().map(|p| ms(p.dep_graph)).collect());
        let first = median(s.iter().map(|p| ms(p.first_recompute)).collect());
        let warm = median(s.iter().map(|p| ms(p.warm_recompute)).collect());

        // The decision quantity, computed per round and then medianed, so a
        // round that ran dirty inflates both halves and cancels.
        let ratio = median(
            s.iter()
                .map(|p| ms(p.build()) / ms(p.warm_recompute))
                .collect(),
        );
        let share = median(
            s.iter()
                .map(|p| ms(p.warm_recompute) / (ms(p.build()) + ms(p.warm_recompute)))
                .collect(),
        );

        println!(
            "{:>9}  {:>6}  {:>8.1}ms {:>8.1}ms {:>8.1}ms {:>8.1}ms   {:>10.2}x  {:>10.0}%",
            n,
            shape.label(),
            writes,
            dep,
            first,
            warm,
            ratio,
            share * 100.0
        );
    }

    println!(
        "\nbuild:recalc = (writes + depgraph) / recalc-N, median of per-round ratios.\n\
         >1 means a resident workbook (project A) is the larger win; <1 means\n\
         incremental recalculation (project B) is.\n\
         recalc/total = the share of a stateless per-keystroke call that survives\n\
         project A — i.e. what project B would then be attacking.\n\
         Absolute milliseconds are native, release, and true for this machine\n\
         and this minute only; the ratios are the portable part."
    );
}


// =============================================================================
// Follow-on: what a resident workbook would actually pay per keystroke.
//
// The first test measures the phases of today's stateless call. It answers
// "which phase dominates" but not "is a resident workbook enough", because
// under a resident workbook a keystroke would not call recompute_full_ordered
// at all — it would go through the tracked-write path the desktop app already
// uses: set_cell_value_tracked -> note_cell_changed -> recalc_dirty_set, which
// BFSes the dirty subgraph and evaluates only that.
//
// So this times exactly that call, on an already-built workbook, and compares
// it against the full recompute of the same fixture.
//
// The case that matters most is the narrow one, and it is what this harness
// was written to catch. `recalc_dirty_set` used to call
// `topo_order_all_formulas()` — a full topological sort of EVERY formula in
// the workbook — on every edit, however small the change, which cost 78 ms for
// a ONE-dependent edit on 200k formulas. It now orders only the dirty subgraph
// (`dep_graph::topo_order_subset`), and the same edit is ~0.002 ms and flat
// across sizes. (It regressed to 42 ms at 200k when spill settlement began
// scanning every formula for INDIRECT/OFFSET readers on every call; that scan
// is now lazy, done only when a round actually places an array.) A one-dependent edit is the measurement that keeps it honest:
// the dirty set is 1, so anything the edit costs beyond trivial is ordering
// work that should not be happening.
//
// (An edit with NO dependents cannot isolate it — `recalc_dirty_set` returns
// early on an empty dirty set, before the sort.)

/// Build a workbook the same way the first test does, without timing it.
fn built(cells: &[(usize, usize, String)]) -> Workbook {
    let mut wb = Workbook::new();
    let (needed_rows, needed_cols) = cells
        .iter()
        .fold((0, 0), |(r, c), (row, col, _)| (r.max(row + 1), c.max(col + 1)));
    {
        let sheet = &mut wb.sheets_mut()[0];
        sheet.rows = sheet.rows.max(needed_rows);
        sheet.cols = sheet.cols.max(needed_cols);
        for (row, col, raw) in cells {
            sheet.set_value_deferred(*row, *col, raw);
        }
    }
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

#[test]
#[ignore = "measurement, not an assertion — run explicitly with --ignored --nocapture"]
fn single_cell_edit_on_a_resident_workbook() {
    let sizes: Vec<usize> = std::env::var("RTC_BENCH_SIZES")
        .ok()
        .map(|s| s.split(',').filter_map(|x| x.trim().parse().ok()).collect())
        .unwrap_or_else(|| vec![10_000, 50_000, 200_000]);
    let shapes: Vec<Shape> = match std::env::var("RTC_BENCH_SHAPES").ok().as_deref() {
        Some("cheap") => vec![Shape::Cheap],
        Some("range") => vec![Shape::Range],
        _ => vec![Shape::Cheap, Shape::Range],
    };
    let edits: usize = std::env::var("RTC_BENCH_EDITS")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(11);

    println!("\nedits timed per fixture: {edits} (plus one discarded)");
    println!("engine: visigrid-engine (native, release)\n");
    println!(
        "{:>9}  {:>6}  {:>10}  {:>12}  {:>12}   {:>9}",
        "formulas", "shape", "dependents", "edit", "full recalc", "edit/full"
    );
    println!("{}", "-".repeat(76));

    for &n in &sizes {
        for &shape in &shapes {
            let cells = fixture(n, shape);
            let mut wb = built(&cells);

            // Edit the literal at (0, 0). Under the cheap shape exactly one
            // formula reads it; under the range shape every formula in the
            // first column-pair does, because they all read A1:A200.
            let dependents = wb.dep_graph().dependent_count(
                visigrid_engine::cell_id::CellId::new(wb.sheets()[0].id, 0, 0),
            );

            // Prove the edit actually recalculates before timing it. A silent
            // no-op — auto_recalc off, a write that does not mark anything
            // dirty — would otherwise be reported as a very fast keystroke.
            let formula_cell_col = 1;
            let before = wb.sheets()[0].get_display(0, formula_cell_col);
            wb.set_cell_value_tracked(0, 0, 0, "7777");
            let after = wb.sheets()[0].get_display(0, formula_cell_col);
            assert_ne!(
                before, after,
                "the tracked write did not recalculate its dependent — timing it would be meaningless"
            );

            let mut samples = Vec::with_capacity(edits);
            for i in 0..=edits {
                // Alternate the value so no write can be skipped as a no-op.
                let v = format!("{}", 1000 + i);
                let t = Instant::now();
                wb.set_cell_value_tracked(0, 0, 0, &v);
                let dt = t.elapsed();
                if i > 0 {
                    samples.push(ms(dt));
                }
            }
            let edit_ms = median(samples);

            // Same workbook, for a like-for-like comparison in the same round.
            let t = Instant::now();
            wb.recompute_full_ordered();
            let full_ms = ms(t.elapsed());

            println!(
                "{:>9}  {:>6}  {:>10}  {:>10.3}ms  {:>10.1}ms   {:>8.1}%",
                n,
                shape.label(),
                dependents,
                edit_ms,
                full_ms,
                edit_ms / full_ms * 100.0
            );
        }
    }

    println!(
        "\nedit = set_cell_value_tracked on one literal, median of per-edit times.\n\
         This is the per-keystroke cost a resident workbook would pay; the phases\n\
         in the other test become per-LOAD costs once the workbook stops being\n\
         rebuilt. `dependents` is the size of the dirty set the edit triggers.\n\
         A narrow edit should be flat across sizes. When it instead tracks the\n\
         workbook, the ordering is being done over every formula rather than\n\
         over the dirty set (the defect fixed in dep_graph::topo_order_subset)."
    );
}

// #88: volatile formulas (NOW, TODAY, RAND, RANDBETWEEN, INDIRECT, OFFSET) are
// recalculated on every edit, as in Excel. A workbook without any must keep
// the one-dependent edit flat (the first test above); this one shows what a
// workbook WITH them pays: every edit re-evaluates each volatile cell and
// whatever reads it, however unrelated the edit.
#[test]
#[ignore = "measurement, not an assertion — run explicitly with --ignored --nocapture"]
fn single_cell_edit_with_volatile_formulas() {
    let n = 200_000;
    let mut cells = fixture(n, Shape::Cheap);
    // Far right of the fixture: one INDIRECT with a reader, and 100 NOWs.
    let col = 2 * n.div_ceil(ROWS_PER_PAIR) + 1;
    cells.push((0, col, "=INDIRECT(\"A2\")".into()));
    cells.push((1, col, format!("={}1+1", col_name(col))));
    for r in 0..100 {
        cells.push((r + 2, col, "=NOW()".into()));
    }
    let mut wb = built(&cells);
    assert_eq!(wb.volatile_cell_count(), 101);
    let mut samples = Vec::new();
    for i in 0..=11 {
        let v = format!("{}", 1000 + i);
        let t = Instant::now();
        wb.set_cell_value_tracked(0, 0, 0, &v);
        if i > 0 {
            samples.push(ms(t.elapsed()));
        }
    }
    println!(
        "\nformulas={n} cheap + 1 INDIRECT (+1 reader) + 100 NOW: one-dependent edit median {:.3}ms",
        median(samples)
    );
}
