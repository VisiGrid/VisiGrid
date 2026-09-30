//! #29: ranges are indexed instead of expanded into per-cell edges. The
//! oracle: after every incremental edit (set_cell_value_tracked: graph
//! update + dirty-set recalc), the sheet must display exactly what a fresh
//! rebuild_dep_graph + recompute_full_ordered on a copy displays.
//!
//! Formulas only read columns to their left, so the graph is acyclic and
//! the two paths have no reason to differ.

use visigrid_engine::structural::Axis;
use visigrid_engine::workbook::Workbook;

const ROWS: usize = 40;
const COLS: usize = 6;

/// xorshift64*: deterministic, no dependency.
struct Rng(u64);
impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 >> 12;
        self.0 ^= self.0 << 25;
        self.0 ^= self.0 >> 27;
        self.0.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }
    fn below(&mut self, n: usize) -> usize {
        (self.next() % n as u64) as usize
    }
}

fn col_letter(c: usize) -> char {
    (b'A' + c as u8) as char
}

/// Something to put at (row, col): a value, or a formula over columns < col.
fn content(rng: &mut Rng, row: usize, col: usize) -> String {
    if col == 0 || rng.below(4) == 0 {
        return match rng.below(5) {
            0 => String::new(),
            1 => "text".to_string(),
            2 => "=2*3".to_string(), // a formula with no references
            _ => format!("{}", rng.below(100)),
        };
    }
    let src = col_letter(rng.below(col));
    let r = row + 1;
    match rng.below(10) {
        // running total: every row reads all rows above it
        0 => format!("=SUM(${src}$1:{src}{r})"),
        // a rectangle over several columns to the left
        1 => {
            let (a, b) = (rng.below(ROWS) + 1, rng.below(ROWS) + 1);
            format!("=SUM(A{}:{}{})", a.min(b), col_letter(col - 1), a.max(b))
        }
        // a whole column
        2 => format!("=SUM({src}:{src})"),
        // single references
        3 => format!("={src}{}*2+{src}{}", rng.below(ROWS) + 1, rng.below(ROWS) + 1),
        4 => format!("=COUNT({src}1:{src}{})", rng.below(ROWS) + 1),
        // written bottom-up and right-to-left: A9:A2, C5:A1
        5 => {
            let (a, b) = (rng.below(ROWS) + 1, rng.below(ROWS) + 1);
            format!("=SUM({}{}:A{})", col_letter(col - 1), a.max(b), a.min(b))
        }
        // array arithmetic: the ranges sit inside operators, not directly in
        // a function's argument list, so their dependencies must still register
        6 => format!("=SUMPRODUCT(--({src}1:{src}{r}>50))"),
        7 => {
            let (a, b) = (rng.below(ROWS) + 1, rng.below(ROWS) + 1);
            let (a, b) = (a.min(b), a.max(b));
            format!("=SUM(({src}{a}:{src}{b}>20)*{src}{a}:{src}{b})")
        }
        // a wide range: one row, every column to the left (stored in a
        // row tree, not a column tree)
        8 => format!("=SUM(A{0}:{1}{0})", rng.below(ROWS) + 1, col_letter(col - 1)),
        _ => format!("=MAX({src}{}:{src}{})", r.min(ROWS), ROWS),
    }
}

fn snapshot(wb: &Workbook) -> Vec<String> {
    let sheet = wb.active_sheet();
    (0..ROWS)
        .flat_map(|r| (0..COLS).map(move |c| (r, c)))
        .map(|(r, c)| sheet.get_display(r, c))
        .collect()
}

fn oracle(wb: &Workbook) -> Vec<String> {
    let mut fresh = wb.clone();
    fresh.rebuild_dep_graph();
    fresh.recompute_full_ordered();
    snapshot(&fresh)
}

fn run(seed: u64, edits: usize) {
    let mut rng = Rng(seed);
    let mut wb = Workbook::new();
    // Fill column by column so every formula's inputs exist first...
    for c in 0..COLS {
        for r in 0..ROWS {
            let v = content(&mut rng, r, c);
            wb.set_cell_value_tracked(0, r, c, &v);
        }
    }
    assert_eq!(snapshot(&wb), oracle(&wb), "after fill (seed {seed})");
    // ...then edit anywhere, including turning values inside other
    // formulas' ranges into formulas and back.
    for i in 0..edits {
        let (r, c) = (rng.below(ROWS), rng.below(COLS));
        let v = content(&mut rng, r, c);
        wb.set_cell_value_tracked(0, r, c, &v);
        if rng.below(10) == 0 {
            let (r, c) = (rng.below(ROWS), rng.below(COLS));
            wb.clear_cell_tracked(0, r, c);
        }
        // Structural edits rebuild the graph; the edits after them then run
        // on a rebuilt graph. Formulas still only read columns to their left.
        if rng.below(15) == 0 {
            let axis = if rng.below(2) == 0 { Axis::Row } else { Axis::Col };
            let limit = if axis == Axis::Row { ROWS } else { COLS };
            let _ = wb.structural_edit(0, axis, rng.below(limit), 1, rng.below(2) == 0);
        }
        let expected = oracle(&wb);
        let got = snapshot(&wb);
        if got != expected {
            let diffs: Vec<String> = got
                .iter()
                .zip(&expected)
                .enumerate()
                .filter(|(_, (g, e))| g != e)
                .take(5)
                .map(|(i, (g, e))| {
                    let (r, c) = (i / COLS, i % COLS);
                    format!("{}{}: incremental {g:?}, full {e:?} ({:?})", col_letter(c), r + 1, wb.active_sheet().get_raw(r, c))
                })
                .collect();
            panic!("edit {i} (seed {seed}): {}", diffs.join("; "));
        }
    }
}

#[test]
fn incremental_recalc_matches_full_recompute_with_ranges() {
    for seed in 1..=12u64 {
        run(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 150);
    }
}

/// Longer soak: `cargo test --release --test range_deps_oracle -- --ignored`.
#[test]
#[ignore]
fn incremental_recalc_matches_full_recompute_with_ranges_long() {
    for seed in 1..=200u64 {
        run(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 400);
    }
}

/// A range written bottom-up (A5:A1) used to reach the graph reversed:
/// a panic in the range index, and on main a formula that never updated.
#[test]
fn reversed_ranges_are_tracked() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=1+1");
    wb.set_cell_value_tracked(0, 2, 0, "3");
    wb.set_cell_value_tracked(0, 0, 1, "=SUM(A5:A1)");
    wb.set_cell_value_tracked(0, 0, 2, "=SUM(B5:A1)");
    assert_eq!(wb.active_sheet().get_display(0, 1), "5");
    wb.set_cell_value_tracked(0, 2, 0, "13");
    assert_eq!(wb.active_sheet().get_display(0, 1), "15");
    assert_eq!(wb.active_sheet().get_display(0, 2), "30");
    assert_eq!(snapshot(&wb), oracle(&wb));
}

/// The inspector asks `has_cycle_in_upstream` on every render. A running
/// total whose first input sits in a cycle reports it at every row; one
/// beside it does not; and the walk is linear, not quadratic, up the column.
#[test]
fn a_cycle_upstream_is_found_through_ranges() {
    let rows = 5000;
    let mut wb = Workbook::new();
    // A1 and D1 read each other: a cycle.
    wb.set_cell_value_tracked(0, 0, 0, "=D1+1");
    wb.set_cell_value_tracked(0, 0, 3, "=A1+1");
    for r in 1..rows {
        wb.set_cell_value_tracked(0, r, 0, &format!("=A{}+1", r));
    }
    for r in 0..rows {
        // B: running total over the A formulas; C: over values in E.
        wb.set_cell_value_tracked(0, r, 1, &format!("=SUM($A$1:A{})", r + 1));
        wb.set_cell_value_tracked(0, r, 4, &format!("{}", r));
        wb.set_cell_value_tracked(0, r, 2, &format!("=SUM($E$1:E{})", r + 1));
    }
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet = wb.active_sheet().id;
    let t = std::time::Instant::now();
    assert!(wb.has_cycle_in_upstream(sheet, rows - 1, 1), "B reads A, whose top is in a cycle");
    assert!(!wb.has_cycle_in_upstream(sheet, rows - 1, 2), "C reads only values");
    // Base took ~0.5 s here and the first version of this change ~1.3 s;
    // a linear walk is a few milliseconds. Generous for a loaded machine.
    assert!(t.elapsed() < std::time::Duration::from_millis(500), "{:?}", t.elapsed());
}
