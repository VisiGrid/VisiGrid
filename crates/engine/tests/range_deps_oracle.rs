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
    match rng.below(6) {
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
