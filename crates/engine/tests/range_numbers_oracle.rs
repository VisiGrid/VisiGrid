//! #18: SUM-like functions read ranges through Sheet::numbers_in, which
//! visits only stored cells. The oracle is the loop it replaced: for every
//! cell in the rectangle, get_text then parse::<f64>(), in row-major order.
//! Integers stay below 2^63, where the old loop saturated (tested separately).

use visigrid_engine::formula::eval::CellLookup;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;

const ROWS: usize = 60;
const COLS: usize = 5;

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

fn content(rng: &mut Rng) -> String {
    match rng.below(24) {
        16 => "-0".to_string(),         // negative zero, stored as a number
        17 => "=0*-1".to_string(),      // negative zero as a formula result
        18 => " 5".to_string(),         // text that does not parse
        19 => "+5".to_string(),
        20 => "inf".to_string(),
        21 => "1e3".to_string(),
        22 => "=\"-0\"".to_string(),  // text "-0" from a formula
        23 => "=-2^70".to_string(),     // large, non-integral path in to_text
        0 | 1 => String::new(),
        2 => format!("{}", rng.below(1000)),
        3 => format!("{}.{}", rng.below(100), rng.below(1000)),
        4 => format!("-{}", rng.below(50)),
        5 => "text".to_string(),
        6 => "42".to_string(),
        7 => "=1+2".to_string(),
        8 => "=\"5\"".to_string(),    // text result that parses
        9 => "=\"x\"".to_string(),    // text result that doesn't
        10 => "=1>0".to_string(),     // boolean result
        11 => "=1/0".to_string(),     // error result
        12 => format!("=SEQUENCE({})", rng.below(4) + 1), // spills down
        13 => "1e14".to_string(),
        14 => "=A1*0.5".to_string(),
        _ => "0".to_string(),
    }
}

fn slow(sheet: &Sheet, r0: usize, c0: usize, r1: usize, c1: usize) -> Vec<f64> {
    let mut out = Vec::new();
    for r in r0..=r1 {
        for c in c0..=c1 {
            if let Ok(n) = CellLookup::get_text(sheet, r, c).parse::<f64>() {
                out.push(n);
            }
        }
    }
    out
}

fn same(a: &[f64], b: &[f64]) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(x, y)| x.to_bits() == y.to_bits() || (x.is_nan() && y.is_nan()))
}

#[test]
fn numbers_in_matches_the_per_cell_loop() {
    for seed in 1..=40u64 {
        let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15));
        let mut wb = Workbook::new();
        for r in 0..ROWS {
            for c in 0..COLS {
                let v = content(&mut rng);
                wb.set_cell_value_tracked(0, r, c, &v);
            }
        }
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        let sheet = &wb.sheets()[0];
        for _ in 0..200 {
            let (a, b) = (rng.below(ROWS + 10), rng.below(ROWS + 10));
            let (c, d) = (rng.below(COLS + 2), rng.below(COLS + 2));
            let (r0, r1, c0, c1) = (a.min(b), a.max(b), c.min(d), c.max(d));
            let mut fast = Vec::new();
            sheet.numbers_in(r0, c0, r1, c1, &mut fast);
            let expected = slow(sheet, r0, c0, r1, c1);
            assert!(same(&fast, &expected), "seed {seed} rect ({r0},{c0})-({r1},{c1}):\n fast {fast:?}\n slow {expected:?}");
        }
    }
}

/// Cross-sheet reads go through WorkbookLookup::numbers_in_range, which
/// must match the per-cell loop over get_text_sheet, and a missing sheet
/// contributes nothing, as its "#REF!" text did.
#[test]
fn cross_sheet_numbers_match_the_per_cell_loop() {
    use visigrid_engine::sheet::SheetRef;
    let mut rng = Rng(7);
    let mut wb = Workbook::new();
    wb.add_sheet();
    for r in 0..ROWS {
        for c in 0..COLS {
            let v = content(&mut rng);
            wb.set_cell_value_tracked(1, r, c, &v);
        }
    }
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let other = wb.sheets()[1].id;
    let lookup = visigrid_engine::workbook::WorkbookLookup::new(&wb, wb.sheets()[0].id);
    for _ in 0..200 {
        let (a, b) = (rng.below(ROWS + 5), rng.below(ROWS + 5));
        let (c, d) = (rng.below(COLS + 1), rng.below(COLS + 1));
        let (r0, r1, c0, c1) = (a.min(b), a.max(b), c.min(d), c.max(d));
        let mut fast = Vec::new();
        lookup.numbers_in_range(&SheetRef::Id(other), r0, c0, r1, c1, &mut fast).unwrap();
        let mut expected = Vec::new();
        for r in r0..=r1 {
            for c in c0..=c1 {
                if let Ok(n) = lookup.get_text_sheet(other, r, c).parse::<f64>() {
                    expected.push(n);
                }
            }
        }
        assert!(same(&fast, &expected), "rect ({r0},{c0})-({r1},{c1}):\n fast {fast:?}\n slow {expected:?}");
    }
    let mut missing = Vec::new();
    lookup.numbers_in_range(&SheetRef::Id(visigrid_engine::sheet::SheetId(999)), 0, 0, 10, 3, &mut missing).unwrap();
    assert!(missing.is_empty());
}

/// The one intended difference: the per-cell loop formatted integers through
/// i64, so values at or above 2^63 saturated. SUM now agrees with `+`.
#[test]
fn sum_of_huge_integers_matches_addition() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "1e20");
    wb.set_cell_value_tracked(0, 1, 0, "1e20");
    wb.set_cell_value_tracked(0, 2, 0, "=SUM(A1:A2)");
    wb.set_cell_value_tracked(0, 3, 0, "=A1+A2");
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet = &wb.sheets()[0];
    assert_eq!(sheet.get_cached_value(2, 0), sheet.get_cached_value(3, 0));
    assert_eq!(sheet.get_cached_value(2, 0), Some(visigrid_engine::formula::eval::Value::Number(2e20)));
}
