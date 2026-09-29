//! Cell storage behind a [`Sheet`](crate::sheet::Sheet) (#18).
//!
//! Every read and write of a sheet's cells goes through a closed API: point
//! lookups and iteration hand out [`CellRef`](crate::cell::CellRef) views,
//! writes take a closure over a `&mut Cell`, and row and column inserts and
//! deletes are store operations. Nothing outside borrows the representation,
//! so it can change without its callers changing.

mod columns;
mod hash;

pub(crate) use hash::HashStore as CellStore;

#[cfg(test)]
mod differential {
    //! The column store against the hash store it replaces: the same random
    //! operations, applied to both, must leave the same cells behind.

    use std::sync::Arc;

    use super::columns::ColumnStore;
    use super::hash::HashStore;
    use crate::cell::{Cell, CellFormat, CellRef, SpillError, SpillInfo};

    /// Both stores, driven through one API.
    trait Store {
        fn len(&self) -> usize;
        fn cells(&self) -> Vec<((usize, usize), CellRef<'_>)>;
        fn get(&self, row: usize, col: usize) -> Option<CellRef<'_>>;
        fn upsert_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell));
        fn update_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell)) -> bool;
        fn remove(&mut self, row: usize, col: usize) -> Option<Cell>;
        fn retain_where(&mut self, keep: &dyn Fn((usize, usize)) -> bool);
        fn insert_rows(&mut self, at: usize, count: usize, limit: usize);
        fn delete_rows(&mut self, start: usize, count: usize);
        fn insert_cols(&mut self, at: usize, count: usize, limit: usize);
        fn delete_cols(&mut self, start: usize, count: usize);
    }

    macro_rules! impl_store {
        ($t:ty) => {
            impl Store for $t {
                fn len(&self) -> usize { <$t>::len(self) }
                fn cells(&self) -> Vec<((usize, usize), CellRef<'_>)> { self.iter().collect() }
                fn get(&self, row: usize, col: usize) -> Option<CellRef<'_>> { <$t>::get(self, row, col) }
                fn upsert_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell)) {
                    self.upsert(row, col, Cell::default, |c| f(c))
                }
                fn update_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell)) -> bool {
                    self.update(row, col, |c| f(c)).is_some()
                }
                fn remove(&mut self, row: usize, col: usize) -> Option<Cell> { <$t>::remove(self, row, col) }
                fn retain_where(&mut self, keep: &dyn Fn((usize, usize)) -> bool) { self.retain(|pos, _| keep(pos)) }
                fn insert_rows(&mut self, at: usize, count: usize, limit: usize) { <$t>::insert_rows(self, at, count, limit) }
                fn delete_rows(&mut self, start: usize, count: usize) { <$t>::delete_rows(self, start, count) }
                fn insert_cols(&mut self, at: usize, count: usize, limit: usize) { <$t>::insert_cols(self, at, count, limit) }
                fn delete_cols(&mut self, start: usize, count: usize) { <$t>::delete_cols(self, start, count) }
            }
        };
    }
    impl_store!(HashStore);
    impl_store!(ColumnStore);

    /// Everything a reader can observe about one cell.
    #[derive(Debug, PartialEq)]
    struct Seen {
        raw: String,
        formula: bool,
        has_ast: bool,
        format: CellFormat,
        style_id: Option<u32>,
        spill_parent: Option<(usize, usize)>,
        spill_info: Option<(usize, usize)>,
        spill_error: Option<(usize, usize)>,
        frozen: Option<String>,
    }

    fn seen(cell: CellRef<'_>) -> Seen {
        Seen {
            raw: cell.raw_display(),
            formula: cell.value().is_formula(),
            has_ast: cell.value().formula_ast().is_some(),
            format: cell.format().clone(),
            style_id: cell.style_id(),
            spill_parent: cell.spill_parent(),
            spill_info: cell.spill_info().map(|i| (i.rows, i.cols)),
            spill_error: cell.spill_error().map(|e| e.blocked_by),
            frozen: cell.frozen_formula().map(str::to_string),
        }
    }

    fn snapshot(store: &dyn Store) -> Vec<((usize, usize), Seen)> {
        let mut all: Vec<_> = store.cells().into_iter().map(|(pos, c)| (pos, seen(c))).collect();
        all.sort_by_key(|(pos, _)| *pos);
        all
    }

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

    const ROWS: usize = 3_500; // spans four chunks
    const COLS: usize = 5;
    const ROW_LIMIT: usize = 4_000;
    const COL_LIMIT: usize = 8;

    fn value(rng: &mut Rng) -> String {
        match rng.below(9) {
            0 => String::new(),
            1 => format!("{}", rng.below(1000)),
            2 => format!("{}.5", rng.below(100)),
            3 => ["alpha", "beta", "gamma", "delta"][rng.below(4)].to_string(),
            4 => format!("unique-{}", rng.next()),
            5 => format!("=A{}+1", rng.below(50) + 1),
            6 => "=SUM(".to_string(), // invalid: formula without an AST
            7 => "#CYCLE!".to_string(),
            _ => "  7 ".to_string(),
        }
    }

    fn format(rng: &mut Rng) -> CellFormat {
        let mut f = CellFormat::default();
        match rng.below(3) {
            0 => {}
            1 => f.bold = true,
            _ => f.italic = true,
        }
        f
    }

    /// One random operation, applied identically to both stores.
    fn step(rng: &mut Rng, a: &mut dyn Store, b: &mut dyn Store) {
        let (r, c) = (rng.below(ROWS), rng.below(COLS));
        match rng.below(14) {
            0..=3 => {
                let v = value(rng);
                let f = move |cell: &mut Cell| cell.set(&v);
                a.upsert_with(r, c, &f);
                b.upsert_with(r, c, &f);
            }
            4 => {
                let fmt = Arc::new(format(rng));
                let f = move |cell: &mut Cell| cell.format = Arc::clone(&fmt);
                a.upsert_with(r, c, &f);
                b.upsert_with(r, c, &f);
            }
            5 => {
                let which = rng.below(5);
                let f = move |cell: &mut Cell| match which {
                    0 => cell.set_style_id(Some(3)),
                    1 => cell.set_spill_parent(Some((1, 1))),
                    2 => cell.set_spill_info(Some(SpillInfo { rows: 2, cols: 3 })),
                    3 => cell.set_spill_error(Some(SpillError { blocked_by: (4, 2) })),
                    _ => cell.clear_spill_state(),
                };
                a.upsert_with(r, c, &f);
                b.upsert_with(r, c, &f);
            }
            6 => {
                let v = value(rng);
                let f = move |cell: &mut Cell| cell.set(&v);
                assert_eq!(a.update_with(r, c, &f), b.update_with(r, c, &f), "update presence at ({r}, {c})");
            }
            7 => {
                let (x, y) = (a.remove(r, c), b.remove(r, c));
                assert_eq!(x.map(|c| seen(c.as_ref())), y.map(|c| seen(c.as_ref())), "removed cell at ({r}, {c})");
            }
            8 => {
                // Fill a run in one column: forces sparse chunks dense.
                let start = rng.below(ROWS - 300);
                let kind = rng.below(3);
                for row in start..start + 300 {
                    let v = match kind {
                        0 => format!("{row}"),
                        1 => ["alpha", "beta"][row % 2].to_string(),
                        _ => if row % 3 == 0 { "x".to_string() } else { format!("{row}") },
                    };
                    let f = move |cell: &mut Cell| cell.set(&v);
                    a.upsert_with(row, c, &f);
                    b.upsert_with(row, c, &f);
                }
            }
            9 => {
                // Clear most of a run: forces dense chunks sparse again.
                let start = rng.below(ROWS - 300);
                for row in (start..start + 300).filter(|row| row % 13 != 0) {
                    a.remove(row, c);
                    b.remove(row, c);
                }
            }
            10 => {
                let m = rng.below(5) + 2;
                a.retain_where(&|(r, c)| (r + c) % m != 0);
                b.retain_where(&|(r, c)| (r + c) % m != 0);
            }
            11 => {
                let (at, n) = (rng.below(ROWS), rng.below(40) + 1);
                if rng.below(2) == 0 {
                    a.insert_rows(at, n, ROW_LIMIT);
                    b.insert_rows(at, n, ROW_LIMIT);
                } else {
                    a.delete_rows(at, n);
                    b.delete_rows(at, n);
                }
            }
            12 => {
                let (at, n) = (rng.below(COLS), rng.below(2) + 1);
                if rng.below(2) == 0 {
                    a.insert_cols(at, n, COL_LIMIT);
                    b.insert_cols(at, n, COL_LIMIT);
                } else {
                    a.delete_cols(at, n);
                    b.delete_cols(at, n);
                }
            }
            _ => {
                assert_eq!(a.get(r, c).map(seen), b.get(r, c).map(seen), "get at ({r}, {c})");
            }
        }
    }

    fn run(seed: u64, steps: usize) {
        let mut rng = Rng(seed);
        let (mut hash, mut columns) = (HashStore::default(), ColumnStore::default());
        for i in 0..steps {
            step(&mut rng, &mut hash, &mut columns);
            assert_eq!(hash.len(), columns.len(), "len after step {i} (seed {seed})");
            if i % 25 == 0 || i + 1 == steps {
                assert_eq!(snapshot(&hash), snapshot(&columns), "state after step {i} (seed {seed})");
            }
        }
    }

    #[test]
    fn column_store_matches_hash_store() {
        for seed in 1..=24u64 {
            run(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 400);
        }
    }

    /// Longer run for local soak testing: `cargo test -- --ignored`.
    #[test]
    #[ignore]
    fn column_store_matches_hash_store_long() {
        for seed in 1..=200u64 {
            run(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 2_000);
        }
    }
}
