//! Cell storage behind a [`Sheet`](crate::sheet::Sheet) (#18).
//!
//! Every read and write of a sheet's cells goes through a closed API: point
//! lookups and iteration hand out [`CellRef`](crate::cell::CellRef) views,
//! writes take a closure over a `&mut Cell`, and row and column inserts and
//! deletes are store operations. Nothing outside borrows the representation,
//! so it can change without its callers changing.

mod columns;
// The pre-#18 store: kept as the reference the differential tests check the
// column store against.
#[cfg(test)]
mod hash;

pub(crate) use columns::ColumnStore as CellStore;

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
        fn set_format_to(&mut self, row: usize, col: usize, format: Arc<CellFormat>);
        fn update_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell)) -> bool;
        fn remove(&mut self, row: usize, col: usize) -> Option<Cell>;
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
                fn set_format_to(&mut self, row: usize, col: usize, format: Arc<CellFormat>) {
                    self.set_format(row, col, Cell::default, format)
                }
                fn update_with(&mut self, row: usize, col: usize, f: &dyn Fn(&mut Cell)) -> bool {
                    self.update(row, col, |c| f(c)).is_some()
                }
                fn remove(&mut self, row: usize, col: usize) -> Option<Cell> { <$t>::remove(self, row, col) }
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
                if rng.below(2) == 0 {
                    let f = move |cell: &mut Cell| cell.format = Arc::clone(&fmt);
                    a.upsert_with(r, c, &f);
                    b.upsert_with(r, c, &f);
                } else {
                    // The format-only path: value untouched, cell created if absent.
                    a.set_format_to(r, c, Arc::clone(&fmt));
                    b.set_format_to(r, c, fmt);
                }
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
                // Remove every cell matching a condition, across the sheet.
                let m = rng.below(5) + 2;
                let doomed: Vec<_> =
                    a.cells().into_iter().map(|(pos, _)| pos).filter(|(r, c)| (r + c) % m == 0).collect();
                for (r, c) in doomed {
                    a.remove(r, c);
                    b.remove(r, c);
                }
            }
            11 => {
                // Half the time right at a chunk boundary (1,024 rows), where
                // the chunks kept whole meet the ones rebuilt.
                let at = if rng.below(2) == 0 { (rng.below(3) + 1) * 1024 + rng.below(5) - 2 } else { rng.below(ROWS) };
                let n = rng.below(40) + 1;
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

    /// Phase 3: clones share chunks and pool pages until written. Fork a
    /// store mid-run, then edit both copies with different operations: each
    /// must still match its own hash-store oracle, so a write that leaked
    /// through a shared chunk or page fails on one side or the other.
    fn run_forked(seed: u64, steps: usize) {
        let mut rng = Rng(seed);
        let (mut hash, mut columns) = (HashStore::default(), ColumnStore::default());
        for _ in 0..steps / 2 {
            step(&mut rng, &mut hash, &mut columns);
        }
        let (mut hash2, mut columns2) = (hash.clone(), columns.clone());
        let mut rng2 = Rng(seed ^ 0xD1B5_4A32_D192_ED03);
        for i in 0..steps {
            step(&mut rng, &mut hash, &mut columns);
            step(&mut rng2, &mut hash2, &mut columns2);
            if i % 25 == 0 || i + 1 == steps {
                assert_eq!(hash.len(), columns.len(), "original len after step {i} (seed {seed})");
                assert_eq!(hash2.len(), columns2.len(), "clone len after step {i} (seed {seed})");
                assert_eq!(snapshot(&hash), snapshot(&columns), "original after step {i} (seed {seed})");
                assert_eq!(snapshot(&hash2), snapshot(&columns2), "clone after step {i} (seed {seed})");
            }
        }
        // Once the other copy is gone, the survivor owns everything again.
        drop((hash2, columns2));
        for _ in 0..50 {
            step(&mut rng, &mut hash, &mut columns);
        }
        assert_eq!(snapshot(&hash), snapshot(&columns), "survivor (seed {seed})");
    }

    #[test]
    fn clones_are_independent() {
        for seed in 1..=16u64 {
            run_forked(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 300);
        }
    }

    /// A clone that is never written shares everything, including computed
    /// results: results written to the original after the clone stay there.
    #[test]
    fn computed_results_do_not_leak_between_clones() {
        use crate::formula::eval::Value;
        let mut store = ColumnStore::default();
        for r in 0..3000 {
            let mut c = Cell::default();
            c.set(&format!("=A{}+1", r + 1));
            store.upsert(r, 1, Cell::default, move |cell| *cell = c);
            store.set_computed(r, 1, Value::Number(r as f64));
        }
        let snap = store.clone();
        store.set_computed(5, 1, Value::Number(-1.0));
        store.clear_all_computed();
        store.set_computed(2999, 1, Value::Text("new".into()));
        assert_eq!(snap.with_computed(5, 1, |v| v.cloned()), Some(Value::Number(5.0)));
        assert_eq!(snap.with_computed(2999, 1, |v| v.cloned()), Some(Value::Number(2999.0)));
        assert_eq!(snap.computed_count(), 3000);
        assert_eq!(store.computed_count(), 1);
    }

    /// Longer run for local soak testing: `cargo test -- --ignored`.
    #[test]
    #[ignore]
    fn column_store_matches_hash_store_long() {
        for seed in 1..=200u64 {
            run(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15), 2_000);
        }
    }

    // A document naming the same coordinates twice keeps the later cell,
    // as the hash store did, and the counts stay true. Both readers.
    fn cell(value: &str) -> Cell {
        let mut c = Cell::default();
        c.set(value);
        c
    }

    fn check_duplicate_load(store: ColumnStore) {
        assert_eq!(store.len(), 1);
        assert_eq!(store.iter().count(), 1);
        assert_eq!(store.get(0, 0).map(|c| c.raw_display()), Some("second".to_string()));
        let mut store = store;
        assert!(store.remove(0, 0).is_some());
        assert_eq!(store.len(), 0);
        assert_eq!(store.iter().count(), 0);
    }

    #[test]
    fn duplicate_coordinates_through_the_sequence_reader() {
        let json = serde_json::json!([
            [[0, 0], serde_json::to_value(cell("first")).unwrap()],
            [[0, 0], serde_json::to_value(cell("second")).unwrap()],
        ]);
        let store: ColumnStore = serde_json::from_value(json).unwrap();
        check_duplicate_load(store);
    }

    #[test]
    fn duplicate_coordinates_through_the_map_reader() {
        use serde::de::value::{MapDeserializer, SeqDeserializer};
        use serde::Deserialize;
        type Key = SeqDeserializer<std::vec::IntoIter<u32>, serde_json::Error>;
        let entries: Vec<(Key, serde_json::Value)> = vec![
            (SeqDeserializer::new(vec![0u32, 0].into_iter()), serde_json::to_value(cell("first")).unwrap()),
            (SeqDeserializer::new(vec![0u32, 0].into_iter()), serde_json::to_value(cell("second")).unwrap()),
        ];
        let store = ColumnStore::deserialize(MapDeserializer::<_, serde_json::Error>::new(entries.into_iter())).unwrap();
        check_duplicate_load(store);
    }

    // Phase 2: computed results live with their formulas.
    mod computed {
        use super::super::columns::ColumnStore;
        use crate::cell::{Cell, SpillInfo};
        use crate::formula::eval::Value;

        fn store_with_formula(row: usize, col: usize) -> ColumnStore {
            let mut s = ColumnStore::default();
            s.upsert(row, col, Cell::default, |c| c.set("=A1+1"));
            s.set_computed(row, col, Value::Number(42.0));
            s
        }

        fn result(s: &ColumnStore, row: usize, col: usize) -> Option<Value> {
            s.with_computed(row, col, |v| v.cloned())
        }

        #[test]
        fn a_write_that_keeps_the_formula_keeps_its_result() {
            let mut s = store_with_formula(5, 1);
            s.upsert(5, 1, Cell::default, |c| c.set_spill_info(Some(SpillInfo { rows: 2, cols: 1 })));
            assert_eq!(result(&s, 5, 1), Some(Value::Number(42.0)));
            s.update(5, 1, |c| c.set_style_id(Some(3)));
            assert_eq!(result(&s, 5, 1), Some(Value::Number(42.0)));
            s.set_format(5, 1, Cell::default, std::sync::Arc::new(Default::default()));
            assert_eq!(result(&s, 5, 1), Some(Value::Number(42.0)));
        }

        #[test]
        fn an_explicit_clear_or_a_non_formula_value_drops_it() {
            let mut s = store_with_formula(5, 1);
            s.clear_computed(5, 1);
            assert_eq!(result(&s, 5, 1), None);

            let mut s = store_with_formula(5, 1);
            s.upsert(5, 1, Cell::default, |c| c.set("7"));
            assert_eq!(result(&s, 5, 1), None, "no longer a formula");
        }

        #[test]
        fn a_different_formula_starts_uncomputed_even_without_a_clear() {
            let mut s = store_with_formula(5, 1);
            s.upsert(5, 1, Cell::default, |c| c.set("=B1*2"));
            assert_eq!(result(&s, 5, 1), None, "new formula, no stale result");
            let mut s = store_with_formula(5, 1);
            s.update(5, 1, |c| c.set("=A1+1"));
            assert_eq!(result(&s, 5, 1), Some(Value::Number(42.0)), "same formula keeps it");
        }

        #[test]
        fn a_reused_formula_id_starts_uncomputed() {
            let mut s = store_with_formula(5, 1);
            s.remove(5, 1);
            s.upsert(9, 2, Cell::default, |c| c.set("=B1*2"));
            assert_eq!(result(&s, 9, 2), None);
        }

        #[test]
        fn results_move_with_their_cells() {
            let mut s = store_with_formula(5, 1);
            s.insert_rows(0, 3, 1_000);
            assert_eq!(result(&s, 8, 1), Some(Value::Number(42.0)));
            s.insert_cols(0, 2, 100);
            assert_eq!(result(&s, 8, 3), Some(Value::Number(42.0)));
            s.delete_rows(0, 3);
            s.delete_cols(0, 2);
            assert_eq!(result(&s, 5, 1), Some(Value::Number(42.0)));
            s.delete_rows(5, 1);
            assert_eq!(s.computed_count(), 0, "deleted formula takes its result with it");
        }

        #[test]
        fn a_result_where_there_is_no_formula_is_ignored() {
            let mut s = ColumnStore::default();
            s.upsert(0, 0, Cell::default, |c| c.set("text"));
            s.set_computed(0, 0, Value::Number(1.0));
            s.set_computed(3, 3, Value::Number(1.0));
            assert_eq!(result(&s, 0, 0), None);
            assert_eq!(s.computed_count(), 0);
        }

        #[test]
        fn clear_all_forgets_every_result() {
            let mut s = store_with_formula(5, 1);
            s.upsert(6, 1, Cell::default, |c| c.set("=A1*3"));
            s.set_computed(6, 1, Value::Number(3.0));
            assert_eq!(s.computed_count(), 2);
            s.clear_all_computed();
            assert_eq!(s.computed_count(), 0);
        }
    }
}
