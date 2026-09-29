//! Column store: cells kept by column in 1,024-row chunks (#18).
//!
//! A hash entry per cell cost ~100 bytes before its contents. Here a column
//! is a sorted list of its non-empty chunks, and each chunk picks the
//! cheapest representation for what it holds:
//!
//! ```text
//! Sparse   a few cells: sorted (offset, slot) pairs            ~24 B/cell
//! Numbers  only numbers: presence bitmap + f64 per row          ~8 B/cell
//! Texts    only text: presence bitmap + string id per row       ~4 B/cell
//! Mixed    anything: presence bitmap + 16-byte slot per row    ~16 B/cell
//! ```
//!
//! Every dense payload is its own heap buffer, so a chunk costs only what
//! its variant holds. Text is interned in a reference-counted pool, formulas
//! live in a table, formats are ids into a table (one per chunk when they
//! agree, one per row when they don't), and the rare per-cell metadata
//! (spill state, import style, frozen formula) sits in a side map that is
//! only consulted for chunks that have any.
//!
//! Lookups are O(log C) in the column's non-empty chunks, which is a
//! handful of comparisons.

use std::collections::HashMap;
use std::sync::Arc;

use crate::cell::{Cell, CellExtras, CellFormat, CellRef, CellValue, ValueRef};
use crate::formula::parser::ParsedExpr;

const CHUNK_BITS: usize = 10;
const CHUNK_ROWS: usize = 1 << CHUNK_BITS;
const WORDS: usize = CHUNK_ROWS / 64;
/// A sparse chunk becomes dense past this many cells...
const DENSIFY_AT: usize = 128;
/// ...and a dense one sparse again below this many, so a column edited back
/// and forth across one threshold does not re-encode on every edit.
const SPARSIFY_BELOW: usize = 32;

type StrId = u32;
type FormulaId = u32;
type FormatId = u32;

#[inline]
fn split(row: usize) -> (u32, usize) {
    ((row >> CHUNK_BITS) as u32, row & (CHUNK_ROWS - 1))
}

#[inline]
fn join(chunk: u32, off: usize) -> usize {
    ((chunk as usize) << CHUNK_BITS) | off
}

/// One stored value. 16 bytes.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Slot {
    Empty,
    Number(f64),
    Text(StrId),
    Formula(FormulaId),
}

type Bits = Box<[u64; WORDS]>;

fn new_bits() -> Bits {
    Box::new([0; WORDS])
}

#[inline]
fn bit(bits: &[u64; WORDS], off: usize) -> bool {
    bits[off / 64] & (1 << (off % 64)) != 0
}

#[inline]
fn set_bit(bits: &mut [u64; WORDS], off: usize, on: bool) {
    if on {
        bits[off / 64] |= 1 << (off % 64);
    } else {
        bits[off / 64] &= !(1 << (off % 64));
    }
}

#[derive(Debug, Clone)]
enum Cells {
    Sparse(Vec<(u16, Slot)>),
    Numbers { present: Bits, values: Box<[f64]> },
    Texts { present: Bits, ids: Box<[StrId]> },
    Mixed { present: Bits, slots: Box<[Slot]> },
}

impl Cells {
    fn get(&self, off: usize) -> Option<Slot> {
        match self {
            Cells::Sparse(v) => v
                .binary_search_by_key(&(off as u16), |e| e.0)
                .ok()
                .map(|i| v[i].1),
            Cells::Numbers { present, values } => bit(present, off).then(|| Slot::Number(values[off])),
            Cells::Texts { present, ids } => bit(present, off).then(|| Slot::Text(ids[off])),
            Cells::Mixed { present, slots } => bit(present, off).then(|| slots[off]),
        }
    }

    /// Present cells in offset order, without allocating.
    fn slots(&self) -> SlotIter<'_> {
        match self {
            Cells::Sparse(v) => SlotIter::Sparse(v.iter()),
            Cells::Numbers { present, .. } | Cells::Texts { present, .. } | Cells::Mixed { present, .. } => {
                SlotIter::Dense { cells: self, present, word: 0, bits: present[0] }
            }
        }
    }

    fn entries(&self) -> Vec<(usize, Slot)> {
        self.slots().collect()
    }

    /// Dense form of `entries`, specialised when every value agrees in kind.
    fn dense_from(entries: &[(usize, Slot)]) -> Cells {
        let mut present = new_bits();
        for (off, _) in entries {
            set_bit(&mut present, *off, true);
        }
        if entries.iter().all(|(_, s)| matches!(s, Slot::Number(_))) {
            let mut values = vec![0.0; CHUNK_ROWS].into_boxed_slice();
            for (off, s) in entries {
                if let Slot::Number(n) = s {
                    values[*off] = *n;
                }
            }
            Cells::Numbers { present, values }
        } else if entries.iter().all(|(_, s)| matches!(s, Slot::Text(_))) {
            let mut ids = vec![0; CHUNK_ROWS].into_boxed_slice();
            for (off, s) in entries {
                if let Slot::Text(id) = s {
                    ids[*off] = *id;
                }
            }
            Cells::Texts { present, ids }
        } else {
            let mut slots = vec![Slot::Empty; CHUNK_ROWS].into_boxed_slice();
            for (off, s) in entries {
                slots[*off] = *s;
            }
            Cells::Mixed { present, slots }
        }
    }

    /// Store `slot` at `off`, returning what was there. `count_after` is the
    /// number of cells present after the write, for the sparse/dense choice.
    fn set(&mut self, off: usize, slot: Slot, count_after: usize) -> Option<Slot> {
        let fits = matches!(
            (&*self, slot),
            (Cells::Sparse(_), _)
                | (Cells::Mixed { .. }, _)
                | (Cells::Numbers { .. }, Slot::Number(_))
                | (Cells::Texts { .. }, Slot::Text(_))
        );
        if !fits {
            self.to_mixed();
        }
        let prev = match self {
            Cells::Sparse(v) => match v.binary_search_by_key(&(off as u16), |e| e.0) {
                Ok(i) => Some(std::mem::replace(&mut v[i].1, slot)),
                Err(i) => {
                    v.insert(i, (off as u16, slot));
                    None
                }
            },
            Cells::Numbers { present, values } => {
                let prev = bit(present, off).then(|| Slot::Number(values[off]));
                set_bit(present, off, true);
                if let Slot::Number(n) = slot {
                    values[off] = n;
                }
                prev
            }
            Cells::Texts { present, ids } => {
                let prev = bit(present, off).then(|| Slot::Text(ids[off]));
                set_bit(present, off, true);
                if let Slot::Text(id) = slot {
                    ids[off] = id;
                }
                prev
            }
            Cells::Mixed { present, slots } => {
                let prev = bit(present, off).then(|| slots[off]);
                set_bit(present, off, true);
                slots[off] = slot;
                prev
            }
        };
        if let Cells::Sparse(v) = self {
            if count_after > DENSIFY_AT {
                let entries: Vec<(usize, Slot)> = v.iter().map(|e| (e.0 as usize, e.1)).collect();
                *self = Cells::dense_from(&entries);
            }
        }
        prev
    }

    /// Clear `off`, returning what was there. `count_after` is the number of
    /// cells left, for the sparse/dense choice.
    fn remove(&mut self, off: usize, count_after: usize) -> Option<Slot> {
        let prev = self.get(off)?;
        match self {
            Cells::Sparse(v) => {
                if let Ok(i) = v.binary_search_by_key(&(off as u16), |e| e.0) {
                    v.remove(i);
                }
            }
            Cells::Numbers { present, .. } | Cells::Texts { present, .. } | Cells::Mixed { present, .. } => {
                set_bit(present, off, false);
            }
        }
        if !matches!(self, Cells::Sparse(_)) && count_after < SPARSIFY_BELOW {
            let entries = self.entries();
            *self = Cells::Sparse(entries.into_iter().map(|(o, s)| (o as u16, s)).collect());
        }
        Some(prev)
    }

    fn to_mixed(&mut self) {
        let entries = self.entries();
        let mut present = new_bits();
        let mut slots = vec![Slot::Empty; CHUNK_ROWS].into_boxed_slice();
        for (off, s) in entries {
            set_bit(&mut present, off, true);
            slots[off] = s;
        }
        *self = Cells::Mixed { present, slots };
    }
}

enum SlotIter<'a> {
    Sparse(std::slice::Iter<'a, (u16, Slot)>),
    Dense { cells: &'a Cells, present: &'a [u64; WORDS], word: usize, bits: u64 },
}

impl Iterator for SlotIter<'_> {
    type Item = (usize, Slot);

    fn next(&mut self) -> Option<(usize, Slot)> {
        match self {
            SlotIter::Sparse(it) => it.next().map(|&(off, slot)| (off as usize, slot)),
            SlotIter::Dense { cells, present, word, bits } => loop {
                if *bits != 0 {
                    let off = *word * 64 + bits.trailing_zeros() as usize;
                    *bits &= *bits - 1;
                    let slot = match cells {
                        Cells::Numbers { values, .. } => Slot::Number(values[off]),
                        Cells::Texts { ids, .. } => Slot::Text(ids[off]),
                        Cells::Mixed { slots, .. } => slots[off],
                        Cells::Sparse(_) => unreachable!("dense iterator over sparse cells"),
                    };
                    return Some((off, slot));
                }
                *word += 1;
                if *word >= WORDS {
                    return None;
                }
                *bits = present[*word];
            },
        }
    }
}

#[derive(Debug, Clone)]
enum Formats {
    Uniform(FormatId),
    Varied(Box<[FormatId]>),
}

impl Formats {
    fn get(&self, off: usize) -> FormatId {
        match self {
            Formats::Uniform(id) => *id,
            Formats::Varied(ids) => ids[off],
        }
    }

    /// `alone`: this is the only cell in the chunk, so a uniform chunk can
    /// simply take its format.
    fn set(&mut self, off: usize, id: FormatId, alone: bool) {
        match self {
            Formats::Uniform(u) if *u == id => {}
            Formats::Uniform(u) if alone => *u = id,
            Formats::Uniform(u) => {
                let mut ids = vec![*u; CHUNK_ROWS].into_boxed_slice();
                ids[off] = id;
                *self = Formats::Varied(ids);
            }
            Formats::Varied(ids) => ids[off] = id,
        }
    }
}

#[derive(Debug, Clone)]
struct Chunk {
    cells: Cells,
    formats: Formats,
    count: u16,
    /// Cells in this chunk with an entry in the store's extras map.
    extras: u16,
}

impl Chunk {
    fn new(format: FormatId) -> Chunk {
        Chunk { cells: Cells::Sparse(Vec::new()), formats: Formats::Uniform(format), count: 0, extras: 0 }
    }
}

#[derive(Debug, Clone, Default)]
struct Column {
    /// Non-empty chunks, sorted by chunk index.
    chunks: Vec<(u32, Chunk)>,
}

impl Column {
    fn chunk(&self, idx: u32) -> Option<&Chunk> {
        self.chunks.binary_search_by_key(&idx, |c| c.0).ok().map(|i| &self.chunks[i].1)
    }

    fn chunk_mut_or_new(&mut self, idx: u32, format: FormatId) -> &mut Chunk {
        let i = match self.chunks.binary_search_by_key(&idx, |c| c.0) {
            Ok(i) => i,
            Err(i) => {
                self.chunks.insert(i, (idx, Chunk::new(format)));
                i
            }
        };
        &mut self.chunks[i].1
    }
}

/// Interned text with reference counts, so repeated values cost one copy
/// and a cleared cell gives its string back. The index stores only ids and
/// hashes through `entries`, so each distinct string is kept once.
#[derive(Debug, Clone, Default)]
struct StringPool {
    entries: Vec<Option<(Box<str>, u32)>>,
    free: Vec<StrId>,
    index: hashbrown::HashTable<StrId>,
}

fn hash_str(s: &str) -> u64 {
    use std::hash::BuildHasher;
    rustc_hash::FxBuildHasher.hash_one(s)
}

impl StringPool {
    fn text(entries: &[Option<(Box<str>, u32)>], id: StrId) -> &str {
        &entries[id as usize].as_ref().expect("live string id").0
    }

    fn intern(&mut self, s: &str) -> StrId {
        let hash = hash_str(s);
        let entries = &mut self.entries;
        if let Some(&id) = self.index.find(hash, |&id| Self::text(entries, id) == s) {
            if let Some((_, refs)) = &mut entries[id as usize] {
                *refs += 1;
            }
            return id;
        }
        let id = match self.free.pop() {
            Some(id) => {
                entries[id as usize] = Some((Box::from(s), 1));
                id
            }
            None => {
                entries.push(Some((Box::from(s), 1)));
                (entries.len() - 1) as StrId
            }
        };
        let entries = &self.entries;
        self.index.insert_unique(hash, id, |&i| hash_str(Self::text(entries, i)));
        id
    }

    fn get(&self, id: StrId) -> &str {
        Self::text(&self.entries, id)
    }

    fn release(&mut self, id: StrId) {
        let Some((text, refs)) = &mut self.entries[id as usize] else { return };
        *refs -= 1;
        if *refs == 0 {
            let hash = hash_str(text);
            if let Ok(entry) = self.index.find_entry(hash, |&i| i == id) {
                entry.remove();
            }
            self.entries[id as usize] = None;
            self.free.push(id);
        }
    }
}

#[derive(Debug, Clone)]
struct Formula {
    source: String,
    ast: Option<Box<ParsedExpr>>,
}

#[derive(Debug, Clone, Default)]
struct FormulaTable {
    entries: Vec<Option<Formula>>,
    free: Vec<FormulaId>,
}

impl FormulaTable {
    fn insert(&mut self, f: Formula) -> FormulaId {
        match self.free.pop() {
            Some(id) => {
                self.entries[id as usize] = Some(f);
                id
            }
            None => {
                self.entries.push(Some(f));
                (self.entries.len() - 1) as FormulaId
            }
        }
    }

    fn get(&self, id: FormulaId) -> &Formula {
        self.entries[id as usize].as_ref().expect("live formula id")
    }

    fn remove(&mut self, id: FormulaId) -> Formula {
        let f = self.entries[id as usize].take().expect("live formula id");
        self.free.push(id);
        f
    }
}

/// Formats by id. Id 0 is the shared default every unformatted cell uses.
#[derive(Debug, Clone)]
struct FormatTable {
    formats: Vec<Arc<CellFormat>>,
    index: HashMap<Arc<CellFormat>, FormatId>,
    /// Allocation address of each stored format. The table holds those Arcs,
    /// so an incoming Arc at one of these addresses is that same format, and
    /// the 112-byte value need not be hashed. Sheets share format Arcs, so
    /// this is the common case.
    by_address: HashMap<usize, FormatId>,
}

impl Default for FormatTable {
    fn default() -> Self {
        let default = Cell::default().format;
        let mut index = HashMap::new();
        index.insert(Arc::clone(&default), 0);
        let mut by_address = HashMap::new();
        by_address.insert(Arc::as_ptr(&default) as usize, 0);
        FormatTable { formats: vec![default], index, by_address }
    }
}

impl FormatTable {
    fn intern(&mut self, format: Arc<CellFormat>) -> FormatId {
        if let Some(&id) = self.by_address.get(&(Arc::as_ptr(&format) as usize)) {
            return id;
        }
        if let Some(&id) = self.index.get(&format) {
            return id;
        }
        let id = self.formats.len() as FormatId;
        self.by_address.insert(Arc::as_ptr(&format) as usize, id);
        self.formats.push(Arc::clone(&format));
        self.index.insert(format, id);
        id
    }

    fn get(&self, id: FormatId) -> &Arc<CellFormat> {
        &self.formats[id as usize]
    }
}

/// A cell's stored pieces, moved around by structural edits without
/// touching the pools.
struct Raw {
    slot: Slot,
    format: FormatId,
    extras: Option<CellExtras>,
}

/// The cells of one sheet, stored by column.
#[derive(Debug, Clone, Default)]
pub(crate) struct ColumnStore {
    columns: Vec<Column>,
    strings: StringPool,
    formulas: FormulaTable,
    formats: FormatTable,
    extras: HashMap<(u32, u32), CellExtras>,
    len: usize,
}

impl ColumnStore {
    /// Number of stored cells (with a value, a format, or metadata).
    pub fn len(&self) -> usize {
        self.len
    }

    /// Nothing to reserve: chunks are allocated as rows arrive.
    pub fn reserve(&mut self, _additional: usize) {}

    pub fn get(&self, row: usize, col: usize) -> Option<CellRef<'_>> {
        let (idx, off) = split(row);
        let chunk = self.columns.get(col)?.chunk(idx)?;
        let slot = chunk.cells.get(off)?;
        Some(self.view(row, col, chunk, off, slot))
    }

    fn value_ref(&self, slot: Slot) -> ValueRef<'_> {
        match slot {
            Slot::Empty => ValueRef::Empty,
            Slot::Number(n) => ValueRef::Number(n),
            Slot::Text(id) => ValueRef::Text(self.strings.get(id)),
            Slot::Formula(id) => {
                let f = self.formulas.get(id);
                ValueRef::Formula { source: &f.source, ast: f.ast.as_deref() }
            }
        }
    }

    /// Every stored cell, column by column, top to bottom.
    pub fn iter(&self) -> Iter<'_> {
        Iter { store: self, col: 0, chunk: 0, slots: None }
    }

    fn view(&self, row: usize, col: usize, chunk: &Chunk, off: usize, slot: Slot) -> CellRef<'_> {
        let extras = if chunk.extras > 0 { self.extras.get(&(row as u32, col as u32)) } else { None };
        CellRef::from_parts(self.value_ref(slot), self.formats.get(chunk.formats.get(off)), extras)
    }

    /// Change a cell that exists; `None` when there is no cell there.
    pub fn update<R>(&mut self, row: usize, col: usize, f: impl FnOnce(&mut Cell) -> R) -> Option<R> {
        let mut cell = self.remove(row, col)?;
        let out = f(&mut cell);
        self.insert(row, col, cell);
        Some(out)
    }

    /// Change a cell, creating it from `init` first if it does not exist.
    pub fn upsert<R>(
        &mut self,
        row: usize,
        col: usize,
        init: impl FnOnce() -> Cell,
        f: impl FnOnce(&mut Cell) -> R,
    ) -> R {
        let mut cell = self.remove(row, col).unwrap_or_else(init);
        let out = f(&mut cell);
        self.insert(row, col, cell);
        out
    }

    /// Store an owned cell, moving its text and formula into the pools.
    fn insert(&mut self, row: usize, col: usize, cell: Cell) {
        let (value, format, extras) = cell.into_parts();
        let slot = match value {
            CellValue::Empty => Slot::Empty,
            CellValue::Number(n) => Slot::Number(n),
            CellValue::Text(s) => Slot::Text(self.strings.intern(&s)),
            CellValue::Formula { source, ast } => Slot::Formula(self.formulas.insert(Formula { source, ast })),
        };
        let format = self.formats.intern(format);
        let extras = extras.filter(|e| !e.is_empty());
        self.put(row, col, Raw { slot, format, extras });
    }

    /// Place stored pieces at a position that is currently empty.
    fn put(&mut self, row: usize, col: usize, raw: Raw) {
        let (idx, off) = split(row);
        if self.columns.len() <= col {
            self.columns.resize_with(col + 1, Column::default);
        }
        let chunk = self.columns[col].chunk_mut_or_new(idx, raw.format);
        let count_after = chunk.count as usize + 1;
        let prev = chunk.cells.set(off, raw.slot, count_after);
        debug_assert!(prev.is_none(), "put over an existing cell at ({row}, {col})");
        chunk.count += 1;
        chunk.formats.set(off, raw.format, chunk.count == 1);
        if let Some(extras) = raw.extras {
            chunk.extras += 1;
            self.extras.insert((row as u32, col as u32), extras);
        }
        self.len += 1;
    }

    /// Lift a cell's stored pieces out, leaving the position empty. Pool
    /// references stay with the returned pieces.
    fn take(&mut self, row: usize, col: usize) -> Option<Raw> {
        let (idx, off) = split(row);
        let column = self.columns.get_mut(col)?;
        let i = column.chunks.binary_search_by_key(&idx, |c| c.0).ok()?;
        let chunk = &mut column.chunks[i].1;
        let format = chunk.formats.get(off);
        let slot = chunk.cells.remove(off, chunk.count as usize - 1)?;
        chunk.count -= 1;
        let extras = if chunk.extras > 0 {
            let e = self.extras.remove(&(row as u32, col as u32));
            if e.is_some() {
                chunk.extras -= 1;
            }
            e
        } else {
            None
        };
        if chunk.count == 0 {
            column.chunks.remove(i);
        }
        self.len -= 1;
        Some(Raw { slot, format, extras })
    }

    /// Give a removed cell's pool references back.
    fn release(&mut self, raw: &Raw) {
        match raw.slot {
            Slot::Text(id) => self.strings.release(id),
            Slot::Formula(id) => {
                self.formulas.remove(id);
            }
            Slot::Empty | Slot::Number(_) => {}
        }
    }

    pub fn remove(&mut self, row: usize, col: usize) -> Option<Cell> {
        let raw = self.take(row, col)?;
        let value = match raw.slot {
            Slot::Empty => CellValue::Empty,
            Slot::Number(n) => CellValue::Number(n),
            Slot::Text(id) => {
                let text = self.strings.get(id).to_string();
                self.strings.release(id);
                CellValue::Text(text)
            }
            Slot::Formula(id) => {
                let Formula { source, ast } = self.formulas.remove(id);
                CellValue::Formula { source, ast }
            }
        };
        let format = Arc::clone(self.formats.get(raw.format));
        Some(Cell::from_parts(value, format, raw.extras))
    }

    /// Move cells at or below `at` down by `count` rows, dropping any that
    /// would land at or past `limit`.
    pub fn insert_rows(&mut self, at: usize, count: usize, limit: usize) {
        self.shift(|(r, c)| if r >= at { (r + count < limit).then_some((r + count, c)) } else { Some((r, c)) });
    }

    /// Delete `count` rows from `start`; cells below move up.
    pub fn delete_rows(&mut self, start: usize, count: usize) {
        let end = start + count;
        self.shift(|(r, c)| {
            if (start..end).contains(&r) {
                None
            } else if r >= end {
                Some((r - count, c))
            } else {
                Some((r, c))
            }
        });
    }

    /// Move cells at or right of `at` right by `count` columns, dropping any
    /// that would land at or past `limit`.
    pub fn insert_cols(&mut self, at: usize, count: usize, limit: usize) {
        self.shift(|(r, c)| if c >= at { (c + count < limit).then_some((r, c + count)) } else { Some((r, c)) });
    }

    /// Delete `count` columns from `start`; cells to the right move left.
    pub fn delete_cols(&mut self, start: usize, count: usize) {
        let end = start + count;
        self.shift(|(r, c)| {
            if (start..end).contains(&c) {
                None
            } else if c >= end {
                Some((r, c - count))
            } else {
                Some((r, c))
            }
        });
    }

    /// Move every cell through `to`, dropping those it maps to `None`. Cells
    /// that stay put are untouched; the rest are lifted out first and then
    /// placed, so a move never lands on a cell that has not moved yet.
    fn shift(&mut self, to: impl Fn((usize, usize)) -> Option<(usize, usize)>) {
        let moving: Vec<(usize, usize)> =
            self.iter().map(|(pos, _)| pos).filter(|&pos| to(pos) != Some(pos)).collect();
        let lifted: Vec<((usize, usize), Raw)> =
            moving.into_iter().filter_map(|pos| self.take(pos.0, pos.1).map(|raw| (pos, raw))).collect();
        for (pos, raw) in lifted {
            match to(pos) {
                Some((r, c)) => self.put(r, c, raw),
                None => self.release(&raw),
            }
        }
    }
}

/// Iterator over a column store's cells: one flat state machine rather than
/// nested adapters, since exporters and bounds scans walk every cell.
pub(crate) struct Iter<'a> {
    store: &'a ColumnStore,
    col: usize,
    chunk: usize,
    slots: Option<SlotIter<'a>>,
}

impl<'a> Iterator for Iter<'a> {
    type Item = ((usize, usize), CellRef<'a>);

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let column = self.store.columns.get(self.col)?;
            if let Some(slots) = &mut self.slots {
                let (idx, chunk) = &column.chunks[self.chunk];
                if let Some((off, slot)) = slots.next() {
                    let row = join(*idx, off);
                    return Some(((row, self.col), self.store.view(row, self.col, chunk, off, slot)));
                }
                self.slots = None;
                self.chunk += 1;
            }
            match column.chunks.get(self.chunk) {
                Some((_, chunk)) => self.slots = Some(chunk.cells.slots()),
                None => {
                    self.col += 1;
                    self.chunk = 0;
                }
            }
        }
    }
}

/// Serialized exactly as the hash store was (a map from `(row, col)` to
/// `Cell`), so the format does not change with the representation.
/// Deserialization also accepts a sequence of `((row, col), cell)` pairs.
impl serde::Serialize for ColumnStore {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        use serde::ser::SerializeMap;
        let mut map = serializer.serialize_map(Some(self.len))?;
        for ((row, col), cell) in self.iter() {
            map.serialize_entry(&(row as u32, col as u32), &cell.to_cell())?;
        }
        map.end()
    }
}

impl<'de> serde::Deserialize<'de> for ColumnStore {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        struct Cells;
        impl<'de> serde::de::Visitor<'de> for Cells {
            type Value = ColumnStore;
            fn expecting(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
                f.write_str("a map or sequence of (row, col) to cell")
            }
            fn visit_map<A: serde::de::MapAccess<'de>>(self, mut map: A) -> Result<ColumnStore, A::Error> {
                let mut store = ColumnStore::default();
                while let Some(((row, col), cell)) = map.next_entry::<(u32, u32), Cell>()? {
                    store.insert(row as usize, col as usize, cell);
                }
                Ok(store)
            }
            fn visit_seq<A: serde::de::SeqAccess<'de>>(self, mut seq: A) -> Result<ColumnStore, A::Error> {
                let mut store = ColumnStore::default();
                while let Some(((row, col), cell)) = seq.next_element::<((u32, u32), Cell)>()? {
                    store.insert(row as usize, col as usize, cell);
                }
                Ok(store)
            }
        }
        deserializer.deserialize_any(Cells)
    }
}
