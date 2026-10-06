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
//!
//! Cloning is copy-on-write (#18 phase 3). Chunks are shared behind `Arc`,
//! and the string pool and formula table are split into shared pages, so a
//! clone copies pointers, and a later write copies only the chunk or page it
//! lands in. Rewind preview keeps a clone of the whole workbook, and so do
//! operation plans.

use std::cell::RefCell;
use std::collections::HashMap;
use std::sync::Arc;

use crate::cell::{Cell, CellExtras, CellFormat, CellRef, CellValue, ValueRef};
use crate::formula::eval::Value;
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

const PAGE_BITS: usize = 10;
const PAGE: usize = 1 << PAGE_BITS;

/// A vector in pages that clones share until one is written: a clone copies
/// a pointer per page, and a write copies only the page it lands in.
#[derive(Debug)]
struct Paged<T> {
    pages: Vec<Arc<Vec<T>>>,
    len: usize,
}

impl<T> Clone for Paged<T> {
    fn clone(&self) -> Self {
        Paged { pages: self.pages.clone(), len: self.len }
    }
}

impl<T> Default for Paged<T> {
    fn default() -> Self {
        Paged { pages: Vec::new(), len: 0 }
    }
}

impl<T: Clone> Paged<T> {
    fn len(&self) -> usize {
        self.len
    }

    fn get(&self, i: usize) -> Option<&T> {
        self.pages.get(i >> PAGE_BITS)?.get(i & (PAGE - 1))
    }

    fn get_mut(&mut self, i: usize) -> Option<&mut T> {
        let page = self.pages.get_mut(i >> PAGE_BITS)?;
        Arc::make_mut(page).get_mut(i & (PAGE - 1))
    }

    fn push(&mut self, value: T) {
        if self.len.is_multiple_of(PAGE) {
            self.pages.push(Arc::new(Vec::new()));
        }
        Arc::make_mut(self.pages.last_mut().expect("page just ensured")).push(value);
        self.len += 1;
    }

    fn iter(&self) -> impl Iterator<Item = &T> {
        self.pages.iter().flat_map(|p| p.iter())
    }
}

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
    /// Non-empty chunks, sorted by chunk index. Shared with clones until
    /// written.
    chunks: Vec<(u32, Arc<Chunk>)>,
}

impl Column {
    fn chunk(&self, idx: u32) -> Option<&Chunk> {
        self.chunks.binary_search_by_key(&idx, |c| c.0).ok().map(|i| &*self.chunks[i].1)
    }

    fn chunk_mut_or_new(&mut self, idx: u32, format: FormatId) -> &mut Chunk {
        let i = match self.chunks.binary_search_by_key(&idx, |c| c.0) {
            Ok(i) => i,
            Err(i) => {
                self.chunks.insert(i, (idx, Arc::new(Chunk::new(format))));
                i
            }
        };
        Arc::make_mut(&mut self.chunks[i].1)
    }
}

/// Interned text with reference counts, so repeated values cost one copy
/// and a cleared cell gives its string back. The index stores only ids and
/// hashes through `entries`, so each distinct string is kept once. Clones
/// share the index until a distinct string is added or dropped.
#[derive(Debug, Clone, Default)]
struct StringPool {
    entries: Paged<Option<(Box<str>, u32)>>,
    free: Vec<StrId>,
    index: Arc<hashbrown::HashTable<StrId>>,
}

fn hash_str(s: &str) -> u64 {
    use std::hash::BuildHasher;
    rustc_hash::FxBuildHasher.hash_one(s)
}

impl StringPool {
    fn text(entries: &Paged<Option<(Box<str>, u32)>>, id: StrId) -> &str {
        &entries.get(id as usize).and_then(Option::as_ref).expect("live string id").0
    }

    fn intern(&mut self, s: &str) -> StrId {
        let hash = hash_str(s);
        if let Some(&id) = self.index.find(hash, |&id| Self::text(&self.entries, id) == s) {
            if let Some(Some((_, refs))) = self.entries.get_mut(id as usize) {
                *refs += 1;
            }
            return id;
        }
        let id = match self.free.pop() {
            Some(id) => {
                *self.entries.get_mut(id as usize).expect("freed string id") = Some((Box::from(s), 1));
                id
            }
            None => {
                self.entries.push(Some((Box::from(s), 1)));
                (self.entries.len() - 1) as StrId
            }
        };
        let entries = &self.entries;
        Arc::make_mut(&mut self.index).insert_unique(hash, id, |&i| hash_str(Self::text(entries, i)));
        id
    }

    fn get(&self, id: StrId) -> &str {
        Self::text(&self.entries, id)
    }

    fn release(&mut self, id: StrId) {
        let hash = {
            let Some(Some((text, refs))) = self.entries.get_mut(id as usize) else { return };
            *refs -= 1;
            if *refs > 0 {
                return;
            }
            hash_str(text)
        };
        if let Ok(entry) = Arc::make_mut(&mut self.index).find_entry(hash, |&i| i == id) {
            entry.remove();
        }
        *self.entries.get_mut(id as usize).expect("live string id") = None;
        self.free.push(id);
    }
}

#[derive(Debug, Clone)]
struct Formula {
    source: String,
    ast: Option<Box<ParsedExpr>>,
}

#[derive(Debug, Clone, Default)]
struct FormulaTable {
    unparsed: usize,
    entries: Paged<Option<Formula>>,
    free: Vec<FormulaId>,
    /// Each formula's last computed result, by the same id (#18 phase 2).
    /// Written during recalculation through `&self`, as the position-keyed
    /// cache it replaces was. `None` means not computed yet; readers never
    /// evaluate on a miss.
    values: RefCell<Paged<Option<Value>>>,
}

impl FormulaTable {
    fn insert(&mut self, f: Formula) -> FormulaId {
        self.unparsed += usize::from(f.ast.is_none());
        let id = match self.free.pop() {
            Some(id) => {
                *self.entries.get_mut(id as usize).expect("freed formula id") = Some(f);
                id
            }
            None => {
                self.entries.push(Some(f));
                (self.entries.len() - 1) as FormulaId
            }
        };
        // A reused id must not inherit the previous formula's result.
        let values = self.values.get_mut();
        while values.len() <= id as usize {
            values.push(None);
        }
        if values.get(id as usize).is_some_and(Option::is_some) {
            *values.get_mut(id as usize).expect("value slot") = None;
        }
        id
    }

    fn get(&self, id: FormulaId) -> &Formula {
        self.entries.get(id as usize).and_then(Option::as_ref).expect("live formula id")
    }

    fn remove(&mut self, id: FormulaId) -> Formula {
        let f = self.entries.get_mut(id as usize).and_then(Option::take).expect("live formula id");
        let values = self.values.get_mut();
        if values.get(id as usize).is_some_and(Option::is_some) {
            *values.get_mut(id as usize).expect("value slot") = None;
        }
        self.free.push(id);
        self.unparsed -= usize::from(f.ast.is_none());
        f
    }

    fn with_value<R>(&self, id: FormulaId, f: impl FnOnce(Option<&Value>) -> R) -> R {
        let values = self.values.borrow();
        f(values.get(id as usize).and_then(Option::as_ref))
    }

    /// Recalculation rewrites results that have not changed; skipping those
    /// keeps a page shared with a snapshot from being copied for nothing.
    fn set_value(&self, id: FormulaId, value: Option<Value>) {
        let mut values = self.values.borrow_mut();
        match values.get(id as usize) {
            Some(current) if *current != value => {
                *values.get_mut(id as usize).expect("value slot") = value;
            }
            _ => {}
        }
    }

    fn take_value(&self, id: FormulaId) -> Option<Value> {
        let mut values = self.values.borrow_mut();
        if values.get(id as usize)?.is_none() {
            return None;
        }
        values.get_mut(id as usize).and_then(Option::take)
    }

    fn clear_values(&self) {
        let mut values = self.values.borrow_mut();
        for id in 0..values.len() {
            if values.get(id).is_some_and(Option::is_some) {
                *values.get_mut(id).expect("value slot") = None;
            }
        }
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
    fn find(&self, format: &Arc<CellFormat>) -> Option<FormatId> {
        self.by_address.get(&(Arc::as_ptr(format) as usize)).or_else(|| self.index.get(format)).copied()
    }

    /// Add a format `find` did not have.
    fn add(&mut self, format: Arc<CellFormat>) -> FormatId {
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

/// A stored cell as range aggregation sees it.
pub(crate) enum Scalar<'a> {
    Number(f64),
    Text(&'a str),
    /// A formula with a parsed AST: its computed result, `None` if not
    /// computed yet.
    Computed(Option<&'a Value>),
    /// Empty, or a formula that failed to parse.
    Empty,
}

/// Add a cell after every cell already in `chunks` (rows must arrive in
/// ascending order), opening a new chunk when the row crosses into one.
fn append(chunks: &mut Vec<(u32, Arc<Chunk>)>, row: usize, slot: Slot, format: FormatId) {
    let (idx, off) = split(row);
    if chunks.last().map(|c| c.0) != Some(idx) {
        chunks.push((idx, Arc::new(Chunk::new(format))));
    }
    let chunk = Arc::make_mut(&mut chunks.last_mut().expect("chunk just ensured").1);
    chunk.count += 1;
    chunk.cells.set(off, slot, chunk.count as usize);
    chunk.formats.set(off, format, chunk.count == 1);
}

/// A cell's stored pieces, lifted out and put back by edits without
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
    formats: Arc<FormatTable>,
    extras: Arc<HashMap<(u32, u32), CellExtras>>,
    len: usize,
}

impl ColumnStore {
    pub fn has_unparsed_formulas(&self) -> bool { self.formulas.unparsed != 0 }
    pub fn frozen_positions(&self) -> impl Iterator<Item = (usize, usize)> + '_ {
        self.extras.iter().filter(|(_, e)| e.frozen_formula.is_some()).map(|(&(r, c), _)| (r as usize, c as usize))
    }

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

    /// Positions of the stored cells inside a rectangle, column by column.
    /// Visits only the chunks the rectangle overlaps, so a whole-column
    /// query costs that column, not the sheet.
    pub fn positions_in(&self, min_row: usize, max_row: usize, min_col: usize, max_col: usize) -> Vec<(usize, usize)> {
        let mut out = Vec::new();
        self.for_each_in(min_row, max_row, min_col, max_col, |pos, _| out.push(pos));
        out
    }

    /// Visit the stored cells inside a rectangle, column by column and top
    /// to bottom within a column, touching only the chunks it overlaps.
    pub fn for_each_in(
        &self,
        min_row: usize,
        max_row: usize,
        min_col: usize,
        max_col: usize,
        mut f: impl FnMut((usize, usize), CellRef<'_>),
    ) {
        if self.columns.is_empty() || min_row > max_row || min_col > max_col {
            return;
        }
        let last_col = max_col.min(self.columns.len() - 1);
        let (lo, _) = split(min_row);
        let (hi, _) = split(max_row);
        for col in min_col..=last_col {
            let chunks = &self.columns[col].chunks;
            let start = chunks.partition_point(|c| c.0 < lo);
            for (idx, chunk) in &chunks[start..] {
                if *idx > hi {
                    break;
                }
                for (off, slot) in chunk.cells.slots() {
                    let row = join(*idx, off);
                    if (min_row..=max_row).contains(&row) {
                        f((row, col), self.view(row, col, chunk, off, slot));
                    }
                }
            }
        }
    }

    /// Visit the stored cells inside a rectangle as scalars, in the same
    /// order as [`Self::for_each_in`], for range aggregation: a formula
    /// yields its computed result, read by id under one borrow instead of
    /// being looked up again by position and cloned. Summing a column of
    /// formulas was dominated by that lookup (#29).
    pub fn for_each_scalar_in(
        &self,
        min_row: usize,
        max_row: usize,
        min_col: usize,
        max_col: usize,
        mut f: impl FnMut((usize, usize), Scalar<'_>),
    ) {
        if self.columns.is_empty() || min_row > max_row || min_col > max_col {
            return;
        }
        let values = self.formulas.values.borrow();
        let last_col = max_col.min(self.columns.len() - 1);
        let (lo, _) = split(min_row);
        let (hi, _) = split(max_row);
        for col in min_col..=last_col {
            let chunks = &self.columns[col].chunks;
            let start = chunks.partition_point(|c| c.0 < lo);
            for (idx, chunk) in &chunks[start..] {
                if *idx > hi {
                    break;
                }
                for (off, slot) in chunk.cells.slots() {
                    let row = join(*idx, off);
                    if !(min_row..=max_row).contains(&row) {
                        continue;
                    }
                    let scalar = match slot {
                        Slot::Empty => Scalar::Empty,
                        Slot::Number(n) => Scalar::Number(n),
                        Slot::Text(id) => Scalar::Text(self.strings.get(id)),
                        Slot::Formula(id) if self.formulas.get(id).ast.is_some() => {
                            Scalar::Computed(values.get(id as usize).and_then(Option::as_ref))
                        }
                        Slot::Formula(_) => Scalar::Empty,
                    };
                    f((row, col), scalar);
                }
            }
        }
    }

    /// Change a cell that exists; `None` when there is no cell there.
    pub fn update<R>(&mut self, row: usize, col: usize, f: impl FnOnce(&mut Cell) -> R) -> Option<R> {
        let carried = self.take_computed(row, col);
        let mut cell = self.remove(row, col)?;
        let out = f(&mut cell);
        self.insert(row, col, cell);
        self.restore_computed(row, col, carried);
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
        let carried = self.take_computed(row, col);
        let mut cell = self.remove(row, col).unwrap_or_else(init);
        let out = f(&mut cell);
        self.insert(row, col, cell);
        self.restore_computed(row, col, carried);
        out
    }

    fn formula_id(&self, row: usize, col: usize) -> Option<FormulaId> {
        let (idx, off) = split(row);
        match self.columns.get(col)?.chunk(idx)?.cells.get(off)? {
            Slot::Formula(id) => Some(id),
            _ => None,
        }
    }

    /// Read the computed result of the formula at `(row, col)`. `None` when
    /// there is no formula there or it has not been computed.
    pub fn with_computed<R>(&self, row: usize, col: usize, f: impl FnOnce(Option<&Value>) -> R) -> R {
        match self.formula_id(row, col) {
            Some(id) => self.formulas.with_value(id, f),
            None => f(None),
        }
    }

    /// Record a formula's computed result. Ignored where there is no formula:
    /// every reader consults results only for formula cells.
    pub fn set_computed(&self, row: usize, col: usize, value: Value) {
        if let Some(id) = self.formula_id(row, col) {
            self.formulas.set_value(id, Some(value));
        }
    }

    pub fn clear_computed(&self, row: usize, col: usize) {
        if let Some(id) = self.formula_id(row, col) {
            self.formulas.set_value(id, None);
        }
    }

    /// Forget every computed result (before a full recalculation).
    pub fn clear_all_computed(&self) {
        self.formulas.clear_values();
    }

    /// Number of formulas with a computed result (diagnostics).
    pub fn computed_count(&self) -> usize {
        self.formulas.values.borrow().iter().filter(|v| v.is_some()).count()
    }

    /// A write that keeps the same formula keeps its result (a style, spill
    /// or metadata change), as it did when results were keyed by position.
    /// Returns the result with the formula source it belongs to.
    fn take_computed(&self, row: usize, col: usize) -> Option<(Value, String)> {
        let id = self.formula_id(row, col)?;
        let value = self.formulas.take_value(id)?;
        Some((value, self.formulas.get(id).source.clone()))
    }

    /// Put a carried result back only if the cell still holds the same
    /// formula. A different formula starts uncomputed, whether or not the
    /// caller remembered to clear first.
    fn restore_computed(&self, row: usize, col: usize, carried: Option<(Value, String)>) {
        let Some((value, source)) = carried else { return };
        if let Some(id) = self.formula_id(row, col) {
            if self.formulas.get(id).source == source {
                self.formulas.set_value(id, Some(value));
            }
        }
    }

    /// Store a cell, replacing whatever is there. The loaders use this: a
    /// document naming the same coordinates twice keeps the later cell, as
    /// the hash store did, instead of corrupting the counts.
    fn insert_replacing(&mut self, row: usize, col: usize, cell: Cell) {
        if let Some(old) = self.take(row, col) {
            self.release_slot(old.slot);
        }
        self.insert(row, col, cell);
    }

    /// Give a cell a new format without touching its value, creating it from
    /// `init` if there is none. A format edit used to materialize the cell
    /// and write it back, copying and re-interning its text for nothing.
    pub fn set_format(&mut self, row: usize, col: usize, init: impl FnOnce() -> Cell, format: Arc<CellFormat>) {
        let (idx, off) = split(row);
        let present = self
            .columns
            .get(col)
            .and_then(|c| c.chunk(idx))
            .is_some_and(|chunk| chunk.cells.get(off).is_some());
        if !present {
            let mut cell = init();
            cell.format = format;
            self.insert(row, col, cell);
            return;
        }
        let id = self.intern_format(format);
        let column = &mut self.columns[col];
        let i = column.chunks.binary_search_by_key(&idx, |c| c.0).expect("present chunk");
        let chunk = Arc::make_mut(&mut column.chunks[i].1);
        chunk.formats.set(off, id, chunk.count == 1);
    }

    /// Store an owned cell at an empty position, moving its text and formula
    /// into the pools.
    fn insert(&mut self, row: usize, col: usize, cell: Cell) {
        let (value, format, extras) = cell.into_parts();
        let slot = match value {
            CellValue::Empty => Slot::Empty,
            CellValue::Number(n) => Slot::Number(n),
            CellValue::Text(s) => Slot::Text(self.strings.intern(&s)),
            CellValue::Formula { source, ast } => Slot::Formula(self.formulas.insert(Formula { source, ast })),
        };
        let format = self.intern_format(format);
        let extras = extras.filter(|e| !e.is_empty());
        self.put(row, col, Raw { slot, format, extras });
    }

    /// The table is shared with clones, so it is copied only for a format it
    /// does not have yet.
    fn intern_format(&mut self, format: Arc<CellFormat>) -> FormatId {
        match self.formats.find(&format) {
            Some(id) => id,
            None => Arc::make_mut(&mut self.formats).add(format),
        }
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
            Arc::make_mut(&mut self.extras).insert((row as u32, col as u32), extras);
        }
        self.len += 1;
    }

    /// Lift a cell's stored pieces out, leaving the position empty. Pool
    /// references stay with the returned pieces.
    fn take(&mut self, row: usize, col: usize) -> Option<Raw> {
        let (idx, off) = split(row);
        let column = self.columns.get_mut(col)?;
        let i = column.chunks.binary_search_by_key(&idx, |c| c.0).ok()?;
        let chunk = Arc::make_mut(&mut column.chunks[i].1);
        let format = chunk.formats.get(off);
        let slot = chunk.cells.remove(off, chunk.count as usize - 1)?;
        chunk.count -= 1;
        let extras = if chunk.extras > 0 {
            let e = Arc::make_mut(&mut self.extras).remove(&(row as u32, col as u32));
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
        self.remap_rows(at, |r| if r >= at { (r + count < limit).then_some(r + count) } else { Some(r) });
    }

    /// Delete `count` rows from `start`; cells below move up.
    pub fn delete_rows(&mut self, start: usize, count: usize) {
        let end = start + count;
        self.remap_rows(start, |r| {
            if r < start {
                Some(r)
            } else if r < end {
                None
            } else {
                Some(r - count)
            }
        });
    }

    /// Move columns at or right of `at` right by `count`, dropping any that
    /// would land at or past `limit`. Columns move as whole containers; no
    /// cell is touched unless its column is dropped.
    pub fn insert_cols(&mut self, at: usize, count: usize, limit: usize) {
        if at >= self.columns.len() {
            return;
        }
        let cut = at.max(limit.saturating_sub(count));
        if cut < self.columns.len() {
            let dropped = self.columns.split_off(cut);
            for (i, column) in dropped.into_iter().enumerate() {
                self.drop_column(cut + i, column);
            }
        }
        self.rekey_extras_cols(|c| if c >= at { c + count } else { c });
        let at = at.min(self.columns.len());
        self.columns.splice(at..at, std::iter::repeat_with(Column::default).take(count));
    }

    /// Delete `count` columns from `start`; columns to the right move left
    /// as whole containers.
    pub fn delete_cols(&mut self, start: usize, count: usize) {
        if start >= self.columns.len() {
            return;
        }
        let end = (start + count).min(self.columns.len());
        let removed: Vec<Column> = self.columns.drain(start..end).collect();
        for (i, column) in removed.into_iter().enumerate() {
            self.drop_column(start + i, column);
        }
        self.rekey_extras_cols(|c| if c >= start + count { c - count } else { c });
    }

    /// Re-row every column through `to`, which must keep the order of the
    /// rows it keeps (true of row inserts and deletes) and leave rows above
    /// `first_moved` where they are. Chunks wholly above it are kept as they
    /// are, still shared with any clone. The rest of each column is rebuilt
    /// alone, streaming its cells into new chunks while the old chunks are
    /// dropped, so the extra memory is bounded by one column, not the sheet.
    fn remap_rows(&mut self, first_moved: usize, to: impl Fn(usize) -> Option<usize>) {
        let kept = |idx: u32| join(idx + 1, 0) <= first_moved;
        for col in 0..self.columns.len() {
            let old = std::mem::take(&mut self.columns[col].chunks);
            if old.is_empty() {
                continue;
            }
            // This column's metadata, keyed by old row, to follow its cells.
            let mut extras: HashMap<usize, CellExtras> = HashMap::new();
            for (idx, chunk) in old.iter().filter(|(idx, c)| c.extras > 0 && !kept(*idx)) {
                for (off, _) in chunk.cells.slots() {
                    let row = join(*idx, off);
                    if let Some(e) = Arc::make_mut(&mut self.extras).remove(&(row as u32, col as u32)) {
                        extras.insert(row, e);
                    }
                }
            }
            let mut built: Vec<(u32, Arc<Chunk>)> = Vec::new();
            for (idx, chunk) in old {
                if kept(idx) {
                    built.push((idx, chunk));
                    continue;
                }
                for (off, slot) in chunk.cells.slots() {
                    let row = join(idx, off);
                    let extra = extras.remove(&row);
                    match to(row) {
                        Some(new_row) => {
                            append(&mut built, new_row, slot, chunk.formats.get(off));
                            if let Some(e) = extra {
                                Arc::make_mut(&mut built.last_mut().expect("just appended").1).extras += 1;
                                Arc::make_mut(&mut self.extras).insert((new_row as u32, col as u32), e);
                            }
                        }
                        None => {
                            self.release_slot(slot);
                            self.len -= 1;
                        }
                    }
                }
            }
            self.columns[col].chunks = built;
        }
    }

    /// Give back everything a removed column held.
    fn drop_column(&mut self, col: usize, column: Column) {
        for (idx, chunk) in column.chunks {
            for (off, slot) in chunk.cells.slots() {
                self.release_slot(slot);
                self.len -= 1;
                if chunk.extras > 0 {
                    Arc::make_mut(&mut self.extras).remove(&(join(idx, off) as u32, col as u32));
                }
            }
        }
    }

    /// Move metadata keys to follow their columns.
    fn rekey_extras_cols(&mut self, to: impl Fn(usize) -> usize) {
        if self.extras.is_empty() {
            return;
        }
        let old = Arc::unwrap_or_clone(std::mem::take(&mut self.extras));
        self.extras = Arc::new(old.into_iter().map(|((r, c), e)| ((r, to(c as usize) as u32), e)).collect());
    }

    fn release_slot(&mut self, slot: Slot) {
        match slot {
            Slot::Text(id) => self.strings.release(id),
            Slot::Formula(id) => {
                self.formulas.remove(id);
            }
            Slot::Empty | Slot::Number(_) => {}
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
                    store.insert_replacing(row as usize, col as usize, cell);
                }
                Ok(store)
            }
            fn visit_seq<A: serde::de::SeqAccess<'de>>(self, mut seq: A) -> Result<ColumnStore, A::Error> {
                let mut store = ColumnStore::default();
                while let Some(((row, col), cell)) = seq.next_element::<((u32, u32), Cell)>()? {
                    store.insert_replacing(row as usize, col as usize, cell);
                }
                Ok(store)
            }
        }
        deserializer.deserialize_any(Cells)
    }
}
