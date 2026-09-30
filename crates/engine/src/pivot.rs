//! Pivot tables: definition types and the pure aggregation core.
//!
//! Spec: Obsidian `Projects/Visi/VisiGrid Pivot Tables Spec.md` (v1 scope).
//!
//! This module has no knowledge of the grid, ownership or undo. It turns a
//! snapshot of source columns plus a [`PivotDefinition`] into a dense output
//! grid ([`PivotOutput`]). Placement, protection and persistence live
//! elsewhere and consume this output.
//!
//! # Semantics (v1)
//!
//! - **Group keys are typed.** A number, a text value, a boolean, an error and
//!   a blank are different keys even when they display alike. Text groups
//!   case-insensitively (as Excel does); the label shown is the first spelling
//!   met in source order.
//! - **Item order** within a field: numbers ascending, then text
//!   (case-insensitive, then by exact spelling), then FALSE, TRUE, then errors
//!   (by code), then blank last (shown as `(blank)`). Deterministic for any
//!   input order.
//! - **Sum / Average / Min / Max** use numeric values only. Text — including
//!   numeric-looking text — booleans and blanks are ignored. If any source value
//!   in the group is an error, the result is that error (the first one in
//!   source order), not a generic `#VALUE!`.
//! - **Count** counts non-empty entries of the value field: numbers, text,
//!   booleans and errors. This differs from the worksheet `COUNT`, which counts
//!   numbers only; it matches a pivot's "Count of".
//! - **Empty groups.** A group with no numeric values gives 0 for Sum, Min and
//!   Max and `#DIV/0!` for Average. A row/column intersection with no source
//!   rows at all is left blank.
//! - **Grand totals** are computed from the underlying accumulators (sum,
//!   numeric count, min, max, non-empty count), never from displayed subtotals,
//!   so a grand average is sum ÷ count over all contributing rows.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use serde::{Deserialize, Serialize};

use crate::cell::{CellBorder, CellFormat, NegativeStyle, NumberFormat};
use crate::formula::eval::Value;
use crate::sheet::{Sheet, SheetId};

/// Output larger than this many cells (including headers and totals) is
/// refused before any dense grid is allocated. Engineering budget; revisit
/// with the fixture measurements (spec, "Performance and limits").
pub const MAX_OUTPUT_CELLS: usize = 250_000;

/// Output wider than this is refused. A column field with thousands of items
/// is almost always a mistake (a date or an id as the column field).
pub const MAX_OUTPUT_COLS: usize = 2_000;

/// Distinct Count holds a set of keys per output cell and per total. Refuse
/// when the keys held across all of them would exceed this.
pub const MAX_DISTINCT_KEYS: usize = 4_000_000;

/// Label used for blank group keys.
pub const BLANK_LABEL: &str = "(blank)";
/// Label of the grand-total row and column.
pub const GRAND_TOTAL_LABEL: &str = "Grand Total";

/// How a value field is summarized.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Aggregation {
    Sum,
    Count,
    DistinctCount,
    Average,
    Min,
    Max,
}

impl Aggregation {
    pub const ALL: [Aggregation; 6] = [
        Aggregation::Sum,
        Aggregation::Count,
        Aggregation::DistinctCount,
        Aggregation::Average,
        Aggregation::Min,
        Aggregation::Max,
    ];

    /// Counts display as whole numbers whatever the source's format.
    pub fn is_count(self) -> bool {
        matches!(self, Aggregation::Count | Aggregation::DistinctCount)
    }

    /// Parse a user-typed aggregation name: `sum`, `count`, `distinct`
    /// (or `distinct_count`, `countd`), `avg`/`average`/`mean`, `min`, `max`.
    pub fn parse(name: &str) -> Option<Aggregation> {
        match name.trim().to_ascii_lowercase().replace(['-', ' '], "_").as_str() {
            "sum" => Some(Aggregation::Sum),
            "count" => Some(Aggregation::Count),
            "distinct" | "distinct_count" | "distinctcount" | "countd" => Some(Aggregation::DistinctCount),
            "avg" | "average" | "mean" => Some(Aggregation::Average),
            "min" => Some(Aggregation::Min),
            "max" => Some(Aggregation::Max),
            _ => None,
        }
    }

    /// Display name used in value-field headers ("Sum of Amount").
    pub fn label(self) -> &'static str {
        match self {
            Aggregation::Sum => "Sum",
            Aggregation::Count => "Count",
            Aggregation::DistinctCount => "Distinct Count",
            Aggregation::Average => "Average",
            Aggregation::Min => "Min",
            Aggregation::Max => "Max",
        }
    }
}

/// A source field, identified by its column offset within the source range.
/// The header text is recorded so a refresh can detect that the column it
/// meant has moved or been renamed, instead of silently using a different one.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PivotField {
    /// 0-based column offset from the source's first column.
    pub offset: u32,
    /// Header text at definition time.
    pub header: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotValueField {
    pub field: PivotField,
    pub aggregation: Aggregation,
    /// Number format for this field's results. It belongs to the field, so it
    /// follows the results when the layout changes. `None` = General.
    #[serde(default)]
    pub number_format: Option<NumberFormat>,
}

/// The default number format for a new value field: counts are whole
/// numbers (never currency); other aggregations take the source column's
/// format when it has one.
pub fn default_number_format(aggregation: Aggregation, source: &NumberFormat) -> Option<NumberFormat> {
    if aggregation.is_count() {
        return Some(NumberFormat::Number { decimals: 0, thousands: true, negative: NegativeStyle::default() });
    }
    match source {
        NumberFormat::General => None,
        other => Some(other.clone()),
    }
}

impl PivotValueField {
    /// "Sum of Amount".
    pub fn header(&self) -> String {
        format!("{} of {}", self.aggregation.label(), self.field.header)
    }
}

/// Which fields go where. Row fields are ordered (outermost first).
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct PivotDefinition {
    #[serde(default)]
    pub rows: Vec<PivotField>,
    #[serde(default)]
    pub column: Option<PivotField>,
    #[serde(default)]
    pub values: Vec<PivotValueField>,
}

impl PivotDefinition {
    /// Distinct source column offsets the definition reads, ascending.
    pub fn used_offsets(&self) -> Vec<u32> {
        let mut v: Vec<u32> = self
            .rows
            .iter()
            .chain(self.column.iter())
            .map(|f| f.offset)
            .chain(self.values.iter().map(|v| v.field.offset))
            .collect();
        v.sort_unstable();
        v.dedup();
        v
    }

    pub fn is_empty(&self) -> bool {
        self.rows.is_empty() && self.column.is_none() && self.values.is_empty()
    }

    /// Build a definition from header names, for callers that name fields
    /// rather than pick them (the CLI, session clients, agents). Names match
    /// headers case-insensitively after trimming. A value field without an
    /// aggregation gets the desktop's default (Sum for a numeric column, Count
    /// otherwise), and every value field gets the desktop's default number
    /// format, from `profile` (one entry per source column, see
    /// [`column_profile`]).
    pub fn from_names(
        headers: &[String],
        rows: &[String],
        column: Option<&str>,
        values: &[(Option<Aggregation>, String)],
        profile: &[ColumnProfile],
    ) -> Result<PivotDefinition, String> {
        let field = |name: &str| -> Result<PivotField, String> {
            let want = name.trim().to_lowercase();
            headers
                .iter()
                .position(|h| h.trim().to_lowercase() == want)
                .map(|i| PivotField { offset: i as u32, header: headers[i].trim().to_string() })
                .ok_or_else(|| {
                    let known: Vec<&str> = headers.iter().map(|h| h.trim()).filter(|h| !h.is_empty()).collect();
                    format!("no column headed \"{}\" (columns: {})", name.trim(), known.join(", "))
                })
        };
        let def = PivotDefinition {
            rows: rows.iter().map(|r| field(r)).collect::<Result<_, _>>()?,
            column: column.map(field).transpose()?,
            values: values
                .iter()
                .map(|(aggregation, name)| {
                    let field = field(name)?;
                    let p = profile.get(field.offset as usize);
                    let aggregation = aggregation
                        .unwrap_or(if p.is_some_and(|p| p.numeric) { Aggregation::Sum } else { Aggregation::Count });
                    let number_format = default_number_format(aggregation, p.map_or(&NumberFormat::General, |p| &p.format));
                    Ok(PivotValueField { field, aggregation, number_format })
                })
                .collect::<Result<_, String>>()?,
        };
        if def.is_empty() {
            return Err(PivotError::NoFields.to_string());
        }
        Ok(def)
    }
}

/// The source rectangle, inclusive. `start_row` is the header row.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct PivotSource {
    pub sheet_id: SheetId,
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

impl PivotSource {
    pub fn width(&self) -> u32 {
        self.end_col - self.start_col + 1
    }
    /// Number of data rows (excluding the header).
    pub fn data_rows(&self) -> u32 {
        self.end_row - self.start_row
    }
}

/// When a pivot was last refreshed and against what.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RefreshRecord {
    /// Source data rows read.
    pub source_rows: u64,
    /// Seconds since the Unix epoch.
    pub refreshed_at: i64,
}

/// A pivot table, stored on the sheet that shows its output.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotTable {
    /// Stable, workbook-unique id.
    pub id: u64,
    /// Display name, e.g. "PivotTable1".
    pub name: String,
    pub source: PivotSource,
    pub definition: PivotDefinition,
    /// Top-left output cell.
    pub anchor_row: u32,
    pub anchor_col: u32,
    /// Size (rows, cols) of the last committed output. `None` until the first
    /// successful refresh. The owned region is exactly this rectangle.
    #[serde(default)]
    pub extent: Option<(u32, u32)>,
    #[serde(default)]
    pub last_refresh: Option<RefreshRecord>,
    /// The source may have changed since the last refresh. Persisted, so a
    /// reopened workbook never claims to be fresh.
    #[serde(default)]
    pub stale: bool,
    /// Runtime only: the source sheet's edit generation at the last refresh.
    #[serde(skip)]
    pub source_generation: Option<u64>,
}

impl PivotTable {
    /// The owned rectangle, inclusive: (start_row, start_col, end_row, end_col).
    pub fn region(&self) -> Option<(usize, usize, usize, usize)> {
        let (h, w) = self.extent?;
        if h == 0 || w == 0 {
            return None;
        }
        let (r, c) = (self.anchor_row as usize, self.anchor_col as usize);
        Some((r, c, r + h as usize - 1, c + w as usize - 1))
    }

    pub fn contains(&self, row: usize, col: usize) -> bool {
        self.region().is_some_and(|(r0, c0, r1, c1)| row >= r0 && row <= r1 && col >= c0 && col <= c1)
    }

    /// Does the owned region intersect this rectangle (inclusive)?
    pub fn intersects(&self, r0: usize, c0: usize, r1: usize, c1: usize) -> bool {
        self.region().is_some_and(|(a0, b0, a1, b1)| a0 <= r1 && r0 <= a1 && b0 <= c1 && c0 <= b1)
    }
}

/// Snapshot of the source handed to [`aggregate`]: header texts for every
/// source column, and the data rows of the columns the definition uses.
pub struct PivotSnapshot {
    /// One header per source column (index = offset).
    pub headers: Vec<String>,
    /// Data values keyed by source column offset. Every column holds the same
    /// number of rows.
    pub columns: HashMap<u32, Vec<Value>>,
    pub row_count: usize,
}

/// Why a definition cannot be refreshed against a snapshot.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotError {
    NoFields,
    /// A field's offset is outside the source.
    FieldOutOfRange { header: String },
    /// The header at a field's offset no longer matches the recorded header.
    FieldChanged { expected: String, found: String },
    /// The source header row has an empty or duplicate header.
    BadHeader { offset: u32, reason: String },
    /// A field is used twice as a row/column field.
    DuplicateLayoutField { header: String },
    OutputTooLarge { rows: usize, cols: usize, cells: usize },
    OutputTooWide { cols: usize },
    /// Distinct Count would hold more keys than the budget.
    DistinctTooLarge { keys: usize },
}

impl std::fmt::Display for PivotError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PivotError::NoFields => write!(f, "Choose at least one row, column or value field."),
            PivotError::FieldOutOfRange { header } => write!(f, "Field \"{header}\" is no longer inside the source range."),
            PivotError::FieldChanged { expected, found } => write!(
                f,
                "The source column for \"{expected}\" is now headed \"{found}\". Update the field list or the source."
            ),
            PivotError::BadHeader { offset, reason } => write!(f, "Source column {} header: {reason}", offset + 1),
            PivotError::DuplicateLayoutField { header } => {
                write!(f, "\"{header}\" is used more than once as a row or column field.")
            }
            PivotError::OutputTooLarge { rows, cols, cells } => write!(
                f,
                "The pivot would produce {rows} × {cols} = {cells} cells, over the {MAX_OUTPUT_CELLS}-cell limit. Use fewer or coarser fields."
            ),
            PivotError::DistinctTooLarge { keys } => write!(
                f,
                "Distinct Count would track over {keys} distinct values, above the {MAX_DISTINCT_KEYS} limit. Use fewer groups or Count instead."
            ),
            PivotError::OutputTooWide { cols } => write!(
                f,
                "The pivot would be {cols} columns wide, over the {MAX_OUTPUT_COLS}-column limit. The column field has too many distinct items."
            ),
        }
    }
}

/// Check the source headers: every column header must be non-empty and
/// unique (case-insensitively), so a field name resolves to one column.
pub fn validate_headers(headers: &[String]) -> Result<(), PivotError> {
    let mut seen: HashMap<String, u32> = HashMap::new();
    for (i, h) in headers.iter().enumerate() {
        let t = h.trim();
        if t.is_empty() {
            return Err(PivotError::BadHeader { offset: i as u32, reason: "is empty; every source column needs a header".into() });
        }
        if let Some(first) = seen.insert(t.to_lowercase(), i as u32) {
            return Err(PivotError::BadHeader {
                offset: i as u32,
                reason: format!("\"{t}\" repeats the header of column {}", first + 1),
            });
        }
    }
    Ok(())
}

/// What the desktop's field list knows about a source column: the number
/// format and type of its first non-empty data cell (within the first 200
/// rows). Drives default aggregations and value formats everywhere a pivot is
/// created, so a pivot made by an agent or the CLI matches one made by hand.
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnProfile {
    pub format: NumberFormat,
    pub numeric: bool,
}

/// One [`ColumnProfile`] per source column. The caller passes the sheet
/// `source` names.
pub fn column_profile(sheet: &Sheet, source: &PivotSource) -> Vec<ColumnProfile> {
    (source.start_col..=source.end_col)
        .map(|c| {
            let first = (source.start_row + 1..=source.end_row.min(source.start_row + 200))
                .map(|r| (r as usize, c as usize))
                .find(|&(r, c)| !matches!(sheet.get_computed_value(r, c), Value::Empty));
            match first {
                Some((r, c)) => ColumnProfile {
                    format: sheet.get_format(r, c).number_format.clone(),
                    numeric: matches!(sheet.get_computed_value(r, c), Value::Number(_)),
                },
                None => ColumnProfile { format: NumberFormat::General, numeric: false },
            }
        })
        .collect()
}

/// Does this source column hold any number at all? Sum, Average, Min and Max
/// of a column without one are all zeros or errors, which is never what was
/// meant.
pub fn column_has_numbers(sheet: &Sheet, source: &PivotSource, offset: u32) -> bool {
    let c = (source.start_col + offset) as usize;
    (source.start_row as usize + 1..=source.end_row as usize).any(|r| matches!(sheet.get_computed_value(r, c), Value::Number(_)))
}

/// Read a snapshot of `source` from `sheet`: every header, plus the data rows
/// of the columns `def` uses, as computed values in data space (view filters
/// and sorting are ignored). The caller must pass the sheet `source` names.
///
/// Reads only the needed columns; it does not clone the sheet.
pub fn capture_snapshot(sheet: &Sheet, source: &PivotSource, def: &PivotDefinition) -> PivotSnapshot {
    let header_row = source.start_row as usize;
    let headers: Vec<String> = (0..source.width())
        .map(|o| sheet.get_display(header_row, (source.start_col + o) as usize).trim().to_string())
        .collect();
    let rows = source.data_rows() as usize;
    let mut columns = HashMap::new();
    for offset in def.used_offsets() {
        if offset >= source.width() {
            continue;
        }
        let col = (source.start_col + offset) as usize;
        let mut v = Vec::with_capacity(rows);
        for r in 0..rows {
            v.push(sheet.get_computed_value(header_row + 1 + r, col));
        }
        columns.insert(offset, v);
    }
    PivotSnapshot { headers, columns, row_count: rows }
}

// ---------------------------------------------------------------------------
// Group keys
// ---------------------------------------------------------------------------

/// A typed group key. Text is keyed case-insensitively.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum Key {
    Number(u64), // f64 bits, normalized (-0 → 0)
    Text(String), // lowercased
    Bool(bool),
    Error(String),
    Blank,
}

impl Key {
    fn from_value(v: &Value) -> Key {
        match v {
            Value::Empty => Key::Blank,
            Value::Text(s) if s.is_empty() => Key::Blank,
            Value::Number(n) => {
                let n = if *n == 0.0 { 0.0 } else { *n };
                Key::Number(n.to_bits())
            }
            Value::Text(s) => Key::Text(s.to_lowercase()),
            Value::Boolean(b) => Key::Bool(*b),
            Value::Error(e) => Key::Error(e.clone()),
        }
    }

    fn class(&self) -> u8 {
        match self {
            Key::Number(_) => 0,
            Key::Text(_) => 1,
            Key::Bool(_) => 2,
            Key::Error(_) => 3,
            Key::Blank => 4,
        }
    }
}

/// One distinct item of a field: its key and the value to display.
#[derive(Debug, Clone)]
struct Item {
    key: Key,
    label: Value,
}

fn cmp_items(a: &Item, b: &Item) -> Ordering {
    let c = a.key.class().cmp(&b.key.class());
    if c != Ordering::Equal {
        return c;
    }
    match (&a.key, &b.key) {
        (Key::Number(x), Key::Number(y)) => f64::from_bits(*x).total_cmp(&f64::from_bits(*y)),
        (Key::Text(x), Key::Text(y)) => x.cmp(y).then_with(|| label_text(&a.label).cmp(&label_text(&b.label))),
        (Key::Bool(x), Key::Bool(y)) => x.cmp(y),
        (Key::Error(x), Key::Error(y)) => x.cmp(y),
        _ => Ordering::Equal,
    }
}

fn label_text(v: &Value) -> String {
    match v {
        Value::Text(s) => s.clone(),
        _ => String::new(),
    }
}

/// Distinct items of one field, in first-seen order, with a key → index map.
#[derive(Default)]
struct ItemSet {
    items: Vec<Item>,
    index: HashMap<Key, u32>,
    /// Exact text spelling → item, so repeated text is looked up without
    /// lowercasing or allocating on every row.
    exact_text: HashMap<String, u32>,
}

impl ItemSet {
    fn intern(&mut self, v: &Value) -> u32 {
        if let Value::Text(s) = v {
            if let Some(&i) = self.exact_text.get(s.as_str()) {
                return i;
            }
            let i = self.intern_slow(v);
            self.exact_text.insert(s.clone(), i);
            return i;
        }
        self.intern_slow(v)
    }

    fn intern_slow(&mut self, v: &Value) -> u32 {
        let key = Key::from_value(v);
        if let Some(&i) = self.index.get(&key) {
            return i;
        }
        let label = match &key {
            Key::Blank => Value::Text(BLANK_LABEL.to_string()),
            _ => v.clone(),
        };
        let i = self.items.len() as u32;
        self.index.insert(key.clone(), i);
        self.items.push(Item { key, label });
        i
    }

    /// Sorted order: position → item index, and item index → rank.
    fn sorted(&self) -> (Vec<u32>, Vec<u32>) {
        let mut order: Vec<u32> = (0..self.items.len() as u32).collect();
        order.sort_by(|&a, &b| cmp_items(&self.items[a as usize], &self.items[b as usize]));
        let mut rank = vec![0u32; order.len()];
        for (pos, &i) in order.iter().enumerate() {
            rank[i as usize] = pos as u32;
        }
        (order, rank)
    }
}

// ---------------------------------------------------------------------------
// Accumulators
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default)]
struct Acc {
    sum: f64,
    numeric: u64,
    non_empty: u64,
    min: Option<f64>,
    max: Option<f64>,
    error: Option<String>,
    /// At least one source row contributed (even if every value was blank).
    touched: bool,
    /// Distinct non-empty keys, only for Distinct Count fields.
    distinct: Option<HashSet<Key>>,
}

impl Acc {
    fn new(agg: Aggregation) -> Acc {
        Acc {
            distinct: (agg == Aggregation::DistinctCount).then(HashSet::new),
            ..Acc::default()
        }
    }

    fn add(&mut self, v: &Value) {
        self.touched = true;
        if let Some(set) = &mut self.distinct {
            let key = Key::from_value(v);
            if key != Key::Blank {
                set.insert(key);
            }
        }
        match v {
            Value::Empty => {}
            Value::Text(s) if s.is_empty() => {}
            Value::Number(n) => {
                self.non_empty += 1;
                self.numeric += 1;
                self.sum += n;
                self.min = Some(self.min.map_or(*n, |m| m.min(*n)));
                self.max = Some(self.max.map_or(*n, |m| m.max(*n)));
            }
            Value::Text(_) | Value::Boolean(_) => self.non_empty += 1,
            Value::Error(e) => {
                self.non_empty += 1;
                if self.error.is_none() {
                    self.error = Some(e.clone());
                }
            }
        }
    }

    fn merge(&mut self, o: &Acc) {
        self.touched |= o.touched;
        if let (Some(a), Some(b)) = (&mut self.distinct, &o.distinct) {
            a.extend(b.iter().cloned());
        }
        self.sum += o.sum;
        self.numeric += o.numeric;
        self.non_empty += o.non_empty;
        self.min = match (self.min, o.min) {
            (Some(a), Some(b)) => Some(a.min(b)),
            (a, b) => a.or(b),
        };
        self.max = match (self.max, o.max) {
            (Some(a), Some(b)) => Some(a.max(b)),
            (a, b) => a.or(b),
        };
        if self.error.is_none() {
            self.error = o.error.clone();
        }
    }

    fn result(&self, agg: Aggregation) -> Value {
        if !self.touched {
            return Value::Empty;
        }
        if agg == Aggregation::Count {
            return Value::Number(self.non_empty as f64);
        }
        if agg == Aggregation::DistinctCount {
            return Value::Number(self.distinct.as_ref().map_or(0, |d| d.len()) as f64);
        }
        if let Some(e) = &self.error {
            return Value::Error(e.clone());
        }
        match agg {
            Aggregation::Sum => Value::Number(self.sum),
            Aggregation::Average => {
                if self.numeric == 0 {
                    Value::Error("#DIV/0!".to_string())
                } else {
                    Value::Number(self.sum / self.numeric as f64)
                }
            }
            Aggregation::Min => Value::Number(self.min.unwrap_or(0.0)),
            Aggregation::Max => Value::Number(self.max.unwrap_or(0.0)),
            Aggregation::Count | Aggregation::DistinctCount => unreachable!(),
        }
    }
}

static EMPTY_VALUE: Value = Value::Empty;

fn new_accs(def: &PivotDefinition) -> Vec<Acc> {
    def.values.iter().map(|v| Acc::new(v.aggregation)).collect()
}

fn get(col: &[Value], r: usize) -> &Value {
    col.get(r).unwrap_or(&EMPTY_VALUE)
}

// ---------------------------------------------------------------------------
// Creation defaults (shared by the desktop app and headless hosts)
// ---------------------------------------------------------------------------

/// Choose a readable format only where the source supplied no number format.
/// Store it on the value field so it follows subsequent layout changes.
pub fn format_new_pivot_values(definition: &mut PivotDefinition, output: &PivotOutput) {
    let mut fractional = vec![false; definition.values.len()];
    for row in output.cells.iter().skip(output.header_rows) {
        for (col, value) in row.iter().enumerate() {
            if let (Some(Some(field)), Value::Number(n)) = (output.value_columns.get(col), value) {
                if n.is_finite() && (n - n.round()).abs() > 1e-9 {
                    fractional[*field] = true;
                }
            }
        }
    }
    for (i, field) in definition.values.iter_mut().enumerate() {
        if field.number_format.as_ref().is_none_or(|f| matches!(f, NumberFormat::General)) {
            let count = matches!(field.aggregation, Aggregation::Count | Aggregation::DistinctCount);
            let decimals = if !count && (fractional[i] || field.aggregation == Aggregation::Average) { 2 } else { 0 };
            field.number_format = Some(NumberFormat::Number { decimals, thousands: true, negative: Default::default() });
        }
    }
}

/// Creation-only defaults. The blank, styled sheet becomes the create action's
/// sheet snapshot, so redo restores the style without a workbook-sized copy.
/// Refresh deliberately leaves these formats and later user edits alone.
pub fn style_new_pivot(sheet: &mut Sheet, table: &PivotTable, output: &PivotOutput) {
    let (r0, c0) = (table.anchor_row as usize, table.anchor_col as usize);
    let border = CellBorder { color: Some([177, 192, 213, 255]), ..CellBorder::thin() };
    // Fill and foreground travel together, remaining legible in either theme.
    let header = CellFormat {
        bold: true,
        background_color: Some([232, 239, 250, 255]),
        font_color: Some([34, 53, 78, 255]),
        ..CellFormat::default()
    };
    for r in 0..output.header_rows {
        for c in 0..output.width() {
            let mut format = header.clone();
            if r + 1 == output.header_rows { format.border_bottom = border; }
            sheet.set_format(r0 + r, c0 + c, format);
        }
    }
    let has_total = !table.definition.values.is_empty() || table.definition.column.is_some();
    if has_total && output.height() > output.header_rows {
        let total = CellFormat {
            background_color: Some([241, 245, 251, 255]),
            border_top: border,
            ..header.clone()
        };
        for c in 0..output.width() {
            sheet.set_format(r0 + output.height() - 1, c0 + c, total.clone());
        }
    }
    if table.definition.column.is_some() {
        let first_total = output.width() - table.definition.values.len().max(1);
        for r in 0..output.height() {
            for c in first_total..output.width() {
                sheet.set_bold(r0 + r, c0 + c, true);
            }
            sheet.set_border_left(r0 + r, c0 + first_total, border);
        }
    }
}

// ---------------------------------------------------------------------------
// Output
// ---------------------------------------------------------------------------

/// Dense tabular output. `cells[r][c]`; header rows first.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotOutput {
    pub cells: Vec<Vec<Value>>,
    /// Number of leading header rows (1 without a column field, 2 with).
    pub header_rows: usize,
    /// Number of distinct row groups (excluding the grand-total row).
    pub row_groups: usize,
    /// Number of distinct column items (0 without a column field).
    pub column_items: usize,
    /// Source data rows read.
    pub source_rows: usize,
    /// For each output column, the value field whose results it holds (data
    /// rows and the grand-total row), or `None` for label columns.
    pub value_columns: Vec<Option<usize>>,
}

impl PivotOutput {
    pub fn height(&self) -> usize {
        self.cells.len()
    }
    pub fn width(&self) -> usize {
        self.cells.first().map_or(0, |r| r.len())
    }
}

/// Validate a definition against a snapshot's headers.
pub fn validate(def: &PivotDefinition, headers: &[String]) -> Result<(), PivotError> {
    if def.is_empty() {
        return Err(PivotError::NoFields);
    }
    validate_headers(headers)?;
    let check = |f: &PivotField| -> Result<(), PivotError> {
        match headers.get(f.offset as usize) {
            None => Err(PivotError::FieldOutOfRange { header: f.header.clone() }),
            Some(h) if h.trim() != f.header.trim() => {
                Err(PivotError::FieldChanged { expected: f.header.clone(), found: h.clone() })
            }
            Some(_) => Ok(()),
        }
    };
    let mut layout_offsets = Vec::new();
    for f in def.rows.iter().chain(def.column.iter()) {
        check(f)?;
        if layout_offsets.contains(&f.offset) {
            return Err(PivotError::DuplicateLayoutField { header: f.header.clone() });
        }
        layout_offsets.push(f.offset);
    }
    for v in &def.values {
        check(&v.field)?;
    }
    Ok(())
}

/// Compute the pivot. Pure: no grid access, no allocation proportional to the
/// output until the output size has been checked against the budgets.
pub fn aggregate(def: &PivotDefinition, snap: &PivotSnapshot) -> Result<PivotOutput, PivotError> {
    validate(def, &snap.headers)?;
    let n = snap.row_count;
    let empty: Vec<Value> = Vec::new();
    let col_of = |offset: u32| -> &Vec<Value> { snap.columns.get(&offset).unwrap_or(&empty) };

    // 1. Intern row-field items and the column-field items.
    let mut row_sets: Vec<ItemSet> = def.rows.iter().map(|_| ItemSet::default()).collect();
    let mut col_set = ItemSet::default();
    let row_cols: Vec<&Vec<Value>> = def.rows.iter().map(|f| col_of(f.offset)).collect();
    let col_col: Option<&Vec<Value>> = def.column.as_ref().map(|f| col_of(f.offset));
    let val_cols: Vec<&Vec<Value>> = def.values.iter().map(|v| col_of(v.field.offset)).collect();

    // Row key tuple (item indices) → group index; accumulators are
    // [group][col_item + 1 (0 = row total)][value field].
    let mut group_index: HashMap<Vec<u32>, usize> = HashMap::new();
    let mut groups: Vec<Vec<u32>> = Vec::new();
    // Sparse cell accumulators: (group, col_item) → per-value accumulators.
    let mut cell_acc: HashMap<(usize, u32), Vec<Acc>> = HashMap::new();
    let nv = def.values.len();

    let mut tuple: Vec<u32> = Vec::with_capacity(def.rows.len());
    for r in 0..n {
        tuple.clear();
        for (i, c) in row_cols.iter().enumerate() {
            tuple.push(row_sets[i].intern(get(c, r)));
        }
        let g = match group_index.get(&tuple) {
            Some(&g) => g,
            None => {
                let g = groups.len();
                groups.push(tuple.clone());
                group_index.insert(tuple.clone(), g);
                g
            }
        };
        let ci = match col_col {
            Some(c) => col_set.intern(get(c, r)),
            None => 0,
        };
        let accs = cell_acc.entry((g, ci)).or_insert_with(|| new_accs(def));
        for (vi, vc) in val_cols.iter().enumerate() {
            accs[vi].add(get(vc, r));
        }
    }

    // 2. Output size, before allocating.
    let has_col = def.column.is_some();
    let n_col_items = if has_col { col_set.items.len() } else { 0 };
    let label_cols = def.rows.len().max(if has_col || nv > 0 { 1 } else { 0 });
    let value_block = nv.max(if has_col { 1 } else { 0 });
    let data_cols = if has_col { (n_col_items + 1) * value_block } else { nv };
    let width = (label_cols + data_cols).max(1);
    let header_rows = if has_col { 2 } else { 1 };
    let body_rows = if def.rows.is_empty() { 0 } else { groups.len() };
    let total_row = nv > 0 || has_col;
    let height = header_rows + body_rows + usize::from(total_row);
    if width > MAX_OUTPUT_COLS {
        return Err(PivotError::OutputTooWide { cols: width });
    }
    let cells_count = width.saturating_mul(height);
    if cells_count > MAX_OUTPUT_CELLS {
        return Err(PivotError::OutputTooLarge { rows: height, cols: width, cells: cells_count });
    }

    // 3. Sort groups by each row field's item rank, outermost first.
    let row_ranks: Vec<Vec<u32>> = row_sets.iter().map(|s| s.sorted().1).collect();
    let mut group_order: Vec<usize> = (0..groups.len()).collect();
    group_order.sort_by(|&a, &b| {
        for (fi, ranks) in row_ranks.iter().enumerate() {
            let c = ranks[groups[a][fi] as usize].cmp(&ranks[groups[b][fi] as usize]);
            if c != Ordering::Equal {
                return c;
            }
        }
        Ordering::Equal
    });
    let (col_order, _) = col_set.sorted();

    // 4. Totals from underlying accumulators.
    let mut row_totals: Vec<Vec<Acc>> = (0..groups.len()).map(|_| new_accs(def)).collect();
    let mut col_totals: HashMap<u32, Vec<Acc>> = HashMap::new();
    let mut grand: Vec<Acc> = new_accs(def);
    for (&(g, ci), accs) in &cell_acc {
        for vi in 0..nv {
            row_totals[g][vi].merge(&accs[vi]);
            col_totals.entry(ci).or_insert_with(|| new_accs(def))[vi].merge(&accs[vi]);
            grand[vi].merge(&accs[vi]);
        }
    }
    let distinct_keys: usize = cell_acc
        .values()
        .chain(row_totals.iter())
        .chain(col_totals.values())
        .chain(std::iter::once(&grand))
        .flat_map(|accs| accs.iter())
        .map(|a| a.distinct.as_ref().map_or(0, |d| d.len()))
        .sum();
    if distinct_keys > MAX_DISTINCT_KEYS {
        return Err(PivotError::DistinctTooLarge { keys: distinct_keys });
    }
    if n > 0 {
        // A grand total over zero rows stays blank; over ≥1 row it is touched.
        for a in grand.iter_mut() {
            a.touched = true;
        }
    }

    // 5. Lay out.
    let mut out: Vec<Vec<Value>> = vec![vec![Value::Empty; width]; height];
    let text = |s: &str| Value::Text(s.to_string());

    // Headers.
    let hdr = header_rows - 1; // the row carrying field names
    for (i, f) in def.rows.iter().enumerate() {
        out[hdr][i] = text(&f.header);
    }
    if has_col {
        // Two header rows, as Excel lays out a cross-tab:
        // - one value field: row 0 = value header, then the column field's
        //   name over its items; row 1 = row-field names, item labels,
        //   "Grand Total".
        // - several value fields: row 0 = column field's name, then each item
        //   label (and "Grand Total") over its block; row 1 = row-field names,
        //   then the value headers inside each block ("Total …" in the last).
        let col_field = def.column.as_ref().unwrap();
        if nv <= 1 {
            if nv == 1 {
                out[0][0] = text(&def.values[0].header());
            }
            out[0][label_cols] = text(&col_field.header);
            for (pos, &ci) in col_order.iter().enumerate() {
                out[1][label_cols + pos] = col_set.items[ci as usize].label.clone();
            }
            out[1][label_cols + n_col_items] = text(GRAND_TOTAL_LABEL);
        } else {
            out[0][0] = text(&col_field.header);
            for (pos, &ci) in col_order.iter().enumerate() {
                out[0][label_cols + pos * value_block] = col_set.items[ci as usize].label.clone();
                for vb in 0..value_block {
                    out[1][label_cols + pos * value_block + vb] = text(&def.values[vb].header());
                }
            }
            let tot = label_cols + n_col_items * value_block;
            out[0][tot] = text(GRAND_TOTAL_LABEL);
            for vb in 0..value_block {
                out[1][tot + vb] = text(&format!("Total {}", def.values[vb].header()));
            }
        }
    } else {
        for (vi, v) in def.values.iter().enumerate() {
            out[0][label_cols + vi] = text(&v.header());
        }
    }

    // Body.
    let value_at = |accs: Option<&Vec<Acc>>, vi: usize| -> Value {
        match accs {
            Some(a) => a[vi].result(def.values[vi].aggregation),
            None => Value::Empty,
        }
    };
    for (bi, &g) in group_order.iter().enumerate().take(body_rows) {
        let r = header_rows + bi;
        for (fi, &item) in groups[g].iter().enumerate() {
            out[r][fi] = row_sets[fi].items[item as usize].label.clone();
        }
        if has_col {
            for (pos, &ci) in col_order.iter().enumerate() {
                let accs = cell_acc.get(&(g, ci));
                for vi in 0..nv {
                    out[r][label_cols + pos * value_block + vi] = value_at(accs, vi);
                }
            }
            for vi in 0..nv {
                out[r][label_cols + n_col_items * value_block + vi] = row_totals[g][vi].result(def.values[vi].aggregation);
            }
        } else {
            for vi in 0..nv {
                out[r][label_cols + vi] = row_totals[g][vi].result(def.values[vi].aggregation);
            }
        }
    }

    // Grand-total row.
    if total_row {
        let r = height - 1;
        out[r][0] = text(GRAND_TOTAL_LABEL);
        if has_col {
            for (pos, &ci) in col_order.iter().enumerate() {
                let accs = col_totals.get(&ci);
                for vi in 0..nv {
                    out[r][label_cols + pos * value_block + vi] = value_at(accs, vi);
                }
            }
        }
        let base = if has_col { label_cols + n_col_items * value_block } else { label_cols };
        for vi in 0..nv {
            out[r][base + vi] = grand[vi].result(def.values[vi].aggregation);
        }
    }

    let mut value_columns: Vec<Option<usize>> = vec![None; width];
    if nv > 0 {
        for (c, slot) in value_columns.iter_mut().enumerate().skip(label_cols) {
            *slot = Some((c - label_cols) % value_block.max(1));
        }
    }

    Ok(PivotOutput {
        cells: out,
        header_rows,
        row_groups: body_rows,
        column_items: n_col_items,
        source_rows: n,
        value_columns,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn t(s: &str) -> Value {
        Value::Text(s.to_string())
    }
    fn n(x: f64) -> Value {
        Value::Number(x)
    }
    fn e(s: &str) -> Value {
        Value::Error(s.to_string())
    }
    fn f(offset: u32, header: &str) -> PivotField {
        PivotField { offset, header: header.to_string() }
    }
    fn v(offset: u32, header: &str, aggregation: Aggregation) -> PivotValueField {
        PivotValueField { field: f(offset, header), aggregation, number_format: None }
    }

    /// Build a snapshot from header names and row-major data.
    fn snap(headers: &[&str], rows: Vec<Vec<Value>>) -> PivotSnapshot {
        let mut columns: HashMap<u32, Vec<Value>> = HashMap::new();
        for c in 0..headers.len() {
            columns.insert(c as u32, rows.iter().map(|r| r.get(c).cloned().unwrap_or(Value::Empty)).collect());
        }
        PivotSnapshot {
            headers: headers.iter().map(|s| s.to_string()).collect(),
            columns,
            row_count: rows.len(),
        }
    }

    fn sales() -> PivotSnapshot {
        snap(
            &["Region", "Month", "Amount", "Rep"],
            vec![
                vec![t("West"), t("Jan"), n(100.0), t("Ann")],
                vec![t("East"), t("Jan"), n(50.0), t("Bo")],
                vec![t("West"), t("Feb"), n(25.0), t("Ann")],
                vec![t("west"), t("Jan"), n(10.0), t("Cy")],
                vec![t("East"), t("Feb"), n(5.0), t("Bo")],
            ],
        )
    }

    #[test]
    fn rows_and_sum_with_case_insensitive_text_and_grand_total() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: None,
            values: vec![v(2, "Amount", Aggregation::Sum)],
        };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(
            out.cells,
            vec![
                vec![t("Region"), t("Sum of Amount")],
                vec![t("East"), n(55.0)],
                vec![t("West"), n(135.0)], // "west" joins "West"; first spelling shown
                vec![t(GRAND_TOTAL_LABEL), n(190.0)],
            ]
        );
    }

    #[test]
    fn cross_tab_with_column_field_and_totals() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: Some(f(1, "Month")),
            values: vec![v(2, "Amount", Aggregation::Sum)],
        };
        let out = aggregate(&def, &sales()).unwrap();
        // Months sort as text: Feb, Jan.
        assert_eq!(out.header_rows, 2);
        assert_eq!(out.cells[0], vec![t("Sum of Amount"), t("Month"), Value::Empty, Value::Empty]);
        assert_eq!(out.cells[1], vec![t("Region"), t("Feb"), t("Jan"), t(GRAND_TOTAL_LABEL)]);
        assert_eq!(out.cells[2], vec![t("East"), n(5.0), n(50.0), n(55.0)]);
        assert_eq!(out.cells[3], vec![t("West"), n(25.0), n(110.0), n(135.0)]);
        assert_eq!(out.cells[4], vec![t(GRAND_TOTAL_LABEL), n(30.0), n(160.0), n(190.0)]);
    }

    #[test]
    fn multiple_values_with_column_field() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: Some(f(1, "Month")),
            values: vec![v(2, "Amount", Aggregation::Sum), v(2, "Amount", Aggregation::Count)],
        };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(out.width(), 1 + 3 * 2);
        assert_eq!(
            out.cells[0],
            vec![t("Month"), t("Feb"), Value::Empty, t("Jan"), Value::Empty, t(GRAND_TOTAL_LABEL), Value::Empty]
        );
        assert_eq!(
            out.cells[1],
            vec![
                t("Region"),
                t("Sum of Amount"),
                t("Count of Amount"),
                t("Sum of Amount"),
                t("Count of Amount"),
                t("Total Sum of Amount"),
                t("Total Count of Amount"),
            ]
        );
        // West: Feb sum 25 count 1, Jan sum 110 count 2, totals 135 / 3.
        assert_eq!(out.cells[3], vec![t("West"), n(25.0), n(1.0), n(110.0), n(2.0), n(135.0), n(3.0)]);
    }

    #[test]
    fn missing_intersection_is_blank_not_zero() {
        let s = snap(
            &["R", "C", "V"],
            vec![vec![t("a"), t("x"), n(1.0)], vec![t("b"), t("y"), n(2.0)]],
        );
        let def = PivotDefinition { rows: vec![f(0, "R")], column: Some(f(1, "C")), values: vec![v(2, "V", Aggregation::Sum)] };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[2], vec![t("a"), n(1.0), Value::Empty, n(1.0)]);
        assert_eq!(out.cells[3], vec![t("b"), Value::Empty, n(2.0), n(2.0)]);
    }

    #[test]
    fn numeric_looking_text_is_not_summed_but_is_counted() {
        let s = snap(
            &["K", "V"],
            vec![vec![t("a"), n(10.0)], vec![t("a"), t("5")], vec![t("a"), Value::Empty], vec![t("a"), Value::Boolean(true)]],
        );
        let def = PivotDefinition {
            rows: vec![f(0, "K")],
            column: None,
            values: vec![v(1, "V", Aggregation::Sum), v(1, "V", Aggregation::Count), v(1, "V", Aggregation::Average)],
        };
        let out = aggregate(&def, &s).unwrap();
        // Sum ignores "5" and TRUE; Count counts 10, "5", TRUE (not the blank);
        // Average = 10 / 1 numeric.
        assert_eq!(out.cells[1], vec![t("a"), n(10.0), n(3.0), n(10.0)]);
    }

    #[test]
    fn grand_average_uses_underlying_rows_not_group_averages() {
        let s = snap(
            &["K", "V"],
            vec![vec![t("a"), n(10.0)], vec![t("b"), n(1.0)], vec![t("b"), n(1.0)], vec![t("b"), n(1.0)]],
        );
        let def = PivotDefinition { rows: vec![f(0, "K")], column: None, values: vec![v(1, "V", Aggregation::Average)] };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[1], vec![t("a"), n(10.0)]);
        assert_eq!(out.cells[2], vec![t("b"), n(1.0)]);
        // (10+1+1+1)/4 = 3.25, not (10+1)/2 = 5.5
        assert_eq!(out.cells[3], vec![t(GRAND_TOTAL_LABEL), n(3.25)]);
    }

    #[test]
    fn errors_surface_as_themselves_and_count_as_non_empty() {
        let s = snap(
            &["K", "V"],
            vec![vec![t("a"), n(1.0)], vec![t("a"), e("#N/A")], vec![t("b"), n(2.0)]],
        );
        let def = PivotDefinition {
            rows: vec![f(0, "K")],
            column: None,
            values: vec![v(1, "V", Aggregation::Sum), v(1, "V", Aggregation::Count)],
        };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[1], vec![t("a"), e("#N/A"), n(2.0)]);
        assert_eq!(out.cells[2], vec![t("b"), n(2.0), n(1.0)]);
        assert_eq!(out.cells[3], vec![t(GRAND_TOTAL_LABEL), e("#N/A"), n(3.0)]);
    }

    #[test]
    fn empty_numeric_group_rules() {
        let s = snap(&["K", "V"], vec![vec![t("a"), t("x")]]);
        let def = PivotDefinition {
            rows: vec![f(0, "K")],
            column: None,
            values: vec![
                v(1, "V", Aggregation::Sum),
                v(1, "V", Aggregation::Average),
                v(1, "V", Aggregation::Min),
                v(1, "V", Aggregation::Max),
            ],
        };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[1], vec![t("a"), n(0.0), e("#DIV/0!"), n(0.0), n(0.0)]);
    }

    #[test]
    fn typed_keys_do_not_merge_and_sort_numbers_text_bool_error_blank() {
        let s = snap(
            &["K", "V"],
            vec![
                vec![Value::Empty, n(1.0)],
                vec![e("#REF!"), n(1.0)],
                vec![Value::Boolean(true), n(1.0)],
                vec![Value::Boolean(false), n(1.0)],
                vec![t("10"), n(1.0)],
                vec![n(10.0), n(1.0)],
                vec![n(-2.0), n(1.0)],
                vec![t("b"), n(1.0)],
                vec![t("A"), n(1.0)],
            ],
        );
        let def = PivotDefinition { rows: vec![f(0, "K")], column: None, values: vec![v(1, "V", Aggregation::Sum)] };
        let out = aggregate(&def, &s).unwrap();
        let labels: Vec<Value> = out.cells[1..out.height() - 1].iter().map(|r| r[0].clone()).collect();
        assert_eq!(
            labels,
            vec![
                n(-2.0),
                n(10.0),
                t("10"), // text "10" is its own item, not merged with 10
                t("A"),
                t("b"),
                Value::Boolean(false),
                Value::Boolean(true),
                e("#REF!"),
                t(BLANK_LABEL),
            ]
        );
    }

    #[test]
    fn deterministic_regardless_of_input_order() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region"), f(3, "Rep")],
            column: Some(f(1, "Month")),
            values: vec![v(2, "Amount", Aggregation::Sum)],
        };
        let a = aggregate(&def, &sales()).unwrap();
        let mut rev = sales();
        for col in rev.columns.values_mut() {
            col.reverse();
        }
        let b = aggregate(&def, &rev).unwrap();
        // Labels use first-seen spelling ("West" vs "west" may differ); compare numbers.
        let nums = |o: &PivotOutput| -> Vec<Value> {
            o.cells.iter().flatten().filter(|c| matches!(c, Value::Number(_))).cloned().collect()
        };
        assert_eq!(nums(&a), nums(&b));
    }

    #[test]
    fn two_row_fields_nest_outer_then_inner() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region"), f(3, "Rep")],
            column: None,
            values: vec![v(2, "Amount", Aggregation::Sum)],
        };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(out.cells[0], vec![t("Region"), t("Rep"), t("Sum of Amount")]);
        assert_eq!(out.cells[1], vec![t("East"), t("Bo"), n(55.0)]);
        assert_eq!(out.cells[2], vec![t("West"), t("Ann"), n(125.0)]);
        assert_eq!(out.cells[3], vec![t("West"), t("Cy"), n(10.0)]);
        assert_eq!(out.cells[4], vec![t(GRAND_TOTAL_LABEL), Value::Empty, n(190.0)]);
    }

    #[test]
    fn rows_only_lists_distinct_items_without_total_row() {
        let def = PivotDefinition { rows: vec![f(0, "Region")], column: None, values: vec![] };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(out.cells, vec![vec![t("Region")], vec![t("East")], vec![t("West")]]);
    }

    #[test]
    fn values_only_gives_a_single_grand_total_row() {
        let def = PivotDefinition { rows: vec![], column: None, values: vec![v(2, "Amount", Aggregation::Max)] };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(out.cells, vec![vec![Value::Empty, t("Max of Amount")], vec![t(GRAND_TOTAL_LABEL), n(100.0)]]);
    }

    #[test]
    fn distinct_count_totals_deduplicate_across_groups() {
        // Customer "c1" buys in both regions: 2 distinct per region, 3 overall.
        let s = snap(
            &["Region", "Customer"],
            vec![
                vec![t("East"), t("c1")],
                vec![t("East"), t("c2")],
                vec![t("East"), t("C1")], // same customer, different case
                vec![t("West"), t("c1")],
                vec![t("West"), t("c3")],
                vec![t("West"), Value::Empty], // blank is not a customer
            ],
        );
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: None,
            values: vec![v(1, "Customer", Aggregation::DistinctCount), v(1, "Customer", Aggregation::Count)],
        };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[0][1], t("Distinct Count of Customer"));
        assert_eq!(out.cells[1], vec![t("East"), n(2.0), n(3.0)]);
        assert_eq!(out.cells[2], vec![t("West"), n(2.0), n(2.0)]);
        // Grand total is 3 distinct customers, not 2 + 2.
        assert_eq!(out.cells[3], vec![t(GRAND_TOTAL_LABEL), n(3.0), n(5.0)]);
    }

    #[test]
    fn distinct_count_column_totals_deduplicate_too() {
        let s = snap(
            &["R", "C", "K"],
            vec![
                vec![t("a"), t("x"), n(1.0)],
                vec![t("b"), t("x"), n(1.0)],
                vec![t("b"), t("y"), n(2.0)],
            ],
        );
        let def = PivotDefinition { rows: vec![f(0, "R")], column: Some(f(1, "C")), values: vec![v(2, "K", Aggregation::DistinctCount)] };
        let out = aggregate(&def, &s).unwrap();
        // Column x total: {1} → 1 (not 1 + 1); row b total: {1, 2} → 2; grand {1, 2} → 2.
        assert_eq!(out.cells[4], vec![t(GRAND_TOTAL_LABEL), n(1.0), n(1.0), n(2.0)]);
        assert_eq!(out.cells[3], vec![t("b"), n(1.0), n(1.0), n(2.0)]);
    }

    #[test]
    fn value_columns_map_output_columns_to_fields() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: Some(f(1, "Month")),
            values: vec![v(2, "Amount", Aggregation::Sum), v(2, "Amount", Aggregation::Count)],
        };
        let out = aggregate(&def, &sales()).unwrap();
        assert_eq!(out.value_columns, vec![None, Some(0), Some(1), Some(0), Some(1), Some(0), Some(1)]);
    }

    #[test]
    fn default_formats_counts_are_plain_others_inherit() {
        let cur = NumberFormat::Currency { decimals: 2, thousands: true, symbol: None, negative: NegativeStyle::default() };
        assert_eq!(default_number_format(Aggregation::Sum, &cur), Some(cur.clone()));
        assert!(matches!(default_number_format(Aggregation::Count, &cur), Some(NumberFormat::Number { decimals: 0, .. })));
        assert!(matches!(default_number_format(Aggregation::DistinctCount, &cur), Some(NumberFormat::Number { decimals: 0, .. })));
        assert_eq!(default_number_format(Aggregation::Average, &NumberFormat::General), None);
    }

    #[test]
    fn negative_values_min_max() {
        let s = snap(&["K", "V"], vec![vec![t("a"), n(-5.0)], vec![t("a"), n(-1.0)]]);
        let def = PivotDefinition {
            rows: vec![f(0, "K")],
            column: None,
            values: vec![v(1, "V", Aggregation::Min), v(1, "V", Aggregation::Max), v(1, "V", Aggregation::Sum)],
        };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells[1], vec![t("a"), n(-5.0), n(-1.0), n(-6.0)]);
    }

    #[test]
    fn header_validation_and_field_drift() {
        let s = snap(&["A", "a"], vec![]);
        let def = PivotDefinition { rows: vec![f(0, "A")], column: None, values: vec![] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::BadHeader { .. })));

        let s = snap(&["Region", "Amount"], vec![vec![t("x"), n(1.0)]]);
        let def = PivotDefinition { rows: vec![f(0, "Territory")], column: None, values: vec![] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::FieldChanged { .. })));

        let def = PivotDefinition { rows: vec![f(5, "Region")], column: None, values: vec![] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::FieldOutOfRange { .. })));

        let def = PivotDefinition { rows: vec![f(0, "Region")], column: Some(f(0, "Region")), values: vec![] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::DuplicateLayoutField { .. })));

        assert!(matches!(aggregate(&PivotDefinition::default(), &s), Err(PivotError::NoFields)));
    }

    #[test]
    fn expansion_guard_refuses_before_allocating() {
        // 3,000 distinct column items → wider than MAX_OUTPUT_COLS.
        let rows: Vec<Vec<Value>> = (0..3000).map(|i| vec![t("a"), n(i as f64), n(1.0)]).collect();
        let s = snap(&["R", "C", "V"], rows);
        let def = PivotDefinition { rows: vec![f(0, "R")], column: Some(f(1, "C")), values: vec![v(2, "V", Aggregation::Sum)] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::OutputTooWide { .. })));

        // 300,000 distinct row groups → over MAX_OUTPUT_CELLS.
        let rows: Vec<Vec<Value>> = (0..300_000).map(|i| vec![n(i as f64), n(1.0)]).collect();
        let s = snap(&["R", "V"], rows);
        let def = PivotDefinition { rows: vec![f(0, "R")], column: None, values: vec![v(1, "V", Aggregation::Sum)] };
        assert!(matches!(aggregate(&def, &s), Err(PivotError::OutputTooLarge { .. })));
    }

    #[test]
    fn empty_source_gives_headers_and_blank_total() {
        let s = snap(&["K", "V"], vec![]);
        let def = PivotDefinition { rows: vec![f(0, "K")], column: None, values: vec![v(1, "V", Aggregation::Sum)] };
        let out = aggregate(&def, &s).unwrap();
        assert_eq!(out.cells, vec![vec![t("K"), t("Sum of V")], vec![t(GRAND_TOTAL_LABEL), Value::Empty]]);
    }

    #[test]
    fn capture_reads_headers_and_only_used_columns_in_data_space() {
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        for (c, h) in ["Region", "Skip", "Amount"].iter().enumerate() {
            sheet.set_value(0, c, h);
        }
        sheet.set_value(1, 0, "West");
        sheet.set_value(1, 1, "ignored");
        sheet.set_value(1, 2, "12.5");
        sheet.set_value(2, 0, "East");
        sheet.set_value(2, 2, "'7"); // text, not a number
        let source = PivotSource { sheet_id: SheetId(1), start_row: 0, start_col: 0, end_row: 2, end_col: 2 };
        let def = PivotDefinition { rows: vec![f(0, "Region")], column: None, values: vec![v(2, "Amount", Aggregation::Sum)] };
        let snap = capture_snapshot(&sheet, &source, &def);
        assert_eq!(snap.headers, vec!["Region", "Skip", "Amount"]);
        assert_eq!(snap.row_count, 2);
        assert!(!snap.columns.contains_key(&1));
        assert_eq!(snap.columns[&2][0], n(12.5));
        let out = aggregate(&def, &snap).unwrap();
        assert_eq!(out.cells[1], vec![t("East"), n(0.0)]);
        assert_eq!(out.cells[2], vec![t("West"), n(12.5)]);
    }

    /// Spec fixtures ("Performance and limits"). Run with:
    /// `cargo test -p visigrid-engine --release --lib pivot::tests::bench -- --ignored --nocapture`
    #[test]
    #[ignore]
    fn bench_fixtures() {
        for rows in [50_000usize, 1_000_000] {
            let mut sheet = Sheet::new(SheetId(1), rows + 1, 8);
            let headers = ["Region", "Rep", "Month", "Amount", "Qty", "Sku", "Note", "Flag"];
            for (c, h) in headers.iter().enumerate() {
                sheet.set_value(0, c, h);
            }
            // 10 regions × 25 reps = 250 combined row keys; 12 months.
            for r in 0..rows {
                let row = r + 1;
                sheet.set_value(row, 0, &format!("Region {}", r % 10));
                sheet.set_value(row, 1, &format!("Rep {}", (r / 10) % 25));
                sheet.set_value(row, 2, &format!("M{:02}", r % 12 + 1));
                sheet.set_value(row, 3, &format!("{}.25", r % 997));
                sheet.set_value(row, 4, &format!("{}", r % 13));
                sheet.set_value(row, 5, &format!("SKU-{}", r % 5000));
                sheet.set_value(row, 6, "x");
                sheet.set_value(row, 7, if r % 2 == 0 { "TRUE" } else { "FALSE" });
            }
            let source = PivotSource { sheet_id: SheetId(1), start_row: 0, start_col: 0, end_row: rows as u32, end_col: 7 };
            let def = PivotDefinition {
                rows: vec![f(0, "Region"), f(1, "Rep")],
                column: Some(f(2, "Month")),
                values: vec![v(3, "Amount", Aggregation::Sum), v(4, "Qty", Aggregation::Average)],
            };
            let t0 = std::time::Instant::now();
            let snap = capture_snapshot(&sheet, &source, &def);
            let t1 = std::time::Instant::now();
            let out = aggregate(&def, &snap).unwrap();
            let t2 = std::time::Instant::now();
            eprintln!(
                "rows={rows:>9} capture={:>7.1}ms aggregate={:>7.1}ms output={}x{} ({} cells, {} groups)",
                (t1 - t0).as_secs_f64() * 1000.0,
                (t2 - t1).as_secs_f64() * 1000.0,
                out.height(),
                out.width(),
                out.height() * out.width(),
                out.row_groups,
            );
            assert_eq!(out.row_groups, 250);
            assert_eq!(out.column_items, 12);
        }
    }

    #[test]
    fn definition_serde_round_trip() {
        let def = PivotDefinition {
            rows: vec![f(0, "Region")],
            column: Some(f(1, "Month")),
            values: vec![v(2, "Amount", Aggregation::Average)],
        };
        let json = serde_json::to_string(&def).unwrap();
        assert!(json.contains("\"average\""));
        let back: PivotDefinition = serde_json::from_str(&json).unwrap();
        assert_eq!(back, def);
    }
}
