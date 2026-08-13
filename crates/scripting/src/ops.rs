//! Lua operation types and command sink.
//!
//! # Architecture
//!
//! Lua never touches workbook state directly. Instead:
//!
//! 1. Lua calls `sheet:set_value()` etc.
//! 2. These push `LuaOp` entries into a journal
//! 3. Reads consult `pending` shadow map first, then fall back to workbook
//! 4. After Lua returns, VisiGrid applies ops as a batch
//!
//! This ensures:
//! - No borrow checker issues during Lua execution
//! - Single recalc after all mutations
//! - Single undo entry per script
//! - Deterministic, replayable execution

use std::collections::HashMap;

// ============================================================================
// Cell Key (cheap lookup key)
// ============================================================================

/// Packed cell coordinate for HashMap keys.
/// Using u64 = (row << 32) | col for cheap hashing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellKey(u64);

impl CellKey {
    #[inline]
    pub fn new(row: u32, col: u32) -> Self {
        Self(((row as u64) << 32) | (col as u64))
    }

    #[inline]
    pub fn row(&self) -> u32 {
        (self.0 >> 32) as u32
    }

    #[inline]
    pub fn col(&self) -> u32 {
        self.0 as u32
    }

    #[inline]
    pub fn unpack(&self) -> (u32, u32) {
        (self.row(), self.col())
    }
}

impl From<(u32, u32)> for CellKey {
    fn from((row, col): (u32, u32)) -> Self {
        Self::new(row, col)
    }
}

impl From<(usize, usize)> for CellKey {
    fn from((row, col): (usize, usize)) -> Self {
        Self::new(row as u32, col as u32)
    }
}

// ============================================================================
// Lua Cell Value (typed, not display string)
// ============================================================================

/// Typed cell value for Lua interaction.
///
/// This is NOT a display string. Numbers are f64, bools are bool, etc.
/// Nil means "clear the cell" when writing, or "empty cell" when reading.
#[derive(Debug, Clone, PartialEq)]
pub enum LuaCellValue {
    /// Empty cell / clear cell
    Nil,
    /// Numeric value
    Number(f64),
    /// Text value
    String(String),
    /// Boolean value
    Bool(bool),
    /// Error value (from formula evaluation)
    Error(String),
}

impl LuaCellValue {
    /// Check if this is a nil/empty value
    pub fn is_nil(&self) -> bool {
        matches!(self, LuaCellValue::Nil)
    }

    /// Convert to a display string (for debugging/REPL output)
    pub fn display(&self) -> String {
        match self {
            LuaCellValue::Nil => "nil".to_string(),
            LuaCellValue::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{:.0}", n)
                } else {
                    format!("{}", n)
                }
            }
            LuaCellValue::String(s) => s.clone(),
            LuaCellValue::Bool(b) => b.to_string(),
            LuaCellValue::Error(e) => format!("#ERROR: {}", e),
        }
    }
}

// ============================================================================
// Lua Operations (journal entries)
// ============================================================================

/// A single operation that Lua wants to perform on the workbook.
///
/// These are collected during script execution and applied as a batch
/// after the script completes.
#[derive(Debug, Clone)]
pub enum LuaOp {
    /// Set a cell's value (clears any formula)
    SetValue {
        row: u32,
        col: u32,
        value: LuaCellValue,
    },
    /// Set a cell's formula (value will be computed on recalc)
    SetFormula {
        row: u32,
        col: u32,
        formula: String,
    },
    /// Set cell style on a range (format-only, no recalc needed)
    SetCellStyle {
        r1: u32, c1: u32,  // top-left (0-indexed)
        r2: u32, c2: u32,  // bottom-right (0-indexed)
        style: u8,          // CellStyle::to_int() value
    },
}

impl LuaOp {
    /// Get the cell key this operation affects (top-left for range ops)
    pub fn cell_key(&self) -> CellKey {
        match self {
            LuaOp::SetValue { row, col, .. } => CellKey::new(*row, *col),
            LuaOp::SetFormula { row, col, .. } => CellKey::new(*row, *col),
            LuaOp::SetCellStyle { r1, c1, .. } => CellKey::new(*r1, *c1),
        }
    }
}

// ============================================================================
// Pending Cell (shadow map entry)
// ============================================================================

/// What's pending for a cell (for read-after-write correctness).
///
/// Tracks whether a value or formula was set, so reads can see their own writes.
#[derive(Debug, Clone)]
pub enum PendingCell {
    /// A literal value was set
    Value(LuaCellValue),
    /// A formula was set (value unknown until recalc)
    Formula(String),
}

impl PendingCell {
    /// Get the value for in-script reads.
    ///
    /// For Value: returns the value
    /// For Formula: returns Nil (formula result unknown until recalc)
    pub fn read_value(&self) -> LuaCellValue {
        match self {
            PendingCell::Value(v) => v.clone(),
            // Formula result is unknown until recalc; return Nil in-script
            PendingCell::Formula(_) => LuaCellValue::Nil,
        }
    }

    /// Get the formula string if this is a formula
    pub fn formula(&self) -> Option<&str> {
        match self {
            PendingCell::Formula(f) => Some(f.as_str()),
            PendingCell::Value(_) => None,
        }
    }
}

// ============================================================================
// Sheet Reader (read-only workbook access)
// ============================================================================

/// Read-only interface to the workbook for Lua reads.
///
/// This trait allows LuaOpSink to read from the workbook without
/// holding a mutable reference. Implementations can use per-call
/// locking or snapshots.
pub trait SheetReader {
    /// Get the typed value at (row, col)
    fn get_value(&self, row: usize, col: usize) -> LuaCellValue;

    /// Get the formula at (row, col), if any
    fn get_formula(&self, row: usize, col: usize) -> Option<String>;

    /// Number of rows with data
    fn rows(&self) -> usize;

    /// Number of columns with data
    fn cols(&self) -> usize;
}

// ============================================================================
// Lua Op Sink (the bridge)
// ============================================================================

/// The command sink that Lua's `sheet` userdata talks to.
///
/// This is the ONLY path between Lua and the workbook.
/// It owns:
/// - The ops journal (ordered list of mutations)
/// - The pending shadow map (for read-after-write)
/// - A read-only reference to the sheet
pub struct LuaOpSink<'a, R: SheetReader> {
    /// Ordered list of operations (journal)
    pub ops: Vec<LuaOp>,
    /// Shadow map for read-after-write
    pending: HashMap<CellKey, PendingCell>,
    /// Read-only sheet access
    reader: &'a R,
}

impl<'a, R: SheetReader> LuaOpSink<'a, R> {
    /// Create a new sink with a sheet reader
    pub fn new(reader: &'a R) -> Self {
        Self {
            ops: Vec::new(),
            pending: HashMap::new(),
            reader,
        }
    }

    /// Get the number of unique cells modified
    pub fn mutations(&self) -> usize {
        self.pending.len()
    }

    /// Get the ops (consumes self)
    pub fn into_ops(self) -> Vec<LuaOp> {
        self.ops
    }

    // ========================================================================
    // Read operations (check pending first, then fall back to reader)
    // ========================================================================

    /// Get value at (row, col) - checks pending first
    pub fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
        let key = CellKey::from((row, col));

        // Check pending first (read-after-write)
        if let Some(pending) = self.pending.get(&key) {
            return pending.read_value();
        }

        // Fall back to workbook
        self.reader.get_value(row, col)
    }

    /// Get formula at (row, col) - checks pending first
    pub fn get_formula(&self, row: usize, col: usize) -> Option<String> {
        let key = CellKey::from((row, col));

        // Check pending first
        if let Some(pending) = self.pending.get(&key) {
            return pending.formula().map(|s| s.to_string());
        }

        // Fall back to workbook
        self.reader.get_formula(row, col)
    }

    /// Get row count
    pub fn rows(&self) -> usize {
        self.reader.rows()
    }

    /// Get column count
    pub fn cols(&self) -> usize {
        self.reader.cols()
    }

    // ========================================================================
    // Write operations (update pending + append to journal)
    // ========================================================================

    /// Set value at (row, col)
    pub fn set_value(&mut self, row: usize, col: usize, value: LuaCellValue) {
        let key = CellKey::from((row, col));

        // Update shadow map
        self.pending.insert(key, PendingCell::Value(value.clone()));

        // Append to journal
        self.ops.push(LuaOp::SetValue {
            row: row as u32,
            col: col as u32,
            value,
        });
    }

    /// Set formula at (row, col)
    pub fn set_formula(&mut self, row: usize, col: usize, formula: String) {
        let key = CellKey::from((row, col));

        // Update shadow map
        self.pending.insert(key, PendingCell::Formula(formula.clone()));

        // Append to journal
        self.ops.push(LuaOp::SetFormula {
            row: row as u32,
            col: col as u32,
            formula,
        });
    }
}

// ============================================================================
// A1 notation parsing
// ============================================================================

/// Parse A1 notation to (row, col) - 1-indexed for Lua
pub fn parse_a1(a1: &str) -> Option<(usize, usize)> {
    let a1 = a1.trim().to_uppercase();
    if a1.is_empty() {
        return None;
    }

    // Find where letters end and digits begin
    let col_end = a1.chars().take_while(|c| c.is_ascii_alphabetic()).count();
    if col_end == 0 || col_end >= a1.len() {
        return None;
    }

    let col_str = &a1[..col_end];
    let row_str = &a1[col_end..];

    // Parse column (A=1, B=2, ..., Z=26, AA=27, etc.)
    let mut col: usize = 0;
    for c in col_str.chars() {
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }

    // Parse row (1-indexed in A1, we keep 1-indexed for Lua)
    let row: usize = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    Some((row, col))
}

/// Format (row, col) as A1 notation - expects 1-indexed
pub fn format_a1(row: usize, col: usize) -> String {
    if row == 0 || col == 0 {
        return "".to_string();
    }

    let mut col_str = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        col_str.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }

    format!("{}{}", col_str, row)
}

/// Parse a range in A1 notation (e.g., "A1:C5") to ((start_row, start_col), (end_row, end_col)).
/// Returns 1-indexed coordinates for Lua convention.
/// Also accepts single cell "A1" (returns same cell for start and end).
pub fn parse_range(range: &str) -> Option<((usize, usize), (usize, usize))> {
    let range = range.trim();

    if let Some(colon_pos) = range.find(':') {
        // Range like "A1:C5"
        let start = &range[..colon_pos];
        let end = &range[colon_pos + 1..];

        let (start_row, start_col) = parse_a1(start)?;
        let (end_row, end_col) = parse_a1(end)?;

        // Normalize to ensure start <= end
        Some((
            (start_row.min(end_row), start_col.min(end_col)),
            (start_row.max(end_row), start_col.max(end_col)),
        ))
    } else {
        // Single cell like "A1"
        let (row, col) = parse_a1(range)?;
        Some(((row, col), (row, col)))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========================================================================
    // CellKey tests
    // ========================================================================

    #[test]
    fn test_cell_key_pack_unpack() {
        let key = CellKey::new(100, 200);
        assert_eq!(key.row(), 100);
        assert_eq!(key.col(), 200);
        assert_eq!(key.unpack(), (100, 200));
    }

    #[test]
    fn test_cell_key_from_tuple() {
        let key: CellKey = (10u32, 20u32).into();
        assert_eq!(key.unpack(), (10, 20));

        let key: CellKey = (10usize, 20usize).into();
        assert_eq!(key.unpack(), (10, 20));
    }

    // ========================================================================
    // A1 parsing tests
    // ========================================================================

    #[test]
    fn test_parse_a1_simple() {
        assert_eq!(parse_a1("A1"), Some((1, 1)));
        assert_eq!(parse_a1("B2"), Some((2, 2)));
        assert_eq!(parse_a1("Z26"), Some((26, 26)));
    }

    #[test]
    fn test_parse_a1_double_letter() {
        assert_eq!(parse_a1("AA1"), Some((1, 27)));
        assert_eq!(parse_a1("AB1"), Some((1, 28)));
        assert_eq!(parse_a1("AZ1"), Some((1, 52)));
        assert_eq!(parse_a1("BA1"), Some((1, 53)));
    }

    #[test]
    fn test_parse_a1_case_insensitive() {
        assert_eq!(parse_a1("a1"), Some((1, 1)));
        assert_eq!(parse_a1("aA1"), Some((1, 27)));
    }

    #[test]
    fn test_parse_a1_invalid() {
        assert_eq!(parse_a1(""), None);
        assert_eq!(parse_a1("A"), None);
        assert_eq!(parse_a1("1"), None);
        assert_eq!(parse_a1("A0"), None);
    }

    #[test]
    fn test_format_a1() {
        assert_eq!(format_a1(1, 1), "A1");
        assert_eq!(format_a1(2, 2), "B2");
        assert_eq!(format_a1(26, 26), "Z26");
        assert_eq!(format_a1(1, 27), "AA1");
        assert_eq!(format_a1(1, 28), "AB1");
    }

    #[test]
    fn test_a1_roundtrip() {
        for row in 1..=100 {
            for col in 1..=100 {
                let a1 = format_a1(row, col);
                assert_eq!(parse_a1(&a1), Some((row, col)), "Failed for {}", a1);
            }
        }
    }

    // ========================================================================
    // Range parsing tests
    // ========================================================================

    #[test]
    fn test_parse_range_simple() {
        assert_eq!(parse_range("A1:C3"), Some(((1, 1), (3, 3))));
        assert_eq!(parse_range("B2:D5"), Some(((2, 2), (5, 4))));
    }

    #[test]
    fn test_parse_range_single_cell() {
        assert_eq!(parse_range("A1"), Some(((1, 1), (1, 1))));
        assert_eq!(parse_range("B2"), Some(((2, 2), (2, 2))));
    }

    #[test]
    fn test_parse_range_normalizes() {
        // Reversed ranges should normalize to start <= end
        assert_eq!(parse_range("C3:A1"), Some(((1, 1), (3, 3))));
        assert_eq!(parse_range("D5:B2"), Some(((2, 2), (5, 4))));
    }

    #[test]
    fn test_parse_range_invalid() {
        assert_eq!(parse_range(""), None);
        assert_eq!(parse_range(":"), None);
        assert_eq!(parse_range("A1:"), None);
        assert_eq!(parse_range(":B2"), None);
    }

    // ========================================================================
    // LuaCellValue tests
    // ========================================================================

    #[test]
    fn test_lua_cell_value_display() {
        assert_eq!(LuaCellValue::Nil.display(), "nil");
        assert_eq!(LuaCellValue::Number(42.0).display(), "42");
        assert_eq!(LuaCellValue::Number(3.14).display(), "3.14");
        assert_eq!(LuaCellValue::String("hello".to_string()).display(), "hello");
        assert_eq!(LuaCellValue::Bool(true).display(), "true");
        assert_eq!(LuaCellValue::Error("DIV/0".to_string()).display(), "#ERROR: DIV/0");
    }

    // ========================================================================
    // LuaOpSink tests (with mock reader)
    // ========================================================================

    struct MockReader {
        data: HashMap<CellKey, LuaCellValue>,
        formulas: HashMap<CellKey, String>,
    }

    impl MockReader {
        fn new() -> Self {
            Self {
                data: HashMap::new(),
                formulas: HashMap::new(),
            }
        }

        fn set(&mut self, row: usize, col: usize, value: LuaCellValue) {
            self.data.insert(CellKey::from((row, col)), value);
        }

        fn set_formula_direct(&mut self, row: usize, col: usize, formula: &str) {
            self.formulas.insert(CellKey::from((row, col)), formula.to_string());
        }
    }

    impl SheetReader for MockReader {
        fn get_value(&self, row: usize, col: usize) -> LuaCellValue {
            self.data
                .get(&CellKey::from((row, col)))
                .cloned()
                .unwrap_or(LuaCellValue::Nil)
        }

        fn get_formula(&self, row: usize, col: usize) -> Option<String> {
            self.formulas.get(&CellKey::from((row, col))).cloned()
        }

        fn rows(&self) -> usize {
            100
        }

        fn cols(&self) -> usize {
            26
        }
    }

    #[test]
    fn test_sink_read_from_reader() {
        let mut reader = MockReader::new();
        reader.set(1, 1, LuaCellValue::Number(42.0));
        reader.set_formula_direct(2, 2, "=A1*2");

        let sink = LuaOpSink::new(&reader);

        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(42.0));
        assert_eq!(sink.get_value(3, 3), LuaCellValue::Nil);
        assert_eq!(sink.get_formula(2, 2), Some("=A1*2".to_string()));
        assert_eq!(sink.get_formula(1, 1), None);
    }

    #[test]
    fn test_sink_read_after_write() {
        let reader = MockReader::new();
        let mut sink = LuaOpSink::new(&reader);

        // Initially nil
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Nil);

        // Write
        sink.set_value(1, 1, LuaCellValue::Number(100.0));

        // Read should see the pending write
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(100.0));

        // Ops should have one entry
        assert_eq!(sink.ops.len(), 1);
    }

    #[test]
    fn test_sink_write_ordering() {
        let reader = MockReader::new();
        let mut sink = LuaOpSink::new(&reader);

        sink.set_value(1, 1, LuaCellValue::Number(1.0));
        sink.set_value(1, 1, LuaCellValue::Number(2.0));

        // Shadow shows final value
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(2.0));

        // But ops has both (for journal/replay purposes)
        assert_eq!(sink.ops.len(), 2);

        // Mutations count is 1 (one unique cell)
        assert_eq!(sink.mutations(), 1);
    }

    #[test]
    fn test_sink_formula_overrides_value() {
        let reader = MockReader::new();
        let mut sink = LuaOpSink::new(&reader);

        sink.set_value(1, 1, LuaCellValue::Number(100.0));
        sink.set_formula(1, 1, "=A2*2".to_string());

        // Formula is now pending
        assert_eq!(sink.get_formula(1, 1), Some("=A2*2".to_string()));

        // Value reads as Nil (formula result unknown until recalc)
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Nil);
    }

    #[test]
    fn test_sink_value_overrides_formula() {
        let reader = MockReader::new();
        let mut sink = LuaOpSink::new(&reader);

        sink.set_formula(1, 1, "=A2*2".to_string());
        sink.set_value(1, 1, LuaCellValue::Number(100.0));

        // Value is now pending, formula is gone
        assert_eq!(sink.get_formula(1, 1), None);
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(100.0));
    }

    #[test]
    fn test_sink_nil_clears() {
        let mut reader = MockReader::new();
        reader.set(1, 1, LuaCellValue::Number(42.0));

        let mut sink = LuaOpSink::new(&reader);

        // Reader has a value
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Number(42.0));

        // Set to nil
        sink.set_value(1, 1, LuaCellValue::Nil);

        // Now reads as nil (pending overrides reader)
        assert_eq!(sink.get_value(1, 1), LuaCellValue::Nil);
    }
}
