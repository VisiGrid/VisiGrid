// Formula evaluator - evaluates bound expressions (after sheet name resolution)

use crate::sheet::{SheetId, SheetRef};
use super::parser::{BoundExpr, Expr, Op};

/// Result of resolving a named range
#[derive(Debug, Clone)]
pub enum NamedRangeResolution {
    Cell { row: usize, col: usize },
    Range { start_row: usize, start_col: usize, end_row: usize, end_col: usize },
}

pub trait CellLookup {
    fn get_value(&self, row: usize, col: usize) -> f64;
    fn get_text(&self, row: usize, col: usize) -> String;

    /// Get a cell's value from another sheet by SheetId.
    /// Returns Value::Error("#REF!") if sheet doesn't exist.
    /// Default implementation returns #REF! (cross-sheet not supported).
    fn get_value_sheet(&self, _sheet_id: SheetId, _row: usize, _col: usize) -> Value {
        Value::Error("#REF!".to_string())
    }

    /// Get a cell's text from another sheet by SheetId.
    /// Returns "#REF!" if sheet doesn't exist.
    /// Default implementation returns #REF! (cross-sheet not supported).
    fn get_text_sheet(&self, _sheet_id: SheetId, _row: usize, _col: usize) -> String {
        "#REF!".to_string()
    }

    /// Resolve a named range to its target. Returns None if name not defined.
    /// Default implementation returns None (named ranges not supported).
    fn resolve_named_range(&self, _name: &str) -> Option<NamedRangeResolution> {
        None
    }

    /// Return diagnostic context info (for debugging recalc issues).
    /// Default implementation returns empty string.
    fn debug_context(&self) -> String {
        String::new()
    }

    /// Get the current cell being evaluated (for ROW()/COLUMN() without args).
    /// Default implementation returns None (not in cell context).
    fn current_cell(&self) -> Option<(usize, usize)> {
        None
    }

    /// Query merge origin for a cell on the current sheet. Returns None if not merged.
    fn get_merge_start(&self, _row: usize, _col: usize) -> Option<(usize, usize)> {
        None
    }

    /// Query merge origin for a cell on another sheet. Returns None if not merged.
    fn get_merge_start_sheet(&self, _sheet_id: SheetId, _row: usize, _col: usize) -> Option<(usize, usize)> {
        None
    }

    /// Get a cell's typed value from the current sheet.
    /// Returns Value::Empty for empty/missing cells.
    /// Default: falls back to get_text() + parse (backward compat).
    fn get_cell_value(&self, row: usize, col: usize) -> Value {
        let text = self.get_text(row, col);
        if text.is_empty() {
            Value::Empty
        } else if text.starts_with('#') {
            Value::Error(text)
        } else if let Ok(n) = text.parse::<f64>() {
            Value::Number(n)
        } else if text.eq_ignore_ascii_case("TRUE") {
            Value::Boolean(true)
        } else if text.eq_ignore_ascii_case("FALSE") {
            Value::Boolean(false)
        } else {
            Value::Text(text)
        }
    }

    /// Try to evaluate a custom (user-defined) function.
    /// Args are already evaluated by the engine — scalars resolved, ranges snapshotted as typed Values.
    /// Returns None if the function is not a custom function.
    /// Returns Some(EvalResult) if it is (including errors).
    /// Default: no custom functions available.
    fn try_custom_function(&self, _name: &str, _args: &[EvalArg]) -> Option<EvalResult> {
        None
    }
}

/// A lookup that wraps another CellLookup and adds named range resolution
pub struct LookupWithNamedRanges<'a, L: CellLookup, F: Fn(&str) -> Option<NamedRangeResolution>> {
    inner: &'a L,
    resolver: F,
}

impl<'a, L: CellLookup, F: Fn(&str) -> Option<NamedRangeResolution>> LookupWithNamedRanges<'a, L, F> {
    pub fn new(inner: &'a L, resolver: F) -> Self {
        Self { inner, resolver }
    }
}

impl<'a, L: CellLookup, F: Fn(&str) -> Option<NamedRangeResolution>> CellLookup for LookupWithNamedRanges<'a, L, F> {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        self.inner.get_value(row, col)
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        self.inner.get_text(row, col)
    }

    fn get_value_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Value {
        self.inner.get_value_sheet(sheet_id, row, col)
    }

    fn get_text_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> String {
        self.inner.get_text_sheet(sheet_id, row, col)
    }

    fn resolve_named_range(&self, name: &str) -> Option<NamedRangeResolution> {
        (self.resolver)(name)
    }

    fn current_cell(&self) -> Option<(usize, usize)> {
        self.inner.current_cell()
    }

    fn get_merge_start(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.inner.get_merge_start(row, col)
    }

    fn get_merge_start_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Option<(usize, usize)> {
        self.inner.get_merge_start_sheet(sheet_id, row, col)
    }

    fn get_cell_value(&self, row: usize, col: usize) -> Value {
        self.inner.get_cell_value(row, col)
    }

    fn try_custom_function(&self, name: &str, args: &[EvalArg]) -> Option<EvalResult> {
        self.inner.try_custom_function(name, args)
    }
}

/// A lookup wrapper that provides current cell context for ROW()/COLUMN()
pub struct LookupWithContext<'a, L: CellLookup> {
    inner: &'a L,
    current_row: usize,
    current_col: usize,
}

impl<'a, L: CellLookup> LookupWithContext<'a, L> {
    pub fn new(inner: &'a L, current_row: usize, current_col: usize) -> Self {
        Self { inner, current_row, current_col }
    }
}

impl<'a, L: CellLookup> CellLookup for LookupWithContext<'a, L> {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        self.inner.get_value(row, col)
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        self.inner.get_text(row, col)
    }

    fn get_value_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Value {
        self.inner.get_value_sheet(sheet_id, row, col)
    }

    fn get_text_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> String {
        self.inner.get_text_sheet(sheet_id, row, col)
    }

    fn resolve_named_range(&self, name: &str) -> Option<NamedRangeResolution> {
        self.inner.resolve_named_range(name)
    }

    fn current_cell(&self) -> Option<(usize, usize)> {
        Some((self.current_row, self.current_col))
    }

    fn get_merge_start(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.inner.get_merge_start(row, col)
    }

    fn get_merge_start_sheet(&self, sheet_id: SheetId, row: usize, col: usize) -> Option<(usize, usize)> {
        self.inner.get_merge_start_sheet(sheet_id, row, col)
    }

    fn get_cell_value(&self, row: usize, col: usize) -> Value {
        self.inner.get_cell_value(row, col)
    }

    fn try_custom_function(&self, name: &str, args: &[EvalArg]) -> Option<EvalResult> {
        self.inner.try_custom_function(name, args)
    }
}

// =============================================================================
// Value: The scalar primitive for all cell values
// =============================================================================

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(String),
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}

impl Value {
    pub fn to_number(&self) -> Result<f64, String> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Text(s) if s.is_empty() => Ok(0.0),
            Value::Text(s) => {
                // Try numeric parse first
                if let Ok(n) = s.parse::<f64>() {
                    return Ok(n);
                }
                // Try date string parse (ISO: 2023-11-07, US: 11/07/2023)
                if let Some(serial) = super::eval_helpers::try_parse_date_string(s) {
                    return Ok(serial);
                }
                Err(format!("#VALUE! Cannot convert '{}' to number", s))
            }
            Value::Empty => Ok(0.0),
            Value::Error(e) => Err(e.clone()),
        }
    }

    pub fn to_text(&self) -> String {
        match self {
            Value::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            Value::Text(s) => s.clone(),
            Value::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            Value::Empty => String::new(),
            Value::Error(e) => e.clone(),
        }
    }

    pub fn to_bool(&self) -> Result<bool, String> {
        match self {
            Value::Boolean(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Text(s) => {
                let upper = s.to_uppercase();
                if upper == "TRUE" { Ok(true) }
                else if upper == "FALSE" { Ok(false) }
                else { Err(format!("#VALUE! Cannot convert '{}' to boolean", s)) }
            }
            Value::Empty => Ok(false),
            Value::Error(e) => Err(e.clone()),
        }
    }

    pub fn is_error(&self) -> bool {
        matches!(self, Value::Error(_))
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Value::Empty)
    }
}

// =============================================================================
// EvalArg: Argument passed to custom function hooks (already evaluated)
// =============================================================================

/// Argument passed to custom function hooks, already evaluated by the engine.
/// This keeps the engine/app boundary clean: no AST leaks across.
#[derive(Debug, Clone)]
pub enum EvalArg {
    Scalar(Value),
    Range {
        /// Typed cell values, row-major flattened. Uses engine Value (not text).
        values: Vec<Value>,
        num_cells: usize,
    },
}

// =============================================================================
// Array2D: 2D grid of Values (dense storage, row-major)
// =============================================================================

#[derive(Debug, Clone, PartialEq)]
pub struct Array2D {
    data: Vec<Value>,
    rows: usize,
    cols: usize,
}

impl Array2D {
    pub fn new(rows: usize, cols: usize) -> Self {
        Self {
            data: vec![Value::Empty; rows * cols],
            rows,
            cols,
        }
    }

    pub fn from_vec(data: Vec<Vec<Value>>) -> Self {
        if data.is_empty() {
            return Self::new(0, 0);
        }
        let rows = data.len();
        let cols = data[0].len();
        let mut flat = Vec::with_capacity(rows * cols);
        for row in data {
            for val in row {
                flat.push(val);
            }
        }
        Self { data: flat, rows, cols }
    }

    pub fn rows(&self) -> usize {
        self.rows
    }

    pub fn cols(&self) -> usize {
        self.cols
    }

    pub fn get(&self, row: usize, col: usize) -> Option<&Value> {
        if row < self.rows && col < self.cols {
            Some(&self.data[row * self.cols + col])
        } else {
            None
        }
    }

    pub fn set(&mut self, row: usize, col: usize, value: Value) {
        if row < self.rows && col < self.cols {
            self.data[row * self.cols + col] = value;
        }
    }

    /// Get the top-left value (for scalar coercion)
    pub fn top_left(&self) -> Value {
        self.get(0, 0).cloned().unwrap_or(Value::Empty)
    }

    /// Check if this is a 1x1 array (effectively scalar)
    pub fn is_scalar(&self) -> bool {
        self.rows == 1 && self.cols == 1
    }

    /// Convert 1x1 array to scalar Value
    pub fn to_scalar(&self) -> Option<Value> {
        if self.is_scalar() {
            Some(self.top_left())
        } else {
            None
        }
    }
}

// =============================================================================
// EvalResult: The result of formula evaluation (scalar or array)
// =============================================================================

#[derive(Debug, Clone, PartialEq)]
pub enum EvalResult {
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(String),
    Array(Array2D),
}

impl EvalResult {
    pub fn to_display(&self) -> String {
        match self {
            EvalResult::Number(n) => {
                if n.is_nan() {
                    "#NAN".to_string()
                } else if n.is_infinite() {
                    "#INF".to_string()
                } else if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{:.4}", n).trim_end_matches('0').trim_end_matches('.').to_string()
                }
            }
            EvalResult::Text(s) => s.clone(),
            EvalResult::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            EvalResult::Error(e) => {
                if e.starts_with('#') { e.clone() } else { format!("#ERR: {}", e) }
            }
            EvalResult::Array(arr) => {
                // Display top-left value for single-cell display
                arr.top_left().to_text()
            }
        }
    }

    /// Convert result to a number (for arithmetic operations)
    /// Arrays coerce to their top-left value
    /// Also parses ISO date strings (2023-11-07) to Excel serial numbers
    pub fn to_number(&self) -> Result<f64, String> {
        match self {
            EvalResult::Number(n) => Ok(*n),
            EvalResult::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
            EvalResult::Text(s) => {
                // Try numeric parse first
                if let Ok(n) = s.parse::<f64>() {
                    return Ok(n);
                }
                // Try date string parse (ISO: 2023-11-07, US: 11/07/2023)
                if let Some(serial) = super::eval_helpers::try_parse_date_string(s) {
                    return Ok(serial);
                }
                Err(format!("Cannot convert '{}' to number", s))
            }
            EvalResult::Error(e) => Err(e.clone()),
            EvalResult::Array(arr) => arr.top_left().to_number(),
        }
    }

    /// Convert result to a string (for text operations)
    /// Arrays coerce to their top-left value
    pub fn to_text(&self) -> String {
        match self {
            EvalResult::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            EvalResult::Text(s) => s.clone(),
            EvalResult::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            EvalResult::Error(e) => {
                if e.starts_with('#') { e.clone() } else { format!("#ERR: {}", e) }
            }
            EvalResult::Array(arr) => arr.top_left().to_text(),
        }
    }

    /// Convert result to a boolean (for logical operations)
    /// Arrays coerce to their top-left value
    pub fn to_bool(&self) -> Result<bool, String> {
        match self {
            EvalResult::Boolean(b) => Ok(*b),
            EvalResult::Number(n) => Ok(*n != 0.0),
            EvalResult::Text(s) => {
                let upper = s.to_uppercase();
                if upper == "TRUE" { Ok(true) }
                else if upper == "FALSE" { Ok(false) }
                else { Err(format!("Cannot convert '{}' to boolean", s)) }
            }
            EvalResult::Error(e) => Err(e.clone()),
            EvalResult::Array(arr) => arr.top_left().to_bool(),
        }
    }

    /// Check if this is an error
    pub fn is_error(&self) -> bool {
        matches!(self, EvalResult::Error(_))
    }

    /// Check if this is an array result
    pub fn is_array(&self) -> bool {
        matches!(self, EvalResult::Array(_))
    }

    /// Get array dimensions if this is an array, else (1, 1)
    pub fn dimensions(&self) -> (usize, usize) {
        match self {
            EvalResult::Array(arr) => (arr.rows(), arr.cols()),
            _ => (1, 1),
        }
    }

    /// Convert EvalResult to Value (for storage)
    pub fn to_value(&self) -> Value {
        match self {
            EvalResult::Number(n) => Value::Number(*n),
            EvalResult::Text(s) => Value::Text(s.clone()),
            EvalResult::Boolean(b) => Value::Boolean(*b),
            EvalResult::Error(e) => Value::Error(e.clone()),
            EvalResult::Array(arr) => arr.top_left(),
        }
    }

    /// Convert Value to EvalResult
    pub fn from_value(v: &Value) -> EvalResult {
        match v {
            Value::Empty => EvalResult::Number(0.0),
            Value::Number(n) => EvalResult::Number(*n),
            Value::Text(s) => EvalResult::Text(s.clone()),
            Value::Boolean(b) => EvalResult::Boolean(*b),
            Value::Error(e) => EvalResult::Error(e.clone()),
        }
    }
}

pub fn evaluate<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> EvalResult {
    match expr {
        Expr::Number(n) => EvalResult::Number(*n),
        Expr::Text(s) => EvalResult::Text(s.clone()),
        Expr::Boolean(b) => EvalResult::Boolean(*b),
        Expr::CellRef { sheet, col, row, .. } => {
            // Get cell value, potentially from another sheet.
            // Redirect hidden merge cells to origin so =B1 returns the
            // merge origin's value when B1 is hidden inside a merge.
            let text = match sheet {
                SheetRef::Current => {
                    let (r, c) = lookup.get_merge_start(*row, *col).unwrap_or((*row, *col));
                    lookup.get_text(r, c)
                }
                SheetRef::Id(sheet_id) => {
                    let (r, c) = lookup.get_merge_start_sheet(*sheet_id, *row, *col).unwrap_or((*row, *col));
                    lookup.get_text_sheet(*sheet_id, r, c)
                }
                SheetRef::RefError { .. } => return EvalResult::Error("#REF!".to_string()),
            };
            if text.is_empty() {
                EvalResult::Number(0.0)
            } else if text.starts_with('#') {
                // Propagate errors (e.g., #CIRC!, #REF!, #VALUE!)
                EvalResult::Error(text)
            } else if let Ok(n) = text.parse::<f64>() {
                EvalResult::Number(n)
            } else if text.to_uppercase() == "TRUE" {
                EvalResult::Boolean(true)
            } else if text.to_uppercase() == "FALSE" {
                EvalResult::Boolean(false)
            } else {
                EvalResult::Text(text)
            }
        }
        Expr::Range { .. } => {
            // Ranges can't be evaluated directly, only within functions
            EvalResult::Error("Array arithmetic not yet supported. Use SUMPRODUCT(A1:A10, B1:B10) instead of SUM(A1:A10*B1:B10).".to_string())
        }
        Expr::NamedRange(name) => {
            // Resolve the named range and evaluate
            match lookup.resolve_named_range(name) {
                None => EvalResult::Error(format!("#NAME? '{}'", name)),
                Some(NamedRangeResolution::Cell { row, col }) => {
                    // Evaluate like a cell reference
                    let text = lookup.get_text(row, col);
                    if text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if text.starts_with('#') {
                        EvalResult::Error(text)
                    } else if let Ok(n) = text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else if text.to_uppercase() == "TRUE" {
                        EvalResult::Boolean(true)
                    } else if text.to_uppercase() == "FALSE" {
                        EvalResult::Boolean(false)
                    } else {
                        EvalResult::Text(text)
                    }
                }
                Some(NamedRangeResolution::Range { .. }) => {
                    // Ranges can't be evaluated directly, only within functions
                    EvalResult::Error("Array arithmetic not yet supported. Use SUMPRODUCT(A1:A10, B1:B10) instead of SUM(A1:A10*B1:B10).".to_string())
                }
            }
        }
        Expr::Function { name, args } => evaluate_function(name, args, lookup),
        Expr::BinaryOp { op, left, right } => {
            let left_result = evaluate(left, lookup);
            let right_result = evaluate(right, lookup);

            // Check for errors first
            if let EvalResult::Error(e) = &left_result {
                return EvalResult::Error(e.clone());
            }
            if let EvalResult::Error(e) = &right_result {
                return EvalResult::Error(e.clone());
            }

            match op {
                // Arithmetic operators - require numbers
                Op::Add | Op::Sub | Op::Mul | Op::Div | Op::Pow => {
                    let left_val = match left_result.to_number() {
                        Ok(n) => n,
                        Err(e) => return EvalResult::Error(e),
                    };
                    let right_val = match right_result.to_number() {
                        Ok(n) => n,
                        Err(e) => return EvalResult::Error(e),
                    };

                    let result = match op {
                        Op::Add => left_val + right_val,
                        Op::Sub => left_val - right_val,
                        Op::Mul => left_val * right_val,
                        Op::Div => {
                            if right_val == 0.0 {
                                return EvalResult::Error("#DIV/0!".to_string());
                            }
                            left_val / right_val
                        }
                        Op::Pow => left_val.powf(right_val),
                        _ => unreachable!(),
                    };
                    EvalResult::Number(result)
                }

                // Comparison operators
                Op::Lt | Op::Gt | Op::Eq | Op::LtEq | Op::GtEq | Op::NotEq => {
                    // Compare based on types - numbers compare numerically, text alphabetically
                    let result = match (&left_result, &right_result) {
                        (EvalResult::Number(a), EvalResult::Number(b)) => {
                            match op {
                                Op::Lt => a < b,
                                Op::Gt => a > b,
                                Op::Eq => (a - b).abs() < f64::EPSILON,
                                Op::LtEq => a <= b,
                                Op::GtEq => a >= b,
                                Op::NotEq => (a - b).abs() >= f64::EPSILON,
                                _ => unreachable!(),
                            }
                        }
                        (EvalResult::Text(a), EvalResult::Text(b)) => {
                            let a_lower = a.to_lowercase();
                            let b_lower = b.to_lowercase();
                            match op {
                                Op::Lt => a_lower < b_lower,
                                Op::Gt => a_lower > b_lower,
                                Op::Eq => a_lower == b_lower,
                                Op::LtEq => a_lower <= b_lower,
                                Op::GtEq => a_lower >= b_lower,
                                Op::NotEq => a_lower != b_lower,
                                _ => unreachable!(),
                            }
                        }
                        (EvalResult::Boolean(a), EvalResult::Boolean(b)) => {
                            match op {
                                Op::Eq => a == b,
                                Op::NotEq => a != b,
                                _ => return EvalResult::Error("Cannot compare booleans with < > <= >=".to_string()),
                            }
                        }
                        // Mixed type comparisons - convert to common type
                        _ => {
                            // Try numeric comparison first
                            if let (Ok(a), Ok(b)) = (left_result.to_number(), right_result.to_number()) {
                                match op {
                                    Op::Lt => a < b,
                                    Op::Gt => a > b,
                                    Op::Eq => (a - b).abs() < f64::EPSILON,
                                    Op::LtEq => a <= b,
                                    Op::GtEq => a >= b,
                                    Op::NotEq => (a - b).abs() >= f64::EPSILON,
                                    _ => unreachable!(),
                                }
                            } else {
                                // Fall back to text comparison
                                let a = left_result.to_text().to_lowercase();
                                let b = right_result.to_text().to_lowercase();
                                match op {
                                    Op::Lt => a < b,
                                    Op::Gt => a > b,
                                    Op::Eq => a == b,
                                    Op::LtEq => a <= b,
                                    Op::GtEq => a >= b,
                                    Op::NotEq => a != b,
                                    _ => unreachable!(),
                                }
                            }
                        }
                    };
                    EvalResult::Boolean(result)
                }

                // String concatenation
                Op::Concat => {
                    let left_str = left_result.to_text();
                    let right_str = right_result.to_text();
                    EvalResult::Text(format!("{}{}", left_str, right_str))
                }
            }
        }
    }
}


/// Evaluate function arguments into typed EvalArg values for custom function hooks.
fn eval_function_args<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Vec<EvalArg> {
    args.iter().map(|arg| {
        match arg {
            Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                let min_r = (*start_row).min(*end_row);
                let max_r = (*start_row).max(*end_row);
                let min_c = (*start_col).min(*end_col);
                let max_c = (*start_col).max(*end_col);
                let mut values = Vec::new();
                for r in min_r..=max_r {
                    for c in min_c..=max_c {
                        let val = match sheet {
                            SheetRef::Current => lookup.get_cell_value(r, c),
                            SheetRef::Id(sid) => lookup.get_value_sheet(*sid, r, c),
                            SheetRef::RefError { .. } => Value::Error("#REF!".to_string()),
                        };
                        values.push(val);
                    }
                }
                let n = values.len();
                EvalArg::Range { values, num_cells: n }
            }
            _ => {
                EvalArg::Scalar(evaluate(arg, lookup).to_value())
            }
        }
    }).collect()
}

fn evaluate_function<L: CellLookup>(name: &str, args: &[BoundExpr], lookup: &L) -> EvalResult {
    None
        .or_else(|| super::eval_math::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_logical::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_text::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_conditional::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_lookup::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_financial::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_datetime::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_trig::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_statistical::try_evaluate(name, args, lookup))
        .or_else(|| super::eval_array::try_evaluate(name, args, lookup))
        .or_else(|| {
            let eval_args = eval_function_args(args, lookup);
            lookup.try_custom_function(name, &eval_args)
        })
        .unwrap_or_else(|| EvalResult::Error(format!("Unknown function: {}", name)))
}


#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::parser::{parse, bind_expr_same_sheet, BoundExpr};

    /// Parse and bind a formula for testing (same-sheet formulas only)
    fn parse_and_bind(formula: &str) -> BoundExpr {
        let parsed = parse(formula).unwrap();
        bind_expr_same_sheet(&parsed)
    }

    /// Simple lookup for testing: stores values in a 10x10 grid
    struct TestLookup {
        cells: [[String; 10]; 10],
        named_ranges: std::collections::HashMap<String, NamedRangeResolution>,
    }

    impl TestLookup {
        fn new() -> Self {
            Self {
                cells: Default::default(),
                named_ranges: std::collections::HashMap::new(),
            }
        }

        fn set(&mut self, row: usize, col: usize, value: &str) {
            self.cells[row][col] = value.to_string();
        }

        fn define_cell(&mut self, name: &str, row: usize, col: usize) {
            self.named_ranges.insert(
                name.to_lowercase(),
                NamedRangeResolution::Cell { row, col },
            );
        }

        fn define_range(&mut self, name: &str, start_row: usize, start_col: usize, end_row: usize, end_col: usize) {
            self.named_ranges.insert(
                name.to_lowercase(),
                NamedRangeResolution::Range { start_row, start_col, end_row, end_col },
            );
        }
    }

    impl CellLookup for TestLookup {
        fn get_value(&self, row: usize, col: usize) -> f64 {
            self.cells[row][col].parse().unwrap_or(0.0)
        }

        fn get_text(&self, row: usize, col: usize) -> String {
            self.cells[row][col].clone()
        }

        fn resolve_named_range(&self, name: &str) -> Option<NamedRangeResolution> {
            self.named_ranges.get(&name.to_lowercase()).cloned()
        }
    }

    #[test]
    fn test_named_range_cell_reference() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "42");
        lookup.define_cell("Revenue", 0, 0);

        let expr = parse_and_bind("=Revenue");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(42.0));
    }

    #[test]
    fn test_named_range_undefined() {
        let lookup = TestLookup::new();
        let expr = parse_and_bind("=UndefinedName");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Error(e) => assert!(e.contains("#NAME?")),
            _ => panic!("Expected error for undefined name"),
        }
    }

    #[test]
    fn test_named_range_in_function() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");
        lookup.set(2, 0, "30");
        lookup.define_range("Sales", 0, 0, 2, 0);

        let expr = parse_and_bind("=SUM(Sales)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(60.0));
    }

    #[test]
    fn test_named_range_case_insensitive() {
        let mut lookup = TestLookup::new();
        lookup.set(5, 5, "100");
        lookup.define_cell("MyValue", 5, 5);

        // All of these should resolve to the same cell
        let exprs = ["=MyValue", "=myvalue", "=MYVALUE", "=myVALUE"];
        for formula in exprs {
            let expr = parse_and_bind(formula);
            let result = evaluate(&expr, &lookup);
            assert_eq!(result, EvalResult::Number(100.0), "Failed for {}", formula);
        }
    }

    #[test]
    fn test_named_range_in_arithmetic() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "50");
        lookup.set(1, 1, "25");
        lookup.define_cell("Price", 0, 0);
        lookup.define_cell("Discount", 1, 1);

        let expr = parse_and_bind("=Price-Discount");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(25.0));
    }

    // =========================================================================
    // Compat tests for newly implemented functions (SUMIFS, COUNTIFS, IFNA, TEXTJOIN, XLOOKUP)
    // These ensure Excel-compatible behavior for common import scenarios
    // =========================================================================

    #[test]
    fn test_sumifs_single_criteria() {
        let mut lookup = TestLookup::new();
        // Sum range: A1:A5 = [100, 200, 150, 300, 50]
        lookup.set(0, 0, "100");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "150");
        lookup.set(3, 0, "300");
        lookup.set(4, 0, "50");
        // Criteria range: B1:B5 = ["East", "West", "East", "East", "West"]
        lookup.set(0, 1, "East");
        lookup.set(1, 1, "West");
        lookup.set(2, 1, "East");
        lookup.set(3, 1, "East");
        lookup.set(4, 1, "West");

        // SUMIFS(A1:A5, B1:B5, "East") should sum 100+150+300 = 550
        let expr = parse_and_bind(r#"=SUMIFS(A1:A5, B1:B5, "East")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(550.0));
    }

    #[test]
    fn test_sumifs_multiple_criteria() {
        let mut lookup = TestLookup::new();
        // Sum range: A1:A4 = [100, 200, 150, 300]
        lookup.set(0, 0, "100");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "150");
        lookup.set(3, 0, "300");
        // Region: B1:B4 = ["East", "West", "East", "East"]
        lookup.set(0, 1, "East");
        lookup.set(1, 1, "West");
        lookup.set(2, 1, "East");
        lookup.set(3, 1, "East");
        // Status: C1:C4 = ["Active", "Active", "Inactive", "Active"]
        lookup.set(0, 2, "Active");
        lookup.set(1, 2, "Active");
        lookup.set(2, 2, "Inactive");
        lookup.set(3, 2, "Active");

        // SUMIFS(A1:A4, B1:B4, "East", C1:C4, "Active") = 100 + 300 = 400
        let expr = parse_and_bind(r#"=SUMIFS(A1:A4, B1:B4, "East", C1:C4, "Active")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(400.0));
    }

    #[test]
    fn test_countifs_single_criteria() {
        let mut lookup = TestLookup::new();
        // A1:A5 = ["Apple", "Banana", "Apple", "Cherry", "Apple"]
        lookup.set(0, 0, "Apple");
        lookup.set(1, 0, "Banana");
        lookup.set(2, 0, "Apple");
        lookup.set(3, 0, "Cherry");
        lookup.set(4, 0, "Apple");

        // COUNTIFS(A1:A5, "Apple") = 3
        let expr = parse_and_bind(r#"=COUNTIFS(A1:A5, "Apple")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(3.0));
    }

    #[test]
    fn test_countifs_multiple_criteria() {
        let mut lookup = TestLookup::new();
        // Product: A1:A4 = ["Widget", "Gadget", "Widget", "Widget"]
        lookup.set(0, 0, "Widget");
        lookup.set(1, 0, "Gadget");
        lookup.set(2, 0, "Widget");
        lookup.set(3, 0, "Widget");
        // Region: B1:B4 = ["North", "North", "South", "North"]
        lookup.set(0, 1, "North");
        lookup.set(1, 1, "North");
        lookup.set(2, 1, "South");
        lookup.set(3, 1, "North");

        // COUNTIFS(A1:A4, "Widget", B1:B4, "North") = 2
        let expr = parse_and_bind(r#"=COUNTIFS(A1:A4, "Widget", B1:B4, "North")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(2.0));
    }

    #[test]
    fn test_averageif_basic() {
        let mut lookup = TestLookup::new();
        // A1:A5 = [10, 20, 30, 40, 50]
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");
        lookup.set(2, 0, "30");
        lookup.set(3, 0, "40");
        lookup.set(4, 0, "50");
        // B1:B5 = ["Yes", "No", "Yes", "Yes", "No"]
        lookup.set(0, 1, "Yes");
        lookup.set(1, 1, "No");
        lookup.set(2, 1, "Yes");
        lookup.set(3, 1, "Yes");
        lookup.set(4, 1, "No");

        // AVERAGEIF(B1:B5, "Yes", A1:A5) = (10+30+40)/3 = 26.666...
        let expr = parse_and_bind(r#"=AVERAGEIF(B1:B5, "Yes", A1:A5)"#);
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 26.666666666666668).abs() < 0.0001),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_averageif_numeric_criteria() {
        let mut lookup = TestLookup::new();
        // A1:A5 = [10, 20, 30, 40, 50]
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");
        lookup.set(2, 0, "30");
        lookup.set(3, 0, "40");
        lookup.set(4, 0, "50");

        // AVERAGEIF(A1:A5, ">25") = (30+40+50)/3 = 40
        let expr = parse_and_bind(r#"=AVERAGEIF(A1:A5, ">25")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(40.0));
    }

    #[test]
    fn test_averageif_no_matches() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");

        // AVERAGEIF(A1:A2, ">100") = #DIV/0! (no matches)
        let expr = parse_and_bind(r#"=AVERAGEIF(A1:A2, ">100")"#);
        let result = evaluate(&expr, &lookup);
        assert!(matches!(result, EvalResult::Error(_)));
    }

    #[test]
    fn test_averageifs_single_criteria() {
        let mut lookup = TestLookup::new();
        // Values: A1:A5 = [100, 200, 150, 300, 50]
        lookup.set(0, 0, "100");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "150");
        lookup.set(3, 0, "300");
        lookup.set(4, 0, "50");
        // Region: B1:B5 = ["East", "West", "East", "East", "West"]
        lookup.set(0, 1, "East");
        lookup.set(1, 1, "West");
        lookup.set(2, 1, "East");
        lookup.set(3, 1, "East");
        lookup.set(4, 1, "West");

        // AVERAGEIFS(A1:A5, B1:B5, "East") = (100+150+300)/3 = 183.333...
        let expr = parse_and_bind(r#"=AVERAGEIFS(A1:A5, B1:B5, "East")"#);
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 183.33333333333334).abs() < 0.0001),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_averageifs_multiple_criteria() {
        let mut lookup = TestLookup::new();
        // Values: A1:A4 = [100, 200, 150, 300]
        lookup.set(0, 0, "100");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "150");
        lookup.set(3, 0, "300");
        // Region: B1:B4 = ["East", "West", "East", "East"]
        lookup.set(0, 1, "East");
        lookup.set(1, 1, "West");
        lookup.set(2, 1, "East");
        lookup.set(3, 1, "East");
        // Status: C1:C4 = ["Active", "Active", "Inactive", "Active"]
        lookup.set(0, 2, "Active");
        lookup.set(1, 2, "Active");
        lookup.set(2, 2, "Inactive");
        lookup.set(3, 2, "Active");

        // AVERAGEIFS(A1:A4, B1:B4, "East", C1:C4, "Active") = (100+300)/2 = 200
        let expr = parse_and_bind(r#"=AVERAGEIFS(A1:A4, B1:B4, "East", C1:C4, "Active")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(200.0));
    }

    #[test]
    fn test_averageifs_no_matches() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "100");
        lookup.set(0, 1, "East");

        // AVERAGEIFS(A1:A1, B1:B1, "West") = #DIV/0! (no matches)
        let expr = parse_and_bind(r#"=AVERAGEIFS(A1:A1, B1:B1, "West")"#);
        let result = evaluate(&expr, &lookup);
        assert!(matches!(result, EvalResult::Error(_)));
    }

    // =========================================================================
    // Financial function tests
    // =========================================================================

    #[test]
    fn test_pmt_basic() {
        let lookup = TestLookup::new();
        // PMT(0.05/12, 60, 10000) = -188.71 (monthly payment for $10k loan at 5% for 5 years)
        let expr = parse_and_bind("=PMT(0.05/12, 60, 10000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - (-188.71)).abs() < 0.01),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pmt_with_fv() {
        let lookup = TestLookup::new();
        // PMT(0.06/12, 120, 100000, 50000) - loan with balloon payment
        // Payment to pay off 100k principal plus accumulate 50k at end
        let expr = parse_and_bind("=PMT(0.06/12, 120, 100000, 50000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - (-1415.31)).abs() < 0.01),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pmt_zero_rate() {
        let lookup = TestLookup::new();
        // PMT(0, 12, 1200) = -100 (no interest)
        let expr = parse_and_bind("=PMT(0, 12, 1200)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(-100.0));
    }

    #[test]
    fn test_fv_basic() {
        let lookup = TestLookup::new();
        // FV(0.05/12, 60, -200) - saving $200/month at 5% for 5 years
        let expr = parse_and_bind("=FV(0.05/12, 60, -200)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 13601.22).abs() < 0.01),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_fv_with_pv() {
        let lookup = TestLookup::new();
        // FV(0.08/12, 120, -500, -10000) - monthly savings plus initial deposit
        let expr = parse_and_bind("=FV(0.08/12, 120, -500, -10000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 113669.42).abs() < 0.1),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_fv_zero_rate() {
        let lookup = TestLookup::new();
        // FV(0, 12, -100, -1000) = 2200 (no interest)
        let expr = parse_and_bind("=FV(0, 12, -100, -1000)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(2200.0));
    }

    #[test]
    fn test_pv_basic() {
        let lookup = TestLookup::new();
        // PV(0.08/12, 240, -1000) = 119,554.29 (what loan can I afford with $1000/month payment)
        let expr = parse_and_bind("=PV(0.08/12, 240, -1000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 119554.29).abs() < 0.01),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pv_with_fv() {
        let lookup = TestLookup::new();
        // PV(0.1/12, 60, -100, -5000)
        let expr = parse_and_bind("=PV(0.1/12, 60, -100, -5000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Value depends on formula convention
                assert!(n.is_finite());
            },
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pv_zero_rate() {
        let lookup = TestLookup::new();
        // PV(0, 10, -100, -500) = 1500 (no interest)
        let expr = parse_and_bind("=PV(0, 10, -100, -500)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(1500.0));
    }

    #[test]
    fn test_npv_basic() {
        let mut lookup = TestLookup::new();
        // Cash flows: -10000, 3000, 4000, 5000, 6000
        lookup.set(0, 0, "-10000");
        lookup.set(1, 0, "3000");
        lookup.set(2, 0, "4000");
        lookup.set(3, 0, "5000");
        lookup.set(4, 0, "6000");

        // NPV(0.1, A1:A5) at 10% discount rate
        // = -10000/1.1 + 3000/1.1^2 + 4000/1.1^3 + 5000/1.1^4 + 6000/1.1^5
        let expr = parse_and_bind("=NPV(0.1, A1:A5)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 3534.28).abs() < 0.01),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_npv_investment() {
        let mut lookup = TestLookup::new();
        // Initial investment of -1000, then 300 per year for 5 years
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "300");
        lookup.set(2, 0, "300");
        lookup.set(3, 0, "300");
        lookup.set(4, 0, "300");
        lookup.set(5, 0, "300");

        let expr = parse_and_bind("=NPV(0.08, A1:A6)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!(n.is_finite()), // Value varies by convention
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_irr_basic() {
        let mut lookup = TestLookup::new();
        // Cash flows: -100, 30, 35, 40, 45
        lookup.set(0, 0, "-100");
        lookup.set(1, 0, "30");
        lookup.set(2, 0, "35");
        lookup.set(3, 0, "40");
        lookup.set(4, 0, "45");

        let expr = parse_and_bind("=IRR(A1:A5)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 0.1709).abs() < 0.01), // ~17.09%
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_irr_with_guess() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "300");
        lookup.set(3, 0, "400");
        lookup.set(4, 0, "500");

        let expr = parse_and_bind("=IRR(A1:A5, 0.15)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 0.1283).abs() < 0.001), // ~12.83%
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_irr_no_solution() {
        let mut lookup = TestLookup::new();
        // All positive - no IRR exists
        lookup.set(0, 0, "100");
        lookup.set(1, 0, "200");

        let expr = parse_and_bind("=IRR(A1:A2)");
        let result = evaluate(&expr, &lookup);
        assert!(matches!(result, EvalResult::Error(_)));
    }

    #[test]
    fn test_irr_near_zero_return() {
        let mut lookup = TestLookup::new();
        // Cash flows that yield IRR very close to 0
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "500");
        lookup.set(2, 0, "500");

        let expr = parse_and_bind("=IRR(A1:A3)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!(n.abs() < 0.01, "Expected ~0%, got {}", n),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_irr_large_negative_initial() {
        let mut lookup = TestLookup::new();
        // Large initial outlay with small returns — low IRR
        lookup.set(0, 0, "-10000");
        lookup.set(1, 0, "100");
        lookup.set(2, 0, "100");
        lookup.set(3, 0, "100");
        lookup.set(4, 0, "100");
        lookup.set(5, 0, "100");
        lookup.set(6, 0, "100");
        lookup.set(7, 0, "100");
        lookup.set(8, 0, "100");
        lookup.set(9, 0, "10000");

        let expr = parse_and_bind("=IRR(A1:A10)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Should converge to a valid rate (bisection fallback may be needed)
                assert!(n > -1.0 && n.is_finite(), "Expected valid rate, got {}", n);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_irr_haven_cashflows() {
        // Exact cashflows from Haven Hilgard Pro Forma XLSX
        // Expected IRR ≈ 34.9956% (verified in Excel/Google Sheets)
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "-2037280");
        lookup.set(1, 0, "154388");
        lookup.set(2, 0, "318040");
        lookup.set(3, 0, "327581");
        lookup.set(4, 0, "337409");
        lookup.set(5, 0, "6786109");

        let expr = parse_and_bind("=IRR(A1:A6)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Verify NPV at returned rate is ~0
                let cashflows = [-2037280.0, 154388.0, 318040.0, 327581.0, 337409.0, 6786109.0];
                let npv: f64 = cashflows.iter().enumerate()
                    .map(|(i, &cf)| cf / (1.0 + n).powf(i as f64))
                    .sum();
                assert!(npv.abs() < 0.01,
                    "IRR={:.6} but NPV(IRR)={:.6} (should be ~0)", n, npv);
                assert!((n - 0.3500).abs() < 0.005,
                    "Expected IRR ~35.00%, got {:.4}%", n * 100.0);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_irr_error_propagation() {
        // If a cell in the range has an error, IRR should propagate it
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "#DIV/0!");  // Error cell
        lookup.set(3, 0, "400");

        let expr = parse_and_bind("=IRR(A1:A4)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Error(e) => {
                assert_eq!(e, "#DIV/0!", "Expected #DIV/0! propagation, got {:?}", e);
            }
            other => panic!("Expected error propagation, got {:?}", other),
        }
    }

    #[test]
    fn test_ifna_with_na_error() {
        let mut lookup = TestLookup::new();
        // VLOOKUP that won't find match returns #N/A
        lookup.set(0, 0, "NotFound");
        lookup.set(0, 1, "100");

        // IFNA(VLOOKUP("X", A1:B1, 2, FALSE), "Not Available")
        let expr = parse_and_bind(r#"=IFNA(VLOOKUP("X", A1:B1, 2, FALSE), "Not Available")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("Not Available".to_string()));
    }

    #[test]
    fn test_ifna_without_error() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "Found");
        lookup.set(0, 1, "100");

        // IFNA(VLOOKUP("Found", A1:B1, 2, FALSE), "Not Available")
        let expr = parse_and_bind(r#"=IFNA(VLOOKUP("Found", A1:B1, 2, FALSE), "Not Available")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(100.0));
    }

    #[test]
    fn test_isna_with_na_error() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "NotFound");
        lookup.set(0, 1, "100");

        // ISNA(VLOOKUP("X", A1:B1, 2, FALSE)) - should be TRUE (returns #N/A)
        let expr = parse_and_bind(r#"=ISNA(VLOOKUP("X", A1:B1, 2, FALSE))"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Boolean(true));
    }

    #[test]
    fn test_isna_with_other_error() {
        let lookup = TestLookup::new();

        // ISNA(1/0) - should be FALSE (division by zero is not #N/A)
        let expr = parse_and_bind("=ISNA(1/0)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Boolean(false));
    }

    #[test]
    fn test_isna_with_value() {
        let lookup = TestLookup::new();

        // ISNA(42) - should be FALSE (not an error)
        let expr = parse_and_bind("=ISNA(42)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Boolean(false));
    }

    #[test]
    fn test_textjoin_basic() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "Hello");
        lookup.set(1, 0, "World");
        lookup.set(2, 0, "!");

        // TEXTJOIN(", ", TRUE, A1:A3) = "Hello, World, !"
        let expr = parse_and_bind(r#"=TEXTJOIN(", ", TRUE, A1:A3)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("Hello, World, !".to_string()));
    }

    #[test]
    fn test_textjoin_ignore_empty() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "A");
        lookup.set(1, 0, "");  // Empty
        lookup.set(2, 0, "B");
        lookup.set(3, 0, "");  // Empty
        lookup.set(4, 0, "C");

        // TEXTJOIN("-", TRUE, A1:A5) = "A-B-C" (empties ignored)
        let expr = parse_and_bind(r#"=TEXTJOIN("-", TRUE, A1:A5)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("A-B-C".to_string()));

        // TEXTJOIN("-", FALSE, A1:A5) = "A--B--C" (empties included)
        let expr = parse_and_bind(r#"=TEXTJOIN("-", FALSE, A1:A5)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("A--B--C".to_string()));
    }

    #[test]
    fn test_xlookup_basic() {
        let mut lookup = TestLookup::new();
        // Lookup column A1:A4 = ["Apple", "Banana", "Cherry", "Date"]
        lookup.set(0, 0, "Apple");
        lookup.set(1, 0, "Banana");
        lookup.set(2, 0, "Cherry");
        lookup.set(3, 0, "Date");
        // Return column B1:B4 = [1.99, 0.99, 2.49, 3.99]
        lookup.set(0, 1, "1.99");
        lookup.set(1, 1, "0.99");
        lookup.set(2, 1, "2.49");
        lookup.set(3, 1, "3.99");

        // XLOOKUP("Cherry", A1:A4, B1:B4) = 2.49
        let expr = parse_and_bind(r#"=XLOOKUP("Cherry", A1:A4, B1:B4)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(2.49));
    }

    #[test]
    fn test_xlookup_not_found_with_default() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "Apple");
        lookup.set(0, 1, "1.99");

        // XLOOKUP("Orange", A1:A1, B1:B1, "Not found") = "Not found"
        let expr = parse_and_bind(r#"=XLOOKUP("Orange", A1:A1, B1:B1, "Not found")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("Not found".to_string()));
    }

    #[test]
    fn test_xlookup_not_found_returns_na() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "Apple");
        lookup.set(0, 1, "1.99");

        // XLOOKUP("Orange", A1:A1, B1:B1) with no default = #N/A
        let expr = parse_and_bind(r#"=XLOOKUP("Orange", A1:A1, B1:B1)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Error("#N/A".to_string()));
    }

    #[test]
    fn test_xlookup_case_insensitive() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "APPLE");
        lookup.set(0, 1, "1.99");

        // XLOOKUP("apple", ...) should match "APPLE"
        let expr = parse_and_bind(r#"=XLOOKUP("apple", A1:A1, B1:B1)"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(1.99));
    }

    #[test]
    fn test_xlookup_numeric() {
        let mut lookup = TestLookup::new();
        // Lookup by ID: A1:A3 = [101, 102, 103]
        lookup.set(0, 0, "101");
        lookup.set(1, 0, "102");
        lookup.set(2, 0, "103");
        // Names: B1:B3 = ["Alice", "Bob", "Charlie"]
        lookup.set(0, 1, "Alice");
        lookup.set(1, 1, "Bob");
        lookup.set(2, 1, "Charlie");

        // XLOOKUP(102, A1:A3, B1:B3) = "Bob"
        let expr = parse_and_bind("=XLOOKUP(102, A1:A3, B1:B3)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Text("Bob".to_string()));
    }

    // =========================================================================
    // SPARKLINE tests
    // =========================================================================

    #[test]
    fn test_sparkline_bar_basic() {
        let mut lookup = TestLookup::new();
        // Values: 1, 2, 3, 4, 5, 6, 7, 8 (linear progression)
        for i in 0..8 {
            lookup.set(0, i, &format!("{}", i + 1));
        }

        let expr = parse_and_bind("=SPARKLINE(A1:H1)");
        let result = evaluate(&expr, &lookup);
        // Should produce 8 bars from lowest to highest
        assert_eq!(result, EvalResult::Text("▁▂▃▄▅▆▇█".to_string()));
    }

    #[test]
    fn test_sparkline_bar_reverse() {
        let mut lookup = TestLookup::new();
        // Values: 8, 7, 6, 5, 4, 3, 2, 1 (reverse)
        for i in 0..8 {
            lookup.set(0, i, &format!("{}", 8 - i));
        }

        let expr = parse_and_bind("=SPARKLINE(A1:H1)");
        let result = evaluate(&expr, &lookup);
        // Should produce 8 bars from highest to lowest
        assert_eq!(result, EvalResult::Text("█▇▆▅▄▃▂▁".to_string()));
    }

    #[test]
    fn test_sparkline_flat_data() {
        let mut lookup = TestLookup::new();
        // All values are 5
        for i in 0..4 {
            lookup.set(0, i, "5");
        }

        let expr = parse_and_bind("=SPARKLINE(A1:D1)");
        let result = evaluate(&expr, &lookup);
        // All same value - should show middle bars
        assert_eq!(result, EvalResult::Text("▄▄▄▄".to_string()));
    }

    #[test]
    fn test_sparkline_winloss() {
        let mut lookup = TestLookup::new();
        // Values: 10, -5, 0, 15, -3
        lookup.set(0, 0, "10");
        lookup.set(0, 1, "-5");
        lookup.set(0, 2, "0");
        lookup.set(0, 3, "15");
        lookup.set(0, 4, "-3");

        let expr = parse_and_bind(r#"=SPARKLINE(A1:E1, "winloss")"#);
        let result = evaluate(&expr, &lookup);
        // Positive = ▲, Negative = ▼, Zero = ▬
        assert_eq!(result, EvalResult::Text("▲▼▬▲▼".to_string()));
    }

    #[test]
    fn test_sparkline_empty_range() {
        let lookup = TestLookup::new();
        // Empty range
        let expr = parse_and_bind("=SPARKLINE(A1:A1)");
        let result = evaluate(&expr, &lookup);
        // Empty input returns empty string
        assert_eq!(result, EvalResult::Text("".to_string()));
    }

    #[test]
    fn test_sparkline_invalid_type() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "1");
        lookup.set(0, 1, "2");

        let expr = parse_and_bind(r#"=SPARKLINE(A1:B1, "invalid")"#);
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Error(e) => assert!(e.contains("Unknown sparkline type")),
            _ => panic!("Expected error for invalid sparkline type"),
        }
    }

    #[test]
    fn test_sparkline_single_value() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "42");

        let expr = parse_and_bind("=SPARKLINE(A1:A1)");
        let result = evaluate(&expr, &lookup);
        // Single value - should show middle bar
        assert_eq!(result, EvalResult::Text("▄".to_string()));
    }

    // =========================================================================
    // Power (^) and Percent (%) operator tests
    // =========================================================================

    #[test]
    fn test_power_simple() {
        let lookup = TestLookup::new();
        let expr = parse_and_bind("=2^3");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(8.0));
    }

    #[test]
    fn test_power_right_associative() {
        let lookup = TestLookup::new();
        // 2^3^2 should be 2^(3^2) = 2^9 = 512 (right-associative)
        let expr = parse_and_bind("=2^3^2");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(512.0));
    }

    #[test]
    fn test_power_precedence_over_multiply() {
        let lookup = TestLookup::new();
        // 3*2^3 = 3*(2^3) = 3*8 = 24
        let expr = parse_and_bind("=3*2^3");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(24.0));
    }

    #[test]
    fn test_power_fractional_exponent() {
        let lookup = TestLookup::new();
        // 9^0.5 = 3
        let expr = parse_and_bind("=9^0.5");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-10),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_percent_simple() {
        let lookup = TestLookup::new();
        // 50% = 0.5
        let expr = parse_and_bind("=50%");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(0.5));
    }

    #[test]
    fn test_percent_in_expression() {
        let lookup = TestLookup::new();
        // 100*5% = 100*0.05 = 5
        let expr = parse_and_bind("=100*5%");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(5.0));
    }

    #[test]
    fn test_percent_double() {
        let lookup = TestLookup::new();
        // 500%% = 500 * 0.01 * 0.01 = 0.05
        let expr = parse_and_bind("=500%%");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(0.05));
    }

    #[test]
    fn test_percent_with_power() {
        let lookup = TestLookup::new();
        // 50%^2 = (0.5)^2 = 0.25 (% binds tighter than ^)
        let expr = parse_and_bind("=50%^2");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(0.25));
    }

    #[test]
    fn test_negative_percent() {
        let lookup = TestLookup::new();
        // -5% = (-5)*0.01 = -0.05
        let expr = parse_and_bind("=-5%");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(-0.05));
    }

    #[test]
    fn test_power_with_cell_ref() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "3"); // A1 = 3
        let expr = parse_and_bind("=A1^2");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(9.0));
    }

    #[test]
    fn test_percent_with_cell_ref() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "200"); // A1 = 200
        // A1*5% = 200*0.05 = 10
        let expr = parse_and_bind("=A1*5%");
        assert_eq!(evaluate(&expr, &lookup), EvalResult::Number(10.0));
    }

    // =========================================================================
    // IPMT, PPMT, CUMPRINC, CUMIPMT tests
    // =========================================================================

    #[test]
    fn test_ipmt_first_period() {
        let lookup = TestLookup::new();
        // IPMT(0.01, 1, 12, 100000) - interest on first period of 100k loan at 1%/period
        let expr = parse_and_bind("=IPMT(0.01, 1, 12, 100000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => assert!((n - (-1000.0)).abs() < 0.01, "Expected ~-1000, got {}", n),
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_ipmt_later_period() {
        let lookup = TestLookup::new();
        // IPMT decreases over time as principal is paid down
        let expr1 = parse_and_bind("=IPMT(0.01, 1, 12, 100000)");
        let expr6 = parse_and_bind("=IPMT(0.01, 6, 12, 100000)");
        let r1 = evaluate(&expr1, &lookup);
        let r6 = evaluate(&expr6, &lookup);
        match (r1, r6) {
            (EvalResult::Number(n1), EvalResult::Number(n6)) => {
                // Interest should decrease as principal is paid
                assert!(n6 > n1, "Interest should decrease (become less negative): per1={}, per6={}", n1, n6);
            }
            _ => panic!("Expected numbers"),
        }
    }

    #[test]
    fn test_ppmt_first_period() {
        let lookup = TestLookup::new();
        // PPMT(0.01, 1, 12, 100000) - principal portion of first payment
        let expr = parse_and_bind("=PPMT(0.01, 1, 12, 100000)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // PMT ~= -8884.88, IPMT(1) = -1000, so PPMT(1) ~= -7884.88
                assert!((n - (-7884.88)).abs() < 1.0, "Expected ~-7884.88, got {}", n);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_ipmt_plus_ppmt_equals_pmt() {
        let lookup = TestLookup::new();
        // IPMT + PPMT should equal PMT for any period
        let pmt_expr = parse_and_bind("=PMT(0.01, 12, 100000)");
        let ipmt_expr = parse_and_bind("=IPMT(0.01, 5, 12, 100000)");
        let ppmt_expr = parse_and_bind("=PPMT(0.01, 5, 12, 100000)");

        let pmt = evaluate(&pmt_expr, &lookup);
        let ipmt = evaluate(&ipmt_expr, &lookup);
        let ppmt = evaluate(&ppmt_expr, &lookup);

        match (pmt, ipmt, ppmt) {
            (EvalResult::Number(p), EvalResult::Number(i), EvalResult::Number(pp)) => {
                assert!((p - (i + pp)).abs() < 1e-6, "PMT({}) != IPMT({}) + PPMT({})", p, i, pp);
            }
            _ => panic!("Expected numbers"),
        }
    }

    #[test]
    fn test_cumprinc_full_loan() {
        let lookup = TestLookup::new();
        // CUMPRINC over the entire loan should equal -(loan amount)
        // CUMPRINC(0.01, 12, 100000, 1, 12, 0)
        let expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 1, 12, 0)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Total principal paid = -pv = -100000
                assert!((n - (-100000.0)).abs() < 0.01, "Expected ~-100000, got {}", n);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_cumprinc_partial() {
        let lookup = TestLookup::new();
        // CUMPRINC for first 3 periods of a 12-period loan
        let expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 1, 3, 0)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Should be negative (principal outflow) and less than total
                assert!(n < 0.0, "Expected negative, got {}", n);
                assert!(n > -100000.0, "Expected partial, got {}", n);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_cumipmt_full_loan() {
        let lookup = TestLookup::new();
        // CUMIPMT over entire loan = total interest paid
        let expr = parse_and_bind("=CUMIPMT(0.01, 12, 100000, 1, 12, 0)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Total payments = 12 * PMT(0.01,12,100000) ≈ 12 * (-8884.88) ≈ -106618.6
                // Total interest = total payments - (-100000) ≈ -6618.6
                assert!(n < 0.0, "Expected negative interest, got {}", n);
                assert!((n - (-6618.55)).abs() < 1.0, "Expected ~-6618.55, got {}", n);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_cumprinc_plus_cumipmt_equals_total_payments() {
        let lookup = TestLookup::new();
        // CUMPRINC + CUMIPMT = total PMT * nperiods for the same range
        let cp_expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 1, 12, 0)");
        let ci_expr = parse_and_bind("=CUMIPMT(0.01, 12, 100000, 1, 12, 0)");
        let pmt_expr = parse_and_bind("=PMT(0.01, 12, 100000)");

        let cp = evaluate(&cp_expr, &lookup);
        let ci = evaluate(&ci_expr, &lookup);
        let pmt = evaluate(&pmt_expr, &lookup);

        match (cp, ci, pmt) {
            (EvalResult::Number(c), EvalResult::Number(i), EvalResult::Number(p)) => {
                let total = c + i;
                let expected = p * 12.0;
                assert!((total - expected).abs() < 0.01,
                    "CUMPRINC({}) + CUMIPMT({}) = {} != PMT*12 = {}", c, i, total, expected);
            }
            _ => panic!("Expected numbers"),
        }
    }

    #[test]
    fn test_cumprinc_invalid_args() {
        let lookup = TestLookup::new();
        // start > end
        let expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 5, 3, 0)");
        assert!(matches!(evaluate(&expr, &lookup), EvalResult::Error(_)));

        // start < 1
        let expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 0, 3, 0)");
        assert!(matches!(evaluate(&expr, &lookup), EvalResult::Error(_)));

        // end > nper
        let expr = parse_and_bind("=CUMPRINC(0.01, 12, 100000, 1, 15, 0)");
        assert!(matches!(evaluate(&expr, &lookup), EvalResult::Error(_)));
    }

    // =========================================================================
    // Excel golden value tests
    //
    // These tests pin exact outputs matching Excel's financial functions.
    // Values computed from Excel's documented formulas at full precision.
    // Pass criterion: within 0.01 for dollar amounts, 1e-6 for rates.
    // =========================================================================

    /// Helper: assert a formula evaluates to a number within epsilon of expected
    fn assert_excel(formula: &str, lookup: &TestLookup, expected: f64, epsilon: f64) {
        let expr = parse_and_bind(formula);
        let result = evaluate(&expr, lookup);
        match result {
            EvalResult::Number(n) => {
                assert!(
                    (n - expected).abs() < epsilon,
                    "Formula: {}\nExpected: {:.10}\nGot:      {:.10}\nDiff:     {:.2e}",
                    formula, expected, n, (n - expected).abs()
                );
            }
            other => panic!("Formula: {}\nExpected Number({}), got {:?}", formula, expected, other),
        }
    }

    // --- Test case 1: $100,000 annual loan at 6% for 10 years ---
    // Source: FinanceTrain, Excel documentation

    #[test]
    fn test_excel_pmt_annual_loan() {
        let lookup = TestLookup::new();
        // Excel: PMT(0.06, 10, 100000) = -13586.795524...
        assert_excel("=PMT(0.06, 10, 100000)", &lookup, -13586.7955238, 0.01);
    }

    #[test]
    fn test_excel_ipmt_period_1() {
        let lookup = TestLookup::new();
        // Excel: IPMT(0.06, 1, 10, 100000) = -6000.00
        // First period interest = loan_amount * rate
        assert_excel("=IPMT(0.06, 1, 10, 100000)", &lookup, -6000.00, 0.01);
    }

    #[test]
    fn test_excel_ipmt_period_3() {
        let lookup = TestLookup::new();
        // Excel: IPMT(0.06, 3, 10, 100000) = -5062.27...
        assert_excel("=IPMT(0.06, 3, 10, 100000)", &lookup, -5062.27, 0.01);
    }

    #[test]
    fn test_excel_ppmt_period_3() {
        let lookup = TestLookup::new();
        // Excel: PPMT(0.06, 3, 10, 100000) = -8524.52...
        assert_excel("=PPMT(0.06, 3, 10, 100000)", &lookup, -8524.52, 0.01);
    }

    #[test]
    fn test_excel_cumprinc_periods_3_to_6() {
        let lookup = TestLookup::new();
        // Excel: CUMPRINC(0.06, 10, 100000, 3, 6, 0) = -37291.52...
        assert_excel("=CUMPRINC(0.06, 10, 100000, 3, 6, 0)", &lookup, -37291.52, 0.01);
    }

    #[test]
    fn test_excel_cumprinc_full_term() {
        let lookup = TestLookup::new();
        // Total principal over full loan = -pv
        assert_excel("=CUMPRINC(0.06, 10, 100000, 1, 10, 0)", &lookup, -100000.00, 0.01);
    }

    #[test]
    fn test_excel_cumipmt_full_term() {
        let lookup = TestLookup::new();
        // Total interest = nper * PMT - (-pv) = 10 * (-13586.80) + 100000 = -35867.96...
        assert_excel("=CUMIPMT(0.06, 10, 100000, 1, 10, 0)", &lookup, -35867.95, 0.01);
    }

    // --- Test case 2: $10,000 monthly loan at 5%/12 for 60 months ---
    // Source: Exceljet, Wall Street Prep

    #[test]
    fn test_excel_pmt_monthly_loan() {
        let lookup = TestLookup::new();
        // Excel: PMT(0.05/12, 60, 10000) = -188.712...
        // Monthly rate = 0.004166667
        assert_excel("=PMT(0.05/12, 60, 10000)", &lookup, -188.71, 0.01);
    }

    #[test]
    fn test_excel_ipmt_monthly_period_1() {
        let lookup = TestLookup::new();
        // Excel: IPMT(0.05/12, 1, 60, 10000) = -41.67
        // First month interest = 10000 * 0.05/12
        assert_excel("=IPMT(0.05/12, 1, 60, 10000)", &lookup, -41.67, 0.01);
    }

    #[test]
    fn test_excel_ppmt_monthly_period_1() {
        let lookup = TestLookup::new();
        // Excel: PPMT(0.05/12, 1, 60, 10000) = PMT - IPMT = -188.71 - (-41.67) = -147.05
        assert_excel("=PPMT(0.05/12, 1, 60, 10000)", &lookup, -147.05, 0.01);
    }

    #[test]
    fn test_excel_cumprinc_monthly_full() {
        let lookup = TestLookup::new();
        // Total principal over full loan = -10000
        assert_excel("=CUMPRINC(0.05/12, 60, 10000, 1, 60, 0)", &lookup, -10000.00, 0.01);
    }

    // --- Test case 3: PMT edge cases ---

    #[test]
    fn test_excel_pmt_zero_rate() {
        let lookup = TestLookup::new();
        // PMT(0, 12, 1200) = -100.00
        assert_excel("=PMT(0, 12, 1200)", &lookup, -100.00, 0.01);
    }

    #[test]
    fn test_excel_pmt_with_fv() {
        let lookup = TestLookup::new();
        // PMT(0.08/12, 120, 0, 100000) — saving for $100K future value
        // Excel: PMT(0.00666667, 120, 0, 100000) = -546.608...
        assert_excel("=PMT(0.08/12, 120, 0, 100000)", &lookup, -546.61, 0.01);
    }

    // --- Test case 4: IRR golden values ---

    #[test]
    fn test_excel_irr_standard() {
        let mut lookup = TestLookup::new();
        // {-100, 30, 35, 40, 45} — IRR ≈ 17.09%
        // Verify by computing NPV at the returned rate ≈ 0
        lookup.set(0, 0, "-100");
        lookup.set(1, 0, "30");
        lookup.set(2, 0, "35");
        lookup.set(3, 0, "40");
        lookup.set(4, 0, "45");
        let expr = parse_and_bind("=IRR(A1:A5)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Verify NPV at returned rate is ~0
                let cashflows = [-100.0, 30.0, 35.0, 40.0, 45.0];
                let npv: f64 = cashflows.iter().enumerate()
                    .map(|(i, &cf)| cf / (1.0 + n).powf(i as f64))
                    .sum();
                assert!(npv.abs() < 1e-6,
                    "IRR={:.10} but NPV(IRR)={:.10} (should be ~0)", n, npv);
                assert!((n - 0.1709).abs() < 0.001,
                    "IRR expected ~0.1709, got {}", n);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_excel_irr_even_cashflows() {
        let mut lookup = TestLookup::new();
        // {-1000, 300, 300, 300, 300, 300} — annuity IRR
        // NPV = -1000 + 300/r * (1 - 1/(1+r)^5) = 0
        // Excel: IRR ≈ 15.24%
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "300");
        lookup.set(2, 0, "300");
        lookup.set(3, 0, "300");
        lookup.set(4, 0, "300");
        lookup.set(5, 0, "300");
        let expr = parse_and_bind("=IRR(A1:A6)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                assert!((n - 0.1524).abs() < 0.001,
                    "IRR expected ~0.1524, got {}", n);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_excel_irr_with_guess() {
        let mut lookup = TestLookup::new();
        // Same cashflows, explicit guess = 0.2
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "200");
        lookup.set(2, 0, "300");
        lookup.set(3, 0, "400");
        lookup.set(4, 0, "500");
        let expr = parse_and_bind("=IRR(A1:A5, 0.2)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                // Should still converge to same root regardless of guess
                assert!((n - 0.1283).abs() < 0.001,
                    "IRR expected ~0.1283, got {}", n);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    #[test]
    fn test_excel_irr_negative_result() {
        let mut lookup = TestLookup::new();
        // Losing investment: {-1000, 100, 200, 300}
        // Total return = 600 < 1000 → negative IRR
        lookup.set(0, 0, "-1000");
        lookup.set(1, 0, "100");
        lookup.set(2, 0, "200");
        lookup.set(3, 0, "300");
        let expr = parse_and_bind("=IRR(A1:A4)");
        let result = evaluate(&expr, &lookup);
        match result {
            EvalResult::Number(n) => {
                assert!(n < 0.0, "Expected negative IRR, got {}", n);
                assert!(n > -1.0, "IRR should be > -1, got {}", n);
            }
            other => panic!("Expected number, got {:?}", other),
        }
    }

    // --- Test case 5: IPMT + PPMT = PMT identity across all periods ---

    #[test]
    fn test_excel_ipmt_ppmt_sum_all_periods() {
        let lookup = TestLookup::new();
        // Verify IPMT(per) + PPMT(per) = PMT for every period
        // $100,000 at 6% for 10 years
        for per in 1..=10 {
            let ipmt_formula = format!("=IPMT(0.06, {}, 10, 100000)", per);
            let ppmt_formula = format!("=PPMT(0.06, {}, 10, 100000)", per);

            let ipmt_expr = parse_and_bind(&ipmt_formula);
            let ppmt_expr = parse_and_bind(&ppmt_formula);
            let pmt_expr = parse_and_bind("=PMT(0.06, 10, 100000)");

            let ipmt = evaluate(&ipmt_expr, &lookup);
            let ppmt = evaluate(&ppmt_expr, &lookup);
            let pmt = evaluate(&pmt_expr, &lookup);

            match (ipmt, ppmt, pmt) {
                (EvalResult::Number(i), EvalResult::Number(p), EvalResult::Number(total)) => {
                    assert!(
                        (i + p - total).abs() < 1e-8,
                        "Period {}: IPMT({:.6}) + PPMT({:.6}) = {:.6} != PMT({:.6})",
                        per, i, p, i + p, total
                    );
                }
                _ => panic!("Period {}: expected numbers", per),
            }
        }
    }

    // --- Test case 6: CUMPRINC + CUMIPMT = total payments identity ---

    #[test]
    fn test_excel_cum_identity_partial() {
        let lookup = TestLookup::new();
        // CUMPRINC(3,6) + CUMIPMT(3,6) = PMT * 4 (periods 3 through 6)
        let cp = parse_and_bind("=CUMPRINC(0.06, 10, 100000, 3, 6, 0)");
        let ci = parse_and_bind("=CUMIPMT(0.06, 10, 100000, 3, 6, 0)");
        let pmt = parse_and_bind("=PMT(0.06, 10, 100000)");

        let cp_val = evaluate(&cp, &lookup);
        let ci_val = evaluate(&ci, &lookup);
        let pmt_val = evaluate(&pmt, &lookup);

        match (cp_val, ci_val, pmt_val) {
            (EvalResult::Number(c), EvalResult::Number(i), EvalResult::Number(p)) => {
                let total = c + i;
                let expected = p * 4.0; // 4 periods
                assert!(
                    (total - expected).abs() < 0.01,
                    "CUMPRINC + CUMIPMT = {:.2} != PMT*4 = {:.2}",
                    total, expected
                );
            }
            _ => panic!("Expected numbers"),
        }
    }

    // --- Test case 7: Power and percent in financial context ---

    #[test]
    fn test_excel_compound_growth() {
        let lookup = TestLookup::new();
        // 1000*(1+5%)^10 = 1000 * 1.05^10 = 1628.89...
        assert_excel("=1000*(1+5%)^10", &lookup, 1628.89, 0.01);
    }

    #[test]
    fn test_excel_discount_factor() {
        let lookup = TestLookup::new();
        // 1/(1+8%)^5 = 1/1.08^5 = 0.6806...
        assert_excel("=1/(1+8%)^5", &lookup, 0.6806, 0.0001);
    }

    // =========================================================================
    // COUNTIF/SUMIF with single cell refs and named ranges
    // =========================================================================

    #[test]
    fn test_countif_single_cell_ref() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "hello");  // A1
        // COUNTIF(A1, "hello") — single cell as 1x1 range
        let expr = parse_and_bind(r#"=COUNTIF(A1, "hello")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(1.0));
    }

    #[test]
    fn test_countif_single_cell_no_match() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "world");  // A1
        let expr = parse_and_bind(r#"=COUNTIF(A1, "hello")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(0.0));
    }

    #[test]
    fn test_countif_named_range() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");
        lookup.set(2, 0, "30");
        lookup.define_range("Data", 0, 0, 2, 0);
        let expr = parse_and_bind(r#"=COUNTIF(Data, ">15")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(2.0));
    }

    #[test]
    fn test_countif_named_range_single_cell() {
        let mut lookup = TestLookup::new();
        lookup.set(3, 3, "42");
        lookup.define_cell("Target", 3, 3);
        let expr = parse_and_bind(r#"=COUNTIF(Target, "42")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(1.0));
    }

    #[test]
    fn test_sumif_single_cell_ref() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "50");  // A1
        // SUMIF(A1, ">0") — single cell, matches
        let expr = parse_and_bind(r#"=SUMIF(A1, ">0")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(50.0));
    }

    #[test]
    fn test_sumif_named_range() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "10");
        lookup.set(1, 0, "20");
        lookup.set(2, 0, "5");
        lookup.define_range("Amounts", 0, 0, 2, 0);
        let expr = parse_and_bind(r#"=SUMIF(Amounts, ">=10")"#);
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(30.0));
    }

    #[test]
    fn test_countblank_single_cell_ref() {
        let mut lookup = TestLookup::new();
        // A1 is empty
        let expr = parse_and_bind("=COUNTBLANK(A1)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(1.0));

        // A1 has value → not blank
        lookup.set(0, 0, "x");
        let expr = parse_and_bind("=COUNTBLANK(A1)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(0.0));
    }

    // =========================================================================
    // SUMPRODUCT tests
    // =========================================================================

    #[test]
    fn test_sumproduct_two_ranges_1d() {
        let mut lookup = TestLookup::new();
        // A1:A3 = [1, 2, 3], B1:B3 = [4, 5, 6]
        lookup.set(0, 0, "1"); lookup.set(1, 0, "2"); lookup.set(2, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5"); lookup.set(2, 1, "6");
        // 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3, B1:B3)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(32.0));
    }

    #[test]
    fn test_sumproduct_single_range() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "10"); lookup.set(1, 0, "20"); lookup.set(2, 0, "30");
        // Single range = sum of elements: 10 + 20 + 30 = 60
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(60.0));
    }

    #[test]
    fn test_sumproduct_three_ranges() {
        let mut lookup = TestLookup::new();
        // A1:A2 = [2, 3], B1:B2 = [4, 5], C1:C2 = [10, 20]
        lookup.set(0, 0, "2"); lookup.set(1, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5");
        lookup.set(0, 2, "10"); lookup.set(1, 2, "20");
        // 2*4*10 + 3*5*20 = 80 + 300 = 380
        let expr = parse_and_bind("=SUMPRODUCT(A1:A2, B1:B2, C1:C2)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(380.0));
    }

    #[test]
    fn test_sumproduct_2d_ranges() {
        let mut lookup = TestLookup::new();
        // A1:B2 = [[1,2],[3,4]], C1:D2 = [[5,6],[7,8]]
        lookup.set(0, 0, "1"); lookup.set(0, 1, "2");
        lookup.set(1, 0, "3"); lookup.set(1, 1, "4");
        lookup.set(0, 2, "5"); lookup.set(0, 3, "6");
        lookup.set(1, 2, "7"); lookup.set(1, 3, "8");
        // 1*5 + 2*6 + 3*7 + 4*8 = 5 + 12 + 21 + 32 = 70
        let expr = parse_and_bind("=SUMPRODUCT(A1:B2, C1:D2)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(70.0));
    }

    #[test]
    fn test_sumproduct_with_blanks() {
        let mut lookup = TestLookup::new();
        // A1:A3 = [1, "", 3], B1:B3 = [4, 5, 6]
        lookup.set(0, 0, "1"); /* A2 empty */ lookup.set(2, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5"); lookup.set(2, 1, "6");
        // 1*4 + 0*5 + 3*6 = 4 + 0 + 18 = 22
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3, B1:B3)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(22.0));
    }

    #[test]
    fn test_sumproduct_with_text() {
        let mut lookup = TestLookup::new();
        // A1:A3 = [1, "hello", 3], B1:B3 = [4, 5, 6]
        lookup.set(0, 0, "1"); lookup.set(1, 0, "hello"); lookup.set(2, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5"); lookup.set(2, 1, "6");
        // 1*4 + 0*5 + 3*6 = 4 + 0 + 18 = 22 (text → 0)
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3, B1:B3)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(22.0));
    }

    #[test]
    fn test_sumproduct_shape_mismatch() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "1"); lookup.set(1, 0, "2"); lookup.set(2, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5");
        // A1:A3 (3x1) vs B1:B2 (2x1) → shape mismatch
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3, B1:B2)");
        let result = evaluate(&expr, &lookup);
        assert!(matches!(result, EvalResult::Error(_)));
    }

    #[test]
    fn test_sumproduct_named_range() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "2"); lookup.set(1, 0, "3");
        lookup.set(0, 1, "10"); lookup.set(1, 1, "20");
        lookup.define_range("Prices", 0, 0, 1, 0);
        lookup.define_range("Quantities", 0, 1, 1, 1);
        // 2*10 + 3*20 = 20 + 60 = 80
        let expr = parse_and_bind("=SUMPRODUCT(Prices, Quantities)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(80.0));
    }

    #[test]
    fn test_sumproduct_single_cell_refs() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "5"); lookup.set(0, 1, "7");
        // SUMPRODUCT(A1, B1) = 5*7 = 35
        let expr = parse_and_bind("=SUMPRODUCT(A1, B1)");
        let result = evaluate(&expr, &lookup);
        assert_eq!(result, EvalResult::Number(35.0));
    }

    #[test]
    fn test_sumproduct_error_propagation() {
        let mut lookup = TestLookup::new();
        lookup.set(0, 0, "1"); lookup.set(1, 0, "#DIV/0!"); lookup.set(2, 0, "3");
        lookup.set(0, 1, "4"); lookup.set(1, 1, "5"); lookup.set(2, 1, "6");
        // Error in A2 → propagates
        let expr = parse_and_bind("=SUMPRODUCT(A1:A3, B1:B3)");
        let result = evaluate(&expr, &lookup);
        assert!(matches!(result, EvalResult::Error(_)));
    }
}
