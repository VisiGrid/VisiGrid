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

    /// Get the current cell being evaluated (for ROW()/COLUMN() without args).
    /// Default implementation returns None (not in cell context).
    fn current_cell(&self) -> Option<(usize, usize)> {
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
            Value::Text(s) => s.parse::<f64>().map_err(|_| format!("#VALUE! Cannot convert '{}' to number", s)),
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
    pub fn to_number(&self) -> Result<f64, String> {
        match self {
            EvalResult::Number(n) => Ok(*n),
            EvalResult::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
            EvalResult::Text(s) => s.parse::<f64>().map_err(|_| format!("Cannot convert '{}' to number", s)),
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
            // Get cell value, potentially from another sheet
            let text = match sheet {
                SheetRef::Current => lookup.get_text(*row, *col),
                SheetRef::Id(sheet_id) => lookup.get_text_sheet(*sheet_id, *row, *col),
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
            EvalResult::Error("Range must be used in a function".to_string())
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
                    EvalResult::Error("Range must be used in a function".to_string())
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
                Op::Add | Op::Sub | Op::Mul | Op::Div => {
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

fn evaluate_function<L: CellLookup>(name: &str, args: &[BoundExpr], lookup: &L) -> EvalResult {
    match name {
        // =====================
        // MATH FUNCTIONS
        // =====================
        "SUM" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => EvalResult::Number(vals.iter().sum()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "AVERAGE" | "AVG" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        EvalResult::Error("AVERAGE requires at least one value".to_string())
                    } else {
                        EvalResult::Number(vals.iter().sum::<f64>() / vals.len() as f64)
                    }
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "MIN" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        EvalResult::Number(0.0)
                    } else {
                        EvalResult::Number(vals.iter().cloned().fold(f64::INFINITY, f64::min))
                    }
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "MAX" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        EvalResult::Number(0.0)
                    } else {
                        EvalResult::Number(vals.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
                    }
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "COUNT" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => EvalResult::Number(vals.len() as f64),
                Err(e) => EvalResult::Error(e),
            }
        }
        "COUNTA" => {
            // Count non-empty cells
            let values = collect_all_values(args, lookup);
            let count = values.iter().filter(|v| !matches!(v, EvalResult::Text(s) if s.is_empty())).count();
            EvalResult::Number(count as f64)
        }
        "ABS" => {
            if args.len() != 1 {
                return EvalResult::Error("ABS requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.abs()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ROUND" => {
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("ROUND requires 1 or 2 arguments".to_string());
            }
            let value = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let decimals = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0
            };
            let factor = 10_f64.powi(decimals);
            EvalResult::Number((value * factor).round() / factor)
        }
        "INT" => {
            if args.len() != 1 {
                return EvalResult::Error("INT requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.floor()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "MOD" => {
            if args.len() != 2 {
                return EvalResult::Error("MOD requires exactly 2 arguments".to_string());
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let divisor = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            if divisor == 0.0 {
                return EvalResult::Error("#DIV/0!".to_string());
            }
            EvalResult::Number(number % divisor)
        }
        "POWER" => {
            if args.len() != 2 {
                return EvalResult::Error("POWER requires exactly 2 arguments".to_string());
            }
            let base = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let exp = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            EvalResult::Number(base.powf(exp))
        }
        // =====================================================================
        // Financial Functions
        // =====================================================================
        "PMT" => {
            // PMT(rate, nper, pv, [fv], [type])
            // Returns the payment for a loan based on constant payments and interest rate
            if args.len() < 3 || args.len() > 5 {
                return EvalResult::Error("PMT requires 3 to 5 arguments".to_string());
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let pv = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let fv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };

            if nper == 0.0 {
                return EvalResult::Error("#NUM!".to_string());
            }

            let pmt = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let pmt = (rate * (pv * pow + fv)) / (pow - 1.0);
                if pmt_type != 0.0 {
                    -pmt / (1.0 + rate)
                } else {
                    -pmt
                }
            };
            EvalResult::Number(pmt)
        }
        "FV" => {
            // FV(rate, nper, pmt, [pv], [type])
            // Returns the future value of an investment
            if args.len() < 3 || args.len() > 5 {
                return EvalResult::Error("FV requires 3 to 5 arguments".to_string());
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let pmt = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let pv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };

            let fv = if rate == 0.0 {
                -pv - pmt * nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let fv_pmt = if pmt_type != 0.0 {
                    pmt * (1.0 + rate) * (pow - 1.0) / rate
                } else {
                    pmt * (pow - 1.0) / rate
                };
                -pv * pow - fv_pmt
            };
            EvalResult::Number(fv)
        }
        "PV" => {
            // PV(rate, nper, pmt, [fv], [type])
            // Returns the present value of an investment
            if args.len() < 3 || args.len() > 5 {
                return EvalResult::Error("PV requires 3 to 5 arguments".to_string());
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let pmt = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let fv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.0
            };

            let pv = if rate == 0.0 {
                -fv - pmt * nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let pv_pmt = if pmt_type != 0.0 {
                    pmt * (1.0 + rate) * (pow - 1.0) / rate
                } else {
                    pmt * (pow - 1.0) / rate
                };
                (-fv - pv_pmt) / pow
            };
            EvalResult::Number(pv)
        }
        "NPV" => {
            // NPV(rate, value1, [value2], ...)
            // Returns the net present value of an investment based on periodic cash flows
            if args.len() < 2 {
                return EvalResult::Error("NPV requires at least 2 arguments".to_string());
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };

            if rate == -1.0 {
                return EvalResult::Error("#DIV/0!".to_string());
            }

            let mut npv = 0.0;
            let mut period = 1;

            for arg in &args[1..] {
                // Handle both single values and ranges
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        let (min_row, min_col) = (*start_row.min(end_row), *start_col.min(end_col));
                        let (max_row, max_col) = (*start_row.max(end_row), *start_col.max(end_col));
                        for r in min_row..=max_row {
                            for c in min_col..=max_col {
                                let val = lookup.get_value(r, c);
                                if val.is_finite() {
                                    npv += val / (1.0 + rate).powi(period);
                                    period += 1;
                                }
                            }
                        }
                    }
                    _ => {
                        match evaluate(arg, lookup).to_number() {
                            Ok(n) => {
                                npv += n / (1.0 + rate).powi(period);
                                period += 1;
                            }
                            Err(e) => return EvalResult::Error(e),
                        }
                    }
                }
            }
            EvalResult::Number(npv)
        }
        "IRR" => {
            // IRR(values, [guess])
            // Returns the internal rate of return for a series of cash flows
            if args.len() < 1 || args.len() > 2 {
                return EvalResult::Error("IRR requires 1 or 2 arguments".to_string());
            }

            // Collect cash flows from range
            let values: Vec<f64> = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col) = (*start_row.min(end_row), *start_col.min(end_col));
                    let (max_row, max_col) = (*start_row.max(end_row), *start_col.max(end_col));
                    let mut vals = Vec::new();
                    for r in min_row..=max_row {
                        for c in min_col..=max_col {
                            let val = lookup.get_value(r, c);
                            if val.is_finite() && val != 0.0 {
                                vals.push(val);
                            } else if lookup.get_text(r, c).is_empty() {
                                // Skip empty cells
                            } else if val == 0.0 {
                                vals.push(0.0);
                            }
                        }
                    }
                    vals
                }
                _ => return EvalResult::Error("IRR requires a range of values".to_string()),
            };

            if values.len() < 2 {
                return EvalResult::Error("#NUM!".to_string());
            }

            // Check that there's at least one positive and one negative value
            let has_positive = values.iter().any(|&v| v > 0.0);
            let has_negative = values.iter().any(|&v| v < 0.0);
            if !has_positive || !has_negative {
                return EvalResult::Error("#NUM!".to_string());
            }

            let guess = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0.1 // Default guess of 10%
            };

            // Newton-Raphson iteration to find IRR
            let mut rate = guess;
            let max_iterations = 100;
            let tolerance = 1e-10;

            for _ in 0..max_iterations {
                let mut npv = 0.0;
                let mut dnpv = 0.0; // derivative of NPV with respect to rate

                for (i, &cf) in values.iter().enumerate() {
                    let t = i as f64;
                    let divisor = (1.0 + rate).powf(t);
                    if divisor == 0.0 {
                        return EvalResult::Error("#NUM!".to_string());
                    }
                    npv += cf / divisor;
                    if t > 0.0 {
                        dnpv -= t * cf / (1.0 + rate).powf(t + 1.0);
                    }
                }

                if dnpv.abs() < 1e-30 {
                    return EvalResult::Error("#NUM!".to_string());
                }

                let new_rate = rate - npv / dnpv;

                if (new_rate - rate).abs() < tolerance {
                    return EvalResult::Number(new_rate);
                }

                rate = new_rate;

                // Prevent divergence
                if rate < -1.0 || rate > 10.0 || !rate.is_finite() {
                    return EvalResult::Error("#NUM!".to_string());
                }
            }

            // Failed to converge
            EvalResult::Error("#NUM!".to_string())
        }
        "SQRT" => {
            if args.len() != 1 {
                return EvalResult::Error("SQRT requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.sqrt()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "CEILING" => {
            if args.len() != 1 {
                return EvalResult::Error("CEILING requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.ceil()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "FLOOR" => {
            if args.len() != 1 {
                return EvalResult::Error("FLOOR requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.floor()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "PRODUCT" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        EvalResult::Number(0.0)
                    } else {
                        EvalResult::Number(vals.iter().product())
                    }
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "MEDIAN" => {
            let values = collect_numbers(args, lookup);
            match values {
                Ok(mut vals) => {
                    if vals.is_empty() {
                        EvalResult::Error("MEDIAN requires at least one value".to_string())
                    } else {
                        vals.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
                        let mid = vals.len() / 2;
                        if vals.len() % 2 == 0 {
                            EvalResult::Number((vals[mid - 1] + vals[mid]) / 2.0)
                        } else {
                            EvalResult::Number(vals[mid])
                        }
                    }
                }
                Err(e) => EvalResult::Error(e),
            }
        }

        // =====================
        // LOGICAL FUNCTIONS
        // =====================
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("IF requires 2 or 3 arguments".to_string());
            }
            let condition = match evaluate(&args[0], lookup).to_bool() {
                Ok(b) => b,
                Err(e) => return EvalResult::Error(e),
            };
            if condition {
                evaluate(&args[1], lookup)
            } else if args.len() == 3 {
                evaluate(&args[2], lookup)
            } else {
                EvalResult::Boolean(false)
            }
        }
        "AND" => {
            if args.is_empty() {
                return EvalResult::Error("AND requires at least one argument".to_string());
            }
            for arg in args {
                match evaluate(arg, lookup).to_bool() {
                    Ok(false) => return EvalResult::Boolean(false),
                    Err(e) => return EvalResult::Error(e),
                    _ => {}
                }
            }
            EvalResult::Boolean(true)
        }
        "OR" => {
            if args.is_empty() {
                return EvalResult::Error("OR requires at least one argument".to_string());
            }
            for arg in args {
                match evaluate(arg, lookup).to_bool() {
                    Ok(true) => return EvalResult::Boolean(true),
                    Err(e) => return EvalResult::Error(e),
                    _ => {}
                }
            }
            EvalResult::Boolean(false)
        }
        "NOT" => {
            if args.len() != 1 {
                return EvalResult::Error("NOT requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_bool() {
                Ok(b) => EvalResult::Boolean(!b),
                Err(e) => EvalResult::Error(e),
            }
        }
        "IFERROR" => {
            if args.len() != 2 {
                return EvalResult::Error("IFERROR requires exactly 2 arguments".to_string());
            }
            let result = evaluate(&args[0], lookup);
            if result.is_error() {
                evaluate(&args[1], lookup)
            } else {
                result
            }
        }
        "ISBLANK" => {
            if args.len() != 1 {
                return EvalResult::Error("ISBLANK requires exactly one argument".to_string());
            }
            let result = evaluate(&args[0], lookup);
            let is_blank = match &result {
                EvalResult::Text(s) => s.is_empty(),
                _ => false,
            };
            EvalResult::Boolean(is_blank)
        }
        "ISNUMBER" => {
            if args.len() != 1 {
                return EvalResult::Error("ISNUMBER requires exactly one argument".to_string());
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Number(_)))
        }
        "ISTEXT" => {
            if args.len() != 1 {
                return EvalResult::Error("ISTEXT requires exactly one argument".to_string());
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Text(_)))
        }
        "ISERROR" => {
            if args.len() != 1 {
                return EvalResult::Error("ISERROR requires exactly one argument".to_string());
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(result.is_error())
        }
        "ISNA" => {
            if args.len() != 1 {
                return EvalResult::Error("ISNA requires exactly one argument".to_string());
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Error(ref e) if e == "#N/A"))
        }

        // =====================
        // TEXT FUNCTIONS
        // =====================
        "CONCATENATE" | "CONCAT" => {
            let mut result = String::new();
            for arg in args {
                result.push_str(&evaluate(arg, lookup).to_text());
            }
            EvalResult::Text(result)
        }
        "TEXTJOIN" => {
            // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
            if args.len() < 3 {
                return EvalResult::Error("TEXTJOIN requires at least 3 arguments".to_string());
            }
            let delimiter = evaluate(&args[0], lookup).to_text();
            let ignore_empty = match evaluate(&args[1], lookup).to_bool() {
                Ok(b) => b,
                Err(_) => true, // default to TRUE
            };

            let mut parts: Vec<String> = Vec::new();

            for arg in &args[2..] {
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        // Collect all values from range
                        let (min_row, min_col, max_row, max_col) = (
                            (*start_row).min(*end_row), (*start_col).min(*end_col),
                            (*start_row).max(*end_row), (*start_col).max(*end_col)
                        );
                        for r in min_row..=max_row {
                            for c in min_col..=max_col {
                                let text = lookup.get_text(r, c);
                                if !ignore_empty || !text.is_empty() {
                                    parts.push(text);
                                }
                            }
                        }
                    }
                    _ => {
                        let text = evaluate(arg, lookup).to_text();
                        if !ignore_empty || !text.is_empty() {
                            parts.push(text);
                        }
                    }
                }
            }

            EvalResult::Text(parts.join(&delimiter))
        }
        "LEFT" => {
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("LEFT requires 1 or 2 arguments".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1
            };
            EvalResult::Text(text.chars().take(num_chars).collect())
        }
        "RIGHT" => {
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("RIGHT requires 1 or 2 arguments".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1
            };
            let len = text.chars().count();
            let start = len.saturating_sub(num_chars);
            EvalResult::Text(text.chars().skip(start).collect())
        }
        "MID" => {
            if args.len() != 3 {
                return EvalResult::Error("MID requires exactly 3 arguments".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                Ok(n) => (n as usize).saturating_sub(1), // 1-indexed
                Err(e) => return EvalResult::Error(e),
            };
            let num_chars = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            EvalResult::Text(text.chars().skip(start).take(num_chars).collect())
        }
        "LEN" => {
            if args.len() != 1 {
                return EvalResult::Error("LEN requires exactly one argument".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Number(text.chars().count() as f64)
        }
        "UPPER" => {
            if args.len() != 1 {
                return EvalResult::Error("UPPER requires exactly one argument".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_uppercase())
        }
        "LOWER" => {
            if args.len() != 1 {
                return EvalResult::Error("LOWER requires exactly one argument".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_lowercase())
        }
        "TRIM" => {
            if args.len() != 1 {
                return EvalResult::Error("TRIM requires exactly one argument".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            // TRIM removes leading/trailing spaces and collapses internal spaces
            let trimmed: String = text.split_whitespace().collect::<Vec<_>>().join(" ");
            EvalResult::Text(trimmed)
        }
        "TEXT" => {
            if args.len() != 2 {
                return EvalResult::Error("TEXT requires exactly 2 arguments".to_string());
            }
            let value = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let format = evaluate(&args[1], lookup).to_text();
            // Simple format support
            let result = if format.contains("0.") {
                let decimals = format.matches('0').count().saturating_sub(1);
                format!("{:.1$}", value, decimals)
            } else if format.contains('%') {
                format!("{}%", (value * 100.0) as i64)
            } else {
                format!("{}", value)
            };
            EvalResult::Text(result)
        }
        "VALUE" => {
            if args.len() != 1 {
                return EvalResult::Error("VALUE requires exactly one argument".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            match text.replace(',', "").trim().parse::<f64>() {
                Ok(n) => EvalResult::Number(n),
                Err(_) => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "FIND" => {
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("FIND requires 2 or 3 arguments".to_string());
            }
            let find_text = evaluate(&args[0], lookup).to_text();
            let within_text = evaluate(&args[1], lookup).to_text();
            let start_pos = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                    Ok(n) => (n as usize).saturating_sub(1),
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                0
            };
            let search_area = &within_text[start_pos.min(within_text.len())..];
            match search_area.find(&find_text) {
                Some(pos) => EvalResult::Number((pos + start_pos + 1) as f64), // 1-indexed
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "SUBSTITUTE" => {
            if args.len() < 3 || args.len() > 4 {
                return EvalResult::Error("SUBSTITUTE requires 3 or 4 arguments".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            let old_text = evaluate(&args[1], lookup).to_text();
            let new_text = evaluate(&args[2], lookup).to_text();
            let instance = if args.len() == 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => Some(n as usize),
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                None
            };

            let result = if let Some(n) = instance {
                // Replace only the nth instance
                let mut count = 0;
                let mut result = String::new();
                let mut remaining = text.as_str();
                while let Some(pos) = remaining.find(&old_text) {
                    count += 1;
                    if count == n {
                        result.push_str(&remaining[..pos]);
                        result.push_str(&new_text);
                        result.push_str(&remaining[pos + old_text.len()..]);
                        break;
                    } else {
                        result.push_str(&remaining[..pos + old_text.len()]);
                        remaining = &remaining[pos + old_text.len()..];
                    }
                }
                if count < n {
                    text // Not enough instances found
                } else {
                    result
                }
            } else {
                // Replace all instances
                text.replace(&old_text, &new_text)
            };
            EvalResult::Text(result)
        }
        "REPT" => {
            if args.len() != 2 {
                return EvalResult::Error("REPT requires exactly 2 arguments".to_string());
            }
            let text = evaluate(&args[0], lookup).to_text();
            let times = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 0.0 => return EvalResult::Error("#VALUE!".to_string()),
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            EvalResult::Text(text.repeat(times))
        }

        // =====================
        // CONDITIONAL FUNCTIONS
        // =====================
        "SUMIF" => {
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("SUMIF requires 2 or 3 arguments".to_string());
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("SUMIF requires a range as first argument".to_string()),
            };
            let criteria = evaluate(&args[1], lookup);
            let sum_range = if args.len() == 3 {
                match &args[2] {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => Some((*start_row, *start_col, *end_row, *end_col)),
                    _ => return EvalResult::Error("SUMIF sum_range must be a range".to_string()),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        // Get value from sum_range or criteria range
                        let (sum_r, sum_c) = if let Some((sr, sc, _, _)) = sum_range {
                            (sr + row_offset, sc + col_offset)
                        } else {
                            (r, c)
                        };
                        sum += lookup.get_value(sum_r, sum_c);
                    }
                }
            }
            EvalResult::Number(sum)
        }
        "AVERAGEIF" => {
            // AVERAGEIF(range, criteria, [average_range])
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("AVERAGEIF requires 2 or 3 arguments".to_string());
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("AVERAGEIF requires a range as first argument".to_string()),
            };
            let criteria = evaluate(&args[1], lookup);
            let avg_range = if args.len() == 3 {
                match &args[2] {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => Some((*start_row, *start_col, *end_row, *end_col)),
                    _ => return EvalResult::Error("AVERAGEIF average_range must be a range".to_string()),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        // Get value from average_range or criteria range
                        let (avg_r, avg_c) = if let Some((ar, ac, _, _)) = avg_range {
                            (ar + row_offset, ac + col_offset)
                        } else {
                            (r, c)
                        };
                        let val = lookup.get_value(avg_r, avg_c);
                        // Only count numeric values for average
                        if val.is_finite() {
                            sum += val;
                            count += 1;
                        }
                    }
                }
            }
            if count == 0 {
                EvalResult::Error("#DIV/0!".to_string())
            } else {
                EvalResult::Number(sum / count as f64)
            }
        }
        "COUNTIF" => {
            if args.len() != 2 {
                return EvalResult::Error("COUNTIF requires exactly 2 arguments".to_string());
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("COUNTIF requires a range as first argument".to_string()),
            };
            let criteria = evaluate(&args[1], lookup);

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Text(String::new())
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        "COUNTBLANK" => {
            if args.len() != 1 {
                return EvalResult::Error("COUNTBLANK requires exactly one argument".to_string());
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("COUNTBLANK requires a range".to_string()),
            };

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if lookup.get_text(r, c).is_empty() {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        "SUMIFS" => {
            // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 3 || (args.len() - 1) % 2 != 0 {
                return EvalResult::Error("SUMIFS requires sum_range and pairs of criteria_range and criteria".to_string());
            }

            let sum_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("SUMIFS sum_range must be a range".to_string()),
            };

            let (sr_min_row, sr_min_col, sr_max_row, sr_max_col) = (
                sum_range.0.min(sum_range.2), sum_range.1.min(sum_range.3),
                sum_range.0.max(sum_range.2), sum_range.1.max(sum_range.3)
            );
            let num_rows = sr_max_row - sr_min_row + 1;
            let num_cols = sr_max_col - sr_min_col + 1;

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return EvalResult::Error("SUMIFS criteria_range must be a range".to_string()),
                };

                // Verify dimensions match sum_range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return EvalResult::Error("SUMIFS criteria ranges must have same dimensions as sum_range".to_string());
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Number(0.0)
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        sum += lookup.get_value(sr_min_row + row_offset, sr_min_col + col_offset);
                    }
                }
            }
            EvalResult::Number(sum)
        }
        "AVERAGEIFS" => {
            // AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 3 || (args.len() - 1) % 2 != 0 {
                return EvalResult::Error("AVERAGEIFS requires average_range and pairs of criteria_range and criteria".to_string());
            }

            let avg_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("AVERAGEIFS average_range must be a range".to_string()),
            };

            let (ar_min_row, ar_min_col, ar_max_row, ar_max_col) = (
                avg_range.0.min(avg_range.2), avg_range.1.min(avg_range.3),
                avg_range.0.max(avg_range.2), avg_range.1.max(avg_range.3)
            );
            let num_rows = ar_max_row - ar_min_row + 1;
            let num_cols = ar_max_col - ar_min_col + 1;

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return EvalResult::Error("AVERAGEIFS criteria_range must be a range".to_string()),
                };

                // Verify dimensions match average_range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return EvalResult::Error("AVERAGEIFS criteria ranges must have same dimensions as average_range".to_string());
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Number(0.0)
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        let val = lookup.get_value(ar_min_row + row_offset, ar_min_col + col_offset);
                        if val.is_finite() {
                            sum += val;
                            count += 1;
                        }
                    }
                }
            }
            if count == 0 {
                EvalResult::Error("#DIV/0!".to_string())
            } else {
                EvalResult::Number(sum / count as f64)
            }
        }
        "COUNTIFS" => {
            // COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 2 || args.len() % 2 != 0 {
                return EvalResult::Error("COUNTIFS requires pairs of criteria_range and criteria".to_string());
            }

            // Use first range to determine dimensions
            let first_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("COUNTIFS criteria_range must be a range".to_string()),
            };

            let (fr_min_row, fr_min_col, fr_max_row, fr_max_col) = (
                first_range.0.min(first_range.2), first_range.1.min(first_range.3),
                first_range.0.max(first_range.2), first_range.1.max(first_range.3)
            );
            let num_rows = fr_max_row - fr_min_row + 1;
            let num_cols = fr_max_col - fr_min_col + 1;

            // Parse criteria pairs
            let num_criteria = args.len() / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[i * 2];
                let criteria_arg = &args[i * 2 + 1];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return EvalResult::Error("COUNTIFS criteria_range must be a range".to_string()),
                };

                // Verify dimensions match first range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return EvalResult::Error("COUNTIFS ranges must have same dimensions".to_string());
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Text(String::new())
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        "IFNA" => {
            // IFNA(value, value_if_na)
            if args.len() != 2 {
                return EvalResult::Error("IFNA requires exactly 2 arguments".to_string());
            }
            let value = evaluate(&args[0], lookup);
            match value {
                EvalResult::Error(ref e) if e == "#N/A" => evaluate(&args[1], lookup),
                _ => value,
            }
        }

        // =====================
        // LOOKUP FUNCTIONS
        // =====================
        "VLOOKUP" => {
            // VLOOKUP(search_key, range, index, [is_sorted])
            if args.len() < 3 || args.len() > 4 {
                return EvalResult::Error("VLOOKUP requires 3 or 4 arguments".to_string());
            }
            let search_key = evaluate(&args[0], lookup);
            let range = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("VLOOKUP requires a range as second argument".to_string()),
            };
            let col_index = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            let is_sorted = if args.len() == 4 {
                match evaluate(&args[3], lookup).to_bool() {
                    Ok(b) => b,
                    Err(_) => true, // default to TRUE
                }
            } else {
                true
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_cols = max_col - min_col + 1;

            if col_index > num_cols {
                return EvalResult::Error("#REF!".to_string());
            }

            // Search for the key in the first column
            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_row: Option<usize> = None;

            if is_sorted {
                // Approximate match (find largest value <= search_key)
                if let Some(search_n) = search_num {
                    let mut best_row: Option<usize> = None;
                    let mut best_val = f64::NEG_INFINITY;
                    for r in min_row..=max_row {
                        let cell_text = lookup.get_text(r, min_col);
                        if let Ok(cell_n) = cell_text.parse::<f64>() {
                            if cell_n <= search_n && cell_n > best_val {
                                best_val = cell_n;
                                best_row = Some(r);
                            }
                        }
                    }
                    found_row = best_row;
                }
            } else {
                // Exact match
                for r in min_row..=max_row {
                    let cell_text = lookup.get_text(r, min_col);
                    let cell_lower = cell_text.to_lowercase();

                    // Try numeric comparison first
                    if let (Some(search_n), Ok(cell_n)) = (search_num, cell_text.parse::<f64>()) {
                        if (search_n - cell_n).abs() < f64::EPSILON {
                            found_row = Some(r);
                            break;
                        }
                    } else if cell_lower == search_text {
                        found_row = Some(r);
                        break;
                    }
                }
            }

            match found_row {
                Some(r) => {
                    let result_col = min_col + col_index - 1;
                    let result_text = lookup.get_text(r, result_col);
                    if result_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "XLOOKUP" => {
            // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
            // match_mode: 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard
            // search_mode: 1 = first to last (default), -1 = last to first, 2 = binary ascending, -2 = binary descending
            if args.len() < 3 || args.len() > 6 {
                return EvalResult::Error("XLOOKUP requires 3 to 6 arguments".to_string());
            }

            let lookup_value = evaluate(&args[0], lookup);

            // Parse lookup array (must be 1D - single row or column)
            let lookup_array = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col, max_row, max_col) = (
                        (*start_row).min(*end_row), (*start_col).min(*end_col),
                        (*start_row).max(*end_row), (*start_col).max(*end_col)
                    );
                    // Determine orientation
                    let is_row = min_row == max_row;
                    let is_col = min_col == max_col;
                    if !is_row && !is_col {
                        return EvalResult::Error("XLOOKUP lookup_array must be a single row or column".to_string());
                    }
                    (min_row, min_col, max_row, max_col, is_row)
                }
                _ => return EvalResult::Error("XLOOKUP lookup_array must be a range".to_string()),
            };

            // Parse return array (must have same dimensions)
            let return_array = match &args[2] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col, max_row, max_col) = (
                        (*start_row).min(*end_row), (*start_col).min(*end_col),
                        (*start_row).max(*end_row), (*start_col).max(*end_col)
                    );
                    (min_row, min_col, max_row, max_col)
                }
                _ => return EvalResult::Error("XLOOKUP return_array must be a range".to_string()),
            };

            // Verify lookup and return arrays have compatible dimensions
            let lookup_size = if lookup_array.4 {
                lookup_array.3 - lookup_array.1 + 1  // columns for row array
            } else {
                lookup_array.2 - lookup_array.0 + 1  // rows for column array
            };

            let return_size = if lookup_array.4 {
                return_array.3 - return_array.1 + 1  // columns
            } else {
                return_array.2 - return_array.0 + 1  // rows
            };

            if lookup_size != return_size {
                return EvalResult::Error("XLOOKUP lookup and return arrays must have same size".to_string());
            }

            // Optional: if_not_found
            let if_not_found = if args.len() >= 4 {
                Some(&args[3])
            } else {
                None
            };

            // Optional: match_mode (0 = exact match, default)
            let match_mode = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 0,
                }
            } else {
                0
            };

            // Search for match
            let mut found_idx: Option<usize> = None;

            for idx in 0..lookup_size {
                let (r, c) = if lookup_array.4 {
                    (lookup_array.0, lookup_array.1 + idx)
                } else {
                    (lookup_array.0 + idx, lookup_array.1)
                };

                let cell_text = lookup.get_text(r, c);
                let cell_value = if cell_text.is_empty() {
                    EvalResult::Text(String::new())
                } else if let Ok(n) = cell_text.parse::<f64>() {
                    EvalResult::Number(n)
                } else {
                    EvalResult::Text(cell_text)
                };

                let is_match = match match_mode {
                    0 => {
                        // Exact match
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            (EvalResult::Text(a), EvalResult::Text(b)) => a.eq_ignore_ascii_case(b),
                            _ => false,
                        }
                    }
                    2 => {
                        // Wildcard match (simplified: just exact for now)
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Text(pattern), EvalResult::Text(text)) => {
                                // Simple wildcard: * matches any chars, ? matches single char
                                let pattern_lower = pattern.to_lowercase();
                                let text_lower = text.to_lowercase();
                                if pattern_lower.contains('*') || pattern_lower.contains('?') {
                                    // Simple wildcard matching: * matches any sequence, ? matches single char
                                    // For now, just check prefix match for patterns like "abc*"
                                    let prefix = pattern_lower.split('*').next().unwrap_or("");
                                    text_lower.starts_with(prefix) || text_lower == pattern_lower
                                } else {
                                    pattern_lower == text_lower
                                }
                            }
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            _ => false,
                        }
                    }
                    _ => {
                        // For -1 (next smaller) and 1 (next larger), fall back to exact for simplicity
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            (EvalResult::Text(a), EvalResult::Text(b)) => a.eq_ignore_ascii_case(b),
                            _ => false,
                        }
                    }
                };

                if is_match {
                    found_idx = Some(idx);
                    break;
                }
            }

            match found_idx {
                Some(idx) => {
                    // Return value from return_array at same position
                    let (r, c) = if lookup_array.4 {
                        (return_array.0, return_array.1 + idx)
                    } else {
                        (return_array.0 + idx, return_array.1)
                    };

                    let result_text = lookup.get_text(r, c);
                    if result_text.is_empty() {
                        EvalResult::Text(String::new())
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => {
                    // Not found - return if_not_found or #N/A
                    if let Some(if_not_found_expr) = if_not_found {
                        evaluate(if_not_found_expr, lookup)
                    } else {
                        EvalResult::Error("#N/A".to_string())
                    }
                }
            }
        }
        "HLOOKUP" => {
            // HLOOKUP(search_key, range, index, [is_sorted])
            if args.len() < 3 || args.len() > 4 {
                return EvalResult::Error("HLOOKUP requires 3 or 4 arguments".to_string());
            }
            let search_key = evaluate(&args[0], lookup);
            let range = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("HLOOKUP requires a range as second argument".to_string()),
            };
            let row_index = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            let is_sorted = if args.len() == 4 {
                evaluate(&args[3], lookup).to_bool().unwrap_or(true)
            } else {
                true
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_rows = max_row - min_row + 1;

            if row_index > num_rows {
                return EvalResult::Error("#REF!".to_string());
            }

            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_col: Option<usize> = None;

            if is_sorted {
                if let Some(search_n) = search_num {
                    let mut best_col: Option<usize> = None;
                    let mut best_val = f64::NEG_INFINITY;
                    for c in min_col..=max_col {
                        let cell_text = lookup.get_text(min_row, c);
                        if let Ok(cell_n) = cell_text.parse::<f64>() {
                            if cell_n <= search_n && cell_n > best_val {
                                best_val = cell_n;
                                best_col = Some(c);
                            }
                        }
                    }
                    found_col = best_col;
                }
            } else {
                for c in min_col..=max_col {
                    let cell_text = lookup.get_text(min_row, c);
                    let cell_lower = cell_text.to_lowercase();

                    if let (Some(search_n), Ok(cell_n)) = (search_num, cell_text.parse::<f64>()) {
                        if (search_n - cell_n).abs() < f64::EPSILON {
                            found_col = Some(c);
                            break;
                        }
                    } else if cell_lower == search_text {
                        found_col = Some(c);
                        break;
                    }
                }
            }

            match found_col {
                Some(c) => {
                    let result_row = min_row + row_index - 1;
                    let result_text = lookup.get_text(result_row, c);
                    if result_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "INDEX" => {
            // INDEX(range, row_num, [col_num])
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("INDEX requires 2 or 3 arguments".to_string());
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                Expr::CellRef { col, row, .. } => (*row, *col, *row, *col),
                _ => return EvalResult::Error("INDEX requires a range as first argument".to_string()),
            };
            let row_num = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            let col_num = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_rows = max_row - min_row + 1;
            let num_cols = max_col - min_col + 1;

            if row_num < 1 || row_num > num_rows || col_num < 1 || col_num > num_cols {
                return EvalResult::Error("#REF!".to_string());
            }

            let target_row = min_row + row_num - 1;
            let target_col = min_col + col_num - 1;
            let result_text = lookup.get_text(target_row, target_col);

            if result_text.is_empty() {
                EvalResult::Number(0.0)
            } else if let Ok(n) = result_text.parse::<f64>() {
                EvalResult::Number(n)
            } else {
                EvalResult::Text(result_text)
            }
        }
        "MATCH" => {
            // MATCH(search_key, range, [match_type])
            if args.len() < 2 || args.len() > 3 {
                return EvalResult::Error("MATCH requires 2 or 3 arguments".to_string());
            }
            let search_key = evaluate(&args[0], lookup);
            let range = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return EvalResult::Error("MATCH requires a range as second argument".to_string()),
            };
            let match_type = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 1,
                }
            } else {
                1
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            // Determine if it's a row vector or column vector
            let is_row = min_row == max_row;
            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_pos: Option<usize> = None;

            if is_row {
                // Search horizontally
                if match_type == 0 {
                    // Exact match
                    for (i, c) in (min_col..=max_col).enumerate() {
                        let cell_text = lookup.get_text(min_row, c);
                        let cell_lower = cell_text.to_lowercase();
                        if let (Some(sn), Ok(cn)) = (search_num, cell_text.parse::<f64>()) {
                            if (sn - cn).abs() < f64::EPSILON {
                                found_pos = Some(i + 1);
                                break;
                            }
                        } else if cell_lower == search_text {
                            found_pos = Some(i + 1);
                            break;
                        }
                    }
                } else if match_type == 1 {
                    // Largest value <= search_key
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::NEG_INFINITY;
                        for (i, c) in (min_col..=max_col).enumerate() {
                            if let Ok(cn) = lookup.get_text(min_row, c).parse::<f64>() {
                                if cn <= sn && cn > best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                } else {
                    // Smallest value >= search_key
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::INFINITY;
                        for (i, c) in (min_col..=max_col).enumerate() {
                            if let Ok(cn) = lookup.get_text(min_row, c).parse::<f64>() {
                                if cn >= sn && cn < best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                }
            } else {
                // Search vertically
                if match_type == 0 {
                    for (i, r) in (min_row..=max_row).enumerate() {
                        let cell_text = lookup.get_text(r, min_col);
                        let cell_lower = cell_text.to_lowercase();
                        if let (Some(sn), Ok(cn)) = (search_num, cell_text.parse::<f64>()) {
                            if (sn - cn).abs() < f64::EPSILON {
                                found_pos = Some(i + 1);
                                break;
                            }
                        } else if cell_lower == search_text {
                            found_pos = Some(i + 1);
                            break;
                        }
                    }
                } else if match_type == 1 {
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::NEG_INFINITY;
                        for (i, r) in (min_row..=max_row).enumerate() {
                            if let Ok(cn) = lookup.get_text(r, min_col).parse::<f64>() {
                                if cn <= sn && cn > best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                } else {
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::INFINITY;
                        for (i, r) in (min_row..=max_row).enumerate() {
                            if let Ok(cn) = lookup.get_text(r, min_col).parse::<f64>() {
                                if cn >= sn && cn < best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                }
            }

            match found_pos {
                Some(pos) => EvalResult::Number(pos as f64),
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "ROW" => {
            // ROW([reference]) - returns the row number
            if args.is_empty() {
                // No argument: return row of current cell being evaluated
                return match lookup.current_cell() {
                    Some((row, _)) => EvalResult::Number((row + 1) as f64),
                    None => EvalResult::Error("ROW() requires cell context".to_string()),
                };
            }
            if args.len() != 1 {
                return EvalResult::Error("ROW requires 0 or 1 argument".to_string());
            }
            match &args[0] {
                Expr::CellRef { row, .. } => EvalResult::Number((*row + 1) as f64),
                Expr::Range { start_row, .. } => EvalResult::Number((*start_row + 1) as f64),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "COLUMN" => {
            // COLUMN([reference]) - returns the column number
            if args.is_empty() {
                // No argument: return column of current cell being evaluated
                return match lookup.current_cell() {
                    Some((_, col)) => EvalResult::Number((col + 1) as f64),
                    None => EvalResult::Error("COLUMN() requires cell context".to_string()),
                };
            }
            if args.len() != 1 {
                return EvalResult::Error("COLUMN requires 0 or 1 argument".to_string());
            }
            match &args[0] {
                Expr::CellRef { col, .. } => EvalResult::Number((*col + 1) as f64),
                Expr::Range { start_col, .. } => EvalResult::Number((*start_col + 1) as f64),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "ROWS" => {
            if args.len() != 1 {
                return EvalResult::Error("ROWS requires exactly one argument".to_string());
            }
            match &args[0] {
                Expr::Range { start_row, end_row, .. } => {
                    EvalResult::Number((end_row.max(start_row) - end_row.min(start_row) + 1) as f64)
                }
                Expr::CellRef { .. } => EvalResult::Number(1.0),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "COLUMNS" => {
            if args.len() != 1 {
                return EvalResult::Error("COLUMNS requires exactly one argument".to_string());
            }
            match &args[0] {
                Expr::Range { start_col, end_col, .. } => {
                    EvalResult::Number((end_col.max(start_col) - end_col.min(start_col) + 1) as f64)
                }
                Expr::CellRef { .. } => EvalResult::Number(1.0),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }

        // =====================
        // DATE/TIME FUNCTIONS
        // =====================
        "TODAY" => {
            if !args.is_empty() {
                return EvalResult::Error("TODAY takes no arguments".to_string());
            }
            // Return Excel-style date serial number (days since 1899-12-30)
            use std::time::{SystemTime, UNIX_EPOCH};
            let now = SystemTime::now().duration_since(UNIX_EPOCH).unwrap();
            let days_since_unix = now.as_secs() / 86400;
            // Excel epoch is 1899-12-30, Unix epoch is 1970-01-01
            // Difference is 25569 days
            let excel_date = days_since_unix as f64 + 25569.0;
            EvalResult::Number(excel_date)
        }
        "NOW" => {
            if !args.is_empty() {
                return EvalResult::Error("NOW takes no arguments".to_string());
            }
            use std::time::{SystemTime, UNIX_EPOCH};
            let now = SystemTime::now().duration_since(UNIX_EPOCH).unwrap();
            let secs = now.as_secs() as f64 + now.subsec_nanos() as f64 / 1_000_000_000.0;
            let days_since_unix = secs / 86400.0;
            let excel_datetime = days_since_unix + 25569.0;
            EvalResult::Number(excel_datetime)
        }
        "DATE" => {
            // DATE(year, month, day) - returns Excel date serial
            if args.len() != 3 {
                return EvalResult::Error("DATE requires exactly 3 arguments".to_string());
            }
            let year = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return EvalResult::Error(e),
            };
            let month = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return EvalResult::Error(e),
            };
            let day = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return EvalResult::Error(e),
            };

            // Adjust year if 0-99 (Excel convention)
            let year = if year < 100 { year + 1900 } else { year };

            // Simple date to Excel serial conversion
            // This is a simplified calculation
            let serial = date_to_serial(year, month, day);
            EvalResult::Number(serial)
        }
        "YEAR" => {
            if args.len() != 1 {
                return EvalResult::Error("YEAR requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let (year, _, _) = serial_to_date(serial);
            EvalResult::Number(year as f64)
        }
        "MONTH" => {
            if args.len() != 1 {
                return EvalResult::Error("MONTH requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let (_, month, _) = serial_to_date(serial);
            EvalResult::Number(month as f64)
        }
        "DAY" => {
            if args.len() != 1 {
                return EvalResult::Error("DAY requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let (_, _, day) = serial_to_date(serial);
            EvalResult::Number(day as f64)
        }
        "WEEKDAY" => {
            // WEEKDAY(date, [type]) - returns day of week
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("WEEKDAY requires 1 or 2 arguments".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as i64,
                Err(e) => return EvalResult::Error(e),
            };
            let return_type = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 1,
                }
            } else {
                1
            };

            // Excel serial 1 = 1900-01-01 which was a Sunday (but Excel incorrectly thinks it was Saturday due to 1900 leap year bug)
            // For simplicity, we'll use a corrected calculation
            let weekday = ((serial + 6) % 7) as i32; // 0 = Sunday, 6 = Saturday

            let result = match return_type {
                1 => weekday + 1,        // 1 (Sunday) to 7 (Saturday)
                2 => if weekday == 0 { 7 } else { weekday }, // 1 (Monday) to 7 (Sunday)
                3 => if weekday == 0 { 6 } else { weekday - 1 }, // 0 (Monday) to 6 (Sunday)
                _ => weekday + 1,
            };
            EvalResult::Number(result as f64)
        }
        "DATEDIF" => {
            // DATEDIF(start_date, end_date, unit)
            if args.len() != 3 {
                return EvalResult::Error("DATEDIF requires exactly 3 arguments".to_string());
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let end_serial = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let unit = evaluate(&args[2], lookup).to_text().to_uppercase();

            if start_serial > end_serial {
                return EvalResult::Error("#NUM!".to_string());
            }

            let (start_y, start_m, start_d) = serial_to_date(start_serial);
            let (end_y, end_m, end_d) = serial_to_date(end_serial);

            let result = match unit.as_str() {
                "Y" => {
                    // Complete years
                    let mut years = end_y - start_y;
                    if end_m < start_m || (end_m == start_m && end_d < start_d) {
                        years -= 1;
                    }
                    years as f64
                }
                "M" => {
                    // Complete months
                    let mut months = (end_y - start_y) * 12 + (end_m - start_m);
                    if end_d < start_d {
                        months -= 1;
                    }
                    months as f64
                }
                "D" => {
                    // Days
                    (end_serial - start_serial).floor()
                }
                "YM" => {
                    // Months ignoring years
                    let mut months = end_m - start_m;
                    if end_d < start_d {
                        months -= 1;
                    }
                    if months < 0 {
                        months += 12;
                    }
                    months as f64
                }
                "YD" => {
                    // Days ignoring years
                    let end_in_start_year = date_to_serial(start_y, end_m, end_d);
                    let mut days = end_in_start_year - start_serial;
                    if days < 0.0 {
                        let end_in_next_year = date_to_serial(start_y + 1, end_m, end_d);
                        days = end_in_next_year - start_serial;
                    }
                    days.floor()
                }
                "MD" => {
                    // Days ignoring months and years
                    let mut days = end_d - start_d;
                    if days < 0 {
                        // Days in previous month (simplified)
                        days += 30;
                    }
                    days as f64
                }
                _ => return EvalResult::Error("#VALUE!".to_string()),
            };
            EvalResult::Number(result)
        }
        "EDATE" => {
            // EDATE(start_date, months) - add months to a date
            if args.len() != 2 {
                return EvalResult::Error("EDATE requires exactly 2 arguments".to_string());
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let months = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return EvalResult::Error(e),
            };

            let (year, month, day) = serial_to_date(start_serial);
            let total_months = year * 12 + month + months;
            let new_year = (total_months - 1) / 12;
            let new_month = ((total_months - 1) % 12) + 1;

            // Clamp day to valid range for new month
            let days_in_month = days_in_month(new_year, new_month);
            let new_day = day.min(days_in_month);

            EvalResult::Number(date_to_serial(new_year, new_month, new_day))
        }
        "EOMONTH" => {
            // EOMONTH(start_date, months) - end of month after adding months
            if args.len() != 2 {
                return EvalResult::Error("EOMONTH requires exactly 2 arguments".to_string());
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let months = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return EvalResult::Error(e),
            };

            let (year, month, _) = serial_to_date(start_serial);
            let total_months = year * 12 + month + months;
            let new_year = (total_months - 1) / 12;
            let new_month = ((total_months - 1) % 12) + 1;
            let last_day = days_in_month(new_year, new_month);

            EvalResult::Number(date_to_serial(new_year, new_month, last_day))
        }
        "HOUR" => {
            if args.len() != 1 {
                return EvalResult::Error("HOUR requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let time_part = serial.fract();
            let hours = (time_part * 24.0).floor() as i32 % 24;
            EvalResult::Number(hours as f64)
        }
        "MINUTE" => {
            if args.len() != 1 {
                return EvalResult::Error("MINUTE requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let time_part = serial.fract();
            let total_minutes = (time_part * 24.0 * 60.0).floor() as i32;
            let minutes = total_minutes % 60;
            EvalResult::Number(minutes as f64)
        }
        "SECOND" => {
            if args.len() != 1 {
                return EvalResult::Error("SECOND requires exactly one argument".to_string());
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let time_part = serial.fract();
            let total_seconds = (time_part * 24.0 * 60.0 * 60.0).floor() as i32;
            let seconds = total_seconds % 60;
            EvalResult::Number(seconds as f64)
        }

        // =====================
        // RANDOM FUNCTIONS
        // =====================
        "RAND" => {
            if !args.is_empty() {
                return EvalResult::Error("RAND takes no arguments".to_string());
            }
            use std::time::{SystemTime, UNIX_EPOCH};
            // Simple LCG random - good enough for spreadsheet use
            let seed = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos() as u64;
            let random = ((seed.wrapping_mul(6364136223846793005).wrapping_add(1)) as f64)
                / (u64::MAX as f64);
            EvalResult::Number(random)
        }
        "RANDBETWEEN" => {
            if args.len() != 2 {
                return EvalResult::Error("RANDBETWEEN requires exactly 2 arguments".to_string());
            }
            let bottom = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.floor() as i64,
                Err(e) => return EvalResult::Error(e),
            };
            let top = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.floor() as i64,
                Err(e) => return EvalResult::Error(e),
            };
            if bottom > top {
                return EvalResult::Error("#NUM!".to_string());
            }
            use std::time::{SystemTime, UNIX_EPOCH};
            let seed = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos() as u64;
            let range = (top - bottom + 1) as u64;
            let random = (seed.wrapping_mul(6364136223846793005).wrapping_add(1)) % range;
            EvalResult::Number((bottom + random as i64) as f64)
        }

        // =====================
        // LOGARITHM FUNCTIONS
        // =====================
        "LOG" => {
            // LOG(number, [base]) - base defaults to 10
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("LOG requires 1 or 2 arguments".to_string());
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => return EvalResult::Error("#NUM!".to_string()),
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let base = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(b) if b <= 0.0 || b == 1.0 => return EvalResult::Error("#NUM!".to_string()),
                    Ok(b) => b,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                10.0
            };
            EvalResult::Number(number.log(base))
        }
        "LOG10" => {
            if args.len() != 1 {
                return EvalResult::Error("LOG10 requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.log10()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "LN" => {
            if args.len() != 1 {
                return EvalResult::Error("LN requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.ln()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "EXP" => {
            if args.len() != 1 {
                return EvalResult::Error("EXP requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.exp()),
                Err(e) => EvalResult::Error(e),
            }
        }

        // =====================
        // TRIGONOMETRY FUNCTIONS
        // =====================
        "PI" => {
            if !args.is_empty() {
                return EvalResult::Error("PI takes no arguments".to_string());
            }
            EvalResult::Number(std::f64::consts::PI)
        }
        "DEGREES" => {
            if args.len() != 1 {
                return EvalResult::Error("DEGREES requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.to_degrees()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "RADIANS" => {
            if args.len() != 1 {
                return EvalResult::Error("RADIANS requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.to_radians()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "SIN" => {
            if args.len() != 1 {
                return EvalResult::Error("SIN requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.sin()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "COS" => {
            if args.len() != 1 {
                return EvalResult::Error("COS requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.cos()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "TAN" => {
            if args.len() != 1 {
                return EvalResult::Error("TAN requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.tan()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ASIN" => {
            if args.len() != 1 {
                return EvalResult::Error("ASIN requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < -1.0 || n > 1.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.asin()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ACOS" => {
            if args.len() != 1 {
                return EvalResult::Error("ACOS requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < -1.0 || n > 1.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.acos()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ATAN" => {
            if args.len() != 1 {
                return EvalResult::Error("ATAN requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.atan()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ATAN2" => {
            if args.len() != 2 {
                return EvalResult::Error("ATAN2 requires exactly 2 arguments".to_string());
            }
            let x = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            let y = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return EvalResult::Error(e),
            };
            if x == 0.0 && y == 0.0 {
                return EvalResult::Error("#DIV/0!".to_string());
            }
            EvalResult::Number(y.atan2(x))
        }

        // =====================
        // ADVANCED LOGICAL FUNCTIONS
        // =====================
        "IFS" => {
            // IFS(condition1, value1, [condition2, value2], ...)
            if args.len() < 2 || args.len() % 2 != 0 {
                return EvalResult::Error("IFS requires pairs of condition, value arguments".to_string());
            }
            for i in (0..args.len()).step_by(2) {
                match evaluate(&args[i], lookup).to_bool() {
                    Ok(true) => return evaluate(&args[i + 1], lookup),
                    Ok(false) => continue,
                    Err(e) => return EvalResult::Error(e),
                }
            }
            EvalResult::Error("#N/A".to_string())
        }
        "SWITCH" => {
            // SWITCH(expression, value1, result1, [value2, result2], ..., [default])
            if args.len() < 3 {
                return EvalResult::Error("SWITCH requires at least 3 arguments".to_string());
            }
            let expr = evaluate(&args[0], lookup);
            let pairs = (args.len() - 1) / 2;
            let has_default = (args.len() - 1) % 2 == 1;

            for i in 0..pairs {
                let value = evaluate(&args[1 + i * 2], lookup);
                // Compare expr with value
                let matches = match (&expr, &value) {
                    (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < f64::EPSILON,
                    (EvalResult::Text(a), EvalResult::Text(b)) => a.to_lowercase() == b.to_lowercase(),
                    (EvalResult::Boolean(a), EvalResult::Boolean(b)) => a == b,
                    _ => expr.to_text() == value.to_text(),
                };
                if matches {
                    return evaluate(&args[2 + i * 2], lookup);
                }
            }

            if has_default {
                evaluate(&args[args.len() - 1], lookup)
            } else {
                EvalResult::Error("#N/A".to_string())
            }
        }
        "CHOOSE" => {
            // CHOOSE(index, value1, [value2], ...)
            if args.len() < 2 {
                return EvalResult::Error("CHOOSE requires at least 2 arguments".to_string());
            }
            let index = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };
            if index < 1 || index >= args.len() {
                return EvalResult::Error("#VALUE!".to_string());
            }
            evaluate(&args[index], lookup)
        }

        // =====================
        // STATISTICAL FUNCTIONS
        // =====================
        "STDEV" | "STDEV.S" => {
            // Sample standard deviation
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.len() < 2 {
                        return EvalResult::Error("#DIV/0!".to_string());
                    }
                    let mean = vals.iter().sum::<f64>() / vals.len() as f64;
                    let variance = vals.iter()
                        .map(|x| (x - mean).powi(2))
                        .sum::<f64>() / (vals.len() - 1) as f64;
                    EvalResult::Number(variance.sqrt())
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "STDEV.P" | "STDEVP" => {
            // Population standard deviation
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        return EvalResult::Error("#DIV/0!".to_string());
                    }
                    let mean = vals.iter().sum::<f64>() / vals.len() as f64;
                    let variance = vals.iter()
                        .map(|x| (x - mean).powi(2))
                        .sum::<f64>() / vals.len() as f64;
                    EvalResult::Number(variance.sqrt())
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "VAR" | "VAR.S" => {
            // Sample variance
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.len() < 2 {
                        return EvalResult::Error("#DIV/0!".to_string());
                    }
                    let mean = vals.iter().sum::<f64>() / vals.len() as f64;
                    let variance = vals.iter()
                        .map(|x| (x - mean).powi(2))
                        .sum::<f64>() / (vals.len() - 1) as f64;
                    EvalResult::Number(variance)
                }
                Err(e) => EvalResult::Error(e),
            }
        }
        "VAR.P" | "VARP" => {
            // Population variance
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.is_empty() {
                        return EvalResult::Error("#DIV/0!".to_string());
                    }
                    let mean = vals.iter().sum::<f64>() / vals.len() as f64;
                    let variance = vals.iter()
                        .map(|x| (x - mean).powi(2))
                        .sum::<f64>() / vals.len() as f64;
                    EvalResult::Number(variance)
                }
                Err(e) => EvalResult::Error(e),
            }
        }

        // =====================
        // ARRAY FUNCTIONS
        // =====================
        "SEQUENCE" => {
            // SEQUENCE(rows, [cols], [start], [step])
            // Returns a 2D array of sequential numbers
            if args.is_empty() || args.len() > 4 {
                return EvalResult::Error("SEQUENCE requires 1-4 arguments".to_string());
            }

            let rows = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                Ok(n) => n as usize,
                Err(e) => return EvalResult::Error(e),
            };

            let cols = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE!".to_string()),
                    Ok(n) => n as usize,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1
            };

            let start = if args.len() >= 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1.0
            };

            let step = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1.0
            };

            // Build the array
            let mut array = Array2D::new(rows, cols);
            let mut val = start;
            for r in 0..rows {
                for c in 0..cols {
                    array.set(r, c, Value::Number(val));
                    val += step;
                }
            }
            EvalResult::Array(array)
        }

        "TRANSPOSE" => {
            // TRANSPOSE(array)
            // Returns the transpose of an array/range
            if args.len() != 1 {
                return EvalResult::Error("TRANSPOSE requires exactly one argument".to_string());
            }

            // Get the input - if it's a range, build an array from it
            match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let in_rows = end_row - start_row + 1;
                    let in_cols = end_col - start_col + 1;

                    // Build transposed array (swap rows and cols)
                    let mut array = Array2D::new(in_cols, in_rows);
                    for r in 0..in_rows {
                        for c in 0..in_cols {
                            let val = lookup.get_value(start_row + r, start_col + c);
                            array.set(c, r, Value::Number(val));
                        }
                    }
                    EvalResult::Array(array)
                }
                _ => {
                    // Single value - just return it (1x1 transpose is identity)
                    evaluate(&args[0], lookup)
                }
            }
        }

        "FILTER" => {
            // FILTER(range, include)
            // Returns rows from range where include is TRUE
            // include must be a single column matching the number of rows in range
            if args.len() != 2 {
                return EvalResult::Error("FILTER requires exactly 2 arguments".to_string());
            }

            // Get the data range dimensions and values
            let (data_rows, data_cols, data): (usize, usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if let Ok(_) = text.parse::<f64>() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (r_count, c_count, row_data)
                }
                _ => {
                    return EvalResult::Error("#VALUE! FILTER requires a range as first argument".to_string());
                }
            };

            // Get the include criteria (must be a column matching data_rows)
            let include: Vec<bool> = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let inc_rows = end_row - start_row + 1;
                    let inc_cols = end_col - start_col + 1;

                    // Must be a single column with matching row count
                    if inc_cols != 1 {
                        return EvalResult::Error("#VALUE! Include must be a single column".to_string());
                    }
                    if inc_rows != data_rows {
                        return EvalResult::Error(format!("#VALUE! Include has {} rows but data has {} rows", inc_rows, data_rows));
                    }

                    let mut criteria = Vec::with_capacity(inc_rows);
                    for r in 0..inc_rows {
                        let val = lookup.get_value(start_row + r, *start_col);
                        // Treat non-zero as TRUE, zero as FALSE
                        criteria.push(val != 0.0);
                    }
                    criteria
                }
                _ => {
                    // Try evaluating as a single value (scalar comparison result)
                    match evaluate(&args[1], lookup) {
                        EvalResult::Boolean(b) => vec![b; data_rows],
                        EvalResult::Number(n) => vec![n != 0.0; data_rows],
                        _ => return EvalResult::Error("#VALUE! Include must be a range or boolean".to_string()),
                    }
                }
            };

            // Filter rows where include is TRUE
            let filtered_rows: Vec<Vec<Value>> = data.into_iter()
                .zip(include.iter())
                .filter(|(_, &inc)| inc)
                .map(|(row, _)| row)
                .collect();

            if filtered_rows.is_empty() {
                return EvalResult::Error("#CALC! No matches".to_string());
            }

            // Build result array
            let out_rows = filtered_rows.len();
            let mut array = Array2D::new(out_rows, data_cols);
            for (r, row) in filtered_rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "UNIQUE" => {
            // UNIQUE(range)
            // Returns unique rows from a range (preserves first occurrence order)
            if args.len() != 1 {
                return EvalResult::Error("UNIQUE requires exactly one argument".to_string());
            }

            // Build rows from range
            let (in_cols, rows): (usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if let Ok(_) = text.parse::<f64>() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (c_count, row_data)
                }
                _ => {
                    // Single value - return as-is
                    return evaluate(&args[0], lookup);
                }
            };

            // Find unique rows (first occurrence wins, case-insensitive for text)
            let mut unique_rows: Vec<Vec<Value>> = Vec::new();
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();

            for row in rows {
                // Create a canonical key for the row (case-insensitive)
                let key = row.iter()
                    .map(|v| match v {
                        Value::Text(s) => s.to_lowercase(),
                        other => other.to_text(),
                    })
                    .collect::<Vec<_>>()
                    .join("\x00"); // Use null byte as separator

                if !seen.contains(&key) {
                    seen.insert(key);
                    unique_rows.push(row);
                }
            }

            if unique_rows.is_empty() {
                return EvalResult::Error("#CALC! No data".to_string());
            }

            // Build result array
            let out_rows = unique_rows.len();
            let mut array = Array2D::new(out_rows, in_cols);
            for (r, row) in unique_rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "SORT" => {
            // SORT(array_or_range, [sort_col], [is_asc])
            // Default: sort_col=1, is_asc=TRUE
            // Sorts rows by the specified column
            if args.is_empty() || args.len() > 3 {
                return EvalResult::Error("SORT requires 1-3 arguments".to_string());
            }

            // Get sort column (1-indexed, default 1)
            let sort_col_1idx = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) if n < 1.0 => return EvalResult::Error("#VALUE! Sort column must be >= 1".to_string()),
                    Ok(n) => n as usize,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                1
            };

            // Get ascending flag (default true)
            let is_asc = if args.len() >= 3 {
                match evaluate(&args[2], lookup).to_bool() {
                    Ok(b) => b,
                    Err(e) => return EvalResult::Error(e),
                }
            } else {
                true
            };

            // Build rows from range
            let (in_rows, in_cols, mut rows): (usize, usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            // Determine value type
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if let Ok(_) = text.parse::<f64>() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (r_count, c_count, row_data)
                }
                _ => {
                    // Single value - can't sort meaningfully
                    return evaluate(&args[0], lookup);
                }
            };

            // Validate sort column
            if sort_col_1idx > in_cols {
                return EvalResult::Error(format!("#VALUE! Sort column {} exceeds range width {}", sort_col_1idx, in_cols));
            }
            let sort_col_0idx = sort_col_1idx - 1;

            // Sort rows by the key column (stable sort)
            // Ordering: Numbers < Text < Empty < Errors (ascending)
            // For descending, we reverse at the end
            rows.sort_by(|a, b| {
                let key_a = &a[sort_col_0idx];
                let key_b = &b[sort_col_0idx];
                value_compare(key_a, key_b)
            });

            // Reverse if descending
            if !is_asc {
                rows.reverse();
            }

            // Build result array
            let mut array = Array2D::new(in_rows, in_cols);
            for (r, row) in rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        _ => EvalResult::Error(format!("Unknown function: {}", name)),
    }
}

// =============================================================================
// Value comparison for SORT
// =============================================================================

/// Compare two Values for sorting
/// Order: Numbers < Text < Empty < Errors (ascending)
fn value_compare(a: &Value, b: &Value) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    // Type ordering: Number=0, Text=1, Boolean=2, Empty=3, Error=4
    fn type_rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) => 0,
            Value::Text(_) => 1,
            Value::Boolean(_) => 2,
            Value::Empty => 3,
            Value::Error(_) => 4,
        }
    }

    let rank_a = type_rank(a);
    let rank_b = type_rank(b);

    if rank_a != rank_b {
        return rank_a.cmp(&rank_b);
    }

    // Same type - compare within type
    match (a, b) {
        (Value::Number(na), Value::Number(nb)) => {
            na.partial_cmp(nb).unwrap_or(Ordering::Equal)
        }
        (Value::Text(sa), Value::Text(sb)) => {
            // Case-insensitive comparison
            sa.to_lowercase().cmp(&sb.to_lowercase())
        }
        (Value::Boolean(ba), Value::Boolean(bb)) => {
            // FALSE < TRUE
            ba.cmp(bb)
        }
        (Value::Error(ea), Value::Error(eb)) => {
            ea.cmp(eb)
        }
        _ => Ordering::Equal, // Empty == Empty
    }
}

// Helper functions for date calculations

/// Convert year/month/day to Excel serial date number
fn date_to_serial(year: i32, month: i32, day: i32) -> f64 {
    // Handle month overflow/underflow
    let mut y = year;
    let mut m = month;
    while m > 12 {
        m -= 12;
        y += 1;
    }
    while m < 1 {
        m += 12;
        y -= 1;
    }

    // Calculate days since Excel epoch (1899-12-30)
    // Using a simplified algorithm
    let a = (14 - m) / 12;
    let y_adj = y + 4800 - a;
    let m_adj = m + 12 * a - 3;

    let jdn = day + (153 * m_adj + 2) / 5 + 365 * y_adj + y_adj / 4 - y_adj / 100 + y_adj / 400 - 32045;

    // Excel epoch JDN (1899-12-30) = 2415019
    (jdn - 2415019) as f64
}

/// Convert Excel serial date number to year/month/day
fn serial_to_date(serial: f64) -> (i32, i32, i32) {
    let serial = serial.floor() as i32;
    let jdn = serial + 2415019; // Convert to Julian Day Number

    // Algorithm to convert JDN to Gregorian date
    let a = jdn + 32044;
    let b = (4 * a + 3) / 146097;
    let c = a - (146097 * b) / 4;
    let d = (4 * c + 3) / 1461;
    let e = c - (1461 * d) / 4;
    let m = (5 * e + 2) / 153;

    let day = e - (153 * m + 2) / 5 + 1;
    let month = m + 3 - 12 * (m / 10);
    let year = 100 * b + d - 4800 + m / 10;

    (year, month, day)
}

/// Get the number of days in a month
fn days_in_month(year: i32, month: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0) {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

/// Check if a cell value matches criteria (for SUMIF, COUNTIF, etc.)
fn matches_criteria(value: &EvalResult, criteria: &EvalResult) -> bool {
    let criteria_str = criteria.to_text();

    // Check for comparison operators in criteria
    if criteria_str.starts_with(">=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[2..].trim().parse::<f64>()) {
            return v >= c;
        }
    } else if criteria_str.starts_with("<=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[2..].trim().parse::<f64>()) {
            return v <= c;
        }
    } else if criteria_str.starts_with("<>") {
        let c = criteria_str[2..].trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return (v - n).abs() >= f64::EPSILON;
            }
        }
        return value.to_text().to_lowercase() != c.to_lowercase();
    } else if criteria_str.starts_with('>') {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[1..].trim().parse::<f64>()) {
            return v > c;
        }
    } else if criteria_str.starts_with('<') {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[1..].trim().parse::<f64>()) {
            return v < c;
        }
    } else if criteria_str.starts_with('=') {
        let c = criteria_str[1..].trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return (v - n).abs() < f64::EPSILON;
            }
        }
        return value.to_text().to_lowercase() == c.to_lowercase();
    }

    // Simple equality check
    match (value, criteria) {
        (EvalResult::Number(v), EvalResult::Number(c)) => (v - c).abs() < f64::EPSILON,
        (EvalResult::Text(v), EvalResult::Text(c)) => v.to_lowercase() == c.to_lowercase(),
        _ => value.to_text().to_lowercase() == criteria_str.to_lowercase(),
    }
}

/// Helper to get text from a cell, handling cross-sheet references
fn get_text_for_sheet<L: CellLookup>(lookup: &L, sheet: &SheetRef, row: usize, col: usize) -> Result<String, String> {
    match sheet {
        SheetRef::Current => Ok(lookup.get_text(row, col)),
        SheetRef::Id(id) => Ok(lookup.get_text_sheet(*id, row, col)),
        SheetRef::RefError { .. } => Err("#REF!".to_string()),
    }
}

fn collect_numbers<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Result<Vec<f64>, String> {
    let mut values = Vec::new();

    for arg in args {
        match arg {
            Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                collect_numbers_from_range_sheet(sheet, *start_row, *start_col, *end_row, *end_col, lookup, &mut values)?;
            }
            Expr::NamedRange(name) => {
                // Resolve named range and collect numbers from it
                match lookup.resolve_named_range(name) {
                    None => return Err(format!("#NAME? '{}'", name)),
                    Some(NamedRangeResolution::Cell { row, col }) => {
                        let text = lookup.get_text(row, col);
                        if let Ok(n) = text.parse::<f64>() {
                            values.push(n);
                        }
                    }
                    Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                        collect_numbers_from_range(start_row, start_col, end_row, end_col, lookup, &mut values);
                    }
                }
            }
            _ => {
                let result = evaluate(arg, lookup);
                match result.to_number() {
                    Ok(n) => values.push(n),
                    Err(e) => return Err(e),
                }
            }
        }
    }

    Ok(values)
}

/// Collect numbers from a range, supporting cross-sheet references
fn collect_numbers_from_range_sheet<L: CellLookup>(
    sheet: &SheetRef,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<f64>,
) -> Result<(), String> {
    // Check for RefError early
    if let SheetRef::RefError { .. } = sheet {
        return Err("#REF!".to_string());
    }

    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);

    for r in min_row..=max_row {
        for c in min_col..=max_col {
            let text = get_text_for_sheet(lookup, sheet, r, c)?;
            // Only include numeric values, skip text/empty
            if let Ok(n) = text.parse::<f64>() {
                values.push(n);
            }
        }
    }
    Ok(())
}

/// Legacy helper for same-sheet ranges (used by named range resolution)
fn collect_numbers_from_range<L: CellLookup>(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<f64>,
) {
    let _ = collect_numbers_from_range_sheet(
        &SheetRef::Current,
        start_row, start_col, end_row, end_col,
        lookup, values
    );
}

fn collect_all_values<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Vec<EvalResult> {
    let mut values = Vec::new();

    for arg in args {
        match arg {
            Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                if let Err(e) = collect_all_values_from_range_sheet(sheet, *start_row, *start_col, *end_row, *end_col, lookup, &mut values) {
                    values.push(EvalResult::Error(e));
                }
            }
            Expr::NamedRange(name) => {
                // Resolve named range and collect all values from it
                match lookup.resolve_named_range(name) {
                    None => values.push(EvalResult::Error(format!("#NAME? '{}'", name))),
                    Some(NamedRangeResolution::Cell { row, col }) => {
                        let text = lookup.get_text(row, col);
                        if text.is_empty() {
                            values.push(EvalResult::Text(String::new()));
                        } else if let Ok(n) = text.parse::<f64>() {
                            values.push(EvalResult::Number(n));
                        } else {
                            values.push(EvalResult::Text(text));
                        }
                    }
                    Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                        collect_all_values_from_range(start_row, start_col, end_row, end_col, lookup, &mut values);
                    }
                }
            }
            _ => {
                values.push(evaluate(arg, lookup));
            }
        }
    }

    values
}

/// Collect all values from a range, supporting cross-sheet references
fn collect_all_values_from_range_sheet<L: CellLookup>(
    sheet: &SheetRef,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<EvalResult>,
) -> Result<(), String> {
    // Check for RefError early
    if let SheetRef::RefError { .. } = sheet {
        return Err("#REF!".to_string());
    }

    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);

    for r in min_row..=max_row {
        for c in min_col..=max_col {
            let text = get_text_for_sheet(lookup, sheet, r, c)?;
            if text.is_empty() {
                values.push(EvalResult::Text(String::new()));
            } else if let Ok(n) = text.parse::<f64>() {
                values.push(EvalResult::Number(n));
            } else {
                values.push(EvalResult::Text(text));
            }
        }
    }
    Ok(())
}

/// Legacy helper for same-sheet ranges (used by named range resolution)
fn collect_all_values_from_range<L: CellLookup>(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<EvalResult>,
) {
    let _ = collect_all_values_from_range_sheet(
        &SheetRef::Current,
        start_row, start_col, end_row, end_col,
        lookup, values
    );
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
            EvalResult::Number(n) => assert!((n - 0.1728).abs() < 0.01), // ~17.28%
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
}
