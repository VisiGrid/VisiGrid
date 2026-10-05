// Scalar functions over arrays ("lifting"), and error pass-through.
//
// A function whose parameters are plain values — LEN, UPPER, ABS, ROUND, the
// IS* family, YEAR — applies to each element when given a range or an array
// and returns an array, as in Excel: =LEN(A2:A5) spills 3, 2, 2, 2 and
// SUMPRODUCT(--ISNUMBER(B2:B5)) counts the numbers. Before this, a range
// argument evaluated to the error "#VALUE! Array arithmetic not supported",
// and functions that do not check for errors treated that as data: LEN(A2:A5)
// was 38, the length of the message (#43).
//
// The same pass makes those functions return an error argument instead of
// reading it: LEN(#N/A) is #N/A. The IS* functions and IFERROR/IFNA are the
// exception — seeing the error is their job.
//
// Arguments are evaluated once here and handed on as literals, so nothing is
// evaluated twice. Functions that need a reference rather than a value (ROW,
// COLUMN, OFFSET, INDEX, COUNTBLANK) and those that already take or return
// arrays are not in the table and are untouched.

use super::eval::{Array2D, EvalResult, Value};
use super::parser::{BoundExpr, Expr};

/// Functions whose every argument is a plain value.
const VALUE_FUNCTIONS: &[&str] = &[
    // text
    "LEN", "LEFT", "RIGHT", "MID", "UPPER", "LOWER", "TRIM", "PROPER", "REPT",
    "SUBSTITUTE", "FIND", "SEARCH", "EXACT", "VALUE", "TEXT", "TEXTBEFORE",
    "TEXTAFTER", "REPLACE", "REGEXTEST", "REGEXMATCH", "REGEXREPLACE",
    // math and trig
    "ABS", "INT", "SQRT", "EXP", "LN", "LOG", "LOG10", "POWER", "MOD", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "TRUNC", "CEILING", "FLOOR", "SIN", "COS", "TAN",
    "ASIN", "ACOS", "ATAN", "ATAN2", "DEGREES", "RADIANS",
    // dates
    "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "WEEKDAY", "DATEVALUE",
    "EDATE", "EOMONTH",
    // logical and information
    "NOT", "ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISNA",
];

/// Functions that answer for an error rather than returning it.
const ERROR_AWARE: &[&str] = &["ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISNA"];

pub(super) enum Lifted {
    /// The function has been evaluated (per element, or to an argument's error).
    Done(EvalResult),
    /// Evaluate normally with these arguments (literals in place of values).
    Args(Vec<BoundExpr>),
    /// Not a lifted function; evaluate as usual.
    No,
}

/// Decide how `name` evaluates given its arguments.
///
/// `operand` evaluates an argument, turning a range into an array; `call`
/// evaluates the function with a given argument list.
pub(super) fn lift(
    name: &str,
    args: &[BoundExpr],
    operand: &dyn Fn(&BoundExpr) -> EvalResult,
    call: &dyn Fn(&[BoundExpr]) -> EvalResult,
) -> Lifted {
    match name {
        "IF" => return lift_if(args, operand, call),
        "IFERROR" | "IFNA" => return lift_iferror(name, args, operand, call),
        _ => {}
    }
    if !VALUE_FUNCTIONS.contains(&name) {
        return Lifted::No;
    }
    let values: Vec<EvalResult> = args.iter().map(|a| operand(a)).collect();
    if !values.iter().any(is_multi_cell) {
        // Scalars only. An error argument is the answer unless the function
        // inspects errors, in which case it evaluates from the original argument.
        if let Some(EvalResult::Error(e)) = values.iter().find(|v| matches!(v, EvalResult::Error(_))) {
            return if ERROR_AWARE.contains(&name) {
                Lifted::No
            } else {
                Lifted::Done(EvalResult::Error(e.clone()))
            };
        }
        return Lifted::Args(values.iter().map(literal).collect());
    }
    let error_aware = ERROR_AWARE.contains(&name);
    Lifted::Done(per_element(&values, |elems| {
        if let Some(e) = elems.iter().find_map(|v| match v {
            Value::Error(e) => Some(e.clone()),
            _ => None,
        }) {
            return if error_aware {
                // Only these five take one argument, which is the error.
                EvalResult::Boolean(match name {
                    "ISERROR" => true,
                    "ISNA" => e.starts_with("#N/A"),
                    _ => false,
                })
            } else {
                EvalResult::Error(e)
            };
        }
        call(&elems.iter().map(|v| literal(&EvalResult::from_value(v))).collect::<Vec<_>>())
    }))
}

/// IF over an array condition: each element picks from its own branch element.
/// Branches stay lazy for a scalar condition, so IF(x>0, 1/x, 0) is not an
/// error when x is 0.
fn lift_if(
    args: &[BoundExpr],
    operand: &dyn Fn(&BoundExpr) -> EvalResult,
    call: &dyn Fn(&[BoundExpr]) -> EvalResult,
) -> Lifted {
    let Some(cond_expr) = args.first() else { return Lifted::No };
    let cond = operand(cond_expr);
    if !is_multi_cell(&cond) {
        return match cond {
            EvalResult::Error(_) => Lifted::No,
            c => {
                let mut v = args.to_vec();
                v[0] = literal(&c);
                Lifted::Args(v)
            }
        };
    }
    let mut values = vec![cond];
    values.extend(args.iter().skip(1).map(|a| operand(a)));
    Lifted::Done(per_element(&values, |elems| {
        let cond = match &elems[0] {
            Value::Error(e) => return EvalResult::Error(e.clone()),
            c => EvalResult::from_value(c),
        };
        // Which branch IF will take. If that branch's element is an error, it is
        // the answer; the branch not taken may hold errors freely.
        let chosen = match cond.to_bool() {
            Ok(true) => 1,
            Ok(false) => 2,
            Err(e) => return EvalResult::Error(e),
        };
        if let Some(Value::Error(e)) = elems.get(chosen) {
            return EvalResult::Error(e.clone());
        }
        let v: Vec<BoundExpr> = elems
            .iter()
            .map(|x| match x {
                Value::Error(_) => Expr::Empty,
                other => literal(&EvalResult::from_value(other)),
            })
            .collect();
        call(&v)
    }))
}

/// IFERROR/IFNA over an array: each erroring element takes its fallback.
fn lift_iferror(
    name: &str,
    args: &[BoundExpr],
    operand: &dyn Fn(&BoundExpr) -> EvalResult,
    call: &dyn Fn(&[BoundExpr]) -> EvalResult,
) -> Lifted {
    let Some(first) = args.first() else { return Lifted::No };
    let value = operand(first);
    if !is_multi_cell(&value) {
        return match value {
            EvalResult::Error(_) => Lifted::No,
            v => {
                let mut a = args.to_vec();
                a[0] = literal(&v);
                Lifted::Args(a)
            }
        };
    }
    let mut values = vec![value];
    values.extend(args.iter().skip(1).map(|a| operand(a)));
    Lifted::Done(per_element(&values, |elems| {
        let catches = match &elems[0] {
            Value::Error(e) => name == "IFERROR" || e.starts_with("#N/A"),
            _ => false,
        };
        match (catches, elems.get(1)) {
            (true, Some(fallback)) => EvalResult::from_value(fallback),
            (true, None) => EvalResult::Empty,
            (false, _) if matches!(elems[0], Value::Error(_)) => EvalResult::from_value(&elems[0]),
            (false, _) => call(&elems.iter().map(|v| literal(&EvalResult::from_value(v))).collect::<Vec<_>>()),
        }
    }))
}

fn is_multi_cell(v: &EvalResult) -> bool {
    matches!(v, EvalResult::Array(a) if a.rows() * a.cols() > 1)
}

/// A value as a literal argument. A 1x1 array is its single value.
fn literal(v: &EvalResult) -> BoundExpr {
    match v {
        EvalResult::Number(n) => Expr::Number(*n),
        EvalResult::Text(s) => Expr::Text(s.clone()),
        EvalResult::Boolean(b) => Expr::Boolean(*b),
        EvalResult::Empty => Expr::Empty,
        EvalResult::Array(a) => literal(&EvalResult::from_value(&a.top_left())),
        // Callers handle errors before asking for a literal.
        EvalResult::Error(_) => Expr::RefError,
    }
}

/// Apply `f` to each element position, broadcasting scalars and single rows or
/// columns as the operators do; positions two arrays don't share are #N/A.
fn per_element(values: &[EvalResult], f: impl Fn(&[Value]) -> EvalResult) -> EvalResult {
    let dims: Vec<(usize, usize)> = values
        .iter()
        .map(|v| match v {
            EvalResult::Array(a) => (a.rows(), a.cols()),
            _ => (1, 1),
        })
        .collect();
    let size = |pick: fn(&(usize, usize)) -> usize| {
        dims.iter().map(pick).fold(1, |acc, d| if acc == 1 { d } else if d == 1 { acc } else { acc.max(d) })
    };
    let (rows, cols) = (size(|d| d.0), size(|d| d.1));
    if let Err(error) = super::eval_budget::array(rows, cols) { return EvalResult::Error(error); }
    let mut out = Array2D::new(rows, cols);
    for r in 0..rows {
        for c in 0..cols {
            let mut elems = Vec::with_capacity(values.len());
            let mut missing = false;
            for (v, &(vr, vc)) in values.iter().zip(&dims) {
                match v {
                    EvalResult::Array(a) => {
                        let (rr, cc) = (if vr == 1 { 0 } else { r }, if vc == 1 { 0 } else { c });
                        match a.get(rr, cc) {
                            Some(x) => elems.push(x.clone()),
                            None => missing = true,
                        }
                    }
                    other => elems.push(other.to_value()),
                }
            }
            let result = if missing { EvalResult::Error("#N/A".to_string()) } else { f(&elems) };
            out.set(r, c, result.to_value());
        }
    }
    EvalResult::Array(out)
}
