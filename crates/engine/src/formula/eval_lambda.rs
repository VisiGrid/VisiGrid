// LAMBDA helper functions: MAP, REDUCE, SCAN, BYROW, BYCOL.
//
// Each takes arrays (a range, an array constant, a computed array) and a
// LAMBDA, and calls the LAMBDA once per element, row or column. The call
// itself is eval_let.rs's: the LAMBDA's parameters are bound to the values,
// exactly as LAMBDA(x, x+1)(2) binds x. The LAMBDA may also be a LET name
// bound to one, or a built-in function's name (BYROW(A1:C3, SUM)).
//
// The LAMBDA sees values, not references: an element of a range is the
// cell's typed value, and a row passed by BYROW is an array (a single value
// when the array is one column wide). An error element is passed to the
// LAMBDA like any other value; deciding what it means is the LAMBDA's job.
//
// MAP, SCAN, BYROW and BYCOL build arrays of the LAMBDA's answers, and an
// answer that is itself more than one cell is #CALC! (an array of arrays
// cannot spill), as in Excel. REDUCE returns its final accumulator, which
// may be an array. A blank answer is 0, as =A1 shows 0 for a blank A1.
//
// A first argument that is not a LAMBDA, or a LAMBDA with the wrong number of
// parameters, is #VALUE!.

use super::eval::{Array2D, CellLookup, EvalResult, Value};
use super::eval_helpers::arg_values;
use super::eval_let::{call_with_values, is_callable, lambda_arity};
use super::parser::{BoundExpr, Expr};

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "MAP" => {
            // MAP(array1, [array2, ...], lambda): the LAMBDA applied to each
            // element, one parameter per array. Arrays of different sizes
            // pair up as the operators pair them: a single value goes with
            // every element, a single row or column stretches, and where one
            // array has no element the answer is #N/A.
            if args.len() < 2 {
                return Some(EvalResult::Error("MAP requires an array and a LAMBDA".to_string()));
            }
            let (arrays, lambda) = args.split_at(args.len() - 1);
            let lambda = &lambda[0];
            if let Err(e) = check_lambda(lambda, arrays.len(), lookup) {
                return Some(EvalResult::Error(e));
            }
            let mut grids = Vec::with_capacity(arrays.len());
            for arg in arrays {
                match grid(arg, lookup) {
                    Ok(g) => grids.push(g),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            }
            let stretch = |a: usize, b: usize| if a == 1 { b } else if b == 1 { a } else { a.max(b) };
            let rows = grids.iter().map(Array2D::rows).fold(1, stretch);
            let cols = grids.iter().map(Array2D::cols).fold(1, stretch);
            if let Err(error) = super::eval_budget::array(rows, cols) {
                return Some(EvalResult::Error(error));
            }
            let mut out = Array2D::new(rows, cols);
            for r in 0..rows {
                for c in 0..cols {
                    let picked: Option<Vec<EvalResult>> = grids.iter().map(|g| {
                        let gr = if g.rows() == 1 { 0 } else { r };
                        let gc = if g.cols() == 1 { 0 } else { c };
                        g.get(gr, gc).map(EvalResult::from_value)
                    }).collect();
                    let value = match picked {
                        Some(values) => match single(call_with_values(lambda, values, lookup)) {
                            Ok(v) => v,
                            Err(e) => return Some(EvalResult::Error(e)),
                        },
                        None => Value::Error("#N/A".to_string()),
                    };
                    out.set(r, c, value);
                }
            }
            shrink(out)
        }
        "REDUCE" | "SCAN" => {
            // REDUCE([initial_value], array, lambda(accumulator, value)):
            // the LAMBDA folded over the array's elements in row-major order,
            // starting from initial_value (blank when omitted). SCAN takes the
            // same arguments and returns every intermediate accumulator, in
            // an array the shape of the input.
            if args.len() != 3 {
                return Some(EvalResult::Error(format!("{name} requires exactly 3 arguments")));
            }
            if let Err(e) = check_lambda(&args[2], 2, lookup) {
                return Some(EvalResult::Error(e));
            }
            let mut acc = match &args[0] {
                Expr::Empty => EvalResult::Empty,
                initial => match grid(initial, lookup) {
                    Ok(g) => shrink(g),
                    Err(e) => return Some(EvalResult::Error(e)),
                },
            };
            let items = match grid(&args[1], lookup) {
                Ok(g) => g,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if items.rows() * items.cols() == 0 {
                return Some(EvalResult::Error("#CALC! Empty array".to_string()));
            }
            let scan = name == "SCAN";
            if scan {
                if let Err(error) = super::eval_budget::array(items.rows(), items.cols()) {
                    return Some(EvalResult::Error(error));
                }
            }
            let mut out = Array2D::new(if scan { items.rows() } else { 0 }, items.cols());
            for r in 0..items.rows() {
                for c in 0..items.cols() {
                    let item = EvalResult::from_value(items.get(r, c).unwrap_or(&Value::Empty));
                    acc = call_with_values(&args[2], vec![acc, item], lookup);
                    if scan {
                        match single(acc.clone()) {
                            Ok(v) => out.set(r, c, v),
                            Err(e) => return Some(EvalResult::Error(e)),
                        }
                    }
                }
            }
            if scan {
                shrink(out)
            } else {
                match acc {
                    EvalResult::Empty => EvalResult::Number(0.0),
                    other => other,
                }
            }
        }
        "BYROW" | "BYCOL" => {
            // BYROW(array, lambda(row)): the LAMBDA applied to each row,
            // giving a column of answers. BYCOL is the same by column, giving
            // a row. Each answer must be a single value.
            if args.len() != 2 {
                return Some(EvalResult::Error(format!("{name} requires exactly 2 arguments")));
            }
            if let Err(e) = check_lambda(&args[1], 1, lookup) {
                return Some(EvalResult::Error(e));
            }
            let items = match grid(&args[0], lookup) {
                Ok(g) => g,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if items.rows() * items.cols() == 0 {
                return Some(EvalResult::Error("#CALC! Empty array".to_string()));
            }
            let by_row = name == "BYROW";
            let (count, len) = if by_row { (items.rows(), items.cols()) } else { (items.cols(), items.rows()) };
            if let Err(error) = super::eval_budget::array(count, 1) {
                return Some(EvalResult::Error(error));
            }
            let mut out = if by_row { Array2D::new(count, 1) } else { Array2D::new(1, count) };
            for i in 0..count {
                let mut slice = if by_row { Array2D::new(1, len) } else { Array2D::new(len, 1) };
                for j in 0..len {
                    let (r, c) = if by_row { (i, j) } else { (j, i) };
                    let v = items.get(r, c).cloned().unwrap_or(Value::Empty);
                    if by_row { slice.set(0, j, v) } else { slice.set(j, 0, v) }
                }
                let slice = shrink(slice);
                let value = match single(call_with_values(&args[1], vec![slice], lookup)) {
                    Ok(v) => v,
                    Err(e) => return Some(EvalResult::Error(e)),
                };
                if by_row { out.set(i, 0, value) } else { out.set(0, i, value) }
            }
            shrink(out)
        }
        _ => return None,
    };
    Some(result)
}

/// #VALUE! unless `lambda` can be called with `params` arguments. A LAMBDA
/// written in place is checked now; one reached through a LET or a curried
/// call is checked when it is called, which gives the same #VALUE!.
fn check_lambda<L: CellLookup>(lambda: &BoundExpr, params: usize, lookup: &L) -> Result<(), String> {
    if !is_callable(lambda, lookup) {
        return Err("#VALUE! Expected a LAMBDA".to_string());
    }
    match lambda_arity(lambda) {
        Some(n) if n != params => Err(format!("#VALUE! The LAMBDA needs {params} parameter{}", if params == 1 { "" } else { "s" })),
        _ => Ok(()),
    }
}

/// An argument as a grid of values: a range or array as it is, anything else
/// as one value. An error argument (not an error inside an array) is the
/// helper's answer.
fn grid<L: CellLookup>(arg: &BoundExpr, lookup: &L) -> Result<Array2D, String> {
    let got = arg_values(arg, lookup)?;
    if !got.bulk {
        if let Some(Value::Error(e)) = got.values.first() {
            return Err(e.clone());
        }
    }
    let mut out = Array2D::new(got.rows, got.cols);
    for (i, v) in got.values.into_iter().enumerate() {
        out.set(i / got.cols.max(1), i % got.cols.max(1), v);
    }
    Ok(out)
}

/// A one-cell array as its value; anything else unchanged.
fn shrink(a: Array2D) -> EvalResult {
    match a.to_scalar() {
        Some(v) => EvalResult::from_value(&v),
        None => EvalResult::Array(a),
    }
}

/// One LAMBDA answer as a cell value: a one-cell array is its value, a blank
/// is 0, and a larger array is #CALC!.
fn single(result: EvalResult) -> Result<Value, String> {
    match result {
        EvalResult::Array(a) if a.rows() * a.cols() != 1 => {
            Err("#CALC! A LAMBDA returned an array where a single value is needed".to_string())
        }
        EvalResult::Array(a) => Ok(match a.top_left() {
            Value::Empty => Value::Number(0.0),
            v => v,
        }),
        EvalResult::Empty => Ok(Value::Number(0.0)),
        other => Ok(other.to_value()),
    }
}

#[cfg(test)]
mod tests {
    use crate::formula::eval::{evaluate, CellLookup, EvalResult};
    use crate::formula::parser::{bind_expr_same_sheet, parse};

    struct Empty;
    impl CellLookup for Empty {
        fn get_value(&self, _r: usize, _c: usize) -> f64 { 0.0 }
        fn get_text(&self, _r: usize, _c: usize) -> String { String::new() }
    }

    fn eval(formula: &str) -> EvalResult {
        evaluate(&bind_expr_same_sheet(&parse(formula).unwrap()), &Empty)
    }

    fn error(formula: &str) -> String {
        match eval(formula) {
            EvalResult::Error(e) => e,
            other => panic!("{formula} gave {other:?}, expected an error"),
        }
    }

    /// The value table a call binds into is released after every call, so
    /// a long MAP does not hold every element it has seen.
    #[test]
    fn calls_release_their_bound_values() {
        let before = super::super::eval_let::bound_value_count();
        assert!(matches!(eval("=MAP(SEQUENCE(50),LAMBDA(x,SEQUENCE(1,1,x)))"), EvalResult::Array(_)));
        assert_eq!(super::super::eval_let::bound_value_count(), before);
    }

    #[test]
    fn a_lambda_with_the_wrong_parameter_count_is_value() {
        assert!(error("=MAP({1,2},LAMBDA(a,b,a+b))").starts_with("#VALUE!"));
        assert!(error("=REDUCE(0,{1,2},LAMBDA(a,a))").starts_with("#VALUE!"));
        assert!(error("=BYROW({1,2},LAMBDA(a,b,a))").starts_with("#VALUE!"));
        // Through a LET name, the count is only known at the call; same answer.
        assert!(error("=LET(f,LAMBDA(a,b,a+b),MAP({1,2},f))").starts_with("#VALUE!"));
    }
}
