// LET and LAMBDA.
//
// LET(name1, value1, [name2, value2, ...], calculation) names intermediate
// results; LAMBDA(param1, ..., calculation) is a function written in a
// formula, called where it is written — LAMBDA(x, x+1)(2) — or through a LET
// name — LET(f, LAMBDA(x, x*2), f(3)).
//
// Both work by substitution at evaluation time; the stored formula is never
// rewritten. When a LET is evaluated, each value is evaluated once, in order,
// and every later use of its name is replaced by that result as a literal. A
// value that is a reference (A1, A1:B9, a defined name, a table column) is
// substituted as the reference itself, so SUM(x) and COUNTIF(x, ...) keep
// range semantics. A LAMBDA bound to a name is kept as a LAMBDA; calling it
// becomes a LET of its parameters, which is exactly the meaning of a call.
//
// Names are lexically scoped: an inner LET or a LAMBDA parameter of the same
// name shadows an outer one.
//
// Not supported: a LAMBDA passed as a value to MAP/REDUCE/BYROW and the other
// helper functions (which the engine does not implement), recursion (a LAMBDA
// cannot call the LET name it is bound to), and LAMBDAs defined in the Name
// Manager. A LAMBDA that is never called evaluates to #CALC!, as in Excel.

use std::collections::HashMap;

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, Value};
use super::parser::{BoundExpr, Expr, ARRAY_LITERAL, INVOKE};

/// Internal literal for a value with no expression of its own: an error, an
/// empty value, or an array. Never written by a user (underscore prefix).
pub(super) const LITERAL_ERROR: &str = "_LET.ERR";
pub(super) const LITERAL_EMPTY: &str = "_LET.EMPTY";
pub(super) const LITERAL_ARRAY: &str = ARRAY_LITERAL;

#[derive(Clone)]
enum Binding {
    /// A literal or a reference, substituted as an expression.
    Expr(BoundExpr),
    /// A LAMBDA, already closed over the names in scope where it was defined.
    Lambda { params: Vec<String>, body: BoundExpr },
}

type Env = HashMap<String, Binding>;

/// Evaluate the functions this module owns, or None for any other name.
pub(super) fn try_evaluate<L: CellLookup>(name: &str, args: &[BoundExpr], lookup: &L) -> Option<EvalResult> {
    Some(match name {
        "LET" => evaluate_let(args, lookup),
        // A LAMBDA reached by evaluation was never called.
        "LAMBDA" => EvalResult::Error("#CALC!".to_string()),
        _ if name == INVOKE => {
            let expanded = substitute(&Expr::Function { name: INVOKE.to_string(), args: args.to_vec() }, &Env::new());
            if matches!(&expanded, Expr::Function { name, .. } if name == INVOKE) {
                // The callee is not a LAMBDA, so there is nothing to call.
                EvalResult::Error("#VALUE!".to_string())
            } else {
                evaluate(&expanded, lookup)
            }
        }
        LITERAL_EMPTY => EvalResult::Empty,
        LITERAL_ERROR => match args.first() {
            Some(Expr::Text(e)) => EvalResult::Error(e.clone()),
            _ => EvalResult::Error("#VALUE!".to_string()),
        },
        LITERAL_ARRAY => {
            let dim = |i: usize| match args.get(i) {
                Some(Expr::Number(n)) => *n as usize,
                _ => 0,
            };
            let (rows, cols) = (dim(0), dim(1));
            let mut out = Array2D::new(rows, cols);
            for (i, element) in args.iter().skip(2).enumerate().take(rows * cols) {
                out.set(i / cols, i % cols, evaluate(element, lookup).to_value());
            }
            EvalResult::Array(out)
        }
        _ => return None,
    })
}

fn evaluate_let<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> EvalResult {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return EvalResult::Error("#VALUE! LET needs name, value pairs and a calculation".to_string());
    }
    let mut env = Env::new();
    let (pairs, body) = args.split_at(args.len() - 1);
    for pair in pairs.chunks(2) {
        let Expr::NamedRange(name) = &pair[0] else {
            return EvalResult::Error("#VALUE! LET names must be names".to_string());
        };
        let value = substitute(&pair[1], &env);
        let binding = match value {
            Expr::Function { name: ref f, ref args } if f == "LAMBDA" => match lambda_parts(args) {
                Some((params, body)) => Binding::Lambda { params, body },
                None => Binding::Expr(error_literal("#VALUE!")),
            },
            ref v if is_reference(v) => Binding::Expr(value),
            v => Binding::Expr(literal(&evaluate(&v, lookup))),
        };
        env.insert(name.clone(), binding);
    }
    evaluate(&substitute(&body[0], &env), lookup)
}

/// A LAMBDA's parameter names and calculation, or None if malformed.
fn lambda_parts(args: &[BoundExpr]) -> Option<(Vec<String>, BoundExpr)> {
    let (body, params) = args.split_last()?;
    let names = params
        .iter()
        .map(|p| match p {
            Expr::NamedRange(n) => Some(n.clone()),
            _ => None,
        })
        .collect::<Option<Vec<_>>>()?;
    Some((names, body.clone()))
}

fn is_reference(expr: &BoundExpr) -> bool {
    matches!(
        expr,
        Expr::CellRef { .. } | Expr::Range { .. } | Expr::WholeRange { .. } | Expr::NamedRange(_)
            | Expr::StructuredRef(_) | Expr::RefError
    )
}

fn error_literal(e: &str) -> BoundExpr {
    Expr::Function { name: LITERAL_ERROR.to_string(), args: vec![Expr::Text(e.to_string())] }
}

fn value_literal(v: &Value) -> BoundExpr {
    match v {
        Value::Number(n) => Expr::Number(*n),
        Value::Text(s) => Expr::Text(s.clone()),
        Value::Boolean(b) => Expr::Boolean(*b),
        Value::Empty => Expr::Function { name: LITERAL_EMPTY.to_string(), args: Vec::new() },
        Value::Error(e) => error_literal(e),
    }
}

/// An evaluated result as an expression that evaluates back to it.
fn literal(result: &EvalResult) -> BoundExpr {
    match result {
        EvalResult::Array(a) => {
            let mut args = vec![Expr::Number(a.rows() as f64), Expr::Number(a.cols() as f64)];
            for r in 0..a.rows() {
                for c in 0..a.cols() {
                    args.push(value_literal(a.get(r, c).unwrap_or(&Value::Empty)));
                }
            }
            Expr::Function { name: LITERAL_ARRAY.to_string(), args }
        }
        other => value_literal(&other.to_value()),
    }
}

/// A call of a LAMBDA with these (already substituted) arguments: a LET of its
/// parameters, so each argument is evaluated once.
fn call(params: &[String], body: &BoundExpr, args: Vec<BoundExpr>) -> BoundExpr {
    if params.len() != args.len() {
        return error_literal("#VALUE!");
    }
    if params.is_empty() {
        return body.clone();
    }
    let mut let_args = Vec::with_capacity(params.len() * 2 + 1);
    for (param, arg) in params.iter().zip(args) {
        let_args.push(Expr::NamedRange(param.clone()));
        let_args.push(arg);
    }
    let_args.push(body.clone());
    Expr::Function { name: "LET".to_string(), args: let_args }
}

/// Replace names bound in `env` throughout `expr`, honoring shadowing by
/// inner LETs and LAMBDA parameters, and expanding calls of bound LAMBDAs.
fn substitute(expr: &BoundExpr, env: &Env) -> BoundExpr {
    match expr {
        Expr::NamedRange(name) => match env.get(name) {
            Some(Binding::Expr(e)) => e.clone(),
            Some(Binding::Lambda { params, body }) => {
                let mut args: Vec<BoundExpr> = params.iter().map(|p| Expr::NamedRange(p.clone())).collect();
                args.push(body.clone());
                Expr::Function { name: "LAMBDA".to_string(), args }
            }
            None => expr.clone(),
        },
        Expr::BinaryOp { op, left, right } => Expr::BinaryOp {
            op: *op,
            left: Box::new(substitute(left, env)),
            right: Box::new(substitute(right, env)),
        },
        Expr::Function { name, args } if name == "LET" && args.len() >= 3 && !args.len().is_multiple_of(2) => {
            let mut scope = env.clone();
            let mut out = Vec::with_capacity(args.len());
            let (pairs, body) = args.split_at(args.len() - 1);
            for pair in pairs.chunks(2) {
                out.push(pair[0].clone());
                out.push(substitute(&pair[1], &scope));
                if let Expr::NamedRange(n) = &pair[0] {
                    scope.remove(n); // the inner LET's own binding shadows ours
                }
            }
            out.push(substitute(&body[0], &scope));
            Expr::Function { name: name.clone(), args: out }
        }
        Expr::Function { name, args } if name == "LAMBDA" && !args.is_empty() => {
            let mut scope = env.clone();
            let (body, params) = args.split_last().expect("non-empty");
            for p in params {
                if let Expr::NamedRange(n) = p {
                    scope.remove(n);
                }
            }
            let mut out: Vec<BoundExpr> = params.to_vec();
            out.push(substitute(body, &scope));
            Expr::Function { name: name.clone(), args: out }
        }
        Expr::Function { name, args } if name == INVOKE && !args.is_empty() => {
            let callee = substitute(&args[0], env);
            let call_args: Vec<BoundExpr> = args[1..].iter().map(|a| substitute(a, env)).collect();
            match &callee {
                Expr::Function { name: f, args: lambda } if f == "LAMBDA" => match lambda_parts(lambda) {
                    Some((params, body)) => call(&params, &body, call_args),
                    None => error_literal("#VALUE!"),
                },
                // A callee that is a LET returns its calculation, which may be
                // a LAMBDA: LAMBDA(x, LAMBDA(y, x+y))(1)(2), whose inner call
                // has already become LET(x, 1, LAMBDA(y, x+y)). Call that
                // calculation inside the LET. The arguments are bound first,
                // under names no formula can contain, so the LET's own names
                // cannot capture anything they mention.
                Expr::Function { name: f, args: let_args } if f == "LET" && let_args.len() >= 3 => {
                    let held: Vec<String> = (0..call_args.len()).map(|i| format!("\u{1}ARG{i}")).collect();
                    let mut inner = let_args.clone();
                    let body = inner.pop().expect("LET has a calculation");
                    let mut invoke = vec![body];
                    invoke.extend(held.iter().map(|h| Expr::NamedRange(h.clone())));
                    inner.push(Expr::Function { name: INVOKE.to_string(), args: invoke });
                    let inner = Expr::Function { name: "LET".to_string(), args: inner };
                    if held.is_empty() {
                        return inner;
                    }
                    let mut outer = Vec::with_capacity(held.len() * 2 + 1);
                    for (h, a) in held.into_iter().zip(call_args) {
                        outer.push(Expr::NamedRange(h));
                        outer.push(a);
                    }
                    outer.push(inner);
                    Expr::Function { name: "LET".to_string(), args: outer }
                }
                _ => {
                    let mut out = vec![callee];
                    out.extend(call_args);
                    Expr::Function { name: name.clone(), args: out }
                }
            }
        }
        Expr::Function { name, args } => {
            let call_args: Vec<BoundExpr> = args.iter().map(|a| substitute(a, env)).collect();
            match env.get(name) {
                // A name bound to a LAMBDA, called: f(3).
                Some(Binding::Lambda { params, body }) => call(params, body, call_args),
                // Any other binding does not shadow a built-in function name.
                _ => Expr::Function { name: name.clone(), args: call_args },
            }
        }
        _ => expr.clone(),
    }
}
