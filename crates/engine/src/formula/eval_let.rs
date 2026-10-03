// LET and LAMBDA.
//
// LET(name1, value1, [name2, value2, ...], calculation) names intermediate
// results; LAMBDA(param1, ..., calculation) is a function written in a
// formula, called where it is written — LAMBDA(x, x+1)(2) — or through a LET
// name — LET(f, LAMBDA(x, x*2), f(3)).
//
// Both work by substitution at evaluation time; the stored formula is never
// rewritten. When a LET is evaluated, each value is evaluated once, in order,
// and every later use of its name is replaced by that result (a literal, or a
// handle into a per-thread value table for errors, blanks and arrays). A
// value that is a reference (A1, A1:B9, a defined name, a table column) is
// substituted as the reference itself, so SUM(x) and COUNTIF(x, ...) keep
// range semantics. A LAMBDA bound to a name is kept as a LAMBDA; calling it
// becomes a LET of its parameters, which is exactly the meaning of a call.
//
// Names are lexically scoped: an inner LET or a LAMBDA parameter of the same
// name shadows an outer one. A call's arguments are evaluated in the caller's
// scope, and a defined name bound to a LET name is pinned to its cells, so a
// parameter can never capture a name the caller meant.
//
// Substitution is bounded (MAX_EXPANSION_NODES): chained LAMBDAs can grow
// exponentially, and a formula that would expand past the bound is #CALC!.
//
// Not supported: a LAMBDA passed as a value to MAP/REDUCE/BYROW and the other
// helper functions (which the engine does not implement), recursion (a LAMBDA
// cannot call the LET name it is bound to), and LAMBDAs defined in the Name
// Manager. A LAMBDA that is never called evaluates to #CALC!, as in Excel.

use std::collections::HashMap;
use std::rc::Rc;

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, NamedRangeResolution};
use crate::sheet::SheetRef;
use super::parser::{BoundExpr, Expr, ARRAY_LITERAL, INVOKE};

/// Internal literal for a value with no expression of its own: an error, an
/// empty value, or an array. Never written by a user (underscore prefix).
pub(super) const LITERAL_ERROR: &str = "_LET.ERR";
pub(super) const LITERAL_ARRAY: &str = ARRAY_LITERAL;
/// A handle to a value LET has already computed: `_LET.VALUE(index)` into
/// this thread's value table. Two nodes however large the value, so binding
/// SEQUENCE(100000) costs nothing to substitute.
pub(super) const LITERAL_VALUE: &str = "_LET.VALUE";

thread_local! {
    /// Values bound by the LETs being evaluated on this thread. A LET pushes
    /// its values and truncates back when its calculation is done, so a
    /// handle is valid exactly while the expression holding it is evaluated.
    static VALUES: std::cell::RefCell<Vec<Rc<EvalResult>>> = const { std::cell::RefCell::new(Vec::new()) };
}

/// Shared, so cloning a scope (once per nested LET or LAMBDA) never copies
/// the expressions bound in it.
#[derive(Clone)]
enum Binding {
    /// A literal or a reference, substituted as an expression.
    Expr(Rc<BoundExpr>),
    /// A LAMBDA, already closed over the names in scope where it was defined.
    Lambda { params: Rc<[String]>, body: Rc<BoundExpr> },
}

type Env = HashMap<String, Binding>;

/// Evaluate the functions this module owns, or None for any other name.
pub(super) fn try_evaluate<L: CellLookup>(name: &str, args: &[BoundExpr], lookup: &L) -> Option<EvalResult> {
    Some(match name {
        "LET" => evaluate_let(args, lookup),
        // A LAMBDA reached by evaluation was never called.
        "LAMBDA" => EvalResult::Error("#CALC!".to_string()),
        _ if name == INVOKE => {
            let call = Expr::Function { name: INVOKE.to_string(), args: args.to_vec() };
            let Some(expanded) = expand(&call, &Env::new()) else {
                return Some(too_large());
            };
            if matches!(&expanded, Expr::Function { name, .. } if name == INVOKE) {
                // The callee is not a LAMBDA, so there is nothing to call.
                EvalResult::Error("#VALUE!".to_string())
            } else {
                evaluate(&expanded, lookup)
            }
        }
        LITERAL_VALUE => match args.first() {
            Some(Expr::Number(i)) => VALUES
                .with(|v| v.borrow().get(*i as usize).map(|r| (**r).clone()))
                .unwrap_or_else(|| EvalResult::Error("#VALUE!".to_string())),
            _ => EvalResult::Error("#VALUE!".to_string()),
        },
        LITERAL_ERROR => match args.first() {
            Some(Expr::Text(e)) => EvalResult::Error(e.clone()),
            _ => EvalResult::Error("#VALUE!".to_string()),
        },
        LITERAL_ARRAY => array_literal(args, lookup),
        _ => return None,
    })
}

fn evaluate_let<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> EvalResult {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return EvalResult::Error("#VALUE! LET needs name, value pairs and a calculation".to_string());
    }
    let mark = VALUES.with(|v| v.borrow().len());
    let result = evaluate_let_inner(args, lookup);
    VALUES.with(|v| v.borrow_mut().truncate(mark));
    result
}

fn evaluate_let_inner<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> EvalResult {
    let mut env = Env::new();
    let (pairs, body) = args.split_at(args.len() - 1);
    for pair in pairs.chunks(2) {
        let Expr::NamedRange(name) = &pair[0] else {
            return EvalResult::Error("#VALUE! LET names must be names".to_string());
        };
        let Some(value) = expand(&pair[1], &env) else {
            return too_large();
        };
        let binding = match value {
            Expr::Function { name: ref f, ref args } if f == "LAMBDA" => match lambda_parts(args) {
                Some((params, body)) => Binding::Lambda { params: params.into(), body: Rc::new(body) },
                None => Binding::Expr(Rc::new(error_literal("#VALUE!"))),
            },
            // A defined name is pinned to the cells it names now, so a LET or
            // LAMBDA parameter of the same name further in cannot capture it.
            Expr::NamedRange(ref n) if !lookup.is_table_name(n) => match lookup.resolve_named_range(n) {
                Some(NamedRangeResolution::Cell { row, col }) => Binding::Expr(Rc::new(Expr::CellRef {
                    sheet: SheetRef::Current, col, row, col_abs: true, row_abs: true,
                })),
                Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => Binding::Expr(Rc::new(Expr::Range {
                    sheet: SheetRef::Current, start_col, start_row, end_col, end_row,
                    start_col_abs: true, start_row_abs: true, end_col_abs: true, end_row_abs: true,
                })),
                None => Binding::Expr(Rc::new(literal(evaluate(&value, lookup)))),
            },
            ref v if is_reference(v) => Binding::Expr(Rc::new(value)),
            v => Binding::Expr(Rc::new(literal(evaluate(&v, lookup)))),
        };
        env.insert(name.clone(), binding);
    }
    match expand(&body[0], &env) {
        Some(body) => evaluate(&body, lookup),
        None => too_large(),
    }
}

/// The largest expression one expansion may build. Chained LAMBDAs grow by
/// substitution — LET(a, LAMBDA(x,x+x), b, LAMBDA(x,a(a(x))), ...) doubles per
/// level — so thirty levels of an ~800-character formula would build a
/// billion nodes. Expansion counts the nodes it builds and gives up here.
/// Bound values are handles (see LITERAL_VALUE), so this counts formula
/// structure only: a formula of Excel's maximum length parses to far fewer.
const MAX_EXPANSION_NODES: usize = 20_000;

fn too_large() -> EvalResult {
    EvalResult::Error("#CALC! Formula expands too far".to_string())
}

/// `substitute` within the node budget; None when it would exceed it.
fn expand(expr: &BoundExpr, env: &Env) -> Option<BoundExpr> {
    let mut budget = MAX_EXPANSION_NODES;
    substitute(expr, &mut Scope { env, shadowed: Vec::new() }, &mut budget)
}

/// The names in scope during a substitution: `env`, minus any an inner LET
/// or LAMBDA parameter has rebound. Shadowing is a stack pushed and popped
/// around each inner scope, so entering one never copies the environment.
struct Scope<'a> {
    env: &'a Env,
    shadowed: Vec<String>,
}

impl Scope<'_> {
    fn get(&self, name: &str) -> Option<&Binding> {
        if self.shadowed.iter().any(|s| s == name) {
            return None;
        }
        self.env.get(name)
    }
}

/// Charge `n` nodes to the budget; None once it is spent.
fn charge(budget: &mut usize, n: usize) -> Option<()> {
    *budget = budget.checked_sub(n)?;
    Some(())
}

/// Nodes in an expression, stopping early once past `limit`.
fn size_within(expr: &BoundExpr, limit: usize) -> usize {
    fn walk(e: &BoundExpr, n: &mut usize, limit: usize) {
        if *n > limit {
            return;
        }
        *n += 1;
        match e {
            Expr::Function { args, .. } => args.iter().for_each(|a| walk(a, n, limit)),
            Expr::BinaryOp { left, right, .. } => {
                walk(left, n, limit);
                walk(right, n, limit);
            }
            _ => {}
        }
    }
    let mut n = 0;
    walk(expr, &mut n, limit);
    n
}

/// A clone that is charged for its size.
fn charged_clone(expr: &BoundExpr, budget: &mut usize) -> Option<BoundExpr> {
    charge(budget, size_within(expr, *budget))?;
    Some(expr.clone())
}

/// An `_ARRAY(rows, cols, elements...)` value. Its shape is checked against
/// its own arguments before anything is allocated: rows and cols must be
/// positive whole numbers within the sheet, and exactly rows*cols elements
/// must follow, so an array can never be larger than the formula holding it.
fn array_literal<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> EvalResult {
    let dim = |i: usize, max: usize| match args.get(i) {
        Some(Expr::Number(n)) if n.fract() == 0.0 && *n >= 1.0 && *n <= max as f64 => Some(*n as usize),
        _ => None,
    };
    let (Some(rows), Some(cols)) = (dim(0, crate::sheet::NUM_ROWS), dim(1, crate::sheet::NUM_COLS)) else {
        return EvalResult::Error("#VALUE!".to_string());
    };
    if rows.checked_mul(cols).and_then(|n| n.checked_add(2)) != Some(args.len()) {
        return EvalResult::Error("#VALUE!".to_string());
    }
    let mut out = Array2D::new(rows, cols);
    for (i, element) in args[2..].iter().enumerate() {
        out.set(i / cols, i % cols, evaluate(element, lookup).to_value());
    }
    EvalResult::Array(out)
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

/// An evaluated result as an expression that evaluates back to it: a plain
/// literal for a number, text or logical, otherwise a handle into the value
/// table (errors, empty, arrays).
fn literal(result: EvalResult) -> BoundExpr {
    match result {
        EvalResult::Number(n) => Expr::Number(n),
        EvalResult::Text(s) => Expr::Text(s),
        EvalResult::Boolean(b) => Expr::Boolean(b),
        other => {
            let index = VALUES.with(|v| {
                let mut v = v.borrow_mut();
                v.push(Rc::new(other));
                v.len() - 1
            });
            Expr::Function { name: LITERAL_VALUE.to_string(), args: vec![Expr::Number(index as f64)] }
        }
    }
}

/// A name no formula can contain (formula names are ASCII letters, digits,
/// `_`, `.` and `\\`), unique within this thread.
fn hidden_name() -> String {
    thread_local!(static NEXT: std::cell::Cell<u64> = const { std::cell::Cell::new(0) });
    NEXT.with(|n| {
        n.set(n.get() + 1);
        format!("\u{1}ARG{}", n.get())
    })
}

/// LET(h1, a1, h2, a2, ..., inner): each argument evaluated once, in the
/// caller's scope, before `inner` (which refers to them as h1, h2, ...) binds
/// anything of its own. Binding parameters directly as LET(x, a1, y, a2, ...)
/// would evaluate a2 with x already bound to a1, so f(1, x) with a defined
/// name x would read the parameter rather than the name.
fn with_args(args: Vec<BoundExpr>, inner: impl FnOnce(&[String]) -> BoundExpr) -> BoundExpr {
    let held: Vec<String> = args.iter().map(|_| hidden_name()).collect();
    let inner = inner(&held);
    if held.is_empty() {
        return inner;
    }
    let mut out = Vec::with_capacity(held.len() * 2 + 1);
    for (h, a) in held.into_iter().zip(args) {
        out.push(Expr::NamedRange(h));
        out.push(a);
    }
    out.push(inner);
    Expr::Function { name: "LET".to_string(), args: out }
}

/// A call of a LAMBDA with these (already substituted) arguments: a LET of its
/// parameters, so each argument is evaluated once.
fn call(params: &[String], body: &BoundExpr, args: Vec<BoundExpr>, budget: &mut usize) -> Option<BoundExpr> {
    if params.len() != args.len() {
        return Some(error_literal("#VALUE!"));
    }
    let body = charged_clone(body, budget)?;
    charge(budget, params.len() * 3 + 2)?;
    if params.is_empty() {
        return Some(body);
    }
    Some(with_args(args, |held| {
        let mut let_args = Vec::with_capacity(params.len() * 2 + 1);
        for (param, h) in params.iter().zip(held) {
            let_args.push(Expr::NamedRange(param.clone()));
            let_args.push(Expr::NamedRange(h.clone()));
        }
        let_args.push(body);
        Expr::Function { name: "LET".to_string(), args: let_args }
    }))
}

/// Replace names bound in `env` throughout `expr`, honoring shadowing by
/// inner LETs and LAMBDA parameters, and expanding calls of bound LAMBDAs.
/// Every node built is charged to `budget`; None once it is spent.
fn substitute(expr: &BoundExpr, env: &mut Scope, budget: &mut usize) -> Option<BoundExpr> {
    charge(budget, 1)?;
    Some(match expr {
        Expr::NamedRange(name) => match env.get(name) {
            Some(Binding::Expr(e)) => charged_clone(e, budget)?,
            Some(Binding::Lambda { params, body }) => {
                let mut args: Vec<BoundExpr> = params.iter().map(|p| Expr::NamedRange(p.clone())).collect();
                args.push(charged_clone(body, budget)?);
                Expr::Function { name: "LAMBDA".to_string(), args }
            }
            None => expr.clone(),
        },
        Expr::BinaryOp { op, left, right } => Expr::BinaryOp {
            op: *op,
            left: Box::new(substitute(left, env, budget)?),
            right: Box::new(substitute(right, env, budget)?),
        },
        Expr::Function { name, args } if name == "LET" && args.len() >= 3 && !args.len().is_multiple_of(2) => {
            let depth = env.shadowed.len();
            let mut out = Vec::with_capacity(args.len());
            let (pairs, body) = args.split_at(args.len() - 1);
            for pair in pairs.chunks(2) {
                out.push(pair[0].clone());
                let value = substitute(&pair[1], env, budget);
                let Some(value) = value else {
                    env.shadowed.truncate(depth);
                    return None;
                };
                out.push(value);
                if let Expr::NamedRange(n) = &pair[0] {
                    env.shadowed.push(n.clone()); // the inner LET's own binding shadows ours
                }
            }
            let calc = substitute(&body[0], env, budget);
            env.shadowed.truncate(depth);
            out.push(calc?);
            Expr::Function { name: name.clone(), args: out }
        }
        Expr::Function { name, args } if name == "LAMBDA" && !args.is_empty() => {
            let depth = env.shadowed.len();
            let (body, params) = args.split_last().expect("non-empty");
            for p in params {
                if let Expr::NamedRange(n) = p {
                    env.shadowed.push(n.clone());
                }
            }
            let calc = substitute(body, env, budget);
            env.shadowed.truncate(depth);
            let mut out: Vec<BoundExpr> = params.to_vec();
            out.push(calc?);
            Expr::Function { name: name.clone(), args: out }
        }
        Expr::Function { name, args } if name == INVOKE && !args.is_empty() => {
            let callee = substitute(&args[0], env, budget)?;
            let call_args = args[1..].iter().map(|a| substitute(a, env, budget)).collect::<Option<Vec<_>>>()?;
            match &callee {
                Expr::Function { name: f, args: lambda } if f == "LAMBDA" => match lambda_parts(lambda) {
                    Some((params, body)) => call(&params, &body, call_args, budget)?,
                    None => error_literal("#VALUE!"),
                },
                // A callee that is a LET returns its calculation, which may be
                // a LAMBDA: LAMBDA(x, LAMBDA(y, x+y))(1)(2), whose inner call
                // has already become a LET ending in LAMBDA(y, x+y). Call that
                // calculation inside the LET, with the arguments bound first in
                // the caller's scope so the LET's names cannot capture them.
                Expr::Function { name: f, args: let_args } if f == "LET" && let_args.len() >= 3 => {
                    charge(budget, call_args.len() * 3 + 3)?;
                    let mut inner = let_args.clone();
                    let body = inner.pop().expect("LET has a calculation");
                    with_args(call_args, |held| {
                        let mut invoke = vec![body];
                        invoke.extend(held.iter().map(|h| Expr::NamedRange(h.clone())));
                        inner.push(Expr::Function { name: INVOKE.to_string(), args: invoke });
                        Expr::Function { name: "LET".to_string(), args: inner }
                    })
                }
                _ => {
                    let mut out = vec![callee];
                    out.extend(call_args);
                    Expr::Function { name: name.clone(), args: out }
                }
            }
        }
        Expr::Function { name, args } => {
            let call_args = args.iter().map(|a| substitute(a, env, budget)).collect::<Option<Vec<_>>>()?;
            match env.get(name) {
                // A name bound to a LAMBDA, called: f(3).
                Some(Binding::Lambda { params, body }) => call(params, body, call_args, budget)?,
                // Any other binding does not shadow a built-in function name.
                _ => Expr::Function { name: name.clone(), args: call_args },
            }
        }
        _ => expr.clone(),
    })
}
