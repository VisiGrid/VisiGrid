// Formula evaluator - evaluates parsed expressions

use super::parser::{Expr, Op};

pub trait CellLookup {
    fn get_value(&self, row: usize, col: usize) -> f64;
}

#[derive(Debug, Clone)]
pub enum EvalResult {
    Number(f64),
    Error(String),
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
            EvalResult::Error(e) => format!("#ERR: {}", e),
        }
    }
}

pub fn evaluate<L: CellLookup>(expr: &Expr, lookup: &L) -> EvalResult {
    match expr {
        Expr::Number(n) => EvalResult::Number(*n),
        Expr::CellRef { col, row } => {
            let value = lookup.get_value(*row, *col);
            EvalResult::Number(value)
        }
        Expr::Range { .. } => {
            // Ranges can't be evaluated directly, only within functions
            EvalResult::Error("Range must be used in a function".to_string())
        }
        Expr::Function { name, args } => evaluate_function(name, args, lookup),
        Expr::BinaryOp { op, left, right } => {
            let left_val = match evaluate(left, lookup) {
                EvalResult::Number(n) => n,
                err => return err,
            };
            let right_val = match evaluate(right, lookup) {
                EvalResult::Number(n) => n,
                err => return err,
            };

            let result = match op {
                Op::Add => left_val + right_val,
                Op::Sub => left_val - right_val,
                Op::Mul => left_val * right_val,
                Op::Div => {
                    if right_val == 0.0 {
                        return EvalResult::Error("Division by zero".to_string());
                    }
                    left_val / right_val
                }
            };

            EvalResult::Number(result)
        }
    }
}

fn evaluate_function<L: CellLookup>(name: &str, args: &[Expr], lookup: &L) -> EvalResult {
    match name {
        "SUM" => {
            let values = collect_values(args, lookup);
            match values {
                Ok(vals) => EvalResult::Number(vals.iter().sum()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "AVERAGE" | "AVG" => {
            let values = collect_values(args, lookup);
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
            let values = collect_values(args, lookup);
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
            let values = collect_values(args, lookup);
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
            let values = collect_values(args, lookup);
            match values {
                Ok(vals) => EvalResult::Number(vals.len() as f64),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ABS" => {
            if args.len() != 1 {
                return EvalResult::Error("ABS requires exactly one argument".to_string());
            }
            match evaluate(&args[0], lookup) {
                EvalResult::Number(n) => EvalResult::Number(n.abs()),
                err => err,
            }
        }
        "ROUND" => {
            if args.is_empty() || args.len() > 2 {
                return EvalResult::Error("ROUND requires 1 or 2 arguments".to_string());
            }
            let value = match evaluate(&args[0], lookup) {
                EvalResult::Number(n) => n,
                err => return err,
            };
            let decimals = if args.len() == 2 {
                match evaluate(&args[1], lookup) {
                    EvalResult::Number(n) => n as i32,
                    err => return err,
                }
            } else {
                0
            };
            let factor = 10_f64.powi(decimals);
            EvalResult::Number((value * factor).round() / factor)
        }
        _ => EvalResult::Error(format!("Unknown function: {}", name)),
    }
}

fn collect_values<L: CellLookup>(args: &[Expr], lookup: &L) -> Result<Vec<f64>, String> {
    let mut values = Vec::new();

    for arg in args {
        match arg {
            Expr::Range { start_col, start_row, end_col, end_row } => {
                let min_row = (*start_row).min(*end_row);
                let max_row = (*start_row).max(*end_row);
                let min_col = (*start_col).min(*end_col);
                let max_col = (*start_col).max(*end_col);

                for r in min_row..=max_row {
                    for c in min_col..=max_col {
                        values.push(lookup.get_value(r, c));
                    }
                }
            }
            _ => {
                match evaluate(arg, lookup) {
                    EvalResult::Number(n) => values.push(n),
                    EvalResult::Error(e) => return Err(e),
                }
            }
        }
    }

    Ok(values)
}
