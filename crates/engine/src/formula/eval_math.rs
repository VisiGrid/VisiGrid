// Math functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD,
// POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN

use super::eval::{evaluate, CellLookup, EvalResult};
use super::eval_helpers::{collect_numbers, collect_all_values};
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
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
                return Some(EvalResult::Error("ABS requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.abs()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ROUND" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("ROUND requires 1 or 2 arguments".to_string()));
            }
            let value = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let decimals = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0
            };
            let factor = 10_f64.powi(decimals);
            EvalResult::Number((value * factor).round() / factor)
        }
        "INT" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("INT requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.floor()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "MOD" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("MOD requires exactly 2 arguments".to_string()));
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let divisor = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if divisor == 0.0 {
                return Some(EvalResult::Error("#DIV/0!".to_string()));
            }
            EvalResult::Number(number % divisor)
        }
        "POWER" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("POWER requires exactly 2 arguments".to_string()));
            }
            let base = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let exp = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Number(base.powf(exp))
        }
        "SQRT" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("SQRT requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.sqrt()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "CEILING" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("CEILING requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.ceil()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "FLOOR" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("FLOOR requires exactly one argument".to_string()));
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
        _ => return None,
    };
    Some(result)
}
