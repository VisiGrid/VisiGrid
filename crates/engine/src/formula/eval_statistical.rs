// Statistical functions: STDEV, STDEV.S, STDEV.P, STDEVP, VAR, VAR.S, VAR.P,
// VARP, RAND, RANDBETWEEN

use super::eval::{evaluate, CellLookup, EvalResult};
use super::eval_helpers::collect_numbers;
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "STDEV" | "STDEV.S" => {
            // Sample standard deviation
            let values = collect_numbers(args, lookup);
            match values {
                Ok(vals) => {
                    if vals.len() < 2 {
                        return Some(EvalResult::Error("#DIV/0!".to_string()));
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
                        return Some(EvalResult::Error("#DIV/0!".to_string()));
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
                        return Some(EvalResult::Error("#DIV/0!".to_string()));
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
                        return Some(EvalResult::Error("#DIV/0!".to_string()));
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
        "RAND" => {
            if !args.is_empty() {
                return Some(EvalResult::Error("RAND takes no arguments".to_string()));
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
                return Some(EvalResult::Error("RANDBETWEEN requires exactly 2 arguments".to_string()));
            }
            let bottom = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.floor() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let top = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.floor() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if bottom > top {
                return Some(EvalResult::Error("#NUM!".to_string()));
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
        _ => return None,
    };
    Some(result)
}
