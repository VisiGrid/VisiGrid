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
        "NORMSDIST" => {
            // Standard normal cumulative distribution function
            if args.len() != 1 {
                return Some(EvalResult::Error("NORMSDIST requires 1 argument".to_string()));
            }
            let z = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if z.is_nan() || z.is_infinite() {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            EvalResult::Number(norm_s_cdf(z))
        }
        "NORM.S.DIST" => {
            // Standard normal distribution (CDF or PDF)
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("NORM.S.DIST requires 1-2 arguments".to_string()));
            }
            let z = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if z.is_nan() || z.is_infinite() {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let cumulative = if args.len() > 1 {
                match evaluate(&args[1], lookup).to_bool() {
                    Ok(b) => b,
                    Err(_) => return Some(EvalResult::Error("#VALUE!".to_string())),
                }
            } else {
                true
            };
            if cumulative {
                EvalResult::Number(norm_s_cdf(z))
            } else {
                // PDF: (1/sqrt(2*pi)) * exp(-z^2/2)
                EvalResult::Number((1.0 / (2.0 * std::f64::consts::PI).sqrt()) * (-z * z / 2.0).exp())
            }
        }
        _ => return None,
    };
    Some(result)
}

/// Error function approximation (Abramowitz & Stegun 7.1.26, max error ~1.5e-7)
fn erf(x: f64) -> f64 {
    let a1 =  0.254829592;
    let a2 = -0.284496736;
    let a3 =  1.421413741;
    let a4 = -1.453152027;
    let a5 =  1.061405429;
    let p  =  0.3275911;
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let x = x.abs();
    let t = 1.0 / (1.0 + p * x);
    let y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * (-x * x).exp();
    sign * y
}

/// Standard normal CDF: Φ(z) = 0.5 * (1 + erf(z / sqrt(2)))
fn norm_s_cdf(z: f64) -> f64 {
    0.5 * (1.0 + erf(z / std::f64::consts::SQRT_2))
}
