// Statistical functions: STDEV, STDEV.S, STDEV.P, STDEVP, VAR, VAR.S, VAR.P,
// VARP, RAND, RANDBETWEEN, LARGE, SMALL, RANK, RANK.EQ, MODE, MODE.SNGL,
// PERCENTILE, PERCENTILE.INC, QUARTILE, QUARTILE.INC

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, Value};
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
                    let mean = crate::numeric::sum(&vals) / vals.len() as f64;
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
                    let mean = crate::numeric::sum(&vals) / vals.len() as f64;
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
                    let mean = crate::numeric::sum(&vals) / vals.len() as f64;
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
                    let mean = crate::numeric::sum(&vals) / vals.len() as f64;
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
            EvalResult::Number(super::eval_helpers::next_random_f64())
        }
        "RANDBETWEEN" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("RANDBETWEEN requires exactly 2 arguments".to_string()));
            }
            // Excel rounds the bounds inwards: bottom up, top down. Flooring both
            // let RANDBETWEEN(1.5,2.5) return 1, below its own bottom.
            let bottom = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.ceil() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let top = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.floor() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if bottom > top {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let range = (top as i128 - bottom as i128 + 1) as u128;
            // Multiply-shift maps 64 random bits onto the range without modulo bias.
            let offset = (super::eval_helpers::next_random_u64() as u128 * range) >> 64;
            EvalResult::Number((bottom as i128 + offset as i128) as f64)
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
        "LARGE" | "SMALL" => {
            // LARGE(array, k) / SMALL(array, k): the k-th largest or smallest
            // number. Text, logicals and blanks in the array are skipped; an
            // error in it is the result. A fractional k rounds up, as in Excel;
            // k below 1 or past the count is #NUM!. An array of k spills one
            // answer per element.
            if args.len() != 2 {
                return Some(EvalResult::Error(format!("{name} requires exactly 2 arguments")));
            }
            let mut nums = match numbers_strict(&args[0..1], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            nums.sort_by(f64::total_cmp);
            if name == "LARGE" {
                nums.reverse();
            }
            let pick = |k: &EvalResult| -> EvalResult {
                let k = match k.to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(excel_error(e)),
                };
                let k = approx_ceil(k);
                if k < 1.0 || k > nums.len() as f64 {
                    return EvalResult::Error("#NUM!".to_string());
                }
                EvalResult::Number(nums[k as usize - 1])
            };
            per_element(evaluate(&args[1], lookup), pick)
        }
        "RANK" | "RANK.EQ" => {
            // RANK(number, ref, [order]): 1 + how many numbers in ref beat it —
            // larger for order 0 (the default), smaller otherwise. Ties share a
            // rank. A number that is not in ref is #N/A.
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let ascending = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n != 0.0,
                    Err(e) => return Some(EvalResult::Error(excel_error(e))),
                }
            } else {
                false
            };
            let nums = match numbers_strict(&args[1..2], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if !nums.contains(&number) {
                return Some(EvalResult::Error("#N/A".to_string()));
            }
            let beaten = nums.iter().filter(|v| if ascending { **v < number } else { **v > number }).count();
            EvalResult::Number((beaten + 1) as f64)
        }
        "MODE" | "MODE.SNGL" => {
            // MODE(number1, ...): the most frequent number. A tie goes to the
            // value that appears first. No repeated value is #N/A.
            if args.is_empty() {
                return Some(EvalResult::Error(format!("{name} requires at least one argument")));
            }
            let nums = match numbers_strict(args, lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let mut best: Option<(f64, usize)> = None;
            for (i, v) in nums.iter().enumerate() {
                if nums[..i].contains(v) {
                    continue; // counted at its first appearance
                }
                let count = nums[i..].iter().filter(|w| *w == v).count();
                if count > 1 && best.is_none_or(|(_, c)| count > c) {
                    best = Some((*v, count));
                }
            }
            match best {
                Some((v, _)) => EvalResult::Number(v),
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "PERCENTILE" | "PERCENTILE.INC" | "QUARTILE" | "QUARTILE.INC" => {
            // PERCENTILE(array, k): the k-th percentile, 0 <= k <= 1, by linear
            // interpolation between the closest ranks (Excel's inclusive
            // method). QUARTILE(array, quart) is PERCENTILE(array, quart/4) for
            // quart 0..4 (truncated). Out of range, or no numbers, is #NUM!.
            if args.len() != 2 {
                return Some(EvalResult::Error(format!("{name} requires exactly 2 arguments")));
            }
            let mut nums = match numbers_strict(&args[0..1], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if nums.is_empty() {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            nums.sort_by(f64::total_cmp);
            let quartile = name.starts_with("QUARTILE");
            let at = |k: &EvalResult| -> EvalResult {
                let k = match k.to_number() {
                    Ok(n) => n,
                    Err(e) => return EvalResult::Error(excel_error(e)),
                };
                let p = if quartile {
                    let q = k.trunc();
                    if !(0.0..=4.0).contains(&q) {
                        return EvalResult::Error("#NUM!".to_string());
                    }
                    q / 4.0
                } else {
                    if !(0.0..=1.0).contains(&k) {
                        return EvalResult::Error("#NUM!".to_string());
                    }
                    k
                };
                // Snap away binary noise: 0.3 * 3 is 0.8999999999999999, which
                // would make PERCENTILE({1,2,3,4}, 0.3) 1.8999999999999999.
                let pos = ((p * (nums.len() - 1) as f64) * 1e12).round() / 1e12;
                let lo = pos.floor() as usize;
                let frac = pos - lo as f64;
                let value = match nums.get(lo + 1) {
                    Some(hi) if frac > 0.0 => nums[lo] + frac * (hi - nums[lo]),
                    _ => nums[lo],
                };
                EvalResult::Number(value)
            };
            per_element(evaluate(&args[1], lookup), at)
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

/// Round up, treating a value within floating-point noise of an integer as
/// that integer: 2.0000000000000004 is 2, not 3.
fn approx_ceil(x: f64) -> f64 {
    let r = x.round();
    if (x - r).abs() < 1e-9 { r } else { x.ceil() }
}

/// Apply `f` to a scalar argument, or to each element of an array argument
/// (returning an array of answers), so LARGE(A1:A9, {1,2,3}) spills.
fn per_element(arg: EvalResult, f: impl Fn(&EvalResult) -> EvalResult) -> EvalResult {
    match arg {
        EvalResult::Array(a) if a.rows() * a.cols() > 1 => {
            let mut out = Array2D::new(a.rows(), a.cols());
            for r in 0..a.rows() {
                for c in 0..a.cols() {
                    let v = a.get(r, c).cloned().unwrap_or(Value::Empty);
                    out.set(r, c, f(&EvalResult::from_value(&v)).to_value());
                }
            }
            EvalResult::Array(out)
        }
        other => f(&other),
    }
}

/// An evaluation error as Excel reports it. Engine conversion messages
/// ("Cannot convert 'x' to number") become #VALUE!; error values pass through.
fn excel_error(e: String) -> String {
    if e.starts_with('#') { e } else { "#VALUE!".to_string() }
}

/// The numbers the new statistical functions read, Excel's way. In a
/// reference or array, text, logicals and blanks are skipped and an error is
/// the result (`collect_numbers`, which older functions share, skips errors
/// too, so SUM(1, #DIV/0!) is 1 where Excel says #DIV/0!). A value given
/// directly is coerced: a logical is 1 or 0, numeric text is its number, other
/// text is #VALUE!.
fn numbers_strict<L: CellLookup>(args: &[super::parser::BoundExpr], lookup: &L) -> Result<Vec<f64>, String> {
    let mut out = Vec::new();
    for arg in args {
        let got = super::eval_helpers::arg_values(arg, lookup)?;
        for value in got.values {
            match value {
                Value::Error(e) => return Err(e),
                Value::Number(n) => out.push(n),
                _ if got.bulk => {}
                Value::Boolean(b) => out.push(if b { 1.0 } else { 0.0 }),
                Value::Empty => out.push(0.0),
                Value::Text(t) => match crate::cell::parse_finite(&t) {
                    Some(n) => out.push(n),
                    None => return Err("#VALUE!".to_string()),
                },
            }
        }
    }
    Ok(out)
}
