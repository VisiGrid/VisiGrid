// Math functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD,
// POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN, SUMPRODUCT

use super::eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution, Value};
use super::eval_helpers::{
    collect_all_values, collect_numbers, excel_mod, round_to_15_sig, round_to_digits, RoundMode,
};
use super::parser::{BoundExpr, Expr};

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
            EvalResult::Number(round_to_digits(value, decimals, RoundMode::Nearest))
        }
        "ROUNDUP" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("ROUNDUP requires 1 or 2 arguments".to_string()));
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
            EvalResult::Number(round_to_digits(value, decimals, RoundMode::AwayFromZero))
        }
        "ROUNDDOWN" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("ROUNDDOWN requires 1 or 2 arguments".to_string()));
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
            EvalResult::Number(round_to_digits(value, decimals, RoundMode::TowardZero))
        }
        "TRUNC" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("TRUNC requires 1 or 2 arguments".to_string()));
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
            EvalResult::Number(round_to_digits(value, decimals, RoundMode::TowardZero))
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
            EvalResult::Number(excel_mod(number, divisor))
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
            power(base, exp)
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
        "CEILING" | "FLOOR" => {
            // CEILING(number, [significance]) / FLOOR(number, [significance]).
            // Excel requires the significance; it is optional here so sheets written
            // against the old one-argument form (whole-number ceil/floor) still work.
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error(format!("{} requires 1 or 2 arguments", name)));
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let significance = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1.0
            };
            round_to_multiple(number, significance, name == "CEILING")
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
                        vals.sort_by(f64::total_cmp);
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
        "SUMPRODUCT" => {
            // SUMPRODUCT(array1, array2, ..., arrayN) -> Number
            //
            // Contract:
            // - Each arg is a range, cell ref, named range, or an expression that
            //   produces an array (B2:B9*C2:C9, --(A2:A9="x"))
            // - All args must have the same shape (rows × cols)
            // - Iterates row-major, multiplies corresponding cells, sums products
            // - In a range, numeric text counts; empty/other text/bool cells → 0.
            //   In a computed array, only numbers count, which is why booleans
            //   need -- to become 1s and 0s, as in Excel. Errors propagate.
            if args.is_empty() {
                return Some(EvalResult::Error("SUMPRODUCT requires at least one argument".to_string()));
            }
            let mut grids: Vec<Vec<Vec<f64>>> = Vec::with_capacity(args.len());
            for arg in args {
                let grid = match sumproduct_grid(arg, lookup) {
                    Ok(g) => g,
                    Err(e) => return Some(EvalResult::Error(e)),
                };
                grids.push(grid);
            }
            let (rows, cols) = (grids[0].len(), grids[0].first().map_or(0, |r| r.len()));
            for (i, g) in grids.iter().enumerate().skip(1) {
                let (r, c) = (g.len(), g.first().map_or(0, |r| r.len()));
                if r != rows || c != cols {
                    return Some(EvalResult::Error(format!(
                        "SUMPRODUCT ranges must have the same shape. Argument 1 is {}x{}, argument {} is {}x{}.",
                        rows, cols, i + 1, r, c
                    )));
                }
            }
            let mut sum = 0.0;
            for r in 0..rows {
                for c in 0..cols {
                    sum += grids.iter().map(|g| g[r][c]).product::<f64>();
                }
            }
            EvalResult::Number(sum)
        }
        _ => return None,
    };
    Some(result)
}

/// POWER and the ^ operator. Excel never shows NaN or infinity: 0 to a negative power is
/// a division by zero, and anything else without a finite real result is #NUM!.
pub(crate) fn power(base: f64, exp: f64) -> EvalResult {
    if base == 0.0 && exp < 0.0 {
        return EvalResult::Error("#DIV/0!".to_string());
    }
    let result = base.powf(exp);
    if result.is_finite() {
        EvalResult::Number(result)
    } else {
        EvalResult::Error("#NUM!".to_string())
    }
}

/// CEILING and FLOOR with Excel's significance rules.
///
/// The quotient is rounded towards +inf (CEILING) or -inf (FLOOR) and multiplied back,
/// which gives each sign combination Excel's answer: CEILING(-4.2,1) = -4 and
/// CEILING(-4.2,-1) = -5. A positive number with a negative significance is #NUM!.
/// A zero significance gives 0 for CEILING and #DIV/0! for FLOOR, as in Excel.
fn round_to_multiple(number: f64, significance: f64, up: bool) -> EvalResult {
    if number == 0.0 {
        return EvalResult::Number(0.0);
    }
    if significance == 0.0 {
        return if up {
            EvalResult::Number(0.0)
        } else {
            EvalResult::Error("#DIV/0!".to_string())
        };
    }
    if number > 0.0 && significance < 0.0 {
        return EvalResult::Error("#NUM!".to_string());
    }
    // 0.3/0.1 is 2.9999999999999996 in binary; snap before choosing a side.
    let q = round_to_15_sig(number / significance);
    let q = if up { q.ceil() } else { q.floor() };
    // ...and again after, so CEILING(0.25,0.1) is 0.3 rather than 0.30000000000000004.
    EvalResult::Number(round_to_15_sig(q * significance))
}

/// One SUMPRODUCT argument as a grid of numbers.
fn sumproduct_grid<L: CellLookup>(arg: &BoundExpr, lookup: &L) -> Result<Vec<Vec<f64>>, String> {
    // Ranges keep their long-standing reading: numeric text counts, other text,
    // booleans and blanks are 0.
    let rect = match arg {
        Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
            if matches!(sheet, crate::sheet::SheetRef::RefError { .. }) {
                return Err("#REF!".to_string());
            }
            Some((*start_row, *start_col, *end_row, *end_col))
        }
        Expr::CellRef { col, row, .. } => Some((*row, *col, *row, *col)),
        Expr::NamedRange(name) => match lookup.resolve_named_range(name) {
            Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                Some((start_row, start_col, end_row, end_col))
            }
            Some(NamedRangeResolution::Cell { row, col }) => Some((row, col, row, col)),
            None => return Err(format!("#NAME? '{}'", name)),
        },
        _ => None,
    };
    if let Some((r1, c1, r2, c2)) = rect {
        let (r0, r1) = (r1.min(r2), r1.max(r2));
        let (c0, c1) = (c1.min(c2), c1.max(c2));
        return (r0..=r1)
            .map(|r| {
                (c0..=c1)
                    .map(|c| {
                        let text = match arg {
                            Expr::Range { sheet, .. } | Expr::CellRef { sheet, .. } => super::eval_helpers::get_text_for_sheet(lookup, sheet, r, c)?,
                            _ => lookup.get_text(r,c),
                        };
                        if text.starts_with('#') {
                            Err(text)
                        } else {
                            Ok(crate::cell::parse_finite(&text).unwrap_or(0.0))
                        }
                    })
                    .collect()
            })
            .collect();
    }
    let element = |v: &Value| match v {
        Value::Number(n) => Ok(*n),
        Value::Error(e) => Err(e.clone()),
        _ => Ok(0.0),
    };
    match evaluate(arg, lookup) {
        EvalResult::Array(a) => (0..a.rows())
            .map(|r| (0..a.cols()).map(|c| a.get(r, c).map_or(Ok(0.0), element)).collect())
            .collect(),
        EvalResult::Error(e) => Err(e),
        // A lone scalar is a 1x1 array: SUMPRODUCT(5) is 5.
        other => Ok(vec![vec![other.to_number()?]]),
    }
}
