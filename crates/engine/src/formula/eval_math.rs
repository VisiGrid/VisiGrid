// Math functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD,
// POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN, SUMPRODUCT

use super::eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution};
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
            // SUMPRODUCT(range1, range2, ..., rangeN) -> Number
            //
            // Contract:
            // - Each arg must be a range, cell ref, or named range
            // - All args must have the same shape (rows × cols)
            // - Iterates row-major, multiplies corresponding cells, sums products
            // - Empty/text/bool cells → 0, errors propagate
            if args.is_empty() {
                return Some(EvalResult::Error("SUMPRODUCT requires at least one argument".to_string()));
            }

            // Extract rectangular coordinates for each arg
            let mut ranges: Vec<(usize, usize, usize, usize)> = Vec::with_capacity(args.len());
            for (i, arg) in args.iter().enumerate() {
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        ranges.push((*start_row, *start_col, *end_row, *end_col));
                    }
                    Expr::CellRef { col, row, .. } => {
                        ranges.push((*row, *col, *row, *col));
                    }
                    Expr::NamedRange(name) => {
                        match lookup.resolve_named_range(name) {
                            Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                                ranges.push((start_row, start_col, end_row, end_col));
                            }
                            Some(NamedRangeResolution::Cell { row, col }) => {
                                ranges.push((row, col, row, col));
                            }
                            None => return Some(EvalResult::Error(format!("#NAME? '{}'", name))),
                        }
                    }
                    _ => {
                        // Single scalar arg: evaluate and return (SUMPRODUCT(5) = 5)
                        if args.len() == 1 {
                            return Some(match evaluate(arg, lookup).to_number() {
                                Ok(n) => EvalResult::Number(n),
                                Err(e) => EvalResult::Error(e),
                            });
                        }
                        return Some(EvalResult::Error(format!(
                            "SUMPRODUCT argument {} must be a range, cell reference, or named range",
                            i + 1
                        )));
                    }
                }
            }

            // Normalize coordinates (min/max) and compute shape
            let norm: Vec<(usize, usize, usize, usize)> = ranges.iter().map(|&(r1, c1, r2, c2)| {
                (r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2))
            }).collect();

            let num_rows = norm[0].2 - norm[0].0 + 1;
            let num_cols = norm[0].3 - norm[0].1 + 1;

            // Validate all shapes match
            for (i, r) in norm.iter().enumerate().skip(1) {
                let rows = r.2 - r.0 + 1;
                let cols = r.3 - r.1 + 1;
                if rows != num_rows || cols != num_cols {
                    return Some(EvalResult::Error(format!(
                        "SUMPRODUCT ranges must have the same shape. Argument 1 is {}x{}, argument {} is {}x{}.",
                        num_rows, num_cols, i + 1, rows, cols
                    )));
                }
            }

            // Iterate row-major, multiply corresponding cells, accumulate sum
            let mut sum = 0.0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    let mut product = 1.0;
                    for r in &norm {
                        let cell_r = r.0 + row_offset;
                        let cell_c = r.1 + col_offset;
                        let text = lookup.get_text(cell_r, cell_c);
                        if text.is_empty() {
                            product = 0.0;
                            break; // 0 * anything = 0, skip remaining
                        } else if text.starts_with('#') {
                            // Error cell — propagate
                            return Some(EvalResult::Error(text));
                        } else if let Ok(n) = text.parse::<f64>() {
                            product *= n;
                        } else {
                            // Text/bool → 0
                            product = 0.0;
                            break;
                        }
                    }
                    sum += product;
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
