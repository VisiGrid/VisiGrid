// Math functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD,
// POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN, SUMPRODUCT

use super::eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution};
use super::eval_helpers::{collect_numbers, collect_all_values};
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
            let factor = 10_f64.powi(decimals);
            EvalResult::Number((value * factor).round() / factor)
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
            let factor = 10_f64.powi(decimals);
            let result = if value >= 0.0 {
                (value * factor).ceil() / factor
            } else {
                (value * factor).floor() / factor
            };
            EvalResult::Number(result)
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
            let factor = 10_f64.powi(decimals);
            let result = if value >= 0.0 {
                (value * factor).floor() / factor
            } else {
                (value * factor).ceil() / factor
            };
            EvalResult::Number(result)
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
            let factor = 10_f64.powi(decimals);
            let result = if value >= 0.0 {
                (value * factor).floor() / factor
            } else {
                (value * factor).ceil() / factor
            };
            EvalResult::Number(result)
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
