// Conditional aggregate functions: SUMIF, AVERAGEIF, COUNTIF, COUNTBLANK,
// SUMIFS, AVERAGEIFS, COUNTIFS

use super::eval::{evaluate, CellLookup, EvalResult};
use super::eval_helpers::matches_criteria;
use super::parser::{BoundExpr, Expr};

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "SUMIF" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("SUMIF requires 2 or 3 arguments".to_string()));
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("SUMIF requires a range as first argument".to_string())),
            };
            let criteria = evaluate(&args[1], lookup);
            let sum_range = if args.len() == 3 {
                match &args[2] {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => Some((*start_row, *start_col, *end_row, *end_col)),
                    _ => return Some(EvalResult::Error("SUMIF sum_range must be a range".to_string())),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        // Get value from sum_range or criteria range
                        let (sum_r, sum_c) = if let Some((sr, sc, _, _)) = sum_range {
                            (sr + row_offset, sc + col_offset)
                        } else {
                            (r, c)
                        };
                        sum += lookup.get_value(sum_r, sum_c);
                    }
                }
            }
            EvalResult::Number(sum)
        }
        "AVERAGEIF" => {
            // AVERAGEIF(range, criteria, [average_range])
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("AVERAGEIF requires 2 or 3 arguments".to_string()));
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("AVERAGEIF requires a range as first argument".to_string())),
            };
            let criteria = evaluate(&args[1], lookup);
            let avg_range = if args.len() == 3 {
                match &args[2] {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => Some((*start_row, *start_col, *end_row, *end_col)),
                    _ => return Some(EvalResult::Error("AVERAGEIF average_range must be a range".to_string())),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Number(0.0)
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        // Get value from average_range or criteria range
                        let (avg_r, avg_c) = if let Some((ar, ac, _, _)) = avg_range {
                            (ar + row_offset, ac + col_offset)
                        } else {
                            (r, c)
                        };
                        let val = lookup.get_value(avg_r, avg_c);
                        // Only count numeric values for average
                        if val.is_finite() {
                            sum += val;
                            count += 1;
                        }
                    }
                }
            }
            if count == 0 {
                EvalResult::Error("#DIV/0!".to_string())
            } else {
                EvalResult::Number(sum / count as f64)
            }
        }
        "COUNTIF" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("COUNTIF requires exactly 2 arguments".to_string()));
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("COUNTIF requires a range as first argument".to_string())),
            };
            let criteria = evaluate(&args[1], lookup);

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    let cell_text = lookup.get_text(r, c);
                    let cell_value = if cell_text.is_empty() {
                        EvalResult::Text(String::new())
                    } else if let Ok(n) = cell_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(cell_text)
                    };

                    if matches_criteria(&cell_value, &criteria) {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        "COUNTBLANK" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("COUNTBLANK requires exactly one argument".to_string()));
            }
            let range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("COUNTBLANK requires a range".to_string())),
            };

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if lookup.get_text(r, c).is_empty() {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        "SUMIFS" => {
            // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 3 || (args.len() - 1) % 2 != 0 {
                return Some(EvalResult::Error("SUMIFS requires sum_range and pairs of criteria_range and criteria".to_string()));
            }

            let sum_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("SUMIFS sum_range must be a range".to_string())),
            };

            let (sr_min_row, sr_min_col, sr_max_row, sr_max_col) = (
                sum_range.0.min(sum_range.2), sum_range.1.min(sum_range.3),
                sum_range.0.max(sum_range.2), sum_range.1.max(sum_range.3)
            );
            let num_rows = sr_max_row - sr_min_row + 1;
            let num_cols = sr_max_col - sr_min_col + 1;

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return Some(EvalResult::Error("SUMIFS criteria_range must be a range".to_string())),
                };

                // Verify dimensions match sum_range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return Some(EvalResult::Error("SUMIFS criteria ranges must have same dimensions as sum_range".to_string()));
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Number(0.0)
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        sum += lookup.get_value(sr_min_row + row_offset, sr_min_col + col_offset);
                    }
                }
            }
            EvalResult::Number(sum)
        }
        "AVERAGEIFS" => {
            // AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 3 || (args.len() - 1) % 2 != 0 {
                return Some(EvalResult::Error("AVERAGEIFS requires average_range and pairs of criteria_range and criteria".to_string()));
            }

            let avg_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("AVERAGEIFS average_range must be a range".to_string())),
            };

            let (ar_min_row, ar_min_col, ar_max_row, ar_max_col) = (
                avg_range.0.min(avg_range.2), avg_range.1.min(avg_range.3),
                avg_range.0.max(avg_range.2), avg_range.1.max(avg_range.3)
            );
            let num_rows = ar_max_row - ar_min_row + 1;
            let num_cols = ar_max_col - ar_min_col + 1;

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return Some(EvalResult::Error("AVERAGEIFS criteria_range must be a range".to_string())),
                };

                // Verify dimensions match average_range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return Some(EvalResult::Error("AVERAGEIFS criteria ranges must have same dimensions as average_range".to_string()));
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Number(0.0)
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        let val = lookup.get_value(ar_min_row + row_offset, ar_min_col + col_offset);
                        if val.is_finite() {
                            sum += val;
                            count += 1;
                        }
                    }
                }
            }
            if count == 0 {
                EvalResult::Error("#DIV/0!".to_string())
            } else {
                EvalResult::Number(sum / count as f64)
            }
        }
        "COUNTIFS" => {
            // COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
            if args.len() < 2 || args.len() % 2 != 0 {
                return Some(EvalResult::Error("COUNTIFS requires pairs of criteria_range and criteria".to_string()));
            }

            // Use first range to determine dimensions
            let first_range = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                _ => return Some(EvalResult::Error("COUNTIFS criteria_range must be a range".to_string())),
            };

            let (fr_min_row, fr_min_col, fr_max_row, fr_max_col) = (
                first_range.0.min(first_range.2), first_range.1.min(first_range.3),
                first_range.0.max(first_range.2), first_range.1.max(first_range.3)
            );
            let num_rows = fr_max_row - fr_min_row + 1;
            let num_cols = fr_max_col - fr_min_col + 1;

            // Parse criteria pairs
            let num_criteria = args.len() / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[i * 2];
                let criteria_arg = &args[i * 2 + 1];

                let crit_range = match range_arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => (*start_row, *start_col, *end_row, *end_col),
                    _ => return Some(EvalResult::Error("COUNTIFS criteria_range must be a range".to_string())),
                };

                // Verify dimensions match first range
                let (cr_min_row, cr_min_col, cr_max_row, cr_max_col) = (
                    crit_range.0.min(crit_range.2), crit_range.1.min(crit_range.3),
                    crit_range.0.max(crit_range.2), crit_range.1.max(crit_range.3)
                );
                if (cr_max_row - cr_min_row + 1) != num_rows || (cr_max_col - cr_min_col + 1) != num_cols {
                    return Some(EvalResult::Error("COUNTIFS ranges must have same dimensions".to_string()));
                }

                criteria_ranges.push((cr_min_row, cr_min_col));
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    // Check all criteria
                    let mut all_match = true;
                    for (idx, &(cr_row, cr_col)) in criteria_ranges.iter().enumerate() {
                        let r = cr_row + row_offset;
                        let c = cr_col + col_offset;
                        let cell_text = lookup.get_text(r, c);
                        let cell_value = if cell_text.is_empty() {
                            EvalResult::Text(String::new())
                        } else if let Ok(n) = cell_text.parse::<f64>() {
                            EvalResult::Number(n)
                        } else {
                            EvalResult::Text(cell_text)
                        };

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        count += 1;
                    }
                }
            }
            EvalResult::Number(count as f64)
        }
        _ => return None,
    };
    Some(result)
}
