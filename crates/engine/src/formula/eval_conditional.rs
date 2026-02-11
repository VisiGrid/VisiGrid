// Conditional aggregate functions: SUMIF, AVERAGEIF, COUNTIF, COUNTBLANK,
// SUMIFS, AVERAGEIFS, COUNTIFS

use crate::sheet::SheetRef;
use super::eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution};
use super::eval_helpers::matches_criteria;
use super::parser::{BoundExpr, Expr};

/// Extracted range with sheet context.
struct RangeRef {
    sheet: SheetRef,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
}

impl RangeRef {
    fn min_row(&self) -> usize { self.start_row.min(self.end_row) }
    fn min_col(&self) -> usize { self.start_col.min(self.end_col) }
    fn max_row(&self) -> usize { self.start_row.max(self.end_row) }
    fn max_col(&self) -> usize { self.start_col.max(self.end_col) }
    fn num_rows(&self) -> usize { self.max_row() - self.min_row() + 1 }
    fn num_cols(&self) -> usize { self.max_col() - self.min_col() + 1 }
}

/// Extract range coordinates and sheet reference from a range-like expression.
fn extract_range<L: CellLookup>(
    expr: &BoundExpr, lookup: &L, arg_name: &str,
) -> Result<RangeRef, String> {
    match expr {
        Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
            Ok(RangeRef {
                sheet: sheet.clone(),
                start_row: *start_row, start_col: *start_col,
                end_row: *end_row, end_col: *end_col,
            })
        }
        Expr::CellRef { sheet, col, row, .. } => {
            Ok(RangeRef {
                sheet: sheet.clone(),
                start_row: *row, start_col: *col,
                end_row: *row, end_col: *col,
            })
        }
        Expr::NamedRange(name) => {
            match lookup.resolve_named_range(name) {
                Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                    Ok(RangeRef {
                        sheet: SheetRef::Current,
                        start_row, start_col, end_row, end_col,
                    })
                }
                Some(NamedRangeResolution::Cell { row, col }) => {
                    Ok(RangeRef {
                        sheet: SheetRef::Current,
                        start_row: row, start_col: col,
                        end_row: row, end_col: col,
                    })
                }
                None => Err(format!("#NAME? '{}'", name)),
            }
        }
        _ => Err(format!(
            "{} must be a range (e.g. A1:A10), cell reference, or named range",
            arg_name
        )),
    }
}

/// Get cell text from the correct sheet.
fn range_get_text<L: CellLookup>(lookup: &L, sheet: &SheetRef, row: usize, col: usize) -> String {
    match sheet {
        SheetRef::Current => lookup.get_text(row, col),
        SheetRef::Id(sid) => lookup.get_text_sheet(*sid, row, col),
        SheetRef::RefError { .. } => "#REF!".to_string(),
    }
}

/// Get cell numeric value from the correct sheet.
fn range_get_value<L: CellLookup>(lookup: &L, sheet: &SheetRef, row: usize, col: usize) -> f64 {
    match sheet {
        SheetRef::Current => lookup.get_value(row, col),
        SheetRef::Id(sid) => {
            use super::eval::Value;
            match lookup.get_value_sheet(*sid, row, col) {
                Value::Number(n) => n,
                Value::Boolean(true) => 1.0,
                Value::Boolean(false) => 0.0,
                _ => 0.0,
            }
        }
        SheetRef::RefError { .. } => 0.0,
    }
}

/// Parse cell text into an EvalResult for criteria matching.
fn text_to_eval_result(text: &str, empty_as_number: bool) -> EvalResult {
    if text.is_empty() {
        if empty_as_number {
            EvalResult::Number(0.0)
        } else {
            EvalResult::Text(String::new())
        }
    } else if let Ok(n) = text.parse::<f64>() {
        EvalResult::Number(n)
    } else {
        EvalResult::Text(text.to_string())
    }
}

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "SUMIF" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("SUMIF requires 2 or 3 arguments".to_string()));
            }
            let range = match extract_range(&args[0], lookup, "SUMIF range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let criteria = evaluate(&args[1], lookup);
            let sum_range = if args.len() == 3 {
                match extract_range(&args[2], lookup, "SUMIF sum_range") {
                    Ok(r) => Some(r),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let (min_row, min_col, max_row, max_col) = (range.min_row(), range.min_col(), range.max_row(), range.max_col());

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = range_get_text(lookup, &range.sheet, r, c);
                    let cell_value = text_to_eval_result(&cell_text, true);

                    if matches_criteria(&cell_value, &criteria) {
                        if let Some(ref sr) = sum_range {
                            let sum_r = sr.min_row() + row_offset;
                            let sum_c = sr.min_col() + col_offset;
                            sum += range_get_value(lookup, &sr.sheet, sum_r, sum_c);
                        } else {
                            sum += range_get_value(lookup, &range.sheet, r, c);
                        }
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
            let range = match extract_range(&args[0], lookup, "AVERAGEIF range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let criteria = evaluate(&args[1], lookup);
            let avg_range = if args.len() == 3 {
                match extract_range(&args[2], lookup, "AVERAGEIF average_range") {
                    Ok(r) => Some(r),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                None
            };

            let mut sum = 0.0;
            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.min_row(), range.min_col(), range.max_row(), range.max_col());

            for row_offset in 0..=(max_row - min_row) {
                for col_offset in 0..=(max_col - min_col) {
                    let r = min_row + row_offset;
                    let c = min_col + col_offset;
                    let cell_text = range_get_text(lookup, &range.sheet, r, c);
                    let cell_value = text_to_eval_result(&cell_text, true);

                    if matches_criteria(&cell_value, &criteria) {
                        let (avg_r, avg_c, avg_sheet) = if let Some(ref ar) = avg_range {
                            (ar.min_row() + row_offset, ar.min_col() + col_offset, &ar.sheet)
                        } else {
                            (r, c, &range.sheet)
                        };
                        let val = range_get_value(lookup, avg_sheet, avg_r, avg_c);
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
            let range = match extract_range(&args[0], lookup, "COUNTIF range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let criteria = evaluate(&args[1], lookup);

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.min_row(), range.min_col(), range.max_row(), range.max_col());

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    let cell_text = range_get_text(lookup, &range.sheet, r, c);
                    let cell_value = text_to_eval_result(&cell_text, false);

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
            let range = match extract_range(&args[0], lookup, "COUNTBLANK range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let mut count = 0;
            let (min_row, min_col, max_row, max_col) = (range.min_row(), range.min_col(), range.max_row(), range.max_col());

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if range_get_text(lookup, &range.sheet, r, c).is_empty() {
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

            let sum_range = match extract_range(&args[0], lookup, "SUMIFS sum_range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let num_rows = sum_range.num_rows();
            let num_cols = sum_range.num_cols();

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match extract_range(range_arg, lookup, "SUMIFS criteria_range") {
                    Ok(r) => r,
                    Err(e) => return Some(EvalResult::Error(e)),
                };

                if crit_range.num_rows() != num_rows || crit_range.num_cols() != num_cols {
                    return Some(EvalResult::Error("SUMIFS criteria ranges must have same dimensions as sum_range".to_string()));
                }

                criteria_ranges.push(crit_range);
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    let mut all_match = true;
                    for (idx, cr) in criteria_ranges.iter().enumerate() {
                        let r = cr.min_row() + row_offset;
                        let c = cr.min_col() + col_offset;
                        let cell_text = range_get_text(lookup, &cr.sheet, r, c);
                        let cell_value = text_to_eval_result(&cell_text, true);

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        sum += range_get_value(
                            lookup, &sum_range.sheet,
                            sum_range.min_row() + row_offset,
                            sum_range.min_col() + col_offset,
                        );
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

            let avg_range = match extract_range(&args[0], lookup, "AVERAGEIFS average_range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let num_rows = avg_range.num_rows();
            let num_cols = avg_range.num_cols();

            // Parse criteria pairs
            let num_criteria = (args.len() - 1) / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[1 + i * 2];
                let criteria_arg = &args[2 + i * 2];

                let crit_range = match extract_range(range_arg, lookup, "AVERAGEIFS criteria_range") {
                    Ok(r) => r,
                    Err(e) => return Some(EvalResult::Error(e)),
                };

                if crit_range.num_rows() != num_rows || crit_range.num_cols() != num_cols {
                    return Some(EvalResult::Error("AVERAGEIFS criteria ranges must have same dimensions as average_range".to_string()));
                }

                criteria_ranges.push(crit_range);
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut sum = 0.0;
            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    let mut all_match = true;
                    for (idx, cr) in criteria_ranges.iter().enumerate() {
                        let r = cr.min_row() + row_offset;
                        let c = cr.min_col() + col_offset;
                        let cell_text = range_get_text(lookup, &cr.sheet, r, c);
                        let cell_value = text_to_eval_result(&cell_text, true);

                        if !matches_criteria(&cell_value, &criteria_values[idx]) {
                            all_match = false;
                            break;
                        }
                    }

                    if all_match {
                        let val = range_get_value(
                            lookup, &avg_range.sheet,
                            avg_range.min_row() + row_offset,
                            avg_range.min_col() + col_offset,
                        );
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
            let first_range = match extract_range(&args[0], lookup, "COUNTIFS criteria_range") {
                Ok(r) => r,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let num_rows = first_range.num_rows();
            let num_cols = first_range.num_cols();

            // Parse criteria pairs
            let num_criteria = args.len() / 2;
            let mut criteria_ranges = Vec::with_capacity(num_criteria);
            let mut criteria_values = Vec::with_capacity(num_criteria);

            for i in 0..num_criteria {
                let range_arg = &args[i * 2];
                let criteria_arg = &args[i * 2 + 1];

                let crit_range = match extract_range(range_arg, lookup, "COUNTIFS criteria_range") {
                    Ok(r) => r,
                    Err(e) => return Some(EvalResult::Error(e)),
                };

                if crit_range.num_rows() != num_rows || crit_range.num_cols() != num_cols {
                    return Some(EvalResult::Error("COUNTIFS ranges must have same dimensions".to_string()));
                }

                criteria_ranges.push(crit_range);
                criteria_values.push(evaluate(criteria_arg, lookup));
            }

            let mut count = 0;
            for row_offset in 0..num_rows {
                for col_offset in 0..num_cols {
                    let mut all_match = true;
                    for (idx, cr) in criteria_ranges.iter().enumerate() {
                        let r = cr.min_row() + row_offset;
                        let c = cr.min_col() + col_offset;
                        let cell_text = range_get_text(lookup, &cr.sheet, r, c);
                        let cell_value = text_to_eval_result(&cell_text, false);

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
