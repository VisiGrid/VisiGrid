// Array/spill functions: SEQUENCE, TRANSPOSE, FILTER, UNIQUE, SORT, SPARKLINE

use super::eval::{evaluate, CellLookup, EvalResult, Value, Array2D};
use super::eval_helpers::{collect_numbers, value_compare};
use super::parser::{BoundExpr, Expr};

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "SEQUENCE" => {
            // SEQUENCE(rows, [cols], [start], [step])
            // Returns a 2D array of sequential numbers
            if args.is_empty() || args.len() > 4 {
                return Some(EvalResult::Error("SEQUENCE requires 1-4 arguments".to_string()));
            }

            let rows = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let cols = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };

            let start = if args.len() >= 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1.0
            };

            let step = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1.0
            };

            // Build the array
            let mut array = Array2D::new(rows, cols);
            let mut val = start;
            for r in 0..rows {
                for c in 0..cols {
                    array.set(r, c, Value::Number(val));
                    val += step;
                }
            }
            EvalResult::Array(array)
        }

        "TRANSPOSE" => {
            // TRANSPOSE(array)
            // Returns the transpose of an array/range
            if args.len() != 1 {
                return Some(EvalResult::Error("TRANSPOSE requires exactly one argument".to_string()));
            }

            // Get the input - if it's a range, build an array from it
            match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let in_rows = end_row - start_row + 1;
                    let in_cols = end_col - start_col + 1;

                    // Build transposed array (swap rows and cols)
                    let mut array = Array2D::new(in_cols, in_rows);
                    for r in 0..in_rows {
                        for c in 0..in_cols {
                            let val = lookup.get_value(start_row + r, start_col + c);
                            array.set(c, r, Value::Number(val));
                        }
                    }
                    EvalResult::Array(array)
                }
                _ => {
                    // Single value - just return it (1x1 transpose is identity)
                    evaluate(&args[0], lookup)
                }
            }
        }

        "FILTER" => {
            // FILTER(range, include)
            // Returns rows from range where include is TRUE
            if args.len() != 2 {
                return Some(EvalResult::Error("FILTER requires exactly 2 arguments".to_string()));
            }

            // Get the data range dimensions and values
            let (data_rows, data_cols, data): (usize, usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if text.parse::<f64>().is_ok() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (r_count, c_count, row_data)
                }
                _ => {
                    return Some(EvalResult::Error("#VALUE! FILTER requires a range as first argument".to_string()));
                }
            };

            // Get the include criteria (must be a column matching data_rows)
            let include: Vec<bool> = match &args[1] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let inc_rows = end_row - start_row + 1;
                    let inc_cols = end_col - start_col + 1;

                    // Must be a single column with matching row count
                    if inc_cols != 1 {
                        return Some(EvalResult::Error("#VALUE! Include must be a single column".to_string()));
                    }
                    if inc_rows != data_rows {
                        return Some(EvalResult::Error(format!("#VALUE! Include has {} rows but data has {} rows", inc_rows, data_rows)));
                    }

                    let mut criteria = Vec::with_capacity(inc_rows);
                    for r in 0..inc_rows {
                        let val = lookup.get_value(start_row + r, *start_col);
                        // Treat non-zero as TRUE, zero as FALSE
                        criteria.push(val != 0.0);
                    }
                    criteria
                }
                _ => {
                    // Try evaluating as a single value (scalar comparison result)
                    match evaluate(&args[1], lookup) {
                        EvalResult::Boolean(b) => vec![b; data_rows],
                        EvalResult::Number(n) => vec![n != 0.0; data_rows],
                        _ => return Some(EvalResult::Error("#VALUE! Include must be a range or boolean".to_string())),
                    }
                }
            };

            // Filter rows where include is TRUE
            let filtered_rows: Vec<Vec<Value>> = data.into_iter()
                .zip(include.iter())
                .filter(|(_, &inc)| inc)
                .map(|(row, _)| row)
                .collect();

            if filtered_rows.is_empty() {
                return Some(EvalResult::Error("#CALC! No matches".to_string()));
            }

            // Build result array
            let out_rows = filtered_rows.len();
            let mut array = Array2D::new(out_rows, data_cols);
            for (r, row) in filtered_rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "UNIQUE" => {
            // UNIQUE(range)
            // Returns unique rows from a range (preserves first occurrence order)
            if args.len() != 1 {
                return Some(EvalResult::Error("UNIQUE requires exactly one argument".to_string()));
            }

            // Build rows from range
            let (in_cols, rows): (usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if text.parse::<f64>().is_ok() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (c_count, row_data)
                }
                _ => {
                    // Single value - return as-is
                    return Some(evaluate(&args[0], lookup));
                }
            };

            // Find unique rows (first occurrence wins, case-insensitive for text)
            let mut unique_rows: Vec<Vec<Value>> = Vec::new();
            let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();

            for row in rows {
                // Create a canonical key for the row (case-insensitive)
                let key = row.iter()
                    .map(|v| match v {
                        Value::Text(s) => s.to_lowercase(),
                        other => other.to_text(),
                    })
                    .collect::<Vec<_>>()
                    .join("\x00"); // Use null byte as separator

                if !seen.contains(&key) {
                    seen.insert(key);
                    unique_rows.push(row);
                }
            }

            if unique_rows.is_empty() {
                return Some(EvalResult::Error("#CALC! No data".to_string()));
            }

            // Build result array
            let out_rows = unique_rows.len();
            let mut array = Array2D::new(out_rows, in_cols);
            for (r, row) in unique_rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "SORT" => {
            // SORT(array_or_range, [sort_col], [is_asc])
            if args.is_empty() || args.len() > 3 {
                return Some(EvalResult::Error("SORT requires 1-3 arguments".to_string()));
            }

            // Get sort column (1-indexed, default 1)
            let sort_col_1idx = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE! Sort column must be >= 1".to_string())),
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };

            // Get ascending flag (default true)
            let is_asc = if args.len() >= 3 {
                match evaluate(&args[2], lookup).to_bool() {
                    Ok(b) => b,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                true
            };

            // Build rows from range
            let (in_rows, in_cols, mut rows): (usize, usize, Vec<Vec<Value>>) = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let r_count = end_row - start_row + 1;
                    let c_count = end_col - start_col + 1;
                    let mut row_data = Vec::with_capacity(r_count);
                    for r in 0..r_count {
                        let mut row = Vec::with_capacity(c_count);
                        for c in 0..c_count {
                            let text = lookup.get_text(start_row + r, start_col + c);
                            let val = lookup.get_value(start_row + r, start_col + c);
                            // Determine value type
                            if text.is_empty() {
                                row.push(Value::Empty);
                            } else if text.starts_with('#') {
                                row.push(Value::Error(text));
                            } else if text.parse::<f64>().is_ok() {
                                row.push(Value::Number(val));
                            } else {
                                row.push(Value::Text(text));
                            }
                        }
                        row_data.push(row);
                    }
                    (r_count, c_count, row_data)
                }
                _ => {
                    // Single value - can't sort meaningfully
                    return Some(evaluate(&args[0], lookup));
                }
            };

            // Validate sort column
            if sort_col_1idx > in_cols {
                return Some(EvalResult::Error(format!("#VALUE! Sort column {} exceeds range width {}", sort_col_1idx, in_cols)));
            }
            let sort_col_0idx = sort_col_1idx - 1;

            // Sort rows by the key column (stable sort)
            rows.sort_by(|a, b| {
                let key_a = &a[sort_col_0idx];
                let key_b = &b[sort_col_0idx];
                value_compare(key_a, key_b)
            });

            // Reverse if descending
            if !is_asc {
                rows.reverse();
            }

            // Build result array
            let mut array = Array2D::new(in_rows, in_cols);
            for (r, row) in rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "SPARKLINE" => {
            // SPARKLINE(data_range, [type])
            // Creates a Unicode mini-chart from numeric data
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("SPARKLINE requires 1-2 arguments".to_string()));
            }

            // Collect numbers from first argument
            let nums = match collect_numbers(&args[0..1], lookup) {
                Ok(v) if v.is_empty() => return Some(EvalResult::Text(String::new())),
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            // Get chart type (default "bar")
            let chart_type = if args.len() > 1 {
                evaluate(&args[1], lookup).to_text().to_lowercase()
            } else {
                "bar".to_string()
            };

            match chart_type.as_str() {
                "bar" | "line" => {
                    // Bar/line sparkline using Unicode block characters
                    // ▁▂▃▄▅▆▇█ (U+2581 to U+2588) - 8 height levels
                    const BARS: [char; 8] = ['▁', '▂', '▃', '▄', '▅', '▆', '▇', '█'];

                    let min = nums.iter().cloned().fold(f64::INFINITY, f64::min);
                    let max = nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
                    let range = max - min;

                    let sparkline: String = if range == 0.0 {
                        // All values equal - show middle bars
                        BARS[3].to_string().repeat(nums.len())
                    } else {
                        nums.iter().map(|&n| {
                            let normalized = (n - min) / range;
                            let idx = ((normalized * 7.0).round() as usize).min(7);
                            BARS[idx]
                        }).collect()
                    };
                    EvalResult::Text(sparkline)
                }
                "winloss" => {
                    // Win/loss sparkline: ▲ for positive, ▼ for negative, ▬ for zero
                    let sparkline: String = nums.iter().map(|&n| {
                        if n > 0.0 { '▲' }
                        else if n < 0.0 { '▼' }
                        else { '▬' }
                    }).collect();
                    EvalResult::Text(sparkline)
                }
                _ => EvalResult::Error(format!("Unknown sparkline type: {}. Use 'bar', 'line', or 'winloss'", chart_type)),
            }
        }

        _ => return None,
    };
    Some(result)
}
