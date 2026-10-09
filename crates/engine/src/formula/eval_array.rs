// Array/spill functions: SEQUENCE, TRANSPOSE, FILTER, UNIQUE, SORT, SORTBY, SPARKLINE,
// ARRAYFORMULA

use super::eval::{evaluate, CellLookup, EvalResult, Value, Array2D};
use super::eval_helpers::{collect_numbers, read_cell_value, value_compare};
use crate::sheet::SheetRef;
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
            if let Err(error) = super::eval_budget::array(rows, cols) { return Some(EvalResult::Error(error)); }
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
            // TRANSPOSE(array): rows become columns. Values keep their type —
            // this used to read every cell as a number, so text became 0.
            if args.len() != 1 {
                return Some(EvalResult::Error("TRANSPOSE requires exactly one argument".to_string()));
            }
            // A computed array (FILTER(...), UNIQUE(...)) transposes like a range;
            // it used to come back unchanged.
            match grid_values(&args[0], lookup) {
                Ok((in_rows, in_cols, rows)) => {
                    if let Err(error) = super::eval_budget::array(in_cols, in_rows) { return Some(EvalResult::Error(error)); }
                    let mut array = Array2D::new(in_cols, in_rows);
                    for (r, row) in rows.into_iter().enumerate() {
                        for (c, val) in row.into_iter().enumerate() {
                            // An empty cell transposes to 0, as in Excel.
                            let val = if matches!(val, Value::Empty) { Value::Number(0.0) } else { val };
                            array.set(c, r, val);
                        }
                    }
                    EvalResult::Array(array)
                }
                Err(e) => EvalResult::Error(e),
            }
        }

        "FILTER" => {
            // FILTER(array, include, [if_empty])
            // Returns the rows (or columns) of array where include is TRUE.
            // include may be a range or a computed condition such as B2:B9>15.
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("FILTER requires 2 or 3 arguments".to_string()));
            }

            let (data_rows, data_cols, data) = match grid_values(&args[0], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (inc_rows, inc_cols, inc) = match grid_values(&args[1], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            // An error anywhere in the condition is the answer, as in Excel.
            if let Some(Value::Error(e)) = inc.iter().flatten().find(|v| matches!(v, Value::Error(_))) {
                return Some(EvalResult::Error(e.clone()));
            }
            let flags: Vec<bool> = inc.iter().flatten().map(include_flag).collect();

            let kept: Vec<Vec<Value>> = if inc_cols == 1 && (inc_rows == data_rows || inc_rows == 1) {
                // One flag per row (a single flag keeps or drops everything).
                data.into_iter()
                    .enumerate()
                    .filter(|(r, _)| flags[if inc_rows == 1 { 0 } else { *r }])
                    .map(|(_, row)| row)
                    .collect()
            } else if inc_rows == 1 && inc_cols == data_cols {
                // One flag per column.
                if !flags.iter().any(|&f| f) {
                    Vec::new()
                } else {
                    data.into_iter()
                        .map(|row| row.into_iter().zip(&flags).filter(|(_, &f)| f).map(|(v, _)| v).collect())
                        .collect()
                }
            } else {
                return Some(EvalResult::Error(format!(
                    "#VALUE! Include is {}x{} but must match the {} rows or {} columns of the data",
                    inc_rows, inc_cols, data_rows, data_cols
                )));
            };

            if kept.is_empty() {
                if let Some(if_empty) = args.get(2) {
                    return Some(evaluate(if_empty, lookup));
                }
                return Some(EvalResult::Error("#CALC! No matches".to_string()));
            }
            EvalResult::Array(Array2D::from_vec(kept))
        }

        "UNIQUE" => {
            // UNIQUE(range)
            // Returns unique rows from a range (preserves first occurrence order)
            if args.len() != 1 {
                return Some(EvalResult::Error("UNIQUE requires exactly one argument".to_string()));
            }

            // Rows from a range or a computed array, typed. A computed array
            // (UNIQUE(FILTER(...))) used to come back unchanged, duplicates and all.
            let (_, in_cols, rows) = match grid_values(&args[0], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
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
            if let Err(error) = super::eval_budget::array(out_rows, in_cols) { return Some(EvalResult::Error(error)); }
            let mut array = Array2D::new(out_rows, in_cols);
            for (r, row) in unique_rows.iter().enumerate() {
                for (c, val) in row.iter().enumerate() {
                    array.set(r, c, val.clone());
                }
            }

            EvalResult::Array(array)
        }

        "SORT" => {
            // SORT(array, [sort_index], [sort_order])
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

            // Excel's sort_order is 1 (ascending) or -1 (descending). This used to be
            // read as a boolean, and -1 is truthy, so every SORT(x,1,-1) written for
            // Excel sorted ascending. TRUE/FALSE are still accepted for sheets that
            // relied on the old reading.
            let descending = if args.len() >= 3 {
                match evaluate(&args[2], lookup) {
                    EvalResult::Boolean(b) => !b,
                    EvalResult::Number(n) if n == 1.0 => false,
                    EvalResult::Number(n) if n == -1.0 => true,
                    EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                    _ => return Some(EvalResult::Error("#VALUE! Sort order must be 1 or -1".to_string())),
                }
            } else {
                false
            };

            // A computed array sorts like a range. SORT(UNIQUE(x)) used to return
            // UNIQUE's result untouched, which looks sorted until it isn't.
            let (in_rows, in_cols, mut rows) = match grid_values(&args[0], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            // Validate sort column
            if sort_col_1idx > in_cols {
                return Some(EvalResult::Error(format!("#VALUE! Sort column {} exceeds range width {}", sort_col_1idx, in_cols)));
            }
            let key = sort_col_1idx - 1;

            // Stable in both directions: reversing an ascending sort would also
            // reverse the order of rows with equal keys.
            rows.sort_by(|a, b| {
                let ord = value_compare(&a[key], &b[key]);
                if descending { ord.reverse() } else { ord }
            });

            // Build result array
            if let Err(error) = super::eval_budget::array(in_rows, in_cols) { return Some(EvalResult::Error(error)); }
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

        "SORTBY" => {
            // SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2],
            // ...): sort array's rows by one or more key columns (each as tall
            // as array), or its columns by key rows (each as wide). Orders are
            // 1 (ascending, the default) or -1. Stable, so equal keys keep
            // their order, and keys compare as SORT's do. Keys of the wrong
            // shape, or mixed orientations, are #VALUE!.
            if args.len() < 2 {
                return Some(EvalResult::Error("SORTBY requires an array and at least one by_array".to_string()));
            }
            let (rows, cols, data) = match grid_values(&args[0], lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let mut keys: Vec<(Vec<Value>, bool)> = Vec::new();
            let mut by_rows: Option<bool> = None;
            for pair in args[1..].chunks(2) {
                let (kr, kc, kd) = match grid_values(&pair[0], lookup) {
                    Ok(v) => v,
                    Err(e) => return Some(EvalResult::Error(e)),
                };
                let orientation = if kc == 1 && kr == rows {
                    true
                } else if kr == 1 && kc == cols {
                    false
                } else {
                    return Some(EvalResult::Error("#VALUE!".to_string()));
                };
                if by_rows.is_some_and(|o| o != orientation) {
                    return Some(EvalResult::Error("#VALUE!".to_string()));
                }
                by_rows = Some(orientation);
                let descending = match pair.get(1).filter(|a| !matches!(a, Expr::Empty)).map(|a| evaluate(a, lookup)) {
                    None => false,
                    Some(EvalResult::Number(1.0)) => false,
                    Some(EvalResult::Number(-1.0)) => true,
                    Some(EvalResult::Error(e)) => return Some(EvalResult::Error(e)),
                    Some(_) => return Some(EvalResult::Error("#VALUE!".to_string())),
                };
                let flat: Vec<Value> = kd.into_iter().flatten().collect();
                keys.push((flat, descending));
            }
            let by_rows = by_rows.unwrap_or(true);
            let n = if by_rows { rows } else { cols };
            let mut order: Vec<usize> = (0..n).collect();
            order.sort_by(|&a, &b| {
                for (key, descending) in &keys {
                    let ord = value_compare(&key[a], &key[b]);
                    let ord = if *descending { ord.reverse() } else { ord };
                    if ord != std::cmp::Ordering::Equal {
                        return ord;
                    }
                }
                std::cmp::Ordering::Equal
            });
            if let Err(error) = super::eval_budget::array(rows, cols) { return Some(EvalResult::Error(error)); }
            let mut array = Array2D::new(rows, cols);
            for r in 0..rows {
                for c in 0..cols {
                    let (sr, sc) = if by_rows { (order[r], c) } else { (r, order[c]) };
                    array.set(r, c, data[sr][sc].clone());
                }
            }
            EvalResult::Array(array)
        }
        "ARRAYFORMULA" => {
            // ARRAYFORMULA(array_formula): Google Sheets' switch into array
            // mode. Here every formula is already in array mode — operators
            // and value functions go element by element over ranges, and an
            // array result spills — so this evaluates its argument as an
            // operator operand (a bare range becomes an array) and returns
            // it. It exists so formulas written for Sheets keep working.
            if args.len() != 1 {
                return Some(EvalResult::Error("ARRAYFORMULA requires exactly one argument".to_string()));
            }
            super::eval::operand(&args[0], lookup)
        }
        _ => return None,
    };
    Some(result)
}

/// The cells of a range argument, typed and read from the sheet the range names.
///
/// None when the argument is not a range. SORT, TRANSPOSE and FILTER each used to
/// read the current sheet by text and guess the type back, which lost text in
/// TRANSPOSE and ignored the sheet in a reference like Data!A1:A9.
fn range_values<L: CellLookup>(
    arg: &BoundExpr,
    lookup: &L,
) -> Option<Result<(usize, usize, Vec<Vec<Value>>), String>> {
    let Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } = arg else {
        return None;
    };
    if matches!(sheet, SheetRef::RefError { .. }) {
        return Some(Err("#REF!".to_string()));
    }
    let (r0, r1) = (*start_row.min(end_row), *start_row.max(end_row));
    let (c0, c1) = (*start_col.min(end_col), *start_col.max(end_col));
    if let Err(error) = super::eval_budget::array(r1 - r0 + 1, c1 - c0 + 1) { return Some(Err(error)); }
    let rows = (r0..=r1)
        .map(|r| (c0..=c1).map(|c| read_cell_value(lookup, sheet, r, c)).collect())
        .collect();
    Some(Ok((r1 - r0 + 1, c1 - c0 + 1, rows)))
}

/// Whether one cell of FILTER's include column selects its row.
fn include_flag(v: &Value) -> bool {
    match v {
        Value::Boolean(b) => *b,
        Value::Number(n) => *n != 0.0,
        // A typed TRUE is stored as text.
        Value::Text(s) => s.eq_ignore_ascii_case("TRUE"),
        Value::Empty | Value::Error(_) => false,
    }
}

/// A function argument as a grid of values: a range's cells, a computed array's
/// elements, or a single value as a 1x1 grid.
fn grid_values<L: CellLookup>(
    arg: &BoundExpr,
    lookup: &L,
) -> Result<(usize, usize, Vec<Vec<Value>>), String> {
    if let Some(range) = range_values(arg, lookup) {
        return range;
    }
    match evaluate(arg, lookup) {
        EvalResult::Array(a) => {
            super::eval_budget::array(a.rows(), a.cols())?;
            let data: Vec<Vec<Value>> = (0..a.rows())
                .map(|r| (0..a.cols()).map(|c| a.get(r, c).cloned().unwrap_or(Value::Empty)).collect())
                .collect();
            Ok((a.rows(), a.cols(), data))
        }
        EvalResult::Error(e) => Err(e),
        other => Ok((1, 1, vec![vec![other.to_value()]])),
    }
}
