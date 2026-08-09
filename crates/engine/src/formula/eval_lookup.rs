// Lookup/reference functions: VLOOKUP, XLOOKUP, HLOOKUP, INDEX, MATCH,
// ROW, COLUMN, ROWS, COLUMNS

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, Value};
use super::eval_helpers::{get_text_for_sheet, wildcard_match};
use super::parser::{BoundExpr, Expr};
use crate::sheet::SheetRef;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "VLOOKUP" => {
            // VLOOKUP(search_key, range, index, [is_sorted])
            if args.len() < 3 || args.len() > 4 {
                return Some(EvalResult::Error("VLOOKUP requires 3 or 4 arguments".to_string()));
            }
            let search_key = evaluate(&args[0], lookup);
            let (range, sheet_ref) = match &args[1] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => ((*start_row, *start_col, *end_row, *end_col), sheet),
                _ => return Some(EvalResult::Error("VLOOKUP requires a range as second argument".to_string())),
            };
            let col_index = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let is_sorted = if args.len() == 4 {
                match evaluate(&args[3], lookup).to_bool() {
                    Ok(b) => b,
                    Err(_) => true, // default to TRUE
                }
            } else {
                true
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_cols = max_col - min_col + 1;

            if col_index > num_cols {
                return Some(EvalResult::Error("#REF!".to_string()));
            }

            // Search for the key in the first column
            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_row: Option<usize> = None;

            // Helper closure to get text from the correct sheet
            let get = |r: usize, c: usize| -> String {
                match get_text_for_sheet(lookup, sheet_ref, r, c) {
                    Ok(t) => t,
                    Err(_) => String::new(),
                }
            };

            if is_sorted {
                // Approximate match (find largest value <= search_key)
                if let Some(search_n) = search_num {
                    // Numeric key: find largest numeric cell <= search_n
                    let mut best_row: Option<usize> = None;
                    let mut best_val = f64::NEG_INFINITY;
                    for r in min_row..=max_row {
                        let cell_text = get(r, min_col);
                        if let Ok(cell_n) = cell_text.parse::<f64>() {
                            if cell_n <= search_n && cell_n > best_val {
                                best_val = cell_n;
                                best_row = Some(r);
                            }
                        }
                    }
                    found_row = best_row;
                } else {
                    // Text key: find largest text cell <= search_text (case-insensitive)
                    let mut best_row: Option<usize> = None;
                    let mut best_val = String::new();
                    for r in min_row..=max_row {
                        let cell_text = get(r, min_col);
                        let cell_lower = cell_text.to_lowercase();
                        // Skip numeric cells and empties when searching text
                        if cell_text.is_empty() || cell_text.parse::<f64>().is_ok() {
                            continue;
                        }
                        if cell_lower <= search_text && (best_row.is_none() || cell_lower > best_val) {
                            best_val = cell_lower;
                            best_row = Some(r);
                        }
                    }
                    found_row = best_row;
                }
            } else {
                // Exact match
                for r in min_row..=max_row {
                    let cell_text = get(r, min_col);
                    let cell_lower = cell_text.to_lowercase();

                    // Try numeric comparison first
                    if let (Some(search_n), Ok(cell_n)) = (search_num, cell_text.parse::<f64>()) {
                        if (search_n - cell_n).abs() < f64::EPSILON {
                            found_row = Some(r);
                            break;
                        }
                    } else if cell_lower == search_text {
                        found_row = Some(r);
                        break;
                    }
                }
            }

            match found_row {
                Some(r) => {
                    let result_col = min_col + col_index - 1;
                    let result_text = get(r, result_col);
                    if result_text.is_empty() {
                        EvalResult::Empty
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "XLOOKUP" => {
            // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
            if args.len() < 3 || args.len() > 6 {
                return Some(EvalResult::Error("XLOOKUP requires 3 to 6 arguments".to_string()));
            }

            let lookup_value = evaluate(&args[0], lookup);

            // Parse lookup array (must be 1D - single row or column)
            let (lookup_array, xl_lookup_sheet) = match &args[1] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col, max_row, max_col) = (
                        (*start_row).min(*end_row), (*start_col).min(*end_col),
                        (*start_row).max(*end_row), (*start_col).max(*end_col)
                    );
                    // Determine orientation
                    let is_row = min_row == max_row;
                    let is_col = min_col == max_col;
                    if !is_row && !is_col {
                        return Some(EvalResult::Error("XLOOKUP lookup_array must be a single row or column".to_string()));
                    }
                    ((min_row, min_col, max_row, max_col, is_row), sheet)
                }
                _ => return Some(EvalResult::Error("XLOOKUP lookup_array must be a range".to_string())),
            };

            // Parse return array (must have same dimensions)
            let (return_array, xl_return_sheet) = match &args[2] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col, max_row, max_col) = (
                        (*start_row).min(*end_row), (*start_col).min(*end_col),
                        (*start_row).max(*end_row), (*start_col).max(*end_col)
                    );
                    ((min_row, min_col, max_row, max_col), sheet)
                }
                _ => return Some(EvalResult::Error("XLOOKUP return_array must be a range".to_string())),
            };

            // Verify lookup and return arrays have compatible dimensions
            let lookup_size = if lookup_array.4 {
                lookup_array.3 - lookup_array.1 + 1  // columns for row array
            } else {
                lookup_array.2 - lookup_array.0 + 1  // rows for column array
            };

            let return_size = if lookup_array.4 {
                return_array.3 - return_array.1 + 1  // columns
            } else {
                return_array.2 - return_array.0 + 1  // rows
            };

            if lookup_size != return_size {
                return Some(EvalResult::Error("XLOOKUP lookup and return arrays must have same size".to_string()));
            }

            // Optional: if_not_found
            let if_not_found = if args.len() >= 4 {
                Some(&args[3])
            } else {
                None
            };

            // Optional: match_mode (0 = exact match, default)
            let match_mode = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 0,
                }
            } else {
                0
            };

            // Optional: search_mode (1 = first to last, default; -1 = last to first).
            //
            // Previously the arity check accepted this argument and nothing ever read
            // it, so XLOOKUP(...,-1) quietly returned the first match instead of the
            // last — the answer a user asked not to get, with no indication the
            // argument had been disregarded.
            //
            // Excel's 2 and -2 select a binary search over data the caller promises is
            // sorted. On sorted input that finds the same row as a scan in the same
            // direction, so they map onto the forward and reverse scans here; on
            // unsorted input Excel's own result is undefined, which is not a behaviour
            // worth reproducing.
            let search_mode = if args.len() >= 6 {
                match evaluate(&args[5], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 1,
                }
            } else {
                1
            };
            let reverse = search_mode < 0;

            // Search for match
            let mut found_idx: Option<usize> = None;

            let xl_get = |sheet: &SheetRef, r: usize, c: usize| -> String {
                match get_text_for_sheet(lookup, sheet, r, c) {
                    Ok(t) => t,
                    Err(_) => String::new(),
                }
            };

            let scan: Box<dyn Iterator<Item = usize>> = if reverse {
                Box::new((0..lookup_size).rev())
            } else {
                Box::new(0..lookup_size)
            };

            for idx in scan {
                let (r, c) = if lookup_array.4 {
                    (lookup_array.0, lookup_array.1 + idx)
                } else {
                    (lookup_array.0 + idx, lookup_array.1)
                };

                let cell_text = xl_get(xl_lookup_sheet, r, c);
                let cell_value = if cell_text.is_empty() {
                    EvalResult::Text(String::new())
                } else if let Ok(n) = cell_text.parse::<f64>() {
                    EvalResult::Number(n)
                } else {
                    EvalResult::Text(cell_text)
                };

                let is_match = match match_mode {
                    0 => {
                        // Exact match
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            (EvalResult::Text(a), EvalResult::Text(b)) => a.eq_ignore_ascii_case(b),
                            _ => false,
                        }
                    }
                    2 => {
                        // The same matcher COUNTIF/SUMIF criteria use, rather than a
                        // second implementation. The previous one took the text before
                        // the first `*` and asked whether the cell started with it,
                        // which made "*eta" an empty prefix — matching every row and
                        // returning the first. A wrong answer, not a missed one.
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Text(pattern), EvalResult::Text(text)) => {
                                wildcard_match(pattern, text)
                            }
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            _ => false,
                        }
                    }
                    _ => {
                        // For -1 (next smaller) and 1 (next larger), fall back to exact for simplicity
                        match (&lookup_value, &cell_value) {
                            (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < 1e-10,
                            (EvalResult::Text(a), EvalResult::Text(b)) => a.eq_ignore_ascii_case(b),
                            _ => false,
                        }
                    }
                };

                if is_match {
                    found_idx = Some(idx);
                    break;
                }
            }

            match found_idx {
                Some(idx) => {
                    // Return value from return_array at same position
                    let (r, c) = if lookup_array.4 {
                        (return_array.0, return_array.1 + idx)
                    } else {
                        (return_array.0 + idx, return_array.1)
                    };

                    let result_text = xl_get(xl_return_sheet, r, c);
                    if result_text.is_empty() {
                        EvalResult::Empty
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => {
                    // Not found - return if_not_found or #N/A
                    if let Some(if_not_found_expr) = if_not_found {
                        evaluate(if_not_found_expr, lookup)
                    } else {
                        EvalResult::Error("#N/A".to_string())
                    }
                }
            }
        }
        "HLOOKUP" => {
            // HLOOKUP(search_key, range, index, [is_sorted])
            if args.len() < 3 || args.len() > 4 {
                return Some(EvalResult::Error("HLOOKUP requires 3 or 4 arguments".to_string()));
            }
            let search_key = evaluate(&args[0], lookup);
            let (range, hlookup_sheet) = match &args[1] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => ((*start_row, *start_col, *end_row, *end_col), sheet),
                _ => return Some(EvalResult::Error("HLOOKUP requires a range as second argument".to_string())),
            };
            let row_index = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let is_sorted = if args.len() == 4 {
                evaluate(&args[3], lookup).to_bool().unwrap_or(true)
            } else {
                true
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_rows = max_row - min_row + 1;

            if row_index > num_rows {
                return Some(EvalResult::Error("#REF!".to_string()));
            }

            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_col: Option<usize> = None;

            let hget = |r: usize, c: usize| -> String {
                match get_text_for_sheet(lookup, hlookup_sheet, r, c) {
                    Ok(t) => t,
                    Err(_) => String::new(),
                }
            };

            if is_sorted {
                if let Some(search_n) = search_num {
                    let mut best_col: Option<usize> = None;
                    let mut best_val = f64::NEG_INFINITY;
                    for c in min_col..=max_col {
                        let cell_text = hget(min_row, c);
                        if let Ok(cell_n) = cell_text.parse::<f64>() {
                            if cell_n <= search_n && cell_n > best_val {
                                best_val = cell_n;
                                best_col = Some(c);
                            }
                        }
                    }
                    found_col = best_col;
                } else {
                    // Text key: find largest text cell <= search_text (case-insensitive)
                    let mut best_col: Option<usize> = None;
                    let mut best_val = String::new();
                    for c in min_col..=max_col {
                        let cell_text = hget(min_row, c);
                        let cell_lower = cell_text.to_lowercase();
                        if cell_text.is_empty() || cell_text.parse::<f64>().is_ok() {
                            continue;
                        }
                        if cell_lower <= search_text && (best_col.is_none() || cell_lower > best_val) {
                            best_val = cell_lower;
                            best_col = Some(c);
                        }
                    }
                    found_col = best_col;
                }
            } else {
                for c in min_col..=max_col {
                    let cell_text = hget(min_row, c);
                    let cell_lower = cell_text.to_lowercase();

                    if let (Some(search_n), Ok(cell_n)) = (search_num, cell_text.parse::<f64>()) {
                        if (search_n - cell_n).abs() < f64::EPSILON {
                            found_col = Some(c);
                            break;
                        }
                    } else if cell_lower == search_text {
                        found_col = Some(c);
                        break;
                    }
                }
            }

            match found_col {
                Some(c) => {
                    let result_row = min_row + row_index - 1;
                    let result_text = hget(result_row, c);
                    if result_text.is_empty() {
                        EvalResult::Empty
                    } else if let Ok(n) = result_text.parse::<f64>() {
                        EvalResult::Number(n)
                    } else {
                        EvalResult::Text(result_text)
                    }
                }
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "INDEX" => {
            // INDEX(range, row_num, [col_num])
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("INDEX requires 2 or 3 arguments".to_string()));
            }
            let (range, idx_sheet) = match &args[0] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => ((*start_row, *start_col, *end_row, *end_col), sheet.clone()),
                Expr::CellRef { sheet, col, row, .. } => ((*row, *col, *row, *col), sheet.clone()),
                _ => return Some(EvalResult::Error("INDEX requires a range as first argument".to_string())),
            };
            let row_num = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let col_num = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));
            let num_rows = max_row - min_row + 1;
            let num_cols = max_col - min_col + 1;

            if row_num < 1 || row_num > num_rows || col_num < 1 || col_num > num_cols {
                return Some(EvalResult::Error("#REF!".to_string()));
            }

            let target_row = min_row + row_num - 1;
            let target_col = min_col + col_num - 1;
            let result_text = match get_text_for_sheet(lookup, &idx_sheet, target_row, target_col) {
                Ok(t) => t,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            if result_text.is_empty() {
                EvalResult::Empty
            } else if let Ok(n) = result_text.parse::<f64>() {
                EvalResult::Number(n)
            } else {
                EvalResult::Text(result_text)
            }
        }
        "MATCH" => {
            // MATCH(search_key, range, [match_type])
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("MATCH requires 2 or 3 arguments".to_string()));
            }
            let search_key = evaluate(&args[0], lookup);
            let (range, match_sheet) = match &args[1] {
                Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => ((*start_row, *start_col, *end_row, *end_col), sheet),
                _ => return Some(EvalResult::Error("MATCH requires a range as second argument".to_string())),
            };
            let match_type = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 1,
                }
            } else {
                1
            };

            let (min_row, min_col, max_row, max_col) = (range.0.min(range.2), range.1.min(range.3), range.0.max(range.2), range.1.max(range.3));

            // Determine if it's a row vector or column vector
            let is_row = min_row == max_row;
            let search_text = search_key.to_text().to_lowercase();
            let search_num = search_key.to_number().ok();

            let mut found_pos: Option<usize> = None;

            let mget = |r: usize, c: usize| -> String {
                match get_text_for_sheet(lookup, match_sheet, r, c) {
                    Ok(t) => t,
                    Err(_) => String::new(),
                }
            };

            if is_row {
                // Search horizontally
                if match_type == 0 {
                    // Exact match
                    for (i, c) in (min_col..=max_col).enumerate() {
                        let cell_text = mget(min_row, c);
                        let cell_lower = cell_text.to_lowercase();
                        if let (Some(sn), Ok(cn)) = (search_num, cell_text.parse::<f64>()) {
                            if (sn - cn).abs() < f64::EPSILON {
                                found_pos = Some(i + 1);
                                break;
                            }
                        } else if cell_lower == search_text {
                            found_pos = Some(i + 1);
                            break;
                        }
                    }
                } else if match_type == 1 {
                    // Largest value <= search_key
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::NEG_INFINITY;
                        for (i, c) in (min_col..=max_col).enumerate() {
                            if let Ok(cn) = mget(min_row, c).parse::<f64>() {
                                if cn <= sn && cn > best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                } else {
                    // Smallest value >= search_key
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::INFINITY;
                        for (i, c) in (min_col..=max_col).enumerate() {
                            if let Ok(cn) = mget(min_row, c).parse::<f64>() {
                                if cn >= sn && cn < best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                }
            } else {
                // Search vertically
                if match_type == 0 {
                    for (i, r) in (min_row..=max_row).enumerate() {
                        let cell_text = mget(r, min_col);
                        let cell_lower = cell_text.to_lowercase();
                        if let (Some(sn), Ok(cn)) = (search_num, cell_text.parse::<f64>()) {
                            if (sn - cn).abs() < f64::EPSILON {
                                found_pos = Some(i + 1);
                                break;
                            }
                        } else if cell_lower == search_text {
                            found_pos = Some(i + 1);
                            break;
                        }
                    }
                } else if match_type == 1 {
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::NEG_INFINITY;
                        for (i, r) in (min_row..=max_row).enumerate() {
                            if let Ok(cn) = mget(r, min_col).parse::<f64>() {
                                if cn <= sn && cn > best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                } else {
                    if let Some(sn) = search_num {
                        let mut best_pos: Option<usize> = None;
                        let mut best_val = f64::INFINITY;
                        for (i, r) in (min_row..=max_row).enumerate() {
                            if let Ok(cn) = mget(r, min_col).parse::<f64>() {
                                if cn >= sn && cn < best_val {
                                    best_val = cn;
                                    best_pos = Some(i + 1);
                                }
                            }
                        }
                        found_pos = best_pos;
                    }
                }
            }

            match found_pos {
                Some(pos) => EvalResult::Number(pos as f64),
                None => EvalResult::Error("#N/A".to_string()),
            }
        }
        "ROW" => {
            // ROW([reference]) - returns the row number
            if args.is_empty() {
                // No argument: return row of current cell being evaluated
                return Some(match lookup.current_cell() {
                    Some((row, _)) => EvalResult::Number((row + 1) as f64),
                    None => EvalResult::Error("ROW() requires cell context".to_string()),
                });
            }
            if args.len() != 1 {
                return Some(EvalResult::Error("ROW requires 0 or 1 argument".to_string()));
            }
            match &args[0] {
                Expr::CellRef { row, .. } => EvalResult::Number((*row + 1) as f64),
                Expr::Range { start_row, .. } => EvalResult::Number((*start_row + 1) as f64),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "COLUMN" => {
            // COLUMN([reference]) - returns the column number
            if args.is_empty() {
                // No argument: return column of current cell being evaluated
                return Some(match lookup.current_cell() {
                    Some((_, col)) => EvalResult::Number((col + 1) as f64),
                    None => EvalResult::Error("COLUMN() requires cell context".to_string()),
                });
            }
            if args.len() != 1 {
                return Some(EvalResult::Error("COLUMN requires 0 or 1 argument".to_string()));
            }
            match &args[0] {
                Expr::CellRef { col, .. } => EvalResult::Number((*col + 1) as f64),
                Expr::Range { start_col, .. } => EvalResult::Number((*start_col + 1) as f64),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "ROWS" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ROWS requires exactly one argument".to_string()));
            }
            match &args[0] {
                Expr::Range { start_row, end_row, .. } => {
                    EvalResult::Number((end_row.max(start_row) - end_row.min(start_row) + 1) as f64)
                }
                Expr::CellRef { .. } => EvalResult::Number(1.0),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "COLUMNS" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("COLUMNS requires exactly one argument".to_string()));
            }
            match &args[0] {
                Expr::Range { start_col, end_col, .. } => {
                    EvalResult::Number((end_col.max(start_col) - end_col.min(start_col) + 1) as f64)
                }
                Expr::CellRef { .. } => EvalResult::Number(1.0),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "OFFSET" => {
            // OFFSET(reference, rows, cols, [height], [width]) — returns a range shifted
            // from `reference`, optionally resized. Reads the target and returns a scalar
            // for a 1x1 result or an Array for a multi-cell result.
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error("OFFSET requires 3 to 5 arguments".to_string()));
            }
            let (base_row, base_col, base_h, base_w, sheet): (usize, usize, usize, usize, SheetRef) =
                match &args[0] {
                    Expr::CellRef { sheet, row, col, .. } => (*row, *col, 1, 1, sheet.clone()),
                    Expr::Range { sheet, start_row, start_col, end_row, end_col, .. } => {
                        let sr = (*start_row).min(*end_row);
                        let sc = (*start_col).min(*end_col);
                        let h = (*start_row).max(*end_row) - sr + 1;
                        let w = (*start_col).max(*end_col) - sc + 1;
                        (sr, sc, h, w, sheet.clone())
                    }
                    _ => {
                        return Some(EvalResult::Error(
                            "OFFSET requires a cell or range reference".to_string(),
                        ))
                    }
                };
            let rows_off = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.trunc() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let cols_off = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n.trunc() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let height = match args.get(3) {
                Some(a) => match evaluate(a, lookup).to_number() {
                    Ok(n) if n >= 1.0 => n.trunc() as usize,
                    Ok(_) => return Some(EvalResult::Error("#REF!".to_string())),
                    Err(e) => return Some(EvalResult::Error(e)),
                },
                None => base_h,
            };
            let width = match args.get(4) {
                Some(a) => match evaluate(a, lookup).to_number() {
                    Ok(n) if n >= 1.0 => n.trunc() as usize,
                    Ok(_) => return Some(EvalResult::Error("#REF!".to_string())),
                    Err(e) => return Some(EvalResult::Error(e)),
                },
                None => base_w,
            };
            let new_row = base_row as i64 + rows_off;
            let new_col = base_col as i64 + cols_off;
            if new_row < 0 || new_col < 0 {
                return Some(EvalResult::Error("#REF!".to_string()));
            }
            range_to_result(lookup, &sheet, new_row as usize, new_col as usize, height, width)
        }
        "INDIRECT" => {
            // INDIRECT(ref_text, [a1]) — resolves an A1-style text reference on the current
            // sheet. Cross-sheet ("Sheet!A1") and R1C1 style are not yet supported (#REF!).
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("INDIRECT requires 1 or 2 arguments".to_string()));
            }
            let ref_text = evaluate(&args[0], lookup).to_text();
            match parse_a1_ref(&ref_text) {
                Some((sr, sc, er, ec)) => {
                    range_to_result(lookup, &SheetRef::Current, sr, sc, er - sr + 1, ec - sc + 1)
                }
                None => EvalResult::Error("#REF!".to_string()),
            }
        }
        _ => return None,
    };
    Some(result)
}

/// Read a single cell's typed value, honoring the sheet the reference points at.
fn read_cell_value<L: CellLookup>(lookup: &L, sheet: &SheetRef, row: usize, col: usize) -> Value {
    match sheet {
        SheetRef::Current => lookup.get_cell_value(row, col),
        SheetRef::Id(id) => lookup.get_value_sheet(*id, row, col),
        SheetRef::RefError { .. } => Value::Empty,
    }
}

/// Materialize a rectangular region into an EvalResult: a scalar for 1x1, else an Array.
fn range_to_result<L: CellLookup>(
    lookup: &L,
    sheet: &SheetRef,
    start_row: usize,
    start_col: usize,
    height: usize,
    width: usize,
) -> EvalResult {
    if height == 0 || width == 0 {
        return EvalResult::Error("#REF!".to_string());
    }
    if height == 1 && width == 1 {
        return EvalResult::from_value(&read_cell_value(lookup, sheet, start_row, start_col));
    }
    let mut rows_vec: Vec<Vec<Value>> = Vec::with_capacity(height);
    for r in 0..height {
        let mut row_vec = Vec::with_capacity(width);
        for c in 0..width {
            row_vec.push(read_cell_value(lookup, sheet, start_row + r, start_col + c));
        }
        rows_vec.push(row_vec);
    }
    EvalResult::Array(Array2D::from_vec(rows_vec))
}

/// Parse a single A1 cell reference like "A1" or "$B$2" into 0-based (row, col).
fn parse_a1_cell(s: &str) -> Option<(usize, usize)> {
    let clean: String = s.trim().chars().filter(|c| *c != '$').collect();
    let bytes = clean.as_bytes();
    let mut split = 0;
    while split < bytes.len() && bytes[split].is_ascii_alphabetic() {
        split += 1;
    }
    if split == 0 || split == bytes.len() {
        return None;
    }
    let (col_part, row_part) = clean.split_at(split);
    if !row_part.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let mut col: usize = 0;
    for c in col_part.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let col = col.checked_sub(1)?;
    let row = row_part.parse::<usize>().ok()?.checked_sub(1)?;
    Some((row, col))
}

/// Parse an A1 reference ("A1" or "A1:B3") into a normalized 0-based
/// (start_row, start_col, end_row, end_col). Returns None for cross-sheet or malformed refs.
fn parse_a1_ref(s: &str) -> Option<(usize, usize, usize, usize)> {
    let s = s.trim();
    if s.is_empty() || s.contains('!') {
        return None;
    }
    match s.split_once(':') {
        Some((a, b)) => {
            let (r1, c1) = parse_a1_cell(a)?;
            let (r2, c2) = parse_a1_cell(b)?;
            Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
        }
        None => {
            let (r, c) = parse_a1_cell(s)?;
            Some((r, c, r, c))
        }
    }
}
