// Shared helper functions for formula evaluation

use crate::sheet::SheetRef;
use super::eval::{evaluate, CellLookup, EvalResult, Value, NamedRangeResolution};
use super::parser::{BoundExpr, Expr};

/// Compare two Values for sorting
/// Order: Numbers < Text < Empty < Errors (ascending)
pub(crate) fn value_compare(a: &Value, b: &Value) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    // Type ordering: Number=0, Text=1, Boolean=2, Empty=3, Error=4
    fn type_rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) => 0,
            Value::Text(_) => 1,
            Value::Boolean(_) => 2,
            Value::Empty => 3,
            Value::Error(_) => 4,
        }
    }

    let rank_a = type_rank(a);
    let rank_b = type_rank(b);

    if rank_a != rank_b {
        return rank_a.cmp(&rank_b);
    }

    // Same type - compare within type
    match (a, b) {
        (Value::Number(na), Value::Number(nb)) => {
            na.partial_cmp(nb).unwrap_or(Ordering::Equal)
        }
        (Value::Text(sa), Value::Text(sb)) => {
            // Case-insensitive comparison
            sa.to_lowercase().cmp(&sb.to_lowercase())
        }
        (Value::Boolean(ba), Value::Boolean(bb)) => {
            // FALSE < TRUE
            ba.cmp(bb)
        }
        (Value::Error(ea), Value::Error(eb)) => {
            ea.cmp(eb)
        }
        _ => Ordering::Equal, // Empty == Empty
    }
}

// Helper functions for date calculations

/// Convert year/month/day to Excel serial date number
pub(crate) fn date_to_serial(year: i32, month: i32, day: i32) -> f64 {
    // Handle month overflow/underflow
    let mut y = year;
    let mut m = month;
    while m > 12 {
        m -= 12;
        y += 1;
    }
    while m < 1 {
        m += 12;
        y -= 1;
    }

    // Calculate days since Excel epoch (1899-12-30)
    // Using a simplified algorithm
    let a = (14 - m) / 12;
    let y_adj = y + 4800 - a;
    let m_adj = m + 12 * a - 3;

    let jdn = day + (153 * m_adj + 2) / 5 + 365 * y_adj + y_adj / 4 - y_adj / 100 + y_adj / 400 - 32045;

    // Excel epoch JDN (1899-12-30) = 2415019
    (jdn - 2415019) as f64
}

/// Convert Excel serial date number to year/month/day
pub(crate) fn serial_to_date(serial: f64) -> (i32, i32, i32) {
    let serial = serial.floor() as i32;
    let jdn = serial + 2415019; // Convert to Julian Day Number

    // Algorithm to convert JDN to Gregorian date
    let a = jdn + 32044;
    let b = (4 * a + 3) / 146097;
    let c = a - (146097 * b) / 4;
    let d = (4 * c + 3) / 1461;
    let e = c - (1461 * d) / 4;
    let m = (5 * e + 2) / 153;

    let day = e - (153 * m + 2) / 5 + 1;
    let month = m + 3 - 12 * (m / 10);
    let year = 100 * b + d - 4800 + m / 10;

    (year, month, day)
}

/// Get the number of days in a month
pub(crate) fn days_in_month(year: i32, month: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0) {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

/// Check if a cell value matches criteria (for SUMIF, COUNTIF, etc.)
pub(crate) fn matches_criteria(value: &EvalResult, criteria: &EvalResult) -> bool {
    let criteria_str = criteria.to_text();

    // Check for comparison operators in criteria
    if criteria_str.starts_with(">=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[2..].trim().parse::<f64>()) {
            return v >= c;
        }
    } else if criteria_str.starts_with("<=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[2..].trim().parse::<f64>()) {
            return v <= c;
        }
    } else if criteria_str.starts_with("<>") {
        let c = criteria_str[2..].trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return (v - n).abs() >= f64::EPSILON;
            }
        }
        return value.to_text().to_lowercase() != c.to_lowercase();
    } else if criteria_str.starts_with('>') {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[1..].trim().parse::<f64>()) {
            return v > c;
        }
    } else if criteria_str.starts_with('<') {
        if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str[1..].trim().parse::<f64>()) {
            return v < c;
        }
    } else if criteria_str.starts_with('=') {
        let c = criteria_str[1..].trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return (v - n).abs() < f64::EPSILON;
            }
        }
        return value.to_text().to_lowercase() == c.to_lowercase();
    }

    // Simple equality check
    match (value, criteria) {
        (EvalResult::Number(v), EvalResult::Number(c)) => (v - c).abs() < f64::EPSILON,
        (EvalResult::Text(v), EvalResult::Text(c)) => v.to_lowercase() == c.to_lowercase(),
        _ => value.to_text().to_lowercase() == criteria_str.to_lowercase(),
    }
}

/// Helper to get text from a cell, handling cross-sheet references
pub(crate) fn get_text_for_sheet<L: CellLookup>(lookup: &L, sheet: &SheetRef, row: usize, col: usize) -> Result<String, String> {
    match sheet {
        SheetRef::Current => Ok(lookup.get_text(row, col)),
        SheetRef::Id(id) => Ok(lookup.get_text_sheet(*id, row, col)),
        SheetRef::RefError { .. } => Err("#REF!".to_string()),
    }
}

pub(crate) fn collect_numbers<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Result<Vec<f64>, String> {
    let mut values = Vec::new();

    for arg in args {
        match arg {
            Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                collect_numbers_from_range_sheet(sheet, *start_row, *start_col, *end_row, *end_col, lookup, &mut values)?;
            }
            Expr::NamedRange(name) => {
                // Resolve named range and collect numbers from it
                match lookup.resolve_named_range(name) {
                    None => return Err(format!("#NAME? '{}'", name)),
                    Some(NamedRangeResolution::Cell { row, col }) => {
                        let text = lookup.get_text(row, col);
                        if let Ok(n) = text.parse::<f64>() {
                            values.push(n);
                        }
                    }
                    Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                        collect_numbers_from_range(start_row, start_col, end_row, end_col, lookup, &mut values);
                    }
                }
            }
            _ => {
                let result = evaluate(arg, lookup);
                match result.to_number() {
                    Ok(n) => values.push(n),
                    Err(e) => return Err(e),
                }
            }
        }
    }

    Ok(values)
}

/// Collect numbers from a range, supporting cross-sheet references
pub(crate) fn collect_numbers_from_range_sheet<L: CellLookup>(
    sheet: &SheetRef,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<f64>,
) -> Result<(), String> {
    // Check for RefError early
    if let SheetRef::RefError { .. } = sheet {
        return Err("#REF!".to_string());
    }

    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);

    for r in min_row..=max_row {
        for c in min_col..=max_col {
            let text = get_text_for_sheet(lookup, sheet, r, c)?;
            // Only include numeric values, skip text/empty
            if let Ok(n) = text.parse::<f64>() {
                values.push(n);
            }
        }
    }
    Ok(())
}

/// Legacy helper for same-sheet ranges (used by named range resolution)
fn collect_numbers_from_range<L: CellLookup>(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<f64>,
) {
    let _ = collect_numbers_from_range_sheet(
        &SheetRef::Current,
        start_row, start_col, end_row, end_col,
        lookup, values
    );
}

pub(crate) fn collect_all_values<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Vec<EvalResult> {
    let mut values = Vec::new();

    for arg in args {
        match arg {
            Expr::Range { sheet, start_col, start_row, end_col, end_row, .. } => {
                if let Err(e) = collect_all_values_from_range_sheet(sheet, *start_row, *start_col, *end_row, *end_col, lookup, &mut values) {
                    values.push(EvalResult::Error(e));
                }
            }
            Expr::NamedRange(name) => {
                // Resolve named range and collect all values from it
                match lookup.resolve_named_range(name) {
                    None => values.push(EvalResult::Error(format!("#NAME? '{}'", name))),
                    Some(NamedRangeResolution::Cell { row, col }) => {
                        let text = lookup.get_text(row, col);
                        if text.is_empty() {
                            values.push(EvalResult::Text(String::new()));
                        } else if let Ok(n) = text.parse::<f64>() {
                            values.push(EvalResult::Number(n));
                        } else {
                            values.push(EvalResult::Text(text));
                        }
                    }
                    Some(NamedRangeResolution::Range { start_row, start_col, end_row, end_col }) => {
                        collect_all_values_from_range(start_row, start_col, end_row, end_col, lookup, &mut values);
                    }
                }
            }
            _ => {
                values.push(evaluate(arg, lookup));
            }
        }
    }

    values
}

/// Collect all values from a range, supporting cross-sheet references
fn collect_all_values_from_range_sheet<L: CellLookup>(
    sheet: &SheetRef,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<EvalResult>,
) -> Result<(), String> {
    // Check for RefError early
    if let SheetRef::RefError { .. } = sheet {
        return Err("#REF!".to_string());
    }

    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let min_col = start_col.min(end_col);
    let max_col = start_col.max(end_col);

    for r in min_row..=max_row {
        for c in min_col..=max_col {
            let text = get_text_for_sheet(lookup, sheet, r, c)?;
            if text.is_empty() {
                values.push(EvalResult::Text(String::new()));
            } else if let Ok(n) = text.parse::<f64>() {
                values.push(EvalResult::Number(n));
            } else {
                values.push(EvalResult::Text(text));
            }
        }
    }
    Ok(())
}

/// Legacy helper for same-sheet ranges (used by named range resolution)
fn collect_all_values_from_range<L: CellLookup>(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    lookup: &L,
    values: &mut Vec<EvalResult>,
) {
    let _ = collect_all_values_from_range_sheet(
        &SheetRef::Current,
        start_row, start_col, end_row, end_col,
        lookup, values
    );
}
