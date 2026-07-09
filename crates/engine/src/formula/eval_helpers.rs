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

/// Try to parse a date string and return Excel serial number.
/// Supports formats:
/// - ISO: "2023-11-07", "2023/11/07"
/// - US: "11/07/2023", "11-07-2023"
/// Returns None if the string doesn't look like a date.
pub fn try_parse_date_string(s: &str) -> Option<f64> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    // Try ISO format: YYYY-MM-DD or YYYY/MM/DD
    if s.len() >= 8 && s.len() <= 10 {
        // Check for ISO pattern (starts with 4-digit year)
        if let Some(sep_pos) = s.find(|c| c == '-' || c == '/') {
            if sep_pos == 4 {
                let sep = s.chars().nth(sep_pos).unwrap();
                let parts: Vec<&str> = s.split(sep).collect();
                if parts.len() == 3 {
                    if let (Ok(year), Ok(month), Ok(day)) = (
                        parts[0].parse::<i32>(),
                        parts[1].parse::<i32>(),
                        parts[2].parse::<i32>(),
                    ) {
                        if month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 9999 {
                            return Some(date_to_serial(year, month, day));
                        }
                    }
                }
            }
        }

        // Try US format: MM/DD/YYYY or MM-DD-YYYY
        if let Some(sep) = s.chars().find(|&c| c == '/' || c == '-') {
            let parts: Vec<&str> = s.split(sep).collect();
            if parts.len() == 3 && parts[2].len() == 4 {
                if let (Ok(month), Ok(day), Ok(year)) = (
                    parts[0].parse::<i32>(),
                    parts[1].parse::<i32>(),
                    parts[2].parse::<i32>(),
                ) {
                    if month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 9999 {
                        return Some(date_to_serial(year, month, day));
                    }
                }
            }
        }
    }

    None
}

/// Compare two floats for spreadsheet equality.
///
/// Uses a magnitude-relative tolerance rather than a bare `f64::EPSILON` absolute check.
/// `f64::EPSILON` (~2.2e-16) is only meaningful near 1.0: for large values `=` would never
/// hold (rounding error exceeds it) and near zero it is needlessly loose. A relative
/// tolerance tracks the ~15 significant figures a spreadsheet treats as "equal".
pub(crate) fn approx_eq(a: f64, b: f64) -> bool {
    if a == b {
        return true;
    }
    let scale = a.abs().max(b.abs());
    (a - b).abs() <= (scale * 1e-12).max(f64::EPSILON)
}

/// Case-insensitive text match supporting Excel-style wildcards: `*` matches any run of
/// characters, `?` matches exactly one, and `~` escapes the following `*`/`?`/`~` to a
/// literal. With no wildcard characters this is a plain case-insensitive equality.
pub(crate) fn wildcard_match(pattern: &str, text: &str) -> bool {
    if !pattern.contains(['*', '?', '~']) {
        return pattern.eq_ignore_ascii_case(text);
    }

    enum Tok {
        Star,
        AnyOne,
        Lit(char),
    }
    let pchars: Vec<char> = pattern.to_lowercase().chars().collect();
    let mut toks: Vec<Tok> = Vec::with_capacity(pchars.len());
    let mut i = 0;
    while i < pchars.len() {
        match pchars[i] {
            '~' if i + 1 < pchars.len() => {
                toks.push(Tok::Lit(pchars[i + 1]));
                i += 2;
            }
            '*' => {
                toks.push(Tok::Star);
                i += 1;
            }
            '?' => {
                toks.push(Tok::AnyOne);
                i += 1;
            }
            c => {
                toks.push(Tok::Lit(c));
                i += 1;
            }
        }
    }

    let t: Vec<char> = text.to_lowercase().chars().collect();
    let (mut pi, mut ti) = (0usize, 0usize);
    // Backtracking anchor for the most recent `*`.
    let (mut star_pi, mut star_ti): (Option<usize>, usize) = (None, 0);
    while ti < t.len() {
        match toks.get(pi) {
            Some(Tok::Star) => {
                star_pi = Some(pi);
                star_ti = ti;
                pi += 1;
            }
            Some(Tok::AnyOne) => {
                pi += 1;
                ti += 1;
            }
            Some(Tok::Lit(c)) if *c == t[ti] => {
                pi += 1;
                ti += 1;
            }
            _ => match star_pi {
                Some(sp) => {
                    pi = sp + 1;
                    star_ti += 1;
                    ti = star_ti;
                }
                None => return false,
            },
        }
    }
    while matches!(toks.get(pi), Some(Tok::Star)) {
        pi += 1;
    }
    pi == toks.len()
}

/// Check if a cell value matches criteria (for SUMIF, COUNTIF, etc.)
pub(crate) fn matches_criteria(value: &EvalResult, criteria: &EvalResult) -> bool {
    let criteria_str = criteria.to_text();

    // Check for comparison operators in criteria
    if let Some(rest) = criteria_str.strip_prefix(">=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), rest.trim().parse::<f64>()) {
            return v >= c;
        }
    } else if let Some(rest) = criteria_str.strip_prefix("<=") {
        if let (Ok(v), Ok(c)) = (value.to_number(), rest.trim().parse::<f64>()) {
            return v <= c;
        }
    } else if let Some(rest) = criteria_str.strip_prefix("<>") {
        let c = rest.trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return !approx_eq(v, n);
            }
        }
        return !wildcard_match(c, &value.to_text());
    } else if let Some(rest) = criteria_str.strip_prefix('>') {
        if let (Ok(v), Ok(c)) = (value.to_number(), rest.trim().parse::<f64>()) {
            return v > c;
        }
    } else if let Some(rest) = criteria_str.strip_prefix('<') {
        if let (Ok(v), Ok(c)) = (value.to_number(), rest.trim().parse::<f64>()) {
            return v < c;
        }
    } else if let Some(rest) = criteria_str.strip_prefix('=') {
        let c = rest.trim();
        if let Ok(n) = c.parse::<f64>() {
            if let Ok(v) = value.to_number() {
                return approx_eq(v, n);
            }
        }
        return wildcard_match(c, &value.to_text());
    }

    // No operator: numeric equality when both sides are numeric, else wildcard-aware text.
    if let (Ok(v), Ok(c)) = (value.to_number(), criteria_str.parse::<f64>()) {
        return approx_eq(v, c);
    }
    wildcard_match(&criteria_str, &value.to_text())
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
                // A function may return a range as an Array (e.g. OFFSET/INDIRECT over a
                // multi-cell region). Flatten its numeric cells like a Range arg does,
                // skipping text/blanks; scalars keep the strict number-or-error behavior.
                match evaluate(arg, lookup) {
                    EvalResult::Array(arr) => {
                        for r in 0..arr.rows() {
                            for c in 0..arr.cols() {
                                if let Some(v) = arr.get(r, c) {
                                    if let Ok(n) = EvalResult::from_value(v).to_number() {
                                        values.push(n);
                                    }
                                }
                            }
                        }
                    }
                    result => match result.to_number() {
                        Ok(n) => values.push(n),
                        Err(e) => return Err(e),
                    },
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn approx_eq_holds_at_large_magnitude() {
        // The old `(a-b).abs() < f64::EPSILON` check could never hold for large values
        // because rounding error exceeds ~2.2e-16. Relative tolerance fixes that.
        let a = 1_000_000_000.0_f64;
        // Rounding-level difference at large magnitude must compare equal — impossible
        // under the old `abs() < f64::EPSILON` (2.2e-16) check.
        assert!(approx_eq(a, a + 1e-4), "rounding-level diff at 1e9 should compare equal");
        assert!(approx_eq(0.1 + 0.2, 0.3), "0.1+0.2 should equal 0.3");
        assert!(approx_eq(0.0, 0.0));
        // Genuinely different values must still be unequal at every magnitude.
        assert!(!approx_eq(1.0, 1.0001), "small clearly-different values stay unequal");
        assert!(!approx_eq(a, a + 100.0), "large clearly-different values stay unequal");
        assert!(!approx_eq(1_000_000_000.0, 1_000_100_000.0));
    }

    #[test]
    fn matches_criteria_numeric_and_wildcards() {
        let num = |n: f64| EvalResult::Number(n);
        let txt = |s: &str| EvalResult::Text(s.to_string());

        // Numeric operators
        assert!(matches_criteria(&num(5.0), &txt(">=5")));
        assert!(matches_criteria(&num(5.0), &txt("=5")));
        assert!(!matches_criteria(&num(4.0), &txt(">5")));
        assert!(matches_criteria(&num(4.0), &txt("<>5")));

        // Wildcards: * (any run) and ? (single char), case-insensitive
        assert!(matches_criteria(&txt("Apple"), &txt("app*")));
        assert!(matches_criteria(&txt("banana"), &txt("*ana*")));
        assert!(matches_criteria(&txt("cat"), &txt("c?t")));
        assert!(!matches_criteria(&txt("cart"), &txt("c?t")));
        assert!(!matches_criteria(&txt("Apple"), &txt("*z*")));
        // <> with a wildcard = "not matching"
        assert!(matches_criteria(&txt("Apple"), &txt("<>*z*")));
        assert!(!matches_criteria(&txt("Apple"), &txt("<>app*")));
        // ~ escapes a wildcard to a literal
        assert!(matches_criteria(&txt("a*b"), &txt("a~*b")));
        assert!(!matches_criteria(&txt("axb"), &txt("a~*b")));
    }

    #[test]
    fn test_try_parse_date_string_iso() {
        // ISO format: YYYY-MM-DD
        let serial = try_parse_date_string("2023-11-07").unwrap();
        let (y, m, d) = serial_to_date(serial);
        assert_eq!((y, m, d), (2023, 11, 7));

        // ISO format with slashes
        let serial = try_parse_date_string("2024/08/29").unwrap();
        let (y, m, d) = serial_to_date(serial);
        assert_eq!((y, m, d), (2024, 8, 29));
    }

    #[test]
    fn test_try_parse_date_string_us() {
        // US format: MM/DD/YYYY
        let serial = try_parse_date_string("11/07/2023").unwrap();
        let (y, m, d) = serial_to_date(serial);
        assert_eq!((y, m, d), (2023, 11, 7));

        // US format with dashes
        let serial = try_parse_date_string("08-29-2024").unwrap();
        let (y, m, d) = serial_to_date(serial);
        assert_eq!((y, m, d), (2024, 8, 29));
    }

    #[test]
    fn test_try_parse_date_string_invalid() {
        assert!(try_parse_date_string("hello").is_none());
        assert!(try_parse_date_string("123").is_none());
        assert!(try_parse_date_string("").is_none());
        assert!(try_parse_date_string("13/01/2023").is_none()); // Invalid month for US format when year is last
    }

    #[test]
    fn test_date_subtraction() {
        // Test that date subtraction gives correct day count
        let date1 = try_parse_date_string("2023-11-07").unwrap();
        let date2 = try_parse_date_string("2024-08-29").unwrap();
        let days = date2 - date1;
        assert_eq!(days as i32, 296); // 296 days between these dates
    }
}
