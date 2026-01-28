//! Series Fill: Pattern detection and generation for fill handle
//!
//! Implements Excel-compatible series fill behavior:
//! - Single numbers: copy by default, Ctrl+drag for series
//! - Built-in lists (months, weekdays, quarters): series by default
//! - Two+ cells: detect linear step pattern
//! - Alphanumeric: trailing numbers/letters with overflow
//!
//! See: docs/features/series-fill-spec.md

use visigrid_engine::formula::eval::Value;

// ============================================================================
// Core Types
// ============================================================================

/// User intent derived from gesture (Ctrl held, selection size, pattern type)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillIntent {
    /// Copy source cells literally (no pattern extension)
    Copy,
    /// Extend pattern as series
    Series,
}

/// Detected fill pattern
#[derive(Debug, Clone, PartialEq)]
pub enum FillPattern {
    /// No pattern detected, copy source cells literally
    Copy,
    /// Repeat source cells as a cycle (for multi-cell with non-constant step)
    Repeat { values: Vec<Value> },
    /// Linear numeric series
    Linear { start: f64, step: f64 },
    /// Built-in list (months, weekdays, quarters)
    TextList {
        list: BuiltinList,
        start_idx: i32,
        step: i32,
        case: CaseMode,
    },
    /// Prefix + trailing number (Item1, Row-5, etc.)
    AlphaNum {
        prefix: String,
        suffix: AlphaSuffix,
        start: i64,
        step: i64,
        width: Option<usize>, // For zero-padding: "001" → width=3
    },
    /// Quarter with optional year (Q1, Q1 2026)
    QuarterYear {
        quarter: i32,
        year: Option<i32>,
        step: i32,
    },
}

/// Trailing suffix type for alphanumeric patterns
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AlphaSuffix {
    /// Trailing number: Item1 → Item2
    Number,
    /// Trailing letter: Row A → Row B → ... → Row Z → Row AA
    Letter,
}

/// Case preservation mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaseMode {
    Upper, // JAN → FEB
    Lower, // jan → feb
    Title, // Jan → Feb
}

/// Built-in list types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuiltinList {
    MonthsShort,
    MonthsLong,
    WeekdaysShort,
    WeekdaysLong,
    Quarters,
}

/// Source data for pattern detection
#[derive(Debug, Clone)]
pub struct DetectedSource {
    /// Typed values from source cells
    pub values: Vec<Value>,
    /// Original text tokens (only for Value::Text cells, otherwise None)
    pub text_tokens: Vec<Option<String>>,
    // Note: formats would go here but not needed for V1 pattern detection
}

// ============================================================================
// Intent Detection
// ============================================================================

/// Determine fill intent from user gesture
///
/// # Arguments
/// * `ctrl_held` - Whether Ctrl key is pressed
/// * `selection_len` - Number of cells in source selection
/// * `single_is_list` - Whether single cell matches a built-in list
///
/// # Returns
/// `FillIntent::Series` or `FillIntent::Copy` based on Excel-default rules
pub fn fill_intent(ctrl_held: bool, selection_len: usize, single_is_list: bool) -> FillIntent {
    match (ctrl_held, selection_len, single_is_list) {
        // Single number: copy by default, Ctrl for series
        (false, 1, false) => FillIntent::Copy,
        (true, 1, false) => FillIntent::Series,
        // Single list item: series by default, Ctrl for copy
        (false, 1, true) => FillIntent::Series,
        (true, 1, true) => FillIntent::Copy,
        // Multi-cell: series by default, Ctrl for copy
        (false, _, _) => FillIntent::Series,
        (true, _, _) => FillIntent::Copy,
    }
}

// ============================================================================
// Pattern Detection
// ============================================================================

/// Detect pattern from source values
///
/// # Arguments
/// * `source` - Source cell data
/// * `intent` - User intent (Copy or Series)
///
/// # Returns
/// Detected `FillPattern` for generating new values
pub fn detect_pattern(source: &DetectedSource, intent: FillIntent) -> FillPattern {
    if intent == FillIntent::Copy {
        return FillPattern::Copy;
    }

    match source.values.len() {
        0 => FillPattern::Copy,
        1 => detect_single_value_pattern(&source.values[0], source.text_tokens.first().and_then(|t| t.as_deref())),
        2 => detect_two_value_pattern(source),
        _ => detect_multi_value_pattern(source),
    }
}

/// Detect pattern from a single value
fn detect_single_value_pattern(value: &Value, _text_token: Option<&str>) -> FillPattern {
    let text = match value {
        Value::Text(s) => s.as_str(),
        Value::Number(n) => {
            // Single number with series intent → Linear step 1
            return FillPattern::Linear { start: *n, step: 1.0 };
        }
        _ => return FillPattern::Copy,
    };

    // Skip formulas - they go through existing ref-adjustment path
    if text.starts_with('=') {
        return FillPattern::Copy;
    }

    // Check quarter/year pattern FIRST (more specific than built-in list)
    if let Some((q, year)) = parse_quarter(text) {
        return FillPattern::QuarterYear {
            quarter: q,
            year,
            step: 1,
        };
    }

    // Check other built-in lists (months, weekdays - but not quarters since handled above)
    if let Some((list, index, case)) = find_in_builtin_list(text) {
        // Skip if it's a quarter (already handled by QuarterYear)
        if list != BuiltinList::Quarters {
            return FillPattern::TextList {
                list,
                start_idx: index,
                step: 1,
                case,
            };
        }
    }

    // Check alphanumeric with trailing number (supports negatives)
    if let Some((prefix, num, width)) = extract_trailing_number(text) {
        return FillPattern::AlphaNum {
            prefix,
            suffix: AlphaSuffix::Number,
            start: num,
            step: 1,
            width,
        };
    }

    // Check alphanumeric with trailing letter
    if let Some((prefix, letter_idx)) = extract_trailing_letter(text) {
        return FillPattern::AlphaNum {
            prefix,
            suffix: AlphaSuffix::Letter,
            start: letter_idx,
            step: 1,
            width: None,
        };
    }

    FillPattern::Copy
}

/// Detect pattern from two values
fn detect_two_value_pattern(source: &DetectedSource) -> FillPattern {
    let v1 = &source.values[0];
    let v2 = &source.values[1];

    // Two numbers → linear step (no growth detection in V1)
    if let (Value::Number(n1), Value::Number(n2)) = (v1, v2) {
        let step = n2 - n1;
        // Invariant: start = last value of source
        return FillPattern::Linear { start: *n2, step };
    }

    // Get text representations
    let t1 = match v1 {
        Value::Text(s) => s.as_str(),
        _ => return FillPattern::Repeat { values: source.values.clone() },
    };
    let t2 = match v2 {
        Value::Text(s) => s.as_str(),
        _ => return FillPattern::Repeat { values: source.values.clone() },
    };

    // Skip formulas - they go through existing ref-adjustment path
    if t1.starts_with('=') || t2.starts_with('=') {
        return FillPattern::Repeat { values: source.values.clone() };
    }

    // Two quarters (check BEFORE built-in list since quarters are in the list)
    if let (Some((q1, y1)), Some((q2, y2))) = (parse_quarter(t1), parse_quarter(t2)) {
        // Both must have year or both must not
        if y1.is_some() == y2.is_some() {
            let step = q2 - q1 + (y2.unwrap_or(0) - y1.unwrap_or(0)) * 4;
            return FillPattern::QuarterYear {
                quarter: q2,
                year: y2,
                step,
            };
        }
    }

    // Two list items in same list → calculate step (but not quarters - handled above)
    if let (Some((list1, idx1, case)), Some((list2, idx2, _))) =
        (find_in_builtin_list(t1), find_in_builtin_list(t2))
    {
        if list1 == list2 && list1 != BuiltinList::Quarters {
            let step = idx2 - idx1;
            // Invariant: start = last value (idx2)
            return FillPattern::TextList {
                list: list1,
                start_idx: idx2,
                step,
                case,
            };
        }
    }

    // Two alphanumeric with matching prefix
    if let (Some((p1, n1, w1)), Some((p2, n2, _))) =
        (extract_trailing_number(t1), extract_trailing_number(t2))
    {
        if p1 == p2 {
            let step = n2 - n1;
            // Invariant: start = last value (n2)
            return FillPattern::AlphaNum {
                prefix: p1,
                suffix: AlphaSuffix::Number,
                start: n2,
                step,
                width: w1,
            };
        }
    }

    // Two trailing letters with matching prefix
    if let (Some((p1, l1)), Some((p2, l2))) =
        (extract_trailing_letter(t1), extract_trailing_letter(t2))
    {
        if p1 == p2 {
            let step = l2 - l1;
            return FillPattern::AlphaNum {
                prefix: p1,
                suffix: AlphaSuffix::Letter,
                start: l2,
                step,
                width: None,
            };
        }
    }

    FillPattern::Repeat { values: source.values.clone() }
}

/// Detect pattern from 3+ values
fn detect_multi_value_pattern(source: &DetectedSource) -> FillPattern {
    if source.values.len() < 2 {
        return FillPattern::Repeat { values: source.values.clone() };
    }

    // Try to detect constant step from first two values
    let two_source = DetectedSource {
        values: vec![source.values[0].clone(), source.values[1].clone()],
        text_tokens: vec![
            source.text_tokens.get(0).cloned().flatten(),
            source.text_tokens.get(1).cloned().flatten(),
        ],
    };

    let pattern = detect_two_value_pattern(&two_source);

    // Validate remaining values have same step
    if validate_constant_step(&pattern, source) {
        // Update pattern to use last value as start
        update_pattern_start(&pattern, source)
    } else {
        FillPattern::Repeat { values: source.values.clone() }
    }
}

/// Validate that all values in source follow the detected pattern's step
fn validate_constant_step(pattern: &FillPattern, source: &DetectedSource) -> bool {
    match pattern {
        FillPattern::Linear { step, .. } => {
            for i in 1..source.values.len() {
                if let (Value::Number(n1), Value::Number(n2)) =
                    (&source.values[i - 1], &source.values[i])
                {
                    if (n2 - n1 - step).abs() > f64::EPSILON {
                        return false;
                    }
                } else {
                    return false;
                }
            }
            true
        }
        FillPattern::TextList { list, step, .. } => {
            for i in 1..source.values.len() {
                if let (Value::Text(t1), Value::Text(t2)) =
                    (&source.values[i - 1], &source.values[i])
                {
                    if let (Some((l1, idx1, _)), Some((l2, idx2, _))) =
                        (find_in_builtin_list(t1), find_in_builtin_list(t2))
                    {
                        if l1 != *list || l2 != *list || (idx2 - idx1) != *step {
                            return false;
                        }
                    } else {
                        return false;
                    }
                } else {
                    return false;
                }
            }
            true
        }
        FillPattern::AlphaNum { prefix, suffix, step, .. } => {
            for i in 1..source.values.len() {
                if let (Value::Text(t1), Value::Text(t2)) =
                    (&source.values[i - 1], &source.values[i])
                {
                    match suffix {
                        AlphaSuffix::Number => {
                            if let (Some((p1, n1, _)), Some((p2, n2, _))) =
                                (extract_trailing_number(t1), extract_trailing_number(t2))
                            {
                                if p1 != *prefix || p2 != *prefix || (n2 - n1) != *step {
                                    return false;
                                }
                            } else {
                                return false;
                            }
                        }
                        AlphaSuffix::Letter => {
                            if let (Some((p1, l1)), Some((p2, l2))) =
                                (extract_trailing_letter(t1), extract_trailing_letter(t2))
                            {
                                if p1 != *prefix || p2 != *prefix || (l2 - l1) != *step {
                                    return false;
                                }
                            } else {
                                return false;
                            }
                        }
                    }
                } else {
                    return false;
                }
            }
            true
        }
        _ => false,
    }
}

/// Update pattern to use last value of source as start
fn update_pattern_start(pattern: &FillPattern, source: &DetectedSource) -> FillPattern {
    let last = source.values.last().unwrap();

    match pattern {
        FillPattern::Linear { step, .. } => {
            if let Value::Number(n) = last {
                FillPattern::Linear { start: *n, step: *step }
            } else {
                pattern.clone()
            }
        }
        FillPattern::TextList { list, step, case, .. } => {
            if let Value::Text(t) = last {
                if let Some((_, idx, _)) = find_in_builtin_list(t) {
                    return FillPattern::TextList {
                        list: *list,
                        start_idx: idx,
                        step: *step,
                        case: *case,
                    };
                }
            }
            pattern.clone()
        }
        FillPattern::AlphaNum { prefix, suffix, step, width, .. } => {
            if let Value::Text(t) = last {
                match suffix {
                    AlphaSuffix::Number => {
                        if let Some((_, n, _)) = extract_trailing_number(t) {
                            return FillPattern::AlphaNum {
                                prefix: prefix.clone(),
                                suffix: *suffix,
                                start: n,
                                step: *step,
                                width: *width,
                            };
                        }
                    }
                    AlphaSuffix::Letter => {
                        if let Some((_, l)) = extract_trailing_letter(t) {
                            return FillPattern::AlphaNum {
                                prefix: prefix.clone(),
                                suffix: *suffix,
                                start: l,
                                step: *step,
                                width: *width,
                            };
                        }
                    }
                }
            }
            pattern.clone()
        }
        _ => pattern.clone(),
    }
}

// ============================================================================
// Value Generation
// ============================================================================

/// Generate the k-th value after the source (1-indexed)
pub fn generate(pattern: &FillPattern, k: usize) -> Value {
    let k = k as i64;

    match pattern {
        FillPattern::Copy => panic!("generate() called on Copy pattern"),

        FillPattern::Repeat { values } => {
            values[(k as usize - 1) % values.len()].clone()
        }

        FillPattern::Linear { start, step } => {
            Value::Number(start + step * k as f64)
        }

        FillPattern::TextList { list, start_idx, step, case } => {
            let list_values = get_builtin_list(*list);
            let list_len = list_values.len() as i32;
            let new_idx = (start_idx + step * k as i32).rem_euclid(list_len) as usize;
            let value = list_values[new_idx];
            Value::Text(apply_case(value, *case))
        }

        FillPattern::AlphaNum { prefix, suffix, start, step, width } => {
            let new_val = start + step * k;
            let suffix_str = match suffix {
                AlphaSuffix::Number => format_number_with_width(new_val, *width),
                AlphaSuffix::Letter => index_to_letters(new_val),
            };
            Value::Text(format!("{}{}", prefix, suffix_str))
        }

        FillPattern::QuarterYear { quarter, year, step } => {
            let total_quarters = *quarter + step * k as i32;
            let new_q = ((total_quarters - 1).rem_euclid(4)) + 1; // 1-4
            let year_offset = (total_quarters - 1).div_euclid(4);
            match year {
                Some(y) => Value::Text(format!("Q{} {}", new_q, y + year_offset)),
                None => Value::Text(format!("Q{}", new_q)),
            }
        }
    }
}

// ============================================================================
// Built-in Lists
// ============================================================================

const MONTHS_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

const MONTHS_LONG: [&str; 12] = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
];

const WEEKDAYS_SHORT: [&str; 7] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

const WEEKDAYS_LONG: [&str; 7] = [
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
];

const QUARTERS: [&str; 4] = ["Q1", "Q2", "Q3", "Q4"];

fn get_builtin_list(list: BuiltinList) -> &'static [&'static str] {
    match list {
        BuiltinList::MonthsShort => &MONTHS_SHORT,
        BuiltinList::MonthsLong => &MONTHS_LONG,
        BuiltinList::WeekdaysShort => &WEEKDAYS_SHORT,
        BuiltinList::WeekdaysLong => &WEEKDAYS_LONG,
        BuiltinList::Quarters => &QUARTERS,
    }
}

/// Find text in a built-in list, return (list, 0-indexed position, case mode)
fn find_in_builtin_list(text: &str) -> Option<(BuiltinList, i32, CaseMode)> {
    let normalized = text.trim();

    // Detect case mode from input
    let case = if normalized.chars().all(|c| c.is_uppercase() || !c.is_alphabetic()) {
        CaseMode::Upper
    } else if normalized.chars().all(|c| c.is_lowercase() || !c.is_alphabetic()) {
        CaseMode::Lower
    } else {
        CaseMode::Title
    };

    let lower = normalized.to_lowercase();

    // Check each list
    for (list, values) in [
        (BuiltinList::MonthsShort, MONTHS_SHORT.as_slice()),
        (BuiltinList::MonthsLong, MONTHS_LONG.as_slice()),
        (BuiltinList::WeekdaysShort, WEEKDAYS_SHORT.as_slice()),
        (BuiltinList::WeekdaysLong, WEEKDAYS_LONG.as_slice()),
        (BuiltinList::Quarters, QUARTERS.as_slice()),
    ] {
        for (i, &val) in values.iter().enumerate() {
            if val.to_lowercase() == lower {
                return Some((list, i as i32, case));
            }
        }
    }

    None
}

/// Parse quarter notation: "Q1", "Q2 2026", etc.
fn parse_quarter(text: &str) -> Option<(i32, Option<i32>)> {
    let text = text.trim();

    // Match Q followed by 1-4, optionally followed by year
    if !text.starts_with('Q') && !text.starts_with('q') {
        return None;
    }

    let rest = &text[1..];
    let parts: Vec<&str> = rest.split_whitespace().collect();

    match parts.len() {
        1 => {
            // Just Q1, Q2, etc.
            let q: i32 = parts[0].parse().ok()?;
            if (1..=4).contains(&q) {
                Some((q, None))
            } else {
                None
            }
        }
        2 => {
            // Q1 2026
            let q: i32 = parts[0].parse().ok()?;
            let year: i32 = parts[1].parse().ok()?;
            if (1..=4).contains(&q) {
                Some((q, Some(year)))
            } else {
                None
            }
        }
        _ => None,
    }
}

// ============================================================================
// Alphanumeric Helpers
// ============================================================================

/// Extract trailing number from text: "Item-1" → ("Item", -1, None)
/// Returns (prefix, number, width for zero-padded text numbers)
fn extract_trailing_number(text: &str) -> Option<(String, i64, Option<usize>)> {
    // Find where the trailing number starts (including optional minus)
    let chars: Vec<char> = text.chars().collect();
    if chars.is_empty() {
        return None;
    }

    // Find the last sequence of digits (possibly preceded by minus)
    let mut num_start = chars.len();
    let mut has_minus = false;

    // Walk backwards to find digits
    for i in (0..chars.len()).rev() {
        if chars[i].is_ascii_digit() {
            num_start = i;
        } else if chars[i] == '-' && i + 1 < chars.len() && chars[i + 1].is_ascii_digit() && num_start == i + 1 {
            // Minus immediately before digits
            num_start = i;
            has_minus = true;
            break;
        } else if num_start < chars.len() {
            break;
        }
    }

    if num_start >= chars.len() {
        return None;
    }

    let prefix: String = chars[..num_start].iter().collect();
    let num_str: String = chars[num_start..].iter().collect();

    // Parse the number
    let num: i64 = num_str.parse().ok()?;

    // Calculate width (digits only, excluding minus)
    let digit_count = if has_minus {
        num_str.len() - 1
    } else {
        num_str.len()
    };

    // Only use width if there are leading zeros
    let width = if digit_count > 1 && num_str.trim_start_matches('-').starts_with('0') {
        Some(digit_count)
    } else {
        None
    };

    Some((prefix, num, width))
}

/// Extract trailing letter from text: "Row A" → ("Row ", 1)
/// Returns (prefix, 1-indexed letter value: A=1, Z=26, AA=27)
fn extract_trailing_letter(text: &str) -> Option<(String, i64)> {
    let chars: Vec<char> = text.chars().collect();
    if chars.is_empty() {
        return None;
    }

    // Find trailing letter sequence
    let mut letter_start = chars.len();
    for i in (0..chars.len()).rev() {
        if chars[i].is_ascii_alphabetic() {
            letter_start = i;
        } else if letter_start < chars.len() {
            break;
        }
    }

    if letter_start >= chars.len() {
        return None;
    }

    // Must have at least one non-letter prefix character or start at position 0
    // Actually, we allow pure letters like "A" to not match as alphanum
    // We need a prefix to distinguish from built-in lists
    if letter_start == 0 {
        return None;
    }

    let prefix: String = chars[..letter_start].iter().collect();
    let letter_str: String = chars[letter_start..].iter().collect();

    // Convert letters to 1-indexed value (A=1, Z=26, AA=27)
    let letter_val = letters_to_index(&letter_str.to_uppercase());

    Some((prefix, letter_val))
}

/// Convert letter string to 1-indexed value: A=1, Z=26, AA=27
fn letters_to_index(letters: &str) -> i64 {
    letters.chars().fold(0i64, |acc, c| {
        acc * 26 + (c as i64 - 'A' as i64 + 1)
    })
}

/// Convert 1-indexed value to letter string: 1=A, 26=Z, 27=AA
fn index_to_letters(n: i64) -> String {
    if n <= 0 {
        return "A".to_string(); // Clamp to A for negative/zero
    }
    let mut result = String::new();
    let mut n = n;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}

/// Format number with optional zero-padding and sign preservation
fn format_number_with_width(n: i64, width: Option<usize>) -> String {
    match width {
        Some(w) => {
            if n < 0 {
                format!("-{:0>width$}", -n, width = w)
            } else {
                format!("{:0>width$}", n, width = w)
            }
        }
        None => n.to_string(),
    }
}

/// Apply case mode to a string
fn apply_case(text: &str, case: CaseMode) -> String {
    match case {
        CaseMode::Upper => text.to_uppercase(),
        CaseMode::Lower => text.to_lowercase(),
        CaseMode::Title => {
            // Title case: first letter uppercase, rest lowercase
            let mut chars = text.chars();
            match chars.next() {
                None => String::new(),
                Some(first) => first.to_uppercase().chain(chars.flat_map(|c| c.to_lowercase())).collect(),
            }
        }
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    /// Helper to create a numeric Value
    fn num(n: f64) -> Value {
        Value::Number(n)
    }

    /// Helper to create a text Value
    fn text(s: &str) -> Value {
        Value::Text(s.to_string())
    }

    /// Helper to create DetectedSource from values
    fn source(values: Vec<Value>) -> DetectedSource {
        let text_tokens = values.iter().map(|v| {
            if let Value::Text(s) = v {
                Some(s.clone())
            } else {
                None
            }
        }).collect();
        DetectedSource { values, text_tokens }
    }

    // ========================================================================
    // Test 1-5: Single Cell (Excel-default behavior)
    // ========================================================================

    #[test]
    fn test_01_single_number_drag_copies() {
        // 1: drag down 3 → 1, 1, 1 (Copy by default)
        let intent = fill_intent(false, 1, false);
        assert_eq!(intent, FillIntent::Copy);
    }

    #[test]
    fn test_02_single_number_ctrl_drag_series() {
        // 1: Ctrl+drag down 3 → 2, 3, 4
        let intent = fill_intent(true, 1, false);
        assert_eq!(intent, FillIntent::Series);

        let src = source(vec![num(1.0)]);
        let pattern = detect_pattern(&src, intent);
        assert_eq!(pattern, FillPattern::Linear { start: 1.0, step: 1.0 });

        assert_eq!(generate(&pattern, 1), num(2.0));
        assert_eq!(generate(&pattern, 2), num(3.0));
        assert_eq!(generate(&pattern, 3), num(4.0));
    }

    #[test]
    fn test_03_single_month_drag_series() {
        // Jan: drag down 3 → Feb, Mar, Apr
        let intent = fill_intent(false, 1, true);
        assert_eq!(intent, FillIntent::Series);

        let src = source(vec![text("Jan")]);
        let pattern = detect_pattern(&src, intent);

        if let FillPattern::TextList { list, start_idx, step, case } = &pattern {
            assert_eq!(*list, BuiltinList::MonthsShort);
            assert_eq!(*start_idx, 0); // Jan is index 0
            assert_eq!(*step, 1);
            assert_eq!(*case, CaseMode::Title);
        } else {
            panic!("Expected TextList pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Feb"));
        assert_eq!(generate(&pattern, 2), text("Mar"));
        assert_eq!(generate(&pattern, 3), text("Apr"));
    }

    #[test]
    fn test_04_single_month_ctrl_drag_copies() {
        // Jan: Ctrl+drag → Jan, Jan, Jan
        let intent = fill_intent(true, 1, true);
        assert_eq!(intent, FillIntent::Copy);
    }

    #[test]
    fn test_05_single_text_no_pattern_copies() {
        // Hello: drag → Hello, Hello, Hello
        let src = source(vec![text("Hello")]);
        let pattern = detect_pattern(&src, FillIntent::Series);
        assert_eq!(pattern, FillPattern::Copy);
    }

    // ========================================================================
    // Test 6-10: Two Cell (Pattern Detection)
    // ========================================================================

    #[test]
    fn test_06_two_numbers_linear() {
        // 1, 2: drag down 3 → 3, 4, 5
        let src = source(vec![num(1.0), num(2.0)]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        // start = last value (2), step = 1
        assert_eq!(pattern, FillPattern::Linear { start: 2.0, step: 1.0 });

        assert_eq!(generate(&pattern, 1), num(3.0));
        assert_eq!(generate(&pattern, 2), num(4.0));
        assert_eq!(generate(&pattern, 3), num(5.0));
    }

    #[test]
    fn test_07_two_numbers_negative_step() {
        // 5, 3: drag down 3 → 1, -1, -3
        let src = source(vec![num(5.0), num(3.0)]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        // start = 3, step = -2
        assert_eq!(pattern, FillPattern::Linear { start: 3.0, step: -2.0 });

        assert_eq!(generate(&pattern, 1), num(1.0));
        assert_eq!(generate(&pattern, 2), num(-1.0));
        assert_eq!(generate(&pattern, 3), num(-3.0));
    }

    #[test]
    fn test_08_two_numbers_step_10() {
        // 10, 20: drag down 3 → 30, 40, 50
        let src = source(vec![num(10.0), num(20.0)]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        assert_eq!(pattern, FillPattern::Linear { start: 20.0, step: 10.0 });

        assert_eq!(generate(&pattern, 1), num(30.0));
        assert_eq!(generate(&pattern, 2), num(40.0));
        assert_eq!(generate(&pattern, 3), num(50.0));
    }

    #[test]
    fn test_09_two_months_step_2() {
        // Jan, Mar: drag down 3 → May, Jul, Sep
        let src = source(vec![text("Jan"), text("Mar")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::TextList { start_idx, step, .. } = &pattern {
            assert_eq!(*start_idx, 2); // Mar is index 2
            assert_eq!(*step, 2);
        } else {
            panic!("Expected TextList pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("May"));
        assert_eq!(generate(&pattern, 2), text("Jul"));
        assert_eq!(generate(&pattern, 3), text("Sep"));
    }

    #[test]
    fn test_10_two_quarters_wrap() {
        // Q3, Q4: drag right 2 → Q1, Q2
        let src = source(vec![text("Q3"), text("Q4")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::QuarterYear { quarter, year, step } = &pattern {
            assert_eq!(*quarter, 4);
            assert_eq!(*year, None);
            assert_eq!(*step, 1);
        } else {
            panic!("Expected QuarterYear pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Q1"));
        assert_eq!(generate(&pattern, 2), text("Q2"));
    }

    // ========================================================================
    // Test 11-12: Multi Cell (Validate or Repeat)
    // ========================================================================

    #[test]
    fn test_11_multi_constant_step() {
        // 1, 2, 3: drag down 3 → 4, 5, 6
        let src = source(vec![num(1.0), num(2.0), num(3.0)]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        // start = 3 (last), step = 1
        assert_eq!(pattern, FillPattern::Linear { start: 3.0, step: 1.0 });

        assert_eq!(generate(&pattern, 1), num(4.0));
        assert_eq!(generate(&pattern, 2), num(5.0));
        assert_eq!(generate(&pattern, 3), num(6.0));
    }

    #[test]
    fn test_12_multi_non_constant_repeats() {
        // 1, 2, 4: drag → 1, 2, 4 (repeat)
        let src = source(vec![num(1.0), num(2.0), num(4.0)]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::Repeat { values } = &pattern {
            assert_eq!(values.len(), 3);
        } else {
            panic!("Expected Repeat pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), num(1.0));
        assert_eq!(generate(&pattern, 2), num(2.0));
        assert_eq!(generate(&pattern, 3), num(4.0));
        assert_eq!(generate(&pattern, 4), num(1.0)); // Wraps
    }

    // ========================================================================
    // Test 13-15: Alphanumeric
    // ========================================================================

    #[test]
    fn test_13_trailing_number() {
        // Item1: drag down 3 → Item2, Item3, Item4
        let src = source(vec![text("Item1")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::AlphaNum { prefix, suffix, start, step, .. } = &pattern {
            assert_eq!(prefix, "Item");
            assert_eq!(*suffix, AlphaSuffix::Number);
            assert_eq!(*start, 1);
            assert_eq!(*step, 1);
        } else {
            panic!("Expected AlphaNum pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Item2"));
        assert_eq!(generate(&pattern, 2), text("Item3"));
        assert_eq!(generate(&pattern, 3), text("Item4"));
    }

    #[test]
    fn test_14_trailing_letter() {
        // Row A: drag down 3 → Row B, Row C, Row D
        let src = source(vec![text("Row A")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::AlphaNum { prefix, suffix, start, step, .. } = &pattern {
            assert_eq!(prefix, "Row ");
            assert_eq!(*suffix, AlphaSuffix::Letter);
            assert_eq!(*start, 1); // A = 1
            assert_eq!(*step, 1);
        } else {
            panic!("Expected AlphaNum pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Row B"));
        assert_eq!(generate(&pattern, 2), text("Row C"));
        assert_eq!(generate(&pattern, 3), text("Row D"));
    }

    #[test]
    fn test_15_letter_overflow() {
        // Row Z: drag down 2 → Row AA, Row AB
        let src = source(vec![text("Row Z")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::AlphaNum { start, .. } = &pattern {
            assert_eq!(*start, 26); // Z = 26
        } else {
            panic!("Expected AlphaNum pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Row AA"));
        assert_eq!(generate(&pattern, 2), text("Row AB"));
    }

    // ========================================================================
    // Test 16: Negative Trailing Numbers
    // ========================================================================

    #[test]
    fn test_16_negative_trailing_number() {
        // Item-1, Item0: drag down 2 → Item1, Item2
        let src = source(vec![text("Item-1"), text("Item0")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::AlphaNum { prefix, start, step, .. } = &pattern {
            assert_eq!(prefix, "Item");
            assert_eq!(*start, 0); // Last value
            assert_eq!(*step, 1);
        } else {
            panic!("Expected AlphaNum pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Item1"));
        assert_eq!(generate(&pattern, 2), text("Item2"));
    }

    // ========================================================================
    // Test 17: Leading Zeros (Text)
    // ========================================================================

    #[test]
    fn test_17_leading_zeros() {
        // "001", "002": drag down 2 → "003", "004"
        let src = source(vec![text("001"), text("002")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::AlphaNum { prefix, start, step, width, .. } = &pattern {
            assert_eq!(prefix, "");
            assert_eq!(*start, 2);
            assert_eq!(*step, 1);
            assert_eq!(*width, Some(3));
        } else {
            panic!("Expected AlphaNum pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("003"));
        assert_eq!(generate(&pattern, 2), text("004"));
    }

    // ========================================================================
    // Test 18-19: Quarter/Year
    // ========================================================================

    #[test]
    fn test_18_quarter_wrap() {
        // Q1: drag right 4 → Q2, Q3, Q4, Q1
        let src = source(vec![text("Q1")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::QuarterYear { quarter, year, step } = &pattern {
            assert_eq!(*quarter, 1);
            assert_eq!(*year, None);
            assert_eq!(*step, 1);
        } else {
            panic!("Expected QuarterYear pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Q2"));
        assert_eq!(generate(&pattern, 2), text("Q3"));
        assert_eq!(generate(&pattern, 3), text("Q4"));
        assert_eq!(generate(&pattern, 4), text("Q1")); // Wrap
    }

    #[test]
    fn test_19_quarter_year_wrap() {
        // Q4 2026: drag right 2 → Q1 2027, Q2 2027
        let src = source(vec![text("Q4 2026")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::QuarterYear { quarter, year, step } = &pattern {
            assert_eq!(*quarter, 4);
            assert_eq!(*year, Some(2026));
            assert_eq!(*step, 1);
        } else {
            panic!("Expected QuarterYear pattern, got {:?}", pattern);
        }

        assert_eq!(generate(&pattern, 1), text("Q1 2027"));
        assert_eq!(generate(&pattern, 2), text("Q2 2027"));
    }

    // ========================================================================
    // Test 20-22: Edge Cases
    // ========================================================================

    #[test]
    fn test_20_formula_not_detected() {
        // Formulas should go through existing ref-adjustment path, not series fill
        // This test just verifies formula text isn't detected as a pattern
        let src = source(vec![text("=A1")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        // Should be Copy (no alphanumeric pattern matches)
        assert_eq!(pattern, FillPattern::Copy);
    }

    #[test]
    fn test_21_blank_in_source_repeats() {
        // 1, (empty): drag → repeat pattern
        let src = source(vec![num(1.0), Value::Empty]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::Repeat { values } = &pattern {
            assert_eq!(values.len(), 2);
        } else {
            panic!("Expected Repeat pattern, got {:?}", pattern);
        }
    }

    #[test]
    fn test_22_mixed_types_repeats() {
        // 1, "text": drag → repeat pattern
        let src = source(vec![num(1.0), text("text")]);
        let pattern = detect_pattern(&src, FillIntent::Series);

        if let FillPattern::Repeat { values } = &pattern {
            assert_eq!(values.len(), 2);
        } else {
            panic!("Expected Repeat pattern, got {:?}", pattern);
        }
    }

    // ========================================================================
    // Helper function tests
    // ========================================================================

    #[test]
    fn test_letters_to_index() {
        assert_eq!(letters_to_index("A"), 1);
        assert_eq!(letters_to_index("Z"), 26);
        assert_eq!(letters_to_index("AA"), 27);
        assert_eq!(letters_to_index("AB"), 28);
        assert_eq!(letters_to_index("AZ"), 52);
        assert_eq!(letters_to_index("BA"), 53);
    }

    #[test]
    fn test_index_to_letters() {
        assert_eq!(index_to_letters(1), "A");
        assert_eq!(index_to_letters(26), "Z");
        assert_eq!(index_to_letters(27), "AA");
        assert_eq!(index_to_letters(28), "AB");
        assert_eq!(index_to_letters(52), "AZ");
        assert_eq!(index_to_letters(53), "BA");
        assert_eq!(index_to_letters(0), "A"); // Clamp
        assert_eq!(index_to_letters(-5), "A"); // Clamp
    }

    #[test]
    fn test_format_number_with_width() {
        assert_eq!(format_number_with_width(3, Some(3)), "003");
        assert_eq!(format_number_with_width(-3, Some(3)), "-003");
        assert_eq!(format_number_with_width(42, None), "42");
        assert_eq!(format_number_with_width(-42, None), "-42");
    }

    #[test]
    fn test_extract_trailing_number() {
        assert_eq!(extract_trailing_number("Item1"), Some(("Item".to_string(), 1, None)));
        assert_eq!(extract_trailing_number("Item-1"), Some(("Item".to_string(), -1, None)));
        assert_eq!(extract_trailing_number("Row 42"), Some(("Row ".to_string(), 42, None)));
        assert_eq!(extract_trailing_number("001"), Some(("".to_string(), 1, Some(3))));
        assert_eq!(extract_trailing_number("-001"), Some(("".to_string(), -1, Some(3))));
        assert_eq!(extract_trailing_number("NoNumber"), None);
    }

    #[test]
    fn test_case_preservation() {
        // JAN → FEB (upper)
        let src = source(vec![text("JAN")]);
        let pattern = detect_pattern(&src, FillIntent::Series);
        assert_eq!(generate(&pattern, 1), text("FEB"));

        // jan → feb (lower)
        let src = source(vec![text("jan")]);
        let pattern = detect_pattern(&src, FillIntent::Series);
        assert_eq!(generate(&pattern, 1), text("feb"));

        // Jan → Feb (title)
        let src = source(vec![text("Jan")]);
        let pattern = detect_pattern(&src, FillIntent::Series);
        assert_eq!(generate(&pattern, 1), text("Feb"));
    }
}
