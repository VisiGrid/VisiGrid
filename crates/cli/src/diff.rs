// Reconciliation logic for `visigrid diff`
// Pure functions: two datasets in, matched/unmatched/diff rows out.
// No IO, no clap, no formatting.

use std::collections::HashMap;

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct DiffOptions {
    pub key_col: usize,
    pub compare_cols: Option<Vec<usize>>,
    pub match_mode: MatchMode,
    pub key_transform: KeyTransform,
    pub on_ambiguous: AmbiguityPolicy,
    pub tolerance: f64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MatchMode {
    Exact,
    Contains,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KeyTransform {
    None,
    Trim,
    Digits,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AmbiguityPolicy {
    Error,
    Report,
}

// ---------------------------------------------------------------------------
// Input
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct DataRow {
    pub key_raw: String,
    pub key_norm: String,
    pub values: HashMap<String, String>,
}

// ---------------------------------------------------------------------------
// Results
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct DiffResult {
    pub results: Vec<DiffRow>,
    pub summary: DiffSummary,
    pub ambiguous_keys: Vec<AmbiguousKey>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowStatus {
    Matched,
    OnlyLeft,
    OnlyRight,
    Diff,
    Ambiguous,
}

impl RowStatus {
    pub fn as_str(&self) -> &'static str {
        match self {
            RowStatus::Matched => "matched",
            RowStatus::OnlyLeft => "only_left",
            RowStatus::OnlyRight => "only_right",
            RowStatus::Diff => "diff",
            RowStatus::Ambiguous => "ambiguous",
        }
    }
}

#[derive(Debug, Clone)]
pub struct DiffRow {
    pub status: RowStatus,
    pub key: String,
    pub left: Option<HashMap<String, String>>,
    pub right: Option<HashMap<String, String>>,
    pub diffs: Vec<ColumnDiff>,
    pub match_explain: Option<MatchExplain>,
    pub candidates: Option<Vec<Candidate>>,
}

#[derive(Debug, Clone)]
pub struct ColumnDiff {
    pub column: String,
    pub left: String,
    pub right: String,
    pub delta: Option<f64>,
    pub within_tolerance: bool,
}

#[derive(Debug, Clone)]
pub struct MatchExplain {
    pub mode: String,
    pub left_key_raw: String,
    pub right_key_raw: String,
    pub left_key_norm: String,
    pub right_key_norm: String,
}

#[derive(Debug, Clone)]
pub struct Candidate {
    pub right_key_raw: String,
    pub right_row_index: usize,
}

#[derive(Debug, Clone)]
pub struct AmbiguousKey {
    pub key: String,
    pub candidates: Vec<Candidate>,
}

#[derive(Debug, Clone)]
pub struct DuplicateKey {
    pub side: Side,
    pub key: String,
    pub count: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Side {
    Left,
    Right,
}

impl Side {
    pub fn as_str(&self) -> &'static str {
        match self {
            Side::Left => "left",
            Side::Right => "right",
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct DiffSummary {
    pub left_rows: usize,
    pub right_rows: usize,
    pub matched: usize,
    pub only_left: usize,
    pub only_right: usize,
    pub diff: usize,
    pub diff_outside_tolerance: usize,
    pub ambiguous: usize,
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub enum DiffError {
    DuplicateKeys(Vec<DuplicateKey>),
}

impl std::fmt::Display for DiffError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DiffError::DuplicateKeys(dups) => {
                writeln!(f, "duplicate keys found:")?;
                for dup in dups {
                    writeln!(f, "  {} key {:?} appears {} times", dup.side.as_str(), dup.key, dup.count)?;
                }
                Ok(())
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Key transform
// ---------------------------------------------------------------------------

pub fn apply_key_transform(raw: &str, transform: KeyTransform) -> String {
    match transform {
        KeyTransform::None => raw.to_string(),
        KeyTransform::Trim => raw.trim().to_string(),
        KeyTransform::Digits => raw.chars().filter(|c| c.is_ascii_digit()).collect(),
    }
}

// ---------------------------------------------------------------------------
// Numeric parsing for financial values
// ---------------------------------------------------------------------------

/// Parse a financial number string:
/// - Strip `$`, commas, whitespace
/// - Handle `(123.45)` → `-123.45`
/// - Returns None if non-numeric characters remain after stripping
pub fn parse_financial_number(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Check for parenthesized negatives: (123.45) → -123.45
    let (is_negative, inner) = if trimmed.starts_with('(') && trimmed.ends_with(')') {
        (true, &trimmed[1..trimmed.len() - 1])
    } else {
        (false, trimmed)
    };

    // Strip allowed non-numeric characters: $, commas, whitespace
    let cleaned: String = inner
        .chars()
        .filter(|c| *c != '$' && *c != ',' && !c.is_whitespace())
        .collect();

    if cleaned.is_empty() {
        return None;
    }

    // After stripping, only digits, '.', '-', '+' should remain
    // Allow leading minus (but not if already negative from parens)
    for (i, c) in cleaned.chars().enumerate() {
        match c {
            '0'..='9' | '.' => {}
            '-' | '+' if i == 0 && !is_negative => {}
            _ => return None, // Non-numeric character → treat as string
        }
    }

    let value: f64 = cleaned.parse().ok()?;
    Some(if is_negative { -value } else { value })
}

// ---------------------------------------------------------------------------
// Core reconciliation
// ---------------------------------------------------------------------------

pub fn reconcile(
    left_rows: &[DataRow],
    right_rows: &[DataRow],
    headers: &[String],
    options: &DiffOptions,
) -> Result<DiffResult, DiffError> {
    // 1. Check for duplicate keys in each side
    let mut duplicates = Vec::new();
    check_duplicates(left_rows, Side::Left, &mut duplicates);
    check_duplicates(right_rows, Side::Right, &mut duplicates);
    if !duplicates.is_empty() {
        return Err(DiffError::DuplicateKeys(duplicates));
    }

    // 2. Determine which columns to compare
    let compare_cols: Vec<String> = match &options.compare_cols {
        Some(indices) => indices
            .iter()
            .filter_map(|&i| headers.get(i).cloned())
            .collect(),
        None => headers
            .iter()
            .enumerate()
            .filter(|(i, _)| *i != options.key_col)
            .map(|(_, h)| h.clone())
            .collect(),
    };

    // 3. Build index on right side
    let mut right_index: HashMap<String, usize> = HashMap::new();
    for (i, row) in right_rows.iter().enumerate() {
        right_index.insert(row.key_norm.clone(), i);
    }

    let mut right_consumed = vec![false; right_rows.len()];
    let mut results = Vec::new();
    let mut ambiguous_keys = Vec::new();

    // 4. Match left rows against right
    for left_row in left_rows {
        match options.match_mode {
            MatchMode::Exact => {
                if let Some(&right_idx) = right_index.get(&left_row.key_norm) {
                    right_consumed[right_idx] = true;
                    let right_row = &right_rows[right_idx];
                    let diffs = compare_values(left_row, right_row, &compare_cols, options.tolerance);
                    let status = if diffs.is_empty() {
                        RowStatus::Matched
                    } else {
                        RowStatus::Diff
                    };
                    results.push(DiffRow {
                        status,
                        key: left_row.key_norm.clone(),
                        left: Some(left_row.values.clone()),
                        right: Some(right_row.values.clone()),
                        diffs,
                        match_explain: None,
                        candidates: None,
                    });
                } else {
                    results.push(DiffRow {
                        status: RowStatus::OnlyLeft,
                        key: left_row.key_norm.clone(),
                        left: Some(left_row.values.clone()),
                        right: None,
                        diffs: Vec::new(),
                        match_explain: None,
                        candidates: None,
                    });
                }
            }
            MatchMode::Contains => {
                // Left key must be substring of right key (directional)
                let mut matches: Vec<(usize, &DataRow)> = Vec::new();
                for (i, right_row) in right_rows.iter().enumerate() {
                    if !right_consumed[i] && right_row.key_norm.contains(&left_row.key_norm) {
                        matches.push((i, right_row));
                    }
                }

                match matches.len() {
                    0 => {
                        results.push(DiffRow {
                            status: RowStatus::OnlyLeft,
                            key: left_row.key_norm.clone(),
                            left: Some(left_row.values.clone()),
                            right: None,
                            diffs: Vec::new(),
                            match_explain: Some(make_explain("contains", left_row, left_row)),
                            candidates: None,
                        });
                    }
                    1 => {
                        let (right_idx, right_row) = matches[0];
                        right_consumed[right_idx] = true;
                        let diffs = compare_values(left_row, right_row, &compare_cols, options.tolerance);
                        let status = if diffs.is_empty() {
                            RowStatus::Matched
                        } else {
                            RowStatus::Diff
                        };
                        results.push(DiffRow {
                            status,
                            key: left_row.key_norm.clone(),
                            left: Some(left_row.values.clone()),
                            right: Some(right_row.values.clone()),
                            diffs,
                            match_explain: Some(make_explain_pair("contains", left_row, right_row)),
                            candidates: None,
                        });
                    }
                    _ => {
                        // Ambiguous
                        let candidates: Vec<Candidate> = matches
                            .iter()
                            .map(|(idx, r)| Candidate {
                                right_key_raw: r.key_raw.clone(),
                                right_row_index: *idx,
                            })
                            .collect();

                        ambiguous_keys.push(AmbiguousKey {
                            key: left_row.key_norm.clone(),
                            candidates: candidates.clone(),
                        });

                        if options.on_ambiguous == AmbiguityPolicy::Report {
                            results.push(DiffRow {
                                status: RowStatus::Ambiguous,
                                key: left_row.key_norm.clone(),
                                left: Some(left_row.values.clone()),
                                right: None,
                                diffs: Vec::new(),
                                match_explain: None,
                                candidates: Some(candidates),
                            });
                        }
                    }
                }
            }
        }
    }

    // 5. Any right rows not consumed → only_right
    for (i, right_row) in right_rows.iter().enumerate() {
        if !right_consumed[i] {
            results.push(DiffRow {
                status: RowStatus::OnlyRight,
                key: right_row.key_norm.clone(),
                left: None,
                right: Some(right_row.values.clone()),
                diffs: Vec::new(),
                match_explain: None,
                candidates: None,
            });
        }
    }

    // 6. Build summary
    let diff_rows: Vec<&DiffRow> = results.iter().filter(|r| r.status == RowStatus::Diff).collect();
    let diff_outside_tolerance = diff_rows.iter()
        .filter(|r| r.diffs.iter().any(|d| !d.within_tolerance))
        .count();
    let summary = DiffSummary {
        left_rows: left_rows.len(),
        right_rows: right_rows.len(),
        matched: results.iter().filter(|r| r.status == RowStatus::Matched).count(),
        only_left: results.iter().filter(|r| r.status == RowStatus::OnlyLeft).count(),
        only_right: results.iter().filter(|r| r.status == RowStatus::OnlyRight).count(),
        diff: diff_rows.len(),
        diff_outside_tolerance,
        ambiguous: ambiguous_keys.len(),
    };

    Ok(DiffResult {
        results,
        summary,
        ambiguous_keys,
    })
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn check_duplicates(rows: &[DataRow], side: Side, out: &mut Vec<DuplicateKey>) {
    let mut counts: HashMap<&str, usize> = HashMap::new();
    for row in rows {
        *counts.entry(&row.key_norm).or_insert(0) += 1;
    }
    for (key, count) in counts {
        if count > 1 {
            out.push(DuplicateKey {
                side,
                key: key.to_string(),
                count,
            });
        }
    }
}

fn compare_values(
    left: &DataRow,
    right: &DataRow,
    compare_cols: &[String],
    tolerance: f64,
) -> Vec<ColumnDiff> {
    let mut diffs = Vec::new();

    for col_name in compare_cols {
        let left_val = left.values.get(col_name).map(|s| s.as_str()).unwrap_or("");
        let right_val = right.values.get(col_name).map(|s| s.as_str()).unwrap_or("");

        // Both empty = match
        if left_val.is_empty() && right_val.is_empty() {
            continue;
        }

        // Try numeric comparison
        let left_num = parse_financial_number(left_val);
        let right_num = parse_financial_number(right_val);

        match (left_num, right_num) {
            (Some(l), Some(r)) => {
                let delta = (l - r).abs();
                // Epsilon-inclusive comparison: preserve human-decimal boundary
                // semantics under IEEE-754 float representation.
                let scale = 1.0_f64
                    .max(l.abs())
                    .max(r.abs())
                    .max(delta)
                    .max(tolerance);
                let eps = f64::EPSILON * 16.0 * scale;
                let within = delta <= tolerance + eps;
                if !within || delta > 0.0 {
                    // Report diff if values aren't identical (even if within tolerance)
                    // but mark within_tolerance accordingly
                    if left_val != right_val {
                        diffs.push(ColumnDiff {
                            column: col_name.clone(),
                            left: left_val.to_string(),
                            right: right_val.to_string(),
                            delta: Some(delta),
                            within_tolerance: within,
                        });
                    }
                }
            }
            _ => {
                // String comparison
                if left_val != right_val {
                    diffs.push(ColumnDiff {
                        column: col_name.clone(),
                        left: left_val.to_string(),
                        right: right_val.to_string(),
                        delta: None,
                        within_tolerance: false,
                    });
                }
            }
        }
    }

    diffs
}

fn make_explain(mode: &str, left: &DataRow, _right: &DataRow) -> MatchExplain {
    MatchExplain {
        mode: mode.to_string(),
        left_key_raw: left.key_raw.clone(),
        right_key_raw: String::new(),
        left_key_norm: left.key_norm.clone(),
        right_key_norm: String::new(),
    }
}

fn make_explain_pair(mode: &str, left: &DataRow, right: &DataRow) -> MatchExplain {
    MatchExplain {
        mode: mode.to_string(),
        left_key_raw: left.key_raw.clone(),
        right_key_raw: right.key_raw.clone(),
        left_key_norm: left.key_norm.clone(),
        right_key_norm: right.key_norm.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_financial_number_basic() {
        assert_eq!(parse_financial_number("123.45"), Some(123.45));
        assert_eq!(parse_financial_number("-50"), Some(-50.0));
        assert_eq!(parse_financial_number("0"), Some(0.0));
    }

    #[test]
    fn test_parse_financial_number_currency() {
        assert_eq!(parse_financial_number("$685.00"), Some(685.0));
        assert_eq!(parse_financial_number("$1,234.56"), Some(1234.56));
    }

    #[test]
    fn test_parse_financial_number_parens() {
        assert_eq!(parse_financial_number("(500.00)"), Some(-500.0));
        assert_eq!(parse_financial_number("(1,234.56)"), Some(-1234.56));
    }

    #[test]
    fn test_parse_financial_number_whitespace() {
        assert_eq!(parse_financial_number("  123.45  "), Some(123.45));
        assert_eq!(parse_financial_number(""), None);
        assert_eq!(parse_financial_number("   "), None);
    }

    #[test]
    fn test_parse_financial_number_non_numeric() {
        assert_eq!(parse_financial_number("abc"), None);
        assert_eq!(parse_financial_number("12abc34"), None);
        assert_eq!(parse_financial_number("N/A"), None);
    }

    #[test]
    fn test_key_transform_none() {
        assert_eq!(apply_key_transform("  INV-123  ", KeyTransform::None), "  INV-123  ");
    }

    #[test]
    fn test_key_transform_trim() {
        assert_eq!(apply_key_transform("  INV-123  ", KeyTransform::Trim), "INV-123");
    }

    #[test]
    fn test_key_transform_digits() {
        assert_eq!(apply_key_transform("INV-123-AB", KeyTransform::Digits), "123");
        assert_eq!(apply_key_transform("100154662", KeyTransform::Digits), "100154662");
    }
}
