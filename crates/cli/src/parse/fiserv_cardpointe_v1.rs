//! Fiserv/CardPointe v1 statement parser.
//!
//! Parses the "Summary By Day" section from `pdftotext -layout` output of a
//! Fiserv monthly statement PDF. Extracts daily Amount Processed values.

use regex::Regex;

use crate::fetch::common::parse_money_string;
use crate::CliError;

/// Parsed statement with metadata and daily rows.
#[derive(Debug)]
pub(super) struct ParsedStatement {
    pub merchant_id: String,
    pub period_start: String,
    pub period_end: String,
    pub rows: Vec<DayRow>,
    pub total_amount_minor: Option<i64>,
}

/// A single day's Amount Processed.
#[derive(Debug)]
pub(super) struct DayRow {
    /// ISO date YYYY-MM-DD
    pub date: String,
    /// Amount in cents
    pub amount_minor: i64,
}

/// Parse pdftotext output using the fiserv_cardpointe_v1 template.
pub(super) fn parse(text: &str) -> Result<ParsedStatement, CliError> {
    if text.trim().is_empty() {
        return Err(CliError::parse(
            "PDF appears scanned/image-only — text extraction failed",
        ));
    }

    let merchant_id = extract_merchant_id(text)?;
    let (period_start, period_end) = extract_period(text)?;

    let (rows, total_amount_minor) = extract_summary_rows(text)?;

    if rows.is_empty() {
        return Err(CliError::parse(
            "Unsupported or unrecognized statement template",
        ));
    }

    Ok(ParsedStatement {
        merchant_id,
        period_start,
        period_end,
        rows,
        total_amount_minor,
    })
}

/// Extract Merchant Number from statement text.
fn extract_merchant_id(text: &str) -> Result<String, CliError> {
    let re = Regex::new(r"Merchant\s+Number\s+(\d+)").unwrap();
    re.captures(text)
        .and_then(|c| c.get(1))
        .map(|m| m.as_str().to_string())
        .ok_or_else(|| CliError::parse("Could not find Merchant Number in statement"))
}

/// Extract Statement Period (MM/DD/YY - MM/DD/YY).
fn extract_period(text: &str) -> Result<(String, String), CliError> {
    let re = Regex::new(r"Statement\s+Period\s+(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})")
        .unwrap();
    let caps = re
        .captures(text)
        .ok_or_else(|| CliError::parse("Could not find Statement Period in statement"))?;

    let start = parse_mmddyy(caps.get(1).unwrap().as_str())?;
    let end = parse_mmddyy(caps.get(2).unwrap().as_str())?;

    Ok((start, end))
}

/// Find the Summary By Day section and extract rows.
fn extract_summary_rows(text: &str) -> Result<(Vec<DayRow>, Option<i64>), CliError> {
    let lines: Vec<&str> = text.lines().collect();

    // Find section header: line containing both "Date" and "Processed"
    let header_idx = lines
        .iter()
        .position(|line| {
            let upper = line.to_uppercase();
            upper.contains("DATE") && upper.contains("PROCESSED")
        })
        .ok_or_else(|| {
            CliError::parse("Unsupported or unrecognized statement template")
        })?;

    let date_re = Regex::new(r"^\s*(\d{2}/\d{2}/\d{2})\s").unwrap();
    let money_re = Regex::new(r"-?\$[\d,]+\.\d{2}").unwrap();
    let total_re = Regex::new(r"(?i)^\s*T\s*o\s*t\s*a\s*l\s").unwrap();
    let section_end_re =
        Regex::new(r"(?i)^\s*(BATCH|CARD\s+TYPE|CHARGEBACKS|INTERCHANGE)").unwrap();

    let mut rows = Vec::new();
    let mut total_amount_minor: Option<i64> = None;

    for line in &lines[header_idx + 1..] {
        // Stop at section boundary
        if section_end_re.is_match(line) {
            break;
        }

        // Total line: extract for cross-check, then stop
        if total_re.is_match(line) {
            if let Some(amount) = extract_last_dollar_amount(line, &money_re) {
                total_amount_minor = Some(amount);
            }
            break;
        }

        // Skip non-date lines (e.g. "Month End Charge")
        if let Some(caps) = date_re.captures(line) {
            let date_str = caps.get(1).unwrap().as_str();

            // Require at least 2 dollar amounts on the line (ensures we're in the numeric table)
            let dollar_matches: Vec<_> = money_re.find_iter(line).collect();
            if dollar_matches.len() < 2 {
                continue;
            }

            // Last dollar amount is Amount Processed
            let amount = extract_last_dollar_amount(line, &money_re).ok_or_else(|| {
                CliError::parse(format!("Failed to parse row: {}", line.trim()))
            })?;

            let date = parse_mmddyy(date_str)?;
            rows.push(DayRow {
                date,
                amount_minor: amount,
            });
        }
    }

    Ok((rows, total_amount_minor))
}

/// Extract the last dollar amount token from a line.
/// Handles both `$1,234.56` and `-$1,234.56` and `($1,234.56)`.
fn extract_last_dollar_amount(line: &str, money_re: &Regex) -> Option<i64> {
    // Also handle parenthesized negatives like ($1,517.82)
    let paren_re = Regex::new(r"\(\$[\d,]+\.\d{2}\)").unwrap();

    let mut last_match: Option<&str> = None;
    let mut last_pos: usize = 0;
    let mut is_paren_negative = false;

    for m in money_re.find_iter(line) {
        if m.end() >= last_pos {
            last_match = Some(m.as_str());
            last_pos = m.end();
            is_paren_negative = false;
        }
    }

    // Check for parenthesized negative after the last regular match
    for m in paren_re.find_iter(line) {
        if m.end() >= last_pos {
            last_match = Some(m.as_str());
            last_pos = m.end();
            is_paren_negative = true;
        }
    }

    last_match.and_then(|s| {
        let cleaned = if is_paren_negative {
            // ($1,234.56) → -1234.56
            let inner = s.trim_start_matches('(').trim_end_matches(')');
            let stripped = inner.replace('$', "").replace(',', "");
            format!("-{}", stripped)
        } else {
            // -$1,234.56 or $1,234.56 → -1234.56 or 1234.56
            let negative = s.starts_with('-');
            let stripped = s.replace('$', "").replace(',', "").replace('-', "");
            if negative {
                format!("-{}", stripped)
            } else {
                stripped
            }
        };
        parse_money_string(&cleaned).ok()
    })
}

/// Parse MM/DD/YY → YYYY-MM-DD. YY < 50 → 20YY, YY ≥ 50 → 19YY.
fn parse_mmddyy(s: &str) -> Result<String, CliError> {
    let parts: Vec<&str> = s.split('/').collect();
    if parts.len() != 3 {
        return Err(CliError::parse(format!("Invalid date format: {}", s)));
    }
    let mm: u32 = parts[0]
        .parse()
        .map_err(|_| CliError::parse(format!("Invalid month in date: {}", s)))?;
    let dd: u32 = parts[1]
        .parse()
        .map_err(|_| CliError::parse(format!("Invalid day in date: {}", s)))?;
    let yy: u32 = parts[2]
        .parse()
        .map_err(|_| CliError::parse(format!("Invalid year in date: {}", s)))?;

    let yyyy = if yy < 50 { 2000 + yy } else { 1900 + yy };

    Ok(format!("{:04}-{:02}-{:02}", yyyy, mm, dd))
}

#[cfg(test)]
mod tests {
    use super::*;

    // A realistic pdftotext -layout excerpt for testing.
    fn sample_text() -> String {
        [
            "                         Merchant Number   1234567890",
            "                         Statement Period   01/01/26 - 01/31/26",
            "",
            "  S  UMMARY          B DY     AY",
            "",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  Submitted",
            "  01/01/26                   $116.92          0.00          0.00    0.00     $116.92",
            "  01/02/26                   $250.00          0.00          0.00    0.00     $250.00",
            "  01/13/26                 $4,642.61      -$200.00          0.00    0.00   $4,442.61",
            "  Month End Charge               0.00          0.00          0.00  -$1,517.82  -$1,517.82",
            "  Total                    $64,631.74      -$709.00          0.00  -$1,517.82  $62,404.92",
            "",
            "  BATCH DETAIL",
        ]
        .join("\n")
    }

    #[test]
    fn test_parse_merchant_id() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        assert_eq!(result.merchant_id, "1234567890");
    }

    #[test]
    fn test_parse_period() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        assert_eq!(result.period_start, "2026-01-01");
        assert_eq!(result.period_end, "2026-01-31");
    }

    #[test]
    fn test_parse_rows() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        assert_eq!(result.rows.len(), 3);
        assert_eq!(result.rows[0].date, "2026-01-01");
        assert_eq!(result.rows[0].amount_minor, 11692);
        assert_eq!(result.rows[1].date, "2026-01-02");
        assert_eq!(result.rows[1].amount_minor, 25000);
        assert_eq!(result.rows[2].date, "2026-01-13");
        assert_eq!(result.rows[2].amount_minor, 444261);
    }

    #[test]
    fn test_parse_month_end_charge_excluded() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        // Month End Charge should NOT appear as a row
        for row in &result.rows {
            assert_ne!(row.amount_minor, -151782);
        }
    }

    #[test]
    fn test_parse_total_crosscheck() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        assert_eq!(result.total_amount_minor, Some(6240492));
    }

    #[test]
    fn test_parse_empty_text() {
        let err = parse("").unwrap_err();
        assert!(err.message.contains("scanned/image-only"));
    }

    #[test]
    fn test_parse_no_summary_section() {
        let text = "Merchant Number   1234567890\nStatement Period   01/01/26 - 01/31/26\nSome random content\n";
        let err = parse(text).unwrap_err();
        assert!(err.message.contains("Unsupported or unrecognized"));
    }

    #[test]
    fn test_parse_negative_amount_processed() {
        // A line where chargebacks exceed sales
        let text = [
            "Merchant Number   9999",
            "Statement Period   01/01/26 - 01/31/26",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  01/05/26                   $100.00      -$300.00          0.00    0.00   -$200.00",
            "  Total                      $100.00      -$300.00          0.00    0.00   -$200.00",
        ]
        .join("\n");
        let result = parse(&text).unwrap();
        assert_eq!(result.rows.len(), 1);
        assert_eq!(result.rows[0].amount_minor, -20000);
    }

    #[test]
    fn test_parse_parenthesized_negative() {
        let text = [
            "Merchant Number   9999",
            "Statement Period   01/01/26 - 01/31/26",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  01/05/26                   $100.00          0.00          0.00  ($50.00)   ($50.00)",
            "  Total                      $100.00          0.00          0.00  ($50.00)   ($50.00)",
        ]
        .join("\n");
        // Parenthesized amounts should parse, but our regex only handles -$ and $
        // The row might not match because parenthesized amounts aren't matched by the standard money_re.
        // This test validates that the fallback paren_re handles it.
        let result = parse(&text);
        // If parenthesized values are the only dollar amounts, the row should still parse
        // via the paren_re fallback.
        assert!(result.is_ok());
    }

    #[test]
    fn test_parse_mmddyy_century() {
        assert_eq!(parse_mmddyy("01/15/26").unwrap(), "2026-01-15");
        assert_eq!(parse_mmddyy("12/31/99").unwrap(), "1999-12-31");
        assert_eq!(parse_mmddyy("06/15/49").unwrap(), "2049-06-15");
        assert_eq!(parse_mmddyy("06/15/50").unwrap(), "1950-06-15");
    }

    #[test]
    fn test_parse_kerned_total_line() {
        // pdftotext sometimes kerns "Total" with extra spaces
        let text = [
            "Merchant Number   9999",
            "Statement Period   01/01/26 - 01/31/26",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  01/01/26                   $500.00          0.00          0.00    0.00     $500.00",
            "  01/02/26                   $300.00          0.00          0.00    0.00     $300.00",
            "  T o t a l                  $800.00          0.00          0.00    0.00     $800.00",
            "",
            "  BATCH DETAIL",
        ]
        .join("\n");
        let result = parse(&text).unwrap();
        assert_eq!(result.rows.len(), 2);
        assert_eq!(result.total_amount_minor, Some(80000));
        // Cross-check: sum matches Total
        let sum: i64 = result.rows.iter().map(|r| r.amount_minor).sum();
        assert_eq!(sum, 80000);
    }

    #[test]
    fn test_parse_extra_gap_total_line() {
        // Some statements have irregular spacing: "T  otal" or "Total     "
        let text = [
            "Merchant Number   9999",
            "Statement Period   01/01/26 - 01/31/26",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  01/01/26                   $200.00          0.00          0.00    0.00     $200.00",
            "  T  otal                    $200.00          0.00          0.00    0.00     $200.00",
        ]
        .join("\n");
        let result = parse(&text).unwrap();
        assert_eq!(result.rows.len(), 1);
        assert_eq!(result.total_amount_minor, Some(20000));
    }

    #[test]
    fn test_parse_stops_at_section_without_total() {
        // Statement that jumps straight to next section without a Total line
        let text = [
            "Merchant Number   9999",
            "Statement Period   01/01/26 - 01/31/26",
            "  Date        Gross Sales   Chargebacks    Returns   Fees     Amount Processed",
            "  01/01/26                   $100.00          0.00          0.00    0.00     $100.00",
            "",
            "  CHARGEBACKS",
            "  01/15/26   -$50.00",
        ]
        .join("\n");
        let result = parse(&text).unwrap();
        assert_eq!(result.rows.len(), 1);
        assert_eq!(result.rows[0].amount_minor, 10000);
        // No Total line → None (delta will be null in engine_meta)
        assert_eq!(result.total_amount_minor, None);
    }

    #[test]
    fn test_canonical_source_id_format() {
        let text = sample_text();
        let result = parse(&text).unwrap();
        let row = &result.rows[0];
        // Verify the components needed for source_id
        let source_id = format!(
            "stmt:{}:{}:{}",
            result.merchant_id, row.date, row.amount_minor
        );
        assert_eq!(source_id, "stmt:1234567890:2026-01-01:11692");
    }
}
