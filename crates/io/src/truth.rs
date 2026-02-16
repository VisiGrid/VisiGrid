//! Canonical truth schema for financial verification.
//!
//! This module defines the **verification schema** — a minimal, deterministic
//! representation of financial transactions from external truth sources
//! (Stripe, Mercury, Brex, bank CSVs, SFTP feeds).
//!
//! Two core types:
//! - [`TruthTransaction`] — individual canonical transaction
//! - [`DailyTotals`] — deterministic daily aggregation (the real API surface)
//!
//! ## Amount convention
//!
//! All amounts are stored as `i64` in **micro-units** (1e-6 of the currency unit).
//! For example, $100.00 USD = 100_000_000 micro-units.
//!
//! This provides 6 decimal places of precision regardless of currency, avoiding
//! the need for per-currency scale tables (USD=2, JPY=0, BTC=8, etc.).
//!
//! CSV output renders amounts with exactly 6 decimal places for deterministic
//! hashing: `100_000_000` → `"100.000000"`.
//!
//! ## Design invariants
//!
//! - `amount_net` is always non-negative; `direction` carries the sign
//! - `amount_gross` and `fee_amount` are non-negative when present
//! - Aggregation always uses `occurred_at` (never `posted_at`)
//! - All decimal output uses fixed 6-decimal-place formatting
//! - Sorting is deterministic: date ASC, currency ASC, source_account ASC

use std::collections::BTreeMap;
use std::io::Write;

use blake3::Hasher;
use chrono::NaiveDate;
use serde::{Deserialize, Serialize};

/// Scale factor: 1 currency unit = 1_000_000 micro-units.
pub const MICRO_UNIT_SCALE: i64 = 1_000_000;

// ── Direction ───────────────────────────────────────────────────────

/// Transaction direction. Removes sign ambiguity from amounts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Direction {
    Credit,
    Debit,
}

impl Direction {
    /// Returns +1 for credit, -1 for debit. Used in signed aggregation.
    pub fn sign(&self) -> i64 {
        match self {
            Direction::Credit => 1,
            Direction::Debit => -1,
        }
    }

    pub fn as_str(&self) -> &'static str {
        match self {
            Direction::Credit => "credit",
            Direction::Debit => "debit",
        }
    }
}

impl std::fmt::Display for Direction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.as_str())
    }
}

// ── TruthTransaction ────────────────────────────────────────────────

/// A single canonical transaction from an external truth source.
///
/// Amounts are stored in **micro-units** (1e-6 of the currency unit) as `i64`.
/// All amounts are non-negative — `direction` carries the sign.
///
/// `raw_hash` is a BLAKE3 hash of the original source row, providing an
/// immutability anchor from source file → canonical row → daily totals → proof.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TruthTransaction {
    /// Source system identifier: "stripe", "mercury", "brex", "bank_csv"
    pub source: String,
    /// Account identifier within the source (e.g. Stripe account ID, bank name)
    pub source_account: String,
    /// Unique transaction ID from the source system
    pub source_id: String,
    /// When the transaction occurred (UTC). Aggregation key.
    pub occurred_at: NaiveDate,
    /// When the transaction posted/settled (optional, banks only)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub posted_at: Option<NaiveDate>,
    /// ISO 4217 currency code (uppercase)
    pub currency: String,
    /// Credit or debit
    pub direction: Direction,
    /// Gross amount in micro-units (non-negative). None if fee is unknown/embedded.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub amount_gross: Option<i64>,
    /// Fee in micro-units (non-negative). None if unknown.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fee_amount: Option<i64>,
    /// Net amount in micro-units (non-negative). Always required.
    pub amount_net: i64,
    /// Counterparty name/identifier (optional)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub counterparty: Option<String>,
    /// Human-readable description (optional)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    /// BLAKE3 hash of the original source row (hex-encoded)
    pub raw_hash: String,
}

/// Validation errors for a TruthTransaction.
#[derive(Debug, PartialEq)]
pub enum ValidationError {
    EmptySource,
    EmptySourceAccount,
    EmptySourceId,
    EmptyCurrency,
    InvalidCurrency(String),
    NegativeAmountNet,
    NegativeAmountGross,
    NegativeFeeAmount,
    EmptyRawHash,
    GrossNetFeeMismatch,
}

impl std::fmt::Display for ValidationError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::EmptySource => write!(f, "source must not be empty"),
            Self::EmptySourceAccount => write!(f, "source_account must not be empty"),
            Self::EmptySourceId => write!(f, "source_id must not be empty"),
            Self::EmptyCurrency => write!(f, "currency must not be empty"),
            Self::InvalidCurrency(c) => write!(f, "currency must be 3 uppercase letters, got: {c}"),
            Self::NegativeAmountNet => write!(f, "amount_net must be non-negative"),
            Self::NegativeAmountGross => write!(f, "amount_gross must be non-negative"),
            Self::NegativeFeeAmount => write!(f, "fee_amount must be non-negative"),
            Self::EmptyRawHash => write!(f, "raw_hash must not be empty"),
            Self::GrossNetFeeMismatch => {
                write!(f, "amount_gross - fee_amount != amount_net")
            }
        }
    }
}

impl TruthTransaction {
    /// Validate all schema invariants. Returns all errors found.
    pub fn validate(&self) -> Vec<ValidationError> {
        let mut errors = Vec::new();

        if self.source.is_empty() {
            errors.push(ValidationError::EmptySource);
        }
        if self.source_account.is_empty() {
            errors.push(ValidationError::EmptySourceAccount);
        }
        if self.source_id.is_empty() {
            errors.push(ValidationError::EmptySourceId);
        }
        if self.currency.is_empty() {
            errors.push(ValidationError::EmptyCurrency);
        } else if self.currency.len() != 3 || self.currency != self.currency.to_uppercase() {
            errors.push(ValidationError::InvalidCurrency(self.currency.clone()));
        }
        if self.amount_net < 0 {
            errors.push(ValidationError::NegativeAmountNet);
        }
        if let Some(gross) = self.amount_gross {
            if gross < 0 {
                errors.push(ValidationError::NegativeAmountGross);
            }
        }
        if let Some(fee) = self.fee_amount {
            if fee < 0 {
                errors.push(ValidationError::NegativeFeeAmount);
            }
        }
        if self.raw_hash.is_empty() {
            errors.push(ValidationError::EmptyRawHash);
        }

        // If both gross and fee are present, verify gross - fee == net
        if let (Some(gross), Some(fee)) = (self.amount_gross, self.fee_amount) {
            if gross >= 0 && fee >= 0 && gross - fee != self.amount_net {
                errors.push(ValidationError::GrossNetFeeMismatch);
            }
        }

        errors
    }

    /// Returns true if all invariants hold.
    pub fn is_valid(&self) -> bool {
        self.validate().is_empty()
    }
}

// ── Conversion helpers ──────────────────────────────────────────────

/// Convert a minor-unit amount (e.g. cents) to micro-units.
///
/// `minor_scale` is the number of decimal places in the source minor unit.
/// For USD (cents), `minor_scale = 2`: 15042 cents → 150_420_000 micro-units.
/// For JPY (no minor unit), `minor_scale = 0`: 1500 → 1_500_000_000 micro-units.
pub fn minor_to_micro(amount: i64, minor_scale: u32) -> i64 {
    let shift = 6u32.saturating_sub(minor_scale);
    amount * 10i64.pow(shift)
}

// ── DailyTotals ─────────────────────────────────────────────────────

/// Deterministic daily aggregation of truth transactions.
///
/// Grouped by (date, currency, source_account), sorted deterministically.
/// This is the primary comparison surface for reconciliation.
///
/// `total_net` is **signed**: credits are positive, debits are negative.
/// `total_fee` is always non-negative (fees are costs).
/// All amounts in micro-units.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct DailyTotals {
    pub date: NaiveDate,
    pub currency: String,
    pub source_account: String,
    /// Sum of (amount_net * direction.sign()) — signed total in micro-units.
    pub total_net: i64,
    /// Sum of (amount_gross * direction.sign()) when present, else derived from net.
    pub total_gross: i64,
    /// Sum of fee_amount (always non-negative) in micro-units.
    pub total_fee: i64,
    /// Number of transactions in this group.
    pub transaction_count: u32,
}

/// Compute daily totals from a slice of truth transactions.
///
/// Enforces single-account: returns an error if more than one `source_account`
/// is present in the input.
///
/// Output is deterministically sorted: date ASC, currency ASC, source_account ASC.
pub fn compute_daily_totals(
    transactions: &[TruthTransaction],
) -> Result<Vec<DailyTotals>, String> {
    if transactions.is_empty() {
        return Ok(Vec::new());
    }

    // Enforce single-account constraint
    let first_account = &transactions[0].source_account;
    if let Some(tx) = transactions
        .iter()
        .find(|tx| tx.source_account != *first_account)
    {
        return Err(format!(
            "single-account constraint violated: found '{}' and '{}'. \
             Use --allow-multi or run separately per account.",
            first_account, tx.source_account,
        ));
    }

    // Group by (date, currency, source_account) using BTreeMap for deterministic order
    type Key = (NaiveDate, String, String);
    let mut groups: BTreeMap<Key, DailyTotals> = BTreeMap::new();

    for tx in transactions {
        let key = (
            tx.occurred_at,
            tx.currency.clone(),
            tx.source_account.clone(),
        );

        let entry = groups.entry(key).or_insert_with(|| DailyTotals {
            date: tx.occurred_at,
            currency: tx.currency.clone(),
            source_account: tx.source_account.clone(),
            total_net: 0,
            total_gross: 0,
            total_fee: 0,
            transaction_count: 0,
        });

        let sign = tx.direction.sign();
        entry.total_net += tx.amount_net * sign;
        entry.total_gross += tx.amount_gross.unwrap_or(tx.amount_net) * sign;
        entry.total_fee += tx.fee_amount.unwrap_or(0);
        entry.transaction_count += 1;
    }

    // BTreeMap iteration is already sorted by key (date, currency, account)
    Ok(groups.into_values().collect())
}

// ── CSV output ──────────────────────────────────────────────────────

/// Format a micro-unit amount as a fixed 6-decimal-place string.
///
/// Micro-units are 1e-6 of the currency unit.
/// Example: 100_000_000 → "100.000000" ($100.00)
/// Example: 150_420_000 → "150.420000"
/// Example:       1_000 → "0.001000"
///
/// This fixed formatting prevents hash drift from "1.2" vs "1.200000".
fn format_amount(micro_units: i64) -> String {
    let is_negative = micro_units < 0;
    let abs = micro_units.unsigned_abs();
    let whole = abs / 1_000_000;
    let frac = abs % 1_000_000;
    if is_negative {
        format!("-{whole}.{frac:06}")
    } else {
        format!("{whole}.{frac:06}")
    }
}

/// Parse a fixed 6-decimal-place amount string back to micro-units.
///
/// Inverse of `format_amount`. Returns `None` on parse failure.
pub fn parse_amount(s: &str) -> Option<i64> {
    let (is_negative, s) = if let Some(rest) = s.strip_prefix('-') {
        (true, rest)
    } else {
        (false, s)
    };

    let (whole_str, frac_str) = s.split_once('.')?;
    let whole: u64 = whole_str.parse().ok()?;
    // Pad or truncate fractional part to exactly 6 digits
    let frac_padded = format!("{:0<6}", frac_str);
    if frac_padded.len() > 6 {
        return None; // too many decimal places
    }
    let frac: u64 = frac_padded.parse().ok()?;
    let abs_val = whole * 1_000_000 + frac;
    let val = abs_val as i64;
    Some(if is_negative { -val } else { val })
}

/// CSV header for truth_transactions.csv
const TRANSACTIONS_HEADER: &[&str] = &[
    "source",
    "source_account",
    "source_id",
    "occurred_at",
    "posted_at",
    "currency",
    "direction",
    "amount_gross",
    "fee_amount",
    "amount_net",
    "counterparty",
    "description",
    "raw_hash",
];

/// CSV header for truth_daily_totals.csv
const DAILY_TOTALS_HEADER: &[&str] = &[
    "date",
    "currency",
    "source_account",
    "total_gross",
    "total_fee",
    "total_net",
    "transaction_count",
];

/// Write truth transactions to CSV with deterministic formatting.
///
/// Transactions are sorted by (occurred_at, source_id) for deterministic output.
pub fn write_transactions_csv(
    transactions: &[TruthTransaction],
    writer: impl Write,
) -> Result<(), String> {
    let mut sorted: Vec<&TruthTransaction> = transactions.iter().collect();
    sorted.sort_by(|a, b| {
        a.occurred_at
            .cmp(&b.occurred_at)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    let mut csv = csv::WriterBuilder::new()
        .terminator(csv::Terminator::Any(b'\n'))
        .from_writer(writer);

    csv.write_record(TRANSACTIONS_HEADER)
        .map_err(|e| format!("CSV write error: {e}"))?;

    for tx in &sorted {
        csv.write_record(&[
            &tx.source,
            &tx.source_account,
            &tx.source_id,
            &tx.occurred_at.format("%Y-%m-%d").to_string(),
            &tx.posted_at
                .map(|d| d.format("%Y-%m-%d").to_string())
                .unwrap_or_default(),
            &tx.currency,
            tx.direction.as_str(),
            &tx.amount_gross.map(format_amount).unwrap_or_default(),
            &tx.fee_amount.map(format_amount).unwrap_or_default(),
            &format_amount(tx.amount_net),
            tx.counterparty.as_deref().unwrap_or(""),
            tx.description.as_deref().unwrap_or(""),
            &tx.raw_hash,
        ])
        .map_err(|e| format!("CSV write error: {e}"))?;
    }

    csv.flush().map_err(|e| format!("CSV flush error: {e}"))?;
    Ok(())
}

/// Write daily totals to CSV with deterministic formatting.
///
/// Totals must already be sorted (as returned by [`compute_daily_totals`]).
pub fn write_daily_totals_csv(
    totals: &[DailyTotals],
    writer: impl Write,
) -> Result<(), String> {
    let mut csv = csv::WriterBuilder::new()
        .terminator(csv::Terminator::Any(b'\n'))
        .from_writer(writer);

    csv.write_record(DAILY_TOTALS_HEADER)
        .map_err(|e| format!("CSV write error: {e}"))?;

    for row in totals {
        csv.write_record(&[
            row.date.format("%Y-%m-%d").to_string(),
            row.currency.clone(),
            row.source_account.clone(),
            format_amount(row.total_gross),
            format_amount(row.total_fee),
            format_amount(row.total_net),
            row.transaction_count.to_string(),
        ])
        .map_err(|e| format!("CSV write error: {e}"))?;
    }

    csv.flush().map_err(|e| format!("CSV flush error: {e}"))?;
    Ok(())
}

/// Read daily totals from a CSV file.
///
/// Parses the deterministic CSV format produced by [`write_daily_totals_csv`].
/// Used by `vgrid verify totals` to load both truth and warehouse files.
pub fn read_daily_totals_csv(reader: impl std::io::Read) -> Result<Vec<DailyTotals>, String> {
    let mut csv = csv::ReaderBuilder::new()
        .has_headers(true)
        .from_reader(reader);

    let mut totals = Vec::new();

    for (i, result) in csv.records().enumerate() {
        let record = result.map_err(|e| format!("CSV parse error at row {}: {e}", i + 1))?;

        if record.len() < 7 {
            return Err(format!(
                "row {} has {} columns, expected 7",
                i + 1,
                record.len()
            ));
        }

        let date = NaiveDate::parse_from_str(&record[0], "%Y-%m-%d")
            .map_err(|e| format!("row {}: invalid date '{}': {e}", i + 1, &record[0]))?;

        let total_gross = parse_amount(&record[3])
            .ok_or_else(|| format!("row {}: invalid total_gross '{}'", i + 1, &record[3]))?;
        let total_fee = parse_amount(&record[4])
            .ok_or_else(|| format!("row {}: invalid total_fee '{}'", i + 1, &record[4]))?;
        let total_net = parse_amount(&record[5])
            .ok_or_else(|| format!("row {}: invalid total_net '{}'", i + 1, &record[5]))?;
        let transaction_count: u32 = record[6]
            .parse()
            .map_err(|e| format!("row {}: invalid transaction_count '{}': {e}", i + 1, &record[6]))?;

        totals.push(DailyTotals {
            date,
            currency: record[1].to_string(),
            source_account: record[2].to_string(),
            total_gross,
            total_fee,
            total_net,
            transaction_count,
        });
    }

    Ok(totals)
}

// ── Hashing ─────────────────────────────────────────────────────────

/// Compute a BLAKE3 hash of a raw source row (for `raw_hash` field).
///
/// Input should be the original CSV row or JSON blob as bytes.
pub fn hash_raw_row(data: &[u8]) -> String {
    let mut hasher = Hasher::new();
    hasher.update(data);
    hasher.finalize().to_hex().to_string()
}

/// Compute a BLAKE3 hash of the daily totals CSV output.
///
/// This is the fingerprint used in proof artifacts.
pub fn hash_daily_totals(totals: &[DailyTotals]) -> Result<String, String> {
    let mut buf = Vec::new();
    write_daily_totals_csv(totals, &mut buf)?;
    Ok(hash_raw_row(&buf))
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;

    /// Helper: amounts in USD cents → micro-units (scale=2 → scale=6)
    fn usd(cents: i64) -> i64 {
        minor_to_micro(cents, 2)
    }

    fn make_tx(
        source_id: &str,
        date: &str,
        direction: Direction,
        amount_net: i64,
        amount_gross: Option<i64>,
        fee: Option<i64>,
    ) -> TruthTransaction {
        TruthTransaction {
            source: "stripe".to_string(),
            source_account: "acct_test123".to_string(),
            source_id: source_id.to_string(),
            occurred_at: NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
            posted_at: None,
            currency: "USD".to_string(),
            direction,
            amount_gross,
            fee_amount: fee,
            amount_net,
            counterparty: Some("Customer Inc".to_string()),
            description: Some("Payment".to_string()),
            raw_hash: hash_raw_row(source_id.as_bytes()),
        }
    }

    #[test]
    fn test_minor_to_micro() {
        // USD: 2 decimal places → 4 more zeros
        assert_eq!(minor_to_micro(100, 2), 1_000_000); // $1.00
        assert_eq!(minor_to_micro(15042, 2), 150_420_000); // $150.42
        // JPY: 0 decimal places → 6 more zeros
        assert_eq!(minor_to_micro(1500, 0), 1_500_000_000); // ¥1500
        // Already micro-units (scale=6) → no change
        assert_eq!(minor_to_micro(100_000_000, 6), 100_000_000);
    }

    #[test]
    fn test_validation_valid() {
        let tx = make_tx(
            "txn_001",
            "2026-01-15",
            Direction::Credit,
            usd(10000),
            Some(usd(10300)),
            Some(usd(300)),
        );
        assert!(tx.is_valid());
    }

    #[test]
    fn test_validation_empty_source() {
        let mut tx = make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None);
        tx.source = String::new();
        let errors = tx.validate();
        assert!(errors.contains(&ValidationError::EmptySource));
    }

    #[test]
    fn test_validation_bad_currency() {
        let mut tx = make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None);
        tx.currency = "usd".to_string();
        let errors = tx.validate();
        assert!(errors.contains(&ValidationError::InvalidCurrency("usd".to_string())));
    }

    #[test]
    fn test_validation_negative_net() {
        let mut tx = make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None);
        tx.amount_net = -100;
        let errors = tx.validate();
        assert!(errors.contains(&ValidationError::NegativeAmountNet));
    }

    #[test]
    fn test_validation_gross_net_fee_mismatch() {
        let tx = make_tx(
            "txn_001",
            "2026-01-15",
            Direction::Credit,
            usd(10000),
            Some(usd(10500)),
            Some(usd(300)),
        );
        let errors = tx.validate();
        assert!(errors.contains(&ValidationError::GrossNetFeeMismatch));
    }

    #[test]
    fn test_format_amount() {
        assert_eq!(format_amount(0), "0.000000");
        assert_eq!(format_amount(1_000_000), "1.000000");
        assert_eq!(format_amount(100_000_000), "100.000000");
        assert_eq!(format_amount(150_420_000), "150.420000");
        assert_eq!(format_amount(1_000), "0.001000");
        assert_eq!(format_amount(1), "0.000001");
        assert_eq!(format_amount(-5_000_000), "-5.000000");
        assert_eq!(format_amount(-10_000), "-0.010000");
    }

    #[test]
    fn test_parse_amount() {
        assert_eq!(parse_amount("0.000000"), Some(0));
        assert_eq!(parse_amount("1.000000"), Some(1_000_000));
        assert_eq!(parse_amount("100.000000"), Some(100_000_000));
        assert_eq!(parse_amount("150.420000"), Some(150_420_000));
        assert_eq!(parse_amount("0.001000"), Some(1_000));
        assert_eq!(parse_amount("-5.000000"), Some(-5_000_000));
        assert_eq!(parse_amount("-0.010000"), Some(-10_000));
        // Round-trip
        assert_eq!(parse_amount(&format_amount(123_456_789)), Some(123_456_789));
        assert_eq!(parse_amount(&format_amount(-42_000_000)), Some(-42_000_000));
    }

    #[test]
    fn test_daily_totals_basic() {
        let txs = vec![
            make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None),
            make_tx("txn_002", "2026-01-15", Direction::Credit, usd(5000), None, None),
            make_tx("txn_003", "2026-01-15", Direction::Debit, usd(2000), None, None),
        ];

        let totals = compute_daily_totals(&txs).unwrap();
        assert_eq!(totals.len(), 1);
        assert_eq!(totals[0].total_net, usd(13000)); // +10000 +5000 -2000
        assert_eq!(totals[0].transaction_count, 3);
    }

    #[test]
    fn test_daily_totals_multi_day() {
        let txs = vec![
            make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None),
            make_tx("txn_002", "2026-01-16", Direction::Credit, usd(5000), None, None),
        ];

        let totals = compute_daily_totals(&txs).unwrap();
        assert_eq!(totals.len(), 2);
        assert_eq!(totals[0].date, NaiveDate::from_ymd_opt(2026, 1, 15).unwrap());
        assert_eq!(totals[1].date, NaiveDate::from_ymd_opt(2026, 1, 16).unwrap());
    }

    #[test]
    fn test_daily_totals_with_fees() {
        let txs = vec![
            make_tx(
                "txn_001",
                "2026-01-15",
                Direction::Credit,
                usd(9700),
                Some(usd(10000)),
                Some(usd(300)),
            ),
            make_tx(
                "txn_002",
                "2026-01-15",
                Direction::Credit,
                usd(4850),
                Some(usd(5000)),
                Some(usd(150)),
            ),
        ];

        let totals = compute_daily_totals(&txs).unwrap();
        assert_eq!(totals.len(), 1);
        assert_eq!(totals[0].total_net, usd(14550)); // 9700 + 4850
        assert_eq!(totals[0].total_gross, usd(15000)); // 10000 + 5000
        assert_eq!(totals[0].total_fee, usd(450)); // 300 + 150
    }

    #[test]
    fn test_single_account_enforcement() {
        let mut txs = vec![make_tx(
            "txn_001",
            "2026-01-15",
            Direction::Credit,
            usd(10000),
            None,
            None,
        )];
        txs[0].source_account = "acct_a".to_string();

        let mut tx2 = make_tx("txn_002", "2026-01-15", Direction::Credit, usd(5000), None, None);
        tx2.source_account = "acct_b".to_string();
        txs.push(tx2);

        let err = compute_daily_totals(&txs).unwrap_err();
        assert!(err.contains("single-account constraint violated"));
    }

    #[test]
    fn test_deterministic_output_order_independent() {
        // Core invariant: same transactions in different order → identical CSV bytes
        let tx1 = make_tx(
            "txn_001",
            "2026-01-15",
            Direction::Credit,
            usd(10000),
            Some(usd(10300)),
            Some(usd(300)),
        );
        let tx2 = make_tx("txn_002", "2026-01-15", Direction::Debit, usd(2000), None, None);
        let tx3 = make_tx("txn_003", "2026-01-16", Direction::Credit, usd(5000), None, None);

        // Order A: 1, 2, 3
        let order_a = vec![tx1.clone(), tx2.clone(), tx3.clone()];
        let totals_a = compute_daily_totals(&order_a).unwrap();
        let mut csv_a = Vec::new();
        write_daily_totals_csv(&totals_a, &mut csv_a).unwrap();

        // Order B: 3, 1, 2
        let order_b = vec![tx3.clone(), tx1.clone(), tx2.clone()];
        let totals_b = compute_daily_totals(&order_b).unwrap();
        let mut csv_b = Vec::new();
        write_daily_totals_csv(&totals_b, &mut csv_b).unwrap();

        // Order C: 2, 3, 1
        let order_c = vec![tx2, tx3, tx1];
        let totals_c = compute_daily_totals(&order_c).unwrap();
        let mut csv_c = Vec::new();
        write_daily_totals_csv(&totals_c, &mut csv_c).unwrap();

        assert_eq!(csv_a, csv_b, "order A vs B produced different bytes");
        assert_eq!(csv_b, csv_c, "order B vs C produced different bytes");

        // Also verify the hash is stable
        let hash_a = hash_daily_totals(&totals_a).unwrap();
        let hash_b = hash_daily_totals(&totals_b).unwrap();
        let hash_c = hash_daily_totals(&totals_c).unwrap();
        assert_eq!(hash_a, hash_b);
        assert_eq!(hash_b, hash_c);
    }

    #[test]
    fn test_deterministic_transactions_csv_order_independent() {
        let tx1 = make_tx("txn_001", "2026-01-15", Direction::Credit, usd(10000), None, None);
        let tx2 = make_tx("txn_002", "2026-01-15", Direction::Debit, usd(2000), None, None);

        let mut csv_a = Vec::new();
        write_transactions_csv(&[tx1.clone(), tx2.clone()], &mut csv_a).unwrap();

        let mut csv_b = Vec::new();
        write_transactions_csv(&[tx2, tx1], &mut csv_b).unwrap();

        assert_eq!(csv_a, csv_b, "transaction CSV must be order-independent");
    }

    #[test]
    fn test_hash_raw_row() {
        let hash = hash_raw_row(b"txn_001,2026-01-15,10000,USD");
        assert_eq!(hash.len(), 64); // BLAKE3 hex = 64 chars
        let hash2 = hash_raw_row(b"txn_001,2026-01-15,10000,USD");
        assert_eq!(hash, hash2);
    }

    #[test]
    fn test_empty_input() {
        let totals = compute_daily_totals(&[]).unwrap();
        assert!(totals.is_empty());

        let mut csv = Vec::new();
        write_daily_totals_csv(&[], &mut csv).unwrap();
        let output = String::from_utf8(csv).unwrap();
        assert!(output.starts_with("date,currency,source_account"));
        assert_eq!(output.lines().count(), 1);
    }

    #[test]
    fn test_csv_format_snapshot() {
        let txs = vec![make_tx(
            "txn_001",
            "2026-01-15",
            Direction::Credit,
            usd(10000),     // $100.00 net
            Some(usd(10300)), // $103.00 gross
            Some(usd(300)),   // $3.00 fee
        )];

        let totals = compute_daily_totals(&txs).unwrap();
        let mut csv = Vec::new();
        write_daily_totals_csv(&totals, &mut csv).unwrap();
        let output = String::from_utf8(csv).unwrap();

        let lines: Vec<&str> = output.lines().collect();
        assert_eq!(lines.len(), 2);
        assert_eq!(
            lines[0],
            "date,currency,source_account,total_gross,total_fee,total_net,transaction_count"
        );
        assert_eq!(
            lines[1],
            "2026-01-15,USD,acct_test123,103.000000,3.000000,100.000000,1"
        );
    }

    #[test]
    fn test_read_daily_totals_csv_roundtrip() {
        let txs = vec![
            make_tx(
                "txn_001",
                "2026-01-15",
                Direction::Credit,
                usd(10000),
                Some(usd(10300)),
                Some(usd(300)),
            ),
            make_tx("txn_002", "2026-01-15", Direction::Debit, usd(2000), None, None),
            make_tx("txn_003", "2026-01-16", Direction::Credit, usd(5000), None, None),
        ];

        let totals = compute_daily_totals(&txs).unwrap();

        // Write to bytes
        let mut buf = Vec::new();
        write_daily_totals_csv(&totals, &mut buf).unwrap();

        // Read back
        let parsed = read_daily_totals_csv(buf.as_slice()).unwrap();
        assert_eq!(totals, parsed);
    }

    #[test]
    fn test_direction_sign() {
        assert_eq!(Direction::Credit.sign(), 1);
        assert_eq!(Direction::Debit.sign(), -1);
    }
}
