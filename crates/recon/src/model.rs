use std::collections::HashMap;

use chrono::NaiveDate;
use serde::Serialize;

// ---------------------------------------------------------------------------
// Input
// ---------------------------------------------------------------------------

/// A single normalized row from any role's CSV.
#[derive(Debug, Clone)]
pub struct ReconRow {
    pub role: String,
    pub record_id: String,
    pub match_key: String,
    pub date: NaiveDate,
    pub amount_cents: i64,
    pub currency: String,
    pub kind: String,
    pub raw_fields: HashMap<String, String>,
}

/// Pre-loaded records grouped by role name.
pub struct ReconInput {
    pub records: HashMap<String, Vec<ReconRow>>,
}

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

/// Aggregate key = (match_key, currency). Never matches across currencies.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct AggregateKey {
    pub match_key: String,
    pub currency: String,
}

/// Aggregated records sharing the same (match_key, currency).
#[derive(Debug, Clone, Serialize)]
pub struct Aggregate {
    pub role: String,
    pub match_key: String,
    pub currency: String,
    pub date: NaiveDate,
    pub total_cents: i64,
    pub record_count: usize,
    pub record_ids: Vec<String>,
}

// ---------------------------------------------------------------------------
// Pair matching
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize)]
pub struct MatchedPair {
    pub left: Aggregate,
    pub right: Aggregate,
    pub delta_cents: i64,
    pub date_offset_days: i32,
    pub within_tolerance: bool,
    pub within_window: bool,
}

#[derive(Debug)]
pub struct PairMatchOutput {
    pub matched: Vec<MatchedPair>,
    pub left_only: Vec<Aggregate>,
    pub right_only: Vec<Aggregate>,
}

// ---------------------------------------------------------------------------
// Classification
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum ReconBucket {
    MatchedTwoWay,
    MatchedThreeWay,
    ProcessorLedgerOnly,
    ProcessorBankOnly,
    LedgerOnly,
    BankOnly,
    AmountMismatch,
    TimingMismatch,
}

impl std::fmt::Display for ReconBucket {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MatchedTwoWay => write!(f, "matched_two_way"),
            Self::MatchedThreeWay => write!(f, "matched_three_way"),
            Self::ProcessorLedgerOnly => write!(f, "processor_ledger_only"),
            Self::ProcessorBankOnly => write!(f, "processor_bank_only"),
            Self::LedgerOnly => write!(f, "ledger_only"),
            Self::BankOnly => write!(f, "bank_only"),
            Self::AmountMismatch => write!(f, "amount_mismatch"),
            Self::TimingMismatch => write!(f, "timing_mismatch"),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct Deltas {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub delta_cents: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date_offset_days: Option<i32>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ClassifiedResult {
    pub bucket: ReconBucket,
    pub match_key: String,
    pub currency: String,
    pub aggregates: HashMap<String, Aggregate>,
    pub deltas: Deltas,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settlement: Option<SettlementClassification>,
}

// ---------------------------------------------------------------------------
// Settlement classification
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum SettlementState {
    Matched,
    Pending,
    Stale,
    Error,
}

impl std::fmt::Display for SettlementState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Matched => write!(f, "matched"),
            Self::Pending => write!(f, "pending"),
            Self::Stale => write!(f, "stale"),
            Self::Error => write!(f, "error"),
        }
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct SettlementClassification {
    pub state: SettlementState,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub age_days: Option<i64>,
    pub sla_days: u32,
}

#[derive(Debug, Clone, Serialize)]
pub struct SettlementSummary {
    pub matched: usize,
    pub pending: usize,
    pub stale: usize,
    pub errors: usize,
}

// ---------------------------------------------------------------------------
// Summary + Output
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize)]
pub struct ReconSummary {
    pub total_groups: usize,
    pub matched: usize,
    pub amount_mismatches: usize,
    pub timing_mismatches: usize,
    pub left_only: usize,
    pub right_only: usize,
    pub bucket_counts: HashMap<String, usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settlement: Option<SettlementSummary>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ReconResult {
    pub meta: ReconMeta,
    pub summary: ReconSummary,
    pub groups: Vec<ClassifiedResult>,
    #[serde(skip_serializing_if = "DerivedOutputs::is_empty")]
    pub derived: DerivedOutputs,
}

// ---------------------------------------------------------------------------
// Derived outputs — future-proof surface for computed analyses
// ---------------------------------------------------------------------------

/// Derived analyses computed from recon groups. Each field is an optional
/// vec that only serializes when populated. Initially ships empty — the
/// shape exists so downstream consumers (Rails, React) can type against it
/// without a binary update when implementations land.
#[derive(Debug, Clone, Default, Serialize)]
pub struct DerivedOutputs {
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub payout_rollup: Vec<serde_json::Value>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub clearing_delta: Vec<serde_json::Value>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub revenue_rollforward: Vec<serde_json::Value>,
}

impl DerivedOutputs {
    pub fn is_empty(&self) -> bool {
        self.payout_rollup.is_empty()
            && self.clearing_delta.is_empty()
            && self.revenue_rollforward.is_empty()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct ReconMeta {
    pub config_name: String,
    pub way: u8,
    pub engine_version: String,
    pub run_at: String,
    /// Which role's date was used as the settlement clock.
    /// Only present when settlement classification is enabled.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settlement_clock: Option<crate::config::SettlementClock>,
}
