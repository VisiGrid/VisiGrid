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

#[derive(Debug, Clone, Serialize)]
pub struct DerivedDataset {
    pub schema: &'static str,
    pub version: u32,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub rows: Vec<serde_json::Value>,
}

impl DerivedDataset {
    pub fn new(schema: &'static str) -> Self {
        Self { schema, version: 1, rows: vec![] }
    }
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct DerivedOutputs {
    #[serde(skip_serializing_if = "DerivedDataset::is_empty")]
    pub payout_rollup: DerivedDataset,
    #[serde(skip_serializing_if = "DerivedDataset::is_empty")]
    pub clearing_delta: DerivedDataset,
    #[serde(skip_serializing_if = "DerivedDataset::is_empty")]
    pub revenue_rollforward: DerivedDataset,
}

impl Default for DerivedOutputs {
    fn default() -> Self {
        Self {
            payout_rollup: DerivedDataset::new("payout_rollup"),
            clearing_delta: DerivedDataset::new("clearing_delta"),
            revenue_rollforward: DerivedDataset::new("revenue_rollforward"),
        }
    }
}

impl DerivedOutputs {
    pub fn is_empty(&self) -> bool {
        self.payout_rollup.is_empty()
            && self.clearing_delta.is_empty()
            && self.revenue_rollforward.is_empty()
    }
}

// ---------------------------------------------------------------------------
// Composite result types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum StepStatus {
    Pass,
    Warn,
    Fail,
    Error,
}

impl StepStatus {
    pub fn from_recon_result(result: &ReconResult) -> Self {
        let s = &result.summary;

        // Settlement-aware path: settlement config changes the exit semantics
        if let Some(ref settlement) = s.settlement {
            if settlement.errors > 0 || s.amount_mismatches > 0
                || s.left_only > 0 || s.right_only > 0
            {
                return Self::Fail;
            }
            if settlement.stale > 0 {
                return Self::Warn;
            }
            // pending only → pass (matches CLI: pending-only is not an error)
            return Self::Pass;
        }

        // Non-settlement path: any mismatch is a failure (matches CLI exit 1)
        if s.amount_mismatches > 0 || s.timing_mismatches > 0
            || s.left_only > 0 || s.right_only > 0
        {
            return Self::Fail;
        }

        Self::Pass
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct StepResult {
    pub name: String,
    pub status: StepStatus,
    pub duration_ms: u64,
    pub config_path: String,
    pub result: ReconResult,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum CompositeVerdict {
    Pass,
    Warn,
    Fail,
}

impl CompositeVerdict {
    pub fn from_steps(steps: &[StepResult]) -> Self {
        if steps
            .iter()
            .any(|s| matches!(s.status, StepStatus::Fail | StepStatus::Error))
        {
            Self::Fail
        } else if steps.iter().any(|s| matches!(s.status, StepStatus::Warn)) {
            Self::Warn
        } else {
            Self::Pass
        }
    }

    /// The CLI exit code for this verdict.
    /// Fail → 1 (EXIT_RECON_MISMATCH), Warn → 61 (EXIT_RECON_STALE), Pass → 0.
    pub fn exit_code(&self) -> u8 {
        match self {
            Self::Fail => 1,
            Self::Warn => 61,
            Self::Pass => 0,
        }
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct CompositeResult {
    pub name: String,
    pub engine_version: String,
    pub run_at: String,
    pub verdict: CompositeVerdict,
    pub exit_code: u8,
    pub steps: Vec<StepResult>,
}

// ---------------------------------------------------------------------------
// Meta
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn make_summary(
        matched: usize,
        amount_mismatches: usize,
        timing_mismatches: usize,
        left_only: usize,
        right_only: usize,
        settlement: Option<SettlementSummary>,
    ) -> ReconResult {
        ReconResult {
            meta: ReconMeta {
                config_name: "test".into(),
                way: 2,
                engine_version: "0.0.0".into(),
                run_at: "2026-01-01T00:00:00Z".into(),
                settlement_clock: None,
            },
            summary: ReconSummary {
                total_groups: matched + amount_mismatches + timing_mismatches + left_only + right_only,
                matched,
                amount_mismatches,
                timing_mismatches,
                left_only,
                right_only,
                bucket_counts: HashMap::new(),
                settlement,
            },
            groups: vec![],
            derived: DerivedOutputs::default(),
        }
    }

    #[test]
    fn step_status_pass() {
        let r = make_summary(5, 0, 0, 0, 0, None);
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Pass);
    }

    #[test]
    fn step_status_fail_on_amount_mismatch() {
        let r = make_summary(3, 1, 0, 0, 0, None);
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Fail);
    }

    #[test]
    fn step_status_fail_on_left_only() {
        let r = make_summary(3, 0, 0, 1, 0, None);
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Fail);
    }

    #[test]
    fn step_status_fail_on_right_only() {
        let r = make_summary(3, 0, 0, 0, 2, None);
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Fail);
    }

    #[test]
    fn step_status_fail_on_timing_without_settlement() {
        // Without settlement config, timing mismatches are failures (matches CLI exit 1)
        let r = make_summary(3, 0, 1, 0, 0, None);
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Fail);
    }

    #[test]
    fn step_status_fail_on_settlement_errors() {
        let settlement = SettlementSummary { matched: 1, pending: 0, stale: 0, errors: 1 };
        let r = make_summary(2, 0, 0, 0, 0, Some(settlement));
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Fail);
    }

    #[test]
    fn step_status_warn_on_settlement_stale() {
        let settlement = SettlementSummary { matched: 1, pending: 0, stale: 1, errors: 0 };
        let r = make_summary(2, 0, 0, 0, 0, Some(settlement));
        assert_eq!(StepStatus::from_recon_result(&r), StepStatus::Warn);
    }

    #[test]
    fn composite_verdict_all_pass() {
        let steps = vec![
            StepResult {
                name: "a".into(),
                status: StepStatus::Pass,
                duration_ms: 1,
                config_path: "a.toml".into(),
                result: make_summary(1, 0, 0, 0, 0, None),
            },
            StepResult {
                name: "b".into(),
                status: StepStatus::Pass,
                duration_ms: 1,
                config_path: "b.toml".into(),
                result: make_summary(1, 0, 0, 0, 0, None),
            },
        ];
        assert_eq!(CompositeVerdict::from_steps(&steps), CompositeVerdict::Pass);
    }

    #[test]
    fn composite_verdict_any_fail() {
        let steps = vec![
            StepResult {
                name: "a".into(),
                status: StepStatus::Pass,
                duration_ms: 1,
                config_path: "a.toml".into(),
                result: make_summary(1, 0, 0, 0, 0, None),
            },
            StepResult {
                name: "b".into(),
                status: StepStatus::Fail,
                duration_ms: 1,
                config_path: "b.toml".into(),
                result: make_summary(0, 1, 0, 0, 0, None),
            },
        ];
        assert_eq!(CompositeVerdict::from_steps(&steps), CompositeVerdict::Fail);
    }

    #[test]
    fn composite_verdict_warn_without_fail() {
        let settlement_stale = SettlementSummary { matched: 1, pending: 0, stale: 1, errors: 0 };
        let steps = vec![
            StepResult {
                name: "a".into(),
                status: StepStatus::Pass,
                duration_ms: 1,
                config_path: "a.toml".into(),
                result: make_summary(1, 0, 0, 0, 0, None),
            },
            StepResult {
                name: "b".into(),
                status: StepStatus::Warn,
                duration_ms: 1,
                config_path: "b.toml".into(),
                result: make_summary(2, 0, 0, 0, 0, Some(settlement_stale)),
            },
        ];
        assert_eq!(CompositeVerdict::from_steps(&steps), CompositeVerdict::Warn);
    }

    #[test]
    fn composite_verdict_error_means_fail() {
        let steps = vec![
            StepResult {
                name: "a".into(),
                status: StepStatus::Error,
                duration_ms: 0,
                config_path: "a.toml".into(),
                result: make_summary(0, 0, 0, 0, 0, None),
            },
        ];
        assert_eq!(CompositeVerdict::from_steps(&steps), CompositeVerdict::Fail);
    }

    // -----------------------------------------------------------------
    // Lockstep: StepStatus must agree with CLI exit code semantics
    // -----------------------------------------------------------------

    /// Reference implementation of CLI exit-code logic for a single recon result.
    /// Returns: 0 = pass, 1 = fail/mismatch, 61 = stale/warn.
    fn cli_exit_code(result: &ReconResult) -> u8 {
        let s = &result.summary;

        // Settlement-aware path (matches cmd_recon_run_single)
        if let Some(ref settlement) = s.settlement {
            if settlement.errors > 0 {
                return 1; // EXIT_RECON_MISMATCH
            }
            if settlement.stale > 0 {
                return 61; // EXIT_RECON_STALE
            }
            return 0;
        }

        // Fallback: original logic
        if s.amount_mismatches > 0 || s.timing_mismatches > 0
            || s.left_only > 0 || s.right_only > 0
        {
            return 1; // EXIT_RECON_MISMATCH
        }

        0
    }

    /// Map StepStatus to the expected exit code.
    fn status_to_exit(status: &StepStatus) -> u8 {
        match status {
            StepStatus::Pass => 0,
            StepStatus::Warn => 61,
            StepStatus::Fail | StepStatus::Error => 1,
        }
    }

    #[test]
    fn lockstep_step_status_matches_cli_exit_code() {
        // Exhaustive scenarios: each (summary shape, expected agreement)
        let cases: Vec<(ReconResult, &str)> = vec![
            (make_summary(5, 0, 0, 0, 0, None), "all matched, no settlement"),
            (make_summary(3, 1, 0, 0, 0, None), "amount mismatch"),
            (make_summary(3, 0, 1, 0, 0, None), "timing mismatch"),
            (make_summary(3, 0, 0, 1, 0, None), "left only"),
            (make_summary(3, 0, 0, 0, 2, None), "right only"),
            (
                make_summary(2, 0, 0, 0, 0, Some(SettlementSummary {
                    matched: 2, pending: 0, stale: 0, errors: 0,
                })),
                "settlement all matched",
            ),
            (
                make_summary(2, 0, 0, 0, 0, Some(SettlementSummary {
                    matched: 1, pending: 0, stale: 1, errors: 0,
                })),
                "settlement stale",
            ),
            (
                make_summary(2, 0, 0, 0, 0, Some(SettlementSummary {
                    matched: 1, pending: 0, stale: 0, errors: 1,
                })),
                "settlement errors",
            ),
            (
                make_summary(2, 0, 0, 0, 0, Some(SettlementSummary {
                    matched: 1, pending: 1, stale: 0, errors: 0,
                })),
                "settlement pending only",
            ),
        ];

        for (result, label) in &cases {
            let status = StepStatus::from_recon_result(result);
            let step_exit = status_to_exit(&status);
            let cli_exit = cli_exit_code(result);
            assert_eq!(
                step_exit, cli_exit,
                "lockstep mismatch for '{}': StepStatus::{:?} (exit {}) vs CLI exit {}",
                label, status, step_exit, cli_exit,
            );
        }
    }
}
