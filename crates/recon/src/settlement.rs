//! Settlement classification — classifies recon groups into settlement states.
//!
//! Four states: Matched, Pending, Stale, Error.
//! - Matched: fully reconciled (MatchedTwoWay / MatchedThreeWay)
//! - Pending: unmatched but within SLA window (age ≤ sla_days)
//! - Stale: unmatched and past SLA window (age > sla_days)
//! - Error: structural mismatch (amount/timing mismatch, any future non-matched bucket).
//!          No age — errors are invariant violations, not timing lag.

use chrono::NaiveDate;

use crate::config::SettlementConfig;
use crate::model::{
    ClassifiedResult, ReconBucket, SettlementClassification, SettlementState, SettlementSummary,
};

/// Classify all results with settlement state, mutating in place.
pub fn classify_settlement(results: &mut [ClassifiedResult], config: &SettlementConfig) {
    for result in results.iter_mut() {
        result.settlement = Some(classify_one(result, config));
    }
}

fn classify_one(
    result: &ClassifiedResult,
    config: &SettlementConfig,
) -> SettlementClassification {
    match &result.bucket {
        // Fully matched — no age needed
        ReconBucket::MatchedTwoWay | ReconBucket::MatchedThreeWay => SettlementClassification {
            state: SettlementState::Matched,
            age_days: None,
            sla_days: config.sla_days,
        },

        // Unmatched — pending vs stale based on age
        ReconBucket::ProcessorLedgerOnly
        | ReconBucket::ProcessorBankOnly
        | ReconBucket::LedgerOnly
        | ReconBucket::BankOnly => {
            match compute_age_for_only(result, config.reference_date) {
                Some(age) => SettlementClassification {
                    state: if age <= config.sla_days as i64 {
                        SettlementState::Pending
                    } else {
                        SettlementState::Stale
                    },
                    age_days: Some(age),
                    sla_days: config.sla_days,
                },
                None => SettlementClassification {
                    state: SettlementState::Pending, // can't compute → don't false-fail
                    age_days: None,
                    sla_days: config.sla_days,
                },
            }
        }

        // Everything else = structural error. No age — errors are invariant violations,
        // not timing lag. Including age here would conflate two different concerns.
        _ => SettlementClassification {
            state: SettlementState::Error,
            age_days: None,
            sla_days: config.sla_days,
        },
    }
}

/// For unmatched groups: use the single aggregate's date.
fn compute_age_for_only(result: &ClassifiedResult, reference_date: NaiveDate) -> Option<i64> {
    result
        .aggregates
        .values()
        .next()
        .map(|a| (reference_date - a.date).num_days())
}

/// Compute settlement summary counts from classified results.
pub fn compute_settlement_summary(results: &[ClassifiedResult]) -> SettlementSummary {
    let mut matched = 0;
    let mut pending = 0;
    let mut stale = 0;
    let mut errors = 0;

    for r in results {
        if let Some(ref s) = r.settlement {
            match s.state {
                SettlementState::Matched => matched += 1,
                SettlementState::Pending => pending += 1,
                SettlementState::Stale => stale += 1,
                SettlementState::Error => errors += 1,
            }
        }
    }

    SettlementSummary {
        matched,
        pending,
        stale,
        errors,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::SettlementClock;
    use crate::model::{Aggregate, Deltas};
    use std::collections::HashMap;

    fn make_config(ref_date: &str, sla: u32) -> SettlementConfig {
        SettlementConfig {
            reference_date: NaiveDate::parse_from_str(ref_date, "%Y-%m-%d").unwrap(),
            sla_days: sla,
            clock: SettlementClock::Processor,
        }
    }

    fn make_result(bucket: ReconBucket, aggregates: HashMap<String, Aggregate>) -> ClassifiedResult {
        ClassifiedResult {
            bucket,
            match_key: "k".into(),
            currency: "USD".into(),
            aggregates,
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
            settlement: None,
        }
    }

    fn agg(role: &str, date: &str) -> Aggregate {
        Aggregate {
            role: role.into(),
            match_key: "k".into(),
            currency: "USD".into(),
            date: NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
            total_cents: 1000,
            record_count: 1,
            record_ids: vec!["id_1".into()],
        }
    }

    #[test]
    fn matched_gets_matched_state() {
        let config = make_config("2026-01-31", 5);
        let mut results = vec![make_result(ReconBucket::MatchedTwoWay, HashMap::new())];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Matched);
        assert_eq!(s.age_days, None);
    }

    #[test]
    fn three_way_matched_gets_matched_state() {
        let config = make_config("2026-01-31", 5);
        let mut results = vec![make_result(ReconBucket::MatchedThreeWay, HashMap::new())];
        classify_settlement(&mut results, &config);
        assert_eq!(
            results[0].settlement.as_ref().unwrap().state,
            SettlementState::Matched,
        );
    }

    #[test]
    fn only_within_sla_is_pending() {
        let config = make_config("2026-01-31", 5);
        let mut aggs = HashMap::new();
        aggs.insert("processor".into(), agg("processor", "2026-01-28")); // 3 days old
        let mut results = vec![make_result(ReconBucket::ProcessorLedgerOnly, aggs)];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Pending);
        assert_eq!(s.age_days, Some(3));
    }

    #[test]
    fn only_past_sla_is_stale() {
        let config = make_config("2026-01-31", 5);
        let mut aggs = HashMap::new();
        aggs.insert("processor".into(), agg("processor", "2026-01-20")); // 11 days old
        let mut results = vec![make_result(ReconBucket::ProcessorLedgerOnly, aggs)];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Stale);
        assert_eq!(s.age_days, Some(11));
    }

    #[test]
    fn only_at_exact_sla_boundary_is_pending() {
        let config = make_config("2026-01-31", 5);
        let mut aggs = HashMap::new();
        aggs.insert("bank".into(), agg("bank", "2026-01-26")); // exactly 5 days
        let mut results = vec![make_result(ReconBucket::BankOnly, aggs)];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Pending);
        assert_eq!(s.age_days, Some(5));
    }

    #[test]
    fn amount_mismatch_is_error_with_no_age() {
        let config = make_config("2026-01-31", 5);
        let mut aggs = HashMap::new();
        aggs.insert("processor".into(), agg("processor", "2026-01-28"));
        aggs.insert("ledger".into(), agg("ledger", "2026-01-29"));
        let mut results = vec![make_result(ReconBucket::AmountMismatch, aggs)];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Error);
        assert_eq!(s.age_days, None); // errors don't carry age
    }

    #[test]
    fn timing_mismatch_is_error() {
        let config = make_config("2026-01-31", 5);
        let mut aggs = HashMap::new();
        aggs.insert("processor".into(), agg("processor", "2026-01-25"));
        let mut results = vec![make_result(ReconBucket::TimingMismatch, aggs)];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Error);
        assert_eq!(s.age_days, None);
    }

    #[test]
    fn no_aggregates_defaults_to_pending() {
        let config = make_config("2026-01-31", 5);
        let mut results = vec![make_result(ReconBucket::LedgerOnly, HashMap::new())];
        classify_settlement(&mut results, &config);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Pending);
        assert_eq!(s.age_days, None);
    }

    #[test]
    fn summary_counts() {
        let config = make_config("2026-01-31", 5);
        let mut aggs_pending = HashMap::new();
        aggs_pending.insert("processor".into(), agg("processor", "2026-01-28"));
        let mut aggs_stale = HashMap::new();
        aggs_stale.insert("bank".into(), agg("bank", "2026-01-10"));

        let mut results = vec![
            make_result(ReconBucket::MatchedTwoWay, HashMap::new()),
            make_result(ReconBucket::MatchedThreeWay, HashMap::new()),
            make_result(ReconBucket::ProcessorLedgerOnly, aggs_pending),
            make_result(ReconBucket::BankOnly, aggs_stale),
            make_result(ReconBucket::AmountMismatch, HashMap::new()),
        ];
        classify_settlement(&mut results, &config);
        let summary = compute_settlement_summary(&results);
        assert_eq!(summary.matched, 2);
        assert_eq!(summary.pending, 1);
        assert_eq!(summary.stale, 1);
        assert_eq!(summary.errors, 1);
    }

    // ── Payout-only aging simulation ─────────────────────────────────
    // Simulates a real Stripe payout that has no bank deposit yet.
    // At day 3 (within SLA=5) → Pending. At day 8 (past SLA) → Stale.

    #[test]
    fn payout_only_pending_then_stale_as_reference_advances() {
        // Day 3: payout is 3 days old, SLA is 5 → Pending
        let config_day3 = SettlementConfig {
            reference_date: NaiveDate::parse_from_str("2026-01-20", "%Y-%m-%d").unwrap(),
            sla_days: 5,
            clock: SettlementClock::Processor,
        };
        let mut aggs = HashMap::new();
        aggs.insert("processor".into(), agg("processor", "2026-01-17")); // payout date
        let mut results = vec![make_result(ReconBucket::ProcessorBankOnly, aggs.clone())];
        classify_settlement(&mut results, &config_day3);
        let s = results[0].settlement.as_ref().unwrap();
        assert_eq!(s.state, SettlementState::Pending);
        assert_eq!(s.age_days, Some(3));

        // Day 8: same payout is now 8 days old, SLA is 5 → Stale
        let config_day8 = SettlementConfig {
            reference_date: NaiveDate::parse_from_str("2026-01-25", "%Y-%m-%d").unwrap(),
            sla_days: 5,
            clock: SettlementClock::Processor,
        };
        let mut results2 = vec![make_result(ReconBucket::ProcessorBankOnly, aggs)];
        classify_settlement(&mut results2, &config_day8);
        let s2 = results2[0].settlement.as_ref().unwrap();
        assert_eq!(s2.state, SettlementState::Stale);
        assert_eq!(s2.age_days, Some(8));
    }
}
