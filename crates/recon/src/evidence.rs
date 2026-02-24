use std::collections::HashMap;

use crate::model::{ClassifiedResult, ReconBucket, ReconSummary};
use crate::settlement;

/// Compute summary statistics from classified results.
pub fn compute_summary(results: &[ClassifiedResult]) -> ReconSummary {
    let mut bucket_counts: HashMap<String, usize> = HashMap::new();
    let mut matched = 0;
    let mut amount_mismatches = 0;
    let mut timing_mismatches = 0;
    let mut ambiguous = 0;
    let mut left_only = 0;
    let mut right_only = 0;

    for r in results {
        *bucket_counts.entry(r.bucket.to_string()).or_insert(0) += 1;

        match r.bucket {
            ReconBucket::MatchedTwoWay | ReconBucket::MatchedThreeWay => matched += 1,
            ReconBucket::AmountMismatch => amount_mismatches += 1,
            ReconBucket::TimingMismatch => timing_mismatches += 1,
            ReconBucket::Ambiguous => ambiguous += 1,
            ReconBucket::ProcessorLedgerOnly | ReconBucket::ProcessorBankOnly => left_only += 1,
            ReconBucket::LedgerOnly | ReconBucket::BankOnly => right_only += 1,
        }
    }

    let settlement = if results.iter().any(|r| r.settlement.is_some()) {
        Some(settlement::compute_settlement_summary(results))
    } else {
        None
    };

    ReconSummary {
        total_groups: results.len(),
        matched,
        amount_mismatches,
        timing_mismatches,
        ambiguous,
        left_only,
        right_only,
        bucket_counts,
        settlement,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{ClassifiedResult, Deltas, ReconBucket};

    fn result(bucket: ReconBucket) -> ClassifiedResult {
        ClassifiedResult {
            bucket,
            match_key: "k".into(),
            currency: "USD".into(),
            aggregates: HashMap::new(),
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
            settlement: None,
            proof: None,
            leg_proofs: HashMap::new(),
        }
    }

    #[test]
    fn summary_counts() {
        let results = vec![
            result(ReconBucket::MatchedTwoWay),
            result(ReconBucket::MatchedTwoWay),
            result(ReconBucket::AmountMismatch),
            result(ReconBucket::LedgerOnly),
            result(ReconBucket::ProcessorLedgerOnly),
        ];
        let summary = compute_summary(&results);
        assert_eq!(summary.total_groups, 5);
        assert_eq!(summary.matched, 2);
        assert_eq!(summary.amount_mismatches, 1);
        assert_eq!(summary.left_only, 1);
        assert_eq!(summary.right_only, 1);
    }
}
