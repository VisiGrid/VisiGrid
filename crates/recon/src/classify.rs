use std::collections::HashMap;

use crate::model::{
    ClassifiedResult, Deltas, MatchedPair, PairMatchOutput, ReconBucket,
};

/// Classify a 2-way pair match into buckets.
pub fn classify_two_way(
    pair_output: &PairMatchOutput,
    left_role: &str,
    right_role: &str,
) -> Vec<ClassifiedResult> {
    let mut results = Vec::new();

    for m in &pair_output.matched {
        let bucket = if !m.within_tolerance {
            ReconBucket::AmountMismatch
        } else if !m.within_window {
            ReconBucket::TimingMismatch
        } else {
            ReconBucket::MatchedTwoWay
        };

        let mut aggregates = HashMap::new();
        aggregates.insert(left_role.to_string(), m.left.clone());
        aggregates.insert(right_role.to_string(), m.right.clone());

        results.push(ClassifiedResult {
            bucket,
            match_key: m.left.match_key.clone(),
            currency: m.left.currency.clone(),
            aggregates,
            deltas: Deltas {
                delta_cents: Some(m.delta_cents),
                date_offset_days: Some(m.date_offset_days),
            },
        });
    }

    for agg in &pair_output.left_only {
        let mut aggregates = HashMap::new();
        aggregates.insert(left_role.to_string(), agg.clone());
        let bucket = bucket_for_only_role(left_role);
        results.push(ClassifiedResult {
            bucket,
            match_key: agg.match_key.clone(),
            currency: agg.currency.clone(),
            aggregates,
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
        });
    }

    for agg in &pair_output.right_only {
        let mut aggregates = HashMap::new();
        aggregates.insert(right_role.to_string(), agg.clone());
        let bucket = bucket_for_only_role(right_role);
        results.push(ClassifiedResult {
            bucket,
            match_key: agg.match_key.clone(),
            currency: agg.currency.clone(),
            aggregates,
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
        });
    }

    results
}

/// Merge two pair outputs into 3-way classified results.
///
/// pair_pl = processor↔ledger, pair_pb = processor↔bank.
/// Per processor aggregate key:
/// - Matched in both → MatchedThreeWay
/// - Matched in pl only → ProcessorBankOnly (bank missing)
/// - Matched in pb only → ProcessorLedgerOnly (ledger missing)
/// - AmountMismatch/TimingMismatch propagated from either pair
/// - Unmatched from ledger → LedgerOnly
/// - Unmatched from bank → BankOnly
pub fn merge_three_way(
    pair_pl: &PairMatchOutput,
    pair_pb: &PairMatchOutput,
    processor_role: &str,
    ledger_role: &str,
    bank_role: &str,
) -> Vec<ClassifiedResult> {
    let mut results = Vec::new();

    // Index matched pairs by (match_key, currency) from the processor side.
    let pl_by_key: HashMap<(String, String), &MatchedPair> = pair_pl
        .matched
        .iter()
        .map(|m| ((m.left.match_key.clone(), m.left.currency.clone()), m))
        .collect();

    let pb_by_key: HashMap<(String, String), &MatchedPair> = pair_pb
        .matched
        .iter()
        .map(|m| ((m.left.match_key.clone(), m.left.currency.clone()), m))
        .collect();

    // All processor keys from both pairs + processor-only from both
    let mut all_proc_keys: Vec<(String, String)> = Vec::new();
    for m in &pair_pl.matched {
        let k = (m.left.match_key.clone(), m.left.currency.clone());
        if !all_proc_keys.contains(&k) {
            all_proc_keys.push(k);
        }
    }
    for m in &pair_pb.matched {
        let k = (m.left.match_key.clone(), m.left.currency.clone());
        if !all_proc_keys.contains(&k) {
            all_proc_keys.push(k);
        }
    }
    for a in &pair_pl.left_only {
        let k = (a.match_key.clone(), a.currency.clone());
        if !all_proc_keys.contains(&k) {
            all_proc_keys.push(k);
        }
    }
    for a in &pair_pb.left_only {
        let k = (a.match_key.clone(), a.currency.clone());
        if !all_proc_keys.contains(&k) {
            all_proc_keys.push(k);
        }
    }

    for key in &all_proc_keys {
        let in_pl = pl_by_key.get(key);
        let in_pb = pb_by_key.get(key);

        match (in_pl, in_pb) {
            (Some(pl), Some(pb)) => {
                // Both matched — check tolerances
                let bucket = if !pl.within_tolerance || !pb.within_tolerance {
                    ReconBucket::AmountMismatch
                } else if !pl.within_window || !pb.within_window {
                    ReconBucket::TimingMismatch
                } else {
                    ReconBucket::MatchedThreeWay
                };

                let mut aggregates = HashMap::new();
                aggregates.insert(processor_role.to_string(), pl.left.clone());
                aggregates.insert(ledger_role.to_string(), pl.right.clone());
                aggregates.insert(bank_role.to_string(), pb.right.clone());

                results.push(ClassifiedResult {
                    bucket,
                    match_key: key.0.clone(),
                    currency: key.1.clone(),
                    aggregates,
                    deltas: Deltas {
                        delta_cents: Some(pl.delta_cents),
                        date_offset_days: Some(pl.date_offset_days),
                    },
                });
            }
            (Some(pl), None) => {
                // Matched processor↔ledger but bank missing
                let bucket = if !pl.within_tolerance {
                    ReconBucket::AmountMismatch
                } else if !pl.within_window {
                    ReconBucket::TimingMismatch
                } else {
                    ReconBucket::ProcessorBankOnly
                };

                let mut aggregates = HashMap::new();
                aggregates.insert(processor_role.to_string(), pl.left.clone());
                aggregates.insert(ledger_role.to_string(), pl.right.clone());

                results.push(ClassifiedResult {
                    bucket,
                    match_key: key.0.clone(),
                    currency: key.1.clone(),
                    aggregates,
                    deltas: Deltas {
                        delta_cents: Some(pl.delta_cents),
                        date_offset_days: Some(pl.date_offset_days),
                    },
                });
            }
            (None, Some(pb)) => {
                // Matched processor↔bank but ledger missing
                let bucket = if !pb.within_tolerance {
                    ReconBucket::AmountMismatch
                } else if !pb.within_window {
                    ReconBucket::TimingMismatch
                } else {
                    ReconBucket::ProcessorLedgerOnly
                };

                let mut aggregates = HashMap::new();
                aggregates.insert(processor_role.to_string(), pb.left.clone());
                aggregates.insert(bank_role.to_string(), pb.right.clone());

                results.push(ClassifiedResult {
                    bucket,
                    match_key: key.0.clone(),
                    currency: key.1.clone(),
                    aggregates,
                    deltas: Deltas {
                        delta_cents: Some(pb.delta_cents),
                        date_offset_days: Some(pb.date_offset_days),
                    },
                });
            }
            (None, None) => {
                // Processor-only (in left_only of both pairs — rare but possible)
                // Find the aggregate from either pair's left_only
                let proc_agg = pair_pl
                    .left_only
                    .iter()
                    .find(|a| a.match_key == key.0 && a.currency == key.1)
                    .or_else(|| {
                        pair_pb
                            .left_only
                            .iter()
                            .find(|a| a.match_key == key.0 && a.currency == key.1)
                    });
                if let Some(agg) = proc_agg {
                    let mut aggregates = HashMap::new();
                    aggregates.insert(processor_role.to_string(), agg.clone());
                    results.push(ClassifiedResult {
                        bucket: ReconBucket::ProcessorLedgerOnly,
                        match_key: key.0.clone(),
                        currency: key.1.clone(),
                        aggregates,
                        deltas: Deltas {
                            delta_cents: None,
                            date_offset_days: None,
                        },
                    });
                }
            }
        }
    }

    // Ledger-only (unmatched from right side of pair_pl)
    for agg in &pair_pl.right_only {
        let mut aggregates = HashMap::new();
        aggregates.insert(ledger_role.to_string(), agg.clone());
        results.push(ClassifiedResult {
            bucket: ReconBucket::LedgerOnly,
            match_key: agg.match_key.clone(),
            currency: agg.currency.clone(),
            aggregates,
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
        });
    }

    // Bank-only (unmatched from right side of pair_pb)
    for agg in &pair_pb.right_only {
        let mut aggregates = HashMap::new();
        aggregates.insert(bank_role.to_string(), agg.clone());
        results.push(ClassifiedResult {
            bucket: ReconBucket::BankOnly,
            match_key: agg.match_key.clone(),
            currency: agg.currency.clone(),
            aggregates,
            deltas: Deltas {
                delta_cents: None,
                date_offset_days: None,
            },
        });
    }

    results
}

fn bucket_for_only_role(role: &str) -> ReconBucket {
    match role {
        "processor" => ReconBucket::ProcessorLedgerOnly,
        "ledger" => ReconBucket::LedgerOnly,
        "bank" => ReconBucket::BankOnly,
        _ => {
            // For generic roles, use the role name heuristic
            if role.contains("ledger") {
                ReconBucket::LedgerOnly
            } else if role.contains("bank") {
                ReconBucket::BankOnly
            } else {
                // Default: treat left as processor-side, right as ledger-side
                ReconBucket::ProcessorLedgerOnly
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Aggregate;
    use chrono::NaiveDate;

    fn agg(role: &str, key: &str, cur: &str, cents: i64, date: &str) -> Aggregate {
        Aggregate {
            role: role.into(),
            match_key: key.into(),
            currency: cur.into(),
            date: NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
            total_cents: cents,
            record_count: 1,
            record_ids: vec![format!("{key}_1")],
        }
    }

    fn mp(
        left: Aggregate,
        right: Aggregate,
        delta: i64,
        date_off: i32,
        tol: bool,
        win: bool,
    ) -> MatchedPair {
        MatchedPair {
            left,
            right,
            delta_cents: delta,
            date_offset_days: date_off,
            within_tolerance: tol,
            within_window: win,
        }
    }

    #[test]
    fn two_way_all_matched() {
        let pair = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-17"),
                agg("ledger", "po_1", "USD", 7210, "2026-01-17"),
                0, 0, true, true,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let classified = classify_two_way(&pair, "processor", "ledger");
        assert_eq!(classified.len(), 1);
        assert_eq!(classified[0].bucket, ReconBucket::MatchedTwoWay);
    }

    #[test]
    fn two_way_amount_mismatch() {
        let pair = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-17"),
                agg("ledger", "po_1", "USD", 7200, "2026-01-17"),
                10, 0, false, true,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let classified = classify_two_way(&pair, "processor", "ledger");
        assert_eq!(classified[0].bucket, ReconBucket::AmountMismatch);
    }

    #[test]
    fn two_way_timing_mismatch() {
        let pair = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-15"),
                agg("ledger", "po_1", "USD", 7210, "2026-01-20"),
                0, -5, true, false,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let classified = classify_two_way(&pair, "processor", "ledger");
        assert_eq!(classified[0].bucket, ReconBucket::TimingMismatch);
    }

    #[test]
    fn two_way_only_sides() {
        let pair = PairMatchOutput {
            matched: vec![],
            left_only: vec![agg("proc", "po_1", "USD", 7210, "2026-01-17")],
            right_only: vec![agg("ledger", "dep_9", "USD", 3000, "2026-01-19")],
        };
        let classified = classify_two_way(&pair, "processor", "ledger");
        assert_eq!(classified.len(), 2);
        assert_eq!(classified[0].bucket, ReconBucket::ProcessorLedgerOnly);
        assert_eq!(classified[1].bucket, ReconBucket::LedgerOnly);
    }

    #[test]
    fn three_way_all_matched() {
        let pair_pl = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-17"),
                agg("ledger", "po_1", "USD", 7210, "2026-01-18"),
                0, -1, true, true,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let pair_pb = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-17"),
                agg("bank", "po_1", "USD", 7210, "2026-01-17"),
                0, 0, true, true,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let classified = merge_three_way(&pair_pl, &pair_pb, "processor", "ledger", "bank");
        assert_eq!(classified.len(), 1);
        assert_eq!(classified[0].bucket, ReconBucket::MatchedThreeWay);
        assert_eq!(classified[0].aggregates.len(), 3);
    }

    #[test]
    fn three_way_bank_missing() {
        let pair_pl = PairMatchOutput {
            matched: vec![mp(
                agg("proc", "po_1", "USD", 7210, "2026-01-17"),
                agg("ledger", "po_1", "USD", 7210, "2026-01-18"),
                0, -1, true, true,
            )],
            left_only: vec![],
            right_only: vec![],
        };
        let pair_pb = PairMatchOutput {
            matched: vec![],
            left_only: vec![agg("proc", "po_1", "USD", 7210, "2026-01-17")],
            right_only: vec![],
        };
        let classified = merge_three_way(&pair_pl, &pair_pb, "processor", "ledger", "bank");
        assert_eq!(classified.len(), 1);
        assert_eq!(classified[0].bucket, ReconBucket::ProcessorBankOnly);
    }

    #[test]
    fn three_way_ledger_only() {
        let pair_pl = PairMatchOutput {
            matched: vec![],
            left_only: vec![],
            right_only: vec![agg("ledger", "dep_9", "USD", 3000, "2026-01-19")],
        };
        let pair_pb = PairMatchOutput {
            matched: vec![],
            left_only: vec![],
            right_only: vec![],
        };
        let classified = merge_three_way(&pair_pl, &pair_pb, "processor", "ledger", "bank");
        assert_eq!(classified.len(), 1);
        assert_eq!(classified[0].bucket, ReconBucket::LedgerOnly);
    }
}
