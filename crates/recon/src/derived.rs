//! Derived dataset builders — computed analyses layered on top of classified results.

use serde_json::json;

use crate::config::{ReconConfig, RoleKind};
use crate::model::{ClassifiedResult, DerivedDataset};

/// Build the `payout_rollup.v1` derived dataset.
///
/// Only produces rows when the config has at least one role with
/// `kind = processor` or `kind = bank` (settlement-aware configs).
/// Returns an empty dataset otherwise.
pub fn build_payout_rollup(groups: &[ClassifiedResult], config: &ReconConfig) -> DerivedDataset {
    let has_settlement = config.roles.values().any(|r| {
        matches!(r.kind, RoleKind::Processor | RoleKind::Bank)
    });

    let mut dataset = DerivedDataset::new("payout_rollup");
    if !has_settlement {
        return dataset;
    }

    for group in groups {
        // Use the processor aggregate if available, otherwise first available
        let primary_agg = group
            .aggregates
            .get("processor")
            .or_else(|| group.aggregates.values().next());

        let Some(agg) = primary_agg else {
            continue;
        };

        let mut row = json!({
            "payout_id": group.match_key,
            "currency": group.currency.to_lowercase(),
            "bucket": group.bucket.to_string(),
            "status": settlement_status(group),
            "total_cents": agg.total_cents,
            "delta_cents": group.deltas.delta_cents.unwrap_or(0),
            "record_count": agg.record_count,
            "earliest_date": agg.date.to_string(),
        });

        // Optional settlement fields — only present when settlement classification exists
        if let Some(ref sc) = group.settlement {
            if let Some(age) = sc.age_days {
                row["settlement_age_days"] = json!(age);
            }
            row["settlement_sla_days"] = json!(sc.sla_days);
        }

        dataset.rows.push(row);
    }

    dataset.enforce_limit();
    dataset
}

fn settlement_status(group: &ClassifiedResult) -> &'static str {
    if let Some(ref sc) = group.settlement {
        match sc.state {
            crate::model::SettlementState::Matched => "matched",
            crate::model::SettlementState::Pending => "pending",
            crate::model::SettlementState::Stale => "stale",
            crate::model::SettlementState::Error => "error",
        }
    } else {
        // Without settlement config, derive from bucket
        if group.bucket.to_string().starts_with("matched") {
            "matched"
        } else {
            "unclassified"
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use crate::model::*;
    use crate::config::*;
    use chrono::NaiveDate;

    fn test_config(kind: RoleKind) -> ReconConfig {
        ReconConfig {
            name: "test".into(),
            way: 2,
            roles: HashMap::from([
                ("processor".into(), RoleConfig {
                    kind,
                    file: "test.csv".into(),
                    columns: ColumnMapping {
                        record_id: "id".into(),
                        match_key: "key".into(),
                        amount: "amt".into(),
                        date: "dt".into(),
                        currency: "cur".into(),
                        kind: "k".into(),
                    },
                    filter: None,
                    transform: None,
                }),
                ("ledger".into(), RoleConfig {
                    kind: RoleKind::Ledger,
                    file: "ledger.csv".into(),
                    columns: ColumnMapping {
                        record_id: "id".into(),
                        match_key: "key".into(),
                        amount: "amt".into(),
                        date: "dt".into(),
                        currency: "cur".into(),
                        kind: "k".into(),
                    },
                    filter: None,
                    transform: None,
                }),
            ]),
            pairs: HashMap::from([
                ("pl".into(), PairConfig {
                    left: "processor".into(),
                    right: "ledger".into(),
                    strategy: MatchStrategy::ExactKey,
                    windowed_nm: None,
                }),
            ]),
            tolerance: ToleranceConfig::default(),
            output: OutputConfig::default(),
            settlement: None,
            fail_on_ambiguous: false,
        }
    }

    fn make_group(match_key: &str, bucket: ReconBucket) -> ClassifiedResult {
        ClassifiedResult {
            bucket,
            match_key: match_key.into(),
            currency: "USD".into(),
            aggregates: HashMap::from([
                ("processor".into(), Aggregate {
                    role: "processor".into(),
                    match_key: match_key.into(),
                    currency: "USD".into(),
                    date: NaiveDate::from_ymd_opt(2026, 1, 15).unwrap(),
                    total_cents: 10000,
                    record_count: 3,
                    record_ids: vec!["a".into(), "b".into(), "c".into()],
                }),
            ]),
            deltas: Deltas { delta_cents: Some(0), date_offset_days: None },
            settlement: None,
            proof: None,
            leg_proofs: HashMap::new(),
        }
    }

    #[test]
    fn empty_when_no_settlement_roles() {
        let config = test_config(RoleKind::Ledger);
        let groups = vec![make_group("po_1", ReconBucket::MatchedTwoWay)];
        let ds = build_payout_rollup(&groups, &config);
        assert!(ds.is_empty());
    }

    #[test]
    fn builds_rows_for_processor_config() {
        let config = test_config(RoleKind::Processor);
        let groups = vec![
            make_group("po_1", ReconBucket::MatchedTwoWay),
            make_group("po_2", ReconBucket::LedgerOnly),
        ];
        let ds = build_payout_rollup(&groups, &config);
        assert_eq!(ds.rows.len(), 2);
        assert_eq!(ds.schema, "payout_rollup");
        assert_eq!(ds.version, 1);

        let row = &ds.rows[0];
        assert_eq!(row["payout_id"], "po_1");
        assert_eq!(row["currency"], "usd");
        assert_eq!(row["bucket"], "matched_two_way");
        assert_eq!(row["total_cents"], 10000);
        assert_eq!(row["record_count"], 3);
    }

    #[test]
    fn builds_rows_for_bank_config() {
        let config = test_config(RoleKind::Bank);
        let groups = vec![make_group("po_1", ReconBucket::MatchedTwoWay)];
        let ds = build_payout_rollup(&groups, &config);
        assert_eq!(ds.rows.len(), 1);
    }

    #[test]
    fn truncates_at_max_rows() {
        use crate::model::DerivedDataset;
        let config = test_config(RoleKind::Processor);
        let groups: Vec<ClassifiedResult> = (0..DerivedDataset::MAX_ROWS + 500)
            .map(|i| make_group(&format!("po_{i}"), ReconBucket::MatchedTwoWay))
            .collect();
        let ds = build_payout_rollup(&groups, &config);
        assert_eq!(ds.rows.len(), DerivedDataset::MAX_ROWS);
        assert!(ds.truncated);
    }

    #[test]
    fn not_truncated_under_limit() {
        let config = test_config(RoleKind::Processor);
        let groups = vec![make_group("po_1", ReconBucket::MatchedTwoWay)];
        let ds = build_payout_rollup(&groups, &config);
        assert!(!ds.truncated);
    }

    #[test]
    fn includes_settlement_fields_when_present() {
        let config = test_config(RoleKind::Processor);
        let mut group = make_group("po_1", ReconBucket::MatchedTwoWay);
        group.settlement = Some(SettlementClassification {
            state: SettlementState::Pending,
            age_days: Some(2),
            sla_days: 5,
        });
        let ds = build_payout_rollup(&[group], &config);
        let row = &ds.rows[0];
        assert_eq!(row["settlement_age_days"], 2);
        assert_eq!(row["settlement_sla_days"], 5);
        assert_eq!(row["status"], "pending");
    }
}
