use std::collections::HashMap;

use crate::aggregate::aggregate_records;
use crate::classify::{classify_two_way, merge_three_way};
use crate::config::{MatchStrategy, ReconConfig};
use crate::error::ReconError;
use crate::evidence::compute_summary;
use crate::matcher::{match_exact_key, match_fuzzy_amount_date};
use crate::model::{Aggregate, ClassifiedResult, DerivedOutputs, ReconInput, ReconMeta, ReconResult, ReconRow};

/// Run reconciliation per config. Returns classified results + summary.
pub fn run(config: &ReconConfig, input: &ReconInput) -> Result<ReconResult, ReconError> {
    // Aggregate all roles
    let mut aggregates: HashMap<String, Vec<Aggregate>> = HashMap::new();
    for (role_name, rows) in &input.records {
        aggregates.insert(role_name.clone(), aggregate_records(role_name, rows));
    }

    let mut classified = if config.way == 2 {
        run_two_way(config, &aggregates, input)?
    } else {
        run_three_way(config, &aggregates, input)?
    };

    if let Some(ref settlement_config) = config.settlement {
        crate::settlement::classify_settlement(&mut classified, settlement_config);
    }

    let summary = compute_summary(&classified);

    let mut derived = DerivedOutputs::default();
    derived.payout_rollup = crate::derived::build_payout_rollup(&classified, config);

    Ok(ReconResult {
        meta: ReconMeta {
            config_name: config.name.clone(),
            way: config.way,
            engine_version: env!("CARGO_PKG_VERSION").to_string(),
            run_at: chrono::Utc::now().to_rfc3339(),
            settlement_clock: config.settlement.as_ref().map(|s| s.clock),
        },
        summary,
        groups: classified,
        derived,
    })
}

fn run_two_way(
    config: &ReconConfig,
    aggregates: &HashMap<String, Vec<Aggregate>>,
    input: &ReconInput,
) -> Result<Vec<ClassifiedResult>, ReconError> {
    let (pair_name, pair) = config.pairs.iter().next().unwrap();
    let left_aggs = aggregates.get(&pair.left).ok_or_else(|| {
        ReconError::UnknownRole(format!("pair '{pair_name}': left role '{}' has no data", pair.left))
    })?;
    let right_aggs = aggregates.get(&pair.right).ok_or_else(|| {
        ReconError::UnknownRole(format!(
            "pair '{pair_name}': right role '{}' has no data",
            pair.right
        ))
    })?;

    let pair_output = match pair.strategy {
        MatchStrategy::ExactKey => match_exact_key(left_aggs, right_aggs, &config.tolerance),
        MatchStrategy::FuzzyAmountDate => {
            match_fuzzy_amount_date(left_aggs, right_aggs, &config.tolerance)
        }
        MatchStrategy::WindowedNm => {
            let wnm_config = pair.windowed_nm.clone().unwrap_or_default();
            let left_rows = input.records.get(&pair.left).map(|v| v.as_slice()).unwrap_or(&[]);
            let right_rows = input.records.get(&pair.right).map(|v| v.as_slice()).unwrap_or(&[]);
            crate::windowed_nm::match_windowed_nm(left_rows, right_rows, &config.tolerance, &wnm_config)
        }
    };

    Ok(classify_two_way(&pair_output, &pair.left, &pair.right))
}

fn run_three_way(
    config: &ReconConfig,
    aggregates: &HashMap<String, Vec<Aggregate>>,
    input: &ReconInput,
) -> Result<Vec<ClassifiedResult>, ReconError> {
    // Find the two pairs. Convention: first pair = processor↔ledger, second = processor↔bank
    let pairs: Vec<_> = config.pairs.iter().collect();
    if pairs.len() != 2 {
        return Err(ReconError::WayMismatch {
            way: 3,
            pairs: pairs.len(),
        });
    }

    let (_, pair_0) = pairs[0];
    let (_, pair_1) = pairs[1];

    // Determine which is processor↔ledger and which is processor↔bank
    // by finding the common role (processor) across both pairs
    let (pl_pair, pb_pair, processor_role, ledger_role, bank_role) =
        identify_three_way_roles(pair_0, pair_1, config)?;

    let proc_aggs = aggregates.get(&processor_role).ok_or_else(|| {
        ReconError::UnknownRole(format!("processor role '{processor_role}' has no data"))
    })?;
    let ledger_aggs = aggregates.get(&ledger_role).ok_or_else(|| {
        ReconError::UnknownRole(format!("ledger role '{ledger_role}' has no data"))
    })?;
    let bank_aggs = aggregates.get(&bank_role).ok_or_else(|| {
        ReconError::UnknownRole(format!("bank role '{bank_role}' has no data"))
    })?;

    let match_fn = |left: &[Aggregate], right: &[Aggregate], pair: &crate::config::PairConfig| {
        match pair.strategy {
            MatchStrategy::ExactKey => match_exact_key(left, right, &config.tolerance),
            MatchStrategy::FuzzyAmountDate => {
                match_fuzzy_amount_date(left, right, &config.tolerance)
            }
            MatchStrategy::WindowedNm => {
                let wnm_config = pair.windowed_nm.clone().unwrap_or_default();
                let left_rows = input.records.get(&pair.left).map(|v| v.as_slice()).unwrap_or(&[]);
                let right_rows = input.records.get(&pair.right).map(|v| v.as_slice()).unwrap_or(&[]);
                crate::windowed_nm::match_windowed_nm(left_rows, right_rows, &config.tolerance, &wnm_config)
            }
        }
    };

    let pair_pl = match_fn(proc_aggs, ledger_aggs, pl_pair);
    let pair_pb = match_fn(proc_aggs, bank_aggs, pb_pair);

    Ok(merge_three_way(
        &pair_pl,
        &pair_pb,
        &processor_role,
        &ledger_role,
        &bank_role,
    ))
}

/// Identify the shared (processor) role and the two unique roles across a 3-way pair.
fn identify_three_way_roles<'a>(
    pair_0: &'a crate::config::PairConfig,
    pair_1: &'a crate::config::PairConfig,
    config: &'a ReconConfig,
) -> Result<
    (
        &'a crate::config::PairConfig,
        &'a crate::config::PairConfig,
        String,
        String,
        String,
    ),
    ReconError,
> {
    // Collect all role refs
    let roles_0 = [&pair_0.left, &pair_0.right];
    let roles_1 = [&pair_1.left, &pair_1.right];

    // Find the shared role (appears in both pairs)
    let shared: Vec<&String> = roles_0
        .iter()
        .filter(|r| roles_1.contains(r))
        .copied()
        .collect();

    if shared.len() != 1 {
        return Err(ReconError::ConfigValidation(
            "3-way recon requires exactly one shared role across both pairs".into(),
        ));
    }

    let processor_role = shared[0].clone();

    // The other role in pair_0 and pair_1
    let other_0 = if pair_0.left == processor_role {
        &pair_0.right
    } else {
        &pair_0.left
    };
    let other_1 = if pair_1.left == processor_role {
        &pair_1.right
    } else {
        &pair_1.left
    };

    // Determine which is "ledger" and which is "bank" by role kind
    let kind_0 = config
        .roles
        .get(other_0)
        .map(|r| r.kind)
        .ok_or_else(|| ReconError::UnknownRole(other_0.clone()))?;
    let kind_1 = config
        .roles
        .get(other_1)
        .map(|r| r.kind)
        .ok_or_else(|| ReconError::UnknownRole(other_1.clone()))?;

    use crate::config::RoleKind;

    // Assign: ledger-kind role pairs with processor as pair_pl, bank-kind as pair_pb
    // If both are the same kind, use order as-is (pair_0=pl, pair_1=pb)
    let (pl_pair, pb_pair, ledger_role, bank_role) = match (kind_0, kind_1) {
        (RoleKind::Ledger, RoleKind::Bank) => (pair_0, pair_1, other_0.clone(), other_1.clone()),
        (RoleKind::Bank, RoleKind::Ledger) => (pair_1, pair_0, other_1.clone(), other_0.clone()),
        _ => {
            // Default: first pair = processor↔ledger, second = processor↔bank
            (pair_0, pair_1, other_0.clone(), other_1.clone())
        }
    };

    Ok((pl_pair, pb_pair, processor_role, ledger_role, bank_role))
}

/// Load CSV rows into ReconRows, applying column mapping, filter, and transform.
pub fn load_csv_rows(
    role_name: &str,
    csv_data: &str,
    role_config: &crate::config::RoleConfig,
) -> Result<Vec<ReconRow>, ReconError> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(true)
        .from_reader(csv_data.as_bytes());

    let headers: Vec<String> = reader
        .headers()
        .map_err(|e| ReconError::Io(e.to_string()))?
        .iter()
        .map(|h| h.to_string())
        .collect();

    let col = &role_config.columns;

    let idx = |name: &str| -> Result<usize, ReconError> {
        headers.iter().position(|h| h == name).ok_or_else(|| {
            ReconError::MissingColumn {
                role: role_name.into(),
                column: name.into(),
            }
        })
    };

    let record_id_idx = idx(&col.record_id)?;
    let match_key_idx = idx(&col.match_key)?;
    let amount_idx = idx(&col.amount)?;
    let date_idx = idx(&col.date)?;
    let currency_idx = idx(&col.currency)?;
    let kind_idx = idx(&col.kind)?;

    // Filter column index (if configured)
    let filter_idx = if let Some(ref filter) = role_config.filter {
        Some(idx(&filter.column)?)
    } else {
        None
    };

    // Transform when_column index (if configured)
    let transform_when_idx = if let Some(ref xf) = role_config.transform {
        if let Some(ref when_col) = xf.when_column {
            Some(idx(when_col)?)
        } else {
            None
        }
    } else {
        None
    };

    let mut rows = Vec::new();

    for record in reader.records() {
        let record = record.map_err(|e| ReconError::Io(e.to_string()))?;

        // Apply filter
        if let (Some(ref filter), Some(fi)) = (&role_config.filter, filter_idx) {
            let val = record.get(fi).unwrap_or("");
            if !filter.values.iter().any(|v| v == val) {
                continue;
            }
        }

        let record_id = record.get(record_id_idx).unwrap_or("").to_string();
        let match_key = record.get(match_key_idx).unwrap_or("").to_string();
        let currency = record.get(currency_idx).unwrap_or("").to_string();
        let kind = record.get(kind_idx).unwrap_or("").to_string();

        let date_str = record.get(date_idx).unwrap_or("");
        let date = chrono::NaiveDate::parse_from_str(date_str, "%Y-%m-%d").map_err(|_| {
            ReconError::DateParse {
                role: role_name.into(),
                record_id: record_id.clone(),
                value: date_str.into(),
            }
        })?;

        let amount_str = record.get(amount_idx).unwrap_or("");
        let mut amount_cents: i64 = amount_str.parse().map_err(|_| ReconError::AmountParse {
            role: role_name.into(),
            record_id: record_id.clone(),
            value: amount_str.into(),
        })?;

        // Apply transform
        if let Some(ref xf) = role_config.transform {
            let should_apply = match (&xf.when_column, &xf.when_values, transform_when_idx) {
                (Some(_), Some(values), Some(wi)) => {
                    let val = record.get(wi).unwrap_or("");
                    values.iter().any(|v| v == val)
                }
                (None, _, _) | (_, None, _) => true,
                _ => true,
            };

            if should_apply {
                if let Some(mult) = xf.multiply {
                    amount_cents *= mult;
                }
            }
        }

        // Build raw_fields
        let mut raw_fields = HashMap::new();
        for (i, h) in headers.iter().enumerate() {
            if let Some(val) = record.get(i) {
                raw_fields.insert(h.clone(), val.to_string());
            }
        }

        rows.push(ReconRow {
            role: role_name.into(),
            record_id,
            match_key,
            date,
            amount_cents,
            currency,
            kind,
            raw_fields,
        });
    }

    Ok(rows)
}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn load_csv_basic() {
        let csv = "\
source_id,group_id,amount_minor,effective_date,currency,type
txn_1,po_1,10000,2026-01-15,USD,charge
txn_2,po_1,-290,2026-01-15,USD,fee
txn_3,po_1,-7210,2026-01-17,USD,payout
";
        let role_config = crate::config::RoleConfig {
            kind: crate::config::RoleKind::Processor,
            file: "stripe.csv".into(),
            columns: crate::config::ColumnMapping {
                record_id: "source_id".into(),
                match_key: "group_id".into(),
                amount: "amount_minor".into(),
                date: "effective_date".into(),
                currency: "currency".into(),
                kind: "type".into(),
            },
            filter: None,
            transform: None,
        };

        let rows = load_csv_rows("processor", csv, &role_config).unwrap();
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0].record_id, "txn_1");
        assert_eq!(rows[0].amount_cents, 10000);
        assert_eq!(rows[2].amount_cents, -7210);
    }

    #[test]
    fn load_csv_with_filter() {
        let csv = "\
source_id,group_id,amount_minor,effective_date,currency,type
txn_1,po_1,10000,2026-01-15,USD,charge
txn_2,po_1,-290,2026-01-15,USD,fee
txn_3,po_1,-7210,2026-01-17,USD,payout
";
        let role_config = crate::config::RoleConfig {
            kind: crate::config::RoleKind::Processor,
            file: "stripe.csv".into(),
            columns: crate::config::ColumnMapping {
                record_id: "source_id".into(),
                match_key: "group_id".into(),
                amount: "amount_minor".into(),
                date: "effective_date".into(),
                currency: "currency".into(),
                kind: "type".into(),
            },
            filter: Some(crate::config::RowFilter {
                column: "type".into(),
                values: vec!["payout".into()],
            }),
            transform: None,
        };

        let rows = load_csv_rows("processor", csv, &role_config).unwrap();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].kind, "payout");
    }

    #[test]
    fn load_csv_with_transform() {
        let csv = "\
source_id,group_id,amount_minor,effective_date,currency,type
txn_1,po_1,-7210,2026-01-17,USD,payout
txn_2,po_1,10000,2026-01-15,USD,charge
";
        let role_config = crate::config::RoleConfig {
            kind: crate::config::RoleKind::Processor,
            file: "stripe.csv".into(),
            columns: crate::config::ColumnMapping {
                record_id: "source_id".into(),
                match_key: "group_id".into(),
                amount: "amount_minor".into(),
                date: "effective_date".into(),
                currency: "currency".into(),
                kind: "type".into(),
            },
            filter: None,
            transform: Some(crate::config::AmountTransform {
                multiply: Some(-1),
                when_column: Some("type".into()),
                when_values: Some(vec!["payout".into()]),
            }),
        };

        let rows = load_csv_rows("processor", csv, &role_config).unwrap();
        assert_eq!(rows.len(), 2);
        // payout: -7210 * -1 = 7210
        assert_eq!(rows[0].amount_cents, 7210);
        // charge: not transformed
        assert_eq!(rows[1].amount_cents, 10000);
    }

    #[test]
    fn integration_two_way() {
        let stripe_csv = "\
source_id,group_id,amount_minor,effective_date,currency,type
ch_1,po_1,10000,2026-01-15,USD,charge
fee_1,po_1,-290,2026-01-15,USD,fee
ref_1,po_1,-2500,2026-01-16,USD,refund
po_1,po_1,-7210,2026-01-17,USD,payout
ch_2,po_2,5000,2026-01-16,USD,charge
fee_2,po_2,-145,2026-01-16,USD,fee
po_2,po_2,-4855,2026-01-18,USD,payout
";
        let qbo_csv = "\
source_id,group_id,amount_minor,effective_date,currency,type
dep_1,dep_1,7210,2026-01-18,USD,deposit
dep_2,dep_2,4855,2026-01-19,USD,deposit
";
        let config_toml = r#"
name = "Integration Test"
way = 2

[roles.processor]
kind = "processor"
file = "stripe.csv"
[roles.processor.columns]
record_id  = "source_id"
match_key  = "group_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"
[roles.processor.filter]
column = "type"
values = ["payout"]
[roles.processor.transform]
multiply = -1
when_column = "type"
when_values = ["payout"]

[roles.ledger]
kind = "ledger"
file = "qbo.csv"
[roles.ledger.columns]
record_id  = "source_id"
match_key  = "source_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"

[pairs.processor_ledger]
left = "processor"
right = "ledger"
strategy = "fuzzy_amount_date"

[tolerance]
amount_cents = 0
date_window_days = 2
"#;
        let config = crate::config::ReconConfig::from_toml(config_toml).unwrap();
        let proc_rows = load_csv_rows("processor", stripe_csv, &config.roles["processor"]).unwrap();
        let ledger_rows = load_csv_rows("ledger", qbo_csv, &config.roles["ledger"]).unwrap();

        let input = ReconInput {
            records: HashMap::from([
                ("processor".into(), proc_rows),
                ("ledger".into(), ledger_rows),
            ]),
        };

        let result = run(&config, &input).unwrap();
        assert_eq!(result.meta.way, 2);
        assert_eq!(result.summary.matched, 2);
        assert_eq!(result.summary.amount_mismatches, 0);
    }

    #[test]
    fn integration_two_way_windowed_nm() {
        // Settlement scenario: 3 processor payments sum to 1 bank deposit,
        // and 2 processor payments sum to another bank deposit.
        let settlements_csv = "\
settlement_id,net_minor,payout_date,currency,type
s1,3000,2026-01-15,USD,payout
s2,4000,2026-01-15,USD,payout
s3,3000,2026-01-16,USD,payout
s4,2500,2026-01-17,USD,payout
s5,2500,2026-01-17,USD,payout
";
        let deposits_csv = "\
txn_id,amount_minor,posted_date,currency,type
d1,10000,2026-01-16,USD,deposit
d2,5000,2026-01-18,USD,deposit
";
        let config_toml = r#"
name = "Settlement Recon"
way = 2

[roles.processor]
kind = "processor"
file = "settlements.csv"
[roles.processor.columns]
record_id  = "settlement_id"
match_key  = "settlement_id"
amount     = "net_minor"
date       = "payout_date"
currency   = "currency"
kind       = "type"

[roles.bank]
kind = "bank"
file = "deposits.csv"
[roles.bank.columns]
record_id  = "txn_id"
match_key  = "txn_id"
amount     = "amount_minor"
date       = "posted_date"
currency   = "currency"
kind       = "type"

[pairs.settlement_bank]
left = "processor"
right = "bank"
strategy = "windowed_nm"

[pairs.settlement_bank.windowed_nm]
max_group_size = 6
max_bucket_size = 50
max_nodes = 50000

[tolerance]
amount_cents = 0
date_window_days = 3
"#;
        let config = crate::config::ReconConfig::from_toml(config_toml).unwrap();
        let proc_rows =
            load_csv_rows("processor", settlements_csv, &config.roles["processor"]).unwrap();
        let bank_rows =
            load_csv_rows("bank", deposits_csv, &config.roles["bank"]).unwrap();

        assert_eq!(proc_rows.len(), 5);
        assert_eq!(bank_rows.len(), 2);

        let input = ReconInput {
            records: HashMap::from([
                ("processor".into(), proc_rows),
                ("bank".into(), bank_rows),
            ]),
        };

        let result = run(&config, &input).unwrap();
        assert_eq!(result.meta.way, 2);
        // s1(3000)+s2(4000)+s3(3000)=10000=d1, s4(2500)+s5(2500)=5000=d2
        assert_eq!(result.summary.matched, 2, "expected 2 matched groups");
        assert_eq!(result.summary.amount_mismatches, 0);
        assert_eq!(result.summary.left_only, 0);
        assert_eq!(result.summary.right_only, 0);

        // Verify proofs exist
        for group in &result.groups {
            if group.bucket == crate::model::ReconBucket::MatchedTwoWay {
                assert!(group.proof.is_some(), "matched groups should have proofs");
                let proof = group.proof.as_ref().unwrap();
                assert_eq!(proof.strategy, "windowed_nm");
            }
        }
    }
}
