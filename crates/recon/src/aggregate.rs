use std::collections::BTreeMap;

use chrono::NaiveDate;

use crate::model::{Aggregate, AggregateKey, ReconRow};

/// Group records by (match_key, currency), sum amounts, track earliest date.
pub fn aggregate_records(role: &str, rows: &[ReconRow]) -> Vec<Aggregate> {
    let mut groups: BTreeMap<AggregateKey, (i64, NaiveDate, usize, Vec<String>)> = BTreeMap::new();

    for row in rows {
        let key = AggregateKey {
            match_key: row.match_key.clone(),
            currency: row.currency.clone(),
        };
        let entry = groups.entry(key).or_insert_with(|| (0, row.date, 0, Vec::new()));
        entry.0 += row.amount_cents;
        if row.date < entry.1 {
            entry.1 = row.date;
        }
        entry.2 += 1;
        entry.3.push(row.record_id.clone());
    }

    groups
        .into_iter()
        .map(|(key, (total_cents, date, count, ids))| Aggregate {
            role: role.to_string(),
            match_key: key.match_key,
            currency: key.currency,
            date,
            total_cents,
            record_count: count,
            record_ids: ids,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;
    use std::collections::HashMap;

    fn row(match_key: &str, amount: i64, date: &str, currency: &str) -> ReconRow {
        ReconRow {
            role: "test".into(),
            record_id: format!("r_{match_key}_{amount}"),
            match_key: match_key.into(),
            date: NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
            amount_cents: amount,
            currency: currency.into(),
            kind: "payment".into(),
            raw_fields: HashMap::new(),
        }
    }

    #[test]
    fn basic_aggregation() {
        let rows = vec![
            row("po_1", 1000, "2026-01-15", "USD"),
            row("po_1", -290, "2026-01-15", "USD"),
            row("po_1", -2500, "2026-01-16", "USD"),
        ];
        let aggs = aggregate_records("processor", &rows);
        assert_eq!(aggs.len(), 1);
        assert_eq!(aggs[0].total_cents, 1000 - 290 - 2500);
        assert_eq!(aggs[0].record_count, 3);
        assert_eq!(aggs[0].date, NaiveDate::from_ymd_opt(2026, 1, 15).unwrap());
    }

    #[test]
    fn currency_separation() {
        let rows = vec![
            row("po_1", 1000, "2026-01-15", "USD"),
            row("po_1", 500, "2026-01-15", "CAD"),
        ];
        let aggs = aggregate_records("processor", &rows);
        assert_eq!(aggs.len(), 2);
        // BTreeMap ordering: CAD before USD
        assert_eq!(aggs[0].currency, "CAD");
        assert_eq!(aggs[0].total_cents, 500);
        assert_eq!(aggs[1].currency, "USD");
        assert_eq!(aggs[1].total_cents, 1000);
    }

    #[test]
    fn earliest_date_used() {
        let rows = vec![
            row("po_1", 100, "2026-01-17", "USD"),
            row("po_1", 200, "2026-01-15", "USD"),
            row("po_1", 300, "2026-01-16", "USD"),
        ];
        let aggs = aggregate_records("processor", &rows);
        assert_eq!(aggs[0].date, NaiveDate::from_ymd_opt(2026, 1, 15).unwrap());
    }
}
