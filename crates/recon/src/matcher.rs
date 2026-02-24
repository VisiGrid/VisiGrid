use std::collections::BTreeMap;

use crate::config::ToleranceConfig;
use crate::model::{Aggregate, AggregateKey, MatchedPair, PairMatchOutput};

/// Match two sets of aggregates by exact (match_key, currency).
pub fn match_exact_key(
    left: &[Aggregate],
    right: &[Aggregate],
    tolerance: &ToleranceConfig,
) -> PairMatchOutput {
    let left_map: BTreeMap<AggregateKey, &Aggregate> = left
        .iter()
        .map(|a| {
            (
                AggregateKey {
                    match_key: a.match_key.clone(),
                    currency: a.currency.clone(),
                },
                a,
            )
        })
        .collect();

    let right_map: BTreeMap<AggregateKey, &Aggregate> = right
        .iter()
        .map(|a| {
            (
                AggregateKey {
                    match_key: a.match_key.clone(),
                    currency: a.currency.clone(),
                },
                a,
            )
        })
        .collect();

    let mut matched = Vec::new();
    let mut left_only = Vec::new();
    let mut right_only = Vec::new();

    for (key, left_agg) in &left_map {
        if let Some(right_agg) = right_map.get(key) {
            let delta_cents = left_agg.total_cents - right_agg.total_cents;
            let date_offset_days = (left_agg.date - right_agg.date).num_days() as i32;
            let within_tolerance = delta_cents.abs() <= tolerance.amount_cents;
            let within_window =
                (date_offset_days.unsigned_abs()) <= tolerance.date_window_days;

            matched.push(MatchedPair {
                left: (*left_agg).clone(),
                right: (*right_agg).clone(),
                delta_cents,
                date_offset_days,
                within_tolerance,
                within_window,
                proof: None,
            });
        } else {
            left_only.push((*left_agg).clone());
        }
    }

    for (key, right_agg) in &right_map {
        if !left_map.contains_key(key) {
            right_only.push((*right_agg).clone());
        }
    }

    PairMatchOutput {
        matched,
        left_only,
        right_only,
    }
}

/// Match by fuzzy amount + date window (no key required).
/// Finds best match for each left aggregate from unmatched right aggregates.
pub fn match_fuzzy_amount_date(
    left: &[Aggregate],
    right: &[Aggregate],
    tolerance: &ToleranceConfig,
) -> PairMatchOutput {
    let mut right_used = vec![false; right.len()];
    let mut matched = Vec::new();
    let mut left_only = Vec::new();

    for left_agg in left {
        let mut best: Option<(usize, i64, i32)> = None;

        for (ri, right_agg) in right.iter().enumerate() {
            if right_used[ri] {
                continue;
            }
            // Must be same currency
            if left_agg.currency != right_agg.currency {
                continue;
            }

            let delta = left_agg.total_cents - right_agg.total_cents;
            let date_off = (left_agg.date - right_agg.date).num_days() as i32;

            if delta.abs() <= tolerance.amount_cents
                && date_off.unsigned_abs() <= tolerance.date_window_days
            {
                let score = delta.abs() + (date_off.abs() as i64);
                if best.is_none() || score < best.unwrap().1 + best.unwrap().2.abs() as i64 {
                    best = Some((ri, delta, date_off));
                }
            }
        }

        if let Some((ri, delta_cents, date_offset_days)) = best {
            right_used[ri] = true;
            matched.push(MatchedPair {
                left: left_agg.clone(),
                right: right[ri].clone(),
                delta_cents,
                date_offset_days,
                within_tolerance: delta_cents.abs() <= tolerance.amount_cents,
                within_window: date_offset_days.unsigned_abs() <= tolerance.date_window_days,
                proof: None,
            });
        } else {
            left_only.push(left_agg.clone());
        }
    }

    let right_only: Vec<Aggregate> = right
        .iter()
        .enumerate()
        .filter(|(i, _)| !right_used[*i])
        .map(|(_, a)| a.clone())
        .collect();

    PairMatchOutput {
        matched,
        left_only,
        right_only,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
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

    #[test]
    fn exact_key_match() {
        let left = vec![
            agg("proc", "po_1", "USD", 7210, "2026-01-17"),
            agg("proc", "po_2", "USD", 5000, "2026-01-18"),
        ];
        let right = vec![
            agg("ledger", "po_1", "USD", 7210, "2026-01-18"),
            agg("ledger", "po_3", "USD", 3000, "2026-01-19"),
        ];
        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 2 };
        let out = match_exact_key(&left, &right, &tol);
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.matched[0].delta_cents, 0);
        assert_eq!(out.matched[0].date_offset_days, -1); // 17 - 18
        assert!(out.matched[0].within_tolerance);
        assert!(out.matched[0].within_window);
        assert_eq!(out.left_only.len(), 1);
        assert_eq!(out.left_only[0].match_key, "po_2");
        assert_eq!(out.right_only.len(), 1);
        assert_eq!(out.right_only[0].match_key, "po_3");
    }

    #[test]
    fn exact_key_no_cross_currency() {
        let left = vec![agg("proc", "po_1", "USD", 7210, "2026-01-17")];
        let right = vec![agg("ledger", "po_1", "CAD", 7210, "2026-01-17")];
        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 0 };
        let out = match_exact_key(&left, &right, &tol);
        assert_eq!(out.matched.len(), 0);
        assert_eq!(out.left_only.len(), 1);
        assert_eq!(out.right_only.len(), 1);
    }

    #[test]
    fn amount_mismatch_detected() {
        let left = vec![agg("proc", "po_1", "USD", 7210, "2026-01-17")];
        let right = vec![agg("ledger", "po_1", "USD", 7200, "2026-01-17")];
        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 0 };
        let out = match_exact_key(&left, &right, &tol);
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.matched[0].delta_cents, 10);
        assert!(!out.matched[0].within_tolerance);
    }

    #[test]
    fn date_window_exceeded() {
        let left = vec![agg("proc", "po_1", "USD", 7210, "2026-01-15")];
        let right = vec![agg("ledger", "po_1", "USD", 7210, "2026-01-20")];
        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 2 };
        let out = match_exact_key(&left, &right, &tol);
        assert_eq!(out.matched.len(), 1);
        assert!(out.matched[0].within_tolerance);
        assert!(!out.matched[0].within_window); // 5 days > 2
    }

    #[test]
    fn fuzzy_match_basic() {
        let left = vec![
            agg("proc", "po_1", "USD", 7210, "2026-01-17"),
            agg("proc", "po_2", "USD", 5000, "2026-01-18"),
        ];
        let right = vec![
            agg("ledger", "dep_a", "USD", 7210, "2026-01-18"),
            agg("ledger", "dep_b", "USD", 9999, "2026-01-18"),
        ];
        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 2 };
        let out = match_fuzzy_amount_date(&left, &right, &tol);
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.matched[0].left.match_key, "po_1");
        assert_eq!(out.matched[0].right.match_key, "dep_a");
        assert_eq!(out.left_only.len(), 1); // po_2 unmatched (5000 != 9999)
        assert_eq!(out.right_only.len(), 1); // dep_b unmatched
    }

    #[test]
    fn fuzzy_no_cross_currency() {
        let left = vec![agg("proc", "po_1", "USD", 7210, "2026-01-17")];
        let right = vec![agg("ledger", "dep_a", "CAD", 7210, "2026-01-17")];
        let tol = ToleranceConfig { amount_cents: 10, date_window_days: 5 };
        let out = match_fuzzy_amount_date(&left, &right, &tol);
        assert_eq!(out.matched.len(), 0);
    }
}
