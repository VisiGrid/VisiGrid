// Property-based tests for diff reconciliation logic.
// CI: 256 cases (default). Soak: PROPTEST_CASES=10000 cargo test --release

use std::collections::{BTreeMap, HashSet};

use proptest::prelude::*;
use visigrid_cli::diff::*;

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

fn config_256() -> ProptestConfig {
    ProptestConfig {
        cases: std::env::var("PROPTEST_CASES")
            .ok()
            .and_then(|s| s.parse().ok())
            .unwrap_or(256),
        failure_persistence: None,
        ..ProptestConfig::default()
    }
}

fn config_128() -> ProptestConfig {
    ProptestConfig {
        cases: std::env::var("PROPTEST_CASES")
            .ok()
            .and_then(|s| s.parse().ok())
            .unwrap_or(128),
        failure_persistence: None,
        ..ProptestConfig::default()
    }
}

// ---------------------------------------------------------------------------
// Shared headers
// ---------------------------------------------------------------------------

fn headers() -> Vec<String> {
    vec![
        "key".to_string(),
        "amount".to_string(),
        "label".to_string(),
        "qty".to_string(),
    ]
}

// ---------------------------------------------------------------------------
// Generators
// ---------------------------------------------------------------------------

/// Arbitrary value: mostly numeric, sometimes text, sometimes empty.
fn arb_value() -> impl Strategy<Value = String> {
    prop_oneof![
        3 => r"-?[0-9]{1,6}(\.[0-9]{1,2})?",
        1 => r"[a-zA-Z ]{0,15}",
        1 => Just("".to_string()),
    ]
}

/// Arbitrary tolerance: usually 0, sometimes positive.
fn arb_tolerance() -> impl Strategy<Value = f64> {
    prop_oneof![
        3 => Just(0.0),
        1 => 0.001..1000.0f64,
    ]
}

/// Build a DataRow from key and value columns.
fn make_row(key: &str, amount: &str, label: &str, qty: &str) -> DataRow {
    let mut values = std::collections::HashMap::new();
    values.insert("key".to_string(), key.to_string());
    values.insert("amount".to_string(), amount.to_string());
    values.insert("label".to_string(), label.to_string());
    values.insert("qty".to_string(), qty.to_string());
    DataRow {
        key_raw: key.to_string(),
        key_norm: key.to_string(),
        values,
    }
}

/// Category assignment for each key in exact-mode datasets.
#[derive(Debug, Clone, Copy, PartialEq)]
enum KeyCategory {
    Both,      // appears in both sides
    LeftOnly,  // left only
    RightOnly, // right only
}

/// Generate an exact-mode dataset with unique keys.
/// Returns (left_rows, right_rows, forced categories for verification).
fn arb_exact_dataset(
    max_keys: usize,
) -> impl Strategy<Value = (Vec<DataRow>, Vec<DataRow>, Vec<(String, KeyCategory)>)> {
    // Generate 1..=max_keys unique keys
    proptest::collection::hash_set(r"[A-Za-z0-9]{1,10}", 1..=max_keys)
        .prop_flat_map(move |keys| {
            let keys_vec: Vec<String> = keys.into_iter().collect();
            let n = keys_vec.len();
            // For each key: category + left values + right values
            let cats = if n >= 4 {
                // Force at least one of each: matched, diff, only_left, only_right
                // First 4 keys get forced categories
                let forced = vec![0u32, 0, 1, 2]; // 0=both, 1=left-only, 2=right-only
                let rest = proptest::collection::vec(0u32..3, n - 4);
                (Just(forced), rest)
                    .prop_map(|(mut f, r)| {
                        f.extend(r);
                        f
                    })
                    .boxed()
            } else {
                proptest::collection::vec(0u32..3, n).boxed()
            };
            let vals = proptest::collection::vec(
                (arb_value(), arb_value(), arb_value(), arb_value(), arb_value(), arb_value(), prop::bool::ANY),
                n,
            );
            (Just(keys_vec), cats, vals)
        })
        .prop_map(|(keys, cats, vals)| {
            let mut left = Vec::new();
            let mut right = Vec::new();
            let mut categories = Vec::new();

            for (i, key) in keys.iter().enumerate() {
                let cat = match cats[i] {
                    0 => KeyCategory::Both,
                    1 => KeyCategory::LeftOnly,
                    _ => KeyCategory::RightOnly,
                };
                categories.push((key.clone(), cat));

                let (la, ll, lq, ra, rl, rq, same) = &vals[i];

                match cat {
                    KeyCategory::Both => {
                        if *same {
                            // matched: identical values
                            left.push(make_row(key, la, ll, lq));
                            right.push(make_row(key, la, ll, lq));
                        } else {
                            // diff: different values
                            left.push(make_row(key, la, ll, lq));
                            right.push(make_row(key, ra, rl, rq));
                        }
                    }
                    KeyCategory::LeftOnly => {
                        left.push(make_row(key, la, ll, lq));
                    }
                    KeyCategory::RightOnly => {
                        right.push(make_row(key, ra, rl, rq));
                    }
                }
            }

            (left, right, categories)
        })
}

/// Generate a dataset that will have duplicate keys (for error testing).
fn arb_dataset_with_duplicates(
) -> impl Strategy<Value = (Vec<DataRow>, Vec<DataRow>)> {
    arb_exact_dataset(8).prop_flat_map(|(left, right, _)| {
        // Pick which side to duplicate and which row
        let left_len = left.len();
        let right_len = right.len();
        let total = left_len + right_len;
        if total == 0 {
            return Just((left, right)).boxed();
        }
        (0..total).prop_map(move |idx| {
            let mut l = left.clone();
            let mut r = right.clone();
            if idx < l.len() && !l.is_empty() {
                // Duplicate a left row
                let dup = l[idx % l.len()].clone();
                l.push(dup);
            } else if !r.is_empty() {
                // Duplicate a right row
                let ridx = if r.is_empty() { 0 } else { idx % r.len() };
                let dup = r[ridx].clone();
                r.push(dup);
            } else {
                // Both empty edge case — duplicate from left if possible
                if !l.is_empty() {
                    let dup = l[0].clone();
                    l.push(dup);
                }
            }
            (l, r)
        }).boxed()
    })
}

/// Generate a contains-mode dataset with forced unique matches, ambiguous matches,
/// and only-left entries.
fn arb_contains_dataset(
) -> impl Strategy<Value = (Vec<DataRow>, Vec<DataRow>)> {
    // Generate base keys: short alphanumeric strings
    let base_keys = proptest::collection::vec(r"[a-z]{2,5}", 3..=8);
    let extra_vals = proptest::collection::vec(
        (arb_value(), arb_value(), arb_value()),
        10..=20,
    );
    (base_keys, extra_vals).prop_map(|(bases, vals)| {
        let mut left = Vec::new();
        let mut right = Vec::new();
        let mut val_idx = 0;
        let next_val = |idx: &mut usize| -> (String, String, String) {
            let v = if *idx < vals.len() {
                vals[*idx].clone()
            } else {
                ("0".to_string(), "x".to_string(), "1".to_string())
            };
            *idx += 1;
            v
        };

        // 1. Unique match: left key = base, right key = prefix+base+suffix (only one)
        if let Some(base) = bases.first() {
            let left_key = base.clone();
            let right_key = format!("X{}Y", base);
            let (a, l, q) = next_val(&mut val_idx);
            left.push(make_row(&left_key, &a, &l, &q));
            let (a2, l2, q2) = next_val(&mut val_idx);
            right.push(make_row(&right_key, &a2, &l2, &q2));
        }

        // 2. Ambiguous match: left key = base, right has 2+ keys containing it
        if bases.len() >= 2 {
            let base = &bases[1];
            let left_key = base.clone();
            let right_key1 = format!("P{}", base);
            let right_key2 = format!("Q{}", base);
            let (a, l, q) = next_val(&mut val_idx);
            left.push(make_row(&left_key, &a, &l, &q));
            let (a2, l2, q2) = next_val(&mut val_idx);
            right.push(make_row(&right_key1, &a2, &l2, &q2));
            let (a3, l3, q3) = next_val(&mut val_idx);
            right.push(make_row(&right_key2, &a3, &l3, &q3));
        }

        // 3. Only-left: left key that won't match any right key
        if bases.len() >= 3 {
            let unique_key = format!("ZZZUNIQUE{}", &bases[2]);
            let (a, l, q) = next_val(&mut val_idx);
            left.push(make_row(&unique_key, &a, &l, &q));
            // Don't add to right — no right key contains "ZZZUNIQUE..."
        }

        // Add some extra right-only rows
        for i in 3..bases.len().min(6) {
            let right_key = format!("RONLY{}", &bases[i]);
            let (a, l, q) = next_val(&mut val_idx);
            right.push(make_row(&right_key, &a, &l, &q));
        }

        (left, right)
    })
}

/// Generate a financial number string from a known f64 value.
fn arb_financial_string() -> impl Strategy<Value = (f64, String)> {
    let value = (-9_999_999.0f64..9_999_999.0f64).prop_map(|v| (v * 100.0).round() / 100.0);
    let style = 0u32..8;
    (value, style).prop_map(|(v, style)| {
        let formatted = format_financial(v, style);
        (v, formatted)
    })
}

fn format_financial(v: f64, style: u32) -> String {
    let abs = v.abs();
    let is_neg = v < 0.0;
    let is_whole = abs == abs.floor();

    match style {
        0 => {
            // Plain: 1234.56 or -50.00
            format!("{:.2}", v)
        }
        1 => {
            // Commas: 1,234.56
            let base = if is_neg {
                format!("-{}", add_commas(abs, !is_whole))
            } else {
                add_commas(abs, !is_whole)
            };
            base
        }
        2 => {
            // Dollar: $1,234.56
            if is_neg {
                format!("$-{}", add_commas(abs, true))
            } else {
                format!("${}", add_commas(abs, true))
            }
        }
        3 => {
            // Parens for negative: (500.00) or (1,234.56)
            if is_neg {
                format!("({})", add_commas(abs, true))
            } else {
                add_commas(abs, true)
            }
        }
        4 => {
            // Dollar + parens: ($1,234.56)
            if is_neg {
                format!("(${})", add_commas(abs, true))
            } else {
                format!("${}", add_commas(abs, true))
            }
        }
        5 => {
            // Integer+commas when whole
            if is_whole {
                if is_neg {
                    format!("-{}", add_commas(abs, false))
                } else {
                    add_commas(abs, false)
                }
            } else {
                format!("{:.2}", v)
            }
        }
        6 => {
            // Whitespace-padded
            format!("  {:.2}  ", v)
        }
        7 => {
            // Whitespace-padded + dollar
            if is_neg {
                format!("  $-{}  ", add_commas(abs, true))
            } else {
                format!("  ${}  ", add_commas(abs, true))
            }
        }
        _ => format!("{:.2}", v),
    }
}

fn add_commas(v: f64, with_decimals: bool) -> String {
    let int_part = v.floor() as u64;
    let s = int_part.to_string();
    let mut result = String::new();
    for (i, ch) in s.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(ch);
    }
    let result: String = result.chars().rev().collect();
    if with_decimals {
        let frac = ((v - v.floor()) * 100.0).round() as u64;
        format!("{}.{:02}", result, frac)
    } else {
        result
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn exact_opts(tolerance: f64) -> DiffOptions {
    DiffOptions {
        key_col: 0,
        compare_cols: None,
        match_mode: MatchMode::Exact,
        key_transform: KeyTransform::None,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance,
    }
}

fn contains_opts(policy: AmbiguityPolicy) -> DiffOptions {
    DiffOptions {
        key_col: 0,
        compare_cols: None,
        match_mode: MatchMode::Contains,
        key_transform: KeyTransform::None,
        on_ambiguous: policy,
        tolerance: 0.0,
    }
}

/// Build a map from key → (status, diffs) for by-key comparison.
fn build_key_map(result: &DiffResult) -> BTreeMap<String, &DiffRow> {
    let mut map = BTreeMap::new();
    for row in &result.results {
        map.insert(row.key.clone(), row);
    }
    map
}

// ===========================================================================
// Phase 1A — Core (256 cases)
// ===========================================================================

// Test 1: Determinism
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn determinism(
        (left, right, _cats) in arb_exact_dataset(20),
        tolerance in arb_tolerance(),
    ) {
        let hdrs = headers();
        let opts = exact_opts(tolerance);

        let r1 = reconcile(&left, &right, &hdrs, &opts);
        let r2 = reconcile(&left, &right, &hdrs, &opts);

        match (r1, r2) {
            (Ok(r1), Ok(r2)) => {
                // Results length identical
                prop_assert_eq!(r1.results.len(), r2.results.len(),
                    "Result count mismatch");

                // Per-row structural equality
                for (i, (a, b)) in r1.results.iter().zip(r2.results.iter()).enumerate() {
                    prop_assert_eq!(a.status, b.status,
                        "Row {} status mismatch", i);
                    prop_assert_eq!(&a.key, &b.key,
                        "Row {} key mismatch", i);
                    prop_assert_eq!(a.diffs.len(), b.diffs.len(),
                        "Row {} diffs count mismatch", i);
                    for (j, (d1, d2)) in a.diffs.iter().zip(b.diffs.iter()).enumerate() {
                        prop_assert_eq!(&d1.column, &d2.column,
                            "Row {} diff {} column mismatch", i, j);
                        prop_assert_eq!(&d1.left, &d2.left,
                            "Row {} diff {} left mismatch", i, j);
                        prop_assert_eq!(&d1.right, &d2.right,
                            "Row {} diff {} right mismatch", i, j);
                        prop_assert_eq!(d1.delta, d2.delta,
                            "Row {} diff {} delta mismatch", i, j);
                        prop_assert_eq!(d1.within_tolerance, d2.within_tolerance,
                            "Row {} diff {} within_tolerance mismatch", i, j);
                    }
                    prop_assert_eq!(&a.left, &b.left,
                        "Row {} left values mismatch", i);
                    prop_assert_eq!(&a.right, &b.right,
                        "Row {} right values mismatch", i);
                }

                // Ambiguous keys identical
                prop_assert_eq!(r1.ambiguous_keys.len(), r2.ambiguous_keys.len());
                for (a, b) in r1.ambiguous_keys.iter().zip(r2.ambiguous_keys.iter()) {
                    prop_assert_eq!(&a.key, &b.key);
                }

                // Summary identical
                prop_assert_eq!(r1.summary.left_rows, r2.summary.left_rows);
                prop_assert_eq!(r1.summary.right_rows, r2.summary.right_rows);
                prop_assert_eq!(r1.summary.matched, r2.summary.matched);
                prop_assert_eq!(r1.summary.only_left, r2.summary.only_left);
                prop_assert_eq!(r1.summary.only_right, r2.summary.only_right);
                prop_assert_eq!(r1.summary.diff, r2.summary.diff);
                prop_assert_eq!(r1.summary.diff_outside_tolerance, r2.summary.diff_outside_tolerance);
                prop_assert_eq!(r1.summary.ambiguous, r2.summary.ambiguous);
            }
            (Err(_), Err(_)) => {
                // Both error — fine, deterministic
            }
            _ => {
                prop_assert!(false, "One succeeded and one failed");
            }
        }
    }
}

// Test 2: Tolerance monotonicity
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn tolerance_monotonicity(
        (left, right, _cats) in arb_exact_dataset(15),
        t1 in 0.0..500.0f64,
        gap in 0.001..500.0f64,
    ) {
        let hdrs = headers();
        let t2 = t1 + gap; // t2 > t1

        let opts1 = exact_opts(t1);
        let opts2 = exact_opts(t2);

        let r1 = reconcile(&left, &right, &hdrs, &opts1);
        let r2 = reconcile(&left, &right, &hdrs, &opts2);

        if let (Ok(r1), Ok(r2)) = (r1, r2) {
            // Count diffs outside tolerance
            let outside1 = r1.results.iter()
                .filter(|r| r.status == RowStatus::Diff && r.diffs.iter().any(|d| !d.within_tolerance))
                .count();
            let outside2 = r2.results.iter()
                .filter(|r| r.status == RowStatus::Diff && r.diffs.iter().any(|d| !d.within_tolerance))
                .count();

            prop_assert!(outside2 <= outside1,
                "Larger tolerance t2={} has {} outside vs t1={} has {} outside",
                t2, outside2, t1, outside1);
        }
    }
}

// Test 3: Exit code correctness (summary vs per-row recount)
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn exit_code_correctness(
        (left, right, _cats) in arb_exact_dataset(20),
        tolerance in arb_tolerance(),
    ) {
        let hdrs = headers();
        let opts = exact_opts(tolerance);
        let result = reconcile(&left, &right, &hdrs, &opts);

        if let Ok(result) = result {
            let summary = &result.summary;

            // Recompute from per-row data
            let actual_matched = result.results.iter()
                .filter(|r| r.status == RowStatus::Matched).count();
            let actual_diff = result.results.iter()
                .filter(|r| r.status == RowStatus::Diff).count();
            let actual_only_left = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyLeft).count();
            let actual_only_right = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyRight).count();
            let actual_diff_outside = result.results.iter()
                .filter(|r| r.status == RowStatus::Diff && r.diffs.iter().any(|d| !d.within_tolerance))
                .count();

            // Summary matches per-row reality
            prop_assert_eq!(summary.matched, actual_matched,
                "summary.matched mismatch");
            prop_assert_eq!(summary.diff, actual_diff,
                "summary.diff mismatch");
            prop_assert_eq!(summary.only_left, actual_only_left,
                "summary.only_left mismatch");
            prop_assert_eq!(summary.only_right, actual_only_right,
                "summary.only_right mismatch");
            prop_assert_eq!(summary.diff_outside_tolerance, actual_diff_outside,
                "summary.diff_outside_tolerance mismatch");
            prop_assert_eq!(summary.left_rows, left.len(),
                "summary.left_rows mismatch");
            prop_assert_eq!(summary.right_rows, right.len(),
                "summary.right_rows mismatch");

            // Non-strict: fail iff material drift
            let should_fail_nonstrict = actual_only_left > 0
                || actual_only_right > 0
                || actual_diff_outside > 0;
            // Strict: fail iff any diff at all
            let should_fail_strict = actual_only_left > 0
                || actual_only_right > 0
                || actual_diff > 0;
            // Strict is always stricter
            if should_fail_nonstrict {
                prop_assert!(should_fail_strict,
                    "Non-strict failure implies strict failure");
            }
        }
    }
}

// Test 4: Duplicate keys always error
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn duplicate_keys_always_error(
        (left, right) in arb_dataset_with_duplicates(),
    ) {
        let hdrs = headers();
        let opts = exact_opts(0.0);

        // Verify we actually have duplicates on at least one side
        let has_left_dups = {
            let mut seen = HashSet::new();
            left.iter().any(|r| !seen.insert(&r.key_norm))
        };
        let has_right_dups = {
            let mut seen = HashSet::new();
            right.iter().any(|r| !seen.insert(&r.key_norm))
        };

        if has_left_dups || has_right_dups {
            let result = reconcile(&left, &right, &hdrs, &opts);
            match result {
                Err(DiffError::DuplicateKeys(dups)) => {
                    prop_assert!(!dups.is_empty(),
                        "DuplicateKeys error should have non-empty dups");
                }
                Ok(_) => {
                    prop_assert!(false,
                        "Expected DuplicateKeys error but got Ok");
                }
            }
        }
        // If the generator happened not to produce duplicates, skip
    }
}

// ===========================================================================
// Phase 1B — Symmetry + accounting (128 cases)
// ===========================================================================

// Test 5: Symmetry (exact mode, by key)
proptest! {
    #![proptest_config(config_128())]
    #[test]
    fn symmetry_exact_mode(
        (left, right, _cats) in arb_exact_dataset(15),
        tolerance in arb_tolerance(),
    ) {
        let hdrs = headers();
        let opts = exact_opts(tolerance);

        let r_lr = reconcile(&left, &right, &hdrs, &opts);
        let r_rl = reconcile(&right, &left, &hdrs, &opts);

        if let (Ok(r_lr), Ok(r_rl)) = (r_lr, r_rl) {
            let forward_map = build_key_map(&r_lr);
            let reverse_map = build_key_map(&r_rl);

            for (key, fwd) in &forward_map {
                let rev = reverse_map.get(key);
                prop_assert!(rev.is_some(),
                    "Key {:?} in forward but not reverse", key);
                let rev = rev.unwrap();

                match fwd.status {
                    RowStatus::Matched => {
                        prop_assert_eq!(rev.status, RowStatus::Matched,
                            "Key {:?}: forward Matched but reverse {:?}", key, rev.status);
                    }
                    RowStatus::Diff => {
                        prop_assert_eq!(rev.status, RowStatus::Diff,
                            "Key {:?}: forward Diff but reverse {:?}", key, rev.status);
                        // Deltas identical (abs), left/right swap
                        prop_assert_eq!(fwd.diffs.len(), rev.diffs.len(),
                            "Key {:?}: diff count mismatch", key);
                        for (fd, rd) in fwd.diffs.iter().zip(rev.diffs.iter()) {
                            prop_assert_eq!(fd.delta, rd.delta,
                                "Key {:?} col {:?}: delta mismatch", key, fd.column);
                            prop_assert_eq!(fd.within_tolerance, rd.within_tolerance,
                                "Key {:?} col {:?}: tolerance mismatch", key, fd.column);
                            prop_assert_eq!(&fd.left, &rd.right,
                                "Key {:?} col {:?}: left/right swap failed", key, fd.column);
                            prop_assert_eq!(&fd.right, &rd.left,
                                "Key {:?} col {:?}: right/left swap failed", key, fd.column);
                        }
                    }
                    RowStatus::OnlyLeft => {
                        prop_assert_eq!(rev.status, RowStatus::OnlyRight,
                            "Key {:?}: forward OnlyLeft but reverse {:?}", key, rev.status);
                    }
                    RowStatus::OnlyRight => {
                        prop_assert_eq!(rev.status, RowStatus::OnlyLeft,
                            "Key {:?}: forward OnlyRight but reverse {:?}", key, rev.status);
                    }
                    _ => {
                        prop_assert!(false, "Unexpected status in exact mode: {:?}", fwd.status);
                    }
                }
            }

            // Summary symmetry
            prop_assert_eq!(r_lr.summary.matched, r_rl.summary.matched);
            prop_assert_eq!(r_lr.summary.diff, r_rl.summary.diff);
            prop_assert_eq!(r_lr.summary.only_left, r_rl.summary.only_right);
            prop_assert_eq!(r_lr.summary.only_right, r_rl.summary.only_left);
            prop_assert_eq!(r_lr.ambiguous_keys.len(), 0);
            prop_assert_eq!(r_rl.ambiguous_keys.len(), 0);
        }
    }
}

// Test 6: No silent dropping — exact mode
proptest! {
    #![proptest_config(config_128())]
    #[test]
    fn no_silent_dropping_exact(
        (left, right, _cats) in arb_exact_dataset(20),
        tolerance in arb_tolerance(),
    ) {
        let hdrs = headers();
        let opts = exact_opts(tolerance);
        let result = reconcile(&left, &right, &hdrs, &opts);

        if let Ok(result) = result {
            let actual_matched = result.results.iter()
                .filter(|r| r.status == RowStatus::Matched).count();
            let actual_diff = result.results.iter()
                .filter(|r| r.status == RowStatus::Diff).count();
            let actual_only_left = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyLeft).count();
            let actual_only_right = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyRight).count();

            // Accounting identities
            prop_assert_eq!(
                actual_matched + actual_diff + actual_only_left,
                left.len(),
                "Left accounting: {} matched + {} diff + {} only_left != {} left rows",
                actual_matched, actual_diff, actual_only_left, left.len()
            );
            prop_assert_eq!(
                actual_matched + actual_diff + actual_only_right,
                right.len(),
                "Right accounting: {} matched + {} diff + {} only_right != {} right rows",
                actual_matched, actual_diff, actual_only_right, right.len()
            );
            prop_assert_eq!(result.ambiguous_keys.len(), 0,
                "Exact mode should have no ambiguous keys");

            // Summary matches reality
            prop_assert_eq!(result.summary.matched, actual_matched);
            prop_assert_eq!(result.summary.diff, actual_diff);
            prop_assert_eq!(result.summary.only_left, actual_only_left);
            prop_assert_eq!(result.summary.only_right, actual_only_right);
            prop_assert_eq!(result.summary.left_rows, left.len());
            prop_assert_eq!(result.summary.right_rows, right.len());
        }
    }
}

// Test 7: No silent dropping — contains mode
proptest! {
    #![proptest_config(config_128())]
    #[test]
    fn no_silent_dropping_contains(
        (left, right) in arb_contains_dataset(),
        policy_idx in 0u32..2,
    ) {
        let policy = if policy_idx == 0 {
            AmbiguityPolicy::Error
        } else {
            AmbiguityPolicy::Report
        };
        let hdrs = headers();
        let opts = contains_opts(policy);

        let result = reconcile(&left, &right, &hdrs, &opts);
        if let Ok(result) = result {
            let actual_matched = result.results.iter()
                .filter(|r| r.status == RowStatus::Matched).count();
            let actual_diff = result.results.iter()
                .filter(|r| r.status == RowStatus::Diff).count();
            let actual_only_left = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyLeft).count();
            let actual_only_right = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyRight).count();
            let actual_ambiguous_in_results = result.results.iter()
                .filter(|r| r.status == RowStatus::Ambiguous).count();
            let ambiguous_count = result.ambiguous_keys.len();

            // Left accounting: matched + diff + only_left + ambiguous == left_rows
            prop_assert_eq!(
                actual_matched + actual_diff + actual_only_left + ambiguous_count,
                left.len(),
                "Left accounting: {} matched + {} diff + {} only_left + {} ambiguous != {} left",
                actual_matched, actual_diff, actual_only_left, ambiguous_count, left.len()
            );

            // Right accounting: matched + diff + only_right == right_rows
            prop_assert_eq!(
                actual_matched + actual_diff + actual_only_right,
                right.len(),
                "Right accounting: {} matched + {} diff + {} only_right != {} right",
                actual_matched, actual_diff, actual_only_right, right.len()
            );

            // Policy-specific: Error → ambiguous NOT in results; Report → in results
            match policy {
                AmbiguityPolicy::Error => {
                    prop_assert_eq!(actual_ambiguous_in_results, 0,
                        "Error policy: ambiguous rows should not be in results");
                }
                AmbiguityPolicy::Report => {
                    prop_assert_eq!(actual_ambiguous_in_results, ambiguous_count,
                        "Report policy: ambiguous in results should match ambiguous_keys count");
                }
            }

            // Summary consistency
            prop_assert_eq!(result.summary.ambiguous, ambiguous_count);
        }
    }
}

// ===========================================================================
// Phase 1C — Parsing + ordering
// ===========================================================================

// Test 8: parse_financial_number roundtrip
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn financial_number_roundtrip(
        (expected, formatted) in arb_financial_string(),
    ) {
        let parsed = parse_financial_number(&formatted);
        prop_assert!(parsed.is_some(),
            "Failed to parse {:?} (expected {})", formatted, expected);
        let parsed = parsed.unwrap();

        // Exact cents equality
        let expected_cents = (expected * 100.0).round() as i64;
        let parsed_cents = (parsed * 100.0).round() as i64;
        prop_assert_eq!(expected_cents, parsed_cents,
            "Parsed {} from {:?}, expected {} (cents: {} vs {})",
            parsed, formatted, expected, parsed_cents, expected_cents);
    }
}

// Rejection tests for parse_financial_number
#[test]
fn financial_number_rejects_alpha() {
    assert_eq!(parse_financial_number("hello"), None);
    assert_eq!(parse_financial_number("abc"), None);
    assert_eq!(parse_financial_number("N/A"), None);
}

#[test]
fn financial_number_rejects_empty() {
    assert_eq!(parse_financial_number(""), None);
}

#[test]
fn financial_number_rejects_whitespace() {
    assert_eq!(parse_financial_number("   "), None);
    assert_eq!(parse_financial_number("\t\n"), None);
}

// Test 9: apply_key_transform idempotence
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn key_transform_idempotence(
        raw in r"[ \tA-Za-z0-9\-]{0,20}",
        transform_idx in 0u32..3,
    ) {
        let transform = match transform_idx {
            0 => KeyTransform::None,
            1 => KeyTransform::Trim,
            _ => KeyTransform::Digits,
        };

        let once = apply_key_transform(&raw, transform);
        let twice = apply_key_transform(&once, transform);
        prop_assert_eq!(&once, &twice,
            "Transform {:?} not idempotent: {:?} -> {:?} -> {:?}",
            transform, raw, once, twice);
    }
}

// Test 10: Stable result ordering (exact mode)
proptest! {
    #![proptest_config(config_256())]
    #[test]
    fn stable_result_ordering(
        (left, right, _cats) in arb_exact_dataset(15),
    ) {
        let hdrs = headers();
        let opts = exact_opts(0.0);
        let result = reconcile(&left, &right, &hdrs, &opts);

        if let Ok(result) = result {
            // Build expected order: left-derived results in left-input order,
            // then only_right in right-input order.

            let left_keys: Vec<&str> = left.iter().map(|r| r.key_norm.as_str()).collect();
            let right_keys: Vec<&str> = right.iter().map(|r| r.key_norm.as_str()).collect();

            // Split results into left-derived (Matched, Diff, OnlyLeft) and right-only
            let left_derived: Vec<&DiffRow> = result.results.iter()
                .filter(|r| matches!(r.status, RowStatus::Matched | RowStatus::Diff | RowStatus::OnlyLeft))
                .collect();
            let right_only: Vec<&DiffRow> = result.results.iter()
                .filter(|r| r.status == RowStatus::OnlyRight)
                .collect();

            // Left-derived count should equal left input count
            prop_assert_eq!(left_derived.len(), left.len(),
                "Left-derived count {} != left input count {}", left_derived.len(), left.len());

            // Left-derived keys should be in left-input order
            let left_derived_keys: Vec<&str> = left_derived.iter()
                .map(|r| r.key.as_str()).collect();
            prop_assert_eq!(left_derived_keys, left_keys,
                "Left-derived keys not in left-input order");

            // Right-only keys should be in right-input order (preserving only unmatched)
            let consumed_keys: HashSet<&str> = left_derived.iter()
                .filter(|r| matches!(r.status, RowStatus::Matched | RowStatus::Diff))
                .map(|r| r.key.as_str())
                .collect();
            let expected_right_only: Vec<&str> = right_keys.iter()
                .filter(|k| !consumed_keys.contains(*k))
                .copied()
                .collect();
            let actual_right_only: Vec<&str> = right_only.iter()
                .map(|r| r.key.as_str()).collect();
            prop_assert_eq!(actual_right_only, expected_right_only,
                "Right-only keys not in right-input order");
        }
    }
}

// ===========================================================================
// Phase 2: Metamorphic + fixture tests
// ===========================================================================

// ---------------------------------------------------------------------------
// Phase 2 helpers
// ---------------------------------------------------------------------------

/// Reconcile single-row pair comparing only the "amount" column.
fn reconcile_pair(left_val: &str, right_val: &str, tol: f64) -> (RowStatus, Vec<ColumnDiff>) {
    let hdrs = headers();
    let mut opts = exact_opts(tol);
    opts.compare_cols = Some(vec![1]); // index 1 = "amount"
    let left = vec![make_row("K", left_val, "", "")];
    let right = vec![make_row("K", right_val, "", "")];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();
    let row = &result.results[0];
    (row.status, row.diffs.clone())
}

fn assert_matched(left: &str, right: &str) {
    let (status, diffs) = reconcile_pair(left, right, 0.0);
    assert_eq!(
        status,
        RowStatus::Matched,
        "{:?} vs {:?}: expected Matched, got {:?} with diffs {:?}",
        left,
        right,
        status,
        diffs
    );
}

fn assert_diff_numeric(left: &str, right: &str, tol: f64, expect_within: bool) {
    let (status, diffs) = reconcile_pair(left, right, tol);
    assert_eq!(
        status,
        RowStatus::Diff,
        "{:?} vs {:?}: expected Diff, got {:?}",
        left,
        right,
        status
    );
    assert_eq!(
        diffs.len(),
        1,
        "{:?} vs {:?}: expected 1 diff, got {}",
        left,
        right,
        diffs.len()
    );
    assert!(
        diffs[0].delta.is_some(),
        "{:?} vs {:?}: expected numeric delta",
        left,
        right
    );
    assert_eq!(
        diffs[0].within_tolerance, expect_within,
        "{:?} vs {:?}: within_tolerance mismatch",
        left,
        right
    );
}

fn assert_diff_nonnumeric(left: &str, right: &str) {
    let (status, diffs) = reconcile_pair(left, right, 0.0);
    assert_eq!(
        status,
        RowStatus::Diff,
        "{:?} vs {:?}: expected Diff, got {:?}",
        left,
        right,
        status
    );
    assert_eq!(
        diffs.len(),
        1,
        "{:?} vs {:?}: expected 1 diff, got {}",
        left,
        right,
        diffs.len()
    );
    assert!(
        diffs[0].delta.is_none(),
        "{:?} vs {:?}: expected delta=None (non-numeric)",
        left,
        right
    );
}

// ---------------------------------------------------------------------------
// Group 1: Positive equivalences
// ---------------------------------------------------------------------------

#[test]
fn bank_export_positive_equivalences() {
    assert_matched("$1,234.50", "1234.50");
    assert_matched("1,234.50", "1234.5");
    assert_matched("$1,234.50", "1234.5");
    assert_matched("  $1,234.50  ", "1234.50");
}

#[test]
fn bank_export_integer_vs_decimal() {
    assert_matched("1234", "1234.00");
    assert_matched("1,234", "1234.00");
    assert_matched("$1,234", "1234");
    assert_matched("0", "0.00");
}

#[test]
fn bank_export_leading_zeros() {
    assert_matched("01234.56", "1234.56");
    assert_matched("007", "7");
    assert_matched("00.50", "0.50");
}

#[test]
fn bank_export_plus_sign() {
    // '+' is explicitly allowed at position 0 (diff.rs:229). Lock in existing behavior.
    assert_matched("+1234.56", "1234.56");
    assert_matched("+0", "0");
}

// ---------------------------------------------------------------------------
// Group 2: Negative equivalences
// ---------------------------------------------------------------------------

#[test]
fn bank_export_negative_equivalences() {
    assert_matched("(1,234.50)", "-1234.50");
    assert_matched("($1,234.50)", "-1234.5");
    assert_matched("-$1,234.50", "-1234.50"); // strip $ → -1234.50 → parses ok
}

#[test]
fn bank_export_negative_zero() {
    assert_matched("-0.00", "0.00"); // f64: (-0.0 - 0.0).abs() == 0.0
    assert_matched("(0.00)", "0.00");
    assert_matched("-0", "0");
    assert_matched("$0.00", "-$0.00");
}

#[test]
fn bank_export_negative_paren_dollar() {
    assert_matched("($500.00)", "(500.00)"); // both → -500.0
    assert_matched("($1,234.56)", "-1234.56");
}

// ---------------------------------------------------------------------------
// Group 3: Diffs detected
// ---------------------------------------------------------------------------

#[test]
fn bank_export_real_diff() {
    // Different values, correct delta
    assert_diff_numeric("$1,234.56", "$1,234.00", 0.0, false);
    assert_diff_numeric("(500.00)", "-499.50", 0.0, false);
}

#[test]
fn bank_export_mixed_type_diff() {
    // One numeric, one non-numeric → string comparison
    assert_diff_nonnumeric("1234.56", "N/A");
    assert_diff_nonnumeric("$0.00", "pending");
    assert_diff_nonnumeric("N/A", "TBD");
}

// ---------------------------------------------------------------------------
// Group 4: Empty cells
// ---------------------------------------------------------------------------

#[test]
fn compare_values_empty_cells() {
    assert_matched("", ""); // both empty → continue, no diff
    assert_diff_nonnumeric("", "0"); // empty → parse returns None → string comparison
    assert_diff_nonnumeric("0", "");
}

// ---------------------------------------------------------------------------
// Group 5: Tolerance + formatting
// ---------------------------------------------------------------------------

#[test]
fn bank_export_tolerance_with_formatting() {
    // tol=1.0
    assert_diff_numeric("$1,234.56", "1234.00", 1.0, true); // delta ≈ 0.56
    assert_diff_numeric("(500.00)", "-500.99", 1.0, true); // delta = 0.99
    assert_diff_numeric("$1,234.56", "1232.00", 1.0, false); // delta ≈ 2.56
}

// ---------------------------------------------------------------------------
// Group 6: Rounding edge (3-decimal strings)
// ---------------------------------------------------------------------------

#[test]
fn bank_export_rounding_edge() {
    // Classic "bank export had 3 decimals"
    assert_diff_numeric("1.005", "1.01", 0.0, false); // delta=0.005
    assert_diff_numeric("1.005", "1.00", 0.0, false); // delta=0.005
    assert_diff_numeric("1.005", "1.01", 0.01, true); // 0.005 <= 0.01
    assert_diff_numeric("1.005", "1.00", 0.01, true); // 0.005 <= 0.01
}

// ---------------------------------------------------------------------------
// Group 7: Unicode
// ---------------------------------------------------------------------------

#[test]
fn unicode_nfc_nfd_key_mismatch() {
    // Unicode normalization is not applied. Keys must be byte-identical.
    // NFC "José" vs NFD "Jose\u{0301}" are byte-different → 1 only_left + 1 only_right.
    let hdrs = headers();
    let opts = exact_opts(0.0);
    let left = vec![make_row("Jos\u{00e9}", "100", "", "")];
    let right = vec![make_row("Jose\u{0301}", "100", "", "")];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    let only_left = result
        .results
        .iter()
        .filter(|r| r.status == RowStatus::OnlyLeft)
        .count();
    let only_right = result
        .results
        .iter()
        .filter(|r| r.status == RowStatus::OnlyRight)
        .count();
    assert_eq!(only_left, 1, "expected 1 only_left for NFC key");
    assert_eq!(only_right, 1, "expected 1 only_right for NFD key");
}

#[test]
fn unicode_nfc_nfd_value_diff() {
    // Unicode normalization is not applied to compared values either.
    // Same key, values differ only by NFC/NFD → string diff, delta=None.
    let hdrs = headers();
    let mut opts = exact_opts(0.0);
    opts.compare_cols = Some(vec![1]); // compare "amount" column only
    let left = vec![make_row("K", "Jos\u{00e9}", "", "")];
    let right = vec![make_row("K", "Jose\u{0301}", "", "")];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.results.len(), 1);
    let row = &result.results[0];
    assert_eq!(row.status, RowStatus::Diff);
    assert_eq!(row.diffs.len(), 1);
    assert!(
        row.diffs[0].delta.is_none(),
        "expected delta=None for non-numeric NFC/NFD value diff"
    );
}

// ---------------------------------------------------------------------------
// Group 8: NBSP / weird whitespace
// ---------------------------------------------------------------------------

#[test]
fn bank_export_nbsp_whitespace() {
    // Rust's trim() and is_whitespace() handle NBSP.
    assert_matched("$1,234.50\u{00A0}", "1234.50"); // NBSP trimmed/stripped
    assert_matched("\u{00A0}1234.50\u{00A0}", "1234.50");
}

// ---------------------------------------------------------------------------
// Group 9: Full reconcile integration
// ---------------------------------------------------------------------------

#[test]
fn reconcile_formatting_match_is_matched_status() {
    // 3-row dataset, left uses $ formatting, right plain.
    let hdrs = headers();
    let opts = exact_opts(0.0);
    let left = vec![
        make_row("INV-001", "$1,234.56", "", ""),
        make_row("INV-002", "($500.00)", "", ""),
        make_row("INV-003", "$0.00", "", ""),
    ];
    let right = vec![
        make_row("INV-001", "1234.56", "", ""),
        make_row("INV-002", "-500", "", ""),
        make_row("INV-003", "0", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.matched, 3, "expected 3 matched rows");
    assert_eq!(result.summary.diff, 0, "expected 0 diff rows");
    assert_eq!(result.summary.only_left, 0);
    assert_eq!(result.summary.only_right, 0);
}

// ===========================================================================
// Phase 3: Feature interaction fixtures
// ===========================================================================
//
// Pairwise coverage of interacting features:
//   MatchMode (Exact/Contains) × KeyTransform (None/Trim/Digits)
//   × Tolerance (0/small/larger) × AmbiguityPolicy × compare_cols

// ---------------------------------------------------------------------------
// Phase 3 helpers
// ---------------------------------------------------------------------------

/// Build a DataRow with key_norm pre-computed via apply_key_transform.
fn make_row_transformed(
    key_raw: &str,
    transform: KeyTransform,
    amount: &str,
    label: &str,
    qty: &str,
) -> DataRow {
    let key_norm = apply_key_transform(key_raw, transform);
    let mut values = std::collections::HashMap::new();
    values.insert("key".to_string(), key_raw.to_string());
    values.insert("amount".to_string(), amount.to_string());
    values.insert("label".to_string(), label.to_string());
    values.insert("qty".to_string(), qty.to_string());
    DataRow {
        key_raw: key_raw.to_string(),
        key_norm,
        values,
    }
}

// ---------------------------------------------------------------------------
// Interaction tests
// ---------------------------------------------------------------------------

#[test]
fn combo_exact_digits_tolerance() {
    // Exact + Digits + tolerance=1.0
    // Keys "INV-001"/"INV-002" → digits → "001"/"002"
    // Use exact f64 deltas to avoid floating-point boundary surprises.
    let hdrs = headers();
    let t = KeyTransform::Digits;
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Exact,
        key_transform: t,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 1.0,
    };
    let left = vec![
        make_row_transformed("INV-001", t, "$101.00", "", ""),
        make_row_transformed("INV-002", t, "200", "", ""),
    ];
    let right = vec![
        make_row_transformed("PO-001", t, "100.00", "", ""),
        make_row_transformed("PO-002", t, "200.00", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    // "001" matches: delta=1.0, within tolerance (1.0 <= 1.0)
    // "002" matches: numerically equal → Matched
    assert_eq!(result.summary.matched, 1);
    assert_eq!(result.summary.diff, 1);
    assert_eq!(result.summary.diff_outside_tolerance, 0);
}

#[test]
fn combo_exact_digits_tolerance_zero() {
    // Same as above but tolerance=0 → the 0.01 diff is outside tolerance
    let hdrs = headers();
    let t = KeyTransform::Digits;
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Exact,
        key_transform: t,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 0.0,
    };
    let left = vec![
        make_row_transformed("INV-001", t, "$100.50", "", ""),
        make_row_transformed("INV-002", t, "200", "", ""),
    ];
    let right = vec![
        make_row_transformed("PO-001", t, "100.49", "", ""),
        make_row_transformed("PO-002", t, "200.00", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.matched, 1);
    assert_eq!(result.summary.diff, 1);
    assert_eq!(result.summary.diff_outside_tolerance, 1);
}

#[test]
fn combo_contains_digits_ambiguity_error() {
    // Contains + Digits + tolerance=0.01 + Error policy
    // Left "12" → digits "12". Right "A-123"→"123", "B-125"→"125".
    // "123".contains("12") and "125".contains("12") → ambiguous.
    let hdrs = headers();
    let t = KeyTransform::Digits;
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Contains,
        key_transform: t,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 0.01,
    };
    let left = vec![make_row_transformed("12", t, "$100.00", "", "")];
    let right = vec![
        make_row_transformed("A-123", t, "100.01", "", ""),
        make_row_transformed("B-125", t, "99.99", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    // Ambiguous → not matched, right rows unconsumed → OnlyRight
    assert_eq!(result.summary.ambiguous, 1);
    assert_eq!(result.summary.only_right, 2);
    assert_eq!(result.summary.matched, 0);
    assert_eq!(result.summary.diff, 0);
    // Error policy: ambiguous rows do NOT appear in results
    assert!(
        result.results.iter().all(|r| r.status != RowStatus::Ambiguous),
        "Error policy should not emit Ambiguous rows in results"
    );
}

#[test]
fn combo_contains_digits_ambiguity_report() {
    // Same as above but Report policy → Ambiguous row appears in results
    let hdrs = headers();
    let t = KeyTransform::Digits;
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Contains,
        key_transform: t,
        on_ambiguous: AmbiguityPolicy::Report,
        tolerance: 0.01,
    };
    let left = vec![make_row_transformed("12", t, "$100.00", "", "")];
    let right = vec![
        make_row_transformed("A-123", t, "100.01", "", ""),
        make_row_transformed("B-125", t, "99.99", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.ambiguous, 1);
    assert_eq!(result.summary.only_right, 2);
    // Report policy: ambiguous row IS in results
    let ambiguous_in_results = result
        .results
        .iter()
        .filter(|r| r.status == RowStatus::Ambiguous)
        .count();
    assert_eq!(
        ambiguous_in_results, 1,
        "Report policy should emit 1 Ambiguous row in results"
    );
}

#[test]
fn combo_contains_trim_tolerance() {
    // Contains + Trim + tolerance=1.0
    // Left "  INV  " → trim → "INV"
    // Right " PREFIX-INV-SUFFIX " → trim → "PREFIX-INV-SUFFIX"
    // "PREFIX-INV-SUFFIX".contains("INV") → single match
    // Use exact f64 delta (1.0) to avoid floating-point boundary issues.
    let hdrs = headers();
    let t = KeyTransform::Trim;
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Contains,
        key_transform: t,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 1.0,
    };
    let left = vec![make_row_transformed("  INV  ", t, "$101.00", "", "")];
    let right = vec![make_row_transformed(" PREFIX-INV-SUFFIX ", t, "100.00", "", "")];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    // delta=1.0, within tolerance (1.0 <= 1.0) → Diff but within_tolerance
    assert_eq!(result.summary.diff, 1);
    assert_eq!(result.summary.diff_outside_tolerance, 0);
    assert_eq!(result.summary.matched, 0);
    assert_eq!(result.summary.only_left, 0);
    assert_eq!(result.summary.only_right, 0);
}

#[test]
fn combo_exact_compare_cols_tolerance() {
    // Exact + compare_cols=[1] (amount only) + tolerance=1.0
    // Verifies that non-compared columns don't affect status or diff counts.
    // Use exact f64 deltas to avoid floating-point boundary issues.
    let hdrs = headers();
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]), // only "amount"
        match_mode: MatchMode::Exact,
        key_transform: KeyTransform::None,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 1.0,
    };
    let left = vec![
        // K1: amount diff within tolerance, label/qty differ (but not compared)
        make_row("K1", "$101.00", "foo", "1"),
        // K2: amount diff outside tolerance
        make_row("K2", "$105.00", "x", "1"),
    ];
    let right = vec![
        make_row("K1", "100.00", "bar", "999"),
        make_row("K2", "100.00", "x", "1"),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    // K1: 1 diff (amount, delta=1.0, within tolerance). Label/qty not compared.
    // K2: 1 diff (amount, delta=5.0, outside tolerance).
    assert_eq!(result.summary.diff, 2);
    assert_eq!(result.summary.diff_outside_tolerance, 1, "only K2 outside tolerance");
    assert_eq!(result.summary.matched, 0);

    // Verify K1 has exactly 1 diff entry (amount, not label or qty)
    let k1 = result.results.iter().find(|r| r.key == "K1").unwrap();
    assert_eq!(k1.diffs.len(), 1, "K1 should have 1 diff (amount only)");
    assert_eq!(k1.diffs[0].column, "amount");
    assert!(k1.diffs[0].within_tolerance);
}

#[test]
fn combo_exact_tolerance_boundary() {
    // delta <= tolerance is within (inclusive). Verify the boundary.
    let hdrs = headers();
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Exact,
        key_transform: KeyTransform::None,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 0.25,
    };
    let left = vec![
        make_row("K1", "100.50", "", ""), // delta=0.25 → exactly at boundary
        make_row("K2", "100.50", "", ""), // delta=0.50 → outside
        make_row("K3", "100.50", "", ""), // delta=0.0 → Matched
    ];
    let right = vec![
        make_row("K1", "100.25", "", ""),
        make_row("K2", "100.00", "", ""),
        make_row("K3", "100.50", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.matched, 1, "K3 identical → Matched");
    assert_eq!(result.summary.diff, 2, "K1 + K2 → Diff");
    assert_eq!(
        result.summary.diff_outside_tolerance, 1,
        "K2 outside, K1 at boundary (within)"
    );

    let k1 = result.results.iter().find(|r| r.key == "K1").unwrap();
    assert!(k1.diffs[0].within_tolerance, "delta=0.25 at tol=0.25 must be within");

    let k2 = result.results.iter().find(|r| r.key == "K2").unwrap();
    assert!(!k2.diffs[0].within_tolerance, "delta=0.50 at tol=0.25 must be outside");
}

#[test]
fn combo_mixed_type_with_tolerance() {
    // Tolerance only applies when both sides are numeric.
    // Mixed type or both non-numeric → string diff, delta=None, within_tolerance=false.
    let hdrs = headers();
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Exact,
        key_transform: KeyTransform::None,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 1.0, // generous, but shouldn't matter for non-numeric
    };
    let left = vec![
        make_row("K1", "1234.56", "", ""), // numeric vs non-numeric
        make_row("K2", "N/A", "", ""),     // non-numeric vs non-numeric
        make_row("K3", "100.00", "", ""),  // identical → Matched
    ];
    let right = vec![
        make_row("K1", "N/A", "", ""),
        make_row("K2", "pending", "", ""),
        make_row("K3", "100.00", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.matched, 1, "K3 identical");
    assert_eq!(result.summary.diff, 2, "K1 + K2");
    assert_eq!(
        result.summary.diff_outside_tolerance, 2,
        "non-numeric diffs are always outside tolerance"
    );

    // Verify delta=None for both non-numeric diffs
    let k1 = result.results.iter().find(|r| r.key == "K1").unwrap();
    assert!(k1.diffs[0].delta.is_none(), "mixed type → delta=None");
    assert!(!k1.diffs[0].within_tolerance, "mixed type → outside tolerance");

    let k2 = result.results.iter().find(|r| r.key == "K2").unwrap();
    assert!(k2.diffs[0].delta.is_none(), "both non-numeric → delta=None");
    assert!(!k2.diffs[0].within_tolerance, "both non-numeric → outside tolerance");
}

// ===========================================================================
// Phase 4: IEEE-754 boundary tolerance
// ===========================================================================
//
// Verifies that tolerance comparison uses epsilon-inclusive semantics:
// values whose decimal difference equals the tolerance are classified as
// within tolerance, even when IEEE-754 representation makes the f64 delta
// slightly larger than the f64 tolerance.

#[test]
fn tolerance_boundary_ieee754() {
    // --- Per-pair checks (within_tolerance on individual diffs) ---

    // These values produce f64 deltas slightly above 0.01 due to
    // IEEE-754 representation. Without epsilon, they'd be classified
    // as outside tolerance.
    assert_diff_numeric("100.50", "100.49", 0.01, true);
    assert_diff_numeric("$1,234.56", "1234.55", 0.01, true);
    assert_diff_numeric("999999.50", "999999.49", 0.01, true);
    assert_diff_numeric("1000000.01", "1000000.00", 0.01, true);

    // Clearly outside — epsilon must not forgive real differences.
    assert_diff_numeric("100.00", "100.02", 0.01, false);
    assert_diff_numeric("1000.00", "1001.10", 1.0, false);
}

#[test]
fn tolerance_boundary_ieee754_summary() {
    // Full reconcile: verify diff_outside_tolerance counts correctly
    // when boundary cases are present alongside clearly-outside cases.
    let hdrs = headers();
    let opts = DiffOptions {
        key_col: 0,
        compare_cols: Some(vec![1]),
        match_mode: MatchMode::Exact,
        key_transform: KeyTransform::None,
        on_ambiguous: AmbiguityPolicy::Error,
        tolerance: 0.01,
    };
    let left = vec![
        make_row("K1", "100.50", "", ""),      // boundary: delta ≈ 0.01 → within
        make_row("K2", "999999.50", "", ""),    // boundary at large magnitude → within
        make_row("K3", "100.00", "", ""),       // clearly outside: delta = 0.02
        make_row("K4", "50.00", "", ""),        // identical → Matched
    ];
    let right = vec![
        make_row("K1", "100.49", "", ""),
        make_row("K2", "999999.49", "", ""),
        make_row("K3", "100.02", "", ""),
        make_row("K4", "50.00", "", ""),
    ];
    let result = reconcile(&left, &right, &hdrs, &opts).unwrap();

    assert_eq!(result.summary.matched, 1, "K4 identical");
    assert_eq!(result.summary.diff, 3, "K1 + K2 + K3 are diffs");
    assert_eq!(
        result.summary.diff_outside_tolerance, 1,
        "only K3 is outside tolerance; K1 and K2 are boundary-within"
    );

    // Verify K1 boundary: Diff status, within_tolerance=true
    let k1 = result.results.iter().find(|r| r.key == "K1").unwrap();
    assert_eq!(k1.status, RowStatus::Diff);
    assert!(k1.diffs[0].within_tolerance, "K1 boundary must be within tolerance");

    // Verify K3 clearly outside
    let k3 = result.results.iter().find(|r| r.key == "K3").unwrap();
    assert_eq!(k3.status, RowStatus::Diff);
    assert!(!k3.diffs[0].within_tolerance, "K3 must be outside tolerance");
}

// ===========================================================================
// Contract guard
// ===========================================================================

#[test]
fn contract_version_guard() {
    let bin = env!("CARGO_BIN_EXE_vgrid");
    let base = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/cli/diff/contract-json-schema"
    );
    let output = std::process::Command::new(bin)
        .args([
            "diff",
            &format!("{base}/left.csv"),
            &format!("{base}/right.csv"),
            "--key",
            "id",
            "--out",
            "json",
            "--summary",
            "none",
        ])
        .output()
        .expect("failed to run vgrid");
    // diff exits 1 when there are material diffs — that's expected
    assert!(
        output.status.code() == Some(0) || output.status.code() == Some(1),
        "unexpected exit code: {:?}",
        output.status.code()
    );
    let json: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("invalid JSON output");
    assert_eq!(
        json["contract_version"], 1,
        "contract_version must be 1"
    );
    // Top-level keys: contract_version, results, summary
    let obj = json.as_object().unwrap();
    assert!(
        obj.contains_key("contract_version"),
        "contract_version key must exist"
    );
    assert!(obj.contains_key("summary"), "summary key must exist");
    assert!(obj.contains_key("results"), "results key must exist");
}
