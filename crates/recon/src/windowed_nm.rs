use std::collections::HashMap;

use chrono::NaiveDate;

use crate::config::{ToleranceConfig, WindowedNmConfig};
use crate::model::{Aggregate, AmbiguityReason, MatchProof, MatchedPair, PairMatchOutput, ReconRow};

/// Compute the ambiguity reason from the combination of tied solutions and cap hit.
fn ambiguity_reason(num_equivalent: usize, cap_hit: bool) -> Option<AmbiguityReason> {
    let tied = num_equivalent > 1;
    match (tied, cap_hit) {
        (true, true) => Some(AmbiguityReason::TiedAndCapHit),
        (true, false) => Some(AmbiguityReason::TiedSolutions),
        (false, true) => Some(AmbiguityReason::SearchCapHit),
        (false, false) => None,
    }
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

pub fn match_windowed_nm(
    left_rows: &[ReconRow],
    right_rows: &[ReconRow],
    tolerance: &ToleranceConfig,
    config: &WindowedNmConfig,
) -> PairMatchOutput {
    let mut all_matched = Vec::new();
    let mut all_left_only = Vec::new();
    let mut all_right_only = Vec::new();

    // Partition by currency
    let mut left_by_cur: HashMap<&str, Vec<&ReconRow>> = HashMap::new();
    let mut right_by_cur: HashMap<&str, Vec<&ReconRow>> = HashMap::new();
    for r in left_rows {
        left_by_cur.entry(&r.currency).or_default().push(r);
    }
    for r in right_rows {
        right_by_cur.entry(&r.currency).or_default().push(r);
    }

    // Collect all currencies
    let mut currencies: Vec<&str> = left_by_cur.keys().chain(right_by_cur.keys()).copied().collect();
    currencies.sort_unstable();
    currencies.dedup();

    let mut group_counter: usize = 0;

    for currency in currencies {
        let left = left_by_cur.get(currency).map(|v| v.as_slice()).unwrap_or(&[]);
        let right = right_by_cur.get(currency).map(|v| v.as_slice()).unwrap_or(&[]);

        let buckets = build_buckets(left, right, tolerance.date_window_days, currency);

        for bucket in &buckets {
            if bucket.left.len() + bucket.right.len() > config.max_bucket_size {
                // Bucket too large — produce ambiguous matches with BucketTooLarge proof.
                // If both sides have rows, pair them as one oversized ambiguous group.
                // If only one side, they're truly unmatched.
                if !bucket.left.is_empty() && !bucket.right.is_empty() {
                    let left_agg = rows_to_aggregate(&bucket.left, &mut group_counter);
                    let right_agg = rows_to_aggregate(&bucket.right, &mut group_counter);
                    let left_sum: i64 = bucket.left.iter().map(|r| r.amount_cents).sum();
                    let right_sum: i64 = bucket.right.iter().map(|r| r.amount_cents).sum();
                    let delta = left_sum - right_sum;
                    let date_off = compute_date_offset(&bucket.left, &bucket.right);

                    let proof = MatchProof {
                        strategy: "windowed_nm".into(),
                        pass: "bucket_too_large".into(),
                        bucket_id: bucket.bucket_id(),
                        nodes_visited: 0,
                        nodes_pruned: 0,
                        cap_hit: false,
                        ambiguous: true,
                        num_equivalent_solutions: 0,
                        ambiguity_reason: Some(AmbiguityReason::BucketTooLarge),
                        tie_break_reason: Some(format!(
                            "bucket_size={} exceeds max={}",
                            bucket.left.len() + bucket.right.len(),
                            config.max_bucket_size
                        )),
                    };
                    all_matched.push(MatchedPair {
                        left: left_agg,
                        right: right_agg,
                        delta_cents: delta,
                        date_offset_days: date_off,
                        within_tolerance: delta.abs() <= tolerance.amount_cents,
                        within_window: date_off.unsigned_abs() <= tolerance.date_window_days,
                        proof: Some(proof),
                    });
                } else {
                    for r in &bucket.left {
                        all_left_only.push(row_to_aggregate(r, &mut group_counter));
                    }
                    for r in &bucket.right {
                        all_right_only.push(row_to_aggregate(r, &mut group_counter));
                    }
                }
                continue;
            }

            let result = solve_bucket(bucket, tolerance, config, &mut group_counter);
            all_matched.extend(result.matched);
            all_left_only.extend(result.left_only);
            all_right_only.extend(result.right_only);
        }
    }

    PairMatchOutput {
        matched: all_matched,
        left_only: all_left_only,
        right_only: all_right_only,
    }
}

// ---------------------------------------------------------------------------
// Bucket types
// ---------------------------------------------------------------------------

#[derive(Debug)]
struct Bucket<'a> {
    left: Vec<&'a ReconRow>,
    right: Vec<&'a ReconRow>,
    currency: String,
    window_start: NaiveDate,
    window_end: NaiveDate,
}

impl<'a> Bucket<'a> {
    fn bucket_id(&self) -> String {
        format!("{}:{}..{}", self.currency, self.window_start, self.window_end)
    }
}

// ---------------------------------------------------------------------------
// Bucket building — sliding window over sorted timeline
// ---------------------------------------------------------------------------

fn build_buckets<'a>(
    left: &[&'a ReconRow],
    right: &[&'a ReconRow],
    date_window_days: u32,
    currency: &str,
) -> Vec<Bucket<'a>> {
    // Merge all rows with side tag, sort by (date, amount, record_id)
    #[derive(Clone, Copy)]
    enum Side {
        Left,
        Right,
    }
    let mut all: Vec<(&ReconRow, Side)> = Vec::with_capacity(left.len() + right.len());
    for r in left {
        all.push((r, Side::Left));
    }
    for r in right {
        all.push((r, Side::Right));
    }

    all.sort_by(|a, b| {
        a.0.date
            .cmp(&b.0.date)
            .then_with(|| a.0.amount_cents.cmp(&b.0.amount_cents))
            .then_with(|| a.0.record_id.cmp(&b.0.record_id))
    });

    if all.is_empty() {
        return vec![];
    }

    // Sliding window: group rows whose dates fall within date_window_days of each other.
    // We use a greedy approach: start a new bucket at the first unassigned row,
    // include all rows within date_window_days of the first row's date.
    let window = chrono::Duration::days(date_window_days as i64);
    let mut buckets = Vec::new();
    let mut assigned = vec![false; all.len()];

    for i in 0..all.len() {
        if assigned[i] {
            continue;
        }

        let anchor_date = all[i].0.date;
        let window_end = anchor_date + window;

        let mut bucket_left = Vec::new();
        let mut bucket_right = Vec::new();
        let mut actual_start = anchor_date;
        let mut actual_end = anchor_date;

        for j in i..all.len() {
            if assigned[j] {
                continue;
            }
            if all[j].0.date > window_end {
                break;
            }
            assigned[j] = true;
            if all[j].0.date < actual_start {
                actual_start = all[j].0.date;
            }
            if all[j].0.date > actual_end {
                actual_end = all[j].0.date;
            }
            match all[j].1 {
                Side::Left => bucket_left.push(all[j].0),
                Side::Right => bucket_right.push(all[j].0),
            }
        }

        if !bucket_left.is_empty() || !bucket_right.is_empty() {
            buckets.push(Bucket {
                left: bucket_left,
                right: bucket_right,
                currency: currency.to_string(),
                window_start: actual_start,
                window_end: actual_end,
            });
        }
    }

    buckets
}

// ---------------------------------------------------------------------------
// Solver — 4-pass per bucket
// ---------------------------------------------------------------------------

struct SolveResult {
    matched: Vec<MatchedPair>,
    left_only: Vec<Aggregate>,
    right_only: Vec<Aggregate>,
}

fn solve_bucket(
    bucket: &Bucket<'_>,
    tolerance: &ToleranceConfig,
    config: &WindowedNmConfig,
    group_counter: &mut usize,
) -> SolveResult {
    let bucket_id = bucket.bucket_id();

    // Sort rows within bucket by (date, amount, record_id) for determinism
    let mut left: Vec<&ReconRow> = bucket.left.clone();
    let mut right: Vec<&ReconRow> = bucket.right.clone();
    sort_rows(&mut left);
    sort_rows(&mut right);

    let effective_max_group = if config.allow_mixed_sign {
        config.max_group_size.min(4)
    } else {
        config.max_group_size
    };

    let mut matched = Vec::new();
    let mut total_nodes: u64 = 0;
    let mut total_pruned: u64 = 0;
    let mut any_cap_hit = false;

    // ----- Pass 1: Exact 1:1 -----
    let mut left_used = vec![false; left.len()];
    let mut right_used = vec![false; right.len()];

    for (li, lr) in left.iter().enumerate() {
        if left_used[li] {
            continue;
        }
        let mut best: Option<(usize, i64, i32)> = None;
        let mut best_score: i64 = i64::MAX;
        let mut num_equivalent: usize = 0;

        for (ri, rr) in right.iter().enumerate() {
            if right_used[ri] {
                continue;
            }
            let delta = lr.amount_cents - rr.amount_cents;
            let date_off = (lr.date - rr.date).num_days() as i32;

            if delta.abs() <= tolerance.amount_cents {
                let score = delta.abs() * 1000 + date_off.unsigned_abs() as i64;
                if score < best_score {
                    best = Some((ri, delta, date_off));
                    best_score = score;
                    num_equivalent = 1;
                } else if score == best_score {
                    num_equivalent += 1;
                }
            }
        }
        if let Some((ri, delta, date_off)) = best {
            left_used[li] = true;
            right_used[ri] = true;
            let ambiguous = num_equivalent > 1;
            let reason = ambiguity_reason(num_equivalent, false);
            let proof = MatchProof {
                strategy: "windowed_nm".into(),
                pass: "exact_1_1".into(),
                bucket_id: bucket_id.clone(),
                nodes_visited: 1,
                nodes_pruned: 0,
                cap_hit: false,
                ambiguous,
                num_equivalent_solutions: num_equivalent,
                ambiguity_reason: reason,
                tie_break_reason: if ambiguous {
                    Some("record_id_order".into())
                } else {
                    None
                },
            };
            matched.push(make_matched_pair(
                &[*lr],
                &[right[ri]],
                delta,
                date_off,
                tolerance,
                group_counter,
                proof,
            ));
        }
    }

    // Collect remaining rows
    let mut rem_left: Vec<&ReconRow> = left
        .iter()
        .enumerate()
        .filter(|(i, _)| !left_used[*i])
        .map(|(_, r)| *r)
        .collect();
    let mut rem_right: Vec<&ReconRow> = right
        .iter()
        .enumerate()
        .filter(|(i, _)| !right_used[*i])
        .map(|(_, r)| *r)
        .collect();

    // ----- Pass 2: k:1 (subset of left sums to one right) -----
    if !rem_left.is_empty() && !rem_right.is_empty() {
        let mut right_matched = vec![false; rem_right.len()];

        // Sort right by amount descending for deterministic processing
        let mut right_indices: Vec<usize> = (0..rem_right.len()).collect();
        right_indices.sort_by(|&a, &b| {
            rem_right[b]
                .amount_cents
                .abs()
                .cmp(&rem_right[a].amount_cents.abs())
                .then_with(|| rem_right[a].record_id.cmp(&rem_right[b].record_id))
        });

        let mut left_consumed = vec![false; rem_left.len()];

        for &ri in &right_indices {
            if right_matched[ri] {
                continue;
            }
            let target = rem_right[ri].amount_cents;
            let available: Vec<usize> = (0..rem_left.len())
                .filter(|&i| !left_consumed[i])
                .filter(|&i| {
                    if !config.allow_mixed_sign {
                        same_sign(rem_left[i].amount_cents, target)
                    } else {
                        true
                    }
                })
                .collect();

            if available.is_empty() {
                continue;
            }

            let search = subset_sum_search(
                &available.iter().map(|&i| rem_left[i].amount_cents).collect::<Vec<_>>(),
                target,
                tolerance.amount_cents,
                effective_max_group,
                config.max_nodes,
            );
            total_nodes += search.nodes_visited;
            total_pruned += search.nodes_pruned;
            if search.cap_hit {
                any_cap_hit = true;
            }

            if !search.tied_best.is_empty() {
                // Re-rank using full score tuple
                let avail_rows: Vec<&ReconRow> = available.iter().map(|&i| rem_left[i]).collect();
                let target_rows: Vec<&ReconRow> = vec![rem_right[ri]];
                let (best_indices, num_equiv) = pick_best_solution(
                    &search.tied_best,
                    &avail_rows,
                    &target_rows,
                    &config.evidence_fields,
                );
                // Map back to rem_left indices
                let chosen: Vec<usize> = best_indices.iter().map(|&si| available[si]).collect();
                if chosen.len() >= 2 {
                    let left_group: Vec<&ReconRow> =
                        chosen.iter().map(|&i| rem_left[i]).collect();
                    let delta: i64 = left_group.iter().map(|r| r.amount_cents).sum::<i64>() - target;
                    let date_off = compute_date_offset(&left_group, &[rem_right[ri]]);

                    let ambiguous = num_equiv > 1 || search.cap_hit;
                    let reason = ambiguity_reason(num_equiv, search.cap_hit);
                    let proof = MatchProof {
                        strategy: "windowed_nm".into(),
                        pass: "k_1".into(),
                        bucket_id: bucket_id.clone(),
                        nodes_visited: search.nodes_visited,
                        nodes_pruned: search.nodes_pruned,
                        cap_hit: search.cap_hit,
                        ambiguous,
                        num_equivalent_solutions: num_equiv,
                        ambiguity_reason: reason,
                        tie_break_reason: if ambiguous {
                            Some("full_score_tuple".into())
                        } else {
                            None
                        },
                    };
                    matched.push(make_matched_pair(
                        &left_group,
                        &[rem_right[ri]],
                        delta,
                        date_off,
                        tolerance,
                        group_counter,
                        proof,
                    ));

                    for &ci in &chosen {
                        left_consumed[ci] = true;
                    }
                    right_matched[ri] = true;
                }
            }
        }

        // Update remaining
        rem_left = rem_left
            .into_iter()
            .enumerate()
            .filter(|(i, _)| !left_consumed[*i])
            .map(|(_, r)| r)
            .collect();
        rem_right = rem_right
            .into_iter()
            .enumerate()
            .filter(|(i, _)| !right_matched[*i])
            .map(|(_, r)| r)
            .collect();
    }

    // ----- Pass 3: 1:k (one left matches subset of right) -----
    if !rem_left.is_empty() && !rem_right.is_empty() {
        let mut left_matched = vec![false; rem_left.len()];

        let mut left_indices: Vec<usize> = (0..rem_left.len()).collect();
        left_indices.sort_by(|&a, &b| {
            rem_left[b]
                .amount_cents
                .abs()
                .cmp(&rem_left[a].amount_cents.abs())
                .then_with(|| rem_left[a].record_id.cmp(&rem_left[b].record_id))
        });

        let mut right_consumed = vec![false; rem_right.len()];

        for &li in &left_indices {
            if left_matched[li] {
                continue;
            }
            let target = rem_left[li].amount_cents;
            let available: Vec<usize> = (0..rem_right.len())
                .filter(|&i| !right_consumed[i])
                .filter(|&i| {
                    if !config.allow_mixed_sign {
                        same_sign(rem_right[i].amount_cents, target)
                    } else {
                        true
                    }
                })
                .collect();

            if available.is_empty() {
                continue;
            }

            let search = subset_sum_search(
                &available.iter().map(|&i| rem_right[i].amount_cents).collect::<Vec<_>>(),
                target,
                tolerance.amount_cents,
                effective_max_group,
                config.max_nodes,
            );
            total_nodes += search.nodes_visited;
            total_pruned += search.nodes_pruned;
            if search.cap_hit {
                any_cap_hit = true;
            }

            if !search.tied_best.is_empty() {
                let avail_rows: Vec<&ReconRow> = available.iter().map(|&i| rem_right[i]).collect();
                let target_rows: Vec<&ReconRow> = vec![rem_left[li]];
                let (best_indices, num_equiv) = pick_best_solution(
                    &search.tied_best,
                    &avail_rows,
                    &target_rows,
                    &config.evidence_fields,
                );
                let chosen: Vec<usize> = best_indices.iter().map(|&si| available[si]).collect();
                if chosen.len() >= 2 {
                    let right_group: Vec<&ReconRow> =
                        chosen.iter().map(|&i| rem_right[i]).collect();
                    let delta: i64 =
                        target - right_group.iter().map(|r| r.amount_cents).sum::<i64>();
                    let date_off = compute_date_offset(&[rem_left[li]], &right_group);

                    let ambiguous = num_equiv > 1 || search.cap_hit;
                    let reason = ambiguity_reason(num_equiv, search.cap_hit);
                    let proof = MatchProof {
                        strategy: "windowed_nm".into(),
                        pass: "1_k".into(),
                        bucket_id: bucket_id.clone(),
                        nodes_visited: search.nodes_visited,
                        nodes_pruned: search.nodes_pruned,
                        cap_hit: search.cap_hit,
                        ambiguous,
                        num_equivalent_solutions: num_equiv,
                        ambiguity_reason: reason,
                        tie_break_reason: if ambiguous {
                            Some("full_score_tuple".into())
                        } else {
                            None
                        },
                    };
                    matched.push(make_matched_pair(
                        &[rem_left[li]],
                        &right_group,
                        delta,
                        date_off,
                        tolerance,
                        group_counter,
                        proof,
                    ));

                    left_matched[li] = true;
                    for &ci in &chosen {
                        right_consumed[ci] = true;
                    }
                }
            }
        }

        rem_left = rem_left
            .into_iter()
            .enumerate()
            .filter(|(i, _)| !left_matched[*i])
            .map(|(_, r)| r)
            .collect();
        rem_right = rem_right
            .into_iter()
            .enumerate()
            .filter(|(i, _)| !right_consumed[*i])
            .map(|(_, r)| r)
            .collect();
    }

    // ----- Pass 4: k:k (bounded DFS on small remainders) -----
    if !rem_left.is_empty()
        && !rem_right.is_empty()
        && rem_left.len() + rem_right.len() <= effective_max_group * 2
    {
        let kk = kk_search(
            &rem_left,
            &rem_right,
            tolerance,
            effective_max_group,
            config.max_nodes,
        );
        total_nodes += kk.nodes_visited;
        total_pruned += kk.nodes_pruned;
        if kk.cap_hit {
            any_cap_hit = true;
        }

        if !kk.tied_best.is_empty() {
            // Re-rank tied solutions using full score tuple
            let (left_idx, right_idx, num_equiv) = pick_best_kk_solution(
                &kk.tied_best,
                &rem_left,
                &rem_right,
                &config.evidence_fields,
            );

            let left_group: Vec<&ReconRow> = left_idx.iter().map(|&i| rem_left[i]).collect();
            let right_group: Vec<&ReconRow> = right_idx.iter().map(|&i| rem_right[i]).collect();
            let left_sum: i64 = left_group.iter().map(|r| r.amount_cents).sum();
            let right_sum: i64 = right_group.iter().map(|r| r.amount_cents).sum();
            let delta = left_sum - right_sum;
            let date_off = compute_date_offset(&left_group, &right_group);

            let ambiguous = num_equiv > 1 || kk.cap_hit;
            let reason = ambiguity_reason(num_equiv, kk.cap_hit);
            let proof = MatchProof {
                strategy: "windowed_nm".into(),
                pass: "k_k".into(),
                bucket_id: bucket_id.clone(),
                nodes_visited: kk.nodes_visited,
                nodes_pruned: kk.nodes_pruned,
                cap_hit: kk.cap_hit,
                ambiguous,
                num_equivalent_solutions: num_equiv,
                ambiguity_reason: reason,
                tie_break_reason: if ambiguous {
                    Some("full_score_tuple".into())
                } else {
                    None
                },
            };
            matched.push(make_matched_pair(
                &left_group,
                &right_group,
                delta,
                date_off,
                tolerance,
                group_counter,
                proof,
            ));

            // Remove matched from remaining
            let left_set: Vec<bool> = (0..rem_left.len())
                .map(|i| left_idx.contains(&i))
                .collect();
            let right_set: Vec<bool> = (0..rem_right.len())
                .map(|i| right_idx.contains(&i))
                .collect();
            rem_left = rem_left
                .into_iter()
                .enumerate()
                .filter(|(i, _)| !left_set[*i])
                .map(|(_, r)| r)
                .collect();
            rem_right = rem_right
                .into_iter()
                .enumerate()
                .filter(|(i, _)| !right_set[*i])
                .map(|(_, r)| r)
                .collect();
        }
    }

    // Anything still remaining is unmatched
    let left_only: Vec<Aggregate> = rem_left
        .iter()
        .map(|r| row_to_aggregate(r, group_counter))
        .collect();
    let right_only: Vec<Aggregate> = rem_right
        .iter()
        .map(|r| row_to_aggregate(r, group_counter))
        .collect();

    let _ = (total_nodes, total_pruned, any_cap_hit); // used in proofs above

    SolveResult {
        matched,
        left_only,
        right_only,
    }
}

// ---------------------------------------------------------------------------
// Subset sum search (bounded DFS)
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Tie-break score tuple
// ---------------------------------------------------------------------------

/// Full tie-break score. Lower is better (lexicographic comparison).
/// Two solutions are "equivalent" iff they match on all fields except lex_ids.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct SolutionScore {
    total_records: usize,
    date_span_days: i64,
    sum_date_distance: i64,
    neg_evidence_score: i64, // negated so lower = more evidence
    lex_ids: String,
}

impl SolutionScore {
    /// Score a candidate subset of rows against a target group.
    fn from_rows(
        candidate_rows: &[&ReconRow],
        target_rows: &[&ReconRow],
        evidence_fields: &[String],
    ) -> Self {
        let total_records = candidate_rows.len() + target_rows.len();

        let all_dates: Vec<NaiveDate> = candidate_rows
            .iter()
            .chain(target_rows.iter())
            .map(|r| r.date)
            .collect();
        let date_span_days = if all_dates.len() >= 2 {
            let min = *all_dates.iter().min().unwrap();
            let max = *all_dates.iter().max().unwrap();
            (max - min).num_days()
        } else {
            0
        };

        // Sum of absolute date distances from the group centroid
        let sum_date_distance = if !all_dates.is_empty() {
            let min_date = *all_dates.iter().min().unwrap();
            all_dates.iter().map(|d| (*d - min_date).num_days()).sum()
        } else {
            0
        };

        let evidence_score = compute_evidence_score(candidate_rows, target_rows, evidence_fields);

        let mut ids: Vec<&str> = candidate_rows
            .iter()
            .chain(target_rows.iter())
            .map(|r| r.record_id.as_str())
            .collect();
        ids.sort_unstable();
        let lex_ids = ids.join(",");

        SolutionScore {
            total_records,
            date_span_days,
            sum_date_distance,
            neg_evidence_score: -(evidence_score as i64),
            lex_ids,
        }
    }

    /// Two scores are "equivalent" for ambiguity purposes if everything except
    /// lex_ids matches. lex_ids is the final deterministic tiebreak — if we need
    /// it to distinguish solutions, they're operationally ambiguous.
    fn equivalent_ignoring_ids(&self, other: &Self) -> bool {
        self.total_records == other.total_records
            && self.date_span_days == other.date_span_days
            && self.sum_date_distance == other.sum_date_distance
            && self.neg_evidence_score == other.neg_evidence_score
    }
}

/// Count matching tokens in evidence_fields between two row groups.
fn compute_evidence_score(
    left: &[&ReconRow],
    right: &[&ReconRow],
    evidence_fields: &[String],
) -> usize {
    if evidence_fields.is_empty() {
        return 0;
    }
    let mut score = 0;
    for field in evidence_fields {
        let left_tokens: Vec<&str> = left
            .iter()
            .filter_map(|r| r.raw_fields.get(field).map(|s| s.as_str()))
            .filter(|s| !s.is_empty())
            .collect();
        let right_tokens: Vec<&str> = right
            .iter()
            .filter_map(|r| r.raw_fields.get(field).map(|s| s.as_str()))
            .filter(|s| !s.is_empty())
            .collect();
        for lt in &left_tokens {
            for rt in &right_tokens {
                if lt == rt {
                    score += 1;
                }
            }
        }
    }
    score
}

/// Given tied-best subsets from the DFS (index sets into `candidates`),
/// re-rank using the full score tuple and return the best + ambiguity info.
/// `candidates` are the rows the subsets index into.
/// `targets` are the rows on the other side of the match.
fn pick_best_solution<'a>(
    tied: &[Vec<usize>],
    candidates: &[&'a ReconRow],
    targets: &[&'a ReconRow],
    evidence_fields: &[String],
) -> (Vec<usize>, usize) {
    if tied.len() <= 1 {
        return (tied.first().cloned().unwrap_or_default(), tied.len());
    }

    let mut scored: Vec<(SolutionScore, &Vec<usize>)> = tied
        .iter()
        .map(|indices| {
            let rows: Vec<&ReconRow> = indices.iter().map(|&i| candidates[i]).collect();
            let score = SolutionScore::from_rows(&rows, targets, evidence_fields);
            (score, indices)
        })
        .collect();
    scored.sort_by(|a, b| a.0.cmp(&b.0));

    let best_score = &scored[0].0;
    let num_equivalent = scored
        .iter()
        .filter(|(s, _)| best_score.equivalent_ignoring_ids(s))
        .count();

    (scored[0].1.clone(), num_equivalent)
}

/// Re-rank tied k:k solutions using full score tuple.
/// Returns (best_left_indices, best_right_indices, num_equivalent).
fn pick_best_kk_solution(
    tied: &[(Vec<usize>, Vec<usize>)],
    left_rows: &[&ReconRow],
    right_rows: &[&ReconRow],
    evidence_fields: &[String],
) -> (Vec<usize>, Vec<usize>, usize) {
    if tied.len() <= 1 {
        let (li, ri) = tied.first().cloned().unwrap_or_default();
        return (li, ri, tied.len());
    }

    let mut scored: Vec<(SolutionScore, usize)> = tied
        .iter()
        .enumerate()
        .map(|(idx, (li, ri))| {
            let l: Vec<&ReconRow> = li.iter().map(|&i| left_rows[i]).collect();
            let r: Vec<&ReconRow> = ri.iter().map(|&i| right_rows[i]).collect();
            let score = SolutionScore::from_rows(&l, &r, evidence_fields);
            (score, idx)
        })
        .collect();
    scored.sort_by(|a, b| a.0.cmp(&b.0));

    let best_score = &scored[0].0;
    let num_equivalent = scored
        .iter()
        .filter(|(s, _)| best_score.equivalent_ignoring_ids(s))
        .count();

    let best_idx = scored[0].1;
    let (li, ri) = tied[best_idx].clone();
    (li, ri, num_equivalent)
}

/// Maximum number of tied-best solutions to collect for caller re-ranking.
const MAX_TIED_SOLUTIONS: usize = 16;

struct SearchResult {
    /// All solutions tied at the best (delta, size) level, up to MAX_TIED_SOLUTIONS.
    /// The caller re-ranks these using full context (dates, evidence).
    tied_best: Vec<Vec<usize>>,
    /// Total number of solutions at the best level (may exceed tied_best.len()).
    /// Used for diagnostics; callers re-count via pick_best_solution.
    #[allow(dead_code)]
    num_equivalent_solutions: usize,
    nodes_visited: u64,
    nodes_pruned: u64,
    cap_hit: bool,
}

impl SearchResult {
    fn best(&self) -> Option<&Vec<usize>> {
        self.tied_best.first()
    }
}

fn subset_sum_search(
    amounts: &[i64],
    target: i64,
    tolerance: i64,
    max_group_size: usize,
    max_nodes: usize,
) -> SearchResult {
    let mut tied_best: Vec<Vec<usize>> = Vec::new();
    let mut best_delta: i64 = i64::MAX;
    let mut best_len: usize = usize::MAX;
    let mut num_equivalent: usize = 0;
    let mut nodes_visited: u64 = 0;
    let mut nodes_pruned: u64 = 0;
    let mut cap_hit = false;

    let mut stack: Vec<usize> = Vec::new();

    fn dfs(
        amounts: &[i64],
        target: i64,
        tolerance: i64,
        max_group_size: usize,
        max_nodes: usize,
        start: usize,
        current_sum: i64,
        stack: &mut Vec<usize>,
        tied_best: &mut Vec<Vec<usize>>,
        best_delta: &mut i64,
        best_len: &mut usize,
        num_equivalent: &mut usize,
        nodes_visited: &mut u64,
        nodes_pruned: &mut u64,
        cap_hit: &mut bool,
    ) {
        if *cap_hit {
            return;
        }

        *nodes_visited += 1;
        if *nodes_visited >= max_nodes as u64 {
            *cap_hit = true;
            return;
        }

        let delta = (current_sum - target).abs();
        if delta <= tolerance && !stack.is_empty() {
            if delta < *best_delta
                || (delta == *best_delta && stack.len() < *best_len)
            {
                // Strictly better — replace all tied solutions
                tied_best.clear();
                tied_best.push(stack.clone());
                *best_delta = delta;
                *best_len = stack.len();
                *num_equivalent = 1;
            } else if delta == *best_delta && stack.len() == *best_len {
                // Equivalent — collect for caller re-ranking
                *num_equivalent += 1;
                if tied_best.len() < MAX_TIED_SOLUTIONS {
                    tied_best.push(stack.clone());
                }
            }
        }

        if stack.len() >= max_group_size {
            *nodes_pruned += 1;
            return;
        }

        for i in start..amounts.len() {
            stack.push(i);
            dfs(
                amounts,
                target,
                tolerance,
                max_group_size,
                max_nodes,
                i + 1,
                current_sum + amounts[i],
                stack,
                tied_best,
                best_delta,
                best_len,
                num_equivalent,
                nodes_visited,
                nodes_pruned,
                cap_hit,
            );
            stack.pop();

            if *cap_hit {
                return;
            }
        }
    }

    dfs(
        amounts,
        target,
        tolerance,
        max_group_size,
        max_nodes,
        0,
        0,
        &mut stack,
        &mut tied_best,
        &mut best_delta,
        &mut best_len,
        &mut num_equivalent,
        &mut nodes_visited,
        &mut nodes_pruned,
        &mut cap_hit,
    );

    SearchResult {
        tied_best,
        num_equivalent_solutions: num_equivalent,
        nodes_visited,
        nodes_pruned,
        cap_hit,
    }
}

// ---------------------------------------------------------------------------
// k:k search (bounded DFS over partition pairs)
// ---------------------------------------------------------------------------

struct KkResult {
    /// Tied-best (left_indices, right_indices) pairs for caller re-ranking.
    tied_best: Vec<(Vec<usize>, Vec<usize>)>,
    nodes_visited: u64,
    nodes_pruned: u64,
    cap_hit: bool,
    /// Used for diagnostics; callers re-count via pick_best_kk_solution.
    #[allow(dead_code)]
    num_equivalent_solutions: usize,
}

fn kk_search(
    left: &[&ReconRow],
    right: &[&ReconRow],
    tolerance: &ToleranceConfig,
    max_group_size: usize,
    max_nodes: usize,
) -> KkResult {
    let mut tied_best: Vec<(Vec<usize>, Vec<usize>)> = Vec::new();
    let mut best_score: (usize, i64) = (usize::MAX, i64::MAX); // (total_count, delta)
    let mut num_equivalent: usize = 0;
    let mut nodes_visited: u64 = 0;
    let mut nodes_pruned: u64 = 0;
    let mut cap_hit = false;

    // Enumerate subsets of left (up to max_group_size)
    let max_left = left.len().min(max_group_size);
    let max_right = right.len().min(max_group_size);

    for left_size in 2..=max_left {
        if cap_hit {
            break;
        }
        for left_combo in combinations(left.len(), left_size) {
            if cap_hit {
                break;
            }
            nodes_visited += 1;
            if nodes_visited >= max_nodes as u64 {
                cap_hit = true;
                break;
            }

            let left_sum: i64 = left_combo.iter().map(|&i| left[i].amount_cents).sum();

            // For this left subset, find a right subset that sums close
            let right_amounts: Vec<i64> = right.iter().map(|r| r.amount_cents).collect();
            let sub = subset_sum_search(
                &right_amounts,
                left_sum,
                tolerance.amount_cents,
                max_right,
                (max_nodes as u64 - nodes_visited) as usize,
            );
            nodes_visited += sub.nodes_visited;
            nodes_pruned += sub.nodes_pruned;
            if sub.cap_hit {
                cap_hit = true;
            }

            if let Some(right_combo) = sub.best() {
                let right_combo = right_combo.clone();
                if right_combo.len() >= 2 || left_combo.len() >= 2 {
                    let right_sum: i64 = right_combo.iter().map(|&i| right[i].amount_cents).sum();
                    let delta = (left_sum - right_sum).abs();
                    let total_count = left_combo.len() + right_combo.len();
                    let score = (total_count, delta);

                    if score < best_score {
                        best_score = score;
                        tied_best.clear();
                        tied_best.push((left_combo.clone(), right_combo));
                        num_equivalent = 1;
                    } else if score == best_score {
                        num_equivalent += 1;
                        if tied_best.len() < MAX_TIED_SOLUTIONS {
                            tied_best.push((left_combo.clone(), right_combo));
                        }
                    }
                }
            }
        }
    }

    KkResult {
        tied_best,
        nodes_visited,
        nodes_pruned,
        cap_hit,
        num_equivalent_solutions: num_equivalent,
    }
}

/// Generate all combinations of `k` items from `0..n`.
fn combinations(n: usize, k: usize) -> Vec<Vec<usize>> {
    let mut result = Vec::new();
    let mut combo = Vec::with_capacity(k);

    fn gen(start: usize, n: usize, k: usize, combo: &mut Vec<usize>, result: &mut Vec<Vec<usize>>) {
        if combo.len() == k {
            result.push(combo.clone());
            return;
        }
        for i in start..n {
            combo.push(i);
            gen(i + 1, n, k, combo, result);
            combo.pop();
        }
    }

    gen(0, n, k, &mut combo, &mut result);
    result
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn sort_rows(rows: &mut [&ReconRow]) {
    rows.sort_by(|a, b| {
        a.date
            .cmp(&b.date)
            .then_with(|| a.amount_cents.cmp(&b.amount_cents))
            .then_with(|| a.record_id.cmp(&b.record_id))
    });
}

fn same_sign(a: i64, b: i64) -> bool {
    (a >= 0 && b >= 0) || (a < 0 && b < 0)
}

fn compute_date_offset(left: &[&ReconRow], right: &[&ReconRow]) -> i32 {
    let left_date = left.iter().map(|r| r.date).min().unwrap_or(NaiveDate::MIN);
    let right_date = right.iter().map(|r| r.date).min().unwrap_or(NaiveDate::MIN);
    (left_date - right_date).num_days() as i32
}

fn row_to_aggregate(row: &ReconRow, counter: &mut usize) -> Aggregate {
    let id = format!("wnm_{}", *counter);
    *counter += 1;
    Aggregate {
        role: row.role.clone(),
        match_key: id,
        currency: row.currency.clone(),
        date: row.date,
        total_cents: row.amount_cents,
        record_count: 1,
        record_ids: vec![row.record_id.clone()],
    }
}

fn rows_to_aggregate(rows: &[&ReconRow], counter: &mut usize) -> Aggregate {
    let id = format!("wnm_{}", *counter);
    *counter += 1;
    let total: i64 = rows.iter().map(|r| r.amount_cents).sum();
    let date = rows.iter().map(|r| r.date).min().unwrap_or(NaiveDate::MIN);
    let role = rows.first().map(|r| r.role.as_str()).unwrap_or("unknown");
    let record_ids: Vec<String> = rows.iter().map(|r| r.record_id.clone()).collect();
    Aggregate {
        role: role.to_string(),
        match_key: id,
        currency: rows.first().map(|r| r.currency.clone()).unwrap_or_default(),
        date,
        total_cents: total,
        record_count: rows.len(),
        record_ids,
    }
}

fn make_matched_pair(
    left_rows: &[&ReconRow],
    right_rows: &[&ReconRow],
    delta_cents: i64,
    date_offset_days: i32,
    tolerance: &ToleranceConfig,
    group_counter: &mut usize,
    proof: MatchProof,
) -> MatchedPair {
    let left_agg = rows_to_aggregate(left_rows, group_counter);
    let right_agg = rows_to_aggregate(right_rows, group_counter);

    MatchedPair {
        left: left_agg,
        right: right_agg,
        delta_cents,
        date_offset_days,
        within_tolerance: delta_cents.abs() <= tolerance.amount_cents,
        within_window: date_offset_days.unsigned_abs() <= tolerance.date_window_days,
        proof: Some(proof),
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    fn row(role: &str, id: &str, amount: i64, date: &str, currency: &str) -> ReconRow {
        ReconRow {
            role: role.into(),
            record_id: id.into(),
            match_key: id.into(),
            date: chrono::NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
            amount_cents: amount,
            currency: currency.into(),
            kind: "payment".into(),
            raw_fields: HashMap::new(),
        }
    }

    fn default_tol() -> ToleranceConfig {
        ToleranceConfig {
            amount_cents: 0,
            date_window_days: 3,
        }
    }

    fn default_config() -> WindowedNmConfig {
        WindowedNmConfig::default()
    }

    // 1. exact_1_1_passthrough
    #[test]
    fn exact_1_1_passthrough() {
        let left = vec![row("proc", "L1", 10000, "2026-01-15", "USD")];
        let right = vec![row("bank", "R1", 10000, "2026-01-15", "USD")];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);
        assert_eq!(out.matched[0].delta_cents, 0);
        let proof = out.matched[0].proof.as_ref().unwrap();
        assert_eq!(proof.pass, "exact_1_1");
    }

    // 2. merge_2_1
    #[test]
    fn merge_2_1() {
        let left = vec![
            row("proc", "L1", 3000, "2026-01-15", "USD"),
            row("proc", "L2", 7000, "2026-01-15", "USD"),
        ];
        let right = vec![row("bank", "R1", 10000, "2026-01-15", "USD")];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);
        assert_eq!(out.matched[0].left.record_count, 2);
        assert_eq!(out.matched[0].right.record_count, 1);
        let proof = out.matched[0].proof.as_ref().unwrap();
        assert_eq!(proof.pass, "k_1");
    }

    // 3. split_1_3
    #[test]
    fn split_1_3() {
        let left = vec![row("proc", "L1", 15000, "2026-01-15", "USD")];
        let right = vec![
            row("bank", "R1", 5000, "2026-01-15", "USD"),
            row("bank", "R2", 5000, "2026-01-16", "USD"),
            row("bank", "R3", 5000, "2026-01-16", "USD"),
        ];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);
        assert_eq!(out.matched[0].left.record_count, 1);
        assert_eq!(out.matched[0].right.record_count, 3);
        let proof = out.matched[0].proof.as_ref().unwrap();
        assert_eq!(proof.pass, "1_k");
    }

    // 4. true_3_2
    #[test]
    fn true_3_2() {
        let left = vec![
            row("proc", "L1", 2000, "2026-01-15", "USD"),
            row("proc", "L2", 3000, "2026-01-15", "USD"),
            row("proc", "L3", 5000, "2026-01-16", "USD"),
        ];
        let right = vec![
            row("bank", "R1", 4000, "2026-01-15", "USD"),
            row("bank", "R2", 6000, "2026-01-16", "USD"),
        ];
        // 2000+3000+5000 = 10000 = 4000+6000
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);
        let proof = out.matched[0].proof.as_ref().unwrap();
        assert_eq!(proof.pass, "k_k");
    }

    // 5. cross_currency_isolation
    #[test]
    fn cross_currency_isolation() {
        let left = vec![
            row("proc", "L1", 10000, "2026-01-15", "USD"),
            row("proc", "L2", 5000, "2026-01-15", "EUR"),
        ];
        let right = vec![
            row("bank", "R1", 10000, "2026-01-15", "EUR"),
            row("bank", "R2", 5000, "2026-01-15", "USD"),
        ];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        // USD: L1(10000) matches R2(5000)? No — different amounts.
        // EUR: L2(5000) matches R1(10000)? No — different amounts.
        // Actually: USD L1=10000 vs R2=5000 won't match. EUR L2=5000 vs R1=10000 won't match.
        assert_eq!(out.matched.len(), 0);
        assert_eq!(out.left_only.len(), 2);
        assert_eq!(out.right_only.len(), 2);
    }

    // 6. date_window_enforcement
    #[test]
    fn date_window_enforcement() {
        let left = vec![row("proc", "L1", 10000, "2026-01-10", "USD")];
        let right = vec![row("bank", "R1", 10000, "2026-01-20", "USD")];
        // Window is 3 days, dates are 10 days apart — they'll be in different buckets
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched.len(), 0);
        assert_eq!(out.left_only.len(), 1);
        assert_eq!(out.right_only.len(), 1);
    }

    // 7. tolerance_acceptance
    #[test]
    fn tolerance_acceptance() {
        let left = vec![row("proc", "L1", 10000, "2026-01-15", "USD")];
        let right = vec![row("bank", "R1", 10030, "2026-01-15", "USD")];
        let tol = ToleranceConfig {
            amount_cents: 50,
            date_window_days: 3,
        };
        let out = match_windowed_nm(&left, &right, &tol, &default_config());
        assert_eq!(out.matched.len(), 1);
        assert!(out.matched[0].within_tolerance);
        assert_eq!(out.matched[0].delta_cents, -30);
    }

    // 8. max_group_size_prune
    #[test]
    fn max_group_size_prune() {
        // 7 left rows that sum to 1 right, but max_group_size=3
        let left: Vec<ReconRow> = (1..=7)
            .map(|i| row("proc", &format!("L{i}"), 1000, "2026-01-15", "USD"))
            .collect();
        let right = vec![row("bank", "R1", 7000, "2026-01-15", "USD")];
        let mut cfg = default_config();
        cfg.max_group_size = 3;
        let out = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        // Can't form a group of 7, and no subset of ≤3 sums to 7000, so no match
        // (3*1000=3000 ≠ 7000)
        assert_eq!(out.matched.len(), 0);
    }

    // 9. deterministic_tiebreak
    #[test]
    fn deterministic_tiebreak() {
        // Two identical-amount right rows, one left
        let left = vec![row("proc", "L1", 5000, "2026-01-15", "USD")];
        let right = vec![
            row("bank", "R1", 5000, "2026-01-15", "USD"),
            row("bank", "R2", 5000, "2026-01-15", "USD"),
        ];
        let out1 = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        let out2 = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out1.matched.len(), 1);
        assert_eq!(out2.matched.len(), 1);
        // Must pick the same one every time
        assert_eq!(
            out1.matched[0].right.record_ids,
            out2.matched[0].right.record_ids
        );
    }

    // 10. collision_identical_amounts
    #[test]
    fn collision_identical_amounts() {
        let left = vec![
            row("proc", "L1", 5000, "2026-01-15", "USD"),
            row("proc", "L2", 5000, "2026-01-15", "USD"),
        ];
        let right = vec![
            row("bank", "R1", 5000, "2026-01-15", "USD"),
            row("bank", "R2", 5000, "2026-01-15", "USD"),
        ];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        // Should match 2 pairs deterministically
        assert_eq!(out.matched.len(), 2);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);

        // Run again — same result
        let out2 = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched[0].left.record_ids, out2.matched[0].left.record_ids);
        assert_eq!(out.matched[1].left.record_ids, out2.matched[1].left.record_ids);
    }

    // 11. cap_behavior_max_nodes
    #[test]
    fn cap_behavior_max_nodes() {
        // Create a scenario with many rows but very low max_nodes
        let left: Vec<ReconRow> = (1..=10)
            .map(|i| row("proc", &format!("L{i}"), i * 1000, "2026-01-15", "USD"))
            .collect();
        let right: Vec<ReconRow> = (1..=10)
            .map(|i| row("bank", &format!("R{i}"), i * 1000, "2026-01-15", "USD"))
            .collect();
        let mut cfg = default_config();
        cfg.max_nodes = 10; // Very low cap

        let out = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        // Should return a stable result — some matches found before cap, rest unmatched
        // The key property: no panic, deterministic
        let out2 = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        assert_eq!(out.matched.len(), out2.matched.len());
    }

    // 12. cap_behavior_max_bucket
    #[test]
    fn cap_behavior_max_bucket() {
        // Create more rows than max_bucket_size
        let left: Vec<ReconRow> = (1..=30)
            .map(|i| row("proc", &format!("L{i}"), i * 100, "2026-01-15", "USD"))
            .collect();
        let right: Vec<ReconRow> = (1..=30)
            .map(|i| row("bank", &format!("R{i}"), i * 100, "2026-01-15", "USD"))
            .collect();
        let mut cfg = default_config();
        cfg.max_bucket_size = 10; // Very low — 60 rows >> 10

        let out = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        // Oversized bucket produces a single ambiguous match with BucketTooLarge proof
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);

        let proof = out.matched[0].proof.as_ref().unwrap();
        assert!(proof.ambiguous);
        assert_eq!(proof.pass, "bucket_too_large");
        assert_eq!(
            proof.ambiguity_reason,
            Some(crate::model::AmbiguityReason::BucketTooLarge)
        );
        assert!(
            proof.tie_break_reason.as_ref().unwrap().contains("exceeds max"),
            "tie_break_reason should explain the bucket size"
        );

        // The ambiguous match should aggregate all rows
        assert_eq!(out.matched[0].left.record_count, 30);
        assert_eq!(out.matched[0].right.record_count, 30);
    }

    // -----------------------------------------------------------------------
    // Proof minimum viable contract
    // -----------------------------------------------------------------------

    // 13. proof_contract — every WindowedNM match has a complete proof
    #[test]
    fn proof_contract() {
        // Mix of passes: 1:1, k:1, 1:k
        let left = vec![
            row("proc", "L1", 5000, "2026-01-15", "USD"),  // 1:1
            row("proc", "L2", 3000, "2026-01-15", "USD"),  // k:1 part
            row("proc", "L3", 7000, "2026-01-16", "USD"),  // k:1 part
        ];
        let right = vec![
            row("bank", "R1", 5000, "2026-01-15", "USD"),  // 1:1
            row("bank", "R2", 10000, "2026-01-16", "USD"), // k:1 target
        ];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());

        for m in &out.matched {
            let proof = m.proof.as_ref().expect("every windowed_nm match must have proof");

            // Strategy must be identified
            assert_eq!(proof.strategy, "windowed_nm");

            // Pass must be one of the valid phases
            assert!(
                ["exact_1_1", "k_1", "1_k", "k_k"].contains(&proof.pass.as_str()),
                "unexpected pass: {}",
                proof.pass
            );

            // Bucket ID must be present and parseable
            assert!(proof.bucket_id.contains(':'), "bucket_id must contain currency separator");

            // nodes_visited must be > 0 (we always visit at least one node)
            assert!(proof.nodes_visited > 0, "nodes_visited must be positive");

            // cap_hit must be explicitly recorded (not just defaulted)
            // This is a bool so it's always present — just verify we can read it
            let _ = proof.cap_hit;

            // ambiguous must be explicitly recorded
            let _ = proof.ambiguous;
            let _ = proof.num_equivalent_solutions;

            // If ambiguous, both reason and tie_break must explain why
            if proof.ambiguous {
                assert!(
                    proof.ambiguity_reason.is_some(),
                    "ambiguous proof must have ambiguity_reason"
                );
                assert!(
                    proof.tie_break_reason.is_some(),
                    "ambiguous proof must have tie_break_reason"
                );
            } else {
                assert!(
                    proof.ambiguity_reason.is_none(),
                    "non-ambiguous proof must not have ambiguity_reason"
                );
            }
        }
    }

    // -----------------------------------------------------------------------
    // Regression fixture 1: Same-amount collision day
    // -----------------------------------------------------------------------

    // 14. Two deposits same amount, two possible groupings, no evidence fields.
    // Must either match deterministically AND flag as ambiguous, or refuse to match.
    #[test]
    fn regression_same_amount_collision_day() {
        // Two processor payouts of $10,000 each
        // Two bank deposits of $10,000 each, same day
        // Any 1:1 pairing is equally valid — engine must flag ambiguity.
        let left = vec![
            row("proc", "P1", 10000, "2026-01-15", "USD"),
            row("proc", "P2", 10000, "2026-01-15", "USD"),
        ];
        let right = vec![
            row("bank", "D1", 10000, "2026-01-15", "USD"),
            row("bank", "D2", 10000, "2026-01-15", "USD"),
        ];
        let out = match_windowed_nm(&left, &right, &default_tol(), &default_config());

        // Must match both
        assert_eq!(out.matched.len(), 2);
        assert_eq!(out.left_only.len(), 0);
        assert_eq!(out.right_only.len(), 0);

        // At least the first match must be flagged as ambiguous
        // (the second match has only 1 candidate remaining, so it's not ambiguous)
        let proof0 = out.matched[0].proof.as_ref().unwrap();
        assert!(
            proof0.ambiguous,
            "first match in same-amount collision must be ambiguous"
        );
        assert!(
            proof0.num_equivalent_solutions > 1,
            "must record multiple equivalent solutions"
        );
        assert_eq!(
            proof0.ambiguity_reason,
            Some(crate::model::AmbiguityReason::TiedSolutions),
            "reason must be TiedSolutions (not cap hit)"
        );

        // Second match: only one candidate left, not ambiguous
        let proof1 = out.matched[1].proof.as_ref().unwrap();
        assert!(
            !proof1.ambiguous,
            "second match should not be ambiguous (only one candidate left)"
        );
        assert_eq!(
            proof1.ambiguity_reason, None,
            "non-ambiguous match must have no reason"
        );

        // Deterministic: same result on re-run
        let out2 = match_windowed_nm(&left, &right, &default_tol(), &default_config());
        assert_eq!(out.matched[0].right.record_ids, out2.matched[0].right.record_ids);
        assert_eq!(out.matched[1].right.record_ids, out2.matched[1].right.record_ids);
    }

    // -----------------------------------------------------------------------
    // Evidence scoring breaks ties
    // -----------------------------------------------------------------------

    // 16. Two k:1 solutions identical on (delta, size, date_span) but
    //     evidence_fields picks the one with matching merchant_id.
    #[test]
    fn evidence_fields_break_tie() {
        fn row_with_fields(
            role: &str,
            id: &str,
            amount: i64,
            date: &str,
            currency: &str,
            fields: Vec<(&str, &str)>,
        ) -> ReconRow {
            let mut raw_fields = HashMap::new();
            for (k, v) in fields {
                raw_fields.insert(k.to_string(), v.to_string());
            }
            ReconRow {
                role: role.into(),
                record_id: id.into(),
                match_key: id.into(),
                date: chrono::NaiveDate::parse_from_str(date, "%Y-%m-%d").unwrap(),
                amount_cents: amount,
                currency: currency.into(),
                kind: "payment".into(),
                raw_fields,
            }
        }

        // Two processor payments of 5000 each, same date
        // One bank deposit of 10000
        // Both L1+L2 and L1+L3 sum to 10000 (L3 has same amount as L2)
        // But L1 and L2 share merchant_id "ACME" with the bank row
        let left = vec![
            row_with_fields("proc", "L1", 5000, "2026-01-15", "USD", vec![("merchant_id", "ACME")]),
            row_with_fields("proc", "L2", 5000, "2026-01-15", "USD", vec![("merchant_id", "ACME")]),
            row_with_fields("proc", "L3", 5000, "2026-01-15", "USD", vec![("merchant_id", "OTHER")]),
        ];
        let right = vec![
            row_with_fields("bank", "R1", 10000, "2026-01-15", "USD", vec![("merchant_id", "ACME")]),
        ];
        let mut cfg = default_config();
        cfg.evidence_fields = vec!["merchant_id".into()];

        let out = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        assert_eq!(out.matched.len(), 1);
        assert_eq!(out.matched[0].left.record_count, 2);

        // Evidence scoring should prefer L1+L2 (both match "ACME") over L1+L3 or L2+L3
        let matched_ids = &out.matched[0].left.record_ids;
        assert!(
            matched_ids.contains(&"L1".to_string()) && matched_ids.contains(&"L2".to_string()),
            "evidence fields should prefer L1+L2 (ACME match), got {:?}",
            matched_ids
        );

        // L3 should be unmatched
        assert_eq!(out.left_only.len(), 1);
        assert_eq!(out.left_only[0].record_ids, vec!["L3"]);
    }

    // -----------------------------------------------------------------------
    // Regression fixture 2: Cap-hit bucket
    // -----------------------------------------------------------------------

    // 15. Large bucket hits max_nodes → stable output, proof records cap hit.
    #[test]
    fn regression_cap_hit_bucket() {
        // 15 left rows, 5 right rows in same window.
        // With max_nodes=50 the k:1 / 1:k / k:k passes will hit the cap.
        let left: Vec<ReconRow> = (1..=15)
            .map(|i| row("proc", &format!("P{i}"), 1000 + (i as i64) * 7, "2026-01-15", "USD"))
            .collect();
        let right: Vec<ReconRow> = vec![
            row("bank", "D1", 1007, "2026-01-15", "USD"),  // matches P1 exactly
            row("bank", "D2", 1014, "2026-01-15", "USD"),  // matches P2 exactly
            row("bank", "D3", 5000, "2026-01-16", "USD"),  // requires subset sum
            row("bank", "D4", 8000, "2026-01-16", "USD"),  // requires subset sum
            row("bank", "D5", 12000, "2026-01-16", "USD"), // requires subset sum
        ];
        let mut cfg = default_config();
        cfg.max_nodes = 50; // Low cap to force cap-hit

        let out1 = match_windowed_nm(&left, &right, &default_tol(), &cfg);
        let out2 = match_windowed_nm(&left, &right, &default_tol(), &cfg);

        // Stable: same number of matches across runs
        assert_eq!(out1.matched.len(), out2.matched.len());

        // At least the 1:1 exact matches should succeed (P1↔D1, P2↔D2)
        assert!(out1.matched.len() >= 2, "at least 2 exact 1:1 matches expected");

        // Any match that hit the cap must record it in proof
        for m in &out1.matched {
            let proof = m.proof.as_ref().unwrap();
            // If this match's search hit the cap, it must be flagged
            if proof.cap_hit {
                assert!(
                    proof.ambiguous,
                    "cap_hit matches must be flagged ambiguous"
                );
            }
        }

        // Stable record-level determinism
        for (m1, m2) in out1.matched.iter().zip(out2.matched.iter()) {
            assert_eq!(m1.left.record_ids, m2.left.record_ids);
            assert_eq!(m1.right.record_ids, m2.right.record_ids);
        }
    }

    // -----------------------------------------------------------------------
    // Determinism under randomized input order
    // -----------------------------------------------------------------------

    /// Canonical fingerprint of match output for comparison.
    /// Sorted by (left_ids, right_ids) to be order-independent.
    fn fingerprint(out: &PairMatchOutput) -> String {
        let mut match_strs: Vec<String> = out
            .matched
            .iter()
            .map(|m| {
                let mut l = m.left.record_ids.clone();
                l.sort();
                let mut r = m.right.record_ids.clone();
                r.sort();
                let proof_summary = m.proof.as_ref().map(|p| {
                    format!(
                        "{}:{}:ambig={}:equiv={}:cap={}",
                        p.strategy, p.pass, p.ambiguous, p.num_equivalent_solutions, p.cap_hit
                    )
                }).unwrap_or_default();
                format!(
                    "M[{:?}↔{:?} Δ={} d={} {}]",
                    l, r, m.delta_cents, m.date_offset_days, proof_summary
                )
            })
            .collect();
        match_strs.sort();

        let mut lo_strs: Vec<String> = out
            .left_only
            .iter()
            .map(|a| {
                let mut ids = a.record_ids.clone();
                ids.sort();
                format!("L{:?}", ids)
            })
            .collect();
        lo_strs.sort();

        let mut ro_strs: Vec<String> = out
            .right_only
            .iter()
            .map(|a| {
                let mut ids = a.record_ids.clone();
                ids.sort();
                format!("R{:?}", ids)
            })
            .collect();
        ro_strs.sort();

        format!("{:?}|{:?}|{:?}", match_strs, lo_strs, ro_strs)
    }

    /// Simple seeded shuffle (Fisher-Yates with xorshift64).
    fn shuffle_seeded<T>(data: &mut [T], seed: u64) {
        let mut s = seed;
        for i in (1..data.len()).rev() {
            // xorshift64
            s ^= s << 13;
            s ^= s >> 7;
            s ^= s << 17;
            let j = (s as usize) % (i + 1);
            data.swap(i, j);
        }
    }

    // 17. Randomized-input determinism: shuffled inputs produce identical output.
    #[test]
    fn determinism_randomized_input_order() {
        // Non-trivial fixture: mix of 1:1, k:1, and unmatched across 2 currencies
        let base_left = vec![
            row("proc", "L1", 10000, "2026-01-15", "USD"),
            row("proc", "L2", 3000, "2026-01-15", "USD"),
            row("proc", "L3", 7000, "2026-01-16", "USD"),
            row("proc", "L4", 5000, "2026-01-15", "EUR"),
            row("proc", "L5", 2000, "2026-01-17", "USD"),
            row("proc", "L6", 8000, "2026-01-16", "EUR"),
        ];
        let base_right = vec![
            row("bank", "R1", 10000, "2026-01-15", "USD"),  // 1:1 with L1
            row("bank", "R2", 10000, "2026-01-16", "USD"),  // k:1 with L2+L3
            row("bank", "R3", 5000, "2026-01-15", "EUR"),   // 1:1 with L4
            row("bank", "R4", 9000, "2026-01-17", "USD"),   // no match
            row("bank", "R5", 8000, "2026-01-16", "EUR"),   // 1:1 with L6
        ];

        let tol = ToleranceConfig { amount_cents: 0, date_window_days: 3 };
        let cfg = default_config();

        // Get reference fingerprint from sorted input
        let reference = fingerprint(&match_windowed_nm(&base_left, &base_right, &tol, &cfg));

        // Run with 10 different shuffles
        for seed in 1..=10u64 {
            let mut left = base_left.clone();
            let mut right = base_right.clone();
            shuffle_seeded(&mut left, seed);
            shuffle_seeded(&mut right, seed.wrapping_mul(31));

            let out = match_windowed_nm(&left, &right, &tol, &cfg);
            let fp = fingerprint(&out);
            assert_eq!(
                reference, fp,
                "determinism broken with seed={}: output differs from reference",
                seed
            );
        }
    }
}
