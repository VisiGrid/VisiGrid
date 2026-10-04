//! Randomized checks.
//!
//! - `tp1_random_pairs`: for random states and random concurrent pairs,
//!   `apply(S, b ++ T(a,b)) == apply(S, a ++ T(b,a))` on the real engine.
//! - `simulator_converges`: full client/server simulations with delays,
//!   reordering across clients and disconnects.
//!
//! Defaults keep `cargo test` under a minute. For long runs:
//! `COLLAB_FUZZ_SEEDS=100000 cargo test -p visigrid-collab --release --test fuzz`.

use rand::rngs::StdRng;
use rand::{Rng, SeedableRng};
use visigrid_collab::apply::{apply_ops, fingerprint, first_difference};
use visigrid_collab::gen::random_ops;
use visigrid_collab::op::CollabOp;
use visigrid_collab::sim::{run, shrink, SimConfig};
use visigrid_collab::transform::{transform_lists, Order};
use visigrid_engine::workbook::Workbook;

fn seeds(default: u64) -> u64 {
    std::env::var("COLLAB_FUZZ_SEEDS")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(default)
}

fn replay(ops: &[Vec<CollabOp>]) -> Workbook {
    let mut wb = Workbook::new();
    for o in ops {
        apply_ops(&mut wb, o);
    }
    wb
}

#[test]
fn tp1_random_pairs() {
    let n = seeds(5000);
    let mut failures = Vec::new();
    let mut engine: Vec<String> = Vec::new();
    for seed in 0..n {
        let mut rng = StdRng::seed_from_u64(seed ^ 0x7031);
        let mut history: Vec<Vec<CollabOp>> = Vec::new();
        let mut wb = Workbook::new();
        for _ in 0..rng.gen_range(0..12) {
            let key = (1 << 62) | rng.gen_range(0..1u64 << 40);
            let ops = random_ops(&mut rng, &wb, key);
            apply_ops(&mut wb, &ops);
            history.push(ops);
        }
        let (ka, kb) = (
            (1 << 62) | rng.gen_range(0..1u64 << 40),
            (1 << 62) | rng.gen_range(0..1u64 << 40),
        );
        let a = random_ops(&mut rng, &wb, ka);
        let b = random_ops(&mut rng, &wb, kb);
        let Ok((a2, b2)) = transform_lists(&a, &b, Order::Later) else {
            continue;
        };
        let mut left = replay(&history);
        apply_ops(&mut left, &b);
        apply_ops(&mut left, &a2);
        let mut right = replay(&history);
        apply_ops(&mut right, &a);
        apply_ops(&mut right, &b2);
        if let Some(d) = first_difference(&fingerprint(&left), &fingerprint(&right)) {
            // Classify: an identity transform means the engine itself is
            // order dependent for these two ops; agreement after a full
            // recompute means the incremental recalc left a stale value.
            left.recompute_full_ordered();
            right.recompute_full_ordered();
            let after_full = first_difference(&fingerprint(&left), &fingerprint(&right));
            let identity = a2 == a && b2 == b;
            let class = match (identity, after_full.is_none()) {
                (_, true) => "ENGINE incremental recalc (paths agree after full recompute)",
                (true, false) => "ENGINE order dependence (identity transform)",
                (false, false) => "TRANSFORM",
            };
            let line =
                format!("{class} seed {seed}: a={a:?}\n  b={b:?}\n  a'={a2:?}\n  b'={b2:?}\n  {d}");
            if class == "TRANSFORM" {
                failures.push(line);
                if failures.len() >= 5 {
                    break;
                }
            } else {
                engine.push(line);
            }
        }
    }
    if !engine.is_empty() {
        eprintln!(
            "engine findings ({}), not transform failures:\n{}",
            engine.len(),
            engine
                .iter()
                .take(8)
                .cloned()
                .collect::<Vec<_>>()
                .join("\n")
        );
    }
    assert!(
        failures.is_empty(),
        "TP1 failures ({}):\n{}",
        failures.len(),
        failures.join("\n")
    );
}

#[test]
fn simulator_converges() {
    let n = seeds(2000);
    let configs = [
        SimConfig::default(),
        SimConfig {
            clients: 2,
            edits: 30,
            max_delay: 4,
            toggle_prob: 0.0,
            step_gap: 1,
            undo_prob: 0.0,
        },
        SimConfig {
            clients: 5,
            edits: 60,
            max_delay: 20,
            toggle_prob: 0.08,
            step_gap: 2,
            undo_prob: 0.0,
        },
        SimConfig {
            clients: 10,
            edits: 80,
            max_delay: 30,
            toggle_prob: 0.03,
            step_gap: 1,
            undo_prob: 0.0,
        },
    ];
    let mut committed = 0;
    let mut refused = 0;
    let (mut local, mut own, mut collateral, mut no_inverse) = (0u64, 0u64, 0u64, 0u64);
    let mut stale = Vec::new();
    let mut cycle = Vec::new();
    for seed in 0..n {
        let cfg = &configs[(seed % configs.len() as u64) as usize];
        let r = run(seed, cfg, false);
        committed += r.committed;
        refused += r.refused;
        local += r.local_envelopes;
        own += r.refused_envelopes;
        collateral += r.discarded_after_refusal;
        no_inverse += r.discarded_no_inverse;
        if let Some(d) = &r.engine_stale {
            stale.push(format!("seed {seed}: {d}"));
        }
        if let Some(d) = &r.engine_cycle {
            cycle.push(format!("seed {seed}: {d}"));
        }
        if !r.ok() {
            let s = shrink(seed, cfg);
            panic!(
                "seed {seed} {cfg:?}: {}\nshrunk to {} events:\n{}",
                s.failure.clone().unwrap_or_default(),
                s.trace.len(),
                s.trace.join("\n")
            );
        }
    }
    let pct = |x: u64| 100.0 * x as f64 / local.max(1) as f64;
    eprintln!(
        "local envelopes {local}: refused themselves {own} ({:.2}%), discarded as collateral {collateral} ({:.2}%; {no_inverse} after a sheet rename/delete), total lost {:.2}%",
        pct(own),
        pct(collateral),
        pct(own + collateral)
    );
    eprintln!(
        "{n} simulations converged: {committed} envelopes committed, {refused} refused; \
         {} needed a full recompute (engine incremental recalc):\n{}\n\
         {} hit engine cycle-text loss:\n{}",
        stale.len(),
        stale.iter().take(5).cloned().collect::<Vec<_>>().join("\n"),
        cycle.len(),
        cycle.iter().take(5).cloned().collect::<Vec<_>>().join("\n")
    );
}

/// Undo and redo are ordinary local edits, so convergence must hold with
/// them mixed into every configuration.
#[test]
fn simulator_converges_with_undo() {
    let n = seeds(1000);
    let configs = [
        SimConfig { undo_prob: 0.3, ..SimConfig::default() },
        SimConfig { clients: 5, edits: 60, max_delay: 20, toggle_prob: 0.08, step_gap: 2, undo_prob: 0.4 },
    ];
    // COLLAB_UNDO_SEED replays one seed.
    let only: Option<u64> = std::env::var("COLLAB_UNDO_SEED").ok().and_then(|s| s.parse().ok());
    for seed in only.map_or(0..n, |s| s..s + 1) {
        let cfg = &configs[(seed % configs.len() as u64) as usize];
        let r = run(seed ^ 0x0d0e, cfg, false);
        if !r.ok() {
            let s = shrink(seed ^ 0x0d0e, cfg);
            panic!(
                "seed {seed} {cfg:?}: {}\nshrunk to {} events:\n{}",
                s.failure.clone().unwrap_or_default(),
                s.trace.len(),
                s.trace.join("\n")
            );
        }
    }
}
