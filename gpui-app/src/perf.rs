//! Navigation latency instrumentation
//!
//! Records timestamps at key points in the keystroke-to-paint pipeline:
//! - t0: key action handler entry (KeyDown received)
//! - t1: selection state updated
//! - t2: render() called (frame start)
//!
//! Use `NavLatencyTracker::report()` to get p50/p95 stats and events-per-frame count.
//! Enabled via `VISIGRID_PERF=nav` environment variable.

use std::time::Instant;
use std::sync::OnceLock;

/// Check if nav perf tracking is enabled via VISIGRID_PERF=nav env var.
static NAV_PERF_ENABLED: OnceLock<bool> = OnceLock::new();

pub fn is_nav_perf_enabled() -> bool {
    *NAV_PERF_ENABLED.get_or_init(|| {
        std::env::var("VISIGRID_PERF").ok().as_deref() == Some("nav")
    })
}

const RING_SIZE: usize = 128;

/// A completed latency sample (key action → render).
struct LatencySample {
    /// Time from key action to render start (the number users feel).
    key_to_render_us: u64,
    /// Time from key action to state update (input handling cost).
    key_to_state_us: u64,
    /// Number of key events coalesced into this render frame.
    events_this_frame: u32,
    /// Number of cell moves applied in this render frame.
    moves_this_frame: u32,
}

/// Tracks navigation keystroke-to-paint latency.
///
/// Call sites:
/// - `mark_key_action()` — arrow key action handler entry
/// - `mark_state_updated()` — after move_selection completes
/// - `mark_render()` — top of Render::render()
pub struct NavLatencyTracker {
    /// Ring buffer of completed samples.
    samples: Vec<LatencySample>,
    write_idx: usize,

    /// In-flight tracking for current frame.
    frame_key_time: Option<Instant>,
    frame_state_time: Option<Instant>,
    frame_event_count: u32,
    frame_moves: u32,

    /// Lifetime counters.
    total_frames: u64,
    total_events: u64,
    total_coalesced: u64,
    total_moves: u64,
}

impl Default for NavLatencyTracker {
    fn default() -> Self {
        Self {
            samples: Vec::new(),
            write_idx: 0,
            frame_key_time: None,
            frame_state_time: None,
            frame_event_count: 0,
            frame_moves: 0,
            total_frames: 0,
            total_events: 0,
            total_coalesced: 0,
            total_moves: 0,
        }
    }
}

impl NavLatencyTracker {
    /// Called at the start of an arrow key action handler.
    /// Records the first key timestamp per frame (subsequent events in the
    /// same frame increment the event counter but don't overwrite t0).
    pub fn mark_key_action(&mut self) {
        if !is_nav_perf_enabled() { return; }
        if self.frame_key_time.is_none() {
            self.frame_key_time = Some(Instant::now());
        }
        self.frame_event_count += 1;
        self.total_events += 1;
    }

    /// Called after selection state is updated (move_selection complete).
    pub fn mark_state_updated(&mut self) {
        if !is_nav_perf_enabled() { return; }
        self.frame_state_time = Some(Instant::now());
    }

    /// Called at the top of Render::render(). Flushes the in-flight sample
    /// and resets for the next frame.
    pub fn mark_render(&mut self) {
        if !is_nav_perf_enabled() { return; }

        if let Some(key_time) = self.frame_key_time.take() {
            let now = Instant::now();
            let key_to_render = now.duration_since(key_time).as_micros() as u64;
            let key_to_state = self.frame_state_time
                .map(|t| t.duration_since(key_time).as_micros() as u64)
                .unwrap_or(0);

            let sample = LatencySample {
                key_to_render_us: key_to_render,
                key_to_state_us: key_to_state,
                events_this_frame: self.frame_event_count,
                moves_this_frame: self.frame_moves,
            };

            if self.samples.len() < RING_SIZE {
                self.samples.push(sample);
            } else {
                self.samples[self.write_idx] = sample;
            }
            self.write_idx = (self.write_idx + 1) % RING_SIZE;
            self.total_frames += 1;

            if self.frame_event_count > 1 {
                self.total_coalesced += (self.frame_event_count - 1) as u64;
            }
        }

        // Reset for next frame
        self.frame_state_time = None;
        self.frame_event_count = 0;
        self.frame_moves = 0;
    }

    /// Called after flushing batched navigation moves.
    /// `count` is the number of move_selection calls applied this frame.
    pub fn mark_moves_applied(&mut self, count: u32) {
        if !is_nav_perf_enabled() || count == 0 { return; }
        self.frame_moves += count;
        self.total_moves += count as u64;
    }

    /// Generate a human-readable report of latency stats.
    /// Returns None if not enough samples or perf tracking is disabled.
    pub fn report(&self) -> Option<String> {
        if !is_nav_perf_enabled() || self.samples.is_empty() {
            return None;
        }

        let n = self.samples.len();

        // Collect key-to-render latencies
        let mut k2r: Vec<u64> = self.samples.iter().map(|s| s.key_to_render_us).collect();
        k2r.sort_unstable();

        // Collect key-to-state latencies
        let mut k2s: Vec<u64> = self.samples.iter().map(|s| s.key_to_state_us).collect();
        k2s.sort_unstable();

        // Collect moves-per-frame (only frames that had moves)
        let mut mpf: Vec<u32> = self.samples.iter()
            .map(|s| s.moves_this_frame)
            .filter(|&m| m > 0)
            .collect();
        mpf.sort_unstable();

        let p50 = |v: &[u64]| v[v.len() / 2];
        let p95 = |v: &[u64]| v[(v.len() as f64 * 0.95) as usize];

        // Format as ms with one decimal place
        let us_to_ms = |us: u64| -> String {
            if us < 1000 {
                format!("{}µs", us)
            } else {
                format!("{:.1}ms", us as f64 / 1000.0)
            }
        };

        let moves_str = if mpf.is_empty() {
            "0".to_string()
        } else {
            let avg = self.total_moves as f64 / self.total_frames.max(1) as f64;
            format!("{:.1}/f ({}tot)", avg, self.total_moves)
        };

        Some(format!(
            "Nav p50={} p95={} | state={} | moves={} | {} frames, {} coalesced",
            us_to_ms(p50(&k2r)), us_to_ms(p95(&k2r)),
            us_to_ms(p50(&k2s)),
            moves_str,
            self.total_frames, self.total_coalesced,
        ))
    }
}
