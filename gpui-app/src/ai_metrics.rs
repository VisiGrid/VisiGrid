//! Local AI feature usage metrics.
//!
//! Tracks how AI terminal features are used, stored in
//! `~/.config/visigrid/ai_metrics.json`. No network, no telemetry service.
//! This is purely for the user (and us) to understand what's working.
//!
//! Writes are debounced: updates accumulate in memory and flush to disk
//! at most every 5 seconds, or on process exit via the Drop guard.

use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use std::sync::Mutex;
use std::time::Instant;

/// Debounce interval: don't write more often than this.
const FLUSH_INTERVAL_SECS: u64 = 5;

struct MetricsState {
    data: AiMetrics,
    dirty: bool,
    last_write: Instant,
}

static METRICS: Mutex<Option<MetricsState>> = Mutex::new(None);

/// Local AI usage counters.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AiMetrics {
    /// Number of times "Launch AI" was invoked
    #[serde(default)]
    pub launch_ai: u64,
    /// Number of times "AI: Explain Selection" was invoked
    #[serde(default)]
    pub explain_selection: u64,
    /// Number of times any paste-to-terminal command was used
    #[serde(default)]
    pub paste_context: u64,
    /// Number of times "Generate AI Context Files" was used
    #[serde(default)]
    pub generate_context_files: u64,
    /// Last CLI binary that was successfully launched
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_cli: Option<String>,
    /// Cumulative error counts by category
    #[serde(default, skip_serializing_if = "std::collections::HashMap::is_empty")]
    pub errors: std::collections::HashMap<String, u64>,
}

/// Returns the metrics file path for display or opening.
pub fn metrics_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("ai_metrics.json")
}

fn load() -> AiMetrics {
    let path = metrics_path();
    if let Ok(data) = std::fs::read_to_string(&path) {
        serde_json::from_str(&data).unwrap_or_default()
    } else {
        AiMetrics::default()
    }
}

fn write_to_disk(metrics: &AiMetrics) {
    let path = metrics_path();
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(metrics) {
        let _ = std::fs::write(&path, json);
    }
}

fn with_state<F>(f: F)
where
    F: FnOnce(&mut MetricsState),
{
    let mut guard = METRICS.lock().unwrap_or_else(|e| e.into_inner());
    let state = guard.get_or_insert_with(|| MetricsState {
        data: load(),
        dirty: false,
        last_write: Instant::now(),
    });
    f(state);

    // Debounced flush: write only if dirty AND enough time has passed
    if state.dirty && state.last_write.elapsed().as_secs() >= FLUSH_INTERVAL_SECS {
        write_to_disk(&state.data);
        state.dirty = false;
        state.last_write = Instant::now();
    }
}

/// Record a metric event. Accumulates in memory, flushes to disk on a 5s debounce.
pub fn record(event: AiMetricEvent) {
    with_state(|state| {
        match event {
            AiMetricEvent::LaunchAi { cli } => {
                state.data.launch_ai += 1;
                state.data.last_cli = Some(cli.to_string());
            }
            AiMetricEvent::ExplainSelection => {
                state.data.explain_selection += 1;
            }
            AiMetricEvent::PasteContext => {
                state.data.paste_context += 1;
            }
            AiMetricEvent::GenerateContextFiles => {
                state.data.generate_context_files += 1;
            }
            AiMetricEvent::Error { category } => {
                *state.data.errors.entry(category.to_string()).or_insert(0) += 1;
            }
        }
        state.dirty = true;
    });
}

/// Flush any pending metrics to disk. Call on app exit.
pub fn flush() {
    let mut guard = METRICS.lock().unwrap_or_else(|e| e.into_inner());
    if let Some(state) = guard.as_mut() {
        if state.dirty {
            write_to_disk(&state.data);
            state.dirty = false;
            state.last_write = Instant::now();
        }
    }
}

/// Metric event types.
pub enum AiMetricEvent<'a> {
    LaunchAi { cli: &'a str },
    ExplainSelection,
    PasteContext,
    GenerateContextFiles,
    Error { category: &'a str },
}
