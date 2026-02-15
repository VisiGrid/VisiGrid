//! `vgrid fetch brex-card` and `vgrid fetch brex-bank` — Brex adapters.
//!
//! Shared constants and helpers for both card and cash (bank) transaction APIs.
//! Both use the same API key, base URL, and bearer auth.

pub mod bank;
pub mod card;

use super::common;
use crate::CliError;

// ── Shared constants ────────────────────────────────────────────────

pub(super) const BREX_API_BASE: &str = "https://platform.brexapis.com";
pub(super) const PAGE_LIMIT: u32 = 100;

// ── Shared helpers ──────────────────────────────────────────────────

pub(super) fn extract_brex_error(body: &serde_json::Value, status: u16) -> String {
    body["message"]
        .as_str()
        .or_else(|| body["error"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

pub(super) fn resolve_api_key(flag: Option<String>) -> Result<String, CliError> {
    common::resolve_api_key(flag, "Brex", "BREX_API_KEY")
}
