//! `vgrid fetch ramp-card` and `vgrid fetch ramp-bank` — Ramp adapters.
//!
//! Shared constants and helpers for both card and business account transaction APIs.
//! Both use the same OAuth API at api.ramp.com with bearer auth.

pub mod bank;
pub mod card;

use super::common;
use crate::CliError;

// ── Shared constants ────────────────────────────────────────────────

pub(super) const RAMP_API_BASE: &str = "https://api.ramp.com";
pub(super) const PAGE_SIZE: u32 = 100;

// ── Shared helpers ──────────────────────────────────────────────────

pub(super) fn extract_ramp_error(body: &serde_json::Value, status: u16) -> String {
    body["error"]["message"]
        .as_str()
        .or_else(|| body["error"].as_str())
        .or_else(|| body["message"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

pub(super) fn resolve_api_key(flag: Option<String>) -> Result<String, CliError> {
    common::resolve_api_key(flag, "Ramp", "RAMP_ACCESS_TOKEN")
}

/// Parse the Ramp amount field, which can be either:
/// - A float in decimal dollars (e.g. `90.0`) → convert to cents
/// - An object `{ "amount": cents, "currency": "USD" }` → use cents directly
///
/// Returns (amount_cents, currency).
pub(super) fn parse_ramp_amount(
    amount_val: &serde_json::Value,
    currency_fallback: &str,
) -> Result<(i64, String), String> {
    // Object variant: { amount: cents, currency: "USD" }
    if amount_val.is_object() {
        let cents = amount_val["amount"]
            .as_i64()
            .or_else(|| amount_val["amount"].as_f64().map(|f| f.round() as i64))
            .ok_or_else(|| "missing amount in amount object".to_string())?;
        let currency = amount_val["currency_code"]
            .as_str()
            .or_else(|| amount_val["currency"].as_str())
            .unwrap_or(currency_fallback)
            .to_uppercase();
        return Ok((cents, currency));
    }

    // Float variant: decimal dollars (e.g. 90.0)
    if let Some(f) = amount_val.as_f64() {
        let cents = (f * 100.0).round() as i64;
        return Ok((cents, currency_fallback.to_uppercase()));
    }

    // String variant: "90.0"
    if let Some(s) = amount_val.as_str() {
        let f: f64 = s.parse().map_err(|e| format!("bad amount string: {}", e))?;
        let cents = (f * 100.0).round() as i64;
        return Ok((cents, currency_fallback.to_uppercase()));
    }

    Err(format!("unexpected amount format: {}", amount_val))
}
