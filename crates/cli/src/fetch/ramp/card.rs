//! `vgrid fetch ramp-card` — fetch Ramp card transactions into canonical CSV.
//!
//! Ramp returns transaction amounts as decimal dollars (e.g. `90.0`).
//! We convert to minor units (cents) for the canonical format.
//!
//! API: GET https://api.ramp.com/developer/v1/transactions
//! Auth: Bearer token (OAuth access token with `transactions:read` scope)
//! Pagination: cursor-based (`page.next`)
//! Docs: https://docs.ramp.com/developer-api/v1/api-reference/transaction

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::super::common::{self, CanonicalRow, FetchClient};
use super::{extract_ramp_error, parse_ramp_amount, RAMP_API_BASE, PAGE_SIZE};

// ── Internal transaction representation ─────────────────────────────

#[derive(Debug)]
struct RawTransaction {
    effective_date: String,
    posted_date: String,
    amount_minor: i64,
    currency: String,
    canonical_type: String,
    source_id: String,
    card_id: String,
    description: String,
}

// ── Ramp card client ────────────────────────────────────────────────

pub struct RampCardClient {
    client: FetchClient,
    api_key: String,
    base_url: String,
}

impl RampCardClient {
    pub fn new(api_key: String) -> Self {
        Self::with_base_url(api_key, RAMP_API_BASE.to_string())
    }

    pub fn with_base_url(api_key: String, base_url: String) -> Self {
        Self {
            client: FetchClient::new("Ramp", extract_ramp_error),
            api_key,
            base_url,
        }
    }

    /// Fetch all card transactions in the given date range.
    fn fetch_transactions(
        &self,
        from_date: &NaiveDate,
        to_date: &NaiveDate,
        state: &str,
        card_id: Option<&str>,
        entity_id: Option<&str>,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut cursor: Option<String> = None;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        let from_str = from_date.format("%Y-%m-%dT00:00:00Z").to_string();
        // to_date is exclusive in our convention
        let to_str = to_date.format("%Y-%m-%dT00:00:00Z").to_string();
        let page_size_str = PAGE_SIZE.to_string();

        loop {
            page += 1;
            let url = format!("{}/developer/v1/transactions", self.base_url);
            let api_key = self.api_key.clone();
            let from_str = from_str.clone();
            let to_str = to_str.clone();
            let cursor_clone = cursor.clone();
            let page_size_str = page_size_str.clone();
            let state = state.to_string();
            let card_id = card_id.map(|s| s.to_string());
            let entity_id = entity_id.map(|s| s.to_string());

            let body = self.client.request_with_retry(|http| {
                let mut req = http
                    .get(&url)
                    .bearer_auth(&api_key)
                    .query(&[
                        ("from_date", from_str.as_str()),
                        ("to_date", to_str.as_str()),
                        ("page_size", page_size_str.as_str()),
                        ("order_by_date_asc", "true"),
                        ("state", state.as_str()),
                    ]);
                if let Some(ref c) = card_id {
                    req = req.query(&[("card_id", c.as_str())]);
                }
                if let Some(ref e) = entity_id {
                    req = req.query(&[("entity_id", e.as_str())]);
                }
                if let Some(ref c) = cursor_clone {
                    req = req.query(&[("start", c.as_str())]);
                }
                req
            })?;

            let items = body["data"]
                .as_array()
                .ok_or_else(|| CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Ramp response missing 'data' array".into(),
                    hint: None,
                })?;

            let next_cursor = body["page"]["next"]
                .as_str()
                .map(|s| s.to_string());

            // Guard: next_cursor but empty items = malformed response
            if next_cursor.is_some() && items.is_empty() {
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Ramp returned page.next with empty data (malformed response)"
                        .into(),
                    hint: None,
                });
            }

            if show_progress {
                eprintln!("  page {}: {} transactions", page, items.len());
            }

            for item in items {
                all_txns.push(parse_transaction(item)?);
            }

            match next_cursor {
                Some(ref nc) if !nc.is_empty() => {
                    // Infinite loop protection: detect repeated cursor
                    if cursor.as_deref() == Some(nc) {
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "Ramp pagination stuck: cursor={} repeated",
                                nc
                            ),
                            hint: None,
                        });
                    }
                    cursor = Some(nc.clone());
                }
                _ => break,
            }
        }

        Ok(all_txns)
    }
}

// ── Parse a single Ramp card transaction ────────────────────────────

fn parse_transaction(item: &serde_json::Value) -> Result<RawTransaction, CliError> {
    let id = item["id"].as_str().unwrap_or("").to_string();

    // user_transaction_time: ISO 8601 → YYYY-MM-DD
    let effective_date = item["user_transaction_time"]
        .as_str()
        .and_then(|s| s.get(..10))
        .unwrap_or("")
        .to_string();

    // settlement_date: YYYY-MM-DD, fallback to effective_date
    let posted_date = item["settlement_date"]
        .as_str()
        .and_then(|s| s.get(..10))
        .unwrap_or("")
        .to_string();

    let posted_date = if posted_date.is_empty() {
        effective_date.clone()
    } else {
        posted_date
    };

    if effective_date.is_empty() && posted_date.is_empty() {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Ramp transaction {} missing date fields", id),
            hint: None,
        });
    }

    let effective_date = if effective_date.is_empty() {
        posted_date.clone()
    } else {
        effective_date
    };

    // Amount: decimal dollars or object
    let currency_code = item["currency_code"]
        .as_str()
        .unwrap_or("USD");

    let (raw_cents, currency) = parse_ramp_amount(&item["amount"], currency_code)
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Ramp transaction {} bad amount: {}", id, e),
            hint: None,
        })?;

    // Determine type and sign:
    // - If amount from API is negative → refund (positive in canonical)
    // - Otherwise → purchase (negative in canonical, outflow)
    let (canonical_type, amount_minor) = if raw_cents < 0 {
        ("refund".to_string(), raw_cents.abs())
    } else {
        ("purchase".to_string(), -raw_cents.abs())
    };

    let description = item["merchant_name"]
        .as_str()
        .or_else(|| item["merchant_descriptor"].as_str())
        .or_else(|| item["memo"].as_str())
        .unwrap_or("")
        .to_string();

    let card_id = item["card_id"]
        .as_str()
        .unwrap_or("")
        .to_string();

    Ok(RawTransaction {
        effective_date,
        posted_date,
        amount_minor,
        currency,
        canonical_type,
        source_id: id,
        card_id,
        description,
    })
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_ramp_card(
    from: String,
    to: String,
    api_key: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
    state: Option<String>,
    card: Option<String>,
    entity: Option<String>,
) -> Result<(), CliError> {
    // 1. Resolve API key
    let key = super::resolve_api_key(api_key)?;

    // 2. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Ramp card transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let state_filter = state.as_deref().unwrap_or("CLEARED");
    let client = RampCardClient::new(key);
    let mut txns = client.fetch_transactions(
        &from_date,
        &to_date,
        state_filter,
        card.as_deref(),
        entity.as_deref(),
        quiet,
    )?;

    // 4. Sort: (effective_date ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 5. Build canonical rows
    let rows: Vec<CanonicalRow> = txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: txn.effective_date.clone(),
            posted_date: txn.posted_date.clone(),
            amount_minor: txn.amount_minor,
            currency: txn.currency.clone(),
            r#type: txn.canonical_type.clone(),
            source: "ramp_card".to_string(),
            source_id: txn.source_id.clone(),
            group_id: txn.card_id.clone(),
            description: txn.description.clone(),
        })
        .collect();

    // 6. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!("Done: {} transactions written to {}", rows.len(), out_label);
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    #[test]
    fn test_parse_transaction_purchase_float_dollars() {
        let item = serde_json::json!({
            "id": "txn_ramp_001",
            "user_transaction_time": "2026-01-15T14:30:00Z",
            "settlement_date": "2026-01-16",
            "amount": 90.0,
            "currency_code": "USD",
            "state": "CLEARED",
            "merchant_name": "ACME Corp",
            "card_id": "card_ramp_123"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "purchase");
        assert_eq!(txn.amount_minor, -9000); // negated (outflow)
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.card_id, "card_ramp_123");
        assert_eq!(txn.description, "ACME Corp");
        assert_eq!(txn.effective_date, "2026-01-15");
        assert_eq!(txn.posted_date, "2026-01-16");
    }

    #[test]
    fn test_parse_transaction_refund_negative_amount() {
        let item = serde_json::json!({
            "id": "txn_ramp_002",
            "user_transaction_time": "2026-01-20T10:00:00Z",
            "settlement_date": "2026-01-21",
            "amount": -25.5,
            "currency_code": "USD",
            "state": "CLEARED",
            "merchant_name": "Vendor Refund",
            "card_id": "card_ramp_123"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "refund");
        assert_eq!(txn.amount_minor, 2550); // positive (inflow)
    }

    #[test]
    fn test_parse_transaction_object_amount() {
        let item = serde_json::json!({
            "id": "txn_ramp_003",
            "user_transaction_time": "2026-01-10T08:00:00Z",
            "amount": { "amount": 5000, "currency_code": "USD" },
            "currency_code": "USD",
            "state": "CLEARED",
            "merchant_descriptor": "OFFICE SUPPLIES",
            "card_id": "card_ramp_456"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "purchase");
        assert_eq!(txn.amount_minor, -5000); // negated (outflow), already in cents
        assert_eq!(txn.description, "OFFICE SUPPLIES");
        // No settlement_date → falls back to effective_date
        assert_eq!(txn.posted_date, "2026-01-10");
    }

    #[test]
    fn test_parse_transaction_memo_fallback() {
        let item = serde_json::json!({
            "id": "txn_ramp_004",
            "user_transaction_time": "2026-01-12T12:00:00Z",
            "amount": 42.0,
            "currency_code": "USD",
            "memo": "Team lunch",
            "card_id": "card_ramp_789"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.description, "Team lunch");
    }

    #[test]
    fn test_resolve_api_key_from_flag() {
        let key = super::super::resolve_api_key(Some("  ramp_token_123  ".into())).unwrap();
        assert_eq!(key, "ramp_token_123");
    }

    #[test]
    fn test_resolve_api_key_empty_flag() {
        let err = super::super::resolve_api_key(Some("  ".into())).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_resolve_api_key_missing() {
        std::env::remove_var("RAMP_ACCESS_TOKEN");
        let err = super::super::resolve_api_key(None).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("missing Ramp API key"));
    }

    // ── Helper: build a Ramp-shaped transaction JSON ────────────────

    fn mock_txn(id: &str, date: &str, amount: f64, merchant: &str) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "user_transaction_time": format!("{}T12:00:00Z", date),
            "settlement_date": date,
            "amount": amount,
            "currency_code": "USD",
            "state": "CLEARED",
            "merchant_name": merchant,
            "card_id": "card_001"
        })
    }

    fn ramp_list_response(
        items: Vec<serde_json::Value>,
        next_cursor: Option<&str>,
    ) -> serde_json::Value {
        let mut resp = serde_json::json!({ "data": items, "page": {} });
        if let Some(cursor) = next_cursor {
            resp["page"]["next"] = serde_json::Value::String(cursor.to_string());
        }
        resp
    }

    // ── Test: Pagination across 2 pages ─────────────────────────────

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        // Page 1: 2 transactions, has next cursor
        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions")
                .query_param_exists("from_date")
                .query_param_exists("to_date")
                .query_param_missing("start");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(ramp_list_response(
                    vec![
                        mock_txn("txn_001", "2026-01-10", 50.0, "Vendor A"),
                        mock_txn("txn_002", "2026-01-11", 30.0, "Vendor B"),
                    ],
                    Some("cursor_page2"),
                ));
        });

        // Page 2: 1 transaction, no next cursor
        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions")
                .query_param("start", "cursor_page2");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(ramp_list_response(
                    vec![mock_txn("txn_003", "2026-01-12", -10.0, "Refund Co")],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampCardClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let txns = client
            .fetch_transactions(&from, &to, "CLEARED", None, None, true)
            .unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].source_id, "txn_001");
        assert_eq!(txns[0].amount_minor, -5000); // purchase
        assert_eq!(txns[1].source_id, "txn_002");
        assert_eq!(txns[2].source_id, "txn_003");
        assert_eq!(txns[2].canonical_type, "refund");
        assert_eq!(txns[2].amount_minor, 1000); // refund (positive)
    }

    // ── Test: Retry on 429 exhausted ────────────────────────────────

    #[test]
    fn test_retry_on_429_exhausted() {
        let server = MockServer::start();

        let rate_limit_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions");
            then.status(429)
                .header("retry-after", "0")
                .json_body(serde_json::json!({
                    "error": { "message": "Too many requests" }
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampCardClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, "CLEARED", None, None, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_RATE_LIMIT);
        assert!(err.message.contains("rate limited"));
        rate_limit_mock.assert_calls(4);
    }

    // ── Test: Auth failure → exit 51 ────────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions");
            then.status(401)
                .json_body(serde_json::json!({
                    "error": { "message": "Invalid access token" }
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampCardClient::with_base_url(
            "bad_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, "CLEARED", None, None, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(err.message.contains("Ramp auth failed (401)"));
    }

    // ── Test: page.next + empty data → error ────────────────────────

    #[test]
    fn test_next_cursor_empty_data_error() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "data": [],
                    "page": { "next": "bad_cursor" }
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampCardClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, "CLEARED", None, None, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_UPSTREAM);
        assert!(err.message.contains("page.next with empty data"));
    }

    // ── Test: Deterministic sorting ─────────────────────────────────

    #[test]
    fn test_deterministic_sorting() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(ramp_list_response(
                    vec![
                        mock_txn("txn_c", "2026-01-15", 10.0, "Vendor C"),
                        mock_txn("txn_a", "2026-01-15", 20.0, "Vendor A"),
                        mock_txn("txn_b", "2026-01-10", 5.0, "Vendor B"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampCardClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let mut txns = client
            .fetch_transactions(&from, &to, "CLEARED", None, None, true)
            .unwrap();
        txns.sort_by(|a, b| {
            a.effective_date
                .cmp(&b.effective_date)
                .then_with(|| a.source_id.cmp(&b.source_id))
        });

        assert_eq!(txns[0].source_id, "txn_b"); // Jan 10
        assert_eq!(txns[1].source_id, "txn_a"); // Jan 15, "a" < "c"
        assert_eq!(txns[2].source_id, "txn_c"); // Jan 15
    }
}
