//! `vgrid fetch brex-card` — fetch Brex card transactions into canonical CSV.
//!
//! Brex returns only settled transactions. Amounts are already in minor
//! units (cents), so no float parsing is needed.
//!
//! API: GET https://platform.brexapis.com/v2/transactions/card/primary
//! Auth: Bearer token (API key with `transactions.card.readonly` scope)
//! Pagination: cursor-based (`next_cursor`)
//! Docs: https://developer.brex.com/openapi/transactions_api/

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::super::common::{self, CanonicalRow, FetchClient};
use super::{extract_brex_error, BREX_API_BASE, PAGE_LIMIT};

// ── Internal transaction representation ─────────────────────────────

#[derive(Debug)]
struct RawTransaction {
    posted_date: String,
    amount_minor: i64,
    currency: String,
    canonical_type: String,
    source_id: String,
    card_id: String,
    description: String,
}

// ── Type mapping ────────────────────────────────────────────────────

fn map_brex_type(brex_type: &str) -> &'static str {
    match brex_type {
        "PURCHASE" => "purchase",
        "REFUND" => "refund",
        "CHARGEBACK" => "chargeback",
        "COLLECTION" => "collection",
        "REWARDS_CREDIT" => "credit",
        _ => "adjustment",
    }
}

// ── Brex card client ────────────────────────────────────────────────

pub struct BrexCardClient {
    client: FetchClient,
    api_key: String,
    base_url: String,
}

impl BrexCardClient {
    pub fn new(api_key: String) -> Self {
        Self::with_base_url(api_key, BREX_API_BASE.to_string())
    }

    pub fn with_base_url(api_key: String, base_url: String) -> Self {
        Self {
            client: FetchClient::new("Brex", extract_brex_error),
            api_key,
            base_url,
        }
    }

    /// Fetch all settled card transactions in the given date range.
    fn fetch_transactions(
        &self,
        from_date: &NaiveDate,
        to_date: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut cursor: Option<String> = None;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        // Brex filters by posted_at_start (inclusive YYYY-MM-DD)
        let from_str = from_date.format("%Y-%m-%d").to_string();
        // to_date is exclusive in our convention; Brex uses posted_at_end as inclusive,
        // so subtract one day to get the last inclusive date.
        let to_inclusive = *to_date - chrono::Duration::days(1);
        let to_str = to_inclusive.format("%Y-%m-%d").to_string();

        loop {
            page += 1;
            let url = format!("{}/v2/transactions/card/primary", self.base_url);
            let api_key = self.api_key.clone();
            let from_str = from_str.clone();
            let to_str = to_str.clone();
            let cursor_clone = cursor.clone();
            let limit = PAGE_LIMIT.to_string();

            let body = self.client.request_with_retry(|http| {
                let mut req = http
                    .get(&url)
                    .bearer_auth(&api_key)
                    .query(&[
                        ("posted_at_start", from_str.as_str()),
                        ("posted_at_end", to_str.as_str()),
                        ("limit", limit.as_str()),
                    ]);
                if let Some(ref c) = cursor_clone {
                    req = req.query(&[("cursor", c.as_str())]);
                }
                req
            })?;

            let items = body["items"]
                .as_array()
                .ok_or_else(|| CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Brex response missing 'items' array".into(),
                    hint: None,
                })?;

            let next_cursor = body["next_cursor"].as_str().map(|s| s.to_string());

            // Guard: next_cursor but empty items = malformed response
            if next_cursor.is_some() && items.is_empty() {
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Brex returned next_cursor with empty items (malformed response)"
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
                                "Brex pagination stuck: cursor={} repeated",
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

// ── Parse a single Brex card transaction ────────────────────────────

fn parse_transaction(item: &serde_json::Value) -> Result<RawTransaction, CliError> {
    let id = item["id"].as_str().unwrap_or("").to_string();

    let posted_at_date = item["posted_at_date"]
        .as_str()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Brex transaction {} missing 'posted_at_date'", id),
            hint: None,
        })?
        .to_string();

    // Amount is nested: { "amount": 5000, "currency": "USD" }
    let amount_obj = &item["amount"];
    let amount_cents = amount_obj["amount"]
        .as_i64()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Brex transaction {} missing 'amount.amount'", id),
            hint: None,
        })?;

    let currency = amount_obj["currency"]
        .as_str()
        .unwrap_or("USD")
        .to_uppercase();

    let raw_type = item["type"].as_str().unwrap_or("PURCHASE");
    let canonical_type = map_brex_type(raw_type).to_string();

    // Negate purchases to be consistent (outflows are negative)
    let amount_minor = match raw_type {
        "REFUND" | "CHARGEBACK" | "REWARDS_CREDIT" => {
            // These are typically positive from Brex (money coming back);
            // keep as-is since they represent inflows
            amount_cents
        }
        _ => {
            // Purchases are negative (outflow), Brex sends positive amounts
            // Negate to match our outflow convention
            -amount_cents.abs()
        }
    };

    let description = item["description"]
        .as_str()
        .unwrap_or("")
        .to_string();

    let card_id = item["card_id"]
        .as_str()
        .unwrap_or("")
        .to_string();

    Ok(RawTransaction {
        posted_date: posted_at_date,
        amount_minor,
        currency,
        canonical_type,
        source_id: id,
        card_id,
        description,
    })
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_brex_card(
    from: String,
    to: String,
    api_key: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
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
            "Fetching Brex card transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let client = BrexCardClient::new(key);
    let mut txns = client.fetch_transactions(&from_date, &to_date, quiet)?;

    // 4. Sort: (posted_date ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.posted_date
            .cmp(&b.posted_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 5. Build canonical rows
    let rows: Vec<CanonicalRow> = txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: txn.posted_date.clone(),
            posted_date: txn.posted_date.clone(),
            amount_minor: txn.amount_minor,
            currency: txn.currency.clone(),
            r#type: txn.canonical_type.clone(),
            source: "brex_card".to_string(),
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
    fn test_map_brex_type() {
        assert_eq!(map_brex_type("PURCHASE"), "purchase");
        assert_eq!(map_brex_type("REFUND"), "refund");
        assert_eq!(map_brex_type("CHARGEBACK"), "chargeback");
        assert_eq!(map_brex_type("COLLECTION"), "collection");
        assert_eq!(map_brex_type("REWARDS_CREDIT"), "credit");
        assert_eq!(map_brex_type("SOMETHING_NEW"), "adjustment");
    }

    #[test]
    fn test_parse_transaction_purchase() {
        let item = serde_json::json!({
            "id": "txn_brex_001",
            "posted_at_date": "2026-01-15",
            "amount": { "amount": 5000, "currency": "USD" },
            "type": "PURCHASE",
            "description": "ACME Corp Office Supplies",
            "card_id": "card_123"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "purchase");
        assert_eq!(txn.amount_minor, -5000); // negated (outflow)
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.card_id, "card_123");
        assert_eq!(txn.description, "ACME Corp Office Supplies");
    }

    #[test]
    fn test_parse_transaction_refund() {
        let item = serde_json::json!({
            "id": "txn_brex_002",
            "posted_at_date": "2026-01-20",
            "amount": { "amount": 2500, "currency": "USD" },
            "type": "REFUND",
            "description": "Refund from vendor",
            "card_id": "card_123"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "refund");
        assert_eq!(txn.amount_minor, 2500); // positive (inflow)
    }

    #[test]
    fn test_resolve_api_key_from_flag() {
        let key = super::super::resolve_api_key(Some("  brex_token_123  ".into())).unwrap();
        assert_eq!(key, "brex_token_123");
    }

    #[test]
    fn test_resolve_api_key_empty_flag() {
        let err = super::super::resolve_api_key(Some("  ".into())).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_resolve_api_key_missing() {
        std::env::remove_var("BREX_API_KEY");
        let err = super::super::resolve_api_key(None).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("missing Brex API key"));
    }

    // ── Helper: build a Brex-shaped transaction JSON ────────────────

    fn mock_txn(id: &str, date: &str, amount: i64, txn_type: &str) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "posted_at_date": date,
            "amount": { "amount": amount, "currency": "USD" },
            "type": txn_type,
            "description": format!("Txn {}", id),
            "card_id": "card_001"
        })
    }

    fn brex_list_response(
        items: Vec<serde_json::Value>,
        next_cursor: Option<&str>,
    ) -> serde_json::Value {
        let mut resp = serde_json::json!({ "items": items });
        if let Some(cursor) = next_cursor {
            resp["next_cursor"] = serde_json::Value::String(cursor.to_string());
        }
        resp
    }

    // ── Test: Pagination across 2 pages ─────────────────────────────

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        // Page 1: 2 transactions, has next_cursor
        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/card/primary")
                .query_param_exists("posted_at_start")
                .query_param_exists("posted_at_end")
                .query_param_missing("cursor");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![
                        mock_txn("txn_001", "2026-01-10", 5000, "PURCHASE"),
                        mock_txn("txn_002", "2026-01-11", 3000, "PURCHASE"),
                    ],
                    Some("cursor_page2"),
                ));
        });

        // Page 2: 1 transaction, no next_cursor
        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/card/primary")
                .query_param("cursor", "cursor_page2");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![mock_txn("txn_003", "2026-01-12", 100, "REFUND")],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexCardClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let txns = client.fetch_transactions(&from, &to, true).unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].source_id, "txn_001");
        assert_eq!(txns[1].source_id, "txn_002");
        assert_eq!(txns[2].source_id, "txn_003");
        assert_eq!(txns[2].canonical_type, "refund");
    }

    // ── Test: Retry on 429 exhausted ────────────────────────────────

    #[test]
    fn test_retry_on_429_exhausted() {
        let server = MockServer::start();

        let rate_limit_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/card/primary");
            then.status(429)
                .header("retry-after", "0")
                .json_body(serde_json::json!({
                    "message": "Too many requests"
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexCardClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, true)
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
                .path("/v2/transactions/card/primary");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Invalid token"
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexCardClient::with_base_url(
            "bad_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(err.message.contains("Brex auth failed (401)"));
    }

    // ── Test: next_cursor + empty items → error ─────────────────────

    #[test]
    fn test_next_cursor_empty_items_error() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/card/primary");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "items": [],
                    "next_cursor": "bad_cursor"
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexCardClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_UPSTREAM);
        assert!(err.message.contains("next_cursor with empty items"));
    }

    // ── Test: Deterministic sorting ─────────────────────────────────

    #[test]
    fn test_deterministic_sorting() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/card/primary");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![
                        mock_txn("txn_c", "2026-01-15", 100, "PURCHASE"),
                        mock_txn("txn_a", "2026-01-15", 200, "PURCHASE"),
                        mock_txn("txn_b", "2026-01-10", 50, "PURCHASE"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexCardClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let mut txns = client.fetch_transactions(&from, &to, true).unwrap();
        txns.sort_by(|a, b| {
            a.posted_date
                .cmp(&b.posted_date)
                .then_with(|| a.source_id.cmp(&b.source_id))
        });

        assert_eq!(txns[0].source_id, "txn_b"); // Jan 10
        assert_eq!(txns[1].source_id, "txn_a"); // Jan 15, "a" < "c"
        assert_eq!(txns[2].source_id, "txn_c"); // Jan 15
    }
}
