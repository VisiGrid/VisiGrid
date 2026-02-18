//! `vgrid fetch ramp-bank` — fetch Ramp business account transactions into canonical CSV.
//!
//! Uses the same transactions API as ramp-card, but treats transactions as
//! business account activity (deposits, withdrawals) rather than card purchases.
//!
//! API: GET https://api.ramp.com/developer/v1/transactions
//! Auth: Bearer token (OAuth access token with `transactions:read` scope)
//! Pagination: cursor-based (`page.next`)

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
    description: String,
}

// ── Ramp bank client ────────────────────────────────────────────────

pub struct RampBankClient {
    client: FetchClient,
    api_key: String,
    base_url: String,
}

impl RampBankClient {
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

    /// Fetch all business account transactions in the given date range.
    fn fetch_transactions(
        &self,
        from_date: &NaiveDate,
        to_date: &NaiveDate,
        state: &str,
        entity_id: Option<&str>,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut cursor: Option<String> = None;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        let from_str = from_date.format("%Y-%m-%dT00:00:00Z").to_string();
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

// ── Parse a single Ramp bank transaction ────────────────────────────

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

    let (amount_minor, currency) = parse_ramp_amount(&item["amount"], currency_code)
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Ramp transaction {} bad amount: {}", id, e),
            hint: None,
        })?;

    // Sign-based type mapping (like brex-bank):
    // positive = inflow (deposit), negative = outflow (withdrawal)
    let canonical_type = if amount_minor >= 0 {
        "deposit".to_string()
    } else {
        "withdrawal".to_string()
    };

    let description = item["merchant_name"]
        .as_str()
        .or_else(|| item["memo"].as_str())
        .or_else(|| item["merchant_descriptor"].as_str())
        .unwrap_or("")
        .to_string();

    Ok(RawTransaction {
        effective_date,
        posted_date,
        amount_minor,
        currency,
        canonical_type,
        source_id: id,
        description,
    })
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_ramp_bank(
    from: String,
    to: String,
    api_key: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
    state: Option<String>,
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
            "Fetching Ramp bank transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let state_filter = state.as_deref().unwrap_or("CLEARED");
    let client = RampBankClient::new(key);
    let mut txns = client.fetch_transactions(
        &from_date,
        &to_date,
        state_filter,
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
            source: "ramp".to_string(),
            source_id: txn.source_id.clone(),
            group_id: String::new(),
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
    fn test_parse_transaction_deposit() {
        let item = serde_json::json!({
            "id": "txn_bank_001",
            "user_transaction_time": "2026-01-15T14:30:00Z",
            "settlement_date": "2026-01-16",
            "amount": 1500.0,
            "currency_code": "USD",
            "merchant_name": "ACH DEPOSIT"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "deposit");
        assert_eq!(txn.amount_minor, 150000);
        assert_eq!(txn.posted_date, "2026-01-16");
        assert_eq!(txn.effective_date, "2026-01-15");
    }

    #[test]
    fn test_parse_transaction_withdrawal() {
        let item = serde_json::json!({
            "id": "txn_bank_002",
            "user_transaction_time": "2026-01-20T10:00:00Z",
            "settlement_date": "2026-01-21",
            "amount": -500.25,
            "currency_code": "USD",
            "memo": "PAYROLL"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "withdrawal");
        assert_eq!(txn.amount_minor, -50025);
        assert_eq!(txn.description, "PAYROLL");
    }

    #[test]
    fn test_parse_transaction_object_amount() {
        let item = serde_json::json!({
            "id": "txn_bank_003",
            "user_transaction_time": "2026-01-10T08:00:00Z",
            "amount": { "amount": 250000, "currency_code": "USD" },
            "currency_code": "USD",
            "merchant_descriptor": "WIRE TRANSFER"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "deposit");
        assert_eq!(txn.amount_minor, 250000);
        assert_eq!(txn.description, "WIRE TRANSFER");
    }

    #[test]
    fn test_parse_transaction_fallback_dates() {
        // No settlement_date → falls back to effective_date
        let item = serde_json::json!({
            "id": "txn_bank_004",
            "user_transaction_time": "2026-01-15T12:00:00Z",
            "amount": 100.0,
            "currency_code": "USD"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.posted_date, "2026-01-15");
        assert_eq!(txn.effective_date, "2026-01-15");
    }

    #[test]
    fn test_parse_transaction_description_fallback_order() {
        // merchant_name takes priority over memo
        let item = serde_json::json!({
            "id": "txn_bank_005",
            "user_transaction_time": "2026-01-15T12:00:00Z",
            "amount": 50.0,
            "currency_code": "USD",
            "merchant_name": "Primary Name",
            "memo": "Secondary memo",
            "merchant_descriptor": "Tertiary descriptor"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.description, "Primary Name");

        // Memo when no merchant_name
        let item2 = serde_json::json!({
            "id": "txn_bank_006",
            "user_transaction_time": "2026-01-15T12:00:00Z",
            "amount": 50.0,
            "currency_code": "USD",
            "memo": "Secondary memo",
            "merchant_descriptor": "Tertiary descriptor"
        });
        let txn2 = parse_transaction(&item2).unwrap();
        assert_eq!(txn2.description, "Secondary memo");
    }

    // ── Mock helpers ────────────────────────────────────────────────

    fn mock_bank_txn(
        id: &str,
        date: &str,
        amount: f64,
        desc: &str,
    ) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "user_transaction_time": format!("{}T12:00:00Z", date),
            "settlement_date": date,
            "amount": amount,
            "currency_code": "USD",
            "state": "CLEARED",
            "merchant_name": desc
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
                        mock_bank_txn("txn_001", "2026-01-10", 1500.0, "ACH DEPOSIT"),
                        mock_bank_txn("txn_002", "2026-01-15", -500.0, "PAYROLL"),
                    ],
                    Some("cursor_page2"),
                ));
        });

        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions")
                .query_param("start", "cursor_page2");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(ramp_list_response(
                    vec![
                        mock_bank_txn("txn_003", "2026-01-20", -25.0, "Wire fee"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampBankClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let txns = client
            .fetch_transactions(&from, &to, "CLEARED", None, true)
            .unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].canonical_type, "deposit");
        assert_eq!(txns[0].amount_minor, 150000);
        assert_eq!(txns[1].canonical_type, "withdrawal");
        assert_eq!(txns[1].amount_minor, -50000);
        assert_eq!(txns[2].canonical_type, "withdrawal");
        assert_eq!(txns[2].amount_minor, -2500);
    }

    // ── Test: Auth failure ──────────────────────────────────────────

    #[test]
    fn test_auth_failure() {
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

        let client = RampBankClient::with_base_url(
            "bad_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions(&from, &to, "CLEARED", None, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(err.message.contains("Ramp auth failed (401)"));
    }

    // ── Test: Mixed deposit/withdrawal ──────────────────────────────

    #[test]
    fn test_mixed_transaction_types() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/developer/v1/transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(ramp_list_response(
                    vec![
                        mock_bank_txn("txn_a", "2026-01-10", 5000.0, "Wire in"),
                        mock_bank_txn("txn_b", "2026-01-15", -800.0, "Vendor payment"),
                        mock_bank_txn("txn_c", "2026-01-20", 0.0, "Zero amount"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = RampBankClient::with_base_url(
            "ramp_test_key".into(),
            server.base_url(),
        );

        let txns = client
            .fetch_transactions(&from, &to, "CLEARED", None, true)
            .unwrap();

        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].canonical_type, "deposit");
        assert_eq!(txns[0].amount_minor, 500000);
        assert_eq!(txns[1].canonical_type, "withdrawal");
        assert_eq!(txns[1].amount_minor, -80000);
        assert_eq!(txns[2].canonical_type, "deposit"); // zero → deposit
        assert_eq!(txns[2].amount_minor, 0);
    }
}
