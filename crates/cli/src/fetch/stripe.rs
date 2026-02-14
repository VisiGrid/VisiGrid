//! `vgrid fetch stripe` — fetch Stripe balance transactions into canonical CSV.

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const STRIPE_API_BASE: &str = "https://api.stripe.com";
const PAGE_LIMIT: u32 = 100;

// ── Internal transaction representation ─────────────────────────────

/// Internal representation with epoch for sorting.
#[derive(Debug)]
struct RawTransaction {
    created_epoch: i64,
    available_on_epoch: i64,
    amount_minor: i64,
    currency: String,
    canonical_type: String,
    source_id: String,
    group_id: String,
    description: String,
}

// ── Type mapping ────────────────────────────────────────────────────

fn map_stripe_type(stripe_type: &str) -> &'static str {
    match stripe_type {
        "charge" | "payment" => "charge",
        "payout" => "payout",
        "stripe_fee" | "application_fee" => "fee",
        "refund" => "refund",
        _ => "adjustment",
    }
}

// ── Epoch ↔ date helpers ────────────────────────────────────────────

fn epoch_to_date(epoch: i64) -> String {
    chrono::DateTime::from_timestamp(epoch, 0)
        .map(|dt| dt.format("%Y-%m-%d").to_string())
        .unwrap_or_default()
}

fn date_to_epoch(date: &NaiveDate) -> i64 {
    date.and_hms_opt(0, 0, 0)
        .unwrap()
        .and_utc()
        .timestamp()
}

// ── Stripe client ───────────────────────────────────────────────────

pub struct StripeClient {
    client: FetchClient,
    api_key: String,
    account: Option<String>,
    base_url: String,
}

impl StripeClient {
    pub fn new(api_key: String, account: Option<String>) -> Self {
        Self::with_base_url(api_key, account, STRIPE_API_BASE.to_string())
    }

    pub fn with_base_url(api_key: String, account: Option<String>, base_url: String) -> Self {
        Self {
            client: FetchClient::new("Stripe", extract_stripe_error),
            api_key,
            account,
            base_url,
        }
    }

    /// Fetch all balance transactions in the given epoch range.
    fn fetch_balance_transactions(
        &self,
        from_epoch: i64,
        to_epoch: i64,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut starting_after: Option<String> = None;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        loop {
            page += 1;
            let mut params = vec![
                ("created[gte]".to_string(), from_epoch.to_string()),
                ("created[lt]".to_string(), to_epoch.to_string()),
                ("limit".to_string(), PAGE_LIMIT.to_string()),
            ];
            if let Some(ref after) = starting_after {
                params.push(("starting_after".to_string(), after.clone()));
            }

            let url = format!("{}/v1/balance_transactions", self.base_url);
            let api_key = self.api_key.clone();
            let account = self.account.clone();

            let body = self.client.request_with_retry(|http| {
                let mut req = http
                    .get(&url)
                    .basic_auth(&api_key, Some(""))
                    .query(&params);
                if let Some(ref acct) = account {
                    req = req.header("Stripe-Account", acct);
                }
                req
            })?;

            let data = body["data"]
                .as_array()
                .ok_or_else(|| CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Stripe response missing 'data' array".into(),
                    hint: None,
                })?;

            let has_more = body["has_more"].as_bool().unwrap_or(false);

            // Guard: has_more but empty data = malformed response
            if has_more && data.is_empty() {
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Stripe returned has_more=true with empty data (malformed response)"
                        .into(),
                    hint: None,
                });
            }

            if show_progress {
                eprintln!("  page {}: {} transactions", page, data.len());
            }

            for item in data {
                all_txns.push(parse_transaction(item)?);
            }

            if !has_more {
                break;
            }

            // Pagination: use last item's ID
            let last_id = data
                .last()
                .and_then(|item| item["id"].as_str())
                .map(|s| s.to_string())
                .ok_or_else(|| CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Stripe transaction missing 'id' field for pagination".into(),
                    hint: None,
                })?;

            // Infinite loop protection: detect repeated starting_after
            if starting_after.as_deref() == Some(&last_id) {
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: format!(
                        "Stripe pagination stuck: starting_after={} repeated",
                        last_id
                    ),
                    hint: None,
                });
            }

            starting_after = Some(last_id);
        }

        Ok(all_txns)
    }
}

// ── Parse a single Stripe transaction ───────────────────────────────

fn parse_transaction(item: &serde_json::Value) -> Result<RawTransaction, CliError> {
    let created = item["created"].as_i64().ok_or_else(|| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: "Stripe transaction missing 'created' field".into(),
        hint: None,
    })?;

    let available_on = item["available_on"].as_i64().ok_or_else(|| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: "Stripe transaction missing 'available_on' field".into(),
        hint: None,
    })?;

    let amount = item["amount"].as_i64().ok_or_else(|| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: "Stripe transaction missing 'amount' field".into(),
        hint: None,
    })?;

    let currency = item["currency"]
        .as_str()
        .unwrap_or("usd")
        .to_uppercase();

    let raw_type = item["type"].as_str().unwrap_or("unknown");
    let canonical_type = map_stripe_type(raw_type).to_string();

    let source_id = item["id"].as_str().unwrap_or("").to_string();

    let raw_description = item["description"].as_str().unwrap_or("").to_string();

    // Annotate description for unknown types
    let description = if canonical_type == "adjustment" && raw_type != "adjustment" {
        if raw_description.is_empty() {
            format!("[stripe_type: {}]", raw_type)
        } else {
            format!("{} [stripe_type: {}]", raw_description, raw_type)
        }
    } else {
        raw_description
    };

    // group_id extraction
    let group_id = if let Some(payout_id) = item["payout"].as_str() {
        payout_id.to_string()
    } else if canonical_type == "payout" {
        source_id.clone()
    } else {
        String::new()
    };

    Ok(RawTransaction {
        created_epoch: created,
        available_on_epoch: available_on,
        amount_minor: amount,
        currency,
        canonical_type,
        source_id,
        group_id,
        description,
    })
}

fn extract_stripe_error(body: &serde_json::Value, status: u16) -> String {
    body["error"]["message"]
        .as_str()
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_stripe(
    from: String,
    to: String,
    api_key: Option<String>,
    out: Option<PathBuf>,
    account: Option<String>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Resolve API key
    let key = resolve_api_key(api_key)?;

    // 2. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    let from_epoch = date_to_epoch(&from_date);
    let to_epoch = date_to_epoch(&to_date);

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Stripe balance transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let client = StripeClient::new(key, account);
    let mut txns = client.fetch_balance_transactions(from_epoch, to_epoch, quiet)?;

    // 4. Sort: (created_epoch ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.created_epoch
            .cmp(&b.created_epoch)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 5. Build canonical rows
    let rows: Vec<CanonicalRow> = txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: epoch_to_date(txn.created_epoch),
            posted_date: epoch_to_date(txn.available_on_epoch),
            amount_minor: txn.amount_minor,
            currency: txn.currency.clone(),
            r#type: txn.canonical_type.clone(),
            source: "stripe".to_string(),
            source_id: txn.source_id.clone(),
            group_id: txn.group_id.clone(),
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

fn resolve_api_key(flag: Option<String>) -> Result<String, CliError> {
    common::resolve_api_key(flag, "Stripe", "STRIPE_API_KEY")
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    #[test]
    fn test_map_stripe_type() {
        assert_eq!(map_stripe_type("charge"), "charge");
        assert_eq!(map_stripe_type("payment"), "charge");
        assert_eq!(map_stripe_type("payout"), "payout");
        assert_eq!(map_stripe_type("stripe_fee"), "fee");
        assert_eq!(map_stripe_type("application_fee"), "fee");
        assert_eq!(map_stripe_type("refund"), "refund");
        assert_eq!(map_stripe_type("issuing_transaction"), "adjustment");
        assert_eq!(map_stripe_type("unknown_thing"), "adjustment");
    }

    #[test]
    fn test_epoch_to_date() {
        assert_eq!(epoch_to_date(1768435200), "2026-01-15");
    }

    #[test]
    fn test_date_to_epoch() {
        let d = NaiveDate::from_ymd_opt(2026, 1, 15).unwrap();
        assert_eq!(date_to_epoch(&d), 1768435200);
    }

    #[test]
    fn test_parse_transaction_charge() {
        let item = serde_json::json!({
            "id": "txn_123",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": 5000,
            "currency": "usd",
            "type": "charge",
            "description": "Payment from customer",
            "payout": "po_abc"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "charge");
        assert_eq!(txn.amount_minor, 5000);
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.group_id, "po_abc");
        assert_eq!(txn.description, "Payment from customer");
    }

    #[test]
    fn test_parse_transaction_payout_self_groups() {
        let item = serde_json::json!({
            "id": "txn_po_456",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": -10000,
            "currency": "usd",
            "type": "payout",
            "description": "STRIPE PAYOUT",
            "payout": serde_json::Value::Null
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "payout");
        assert_eq!(txn.group_id, "txn_po_456");
    }

    #[test]
    fn test_resolve_api_key_from_flag() {
        let key = resolve_api_key(Some("  sk_test_123  ".into())).unwrap();
        assert_eq!(key, "sk_test_123");
    }

    #[test]
    fn test_resolve_api_key_empty_flag() {
        let err = resolve_api_key(Some("  ".into())).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_resolve_api_key_missing() {
        std::env::remove_var("STRIPE_API_KEY");
        let err = resolve_api_key(None).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    // ── Helper: build a Stripe-shaped transaction JSON ──────────────

    fn mock_txn(id: &str, created: i64, amount: i64, txn_type: &str) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "created": created,
            "available_on": created + 86400,
            "amount": amount,
            "currency": "usd",
            "type": txn_type,
            "description": format!("Txn {}", id),
            "payout": serde_json::Value::Null
        })
    }

    fn stripe_list_response(data: Vec<serde_json::Value>, has_more: bool) -> serde_json::Value {
        serde_json::json!({
            "object": "list",
            "data": data,
            "has_more": has_more,
            "url": "/v1/balance_transactions"
        })
    }

    // ── Test 1: Pagination across 2 pages ───────────────────────────

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        // Page 1: 2 transactions, has_more = true
        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions")
                .query_param_exists("created[gte]")
                .query_param_exists("created[lt]")
                .query_param_missing("starting_after");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(stripe_list_response(
                    vec![
                        mock_txn("txn_001", 1000, 5000, "charge"),
                        mock_txn("txn_002", 1001, 3000, "charge"),
                    ],
                    true,
                ));
        });

        // Page 2: 1 transaction, has_more = false
        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions")
                .query_param("starting_after", "txn_002");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(stripe_list_response(
                    vec![mock_txn("txn_003", 1002, -100, "stripe_fee")],
                    false,
                ));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let txns = client.fetch_balance_transactions(0, 2000, true).unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].source_id, "txn_001");
        assert_eq!(txns[1].source_id, "txn_002");
        assert_eq!(txns[2].source_id, "txn_003");
        assert_eq!(txns[2].canonical_type, "fee");
    }

    // ── Test 2: Retry on 429 with Retry-After ───────────────────────
    //
    // Verifies that 429 responses exhaust retries and produce exit 53.
    // (Testing successful retry after 429 requires mock sequencing which
    // httpmock doesn't support natively. The retry loop itself is simple
    // enough that testing the terminal failure path is sufficient.)

    #[test]
    fn test_retry_on_429_exhausted() {
        let server = MockServer::start();

        // All requests return 429
        let rate_limit_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions");
            then.status(429)
                .header("retry-after", "0")
                .json_body(serde_json::json!({
                    "error": {
                        "type": "rate_limit",
                        "message": "Too many requests"
                    }
                }));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let err = client
            .fetch_balance_transactions(0, 2000, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_RATE_LIMIT);
        assert!(
            err.message.contains("rate limited"),
            "message: {}",
            err.message,
        );
        // Should have been called 1 initial + 3 retries = 4 times
        rate_limit_mock.assert_calls(4);
    }

    // ── Test 3: Unknown type maps to adjustment ─────────────────────

    #[test]
    fn test_unknown_type_maps_to_adjustment() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(stripe_list_response(
                    vec![serde_json::json!({
                        "id": "txn_issuing",
                        "created": 1000,
                        "available_on": 2000,
                        "amount": -500,
                        "currency": "usd",
                        "type": "issuing_transaction",
                        "description": "Card purchase"
                    })],
                    false,
                ));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let txns = client.fetch_balance_transactions(0, 2000, true).unwrap();
        assert_eq!(txns.len(), 1);
        assert_eq!(txns[0].canonical_type, "adjustment");
        assert_eq!(
            txns[0].description,
            "Card purchase [stripe_type: issuing_transaction]"
        );
    }

    // ── Test 4: Deterministic sorting ───────────────────────────────

    #[test]
    fn test_deterministic_sorting() {
        let server = MockServer::start();

        // Return transactions in scrambled order
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(stripe_list_response(
                    vec![
                        mock_txn("txn_c", 1000, 100, "charge"),
                        mock_txn("txn_a", 1000, 200, "charge"),
                        mock_txn("txn_b", 500, 50, "stripe_fee"),
                    ],
                    false,
                ));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let mut txns = client.fetch_balance_transactions(0, 2000, true).unwrap();
        txns.sort_by(|a, b| {
            a.created_epoch
                .cmp(&b.created_epoch)
                .then_with(|| a.source_id.cmp(&b.source_id))
        });

        assert_eq!(txns[0].source_id, "txn_b"); // epoch 500
        assert_eq!(txns[1].source_id, "txn_a"); // epoch 1000, "a" < "c"
        assert_eq!(txns[2].source_id, "txn_c"); // epoch 1000

        // Write to CSV and verify byte-identical on second run
        let csv1 = write_txns_to_csv(&txns);

        let mut txns2 = client.fetch_balance_transactions(0, 2000, true).unwrap();
        txns2.sort_by(|a, b| {
            a.created_epoch
                .cmp(&b.created_epoch)
                .then_with(|| a.source_id.cmp(&b.source_id))
        });
        let csv2 = write_txns_to_csv(&txns2);

        assert_eq!(csv1, csv2, "CSV output must be byte-identical across runs");
    }

    /// Helper: write sorted transactions to a CSV string.
    fn write_txns_to_csv(txns: &[RawTransaction]) -> String {
        let mut buf = Vec::new();
        {
            let mut wtr = csv::WriterBuilder::new()
                .terminator(csv::Terminator::Any(b'\n'))
                .from_writer(&mut buf);
            for txn in txns {
                let row = CanonicalRow {
                    effective_date: epoch_to_date(txn.created_epoch),
                    posted_date: epoch_to_date(txn.available_on_epoch),
                    amount_minor: txn.amount_minor,
                    currency: txn.currency.clone(),
                    r#type: txn.canonical_type.clone(),
                    source_id: txn.source_id.clone(),
                    source: "stripe".to_string(),
                    group_id: txn.group_id.clone(),
                    description: txn.description.clone(),
                };
                wtr.serialize(&row).unwrap();
            }
            wtr.flush().unwrap();
        }
        String::from_utf8(buf).unwrap()
    }

    // ── Test 5: Missing key → exit 50 ───────────────────────────────

    #[test]
    fn test_missing_key_exit_50() {
        std::env::remove_var("STRIPE_API_KEY");
        let err = resolve_api_key(None).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(
            err.message.contains("missing Stripe API key"),
            "message: {}",
            err.message,
        );
    }

    // ── Test 6: Auth failure → exit 51 ──────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions");
            then.status(401)
                .json_body(serde_json::json!({
                    "error": {
                        "type": "invalid_request_error",
                        "message": "Invalid API Key provided: sk_test_****_bad"
                    }
                }));
        });

        let client = StripeClient::with_base_url(
            "sk_test_bad".into(),
            None,
            server.base_url(),
        );

        let err = client
            .fetch_balance_transactions(0, 2000, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("Stripe auth failed (401)"),
            "message: {}",
            err.message,
        );
        assert!(
            err.message.contains("Invalid API Key"),
            "message: {}",
            err.message,
        );
    }

    // ── Test 7: has_more + empty data → error ───────────────────────

    #[test]
    fn test_has_more_empty_data_error() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "object": "list",
                    "data": [],
                    "has_more": true,
                    "url": "/v1/balance_transactions"
                }));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let err = client
            .fetch_balance_transactions(0, 2000, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_UPSTREAM);
        assert!(
            err.message.contains("has_more=true with empty data"),
            "message: {}",
            err.message,
        );
    }
}
