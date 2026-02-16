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
    btxn_id: String,
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

    /// Fetch the balance transaction IDs that belong to a specific payout.
    ///
    /// Uses `/v1/balance_transactions?payout=po_xxx` — the only reliable way
    /// to associate btxns with their payout, since the Balance Transaction
    /// object does not carry a `payout` field.
    fn fetch_payout_member_ids(
        &self,
        payout_id: &str,
        quiet: bool,
    ) -> Result<Vec<String>, CliError> {
        let mut btxn_ids = Vec::new();
        let mut starting_after: Option<String> = None;

        loop {
            let mut params = vec![
                ("payout".to_string(), payout_id.to_string()),
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

            for item in data {
                if let Some(id) = item["id"].as_str() {
                    btxn_ids.push(id.to_string());
                }
            }

            if !has_more || data.is_empty() {
                break;
            }

            let last_id = data
                .last()
                .and_then(|item| item["id"].as_str())
                .map(|s| s.to_string());

            if last_id == starting_after.as_deref().map(|s| s.to_string()) {
                break; // stuck pagination
            }
            starting_after = last_id;
        }

        if !quiet && atty::is(atty::Stream::Stderr) {
            eprintln!("  payout {}: {} members", payout_id, btxn_ids.len());
        }

        Ok(btxn_ids)
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
                all_txns.extend(parse_transaction(item)?);
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

fn parse_transaction(item: &serde_json::Value) -> Result<Vec<RawTransaction>, CliError> {
    let btxn_id = item["id"]
        .as_str()
        .unwrap_or("")
        .to_string();

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

    let fee = item["fee"].as_i64().unwrap_or(0);

    let currency = item["currency"]
        .as_str()
        .unwrap_or("usd")
        .to_uppercase();

    let raw_type = item["type"].as_str().unwrap_or("unknown");
    let canonical_type = map_stripe_type(raw_type).to_string();

    // Prefer the source object ID (ch_xxx, po_xxx) over the balance
    // transaction ID (txn_xxx).  The `source` field may be a string or an
    // expanded object; fall back to `id` when absent.
    let source_id = item["source"]
        .as_str()
        .or_else(|| item["source"]["id"].as_str())
        .or_else(|| item["id"].as_str())
        .unwrap_or("")
        .to_string();

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

    // group_id: payout btxns self-group via source_id (po_xxx).
    // All other btxns get group_id populated later via per-payout API calls,
    // since the Balance Transaction object does NOT carry a `payout` field.
    let group_id = if canonical_type == "payout" {
        source_id.clone()
    } else {
        String::new()
    };

    let mut rows = vec![RawTransaction {
        btxn_id: btxn_id.clone(),
        created_epoch: created,
        available_on_epoch: available_on,
        amount_minor: amount,
        currency: currency.clone(),
        canonical_type: canonical_type.clone(),
        source_id: source_id.clone(),
        group_id: group_id.clone(),
        description,
    }];

    // Emit a synthetic fee row when the balance transaction carries a non-zero
    // fee.  Stripe embeds the processing fee as a field on the charge (not as a
    // separate balance transaction), so without this the payout-group rollup
    // (charges + fees + payout = 0) can never balance.  Skip for transactions
    // that are already fee-typed to avoid double-counting.
    if fee != 0 && canonical_type != "fee" {
        rows.push(RawTransaction {
            btxn_id: String::new(), // synthetic — no Stripe ID
            created_epoch: created,
            available_on_epoch: available_on,
            amount_minor: -fee,
            currency,
            canonical_type: "fee".to_string(),
            source_id: format!("{}_fee", source_id),
            group_id,
            description: "Stripe processing fee".to_string(),
        });
    }

    Ok(rows)
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

    // 4. Populate group_id via per-payout API calls.
    //    The Balance Transaction object does NOT carry a `payout` field, so we
    //    must ask Stripe which btxns belong to each payout.
    let payout_ids: Vec<String> = txns
        .iter()
        .filter(|t| t.canonical_type == "payout")
        .map(|t| t.source_id.clone())
        .collect();

    if !payout_ids.is_empty() {
        if show_progress {
            eprintln!("Resolving payout membership for {} payouts...", payout_ids.len());
        }

        // Build btxn_id → payout_id map
        let mut btxn_to_payout = std::collections::HashMap::new();
        for payout_id in &payout_ids {
            let member_ids = client.fetch_payout_member_ids(payout_id, quiet)?;
            for btxn_id in member_ids {
                btxn_to_payout.insert(btxn_id, payout_id.clone());
            }
        }

        // Tag real btxns with their payout group
        for txn in &mut txns {
            if txn.group_id.is_empty() && !txn.btxn_id.is_empty() {
                if let Some(payout_id) = btxn_to_payout.get(&txn.btxn_id) {
                    txn.group_id = payout_id.clone();
                }
            }
        }

        // Propagate group_id to synthetic fee rows: their source_id is
        // "{parent_source_id}_fee", so strip the suffix and look up the parent.
        let source_to_group: std::collections::HashMap<String, String> = txns
            .iter()
            .filter(|t| !t.group_id.is_empty() && !t.btxn_id.is_empty())
            .map(|t| (t.source_id.clone(), t.group_id.clone()))
            .collect();

        for txn in &mut txns {
            if txn.group_id.is_empty() && txn.btxn_id.is_empty() {
                if let Some(parent_source) = txn.source_id.strip_suffix("_fee") {
                    if let Some(payout_id) = source_to_group.get(parent_source) {
                        txn.group_id = payout_id.clone();
                    }
                }
            }
        }
    }

    // 5. Sort: (created_epoch ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.created_epoch
            .cmp(&b.created_epoch)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 6. Build canonical rows
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

    // 7. Write CSV
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
        // parse_transaction does NOT set group_id for charges — that's done
        // later via per-payout API calls in cmd_fetch_stripe.
        let item = serde_json::json!({
            "id": "txn_123",
            "source": "ch_123",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": 5000,
            "fee": 175,
            "currency": "usd",
            "type": "charge",
            "description": "Payment from customer"
        });
        let rows = parse_transaction(&item).unwrap();
        assert_eq!(rows.len(), 2, "charge with fee should emit charge + synthetic fee row");

        let txn = &rows[0];
        assert_eq!(txn.btxn_id, "txn_123");
        assert_eq!(txn.canonical_type, "charge");
        assert_eq!(txn.amount_minor, 5000);
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.source_id, "ch_123");
        assert_eq!(txn.group_id, "", "group_id populated later via payout membership API");
        assert_eq!(txn.description, "Payment from customer");

        let fee_row = &rows[1];
        assert_eq!(fee_row.btxn_id, "", "synthetic fee has no Stripe ID");
        assert_eq!(fee_row.canonical_type, "fee");
        assert_eq!(fee_row.amount_minor, -175);
        assert_eq!(fee_row.group_id, "", "propagated from parent later");
        assert_eq!(fee_row.source_id, "ch_123_fee");
    }

    #[test]
    fn test_parse_transaction_charge_zero_fee() {
        let item = serde_json::json!({
            "id": "txn_nofee",
            "source": "ch_nofee",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": 5000,
            "fee": 0,
            "currency": "usd",
            "type": "charge",
            "description": "Zero-fee charge"
        });
        let rows = parse_transaction(&item).unwrap();
        assert_eq!(rows.len(), 1, "charge with fee=0 should not emit synthetic fee row");
        assert_eq!(rows[0].source_id, "ch_nofee");
    }

    #[test]
    fn test_parse_transaction_stripe_fee_no_double_count() {
        let item = serde_json::json!({
            "id": "txn_sf",
            "source": "fee_sf",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": -500,
            "fee": 0,
            "currency": "usd",
            "type": "stripe_fee",
            "description": "Stripe fee"
        });
        let rows = parse_transaction(&item).unwrap();
        assert_eq!(rows.len(), 1, "stripe_fee type should not emit synthetic fee row");
        assert_eq!(rows[0].canonical_type, "fee");
        assert_eq!(rows[0].source_id, "fee_sf");
    }

    #[test]
    fn test_parse_transaction_payout_self_groups() {
        // Payout btxns self-group via their source_id (po_xxx).
        let item = serde_json::json!({
            "id": "txn_po_456",
            "source": "po_456",
            "created": 1768435200,
            "available_on": 1768521600,
            "amount": -10000,
            "fee": 0,
            "currency": "usd",
            "type": "payout",
            "description": "STRIPE PAYOUT"
        });
        let rows = parse_transaction(&item).unwrap();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].btxn_id, "txn_po_456");
        assert_eq!(rows[0].canonical_type, "payout");
        assert_eq!(rows[0].source_id, "po_456");
        assert_eq!(rows[0].group_id, "po_456");
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
            "description": format!("Txn {}", id)
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

    // ── Test 8: Payout membership resolution ─────────────────────

    #[test]
    fn test_fetch_payout_member_ids() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/balance_transactions")
                .query_param("payout", "po_test");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(stripe_list_response(
                    vec![
                        mock_txn("txn_ch1", 1000, 5000, "charge"),
                        mock_txn("txn_fee1", 1000, -145, "stripe_fee"),
                        mock_txn("txn_po", 1001, -4855, "payout"),
                    ],
                    false,
                ));
        });

        let client = StripeClient::with_base_url(
            "sk_test_key".into(),
            None,
            server.base_url(),
        );

        let ids = client.fetch_payout_member_ids("po_test", true).unwrap();
        assert_eq!(ids, vec!["txn_ch1", "txn_fee1", "txn_po"]);
    }

    // ── Test 9: Group ID propagation to synthetic fees ───────────

    #[test]
    fn test_group_id_propagation() {
        // Simulate the post-fetch group_id population logic
        let mut txns = vec![
            RawTransaction {
                btxn_id: "txn_ch1".into(),
                created_epoch: 1000, available_on_epoch: 2000,
                amount_minor: 5000, currency: "USD".into(),
                canonical_type: "charge".into(), source_id: "ch_001".into(),
                group_id: String::new(), description: "Charge".into(),
            },
            RawTransaction {
                btxn_id: String::new(), // synthetic
                created_epoch: 1000, available_on_epoch: 2000,
                amount_minor: -145, currency: "USD".into(),
                canonical_type: "fee".into(), source_id: "ch_001_fee".into(),
                group_id: String::new(), description: "Stripe processing fee".into(),
            },
            RawTransaction {
                btxn_id: "txn_po1".into(),
                created_epoch: 1001, available_on_epoch: 2001,
                amount_minor: -4855, currency: "USD".into(),
                canonical_type: "payout".into(), source_id: "po_001".into(),
                group_id: "po_001".into(), description: "STRIPE PAYOUT".into(),
            },
        ];

        // Simulate btxn→payout mapping from API
        let mut btxn_to_payout = std::collections::HashMap::new();
        btxn_to_payout.insert("txn_ch1".to_string(), "po_001".to_string());
        btxn_to_payout.insert("txn_po1".to_string(), "po_001".to_string());

        // Tag real btxns
        for txn in &mut txns {
            if txn.group_id.is_empty() && !txn.btxn_id.is_empty() {
                if let Some(payout_id) = btxn_to_payout.get(&txn.btxn_id) {
                    txn.group_id = payout_id.clone();
                }
            }
        }

        // Propagate to synthetic fees
        let source_to_group: std::collections::HashMap<String, String> = txns
            .iter()
            .filter(|t| !t.group_id.is_empty() && !t.btxn_id.is_empty())
            .map(|t| (t.source_id.clone(), t.group_id.clone()))
            .collect();

        for txn in &mut txns {
            if txn.group_id.is_empty() && txn.btxn_id.is_empty() {
                if let Some(parent_source) = txn.source_id.strip_suffix("_fee") {
                    if let Some(payout_id) = source_to_group.get(parent_source) {
                        txn.group_id = payout_id.clone();
                    }
                }
            }
        }

        // All should now have group_id = "po_001"
        assert_eq!(txns[0].group_id, "po_001", "charge should get group_id from payout membership");
        assert_eq!(txns[1].group_id, "po_001", "synthetic fee should inherit from parent charge");
        assert_eq!(txns[2].group_id, "po_001", "payout self-groups");

        // Verify rollup: charges + fees + payout = 0
        let sum: i64 = txns.iter().map(|t| t.amount_minor).sum();
        assert_eq!(sum, 0, "payout group must balance");
    }
}
