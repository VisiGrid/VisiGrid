//! `vgrid fetch brex-bank` — fetch Brex cash (bank) transactions into canonical CSV.
//!
//! Targets the Brex cash account transactions API. Amounts are in minor
//! units (cents). This adapter treats Brex as a bank account, similar to
//! Mercury — it emits deposits, withdrawals, transfers, etc.
//!
//! API: GET https://platform.brexapis.com/v2/transactions/cash/{cash_account_id}
//! Accounts: GET https://platform.brexapis.com/v2/accounts/cash
//! Auth: Bearer token (API key)
//! Pagination: cursor-based (`next_cursor`)
//! Docs: https://developer.brex.com/openapi/transactions_api/

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::super::common::{self, CanonicalRow, FetchClient};
use super::{extract_brex_error, BREX_API_BASE, PAGE_LIMIT};

// ── Internal representations ────────────────────────────────────────

#[derive(Debug)]
struct RawTransaction {
    posted_date: String,
    initiated_date: String,
    amount_minor: i64,
    currency: String,
    canonical_type: String,
    source_id: String,
    description: String,
}

#[derive(Debug)]
struct BrexCashAccount {
    id: String,
    name: String,
    status: String,
}

// ── Type mapping ────────────────────────────────────────────────────

/// Map Brex cash transaction type to canonical type.
///
/// Brex cash types are inferred from their API. For sign-ambiguous types
/// we use the amount sign to distinguish deposits from withdrawals.
fn map_brex_cash_type(brex_type: &str, amount: i64) -> &'static str {
    match brex_type {
        // Outgoing payments (ACH, wire, check)
        "PAYMENT" | "PAYMENT_FAILED" => "withdrawal",
        // Incoming — receiving money
        "DEPOSIT" | "RECEIVED" => "deposit",
        // Transfers between accounts
        "TRANSFER" | "BOOK_TRANSFER" => "transfer",
        // Card settlements hitting the cash account
        "CARD_SETTLEMENT" => "expense",
        // Fees
        "FEE" => "fee",
        // Interest earned
        "INTEREST" => "interest",
        // Reversal/return
        "RETURN" | "REVERSAL" => {
            if amount >= 0 { "deposit" } else { "withdrawal" }
        }
        // ACH — direction depends on sign
        "ACH_TRANSFER" | "ACH" => {
            if amount >= 0 { "deposit" } else { "withdrawal" }
        }
        // Wire — direction depends on sign
        "WIRE_TRANSFER" | "WIRE" => {
            if amount >= 0 { "deposit" } else { "withdrawal" }
        }
        // Catch-all: use sign
        _ => {
            if amount >= 0 { "deposit" } else { "withdrawal" }
        }
    }
}

// ── Brex bank client ────────────────────────────────────────────────

pub struct BrexBankClient {
    client: FetchClient,
    api_key: String,
    base_url: String,
}

impl BrexBankClient {
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

    /// List all Brex cash accounts.
    fn list_accounts(&self) -> Result<Vec<BrexCashAccount>, CliError> {
        let url = format!("{}/v2/accounts/cash", self.base_url);
        let api_key = self.api_key.clone();

        let body = self.client.request_with_retry(|http| {
            http.get(&url).bearer_auth(&api_key)
        })?;

        let items = body["items"]
            .as_array()
            .ok_or_else(|| CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message: "Brex response missing 'items' array for accounts".into(),
                hint: None,
            })?;

        let mut result = Vec::new();
        for acct in items {
            result.push(BrexCashAccount {
                id: acct["id"].as_str().unwrap_or("").to_string(),
                name: acct["name"].as_str().unwrap_or("").to_string(),
                status: acct["status"].as_str().unwrap_or("").to_uppercase(),
            });
        }

        Ok(result)
    }

    /// Fetch all cash transactions for an account in the given date range.
    fn fetch_transactions(
        &self,
        account_id: &str,
        from_date: &NaiveDate,
        to_date: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut cursor: Option<String> = None;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        let from_str = from_date.format("%Y-%m-%d").to_string();
        // to_date is exclusive in our convention; Brex uses posted_at_end as inclusive
        let to_inclusive = *to_date - chrono::Duration::days(1);
        let to_str = to_inclusive.format("%Y-%m-%d").to_string();

        loop {
            page += 1;
            let url = format!(
                "{}/v2/transactions/cash/{}",
                self.base_url, account_id,
            );
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

// ── Parse a single Brex cash transaction ────────────────────────────

fn parse_transaction(item: &serde_json::Value) -> Result<RawTransaction, CliError> {
    let id = item["id"].as_str().unwrap_or("").to_string();

    let posted_at_date = item["posted_at_date"]
        .as_str()
        .unwrap_or("");

    let initiated_at_date = item["initiated_at_date"]
        .as_str()
        .unwrap_or("");

    // Use posted_at_date if available, fall back to initiated_at_date
    let posted_date = if !posted_at_date.is_empty() {
        posted_at_date.to_string()
    } else if !initiated_at_date.is_empty() {
        initiated_at_date.to_string()
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Brex transaction {} missing date fields", id),
            hint: None,
        });
    };

    let initiated_date = if !initiated_at_date.is_empty() {
        initiated_at_date.to_string()
    } else {
        posted_date.clone()
    };

    // Amount is nested: { "amount": 5000, "currency": "USD" }
    let amount_obj = &item["amount"];
    let amount_minor = amount_obj["amount"]
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

    let raw_type = item["type"].as_str().unwrap_or("UNKNOWN");
    let canonical_type = map_brex_cash_type(raw_type, amount_minor).to_string();

    let description = item["description"]
        .as_str()
        .or_else(|| item["memo"].as_str())
        .unwrap_or("")
        .to_string();

    Ok(RawTransaction {
        posted_date,
        initiated_date,
        amount_minor,
        currency,
        canonical_type,
        source_id: id,
        description,
    })
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_brex_bank(
    from: String,
    to: String,
    api_key: Option<String>,
    out: Option<PathBuf>,
    account: Option<String>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Resolve API key
    let key = super::resolve_api_key(api_key)?;

    // 2. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    let client = BrexBankClient::new(key);

    // 3. Resolve account
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    let account_id = match account {
        Some(id) => id,
        None => {
            let accounts = client.list_accounts()?;
            let active: Vec<_> = accounts
                .iter()
                .filter(|a| a.status == "ACTIVE")
                .collect();
            match active.len() {
                0 => {
                    return Err(CliError::args(
                        "no active Brex cash accounts found".to_string(),
                    ));
                }
                1 => {
                    if show_progress {
                        eprintln!("Using account: {} ({})", active[0].name, active[0].id);
                    }
                    active[0].id.clone()
                }
                _ => {
                    let names: Vec<_> = active.iter().map(|a| a.name.as_str()).collect();
                    return Err(CliError::args(format!(
                        "multiple Brex cash accounts found; use --account to specify ({})",
                        names.join(", "),
                    )));
                }
            }
        }
    };

    // 4. Fetch
    if show_progress {
        eprintln!(
            "Fetching Brex bank transactions ({} to {})...",
            from_date, to_date,
        );
        let display_id = if account_id.len() > 12 {
            format!("{}...", &account_id[..12])
        } else {
            account_id.clone()
        };
        eprintln!("  Account: {}", display_id);
    }

    let mut txns = client.fetch_transactions(&account_id, &from_date, &to_date, quiet)?;

    // 5. Sort: (posted_date ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.posted_date
            .cmp(&b.posted_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 6. Build canonical rows
    let rows: Vec<CanonicalRow> = txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: txn.initiated_date.clone(),
            posted_date: txn.posted_date.clone(),
            amount_minor: txn.amount_minor,
            currency: txn.currency.clone(),
            r#type: txn.canonical_type.clone(),
            source: "brex".to_string(),
            source_id: txn.source_id.clone(),
            group_id: String::new(),
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

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── Unit tests ──────────────────────────────────────────────────

    #[test]
    fn test_map_brex_cash_type() {
        // Fixed mappings
        assert_eq!(map_brex_cash_type("PAYMENT", -5000), "withdrawal");
        assert_eq!(map_brex_cash_type("DEPOSIT", 10000), "deposit");
        assert_eq!(map_brex_cash_type("RECEIVED", 3000), "deposit");
        assert_eq!(map_brex_cash_type("TRANSFER", 1000), "transfer");
        assert_eq!(map_brex_cash_type("BOOK_TRANSFER", -500), "transfer");
        assert_eq!(map_brex_cash_type("CARD_SETTLEMENT", -2000), "expense");
        assert_eq!(map_brex_cash_type("FEE", -25), "fee");
        assert_eq!(map_brex_cash_type("INTEREST", 100), "interest");

        // Sign-dependent
        assert_eq!(map_brex_cash_type("ACH_TRANSFER", 5000), "deposit");
        assert_eq!(map_brex_cash_type("ACH_TRANSFER", -5000), "withdrawal");
        assert_eq!(map_brex_cash_type("ACH", 3000), "deposit");
        assert_eq!(map_brex_cash_type("ACH", -3000), "withdrawal");
        assert_eq!(map_brex_cash_type("WIRE_TRANSFER", 10000), "deposit");
        assert_eq!(map_brex_cash_type("WIRE_TRANSFER", -10000), "withdrawal");
        assert_eq!(map_brex_cash_type("RETURN", 500), "deposit");
        assert_eq!(map_brex_cash_type("RETURN", -500), "withdrawal");
        assert_eq!(map_brex_cash_type("REVERSAL", 200), "deposit");

        // Unknown → sign-based
        assert_eq!(map_brex_cash_type("SOMETHING_NEW", 100), "deposit");
        assert_eq!(map_brex_cash_type("SOMETHING_NEW", -100), "withdrawal");
    }

    #[test]
    fn test_parse_transaction_deposit() {
        let item = serde_json::json!({
            "id": "txn_cash_001",
            "posted_at_date": "2026-01-15",
            "initiated_at_date": "2026-01-14",
            "amount": { "amount": 150000, "currency": "USD" },
            "type": "ACH_TRANSFER",
            "description": "STRIPE TRANSFER"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "deposit");
        assert_eq!(txn.amount_minor, 150000);
        assert_eq!(txn.posted_date, "2026-01-15");
        assert_eq!(txn.initiated_date, "2026-01-14");
    }

    #[test]
    fn test_parse_transaction_withdrawal() {
        let item = serde_json::json!({
            "id": "txn_cash_002",
            "posted_at_date": "2026-01-20",
            "initiated_at_date": "2026-01-19",
            "amount": { "amount": -50000, "currency": "USD" },
            "type": "PAYMENT",
            "description": "GUSTO PAYROLL"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.canonical_type, "withdrawal");
        assert_eq!(txn.amount_minor, -50000);
    }

    #[test]
    fn test_parse_transaction_fallback_dates() {
        // Only posted_at_date, no initiated_at_date
        let item = serde_json::json!({
            "id": "txn_cash_003",
            "posted_at_date": "2026-01-15",
            "amount": { "amount": 1000, "currency": "USD" },
            "type": "INTEREST",
            "description": "Interest payment"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.posted_date, "2026-01-15");
        assert_eq!(txn.initiated_date, "2026-01-15"); // falls back to posted

        // Only initiated_at_date, no posted_at_date
        let item2 = serde_json::json!({
            "id": "txn_cash_004",
            "initiated_at_date": "2026-01-10",
            "amount": { "amount": -2000, "currency": "USD" },
            "type": "PAYMENT",
            "description": "Vendor payment"
        });
        let txn2 = parse_transaction(&item2).unwrap();
        assert_eq!(txn2.posted_date, "2026-01-10"); // falls back to initiated
        assert_eq!(txn2.initiated_date, "2026-01-10");
    }

    #[test]
    fn test_parse_transaction_memo_fallback() {
        let item = serde_json::json!({
            "id": "txn_cash_005",
            "posted_at_date": "2026-01-15",
            "amount": { "amount": -500, "currency": "USD" },
            "type": "FEE",
            "memo": "Monthly account fee"
        });
        let txn = parse_transaction(&item).unwrap();
        assert_eq!(txn.description, "Monthly account fee");
    }

    // ── Mock helpers ────────────────────────────────────────────────

    fn mock_cash_txn(
        id: &str,
        posted_date: &str,
        amount: i64,
        txn_type: &str,
        desc: &str,
    ) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "posted_at_date": posted_date,
            "initiated_at_date": posted_date,
            "amount": { "amount": amount, "currency": "USD" },
            "type": txn_type,
            "description": desc
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

    // ── Test: Account auto-detect (single active) ───────────────────

    #[test]
    fn test_auto_detect_single_account() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET).path("/v2/accounts/cash");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "items": [
                        { "id": "cash_001", "name": "Operating", "status": "ACTIVE" },
                        { "id": "cash_002", "name": "Closed", "status": "CLOSED" }
                    ]
                }));
        });

        let client = BrexBankClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let accounts = client.list_accounts().unwrap();
        let active: Vec<_> = accounts.iter().filter(|a| a.status == "ACTIVE").collect();
        assert_eq!(active.len(), 1);
        assert_eq!(active[0].id, "cash_001");
        assert_eq!(active[0].name, "Operating");
    }

    // ── Test: Multiple active accounts → error ──────────────────────

    #[test]
    fn test_multiple_active_accounts() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET).path("/v2/accounts/cash");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "items": [
                        { "id": "cash_001", "name": "Operating", "status": "ACTIVE" },
                        { "id": "cash_002", "name": "Savings", "status": "ACTIVE" }
                    ]
                }));
        });

        let client = BrexBankClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let accounts = client.list_accounts().unwrap();
        let active: Vec<_> = accounts.iter().filter(|a| a.status == "ACTIVE").collect();
        assert_eq!(active.len(), 2);
    }

    // ── Test: Pagination across 2 pages ─────────────────────────────

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/cash/cash_001")
                .query_param_exists("posted_at_start")
                .query_param_exists("posted_at_end")
                .query_param_missing("cursor");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![
                        mock_cash_txn("txn_001", "2026-01-10", 150000, "ACH_TRANSFER", "STRIPE TRANSFER"),
                        mock_cash_txn("txn_002", "2026-01-15", -50000, "PAYMENT", "GUSTO PAYROLL"),
                    ],
                    Some("cursor_page2"),
                ));
        });

        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/cash/cash_001")
                .query_param("cursor", "cursor_page2");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![
                        mock_cash_txn("txn_003", "2026-01-20", -2500, "FEE", "Wire fee"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexBankClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let txns = client
            .fetch_transactions("cash_001", &from, &to, true)
            .unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 3);
        assert_eq!(txns[0].canonical_type, "deposit");
        assert_eq!(txns[0].amount_minor, 150000);
        assert_eq!(txns[1].canonical_type, "withdrawal");
        assert_eq!(txns[1].amount_minor, -50000);
        assert_eq!(txns[2].canonical_type, "fee");
        assert_eq!(txns[2].amount_minor, -2500);
    }

    // ── Test: Auth failure ──────────────────────────────────────────

    #[test]
    fn test_auth_failure() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/cash/cash_001");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Invalid token"
                }));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexBankClient::with_base_url(
            "bad_key".into(),
            server.base_url(),
        );

        let err = client
            .fetch_transactions("cash_001", &from, &to, true)
            .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(err.message.contains("Brex auth failed (401)"));
    }

    // ── Test: Mixed transaction types ───────────────────────────────

    #[test]
    fn test_mixed_transaction_types() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v2/transactions/cash/cash_001");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(brex_list_response(
                    vec![
                        mock_cash_txn("txn_a", "2026-01-10", 150000, "ACH_TRANSFER", "Stripe payout"),
                        mock_cash_txn("txn_b", "2026-01-15", -80000, "PAYMENT", "Vendor payment"),
                        mock_cash_txn("txn_c", "2026-01-15", -25000, "CARD_SETTLEMENT", "Card settlement"),
                        mock_cash_txn("txn_d", "2026-01-20", 500, "INTEREST", "Interest"),
                        mock_cash_txn("txn_e", "2026-01-25", -100, "FEE", "Wire fee"),
                        mock_cash_txn("txn_f", "2026-01-28", 5000, "WIRE_TRANSFER", "Incoming wire"),
                    ],
                    None,
                ));
        });

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();

        let client = BrexBankClient::with_base_url(
            "brex_test_key".into(),
            server.base_url(),
        );

        let txns = client
            .fetch_transactions("cash_001", &from, &to, true)
            .unwrap();

        assert_eq!(txns.len(), 6);
        assert_eq!(txns[0].canonical_type, "deposit");    // ACH in
        assert_eq!(txns[1].canonical_type, "withdrawal");  // PAYMENT
        assert_eq!(txns[2].canonical_type, "expense");     // CARD_SETTLEMENT
        assert_eq!(txns[3].canonical_type, "interest");    // INTEREST
        assert_eq!(txns[4].canonical_type, "fee");         // FEE
        assert_eq!(txns[5].canonical_type, "deposit");     // WIRE in
    }
}
