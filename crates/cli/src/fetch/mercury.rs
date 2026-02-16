//! `vgrid fetch mercury` — fetch Mercury bank transactions into canonical CSV.

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const MERCURY_API_BASE: &str = "https://api.mercury.com/api/v1";
const PAGE_LIMIT: u32 = 500;

// ── Internal transaction representation ─────────────────────────────

/// Internal representation for sorting before CSV output.
#[derive(Debug)]
struct RawTransaction {
    created_iso: String,
    posted_date: String,
    amount_minor: i64,
    canonical_type: String,
    source_id: String,
    description: String,
}

/// Mercury account from the /accounts endpoint.
#[derive(Debug)]
struct MercuryAccount {
    id: String,
    name: String,
    status: String,
}

// ── Type mapping ────────────────────────────────────────────────────

fn map_mercury_kind(kind: &str, amount: f64) -> &'static str {
    match kind {
        "externalTransfer" => {
            if amount > 0.0 { "deposit" } else { "withdrawal" }
        }
        "internalTransfer" => "transfer",
        "outgoingPayment" => "withdrawal",
        "incomingDomesticWire" | "incomingInternationalWire" => "deposit",
        "debitCardTransaction" | "cardTransaction" => "expense",
        "fee" => "fee",
        "check" => {
            if amount < 0.0 { "withdrawal" } else { "deposit" }
        }
        // Mercury returns "other" for ACH credits/debits (Stripe transfers,
        // Forte deposits, etc.). Fall back to sign-based classification.
        _ => {
            if amount > 0.0 { "deposit" } else { "withdrawal" }
        }
    }
}

// ── Helpers ─────────────────────────────────────────────────────────

/// Convert ISO 8601 datetime to `YYYY-MM-DD` by truncating at `T`.
fn parse_iso_date(iso: &str) -> String {
    iso.split('T').next().unwrap_or(iso).to_string()
}

/// Convert Mercury float amount (2dp) to integer minor units (cents).
fn amount_to_minor(amount: f64) -> i64 {
    (amount * 100.0).round() as i64
}

/// Pick the first non-empty string from candidates, or empty.
fn pick_description(bank_description: &str, counterparty_name: &str, note: &str) -> String {
    if !bank_description.is_empty() {
        bank_description.to_string()
    } else if !counterparty_name.is_empty() {
        counterparty_name.to_string()
    } else if !note.is_empty() {
        note.to_string()
    } else {
        String::new()
    }
}

// ── Mercury client ──────────────────────────────────────────────────

pub struct MercuryClient {
    client: FetchClient,
    api_key: String,
    base_url: String,
}

impl MercuryClient {
    pub fn new(api_key: String) -> Self {
        Self::with_base_url(api_key, MERCURY_API_BASE.to_string())
    }

    pub fn with_base_url(api_key: String, base_url: String) -> Self {
        Self {
            client: FetchClient::new("Mercury", extract_mercury_error),
            api_key,
            base_url,
        }
    }

    /// List all Mercury accounts.
    fn list_accounts(&self) -> Result<Vec<MercuryAccount>, CliError> {
        let url = format!("{}/accounts", self.base_url);
        let api_key = self.api_key.clone();

        let body = self.client.request_with_retry(|http| {
            http.get(&url).bearer_auth(&api_key)
        })?;

        let accounts = body["accounts"]
            .as_array()
            .ok_or_else(|| CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message: "Mercury response missing 'accounts' array".into(),
                hint: None,
            })?;

        let mut result = Vec::new();
        for acct in accounts {
            result.push(MercuryAccount {
                id: acct["id"].as_str().unwrap_or("").to_string(),
                name: acct["name"].as_str().unwrap_or("").to_string(),
                status: acct["status"].as_str().unwrap_or("").to_string(),
            });
        }

        Ok(result)
    }

    /// Fetch all transactions for an account in the given date range.
    fn fetch_transactions(
        &self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawTransaction>, CliError> {
        let mut all_txns = Vec::new();
        let mut offset: u32 = 0;
        let mut page = 0u32;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;
        let mut prev_first_id: Option<String> = None;

        loop {
            page += 1;
            let url = format!(
                "{}/account/{}/transactions",
                self.base_url, account_id,
            );
            let params = vec![
                ("start".to_string(), from.to_string()),
                ("end".to_string(), to.to_string()),
                ("limit".to_string(), PAGE_LIMIT.to_string()),
                ("offset".to_string(), offset.to_string()),
            ];
            let api_key = self.api_key.clone();

            let body = self.client.request_with_retry(|http| {
                http.get(&url)
                    .bearer_auth(&api_key)
                    .query(&params)
            })?;

            let transactions = body["transactions"]
                .as_array()
                .ok_or_else(|| CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: "Mercury response missing 'transactions' array".into(),
                    hint: None,
                })?;

            let count = transactions.len() as u32;

            if show_progress {
                eprintln!("  page {}: {} transactions", page, count);
            }

            // Pagination guard: detect stuck pagination
            if count == PAGE_LIMIT {
                let first_id = transactions
                    .first()
                    .and_then(|t| t["id"].as_str())
                    .map(|s| s.to_string());
                if first_id.is_some() && first_id == prev_first_id {
                    return Err(CliError {
                        code: exit_codes::EXIT_FETCH_UPSTREAM,
                        message: "Mercury pagination stuck: same page returned twice".into(),
                        hint: None,
                    });
                }
                prev_first_id = first_id;
            }

            for item in transactions {
                if let Some(txn) = parse_transaction(item)? {
                    all_txns.push(txn);
                }
            }

            if count < PAGE_LIMIT {
                break;
            }

            offset += PAGE_LIMIT;
        }

        Ok(all_txns)
    }
}

// ── Parse a single Mercury transaction ──────────────────────────────

/// Parse a Mercury transaction JSON. Returns `None` for filtered-out statuses.
fn parse_transaction(item: &serde_json::Value) -> Result<Option<RawTransaction>, CliError> {
    // Status filtering: only include sent and pending
    let status = item["status"].as_str().unwrap_or("");
    if status != "sent" && status != "pending" {
        return Ok(None);
    }

    let id = item["id"]
        .as_str()
        .unwrap_or("")
        .to_string();

    let amount = item["amount"].as_f64().ok_or_else(|| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: "Mercury transaction missing 'amount' field".into(),
        hint: None,
    })?;

    let kind = item["kind"].as_str().unwrap_or("unknown");
    let canonical_type = map_mercury_kind(kind, amount);

    let created_at = item["createdAt"]
        .as_str()
        .unwrap_or("");

    let posted_at = item["postedAt"]
        .as_str()
        .filter(|s| !s.is_empty());
    let posted_source = posted_at.unwrap_or(created_at);

    let bank_description = item["bankDescription"].as_str().unwrap_or("");
    let counterparty_name = item["counterpartyName"].as_str().unwrap_or("");
    let note = item["note"].as_str().unwrap_or("");

    let description = pick_description(bank_description, counterparty_name, note);

    Ok(Some(RawTransaction {
        created_iso: created_at.to_string(),
        posted_date: parse_iso_date(posted_source),
        amount_minor: amount_to_minor(amount),
        canonical_type: canonical_type.to_string(),
        source_id: id,
        description,
    }))
}

fn extract_mercury_error(body: &serde_json::Value, status: u16) -> String {
    body["message"]
        .as_str()
        .or_else(|| body["error"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_mercury(
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

    let client = MercuryClient::new(key);

    // 3. Resolve account
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    let account_id = match account {
        Some(id) => id,
        None => {
            let accounts = client.list_accounts()?;
            let active: Vec<_> = accounts.iter().filter(|a| a.status == "active").collect();
            match active.len() {
                0 => {
                    return Err(CliError::args(
                        "no active Mercury accounts found".to_string(),
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
                        "multiple Mercury accounts found; use --account to specify ({})",
                        names.join(", "),
                    )));
                }
            }
        }
    };

    // 4. Fetch
    if show_progress {
        eprintln!(
            "Fetching Mercury transactions ({} to {})...",
            from_date, to_date,
        );
        let display_id = if account_id.len() > 10 {
            format!("{}...", &account_id[..10])
        } else {
            account_id.clone()
        };
        eprintln!("  Account: {}", display_id);
    }

    let mut txns = client.fetch_transactions(&account_id, &from_date, &to_date, quiet)?;

    // 5. Sort: (created_iso ASC, source_id ASC)
    txns.sort_by(|a, b| {
        a.created_iso
            .cmp(&b.created_iso)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 6. Build canonical rows
    let rows: Vec<CanonicalRow> = txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: parse_iso_date(&txn.created_iso),
            posted_date: txn.posted_date.clone(),
            amount_minor: txn.amount_minor,
            currency: "USD".to_string(),
            r#type: txn.canonical_type.clone(),
            source: "mercury".to_string(),
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

fn resolve_api_key(flag: Option<String>) -> Result<String, CliError> {
    common::resolve_api_key(flag, "Mercury", "MERCURY_API_KEY")
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── Unit tests ──────────────────────────────────────────────────

    #[test]
    fn test_amount_to_minor() {
        assert_eq!(amount_to_minor(150.00), 15000);
        assert_eq!(amount_to_minor(-4.35), -435);
        assert_eq!(amount_to_minor(0.01), 1);
        assert_eq!(amount_to_minor(0.0), 0);
    }

    #[test]
    fn test_map_mercury_kind() {
        // externalTransfer — sign-dependent
        assert_eq!(map_mercury_kind("externalTransfer", 100.0), "deposit");
        assert_eq!(map_mercury_kind("externalTransfer", -100.0), "withdrawal");

        // fixed mappings
        assert_eq!(map_mercury_kind("internalTransfer", 50.0), "transfer");
        assert_eq!(map_mercury_kind("outgoingPayment", -200.0), "withdrawal");
        assert_eq!(map_mercury_kind("incomingDomesticWire", 500.0), "deposit");
        assert_eq!(map_mercury_kind("incomingInternationalWire", 1000.0), "deposit");
        assert_eq!(map_mercury_kind("debitCardTransaction", -25.0), "expense");
        assert_eq!(map_mercury_kind("cardTransaction", -10.0), "expense");
        assert_eq!(map_mercury_kind("fee", -5.0), "fee");

        // check — sign-dependent
        assert_eq!(map_mercury_kind("check", -50.0), "withdrawal");
        assert_eq!(map_mercury_kind("check", 50.0), "deposit");

        // unknown/other — sign-dependent fallback
        assert_eq!(map_mercury_kind("other", 100.0), "deposit");
        assert_eq!(map_mercury_kind("other", -50.0), "withdrawal");
        assert_eq!(map_mercury_kind("treasuryCredit", 100.0), "deposit");
        assert_eq!(map_mercury_kind("somethingNew", -50.0), "withdrawal");
    }

    #[test]
    fn test_parse_iso_date() {
        assert_eq!(parse_iso_date("2026-01-15T10:30:00Z"), "2026-01-15");
        assert_eq!(parse_iso_date("2026-12-31T23:59:59.000Z"), "2026-12-31");
        assert_eq!(parse_iso_date("2026-01-01"), "2026-01-01");
    }

    #[test]
    fn test_description_priority() {
        // bankDescription first
        assert_eq!(
            pick_description("Bank desc", "Counterparty", "Note"),
            "Bank desc",
        );
        // fall through to counterpartyName
        assert_eq!(
            pick_description("", "Counterparty", "Note"),
            "Counterparty",
        );
        // fall through to note
        assert_eq!(
            pick_description("", "", "Note"),
            "Note",
        );
        // all empty
        assert_eq!(
            pick_description("", "", ""),
            "",
        );
    }

    #[test]
    fn test_resolve_api_key_from_flag() {
        let key = resolve_api_key(Some("  secret-token:mercury_test  ".into())).unwrap();
        assert_eq!(key, "secret-token:mercury_test");
    }

    #[test]
    fn test_resolve_api_key_empty_flag() {
        let err = resolve_api_key(Some("  ".into())).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_resolve_api_key_missing() {
        std::env::remove_var("MERCURY_API_KEY");
        let err = resolve_api_key(None).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("missing Mercury API key"));
    }

    #[test]
    fn test_status_filtering() {
        let make_txn = |status: &str| -> serde_json::Value {
            serde_json::json!({
                "id": format!("txn_{}", status),
                "amount": 100.0,
                "kind": "externalTransfer",
                "status": status,
                "createdAt": "2026-01-15T10:00:00Z",
                "postedAt": null,
                "bankDescription": "Test",
                "counterpartyName": "",
                "note": ""
            })
        };

        // sent → included
        assert!(parse_transaction(&make_txn("sent")).unwrap().is_some());
        // pending → included
        assert!(parse_transaction(&make_txn("pending")).unwrap().is_some());
        // cancelled → excluded
        assert!(parse_transaction(&make_txn("cancelled")).unwrap().is_none());
        // failed → excluded
        assert!(parse_transaction(&make_txn("failed")).unwrap().is_none());
    }

    // ── httpmock tests ──────────────────────────────────────────────

    /// Helper: build a Mercury-shaped transaction JSON.
    fn mock_mercury_txn(id: &str, amount: f64, kind: &str) -> serde_json::Value {
        serde_json::json!({
            "id": id,
            "amount": amount,
            "kind": kind,
            "status": "sent",
            "createdAt": "2026-01-15T10:00:00Z",
            "postedAt": "2026-01-16T00:00:00Z",
            "bankDescription": format!("Txn {}", id),
            "counterpartyName": "",
            "counterpartyNickname": "",
            "externalMemo": "",
            "note": "",
            "feeId": null,
            "estimatedDeliveryDate": null,
            "failedAt": null,
            "reasonForFailure": null,
            "dashboardLink": ""
        })
    }

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        // Build 500 transactions for page 1
        let page1_txns: Vec<serde_json::Value> = (0..500)
            .map(|i| mock_mercury_txn(&format!("txn_{:04}", i), 10.0, "externalTransfer"))
            .collect();

        // Page 1: 500 items (full page)
        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/transactions")
                .query_param("offset", "0");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({ "transactions": page1_txns }));
        });

        // Build 47 transactions for page 2
        let page2_txns: Vec<serde_json::Value> = (500..547)
            .map(|i| mock_mercury_txn(&format!("txn_{:04}", i), 20.0, "externalTransfer"))
            .collect();

        // Page 2: 47 items (partial page → stop)
        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/transactions")
                .query_param("offset", "500");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({ "transactions": page2_txns }));
        });

        let client = MercuryClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_transactions("acc_123", &from, &to, true).unwrap();

        page1_mock.assert();
        page2_mock.assert();
        assert_eq!(txns.len(), 547);
    }

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/transactions");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Invalid token"
                }));
        });

        let client = MercuryClient::with_base_url(
            "bad_key".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_transactions("acc_123", &from, &to, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("Mercury auth failed (401)"),
            "message: {}",
            err.message,
        );
        assert!(
            err.message.contains("Invalid token"),
            "message: {}",
            err.message,
        );
    }

    #[test]
    fn test_unknown_kind_sign_fallback() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/transactions");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "transactions": [
                        {
                            "id": "txn_treasury",
                            "amount": 500.0,
                            "kind": "treasuryCredit",
                            "status": "sent",
                            "createdAt": "2026-01-15T10:00:00Z",
                            "postedAt": "2026-01-16T00:00:00Z",
                            "bankDescription": "Treasury deposit",
                            "counterpartyName": "",
                            "note": ""
                        },
                        {
                            "id": "txn_other_debit",
                            "amount": -75.0,
                            "kind": "other",
                            "status": "sent",
                            "createdAt": "2026-01-15T11:00:00Z",
                            "postedAt": "2026-01-16T00:00:00Z",
                            "bankDescription": "Some debit",
                            "counterpartyName": "",
                            "note": ""
                        }
                    ]
                }));
        });

        let client = MercuryClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_transactions("acc_123", &from, &to, true).unwrap();

        assert_eq!(txns.len(), 2);
        // Unknown positive → deposit
        assert_eq!(txns[0].canonical_type, "deposit");
        // "other" negative → withdrawal
        assert_eq!(txns[1].canonical_type, "withdrawal");
    }

    #[test]
    fn test_auto_detect_single_account() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/accounts");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "accounts": [
                        { "id": "acc_001", "name": "Checking", "status": "active" },
                        { "id": "acc_002", "name": "Closed Account", "status": "closed" }
                    ]
                }));
        });

        let client = MercuryClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let accounts = client.list_accounts().unwrap();
        let active: Vec<_> = accounts.iter().filter(|a| a.status == "active").collect();
        assert_eq!(active.len(), 1);
        assert_eq!(active[0].id, "acc_001");
        assert_eq!(active[0].name, "Checking");
    }

    #[test]
    fn test_auto_detect_multiple_errors() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/accounts");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({
                    "accounts": [
                        { "id": "acc_001", "name": "Checking", "status": "active" },
                        { "id": "acc_002", "name": "Savings", "status": "active" }
                    ]
                }));
        });

        let client = MercuryClient::with_base_url(
            "test_key".into(),
            server.base_url(),
        );

        let accounts = client.list_accounts().unwrap();
        let active: Vec<_> = accounts.iter().filter(|a| a.status == "active").collect();
        assert_eq!(active.len(), 2);

        // Verify the error message that cmd_fetch_mercury would produce
        let names: Vec<_> = active.iter().map(|a| a.name.as_str()).collect();
        let msg = format!(
            "multiple Mercury accounts found; use --account to specify ({})",
            names.join(", "),
        );
        assert!(msg.contains("Checking"));
        assert!(msg.contains("Savings"));
    }
}
