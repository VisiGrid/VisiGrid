//! `vgrid fetch authorizenet` — fetch Authorize.net settled transactions into canonical CSV.

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const API_BASE_PROD: &str = "https://api.authorize.net/xml/v1/request.api";
const API_BASE_SANDBOX: &str = "https://apitest.authorize.net/xml/v1/request.api";

// ── Type mapping ────────────────────────────────────────────────────

fn map_txn_type(authnet_type: &str) -> &'static str {
    match authnet_type {
        "authCaptureTransaction" | "captureOnlyTransaction" => "charge",
        "refundTransaction" => "refund",
        "voidTransaction" => "void",
        "priorAuthCaptureTransaction" => "capture",
        _ => "adjustment",
    }
}

// ── Internal transaction representation ─────────────────────────────

#[derive(Debug)]
struct RawTransaction {
    trans_id: String,
    effective_date: String,
    posted_date: String,
    amount_minor: i64,
    canonical_type: String,
    batch_id: String,
    description: String,
}

// ── Client ──────────────────────────────────────────────────────────

struct AuthorizeNetClient {
    client: FetchClient,
    api_login_id: String,
    transaction_key: String,
    api_url: String,
}

impl AuthorizeNetClient {
    fn new(api_login_id: String, transaction_key: String, sandbox: bool) -> Self {
        let api_url = if sandbox {
            API_BASE_SANDBOX
        } else {
            API_BASE_PROD
        };
        Self {
            client: FetchClient::new("Authorize.net", extract_error),
            api_login_id,
            transaction_key,
            api_url: api_url.to_string(),
        }
    }

    #[cfg(test)]
    fn with_base_url(api_login_id: String, transaction_key: String, api_url: String) -> Self {
        Self {
            client: FetchClient::new("Authorize.net", extract_error),
            api_login_id,
            transaction_key,
            api_url,
        }
    }

    fn auth_block(&self) -> serde_json::Value {
        serde_json::json!({
            "name": self.api_login_id,
            "transactionKey": self.transaction_key,
        })
    }

    /// Fetch settled batch list for a date range.
    fn fetch_batch_list(
        &self,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        // Authorize.net XML schema requires merchantAuthentication first.
        // serde_json::json! uses BTreeMap (alphabetical), so build manually.
        let mut inner = serde_json::Map::new();
        inner.insert("merchantAuthentication".into(), self.auth_block());
        inner.insert("firstSettlementDate".into(), format!("{}T00:00:00Z", from).into());
        inner.insert("lastSettlementDate".into(), format!("{}T00:00:00Z", to).into());
        let mut payload = serde_json::Map::new();
        payload.insert("getSettledBatchListRequest".into(), serde_json::Value::Object(inner));
        let payload = serde_json::Value::Object(payload);

        let url = self.api_url.clone();
        let body = self.client.request_with_retry(|http| {
            http.post(&url)
                .header("Content-Type", "application/json")
                .json(&payload)
        })?;

        check_authnet_result(&body)?;

        let batches = body["batchList"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        let show_progress = !quiet && atty::is(atty::Stream::Stderr);
        if show_progress {
            eprintln!("  found {} settled batches", batches.len());
        }

        Ok(batches)
    }

    /// Fetch transactions for a single batch.
    fn fetch_transaction_list(
        &self,
        batch_id: &str,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        // Authorize.net XML schema requires merchantAuthentication first.
        let mut inner = serde_json::Map::new();
        inner.insert("merchantAuthentication".into(), self.auth_block());
        inner.insert("batchId".into(), batch_id.into());
        let mut payload = serde_json::Map::new();
        payload.insert("getTransactionListRequest".into(), serde_json::Value::Object(inner));
        let payload = serde_json::Value::Object(payload);

        let url = self.api_url.clone();
        let body = self.client.request_with_retry(|http| {
            http.post(&url)
                .header("Content-Type", "application/json")
                .json(&payload)
        })?;

        check_authnet_result(&body)?;

        let txns = body["transactions"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        let show_progress = !quiet && atty::is(atty::Stream::Stderr);
        if show_progress {
            eprintln!("  batch {}: {} transactions", batch_id, txns.len());
        }

        Ok(txns)
    }
}

/// Check the Authorize.net `resultCode` and return an error if it's not "Ok".
fn check_authnet_result(body: &serde_json::Value) -> Result<(), CliError> {
    let messages = &body["messages"];
    let result_code = messages["resultCode"].as_str().unwrap_or("");

    if result_code == "Ok" {
        return Ok(());
    }

    let msg = messages["message"]
        .as_array()
        .and_then(|arr| arr.first())
        .and_then(|m| m["text"].as_str())
        .unwrap_or("unknown error");

    let code_str = messages["message"]
        .as_array()
        .and_then(|arr| arr.first())
        .and_then(|m| m["code"].as_str())
        .unwrap_or("");

    // E00007 = auth error
    let exit_code = if code_str == "E00007" {
        exit_codes::EXIT_FETCH_AUTH
    } else {
        exit_codes::EXIT_FETCH_UPSTREAM
    };

    Err(CliError {
        code: exit_code,
        message: format!("Authorize.net error ({}): {}", code_str, msg),
        hint: None,
    })
}

fn extract_error(body: &serde_json::Value, status: u16) -> String {
    body["messages"]["message"]
        .as_array()
        .and_then(|arr| arr.first())
        .and_then(|m| m["text"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

/// Parse a datetime string like "2026-01-15T12:34:56Z" or "2026-01-15T12:34:56" to a date.
fn parse_datetime_to_date(s: &str) -> String {
    // Try full ISO 8601 with Z
    if let Ok(dt) = chrono::DateTime::parse_from_rfc3339(s) {
        return dt.format("%Y-%m-%d").to_string();
    }
    // Try without timezone
    if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
        return dt.format("%Y-%m-%d").to_string();
    }
    // Try date only
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return d.to_string();
    }
    s.to_string()
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_authorizenet(
    from: String,
    to: String,
    api_login_id: Option<String>,
    transaction_key: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
    sandbox: bool,
) -> Result<(), CliError> {
    // 1. Resolve credentials
    let login_id = common::resolve_api_key(
        api_login_id,
        "Authorize.net",
        "AUTHORIZENET_API_LOGIN_ID",
    )?;
    let txn_key = common::resolve_api_key(
        transaction_key,
        "Authorize.net",
        "AUTHORIZENET_TRANSACTION_KEY",
    )?;

    // 2. Parse dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Authorize.net settled transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let client = AuthorizeNetClient::new(login_id, txn_key, sandbox);

    // Phase 1: get batch list
    let batches = client.fetch_batch_list(&from_date, &to_date, quiet)?;

    // Phase 2: get transactions per batch
    let mut all_txns: Vec<RawTransaction> = Vec::new();

    for batch in &batches {
        let batch_id = batch["batchId"].as_str().unwrap_or("").to_string();
        let settlement_time = batch["settlementTimeUTC"]
            .as_str()
            .or_else(|| batch["settlementTimeLocal"].as_str())
            .unwrap_or("");
        let posted_date = parse_datetime_to_date(settlement_time);

        if batch_id.is_empty() {
            continue;
        }

        let txns = client.fetch_transaction_list(&batch_id, quiet)?;

        for txn in &txns {
            let trans_id = txn["transId"].as_str().unwrap_or("").to_string();
            let submit_time = txn["submitTimeUTC"]
                .as_str()
                .or_else(|| txn["submitTimeLocal"].as_str())
                .unwrap_or("");
            let effective_date = parse_datetime_to_date(submit_time);

            let settle_amount = txn["settleAmount"]
                .as_str()
                .or_else(|| txn["settleAmount"].as_f64().map(|_| ""))
                .unwrap_or("0");

            // settleAmount may be a number or string
            let amount_str = if let Some(n) = txn["settleAmount"].as_f64() {
                format!("{:.2}", n)
            } else {
                txn["settleAmount"]
                    .as_str()
                    .unwrap_or("0")
                    .to_string()
            };

            let amount_minor = common::parse_money_string(&amount_str).map_err(|e| {
                CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: format!(
                        "Authorize.net bad settleAmount {:?} for txn {}: {}",
                        settle_amount, trans_id, e,
                    ),
                    hint: None,
                }
            })?;

            let raw_type = txn["transactionType"].as_str().unwrap_or("");
            let canonical_type = map_txn_type(raw_type).to_string();

            // Refunds and voids are negative
            let signed_amount = if canonical_type == "refund" || canonical_type == "void" {
                -amount_minor.abs()
            } else {
                amount_minor.abs()
            };

            let account_type = txn["accountType"].as_str().unwrap_or("");
            let account_number = txn["accountNumber"].as_str().unwrap_or("");
            let description = if !account_type.is_empty() || !account_number.is_empty() {
                format!("{} {}", account_type, account_number).trim().to_string()
            } else {
                String::new()
            };

            all_txns.push(RawTransaction {
                trans_id,
                effective_date,
                posted_date: posted_date.clone(),
                amount_minor: signed_amount,
                canonical_type,
                batch_id: batch_id.clone(),
                description,
            });
        }
    }

    // 4. Sort: (effective_date, trans_id)
    all_txns.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.trans_id.cmp(&b.trans_id))
    });

    // 5. Build canonical rows
    let rows: Vec<CanonicalRow> = all_txns
        .iter()
        .map(|txn| CanonicalRow {
            effective_date: txn.effective_date.clone(),
            posted_date: txn.posted_date.clone(),
            amount_minor: txn.amount_minor,
            currency: "USD".to_string(),
            r#type: txn.canonical_type.clone(),
            source: "authorizenet".to_string(),
            source_id: txn.trans_id.clone(),
            group_id: txn.batch_id.clone(),
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
    fn test_map_txn_type() {
        assert_eq!(map_txn_type("authCaptureTransaction"), "charge");
        assert_eq!(map_txn_type("captureOnlyTransaction"), "charge");
        assert_eq!(map_txn_type("refundTransaction"), "refund");
        assert_eq!(map_txn_type("voidTransaction"), "void");
        assert_eq!(map_txn_type("priorAuthCaptureTransaction"), "capture");
        assert_eq!(map_txn_type("something_else"), "adjustment");
    }

    #[test]
    fn test_parse_datetime_to_date() {
        assert_eq!(
            parse_datetime_to_date("2026-01-15T12:34:56Z"),
            "2026-01-15"
        );
        assert_eq!(
            parse_datetime_to_date("2026-01-15T12:34:56"),
            "2026-01-15"
        );
        assert_eq!(parse_datetime_to_date("2026-01-15"), "2026-01-15");
    }

    fn authnet_success(data: serde_json::Value) -> serde_json::Value {
        let mut result = serde_json::json!({
            "messages": {
                "resultCode": "Ok",
                "message": [{"code": "I00001", "text": "Successful."}]
            }
        });
        // Merge data into result
        if let (Some(r), Some(d)) = (result.as_object_mut(), data.as_object()) {
            for (k, v) in d {
                r.insert(k.clone(), v.clone());
            }
        }
        result
    }

    #[test]
    fn test_fetch_batch_list() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(POST).path("/");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(authnet_success(serde_json::json!({
                    "batchList": [
                        {
                            "batchId": "12345",
                            "settlementTimeUTC": "2026-01-15T10:00:00Z",
                            "settlementState": "settledSuccessfully"
                        }
                    ]
                })));
        });

        let client = AuthorizeNetClient::with_base_url(
            "login".into(),
            "key".into(),
            format!("{}/", server.base_url()),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let batches = client.fetch_batch_list(&from, &to, true).unwrap();

        assert_eq!(batches.len(), 1);
        assert_eq!(batches[0]["batchId"].as_str().unwrap(), "12345");
    }

    #[test]
    fn test_fetch_transaction_list() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(POST).path("/");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(authnet_success(serde_json::json!({
                    "transactions": [
                        {
                            "transId": "111",
                            "submitTimeUTC": "2026-01-15T09:00:00Z",
                            "settleAmount": 50.00,
                            "transactionType": "authCaptureTransaction",
                            "accountType": "Visa",
                            "accountNumber": "XXXX1234"
                        }
                    ]
                })));
        });

        let client = AuthorizeNetClient::with_base_url(
            "login".into(),
            "key".into(),
            format!("{}/", server.base_url()),
        );

        let txns = client.fetch_transaction_list("12345", true).unwrap();
        assert_eq!(txns.len(), 1);
        assert_eq!(txns[0]["transId"].as_str().unwrap(), "111");
    }

    #[test]
    fn test_authnet_error_response() {
        let body = serde_json::json!({
            "messages": {
                "resultCode": "Error",
                "message": [{"code": "E00007", "text": "User authentication failed."}]
            }
        });
        let err = check_authnet_result(&body).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(err.message.contains("User authentication failed"));
    }

    #[test]
    fn test_refund_amount_negative() {
        let amount: i64 = 5000;
        let canonical_type = "refund";
        let signed = if canonical_type == "refund" || canonical_type == "void" {
            -amount.abs()
        } else {
            amount.abs()
        };
        assert_eq!(signed, -5000);
    }

    #[test]
    fn test_resolve_credentials() {
        std::env::remove_var("AUTHORIZENET_API_LOGIN_ID");
        let err = common::resolve_api_key(
            None,
            "Authorize.net",
            "AUTHORIZENET_API_LOGIN_ID",
        )
        .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }
}
