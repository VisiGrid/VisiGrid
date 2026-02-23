//! `vgrid fetch netsuite` — fetch general ledger transactions from NetSuite via SuiteTalk REST.

use std::path::PathBuf;

use chrono::NaiveDate;
use hmac::{Hmac, Mac};
use sha2::Sha256;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Type mapping ────────────────────────────────────────────────────

fn map_txn_type(ns_type: &str) -> &'static str {
    match ns_type {
        "CustInvc" | "Invoice" => "charge",
        "CustCred" | "CreditMemo" => "refund",
        "CustPymt" | "Payment" => "payment",
        "Journal" | "JrnlEntr" => "journal",
        "VendBill" | "Bill" => "bill",
        "VendPymt" | "VendorPayment" => "vendor_payment",
        "Deposit" => "deposit",
        "Transfer" => "transfer",
        "Check" => "check",
        _ => "adjustment",
    }
}

// ── OAuth 1.0 TBA signing ───────────────────────────────────────────

fn percent_encode(s: &str) -> String {
    let mut encoded = String::with_capacity(s.len() * 2);
    for byte in s.bytes() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' => {
                encoded.push(byte as char);
            }
            _ => {
                encoded.push_str(&format!("%{:02X}", byte));
            }
        }
    }
    encoded
}

fn oauth_header(
    method: &str,
    url: &str,
    account_id: &str,
    consumer_key: &str,
    consumer_secret: &str,
    token_id: &str,
    token_secret: &str,
) -> String {
    let nonce: String = (0..32)
        .map(|_| format!("{:x}", rand::random::<u8>() % 16))
        .collect();
    let timestamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_secs()
        .to_string();

    let mut params = vec![
        ("oauth_consumer_key", consumer_key.to_string()),
        ("oauth_token", token_id.to_string()),
        ("oauth_nonce", nonce.clone()),
        ("oauth_timestamp", timestamp.clone()),
        ("oauth_signature_method", "HMAC-SHA256".to_string()),
        ("oauth_version", "1.0".to_string()),
    ];
    params.sort_by(|a, b| a.0.cmp(&b.0));

    let sorted_str: String = params
        .iter()
        .map(|(k, v)| format!("{}={}", percent_encode(k), percent_encode(v)))
        .collect::<Vec<_>>()
        .join("&");

    let base_string = format!(
        "{}&{}&{}",
        method.to_uppercase(),
        percent_encode(url),
        percent_encode(&sorted_str)
    );

    let signing_key = format!(
        "{}&{}",
        percent_encode(consumer_secret),
        percent_encode(token_secret)
    );

    type HmacSha256 = Hmac<Sha256>;
    let mut mac = HmacSha256::new_from_slice(signing_key.as_bytes()).unwrap();
    mac.update(base_string.as_bytes());
    let signature = base64::Engine::encode(
        &base64::engine::general_purpose::STANDARD,
        mac.finalize().into_bytes(),
    );

    let header_params: String = [
        ("oauth_consumer_key", consumer_key),
        ("oauth_nonce", &nonce),
        ("oauth_signature", &signature),
        ("oauth_signature_method", "HMAC-SHA256"),
        ("oauth_timestamp", &timestamp),
        ("oauth_token", token_id),
        ("oauth_version", "1.0"),
    ]
    .iter()
    .map(|(k, v)| format!("{}=\"{}\"", percent_encode(k), percent_encode(v)))
    .collect::<Vec<_>>()
    .join(", ");

    format!("OAuth realm=\"{}\", {}", account_id, header_params)
}

// ── Client ──────────────────────────────────────────────────────────

struct NetSuiteClient {
    client: FetchClient,
    account_id: String,
    consumer_key: String,
    consumer_secret: String,
    token_id: String,
    token_secret: String,
    base_url: String,
}

impl NetSuiteClient {
    fn new(
        account_id: String,
        consumer_key: String,
        consumer_secret: String,
        token_id: String,
        token_secret: String,
    ) -> Self {
        let normalized = account_id.to_lowercase().replace('_', "-");
        let base_url = format!("https://{}.suitetalk.api.netsuite.com", normalized);
        Self {
            client: FetchClient::new("NetSuite", extract_error),
            account_id: normalized,
            consumer_key,
            consumer_secret,
            token_id,
            token_secret,
            base_url,
        }
    }

    #[cfg(test)]
    fn with_base_url(
        account_id: String,
        consumer_key: String,
        consumer_secret: String,
        token_id: String,
        token_secret: String,
        base_url: String,
    ) -> Self {
        let normalized = account_id.to_lowercase().replace('_', "-");
        Self {
            client: FetchClient::new("NetSuite", extract_error),
            account_id: normalized,
            consumer_key,
            consumer_secret,
            token_id,
            token_secret,
            base_url,
        }
    }

    fn suiteql_query(
        &self,
        query: &str,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        let url = format!("{}/services/rest/query/v1/suiteql", self.base_url);
        let mut all_items = Vec::new();
        let mut offset = 0;
        let limit = 1000;

        loop {
            let paged_url = format!("{}?limit={}&offset={}", url, limit, offset);
            let auth = oauth_header(
                "POST",
                &paged_url,
                &self.account_id,
                &self.consumer_key,
                &self.consumer_secret,
                &self.token_id,
                &self.token_secret,
            );

            let payload = serde_json::json!({ "q": query });
            let auth_clone = auth.clone();
            let body = self.client.request_with_retry(|http| {
                http.post(&paged_url)
                    .header("Authorization", &auth_clone)
                    .header("Content-Type", "application/json")
                    .header("Accept", "application/json")
                    .header("Prefer", "transient")
                    .json(&payload)
            })?;

            let items = body["items"]
                .as_array()
                .cloned()
                .unwrap_or_default();

            let show_progress = !quiet && atty::is(atty::Stream::Stderr);
            if show_progress {
                eprintln!("  fetched {} records (offset {})", items.len(), offset);
            }

            let count = items.len();
            all_items.extend(items);

            if body["hasMore"].as_bool().unwrap_or(false) && count > 0 {
                offset += count;
            } else {
                break;
            }
        }

        Ok(all_items)
    }
}

fn extract_error(body: &serde_json::Value, status: u16) -> String {
    body["o:errorDetails"]
        .as_array()
        .and_then(|arr| arr.first())
        .and_then(|e| e["detail"].as_str())
        .or_else(|| body["title"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_netsuite(
    from: String,
    to: String,
    account_id: Option<String>,
    consumer_key: Option<String>,
    consumer_secret: Option<String>,
    token_id: Option<String>,
    token_secret: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Resolve credentials
    let account_id = common::resolve_api_key(account_id, "NetSuite", "NETSUITE_ACCOUNT_ID")?;
    let consumer_key = common::resolve_api_key(consumer_key, "NetSuite", "NETSUITE_CONSUMER_KEY")?;
    let consumer_secret =
        common::resolve_api_key(consumer_secret, "NetSuite", "NETSUITE_CONSUMER_SECRET")?;
    let token_id = common::resolve_api_key(token_id, "NetSuite", "NETSUITE_TOKEN_ID")?;
    let token_secret =
        common::resolve_api_key(token_secret, "NetSuite", "NETSUITE_TOKEN_SECRET")?;

    // 2. Parse dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching NetSuite transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let client = NetSuiteClient::new(
        account_id,
        consumer_key,
        consumer_secret,
        token_id,
        token_secret,
    );

    let query = format!(
        "SELECT id, trandate, postingperiod, amount, currency, type, tranid, memo \
         FROM transaction \
         WHERE trandate >= '{}' AND trandate < '{}'",
        from_date, to_date,
    );

    let items = client.suiteql_query(&query, quiet)?;

    // 4. Build canonical rows
    let mut rows: Vec<CanonicalRow> = items
        .iter()
        .map(|item| {
            let id = item["id"]
                .as_str()
                .or_else(|| item["id"].as_i64().map(|_| ""))
                .unwrap_or("");
            let id_str = if let Some(n) = item["id"].as_i64() {
                n.to_string()
            } else {
                id.to_string()
            };

            let trandate = item["trandate"].as_str().unwrap_or("").to_string();
            let amount_str = if let Some(n) = item["amount"].as_f64() {
                format!("{:.2}", n)
            } else {
                item["amount"].as_str().unwrap_or("0").to_string()
            };
            let amount_minor = common::parse_money_string(&amount_str).unwrap_or(0);

            let currency = item["currency"]
                .as_str()
                .map(|s| s.to_string())
                .unwrap_or_else(|| "USD".to_string());

            let ns_type = item["type"].as_str().unwrap_or("");
            let tranid = item["tranid"].as_str().unwrap_or("").to_string();
            let memo = item["memo"].as_str().unwrap_or("").to_string();

            CanonicalRow {
                effective_date: trandate.clone(),
                posted_date: trandate,
                amount_minor,
                currency,
                r#type: map_txn_type(ns_type).to_string(),
                source: "netsuite".to_string(),
                source_id: id_str,
                group_id: tranid,
                description: memo,
            }
        })
        .collect();

    // 5. Sort
    rows.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 6. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!("Done: {} transactions written to {}", rows.len(), out_label);
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_map_txn_type() {
        assert_eq!(map_txn_type("CustInvc"), "charge");
        assert_eq!(map_txn_type("CustCred"), "refund");
        assert_eq!(map_txn_type("Journal"), "journal");
        assert_eq!(map_txn_type("VendBill"), "bill");
        assert_eq!(map_txn_type("Deposit"), "deposit");
        assert_eq!(map_txn_type("unknown"), "adjustment");
    }

    #[test]
    fn test_percent_encode() {
        assert_eq!(percent_encode("hello world"), "hello%20world");
        assert_eq!(percent_encode("foo=bar&baz"), "foo%3Dbar%26baz");
        assert_eq!(percent_encode("simple"), "simple");
    }

    #[test]
    fn test_oauth_header_format() {
        let header = oauth_header(
            "GET",
            "https://test.suitetalk.api.netsuite.com/services/rest/record/v1/metadata-catalog/",
            "test-account",
            "consumer_key",
            "consumer_secret",
            "token_id",
            "token_secret",
        );
        assert!(header.starts_with("OAuth realm=\"test-account\""));
        assert!(header.contains("oauth_consumer_key="));
        assert!(header.contains("oauth_signature="));
        assert!(header.contains("oauth_signature_method=\"HMAC-SHA256\""));
    }
}
