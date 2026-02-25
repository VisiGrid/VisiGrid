//! `vgrid fetch fiserv` — fetch Fiserv/CardPointe settled transactions into canonical CSV.

use std::path::PathBuf;

use chrono::NaiveDate;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const DEFAULT_API_URL: &str = "https://fts.cardconnect.com";

// ── Type mapping ────────────────────────────────────────────────────

fn map_txn_type(fiserv_type: &str) -> &'static str {
    match fiserv_type.to_lowercase().as_str() {
        "sale" => "charge",
        "refund" => "refund",
        "void" => "void",
        _ => "adjustment",
    }
}

// ── Internal transaction representation ─────────────────────────────

#[derive(Debug)]
struct RawTransaction {
    retref: String,
    effective_date: String,
    posted_date: String,
    amount_minor: i64,
    canonical_type: String,
    batch_id: String,
    description: String,
}

// ── Date helpers ────────────────────────────────────────────────────

/// Convert MMDDYYYY to YYYY-MM-DD.
fn parse_authdate(s: &str) -> String {
    if s.len() == 8 {
        format!("{}-{}-{}", &s[4..8], &s[0..2], &s[2..4])
    } else {
        s.to_string()
    }
}

/// Format a NaiveDate as MMDD for the Fiserv API query parameter.
fn date_to_mmdd(d: &NaiveDate) -> String {
    d.format("%m%d").to_string()
}

// ── Client ──────────────────────────────────────────────────────────

struct FiservClient {
    client: FetchClient,
    api_url: String,
    merchant_id: String,
    api_username: String,
    api_password: String,
}

impl FiservClient {
    fn new(
        api_url: String,
        merchant_id: String,
        api_username: String,
        api_password: String,
    ) -> Self {
        Self {
            client: FetchClient::new("Fiserv", extract_error),
            api_url,
            merchant_id,
            api_username,
            api_password,
        }
    }

    /// Fetch settlement status for a single date.
    fn fetch_settlestat(
        &self,
        date: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        let mmdd = date_to_mmdd(date);
        let url = format!(
            "{}/cardconnect/rest/settlestat?merchid={}&date={}",
            self.api_url, self.merchant_id, mmdd,
        );

        let username = self.api_username.clone();
        let password = self.api_password.clone();
        let url_clone = url.clone();

        let body = self.client.request_with_retry_text(|http| {
            http.get(&url_clone)
                .basic_auth(&username, Some(&password))
        })?;

        let show_progress = !quiet && atty::is(atty::Stream::Stderr);

        // CardPointe returns empty string, "Null batches", or other non-JSON
        // for dates with no settlements
        let trimmed = body.trim();
        if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("null batches") {
            if show_progress {
                eprintln!("  {}: no settlement data", date);
            }
            return Ok(Vec::new());
        }

        let parsed: serde_json::Value = serde_json::from_str(trimmed).map_err(|e| {
            CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message: format!(
                    "failed to parse Fiserv settlestat response for {}: {} (body: {})",
                    date, e, &trimmed[..trimmed.len().min(200)],
                ),
                hint: None,
            }
        })?;

        // Response is an array of batch objects
        let batches = if let Some(arr) = parsed.as_array() {
            arr.clone()
        } else if parsed.is_object() {
            // Single batch object or error
            if parsed.get("txns").is_some() {
                vec![parsed]
            } else {
                // Possibly an error or empty response
                if show_progress {
                    eprintln!("  {}: no settlement data", date);
                }
                return Ok(Vec::new());
            }
        } else {
            return Ok(Vec::new());
        };

        if show_progress {
            let txn_count: usize = batches
                .iter()
                .filter_map(|b| b["txns"].as_array())
                .map(|t| t.len())
                .sum();
            eprintln!("  {}: {} batches, {} transactions", date, batches.len(), txn_count);
        }

        Ok(batches)
    }
}

fn extract_error(body: &serde_json::Value, status: u16) -> String {
    body["resptext"]
        .as_str()
        .or_else(|| body["message"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_fiserv(
    from: String,
    to: String,
    api_url: Option<String>,
    merchant_id: Option<String>,
    api_username: Option<String>,
    api_password: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
    group_by_batch: bool,
) -> Result<(), CliError> {
    // 1. Resolve credentials
    let url = api_url
        .or_else(|| std::env::var("FISERV_API_URL").ok())
        .unwrap_or_else(|| DEFAULT_API_URL.to_string())
        .trim_end_matches('/')
        .to_string();

    let mid = common::resolve_api_key(
        merchant_id,
        "Fiserv",
        "FISERV_MERCHANT_ID",
    )?;
    let username = common::resolve_api_key(
        api_username,
        "Fiserv",
        "FISERV_API_USERNAME",
    )?;
    let password = common::resolve_api_key(
        api_password,
        "Fiserv",
        "FISERV_API_PASSWORD",
    )?;

    // 2. Parse dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Fiserv/CardPointe settled transactions ({} to {})...",
            from_date, to_date,
        );
    }

    let client = FiservClient::new(url, mid, username, password);

    let mut all_txns: Vec<RawTransaction> = Vec::new();

    // Iterate each date in the range
    let mut current = from_date;
    while current < to_date {
        let batches = client.fetch_settlestat(&current, quiet)?;

        for batch in &batches {
            let batch_id = batch["batchid"]
                .as_str()
                .unwrap_or("")
                .to_string();

            let txns = batch["txns"].as_array();
            if let Some(txns) = txns {
                for txn in txns {
                    let retref = txn["retref"]
                        .as_str()
                        .unwrap_or("")
                        .to_string();

                    // settlestat uses the query date as the settlement date;
                    // authdate may be present (MMDDYYYY) or absent.
                    let effective_date = txn["authdate"]
                        .as_str()
                        .filter(|s| !s.is_empty())
                        .map(|s| parse_authdate(s))
                        .unwrap_or_else(|| current.to_string());

                    let posted_date = current.to_string();

                    // settlestat returns "setlamount" (not "amount")
                    let amount_field = txn["setlamount"]
                        .as_str()
                        .or_else(|| txn["amount"].as_str());
                    let amount_value = if let Some(s) = amount_field {
                        s.to_string()
                    } else if let Some(n) = txn["setlamount"].as_f64().or_else(|| txn["amount"].as_f64()) {
                        format!("{:.2}", n)
                    } else {
                        "0".to_string()
                    };

                    let amount_minor =
                        common::parse_money_string(&amount_value).map_err(|e| CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "Fiserv bad amount {:?} for txn {}: {}",
                                amount_value, retref, e,
                            ),
                            hint: None,
                        })?;

                    // settlestat "setlstat" is Y/N (settled or not), not a txn type.
                    // Use "type" if present, otherwise infer from amount sign or
                    // default to "charge" for settled transactions.
                    let canonical_type = if let Some(t) = txn["type"].as_str() {
                        map_txn_type(t).to_string()
                    } else if amount_minor < 0 {
                        "refund".to_string()
                    } else {
                        "charge".to_string()
                    };

                    // Refunds and voids are negative
                    let signed_amount = if canonical_type == "refund" || canonical_type == "void" {
                        -amount_minor.abs()
                    } else {
                        amount_minor.abs()
                    };

                    // Description: card type + salesdoc (order ref) or last 4 of token
                    let cardtype = txn["cardtype"].as_str().unwrap_or("");
                    let salesdoc = txn["salesdoc"].as_str().unwrap_or("");
                    let token = txn["token"].as_str().unwrap_or("");
                    let last4 = if token.len() >= 4 {
                        &token[token.len() - 4..]
                    } else {
                        token
                    };
                    // Prefer salesdoc (order ref) for description, fall back to card+last4
                    let description = if !salesdoc.is_empty() {
                        format!("{} {}", cardtype, salesdoc).trim().to_string()
                    } else if !cardtype.is_empty() || !last4.is_empty() {
                        format!("{} {}", cardtype, last4).trim().to_string()
                    } else {
                        String::new()
                    };

                    all_txns.push(RawTransaction {
                        retref,
                        effective_date,
                        posted_date,
                        amount_minor: signed_amount,
                        canonical_type,
                        batch_id: batch_id.clone(),
                        description,
                    });
                }
            }
        }

        current = current.succ_opt().unwrap_or(current);
    }

    // 4. Build canonical rows
    let rows: Vec<CanonicalRow> = if group_by_batch {
        // Aggregate transactions by (posted_date, batch_id) → one row per batch deposit
        use std::collections::BTreeMap;
        let mut batches: BTreeMap<(String, String), (i64, usize)> = BTreeMap::new();
        for txn in &all_txns {
            let key = (txn.posted_date.clone(), txn.batch_id.clone());
            let entry = batches.entry(key).or_insert((0, 0));
            entry.0 += txn.amount_minor;
            entry.1 += 1;
        }
        batches
            .into_iter()
            .map(|((date, batch_id), (net, count))| CanonicalRow {
                effective_date: date.clone(),
                posted_date: date,
                amount_minor: net,
                currency: "USD".to_string(),
                r#type: "deposit".to_string(),
                source: "fiserv".to_string(),
                source_id: format!("batch:{}", batch_id),
                group_id: batch_id,
                description: format!("{} transactions", count),
            })
            .collect()
    } else {
        // Individual transactions (default)
        all_txns.sort_by(|a, b| {
            a.effective_date
                .cmp(&b.effective_date)
                .then_with(|| a.retref.cmp(&b.retref))
        });
        all_txns
            .iter()
            .map(|txn| CanonicalRow {
                effective_date: txn.effective_date.clone(),
                posted_date: txn.posted_date.clone(),
                amount_minor: txn.amount_minor,
                currency: "USD".to_string(),
                r#type: txn.canonical_type.clone(),
                source: "fiserv".to_string(),
                source_id: txn.retref.clone(),
                group_id: txn.batch_id.clone(),
                description: txn.description.clone(),
            })
            .collect()
    };

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
        assert_eq!(map_txn_type("sale"), "charge");
        assert_eq!(map_txn_type("Sale"), "charge");
        assert_eq!(map_txn_type("refund"), "refund");
        assert_eq!(map_txn_type("Refund"), "refund");
        assert_eq!(map_txn_type("void"), "void");
        assert_eq!(map_txn_type("something"), "adjustment");
    }

    #[test]
    fn test_parse_authdate() {
        assert_eq!(parse_authdate("01152026"), "2026-01-15");
        assert_eq!(parse_authdate("12312025"), "2025-12-31");
        assert_eq!(parse_authdate("short"), "short");
    }

    #[test]
    fn test_date_to_mmdd() {
        let d = NaiveDate::from_ymd_opt(2026, 1, 5).unwrap();
        assert_eq!(date_to_mmdd(&d), "0105");
        let d = NaiveDate::from_ymd_opt(2026, 12, 31).unwrap();
        assert_eq!(date_to_mmdd(&d), "1231");
    }

    #[test]
    fn test_fetch_settlestat() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/cardconnect/rest/settlestat")
                .query_param("merchid", "merchant123")
                .query_param("date", "0115");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!([
                    {
                        "batchid": "B001",
                        "txns": [
                            {
                                "retref": "RR001",
                                "setlamount": "50.00",
                                "setlstat": "Y",
                                "salesdoc": "O2026011509275293EB",
                                "cardtype": "VISA"
                            }
                        ]
                    }
                ]));
        });

        let client = FiservClient::new(
            server.base_url(),
            "merchant123".into(),
            "user".into(),
            "pass".into(),
        );

        let date = NaiveDate::from_ymd_opt(2026, 1, 15).unwrap();
        let batches = client.fetch_settlestat(&date, true).unwrap();

        assert_eq!(batches.len(), 1);
        assert_eq!(batches[0]["batchid"].as_str().unwrap(), "B001");
        let txns = batches[0]["txns"].as_array().unwrap();
        assert_eq!(txns.len(), 1);
        assert_eq!(txns[0]["retref"].as_str().unwrap(), "RR001");
    }

    #[test]
    fn test_refund_amount_negative() {
        let amount: i64 = 2500;
        let canonical_type = "refund";
        let signed = if canonical_type == "refund" || canonical_type == "void" {
            -amount.abs()
        } else {
            amount.abs()
        };
        assert_eq!(signed, -2500);
    }

    #[test]
    fn test_description_formatting() {
        let cardtype = "VISA";
        let token = "9876543210001234";
        let last4 = &token[token.len() - 4..];
        let desc = format!("{} {}", cardtype, last4).trim().to_string();
        assert_eq!(desc, "VISA 1234");
    }

    #[test]
    fn test_resolve_credentials() {
        std::env::remove_var("FISERV_MERCHANT_ID");
        let err = common::resolve_api_key(
            None,
            "Fiserv",
            "FISERV_MERCHANT_ID",
        )
        .unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }
}
