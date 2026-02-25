//! `vgrid fetch fiserv` — fetch Fiserv/CardPointe settled transactions into canonical CSV.

use std::io::Write;
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

// ── Funding deposit representation ──────────────────────────────────

#[derive(Debug)]
struct FundingDeposit {
    funding_id: String,
    funding_date: String,
    amount_minor: i64,
    merchant_id: String,
    currency: String,
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

/// Format a NaiveDate as YYYYMMDD for the Fiserv funding API.
fn date_to_yyyymmdd(d: &NaiveDate) -> String {
    d.format("%Y%m%d").to_string()
}

/// Case-insensitive key lookup on a JSON object.
fn get_ci<'a>(obj: &'a serde_json::Value, candidates: &[&str]) -> Option<&'a serde_json::Value> {
    let map = obj.as_object()?;
    for candidate in candidates {
        let lower = candidate.to_lowercase();
        for (key, value) in map {
            if key.to_lowercase() == lower {
                return Some(value);
            }
        }
    }
    None
}

/// Extract a string value via case-insensitive key lookup.
fn get_ci_str<'a>(obj: &'a serde_json::Value, candidates: &[&str]) -> Option<&'a str> {
    get_ci(obj, candidates)?.as_str()
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

    /// Fetch funding data for a single date.
    ///
    /// Returns `(raw_body_text, date_format_used, parsed_deposits)`.
    /// `raw_body_text` is always returned (even for empty/error responses) for artifact saving.
    fn fetch_funding(
        &self,
        date: &NaiveDate,
        quiet: bool,
    ) -> Result<(String, &'static str, Vec<FundingDeposit>), CliError> {
        let show_progress = !quiet && atty::is(atty::Stream::Stderr);

        // Try YYYYMMDD first, fallback to MMDD
        let yyyymmdd = date_to_yyyymmdd(date);
        let mmdd = date_to_mmdd(date);

        let (body, date_format) = match self.try_funding_request(&yyyymmdd) {
            Ok(body) => (body, "YYYYMMDD"),
            Err(_) => {
                if show_progress {
                    eprintln!("  {}: YYYYMMDD format failed, trying MMDD", date);
                }
                let body = self.try_funding_request(&mmdd)?;
                (body, "MMDD")
            }
        };

        let trimmed = body.trim();

        // Handle empty/error responses — these are normal (no deposits for this date)
        if trimmed.is_empty()
            || trimmed.eq_ignore_ascii_case("no deposits match your request")
            || trimmed.eq_ignore_ascii_case("null")
        {
            if show_progress {
                eprintln!("  {}: no funding data (format: {})", date, date_format);
            }
            return Ok((body.clone(), date_format, Vec::new()));
        }

        // Try to parse as JSON
        let parsed: serde_json::Value = match serde_json::from_str(trimmed) {
            Ok(v) => v,
            Err(_) => {
                // Non-JSON response (e.g., text error message) — treat as empty
                if show_progress {
                    eprintln!("  {}: non-JSON funding response (format: {})", date, date_format);
                }
                return Ok((body.clone(), date_format, Vec::new()));
            }
        };

        // If parsed is a string message (e.g., error envelope), treat as empty
        if parsed.is_string() {
            if show_progress {
                eprintln!("  {}: no funding data (format: {})", date, date_format);
            }
            return Ok((body.clone(), date_format, Vec::new()));
        }

        let deposits = parse_funding_deposits(&parsed, &self.merchant_id, date)?;

        if show_progress {
            eprintln!("  {}: {} deposits (format: {})", date, deposits.len(), date_format);
        }

        Ok((body.clone(), date_format, deposits))
    }

    /// Attempt a single funding request with the given date parameter.
    fn try_funding_request(&self, date_param: &str) -> Result<String, CliError> {
        let url = format!(
            "{}/cardconnect/rest/funding?merchid={}&date={}",
            self.api_url, self.merchant_id, date_param,
        );

        let username = self.api_username.clone();
        let password = self.api_password.clone();
        let url_clone = url.clone();

        self.client.request_with_retry_text(|http| {
            http.get(&url_clone)
                .basic_auth(&username, Some(&password))
        })
    }
}

fn extract_error(body: &serde_json::Value, status: u16) -> String {
    body["resptext"]
        .as_str()
        .or_else(|| body["message"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Funding parsing ─────────────────────────────────────────────────

/// Parse funding deposits from a JSON response.
///
/// The response may be an array of deposit objects or a single object.
/// Uses case-insensitive key lookup for resilience against API changes.
///
/// For deposits without a `fundingid`, a stable fallback ID is constructed
/// as `{merchid}:{date}:{seq}`. To ensure `seq` is deterministic regardless
/// of JSON array ordering, items without a fundingid are sorted by
/// `(funding_date, amount_minor, description)` before seq assignment.
fn parse_funding_deposits(
    parsed: &serde_json::Value,
    merchant_id: &str,
    date: &NaiveDate,
) -> Result<Vec<FundingDeposit>, CliError> {
    let items: Vec<&serde_json::Value> = if let Some(arr) = parsed.as_array() {
        arr.iter().collect()
    } else if parsed.is_object() {
        // Check for nested array under common keys
        if let Some(arr) = parsed.get("fundings").and_then(|v| v.as_array())
            .or_else(|| parsed.get("txns").and_then(|v| v.as_array()))
            .or_else(|| parsed.get("deposits").and_then(|v| v.as_array()))
        {
            arr.iter().collect()
        } else {
            vec![parsed]
        }
    } else {
        return Ok(Vec::new());
    };

    // First pass: extract fields from each item without assigning fallback IDs
    struct ParsedItem {
        explicit_funding_id: Option<String>,
        funding_date: String,
        amount_minor: i64,
        merchant_id: String,
        description: String,
    }

    let mut parsed_items = Vec::new();

    for (idx, item) in items.iter().enumerate() {
        // Extract date — try multiple field names in preference order
        let funding_date = get_ci_str(item, &["fundingdate", "fundingDate", "depositdate", "depositDate", "achdate", "achDate"])
            .map(|s| normalize_funding_date(s, date));

        // Extract amount — try multiple field names
        let amount_str = get_ci_str(item, &["amount", "depositamount", "depositAmount", "netamount", "netAmount", "fundedamount", "fundedAmount"]);

        // Both date and amount must be extractable
        let funding_date = match funding_date {
            Some(d) => d,
            None => {
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_UPSTREAM,
                    message: format!(
                        "unknown funding schema — no recognizable date field in deposit #{} for {}",
                        idx, date,
                    ),
                    hint: Some("Raw JSON was saved if --save-raw was set. Inspect it to determine field names.".into()),
                });
            }
        };

        let amount_value = match amount_str {
            Some(s) => s.to_string(),
            None => {
                // Try numeric values
                if let Some(val) = get_ci(item, &["amount", "depositamount", "depositAmount", "netamount", "netAmount", "fundedamount", "fundedAmount"]) {
                    if let Some(n) = val.as_f64() {
                        format!("{:.2}", n)
                    } else {
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "unknown funding schema — no recognizable amount field in deposit #{} for {}",
                                idx, date,
                            ),
                            hint: Some("Raw JSON was saved if --save-raw was set. Inspect it to determine field names.".into()),
                        });
                    }
                } else {
                    return Err(CliError {
                        code: exit_codes::EXIT_FETCH_UPSTREAM,
                        message: format!(
                            "unknown funding schema — no recognizable amount field in deposit #{} for {}",
                            idx, date,
                        ),
                        hint: Some("Raw JSON was saved if --save-raw was set. Inspect it to determine field names.".into()),
                    });
                }
            }
        };

        let amount_minor = common::parse_money_string(&amount_value).map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!(
                "Fiserv funding bad amount {:?} for deposit #{}: {}",
                amount_value, idx, e,
            ),
            hint: None,
        })?;

        let explicit_funding_id = get_ci_str(item, &["fundingid", "fundingId", "fundingID"])
            .map(|s| s.to_string());

        let description = build_funding_description(item);

        let mid = get_ci_str(item, &["merchid", "merchantid", "merchantId"])
            .unwrap_or(merchant_id)
            .to_string();

        parsed_items.push(ParsedItem {
            explicit_funding_id,
            funding_date,
            amount_minor,
            merchant_id: mid,
            description,
        });
    }

    // Sort for deterministic seq assignment: (funding_date, amount_minor, description)
    // This ensures the same data produces the same fallback IDs regardless of JSON order.
    parsed_items.sort_by(|a, b| {
        a.funding_date.cmp(&b.funding_date)
            .then_with(|| a.amount_minor.cmp(&b.amount_minor))
            .then_with(|| a.description.cmp(&b.description))
    });

    // Second pass: assign fallback IDs using stable seq from sorted order
    let mut deposits = Vec::new();
    for (seq, item) in parsed_items.iter().enumerate() {
        let funding_id = item.explicit_funding_id.clone().unwrap_or_else(|| {
            format!("{}:{}:{}", item.merchant_id, item.funding_date, seq)
        });

        deposits.push(FundingDeposit {
            funding_id,
            funding_date: item.funding_date.clone(),
            amount_minor: item.amount_minor,
            merchant_id: item.merchant_id.clone(),
            currency: "USD".to_string(),
            description: item.description.clone(),
        });
    }

    Ok(deposits)
}

/// Normalize a funding date to YYYY-MM-DD format.
/// Handles YYYYMMDD, MMDDYYYY, YYYY-MM-DD, and MM/DD/YYYY.
fn normalize_funding_date(s: &str, fallback: &NaiveDate) -> String {
    let s = s.trim();
    // Already ISO format
    if s.len() == 10 && s.chars().nth(4) == Some('-') {
        return s.to_string();
    }
    // YYYYMMDD
    if s.len() == 8 && s.chars().all(|c| c.is_ascii_digit()) {
        if let Ok(d) = NaiveDate::parse_from_str(s, "%Y%m%d") {
            return d.format("%Y-%m-%d").to_string();
        }
        // Try MMDDYYYY
        return parse_authdate(s);
    }
    // MM/DD/YYYY
    if s.len() == 10 && s.chars().nth(2) == Some('/') {
        if let Ok(d) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
            return d.format("%Y-%m-%d").to_string();
        }
    }
    // Fallback to query date
    fallback.format("%Y-%m-%d").to_string()
}

/// Build a human-readable description from available funding fields.
fn build_funding_description(item: &serde_json::Value) -> String {
    let mut parts = Vec::new();

    if let Some(s) = get_ci_str(item, &["fundingmasterid", "fundingMasterId"]) {
        parts.push(format!("master:{}", s));
    }
    if let Some(s) = get_ci_str(item, &["bankname", "bankName"]) {
        parts.push(s.to_string());
    }
    if let Some(s) = get_ci_str(item, &["currency"]) {
        if s != "USD" {
            parts.push(s.to_string());
        }
    }

    parts.join(" ").trim().to_string()
}

// ── Raw JSON saving ─────────────────────────────────────────────────

/// Write raw JSON response to a file in the save_raw directory.
fn save_raw_json(save_raw: &PathBuf, date: &NaiveDate, body: &str) -> Result<(), CliError> {
    let filename = format!("fiserv_funding_raw_{}.json", date.format("%Y-%m-%d"));
    let path = save_raw.join(&filename);
    let mut f = std::fs::File::create(&path).map_err(|e| {
        CliError::io(format!("cannot create {}: {}", path.display(), e))
    })?;
    f.write_all(body.as_bytes()).map_err(|e| {
        CliError::io(format!("cannot write {}: {}", path.display(), e))
    })?;
    Ok(())
}

/// Write engine_meta.json summarizing the funding run.
fn save_engine_meta(
    save_raw: &PathBuf,
    date_format_used: &str,
    dates_fetched: u32,
    dates_with_data: u32,
    total_deposits: u32,
) -> Result<(), CliError> {
    let meta = serde_json::json!({
        "fiserv_mode": "funding",
        "date_format_used": date_format_used,
        "dates_fetched": dates_fetched,
        "dates_with_data": dates_with_data,
        "total_deposits": total_deposits,
        "vgrid_version": env!("CARGO_PKG_VERSION"),
    });
    let path = save_raw.join("engine_meta.json");
    let mut f = std::fs::File::create(&path).map_err(|e| {
        CliError::io(format!("cannot create {}: {}", path.display(), e))
    })?;
    f.write_all(serde_json::to_string_pretty(&meta).unwrap().as_bytes())
        .map_err(|e| CliError::io(format!("cannot write {}: {}", path.display(), e)))?;
    Ok(())
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
    funding: bool,
    save_raw: Option<PathBuf>,
    quiet: bool,
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

    // 3. Create save_raw directory if needed
    if let Some(ref dir) = save_raw {
        std::fs::create_dir_all(dir).map_err(|e| {
            CliError::io(format!("cannot create --save-raw dir {}: {}", dir.display(), e))
        })?;
    }

    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    let client = FiservClient::new(url, mid, username, password);

    if funding {
        cmd_fetch_fiserv_funding(&client, from_date, to_date, out, save_raw, quiet, show_progress)
    } else {
        cmd_fetch_fiserv_settlestat(&client, from_date, to_date, out, quiet, show_progress)
    }
}

fn cmd_fetch_fiserv_funding(
    client: &FiservClient,
    from_date: NaiveDate,
    to_date: NaiveDate,
    out: Option<PathBuf>,
    save_raw: Option<PathBuf>,
    quiet: bool,
    show_progress: bool,
) -> Result<(), CliError> {
    if show_progress {
        eprintln!(
            "Fetching Fiserv/CardPointe funding data ({} to {})...",
            from_date, to_date,
        );
    }

    let mut all_deposits: Vec<FundingDeposit> = Vec::new();
    let mut dates_fetched: u32 = 0;
    let mut dates_with_data: u32 = 0;
    let mut last_date_format = "YYYYMMDD";

    let mut current = from_date;
    while current < to_date {
        let (raw_body, date_format, deposits) = client.fetch_funding(&current, quiet)?;
        dates_fetched += 1;
        last_date_format = date_format;

        if !deposits.is_empty() {
            dates_with_data += 1;
        }

        // Always save raw JSON (even empty responses)
        if let Some(ref dir) = save_raw {
            save_raw_json(dir, &current, &raw_body)?;
        }

        all_deposits.extend(deposits);
        current = current.succ_opt().unwrap_or(current);
    }

    // Save engine_meta
    if let Some(ref dir) = save_raw {
        save_engine_meta(
            dir,
            last_date_format,
            dates_fetched,
            dates_with_data,
            all_deposits.len() as u32,
        )?;
    }

    // Sort: (funding_date, funding_id)
    all_deposits.sort_by(|a, b| {
        a.funding_date
            .cmp(&b.funding_date)
            .then_with(|| a.funding_id.cmp(&b.funding_id))
    });

    // Build canonical rows
    let rows: Vec<CanonicalRow> = all_deposits
        .iter()
        .map(|d| CanonicalRow {
            effective_date: d.funding_date.clone(),
            posted_date: d.funding_date.clone(),
            amount_minor: d.amount_minor,
            currency: d.currency.clone(),
            r#type: "funding".to_string(),
            source: "fiserv".to_string(),
            source_id: d.funding_id.clone(),
            group_id: String::new(),
            description: d.description.clone(),
        })
        .collect();

    // Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!("Done: {} funding deposits written to {}", rows.len(), out_label);
    }

    Ok(())
}

fn cmd_fetch_fiserv_settlestat(
    client: &FiservClient,
    from_date: NaiveDate,
    to_date: NaiveDate,
    out: Option<PathBuf>,
    quiet: bool,
    show_progress: bool,
) -> Result<(), CliError> {
    if show_progress {
        eprintln!(
            "Fetching Fiserv/CardPointe settled transactions ({} to {})...",
            from_date, to_date,
        );
    }

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

    // 4. Sort: (effective_date, retref)
    all_txns.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.retref.cmp(&b.retref))
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
            source: "fiserv".to_string(),
            source_id: txn.retref.clone(),
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

    // ── Funding tests ───────────────────────────────────────────────

    #[test]
    fn test_fetch_funding() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/cardconnect/rest/funding")
                .query_param("merchid", "merchant123")
                .query_param("date", "20260215");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!([
                    {
                        "fundingid": "F001",
                        "fundingdate": "20260215",
                        "amount": "1500.00",
                        "merchid": "merchant123"
                    },
                    {
                        "fundingid": "F002",
                        "fundingdate": "20260215",
                        "amount": "250.75",
                        "merchid": "merchant123"
                    }
                ]));
        });

        let client = FiservClient::new(
            server.base_url(),
            "merchant123".into(),
            "user".into(),
            "pass".into(),
        );

        let date = NaiveDate::from_ymd_opt(2026, 2, 15).unwrap();
        let (raw_body, date_format, deposits) = client.fetch_funding(&date, true).unwrap();

        assert_eq!(date_format, "YYYYMMDD");
        assert!(!raw_body.is_empty());
        assert_eq!(deposits.len(), 2);
        // Sorted by (date, amount): F002 (25075) before F001 (150000)
        assert_eq!(deposits[0].funding_id, "F002");
        assert_eq!(deposits[0].funding_date, "2026-02-15");
        assert_eq!(deposits[0].amount_minor, 25075);
        assert_eq!(deposits[1].funding_id, "F001");
        assert_eq!(deposits[1].amount_minor, 150000);
    }

    #[test]
    fn test_fetch_funding_empty() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/cardconnect/rest/funding")
                .query_param("merchid", "merchant123")
                .query_param("date", "20260215");
            then.status(200)
                .header("content-type", "text/plain")
                .body("No deposits match your request");
        });

        let client = FiservClient::new(
            server.base_url(),
            "merchant123".into(),
            "user".into(),
            "pass".into(),
        );

        let date = NaiveDate::from_ymd_opt(2026, 2, 15).unwrap();
        let (raw_body, _, deposits) = client.fetch_funding(&date, true).unwrap();

        assert!(deposits.is_empty());
        assert!(raw_body.contains("No deposits match your request"));
    }

    #[test]
    fn test_funding_source_id_fallback() {
        // No fundingid field — should construct stable key.
        // Items are sorted by (date, amount, description) before seq assignment,
        // so the 300.00 item gets seq=0 and 500.00 gets seq=1.
        let date = NaiveDate::from_ymd_opt(2026, 2, 15).unwrap();
        let json = serde_json::json!([
            {
                "fundingdate": "20260215",
                "amount": "500.00",
                "merchid": "M001"
            },
            {
                "fundingdate": "20260215",
                "amount": "300.00",
                "merchid": "M001"
            }
        ]);

        let deposits = parse_funding_deposits(&json, "M001", &date).unwrap();
        assert_eq!(deposits.len(), 2);
        // Sorted by amount: 300.00 (30000 cents) first, then 500.00 (50000 cents)
        assert_eq!(deposits[0].funding_id, "M001:2026-02-15:0");
        assert_eq!(deposits[0].amount_minor, 30000);
        assert_eq!(deposits[1].funding_id, "M001:2026-02-15:1");
        assert_eq!(deposits[1].amount_minor, 50000);

        // Verify stability: reversed JSON order produces identical IDs
        let json_reversed = serde_json::json!([
            {
                "fundingdate": "20260215",
                "amount": "300.00",
                "merchid": "M001"
            },
            {
                "fundingdate": "20260215",
                "amount": "500.00",
                "merchid": "M001"
            }
        ]);
        let deposits2 = parse_funding_deposits(&json_reversed, "M001", &date).unwrap();
        assert_eq!(deposits[0].funding_id, deposits2[0].funding_id);
        assert_eq!(deposits[1].funding_id, deposits2[1].funding_id);
    }

    #[test]
    fn test_funding_canonical_output() {
        let deposit = FundingDeposit {
            funding_id: "F001".to_string(),
            funding_date: "2026-02-15".to_string(),
            amount_minor: 150000,
            merchant_id: "M001".to_string(),
            currency: "USD".to_string(),
            description: "".to_string(),
        };

        let row = CanonicalRow {
            effective_date: deposit.funding_date.clone(),
            posted_date: deposit.funding_date.clone(),
            amount_minor: deposit.amount_minor,
            currency: deposit.currency.clone(),
            r#type: "funding".to_string(),
            source: "fiserv".to_string(),
            source_id: deposit.funding_id.clone(),
            group_id: String::new(),
            description: deposit.description.clone(),
        };

        assert_eq!(row.r#type, "funding");
        assert_eq!(row.source, "fiserv");
        assert_eq!(row.effective_date, "2026-02-15");
        assert_eq!(row.posted_date, "2026-02-15");
        assert_eq!(row.amount_minor, 150000);
        assert_eq!(row.source_id, "F001");
        assert!(row.group_id.is_empty());
    }

    #[test]
    fn test_funding_date_format_fallback() {
        let server = MockServer::start();

        // YYYYMMDD returns 400
        server.mock(|when, then| {
            when.method(GET)
                .path("/cardconnect/rest/funding")
                .query_param("merchid", "merchant123")
                .query_param("date", "20260215");
            then.status(400)
                .header("content-type", "application/json")
                .json_body(serde_json::json!({"message": "Invalid date format"}));
        });

        // MMDD succeeds
        server.mock(|when, then| {
            when.method(GET)
                .path("/cardconnect/rest/funding")
                .query_param("merchid", "merchant123")
                .query_param("date", "0215");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!([
                    {
                        "fundingid": "F003",
                        "fundingdate": "20260215",
                        "amount": "750.00",
                        "merchid": "merchant123"
                    }
                ]));
        });

        let client = FiservClient::new(
            server.base_url(),
            "merchant123".into(),
            "user".into(),
            "pass".into(),
        );

        let date = NaiveDate::from_ymd_opt(2026, 2, 15).unwrap();
        let (_, date_format, deposits) = client.fetch_funding(&date, true).unwrap();

        assert_eq!(date_format, "MMDD");
        assert_eq!(deposits.len(), 1);
        assert_eq!(deposits[0].funding_id, "F003");
        assert_eq!(deposits[0].amount_minor, 75000);
    }

    #[test]
    fn test_normalize_funding_date() {
        let fallback = NaiveDate::from_ymd_opt(2026, 2, 15).unwrap();

        // YYYYMMDD
        assert_eq!(normalize_funding_date("20260215", &fallback), "2026-02-15");
        // Already ISO
        assert_eq!(normalize_funding_date("2026-02-15", &fallback), "2026-02-15");
        // MMDDYYYY
        assert_eq!(normalize_funding_date("02152026", &fallback), "2026-02-15");
        // MM/DD/YYYY
        assert_eq!(normalize_funding_date("02/15/2026", &fallback), "2026-02-15");
    }

    #[test]
    fn test_get_ci_str() {
        let json = serde_json::json!({
            "FundingId": "F001",
            "Amount": "500.00"
        });

        assert_eq!(get_ci_str(&json, &["fundingid", "fundingId"]), Some("F001"));
        assert_eq!(get_ci_str(&json, &["amount"]), Some("500.00"));
        assert_eq!(get_ci_str(&json, &["nonexistent"]), None);
    }
}
