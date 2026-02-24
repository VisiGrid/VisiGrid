//! `vgrid fetch digits` — fetch ledger entries from Digits into canonical CSV.

use std::path::PathBuf;

use chrono::{NaiveDate, Utc};
use serde::{Deserialize, Serialize};

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const DIGITS_API_BASE: &str = "https://api.digits.com";
const DIGITS_TOKEN_URL: &str = "https://auth.digits.com/oauth2/token";

// ── Credentials ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct DigitsCredentials {
    pub client_id: String,
    pub client_secret: String,
    pub access_token: String,
    pub refresh_token: String,
    pub legal_entity_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub access_token_expires_at: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub refresh_token_expires_at: Option<String>,
}

fn load_credentials(path: &PathBuf) -> Result<DigitsCredentials, CliError> {
    let content = std::fs::read_to_string(path).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_NOT_AUTH,
        message: format!(
            "cannot read credentials file {}: {}",
            path.display(),
            e,
        ),
        hint: None,
    })?;

    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        if let Ok(meta) = std::fs::metadata(path) {
            let mode = meta.permissions().mode();
            if mode & 0o077 != 0 {
                eprintln!(
                    "warning: credentials file {} is accessible by others (mode {:o}), consider chmod 600",
                    path.display(),
                    mode & 0o777,
                );
            }
        }
    }

    serde_json::from_str(&content).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_NOT_AUTH,
        message: format!(
            "invalid credentials JSON in {}: {}",
            path.display(),
            e,
        ),
        hint: None,
    })
}

fn save_credentials(creds: &DigitsCredentials, path: &PathBuf) -> Result<(), CliError> {
    let json = serde_json::to_string_pretty(creds).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("failed to serialize credentials: {}", e),
        hint: None,
    })?;
    std::fs::write(path, json).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!(
            "failed to write credentials to {}: {}",
            path.display(),
            e,
        ),
        hint: None,
    })?;
    Ok(())
}

fn refresh_access_token(
    creds: &DigitsCredentials,
    http: &reqwest::blocking::Client,
    token_url: &str,
) -> Result<DigitsCredentials, CliError> {
    use base64::Engine;

    let basic = base64::engine::general_purpose::STANDARD
        .encode(format!("{}:{}", creds.client_id, creds.client_secret));

    let resp = http
        .post(token_url)
        .header("Authorization", format!("Basic {}", basic))
        .header("Content-Type", "application/x-www-form-urlencoded")
        .body(format!(
            "grant_type=refresh_token&refresh_token={}",
            creds.refresh_token,
        ))
        .send()
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: format!("Digits token refresh request failed: {}", e),
            hint: None,
        })?;

    let status = resp.status().as_u16();
    if status != 200 {
        let body: serde_json::Value = resp.json().unwrap_or(serde_json::Value::Null);
        let msg = body["error_description"]
            .as_str()
            .or_else(|| body["error"].as_str())
            .unwrap_or("unknown error");
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: format!(
                "Digits token refresh failed ({}): {}",
                status, msg,
            ),
            hint: Some(
                "Refresh token expired or revoked. Reconnect Digits \
                 in your OAuth app (or your account settings) to regenerate credentials."
                    .into(),
            ),
        });
    }

    let body: serde_json::Value = resp.json().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_AUTH,
        message: format!("Digits token refresh response invalid: {}", e),
        hint: None,
    })?;

    let new_access = body["access_token"]
        .as_str()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: "Digits token refresh response missing access_token".into(),
            hint: None,
        })?;

    let new_refresh = body["refresh_token"]
        .as_str()
        .unwrap_or(&creds.refresh_token);

    let access_token_expires_at = body["expires_in"]
        .as_i64()
        .map(|secs| {
            (Utc::now() + chrono::Duration::seconds(secs))
                .to_rfc3339()
        });

    let refresh_token_expires_at = body["x_refresh_token_expires_in"]
        .as_i64()
        .map(|secs| {
            let expires = Utc::now() + chrono::Duration::seconds(secs);
            let days_left = (expires - Utc::now()).num_days();
            if days_left <= 30 {
                eprintln!(
                    "warning: Digits refresh token expires {}, re-authorize soon",
                    expires.format("%Y-%m-%d"),
                );
            }
            expires.to_rfc3339()
        })
        .or_else(|| creds.refresh_token_expires_at.clone());

    Ok(DigitsCredentials {
        client_id: creds.client_id.clone(),
        client_secret: creds.client_secret.clone(),
        access_token: new_access.to_string(),
        refresh_token: new_refresh.to_string(),
        legal_entity_id: creds.legal_entity_id.clone(),
        access_token_expires_at,
        refresh_token_expires_at,
    })
}

// ── Error extraction ────────────────────────────────────────────────

fn extract_digits_error(body: &serde_json::Value, status: u16) -> String {
    body["error"]["message"]
        .as_str()
        .or_else(|| body["message"].as_str())
        .or_else(|| body["error"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Internal entry row ──────────────────────────────────────────────

#[derive(Debug)]
struct RawDigitsEntry {
    effective_date: String,
    posted_date: String,
    amount_minor: i64,
    currency: String,
    entry_type: String,
    source_id: String,
    group_id: String,
    description: String,
}

// ── Entry parsing ───────────────────────────────────────────────────

fn parse_entry(entry: &serde_json::Value) -> Result<RawDigitsEntry, CliError> {
    let id = entry["id"].as_str().unwrap_or("").to_string();
    let effective_date = entry["effectiveDate"].as_str().unwrap_or("").to_string();
    let posted_date = entry["postedDate"].as_str().unwrap_or("").to_string();

    let amount_raw = entry["amount"]["amount"]
        .as_i64()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("entry {} missing amount.amount", id),
            hint: None,
        })?;

    let currency = entry["amount"]["currency"]
        .as_str()
        .unwrap_or("USD")
        .to_string();

    let entry_type_raw = entry["type"].as_str().unwrap_or("").to_string();
    let entry_type = entry_type_raw.to_lowercase();

    // Sign convention: Credit = positive, Debit = negative
    let amount_minor = if entry_type == "debit" {
        -amount_raw.abs()
    } else {
        amount_raw.abs()
    };

    let transaction_id = entry["transactionId"]
        .as_str()
        .unwrap_or("")
        .to_string();

    let description = entry["description"]
        .as_str()
        .unwrap_or("")
        .to_string();

    Ok(RawDigitsEntry {
        effective_date,
        posted_date,
        amount_minor,
        currency,
        entry_type,
        source_id: format!("entry:{}", id),
        group_id: transaction_id,
        description,
    })
}

// ── Digits client ───────────────────────────────────────────────────

pub struct DigitsClient {
    client: FetchClient,
    access_token: String,
    legal_entity_id: String,
    base_url: String,
    token_url: String,
    creds: Option<DigitsCredentials>,
    creds_path: Option<PathBuf>,
}

impl DigitsClient {
    pub fn new(access_token: String, legal_entity_id: String) -> Self {
        Self::with_base_url(
            access_token,
            legal_entity_id,
            DIGITS_API_BASE.to_string(),
            DIGITS_TOKEN_URL.to_string(),
        )
    }

    fn with_base_url(
        access_token: String,
        legal_entity_id: String,
        base_url: String,
        token_url: String,
    ) -> Self {
        Self {
            client: FetchClient::new("Digits", extract_digits_error),
            access_token,
            legal_entity_id,
            base_url,
            token_url,
            creds: None,
            creds_path: None,
        }
    }

    fn from_credentials(creds: DigitsCredentials, path: PathBuf) -> Self {
        Self {
            client: FetchClient::new("Digits", extract_digits_error),
            access_token: creds.access_token.clone(),
            legal_entity_id: creds.legal_entity_id.clone(),
            base_url: DIGITS_API_BASE.to_string(),
            token_url: DIGITS_TOKEN_URL.to_string(),
            creds: Some(creds),
            creds_path: Some(path),
        }
    }

    #[cfg(test)]
    fn from_credentials_with_base_url(
        creds: DigitsCredentials,
        path: PathBuf,
        base_url: String,
    ) -> Self {
        let token_url = format!("{}/oauth2/token", base_url);
        Self {
            client: FetchClient::new("Digits", extract_digits_error),
            access_token: creds.access_token.clone(),
            legal_entity_id: creds.legal_entity_id.clone(),
            base_url,
            token_url,
            creds: Some(creds),
            creds_path: Some(path),
        }
    }

    fn try_refresh(&mut self) -> Result<(), CliError> {
        let creds = self.creds.as_ref().ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: "cannot refresh token without credentials file".into(),
            hint: None,
        })?;
        let path = self.creds_path.as_ref().unwrap();

        let new_creds = refresh_access_token(creds, &self.client.http, &self.token_url)?;
        save_credentials(&new_creds, path)?;
        self.access_token = new_creds.access_token.clone();
        self.creds = Some(new_creds);
        Ok(())
    }

    fn fetch_entries(
        &mut self,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawDigitsEntry>, CliError> {
        let mut all = Vec::new();
        let mut cursor: Option<String> = None;
        let mut refreshed = false;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;
        let mut page = 0u32;

        loop {
            let url = format!("{}/v1/ledger/entries", self.base_url);
            let token = self.access_token.clone();
            let from_str = from.to_string();
            let to_str = to.to_string();
            let lei = self.legal_entity_id.clone();
            let cursor_clone = cursor.clone();

            let result = self.client.request_with_retry(|http| {
                let mut req = http
                    .get(&url)
                    .bearer_auth(&token)
                    .header("Accept", "application/json")
                    .query(&[
                        ("startDate", from_str.as_str()),
                        ("endDate", to_str.as_str()),
                        ("legalEntityId", lei.as_str()),
                    ]);
                if let Some(ref c) = cursor_clone {
                    req = req.query(&[("cursor", c.as_str())]);
                }
                req
            });

            let body = match result {
                Ok(body) => body,
                Err(e)
                    if e.code == exit_codes::EXIT_FETCH_AUTH
                        && !refreshed
                        && self.creds.is_some() =>
                {
                    self.try_refresh()?;
                    refreshed = true;
                    let token = self.access_token.clone();
                    let cursor_clone = cursor.clone();
                    self.client.request_with_retry(|http| {
                        let mut req = http
                            .get(&url)
                            .bearer_auth(&token)
                            .header("Accept", "application/json")
                            .query(&[
                                ("startDate", from_str.as_str()),
                                ("endDate", to_str.as_str()),
                                ("legalEntityId", lei.as_str()),
                            ]);
                        if let Some(ref c) = cursor_clone {
                            req = req.query(&[("cursor", c.as_str())]);
                        }
                        req
                    })?
                }
                Err(e) => return Err(e),
            };

            let entries = body["entries"]
                .as_array()
                .cloned()
                .unwrap_or_default();

            let count = entries.len();

            if show_progress {
                page += 1;
                eprintln!("  page {}: {} entries", page, count);
            }

            for entry in &entries {
                all.push(parse_entry(entry)?);
            }

            let more = body["next"]["more"].as_bool().unwrap_or(false);
            if !more {
                break;
            }
            cursor = body["next"]["cursor"].as_str().map(String::from);
            if cursor.is_none() {
                break;
            }
        }

        Ok(all)
    }
}

// ── Include filter ──────────────────────────────────────────────────

fn parse_include(include: &str) -> Result<Vec<String>, CliError> {
    let mut types = Vec::new();
    for part in include.split(',') {
        let t = part.trim().to_lowercase();
        match t.as_str() {
            "credit" | "debit" => types.push(t),
            "" => {}
            other => {
                return Err(CliError {
                    code: exit_codes::EXIT_USAGE,
                    message: format!(
                        "Unknown entry type '{}'. Valid: credit, debit",
                        other,
                    ),
                    hint: None,
                });
            }
        }
    }
    if types.is_empty() {
        types = vec!["credit".to_string(), "debit".to_string()];
    }
    Ok(types)
}

// ── Entry point ─────────────────────────────────────────────────────

#[allow(clippy::too_many_arguments)]
pub fn cmd_fetch_digits(
    from: String,
    to: String,
    credentials: Option<PathBuf>,
    access_token: Option<String>,
    legal_entity_id: Option<String>,
    _account: Option<String>,
    _account_id: Option<String>,
    include: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 2. Resolve auth
    let mut client = if let Some(token) = access_token {
        let token = token.trim().to_string();
        if token.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing Digits access token (--access-token is empty)".into(),
                hint: None,
            });
        }
        let lei = legal_entity_id.ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "missing --legal-entity-id (required with --access-token)".into(),
            hint: None,
        })?;
        let lei = lei.trim().to_string();
        if lei.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing Digits legal entity ID (--legal-entity-id is empty)".into(),
                hint: None,
            });
        }
        DigitsClient::new(token, lei)
    } else if let Some(ref path) = credentials {
        let expanded =
            shellexpand::tilde(&path.to_string_lossy()).to_string();
        let creds_path = PathBuf::from(&expanded);
        let creds = load_credentials(&creds_path)?;
        DigitsClient::from_credentials(creds, creds_path)
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "Use --credentials or --access-token + --legal-entity-id".into(),
            hint: None,
        });
    };

    // 3. Parse include filter
    let include_types = parse_include(
        include.as_deref().unwrap_or("credit,debit"),
    )?;

    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Digits ledger entries ({} to {})...",
            from_date, to_date,
        );
    }

    // 4. Fetch all entries
    let all_entries = client.fetch_entries(&from_date, &to_date, quiet)?;

    // 5. Filter by account name/ID if specified
    let filtered: Vec<&RawDigitsEntry> = all_entries
        .iter()
        .filter(|e| {
            // Filter by include types
            if !include_types.contains(&e.entry_type) {
                return false;
            }
            true
        })
        .collect();

    // 6. Sort: (effective_date ASC, source_id ASC)
    let mut sorted: Vec<&RawDigitsEntry> = filtered;
    sorted.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 7. Build canonical rows
    let rows: Vec<CanonicalRow> = sorted
        .iter()
        .map(|r| CanonicalRow {
            effective_date: r.effective_date.clone(),
            posted_date: r.posted_date.clone(),
            amount_minor: r.amount_minor,
            currency: r.currency.clone(),
            r#type: r.entry_type.clone(),
            source: "digits".to_string(),
            source_id: r.source_id.clone(),
            group_id: r.group_id.clone(),
            description: r.description.clone(),
        })
        .collect();

    // 8. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        let credit_count = sorted.iter().filter(|e| e.entry_type == "credit").count();
        let debit_count = sorted.iter().filter(|e| e.entry_type == "debit").count();
        eprintln!(
            "Done: {} credits + {} debits = {} rows written to {}",
            credit_count, debit_count, rows.len(), out_label,
        );
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── parse_entry ─────────────────────────────────────────────────

    #[test]
    fn test_parse_entry_credit() {
        let entry = serde_json::json!({
            "id": "ent_001",
            "effectiveDate": "2026-01-15",
            "postedDate": "2026-01-16",
            "amount": { "amount": 150000, "currency": "USD" },
            "type": "Credit",
            "transactionId": "txn_abc",
            "description": "Wire payment received",
            "account": { "id": "acc_1", "name": "Operating Account" }
        });

        let raw = parse_entry(&entry).unwrap();
        assert_eq!(raw.effective_date, "2026-01-15");
        assert_eq!(raw.posted_date, "2026-01-16");
        assert_eq!(raw.amount_minor, 150000); // positive for credit
        assert_eq!(raw.currency, "USD");
        assert_eq!(raw.entry_type, "credit");
        assert_eq!(raw.source_id, "entry:ent_001");
        assert_eq!(raw.group_id, "txn_abc");
        assert_eq!(raw.description, "Wire payment received");
    }

    #[test]
    fn test_parse_entry_debit() {
        let entry = serde_json::json!({
            "id": "ent_002",
            "effectiveDate": "2026-01-20",
            "postedDate": "2026-01-20",
            "amount": { "amount": 5000, "currency": "USD" },
            "type": "Debit",
            "transactionId": "txn_def",
            "description": "Office supplies",
            "account": { "id": "acc_2", "name": "Expenses" }
        });

        let raw = parse_entry(&entry).unwrap();
        assert_eq!(raw.amount_minor, -5000); // negative for debit
        assert_eq!(raw.entry_type, "debit");
        assert_eq!(raw.source_id, "entry:ent_002");
    }

    #[test]
    fn test_parse_entry_missing_optional_fields() {
        let entry = serde_json::json!({
            "id": "ent_003",
            "effectiveDate": "2026-02-01",
            "amount": { "amount": 100, "currency": "EUR" },
            "type": "Credit",
        });

        let raw = parse_entry(&entry).unwrap();
        assert_eq!(raw.effective_date, "2026-02-01");
        assert_eq!(raw.posted_date, "");
        assert_eq!(raw.amount_minor, 100);
        assert_eq!(raw.currency, "EUR");
        assert_eq!(raw.group_id, "");
        assert_eq!(raw.description, "");
    }

    // ── include filter ──────────────────────────────────────────────

    #[test]
    fn test_include_filter() {
        let types = parse_include("credit").unwrap();
        assert_eq!(types, vec!["credit"]);

        let types = parse_include("credit,debit").unwrap();
        assert_eq!(types, vec!["credit", "debit"]);

        let types = parse_include("debit").unwrap();
        assert_eq!(types, vec!["debit"]);

        let err = parse_include("credit,refund").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_USAGE);
        assert!(err.message.contains("refund"));
    }

    // ── Credential loading ──────────────────────────────────────────

    #[test]
    fn test_credential_loading_valid() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("digits.json");
        std::fs::write(
            &path,
            serde_json::json!({
                "client_id": "cid",
                "client_secret": "csec",
                "access_token": "at",
                "refresh_token": "rt",
                "legal_entity_id": "le_123"
            })
            .to_string(),
        )
        .unwrap();
        let creds = load_credentials(&path.to_path_buf()).unwrap();
        assert_eq!(creds.client_id, "cid");
        assert_eq!(creds.legal_entity_id, "le_123");
        assert!(creds.access_token_expires_at.is_none());
    }

    #[test]
    fn test_credential_loading_missing_file() {
        let path = PathBuf::from("/tmp/nonexistent-digits-creds.json");
        let err = load_credentials(&path).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("cannot read credentials file"));
    }

    #[test]
    fn test_credential_loading_invalid_json() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bad.json");
        std::fs::write(&path, "not json").unwrap();
        let err = load_credentials(&path.to_path_buf()).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("invalid credentials JSON"));
    }

    // ── Pagination (httpmock) ───────────────────────────────────────

    #[test]
    fn test_single_page_fetch() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries");
            then.status(200)
                .json_body(serde_json::json!({
                    "entries": [
                        {
                            "id": "ent_1",
                            "effectiveDate": "2026-01-15",
                            "postedDate": "2026-01-15",
                            "amount": { "amount": 10000, "currency": "USD" },
                            "type": "Credit",
                            "transactionId": "txn_1",
                            "description": "Payment"
                        }
                    ],
                    "next": { "more": false }
                }));
        });

        let mut client = DigitsClient::with_base_url(
            "test_token".into(),
            "le_123".into(),
            server.base_url(),
            format!("{}/oauth2/token", server.base_url()),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let entries = client.fetch_entries(&from, &to, true).unwrap();
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].amount_minor, 10000);
        assert_eq!(entries[0].source_id, "entry:ent_1");
    }

    #[test]
    fn test_cursor_pagination() {
        let server = MockServer::start();

        // Page 2: last page (register first so it takes priority when cursor param is present)
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries")
                .query_param("cursor", "page2");
            then.status(200)
                .json_body(serde_json::json!({
                    "entries": [{
                        "id": "ent_2",
                        "effectiveDate": "2026-01-20",
                        "postedDate": "2026-01-20",
                        "amount": { "amount": 3000, "currency": "USD" },
                        "type": "Debit",
                        "transactionId": "txn_2",
                        "description": "Second"
                    }],
                    "next": { "more": false }
                }));
        });

        // Page 1: no cursor param → first page
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries")
                .query_param_exists("startDate");
            then.status(200)
                .json_body(serde_json::json!({
                    "entries": [{
                        "id": "ent_1",
                        "effectiveDate": "2026-01-10",
                        "postedDate": "2026-01-10",
                        "amount": { "amount": 5000, "currency": "USD" },
                        "type": "Credit",
                        "transactionId": "txn_1",
                        "description": "First"
                    }],
                    "next": { "cursor": "page2", "more": true }
                }));
        });

        let mut client = DigitsClient::with_base_url(
            "test_token".into(),
            "le_123".into(),
            server.base_url(),
            format!("{}/oauth2/token", server.base_url()),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let entries = client.fetch_entries(&from, &to, true).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].source_id, "entry:ent_1");
        assert_eq!(entries[0].amount_minor, 5000);
        assert_eq!(entries[1].source_id, "entry:ent_2");
        assert_eq!(entries[1].amount_minor, -3000);
    }

    #[test]
    fn test_empty_result_set() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries");
            then.status(200)
                .json_body(serde_json::json!({
                    "entries": [],
                    "next": { "more": false }
                }));
        });

        let mut client = DigitsClient::with_base_url(
            "test_token".into(),
            "le_123".into(),
            server.base_url(),
            format!("{}/oauth2/token", server.base_url()),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let entries = client.fetch_entries(&from, &to, true).unwrap();
        assert_eq!(entries.len(), 0);
    }

    // ── Auth failure ────────────────────────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries");
            then.status(401)
                .json_body(serde_json::json!({
                    "error": { "message": "Invalid access token" }
                }));
        });

        let mut client = DigitsClient::with_base_url(
            "bad_token".into(),
            "le_123".into(),
            server.base_url(),
            format!("{}/oauth2/token", server.base_url()),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client.fetch_entries(&from, &to, true).unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("Digits auth failed (401)"),
            "message: {}",
            err.message,
        );
    }

    // ── Token refresh on 401 ────────────────────────────────────────

    #[test]
    fn test_token_refresh_on_401() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries")
                .header("Authorization", "Bearer old_token");
            then.status(401)
                .json_body(serde_json::json!({
                    "error": { "message": "Token expired" }
                }));
        });

        // Refresh → new token
        server.mock(|when, then| {
            when.method(POST)
                .path("/oauth2/token");
            then.status(200)
                .json_body(serde_json::json!({
                    "access_token": "new_token",
                    "refresh_token": "new_refresh",
                    "expires_in": 3600
                }));
        });

        // Retry with new token → success
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries")
                .header("Authorization", "Bearer new_token");
            then.status(200)
                .json_body(serde_json::json!({
                    "entries": [{
                        "id": "ent_1",
                        "effectiveDate": "2026-01-15",
                        "postedDate": "2026-01-15",
                        "amount": { "amount": 50000, "currency": "USD" },
                        "type": "Credit",
                        "transactionId": "txn_1",
                        "description": "Payment"
                    }],
                    "next": { "more": false }
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("digits.json");
        let creds = DigitsCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "old_refresh".into(),
            legal_entity_id: "le_123".into(),
            access_token_expires_at: None,
            refresh_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = DigitsClient::from_credentials_with_base_url(
            creds,
            creds_path.clone(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let entries = client.fetch_entries(&from, &to, true).unwrap();

        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].amount_minor, 50000);

        // Verify credentials were updated on disk
        let saved: DigitsCredentials = serde_json::from_str(
            &std::fs::read_to_string(&creds_path).unwrap(),
        )
        .unwrap();
        assert_eq!(saved.access_token, "new_token");
        assert_eq!(saved.refresh_token, "new_refresh");
    }

    // ── Refresh failure hint ────────────────────────────────────────

    #[test]
    fn test_refresh_failure_includes_reconnect_hint() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path("/v1/ledger/entries");
            then.status(401)
                .json_body(serde_json::json!({
                    "error": { "message": "Unauthorized" }
                }));
        });

        // Refresh → revoked
        server.mock(|when, then| {
            when.method(POST)
                .path("/oauth2/token");
            then.status(400)
                .json_body(serde_json::json!({
                    "error": "invalid_grant",
                    "error_description": "refresh token has been revoked"
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("digits.json");
        let creds = DigitsCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "revoked".into(),
            legal_entity_id: "le_123".into(),
            access_token_expires_at: None,
            refresh_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = DigitsClient::from_credentials_with_base_url(
            creds,
            creds_path,
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client.fetch_entries(&from, &to, true).unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("token refresh failed"),
            "message: {}",
            err.message,
        );
        assert!(
            err.hint.as_deref().unwrap_or("").contains("Reconnect"),
            "hint should mention reconnect: {:?}",
            err.hint,
        );
    }
}
