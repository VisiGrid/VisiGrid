//! `vgrid fetch xero` — fetch bank transactions from Xero into canonical CSV.

use std::path::PathBuf;

use chrono::{NaiveDate, Utc};
use serde::{Deserialize, Serialize};

use crate::exit_codes;
use crate::CliError;

use super::common::{self, parse_money_string, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const XERO_API_BASE: &str = "https://api.xero.com/api.xro/2.0";
const XERO_TOKEN_URL: &str = "https://identity.xero.com/connect/token";
const XERO_PAGE_SIZE: u32 = 100;

// ── Credentials ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct XeroCredentials {
    pub client_id: String,
    pub client_secret: String,
    pub access_token: String,
    pub refresh_token: String,
    pub tenant_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub access_token_expires_at: Option<String>,
}

fn load_credentials(path: &PathBuf) -> Result<XeroCredentials, CliError> {
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

fn save_credentials(creds: &XeroCredentials, path: &PathBuf) -> Result<(), CliError> {
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
    creds: &XeroCredentials,
    http: &reqwest::blocking::Client,
    token_url: &str,
) -> Result<XeroCredentials, CliError> {
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
            message: format!("Xero token refresh request failed: {}", e),
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
                "Xero token refresh failed ({}): {}",
                status, msg,
            ),
            hint: Some(
                "Refresh token expired or revoked. Reconnect Xero \
                 in your OAuth app (or VisiHub settings) to regenerate credentials."
                    .into(),
            ),
        });
    }

    let body: serde_json::Value = resp.json().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_AUTH,
        message: format!("Xero token refresh response invalid: {}", e),
        hint: None,
    })?;

    let new_access = body["access_token"]
        .as_str()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: "Xero token refresh response missing access_token".into(),
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

    Ok(XeroCredentials {
        client_id: creds.client_id.clone(),
        client_secret: creds.client_secret.clone(),
        access_token: new_access.to_string(),
        refresh_token: new_refresh.to_string(),
        tenant_id: creds.tenant_id.clone(),
        access_token_expires_at,
    })
}

// ── Date parsing ─────────────────────────────────────────────────────

/// Parse Xero's .NET `/Date(1609459200000+0000)/` format to YYYY-MM-DD.
fn parse_xero_date(s: &str) -> Option<String> {
    // Format: /Date(millis+offset)/ or /Date(millis)/
    let inner = s.strip_prefix("/Date(")?.strip_suffix(")/")?;
    // Strip timezone offset if present (e.g. "+0000")
    let millis_str = inner.split('+').next()?.split('-').next()?;
    let millis: i64 = millis_str.parse().ok()?;
    let secs = millis / 1000;
    let dt = chrono::DateTime::from_timestamp(secs, 0)?;
    Some(dt.format("%Y-%m-%d").to_string())
}

// ── Amount parsing ───────────────────────────────────────────────────

/// Extract a Xero amount from a serde_json::Value (typically f64).
/// Formats to 2 decimal places then reuses parse_money_string for cents conversion.
fn extract_amount(val: &serde_json::Value) -> Result<i64, String> {
    if let Some(s) = val.as_str() {
        parse_money_string(s)
    } else if val.is_number() {
        parse_money_string(&format!("{:.2}", val.as_f64().unwrap_or(0.0)))
    } else {
        Err(format!("expected number or string, got {:?}", val))
    }
}

// ── Internal transaction row ────────────────────────────────────────

#[derive(Debug)]
struct RawXeroTransaction {
    effective_date: String,
    amount_minor: i64,
    currency: String,
    txn_type: String,
    source_id: String,
    description: String,
}

fn extract_xero_error(body: &serde_json::Value, status: u16) -> String {
    body["Detail"]
        .as_str()
        .or_else(|| body["Message"].as_str())
        .or_else(|| body["Title"].as_str())
        .or_else(|| body["message"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entity parsing ──────────────────────────────────────────────────

fn parse_bank_transaction(entity: &serde_json::Value) -> Result<RawXeroTransaction, CliError> {
    let id = entity["BankTransactionID"].as_str().unwrap_or("").to_string();
    let txn_type_raw = entity["Type"].as_str().unwrap_or("");

    let date_str = entity["Date"]
        .as_str()
        .and_then(|s| parse_xero_date(s))
        .or_else(|| entity["DateString"].as_str().map(|s| s[..10].to_string()))
        .unwrap_or_default();

    let total = extract_amount(&entity["Total"]).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("bank transaction {} bad Total: {}", id, e),
        hint: None,
    })?;

    let currency = entity["CurrencyCode"]
        .as_str()
        .unwrap_or("USD")
        .to_string();

    // RECEIVE = money in (positive), SPEND = money out (negative)
    let (signed_amount, canonical_type) = match txn_type_raw {
        "RECEIVE" => (total, "deposit"),
        "SPEND" => (-total, "expense"),
        _ => (total, "other"),
    };

    // Description: Contact.Name → LineItems[0].Description → Reference
    let description = entity["Contact"]["Name"]
        .as_str()
        .filter(|s| !s.is_empty())
        .or_else(|| {
            entity["LineItems"]
                .as_array()
                .and_then(|items| items.first())
                .and_then(|item| item["Description"].as_str())
        })
        .or_else(|| entity["Reference"].as_str())
        .unwrap_or("")
        .to_string();

    Ok(RawXeroTransaction {
        effective_date: date_str,
        amount_minor: signed_amount,
        currency,
        txn_type: canonical_type.to_string(),
        source_id: format!("banktxn:{}", id),
        description,
    })
}

fn parse_bank_transfer(
    entity: &serde_json::Value,
    our_account_id: &str,
) -> Result<RawXeroTransaction, CliError> {
    let id = entity["BankTransferID"].as_str().unwrap_or("").to_string();

    let date_str = entity["Date"]
        .as_str()
        .and_then(|s| parse_xero_date(s))
        .or_else(|| entity["DateString"].as_str().map(|s| s[..10].to_string()))
        .unwrap_or_default();

    let amount = extract_amount(&entity["Amount"]).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("bank transfer {} bad Amount: {}", id, e),
        hint: None,
    })?;

    let from_id = entity["FromBankAccount"]["AccountID"].as_str().unwrap_or("");
    let to_id = entity["ToBankAccount"]["AccountID"].as_str().unwrap_or("");

    let (signed_amount, description) = if from_id == our_account_id {
        let to_name = entity["ToBankAccount"]["Name"]
            .as_str()
            .unwrap_or("unknown");
        (-amount, format!("Transfer to {}", to_name))
    } else if to_id == our_account_id {
        let from_name = entity["FromBankAccount"]["Name"]
            .as_str()
            .unwrap_or("unknown");
        (amount, format!("Transfer from {}", from_name))
    } else {
        (amount, "Transfer".to_string())
    };

    let currency = entity["CurrencyCode"]
        .as_str()
        .unwrap_or("USD")
        .to_string();

    Ok(RawXeroTransaction {
        effective_date: date_str,
        amount_minor: signed_amount,
        currency,
        txn_type: "transfer".to_string(),
        source_id: format!("transfer:{}", id),
        description,
    })
}

// ── Xero client ─────────────────────────────────────────────────────

pub struct XeroClient {
    client: FetchClient,
    access_token: String,
    tenant_id: String,
    base_url: String,
    token_url: String,
    creds: Option<XeroCredentials>,
    creds_path: Option<PathBuf>,
}

impl XeroClient {
    pub fn new(access_token: String, tenant_id: String) -> Self {
        Self::with_base_url(access_token, tenant_id, XERO_API_BASE.to_string())
    }

    pub fn with_base_url(
        access_token: String,
        tenant_id: String,
        base_url: String,
    ) -> Self {
        Self {
            client: FetchClient::new("Xero", extract_xero_error),
            access_token,
            tenant_id,
            base_url,
            token_url: XERO_TOKEN_URL.to_string(),
            creds: None,
            creds_path: None,
        }
    }

    fn from_credentials(
        creds: XeroCredentials,
        path: PathBuf,
    ) -> Self {
        Self {
            client: FetchClient::new("Xero", extract_xero_error),
            access_token: creds.access_token.clone(),
            tenant_id: creds.tenant_id.clone(),
            base_url: XERO_API_BASE.to_string(),
            token_url: XERO_TOKEN_URL.to_string(),
            creds: Some(creds),
            creds_path: Some(path),
        }
    }

    #[cfg(test)]
    fn from_credentials_with_base_url(
        creds: XeroCredentials,
        path: PathBuf,
        base_url: String,
    ) -> Self {
        let token_url = format!("{}/connect/token", base_url);
        Self {
            client: FetchClient::new("Xero", extract_xero_error),
            access_token: creds.access_token.clone(),
            tenant_id: creds.tenant_id.clone(),
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

    /// Build the Xero date filter for `where` clause.
    /// Format: `Date >= DateTime(2026,1,1) AND Date < DateTime(2026,2,1)`
    fn date_filter(from: &NaiveDate, to: &NaiveDate) -> String {
        format!(
            "Date >= DateTime({},{},{}) AND Date < DateTime({},{},{})",
            from.format("%Y"),
            from.format("%-m"),
            from.format("%-d"),
            to.format("%Y"),
            to.format("%-m"),
            to.format("%-d"),
        )
    }

    /// Make a paginated GET request to a Xero endpoint.
    /// Returns all items from all pages.
    fn get_paginated(
        &mut self,
        endpoint: &str,
        where_clause: &str,
        response_key: &str,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        let mut all = Vec::new();
        let mut page = 1u32;
        let mut refreshed = false;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        loop {
            let url = format!("{}/{}", self.base_url, endpoint);
            let token = self.access_token.clone();
            let tenant_id = self.tenant_id.clone();
            let where_str = where_clause.to_string();
            let current_page = page;

            let result = self.client.request_with_retry(|http| {
                http.get(&url)
                    .bearer_auth(&token)
                    .header("Xero-Tenant-Id", &tenant_id)
                    .header("Accept", "application/json")
                    .query(&[
                        ("where", where_str.as_str()),
                        ("page", &current_page.to_string()),
                    ])
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
                    let tenant_id = self.tenant_id.clone();
                    self.client.request_with_retry(|http| {
                        http.get(&url)
                            .bearer_auth(&token)
                            .header("Xero-Tenant-Id", &tenant_id)
                            .header("Accept", "application/json")
                            .query(&[
                                ("where", where_str.as_str()),
                                ("page", &current_page.to_string()),
                            ])
                    })?
                }
                Err(e) => return Err(e),
            };

            let entities = body[response_key]
                .as_array()
                .cloned()
                .unwrap_or_default();

            let count = entities.len() as u32;

            if show_progress {
                eprintln!(
                    "  {} page {}: {} results",
                    endpoint, page, count,
                );
            }

            all.extend(entities);

            if count < XERO_PAGE_SIZE {
                break;
            }

            page += 1;
        }

        Ok(all)
    }

    /// Resolve an account name to its Xero AccountID.
    fn resolve_account(
        &mut self,
        name: &str,
        quiet: bool,
    ) -> Result<String, CliError> {
        let where_clause = format!(
            "Name==\"{}\" AND Type==\"BANK\"",
            name.replace('"', "\\\""),
        );
        let url = format!("{}/Accounts", self.base_url);
        let token = self.access_token.clone();
        let tenant_id = self.tenant_id.clone();

        let body = self.client.request_with_retry(|http| {
            http.get(&url)
                .bearer_auth(&token)
                .header("Xero-Tenant-Id", &tenant_id)
                .header("Accept", "application/json")
                .query(&[("where", where_clause.as_str())])
        })?;

        let accounts = body["Accounts"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        match accounts.len() {
            0 => Err(CliError {
                code: exit_codes::EXIT_FETCH_VALIDATION,
                message: format!(
                    "No bank account named '{}' found in Xero. Use --account-id to specify by ID.",
                    name,
                ),
                hint: None,
            }),
            1 => {
                let id = accounts[0]["AccountID"]
                    .as_str()
                    .unwrap_or("")
                    .to_string();
                if show_progress {
                    let resolved_name = accounts[0]["Name"].as_str().unwrap_or(name);
                    eprintln!("  Resolved account '{}' → ID {}", resolved_name, id);
                }
                Ok(id)
            }
            _ => {
                let ids: Vec<String> = accounts
                    .iter()
                    .map(|a| {
                        let aname = a["Name"].as_str().unwrap_or("?");
                        format!(
                            "{} (ID {})",
                            aname,
                            a["AccountID"].as_str().unwrap_or("?"),
                        )
                    })
                    .collect();
                Err(CliError {
                    code: exit_codes::EXIT_FETCH_VALIDATION,
                    message: format!(
                        "Multiple accounts named '{}': {}. Use --account-id to specify.",
                        name,
                        ids.join(", "),
                    ),
                    hint: None,
                })
            }
        }
    }

    /// Fetch BankTransaction entities for the given account and date range.
    fn fetch_bank_transactions(
        &mut self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawXeroTransaction>, CliError> {
        let where_clause = format!(
            "BankAccount.AccountID=guid(\"{}\") AND {}",
            account_id,
            Self::date_filter(from, to),
        );
        let entities = self.get_paginated(
            "BankTransactions",
            &where_clause,
            "BankTransactions",
            quiet,
        )?;
        entities.iter().map(parse_bank_transaction).collect()
    }

    /// Fetch BankTransfer entities for the given account and date range.
    /// Queries both FromBankAccount and ToBankAccount, deduplicating by ID.
    fn fetch_bank_transfers(
        &mut self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawXeroTransaction>, CliError> {
        let mut all = Vec::new();

        // Transfers FROM this account
        let where_from = format!(
            "FromBankAccount.AccountID=guid(\"{}\") AND {}",
            account_id,
            Self::date_filter(from, to),
        );
        let from_entities = self.get_paginated(
            "BankTransfers",
            &where_from,
            "BankTransfers",
            quiet,
        )?;
        for entity in &from_entities {
            all.push(parse_bank_transfer(entity, account_id)?);
        }

        // Transfers TO this account
        let where_to = format!(
            "ToBankAccount.AccountID=guid(\"{}\") AND {}",
            account_id,
            Self::date_filter(from, to),
        );
        let to_entities = self.get_paginated(
            "BankTransfers",
            &where_to,
            "BankTransfers",
            quiet,
        )?;
        for entity in &to_entities {
            let id = entity["BankTransferID"].as_str().unwrap_or("");
            let source_id = format!("transfer:{}", id);
            if !all.iter().any(|t| t.source_id == source_id) {
                all.push(parse_bank_transfer(entity, account_id)?);
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
            "transaction" | "transfer" => types.push(t),
            "" => {}
            other => {
                return Err(CliError {
                    code: exit_codes::EXIT_USAGE,
                    message: format!(
                        "Unknown entity type '{}'. Valid: transaction, transfer",
                        other,
                    ),
                    hint: None,
                });
            }
        }
    }
    if types.is_empty() {
        types = vec![
            "transaction".to_string(),
            "transfer".to_string(),
        ];
    }
    Ok(types)
}

// ── Entry point ─────────────────────────────────────────────────────

#[allow(clippy::too_many_arguments)]
pub fn cmd_fetch_xero(
    from: String,
    to: String,
    credentials: Option<PathBuf>,
    access_token: Option<String>,
    tenant_id: Option<String>,
    account: Option<String>,
    account_id: Option<String>,
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
                message: "missing Xero access token (--access-token is empty)".into(),
                hint: None,
            });
        }
        let tid = tenant_id.ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "missing --tenant-id (required with --access-token)".into(),
            hint: None,
        })?;
        let tid = tid.trim().to_string();
        if tid.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing Xero tenant ID (--tenant-id is empty)".into(),
                hint: None,
            });
        }
        XeroClient::new(token, tid)
    } else if let Some(ref path) = credentials {
        let expanded =
            shellexpand::tilde(&path.to_string_lossy()).to_string();
        let creds_path = PathBuf::from(&expanded);
        let creds = load_credentials(&creds_path)?;
        XeroClient::from_credentials(creds, creds_path)
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "Use --credentials or --access-token + --tenant-id".into(),
            hint: None,
        });
    };

    // 3. Parse include list
    let include_types = parse_include(
        include.as_deref().unwrap_or("transaction,transfer"),
    )?;

    // 4. Resolve account
    let resolved_account_id = if let Some(aid) = account_id {
        let aid = aid.trim().to_string();
        if aid.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_VALIDATION,
                message: "--account-id is empty".into(),
                hint: None,
            });
        }
        aid
    } else if let Some(ref name) = account {
        client.resolve_account(name, quiet)?
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_VALIDATION,
            message: "missing account: use --account or --account-id".into(),
            hint: None,
        });
    };

    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Xero transactions ({} to {}, account {})...",
            from_date, to_date, resolved_account_id,
        );
    }

    // 5. Query each included entity type
    let mut all_txns: Vec<RawXeroTransaction> = Vec::new();
    let mut transaction_count = 0usize;
    let mut transfer_count = 0usize;

    for entity_type in &include_types {
        match entity_type.as_str() {
            "transaction" => {
                let txns = client.fetch_bank_transactions(
                    &resolved_account_id,
                    &from_date,
                    &to_date,
                    quiet,
                )?;
                transaction_count = txns.len();
                all_txns.extend(txns);
            }
            "transfer" => {
                let txns = client.fetch_bank_transfers(
                    &resolved_account_id,
                    &from_date,
                    &to_date,
                    quiet,
                )?;
                transfer_count = txns.len();
                all_txns.extend(txns);
            }
            _ => unreachable!(),
        }
    }

    // 6. Sort: (effective_date ASC, source_id ASC)
    all_txns.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 7. Build canonical rows
    let rows: Vec<CanonicalRow> = all_txns
        .iter()
        .map(|r| CanonicalRow {
            effective_date: r.effective_date.clone(),
            posted_date: r.effective_date.clone(),
            amount_minor: r.amount_minor,
            currency: r.currency.clone(),
            r#type: r.txn_type.clone(),
            source: "xero".to_string(),
            source_id: r.source_id.clone(),
            group_id: String::new(),
            description: r.description.clone(),
        })
        .collect();

    // 8. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!(
            "Done: {} transactions + {} transfers = {} rows written to {}",
            transaction_count, transfer_count, rows.len(), out_label,
        );
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── parse_xero_date ─────────────────────────────────────────────

    #[test]
    fn test_parse_xero_date_basic() {
        // 2021-01-01 00:00:00 UTC = 1609459200000 ms
        assert_eq!(
            parse_xero_date("/Date(1609459200000+0000)/"),
            Some("2021-01-01".to_string()),
        );
    }

    #[test]
    fn test_parse_xero_date_no_offset() {
        assert_eq!(
            parse_xero_date("/Date(1609459200000)/"),
            Some("2021-01-01".to_string()),
        );
    }

    #[test]
    fn test_parse_xero_date_invalid() {
        assert_eq!(parse_xero_date("2026-01-01"), None);
        assert_eq!(parse_xero_date("not-a-date"), None);
        assert_eq!(parse_xero_date(""), None);
    }

    #[test]
    fn test_parse_xero_date_mid_year() {
        // 2026-06-15 12:00:00 UTC = 1781870400000 ms
        let result = parse_xero_date("/Date(1781481600000+0000)/");
        assert_eq!(result, Some("2026-06-15".to_string()));
    }

    // ── parse_bank_transaction ──────────────────────────────────────

    #[test]
    fn test_parse_bank_transaction_receive() {
        let entity = serde_json::json!({
            "BankTransactionID": "abc-123",
            "Type": "RECEIVE",
            "Date": "/Date(1769904000000+0000)/",
            "Total": 2500.00,
            "CurrencyCode": "USD",
            "Contact": { "Name": "Client Corp" },
            "LineItems": [
                { "Description": "Invoice payment" }
            ],
            "BankAccount": { "AccountID": "bank-001" }
        });

        let txn = parse_bank_transaction(&entity).unwrap();
        assert_eq!(txn.effective_date, "2026-02-01");
        assert_eq!(txn.amount_minor, 250000); // positive
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.txn_type, "deposit");
        assert_eq!(txn.source_id, "banktxn:abc-123");
        assert_eq!(txn.description, "Client Corp");
    }

    #[test]
    fn test_parse_bank_transaction_spend() {
        let entity = serde_json::json!({
            "BankTransactionID": "def-456",
            "Type": "SPEND",
            "Date": "/Date(1769904000000+0000)/",
            "Total": 49.99,
            "CurrencyCode": "USD",
            "Contact": { "Name": "Office Supplies" },
            "LineItems": [],
            "BankAccount": { "AccountID": "bank-001" }
        });

        let txn = parse_bank_transaction(&entity).unwrap();
        assert_eq!(txn.amount_minor, -4999); // negative
        assert_eq!(txn.txn_type, "expense");
        assert_eq!(txn.source_id, "banktxn:def-456");
        assert_eq!(txn.description, "Office Supplies");
    }

    #[test]
    fn test_parse_bank_transaction_description_fallback() {
        // No Contact.Name → falls back to LineItems[0].Description
        let entity = serde_json::json!({
            "BankTransactionID": "ghi-789",
            "Type": "RECEIVE",
            "DateString": "2026-01-15T00:00:00",
            "Total": "1234.56",
            "LineItems": [
                { "Description": "Wire from client" }
            ]
        });

        let txn = parse_bank_transaction(&entity).unwrap();
        assert_eq!(txn.amount_minor, 123456);
        assert_eq!(txn.description, "Wire from client");
        assert_eq!(txn.effective_date, "2026-01-15");
    }

    #[test]
    fn test_parse_bank_transaction_reference_fallback() {
        // No Contact, no LineItems → falls back to Reference
        let entity = serde_json::json!({
            "BankTransactionID": "jkl-012",
            "Type": "SPEND",
            "DateString": "2026-01-20T00:00:00",
            "Total": 100.00,
            "Reference": "REF-12345"
        });

        let txn = parse_bank_transaction(&entity).unwrap();
        assert_eq!(txn.description, "REF-12345");
    }

    // ── parse_bank_transfer ────────────────────────────────────────

    #[test]
    fn test_parse_bank_transfer_outbound() {
        let entity = serde_json::json!({
            "BankTransferID": "xfer-001",
            "Date": "/Date(1769904000000+0000)/",
            "Amount": 5000.00,
            "FromBankAccount": { "AccountID": "bank-001", "Name": "Checking" },
            "ToBankAccount": { "AccountID": "bank-002", "Name": "Savings" },
            "CurrencyCode": "USD"
        });

        let txn = parse_bank_transfer(&entity, "bank-001").unwrap();
        assert_eq!(txn.amount_minor, -500000); // outbound = negative
        assert_eq!(txn.txn_type, "transfer");
        assert_eq!(txn.source_id, "transfer:xfer-001");
        assert_eq!(txn.description, "Transfer to Savings");
    }

    #[test]
    fn test_parse_bank_transfer_inbound() {
        let entity = serde_json::json!({
            "BankTransferID": "xfer-002",
            "Date": "/Date(1769904000000+0000)/",
            "Amount": 3000.00,
            "FromBankAccount": { "AccountID": "bank-002", "Name": "Savings" },
            "ToBankAccount": { "AccountID": "bank-001", "Name": "Checking" },
            "CurrencyCode": "USD"
        });

        let txn = parse_bank_transfer(&entity, "bank-001").unwrap();
        assert_eq!(txn.amount_minor, 300000); // inbound = positive
        assert_eq!(txn.description, "Transfer from Savings");
    }

    // ── Golden transfer fixture ────────────────────────────────────

    fn golden_transfer() -> serde_json::Value {
        serde_json::json!({
            "BankTransferID": "xfer-400",
            "Date": "/Date(1768867200000+0000)/",
            "Amount": 5000.00,
            "FromBankAccount": { "AccountID": "bank-001", "Name": "Checking" },
            "ToBankAccount":   { "AccountID": "bank-002", "Name": "Savings" },
            "CurrencyCode": "USD"
        })
    }

    #[test]
    fn test_golden_transfer_from_checking_perspective() {
        let txn = parse_bank_transfer(&golden_transfer(), "bank-001").unwrap();
        assert_eq!(txn.amount_minor, -500000, "Checking→Savings should be negative from Checking's perspective");
        assert_eq!(txn.description, "Transfer to Savings");
    }

    #[test]
    fn test_golden_transfer_from_savings_perspective() {
        let txn = parse_bank_transfer(&golden_transfer(), "bank-002").unwrap();
        assert_eq!(txn.amount_minor, 500000, "Checking→Savings should be positive from Savings' perspective");
        assert_eq!(txn.description, "Transfer from Checking");
    }

    // ── Account resolution (httpmock) ──────────────────────────────

    #[test]
    fn test_account_resolution_single() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/Accounts");
            then.status(200)
                .json_body(serde_json::json!({
                    "Accounts": [
                        { "AccountID": "bank-001", "Name": "Checking", "Type": "BANK" }
                    ]
                }));
        });

        let mut client = XeroClient::with_base_url(
            "test_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let id = client.resolve_account("Checking", true).unwrap();
        assert_eq!(id, "bank-001");
    }

    #[test]
    fn test_account_resolution_ambiguous() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/Accounts");
            then.status(200)
                .json_body(serde_json::json!({
                    "Accounts": [
                        { "AccountID": "bank-001", "Name": "Checking", "Type": "BANK" },
                        { "AccountID": "bank-003", "Name": "Checking", "Type": "BANK" }
                    ]
                }));
        });

        let mut client = XeroClient::with_base_url(
            "test_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let err = client.resolve_account("Checking", true).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_VALIDATION);
        assert!(err.message.contains("Multiple accounts"), "message: {}", err.message);
        assert!(err.message.contains("bank-001"), "should include IDs: {}", err.message);
        assert!(err.message.contains("bank-003"), "should include IDs: {}", err.message);
    }

    #[test]
    fn test_account_resolution_not_found() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/Accounts");
            then.status(200)
                .json_body(serde_json::json!({
                    "Accounts": []
                }));
        });

        let mut client = XeroClient::with_base_url(
            "test_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let err = client.resolve_account("Nonexistent", true).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_VALIDATION);
        assert!(err.message.contains("No bank account named"));
    }

    // ── Pagination (httpmock) ──────────────────────────────────────

    #[test]
    fn test_pagination_single_page() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions");
            then.status(200)
                .json_body(serde_json::json!({
                    "BankTransactions": (0..50).map(|i| {
                        serde_json::json!({
                            "BankTransactionID": format!("txn-{}", i),
                            "Type": "RECEIVE",
                            "DateString": "2026-01-15T00:00:00",
                            "Total": 100.00,
                            "BankAccount": { "AccountID": "bank-001" }
                        })
                    }).collect::<Vec<_>>()
                }));
        });

        let mut client = XeroClient::with_base_url(
            "test_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_bank_transactions("bank-001", &from, &to, true).unwrap();
        assert_eq!(txns.len(), 50);
    }

    #[test]
    fn test_empty_result_set() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions");
            then.status(200)
                .json_body(serde_json::json!({
                    "BankTransactions": []
                }));
        });

        let mut client = XeroClient::with_base_url(
            "test_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_bank_transactions("bank-001", &from, &to, true).unwrap();
        assert_eq!(txns.len(), 0);
    }

    // ── Auth failure (httpmock) ────────────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions");
            then.status(401)
                .json_body(serde_json::json!({
                    "Title": "Unauthorized",
                    "Detail": "Token expired"
                }));
        });

        let mut client = XeroClient::with_base_url(
            "bad_token".into(),
            "tenant_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_bank_transactions("bank-001", &from, &to, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("Xero auth failed (401)"),
            "message: {}",
            err.message,
        );
    }

    // ── Token refresh on 401 (httpmock) ────────────────────────────

    #[test]
    fn test_token_refresh_on_401() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions")
                .header("Authorization", "Bearer old_token");
            then.status(401)
                .json_body(serde_json::json!({
                    "Title": "Unauthorized",
                    "Detail": "Token expired"
                }));
        });

        // Refresh → new token
        server.mock(|when, then| {
            when.method(POST)
                .path("/connect/token");
            then.status(200)
                .json_body(serde_json::json!({
                    "access_token": "new_token",
                    "refresh_token": "new_refresh",
                    "expires_in": 1800
                }));
        });

        // Retry with new token → success
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions")
                .header("Authorization", "Bearer new_token");
            then.status(200)
                .json_body(serde_json::json!({
                    "BankTransactions": [{
                        "BankTransactionID": "txn-101",
                        "Type": "RECEIVE",
                        "DateString": "2026-01-15T00:00:00",
                        "Total": 500.00,
                        "BankAccount": { "AccountID": "bank-001" }
                    }]
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("xero.json");
        let creds = XeroCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "old_refresh".into(),
            tenant_id: "tenant_123".into(),
            access_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = XeroClient::from_credentials_with_base_url(
            creds,
            creds_path.clone(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_bank_transactions("bank-001", &from, &to, true).unwrap();

        assert_eq!(txns.len(), 1);
        assert_eq!(txns[0].amount_minor, 50000);

        // Verify credentials were updated on disk
        let saved: XeroCredentials = serde_json::from_str(
            &std::fs::read_to_string(&creds_path).unwrap(),
        )
        .unwrap();
        assert_eq!(saved.access_token, "new_token");
        assert_eq!(saved.refresh_token, "new_refresh");
    }

    // ── Credential loading ─────────────────────────────────────────

    #[test]
    fn test_credential_loading_valid() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("xero.json");
        std::fs::write(
            &path,
            serde_json::json!({
                "client_id": "cid",
                "client_secret": "csec",
                "access_token": "at",
                "refresh_token": "rt",
                "tenant_id": "tid"
            })
            .to_string(),
        )
        .unwrap();
        let creds = load_credentials(&path.to_path_buf()).unwrap();
        assert_eq!(creds.client_id, "cid");
        assert_eq!(creds.tenant_id, "tid");
        assert!(creds.access_token_expires_at.is_none());
    }

    #[test]
    fn test_credential_loading_missing_file() {
        let path = PathBuf::from("/tmp/nonexistent-xero-creds.json");
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

    // ── Include filter ─────────────────────────────────────────────

    #[test]
    fn test_include_filter() {
        let types = parse_include("transaction").unwrap();
        assert_eq!(types, vec!["transaction"]);

        let types = parse_include("transaction,transfer").unwrap();
        assert_eq!(types, vec!["transaction", "transfer"]);

        let types = parse_include("transfer").unwrap();
        assert_eq!(types, vec!["transfer"]);

        let err = parse_include("transaction,invoice").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_USAGE);
        assert!(err.message.contains("invoice"));
    }

    // ── Refresh failure hint ───────────────────────────────────────

    #[test]
    fn test_refresh_failure_includes_reconnect_hint() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/BankTransactions");
            then.status(401)
                .json_body(serde_json::json!({
                    "Title": "Unauthorized"
                }));
        });

        // Refresh → revoked
        server.mock(|when, then| {
            when.method(POST)
                .path("/connect/token");
            then.status(400)
                .json_body(serde_json::json!({
                    "error": "invalid_grant",
                    "error_description": "refresh token has been revoked"
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("xero.json");
        let creds = XeroCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "revoked".into(),
            tenant_id: "tenant_123".into(),
            access_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = XeroClient::from_credentials_with_base_url(
            creds,
            creds_path,
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_bank_transactions("bank-001", &from, &to, true)
            .unwrap_err();

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

    // ── Date filter format ─────────────────────────────────────────

    #[test]
    fn test_date_filter_format() {
        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 2, 1).unwrap();
        let filter = XeroClient::date_filter(&from, &to);
        assert_eq!(
            filter,
            "Date >= DateTime(2026,1,1) AND Date < DateTime(2026,2,1)",
        );
    }

    #[test]
    fn test_date_filter_double_digit() {
        let from = NaiveDate::from_ymd_opt(2026, 11, 15).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 12, 15).unwrap();
        let filter = XeroClient::date_filter(&from, &to);
        assert_eq!(
            filter,
            "Date >= DateTime(2026,11,15) AND Date < DateTime(2026,12,15)",
        );
    }
}
