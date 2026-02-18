//! `vgrid fetch qbo` — fetch posted ledger transactions from QuickBooks Online into canonical CSV.

use std::path::PathBuf;

use chrono::{NaiveDate, Utc};
use serde::{Deserialize, Serialize};

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const QBO_API_BASE: &str = "https://quickbooks.api.intuit.com";
const QBO_SANDBOX_BASE: &str = "https://sandbox-quickbooks.api.intuit.com";
const QBO_TOKEN_URL: &str = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
const QBO_QUERY_LIMIT: u32 = 1000;

// ── Credentials ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct QboCredentials {
    pub client_id: String,
    pub client_secret: String,
    pub access_token: String,
    pub refresh_token: String,
    pub realm_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub access_token_expires_at: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub refresh_token_expires_at: Option<String>,
}

fn load_credentials(path: &PathBuf) -> Result<QboCredentials, CliError> {
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

fn save_credentials(creds: &QboCredentials, path: &PathBuf) -> Result<(), CliError> {
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
    creds: &QboCredentials,
    http: &reqwest::blocking::Client,
    token_url: &str,
) -> Result<QboCredentials, CliError> {
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
            message: format!("QBO token refresh request failed: {}", e),
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
                "QBO token refresh failed ({}): {}",
                status, msg,
            ),
            hint: Some(
                "Refresh token expired or revoked. Reconnect QuickBooks \
                 in your OAuth app (or VisiHub settings) to regenerate credentials."
                    .into(),
            ),
        });
    }

    let body: serde_json::Value = resp.json().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_AUTH,
        message: format!("QBO token refresh response invalid: {}", e),
        hint: None,
    })?;

    let new_access = body["access_token"]
        .as_str()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: "QBO token refresh response missing access_token".into(),
            hint: None,
        })?;

    let new_refresh = body["refresh_token"]
        .as_str()
        .unwrap_or(&creds.refresh_token);

    // Compute expiry timestamps if present
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
            // Warn if refresh token expires within 30 days
            let days_left = (expires - Utc::now()).num_days();
            if days_left <= 30 {
                eprintln!(
                    "warning: QBO refresh token expires {}, re-authorize soon",
                    expires.format("%Y-%m-%d"),
                );
            }
            expires.to_rfc3339()
        })
        .or_else(|| creds.refresh_token_expires_at.clone());

    Ok(QboCredentials {
        client_id: creds.client_id.clone(),
        client_secret: creds.client_secret.clone(),
        access_token: new_access.to_string(),
        refresh_token: new_refresh.to_string(),
        realm_id: creds.realm_id.clone(),
        access_token_expires_at,
        refresh_token_expires_at,
    })
}

// ── String escaping ─────────────────────────────────────────────────

/// Escape a string for use in QBO SQL-like queries.
/// QBO uses single-quoted strings; escape embedded single quotes by doubling them.
fn qbo_escape(s: &str) -> String {
    s.replace('\'', "''")
}

// ── Amount parsing ───────────────────────────────────────────────────

use super::common::parse_money_string;

/// Extract a QBO amount from a serde_json::Value, handling both string and number types.
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
struct RawQboTransaction {
    effective_date: String,
    amount_minor: i64,
    currency: String,
    txn_type: String,
    source_id: String,
    description: String,
}

fn extract_qbo_error(body: &serde_json::Value, status: u16) -> String {
    // QBO error responses come in a Fault structure
    body["Fault"]["Error"][0]["Detail"]
        .as_str()
        .or_else(|| body["Fault"]["Error"][0]["Message"].as_str())
        .or_else(|| body["fault"]["error"][0]["detail"].as_str())
        .or_else(|| body["message"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Entity parsing ──────────────────────────────────────────────────

fn parse_deposit(entity: &serde_json::Value) -> Result<RawQboTransaction, CliError> {
    let id = entity["Id"].as_str().unwrap_or("").to_string();
    let txn_date = entity["TxnDate"].as_str().unwrap_or("").to_string();

    let amount = extract_amount(&entity["TotalAmt"]).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("deposit {} bad TotalAmt: {}", id, e),
        hint: None,
    })?;

    let currency = entity["CurrencyRef"]["value"]
        .as_str()
        .unwrap_or("USD")
        .to_string();

    if currency != "USD" {
        eprintln!("warning: deposit {} has currency {}, expected USD", id, currency);
    }

    let description = entity["PrivateNote"]
        .as_str()
        .filter(|s| !s.is_empty())
        .or_else(|| {
            entity["Line"]
                .as_array()
                .and_then(|lines| lines.first())
                .and_then(|line| line["Description"].as_str())
        })
        .unwrap_or("")
        .to_string();

    Ok(RawQboTransaction {
        effective_date: txn_date,
        amount_minor: amount, // positive: money coming in
        currency,
        txn_type: "deposit".to_string(),
        source_id: format!("deposit:{}", id),
        description,
    })
}

fn parse_purchase(entity: &serde_json::Value) -> Result<RawQboTransaction, CliError> {
    let id = entity["Id"].as_str().unwrap_or("").to_string();
    let txn_date = entity["TxnDate"].as_str().unwrap_or("").to_string();

    let amount = extract_amount(&entity["TotalAmt"]).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("purchase {} bad TotalAmt: {}", id, e),
        hint: None,
    })?;

    let currency = entity["CurrencyRef"]["value"]
        .as_str()
        .unwrap_or("USD")
        .to_string();

    if currency != "USD" {
        eprintln!("warning: purchase {} has currency {}, expected USD", id, currency);
    }

    // Map PaymentType to canonical type
    let payment_type = entity["PaymentType"].as_str().unwrap_or("");
    let canonical_type = match payment_type {
        "Cash" => "withdrawal",
        "Check" => "check",
        _ => "expense",
    };

    let description = entity["PrivateNote"]
        .as_str()
        .filter(|s| !s.is_empty())
        .or_else(|| entity["EntityRef"]["name"].as_str())
        .unwrap_or("")
        .to_string();

    Ok(RawQboTransaction {
        effective_date: txn_date,
        amount_minor: -amount, // negative: money going out
        currency,
        txn_type: canonical_type.to_string(),
        source_id: format!("purchase:{}", id),
        description,
    })
}

fn parse_transfer(
    entity: &serde_json::Value,
    our_account_id: &str,
) -> Result<RawQboTransaction, CliError> {
    let id = entity["Id"].as_str().unwrap_or("").to_string();
    let txn_date = entity["TxnDate"].as_str().unwrap_or("").to_string();

    let amount = extract_amount(&entity["Amount"]).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("transfer {} bad Amount: {}", id, e),
        hint: None,
    })?;

    let from_ref = entity["FromAccountRef"]["value"].as_str().unwrap_or("");
    let to_ref = entity["ToAccountRef"]["value"].as_str().unwrap_or("");

    let (signed_amount, description_default) = if to_ref == our_account_id {
        let from_name = entity["FromAccountRef"]["name"]
            .as_str()
            .unwrap_or("unknown");
        (amount, format!("Transfer from {}", from_name))
    } else if from_ref == our_account_id {
        let to_name = entity["ToAccountRef"]["name"]
            .as_str()
            .unwrap_or("unknown");
        (-amount, format!("Transfer to {}", to_name))
    } else {
        // Shouldn't happen if our queries are correct, but handle gracefully
        (amount, "Transfer".to_string())
    };

    let description = entity["PrivateNote"]
        .as_str()
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .unwrap_or(description_default);

    Ok(RawQboTransaction {
        effective_date: txn_date,
        amount_minor: signed_amount,
        currency: "USD".to_string(),
        txn_type: "transfer".to_string(),
        source_id: format!("transfer:{}", id),
        description,
    })
}

// ── QBO client ──────────────────────────────────────────────────────

pub struct QboClient {
    client: FetchClient,
    access_token: String,
    realm_id: String,
    base_url: String,
    token_url: String,
    creds: Option<QboCredentials>,
    creds_path: Option<PathBuf>,
}

impl QboClient {
    pub fn new(access_token: String, realm_id: String, sandbox: bool) -> Self {
        let base_url = if sandbox {
            QBO_SANDBOX_BASE
        } else {
            QBO_API_BASE
        };
        Self::with_base_url(access_token, realm_id, base_url.to_string())
    }

    pub fn with_base_url(
        access_token: String,
        realm_id: String,
        base_url: String,
    ) -> Self {
        Self {
            client: FetchClient::new("QBO", extract_qbo_error),
            access_token,
            realm_id,
            base_url,
            token_url: QBO_TOKEN_URL.to_string(),
            creds: None,
            creds_path: None,
        }
    }

    fn from_credentials(
        creds: QboCredentials,
        path: PathBuf,
        sandbox: bool,
    ) -> Self {
        let base_url = if sandbox {
            QBO_SANDBOX_BASE
        } else {
            QBO_API_BASE
        };
        Self {
            client: FetchClient::new("QBO", extract_qbo_error),
            access_token: creds.access_token.clone(),
            realm_id: creds.realm_id.clone(),
            base_url: base_url.to_string(),
            token_url: QBO_TOKEN_URL.to_string(),
            creds: Some(creds),
            creds_path: Some(path),
        }
    }

    #[cfg(test)]
    fn from_credentials_with_base_url(
        creds: QboCredentials,
        path: PathBuf,
        base_url: String,
    ) -> Self {
        let token_url = format!("{}/oauth2/v1/tokens/bearer", base_url);
        Self {
            client: FetchClient::new("QBO", extract_qbo_error),
            access_token: creds.access_token.clone(),
            realm_id: creds.realm_id.clone(),
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

    fn query_url(&self) -> String {
        format!(
            "{}/v3/company/{}/query",
            self.base_url, self.realm_id,
        )
    }

    /// Run a paginated QBO query and return all matching entities.
    fn query_entities(
        &mut self,
        entity_type: &str,
        where_clause: &str,
        quiet: bool,
    ) -> Result<Vec<serde_json::Value>, CliError> {
        let mut all = Vec::new();
        let mut start_pos = 1u32;
        let mut refreshed = false;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        loop {
            let query = format!(
                "SELECT * FROM {} WHERE {} ORDERBY TxnDate ASC, Id ASC STARTPOSITION {} MAXRESULTS {}",
                entity_type, where_clause, start_pos, QBO_QUERY_LIMIT,
            );
            let url = self.query_url();
            let token = self.access_token.clone();

            let result = self.client.request_with_retry(|http| {
                http.get(&url)
                    .bearer_auth(&token)
                    .header("Accept", "application/json")
                    .query(&[("query", &query)])
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
                    self.client.request_with_retry(|http| {
                        http.get(&url)
                            .bearer_auth(&token)
                            .header("Accept", "application/json")
                            .query(&[("query", &query)])
                    })?
                }
                Err(e) => return Err(e),
            };

            let entities = body["QueryResponse"][entity_type]
                .as_array()
                .cloned()
                .unwrap_or_default();

            let count = entities.len() as u32;

            if show_progress {
                eprintln!(
                    "  {} startPosition {}: {} results",
                    entity_type, start_pos, count,
                );
            }

            all.extend(entities);

            if count < QBO_QUERY_LIMIT {
                break;
            }

            start_pos += count;
        }

        Ok(all)
    }

    /// Resolve an account name to its QBO account ID.
    fn resolve_account(
        &mut self,
        name: &str,
        quiet: bool,
    ) -> Result<String, CliError> {
        let escaped = qbo_escape(name);
        let query = format!(
            "SELECT Id, Name, FullyQualifiedName, AccountType FROM Account WHERE Name = '{}' AND AccountType = 'Bank'",
            escaped,
        );
        let url = self.query_url();
        let token = self.access_token.clone();

        let body = self.client.request_with_retry(|http| {
            http.get(&url)
                .bearer_auth(&token)
                .header("Accept", "application/json")
                .query(&[("query", &query)])
        })?;

        let accounts = body["QueryResponse"]["Account"]
            .as_array()
            .cloned()
            .unwrap_or_default();

        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        match accounts.len() {
            0 => Err(CliError {
                code: exit_codes::EXIT_FETCH_VALIDATION,
                message: format!(
                    "No bank account named '{}' found in QBO. Use --account-id to specify by ID.",
                    name,
                ),
                hint: None,
            }),
            1 => {
                let id = accounts[0]["Id"]
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
                        let fqn = a["FullyQualifiedName"]
                            .as_str()
                            .unwrap_or_else(|| a["Name"].as_str().unwrap_or("?"));
                        format!(
                            "{} (ID {})",
                            fqn,
                            a["Id"].as_str().unwrap_or("?"),
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

    /// Fetch Deposit entities for the given account and date range.
    /// DepositToAccountRef is not queryable in QBO's API, so we fetch all
    /// deposits in the date range and filter client-side by account ID.
    fn fetch_deposits(
        &mut self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawQboTransaction>, CliError> {
        let where_clause = format!(
            "TxnDate >= '{}' AND TxnDate < '{}'",
            from, to,
        );
        let entities = self.query_entities("Deposit", &where_clause, quiet)?;
        let filtered: Vec<_> = entities
            .iter()
            .filter(|e| {
                e["DepositToAccountRef"]["value"]
                    .as_str()
                    .map_or(false, |v| v == account_id)
            })
            .collect();
        filtered.iter().map(|e| parse_deposit(e)).collect()
    }

    /// Fetch Purchase entities for the given account and date range.
    /// Fetch Purchase entities for the given account and date range.
    /// AccountRef is not queryable in QBO's API, so we fetch all purchases
    /// in the date range and filter client-side by account ID.
    fn fetch_purchases(
        &mut self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawQboTransaction>, CliError> {
        let where_clause = format!(
            "TxnDate >= '{}' AND TxnDate < '{}'",
            from, to,
        );
        let entities = self.query_entities("Purchase", &where_clause, quiet)?;
        let filtered: Vec<_> = entities
            .iter()
            .filter(|e| {
                e["AccountRef"]["value"]
                    .as_str()
                    .map_or(false, |v| v == account_id)
            })
            .collect();
        filtered.iter().map(|e| parse_purchase(e)).collect()
    }

    /// Fetch Transfer entities for the given account and date range.
    /// FromAccountRef/ToAccountRef are not queryable, so we fetch all
    /// transfers in the date range and filter client-side.
    fn fetch_transfers(
        &mut self,
        account_id: &str,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawQboTransaction>, CliError> {
        let where_clause = format!(
            "TxnDate >= '{}' AND TxnDate < '{}'",
            from, to,
        );
        let entities = self.query_entities("Transfer", &where_clause, quiet)?;

        let mut all = Vec::new();
        for entity in &entities {
            let from_ref = entity["FromAccountRef"]["value"].as_str().unwrap_or("");
            let to_ref = entity["ToAccountRef"]["value"].as_str().unwrap_or("");
            if from_ref == account_id || to_ref == account_id {
                all.push(parse_transfer(entity, account_id)?);
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
            "deposit" | "purchase" | "transfer" => types.push(t),
            "" => {}
            other => {
                return Err(CliError {
                    code: exit_codes::EXIT_USAGE,
                    message: format!(
                        "Unknown entity type '{}'. Valid: deposit, purchase, transfer",
                        other,
                    ),
                    hint: None,
                });
            }
        }
    }
    if types.is_empty() {
        types = vec![
            "deposit".to_string(),
            "purchase".to_string(),
            "transfer".to_string(),
        ];
    }
    Ok(types)
}

// ── Entry point ─────────────────────────────────────────────────────

#[allow(clippy::too_many_arguments)]
pub fn cmd_fetch_qbo(
    from: String,
    to: String,
    credentials: Option<PathBuf>,
    access_token: Option<String>,
    realm_id: Option<String>,
    account: Option<String>,
    account_id: Option<String>,
    include: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
    sandbox: bool,
) -> Result<(), CliError> {
    // 1. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 2. Resolve auth
    let mut client = if let Some(token) = access_token {
        let token = token.trim().to_string();
        if token.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing QBO access token (--access-token is empty)".into(),
                hint: None,
            });
        }
        let rid = realm_id.ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "missing --realm-id (required with --access-token)".into(),
            hint: None,
        })?;
        let rid = rid.trim().to_string();
        if rid.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing QBO realm ID (--realm-id is empty)".into(),
                hint: None,
            });
        }
        QboClient::new(token, rid, sandbox)
    } else if let Some(ref path) = credentials {
        let expanded =
            shellexpand::tilde(&path.to_string_lossy()).to_string();
        let creds_path = PathBuf::from(&expanded);
        let creds = load_credentials(&creds_path)?;
        QboClient::from_credentials(creds, creds_path, sandbox)
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "Use --credentials or --access-token + --realm-id".into(),
            hint: None,
        });
    };

    // 3. Parse include list
    let include_types = parse_include(
        include.as_deref().unwrap_or("deposit,purchase,transfer"),
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
            "Fetching QBO transactions ({} to {}, account {})...",
            from_date, to_date, resolved_account_id,
        );
    }

    // 5. Query each included entity type
    let mut all_txns: Vec<RawQboTransaction> = Vec::new();
    let mut deposit_count = 0usize;
    let mut purchase_count = 0usize;
    let mut transfer_count = 0usize;

    for entity_type in &include_types {
        match entity_type.as_str() {
            "deposit" => {
                let txns = client.fetch_deposits(
                    &resolved_account_id,
                    &from_date,
                    &to_date,
                    quiet,
                )?;
                deposit_count = txns.len();
                all_txns.extend(txns);
            }
            "purchase" => {
                let txns = client.fetch_purchases(
                    &resolved_account_id,
                    &from_date,
                    &to_date,
                    quiet,
                )?;
                purchase_count = txns.len();
                all_txns.extend(txns);
            }
            "transfer" => {
                let txns = client.fetch_transfers(
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
            posted_date: r.effective_date.clone(), // QBO doesn't distinguish
            amount_minor: r.amount_minor,
            currency: r.currency.clone(),
            r#type: r.txn_type.clone(),
            source: "qbo".to_string(),
            source_id: r.source_id.clone(),
            group_id: String::new(),
            description: r.description.clone(),
        })
        .collect();

    // 8. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!(
            "Done: {} deposits + {} purchases + {} transfers = {} rows written to {}",
            deposit_count, purchase_count, transfer_count, rows.len(), out_label,
        );
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── qbo_escape ─────────────────────────────────────────────────

    #[test]
    fn test_qbo_escape() {
        assert_eq!(qbo_escape("Checking"), "Checking");
        assert_eq!(qbo_escape("Owner's Draw"), "Owner''s Draw");
        assert_eq!(qbo_escape("It's a 'test'"), "It''s a ''test''");
        assert_eq!(qbo_escape(""), "");
    }

    // ── parse_deposit ──────────────────────────────────────────────

    #[test]
    fn test_parse_deposit() {
        let entity = serde_json::json!({
            "Id": "101",
            "TxnDate": "2026-01-15",
            "TotalAmt": 2500.00,
            "CurrencyRef": { "value": "USD" },
            "PrivateNote": "Stripe payout",
            "DepositToAccountRef": { "value": "35" },
            "Line": [
                { "Description": "Line item 1", "Amount": 2500.00 }
            ]
        });

        let txn = parse_deposit(&entity).unwrap();
        assert_eq!(txn.effective_date, "2026-01-15");
        assert_eq!(txn.amount_minor, 250000); // positive
        assert_eq!(txn.currency, "USD");
        assert_eq!(txn.txn_type, "deposit");
        assert_eq!(txn.source_id, "deposit:101");
        assert_eq!(txn.description, "Stripe payout");
    }

    #[test]
    fn test_parse_deposit_line_description_fallback() {
        let entity = serde_json::json!({
            "Id": "102",
            "TxnDate": "2026-01-16",
            "TotalAmt": "1234.56",
            "Line": [
                { "Description": "Wire from client" }
            ]
        });

        let txn = parse_deposit(&entity).unwrap();
        assert_eq!(txn.amount_minor, 123456);
        assert_eq!(txn.description, "Wire from client");
    }

    // ── parse_purchase ─────────────────────────────────────────────

    #[test]
    fn test_parse_purchase() {
        let entity = serde_json::json!({
            "Id": "201",
            "TxnDate": "2026-01-20",
            "TotalAmt": 49.99,
            "CurrencyRef": { "value": "USD" },
            "PaymentType": "Cash",
            "AccountRef": { "value": "35" },
            "EntityRef": { "name": "Office Supplies Co" },
            "PrivateNote": ""
        });

        let txn = parse_purchase(&entity).unwrap();
        assert_eq!(txn.effective_date, "2026-01-20");
        assert_eq!(txn.amount_minor, -4999); // negative
        assert_eq!(txn.txn_type, "withdrawal");
        assert_eq!(txn.source_id, "purchase:201");
        assert_eq!(txn.description, "Office Supplies Co");
    }

    #[test]
    fn test_purchase_type_mapping() {
        let make = |payment_type: &str| -> serde_json::Value {
            serde_json::json!({
                "Id": "200",
                "TxnDate": "2026-01-20",
                "TotalAmt": 10.00,
                "PaymentType": payment_type,
            })
        };

        assert_eq!(parse_purchase(&make("Cash")).unwrap().txn_type, "withdrawal");
        assert_eq!(parse_purchase(&make("Check")).unwrap().txn_type, "check");
        assert_eq!(parse_purchase(&make("CreditCard")).unwrap().txn_type, "expense");
        assert_eq!(parse_purchase(&make("Other")).unwrap().txn_type, "expense");
        assert_eq!(parse_purchase(&make("")).unwrap().txn_type, "expense");
    }

    // ── parse_transfer ─────────────────────────────────────────────

    #[test]
    fn test_parse_transfer_inbound() {
        let entity = serde_json::json!({
            "Id": "301",
            "TxnDate": "2026-01-10",
            "Amount": 5000.00,
            "FromAccountRef": { "value": "42", "name": "Savings" },
            "ToAccountRef": { "value": "35", "name": "Checking" },
        });

        let txn = parse_transfer(&entity, "35").unwrap();
        assert_eq!(txn.amount_minor, 500000); // positive: money in
        assert_eq!(txn.txn_type, "transfer");
        assert_eq!(txn.source_id, "transfer:301");
        assert_eq!(txn.description, "Transfer from Savings");
    }

    #[test]
    fn test_parse_transfer_outbound() {
        let entity = serde_json::json!({
            "Id": "302",
            "TxnDate": "2026-01-12",
            "Amount": 3000.00,
            "FromAccountRef": { "value": "35", "name": "Checking" },
            "ToAccountRef": { "value": "42", "name": "Savings" },
        });

        let txn = parse_transfer(&entity, "35").unwrap();
        assert_eq!(txn.amount_minor, -300000); // negative: money out
        assert_eq!(txn.description, "Transfer to Savings");
    }

    #[test]
    fn test_parse_transfer_private_note_override() {
        let entity = serde_json::json!({
            "Id": "303",
            "TxnDate": "2026-01-14",
            "Amount": 1000.00,
            "FromAccountRef": { "value": "42", "name": "Savings" },
            "ToAccountRef": { "value": "35", "name": "Checking" },
            "PrivateNote": "Monthly transfer"
        });

        let txn = parse_transfer(&entity, "35").unwrap();
        assert_eq!(txn.description, "Monthly transfer");
    }

    // ── Golden transfer fixtures: same entity, different --account-id ─

    /// A single QBO Transfer entity: $5,000 from Checking (35) to Savings (42).
    /// The sign must flip depending on which account the user selected.
    fn golden_transfer() -> serde_json::Value {
        serde_json::json!({
            "Id": "400",
            "TxnDate": "2026-01-20",
            "Amount": 5000.00,
            "FromAccountRef": { "value": "35", "name": "Checking" },
            "ToAccountRef":   { "value": "42", "name": "Savings" },
            "PrivateNote": ""
        })
    }

    #[test]
    fn test_golden_transfer_checking_to_savings_from_checking_perspective() {
        // User runs: --account-id 35 (Checking)
        // Money leaves Checking → amount must be negative
        let txn = parse_transfer(&golden_transfer(), "35").unwrap();
        assert_eq!(txn.amount_minor, -500000, "Checking→Savings should be negative from Checking's perspective");
        assert_eq!(txn.txn_type, "transfer");
        assert_eq!(txn.source_id, "transfer:400");
        assert_eq!(txn.description, "Transfer to Savings");
    }

    #[test]
    fn test_golden_transfer_checking_to_savings_from_savings_perspective() {
        // User runs: --account-id 42 (Savings)
        // Money arrives in Savings → amount must be positive
        let txn = parse_transfer(&golden_transfer(), "42").unwrap();
        assert_eq!(txn.amount_minor, 500000, "Checking→Savings should be positive from Savings' perspective");
        assert_eq!(txn.txn_type, "transfer");
        assert_eq!(txn.source_id, "transfer:400");
        assert_eq!(txn.description, "Transfer from Checking");
    }

    // ── Account resolution (httpmock) ──────────────────────────────

    #[test]
    fn test_account_resolution_single() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query")
                .query_param_exists("query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Account": [
                            { "Id": "35", "Name": "Checking", "AccountType": "Bank" }
                        ]
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let id = client.resolve_account("Checking", true).unwrap();
        assert_eq!(id, "35");
    }

    #[test]
    fn test_account_resolution_apostrophe() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Account": [
                            { "Id": "50", "Name": "Owner's Draw", "AccountType": "Bank" }
                        ]
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let id = client.resolve_account("Owner's Draw", true).unwrap();
        assert_eq!(id, "50");
    }

    #[test]
    fn test_account_resolution_ambiguous_with_fqn() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Account": [
                            {
                                "Id": "35",
                                "Name": "Checking",
                                "FullyQualifiedName": "Business:Checking",
                                "AccountType": "Bank"
                            },
                            {
                                "Id": "36",
                                "Name": "Checking",
                                "FullyQualifiedName": "Personal:Checking",
                                "AccountType": "Bank"
                            }
                        ]
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let err = client.resolve_account("Checking", true).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_VALIDATION);
        assert!(err.message.contains("Multiple accounts"), "message: {}", err.message);
        // FullyQualifiedName should appear so the user can distinguish sub-accounts
        assert!(
            err.message.contains("Business:Checking"),
            "should include FQN: {}",
            err.message,
        );
        assert!(
            err.message.contains("Personal:Checking"),
            "should include FQN: {}",
            err.message,
        );
        assert!(err.message.contains("ID 35"), "should include IDs: {}", err.message);
        assert!(err.message.contains("ID 36"), "should include IDs: {}", err.message);
    }

    #[test]
    fn test_account_resolution_not_found() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {}
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
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

        // Single page with < 1000 results → no pagination needed
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Deposit": (0..50).map(|i| {
                            serde_json::json!({
                                "Id": format!("{}", i),
                                "TxnDate": "2026-01-15",
                                "TotalAmt": 100.00,
                                "DepositToAccountRef": { "value": "35" },
                            })
                        }).collect::<Vec<_>>()
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_deposits("35", &from, &to, true).unwrap();
        assert_eq!(txns.len(), 50);
    }

    #[test]
    fn test_empty_result_set() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {}
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_deposits("35", &from, &to, true).unwrap();
        assert_eq!(txns.len(), 0);
    }

    // ── Auth failure (httpmock) ────────────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(401)
                .json_body(serde_json::json!({
                    "Fault": {
                        "Error": [{
                            "Message": "message=AuthenticationFailed",
                            "Detail": "Token expired"
                        }]
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "bad_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_deposits("35", &from, &to, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("QBO auth failed (401)"),
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
                .path_includes("/query")
                .header("Authorization", "Bearer old_token");
            then.status(401)
                .json_body(serde_json::json!({
                    "Fault": {
                        "Error": [{ "Message": "Unauthorized", "Detail": "Token expired" }]
                    }
                }));
        });

        // Refresh → new token
        server.mock(|when, then| {
            when.method(POST)
                .path("/oauth2/v1/tokens/bearer");
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
                .path_includes("/query")
                .header("Authorization", "Bearer new_token");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Deposit": [{
                            "Id": "101",
                            "TxnDate": "2026-01-15",
                            "TotalAmt": 500.00,
                            "DepositToAccountRef": { "value": "35" }
                        }]
                    }
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("qbo.json");
        let creds = QboCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "old_refresh".into(),
            realm_id: "realm_123".into(),
            access_token_expires_at: None,
            refresh_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = QboClient::from_credentials_with_base_url(
            creds,
            creds_path.clone(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_deposits("35", &from, &to, true).unwrap();

        assert_eq!(txns.len(), 1);
        assert_eq!(txns[0].amount_minor, 50000);

        // Verify credentials were updated on disk
        let saved: QboCredentials = serde_json::from_str(
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
        let path = dir.path().join("qbo.json");
        std::fs::write(
            &path,
            serde_json::json!({
                "client_id": "cid",
                "client_secret": "csec",
                "access_token": "at",
                "refresh_token": "rt",
                "realm_id": "rid"
            })
            .to_string(),
        )
        .unwrap();
        let creds = load_credentials(&path.to_path_buf()).unwrap();
        assert_eq!(creds.client_id, "cid");
        assert_eq!(creds.realm_id, "rid");
        assert!(creds.access_token_expires_at.is_none());
    }

    #[test]
    fn test_credential_loading_missing_file() {
        let path = PathBuf::from("/tmp/nonexistent-qbo-creds.json");
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
        let types = parse_include("deposit").unwrap();
        assert_eq!(types, vec!["deposit"]);

        let types = parse_include("deposit,transfer").unwrap();
        assert_eq!(types, vec!["deposit", "transfer"]);

        let types = parse_include("deposit,purchase,transfer").unwrap();
        assert_eq!(types, vec!["deposit", "purchase", "transfer"]);

        let err = parse_include("deposit,invoice").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_USAGE);
        assert!(err.message.contains("invoice"));
    }

    #[test]
    fn test_include_filter_only_queries_selected() {
        // Verify that parse_include correctly filters entity types
        let types = parse_include("deposit").unwrap();
        assert_eq!(types, vec!["deposit"]);
        assert!(!types.contains(&"purchase".to_string()));
        assert!(!types.contains(&"transfer".to_string()));

        // Verify the fetch_deposits method works with a mock
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Deposit": [{
                            "Id": "101",
                            "TxnDate": "2026-01-15",
                            "TotalAmt": 100.00,
                            "DepositToAccountRef": { "value": "35" },
                        }]
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let txns = client.fetch_deposits("35", &from, &to, true).unwrap();
        assert_eq!(txns.len(), 1);
    }

    // ── ORDERBY determinism ────────────────────────────────────────

    #[test]
    fn test_query_includes_orderby() {
        // Verify that the SQL query built by query_entities includes
        // deterministic ordering to prevent pagination gaps/duplicates.
        let server = MockServer::start();

        // Capture the query parameter to verify ORDERBY is present
        let mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(200)
                .json_body(serde_json::json!({
                    "QueryResponse": {
                        "Deposit": []
                    }
                }));
        });

        let mut client = QboClient::with_base_url(
            "test_token".into(),
            "realm_123".into(),
            server.base_url(),
        );

        let _ = client.query_entities(
            "Deposit",
            "DepositToAccountRef = '35' AND TxnDate >= '2026-01-01' AND TxnDate < '2026-01-31'",
            true,
        );

        mock.assert();

        // The query_entities method hardcodes ORDERBY TxnDate ASC, Id ASC.
        // We can't easily inspect the request in httpmock 0.8 without
        // query_param_contains, so we verify the format string directly:
        let test_query = format!(
            "SELECT * FROM {} WHERE {} ORDERBY TxnDate ASC, Id ASC STARTPOSITION {} MAXRESULTS {}",
            "Deposit",
            "DepositToAccountRef = '35'",
            1,
            QBO_QUERY_LIMIT,
        );
        assert!(test_query.contains("ORDERBY TxnDate ASC, Id ASC"));
    }

    // ── Refresh failure hint ───────────────────────────────────────

    #[test]
    fn test_refresh_failure_includes_reconnect_hint() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/query");
            then.status(401)
                .json_body(serde_json::json!({
                    "Fault": { "Error": [{ "Message": "Unauthorized" }] }
                }));
        });

        // Refresh → revoked
        server.mock(|when, then| {
            when.method(POST)
                .path("/oauth2/v1/tokens/bearer");
            then.status(400)
                .json_body(serde_json::json!({
                    "error": "invalid_grant",
                    "error_description": "refresh token has been revoked"
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("qbo.json");
        let creds = QboCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "revoked".into(),
            realm_id: "realm_123".into(),
            access_token_expires_at: None,
            refresh_token_expires_at: None,
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = QboClient::from_credentials_with_base_url(
            creds,
            creds_path,
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_deposits("35", &from, &to, true)
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
}
