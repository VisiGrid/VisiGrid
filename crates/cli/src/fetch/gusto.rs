//! `vgrid fetch gusto` — fetch Gusto payroll data into canonical CSV.

use std::path::PathBuf;

use chrono::NaiveDate;
use serde::{Deserialize, Serialize};

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const GUSTO_API_BASE: &str = "https://api.gusto.com";
const GUSTO_API_VERSION: &str = "2024-04-01";
const PAGE_LIMIT: u32 = 100;

// ── Credentials ─────────────────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct GustoCredentials {
    pub client_id: String,
    pub client_secret: String,
    pub access_token: String,
    pub refresh_token: String,
    pub company_uuid: String,
}

fn load_credentials(path: &PathBuf) -> Result<GustoCredentials, CliError> {
    let content = std::fs::read_to_string(path).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_NOT_AUTH,
        message: format!(
            "cannot read credentials file {}: {}",
            path.display(),
            e,
        ),
        hint: None,
    })?;

    // Warn if file is world-readable (Unix only)
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

fn save_credentials(creds: &GustoCredentials, path: &PathBuf) -> Result<(), CliError> {
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
    creds: &GustoCredentials,
    http: &reqwest::blocking::Client,
    base_url: &str,
) -> Result<GustoCredentials, CliError> {
    let url = format!("{}/oauth/token", base_url);
    let resp = http
        .post(&url)
        .form(&[
            ("grant_type", "refresh_token"),
            ("client_id", creds.client_id.as_str()),
            ("client_secret", creds.client_secret.as_str()),
            ("refresh_token", creds.refresh_token.as_str()),
        ])
        .send()
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: format!("Gusto token refresh request failed: {}", e),
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
            message: format!("Gusto token refresh failed ({}): {}", status, msg),
            hint: None,
        });
    }

    let body: serde_json::Value = resp.json().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_AUTH,
        message: format!("Gusto token refresh response invalid: {}", e),
        hint: None,
    })?;

    let new_access = body["access_token"]
        .as_str()
        .ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_AUTH,
            message: "Gusto token refresh response missing access_token".into(),
            hint: None,
        })?;

    let new_refresh = body["refresh_token"]
        .as_str()
        .unwrap_or(&creds.refresh_token);

    Ok(GustoCredentials {
        client_id: creds.client_id.clone(),
        client_secret: creds.client_secret.clone(),
        access_token: new_access.to_string(),
        refresh_token: new_refresh.to_string(),
        company_uuid: creds.company_uuid.clone(),
    })
}

// ── Amount parsing (string-to-cents, no f64) ────────────────────────

fn parse_money_string(s: &str) -> Result<i64, String> {
    let s = s.trim();
    let negative = s.starts_with('-');
    let s = s.trim_start_matches('-');
    let (dollars, cents) = if let Some(dot) = s.find('.') {
        let d: i64 = s[..dot]
            .parse()
            .map_err(|e| format!("bad dollars: {}", e))?;
        let frac = &s[dot + 1..];
        let c: i64 = match frac.len() {
            0 => 0,
            1 => {
                frac.parse::<i64>()
                    .map_err(|e| format!("bad cents: {}", e))?
                    * 10
            }
            2 => frac.parse().map_err(|e| format!("bad cents: {}", e))?,
            _ => return Err(format!("too many decimal places: {}", s)),
        };
        (d, c)
    } else {
        (s.parse().map_err(|e| format!("bad amount: {}", e))?, 0)
    };
    let minor = dollars * 100 + cents;
    Ok(if negative { -minor } else { minor })
}

// ── Internal payroll row ────────────────────────────────────────────

#[derive(Debug)]
pub(super) struct RawPayrollRow {
    effective_date: String,
    amount_minor: i64,
    row_type: String,
    source_id: String,
    group_id: String,
    description: String,
}

/// Parse a single Gusto payroll JSON into 3 canonical rows (net, tax, other).
fn parse_payroll(item: &serde_json::Value) -> Result<Vec<RawPayrollRow>, CliError> {
    let uuid = item["uuid"].as_str().unwrap_or("").to_string();
    let check_date = item["check_date"].as_str().unwrap_or("").to_string();

    let pay_period = &item["pay_period"];
    let start = pay_period["start_date"].as_str().unwrap_or("");
    let end = pay_period["end_date"].as_str().unwrap_or("");

    let totals = &item["totals"];

    let company_debit_str = totals["company_debit"]
        .as_str()
        .unwrap_or("0");
    let net_pay_str = totals["net_pay_debit"]
        .as_str()
        .unwrap_or("0");
    let tax_str = totals["tax_debit"]
        .as_str()
        .unwrap_or("0");

    let company_debit = parse_money_string(company_debit_str).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!(
            "payroll {} bad company_debit {:?}: {}",
            uuid, company_debit_str, e,
        ),
        hint: None,
    })?;
    let net_pay = parse_money_string(net_pay_str).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!(
            "payroll {} bad net_pay_debit {:?}: {}",
            uuid, net_pay_str, e,
        ),
        hint: None,
    })?;
    let tax = parse_money_string(tax_str).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!(
            "payroll {} bad tax_debit {:?}: {}",
            uuid, tax_str, e,
        ),
        hint: None,
    })?;

    let mut other = company_debit - net_pay - tax;
    if other < 0 {
        eprintln!(
            "warning: payroll {} derived other amount is negative ({}), clamping to 0",
            uuid, other,
        );
        other = 0;
    }

    let group_id = format!("payroll:{}", uuid);

    Ok(vec![
        RawPayrollRow {
            effective_date: check_date.clone(),
            amount_minor: -net_pay,
            row_type: "payroll_net".to_string(),
            source_id: format!("payroll:{}:net", uuid),
            group_id: group_id.clone(),
            description: format!("Net Pay {} - {}", start, end),
        },
        RawPayrollRow {
            effective_date: check_date.clone(),
            amount_minor: -tax,
            row_type: "payroll_tax".to_string(),
            source_id: format!("payroll:{}:tax", uuid),
            group_id: group_id.clone(),
            description: format!("Taxes {} - {}", start, end),
        },
        RawPayrollRow {
            effective_date: check_date,
            amount_minor: -other,
            row_type: "payroll_other".to_string(),
            source_id: format!("payroll:{}:other", uuid),
            group_id,
            description: format!("Fees/Other {} - {}", start, end),
        },
    ])
}

fn extract_gusto_error(body: &serde_json::Value, status: u16) -> String {
    body["error_message"]
        .as_str()
        .or_else(|| body["message"].as_str())
        .or_else(|| body["error"].as_str())
        .unwrap_or(&format!("HTTP {}", status))
        .to_string()
}

// ── Gusto client ────────────────────────────────────────────────────

pub struct GustoClient {
    client: FetchClient,
    access_token: String,
    company_uuid: String,
    base_url: String,
    creds: Option<GustoCredentials>,
    creds_path: Option<PathBuf>,
}

impl GustoClient {
    pub fn new(access_token: String, company_uuid: String) -> Self {
        Self::with_base_url(access_token, company_uuid, GUSTO_API_BASE.to_string())
    }

    pub fn with_base_url(
        access_token: String,
        company_uuid: String,
        base_url: String,
    ) -> Self {
        Self {
            client: FetchClient::new("Gusto", extract_gusto_error),
            access_token,
            company_uuid,
            base_url,
            creds: None,
            creds_path: None,
        }
    }

    fn from_credentials(
        creds: GustoCredentials,
        path: PathBuf,
        base_url: String,
    ) -> Self {
        Self {
            client: FetchClient::new("Gusto", extract_gusto_error),
            access_token: creds.access_token.clone(),
            company_uuid: creds.company_uuid.clone(),
            base_url,
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

        let new_creds =
            refresh_access_token(creds, &self.client.http, &self.base_url)?;
        save_credentials(&new_creds, path)?;
        self.access_token = new_creds.access_token.clone();
        self.creds = Some(new_creds);
        Ok(())
    }

    /// Fetch all processed payrolls in the given date range.
    pub(super) fn fetch_payrolls(
        &mut self,
        from: &NaiveDate,
        to: &NaiveDate,
        quiet: bool,
    ) -> Result<Vec<RawPayrollRow>, CliError> {
        let mut all_rows = Vec::new();
        let mut page = 1u32;
        let mut refreshed = false;
        let stderr_tty = atty::is(atty::Stream::Stderr);
        let show_progress = !quiet && stderr_tty;

        loop {
            let url = format!(
                "{}/v1/companies/{}/payrolls",
                self.base_url, self.company_uuid,
            );
            let params = vec![
                ("processing_statuses", "processed".to_string()),
                ("start_date", from.to_string()),
                ("end_date", to.to_string()),
                ("page", page.to_string()),
                ("per", PAGE_LIMIT.to_string()),
            ];
            let token = self.access_token.clone();

            let result = self.client.request_with_retry(|http| {
                http.get(&url)
                    .bearer_auth(&token)
                    .header("X-Gusto-API-Version", GUSTO_API_VERSION)
                    .query(&params)
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
                    // Retry same page with refreshed token
                    let token = self.access_token.clone();
                    self.client.request_with_retry(|http| {
                        http.get(&url)
                            .bearer_auth(&token)
                            .header("X-Gusto-API-Version", GUSTO_API_VERSION)
                            .query(&params)
                    })?
                }
                Err(e) => return Err(e),
            };

            let payrolls = body.as_array().ok_or_else(|| CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message: "Gusto response is not an array".into(),
                hint: None,
            })?;

            let count = payrolls.len() as u32;

            if show_progress {
                eprintln!("  page {}: {} payrolls", page, count);
            }

            for item in payrolls {
                all_rows.extend(parse_payroll(item)?);
            }

            if count < PAGE_LIMIT {
                break;
            }

            page += 1;
        }

        Ok(all_rows)
    }
}

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_gusto(
    from: String,
    to: String,
    credentials: Option<PathBuf>,
    access_token: Option<String>,
    company_uuid: Option<String>,
    out: Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Parse and validate dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // Gusto limits date range to max 1 year
    let days = (to_date - from_date).num_days();
    if days > 366 {
        return Err(CliError::args(format!(
            "Gusto limits date range to 1 year ({} days requested)",
            days,
        )));
    }

    // 2. Resolve auth
    let mut client = if let Some(token) = access_token {
        let token = token.trim().to_string();
        if token.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing Gusto access token (--access-token is empty)"
                    .into(),
                hint: None,
            });
        }
        let uuid = company_uuid.ok_or_else(|| CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message: "missing --company-uuid (required with --access-token)"
                .into(),
            hint: None,
        })?;
        let uuid = uuid.trim().to_string();
        if uuid.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: "missing Gusto company UUID (--company-uuid is empty)"
                    .into(),
                hint: None,
            });
        }
        GustoClient::new(token, uuid)
    } else if let Some(ref path) = credentials {
        let expanded =
            shellexpand::tilde(&path.to_string_lossy()).to_string();
        let creds_path = PathBuf::from(&expanded);
        let creds = load_credentials(&creds_path)?;
        GustoClient::from_credentials(
            creds,
            creds_path,
            GUSTO_API_BASE.to_string(),
        )
    } else {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_NOT_AUTH,
            message:
                "missing Gusto credentials (use --credentials or --access-token)"
                    .into(),
            hint: None,
        });
    };

    // 3. Fetch
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !quiet && stderr_tty;

    if show_progress {
        eprintln!(
            "Fetching Gusto payrolls ({} to {})...",
            from_date, to_date,
        );
    }

    let mut rows_raw =
        client.fetch_payrolls(&from_date, &to_date, quiet)?;

    // 4. Sort: (effective_date ASC, source_id ASC)
    rows_raw.sort_by(|a, b| {
        a.effective_date
            .cmp(&b.effective_date)
            .then_with(|| a.source_id.cmp(&b.source_id))
    });

    // 5. Build canonical rows
    let rows: Vec<CanonicalRow> = rows_raw
        .iter()
        .map(|r| CanonicalRow {
            effective_date: r.effective_date.clone(),
            posted_date: r.effective_date.clone(), // Gusto doesn't distinguish
            amount_minor: r.amount_minor,
            currency: "USD".to_string(),
            r#type: r.row_type.clone(),
            source: "gusto".to_string(),
            source_id: r.source_id.clone(),
            group_id: r.group_id.clone(),
            description: r.description.clone(),
        })
        .collect();

    // 6. Write CSV
    let out_label = common::write_csv(&rows, &out)?;

    if show_progress {
        eprintln!("Done: {} rows written to {}", rows.len(), out_label);
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use httpmock::prelude::*;

    // ── parse_money_string ─────────────────────────────────────────

    #[test]
    fn test_parse_money_string_basic() {
        assert_eq!(parse_money_string("1080.47").unwrap(), 108047);
        assert_eq!(parse_money_string("0.01").unwrap(), 1);
        assert_eq!(parse_money_string("100").unwrap(), 10000);
        assert_eq!(parse_money_string("0").unwrap(), 0);
        assert_eq!(parse_money_string("0.00").unwrap(), 0);
    }

    #[test]
    fn test_parse_money_string_negative() {
        assert_eq!(parse_money_string("-500.25").unwrap(), -50025);
        assert_eq!(parse_money_string("-0.01").unwrap(), -1);
    }

    #[test]
    fn test_parse_money_string_one_decimal() {
        assert_eq!(parse_money_string("10.5").unwrap(), 1050);
        assert_eq!(parse_money_string("0.5").unwrap(), 50);
    }

    #[test]
    fn test_parse_money_string_no_decimal() {
        assert_eq!(parse_money_string("42").unwrap(), 4200);
    }

    #[test]
    fn test_parse_money_string_trailing_dot() {
        assert_eq!(parse_money_string("100.").unwrap(), 10000);
    }

    #[test]
    fn test_parse_money_string_whitespace() {
        assert_eq!(parse_money_string("  1080.47  ").unwrap(), 108047);
    }

    #[test]
    fn test_parse_money_string_too_many_decimals() {
        assert!(parse_money_string("10.123").is_err());
    }

    #[test]
    fn test_parse_money_string_bad_input() {
        assert!(parse_money_string("abc").is_err());
    }

    // ── parse_payroll → 3 rows ─────────────────────────────────────

    fn mock_payroll_json(
        uuid: &str,
        check_date: &str,
        company_debit: &str,
        net_pay: &str,
        tax: &str,
    ) -> serde_json::Value {
        serde_json::json!({
            "uuid": uuid,
            "check_date": check_date,
            "pay_period": {
                "start_date": "2026-01-01",
                "end_date": "2026-01-15"
            },
            "totals": {
                "company_debit": company_debit,
                "net_pay_debit": net_pay,
                "tax_debit": tax
            }
        })
    }

    #[test]
    fn test_parse_payroll_three_rows() {
        let item = mock_payroll_json("pr_001", "2026-01-15", "1819.47", "1520.00", "284.00");
        let rows = parse_payroll(&item).unwrap();

        assert_eq!(rows.len(), 3);

        // Net Pay
        assert_eq!(rows[0].row_type, "payroll_net");
        assert_eq!(rows[0].amount_minor, -152000);
        assert_eq!(rows[0].source_id, "payroll:pr_001:net");
        assert_eq!(rows[0].group_id, "payroll:pr_001");
        assert_eq!(rows[0].effective_date, "2026-01-15");

        // Tax
        assert_eq!(rows[1].row_type, "payroll_tax");
        assert_eq!(rows[1].amount_minor, -28400);
        assert_eq!(rows[1].source_id, "payroll:pr_001:tax");

        // Other (1819.47 - 1520.00 - 284.00 = 15.47 → 1547 cents)
        assert_eq!(rows[2].row_type, "payroll_other");
        assert_eq!(rows[2].amount_minor, -1547);
        assert_eq!(rows[2].source_id, "payroll:pr_001:other");
    }

    #[test]
    fn test_parse_payroll_zero_other() {
        // company_debit exactly equals net + tax
        let item = mock_payroll_json("pr_002", "2026-01-31", "1804.00", "1520.00", "284.00");
        let rows = parse_payroll(&item).unwrap();

        assert_eq!(rows[2].row_type, "payroll_other");
        assert_eq!(rows[2].amount_minor, 0); // -0 = 0
    }

    // ── Credential loading ─────────────────────────────────────────

    #[test]
    fn test_load_credentials_missing_file() {
        let path = PathBuf::from("/tmp/nonexistent-gusto-creds.json");
        let err = load_credentials(&path).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("cannot read credentials file"));
    }

    #[test]
    fn test_load_credentials_invalid_json() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("bad.json");
        std::fs::write(&path, "not json").unwrap();
        let err = load_credentials(&path.to_path_buf()).unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("invalid credentials JSON"));
    }

    #[test]
    fn test_load_credentials_valid() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("gusto.json");
        std::fs::write(
            &path,
            serde_json::json!({
                "client_id": "cid",
                "client_secret": "csec",
                "access_token": "at",
                "refresh_token": "rt",
                "company_uuid": "cu"
            })
            .to_string(),
        )
        .unwrap();
        let creds = load_credentials(&path.to_path_buf()).unwrap();
        assert_eq!(creds.client_id, "cid");
        assert_eq!(creds.company_uuid, "cu");
    }

    // ── Mock pagination ────────────────────────────────────────────

    fn mock_gusto_payroll(uuid: &str) -> serde_json::Value {
        serde_json::json!({
            "uuid": uuid,
            "check_date": "2026-01-15",
            "pay_period": {
                "start_date": "2026-01-01",
                "end_date": "2026-01-15"
            },
            "totals": {
                "company_debit": "1000.00",
                "net_pay_debit": "800.00",
                "tax_debit": "150.00"
            }
        })
    }

    #[test]
    fn test_pagination_two_pages() {
        let server = MockServer::start();

        // Build 100 payrolls for page 1
        let page1: Vec<serde_json::Value> = (0..100)
            .map(|i| mock_gusto_payroll(&format!("pr_{:04}", i)))
            .collect();

        let page1_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls")
                .query_param("page", "1");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!(page1));
        });

        // Build 3 payrolls for page 2
        let page2: Vec<serde_json::Value> = (100..103)
            .map(|i| mock_gusto_payroll(&format!("pr_{:04}", i)))
            .collect();

        let page2_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls")
                .query_param("page", "2");
            then.status(200)
                .header("content-type", "application/json")
                .json_body(serde_json::json!(page2));
        });

        let mut client = GustoClient::with_base_url(
            "test_token".into(),
            "company_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let rows = client.fetch_payrolls(&from, &to, true).unwrap();

        page1_mock.assert();
        page2_mock.assert();
        // 103 payrolls × 3 rows each = 309 rows
        assert_eq!(rows.len(), 309);
    }

    // ── Auth failure ───────────────────────────────────────────────

    #[test]
    fn test_auth_failure_exit_51() {
        let server = MockServer::start();

        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Unauthorized"
                }));
        });

        let mut client = GustoClient::with_base_url(
            "bad_token".into(),
            "company_123".into(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_payrolls(&from, &to, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("Gusto auth failed (401)"),
            "message: {}",
            err.message,
        );
    }

    // ── Token refresh ──────────────────────────────────────────────

    #[test]
    fn test_token_refresh_on_401() {
        let server = MockServer::start();

        // First request → 401
        let first_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls")
                .header("Authorization", "Bearer old_token");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Unauthorized"
                }));
        });

        // Refresh → new token
        let refresh_mock = server.mock(|when, then| {
            when.method(POST)
                .path("/oauth/token");
            then.status(200)
                .json_body(serde_json::json!({
                    "access_token": "new_token",
                    "refresh_token": "new_refresh"
                }));
        });

        // Retry with new token → success
        let retry_mock = server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls")
                .header("Authorization", "Bearer new_token");
            then.status(200)
                .json_body(serde_json::json!([
                    {
                        "uuid": "pr_001",
                        "check_date": "2026-01-15",
                        "pay_period": {
                            "start_date": "2026-01-01",
                            "end_date": "2026-01-15"
                        },
                        "totals": {
                            "company_debit": "1000.00",
                            "net_pay_debit": "800.00",
                            "tax_debit": "150.00"
                        }
                    }
                ]));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("gusto.json");
        let creds = GustoCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "old_refresh".into(),
            company_uuid: "company_123".into(),
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = GustoClient::from_credentials(
            creds,
            creds_path.clone(),
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let rows = client.fetch_payrolls(&from, &to, true).unwrap();

        first_mock.assert();
        refresh_mock.assert();
        retry_mock.assert();
        assert_eq!(rows.len(), 3); // 1 payroll × 3 rows

        // Verify credentials were updated on disk
        let saved: GustoCredentials = serde_json::from_str(
            &std::fs::read_to_string(&creds_path).unwrap(),
        )
        .unwrap();
        assert_eq!(saved.access_token, "new_token");
        assert_eq!(saved.refresh_token, "new_refresh");
    }

    #[test]
    fn test_token_refresh_failure_exits_51() {
        let server = MockServer::start();

        // First request → 401
        server.mock(|when, then| {
            when.method(GET)
                .path_includes("/payrolls");
            then.status(401)
                .json_body(serde_json::json!({
                    "message": "Unauthorized"
                }));
        });

        // Refresh → failure
        server.mock(|when, then| {
            when.method(POST)
                .path("/oauth/token");
            then.status(400)
                .json_body(serde_json::json!({
                    "error": "invalid_grant",
                    "error_description": "refresh token has been revoked"
                }));
        });

        let dir = tempfile::tempdir().unwrap();
        let creds_path = dir.path().join("gusto.json");
        let creds = GustoCredentials {
            client_id: "cid".into(),
            client_secret: "csec".into(),
            access_token: "old_token".into(),
            refresh_token: "revoked_refresh".into(),
            company_uuid: "company_123".into(),
        };
        std::fs::write(
            &creds_path,
            serde_json::to_string(&creds).unwrap(),
        )
        .unwrap();

        let mut client = GustoClient::from_credentials(
            creds,
            creds_path,
            server.base_url(),
        );

        let from = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
        let to = NaiveDate::from_ymd_opt(2026, 1, 31).unwrap();
        let err = client
            .fetch_payrolls(&from, &to, true)
            .unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_AUTH);
        assert!(
            err.message.contains("token refresh failed"),
            "message: {}",
            err.message,
        );
    }
}
