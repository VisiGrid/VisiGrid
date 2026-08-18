//! Shared infrastructure for `vgrid fetch` adapters.
//!
//! Each adapter (stripe, mercury, http, …) reuses:
//! - `FetchClient` — HTTP client with retry / backoff / error classification
//! - `CanonicalRow` — the 10-column CSV schema all adapters emit
//! - `resolve_api_key` — flag > env > error
//! - `parse_date_range` — parse + validate `--from` / `--to`
//! - `write_csv` — open output, write header + rows, flush
//!
//! # CanonicalRow Contract
//!
//! Every fetch adapter MUST produce rows conforming to this contract.
//! Downstream reconciliation, fingerprinting, and diffing depend on it.
//! Breaking changes here break production fingerprints silently.
//!
//! ## Columns (in order)
//!
//! | #  | Column           | Type     | Required | Description                          |
//! |----|------------------|----------|----------|--------------------------------------|
//! | 1  | `effective_date` | `String` | Yes      | When the transaction occurred         |
//! | 2  | `posted_date`    | `String` | No       | When the transaction settled/posted   |
//! | 3  | `amount_minor`   | `i64`    | Yes      | Amount in minor units (cents). Never float. |
//! | 4  | `currency`       | `String` | Yes      | ISO 4217 uppercase (USD, EUR, GBP)   |
//! | 5  | `type`           | `String` | Yes      | Transaction type (charge, refund, …)  |
//! | 6  | `source`         | `String` | Yes      | Adapter name (stripe, mercury, …)     |
//! | 7  | `source_id`      | `String` | Yes      | Upstream id of the SUBJECT (charge, payment). NOT unique — several transactions can share one |
//! | 8  | `group_id`       | `String` | No       | Grouping key (payout ID, invoice, …)  |
//! | 9  | `description`    | `String` | No       | Human-readable memo                   |
//! | 10 | `btxn_id`        | `String` | No       | Provider's per-transaction id, when it has one |
//!
//! ## Invariants
//!
//! - **Column order**: Fixed. Serialized by `serde` in struct field order.
//!   CSV header is always `effective_date,posted_date,amount_minor,…`.
//! - **Sort order**: Deterministic. Default: `group_id`, `effective_date`,
//!   `source_id`. Adapters may override via `sort_by`. Ties are stable
//!   (Rust's `sort_by` is stable). Two runs over the same data MUST
//!   produce byte-identical CSV.
//! - **Dates**: ISO 8601 date strings (`YYYY-MM-DD`). Timezone handling
//!   is the adapter's responsibility — convert to UTC or local date before
//!   writing. Empty string for missing optional dates.
//! - **Amounts**: Always `i64` minor units. Use `parse_money_string()` for
//!   decimal-to-cents conversion (integer math, no floats, max 2 decimal
//!   places). Negative values for refunds/credits.
//! - **Optional columns**: Empty string `""` when absent. Never `null`,
//!   never omitted. CSV always has 10 columns per row.
//! - **Identity**: `btxn_id` where non-empty, else `source_id`. `source_id`
//!   alone is NOT an identity — Stripe posts a charge, its ACH return and a
//!   later dispute against the same source, and a consumer keying on it drops
//!   all but the first while reporting success.
//! - **Encoding**: UTF-8. The `csv` crate handles quoting/escaping.

use std::io::Write;
use std::path::PathBuf;
use std::thread;
use std::time::Duration;

use chrono::NaiveDate;
use serde::Serialize;

use crate::exit_codes;
use crate::CliError;

// ── Constants ───────────────────────────────────────────────────────

pub(super) const MAX_RETRIES: u32 = 3;
pub(super) const USER_AGENT: &str = concat!("vgrid/", env!("CARGO_PKG_VERSION"));

// ── Canonical CSV row ───────────────────────────────────────────────

#[derive(Debug, Serialize)]
pub(crate) struct CanonicalRow {
    pub effective_date: String,
    pub posted_date: String,
    pub amount_minor: i64,
    pub currency: String,
    pub r#type: String,
    pub source: String,
    pub source_id: String,
    pub group_id: String,
    pub description: String,
    /// Provider's own id for this transaction, where the provider has one.
    ///
    /// `source_id` names the *thing the transaction is about* — a charge, a
    /// payment — and is NOT unique: Stripe posts a charge, then its ACH
    /// return, then a dispute, all against the same source. Keying on it made
    /// the importer treat the second row as already-seen and drop it,
    /// reporting success. Fees escaped only because they were special-cased
    /// into `{source_id}_fee`.
    ///
    /// Empty for rows the provider never assigned an id to — synthetic fee
    /// rows, and adapters whose API exposes no per-transaction identifier.
    /// Consumers fall back to `source_id` when this is blank.
    ///
    /// APPEND-ONLY. New fields go after this one, never between existing
    /// ones. Serde derives the CSV header from this order, and the empty-rows
    /// header below is written by hand — they must agree.
    pub btxn_id: String,
}

// ── FetchClient ─────────────────────────────────────────────────────

/// Shared HTTP client that handles retry, backoff, and error classification.
///
/// Adapters own their API key, base URL, and auth method. They pass a
/// request-building closure to [`request_with_retry`] which handles
/// the retry loop and maps HTTP status codes to the standard exit codes.
pub(super) struct FetchClient {
    pub(super) http: reqwest::blocking::Client,
    source_name: String,
    error_extractor: fn(&serde_json::Value, u16) -> String,
}

impl FetchClient {
    pub(super) fn new(
        source_name: &str,
        error_extractor: fn(&serde_json::Value, u16) -> String,
    ) -> Self {
        let http = reqwest::blocking::Client::builder()
            .timeout(Duration::from_secs(30))
            .user_agent(USER_AGENT)
            .build()
            .expect("failed to build HTTP client");

        Self {
            http,
            source_name: source_name.to_string(),
            error_extractor,
        }
    }

    /// Make a GET request with retry + exponential backoff.
    ///
    /// `build_request` is called once per attempt. It receives the
    /// underlying `reqwest::blocking::Client` and must return a fully
    /// configured `RequestBuilder` (URL, auth, headers, query params).
    pub(super) fn request_with_retry(
        &self,
        build_request: impl Fn(&reqwest::blocking::Client) -> reqwest::blocking::RequestBuilder,
    ) -> Result<serde_json::Value, CliError> {
        let mut backoff_secs = 1u64;

        for attempt in 0..=MAX_RETRIES {
            let req = build_request(&self.http);
            let result = req.send();

            match result {
                Ok(resp) => {
                    let status = resp.status().as_u16();

                    // Auth errors: fail immediately
                    if status == 401 || status == 403 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_AUTH,
                            message: format!(
                                "{} auth failed ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    // Bad request: fail immediately
                    if status == 400 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_VALIDATION,
                            message: format!(
                                "{} request rejected ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    // Other 4xx (not 429): fail immediately
                    if (400..500).contains(&status) && status != 429 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "{} error ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    // Retryable: 429, 5xx
                    if status == 429 || status >= 500 {
                        if attempt == MAX_RETRIES {
                            let exit_code = if status == 429 {
                                exit_codes::EXIT_FETCH_RATE_LIMIT
                            } else {
                                exit_codes::EXIT_FETCH_UPSTREAM
                            };
                            return Err(CliError {
                                code: exit_code,
                                message: format!(
                                    "{} {} after {} attempts ({})",
                                    self.source_name,
                                    if status == 429 {
                                        "rate limited"
                                    } else {
                                        "upstream error"
                                    },
                                    MAX_RETRIES,
                                    status,
                                ),
                                hint: None,
                            });
                        }

                        // Respect Retry-After header for 429
                        let wait = if status == 429 {
                            resp.headers()
                                .get("retry-after")
                                .and_then(|v| v.to_str().ok())
                                .and_then(|v| v.parse::<u64>().ok())
                                .unwrap_or(backoff_secs)
                        } else {
                            backoff_secs
                        };

                        eprintln!(
                            "warning: retry {}/{} in {}s (HTTP {})",
                            attempt + 1,
                            MAX_RETRIES,
                            wait,
                            status,
                        );
                        thread::sleep(Duration::from_secs(wait));
                        backoff_secs *= 2;
                        continue;
                    }

                    // Success: parse JSON (read as text first to handle
                    // BOM-prefixed responses from providers like Authorize.net)
                    let text = resp.text().map_err(|e| CliError {
                        code: exit_codes::EXIT_FETCH_UPSTREAM,
                        message: format!(
                            "failed to read {} response body: {}",
                            self.source_name, e,
                        ),
                        hint: None,
                    })?;
                    let trimmed = text.trim_start_matches('\u{feff}');
                    let body: serde_json::Value =
                        serde_json::from_str(trimmed).map_err(|e| CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "failed to parse {} JSON response: {} (body: {})",
                                self.source_name,
                                e,
                                &trimmed[..trimmed.len().min(200)],
                            ),
                            hint: None,
                        })?;

                    return Ok(body);
                }
                Err(e) => {
                    // Network/timeout errors: retry
                    if attempt == MAX_RETRIES {
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "{} upstream error after {} attempts: {}",
                                self.source_name, MAX_RETRIES, e,
                            ),
                            hint: None,
                        });
                    }

                    eprintln!(
                        "warning: retry {}/{} in {}s ({})",
                        attempt + 1,
                        MAX_RETRIES,
                        backoff_secs,
                        e,
                    );
                    thread::sleep(Duration::from_secs(backoff_secs));
                    backoff_secs *= 2;
                }
            }
        }

        unreachable!()
    }

    /// Like `request_with_retry`, but returns the raw response body as a
    /// `String` instead of parsing JSON.  Useful when the upstream may
    /// return empty or non-JSON responses on success (e.g. Fiserv
    /// settlestat for dates with no settlements).
    pub(super) fn request_with_retry_text(
        &self,
        build_request: impl Fn(&reqwest::blocking::Client) -> reqwest::blocking::RequestBuilder,
    ) -> Result<String, CliError> {
        let mut backoff_secs = 1u64;

        for attempt in 0..=MAX_RETRIES {
            let req = build_request(&self.http);
            let result = req.send();

            match result {
                Ok(resp) => {
                    let status = resp.status().as_u16();

                    if status == 401 || status == 403 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_AUTH,
                            message: format!(
                                "{} auth failed ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    if status == 400 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_VALIDATION,
                            message: format!(
                                "{} request rejected ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    if (400..500).contains(&status) && status != 429 {
                        let body: serde_json::Value =
                            resp.json().unwrap_or(serde_json::Value::Null);
                        let msg = (self.error_extractor)(&body, status);
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "{} error ({}): {}",
                                self.source_name, status, msg,
                            ),
                            hint: None,
                        });
                    }

                    if status >= 500 || status == 429 {
                        if attempt == MAX_RETRIES {
                            return Err(CliError {
                                code: exit_codes::EXIT_FETCH_UPSTREAM,
                                message: format!(
                                    "{} error (HTTP {}) after {} attempts",
                                    self.source_name, status, MAX_RETRIES,
                                ),
                                hint: None,
                            });
                        }

                        let wait = if status == 429 {
                            resp.headers()
                                .get("retry-after")
                                .and_then(|v| v.to_str().ok())
                                .and_then(|v| v.parse::<u64>().ok())
                                .unwrap_or(backoff_secs)
                        } else {
                            backoff_secs
                        };

                        eprintln!(
                            "warning: retry {}/{} in {}s (HTTP {})",
                            attempt + 1,
                            MAX_RETRIES,
                            wait,
                            status,
                        );
                        thread::sleep(Duration::from_secs(wait));
                        backoff_secs *= 2;
                        continue;
                    }

                    // Success: return raw text
                    let text = resp.text().map_err(|e| CliError {
                        code: exit_codes::EXIT_FETCH_UPSTREAM,
                        message: format!(
                            "failed to read {} response body: {}",
                            self.source_name, e,
                        ),
                        hint: None,
                    })?;

                    return Ok(text);
                }
                Err(e) => {
                    if attempt == MAX_RETRIES {
                        return Err(CliError {
                            code: exit_codes::EXIT_FETCH_UPSTREAM,
                            message: format!(
                                "{} upstream error after {} attempts: {}",
                                self.source_name, MAX_RETRIES, e,
                            ),
                            hint: None,
                        });
                    }

                    eprintln!(
                        "warning: retry {}/{} in {}s ({})",
                        attempt + 1,
                        MAX_RETRIES,
                        backoff_secs,
                        e,
                    );
                    thread::sleep(Duration::from_secs(backoff_secs));
                    backoff_secs *= 2;
                }
            }
        }

        unreachable!()
    }
}

// ── Shared helpers ──────────────────────────────────────────────────

/// Resolve an API key: flag value > environment variable > error.
pub(super) fn resolve_api_key(
    flag: Option<String>,
    source_name: &str,
    env_var: &str,
) -> Result<String, CliError> {
    if let Some(key) = flag {
        let trimmed = key.trim().to_string();
        if trimmed.is_empty() {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: format!(
                    "missing {} API key (use --api-key or set {})",
                    source_name, env_var,
                ),
                hint: None,
            });
        }
        return Ok(trimmed);
    }

    if let Ok(key) = std::env::var(env_var) {
        let trimmed = key.trim().to_string();
        if !trimmed.is_empty() {
            return Ok(trimmed);
        }
    }

    Err(CliError {
        code: exit_codes::EXIT_FETCH_NOT_AUTH,
        message: format!(
            "missing {} API key (use --api-key or set {})",
            source_name, env_var,
        ),
        hint: None,
    })
}

/// Parse and validate `--from` / `--to` date strings.
pub(super) fn parse_date_range(
    from: &str,
    to: &str,
) -> Result<(NaiveDate, NaiveDate), CliError> {
    let from_date = NaiveDate::parse_from_str(from, "%Y-%m-%d").map_err(|e| {
        CliError::args(format!("invalid --from date {:?}: {}", from, e))
    })?;
    let to_date = NaiveDate::parse_from_str(to, "%Y-%m-%d").map_err(|e| {
        CliError::args(format!("invalid --to date {:?}: {}", to, e))
    })?;

    if from_date >= to_date {
        return Err(CliError::args(format!(
            "--from ({}) must be before --to ({})",
            from_date, to_date,
        )));
    }

    Ok((from_date, to_date))
}

/// Write canonical rows to CSV (file or stdout). Returns the output label
/// for use in progress messages.
pub(crate) fn write_csv(
    rows: &[CanonicalRow],
    out: &Option<PathBuf>,
) -> Result<String, CliError> {
    let out_label = out
        .as_ref()
        .map(|p| p.display().to_string())
        .unwrap_or_else(|| "stdout".to_string());

    let writer: Box<dyn Write> = match out {
        Some(path) => {
            let f = std::fs::File::create(path).map_err(|e| {
                CliError::io(format!("cannot create {}: {}", path.display(), e))
            })?;
            Box::new(std::io::BufWriter::new(f))
        }
        None => Box::new(std::io::BufWriter::new(std::io::stdout().lock())),
    };

    let mut csv_writer = csv::WriterBuilder::new()
        .terminator(csv::Terminator::Any(b'\n'))
        .from_writer(writer);

    // Always write header, even with zero rows
    if rows.is_empty() {
        csv_writer
            .write_record([
                "effective_date",
                "posted_date",
                "amount_minor",
                "currency",
                "type",
                "source",
                "source_id",
                "group_id",
                "description",
                // Must match CanonicalRow's field order exactly: serde writes
                // the header for non-empty output, this list writes it for
                // empty output, and nothing checks that the two agree.
                "btxn_id",
            ])
            .map_err(|e| CliError::io(format!("CSV write error: {}", e)))?;
    }

    for row in rows {
        csv_writer
            .serialize(row)
            .map_err(|e| CliError::io(format!("CSV write error: {}", e)))?;
    }

    csv_writer
        .flush()
        .map_err(|e| CliError::io(format!("CSV flush error: {}", e)))?;

    Ok(out_label)
}

// ── Amount parsing (string-to-cents, no f64) ────────────────────────

/// Parse a decimal amount string to i64 minor units (cents).
/// Handles "1234.56", "1234.5", "1234", "-1234.56".
pub(crate) fn parse_money_string(s: &str) -> Result<i64, String> {
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

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_api_key_flag_priority() {
        let key = resolve_api_key(Some("  token_123  ".into()), "Test", "TEST_KEY").unwrap();
        assert_eq!(key, "token_123");
    }

    #[test]
    fn test_resolve_api_key_empty_flag() {
        let err = resolve_api_key(Some("  ".into()), "Test", "TEST_KEY").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
        assert!(err.message.contains("missing Test API key"));
    }

    #[test]
    fn test_resolve_api_key_missing() {
        std::env::remove_var("__VGRID_TEST_KEY_MISSING");
        let err = resolve_api_key(None, "Test", "__VGRID_TEST_KEY_MISSING").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_parse_date_range_valid() {
        let (from, to) = parse_date_range("2026-01-01", "2026-01-31").unwrap();
        assert_eq!(from.to_string(), "2026-01-01");
        assert_eq!(to.to_string(), "2026-01-31");
    }

    #[test]
    fn test_parse_date_range_invalid_order() {
        let err = parse_date_range("2026-01-31", "2026-01-01").unwrap_err();
        assert!(err.message.contains("--from"));
    }

    #[test]
    fn test_parse_date_range_bad_format() {
        let err = parse_date_range("not-a-date", "2026-01-31").unwrap_err();
        assert!(err.message.contains("invalid --from date"));
    }

    #[test]
    fn test_parse_money_string() {
        assert_eq!(parse_money_string("1080.47").unwrap(), 108047);
        assert_eq!(parse_money_string("0.01").unwrap(), 1);
        assert_eq!(parse_money_string("100").unwrap(), 10000);
        assert_eq!(parse_money_string("0").unwrap(), 0);
        assert_eq!(parse_money_string("0.00").unwrap(), 0);
        assert_eq!(parse_money_string("-500.25").unwrap(), -50025);
        assert_eq!(parse_money_string("10.5").unwrap(), 1050);
        assert_eq!(parse_money_string("100.").unwrap(), 10000);
        assert_eq!(parse_money_string("  42  ").unwrap(), 4200);
        assert!(parse_money_string("10.123").is_err());
        assert!(parse_money_string("abc").is_err());
    }

    /// The header is produced two ways: serde derives it from CanonicalRow
    /// when there are rows, and the empty case writes a hand-typed list. Two
    /// copies of one truth, with nothing keeping them in step — adding
    /// btxn_id updated the struct and left the literal stale until this was
    /// noticed. A consumer reading an empty fetch would have seen a schema
    /// that no longer existed.
    #[test]
    fn empty_and_populated_output_agree_on_the_header() {
        let dir = std::env::temp_dir().join(format!("vgrid_hdr_{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let empty_path = dir.join("empty.csv");
        let full_path = dir.join("full.csv");

        write_csv(&[], &Some(empty_path.clone())).unwrap();
        write_csv(
            &[CanonicalRow {
                effective_date: "2026-01-01".into(),
                posted_date: "2026-01-02".into(),
                amount_minor: 1,
                currency: "USD".into(),
                r#type: "charge".into(),
                source: "test".into(),
                source_id: "src_1".into(),
                group_id: String::new(),
                description: "x".into(),
                btxn_id: "txn_1".into(),
            }],
            &Some(full_path.clone()),
        )
        .unwrap();

        let empty_header = std::fs::read_to_string(&empty_path).unwrap();
        let full = std::fs::read_to_string(&full_path).unwrap();
        let empty_header = empty_header.lines().next().unwrap().trim().to_string();
        let full_header = full.lines().next().unwrap().trim().to_string();

        let _ = std::fs::remove_dir_all(&dir);
        assert_eq!(
            empty_header, full_header,
            "the hand-written empty-case header has drifted from CanonicalRow"
        );
    }
}
