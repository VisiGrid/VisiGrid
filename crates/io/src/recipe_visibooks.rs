//! A VisiBooks report as a recipe source: the trial balance, general ledger
//! or an aging of one entity, read through the VisiBooks API.
//!
//! The API key lives in the OS keychain under the exact server origin, never
//! in the recipe or the workbook: a recipe naming another server finds no
//! key. Requests are GETs against read-only endpoints. The response body is
//! the run's snapshot, so a refresh that read the same data changes nothing
//! and a blocked run can be retried against what was read.

use chrono::{Datelike, NaiveDate};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use super::{Frame, OutColumn, Snapshot, ValueKind};
use crate::csv_import::{ColumnRule, DateOrder};

/// VisiAPI, where VisiBooks lives.
pub const DEFAULT_SERVER: &str = "https://api.visiapi.com";
/// The environment variable a headless run (CLI, CI) may hold the key in.
pub const KEY_ENV: &str = "VISIBOOKS_API_KEY";
/// With `KEY_ENV`: the server that key belongs to (default `DEFAULT_SERVER`).
pub const SERVER_ENV: &str = "VISIBOOKS_API_SERVER";
/// The largest response read: a year of a busy ledger is far below this.
const MAX_RESPONSE_BYTES: u64 = 256 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct VisibooksSource {
    #[serde(default = "default_server", skip_serializing_if = "is_default_server")]
    pub server: String,
    /// The entity id, as VisiBooks lists them for the key (`personal` for
    /// the owner ledger).
    pub entity: String,
    pub report: Report,
    /// Trial balance and agings: the date. Empty means today. Also
    /// `today`, `month_end`, `last_month_end`, `year_end`, `last_year_end`.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub as_of: String,
    /// General ledger: the period. Empty `from` is the start of this month,
    /// empty `to` today; also `month_start`, `last_month_start`,
    /// `year_start`, `last_year_start` and the `as_of` words.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub from: String,
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub to: String,
    /// `accrual` or `cash`; empty uses the entity's default.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub basis: String,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
}

fn default_server() -> String {
    DEFAULT_SERVER.to_string()
}
fn is_default_server(s: &String) -> bool {
    s == DEFAULT_SERVER
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Report {
    TrialBalance,
    GeneralLedger,
    ArAging,
    ApAging,
}

impl Report {
    pub const ALL: [Report; 4] = [Report::TrialBalance, Report::GeneralLedger, Report::ArAging, Report::ApAging];

    pub fn label(self) -> &'static str {
        match self {
            Report::TrialBalance => "Trial balance",
            Report::GeneralLedger => "General ledger",
            Report::ArAging => "AR aging",
            Report::ApAging => "AP aging",
        }
    }

    /// Whether the report covers a period (from/to) rather than a date.
    pub fn is_period(self) -> bool {
        self == Report::GeneralLedger
    }
}

impl VisibooksSource {
    pub fn new(entity: String, report: Report) -> Self {
        VisibooksSource {
            server: default_server(),
            entity,
            report,
            as_of: String::new(),
            from: String::new(),
            to: String::new(),
            basis: String::new(),
            columns: Vec::new(),
        }
    }

    /// What the recipe reads, for approvals and reports:
    /// `visibooks:https://api.visiapi.com/42/trial_balance`.
    pub fn identity(&self) -> String {
        #[cfg(feature = "native")]
        let origin = origin(&self.server).unwrap_or_else(|_| self.server.clone());
        #[cfg(not(feature = "native"))]
        let origin = self.server.trim().trim_end_matches('/').to_string();
        format!("visibooks:{origin}/{}/{}", self.entity.trim(), report_key(self.report))
    }

    /// In words, for the approval card and the strip.
    pub fn describe(&self) -> String {
        let when = if self.report.is_period() {
            format!("{} to {}", date_word(&self.from, "month_start"), date_word(&self.to, "today"))
        } else {
            format!("as of {}", date_word(&self.as_of, "today"))
        };
        #[cfg(feature = "native")]
        let host = origin(&self.server).unwrap_or_else(|_| self.server.clone());
        #[cfg(not(feature = "native"))]
        let host = self.server.trim().to_string();
        format!("{} of VisiBooks entity {}, {when}, from {host}", self.report.label(), self.entity.trim())
    }
}

fn report_key(r: Report) -> &'static str {
    match r {
        Report::TrialBalance => "trial_balance",
        Report::GeneralLedger => "general_ledger",
        Report::ArAging => "ar_aging",
        Report::ApAging => "ap_aging",
    }
}

fn date_word<'a>(s: &'a str, default: &'a str) -> &'a str {
    if s.trim().is_empty() { default } else { s.trim() }
}

#[cfg(feature = "native")]
/// `https://host[:port]` of a server address: https only, except a local
/// development server; no user, path or query, so a key bound to it can't
/// be sent anywhere else by a crafted address.
pub fn origin(server: &str) -> Result<String, String> {
    let url = reqwest::Url::parse(server.trim()).map_err(|e| format!("{server:?} is not a server address: {e}"))?;
    let host = url.host_str().ok_or_else(|| format!("{server:?} has no host"))?;
    let local = matches!(host, "localhost" | "127.0.0.1" | "[::1]");
    match url.scheme() {
        "https" => {}
        "http" if local => {}
        _ => return Err(format!("{server:?}: VisiBooks is read over https")),
    }
    if !url.username().is_empty() || url.password().is_some() {
        return Err(format!("{server:?}: put the API key in the keychain, not in the address"));
    }
    if url.path() != "/" && !url.path().is_empty() || url.query().is_some() || url.fragment().is_some() {
        return Err(format!("{server:?}: give only the server (like {DEFAULT_SERVER}), no path"));
    }
    Ok(match url.port() {
        Some(p) => format!("{}://{host}:{p}", url.scheme()),
        None => format!("{}://{host}", url.scheme()),
    })
}

/// A date word or YYYY-MM-DD as a date, relative to `today`.
pub fn resolve_date(s: &str, today: NaiveDate) -> Result<NaiveDate, String> {
    let first = |y: i32, m: u32| NaiveDate::from_ymd_opt(y, m, 1).unwrap();
    let month_start = first(today.year(), today.month());
    let last_month_start = if today.month() == 1 { first(today.year() - 1, 12) } else { first(today.year(), today.month() - 1) };
    let next_month_start = if today.month() == 12 { first(today.year() + 1, 1) } else { first(today.year(), today.month() + 1) };
    let day_before = |d: NaiveDate| d.pred_opt().unwrap();
    Ok(match s.trim() {
        "" | "today" => today,
        "month_start" => month_start,
        "month_end" => day_before(next_month_start),
        "last_month_start" => last_month_start,
        "last_month_end" => day_before(month_start),
        "year_start" => first(today.year(), 1),
        "year_end" => NaiveDate::from_ymd_opt(today.year(), 12, 31).unwrap(),
        "last_year_start" => first(today.year() - 1, 1),
        "last_year_end" => NaiveDate::from_ymd_opt(today.year() - 1, 12, 31).unwrap(),
        other => NaiveDate::parse_from_str(other, "%Y-%m-%d").map_err(|_| {
            format!("{other:?} is not a date; use YYYY-MM-DD or a word such as month_start, last_month_end, year_start, today")
        })?,
    })
}

#[cfg(feature = "native")]
/// The URL a run requests, dates resolved against `today`.
pub fn request_url(src: &VisibooksSource, today: NaiveDate) -> Result<String, String> {
    let origin = origin(&src.server)?;
    let entity = src.entity.trim();
    if entity.is_empty() || !entity.chars().all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-') {
        return Err("the recipe names no VisiBooks entity; set source.entity to an id from the key's entity list".into());
    }
    let basis = match src.basis.trim() {
        "" => None,
        b @ ("accrual" | "cash") => Some(b),
        other => return Err(format!("basis {other:?}: use accrual or cash")),
    };
    let mut query: Vec<(&str, String)> = vec![("entity_id", entity.to_string())];
    let path = match src.report {
        Report::TrialBalance => {
            query.push(("as_of", resolve_date(&src.as_of, today)?.to_string()));
            "trial_balance"
        }
        Report::GeneralLedger => {
            let from = resolve_date(if src.from.trim().is_empty() { "month_start" } else { &src.from }, today)?;
            let to = resolve_date(&src.to, today)?;
            if from > to {
                return Err(format!("the period runs backwards: from {from} to {to}"));
            }
            query.push(("from", from.to_string()));
            query.push(("to", to.to_string()));
            "general_ledger"
        }
        Report::ArAging | Report::ApAging => {
            query.push(("side", if src.report == Report::ArAging { "ar" } else { "ap" }.to_string()));
            query.push(("as_of", resolve_date(&src.as_of, today)?.to_string()));
            "aging_detail"
        }
    };
    if let Some(b) = basis {
        query.push(("basis", b.to_string()));
    }
    let mut url = reqwest::Url::parse(&format!("{origin}/api/v1/books/agent/{path}")).map_err(|e| e.to_string())?;
    url.query_pairs_mut().extend_pairs(query.iter().map(|(k, v)| (*k, v.as_str())));
    Ok(url.to_string())
}

#[cfg(feature = "native")]
/// The keychain account a server's key is stored under.
pub fn keychain_account(origin: &str) -> String {
    format!("visibooks/{origin}")
}

#[cfg(feature = "native")]
/// The API key for a server: the keychain entry for exactly that origin,
/// else `VISIBOOKS_API_KEY` when `VISIBOOKS_API_SERVER` (default VisiAPI)
/// is that origin.
pub fn api_key(origin: &str) -> Result<String, String> {
    if let Some(key) = visigrid_config::secrets::get(&keychain_account(origin)) {
        return Ok(key);
    }
    if let Ok(key) = std::env::var(KEY_ENV) {
        let env_server = std::env::var(SERVER_ENV).unwrap_or_else(|_| DEFAULT_SERVER.to_string());
        if !key.trim().is_empty() && self::origin(&env_server).ok().as_deref() == Some(origin) {
            return Ok(key.trim().to_string());
        }
    }
    Err(format!(
        "no VisiBooks API key for {origin}; save one with `vgrid visibooks key` (or in VisiGrid), or set {KEY_ENV} for headless runs"
    ))
}

#[cfg(feature = "native")]
/// Read the report: one GET, the body is the snapshot.
pub fn fetch(src: &VisibooksSource, today: NaiveDate) -> Result<Snapshot, String> {
    // No key is the likelier problem than anything in the recipe: say so first
    let origin = origin(&src.server)?;
    api_key(&origin)?;
    let url = request_url(src, today)?;
    let body = get(&origin, &url, Some(src.entity.trim()))?;
    Ok(Snapshot::from_bytes(std::path::Path::new(&src.identity()), body))
}

#[cfg(feature = "native")]
/// The entities the saved key reaches, as (id, name).
pub fn entities(origin: &str) -> Result<Vec<(String, String)>, String> {
    let body = get(origin, &format!("{origin}/api/v1/books/agent/entities"), None)?;
    let json: Value = serde_json::from_slice(&body).map_err(|e| format!("VisiBooks' answer isn't JSON: {e}"))?;
    Ok(json
        .get("entities")
        .and_then(Value::as_array)
        .map(|list| list.iter().map(|e| (text(&e["entity_id"]), text(&e["name"]))).collect())
        .unwrap_or_default())
}

#[cfg(feature = "native")]
/// One authenticated GET to `url` on `origin`, read up to the size cap.
fn get(origin: &str, url: &str, entity: Option<&str>) -> Result<Vec<u8>, String> {
    let key = api_key(origin)?;
    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(300))
        // Edge proxies reject requests without a User-Agent
        .user_agent(concat!("VisiGrid/", env!("CARGO_PKG_VERSION")))
        // The key goes to this origin only, never to a redirect target
        .redirect(reqwest::redirect::Policy::none())
        .build()
        .map_err(|e| e.to_string())?;
    let response = client
        .get(url)
        .bearer_auth(key)
        .header("Accept", "application/json")
        .send()
        .map_err(|e| format!("couldn't reach VisiBooks at {origin}: {e}"))?;
    let status = response.status();
    let too_big = || format!("VisiBooks sent more than {} MB; use a shorter period", MAX_RESPONSE_BYTES / (1024 * 1024));
    if response.content_length().is_some_and(|n| n > MAX_RESPONSE_BYTES) {
        return Err(too_big());
    }
    let mut body = Vec::new();
    use std::io::Read;
    response.take(MAX_RESPONSE_BYTES + 1).read_to_end(&mut body).map_err(|e| format!("reading VisiBooks' answer: {e}"))?;
    if body.len() as u64 > MAX_RESPONSE_BYTES {
        return Err(too_big());
    }
    if !status.is_success() {
        let message = serde_json::from_slice::<Value>(&body)
            .ok()
            .and_then(|v| {
                v.pointer("/error/message").or_else(|| v.get("message")).or_else(|| v.get("error")).and_then(Value::as_str).map(str::to_string)
            })
            .unwrap_or_else(|| status.canonical_reason().unwrap_or("error").to_string());
        return Err(match (status.as_u16(), entity) {
            (401, _) => format!("VisiBooks at {origin} didn't accept the API key; save a new one"),
            (403, Some(e)) => format!("the API key can't read entity {e}: {message}"),
            (300..=399, _) => format!("VisiBooks at {origin} answered with a redirect; check the server address"),
            _ => format!("VisiBooks answered {}: {message}", status.as_u16()),
        });
    }
    Ok(body)
}

/// Cents as a decimal: -123456 -> "-1234.56".
fn money(minor: i64) -> String {
    let sign = if minor < 0 { "-" } else { "" };
    let m = minor.unsigned_abs();
    format!("{sign}{}.{:02}", m / 100, m % 100)
}

fn text(v: &Value) -> String {
    match v {
        Value::Null => String::new(),
        Value::String(s) => s.clone(),
        other => other.to_string(),
    }
}

fn minor(v: &Value, key: &str) -> Result<i64, String> {
    match v.get(key) {
        // A missing field is an API change, not a zero: never load a table
        // of zeros that still balances
        None => Err(format!("VisiBooks' answer has no {key}; this VisiGrid may be older than the server")),
        Some(Value::Null) => Ok(0),
        Some(n) => n.as_i64().ok_or_else(|| format!("VisiBooks sent {key} = {n}, not a whole number of cents")),
    }
}

/// The snapshot's rows as a recipe frame.
pub(super) fn read_frame(src: &VisibooksSource, snapshot: &Snapshot) -> Result<Frame, String> {
    let json: Value = serde_json::from_slice(&snapshot.bytes).map_err(|e| format!("VisiBooks' answer isn't JSON: {e}"))?;
    let (key, cols): (&str, Vec<(&str, ColumnRule)>) = match src.report {
        Report::TrialBalance => (
            "lines",
            vec![
                ("Account", ColumnRule::Text),
                ("Name", ColumnRule::Text),
                ("Type", ColumnRule::Text),
                ("Debit", ColumnRule::Number),
                ("Credit", ColumnRule::Number),
                ("Balance", ColumnRule::Number),
            ],
        ),
        Report::GeneralLedger => (
            "lines",
            vec![
                ("Date", ColumnRule::Date(DateOrder::Ymd)),
                ("Entry", ColumnRule::Text),
                ("Account", ColumnRule::Text),
                ("Name", ColumnRule::Text),
                ("Type", ColumnRule::Text),
                ("Memo", ColumnRule::Text),
                ("Contact", ColumnRule::Text),
                ("Debit", ColumnRule::Number),
                ("Credit", ColumnRule::Number),
            ],
        ),
        Report::ArAging | Report::ApAging => (
            "rows",
            vec![
                (if src.report == Report::ArAging { "Customer" } else { "Vendor" }, ColumnRule::Text),
                (if src.report == Report::ArAging { "Invoice" } else { "Bill" }, ColumnRule::Text),
                ("Issue date", ColumnRule::Date(DateOrder::Ymd)),
                ("Due date", ColumnRule::Date(DateOrder::Ymd)),
                ("Days past due", ColumnRule::Number),
                ("Bucket", ColumnRule::Text),
                ("Currency", ColumnRule::Text),
                ("Balance", ColumnRule::Number),
            ],
        ),
    };
    let items = json.get(key).and_then(Value::as_array).ok_or_else(|| format!("VisiBooks' answer has no {key}"))?;
    let mut rows = Vec::with_capacity(items.len());
    for item in items {
        rows.push(match src.report {
            Report::TrialBalance => {
                let (d, c) = (minor(item, "debit_minor")?, minor(item, "credit_minor")?);
                vec![
                    text(&item["account_code"]),
                    text(&item["account_name"]),
                    text(&item["account_type"]),
                    money(d),
                    money(c),
                    money(d - c),
                ]
            }
            Report::GeneralLedger => vec![
                text(&item["date"]),
                text(&item["entry_id"]),
                text(&item["account_code"]),
                text(&item["account_name"]),
                text(&item["account_type"]),
                text(&item["memo"]),
                text(&item["contact"]),
                money(minor(item, "debit_minor")?),
                money(minor(item, "credit_minor")?),
            ],
            Report::ArAging | Report::ApAging => vec![
                text(&item["contact"]),
                text(&item["number"]),
                text(&item["issue_date"]),
                text(&item["due_date"]),
                text(&item["days_past_due"]),
                text(&item["bucket"]),
                text(&item["currency"]),
                money(minor(item, "balance_due_minor")?),
            ],
        });
    }
    let lines = (1..=rows.len()).collect();
    Ok(Frame {
        columns: cols
            .into_iter()
            .map(|(name, rule)| OutColumn {
                name: name.to_string(),
                rule,
                kind: if rule == ColumnRule::Number { ValueKind::Number } else { ValueKind::Plain },
            })
            .collect(),
        rows,
        lines,
        files: Vec::new(),
        file_names: Vec::new(),
        decimal_comma: false,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn d(s: &str) -> NaiveDate {
        NaiveDate::parse_from_str(s, "%Y-%m-%d").unwrap()
    }

    #[cfg(feature = "native")]
    #[test]
    fn origins_are_https_servers_only() {
        assert_eq!(origin("https://api.visiapi.com").unwrap(), "https://api.visiapi.com");
        assert_eq!(origin("https://api.visiapi.com/").unwrap(), "https://api.visiapi.com");
        assert_eq!(origin("http://localhost:3000").unwrap(), "http://localhost:3000");
        assert!(origin("http://api.visiapi.com").is_err());
        assert!(origin("https://user:pw@api.visiapi.com").is_err());
        assert!(origin("https://api.visiapi.com/api/v1").is_err());
        assert!(origin("https://api.visiapi.com?x=1").is_err());
        assert!(origin("file:///etc/passwd").is_err());
    }

    #[test]
    fn dates_resolve_relative_to_today() {
        let today = d("2026-03-15");
        assert_eq!(resolve_date("", today).unwrap(), today);
        assert_eq!(resolve_date("month_start", today).unwrap(), d("2026-03-01"));
        assert_eq!(resolve_date("month_end", today).unwrap(), d("2026-03-31"));
        assert_eq!(resolve_date("last_month_start", today).unwrap(), d("2026-02-01"));
        assert_eq!(resolve_date("last_month_end", today).unwrap(), d("2026-02-28"));
        assert_eq!(resolve_date("last_month_start", d("2026-01-10")).unwrap(), d("2025-12-01"));
        assert_eq!(resolve_date("last_year_end", today).unwrap(), d("2025-12-31"));
        assert_eq!(resolve_date("2026-02-01", today).unwrap(), d("2026-02-01"));
        assert!(resolve_date("02/01/2026", today).is_err());
    }

    #[cfg(feature = "native")]
    #[test]
    fn requests_name_only_read_endpoints() {
        let today = d("2026-03-15");
        let mut src = VisibooksSource::new("42".into(), Report::GeneralLedger);
        src.from = "last_month_start".into();
        src.to = "last_month_end".into();
        src.basis = "cash".into();
        assert_eq!(
            request_url(&src, today).unwrap(),
            "https://api.visiapi.com/api/v1/books/agent/general_ledger?entity_id=42&from=2026-02-01&to=2026-02-28&basis=cash"
        );
        src.report = Report::ApAging;
        src.basis.clear();
        assert_eq!(request_url(&src, today).unwrap(), "https://api.visiapi.com/api/v1/books/agent/aging_detail?entity_id=42&side=ap&as_of=2026-03-15");
        src.entity = "42&side=ar".into();
        assert!(request_url(&src, today).is_err(), "an entity can't add parameters");
        src.entity = "personal".into();
        src.report = Report::TrialBalance;
        src.as_of = "year_end".into();
        assert!(request_url(&src, today).unwrap().ends_with("trial_balance?entity_id=personal&as_of=2026-12-31"));
    }

    #[test]
    fn reports_become_typed_frames_with_exact_amounts() {
        let src = VisibooksSource::new("42".into(), Report::TrialBalance);
        let body = br#"{"basis":"accrual","as_of":"2026-03-15","lines":[
            {"account_code":"1010","account_name":"Cash","account_type":"asset","debit_minor":123456,"credit_minor":0},
            {"account_code":"4010","account_name":"Sales","account_type":"revenue","debit_minor":5,"credit_minor":123461}]}"#;
        let snap = Snapshot::from_bytes(std::path::Path::new(&src.identity()), body.to_vec());
        let frame = read_frame(&src, &snap).unwrap();
        assert_eq!(frame.columns.iter().map(|c| c.name.as_str()).collect::<Vec<_>>(), ["Account", "Name", "Type", "Debit", "Credit", "Balance"]);
        assert_eq!(frame.rows[0], ["1010", "Cash", "asset", "1234.56", "0.00", "1234.56"]);
        assert_eq!(frame.rows[1][5], "-1234.56");
        assert_eq!(frame.columns[0].rule, ColumnRule::Text, "account codes stay text");
        let bad = Snapshot::from_bytes(std::path::Path::new("x"), br#"{"lines":[{"debit_minor":1.5,"credit_minor":0}]}"#.to_vec());
        assert!(read_frame(&src, &bad).err().unwrap().contains("whole number of cents"));
        // A renamed field fails rather than reading as 0
        let renamed = Snapshot::from_bytes(std::path::Path::new("x"), br#"{"lines":[{"debit_cents":100,"credit_minor":0}]}"#.to_vec());
        assert!(read_frame(&src, &renamed).err().unwrap().contains("has no debit_minor"));
        // An explicit null is a zero
        let null = Snapshot::from_bytes(std::path::Path::new("x"), br#"{"lines":[{"debit_minor":null,"credit_minor":5}]}"#.to_vec());
        assert_eq!(read_frame(&src, &null).unwrap().rows[0][3], "0.00");
    }
}
