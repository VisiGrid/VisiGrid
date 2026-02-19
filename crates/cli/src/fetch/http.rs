//! `vgrid fetch http` — fetch data from any HTTP API into canonical CSV.
//!
//! Uses a mapping file to transform JSON responses into the standard
//! 9-column CanonicalRow format. Auth credentials are resolved from
//! environment variables only (never inline secrets).

use std::collections::HashMap;
use std::path::PathBuf;

use serde::Deserialize;

use crate::exit_codes;
use crate::CliError;

use super::common::{self, CanonicalRow, FetchClient};

// ── Constants ───────────────────────────────────────────────────────

const MAX_RESPONSE_BYTES: usize = 10 * 1024 * 1024; // 10 MB
const DEFAULT_MAX_ITEMS: usize = 10_000;
const DEFAULT_TIMEOUT_SECS: u64 = 15;

// ── Mapping config types ────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct MappingConfig {
    /// JSONPath-like selector for the root array (e.g., "$.payments")
    pub root: String,

    /// How --from/--to map to query parameters
    #[serde(default)]
    pub params: HashMap<String, ParamMapping>,

    /// Column mappings: key is the canonical column name
    pub columns: HashMap<String, ColumnMapping>,

    /// Sort order for deterministic output
    #[serde(default)]
    pub sort_by: Vec<String>,
}

#[derive(Debug, Deserialize)]
pub struct ParamMapping {
    /// Query parameter name
    pub query: String,

    /// Date format: "iso" (YYYY-MM-DD) or "unix_s" or "unix_ms"
    #[serde(default = "default_date_format")]
    pub format: String,
}

fn default_date_format() -> String {
    "iso".to_string()
}

#[derive(Debug, Deserialize)]
#[serde(untagged)]
pub enum ColumnMapping {
    /// Simple path shorthand: "$.field.name"
    Path(String),

    /// Full column spec with type, transform, optional flag
    Spec(ColumnSpec),
}

#[derive(Debug, Deserialize)]
pub struct ColumnSpec {
    /// JSONPath-like path (e.g., "$.amount")
    #[serde(default)]
    pub path: Option<String>,

    /// Constant value (e.g., "booking_api")
    #[serde(rename = "const")]
    #[serde(default)]
    pub const_value: Option<String>,

    /// Value type: "string", "int", "datetime", "decimal"
    #[serde(rename = "type")]
    #[serde(default = "default_type")]
    pub col_type: String,

    /// Transform: "cents", "upper", "lower"
    #[serde(default)]
    pub transform: Option<String>,

    /// Value mapping for enums: {"payment": "charge", "*": "adjustment"}
    #[serde(default)]
    pub map: Option<HashMap<String, String>>,

    /// Date format hint for parsing: "iso", "unix_s", "unix_ms"
    #[serde(default)]
    pub format: Option<String>,

    /// If true, missing field produces empty string instead of error
    #[serde(default)]
    pub optional: bool,
}

fn default_type() -> String {
    "string".to_string()
}

// ── Auth resolution ─────────────────────────────────────────────────

#[derive(Debug)]
enum AuthMethod {
    None,
    Bearer(String),
    Header(String, String),
    Basic(String, String),
}

fn resolve_auth(auth_str: &str) -> Result<AuthMethod, CliError> {
    if auth_str == "none" {
        return Ok(AuthMethod::None);
    }

    if let Some(env_var) = auth_str.strip_prefix("bearer-env:") {
        let token = resolve_env(env_var, "bearer token")?;
        return Ok(AuthMethod::Bearer(token));
    }

    if let Some(rest) = auth_str.strip_prefix("header-env:") {
        let parts: Vec<&str> = rest.splitn(2, ':').collect();
        if parts.len() != 2 {
            return Err(CliError {
                code: exit_codes::EXIT_USAGE,
                message: "header-env format: header-env:HEADER_NAME:ENV_VAR".into(),
                hint: Some("example: --auth header-env:X-API-Key:MY_API_KEY".into()),
            });
        }
        let value = resolve_env(parts[1], &format!("header {}", parts[0]))?;
        return Ok(AuthMethod::Header(parts[0].to_string(), value));
    }

    if let Some(rest) = auth_str.strip_prefix("basic-env:") {
        let parts: Vec<&str> = rest.splitn(2, ':').collect();
        if parts.len() != 2 {
            return Err(CliError {
                code: exit_codes::EXIT_USAGE,
                message: "basic-env format: basic-env:USER_ENV:PASS_ENV".into(),
                hint: Some("example: --auth basic-env:API_USER:API_PASS".into()),
            });
        }
        let user = resolve_env(parts[0], "basic auth username")?;
        let pass = resolve_env(parts[1], "basic auth password")?;
        return Ok(AuthMethod::Basic(user, pass));
    }

    Err(CliError {
        code: exit_codes::EXIT_USAGE,
        message: format!("unknown auth method: {}", auth_str),
        hint: Some("supported: none, bearer-env:VAR, header-env:NAME:VAR, basic-env:USER:PASS".into()),
    })
}

fn resolve_env(var_name: &str, label: &str) -> Result<String, CliError> {
    std::env::var(var_name).map_err(|_| CliError {
        code: exit_codes::EXIT_FETCH_NOT_AUTH,
        message: format!("environment variable {} not set (needed for {})", var_name, label),
        hint: Some(format!("export {}=<value>", var_name)),
    }).and_then(|v| {
        let trimmed = v.trim().to_string();
        if trimmed.is_empty() {
            Err(CliError {
                code: exit_codes::EXIT_FETCH_NOT_AUTH,
                message: format!("environment variable {} is empty (needed for {})", var_name, label),
                hint: None,
            })
        } else {
            Ok(trimmed)
        }
    })
}

// ── JSON path extraction ────────────────────────────────────────────

/// Extract a value from a JSON object using a simple dot-path selector.
///
/// Supports: `$.field`, `$.nested.field`, `$.array[0].field`
/// Does NOT support full JSONPath (filters, wildcards, etc.).
fn json_extract<'a>(value: &'a serde_json::Value, path: &str) -> Option<&'a serde_json::Value> {
    let path = path.strip_prefix("$.").unwrap_or(path);
    let mut current = value;

    for segment in path.split('.') {
        // Handle array indexing: field[0]
        if let Some(bracket_pos) = segment.find('[') {
            let field = &segment[..bracket_pos];
            let idx_str = &segment[bracket_pos + 1..segment.len() - 1];

            if !field.is_empty() {
                current = current.get(field)?;
            }

            let idx: usize = idx_str.parse().ok()?;
            current = current.get(idx)?;
        } else {
            current = current.get(segment)?;
        }
    }

    Some(current)
}

/// Convert a JSON value to a string representation for CSV output.
fn json_value_to_string(value: &serde_json::Value) -> String {
    match value {
        serde_json::Value::String(s) => s.clone(),
        serde_json::Value::Number(n) => n.to_string(),
        serde_json::Value::Bool(b) => b.to_string(),
        serde_json::Value::Null => String::new(),
        other => other.to_string(),
    }
}

// ── Column extraction ───────────────────────────────────────────────

fn extract_column(
    item: &serde_json::Value,
    col_name: &str,
    mapping: &ColumnMapping,
) -> Result<String, CliError> {
    match mapping {
        ColumnMapping::Path(path) => {
            match json_extract(item, path) {
                Some(v) => Ok(json_value_to_string(v)),
                None => Err(mapping_error(format!(
                    "missing required field '{}' (path: {})",
                    col_name, path
                ))),
            }
        }
        ColumnMapping::Spec(spec) => {
            // Constant value — no extraction needed
            if let Some(ref const_val) = spec.const_value {
                return Ok(const_val.clone());
            }

            let path = spec.path.as_deref().ok_or_else(|| {
                mapping_error(format!(
                    "column '{}' needs either 'path' or 'const'",
                    col_name
                ))
            })?;

            let raw = match json_extract(item, path) {
                Some(v) => json_value_to_string(v),
                None => {
                    if spec.optional {
                        return Ok(String::new());
                    }
                    return Err(mapping_error(format!(
                        "missing required field '{}' (path: {})",
                        col_name, path
                    )));
                }
            };

            // Apply value mapping (enum translation)
            let mapped = if let Some(ref value_map) = spec.map {
                value_map
                    .get(&raw)
                    .or_else(|| value_map.get("*"))
                    .cloned()
                    .unwrap_or(raw)
            } else {
                raw
            };

            // Apply transform
            let transformed = match spec.transform.as_deref() {
                Some("upper") => mapped.to_uppercase(),
                Some("lower") => mapped.to_lowercase(),
                Some("cents") => {
                    // Value is already in cents (minor units) — pass through as integer
                    let n: i64 = mapped.parse().map_err(|_| {
                        mapping_error(format!(
                            "column '{}': cannot parse '{}' as integer for cents transform",
                            col_name, mapped
                        ))
                    })?;
                    n.to_string()
                }
                Some("dollars_to_cents") => {
                    // Convert dollar amount string to cents
                    let cents = common::parse_money_string(&mapped).map_err(|e| {
                        mapping_error(format!(
                            "column '{}': cannot parse '{}' as dollar amount: {}",
                            col_name, mapped, e
                        ))
                    })?;
                    cents.to_string()
                }
                Some(other) => {
                    return Err(mapping_error(format!(
                        "column '{}': unknown transform '{}'",
                        col_name, other
                    )));
                }
                None => mapped,
            };

            Ok(transformed)
        }
    }
}

fn mapping_error(msg: String) -> CliError {
    CliError {
        code: exit_codes::EXIT_FETCH_MAPPING,
        message: msg,
        hint: None,
    }
}

// ── Build CanonicalRow from JSON item ───────────────────────────────

fn item_to_row(
    item: &serde_json::Value,
    config: &MappingConfig,
) -> Result<CanonicalRow, CliError> {
    Ok(CanonicalRow {
        effective_date: extract_column(item, "effective_date", col(config, "effective_date")?)?,
        posted_date: extract_column(item, "posted_date", col(config, "posted_date")?)?,
        amount_minor: extract_column(item, "amount_minor", col(config, "amount_minor")?)?
            .parse::<i64>()
            .map_err(|e| mapping_error(format!("amount_minor not a valid integer: {}", e)))?,
        currency: extract_column(item, "currency", col(config, "currency")?)?,
        r#type: extract_column(item, "type", col(config, "type")?)?,
        source: extract_column(item, "source", col(config, "source")?)?,
        source_id: extract_column(item, "source_id", col(config, "source_id")?)?,
        group_id: extract_column(item, "group_id", col(config, "group_id")?)?,
        description: extract_column(item, "description", col(config, "description")?)?,
    })
}

fn col<'a>(
    config: &'a MappingConfig,
    name: &str,
) -> Result<&'a ColumnMapping, CliError> {
    config.columns.get(name).ok_or_else(|| {
        mapping_error(format!(
            "mapping file missing required column '{}'",
            name
        ))
    })
}

// ── Sort rows deterministically ─────────────────────────────────────

fn sort_rows(rows: &mut [CanonicalRow], sort_by: &[String]) {
    let keys: Vec<&str> = if sort_by.is_empty() {
        vec!["group_id", "effective_date", "source_id"]
    } else {
        sort_by.iter().map(|s| s.as_str()).collect()
    };

    rows.sort_by(|a, b| {
        for key in &keys {
            let cmp = match *key {
                "effective_date" => a.effective_date.cmp(&b.effective_date),
                "posted_date" => a.posted_date.cmp(&b.posted_date),
                "amount_minor" => a.amount_minor.cmp(&b.amount_minor),
                "currency" => a.currency.cmp(&b.currency),
                "type" => a.r#type.cmp(&b.r#type),
                "source" => a.source.cmp(&b.source),
                "source_id" => a.source_id.cmp(&b.source_id),
                "group_id" => a.group_id.cmp(&b.group_id),
                "description" => a.description.cmp(&b.description),
                _ => std::cmp::Ordering::Equal,
            };
            if cmp != std::cmp::Ordering::Equal {
                return cmp;
            }
        }
        std::cmp::Ordering::Equal
    });
}

// ── Main command ────────────────────────────────────────────────────

pub fn cmd_fetch_http(
    url: String,
    auth: String,
    from: String,
    to: String,
    map_file: PathBuf,
    out: Option<PathBuf>,
    save_raw: Option<PathBuf>,
    sample: bool,
    timeout: Option<u64>,
    max_items: Option<usize>,
    quiet: bool,
) -> Result<(), CliError> {
    // 1. Validate HTTPS
    if !url.starts_with("https://") {
        return Err(CliError {
            code: exit_codes::EXIT_USAGE,
            message: "only HTTPS URLs are allowed".into(),
            hint: Some(format!("change {} to https://", url)),
        });
    }

    // 2. Parse dates
    let (from_date, to_date) = common::parse_date_range(&from, &to)?;

    // 3. Load mapping config
    let mapping_str = std::fs::read_to_string(&map_file).map_err(|e| CliError {
        code: exit_codes::EXIT_USAGE,
        message: format!("cannot read mapping file {}: {}", map_file.display(), e),
        hint: None,
    })?;
    let config: MappingConfig = serde_json::from_str(&mapping_str).map_err(|e| {
        mapping_error(format!("invalid mapping file: {}", e))
    })?;

    // 4. Resolve auth
    let auth_method = resolve_auth(&auth)?;

    // 5. Build request URL with query params
    let timeout_secs = timeout.unwrap_or(DEFAULT_TIMEOUT_SECS);
    let item_cap = max_items.unwrap_or(DEFAULT_MAX_ITEMS);

    let mut request_url = url::Url::parse(&url).map_err(|e| {
        CliError::args(format!("invalid URL: {}", e))
    })?;

    // Map --from/--to into query params via mapping config
    for (param_key, param_mapping) in &config.params {
        let date = match param_key.as_str() {
            "from" => from_date,
            "to" => to_date,
            other => {
                return Err(mapping_error(format!(
                    "unknown param key '{}' (expected 'from' or 'to')",
                    other
                )));
            }
        };
        let value = match param_mapping.format.as_str() {
            "iso" => date.format("%Y-%m-%d").to_string(),
            "unix_s" => date
                .and_hms_opt(0, 0, 0)
                .unwrap()
                .and_utc()
                .timestamp()
                .to_string(),
            "unix_ms" => (date
                .and_hms_opt(0, 0, 0)
                .unwrap()
                .and_utc()
                .timestamp()
                * 1000)
                .to_string(),
            other => {
                return Err(mapping_error(format!(
                    "unknown date format '{}' for param '{}'",
                    other, param_key
                )));
            }
        };
        request_url
            .query_pairs_mut()
            .append_pair(&param_mapping.query, &value);
    }

    if !quiet {
        eprintln!("Fetching {}...", request_url.as_str());
    }

    // 6. Execute request
    let client = FetchClient::new("HTTP", |body, status| {
        // Generic error extractor for unknown APIs
        if let Some(msg) = body.get("message").and_then(|v| v.as_str()) {
            msg.to_string()
        } else if let Some(err) = body.get("error").and_then(|v| v.as_str()) {
            err.to_string()
        } else {
            format!("HTTP {}", status)
        }
    });

    // Build a custom client with the right timeout
    let http = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(timeout_secs))
        .user_agent(common::USER_AGENT)
        .build()
        .map_err(|e| CliError {
            code: exit_codes::EXIT_ERROR,
            message: format!("failed to build HTTP client: {}", e),
            hint: None,
        })?;

    let response_body = client.request_with_retry(|_| {
        let mut req = http.get(request_url.as_str());

        // Apply auth
        req = match &auth_method {
            AuthMethod::None => req,
            AuthMethod::Bearer(token) => req.bearer_auth(token),
            AuthMethod::Header(name, value) => req.header(name.as_str(), value.as_str()),
            AuthMethod::Basic(user, pass) => req.basic_auth(user, Some(pass)),
        };

        req
    })?;

    // Check response size (approximate — already parsed, but guards against huge objects)
    let raw_json = serde_json::to_string(&response_body).unwrap_or_default();
    if raw_json.len() > MAX_RESPONSE_BYTES {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_OVERFLOW,
            message: format!(
                "response too large ({} bytes, max {} bytes)",
                raw_json.len(),
                MAX_RESPONSE_BYTES
            ),
            hint: Some("narrow the date range or increase --max-response-bytes".into()),
        });
    }

    // 7. Save raw response if requested
    if let Some(ref raw_path) = save_raw {
        let pretty = serde_json::to_string_pretty(&response_body).unwrap_or_default();
        std::fs::write(raw_path, pretty.as_bytes()).map_err(|e| {
            CliError::io(format!("cannot write raw response to {}: {}", raw_path.display(), e))
        })?;
        if !quiet {
            eprintln!("Raw response saved to {}", raw_path.display());
        }
    }

    // 8. Sample mode — print raw and exit
    if sample {
        let pretty = serde_json::to_string_pretty(&response_body).unwrap_or_default();
        println!("{}", pretty);
        return Ok(());
    }

    // 9. Extract root array
    let root_path = &config.root;
    let items_value = json_extract(&response_body, root_path).ok_or_else(|| {
        mapping_error(format!(
            "root path '{}' not found in response",
            root_path
        ))
    })?;

    let items = items_value.as_array().ok_or_else(|| {
        mapping_error(format!(
            "root path '{}' resolved to {}, expected array",
            root_path,
            match items_value {
                serde_json::Value::Object(_) => "an object",
                serde_json::Value::String(_) => "a string",
                serde_json::Value::Number(_) => "a number",
                serde_json::Value::Bool(_) => "a boolean",
                serde_json::Value::Null => "null",
                _ => "non-array",
            }
        ))
    })?;

    // 10. Check item cap
    if items.len() > item_cap {
        return Err(CliError {
            code: exit_codes::EXIT_FETCH_OVERFLOW,
            message: format!(
                "response contains {} items, max {} allowed",
                items.len(),
                item_cap
            ),
            hint: Some("narrow the date range or increase --max-items".into()),
        });
    }

    if !quiet {
        eprintln!("Extracted {} items from {}", items.len(), root_path);
    }

    // 11. Map each item to CanonicalRow
    let mut rows: Vec<CanonicalRow> = Vec::with_capacity(items.len());
    for (idx, item) in items.iter().enumerate() {
        match item_to_row(item, &config) {
            Ok(row) => rows.push(row),
            Err(mut e) => {
                e.message = format!("item [{}]: {}", idx, e.message);
                return Err(e);
            }
        }
    }

    // 12. Sort deterministically
    sort_rows(&mut rows, &config.sort_by);

    // 13. Write CSV output
    let out_label = common::write_csv(&rows, &out)?;

    if !quiet {
        eprintln!(
            "Wrote {} rows to {}",
            rows.len(),
            out_label,
        );
    }

    Ok(())
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_item() -> serde_json::Value {
        serde_json::json!({
            "id": "pay_001",
            "booking_id": "bk_100",
            "amount": 5000,
            "currency": "usd",
            "kind": "payment",
            "created_at": "2026-01-15",
            "settled_at": "2026-01-17",
            "memo": "Room charge"
        })
    }

    fn sample_mapping() -> MappingConfig {
        serde_json::from_str(r#"{
            "root": "$.payments",
            "params": {
                "from": { "query": "start_date", "format": "iso" },
                "to": { "query": "end_date", "format": "iso" }
            },
            "columns": {
                "effective_date": { "path": "$.created_at", "type": "datetime", "format": "iso" },
                "posted_date": { "path": "$.settled_at", "type": "datetime", "format": "iso", "optional": true },
                "amount_minor": { "path": "$.amount", "type": "int", "transform": "cents" },
                "currency": { "path": "$.currency", "type": "string", "transform": "upper" },
                "type": { "path": "$.kind", "type": "string", "map": { "payment": "charge", "refund": "refund", "*": "adjustment" } },
                "source": { "const": "booking_api" },
                "source_id": { "path": "$.id", "type": "string" },
                "group_id": { "path": "$.booking_id", "type": "string", "optional": true },
                "description": { "path": "$.memo", "type": "string", "optional": true }
            },
            "sort_by": ["effective_date", "source_id"]
        }"#).unwrap()
    }

    #[test]
    fn test_json_extract_simple() {
        let json = serde_json::json!({"a": {"b": {"c": 42}}});
        assert_eq!(json_extract(&json, "$.a.b.c").unwrap(), &serde_json::json!(42));
    }

    #[test]
    fn test_json_extract_array_index() {
        let json = serde_json::json!({"items": [{"id": 1}, {"id": 2}]});
        assert_eq!(
            json_extract(&json, "$.items[1].id").unwrap(),
            &serde_json::json!(2)
        );
    }

    #[test]
    fn test_json_extract_missing() {
        let json = serde_json::json!({"a": 1});
        assert!(json_extract(&json, "$.b.c").is_none());
    }

    #[test]
    fn test_item_to_row_golden() {
        let item = sample_item();
        let config = sample_mapping();
        let row = item_to_row(&item, &config).unwrap();

        assert_eq!(row.effective_date, "2026-01-15");
        assert_eq!(row.posted_date, "2026-01-17");
        assert_eq!(row.amount_minor, 5000);
        assert_eq!(row.currency, "USD");
        assert_eq!(row.r#type, "charge");
        assert_eq!(row.source, "booking_api");
        assert_eq!(row.source_id, "pay_001");
        assert_eq!(row.group_id, "bk_100");
        assert_eq!(row.description, "Room charge");
    }

    #[test]
    fn test_item_to_row_value_map_wildcard() {
        let item = serde_json::json!({
            "id": "pay_002",
            "booking_id": "bk_101",
            "amount": 1000,
            "currency": "eur",
            "kind": "credit_note",
            "created_at": "2026-01-16",
            "settled_at": "2026-01-18",
            "memo": "Adjustment"
        });
        let config = sample_mapping();
        let row = item_to_row(&item, &config).unwrap();

        assert_eq!(row.r#type, "adjustment"); // wildcard match
    }

    #[test]
    fn test_item_to_row_optional_missing() {
        let item = serde_json::json!({
            "id": "pay_003",
            "amount": 2000,
            "currency": "usd",
            "kind": "payment",
            "created_at": "2026-01-20"
        });
        let config = sample_mapping();
        let row = item_to_row(&item, &config).unwrap();

        assert_eq!(row.posted_date, "");
        assert_eq!(row.group_id, "");
        assert_eq!(row.description, "");
    }

    #[test]
    fn test_item_to_row_required_missing() {
        let item = serde_json::json!({
            "id": "pay_004",
            "amount": 3000,
            "currency": "usd",
            "kind": "payment"
            // missing created_at (required effective_date)
        });
        let config = sample_mapping();
        let err = item_to_row(&item, &config).unwrap_err();

        assert_eq!(err.code, exit_codes::EXIT_FETCH_MAPPING);
        assert!(err.message.contains("effective_date"));
    }

    #[test]
    fn test_sort_determinism() {
        let mut rows = vec![
            CanonicalRow {
                effective_date: "2026-01-15".into(),
                posted_date: "".into(),
                amount_minor: 100,
                currency: "USD".into(),
                r#type: "charge".into(),
                source: "api".into(),
                source_id: "b".into(),
                group_id: "".into(),
                description: "".into(),
            },
            CanonicalRow {
                effective_date: "2026-01-15".into(),
                posted_date: "".into(),
                amount_minor: 200,
                currency: "USD".into(),
                r#type: "charge".into(),
                source: "api".into(),
                source_id: "a".into(),
                group_id: "".into(),
                description: "".into(),
            },
        ];

        sort_rows(&mut rows, &["effective_date".into(), "source_id".into()]);

        assert_eq!(rows[0].source_id, "a");
        assert_eq!(rows[1].source_id, "b");
    }

    #[test]
    fn test_sort_default_order() {
        let mut rows = vec![
            CanonicalRow {
                effective_date: "2026-01-15".into(),
                posted_date: "".into(),
                amount_minor: 100,
                currency: "USD".into(),
                r#type: "charge".into(),
                source: "api".into(),
                source_id: "b".into(),
                group_id: "g2".into(),
                description: "".into(),
            },
            CanonicalRow {
                effective_date: "2026-01-15".into(),
                posted_date: "".into(),
                amount_minor: 200,
                currency: "USD".into(),
                r#type: "charge".into(),
                source: "api".into(),
                source_id: "a".into(),
                group_id: "g1".into(),
                description: "".into(),
            },
        ];

        sort_rows(&mut rows, &[]); // default: group_id, effective_date, source_id

        assert_eq!(rows[0].group_id, "g1"); // g1 < g2
        assert_eq!(rows[1].group_id, "g2");
    }

    #[test]
    fn test_mapping_config_parse() {
        let json = r#"{
            "root": "$.data",
            "columns": {
                "effective_date": "$.date",
                "posted_date": { "path": "$.posted", "optional": true },
                "amount_minor": { "path": "$.amount", "transform": "cents" },
                "currency": { "const": "USD" },
                "type": "$.type",
                "source": { "const": "test" },
                "source_id": "$.id",
                "group_id": { "path": "$.group", "optional": true },
                "description": { "path": "$.desc", "optional": true }
            }
        }"#;

        let config: MappingConfig = serde_json::from_str(json).unwrap();
        assert_eq!(config.root, "$.data");
        assert_eq!(config.columns.len(), 9);
        assert!(config.sort_by.is_empty()); // defaults
    }

    #[test]
    fn test_auth_bearer_env() {
        std::env::set_var("__TEST_TOKEN", "secret123");
        let auth = resolve_auth("bearer-env:__TEST_TOKEN").unwrap();
        match auth {
            AuthMethod::Bearer(t) => assert_eq!(t, "secret123"),
            _ => panic!("expected Bearer"),
        }
        std::env::remove_var("__TEST_TOKEN");
    }

    #[test]
    fn test_auth_missing_env() {
        std::env::remove_var("__TEST_MISSING_TOKEN");
        let err = resolve_auth("bearer-env:__TEST_MISSING_TOKEN").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_FETCH_NOT_AUTH);
    }

    #[test]
    fn test_auth_none() {
        let auth = resolve_auth("none").unwrap();
        matches!(auth, AuthMethod::None);
    }

    #[test]
    fn test_auth_unknown() {
        let err = resolve_auth("oauth:something").unwrap_err();
        assert_eq!(err.code, exit_codes::EXIT_USAGE);
    }

    /// Golden output snapshot: full pipeline from multi-item JSON response
    /// through mapping, sorting, and CSV serialization. If this test breaks,
    /// reconciliation fingerprints will break in production.
    #[test]
    fn test_golden_csv_output() {
        // Simulate a 4-item API response with intentionally unordered items,
        // mixed types (charge, refund, wildcard), optional nulls, and
        // dollar-to-cents conversion to exercise every code path.
        let response = serde_json::json!({
            "transactions": [
                {
                    "txn_id": "txn_004",
                    "ref": "inv_200",
                    "total": "100.00",
                    "cur": "eur",
                    "category": "credit_note",
                    "date": "2026-01-18",
                    "settled": "2026-01-20",
                    "note": "Goodwill credit"
                },
                {
                    "txn_id": "txn_002",
                    "ref": "inv_100",
                    "total": "250.75",
                    "cur": "usd",
                    "category": "refund",
                    "date": "2026-01-15",
                    "settled": null,
                    "note": "Refund for overbilling"
                },
                {
                    "txn_id": "txn_001",
                    "ref": "inv_100",
                    "total": "1080.47",
                    "cur": "usd",
                    "category": "payment",
                    "date": "2026-01-15",
                    "settled": "2026-01-17",
                    "note": "Invoice payment"
                },
                {
                    "txn_id": "txn_003",
                    "total": "50.00",
                    "cur": "usd",
                    "category": "payment",
                    "date": "2026-01-15",
                    "note": "Ad-hoc charge"
                }
            ]
        });

        let config: MappingConfig = serde_json::from_str(r#"{
            "root": "$.transactions",
            "columns": {
                "effective_date": "$.date",
                "posted_date": { "path": "$.settled", "optional": true },
                "amount_minor": { "path": "$.total", "transform": "dollars_to_cents" },
                "currency": { "path": "$.cur", "transform": "upper" },
                "type": { "path": "$.category", "map": { "payment": "charge", "refund": "refund", "*": "adjustment" } },
                "source": { "const": "billing_api" },
                "source_id": "$.txn_id",
                "group_id": { "path": "$.ref", "optional": true },
                "description": { "path": "$.note", "optional": true }
            },
            "sort_by": ["effective_date", "source_id"]
        }"#).unwrap();

        // Extract root array
        let items = json_extract(&response, &config.root)
            .unwrap()
            .as_array()
            .unwrap();

        // Map items
        let mut rows: Vec<CanonicalRow> = items
            .iter()
            .map(|item| item_to_row(item, &config).unwrap())
            .collect();

        // Sort
        sort_rows(&mut rows, &config.sort_by);

        // Serialize to CSV
        let mut buf = Vec::new();
        {
            let mut wtr = csv::WriterBuilder::new()
                .terminator(csv::Terminator::Any(b'\n'))
                .from_writer(&mut buf);
            for row in &rows {
                wtr.serialize(row).unwrap();
            }
            wtr.flush().unwrap();
        }
        let csv_output = String::from_utf8(buf).unwrap();

        // The golden output. Column order matches CanonicalRow struct field
        // order (serde default). Sort order is effective_date then source_id.
        //
        // If you change CanonicalRow fields, column order, sort logic, or
        // transform behavior, this test MUST be updated deliberately.
        let expected = "\
effective_date,posted_date,amount_minor,currency,type,source,source_id,group_id,description
2026-01-15,2026-01-17,108047,USD,charge,billing_api,txn_001,inv_100,Invoice payment
2026-01-15,,25075,USD,refund,billing_api,txn_002,inv_100,Refund for overbilling
2026-01-15,,5000,USD,charge,billing_api,txn_003,,Ad-hoc charge
2026-01-18,2026-01-20,10000,EUR,adjustment,billing_api,txn_004,inv_200,Goodwill credit
";

        assert_eq!(
            csv_output, expected,
            "\n\nGolden CSV output mismatch!\n\
             If this is intentional (column order, sort, or transform change),\n\
             update the expected string. If not, you broke reproducibility.\n\n\
             GOT:\n{}\nEXPECTED:\n{}",
            csv_output, expected
        );
    }

    #[test]
    fn test_dollars_to_cents_transform() {
        let item = serde_json::json!({
            "id": "pay_005",
            "booking_id": "bk_102",
            "amount_str": "42.50",
            "currency": "usd",
            "kind": "payment",
            "created_at": "2026-01-20",
            "memo": "test"
        });

        let config: MappingConfig = serde_json::from_str(r#"{
            "root": "$.payments",
            "columns": {
                "effective_date": "$.created_at",
                "posted_date": { "path": "$.settled", "optional": true },
                "amount_minor": { "path": "$.amount_str", "transform": "dollars_to_cents" },
                "currency": { "path": "$.currency", "transform": "upper" },
                "type": { "path": "$.kind", "map": { "payment": "charge", "*": "adjustment" } },
                "source": { "const": "test_api" },
                "source_id": "$.id",
                "group_id": { "path": "$.booking_id", "optional": true },
                "description": { "path": "$.memo", "optional": true }
            }
        }"#).unwrap();

        let row = item_to_row(&item, &config).unwrap();
        assert_eq!(row.amount_minor, 4250);
    }
}
