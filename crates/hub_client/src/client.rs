//! VisiHub HTTP client.
//!
//! Blocking reqwest client (no Tokio runtime required).
//! Covers the full publish flow: create revision → upload → complete → poll.

use std::path::Path;
use std::thread;
use std::time::Duration;

use crate::auth::{load_auth, AuthCredentials};

/// VisiHub API client (blocking).
#[derive(Clone)]
pub struct HubClient {
    http: reqwest::blocking::Client,
    api_base: String,
    token: String,
}

/// Error type for hub operations.
#[derive(Debug)]
pub enum HubError {
    /// No auth credentials configured
    NotAuthenticated,
    /// Network error
    Network(String),
    /// HTTP error with status code
    Http(u16, String),
    /// JSON parsing error
    Parse(String),
    /// File I/O error
    Io(String),
    /// Server returned a validation error (4xx with message)
    Validation(String),
    /// Timeout waiting for processing
    Timeout(String),
}

impl std::fmt::Display for HubError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            HubError::NotAuthenticated => write!(f, "Not authenticated — run `vgrid login` first"),
            HubError::Network(msg) => write!(f, "Network error: {}", msg),
            HubError::Http(code, msg) => write!(f, "HTTP {}: {}", code, msg),
            HubError::Parse(msg) => write!(f, "Parse error: {}", msg),
            HubError::Io(msg) => write!(f, "I/O error: {}", msg),
            HubError::Validation(msg) => write!(f, "{}", msg),
            HubError::Timeout(msg) => write!(f, "Timeout: {}", msg),
        }
    }
}

impl std::error::Error for HubError {}

/// Options for creating a revision (publish flow step 1).
#[derive(Debug, Clone, Default)]
pub struct CreateRevisionOptions {
    pub source_type: Option<String>,
    pub source_identity: Option<String>,
    pub query_hash: Option<String>,
}

/// Status of a run (from the runs API).
#[derive(Debug, Clone, serde::Serialize)]
pub struct RunResult {
    pub run_id: String,
    pub version: u64,
    pub status: String,
    pub check_status: Option<String>,
    pub diff_summary: Option<serde_json::Value>,
    pub row_count: Option<u64>,
    pub col_count: Option<u64>,
    pub content_hash: Option<String>,
    pub source_metadata: Option<serde_json::Value>,
    pub proof_url: String,
}

/// User info from /api/desktop/me
#[derive(Debug, Clone, serde::Deserialize, serde::Serialize)]
pub struct UserInfo {
    #[serde(alias = "user_slug")]
    pub slug: String,
    pub email: String,
    pub plan: String,
}

/// Repository info
#[derive(Debug, Clone)]
pub struct RepoInfo {
    pub owner: String,
    pub slug: String,
    pub name: String,
}

/// Dataset info
#[derive(Debug, Clone)]
pub struct DatasetInfo {
    pub id: String,
    pub name: String,
}

impl HubClient {
    /// Create a new client using saved auth credentials.
    pub fn from_saved_auth() -> Result<Self, HubError> {
        let creds = load_auth().ok_or(HubError::NotAuthenticated)?;
        Ok(Self::new(creds))
    }

    /// Create a new client with explicit credentials.
    pub fn new(creds: AuthCredentials) -> Self {
        let http = reqwest::blocking::Client::builder()
            .user_agent(format!("vgrid/{}", env!("CARGO_PKG_VERSION")))
            .timeout(Duration::from_secs(60))
            .build()
            .expect("Failed to create HTTP client");

        Self {
            http,
            api_base: creds.api_base,
            token: creds.token,
        }
    }

    /// Verify the current token and get user info.
    pub fn verify_token(&self) -> Result<UserInfo, HubError> {
        let url = format!("{}/api/desktop/me", self.api_base);
        let resp = self.get(&url)?;
        resp.json::<UserInfo>().map_err(|e| HubError::Parse(e.to_string()))
    }

    /// List datasets in a repo.
    pub fn list_datasets(&self, owner: &str, slug: &str) -> Result<Vec<DatasetInfo>, HubError> {
        let url = format!("{}/api/desktop/repos/{}/{}/datasets", self.api_base, owner, slug);
        let resp = self.get(&url)?;
        let json: serde_json::Value = resp.json().map_err(|e| HubError::Parse(e.to_string()))?;

        let datasets = json.as_array()
            .unwrap_or(&vec![])
            .iter()
            .filter_map(|d| {
                let id = d["id"].as_i64()
                    .map(|n| n.to_string())
                    .or_else(|| d["id"].as_str().map(String::from))?;
                Some(DatasetInfo {
                    id,
                    name: d["name"].as_str()?.to_string(),
                })
            })
            .collect();

        Ok(datasets)
    }

    /// Create a new dataset in a repo.
    pub fn create_dataset(&self, owner: &str, slug: &str, name: &str) -> Result<String, HubError> {
        let url = format!("{}/api/desktop/repos/{}/{}/datasets", self.api_base, owner, slug);
        let resp = self.post_json(&url, &serde_json::json!({ "name": name }))?;
        let json: serde_json::Value = resp.json().map_err(|e| HubError::Parse(e.to_string()))?;

        json["dataset_id"].as_i64()
            .map(|n| n.to_string())
            .or_else(|| json["dataset_id"].as_str().map(String::from))
            .ok_or_else(|| HubError::Parse("Missing dataset_id in response".into()))
    }

    /// Create a new revision (publish flow step 1).
    /// Returns (revision_id, upload_url, upload_headers).
    pub fn create_revision(
        &self,
        dataset_id: &str,
        content_hash: &str,
        byte_size: u64,
        opts: &CreateRevisionOptions,
    ) -> Result<(String, String, serde_json::Value), HubError> {
        let url = format!("{}/api/desktop/datasets/{}/revisions", self.api_base, dataset_id);

        let mut body = serde_json::json!({
            "content_hash": content_hash,
            "byte_size": byte_size,
        });

        // Attach source metadata if provided
        if opts.source_type.is_some() || opts.source_identity.is_some() || opts.query_hash.is_some() {
            let mut sm = serde_json::Map::new();
            if let Some(ref t) = opts.source_type {
                sm.insert("type".into(), serde_json::Value::String(t.clone()));
            }
            if let Some(ref id) = opts.source_identity {
                sm.insert("identity".into(), serde_json::Value::String(id.clone()));
            }
            if let Some(ref qh) = opts.query_hash {
                sm.insert("query_hash".into(), serde_json::Value::String(qh.clone()));
            }
            sm.insert("timestamp".into(), serde_json::Value::String(
                chrono_now_utc()
            ));
            body["source_metadata"] = serde_json::Value::Object(sm);
        }

        let resp = self.post_json(&url, &body)?;
        let json: serde_json::Value = resp.json().map_err(|e| HubError::Parse(e.to_string()))?;

        let revision_id = json_str(&json, "revision_id")?;
        let upload_url = json_str(&json, "upload_url")?;
        let upload_headers = json.get("upload_headers")
            .cloned()
            .unwrap_or(serde_json::Value::Object(serde_json::Map::new()));

        Ok((revision_id, upload_url, upload_headers))
    }

    /// Upload file bytes to signed URL (publish flow step 2).
    pub fn upload_bytes(&self, upload_url: &str, data: Vec<u8>, headers: &serde_json::Value) -> Result<(), HubError> {
        let mut req = self.http.put(upload_url)
            .header("Content-Type", "application/octet-stream");

        // Apply upload headers from the server
        if let Some(obj) = headers.as_object() {
            for (k, v) in obj {
                if let Some(val) = v.as_str() {
                    req = req.header(k.as_str(), val);
                }
            }
        }

        let response = req.body(data)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        Ok(())
    }

    /// Complete a revision after upload (publish flow step 3).
    pub fn complete_revision(&self, revision_id: &str, content_hash: &str) -> Result<(), HubError> {
        let url = format!("{}/api/desktop/revisions/{}/complete", self.api_base, revision_id);
        self.post_json(&url, &serde_json::json!({ "content_hash": content_hash }))?;
        Ok(())
    }

    /// Poll until the run reaches a terminal state (publish flow step 4).
    /// Returns the final run result, or timeout error.
    pub fn poll_run(
        &self,
        owner: &str,
        slug: &str,
        revision_id: &str,
        timeout: Duration,
    ) -> Result<RunResult, HubError> {
        let start = std::time::Instant::now();
        let poll_interval = Duration::from_secs(3);
        let proof_url = format!("{}/api/repos/{}/{}/runs/{}/proof",
            self.api_base, owner, slug, revision_id);

        loop {
            if start.elapsed() > timeout {
                return Err(HubError::Timeout(
                    format!("Import did not complete within {}s", timeout.as_secs())
                ));
            }

            let url = format!("{}/api/repos/{}/{}/runs?limit=5", self.api_base, owner, slug);
            let resp = self.get(&url)?;
            let json: serde_json::Value = resp.json()
                .map_err(|e| HubError::Parse(e.to_string()))?;

            let rev_id_num: i64 = revision_id.parse().unwrap_or(-1);

            if let Some(runs) = json["runs"].as_array() {
                if let Some(run) = runs.iter().find(|r| r["id"].as_i64() == Some(rev_id_num)) {
                    let status = run["status"].as_str().unwrap_or("unknown");
                    match status {
                        "verified" | "completed" => {
                            return Ok(RunResult {
                                run_id: revision_id.to_string(),
                                version: run["version"].as_u64().unwrap_or(0),
                                status: status.to_string(),
                                check_status: run["check_status"].as_str().map(String::from),
                                diff_summary: run.get("diff_summary").cloned(),
                                row_count: run["row_count"].as_u64(),
                                col_count: run["col_count"].as_u64(),
                                content_hash: run["content_hash"].as_str().map(String::from),
                                source_metadata: run.get("source_metadata").cloned(),
                                proof_url: proof_url.clone(),
                            });
                        }
                        "failed" => {
                            return Err(HubError::Http(500, "Import failed on server".into()));
                        }
                        _ => {
                            // Still processing
                        }
                    }
                }
            }

            thread::sleep(poll_interval);
        }
    }

    /// Build the proof URL for a run.
    pub fn proof_url(&self, owner: &str, slug: &str, run_id: &str) -> String {
        format!("{}/api/repos/{}/{}/runs/{}/proof", self.api_base, owner, slug, run_id)
    }

    // ── Internal helpers ────────────────────────────────────────────

    fn get(&self, url: &str) -> Result<reqwest::blocking::Response, HubError> {
        let response = self.http.get(url)
            .bearer_auth(&self.token)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            if status == 422 || status == 400 {
                return Err(HubError::Validation(body));
            }
            return Err(HubError::Http(status, body));
        }

        Ok(response)
    }

    fn post_json(&self, url: &str, body: &serde_json::Value) -> Result<reqwest::blocking::Response, HubError> {
        let response = self.http.post(url)
            .bearer_auth(&self.token)
            .json(body)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            if status == 422 || status == 400 {
                return Err(HubError::Validation(body));
            }
            return Err(HubError::Http(status, body));
        }

        Ok(response)
    }
}

// ── Free functions ──────────────────────────────────────────────────

/// Compute blake3 hash of a file (with algorithm prefix).
pub fn hash_file(path: &Path) -> Result<String, HubError> {
    let contents = std::fs::read(path)
        .map_err(|e| HubError::Io(e.to_string()))?;
    Ok(format!("blake3:{}", blake3::hash(&contents).to_hex()))
}

/// Compute blake3 hash of bytes (with algorithm prefix).
pub fn hash_bytes(data: &[u8]) -> String {
    format!("blake3:{}", blake3::hash(data).to_hex())
}

/// Get current UTC time as ISO 8601 string (no chrono dependency).
fn chrono_now_utc() -> String {
    // Use std::time to avoid adding chrono as a dependency
    let now = std::time::SystemTime::now();
    let duration = now.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
    let secs = duration.as_secs();

    // Simple UTC formatting
    let days = secs / 86400;
    let time_secs = secs % 86400;
    let hours = time_secs / 3600;
    let minutes = (time_secs % 3600) / 60;
    let seconds = time_secs % 60;

    // Calculate year/month/day from days since epoch (1970-01-01)
    let (year, month, day) = days_to_ymd(days);

    format!("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        year, month, day, hours, minutes, seconds)
}

fn days_to_ymd(mut days: u64) -> (u64, u64, u64) {
    // Simplified date calculation
    let mut year = 1970;
    loop {
        let days_in_year = if is_leap(year) { 366 } else { 365 };
        if days < days_in_year { break; }
        days -= days_in_year;
        year += 1;
    }
    let month_days: &[u64] = if is_leap(year) {
        &[31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        &[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };
    let mut month = 1;
    for &md in month_days {
        if days < md { break; }
        days -= md;
        month += 1;
    }
    (year, month, days + 1)
}

fn is_leap(year: u64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn json_str(json: &serde_json::Value, key: &str) -> Result<String, HubError> {
    json[key].as_i64()
        .map(|n| n.to_string())
        .or_else(|| json[key].as_str().map(String::from))
        .ok_or_else(|| HubError::Parse(format!("Missing {} in response", key)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hash_bytes() {
        let hash1 = hash_bytes(b"hello world");
        let hash2 = hash_bytes(b"hello world");
        let hash3 = hash_bytes(b"different");

        assert_eq!(hash1, hash2);
        assert_ne!(hash1, hash3);
        assert!(hash1.starts_with("blake3:"));
        assert_eq!(hash1.len(), 7 + 64);
    }

    #[test]
    fn test_chrono_now_utc_format() {
        let ts = chrono_now_utc();
        // Should match YYYY-MM-DDTHH:MM:SSZ
        assert_eq!(ts.len(), 20);
        assert!(ts.ends_with('Z'));
        assert_eq!(&ts[4..5], "-");
        assert_eq!(&ts[7..8], "-");
        assert_eq!(&ts[10..11], "T");
    }

    #[test]
    fn test_days_to_ymd() {
        // 1970-01-01
        assert_eq!(days_to_ymd(0), (1970, 1, 1));
        // 2024-01-01 = 19723 days since epoch
        let (y, m, d) = days_to_ymd(19723);
        assert_eq!(y, 2024);
        assert_eq!(m, 1);
        assert_eq!(d, 1);
    }

    #[test]
    fn test_check_failed_run_result() {
        // Golden test: a "check failed" run must have check_status="fail"
        // and the CLI keys off this exact string value for exit code 41.
        let result = RunResult {
            run_id: "99".into(),
            version: 5,
            status: "verified".into(),
            check_status: Some("fail".into()),
            diff_summary: Some(serde_json::json!({
                "row_count_change": -50,
                "col_count_change": 2,
            })),
            row_count: Some(950),
            col_count: Some(17),
            content_hash: Some("blake3:deadbeef".into()),
            source_metadata: Some(serde_json::json!({"type": "dbt", "identity": "models/payments"})),
            proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/99/proof".into(),
        };

        let json = serde_json::to_value(&result).unwrap();
        assert_eq!(json["check_status"].as_str(), Some("fail"));
        assert_eq!(json["status"].as_str(), Some("verified"));

        // The CLI check: check_status == Some("fail") triggers exit code 41
        assert_eq!(result.check_status.as_deref(), Some("fail"));

        // Passing run must NOT trigger failure
        let passing = RunResult {
            check_status: Some("pass".into()),
            ..result
        };
        assert_ne!(passing.check_status.as_deref(), Some("fail"));
    }

    #[test]
    fn test_run_result_json_schema() {
        let result = RunResult {
            run_id: "42".into(),
            version: 3,
            status: "verified".into(),
            check_status: Some("pass".into()),
            diff_summary: Some(serde_json::json!({
                "row_count_change": 10,
                "col_count_change": 0,
            })),
            row_count: Some(1000),
            col_count: Some(15),
            content_hash: Some("blake3:abc123".into()),
            source_metadata: Some(serde_json::json!({"type": "dbt"})),
            proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/42/proof".into(),
        };

        let json = serde_json::to_value(&result).unwrap();
        assert!(json["run_id"].is_string());
        assert!(json["version"].is_number());
        assert!(json["status"].is_string());
        assert!(json["check_status"].is_string());
        assert!(json["diff_summary"].is_object());
        assert!(json["proof_url"].is_string());
        assert!(json["source_metadata"].is_object());
    }
}
