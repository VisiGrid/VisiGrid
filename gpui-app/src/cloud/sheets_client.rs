// Sheets API client for cloud sync.
//
// Follows the same pattern as hub/client.rs: blocking reqwest + smol::unblock for async.

use std::collections::HashMap;

use crate::hub::auth::{load_auth, AuthCredentials};
use crate::hub::client::HubError;

/// Cloud sheets API client
#[derive(Clone)]
pub struct SheetsClient {
    http: reqwest::blocking::Client,
    api_base: String,
    token: String,
}

/// Sheet metadata returned by the API
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub id: i64,
    pub public_id: String,
    pub name: String,
    pub slug: String,
    pub byte_size: Option<i64>,
    pub last_edited_at: Option<String>,
    pub revision: Option<i64>,
    /// Set for sheets listed from Grid (Loco): their pid. Rails sheets: None.
    pub grid_pid: Option<String>,
}

/// Response from the save endpoint (presigned upload URL)
#[derive(Debug, Clone)]
pub struct SaveResponse {
    pub upload_url: String,
    pub headers: HashMap<String, String>,
    pub blob_key: String,
    pub revision: i64,
}

impl SheetsClient {
    /// Create a new client using saved auth credentials.
    pub fn from_saved_auth() -> Result<Self, HubError> {
        let creds = load_auth().ok_or(HubError::NotAuthenticated)?;
        Ok(Self::new(creds))
    }

    /// Create a new client with explicit credentials.
    pub fn new(creds: AuthCredentials) -> Self {
        let http = reqwest::blocking::Client::builder()
            .user_agent("VisiGrid/0.1")
            .timeout(std::time::Duration::from_secs(30))
            .build()
            .expect("Failed to create HTTP client");

        Self {
            http,
            api_base: creds.api_base,
            token: creds.token,
        }
    }

    /// Create a new cloud sheet.
    /// POST /api/desktop/sheets
    pub fn create_sheet(&self, name: &str) -> Result<SheetInfo, HubError> {
        let url = format!("{}/api/desktop/sheets", self.api_base);

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&serde_json::json!({ "name": name }))
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;

        parse_sheet_info(&json["sheet"])
            .ok_or_else(|| HubError::Parse("Invalid sheet response".to_string()))
    }

    /// List all sheets for the current user.
    /// GET /api/desktop/sheets
    pub fn list_sheets(&self) -> Result<Vec<SheetInfo>, HubError> {
        let url = format!("{}/api/desktop/sheets", self.api_base);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;

        let sheets = json["sheets"].as_array()
            .unwrap_or(&vec![])
            .iter()
            .filter_map(parse_sheet_info)
            .collect();

        Ok(sheets)
    }

    /// Get a single sheet's metadata.
    /// GET /api/desktop/sheets/:id
    pub fn get_sheet(&self, sheet_id: i64) -> Result<SheetInfo, HubError> {
        let url = format!("{}/api/desktop/sheets/{}", self.api_base, sheet_id);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;

        parse_sheet_info(&json["sheet"])
            .ok_or_else(|| HubError::Parse("Invalid sheet response".to_string()))
    }

    /// Request a presigned upload URL for saving sheet data.
    /// POST /api/desktop/sheets/:id/save
    ///
    /// `expected_revision` is the server revision the local file last matched.
    /// The server answers 409 when the sheet has moved past it (see
    /// [`conflict_revision`]). `None` skips the check — only for a deliberate
    /// overwrite.
    pub fn save_sheet(&self, sheet_id: i64, byte_size: u64, expected_revision: Option<i64>) -> Result<SaveResponse, HubError> {
        let url = format!("{}/api/desktop/sheets/{}/save", self.api_base, sheet_id);

        let mut body = serde_json::json!({
            "byte_size": byte_size,
            "client_supports_complete": true,
        });
        if let Some(rev) = expected_revision {
            body["expected_revision"] = serde_json::json!(rev);
        }

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&body)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;

        let upload_url = json["upload_url"].as_str()
            .ok_or_else(|| HubError::Parse("Missing upload_url".to_string()))?
            .to_string();

        let blob_key = json["blob_key"].as_str()
            .ok_or_else(|| HubError::Parse("Missing blob_key".to_string()))?
            .to_string();

        let revision = json["revision"].as_i64().unwrap_or(0);

        let headers = json["headers"].as_object()
            .map(|obj| {
                obj.iter()
                    .filter_map(|(k, v)| Some((k.clone(), v.as_str()?.to_string())))
                    .collect()
            })
            .unwrap_or_default();

        Ok(SaveResponse { upload_url, headers, blob_key, revision })
    }

    /// Confirm the presigned PUT finished so the server can move the pointer.
    /// POST /api/desktop/sheets/:id/save_complete
    pub fn complete_save(&self, sheet_id: i64, blob_key: &str, byte_size: u64) -> Result<i64, HubError> {
        let url = format!("{}/api/desktop/sheets/{}/save_complete", self.api_base, sheet_id);

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&serde_json::json!({
                "blob_key": blob_key,
                "byte_size": byte_size,
            }))
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;
        Ok(json["revision"].as_i64().unwrap_or(0))
    }

    /// Get presigned download URL for sheet data.
    /// GET /api/desktop/sheets/:id/data
    pub fn get_data_url(&self, sheet_id: i64) -> Result<Option<String>, HubError> {
        let url = format!("{}/api/desktop/sheets/{}/data", self.api_base, sheet_id);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json()
            .map_err(|e| HubError::Parse(e.to_string()))?;

        Ok(json["url"].as_str().map(String::from))
    }

    /// Upload file bytes to a presigned URL with custom headers.
    pub fn upload_to_url(&self, url: &str, headers: &HashMap<String, String>, data: Vec<u8>) -> Result<(), HubError> {
        let mut request = self.http.put(url);

        for (key, value) in headers {
            request = request.header(key.as_str(), value.as_str());
        }

        let response = request
            .body(data)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        Ok(())
    }

    /// Download file bytes from a presigned URL.
    pub fn download_from_url(&self, url: &str) -> Result<Vec<u8>, HubError> {
        let response = self.http
            .get(url)
            .send()
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        response.bytes()
            .map(|b| b.to_vec())
            .map_err(|e| HubError::Network(e.to_string()))
    }
}

fn parse_sheet_info(v: &serde_json::Value) -> Option<SheetInfo> {
    Some(SheetInfo {
        id: v["id"].as_i64()?,
        public_id: v["public_id"].as_str()?.to_string(),
        name: v["name"].as_str()?.to_string(),
        slug: v["slug"].as_str()?.to_string(),
        byte_size: v["byte_size"].as_i64(),
        last_edited_at: v["last_edited_at"].as_str().map(String::from),
        revision: v["revision"].as_i64(),
        grid_pid: None,
    })
}

/// The server rejected the saved token (revoked, or from another environment).
/// Signing in again is the only fix, and "Hub: Sign In" refuses while a token
/// file exists, so callers clear it via `cloud_sign_in_expired`.
pub fn is_unauthorized(err: &HubError) -> bool {
    matches!(err, HubError::Http(401, _))
}

/// The server's current revision when a save was refused because the sheet
/// changed since this client last synced (HTTP 409 from the save endpoint).
pub fn conflict_revision(err: &HubError) -> Option<i64> {
    match err {
        HubError::Http(409, body) => serde_json::from_str::<serde_json::Value>(body)
            .ok()
            .and_then(|v| v["current_revision"].as_i64()),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn conflict_revision_reads_the_servers_current_revision() {
        let err = HubError::Http(409, r#"{"error":"This sheet changed since you opened it.","current_revision":9}"#.to_string());
        assert_eq!(conflict_revision(&err), Some(9));
    }

    #[test]
    fn unauthorized_is_only_a_401() {
        assert!(is_unauthorized(&HubError::Http(401, r#"{"error":"Unauthorized"}"#.to_string())));
        assert!(!is_unauthorized(&HubError::Http(403, String::new())));
        assert!(!is_unauthorized(&HubError::NotAuthenticated));
    }

    #[test]
    fn conflict_revision_ignores_other_failures() {
        // save_complete also answers 409 (size unavailable) without a revision.
        assert_eq!(conflict_revision(&HubError::Http(409, r#"{"error":"Upload verification failed"}"#.to_string())), None);
        assert_eq!(conflict_revision(&HubError::Http(500, r#"{"current_revision":9}"#.to_string())), None);
        assert_eq!(conflict_revision(&HubError::Network("down".to_string())), None);
    }
}
