// VisiHub API client
//
// HTTP client for communicating with the VisiHub API.
// Uses reqwest for async HTTP requests.

use std::path::Path;

use crate::hub::auth::{load_auth, AuthCredentials};
use crate::hub::types::RemoteStatus;

/// VisiHub API client
pub struct HubClient {
    http: reqwest::Client,
    api_base: String,
    token: String,
}

/// Error type for hub operations
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
}

impl std::fmt::Display for HubError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            HubError::NotAuthenticated => write!(f, "Not authenticated to VisiHub"),
            HubError::Network(msg) => write!(f, "Network error: {}", msg),
            HubError::Http(code, msg) => write!(f, "HTTP {}: {}", code, msg),
            HubError::Parse(msg) => write!(f, "Parse error: {}", msg),
            HubError::Io(msg) => write!(f, "I/O error: {}", msg),
        }
    }
}

impl std::error::Error for HubError {}

impl HubClient {
    /// Create a new client using saved auth credentials.
    /// Returns NotAuthenticated error if no credentials are saved.
    pub fn from_saved_auth() -> Result<Self, HubError> {
        let creds = load_auth().ok_or(HubError::NotAuthenticated)?;
        Ok(Self::new(creds))
    }

    /// Create a new client with explicit credentials.
    pub fn new(creds: AuthCredentials) -> Self {
        let http = reqwest::Client::builder()
            .user_agent("VisiGrid/0.1")
            .build()
            .expect("Failed to create HTTP client");

        Self {
            http,
            api_base: creds.api_base,
            token: creds.token,
        }
    }

    /// Verify the current token and get user info.
    /// GET /api/desktop/me
    pub async fn verify_token(&self) -> Result<UserInfo, HubError> {
        let url = format!("{}/api/desktop/me", self.api_base);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .await
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().await.unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        response.json::<UserInfo>().await
            .map_err(|e| HubError::Parse(e.to_string()))
    }

    /// Get the status of a dataset.
    /// GET /api/desktop/datasets/:id/status
    pub async fn get_dataset_status(&self, dataset_id: &str) -> Result<RemoteStatus, HubError> {
        let url = format!("{}/api/desktop/datasets/{}/status", self.api_base, dataset_id);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .await
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().await.unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        let json: serde_json::Value = response.json().await
            .map_err(|e| HubError::Parse(e.to_string()))?;

        Ok(RemoteStatus {
            current_revision_id: json["current_revision_id"].as_str().map(String::from),
            content_hash: json["content_hash"].as_str().map(String::from),
            byte_size: json["byte_size"].as_u64(),
            updated_at: json["updated_at"].as_str().map(String::from),
            updated_by: json["updated_by"].as_str().map(String::from),
        })
    }

    /// Download a revision's content.
    /// GET /api/desktop/revisions/:id/download
    /// Returns the raw bytes of the .sheet file.
    pub async fn download_revision(&self, revision_id: &str) -> Result<Vec<u8>, HubError> {
        let url = format!("{}/api/desktop/revisions/{}/download", self.api_base, revision_id);

        let response = self.http
            .get(&url)
            .bearer_auth(&self.token)
            .send()
            .await
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().await.unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        response.bytes().await
            .map(|b| b.to_vec())
            .map_err(|e| HubError::Network(e.to_string()))
    }
}

/// User info from /api/desktop/me
#[derive(Debug, Clone, serde::Deserialize)]
pub struct UserInfo {
    pub slug: String,
    pub email: String,
    pub plan: String,
}

/// Compute blake3 hash of a file
pub fn hash_file(path: &Path) -> Result<String, HubError> {
    let contents = std::fs::read(path)
        .map_err(|e| HubError::Io(e.to_string()))?;
    Ok(blake3::hash(&contents).to_hex().to_string())
}

/// Compute blake3 hash of bytes
pub fn hash_bytes(data: &[u8]) -> String {
    blake3::hash(data).to_hex().to_string()
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
        assert_eq!(hash1.len(), 64); // blake3 hex is 64 chars
    }
}
