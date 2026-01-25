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

    /// List available repos.
    /// GET /api/desktop/repos
    pub async fn list_repos(&self) -> Result<Vec<RepoInfo>, HubError> {
        let url = format!("{}/api/desktop/repos", self.api_base);

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

        let repos = json.as_array()
            .unwrap_or(&vec![])
            .iter()
            .filter_map(|r| {
                Some(RepoInfo {
                    owner: r["owner"].as_str()?.to_string(),
                    slug: r["slug"].as_str()?.to_string(),
                    name: r["name"].as_str()?.to_string(),
                })
            })
            .collect();

        Ok(repos)
    }

    /// List datasets in a repo.
    /// GET /api/desktop/repos/:owner/:slug/datasets
    pub async fn list_datasets(&self, owner: &str, slug: &str) -> Result<Vec<DatasetInfo>, HubError> {
        let url = format!("{}/api/desktop/repos/{}/{}/datasets", self.api_base, owner, slug);

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

        let datasets = json.as_array()
            .unwrap_or(&vec![])
            .iter()
            .filter_map(|d| {
                Some(DatasetInfo {
                    id: d["id"].as_str()?.to_string(),
                    name: d["name"].as_str()?.to_string(),
                })
            })
            .collect();

        Ok(datasets)
    }

    /// Create a new dataset in a repo.
    /// POST /api/desktop/repos/:owner/:slug/datasets
    pub async fn create_dataset(&self, owner: &str, slug: &str, name: &str) -> Result<String, HubError> {
        let url = format!("{}/api/desktop/repos/{}/{}/datasets", self.api_base, owner, slug);

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&serde_json::json!({ "name": name }))
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

        json["dataset_id"].as_str()
            .map(String::from)
            .ok_or_else(|| HubError::Parse("Missing dataset_id in response".to_string()))
    }

    /// Create a new revision for publishing.
    /// POST /api/desktop/datasets/:id/revisions
    /// Returns (revision_id, upload_url)
    pub async fn create_revision(
        &self,
        dataset_id: &str,
        content_hash: &str,
        byte_size: u64,
    ) -> Result<(String, String), HubError> {
        let url = format!("{}/api/desktop/datasets/{}/revisions", self.api_base, dataset_id);

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&serde_json::json!({
                "content_hash": content_hash,
                "byte_size": byte_size
            }))
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

        let revision_id = json["revision_id"].as_str()
            .ok_or_else(|| HubError::Parse("Missing revision_id".to_string()))?
            .to_string();

        let upload_url = json["upload_url"].as_str()
            .ok_or_else(|| HubError::Parse("Missing upload_url".to_string()))?
            .to_string();

        Ok((revision_id, upload_url))
    }

    /// Upload file bytes to signed URL.
    /// PUT to the signed R2 URL.
    pub async fn upload_to_signed_url(&self, upload_url: &str, data: Vec<u8>) -> Result<(), HubError> {
        let response = self.http
            .put(upload_url)
            .header("Content-Type", "application/octet-stream")
            .body(data)
            .send()
            .await
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().await.unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        Ok(())
    }

    /// Complete a revision after upload.
    /// POST /api/desktop/revisions/:id/complete
    pub async fn complete_revision(&self, revision_id: &str, content_hash: &str) -> Result<(), HubError> {
        let url = format!("{}/api/desktop/revisions/{}/complete", self.api_base, revision_id);

        let response = self.http
            .post(&url)
            .bearer_auth(&self.token)
            .json(&serde_json::json!({ "content_hash": content_hash }))
            .send()
            .await
            .map_err(|e| HubError::Network(e.to_string()))?;

        let status = response.status().as_u16();
        if !response.status().is_success() {
            let body = response.text().await.unwrap_or_default();
            return Err(HubError::Http(status, body));
        }

        Ok(())
    }
}

/// User info from /api/desktop/me
#[derive(Debug, Clone, serde::Deserialize)]
pub struct UserInfo {
    pub slug: String,
    pub email: String,
    pub plan: String,
}

/// Repository info from /api/desktop/repos
#[derive(Debug, Clone)]
pub struct RepoInfo {
    pub owner: String,
    pub slug: String,
    pub name: String,
}

/// Dataset info from /api/desktop/repos/:owner/:slug/datasets
#[derive(Debug, Clone)]
pub struct DatasetInfo {
    pub id: String,
    pub name: String,
}

/// Compute blake3 hash of a file (with algorithm prefix for future-proofing)
pub fn hash_file(path: &Path) -> Result<String, HubError> {
    let contents = std::fs::read(path)
        .map_err(|e| HubError::Io(e.to_string()))?;
    Ok(format!("blake3:{}", blake3::hash(&contents).to_hex()))
}

/// Compute blake3 hash of bytes (with algorithm prefix for future-proofing)
pub fn hash_bytes(data: &[u8]) -> String {
    format!("blake3:{}", blake3::hash(data).to_hex())
}

/// Check if two hashes match (handles prefix comparison)
pub fn hashes_match(a: &str, b: &str) -> bool {
    // Strip prefix if present for comparison
    fn normalize(s: &str) -> &str {
        s.strip_prefix("blake3:").unwrap_or(s)
    }
    normalize(a) == normalize(b)
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
        assert_eq!(hash1.len(), 7 + 64); // "blake3:" prefix + 64 char hex
    }

    #[test]
    fn test_hashes_match() {
        // Same hash, same format
        assert!(hashes_match("blake3:abc123", "blake3:abc123"));

        // Same hash, different formats (with and without prefix)
        assert!(hashes_match("blake3:abc123", "abc123"));
        assert!(hashes_match("abc123", "blake3:abc123"));

        // Different hashes
        assert!(!hashes_match("blake3:abc123", "blake3:def456"));
        assert!(!hashes_match("abc123", "def456"));
    }
}
