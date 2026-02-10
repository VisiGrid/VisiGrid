//! Token storage — shared with desktop app.
//!
//! Reads/writes ~/.config/visigrid/auth.json (0600 on Unix).
//! If desktop app has already logged in, CLI picks it up automatically.

use std::path::PathBuf;
use serde::{Deserialize, Serialize};

/// Authentication credentials stored locally.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AuthCredentials {
    /// Bearer token for VisiHub API
    pub token: String,
    /// API base URL (e.g., "https://api.visihub.app")
    pub api_base: String,
    /// User slug (for display)
    #[serde(default)]
    pub user_slug: Option<String>,
    /// User email (for display)
    #[serde(default)]
    pub email: Option<String>,
}

impl AuthCredentials {
    pub fn new(token: String, api_base: String) -> Self {
        Self { token, api_base, user_slug: None, email: None }
    }
}

/// Returns the path to the auth credentials file.
pub fn auth_file_path() -> Option<PathBuf> {
    dirs::config_dir().map(|c| c.join("visigrid/auth.json"))
}

/// Load saved auth credentials from disk.
/// Returns None if no credentials are saved or if the file is invalid.
pub fn load_auth() -> Option<AuthCredentials> {
    let path = auth_file_path()?;
    let contents = std::fs::read_to_string(&path).ok()?;
    serde_json::from_str(&contents).ok()
}

/// Save auth credentials to disk.
/// Creates the parent directory if it doesn't exist.
/// Sets 0600 permissions on Unix.
pub fn save_auth(creds: &AuthCredentials) -> Result<(), String> {
    let path = auth_file_path().ok_or("Could not determine config directory")?;

    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("Failed to create config directory: {}", e))?;
    }

    let contents = serde_json::to_string_pretty(creds)
        .map_err(|e| format!("Failed to serialize credentials: {}", e))?;

    std::fs::write(&path, &contents)
        .map_err(|e| format!("Failed to write auth file: {}", e))?;

    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let permissions = std::fs::Permissions::from_mode(0o600);
        std::fs::set_permissions(&path, permissions)
            .map_err(|e| format!("Failed to set file permissions: {}", e))?;
    }

    Ok(())
}

/// Delete saved auth credentials.
pub fn delete_auth() -> Result<(), String> {
    let Some(path) = auth_file_path() else {
        return Ok(());
    };
    if path.exists() {
        std::fs::remove_file(&path)
            .map_err(|e| format!("Failed to delete auth file: {}", e))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_auth_credentials_roundtrip() {
        let creds = AuthCredentials {
            token: "test-token".into(),
            api_base: "https://api.visihub.app".into(),
            user_slug: Some("alice".into()),
            email: Some("alice@example.com".into()),
        };

        let json = serde_json::to_string_pretty(&creds).unwrap();
        let parsed: AuthCredentials = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.token, "test-token");
        assert_eq!(parsed.api_base, "https://api.visihub.app");
        assert_eq!(parsed.user_slug.as_deref(), Some("alice"));
        assert_eq!(parsed.email.as_deref(), Some("alice@example.com"));
    }

    #[test]
    fn test_auth_credentials_missing_optional_fields() {
        let json = r#"{"token":"tok","api_base":"https://api.visihub.app"}"#;
        let parsed: AuthCredentials = serde_json::from_str(json).unwrap();
        assert_eq!(parsed.token, "tok");
        assert!(parsed.user_slug.is_none());
        assert!(parsed.email.is_none());
    }

    #[test]
    fn test_auth_file_path_exists() {
        let path = auth_file_path();
        assert!(path.is_some());
        let path = path.unwrap();
        assert!(path.to_string_lossy().contains("visigrid"));
        assert!(path.to_string_lossy().contains("auth.json"));
    }

    #[test]
    fn test_save_and_load_auth() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("auth.json");

        // Manually write and read since save_auth uses the real config path
        let creds = AuthCredentials::new("tok123".into(), "https://api.test".into());
        let json = serde_json::to_string_pretty(&creds).unwrap();
        std::fs::write(&path, &json).unwrap();

        let contents = std::fs::read_to_string(&path).unwrap();
        let loaded: AuthCredentials = serde_json::from_str(&contents).unwrap();
        assert_eq!(loaded.token, "tok123");
        assert_eq!(loaded.api_base, "https://api.test");
    }
}
