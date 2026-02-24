// Device token authentication
//
// Stores and retrieves the device token for API authentication.
// Token is stored in ~/.config/visigrid/auth.json

use std::path::PathBuf;
use serde::{Deserialize, Serialize};

/// Authentication credentials stored locally
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AuthCredentials {
    /// Bearer token for API
    pub token: String,
    /// API base URL (e.g., "https://api.visiapi.com")
    pub api_base: String,
    /// User slug (for display purposes)
    pub user_slug: Option<String>,
    /// User email (for display purposes)
    pub email: Option<String>,
}

impl AuthCredentials {
    pub fn new(token: String, api_base: String) -> Self {
        Self {
            token,
            api_base,
            user_slug: None,
            email: None,
        }
    }
}

/// Returns the path to the auth credentials file.
/// On macOS: ~/.config/visigrid/auth.json
/// On Windows: %APPDATA%/visigrid/auth.json
/// On Linux: ~/.config/visigrid/auth.json
pub fn auth_file_path() -> Option<PathBuf> {
    #[cfg(target_os = "macos")]
    {
        dirs::home_dir().map(|h| h.join(".config/visigrid/auth.json"))
    }
    #[cfg(target_os = "windows")]
    {
        dirs::config_dir().map(|c| c.join("visigrid/auth.json"))
    }
    #[cfg(target_os = "linux")]
    {
        dirs::config_dir().map(|c| c.join("visigrid/auth.json"))
    }
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
pub fn save_auth(creds: &AuthCredentials) -> Result<(), String> {
    let path = auth_file_path().ok_or("Could not determine config directory")?;

    // Create parent directory if needed
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("Failed to create config directory: {}", e))?;
    }

    let contents = serde_json::to_string_pretty(creds)
        .map_err(|e| format!("Failed to serialize credentials: {}", e))?;

    std::fs::write(&path, contents)
        .map_err(|e| format!("Failed to write auth file: {}", e))?;

    // Set restrictive permissions on Unix
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

/// Check if user is authenticated (has saved credentials).
pub fn is_authenticated() -> bool {
    load_auth().is_some()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_auth_credentials_new() {
        let creds = AuthCredentials::new(
            "test-token".to_string(),
            "https://api.visiapi.com".to_string(),
        );
        assert_eq!(creds.token, "test-token");
        assert_eq!(creds.api_base, "https://api.visiapi.com");
        assert!(creds.user_slug.is_none());
        assert!(creds.email.is_none());
    }

    #[test]
    fn test_auth_file_path() {
        let path = auth_file_path();
        assert!(path.is_some());
        let path = path.unwrap();
        assert!(path.to_string_lossy().contains("visigrid"));
        assert!(path.to_string_lossy().contains("auth.json"));
    }
}
