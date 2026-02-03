//! Discovery file management for session server.
//!
//! Discovery files allow clients to find running VisiGrid sessions.
//! Each running GUI writes a JSON file containing connection info.
//!
//! Platform paths:
//! - Linux: $XDG_STATE_HOME/visigrid/sessions/<id>.json
//! - macOS: ~/Library/Application Support/VisiGrid/sessions/<id>.json
//! - Windows: %LOCALAPPDATA%\VisiGrid\sessions\<id>.json

use std::fs;
use std::io::Write;
use std::path::PathBuf;

use base64::Engine;
use chrono::{DateTime, Utc};
use rand::RngCore;
use serde::{Deserialize, Serialize};
use uuid::Uuid;

/// Discovery file contents - written atomically, read by clients.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiscoveryFile {
    /// Unique session ID (random UUID).
    pub session_id: Uuid,
    /// TCP port the server is listening on.
    pub port: u16,
    /// Process ID of the GUI.
    pub pid: u32,
    /// Authentication token (base64-encoded 32 bytes).
    /// Clients must include this in the hello message.
    pub token: String,
    /// Path to the open workbook (if any).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_path: Option<PathBuf>,
    /// Display title of the workbook.
    pub workbook_title: String,
    /// When the session was created.
    pub created_at: DateTime<Utc>,
    /// Protocol version supported by this session.
    pub protocol_version: u32,
}

impl DiscoveryFile {
    /// Create a new discovery file with a fresh session ID and token.
    pub fn new(port: u16, workbook_path: Option<PathBuf>, workbook_title: String) -> Self {
        Self {
            session_id: Uuid::new_v4(),
            port,
            pid: std::process::id(),
            token: generate_token(),
            workbook_path,
            workbook_title,
            created_at: Utc::now(),
            protocol_version: 1,
        }
    }

    /// Get the raw token bytes (decode from base64).
    pub fn token_bytes(&self) -> Option<Vec<u8>> {
        base64::engine::general_purpose::STANDARD
            .decode(&self.token)
            .ok()
    }

    /// Verify a token string matches this session's token.
    pub fn verify_token(&self, token: &str) -> bool {
        // Constant-time comparison to prevent timing attacks
        use subtle::ConstantTimeEq;
        if let (Some(expected), Ok(provided)) = (
            self.token_bytes(),
            base64::engine::general_purpose::STANDARD.decode(token),
        ) {
            expected.ct_eq(&provided).into()
        } else {
            false
        }
    }
}

/// Generate a cryptographically random 32-byte token, base64-encoded.
fn generate_token() -> String {
    let mut bytes = [0u8; 32];
    rand::thread_rng().fill_bytes(&mut bytes);
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

/// Manages discovery file lifecycle.
pub struct DiscoveryManager {
    /// Path to this session's discovery file.
    path: PathBuf,
    /// The discovery file contents.
    discovery: DiscoveryFile,
}

impl DiscoveryManager {
    /// Create a new discovery manager and write the discovery file.
    pub fn new(
        port: u16,
        workbook_path: Option<PathBuf>,
        workbook_title: String,
    ) -> std::io::Result<Self> {
        let discovery = DiscoveryFile::new(port, workbook_path, workbook_title);
        let path = discovery_file_path(&discovery.session_id)?;

        let manager = Self { path, discovery };
        manager.write()?;

        Ok(manager)
    }

    /// Get the discovery file contents.
    pub fn discovery(&self) -> &DiscoveryFile {
        &self.discovery
    }

    /// Get the session ID.
    pub fn session_id(&self) -> Uuid {
        self.discovery.session_id
    }

    /// Get the authentication token.
    pub fn token(&self) -> &str {
        &self.discovery.token
    }

    /// Verify a token matches this session.
    pub fn verify_token(&self, token: &str) -> bool {
        self.discovery.verify_token(token)
    }

    /// Update the workbook info and rewrite the discovery file.
    pub fn update_workbook(
        &mut self,
        path: Option<PathBuf>,
        title: String,
    ) -> std::io::Result<()> {
        self.discovery.workbook_path = path;
        self.discovery.workbook_title = title;
        self.write()
    }

    /// Write the discovery file atomically (write to temp, then rename).
    fn write(&self) -> std::io::Result<()> {
        // Ensure parent directory exists
        if let Some(parent) = self.path.parent() {
            fs::create_dir_all(parent)?;
        }

        // Write to temp file
        let temp_path = self.path.with_extension("json.tmp");
        {
            let mut file = fs::File::create(&temp_path)?;
            let json = serde_json::to_string_pretty(&self.discovery)
                .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
            file.write_all(json.as_bytes())?;
            file.sync_all()?;
        }

        // Atomic rename
        fs::rename(&temp_path, &self.path)?;

        Ok(())
    }

    /// Remove the discovery file (called on shutdown).
    pub fn cleanup(&self) -> std::io::Result<()> {
        if self.path.exists() {
            fs::remove_file(&self.path)?;
        }
        Ok(())
    }
}

impl Drop for DiscoveryManager {
    fn drop(&mut self) {
        // Best-effort cleanup on drop
        let _ = self.cleanup();
    }
}

/// Get the platform-specific directory for discovery files.
pub fn discovery_dir() -> std::io::Result<PathBuf> {
    #[cfg(target_os = "linux")]
    {
        // $XDG_STATE_HOME/visigrid/sessions or ~/.local/state/visigrid/sessions
        let base = std::env::var("XDG_STATE_HOME")
            .map(PathBuf::from)
            .unwrap_or_else(|_| {
                dirs::home_dir()
                    .unwrap_or_else(|| PathBuf::from("/tmp"))
                    .join(".local/state")
            });
        Ok(base.join("visigrid/sessions"))
    }

    #[cfg(target_os = "macos")]
    {
        // ~/Library/Application Support/VisiGrid/sessions
        let base = dirs::data_dir().unwrap_or_else(|| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("/tmp"))
                .join("Library/Application Support")
        });
        Ok(base.join("VisiGrid/sessions"))
    }

    #[cfg(target_os = "windows")]
    {
        // %LOCALAPPDATA%\VisiGrid\sessions
        let base = dirs::data_local_dir().unwrap_or_else(|| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("C:\\"))
                .join("AppData\\Local")
        });
        Ok(base.join("VisiGrid\\sessions"))
    }

    #[cfg(not(any(target_os = "linux", target_os = "macos", target_os = "windows")))]
    {
        // Fallback: temp directory
        Ok(std::env::temp_dir().join("visigrid/sessions"))
    }
}

/// Get the path for a specific session's discovery file.
fn discovery_file_path(session_id: &Uuid) -> std::io::Result<PathBuf> {
    let dir = discovery_dir()?;
    Ok(dir.join(format!("{}.json", session_id)))
}

/// List all discovery files (for CLI `sessions` command).
pub fn list_sessions() -> std::io::Result<Vec<DiscoveryFile>> {
    let dir = discovery_dir()?;
    if !dir.exists() {
        return Ok(Vec::new());
    }

    let mut sessions = Vec::new();
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.extension().map_or(false, |ext| ext == "json") {
            if let Ok(contents) = fs::read_to_string(&path) {
                if let Ok(discovery) = serde_json::from_str::<DiscoveryFile>(&contents) {
                    // Check if the process is still running
                    if is_process_alive(discovery.pid) {
                        sessions.push(discovery);
                    } else {
                        // Stale discovery file - clean it up
                        let _ = fs::remove_file(&path);
                    }
                }
            }
        }
    }

    Ok(sessions)
}

/// Check if a process is still running.
fn is_process_alive(pid: u32) -> bool {
    #[cfg(unix)]
    {
        // On Unix, kill with signal 0 checks if process exists
        unsafe { libc::kill(pid as i32, 0) == 0 }
    }

    #[cfg(windows)]
    {
        use windows_sys::Win32::Foundation::{CloseHandle, HANDLE};
        use windows_sys::Win32::System::Threading::{OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION};

        unsafe {
            let handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid);
            if handle != 0 {
                CloseHandle(handle);
                true
            } else {
                false
            }
        }
    }

    #[cfg(not(any(unix, windows)))]
    {
        // Assume alive if we can't check
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_token_generation() {
        let token = generate_token();
        // Base64 of 32 bytes = 44 characters (with padding)
        assert_eq!(token.len(), 44);

        // Should decode back to 32 bytes
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(&token)
            .unwrap();
        assert_eq!(bytes.len(), 32);
    }

    #[test]
    fn test_discovery_file_serialization() {
        let discovery = DiscoveryFile::new(12345, Some(PathBuf::from("/tmp/test.sheet")), "Test".to_string());

        let json = serde_json::to_string(&discovery).unwrap();
        let parsed: DiscoveryFile = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.session_id, discovery.session_id);
        assert_eq!(parsed.port, 12345);
        assert_eq!(parsed.token, discovery.token);
    }

    #[test]
    fn test_token_verification() {
        let discovery = DiscoveryFile::new(12345, None, "Test".to_string());

        // Correct token should verify
        assert!(discovery.verify_token(&discovery.token));

        // Wrong token should not verify
        assert!(!discovery.verify_token("wrong_token"));

        // Different valid token should not verify
        let other_token = generate_token();
        assert!(!discovery.verify_token(&other_token));
    }
}
