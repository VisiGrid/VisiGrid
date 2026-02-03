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
///
/// SECURITY: This file is readable by any process on the system.
/// It intentionally does NOT contain the authentication token.
/// Tokens are distributed out-of-band (env var, secure channel, etc.).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiscoveryFile {
    /// Unique session ID (random UUID).
    pub session_id: Uuid,
    /// TCP port the server is listening on.
    pub port: u16,
    /// Process ID of the GUI.
    pub pid: u32,
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
    /// Create a new discovery file with a fresh session ID.
    pub fn new(port: u16, workbook_path: Option<PathBuf>, workbook_title: String) -> Self {
        Self {
            session_id: Uuid::new_v4(),
            port,
            pid: std::process::id(),
            workbook_path,
            workbook_title,
            created_at: Utc::now(),
            protocol_version: 1,
        }
    }
}

/// Generate cryptographically random 32 bytes for authentication token.
fn generate_token_bytes() -> [u8; 32] {
    let mut bytes = [0u8; 32];
    rand::thread_rng().fill_bytes(&mut bytes);
    bytes
}

/// Manages discovery file lifecycle and authentication token.
///
/// The token is stored in memory only - never written to disk.
/// For CI/automation, pass the token via VISIGRID_SESSION_TOKEN env var.
pub struct DiscoveryManager {
    /// Path to this session's discovery file.
    path: PathBuf,
    /// The discovery file contents (public info only).
    discovery: DiscoveryFile,
    /// Authentication token (base64-encoded, kept in memory only).
    token: String,
    /// Raw token bytes for verification.
    token_bytes: Vec<u8>,
}

impl DiscoveryManager {
    /// Create a new discovery manager and write the discovery file.
    ///
    /// If `token_override` is Some, uses that token (for test harness).
    /// Otherwise generates a fresh cryptographic token.
    pub fn new(
        port: u16,
        workbook_path: Option<PathBuf>,
        workbook_title: String,
        token_override: Option<String>,
    ) -> std::io::Result<Self> {
        let discovery = DiscoveryFile::new(port, workbook_path, workbook_title);
        let path = discovery_file_path(&discovery.session_id)?;

        // Use provided token or generate a fresh one
        let (token, token_bytes) = if let Some(t) = token_override {
            let bytes = base64::engine::general_purpose::STANDARD
                .decode(&t)
                .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
            (t, bytes)
        } else {
            let bytes = generate_token_bytes();
            let encoded = base64::engine::general_purpose::STANDARD.encode(&bytes);
            (encoded, bytes.to_vec())
        };

        let manager = Self { path, discovery, token, token_bytes };
        manager.write()?;

        Ok(manager)
    }

    /// Get the discovery file contents.
    pub fn discovery(&self) -> &DiscoveryFile {
        &self.discovery
    }

    /// Get the path to the discovery file.
    pub fn path(&self) -> &PathBuf {
        &self.path
    }

    /// Get the session ID.
    pub fn session_id(&self) -> Uuid {
        self.discovery.session_id
    }

    /// Get the authentication token (base64-encoded).
    pub fn token(&self) -> &str {
        &self.token
    }

    /// Verify a token matches this session (constant-time comparison).
    pub fn verify_token(&self, token: &str) -> bool {
        use subtle::ConstantTimeEq;
        if let Ok(provided) = base64::engine::general_purpose::STANDARD.decode(token) {
            self.token_bytes.ct_eq(&provided).into()
        } else {
            false
        }
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
        let bytes = generate_token_bytes();
        assert_eq!(bytes.len(), 32);

        // Base64 encoding should produce 44 characters
        let encoded = base64::engine::general_purpose::STANDARD.encode(bytes);
        assert_eq!(encoded.len(), 44);
    }

    #[test]
    fn test_discovery_file_serialization() {
        let discovery = DiscoveryFile::new(12345, Some(PathBuf::from("/tmp/test.sheet")), "Test".to_string());

        let json = serde_json::to_string(&discovery).unwrap();
        let parsed: DiscoveryFile = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.session_id, discovery.session_id);
        assert_eq!(parsed.port, 12345);

        // Verify token is NOT in serialized output
        assert!(!json.contains("token"));
    }

    #[test]
    fn test_discovery_file_no_token_field() {
        // Verify the discovery file struct has no token field
        let discovery = DiscoveryFile::new(12345, None, "Test".to_string());
        let json = serde_json::to_string_pretty(&discovery).unwrap();

        // Token should never appear in the JSON
        assert!(!json.contains("token"), "Discovery file should not contain token: {}", json);
    }
}
