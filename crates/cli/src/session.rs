//! Session discovery and client for connecting to running VisiGrid instances.
//!
//! This module provides CLI commands for interacting with VisiGrid GUI sessions:
//! - `sessions` - List running sessions
//! - `attach` - Connect and show session info
//! - `apply` - Apply operations to a session
//! - `inspect` - Query cell state

use std::fs;
use std::io::{BufReader, BufWriter, Read, Write};
use std::net::TcpStream;
use std::path::PathBuf;
use std::time::Duration;

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use uuid::Uuid;

// Re-export protocol types from the shared crate
pub use visigrid_protocol::{
    // Client messages
    ClientMessage, HelloMessage, ApplyOpsMessage, InspectMessage, PingMessage,
    StatsMessage,
    // Server messages
    ServerMessage, ApplyOpsResultMessage, InspectResultMessage, StatsResultMessage,
    // Shared types
    Op, InspectTarget,
    // Constants
    PROTOCOL_VERSION,
};

// ============================================================================
// Discovery (mirrors gpui-app/src/session_server/discovery.rs)
// ============================================================================

/// Discovery file contents - read from disk to find running sessions.
///
/// Note: This uses chrono::DateTime<Utc> for richer CLI handling,
/// while the wire format uses ISO 8601 strings.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiscoveryFile {
    pub session_id: Uuid,
    pub port: u16,
    pub pid: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_path: Option<PathBuf>,
    pub workbook_title: String,
    pub created_at: DateTime<Utc>,
    pub protocol_version: u32,
}

/// Get the platform-specific directory for discovery files.
pub fn discovery_dir() -> std::io::Result<PathBuf> {
    #[cfg(target_os = "linux")]
    {
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
        let base = dirs::data_dir().unwrap_or_else(|| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("/tmp"))
                .join("Library/Application Support")
        });
        Ok(base.join("VisiGrid/sessions"))
    }

    #[cfg(target_os = "windows")]
    {
        let base = dirs::data_local_dir().unwrap_or_else(|| {
            dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("C:\\"))
                .join("AppData\\Local")
        });
        Ok(base.join("VisiGrid\\sessions"))
    }

    #[cfg(not(any(target_os = "linux", target_os = "macos", target_os = "windows")))]
    {
        Ok(std::env::temp_dir().join("visigrid/sessions"))
    }
}

/// Check if a process is still running.
fn is_process_alive(pid: u32) -> bool {
    #[cfg(unix)]
    {
        unsafe { libc::kill(pid as i32, 0) == 0 }
    }

    #[cfg(windows)]
    {
        use windows_sys::Win32::Foundation::CloseHandle;
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
        true
    }
}

/// List all discovery files, cleaning up stale ones.
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

    // Sort by created_at descending (most recent first)
    sessions.sort_by(|a, b| b.created_at.cmp(&a.created_at));

    Ok(sessions)
}

/// Find a session by ID prefix (supports partial match).
pub fn find_session(id_prefix: &str) -> std::io::Result<Option<DiscoveryFile>> {
    let sessions = list_sessions()?;

    // Try exact match first
    if let Some(session) = sessions.iter().find(|s| s.session_id.to_string() == id_prefix) {
        return Ok(Some(session.clone()));
    }

    // Try prefix match
    let matches: Vec<_> = sessions
        .iter()
        .filter(|s| s.session_id.to_string().starts_with(id_prefix))
        .collect();

    match matches.len() {
        0 => Ok(None),
        1 => Ok(Some(matches[0].clone())),
        _ => Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            format!("Ambiguous session ID '{}' matches {} sessions", id_prefix, matches.len()),
        )),
    }
}

// ============================================================================
// Session Client
// ============================================================================

/// A client connection to a VisiGrid session.
pub struct SessionClient {
    reader: BufReader<TcpStream>,
    writer: BufWriter<TcpStream>,
    session_id: String,
    revision: u64,
    capabilities: Vec<String>,
    next_id: u64,
}

impl SessionClient {
    /// Connect to a session and perform the hello handshake.
    pub fn connect(discovery: &DiscoveryFile, token: &str) -> Result<Self, SessionError> {
        let addr = format!("127.0.0.1:{}", discovery.port);
        let stream = TcpStream::connect_timeout(
            &addr.parse().map_err(|_| SessionError::ConnectionFailed("Invalid address".into()))?,
            Duration::from_secs(5),
        ).map_err(|e| SessionError::ConnectionFailed(e.to_string()))?;

        stream.set_read_timeout(Some(Duration::from_secs(30)))
            .map_err(|e| SessionError::ConnectionFailed(e.to_string()))?;
        stream.set_write_timeout(Some(Duration::from_secs(30)))
            .map_err(|e| SessionError::ConnectionFailed(e.to_string()))?;

        let reader = BufReader::new(stream.try_clone()
            .map_err(|e| SessionError::ConnectionFailed(e.to_string()))?);
        let writer = BufWriter::new(stream);

        let mut client = Self {
            reader,
            writer,
            session_id: String::new(),
            revision: 0,
            capabilities: Vec::new(),
            next_id: 1,
        };

        // Send hello
        let hello = ClientMessage::Hello(HelloMessage {
            id: client.next_request_id(),
            client: "visigrid-cli".to_string(),
            version: env!("CARGO_PKG_VERSION").to_string(),
            token: token.to_string(),
            protocol_version: PROTOCOL_VERSION,
        });
        client.send(&hello)?;

        // Read response
        let response = client.receive()?;
        match response {
            ServerMessage::Welcome(welcome) => {
                client.session_id = welcome.session_id;
                client.revision = welcome.revision;
                client.capabilities = welcome.capabilities;
                Ok(client)
            }
            ServerMessage::Error(err) => {
                Err(SessionError::AuthFailed(err.message))
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to hello".into())),
        }
    }

    /// Get the current revision.
    pub fn revision(&self) -> u64 {
        self.revision
    }

    /// Get the session ID.
    pub fn session_id(&self) -> &str {
        &self.session_id
    }

    /// Get server capabilities.
    pub fn capabilities(&self) -> &[String] {
        &self.capabilities
    }

    /// Apply operations to the session.
    pub fn apply_ops(
        &mut self,
        ops: Vec<Op>,
        atomic: bool,
        expected_revision: Option<u64>,
    ) -> Result<ApplyOpsResultMessage, SessionError> {
        let msg = ClientMessage::ApplyOps(ApplyOpsMessage {
            id: self.next_request_id(),
            ops,
            atomic,
            expected_revision,
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::ApplyOpsResult(result) => {
                self.revision = result.revision;
                Ok(result)
            }
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to apply_ops".into())),
        }
    }

    /// Inspect a single cell.
    pub fn inspect_cell(&mut self, sheet: usize, row: usize, col: usize) -> Result<InspectResultMessage, SessionError> {
        let msg = ClientMessage::Inspect(InspectMessage {
            id: self.next_request_id(),
            target: InspectTarget::Cell { sheet, row, col },
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::InspectResult(result) => {
                self.revision = result.revision;
                Ok(result)
            }
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to inspect".into())),
        }
    }

    /// Inspect a range of cells.
    pub fn inspect_range(
        &mut self,
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<InspectResultMessage, SessionError> {
        let msg = ClientMessage::Inspect(InspectMessage {
            id: self.next_request_id(),
            target: InspectTarget::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
            },
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::InspectResult(result) => {
                self.revision = result.revision;
                Ok(result)
            }
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to inspect".into())),
        }
    }

    /// Inspect workbook metadata.
    pub fn inspect_workbook(&mut self) -> Result<InspectResultMessage, SessionError> {
        let msg = ClientMessage::Inspect(InspectMessage {
            id: self.next_request_id(),
            target: InspectTarget::Workbook,
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::InspectResult(result) => {
                self.revision = result.revision;
                Ok(result)
            }
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to inspect".into())),
        }
    }

    /// Ping the server.
    pub fn ping(&mut self) -> Result<(), SessionError> {
        let msg = ClientMessage::Ping(PingMessage {
            id: self.next_request_id(),
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::Pong(_) => Ok(()),
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to ping".into())),
        }
    }

    /// Get server statistics.
    pub fn stats(&mut self) -> Result<StatsResultMessage, SessionError> {
        let msg = ClientMessage::Stats(StatsMessage {
            id: self.next_request_id(),
        });
        self.send(&msg)?;

        let response = self.receive()?;
        match response {
            ServerMessage::StatsResult(stats) => Ok(stats),
            ServerMessage::Error(err) => {
                Err(SessionError::ServerError { code: err.code, message: err.message, retry_after_ms: err.retry_after_ms })
            }
            _ => Err(SessionError::ProtocolError("Unexpected response to stats".into())),
        }
    }

    fn next_request_id(&mut self) -> String {
        let id = self.next_id;
        self.next_id += 1;
        id.to_string()
    }

    fn send(&mut self, msg: &ClientMessage) -> Result<(), SessionError> {
        let json = serde_json::to_string(msg)
            .map_err(|e| SessionError::ProtocolError(e.to_string()))?;
        writeln!(self.writer, "{}", json)
            .map_err(|e| SessionError::IoError(e.to_string()))?;
        self.writer.flush()
            .map_err(|e| SessionError::IoError(e.to_string()))?;
        Ok(())
    }

    /// Maximum line size (10MB). Protects against memory exhaustion from malformed/hostile messages.
    const MAX_LINE_BYTES: usize = 10 * 1024 * 1024;

    fn receive(&mut self) -> Result<ServerMessage, SessionError> {
        let line = self.receive_line_bounded()?;

        serde_json::from_str(&line)
            .map_err(|e| SessionError::ProtocolError(format!("Invalid JSON: {}", e)))
    }

    /// Read a line with bounded size to prevent memory exhaustion.
    fn receive_line_bounded(&mut self) -> Result<String, SessionError> {
        let mut buf = Vec::with_capacity(4096);

        loop {
            let mut byte = [0u8; 1];
            match self.reader.read(&mut byte) {
                Ok(0) => {
                    // Connection closed
                    if buf.is_empty() {
                        return Err(SessionError::ConnectionClosed);
                    } else {
                        // Connection closed mid-frame
                        return Err(SessionError::ProtocolError(
                            "connection closed mid-frame (no newline)".to_string()
                        ));
                    }
                }
                Ok(_) => {
                    if byte[0] == b'\n' {
                        // End of line
                        break;
                    }
                    buf.push(byte[0]);

                    if buf.len() > Self::MAX_LINE_BYTES {
                        return Err(SessionError::ProtocolError(format!(
                            "message exceeds {}MB limit",
                            Self::MAX_LINE_BYTES / (1024 * 1024)
                        )));
                    }
                }
                Err(e) if e.kind() == std::io::ErrorKind::Interrupted => {
                    // Retry on interrupt
                    continue;
                }
                Err(e) => {
                    return Err(SessionError::IoError(e.to_string()));
                }
            }
        }

        String::from_utf8(buf)
            .map_err(|e| SessionError::ProtocolError(format!("Invalid UTF-8: {}", e)))
    }
}

/// Errors that can occur when interacting with a session.
#[derive(Debug)]
pub enum SessionError {
    ConnectionFailed(String),
    ConnectionClosed,
    AuthFailed(String),
    IoError(String),
    ProtocolError(String),
    ServerError {
        code: String,
        message: String,
        /// Retry hint from server (e.g., for writer_conflict, rate_limited).
        retry_after_ms: Option<u64>,
    },
}

impl std::fmt::Display for SessionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            SessionError::ConnectionFailed(msg) => write!(f, "Connection failed: {}", msg),
            SessionError::ConnectionClosed => write!(f, "Connection closed by server"),
            SessionError::AuthFailed(msg) => write!(f, "Authentication failed: {}", msg),
            SessionError::IoError(msg) => write!(f, "I/O error: {}", msg),
            SessionError::ProtocolError(msg) => write!(f, "Protocol error: {}", msg),
            SessionError::ServerError { code, message, .. } => write!(f, "Server error [{}]: {}", code, message),
        }
    }
}

impl std::error::Error for SessionError {}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    /// Helper to test bounded line reading with a fake reader.
    fn read_line_bounded_from_bytes(data: &[u8], max_bytes: usize) -> Result<String, SessionError> {
        let mut reader = Cursor::new(data.to_vec());
        let mut buf = Vec::with_capacity(4096);

        loop {
            let mut byte = [0u8; 1];
            match reader.read(&mut byte) {
                Ok(0) => {
                    if buf.is_empty() {
                        return Err(SessionError::ConnectionClosed);
                    } else {
                        return Err(SessionError::ProtocolError(
                            "connection closed mid-frame (no newline)".to_string()
                        ));
                    }
                }
                Ok(_) => {
                    if byte[0] == b'\n' {
                        break;
                    }
                    buf.push(byte[0]);

                    if buf.len() > max_bytes {
                        return Err(SessionError::ProtocolError(format!(
                            "message exceeds {} byte limit",
                            max_bytes
                        )));
                    }
                }
                Err(e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(e) => return Err(SessionError::IoError(e.to_string())),
            }
        }

        String::from_utf8(buf)
            .map_err(|e| SessionError::ProtocolError(format!("Invalid UTF-8: {}", e)))
    }

    #[test]
    fn test_bounded_read_normal() {
        let data = b"{\"type\":\"pong\",\"id\":\"1\"}\n";
        let result = read_line_bounded_from_bytes(data, 1024);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), "{\"type\":\"pong\",\"id\":\"1\"}");
    }

    #[test]
    fn test_bounded_read_oversize_triggers_error() {
        // Create a line that exceeds the limit
        let mut data = vec![b'x'; 100];
        data.push(b'\n');

        // Set limit to 50 bytes
        let result = read_line_bounded_from_bytes(&data, 50);

        assert!(result.is_err());
        let err = result.unwrap_err();
        match err {
            SessionError::ProtocolError(msg) => {
                assert!(msg.contains("exceeds"), "Expected 'exceeds' in error: {}", msg);
            }
            _ => panic!("Expected ProtocolError, got {:?}", err),
        }
    }

    #[test]
    fn test_bounded_read_connection_close_mid_frame() {
        // Data without newline (simulates connection close mid-frame)
        let data = b"{\"type\":\"pong\"";

        let result = read_line_bounded_from_bytes(data, 1024);

        assert!(result.is_err());
        let err = result.unwrap_err();
        match err {
            SessionError::ProtocolError(msg) => {
                assert!(msg.contains("mid-frame"), "Expected 'mid-frame' in error: {}", msg);
            }
            _ => panic!("Expected ProtocolError, got {:?}", err),
        }
    }

    #[test]
    fn test_bounded_read_empty_connection_close() {
        // Empty data (clean connection close)
        let data = b"";

        let result = read_line_bounded_from_bytes(data, 1024);

        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), SessionError::ConnectionClosed));
    }
}
