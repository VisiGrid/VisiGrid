//! CLI Exit Code Registry
//!
//! This is the single source of truth for all CLI exit codes.
//! Exit codes are part of the shell contract — scripts rely on them.
//!
//! # Exit Code Ranges
//!
//! | Range   | Domain           | Description                              |
//! |---------|------------------|------------------------------------------|
//! | 0       | Universal        | Success                                  |
//! | 1       | Universal        | General error (unspecified)              |
//! | 2       | Universal        | CLI usage error (bad args, missing file) |
//! | 3-9     | diff             | Reconciliation-specific codes            |
//! | 10-19   | ai               | AI provider/keychain codes               |
//! | 20-29   | session          | Session server codes                     |
//! | 30-39   | replay           | Provenance replay codes                  |
//!
//! # Adding New Exit Codes
//!
//! 1. Add the constant in the appropriate range
//! 2. Document what triggers it
//! 3. Update the table above
//! 4. Wire it into the relevant command's error handling

// =============================================================================
// Universal (0-2)
// =============================================================================

/// Success - command completed without errors.
pub const EXIT_SUCCESS: u8 = 0;

/// General error - unspecified failure.
/// Avoid using this; prefer a specific error code.
pub const EXIT_ERROR: u8 = 1;

/// Usage error - bad arguments, missing required options.
pub const EXIT_USAGE: u8 = 2;

// =============================================================================
// Diff (3-9) — per cli-diff.md spec
// =============================================================================

/// Diff found differences (outside tolerance).
/// Like `diff(1)`, exit 1 means "files differ."
pub const EXIT_DIFF_DIFFS: u8 = 1;

/// Duplicate keys found in input.
pub const EXIT_DIFF_DUPLICATE: u8 = 3;

/// Ambiguous match (multiple candidates for a key).
pub const EXIT_DIFF_AMBIGUOUS: u8 = 4;

/// Parse error reading input files.
pub const EXIT_DIFF_PARSE: u8 = 5;

// =============================================================================
// AI (10-19)
// =============================================================================

/// AI disabled (provider=none) — not an error, just informational.
pub const EXIT_AI_DISABLED: u8 = 10;

/// AI provider configured but API key missing.
pub const EXIT_AI_MISSING_KEY: u8 = 11;

/// Keychain error (cannot read/write credentials).
pub const EXIT_AI_KEYCHAIN_ERR: u8 = 12;

// =============================================================================
// Session (20-29)
// =============================================================================

/// Cannot connect to session server (no server, connection refused).
pub const EXIT_SESSION_CONNECT: u8 = 20;

/// Protocol error (bad framing, version mismatch, malformed message).
pub const EXIT_SESSION_PROTOCOL: u8 = 21;

/// Authentication failed (invalid or missing token).
pub const EXIT_SESSION_AUTH: u8 = 22;

/// Write conflict (another writer holds the lease) or revision mismatch.
pub const EXIT_SESSION_CONFLICT: u8 = 23;

/// Partial apply (non-atomic operation had some rejections).
pub const EXIT_SESSION_PARTIAL: u8 = 24;

/// Invalid input (bad op schema, invalid cell reference).
pub const EXIT_SESSION_INPUT: u8 = 25;

/// Operation timed out.
pub const EXIT_SESSION_TIMEOUT: u8 = 26;

// =============================================================================
// Replay (30-39)
// =============================================================================

/// Fingerprint verification failed.
pub const EXIT_REPLAY_VERIFY_FAILED: u8 = 30;

/// Script execution error (Lua runtime error).
pub const EXIT_REPLAY_SCRIPT_ERROR: u8 = 31;

/// Nondeterministic operation detected (NOW(), RAND(), etc.).
pub const EXIT_REPLAY_NONDETERMINISTIC: u8 = 32;

// =============================================================================
// Session Error Types
// =============================================================================

use crate::session::SessionError;

/// Map a SessionError to its exit code.
pub fn session_exit_code(err: &SessionError) -> u8 {
    match err {
        SessionError::ConnectionFailed(_) => EXIT_SESSION_CONNECT,
        SessionError::ConnectionClosed => EXIT_SESSION_CONNECT,
        SessionError::AuthFailed(_) => EXIT_SESSION_AUTH,
        SessionError::IoError(_) => EXIT_SESSION_CONNECT, // Network I/O issues
        SessionError::ProtocolError(_) => EXIT_SESSION_PROTOCOL,
        SessionError::ServerError { code, .. } => {
            match code.as_str() {
                "auth_failed" => EXIT_SESSION_AUTH,
                "protocol_mismatch" => EXIT_SESSION_PROTOCOL,
                "writer_conflict" | "revision_mismatch" => EXIT_SESSION_CONFLICT,
                "rate_limited" => EXIT_SESSION_CONFLICT, // Treat as temporary conflict
                "formula_parse_error" | "invalid_reference" => EXIT_SESSION_INPUT,
                "malformed_message" | "message_too_large" => EXIT_SESSION_PROTOCOL,
                _ => EXIT_ERROR, // Unknown server error
            }
        }
    }
}

/// Structured error output for session commands.
/// Designed for both human-readable and machine-parseable output.
#[derive(Debug, serde::Serialize)]
pub struct SessionErrorOutput {
    pub error: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub retry_after_ms: Option<u64>,
    pub exit_code: u8,
}

impl SessionErrorOutput {
    pub fn from_session_error(err: &SessionError) -> Self {
        let (error, message, retry_after_ms) = match err {
            SessionError::ConnectionFailed(msg) => {
                ("connect_failed".to_string(), msg.clone(), None)
            }
            SessionError::ConnectionClosed => {
                ("connection_closed".to_string(), "Connection closed by server".to_string(), None)
            }
            SessionError::AuthFailed(msg) => {
                ("auth_failed".to_string(), msg.clone(), None)
            }
            SessionError::IoError(msg) => {
                ("io_error".to_string(), msg.clone(), None)
            }
            SessionError::ProtocolError(msg) => {
                ("protocol_error".to_string(), msg.clone(), None)
            }
            SessionError::ServerError { code, message, retry_after_ms } => {
                (code.clone(), message.clone(), *retry_after_ms)
            }
        };

        Self {
            error,
            message,
            retry_after_ms,
            exit_code: session_exit_code(err),
        }
    }

    /// Print error to stderr (human-readable by default).
    pub fn print(&self, json: bool) {
        if json {
            if let Ok(output) = serde_json::to_string(self) {
                eprintln!("{}", output);
            }
        } else {
            if let Some(retry) = self.retry_after_ms {
                eprintln!("error: {} (retry after {}ms)", self.message, retry);
            } else {
                eprintln!("error: {}", self.message);
            }
        }
    }
}
