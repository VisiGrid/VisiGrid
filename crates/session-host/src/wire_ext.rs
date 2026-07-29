//! Server-side wire helpers that are NOT part of the shared frozen wire
//! contract: internal cell refs for broadcast coalescing and the error-code
//! taxonomy. Moved from the GUI's protocol mirror 2026-07-29; the wire types
//! themselves now come from visigrid-protocol (mirror deleted).

use serde::{Deserialize, Serialize};

/// Maximum accepted line length on the wire (server-side enforcement).
pub const MAX_MESSAGE_SIZE: usize = 10 * 1024 * 1024;
use visigrid_protocol::ErrorMessage;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CellRef {
    pub sheet: usize,
    pub row: usize,
    pub col: usize,
}

/// Protocol error codes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtocolError {
    /// Invalid or missing token.
    AuthFailed,
    /// Unsupported protocol version.
    ProtocolMismatch,
    /// Rate limit exceeded.
    RateLimited,
    /// expected_revision doesn't match current.
    RevisionMismatch,
    /// Formula syntax error.
    FormulaParseError,
    /// Invalid sheet/row/col reference.
    InvalidReference,
    /// Message too large.
    MessageTooLarge,
    /// Malformed JSON.
    MalformedMessage,
    /// Server is in read-only mode.
    ReadOnlyMode,
    /// Writer lease held by another connection.
    WriterConflict,
    /// Unknown error.
    InternalError,
}

/// All protocol error variants. Used for exhaustive testing.
/// IMPORTANT: Update this array when adding/removing error codes.
pub const ALL_ERROR_CODES: &[ProtocolError] = &[
    ProtocolError::AuthFailed,
    ProtocolError::ProtocolMismatch,
    ProtocolError::RateLimited,
    ProtocolError::RevisionMismatch,
    ProtocolError::FormulaParseError,
    ProtocolError::InvalidReference,
    ProtocolError::MessageTooLarge,
    ProtocolError::MalformedMessage,
    ProtocolError::ReadOnlyMode,
    ProtocolError::WriterConflict,
    ProtocolError::InternalError,
];

impl ProtocolError {
    pub fn code(&self) -> &'static str {
        match self {
            Self::AuthFailed => "auth_failed",
            Self::ProtocolMismatch => "protocol_mismatch",
            Self::RateLimited => "rate_limited",
            Self::RevisionMismatch => "revision_mismatch",
            Self::FormulaParseError => "formula_parse_error",
            Self::InvalidReference => "invalid_reference",
            Self::MessageTooLarge => "message_too_large",
            Self::MalformedMessage => "malformed_message",
            Self::ReadOnlyMode => "read_only_mode",
            Self::WriterConflict => "writer_conflict",
            Self::InternalError => "internal_error",
        }
    }

    pub fn message(&self) -> &'static str {
        match self {
            Self::AuthFailed => "Invalid or missing authentication token",
            Self::ProtocolMismatch => "Unsupported protocol version",
            Self::RateLimited => "Rate limit exceeded",
            Self::RevisionMismatch => "Expected revision does not match current revision",
            Self::FormulaParseError => "Formula syntax error",
            Self::InvalidReference => "Invalid sheet, row, or column reference",
            Self::MessageTooLarge => "Message exceeds maximum size",
            Self::MalformedMessage => "Malformed JSON message",
            Self::ReadOnlyMode => "Server is in read-only mode",
            Self::WriterConflict => "Write lease held by another connection",
            Self::InternalError => "Internal server error",
        }
    }

    pub fn to_error_message(&self, id: Option<String>) -> ErrorMessage {
        ErrorMessage {
            id,
            code: self.code().to_string(),
            message: self.message().to_string(),
            retry_after_ms: None,
        }
    }

    /// Create a rate limited error message with retry information.
    pub fn rate_limited_error(id: Option<String>, retry_after_ms: u64) -> ErrorMessage {
        ErrorMessage {
            id,
            code: Self::RateLimited.code().to_string(),
            message: format!("Rate limit exceeded. Retry after {} ms", retry_after_ms),
            retry_after_ms: Some(retry_after_ms),
        }
    }
}

