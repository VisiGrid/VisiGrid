//! Protocol definitions for session server communication.
//!
//! Messages are JSONL (newline-delimited JSON) over TCP.
//! Each message has an `id` for request/response correlation.

use serde::{Deserialize, Serialize};

/// Protocol version. Increment on breaking changes.
pub const PROTOCOL_VERSION: u32 = 1;

/// Maximum message size (10 MB).
pub const MAX_MESSAGE_SIZE: usize = 10 * 1024 * 1024;

/// Messages from client to server.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ClientMessage {
    /// Initial handshake - must be first message.
    Hello(HelloMessage),

    /// Apply a batch of operations.
    ApplyOps(ApplyOpsMessage),

    /// Subscribe to events.
    Subscribe(SubscribeMessage),

    /// Unsubscribe from events.
    Unsubscribe(UnsubscribeMessage),

    /// Query current state (e.g., cell values).
    Inspect(InspectMessage),

    /// Ping for keepalive.
    Ping(PingMessage),
}

/// Messages from server to client.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ServerMessage {
    /// Response to Hello.
    Welcome(WelcomeMessage),

    /// Response to ApplyOps.
    ApplyOpsResult(ApplyOpsResultMessage),

    /// Response to Subscribe.
    Subscribed(SubscribedMessage),

    /// Response to Unsubscribe.
    Unsubscribed(UnsubscribedMessage),

    /// Response to Inspect.
    InspectResult(InspectResultMessage),

    /// Response to Ping.
    Pong(PongMessage),

    /// Push event (cells changed, etc.).
    Event(EventMessage),

    /// Error response.
    Error(ErrorMessage),
}

// ============================================================================
// Hello / Welcome
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HelloMessage {
    /// Request ID for correlation.
    pub id: String,
    /// Client identifier (e.g., "visigrid-cli", "my-agent").
    pub client: String,
    /// Client version.
    pub version: String,
    /// Authentication token (from discovery file).
    pub token: String,
    /// Protocol version the client supports.
    #[serde(default = "default_protocol_version")]
    pub protocol_version: u32,
}

fn default_protocol_version() -> u32 {
    1
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WelcomeMessage {
    /// Echoed request ID.
    pub id: String,
    /// Session ID.
    pub session_id: String,
    /// Protocol version in use (min of client and server).
    pub protocol_version: u32,
    /// Current revision number.
    pub revision: u64,
    /// Server capabilities.
    pub capabilities: Vec<String>,
}

// ============================================================================
// Apply Ops
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyOpsMessage {
    /// Request ID for correlation.
    pub id: String,
    /// Operations to apply.
    pub ops: Vec<Op>,
    /// If true, all-or-nothing. If false, partial apply on error.
    #[serde(default)]
    pub atomic: bool,
    /// Expected revision for optimistic concurrency. If set and doesn't match,
    /// the request is rejected with revision_mismatch.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub expected_revision: Option<u64>,
}

/// A single operation to apply.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum Op {
    /// Set a cell's value (auto-detects formulas).
    SetCellValue {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
        value: String,
    },
    /// Set a cell's formula explicitly.
    SetCellFormula {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
        formula: String,
    },
    /// Clear a cell.
    ClearCell {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
    },
    /// Set number format for a range.
    SetNumberFormat {
        #[serde(default)]
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        format: String,
    },
    /// Set style (bold, italic, etc.) for a range.
    SetStyle {
        #[serde(default)]
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        #[serde(skip_serializing_if = "Option::is_none")]
        bold: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        italic: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        underline: Option<bool>,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyOpsResultMessage {
    /// Echoed request ID.
    pub id: String,
    /// Number of ops successfully applied.
    pub applied: usize,
    /// Total number of ops in the request.
    pub total: usize,
    /// Revision after apply (new on success, unchanged on full rollback).
    pub revision: u64,
    /// Error if any op failed.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<OpError>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OpError {
    /// Error code (e.g., "formula_parse_error", "revision_mismatch").
    pub code: String,
    /// Human-readable message.
    pub message: String,
    /// Index of the failing op (0-based).
    pub op_index: usize,
    /// Optional suggestion for fixing the error.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suggestion: Option<String>,
}

// ============================================================================
// Subscribe / Unsubscribe
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SubscribeMessage {
    /// Request ID for correlation.
    pub id: String,
    /// Topics to subscribe to.
    pub topics: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SubscribedMessage {
    /// Echoed request ID.
    pub id: String,
    /// Topics successfully subscribed.
    pub topics: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsubscribeMessage {
    /// Request ID for correlation.
    pub id: String,
    /// Topics to unsubscribe from.
    pub topics: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsubscribedMessage {
    /// Echoed request ID.
    pub id: String,
    /// Topics successfully unsubscribed.
    pub topics: Vec<String>,
}

// ============================================================================
// Events
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventMessage {
    /// Event topic.
    pub topic: String,
    /// Revision that produced this event.
    pub revision: u64,
    /// Event payload.
    pub payload: EventPayload,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "event", rename_all = "snake_case")]
pub enum EventPayload {
    /// Cells changed (value or formula).
    CellsChanged {
        /// List of changed cells.
        cells: Vec<CellRef>,
    },
    /// Revision changed.
    RevisionChanged {
        /// Previous revision.
        previous: u64,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellRef {
    pub sheet: usize,
    pub row: usize,
    pub col: usize,
}

// ============================================================================
// Inspect
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectMessage {
    /// Request ID for correlation.
    pub id: String,
    /// What to inspect.
    pub target: InspectTarget,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "target", rename_all = "snake_case")]
pub enum InspectTarget {
    /// Inspect a single cell.
    Cell { sheet: usize, row: usize, col: usize },
    /// Inspect a range of cells.
    Range {
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
    /// Inspect workbook metadata.
    Workbook,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectResultMessage {
    /// Echoed request ID.
    pub id: String,
    /// Current revision.
    pub revision: u64,
    /// Inspection result.
    pub result: InspectResult,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum InspectResult {
    /// Single cell info.
    Cell(CellInfo),
    /// Range of cells.
    Range { cells: Vec<CellInfo> },
    /// Workbook metadata.
    Workbook(WorkbookInfo),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellInfo {
    pub sheet: usize,
    pub row: usize,
    pub col: usize,
    /// Display value (formatted).
    pub display: String,
    /// Raw value or formula.
    pub raw: String,
    /// True if cell contains a formula.
    pub is_formula: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookInfo {
    /// Number of sheets.
    pub sheet_count: usize,
    /// Sheet names.
    pub sheets: Vec<String>,
    /// Current revision.
    pub revision: u64,
}

// ============================================================================
// Ping / Pong
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PingMessage {
    /// Request ID for correlation.
    pub id: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PongMessage {
    /// Echoed request ID.
    pub id: String,
}

// ============================================================================
// Error
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ErrorMessage {
    /// Request ID (if available).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    /// Error code.
    pub code: String,
    /// Human-readable message.
    pub message: String,
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
    /// Unknown error.
    InternalError,
}

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
            Self::InternalError => "Internal server error",
        }
    }

    pub fn to_error_message(&self, id: Option<String>) -> ErrorMessage {
        ErrorMessage {
            id,
            code: self.code().to_string(),
            message: self.message().to_string(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_client_message_serialization() {
        let msg = ClientMessage::Hello(HelloMessage {
            id: "1".to_string(),
            client: "test-client".to_string(),
            version: "1.0.0".to_string(),
            token: "abc123".to_string(),
            protocol_version: 1,
        });

        let json = serde_json::to_string(&msg).unwrap();
        assert!(json.contains(r#""type":"hello""#));

        let parsed: ClientMessage = serde_json::from_str(&json).unwrap();
        if let ClientMessage::Hello(h) = parsed {
            assert_eq!(h.client, "test-client");
        } else {
            panic!("Expected Hello message");
        }
    }

    #[test]
    fn test_op_serialization() {
        let op = Op::SetCellValue {
            sheet: 0,
            row: 5,
            col: 3,
            value: "hello".to_string(),
        };

        let json = serde_json::to_string(&op).unwrap();
        assert!(json.contains(r#""op":"set_cell_value""#));

        let parsed: Op = serde_json::from_str(&json).unwrap();
        if let Op::SetCellValue { row, col, value, .. } = parsed {
            assert_eq!(row, 5);
            assert_eq!(col, 3);
            assert_eq!(value, "hello");
        } else {
            panic!("Expected SetCellValue op");
        }
    }

    #[test]
    fn test_apply_ops_message() {
        let msg = ApplyOpsMessage {
            id: "req-1".to_string(),
            ops: vec![
                Op::SetCellValue {
                    sheet: 0,
                    row: 0,
                    col: 0,
                    value: "Revenue".to_string(),
                },
                Op::SetCellFormula {
                    sheet: 0,
                    row: 1,
                    col: 0,
                    formula: "=A1*1.1".to_string(),
                },
            ],
            atomic: true,
            expected_revision: Some(5),
        };

        let json = serde_json::to_string_pretty(&msg).unwrap();
        let parsed: ApplyOpsMessage = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.ops.len(), 2);
        assert!(parsed.atomic);
        assert_eq!(parsed.expected_revision, Some(5));
    }
}
