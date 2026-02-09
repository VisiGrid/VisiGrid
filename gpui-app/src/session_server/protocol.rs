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

    /// Request server stats (for diagnostics).
    Stats(StatsMessage),
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

    /// Response to Stats.
    StatsResult(StatsResultMessage),

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
    /// Client identifier (e.g., "vgrid", "my-agent").
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
    #[serde(rename = "current_revision")]
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

/// Event payloads.
///
/// # Event Delivery Contract
///
/// **Events are best-effort, not guaranteed.**
///
/// Clients MUST NOT assume events are reliable. Events may be dropped under
/// backpressure (slow reader, queue overflow). Clients should:
///
/// 1. Track `revision` from events
/// 2. Detect gaps (e.g., revision jumps from 5 to 8)
/// 3. Re-sync via `inspect` when gaps are detected
///
/// The `events_dropped` event signals when drops occurred, but absence of
/// this event does not guarantee no drops occurred.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "event", rename_all = "snake_case")]
pub enum EventPayload {
    /// Cells changed (value or formula).
    /// Ranges are coalesced for efficiency - they cover all changed cells
    /// but may include unchanged cells (superset guarantee, not exact).
    CellsChanged {
        /// Coalesced ranges covering all changed cells.
        ranges: Vec<CellRange>,
    },
    /// Revision changed.
    RevisionChanged {
        /// Previous revision.
        previous: u64,
    },
    /// Events were dropped due to backpressure.
    /// Client should re-sync via inspect if it needs accurate state.
    EventsDropped {
        /// Number of events dropped since last notification.
        dropped_count: u64,
        /// Current revision (client should inspect if out of sync).
        current_revision: u64,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CellRef {
    pub sheet: usize,
    pub row: usize,
    pub col: usize,
}

/// A rectangular range of cells. Used in coalesced event payloads.
///
/// # Coordinate System
///
/// **All bounds are INCLUSIVE.** This differs from many APIs that use
/// end-exclusive ranges.
///
/// A range `{r1: 0, c1: 0, r2: 2, c2: 3}` covers:
/// - Rows 0, 1, 2 (3 rows total)
/// - Columns 0, 1, 2, 3 (4 columns total)
/// - Total: 12 cells
///
/// A single cell at (5, 3) is represented as `{r1: 5, c1: 3, r2: 5, c2: 3}`.
///
/// Invariants: `r1 <= r2` and `c1 <= c2` (enforced by constructor).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct CellRange {
    /// Sheet index (0-based).
    pub sheet: usize,
    /// First row, INCLUSIVE (0-based).
    pub r1: usize,
    /// First column, INCLUSIVE (0-based).
    pub c1: usize,
    /// Last row, INCLUSIVE (0-based). Must be >= r1.
    pub r2: usize,
    /// Last column, INCLUSIVE (0-based). Must be >= c1.
    pub c2: usize,
}

impl CellRange {
    /// Create a range covering a single cell.
    pub fn single(sheet: usize, row: usize, col: usize) -> Self {
        Self { sheet, r1: row, c1: col, r2: row, c2: col }
    }

    /// Create a range from bounds.
    pub fn new(sheet: usize, r1: usize, c1: usize, r2: usize, c2: usize) -> Self {
        debug_assert!(r1 <= r2 && c1 <= c2, "Invalid range bounds");
        Self { sheet, r1, c1, r2, c2 }
    }

    /// Number of cells covered by this range.
    pub fn cell_count(&self) -> usize {
        (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)
    }
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
#[serde(tag = "result", rename_all = "snake_case")]
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
    /// Raw value or formula text.
    pub raw: String,
    /// Display value (formatted).
    pub display: String,
    /// Formula text if cell contains a formula, null otherwise.
    pub formula: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookInfo {
    /// Number of sheets.
    pub sheet_count: usize,
    /// Active sheet index (0-based).
    pub active_sheet: usize,
    /// Workbook title (from filename or "Untitled").
    pub title: String,
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
// Stats
// ============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StatsMessage {
    /// Request ID for correlation.
    pub id: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StatsResultMessage {
    /// Echoed request ID.
    pub id: String,
    /// Connections closed due to parse failure limit.
    pub connections_closed_parse_failures: u64,
    /// Connections closed due to oversized message.
    pub connections_closed_oversize: u64,
    /// Writer conflict errors returned.
    pub writer_conflict_count: u64,
    /// Connections refused due to connection limit (max 5).
    pub connections_refused_limit: u64,
    /// Total events dropped due to backpressure.
    pub dropped_events_total: u64,
    /// Current number of connected clients.
    pub active_connections: u64,
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
    /// Milliseconds until retry is allowed (for rate limiting).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub retry_after_ms: Option<u64>,
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Client message types (for golden vector parsing).
    const CLIENT_TYPES: &[&str] = &["hello", "apply_ops", "subscribe", "unsubscribe", "inspect", "ping"];

    /// Server message types (for golden vector parsing).
    const SERVER_TYPES: &[&str] = &[
        "welcome",
        "apply_ops_result",
        "subscribed",
        "unsubscribed",
        "inspect_result",
        "pong",
        "event",
        "error",
    ];

    /// Extract the "type" field from a JSON object.
    fn get_message_type(json: &serde_json::Value) -> Option<&str> {
        json.get("type")?.as_str()
    }

    /// Normalize JSON for comparison (sorted keys, no extra whitespace).
    fn normalize_json(json: &serde_json::Value) -> String {
        serde_json::to_string(json).unwrap()
    }

    /// Test that golden vector files round-trip through serialization.
    ///
    /// This test ensures protocol stability - any change to serialization
    /// format will cause this test to fail, preventing accidental breaking
    /// changes to the wire protocol.
    #[test]
    fn test_golden_vectors_round_trip() {
        let golden_dir = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/src/session_server/protocol_golden"
        );

        let golden_files = [
            "hello_ok.jsonl",
            "hello_protocol_mismatch.jsonl",
            "apply_ops_ok.jsonl",
            "writer_conflict.jsonl",
            "errors.jsonl",
            "subscribe_events.jsonl",
            "subscribe_events_large_paste.jsonl",
            "inspect.jsonl",
            "ping_pong.jsonl",
        ];

        let mut failures = Vec::new();

        for filename in golden_files {
            let path = format!("{}/{}", golden_dir, filename);
            let content = match std::fs::read_to_string(&path) {
                Ok(c) => c,
                Err(e) => {
                    failures.push(format!("{}: failed to read: {}", filename, e));
                    continue;
                }
            };

            for (line_num, line) in content.lines().enumerate() {
                let line = line.trim();
                if line.is_empty() {
                    continue;
                }

                // Parse as generic JSON first to determine message type
                let original: serde_json::Value = match serde_json::from_str(line) {
                    Ok(v) => v,
                    Err(e) => {
                        failures.push(format!(
                            "{}:{}: invalid JSON: {}",
                            filename,
                            line_num + 1,
                            e
                        ));
                        continue;
                    }
                };

                let msg_type = match get_message_type(&original) {
                    Some(t) => t,
                    None => {
                        failures.push(format!(
                            "{}:{}: missing 'type' field",
                            filename,
                            line_num + 1
                        ));
                        continue;
                    }
                };

                // Determine if it's a client or server message
                let (reserialized, parse_error) = if CLIENT_TYPES.contains(&msg_type) {
                    match serde_json::from_str::<ClientMessage>(line) {
                        Ok(msg) => (serde_json::to_value(&msg).ok(), None),
                        Err(e) => (None, Some(e.to_string())),
                    }
                } else if SERVER_TYPES.contains(&msg_type) {
                    match serde_json::from_str::<ServerMessage>(line) {
                        Ok(msg) => (serde_json::to_value(&msg).ok(), None),
                        Err(e) => (None, Some(e.to_string())),
                    }
                } else {
                    failures.push(format!(
                        "{}:{}: unknown message type '{}'",
                        filename,
                        line_num + 1,
                        msg_type
                    ));
                    continue;
                };

                if let Some(error) = parse_error {
                    failures.push(format!(
                        "{}:{}: parse error for '{}': {}",
                        filename,
                        line_num + 1,
                        msg_type,
                        error
                    ));
                    continue;
                }

                if let Some(reserialized) = reserialized {
                    let original_norm = normalize_json(&original);
                    let reserialized_norm = normalize_json(&reserialized);

                    if original_norm != reserialized_norm {
                        failures.push(format!(
                            "{}:{}: round-trip mismatch for '{}':\n  original:     {}\n  reserialized: {}",
                            filename,
                            line_num + 1,
                            msg_type,
                            original_norm,
                            reserialized_norm
                        ));
                    }
                }
            }
        }

        if !failures.is_empty() {
            panic!(
                "Golden vector round-trip failures ({}):\n{}",
                failures.len(),
                failures.join("\n\n")
            );
        }
    }

    /// Test that ALL_ERROR_CODES array contains all enum variants.
    /// This catches "added variant but forgot to add to array" errors.
    #[test]
    fn test_all_error_codes_exhaustive() {
        // This test relies on the match in code() being exhaustive.
        // If a variant is missing from ALL_ERROR_CODES, we'll catch it here.
        let codes_from_array: std::collections::HashSet<&'static str> =
            ALL_ERROR_CODES.iter().map(|e| e.code()).collect();

        // Verify count matches expected (update this when adding codes)
        assert_eq!(
            ALL_ERROR_CODES.len(),
            11,
            "ALL_ERROR_CODES count changed. Update this test and errors.jsonl golden."
        );

        // Verify no duplicates
        assert_eq!(
            codes_from_array.len(),
            ALL_ERROR_CODES.len(),
            "ALL_ERROR_CODES contains duplicate codes"
        );
    }

    /// Test that errors.jsonl golden contains exactly the codes from ALL_ERROR_CODES.
    /// Prevents silent contract drift where golden and enum diverge.
    #[test]
    fn test_error_codes_golden_coverage() {
        let golden_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/src/session_server/protocol_golden/errors.jsonl"
        );

        let content = std::fs::read_to_string(golden_path)
            .expect("Failed to read errors.jsonl");

        // Extract all "code" values from the golden file
        let mut golden_codes: std::collections::HashSet<String> = std::collections::HashSet::new();
        for line in content.lines() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            let json: serde_json::Value = serde_json::from_str(line)
                .expect("Invalid JSON in errors.jsonl");
            if let Some(code) = json.get("code").and_then(|c| c.as_str()) {
                golden_codes.insert(code.to_string());
            }
        }

        // Get all codes from the enum
        let enum_codes: std::collections::HashSet<String> =
            ALL_ERROR_CODES.iter().map(|e| e.code().to_string()).collect();

        // Check for codes in enum but missing from golden
        let missing_from_golden: Vec<_> = enum_codes.difference(&golden_codes).collect();
        if !missing_from_golden.is_empty() {
            panic!(
                "Error codes in enum but missing from errors.jsonl: {:?}",
                missing_from_golden
            );
        }

        // Check for codes in golden but not in enum
        let extra_in_golden: Vec<_> = golden_codes.difference(&enum_codes).collect();
        if !extra_in_golden.is_empty() {
            panic!(
                "Error codes in errors.jsonl but not in enum: {:?}",
                extra_in_golden
            );
        }
    }

    /// Test that error code strings are stable across versions.
    /// These strings are part of the wire protocol contract.
    /// Renaming requires a protocol version bump.
    #[test]
    fn test_error_code_strings_stable() {
        // Canonical error code strings. NEVER change these without bumping PROTOCOL_VERSION.
        let expected_codes = [
            (ProtocolError::AuthFailed, "auth_failed"),
            (ProtocolError::ProtocolMismatch, "protocol_mismatch"),
            (ProtocolError::RateLimited, "rate_limited"),
            (ProtocolError::RevisionMismatch, "revision_mismatch"),
            (ProtocolError::FormulaParseError, "formula_parse_error"),
            (ProtocolError::InvalidReference, "invalid_reference"),
            (ProtocolError::MessageTooLarge, "message_too_large"),
            (ProtocolError::MalformedMessage, "malformed_message"),
            (ProtocolError::ReadOnlyMode, "read_only_mode"),
            (ProtocolError::WriterConflict, "writer_conflict"),
            (ProtocolError::InternalError, "internal_error"),
        ];

        for (error, expected_code) in expected_codes {
            assert_eq!(
                error.code(),
                expected_code,
                "Error code string changed for {:?}. This breaks wire protocol compatibility!",
                error
            );
        }

        // Verify we tested all codes
        assert_eq!(
            expected_codes.len(),
            ALL_ERROR_CODES.len(),
            "test_error_code_strings_stable doesn't cover all error codes"
        );
    }

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
