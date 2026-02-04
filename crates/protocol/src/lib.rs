//! VisiGrid Session Server Protocol — v1 Frozen Wire Format
//!
//! This crate defines the canonical protocol types for CLI ↔ GUI communication.
//! The wire format is JSONL (newline-delimited JSON) over TCP localhost.
//!
//! # Protocol Version
//!
//! This is **protocol v1** — the wire format is frozen. Changes require:
//! 1. Version bump in PROTOCOL_VERSION
//! 2. New golden vectors in `gpui-app/src/session_server/protocol_golden/`
//! 3. Backward compatibility handling
//!
//! # Usage
//!
//! ```ignore
//! use visigrid_protocol::{ClientMessage, ServerMessage, PROTOCOL_VERSION};
//!
//! // Serialize a client message
//! let msg = ClientMessage::Ping(PingMessage { id: "1".into() });
//! let json = serde_json::to_string(&msg)?;
//!
//! // Deserialize a server message
//! let response: ServerMessage = serde_json::from_str(&line)?;
//! ```

use serde::{Deserialize, Serialize};

/// Current protocol version. Increment for breaking changes.
pub const PROTOCOL_VERSION: u32 = 1;

// =============================================================================
// Client → Server Messages
// =============================================================================

/// Messages sent from client (CLI) to server (GUI).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ClientMessage {
    Hello(HelloMessage),
    ApplyOps(ApplyOpsMessage),
    Inspect(InspectMessage),
    Ping(PingMessage),
    Subscribe(SubscribeMessage),
    Unsubscribe(UnsubscribeMessage),
    Stats(StatsMessage),
}

/// Initial handshake from client.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HelloMessage {
    pub id: String,
    pub client: String,
    pub version: String,
    pub token: String,
    #[serde(default = "default_protocol_version")]
    pub protocol_version: u32,
}

fn default_protocol_version() -> u32 {
    1
}

/// Request to apply operations to the workbook.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyOpsMessage {
    pub id: String,
    pub ops: Vec<Op>,
    #[serde(default)]
    pub atomic: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub expected_revision: Option<u64>,
}

/// A single operation to apply.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum Op {
    SetCellValue {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
        value: String,
    },
    SetCellFormula {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
        formula: String,
    },
    ClearCell {
        #[serde(default)]
        sheet: usize,
        row: usize,
        col: usize,
    },
    SetNumberFormat {
        #[serde(default)]
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        format: String,
    },
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

/// Request to inspect cell/range/workbook state.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectMessage {
    pub id: String,
    pub target: InspectTarget,
}

/// What to inspect.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "target", rename_all = "snake_case")]
pub enum InspectTarget {
    Cell {
        sheet: usize,
        row: usize,
        col: usize,
    },
    Range {
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
    Workbook,
}

/// Ping for keepalive.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PingMessage {
    pub id: String,
}

/// Subscribe to event topics.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SubscribeMessage {
    pub id: String,
    pub topics: Vec<String>,
}

/// Unsubscribe from event topics.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsubscribeMessage {
    pub id: String,
    pub topics: Vec<String>,
}

/// Request server statistics.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StatsMessage {
    pub id: String,
}

// =============================================================================
// Server → Client Messages
// =============================================================================

/// Messages sent from server (GUI) to client (CLI).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ServerMessage {
    Welcome(WelcomeMessage),
    ApplyOpsResult(ApplyOpsResultMessage),
    InspectResult(InspectResultMessage),
    Pong(PongMessage),
    Error(ErrorMessage),
    Subscribed(SubscribedMessage),
    Unsubscribed(UnsubscribedMessage),
    Event(EventMessage),
    StatsResult(StatsResultMessage),
}

/// Welcome response after successful hello.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WelcomeMessage {
    pub id: String,
    pub session_id: String,
    pub protocol_version: u32,
    pub revision: u64,
    pub capabilities: Vec<String>,
}

/// Result of apply_ops request.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ApplyOpsResultMessage {
    pub id: String,
    pub applied: usize,
    pub total: usize,
    #[serde(rename = "current_revision")]
    pub revision: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<OpError>,
}

/// Error applying a specific operation.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OpError {
    pub code: String,
    pub message: String,
    pub op_index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suggestion: Option<String>,
}

/// Result of inspect request.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InspectResultMessage {
    pub id: String,
    pub revision: u64,
    pub result: InspectResult,
}

/// Inspection result variants.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "result", rename_all = "snake_case")]
pub enum InspectResult {
    Cell(CellInfo),
    Range { cells: Vec<CellInfo> },
    Workbook(WorkbookInfo),
}

/// Information about a single cell.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellInfo {
    pub raw: String,
    pub display: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
}

/// Information about the workbook.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookInfo {
    pub sheet_count: usize,
    pub active_sheet: usize,
    pub title: String,
}

/// Pong response to ping.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PongMessage {
    pub id: String,
}

/// Error response.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ErrorMessage {
    pub id: String,
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub retry_after_ms: Option<u64>,
}

/// Confirmation of subscription.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SubscribedMessage {
    pub id: String,
    pub topics: Vec<String>,
}

/// Confirmation of unsubscription.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsubscribedMessage {
    pub id: String,
    pub topics: Vec<String>,
}

/// Server-sent event.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventMessage {
    pub topic: String,
    pub revision: u64,
    pub payload: EventPayload,
}

/// Event payload variants.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "event", rename_all = "snake_case")]
pub enum EventPayload {
    CellsChanged { ranges: Vec<CellRange> },
    EventsDropped {
        dropped_count: u64,
        current_revision: u64,
    },
}

/// A rectangular range of cells.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellRange {
    pub sheet: usize,
    pub r1: usize,
    pub c1: usize,
    pub r2: usize,
    pub c2: usize,
}

/// Server statistics result.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StatsResultMessage {
    pub id: String,
    pub connections_closed_parse_failures: u64,
    pub connections_closed_oversize: u64,
    pub writer_conflict_count: u64,
    pub connections_refused_limit: u64,
    pub dropped_events_total: u64,
    pub active_connections: u64,
}

// =============================================================================
// Discovery File Format
// =============================================================================

/// Discovery file written by GUI, read by CLI to find running sessions.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DiscoveryFile {
    pub session_id: String,
    pub port: u16,
    pub pid: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub workbook_path: Option<std::path::PathBuf>,
    pub workbook_title: String,
    pub created_at: String, // ISO 8601 format
    pub protocol_version: u32,
}
