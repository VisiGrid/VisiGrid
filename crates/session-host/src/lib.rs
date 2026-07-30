//! Session host: the TCP session server plus the request handlers that give
//! any process — the GUI or a headless `vgrid serve` — an agent/CLI control
//! surface over a `visigrid_engine::Workbook`.
//!
//! Wire types come from `visigrid-protocol` (single source of truth; the
//! GUI's historical mirror copy was deleted when this crate was extracted,
//! 2026-07-29). Golden vectors in the CLI pin the wire format.
//!
//! Hosts differ only in how they pump `SessionRequest`s from the bridge and
//! what they do with `ApplyOutcome` (the GUI records undo history and
//! notifies its views; headless hosts ignore both).

pub mod bridge;
pub mod coalesce;
pub mod discovery;
pub mod events;
pub mod handlers;
pub mod rate_limiter;
pub mod server;
pub mod wire_ext;

pub use bridge::{
    ApplyOpsRequest, ApplyOpsResponse, ApplyOpsError, BridgeError,
    InspectRequest, InspectResponse, InspectError,
    SessionBridgeHandle, SessionRequest, SaveOutcome, HistoryOutcome, StructureOutcome,
    SubscribeRequest, SubscribeResponse, UnsubscribeRequest, UnsubscribeResponse,
};
pub use coalesce::coalesce_cells_to_ranges;
pub use discovery::{DiscoveryFile, DiscoveryManager, discovery_dir, list_sessions};
pub use events::{BroadcastEvent, ConnectionSubscriptions, EventBroadcaster, TOPIC_CELLS};
pub use handlers::{
    apply_ops, apply_structure, inspect, structure_target_sheet, validate_inspect_target,
    validate_session_op, validate_structure_op, MAX_STRUCTURE_COUNT,
    ApplyOutcome, FormatPatch, ValueChange,
    MAX_SESSION_FORMAT_CELLS, MAX_SESSION_INSPECT_CELLS, NUM_COLS, NUM_ROWS,
};
pub use rate_limiter::{RateLimitedError, RateLimiter, RateLimiterConfig};
pub use server::{EventRegistry, ServerMode, SessionServer, SessionServerConfig};
pub use wire_ext::{CellRef, ProtocolError, MAX_MESSAGE_SIZE};
