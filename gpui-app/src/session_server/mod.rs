//! Session Server for external control of VisiGrid.
//!
//! Exposes a local TCP endpoint that allows external clients (CLI, agents, scripts)
//! to interact with a running VisiGrid instance:
//!
//! - Discover running sessions via discovery files
//! - Apply batches of spreadsheet operations
//! - Subscribe to live cell change events
//!
//! See: docs/future/phase-1-session-server.md

mod bridge;
mod discovery;
mod protocol;
mod server;

pub use bridge::{
    SessionBridgeHandle, SessionRequest, BridgeError,
    ApplyOpsRequest, ApplyOpsResponse, ApplyOpsError,
    InspectRequest, InspectResponse,
    SubscribeRequest, SubscribeResponse,
    UnsubscribeRequest, UnsubscribeResponse,
};
pub use discovery::{DiscoveryFile, DiscoveryManager, discovery_dir, list_sessions};
pub use protocol::{
    ClientMessage, ServerMessage, ProtocolError, Op,
    InspectTarget, InspectResult, CellInfo, WorkbookInfo,
    PROTOCOL_VERSION, MAX_MESSAGE_SIZE,
};
pub use server::{SessionServer, SessionServerConfig, ServerMode};
