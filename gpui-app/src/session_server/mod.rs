//! Session server for external control of VisiGrid.
//!
//! The implementation lives in the `visigrid-session-host` crate (extracted
//! 2026-07-29 so headless hosts — `vgrid serve` — share it). This module is
//! a re-export shim: everything the GUI referenced as
//! `crate::session_server::X` still resolves. Wire types come from
//! `visigrid-protocol`; the GUI's historical mirror copy is gone.

pub use visigrid_session_host::bridge;

pub use visigrid_session_host::{
    SessionBridgeHandle, SessionRequest, BridgeError,
    ApplyOpsRequest, ApplyOpsResponse, ApplyOpsError,
    InspectRequest, InspectResponse, InspectError,
    SubscribeRequest, SubscribeResponse,
    UnsubscribeRequest, UnsubscribeResponse,
    coalesce_cells_to_ranges,
    DiscoveryFile, DiscoveryManager, discovery_dir, list_sessions,
    SessionServer, SessionServerConfig, ServerMode, EventRegistry,
    RateLimiter, RateLimiterConfig, RateLimitedError,
    EventBroadcaster, BroadcastEvent, ConnectionSubscriptions, TOPIC_CELLS,
    CellRef, ProtocolError, MAX_MESSAGE_SIZE,
};

pub use visigrid_protocol::{
    ClientMessage, ServerMessage, Op, OpError,
    InspectTarget, InspectResult, CellInfo, WorkbookInfo,
    CellRange, PROTOCOL_VERSION,
};
