//! Bridge types for session server ↔ engine communication.
//!
//! The TCP server runs on a separate thread and cannot mutate the workbook
//! directly. Instead, it sends `SessionRequest` messages through an mpsc
//! channel to the GUI/engine thread, which processes them using the
//! canonical mutation path and sends responses back via oneshot channels.
//!
//! This ensures:
//! 1. All mutations go through the same code path (GUI and session server)
//! 2. No deadlocks from cross-thread workbook access
//! 3. Proper revision tracking and event emission

use std::sync::mpsc;

use super::protocol::{Op, InspectTarget, InspectResult, OpError};

/// A simple oneshot channel for single-use responses.
/// Uses std::sync::mpsc under the hood.
pub mod oneshot {
    use std::sync::mpsc;

    pub struct Sender<T>(mpsc::SyncSender<T>);
    pub struct Receiver<T>(mpsc::Receiver<T>);

    impl<T> Sender<T> {
        pub fn send(self, value: T) -> Result<(), T> {
            self.0.send(value).map_err(|e| e.0)
        }
    }

    impl<T> Receiver<T> {
        pub fn blocking_recv(self) -> Result<T, RecvError> {
            self.0.recv().map_err(|_| RecvError)
        }
    }

    #[derive(Debug, Clone, Copy)]
    pub struct RecvError;

    pub fn channel<T>() -> (Sender<T>, Receiver<T>) {
        // Buffer of 1 for oneshot semantics
        let (tx, rx) = mpsc::sync_channel(1);
        (Sender(tx), Receiver(rx))
    }
}

/// Handle passed to the session server for sending requests to the engine.
#[derive(Clone)]
pub struct SessionBridgeHandle {
    /// Channel for sending requests to the engine thread.
    pub tx: mpsc::Sender<SessionRequest>,
}

impl SessionBridgeHandle {
    pub fn new(tx: mpsc::Sender<SessionRequest>) -> Self {
        Self { tx }
    }

    /// Send an apply_ops request and wait for the response.
    pub fn apply_ops(&self, req: ApplyOpsRequest) -> Result<ApplyOpsResponse, BridgeError> {
        let (reply_tx, reply_rx) = oneshot::channel();
        self.tx
            .send(SessionRequest::ApplyOps { req, reply: reply_tx })
            .map_err(|_| BridgeError::ChannelClosed)?;
        reply_rx.blocking_recv().map_err(|_| BridgeError::ChannelClosed)
    }

    /// Send an inspect request and wait for the response.
    pub fn inspect(&self, req: InspectRequest) -> Result<InspectResponse, BridgeError> {
        let (reply_tx, reply_rx) = oneshot::channel();
        self.tx
            .send(SessionRequest::Inspect { req, reply: reply_tx })
            .map_err(|_| BridgeError::ChannelClosed)?;
        reply_rx.blocking_recv().map_err(|_| BridgeError::ChannelClosed)
    }

    /// Send a subscribe request (fire-and-forget for now).
    pub fn subscribe(&self, req: SubscribeRequest) -> Result<SubscribeResponse, BridgeError> {
        let (reply_tx, reply_rx) = oneshot::channel();
        self.tx
            .send(SessionRequest::Subscribe { req, reply: reply_tx })
            .map_err(|_| BridgeError::ChannelClosed)?;
        reply_rx.blocking_recv().map_err(|_| BridgeError::ChannelClosed)
    }

    /// Send an unsubscribe request (fire-and-forget for now).
    pub fn unsubscribe(&self, req: UnsubscribeRequest) -> Result<UnsubscribeResponse, BridgeError> {
        let (reply_tx, reply_rx) = oneshot::channel();
        self.tx
            .send(SessionRequest::Unsubscribe { req, reply: reply_tx })
            .map_err(|_| BridgeError::ChannelClosed)?;
        reply_rx.blocking_recv().map_err(|_| BridgeError::ChannelClosed)
    }
}

/// Errors from bridge communication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BridgeError {
    /// The channel to the engine thread was closed.
    ChannelClosed,
}

/// Requests from session server to engine.
pub enum SessionRequest {
    /// Apply a batch of operations.
    ApplyOps {
        req: ApplyOpsRequest,
        reply: oneshot::Sender<ApplyOpsResponse>,
    },
    /// Inspect workbook state.
    Inspect {
        req: InspectRequest,
        reply: oneshot::Sender<InspectResponse>,
    },
    /// Subscribe to events.
    Subscribe {
        req: SubscribeRequest,
        reply: oneshot::Sender<SubscribeResponse>,
    },
    /// Unsubscribe from events.
    Unsubscribe {
        req: UnsubscribeRequest,
        reply: oneshot::Sender<UnsubscribeResponse>,
    },
}

// ============================================================================
// ApplyOps
// ============================================================================

/// Request to apply a batch of operations.
#[derive(Debug, Clone)]
pub struct ApplyOpsRequest {
    /// Request ID for correlation (from wire protocol).
    pub request_id: String,
    /// Human-readable batch name for undo stack.
    pub batch_name: String,
    /// If true, all-or-nothing semantics.
    pub atomic: bool,
    /// Expected revision for optimistic concurrency.
    /// If set and doesn't match current revision, request is rejected.
    pub expected_revision: Option<u64>,
    /// Operations to apply.
    pub ops: Vec<Op>,
}

/// Response to apply_ops request.
#[derive(Debug, Clone)]
pub struct ApplyOpsResponse {
    /// Number of ops successfully applied.
    pub applied: usize,
    /// Total number of ops in the request.
    pub total: usize,
    /// Current revision after operation (whether successful or not).
    /// INVARIANT: Always present per spec.
    pub current_revision: u64,
    /// Error if any op failed.
    pub error: Option<ApplyOpsError>,
}

/// Error details for apply_ops.
#[derive(Debug, Clone)]
pub enum ApplyOpsError {
    /// Expected revision didn't match current revision.
    RevisionMismatch {
        expected: u64,
        actual: u64,
    },
    /// An operation failed.
    OpFailed(OpError),
}

// ============================================================================
// Inspect
// ============================================================================

/// Request to inspect workbook state.
#[derive(Debug, Clone)]
pub struct InspectRequest {
    /// Request ID for correlation.
    pub request_id: String,
    /// What to inspect.
    pub target: InspectTarget,
}

/// Response to inspect request.
#[derive(Debug, Clone)]
pub struct InspectResponse {
    /// Current revision at time of inspection.
    /// INVARIANT: Always present per spec.
    pub current_revision: u64,
    /// Inspection result.
    pub result: InspectResult,
}

// ============================================================================
// Subscribe / Unsubscribe
// ============================================================================

/// Request to subscribe to events.
#[derive(Debug, Clone)]
pub struct SubscribeRequest {
    /// Request ID for correlation.
    pub request_id: String,
    /// Topics to subscribe to.
    pub topics: Vec<String>,
}

/// Response to subscribe request.
#[derive(Debug, Clone)]
pub struct SubscribeResponse {
    /// Topics successfully subscribed.
    pub topics: Vec<String>,
    /// Current revision at time of subscription.
    pub current_revision: u64,
}

/// Request to unsubscribe from events.
#[derive(Debug, Clone)]
pub struct UnsubscribeRequest {
    /// Request ID for correlation.
    pub request_id: String,
    /// Topics to unsubscribe from.
    pub topics: Vec<String>,
}

/// Response to unsubscribe request.
#[derive(Debug, Clone)]
pub struct UnsubscribeResponse {
    /// Topics successfully unsubscribed.
    pub topics: Vec<String>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::mpsc;

    #[test]
    fn test_bridge_handle_creation() {
        let (tx, _rx) = mpsc::channel();
        let handle = SessionBridgeHandle::new(tx);
        // Handle should be cloneable
        let _handle2 = handle.clone();
    }

    #[test]
    fn test_apply_ops_request_construction() {
        let req = ApplyOpsRequest {
            request_id: "req-1".to_string(),
            batch_name: "Set cell A1".to_string(),
            atomic: true,
            expected_revision: Some(5),
            ops: vec![Op::SetCellValue {
                sheet: 0,
                row: 0,
                col: 0,
                value: "Hello".to_string(),
            }],
        };
        assert_eq!(req.ops.len(), 1);
        assert!(req.atomic);
        assert_eq!(req.expected_revision, Some(5));
    }

    #[test]
    fn test_apply_ops_response_always_has_revision() {
        // Success case
        let success = ApplyOpsResponse {
            applied: 3,
            total: 3,
            current_revision: 10,
            error: None,
        };
        assert_eq!(success.current_revision, 10);

        // Error case - still has revision
        let failure = ApplyOpsResponse {
            applied: 0,
            total: 3,
            current_revision: 9, // Unchanged from before
            error: Some(ApplyOpsError::RevisionMismatch {
                expected: 8,
                actual: 9,
            }),
        };
        assert_eq!(failure.current_revision, 9);
    }

    #[test]
    fn test_inspect_response_always_has_revision() {
        let response = InspectResponse {
            current_revision: 42,
            result: InspectResult::Workbook(super::super::protocol::WorkbookInfo {
                sheet_count: 1,
                sheets: vec!["Sheet1".to_string()],
                revision: 42,
            }),
        };
        assert_eq!(response.current_revision, 42);
    }
}
