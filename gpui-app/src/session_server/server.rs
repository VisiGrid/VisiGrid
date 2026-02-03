//! TCP server for session server protocol.
//!
//! Binds to 127.0.0.1:<random_port> and handles JSONL messages.

use std::io::{BufRead, BufReader, Write};
use std::net::{TcpListener, TcpStream, SocketAddr};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::thread::{self, JoinHandle};

use crate::session_server::bridge::{
    SessionBridgeHandle, ApplyOpsRequest, InspectRequest,
    SubscribeRequest, UnsubscribeRequest,
};
use crate::session_server::discovery::DiscoveryManager;
use crate::session_server::protocol::*;

/// Server operating mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ServerMode {
    /// Server is not running.
    #[default]
    Off,
    /// Server accepts connections but only allows read operations (inspect).
    ReadOnly,
    /// Server accepts connections and allows mutations (apply_ops).
    Apply,
}

/// Configuration for the session server.
#[derive(Clone)]
pub struct SessionServerConfig {
    /// Operating mode.
    pub mode: ServerMode,
    /// Workbook path (if saved).
    pub workbook_path: Option<PathBuf>,
    /// Workbook title for display.
    pub workbook_title: String,
    /// Bridge handle for engine communication.
    /// Required when mode != Off.
    pub bridge: Option<SessionBridgeHandle>,
}

impl std::fmt::Debug for SessionServerConfig {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SessionServerConfig")
            .field("mode", &self.mode)
            .field("workbook_path", &self.workbook_path)
            .field("workbook_title", &self.workbook_title)
            .field("bridge", &self.bridge.as_ref().map(|_| "..."))
            .finish()
    }
}

impl Default for SessionServerConfig {
    fn default() -> Self {
        Self {
            mode: ServerMode::Off,
            workbook_path: None,
            workbook_title: "Untitled".to_string(),
            bridge: None,
        }
    }
}

/// The session server - manages TCP listener and client connections.
pub struct SessionServer {
    /// Discovery manager (writes discovery file).
    discovery: Option<DiscoveryManager>,
    /// TCP listener handle.
    listener_handle: Option<JoinHandle<()>>,
    /// Shutdown signal.
    shutdown: Arc<AtomicBool>,
    /// Current mode.
    mode: ServerMode,
    /// Bound address (if running).
    bound_addr: Option<SocketAddr>,
}

impl SessionServer {
    /// Create a new session server (not started).
    pub fn new() -> Self {
        Self {
            discovery: None,
            listener_handle: None,
            shutdown: Arc::new(AtomicBool::new(false)),
            mode: ServerMode::Off,
            bound_addr: None,
        }
    }

    /// Start the server with the given configuration.
    ///
    /// If mode != Off, the bridge handle is required.
    pub fn start(&mut self, config: SessionServerConfig) -> std::io::Result<()> {
        if self.is_running() {
            return Ok(());
        }

        if config.mode == ServerMode::Off {
            return Ok(());
        }

        // Bridge is required when server is active
        let bridge = config.bridge.ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "SessionBridgeHandle required when mode != Off",
            )
        })?;

        self.mode = config.mode;
        self.shutdown.store(false, Ordering::SeqCst);

        // Bind to localhost on a random port
        let listener = TcpListener::bind("127.0.0.1:0")?;
        let addr = listener.local_addr()?;
        self.bound_addr = Some(addr);

        // Set non-blocking so we can check shutdown flag
        listener.set_nonblocking(true)?;

        // Create discovery file
        self.discovery = Some(DiscoveryManager::new(
            addr.port(),
            config.workbook_path,
            config.workbook_title,
        )?);

        // Spawn listener thread
        let shutdown = Arc::clone(&self.shutdown);
        let mode = self.mode;
        let token = self.discovery.as_ref().unwrap().token().to_string();
        let session_id = self.discovery.as_ref().unwrap().session_id().to_string();

        self.listener_handle = Some(thread::spawn(move || {
            run_listener(listener, shutdown, mode, token, session_id, bridge);
        }));

        log::info!(
            "Session server started on {} (mode: {:?})",
            addr,
            self.mode
        );

        Ok(())
    }

    /// Stop the server.
    pub fn stop(&mut self) {
        if !self.is_running() {
            return;
        }

        // Signal shutdown
        self.shutdown.store(true, Ordering::SeqCst);

        // Wait for listener thread
        if let Some(handle) = self.listener_handle.take() {
            let _ = handle.join();
        }

        // Cleanup discovery file
        if let Some(discovery) = self.discovery.take() {
            let _ = discovery.cleanup();
        }

        self.mode = ServerMode::Off;
        self.bound_addr = None;

        log::info!("Session server stopped");
    }

    /// Check if the server is running.
    pub fn is_running(&self) -> bool {
        self.listener_handle.is_some() && !self.shutdown.load(Ordering::SeqCst)
    }

    /// Get the current mode.
    pub fn mode(&self) -> ServerMode {
        self.mode
    }

    /// Get the bound address (if running).
    pub fn bound_addr(&self) -> Option<SocketAddr> {
        self.bound_addr
    }

    /// Get the session ID (if running).
    pub fn session_id(&self) -> Option<uuid::Uuid> {
        self.discovery.as_ref().map(|d| d.session_id())
    }

    /// Get the token (if running). Only for display in GUI.
    pub fn token(&self) -> Option<&str> {
        self.discovery.as_ref().map(|d| d.token())
    }

    /// Update workbook info (called when file is saved/opened).
    pub fn update_workbook(&mut self, path: Option<PathBuf>, title: String) {
        if let Some(discovery) = &mut self.discovery {
            let _ = discovery.update_workbook(path, title);
        }
    }
}

impl Default for SessionServer {
    fn default() -> Self {
        Self::new()
    }
}

impl Drop for SessionServer {
    fn drop(&mut self) {
        self.stop();
    }
}

/// Run the listener loop in a separate thread.
fn run_listener(
    listener: TcpListener,
    shutdown: Arc<AtomicBool>,
    mode: ServerMode,
    token: String,
    session_id: String,
    bridge: SessionBridgeHandle,
) {
    while !shutdown.load(Ordering::SeqCst) {
        match listener.accept() {
            Ok((stream, addr)) => {
                log::debug!("Accepted connection from {}", addr);
                let token = token.clone();
                let session_id = session_id.clone();
                let bridge = bridge.clone();
                let mode = mode;

                // Handle each connection in its own thread
                thread::spawn(move || {
                    if let Err(e) = handle_connection(stream, mode, &token, &session_id, &bridge) {
                        log::warn!("Connection error from {}: {}", addr, e);
                    }
                });
            }
            Err(ref e) if e.kind() == std::io::ErrorKind::WouldBlock => {
                // No connection ready, sleep briefly
                thread::sleep(std::time::Duration::from_millis(50));
            }
            Err(e) => {
                log::error!("Accept error: {}", e);
                break;
            }
        }
    }
}

/// Handle a single client connection.
fn handle_connection(
    mut stream: TcpStream,
    mode: ServerMode,
    expected_token: &str,
    session_id: &str,
    bridge: &SessionBridgeHandle,
) -> std::io::Result<()> {
    stream.set_nonblocking(false)?;
    stream.set_read_timeout(Some(std::time::Duration::from_secs(30)))?;
    stream.set_write_timeout(Some(std::time::Duration::from_secs(10)))?;

    let reader = BufReader::new(stream.try_clone()?);
    let mut authenticated = false;

    for line in reader.lines() {
        let line = line?;

        // Check message size
        if line.len() > MAX_MESSAGE_SIZE {
            send_error(&mut stream, None, ProtocolError::MessageTooLarge)?;
            continue;
        }

        // Parse message
        let msg: ClientMessage = match serde_json::from_str(&line) {
            Ok(m) => m,
            Err(e) => {
                log::debug!("Malformed message: {}", e);
                send_error(&mut stream, None, ProtocolError::MalformedMessage)?;
                continue;
            }
        };

        // First message must be Hello
        if !authenticated {
            match msg {
                ClientMessage::Hello(hello) => {
                    if hello.token != expected_token {
                        send_error(&mut stream, Some(hello.id), ProtocolError::AuthFailed)?;
                        return Ok(());
                    }

                    // Check protocol version
                    if hello.protocol_version > PROTOCOL_VERSION {
                        send_error(&mut stream, Some(hello.id), ProtocolError::ProtocolMismatch)?;
                        return Ok(());
                    }

                    authenticated = true;

                    // Get current revision from engine via inspect
                    let revision = match bridge.inspect(InspectRequest {
                        request_id: hello.id.clone(),
                        target: InspectTarget::Workbook,
                    }) {
                        Ok(resp) => resp.current_revision,
                        Err(_) => 0, // Fallback if bridge error
                    };

                    let response = ServerMessage::Welcome(WelcomeMessage {
                        id: hello.id,
                        session_id: session_id.to_string(),
                        protocol_version: hello.protocol_version.min(PROTOCOL_VERSION),
                        revision,
                        capabilities: vec!["apply_ops".to_string(), "inspect".to_string()],
                    });
                    send_message(&mut stream, &response)?;
                }
                _ => {
                    send_error(&mut stream, None, ProtocolError::AuthFailed)?;
                    return Ok(());
                }
            }
            continue;
        }

        // Handle authenticated messages
        let response = handle_message(msg, mode, bridge);
        send_message(&mut stream, &response)?;
    }

    Ok(())
}

/// Handle a single message and return the response.
fn handle_message(msg: ClientMessage, mode: ServerMode, bridge: &SessionBridgeHandle) -> ServerMessage {
    match msg {
        ClientMessage::Hello(h) => {
            // Already authenticated, treat as error
            ServerMessage::Error(ErrorMessage {
                id: Some(h.id),
                code: "already_authenticated".to_string(),
                message: "Already authenticated".to_string(),
            })
        }
        ClientMessage::ApplyOps(apply) => {
            if mode == ServerMode::ReadOnly {
                return ServerMessage::Error(
                    ProtocolError::ReadOnlyMode.to_error_message(Some(apply.id)),
                );
            }

            // Send through bridge to engine thread
            let req = ApplyOpsRequest {
                request_id: apply.id.clone(),
                batch_name: format!("Session: {} ops", apply.ops.len()),
                atomic: apply.atomic,
                expected_revision: apply.expected_revision,
                ops: apply.ops.clone(),
            };

            match bridge.apply_ops(req) {
                Ok(resp) => {
                    ServerMessage::ApplyOpsResult(ApplyOpsResultMessage {
                        id: apply.id,
                        applied: resp.applied,
                        total: resp.total,
                        revision: resp.current_revision,
                        error: resp.error.map(|e| match e {
                            crate::session_server::bridge::ApplyOpsError::RevisionMismatch { expected, actual } => {
                                OpError {
                                    code: "revision_mismatch".to_string(),
                                    message: format!("Expected revision {} but current is {}", expected, actual),
                                    op_index: 0,
                                    suggestion: Some("Retry with updated revision".to_string()),
                                }
                            }
                            crate::session_server::bridge::ApplyOpsError::OpFailed(op_err) => op_err,
                        }),
                    })
                }
                Err(_) => {
                    ServerMessage::Error(ErrorMessage {
                        id: Some(apply.id),
                        code: "internal_error".to_string(),
                        message: "Bridge communication failed".to_string(),
                    })
                }
            }
        }
        ClientMessage::Subscribe(sub) => {
            let req = SubscribeRequest {
                request_id: sub.id.clone(),
                topics: sub.topics.clone(),
            };

            match bridge.subscribe(req) {
                Ok(resp) => ServerMessage::Subscribed(SubscribedMessage {
                    id: sub.id,
                    topics: resp.topics,
                }),
                Err(_) => ServerMessage::Error(ErrorMessage {
                    id: Some(sub.id),
                    code: "internal_error".to_string(),
                    message: "Bridge communication failed".to_string(),
                }),
            }
        }
        ClientMessage::Unsubscribe(unsub) => {
            let req = UnsubscribeRequest {
                request_id: unsub.id.clone(),
                topics: unsub.topics.clone(),
            };

            match bridge.unsubscribe(req) {
                Ok(resp) => ServerMessage::Unsubscribed(UnsubscribedMessage {
                    id: unsub.id,
                    topics: resp.topics,
                }),
                Err(_) => ServerMessage::Error(ErrorMessage {
                    id: Some(unsub.id),
                    code: "internal_error".to_string(),
                    message: "Bridge communication failed".to_string(),
                }),
            }
        }
        ClientMessage::Inspect(inspect) => {
            let req = InspectRequest {
                request_id: inspect.id.clone(),
                target: inspect.target.clone(),
            };

            match bridge.inspect(req) {
                Ok(resp) => ServerMessage::InspectResult(InspectResultMessage {
                    id: inspect.id,
                    revision: resp.current_revision,
                    result: resp.result,
                }),
                Err(_) => ServerMessage::Error(ErrorMessage {
                    id: Some(inspect.id),
                    code: "internal_error".to_string(),
                    message: "Bridge communication failed".to_string(),
                }),
            }
        }
        ClientMessage::Ping(ping) => ServerMessage::Pong(PongMessage { id: ping.id }),
    }
}

/// Send a message to the client.
fn send_message(stream: &mut TcpStream, msg: &ServerMessage) -> std::io::Result<()> {
    let json = serde_json::to_string(msg)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    writeln!(stream, "{}", json)?;
    stream.flush()
}

/// Send an error message to the client.
fn send_error(stream: &mut TcpStream, id: Option<String>, error: ProtocolError) -> std::io::Result<()> {
    let msg = ServerMessage::Error(error.to_error_message(id));
    send_message(stream, &msg)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::session_server::bridge::{
        SessionRequest, ApplyOpsResponse, InspectResponse,
        SubscribeResponse, UnsubscribeResponse,
    };
    use std::io::{BufRead, BufReader, Write};
    use std::net::TcpStream;
    use std::sync::mpsc;

    /// Creates a test bridge with a mock handler that responds to requests.
    fn create_test_bridge() -> (SessionBridgeHandle, thread::JoinHandle<()>) {
        let (tx, rx) = mpsc::channel::<SessionRequest>();
        let handle = SessionBridgeHandle::new(tx);

        let handler = thread::spawn(move || {
            while let Ok(req) = rx.recv() {
                match req {
                    SessionRequest::ApplyOps { req, reply } => {
                        let _ = reply.send(ApplyOpsResponse {
                            applied: req.ops.len(),
                            total: req.ops.len(),
                            current_revision: 1,
                            error: None,
                        });
                    }
                    SessionRequest::Inspect { req: _, reply } => {
                        let _ = reply.send(InspectResponse {
                            current_revision: 0,
                            result: InspectResult::Workbook(WorkbookInfo {
                                sheet_count: 1,
                                sheets: vec!["Sheet1".to_string()],
                                revision: 0,
                            }),
                        });
                    }
                    SessionRequest::Subscribe { req, reply } => {
                        let _ = reply.send(SubscribeResponse {
                            topics: req.topics,
                            current_revision: 0,
                        });
                    }
                    SessionRequest::Unsubscribe { req, reply } => {
                        let _ = reply.send(UnsubscribeResponse {
                            topics: req.topics,
                        });
                    }
                }
            }
        });

        (handle, handler)
    }

    #[test]
    fn test_server_lifecycle() {
        let (bridge, _handler) = create_test_bridge();
        let mut server = SessionServer::new();
        assert!(!server.is_running());

        // Start in Apply mode
        server
            .start(SessionServerConfig {
                mode: ServerMode::Apply,
                workbook_path: None,
                workbook_title: "Test".to_string(),
                bridge: Some(bridge),
            })
            .unwrap();

        assert!(server.is_running());
        assert!(server.bound_addr().is_some());
        assert!(server.token().is_some());

        // Stop
        server.stop();
        assert!(!server.is_running());
    }

    #[test]
    fn test_server_requires_bridge() {
        let mut server = SessionServer::new();

        // Starting without bridge should fail
        let result = server.start(SessionServerConfig {
            mode: ServerMode::Apply,
            workbook_path: None,
            workbook_title: "Test".to_string(),
            bridge: None,
        });

        assert!(result.is_err());
        assert!(!server.is_running());
    }

    #[test]
    fn test_server_connection() {
        let (bridge, _handler) = create_test_bridge();
        let mut server = SessionServer::new();
        server
            .start(SessionServerConfig {
                mode: ServerMode::Apply,
                workbook_path: None,
                workbook_title: "Test".to_string(),
                bridge: Some(bridge),
            })
            .unwrap();

        let addr = server.bound_addr().unwrap();
        let token = server.token().unwrap().to_string();

        // Connect
        let mut stream = TcpStream::connect(addr).unwrap();
        stream
            .set_read_timeout(Some(std::time::Duration::from_secs(5)))
            .unwrap();

        // Send hello
        let hello = serde_json::json!({
            "type": "hello",
            "id": "1",
            "client": "test",
            "version": "1.0.0",
            "token": token,
            "protocol_version": 1
        });
        writeln!(stream, "{}", hello).unwrap();

        // Read welcome
        let mut reader = BufReader::new(stream.try_clone().unwrap());
        let mut response = String::new();
        reader.read_line(&mut response).unwrap();

        let msg: ServerMessage = serde_json::from_str(&response).unwrap();
        assert!(matches!(msg, ServerMessage::Welcome(_)));

        // Send ping
        let ping = serde_json::json!({
            "type": "ping",
            "id": "2"
        });
        writeln!(stream, "{}", ping).unwrap();

        // Read pong
        response.clear();
        reader.read_line(&mut response).unwrap();

        let msg: ServerMessage = serde_json::from_str(&response).unwrap();
        assert!(matches!(msg, ServerMessage::Pong(_)));

        server.stop();
    }

    #[test]
    fn test_server_auth_failure() {
        let (bridge, _handler) = create_test_bridge();
        let mut server = SessionServer::new();
        server
            .start(SessionServerConfig {
                mode: ServerMode::Apply,
                workbook_path: None,
                workbook_title: "Test".to_string(),
                bridge: Some(bridge),
            })
            .unwrap();

        let addr = server.bound_addr().unwrap();

        // Connect
        let mut stream = TcpStream::connect(addr).unwrap();
        stream
            .set_read_timeout(Some(std::time::Duration::from_secs(5)))
            .unwrap();

        // Send hello with wrong token
        let hello = serde_json::json!({
            "type": "hello",
            "id": "1",
            "client": "test",
            "version": "1.0.0",
            "token": "wrong_token",
            "protocol_version": 1
        });
        writeln!(stream, "{}", hello).unwrap();

        // Read error
        let mut reader = BufReader::new(stream);
        let mut response = String::new();
        reader.read_line(&mut response).unwrap();

        let msg: ServerMessage = serde_json::from_str(&response).unwrap();
        if let ServerMessage::Error(e) = msg {
            assert_eq!(e.code, "auth_failed");
        } else {
            panic!("Expected error message");
        }

        server.stop();
    }

    #[test]
    fn test_server_apply_ops_via_bridge() {
        let (bridge, _handler) = create_test_bridge();
        let mut server = SessionServer::new();
        server
            .start(SessionServerConfig {
                mode: ServerMode::Apply,
                workbook_path: None,
                workbook_title: "Test".to_string(),
                bridge: Some(bridge),
            })
            .unwrap();

        let addr = server.bound_addr().unwrap();
        let token = server.token().unwrap().to_string();

        // Connect and authenticate
        let mut stream = TcpStream::connect(addr).unwrap();
        stream.set_read_timeout(Some(std::time::Duration::from_secs(5))).unwrap();

        let hello = serde_json::json!({
            "type": "hello",
            "id": "1",
            "client": "test",
            "version": "1.0.0",
            "token": token,
            "protocol_version": 1
        });
        writeln!(stream, "{}", hello).unwrap();

        let mut reader = BufReader::new(stream.try_clone().unwrap());
        let mut response = String::new();
        reader.read_line(&mut response).unwrap();

        // Send apply_ops
        let apply = serde_json::json!({
            "type": "apply_ops",
            "id": "2",
            "ops": [
                {"op": "set_cell_value", "row": 0, "col": 0, "value": "Hello"}
            ],
            "atomic": true
        });
        writeln!(stream, "{}", apply).unwrap();

        // Read result
        response.clear();
        reader.read_line(&mut response).unwrap();

        let msg: ServerMessage = serde_json::from_str(&response).unwrap();
        if let ServerMessage::ApplyOpsResult(result) = msg {
            assert_eq!(result.applied, 1);
            assert_eq!(result.total, 1);
            assert_eq!(result.revision, 1);
            assert!(result.error.is_none());
        } else {
            panic!("Expected ApplyOpsResult, got {:?}", msg);
        }

        server.stop();
    }

    #[test]
    fn test_server_inspect_via_bridge() {
        let (bridge, _handler) = create_test_bridge();
        let mut server = SessionServer::new();
        server
            .start(SessionServerConfig {
                mode: ServerMode::Apply,
                workbook_path: None,
                workbook_title: "Test".to_string(),
                bridge: Some(bridge),
            })
            .unwrap();

        let addr = server.bound_addr().unwrap();
        let token = server.token().unwrap().to_string();

        // Connect and authenticate
        let mut stream = TcpStream::connect(addr).unwrap();
        stream.set_read_timeout(Some(std::time::Duration::from_secs(5))).unwrap();

        let hello = serde_json::json!({
            "type": "hello",
            "id": "1",
            "client": "test",
            "version": "1.0.0",
            "token": token,
            "protocol_version": 1
        });
        writeln!(stream, "{}", hello).unwrap();

        let mut reader = BufReader::new(stream.try_clone().unwrap());
        let mut response = String::new();
        reader.read_line(&mut response).unwrap();

        // Send inspect
        let inspect = serde_json::json!({
            "type": "inspect",
            "id": "2",
            "target": {"target": "workbook"}
        });
        writeln!(stream, "{}", inspect).unwrap();

        // Read result
        response.clear();
        reader.read_line(&mut response).unwrap();

        let msg: ServerMessage = serde_json::from_str(&response).unwrap();
        if let ServerMessage::InspectResult(result) = msg {
            assert_eq!(result.revision, 0);
            if let InspectResult::Workbook(info) = result.result {
                assert_eq!(info.sheet_count, 1);
                assert_eq!(info.sheets, vec!["Sheet1".to_string()]);
            } else {
                panic!("Expected Workbook result");
            }
        } else {
            panic!("Expected InspectResult, got {:?}", msg);
        }

        server.stop();
    }
}
