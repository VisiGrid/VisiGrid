//! Native Go collaboration worker. No socket or HTTP call runs on the UI thread.
use crate::hub::auth::AuthCredentials;
use serde::Deserialize;
use std::sync::{
    atomic::{AtomicBool, Ordering},
    mpsc, Arc,
};
use visigrid_collab::{
    socket::LiveSocket,
    wire::{Received, WireReplica},
};
use visigrid_engine::workbook::Workbook;

pub enum Command {
    Cell {
        sheet: u64,
        row: usize,
        col: usize,
        raw: String,
    },
}
pub enum Event {
    State {
        ready: bool,
        writable: bool,
        message: String,
    },
    Workbook {
        workbook: Workbook,
        pending: usize,
    },
}
pub struct LiveSession {
    pub commands: mpsc::Sender<Command>,
    pub events: mpsc::Receiver<Event>,
    stop: Arc<AtomicBool>,
}
impl Drop for LiveSession {
    fn drop(&mut self) {
        self.stop.store(true, Ordering::Relaxed);
    }
}
#[derive(Deserialize)]
struct Metadata {
    role: String,
}
#[derive(Deserialize)]
struct Snapshot {
    revision: u64,
    document: Box<serde_json::value::RawValue>,
}

impl LiveSession {
    pub fn start(auth: AuthCredentials, pid: String, initial: Workbook) -> Self {
        let (send, commands) = mpsc::channel();
        let (events, receive) = mpsc::channel();
        let stop = Arc::new(AtomicBool::new(false));
        let stopping = stop.clone();
        std::thread::spawn(move || {
            if let Err(message) = run(auth, pid, initial, commands, &events, &stopping) {
                let _ = events.send(Event::State {
                    ready: false,
                    writable: false,
                    message,
                });
            }
        });
        Self {
            commands: send,
            events: receive,
            stop,
        }
    }
}
fn run(
    auth: AuthCredentials,
    pid: String,
    initial: Workbook,
    commands: mpsc::Receiver<Command>,
    events: &mpsc::Sender<Event>,
    stop: &AtomicBool,
) -> Result<(), String> {
    let base = reqwest::Url::parse(&auth.api_base).map_err(|_| "Invalid API origin")?;
    visigrid_collab::socket::validate_api_origin(&auth.api_base).map_err(|error| error.to_string())?;
    let pid = uuid::Uuid::parse_str(&pid)
        .map_err(|_| "A Go workbook UUID is required for live editing")?
        .to_string();
    let http = reqwest::blocking::Client::builder()
        .redirect(reqwest::redirect::Policy::none())
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .map_err(|_| "Could not create API client")?;
    let metadata_url = base
        .join(&format!("/api/sheets/{pid}"))
        .map_err(|_| "Invalid workbook URL")?;
    let metadata: Metadata = http
        .get(metadata_url)
        .bearer_auth(&auth.token)
        .send()
        .and_then(reqwest::blocking::Response::error_for_status)
        .map_err(|_| "Could not authorize live workbook access")?
        .json()
        .map_err(|_| "Invalid workbook permissions")?;
    let writable = matches!(metadata.role.as_str(), "owner" | "editor");
    initial.ensure_writable()?;
    let mut replica = WireReplica::new(initial, 0, writable);
    let mut sequence_known = false;
    let mut refusal: Option<String> = None;
    let mut socket_url = base
        .join(&format!("/api/sheets/{pid}/collab"))
        .map_err(|_| "Invalid collaboration URL")?;
    socket_url
        .set_scheme(if base.scheme() == "https" {
            "wss"
        } else {
            "ws"
        })
        .map_err(|_| "Invalid collaboration scheme")?;
    let mut backoff = visigrid_collab::socket::ReconnectBackoff::default();
    let snapshot_path = format!("/api/sheets/{pid}/snapshot");
    while !stop.load(Ordering::Relaxed) {
        let _ = events.send(Event::State {
            ready: false,
            writable,
            message: "Connecting to shared workbook…".into(),
        });
        let mut socket = match LiveSocket::connect(socket_url.as_str(), &auth.token) {
            Ok(socket) => socket,
            Err(message) => {
                // Authentication and protocol refusals need user action;
                // transient connection failures must preserve pending edits.
                if message.is_fatal() { return Err(message.to_string()); }
                replica.disconnect();
                let _ = events.send(Event::State {
                    ready: false,
                    writable,
                    message: "Service unavailable — reconnecting to shared workbook…".into(),
                });
                reconnect_pause(stop, &mut backoff);
                continue;
            }
        };
        if socket
            .send(&replica.hello(env!("VISIGRID_ENGINE_COMMIT"), sequence_known))
            .is_err()
        {
            replica.disconnect();
            reconnect_pause(stop, &mut backoff);
            continue;
        }
        let mut ready = false;
        while !stop.load(Ordering::Relaxed) {
            if ready {
                for command in commands.try_iter() {
                    refusal = None;
                    match command {
                        Command::Cell {
                            sheet,
                            row,
                            col,
                            raw,
                        } => replica.set_cell(uuid::Uuid::new_v4(), sheet, row, col, raw)?,
                    }
                    let _ = events.send(Event::Workbook {
                        workbook: replica.client.wb.clone(),
                        pending: replica.client.pending_count(),
                    });
                }
                if let Some(frame) = replica.poll_send() {
                    if socket.send(&frame).is_err() {
                        break;
                    }
                }
            }
            let frame = match socket.read() {
                Ok(Some(frame)) => frame,
                Ok(None) => continue,
                Err(_) => break,
            };
            if frame["type"] == "welcome" { backoff.reset(); }
            match replica.receive(&frame)? {
                Received::Snapshot(request) => {
                    let mut url = base
                        .join(&request.url)
                        .map_err(|_| "Invalid snapshot URL")?;
                    if url.origin() != base.origin()
                        || (url.path() != snapshot_path
                            && url.path() != format!("/api/sheets/{pid}/download"))
                    {
                        return Err("Unexpected snapshot origin or workbook".into());
                    }
                    // Download may advertise an R2 transport. Use the API's
                    // canonical inline snapshot without forwarding credentials.
                    url.set_path(&snapshot_path);
                    url.set_query(None);
                    let text = http
                        .get(url)
                        .bearer_auth(&auth.token)
                        .send()
                        .and_then(reqwest::blocking::Response::error_for_status)
                        .map_err(|_| "Could not load collaboration snapshot")?
                        .text()
                        .map_err(|_| "Could not read collaboration snapshot")?;
                    let snapshot: Snapshot = serde_json::from_str(&text)
                        .map_err(|_| "Invalid collaboration snapshot")?;
                    if request.revision.is_some_and(|r| r != snapshot.revision) {
                        return Err(
                            "The snapshot changed during startup; reopen the shared workbook"
                                .into(),
                        );
                    }
                    let (mut workbook, layouts, _) =
                        visigrid_io::json::import_any(snapshot.document.get())?;
                    for (index, layout) in layouts.iter().enumerate() {
                        if let Some(sheet) = workbook.sheet_mut(index) {
                            sheet.layout = layout.line_layout();
                        }
                    }
                    if let Err(message) = workbook.ensure_writable() {
                        let _ = events.send(Event::Workbook {
                            workbook,
                            pending: 0,
                        });
                        return Err(message);
                    }
                    replica.load_snapshot(workbook, request.seq)?;
                    sequence_known = true;
                    ready = true;
                }
                Received::Reconnect => break,
                Received::Updated => {
                    if frame["type"] == "welcome" {
                        sequence_known = true;
                        ready = true;
                    }
                }
            }
            if frame["type"] == "rejected" {
                refusal = Some(format!(
                    "Edit was not committed: {}",
                    frame
                        .get("reason")
                        .and_then(serde_json::Value::as_str)
                        .filter(|reason| !reason.is_empty())
                        .unwrap_or("the server refused this edit")
                ));
            }
            let _ = events.send(Event::Workbook {
                workbook: replica.client.wb.clone(),
                pending: replica.client.pending_count(),
            });
            if ready {
                let _ = events.send(Event::State {
                    ready: true,
                    writable,
                    message: refusal.clone().unwrap_or_else(|| {
                        if replica.client.pending_count() == 0 {
                            "Changes committed"
                        } else {
                            "Waiting for acknowledgement"
                        }
                        .into()
                    }),
                });
            }
        }
        replica.disconnect();
        let _ = events.send(Event::State {
            ready: false,
            writable,
            message: "Disconnected — reconnecting to shared workbook…".into(),
        });
        reconnect_pause(stop, &mut backoff);
    }
    Ok(())
}

fn read_private_live_auth(path: &std::path::Path) -> std::io::Result<String> {
    let mut file = std::fs::File::open(path)?;
    #[cfg(unix)] {
        use std::os::unix::fs::PermissionsExt;
        if file.metadata()?.permissions().mode() & 0o077 != 0 {
            return Err(std::io::Error::new(std::io::ErrorKind::PermissionDenied, "Live auth file must be private"));
        }
    }
    use std::io::Read;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    Ok(contents)
}

fn reconnect_pause(stop: &AtomicBool, backoff: &mut visigrid_collab::socket::ReconnectBackoff) {
    let delay = backoff.next_delay(uuid::Uuid::new_v4().as_u128() as u64);
    let until = std::time::Instant::now() + delay;
    while !stop.load(Ordering::Relaxed) {
        let Some(remaining) = until.checked_duration_since(std::time::Instant::now()) else { break; };
        std::thread::sleep(remaining.min(std::time::Duration::from_millis(50)));
    }
}

impl crate::app::Spreadsheet {
    pub(crate) fn cloud_live_enabled(&self) -> bool {
        std::env::var("VISIGRID_LIVE_COLLAB").is_ok_and(|v| v == "1")
            && self.cloud_identity.is_some()
    }
    pub(crate) fn cloud_live_stop(&mut self) {
        self.cloud_live_generation += 1;
        self.cloud_live = None;
        self.cloud_live_ready = false;
        self.cloud_live_writable = false;
    }
    pub(crate) fn cloud_live_start(&mut self, cx: &mut gpui::Context<Self>) {
        if !self.cloud_live_enabled() {
            return;
        }
        self.cloud_live_stop();
        let Some(identity) = self.cloud_identity.clone() else {
            return;
        };
        // A disposable QA backend can use its own credentials without changing
        // the user's saved sign-in. An invalid override must never fall back to
        // production credentials.
        let auth = match std::env::var_os("VISIGRID_LIVE_AUTH_FILE") {
            Some(path) => read_private_live_auth(std::path::Path::new(&path))
                .ok()
                .and_then(|contents| serde_json::from_str::<AuthCredentials>(&contents).ok()),
            None => crate::hub::auth::load_auth(),
        };
        let Some(auth) = auth else {
            self.status_message = Some("Sign in or provide valid live test credentials".into());
            cx.notify();
            return;
        };
        if auth.api_base.trim_end_matches('/') != identity.api_base.trim_end_matches('/') {
            self.status_message =
                Some("Sign in to this workbook's API before starting collaboration".into());
            cx.notify();
            return;
        }
        if let Err(error) = self.wb(cx).ensure_writable() {
            self.status_message = Some(error);
            cx.notify();
            return;
        }
        self.cloud_live = Some(LiveSession::start(
            auth,
            identity.public_id.clone(),
            self.wb(cx).clone(),
        ));
        self.status_message = Some("Connecting to shared workbook…".into());
        let generation = self.cloud_live_generation;
        cx.spawn(async move |this, cx| loop {
            smol::Timer::after(std::time::Duration::from_millis(16)).await;
            let keep = this.update(cx, |this, cx| {
                if this.cloud_live_generation != generation {
                    return false;
                }
                let Some(session) = &this.cloud_live else {
                    return false;
                };
                let events: Vec<_> = session.events.try_iter().collect();
                for event in events {
                    match event {
                        Event::State {
                            ready,
                            writable,
                            message,
                        } => {
                            this.cloud_live_ready = ready;
                            this.cloud_live_writable = writable;
                            this.status_message = Some(message);
                        }
                        Event::Workbook {
                            mut workbook,
                            pending,
                        } => {
                            let active = this.wb(cx).active_sheet().id;
                            if let Some(index) = workbook.idx_for_sheet_id(active) {
                                workbook.set_active_sheet(index);
                            }
                            this.workbook.update(cx, |wb, _| *wb = workbook);
                            this.update_cached_sheet_id(cx);
                            this.status_message = Some(
                                if pending == 0 {
                                    "Changes committed"
                                } else {
                                    "Waiting for acknowledgement"
                                }
                                .into(),
                            );
                        }
                    }
                    cx.notify();
                }
                true
            });
            if !matches!(keep, Ok(true)) {
                break;
            }
        })
        .detach();
        cx.notify();
    }
    pub(crate) fn cloud_live_cell(
        &mut self,
        row: usize,
        col: usize,
        raw: &str,
        cx: &mut gpui::Context<Self>,
    ) -> bool {
        if !self.cloud_live_enabled() {
            return false;
        }
        if !self.cloud_live_ready || !self.cloud_live_writable {
            self.status_message = Some("The live workbook is read-only or reconnecting".into());
            cx.notify();
            return true;
        }
        let sheet = self.wb(cx).active_sheet().id.0;
        if let Some(session) = &self.cloud_live {
            if session
                .commands
                .send(Command::Cell {
                    sheet,
                    row,
                    col,
                    raw: raw.into(),
                })
                .is_err()
            {
                self.cloud_live_ready = false;
                self.status_message =
                    Some("Live collaboration stopped; reopen the workbook to reconnect".into());
            }
        }
        cx.notify();
        true
    }
    pub(crate) fn block_live_read_only(&mut self, cx: &mut gpui::Context<Self>) -> bool {
        if self.cloud_live_enabled() && (!self.cloud_live_ready || !self.cloud_live_writable) {
            self.status_message = Some("The live workbook is read-only or reconnecting".into());
            cx.notify();
            true
        } else {
            false
        }
    }
}
