//! Terminal panel state.
//!
//! Manages UI state for the terminal panel, mirroring the pattern from
//! `scripting/console_state.rs`.

use std::path::PathBuf;
use std::sync::{Arc, OnceLock};
use std::sync::atomic::{AtomicU64, Ordering};

use alacritty_terminal::event::WindowSize;
use alacritty_terminal::event_loop::{EventLoopSender, Msg};
use alacritty_terminal::sync::FairMutex;
use alacritty_terminal::term::Term;

use super::pty::TerminalEventProxy;

/// Default terminal panel height in pixels.
pub const DEFAULT_TERMINAL_HEIGHT: f32 = 300.0;
/// Minimum terminal panel height in pixels.
pub const MIN_TERMINAL_HEIGHT: f32 = 100.0;
/// Maximum terminal panel height in pixels.
pub const MAX_TERMINAL_HEIGHT: f32 = 600.0;

/// Terminal panel state (one per Spreadsheet window).
pub struct TerminalState {
    // Panel UI
    pub visible: bool,
    pub height: f32,
    pub is_maximized: bool,
    pub restore_height: f32,
    pub resizing: bool,
    pub resize_start_y: f32,
    pub resize_start_height: f32,

    // Terminal emulator
    pub term: Option<Arc<FairMutex<Term<TerminalEventProxy>>>>,
    pub event_loop_sender: Option<EventLoopSender>,

    // Lifecycle
    pub exited: bool,
    pub exit_code: Option<i32>,
    pub cwd: Option<PathBuf>,

    // Workspace (Phase 2)
    pub workspace_root: Option<PathBuf>,
    pub last_sent_cwd: Option<PathBuf>,

    // Output epoch: incremented on every Wakeup event (new PTY output).
    // Used for O(1) readiness checks instead of scanning the grid.
    pub output_epoch: AtomicU64,
}

impl Default for TerminalState {
    fn default() -> Self {
        Self {
            visible: false,
            height: DEFAULT_TERMINAL_HEIGHT,
            is_maximized: false,
            restore_height: DEFAULT_TERMINAL_HEIGHT,
            resizing: false,
            resize_start_y: 0.0,
            resize_start_height: 0.0,
            term: None,
            event_loop_sender: None,
            exited: false,
            exit_code: None,
            cwd: None,
            workspace_root: None,
            last_sent_cwd: None,
            output_epoch: AtomicU64::new(0),
        }
    }
}

#[allow(dead_code)]
impl TerminalState {
    /// Toggle terminal panel visibility.
    pub fn toggle(&mut self) {
        self.visible = !self.visible;
    }

    /// Show the terminal panel.
    pub fn show(&mut self) {
        self.visible = true;
    }

    /// Hide the terminal panel.
    pub fn hide(&mut self) {
        self.visible = false;
    }

    /// Toggle maximize/restore.
    pub fn toggle_maximize(&mut self, effective_max: f32) {
        if self.is_maximized {
            self.height = self.restore_height;
            self.is_maximized = false;
        } else {
            self.restore_height = self.height;
            self.height = effective_max;
            self.is_maximized = true;
        }
    }

    /// Set height from drag resize. Exits maximize mode.
    pub fn set_height_from_drag(&mut self, new_height: f32) {
        self.height = new_height.max(MIN_TERMINAL_HEIGHT).min(MAX_TERMINAL_HEIGHT);
        self.is_maximized = false;
    }

    /// Get the current output epoch. Incremented on every PTY Wakeup event.
    /// Compare snapshots to detect new output without locking the grid.
    pub fn output_epoch(&self) -> u64 {
        self.output_epoch.load(Ordering::Relaxed)
    }

    /// Bump the output epoch. Called from the event bridge on TermEvent::Wakeup.
    pub fn bump_output_epoch(&self) {
        self.output_epoch.fetch_add(1, Ordering::Relaxed);
    }

    /// Write bytes to the PTY.
    pub fn write_to_pty(&self, data: &[u8]) {
        if let Some(ref sender) = self.event_loop_sender {
            let _ = sender.send(Msg::Input(data.to_vec().into()));
        }
    }

    /// Resize the PTY to new dimensions.
    pub fn resize_pty(&self, size: WindowSize) {
        if let Some(ref sender) = self.event_loop_sender {
            let _ = sender.send(Msg::Resize(size));
        }
    }

    /// Shut down the PTY session.
    pub fn shutdown(&mut self) {
        if let Some(ref sender) = self.event_loop_sender {
            let _ = sender.send(Msg::Shutdown);
        }
        self.event_loop_sender = None;
        self.term = None;
        self.exited = true;
    }

    /// Update the workspace root. Does NOT auto-cd — call `ensure_cwd()` for that.
    pub fn set_workspace_root(&mut self, root: PathBuf) {
        self.workspace_root = Some(root);
    }

    /// If a live PTY is running and the workspace root differs from last sent CWD,
    /// send `cd "<root>"\n` to the shell. Avoids redundant cd spam.
    pub fn ensure_cwd(&mut self) {
        let Some(ref root) = self.workspace_root else { return };
        if self.term.is_none() || self.exited {
            return;
        }
        // Skip if we already sent cd to this path
        if self.last_sent_cwd.as_ref() == Some(root) {
            return;
        }
        let root_str = root.display().to_string();
        let cmd = format!("cd {}\n", shell_quote(&root_str));
        self.write_to_pty(cmd.as_bytes());
        self.last_sent_cwd = Some(root.clone());
    }
}

/// Resolve the workspace root for a given workbook file path.
///
/// Policy:
/// - Saved workbook: `<workbook_dir>/.visigrid/workspace/`
/// - Unsaved / no path: `~/.visigrid/workspaces/untitled-<YYYYMMDD-HHMMSS>/`
///
/// Creates the directory if it doesn't exist.
pub fn resolve_workspace_root(current_file: Option<&std::path::Path>) -> PathBuf {
    let root = if let Some(file) = current_file {
        if let Some(dir) = file.parent() {
            dir.join(".visigrid").join("workspace")
        } else {
            fallback_workspace()
        }
    } else {
        fallback_workspace()
    };

    // Ensure the directory exists
    if let Err(e) = std::fs::create_dir_all(&root) {
        eprintln!("Failed to create workspace directory {:?}: {}", root, e);
    }

    root
}

/// Fallback workspace for unsaved workbooks.
/// Uses a timestamped directory name, stable within a process session.
fn fallback_workspace() -> PathBuf {
    static FALLBACK: OnceLock<PathBuf> = OnceLock::new();
    FALLBACK.get_or_init(|| {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default();
        let secs = now.as_secs();
        // Format as untitled-YYYYMMDD-HHMMSS (approximate from epoch)
        // Use simple epoch-based naming to avoid chrono dependency
        let name = format!("untitled-{}", secs);
        dirs::home_dir()
            .unwrap_or_else(|| PathBuf::from("/tmp"))
            .join(".visigrid")
            .join("workspaces")
            .join(name)
    }).clone()
}

/// Shell-quote a string for safe use in `cd` commands.
/// Wraps in single quotes, escaping any embedded single quotes.
fn shell_quote(s: &str) -> String {
    if s.contains('\'') {
        format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\""))
    } else {
        format!("'{}'", s)
    }
}

impl std::fmt::Debug for TerminalState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TerminalState")
            .field("visible", &self.visible)
            .field("height", &self.height)
            .field("has_term", &self.term.is_some())
            .field("exited", &self.exited)
            .field("workspace_root", &self.workspace_root)
            .finish()
    }
}
