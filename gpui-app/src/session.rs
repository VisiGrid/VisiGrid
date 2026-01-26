//! Session persistence for VisiGrid
//!
//! Tracks and restores application state across restarts:
//! - Open windows and their files
//! - Scroll positions, selections
//! - Panel visibility, inspector tab
//! - Window bounds (position, size, maximized state)
//!
//! Session is stored in `~/.config/visigrid/session.json`.
//! Design principle: unbreakable - never crash-loop, graceful degradation.

use std::fs;
use std::path::PathBuf;
use std::time::{Duration, Instant};

use gpui::{Bounds, Global, Pixels, Point, Size, px};
use serde::{Deserialize, Serialize};

use crate::mode::InspectorTab;

// ============================================================================
// Session Data Structures
// ============================================================================

/// Global session state - tracks all windows
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Session {
    /// Version for future migrations
    pub version: u32,
    /// All open windows
    pub windows: Vec<WindowSession>,
    /// Which window was focused last (index into windows)
    pub focused_window: Option<usize>,
    /// Session metadata
    #[serde(default)]
    pub meta: SessionMeta,
}

/// Metadata about the session itself
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SessionMeta {
    /// When the session was last saved (Unix timestamp)
    #[serde(default)]
    pub last_saved: u64,
    /// Application version that saved this session
    #[serde(default)]
    pub app_version: String,
    /// Git commit SHA (if available at build time)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub git_sha: Option<String>,
    /// Platform (os/arch, e.g., "linux/x86_64")
    #[serde(default)]
    pub platform: String,
}

/// State for a single window
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WindowSession {
    /// Open file path (None = untitled)
    pub file: Option<PathBuf>,
    /// Workbook/sheet state
    pub sheets: Vec<SheetSession>,
    /// Active sheet index
    pub active_sheet: usize,
    /// Window geometry (if available)
    pub bounds: Option<SerializableBounds>,
    /// Whether window was maximized
    #[serde(default)]
    pub maximized: bool,
    /// Whether window was fullscreen
    #[serde(default)]
    pub fullscreen: bool,
    /// Panel visibility
    pub panels: PanelState,
    /// Zoom level (default 1.0 = 100%)
    #[serde(default = "default_zoom")]
    pub zoom_level: f32,
}

fn default_zoom() -> f32 {
    1.0
}

/// State for a single sheet within a workbook
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetSession {
    /// Sheet name (for correlation on restore)
    pub name: String,
    /// Viewport scroll position
    pub scroll: ScrollPosition,
    /// Selection state
    pub selection: SelectionState,
}

/// Scroll position within a sheet
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ScrollPosition {
    pub row: usize,
    pub col: usize,
}

/// Selection state within a sheet
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SelectionState {
    /// Anchor cell (single selection or range start)
    pub anchor: (usize, usize),
    /// Range end (None = single cell selection)
    pub end: Option<(usize, usize)>,
    /// Additional selections from Ctrl+Click (optional, may be empty)
    #[serde(default)]
    pub additional: Vec<((usize, usize), Option<(usize, usize)>)>,
}

/// Panel visibility state
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PanelState {
    /// Inspector panel visible
    #[serde(default)]
    pub inspector_visible: bool,
    /// Active inspector tab
    #[serde(default)]
    pub inspector_tab: SerializableInspectorTab,
    // Future: zen_mode, formula_bar_visible, etc.
}

/// Serializable wrapper for InspectorTab
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum SerializableInspectorTab {
    #[default]
    Inspector,
    Format,
    Names,
    History,
}

impl From<InspectorTab> for SerializableInspectorTab {
    fn from(tab: InspectorTab) -> Self {
        match tab {
            InspectorTab::Inspector => SerializableInspectorTab::Inspector,
            InspectorTab::Format => SerializableInspectorTab::Format,
            InspectorTab::Names => SerializableInspectorTab::Names,
            InspectorTab::History => SerializableInspectorTab::History,
        }
    }
}

impl From<SerializableInspectorTab> for InspectorTab {
    fn from(tab: SerializableInspectorTab) -> Self {
        match tab {
            SerializableInspectorTab::Inspector => InspectorTab::Inspector,
            SerializableInspectorTab::Format => InspectorTab::Format,
            SerializableInspectorTab::Names => InspectorTab::Names,
            SerializableInspectorTab::History => InspectorTab::History,
        }
    }
}

/// Serializable window bounds (GPUI's Bounds<Pixels> isn't directly serializable)
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SerializableBounds {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

impl SerializableBounds {
    pub fn from_gpui(bounds: Bounds<Pixels>) -> Self {
        Self {
            x: f32::from(bounds.origin.x),
            y: f32::from(bounds.origin.y),
            width: f32::from(bounds.size.width),
            height: f32::from(bounds.size.height),
        }
    }

    pub fn to_gpui(&self) -> Bounds<Pixels> {
        Bounds {
            origin: Point::new(px(self.x), px(self.y)),
            size: Size {
                width: px(self.width),
                height: px(self.height),
            },
        }
    }
}

// ============================================================================
// Session Persistence
// ============================================================================

const SESSION_VERSION: u32 = 1;
const SESSION_FILENAME: &str = "session.json";

/// Check if debug logging is enabled via VISIGRID_DEBUG_SESSION=1
fn debug_enabled() -> bool {
    std::env::var("VISIGRID_DEBUG_SESSION")
        .map(|v| v == "1")
        .unwrap_or(false)
}

/// Log debug message if VISIGRID_DEBUG_SESSION=1
macro_rules! session_debug {
    ($($arg:tt)*) => {
        if debug_enabled() {
            eprintln!("[session:debug] {}", format!($($arg)*));
        }
    };
}

/// Get the path to the session file
pub fn session_path() -> Option<PathBuf> {
    dirs::config_dir().map(|p| p.join("visigrid").join(SESSION_FILENAME))
}

/// Load session from disk
///
/// Returns None if:
/// - File doesn't exist (first run)
/// - File can't be read (permissions, corruption)
/// - JSON is invalid (schema mismatch, corruption)
///
/// On corruption: renames bad file to session.json.bad-<timestamp> and returns None.
/// This is intentional - session errors should NEVER prevent app startup.
pub fn load_session() -> Option<Session> {
    let path = session_path()?;
    session_debug!("load_session: path={:?}", path);

    let content = match fs::read_to_string(&path) {
        Ok(c) => {
            session_debug!("load_session: read {} bytes", c.len());
            c
        }
        Err(e) => {
            if e.kind() != std::io::ErrorKind::NotFound {
                eprintln!("[session] Failed to read {:?}: {}", path, e);
            } else {
                session_debug!("load_session: file not found (first run)");
            }
            return None;
        }
    };

    match serde_json::from_str::<Session>(&content) {
        Ok(session) => {
            // Version check - if major incompatibility, backup and return None
            if session.version > SESSION_VERSION {
                eprintln!(
                    "[session] Session version {} is newer than supported {}",
                    session.version, SESSION_VERSION
                );
                backup_corrupt_session(&path);
                return None;
            }
            session_debug!(
                "load_session: loaded v{} with {} windows, meta={{app_version={}, platform={}, git_sha={:?}}}",
                session.version,
                session.windows.len(),
                session.meta.app_version,
                session.meta.platform,
                session.meta.git_sha
            );
            Some(session)
        }
        Err(e) => {
            eprintln!("[session] Failed to parse session JSON: {}", e);
            backup_corrupt_session(&path);
            None
        }
    }
}

/// Backup a corrupt session file to session.json.bad-<timestamp>
fn backup_corrupt_session(path: &std::path::Path) {
    let timestamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);

    let backup_name = format!("{}.bad-{}", path.display(), timestamp);
    let backup_path = std::path::PathBuf::from(&backup_name);
    session_debug!("backup_corrupt_session: backing up to {}", backup_name);

    if let Err(e) = fs::rename(path, &backup_path) {
        eprintln!("[session] Failed to backup corrupt session: {}", e);
        // If backup fails, try to delete so we don't crash-loop
        let _ = fs::remove_file(path);
    } else {
        eprintln!("[session] Corrupt session backed up to: {}", backup_name);
    }
}

/// Save session to disk (atomic write)
///
/// Uses write-to-temp-then-rename pattern to prevent corruption.
/// On Windows, backs up existing file before removing since rename doesn't overwrite.
/// Silently handles errors - session save failures are not critical.
pub fn save_session(session: &Session) {
    let Some(path) = session_path() else {
        return;
    };
    session_debug!("save_session: path={:?}", path);

    // Ensure config directory exists
    if let Some(parent) = path.parent() {
        if let Err(e) = fs::create_dir_all(parent) {
            eprintln!("[session] Failed to create config dir: {}", e);
            return;
        }
    }

    // Serialize with pretty printing for debuggability
    let json = match serde_json::to_string_pretty(session) {
        Ok(j) => j,
        Err(e) => {
            eprintln!("[session] Failed to serialize session: {}", e);
            return;
        }
    };
    session_debug!("save_session: serialized {} bytes, {} windows", json.len(), session.windows.len());

    // Atomic write: temp file + rename
    let temp_path = path.with_extension("json.tmp");
    if let Err(e) = fs::write(&temp_path, &json) {
        eprintln!("[session] Failed to write temp file: {}", e);
        return;
    }

    // On Windows, fs::rename fails if target exists.
    // Backup existing to .bak, then remove. If rename succeeds, delete .bak.
    // If rename fails, restore from .bak.
    #[cfg(target_os = "windows")]
    {
        let backup_path = path.with_extension("json.bak");
        let had_existing = path.exists();

        if had_existing {
            // Backup existing session before removing
            if let Err(e) = fs::rename(&path, &backup_path) {
                session_debug!("save_session: backup failed: {}", e);
                // Can't backup? Try direct remove (old behavior)
                let _ = fs::remove_file(&path);
            }
        }

        if let Err(e) = fs::rename(&temp_path, &path) {
            eprintln!("[session] Failed to rename temp to session: {}", e);
            // Restore from backup if we made one
            if had_existing && backup_path.exists() {
                let _ = fs::rename(&backup_path, &path);
                session_debug!("save_session: restored from backup");
            }
            let _ = fs::remove_file(&temp_path);
            return;
        }

        // Success - clean up backup
        if had_existing {
            let _ = fs::remove_file(&backup_path);
        }
        session_debug!("save_session: saved successfully (Windows path)");
    }

    #[cfg(not(target_os = "windows"))]
    {
        if let Err(e) = fs::rename(&temp_path, &path) {
            eprintln!("[session] Failed to rename temp to session: {}", e);
            let _ = fs::remove_file(&temp_path);
            return;
        }
        session_debug!("save_session: saved successfully");
    }
}

/// Explicitly reset (delete) session file.
/// Used by --reset-session CLI flag.
/// This is the ONLY function that should delete the session file.
pub fn reset_session() {
    if let Some(path) = session_path() {
        match fs::remove_file(&path) {
            Ok(()) => eprintln!("[session] Session file deleted: {:?}", path),
            Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
                // Already gone, that's fine
            }
            Err(e) => eprintln!("[session] Failed to delete session: {}", e),
        }
    }
}

/// Dump session to stdout (for --dump-session debugging)
pub fn dump_session() {
    match load_session() {
        Some(session) => {
            println!("{}", serde_json::to_string_pretty(&session).unwrap_or_default());
        }
        None => {
            eprintln!("No session file found or session is invalid.");
            std::process::exit(1);
        }
    }
}

// ============================================================================
// Session Manager (debounced autosave)
// ============================================================================

/// Manages session autosave with debouncing
pub struct SessionManager {
    /// Current session state
    session: Session,
    /// Has the session been modified since last save?
    dirty: bool,
    /// When was the session last marked dirty?
    dirty_since: Option<Instant>,
    /// Debounce interval
    debounce: Duration,
}

impl SessionManager {
    /// Create a new session manager
    pub fn new() -> Self {
        let session = load_session().unwrap_or_else(|| Session {
            version: SESSION_VERSION,
            ..Default::default()
        });

        Self {
            session,
            dirty: false,
            dirty_since: None,
            debounce: Duration::from_secs(5),
        }
    }

    /// Create with custom debounce (for testing)
    pub fn with_debounce(debounce: Duration) -> Self {
        let mut mgr = Self::new();
        mgr.debounce = debounce;
        mgr
    }

    /// Create a fresh session manager for testing (does NOT load from disk)
    #[cfg(test)]
    pub fn new_empty_for_test(debounce: Duration) -> Self {
        Self {
            session: Session {
                version: SESSION_VERSION,
                ..Default::default()
            },
            dirty: false,
            dirty_since: None,
            debounce,
        }
    }

    /// Get the current session (read-only)
    pub fn session(&self) -> &Session {
        &self.session
    }

    /// Get mutable session and mark dirty
    pub fn session_mut(&mut self) -> &mut Session {
        self.mark_dirty();
        &mut self.session
    }

    /// Mark session as dirty (needs save)
    pub fn mark_dirty(&mut self) {
        if !self.dirty {
            self.dirty = true;
            self.dirty_since = Some(Instant::now());
        }
    }

    /// Check if we should save (debounce elapsed)
    pub fn should_save(&self) -> bool {
        if !self.dirty {
            return false;
        }
        match self.dirty_since {
            Some(since) => since.elapsed() >= self.debounce,
            None => false,
        }
    }

    /// Save if debounce has elapsed
    pub fn maybe_save(&mut self) {
        if self.should_save() {
            self.save_now();
        }
    }

    /// Force save immediately (for quit, window close)
    pub fn save_now(&mut self) {
        // Update metadata
        self.session.meta.last_saved = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs())
            .unwrap_or(0);
        self.session.meta.app_version = env!("CARGO_PKG_VERSION").to_string();
        self.session.meta.git_sha = option_env!("VISIGRID_GIT_SHA").map(|s| s.to_string());
        self.session.meta.platform = format!("{}/{}", std::env::consts::OS, std::env::consts::ARCH);

        save_session(&self.session);
        self.dirty = false;
        self.dirty_since = None;
    }

    /// Clear session (for fresh start)
    pub fn clear(&mut self) {
        self.session = Session {
            version: SESSION_VERSION,
            ..Default::default()
        };
        self.dirty = true;
        self.dirty_since = Some(Instant::now());
    }
}

impl Default for SessionManager {
    fn default() -> Self {
        Self::new()
    }
}

impl Global for SessionManager {}

// ============================================================================
// Snapshot & Apply (the core logic)
// ============================================================================

use crate::app::Spreadsheet;
use gpui::{Window, WindowBounds};

impl Spreadsheet {
    /// Capture current state as a WindowSession snapshot
    pub fn snapshot(&self, window: &Window) -> WindowSession {
        self.snapshot_with_bounds(Some(window.window_bounds()))
    }

    /// Capture current state using cached window bounds (for use without Window access)
    pub fn snapshot_cached(&self) -> WindowSession {
        self.snapshot_with_bounds(self.cached_window_bounds)
    }

    /// Internal: create snapshot with given window bounds
    fn snapshot_with_bounds(&self, window_bounds: Option<WindowBounds>) -> WindowSession {
        // Capture sheet states
        let sheets: Vec<SheetSession> = self.workbook.sheets()
            .iter()
            .enumerate()
            .map(|(idx, sheet)| {
                // For the active sheet, use current viewport state
                // For other sheets, we'd need per-sheet scroll/selection tracking
                // (not implemented yet - for now, only active sheet has accurate state)
                if idx == self.workbook.active_sheet_index() {
                    SheetSession {
                        name: sheet.name.clone(),
                        scroll: ScrollPosition {
                            row: self.view_state.scroll_row,
                            col: self.view_state.scroll_col,
                        },
                        selection: SelectionState {
                            anchor: self.view_state.selected,
                            end: self.view_state.selection_end,
                            additional: self.view_state.additional_selections.clone(),
                        },
                    }
                } else {
                    // Other sheets: just name, defaults for rest
                    SheetSession {
                        name: sheet.name.clone(),
                        ..Default::default()
                    }
                }
            })
            .collect();

        // Capture window bounds from provided or default
        let (bounds, maximized, fullscreen) = match window_bounds {
            Some(WindowBounds::Windowed(b)) => (Some(SerializableBounds::from_gpui(b)), false, false),
            Some(WindowBounds::Maximized(b)) => (Some(SerializableBounds::from_gpui(b)), true, false),
            Some(WindowBounds::Fullscreen(b)) => (Some(SerializableBounds::from_gpui(b)), false, true),
            None => (None, false, false),
        };

        WindowSession {
            file: self.current_file.clone(),
            sheets,
            active_sheet: self.workbook.active_sheet_index(),
            bounds,
            maximized,
            fullscreen,
            panels: PanelState {
                inspector_visible: self.inspector_visible,
                inspector_tab: self.inspector_tab.into(),
            },
            zoom_level: self.view_state.zoom_level,
        }
    }

    /// Apply a WindowSession to restore state
    ///
    /// Call this after loading the file but before first render.
    /// Does NOT load the file - caller must do that first.
    pub fn apply(&mut self, session: &WindowSession) {
        // Find matching sheet by name (handles sheet reordering)
        let active_sheet_name = session.sheets
            .get(session.active_sheet)
            .map(|s| s.name.as_str());

        // Switch to active sheet if found
        if let Some(name) = active_sheet_name {
            if let Some(idx) = self.workbook.sheets()
                .iter()
                .position(|s| s.name == name)
            {
                self.workbook.set_active_sheet(idx);
            }
        }

        // Apply scroll and selection for active sheet
        let active_idx = self.workbook.active_sheet_index();
        if let Some(sheet_session) = session.sheets.get(active_idx) {
            // Clamp to valid range (sheet may have shrunk)
            let sheet = self.workbook.active_sheet();
            let max_row = sheet.rows;
            let max_col = sheet.cols;

            self.view_state.scroll_row = sheet_session.scroll.row.min(max_row.saturating_sub(1));
            self.view_state.scroll_col = sheet_session.scroll.col.min(max_col.saturating_sub(1));

            self.view_state.selected = (
                sheet_session.selection.anchor.0.min(max_row.saturating_sub(1)),
                sheet_session.selection.anchor.1.min(max_col.saturating_sub(1)),
            );

            self.view_state.selection_end = sheet_session.selection.end.map(|(r, c)| {
                (r.min(max_row.saturating_sub(1)), c.min(max_col.saturating_sub(1)))
            });

            // Restore additional selections (clamped)
            self.view_state.additional_selections = sheet_session.selection.additional
                .iter()
                .map(|(anchor, end)| {
                    let clamped_anchor = (
                        anchor.0.min(max_row.saturating_sub(1)),
                        anchor.1.min(max_col.saturating_sub(1)),
                    );
                    let clamped_end = end.map(|(r, c)| {
                        (r.min(max_row.saturating_sub(1)), c.min(max_col.saturating_sub(1)))
                    });
                    (clamped_anchor, clamped_end)
                })
                .collect();
        }

        // Apply panel state
        self.inspector_visible = session.panels.inspector_visible;
        self.inspector_tab = session.panels.inspector_tab.into();

        // Apply zoom level (clamp to valid range)
        use crate::app::{ZOOM_STEPS, GridMetrics};
        let zoom = session.zoom_level
            .max(ZOOM_STEPS[0])
            .min(ZOOM_STEPS[ZOOM_STEPS.len() - 1]);
        self.view_state.zoom_level = zoom;
        self.metrics = GridMetrics::new(zoom);
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_session_roundtrip() {
        let session = Session {
            version: SESSION_VERSION,
            windows: vec![
                WindowSession {
                    file: Some(PathBuf::from("/tmp/test.sheet")),
                    sheets: vec![
                        SheetSession {
                            name: "Sheet1".to_string(),
                            scroll: ScrollPosition { row: 10, col: 5 },
                            selection: SelectionState {
                                anchor: (10, 5),
                                end: Some((20, 10)),
                                additional: vec![],
                            },
                        },
                    ],
                    active_sheet: 0,
                    bounds: Some(SerializableBounds {
                        x: 100.0,
                        y: 100.0,
                        width: 1200.0,
                        height: 800.0,
                    }),
                    maximized: false,
                    fullscreen: false,
                    panels: PanelState {
                        inspector_visible: true,
                        inspector_tab: SerializableInspectorTab::Format,
                    },
                    zoom_level: 1.0,
                },
            ],
            focused_window: Some(0),
            meta: SessionMeta::default(),
        };

        let json = serde_json::to_string_pretty(&session).unwrap();
        let parsed: Session = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.version, SESSION_VERSION);
        assert_eq!(parsed.windows.len(), 1);
        assert_eq!(parsed.windows[0].sheets[0].scroll.row, 10);
        assert_eq!(parsed.windows[0].panels.inspector_visible, true);
    }

    #[test]
    fn test_missing_fields_use_defaults() {
        // Simulate older session format missing new fields
        let json = r#"{
            "version": 1,
            "windows": [{
                "file": null,
                "sheets": [],
                "active_sheet": 0,
                "panels": {}
            }],
            "focused_window": null
        }"#;

        let session: Session = serde_json::from_str(json).unwrap();
        assert_eq!(session.windows[0].maximized, false);
        assert_eq!(session.windows[0].fullscreen, false);
    }

    #[test]
    fn test_serializable_bounds_conversion() {
        let gpui_bounds = Bounds {
            origin: Point::new(px(100.0), px(200.0)),
            size: Size {
                width: px(800.0),
                height: px(600.0),
            },
        };

        let serializable = SerializableBounds::from_gpui(gpui_bounds);
        assert_eq!(serializable.x, 100.0);
        assert_eq!(serializable.y, 200.0);
        assert_eq!(serializable.width, 800.0);
        assert_eq!(serializable.height, 600.0);

        let back = serializable.to_gpui();
        assert_eq!(f32::from(back.origin.x), 100.0);
        assert_eq!(f32::from(back.size.height), 600.0);
    }

    #[test]
    fn test_session_manager_debounce() {
        let mut mgr = SessionManager::with_debounce(Duration::from_millis(100));

        // Initially not dirty
        assert!(!mgr.should_save());

        // Mark dirty
        mgr.mark_dirty();
        assert!(!mgr.should_save()); // Not yet - debounce hasn't elapsed

        // Wait for debounce
        std::thread::sleep(Duration::from_millis(150));
        assert!(mgr.should_save());
    }

    #[test]
    fn test_corrupt_json_is_tolerated() {
        // Various forms of invalid JSON that should not panic
        let invalid_jsons = vec![
            "",                           // Empty
            "{",                          // Truncated
            "null",                       // Wrong type
            "[]",                         // Array instead of object
            "{\"version\": \"bad\"}",     // Wrong type for version
            "{\"version\": 999}",         // Future version
            "garbage bytes \x00\x01\x02", // Binary garbage
        ];

        for json in invalid_jsons {
            // Should return None, not panic
            let result: Result<Session, _> = serde_json::from_str(json);
            // Either fails to parse or version check would fail
            assert!(result.is_err() || result.as_ref().map(|s| s.version > SESSION_VERSION).unwrap_or(false),
                "Should reject invalid JSON: {}", json);
        }
    }

    #[test]
    fn test_session_unknown_fields_ignored() {
        // Future sessions might have extra fields - they should be ignored
        let json = r#"{
            "version": 1,
            "windows": [],
            "focused_window": null,
            "meta": {},
            "future_field_1": "ignored",
            "future_field_2": [1, 2, 3]
        }"#;

        let session: Session = serde_json::from_str(json).unwrap();
        assert_eq!(session.version, 1);
        assert!(session.windows.is_empty());
    }

    #[test]
    fn test_window_session_unknown_fields_ignored() {
        // Nested objects also tolerate unknown fields
        let json = r#"{
            "version": 1,
            "windows": [{
                "file": null,
                "sheets": [],
                "active_sheet": 0,
                "panels": {},
                "future_window_field": "ignored"
            }],
            "focused_window": null
        }"#;

        let session: Session = serde_json::from_str(json).unwrap();
        assert_eq!(session.windows.len(), 1);
    }

    #[test]
    fn test_selection_state_clamping_logic() {
        // Test the clamping math directly
        let max_row: usize = 100;
        let max_col: usize = 50;

        // Values within bounds stay the same
        assert_eq!(50usize.min(max_row.saturating_sub(1)), 50);
        assert_eq!(25usize.min(max_col.saturating_sub(1)), 25);

        // Values beyond bounds get clamped
        assert_eq!(200usize.min(max_row.saturating_sub(1)), 99);
        assert_eq!(100usize.min(max_col.saturating_sub(1)), 49);

        // Edge case: empty sheet (0 rows/cols)
        let empty_max: usize = 0;
        assert_eq!(50usize.min(empty_max.saturating_sub(1)), 0);
    }

    #[test]
    fn test_session_with_multiple_windows() {
        let session = Session {
            version: SESSION_VERSION,
            windows: vec![
                WindowSession {
                    file: Some(PathBuf::from("/tmp/file1.sheet")),
                    sheets: vec![SheetSession {
                        name: "Sheet1".to_string(),
                        ..Default::default()
                    }],
                    ..Default::default()
                },
                WindowSession {
                    file: Some(PathBuf::from("/tmp/file2.sheet")),
                    sheets: vec![SheetSession {
                        name: "Data".to_string(),
                        scroll: ScrollPosition { row: 100, col: 20 },
                        ..Default::default()
                    }],
                    ..Default::default()
                },
            ],
            focused_window: Some(1),
            ..Default::default()
        };

        let json = serde_json::to_string_pretty(&session).unwrap();
        let parsed: Session = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.windows.len(), 2);
        assert_eq!(parsed.focused_window, Some(1));
        assert_eq!(parsed.windows[1].sheets[0].scroll.row, 100);
    }

    #[test]
    fn test_additional_selections_roundtrip() {
        let selection = SelectionState {
            anchor: (5, 5),
            end: Some((10, 10)),
            additional: vec![
                ((15, 15), None),           // Single cell
                ((20, 5), Some((25, 10))),  // Range
            ],
        };

        let json = serde_json::to_string(&selection).unwrap();
        let parsed: SelectionState = serde_json::from_str(&json).unwrap();

        assert_eq!(parsed.anchor, (5, 5));
        assert_eq!(parsed.end, Some((10, 10)));
        assert_eq!(parsed.additional.len(), 2);
        assert_eq!(parsed.additional[0], ((15, 15), None));
        assert_eq!(parsed.additional[1], ((20, 5), Some((25, 10))));
    }

    #[test]
    fn test_inspector_tab_serialization() {
        // Test all tab variants serialize correctly
        let tabs = vec![
            (SerializableInspectorTab::Inspector, "\"inspector\""),
            (SerializableInspectorTab::Format, "\"format\""),
            (SerializableInspectorTab::Names, "\"names\""),
        ];

        for (tab, expected) in tabs {
            let json = serde_json::to_string(&tab).unwrap();
            assert_eq!(json, expected);

            let parsed: SerializableInspectorTab = serde_json::from_str(&json).unwrap();
            assert_eq!(parsed, tab);
        }
    }

    #[test]
    fn test_session_manager_clear() {
        // Use test-only constructor that doesn't load from disk
        let mut mgr = SessionManager::new_empty_for_test(Duration::from_secs(5));

        // Add some state
        mgr.session_mut().windows.push(WindowSession::default());
        assert_eq!(mgr.session().windows.len(), 1);

        // Clear should reset to empty
        mgr.clear();
        assert_eq!(mgr.session().windows.len(), 0);
        assert_eq!(mgr.session().version, SESSION_VERSION);
        assert!(mgr.dirty); // Should be marked dirty
    }
}
