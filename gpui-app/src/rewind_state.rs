//! Soft-rewind preview + rewind confirmation dialog state.
//! Extracted from app.rs 2026-07-29 (pure move).

use gpui::*;
use crate::ai_dialog_state::chrono_lite_utc;
use std::collections::HashMap;
use std::path::PathBuf;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::sheet::SheetId;
use crate::history::HistoryFingerprint;
use crate::app::{DocumentSource, next_book_name};

// ============================================================================
// Soft-Rewind Preview (Phase 8A)
// ============================================================================

/// State machine for soft-rewind preview
#[derive(Clone, Debug, Default)]
pub enum RewindPreviewState {
    /// No preview active, no entry armed
    #[default]
    Off,
    /// Preview is active - showing historical state
    On(RewindPreviewSession),
}

/// Quality indicator for a preview build.
/// Degraded previews should block hard rewind.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PreviewQuality {
    /// Preview is complete and trustworthy
    Ok,
    /// Preview was truncated or had issues - block rewind
    Degraded(String),
}

/// Active preview session data
#[derive(Clone, Debug)]
pub struct RewindPreviewSession {
    /// The history entry we're previewing "before"
    pub entry_id: u64,
    /// Global index in canonical history (not filtered list)
    pub target_global_index: usize,
    /// Summary of the action for banner display
    pub action_summary: String,
    /// The reconstructed workbook state (before the action)
    pub snapshot: Workbook,
    /// Preview view state (row order + sort per sheet)
    pub view_state: PreviewViewState,
    /// Live state to restore on exit
    pub live_focus: PreviewFocus,
    /// History fingerprint at preview time (for detecting concurrent changes).
    /// 128-bit blake3 hash ensures collision resistance.
    pub history_fingerprint: HistoryFingerprint,
    /// Number of actions that were replayed to build the preview
    pub replay_count: usize,
    /// Time spent building the preview (milliseconds)
    pub build_ms: u64,
    /// Preview quality (Ok or Degraded with reason)
    pub quality: PreviewQuality,
}

/// Plan for atomic hard rewind operation.
/// Built before commit, applied atomically.
#[derive(Clone, Debug)]
pub struct RewindPlan {
    /// The workbook state to restore
    pub new_workbook: Workbook,
    /// View state (row order + sort per sheet)
    pub new_view_state: PreviewViewState,
    /// Where to truncate history (entries [0..truncate_at) are kept)
    pub truncate_at: usize,
    /// The rewind audit action to append
    pub audit_action: crate::history::UndoAction,
    /// Number of entries being discarded
    pub discarded_count: usize,
    /// Focus to restore after rewind
    pub focus: PreviewFocus,
}

/// Preview-only view state (row order + sort per sheet)
/// Lightweight alternative to snapshotting full app state
#[derive(Clone, Debug, Default)]
pub struct PreviewViewState {
    pub per_sheet: Vec<PreviewSheetView>,
}

/// Per-sheet view state for preview rendering
#[derive(Clone, Debug, Default)]
pub struct PreviewSheetView {
    /// Row order permutation (None = identity order)
    pub row_order: Option<Vec<usize>>,
    /// Sort state (column, is_ascending) - None = no sort
    pub sort: Option<(usize, bool)>,
}

/// Preserved focus state for restoring after preview
#[derive(Clone, Debug)]
pub struct PreviewFocus {
    pub sheet_index: usize,
    pub selected: (usize, usize),
    pub selection_end: Option<(usize, usize)>,
    pub scroll_row: usize,
    pub scroll_col: usize,
}

/// Maximum history actions to replay for preview (safety valve)
pub const MAX_PREVIEW_REPLAY: usize = 10_000;
/// Maximum time budget for building preview snapshot (ms)
pub const MAX_PREVIEW_BUILD_MS: u64 = 200;
/// Consistent message for all blocked commands during preview
pub const PREVIEW_BLOCK_MSG: &str = "Preview mode — release Space to edit";

/// Document metadata for title bar display.
#[derive(Clone, Debug)]
pub struct DocumentMeta {
    /// Display name - FULL filename with extension (e.g., "budget.xlsx", not "budget")
    /// For unsaved documents, this is "Book1", "Book2", etc. (no extension)
    pub display_name: String,
    /// Document has been saved at least once (to native format)
    pub is_saved: bool,
    /// Document is read-only
    pub is_read_only: bool,
    /// How the document was opened/created (only for non-native sources)
    pub source: Option<DocumentSource>,
    /// Full path if saved
    pub path: Option<PathBuf>,
}

impl Default for DocumentMeta {
    fn default() -> Self {
        Self {
            display_name: next_book_name(),
            is_saved: false,
            is_read_only: false,
            source: None,
            path: None,
        }
    }
}

impl DocumentMeta {
    /// Generate the window title string for macOS (includes provenance)
    pub fn title_string_full(&self, is_dirty: bool) -> String {
        let mut title = self.display_name.clone();

        // Dirty indicator
        if is_dirty {
            title.push_str(" \u{25CF}"); // ●
        }

        // Unsaved suffix (new document, never saved)
        if !self.is_saved && self.source.is_none() {
            title.push_str(" \u{2014} unsaved"); // —
        }

        // Provenance subtitle (only for imported/recovered)
        if let Some(source) = &self.source {
            match source {
                DocumentSource::Imported { filename } => {
                    title.push_str(&format!(" \u{2014} imported from {}", filename));
                }
                DocumentSource::Recovered => {
                    title.push_str(" \u{2014} recovered session");
                }
            }
        }

        // Read-only indicator
        if self.is_read_only {
            title.push_str(" \u{2014} read-only");
        }

        title
    }

    /// Generate the window title string for Windows/Linux (compact, no provenance)
    ///
    /// Provenance is omitted because:
    /// - Window titles get truncated aggressively on these platforms
    /// - Long titles pollute task switchers (Alt+Tab, taskbar)
    pub fn title_string_short(&self, is_dirty: bool) -> String {
        let mut title = self.display_name.clone();

        // Dirty indicator
        if is_dirty {
            title.push_str(" \u{25CF}"); // ●
        }

        // Unsaved suffix
        if !self.is_saved && self.source.is_none() {
            title.push_str(" \u{2014} unsaved");
        }

        // Read-only indicator (important enough to keep)
        if self.is_read_only {
            title.push_str(" \u{2014} read-only");
        }

        // App name suffix (Windows/Linux convention)
        title.push_str(" \u{2014} VisiGrid");

        title
    }

    /// Platform-appropriate title string
    pub fn title_string(&self, is_dirty: bool) -> String {
        #[cfg(target_os = "macos")]
        { self.title_string_full(is_dirty) }

        #[cfg(not(target_os = "macos"))]
        { self.title_string_short(is_dirty) }
    }

    /// Primary title part: filename + dirty indicator + unsaved/read-only
    /// Used for prominent display in custom titlebar
    pub fn title_primary(&self, is_dirty: bool) -> String {
        let mut title = self.display_name.clone();

        if is_dirty {
            title.push_str(" \u{25CF}"); // ●
        }

        if !self.is_saved && self.source.is_none() {
            title.push_str(" — unsaved");
        }

        if self.is_read_only {
            title.push_str(" — read-only");
        }

        title
    }

    /// Secondary title part: provenance/context info
    /// Returns None if no provenance, Some("imported from X") otherwise
    /// Used for quieter display in custom titlebar (no dash - hierarchy via size/color)
    pub fn title_secondary(&self) -> Option<String> {
        match &self.source {
            Some(DocumentSource::Imported { filename }) => {
                Some(format!("imported from {}", filename))
            }
            Some(DocumentSource::Recovered) => {
                Some("recovered session".to_string())
            }
            None => None,
        }
    }
}

/// Tri-state value for properties across multiple cells
#[derive(Debug, Clone, PartialEq)]
pub enum TriState<T> {
    /// All cells have the same value
    Uniform(T),
    /// Cells have different values
    Mixed,
    /// No cells in selection (shouldn't happen)
    Empty,
}

impl<T: PartialEq + Clone> TriState<T> {
    /// Combine with another value
    pub fn combine(&self, other: &T) -> Self {
        match self {
            TriState::Empty => TriState::Uniform(other.clone()),
            TriState::Uniform(v) if v == other => TriState::Uniform(v.clone()),
            TriState::Uniform(_) => TriState::Mixed,
            TriState::Mixed => TriState::Mixed,
        }
    }

    /// Get the uniform value if present
    pub fn uniform(&self) -> Option<&T> {
        match self {
            TriState::Uniform(v) => Some(v),
            _ => None,
        }
    }

    pub fn is_mixed(&self) -> bool {
        matches!(self, TriState::Mixed)
    }
}

/// Which field has focus in the Create Named Range dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CreateNameFocus {
    #[default]
    Name,        // Name input field
    Description, // Description input field
}

/// Which tab is active in the bottom panel (Lua console / Terminal).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BottomPanelTab {
    #[default]
    Lua,
    Terminal,
    Problems,
}

/// One row in the Problems panel: a cell whose formula evaluates to an error.
#[derive(Debug, Clone)]
pub struct Problem {
    pub sheet_idx: usize,
    pub sheet_name: String,
    pub row: usize,
    pub col: usize,
    /// Error code, e.g. "#DIV/0!", "#REF!", "#CYCLE!"
    pub error: String,
    /// The formula (raw text) that produced it.
    pub formula: String,
}

/// Which editor surface is active (for popup anchoring and input routing)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum EditorSurface {
    #[default]
    Cell,       // Editing in the cell itself
    FormulaBar, // Editing in the formula bar
}

/// Fill handle drag axis (locked after first significant movement)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillAxis {
    Row,  // Filling vertically (down or up)
    Col,  // Filling horizontally (right or left)
}

/// Fill handle drag state machine
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FillDrag {
    #[default]
    None,
    Dragging {
        /// The start of source range when drag started
        anchor: (usize, usize),
        /// The end of source range (same as anchor for single cell)
        source_end: (usize, usize),
        /// Current hover cell during drag
        current: (usize, usize),
        /// Axis lock (None until threshold crossed, then locked)
        axis: Option<FillAxis>,
    },
}

use visigrid_engine::cell::{Alignment, CellBorder, CellStyle, NegativeStyle, VerticalAlignment, TextOverflow, NumberFormat, max_border};

/// Which context menu variant to display on right-click.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContextMenuKind {
    Cell,
    RowHeader,
    ColHeader,
}

/// State for the right-click context menu on cells/headers.
#[derive(Debug, Clone, Copy)]
pub struct ContextMenuState {
    pub kind: ContextMenuKind,
    pub position: Point<Pixels>,
}

/// State of the "set as default app" prompt in the title bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DefaultAppPromptState {
    /// Not showing (hidden or not applicable)
    #[default]
    Hidden,
    /// Showing the prompt for a specific file type
    Showing,
    /// User clicked "Make default" - show success briefly
    Success,
    /// User clicked but needs to finish in System Settings
    NeedsSettings,
}

/// Format state for a selection of cells
#[derive(Debug, Clone)]
pub struct SelectionFormatState {
    pub cell_count: usize,
    // Value state
    pub raw_value: TriState<String>,      // Raw input value
    pub display_value: Option<String>,    // Formatted display (only if uniform)
    // Format properties
    pub bold: TriState<bool>,
    pub italic: TriState<bool>,
    pub underline: TriState<bool>,
    pub strikethrough: TriState<bool>,
    pub font_family: TriState<Option<String>>,
    pub alignment: TriState<Alignment>,
    pub vertical_alignment: TriState<VerticalAlignment>,
    pub text_overflow: TriState<TextOverflow>,
    pub number_format: TriState<NumberFormat>,
    pub background_color: TriState<Option<[u8; 4]>>,
    pub font_size: TriState<Option<f32>>,
    pub font_color: TriState<Option<[u8; 4]>>,
    pub cell_style: TriState<CellStyle>,
    /// Active cell numeric value for preview (None if non-numeric or multi-cell)
    pub preview_value: Option<f64>,
}

impl Default for SelectionFormatState {
    fn default() -> Self {
        Self {
            cell_count: 0,
            raw_value: TriState::Empty,
            display_value: None,
            bold: TriState::Empty,
            italic: TriState::Empty,
            underline: TriState::Empty,
            strikethrough: TriState::Empty,
            font_family: TriState::Empty,
            alignment: TriState::Empty,
            vertical_alignment: TriState::Empty,
            text_overflow: TriState::Empty,
            number_format: TriState::Empty,
            background_color: TriState::Empty,
            font_size: TriState::Empty,
            font_color: TriState::Empty,
            cell_style: TriState::Empty,
            preview_value: None,
        }
    }
}

// ============================================================================
// Rewind Confirmation Dialog State (Phase 8C)
// ============================================================================

/// State for the destructive rewind confirmation dialog.
/// Shows number of actions to be discarded and requires explicit confirmation.
#[derive(Clone, Debug, Default)]
pub struct RewindConfirmState {
    /// Whether the dialog is visible
    pub visible: bool,
    /// Number of actions that will be discarded
    pub discard_count: usize,
    /// Summary of the target action we're rewinding to (before this)
    pub target_summary: String,
    /// Sheet name where the target action occurred (if available)
    pub sheet_name: Option<String>,
    /// Cell range affected by the target action (if available)
    pub location: Option<String>,
    /// Entry ID of the target action
    pub target_entry_id: u64,
    /// Number of actions replayed for preview
    pub replay_count: usize,
    /// Time spent building preview (ms)
    pub build_ms: u64,
    /// History fingerprint at preview time
    pub fingerprint: HistoryFingerprint,
    /// Pre-built rewind plan (if available)
    pub plan: Option<RewindPlan>,
}

impl RewindConfirmState {
    /// Show the confirmation dialog with the given plan and context
    pub fn show(
        &mut self,
        discard_count: usize,
        target_summary: String,
        sheet_name: Option<String>,
        location: Option<String>,
        target_entry_id: u64,
        replay_count: usize,
        build_ms: u64,
        fingerprint: HistoryFingerprint,
        plan: RewindPlan,
    ) {
        self.visible = true;
        self.discard_count = discard_count;
        self.target_summary = target_summary;
        self.sheet_name = sheet_name;
        self.location = location;
        self.target_entry_id = target_entry_id;
        self.replay_count = replay_count;
        self.build_ms = build_ms;
        self.fingerprint = fingerprint;
        self.plan = Some(plan);
    }

    /// Hide the dialog and clear state
    pub fn hide(&mut self) {
        self.visible = false;
        self.plan = None;
    }
}

/// State for the merge cells confirmation dialog.
/// Shown when merging would discard non-empty cell values.
#[derive(Clone, Debug, Default)]
pub struct MergeConfirmState {
    /// Whether the dialog is visible
    pub visible: bool,
    /// Cell addresses whose values will be lost (display strings for the dialog)
    pub affected_cells: Vec<String>,
    /// The selection range to merge: ((min_row, min_col), (max_row, max_col))
    pub merge_range: Option<((usize, usize), (usize, usize))>,
}

/// Banner shown briefly after a successful rewind.
/// Displays count and provides "Copy audit" for audit trail.
#[derive(Clone, Debug, Default)]
pub struct RewindSuccessBanner {
    /// Whether the banner is visible
    pub visible: bool,
    /// Number of actions that were discarded
    pub discarded_count: usize,
    /// Summary of the target action
    pub target_summary: String,
    /// Full audit details for clipboard copy (single-line format)
    pub audit_details: String,
    /// When the banner was shown (for auto-dismiss)
    pub shown_at: Option<std::time::Instant>,
}

/// Audit data for rewind banner display and clipboard copy
pub struct RewindAuditData {
    pub target_entry_id: u64,
    pub target_summary: String,
    pub discarded_count: usize,
    pub replay_count: usize,
    pub build_ms: u64,
    pub fingerprint: HistoryFingerprint,
}

impl RewindSuccessBanner {
    /// Show the banner with full audit details.
    /// Formats a single-line audit record suitable for logs or clipboard.
    pub fn show(&mut self, audit: RewindAuditData) {
        self.visible = true;
        self.discarded_count = audit.discarded_count;
        self.target_summary = audit.target_summary.clone();

        // Format UTC timestamp (ISO 8601 compact)
        let utc_timestamp = chrono_lite_utc();

        // Format fingerprint as hex (first 16 chars for readability)
        let fp_short = format!("{:016x}", audit.fingerprint.hash_hi);

        // Single-line audit format for clipboard:
        // UTC | Rewind to #ID (Before "Summary") | Discarded N | Replay M actions | Xms | Fingerprint abc...
        self.audit_details = format!(
            "{} | Rewind to #{} (Before \"{}\") | Discarded {} | Replay {} actions | {}ms | Fingerprint {}",
            utc_timestamp,
            audit.target_entry_id,
            audit.target_summary,
            audit.discarded_count,
            audit.replay_count,
            audit.build_ms,
            fp_short
        );
        self.shown_at = Some(std::time::Instant::now());
    }

    /// Hide the banner
    pub fn hide(&mut self) {
        self.visible = false;
        self.shown_at = None;
    }

    /// Check if banner should auto-dismiss (after 5 seconds)
    pub fn should_dismiss(&self) -> bool {
        self.shown_at.map(|t| t.elapsed().as_secs() >= 5).unwrap_or(false)
    }
}

