use gpui::*;
use std::collections::HashMap;
use std::path::PathBuf;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::CellLookup;
use visigrid_engine::filter::{RowView, FilterState};
use visigrid_engine::sheet::SheetId;
use visigrid_engine::cell::{CellBorder, CellStyle, max_border, NumberFormat, NegativeStyle};

use crate::clipboard::InternalClipboard;
use crate::find_replace::MatchHit;
use crate::formatting::BorderApplyMode;
use crate::history::{History, HistoryFingerprint};
use crate::mode::Mode;
use crate::repeat::RepeatAction;
use crate::search::{SearchEngine, SearchAction, CommandId, CommandSearchProvider, GoToSearchProvider, SearchItem, MenuCategory};
use crate::settings::{
    user_settings_path, open_settings_file, user_settings, update_user_settings,
    observe_settings, TipId,
};
use crate::theme::{Theme, TokenKey, default_theme, get_theme, SYSTEM_THEME_ID, resolve_system_theme_id};
use crate::views;
use crate::workbook_view::WorkbookViewState;

// Re-export from autocomplete module for external access
pub use crate::autocomplete::{SignatureHelpInfo, FormulaErrorInfo};

// Re-export from formula_refs module
pub use crate::formula_refs::{RefKey, FormulaRef, REF_COLORS};

// ============================================================================
// Global Book Counter (for "Book1", "Book2", etc.)
// ============================================================================

use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::OnceLock;

pub(crate) use crate::ai_cli::*;
pub use crate::rewind_state::*;
pub use crate::validation_state::*;
pub use crate::ai_dialog_state::*;

/// Generate the next book name (e.g., "Book1", "Book2", ...)
/// Session-level counter for new workbook names.
/// Increments each time a new workbook is created: Book1, Book2, Book3...
static NEXT_BOOK_NUMBER: AtomicU32 = AtomicU32::new(1);

pub fn next_book_name() -> String {
    let n = NEXT_BOOK_NUMBER.fetch_add(1, Ordering::Relaxed);
    format!("Book{}", n)
}

// ============================================================================
// Smoke Mode Recalc (Phase 1.5 - headless dogfooding)
// ============================================================================

/// Check if smoke recalc is enabled via VISIGRID_RECALC=full env var.
static SMOKE_RECALC_ENABLED: OnceLock<bool> = OnceLock::new();

pub(crate) fn is_smoke_recalc_enabled() -> bool {
    *SMOKE_RECALC_ENABLED.get_or_init(|| {
        std::env::var("VISIGRID_RECALC").ok().as_deref() == Some("full")
    })
}

// ============================================================================
// Palette Scope (for Alt accelerator filtering)
// ============================================================================

/// Palette scope for filtering Command Palette results.
///
/// This abstraction supports menu scoping now and can be extended
/// for selection-scoped commands, contextual palettes, etc.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PaletteScope {
    /// Filter to commands in a specific menu category (Alt accelerators)
    Menu(MenuCategory),
    /// Ctrl+K / Cmd+K: default provider = recent files
    QuickOpen,
}

// ============================================================================
// Document Identity (for title bar display)
// ============================================================================

/// Sentinel value for unassigned session window IDs.
/// Any Spreadsheet with this value has not been registered with SessionManager.
pub const WINDOW_ID_UNSET: u64 = u64::MAX;

/// Native file extension for VisiGrid documents
#[allow(dead_code)]
pub const NATIVE_EXT: &str = "vgrid";

/// Returns true if the extension is considered "native" (no provenance needed).
/// Native formats: vgrid (our format), xlsx/xls (Excel, first-class support)
pub fn is_native_ext(ext: &str) -> bool {
    matches!(ext.to_lowercase().as_str(), "vgrid" | "xlsx" | "xls" | "xlsb" | "xlsm" | "sheet")
}

/// Extract display filename from path (full name with extension)
pub(crate) fn display_filename(path: &std::path::Path) -> String {
    path.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("Untitled")
        .to_string()
}

/// Extract lowercase extension from path
pub(crate) fn ext_lower(path: &std::path::Path) -> Option<String> {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_lowercase())
}

/// Source of the document (for provenance display).
///
/// Only used for non-native formats that were imported/converted.
/// Native formats (vgrid, xlsx) have no provenance - they're first-class.
#[derive(Clone, Debug, PartialEq)]
pub enum DocumentSource {
    /// Imported from a non-native format (CSV, TSV, JSON)
    /// These are converted on load and need "Save As" to persist as native.
    Imported { filename: String },
    /// Recovered from session restore (unsaved work from crash/quit)
    Recovered,
}

/// History panel filter mode
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum HistoryFilterMode {
    #[default]
    All,
    CurrentSheet,
    ValidationOnly,
    DataEditsOnly,
    TransformsOnly,
}

/// Semantic verification status based on expected fingerprint.
///
/// When a file has an expected semantic fingerprint (from CLI --stamp or GUI Approve),
/// the GUI compares it against the current computed fingerprint.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerificationStatus {
    /// No expected fingerprint - file hasn't been stamped/approved
    Unverified,
    /// Current fingerprint matches expected - file unchanged since stamp/approval
    Verified,
    /// Current fingerprint doesn't match expected - file has been modified
    Drifted,
}

// Legacy alias for migration - TODO: remove after updating all usages
pub type ApprovalStatus = VerificationStatus;


// Grid configuration
pub use visigrid_session_host::{NUM_ROWS, NUM_COLS, MAX_SESSION_FORMAT_CELLS, MAX_SESSION_INSPECT_CELLS};

/// Auto-fit bounds. The maximum keeps one pathological cell from pushing a
/// column off-screen (Excel caps column width similarly).
pub const MIN_AUTOFIT_WIDTH: f32 = 40.0;
pub const MAX_AUTOFIT_WIDTH: f32 = 600.0;
/// Cell padding (px_1 each side) plus a little slack so text is not flush.
const AUTOFIT_PADDING: f32 = 12.0;

/// Width estimate for paths with no window to shape text with.
/// Counts CHARACTERS, and counts East-Asian wide characters double — the
/// previous estimator multiplied byte length, which over-measured every
/// non-ASCII string.
fn estimate_text_width(text: &str, font_size: f32, bold: bool) -> f32 {
    let units: f32 = text
        .chars()
        .map(|c| {
            let wide = matches!(c as u32,
                0x1100..=0x115F | 0x2E80..=0xA4CF | 0xAC00..=0xD7A3 |
                0xF900..=0xFAFF | 0xFE30..=0xFE6F | 0xFF00..=0xFF60 | 0xFFE0..=0xFFE6);
            if wide { 2.0 } else { 1.0 }
        })
        .sum();
    let per_unit = font_size * 0.55 * if bold { 1.06 } else { 1.0 };
    units * per_unit
}

/// A pending pairing approval dialog: a client asked to control this
/// workbook and the user hasn't answered yet.
pub struct PairingPrompt {
    /// Sanitized client name to display.
    pub client_name: String,
    /// Reply channel back to the TCP thread (true = approve).
    pub reply: Option<crate::session_server::bridge::oneshot::Sender<bool>>,
}

pub const CELL_WIDTH: f32 = 80.0;
pub const CELL_HEIGHT: f32 = 24.0;
pub const HEADER_WIDTH: f32 = 50.0;
pub const MENU_BAR_HEIGHT: f32 = 28.0;
pub const FORMULA_BAR_HEIGHT: f32 = 28.0;
pub const COLUMN_HEADER_HEIGHT: f32 = 24.0;
pub const STATUS_BAR_HEIGHT: f32 = 24.0;
pub const MACOS_TITLEBAR_HEIGHT: f32 = 34.0;

// Resize grab zones — the clickable area at header edges for resizing.
// Must be less than half the minimum row/col dimension to avoid swallowing selection clicks.
pub const ROW_RESIZE_GRAB_PX: f32 = 4.0;
pub const COL_RESIZE_GRAB_PX: f32 = 6.0;

// Formula bar layout (single source of truth for hit-testing + rendering)
pub const FORMULA_BAR_CELL_REF_WIDTH: f32 = 60.0;
pub const FORMULA_BAR_FX_WIDTH: f32 = 30.0;
pub const FORMULA_BAR_PADDING: f32 = 8.0;  // px_2
/// X offset where text content starts (cell ref + fx button + padding)
pub const FORMULA_BAR_TEXT_LEFT: f32 = FORMULA_BAR_CELL_REF_WIDTH + FORMULA_BAR_FX_WIDTH + FORMULA_BAR_PADDING;

// Zoom configuration
pub const ZOOM_STEPS: &[f32] = &[0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0];
pub const DEFAULT_ZOOM: f32 = 1.0;

/// Cached grid metrics scaled by zoom level.
/// Single source of truth for all scaled geometry.
/// Dimensions are pixel-snapped to the device scale factor to eliminate
/// sub-pixel shimmer at fractional zoom levels.
#[derive(Clone, Copy)]
pub struct GridMetrics {
    pub zoom: f32,
    /// Device scale factor (e.g. 2.0 on Retina). Used for pixel snapping.
    pub scale: f32,
    pub cell_w: f32,
    pub cell_h: f32,
    pub header_w: f32,
    pub header_h: f32,
    pub font_size: f32,
}

impl GridMetrics {
    pub fn new(zoom: f32) -> Self {
        Self::with_scale(zoom, 1.0)
    }

    pub fn with_scale(zoom: f32, scale: f32) -> Self {
        Self {
            zoom,
            scale,
            cell_w: Self::snap(CELL_WIDTH * zoom, scale),
            cell_h: Self::snap(CELL_HEIGHT * zoom, scale),
            header_w: Self::snap(HEADER_WIDTH * zoom, scale),
            header_h: Self::snap(COLUMN_HEADER_HEIGHT * zoom, scale),
            font_size: 13.0 * zoom, // font size doesn't snap
        }
    }

    /// Snap a logical dimension to the nearest device pixel boundary (round).
    /// Use for widths/heights so cells have consistent integer-pixel sizes.
    pub fn snap(logical: f32, scale: f32) -> f32 {
        if scale <= 0.0 { return logical; }
        (logical * scale).round() / scale
    }

    /// Snap a logical position to a device pixel boundary (floor).
    /// Use for accumulated offsets so positions are stable during scroll.
    pub fn snap_floor(logical: f32, scale: f32) -> f32 {
        if scale <= 0.0 { return logical; }
        (logical * scale).floor() / scale
    }

    /// Get scaled width for a column (model width * zoom), pixel-snapped.
    pub fn col_width(&self, model_width: f32) -> f32 {
        Self::snap(model_width * self.zoom, self.scale)
    }

    /// Get scaled height for a row (model height * zoom), pixel-snapped.
    pub fn row_height(&self, model_height: f32) -> f32 {
        Self::snap(model_height * self.zoom, self.scale)
    }
}

impl Default for GridMetrics {
    fn default() -> Self {
        Self::new(DEFAULT_ZOOM)
    }
}

/// Cached layout measurements for hit-testing (updated each render)
#[derive(Clone, Copy, Default)]
pub struct GridLayout {
    /// Grid body origin in window coordinates (top-left of first cell)
    pub grid_body_origin: (f32, f32),
    /// Viewport size for the grid body (for limiting iteration)
    pub viewport_size: (f32, f32),
}

/// A cell's bounding rectangle in grid-relative coordinates.
/// Used for positioning popups and overlays relative to cells.
#[derive(Clone, Copy, Debug, Default)]
pub struct CellRect {
    /// Left edge X position (relative to grid origin)
    pub x: f32,
    /// Top edge Y position (relative to grid origin)
    pub y: f32,
    /// Cell width
    pub width: f32,
    /// Cell height
    pub height: f32,
}

impl CellRect {
    /// Bottom edge Y position
    pub fn bottom(&self) -> f32 {
        self.y + self.height
    }

    /// Right edge X position
    pub fn right(&self) -> f32 {
        self.x + self.width
    }
}

/// Transient UI state that is never serialized.
///
/// Non-persisted, ephemeral view state for pickers and dialogs.
///
/// **Rule:** Focus handles, query strings, cursor positions, selection
/// indices, and recent-item lists belong here — NOT on `Spreadsheet`.
/// `Spreadsheet` owns the document view-model (workbook, history,
/// selection, scroll, mode). `UiState` owns transient dialog chrome
/// that is never serialized and has no undo semantics.
///
/// Color picker is the first occupant. Font picker, theme picker,
/// goto/find dialogs, command palette, etc. should migrate here
/// incrementally (opportunistic, not a scheduled refactor).
pub struct UiState {
    pub color_picker: crate::color_palette::ColorPickerState,
    pub format_bar: FormatBarState,
    /// Format dropdown menu in header bar (Bold/Italic/Underline/Alignment)
    pub format_menu_open: bool,
}

/// Transient UI state for the format bar (font size input, dropdown).
/// Never serialized, no undo semantics.
pub struct FormatBarState {
    pub size_input: String,
    pub size_editing: bool,
    pub size_dropdown: bool,
    pub size_focus: FocusHandle,
    /// True on the first keypress after entering edit mode — clears the buffer
    /// so the user can type a replacement value without manually selecting all.
    pub size_replace_next: bool,
    /// Number format quick-menu dropdown (123 ▾ button).
    pub number_format_menu_open: bool,
    /// Cell style quick-menu dropdown (Styles ▾ button).
    pub cell_style_menu_open: bool,
}

impl FormatBarState {
    /// Returns true when the format bar owns focus (editing or dropdown open).
    /// Used to gate grid keyboard and mouse handling.
    pub fn is_active(&self, window: &Window) -> bool {
        self.size_editing || self.size_dropdown || self.number_format_menu_open || self.cell_style_menu_open || self.size_focus.is_focused(window)
    }
}

pub struct Spreadsheet {
    // Core data
    /// The shared workbook entity. All mutations must go through update(cx, ...).
    /// This enables future multi-view support where multiple views share the same workbook.
    pub workbook: Entity<Workbook>,
    pub history: History,
    /// Base workbook state for replay (captured on load/new, never mutated)
    pub base_workbook: Workbook,
    /// Soft-rewind preview state (Phase 8A)
    pub rewind_preview: RewindPreviewState,

    // Role-based auto-styling (agent metadata)
    /// Cell metadata loaded from .sheet file (target -> {key: value})
    pub cell_metadata: crate::role_styles::CellMetadataMap,
    /// Role -> style mapping (singleton, could become per-doc)
    pub role_style_map: crate::role_styles::RoleStyleMap,

    // Lua script invariants
    /// Scripts attached to this workbook (persisted in .sheet file).
    pub attached_scripts: Vec<visigrid_io::scripting::ScriptMeta>,
    /// Run records created since last save (not yet persisted).
    pub pending_run_records: Vec<visigrid_io::scripting::RunRecord>,
    /// Run records loaded from .sheet file (already persisted).
    pub loaded_run_records: Vec<visigrid_io::scripting::RunRecord>,

    // Row view layer (for sort/filter)
    // Maps view rows to data rows, handles visibility
    pub row_view: RowView,
    pub filter_state: FilterState,
    /// Which column's filter dropdown is currently open (None = closed)
    pub filter_dropdown_col: Option<usize>,
    /// Search text in the filter dropdown
    pub filter_search_text: String,
    /// Currently checked items in the filter dropdown (indexes into unique values)
    pub filter_checked_items: std::collections::HashSet<usize>,

    // View state (selection, scroll, zoom, freeze panes)
    // This will become Entity<WorkbookView> in future phases for multi-tab support
    pub view_state: WorkbookViewState,

    /// Column the cursor was in before Shift+Space (select row).
    /// Used to restore the column when the selection collapses (e.g. arrow key).
    /// None when no row selection is active.
    pub pre_row_select_col: Option<usize>,

    // Split view state (Ctrl+\ to split right)
    pub split_pane: Option<crate::split_view::SplitPane>,
    pub split_active_side: crate::split_view::SplitSide,

    // Dependency tracing (Alt+T to toggle)
    pub trace_enabled: bool,
    pub trace_cache: Option<crate::trace::TraceCache>,

    // Mode & editing
    pub mode: Mode,
    pub edit_value: String,
    pub edit_cursor: usize,  // Cursor position within edit_value (byte offset, 0..=len)
    pub edit_selection_anchor: Option<usize>,  // Selection start (None = no selection)
    pub edit_original: String,
    pub edit_scroll_x: f32,  // Horizontal scroll offset for in-cell editor (<=0, updated by ensure_caret_visible)
    pub(crate) edit_scroll_dirty: bool, // True when caret/text changed; triggers ensure_caret_visible once

    // Caret blink state
    pub caret_visible: bool,
    pub caret_last_activity: std::time::Instant,
    pub(crate) caret_blink_task: Option<gpui::Task<()>>,

    // KeyTips state (macOS Option+Space accelerator hints)
    /// True when KeyTips overlay is visible
    pub keytips_active: bool,
    /// Auto-dismiss deadline (3 seconds after activation)
    pub keytips_deadline_at: Option<std::time::Instant>,
    /// Last scope opened via KeyTips (for Enter/Space repeat)
    pub last_keytips_scope: Option<crate::search::MenuCategory>,
    /// True after KeyTips discovery hint has been shown (once per session)
    pub keytips_hint_shown: bool,

    pub goto_input: String,
    pub find_input: String,
    pub find_results: Vec<MatchHit>,
    pub find_index: usize,
    pub replace_input: String,
    pub find_replace_mode: bool,      // true = Find & Replace (Ctrl+H), false = Find only (Ctrl+F)
    pub find_focus_replace: bool,     // true = replace input has focus, false = find input

    // Command palette
    pub palette_query: String,
    pub palette_selected: usize,
    /// First visible row of the palette results (windowed list scroll).
    pub palette_scroll_offset: usize,
    pub palette_scope: Option<PaletteScope>,  // Menu scope for Alt accelerators
    pub(crate) search_engine: SearchEngine,
    pub(crate) palette_results: Vec<SearchItem>,
    pub palette_total_results: usize,  // Total matches before truncation
    // Pre-palette state for preview/restore
    pub(crate) palette_pre_selection: (usize, usize),
    pub(crate) palette_pre_selection_end: Option<(usize, usize)>,
    pub(crate) palette_pre_scroll: (usize, usize),
    pub palette_previewing: bool,  // True if user has previewed (Shift+Enter)

    // Clipboard
    pub internal_clipboard: Option<InternalClipboard>,
    /// Visual range for copy/cut dashed border overlay (r1, c1, r2, c2).
    /// Set on Copy/Cut, cleared on Paste/Escape/edit start/confirm/delete.
    pub clipboard_visual_range: Option<(usize, usize, usize, usize)>,

    // File state
    /// Unique ID for session matching (assigned at startup).
    /// Initialized to WINDOW_ID_UNSET — must be assigned via SessionManager::next_window_id()
    /// before the first snapshot/save.
    pub session_window_id: u64,
    pub current_file: Option<PathBuf>,
    pub is_modified: bool,  // Legacy - use is_dirty() for title bar
    pub close_after_save: bool,  // Set by save_and_close() to close window after Save As completes
    pub window_handle: gpui::AnyWindowHandle,  // Handle for closing window from async contexts
    pub recent_files: Vec<PathBuf>,  // Recently opened files (most recent first)
    pub recent_commands: Vec<CommandId>,  // Recently executed commands (most recent first)

    // Document identity (for title bar)
    pub document_meta: DocumentMeta,
    pub(crate) cached_title: Option<String>,  // For debouncing title updates
    pub(crate) pending_title_refresh: bool,   // Set true + notify() when title may have changed without window access

    // UI state
    pub focus_handle: FocusHandle,
    pub console_focus_handle: FocusHandle,
    pub script_view_focus_handle: FocusHandle,
    pub status_message: Option<String>,
    pub window_size: Size<Pixels>,
    pub cached_window_bounds: Option<WindowBounds>,  // Cached for session snapshot

    // Column/row sizing (per-sheet)
    // Each sheet has independent column widths and row heights.
    // New sheets start with defaults (Excel behavior), not inherited from current sheet.
    pub col_widths: HashMap<SheetId, HashMap<usize, f32>>,   // SheetId -> col -> width
    pub row_heights: HashMap<SheetId, HashMap<usize, f32>>,  // SheetId -> row -> height

    // Hidden rows/columns (per-sheet, user-controlled, separate from AutoFilter)
    pub hidden_rows: HashMap<SheetId, std::collections::BTreeSet<usize>>,
    pub hidden_cols: HashMap<SheetId, std::collections::BTreeSet<usize>>,

    /// Cached active sheet ID for fast lookups without context.
    /// Updated whenever the active sheet changes.
    cached_sheet_id: SheetId,

    // Resize drag state
    pub resizing_col: Option<usize>,       // Column being resized (by right edge)
    pub resizing_row: Option<usize>,       // Row being resized (by bottom edge)
    pub resize_start_pos: f32,             // Mouse position at drag start
    pub resize_start_size: f32,            // Original size at drag start
    pub resize_start_original: Option<f32>, // Original map value (None = was default)

    // Menu bar state (Excel 2003 style dropdown menus)
    pub open_menu: Option<crate::mode::Menu>,
    pub menu_highlight: Option<usize>,

    // Sheet tab state
    pub renaming_sheet: Option<usize>,     // Index of sheet being renamed
    pub sheet_rename_input: String,        // Current rename input value
    pub sheet_rename_cursor: usize,        // Cursor position (byte index)
    pub sheet_rename_select_all: bool,     // Text is fully selected (typing replaces all)
    pub sheet_context_menu: Option<usize>, // Index of sheet with open context menu
    pub context_menu: Option<ContextMenuState>, // Right-click context menu on cells/headers

    // Font picker state
    pub available_fonts: Vec<String>,      // System fonts
    pub font_picker_query: String,         // Filter query
    pub font_picker_selected: usize,       // Selected item index
    pub font_picker_scroll_offset: usize,  // First visible item in list
    pub font_picker_focus: FocusHandle,    // Focus handle for the picker dialog

    // Transient UI state (not serialized — see UiState doc)
    pub ui: UiState,

    // Theme picker state
    pub theme_picker_query: String,        // Filter query
    pub theme_picker_selected: usize,      // Selected item index

    // Drag selection state
    pub dragging_selection: bool,          // Currently dragging to select cells

    // Fill handle drag state
    pub fill_drag: FillDrag,

    // Row/column header drag selection state
    pub dragging_row_header: bool,         // Currently dragging row headers
    pub dragging_col_header: bool,         // Currently dragging column headers
    pub row_header_anchor: Option<usize>,  // Anchor row for drag (stable during drag)
    pub col_header_anchor: Option<usize>,  // Anchor col for drag (stable during drag)

    // Layout cache for hit-testing
    pub grid_layout: GridLayout,

    // Formula reference selection state (for pointing mode)
    pub formula_ref_cell: Option<(usize, usize)>,      // Current reference cell (or range start)
    pub formula_ref_end: Option<(usize, usize)>,       // Range end (None = single cell)
    pub formula_ref_start_cursor: usize,               // Cursor position where reference started
    pub formula_nav_mode: crate::mode::FormulaNavMode, // Caret vs Point submode in Formula mode
    pub formula_nav_manual_override: Option<crate::mode::FormulaNavMode>, // F2 toggle latch - wins over auto-switch
    pub formula_home_sheet: Option<usize>,              // Sheet where formula is being entered (for cross-sheet refs)
    pub formula_edit_cell: Option<(usize, usize)>,     // Cell being edited (preserved across sheet switches in formula mode)
    pub formula_ref_sheet: Option<usize>,               // Sheet where current ref target lives (None = home sheet)
    pub formula_cross_sheet_name: Option<String>,       // Target sheet name when picking cross-sheet refs (None = same sheet)

    // Highlighted formula references (for existing formulas when editing)
    // Each entry has color index, cell bounds, and text position for formula bar coloring
    pub formula_highlighted_refs: Vec<FormulaRef>,

    // Persistent color assignment for formula references during editing
    // Ensures colors don't "jump" as user types - same RefKey keeps same color
    pub formula_ref_color_map: std::collections::HashMap<RefKey, usize>,
    pub formula_ref_next_color: usize,

    // Formula bar display cache (avoids re-parsing on every render)
    // Only used when NOT editing - caches parsed refs for the currently selected cell
    pub formula_bar_cache_cell: Option<(usize, usize)>,
    pub formula_bar_cache_formula: String,
    pub formula_bar_cache_refs: Vec<FormulaRef>,

    // Formula bar editing state (click-to-place caret, drag-to-select)
    pub active_editor: EditorSurface,
    pub formula_bar_scroll_x: f32,
    pub formula_bar_text_rect: gpui::Bounds<gpui::Pixels>,  // Text area rect in window coords (for hit-testing)
    pub(crate) formula_bar_cache_dirty: bool,
    pub(crate) formula_bar_char_boundaries: Vec<usize>,  // Byte offsets: [0, 1, 2, ..., len]
    pub(crate) formula_bar_boundary_xs: Vec<f32>,        // X positions aligned to boundaries
    pub formula_bar_text_width: f32,
    pub formula_bar_drag_anchor: Option<usize>,  // None = not dragging, Some(byte) = drag start anchor
    /// Formula bar expanded mode (shows 2-3 lines for long formulas)
    pub formula_bar_expanded: bool,

    // Name box (cell selector) editing state
    /// Whether the name box is being edited
    pub name_box_editing: bool,
    /// Current input value in name box
    pub name_box_input: String,
    /// Focus handle for name box keyboard events
    pub name_box_focus: FocusHandle,
    /// Replace on next keypress (select-all mode)
    pub name_box_replace_next: bool,

    // Formula autocomplete state
    pub autocomplete_visible: bool,
    pub autocomplete_suppressed: bool,  // Prevents autocomplete from reopening until text edit
    pub autocomplete_selected: usize,
    pub autocomplete_replace_range: std::ops::Range<usize>,

    // Formula hover documentation state
    pub hover_function: Option<&'static crate::formula_context::FunctionInfo>,

    // Document-level settings (persisted in sidecar file)
    pub doc_settings: crate::settings::DocumentSettings,

    // Minimap state (row-density navigator)
    pub minimap_visible: bool,
    pub minimap_cache: crate::minimap::MinimapCache,
    pub minimap_dragging: bool,
    pub minimap_drag_offset_y: f32,

    // Profiler panel state
    pub profiler_visible: bool,
    pub profiler_report: Option<visigrid_engine::recalc::RecalcReport>,
    pub profiler_hotspots: Vec<visigrid_engine::recalc::HotspotEntry>,
    pub profiler_capture_next: bool,

    // Locked feature panel dismiss (session-only)
    pub locked_panels_dismissed: bool,

    // Inspector panel state
    pub inspector_visible: bool,
    pub inspector_tab: crate::mode::InspectorTab,
    pub inspector_pinned: Option<(usize, usize)>,  // Pinned cell (None = follows selection)
    pub format_painter: Option<crate::formatting::FormatPaintState>,  // Format Painter state (snapshot + locked)
    /// Current border color for new borders. None = "Automatic" (theme default).
    pub current_border_color: Option<[u8; 4]>,
    pub tab_chain_origin_col: Option<usize>,  // Tab-chain return: origin column for Enter key
    pub inspector_hover_cell: Option<(usize, usize)>,  // Cell being hovered in inspector (for grid highlight)
    pub inspector_trace_path: Option<Vec<visigrid_engine::cell_id::CellId>>,  // Path trace highlight (Phase 3.5b)
    pub inspector_trace_incomplete: bool,  // True if trace has dynamic refs or was truncated
    pub names_filter_query: String,  // Filter query for Names tab
    pub selected_named_range: Option<String>,  // Selected named range in Names tab (Phase 5)
    pub selected_history_id: Option<u64>,  // Selected entry in History tab (Phase 4.3)
    pub history_filter_query: String,  // Filter query for History tab (Phase 4.3)
    pub history_filter_mode: HistoryFilterMode,  // Filter mode (Phase 7B)
    pub history_view_start: usize,  // Virtual scroll start index (Phase 7C)
    /// Highlighted range for history entry preview (sheet_index, start_row, start_col, end_row, end_col)
    pub history_highlight_range: Option<(usize, usize, usize, usize, usize)>,
    /// Current diff report (Explain Differences feature)
    pub diff_report: Option<crate::diff::DiffReport>,
    /// Filter diff report to show AI-touched changes only
    pub diff_ai_only_filter: bool,
    /// Selected entry in diff report (for highlighting, sheet_index, row, col)
    pub diff_selected_entry: Option<(usize, usize, usize)>,
    /// AI-generated summary of the diff (Phase 3)
    pub diff_ai_summary: Option<String>,
    /// Whether AI summary is currently being generated
    pub diff_ai_summary_loading: bool,
    /// Error from AI summary generation
    pub diff_ai_summary_error: Option<String>,
    /// Per-entry AI explanations cache: (sheet_index, row, col) → explanation
    pub diff_entry_explanations: std::collections::HashMap<(usize, usize, usize), String>,
    /// Entry currently being explained (sheet_index, row, col)
    pub diff_explaining_entry: Option<(usize, usize, usize)>,
    /// Entry ID for history context menu (right-click)
    pub history_context_menu_entry_id: Option<u64>,

    // Transform diff preview (Pro)
    pub transform_preview: Option<crate::transforms::TransformPreview>,

    // Zen mode (distraction-free editing)
    pub zen_mode: bool,

    // F1 context help (hold-to-peek)
    pub f1_help_visible: bool,

    // Zoom (zoom_level is in view_state, metrics is derived)
    pub metrics: GridMetrics,
    /// Debug overlay: draws pixel-alignment reference lines on the grid.
    /// Toggle via Cmd+Alt+Shift+G (dev use only — verifies cell boundary snapping).
    pub debug_grid_alignment: bool,
    /// Debug border instrumentation (only in debug builds).
    /// Uses Cell for interior mutability since render_cell takes &Spreadsheet.
    /// Toggle Cmd+Alt+Shift+G to print once/sec:
    ///   borders_calls=… gridline_cells=… userborder_cells=… frames=…
    #[cfg(debug_assertions)]
    pub debug_border_call_count: std::cell::Cell<u32>,
    #[cfg(debug_assertions)]
    pub debug_gridline_cells: std::cell::Cell<u32>,
    #[cfg(debug_assertions)]
    pub debug_userborder_cells: std::cell::Cell<u32>,
    #[cfg(debug_assertions)]
    debug_border_frames: std::cell::Cell<u32>,
    #[cfg(debug_assertions)]
    debug_border_last_report: std::cell::Cell<std::time::Instant>,
    /// Consecutive 1-second windows where has_any_borders=true but userborder_cells=0.
    /// Triggers a loud warning at 3 consecutive hits (likely stale flag).
    #[cfg(debug_assertions)]
    debug_border_stale_streak: u32,
    zoom_wheel_accumulator: f32,  // For smooth wheel zoom debounce

    // Navigation batching: accumulate repeat arrow events, flush at render start
    pub(crate) pending_nav_dx: i32,
    pub(crate) pending_nav_dy: i32,
    // Navigation coalescing: scroll adjustment deferred to render start
    pub(crate) nav_scroll_dirty: bool,
    // Navigation latency instrumentation (env VISIGRID_PERF=nav)
    pub(crate) nav_perf: crate::perf::NavLatencyTracker,

    // Link opening state (debounce rapid Ctrl+Enter)
    pub link_open_in_flight: bool,

    // Theme
    pub theme: Theme,
    pub theme_preview: Option<Theme>,  // For live preview in picker

    // Cell search cache (generation-based freshness)
    pub(crate) cells_rev: u64,  // Monotonically increasing; bumped on any cell value change
    pub(crate) cell_search_cache: CellSearchCache,
    pub(crate) named_range_usage_cache: NamedRangeUsageCache,

    // Rename symbol state (Ctrl+Shift+R)
    pub rename_original_name: String,      // The named range being renamed
    pub rename_new_name: String,           // User's typed new name
    pub rename_select_all: bool,           // True = typing replaces entire name
    pub rename_affected_cells: Vec<(usize, usize)>,  // Cells with formulas referencing this name
    pub rename_validation_error: Option<String>,     // Current validation error (if any)

    // Add conditional format state
    pub cf_input: String,                          // Typed rule: "=PRED -> STYLE"
    pub cf_input_error: Option<String>,            // Parse error shown in dialog
    pub cf_target: Vec<visigrid_engine::validation::CellRange>,  // Selection when opened
    pub cf_preview_id: Option<u64>,                // Live-preview rule currently in the store
    pub cf_preview_matches: Option<(usize, usize)>, // (matching, scanned) for the preview
    pub cf_panel_visible: bool,                    // Rules management drawer
    pub(crate) cf_rules_rev: u64,                  // Bumped on any CF rule mutation (cache key)
    /// Per-cell conditional format override cache, keyed by (cells_rev, cf_rules_rev).
    /// Heavy predicates (COUNTIF over large ranges) are evaluated once per
    /// edit/rule-change instead of once per frame per cell.
    pub(crate) cf_cache: std::cell::RefCell<std::collections::HashMap<(usize, usize), Option<visigrid_engine::cell::CellFormatOverride>>>,
    pub(crate) cf_cache_key: std::cell::Cell<(u64, u64)>,
    pub cf_edit_backup: Option<(usize, visigrid_engine::cond_format::CondFormatRule)>, // Rule pulled for editing (index, rule) — restored on cancel

    // Create named range state (Ctrl+Shift+N)
    pub create_name_name: String,           // User-typed name
    pub create_name_description: String,    // Optional description
    pub create_name_target: String,         // Auto-filled from selection (e.g., "A1:B10")
    pub create_name_validation_error: Option<String>,
    pub create_name_focus: CreateNameFocus, // Which field has focus

    // Edit description state
    pub edit_description_name: String,           // Name of the named range being edited
    pub edit_description_value: String,          // Current description input
    pub edit_description_original: Option<String>, // Original description (for undo)

    // Tour state
    pub tour_step: usize,                        // Current step (0-3)
    pub tour_completed: bool,                    // Has the tour been completed this session?
    pub show_f2_tip: bool,                       // Should we show the F2 tip this frame?

    // Settings subscription (for observing global settings changes)
    #[allow(dead_code)]
    settings_subscription: gpui::Subscription,

    // OS appearance observer — kept alive so System theme tracks OS dark/light
    #[allow(dead_code)]
    appearance_subscription: Option<gpui::Subscription>,

    // Impact preview state
    pub impact_preview_action: Option<crate::views::impact_preview::ImpactAction>,
    pub impact_preview_usages: Vec<crate::views::impact_preview::ImpactedFormula>,

    // Refactor log
    pub refactor_log: Vec<crate::views::refactor_log::RefactorLogEntry>,

    // Extract Named Range state
    pub extract_range_literal: String,           // The detected range literal (e.g., "A1:A100")
    pub extract_name: String,                    // User-entered name
    pub extract_description: String,             // User-entered description (optional)
    pub extract_affected_cells: Vec<(usize, usize)>,  // Cells containing this range
    pub extract_occurrence_count: usize,         // Total occurrences across all cells
    pub extract_validation_error: Option<String>,
    pub extract_select_all: bool,                // Type-to-replace for name field
    pub extract_focus: CreateNameFocus,          // Which field has focus (reusing enum)

    // Import report state (for Excel imports)
    pub import_result: Option<visigrid_io::xlsx::ImportResult>,
    pub import_filename: Option<String>,         // Original filename for display
    pub import_source_dir: Option<PathBuf>,      // Original directory for Save As default

    // Background import state
    pub import_in_progress: bool,
    pub import_overlay_visible: bool,
    pub import_started_at: Option<std::time::Instant>,

    // Startup timing (cold start measurement)
    pub startup_instant: Option<std::time::Instant>,
    pub cold_start_ms: Option<u128>,

    // Export report state (for Excel exports with warnings)
    pub export_result: Option<visigrid_io::xlsx::ExportResult>,
    pub export_filename: Option<String>,  // Exported filename for display

    // Keyboard hints state (Vimium-style jump)
    pub hint_state: crate::hints::HintState,

    // Bottom panel (shared area for Lua console + Terminal)
    pub bottom_panel_visible: bool,
    pub bottom_panel_tab: BottomPanelTab,

    // Terminal panel state
    pub terminal: crate::terminal::TerminalState,
    pub terminal_focus_handle: FocusHandle,
    /// Explicit boolean tracking terminal focus — secondary check for platforms
    /// where `FocusHandle::is_focused()` may not reflect focus correctly (macOS).
    pub terminal_focused: bool,

    // Lua scripting state
    pub lua_runtime: crate::scripting::LuaRuntime,
    pub lua_console: crate::scripting::ConsoleState,
    pub script: crate::scripting::ScriptState,
    pub custom_fn_registry: crate::scripting::CustomFunctionRegistry,

    // License dialog state
    pub license_input: String,
    pub license_error: Option<String>,

    // Trial CTA state (inline confirm in locked feature panels)
    pub trial_confirm_visible: bool,

    // Default app prompt state (macOS title bar chip)
    pub default_app_prompt_state: DefaultAppPromptState,
    pub default_app_prompt_file_type: Option<crate::default_app::SpreadsheetFileType>,
    pub(crate) default_app_prompt_success_timer: Option<std::time::Instant>,
    /// Timestamp when we entered NeedsSettings state (for backoff cutoff)
    pub(crate) needs_settings_entered_at: Option<std::time::Instant>,
    /// How many checks we've done in NeedsSettings (for exponential backoff)
    pub(crate) needs_settings_check_count: u8,

    // Smoke mode recalc guard (prevents reentrant recalc)
    pub(crate) in_smoke_recalc: bool,

    // Phase 2: Verified Mode - deterministic ordered recalc with visible status
    pub verified_mode: bool,
    pub last_recalc_report: Option<visigrid_engine::recalc::RecalcReport>,

    // Semantic verification state (persisted expected fingerprint)
    // Loaded from .sheet file on open, saved when approving/stamping.
    // Contains the expected semantic fingerprint that the current state is compared against.
    pub semantic_verification: visigrid_io::native::SemanticVerification,
    // UI state for approval dialogs
    pub approval_confirm_visible: bool,  // Confirmation dialog when re-approving after drift
    pub approval_drift_visible: bool,    // "Why drifted?" panel showing changes since approval
    pub approval_label_input: String,    // Label input for approval dialog
    // Legacy fields kept for history diff (shows what changed since approval)
    pub approved_fingerprint: Option<crate::history::HistoryFingerprint>,
    pub approval_history_len: usize,  // History length at time of approval (for drift diff)

    // Cloud sync state
    pub cloud_identity: Option<crate::cloud::CloudIdentity>,
    pub cloud_sync_state: crate::cloud::CloudSyncState,
    pub cloud_upload_generation: u64,
    pub cloud_last_error: Option<String>,
    pub cloud_sheets_list: Vec<crate::cloud::SheetInfo>,
    pub cloud_selected_sheet: Option<usize>,
    pub cloud_sheets_loading: bool,

    // Hub sync state
    pub hub_link: Option<crate::hub::HubLink>,
    pub hub_status: crate::hub::HubStatus,
    pub hub_activity: Option<crate::hub::HubActivity>,
    pub hub_last_check: Option<std::time::Instant>,
    pub hub_last_error: Option<String>,
    pub(crate) hub_check_in_progress: bool,

    // Hub auth/link dialog state
    pub hub_token_input: String,
    pub hub_repos: Vec<crate::hub::RepoInfo>,
    pub hub_selected_repo: Option<usize>,
    pub hub_datasets: Vec<crate::hub::DatasetInfo>,
    pub hub_selected_dataset: Option<usize>,
    pub hub_new_dataset_name: String,
    pub hub_link_loading: bool,

    // Validation dropdown state (data validation list picker)
    pub validation_dropdown: crate::validation_dropdown::ValidationDropdownState,

    // Validation dialog state (Phase 4: Data > Validation menu)
    pub validation_dialog: ValidationDialogState,

    // Paste Special dialog state (Ctrl+Alt+V)
    pub paste_special_dialog: PasteSpecialDialogState,

    // Convert picker dialog state (palette → Convert)
    pub convert_picker_selected: u8,

    // Number Format Editor dialog state (Ctrl+1 escalation)
    pub number_format_editor: NumberFormatEditorState,
    /// Last selected paste type for session memory (remembered within session)
    pub last_paste_special_mode: PasteType,

    // Validation failure navigation (Phase 6B: F8/Shift+F8 to cycle through invalid cells)
    pub validation_failures: Vec<(usize, usize)>,  // (row, col) of failed cells
    pub validation_failure_index: usize,           // Current index for cycling

    // Invalid cell markers (Phase 6C: visible red corner marks)
    pub invalid_cells: std::collections::HashMap<(usize, usize), visigrid_engine::validation::ValidationFailureReason>,

    // Rewind confirmation dialog (Phase 8C: hard rewind)
    pub rewind_confirm: RewindConfirmState,
    // Rewind success banner (Phase 8C: post-rewind feedback)
    pub rewind_success: RewindSuccessBanner,

    // Cycle banner state (cycle/freeze/iteration status)
    pub cycle_banner: CycleBannerState,

    // Merge cells confirmation dialog
    pub merge_confirm: MergeConfirmState,

    // Close-window save confirmation dialog
    pub close_confirm_visible: bool,
    /// Focused button index: 0=Cancel, 1=Don't Save, 2=Save
    pub close_confirm_focused: u8,

    // AI Settings dialog state
    pub ai_settings: AISettingsDialogState,
    pub ask_ai: AskAIDialogState,
    /// Session flag: AI key was validated/set in this session (workaround for keychain timing)
    pub ai_key_validated_this_session: bool,
    /// Cached API key from this session (workaround for keychain timing)
    pub ai_session_key: Option<String>,

    // Session server state (TCP server for external control)
    /// Session server instance (manages TCP listener and discovery file).
    pub session_server: crate::session_server::SessionServer,
    /// Receiver for session requests from TCP server (bridge).
    /// Messages are drained in render() and processed via canonical mutation path.
    pub(crate) session_request_rx: std::sync::mpsc::Receiver<crate::session_server::SessionRequest>,
    /// Sender for session requests (cloned to give to server).
    /// Kept here so we can create bridge handles on demand.
    pub(crate) session_request_tx: std::sync::mpsc::Sender<crate::session_server::SessionRequest>,
    /// Pending pairing approval dialog (client name + reply channel).
    pub pairing_prompt: Option<PairingPrompt>,
    /// Last repeatable command, for F4. See repeat.rs for why it is written
    /// by the mutation methods rather than at a dispatch layer.
    pub repeat_action: Option<crate::repeat::RepeatAction>,
    /// Set while re-applying (or while an agent drives a GUI mutation path)
    /// so the slot is not overwritten by its own replay.
    pub(crate) suppress_repeat_capture: bool,
}

/// Cache for cell search results, invalidated by cells_rev
pub(crate) struct CellSearchCache {
    cached_rev: u64,
    pub(crate) entries: Vec<crate::search::CellEntry>,
}

impl Default for CellSearchCache {
    fn default() -> Self {
        Self {
            cached_rev: 0,
            entries: Vec::new(),
        }
    }
}

/// Cache for named range usage counts, invalidated by cells_rev
pub(crate) struct NamedRangeUsageCache {
    pub(crate) cached_rev: u64,
    /// Map from lowercase name to usage count
    pub(crate) counts: std::collections::HashMap<String, usize>,
}

impl Default for NamedRangeUsageCache {
    fn default() -> Self {
        Self {
            cached_rev: 0,
            counts: std::collections::HashMap::new(),
        }
    }
}

impl Spreadsheet {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let workbook_data = Workbook::new();
        let initial_sheet_id = workbook_data.active_sheet().id;
        let base_workbook = workbook_data.clone(); // Capture initial state for replay
        let workbook = cx.new(|_| workbook_data);

        let focus_handle = cx.focus_handle();
        let console_focus_handle = cx.focus_handle();
        let terminal_focus_handle = cx.focus_handle();
        let script_view_focus_handle = cx.focus_handle();
        let font_picker_focus = cx.focus_handle();
        let ui = UiState {
            color_picker: crate::color_palette::ColorPickerState::new(cx.focus_handle()),
            format_bar: FormatBarState {
                size_input: String::new(),
                size_editing: false,
                size_dropdown: false,
                size_focus: cx.focus_handle(),
                size_replace_next: false,
                number_format_menu_open: false,
                cell_style_menu_open: false,
            },
            format_menu_open: false,
        };
        window.focus(&focus_handle, cx);
        let window_size = window.viewport_size();
        let window_handle = window.window_handle();

        // Get theme from global settings store (resolve "system" to OS-appropriate theme)
        // None (no user preference) is treated as "system" — respects OS dark/light mode.
        let theme = match user_settings(cx).appearance.theme_id.as_value() {
            Some(id) if id == SYSTEM_THEME_ID => {
                let resolved_id = resolve_system_theme_id(window.appearance());
                get_theme(resolved_id).unwrap_or_else(default_theme)
            }
            Some(id) => get_theme(id).unwrap_or_else(default_theme),
            None => {
                let resolved_id = resolve_system_theme_id(window.appearance());
                get_theme(resolved_id).unwrap_or_else(default_theme)
            }
        };

        // Subscribe to global settings changes - trigger re-render when settings change
        let settings_subscription = observe_settings(cx, |cx| {
            // Notify all windows to re-render when settings change
            cx.refresh_windows();
        });

        // Observe OS appearance changes so System theme switches live.
        // Triggers for explicit "system" selection OR when no theme is set (default = system).
        let appearance_subscription = cx.observe_window_appearance(window, |this, window, cx| {
            let is_system = user_settings(cx).appearance.theme_id
                .as_value()
                .map_or(true, |id| id == SYSTEM_THEME_ID);
            if is_system {
                let resolved_id = resolve_system_theme_id(window.appearance());
                if this.theme.meta.id != resolved_id {
                    if let Some(resolved) = get_theme(resolved_id) {
                        this.theme = resolved;
                        cx.notify();
                    }
                }
            }
        });

        // Session server channel: requests from TCP server → GUI thread
        let (session_tx, session_rx) = std::sync::mpsc::channel();
        let session_server = crate::session_server::SessionServer::new();

        let mut app = Self {
            workbook,
            history: History::new(),
            base_workbook,
            rewind_preview: RewindPreviewState::Off,
            cell_metadata: crate::role_styles::CellMetadataMap::new(),
            role_style_map: crate::role_styles::RoleStyleMap::new(),
            row_view: RowView::new(NUM_ROWS),  // Identity mapping, all visible
            filter_state: FilterState::default(),
            filter_dropdown_col: None,
            filter_search_text: String::new(),
            filter_checked_items: std::collections::HashSet::new(),
            view_state: WorkbookViewState::default(),
            pre_row_select_col: None,
            split_pane: None,
            split_active_side: crate::split_view::SplitSide::Left,
            trace_enabled: false,
            trace_cache: None,
            mode: Mode::Navigation,
            edit_value: String::new(),
            edit_cursor: 0,
            edit_selection_anchor: None,
            edit_original: String::new(),
            edit_scroll_x: 0.0,
            edit_scroll_dirty: false,
            caret_visible: true,
            caret_last_activity: std::time::Instant::now(),
            caret_blink_task: None,
            keytips_active: false,
            keytips_deadline_at: None,
            last_keytips_scope: None,
            keytips_hint_shown: false,
            goto_input: String::new(),
            find_input: String::new(),
            find_results: Vec::new(),
            find_index: 0,
            replace_input: String::new(),
            find_replace_mode: false,
            find_focus_replace: false,
            palette_query: String::new(),
            palette_selected: 0,
            palette_scroll_offset: 0,
            palette_scope: None,
            search_engine: Self::create_search_engine(),
            palette_results: Vec::new(),
            palette_total_results: 0,
            palette_pre_selection: (0, 0),
            palette_pre_selection_end: None,
            palette_pre_scroll: (0, 0),
            palette_previewing: false,
            internal_clipboard: None,
            clipboard_visual_range: None,
            session_window_id: WINDOW_ID_UNSET,
            current_file: None,
            is_modified: false,
            close_after_save: false,
            window_handle: window_handle.into(),
            recent_files: Vec::new(),
            recent_commands: Vec::new(),
            document_meta: DocumentMeta::default(),
            cached_title: None,
            pending_title_refresh: false,
            focus_handle,
            console_focus_handle,
            terminal_focus_handle,
            terminal_focused: false,
            script_view_focus_handle,
            font_picker_focus,
            ui,
            status_message: None,
            window_size,
            cached_window_bounds: Some(window.window_bounds()),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            hidden_rows: HashMap::new(),
            hidden_cols: HashMap::new(),
            cached_sheet_id: initial_sheet_id,
            resizing_col: None,
            resizing_row: None,
            resize_start_pos: 0.0,
            resize_start_size: 0.0,
            resize_start_original: None,
            open_menu: None,
            menu_highlight: None,
            renaming_sheet: None,
            sheet_rename_input: String::new(),
            sheet_rename_cursor: 0,
            sheet_rename_select_all: false,
            sheet_context_menu: None,
            context_menu: None,
            available_fonts: Self::enumerate_fonts(),
            font_picker_query: String::new(),
            font_picker_selected: 0,
            font_picker_scroll_offset: 0,
            theme_picker_query: String::new(),
            theme_picker_selected: 0,
            dragging_selection: false,
            fill_drag: FillDrag::None,
            dragging_row_header: false,
            dragging_col_header: false,
            row_header_anchor: None,
            col_header_anchor: None,
            grid_layout: GridLayout::default(),
            formula_ref_cell: None,
            formula_ref_end: None,
            formula_ref_start_cursor: 0,
            formula_nav_mode: crate::mode::FormulaNavMode::default(),
            formula_nav_manual_override: None,
            formula_home_sheet: None,
            formula_edit_cell: None,
            formula_ref_sheet: None,
            formula_cross_sheet_name: None,
            formula_highlighted_refs: Vec::new(),
            formula_ref_color_map: std::collections::HashMap::new(),
            formula_ref_next_color: 0,
            formula_bar_cache_cell: None,
            formula_bar_cache_formula: String::new(),
            formula_bar_cache_refs: Vec::new(),
            active_editor: EditorSurface::Cell,
            formula_bar_scroll_x: 0.0,
            formula_bar_text_rect: gpui::Bounds::default(),
            formula_bar_cache_dirty: false,
            formula_bar_char_boundaries: Vec::new(),
            formula_bar_boundary_xs: Vec::new(),
            formula_bar_text_width: 0.0,
            formula_bar_drag_anchor: None,
            formula_bar_expanded: false,
            name_box_editing: false,
            name_box_input: String::new(),
            name_box_focus: cx.focus_handle(),
            name_box_replace_next: false,
            autocomplete_visible: false,
            autocomplete_suppressed: false,
            autocomplete_selected: 0,
            autocomplete_replace_range: 0..0,
            hover_function: None,
            doc_settings: crate::settings::DocumentSettings::default(),
            minimap_visible: false,
            minimap_cache: crate::minimap::MinimapCache::default(),
            minimap_dragging: false,
            minimap_drag_offset_y: 0.0,
            profiler_visible: false,
            profiler_report: None,
            profiler_hotspots: Vec::new(),
            profiler_capture_next: false,
            locked_panels_dismissed: false,
            inspector_visible: false,
            inspector_tab: crate::mode::InspectorTab::default(),
            inspector_pinned: None,
            format_painter: None,
            current_border_color: None,  // Automatic (theme default)
            tab_chain_origin_col: None,
            inspector_hover_cell: None,
            inspector_trace_path: None,
            inspector_trace_incomplete: false,
            names_filter_query: String::new(),
            selected_named_range: None,
            selected_history_id: None,
            history_filter_query: String::new(),
            history_filter_mode: HistoryFilterMode::default(),
            history_view_start: 0,
            history_highlight_range: None,
            diff_report: None,
            diff_ai_only_filter: false,
            diff_selected_entry: None,
            diff_ai_summary: None,
            diff_ai_summary_loading: false,
            diff_ai_summary_error: None,
            diff_entry_explanations: std::collections::HashMap::new(),
            diff_explaining_entry: None,
            history_context_menu_entry_id: None,
            transform_preview: None,
            theme,
            theme_preview: None,
            cells_rev: 1,  // Start at 1 so cache (starting at 0) is immediately stale
            cell_search_cache: CellSearchCache::default(),
            named_range_usage_cache: NamedRangeUsageCache::default(),
            rename_original_name: String::new(),
            rename_new_name: String::new(),
            rename_select_all: false,
            rename_affected_cells: Vec::new(),
            rename_validation_error: None,
            cf_input: String::new(),
            cf_input_error: None,
            cf_target: Vec::new(),
            cf_preview_id: None,
            cf_preview_matches: None,
            cf_panel_visible: false,
            cf_edit_backup: None,
            cf_rules_rev: 1,
            cf_cache: std::cell::RefCell::new(std::collections::HashMap::new()),
            cf_cache_key: std::cell::Cell::new((0, 0)),
            create_name_name: String::new(),
            create_name_description: String::new(),
            create_name_target: String::new(),
            create_name_validation_error: None,
            create_name_focus: CreateNameFocus::default(),

            edit_description_name: String::new(),
            edit_description_value: String::new(),
            edit_description_original: None,

            tour_step: 0,
            tour_completed: false,
            show_f2_tip: false,
            settings_subscription,
            appearance_subscription: Some(appearance_subscription),

            impact_preview_action: None,
            impact_preview_usages: Vec::new(),

            refactor_log: Vec::new(),

            extract_range_literal: String::new(),
            extract_name: String::new(),
            extract_description: String::new(),
            extract_affected_cells: Vec::new(),
            extract_occurrence_count: 0,
            extract_validation_error: None,
            extract_select_all: false,
            extract_focus: CreateNameFocus::default(),

            import_result: None,
            import_filename: None,
            import_source_dir: None,

            import_in_progress: false,
            import_overlay_visible: false,
            import_started_at: None,

            startup_instant: None,
            cold_start_ms: None,

            export_result: None,
            export_filename: None,

            hint_state: crate::hints::HintState::default(),

            zen_mode: false,
            f1_help_visible: false,
            metrics: GridMetrics::default(),
            debug_grid_alignment: false,
            #[cfg(debug_assertions)]
            debug_border_call_count: std::cell::Cell::new(0),
            #[cfg(debug_assertions)]
            debug_gridline_cells: std::cell::Cell::new(0),
            #[cfg(debug_assertions)]
            debug_userborder_cells: std::cell::Cell::new(0),
            #[cfg(debug_assertions)]
            debug_border_frames: std::cell::Cell::new(0),
            #[cfg(debug_assertions)]
            debug_border_last_report: std::cell::Cell::new(std::time::Instant::now()),
            #[cfg(debug_assertions)]
            debug_border_stale_streak: 0,
            zoom_wheel_accumulator: 0.0,
            pending_nav_dx: 0,
            pending_nav_dy: 0,
            nav_scroll_dirty: false,
            nav_perf: crate::perf::NavLatencyTracker::default(),
            link_open_in_flight: false,

            bottom_panel_visible: false,
            bottom_panel_tab: BottomPanelTab::default(),

            terminal: crate::terminal::TerminalState::default(),

            lua_runtime: crate::scripting::LuaRuntime::default(),
            lua_console: crate::scripting::ConsoleState::default(),
            script: crate::scripting::ScriptState::default(),
            custom_fn_registry: crate::scripting::CustomFunctionRegistry::empty(),

            attached_scripts: Vec::new(),
            pending_run_records: Vec::new(),
            loaded_run_records: Vec::new(),

            license_input: String::new(),
            license_error: None,
            trial_confirm_visible: false,

            default_app_prompt_state: DefaultAppPromptState::Hidden,
            default_app_prompt_file_type: None,
            default_app_prompt_success_timer: None,
            needs_settings_entered_at: None,
            needs_settings_check_count: 0,

            in_smoke_recalc: false,

            verified_mode: false,
            last_recalc_report: None,

            semantic_verification: visigrid_io::native::SemanticVerification::default(),
            approval_confirm_visible: false,
            approval_drift_visible: false,
            approval_label_input: String::new(),
            approved_fingerprint: None,
            approval_history_len: 0,

            cloud_identity: None,
            cloud_sync_state: crate::cloud::CloudSyncState::Local,
            cloud_upload_generation: 0,
            cloud_last_error: None,
            cloud_sheets_list: Vec::new(),
            cloud_selected_sheet: None,
            cloud_sheets_loading: false,

            hub_link: None,
            hub_status: crate::hub::HubStatus::Unlinked,
            hub_activity: None,
            hub_last_check: None,
            hub_last_error: None,
            hub_check_in_progress: false,

            hub_token_input: String::new(),
            hub_repos: Vec::new(),
            hub_selected_repo: None,
            hub_datasets: Vec::new(),
            hub_selected_dataset: None,
            hub_new_dataset_name: String::new(),
            hub_link_loading: false,

            validation_dropdown: crate::validation_dropdown::ValidationDropdownState::default(),

            validation_dialog: ValidationDialogState::default(),

            paste_special_dialog: PasteSpecialDialogState::default(),
            convert_picker_selected: 0,
            number_format_editor: NumberFormatEditorState::default(),
            last_paste_special_mode: PasteType::All,

            validation_failures: Vec::new(),
            validation_failure_index: 0,

            invalid_cells: std::collections::HashMap::new(),

            rewind_confirm: RewindConfirmState::default(),
            rewind_success: RewindSuccessBanner::default(),
            cycle_banner: CycleBannerState::default(),

            merge_confirm: MergeConfirmState::default(),
            close_confirm_visible: false,
            close_confirm_focused: 2, // Default to Save button

            ai_settings: AISettingsDialogState::default(),
            ask_ai: AskAIDialogState::default(),
            ai_key_validated_this_session: false,
            ai_session_key: None,

            // Session server: initialized below
            session_server: session_server,
            session_request_rx: session_rx,
            pairing_prompt: None,
            repeat_action: None,
            suppress_repeat_capture: false,
            session_request_tx: session_tx,
        };

        // Load custom functions from ~/.config/visigrid/functions.lua
        match crate::scripting::custom_functions::load_custom_functions(app.lua_runtime.lua()) {
            Ok(registry) => {
                if !registry.functions.is_empty() {
                    app.status_message = Some(format!(
                        "Loaded {} custom function{}",
                        registry.functions.len(),
                        if registry.functions.len() == 1 { "" } else { "s" },
                    ));
                }
                app.custom_fn_registry = registry;
            }
            Err(e) => {
                eprintln!("Custom functions error: {}", e);
            }
        }

        app
    }

    // ========================================================================
    // Terminal Panel
    // ========================================================================




    // ========================================================================
    // Session Server
    // ========================================================================













    /// End a workbook batch and broadcast changes to session server subscribers.
    ///
    /// This is the canonical way to end a batch when session server may be running.
    /// It ensures all mutation paths (user edits, paste, import, session ops) broadcast
    /// their changes to subscribers.
    ///
    /// Returns the number of changed cells (0 if no changes or nested batch).
    pub fn end_batch_and_broadcast(&mut self, cx: &mut Context<Self>) -> usize {
        let changed = self.workbook.update(cx, |wb, _| wb.end_batch());
        let count = changed.len();

        if !changed.is_empty() && self.session_server.is_running() {
            let revision = self.workbook.read(cx).revision();
            let cells: Vec<crate::session_server::CellRef> = changed
                .into_iter()
                .map(|c| crate::session_server::CellRef {
                    sheet: c.sheet.0 as usize, // SheetId(u64) -> usize
                    row: c.row,
                    col: c.col,
                })
                .collect();
            self.session_server.broadcast_cells(revision, cells);
        }

        count
    }

    /// Get the active theme (preview if set, otherwise current)
    pub fn active_theme(&self) -> &Theme {
        self.theme_preview.as_ref().unwrap_or(&self.theme)
    }

    /// Get a theme token color
    pub fn token(&self, key: TokenKey) -> Hsla {
        self.active_theme().get(key)
    }

    // ========================================================================
    // Validation Dropdown
    // ========================================================================

    /// Close the validation dropdown if open.
    ///
    /// Call this from all invalidation points:
    /// - Selection change
    /// - Sheet switch
    /// - Scroll/zoom
    /// - Modal open
    /// - Click outside
    pub fn close_validation_dropdown(
        &mut self,
        reason: crate::validation_dropdown::DropdownCloseReason,
        cx: &mut Context<Self>,
    ) {
        use crate::validation_dropdown::DropdownCloseReason;

        if self.validation_dropdown.is_open() {
            self.validation_dropdown.close();

            // Show status message for specific close reasons
            if reason == DropdownCloseReason::SourceChanged {
                self.status_message = Some("List updated".to_string());
            }

            cx.notify();
        }
    }

    /// Check if validation dropdown is open
    pub fn is_validation_dropdown_open(&self) -> bool {
        self.validation_dropdown.is_open()
    }

    /// Open dropdown for the current cell (Alt+Down - Excel muscle memory)
    ///
    /// Priority:
    /// 1. If cell has list validation → open validation dropdown
    /// 2. If column has AutoFilter active → open filter dropdown
    /// 3. Else → show "No dropdown" message
    pub fn open_validation_dropdown(&mut self, cx: &mut Context<Self>) {
        use crate::validation_dropdown::ValidationDropdownState;

        let (row, col) = self.view_state.selected;
        let sheet_index = self.sheet_index(cx);

        // Priority 1: Check for list validation
        let resolved = self.wb(cx).get_list_items(sheet_index, row, col);
        match resolved {
            Some(list) if !list.items.is_empty() => {
                // Open validation dropdown
                self.validation_dropdown = ValidationDropdownState::open(
                    (row, col),
                    std::sync::Arc::new(list),
                );
                cx.notify();
                return;
            }
            Some(_) => {
                // Has list validation but list is empty
                self.status_message = Some("Validation list is empty".to_string());
                cx.notify();
                return;
            }
            None => {
                // No list validation - fall through to check filter
            }
        };

        // Priority 2: Check for AutoFilter on this column
        if self.filter_state.is_enabled() {
            if let Some((_, col_start, _, col_end)) = self.filter_state.data_range() {
                if col >= col_start && col <= col_end {
                    // Column is in filter range - open filter dropdown
                    self.open_filter_dropdown(col, cx);
                    return;
                }
            }
        }

        // No dropdown available
        self.status_message = Some("No dropdown available".to_string());
        cx.notify();
    }

    /// Check if the validation dropdown source has changed (fingerprint mismatch).
    /// Call this during render or update cycle to detect stale data.
    pub fn check_dropdown_staleness(&mut self, cx: &mut Context<Self>) {
        use crate::validation_dropdown::DropdownCloseReason;

        let open_state = match self.validation_dropdown.as_open() {
            Some(state) => state,
            None => return,
        };

        let (row, col) = open_state.cell;
        let stored_fingerprint = open_state.source_fingerprint;
        let sheet_index = self.sheet_index(cx);

        // Get current fingerprint from source
        if let Some(current_list) = self.wb(cx).get_list_items(sheet_index, row, col) {
            if current_list.source_fingerprint != stored_fingerprint {
                self.close_validation_dropdown(DropdownCloseReason::SourceChanged, cx);
            }
        } else {
            // Source no longer exists - close dropdown
            self.close_validation_dropdown(DropdownCloseReason::SourceChanged, cx);
        }
    }

    /// Route a key event through the dropdown first.
    ///
    /// Returns true if the event was consumed (dropdown handled it).
    /// Call this BEFORE any other key handling.
    pub fn route_dropdown_key_event(
        &mut self,
        key: &str,
        modifiers: crate::validation_dropdown::KeyModifiers,
        cx: &mut Context<Self>,
    ) -> bool {
        use crate::validation_dropdown::DropdownOutcome;

        let open_state = match self.validation_dropdown.as_open_mut() {
            Some(state) => state,
            None => return false, // Dropdown not open
        };

        let outcome = open_state.handle_key(key, modifiers);

        match outcome {
            DropdownOutcome::Consumed => {
                cx.notify();
                true
            }
            DropdownOutcome::CloseNoCommit => {
                self.validation_dropdown.close();
                cx.notify();
                // For Tab, return false so grid can handle navigation
                key == "tab"
            }
            DropdownOutcome::CommitValue(value) => {
                // Use the same commit path as click-to-select (undo, dep graph)
                self.commit_validation_value(&value, cx);
                true
            }
            DropdownOutcome::NotConsumed => false,
        }
    }

    /// Route a character input through the dropdown first.
    ///
    /// Returns true if the event was consumed.
    pub fn route_dropdown_char_event(
        &mut self,
        ch: char,
        cx: &mut Context<Self>,
    ) -> bool {
        use crate::validation_dropdown::DropdownOutcome;

        let open_state = match self.validation_dropdown.as_open_mut() {
            Some(state) => state,
            None => return false,
        };

        let outcome = open_state.handle_char(ch);

        match outcome {
            DropdownOutcome::Consumed => {
                cx.notify();
                true
            }
            _ => false,
        }
    }

    /// Commit a value from the validation dropdown (called when user clicks an item).
    ///
    /// Uses the same undo/recalc pipeline as normal cell editing to ensure:
    /// - Undo/redo works correctly
    /// - Dependency graph is updated
    /// - Dirty state is tracked via history
    pub fn commit_validation_value(&mut self, value: &str, cx: &mut Context<Self>) {
        use crate::validation_dropdown::DropdownCloseReason;

        // Close dropdown first
        self.close_validation_dropdown(DropdownCloseReason::Committed, cx);

        // Commit value using the same path as normal cell editing
        let (row, col) = self.view_state.selected;
        let old_value = self.sheet(cx).get_raw(row, col);

        // Record for undo (same as confirm_edit)
        self.history.record_change(self.sheet_index(cx), row, col, old_value, value.to_string());

        // Set value and update dependency graph (same as confirm_edit)
        self.set_cell_value(row, col, value, cx);

        // Bump revision for render cache invalidation
        self.cells_rev = self.cells_rev.wrapping_add(1);
        cx.notify();
    }

    // ========================================================================
    // Validation Failure Navigation (Phase 6B: F8/Shift+F8)
    // ========================================================================

    /// Store validation failures for navigation and display.
    /// Called after paste/fill operations that may cause validation failures.
    /// Populates both the navigation list (F8) and the invalid_cells map (red markers).
    pub fn store_validation_failures(&mut self, failures: &visigrid_engine::workbook::ValidationFailures) {
        // Store for F8 navigation
        self.validation_failures = failures.failures.iter()
            .map(|f| (f.row, f.col))
            .collect();
        self.validation_failure_index = 0;

        // Store for red corner markers (adds to existing, doesn't clear)
        for f in &failures.failures {
            self.invalid_cells.insert((f.row, f.col), f.reason);
        }
    }

    /// Clear all invalid cell markers.
    pub fn clear_invalid_circles(&mut self, cx: &mut Context<Self>) {
        let count = self.invalid_cells.len();
        self.invalid_cells.clear();
        self.validation_failures.clear();
        self.validation_failure_index = 0;
        self.status_message = Some(format!("Cleared {} invalid cell markers", count));
        cx.notify();
    }

    /// Circle Invalid Data: validate all cells with validation rules and mark invalid ones.
    pub fn circle_invalid_data(&mut self, cx: &mut Context<Self>) {
        use visigrid_engine::validation::ValidationResult;
        use visigrid_engine::workbook::Workbook;

        // Clear existing markers
        self.invalid_cells.clear();
        self.validation_failures.clear();
        self.validation_failure_index = 0;

        // Collect validation ranges first (to avoid borrow conflict)
        let ranges: Vec<_> = self.sheet(cx).validations.iter()
            .map(|(range, _)| range.clone())
            .collect();

        // Validate each cell with a rule
        let sheet_idx = self.sheet_index(cx);
        for target in ranges {
            for row in target.start_row..=target.end_row {
                for col in target.start_col..=target.end_col {
                    let display_value = self.sheet(cx).get_display(row, col);
                    // Skip empty cells
                    if display_value.is_empty() {
                        continue;
                    }
                    let result = self.wb(cx).validate_cell_input(sheet_idx, row, col, &display_value);
                    if let ValidationResult::Invalid { reason, .. } = result {
                        // Classify the failure reason
                        let failure_reason = Workbook::classify_failure_reason(&reason);
                        self.invalid_cells.insert((row, col), failure_reason);
                        self.validation_failures.push((row, col));
                    }
                }
            }
        }

        // Sort failures in row-major order for predictable navigation
        self.validation_failures.sort_by_key(|&(r, c)| (r, c));

        let count = self.invalid_cells.len();
        if count == 0 {
            self.status_message = Some("All cells are valid".to_string());
        } else {
            self.status_message = Some(format!("Found {} invalid cells. Press F8 to navigate.", count));
        }
        cx.notify();
    }

    /// Check if a cell is marked as invalid (for rendering).
    pub fn is_cell_invalid(&self, row: usize, col: usize) -> bool {
        self.invalid_cells.contains_key(&(row, col))
    }

    /// Get the invalid reason for a cell (for inspector).
    pub fn get_invalid_reason(&self, row: usize, col: usize) -> Option<visigrid_engine::validation::ValidationFailureReason> {
        self.invalid_cells.get(&(row, col)).copied()
    }

    /// Clear invalid marker for a specific cell (called when cell is edited to valid value).
    pub fn clear_cell_invalid(&mut self, row: usize, col: usize) {
        self.invalid_cells.remove(&(row, col));
        // Also remove from navigation list
        self.validation_failures.retain(|&(r, c)| r != row || c != col);
        // Adjust index if needed
        if !self.validation_failures.is_empty() && self.validation_failure_index >= self.validation_failures.len() {
            self.validation_failure_index = 0;
        }
    }

    /// Jump to the next invalid cell (F8).
    pub fn next_invalid_cell(&mut self, cx: &mut Context<Self>) {
        if self.validation_failures.is_empty() {
            self.status_message = Some("No validation failures to navigate".to_string());
            cx.notify();
            return;
        }

        // Move to next failure (with wrap-around)
        self.validation_failure_index = (self.validation_failure_index + 1) % self.validation_failures.len();
        let (row, col) = self.validation_failures[self.validation_failure_index];

        // Select the cell and scroll into view
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.ensure_visible(cx);

        // Get failure reason for status message
        let reason_str = self.invalid_cells.get(&(row, col))
            .map(|r| Self::failure_reason_short(*r))
            .unwrap_or_default();

        self.status_message = Some(format!(
            "Invalid {} of {}: {} — F8 next, Shift+F8 prev",
            self.validation_failure_index + 1,
            self.validation_failures.len(),
            reason_str
        ));
        cx.notify();
    }

    /// Short human-readable description of validation failure reason.
    fn failure_reason_short(reason: visigrid_engine::validation::ValidationFailureReason) -> String {
        use visigrid_engine::validation::ValidationFailureReason;
        match reason {
            ValidationFailureReason::InvalidValue => "Value doesn't match rule".to_string(),
            ValidationFailureReason::ConstraintBlank => "Constraint cell is blank".to_string(),
            ValidationFailureReason::ConstraintNotNumeric => "Constraint is not numeric".to_string(),
            ValidationFailureReason::InvalidReference => "Invalid reference".to_string(),
            ValidationFailureReason::FormulaNotSupported => "Formula constraint not supported".to_string(),
            ValidationFailureReason::ListEmpty => "List is empty".to_string(),
            ValidationFailureReason::NotInList => "Not in list".to_string(),
        }
    }

    /// Jump to the previous invalid cell (Shift+F8).
    pub fn prev_invalid_cell(&mut self, cx: &mut Context<Self>) {
        if self.validation_failures.is_empty() {
            self.status_message = Some("No validation failures to navigate".to_string());
            cx.notify();
            return;
        }

        // Move to previous failure (with wrap-around)
        if self.validation_failure_index == 0 {
            self.validation_failure_index = self.validation_failures.len() - 1;
        } else {
            self.validation_failure_index -= 1;
        }
        let (row, col) = self.validation_failures[self.validation_failure_index];

        // Select the cell and scroll into view
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.ensure_visible(cx);

        // Get failure reason for status message
        let reason_str = self.invalid_cells.get(&(row, col))
            .map(|r| Self::failure_reason_short(*r))
            .unwrap_or_default();

        self.status_message = Some(format!(
            "Invalid {} of {}: {} — F8 next, Shift+F8 prev",
            self.validation_failure_index + 1,
            self.validation_failures.len(),
            reason_str
        ));
        cx.notify();
    }

    // ========================================================================
    // Document settings accessors (resolve Setting<T> to concrete values)
    // ========================================================================

    /// Whether to show formulas instead of calculated values
    pub fn show_formulas(&self) -> bool {
        use crate::settings::Setting;
        match &self.doc_settings.display.show_formulas {
            Setting::Value(v) => *v,
            Setting::Inherit => false, // Default: show values, not formulas
        }
    }

    /// Whether to show zero values (vs blank cells)
    pub fn show_zeros(&self) -> bool {
        use crate::settings::Setting;
        match &self.doc_settings.display.show_zeros {
            Setting::Value(v) => *v,
            Setting::Inherit => true, // Default: show zeros (like Excel)
        }
    }

    /// Toggle the show_formulas document setting
    pub fn toggle_show_formulas(&mut self, cx: &mut Context<Self>) {
        use crate::settings::Setting;
        let current = self.show_formulas();
        self.doc_settings.display.show_formulas = Setting::Value(!current);
        self.save_doc_settings_if_needed();
        cx.notify();
    }

    /// Toggle the show_zeros document setting
    pub fn toggle_show_zeros(&mut self, cx: &mut Context<Self>) {
        use crate::settings::Setting;
        let current = self.show_zeros();
        self.doc_settings.display.show_zeros = Setting::Value(!current);
        self.save_doc_settings_if_needed();
        cx.notify();
    }

    /// Toggle the format bar visibility (user setting, persisted)
    pub fn toggle_format_bar(&mut self, cx: &mut Context<Self>) {
        use crate::settings::Setting;
        let current = match &user_settings(cx).appearance.show_format_bar {
            Setting::Value(v) => *v,
            Setting::Inherit => true,
        };
        update_user_settings(cx, |s| {
            s.appearance.show_format_bar = Setting::Value(!current);
        });
        cx.notify();
    }

    // =========================================================================
    // Zoom
    // =========================================================================

    /// Set zoom level (all zoom changes go through this)
    pub fn set_zoom(&mut self, new_zoom: f32, cx: &mut Context<Self>) {
        // Close validation dropdown on zoom
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::Zoom,
            cx,
        );

        // Clamp to valid range
        let clamped = new_zoom.max(ZOOM_STEPS[0]).min(ZOOM_STEPS[ZOOM_STEPS.len() - 1]);
        if (clamped - self.view_state.zoom_level).abs() < 0.001 {
            return; // No change
        }
        self.view_state.zoom_level = clamped;
        self.metrics = GridMetrics::with_scale(clamped, self.metrics.scale);
        self.ensure_visible(cx);
        // Show status message
        let percent = (clamped * 100.0).round() as i32;
        self.status_message = Some(format!("Zoom: {}%", percent));
        cx.notify();
    }

    /// Zoom in to next step on the ladder
    pub fn zoom_in(&mut self, cx: &mut Context<Self>) {
        if let Some(&next) = ZOOM_STEPS.iter().find(|&&z| z > self.view_state.zoom_level + 0.001) {
            self.set_zoom(next, cx);
        }
    }

    /// Zoom out to previous step on the ladder
    pub fn zoom_out(&mut self, cx: &mut Context<Self>) {
        if let Some(&prev) = ZOOM_STEPS.iter().rev().find(|&&z| z < self.view_state.zoom_level - 0.001) {
            self.set_zoom(prev, cx);
        }
    }

    /// Reset zoom to 100%
    pub fn zoom_reset(&mut self, cx: &mut Context<Self>) {
        self.set_zoom(DEFAULT_ZOOM, cx);
    }

    /// Handle zoom from mouse wheel (with debounce/accumulation)
    pub fn zoom_wheel(&mut self, delta_y: f32, cx: &mut Context<Self>) {
        // Accumulate wheel delta - threshold before stepping
        self.zoom_wheel_accumulator += delta_y;
        let threshold = 50.0; // Pixels of wheel movement to trigger one step
        if self.zoom_wheel_accumulator > threshold {
            self.zoom_wheel_accumulator = 0.0;
            self.zoom_out(cx);
        } else if self.zoom_wheel_accumulator < -threshold {
            self.zoom_wheel_accumulator = 0.0;
            self.zoom_in(cx);
        }
    }

    /// Get zoom percentage for display (e.g., "100%")
    pub fn zoom_display(&self) -> String {
        let percent = (self.view_state.zoom_level * 100.0).round() as i32;
        format!("{}%", percent)
    }

    /// Enumerate available system fonts.
    ///
    /// Uses platform-native APIs where available (macOS Core Text, Linux fontconfig),
    /// with hardcoded fallbacks for safety.
    fn enumerate_fonts() -> Vec<String> {
        let mut fonts = Self::enumerate_system_fonts();
        fonts.sort();
        fonts.dedup();
        // Filter out hidden/internal fonts (starting with '.' or '#')
        fonts.retain(|f| !f.starts_with('.') && !f.starts_with('#') && !f.is_empty());
        fonts
    }

    #[cfg(target_os = "macos")]
    fn enumerate_system_fonts() -> Vec<String> {
        use core_text::font_manager;

        let cf_names = font_manager::copy_available_font_family_names();
        let count = cf_names.len();
        let mut names = Vec::with_capacity(count as usize);
        for i in 0..count {
            if let Some(name) = cf_names.get(i) {
                let s: String = name.to_string();
                if !s.is_empty() {
                    names.push(s);
                }
            }
        }

        if names.is_empty() {
            // Fallback if Core Text fails
            return vec![
                "Menlo".into(), "Monaco".into(), "Courier New".into(),
                "Helvetica".into(), "Arial".into(), "Times New Roman".into(),
                "Georgia".into(), "Verdana".into(),
            ];
        }

        names
    }

    #[cfg(target_os = "linux")]
    fn enumerate_system_fonts() -> Vec<String> {
        // Use fontconfig CLI (standard on Linux desktops)
        if let Ok(output) = std::process::Command::new("fc-list")
            .args([":family", "--format=%{family}\n"])
            .output()
        {
            if output.status.success() {
                let text = String::from_utf8_lossy(&output.stdout);
                let names: Vec<String> = text
                    .lines()
                    .filter(|l| !l.is_empty())
                    // fc-list returns comma-separated variants; take first
                    .map(|l| l.split(',').next().unwrap_or(l).trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();

                if !names.is_empty() {
                    return names;
                }
            }
        }

        // Fallback
        vec![
            "DejaVu Sans".into(), "DejaVu Sans Mono".into(), "DejaVu Serif".into(),
            "Liberation Mono".into(), "Liberation Sans".into(), "Liberation Serif".into(),
            "Noto Sans".into(), "Noto Sans Mono".into(),
        ]
    }

    #[cfg(target_os = "windows")]
    fn enumerate_system_fonts() -> Vec<String> {
        // No easy zero-dep enumeration on Windows; use safe defaults
        // These fonts ship with every Windows installation since Vista+
        vec![
            "Consolas".into(), "Cascadia Mono".into(), "Courier New".into(),
            "Arial".into(), "Calibri".into(), "Cambria".into(),
            "Times New Roman".into(), "Georgia".into(), "Verdana".into(),
            "Segoe UI".into(), "Tahoma".into(), "Trebuchet MS".into(),
            "Lucida Console".into(), "Comic Sans MS".into(),
        ]
    }

    #[cfg(not(any(target_os = "macos", target_os = "linux", target_os = "windows")))]
    fn enumerate_system_fonts() -> Vec<String> {
        vec![
            "Courier New".into(), "Arial".into(), "Times New Roman".into(),
            "Georgia".into(), "Verdana".into(),
        ]
    }

    /// Create and configure the search engine with all providers
    fn create_search_engine() -> SearchEngine {
        use crate::search::{FormulaSearchProvider, SettingsSearchProvider};
        let mut engine = SearchEngine::new();
        engine.register(Box::new(CommandSearchProvider));
        engine.register(Box::new(GoToSearchProvider));
        engine.register(Box::new(FormulaSearchProvider));
        engine.register(Box::new(SettingsSearchProvider));
        engine
    }

    /// Bump the cell revision counter (call after any cell value change)
    /// This invalidates the cell search cache, ensuring fresh results.
    #[inline]
    pub(crate) fn bump_cells_rev(&mut self) {
        self.cells_rev = self.cells_rev.wrapping_add(1);
    }

    pub(crate) fn bump_cf_rules_rev(&mut self) {
        self.cf_rules_rev = self.cf_rules_rev.wrapping_add(1);
    }

    /// Effective format (base + conditional rules) with per-cell caching.
    /// Cache is invalidated wholesale when cell contents or rules change.
    pub(crate) fn effective_format_cached(&self, row: usize, col: usize, cx: &App) -> visigrid_engine::cell::CellFormat {
        let sheet = self.sheet(cx);
        let base = sheet.get_format(row, col);
        if !sheet.cond_formats.any_rule_covers(row, col) {
            return base;
        }

        let key = (self.cells_rev, self.cf_rules_rev);
        if self.cf_cache_key.get() != key {
            self.cf_cache.borrow_mut().clear();
            self.cf_cache_key.set(key);
        }

        let cached = self.cf_cache.borrow().get(&(row, col)).cloned();
        let override_opt = match cached {
            Some(ov) => ov,
            None => {
                let ov = sheet.cond_formats.override_for_cell(row, col, sheet);
                self.cf_cache.borrow_mut().insert((row, col), ov.clone());
                ov
            }
        };

        match override_opt {
            Some(ov) => {
                let mut f = base;
                ov.apply_to(&mut f);
                f
            }
            None => base,
        }
    }

    /// Ensure cell search cache is fresh (rebuilds if cells_rev changed)
    /// Returns a reference to the cached entries.
    pub(crate) fn ensure_cell_search_cache_fresh(&mut self, cx: &App) -> &[crate::search::CellEntry] {
        use crate::search::CellEntry;
        use visigrid_engine::cell::CellValue;

        if self.cell_search_cache.cached_rev != self.cells_rev {
            // Cache is stale, rebuild from sparse storage
            let sheet = self.sheet(cx);
            let entries: Vec<CellEntry> = sheet.cells_iter()
                .filter(|(_, cell)| !matches!(cell.value, CellValue::Empty))
                .take(1000)  // Cap cells scanned for performance
                .map(|(&(row, col), cell)| {
                    let display = sheet.get_display(row, col);
                    let formula = match &cell.value {
                        CellValue::Formula { source, .. } => Some(source.clone()),
                        _ => None,
                    };
                    CellEntry::new(row, col, display, formula)
                })
                .collect();

            self.cell_search_cache.entries = entries;
            self.cell_search_cache.cached_rev = self.cells_rev;
        }

        &self.cell_search_cache.entries
    }

    /// Execute a search action from the command palette
    pub fn dispatch_action(&mut self, action: SearchAction, window: &mut Window, cx: &mut Context<Self>) {
        match action {
            SearchAction::RunCommand(cmd) => self.dispatch_command(cmd, window, cx),
            SearchAction::JumpToCell { row, col } => {
                self.view_state.selected = (row, col);
                self.view_state.selection_end = None;
                self.ensure_cell_visible(row, col);
                cx.notify();
            }
            SearchAction::InsertFormula { name, signature } => {
                // Context-aware insertion
                if self.mode.is_formula() || (self.mode.is_editing() && self.edit_value.starts_with('=')) {
                    // Already editing a formula: insert function name at cursor (byte-indexed)
                    let func_text = format!("{}(", name);
                    let cursor_byte = self.edit_cursor.min(self.edit_value.len());
                    let before = &self.edit_value[..cursor_byte];
                    let after = &self.edit_value[cursor_byte..];
                    self.edit_value = format!("{}{}{}", before, func_text, after);
                    self.edit_cursor += func_text.len();  // Byte length
                } else {
                    // Grid navigation: start formula edit with =FUNC(
                    self.edit_original = self.sheet(cx).get_raw(self.view_state.selected.0, self.view_state.selected.1);
                    self.edit_value = format!("={}(", name);
                    self.edit_cursor = self.edit_value.len();  // Byte offset at end
                    self.mode = Mode::Formula;
                }
                // Show signature in status for reference
                self.status_message = Some(signature);
                cx.notify();
            }
            SearchAction::OpenFile(path) => {
                self.load_file(&path, cx);
            }
            SearchAction::JumpToNamedRange { .. } => {
                // Future: implement named range navigation
            }
            SearchAction::OpenSetting { key } => {
                // Copy key to clipboard so user doesn't have to hunt
                cx.write_to_clipboard(ClipboardItem::new_string(key.clone()));

                // Open settings file in system editor
                let filename = user_settings_path()
                    .and_then(|p| p.file_name().map(|n| n.to_string_lossy().into_owned()))
                    .unwrap_or_else(|| "settings.json".to_string());

                match open_settings_file() {
                    Ok(()) => {
                        self.status_message = Some(format!(
                            "Copied \"{}\" to clipboard — paste into {}",
                            key, filename
                        ));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Failed to open settings: {}", e));
                    }
                }
                cx.notify();
            }
            SearchAction::CopyToClipboard { text, description } => {
                cx.write_to_clipboard(ClipboardItem::new_string(text));
                self.status_message = Some(description);
                cx.notify();
            }
            SearchAction::ShowFunctionHelp { name, signature, description } => {
                // Show detailed function help in status
                self.status_message = Some(format!("{}{} — {}", name, signature, description));
                cx.notify();
            }
            SearchAction::ShowReferences { row, col } => {
                self.show_references(row, col, cx);
            }
            SearchAction::ShowPrecedents { row, col } => {
                self.show_precedents(row, col, cx);
            }
        }
    }

    /// Execute a command by its stable ID
    pub fn dispatch_command(&mut self, cmd: CommandId, window: &mut Window, cx: &mut Context<Self>) {
        // Track as recently used command
        self.add_recent_command(cmd.clone());

        match cmd {
            // Navigation
            CommandId::GoToCell => self.show_goto(cx),
            CommandId::FindInCells => self.show_find(cx),
            CommandId::GoToStart => {
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                cx.notify();
            }
            CommandId::SelectAll => self.select_all(cx),
            CommandId::AddConditionalFormat => self.show_add_cond_format(cx),
            CommandId::ManageConditionalFormats => self.toggle_cf_panel(cx),
            CommandId::ClearConditionalFormats => self.clear_cond_formats_in_selection(cx),
            CommandId::SelectBlanks => self.select_blanks(cx),
            CommandId::SelectCurrentRegion => self.select_current_region(cx),
            CommandId::HideRows => self.hide_rows(cx),
            CommandId::UnhideRows => self.unhide_rows(cx),
            CommandId::HideCols => self.hide_cols(cx),
            CommandId::UnhideCols => self.unhide_cols(cx),

            // Editing
            CommandId::FillDown => self.fill_down(cx),
            CommandId::FillRight => self.fill_right(cx),
            CommandId::ClearCells => self.delete_selection(cx),
            CommandId::TrimWhitespace => self.trim_whitespace(cx),
            CommandId::TransformUppercase => self.apply_transform_pro(crate::transforms::TransformOp::Uppercase, cx),
            CommandId::TransformLowercase => self.apply_transform_pro(crate::transforms::TransformOp::Lowercase, cx),
            CommandId::TransformTitleCase => self.apply_transform_pro(crate::transforms::TransformOp::TitleCase, cx),
            CommandId::TransformSentenceCase => self.apply_transform_pro(crate::transforms::TransformOp::SentenceCase, cx),
            CommandId::Undo => self.undo(cx),
            CommandId::Redo => self.redo(cx),
            CommandId::AutoSum => self.autosum(cx),

            // Clipboard
            CommandId::Copy => self.copy(cx),
            CommandId::Cut => self.cut(cx),
            CommandId::Paste => self.paste(cx),
            CommandId::PasteValues => self.paste_values(cx),
            CommandId::TogglePasteValuesDefault => self.toggle_paste_values_default(cx),
            CommandId::PasteSpecial => self.show_paste_special(cx),
            CommandId::PasteFormulas => self.paste_formulas(cx),
            CommandId::PasteFormats => self.paste_formats(cx),

            // Formatting
            CommandId::ToggleBold => self.toggle_bold(cx),
            CommandId::RepeatLastAction => self.repeat_last_action(window, cx),
            CommandId::FitColumnWidth => self.fit_selection_columns(window, cx),
            CommandId::AlignLeft => self.set_alignment_selection(visigrid_engine::cell::Alignment::Left, cx),
            CommandId::AlignCenter => self.set_alignment_selection(visigrid_engine::cell::Alignment::Center, cx),
            CommandId::AlignRight => self.set_alignment_selection(visigrid_engine::cell::Alignment::Right, cx),
            CommandId::AlignGeneral => self.set_alignment_selection(visigrid_engine::cell::Alignment::General, cx),
            CommandId::CenterAcrossSelection => self.center_across_selection_toggle(cx),
            CommandId::ToggleItalic => self.toggle_italic(cx),
            CommandId::ToggleUnderline => self.toggle_underline(cx),
            CommandId::FormatCurrency => self.format_currency(cx),
            CommandId::FormatPercent => self.format_percent(cx),
            CommandId::FormatCells => {
                // Open inspector to format tab
                self.inspector_visible = true;
                self.inspector_tab = crate::mode::InspectorTab::Format;
                cx.notify();
            }

            CommandId::ClearFormatting => self.clear_formatting_selection(cx),
            CommandId::FormatPainter => self.start_format_painter(cx),
            CommandId::FormatPainterLocked => self.start_format_painter_locked(cx),
            CommandId::CopyFormat => self.copy_format(cx),
            CommandId::PasteFormat => self.paste_format(cx),

            // Background colors
            CommandId::FillColor => self.show_color_picker(crate::color_palette::ColorTarget::Fill, window, cx),
            CommandId::ClearBackground => self.set_background_color(None, cx),
            CommandId::BackgroundYellow => self.set_background_color(Some([255, 255, 0, 255]), cx),
            CommandId::BackgroundGreen => self.set_background_color(Some([198, 239, 206, 255]), cx),
            CommandId::BackgroundBlue => self.set_background_color(Some([189, 215, 238, 255]), cx),
            CommandId::BackgroundRed => self.set_background_color(Some([255, 199, 206, 255]), cx),
            CommandId::BackgroundOrange => self.set_background_color(Some([255, 235, 156, 255]), cx),
            CommandId::BackgroundPurple => self.set_background_color(Some([204, 192, 218, 255]), cx),
            CommandId::BackgroundGray => self.set_background_color(Some([217, 217, 217, 255]), cx),
            CommandId::BackgroundCyan => self.set_background_color(Some([183, 222, 232, 255]), cx),

            // Borders
            CommandId::BordersAll => self.apply_borders(BorderApplyMode::All, cx),
            CommandId::BordersOutline => self.apply_borders(BorderApplyMode::Outline, cx),
            CommandId::BordersInside => self.apply_borders(BorderApplyMode::Inside, cx),
            CommandId::BordersTop => self.apply_borders(BorderApplyMode::Top, cx),
            CommandId::BordersBottom => self.apply_borders(BorderApplyMode::Bottom, cx),
            CommandId::BordersLeft => self.apply_borders(BorderApplyMode::Left, cx),
            CommandId::BordersRight => self.apply_borders(BorderApplyMode::Right, cx),
            CommandId::BordersClear => self.apply_borders(BorderApplyMode::Clear, cx),

            // Cell styles
            CommandId::StyleDefault => self.set_cell_style_selection(CellStyle::None, cx),
            CommandId::StyleError => self.set_cell_style_selection(CellStyle::Error, cx),
            CommandId::StyleWarning => self.set_cell_style_selection(CellStyle::Warning, cx),
            CommandId::StyleSuccess => self.set_cell_style_selection(CellStyle::Success, cx),
            CommandId::StyleInput => self.set_cell_style_selection(CellStyle::Input, cx),
            CommandId::StyleTotal => self.set_cell_style_selection(CellStyle::Total, cx),
            CommandId::StyleNote => self.set_cell_style_selection(CellStyle::Note, cx),
            CommandId::StyleClear => self.set_cell_style_selection(CellStyle::None, cx),

            // File
            // NewWindow dispatches the action which propagates to App-level handler
            CommandId::NewWindow => cx.dispatch_action(&crate::actions::NewWindow),
            CommandId::OpenFile => self.open_file(cx),
            CommandId::Save => self.save(cx),
            CommandId::SaveAs => self.save_as(cx),
            CommandId::ExportCsv => self.export_csv(cx),
            CommandId::ExportTsv => self.export_tsv(cx),
            CommandId::ExportJson => self.export_json(cx),

            // Appearance
            CommandId::SelectTheme => self.show_theme_picker(cx),
            CommandId::SelectFont => self.show_font_picker(window, cx),

            // View
            CommandId::ToggleInspector => {
                self.inspector_visible = !self.inspector_visible;
                if self.inspector_visible { self.profiler_visible = false; }
                cx.notify();
            }
            CommandId::ToggleProfiler => {
                self.profiler_visible = !self.profiler_visible;
                if self.profiler_visible { self.inspector_visible = false; }
                cx.notify();
            }
            CommandId::ProfileNextRecalc => self.profile_next_recalc(cx),
            CommandId::ClearProfiler => {
                self.profiler_report = None;
                self.profiler_hotspots = Vec::new();
                cx.notify();
            }
            CommandId::ToggleMinimap => {
                self.minimap_visible = !self.minimap_visible;
                cx.notify();
            }
            CommandId::ToggleZenMode => {
                self.zen_mode = !self.zen_mode;
                cx.notify();
            }
            CommandId::ZoomIn => self.zoom_in(cx),
            CommandId::ZoomOut => self.zoom_out(cx),
            CommandId::ZoomReset => self.zoom_reset(cx),
            CommandId::FreezeTopRow => self.freeze_top_row(cx),
            CommandId::FreezeFirstColumn => self.freeze_first_column(cx),
            CommandId::FreezePanes => self.freeze_panes(cx),
            CommandId::UnfreezePanes => self.unfreeze_panes(cx),
            CommandId::SplitRight => self.split_right(cx),
            CommandId::CloseSplit => self.close_split(cx),
            CommandId::ToggleTrace => self.toggle_trace(cx),
            CommandId::CycleTracePrecedent => self.cycle_trace_precedent(false, cx),
            CommandId::CycleTraceDependent => self.cycle_trace_dependent(false, cx),
            CommandId::ReturnToTraceSource => self.return_to_trace_source(cx),
            CommandId::ToggleVerifiedMode => self.toggle_verified_mode(cx),
            CommandId::ToggleVimMode => self.toggle_vim_mode(cx),
            CommandId::Recalculate => self.recalculate(cx),
            CommandId::ReloadCustomFunctions => self.reload_custom_functions(cx),
            CommandId::ApproveModel => self.approve_model(None, cx),
            CommandId::ClearApproval => self.clear_approval(cx),
            CommandId::NavPerfReport => {
                let msg = self.nav_perf.report()
                    .unwrap_or_else(|| "Nav perf tracking disabled. Set VISIGRID_PERF=nav and restart.".into());
                self.status_message = Some(msg);
                cx.notify();
            }

            // Window - dispatch to App-level handler
            CommandId::SwitchWindow => cx.dispatch_action(&crate::actions::SwitchWindow),

            // Help
            CommandId::ShowShortcuts => {
                #[cfg(target_os = "macos")]
                {
                    self.status_message = Some("Shortcuts: Cmd+D Fill Down, Cmd+R Fill Right, Cmd+Enter Multi-edit, Cmd+` Switch Window".into());
                }
                #[cfg(not(target_os = "macos"))]
                {
                    self.status_message = Some("Shortcuts: Ctrl+D Fill Down, Ctrl+R Fill Right, Ctrl+Enter Multi-edit, Ctrl+` Switch Window".into());
                }
                cx.notify();
            }
            CommandId::OpenKeybindings => {
                self.open_keybindings(cx);
            }
            CommandId::OpenDocs => {
                let _ = open::that(crate::docs_links::DOCS_HOME);
            }
            CommandId::ShowAbout => {
                self.show_about(cx);
            }
            CommandId::TourNamedRanges => {
                self.show_tour(cx);
            }
            CommandId::ShowRefactorLog => {
                self.show_refactor_log(cx);
            }
            CommandId::ShowAISettings => {
                self.show_ai_settings(cx);
            }
            CommandId::InsertFormulaAI => {
                self.show_ask_ai(cx);
            }
            CommandId::AnalyzeAI => {
                self.show_analyze(cx);
            }
            CommandId::ExtractNamedRange => {
                self.show_extract_named_range(cx);
            }

            // Sheets
            CommandId::NextSheet => self.next_sheet(cx),
            CommandId::PrevSheet => self.prev_sheet(cx),
            CommandId::AddSheet => self.add_sheet(cx),

            // Data (sort/filter)
            CommandId::SortAscending => {
                self.sort_by_current_column(visigrid_engine::filter::SortDirection::Ascending, cx);
            }
            CommandId::SortDescending => {
                self.sort_by_current_column(visigrid_engine::filter::SortDirection::Descending, cx);
            }
            CommandId::ToggleAutoFilter => self.toggle_auto_filter(cx),
            CommandId::ClearSort => self.clear_sort(cx),

            // Data (validation)
            CommandId::ValidationDialog => self.show_validation_dialog(cx),
            CommandId::ExcludeFromValidation => self.exclude_from_validation(cx),
            CommandId::ClearValidationExclusions => self.clear_validation_exclusions(cx),
            CommandId::CircleInvalidData => self.circle_invalid_data(cx),
            CommandId::ClearInvalidCircles => self.clear_invalid_circles(cx),
            CommandId::OpenDiffResults => self.open_diff_results(window, cx),
            CommandId::RerunDiff => self.rerun_diff(window, cx),
            CommandId::RerunDiffRun => self.rerun_diff_run(window, cx),
            CommandId::OpenThisDiffFile => self.open_this_diff_file(cx),
            CommandId::RefreshDiffResults => self.refresh_diff_results(window, cx),

            // Hub sync
            CommandId::HubCheckStatus => self.hub_check_status(cx),
            CommandId::HubPull => self.hub_pull(cx),
            CommandId::HubPublish => self.hub_publish(cx),
            CommandId::HubOpenRemoteAsCopy => self.hub_open_remote_as_copy(cx),
            CommandId::HubUnlink => self.hub_unlink(cx),
            CommandId::HubDiagnostics => self.hub_diagnostics(cx),
            CommandId::HubSignIn => self.hub_sign_in(cx),
            CommandId::HubSignOut => self.hub_sign_out(cx),
            CommandId::HubLinkDialog => self.hub_show_link_dialog(cx),

            // Phase 5: Open Result in Grid
            CommandId::ImportTerminalOutput => self.import_terminal_output(window, cx),
            CommandId::RunVgridPeekJson => self.run_vgrid_peek_json(window, cx),
            CommandId::RunVgridDiffJson => self.run_vgrid_diff_json(window, cx),
            CommandId::RunVgridCalcJson => self.run_vgrid_calc_json(window, cx),

            // Phase 4: Palette-driven terminal
            CommandId::OpenTerminal => self.open_terminal(window, cx),
            CommandId::VerifyIntegrity => self.verify_integrity(cx),
            CommandId::Convert => self.show_convert_picker(window, cx),

            // Phase 6: AI Lua capture
            CommandId::CaptureAiLua => self.capture_ai_lua(window, cx),
            CommandId::PreviewLastLua => self.preview_last_lua(window, cx),
            CommandId::BuildModelWithLua => self.build_model_with_lua(window, cx),

            // AI in Terminal
            CommandId::LaunchAI => self.launch_ai_terminal(window, cx),
            CommandId::PasteSelectionToTerminal => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::PasteContext);
                self.paste_selection_context(window, cx);
            }
            CommandId::PasteHeadersToTerminal => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::PasteContext);
                self.paste_headers_context(window, cx);
            }
            CommandId::PasteFilePathToTerminal => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::PasteContext);
                self.paste_file_path_context(window, cx);
            }
            CommandId::PasteVisiGridContext => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::PasteContext);
                self.paste_full_context(window, cx);
            }
            CommandId::GenerateAiContextFiles => {
                crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::GenerateContextFiles);
                self.generate_ai_context_files(window, cx);
            }
            CommandId::OpenAiContextFolder => self.open_ai_context_folder(window, cx),
            CommandId::AiExplainSelection => self.ai_explain_selection(window, cx),
            CommandId::OpenAiMetrics => self.open_ai_metrics(window, cx),
        }

        // Ensure title reflects any state changes from this command.
        // The flag + cache debounce makes this cheap for non-state-changing commands.
        self.request_title_refresh(cx);
    }

    // =========================================================================
    // AI in Terminal
    // =========================================================================


    /// Write text to the terminal PTY, optionally wrapped in bracketed paste mode.
    /// Bracketed paste prevents shells from executing newlines as commands.
    /// Controlled by the `terminal.bracketed_paste` user setting (default ON).
    pub(crate) fn write_to_pty_bracketed(&mut self, text: &str, cx: &Context<Self>) {
        let use_bracketed = crate::settings::user_settings(cx)
            .terminal
            .bracketed_paste
            .as_value()
            .copied()
            .unwrap_or(true);

        if use_bracketed {
            self.terminal.write_to_pty(b"\x1b[200~");
        }
        self.terminal.write_to_pty(text.as_bytes());
        if use_bracketed {
            self.terminal.write_to_pty(b"\x1b[201~");
        }
    }

    /// Max rows/cols to paste into terminal to avoid UI freeze.
    pub(crate) const PASTE_MAX_ROWS: usize = 200;
    pub(crate) const PASTE_MAX_COLS: usize = 50;
    pub(crate) const PASTE_MAX_CHARS: usize = 50_000;













    // ========================================================================
    // Phase 5: Open Result in Grid
    // ========================================================================

    /// Extract structured result from terminal output (auto-detect or manual).
    pub fn try_extract_structured_result(&mut self, cx: &mut Context<Self>) {
        let Some(ref term_arc) = self.terminal.term else {
            self.terminal.watching_for_result = false;
            self.status_message = Some("No terminal session.".into());
            cx.notify();
            return;
        };

        let text = crate::terminal::extract::extract_recent_text(term_arc, 2000);
        match crate::structured_results::parse(&text) {
            Some(result) => {
                let had_previous = self.terminal.pending_result.is_some();
                self.terminal.pending_result = Some(
                    crate::terminal::state::PendingResult::Structured(result)
                );
                self.terminal.watching_for_result = false;
                if had_previous {
                    self.status_message = Some(
                        "New structured result detected (previous replaced).".into()
                    );
                }
                // Auto-open if setting is enabled
                let auto_open = crate::settings::user_settings(cx)
                    .terminal.auto_open_structured_results
                    .as_value()
                    .copied()
                    .unwrap_or(false);
                if auto_open {
                    self.open_structured_result_inner(cx);
                } else {
                    cx.notify();
                }
            }
            None => {
                self.terminal.watching_for_result = false;
                self.status_message = Some("No structured output detected.".into());
                cx.notify();
            }
        }
    }

    /// Open the pending structured result as a new sheet.
    pub fn open_structured_result(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        self.open_structured_result_inner(cx);
    }

    /// Inner method that opens the pending result without requiring window access.
    pub(crate) fn open_structured_result_inner(&mut self, cx: &mut Context<Self>) {
        let result = match self.terminal.pending_result.take() {
            Some(crate::terminal::state::PendingResult::Structured(r)) => r,
            other => {
                // Put it back if it's not a Structured result
                self.terminal.pending_result = other;
                return;
            }
        };
        let source_file = self.current_file.as_ref().map(|p| p.display().to_string());
        let command = self.terminal.last_injected_command.take();
        let meta = self.workbook.update(cx, |wb, _| {
            let meta = crate::structured_results::render_to_sheet(&result, wb);
            crate::structured_results::append_run_log(wb, &meta, source_file.as_deref(), command.as_deref());
            meta
        });
        self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(meta.sheet_idx); });
        self.status_message = Some(format!("Opened {} as new sheet.", meta.sheet_name));
        cx.notify();
    }

    /// Dismiss the pending structured result.
    pub fn dismiss_structured_result(&mut self, cx: &mut Context<Self>) {
        self.terminal.pending_result = None;
        cx.notify();
    }


    /// Inject a vgrid command into the terminal and watch for JSON result.
    fn inject_vgrid_command(&mut self, subcmd: &str, window: &mut Window, cx: &mut Context<Self>) {
        let Some(ref file) = self.current_file else {
            self.status_message = Some("Save workbook first.".into());
            cx.notify();
            return;
        };

        let file_str = file.display().to_string();

        // Ensure terminal is open
        self.open_terminal(window, cx);

        // Build command
        let cmd = format!("vgrid {} \"{}\" --json\n", subcmd, file_str);
        self.terminal.write_to_pty(cmd.as_bytes());

        // Set watching
        self.terminal.watching_for_result = true;
        self.terminal.watch_generation += 1;
        self.terminal.pending_result = None;
        self.terminal.result_settle_task = None;
        self.terminal.last_injected_command = Some(cmd.trim().to_string());
    }

    /// Palette command: Run vgrid peek --json on current file.
    pub fn run_vgrid_peek_json(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.inject_vgrid_command("peek --headers", window, cx);
    }


    /// Palette command: Run vgrid calc --json. Injects a template command for the user to edit.
    pub fn run_vgrid_calc_json(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let Some(ref file) = self.current_file else {
            self.status_message = Some("Save workbook first.".into());
            cx.notify();
            return;
        };
        let file_str = file.display().to_string();

        // Ensure terminal is open
        self.open_terminal(window, cx);

        // Inject a template — user needs to fill in the formula
        let cmd = format!("cat \"{}\" | vgrid calc '=FORMULA' -f csv --headers --json", file_str);
        self.terminal.write_to_pty(cmd.as_bytes());
        // Don't write \n — let the user edit the formula first

        // Set watching for when they press Enter
        self.terminal.watching_for_result = true;
        self.terminal.watch_generation += 1;
        self.terminal.pending_result = None;
        self.terminal.result_settle_task = None;
        self.terminal.last_injected_command = Some(cmd.clone());

        self.status_message = Some("Edit the formula in the command, then press Enter.".into());
        cx.notify();
    }


    // ========================================================================
    // Phase 6: AI → Lua → Preview → Apply
    // ========================================================================







    /// Verify integrity inline — checks semantic fingerprint and shows status message.
    pub fn verify_integrity(&mut self, cx: &mut Context<Self>) {
        if self.current_file.is_none() && self.document_meta.display_name.starts_with("Book") {
            self.status_message = Some("No active file to verify.".to_string());
            cx.notify();
            return;
        }

        let status = self.verification_status(cx);
        let sv = &self.semantic_verification;

        let msg = match status {
            VerificationStatus::Unverified => {
                "No verification baseline. Use 'Approve Model' first.".to_string()
            }
            VerificationStatus::Verified => {
                let mut m = "\u{2713} Verification passed".to_string();
                if let Some(label) = &sv.label {
                    m.push_str(&format!(" \u{2014} matches Approved: \"{}\"", label));
                }
                if let Some(ts) = &sv.timestamp {
                    m.push_str(&format!(" ({})", ts));
                }
                m
            }
            VerificationStatus::Drifted => {
                let mut m = "\u{2717} Verification failed \u{2014} drifted".to_string();
                if let Some(label) = &sv.label {
                    m.push_str(&format!(" from Approved: \"{}\"", label));
                }
                if let Some(ts) = &sv.timestamp {
                    m.push_str(&format!(" ({})", ts));
                }
                m
            }
        };

        self.status_message = Some(msg);
        cx.notify();
    }

    /// Show the convert format picker dialog.
    pub fn show_convert_picker(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        // Validate: need an active file
        if self.current_file.is_none() {
            self.status_message = Some("No active workbook to convert.".to_string());
            cx.notify();
            return;
        }
        // Validate: must be saved (not dirty)
        if self.history.is_dirty() {
            self.status_message = Some("Save the file before converting.".to_string());
            cx.notify();
            return;
        }

        self.convert_picker_selected = 0;
        self.mode = Mode::ConvertPicker;
        cx.notify();
    }

    /// Cancel the convert picker dialog.
    pub fn cancel_convert_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Execute conversion: insert command into terminal.
    pub fn execute_convert(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let format_idx = self.convert_picker_selected;
        self.mode = Mode::Navigation;

        let (ext, format_flag) = match format_idx {
            0 => ("csv", "csv"),
            1 => ("tsv", "tsv"),
            2 => ("json", "json"),
            3 => ("xlsx", "xlsx"),
            _ => ("csv", "csv"),
        };

        let file_path = match &self.current_file {
            Some(p) => p.clone(),
            None => {
                self.status_message = Some("No active workbook to convert.".to_string());
                cx.notify();
                return;
            }
        };

        // Build input path: workspace-relative if possible, else quoted absolute
        let input_str = if let Some(ws) = &self.terminal.workspace_root {
            if let Ok(rel) = file_path.strip_prefix(ws) {
                rel.display().to_string()
            } else {
                format!("\"{}\"", file_path.display())
            }
        } else {
            format!("\"{}\"", file_path.display())
        };

        // Build output filename
        let basename = file_path.file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("output");
        let output_name = format!("{}.{}", basename, ext);

        let command = format!(
            "vgrid convert {} -t {} -o {}  # Output: ./{}",
            input_str, format_flag, output_name, output_name
        );

        // Ensure terminal is open
        self.bottom_panel_visible = true;
        self.bottom_panel_tab = BottomPanelTab::Terminal;
        self.lua_console.visible = false;
        self.terminal.visible = true;
        if self.terminal.term.is_none() && !self.terminal.exited {
            self.spawn_terminal(window, cx);
        } else {
            self.terminal.ensure_cwd();
        }

        // Clear line (Ctrl+U) then insert command (no auto-run)
        self.terminal.write_to_pty(b"\x15");
        self.terminal.write_to_pty(command.as_bytes());

        self.terminal_focused = true;
        window.focus(&self.terminal_focus_handle, cx);
        self.status_message = Some(format!("Convert command inserted \u{2014} press Enter to run."));
        cx.notify();
    }







    // Menu methods
    pub fn toggle_menu(&mut self, menu: crate::mode::Menu, cx: &mut Context<Self>) {
        if self.open_menu == Some(menu) {
            self.open_menu = None;
        } else {
            self.open_menu = Some(menu);
        }
        self.menu_highlight = None;
        cx.notify();
    }

    pub fn close_menu(&mut self, cx: &mut Context<Self>) {
        if self.open_menu.is_some() {
            self.open_menu = None;
            self.menu_highlight = None;
            cx.notify();
        }
    }

    /// Close the Format dropdown menu in the header bar.
    /// Called by: backdrop click, Escape key, mode switches (Find/GoTo), opening other popovers.
    pub fn close_format_menu(&mut self, cx: &mut Context<Self>) {
        if self.ui.format_menu_open {
            self.ui.format_menu_open = false;
            cx.notify();
        }
    }

    /// Open the Format dropdown menu in the header bar.
    pub fn open_format_menu(&mut self, cx: &mut Context<Self>) {
        self.ui.format_menu_open = true;
        cx.notify();
    }

    /// Toggle the Format dropdown menu in the header bar.
    pub fn toggle_format_menu(&mut self, cx: &mut Context<Self>) {
        self.ui.format_menu_open = !self.ui.format_menu_open;
        self.ui.format_bar.cell_style_menu_open = false;
        cx.notify();
    }

    // Menu keyboard navigation methods

    pub fn menu_highlight_next(&mut self, cx: &mut Context<Self>) {
        if let Some(menu) = self.open_menu {
            let count = crate::menu_model::menu_item_count(menu);
            if count == 0 { return; }
            self.menu_highlight = Some(match self.menu_highlight {
                None => 0,
                Some(i) => if i + 1 >= count { 0 } else { i + 1 },
            });
            cx.notify();
        }
    }

    pub fn menu_highlight_prev(&mut self, cx: &mut Context<Self>) {
        if let Some(menu) = self.open_menu {
            let count = crate::menu_model::menu_item_count(menu);
            if count == 0 { return; }
            self.menu_highlight = Some(match self.menu_highlight {
                None => count - 1,
                Some(0) => count - 1,
                Some(i) => i - 1,
            });
            cx.notify();
        }
    }

    pub fn menu_switch_next(&mut self, cx: &mut Context<Self>) {
        if let Some(current) = self.open_menu {
            self.open_menu = Some(Self::next_active_menu(current));
            self.menu_highlight = None;
            cx.notify();
        }
    }

    pub fn menu_switch_prev(&mut self, cx: &mut Context<Self>) {
        if let Some(current) = self.open_menu {
            self.open_menu = Some(Self::prev_active_menu(current));
            self.menu_highlight = None;
            cx.notify();
        }
    }

    pub fn menu_execute_highlighted(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if let (Some(menu), Some(index)) = (self.open_menu, self.menu_highlight) {
            self.close_menu(cx);
            crate::menu_model::execute_menu_action(self, menu, index, window, cx);
        }
    }

    /// Try to execute a menu item by its accelerator letter.
    /// Returns true if a matching item was found and executed.
    pub fn menu_execute_by_letter(&mut self, letter: char, window: &mut Window, cx: &mut Context<Self>) -> bool {
        use crate::menu_model::{MenuEntry, menu_entries, execute_menu_action, resolve_accel};

        if let Some(menu) = self.open_menu {
            let entries = menu_entries(menu);
            let mut selectable_idx = 0;
            for entry in &entries {
                match entry {
                    MenuEntry::Item { label, accel, .. } | MenuEntry::Color { label, accel, .. } => {
                        let item_letter = resolve_accel(label, *accel);
                        if item_letter == letter {
                            self.close_menu(cx);
                            execute_menu_action(self, menu, selectable_idx, window, cx);
                            return true;
                        }
                        selectable_idx += 1;
                    }
                    _ => {}
                }
            }
        }
        false
    }

    fn next_active_menu(start: crate::mode::Menu) -> crate::mode::Menu {
        let mut m = start.next();
        for _ in 0..7 {
            if crate::menu_model::menu_item_count(m) > 0 { return m; }
            m = m.next();
        }
        start
    }

    fn prev_active_menu(start: crate::mode::Menu) -> crate::mode::Menu {
        let mut m = start.prev();
        for _ in 0..7 {
            if crate::menu_model::menu_item_count(m) > 0 { return m; }
            m = m.prev();
        }
        start
    }

    /// Get width for a column (custom or default) for the current sheet
    pub fn col_width(&self, col: usize) -> f32 {
        self.col_widths
            .get(&self.cached_sheet_id)
            .and_then(|sheet_widths| sheet_widths.get(&col))
            .copied()
            .unwrap_or(CELL_WIDTH)
    }

    /// Get height for a row (custom or default) for the current sheet
    pub fn row_height(&self, row: usize) -> f32 {
        self.row_heights
            .get(&self.cached_sheet_id)
            .and_then(|sheet_heights| sheet_heights.get(&row))
            .copied()
            .unwrap_or(CELL_HEIGHT)
    }

    /// Set column width for the current sheet
    pub fn set_col_width(&mut self, col: usize, width: f32) {
        // Deliberately not `clamp`. For a NaN input `max(20.0)` yields
        // 20.0, because f32::max prefers the non-NaN operand, while
        // `clamp` propagates NaN — and the test below sends NaN down the
        // `insert` arm rather than `remove`, writing NaN into the size map.
        // Sizes are persisted with the session, so it would outlive a restart.
        #[allow(clippy::manual_clamp)]
        let width = width.max(20.0).min(500.0); // 20-500px
        let sheet_widths = self.col_widths.entry(self.cached_sheet_id).or_insert_with(HashMap::new);
        if (width - CELL_WIDTH).abs() < 1.0 {
            sheet_widths.remove(&col); // Remove if close to default
        } else {
            sheet_widths.insert(col, width);
        }
    }

    /// Set row height for the current sheet
    pub fn set_row_height(&mut self, row: usize, height: f32) {
        // Deliberately not `clamp`. For a NaN input `max(12.0)` yields
        // 12.0, because f32::max prefers the non-NaN operand, while
        // `clamp` propagates NaN — and the test below sends NaN down the
        // `insert` arm rather than `remove`, writing NaN into the size map.
        // Sizes are persisted with the session, so it would outlive a restart.
        #[allow(clippy::manual_clamp)]
        let height = height.max(12.0).min(200.0); // 12-200px
        let sheet_heights = self.row_heights.entry(self.cached_sheet_id).or_insert_with(HashMap::new);
        if (height - CELL_HEIGHT).abs() < 1.0 {
            sheet_heights.remove(&row); // Remove if close to default
        } else {
            sheet_heights.insert(row, height);
        }
    }

    /// Record a column width change to history (for undo/redo support).
    /// Called on mouse up after a resize drag to coalesce all drag events into one history entry.
    pub fn record_col_width_change(&mut self, col: usize, old: Option<f32>, cx: &mut Context<Self>) {
        // Get the current value from the map
        let new = self.col_widths
            .get(&self.cached_sheet_id)
            .and_then(|m| m.get(&col))
            .copied();

        // Only record if something actually changed
        if old != new {
            // Use SheetId (stable across sheet reorder/delete) instead of index
            let sheet_id = self.cached_sheet_id;
            self.history.record_action_with_provenance(
                crate::history::UndoAction::ColumnWidthSet {
                    sheet_id,
                    col,
                    old,
                    new,
                },
                None,
            );
            self.is_modified = true;
        }
    }

    /// Record a row height change to history (for undo/redo support).
    /// Called on mouse up after a resize drag to coalesce all drag events into one history entry.
    pub fn record_row_height_change(&mut self, row: usize, old: Option<f32>, cx: &mut Context<Self>) {
        // Get the current value from the map
        let new = self.row_heights
            .get(&self.cached_sheet_id)
            .and_then(|m| m.get(&row))
            .copied();

        // Only record if something actually changed
        if old != new {
            // Use SheetId (stable across sheet reorder/delete) instead of index
            let sheet_id = self.cached_sheet_id;
            self.history.record_action_with_provenance(
                crate::history::UndoAction::RowHeightSet {
                    sheet_id,
                    row,
                    old,
                    new,
                },
                None,
            );
            self.is_modified = true;
        }
    }

    /// Get mutable reference to column widths map for the current sheet
    /// Creates the map if it doesn't exist
    pub fn sheet_col_widths_mut(&mut self) -> &mut HashMap<usize, f32> {
        self.col_widths.entry(self.cached_sheet_id).or_insert_with(HashMap::new)
    }

    /// Get mutable reference to row heights map for the current sheet
    /// Creates the map if it doesn't exist
    pub fn sheet_row_heights_mut(&mut self) -> &mut HashMap<usize, f32> {
        self.row_heights.entry(self.cached_sheet_id).or_insert_with(HashMap::new)
    }

    /// Check if current sheet has any custom row heights
    pub fn has_custom_row_heights(&self) -> bool {
        self.row_heights.get(&self.cached_sheet_id).map_or(false, |h| !h.is_empty())
    }

    /// Check if a row is hidden on the current sheet
    pub fn is_row_hidden(&self, row: usize) -> bool {
        self.hidden_rows
            .get(&self.cached_sheet_id)
            .map_or(false, |set| set.contains(&row))
    }

    /// Check if a column is hidden on the current sheet
    pub fn is_col_hidden(&self, col: usize) -> bool {
        self.hidden_cols
            .get(&self.cached_sheet_id)
            .map_or(false, |set| set.contains(&col))
    }

    /// Check if current sheet has any hidden rows
    pub fn has_hidden_rows(&self) -> bool {
        self.hidden_rows.get(&self.cached_sheet_id).map_or(false, |s| !s.is_empty())
    }

    /// Check if current sheet has any hidden columns
    pub fn has_hidden_cols(&self) -> bool {
        self.hidden_cols.get(&self.cached_sheet_id).map_or(false, |s| !s.is_empty())
    }

    /// Get the nth visible column starting from scroll_col, skipping hidden columns.
    /// Returns the actual column index, or None if out of bounds.
    pub fn nth_visible_col(&self, visible_index: usize, scroll_col: usize) -> Option<usize> {
        if !self.has_hidden_cols() {
            let col = scroll_col + visible_index;
            return if col < NUM_COLS { Some(col) } else { None };
        }
        let hidden = self.hidden_cols.get(&self.cached_sheet_id).unwrap();
        let mut count = 0;
        let mut col = scroll_col;
        while col < NUM_COLS {
            if !hidden.contains(&col) {
                if count == visible_index {
                    return Some(col);
                }
                count += 1;
            }
            col += 1;
        }
        None
    }

    /// Get the nth visible row composing RowView filtering with user-hidden rows.
    /// Returns (view_row, data_row) or None if out of bounds.
    pub fn nth_visible_row_with_hidden(&self, visible_index: usize, cx: &gpui::App) -> Option<(usize, usize)> {
        if !self.has_hidden_rows() {
            return self.nth_visible_row(visible_index, cx);
        }
        let hidden = self.hidden_rows.get(&self.cached_sheet_id).unwrap();
        let mut count = 0;
        let mut idx = 0;
        loop {
            let (view_row, data_row) = self.nth_visible_row(idx, cx)?;
            if !hidden.contains(&data_row) {
                if count == visible_index {
                    return Some((view_row, data_row));
                }
                count += 1;
            }
            idx += 1;
        }
    }

    /// Update cached sheet ID from the workbook.
    /// Call this after switching sheets.
    pub fn update_cached_sheet_id(&mut self, cx: &mut Context<Self>) {
        self.cached_sheet_id = self.workbook.read(cx).active_sheet().id;
    }

    /// Get the cached sheet ID (for use in views without context access)
    pub fn cached_sheet_id(&self) -> SheetId {
        self.cached_sheet_id
    }

    /// Debug assertion: verify cached_sheet_id matches the workbook's active sheet.
    /// Call this in hot paths (render, selection change) to catch desync bugs early.
    /// Only runs in debug builds.
    #[inline]
    pub fn debug_assert_sheet_cache_sync(&self, cx: &Context<Self>) {
        #[cfg(debug_assertions)]
        {
            let actual_id = self.workbook.read(cx).active_sheet().id;
            debug_assert_eq!(
                self.cached_sheet_id, actual_id,
                "cached_sheet_id desync! cached={:?}, actual={:?}. \
                 A code path changed the active sheet without calling update_cached_sheet_id().",
                self.cached_sheet_id, actual_id
            );
        }
    }

    /// Get mutable reference to column widths map for a specific sheet by index
    /// Used by undo/redo operations that need to access a specific sheet's sizing.
    /// Creates the map if it doesn't exist.
    pub fn sheet_col_widths_for_index_mut(&mut self, sheet_index: usize, cx: &mut Context<Self>) -> &mut HashMap<usize, f32> {
        let sheet_id = self.workbook.read(cx).sheets().get(sheet_index)
            .map(|s| s.id)
            .unwrap_or(self.cached_sheet_id);
        self.col_widths.entry(sheet_id).or_insert_with(HashMap::new)
    }

    /// Get mutable reference to row heights map for a specific sheet by index
    /// Used by undo/redo operations that need to access a specific sheet's sizing.
    /// Creates the map if it doesn't exist.
    pub fn sheet_row_heights_for_index_mut(&mut self, sheet_index: usize, cx: &mut Context<Self>) -> &mut HashMap<usize, f32> {
        let sheet_id = self.workbook.read(cx).sheets().get(sheet_index)
            .map(|s| s.id)
            .unwrap_or(self.cached_sheet_id);
        self.row_heights.entry(sheet_id).or_insert_with(HashMap::new)
    }

    /// Get the X position of a column's left edge (relative to start of grid, after row header)
    /// Returns scaled (zoomed) position for rendering.
    pub fn col_x_offset(&self, target_col: usize) -> f32 {
        let mut x = 0.0;
        for col in self.view_state.scroll_col..target_col {
            x += self.metrics.col_width(self.col_width(col));
        }
        GridMetrics::snap_floor(x, self.metrics.scale)
    }

    /// Get the Y position of a row's top edge (relative to start of grid, after column header)
    /// Returns scaled (zoomed) position for rendering.
    pub fn row_y_offset(&self, target_row: usize) -> f32 {
        let mut y = 0.0;
        for row in self.view_state.scroll_row..target_row {
            y += self.metrics.row_height(self.row_height(row));
        }
        GridMetrics::snap_floor(y, self.metrics.scale)
    }

    /// Get the bounding rect of a cell in grid-relative coordinates.
    /// This is the single source of truth for cell position within the grid viewport.
    /// Used for positioning popups, overlays, and other elements relative to cells.
    pub fn cell_rect(&self, row: usize, col: usize) -> CellRect {
        CellRect {
            x: self.col_x_offset(col),
            y: self.row_y_offset(row),
            width: self.metrics.col_width(self.col_width(col)),
            height: self.metrics.row_height(self.row_height(row)),
        }
    }

    /// Get the bounding rect of the currently selected (active) cell in grid-relative coordinates.
    pub fn active_cell_rect(&self) -> CellRect {
        let (row, col) = self.view_state.selected;
        self.cell_rect(row, col)
    }

    /// Get the viewport rect for the grid body (for clamp/flip calculations).
    /// Returns (width, height) of the visible grid area.
    pub fn viewport_rect(&self) -> (f32, f32) {
        self.grid_layout.viewport_size
    }

    /// Convert window X position to column index.
    /// Uses measured grid_layout.grid_body_origin for accuracy.
    /// Uses scaled (zoomed) column widths for hit-testing.
    pub fn col_from_window_x(&self, window_x: f32) -> Option<usize> {
        let x = window_x - self.grid_layout.grid_body_origin.0;
        if x < 0.0 { return None; }

        let viewport_width = self.grid_layout.viewport_size.0;
        let mut current_x = 0.0;
        for col in self.view_state.scroll_col..NUM_COLS {
            if current_x > viewport_width { break; }
            // Use scaled width for hit-testing in screen coordinates
            let width = self.metrics.col_width(self.col_width(col));
            if x < current_x + width {
                return Some(col);
            }
            current_x += width;
        }
        Some(NUM_COLS - 1)  // Clamp to last column if beyond viewport
    }

    /// Convert window Y position to row index.
    /// O(1) for uniform heights, O(visible rows) for variable heights.
    /// Uses scaled (zoomed) row heights for hit-testing.
    pub fn row_from_window_y(&self, window_y: f32) -> Option<usize> {
        let y = window_y - self.grid_layout.grid_body_origin.1;
        if y < 0.0 { return None; }

        // O(1) fast path: uniform row heights (use scaled cell height)
        if !self.has_custom_row_heights() {
            let row = self.view_state.scroll_row + (y / self.metrics.cell_h).floor() as usize;
            return Some(row.min(NUM_ROWS - 1));
        }

        // O(visible rows) slow path: variable heights, stop at viewport bottom
        let viewport_height = self.grid_layout.viewport_size.1;
        let mut current_y = 0.0;
        let mut last_row = self.view_state.scroll_row;
        for row in self.view_state.scroll_row..NUM_ROWS {
            if current_y > viewport_height { break; }
            last_row = row;
            // Use scaled height for hit-testing in screen coordinates
            let height = self.metrics.row_height(self.row_height(row));
            if y < current_y + height {
                return Some(row);
            }
            current_y += height;
        }
        Some(last_row)
    }

    /// Auto-fit column width to content
    /// Measure the width a set of columns needs to show their content.
    ///
    /// One pass over the sheet's POPULATED cells (not 65,536 rows per column),
    /// measuring the same text the grid renders. With a `window` the text is
    /// really shaped by the font system, honouring size and bold; without one
    /// (the import-time path has no window) it falls back to an estimate over
    /// CHARACTERS — the old code multiplied `str::len()`, which is bytes, so
    /// "café" measured as 5 and CJK as 3× its true width.
    fn measure_columns(
        &self,
        cols: &[usize],
        window: Option<&Window>,
        cx: &App,
    ) -> HashMap<usize, f32> {
        use std::collections::HashSet;

        let wanted: HashSet<usize> = cols.iter().copied().collect();
        let mut widths: HashMap<usize, f32> = cols.iter().map(|c| (*c, MIN_AUTOFIT_WIDTH)).collect();
        let sheet = self.sheet(cx);

        let coords: Vec<(usize, usize)> = sheet
            .cells_iter()
            .map(|(&rc, _)| rc)
            .filter(|(_, c)| wanted.contains(c))
            .collect();

        for (row, col) in coords {
            let text = sheet.get_formatted_display(row, col);
            if text.is_empty() {
                continue;
            }
            let format = sheet.get_format(row, col);
            let font_size = format.font_size.unwrap_or(self.metrics.font_size);
            let width = match window {
                Some(w) => {
                    let shared: SharedString = text.into();
                    let len = shared.len();
                    let font = Font {
                        weight: if format.bold { FontWeight::BOLD } else { FontWeight::NORMAL },
                        style: if format.italic { FontStyle::Italic } else { FontStyle::Normal },
                        ..Font::default()
                    };
                    let shaped = w.text_system().shape_line(
                        shared,
                        px(font_size),
                        &[TextRun {
                            len,
                            font,
                            color: Hsla::default(),
                            background_color: None,
                            underline: None,
                            strikethrough: None,
                        }],
                        None,
                    );
                    let w: f32 = shaped.width.into();
                    w
                }
                None => estimate_text_width(&text, font_size, format.bold),
            };
            let entry = widths.entry(col).or_insert(MIN_AUTOFIT_WIDTH);
            *entry = entry.max(width + AUTOFIT_PADDING);
        }

        for w in widths.values_mut() {
            *w = w.clamp(MIN_AUTOFIT_WIDTH, MAX_AUTOFIT_WIDTH);
        }
        widths
    }

    /// Apply measured widths to `cols`, recording one undo entry for the lot.
    fn fit_columns(&mut self, cols: Vec<usize>, window: Option<&Window>, cx: &mut Context<Self>) {
        if cols.is_empty() {
            return;
        }
        let sheet_id = self.cached_sheet_id;
        let old_widths: Vec<(usize, Option<f32>)> = cols
            .iter()
            .map(|&col| (col, self.col_widths.get(&sheet_id).and_then(|m| m.get(&col)).copied()))
            .collect();

        let measured = self.measure_columns(&cols, window, cx);
        for (&col, &width) in measured.iter() {
            self.set_col_width(col, width);
        }

        let mut actions = Vec::new();
        for (col, old) in old_widths {
            let new = self.col_widths.get(&sheet_id).and_then(|m| m.get(&col)).copied();
            if old != new {
                actions.push(crate::history::UndoAction::ColumnWidthSet { sheet_id, col, old, new });
            }
        }
        if !actions.is_empty() {
            let count = actions.len();
            if count == 1 {
                self.history.record_action_with_provenance(actions.remove(0), None);
            } else {
                self.history.record_action_with_provenance(
                    crate::history::UndoAction::Group {
                        actions,
                        description: "Auto-fit column widths".to_string(),
                    },
                    None,
                );
            }
            self.is_modified = true;
            self.status_message = Some(if count == 1 {
                "Fit 1 column to its content".to_string()
            } else {
                format!("Fit {} columns to their content", count)
            });
        } else {
            self.status_message = Some("Columns already fit their content".to_string());
        }
        cx.notify();
    }

    /// Fit every column touched by the current selection (or the active
    /// cell's column when nothing wider is selected). The palette and
    /// keyboard entry point — the header double-click uses
    /// `auto_fit_selected_col_widths`.
    pub fn fit_selection_columns(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.set_repeat(crate::repeat::RepeatAction::FitColumnWidth);
        let mut cols: Vec<usize> = Vec::new();
        for ((_, min_col), (_, max_col)) in self.all_selection_ranges() {
            for col in min_col..=max_col {
                if !cols.contains(&col) {
                    cols.push(col);
                }
            }
        }
        if cols.is_empty() {
            cols.push(self.view_state.active_cell().1);
        }
        // A whole-sheet selection (Ctrl+A) would mean 256 columns; fit only
        // those that actually hold something.
        if cols.len() > 1 {
            let populated: std::collections::HashSet<usize> =
                self.sheet(cx).cells_iter().map(|(&(_, c), _)| c).collect();
            cols.retain(|c| populated.contains(c));
        }
        self.fit_columns(cols, Some(window), cx);
    }

    pub fn auto_fit_col_width(&mut self, col: usize, window: Option<&Window>, cx: &mut Context<Self>) {
        self.fit_columns(vec![col], window, cx);
    }

    /// Auto-fit column width - if the clicked column is part of the selection,
    /// auto-fit all selected columns (Excel behavior)
    pub fn auto_fit_selected_col_widths(
        &mut self,
        clicked_col: usize,
        window: Option<&Window>,
        cx: &mut Context<Self>,
    ) {
        let cols = if self.is_col_header_selected(clicked_col) {
            let mut cols = Vec::new();
            for ((_, min_col), (_, max_col)) in self.all_selection_ranges() {
                for col in min_col..=max_col {
                    if !cols.contains(&col) {
                        cols.push(col);
                    }
                }
            }
            cols
        } else {
            vec![clicked_col]
        };
        self.fit_columns(cols, window, cx);
    }

    /// Auto-fit all columns that have content (for agent-built sheets).
    /// Runs at import time, where no window is available — see
    /// `measure_columns` for what that costs.
    pub fn auto_fit_all_data_columns(&mut self, cx: &App) {
        let cols: Vec<usize> = {
            let mut seen: Vec<usize> = self.sheet(cx).cells_iter().map(|(&(_, c), _)| c).collect();
            seen.sort_unstable();
            seen.dedup();
            seen
        };
        if cols.is_empty() {
            return;
        }
        let measured = self.measure_columns(&cols, None, cx);
        for (&col, &width) in measured.iter() {
            self.set_col_width(col, width);
        }
    }

    /// Auto-fit row height. Rows reset to the default height: VisiGrid has no
    /// multi-line cell text yet, so there is nothing taller to fit to.
    pub fn auto_fit_row_height(&mut self, row: usize, cx: &mut Context<Self>) {
        self.sheet_row_heights_mut().remove(&row);
        cx.notify();
    }

    /// Auto-fit row height - if row is selected and multiple rows are selected,
    /// auto-fit all selected rows (Excel behavior)
    pub fn auto_fit_selected_row_heights(&mut self, clicked_row: usize, cx: &mut Context<Self>) {
        // Check if clicked row is part of selection
        if self.is_row_header_selected(clicked_row) {
            // Collect all selected rows
            let mut rows_to_fit = Vec::new();
            for ((min_row, _), (max_row, _)) in self.all_selection_ranges() {
                for row in min_row..=max_row {
                    if !rows_to_fit.contains(&row) {
                        rows_to_fit.push(row);
                    }
                }
            }
            // Auto-fit each selected row
            for row in rows_to_fit {
                self.auto_fit_row_height_no_notify(row);
            }
            cx.notify();
        } else {
            // Not part of selection, just auto-fit the clicked row
            self.auto_fit_row_height(clicked_row, cx);
        }
    }

    /// Auto-fit row height without notifying (for batch operations)
    fn auto_fit_row_height_no_notify(&mut self, row: usize) {
        // For now, just reset to default since we don't support multi-line
        self.sheet_row_heights_mut().remove(&row);
    }

    /// Check if edit_value starts with = or + (formula indicator)
    pub fn is_formula_content(&self) -> bool {
        self.edit_value.starts_with('=') || self.edit_value.starts_with('+')
    }

    /// Check if grid navigation should be blocked (modal is open).
    /// Use this in action handlers to prevent keyboard leaks to the grid.
    ///
    /// IMPORTANT: When adding a new modal mode:
    /// 1. Add it to Mode::is_overlay() in mode.rs
    /// 2. This method will then correctly block grid navigation
    /// 3. For text-input modals, also add cursor handling in MoveLeft/MoveRight handlers
    #[inline]
    pub fn should_block_grid_navigation(&self) -> bool {
        self.mode.is_overlay()
    }

    // =========================================================================
    // KeyTips (macOS Option double-tap accelerators)
    // =========================================================================

    /// Toggle KeyTips overlay (Option+Space on macOS).
    /// Shows keyboard accelerator hints for menu navigation.
    #[cfg(target_os = "macos")]
    pub fn toggle_keytips(&mut self, cx: &mut Context<Self>) {
        // Don't show KeyTips if text input is active
        if !self.should_handle_option_accelerators() {
            return;
        }

        if self.keytips_active {
            self.dismiss_keytips(cx);
            return;
        }

        // Show KeyTips overlay
        self.keytips_active = true;
        let now = std::time::Instant::now();
        self.keytips_deadline_at = Some(now + std::time::Duration::from_secs(3));
        cx.notify();

        // Schedule auto-dismiss after 3 seconds
        cx.spawn(async move |this, cx| {
            cx.background_executor().timer(std::time::Duration::from_secs(3)).await;
            let _ = this.update(cx, |this, cx| {
                if this.keytips_active {
                    this.dismiss_keytips(cx);
                }
            });
        }).detach();
    }

    /// Stub for non-macOS (KeyTips is macOS-only)
    #[cfg(not(target_os = "macos"))]
    pub fn toggle_keytips(&mut self, _cx: &mut Context<Self>) {}

    /// Dismiss KeyTips overlay
    pub fn dismiss_keytips(&mut self, cx: &mut Context<Self>) {
        if self.keytips_active {
            self.keytips_active = false;
            self.keytips_deadline_at = None;
            cx.notify();
        }
    }

    /// Handle key press while KeyTips is active.
    /// Returns true if the key was handled (caller should stop propagation).
    pub fn keytips_handle_key(&mut self, key: &str, cx: &mut Context<Self>) -> bool {
        if !self.keytips_active {
            return false;
        }

        // Map key to menu category
        // STABLE MAPPING: These letters are locked and will not change.
        // Users build muscle memory; changing mappings breaks trust.
        let category = match key.to_lowercase().as_str() {
            "f" => Some(crate::search::MenuCategory::File),
            "e" => Some(crate::search::MenuCategory::Edit),
            "v" => Some(crate::search::MenuCategory::View),
            "o" => Some(crate::search::MenuCategory::Format),  // O for fOrmat (F taken by File)
            "d" => Some(crate::search::MenuCategory::Data),
            "t" => Some(crate::search::MenuCategory::Tools),
            "h" => Some(crate::search::MenuCategory::Help),
            // Enter or Space: repeat last scope (power-user speed)
            "enter" | "space" => {
                self.keytips_active = false;
                self.keytips_deadline_at = None;
                if let Some(scope) = self.last_keytips_scope {
                    self.apply_menu_scope(scope, cx);
                } else {
                    // No previous scope - just dismiss
                    cx.notify();
                }
                return true;
            }
            "escape" => {
                self.dismiss_keytips(cx);
                return true;
            }
            _ => {
                // Unknown key - dismiss (snappy, avoids stuck overlay)
                self.dismiss_keytips(cx);
                return true;
            }
        };

        // Dismiss and open scoped palette
        self.keytips_active = false;
        self.keytips_deadline_at = None;

        if let Some(cat) = category {
            // Store for repeat-last-scope
            self.last_keytips_scope = Some(cat);
            self.apply_menu_scope(cat, cx);
        }

        true
    }

    /// Check if Option+letter accelerators should be handled.
    /// Returns false if any text input is active, preventing conflicts
    /// with macOS character composition (accents, special characters).
    ///
    /// This is the central guard for all Option-based accelerators on macOS.
    /// When this returns false, Option+letter events should pass through
    /// to the OS for normal text input handling.
    #[inline]
    pub fn should_handle_option_accelerators(&self) -> bool {
        // Block if mode has text input
        if self.mode.has_text_input() {
            return false;
        }

        // Block if Lua console is visible (has text input)
        if self.lua_console.visible {
            return false;
        }

        // Block if filter dropdown search is active
        if self.filter_dropdown_col.is_some() && !self.filter_search_text.is_empty() {
            return false;
        }

        // Block if sheet rename is active
        if self.renaming_sheet.is_some() {
            return false;
        }

        // Block if validation dropdown is open (may have text)
        if self.is_validation_dropdown_open() {
            return false;
        }

        // Safe to handle Option accelerators
        true
    }




    /// Total height of all UI chrome above the grid body.
    ///
    /// Must match the actual rendered layout in views/mod.rs (top to bottom):
    ///   macOS titlebar (MACOS_TITLEBAR_HEIGHT, macOS only)
    ///   Menu bar       (MENU_BAR_HEIGHT, Linux only, hidden in zen mode)
    ///   Formula bar    (FORMULA_BAR_HEIGHT, hidden in zen mode)
    ///   Format bar     (FORMAT_BAR_HEIGHT, hidden in zen mode or when disabled)
    ///   Column headers (metrics.header_h, always visible, scales with zoom)
    ///
    /// This is the single source of truth for grid_body_origin.y and visible_rows().
    pub fn top_chrome_height(&self, cx: &App) -> f32 {
        if self.zen_mode {
            // Zen hides menu, formula bar, format bar — only column headers remain
            return self.metrics.header_h;
        }
        let titlebar_h = if cfg!(target_os = "macos") { MACOS_TITLEBAR_HEIGHT } else { 0.0 };
        let menu_h = if cfg!(target_os = "macos") { 0.0 } else { MENU_BAR_HEIGHT };
        let formula_h = FORMULA_BAR_HEIGHT;
        let format_h = {
            use crate::settings::Setting;
            match &user_settings(cx).appearance.show_format_bar {
                Setting::Value(v) => if *v { crate::views::format_bar::FORMAT_BAR_HEIGHT } else { 0.0 },
                Setting::Inherit => crate::views::format_bar::FORMAT_BAR_HEIGHT,
            }
        };
        titlebar_h + menu_h + formula_h + format_h + self.metrics.header_h
    }

    /// Total height of UI chrome below the grid body.
    /// Currently just the status bar, which is hidden in zen mode.
    pub fn bottom_chrome_height(&self) -> f32 {
        if self.zen_mode { 0.0 } else { STATUS_BAR_HEIGHT }
    }

    /// Calculate visible rows based on window height.
    /// Uses cached grid_layout.viewport_size computed each render by top/bottom_chrome_height.
    pub fn visible_rows(&self) -> usize {
        let available = self.grid_layout.viewport_size.1;
        if available <= 0.0 {
            return 1;
        }
        let rows = (available / self.metrics.cell_h).floor() as usize;
        rows.clamp(1, NUM_ROWS)
    }

    /// Calculate visible columns based on window width and actual column widths.
    /// Sums real column widths starting from the current scroll position to determine
    /// how many columns fit in the viewport. Adds 1 extra column for partial visibility.
    pub fn visible_cols(&self) -> usize {
        let width: f32 = self.window_size.width.into();
        let available_width = width - self.metrics.header_w;
        let scroll_col = self.view_state.scroll_col;
        let frozen_cols = self.view_state.frozen_cols;

        // Account for frozen columns first (they consume space before scrollable area)
        let mut used = 0.0_f32;
        for fc in 0..frozen_cols {
            if !self.is_col_hidden(fc) {
                used += self.metrics.col_width(self.col_width(fc));
            }
        }

        // Sum actual column widths from scroll position until we exceed available width
        // Count represents the number of visible (non-hidden) columns
        let mut count = frozen_cols;
        let mut col = scroll_col;
        while used < available_width && col < NUM_COLS {
            if !self.is_col_hidden(col) {
                used += self.metrics.col_width(self.col_width(col));
                count += 1;
            }
            col += 1;
        }

        // Add 1 extra for partially visible columns at the edge
        count = count.saturating_add(1);
        count.clamp(1, NUM_COLS)
    }

    /// Update window size (called on resize)
    pub fn update_window_size(&mut self, size: Size<Pixels>, cx: &mut Context<Self>) {
        self.window_size = size;
        cx.notify();
    }

    // Column letter (A, B, ..., Z, AA, AB, ...)
    pub fn col_letter(col: usize) -> String {
        let mut result = String::new();
        let mut c = col;
        loop {
            result.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 { break; }
            c = c / 26 - 1;
        }
        result
    }

    // Cell reference (A1, B2, etc.)
    pub fn cell_ref(&self) -> String {
        format!("{}{}", Self::col_letter(self.view_state.selected.1), self.view_state.selected.0 + 1)
    }


    // Formatting (applies to all discontiguous selection ranges)
    pub fn toggle_bold(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_bold(row, col));
                }
            }
        }
        // A toggle over a mixed selection has no single "new value", so the
        // repeat slot takes the ACTIVE cell's resolved state — that is the
        // one the user was looking at when they pressed the key.
        let (r, c) = self.view_state.active_cell();
        let resolved = self.sheet(cx).get_format(r, c).bold;
        self.set_repeat(RepeatAction::Bold(resolved));
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_italic(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_italic(row, col));
                }
            }
        }
        let (r, c) = self.view_state.active_cell();
        let resolved = self.sheet(cx).get_format(r, c).italic;
        self.set_repeat(RepeatAction::Italic(resolved));
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_underline(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_underline(row, col));
                }
            }
        }
        let (r, c) = self.view_state.active_cell();
        let resolved = self.sheet(cx).get_format(r, c).underline;
        self.set_repeat(RepeatAction::Underline(resolved));
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_strikethrough(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_strikethrough(row, col));
                }
            }
        }
        let (r, c) = self.view_state.active_cell();
        let resolved = self.sheet(cx).get_format(r, c).strikethrough;
        self.set_repeat(RepeatAction::Strikethrough(resolved));
        self.is_modified = true;
        cx.notify();
    }

    pub fn format_currency(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, NumberFormat::currency(2)));
                }
            }
        }
        self.set_repeat(RepeatAction::NumberFormat(NumberFormat::currency(2)));
        self.is_modified = true;
        cx.notify();
    }

    pub fn format_percent(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, NumberFormat::Percent { decimals: 2 }));
                }
            }
        }
        self.set_repeat(RepeatAction::NumberFormat(NumberFormat::Percent { decimals: 2 }));
        self.is_modified = true;
        cx.notify();
    }

    pub fn format_date_shortcut(&mut self, cx: &mut Context<Self>) {
        self.set_number_format_selection(NumberFormat::Date { style: visigrid_engine::cell::DateStyle::Short }, cx);
    }

    pub fn format_number_shortcut(&mut self, cx: &mut Context<Self>) {
        self.set_number_format_selection(NumberFormat::Number { decimals: 2, thousands: true, negative: visigrid_engine::cell::NegativeStyle::Minus }, cx);
    }

    pub fn format_general_shortcut(&mut self, cx: &mut Context<Self>) {
        self.set_number_format_selection(NumberFormat::General, cx);
    }

    pub fn format_scientific_shortcut(&mut self, cx: &mut Context<Self>) {
        self.set_number_format_selection(NumberFormat::Custom("0.00E+00".to_string()), cx);
    }

    pub fn format_time_shortcut(&mut self, cx: &mut Context<Self>) {
        self.set_number_format_selection(NumberFormat::Time, cx);
    }

    /// Insert current date as text (Ctrl+;)
    pub fn insert_date(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        let now = chrono::Local::now();
        let date_str = now.format("%-m/%-d/%Y").to_string();
        let (row, col) = self.view_state.active_cell();
        let old_value = self.sheet(cx).get_raw(row, col);
        self.set_cell_value(row, col, &date_str, cx);
        self.history.record_change(self.sheet_index(cx), row, col, old_value, date_str);
        self.is_modified = true;
        self.status_message = Some("Date inserted".to_string());
        cx.notify();
    }

    /// Insert current time as text (Ctrl+Shift+;)
    pub fn insert_time(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) { return; }
        let now = chrono::Local::now();
        let time_str = now.format("%-I:%M %p").to_string();
        let (row, col) = self.view_state.active_cell();
        let old_value = self.sheet(cx).get_raw(row, col);
        self.set_cell_value(row, col, &time_str, cx);
        self.history.record_change(self.sheet_index(cx), row, col, old_value, time_str);
        self.is_modified = true;
        self.status_message = Some("Time inserted".to_string());
        cx.notify();
    }



    /// Extend selection to cell A1 (Ctrl+Shift+Home)
    pub fn extend_to_start(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        self.view_state.selection_end = Some((0, 0));
        self.view_state.scroll_row = 0;
        self.view_state.scroll_col = 0;
        cx.notify();
    }

    /// Extend selection to last cell (Ctrl+Shift+End)
    pub fn extend_to_end(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        self.view_state.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        self.ensure_visible(cx);
    }

    /// Select current region (Ctrl+Shift+*) — selects contiguous data block around active cell
    pub fn select_current_region(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        let (row, col) = self.view_state.selected;
        let (min_row, min_col, max_row, max_col) =
            crate::ai::find_current_region(self.sheet(cx), row, col);
        // Anchor at top-left, extend to bottom-right
        self.view_state.selected = (min_row, min_col);
        self.view_state.selection_end = if (min_row, min_col) == (max_row, max_col) {
            None  // Single cell — no range needed
        } else {
            Some((max_row, max_col))
        };
        self.ensure_visible(cx);
        let rows = max_row - min_row + 1;
        let cols = max_col - min_col + 1;
        self.status_message = Some(format!("Selected region: {}×{}", rows, cols));
        cx.notify();
    }

    /// Insert a newline character into the edit buffer (Alt+Enter)
    pub fn insert_newline(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // Delete selection if any
            self.delete_edit_selection();
            let byte_idx = self.edit_cursor.min(self.edit_value.len());
            self.edit_value.insert(byte_idx, '\n');
            self.edit_cursor = byte_idx + 1;
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;
            cx.notify();
        } else {
            // Start editing then insert newline
            let (row, col) = self.view_state.selected;
            self.edit_original = self.sheet(cx).get_raw(row, col);
            self.edit_value = self.edit_original.clone();
            self.edit_cursor = self.edit_value.len();
            self.mode = Mode::Edit;
            // Now insert the newline
            let byte_idx = self.edit_cursor.min(self.edit_value.len());
            self.edit_value.insert(byte_idx, '\n');
            self.edit_cursor = byte_idx + 1;
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;
            cx.notify();
        }
    }

    /// Open context menu at active cell (Shift+F10)
    pub fn open_context_menu(&mut self, cx: &mut Context<Self>) {
        // Calculate the active cell's position in window coordinates
        let rect = self.active_cell_rect();
        let origin = self.grid_layout.grid_body_origin;
        let position = gpui::Point {
            x: gpui::px(origin.0 + rect.x + rect.width * 0.5),
            y: gpui::px(origin.1 + rect.y + rect.height),
        };
        self.show_context_menu(ContextMenuKind::Cell, position, cx);
    }

    // =========================================================================
    // F2 Function Key Tip (macOS only)
    // =========================================================================

    /// Check if F2 tip should be shown (macOS only, not dismissed, tip was triggered)
    #[cfg(target_os = "macos")]
    pub fn should_show_f2_tip(&self, cx: &gpui::App) -> bool {
        self.show_f2_tip && !user_settings(cx).is_tip_dismissed(TipId::F2Edit)
    }

    #[cfg(not(target_os = "macos"))]
    pub fn should_show_f2_tip(&self, _cx: &gpui::App) -> bool {
        false
    }

    /// Called when user edits via non-F2 path on macOS (double-click, Ctrl+U, menu)
    /// Shows tip suggesting they enable standard function keys
    #[cfg(target_os = "macos")]
    pub fn maybe_show_f2_tip(&mut self, cx: &mut Context<Self>) {
        if !user_settings(cx).is_tip_dismissed(TipId::F2Edit) {
            self.show_f2_tip = true;
            cx.notify();
        }
    }

    #[cfg(not(target_os = "macos"))]
    pub fn maybe_show_f2_tip(&mut self, _cx: &mut Context<Self>) {
        // No-op on non-macOS
    }

    /// Dismiss the F2 tip permanently
    pub fn dismiss_f2_tip(&mut self, cx: &mut Context<Self>) {
        update_user_settings(cx, |settings| {
            settings.dismiss_tip(TipId::F2Edit);
        });
        self.show_f2_tip = false;
        cx.notify();
    }

    /// Hide F2 tip without permanently dismissing
    pub fn hide_f2_tip(&mut self, cx: &mut Context<Self>) {
        self.show_f2_tip = false;
        cx.notify();
    }


    /// Reset all tips (for Preferences UI)
    pub fn reset_all_tips(&mut self, cx: &mut Context<Self>) {
        update_user_settings(cx, |settings| {
            settings.reset_all_tips();
        });
        cx.notify();
    }

    // =========================================================================
    // Rewind Preview (Phase 8A)
    // =========================================================================



    /// Get the workbook to display - preview snapshot if previewing, else live workbook.
    /// Requires context to access the Entity<Workbook> - pass &**cx from Context.
    pub fn display_workbook<'a>(&'a self, cx: &'a App) -> &'a Workbook {
        match &self.rewind_preview {
            RewindPreviewState::On(session) => &session.snapshot,
            RewindPreviewState::Off => self.wb(cx),
        }
    }

    /// Check if editing is allowed (blocked during preview)
    pub fn can_edit(&self) -> bool {
        !self.is_previewing()
    }


    /// Block a bulk operation when the active sheet contains merged cells.
    /// Returns true (and sets status message) if merges exist, false otherwise.
    /// `op_name` is a user-facing verb phrase like "sort", "fill", "replace".
    pub fn block_if_merged(&mut self, op_name: &str, cx: &mut Context<Self>) -> bool {
        if !self.sheet(cx).merged_regions.is_empty() {
            self.status_message = Some(format!(
                "Cannot {op_name}: this operation can't be applied to merged cells. Unmerge first."
            ));
            cx.notify();
            true
        } else {
            false
        }
    }
}













impl Render for Spreadsheet {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        // Drain pending session server requests (TCP → GUI bridge)
        self.drain_session_requests(cx);

        // Flush batched navigation moves (multiple arrow repeats → one batch per frame)
        self.flush_pending_nav_moves(cx);
        // Flush deferred scroll adjustment (coalesces multiple nav moves per frame)
        self.flush_nav_scroll();
        // Record render timestamp for latency instrumentation
        self.nav_perf.mark_render();

        // Cold start measurement (fires once on first render)
        if self.cold_start_ms.is_none() {
            if let Some(start) = self.startup_instant {
                let ms = start.elapsed().as_millis();
                self.cold_start_ms = Some(ms);
                // Show for empty launches (no file loaded yet).
                //
                // Alongside whatever else has something to say, rather than
                // instead of it: this used to require status_message to be
                // empty, so any restore message won and the number vanished.
                // A fresh profile showed it and a machine in daily use never
                // did — which is the wrong way round for a measurement.
                if self.current_file.is_none() {
                    self.status_message = Some(match self.status_message.take() {
                        Some(existing) => format!("{} · Ready in {}ms", existing, ms),
                        None => format!("Ready in {}ms", ms),
                    });
                }
            }
        }

        // One-shot title refresh (triggered by async operations without window access)
        if self.pending_title_refresh {
            self.pending_title_refresh = false;
            self.update_title_if_needed(window, cx);
        }

        // Update window size if changed (handles resize)
        let current_size = window.viewport_size();
        if self.window_size != current_size {
            self.window_size = current_size;
            // Re-validate edit scroll on resize (caret may now be offscreen)
            if self.mode.is_editing() {
                self.edit_scroll_dirty = true;
                self.update_edit_scroll(window);
            }
        }

        // Update grid metrics if display scale factor changed (e.g. window moved to Retina display)
        let sf = window.scale_factor();
        if (sf - self.metrics.scale).abs() > 0.001 {
            self.metrics = GridMetrics::with_scale(self.metrics.zoom, sf);
        }

        // Debug: report border instrumentation (once per second, only when debug overlay is on).
        // All counters accumulate across frames; reset only on print.
        // 0 borders_calls = fast path active (sheet has no borders).
        #[cfg(debug_assertions)]
        if self.debug_grid_alignment {
            self.debug_border_frames.set(self.debug_border_frames.get() + 1);
            let now = std::time::Instant::now();
            let last = self.debug_border_last_report.get();
            if now.duration_since(last).as_secs() >= 1 {
                let calls = self.debug_border_call_count.get();
                let gridlines = self.debug_gridline_cells.get();
                let overlays = self.debug_userborder_cells.get();
                let frames = self.debug_border_frames.get();
                let has_flag = self.sheet(cx).has_any_borders;
                eprintln!(
                    "[border-debug] borders_calls={} gridline_cells={} userborder_cells={} frames={} has_any_borders={}",
                    calls, gridlines, overlays, frames, has_flag,
                );
                // Stale flag tripwire: has_any_borders=true but nothing actually drawn.
                // 3 consecutive 1-second windows triggers a loud warning.
                if has_flag && overlays == 0 && calls > 0 {
                    self.debug_border_stale_streak += 1;
                    if self.debug_border_stale_streak >= 3 {
                        eprintln!(
                            "[border-debug][WARN] has_any_borders=true but no user borders drawn \
                             for {}s; likely stale flag. Consider scan_border_flag().",
                            self.debug_border_stale_streak,
                        );
                    }
                } else {
                    self.debug_border_stale_streak = 0;
                }
                self.debug_border_call_count.set(0);
                self.debug_gridline_cells.set(0);
                self.debug_userborder_cells.set(0);
                self.debug_border_frames.set(0);
                self.debug_border_last_report.set(now);
            }
        }

        // Cache window bounds for session snapshot (updated each render)
        self.cached_window_bounds = Some(window.window_bounds());

        // Modal focus guard: when an overlay modal is open, grid navigation should be blocked.
        // Note: bottom_panel_visible no longer blocks grid nav — per-action terminal_has_focus()
        // guards in actions_nav.rs and actions_edit.rs handle terminal focus routing.

        // Update grid layout cache for hit-testing.
        // top_chrome_height / bottom_chrome_height are the single source of truth;
        // visible_rows() reads the cached viewport_size set here.
        let grid_body_y = self.top_chrome_height(cx);
        let grid_body_x = self.metrics.header_w;

        let window_height: f32 = current_size.height.into();
        let window_width: f32 = current_size.width.into();

        let right_panel_width = if self.inspector_visible || self.profiler_visible {
            crate::views::inspector_panel::PANEL_WIDTH
        } else {
            0.0
        };
        let bottom_status_height = self.bottom_chrome_height();

        let grid_viewport_width = (window_width - grid_body_x - right_panel_width).max(0.0);
        let grid_viewport_height = (window_height - grid_body_y - bottom_status_height).max(0.0);

        self.grid_layout = GridLayout {
            grid_body_origin: (grid_body_x, grid_body_y),
            viewport_size: (grid_viewport_width, grid_viewport_height),
        };

        // Update formula bar text rect for click-to-place-caret hit-testing
        // Uses centralized constants: FORMULA_BAR_TEXT_LEFT, FORMULA_BAR_PADDING
        let formula_bar_input_left = FORMULA_BAR_CELL_REF_WIDTH + FORMULA_BAR_FX_WIDTH;
        let formula_bar_text_width = (window_width - formula_bar_input_left - FORMULA_BAR_PADDING * 2.0 - right_panel_width).max(0.0);
        // Formula bar sits directly below the menu bar (Linux) or titlebar (macOS)
        let formula_bar_y = if cfg!(target_os = "macos") {
            MACOS_TITLEBAR_HEIGHT
        } else if self.zen_mode {
            0.0
        } else {
            MENU_BAR_HEIGHT
        };
        self.formula_bar_text_rect = gpui::Bounds {
            origin: gpui::point(gpui::px(FORMULA_BAR_TEXT_LEFT), gpui::px(formula_bar_y)),
            size: gpui::size(gpui::px(formula_bar_text_width), gpui::px(FORMULA_BAR_HEIGHT)),
        };

        // Update formula bar display cache (only when not editing)
        // This avoids re-parsing on every render
        if !self.mode.is_editing() {
            let cell = self.view_state.selected;
            let formula = self.sheet(cx).get_raw(cell.0, cell.1);

            // Only update cache if cell or formula changed
            let cache_valid = self.formula_bar_cache_cell == Some(cell)
                && self.formula_bar_cache_formula == formula;

            if !cache_valid {
                self.formula_bar_cache_cell = Some(cell);
                self.formula_bar_cache_formula = formula.clone();
                self.formula_bar_cache_refs = if formula.starts_with('=') || formula.starts_with('+') {
                    Self::parse_formula_refs(&formula)
                } else {
                    Vec::new()
                };
            }
        }

        views::render_spreadsheet(self, window, cx)
    }
}

// ============================================================================
// Phase 6: Helpers for AI Lua capture/preview/apply
// ============================================================================

/// Extract the last ` ```lua ` ... ` ``` ` fenced code block from text.
/// Case-insensitive on the language tag. Unclosed blocks are skipped.
pub(crate) fn extract_last_lua_block(text: &str) -> Option<String> {
    let mut last_block: Option<String> = None;
    let mut in_block = false;
    let mut current_lines: Vec<&str> = Vec::new();

    for line in text.lines() {
        let trimmed = line.trim();
        if !in_block {
            // Check for opening fence: ```lua (case-insensitive on language tag)
            if trimmed.starts_with("```") {
                let after = trimmed[3..].trim();
                if after.eq_ignore_ascii_case("lua") {
                    in_block = true;
                    current_lines.clear();
                }
            }
        } else {
            // Check for closing fence
            if trimmed == "```" {
                if !current_lines.is_empty() {
                    last_block = Some(current_lines.join("\n"));
                }
                in_block = false;
                current_lines.clear();
            } else {
                current_lines.push(line);
            }
        }
    }

    last_block
}

/// Save an AI-generated Lua script to `.visigrid/ai/generated/` and copy as `last.lua`.
pub(crate) fn save_ai_lua_script(code: &str, workspace_root: &Option<std::path::PathBuf>) -> std::path::PathBuf {
    let base = workspace_root.clone().unwrap_or_else(|| {
        dirs::home_dir()
            .unwrap_or_else(|| std::path::PathBuf::from("/tmp"))
            .join(".visigrid")
    });
    let dir = base.join("ai").join("generated");
    let _ = std::fs::create_dir_all(&dir);

    let epoch = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let filename = format!("model_{}.lua", epoch);
    let path = dir.join(&filename);
    let _ = std::fs::write(&path, code);

    // Also write last.lua for easy re-run
    let last_path = dir.join("last.lua");
    let _ = std::fs::write(&last_path, code);

    path
}

/// Fingerprint of sheet state for drift detection.
///
/// Samples deterministically across the full used range:
/// - Sheet name + cell count + used range bounds
/// - First 20 populated cells (head)
/// - Last 20 populated cells (tail)
/// - 88 evenly spaced cells across the middle (128 total samples)
/// - Hashes raw values (value OR formula), not display strings
pub(crate) fn sheet_fingerprint(sheet: &visigrid_engine::sheet::Sheet) -> u64 {
    use std::hash::{Hash, Hasher};
    let mut hasher = std::collections::hash_map::DefaultHasher::new();

    // Sheet identity + dimensions
    sheet.name.hash(&mut hasher);

    // Collect cell positions for sampling — cells_iter yields populated cells only
    let cells: Vec<(&(usize, usize), &visigrid_engine::cell::Cell)> =
        sheet.cells_iter().collect();
    let count = cells.len();
    count.hash(&mut hasher);

    if count == 0 {
        return hasher.finish();
    }

    // Used range bounds (min/max row/col across all populated cells)
    let (mut min_r, mut min_c) = (usize::MAX, usize::MAX);
    let (mut max_r, mut max_c) = (0usize, 0usize);
    for (&(r, c), _) in &cells {
        min_r = min_r.min(r);
        min_c = min_c.min(c);
        max_r = max_r.max(r);
        max_c = max_c.max(c);
    }
    min_r.hash(&mut hasher);
    min_c.hash(&mut hasher);
    max_r.hash(&mut hasher);
    max_c.hash(&mut hasher);

    const HEAD: usize = 20;
    const TAIL: usize = 20;
    const TOTAL_SAMPLES: usize = 128;
    let middle_budget = TOTAL_SAMPLES.saturating_sub(HEAD).saturating_sub(TAIL);

    // Hash a cell's raw content (value + formula, not display)
    let hash_cell = |h: &mut std::collections::hash_map::DefaultHasher,
                     &(r, c): &(usize, usize),
                     cell: &visigrid_engine::cell::Cell| {
        r.hash(h);
        c.hash(h);
        cell.value.raw_display().hash(h);
    };

    if count <= TOTAL_SAMPLES {
        // Small sheet: hash everything
        for (&pos, cell) in &cells {
            hash_cell(&mut hasher, &pos, cell);
        }
    } else {
        // Head
        for i in 0..HEAD {
            let (&pos, cell) = cells[i];
            hash_cell(&mut hasher, &pos, cell);
        }
        // Evenly spaced middle
        let middle_start = HEAD;
        let middle_end = count - TAIL;
        let middle_len = middle_end - middle_start;
        for i in 0..middle_budget {
            let idx = middle_start + (i * middle_len) / middle_budget;
            let (&pos, cell) = cells[idx];
            hash_cell(&mut hasher, &pos, cell);
        }
        // Tail
        for i in (count - TAIL)..count {
            let (&pos, cell) = cells[i];
            hash_cell(&mut hasher, &pos, cell);
        }
    }

    hasher.finish()
}

/// Count how many cells in ops would overwrite non-empty cells.
pub(crate) fn count_lua_overwrites(ops: &[crate::scripting::LuaOp], sheet: &visigrid_engine::sheet::Sheet) -> usize {
    use crate::scripting::LuaOp;
    let mut seen = std::collections::HashSet::new();
    let mut count = 0;
    for op in ops {
        let (row, col) = match op {
            LuaOp::SetValue { row, col, .. } => (*row as usize, *col as usize),
            LuaOp::SetFormula { row, col, .. } => (*row as usize, *col as usize),
            _ => continue,
        };
        if seen.insert((row, col)) && !sheet.get_raw(row, col).is_empty() {
            count += 1;
        }
    }
    count
}

#[cfg(test)]
mod paste_values_tests {
    use super::Spreadsheet;
    use visigrid_engine::formula::eval::Value;

    // =========================================================================
    // PASTE VALUES: External value parsing (leading-zero guard, booleans, etc.)
    // =========================================================================

    #[test]
    fn test_parse_external_value_leading_zero_preserved() {
        // Leading zeros should be preserved as text
        assert!(matches!(Spreadsheet::parse_external_value("007"), Value::Text(s) if s == "007"));
        assert!(matches!(Spreadsheet::parse_external_value("00123"), Value::Text(s) if s == "00123"));
        assert!(matches!(Spreadsheet::parse_external_value("000"), Value::Text(s) if s == "000"));
    }

    #[test]
    fn test_parse_external_value_single_zero_is_number() {
        // Single zero is a number, not text
        assert!(matches!(Spreadsheet::parse_external_value("0"), Value::Number(n) if n == 0.0));
    }

    #[test]
    fn test_parse_external_value_zero_decimal_is_number() {
        // 0.5, 0.123 are numbers (the second char is '.')
        assert!(matches!(Spreadsheet::parse_external_value("0.5"), Value::Number(n) if (n - 0.5).abs() < 0.001));
        assert!(matches!(Spreadsheet::parse_external_value("0.123"), Value::Number(n) if (n - 0.123).abs() < 0.001));
    }

    #[test]
    fn test_parse_external_value_boolean() {
        // TRUE/FALSE (case insensitive) become booleans
        assert!(matches!(Spreadsheet::parse_external_value("TRUE"), Value::Boolean(true)));
        assert!(matches!(Spreadsheet::parse_external_value("FALSE"), Value::Boolean(false)));
        assert!(matches!(Spreadsheet::parse_external_value("true"), Value::Boolean(true)));
        assert!(matches!(Spreadsheet::parse_external_value("false"), Value::Boolean(false)));
        assert!(matches!(Spreadsheet::parse_external_value("True"), Value::Boolean(true)));
    }

    #[test]
    fn test_parse_external_value_number() {
        // Regular numbers
        assert!(matches!(Spreadsheet::parse_external_value("42"), Value::Number(n) if n == 42.0));
        assert!(matches!(Spreadsheet::parse_external_value("-3.14"), Value::Number(n) if (n - (-3.14)).abs() < 0.001));
        assert!(matches!(Spreadsheet::parse_external_value("1e6"), Value::Number(n) if n == 1_000_000.0));
    }

    #[test]
    fn test_parse_external_value_text() {
        // Regular text
        assert!(matches!(Spreadsheet::parse_external_value("hello"), Value::Text(s) if s == "hello"));
        assert!(matches!(Spreadsheet::parse_external_value("ABC"), Value::Text(s) if s == "ABC"));
    }

    #[test]
    fn test_parse_external_value_empty() {
        assert!(matches!(Spreadsheet::parse_external_value(""), Value::Empty));
        assert!(matches!(Spreadsheet::parse_external_value("   "), Value::Empty));
    }

    #[test]
    fn test_parse_external_value_formula_prefix_becomes_text() {
        // Formula prefix is preserved as literal text (not executed)
        assert!(matches!(Spreadsheet::parse_external_value("=SUM(A1:A10)"), Value::Text(s) if s == "=SUM(A1:A10)"));
        assert!(matches!(Spreadsheet::parse_external_value("=A1+B1"), Value::Text(s) if s == "=A1+B1"));
    }

    #[test]
    fn test_parse_external_value_whitespace_trimmed() {
        // Whitespace should be trimmed
        assert!(matches!(Spreadsheet::parse_external_value("  42  "), Value::Number(n) if n == 42.0));
        assert!(matches!(Spreadsheet::parse_external_value("  hello  "), Value::Text(s) if s == "hello"));
    }

    // =========================================================================
    // PASTE VALUES: Canonical string representation
    // =========================================================================

    #[test]
    fn test_value_to_canonical_string_number() {
        // Integers should not have decimal places
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(42.0)), "42");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(0.0)), "0");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(-100.0)), "-100");

        // Decimals preserved
        let result = Spreadsheet::value_to_canonical_string(&Value::Number(3.14159));
        assert!(result.starts_with("3.14"));
    }

    #[test]
    fn test_value_to_canonical_string_boolean() {
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Boolean(true)), "TRUE");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Boolean(false)), "FALSE");
    }

    #[test]
    fn test_value_to_canonical_string_text() {
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Text("hello".to_string())), "hello");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Text("007".to_string())), "007");
    }

    #[test]
    fn test_value_to_canonical_string_error() {
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Error("#VALUE!".to_string())), "#VALUE!");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Error("#REF!".to_string())), "#REF!");
    }

    #[test]
    fn test_value_to_canonical_string_empty() {
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Empty), "");
    }

    // =========================================================================
    // CORRECTNESS: Exponent avoidance (never emit scientific notation)
    // =========================================================================

    #[test]
    fn test_value_to_canonical_string_no_scientific_notation_large() {
        // Large numbers must be full decimal, not scientific
        assert_eq!(
            Spreadsheet::value_to_canonical_string(&Value::Number(1e15)),
            "1000000000000000"
        );
        assert_eq!(
            Spreadsheet::value_to_canonical_string(&Value::Number(1234567890123456.0)),
            "1234567890123456"
        );
    }

    #[test]
    fn test_value_to_canonical_string_no_scientific_notation_small() {
        // Small decimals must be full decimal, not scientific
        let result = Spreadsheet::value_to_canonical_string(&Value::Number(0.000001));
        assert_eq!(result, "0.000001");
        assert!(!result.contains('e') && !result.contains('E'), "must not contain exponent");

        let result2 = Spreadsheet::value_to_canonical_string(&Value::Number(1e-6));
        assert_eq!(result2, "0.000001");
    }

    #[test]
    fn test_value_to_canonical_string_negative_zero_normalized() {
        // -0.0 must become "0", not "-0"
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(-0.0)), "0");
    }

    #[test]
    fn test_value_to_canonical_string_special_values() {
        // NaN and Infinity get explicit string representations
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(f64::NAN)), "NaN");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(f64::INFINITY)), "INF");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(f64::NEG_INFINITY)), "-INF");
    }

    #[test]
    fn test_value_to_canonical_string_trailing_zeros_trimmed() {
        // Trailing zeros after decimal should be trimmed
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(12.5)), "12.5");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(12.500)), "12.5");
        assert_eq!(Spreadsheet::value_to_canonical_string(&Value::Number(1.0)), "1");
    }

    // =========================================================================
    // CORRECTNESS: Clipboard metadata ID matching
    // =========================================================================

    #[test]
    fn test_clipboard_id_format() {
        // Verify the ID format we write to clipboard metadata
        let id: u128 = 12345678901234567890;
        let expected = format!("\"{}\"", id);
        assert_eq!(expected, "\"12345678901234567890\"");
        // This is valid JSON string format
    }
}

#[cfg(test)]
mod autofit_tests {
    use super::estimate_text_width;

    // The window-less estimator is what runs at import time. The old one
    // multiplied str::len() — BYTES — so any non-ASCII text measured far too
    // wide. These pin the properties that bug violated.
    #[test]
    fn counts_characters_not_bytes() {
        // "café" is 5 bytes, 4 chars. It must not measure wider than "cafe".
        let accented = estimate_text_width("café", 14.0, false);
        let ascii = estimate_text_width("cafe", 14.0, false);
        assert!(
            (accented - ascii).abs() < 0.01,
            "accented text measured {} vs plain {} — byte length leaking in",
            accented, ascii
        );
    }

    #[test]
    fn east_asian_text_is_double_width_not_triple() {
        // Each CJK char is 3 bytes but renders about 2 columns wide.
        let cjk = estimate_text_width("日本語", 14.0, false);
        let three_ascii = estimate_text_width("abc", 14.0, false);
        let ratio = cjk / three_ascii;
        assert!(
            (1.9..=2.1).contains(&ratio),
            "CJK measured {}x ASCII of the same length; expected ~2x",
            ratio
        );
    }

    #[test]
    fn scales_with_font_size_and_weight() {
        let small = estimate_text_width("hello", 10.0, false);
        let large = estimate_text_width("hello", 20.0, false);
        assert!(large > small * 1.9, "width should track font size");
        assert!(
            estimate_text_width("hello", 14.0, true) > estimate_text_width("hello", 14.0, false),
            "bold text is wider"
        );
    }

    #[test]
    fn empty_text_has_no_width() {
        assert_eq!(estimate_text_width("", 14.0, false), 0.0);
    }
}

#[cfg(test)]
mod layout_geometry_tests {
    use super::{
        CELL_HEIGHT, CELL_WIDTH, COLUMN_HEADER_HEIGHT, FORMULA_BAR_HEIGHT,
        MACOS_TITLEBAR_HEIGHT, MENU_BAR_HEIGHT, ROW_RESIZE_GRAB_PX, COL_RESIZE_GRAB_PX,
    };

    // =========================================================================
    // LAYOUT GEOMETRY: Ensure hit-testing coordinates stay aligned with rendering.
    // These tests catch the class of bug where top_chrome_height drifts from
    // the actual rendered layout (missing bars, unscaled constants, etc.).
    // =========================================================================

    /// Resize grab zone must be less than half the minimum row height.
    /// If this fails, a click at the vertical center of a row header
    /// would land in the resize zone, making row selection impossible.
    #[test]
    fn resize_grab_smaller_than_half_row_height() {
        assert!(
            ROW_RESIZE_GRAB_PX < CELL_HEIGHT / 2.0,
            "ROW_RESIZE_GRAB_PX ({}) must be < CELL_HEIGHT/2 ({}) \
             or center clicks will hit the resize zone",
            ROW_RESIZE_GRAB_PX, CELL_HEIGHT / 2.0,
        );
        assert!(
            COL_RESIZE_GRAB_PX < CELL_WIDTH / 2.0,
            "COL_RESIZE_GRAB_PX ({}) must be < CELL_WIDTH/2 ({}) \
             or center clicks will hit the resize zone",
            COL_RESIZE_GRAB_PX, CELL_WIDTH / 2.0,
        );
    }

    /// Simulate the row header resize-area check using the same math as
    /// headers.rs. A click at the vertical center of a row must NOT trigger
    /// the resize early-return, regardless of how much chrome is above.
    #[test]
    fn row_header_center_click_is_not_resize() {
        // Worst-case chrome height: Linux, all bars visible, default zoom
        let grid_body_y = MENU_BAR_HEIGHT
            + FORMULA_BAR_HEIGHT
            + crate::views::format_bar::FORMAT_BAR_HEIGHT
            + COLUMN_HEADER_HEIGHT;
        let row_height = CELL_HEIGHT; // default row height
        let row_y_offset = 0.0; // first visible row

        // Click at exact vertical center
        let click_y = grid_body_y + row_y_offset + row_height / 2.0;
        let row_end_y = grid_body_y + row_y_offset + row_height;
        let resize_start = row_end_y - ROW_RESIZE_GRAB_PX;

        assert!(
            click_y < resize_start,
            "Center click y={click_y} must be below resize_start={resize_start} \
             (grid_body_y={grid_body_y}, row_height={row_height}, grab={ROW_RESIZE_GRAB_PX})"
        );
    }

    /// The same check for macOS (includes titlebar).
    #[test]
    fn row_header_center_click_macos_layout() {
        let grid_body_y = MACOS_TITLEBAR_HEIGHT
            + FORMULA_BAR_HEIGHT
            + crate::views::format_bar::FORMAT_BAR_HEIGHT
            + COLUMN_HEADER_HEIGHT;
        let row_height = CELL_HEIGHT;

        let click_y = grid_body_y + row_height / 2.0;
        let row_end_y = grid_body_y + row_height;
        let resize_start = row_end_y - ROW_RESIZE_GRAB_PX;

        assert!(click_y < resize_start);
    }

    /// Verify that top_chrome_height components add up correctly for Linux.
    /// If someone adds a new bar above the grid and forgets top_chrome_height(),
    /// this test's expected value will be stale — that's the point.
    #[test]
    fn linux_top_chrome_components() {
        // Linux, not zen, format bar visible, zoom 1.0
        let expected = MENU_BAR_HEIGHT            // 28
            + FORMULA_BAR_HEIGHT                   // 28
            + crate::views::format_bar::FORMAT_BAR_HEIGHT // 28
            + COLUMN_HEADER_HEIGHT;                // 24  → 108
        assert_eq!(expected, 108.0,
            "Linux top chrome changed — update top_chrome_height() if you added/removed a bar");
    }

    /// Verify macOS chrome height.
    #[test]
    fn macos_top_chrome_components() {
        let expected = MACOS_TITLEBAR_HEIGHT       // 34
            + FORMULA_BAR_HEIGHT                   // 28
            + crate::views::format_bar::FORMAT_BAR_HEIGHT // 28
            + COLUMN_HEADER_HEIGHT;                // 24  → 114
        assert_eq!(expected, 114.0,
            "macOS top chrome changed — update top_chrome_height() if you added/removed a bar");
    }

    /// Zen mode: only column headers remain.
    #[test]
    fn zen_mode_chrome_is_just_col_headers() {
        // In zen mode, top_chrome_height returns only metrics.header_h.
        // At zoom 1.0, that equals COLUMN_HEADER_HEIGHT.
        assert_eq!(COLUMN_HEADER_HEIGHT, 24.0);
    }
}
