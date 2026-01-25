use gpui::*;
use std::collections::HashMap;
use std::path::PathBuf;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::CellLookup;
use visigrid_engine::filter::{RowView, FilterState};
use visigrid_engine::provenance::{MutationOp, SortKey};

use crate::clipboard::InternalClipboard;
use crate::find_replace::MatchHit;
use crate::formatting::BorderApplyMode;
use crate::history::{History, CellChange};
use crate::mode::Mode;
use crate::search::{SearchEngine, SearchAction, CommandId, CommandSearchProvider, GoToSearchProvider, SearchItem, MenuCategory};
use crate::session::SessionManager;
use crate::settings::{
    user_settings_path, open_settings_file, user_settings, update_user_settings,
    observe_settings, Setting, TipId,
};
use crate::theme::{Theme, TokenKey, default_theme, builtin_themes, get_theme};
use crate::views;
use crate::links;
use crate::formula_context::{tokenize_for_highlight, TokenType, char_to_byte};
use crate::workbook_view::WorkbookViewState;

// Re-export from autocomplete module for external access
pub use crate::autocomplete::{SignatureHelpInfo, FormulaErrorInfo};

// ============================================================================
// Global Book Counter (for "Book1", "Book2", etc.)
// ============================================================================

use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::OnceLock;

/// Session-level counter for new workbook names.
/// Increments each time a new workbook is created: Book1, Book2, Book3...
static NEXT_BOOK_NUMBER: AtomicU32 = AtomicU32::new(1);

/// Generate the next book name (e.g., "Book1", "Book2", ...)
pub fn next_book_name() -> String {
    let n = NEXT_BOOK_NUMBER.fetch_add(1, Ordering::Relaxed);
    format!("Book{}", n)
}

// ============================================================================
// Smoke Mode Recalc (Phase 1.5 - headless dogfooding)
// ============================================================================

/// Check if smoke recalc is enabled via VISIGRID_RECALC=full env var.
static SMOKE_RECALC_ENABLED: OnceLock<bool> = OnceLock::new();

fn is_smoke_recalc_enabled() -> bool {
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
    // Future: Selection, Context, History, etc.
}

// ============================================================================
// Document Identity (for title bar display)
// ============================================================================

/// Native file extension for VisiGrid documents
pub const NATIVE_EXT: &str = "vgrid";

/// Returns true if the extension is considered "native" (no provenance needed).
/// Native formats: vgrid (our format), xlsx/xls (Excel, first-class support)
pub fn is_native_ext(ext: &str) -> bool {
    matches!(ext.to_lowercase().as_str(), "vgrid" | "xlsx" | "xls" | "xlsb" | "xlsm" | "sheet")
}

/// Extract display filename from path (full name with extension)
fn display_filename(path: &std::path::Path) -> String {
    path.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("Untitled")
        .to_string()
}

/// Extract lowercase extension from path
fn ext_lower(path: &std::path::Path) -> Option<String> {
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

/// Stable key for formula reference deduplication - same ref gets same color
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum RefKey {
    Cell { row: usize, col: usize },
    Range { r1: usize, c1: usize, r2: usize, c2: usize },  // normalized min/max
}

/// Formula reference with color assignment and text position
#[derive(Clone, Debug)]
pub struct FormulaRef {
    pub key: RefKey,
    pub start: (usize, usize),                // top-left of range
    pub end: Option<(usize, usize)>,          // bottom-right (None for single cell)
    pub color_index: usize,                   // 0-7 rotating
    pub text_range: std::ops::Range<usize>,   // char range in formula text
}

/// Color palette for formula references (Excel-like)
pub const REF_COLORS: [u32; 8] = [
    0x4472C4,  // 0: Blue
    0xED7D31,  // 1: Orange
    0x9B59B6,  // 2: Purple
    0x70AD47,  // 3: Green
    0x00B0F0,  // 4: Cyan
    0xFFC000,  // 5: Yellow
    0xFF6B9D,  // 6: Pink
    0x00B294,  // 7: Teal
];

/// Which field has focus in the Create Named Range dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CreateNameFocus {
    #[default]
    Name,        // Name input field
    Description, // Description input field
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
        /// The active cell when drag started (source of fill)
        anchor: (usize, usize),
        /// Current hover cell during drag
        current: (usize, usize),
        /// Axis lock (None until threshold crossed, then locked)
        axis: Option<FillAxis>,
    },
}

use visigrid_engine::cell::{Alignment, VerticalAlignment, TextOverflow, NumberFormat};

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
    pub font_family: TriState<Option<String>>,
    pub alignment: TriState<Alignment>,
    pub vertical_alignment: TriState<VerticalAlignment>,
    pub text_overflow: TriState<TextOverflow>,
    pub number_format: TriState<NumberFormat>,
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
            font_family: TriState::Empty,
            alignment: TriState::Empty,
            vertical_alignment: TriState::Empty,
            text_overflow: TriState::Empty,
            number_format: TriState::Empty,
        }
    }
}

// Grid configuration
pub const NUM_ROWS: usize = 65536;
pub const NUM_COLS: usize = 256;
pub const CELL_WIDTH: f32 = 80.0;
pub const CELL_HEIGHT: f32 = 24.0;
pub const HEADER_WIDTH: f32 = 50.0;
pub const MENU_BAR_HEIGHT: f32 = 28.0;
pub const FORMULA_BAR_HEIGHT: f32 = 28.0;
pub const COLUMN_HEADER_HEIGHT: f32 = 24.0;
pub const STATUS_BAR_HEIGHT: f32 = 24.0;

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
#[derive(Clone, Copy)]
pub struct GridMetrics {
    pub zoom: f32,
    pub cell_w: f32,
    pub cell_h: f32,
    pub header_w: f32,
    pub header_h: f32,
    pub font_size: f32,
}

impl GridMetrics {
    pub fn new(zoom: f32) -> Self {
        Self {
            zoom,
            cell_w: CELL_WIDTH * zoom,
            cell_h: CELL_HEIGHT * zoom,
            header_w: HEADER_WIDTH * zoom,
            header_h: COLUMN_HEADER_HEIGHT * zoom,
            font_size: 13.0 * zoom,
        }
    }

    /// Get scaled width for a column (model width * zoom)
    pub fn col_width(&self, model_width: f32) -> f32 {
        model_width * self.zoom
    }

    /// Get scaled height for a row (model height * zoom)
    pub fn row_height(&self, model_height: f32) -> f32 {
        model_height * self.zoom
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

pub struct Spreadsheet {
    // Core data
    pub workbook: Workbook,
    pub history: History,

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
    caret_blink_task: Option<gpui::Task<()>>,

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

    // File state
    pub current_file: Option<PathBuf>,
    pub is_modified: bool,  // Legacy - use is_dirty() for title bar
    pub recent_files: Vec<PathBuf>,  // Recently opened files (most recent first)
    pub recent_commands: Vec<CommandId>,  // Recently executed commands (most recent first)

    // Document identity (for title bar)
    pub document_meta: DocumentMeta,
    cached_title: Option<String>,  // For debouncing title updates
    pending_title_refresh: bool,   // Set true + notify() when title may have changed without window access

    // UI state
    pub focus_handle: FocusHandle,
    pub console_focus_handle: FocusHandle,
    pub status_message: Option<String>,
    pub window_size: Size<Pixels>,
    pub cached_window_bounds: Option<WindowBounds>,  // Cached for session snapshot

    // Column/row sizing
    pub col_widths: HashMap<usize, f32>,   // Custom column widths (default: CELL_WIDTH)
    pub row_heights: HashMap<usize, f32>,  // Custom row heights (default: CELL_HEIGHT)

    // Resize drag state
    pub resizing_col: Option<usize>,       // Column being resized (by right edge)
    pub resizing_row: Option<usize>,       // Row being resized (by bottom edge)
    pub resize_start_pos: f32,             // Mouse position at drag start
    pub resize_start_size: f32,            // Original size at drag start

    // Menu bar state (Excel 2003 style dropdown menus)
    pub open_menu: Option<crate::mode::Menu>,

    // Sheet tab state
    pub renaming_sheet: Option<usize>,     // Index of sheet being renamed
    pub sheet_rename_input: String,        // Current rename input value
    pub sheet_context_menu: Option<usize>, // Index of sheet with open context menu

    // Font picker state
    pub available_fonts: Vec<String>,      // System fonts
    pub font_picker_query: String,         // Filter query
    pub font_picker_selected: usize,       // Selected item index

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

    // Highlighted formula references (for existing formulas when editing)
    // Each entry has color index, cell bounds, and text position for formula bar coloring
    pub formula_highlighted_refs: Vec<FormulaRef>,

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

    // Formula autocomplete state
    pub autocomplete_visible: bool,
    pub autocomplete_suppressed: bool,  // Prevents autocomplete from reopening until text edit
    pub autocomplete_selected: usize,
    pub autocomplete_replace_range: std::ops::Range<usize>,

    // Formula hover documentation state
    pub hover_function: Option<&'static crate::formula_context::FunctionInfo>,

    // Document-level settings (persisted in sidecar file)
    pub doc_settings: crate::settings::DocumentSettings,

    // Inspector panel state
    pub inspector_visible: bool,
    pub inspector_tab: crate::mode::InspectorTab,
    pub inspector_pinned: Option<(usize, usize)>,  // Pinned cell (None = follows selection)
    pub inspector_hover_cell: Option<(usize, usize)>,  // Cell being hovered in inspector (for grid highlight)
    pub inspector_trace_path: Option<Vec<visigrid_engine::cell_id::CellId>>,  // Path trace highlight (Phase 3.5b)
    pub inspector_trace_incomplete: bool,  // True if trace has dynamic refs or was truncated
    pub names_filter_query: String,  // Filter query for Names tab

    // Zen mode (distraction-free editing)
    pub zen_mode: bool,

    // F1 context help (hold-to-peek)
    pub f1_help_visible: bool,

    // Zoom (zoom_level is in view_state, metrics is derived)
    pub metrics: GridMetrics,
    zoom_wheel_accumulator: f32,  // For smooth wheel zoom debounce

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

    // Export report state (for Excel exports with warnings)
    pub export_result: Option<visigrid_io::xlsx::ExportResult>,
    pub export_filename: Option<String>,  // Exported filename for display

    // Keyboard hints state (Vimium-style jump)
    pub hint_state: crate::hints::HintState,

    // Lua scripting state
    pub lua_runtime: crate::scripting::LuaRuntime,
    pub lua_console: crate::scripting::ConsoleState,

    // License dialog state
    pub license_input: String,
    pub license_error: Option<String>,

    // Default app prompt state (macOS title bar chip)
    pub default_app_prompt_state: DefaultAppPromptState,
    pub default_app_prompt_file_type: Option<crate::default_app::SpreadsheetFileType>,
    default_app_prompt_success_timer: Option<std::time::Instant>,
    /// Timestamp when we entered NeedsSettings state (for backoff cutoff)
    needs_settings_entered_at: Option<std::time::Instant>,
    /// How many checks we've done in NeedsSettings (for exponential backoff)
    needs_settings_check_count: u8,

    // Smoke mode recalc guard (prevents reentrant recalc)
    in_smoke_recalc: bool,

    // Phase 2: Verified Mode - deterministic ordered recalc with visible status
    pub verified_mode: bool,
    pub last_recalc_report: Option<visigrid_engine::recalc::RecalcReport>,

    // VisiHub sync state
    pub hub_link: Option<crate::hub::HubLink>,
    pub hub_status: crate::hub::HubStatus,
    pub hub_activity: Option<crate::hub::HubActivity>,
    pub hub_last_check: Option<std::time::Instant>,
    pub hub_last_error: Option<String>,
    pub(crate) hub_check_in_progress: bool,

    // VisiHub auth/link dialog state
    pub hub_token_input: String,
    pub hub_repos: Vec<crate::hub::RepoInfo>,
    pub hub_selected_repo: Option<usize>,
    pub hub_datasets: Vec<crate::hub::DatasetInfo>,
    pub hub_selected_dataset: Option<usize>,
    pub hub_new_dataset_name: String,
    pub hub_link_loading: bool,
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
        let workbook = Workbook::new();

        let focus_handle = cx.focus_handle();
        let console_focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);
        let window_size = window.viewport_size();

        // Get theme from global settings store
        let theme = user_settings(cx).appearance.theme_id
            .as_value()
            .and_then(|id| get_theme(id))
            .unwrap_or_else(default_theme);

        // Subscribe to global settings changes - trigger re-render when settings change
        let settings_subscription = observe_settings(cx, |cx| {
            // Notify all windows to re-render when settings change
            cx.refresh_windows();
        });

        Self {
            workbook,
            history: History::new(),
            row_view: RowView::new(NUM_ROWS),  // Identity mapping, all visible
            filter_state: FilterState::default(),
            filter_dropdown_col: None,
            filter_search_text: String::new(),
            filter_checked_items: std::collections::HashSet::new(),
            view_state: WorkbookViewState::default(),
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
            goto_input: String::new(),
            find_input: String::new(),
            find_results: Vec::new(),
            find_index: 0,
            replace_input: String::new(),
            find_replace_mode: false,
            find_focus_replace: false,
            palette_query: String::new(),
            palette_selected: 0,
            palette_scope: None,
            search_engine: Self::create_search_engine(),
            palette_results: Vec::new(),
            palette_total_results: 0,
            palette_pre_selection: (0, 0),
            palette_pre_selection_end: None,
            palette_pre_scroll: (0, 0),
            palette_previewing: false,
            internal_clipboard: None,
            current_file: None,
            is_modified: false,
            recent_files: Vec::new(),
            recent_commands: Vec::new(),
            document_meta: DocumentMeta::default(),
            cached_title: None,
            pending_title_refresh: false,
            focus_handle,
            console_focus_handle,
            status_message: None,
            window_size,
            cached_window_bounds: Some(window.window_bounds()),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            resizing_col: None,
            resizing_row: None,
            resize_start_pos: 0.0,
            resize_start_size: 0.0,
            open_menu: None,
            renaming_sheet: None,
            sheet_rename_input: String::new(),
            sheet_context_menu: None,
            available_fonts: Self::enumerate_fonts(),
            font_picker_query: String::new(),
            font_picker_selected: 0,
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
            formula_highlighted_refs: Vec::new(),
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
            autocomplete_visible: false,
            autocomplete_suppressed: false,
            autocomplete_selected: 0,
            autocomplete_replace_range: 0..0,
            hover_function: None,
            doc_settings: crate::settings::DocumentSettings::default(),
            inspector_visible: false,
            inspector_tab: crate::mode::InspectorTab::default(),
            inspector_pinned: None,
            inspector_hover_cell: None,
            inspector_trace_path: None,
            inspector_trace_incomplete: false,
            names_filter_query: String::new(),
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

            export_result: None,
            export_filename: None,

            hint_state: crate::hints::HintState::default(),

            zen_mode: false,
            f1_help_visible: false,
            metrics: GridMetrics::default(),
            zoom_wheel_accumulator: 0.0,
            link_open_in_flight: false,

            lua_runtime: crate::scripting::LuaRuntime::default(),
            lua_console: crate::scripting::ConsoleState::default(),

            license_input: String::new(),
            license_error: None,

            default_app_prompt_state: DefaultAppPromptState::Hidden,
            default_app_prompt_file_type: None,
            default_app_prompt_success_timer: None,
            needs_settings_entered_at: None,
            needs_settings_check_count: 0,

            in_smoke_recalc: false,

            verified_mode: false,
            last_recalc_report: None,

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
        }
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
    // Row View Layer (view space <-> data space conversion)
    // ========================================================================
    // UI uses VIEW space (what user sees after sort/filter)
    // Storage uses DATA space (canonical row numbers)
    // Convert at boundaries only

    /// Convert view row to data row
    #[inline]
    pub fn view_to_data(&self, view_row: usize) -> usize {
        self.row_view.view_to_data(view_row)
    }

    /// Convert data row to view row (None if hidden by filter)
    #[inline]
    pub fn data_to_view(&self, data_row: usize) -> Option<usize> {
        self.row_view.data_to_view(data_row)
    }

    /// Get visible row count
    #[inline]
    pub fn visible_row_count(&self) -> usize {
        self.row_view.visible_count()
    }

    /// Get the nth visible row (view_row, data_row) for rendering
    /// Returns None if index is out of bounds
    #[inline]
    pub fn nth_visible_row(&self, visible_index: usize) -> Option<(usize, usize)> {
        let view_row = self.row_view.nth_visible(visible_index)?;
        let data_row = self.row_view.view_to_data(view_row);
        Some((view_row, data_row))
    }

    /// View row indices that are visible after filtering
    /// (Not to be confused with visible_rows() which returns screen row count)
    #[inline]
    pub fn filtered_row_indices(&self) -> &[usize] {
        self.row_view.visible_rows()
    }

    /// Ensure row_view has enough capacity for current sheet
    pub fn ensure_row_view_capacity(&mut self) {
        // For now, use a large default. Later this can track actual data extent.
        let needed = 100000;
        if self.row_view.row_count() < needed {
            self.row_view.resize(needed);
        }
    }

    // ========================================================================
    // Sort Operations
    // ========================================================================

    /// Sort by the column containing the active cell
    ///
    /// This is the ONE command UI calls. Engine owns the permutation.
    pub fn sort_by_current_column(
        &mut self,
        direction: visigrid_engine::filter::SortDirection,
        cx: &mut Context<Self>,
    ) {
        use visigrid_engine::filter::{sort_by_column, SortState};

        // Ensure filter range is set (use current selection if not)
        if self.filter_state.filter_range.is_none() {
            // Auto-detect range: from row 0 to last non-empty row in current column
            let col = self.view_state.selected.1;
            let max_row = self.find_last_data_row(col);
            if max_row == 0 {
                self.status_message = Some("No data to sort".to_string());
                cx.notify();
                return;
            }
            // Set filter range covering the data extent
            self.filter_state.filter_range = Some((0, col, max_row, col));
        }

        let col = self.view_state.selected.1;

        // Create value_at closure (captures sheet for computed values)
        let sheet = self.sheet();
        let value_at = |data_row: usize, c: usize| -> visigrid_engine::formula::eval::Value {
            sheet.get_computed_value(data_row, c)
        };

        // Call engine's sort function
        let (new_order, undo_item) = sort_by_column(
            &self.row_view,
            &self.filter_state,
            value_at,
            col,
            direction,
        );

        // Record undo for sort operation
        let previous_sort_state = undo_item.previous_sort_state.map(|s| {
            (s.column, s.direction == visigrid_engine::filter::SortDirection::Ascending)
        });
        let is_ascending = direction == visigrid_engine::filter::SortDirection::Ascending;

        // Build provenance
        let provenance = if let Some((start_row, start_col, end_row, end_col)) = self.filter_state.filter_range {
            Some(MutationOp::Sort {
                sheet: self.sheet().id,
                range_start_row: start_row,
                range_start_col: start_col,
                range_end_row: end_row,
                range_end_col: end_col,
                keys: vec![SortKey { col, ascending: is_ascending }],
                has_header: false,  // TODO: detect header row
            }.to_provenance(&self.sheet().name))
        } else {
            None
        };

        self.history.record_action_with_provenance(crate::history::UndoAction::SortApplied {
            previous_row_order: undo_item.previous_row_order,
            previous_sort_state,
            new_row_order: new_order.clone(),
            new_sort_state: (col, is_ascending),
        }, provenance);

        // Apply the sort
        self.row_view.apply_sort(new_order);

        // Update filter_state.sort
        self.filter_state.sort = Some(SortState { column: col, direction });

        // Invalidate caches
        self.filter_state.invalidate_all_caches();

        self.is_modified = true;
        self.status_message = Some(format!(
            "Sorted by column {} {}",
            Self::col_to_letter(col),
            if direction == visigrid_engine::filter::SortDirection::Ascending { "A→Z" } else { "Z→A" }
        ));
        cx.notify();
    }

    /// Find the last non-empty row in a column (for auto-detecting sort range)
    fn find_last_data_row(&self, col: usize) -> usize {
        let sheet = self.sheet();
        let mut last_row = 0;
        // Scan up to a reasonable limit (or could use sheet's actual data extent)
        for row in 0..10000 {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                last_row = row;
            }
        }
        last_row
    }

    /// Toggle AutoFilter on/off for current selection
    pub fn toggle_auto_filter(&mut self, cx: &mut Context<Self>) {
        if self.filter_state.is_enabled() {
            // Disable: restore original order, clear filters
            self.row_view.clear_sort();
            self.row_view.clear_filter();
            self.filter_state.disable();
            self.status_message = Some("AutoFilter disabled".to_string());
        } else {
            // Enable: set filter range based on selection or data region
            let (row, col) = self.view_state.selected;
            let max_row = self.find_last_data_row(col);
            let max_col = self.find_last_data_col(row);

            if max_row == 0 && max_col == 0 {
                self.status_message = Some("No data for AutoFilter".to_string());
                cx.notify();
                return;
            }

            // Filter range: row 0 is header, data starts at row 1
            self.filter_state.filter_range = Some((0, 0, max_row, max_col));
            self.status_message = Some(format!(
                "AutoFilter enabled: A1:{}{}",
                Self::col_to_letter(max_col),
                max_row + 1
            ));
        }
        cx.notify();
    }

    /// Open the filter dropdown for a column
    pub fn open_filter_dropdown(&mut self, col: usize, cx: &mut Context<Self>) {
        if !self.filter_state.is_enabled() {
            return;
        }

        // Collect values for the column first (to avoid borrow conflicts)
        let values: Vec<(usize, visigrid_engine::formula::eval::Value)> = if let Some((data_start, _, data_end, _)) = self.filter_state.data_range() {
            let sheet = self.sheet();
            (data_start..=data_end)
                .map(|data_row| (data_row, sheet.get_computed_value(data_row, col)))
                .collect()
        } else {
            Vec::new()
        };

        // Build unique values cache (max 500 unique values, sorted by frequency)
        self.filter_state.build_unique_values_from_vec(col, &values, 500);

        // Initialize checked items: all checked if no filter, or match current filter
        self.filter_checked_items.clear();
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            let col_filter = self.filter_state.column_filters.get(&col);
            for (idx, entry) in unique_vals.iter().enumerate() {
                // If no filter for this column, all items are checked
                // If filter exists, check if this value is in selected set
                let should_check = match col_filter {
                    None => true,
                    Some(cf) => match &cf.selected {
                        None => true, // No selection = all pass
                        Some(selected) => selected.contains(&entry.key),
                    },
                };
                if should_check {
                    self.filter_checked_items.insert(idx);
                }
            }
        }

        self.filter_dropdown_col = Some(col);
        self.filter_search_text.clear();
        cx.notify();
    }

    /// Close the filter dropdown without applying
    pub fn close_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Toggle a value in the filter dropdown
    pub fn toggle_filter_item(&mut self, idx: usize, cx: &mut Context<Self>) {
        if self.filter_checked_items.contains(&idx) {
            self.filter_checked_items.remove(&idx);
        } else {
            self.filter_checked_items.insert(idx);
        }
        cx.notify();
    }

    /// Select all items in filter dropdown
    pub fn filter_select_all(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            self.filter_checked_items.clear();
            for idx in 0..unique_vals.len() {
                self.filter_checked_items.insert(idx);
            }
        }
        cx.notify();
    }

    /// Clear all items in filter dropdown
    pub fn filter_clear_all(&mut self, cx: &mut Context<Self>) {
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Apply the current filter dropdown selection
    pub fn apply_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        let Some(unique_vals) = self.filter_state.get_unique_values(col) else {
            self.close_filter_dropdown(cx);
            return;
        };

        // Build selected set from checked items
        let all_checked = self.filter_checked_items.len() == unique_vals.len();

        if all_checked {
            // All checked = no filter (remove filter for this column)
            self.filter_state.clear_column_filter(col);
        } else {
            // Build HashSet of selected normalized keys
            let selected: std::collections::HashSet<_> = self
                .filter_checked_items
                .iter()
                .filter_map(|&idx| unique_vals.get(idx).map(|e| e.key.clone()))
                .collect();

            self.filter_state.column_filters.insert(
                col,
                visigrid_engine::filter::ColumnFilter {
                    selected: Some(selected),
                    text_filter: None,
                },
            );
        }

        // Apply filters to row_view
        self.apply_all_filters();

        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        self.is_modified = true;
        cx.notify();
    }

    /// Apply all column filters to update visible_mask
    fn apply_all_filters(&mut self) {
        let Some((data_start, min_col, data_end, max_col)) = self.filter_state.data_range() else {
            // No filter range - all visible
            self.row_view.clear_filter();
            return;
        };

        // Build visible_mask for all data rows
        let row_count = self.row_view.row_count();
        let mut visible_mask = vec![true; row_count];

        // Header row always visible
        if let Some(header) = self.filter_state.header_row() {
            visible_mask[header] = true;
        }

        // Check each data row against all column filters
        for data_row in data_start..=data_end {
            if data_row >= row_count {
                break;
            }

            let mut passes = true;
            for col in min_col..=max_col {
                if let Some(col_filter) = self.filter_state.column_filters.get(&col) {
                    if col_filter.is_active() {
                        let value = self.sheet().get_computed_value(data_row, col);
                        let filter_key = visigrid_engine::filter::FilterKey::from_value(&value);
                        if !col_filter.passes(&filter_key) {
                            passes = false;
                            break;
                        }
                    }
                }
            }
            visible_mask[data_row] = passes;
        }

        self.row_view.apply_filter(visible_mask);
    }

    /// Check if a column has an active filter
    pub fn column_has_filter(&self, col: usize) -> bool {
        self.filter_state
            .column_filters
            .get(&col)
            .map_or(false, |f| f.is_active())
    }

    /// Find the last non-empty column in a row
    fn find_last_data_col(&self, row: usize) -> usize {
        let sheet = self.sheet();
        let mut last_col = 0;
        for col in 0..256 {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                last_col = col;
            }
        }
        last_col
    }

    /// Clear sort (restore original data order)
    pub fn clear_sort(&mut self, cx: &mut Context<Self>) {
        self.row_view.clear_sort();
        self.filter_state.sort = None;
        self.is_modified = true;
        self.status_message = Some("Sort cleared".to_string());
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

    // =========================================================================
    // Zoom
    // =========================================================================

    /// Set zoom level (all zoom changes go through this)
    pub fn set_zoom(&mut self, new_zoom: f32, cx: &mut Context<Self>) {
        // Clamp to valid range
        let clamped = new_zoom.max(ZOOM_STEPS[0]).min(ZOOM_STEPS[ZOOM_STEPS.len() - 1]);
        if (clamped - self.view_state.zoom_level).abs() < 0.001 {
            return; // No change
        }
        self.view_state.zoom_level = clamped;
        self.metrics = GridMetrics::new(clamped);
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

    // =========================================================================
    // Freeze Panes
    // =========================================================================

    /// Freeze the top row (row 0)
    pub fn freeze_top_row(&mut self, cx: &mut Context<Self>) {
        self.view_state.frozen_rows = 1;
        self.view_state.frozen_cols = 0;
        self.clamp_scroll_to_freeze(cx);
        self.status_message = Some("Frozen top row".to_string());
        cx.notify();
    }

    /// Freeze the first column (column A)
    pub fn freeze_first_column(&mut self, cx: &mut Context<Self>) {
        self.view_state.frozen_rows = 0;
        self.view_state.frozen_cols = 1;
        self.clamp_scroll_to_freeze(cx);
        self.status_message = Some("Frozen first column".to_string());
        cx.notify();
    }

    /// Freeze panes at the current selection
    /// Freezes all rows above and all columns to the left of the active cell
    pub fn freeze_panes(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selected;
        if row == 0 && col == 0 {
            // Nothing to freeze - show message
            self.status_message = Some("Select a cell to freeze rows above and columns to the left".to_string());
            cx.notify();
            return;
        }
        self.view_state.frozen_rows = row;
        self.view_state.frozen_cols = col;
        self.clamp_scroll_to_freeze(cx);
        let msg = match (row, col) {
            (0, c) => format!("Frozen {} column{}", c, if c == 1 { "" } else { "s" }),
            (r, 0) => format!("Frozen {} row{}", r, if r == 1 { "" } else { "s" }),
            (r, c) => format!("Frozen {} row{} and {} column{}", r, if r == 1 { "" } else { "s" }, c, if c == 1 { "" } else { "s" }),
        };
        self.status_message = Some(msg);
        cx.notify();
    }

    /// Remove all freeze panes
    pub fn unfreeze_panes(&mut self, cx: &mut Context<Self>) {
        if self.view_state.frozen_rows == 0 && self.view_state.frozen_cols == 0 {
            self.status_message = Some("No frozen panes to unfreeze".to_string());
            cx.notify();
            return;
        }
        self.view_state.frozen_rows = 0;
        self.view_state.frozen_cols = 0;
        self.status_message = Some("Unfrozen all panes".to_string());
        cx.notify();
    }

    /// Clamp scroll position to ensure it doesn't overlap with frozen regions
    fn clamp_scroll_to_freeze(&mut self, _cx: &mut Context<Self>) {
        // When freeze panes are active, scrollable region starts after frozen rows/cols
        // Ensure scroll position doesn't show frozen rows/cols in the scrollable area
        if self.view_state.frozen_rows > 0 && self.view_state.scroll_row < self.view_state.frozen_rows {
            self.view_state.scroll_row = self.view_state.frozen_rows;
        }
        if self.view_state.frozen_cols > 0 && self.view_state.scroll_col < self.view_state.frozen_cols {
            self.view_state.scroll_col = self.view_state.frozen_cols;
        }
    }

    /// Check if freeze panes are active
    pub fn has_frozen_panes(&self) -> bool {
        self.view_state.frozen_rows > 0 || self.view_state.frozen_cols > 0
    }

    /// Save document settings to sidecar if document has a path
    fn save_doc_settings_if_needed(&self) {
        if let Some(ref path) = self.current_file {
            // Best-effort save - don't block on errors
            let _ = crate::settings::save_doc_settings(path, &self.doc_settings);
        }
    }

    /// Enumerate available system fonts
    fn enumerate_fonts() -> Vec<String> {
        // Fonts commonly installed on Linux systems
        // TODO: Could use fontconfig to enumerate dynamically
        vec![
            "Adwaita Mono".to_string(),
            "Adwaita Sans".to_string(),
            "CaskaydiaMono Nerd Font".to_string(),
            "iA Writer Mono S".to_string(),
            "iA Writer Duo S".to_string(),
            "iA Writer Quattro S".to_string(),
            "Liberation Mono".to_string(),
            "Liberation Sans".to_string(),
            "Liberation Serif".to_string(),
            "Nimbus Mono PS".to_string(),
            "Nimbus Sans".to_string(),
            "Nimbus Roman".to_string(),
            "Noto Sans Mono".to_string(),
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

    /// Ensure cell search cache is fresh (rebuilds if cells_rev changed)
    /// Returns a reference to the cached entries.
    pub(crate) fn ensure_cell_search_cache_fresh(&mut self) -> &[crate::search::CellEntry] {
        use crate::search::CellEntry;
        use visigrid_engine::cell::CellValue;

        if self.cell_search_cache.cached_rev != self.cells_rev {
            // Cache is stale, rebuild from sparse storage
            let sheet = self.sheet();
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
    pub fn dispatch_action(&mut self, action: SearchAction, cx: &mut Context<Self>) {
        match action {
            SearchAction::RunCommand(cmd) => self.dispatch_command(cmd, cx),
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
                    self.edit_original = self.sheet().get_raw(self.view_state.selected.0, self.view_state.selected.1);
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
    pub fn dispatch_command(&mut self, cmd: CommandId, cx: &mut Context<Self>) {
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
            CommandId::SelectBlanks => self.select_blanks(cx),

            // Editing
            CommandId::FillDown => self.fill_down(cx),
            CommandId::FillRight => self.fill_right(cx),
            CommandId::ClearCells => self.delete_selection(cx),
            CommandId::TrimWhitespace => self.trim_whitespace(cx),
            CommandId::Undo => self.undo(cx),
            CommandId::Redo => self.redo(cx),
            CommandId::AutoSum => self.autosum(cx),

            // Clipboard
            CommandId::Copy => self.copy(cx),
            CommandId::Cut => self.cut(cx),
            CommandId::Paste => self.paste(cx),
            CommandId::PasteValues => self.paste_values(cx),

            // Formatting
            CommandId::ToggleBold => self.toggle_bold(cx),
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

            // Background colors
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
            CommandId::BordersClear => self.apply_borders(BorderApplyMode::Clear, cx),

            // File
            CommandId::NewFile => self.new_file(cx),
            CommandId::OpenFile => self.open_file(cx),
            CommandId::Save => self.save(cx),
            CommandId::SaveAs => self.save_as(cx),
            CommandId::ExportCsv => self.export_csv(cx),
            CommandId::ExportTsv => self.export_tsv(cx),
            CommandId::ExportJson => self.export_json(cx),

            // Appearance
            CommandId::SelectTheme => self.show_theme_picker(cx),
            CommandId::SelectFont => self.show_font_picker(cx),

            // View
            CommandId::ToggleInspector => {
                self.inspector_visible = !self.inspector_visible;
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

            // Help
            CommandId::ShowShortcuts => {
                self.status_message = Some("Shortcuts: Ctrl+D Fill Down, Ctrl+R Fill Right, Ctrl+Enter Multi-edit".into());
                cx.notify();
            }
            CommandId::OpenKeybindings => {
                self.open_keybindings(cx);
            }
            CommandId::ShowAbout => {
                self.status_message = Some("VisiGrid - A spreadsheet for power users".into());
                cx.notify();
            }
            CommandId::TourNamedRanges => {
                self.show_tour(cx);
            }
            CommandId::ShowRefactorLog => {
                self.show_refactor_log(cx);
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

            // VisiHub sync
            CommandId::HubCheckStatus => self.hub_check_status(cx),
            CommandId::HubPull => self.hub_pull(cx),
            CommandId::HubPublish => self.hub_publish(cx),
            CommandId::HubOpenRemoteAsCopy => self.hub_open_remote_as_copy(cx),
            CommandId::HubUnlink => self.hub_unlink(cx),
            CommandId::HubDiagnostics => self.hub_diagnostics(cx),
            CommandId::HubSignIn => self.hub_sign_in(cx),
            CommandId::HubSignOut => self.hub_sign_out(cx),
            CommandId::HubLinkDialog => self.hub_show_link_dialog(cx),
        }

        // Ensure title reflects any state changes from this command.
        // The flag + cache debounce makes this cheap for non-state-changing commands.
        self.request_title_refresh(cx);
    }

    // Menu methods
    pub fn toggle_menu(&mut self, menu: crate::mode::Menu, cx: &mut Context<Self>) {
        if self.open_menu == Some(menu) {
            self.open_menu = None;
        } else {
            self.open_menu = Some(menu);
        }
        cx.notify();
    }

    pub fn close_menu(&mut self, cx: &mut Context<Self>) {
        if self.open_menu.is_some() {
            self.open_menu = None;
            cx.notify();
        }
    }

    // ========================================================================
    // Session persistence
    // ========================================================================

    /// Update the global session with this window's current state.
    /// Called on significant state changes (file open/save, panel toggles).
    pub fn update_session(&self, window: &Window, cx: &mut Context<Self>) {
        let snapshot = self.snapshot(window);
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Update session using cached window bounds (for use without Window access).
    /// Useful from file_ops or other places where Window isn't available.
    pub fn update_session_cached(&self, cx: &mut Context<Self>) {
        let snapshot = self.snapshot_cached();
        self.update_session_with_snapshot(snapshot, cx);
    }

    /// Internal: update session with a snapshot
    fn update_session_with_snapshot(&self, snapshot: crate::session::WindowSession, cx: &mut Context<Self>) {
        cx.update_global::<SessionManager, _>(|mgr, _| {
            // Find and update this window's entry, or add a new one
            // For now, we use the file path as the key (simple single-window case)
            let session = mgr.session_mut();

            // Find existing window by file path, or add new
            let idx = session.windows.iter().position(|w| w.file == snapshot.file);

            if let Some(idx) = idx {
                session.windows[idx] = snapshot;
            } else {
                session.windows.push(snapshot);
            }
        });
    }

    /// Save session immediately (for quit/close).
    /// This saves the session to disk synchronously.
    pub fn save_session_now(&self, window: &Window, cx: &mut Context<Self>) {
        self.update_session(window, cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    /// Save session using cached window bounds (for use without Window access).
    pub fn save_session_cached(&self, cx: &mut Context<Self>) {
        self.update_session_cached(cx);
        cx.update_global::<SessionManager, _>(|mgr, _| {
            mgr.save_now();
        });
    }

    // Sheet access convenience methods
    /// Get a reference to the active sheet
    pub fn sheet(&self) -> &Sheet {
        self.workbook.active_sheet()
    }

    /// Get a mutable reference to the active sheet
    pub fn sheet_mut(&mut self) -> &mut Sheet {
        self.workbook.active_sheet_mut()
    }

    /// Get the active sheet index (for undo history)
    pub fn sheet_index(&self) -> usize {
        self.workbook.active_sheet_index()
    }

    /// Set a cell value and update the dependency graph.
    /// This is the preferred way to set cell values - it ensures the dep graph stays in sync.
    pub fn set_cell_value(&mut self, row: usize, col: usize, value: &str) {
        let sheet_id = self.workbook.active_sheet_id();
        self.workbook.active_sheet_mut().set_value(row, col, value);
        self.workbook.update_cell_deps(sheet_id, row, col);
    }

    // =========================================================================
    // Document identity and title bar
    // =========================================================================

    /// Returns true if document has unsaved changes.
    /// Computed from history state, not tracked manually.
    pub fn is_dirty(&self) -> bool {
        self.history.is_dirty()
    }

    /// Update window title if it changed (debounced).
    /// This is the ONLY way titles should update.
    pub fn update_title_if_needed(&mut self, window: &mut Window) {
        let title = self.document_meta.title_string(self.is_dirty());
        if self.cached_title.as_deref() != Some(&title) {
            window.set_window_title(&title);
            self.cached_title = Some(title);
        }
    }

    /// Invalidate the title cache (forces update on next update_title_if_needed call).
    /// Use this when title-affecting state changes but you don't have window access.
    /// Request a title refresh on the next UI pass.
    /// Use when title-affecting state changes but you don't have window access.
    pub fn request_title_refresh(&mut self, cx: &mut Context<Self>) {
        self.cached_title = None;
        self.pending_title_refresh = true;
        cx.notify();
    }

    /// Finalize document state after loading a file
    pub fn finalize_load(&mut self, path: &std::path::Path) {
        let ext = ext_lower(path);
        let filename = display_filename(path);
        let is_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

        // Determine source and saved state based on file type
        let (source, is_saved) = if is_native {
            // Native formats - no provenance, considered "saved"
            (None, true)
        } else {
            // Import formats - show provenance, not "saved" until Save As
            (Some(DocumentSource::Imported { filename: filename.clone() }), false)
        };

        self.document_meta = DocumentMeta {
            display_name: filename,
            is_saved,
            is_read_only: false,
            source,
            path: Some(path.to_path_buf()),
        };

        // Keep current_file in sync (legacy)
        self.current_file = Some(path.to_path_buf());

        // CRITICAL: Set save point AFTER load completes
        // This ensures the document starts "clean" (not dirty)
        self.history.mark_saved();
    }

    /// Finalize document state after saving
    pub fn finalize_save(&mut self, path: &std::path::Path) {
        let ext = ext_lower(path);
        let becomes_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

        self.document_meta.display_name = display_filename(path);
        self.document_meta.path = Some(path.to_path_buf());
        self.history.mark_saved();

        if becomes_native {
            // Saving to native format clears import provenance
            self.document_meta.source = None;
            self.document_meta.is_saved = true;
        }
        // Note: Exporting to CSV/JSON does NOT clear provenance or mark as saved

        // Keep legacy fields in sync
        self.current_file = Some(path.to_path_buf());
        self.is_modified = false;
    }

    // Sheet navigation methods
    /// Move to the next sheet
    pub fn next_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.next_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Move to the previous sheet
    pub fn prev_sheet(&mut self, cx: &mut Context<Self>) {
        if self.workbook.prev_sheet() {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Switch to a specific sheet by index
    pub fn goto_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        // Commit any pending edit before switching sheets
        self.commit_pending_edit();
        if self.workbook.set_active_sheet(index) {
            self.clear_selection_state();
            cx.notify();
        }
    }

    /// Add a new sheet and switch to it
    pub fn add_sheet(&mut self, cx: &mut Context<Self>) {
        let new_index = self.workbook.add_sheet();
        self.workbook.set_active_sheet(new_index);
        self.clear_selection_state();
        self.is_modified = true;
        cx.notify();
    }

    /// Clear selection state when switching sheets
    fn clear_selection_state(&mut self) {
        self.view_state.selected = (0, 0);
        self.view_state.selection_end = None;
        self.view_state.scroll_row = 0;
        self.view_state.scroll_col = 0;
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
    }

    // Sheet rename methods
    /// Start renaming a sheet (double-click on tab)
    pub fn start_sheet_rename(&mut self, index: usize, cx: &mut Context<Self>) {
        if let Some(name) = self.workbook.sheet_names().get(index) {
            self.renaming_sheet = Some(index);
            self.sheet_rename_input = name.to_string();
            self.sheet_context_menu = None;
            cx.notify();
        }
    }

    /// Confirm the sheet rename
    pub fn confirm_sheet_rename(&mut self, cx: &mut Context<Self>) {
        if let Some(index) = self.renaming_sheet {
            let new_name = self.sheet_rename_input.trim();
            if !new_name.is_empty() {
                self.workbook.rename_sheet(index, new_name);
                self.is_modified = true;
            }
            self.renaming_sheet = None;
            self.sheet_rename_input.clear();
            self.request_title_refresh(cx);
        }
    }

    /// Cancel the sheet rename
    pub fn cancel_sheet_rename(&mut self, cx: &mut Context<Self>) {
        self.renaming_sheet = None;
        self.sheet_rename_input.clear();
        cx.notify();
    }

    /// Handle input for sheet rename
    pub fn sheet_rename_input_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.push(c);
            cx.notify();
        }
    }

    /// Handle backspace for sheet rename
    pub fn sheet_rename_backspace(&mut self, cx: &mut Context<Self>) {
        if self.renaming_sheet.is_some() {
            self.sheet_rename_input.pop();
            cx.notify();
        }
    }

    // Sheet context menu methods
    /// Show context menu for a sheet tab
    pub fn show_sheet_context_menu(&mut self, index: usize, cx: &mut Context<Self>) {
        self.sheet_context_menu = Some(index);
        self.renaming_sheet = None;
        cx.notify();
    }

    /// Hide sheet context menu
    pub fn hide_sheet_context_menu(&mut self, cx: &mut Context<Self>) {
        self.sheet_context_menu = None;
        cx.notify();
    }

    /// Delete a sheet
    pub fn delete_sheet(&mut self, index: usize, cx: &mut Context<Self>) {
        if self.workbook.delete_sheet(index) {
            self.is_modified = true;
            self.sheet_context_menu = None;
            self.request_title_refresh(cx);
        } else {
            self.status_message = Some("Cannot delete the last sheet".to_string());
            self.sheet_context_menu = None;
            cx.notify();
        }
    }

    /// Get width for a column (custom or default)
    pub fn col_width(&self, col: usize) -> f32 {
        *self.col_widths.get(&col).unwrap_or(&CELL_WIDTH)
    }

    /// Get height for a row (custom or default)
    pub fn row_height(&self, row: usize) -> f32 {
        *self.row_heights.get(&row).unwrap_or(&CELL_HEIGHT)
    }

    /// Set column width
    pub fn set_col_width(&mut self, col: usize, width: f32) {
        let width = width.max(20.0).min(500.0); // Clamp between 20-500px
        if (width - CELL_WIDTH).abs() < 1.0 {
            self.col_widths.remove(&col); // Remove if close to default
        } else {
            self.col_widths.insert(col, width);
        }
    }

    /// Set row height
    pub fn set_row_height(&mut self, row: usize, height: f32) {
        let height = height.max(12.0).min(200.0); // Clamp between 12-200px
        if (height - CELL_HEIGHT).abs() < 1.0 {
            self.row_heights.remove(&row); // Remove if close to default
        } else {
            self.row_heights.insert(row, height);
        }
    }

    /// Get the X position of a column's left edge (relative to start of grid, after row header)
    /// Returns scaled (zoomed) position for rendering.
    pub fn col_x_offset(&self, target_col: usize) -> f32 {
        let mut x = 0.0;
        for col in self.view_state.scroll_col..target_col {
            x += self.metrics.col_width(self.col_width(col));
        }
        x
    }

    /// Get the Y position of a row's top edge (relative to start of grid, after column header)
    /// Returns scaled (zoomed) position for rendering.
    pub fn row_y_offset(&self, target_row: usize) -> f32 {
        let mut y = 0.0;
        for row in self.view_state.scroll_row..target_row {
            y += self.metrics.row_height(self.row_height(row));
        }
        y
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
        if self.row_heights.is_empty() {
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
    pub fn auto_fit_col_width(&mut self, col: usize, cx: &mut Context<Self>) {
        let mut max_width: f32 = 40.0; // Minimum width

        // Check all rows for content in this column
        for row in 0..NUM_ROWS {
            let text = self.sheet().get_text(row, col);
            if !text.is_empty() {
                // Estimate width: ~7px per character + padding
                let estimated_width = text.len() as f32 * 7.5 + 16.0;
                max_width = max_width.max(estimated_width);
            }
        }

        self.set_col_width(col, max_width);
        cx.notify();
    }

    /// Auto-fit row height to content (for multi-line text in future)
    pub fn auto_fit_row_height(&mut self, row: usize, cx: &mut Context<Self>) {
        // For now, just reset to default since we don't support multi-line
        self.row_heights.remove(&row);
        cx.notify();
    }

    /// Auto-fit column width - if column is selected and multiple columns are selected,
    /// auto-fit all selected columns (Excel behavior)
    pub fn auto_fit_selected_col_widths(&mut self, clicked_col: usize, cx: &mut Context<Self>) {
        // Check if clicked column is part of selection
        if self.is_col_header_selected(clicked_col) {
            // Collect all selected columns
            let mut cols_to_fit = Vec::new();
            for ((_, min_col), (_, max_col)) in self.all_selection_ranges() {
                for col in min_col..=max_col {
                    if !cols_to_fit.contains(&col) {
                        cols_to_fit.push(col);
                    }
                }
            }
            // Auto-fit each selected column
            for col in cols_to_fit {
                self.auto_fit_col_width_no_notify(col);
            }
            cx.notify();
        } else {
            // Not part of selection, just auto-fit the clicked column
            self.auto_fit_col_width(clicked_col, cx);
        }
    }

    /// Auto-fit column width without notifying (for batch operations)
    fn auto_fit_col_width_no_notify(&mut self, col: usize) {
        let mut max_width: f32 = 40.0; // Minimum width
        for row in 0..NUM_ROWS {
            let text = self.sheet().get_text(row, col);
            if !text.is_empty() {
                let estimated_width = text.len() as f32 * 7.5 + 16.0;
                max_width = max_width.max(estimated_width);
            }
        }
        self.set_col_width(col, max_width);
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
        self.row_heights.remove(&row);
    }

    // ========================================================================
    // Cell Reference Helpers (for formula mode)
    // ========================================================================

    /// Convert column index to Excel-style letter(s): 0 -> A, 25 -> Z, 26 -> AA
    pub fn col_to_letter(col: usize) -> String {
        let mut result = String::new();
        let mut n = col;
        loop {
            result.insert(0, (b'A' + (n % 26) as u8) as char);
            if n < 26 {
                break;
            }
            n = n / 26 - 1;
        }
        result
    }

    /// Convert (row, col) to cell reference string: (0, 0) -> "A1"
    pub fn make_cell_ref(row: usize, col: usize) -> String {
        format!("{}{}", Self::col_to_letter(col), row + 1)
    }

    /// Convert range to reference string: ((0, 0), (2, 3)) -> "A1:D3"
    pub fn make_range_ref(start: (usize, usize), end: (usize, usize)) -> String {
        let (r1, c1) = (start.0.min(end.0), start.1.min(end.1));
        let (r2, c2) = (start.0.max(end.0), start.1.max(end.1));
        if r1 == r2 && c1 == c2 {
            Self::make_cell_ref(r1, c1)
        } else {
            format!("{}:{}", Self::make_cell_ref(r1, c1), Self::make_cell_ref(r2, c2))
        }
    }

    /// Check if edit_value starts with = or + (formula indicator)
    pub fn is_formula_content(&self) -> bool {
        self.edit_value.starts_with('=') || self.edit_value.starts_with('+')
    }

    /// Check if a cell is the active reference navigation target (formula_ref_cell).
    /// This is distinct from parsed formula refs - it's where arrow keys are pointing RIGHT NOW.
    /// Used for rendering a bright "target" indicator during formula reference navigation.
    pub fn is_active_ref_target(&self, row: usize, col: usize) -> bool {
        if !self.mode.is_formula() {
            return false;
        }

        if let Some(rect) = self.ref_target_rect() {
            crate::ref_target::contains(&rect, row, col)
        } else {
            false
        }
    }

    /// Get the normalized rectangle for the current ref target, if any.
    fn ref_target_rect(&self) -> Option<crate::ref_target::Rect> {
        let (ref_row, ref_col) = self.formula_ref_cell?;
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((ref_row, ref_col));
        Some(crate::ref_target::normalize_rect((ref_row, ref_col), (end_row, end_col)))
    }

    /// Get the border edges to draw for the active ref target (like selection_borders but for ref target)
    pub fn ref_target_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        if !self.mode.is_formula() {
            return (false, false, false, false);
        }

        let Some(rect) = self.ref_target_rect() else {
            return (false, false, false, false);
        };

        let edges = crate::ref_target::borders(&rect, row, col);
        (edges.top, edges.right, edges.bottom, edges.left)
    }

    /// Check if a cell is within any formula reference (for highlighting)
    /// This includes both the live pointing reference AND parsed refs from existing formulas
    pub fn is_formula_ref(&self, row: usize, col: usize) -> bool {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return false;
        }

        // Check the live pointing reference first
        if let Some((ref_row, ref_col)) = self.formula_ref_cell {
            if let Some((end_row, end_col)) = self.formula_ref_end {
                // Range reference - check if cell is within the range
                let (min_row, max_row) = (ref_row.min(end_row), ref_row.max(end_row));
                let (min_col, max_col) = (ref_col.min(end_col), ref_col.max(end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    return true;
                }
            } else {
                // Single cell reference
                if row == ref_row && col == ref_col {
                    return true;
                }
            }
        }

        // Check the highlighted refs from parsed formula
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                // Range - check if cell is within
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    return true;
                }
            } else {
                // Single cell
                if row == fref.start.0 && col == fref.start.1 {
                    return true;
                }
            }
        }

        false
    }

    /// Get the color index for a formula reference at this cell (for multi-color highlighting).
    /// Returns the earliest ref's color (by text position) to avoid muddy overlap.
    /// Returns None if cell is not a formula ref.
    pub fn formula_ref_color(&self, row: usize, col: usize) -> Option<usize> {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return None;
        }

        // Check the highlighted refs (already sorted by text position, so first match = earliest)
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                // Range
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    return Some(fref.color_index);
                }
            } else {
                // Single cell
                if row == fref.start.0 && col == fref.start.1 {
                    return Some(fref.color_index);
                }
            }
        }

        None
    }

    /// Get which borders should be drawn for a formula ref cell
    /// Returns (top, right, bottom, left) - true means draw that border
    pub fn formula_ref_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        // Must be in formula mode or editing a formula
        let is_formula_editing = self.mode.is_formula() ||
            (self.mode.is_editing() && self.is_formula_content());

        if !is_formula_editing {
            return (false, false, false, false);
        }

        let mut top = false;
        let mut right = false;
        let mut bottom = false;
        let mut left = false;

        // Check the live pointing reference
        if let Some((ref_row, ref_col)) = self.formula_ref_cell {
            if let Some((end_row, end_col)) = self.formula_ref_end {
                let (min_row, max_row) = (ref_row.min(end_row), ref_row.max(end_row));
                let (min_col, max_col) = (ref_col.min(end_col), ref_col.max(end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    if row == min_row { top = true; }
                    if row == max_row { bottom = true; }
                    if col == min_col { left = true; }
                    if col == max_col { right = true; }
                }
            } else {
                // Single cell - all borders
                if row == ref_row && col == ref_col {
                    top = true; right = true; bottom = true; left = true;
                }
            }
        }

        // Check the highlighted refs from parsed formula
        for fref in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = fref.end {
                if row >= fref.start.0 && row <= end_row && col >= fref.start.1 && col <= end_col {
                    if row == fref.start.0 { top = true; }
                    if row == end_row { bottom = true; }
                    if col == fref.start.1 { left = true; }
                    if col == end_col { right = true; }
                }
            } else {
                // Single cell - all borders
                if row == fref.start.0 && col == fref.start.1 {
                    top = true; right = true; bottom = true; left = true;
                }
            }
        }

        (top, right, bottom, left)
    }

    /// Calculate which borders to draw for a selected cell.
    /// Returns (top, right, bottom, left) indicating which borders to draw.
    ///
    /// Strategy:
    /// - Always draw right+bottom (internal gridlines within selection)
    /// - Draw top only if cell above is NOT selected (outer edge)
    /// - Draw left only if cell to left is NOT selected (outer edge)
    /// This maintains the grid appearance while avoiding double borders at edges.
    pub fn selection_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        // Check if adjacent cells are also selected
        let cell_above_selected = row > 0 && self.is_selected(row - 1, col);
        let cell_left_selected = col > 0 && self.is_selected(row, col - 1);

        // Top/left: only at outer edges of selection
        let top = !cell_above_selected;
        let left = !cell_left_selected;

        // Right/bottom: always draw for internal gridlines
        let right = true;
        let bottom = true;

        (top, right, bottom, left)
    }

    /// Compute which user-defined borders to draw for a cell using adjacency logic.
    ///
    /// Returns (top, right, bottom, left) flags indicating which borders to draw.
    /// Uses the precedence rule: right/bottom takes precedence over left/top of adjacent cell.
    ///
    /// - Own right and bottom: always draw if set
    /// - Own top: only draw if cell above has no bottom border
    /// - Own left: only draw if cell to left has no right border
    pub fn cell_user_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        let format = self.sheet().get_format(row, col);

        // Right and bottom: always draw if set
        let right = format.border_right.is_set();
        let bottom = format.border_bottom.is_set();

        // Top: only draw if cell above has no bottom border
        let top = if format.border_top.is_set() {
            if row > 0 {
                let above_format = self.sheet().get_format(row - 1, col);
                !above_format.border_bottom.is_set()
            } else {
                true // No cell above, draw top border
            }
        } else {
            false
        };

        // Left: only draw if cell to left has no right border
        let left = if format.border_left.is_set() {
            if col > 0 {
                let left_format = self.sheet().get_format(row, col - 1);
                !left_format.border_right.is_set()
            } else {
                true // No cell to left, draw left border
            }
        } else {
            false
        };

        (top, right, bottom, left)
    }

    /// Check if any user-defined border is set for this cell
    pub fn has_user_borders(&self, row: usize, col: usize) -> bool {
        let format = self.sheet().get_format(row, col);
        format.border_top.is_set() ||
        format.border_right.is_set() ||
        format.border_bottom.is_set() ||
        format.border_left.is_set()
    }

    /// Calculate visible rows based on window height
    pub fn visible_rows(&self) -> usize {
        let height: f32 = self.window_size.height.into();
        // Menu bar, formula bar, status bar don't scale; column header does
        let available_height = height
            - MENU_BAR_HEIGHT
            - FORMULA_BAR_HEIGHT
            - self.metrics.header_h  // Column header scales with zoom
            - STATUS_BAR_HEIGHT;
        let rows = (available_height / self.metrics.cell_h).floor() as usize;
        rows.max(1).min(NUM_ROWS)
    }

    /// Calculate visible columns based on window width
    pub fn visible_cols(&self) -> usize {
        let width: f32 = self.window_size.width.into();
        let available_width = width - self.metrics.header_w;  // Row header scales with zoom
        let cols = (available_width / self.metrics.cell_w).floor() as usize;
        cols.max(1).min(NUM_COLS)
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

    // Navigation
    pub fn move_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selected;
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.view_state.selected = (new_row, new_col);
        self.view_state.selection_end = None;  // Clear range selection
        self.view_state.additional_selections.clear();  // Clear discontiguous selections

        self.ensure_visible(cx);
    }

    pub fn extend_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.view_state.selection_end = Some((new_row, new_col));

        self.ensure_visible(cx);
    }

    pub fn page_up(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        let visible_rows = self.visible_rows() as i32;
        self.move_selection(-visible_rows, 0, cx);
    }

    pub fn page_down(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }
        let visible_rows = self.visible_rows() as i32;
        self.move_selection(visible_rows, 0, cx);
    }

    /// Find the data boundary in a direction (used by Ctrl+Arrow and Ctrl+Shift+Arrow)
    pub(crate) fn find_data_boundary(&self, start_row: usize, start_col: usize, dr: i32, dc: i32) -> (usize, usize) {
        let mut row = start_row;
        let mut col = start_col;
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        (row, col)
    }

    /// Jump to edge of data region or sheet boundary (Excel-style Ctrl+Arrow)
    pub fn jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (mut row, mut col) = self.view_state.selected;
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.ensure_visible(cx);
    }

    /// Extend selection to edge of data region (Excel-style Ctrl+Shift+Arrow)
    pub fn extend_jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Start from current selection end (or selected if no selection)
        let (mut row, mut col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let current_empty = self.sheet().get_cell(row, col).value.raw_display().is_empty();

        // Check if next cell exists and what it contains
        let peek_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let peek_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        let next_empty = if peek_row == row && peek_col == col {
            true // At edge
        } else {
            self.sheet().get_cell(peek_row, peek_col).value.raw_display().is_empty()
        };

        // Determine search mode: looking for non-empty or looking for empty
        let looking_for_nonempty = current_empty || next_empty;

        loop {
            let next_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let next_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

            // Stop if we hit the edge
            if next_row == row && next_col == col {
                break;
            }

            let cell_empty = self.sheet().get_cell(next_row, next_col).value.raw_display().is_empty();

            if looking_for_nonempty {
                // Scanning through empty space: stop at first non-empty or edge
                row = next_row;
                col = next_col;
                if !cell_empty {
                    break;
                }
            } else {
                // Scanning through data: stop at last non-empty before empty
                if cell_empty {
                    break;
                }
                row = next_row;
                col = next_col;
            }
        }

        // Extend selection to this point (don't move selected, just selection_end)
        self.view_state.selection_end = Some((row, col));
        self.ensure_visible(cx);
    }

    pub fn ensure_visible(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // When freeze panes are active, calculate scrollable region
        let scrollable_visible_rows = visible_rows.saturating_sub(self.view_state.frozen_rows);
        let scrollable_visible_cols = visible_cols.saturating_sub(self.view_state.frozen_cols);

        // Vertical scroll - frozen rows are always visible, only scroll for rows in scrollable region
        if row < self.view_state.frozen_rows {
            // Row is in frozen region - always visible, but ensure scroll_row is valid
            self.view_state.scroll_row = self.view_state.scroll_row.max(self.view_state.frozen_rows);
        } else if row < self.view_state.scroll_row {
            self.view_state.scroll_row = row;
        } else if scrollable_visible_rows > 0 && row >= self.view_state.scroll_row + scrollable_visible_rows {
            self.view_state.scroll_row = row - scrollable_visible_rows + 1;
        }

        // Horizontal scroll - frozen cols are always visible, only scroll for cols in scrollable region
        if col < self.view_state.frozen_cols {
            // Col is in frozen region - always visible, but ensure scroll_col is valid
            self.view_state.scroll_col = self.view_state.scroll_col.max(self.view_state.frozen_cols);
        } else if col < self.view_state.scroll_col {
            self.view_state.scroll_col = col;
        } else if scrollable_visible_cols > 0 && col >= self.view_state.scroll_col + scrollable_visible_cols {
            self.view_state.scroll_col = col - scrollable_visible_cols + 1;
        }

        // Ensure scroll positions don't go below freeze bounds
        self.view_state.scroll_row = self.view_state.scroll_row.max(self.view_state.frozen_rows);
        self.view_state.scroll_col = self.view_state.scroll_col.max(self.view_state.frozen_cols);

        cx.notify();
    }

    pub fn select_cell(&mut self, row: usize, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            self.view_state.selection_end = Some((row, col));
        } else {
            self.view_state.selected = (row, col);
            self.view_state.selection_end = None;
            self.view_state.additional_selections.clear();  // Clear Ctrl+Click selections
            // Clear trace path when selection changes (unless inspector is pinned)
            if self.inspector_pinned.is_none() && self.inspector_trace_path.is_some() {
                self.inspector_trace_path = None;
                self.inspector_trace_incomplete = false;
            }
        }
        cx.notify();
    }

    /// Ctrl+Click to add/toggle cell in selection (discontiguous selection)
    pub fn ctrl_click_cell(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        // Save current selection to additional_selections
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        // Start new selection at clicked cell
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        cx.notify();
    }

    /// Start drag selection - called on mouse_down
    pub fn start_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.view_state.additional_selections.clear();  // Clear Ctrl+Click selections on new drag
        cx.notify();
    }

    /// Start drag selection with Ctrl held (add to existing selections)
    pub fn start_ctrl_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        // Save current selection to additional_selections
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        // Start new selection at clicked cell
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        cx.notify();
    }

    /// Continue drag selection - called on mouse_move while dragging
    pub fn continue_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_selection {
            return;
        }
        // Only update if the cell changed to avoid unnecessary redraws
        if self.view_state.selection_end != Some((row, col)) {
            self.view_state.selection_end = Some((row, col));
            cx.notify();
        }
    }

    /// End drag selection - called on mouse_up
    pub fn end_drag_selection(&mut self, cx: &mut Context<Self>) {
        if self.dragging_selection {
            self.dragging_selection = false;
            cx.notify();
        }
    }

    pub fn select_all(&mut self, cx: &mut Context<Self>) {
        self.view_state.selected = (0, 0);
        self.view_state.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        self.view_state.additional_selections.clear();  // Clear discontiguous selections
        cx.notify();
    }

    // Scrolling
    pub fn scroll(&mut self, delta_rows: i32, delta_cols: i32, cx: &mut Context<Self>) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // When freeze panes are active, scrollable region starts after frozen rows/cols
        let min_scroll_row = self.view_state.frozen_rows;
        let min_scroll_col = self.view_state.frozen_cols;

        let new_row = (self.view_state.scroll_row as i32 + delta_rows)
            .max(min_scroll_row as i32)
            .min((NUM_ROWS.saturating_sub(visible_rows)) as i32) as usize;
        let new_col = (self.view_state.scroll_col as i32 + delta_cols)
            .max(min_scroll_col as i32)
            .min((NUM_COLS.saturating_sub(visible_cols)) as i32) as usize;

        if new_row != self.view_state.scroll_row || new_col != self.view_state.scroll_col {
            self.view_state.scroll_row = new_row;
            self.view_state.scroll_col = new_col;
            cx.notify();
        }
    }

    // Editing
    pub fn start_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selected;

        // Block editing spill receivers - show message and redirect to parent
        if let Some((parent_row, parent_col)) = self.sheet().get_spill_parent(row, col) {
            let parent_ref = self.cell_ref_at(parent_row, parent_col);
            self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
            cx.notify();
            return;
        }

        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = self.edit_original.clone();
        self.edit_cursor = self.edit_value.len();  // Cursor at end (byte offset)
        self.edit_scroll_x = 0.0;
        self.edit_scroll_dirty = true;  // Trigger scroll update to show caret
        self.formula_bar_cache_dirty = true;  // Rebuild hit-test cache
        self.formula_bar_scroll_x = 0.0;
        self.active_editor = EditorSurface::Cell;  // Default to cell editor
        self.edit_selection_anchor = None;

        // Debug assert: cursor must be valid
        debug_assert!(
            self.edit_cursor <= self.edit_value.len(),
            "edit_cursor {} exceeds edit_value.len() {}",
            self.edit_cursor, self.edit_value.len()
        );

        // Parse and highlight formula references if editing a formula
        if self.edit_value.starts_with('=') || self.edit_value.starts_with('+') {
            self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
        } else {
            self.formula_highlighted_refs.clear();
        }

        self.mode = Mode::Edit;
        self.start_caret_blink(cx);
        cx.notify();
    }

    pub fn start_edit_clear(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.view_state.selected;

        // Block editing spill receivers - show message and redirect to parent
        if let Some((parent_row, parent_col)) = self.sheet().get_spill_parent(row, col) {
            let parent_ref = self.cell_ref_at(parent_row, parent_col);
            self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
            cx.notify();
            return;
        }

        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = String::new();
        self.edit_cursor = 0;
        self.edit_scroll_x = 0.0;
        self.edit_scroll_dirty = true;  // Trigger scroll update
        self.formula_bar_cache_dirty = true;  // Rebuild hit-test cache
        self.formula_bar_scroll_x = 0.0;
        self.active_editor = EditorSurface::Cell;  // Default to cell editor
        self.edit_selection_anchor = None;
        self.formula_highlighted_refs.clear();  // No formula to highlight
        self.mode = Mode::Edit;
        self.start_caret_blink(cx);
        cx.notify();
    }

    pub fn confirm_edit(&mut self, cx: &mut Context<Self>) {
        // Multi-edit: If multiple cells selected, apply to all (the "wow" moment)
        if self.is_multi_selection() {
            self.confirm_edit_in_place(cx);
        } else {
            self.confirm_edit_and_move(1, 0, cx);  // Enter moves down
        }
    }

    pub fn confirm_edit_up(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(-1, 0, cx);  // Shift+Enter moves up
    }

    /// Commit any pending edit without moving the cursor.
    /// Call this before file operations (Save, Export) to ensure unsaved edits are captured.
    pub fn commit_pending_edit(&mut self) {
        if !self.mode.is_editing() {
            return;
        }

        let (row, col) = self.view_state.selected;
        let old_value = self.edit_original.clone();

        // Convert leading + to = for formulas (Excel compatibility)
        let mut new_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if new_value.starts_with('=') {
            let open_count = new_value.chars().filter(|&c| c == '(').count();
            let close_count = new_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    new_value.push(')');
                }
            }
        }

        self.history.record_change(self.sheet_index(), row, col, old_value, new_value.clone());
        self.set_cell_value(row, col, &new_value);  // Use helper that updates dep graph
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.bump_cells_rev();
        self.is_modified = true;
        // Clear formula state
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
        self.formula_ref_start_cursor = 0;
        self.formula_highlighted_refs.clear();

        // Smoke mode: trigger full ordered recompute for dogfooding
        self.maybe_smoke_recalc();
    }

    /// Run full ordered recompute if enabled (smoke mode or verified mode).
    ///
    /// - Smoke mode (VISIGRID_RECALC=full): Logs to file for dogfooding
    /// - Verified mode: Updates last_recalc_report for status bar display
    pub(crate) fn maybe_smoke_recalc(&mut self) {
        let smoke_enabled = is_smoke_recalc_enabled();

        // Skip if neither mode is active or we're already in a recalc
        if (!smoke_enabled && !self.verified_mode) || self.in_smoke_recalc {
            return;
        }

        self.in_smoke_recalc = true;
        let report = self.workbook.recompute_full_ordered();

        // Store report for verified mode status bar
        if self.verified_mode {
            self.last_recalc_report = Some(report.clone());
        }

        // Smoke mode logging
        if smoke_enabled {
            let log_line = report.log_line();

            // On Linux/macOS: print to stderr (visible in terminal)
            #[cfg(not(target_os = "windows"))]
            eprintln!("{}", log_line);

            // On all platforms: also write to file (Windows GUI apps don't have stderr)
            use std::io::Write;
            let log_path = dirs::home_dir()
                .map(|p| p.join("smoke.log"))
                .unwrap_or_else(|| std::path::PathBuf::from("smoke.log"));
            if let Ok(mut f) = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(&log_path)
            {
                let _ = writeln!(f, "{}", log_line);
            }
        }

        self.in_smoke_recalc = false;
    }

    /// Toggle verified mode on/off.
    pub fn toggle_verified_mode(&mut self, cx: &mut Context<Self>) {
        self.verified_mode = !self.verified_mode;
        if self.verified_mode {
            // Run initial recalc when enabling
            self.in_smoke_recalc = true;
            let report = self.workbook.recompute_full_ordered();
            self.last_recalc_report = Some(report);
            self.in_smoke_recalc = false;
            self.status_message = Some("Verified mode enabled".to_string());
        } else {
            self.last_recalc_report = None;
            self.status_message = Some("Verified mode disabled".to_string());
        }
        cx.notify();
    }

    pub fn confirm_edit_and_move_right(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, 1, cx);  // Tab moves right
    }

    pub fn confirm_edit_and_move_left(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, -1, cx);  // Shift+Tab moves left
    }

    /// Ctrl+Enter: Multi-edit commit / Fill selection / Open link
    ///
    /// Behavior (Excel muscle memory):
    /// - If editing: apply edit to ALL selected cells with formula shifting
    /// - If navigation + multi-selection: fill selection from primary cell
    /// - If navigation + single cell + link: open link
    /// - If navigation + single cell + no link: start editing
    ///
    /// Multi-edit semantics:
    /// - Applies edited value to all cells in primary selection AND additional_selections
    /// - For formulas: shifts relative references for each target cell
    ///   (e.g., =A1 typed at B2, applied to C3, becomes =B2)
    /// - Absolute references ($A$1) are preserved unchanged
    /// - One undo step for all changes
    pub fn confirm_edit_in_place(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // Navigation mode: fill selection or open link
            if self.is_multi_selection() {
                // Multi-selection: fill from primary cell (Excel Ctrl+Enter)
                self.fill_selection_from_primary(cx);
                return;
            }
            // Single cell: try to open link, else start editing
            if self.try_open_link(cx) {
                return;
            }
            self.start_edit(cx);
            return;
        }

        // Convert leading + to = for formulas (Excel compatibility)
        let mut base_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if base_value.starts_with('=') {
            let open_count = base_value.chars().filter(|&c| c == '(').count();
            let close_count = base_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    base_value.push(')');
                }
            }
        }

        let is_formula = base_value.starts_with('=');
        let primary_cell = self.view_state.selected;  // Base cell for formula reference shifting

        // Collect all target cells from primary selection and additional_selections
        let mut target_cells: Vec<(usize, usize)> = Vec::new();

        // Primary selection rectangle
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                target_cells.push((row, col));
            }
        }

        // Additional selections (Ctrl+Click)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_r = start.0.min(end.0);
            let max_r = start.0.max(end.0);
            let min_c = start.1.min(end.1);
            let max_c = start.1.max(end.1);
            for row in min_r..=max_r {
                for col in min_c..=max_c {
                    // Avoid duplicates (primary selection might overlap)
                    if !target_cells.contains(&(row, col)) {
                        target_cells.push((row, col));
                    }
                }
            }
        }

        let mut changes = Vec::new();

        // Apply to all target cells
        for (row, col) in &target_cells {
            // Skip spill receivers
            if self.sheet().get_spill_parent(*row, *col).is_some() {
                continue;
            }

            let old_value = self.sheet().get_raw(*row, *col);

            // For formulas, shift relative references based on delta from primary cell
            let new_value = if is_formula {
                let delta_row = *row as i32 - primary_cell.0 as i32;
                let delta_col = *col as i32 - primary_cell.1 as i32;
                self.adjust_formula_refs(&base_value, delta_row, delta_col)
            } else {
                base_value.clone()
            };

            if old_value != new_value {
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: new_value.clone(),
                });
            }
            self.sheet_mut().set_value(*row, *col, &new_value);
        }

        // Only attach provenance if multiple cells edited
        let provenance = if target_cells.len() > 1 {
            Some(MutationOp::MultiEdit {
                sheet: self.sheet().id,
                cells: target_cells.clone(),
                value: base_value.clone(),
            }.to_provenance(&self.sheet().name))
        } else {
            None
        };
        self.history.record_batch_with_provenance(self.sheet_index(), changes, provenance);
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.view_state.additional_selections.clear();  // Clear multi-selection after commit
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        // Clear formula highlighting state
        self.formula_highlighted_refs.clear();
        // Reset editor surface
        self.active_editor = EditorSurface::Cell;
        // Stop caret blinking
        self.stop_caret_blink();

        let cell_count = target_cells.len();
        if cell_count > 1 {
            self.status_message = Some(format!("Edited {} cells", cell_count));
        }

        // Smoke mode: trigger full ordered recompute for dogfooding
        self.maybe_smoke_recalc();

        cx.notify();
    }

    /// Fill selection from primary cell (Ctrl+Enter in navigation mode with multi-selection)
    ///
    /// Excel muscle memory: select range, type in first cell, Ctrl+Enter fills all.
    /// This is the navigation-mode equivalent - fills from existing primary cell content.
    ///
    /// - Fills all selected cells with primary cell's content
    /// - If primary is blank, clears all selected cells (Excel behavior)
    /// - Formula references shift relative to primary cell
    /// - One undo step
    fn fill_selection_from_primary(&mut self, cx: &mut Context<Self>) {
        let primary_cell = self.view_state.selected;
        let base_value = self.sheet().get_raw(primary_cell.0, primary_cell.1);

        // If primary cell is empty, we still fill (clears the selection - Excel behavior)

        let is_formula = base_value.starts_with('=');

        // Collect all target cells (excluding primary cell itself)
        let mut target_cells: Vec<(usize, usize)> = Vec::new();

        // Primary selection rectangle
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if (row, col) != primary_cell {
                    target_cells.push((row, col));
                }
            }
        }

        // Additional selections (Ctrl+Click)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_r = start.0.min(end.0);
            let max_r = start.0.max(end.0);
            let min_c = start.1.min(end.1);
            let max_c = start.1.max(end.1);
            for row in min_r..=max_r {
                for col in min_c..=max_c {
                    if (row, col) != primary_cell && !target_cells.contains(&(row, col)) {
                        target_cells.push((row, col));
                    }
                }
            }
        }

        if target_cells.is_empty() {
            return;
        }

        let mut changes = Vec::new();
        let mut filled_count = 0;
        let mut skipped_spill = 0;

        for (row, col) in &target_cells {
            // Skip spill receivers
            if self.sheet().get_spill_parent(*row, *col).is_some() {
                skipped_spill += 1;
                continue;
            }

            let old_value = self.sheet().get_raw(*row, *col);

            // For formulas, shift relative references based on delta from primary cell
            let new_value = if is_formula {
                let delta_row = *row as i32 - primary_cell.0 as i32;
                let delta_col = *col as i32 - primary_cell.1 as i32;
                self.adjust_formula_refs(&base_value, delta_row, delta_col)
            } else {
                base_value.clone()
            };

            if old_value != new_value {
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: new_value.clone(),
                });
            }
            self.sheet_mut().set_value(*row, *col, &new_value);
            filled_count += 1;
        }

        if !changes.is_empty() {
            let provenance = MutationOp::MultiEdit {
                sheet: self.sheet().id,
                cells: target_cells.clone(),
                value: base_value.clone(),
            }.to_provenance(&self.sheet().name);
            self.history.record_batch_with_provenance(self.sheet_index(), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc();
        }

        self.view_state.additional_selections.clear();

        // Status message with optional spill skip note
        let status = if skipped_spill > 0 {
            format!("Filled {} cells (skipped {} spill)", filled_count, skipped_spill)
        } else {
            format!("Filled {} cells", filled_count)
        };
        self.status_message = Some(status);
        cx.notify();
    }

    /// Try to open a detected link in the current cell.
    /// Returns true if a link was found and opened, false otherwise.
    ///
    /// Guards:
    /// - Only works with single-cell selection (multi-selection returns false)
    /// - Debounced: ignores rapid Ctrl+Enter if open is already in-flight
    pub fn try_open_link(&mut self, cx: &mut Context<Self>) -> bool {
        // Guard: only open links from single-cell selection
        // This prevents accidental opens when multi-selecting
        if self.is_multi_selection() {
            return false;
        }

        // Guard: debounce - ignore if already opening a link
        if self.link_open_in_flight {
            return false;
        }

        let (row, col) = self.view_state.selected;
        let cell_value = self.sheet().get_display(row, col);

        if let Some(target) = links::detect_link(&cell_value) {
            let open_string = target.open_string();
            let target_desc = match &target {
                links::LinkTarget::Url(_) => "Opening URL...",
                links::LinkTarget::Email(_) => "Opening email...",
                links::LinkTarget::Path(_) => "Opening file...",
            };

            // Mark as in-flight
            self.link_open_in_flight = true;

            // Open asynchronously to avoid blocking the UI
            // Note: open::that() is non-blocking on most platforms (sends to OS and returns)
            cx.spawn(async move |this, cx| {
                // Run open and capture result
                let result = open::that(&open_string);

                // Always reset in-flight and update status in a single update call
                // If update fails (entity dropped), the flag doesn't matter
                let _ = this.update(cx, |this, cx| {
                    this.link_open_in_flight = false;
                    this.status_message = Some(match result {
                        Ok(()) => format!("Opened: {}", open_string),
                        Err(e) => format!("Couldn't open link: {}", e),
                    });
                    cx.notify();
                });
            }).detach();

            self.status_message = Some(target_desc.to_string());
            cx.notify();
            true
        } else {
            false
        }
    }

    /// Detect link in current cell (for status bar hint)
    pub fn detected_link(&self) -> Option<links::LinkTarget> {
        let (row, col) = self.view_state.selected;
        let cell_value = self.sheet().get_display(row, col);
        links::detect_link(&cell_value)
    }

    fn confirm_edit_and_move(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // Not editing - just move (Excel behavior)
            self.move_selection(dr, dc, cx);
            return;
        }

        let (row, col) = self.view_state.selected;
        let old_value = self.edit_original.clone();

        // Convert leading + to = for formulas (Excel compatibility)
        let mut new_value = if self.edit_value.starts_with('+') {
            format!("={}", &self.edit_value[1..])
        } else {
            self.edit_value.clone()
        };

        // Auto-close unmatched parentheses (Excel compatibility)
        if new_value.starts_with('=') {
            let open_count = new_value.chars().filter(|&c| c == '(').count();
            let close_count = new_value.chars().filter(|&c| c == ')').count();
            if open_count > close_count {
                for _ in 0..(open_count - close_count) {
                    new_value.push(')');
                }
            }
        }

        self.history.record_change(self.sheet_index(), row, col, old_value, new_value.clone());
        self.set_cell_value(row, col, &new_value);  // Use helper that updates dep graph
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        // Clear formula reference state
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
        self.formula_ref_start_cursor = 0;
        // Clear formula highlighting state
        self.formula_highlighted_refs.clear();
        // Reset editor surface
        self.active_editor = EditorSurface::Cell;
        // Stop caret blinking
        self.stop_caret_blink();

        // Smoke mode: trigger full ordered recompute for dogfooding
        self.maybe_smoke_recalc();

        // Move after confirming
        self.move_selection(dr, dc, cx);
    }

    pub fn cancel_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.mode = Mode::Navigation;
            self.edit_value.clear();
            self.edit_cursor = 0;
            // Clear formula reference state
            self.formula_ref_cell = None;
            self.formula_ref_end = None;
            self.formula_ref_start_cursor = 0;
            // Clear formula highlighting state
            self.formula_highlighted_refs.clear();
            // Clear autocomplete state
            self.autocomplete_visible = false;
            self.autocomplete_selected = 0;
            // Reset editor surface
            self.active_editor = EditorSurface::Cell;
            // Stop caret blinking
            self.stop_caret_blink();
            cx.notify();
        }
    }

    // ========================================================================
    // Caret Blinking
    // ========================================================================

    /// Start the caret blink timer. Called when entering edit mode.
    pub fn start_caret_blink(&mut self, cx: &mut Context<Self>) {
        use std::time::Duration;

        self.caret_visible = true;
        self.caret_last_activity = std::time::Instant::now();

        // Cancel any existing blink task
        self.caret_blink_task = None;

        // Spawn repeating blink task
        let task = cx.spawn(async move |this, cx| {
            let blink_interval = Duration::from_millis(530);
            let idle_delay = Duration::from_millis(500);

            loop {
                // Wait for blink interval
                smol::Timer::after(blink_interval).await;

                // Update caret visibility
                let should_continue = this.update(cx, |this, cx| {
                    // Don't blink if not editing
                    if !this.mode.is_editing() {
                        return false;
                    }

                    // Don't blink if there's a text selection
                    if this.edit_selection_anchor.is_some() {
                        this.caret_visible = true;
                        cx.notify();
                        return true;
                    }

                    // Don't blink during active typing (wait for idle)
                    if this.caret_last_activity.elapsed() < idle_delay {
                        this.caret_visible = true;
                        cx.notify();
                        return true;
                    }

                    // Toggle visibility
                    this.caret_visible = !this.caret_visible;
                    cx.notify();
                    true
                });

                match should_continue {
                    Ok(true) => continue,
                    _ => break,
                }
            }
        });

        self.caret_blink_task = Some(task);
    }

    /// Stop the caret blink timer. Called when leaving edit mode.
    pub fn stop_caret_blink(&mut self) {
        self.caret_blink_task = None;
        self.caret_visible = true;
    }

    /// Reset caret activity timestamp. Called on text edits and cursor moves.
    /// Keeps caret visible and resets the idle timer.
    pub fn reset_caret_activity(&mut self) {
        self.caret_visible = true;
        self.caret_last_activity = std::time::Instant::now();
    }

    /// Delete selected text and return true if there was a selection
    fn delete_edit_selection(&mut self) -> bool {
        if let Some((start_byte, end_byte)) = self.edit_selection_range() {
            // start_byte and end_byte are already byte offsets
            let start_byte = start_byte.min(self.edit_value.len());
            let end_byte = end_byte.min(self.edit_value.len());
            self.edit_value.replace_range(start_byte..end_byte, "");
            self.edit_cursor = start_byte;
            self.edit_selection_anchor = None;
            true
        } else {
            false
        }
    }

    pub fn backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // Text edit: clear ref_target and suppression so autocomplete can reopen
            if self.mode.is_formula() {
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            }
            self.autocomplete_suppressed = false;
            self.reset_caret_activity();

            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
                return;
            }
            // Otherwise delete char before cursor (byte-indexed)
            if self.edit_cursor > 0 {
                let prev_byte = self.prev_char_boundary(self.edit_cursor);
                let curr_byte = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.replace_range(prev_byte..curr_byte, "");
                self.edit_cursor = prev_byte;
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
            }
        }
    }

    pub fn delete_char(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // Text edit: clear ref_target and suppression so autocomplete can reopen
            if self.mode.is_formula() {
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            }
            self.autocomplete_suppressed = false;
            self.reset_caret_activity();

            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
                return;
            }
            // Otherwise delete char at cursor (byte-indexed)
            let len = self.edit_value.len();
            if self.edit_cursor < len {
                let curr_byte = self.edit_cursor;
                let next_byte = self.next_char_boundary(curr_byte);
                self.edit_value.replace_range(curr_byte..next_byte, "");
                // Cursor stays at same byte position (deleted forward)
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.edit_scroll_dirty = true;
                self.formula_bar_cache_dirty = true;
                self.update_autocomplete(cx);
                cx.notify();
            }
        }
    }

    pub fn insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // In Formula mode, typing an operator finalizes the current reference
            if self.mode.is_formula() && self.formula_ref_cell.is_some() {
                if Self::is_formula_operator(c) {
                    self.finalize_formula_reference();
                } else {
                    // Non-operator character: clear ref_target since we're typing, not navigating
                    self.formula_ref_cell = None;
                    self.formula_ref_end = None;
                }
            }

            // Delete selection if any (replaces selected text)
            self.delete_edit_selection();

            // Insert at cursor byte position
            let byte_idx = self.edit_cursor.min(self.edit_value.len());
            self.edit_value.insert(byte_idx, c);
            self.edit_cursor = byte_idx + c.len_utf8();  // Advance by byte length of char

            // Update highlighted refs for formulas
            if self.is_formula_content() {
                self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
            }

            // Text edit: clear suppression so autocomplete can reopen
            self.autocomplete_suppressed = false;

            // Reset caret blink (keep visible while typing)
            self.reset_caret_activity();

            // Mark scroll/cache dirty
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;

            // Update autocomplete for formulas
            self.update_autocomplete(cx);
        } else {
            // Start editing with this character
            let (row, col) = self.view_state.selected;

            // Block editing spill receivers
            if let Some((parent_row, parent_col)) = self.sheet().get_spill_parent(row, col) {
                let parent_ref = self.cell_ref_at(parent_row, parent_col);
                self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
                cx.notify();
                return;
            }

            self.edit_original = self.sheet().get_raw(row, col);
            self.edit_value = c.to_string();
            self.edit_cursor = c.len_utf8();  // Byte offset after first char

            // Enter Formula mode if starting with = or +
            if c == '=' || c == '+' {
                self.mode = Mode::Formula;
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            } else {
                self.mode = Mode::Edit;
            }

            // Mark scroll/cache dirty
            self.edit_scroll_dirty = true;
            self.formula_bar_cache_dirty = true;
            self.formula_bar_scroll_x = 0.0;
            self.active_editor = EditorSurface::Cell;

            // Start caret blinking
            self.start_caret_blink(cx);

            // Update autocomplete for formulas
            self.update_autocomplete(cx);
        }
        cx.notify();
    }

    /// Check if character is a formula operator that finalizes a reference
    fn is_formula_operator(c: char) -> bool {
        matches!(c, '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | ',' | '(' | ')' | ':' | ';')
    }

    /// Finalize the current formula reference (clear the active reference state)
    fn finalize_formula_reference(&mut self) {
        self.formula_ref_cell = None;
        self.formula_ref_end = None;
    }

    // Selection helpers
    pub fn selection_range(&self) -> ((usize, usize), (usize, usize)) {
        let start = self.view_state.selected;
        let end = self.view_state.selection_end.unwrap_or(start);
        let min_row = start.0.min(end.0);
        let max_row = start.0.max(end.0);
        let min_col = start.1.min(end.1);
        let max_col = start.1.max(end.1);
        ((min_row, min_col), (max_row, max_col))
    }

    /// Clamp selection to valid bounds after operations that might invalidate it.
    /// Preserves column where possible (user mental model), clamps row to valid range.
    ///
    /// TODO: Centralize view state normalization. Any operation that changes row order
    /// or sheet dimensions (sort, filter, insert/delete rows, paste that expands) should
    /// call a single `normalize_view_state()` that: clamps selection, clamps scroll offsets,
    /// clears hover state if out of bounds, invalidates caches. Currently clamp_selection()
    /// is called in multiple undo/redo handlers - this is the invariant to enforce centrally.
    pub fn clamp_selection(&mut self) {
        // Clamp selected cell
        self.view_state.selected.0 = self.view_state.selected.0.min(NUM_ROWS - 1);
        self.view_state.selected.1 = self.view_state.selected.1.min(NUM_COLS - 1);

        // Clamp selection_end if present
        if let Some(ref mut end) = self.view_state.selection_end {
            end.0 = end.0.min(NUM_ROWS - 1);
            end.1 = end.1.min(NUM_COLS - 1);
        }

        // Clamp additional selections
        for (start, end) in &mut self.view_state.additional_selections {
            start.0 = start.0.min(NUM_ROWS - 1);
            start.1 = start.1.min(NUM_COLS - 1);
            if let Some(ref mut e) = end {
                e.0 = e.0.min(NUM_ROWS - 1);
                e.1 = e.1.min(NUM_COLS - 1);
            }
        }
    }

    /// Returns true if more than one cell is selected.
    /// This includes range selections and Ctrl+Click additional selections.
    pub fn is_multi_selection(&self) -> bool {
        // Check if primary selection is a range (more than one cell)
        if let Some(end) = self.view_state.selection_end {
            if end != self.view_state.selected {
                return true;
            }
        }
        // Check if there are additional Ctrl+Click selections
        if !self.view_state.additional_selections.is_empty() {
            return true;
        }
        false
    }

    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        // Check active selection
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            return true;
        }
        // Check additional selections (Ctrl+Click ranges)
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                return true;
            }
        }
        false
    }

    /// Get multi-edit preview for a cell during editing.
    /// Returns the value that will be applied to this cell when edit is confirmed.
    /// Returns None if not in multi-edit mode or if this is the active cell.
    pub fn multi_edit_preview(&self, row: usize, col: usize) -> Option<String> {
        // Only in editing mode with multi-selection
        if !self.mode.is_editing() || !self.is_multi_selection() {
            return None;
        }
        // Skip the active cell (it shows the real edit_value)
        if (row, col) == self.view_state.selected {
            return None;
        }
        // Only for selected cells
        if !self.is_selected(row, col) {
            return None;
        }

        // Compute delta from primary cell
        let delta_row = row as i32 - self.view_state.selected.0 as i32;
        let delta_col = col as i32 - self.view_state.selected.1 as i32;

        // If it's a formula, adjust references
        if self.edit_value.starts_with('=') {
            Some(self.adjust_formula_refs(&self.edit_value, delta_row, delta_col))
        } else {
            // Plain text: same value for all cells
            Some(self.edit_value.clone())
        }
    }

    /// Get all selection ranges (for operations that apply to all selected cells)
    pub fn all_selection_ranges(&self) -> Vec<((usize, usize), (usize, usize))> {
        let mut ranges = Vec::new();
        // Add active selection
        ranges.push(self.selection_range());
        // Add additional selections
        for (start, end) in &self.view_state.additional_selections {
            let end = end.unwrap_or(*start);
            let min_row = start.0.min(end.0);
            let max_row = start.0.max(end.0);
            let min_col = start.1.min(end.1);
            let max_col = start.1.max(end.1);
            ranges.push(((min_row, min_col), (max_row, max_col)));
        }
        ranges
    }

    // Formatting (applies to all discontiguous selection ranges)
    pub fn toggle_bold(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_bold(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_italic(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_italic(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_underline(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().toggle_underline(row, col);
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn format_currency(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().set_number_format(row, col, NumberFormat::Currency { decimals: 2 });
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn format_percent(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.sheet_mut().set_number_format(row, col, NumberFormat::Percent { decimals: 2 });
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    // Row/Column header selection helpers

    /// Check if the active selection spans all columns (row selection)
    pub fn is_row_selection(&self) -> bool {
        let ((_, min_col), (_, max_col)) = self.selection_range();
        min_col == 0 && max_col == NUM_COLS - 1
    }

    /// Check if the active selection spans all rows (column selection)
    pub fn is_col_selection(&self) -> bool {
        let ((min_row, _), (max_row, _)) = self.selection_range();
        min_row == 0 && max_row == NUM_ROWS - 1
    }

    /// Check if row header should be highlighted (checks all selections)
    pub fn is_row_header_selected(&self, row: usize) -> bool {
        for ((min_row, _), (max_row, _)) in self.all_selection_ranges() {
            if row >= min_row && row <= max_row {
                return true;
            }
        }
        false
    }

    /// Check if column header should be highlighted (checks all selections)
    pub fn is_col_header_selected(&self, col: usize) -> bool {
        for ((_, min_col), (_, max_col)) in self.all_selection_ranges() {
            if col >= min_col && col <= max_col {
                return true;
            }
        }
        false
    }

    /// Select entire row. If extend=true, extends from current anchor row.
    pub fn select_row(&mut self, row: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            // Extend from the current anchor (self.view_state.selected.0 before this call)
            let anchor_row = self.view_state.selected.0;
            self.view_state.selected = (anchor_row.min(row), 0);
            self.view_state.selection_end = Some((anchor_row.max(row), NUM_COLS - 1));
        } else {
            self.view_state.selected = (row, 0);
            self.view_state.selection_end = Some((row, NUM_COLS - 1));
            self.view_state.additional_selections.clear();
        }
        cx.notify();
    }

    /// Select entire column. If extend=true, extends from current anchor col.
    pub fn select_col(&mut self, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            let anchor_col = self.view_state.selected.1;
            self.view_state.selected = (0, anchor_col.min(col));
            self.view_state.selection_end = Some((NUM_ROWS - 1, anchor_col.max(col)));
        } else {
            self.view_state.selected = (0, col);
            self.view_state.selection_end = Some((NUM_ROWS - 1, col));
            self.view_state.additional_selections.clear();
        }
        cx.notify();
    }

    /// Start row header drag - stores stable anchor
    pub fn start_row_header_drag(&mut self, row: usize, cx: &mut Context<Self>) {
        self.dragging_row_header = true;
        self.dragging_col_header = false;
        self.dragging_selection = false;
        self.row_header_anchor = Some(row);
        self.select_row(row, false, cx);
    }

    /// Continue row header drag - uses stored anchor
    pub fn continue_row_header_drag(&mut self, row: usize, cx: &mut Context<Self>) {
        if !self.dragging_row_header { return; }
        let anchor = self.row_header_anchor.unwrap_or(row);
        let min_r = anchor.min(row);
        let max_r = anchor.max(row);
        self.view_state.selected = (min_r, 0);
        self.view_state.selection_end = Some((max_r, NUM_COLS - 1));
        cx.notify();
    }

    /// End row header drag
    pub fn end_row_header_drag(&mut self, _cx: &mut Context<Self>) {
        self.dragging_row_header = false;
        self.row_header_anchor = None;
    }

    /// Start column header drag - stores stable anchor
    pub fn start_col_header_drag(&mut self, col: usize, cx: &mut Context<Self>) {
        self.dragging_col_header = true;
        self.dragging_row_header = false;
        self.dragging_selection = false;
        self.col_header_anchor = Some(col);
        self.select_col(col, false, cx);
    }

    /// Continue column header drag - uses stored anchor
    pub fn continue_col_header_drag(&mut self, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_col_header { return; }
        let anchor = self.col_header_anchor.unwrap_or(col);
        let min_c = anchor.min(col);
        let max_c = anchor.max(col);
        self.view_state.selected = (0, min_c);
        self.view_state.selection_end = Some((NUM_ROWS - 1, max_c));
        cx.notify();
    }

    /// End column header drag
    pub fn end_col_header_drag(&mut self, _cx: &mut Context<Self>) {
        self.dragging_col_header = false;
        self.col_header_anchor = None;
    }

    /// Ctrl+click on row header - add row to additional selections
    pub fn ctrl_click_row(&mut self, row: usize, cx: &mut Context<Self>) {
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        self.select_row(row, false, cx);
    }

    /// Ctrl+click on column header - add column to additional selections
    pub fn ctrl_click_col(&mut self, col: usize, cx: &mut Context<Self>) {
        self.view_state.additional_selections.push((self.view_state.selected, self.view_state.selection_end));
        self.select_col(col, false, cx);
    }

    // Row/Column insert/delete operations (Ctrl+= / Ctrl+-)

    /// Insert rows or columns based on current selection (Ctrl+=)
    pub fn insert_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            // Insert rows above selection
            let ((min_row, _), (max_row, _)) = self.selection_range();
            let count = max_row - min_row + 1;
            self.insert_rows(min_row, count, cx);
        } else if self.is_col_selection() {
            // Insert columns left of selection
            let ((_, min_col), (_, max_col)) = self.selection_range();
            let count = max_col - min_col + 1;
            self.insert_cols(min_col, count, cx);
        } else {
            // v1: No dialog, just show status message
            self.status_message = Some("Select entire row (Shift+Space) or column (Ctrl+Space) first".to_string());
            cx.notify();
        }
    }

    /// Delete rows or columns based on current selection (Ctrl+-)
    pub fn delete_rows_or_cols(&mut self, cx: &mut Context<Self>) {
        // v1: Only operate on primary selection, ignore additional selections
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Delete not supported with multiple selections".to_string());
            cx.notify();
            return;
        }

        if self.is_row_selection() {
            // Delete selected rows
            let ((min_row, _), (max_row, _)) = self.selection_range();
            let count = max_row - min_row + 1;
            self.delete_rows(min_row, count, cx);
        } else if self.is_col_selection() {
            // Delete selected columns
            let ((_, min_col), (_, max_col)) = self.selection_range();
            let count = max_col - min_col + 1;
            self.delete_cols(min_col, count, cx);
        } else {
            // v1: No dialog, just show status message
            self.status_message = Some("Select entire row (Shift+Space) or column (Ctrl+Space) first".to_string());
            cx.notify();
        }
    }

    /// Insert rows at position with undo support
    fn insert_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Perform the insert
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.insert_rows(at_row, count);
        }

        // Shift row heights down (from bottom to avoid overwriting)
        let heights_to_shift: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row)
            .map(|(r, h)| (*r, *h))
            .collect();
        for (r, _) in &heights_to_shift {
            self.row_heights.remove(r);
        }
        for (r, h) in heights_to_shift {
            let new_row = r + count;
            if new_row < NUM_ROWS {
                self.row_heights.insert(new_row, h);
            }
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::RowsInserted {
            sheet_index,
            at_row,
            count,
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} row(s)", count));
        cx.notify();
    }

    /// Delete rows at position with undo support
    fn delete_rows(&mut self, at_row: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        if let Some(sheet) = self.workbook.sheet(sheet_index) {
            for row in at_row..at_row + count {
                for col in 0..NUM_COLS {
                    let raw = sheet.get_raw(row, col);
                    let format = sheet.get_format(row, col);
                    // Only store non-empty cells
                    if !raw.is_empty() || format != Default::default() {
                        deleted_cells.push((row, col, raw, format));
                    }
                }
            }
        }

        // Capture row heights for deleted rows
        let deleted_row_heights: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row && **r < at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();

        // Remove heights for deleted rows and shift remaining up
        let heights_to_shift: Vec<_> = self.row_heights
            .iter()
            .filter(|(r, _)| **r >= at_row + count)
            .map(|(r, h)| (*r, *h))
            .collect();
        // Remove all affected heights
        for r in at_row..NUM_ROWS {
            self.row_heights.remove(&r);
        }
        // Re-insert shifted heights
        for (r, h) in heights_to_shift {
            self.row_heights.insert(r - count, h);
        }

        // Perform the delete
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.delete_rows(at_row, count);
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::RowsDeleted {
            sheet_index,
            at_row,
            count,
            deleted_cells,
            deleted_row_heights,
        });

        // Move selection up if needed
        if self.view_state.selected.0 >= at_row + count {
            self.view_state.selected.0 -= count;
        } else if self.view_state.selected.0 >= at_row {
            self.view_state.selected.0 = at_row.saturating_sub(1);
        }
        self.view_state.selection_end = None;

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} row(s)", count));
        cx.notify();
    }

    /// Insert columns at position with undo support
    fn insert_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Perform the insert
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.insert_cols(at_col, count);
        }

        // Shift column widths right (from right to avoid overwriting)
        let widths_to_shift: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col)
            .map(|(c, w)| (*c, *w))
            .collect();
        for (c, _) in &widths_to_shift {
            self.col_widths.remove(c);
        }
        for (c, w) in widths_to_shift {
            let new_col = c + count;
            if new_col < NUM_COLS {
                self.col_widths.insert(new_col, w);
            }
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::ColsInserted {
            sheet_index,
            at_col,
            count,
        });

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Inserted {} column(s)", count));
        cx.notify();
    }

    /// Delete columns at position with undo support
    fn delete_cols(&mut self, at_col: usize, count: usize, cx: &mut Context<Self>) {
        let sheet_index = self.workbook.active_sheet_index();

        // Capture cells to be deleted for undo
        let mut deleted_cells = Vec::new();
        if let Some(sheet) = self.workbook.sheet(sheet_index) {
            for col in at_col..at_col + count {
                for row in 0..NUM_ROWS {
                    let raw = sheet.get_raw(row, col);
                    let format = sheet.get_format(row, col);
                    // Only store non-empty cells
                    if !raw.is_empty() || format != Default::default() {
                        deleted_cells.push((row, col, raw, format));
                    }
                }
            }
        }

        // Capture column widths for deleted columns
        let deleted_col_widths: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col && **c < at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();

        // Remove widths for deleted columns and shift remaining left
        let widths_to_shift: Vec<_> = self.col_widths
            .iter()
            .filter(|(c, _)| **c >= at_col + count)
            .map(|(c, w)| (*c, *w))
            .collect();
        // Remove all affected widths
        for c in at_col..NUM_COLS {
            self.col_widths.remove(&c);
        }
        // Re-insert shifted widths
        for (c, w) in widths_to_shift {
            self.col_widths.insert(c - count, w);
        }

        // Perform the delete
        if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
            sheet.delete_cols(at_col, count);
        }

        // Record undo entry
        self.history.record_named_range_action(crate::history::UndoAction::ColsDeleted {
            sheet_index,
            at_col,
            count,
            deleted_cells,
            deleted_col_widths,
        });

        // Move selection left if needed
        if self.view_state.selected.1 >= at_col + count {
            self.view_state.selected.1 -= count;
        } else if self.view_state.selected.1 >= at_col {
            self.view_state.selected.1 = at_col.saturating_sub(1);
        }
        self.view_state.selection_end = None;

        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(format!("Deleted {} column(s)", count));
        cx.notify();
    }

    // Go To cell dialog
    pub fn show_goto(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::GoTo;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn hide_goto(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn confirm_goto(&mut self, cx: &mut Context<Self>) {
        if let Some((row, col)) = Self::parse_cell_ref(&self.goto_input) {
            if row < NUM_ROWS && col < NUM_COLS {
                self.view_state.selected = (row, col);
                self.view_state.selection_end = None;
                self.ensure_visible(cx);
                self.status_message = Some(format!("Jumped to {}", self.cell_ref()));
            } else {
                self.status_message = Some("Cell reference out of range".to_string());
            }
        } else {
            self.status_message = Some("Invalid cell reference".to_string());
        }
        self.mode = Mode::Navigation;
        self.goto_input.clear();
        cx.notify();
    }

    pub fn goto_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == Mode::GoTo {
            self.goto_input.push(c.to_ascii_uppercase());
            cx.notify();
        }
    }

    pub fn goto_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::GoTo {
            self.goto_input.pop();
            cx.notify();
        }
    }

    /// Parse cell reference like "A1", "B25", "AA100"
    pub(crate) fn parse_cell_ref(input: &str) -> Option<(usize, usize)> {
        let input = input.trim().to_uppercase();
        if input.is_empty() {
            return None;
        }

        // Find where letters end and numbers begin
        let letter_end = input.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 || letter_end == input.len() {
            return None;
        }

        let letters = &input[..letter_end];
        let numbers = &input[letter_end..];

        // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
        let col = letters.chars().fold(0usize, |acc, c| {
            acc * 26 + (c as usize - 'A' as usize + 1)
        }) - 1;

        // Parse row (1-based to 0-based)
        let row = numbers.parse::<usize>().ok()?.checked_sub(1)?;

        Some((row, col))
    }

    /// Parse all cell references from a formula with deterministic color assignment.
    /// Returns FormulaRef entries sorted by text position, with first-seen refs getting unique colors.
    fn parse_formula_refs(formula: &str) -> Vec<FormulaRef> {
        if !formula.starts_with('=') && !formula.starts_with('+') {
            return Vec::new();
        }

        let tokens = tokenize_for_highlight(formula);
        // Collect raw refs with text ranges: (RefKey, start, end, text_range)
        let mut parsed_refs: Vec<(RefKey, (usize, usize), Option<(usize, usize)>, std::ops::Range<usize>)> = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let (range, token_type) = &tokens[i];

            if *token_type == TokenType::CellRef {
                // Convert char indices to byte indices for safe slicing
                let byte_start = char_to_byte(formula, range.start);
                let byte_end = char_to_byte(formula, range.end);
                let cell_text = &formula[byte_start..byte_end];
                // Strip any $ signs for absolute references
                let cell_text_clean: String = cell_text.chars().filter(|c| *c != '$').collect();

                if let Some(start_cell) = Self::parse_cell_ref(&cell_text_clean) {
                    // Check if next tokens form a range (: followed by CellRef)
                    if i + 2 < tokens.len() {
                        let (_, next_type) = &tokens[i + 1];
                        let (range2, next_next_type) = &tokens[i + 2];

                        if *next_type == TokenType::Colon && *next_next_type == TokenType::CellRef {
                            // Convert char indices to byte indices for safe slicing
                            let byte_start2 = char_to_byte(formula, range2.start);
                            let byte_end2 = char_to_byte(formula, range2.end);
                            let end_text = &formula[byte_start2..byte_end2];
                            let end_text_clean: String = end_text.chars().filter(|c| *c != '$').collect();

                            if let Some(end_cell) = Self::parse_cell_ref(&end_text_clean) {
                                // Normalize range to min/max for stable RefKey
                                let r1 = start_cell.0.min(end_cell.0);
                                let c1 = start_cell.1.min(end_cell.1);
                                let r2 = start_cell.0.max(end_cell.0);
                                let c2 = start_cell.1.max(end_cell.1);
                                let key = RefKey::Range { r1, c1, r2, c2 };
                                let text_range = range.start..range2.end;
                                parsed_refs.push((key, (r1, c1), Some((r2, c2)), text_range));
                                i += 3;  // Skip the whole range
                                continue;
                            }
                        }
                    }
                    // Single cell reference
                    let key = RefKey::Cell { row: start_cell.0, col: start_cell.1 };
                    parsed_refs.push((key, start_cell, None, range.clone()));
                }
            }
            i += 1;
        }

        // Sort by text position (left-to-right in formula) for deterministic color assignment
        parsed_refs.sort_by_key(|(_, _, _, text_range)| text_range.start);

        // Assign colors: first-seen order, deduplicate by RefKey (same ref = same color)
        let mut color_map: HashMap<RefKey, usize> = HashMap::new();
        let mut next_color = 0;

        parsed_refs.into_iter().map(|(key, start, end, text_range)| {
            let color_index = *color_map.entry(key.clone()).or_insert_with(|| {
                let c = next_color;
                next_color = (next_color + 1) % 8;
                c
            });
            FormulaRef { key, start, end, color_index, text_range }
        }).collect()
    }

    // =========================================================================
    // Preferences panel
    // =========================================================================

    pub fn show_preferences(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::Preferences;
        cx.notify();
    }

    pub fn hide_preferences(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn theme_picker_up(&mut self, cx: &mut Context<Self>) {
        if self.theme_picker_selected > 0 {
            self.theme_picker_selected -= 1;
            self.update_theme_preview(cx);
        }
    }

    pub fn theme_picker_down(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if self.theme_picker_selected + 1 < filtered.len() {
            self.theme_picker_selected += 1;
            self.update_theme_preview(cx);
        }
    }

    pub fn theme_picker_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.theme_picker_query.push(c);
        self.theme_picker_selected = 0;
        self.update_theme_preview(cx);
    }

    pub fn theme_picker_backspace(&mut self, cx: &mut Context<Self>) {
        self.theme_picker_query.pop();
        self.theme_picker_selected = 0;
        self.update_theme_preview(cx);
    }

    pub fn theme_picker_execute(&mut self, cx: &mut Context<Self>) {
        self.apply_theme_at_index(self.theme_picker_selected, cx);
    }

    pub fn apply_theme_at_index(&mut self, index: usize, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if let Some(theme) = filtered.get(index) {
            self.theme = theme.clone();
            self.status_message = Some(format!("Applied theme: {}", theme.meta.name));
            // Persist theme selection to global store
            let theme_id = theme.meta.id.to_string();
            update_user_settings(cx, |settings| {
                settings.appearance.theme_id = Setting::Value(theme_id);
            });
        }
        self.theme_preview = None;
        self.mode = Mode::Navigation;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        cx.notify();
    }

    /// Filter available themes by query
    pub fn filter_themes(&self) -> Vec<Theme> {
        let themes = builtin_themes();
        if self.theme_picker_query.is_empty() {
            return themes;
        }
        let query_lower = self.theme_picker_query.to_lowercase();
        themes
            .into_iter()
            .filter(|t| t.meta.name.to_lowercase().contains(&query_lower))
            .collect()
    }

    /// Update theme preview based on current selection
    fn update_theme_preview(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_themes();
        if let Some(theme) = filtered.get(self.theme_picker_selected) {
            self.theme_preview = Some(theme.clone());
        } else {
            self.theme_preview = None;
        }
        cx.notify();
    }

    // About dialog
    pub fn show_about(&mut self, cx: &mut Context<Self>) {
        // Close console if open (About dialog needs focus)
        self.lua_console.visible = false;
        self.mode = Mode::About;
        cx.notify();
    }

    pub fn hide_about(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    // License dialog
    pub fn show_license(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.license_input.clear();
        self.license_error = None;
        self.mode = Mode::License;
        cx.notify();
    }

    pub fn hide_license(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.license_input.clear();
        self.license_error = None;
        cx.notify();
    }

    pub fn license_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.license_input.push(c);
        self.license_error = None;
        cx.notify();
    }

    pub fn license_backspace(&mut self, cx: &mut Context<Self>) {
        self.license_input.pop();
        self.license_error = None;
        cx.notify();
    }

    pub fn apply_license(&mut self, cx: &mut Context<Self>) {
        use crate::views::license_dialog::user_friendly_error;

        match visigrid_license::load_license(&self.license_input) {
            Ok(validation) => {
                if validation.valid {
                    self.status_message = Some(format!(
                        "License activated: {}",
                        visigrid_license::license_summary()
                    ));
                    self.hide_license(cx);
                } else {
                    // Convert technical error to user-friendly message
                    let raw_error = validation.error.as_deref().unwrap_or("Unknown error");
                    self.license_error = Some(user_friendly_error(raw_error));
                    cx.notify();
                }
            }
            Err(e) => {
                // Convert technical error to user-friendly message
                self.license_error = Some(user_friendly_error(&e));
                cx.notify();
            }
        }
    }

    pub fn clear_license(&mut self, cx: &mut Context<Self>) {
        visigrid_license::clear_license();
        self.status_message = Some("License removed".to_string());
        self.hide_license(cx);
    }

    // Import report dialog
    pub fn show_import_report(&mut self, cx: &mut Context<Self>) {
        if self.import_result.is_some() {
            self.lua_console.visible = false;
            self.mode = Mode::ImportReport;
            cx.notify();
        }
    }

    pub fn hide_import_report(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    // Export report dialog
    pub fn show_export_report(&mut self, cx: &mut Context<Self>) {
        if self.export_result.is_some() {
            self.lua_console.visible = false;
            self.mode = Mode::ExportReport;
            cx.notify();
        }
    }

    pub fn hide_export_report(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Dismiss the import overlay (ESC during background import)
    /// Does NOT cancel the import - just hides the overlay
    pub fn dismiss_import_overlay(&mut self, cx: &mut Context<Self>) {
        self.import_overlay_visible = false;
        cx.notify();
    }

    // Inspector panel methods
    pub fn toggle_inspector_pin(&mut self, cx: &mut Context<Self>) {
        if self.inspector_pinned.is_some() {
            // Unpin: follow selection again
            self.inspector_pinned = None;
        } else {
            // Pin: lock to current selection
            self.inspector_pinned = Some(self.view_state.selected);
        }
        cx.notify();
    }

    /// Set a path trace from clicked input/output to the inspected cell.
    /// `from` is the clicked cell, `to` is the inspected cell.
    /// `forward` is true when tracing from input toward inspected cell.
    pub fn set_trace_path(&mut self, from_row: usize, from_col: usize, to_row: usize, to_col: usize, forward: bool, cx: &mut Context<Self>) {
        use visigrid_engine::cell_id::CellId;

        let sheet_id = self.sheet().id;
        let from = CellId::new(sheet_id, from_row, from_col);
        let to = CellId::new(sheet_id, to_row, to_col);

        let result = self.workbook.find_path(from, to, forward);

        if result.path.is_empty() {
            // No path found - clear any existing trace
            self.inspector_trace_path = None;
            self.inspector_trace_incomplete = result.truncated;
            if result.truncated {
                self.status_message = Some("Trace too large — refine by starting closer".to_string());
            }
        } else {
            self.inspector_trace_path = Some(result.path);
            self.inspector_trace_incomplete = result.has_dynamic_refs || result.truncated;
        }
        cx.notify();
    }

    /// Clear the current trace path.
    pub fn clear_trace_path(&mut self, cx: &mut Context<Self>) {
        if self.inspector_trace_path.is_some() {
            self.inspector_trace_path = None;
            self.inspector_trace_incomplete = false;
            cx.notify();
        }
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
    // Default App Prompt (macOS title bar chip)
    // =========================================================================

    /// Get the extension key for the current file (for per-extension state).
    fn get_prompt_ext_key(&self) -> Option<String> {
        use crate::default_app::SpreadsheetFileType;

        // Use cached file type if available
        if let Some(ft) = self.default_app_prompt_file_type {
            return Some(match ft {
                SpreadsheetFileType::Excel => "xlsx",
                SpreadsheetFileType::Csv => "csv",
                SpreadsheetFileType::Tsv => "tsv",
                SpreadsheetFileType::Native => "vgrid",
            }.to_string());
        }

        // Derive from current file
        let path = self.document_meta.path.as_ref()?;
        let ext = path.extension().and_then(|e| e.to_str())?;
        let file_type = SpreadsheetFileType::from_ext(ext)?;
        Some(match file_type {
            SpreadsheetFileType::Excel => "xlsx",
            SpreadsheetFileType::Csv => "csv",
            SpreadsheetFileType::Tsv => "tsv",
            SpreadsheetFileType::Native => "vgrid",
        }.to_string())
    }

    /// Check if the default app prompt should be shown.
    ///
    /// Returns true when ALL conditions are met:
    /// - macOS only
    /// - File successfully loaded (has path, no import errors showing)
    /// - Not a temporary file
    /// - File type is CSV/TSV/Excel (not native .vgrid)
    /// - User hasn't dismissed the prompt for THIS extension
    /// - Not in cool-down period for THIS extension (7 days after ignoring)
    /// - Not already shown this session
    /// - We haven't already marked this extension as completed
    /// - VisiGrid isn't already the default for this file type
    #[cfg(target_os = "macos")]
    pub fn should_show_default_app_prompt(&self, cx: &gpui::App) -> bool {
        use crate::default_app::{SpreadsheetFileType, is_default_handler, is_temporary_file, shown_this_session};

        // If we're showing success/needs-settings feedback, show that instead
        if self.default_app_prompt_state == DefaultAppPromptState::Success
            || self.default_app_prompt_state == DefaultAppPromptState::NeedsSettings
        {
            return true;
        }

        // Must have a file open
        let path = match &self.document_meta.path {
            Some(p) => p,
            None => return false,
        };

        // Skip if import report dialog is showing (don't prompt during error review)
        if self.mode == Mode::ImportReport {
            return false;
        }

        // Skip unsaved files (new documents)
        if !self.document_meta.is_saved && self.document_meta.source.is_none() {
            return false;
        }

        // Skip temporary files
        if is_temporary_file(path) {
            return false;
        }

        // Get file type from extension
        let ext = path.extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let file_type = match SpreadsheetFileType::from_ext(ext) {
            Some(ft) => ft,
            None => return false,
        };

        // Skip native VisiGrid files
        if file_type == SpreadsheetFileType::Native {
            return false;
        }

        // Get extension key for per-extension state
        let ext_key = match file_type {
            SpreadsheetFileType::Excel => "xlsx",
            SpreadsheetFileType::Csv => "csv",
            SpreadsheetFileType::Tsv => "tsv",
            SpreadsheetFileType::Native => return false,
        };

        let settings = user_settings(cx);

        // Check if user has permanently dismissed for THIS extension
        if settings.is_default_app_prompt_dismissed(ext_key) {
            return false;
        }

        // Check if we've already completed setup for this extension
        if settings.is_default_app_prompt_completed(ext_key) {
            return false;
        }

        // Check cool-down period for THIS extension (7 days after ignoring)
        if settings.is_default_app_prompt_in_cooldown(ext_key) {
            return false;
        }

        // Don't spam within same session
        if shown_this_session() {
            return false;
        }

        // Check if VisiGrid is already the default (do last - can be slow)
        if is_default_handler(file_type) {
            return false;
        }

        true
    }

    #[cfg(not(target_os = "macos"))]
    pub fn should_show_default_app_prompt(&self, _cx: &gpui::App) -> bool {
        false
    }

    /// Get the file type for the current prompt (for display).
    #[cfg(target_os = "macos")]
    pub fn get_prompt_file_type(&self) -> Option<crate::default_app::SpreadsheetFileType> {
        use crate::default_app::SpreadsheetFileType;

        // If we have a cached file type from when we showed the prompt, use that
        if let Some(ft) = self.default_app_prompt_file_type {
            return Some(ft);
        }

        // Otherwise derive from current file
        let path = self.document_meta.path.as_ref()?;
        let ext = path.extension().and_then(|e| e.to_str())?;
        SpreadsheetFileType::from_ext(ext)
    }

    #[cfg(not(target_os = "macos"))]
    pub fn get_prompt_file_type(&self) -> Option<crate::default_app::SpreadsheetFileType> {
        None
    }

    /// Called when the prompt becomes visible - marks session and records timestamp.
    pub fn on_default_app_prompt_shown(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::{mark_shown_this_session, SpreadsheetFileType};

        // Cache the file type
        if let Some(path) = &self.document_meta.path {
            if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
                self.default_app_prompt_file_type = SpreadsheetFileType::from_ext(ext);
            }
        }

        // Mark shown this session
        mark_shown_this_session();

        // Record timestamp for cool-down for THIS extension (in case they ignore)
        if let Some(ext_key) = self.get_prompt_ext_key() {
            update_user_settings(cx, |settings| {
                settings.mark_default_app_prompt_shown(&ext_key);
            });
        }

        self.default_app_prompt_state = DefaultAppPromptState::Showing;
    }

    /// Set VisiGrid as the default handler for the current file type.
    #[cfg(target_os = "macos")]
    pub fn set_as_default_app(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::{set_as_default_handler, is_default_handler};

        let file_type = match self.get_prompt_file_type() {
            Some(ft) => ft,
            None => return,
        };

        let ext_key = self.get_prompt_ext_key();

        match set_as_default_handler(file_type) {
            Ok(()) => {
                // Check if it actually worked (duti succeeded)
                if is_default_handler(file_type) {
                    // Success! Show brief confirmation
                    self.default_app_prompt_state = DefaultAppPromptState::Success;
                    self.default_app_prompt_success_timer = Some(std::time::Instant::now());
                    // Mark completed for this extension (permanent)
                    if let Some(ext) = ext_key {
                        update_user_settings(cx, |settings| {
                            settings.mark_default_app_completed(&ext);
                        });
                    }
                } else {
                    // Needs manual completion in Settings
                    // Note: Do NOT permanently dismiss - just cool-down
                    // The cool-down timestamp is already set from when we showed the prompt
                    self.default_app_prompt_state = DefaultAppPromptState::NeedsSettings;
                    self.needs_settings_entered_at = Some(std::time::Instant::now());
                    self.needs_settings_check_count = 0;
                }
            }
            Err(_) => {
                // Failed - needs Settings
                // Note: Do NOT permanently dismiss - just cool-down
                self.default_app_prompt_state = DefaultAppPromptState::NeedsSettings;
                self.needs_settings_entered_at = Some(std::time::Instant::now());
                self.needs_settings_check_count = 0;
            }
        }

        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn set_as_default_app(&mut self, _cx: &mut Context<Self>) {}

    /// Open System Settings to complete the default app setup.
    #[cfg(target_os = "macos")]
    pub fn open_default_app_settings(&mut self, cx: &mut Context<Self>) {
        use std::process::Command;

        let _ = Command::new("open")
            .args(["x-apple.systempreferences:com.apple.ExtensionsPreferences"])
            .spawn();

        // Keep in NeedsSettings state (don't hide) so we can re-check on focus
        // The prompt will be re-checked when the window regains focus
        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn open_default_app_settings(&mut self, _cx: &mut Context<Self>) {}

    /// Dismiss the default app prompt permanently for this extension (user clicked ✕).
    pub fn dismiss_default_app_prompt(&mut self, cx: &mut Context<Self>) {
        if let Some(ext_key) = self.get_prompt_ext_key() {
            update_user_settings(cx, |settings| {
                settings.dismiss_default_app_prompt(&ext_key);
            });
        }
        self.default_app_prompt_state = DefaultAppPromptState::Hidden;
        self.needs_settings_entered_at = None;
        self.needs_settings_check_count = 0;
        cx.notify();
    }

    /// Re-check default handler after returning from System Settings.
    /// Called when window regains focus while in NeedsSettings state.
    /// Note: This is now mostly handled by check_default_app_prompt_timer(),
    /// but we keep this for explicit calls if needed.
    #[cfg(target_os = "macos")]
    pub fn recheck_default_app_handler(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::is_default_handler;

        // Only re-check if we're in NeedsSettings state
        if self.default_app_prompt_state != DefaultAppPromptState::NeedsSettings {
            return;
        }

        let file_type = match self.get_prompt_file_type() {
            Some(ft) => ft,
            None => {
                self.default_app_prompt_state = DefaultAppPromptState::Hidden;
                self.needs_settings_entered_at = None;
                self.needs_settings_check_count = 0;
                cx.notify();
                return;
            }
        };

        // Check if they completed the setup in Settings
        if is_default_handler(file_type) {
            // Success! Show brief confirmation
            self.default_app_prompt_state = DefaultAppPromptState::Success;
            self.default_app_prompt_success_timer = Some(std::time::Instant::now());
            self.needs_settings_entered_at = None;
            self.needs_settings_check_count = 0;

            // Mark completed for this extension
            if let Some(ext_key) = self.get_prompt_ext_key() {
                update_user_settings(cx, |settings| {
                    settings.mark_default_app_completed(&ext_key);
                });
            }
        } else {
            // Still not default - hide for now (cool-down will handle re-show)
            self.default_app_prompt_state = DefaultAppPromptState::Hidden;
            self.needs_settings_entered_at = None;
            self.needs_settings_check_count = 0;
        }

        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn recheck_default_app_handler(&mut self, _cx: &mut Context<Self>) {}

    /// Check if success timer has expired and hide the prompt.
    /// Also handles NeedsSettings state with exponential backoff:
    /// - Check 1: at 3 seconds
    /// - Check 2: at 8 seconds
    /// - Check 3: at 20 seconds
    /// - Then stop polling (rely on next file open or next session)
    pub fn check_default_app_prompt_timer(&mut self, cx: &mut Context<Self>) {
        if self.default_app_prompt_state == DefaultAppPromptState::Success {
            if let Some(started) = self.default_app_prompt_success_timer {
                if started.elapsed() > std::time::Duration::from_secs(2) {
                    self.default_app_prompt_state = DefaultAppPromptState::Hidden;
                    self.default_app_prompt_success_timer = None;
                    cx.notify();
                }
            }
        } else if self.default_app_prompt_state == DefaultAppPromptState::NeedsSettings {
            #[cfg(target_os = "macos")]
            {
                use crate::default_app::is_default_handler;
                use std::time::Instant;

                // Exponential backoff schedule (seconds since entered_at):
                // Check 0: 3s, Check 1: 8s, Check 2: 20s, then stop
                const CHECK_SCHEDULE: [u64; 3] = [3, 8, 20];

                let now = Instant::now();
                let entered_at = match self.needs_settings_entered_at {
                    Some(t) => t,
                    None => return, // No timestamp means we shouldn't be polling
                };

                let elapsed_secs = now.duration_since(entered_at).as_secs();
                let check_count = self.needs_settings_check_count as usize;

                // Already exhausted all checks? Stop polling.
                if check_count >= CHECK_SCHEDULE.len() {
                    return;
                }

                // Not yet time for the next check?
                let next_check_at = CHECK_SCHEDULE[check_count];
                if elapsed_secs < next_check_at {
                    return;
                }

                // Time for a check - increment counter first
                self.needs_settings_check_count += 1;

                // Re-check handler status
                if let Some(file_type) = self.get_prompt_file_type() {
                    if is_default_handler(file_type) {
                        // User completed setup! Show success briefly
                        self.default_app_prompt_state = DefaultAppPromptState::Success;
                        self.default_app_prompt_success_timer = Some(now);
                        self.needs_settings_entered_at = None;
                        self.needs_settings_check_count = 0;

                        // Mark completed for this extension
                        if let Some(ext_key) = self.get_prompt_ext_key() {
                            update_user_settings(cx, |settings| {
                                settings.mark_default_app_completed(&ext_key);
                            });
                        }

                        cx.notify();
                    }
                    // If not default yet, keep chip visible but stop polling after 3 checks
                }
            }
        }
    }

    // =========================================================================
    // Impact Preview methods
    // =========================================================================

    /// Find all cells that reference a named range
    fn find_named_range_usages(&self, name: &str) -> Vec<crate::views::impact_preview::ImpactedFormula> {
        use crate::views::impact_preview::ImpactedFormula;

        let name_upper = name.to_uppercase();
        let mut usages = Vec::new();

        // Scan all cells for formulas containing the name
        for ((row, col), cell) in self.sheet().cells_iter() {
            let raw = cell.value.raw_display();
            if !raw.starts_with('=') {
                continue;
            }

            let formula_upper = raw.to_uppercase();

            // Check if name appears as a standalone identifier
            let contains_name = formula_upper
                .split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.')
                .any(|word| word == name_upper);

            if contains_name {
                // Format cell reference
                let cell_ref = {
                    let mut col_name = String::new();
                    let mut c = *col;
                    loop {
                        col_name.insert(0, (b'A' + (c % 26) as u8) as char);
                        if c < 26 { break; }
                        c = c / 26 - 1;
                    }
                    format!("{}{}", col_name, *row + 1)
                };

                usages.push(ImpactedFormula {
                    cell_ref,
                    formula: raw.to_string(),
                });
            }
        }

        // Sort by cell reference for consistent display
        usages.sort_by(|a, b| a.cell_ref.cmp(&b.cell_ref));
        usages
    }

    /// Show impact preview for a rename operation
    pub fn show_impact_preview_for_rename(&mut self, old_name: &str, new_name: &str, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let usages = self.find_named_range_usages(old_name);
        self.impact_preview_action = Some(ImpactAction::Rename {
            old_name: old_name.to_string(),
            new_name: new_name.to_string(),
        });
        self.impact_preview_usages = usages;
        self.mode = Mode::ImpactPreview;
        cx.notify();
    }

    /// Show impact preview for a delete operation
    pub fn show_impact_preview_for_delete(&mut self, name: &str, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let usages = self.find_named_range_usages(name);
        self.impact_preview_action = Some(ImpactAction::Delete {
            name: name.to_string(),
        });
        self.impact_preview_usages = usages;
        self.mode = Mode::ImpactPreview;
        cx.notify();
    }

    /// Hide the impact preview modal
    pub fn hide_impact_preview(&mut self, cx: &mut Context<Self>) {
        self.impact_preview_action = None;
        self.impact_preview_usages.clear();
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Apply the previewed action (rename or delete)
    pub fn apply_impact_preview(&mut self, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let action = self.impact_preview_action.take();
        let usage_count = self.impact_preview_usages.len();
        self.impact_preview_usages.clear();

        match action {
            Some(ImpactAction::Rename { old_name, new_name }) => {
                // Perform the rename
                self.apply_rename_internal(&old_name, &new_name, cx);
                self.mode = Mode::Navigation;

                // Show one-time F12 hint after first rename
                if !user_settings(cx).is_tip_dismissed(TipId::RenameF12) {
                    update_user_settings(cx, |settings| {
                        settings.dismiss_tip(TipId::RenameF12);
                    });
                    self.status_message = Some(format!(
                        "Renamed \"{}\" → \"{}\". Tip: Press F12 to jump to this name's definition.",
                        old_name, new_name
                    ));
                } else {
                    self.status_message = Some(if usage_count > 0 {
                        format!("Renamed \"{}\" → \"{}\", updated {} formula{}",
                            old_name, new_name, usage_count, if usage_count == 1 { "" } else { "s" })
                    } else {
                        format!("Renamed \"{}\" → \"{}\"", old_name, new_name)
                    });
                }
            }
            Some(ImpactAction::Delete { name }) => {
                // Perform the delete
                self.delete_named_range_internal(&name, usage_count, cx);
                self.mode = Mode::Navigation;
                self.status_message = Some(if usage_count > 0 {
                    format!("Deleted \"{}\", {} formula{} affected",
                        name, usage_count, if usage_count == 1 { "" } else { "s" })
                } else {
                    format!("Deleted \"{}\"", name)
                });
            }
            None => {
                self.mode = Mode::Navigation;
            }
        }
        cx.notify();
    }

    // =========================================================================
    // Refactor Log methods
    // =========================================================================

    /// Show the refactor log modal
    pub fn show_refactor_log(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::RefactorLog;
        cx.notify();
    }

    /// Hide the refactor log modal
    pub fn hide_refactor_log(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Log a refactor action
    pub fn log_refactor(&mut self, action: &str, details: &str, impact: Option<&str>) {
        use crate::views::refactor_log::RefactorLogEntry;

        let mut entry = RefactorLogEntry::new(action, details);
        if let Some(imp) = impact {
            entry = entry.with_impact(imp);
        }
        self.refactor_log.push(entry);
    }

    // ========================================================================
    // Keyboard hints (Vimium-style jump navigation)
    // ========================================================================

    /// Enter keyboard hint mode - show jump labels on visible cells.
    pub fn enter_hint_mode(&mut self, cx: &mut Context<Self>) {
        self.enter_hint_mode_with_labels(true, cx);
    }

    /// Enter hint/command mode with optional labels.
    ///
    /// - `show_labels`: if true, generate cell labels (full hint mode)
    /// - `show_labels`: if false, command-only mode (for vim gg without labels)
    pub fn enter_hint_mode_with_labels(&mut self, show_labels: bool, cx: &mut Context<Self>) {
        self.hint_state.buffer.clear();

        if show_labels {
            // Full hint mode: generate labels for visible cells
            let visible_rows = self.visible_rows();
            let visible_cols = self.visible_cols();
            self.hint_state.labels = crate::hints::generate_hints(
                self.view_state.scroll_row,
                self.view_state.scroll_col,
                visible_rows,
                visible_cols,
            );
            self.hint_state.viewport = (self.view_state.scroll_row, self.view_state.scroll_col, visible_rows, visible_cols);
            self.status_message = Some("Hint: type letters to jump".into());
        } else {
            // Command-only mode: no labels, just waiting for g-commands (gg, etc.)
            self.hint_state.labels.clear();
            self.hint_state.viewport = (0, 0, 0, 0);
            self.status_message = Some("g-".into());
        }

        self.mode = Mode::Hint;
        cx.notify();
    }

    /// Exit keyboard hint mode without jumping.
    pub fn exit_hint_mode(&mut self, cx: &mut Context<Self>) {
        self.hint_state.clear();
        self.mode = Mode::Navigation;
        self.status_message = None;
        cx.notify();
    }

    /// Handle a key press in hint mode.
    /// Returns true if the key was consumed.
    ///
    /// Uses the resolver architecture from hints.rs:
    /// 1. Exact command match (gg → GotoTop)
    /// 2. Cell label resolution (a, ab, zz)
    /// 3. No match → exit
    pub fn apply_hint_key(&mut self, key: &str, cx: &mut Context<Self>) -> bool {
        use crate::hints::{resolve_hint_buffer, HintResolution, GCommand, HintExitReason};

        match key {
            "escape" => {
                self.hint_state.last_exit_reason = Some(HintExitReason::Cancelled);
                self.exit_hint_mode(cx);
                true
            }
            "backspace" => {
                self.hint_state.buffer.pop();
                self.update_hint_status(cx);
                true
            }
            _ if key.len() == 1 && key.chars().next().map(|c| c.is_ascii_lowercase()).unwrap_or(false) => {
                self.hint_state.buffer.push_str(key);

                // Resolve the buffer through the phase system
                match resolve_hint_buffer(&self.hint_state) {
                    HintResolution::Command(cmd) => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::Command);
                        self.execute_g_command(cmd, cx);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::Jump(row, col) => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::LabelJump);
                        self.view_state.selected = (row, col);
                        self.view_state.selection_end = None;
                        self.view_state.additional_selections.clear();
                        self.ensure_cell_visible(row, col);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::NoMatch => {
                        self.hint_state.last_exit_reason = Some(HintExitReason::NoMatch);
                        self.exit_hint_mode(cx);
                    }
                    HintResolution::Pending => {
                        self.update_hint_status(cx);
                    }
                }
                true
            }
            _ => false, // Unhandled key
        }
    }

    /// Execute a g-prefixed command.
    fn execute_g_command(&mut self, cmd: crate::hints::GCommand, cx: &mut Context<Self>) {
        use crate::hints::GCommand;

        match cmd {
            GCommand::GotoTop => {
                // gg - Go to A1
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.additional_selections.clear();
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                self.status_message = Some("Jumped to A1".into());
                cx.notify();
            }
            // Future commands go here
        }
    }

    /// Update status bar with current hint state.
    fn update_hint_status(&mut self, cx: &mut Context<Self>) {
        let matches = self.hint_state.matching_labels();
        let buffer = &self.hint_state.buffer;

        if buffer.is_empty() {
            self.status_message = Some("Hint: type letters to jump".into());
        } else if matches.is_empty() {
            self.status_message = Some(format!("Hint: {} (no matches)", buffer));
        } else if matches.len() == 1 {
            // This shouldn't happen (we auto-jump on unique match), but handle it
            self.status_message = Some(format!("Hint: {} → jumping", buffer));
        } else {
            self.status_message = Some(format!("Hint: {} ({} matches)", buffer, matches.len()));
        }
        cx.notify();
    }

    /// Check if hints are enabled in settings.
    pub fn keyboard_hints_enabled(&self, cx: &Context<Self>) -> bool {
        use crate::settings::user_settings;
        user_settings(cx)
            .navigation
            .keyboard_hints
            .as_value()
            .copied()
            .unwrap_or(false)
    }

    /// Check if vim mode is enabled in settings.
    pub fn vim_mode_enabled(&self, cx: &Context<Self>) -> bool {
        use crate::settings::user_settings;
        user_settings(cx)
            .navigation
            .vim_mode
            .as_value()
            .copied()
            .unwrap_or(false)
    }

    /// Handle vim-style navigation keys.
    /// Returns true if the key was consumed.
    pub fn apply_vim_key(&mut self, key: &str, cx: &mut Context<Self>) -> bool {
        match key {
            "h" => {
                self.move_selection(0, -1, cx);
                true
            }
            "j" => {
                self.move_selection(1, 0, cx);
                true
            }
            "k" => {
                self.move_selection(-1, 0, cx);
                true
            }
            "l" => {
                self.move_selection(0, 1, cx);
                true
            }
            "i" => {
                // Enter edit mode (like F2 - edit without replacing)
                self.start_edit(cx);
                true
            }
            "0" => {
                // Move to first column
                self.view_state.selected = (self.view_state.selected.0, 0);
                self.view_state.selection_end = None;
                self.ensure_cell_visible(self.view_state.selected.0, 0);
                cx.notify();
                true
            }
            "$" => {
                // Move to last column with data in current row (or last visible)
                let row = self.view_state.selected.0;
                let last_col = self.find_last_data_col_in_row(row);
                self.view_state.selected = (row, last_col);
                self.view_state.selection_end = None;
                self.ensure_cell_visible(row, last_col);
                cx.notify();
                true
            }
            "w" => {
                // Forward jump: Ctrl+Right equivalent
                self.jump_selection(0, 1, cx);
                true
            }
            "b" => {
                // Back jump: Ctrl+Left equivalent
                self.jump_selection(0, -1, cx);
                true
            }
            _ => false,
        }
    }

    /// Find the last column with data in a given row.
    fn find_last_data_col_in_row(&self, row: usize) -> usize {
        let sheet = self.workbook.active_sheet();
        for col in (0..NUM_COLS).rev() {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                return col;
            }
        }
        0 // Default to first column if row is empty
    }
}

/// Convert column index to letter(s) (0 = A, 25 = Z, 26 = AA, etc.)
pub(crate) fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

impl Render for Spreadsheet {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        // One-shot title refresh (triggered by async operations without window access)
        if self.pending_title_refresh {
            self.pending_title_refresh = false;
            self.update_title_if_needed(window);
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

        // Cache window bounds for session snapshot (updated each render)
        self.cached_window_bounds = Some(window.window_bounds());

        // Update grid layout cache for hit-testing
        let menu_height = if cfg!(target_os = "macos") { 0.0 } else { MENU_BAR_HEIGHT };
        let formula_bar_height = FORMULA_BAR_HEIGHT;
        let col_header_height = COLUMN_HEADER_HEIGHT;

        let grid_body_y = menu_height + formula_bar_height + col_header_height;
        let grid_body_x = HEADER_WIDTH;

        let window_height: f32 = current_size.height.into();
        let window_width: f32 = current_size.width.into();

        // Account for side panels and status bar
        let right_panel_width = if self.inspector_visible {
            crate::views::inspector_panel::PANEL_WIDTH
        } else {
            0.0
        };
        let bottom_status_height = STATUS_BAR_HEIGHT;

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
        self.formula_bar_text_rect = gpui::Bounds {
            origin: gpui::point(gpui::px(FORMULA_BAR_TEXT_LEFT), gpui::px(menu_height)),
            size: gpui::size(gpui::px(formula_bar_text_width), gpui::px(formula_bar_height)),
        };

        // Update formula bar display cache (only when not editing)
        // This avoids re-parsing on every render
        if !self.mode.is_editing() {
            let cell = self.view_state.selected;
            let formula = self.sheet().get_raw(cell.0, cell.1);

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
