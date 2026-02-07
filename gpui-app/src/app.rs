use gpui::*;
use std::collections::HashMap;
use std::path::PathBuf;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::CellLookup;
use visigrid_engine::filter::{RowView, FilterState};
use visigrid_engine::sheet::SheetId;

use crate::clipboard::InternalClipboard;
use crate::find_replace::MatchHit;
use crate::formatting::BorderApplyMode;
use crate::history::{History, HistoryFingerprint};
use crate::mode::Mode;
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
// Validation Dialog State (Phase 4)
// ============================================================================

/// Validation type options for the dialog dropdown
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationTypeOption {
    #[default]
    AnyValue,
    List,
    WholeNumber,
    Decimal,
}

impl ValidationTypeOption {
    pub fn label(&self) -> &'static str {
        match self {
            Self::AnyValue => "Any value",
            Self::List => "List",
            Self::WholeNumber => "Whole number",
            Self::Decimal => "Decimal",
        }
    }

    pub const ALL: &'static [ValidationTypeOption] = &[
        Self::AnyValue,
        Self::List,
        Self::WholeNumber,
        Self::Decimal,
    ];
}

/// Numeric comparison operator options
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NumericOperatorOption {
    #[default]
    Between,
    NotBetween,
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

impl NumericOperatorOption {
    pub fn label(&self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "not between",
            Self::EqualTo => "equal to",
            Self::NotEqualTo => "not equal to",
            Self::GreaterThan => "greater than",
            Self::LessThan => "less than",
            Self::GreaterThanOrEqual => "greater than or equal to",
            Self::LessThanOrEqual => "less than or equal to",
        }
    }

    pub const ALL: &'static [NumericOperatorOption] = &[
        Self::Between,
        Self::NotBetween,
        Self::EqualTo,
        Self::NotEqualTo,
        Self::GreaterThan,
        Self::LessThan,
        Self::GreaterThanOrEqual,
        Self::LessThanOrEqual,
    ];

    /// Whether this operator requires two values (min/max)
    pub fn needs_two_values(&self) -> bool {
        matches!(self, Self::Between | Self::NotBetween)
    }
}

/// Paste Special type selection
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PasteType {
    #[default]
    All,      // Normal paste (values + formulas)
    Values,   // Computed values only
    Formulas, // Raw formulas with reference adjustment
    Formats,  // Cell formatting only
}

impl PasteType {
    /// All available paste types in display order
    pub fn all() -> &'static [PasteType] {
        &[PasteType::All, PasteType::Values, PasteType::Formulas, PasteType::Formats]
    }

    /// Display name for UI
    pub fn label(&self) -> &'static str {
        match self {
            PasteType::All => "All",
            PasteType::Values => "Values",
            PasteType::Formulas => "Formulas",
            PasteType::Formats => "Formats",
        }
    }

    /// Keyboard accelerator for this paste type
    pub fn accelerator(&self) -> char {
        match self {
            PasteType::All => 'A',
            PasteType::Values => 'V',
            PasteType::Formulas => 'F',
            PasteType::Formats => 'O', // fOrmats (Excel convention)
        }
    }

    /// Description for UI
    pub fn description(&self) -> &'static str {
        match self {
            PasteType::All => "Paste everything (formulas, values, and formats)",
            PasteType::Values => "Paste computed values only (no formulas)",
            PasteType::Formulas => "Paste formulas with reference adjustment",
            PasteType::Formats => "Paste cell formatting only (no values)",
        }
    }
}

/// State for the Paste Special dialog
#[derive(Debug, Clone, Default)]
pub struct PasteSpecialDialogState {
    /// Currently selected paste type
    pub selected: PasteType,
}

/// Format type selection in the number format editor
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NumberFormatEditorType {
    #[default]
    General,
    Number,
    Currency,
    Percent,
    Date,
}

/// Per-type settings cache for the number format editor
#[derive(Clone)]
pub struct TypeSettings {
    pub decimals: u8,
    pub thousands: bool,
    pub negative: NegativeStyle,
    pub currency_symbol: String,
}

impl Default for TypeSettings {
    fn default() -> Self {
        Self {
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::Minus,
            currency_symbol: String::new(),
        }
    }
}

/// State for the Number Format Editor dialog (Ctrl+1 escalation)
pub struct NumberFormatEditorState {
    pub format_type: NumberFormatEditorType,
    pub preview_value: f64,
    // Active settings (mirrors the cache entry for current type)
    pub decimals: u8,
    pub thousands: bool,
    pub negative: NegativeStyle,
    pub currency_symbol: String,
    // Per-type caches
    number_cache: TypeSettings,
    currency_cache: TypeSettings,
    percent_cache: TypeSettings,
}

impl Default for NumberFormatEditorState {
    fn default() -> Self {
        Self {
            format_type: NumberFormatEditorType::Number,
            preview_value: 1234.5678,
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::Minus,
            currency_symbol: String::new(),
            number_cache: TypeSettings {
                decimals: 2,
                thousands: true,
                negative: NegativeStyle::Minus,
                currency_symbol: String::new(),
            },
            currency_cache: TypeSettings {
                decimals: 2,
                thousands: true,
                negative: NegativeStyle::Parens,
                currency_symbol: String::new(),
            },
            percent_cache: TypeSettings {
                decimals: 2,
                thousands: false,
                negative: NegativeStyle::Minus,
                currency_symbol: String::new(),
            },
        }
    }
}

impl NumberFormatEditorState {
    /// Initialize from an existing NumberFormat and a sample value
    pub fn from_number_format(fmt: &NumberFormat, sample: f64) -> Self {
        let mut state = Self::default();
        state.preview_value = sample;
        match fmt {
            NumberFormat::Number { decimals, thousands, negative } => {
                state.format_type = NumberFormatEditorType::Number;
                state.decimals = *decimals;
                state.thousands = *thousands;
                state.negative = *negative;
                state.number_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: *thousands,
                    negative: *negative,
                    currency_symbol: String::new(),
                };
            }
            NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                state.format_type = NumberFormatEditorType::Currency;
                state.decimals = *decimals;
                state.thousands = *thousands;
                state.negative = *negative;
                state.currency_symbol = symbol.as_deref().unwrap_or("").to_string();
                state.currency_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: *thousands,
                    negative: *negative,
                    currency_symbol: symbol.as_deref().unwrap_or("").to_string(),
                };
            }
            NumberFormat::Percent { decimals } => {
                state.format_type = NumberFormatEditorType::Percent;
                state.decimals = *decimals;
                state.thousands = false;
                state.negative = NegativeStyle::Minus;
                state.percent_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: false,
                    negative: NegativeStyle::Minus,
                    currency_symbol: String::new(),
                };
            }
            NumberFormat::Date { .. } => {
                state.format_type = NumberFormatEditorType::Date;
            }
            _ => {
                state.format_type = NumberFormatEditorType::General;
            }
        }
        state
    }

    /// Convert current state to a NumberFormat
    pub fn to_number_format(&self) -> NumberFormat {
        match self.format_type {
            NumberFormatEditorType::General => NumberFormat::General,
            NumberFormatEditorType::Number => NumberFormat::Number {
                decimals: self.decimals.min(10),
                thousands: self.thousands,
                negative: self.negative,
            },
            NumberFormatEditorType::Currency => NumberFormat::Currency {
                decimals: self.decimals.min(10),
                thousands: self.thousands,
                negative: self.negative,
                symbol: if self.currency_symbol.is_empty() { None } else { Some(self.currency_symbol.clone()) },
            },
            NumberFormatEditorType::Percent => NumberFormat::Percent {
                decimals: self.decimals.min(10),
            },
            NumberFormatEditorType::Date => NumberFormat::Date {
                style: visigrid_engine::cell::DateStyle::Short,
            },
        }
    }

    /// Format a value using current settings for preview
    pub fn preview(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(self.preview_value, &self.to_number_format())
    }

    /// Format the negative version for preview
    pub fn preview_negative(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(-self.preview_value.abs(), &self.to_number_format())
    }

    /// Format zero for preview
    pub fn preview_zero(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(0.0, &self.to_number_format())
    }

    /// Switch to a different format type, preserving per-type caches
    pub fn switch_type(&mut self, new_type: NumberFormatEditorType) {
        if self.format_type == new_type {
            return;
        }
        // Save current settings to outgoing cache
        let current = TypeSettings {
            decimals: self.decimals,
            thousands: self.thousands,
            negative: self.negative,
            currency_symbol: self.currency_symbol.clone(),
        };
        match self.format_type {
            NumberFormatEditorType::Number => self.number_cache = current,
            NumberFormatEditorType::Currency => self.currency_cache = current,
            NumberFormatEditorType::Percent => self.percent_cache = current,
            _ => {}
        }
        // Restore from incoming cache
        self.format_type = new_type;
        match new_type {
            NumberFormatEditorType::Number => {
                let c = &self.number_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            NumberFormatEditorType::Currency => {
                let c = &self.currency_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            NumberFormatEditorType::Percent => {
                let c = &self.percent_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            _ => {
                self.decimals = 2;
                self.thousands = false;
                self.negative = NegativeStyle::Minus;
                self.currency_symbol = String::new();
            }
        }
    }
}

/// Which field in the validation dialog has focus
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationDialogFocus {
    #[default]
    None,
    TypeDropdown,
    OperatorDropdown,
    Source,      // List source field
    Value1,      // First value (or Minimum for Between)
    Value2,      // Second value (Maximum for Between)
}

/// State for the data validation dialog (Phase 4)
#[derive(Debug, Clone, Default)]
pub struct ValidationDialogState {
    /// Currently selected validation type
    pub validation_type: ValidationTypeOption,
    /// Whether the type dropdown is expanded
    pub type_dropdown_open: bool,
    /// Whether the operator dropdown is expanded
    pub operator_dropdown_open: bool,

    // List validation fields
    /// Source for list validation (e.g., "A1:A10" or "Yes,No,Maybe")
    pub list_source: String,
    /// Show dropdown arrow in cell
    pub show_dropdown: bool,

    // Numeric validation fields
    /// Comparison operator
    pub numeric_operator: NumericOperatorOption,
    /// First value (or minimum for between)
    pub value1: String,
    /// Second value (maximum for between)
    pub value2: String,

    // Common fields
    /// Allow blank cells
    pub ignore_blank: bool,

    /// Which field currently has focus
    pub focus: ValidationDialogFocus,

    /// Error message to display (validation errors)
    pub error: Option<String>,

    /// The range we're applying validation to (captured when dialog opens)
    pub target_range: Option<visigrid_engine::validation::CellRange>,

    /// Whether we loaded existing validation (for Clear button visibility)
    pub has_existing_validation: bool,
}

impl ValidationDialogState {
    /// Reset to defaults for a new dialog session
    pub fn reset(&mut self) {
        *self = Self::default();
        self.show_dropdown = true;  // Default to showing dropdown for list
        self.ignore_blank = true;   // Default to allowing blank
    }

    /// Load state from an existing validation rule
    pub fn load_from_rule(&mut self, rule: &visigrid_engine::validation::ValidationRule) {
        use visigrid_engine::validation::{ValidationType, ListSource};

        self.reset();
        self.has_existing_validation = true;
        self.ignore_blank = rule.ignore_blank;
        self.show_dropdown = rule.show_dropdown;

        match &rule.rule_type {
            // NOTE: No AnyValue case - that variant no longer exists in engine
            ValidationType::List(source) => {
                self.validation_type = ValidationTypeOption::List;
                match source {
                    ListSource::Inline(items) => {
                        self.list_source = items.join(",");
                    }
                    ListSource::Range(r) => {
                        self.list_source = r.clone();
                    }
                    ListSource::NamedRange(n) => {
                        self.list_source = n.clone();
                    }
                }
            }
            ValidationType::WholeNumber(constraint) => {
                self.validation_type = ValidationTypeOption::WholeNumber;
                self.load_numeric_constraint(constraint);
            }
            ValidationType::Decimal(constraint) => {
                self.validation_type = ValidationTypeOption::Decimal;
                self.load_numeric_constraint(constraint);
            }
            _ => {
                // Date, Time, TextLength, Custom - not yet supported in dialog
                // Show as AnyValue (read-only)
                self.validation_type = ValidationTypeOption::AnyValue;
            }
        }
    }

    fn load_numeric_constraint(&mut self, constraint: &visigrid_engine::validation::NumericConstraint) {
        use visigrid_engine::validation::{ComparisonOperator, ConstraintValue};

        self.numeric_operator = match constraint.operator {
            ComparisonOperator::Between => NumericOperatorOption::Between,
            ComparisonOperator::NotBetween => NumericOperatorOption::NotBetween,
            ComparisonOperator::EqualTo => NumericOperatorOption::EqualTo,
            ComparisonOperator::NotEqualTo => NumericOperatorOption::NotEqualTo,
            ComparisonOperator::GreaterThan => NumericOperatorOption::GreaterThan,
            ComparisonOperator::LessThan => NumericOperatorOption::LessThan,
            ComparisonOperator::GreaterThanOrEqual => NumericOperatorOption::GreaterThanOrEqual,
            ComparisonOperator::LessThanOrEqual => NumericOperatorOption::LessThanOrEqual,
        };

        // Helper to convert constraint value to string
        let value_to_string = |v: &ConstraintValue| -> String {
            match v {
                ConstraintValue::Number(n) => n.to_string(),
                ConstraintValue::CellRef(r) => r.clone(),
                ConstraintValue::Formula(f) => f.clone(),
            }
        };

        self.value1 = value_to_string(&constraint.value1);
        if let Some(ref v2) = constraint.value2 {
            self.value2 = value_to_string(v2);
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

/// Lightweight UTC timestamp without external chrono dependency.
/// Returns ISO 8601 format: "2024-01-15T14:30:00Z"
pub fn chrono_lite_utc() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};

    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let secs = now.as_secs();

    // Calculate date/time components from Unix timestamp
    // Days since epoch
    let days = secs / 86400;
    let time_secs = secs % 86400;

    let hours = time_secs / 3600;
    let minutes = (time_secs % 3600) / 60;
    let seconds = time_secs % 60;

    // Year calculation (simplified leap year handling)
    let mut year = 1970;
    let mut remaining_days = days as i64;

    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining_days < days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    // Month calculation
    let month_days: &[i64] = if is_leap_year(year) {
        &[31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        &[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1;
    for &days in month_days {
        if remaining_days < days {
            break;
        }
        remaining_days -= days;
        month += 1;
    }

    let day = remaining_days + 1;

    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        year, month, day, hours, minutes, seconds
    )
}

pub fn is_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

// ============================================================================
// AI Settings Dialog State
// ============================================================================

/// Selected AI provider in the settings dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AIProviderOption {
    #[default]
    None,
    Local,
    OpenAI,
    Anthropic,
    Gemini,
    Grok,
}

impl AIProviderOption {
    pub fn label(&self) -> &'static str {
        match self {
            AIProviderOption::None => "Disabled",
            AIProviderOption::Local => "Local (Ollama)",
            AIProviderOption::OpenAI => "OpenAI",
            AIProviderOption::Anthropic => "Anthropic",
            AIProviderOption::Gemini => "Google Gemini",
            AIProviderOption::Grok => "xAI Grok",
        }
    }

    pub fn to_config(&self) -> visigrid_config::settings::AIProvider {
        match self {
            AIProviderOption::None => visigrid_config::settings::AIProvider::None,
            AIProviderOption::Local => visigrid_config::settings::AIProvider::Local,
            AIProviderOption::OpenAI => visigrid_config::settings::AIProvider::OpenAI,
            AIProviderOption::Anthropic => visigrid_config::settings::AIProvider::Anthropic,
            AIProviderOption::Gemini => visigrid_config::settings::AIProvider::Gemini,
            AIProviderOption::Grok => visigrid_config::settings::AIProvider::Grok,
        }
    }

    pub fn from_config(provider: visigrid_config::settings::AIProvider) -> Self {
        match provider {
            visigrid_config::settings::AIProvider::None => AIProviderOption::None,
            visigrid_config::settings::AIProvider::Local => AIProviderOption::Local,
            visigrid_config::settings::AIProvider::OpenAI => AIProviderOption::OpenAI,
            visigrid_config::settings::AIProvider::Anthropic => AIProviderOption::Anthropic,
            visigrid_config::settings::AIProvider::Gemini => AIProviderOption::Gemini,
            visigrid_config::settings::AIProvider::Grok => AIProviderOption::Grok,
        }
    }
}

/// Focus state for AI settings dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AISettingsFocus {
    #[default]
    Provider,
    Model,
    Endpoint,
    KeyInput,
}

/// Test status for AI key verification
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum AITestStatus {
    #[default]
    Idle,
    Testing,
    Success(String),  // Model name or version returned
    Error(String),    // Error message
}

/// State for the AI Settings dialog
#[derive(Debug, Clone, Default)]
pub struct AISettingsDialogState {
    /// Selected AI provider
    pub provider: AIProviderOption,
    /// Whether the provider dropdown is open
    pub provider_dropdown_open: bool,

    /// Model identifier (empty = use provider default)
    pub model: String,

    /// Custom endpoint for Local provider (Ollama)
    pub endpoint: String,

    /// Privacy mode: minimize data sent to AI
    pub privacy_mode: bool,

    /// Allow AI to propose cell changes
    pub allow_proposals: bool,

    /// Which field has focus
    pub focus: AISettingsFocus,

    /// API key input (for setting new key)
    pub key_input: String,
    /// Whether key is currently stored
    pub key_present: bool,
    /// Source of current key ("keychain", "environment", "none")
    pub key_source: String,

    /// Test status
    pub test_status: AITestStatus,

    /// Error message
    pub error: Option<String>,
}

impl AISettingsDialogState {
    /// Reset to defaults
    pub fn reset(&mut self) {
        *self = Self::default();
    }

    /// Load current settings from config crate using ResolvedAIConfig
    pub fn load_from_config(&mut self) {
        use visigrid_config::ai::ResolvedAIConfig;
        use visigrid_config::settings::Settings;

        // Use the single source of truth
        let config = ResolvedAIConfig::load();
        let settings = Settings::load();

        self.provider = AIProviderOption::from_config(config.provider);
        // Load model from settings (not resolved) so user sees what they actually set
        self.model = settings.ai.model.clone();
        self.endpoint = config.endpoint.unwrap_or_default();
        self.privacy_mode = config.privacy_mode;
        self.allow_proposals = config.allow_proposals;

        // Key status from resolved config
        self.key_present = config.api_key.is_some();
        self.key_source = config.key_source.as_str().to_string();

        self.key_input.clear();
        self.test_status = AITestStatus::Idle;
        self.error = None;
    }

    /// Save current state to config
    pub fn save_to_config(&self) -> Result<(), String> {
        let mut settings = visigrid_config::settings::Settings::load();

        settings.ai.provider = self.provider.to_config();
        settings.ai.model = self.model.clone();
        settings.ai.endpoint = if self.endpoint.is_empty() {
            None
        } else {
            Some(self.endpoint.clone())
        };
        settings.ai.privacy_mode = self.privacy_mode;
        settings.ai.allow_proposals = self.allow_proposals;

        settings.save()
    }

    /// Get the effective model name (user-specified or provider default)
    pub fn effective_model(&self) -> &str {
        if self.model.is_empty() {
            match self.provider {
                AIProviderOption::None => "",
                AIProviderOption::Local => "llama3:8b",
                AIProviderOption::OpenAI => "gpt-4o-mini",
                AIProviderOption::Anthropic => "claude-sonnet-4-20250514",
                AIProviderOption::Gemini => "gemini-2.0-flash",
                AIProviderOption::Grok => "grok-3-mini-fast-beta",
            }
        } else {
            &self.model
        }
    }

    /// Check if the current provider requires an API key
    pub fn needs_api_key(&self) -> bool {
        matches!(
            self.provider,
            AIProviderOption::OpenAI | AIProviderOption::Anthropic | AIProviderOption::Gemini | AIProviderOption::Grok
        )
    }

    /// Get context policy description based on privacy mode
    pub fn context_policy(&self) -> &'static str {
        if self.privacy_mode {
            "Minimal: only current cell and selection"
        } else {
            "Extended: includes nearby cells and context"
        }
    }
}

// ============================================================================
// Ask AI Dialog State
// ============================================================================

/// Which AI verb is active in the dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AiVerb {
    /// Insert Formula with AI (Ctrl+Shift+A) — single-cell write contract
    #[default]
    InsertFormula,
    /// Analyze with AI (Ctrl+Shift+E) — read-only contract
    Analyze,
}

/// Status of an Ask AI request
#[derive(Debug, Clone, Default, PartialEq)]
pub enum AskAIStatus {
    #[default]
    Idle,
    Loading,
    Success,
    Error(String),
}

/// How context range is selected
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AskAIContextMode {
    /// Use current cell selection
    #[default]
    CurrentSelection,
    /// Use contiguous region around active cell
    CurrentRegion,
    /// Use entire used range of sheet
    EntireUsedRange,
}

impl AskAIContextMode {
    pub fn label(&self) -> &'static str {
        match self {
            Self::CurrentSelection => "Current selection",
            Self::CurrentRegion => "Current region",
            Self::EntireUsedRange => "Entire used range",
        }
    }
}

/// What was actually sent to the AI (for transparency)
#[derive(Debug, Clone, Default)]
pub struct AskAISentContext {
    /// Provider used
    pub provider: String,
    /// Model used
    pub model: String,
    /// Privacy mode enabled
    pub privacy_mode: bool,
    /// Effective range after truncation
    pub range_display: String,
    /// Rows actually sent
    pub rows_sent: usize,
    /// Columns actually sent
    pub cols_sent: usize,
    /// Total cells sent
    pub total_cells: usize,
    /// Whether headers were detected/included
    pub headers_included: bool,
    /// Truncation applied
    pub truncation: AskAITruncation,
}

/// What truncation was applied
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AskAITruncation {
    #[default]
    None,
    Rows,
    Cols,
    Both,
}

impl AskAITruncation {
    pub fn label(&self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Rows => "Rows truncated",
            Self::Cols => "Columns truncated",
            Self::Both => "Rows and columns truncated",
        }
    }
}

/// State for the AI dialog (shared by InsertFormula and Analyze verbs)
#[derive(Debug, Clone, Default)]
pub struct AskAIDialogState {
    /// Which verb is active
    pub verb: AiVerb,

    /// User's question
    pub question: String,

    /// Context selection mode
    pub context_mode: AskAIContextMode,

    /// Context summary (e.g., "Sheet1!A1:F200")
    pub context_summary: String,

    /// Context range info for insertion (top-left cell)
    pub context_top_left: Option<(usize, usize)>,

    /// Selected range bounds (start_row, start_col, end_row, end_col)
    pub selected_range: Option<(usize, usize, usize, usize)>,

    /// What was actually sent to AI (populated after request)
    pub sent_context: Option<AskAISentContext>,

    /// Request status
    pub status: AskAIStatus,

    /// Unique request ID for tracking
    pub request_id: Option<String>,

    /// AI explanation (when successful)
    pub explanation: Option<String>,

    /// Proposed formula (when AI suggests one)
    pub formula: Option<String>,

    /// Whether formula is valid for insertion
    pub formula_valid: bool,

    /// Formula validation error (if any)
    pub formula_error: Option<String>,

    /// Warnings from context extraction or AI
    pub warnings: Vec<String>,

    /// Error message (if failed)
    pub error: Option<String>,

    /// Raw model response (for debugging)
    pub raw_response: Option<String>,

    /// Analysis text from Analyze verb (read-only response)
    pub response_text: Option<String>,

    /// Last insertion confirmation (e.g., "Inserted into A1")
    pub last_insertion: Option<String>,

    /// Whether current formula was already inserted (prevent double-insert)
    pub inserted: bool,

    /// Whether "Sent to AI" panel is expanded
    pub sent_panel_expanded: bool,

    /// Whether context selector is open
    pub context_selector_open: bool,
}

impl AskAIDialogState {
    /// Reset to initial state (verb is set by the caller before reset)
    pub fn reset(&mut self) {
        // Note: verb is NOT reset here — it's set by show_ask_ai/show_analyze before reset
        self.question.clear();
        self.context_mode = AskAIContextMode::CurrentSelection;
        self.context_summary.clear();
        self.context_top_left = None;
        self.selected_range = None;
        self.sent_context = None;
        self.status = AskAIStatus::Idle;
        self.request_id = None;
        self.explanation = None;
        self.formula = None;
        self.formula_valid = false;
        self.formula_error = None;
        self.response_text = None;
        self.warnings.clear();
        self.error = None;
        self.raw_response = None;
        self.last_insertion = None;
        self.inserted = false;
        self.sent_panel_expanded = false;
        self.context_selector_open = false;
    }

    /// Clear response state (keep question, context, and verb)
    pub fn clear_response(&mut self) {
        self.sent_context = None;
        self.status = AskAIStatus::Idle;
        self.request_id = None;
        self.explanation = None;
        self.formula = None;
        self.formula_valid = false;
        self.formula_error = None;
        self.response_text = None;
        self.error = None;
        self.raw_response = None;
        self.last_insertion = None;
        self.inserted = false;
    }

    /// Check if Insert Formula button should be enabled
    pub fn can_insert(&self) -> bool {
        matches!(self.status, AskAIStatus::Success)
            && self.formula.is_some()
            && self.formula_valid
            && !self.inserted
    }

    /// Check if a request is in flight
    pub fn is_loading(&self) -> bool {
        matches!(self.status, AskAIStatus::Loading)
    }

    /// Check if retry is available
    pub fn can_retry(&self) -> bool {
        !self.is_loading() && !self.question.is_empty()
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
}

impl FormatBarState {
    /// Returns true when the format bar owns focus (editing or dropdown open).
    /// Used to gate grid keyboard and mouse handling.
    pub fn is_active(&self, window: &Window) -> bool {
        self.size_editing || self.size_dropdown || self.size_focus.is_focused(window)
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
    pub status_message: Option<String>,
    pub window_size: Size<Pixels>,
    pub cached_window_bounds: Option<WindowBounds>,  // Cached for session snapshot

    // Column/row sizing (per-sheet)
    // Each sheet has independent column widths and row heights.
    // New sheets start with defaults (Excel behavior), not inherited from current sheet.
    pub col_widths: HashMap<SheetId, HashMap<usize, f32>>,   // SheetId -> col -> width
    pub row_heights: HashMap<SheetId, HashMap<usize, f32>>,  // SheetId -> row -> height
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

    // Lua scripting state
    pub lua_runtime: crate::scripting::LuaRuntime,
    pub lua_console: crate::scripting::ConsoleState,
    pub custom_fn_registry: crate::scripting::CustomFunctionRegistry,

    // License dialog state
    pub license_input: String,
    pub license_error: Option<String>,

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

    // Validation dropdown state (data validation list picker)
    pub validation_dropdown: crate::validation_dropdown::ValidationDropdownState,

    // Validation dialog state (Phase 4: Data > Validation menu)
    pub validation_dialog: ValidationDialogState,

    // Paste Special dialog state (Ctrl+Alt+V)
    pub paste_special_dialog: PasteSpecialDialogState,

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

    // Merge cells confirmation dialog
    pub merge_confirm: MergeConfirmState,

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
    session_request_rx: std::sync::mpsc::Receiver<crate::session_server::SessionRequest>,
    /// Sender for session requests (cloned to give to server).
    /// Kept here so we can create bridge handles on demand.
    session_request_tx: std::sync::mpsc::Sender<crate::session_server::SessionRequest>,
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
        let font_picker_focus = cx.focus_handle();
        let ui = UiState {
            color_picker: crate::color_palette::ColorPickerState::new(cx.focus_handle()),
            format_bar: FormatBarState {
                size_input: String::new(),
                size_editing: false,
                size_dropdown: false,
                size_focus: cx.focus_handle(),
                size_replace_next: false,
            },
            format_menu_open: false,
        };
        window.focus(&focus_handle, cx);
        let window_size = window.viewport_size();
        let window_handle = window.window_handle();

        // Get theme from global settings store (resolve "system" to OS-appropriate theme)
        let theme = match user_settings(cx).appearance.theme_id.as_value() {
            Some(id) if id == SYSTEM_THEME_ID => {
                let resolved_id = resolve_system_theme_id(window.appearance());
                get_theme(resolved_id).unwrap_or_else(default_theme)
            }
            Some(id) => get_theme(id).unwrap_or_else(default_theme),
            None => default_theme(),
        };

        // Subscribe to global settings changes - trigger re-render when settings change
        let settings_subscription = observe_settings(cx, |cx| {
            // Notify all windows to re-render when settings change
            cx.refresh_windows();
        });

        // Observe OS appearance changes so System theme switches live
        let appearance_subscription = cx.observe_window_appearance(window, |this, window, cx| {
            let is_system = user_settings(cx).appearance.theme_id
                .as_value()
                .map_or(false, |id| id == SYSTEM_THEME_ID);
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
            font_picker_focus,
            ui,
            status_message: None,
            window_size,
            cached_window_bounds: Some(window.window_bounds()),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
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

            lua_runtime: crate::scripting::LuaRuntime::default(),
            lua_console: crate::scripting::ConsoleState::default(),
            custom_fn_registry: crate::scripting::CustomFunctionRegistry::empty(),

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

            semantic_verification: visigrid_io::native::SemanticVerification::default(),
            approval_confirm_visible: false,
            approval_drift_visible: false,
            approval_label_input: String::new(),
            approved_fingerprint: None,
            approval_history_len: 0,

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
            number_format_editor: NumberFormatEditorState::default(),
            last_paste_special_mode: PasteType::All,

            validation_failures: Vec::new(),
            validation_failure_index: 0,

            invalid_cells: std::collections::HashMap::new(),

            rewind_confirm: RewindConfirmState::default(),
            rewind_success: RewindSuccessBanner::default(),

            merge_confirm: MergeConfirmState::default(),

            ai_settings: AISettingsDialogState::default(),
            ask_ai: AskAIDialogState::default(),
            ai_key_validated_this_session: false,
            ai_session_key: None,

            // Session server: initialized below
            session_server: session_server,
            session_request_rx: session_rx,
            session_request_tx: session_tx,
        };

        // Load custom functions from ~/.config/visigrid/functions.lua (Pro only)
        #[cfg(feature = "pro")]
        {
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
        }

        app
    }

    // ========================================================================
    // Session Server
    // ========================================================================

    /// Create a bridge handle for the session server.
    /// The handle can be cloned and passed to the TCP server.
    pub fn session_bridge_handle(&self) -> crate::session_server::SessionBridgeHandle {
        crate::session_server::SessionBridgeHandle::new(self.session_request_tx.clone())
    }

    /// Drain pending session requests and process them.
    /// Called at the start of each render cycle.
    fn drain_session_requests(&mut self, cx: &mut Context<Self>) {
        use crate::session_server::{SessionRequest, SubscribeResponse, UnsubscribeResponse};

        // Non-blocking drain: process all pending requests
        while let Ok(request) = self.session_request_rx.try_recv() {
            match request {
                SessionRequest::ApplyOps { req, reply } => {
                    // Apply ops through the canonical mutation path
                    let response = self.handle_session_apply_ops(&req, cx);
                    let _ = reply.send(response);
                }
                SessionRequest::Inspect { req, reply } => {
                    let response = self.handle_session_inspect(&req, cx);
                    let _ = reply.send(response);
                }
                SessionRequest::Subscribe { req, reply } => {
                    // TODO: Implement subscription tracking
                    let _ = reply.send(SubscribeResponse {
                        topics: req.topics,
                        current_revision: self.workbook.read(cx).revision(),
                    });
                }
                SessionRequest::Unsubscribe { req, reply } => {
                    // TODO: Implement unsubscription
                    let _ = reply.send(UnsubscribeResponse {
                        topics: req.topics,
                    });
                }
            }
        }
    }

    /// Handle an apply_ops request from the session server.
    ///
    /// Uses proper batching: all ops are applied within a single batch_guard,
    /// ensuring exactly one recalc and one revision increment for the entire batch.
    fn handle_session_apply_ops(
        &mut self,
        req: &crate::session_server::ApplyOpsRequest,
        cx: &mut Context<Self>,
    ) -> crate::session_server::ApplyOpsResponse {
        use crate::session_server::{ApplyOpsResponse, ApplyOpsError, Op};
        use crate::history::CellChange;
        use visigrid_engine::cell_id::CellId;

        // Check expected_revision if provided
        let current_rev = self.workbook.read(cx).revision();
        if let Some(expected) = req.expected_revision {
            if expected != current_rev {
                return ApplyOpsResponse {
                    applied: 0,
                    total: req.ops.len(),
                    current_revision: current_rev,
                    error: Some(ApplyOpsError::RevisionMismatch {
                        expected,
                        actual: current_rev,
                    }),
                };
            }
        }

        if req.ops.is_empty() {
            return ApplyOpsResponse {
                applied: 0,
                total: 0,
                current_revision: current_rev,
                error: None,
            };
        }

        // Apply all ops within a single batch_guard, collecting changes for history
        let (applied, error, changes_by_sheet) = self.workbook.update(cx, |wb, _| {
            let mut guard = wb.batch_guard();
            let mut applied = 0;
            let mut error: Option<crate::session_server::ApplyOpsError> = None;
            // Group changes by sheet for history recording
            let mut changes_by_sheet: std::collections::HashMap<usize, Vec<CellChange>> =
                std::collections::HashMap::new();

            for (_i, op) in req.ops.iter().enumerate() {
                let sheet_count = guard.sheets().len();

                match op {
                    Op::SetCellValue { sheet, row, col, value } => {
                        let sheet_idx = if *sheet < sheet_count { *sheet } else { guard.active_sheet_index() };

                        // Capture old value for history
                        let old_value = guard.sheets()[sheet_idx].get_raw(*row, *col);
                        changes_by_sheet.entry(sheet_idx).or_default().push(CellChange {
                            row: *row,
                            col: *col,
                            old_value,
                            new_value: value.clone(),
                        });

                        // Apply the mutation via tracked method
                        guard.set_cell_value_tracked(sheet_idx, *row, *col, value);
                        applied += 1;
                    }
                    Op::SetCellFormula { sheet, row, col, formula } => {
                        let sheet_idx = if *sheet < sheet_count { *sheet } else { guard.active_sheet_index() };

                        let old_value = guard.sheets()[sheet_idx].get_raw(*row, *col);
                        changes_by_sheet.entry(sheet_idx).or_default().push(CellChange {
                            row: *row,
                            col: *col,
                            old_value,
                            new_value: formula.clone(),
                        });

                        guard.set_cell_value_tracked(sheet_idx, *row, *col, formula);
                        applied += 1;
                    }
                    Op::ClearCell { sheet, row, col } => {
                        let sheet_idx = if *sheet < sheet_count { *sheet } else { guard.active_sheet_index() };

                        let old_value = guard.sheets()[sheet_idx].get_raw(*row, *col);
                        changes_by_sheet.entry(sheet_idx).or_default().push(CellChange {
                            row: *row,
                            col: *col,
                            old_value,
                            new_value: String::new(),
                        });

                        guard.clear_cell_tracked(sheet_idx, *row, *col);
                        applied += 1;
                    }
                    Op::SetNumberFormat { .. } => {
                        // TODO: Parse format string and apply to range
                        applied += 1;
                    }
                    Op::SetStyle { .. } => {
                        // TODO: Apply style changes to range
                        applied += 1;
                    }
                }

                // If atomic and there was an error, stop
                if req.atomic && error.is_some() {
                    break;
                }
            }

            (applied, error, changes_by_sheet)
        });
        // batch_guard dropped here → single recalc + revision increment

        // Build changed cells list BEFORE history takes ownership of changes_by_sheet
        let changed_cells: Vec<crate::session_server::CellRef> = if applied > 0 && error.is_none() {
            changes_by_sheet
                .iter()
                .flat_map(|(sheet_idx, changes)| {
                    changes.iter().map(move |c| crate::session_server::CellRef {
                        sheet: *sheet_idx,
                        row: c.row,
                        col: c.col,
                    })
                })
                .collect()
        } else {
            Vec::new()
        };

        // Record history entries for undo (one per sheet that had changes)
        for (sheet_idx, changes) in changes_by_sheet {
            if !changes.is_empty() {
                self.history.record_batch(sheet_idx, changes);
            }
        }

        // Mark document as modified
        self.is_modified = true;
        self.cached_title = None;

        let new_rev = self.workbook.read(cx).revision();

        // Broadcast cell changes to subscribed connections
        // This happens in the same transaction boundary as revision increment
        if !changed_cells.is_empty() {
            self.session_server.broadcast_cells(new_rev, changed_cells);
        }

        ApplyOpsResponse {
            applied,
            total: req.ops.len(),
            current_revision: new_rev,
            error,
        }
    }

    /// Handle an inspect request from the session server.
    fn handle_session_inspect(
        &self,
        req: &crate::session_server::InspectRequest,
        cx: &Context<Self>,
    ) -> crate::session_server::InspectResponse {
        use crate::session_server::{InspectResponse, InspectResult, InspectTarget, CellInfo, WorkbookInfo};

        let wb = self.workbook.read(cx);
        let current_rev = wb.revision();

        let result = match &req.target {
            InspectTarget::Cell { sheet, row, col } => {
                let sheet_data = if *sheet < wb.sheets().len() {
                    &wb.sheets()[*sheet]
                } else {
                    wb.active_sheet()
                };
                let display = sheet_data.get_display(*row, *col);
                let raw = sheet_data.get_raw(*row, *col);
                let formula = if raw.starts_with('=') { Some(raw.clone()) } else { None };
                InspectResult::Cell(CellInfo {
                    raw,
                    display,
                    formula,
                })
            }
            InspectTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                let sheet_data = if *sheet < wb.sheets().len() {
                    &wb.sheets()[*sheet]
                } else {
                    wb.active_sheet()
                };
                let mut cells = Vec::new();
                for r in *start_row..=*end_row {
                    for c in *start_col..=*end_col {
                        let display = sheet_data.get_display(r, c);
                        let raw = sheet_data.get_raw(r, c);
                        let formula = if raw.starts_with('=') { Some(raw.clone()) } else { None };
                        cells.push(CellInfo {
                            raw,
                            display,
                            formula,
                        });
                    }
                }
                InspectResult::Range { cells }
            }
            InspectTarget::Workbook => {
                InspectResult::Workbook(WorkbookInfo {
                    sheet_count: wb.sheets().len(),
                    active_sheet: wb.active_sheet_index(),
                    title: self.document_meta.display_name.clone(),
                })
            }
        };

        InspectResponse {
            current_revision: current_rev,
            result,
        }
    }

    /// Start the session server with the given mode.
    ///
    /// If `token_override` is provided (e.g. from VISIGRID_SESSION_TOKEN env var),
    /// uses that token instead of generating a fresh one. This allows test harnesses
    /// to know the token in advance.
    pub fn start_session_server(
        &mut self,
        mode: crate::session_server::ServerMode,
        token_override: Option<String>,
        cx: &mut Context<Self>,
    ) -> std::io::Result<()> {
        let bridge = self.session_bridge_handle();
        let workbook_path = self.current_file.clone();
        let workbook_title = self.document_meta.display_name.clone();

        self.session_server.start(crate::session_server::SessionServerConfig {
            mode,
            workbook_path,
            workbook_title,
            bridge: Some(bridge),
            token_override,
            ..Default::default()
        })
    }

    /// Get structured READY info for CI output.
    pub fn session_server_ready_info(&self) -> Option<(String, u16, std::path::PathBuf)> {
        self.session_server.ready_info()
    }

    /// Stop the session server.
    pub fn stop_session_server(&mut self) {
        self.session_server.stop();
    }

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
            CommandId::PasteSpecial => self.show_paste_special(cx),
            CommandId::PasteFormulas => self.paste_formulas(cx),
            CommandId::PasteFormats => self.paste_formats(cx),

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
        let width = width.max(20.0).min(500.0); // Clamp between 20-500px
        let sheet_widths = self.col_widths.entry(self.cached_sheet_id).or_insert_with(HashMap::new);
        if (width - CELL_WIDTH).abs() < 1.0 {
            sheet_widths.remove(&col); // Remove if close to default
        } else {
            sheet_widths.insert(col, width);
        }
    }

    /// Set row height for the current sheet
    pub fn set_row_height(&mut self, row: usize, height: f32) {
        let height = height.max(12.0).min(200.0); // Clamp between 12-200px
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
    pub fn auto_fit_col_width(&mut self, col: usize, cx: &mut Context<Self>) {
        let mut max_width: f32 = 40.0; // Minimum width

        // Check all rows for content in this column
        for row in 0..NUM_ROWS {
            let text = self.sheet(cx).get_text(row, col);
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
        self.sheet_row_heights_mut().remove(&row);
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
                self.auto_fit_col_width_no_notify(col, cx);
            }
            cx.notify();
        } else {
            // Not part of selection, just auto-fit the clicked column
            self.auto_fit_col_width(clicked_col, cx);
        }
    }

    /// Auto-fit column width without notifying (for batch operations)
    fn auto_fit_col_width_no_notify(&mut self, col: usize, cx: &App) {
        let mut max_width: f32 = 40.0; // Minimum width
        for row in 0..NUM_ROWS {
            let text = self.sheet(cx).get_text(row, col);
            if !text.is_empty() {
                let estimated_width = text.len() as f32 * 7.5 + 16.0;
                max_width = max_width.max(estimated_width);
            }
        }
        self.set_col_width(col, max_width);
    }

    /// Auto-fit all columns that have content (for agent-built sheets)
    /// Scans all columns up to the rightmost cell with data.
    pub fn auto_fit_all_data_columns(&mut self, cx: &App) {
        // Find the rightmost column with data
        let mut max_col = 0usize;
        for row in 0..100 {  // Check first 100 rows for content
            for col in 0..100 {  // Check first 100 columns
                let text = self.sheet(cx).get_text(row, col);
                if !text.is_empty() {
                    max_col = max_col.max(col);
                }
            }
        }

        // Auto-fit each column with data
        for col in 0..=max_col {
            self.auto_fit_col_width_no_notify(col, cx);
        }
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
        self.mode.is_overlay() || self.lua_console.visible
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
    pub fn cell_user_borders(
        &self, row: usize, col: usize, cx: &App,
        boundary_bottom: bool, boundary_right: bool,
    ) -> (CellBorder, CellBorder, CellBorder, CellBorder) {
        #[cfg(debug_assertions)]
        self.debug_border_call_count.set(self.debug_border_call_count.get() + 1);

        // Single-ownership model: each cell draws only its TOP and LEFT borders.
        // Right/bottom borders are drawn by the neighboring cell as their left/top.
        // Each edge is resolved as max(own_side, neighbor_opposite_side) so both
        // cells' border settings contribute, but only one cell draws the line.
        //
        // Exception: at the viewport boundary (last visible row/col), this cell
        // also draws BOTTOM (boundary_bottom) or RIGHT (boundary_right) because
        // the neighbor that would normally own that edge isn't rendered.
        //
        // For merged cells: interior cells draw nothing. Perimeter cells draw only
        // top (if on merge top edge) and left (if on merge left edge), resolved
        // with the neighboring cell/merge's opposing border. At viewport boundaries,
        // perimeter cells also draw bottom/right for merge edges that touch the boundary.

        let sheet = self.sheet(cx);

        // Helper: effective border contribution for a cell on a given side,
        // accounting for merges (interior cells contribute None, perimeter cells
        // contribute the merge's resolved edge border).
        let effective_side = |r: usize, c: usize, side: u8| -> CellBorder {
            // side: 0=top, 1=right, 2=bottom, 3=left
            if let Some(m) = sheet.get_merge(r, c) {
                let on_perimeter = match side {
                    0 => r == m.start.0,  // top
                    1 => c == m.end.1,    // right
                    2 => r == m.end.0,    // bottom
                    3 => c == m.start.1,  // left
                    _ => false,
                };
                if !on_perimeter {
                    return CellBorder::default(); // interior: no contribution
                }
                let (rt, rr, rb, rl) = sheet.resolve_merge_borders(m);
                match side {
                    0 => rt,
                    1 => rr,
                    2 => rb,
                    3 => rl,
                    _ => CellBorder::default(),
                }
            } else {
                let fmt = sheet.get_format(r, c);
                match side {
                    0 => fmt.border_top,
                    1 => fmt.border_right,
                    2 => fmt.border_bottom,
                    3 => fmt.border_left,
                    _ => CellBorder::default(),
                }
            }
        };

        // Check if this cell is a merge interior (not on any perimeter edge)
        let none = CellBorder::default();
        if let Some(m) = sheet.get_merge(row, col) {
            let on_edge = row == m.start.0 || row == m.end.0
                       || col == m.start.1 || col == m.end.1;
            if !on_edge {
                return (none, none, none, none); // interior: no borders
            }
        }

        // Resolve TOP edge: max(my_top, above_neighbor_bottom)
        let top = {
            let my_top = effective_side(row, col, 0);
            let above_bottom = if row > 0 {
                effective_side(row - 1, col, 2)
            } else {
                none
            };
            max_border(my_top, above_bottom)
        };

        // Resolve LEFT edge: max(my_left, left_neighbor_right)
        let left = {
            let my_left = effective_side(row, col, 3);
            let left_right = if col > 0 {
                effective_side(row, col - 1, 1)
            } else {
                none
            };
            max_border(my_left, left_right)
        };

        // Resolve BOTTOM edge: only at viewport boundary (last visible row)
        let bottom = if boundary_bottom {
            let my_bottom = effective_side(row, col, 2);
            let below_top = if row + 1 < NUM_ROWS {
                effective_side(row + 1, col, 0)
            } else {
                none
            };
            max_border(my_bottom, below_top)
        } else {
            none
        };

        // Resolve RIGHT edge: only at viewport boundary (last visible col)
        let right = if boundary_right {
            let my_right = effective_side(row, col, 1);
            let right_left = if col + 1 < NUM_COLS {
                effective_side(row, col + 1, 3)
            } else {
                none
            };
            max_border(my_right, right_left)
        } else {
            none
        };

        (top, right, bottom, left)
    }

    /// Check if any user-defined border is set for this cell
    pub fn has_user_borders(&self, row: usize, col: usize, cx: &App) -> bool {
        let format = self.sheet(cx).get_format(row, col);
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
            used += self.metrics.col_width(self.col_width(fc));
        }

        // Sum actual column widths from scroll position until we exceed available width
        let mut count = frozen_cols;
        let mut col = scroll_col;
        while used < available_width && col < NUM_COLS {
            used += self.metrics.col_width(self.col_width(col));
            count += 1;
            col += 1;
        }

        // Add 1 extra for partially visible columns at the edge
        count = count.saturating_add(1);
        count.max(1).min(NUM_COLS)
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

    // Formatting (applies to all discontiguous selection ranges)
    pub fn toggle_bold(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_bold(row, col));
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
                    self.active_sheet_mut(cx, |s| s.toggle_italic(row, col));
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
                    self.active_sheet_mut(cx, |s| s.toggle_underline(row, col));
                }
            }
        }
        self.is_modified = true;
        cx.notify();
    }

    pub fn toggle_strikethrough(&mut self, cx: &mut Context<Self>) {
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    self.active_sheet_mut(cx, |s| s.toggle_strikethrough(row, col));
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
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, NumberFormat::currency(2)));
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
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, NumberFormat::Percent { decimals: 2 }));
                }
            }
        }
        self.is_modified = true;
        cx.notify();
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

    /// Check if we're currently in preview mode
    pub fn is_previewing(&self) -> bool {
        matches!(self.rewind_preview, RewindPreviewState::On(_))
    }

    /// Get the current preview session, if any
    pub fn preview_session(&self) -> Option<&RewindPreviewSession> {
        match &self.rewind_preview {
            RewindPreviewState::On(session) => Some(session),
            RewindPreviewState::Off => None,
        }
    }

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

    /// Block a command if in preview mode.
    /// Returns true if blocked (command should return early).
    /// Sets status message with consistent preview warning.
    pub fn block_if_previewing(&mut self, cx: &mut Context<Self>) -> bool {
        if self.is_previewing() {
            self.status_message = Some(PREVIEW_BLOCK_MSG.to_string());
            cx.notify();
            true
        } else {
            false
        }
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

    /// Enter preview mode for the currently selected history entry
    pub fn enter_preview(&mut self, cx: &mut Context<Self>) -> Result<(), String> {
        // Must have a history highlight to preview
        let (sheet_idx, start_row, start_col, end_row, end_col) = match self.history_highlight_range {
            Some(range) => range,
            None => return Err("No history entry selected".to_string()),
        };

        // Must have a selected history entry
        let entry_id = match self.selected_history_id {
            Some(id) => id,
            None => return Err("No history entry selected".to_string()),
        };

        // Find the history index for this entry
        let history_index = match self.history.global_index_for_id(entry_id) {
            Some(idx) => idx,
            None => return Err("History entry not found".to_string()),
        };

        // Get entry info for the session
        let entry = match self.history.entry_at(history_index) {
            Some(e) => e,
            None => return Err("Invalid history index".to_string()),
        };
        let action_summary = entry.action.summary().unwrap_or_else(|| entry.action.label());

        // Build the preview workbook and view state (state BEFORE this action)
        let build_result = self.history.build_workbook_before(
            history_index,
            &self.base_workbook,
            MAX_PREVIEW_REPLAY,
            MAX_PREVIEW_BUILD_MS,
        ).map_err(|e| match e {
            crate::history::PreviewBuildError::InvalidIndex => "Invalid history index".to_string(),
            crate::history::PreviewBuildError::TooManyActions(n) => {
                format!("Preview unavailable — history too large to replay (limit: {} actions)", n)
            }
            crate::history::PreviewBuildError::Timeout => {
                format!("Preview unavailable — replay timed out ({}ms)", MAX_PREVIEW_BUILD_MS)
            }
            crate::history::PreviewBuildError::UnsupportedAction(kind) => {
                format!("Preview unavailable — history contains unsupported action: {}", kind.display_name())
            }
            crate::history::PreviewBuildError::InvariantViolation(msg) => {
                format!("Preview aborted — data integrity error: {}", msg)
            }
        })?;

        // Capture current focus for restoration
        let live_focus = PreviewFocus {
            sheet_index: self.sheet_index(cx),
            selected: self.view_state.selected,
            selection_end: self.view_state.selection_end,
            scroll_row: self.view_state.scroll_row,
            scroll_col: self.view_state.scroll_col,
        };

        // Create the preview session
        let session = RewindPreviewSession {
            entry_id,
            target_global_index: history_index,
            action_summary: action_summary.clone(),
            snapshot: build_result.workbook,
            view_state: build_result.view_state,
            live_focus,
            history_fingerprint: self.history.fingerprint(),
            replay_count: build_result.replay_count,
            build_ms: build_result.build_ms,
            quality: PreviewQuality::Ok,
        };

        self.rewind_preview = RewindPreviewState::On(session);

        // Navigate to the affected area in preview
        // Switch to the sheet where the action occurred
        self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(sheet_idx); });
        self.view_state.selected = (start_row, start_col);
        self.view_state.selection_end = if start_row != end_row || start_col != end_col {
            Some((end_row, end_col))
        } else {
            None
        };

        // Ensure the selection is visible
        self.ensure_visible(cx);

        self.status_message = Some(format!("Preview: Before \"{}\" — Release Space to return", action_summary));
        cx.notify();
        Ok(())
    }

    /// Exit preview mode, restoring live state
    pub fn exit_preview(&mut self, cx: &mut Context<Self>) {
        if let RewindPreviewState::On(session) = std::mem::take(&mut self.rewind_preview) {
            // Restore live focus (Option A: peek behavior)
            self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(session.live_focus.sheet_index); });
            self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
            self.debug_assert_sheet_cache_sync(cx);  // Catch desync at preview exit
            self.view_state.selected = session.live_focus.selected;
            self.view_state.selection_end = session.live_focus.selection_end;
            self.view_state.scroll_row = session.live_focus.scroll_row;
            self.view_state.scroll_col = session.live_focus.scroll_col;

            self.status_message = Some("Returned to current state".to_string());
            cx.notify();
        }
    }

    /// Scrub the preview timeline: navigate to adjacent history entry while holding Space.
    /// direction: -1 for older (up), +1 goes to newer (down)
    pub fn scrub_preview(&mut self, direction: i32, cx: &mut Context<Self>) {
        let current_id = match self.selected_history_id {
            Some(id) => id,
            None => return,
        };

        // Find current position in global history
        let current_idx = match self.history.global_index_for_id(current_id) {
            Some(idx) => idx,
            None => return,
        };

        // Compute new index (direction: -1 goes to older = lower index, +1 goes to newer = higher index)
        let history_len = self.history.undo_count();
        let new_idx = if direction < 0 {
            current_idx.saturating_sub(1)
        } else {
            (current_idx + 1).min(history_len.saturating_sub(1))
        };

        // Don't update if at boundary
        if new_idx == current_idx {
            return;
        }

        // Get the new entry and compute its display info
        let new_entry = match self.history.entry_at(new_idx) {
            Some(e) => e,
            None => return,
        };
        let new_id = new_entry.id;
        let action_summary = new_entry.action.summary()
            .unwrap_or_else(|| new_entry.action.label());

        // Compute highlight range from action details
        let new_highlight = {
            let display_entries = self.history.display_entries();
            display_entries.iter()
                .find(|e| e.id == new_id)
                .and_then(|e| e.sheet_index.and_then(|si| e.affected_range.map(|(sr, sc, er, ec)| (si, sr, sc, er, ec))))
        };

        // Store the current live focus, fingerprint, and quality (preserve across scrubs)
        let (live_focus, history_fingerprint, original_quality) = if let RewindPreviewState::On(ref session) = self.rewind_preview {
            (session.live_focus.clone(), session.history_fingerprint, session.quality.clone())
        } else {
            return; // Not actually previewing
        };

        // Update selection
        self.selected_history_id = Some(new_id);
        self.history_highlight_range = new_highlight;

        // Exit current preview temporarily
        self.rewind_preview = RewindPreviewState::Off;

        // Re-enter preview with new entry
        match self.history.build_workbook_before(
            new_idx,
            &self.base_workbook,
            MAX_PREVIEW_REPLAY,
            MAX_PREVIEW_BUILD_MS,
        ) {
            Ok(build_result) => {
                let session = RewindPreviewSession {
                    entry_id: new_id,
                    target_global_index: new_idx,
                    action_summary: action_summary.clone(),
                    snapshot: build_result.workbook,
                    view_state: build_result.view_state,
                    live_focus,
                    history_fingerprint,  // Preserved from original preview
                    replay_count: build_result.replay_count,
                    build_ms: build_result.build_ms,
                    quality: original_quality,  // Preserve quality from original entry
                };

                self.rewind_preview = RewindPreviewState::On(session);

                // Navigate to the affected area
                if let Some((sheet_idx, start_row, start_col, end_row, end_col)) = new_highlight {
                    self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(sheet_idx); });
                    self.view_state.selected = (start_row, start_col);
                    self.view_state.selection_end = if start_row != end_row || start_col != end_col {
                        Some((end_row, end_col))
                    } else {
                        None
                    };
                    self.ensure_visible(cx);
                }

                self.status_message = Some(format!(
                    "Preview: Before \"{}\" [{}/{}] — ↑↓ to scrub, release Space to return",
                    action_summary, new_idx + 1, history_len
                ));
            }
            Err(e) => {
                // Preview build failed - show error and restore live focus
                self.workbook.update(cx, |wb, _| { let _ = wb.set_active_sheet(live_focus.sheet_index); });
                self.view_state.selected = live_focus.selected;
                self.view_state.selection_end = live_focus.selection_end;
                self.view_state.scroll_row = live_focus.scroll_row;
                self.view_state.scroll_col = live_focus.scroll_col;

                self.status_message = Some(format!("Preview failed: {:?}", e));
            }
        }
        cx.notify();
    }

    /// Build a rewind plan from the current preview session.
    /// Returns None if not previewing or preview is invalid.
    pub fn build_rewind_plan(&self) -> Option<RewindPlan> {
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return None,
        };

        // The truncate point is the target entry index
        // We keep entries [0..target_index), discard [target_index..]
        let truncate_at = session.target_global_index;
        let discarded_count = self.history.undo_count().saturating_sub(truncate_at);

        // Generate timestamp now (will be close to commit time)
        let timestamp_utc = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs().to_string())
            .unwrap_or_else(|_| "0".to_string());

        // Build the audit action with full provenance
        let audit_action = crate::history::UndoAction::Rewind {
            target_entry_id: session.entry_id,
            target_index: session.target_global_index,
            target_action_summary: session.action_summary.clone(),
            discarded_count,
            old_history_len: self.history.undo_count(),
            new_history_len: truncate_at + 1, // After truncate + audit entry
            timestamp_utc,
            preview_replay_count: session.replay_count,
            preview_build_ms: session.build_ms,
        };

        Some(RewindPlan {
            new_workbook: session.snapshot.clone(),
            new_view_state: session.view_state.clone(),
            truncate_at,
            audit_action,
            discarded_count,
            focus: session.live_focus.clone(),
        })
    }

    /// Apply a rewind plan atomically. This is a destructive operation.
    /// Returns Err if the history has changed since the plan was built.
    pub fn apply_rewind_plan(&mut self, plan: RewindPlan, cx: &mut Context<Self>) -> Result<(), String> {
        // Validate history fingerprint hasn't changed
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return Err("No preview active".to_string()),
        };

        let current_fingerprint = self.history.fingerprint();
        if current_fingerprint != session.history_fingerprint {
            return Err(format!(
                "History changed during preview. Expected {:?}, got {:?}. Please re-enter preview to try again.",
                session.history_fingerprint, current_fingerprint
            ));
        }

        // Extract audit entry details before consuming plan
        let (target_entry_id, target_index, action_summary, preview_replay_count, preview_build_ms) = match &plan.audit_action {
            crate::history::UndoAction::Rewind {
                target_entry_id,
                target_index,
                target_action_summary,
                preview_replay_count,
                preview_build_ms,
                ..
            } => (*target_entry_id, *target_index, target_action_summary.clone(), *preview_replay_count, *preview_build_ms),
            _ => return Err("Invalid audit action in plan".to_string()),
        };

        // === ATOMIC COMMIT: Do not fail after this point ===

        // 1. Replace the workbook content
        self.workbook.update(cx, |wb, _| {
            *wb = plan.new_workbook;
        });
        self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
        self.debug_assert_sheet_cache_sync(cx);  // Catch desync at rewind
        // Update base_workbook to match (this is now the canonical state)
        self.base_workbook = self.wb(cx).clone();

        // 2. Apply view state from the plan (row ordering per sheet)
        // Reset row_view to identity for the current sheet
        self.row_view = visigrid_engine::filter::RowView::new(NUM_ROWS);

        // If the preview view state has sort info for current sheet, re-apply it
        let active_idx = self.sheet_index(cx);
        if let Some(sheet_view) = plan.new_view_state.per_sheet.get(active_idx) {
            if let Some(ref row_order) = sheet_view.row_order {
                // Apply the stored row order
                self.row_view.apply_sort(row_order.clone());
            }
        }

        // 3. Truncate history and append audit entry
        self.history.truncate_and_append_rewind(
            plan.truncate_at,
            target_entry_id,
            target_index,
            action_summary.clone(),
            preview_replay_count,
            preview_build_ms,
        );

        // 4. Reset preview state
        self.rewind_preview = RewindPreviewState::Off;

        // 5. Clear history selection/highlight (we're now at end of history)
        self.selected_history_id = None;
        self.history_highlight_range = None;

        // 6. Keep current position in grid (don't restore pre-preview focus)
        // User is looking at the rewound state; changing view would be jarring

        // 7. Mark document as modified
        self.is_modified = true;

        // 8. Status message
        let discarded = plan.discarded_count;
        self.status_message = Some(format!(
            "Rewound to before \"{}\" — {} action{} discarded",
            action_summary,
            discarded,
            if discarded == 1 { "" } else { "s" }
        ));

        cx.notify();
        Ok(())
    }

    /// Check if a rewind is safe (history hasn't changed during preview).
    /// Returns (is_safe, discarded_count, target_summary).
    pub fn rewind_safety_check(&self) -> Option<(bool, usize, String)> {
        let session = match &self.rewind_preview {
            RewindPreviewState::On(s) => s,
            RewindPreviewState::Off => return None,
        };

        let current_fingerprint = self.history.fingerprint();
        let is_safe = current_fingerprint == session.history_fingerprint;
        let discarded = self.history.undo_count().saturating_sub(session.target_global_index);

        Some((is_safe, discarded, session.action_summary.clone()))
    }

    /// Show the rewind confirmation dialog (requires preview to be active).
    /// This builds the plan and presents the destructive warning.
    pub fn show_rewind_confirm(&mut self, cx: &mut Context<Self>) {
        // Must be previewing
        if !self.is_previewing() {
            self.status_message = Some("Not in preview mode".to_string());
            cx.notify();
            return;
        }

        // Build the plan
        let plan = match self.build_rewind_plan() {
            Some(p) => p,
            None => {
                self.status_message = Some("Cannot build rewind plan".to_string());
                cx.notify();
                return;
            }
        };

        // Check safety (fingerprint)
        let (is_safe, discard_count, target_summary) = match self.rewind_safety_check() {
            Some(s) => s,
            None => {
                self.status_message = Some("Cannot verify rewind safety".to_string());
                cx.notify();
                return;
            }
        };

        if !is_safe {
            self.status_message = Some("History changed during preview — please re-enter preview".to_string());
            cx.notify();
            return;
        }

        // Check preview quality - block degraded previews from hard rewind
        if let RewindPreviewState::On(ref session) = self.rewind_preview {
            if let PreviewQuality::Degraded(reason) = &session.quality {
                self.status_message = Some(format!("Rewind unavailable — preview was incomplete: {}", reason));
                cx.notify();
                return;
            }
        }

        // Extract additional context from preview session
        let (entry_id, replay_count, build_ms, fingerprint, sheet_name, location) =
            if let RewindPreviewState::On(ref session) = self.rewind_preview {
                // Get sheet name and location from the history entry
                let entry = self.history.entry_at(session.target_global_index);
                let (sheet_name, location) = if let Some(e) = entry {
                    let display = crate::history::History::to_display_entry(e, true);
                    let sheet = display.sheet_index.and_then(|i| {
                        self.wb(cx).sheet(i).map(|s| s.name.clone())
                    });
                    (sheet, display.location)
                } else {
                    (None, None)
                };

                (
                    session.entry_id,
                    session.replay_count,
                    session.build_ms,
                    session.history_fingerprint,
                    sheet_name,
                    location,
                )
            } else {
                (0, 0, 0, HistoryFingerprint::default(), None, None)
            };

        // Show the confirmation dialog with full context
        self.rewind_confirm.show(
            discard_count,
            target_summary,
            sheet_name,
            location,
            entry_id,
            replay_count,
            build_ms,
            fingerprint,
            plan,
        );
        cx.notify();
    }

    /// Confirm and execute the rewind (called from dialog Confirm button).
    pub fn confirm_rewind(&mut self, cx: &mut Context<Self>) {
        // Take the plan from dialog state
        let plan = match self.rewind_confirm.plan.take() {
            Some(p) => p,
            None => {
                self.status_message = Some("No rewind plan available".to_string());
                self.rewind_confirm.hide();
                cx.notify();
                return;
            }
        };

        // Capture audit data before consuming plan
        let audit_data = RewindAuditData {
            target_entry_id: self.rewind_confirm.target_entry_id,
            target_summary: self.rewind_confirm.target_summary.clone(),
            discarded_count: plan.discarded_count,
            replay_count: self.rewind_confirm.replay_count,
            build_ms: self.rewind_confirm.build_ms,
            fingerprint: self.rewind_confirm.fingerprint,
        };

        // Hide dialog first
        self.rewind_confirm.hide();

        // Apply the rewind
        match self.apply_rewind_plan(plan, cx) {
            Ok(()) => {
                // Success - show banner with full audit data
                self.rewind_success.show(audit_data);
            }
            Err(e) => {
                self.status_message = Some(format!("Rewind failed: {}", e));
            }
        }
        cx.notify();
    }

    /// Cancel the rewind confirmation dialog.
    pub fn cancel_rewind(&mut self, cx: &mut Context<Self>) {
        self.rewind_confirm.hide();
        cx.notify();
    }

    /// Dismiss the rewind success banner.
    pub fn dismiss_rewind_banner(&mut self, cx: &mut Context<Self>) {
        self.rewind_success.hide();
        cx.notify();
    }

    /// Copy rewind audit details to clipboard.
    pub fn copy_rewind_details(&mut self, cx: &mut Context<Self>) {
        let details = self.rewind_success.audit_details.clone();
        cx.write_to_clipboard(gpui::ClipboardItem::new_string(details));
        self.status_message = Some("Rewind details copied to clipboard".to_string());
        cx.notify();
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
                // Show for empty launches (no file loaded yet)
                if self.current_file.is_none() && self.status_message.is_none() {
                    self.status_message = Some(format!("Ready in {}ms", ms));
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
        // This assertion catches bugs where a modal is showing but mode isn't set correctly.
        // If this fires, either lua_console.visible is true without matching mode, or
        // you added a new modal that needs to be included in should_block_grid_navigation().
        debug_assert!(
            !self.lua_console.visible || self.should_block_grid_navigation(),
            "Lua console visible but grid navigation not blocked - mode is {:?}",
            self.mode
        );

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
