use gpui::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::PathBuf;
use std::fs;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::Workbook;
use visigrid_engine::formula::eval::CellLookup;
use visigrid_engine::named_range::is_valid_name;

use crate::history::{History, CellChange, UndoAction};
use crate::mode::Mode;
use crate::search::{SearchEngine, SearchAction, CommandId, CommandSearchProvider, GoToSearchProvider, SearchItem};
use crate::theme::{Theme, TokenKey, visigrid_theme, builtin_themes, get_theme};
use crate::views;
use crate::formula_context::{tokenize_for_highlight, TokenType};

// Re-export from autocomplete module for external access
pub use crate::autocomplete::{SignatureHelpInfo, FormulaErrorInfo};

// User settings that persist across sessions
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
struct Settings {
    theme_id: Option<String>,
    #[serde(default)]
    name_tooltip_dismissed: bool,
}

impl Settings {
    fn path() -> Option<PathBuf> {
        dirs::config_dir().map(|p| p.join("visigrid").join("settings.json"))
    }

    fn load() -> Self {
        Self::path()
            .and_then(|p| fs::read_to_string(p).ok())
            .and_then(|s| serde_json::from_str(&s).ok())
            .unwrap_or_default()
    }

    fn save(&self) {
        if let Some(path) = Self::path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }
            let _ = fs::write(path, serde_json::to_string_pretty(self).unwrap_or_default());
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

use visigrid_engine::cell::{Alignment, VerticalAlignment, TextOverflow, NumberFormat};

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

pub struct Spreadsheet {
    // Core data
    pub workbook: Workbook,
    pub history: History,

    // Selection
    pub selected: (usize, usize),                              // Anchor of active selection
    pub selection_end: Option<(usize, usize)>,                 // End of active range selection
    pub additional_selections: Vec<((usize, usize), Option<(usize, usize)>)>,  // Ctrl+Click ranges

    // Viewport
    pub scroll_row: usize,
    pub scroll_col: usize,

    // Mode & editing
    pub mode: Mode,
    pub edit_value: String,
    pub edit_cursor: usize,  // Cursor position within edit_value
    pub edit_selection_anchor: Option<usize>,  // Selection start (None = no selection)
    pub edit_original: String,
    pub goto_input: String,
    pub find_input: String,
    pub find_results: Vec<(usize, usize)>,
    pub find_index: usize,

    // Command palette
    pub palette_query: String,
    pub palette_selected: usize,
    search_engine: SearchEngine,
    palette_results: Vec<SearchItem>,
    pub palette_total_results: usize,  // Total matches before truncation
    // Pre-palette state for preview/restore
    palette_pre_selection: (usize, usize),
    palette_pre_selection_end: Option<(usize, usize)>,
    palette_pre_scroll: (usize, usize),
    pub palette_previewing: bool,  // True if user has previewed (Shift+Enter)

    // Clipboard
    pub clipboard: Option<String>,

    // File state
    pub current_file: Option<PathBuf>,
    pub is_modified: bool,
    pub recent_files: Vec<PathBuf>,  // Recently opened files (most recent first)
    pub recent_commands: Vec<CommandId>,  // Recently executed commands (most recent first)

    // UI state
    pub focus_handle: FocusHandle,
    pub status_message: Option<String>,
    pub window_size: Size<Pixels>,

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

    // Formula reference selection state (for pointing mode)
    pub formula_ref_cell: Option<(usize, usize)>,      // Current reference cell (or range start)
    pub formula_ref_end: Option<(usize, usize)>,       // Range end (None = single cell)
    pub formula_ref_start_cursor: usize,               // Cursor position where reference started

    // Highlighted formula references (for existing formulas when editing)
    // Each entry is (start_cell, optional_end_cell_for_ranges)
    pub formula_highlighted_refs: Vec<((usize, usize), Option<(usize, usize)>)>,

    // Formula autocomplete state
    pub autocomplete_visible: bool,
    pub autocomplete_selected: usize,
    pub autocomplete_replace_range: std::ops::Range<usize>,

    // Formula hover documentation state
    pub hover_function: Option<&'static crate::formula_context::FunctionInfo>,

    // Formula view toggle (Ctrl+`)
    pub show_formulas: bool,

    // Inspector panel state
    pub inspector_visible: bool,
    pub inspector_tab: crate::mode::InspectorTab,
    pub inspector_pinned: Option<(usize, usize)>,  // Pinned cell (None = follows selection)
    pub names_filter_query: String,  // Filter query for Names tab

    // Theme
    pub theme: Theme,
    pub theme_preview: Option<Theme>,  // For live preview in picker

    // Cell search cache (generation-based freshness)
    cells_rev: u64,  // Monotonically increasing; bumped on any cell value change
    cell_search_cache: CellSearchCache,
    named_range_usage_cache: NamedRangeUsageCache,

    // Rename symbol state (Ctrl+Shift+R)
    pub rename_original_name: String,      // The named range being renamed
    pub rename_new_name: String,           // User's typed new name
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
    pub name_tooltip_dismissed: bool,            // Has the first-run tooltip been dismissed?
}

/// Cache for cell search results, invalidated by cells_rev
struct CellSearchCache {
    cached_rev: u64,
    entries: Vec<crate::search::CellEntry>,
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
struct NamedRangeUsageCache {
    cached_rev: u64,
    /// Map from lowercase name to usage count
    counts: std::collections::HashMap<String, usize>,
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
        window.focus(&focus_handle, cx);
        let window_size = window.viewport_size();

        // Load saved theme or default
        let settings = Settings::load();
        let theme = settings.theme_id
            .as_ref()
            .and_then(|id| get_theme(id))
            .unwrap_or_else(visigrid_theme);

        Self {
            workbook,
            history: History::new(),
            selected: (0, 0),
            selection_end: None,
            additional_selections: Vec::new(),
            scroll_row: 0,
            scroll_col: 0,
            mode: Mode::Navigation,
            edit_value: String::new(),
            edit_cursor: 0,
            edit_selection_anchor: None,
            edit_original: String::new(),
            goto_input: String::new(),
            find_input: String::new(),
            find_results: Vec::new(),
            find_index: 0,
            palette_query: String::new(),
            palette_selected: 0,
            search_engine: Self::create_search_engine(),
            palette_results: Vec::new(),
            palette_total_results: 0,
            palette_pre_selection: (0, 0),
            palette_pre_selection_end: None,
            palette_pre_scroll: (0, 0),
            palette_previewing: false,
            clipboard: None,
            current_file: None,
            is_modified: false,
            recent_files: Vec::new(),
            recent_commands: Vec::new(),
            focus_handle,
            status_message: None,
            window_size,
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
            formula_ref_cell: None,
            formula_ref_end: None,
            formula_ref_start_cursor: 0,
            formula_highlighted_refs: Vec::new(),
            autocomplete_visible: false,
            autocomplete_selected: 0,
            autocomplete_replace_range: 0..0,
            hover_function: None,
            show_formulas: false,
            inspector_visible: false,
            inspector_tab: crate::mode::InspectorTab::default(),
            inspector_pinned: None,
            names_filter_query: String::new(),
            theme,
            theme_preview: None,
            cells_rev: 1,  // Start at 1 so cache (starting at 0) is immediately stale
            cell_search_cache: CellSearchCache::default(),
            named_range_usage_cache: NamedRangeUsageCache::default(),
            rename_original_name: String::new(),
            rename_new_name: String::new(),
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
            name_tooltip_dismissed: settings.name_tooltip_dismissed,
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
    fn ensure_cell_search_cache_fresh(&mut self) -> &[crate::search::CellEntry] {
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
                self.selected = (row, col);
                self.selection_end = None;
                self.ensure_cell_visible(row, col);
                cx.notify();
            }
            SearchAction::InsertFormula { name, signature } => {
                // Context-aware insertion
                if self.mode.is_formula() || (self.mode.is_editing() && self.edit_value.starts_with('=')) {
                    // Already editing a formula: insert function name at cursor
                    let func_text = format!("{}(", name);
                    let before: String = self.edit_value.chars().take(self.edit_cursor).collect();
                    let after: String = self.edit_value.chars().skip(self.edit_cursor).collect();
                    self.edit_value = format!("{}{}{}", before, func_text, after);
                    self.edit_cursor += func_text.chars().count();
                } else {
                    // Grid navigation: start formula edit with =FUNC(
                    self.edit_original = self.sheet().get_raw(self.selected.0, self.selected.1);
                    self.edit_value = format!("={}(", name);
                    self.edit_cursor = self.edit_value.chars().count();
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
                if let Some(path) = Settings::path() {
                    // Ensure file exists with defaults
                    if !path.exists() {
                        Settings::default().save();
                    }

                    // Open with system default editor
                    #[cfg(target_os = "linux")]
                    let result = std::process::Command::new("xdg-open")
                        .arg(&path)
                        .spawn();

                    #[cfg(target_os = "macos")]
                    let result = std::process::Command::new("open")
                        .arg(&path)
                        .spawn();

                    #[cfg(target_os = "windows")]
                    let result = std::process::Command::new("cmd")
                        .args(["/C", "start", "", &path.display().to_string()])
                        .spawn();

                    let filename = path.file_name()
                        .and_then(|n| n.to_str())
                        .unwrap_or("settings.json");

                    match result {
                        Ok(_) => {
                            self.status_message = Some(format!(
                                "Copied \"{}\" to clipboard — paste into {}",
                                key, filename
                            ));
                        }
                        Err(e) => {
                            self.status_message = Some(format!("Failed to open settings: {}", e));
                        }
                    }
                } else {
                    self.status_message = Some("Could not determine settings path".to_string());
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
                self.selected = (0, 0);
                self.selection_end = None;
                self.scroll_row = 0;
                self.scroll_col = 0;
                cx.notify();
            }
            CommandId::SelectAll => self.select_all(cx),

            // Editing
            CommandId::FillDown => self.fill_down(cx),
            CommandId::FillRight => self.fill_right(cx),
            CommandId::ClearCells => self.delete_selection(cx),
            CommandId::Undo => self.undo(cx),
            CommandId::Redo => self.redo(cx),
            CommandId::AutoSum => self.autosum(cx),

            // Clipboard
            CommandId::Copy => self.copy(cx),
            CommandId::Cut => self.cut(cx),
            CommandId::Paste => self.paste(cx),

            // Formatting
            CommandId::ToggleBold => self.toggle_bold(cx),
            CommandId::ToggleItalic => self.toggle_italic(cx),
            CommandId::ToggleUnderline => self.toggle_underline(cx),
            CommandId::FormatCells => {
                // Open inspector to format tab
                self.inspector_visible = true;
                self.inspector_tab = crate::mode::InspectorTab::Format;
                cx.notify();
            }

            // File
            CommandId::NewFile => self.new_file(cx),
            CommandId::OpenFile => self.open_file(cx),
            CommandId::Save => self.save(cx),
            CommandId::SaveAs => self.save_as(cx),
            CommandId::ExportCsv => self.export_csv(cx),

            // Appearance
            CommandId::SelectTheme => self.show_theme_picker(cx),
            CommandId::SelectFont => self.show_font_picker(cx),

            // View
            CommandId::ToggleInspector => {
                self.inspector_visible = !self.inspector_visible;
                cx.notify();
            }

            // Help
            CommandId::ShowShortcuts => {
                self.status_message = Some("Shortcuts: Ctrl+D Fill Down, Ctrl+R Fill Right, Ctrl+Enter Multi-edit".into());
                cx.notify();
            }
            CommandId::ShowAbout => {
                self.status_message = Some("VisiGrid - A spreadsheet for power users".into());
                cx.notify();
            }
            CommandId::TourNamedRanges => {
                self.show_tour(cx);
            }

            // Sheets
            CommandId::NextSheet => self.next_sheet(cx),
            CommandId::PrevSheet => self.prev_sheet(cx),
            CommandId::AddSheet => self.add_sheet(cx),
        }
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
        self.selected = (0, 0);
        self.selection_end = None;
        self.scroll_row = 0;
        self.scroll_col = 0;
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
            cx.notify();
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
            cx.notify();
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
    pub fn col_x_offset(&self, target_col: usize) -> f32 {
        let mut x = 0.0;
        for col in self.scroll_col..target_col {
            x += self.col_width(col);
        }
        x
    }

    /// Get the Y position of a row's top edge (relative to start of grid, after column header)
    pub fn row_y_offset(&self, target_row: usize) -> f32 {
        let mut y = 0.0;
        for row in self.scroll_row..target_row {
            y += self.row_height(row);
        }
        y
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
        for (start_cell, end_cell) in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = end_cell {
                // Range - check if cell is within
                let (min_row, max_row) = (start_cell.0.min(*end_row), start_cell.0.max(*end_row));
                let (min_col, max_col) = (start_cell.1.min(*end_col), start_cell.1.max(*end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    return true;
                }
            } else {
                // Single cell
                if row == start_cell.0 && col == start_cell.1 {
                    return true;
                }
            }
        }

        false
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
        for (start_cell, end_cell) in &self.formula_highlighted_refs {
            if let Some((end_row, end_col)) = end_cell {
                let (min_row, max_row) = (start_cell.0.min(*end_row), start_cell.0.max(*end_row));
                let (min_col, max_col) = (start_cell.1.min(*end_col), start_cell.1.max(*end_col));
                if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
                    if row == min_row { top = true; }
                    if row == max_row { bottom = true; }
                    if col == min_col { left = true; }
                    if col == max_col { right = true; }
                }
            } else {
                // Single cell - all borders
                if row == start_cell.0 && col == start_cell.1 {
                    top = true; right = true; bottom = true; left = true;
                }
            }
        }

        (top, right, bottom, left)
    }

    /// Calculate visible rows based on window height
    pub fn visible_rows(&self) -> usize {
        let height: f32 = self.window_size.height.into();
        let available_height = height
            - MENU_BAR_HEIGHT
            - FORMULA_BAR_HEIGHT
            - COLUMN_HEADER_HEIGHT
            - STATUS_BAR_HEIGHT;
        let rows = (available_height / CELL_HEIGHT).floor() as usize;
        rows.max(1).min(NUM_ROWS)
    }

    /// Calculate visible columns based on window width
    pub fn visible_cols(&self) -> usize {
        let width: f32 = self.window_size.width.into();
        let available_width = width - HEADER_WIDTH;
        let cols = (available_width / CELL_WIDTH).floor() as usize;
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
        format!("{}{}", Self::col_letter(self.selected.1), self.selected.0 + 1)
    }

    // Navigation
    pub fn move_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.selected = (new_row, new_col);
        self.selection_end = None;  // Clear range selection
        self.additional_selections.clear();  // Clear discontiguous selections

        self.ensure_visible(cx);
    }

    pub fn extend_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selection_end.unwrap_or(self.selected);
        let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
        self.selection_end = Some((new_row, new_col));

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
    fn find_data_boundary(&self, start_row: usize, start_col: usize, dr: i32, dc: i32) -> (usize, usize) {
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

        let (mut row, mut col) = self.selected;
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

        self.selected = (row, col);
        self.selection_end = None;
        self.ensure_visible(cx);
    }

    /// Extend selection to edge of data region (Excel-style Ctrl+Shift+Arrow)
    pub fn extend_jump_selection(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        // Start from current selection end (or selected if no selection)
        let (mut row, mut col) = self.selection_end.unwrap_or(self.selected);
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
        self.selection_end = Some((row, col));
        self.ensure_visible(cx);
    }

    pub fn ensure_visible(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.selection_end.unwrap_or(self.selected);
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // Vertical scroll
        if row < self.scroll_row {
            self.scroll_row = row;
        } else if row >= self.scroll_row + visible_rows {
            self.scroll_row = row - visible_rows + 1;
        }

        // Horizontal scroll
        if col < self.scroll_col {
            self.scroll_col = col;
        } else if col >= self.scroll_col + visible_cols {
            self.scroll_col = col - visible_cols + 1;
        }

        cx.notify();
    }

    pub fn select_cell(&mut self, row: usize, col: usize, extend: bool, cx: &mut Context<Self>) {
        if extend {
            self.selection_end = Some((row, col));
        } else {
            self.selected = (row, col);
            self.selection_end = None;
            self.additional_selections.clear();  // Clear Ctrl+Click selections
        }
        cx.notify();
    }

    /// Ctrl+Click to add/toggle cell in selection (discontiguous selection)
    pub fn ctrl_click_cell(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        // Save current selection to additional_selections
        self.additional_selections.push((self.selected, self.selection_end));
        // Start new selection at clicked cell
        self.selected = (row, col);
        self.selection_end = None;
        cx.notify();
    }

    /// Start drag selection - called on mouse_down
    pub fn start_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        self.selected = (row, col);
        self.selection_end = None;
        self.additional_selections.clear();  // Clear Ctrl+Click selections on new drag
        cx.notify();
    }

    /// Start drag selection with Ctrl held (add to existing selections)
    pub fn start_ctrl_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        self.dragging_selection = true;
        // Save current selection to additional_selections
        self.additional_selections.push((self.selected, self.selection_end));
        // Start new selection at clicked cell
        self.selected = (row, col);
        self.selection_end = None;
        cx.notify();
    }

    /// Continue drag selection - called on mouse_move while dragging
    pub fn continue_drag_selection(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.dragging_selection {
            return;
        }
        // Only update if the cell changed to avoid unnecessary redraws
        if self.selection_end != Some((row, col)) {
            self.selection_end = Some((row, col));
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
        self.selected = (0, 0);
        self.selection_end = Some((NUM_ROWS - 1, NUM_COLS - 1));
        self.additional_selections.clear();  // Clear discontiguous selections
        cx.notify();
    }

    // Scrolling
    pub fn scroll(&mut self, delta_rows: i32, delta_cols: i32, cx: &mut Context<Self>) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();
        let new_row = (self.scroll_row as i32 + delta_rows)
            .max(0)
            .min((NUM_ROWS.saturating_sub(visible_rows)) as i32) as usize;
        let new_col = (self.scroll_col as i32 + delta_cols)
            .max(0)
            .min((NUM_COLS.saturating_sub(visible_cols)) as i32) as usize;

        if new_row != self.scroll_row || new_col != self.scroll_col {
            self.scroll_row = new_row;
            self.scroll_col = new_col;
            cx.notify();
        }
    }

    // Editing
    pub fn start_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;

        // Block editing spill receivers - show message and redirect to parent
        if let Some((parent_row, parent_col)) = self.sheet().get_spill_parent(row, col) {
            let parent_ref = self.cell_ref_at(parent_row, parent_col);
            self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
            cx.notify();
            return;
        }

        self.edit_original = self.sheet().get_raw(row, col);
        self.edit_value = self.edit_original.clone();
        self.edit_cursor = self.edit_value.len();  // Cursor at end

        // Parse and highlight formula references if editing a formula
        if self.edit_value.starts_with('=') || self.edit_value.starts_with('+') {
            self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
        } else {
            self.formula_highlighted_refs.clear();
        }

        self.mode = Mode::Edit;
        cx.notify();
    }

    pub fn start_edit_clear(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() { return; }

        let (row, col) = self.selected;

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
        self.formula_highlighted_refs.clear();  // No formula to highlight
        self.mode = Mode::Edit;
        cx.notify();
    }

    pub fn confirm_edit(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(1, 0, cx);  // Enter moves down
    }

    pub fn confirm_edit_and_move_right(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, 1, cx);  // Tab moves right
    }

    pub fn confirm_edit_and_move_left(&mut self, cx: &mut Context<Self>) {
        self.confirm_edit_and_move(0, -1, cx);  // Shift+Tab moves left
    }

    /// Ctrl+Enter: Confirm edit and apply to ALL selected cells (multi-edit)
    pub fn confirm_edit_in_place(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            // If not editing, start editing
            self.start_edit(cx);
            return;
        }

        let new_value = self.edit_value.clone();
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        let mut changes = Vec::new();

        // Apply to all cells in selection
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                if old_value != new_value {
                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: new_value.clone(),
                    });
                }
                self.sheet_mut().set_value(row, col, &new_value);
            }
        }

        self.history.record_batch(self.sheet_index(), changes);
        self.mode = Mode::Navigation;
        self.edit_value.clear();
        self.edit_original.clear();
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        // Clear formula highlighting state
        self.formula_highlighted_refs.clear();

        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);
        if cell_count > 1 {
            self.status_message = Some(format!("Applied to {} cells", cell_count));
        }
        cx.notify();
    }

    fn confirm_edit_and_move(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            self.start_edit(cx);
            return;
        }

        let (row, col) = self.selected;
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
        self.sheet_mut().set_value(row, col, &new_value);
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
            cx.notify();
        }
    }

    /// Delete selected text and return true if there was a selection
    fn delete_edit_selection(&mut self) -> bool {
        if let Some((start, end)) = self.edit_selection_range() {
            // Convert char positions to byte positions
            let start_byte = self.edit_value.char_indices()
                .nth(start)
                .map(|(i, _)| i)
                .unwrap_or(0);
            let end_byte = self.edit_value.char_indices()
                .nth(end)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len());
            self.edit_value.replace_range(start_byte..end_byte, "");
            self.edit_cursor = start;
            self.edit_selection_anchor = None;
            true
        } else {
            false
        }
    }

    pub fn backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.update_autocomplete(cx);
                return;
            }
            // Otherwise delete char before cursor
            if self.edit_cursor > 0 {
                let byte_idx = self.edit_value.char_indices()
                    .nth(self.edit_cursor - 1)
                    .map(|(i, _)| i)
                    .unwrap_or(0);
                let next_byte_idx = self.edit_value.char_indices()
                    .nth(self.edit_cursor)
                    .map(|(i, _)| i)
                    .unwrap_or(self.edit_value.len());
                self.edit_value.replace_range(byte_idx..next_byte_idx, "");
                self.edit_cursor -= 1;
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                self.update_autocomplete(cx);
            }
        }
    }

    pub fn delete_char(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            // If there's a selection, delete it
            if self.delete_edit_selection() {
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
                cx.notify();
                return;
            }
            // Otherwise delete char at cursor
            let char_count = self.edit_value.chars().count();
            if self.edit_cursor < char_count {
                let byte_idx = self.edit_value.char_indices()
                    .nth(self.edit_cursor)
                    .map(|(i, _)| i)
                    .unwrap_or(self.edit_value.len());
                let next_byte_idx = self.edit_value.char_indices()
                    .nth(self.edit_cursor + 1)
                    .map(|(i, _)| i)
                    .unwrap_or(self.edit_value.len());
                self.edit_value.replace_range(byte_idx..next_byte_idx, "");
                // Update highlighted refs for formulas
                if self.is_formula_content() {
                    self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
                }
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
                }
            }

            // Delete selection if any (replaces selected text)
            self.delete_edit_selection();

            // Find byte index for cursor position
            let byte_idx = self.edit_value.char_indices()
                .nth(self.edit_cursor)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len());
            self.edit_value.insert(byte_idx, c);
            self.edit_cursor += 1;

            // Update highlighted refs for formulas
            if self.is_formula_content() {
                self.formula_highlighted_refs = Self::parse_formula_refs(&self.edit_value);
            }

            // Update autocomplete for formulas
            self.update_autocomplete(cx);
        } else {
            // Start editing with this character
            let (row, col) = self.selected;

            // Block editing spill receivers
            if let Some((parent_row, parent_col)) = self.sheet().get_spill_parent(row, col) {
                let parent_ref = self.cell_ref_at(parent_row, parent_col);
                self.status_message = Some(format!("Cannot edit spill range. Edit {} instead.", parent_ref));
                cx.notify();
                return;
            }

            self.edit_original = self.sheet().get_raw(row, col);
            self.edit_value = c.to_string();
            self.edit_cursor = 1;

            // Enter Formula mode if starting with = or +
            if c == '=' || c == '+' {
                self.mode = Mode::Formula;
                self.formula_ref_cell = None;
                self.formula_ref_end = None;
            } else {
                self.mode = Mode::Edit;
            }

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

    // ========================================================================
    // Formula Mode Reference Selection
    // ========================================================================

    /// Move formula reference with arrow keys (inserts or updates reference)
    pub fn formula_move_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        let (new_row, new_col) = if let Some((row, col)) = self.formula_ref_cell {
            // Move existing reference
            let new_row = (row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let new_col = (col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
            (new_row, new_col)
        } else {
            // Start new reference from the selected cell (editing cell)
            let (sel_row, sel_col) = self.selected;
            let new_row = (sel_row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
            let new_col = (sel_col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;
            (new_row, new_col)
        };

        // Update the reference
        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((new_row, new_col));
        self.formula_ref_end = None;  // Reset range when moving without shift

        // Insert or update the reference in the formula
        self.update_formula_reference(is_new);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Extend formula reference to range with Shift+arrow
    pub fn formula_extend_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Need an existing reference to extend
        let (anchor_row, anchor_col) = match self.formula_ref_cell {
            Some(cell) => cell,
            None => {
                // If no reference yet, start one first
                self.formula_move_ref(dr, dc, cx);
                return;
            }
        };

        // Get current end or use anchor as start
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((anchor_row, anchor_col));

        // Extend from the end position
        let new_row = (end_row as i32 + dr).max(0).min(NUM_ROWS as i32 - 1) as usize;
        let new_col = (end_col as i32 + dc).max(0).min(NUM_COLS as i32 - 1) as usize;

        self.formula_ref_end = Some((new_row, new_col));

        // Update the reference in the formula (not new, updating existing)
        self.update_formula_reference(false);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Insert formula reference on mouse click
    pub fn formula_click_ref(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((row, col));
        self.formula_ref_end = None;

        self.update_formula_reference(is_new);
        cx.notify();
    }

    /// Extend formula reference to range on Shift+click
    pub fn formula_shift_click_ref(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Need an existing reference to extend
        if self.formula_ref_cell.is_none() {
            // No reference yet, just insert single cell
            self.formula_click_ref(row, col, cx);
            return;
        }

        self.formula_ref_end = Some((row, col));
        self.update_formula_reference(false);
        cx.notify();
    }

    /// Extend formula reference to data boundary with Ctrl+Shift+arrow
    pub fn formula_extend_jump_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        // Need an existing reference to extend (or start one)
        let (anchor_row, anchor_col) = match self.formula_ref_cell {
            Some(cell) => cell,
            None => {
                // If no reference yet, start one first with a jump
                self.formula_jump_ref(dr, dc, cx);
                return;
            }
        };

        // Get current end or use anchor as start
        let (end_row, end_col) = self.formula_ref_end.unwrap_or((anchor_row, anchor_col));

        // Jump to data boundary from end position
        let (new_row, new_col) = self.find_data_boundary(end_row, end_col, dr, dc);

        self.formula_ref_end = Some((new_row, new_col));
        self.update_formula_reference(false);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Move formula reference by jumping to data boundary (Ctrl+arrow in formula mode)
    pub fn formula_jump_ref(&mut self, dr: i32, dc: i32, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        let (start_row, start_col) = if let Some((row, col)) = self.formula_ref_cell {
            (row, col)
        } else {
            self.selected
        };

        let (new_row, new_col) = self.find_data_boundary(start_row, start_col, dr, dc);

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((new_row, new_col));
        self.formula_ref_end = None;

        self.update_formula_reference(is_new);
        self.ensure_cell_visible(new_row, new_col);
        cx.notify();
    }

    /// Start formula range drag selection
    pub fn formula_start_drag(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() {
            return;
        }

        let is_new = self.formula_ref_cell.is_none();
        self.formula_ref_cell = Some((row, col));
        self.formula_ref_end = None;
        self.dragging_selection = true;  // Reuse the drag flag

        self.update_formula_reference(is_new);
        cx.notify();
    }

    /// Continue formula range drag selection
    pub fn formula_continue_drag(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        if !self.mode.is_formula() || !self.dragging_selection {
            return;
        }

        if self.formula_ref_cell.is_none() {
            return;
        }

        // Only update if the cell changed
        if self.formula_ref_end != Some((row, col)) {
            self.formula_ref_end = Some((row, col));
            self.update_formula_reference(false);
            cx.notify();
        }
    }

    /// Update the formula string with the current reference
    fn update_formula_reference(&mut self, is_new: bool) {
        let Some((ref_row, ref_col)) = self.formula_ref_cell else {
            return;
        };

        // Build the reference string
        let ref_text = if let Some((end_row, end_col)) = self.formula_ref_end {
            Self::make_range_ref((ref_row, ref_col), (end_row, end_col))
        } else {
            Self::make_cell_ref(ref_row, ref_col)
        };

        if is_new {
            // Insert new reference at cursor
            let byte_idx = self.edit_value.char_indices()
                .nth(self.edit_cursor)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len());

            self.formula_ref_start_cursor = self.edit_cursor;
            self.edit_value.insert_str(byte_idx, &ref_text);
            self.edit_cursor += ref_text.chars().count();
        } else {
            // Replace existing reference (from formula_ref_start_cursor to edit_cursor)
            let start_cursor = self.formula_ref_start_cursor;
            let end_cursor = self.edit_cursor;

            // Convert cursor positions to byte positions
            let start_byte = self.edit_value.char_indices()
                .nth(start_cursor)
                .map(|(i, _)| i)
                .unwrap_or(0);
            let end_byte = self.edit_value.char_indices()
                .nth(end_cursor)
                .map(|(i, _)| i)
                .unwrap_or(self.edit_value.len());

            self.edit_value.replace_range(start_byte..end_byte, &ref_text);
            self.edit_cursor = start_cursor + ref_text.chars().count();
        }
    }

    /// Ensure a cell is visible (scroll if necessary)
    fn ensure_cell_visible(&mut self, row: usize, col: usize) {
        let visible_rows = self.visible_rows();
        let visible_cols = self.visible_cols();

        // Adjust scroll to keep cell visible
        if row < self.scroll_row {
            self.scroll_row = row;
        } else if row >= self.scroll_row + visible_rows {
            self.scroll_row = row.saturating_sub(visible_rows - 1);
        }

        if col < self.scroll_col {
            self.scroll_col = col;
        } else if col >= self.scroll_col + visible_cols {
            self.scroll_col = col.saturating_sub(visible_cols - 1);
        }
    }

    // Cursor movement in edit mode
    pub fn move_edit_cursor_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            self.edit_cursor -= 1;
            self.edit_selection_anchor = None;  // Clear selection
            cx.notify();
        }
    }

    pub fn move_edit_cursor_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let char_count = self.edit_value.chars().count();
            if self.edit_cursor < char_count {
                self.edit_cursor += 1;
                self.edit_selection_anchor = None;  // Clear selection
                cx.notify();
            }
        }
    }

    pub fn move_edit_cursor_home(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            self.edit_cursor = 0;
            self.edit_selection_anchor = None;  // Clear selection
            cx.notify();
        }
    }

    pub fn move_edit_cursor_end(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let char_count = self.edit_value.chars().count();
            if self.edit_cursor < char_count {
                self.edit_cursor = char_count;
                self.edit_selection_anchor = None;  // Clear selection
                cx.notify();
            }
        }
    }

    // Selection variants (Shift+Arrow)
    pub fn select_edit_cursor_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor -= 1;
            cx.notify();
        }
    }

    pub fn select_edit_cursor_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let char_count = self.edit_value.chars().count();
            if self.edit_cursor < char_count {
                if self.edit_selection_anchor.is_none() {
                    self.edit_selection_anchor = Some(self.edit_cursor);
                }
                self.edit_cursor += 1;
                cx.notify();
            }
        }
    }

    pub fn select_edit_cursor_home(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() && self.edit_cursor > 0 {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = 0;
            cx.notify();
        }
    }

    pub fn select_edit_cursor_end(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            let char_count = self.edit_value.chars().count();
            if self.edit_cursor < char_count {
                if self.edit_selection_anchor.is_none() {
                    self.edit_selection_anchor = Some(self.edit_cursor);
                }
                self.edit_cursor = char_count;
                cx.notify();
            }
        }
    }

    // Word navigation helpers
    fn find_word_boundary_left(&self, from: usize) -> usize {
        if from == 0 {
            return 0;
        }
        let chars: Vec<char> = self.edit_value.chars().collect();
        let mut pos = from - 1;
        // Skip whitespace/punctuation
        while pos > 0 && !chars[pos].is_alphanumeric() {
            pos -= 1;
        }
        // Skip word characters
        while pos > 0 && chars[pos - 1].is_alphanumeric() {
            pos -= 1;
        }
        pos
    }

    fn find_word_boundary_right(&self, from: usize) -> usize {
        let chars: Vec<char> = self.edit_value.chars().collect();
        let len = chars.len();
        if from >= len {
            return len;
        }
        let mut pos = from;
        // Skip current word characters
        while pos < len && chars[pos].is_alphanumeric() {
            pos += 1;
        }
        // Skip whitespace/punctuation
        while pos < len && !chars[pos].is_alphanumeric() {
            pos += 1;
        }
        pos
    }

    // Ctrl+Arrow word navigation
    pub fn move_edit_cursor_word_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_cursor = self.find_word_boundary_left(self.edit_cursor);
            self.edit_selection_anchor = None;
            cx.notify();
        }
    }

    pub fn move_edit_cursor_word_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_cursor = self.find_word_boundary_right(self.edit_cursor);
            self.edit_selection_anchor = None;
            cx.notify();
        }
    }

    // Ctrl+Shift+Arrow word selection
    pub fn select_edit_cursor_word_left(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = self.find_word_boundary_left(self.edit_cursor);
            cx.notify();
        }
    }

    pub fn select_edit_cursor_word_right(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            if self.edit_selection_anchor.is_none() {
                self.edit_selection_anchor = Some(self.edit_cursor);
            }
            self.edit_cursor = self.find_word_boundary_right(self.edit_cursor);
            cx.notify();
        }
    }

    // Get current selection range (start, end) or None
    pub fn edit_selection_range(&self) -> Option<(usize, usize)> {
        self.edit_selection_anchor.map(|anchor| {
            (anchor.min(self.edit_cursor), anchor.max(self.edit_cursor))
        })
    }

    // Select all text in edit mode
    pub fn select_all_edit(&mut self, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            self.edit_selection_anchor = Some(0);
            self.edit_cursor = self.edit_value.chars().count();
            cx.notify();
        }
    }

    /// F4: Cycle cell reference at cursor through A1 → $A$1 → A$1 → $A1 → A1
    pub fn cycle_reference(&mut self, cx: &mut Context<Self>) {
        if !self.mode.is_editing() {
            return;
        }

        // Cell reference pattern: optional $ + column letters + optional $ + row numbers
        let re = regex::Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

        // Find cursor byte position
        let cursor_byte = self.edit_value.char_indices()
            .nth(self.edit_cursor)
            .map(|(i, _)| i)
            .unwrap_or(self.edit_value.len());

        // Find reference at or near cursor
        let mut best_match: Option<(usize, usize, regex::Captures)> = None;

        for caps in re.captures_iter(&self.edit_value) {
            let m = caps.get(0).unwrap();
            let start = m.start();
            let end = m.end();

            // Check if cursor is within or immediately after this reference
            if cursor_byte >= start && cursor_byte <= end {
                best_match = Some((start, end, caps));
                break;
            }
            // Also check if cursor is just before the reference (user may have cursor at start)
            if cursor_byte == start {
                best_match = Some((start, end, caps));
                break;
            }
        }

        // If no direct match, find the nearest reference before cursor
        if best_match.is_none() {
            for caps in re.captures_iter(&self.edit_value) {
                let m = caps.get(0).unwrap();
                let start = m.start();
                let end = m.end();

                if end <= cursor_byte {
                    best_match = Some((start, end, caps));
                }
            }
        }

        if let Some((start, end, caps)) = best_match {
            let col_dollar = caps.get(1).map(|m| m.as_str()).unwrap_or("");
            let col_letters = caps.get(2).map(|m| m.as_str()).unwrap_or("");
            let row_dollar = caps.get(3).map(|m| m.as_str()).unwrap_or("");
            let row_numbers = caps.get(4).map(|m| m.as_str()).unwrap_or("");

            // Determine current state and cycle to next
            // State 0: A1 (relative, relative)
            // State 1: $A$1 (absolute, absolute)
            // State 2: A$1 (relative col, absolute row)
            // State 3: $A1 (absolute col, relative row)
            let current_state = match (col_dollar.is_empty(), row_dollar.is_empty()) {
                (true, true) => 0,    // A1
                (false, false) => 1,  // $A$1
                (true, false) => 2,   // A$1
                (false, true) => 3,   // $A1
            };

            let next_state = (current_state + 1) % 4;

            let new_ref = match next_state {
                0 => format!("{}{}", col_letters, row_numbers),           // A1
                1 => format!("${}${}", col_letters, row_numbers),         // $A$1
                2 => format!("{}${}", col_letters, row_numbers),          // A$1
                3 => format!("${}{}", col_letters, row_numbers),          // $A1
                _ => unreachable!(),
            };

            // Replace the reference in edit_value
            let old_ref_chars = end - start;  // For ASCII cell refs, byte len == char count
            self.edit_value.replace_range(start..end, &new_ref);
            let new_ref_chars = new_ref.chars().count();

            // Adjust cursor if it was after or within the replaced region
            let start_char = self.edit_value[..start].chars().count();

            if self.edit_cursor > start_char {
                // Cursor was within or after the reference
                if self.edit_cursor <= start_char + old_ref_chars {
                    // Cursor was within reference - move to end of new reference
                    self.edit_cursor = start_char + new_ref_chars;
                } else {
                    // Cursor was after reference - adjust by length difference
                    let diff = new_ref_chars as i32 - old_ref_chars as i32;
                    self.edit_cursor = (self.edit_cursor as i32 + diff) as usize;
                }
            }

            cx.notify();
        }
    }

    // Clipboard
    pub fn copy(&mut self, cx: &mut Context<Self>) {
        // If editing, copy selected text (or all if no selection)
        if self.mode.is_editing() {
            let text = if let Some((start, end)) = self.edit_selection_range() {
                self.edit_value.chars().skip(start).take(end - start).collect()
            } else {
                self.edit_value.clone()
            };
            self.clipboard = Some(text.clone());
            cx.write_to_clipboard(ClipboardItem::new_string(text));
            self.status_message = Some("Copied to clipboard".to_string());
            cx.notify();
            return;
        }

        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Build tab-separated values for clipboard
        let mut text = String::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                if col > min_col {
                    text.push('\t');
                }
                text.push_str(&self.sheet().get_display(row, col));
            }
            if row < max_row {
                text.push('\n');
            }
        }

        self.clipboard = Some(text.clone());
        cx.write_to_clipboard(ClipboardItem::new_string(text));
        self.status_message = Some("Copied to clipboard".to_string());
        cx.notify();
    }

    pub fn cut(&mut self, cx: &mut Context<Self>) {
        self.copy(cx);

        // Clear the selected cells and record history
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let mut changes = Vec::new();
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().set_value(row, col, "");
            }
        }
        self.history.record_batch(self.sheet_index(), changes);
        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;
        self.status_message = Some("Cut to clipboard".to_string());
        cx.notify();
    }

    pub fn paste(&mut self, cx: &mut Context<Self>) {
        // If editing, paste into the edit buffer instead
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.clipboard.clone()
        };

        if let Some(text) = text {
            let (start_row, start_col) = self.selected;
            let mut changes = Vec::new();

            // Parse tab-separated values
            for (row_offset, line) in text.lines().enumerate() {
                for (col_offset, value) in line.split('\t').enumerate() {
                    let row = start_row + row_offset;
                    let col = start_col + col_offset;
                    if row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet().get_raw(row, col);
                        let new_value = value.to_string();
                        if old_value != new_value {
                            changes.push(CellChange {
                                row, col, old_value, new_value,
                            });
                        }
                        self.sheet_mut().set_value(row, col, value);
                    }
                }
            }

            self.history.record_batch(self.sheet_index(), changes);
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;
            self.status_message = Some("Pasted from clipboard".to_string());
            cx.notify();
        }
    }

    /// Paste clipboard text into the edit buffer (when in editing mode)
    pub fn paste_into_edit(&mut self, cx: &mut Context<Self>) {
        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.clipboard.clone()
        };

        if let Some(text) = text {
            // Only take first line if multi-line, and trim whitespace
            let text = text.lines().next().unwrap_or("").trim();
            if !text.is_empty() {
                // Append to edit value (no cursor position tracking yet)
                self.edit_value.push_str(text);
                self.status_message = Some(format!("Pasted: {}", text));
                cx.notify();
            }
        }
    }

    pub fn delete_selection(&mut self, cx: &mut Context<Self>) {
        let mut changes = Vec::new();
        let mut skipped_spill_receivers = false;

        // Delete from all selection ranges (including discontiguous Ctrl+Click selections)
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            // Only get cells that actually have data (efficient for large selections)
            let cells_to_delete = self.sheet().cells_in_range(min_row, max_row, min_col, max_col);

            for (row, col) in cells_to_delete {
                // Skip spill receivers - only the parent formula can be deleted
                if self.sheet().is_spill_receiver(row, col) {
                    skipped_spill_receivers = true;
                    continue;
                }

                let old_value = self.sheet().get_raw(row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row, col, old_value, new_value: String::new(),
                    });
                }
                self.sheet_mut().clear_cell(row, col);
            }
        }

        let had_changes = !changes.is_empty();
        if had_changes {
            self.history.record_batch(self.sheet_index(), changes);
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;
        }

        if skipped_spill_receivers && !had_changes {
            self.status_message = Some("Cannot delete spill range. Delete the parent formula instead.".to_string());
        }

        cx.notify();
    }

    // Undo/Redo
    pub fn undo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.undo() {
            match entry.action {
                UndoAction::Values { sheet_index, changes } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for change in changes {
                            sheet.set_value(change.row, change.col, &change.old_value);
                        }
                    }
                    self.bump_cells_rev();  // Invalidate cell search cache
                    self.status_message = Some("Undo".to_string());
                }
                UndoAction::Format { sheet_index, patches, description, .. } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for patch in patches {
                            sheet.set_format(patch.row, patch.col, patch.before);
                        }
                    }
                    self.status_message = Some(format!("Undo: {}", description));
                }
                UndoAction::NamedRangeDeleted { named_range } => {
                    // Restore the deleted named range
                    let name = named_range.name.clone();
                    let _ = self.workbook.named_ranges_mut().set(named_range);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: restored '{}'", name));
                }
                UndoAction::NamedRangeCreated { name } => {
                    // Delete the created named range
                    self.workbook.delete_named_range(&name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: removed '{}'", name));
                }
                UndoAction::NamedRangeRenamed { old_name, new_name } => {
                    // Rename back to original name
                    let _ = self.workbook.rename_named_range(&new_name, &old_name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: renamed back to '{}'", old_name));
                }
                UndoAction::NamedRangeDescriptionChanged { name, old_description, .. } => {
                    // Restore the old description
                    let _ = self.workbook.named_ranges_mut().set_description(&name, old_description.clone());
                    self.status_message = Some(format!("Undo: description of '{}'", name));
                }
            }
            self.is_modified = true;
            cx.notify();
        }
    }

    pub fn redo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.redo() {
            match entry.action {
                UndoAction::Values { sheet_index, changes } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for change in changes {
                            sheet.set_value(change.row, change.col, &change.new_value);
                        }
                    }
                    self.bump_cells_rev();  // Invalidate cell search cache
                    self.status_message = Some("Redo".to_string());
                }
                UndoAction::Format { sheet_index, patches, description, .. } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for patch in patches {
                            sheet.set_format(patch.row, patch.col, patch.after);
                        }
                    }
                    self.status_message = Some(format!("Redo: {}", description));
                }
                UndoAction::NamedRangeDeleted { named_range } => {
                    // Re-delete the named range
                    let name = named_range.name.clone();
                    self.workbook.delete_named_range(&name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: deleted '{}'", name));
                }
                UndoAction::NamedRangeCreated { ref name } => {
                    // Re-create is not possible without the original data
                    // This shouldn't happen in practice (create followed by undo-redo)
                    self.status_message = Some(format!("Redo: recreate '{}' not supported", name));
                }
                UndoAction::NamedRangeRenamed { old_name, new_name } => {
                    // Rename again to new name
                    let _ = self.workbook.rename_named_range(&old_name, &new_name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: renamed to '{}'", new_name));
                }
                UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                    // Apply the new description
                    let _ = self.workbook.named_ranges_mut().set_description(&name, new_description.clone());
                    self.status_message = Some(format!("Redo: description of '{}'", name));
                }
            }
            self.is_modified = true;
            cx.notify();
        }
    }

    // Selection helpers
    pub fn selection_range(&self) -> ((usize, usize), (usize, usize)) {
        let start = self.selected;
        let end = self.selection_end.unwrap_or(start);
        let min_row = start.0.min(end.0);
        let max_row = start.0.max(end.0);
        let min_col = start.1.min(end.1);
        let max_col = start.1.max(end.1);
        ((min_row, min_col), (max_row, max_col))
    }

    pub fn is_selected(&self, row: usize, col: usize) -> bool {
        // Check active selection
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        if row >= min_row && row <= max_row && col >= min_col && col <= max_col {
            return true;
        }
        // Check additional selections (Ctrl+Click ranges)
        for (start, end) in &self.additional_selections {
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

    /// Get all selection ranges (for operations that apply to all selected cells)
    pub fn all_selection_ranges(&self) -> Vec<((usize, usize), (usize, usize))> {
        let mut ranges = Vec::new();
        // Add active selection
        ranges.push(self.selection_range());
        // Add additional selections
        for (start, end) in &self.additional_selections {
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

    // Go To cell dialog
    pub fn show_goto(&mut self, cx: &mut Context<Self>) {
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
                self.selected = (row, col);
                self.selection_end = None;
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
    fn parse_cell_ref(input: &str) -> Option<(usize, usize)> {
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

    /// Parse all cell references from a formula and return as (start, optional_end) tuples
    /// Handles both single cells (A1) and ranges (A1:B5)
    fn parse_formula_refs(formula: &str) -> Vec<((usize, usize), Option<(usize, usize)>)> {
        if !formula.starts_with('=') {
            return Vec::new();
        }

        let tokens = tokenize_for_highlight(formula);
        let mut refs = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let (range, token_type) = &tokens[i];

            if *token_type == TokenType::CellRef {
                let cell_text = &formula[range.clone()];
                // Strip any $ signs for absolute references
                let cell_text_clean: String = cell_text.chars().filter(|c| *c != '$').collect();

                if let Some(start_cell) = Self::parse_cell_ref(&cell_text_clean) {
                    // Check if next tokens form a range (: followed by CellRef)
                    if i + 2 < tokens.len() {
                        let (_, next_type) = &tokens[i + 1];
                        let (range2, next_next_type) = &tokens[i + 2];

                        if *next_type == TokenType::Colon && *next_next_type == TokenType::CellRef {
                            let end_text = &formula[range2.clone()];
                            let end_text_clean: String = end_text.chars().filter(|c| *c != '$').collect();

                            if let Some(end_cell) = Self::parse_cell_ref(&end_text_clean) {
                                refs.push((start_cell, Some(end_cell)));
                                i += 3;  // Skip the whole range
                                continue;
                            }
                        }
                    }
                    // Single cell reference
                    refs.push((start_cell, None));
                }
            }
            i += 1;
        }

        refs
    }

    // Find in cells
    pub fn show_find(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Find;
        self.find_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        cx.notify();
    }

    pub fn hide_find(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn find_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == Mode::Find {
            self.find_input.push(c);
            self.perform_find(cx);
        }
    }

    pub fn find_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::Find {
            self.find_input.pop();
            self.perform_find(cx);
        }
    }

    fn perform_find(&mut self, cx: &mut Context<Self>) {
        self.find_results.clear();
        self.find_index = 0;

        if self.find_input.is_empty() {
            self.status_message = None;
            cx.notify();
            return;
        }

        let query = self.find_input.to_lowercase();

        // Search through all populated cells
        let cell_positions: Vec<_> = self.sheet().cells_iter()
            .map(|(&pos, _)| pos)
            .collect();

        for (row, col) in cell_positions {
            let display = self.sheet().get_display(row, col);
            if display.to_lowercase().contains(&query) {
                self.find_results.push((row, col));
            }
        }

        // Sort results by row, then column
        self.find_results.sort();

        if !self.find_results.is_empty() {
            self.jump_to_find_result(cx);
            self.status_message = Some(format!(
                "Found {} match{}",
                self.find_results.len(),
                if self.find_results.len() == 1 { "" } else { "es" }
            ));
        } else {
            self.status_message = Some("No matches found".to_string());
        }
        cx.notify();
    }

    pub fn find_next(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        self.find_index = (self.find_index + 1) % self.find_results.len();
        self.jump_to_find_result(cx);
    }

    pub fn find_prev(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        if self.find_index == 0 {
            self.find_index = self.find_results.len() - 1;
        } else {
            self.find_index -= 1;
        }
        self.jump_to_find_result(cx);
    }

    fn jump_to_find_result(&mut self, cx: &mut Context<Self>) {
        if let Some(&(row, col)) = self.find_results.get(self.find_index) {
            self.selected = (row, col);
            self.selection_end = None;
            self.ensure_visible(cx);
            self.status_message = Some(format!(
                "Match {} of {}",
                self.find_index + 1,
                self.find_results.len()
            ));
        }
    }

    // Command Palette
    pub fn toggle_palette(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::Command {
            self.hide_palette(cx);
        } else {
            self.show_palette(cx);
        }
    }

    pub fn show_palette(&mut self, cx: &mut Context<Self>) {
        // Save pre-palette state for restore on Esc
        self.palette_pre_selection = self.selected;
        self.palette_pre_selection_end = self.selection_end;
        self.palette_pre_scroll = (self.scroll_row, self.scroll_col);
        self.palette_previewing = false;

        self.mode = Mode::Command;
        self.palette_query.clear();
        self.palette_selected = 0;
        self.update_palette_results();
        cx.notify();
    }

    /// Show cells that reference the given cell (Find References - Shift+F12)
    /// Opens the command palette populated with dependent cells
    pub fn show_references(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        use crate::search::{ReferenceEntry, ReferencesProvider, SearchProvider, SearchQuery};
        use visigrid_engine::formula::parser::{parse, extract_cell_refs};
        use visigrid_engine::cell::CellValue;

        // Get the cell reference for display
        let source_cell_ref = self.cell_ref_at(row, col);

        // Find all cells that reference this cell (dependents)
        let mut references = Vec::new();
        for (&(cell_row, cell_col), cell) in self.sheet().cells_iter() {
            if let CellValue::Formula { source, .. } = &cell.value {
                if let Ok(expr) = parse(source) {
                    let refs = extract_cell_refs(&expr);
                    if refs.contains(&(row, col)) {
                        let cell_ref = self.cell_ref_at(cell_row, cell_col);
                        references.push(ReferenceEntry::new(
                            cell_row,
                            cell_col,
                            cell_ref,
                            source.clone(),
                        ));
                    }
                }
            }
        }

        if references.is_empty() {
            self.status_message = Some(format!("No cells reference {}", source_cell_ref));
            cx.notify();
            return;
        }

        // Sort references by cell position for predictable order
        references.sort_by_key(|r| (r.row, r.col));

        // Save pre-palette state for restore on Esc
        self.palette_pre_selection = self.selected;
        self.palette_pre_selection_end = self.selection_end;
        self.palette_pre_scroll = (self.scroll_row, self.scroll_col);
        self.palette_previewing = false;

        // Build results using the ReferencesProvider
        let provider = ReferencesProvider::new(source_cell_ref.clone(), references);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette with references
        self.mode = Mode::Command;
        self.palette_query = format!("References to {}", source_cell_ref);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    /// Show cells that the given cell references (Go to Precedents - F12)
    /// Opens the command palette populated with precedent cells
    pub fn show_precedents(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        use crate::search::{PrecedentEntry, PrecedentsProvider, SearchProvider, SearchQuery};
        use visigrid_engine::formula::parser::{parse, extract_cell_refs};

        // Get the cell reference for display
        let source_cell_ref = self.cell_ref_at(row, col);

        // Get the raw value of this cell
        let raw_value = self.sheet().get_raw(row, col);

        // Only formulas have precedents
        if !raw_value.starts_with('=') {
            self.status_message = Some(format!("{} is not a formula", source_cell_ref));
            cx.notify();
            return;
        }

        // Parse the formula to extract referenced cells
        let precedent_coords = if let Ok(expr) = parse(&raw_value) {
            let mut refs = extract_cell_refs(&expr);
            refs.sort();
            refs.dedup();
            refs
        } else {
            Vec::new()
        };

        if precedent_coords.is_empty() {
            self.status_message = Some(format!("{} has no cell references", source_cell_ref));
            cx.notify();
            return;
        }

        // Build precedent entries with display values
        let precedents: Vec<PrecedentEntry> = precedent_coords
            .iter()
            .map(|&(r, c)| {
                let cell_ref = self.cell_ref_at(r, c);
                let display = self.sheet().get_display(r, c);
                PrecedentEntry::new(r, c, cell_ref, display)
            })
            .collect();

        // Save pre-palette state for restore on Esc
        self.palette_pre_selection = self.selected;
        self.palette_pre_selection_end = self.selection_end;
        self.palette_pre_scroll = (self.scroll_row, self.scroll_col);
        self.palette_previewing = false;

        // Build results using the PrecedentsProvider
        let provider = PrecedentsProvider::new(source_cell_ref.clone(), precedents);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette with precedents
        self.mode = Mode::Command;
        self.palette_query = format!("Precedents of {}", source_cell_ref);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    /// Extract the identifier (word) at the cursor position in edit_value
    /// Returns the identifier and its range in the edit_value
    fn identifier_at_cursor(&self) -> Option<(String, usize, usize)> {
        if self.edit_value.is_empty() {
            return None;
        }

        let chars: Vec<char> = self.edit_value.chars().collect();
        let cursor = self.edit_cursor.min(chars.len());

        // Find the start of the identifier (scan backwards)
        let mut start = cursor;
        while start > 0 {
            let c = chars[start - 1];
            if c.is_alphanumeric() || c == '_' {
                start -= 1;
            } else {
                break;
            }
        }

        // Find the end of the identifier (scan forwards)
        let mut end = cursor;
        while end < chars.len() {
            let c = chars[end];
            if c.is_alphanumeric() || c == '_' {
                end += 1;
            } else {
                break;
            }
        }

        if start == end {
            return None;
        }

        let identifier: String = chars[start..end].iter().collect();
        Some((identifier, start, end))
    }

    /// Get the named range at the cursor position in edit_value (if any)
    pub fn named_range_at_cursor(&self) -> Option<String> {
        let (identifier, _, _) = self.identifier_at_cursor()?;

        // Check if this identifier is a named range
        if self.workbook.get_named_range(&identifier).is_some() {
            Some(identifier)
        } else {
            None
        }
    }

    /// Go to the definition of a named range (F12 on named range in formula)
    pub fn go_to_named_range_definition(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;

        // Extract data from named range before mutable borrows
        let target_info = self.workbook.get_named_range(name).map(|nr| {
            let (row, col) = match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => (*row, *col),
                NamedRangeTarget::Range { start_row, start_col, .. } => (*start_row, *start_col),
            };
            (row, col, nr.reference_string())
        });

        if let Some((row, col, ref_str)) = target_info {
            // Exit edit mode and jump to the named range's target
            self.mode = Mode::Navigation;
            self.edit_value.clear();
            self.edit_cursor = 0;
            self.selected = (row, col);
            self.selection_end = None;
            self.ensure_cell_visible(row, col);
            self.status_message = Some(format!("'{}' → {}", name, ref_str));
            cx.notify();
        } else {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
        }
    }

    /// Show all formulas that use a named range (Shift+F12 on named range)
    pub fn show_named_range_references(&mut self, name: &str, cx: &mut Context<Self>) {
        use crate::search::{ReferenceEntry, ReferencesProvider, SearchProvider, SearchQuery};
        use visigrid_engine::cell::CellValue;

        let name_upper = name.to_uppercase();

        // Find all cells that use this named range
        let mut references = Vec::new();
        for (&(cell_row, cell_col), cell) in self.sheet().cells_iter() {
            if let CellValue::Formula { source, .. } = &cell.value {
                // Check if formula references this named range (word-boundary aware)
                if self.formula_references_name(source, &name_upper) {
                    let cell_ref = self.cell_ref_at(cell_row, cell_col);
                    references.push(ReferenceEntry::new(
                        cell_row,
                        cell_col,
                        cell_ref,
                        source.clone(),
                    ));
                }
            }
        }

        if references.is_empty() {
            self.status_message = Some(format!("No cells reference '{}'", name));
            cx.notify();
            return;
        }

        // Sort references by cell position
        references.sort_by_key(|r| (r.row, r.col));

        // Save pre-palette state
        self.palette_pre_selection = self.selected;
        self.palette_pre_selection_end = self.selection_end;
        self.palette_pre_scroll = (self.scroll_row, self.scroll_col);
        self.palette_previewing = false;

        // Build results
        let provider = ReferencesProvider::new(format!("${}", name), references);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette with references
        self.mode = Mode::Command;
        self.palette_query = format!("References to ${}", name);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    pub fn hide_palette(&mut self, cx: &mut Context<Self>) {
        // Restore pre-palette state (Esc behavior)
        if self.palette_previewing {
            self.selected = self.palette_pre_selection;
            self.selection_end = self.palette_pre_selection_end;
            self.scroll_row = self.palette_pre_scroll.0;
            self.scroll_col = self.palette_pre_scroll.1;
        }

        self.mode = Mode::Navigation;
        self.palette_query.clear();
        self.palette_selected = 0;
        self.palette_results.clear();
        self.palette_previewing = false;
        cx.notify();
    }

    /// Preview the selected action (Shift+Enter) - jump/scroll but keep palette open
    pub fn palette_preview(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            self.palette_previewing = true;

            // Preview based on action type
            match &item.action {
                SearchAction::JumpToCell { row, col } => {
                    self.selected = (*row, *col);
                    self.selection_end = None;
                    self.ensure_cell_visible(*row, *col);
                }
                SearchAction::RunCommand(cmd) => {
                    // For commands, show a preview hint in status bar
                    self.status_message = Some(format!("Preview: {}", cmd.name()));
                }
                SearchAction::InsertFormula { name, signature } => {
                    // Show full signature for formula preview
                    self.status_message = Some(format!("{}: {}", name, signature));
                }
                SearchAction::OpenFile(path) => {
                    // Show full path for file preview
                    self.status_message = Some(format!("Open: {}", path.display()));
                }
                _ => {
                    // Other actions: show action description
                    self.status_message = Some(format!("Preview: {}", item.title));
                }
            }
            cx.notify();
        }
    }

    /// Update palette search results based on current query
    fn update_palette_results(&mut self) {
        use crate::search::{SearchQuery, SearchProvider, CellSearchProvider, RecentFilesProvider, NamedRangeSearchProvider, NamedRangeEntry};
        use visigrid_engine::named_range::NamedRangeTarget;

        // Clone query string first to avoid borrow conflicts with cache refresh
        let query_str = self.palette_query.clone();
        let query = SearchQuery::parse(&query_str);
        let mut results = self.search_engine.search(&query_str, 12);

        // Add recent files when there's no prefix (commands + recent files)
        if query.prefix.is_none() && !self.recent_files.is_empty() {
            let provider = RecentFilesProvider::new(self.recent_files.clone());
            let recent_results = provider.search(&query, 10);
            results.extend(recent_results);
        }

        // Add cell search with @ prefix (uses generation-based cache for freshness)
        if query.prefix == Some('@') {
            // Ensure cache is fresh (rebuilds only if cells_rev changed)
            self.ensure_cell_search_cache_fresh();

            // Search over cached entries
            let provider = CellSearchProvider::new(self.cell_search_cache.entries.clone());
            let cell_results = provider.search(&query, 50);
            results.extend(cell_results);
        }

        // Add named range search with $ prefix
        if query.prefix == Some('$') {
            let entries: Vec<NamedRangeEntry> = self.workbook.list_named_ranges()
                .into_iter()
                .map(|nr| {
                    let (row, col) = match &nr.target {
                        NamedRangeTarget::Cell { row, col, .. } => (*row, *col),
                        NamedRangeTarget::Range { start_row, start_col, .. } => (*start_row, *start_col),
                    };
                    NamedRangeEntry::new(
                        nr.name.clone(),
                        nr.reference_string(),
                        nr.description.clone(),
                        row,
                        col,
                    )
                })
                .collect();

            let provider = NamedRangeSearchProvider::new(entries);
            let named_results = provider.search(&query, 50);
            results.extend(named_results);
        }

        // Apply recency boost to commands (makes the palette feel "adaptive")
        for result in &mut results {
            if let SearchAction::RunCommand(cmd) = &result.action {
                let boost = self.command_recency_score(cmd);
                result.score += boost;
            }
        }

        // Apply unified sorting: score (desc) → kind priority (asc) → title (asc)
        results.sort_by(|a, b| {
            match b.score.partial_cmp(&a.score) {
                Some(std::cmp::Ordering::Equal) | None => {}
                Some(ord) => return ord,
            }
            match a.kind.priority().cmp(&b.kind.priority()) {
                std::cmp::Ordering::Equal => {}
                ord => return ord,
            }
            a.title.cmp(&b.title)
        });

        // Track total before truncation
        self.palette_total_results = results.len();
        results.truncate(12);

        self.palette_results = results;
    }

    /// Get palette results for rendering (borrows immutably)
    pub fn palette_results(&self) -> &[SearchItem] {
        &self.palette_results
    }

    /// Track a command as recently used (for scoring boost)
    fn add_recent_command(&mut self, cmd: CommandId) {
        const MAX_RECENT_COMMANDS: usize = 20;

        // Remove if already present (we'll add to front)
        self.recent_commands.retain(|c| c != &cmd);

        // Add to front
        self.recent_commands.insert(0, cmd);

        // Limit size
        self.recent_commands.truncate(MAX_RECENT_COMMANDS);
    }

    /// Check if a command was recently used (returns recency score 0.0-1.0)
    pub fn command_recency_score(&self, cmd: &CommandId) -> f32 {
        if let Some(pos) = self.recent_commands.iter().position(|c| c == cmd) {
            // More recent = higher score, decays with position
            // Position 0 (most recent) = 0.15 boost, position 19 = ~0.0 boost
            0.15 * (1.0 - (pos as f32 / 20.0))
        } else {
            0.0
        }
    }

    pub fn palette_up(&mut self, cx: &mut Context<Self>) {
        if self.palette_selected > 0 {
            self.palette_selected -= 1;
            cx.notify();
        }
    }

    pub fn palette_down(&mut self, cx: &mut Context<Self>) {
        let count = self.palette_results.len();
        if self.palette_selected + 1 < count {
            self.palette_selected += 1;
            cx.notify();
        }
    }

    pub fn palette_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.palette_query.push(c);
        self.palette_selected = 0;  // Reset selection on filter change
        self.update_palette_results();
        cx.notify();
    }

    pub fn palette_backspace(&mut self, cx: &mut Context<Self>) {
        // Retain prefix character if it's the only thing left
        // Prefixes: >, =, @, :, #
        let query_len = self.palette_query.chars().count();
        if query_len == 1 {
            let first_char = self.palette_query.chars().next().unwrap();
            if matches!(first_char, '>' | '=' | '@' | ':' | '#') {
                // Don't remove the prefix - user stays in that search mode
                return;
            }
        }
        self.palette_query.pop();
        self.palette_selected = 0;  // Reset selection on filter change
        self.update_palette_results();
        cx.notify();
    }

    pub fn palette_execute(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            // Clear palette state - don't restore since we're executing
            self.palette_query.clear();
            self.palette_selected = 0;
            self.palette_results.clear();
            self.palette_previewing = false;  // Clear previewing flag

            self.dispatch_action(item.action, cx);
            // Only return to Navigation if action didn't change mode
            if self.mode == Mode::Command {
                self.mode = Mode::Navigation;
            }
            cx.notify();
        } else {
            self.hide_palette(cx);
        }
    }

    /// Execute secondary action (Ctrl+Enter) for selected palette item
    pub fn palette_execute_secondary(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            if let Some(secondary) = item.secondary_action {
                // Clear palette state
                self.palette_query.clear();
                self.palette_selected = 0;
                self.palette_results.clear();
                self.palette_previewing = false;

                self.dispatch_action(secondary, cx);
                if self.mode == Mode::Command {
                    self.mode = Mode::Navigation;
                }
                cx.notify();
            } else {
                // No secondary action - show hint
                self.status_message = Some("No secondary action available".to_string());
                cx.notify();
            }
        }
    }

    // Font Picker
    pub fn show_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::FontPicker;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn hide_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_up(&mut self, cx: &mut Context<Self>) {
        if self.font_picker_selected > 0 {
            self.font_picker_selected -= 1;
            cx.notify();
        }
    }

    pub fn font_picker_down(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if self.font_picker_selected + 1 < filtered.len() {
            self.font_picker_selected += 1;
            cx.notify();
        }
    }

    pub fn font_picker_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.font_picker_query.push(c);
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_backspace(&mut self, cx: &mut Context<Self>) {
        self.font_picker_query.pop();
        self.font_picker_selected = 0;
        cx.notify();
    }

    pub fn font_picker_execute(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if let Some(font_name) = filtered.get(self.font_picker_selected) {
            let font = font_name.clone();
            self.apply_font_to_selection(&font, cx);
        }
        self.hide_font_picker(cx);
    }

    /// Filter available fonts by query
    pub fn filter_fonts(&self) -> Vec<String> {
        if self.font_picker_query.is_empty() {
            return self.available_fonts.clone();
        }
        let query_lower = self.font_picker_query.to_lowercase();
        self.available_fonts
            .iter()
            .filter(|f| f.to_lowercase().contains(&query_lower))
            .cloned()
            .collect()
    }

    /// Apply font to all cells in current selection
    pub fn apply_font_to_selection(&mut self, font_name: &str, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let font = if font_name.is_empty() { None } else { Some(font_name.to_string()) };

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                self.sheet_mut().set_font_family(row, col, font.clone());
            }
        }

        self.is_modified = true;
        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);
        self.status_message = Some(format!("Applied font '{}' to {} cell(s)", font_name, cell_count));
        cx.notify();
    }

    /// Clear font from selection (reset to default)
    pub fn clear_font_from_selection(&mut self, cx: &mut Context<Self>) {
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                self.sheet_mut().set_font_family(row, col, None);
            }
        }

        self.is_modified = true;
        self.status_message = Some("Cleared font from selection".to_string());
        cx.notify();
    }

    // Theme Picker
    pub fn show_theme_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::ThemePicker;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        self.theme_preview = None;
        cx.notify();
    }

    pub fn hide_theme_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        self.theme_preview = None;
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
            // Persist theme selection
            let mut settings = Settings::load();
            settings.theme_id = Some(theme.meta.id.to_string());
            settings.save();
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
        self.mode = Mode::About;
        cx.notify();
    }

    pub fn hide_about(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    // Inspector panel methods
    pub fn toggle_inspector_pin(&mut self, cx: &mut Context<Self>) {
        if self.inspector_pinned.is_some() {
            // Unpin: follow selection again
            self.inspector_pinned = None;
        } else {
            // Pin: lock to current selection
            self.inspector_pinned = Some(self.selected);
        }
        cx.notify();
    }

    // =========================================================================
    // Rename Symbol (Ctrl+Shift+R)
    // =========================================================================

    /// Show the rename symbol dialog
    /// If `name` is provided, pre-fill with that named range
    pub fn show_rename_symbol(&mut self, name: Option<&str>, cx: &mut Context<Self>) {
        // Get list of named ranges
        let named_ranges = self.workbook.list_named_ranges();
        if named_ranges.is_empty() {
            self.status_message = Some("No named ranges defined".to_string());
            cx.notify();
            return;
        }

        // If name provided, use it; otherwise try to detect from current cell
        let original = if let Some(n) = name {
            n.to_string()
        } else {
            // Try to find a named range in the current cell's formula
            let sheet = self.workbook.active_sheet();
            let (row, col) = self.selected;
            let cell = sheet.get_cell(row, col);
            let formula_text = self.get_formula_source(&cell.value);
            if let Some(formula) = formula_text {
                // Look for named range references in the formula
                self.find_named_range_in_formula(&formula)
            } else {
                None
            }.unwrap_or_else(|| {
                // No named range in current cell - use first available
                named_ranges.first().map(|nr| nr.name.clone()).unwrap_or_default()
            })
        };

        if original.is_empty() {
            self.status_message = Some("No named range to rename".to_string());
            cx.notify();
            return;
        }

        self.mode = Mode::RenameSymbol;
        self.rename_original_name = original.clone();
        self.rename_new_name = original;
        self.rename_validation_error = None;
        self.update_rename_affected_cells();
        cx.notify();
    }

    /// Extract formula source from a CellValue if it's a formula
    fn get_formula_source(&self, value: &visigrid_engine::cell::CellValue) -> Option<String> {
        match value {
            visigrid_engine::cell::CellValue::Formula { source, .. } => Some(source.clone()),
            _ => None,
        }
    }

    /// Hide the rename symbol dialog
    pub fn hide_rename_symbol(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.rename_original_name.clear();
        self.rename_new_name.clear();
        self.rename_affected_cells.clear();
        self.rename_validation_error = None;
        cx.notify();
    }

    /// Show the edit description modal for a named range
    pub fn show_edit_description(&mut self, name: &str, cx: &mut Context<Self>) {
        // Get the current description
        let current_description = self.workbook.get_named_range(name)
            .and_then(|nr| nr.description.clone());

        self.edit_description_name = name.to_string();
        self.edit_description_value = current_description.clone().unwrap_or_default();
        self.edit_description_original = current_description;
        self.mode = Mode::EditDescription;
        cx.notify();
    }

    /// Hide the edit description modal without saving
    pub fn hide_edit_description(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.edit_description_name.clear();
        self.edit_description_value.clear();
        self.edit_description_original = None;
        cx.notify();
    }

    /// Insert a character into the description
    pub fn edit_description_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.edit_description_value.push(c);
        cx.notify();
    }

    /// Delete the last character from the description
    pub fn edit_description_backspace(&mut self, cx: &mut Context<Self>) {
        self.edit_description_value.pop();
        cx.notify();
    }

    /// Apply the edited description and record undo
    pub fn apply_edit_description(&mut self, cx: &mut Context<Self>) {
        let name = self.edit_description_name.clone();
        let old_description = self.edit_description_original.clone();
        let new_description = if self.edit_description_value.is_empty() {
            None
        } else {
            Some(self.edit_description_value.clone())
        };

        // Only record if there's a change
        if old_description != new_description {
            // Apply the change
            let _ = self.workbook.named_ranges_mut().set_description(&name, new_description.clone());

            // Record for undo
            self.history.record_named_range_action(UndoAction::NamedRangeDescriptionChanged {
                name: name.clone(),
                old_description,
                new_description,
            });

            self.is_modified = true;
            self.status_message = Some(format!("Updated description for '{}'", name));
        }

        // Close the modal
        self.hide_edit_description(cx);
    }

    // ========== Tour Methods ==========

    /// Show the named ranges tour
    pub fn show_tour(&mut self, cx: &mut Context<Self>) {
        self.tour_step = 0;
        self.mode = Mode::Tour;
        cx.notify();
    }

    /// Hide the tour
    pub fn hide_tour(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Go to next tour step
    pub fn tour_next(&mut self, cx: &mut Context<Self>) {
        if self.tour_step < 3 {
            self.tour_step += 1;
            cx.notify();
        }
    }

    /// Go to previous tour step
    pub fn tour_back(&mut self, cx: &mut Context<Self>) {
        if self.tour_step > 0 {
            self.tour_step -= 1;
            cx.notify();
        }
    }

    /// Complete the tour
    pub fn tour_done(&mut self, cx: &mut Context<Self>) {
        self.tour_completed = true;
        self.mode = Mode::Navigation;
        self.status_message = Some("You just refactored a spreadsheet like code.".to_string());
        cx.notify();
    }

    /// Check if the name tooltip should be shown
    pub fn should_show_name_tooltip(&self) -> bool {
        // Show if: not dismissed, no named ranges exist, has a range selection
        !self.name_tooltip_dismissed
            && self.workbook.list_named_ranges().is_empty()
            && self.selection_end.is_some()
    }

    /// Dismiss the name tooltip permanently
    pub fn dismiss_name_tooltip(&mut self, cx: &mut Context<Self>) {
        self.name_tooltip_dismissed = true;
        // Save to settings
        let mut settings = Settings::load();
        settings.name_tooltip_dismissed = true;
        settings.save();
        cx.notify();
    }

    /// Insert a character into the new name
    pub fn rename_symbol_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.rename_new_name.push(c);
        self.validate_rename_name();
        cx.notify();
    }

    /// Delete the last character from the new name
    pub fn rename_symbol_backspace(&mut self, cx: &mut Context<Self>) {
        self.rename_new_name.pop();
        self.validate_rename_name();
        cx.notify();
    }

    /// Validate the current new name
    fn validate_rename_name(&mut self) {
        if self.rename_new_name.is_empty() {
            self.rename_validation_error = Some("Name cannot be empty".to_string());
            return;
        }

        // Check if it's the same as original (case-insensitive comparison for validity)
        if self.rename_new_name.to_lowercase() == self.rename_original_name.to_lowercase() {
            self.rename_validation_error = None;
            return;
        }

        // Check if name is valid
        if let Err(e) = is_valid_name(&self.rename_new_name) {
            self.rename_validation_error = Some(e);
            return;
        }

        // Check if name already exists
        if self.workbook.get_named_range(&self.rename_new_name).is_some() {
            self.rename_validation_error = Some(format!("'{}' already exists", self.rename_new_name));
            return;
        }

        self.rename_validation_error = None;
    }

    /// Update the list of affected cells (formulas using the named range)
    fn update_rename_affected_cells(&mut self) {
        self.rename_affected_cells.clear();

        let name_upper = self.rename_original_name.to_uppercase();
        let sheet = self.workbook.active_sheet();

        // Scan all cells for formulas that reference this named range
        for (&(row, col), cell) in sheet.cells_iter() {
            if let Some(formula) = self.get_formula_source(&cell.value) {
                if self.formula_references_name(&formula, &name_upper) {
                    self.rename_affected_cells.push((row, col));
                }
            }
        }
    }

    /// Check if a formula references a named range (case-insensitive)
    fn formula_references_name(&self, formula: &str, name_upper: &str) -> bool {
        // Simple check: look for the name as a word boundary
        // A proper implementation would parse the formula and check the AST
        let formula_upper = formula.to_uppercase();

        // Check for word boundaries using simple logic
        let name_len = name_upper.len();
        for (i, _) in formula_upper.match_indices(name_upper) {
            // Check if it's a word boundary (not part of a larger identifier)
            let before_ok = i == 0 || {
                let c = formula_upper.chars().nth(i - 1).unwrap_or(' ');
                !c.is_alphanumeric() && c != '_'
            };
            let after_ok = i + name_len >= formula_upper.len() || {
                let c = formula_upper.chars().nth(i + name_len).unwrap_or(' ');
                !c.is_alphanumeric() && c != '_'
            };
            if before_ok && after_ok {
                return true;
            }
        }
        false
    }

    /// Find a named range identifier in a formula string
    fn find_named_range_in_formula(&self, formula: &str) -> Option<String> {
        let named_ranges = self.workbook.list_named_ranges();
        let formula_upper = formula.to_uppercase();

        for nr in &named_ranges {
            let name_upper = nr.name.to_uppercase();
            if self.formula_references_name(&formula_upper, &name_upper) {
                return Some(nr.name.clone());
            }
        }
        None
    }

    /// Apply the rename operation
    pub fn confirm_rename_symbol(&mut self, cx: &mut Context<Self>) {
        // Validate first
        self.validate_rename_name();
        if self.rename_validation_error.is_some() {
            return;
        }

        let old_name = self.rename_original_name.clone();
        let new_name = self.rename_new_name.clone();

        // If names are the same (case-insensitive), just close
        if old_name.to_lowercase() == new_name.to_lowercase() {
            self.hide_rename_symbol(cx);
            return;
        }

        // Collect all formula changes for undo
        let mut changes: Vec<CellChange> = Vec::new();
        let sheet_index = self.workbook.active_sheet_index();
        let old_name_upper = old_name.to_uppercase();

        // Update formulas in all affected cells
        {
            let sheet = self.workbook.active_sheet();
            for &(row, col) in &self.rename_affected_cells {
                let cell = sheet.get_cell(row, col);
                if let Some(formula) = self.get_formula_source(&cell.value) {
                    let new_formula = self.replace_name_in_formula(&formula, &old_name_upper, &new_name);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value: formula,
                        new_value: new_formula,
                    });
                }
            }
        }

        // Apply the formula changes
        {
            let sheet = self.workbook.active_sheet_mut();
            for change in &changes {
                sheet.set_value(change.row, change.col, &change.new_value);
            }
        }

        // Rename the named range itself
        if let Err(e) = self.workbook.rename_named_range(&old_name, &new_name) {
            self.status_message = Some(format!("Failed to rename: {}", e));
            // TODO: Roll back formula changes
            cx.notify();
            return;
        }

        // Record undo action
        if !changes.is_empty() {
            self.history.record_batch(sheet_index, changes);
        }

        self.is_modified = true;
        self.bump_cells_rev();
        self.status_message = Some(format!(
            "Renamed '{}' to '{}' ({} formula{} updated)",
            old_name,
            new_name,
            self.rename_affected_cells.len(),
            if self.rename_affected_cells.len() == 1 { "" } else { "s" }
        ));

        self.hide_rename_symbol(cx);
    }

    /// Replace a named range in a formula with a new name
    /// Handles case-insensitive matching while preserving surrounding text
    fn replace_name_in_formula(&self, formula: &str, old_name_upper: &str, new_name: &str) -> String {
        let mut result = String::with_capacity(formula.len());
        let formula_chars: Vec<char> = formula.chars().collect();
        let old_name_len = old_name_upper.len();
        let mut i = 0;

        while i < formula_chars.len() {
            // Try to match old name at this position
            let remaining: String = formula_chars[i..].iter().collect();
            let remaining_upper = remaining.to_uppercase();

            if remaining_upper.starts_with(old_name_upper) {
                // Check word boundaries
                let before_ok = i == 0 || {
                    let c = formula_chars[i - 1];
                    !c.is_alphanumeric() && c != '_'
                };
                let after_ok = i + old_name_len >= formula_chars.len() || {
                    let c = formula_chars[i + old_name_len];
                    !c.is_alphanumeric() && c != '_'
                };

                if before_ok && after_ok {
                    // Found a match - replace it
                    result.push_str(new_name);
                    i += old_name_len;
                    continue;
                }
            }

            result.push(formula_chars[i]);
            i += 1;
        }

        result
    }

    // ========================================================================
    // Create Named Range (Ctrl+Shift+N)
    // ========================================================================

    /// Show the create named range dialog
    pub fn show_create_named_range(&mut self, cx: &mut Context<Self>) {
        // Build target string from current selection
        let target = self.selection_to_reference_string();

        self.create_name_name = String::new();
        self.create_name_description = String::new();
        self.create_name_target = target;
        self.create_name_validation_error = None;
        self.create_name_focus = CreateNameFocus::Name;
        self.mode = Mode::CreateNamedRange;
        cx.notify();
    }

    /// Hide the create named range dialog
    pub fn hide_create_named_range(&mut self, cx: &mut Context<Self>) {
        self.create_name_name.clear();
        self.create_name_description.clear();
        self.create_name_target.clear();
        self.create_name_validation_error = None;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Insert a character into the currently focused create name field
    pub fn create_name_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        match self.create_name_focus {
            CreateNameFocus::Name => self.create_name_name.push(c),
            CreateNameFocus::Description => self.create_name_description.push(c),
        }
        self.validate_create_name();
        cx.notify();
    }

    /// Backspace in the currently focused create name field
    pub fn create_name_backspace(&mut self, cx: &mut Context<Self>) {
        match self.create_name_focus {
            CreateNameFocus::Name => { self.create_name_name.pop(); }
            CreateNameFocus::Description => { self.create_name_description.pop(); }
        }
        self.validate_create_name();
        cx.notify();
    }

    /// Tab to next field in create named range dialog
    pub fn create_name_tab(&mut self, cx: &mut Context<Self>) {
        self.create_name_focus = match self.create_name_focus {
            CreateNameFocus::Name => CreateNameFocus::Description,
            CreateNameFocus::Description => CreateNameFocus::Name,
        };
        cx.notify();
    }

    /// Validate the name field
    fn validate_create_name(&mut self) {
        use visigrid_engine::named_range::is_valid_name;

        if self.create_name_name.is_empty() {
            self.create_name_validation_error = Some("Name is required".into());
            return;
        }

        if let Err(e) = is_valid_name(&self.create_name_name) {
            self.create_name_validation_error = Some(e);
            return;
        }

        // Check if name already exists
        if self.workbook.get_named_range(&self.create_name_name).is_some() {
            self.create_name_validation_error = Some(format!(
                "'{}' already exists",
                self.create_name_name
            ));
            return;
        }

        self.create_name_validation_error = None;
    }

    /// Confirm creation of the named range
    pub fn confirm_create_named_range(&mut self, cx: &mut Context<Self>) {
        // Validate first
        self.validate_create_name();
        if self.create_name_validation_error.is_some() {
            return;
        }

        let name = self.create_name_name.clone();
        let description = if self.create_name_description.is_empty() {
            None
        } else {
            Some(self.create_name_description.clone())
        };

        // Parse the selection and create the named range
        let (anchor_row, anchor_col) = self.selected;
        let (end_row, end_col) = self.selection_end.unwrap_or(self.selected);
        let (start_row, start_col, end_row, end_col) = (
            anchor_row.min(end_row),
            anchor_col.min(end_col),
            anchor_row.max(end_row),
            anchor_col.max(end_col),
        );
        let sheet = self.workbook.active_sheet_index();

        let result = if start_row == end_row && start_col == end_col {
            // Single cell
            self.workbook.define_name_for_cell(&name, sheet, start_row, start_col)
        } else {
            // Range
            self.workbook.define_name_for_range(
                &name, sheet, start_row, start_col, end_row, end_col
            )
        };

        match result {
            Ok(()) => {
                // Add description if provided
                if let Some(desc) = description {
                    if let Some(nr) = self.workbook.named_ranges_mut().get(&name).cloned() {
                        let mut updated = nr;
                        updated.description = Some(desc);
                        let _ = self.workbook.named_ranges_mut().set(updated);
                    }
                }

                self.is_modified = true;
                self.status_message = Some(format!(
                    "Created named range '{}' → {}",
                    name,
                    self.create_name_target
                ));
                self.hide_create_named_range(cx);
            }
            Err(e) => {
                self.create_name_validation_error = Some(e);
                cx.notify();
            }
        }
    }

    /// Convert current selection to a reference string (e.g., "A1" or "A1:B10")
    fn selection_to_reference_string(&self) -> String {
        let (anchor_row, anchor_col) = self.selected;
        let (end_row, end_col) = self.selection_end.unwrap_or(self.selected);
        let (start_row, start_col, end_row, end_col) = (
            anchor_row.min(end_row),
            anchor_col.min(end_col),
            anchor_row.max(end_row),
            anchor_col.max(end_col),
        );

        let start_ref = format!("{}{}", col_to_letter(start_col), start_row + 1);

        if start_row == end_row && start_col == end_col {
            start_ref
        } else {
            format!("{}:{}{}", start_ref, col_to_letter(end_col), end_row + 1)
        }
    }

    // ========================================================================
    // Named Ranges Panel Actions
    // ========================================================================

    /// Delete a named range by name (with undo support and reference warning)
    pub fn delete_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        // Get the named range first (need to clone for undo)
        let named_range = self.workbook.get_named_range(name).cloned();

        if let Some(nr) = named_range {
            // Count references to this named range in formulas
            let ref_count = self.count_named_range_references(&nr.name);

            // Record undo action BEFORE deleting
            self.history.record_named_range_action(UndoAction::NamedRangeDeleted {
                named_range: nr.clone(),
            });

            // Now delete
            self.workbook.delete_named_range(name);
            self.is_modified = true;
            self.bump_cells_rev();

            // Status message with reference warning
            if ref_count > 0 {
                self.status_message = Some(format!(
                    "Deleted '{}' (used in {} formula{}—will show #NAME? errors)",
                    name,
                    ref_count,
                    if ref_count == 1 { "" } else { "s" }
                ));
            } else {
                self.status_message = Some(format!("Deleted named range '{}'", name));
            }
            cx.notify();
        } else {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
        }
    }

    /// Count how many formula cells reference a named range
    fn count_named_range_references(&self, name: &str) -> usize {
        let name_upper = name.to_uppercase();
        let mut count = 0;

        for ((_, _), cell) in self.sheet().cells_iter() {
            let raw = cell.value.raw_display();
            if raw.starts_with('=') {
                // Simple check: does the formula contain this name as a word?
                // More sophisticated: parse the formula and check identifiers
                // For now, do case-insensitive word boundary check
                let formula_upper = raw.to_uppercase();
                // Check if name appears as a standalone identifier
                // This is a simple heuristic - a proper check would parse the formula
                for word in formula_upper.split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.') {
                    if word == name_upper {
                        count += 1;
                        break; // Count each cell only once
                    }
                }
            }
        }
        count
    }

    /// Get usage count for a named range (with caching)
    pub fn get_named_range_usage_count(&mut self, name: &str) -> usize {
        // Check if cache is stale
        if self.named_range_usage_cache.cached_rev != self.cells_rev {
            self.rebuild_named_range_usage_cache();
        }

        // Return cached count (or 0 if not found)
        self.named_range_usage_cache.counts
            .get(&name.to_lowercase())
            .copied()
            .unwrap_or(0)
    }

    /// Rebuild the usage count cache for all named ranges
    fn rebuild_named_range_usage_cache(&mut self) {
        self.named_range_usage_cache.counts.clear();

        // Get all named range names (lowercase for lookup)
        let names: Vec<String> = self.workbook.list_named_ranges()
            .iter()
            .map(|nr| nr.name.to_lowercase())
            .collect();

        // Also store uppercase versions for matching
        let names_upper: Vec<String> = names.iter()
            .map(|n| n.to_uppercase())
            .collect();

        // Initialize all counts to 0
        for name in &names {
            self.named_range_usage_cache.counts.insert(name.clone(), 0);
        }

        // Collect all formulas first (to avoid borrow issues)
        let formulas: Vec<String> = self.sheet().cells_iter()
            .filter_map(|((_, _), cell)| {
                let raw = cell.value.raw_display();
                if raw.starts_with('=') {
                    Some(raw.to_uppercase())
                } else {
                    None
                }
            })
            .collect();

        // Now process formulas and update counts
        for formula_upper in formulas {
            // Check each named range
            for (i, name_upper) in names_upper.iter().enumerate() {
                // Check if name appears as a standalone identifier
                for word in formula_upper.split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.') {
                    if word == name_upper {
                        if let Some(count) = self.named_range_usage_cache.counts.get_mut(&names[i]) {
                            *count += 1;
                        }
                        break; // Count each cell only once per name
                    }
                }
            }
        }

        // Mark cache as fresh
        self.named_range_usage_cache.cached_rev = self.cells_rev;
    }

    /// Jump to a named range definition and select the whole range
    pub fn jump_to_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;

        let target_info = self.workbook.get_named_range(name).map(|nr| {
            match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => {
                    (*row, *col, *row, *col, nr.reference_string())
                }
                NamedRangeTarget::Range { start_row, start_col, end_row, end_col, .. } => {
                    (*start_row, *start_col, *end_row, *end_col, nr.reference_string())
                }
            }
        });

        if let Some((start_row, start_col, end_row, end_col, ref_str)) = target_info {
            // Select the whole range
            self.selected = (start_row, start_col);
            if start_row == end_row && start_col == end_col {
                self.selection_end = None;
            } else {
                self.selection_end = Some((end_row, end_col));
            }

            // Center the view on the selection
            self.ensure_cell_visible(start_row, start_col);

            self.status_message = Some(format!("'{}' = {}", name, ref_str));
            cx.notify();
        } else {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
        }
    }

    /// Filter named ranges by query (for Names panel search)
    pub fn set_names_filter(&mut self, query: String, cx: &mut Context<Self>) {
        self.names_filter_query = query;
        cx.notify();
    }

    /// Get filtered named ranges for the Names panel
    pub fn filtered_named_ranges(&self) -> Vec<&visigrid_engine::named_range::NamedRange> {
        let query = self.names_filter_query.to_lowercase();
        let mut ranges: Vec<_> = self.workbook.list_named_ranges()
            .into_iter()
            .filter(|nr| {
                if query.is_empty() {
                    return true;
                }
                // Match against name or description
                nr.name.to_lowercase().contains(&query)
                    || nr.description.as_ref()
                        .map(|d| d.to_lowercase().contains(&query))
                        .unwrap_or(false)
            })
            .collect();

        // Sort alphabetically by name
        ranges.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
        ranges
    }
}

/// Convert column index to letter(s) (0 = A, 25 = Z, 26 = AA, etc.)
fn col_to_letter(col: usize) -> String {
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
        // Update window size if changed (handles resize)
        let current_size = window.viewport_size();
        if self.window_size != current_size {
            self.window_size = current_size;
        }
        views::render_spreadsheet(self, cx)
    }
}
