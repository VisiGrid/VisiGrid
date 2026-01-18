use iced::widget::{button, column, container, mouse_area, row, text, text_input, Column, Row};
use iced::keyboard::{self, Key, Modifiers};
use iced::{Color, Element, Font, Length, Subscription, Event, Border, Background, Task};
use iced::font::{Weight, Style as FontStyle};
use std::path::PathBuf;

use crate::config::keybindings::KeybindingManager;
use crate::config::omarchy;
use crate::config::settings::Settings;
use crate::io::{csv, native};

fn formula_input_id() -> text_input::Id {
    text_input::Id::new("formula_input")
}

// Theme colors - aligned with VisiGrid site palette
#[derive(Clone, Copy)]
pub struct ThemeColors {
    pub bg_dark: Color,
    pub bg_header: Color,
    pub bg_cell: Color,
    pub bg_input: Color,
    pub text: Color,
    pub text_dim: Color,
    pub gridline: Color,
    pub accent: Color,
    pub selected: Color,
    pub selected_border: Color,
    pub border: Color,
}

impl ThemeColors {
    pub const fn dark() -> Self {
        Self {
            bg_dark: Color::from_rgb(0.008, 0.024, 0.090),        // #020617 (grid-950)
            bg_header: Color::from_rgb(0.118, 0.161, 0.231),      // #1e293b (grid-800)
            bg_cell: Color::from_rgb(0.059, 0.090, 0.165),        // #0f172a (grid-900)
            bg_input: Color::from_rgb(0.059, 0.090, 0.165),       // #0f172a (grid-900)
            text: Color::from_rgb(0.945, 0.961, 0.976),           // #f1f5f9 (grid-100)
            text_dim: Color::from_rgb(0.392, 0.439, 0.529),       // #64748b (grid-500)
            gridline: Color::from_rgb(0.200, 0.255, 0.333),       // #334155 (grid-700)
            accent: Color::from_rgb(0.231, 0.510, 0.965),         // #3b82f6
            selected: Color::from_rgba(0.231, 0.510, 0.965, 0.2), // #3b82f6 @ 20%
            selected_border: Color::from_rgb(0.231, 0.510, 0.965),// #3b82f6
            border: Color::from_rgb(0.200, 0.255, 0.333),         // #334155 (grid-700)
        }
    }

    pub const fn light() -> Self {
        Self {
            bg_dark: Color::from_rgb(0.973, 0.980, 0.988),        // #f8fafc (grid-50)
            bg_header: Color::from_rgb(0.886, 0.910, 0.941),      // #e2e8f0 (grid-200)
            bg_cell: Color::from_rgb(1.0, 1.0, 1.0),              // #ffffff
            bg_input: Color::from_rgb(0.945, 0.961, 0.976),       // #f1f5f9 (grid-100)
            text: Color::from_rgb(0.059, 0.090, 0.165),           // #0f172a (grid-900)
            text_dim: Color::from_rgb(0.278, 0.333, 0.412),       // #475569 (grid-600)
            gridline: Color::from_rgb(0.886, 0.910, 0.941),       // #e2e8f0 (grid-200)
            accent: Color::from_rgb(0.231, 0.510, 0.965),         // #3b82f6
            selected: Color::from_rgba(0.231, 0.510, 0.965, 0.15),// #3b82f6 @ 15%
            selected_border: Color::from_rgb(0.231, 0.510, 0.965),// #3b82f6
            border: Color::from_rgb(0.796, 0.835, 0.882),         // #cbd5e1 (grid-300)
        }
    }
}

// Keep old colors module for compatibility during refactor (using new palette)
mod colors {
    use iced::Color;

    pub const BG_DARK: Color = Color::from_rgb(0.008, 0.024, 0.090);        // #020617
    pub const BG_HEADER: Color = Color::from_rgb(0.118, 0.161, 0.231);      // #1e293b
    pub const BG_CELL: Color = Color::from_rgb(0.059, 0.090, 0.165);        // #0f172a
    pub const BG_INPUT: Color = Color::from_rgb(0.059, 0.090, 0.165);       // #0f172a
    pub const TEXT: Color = Color::from_rgb(0.945, 0.961, 0.976);           // #f1f5f9
    pub const TEXT_DIM: Color = Color::from_rgb(0.392, 0.439, 0.529);       // #64748b
    pub const GRIDLINE: Color = Color::from_rgb(0.200, 0.255, 0.333);       // #334155
    pub const ACCENT: Color = Color::from_rgb(0.231, 0.510, 0.965);         // #3b82f6
    pub const SELECTED: Color = Color::from_rgba(0.231, 0.510, 0.965, 0.2); // #3b82f6 @ 20%
    pub const BORDER: Color = Color::from_rgb(0.200, 0.255, 0.333);         // #334155
}
use crate::engine::cell::{Alignment, NumberFormat};
use crate::core::selection::Selection;
use crate::engine::sheet::Sheet;
use crate::ui::formula_highlight::{highlight_formula, is_formula};

// Excel 2003 limits: 65,536 rows × 256 columns
const NUM_ROWS: usize = 65536;
const NUM_COLS: usize = 256;

// Viewport settings
const CELL_WIDTH: f32 = 80.0;
const CELL_HEIGHT: f32 = 24.0;
const ROW_HEADER_WIDTH: f32 = 50.0;
const VISIBLE_ROWS: usize = 40;   // Render this many rows
const VISIBLE_COLS: usize = 26;   // Render this many columns (covers ~2100px width)

// Application modes
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    Navigation,  // Grid focus: keystrokes move selection
    Edit,        // Cell editor focus: keystrokes edit text
    Command,     // Command palette open
    GoTo,        // Go to cell dialog
    QuickOpen,   // Quick file open (Ctrl+P)
    Find,        // Find in cells (Ctrl+F)
}

// Command palette search modes (determined by prefix)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaletteMode {
    Mixed,      // No prefix: commands + recent files
    Commands,   // > prefix: commands only
    Cells,      // @ prefix: search cells
    GoTo,       // : prefix: go to cell
    Functions,  // = prefix: search formula functions
}

// Split view direction
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SplitDirection {
    Horizontal,  // Side by side (left/right)
    Vertical,    // Stacked (top/bottom)
}

// Which pane is active in split view
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SplitPane {
    Primary,    // Left or top pane
    Secondary,  // Right or bottom pane
}

// Formula function definition for search
#[derive(Debug, Clone)]
pub struct FormulaFunction {
    pub name: &'static str,
    pub syntax: &'static str,
    pub description: &'static str,
    pub category: &'static str,
}

// Available formula functions
const FORMULA_FUNCTIONS: &[FormulaFunction] = &[
    // Math & Statistics
    FormulaFunction { name: "SUM", syntax: "=SUM(range)", description: "Adds all numbers in a range", category: "Math" },
    FormulaFunction { name: "AVERAGE", syntax: "=AVERAGE(range)", description: "Returns the arithmetic mean of values", category: "Math" },
    FormulaFunction { name: "AVG", syntax: "=AVG(range)", description: "Alias for AVERAGE", category: "Math" },
    FormulaFunction { name: "MIN", syntax: "=MIN(range)", description: "Returns the smallest value in a range", category: "Math" },
    FormulaFunction { name: "MAX", syntax: "=MAX(range)", description: "Returns the largest value in a range", category: "Math" },
    FormulaFunction { name: "COUNT", syntax: "=COUNT(range)", description: "Counts numbers in a range", category: "Math" },
    FormulaFunction { name: "ABS", syntax: "=ABS(value)", description: "Returns the absolute value of a number", category: "Math" },
    FormulaFunction { name: "ROUND", syntax: "=ROUND(value, decimals)", description: "Rounds a number to specified decimals", category: "Math" },
];

// Undo/Redo support
#[derive(Debug, Clone)]
pub struct CellChange {
    row: usize,
    col: usize,
    old_value: String,
    new_value: String,
}

#[derive(Debug, Clone)]
pub struct UndoAction {
    description: String,
    changes: Vec<CellChange>,
}

impl UndoAction {
    fn new(description: &str) -> Self {
        Self {
            description: description.to_string(),
            changes: Vec::new(),
        }
    }

    fn add_change(&mut self, row: usize, col: usize, old_value: String, new_value: String) {
        self.changes.push(CellChange { row, col, old_value, new_value });
    }

    fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }
}

// Cell error for Problems panel
#[derive(Debug, Clone)]
pub struct CellError {
    pub row: usize,
    pub col: usize,
    pub error_type: String,    // e.g., "#REF!", "#DIV/0!"
    pub description: String,   // Human-readable description
}

// Command definitions
#[derive(Debug, Clone)]
pub struct Command {
    pub id: &'static str,
    pub label: &'static str,
    pub shortcut: Option<&'static str>,
}

const COMMANDS: &[Command] = &[
    Command { id: "new", label: "New Workbook", shortcut: Some("Ctrl+N") },
    Command { id: "save", label: "Save", shortcut: Some("Ctrl+S") },
    Command { id: "open", label: "Open File", shortcut: Some("Ctrl+O") },
    Command { id: "copy", label: "Copy Cell", shortcut: Some("Ctrl+C") },
    Command { id: "cut", label: "Cut Cell", shortcut: Some("Ctrl+X") },
    Command { id: "paste", label: "Paste", shortcut: Some("Ctrl+V") },
    Command { id: "delete", label: "Clear Cell", shortcut: Some("Delete") },
    Command { id: "sum", label: "Insert SUM Formula", shortcut: Some("Alt+=") },
    Command { id: "goto", label: "Go to Cell...", shortcut: Some("Ctrl+G") },
    Command { id: "theme_toggle", label: "Toggle Dark/Light Theme", shortcut: None },
    Command { id: "theme_reload", label: "Reload System Theme", shortcut: None },
    Command { id: "settings_open", label: "Open Settings (JSON)", shortcut: None },
    Command { id: "keybindings_open", label: "Open Keyboard Shortcuts (JSON)", shortcut: None },
    Command { id: "problems_toggle", label: "Show/Hide Problems Panel", shortcut: Some("Ctrl+Shift+M") },
    Command { id: "split_toggle", label: "Toggle Split View", shortcut: Some("Ctrl+\\") },
    Command { id: "split_horizontal", label: "Split Editor Horizontally", shortcut: None },
    Command { id: "split_vertical", label: "Split Editor Vertically", shortcut: None },
    Command { id: "split_switch", label: "Switch Split Pane", shortcut: Some("Ctrl+W") },
    Command { id: "zen_mode", label: "Toggle Zen Mode", shortcut: Some("F11") },
    Command { id: "inspector", label: "Toggle Cell Inspector", shortcut: Some("Ctrl+Shift+I") },
    Command { id: "quick_open", label: "Quick Open Recent File", shortcut: Some("Ctrl+P") },
    Command { id: "find", label: "Find in Cells", shortcut: Some("Ctrl+F") },
    Command { id: "fill_down", label: "Fill Down", shortcut: Some("Ctrl+D") },
    Command { id: "fill_right", label: "Fill Right", shortcut: Some("Ctrl+R") },
    Command { id: "select_all", label: "Select All", shortcut: Some("Ctrl+A") },
    Command { id: "help", label: "Show Help", shortcut: Some("F1") },
];

// Menu item definition
#[derive(Debug, Clone)]
pub struct MenuItem {
    pub id: &'static str,
    pub label: &'static str,
    pub shortcut: Option<&'static str>,
    pub separator_after: bool,
}

impl MenuItem {
    const fn new(id: &'static str, label: &'static str, shortcut: Option<&'static str>) -> Self {
        Self { id, label, shortcut, separator_after: false }
    }
    const fn with_separator(id: &'static str, label: &'static str, shortcut: Option<&'static str>) -> Self {
        Self { id, label, shortcut, separator_after: true }
    }
}

// Menu definitions
const FILE_MENU: &[MenuItem] = &[
    MenuItem::new("new", "New", Some("Ctrl+N")),
    MenuItem::new("open", "Open...", Some("Ctrl+O")),
    MenuItem::with_separator("quick_open", "Quick Open...", Some("Ctrl+P")),
    MenuItem::new("save", "Save", Some("Ctrl+S")),
    MenuItem::with_separator("save_as", "Save As...", Some("Ctrl+Shift+S")),
    MenuItem::new("export_csv", "Export as CSV...", None),
    MenuItem::with_separator("quit", "Quit", Some("Ctrl+Q")),
];

const EDIT_MENU: &[MenuItem] = &[
    MenuItem::new("undo", "Undo", Some("Ctrl+Z")),
    MenuItem::with_separator("redo", "Redo", Some("Ctrl+Y")),
    MenuItem::new("cut", "Cut", Some("Ctrl+X")),
    MenuItem::new("copy", "Copy", Some("Ctrl+C")),
    MenuItem::with_separator("paste", "Paste", Some("Ctrl+V")),
    MenuItem::new("delete", "Clear", Some("Delete")),
    MenuItem::with_separator("select_all", "Select All", Some("Ctrl+A")),
    MenuItem::new("find", "Find...", Some("Ctrl+F")),
    MenuItem::new("replace", "Replace...", Some("Ctrl+H")),
];

const VIEW_MENU: &[MenuItem] = &[
    MenuItem::new("theme_toggle", "Toggle Dark/Light", None),
    MenuItem::with_separator("problems_toggle", "Problems Panel", Some("Ctrl+Shift+M")),
    MenuItem::new("zoom_in", "Zoom In", Some("Ctrl++")),
    MenuItem::new("zoom_out", "Zoom Out", Some("Ctrl+-")),
    MenuItem::with_separator("zoom_reset", "Reset Zoom", Some("Ctrl+0")),
    MenuItem::new("fullscreen", "Fullscreen", Some("F11")),
];

const INSERT_MENU: &[MenuItem] = &[
    MenuItem::new("insert_row", "Row Above", None),
    MenuItem::with_separator("insert_col", "Column Left", None),
    MenuItem::new("sum", "SUM Formula", Some("Alt+=")),
    MenuItem::new("insert_function", "Function...", None),
];

const FORMAT_MENU: &[MenuItem] = &[
    MenuItem::new("bold", "Bold", Some("Ctrl+B")),
    MenuItem::new("italic", "Italic", Some("Ctrl+I")),
    MenuItem::with_separator("underline", "Underline", Some("Ctrl+U")),
    MenuItem::new("align_left", "Align Left", None),
    MenuItem::new("align_center", "Align Center", None),
    MenuItem::with_separator("align_right", "Align Right", None),
    MenuItem::new("format_number", "Number Format...", None),
];

const DATA_MENU: &[MenuItem] = &[
    MenuItem::new("sort_asc", "Sort A to Z", None),
    MenuItem::with_separator("sort_desc", "Sort Z to A", None),
    MenuItem::new("filter", "Filter", None),
    MenuItem::with_separator("fill_down", "Fill Down", Some("Ctrl+D")),
    MenuItem::new("fill_right", "Fill Right", Some("Ctrl+R")),
];

const HELP_MENU: &[MenuItem] = &[
    MenuItem::new("shortcuts", "Keyboard Shortcuts", Some("Ctrl+/")),
    MenuItem::with_separator("command_palette", "Command Palette", Some("Ctrl+Shift+P")),
    MenuItem::new("about", "About", None),
];

pub struct App {
    sheet: Sheet,
    selection: Selection,
    mode: Mode,
    input_value: String,
    edit_original: String,  // Original value before edit (for cancel)
    clipboard: Option<String>,
    dark_mode: bool,
    // Command palette state
    palette_query: String,
    palette_selected: usize,
    palette_filtered: Vec<usize>,
    palette_mode: PaletteMode,
    palette_cells: Vec<(usize, usize, String)>,  // Cell search results in palette
    palette_files: Vec<usize>,                    // Recent file indices matching query
    palette_functions: Vec<usize>,                // Formula function indices matching query
    // Menu state
    open_menu: Option<&'static str>,
    // Keybindings
    keybindings: KeybindingManager,
    // File state
    current_file: Option<PathBuf>,
    status_message: Option<String>,
    // Go to cell state
    goto_input: String,
    // Undo/Redo stacks
    undo_stack: Vec<UndoAction>,
    redo_stack: Vec<UndoAction>,
    // Viewport (first visible row/col) - primary pane
    scroll_row: usize,
    scroll_col: usize,
    // Split view
    split_enabled: bool,
    split_direction: SplitDirection,
    split_active_pane: SplitPane,
    split_scroll_row: usize,  // Secondary pane scroll
    split_scroll_col: usize,
    // Zen mode (distraction-free)
    zen_mode: bool,
    // Cell inspector panel
    show_inspector: bool,
    // Formula autocomplete
    autocomplete_visible: bool,
    autocomplete_suggestions: Vec<usize>,  // Indices into FORMULA_FUNCTIONS
    autocomplete_selected: usize,
    autocomplete_start_pos: usize,  // Position in input where function name starts
    // Signature help (parameter hints)
    signature_help_visible: bool,
    signature_help_function: Option<usize>,  // Index into FORMULA_FUNCTIONS
    signature_help_param: usize,             // Current parameter index (0-based)
    // Column/row sizing
    column_widths: Vec<f32>,
    row_heights: Vec<f32>,
    // Resize state
    resizing_col: Option<usize>,
    resizing_row: Option<usize>,
    resize_start_x: f32,
    resize_start_y: f32,
    resize_start_size: f32,
    // Problems panel
    show_problems: bool,
    problems_cache: Vec<CellError>,
    // Formula view mode (Ctrl+`)
    show_formulas: bool,
    // Quick Open state
    recent_files: Vec<PathBuf>,
    quick_open_query: String,
    quick_open_selected: usize,
    quick_open_filtered: Vec<usize>,
    // Find state
    find_query: String,
    find_results: Vec<(usize, usize, String)>,  // row, col, cell display
    find_selected: usize,
    // Omarchy theme integration
    omarchy_theme: Option<ThemeColors>,
    omarchy_theme_mtime: Option<std::time::SystemTime>,
    // Application settings
    settings: Settings,
}

#[derive(Debug, Clone)]
pub enum Message {
    CellClicked(usize, usize, Modifiers),
    InputChanged(String),
    InputSubmitted,
    KeyPressed(Key, Modifiers),
    PaletteQueryChanged(String),
    PaletteSubmit,
    ExecuteCommand(&'static str),
    MenuClicked(&'static str),
    MenuItemClicked(&'static str),
    // File operations
    FileOpened(Option<PathBuf>),
    FileSaved(Option<PathBuf>),
    FileExported(Option<PathBuf>),
    FileLoaded(Result<Sheet, String>, PathBuf),
    FileSaveResult(Result<(), String>),
    // Go to cell
    GoToInputChanged(String),
    GoToSubmit,
    // Formatting
    FormatBold,
    FormatItalic,
    FormatUnderline,
    FormatAlignLeft,
    FormatAlignCenter,
    FormatAlignRight,
    FormatCurrency,
    FormatPercent,
    FormatDecimalIncrease,
    FormatDecimalDecrease,
    // Column/row resizing
    ColResizeStart(usize, f32),      // col index, x position
    ColResizeMove(f32),               // x position
    ColResizeEnd,
    ColAutoSize(usize),               // double-click to auto-size
    RowResizeStart(usize, f32),      // row index, y position
    RowResizeMove(f32),               // y position
    RowResizeEnd,
    RowAutoSize(usize),               // double-click to auto-size
    // Problems panel
    ToggleProblems,
    GoToProblem(usize, usize),        // row, col of error cell
    // Quick Open
    QuickOpenQueryChanged(String),
    QuickOpenSubmit,
    QuickOpenSelect(usize),           // index in filtered list
    // Find in cells
    FindQueryChanged(String),
    FindSubmit,
    FindNext,
    FindPrev,
    FindSelect(usize, usize),         // row, col to jump to
    // Split view
    SplitToggle,                      // Toggle split on/off
    SplitHorizontal,                  // Split side by side
    SplitVertical,                    // Split top/bottom
    SplitSwitchPane,                  // Switch active pane
    SplitPaneClicked(SplitPane),      // Clicked on a specific pane
    // Zen mode
    ZenModeToggle,                    // Toggle distraction-free mode
    // Cell inspector
    InspectorToggle,                  // Toggle cell inspector panel
    // Autocomplete
    AutocompleteSelect(usize),        // Select autocomplete item by index
    AutocompleteApply,                // Apply selected autocomplete
    AutocompleteDismiss,              // Dismiss autocomplete popup
    // Omarchy integration
    OmarchyThemeChanged,              // System theme file changed
}

impl Default for App {
    fn default() -> Self {
        let filtered: Vec<usize> = (0..COMMANDS.len()).collect();
        let keybindings = KeybindingManager::new();
        let settings = Settings::load();

        // Log config paths on startup
        eprintln!("Keybindings config: {:?}", keybindings.config_path());
        eprintln!("Settings config: {}", Settings::config_path_display());

        // Load Omarchy theme if available
        let (omarchy_theme, omarchy_theme_mtime) = if omarchy::is_omarchy() {
            let theme = omarchy::load_theme();
            let mtime = omarchy::theme_mtime();
            if let Some(name) = omarchy::current_theme_name() {
                eprintln!("Omarchy theme detected: {}", name);
                eprintln!("  bg_dark: {:?}", theme.bg_dark);
            }
            (Some(theme), mtime)
        } else {
            (None, None)
        };

        // Use settings for initial column/row sizes
        let default_col_width = settings.default_column_width;
        let default_row_height = settings.row_height;

        Self {
            sheet: Sheet::new(NUM_ROWS, NUM_COLS),
            selection: Selection::new(0, 0),
            mode: Mode::Navigation,
            input_value: String::new(),
            edit_original: String::new(),
            clipboard: None,
            dark_mode: true,
            palette_query: String::new(),
            palette_selected: 0,
            palette_filtered: filtered,
            palette_mode: PaletteMode::Mixed,
            palette_cells: Vec::new(),
            palette_files: Vec::new(),
            palette_functions: Vec::new(),
            open_menu: None,
            keybindings,
            current_file: None,
            status_message: None,
            goto_input: String::new(),
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            scroll_row: 0,
            scroll_col: 0,
            split_enabled: false,
            split_direction: SplitDirection::Horizontal,
            split_active_pane: SplitPane::Primary,
            split_scroll_row: 0,
            split_scroll_col: 0,
            zen_mode: false,
            show_inspector: false,
            autocomplete_visible: false,
            autocomplete_suggestions: Vec::new(),
            autocomplete_selected: 0,
            autocomplete_start_pos: 0,
            signature_help_visible: false,
            signature_help_function: None,
            signature_help_param: 0,
            column_widths: vec![default_col_width; NUM_COLS],
            row_heights: vec![default_row_height; NUM_ROWS],
            resizing_col: None,
            resizing_row: None,
            resize_start_x: 0.0,
            resize_start_y: 0.0,
            resize_start_size: 0.0,
            show_problems: false,
            problems_cache: Vec::new(),
            show_formulas: false,
            recent_files: Self::load_recent_files(),
            quick_open_query: String::new(),
            quick_open_selected: 0,
            quick_open_filtered: Vec::new(),
            find_query: String::new(),
            find_results: Vec::new(),
            find_selected: 0,
            omarchy_theme,
            omarchy_theme_mtime,
            settings,
        }
    }
}

impl App {
    /// Get the current theme colors
    /// Priority: Omarchy theme > dark/light mode
    fn theme(&self) -> ThemeColors {
        if let Some(theme) = self.omarchy_theme {
            theme
        } else if self.dark_mode {
            ThemeColors::dark()
        } else {
            ThemeColors::light()
        }
    }

    // === Recent Files ===

    fn get_recent_files_path() -> PathBuf {
        dirs::config_dir()
            .unwrap_or_else(|| PathBuf::from("."))
            .join("visigrid")
            .join("recent_files.json")
    }

    fn load_recent_files() -> Vec<PathBuf> {
        let path = Self::get_recent_files_path();
        if path.exists() {
            if let Ok(contents) = std::fs::read_to_string(&path) {
                if let Ok(files) = serde_json::from_str::<Vec<String>>(&contents) {
                    return files.into_iter().map(PathBuf::from).collect();
                }
            }
        }
        Vec::new()
    }

    fn save_recent_files(&self) {
        let path = Self::get_recent_files_path();
        if let Some(parent) = path.parent() {
            let _ = std::fs::create_dir_all(parent);
        }
        let files: Vec<String> = self.recent_files.iter()
            .filter_map(|p| p.to_str().map(String::from))
            .collect();
        if let Ok(json) = serde_json::to_string_pretty(&files) {
            let _ = std::fs::write(&path, json);
        }
    }

    fn add_recent_file(&mut self, path: PathBuf) {
        // Remove if already exists (will re-add at front)
        self.recent_files.retain(|p| p != &path);
        // Add to front
        self.recent_files.insert(0, path);
        // Keep only last 20
        self.recent_files.truncate(20);
        // Save
        self.save_recent_files();
    }

    // === Quick Open ===

    fn open_quick_open(&mut self) {
        self.mode = Mode::QuickOpen;
        self.quick_open_query.clear();
        self.quick_open_selected = 0;
        self.filter_quick_open();
    }

    fn filter_quick_open(&mut self) {
        let query = self.quick_open_query.to_lowercase();
        if query.is_empty() {
            // Show all recent files
            self.quick_open_filtered = (0..self.recent_files.len()).collect();
        } else {
            // Fuzzy match on file names
            self.quick_open_filtered = self.recent_files.iter()
                .enumerate()
                .filter(|(_, path)| {
                    let name = path.file_name()
                        .and_then(|n| n.to_str())
                        .unwrap_or("")
                        .to_lowercase();
                    // Simple fuzzy: check if all query chars appear in order
                    let mut chars = query.chars().peekable();
                    for c in name.chars() {
                        if chars.peek() == Some(&c) {
                            chars.next();
                        }
                    }
                    chars.peek().is_none()
                })
                .map(|(i, _)| i)
                .collect();
        }
        // Clamp selection
        if !self.quick_open_filtered.is_empty() {
            self.quick_open_selected = self.quick_open_selected.min(self.quick_open_filtered.len() - 1);
        } else {
            self.quick_open_selected = 0;
        }
    }

    // === Find in Cells ===

    fn open_find(&mut self) {
        self.mode = Mode::Find;
        self.find_query.clear();
        self.find_results.clear();
        self.find_selected = 0;
    }

    fn search_cells(&mut self) {
        self.find_results.clear();
        let query = self.find_query.to_lowercase();
        if query.is_empty() {
            return;
        }

        // Search first 10000 rows for performance
        let scan_limit = 10000.min(NUM_ROWS);
        for row in 0..scan_limit {
            for col in 0..NUM_COLS {
                let display = self.sheet.get_display(row, col);
                if !display.is_empty() && display.to_lowercase().contains(&query) {
                    let cell_ref = format!("{}{}", Self::col_to_letters(col), row + 1);
                    let preview = if display.len() > 40 {
                        format!("{}...", &display[..40])
                    } else {
                        display
                    };
                    self.find_results.push((row, col, format!("{}: {}", cell_ref, preview)));
                }
                // Limit results for performance
                if self.find_results.len() >= 100 {
                    return;
                }
            }
        }
    }

    /// Record current cell values before modification (for undo)
    fn begin_undo_action(&self, description: &str, cells: &[(usize, usize)]) -> UndoAction {
        let mut action = UndoAction::new(description);
        for &(row, col) in cells {
            let old_value = self.sheet.get_raw(row, col);
            action.changes.push(CellChange {
                row,
                col,
                old_value,
                new_value: String::new(), // Will be filled in later
            });
        }
        action
    }

    /// Finalize undo action with new values and push to stack
    fn commit_undo_action(&mut self, mut action: UndoAction) {
        // Fill in new values
        for change in &mut action.changes {
            change.new_value = self.sheet.get_raw(change.row, change.col);
        }
        // Only push if something actually changed
        if action.changes.iter().any(|c| c.old_value != c.new_value) {
            self.undo_stack.push(action);
            self.redo_stack.clear(); // Clear redo stack on new action
        }
    }

    /// Perform undo
    fn undo(&mut self) {
        if let Some(action) = self.undo_stack.pop() {
            // Apply old values
            for change in &action.changes {
                self.sheet.set_value(change.row, change.col, &change.old_value);
            }
            self.status_message = Some(format!("Undo: {}", action.description));
            self.redo_stack.push(action);
            self.sync_input_from_selection();
        } else {
            self.status_message = Some("Nothing to undo".to_string());
        }
    }

    /// Perform redo
    fn redo(&mut self) {
        if let Some(action) = self.redo_stack.pop() {
            // Apply new values
            for change in &action.changes {
                self.sheet.set_value(change.row, change.col, &change.new_value);
            }
            self.status_message = Some(format!("Redo: {}", action.description));
            self.undo_stack.push(action);
            self.sync_input_from_selection();
        } else {
            self.status_message = Some("Nothing to redo".to_string());
        }
    }

    /// Convert iced Key + Modifiers to a keybinding lookup string
    fn key_to_string(key: &Key, modifiers: &Modifiers) -> Option<String> {
        let key_str = match key {
            Key::Character(ch) => ch.as_str().to_lowercase(),
            Key::Named(named) => {
                match named {
                    keyboard::key::Named::Enter => "enter".to_string(),
                    keyboard::key::Named::Tab => "tab".to_string(),
                    keyboard::key::Named::Escape => "escape".to_string(),
                    keyboard::key::Named::Backspace => "backspace".to_string(),
                    keyboard::key::Named::Delete => "delete".to_string(),
                    keyboard::key::Named::ArrowUp => "up".to_string(),
                    keyboard::key::Named::ArrowDown => "down".to_string(),
                    keyboard::key::Named::ArrowLeft => "left".to_string(),
                    keyboard::key::Named::ArrowRight => "right".to_string(),
                    keyboard::key::Named::Home => "home".to_string(),
                    keyboard::key::Named::End => "end".to_string(),
                    keyboard::key::Named::PageUp => "pageup".to_string(),
                    keyboard::key::Named::PageDown => "pagedown".to_string(),
                    keyboard::key::Named::F1 => "f1".to_string(),
                    keyboard::key::Named::F2 => "f2".to_string(),
                    keyboard::key::Named::F3 => "f3".to_string(),
                    keyboard::key::Named::F4 => "f4".to_string(),
                    keyboard::key::Named::F5 => "f5".to_string(),
                    keyboard::key::Named::F6 => "f6".to_string(),
                    keyboard::key::Named::F7 => "f7".to_string(),
                    keyboard::key::Named::F8 => "f8".to_string(),
                    keyboard::key::Named::F9 => "f9".to_string(),
                    keyboard::key::Named::F10 => "f10".to_string(),
                    keyboard::key::Named::F11 => "f11".to_string(),
                    keyboard::key::Named::F12 => "f12".to_string(),
                    _ => return None,
                }
            }
            _ => return None,
        };

        let mut result = String::new();
        if modifiers.control() { result.push_str("ctrl+"); }
        if modifiers.shift() { result.push_str("shift+"); }
        if modifiers.alt() { result.push_str("alt+"); }
        result.push_str(&key_str);

        Some(result)
    }

    /// Execute a command from the keybindings system
    fn execute_keybinding_command(&mut self, command: &str) -> Task<Message> {
        match command {
            // File operations
            "file.new" => {
                self.sheet = Sheet::new(100, 26);
                self.selection = Selection::new(0, 0);
                self.input_value = String::new();
                Task::none()
            }
            "file.open" => {
                Task::perform(
                    async {
                        rfd::AsyncFileDialog::new()
                            .add_filter("Spreadsheets", &["sheet", "csv"])
                            .add_filter("Native format", &["sheet"])
                            .add_filter("CSV files", &["csv"])
                            .add_filter("All files", &["*"])
                            .set_title("Open Spreadsheet")
                            .pick_file()
                            .await
                            .map(|h| h.path().to_path_buf())
                    },
                    Message::FileOpened
                )
            }
            "file.save" => {
                // If we have a current file, save to it; otherwise show save dialog
                if let Some(ref path) = self.current_file {
                    let sheet = self.sheet.clone();
                    let path = path.clone();
                    // Detect format from extension
                    let is_native = path.extension()
                        .map(|e| e.to_string_lossy().to_lowercase() == "sheet")
                        .unwrap_or(false);
                    if is_native {
                        Task::perform(
                            async move { native::save(&sheet, &path) },
                            Message::FileSaveResult
                        )
                    } else {
                        Task::perform(
                            async move { csv::export(&sheet, &path) },
                            Message::FileSaveResult
                        )
                    }
                } else {
                    // Default to native format for new files
                    Task::perform(
                        async {
                            rfd::AsyncFileDialog::new()
                                .add_filter("Native format", &["sheet"])
                                .add_filter("CSV files", &["csv"])
                                .set_title("Save As")
                                .set_file_name("spreadsheet.sheet")
                                .save_file()
                                .await
                                .map(|h| h.path().to_path_buf())
                        },
                        Message::FileSaved
                    )
                }
            }

            // Navigation
            "cell.moveRight" => {
                self.selection.move_by(0, 1, NUM_ROWS, NUM_COLS);
                self.sync_input_from_selection();
                Task::none()
            }
            "cell.moveLeft" => {
                self.selection.move_by(0, -1, NUM_ROWS, NUM_COLS);
                self.sync_input_from_selection();
                Task::none()
            }
            "cell.moveDown" => {
                self.selection.move_by(1, 0, NUM_ROWS, NUM_COLS);
                self.sync_input_from_selection();
                Task::none()
            }
            "cell.moveUp" => {
                self.selection.move_by(-1, 0, NUM_ROWS, NUM_COLS);
                self.sync_input_from_selection();
                Task::none()
            }
            "cell.goToStart" => {
                self.selection.select_cell(0, 0);
                self.sync_input_from_selection();
                Task::none()
            }
            "cell.goToEnd" => {
                // Go to last cell with data (simplified: go to bottom-right of grid)
                self.selection.select_cell(NUM_ROWS - 1, NUM_COLS - 1);
                self.sync_input_from_selection();
                Task::none()
            }

            // Editing
            "edit.copy" => {
                self.copy_selection();
                Task::none()
            }
            "edit.cut" => {
                self.cut_selection();
                Task::none()
            }
            "edit.paste" => {
                self.paste();
                Task::none()
            }
            "edit.undo" => {
                self.undo();
                Task::none()
            }
            "edit.redo" => {
                self.redo();
                Task::none()
            }
            "cell.clear" => {
                self.clear_selection();
                Task::none()
            }
            "cell.edit" => {
                self.enter_edit_mode()
            }

            // Selection
            "select.all" => {
                self.select_all();
                Task::none()
            }
            "select.toEnd" => {
                self.selection.extend_to(NUM_ROWS - 1, NUM_COLS - 1);
                Task::none()
            }

            // Command palette
            "commandPalette.toggle" => {
                self.open_palette();
                Task::none()
            }

            // Navigate
            "navigate.goto" => {
                self.open_goto();
                Task::none()
            }

            // Problems panel
            "problems.toggle" => {
                self.toggle_problems();
                Task::none()
            }

            // Quick Open
            "file.quickOpen" => {
                self.open_quick_open();
                Task::none()
            }

            // Find
            "edit.find" => {
                self.open_find();
                Task::none()
            }

            // Formatting
            "format.bold" => {
                self.toggle_format_bold();
                Task::none()
            }
            "format.italic" => {
                self.toggle_format_italic();
                Task::none()
            }
            "format.underline" => {
                self.toggle_format_underline();
                Task::none()
            }

            // Column/row sizing
            "column.autoSize" => {
                self.auto_size_selected_columns();
                Task::none()
            }

            // Data operations
            "data.fillDown" => {
                self.fill_down();
                Task::none()
            }
            "data.fillRight" => {
                self.fill_right();
                Task::none()
            }

            // Formula
            "formula.autosum" => {
                self.insert_auto_sum();
                Task::none()
            }

            // Theme
            "theme.toggle" => {
                self.dark_mode = !self.dark_mode;
                Task::none()
            }

            // Row/Column selection
            "select.column" => {
                self.selection.select_column(NUM_ROWS);
                self.status_message = Some("Selected column".to_string());
                Task::none()
            }
            "select.row" => {
                self.selection.select_row(NUM_COLS);
                self.status_message = Some("Selected row".to_string());
                Task::none()
            }

            // Insert/Delete rows and columns
            "edit.insertRowCol" => {
                self.insert_row_or_column();
                Task::none()
            }
            "edit.deleteRowCol" => {
                self.delete_row_or_column();
                Task::none()
            }

            // Insert date/time
            "edit.insertDate" => {
                self.insert_current_date();
                Task::none()
            }
            "edit.insertTime" => {
                self.insert_current_time();
                Task::none()
            }

            // Formula view toggle
            "view.toggleFormulas" => {
                self.show_formulas = !self.show_formulas;
                self.status_message = Some(if self.show_formulas {
                    "Formula view ON".to_string()
                } else {
                    "Formula view OFF".to_string()
                });
                Task::none()
            }

            // F4 - cycle cell reference
            "edit.cycleReference" => {
                self.cycle_cell_reference();
                Task::none()
            }

            // Split view
            "view.splitToggle" => {
                self.update(Message::SplitToggle)
            }
            "view.splitSwitch" => {
                self.update(Message::SplitSwitchPane)
            }
            "view.zenMode" => {
                self.update(Message::ZenModeToggle)
            }
            "view.inspector" => {
                self.update(Message::InspectorToggle)
            }

            _ => Task::none()
        }
    }

    pub fn update(&mut self, message: Message) -> Task<Message> {
        match message {
            Message::CellClicked(row, col, modifiers) => {
                // Close any open menu
                self.open_menu = None;

                if self.mode == Mode::Command {
                    return Task::none();
                }

                // If in edit mode, commit first
                if self.mode == Mode::Edit {
                    self.commit_edit();
                }

                // Handle selection based on modifiers
                if modifiers.control() && modifiers.shift() {
                    // Ctrl+Shift+Click: add range from anchor
                    self.selection.add_range_to(row, col);
                } else if modifiers.control() {
                    // Ctrl+Click: add to selection (discontiguous)
                    self.selection.add_cell(row, col);
                } else if modifiers.shift() {
                    // Shift+Click: extend from anchor
                    self.selection.extend_to(row, col);
                } else {
                    // Plain click: single cell
                    self.selection.select_cell(row, col);
                }

                self.sync_input_from_selection();
                self.mode = Mode::Navigation;
                Task::none()
            }

            Message::InputChanged(value) => {
                self.input_value = value;
                // If we're in navigation mode and user types in formula bar, enter edit mode
                if self.mode == Mode::Navigation {
                    self.mode = Mode::Edit;
                    self.edit_original = self.sheet.get_raw(
                        self.selection.active_cell().0,
                        self.selection.active_cell().1
                    );
                }
                // Update autocomplete suggestions
                self.update_autocomplete();
                Task::none()
            }

            Message::InputSubmitted => {
                self.commit_edit();
                self.move_after_commit(1, 0); // Move down
                Task::none()
            }

            Message::PaletteQueryChanged(query) => {
                self.palette_query = query;
                self.filter_commands();
                self.palette_selected = 0;
                Task::none()
            }

            Message::PaletteSubmit => {
                match self.palette_mode {
                    PaletteMode::Commands | PaletteMode::Mixed => {
                        // In mixed mode, check if we're selecting a command or a file
                        let cmd_count = self.palette_filtered.len();
                        if self.palette_selected < cmd_count {
                            if let Some(&idx) = self.palette_filtered.get(self.palette_selected) {
                                let cmd_id = COMMANDS[idx].id;
                                self.close_palette();
                                return self.execute_command(cmd_id);
                            }
                        } else {
                            // Selecting a file
                            let file_idx = self.palette_selected - cmd_count;
                            if let Some(&idx) = self.palette_files.get(file_idx) {
                                let path = self.recent_files[idx].clone();
                                let path_for_msg = path.clone();
                                self.close_palette();
                                return Task::perform(
                                    async move {
                                        if path.extension().map(|e| e.to_string_lossy().to_lowercase() == "sheet").unwrap_or(false) {
                                            native::load(&path).map_err(|e| e.to_string())
                                        } else {
                                            csv::import(&path).map_err(|e| e.to_string())
                                        }
                                    },
                                    move |result| Message::FileLoaded(result, path_for_msg.clone())
                                );
                            }
                        }
                    }
                    PaletteMode::Cells => {
                        if let Some((row, col, _)) = self.palette_cells.get(self.palette_selected) {
                            let r = *row;
                            let c = *col;
                            self.close_palette();
                            self.selection.select_cell(r, c);
                            self.sync_input_from_selection();
                            self.ensure_selection_visible();
                        }
                    }
                    PaletteMode::GoTo => {
                        let cell_ref = self.palette_query[1..].trim().to_string();
                        self.close_palette();
                        self.goto_cell(&cell_ref);
                    }
                    PaletteMode::Functions => {
                        if let Some(&idx) = self.palette_functions.get(self.palette_selected) {
                            let func = &FORMULA_FUNCTIONS[idx];
                            let syntax = func.syntax.to_string();
                            self.close_palette();
                            // Insert function syntax into formula bar and enter edit mode
                            self.input_value = syntax;
                            self.mode = Mode::Edit;
                            self.edit_original = self.sheet.get_raw(
                                self.selection.active_cell().0,
                                self.selection.active_cell().1
                            );
                        }
                    }
                }
                Task::none()
            }

            Message::ExecuteCommand(cmd_id) => {
                self.close_palette();
                self.execute_command(cmd_id)
            }

            Message::MenuClicked(menu) => {
                // Toggle menu open/close
                if self.open_menu == Some(menu) {
                    self.open_menu = None;
                } else {
                    self.open_menu = Some(menu);
                }
                Task::none()
            }

            Message::MenuItemClicked(item_id) => {
                self.open_menu = None;
                self.execute_command(item_id)
            }

            Message::KeyPressed(key, modifiers) => {
                match self.mode {
                    Mode::Command => self.handle_command_mode_key(key, modifiers),
                    Mode::Edit => {
                        self.handle_edit_mode_key(key, modifiers);
                        Task::none()
                    }
                    Mode::Navigation => self.handle_navigation_mode_key(key, modifiers),
                    Mode::GoTo => {
                        self.handle_goto_mode_key(key);
                        Task::none()
                    }
                    Mode::QuickOpen => {
                        self.handle_quick_open_mode_key(key)
                    }
                    Mode::Find => {
                        self.handle_find_mode_key(key)
                    }
                }
            }

            // File dialog results
            Message::FileOpened(path) => {
                if let Some(path) = path {
                    let path_clone = path.clone();
                    // Detect format from extension
                    let is_native = path.extension()
                        .map(|e| e.to_string_lossy().to_lowercase() == "sheet")
                        .unwrap_or(false);
                    if is_native {
                        Task::perform(
                            async move { native::load(&path) },
                            move |result| Message::FileLoaded(result, path_clone.clone())
                        )
                    } else {
                        Task::perform(
                            async move { csv::import(&path) },
                            move |result| Message::FileLoaded(result, path_clone.clone())
                        )
                    }
                } else {
                    Task::none()
                }
            }

            Message::FileLoaded(result, path) => {
                match result {
                    Ok(sheet) => {
                        self.sheet = sheet;
                        self.current_file = Some(path.clone());
                        self.selection = Selection::new(0, 0);
                        self.sync_input_from_selection();
                        self.add_recent_file(path.clone());
                        self.status_message = Some(format!("Opened: {}", path.display()));
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Error: {}", e));
                    }
                }
                Task::none()
            }

            Message::FileSaved(path) => {
                if let Some(path) = path {
                    let sheet = self.sheet.clone();
                    self.current_file = Some(path.clone());
                    // Detect format from extension
                    let is_native = path.extension()
                        .map(|e| e.to_string_lossy().to_lowercase() == "sheet")
                        .unwrap_or(false);
                    if is_native {
                        Task::perform(
                            async move { native::save(&sheet, &path) },
                            Message::FileSaveResult
                        )
                    } else {
                        Task::perform(
                            async move { csv::export(&sheet, &path) },
                            Message::FileSaveResult
                        )
                    }
                } else {
                    Task::none()
                }
            }

            Message::FileExported(path) => {
                if let Some(path) = path {
                    let sheet = self.sheet.clone();
                    Task::perform(
                        async move { csv::export(&sheet, &path) },
                        Message::FileSaveResult
                    )
                } else {
                    Task::none()
                }
            }

            Message::FileSaveResult(result) => {
                match result {
                    Ok(()) => {
                        if let Some(ref path) = self.current_file {
                            self.status_message = Some(format!("Saved: {}", path.display()));
                        } else {
                            self.status_message = Some("File saved".to_string());
                        }
                    }
                    Err(e) => {
                        self.status_message = Some(format!("Error saving: {}", e));
                    }
                }
                Task::none()
            }

            // Go to cell
            Message::GoToInputChanged(value) => {
                self.goto_input = value;
                Task::none()
            }

            Message::GoToSubmit => {
                let input = self.goto_input.clone();
                self.goto_cell(&input);
                Task::none()
            }

            Message::FormatBold => {
                self.toggle_format_bold();
                Task::none()
            }

            Message::FormatItalic => {
                self.toggle_format_italic();
                Task::none()
            }

            Message::FormatUnderline => {
                self.toggle_format_underline();
                Task::none()
            }
            Message::FormatAlignLeft => {
                self.set_alignment(Alignment::Left);
                Task::none()
            }
            Message::FormatAlignCenter => {
                self.set_alignment(Alignment::Center);
                Task::none()
            }
            Message::FormatAlignRight => {
                self.set_alignment(Alignment::Right);
                Task::none()
            }
            Message::FormatCurrency => {
                self.set_number_format(NumberFormat::Currency { decimals: 2 });
                Task::none()
            }
            Message::FormatPercent => {
                self.set_number_format(NumberFormat::Percent { decimals: 0 });
                Task::none()
            }
            Message::FormatDecimalIncrease => {
                self.increase_decimals();
                Task::none()
            }
            Message::FormatDecimalDecrease => {
                self.decrease_decimals();
                Task::none()
            }

            // Column resizing
            Message::ColResizeStart(col, x) => {
                self.resizing_col = Some(col);
                self.resize_start_x = x;
                self.resize_start_size = self.column_widths[col];
                Task::none()
            }
            Message::ColResizeMove(x) => {
                if let Some(col) = self.resizing_col {
                    let delta = x - self.resize_start_x;
                    let new_width = (self.resize_start_size + delta).max(30.0).min(500.0);
                    self.column_widths[col] = new_width;
                }
                Task::none()
            }
            Message::ColResizeEnd => {
                self.resizing_col = None;
                Task::none()
            }
            Message::ColAutoSize(col) => {
                self.auto_size_column(col);
                Task::none()
            }

            // Row resizing
            Message::RowResizeStart(row, y) => {
                self.resizing_row = Some(row);
                self.resize_start_y = y;
                self.resize_start_size = self.row_heights[row];
                Task::none()
            }
            Message::RowResizeMove(y) => {
                if let Some(row) = self.resizing_row {
                    let delta = y - self.resize_start_y;
                    let new_height = (self.resize_start_size + delta).max(16.0).min(200.0);
                    self.row_heights[row] = new_height;
                }
                Task::none()
            }
            Message::RowResizeEnd => {
                self.resizing_row = None;
                Task::none()
            }
            Message::RowAutoSize(row) => {
                self.auto_size_row(row);
                Task::none()
            }
            Message::ToggleProblems => {
                self.toggle_problems();
                Task::none()
            }
            Message::GoToProblem(row, col) => {
                self.selection.select_cell(row, col);
                self.sync_input_from_selection();
                self.ensure_selection_visible();
                self.status_message = Some(format!("Jumped to {}{}", Self::col_to_letters(col), row + 1));
                Task::none()
            }
            // Split view
            Message::SplitToggle => {
                self.split_enabled = !self.split_enabled;
                if self.split_enabled {
                    // Initialize split pane scroll to match primary
                    self.split_scroll_row = self.scroll_row;
                    self.split_scroll_col = self.scroll_col;
                    self.status_message = Some("Split view enabled".to_string());
                } else {
                    self.split_active_pane = SplitPane::Primary;
                    self.status_message = Some("Split view disabled".to_string());
                }
                Task::none()
            }
            Message::SplitHorizontal => {
                self.split_direction = SplitDirection::Horizontal;
                self.split_enabled = true;
                self.split_scroll_row = self.scroll_row;
                self.split_scroll_col = self.scroll_col;
                self.status_message = Some("Split horizontally (side by side)".to_string());
                Task::none()
            }
            Message::SplitVertical => {
                self.split_direction = SplitDirection::Vertical;
                self.split_enabled = true;
                self.split_scroll_row = self.scroll_row;
                self.split_scroll_col = self.scroll_col;
                self.status_message = Some("Split vertically (stacked)".to_string());
                Task::none()
            }
            Message::SplitSwitchPane => {
                if self.split_enabled {
                    self.split_active_pane = match self.split_active_pane {
                        SplitPane::Primary => SplitPane::Secondary,
                        SplitPane::Secondary => SplitPane::Primary,
                    };
                    self.status_message = Some(format!("Switched to {:?} pane", self.split_active_pane));
                }
                Task::none()
            }
            Message::SplitPaneClicked(pane) => {
                if self.split_enabled {
                    self.split_active_pane = pane;
                }
                Task::none()
            }
            // Zen mode
            Message::ZenModeToggle => {
                self.zen_mode = !self.zen_mode;
                self.status_message = Some(if self.zen_mode {
                    "Zen mode ON - press F11 or Escape to exit".to_string()
                } else {
                    "Zen mode OFF".to_string()
                });
                Task::none()
            }
            // Cell inspector
            Message::InspectorToggle => {
                self.show_inspector = !self.show_inspector;
                self.status_message = Some(if self.show_inspector {
                    "Cell Inspector ON".to_string()
                } else {
                    "Cell Inspector OFF".to_string()
                });
                Task::none()
            }
            // Autocomplete
            Message::AutocompleteSelect(idx) => {
                if idx < self.autocomplete_suggestions.len() {
                    self.autocomplete_selected = idx;
                }
                Task::none()
            }
            Message::AutocompleteApply => {
                self.apply_autocomplete();
                Task::none()
            }
            Message::AutocompleteDismiss => {
                self.autocomplete_visible = false;
                self.autocomplete_suggestions.clear();
                Task::none()
            }
            // Omarchy theme auto-reload (polling check)
            Message::OmarchyThemeChanged => {
                // Check if theme file actually changed
                let current_mtime = omarchy::theme_mtime();
                if current_mtime != self.omarchy_theme_mtime {
                    self.omarchy_theme_mtime = current_mtime;
                    self.reload_omarchy_theme();
                }
                Task::none()
            }
            // Quick Open
            Message::QuickOpenQueryChanged(query) => {
                self.quick_open_query = query;
                self.filter_quick_open();
                Task::none()
            }
            Message::QuickOpenSubmit => {
                if !self.quick_open_filtered.is_empty() {
                    let idx = self.quick_open_filtered[self.quick_open_selected];
                    let path = self.recent_files[idx].clone();
                    let path_for_msg = path.clone();
                    self.mode = Mode::Navigation;
                    // Load the file
                    return Task::perform(
                        async move {
                            if path.extension().map(|e| e.to_string_lossy().to_lowercase() == "sheet").unwrap_or(false) {
                                native::load(&path).map_err(|e| e.to_string())
                            } else {
                                csv::import(&path).map_err(|e| e.to_string())
                            }
                        },
                        move |result| Message::FileLoaded(result, path_for_msg.clone())
                    );
                }
                self.mode = Mode::Navigation;
                Task::none()
            }
            Message::QuickOpenSelect(idx) => {
                if idx < self.recent_files.len() {
                    let path = self.recent_files[idx].clone();
                    let path_for_msg = path.clone();
                    self.mode = Mode::Navigation;
                    return Task::perform(
                        async move {
                            if path.extension().map(|e| e.to_string_lossy().to_lowercase() == "sheet").unwrap_or(false) {
                                native::load(&path).map_err(|e| e.to_string())
                            } else {
                                csv::import(&path).map_err(|e| e.to_string())
                            }
                        },
                        move |result| Message::FileLoaded(result, path_for_msg.clone())
                    );
                }
                Task::none()
            }
            // Find
            Message::FindQueryChanged(query) => {
                self.find_query = query;
                self.search_cells();
                self.find_selected = 0;
                Task::none()
            }
            Message::FindSubmit | Message::FindNext => {
                if !self.find_results.is_empty() {
                    let (row, col, _) = self.find_results[self.find_selected].clone();
                    self.selection.select_cell(row, col);
                    self.sync_input_from_selection();
                    self.ensure_selection_visible();
                    // Move to next for next press
                    self.find_selected = (self.find_selected + 1) % self.find_results.len();
                }
                Task::none()
            }
            Message::FindPrev => {
                if !self.find_results.is_empty() {
                    let (row, col, _) = self.find_results[self.find_selected].clone();
                    self.selection.select_cell(row, col);
                    self.sync_input_from_selection();
                    self.ensure_selection_visible();
                    // Move to prev
                    if self.find_selected == 0 {
                        self.find_selected = self.find_results.len() - 1;
                    } else {
                        self.find_selected -= 1;
                    }
                }
                Task::none()
            }
            Message::FindSelect(row, col) => {
                self.selection.select_cell(row, col);
                self.sync_input_from_selection();
                self.ensure_selection_visible();
                self.mode = Mode::Navigation;
                Task::none()
            }
        }
    }

    fn handle_command_mode_key(&mut self, key: Key, _modifiers: Modifiers) -> Task<Message> {
        match &key {
            Key::Named(keyboard::key::Named::Escape) => {
                self.close_palette();
                Task::none()
            }
            Key::Named(keyboard::key::Named::ArrowDown) => {
                let max_items = self.palette_max_items();
                if max_items > 0 && self.palette_selected + 1 < max_items {
                    self.palette_selected += 1;
                }
                Task::none()
            }
            Key::Named(keyboard::key::Named::ArrowUp) => {
                if self.palette_selected > 0 {
                    self.palette_selected -= 1;
                }
                Task::none()
            }
            Key::Named(keyboard::key::Named::Enter) => {
                // Delegate to PaletteSubmit handler
                return self.update(Message::PaletteSubmit);
            }
            _ => Task::none()
        }
    }

    fn palette_max_items(&self) -> usize {
        match self.palette_mode {
            PaletteMode::Commands => self.palette_filtered.len().min(10),
            PaletteMode::Cells => self.palette_cells.len().min(10),
            PaletteMode::Functions => self.palette_functions.len().min(10),
            PaletteMode::GoTo => 0,
            PaletteMode::Mixed => {
                let cmd_limit = if self.palette_files.is_empty() { 8 } else { 5 };
                let cmd_count = self.palette_filtered.len().min(cmd_limit);
                let file_count = self.palette_files.len().min(5);
                cmd_count + file_count
            }
        }
    }

    fn handle_edit_mode_key(&mut self, key: Key, modifiers: Modifiers) {
        // Handle autocomplete navigation when visible
        if self.autocomplete_visible && !self.autocomplete_suggestions.is_empty() {
            match &key {
                Key::Named(keyboard::key::Named::Escape) => {
                    // Dismiss autocomplete
                    self.autocomplete_visible = false;
                    self.autocomplete_suggestions.clear();
                    return;
                }
                Key::Named(keyboard::key::Named::ArrowDown) => {
                    // Navigate down in autocomplete
                    if self.autocomplete_selected + 1 < self.autocomplete_suggestions.len() {
                        self.autocomplete_selected += 1;
                    }
                    return;
                }
                Key::Named(keyboard::key::Named::ArrowUp) => {
                    // Navigate up in autocomplete
                    if self.autocomplete_selected > 0 {
                        self.autocomplete_selected -= 1;
                    }
                    return;
                }
                Key::Named(keyboard::key::Named::Tab) | Key::Named(keyboard::key::Named::Enter) => {
                    // Apply autocomplete selection
                    self.apply_autocomplete();
                    return;
                }
                _ => {}
            }
        }

        match &key {
            Key::Named(keyboard::key::Named::Escape) => {
                // Cancel edit, restore original
                self.input_value = self.edit_original.clone();
                self.autocomplete_visible = false;
                self.autocomplete_suggestions.clear();
                self.mode = Mode::Navigation;
            }
            Key::Named(keyboard::key::Named::Enter) => {
                if modifiers.shift() {
                    self.commit_edit();
                    self.move_after_commit(-1, 0); // Move up
                } else {
                    self.commit_edit();
                    self.move_after_commit(1, 0); // Move down
                }
            }
            Key::Named(keyboard::key::Named::Tab) => {
                if modifiers.shift() {
                    self.commit_edit();
                    self.move_after_commit(0, -1); // Move left
                } else {
                    self.commit_edit();
                    self.move_after_commit(0, 1); // Move right
                }
            }
            _ => {
                // Let text input handle other keys
            }
        }
    }

    fn handle_goto_mode_key(&mut self, key: Key) {
        match &key {
            Key::Named(keyboard::key::Named::Escape) => {
                self.close_goto();
            }
            Key::Named(keyboard::key::Named::Enter) => {
                let input = self.goto_input.clone();
                self.goto_cell(&input);
            }
            _ => {
                // Let text input handle other keys
            }
        }
    }

    fn handle_quick_open_mode_key(&mut self, key: Key) -> Task<Message> {
        match &key {
            Key::Named(keyboard::key::Named::Escape) => {
                self.mode = Mode::Navigation;
            }
            Key::Named(keyboard::key::Named::Enter) => {
                if !self.quick_open_filtered.is_empty() {
                    let idx = self.quick_open_filtered[self.quick_open_selected];
                    let path = self.recent_files[idx].clone();
                    let path_for_msg = path.clone();
                    self.mode = Mode::Navigation;
                    return Task::perform(
                        async move {
                            if path.extension().map(|e| e.to_string_lossy().to_lowercase() == "sheet").unwrap_or(false) {
                                native::load(&path).map_err(|e| e.to_string())
                            } else {
                                csv::import(&path).map_err(|e| e.to_string())
                            }
                        },
                        move |result| Message::FileLoaded(result, path_for_msg.clone())
                    );
                }
            }
            Key::Named(keyboard::key::Named::ArrowUp) => {
                if self.quick_open_selected > 0 {
                    self.quick_open_selected -= 1;
                }
            }
            Key::Named(keyboard::key::Named::ArrowDown) => {
                if !self.quick_open_filtered.is_empty() &&
                   self.quick_open_selected < self.quick_open_filtered.len() - 1 {
                    self.quick_open_selected += 1;
                }
            }
            _ => {}
        }
        Task::none()
    }

    fn handle_find_mode_key(&mut self, key: Key) -> Task<Message> {
        match &key {
            Key::Named(keyboard::key::Named::Escape) => {
                self.mode = Mode::Navigation;
            }
            Key::Named(keyboard::key::Named::Enter) => {
                if !self.find_results.is_empty() {
                    let (row, col, _) = self.find_results[self.find_selected].clone();
                    self.selection.select_cell(row, col);
                    self.sync_input_from_selection();
                    self.ensure_selection_visible();
                    self.find_selected = (self.find_selected + 1) % self.find_results.len();
                }
            }
            Key::Named(keyboard::key::Named::ArrowUp) => {
                if self.find_selected > 0 {
                    self.find_selected -= 1;
                } else if !self.find_results.is_empty() {
                    self.find_selected = self.find_results.len() - 1;
                }
            }
            Key::Named(keyboard::key::Named::ArrowDown) => {
                if !self.find_results.is_empty() {
                    self.find_selected = (self.find_selected + 1) % self.find_results.len();
                }
            }
            _ => {}
        }
        Task::none()
    }

    fn handle_navigation_mode_key(&mut self, key: Key, modifiers: Modifiers) -> Task<Message> {
        // Escape closes menus or exits zen mode
        if let Key::Named(keyboard::key::Named::Escape) = &key {
            if self.zen_mode {
                self.zen_mode = false;
                self.status_message = Some("Zen mode OFF".to_string());
                return Task::none();
            }
            if self.open_menu.is_some() {
                self.open_menu = None;
                return Task::none();
            }
        }

        // Close menus on any navigation key
        self.open_menu = None;

        // Try keybinding lookup first
        if let Some(key_str) = Self::key_to_string(&key, &modifiers) {
            if let Some(command) = self.keybindings.get_command(&key_str).cloned() {
                return self.execute_keybinding_command(&command);
            }
        }

        // Navigation keys with special behavior (shift for extend, ctrl for jump)
        match &key {
            Key::Named(keyboard::key::Named::ArrowUp) => {
                if modifiers.control() {
                    // Ctrl+Arrow: jump navigation
                    let (row, col) = self.selection.active_cell();
                    let target_row = self.find_jump_target_vertical(row, col, -1);
                    if modifiers.shift() {
                        self.selection.extend_to(target_row, col);
                    } else {
                        self.selection.select_cell(target_row, col);
                    }
                } else if modifiers.shift() {
                    self.selection.extend_by(-1, 0, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(-1, 0, NUM_ROWS, NUM_COLS);
                }
                self.ensure_active_cell_visible();
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::ArrowDown) => {
                if modifiers.control() {
                    // Ctrl+Arrow: jump navigation
                    let (row, col) = self.selection.active_cell();
                    let target_row = self.find_jump_target_vertical(row, col, 1);
                    if modifiers.shift() {
                        self.selection.extend_to(target_row, col);
                    } else {
                        self.selection.select_cell(target_row, col);
                    }
                } else if modifiers.shift() {
                    self.selection.extend_by(1, 0, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(1, 0, NUM_ROWS, NUM_COLS);
                }
                self.ensure_active_cell_visible();
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::ArrowLeft) => {
                if modifiers.control() {
                    // Ctrl+Arrow: jump navigation
                    let (row, col) = self.selection.active_cell();
                    let target_col = self.find_jump_target_horizontal(row, col, -1);
                    if modifiers.shift() {
                        self.selection.extend_to(row, target_col);
                    } else {
                        self.selection.select_cell(row, target_col);
                    }
                } else if modifiers.shift() {
                    self.selection.extend_by(0, -1, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(0, -1, NUM_ROWS, NUM_COLS);
                }
                self.ensure_active_cell_visible();
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::ArrowRight) => {
                if modifiers.control() {
                    // TODO: Ctrl+Right crashes for unknown reason - disabled for now
                    // Just move one cell right as fallback
                    self.selection.move_by(0, 1, NUM_ROWS, NUM_COLS);
                } else if modifiers.shift() {
                    self.selection.extend_by(0, 1, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(0, 1, NUM_ROWS, NUM_COLS);
                }
                self.ensure_active_cell_visible();
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::Tab) => {
                if modifiers.shift() {
                    self.selection.move_by(0, -1, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(0, 1, NUM_ROWS, NUM_COLS);
                }
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::Enter) => {
                // Enter: enter edit mode preserving content
                self.enter_edit_mode()
            }
            Key::Named(keyboard::key::Named::Delete) | Key::Named(keyboard::key::Named::Backspace) => {
                self.clear_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::Space) => {
                if modifiers.control() && !modifiers.shift() {
                    // Ctrl+Space: select entire column
                    self.selection.select_column(NUM_ROWS);
                    self.status_message = Some("Selected column".to_string());
                    Task::none()
                } else if modifiers.shift() && !modifiers.control() {
                    // Shift+Space: select entire row
                    self.selection.select_row(NUM_COLS);
                    self.status_message = Some("Selected row".to_string());
                    Task::none()
                } else {
                    Task::none()
                }
            }
            Key::Named(keyboard::key::Named::PageDown) => {
                // Move down by visible rows
                let jump = VISIBLE_ROWS.saturating_sub(1);
                if modifiers.shift() {
                    self.selection.extend_by(jump as isize, 0, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(jump as isize, 0, NUM_ROWS, NUM_COLS);
                }
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Named(keyboard::key::Named::PageUp) => {
                // Move up by visible rows
                let jump = VISIBLE_ROWS.saturating_sub(1);
                if modifiers.shift() {
                    self.selection.extend_by(-(jump as isize), 0, NUM_ROWS, NUM_COLS);
                } else {
                    self.selection.move_by(-(jump as isize), 0, NUM_ROWS, NUM_COLS);
                }
                self.sync_input_from_selection();
                Task::none()
            }
            Key::Character(ch) => {
                // Typing starts edit mode and replaces selection
                if !modifiers.control() && !modifiers.alt() {
                    let c = ch.as_str();
                    if !c.is_empty() {
                        // Multi-edit: typing replaces ALL selected cells
                        if self.selection.cell_count() > 1 {
                            // Record for undo
                            let cells: Vec<_> = self.selection.all_cells().collect();
                            let action = self.begin_undo_action("Type in cells", &cells);

                            // Set all cells to this character
                            for (r, c_idx) in cells {
                                self.sheet.set_value(r, c_idx, ch.as_str());
                            }

                            self.commit_undo_action(action);
                            self.sync_input_from_selection();
                            Task::none()
                        } else {
                            // Single cell: enter edit mode and start with this character
                            self.input_value = ch.as_str().to_string();
                            self.enter_edit_mode()
                        }
                    } else {
                        Task::none()
                    }
                } else {
                    Task::none()
                }
            }
            _ => Task::none()
        }
    }

    fn enter_edit_mode(&mut self) -> Task<Message> {
        self.edit_original = self.sheet.get_raw(
            self.selection.active_cell().0,
            self.selection.active_cell().1
        );
        self.mode = Mode::Edit;
        text_input::focus(formula_input_id())
    }

    fn commit_edit(&mut self) {
        // Clear autocomplete
        self.autocomplete_visible = false;
        self.autocomplete_suggestions.clear();

        // Record for undo
        let cells: Vec<_> = self.selection.all_cells().collect();
        let action = self.begin_undo_action("Edit cells", &cells);

        // Apply edit to all selected cells (multi-edit)
        let value = self.input_value.clone();
        for (r, c) in cells {
            self.sheet.set_value(r, c, &value);
        }

        self.commit_undo_action(action);
        self.mode = Mode::Navigation;
    }

    fn move_after_commit(&mut self, d_row: isize, d_col: isize) {
        self.selection.move_by(d_row, d_col, NUM_ROWS, NUM_COLS);
        self.sync_input_from_selection();
    }

    fn sync_input_from_selection(&mut self) {
        let (row, col) = self.selection.active_cell();
        self.input_value = self.sheet.get_raw(row, col);
        self.ensure_selection_visible();
    }

    /// Scroll viewport to ensure active cell is visible
    fn ensure_selection_visible(&mut self) {
        let (row, col) = self.selection.active_cell();

        // Determine which scroll position to update based on active pane
        let (scroll_row, scroll_col) = if self.split_enabled && self.split_active_pane == SplitPane::Secondary {
            (&mut self.split_scroll_row, &mut self.split_scroll_col)
        } else {
            (&mut self.scroll_row, &mut self.scroll_col)
        };

        // Scroll vertically if needed
        if row < *scroll_row {
            *scroll_row = row;
        } else if row >= *scroll_row + VISIBLE_ROWS {
            *scroll_row = row.saturating_sub(VISIBLE_ROWS - 1);
        }

        // Scroll horizontally if needed
        if col < *scroll_col {
            *scroll_col = col;
        } else if col >= *scroll_col + VISIBLE_COLS {
            *scroll_col = col.saturating_sub(VISIBLE_COLS - 1);
        }
    }

    /// Ensure active cell is visible (alias for ensure_selection_visible)
    fn ensure_active_cell_visible(&mut self) {
        self.ensure_selection_visible();
    }

    /// Find jump target for Ctrl+Up/Down (Excel-style jump navigation)
    /// Direction: -1 for up, 1 for down
    fn find_jump_target_vertical(&self, row: usize, col: usize, direction: isize) -> usize {
        let max_row = NUM_ROWS - 1;

        // Bounds check input
        let row = row.min(max_row);
        let current_filled = !self.sheet.get_display(row, col).is_empty();
        let mut pos = row;

        if direction < 0 {
            // Moving up
            if pos == 0 {
                return 0;
            }

            if current_filled {
                // In filled cell: check if next cell is also filled
                let next_filled = pos > 0 && !self.sheet.get_display(pos - 1, col).is_empty();

                if next_filled {
                    // Both filled: scan until we hit empty or edge
                    pos -= 1;
                    while pos > 0 {
                        if self.sheet.get_display(pos - 1, col).is_empty() {
                            break;
                        }
                        pos -= 1;
                    }
                } else {
                    // Next is empty: scan for next filled cell
                    if pos > 0 {
                        pos -= 1;
                    }
                    while pos > 0 && self.sheet.get_display(pos, col).is_empty() {
                        pos -= 1;
                    }
                }
            } else {
                // In empty cell: scan for next filled cell
                if pos > 0 {
                    pos -= 1;
                }
                while pos > 0 && self.sheet.get_display(pos, col).is_empty() {
                    pos -= 1;
                }
            }
        } else {
            // Moving down
            if pos >= max_row {
                return max_row;
            }

            if current_filled {
                // In filled cell: check if next cell is also filled
                let next_filled = pos < max_row && !self.sheet.get_display(pos + 1, col).is_empty();

                if next_filled {
                    // Both filled: scan until we hit empty or edge
                    pos += 1;
                    while pos < max_row {
                        if self.sheet.get_display(pos + 1, col).is_empty() {
                            break;
                        }
                        pos += 1;
                    }
                } else {
                    // Next is empty: scan for next filled cell
                    pos += 1;
                    while pos < max_row && self.sheet.get_display(pos, col).is_empty() {
                        pos += 1;
                    }
                }
            } else {
                // In empty cell: scan for next filled cell
                pos += 1;
                while pos < max_row && self.sheet.get_display(pos, col).is_empty() {
                    pos += 1;
                }
            }
        }

        pos.min(max_row)
    }

    /// Find jump target for Ctrl+Left/Right (Excel-style jump navigation)
    /// Direction: -1 for left, 1 for right
    fn find_jump_target_horizontal(&self, row: usize, col: usize, direction: isize) -> usize {
        let max_col = NUM_COLS.saturating_sub(1);
        let col = col.min(max_col);
        let current_filled = !self.sheet.get_display(row, col).is_empty();
        let mut pos = col;

        if direction < 0 {
            // Moving left
            if pos == 0 {
                return 0;
            }

            if current_filled {
                let next_filled = pos > 0 && !self.sheet.get_display(row, pos - 1).is_empty();
                if next_filled {
                    pos = pos.saturating_sub(1);
                    while pos > 0 && !self.sheet.get_display(row, pos - 1).is_empty() {
                        pos = pos.saturating_sub(1);
                    }
                } else {
                    pos = pos.saturating_sub(1);
                    while pos > 0 && self.sheet.get_display(row, pos).is_empty() {
                        pos = pos.saturating_sub(1);
                    }
                }
            } else {
                while pos > 0 && self.sheet.get_display(row, pos - 1).is_empty() {
                    pos = pos.saturating_sub(1);
                }
                if pos > 0 {
                    pos = pos.saturating_sub(1);
                }
            }
        } else {
            // Moving right
            if pos >= max_col {
                return max_col;
            }

            if current_filled {
                let next_filled = pos < max_col && !self.sheet.get_display(row, pos + 1).is_empty();
                if next_filled {
                    pos = pos.saturating_add(1).min(max_col);
                    while pos < max_col && !self.sheet.get_display(row, pos + 1).is_empty() {
                        pos = pos.saturating_add(1).min(max_col);
                    }
                } else {
                    pos = pos.saturating_add(1).min(max_col);
                    while pos < max_col && self.sheet.get_display(row, pos).is_empty() {
                        pos = pos.saturating_add(1).min(max_col);
                    }
                }
            } else {
                while pos < max_col && self.sheet.get_display(row, pos + 1).is_empty() {
                    pos = pos.saturating_add(1).min(max_col);
                }
                if pos < max_col {
                    pos = pos.saturating_add(1).min(max_col);
                }
            }
        }

        pos.min(max_col)
    }

    fn clear_selection(&mut self) {
        let cells: Vec<_> = self.selection.all_cells().collect();
        let action = self.begin_undo_action("Clear cells", &cells);

        for (r, c) in cells {
            self.sheet.set_value(r, c, "");
        }

        self.commit_undo_action(action);
        self.sync_input_from_selection();
    }

    /// Insert row(s) or column(s) based on current selection
    /// If entire row is selected, insert rows. If entire column is selected, insert columns.
    fn insert_row_or_column(&mut self) {
        if self.selection.is_full_row(NUM_COLS) {
            // Insert rows
            let (start_row, end_row) = self.selection.row_range();
            let count = end_row - start_row + 1;
            self.sheet.insert_rows(start_row, count);
            self.status_message = Some(format!("Inserted {} row(s)", count));
        } else if self.selection.is_full_column(NUM_ROWS) {
            // Insert columns
            let (start_col, end_col) = self.selection.col_range();
            let count = end_col - start_col + 1;
            self.sheet.insert_cols(start_col, count);
            self.status_message = Some(format!("Inserted {} column(s)", count));
        } else {
            // Default: insert rows at current row
            let (row, _) = self.selection.active_cell();
            self.sheet.insert_rows(row, 1);
            self.status_message = Some("Inserted row".to_string());
        }
        // Clear undo/redo as these operations don't support undo yet
        self.undo_stack.clear();
        self.redo_stack.clear();
    }

    /// Delete row(s) or column(s) based on current selection
    /// If entire row is selected, delete rows. If entire column is selected, delete columns.
    fn delete_row_or_column(&mut self) {
        if self.selection.is_full_row(NUM_COLS) {
            // Delete rows
            let (start_row, end_row) = self.selection.row_range();
            let count = end_row - start_row + 1;
            self.sheet.delete_rows(start_row, count);
            // Move selection to stay in bounds
            let new_row = start_row.min(NUM_ROWS - 1);
            self.selection.select_cell(new_row, 0);
            self.status_message = Some(format!("Deleted {} row(s)", count));
        } else if self.selection.is_full_column(NUM_ROWS) {
            // Delete columns
            let (start_col, end_col) = self.selection.col_range();
            let count = end_col - start_col + 1;
            self.sheet.delete_cols(start_col, count);
            // Move selection to stay in bounds
            let new_col = start_col.min(NUM_COLS - 1);
            self.selection.select_cell(0, new_col);
            self.status_message = Some(format!("Deleted {} column(s)", count));
        } else {
            // Default: delete row at current row
            let (row, col) = self.selection.active_cell();
            self.sheet.delete_rows(row, 1);
            // Keep selection in same position
            let new_row = row.min(NUM_ROWS - 1);
            self.selection.select_cell(new_row, col);
            self.status_message = Some("Deleted row".to_string());
        }
        // Clear undo/redo as these operations don't support undo yet
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.sync_input_from_selection();
    }

    /// Insert current date (Ctrl+;) as hard-coded value
    fn insert_current_date(&mut self) {
        use std::time::{SystemTime, UNIX_EPOCH};

        // Get current date
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // Convert to date components (simplified calculation)
        let days_since_epoch = now / 86400;
        let mut year = 1970;
        let mut remaining_days = days_since_epoch as i64;

        loop {
            let days_in_year = if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) { 366 } else { 365 };
            if remaining_days < days_in_year {
                break;
            }
            remaining_days -= days_in_year;
            year += 1;
        }

        let days_in_months: [i64; 12] = if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) {
            [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        } else {
            [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        };

        let mut month = 1;
        for &days in &days_in_months {
            if remaining_days < days {
                break;
            }
            remaining_days -= days;
            month += 1;
        }
        let day = remaining_days + 1;

        let date_str = format!("{:04}-{:02}-{:02}", year, month, day);

        // Insert into all selected cells
        let cells: Vec<_> = self.selection.all_cells().collect();
        let action = self.begin_undo_action("Insert date", &cells);
        for (r, c) in cells {
            self.sheet.set_value(r, c, &date_str);
        }
        self.commit_undo_action(action);
        self.sync_input_from_selection();
        self.status_message = Some(format!("Inserted date: {}", date_str));
    }

    /// Insert current time (Ctrl+Shift+;) as hard-coded value
    fn insert_current_time(&mut self) {
        use std::time::{SystemTime, UNIX_EPOCH};

        // Get current time
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // Extract time components (UTC)
        let seconds_today = now % 86400;
        let hours = seconds_today / 3600;
        let minutes = (seconds_today % 3600) / 60;
        let seconds = seconds_today % 60;

        let time_str = format!("{:02}:{:02}:{:02}", hours, minutes, seconds);

        // Insert into all selected cells
        let cells: Vec<_> = self.selection.all_cells().collect();
        let action = self.begin_undo_action("Insert time", &cells);
        for (r, c) in cells {
            self.sheet.set_value(r, c, &time_str);
        }
        self.commit_undo_action(action);
        self.sync_input_from_selection();
        self.status_message = Some(format!("Inserted time: {}", time_str));
    }

    /// Cycle cell reference locking (F4) - cycles through A1 -> $A$1 -> A$1 -> $A1 -> A1
    fn cycle_cell_reference(&mut self) {
        if self.mode != Mode::Edit {
            return;
        }

        // Find cell reference at or before cursor position in input_value
        // For simplicity, we'll cycle the last cell reference in the formula
        let input = &self.input_value;
        if !input.starts_with('=') {
            return;
        }

        // Find cell references (pattern: optional $, letter(s), optional $, number(s))
        let mut result = String::new();
        let mut i = 0;
        let chars: Vec<char> = input.chars().collect();
        let mut last_ref_start = None;
        let mut last_ref_end = None;

        while i < chars.len() {
            let start = i;

            // Check for cell reference: [$]?[A-Z]+[$]?[0-9]+
            let has_col_dollar = chars.get(i) == Some(&'$');
            if has_col_dollar {
                i += 1;
            }

            // Read column letters
            let col_start = i;
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                i += 1;
            }
            let col_len = i - col_start;

            if col_len == 0 {
                // Not a cell reference, copy character and continue
                i = start;
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // Check for row dollar
            let has_row_dollar = chars.get(i) == Some(&'$');
            if has_row_dollar {
                i += 1;
            }

            // Read row numbers
            let row_start = i;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
            let row_len = i - row_start;

            if row_len == 0 {
                // Not a valid cell reference (no row number)
                i = start;
                result.push(chars[i]);
                i += 1;
                continue;
            }

            // Valid cell reference found
            last_ref_start = Some(result.len());
            for j in start..i {
                result.push(chars[j]);
            }
            last_ref_end = Some(result.len());
        }

        // Now cycle the last reference if found
        if let (Some(start), Some(end)) = (last_ref_start, last_ref_end) {
            let ref_str = &result[start..end];
            let cycled = Self::cycle_reference(ref_str);

            let mut new_result = result[..start].to_string();
            new_result.push_str(&cycled);
            new_result.push_str(&result[end..]);

            self.input_value = new_result;
            self.status_message = Some(format!("Reference: {}", cycled));
        }
    }

    /// Cycle a single cell reference through locking states
    fn cycle_reference(ref_str: &str) -> String {
        let chars: Vec<char> = ref_str.chars().collect();
        let mut i = 0;

        let has_col_dollar = chars.get(0) == Some(&'$');
        if has_col_dollar {
            i += 1;
        }

        // Read column
        let col_start = i;
        while i < chars.len() && chars[i].is_ascii_alphabetic() {
            i += 1;
        }
        let col: String = chars[col_start..i].iter().collect();

        let has_row_dollar = chars.get(i) == Some(&'$');
        if has_row_dollar {
            i += 1;
        }

        // Read row
        let row: String = chars[i..].iter().collect();

        // Cycle: A1 -> $A$1 -> A$1 -> $A1 -> A1
        match (has_col_dollar, has_row_dollar) {
            (false, false) => format!("${}${}", col, row),  // A1 -> $A$1
            (true, true) => format!("{}${}", col, row),     // $A$1 -> A$1
            (false, true) => format!("${}{}", col, row),    // A$1 -> $A1
            (true, false) => format!("{}{}", col, row),     // $A1 -> A1
        }
    }

    fn toggle_format_bold(&mut self) {
        for (r, c) in self.selection.all_cells() {
            self.sheet.toggle_bold(r, c);
        }
        self.status_message = Some("Toggled bold".to_string());
    }

    fn toggle_format_italic(&mut self) {
        for (r, c) in self.selection.all_cells() {
            self.sheet.toggle_italic(r, c);
        }
        self.status_message = Some("Toggled italic".to_string());
    }

    fn toggle_format_underline(&mut self) {
        for (r, c) in self.selection.all_cells() {
            self.sheet.toggle_underline(r, c);
        }
        self.status_message = Some("Toggled underline".to_string());
    }

    fn set_alignment(&mut self, alignment: Alignment) {
        for (r, c) in self.selection.all_cells() {
            self.sheet.set_alignment(r, c, alignment);
        }
        let name = match alignment {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
        };
        self.status_message = Some(format!("Aligned {}", name));
    }

    fn set_number_format(&mut self, number_format: NumberFormat) {
        for (r, c) in self.selection.all_cells() {
            self.sheet.set_number_format(r, c, number_format);
        }
        let name = match number_format {
            NumberFormat::General => "general",
            NumberFormat::Number { .. } => "number",
            NumberFormat::Currency { .. } => "currency",
            NumberFormat::Percent { .. } => "percent",
        };
        self.status_message = Some(format!("Applied {} format", name));
    }

    fn increase_decimals(&mut self) {
        for (r, c) in self.selection.all_cells() {
            let current = self.sheet.get_number_format(r, c);
            let new_format = match current {
                NumberFormat::General => NumberFormat::Number { decimals: 1 },
                NumberFormat::Number { decimals } => NumberFormat::Number { decimals: (decimals + 1).min(10) },
                NumberFormat::Currency { decimals } => NumberFormat::Currency { decimals: (decimals + 1).min(10) },
                NumberFormat::Percent { decimals } => NumberFormat::Percent { decimals: (decimals + 1).min(10) },
            };
            self.sheet.set_number_format(r, c, new_format);
        }
        self.status_message = Some("Increased decimal places".to_string());
    }

    fn decrease_decimals(&mut self) {
        for (r, c) in self.selection.all_cells() {
            let current = self.sheet.get_number_format(r, c);
            let new_format = match current {
                NumberFormat::General => NumberFormat::General,
                NumberFormat::Number { decimals } => NumberFormat::Number { decimals: decimals.saturating_sub(1) },
                NumberFormat::Currency { decimals } => NumberFormat::Currency { decimals: decimals.saturating_sub(1) },
                NumberFormat::Percent { decimals } => NumberFormat::Percent { decimals: decimals.saturating_sub(1) },
            };
            self.sheet.set_number_format(r, c, new_format);
        }
        self.status_message = Some("Decreased decimal places".to_string());
    }

    /// Auto-size column width based on content
    fn auto_size_column(&mut self, col: usize) {
        // Estimate width based on character count (rough approximation)
        // Each character is roughly 7-8 pixels wide at size 12
        const CHAR_WIDTH: f32 = 7.5;
        const MIN_WIDTH: f32 = 40.0;
        const MAX_WIDTH: f32 = 400.0;
        const PADDING: f32 = 16.0;

        let mut max_len = 0usize;

        // Only scan first 1000 rows for performance (covers most real data)
        let scan_limit = 1000.min(NUM_ROWS);
        for row in 0..scan_limit {
            let content = self.sheet.get_display(row, col);
            if content.len() > max_len {
                max_len = content.len();
            }
        }

        // Also check column header width
        let header_len = Self::col_to_letters(col).len();
        if header_len > max_len {
            max_len = header_len;
        }

        let width = ((max_len as f32) * CHAR_WIDTH + PADDING).clamp(MIN_WIDTH, MAX_WIDTH);
        self.column_widths[col] = width;
        self.status_message = Some(format!("Auto-sized column {}", Self::col_to_letters(col)));
    }

    /// Auto-size all columns in selection
    fn auto_size_selected_columns(&mut self) {
        // Get unique columns from selection
        let mut cols: Vec<usize> = self.selection.ranges()
            .iter()
            .flat_map(|r| r.start_col..=r.end_col)
            .collect();
        cols.sort();
        cols.dedup();

        for col in &cols {
            self.auto_size_column_silent(*col);
        }

        if cols.len() == 1 {
            self.status_message = Some(format!("Auto-sized column {}", Self::col_to_letters(cols[0])));
        } else {
            self.status_message = Some(format!("Auto-sized {} columns", cols.len()));
        }
    }

    /// Auto-size column without status message (for batch operations)
    fn auto_size_column_silent(&mut self, col: usize) {
        const CHAR_WIDTH: f32 = 7.5;
        const MIN_WIDTH: f32 = 40.0;
        const MAX_WIDTH: f32 = 400.0;
        const PADDING: f32 = 16.0;

        let mut max_len = 0usize;
        let scan_limit = 1000.min(NUM_ROWS);
        for row in 0..scan_limit {
            let content = self.sheet.get_display(row, col);
            if content.len() > max_len {
                max_len = content.len();
            }
        }

        let header_len = Self::col_to_letters(col).len();
        if header_len > max_len {
            max_len = header_len;
        }

        let width = ((max_len as f32) * CHAR_WIDTH + PADDING).clamp(MIN_WIDTH, MAX_WIDTH);
        self.column_widths[col] = width;
    }

    /// Auto-size row height based on content
    fn auto_size_row(&mut self, row: usize) {
        // For now, just reset to default height
        // Multi-line content would need more sophisticated handling
        self.row_heights[row] = CELL_HEIGHT;
    }

    /// Scan sheet for formula errors and update problems cache
    fn scan_for_errors(&mut self) {
        self.problems_cache.clear();

        // Scan first 10000 rows for performance
        let scan_limit = 10000.min(NUM_ROWS);

        for row in 0..scan_limit {
            for col in 0..NUM_COLS {
                let display = self.sheet.get_display(row, col);
                if display.starts_with('#') {
                    let (error_type, description) = match display.as_str() {
                        "#ERR" | "#ERR!" => ("#ERR!", "Formula error".to_string()),
                        "#REF!" => ("#REF!", "Invalid cell reference".to_string()),
                        "#DIV/0!" => ("#DIV/0!", "Division by zero".to_string()),
                        "#NAME?" => ("#NAME?", "Unknown function or name".to_string()),
                        "#VALUE!" => ("#VALUE!", "Wrong value type".to_string()),
                        "#NULL!" => ("#NULL!", "Null reference".to_string()),
                        "#N/A" => ("#N/A", "Value not available".to_string()),
                        "#CIRC!" => ("#CIRC!", "Circular reference".to_string()),
                        _ => (&display[..], "Unknown error".to_string()),
                    };

                    self.problems_cache.push(CellError {
                        row,
                        col,
                        error_type: error_type.to_string(),
                        description,
                    });
                }
            }
        }
    }

    /// Toggle problems panel visibility
    fn toggle_problems(&mut self) {
        self.show_problems = !self.show_problems;
        if self.show_problems {
            self.scan_for_errors();
        }
    }

    /// Reload theme from Omarchy system settings
    fn reload_omarchy_theme(&mut self) {
        if omarchy::is_omarchy() {
            self.omarchy_theme = Some(omarchy::load_theme());
            if let Some(name) = omarchy::current_theme_name() {
                self.status_message = Some(format!("Theme: {}", name));
            }
        } else {
            self.status_message = Some("Not running on Omarchy".to_string());
        }
    }

    /// Open settings file in default editor
    fn open_settings_file(&mut self) {
        let path = Settings::config_path();
        if let Err(e) = open::that(&path) {
            self.status_message = Some(format!("Failed to open settings: {}", e));
        } else {
            self.status_message = Some(format!("Opened: {}", path.display()));
        }
    }

    /// Open keybindings file in default editor
    fn open_keybindings_file(&mut self) {
        let path = self.keybindings.config_path().clone();
        if let Err(e) = open::that(&path) {
            self.status_message = Some(format!("Failed to open keybindings: {}", e));
        } else {
            self.status_message = Some(format!("Opened: {}", path.display()));
        }
    }

    fn copy_selection(&mut self) {
        // For now, just copy active cell. TODO: copy rectangular region
        let (r, c) = self.selection.active_cell();
        self.clipboard = Some(self.sheet.get_raw(r, c));
    }

    fn cut_selection(&mut self) {
        self.copy_selection();
        self.clear_selection();
    }

    fn paste(&mut self) {
        if let Some(ref value) = self.clipboard.clone() {
            // Record for undo
            let cells: Vec<_> = self.selection.all_cells().collect();
            let action = self.begin_undo_action("Paste", &cells);

            // If single value in clipboard, fill all selected cells
            for (r, c) in cells {
                self.sheet.set_value(r, c, &value);
            }

            self.commit_undo_action(action);
            self.sync_input_from_selection();
        }
    }

    fn fill_down(&mut self) {
        // Collect cells that will be modified (all except top row of each range)
        let mut cells_to_modify = Vec::new();
        for range in self.selection.ranges() {
            for col in range.start_col..=range.end_col {
                for row in (range.start_row + 1)..=range.end_row {
                    cells_to_modify.push((row, col));
                }
            }
        }

        if cells_to_modify.is_empty() {
            return;
        }

        let action = self.begin_undo_action("Fill down", &cells_to_modify);

        // Copy top row of selection to all rows below
        for range in self.selection.ranges().to_vec() {
            for col in range.start_col..=range.end_col {
                let source_value = self.sheet.get_raw(range.start_row, col);
                for row in (range.start_row + 1)..=range.end_row {
                    self.sheet.set_value(row, col, &source_value);
                }
            }
        }

        self.commit_undo_action(action);
        self.sync_input_from_selection();
    }

    fn fill_right(&mut self) {
        // Collect cells that will be modified (all except left column of each range)
        let mut cells_to_modify = Vec::new();
        for range in self.selection.ranges() {
            for row in range.start_row..=range.end_row {
                for col in (range.start_col + 1)..=range.end_col {
                    cells_to_modify.push((row, col));
                }
            }
        }

        if cells_to_modify.is_empty() {
            return;
        }

        let action = self.begin_undo_action("Fill right", &cells_to_modify);

        // Copy left column of selection to all columns right
        for range in self.selection.ranges().to_vec() {
            for row in range.start_row..=range.end_row {
                let source_value = self.sheet.get_raw(row, range.start_col);
                for col in (range.start_col + 1)..=range.end_col {
                    self.sheet.set_value(row, col, &source_value);
                }
            }
        }

        self.commit_undo_action(action);
        self.sync_input_from_selection();
    }

    fn select_all(&mut self) {
        // Select all cells with data (find extent)
        // For now, just select a reasonable range
        self.selection = Selection::new(0, 0);
        self.selection.extend_to(NUM_ROWS - 1, NUM_COLS - 1);
    }

    fn open_palette(&mut self) {
        self.mode = Mode::Command;
        self.palette_query.clear();
        self.filter_commands();
        self.palette_selected = 0;
    }

    fn close_palette(&mut self) {
        self.mode = Mode::Navigation;
        self.palette_query.clear();
    }

    fn open_goto(&mut self) {
        self.mode = Mode::GoTo;
        self.goto_input.clear();
    }

    fn close_goto(&mut self) {
        self.mode = Mode::Navigation;
        self.goto_input.clear();
    }

    /// Parse cell reference like "A1", "B25", "AA100" into (row, col)
    fn parse_cell_ref(input: &str) -> Option<(usize, usize)> {
        let input = input.trim().to_uppercase();
        if input.is_empty() {
            return None;
        }

        // Find where letters end and digits begin
        let letter_end = input.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 {
            return None;
        }

        let col_str = &input[..letter_end];
        let row_str = &input[letter_end..];

        if row_str.is_empty() {
            return None;
        }

        // Parse column: A=0, B=1, ..., Z=25, AA=26, etc.
        let col: usize = col_str
            .chars()
            .fold(0, |acc, c| acc * 26 + (c as usize - 'A' as usize + 1)) - 1;

        // Parse row (1-indexed in user input)
        let row: usize = row_str.parse::<usize>().ok()?.checked_sub(1)?;

        Some((row, col))
    }

    fn goto_cell(&mut self, input: &str) {
        if let Some((row, col)) = Self::parse_cell_ref(input) {
            // Clamp to grid bounds
            let row = row.min(NUM_ROWS - 1);
            let col = col.min(NUM_COLS - 1);
            self.selection.select_cell(row, col);
            self.sync_input_from_selection();
            self.status_message = Some(format!("Jumped to {}", input.trim().to_uppercase()));
        } else {
            self.status_message = Some(format!("Invalid cell reference: {}", input));
        }
        self.close_goto();
    }

    fn filter_commands(&mut self) {
        let query = &self.palette_query;

        // Detect mode from prefix
        if query.starts_with('>') {
            self.palette_mode = PaletteMode::Commands;
            let search = query[1..].trim().to_lowercase();
            self.filter_commands_only(&search);
        } else if query.starts_with('@') {
            self.palette_mode = PaletteMode::Cells;
            let search = query[1..].trim().to_lowercase();
            self.filter_cells(&search);
        } else if query.starts_with(':') {
            self.palette_mode = PaletteMode::GoTo;
            // GoTo mode doesn't need filtering, just stores the cell reference
        } else if query.starts_with('=') {
            self.palette_mode = PaletteMode::Functions;
            let search = query[1..].trim().to_lowercase();
            self.filter_functions(&search);
        } else {
            self.palette_mode = PaletteMode::Mixed;
            let search = query.to_lowercase();
            self.filter_mixed(&search);
        }
    }

    fn filter_commands_only(&mut self, query: &str) {
        self.palette_cells.clear();
        self.palette_files.clear();

        if query.is_empty() {
            self.palette_filtered = (0..COMMANDS.len()).collect();
        } else {
            self.palette_filtered = COMMANDS
                .iter()
                .enumerate()
                .filter(|(_, cmd)| {
                    cmd.label.to_lowercase().contains(query)
                        || cmd.id.to_lowercase().contains(query)
                })
                .map(|(i, _)| i)
                .collect();
        }
    }

    fn filter_cells(&mut self, query: &str) {
        self.palette_filtered.clear();
        self.palette_files.clear();
        self.palette_cells.clear();

        if query.is_empty() {
            return;
        }

        // Search cells (limit for performance)
        let scan_limit = 5000.min(NUM_ROWS);
        for row in 0..scan_limit {
            for col in 0..NUM_COLS {
                let display = self.sheet.get_display(row, col);
                if !display.is_empty() && display.to_lowercase().contains(query) {
                    let cell_ref = format!("{}{}", Self::col_to_letters(col), row + 1);
                    let preview = if display.len() > 30 {
                        format!("{}...", &display[..30])
                    } else {
                        display
                    };
                    self.palette_cells.push((row, col, format!("{}: {}", cell_ref, preview)));
                }
                if self.palette_cells.len() >= 20 {
                    return;
                }
            }
        }
    }

    fn filter_functions(&mut self, query: &str) {
        self.palette_filtered.clear();
        self.palette_cells.clear();
        self.palette_files.clear();
        self.palette_functions.clear();

        if query.is_empty() {
            // Show all functions when query is empty
            self.palette_functions = (0..FORMULA_FUNCTIONS.len()).collect();
        } else {
            // Fuzzy match on name, syntax, description, and category
            self.palette_functions = FORMULA_FUNCTIONS
                .iter()
                .enumerate()
                .filter(|(_, func)| {
                    func.name.to_lowercase().contains(query)
                        || func.syntax.to_lowercase().contains(query)
                        || func.description.to_lowercase().contains(query)
                        || func.category.to_lowercase().contains(query)
                })
                .map(|(i, _)| i)
                .collect();
        }
    }

    fn filter_mixed(&mut self, query: &str) {
        self.palette_cells.clear();

        // Filter commands
        if query.is_empty() {
            self.palette_filtered = (0..COMMANDS.len()).collect();
            self.palette_files = (0..self.recent_files.len().min(5)).collect();
        } else {
            // Filter commands
            self.palette_filtered = COMMANDS
                .iter()
                .enumerate()
                .filter(|(_, cmd)| {
                    cmd.label.to_lowercase().contains(query)
                        || cmd.id.to_lowercase().contains(query)
                })
                .map(|(i, _)| i)
                .collect();

            // Filter recent files
            self.palette_files = self.recent_files
                .iter()
                .enumerate()
                .filter(|(_, path)| {
                    path.file_name()
                        .and_then(|n| n.to_str())
                        .map(|s| s.to_lowercase().contains(query))
                        .unwrap_or(false)
                })
                .map(|(i, _)| i)
                .take(5)
                .collect();
        }
    }

    fn execute_command(&mut self, cmd_id: &str) -> Task<Message> {
        match cmd_id {
            "undo" => { self.undo(); Task::none() }
            "redo" => { self.redo(); Task::none() }
            "copy" => { self.copy_selection(); Task::none() }
            "cut" => { self.cut_selection(); Task::none() }
            "paste" => { self.paste(); Task::none() }
            "delete" => { self.clear_selection(); Task::none() }
            "sum" => { self.insert_auto_sum(); Task::none() }
            "theme_toggle" => { self.dark_mode = !self.dark_mode; Task::none() }
            "theme_reload" => { self.reload_omarchy_theme(); Task::none() }
            "settings_open" => { self.open_settings_file(); Task::none() }
            "keybindings_open" => { self.open_keybindings_file(); Task::none() }
            "problems_toggle" => { self.toggle_problems(); Task::none() }
            "split_toggle" => { return self.update(Message::SplitToggle); }
            "split_horizontal" => { return self.update(Message::SplitHorizontal); }
            "split_vertical" => { return self.update(Message::SplitVertical); }
            "split_switch" => { return self.update(Message::SplitSwitchPane); }
            "zen_mode" => { return self.update(Message::ZenModeToggle); }
            "inspector" => { return self.update(Message::InspectorToggle); }
            "quick_open" => { self.open_quick_open(); Task::none() }
            "find" => { self.open_find(); Task::none() }
            "fill_down" => { self.fill_down(); Task::none() }
            "fill_right" => { self.fill_right(); Task::none() }
            "select_all" => { self.select_all(); Task::none() }
            "new" => {
                self.sheet = Sheet::new(100, 26);
                self.current_file = None;
                self.selection = Selection::new(0, 0);
                self.input_value = String::new();
                self.status_message = Some("New workbook".to_string());
                Task::none()
            }
            "open" => self.execute_keybinding_command("file.open"),
            "save" => self.execute_keybinding_command("file.save"),
            "save_as" | "export_csv" => {
                Task::perform(
                    async {
                        rfd::AsyncFileDialog::new()
                            .add_filter("CSV files", &["csv"])
                            .set_title("Export as CSV")
                            .set_file_name("spreadsheet.csv")
                            .save_file()
                            .await
                            .map(|h| h.path().to_path_buf())
                    },
                    Message::FileExported
                )
            }
            "command_palette" => { self.open_palette(); Task::none() }
            "goto" => { self.open_goto(); Task::none() }
            _ => Task::none()
        }
    }

    fn insert_auto_sum(&mut self) {
        let (r, c) = self.selection.active_cell();

        // Determine the formula to insert
        let formula = {
            // Try vertical range first (cells above)
            let mut start_row = r;
            if r > 0 {
                for check_r in (0..r).rev() {
                    if self.sheet.get_raw(check_r, c).is_empty() {
                        break;
                    }
                    start_row = check_r;
                }
            }

            if start_row < r {
                let col_letter = (b'A' + c as u8) as char;
                Some(format!("=SUM({}{}:{}{})", col_letter, start_row + 1, col_letter, r))
            } else {
                // Try horizontal range (cells to the left)
                let mut start_col = c;
                if c > 0 {
                    for check_c in (0..c).rev() {
                        if self.sheet.get_raw(r, check_c).is_empty() {
                            break;
                        }
                        start_col = check_c;
                    }
                }

                if start_col < c {
                    let start_letter = (b'A' + start_col as u8) as char;
                    let end_letter = (b'A' + (c - 1) as u8) as char;
                    Some(format!("=SUM({}{}:{}{})", start_letter, r + 1, end_letter, r + 1))
                } else {
                    None
                }
            }
        };

        if let Some(formula) = formula {
            // Record for undo
            let action = self.begin_undo_action("AutoSum", &[(r, c)]);
            self.input_value = formula.clone();
            self.sheet.set_value(r, c, &formula);
            self.commit_undo_action(action);
        } else {
            self.input_value = "=SUM()".to_string();
        }
    }

    pub fn view(&self) -> Element<'_, Message> {
        let theme = self.theme();
        let (sel_row, sel_col) = self.selection.active_cell();
        let col_letter = (b'A' + sel_col as u8) as char;
        let cell_ref = format!("{}{}", col_letter, sel_row + 1);

        // Mode indicator
        let mode_text = match self.mode {
            Mode::Navigation => "NAV",
            Mode::Edit => "EDIT",
            Mode::Command => "CMD",
            Mode::GoTo => "GOTO",
            Mode::QuickOpen => "OPEN",
            Mode::Find => "FIND",
        };

        // Name box (cell reference)
        let name_box = container(
            text(cell_ref).size(13).color(theme.text).center()
        )
        .width(70)
        .padding(6)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_input)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            ..Default::default()
        });

        // Formula bar
        let formula_input = container(
            text_input("", &self.input_value)
                .id(formula_input_id())
                .on_input(Message::InputChanged)
                .on_submit(Message::InputSubmitted)
                .size(13)
                .padding(6)
                .style(move |_, _| text_input::Style {
                    background: Background::Color(theme.bg_input),
                    border: Border {
                        color: theme.border,
                        width: 1.0,
                        radius: 2.0.into(),
                    },
                    icon: theme.text_dim,
                    placeholder: theme.text_dim,
                    value: theme.text,
                    selection: theme.selected,
                })
        )
        .width(Length::Fill);

        // Mode indicator badge
        let mode_badge = container(
            text(mode_text).size(10).color(if self.mode == Mode::Edit {
                theme.accent
            } else {
                theme.text_dim
            }).center()
        )
        .width(40)
        .padding(6)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            ..Default::default()
        });

        let formula_bar = container(
            row![name_box, formula_input, mode_badge].spacing(4)
        )
        .padding(6)
        .width(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.gridline,
                width: 0.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        });

        // Show highlighted formula preview when editing a formula
        let formula_preview: Option<Element<'_, Message>> = if is_formula(&self.input_value) {
            Some(
                container(highlight_formula(&self.input_value, self.dark_mode))
                    .padding(iced::Padding { top: 2.0, right: 10.0, bottom: 4.0, left: 80.0 })
                    .style(move |_| container::Style {
                        background: Some(Background::Color(theme.bg_header)),
                        ..Default::default()
                    })
                    .width(Length::Fill)
                    .into()
            )
        } else {
            None
        };

        let menu_bar = self.view_menu_bar();
        let format_bar = self.view_format_bar();

        // Build grid(s) - single or split view
        let grid: Element<'_, Message> = if self.split_enabled {
            // Create two grids with independent scroll positions
            let is_primary_active = self.split_active_pane == SplitPane::Primary;
            let is_secondary_active = self.split_active_pane == SplitPane::Secondary;

            // Primary pane (left or top)
            let primary_grid = self.view_grid(Some(SplitPane::Primary));
            let primary_pane = mouse_area(
                container(primary_grid)
                    .width(Length::Fill)
                    .height(Length::Fill)
                    .style(move |_| container::Style {
                        border: Border {
                            color: if is_primary_active { theme.selected_border } else { theme.gridline },
                            width: if is_primary_active { 2.0 } else { 1.0 },
                            radius: 0.0.into(),
                        },
                        ..Default::default()
                    })
            )
            .on_press(Message::SplitPaneClicked(SplitPane::Primary));

            // Secondary pane (right or bottom)
            let secondary_grid = self.view_grid(Some(SplitPane::Secondary));
            let secondary_pane = mouse_area(
                container(secondary_grid)
                    .width(Length::Fill)
                    .height(Length::Fill)
                    .style(move |_| container::Style {
                        border: Border {
                            color: if is_secondary_active { theme.selected_border } else { theme.gridline },
                            width: if is_secondary_active { 2.0 } else { 1.0 },
                            radius: 0.0.into(),
                        },
                        ..Default::default()
                    })
            )
            .on_press(Message::SplitPaneClicked(SplitPane::Secondary));

            match self.split_direction {
                SplitDirection::Horizontal => {
                    // Side by side (left/right)
                    Row::with_children(vec![primary_pane.into(), secondary_pane.into()])
                        .spacing(2)
                        .width(Length::Fill)
                        .height(Length::Fill)
                        .into()
                }
                SplitDirection::Vertical => {
                    // Stacked (top/bottom)
                    Column::with_children(vec![primary_pane.into(), secondary_pane.into()])
                        .spacing(2)
                        .width(Length::Fill)
                        .height(Length::Fill)
                        .into()
                }
            }
        } else {
            self.view_grid(None)
        };

        // Wrap grid with inspector panel if enabled (not in zen mode)
        let grid_area: Element<'_, Message> = if self.show_inspector && !self.zen_mode {
            let inspector = self.view_inspector_panel();
            Row::with_children(vec![grid, inspector])
                .spacing(0)
                .width(Length::Fill)
                .height(Length::Fill)
                .into()
        } else {
            grid
        };

        // Sheet tabs at bottom
        let sheet_tabs = self.view_sheet_tabs();

        // Problems panel (optional)
        let problems_panel: Option<Element<'_, Message>> = if self.show_problems && !self.zen_mode {
            Some(self.view_problems_panel())
        } else {
            None
        };

        // Build main content - in zen mode, only show the grid
        let main_content = if self.zen_mode {
            // Zen mode: just the grid, centered
            column![grid_area].spacing(0)
        } else {
            match (formula_preview, problems_panel) {
                (Some(preview), Some(problems)) => {
                    column![menu_bar, formula_bar, preview, format_bar, grid_area, problems, sheet_tabs].spacing(0)
                }
                (Some(preview), None) => {
                    column![menu_bar, formula_bar, preview, format_bar, grid_area, sheet_tabs].spacing(0)
                }
                (None, Some(problems)) => {
                    column![menu_bar, formula_bar, format_bar, grid_area, problems, sheet_tabs].spacing(0)
                }
                (None, None) => {
                    column![menu_bar, formula_bar, format_bar, grid_area, sheet_tabs].spacing(0)
                }
            }
        };

        let styled_content = container(main_content)
            .width(Length::Fill)
            .height(Length::Fill)
            .style(move |_| container::Style {
                background: Some(Background::Color(theme.bg_dark)),
                ..Default::default()
            });

        // Build overlays
        let has_menu_open = self.open_menu.is_some();
        let has_palette_open = self.mode == Mode::Command;
        let has_goto_open = self.mode == Mode::GoTo;
        let has_quick_open = self.mode == Mode::QuickOpen;
        let has_find_open = self.mode == Mode::Find;

        if has_palette_open {
            let palette = self.view_palette();
            iced::widget::stack![
                styled_content,
                container(
                    container(palette)
                        .width(500)
                        .style(move |_| container::Style {
                            background: Some(Background::Color(theme.bg_cell)),
                            border: Border {
                                color: theme.border,
                                width: 1.0,
                                radius: 8.0.into(),
                            },
                            ..Default::default()
                        })
                )
                .width(Length::Fill)
                .padding(iced::Padding { top: 50.0, right: 0.0, bottom: 0.0, left: 0.0 })
                .center_x(Length::Fill)
            ]
            .into()
        } else if has_goto_open {
            let goto = self.view_goto();
            iced::widget::stack![
                styled_content,
                container(
                    container(goto)
                        .width(350)
                        .style(move |_| container::Style {
                            background: Some(Background::Color(theme.bg_header)),
                            border: Border {
                                color: theme.border,
                                width: 1.0,
                                radius: 6.0.into(),
                            },
                            ..Default::default()
                        })
                )
                .width(Length::Fill)
                .padding(iced::Padding { top: 50.0, right: 0.0, bottom: 0.0, left: 0.0 })
                .center_x(Length::Fill)
            ]
            .into()
        } else if has_quick_open {
            let quick_open = self.view_quick_open();
            iced::widget::stack![
                styled_content,
                container(
                    container(quick_open)
                        .width(500)
                        .height(400)
                        .style(move |_| container::Style {
                            background: Some(Background::Color(theme.bg_cell)),
                            border: Border {
                                color: theme.border,
                                width: 1.0,
                                radius: 8.0.into(),
                            },
                            ..Default::default()
                        })
                )
                .width(Length::Fill)
                .padding(iced::Padding { top: 50.0, right: 0.0, bottom: 0.0, left: 0.0 })
                .center_x(Length::Fill)
            ]
            .into()
        } else if has_find_open {
            let find = self.view_find();
            iced::widget::stack![
                styled_content,
                container(
                    container(find)
                        .width(450)
                        .height(350)
                        .style(move |_| container::Style {
                            background: Some(Background::Color(theme.bg_cell)),
                            border: Border {
                                color: theme.border,
                                width: 1.0,
                                radius: 8.0.into(),
                            },
                            ..Default::default()
                        })
                )
                .width(Length::Fill)
                .padding(iced::Padding { top: 50.0, right: 0.0, bottom: 0.0, left: 0.0 })
                .center_x(Length::Fill)
            ]
            .into()
        } else if has_menu_open {
            let dropdown = self.view_open_menu_dropdown();
            iced::widget::stack![
                styled_content,
                dropdown
            ]
            .into()
        } else if self.autocomplete_visible && !self.autocomplete_suggestions.is_empty() {
            // Show autocomplete dropdown below formula bar
            let autocomplete = self.view_autocomplete();
            iced::widget::stack![
                styled_content,
                container(autocomplete)
                    .width(Length::Fill)
                    .padding(iced::Padding { top: 70.0, right: 0.0, bottom: 0.0, left: 80.0 })
            ]
            .into()
        } else {
            styled_content.into()
        }
    }

    fn view_open_menu_dropdown(&self) -> Element<'_, Message> {
        let menus: &[(&'static str, &[MenuItem])] = &[
            ("File", FILE_MENU),
            ("Edit", EDIT_MENU),
            ("View", VIEW_MENU),
            ("Insert", INSERT_MENU),
            ("Format", FORMAT_MENU),
            ("Data", DATA_MENU),
            ("Help", HELP_MENU),
        ];

        if let Some(open_label) = self.open_menu {
            let menu_items = menus.iter()
                .find(|(label, _)| *label == open_label)
                .map(|(_, items)| *items)
                .unwrap_or(&[]);

            // Calculate horizontal offset based on which menu is open
            let offset: f32 = menus.iter()
                .take_while(|(label, _)| *label != open_label)
                .map(|(label, _)| label.len() as f32 * 7.5 + 20.0)
                .sum::<f32>() + 8.0;

            self.view_dropdown_panel(menu_items, offset)
        } else {
            // Empty container if no menu open
            container(text("")).into()
        }
    }

    fn view_dropdown_panel(&self, items: &[MenuItem], offset: f32) -> Element<'_, Message> {
        let theme = self.theme();

        let menu_items: Vec<Element<'_, Message>> = items
            .iter()
            .flat_map(|item| {
                let separator = item.separator_after;
                let mut elements: Vec<Element<'_, Message>> = vec![
                    button(
                        row![
                            text(item.label).size(13).color(theme.text).width(Length::Fill),
                            text(item.shortcut.unwrap_or("")).size(12).color(theme.text_dim),
                        ]
                        .spacing(20)
                    )
                    .width(Length::Fill)
                    .padding(iced::Padding { top: 5.0, right: 14.0, bottom: 5.0, left: 14.0 })
                    .style(move |_, status| {
                        let bg = match status {
                            button::Status::Hovered | button::Status::Pressed => theme.selected,
                            _ => theme.bg_input,
                        };
                        button::Style {
                            background: Some(Background::Color(bg)),
                            text_color: theme.text,
                            border: Border::default(),
                            shadow: Default::default(),
                        }
                    })
                    .on_press(Message::MenuItemClicked(item.id))
                    .into()
                ];

                if separator {
                    elements.push(
                        container(text(""))
                            .width(Length::Fill)
                            .height(1)
                            .style(move |_| container::Style {
                                background: Some(Background::Color(theme.gridline)),
                                ..Default::default()
                            })
                            .into()
                    );
                }

                elements
            })
            .collect();

        container(
            container(
                Column::with_children(menu_items).spacing(0)
            )
            .width(220)
            .padding(4)
            .style(move |_| container::Style {
                background: Some(Background::Color(theme.bg_input)),
                border: Border {
                    color: theme.border,
                    width: 1.0,
                    radius: 4.0.into(),
                },
                ..Default::default()
            })
        )
        .width(Length::Fill)
        .height(Length::Fill)
        .padding(iced::Padding { top: 26.0, right: 0.0, bottom: 0.0, left: offset })
        .into()
    }

    fn view_menu_bar(&self) -> Element<'_, Message> {
        let theme = self.theme();
        let menu_labels = ["File", "Edit", "View", "Insert", "Format", "Data", "Help"];

        let items: Vec<Element<'_, Message>> = menu_labels
            .iter()
            .map(|&label| {
                let is_open = self.open_menu == Some(label);
                button(
                    text(label).size(13).color(theme.text)
                )
                .padding(iced::Padding { top: 6.0, right: 12.0, bottom: 6.0, left: 12.0 })
                .style(move |_, status| {
                    let bg = if is_open {
                        theme.bg_input
                    } else {
                        match status {
                            button::Status::Hovered | button::Status::Pressed => theme.bg_input,
                            _ => theme.bg_header,
                        }
                    };
                    button::Style {
                        background: Some(Background::Color(bg)),
                        text_color: theme.text,
                        border: Border::default(),
                        shadow: Default::default(),
                    }
                })
                .on_press(Message::MenuClicked(label))
                .into()
            })
            .collect();

        container(
            Row::with_children(items).spacing(0)
        )
        .width(Length::Fill)
        .padding(iced::Padding { top: 2.0, right: 8.0, bottom: 2.0, left: 4.0 })
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.gridline,
                width: 0.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_format_bar(&self) -> Element<'_, Message> {
        let theme = self.theme();

        // Get formatting state of active cell for button highlighting
        let (r, c) = self.selection.active_cell();
        let format = self.sheet.get_format(r, c);

        let bold_btn = button(
            text("B").size(12).color(theme.text).font(Font {
                weight: Weight::Bold,
                ..Font::DEFAULT
            }).center()
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatBold)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.bold { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let italic_btn = button(
            text("I").size(12).color(theme.text).font(Font {
                style: FontStyle::Italic,
                ..Font::DEFAULT
            }).center()
        )
        .padding(iced::Padding { top: 4.0, right: 10.0, bottom: 4.0, left: 10.0 })
        .on_press(Message::FormatItalic)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.italic { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let underline_btn = button(
            text("U").size(12).color(theme.text).center()
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatUnderline)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.underline { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        // Separator helper
        let separator = || {
            container(text("").size(1))
                .width(Length::Fixed(1.0))
                .height(Length::Fixed(20.0))
                .style(move |_| container::Style {
                    background: Some(Background::Color(theme.gridline)),
                    ..Default::default()
                })
        };

        // Alignment buttons
        let align_left_btn = button(
            text("≡").size(12).color(theme.text).center()  // Left align icon
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatAlignLeft)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.alignment == Alignment::Left { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let align_center_btn = button(
            text("☰").size(12).color(theme.text).center()  // Center align icon
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatAlignCenter)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.alignment == Alignment::Center { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let align_right_btn = button(
            text("⁝").size(12).color(theme.text).center()  // Right align icon (3 dots)
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatAlignRight)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if format.alignment == Alignment::Right { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        // Number format buttons
        let is_currency = matches!(format.number_format, NumberFormat::Currency { .. });
        let is_percent = matches!(format.number_format, NumberFormat::Percent { .. });

        let currency_btn = button(
            text("$").size(12).color(theme.text).center()
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatCurrency)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if is_currency { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let percent_btn = button(
            text("%").size(12).color(theme.text).center()
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .on_press(Message::FormatPercent)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(if is_percent { theme.selected } else { theme.bg_header })),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let decimal_inc_btn = button(
            text(".0→").size(10).color(theme.text).center()
        )
        .padding(iced::Padding { top: 4.0, right: 6.0, bottom: 4.0, left: 6.0 })
        .on_press(Message::FormatDecimalIncrease)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(theme.bg_header)),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        let decimal_dec_btn = button(
            text("←.0").size(10).color(theme.text).center()
        )
        .padding(iced::Padding { top: 4.0, right: 6.0, bottom: 4.0, left: 6.0 })
        .on_press(Message::FormatDecimalDecrease)
        .style(move |_, _| button::Style {
            background: Some(Background::Color(theme.bg_header)),
            text_color: theme.text,
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 2.0.into(),
            },
            shadow: Default::default(),
        });

        container(
            row![
                bold_btn, italic_btn, underline_btn,
                separator(),
                align_left_btn, align_center_btn, align_right_btn,
                separator(),
                currency_btn, percent_btn,
                separator(),
                decimal_dec_btn, decimal_inc_btn
            ].spacing(2).align_y(iced::Alignment::Center)
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 4.0, left: 8.0 })
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.gridline,
                width: 0.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_problems_panel(&self) -> Element<'_, Message> {
        let theme = self.theme();
        let error_count = self.problems_cache.len();

        // Header with count
        let header = container(
            row![
                text(format!("Problems ({})", error_count))
                    .size(12)
                    .color(theme.text),
                iced::widget::horizontal_space(),
                button(text("×").size(14).color(theme.text_dim).center())
                    .padding(iced::Padding { top: 2.0, right: 6.0, bottom: 2.0, left: 6.0 })
                    .on_press(Message::ToggleProblems)
                    .style(move |_, _| button::Style {
                        background: None,
                        text_color: theme.text_dim,
                        border: Border::default(),
                        shadow: Default::default(),
                    })
            ]
            .align_y(iced::Alignment::Center)
        )
        .padding(iced::Padding { top: 6.0, right: 8.0, bottom: 6.0, left: 12.0 })
        .width(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.border,
                width: 0.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        });

        // Error list
        let error_list: Element<'_, Message> = if self.problems_cache.is_empty() {
            container(
                text("No problems detected").size(12).color(theme.text_dim)
            )
            .padding(12)
            .width(Length::Fill)
            .into()
        } else {
            let mut items = Column::new().spacing(0);
            for error in &self.problems_cache {
                let cell_ref = format!("{}{}", Self::col_to_letters(error.col), error.row + 1);
                let error_type = error.error_type.clone();
                let description = error.description.clone();
                let row = error.row;
                let col = error.col;

                let error_row = button(
                    row![
                        container(
                            text(error_type).size(11).color(Color::from_rgb(0.9, 0.3, 0.3))
                        )
                        .width(60),
                        container(
                            text(cell_ref).size(11).color(theme.accent)
                        )
                        .width(50),
                        text(description).size(11).color(theme.text_dim)
                    ]
                    .spacing(8)
                    .align_y(iced::Alignment::Center)
                )
                .padding(iced::Padding { top: 4.0, right: 12.0, bottom: 4.0, left: 12.0 })
                .width(Length::Fill)
                .on_press(Message::GoToProblem(row, col))
                .style(move |_, status| {
                    let bg = match status {
                        button::Status::Hovered => theme.selected,
                        _ => theme.bg_cell,
                    };
                    button::Style {
                        background: Some(Background::Color(bg)),
                        text_color: theme.text,
                        border: Border::default(),
                        shadow: Default::default(),
                    }
                });

                items = items.push(error_row);
            }

            iced::widget::scrollable(items)
                .height(Length::Fill)
                .into()
        };

        container(
            column![header, error_list].spacing(0)
        )
        .width(Length::Fill)
        .height(150)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_cell)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_autocomplete(&self) -> Element<'_, Message> {
        let theme = self.theme();

        let mut items: Vec<Element<'_, Message>> = Vec::new();

        for (i, &func_idx) in self.autocomplete_suggestions.iter().enumerate() {
            let func = &FORMULA_FUNCTIONS[func_idx];
            let is_selected = i == self.autocomplete_selected;

            let item = button(
                row![
                    container(
                        text(func.name).size(12).color(theme.text)
                    ).width(80),
                    text(func.description).size(11).color(theme.text_dim)
                ]
                .spacing(8)
                .align_y(iced::Alignment::Center)
            )
            .padding(iced::Padding { top: 6.0, right: 12.0, bottom: 6.0, left: 12.0 })
            .width(Length::Fill)
            .on_press(Message::AutocompleteSelect(i))
            .style(move |_, status| {
                let bg = if is_selected {
                    theme.selected
                } else {
                    match status {
                        button::Status::Hovered => theme.selected,
                        _ => theme.bg_cell,
                    }
                };
                button::Style {
                    background: Some(Background::Color(bg)),
                    text_color: theme.text,
                    border: Border::default(),
                    shadow: Default::default(),
                }
            });

            items.push(item.into());
        }

        container(
            Column::with_children(items).spacing(0)
        )
        .width(400)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_cell)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 4.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_inspector_panel(&self) -> Element<'_, Message> {
        use crate::engine::cell::CellValue;

        let theme = self.theme();
        let (row, col) = self.selection.active_cell();
        let cell_ref = Self::cell_ref_string(row, col);
        let cell = self.sheet.get_cell(row, col);

        // Header
        let header = container(
            row![
                text(format!("Inspector: {}", cell_ref))
                    .size(12)
                    .color(theme.text),
                iced::widget::horizontal_space(),
                button(text("×").size(14).color(theme.text_dim).center())
                    .padding(iced::Padding { top: 2.0, right: 6.0, bottom: 2.0, left: 6.0 })
                    .on_press(Message::InspectorToggle)
                    .style(move |_, _| button::Style {
                        background: None,
                        text_color: theme.text_dim,
                        border: Border::default(),
                        shadow: Default::default(),
                    })
            ]
            .align_y(iced::Alignment::Center)
        )
        .padding(iced::Padding { top: 6.0, right: 8.0, bottom: 6.0, left: 12.0 })
        .width(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border::default(),
            ..Default::default()
        });

        // Cell info
        let (formula_text, result_text, cell_type) = match &cell.value {
            CellValue::Empty => (String::new(), String::new(), "Empty"),
            CellValue::Text(s) => (String::new(), s.clone(), "Text"),
            CellValue::Number(n) => (String::new(), format!("{}", n), "Number"),
            CellValue::Formula { source, .. } => {
                let result = self.sheet.get_display(row, col);
                (source.clone(), result, "Formula")
            }
        };

        let mut info_rows: Vec<Element<'_, Message>> = Vec::new();

        // Formula row (if applicable)
        if !formula_text.is_empty() {
            info_rows.push(
                row![
                    container(text("Formula:").size(11).color(theme.text_dim)).width(70),
                    text(formula_text.clone()).size(11).color(theme.text)
                ]
                .spacing(8)
                .into()
            );
        }

        // Result row
        info_rows.push(
            row![
                container(text("Value:").size(11).color(theme.text_dim)).width(70),
                text(result_text.clone()).size(11).color(theme.text)
            ]
            .spacing(8)
            .into()
        );

        // Type row
        info_rows.push(
            row![
                container(text("Type:").size(11).color(theme.text_dim)).width(70),
                text(cell_type).size(11).color(theme.text)
            ]
            .spacing(8)
            .into()
        );

        // Format row
        let mut format_parts = Vec::new();
        if cell.format.bold { format_parts.push("Bold"); }
        if cell.format.italic { format_parts.push("Italic"); }
        if cell.format.underline { format_parts.push("Underline"); }
        let format_str = if format_parts.is_empty() {
            "Default".to_string()
        } else {
            format_parts.join(", ")
        };
        info_rows.push(
            row![
                container(text("Format:").size(11).color(theme.text_dim)).width(70),
                text(format_str).size(11).color(theme.text)
            ]
            .spacing(8)
            .into()
        );

        let info_section = Column::with_children(info_rows)
            .spacing(4)
            .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 8.0, left: 12.0 });

        // Precedents section
        let precedents = self.get_precedents(row, col);
        let prec_header = text("Precedents (depends on):").size(11).color(theme.text_dim);
        let prec_content: Element<'_, Message> = if precedents.is_empty() {
            text("None").size(11).color(theme.text_dim).into()
        } else {
            let refs: Vec<Element<'_, Message>> = precedents.iter().map(|(r, c)| {
                let ref_str = Self::cell_ref_string(*r, *c);
                let r = *r;
                let c = *c;
                button(text(ref_str).size(11).color(theme.accent))
                    .padding(iced::Padding { top: 2.0, right: 6.0, bottom: 2.0, left: 6.0 })
                    .on_press(Message::GoToProblem(r, c))  // Reuse this message to navigate
                    .style(move |_, status| {
                        let bg = match status {
                            button::Status::Hovered => theme.selected,
                            _ => Color::TRANSPARENT,
                        };
                        button::Style {
                            background: Some(Background::Color(bg)),
                            text_color: theme.accent,
                            border: Border::default(),
                            shadow: Default::default(),
                        }
                    })
                    .into()
            }).collect();
            Row::with_children(refs).spacing(4).wrap().into()
        };
        let prec_section = column![prec_header, prec_content]
            .spacing(4)
            .padding(iced::Padding { top: 4.0, right: 12.0, bottom: 8.0, left: 12.0 });

        // Dependents section
        let dependents = self.get_dependents(row, col);
        let dep_header = text("Dependents (used by):").size(11).color(theme.text_dim);
        let dep_content: Element<'_, Message> = if dependents.is_empty() {
            text("None").size(11).color(theme.text_dim).into()
        } else {
            let refs: Vec<Element<'_, Message>> = dependents.iter().map(|(r, c)| {
                let ref_str = Self::cell_ref_string(*r, *c);
                let r = *r;
                let c = *c;
                button(text(ref_str).size(11).color(theme.accent))
                    .padding(iced::Padding { top: 2.0, right: 6.0, bottom: 2.0, left: 6.0 })
                    .on_press(Message::GoToProblem(r, c))
                    .style(move |_, status| {
                        let bg = match status {
                            button::Status::Hovered => theme.selected,
                            _ => Color::TRANSPARENT,
                        };
                        button::Style {
                            background: Some(Background::Color(bg)),
                            text_color: theme.accent,
                            border: Border::default(),
                            shadow: Default::default(),
                        }
                    })
                    .into()
            }).collect();
            Row::with_children(refs).spacing(4).wrap().into()
        };
        let dep_section = column![dep_header, dep_content]
            .spacing(4)
            .padding(iced::Padding { top: 4.0, right: 12.0, bottom: 8.0, left: 12.0 });

        // Divider helper
        let divider = || container(text(""))
            .width(Length::Fill)
            .height(1)
            .style(move |_| container::Style {
                background: Some(Background::Color(theme.border)),
                ..Default::default()
            });

        container(
            column![
                header,
                info_section,
                divider(),
                prec_section,
                divider(),
                dep_section
            ].spacing(0)
        )
        .width(250)
        .height(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_cell)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_sheet_tabs(&self) -> Element<'_, Message> {
        let theme = self.theme();
        let tab = container(
            text("Sheet1").size(12).color(theme.text)
        )
        .padding(iced::Padding { top: 6.0, right: 16.0, bottom: 6.0, left: 16.0 })
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_cell)),
            border: Border {
                color: theme.border,
                width: 1.0,
                radius: iced::border::Radius {
                    top_left: 4.0,
                    top_right: 4.0,
                    bottom_left: 0.0,
                    bottom_right: 0.0,
                },
            },
            ..Default::default()
        });

        let add_tab = container(
            text("+").size(14).color(theme.text_dim)
        )
        .padding(iced::Padding { top: 4.0, right: 10.0, bottom: 4.0, left: 10.0 });

        // Status bar text (filename or status message)
        let status_text = if let Some(ref msg) = self.status_message {
            msg.clone()
        } else if let Some(ref path) = self.current_file {
            path.file_name()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_default()
        } else {
            String::new()
        };

        let status = container(
            text(status_text).size(11).color(theme.text_dim)
        )
        .padding(iced::Padding { top: 6.0, right: 8.0, bottom: 6.0, left: 8.0 });

        container(
            row![
                row![tab, add_tab].spacing(4).align_y(iced::Alignment::End),
                iced::widget::horizontal_space(),
                status
            ]
        )
        .padding(iced::Padding { top: 4.0, right: 8.0, bottom: 0.0, left: 8.0 })
        .width(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_header)),
            border: Border {
                color: theme.gridline,
                width: 1.0,
                radius: 0.0.into(),
            },
            ..Default::default()
        })
        .into()
    }

    fn view_palette(&self) -> Element<'_, Message> {
        let theme = self.theme();

        // Placeholder text based on mode
        let placeholder = match self.palette_mode {
            PaletteMode::Commands => "> Command name...",
            PaletteMode::Cells => "@ Search in cells...",
            PaletteMode::GoTo => ": Cell reference (e.g., A1)",
            PaletteMode::Functions => "= Function name (e.g., SUM, AVERAGE)...",
            PaletteMode::Mixed => "Type > commands, @ cells, : goto, = functions, or search...",
        };

        // Search input
        let search_input = text_input(placeholder, &self.palette_query)
            .on_input(Message::PaletteQueryChanged)
            .on_submit(Message::PaletteSubmit)
            .size(14)
            .padding(12);

        let search_row = container(search_input)
            .width(Length::Fill)
            .padding(iced::Padding { top: 12.0, right: 12.0, bottom: 8.0, left: 12.0 });

        // Build items based on mode
        let mut items: Vec<Element<'_, Message>> = Vec::new();
        let mut total_items = 0;

        match self.palette_mode {
            PaletteMode::Commands => {
                // Show commands only
                for (display_idx, &cmd_idx) in self.palette_filtered.iter().take(10).enumerate() {
                    let cmd = &COMMANDS[cmd_idx];
                    let is_selected = display_idx == self.palette_selected;
                    items.push(self.view_palette_command_item(cmd, is_selected, theme));
                    total_items += 1;
                }
            }
            PaletteMode::Cells => {
                // Show cell search results
                for (display_idx, (row, col, display)) in self.palette_cells.iter().take(10).enumerate() {
                    let is_selected = display_idx == self.palette_selected;
                    let r = *row;
                    let c = *col;
                    items.push(self.view_palette_cell_item(display.clone(), r, c, is_selected, theme));
                    total_items += 1;
                }
                if self.palette_cells.is_empty() && !self.palette_query.is_empty() {
                    items.push(
                        container(text("No cells found").size(12).color(theme.text_dim))
                            .padding(12)
                            .width(Length::Fill)
                            .center_x(Length::Fill)
                            .into()
                    );
                }
            }
            PaletteMode::GoTo => {
                // Show hint for goto
                let cell_ref = self.palette_query.get(1..).unwrap_or("").trim();
                let hint_text = if cell_ref.is_empty() {
                    "Enter a cell reference (e.g., A1, B25, AA100)".to_string()
                } else {
                    format!("Press Enter to go to {}", cell_ref.to_uppercase())
                };
                items.push(
                    container(text(hint_text).size(12).color(theme.text_dim))
                        .padding(12)
                        .width(Length::Fill)
                        .center_x(Length::Fill)
                        .into()
                );
            }
            PaletteMode::Functions => {
                // Show formula functions
                for (display_idx, &func_idx) in self.palette_functions.iter().take(10).enumerate() {
                    let func = &FORMULA_FUNCTIONS[func_idx];
                    let is_selected = display_idx == self.palette_selected;
                    items.push(self.view_palette_function_item(func, is_selected, theme));
                    total_items += 1;
                }
                if self.palette_functions.is_empty() && self.palette_query.len() > 1 {
                    items.push(
                        container(text("No functions found").size(12).color(theme.text_dim))
                            .padding(12)
                            .width(Length::Fill)
                            .center_x(Length::Fill)
                            .into()
                    );
                }
            }
            PaletteMode::Mixed => {
                // Show commands first (limited)
                let cmd_limit = if self.palette_files.is_empty() { 8 } else { 5 };
                for (display_idx, &cmd_idx) in self.palette_filtered.iter().take(cmd_limit).enumerate() {
                    let cmd = &COMMANDS[cmd_idx];
                    let is_selected = display_idx == self.palette_selected;
                    items.push(self.view_palette_command_item(cmd, is_selected, theme));
                    total_items += 1;
                }

                // Show recent files section if we have any
                if !self.palette_files.is_empty() {
                    // Section header
                    items.push(
                        container(text("Recent Files").size(10).color(theme.text_dim))
                            .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 4.0, left: 12.0 })
                            .width(Length::Fill)
                            .into()
                    );

                    for (file_display_idx, &file_idx) in self.palette_files.iter().take(5).enumerate() {
                        let is_selected = (total_items + file_display_idx) == self.palette_selected;
                        if let Some(path) = self.recent_files.get(file_idx) {
                            items.push(self.view_palette_file_item(path, file_idx, is_selected, theme));
                        }
                    }
                }
            }
        }

        let results = Column::with_children(items).spacing(0);
        let results_container = container(results)
            .width(Length::Fill)
            .padding(iced::Padding { top: 4.0, right: 4.0, bottom: 4.0, left: 4.0 });

        // Separator line
        let separator = container(text(""))
            .width(Length::Fill)
            .height(Length::Fixed(1.0))
            .style(move |_| container::Style {
                background: Some(Background::Color(theme.border)),
                ..Default::default()
            });

        // Footer with keyboard hints
        let footer_text = match self.palette_mode {
            PaletteMode::GoTo => "↵ go to cell  ·  esc close",
            _ => "↑↓ navigate  ·  ↵ select  ·  esc close",
        };
        let footer = container(
            text(footer_text).size(11).color(theme.text_dim)
        )
        .width(Length::Fill)
        .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 10.0, left: 12.0 });

        column![
            search_row,
            results_container,
            separator,
            footer,
        ]
        .spacing(0)
        .into()
    }

    fn view_palette_command_item(&self, cmd: &Command, is_selected: bool, theme: ThemeColors) -> Element<'_, Message> {
        let shortcut_text = cmd.shortcut.unwrap_or("");
        let item_content = row![
            text(cmd.label).size(13).color(theme.text),
            iced::widget::horizontal_space(),
            text(shortcut_text).size(11).color(theme.text_dim)
        ]
        .align_y(iced::Alignment::Center)
        .padding(iced::Padding { top: 0.0, right: 12.0, bottom: 0.0, left: 12.0 });

        let bg_color = if is_selected { theme.selected } else { Color::TRANSPARENT };

        button(item_content)
            .width(Length::Fill)
            .padding(iced::Padding { top: 8.0, right: 0.0, bottom: 8.0, left: 0.0 })
            .on_press(Message::ExecuteCommand(cmd.id))
            .style(move |_theme, status| {
                let hover_bg = if is_selected {
                    theme.selected
                } else {
                    Color::from_rgba(1.0, 1.0, 1.0, 0.05)
                };
                button::Style {
                    background: Some(Background::Color(match status {
                        button::Status::Hovered | button::Status::Pressed => hover_bg,
                        _ => bg_color,
                    })),
                    text_color: theme.text,
                    border: Border::default(),
                    ..Default::default()
                }
            })
            .into()
    }

    fn view_palette_cell_item(&self, display: String, row: usize, col: usize, is_selected: bool, theme: ThemeColors) -> Element<'_, Message> {
        let bg_color = if is_selected { theme.selected } else { Color::TRANSPARENT };

        button(
            text(display).size(12).color(theme.text)
        )
        .width(Length::Fill)
        .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 8.0, left: 12.0 })
        .on_press(Message::FindSelect(row, col))
        .style(move |_theme, status| {
            let hover_bg = if is_selected {
                theme.selected
            } else {
                Color::from_rgba(1.0, 1.0, 1.0, 0.05)
            };
            button::Style {
                background: Some(Background::Color(match status {
                    button::Status::Hovered | button::Status::Pressed => hover_bg,
                    _ => bg_color,
                })),
                text_color: theme.text,
                border: Border::default(),
                ..Default::default()
            }
        })
        .into()
    }

    fn view_palette_function_item(&self, func: &FormulaFunction, is_selected: bool, theme: ThemeColors) -> Element<'_, Message> {
        let bg_color = if is_selected { theme.selected } else { Color::TRANSPARENT };

        let content = column![
            row![
                text(func.name).size(13).color(theme.text),
                iced::widget::horizontal_space(),
                text(func.category).size(10).color(theme.text_dim)
            ].align_y(iced::Alignment::Center),
            text(func.syntax).size(11).color(theme.accent),
            text(func.description).size(11).color(theme.text_dim)
        ].spacing(2);

        container(content)
            .width(Length::Fill)
            .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 8.0, left: 12.0 })
            .style(move |_| container::Style {
                background: Some(Background::Color(bg_color)),
                ..Default::default()
            })
            .into()
    }

    fn view_palette_file_item(&self, path: &PathBuf, idx: usize, is_selected: bool, theme: ThemeColors) -> Element<'_, Message> {
        let file_name = path.file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("Unknown")
            .to_string();
        let dir_path = path.parent()
            .and_then(|p| p.to_str())
            .unwrap_or("")
            .to_string();

        let bg_color = if is_selected { theme.selected } else { Color::TRANSPARENT };

        button(
            column![
                text(file_name).size(12).color(theme.text),
                text(dir_path).size(10).color(theme.text_dim)
            ].spacing(1)
        )
        .width(Length::Fill)
        .padding(iced::Padding { top: 6.0, right: 12.0, bottom: 6.0, left: 12.0 })
        .on_press(Message::QuickOpenSelect(idx))
        .style(move |_theme, status| {
            let hover_bg = if is_selected {
                theme.selected
            } else {
                Color::from_rgba(1.0, 1.0, 1.0, 0.05)
            };
            button::Style {
                background: Some(Background::Color(match status {
                    button::Status::Hovered | button::Status::Pressed => hover_bg,
                    _ => bg_color,
                })),
                text_color: theme.text,
                border: Border::default(),
                ..Default::default()
            }
        })
        .into()
    }

    fn view_goto(&self) -> Element<'_, Message> {
        let theme = self.theme();
        let goto_input = text_input("Cell reference (e.g., A1, B25)", &self.goto_input)
            .on_input(Message::GoToInputChanged)
            .on_submit(Message::GoToSubmit)
            .size(16)
            .padding(12);

        let hint = text("Enter a cell reference and press Enter")
            .size(12)
            .color(theme.text_dim);

        column![
            text("Go to Cell").size(14).color(theme.text),
            goto_input,
            container(hint).padding(8),
        ]
        .spacing(8)
        .padding(8)
        .into()
    }

    fn view_quick_open(&self) -> Element<'_, Message> {
        let theme = self.theme();

        // Search input
        let search_input = text_input("Search recent files...", &self.quick_open_query)
            .on_input(Message::QuickOpenQueryChanged)
            .on_submit(Message::QuickOpenSubmit)
            .size(14)
            .padding(10)
            .style(move |_, _| text_input::Style {
                background: Background::Color(theme.bg_input),
                border: Border {
                    color: theme.border,
                    width: 1.0,
                    radius: 4.0.into(),
                },
                icon: theme.text_dim,
                placeholder: theme.text_dim,
                value: theme.text,
                selection: theme.selected,
            });

        // File list
        let file_list: Element<'_, Message> = if self.recent_files.is_empty() {
            container(
                text("No recent files").size(12).color(theme.text_dim)
            )
            .padding(16)
            .width(Length::Fill)
            .center_x(Length::Fill)
            .into()
        } else if self.quick_open_filtered.is_empty() {
            container(
                text("No matching files").size(12).color(theme.text_dim)
            )
            .padding(16)
            .width(Length::Fill)
            .center_x(Length::Fill)
            .into()
        } else {
            let mut items = Column::new().spacing(0);
            for (list_idx, &file_idx) in self.quick_open_filtered.iter().enumerate() {
                let path = &self.recent_files[file_idx];
                let file_name = path.file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or("Unknown");
                let dir_path = path.parent()
                    .and_then(|p| p.to_str())
                    .unwrap_or("");

                let is_selected = list_idx == self.quick_open_selected;
                let idx = file_idx;

                let file_row = button(
                    column![
                        text(file_name).size(13).color(theme.text),
                        text(dir_path).size(10).color(theme.text_dim)
                    ]
                    .spacing(2)
                )
                .padding(iced::Padding { top: 8.0, right: 12.0, bottom: 8.0, left: 12.0 })
                .width(Length::Fill)
                .on_press(Message::QuickOpenSelect(idx))
                .style(move |_, _| {
                    button::Style {
                        background: Some(Background::Color(if is_selected { theme.selected } else { theme.bg_cell })),
                        text_color: theme.text,
                        border: Border::default(),
                        shadow: Default::default(),
                    }
                });

                items = items.push(file_row);
            }

            iced::widget::scrollable(items)
                .height(Length::FillPortion(1))
                .into()
        };

        let hint = text("↑↓ Navigate  Enter Open  Esc Cancel")
            .size(10)
            .color(theme.text_dim);

        column![
            search_input,
            file_list,
            container(hint).padding(8).center_x(Length::Fill)
        ]
        .spacing(4)
        .padding(8)
        .into()
    }

    fn view_find(&self) -> Element<'_, Message> {
        let theme = self.theme();

        // Search input
        let search_input = text_input("Search in cells...", &self.find_query)
            .on_input(Message::FindQueryChanged)
            .on_submit(Message::FindSubmit)
            .size(14)
            .padding(10)
            .style(move |_, _| text_input::Style {
                background: Background::Color(theme.bg_input),
                border: Border {
                    color: theme.border,
                    width: 1.0,
                    radius: 4.0.into(),
                },
                icon: theme.text_dim,
                placeholder: theme.text_dim,
                value: theme.text,
                selection: theme.selected,
            });

        // Results count
        let results_text = if self.find_query.is_empty() {
            "Type to search".to_string()
        } else if self.find_results.is_empty() {
            "No matches found".to_string()
        } else {
            format!("{} matches found", self.find_results.len())
        };

        let results_label = text(results_text)
            .size(11)
            .color(theme.text_dim);

        // Results list
        let results_list: Element<'_, Message> = if self.find_results.is_empty() {
            container(
                text("").size(12)
            )
            .height(Length::FillPortion(1))
            .into()
        } else {
            let mut items = Column::new().spacing(0);
            for (idx, (row, col, display)) in self.find_results.iter().enumerate() {
                let is_selected = idx == self.find_selected;
                let r = *row;
                let c = *col;

                let result_row = button(
                    text(display.clone()).size(12).color(theme.text)
                )
                .padding(iced::Padding { top: 6.0, right: 12.0, bottom: 6.0, left: 12.0 })
                .width(Length::Fill)
                .on_press(Message::FindSelect(r, c))
                .style(move |_, _| {
                    button::Style {
                        background: Some(Background::Color(if is_selected { theme.selected } else { theme.bg_cell })),
                        text_color: theme.text,
                        border: Border::default(),
                        shadow: Default::default(),
                    }
                });

                items = items.push(result_row);
            }

            iced::widget::scrollable(items)
                .height(Length::FillPortion(1))
                .into()
        };

        let hint = text("↑↓ Navigate  Enter Jump to cell  Esc Close")
            .size(10)
            .color(theme.text_dim);

        column![
            search_input,
            results_label,
            results_list,
            container(hint).padding(8).center_x(Length::Fill)
        ]
        .spacing(4)
        .padding(8)
        .into()
    }

    pub fn subscription(&self) -> Subscription<Message> {
        iced::event::listen_with(|event, _status, _id| {
            match event {
                Event::Keyboard(keyboard::Event::KeyPressed { key, modifiers, .. }) => {
                    Some(Message::KeyPressed(key, modifiers))
                }
                _ => None,
            }
        })
    }

    /// Convert column index to Excel-style letters (0=A, 25=Z, 26=AA, 27=AB, ...)
    fn col_to_letters(col: usize) -> String {
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

    /// Format a cell reference as A1-style string
    fn cell_ref_string(row: usize, col: usize) -> String {
        format!("{}{}", Self::col_to_letters(col), row + 1)
    }

    /// Update autocomplete suggestions based on current input
    fn update_autocomplete(&mut self) {
        // Only show autocomplete for formulas (starting with =)
        if !self.input_value.starts_with('=') {
            self.autocomplete_visible = false;
            self.autocomplete_suggestions.clear();
            return;
        }

        // Find the current word being typed (function name context)
        // Look for the start of the current identifier
        let input = &self.input_value[1..];  // Skip the '='

        // Find the last position that could start a function name
        // (after =, (, ,, or operators)
        let mut start_pos = 0;
        let mut in_word = false;
        let mut word_start = 0;

        for (i, c) in input.chars().enumerate() {
            if c.is_ascii_alphabetic() {
                if !in_word {
                    in_word = true;
                    word_start = i;
                }
            } else {
                in_word = false;
            }
        }

        // If we're currently typing alphabetic characters, show suggestions
        if in_word {
            let current_word = &input[word_start..];
            let current_word_upper = current_word.to_uppercase();

            // Filter functions that start with the current word
            let suggestions: Vec<usize> = FORMULA_FUNCTIONS
                .iter()
                .enumerate()
                .filter(|(_, f)| f.name.starts_with(&current_word_upper))
                .map(|(i, _)| i)
                .take(8)  // Limit to 8 suggestions
                .collect();

            if !suggestions.is_empty() && current_word.len() >= 1 {
                self.autocomplete_visible = true;
                self.autocomplete_suggestions = suggestions;
                self.autocomplete_selected = 0;
                self.autocomplete_start_pos = word_start + 1;  // +1 for the '='
            } else {
                self.autocomplete_visible = false;
                self.autocomplete_suggestions.clear();
            }
        } else {
            self.autocomplete_visible = false;
            self.autocomplete_suggestions.clear();
        }
    }

    /// Apply the selected autocomplete suggestion
    fn apply_autocomplete(&mut self) {
        if !self.autocomplete_visible || self.autocomplete_suggestions.is_empty() {
            return;
        }

        let selected_idx = self.autocomplete_suggestions[self.autocomplete_selected];
        let function = &FORMULA_FUNCTIONS[selected_idx];

        // Replace the partial function name with the complete one + opening paren
        let before = &self.input_value[..self.autocomplete_start_pos];
        let after_start = self.input_value[self.autocomplete_start_pos..]
            .find(|c: char| !c.is_ascii_alphabetic())
            .map(|i| self.autocomplete_start_pos + i)
            .unwrap_or(self.input_value.len());
        let after = &self.input_value[after_start..];

        self.input_value = format!("{}{}({}", before, function.name, after);

        // Hide autocomplete after applying
        self.autocomplete_visible = false;
        self.autocomplete_suggestions.clear();
    }

    /// Update signature help based on current cursor position in formula
    fn update_signature_help(&mut self) {
        // Only show for formulas
        if !self.input_value.starts_with('=') {
            self.signature_help_visible = false;
            self.signature_help_function = None;
            return;
        }

        let input = &self.input_value[1..];  // Skip '='

        // Find if we're inside function parentheses
        // Track parenthesis depth and find the innermost unclosed function call
        let mut paren_depth = 0;
        let mut func_start: Option<usize> = None;
        let mut func_paren_pos: Option<usize> = None;

        for (i, c) in input.chars().enumerate() {
            match c {
                '(' => {
                    // Look backward for function name
                    if i > 0 {
                        let before = &input[..i];
                        // Find the function name (letters before the paren)
                        let name_start = before.rfind(|c: char| !c.is_ascii_alphabetic()).map(|p| p + 1).unwrap_or(0);
                        let name = &before[name_start..];
                        if !name.is_empty() && name.chars().all(|c| c.is_ascii_alphabetic()) {
                            func_start = Some(name_start);
                            func_paren_pos = Some(i);
                        }
                    }
                    paren_depth += 1;
                }
                ')' => {
                    paren_depth -= 1;
                    if paren_depth == 0 {
                        // Closed the function call
                        func_start = None;
                        func_paren_pos = None;
                    }
                }
                _ => {}
            }
        }

        // If we're inside a function call (unclosed parens)
        if let (Some(name_start), Some(paren_pos)) = (func_start, func_paren_pos) {
            let func_name = input[name_start..paren_pos].to_uppercase();

            // Find the function in FORMULA_FUNCTIONS
            if let Some(func_idx) = FORMULA_FUNCTIONS.iter().position(|f| f.name == func_name) {
                // Count commas after the opening paren to determine current parameter
                let after_paren = &input[paren_pos + 1..];
                let mut param_idx = 0;
                let mut nested_parens = 0;

                for c in after_paren.chars() {
                    match c {
                        '(' => nested_parens += 1,
                        ')' => {
                            if nested_parens > 0 {
                                nested_parens -= 1;
                            }
                        }
                        ',' if nested_parens == 0 => param_idx += 1,
                        _ => {}
                    }
                }

                self.signature_help_visible = true;
                self.signature_help_function = Some(func_idx);
                self.signature_help_param = param_idx;
                return;
            }
        }

        // Not inside a known function call
        self.signature_help_visible = false;
        self.signature_help_function = None;
    }

    /// Parse parameters from a function syntax string like "=SUM(range)" or "=ROUND(value, decimals)"
    fn parse_function_params(syntax: &str) -> Vec<String> {
        // Find content between parentheses
        if let Some(start) = syntax.find('(') {
            if let Some(end) = syntax.rfind(')') {
                let params_str = &syntax[start + 1..end];
                if params_str.is_empty() {
                    return vec![];
                }
                return params_str.split(',').map(|s| s.trim().to_string()).collect();
            }
        }
        vec![]
    }

    /// Get precedents (cells that this cell depends on)
    fn get_precedents(&self, row: usize, col: usize) -> Vec<(usize, usize)> {
        use crate::engine::cell::CellValue;
        use crate::engine::formula::parser::extract_cell_refs;

        let cell = self.sheet.get_cell(row, col);
        if let CellValue::Formula { ast: Some(ref ast), .. } = cell.value {
            let mut refs = extract_cell_refs(ast);
            refs.sort();
            refs.dedup();
            refs
        } else {
            Vec::new()
        }
    }

    /// Get dependents (cells that depend on this cell)
    fn get_dependents(&self, row: usize, col: usize) -> Vec<(usize, usize)> {
        use crate::engine::cell::CellValue;
        use crate::engine::formula::parser::extract_cell_refs;

        let mut dependents = Vec::new();

        // Scan all cells to find those that reference (row, col)
        // Note: This is O(n*m) but acceptable for reasonable sheet sizes
        for r in 0..NUM_ROWS {
            for c in 0..NUM_COLS {
                let cell = self.sheet.get_cell(r, c);
                if let CellValue::Formula { ast: Some(ref ast), .. } = cell.value {
                    let refs = extract_cell_refs(ast);
                    if refs.contains(&(row, col)) {
                        dependents.push((r, c));
                    }
                }
            }
        }

        dependents.sort();
        dependents
    }

    fn view_grid(&self, pane: Option<SplitPane>) -> Element<'_, Message> {
        let theme = self.theme();
        let mut grid_rows: Vec<Element<'_, Message>> = Vec::new();

        // Calculate visible range based on pane
        let (scroll_row, scroll_col) = match pane {
            Some(SplitPane::Secondary) => (self.split_scroll_row, self.split_scroll_col),
            _ => (self.scroll_row, self.scroll_col),
        };
        let start_row = scroll_row;
        let end_row = (start_row + VISIBLE_ROWS).min(NUM_ROWS);
        let start_col = scroll_col;
        let end_col = (start_col + VISIBLE_COLS).min(NUM_COLS);

        // Track which pane this grid is for (for click handling)
        let _this_pane = pane.unwrap_or(SplitPane::Primary);

        // Column headers (double-click to auto-size)
        let mut header_cells: Vec<Element<'_, Message>> = vec![
            // Empty corner cell
            container(text(""))
                .width(ROW_HEADER_WIDTH)
                .height(CELL_HEIGHT)
                .style(move |_| container::Style {
                    background: Some(Background::Color(theme.bg_header)),
                    border: Border {
                        color: theme.gridline,
                        width: 1.0,
                        radius: 0.0.into(),
                    },
                    ..Default::default()
                })
                .into()
        ];

        for c in start_col..end_col {
            let col_label = Self::col_to_letters(c);
            let col_width = self.column_widths[c];
            let is_col_selected = self.selection.ranges().iter().any(|range| {
                c >= range.start_col && c <= range.end_col
            });

            // Column header with resize handle (right 4px is "resize zone")
            let header_content = container(
                text(col_label)
                    .size(11)
                    .color(if is_col_selected { theme.text } else { theme.text_dim })
                    .center()
            )
            .width(col_width)
            .height(CELL_HEIGHT)
            .center_y(CELL_HEIGHT)
            .style(move |_| container::Style {
                background: Some(Background::Color(
                    if is_col_selected { theme.bg_input } else { theme.bg_header }
                )),
                border: Border {
                    color: theme.gridline,
                    width: 1.0,
                    radius: 0.0.into(),
                },
                ..Default::default()
            });

            // Make header clickable for auto-size (simulates double-click behavior)
            let header_btn = button(header_content)
                .padding(0)
                .on_press(Message::ColAutoSize(c))
                .style(move |_, _| button::Style {
                    background: None,
                    text_color: theme.text,
                    border: Border::default(),
                    shadow: Default::default(),
                });

            header_cells.push(header_btn.into());
        }
        grid_rows.push(Row::with_children(header_cells).spacing(0).into());

        // Data rows (only visible ones)
        for r in start_row..end_row {
            let row_height = self.row_heights[r];
            let is_row_selected = self.selection.ranges().iter().any(|range| {
                r >= range.start_row && r <= range.end_row
            });

            // Row header (clickable for auto-size)
            let row_header_content = container(
                text(format!("{}", r + 1))
                    .size(11)
                    .color(if is_row_selected { theme.text } else { theme.text_dim })
                    .center()
            )
            .width(ROW_HEADER_WIDTH)
            .height(row_height)
            .center_y(row_height)
            .style(move |_| container::Style {
                background: Some(Background::Color(
                    if is_row_selected { theme.bg_input } else { theme.bg_header }
                )),
                border: Border {
                    color: theme.gridline,
                    width: 1.0,
                    radius: 0.0.into(),
                },
                ..Default::default()
            });

            let row_header_btn = button(row_header_content)
                .padding(0)
                .on_press(Message::RowAutoSize(r))
                .style(move |_, _| button::Style {
                    background: None,
                    text_color: theme.text,
                    border: Border::default(),
                    shadow: Default::default(),
                });

            let mut row_cells: Vec<Element<'_, Message>> = vec![row_header_btn.into()];

            for c in start_col..end_col {
                let col_width = self.column_widths[c];
                // In formula view mode, show raw formulas; otherwise show computed values
                let value = if self.show_formulas {
                    self.sheet.get_raw(r, c)
                } else {
                    self.sheet.get_formatted_display(r, c)
                };
                let format = self.sheet.get_format(r, c);
                let is_selected = self.selection.contains(r, c);
                let is_active = self.selection.active_cell() == (r, c);

                // Build font based on formatting
                let cell_font = Font {
                    weight: if format.bold { Weight::Bold } else { Weight::Normal },
                    style: if format.italic { FontStyle::Italic } else { FontStyle::Normal },
                    ..Font::DEFAULT
                };

                // Build text with formatting (underline simulated with combining char)
                let display_value = if format.underline && !value.is_empty() {
                    value.chars().map(|c| format!("{}\u{0332}", c)).collect::<String>()
                } else {
                    value
                };

                // Determine horizontal alignment
                let h_alignment = match format.alignment {
                    Alignment::Left => iced::alignment::Horizontal::Left,
                    Alignment::Center => iced::alignment::Horizontal::Center,
                    Alignment::Right => iced::alignment::Horizontal::Right,
                };

                let cell_content = container(
                    text(display_value)
                        .size(12)
                        .color(theme.text)
                        .font(cell_font)
                        .align_x(h_alignment)
                        .width(Length::Fill)
                )
                .width(col_width)
                .height(row_height)
                .padding(iced::Padding { top: 0.0, right: 4.0, bottom: 0.0, left: 4.0 })
                .center_y(row_height)
                .style(move |_| container::Style {
                    background: Some(Background::Color(
                        if is_selected {
                            theme.selected
                        } else {
                            theme.bg_cell
                        }
                    )),
                    border: Border {
                        color: if is_active {
                            theme.selected_border
                        } else if is_selected {
                            theme.selected_border
                        } else {
                            theme.gridline
                        },
                        width: if is_active { 2.0 } else { 1.0 },
                        radius: 0.0.into(),
                    },
                    ..Default::default()
                });

                // Wrap in button for click handling
                let cell = button(cell_content)
                    .padding(0)
                    .on_press_with(move || Message::CellClicked(r, c, Modifiers::default()))
                    .style(move |_, _| button::Style {
                        background: None,
                        text_color: theme.text,
                        border: Border::default(),
                        shadow: Default::default(),
                    });

                row_cells.push(cell.into());
            }

            grid_rows.push(Row::with_children(row_cells).spacing(0).into());
        }

        container(
            Column::with_children(grid_rows).spacing(0)
        )
        .width(Length::Fill)
        .height(Length::Fill)
        .style(move |_| container::Style {
            background: Some(Background::Color(theme.bg_cell)),
            ..Default::default()
        })
        .into()
    }
}
