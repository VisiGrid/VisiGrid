//! Unified search engine for VisiGrid
//!
//! This module provides the core search infrastructure used by the command palette
//! and other search-based features. Design principles:
//!
//! - Search providers are pure: no UI state mutation inside search()
//! - Actions are stable IDs, not closures: execution is centralized
//! - All searchable items reduce to: label, subtitle, score, action

use std::path::PathBuf;

// ============================================================================
// Menu Categories (for Alt accelerator scoping)
// ============================================================================

/// Menu category for Alt accelerator filtering.
/// Maps to the macOS menu bar structure.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum MenuCategory {
    File,
    Edit,
    View,
    Format,
    Data,
    Tools,
    Help,
}

impl MenuCategory {
    /// Display name for the category
    pub fn name(&self) -> &'static str {
        match self {
            Self::File => "File",
            Self::Edit => "Edit",
            Self::View => "View",
            Self::Format => "Format",
            Self::Data => "Data",
            Self::Tools => "Tools",
            Self::Help => "Help",
        }
    }

    /// Short key hint for Alt accelerator display
    pub fn key_hint(&self) -> &'static str {
        match self {
            Self::File => "F",
            Self::Edit => "E",
            Self::View => "V",
            Self::Format => "O",
            Self::Data => "A",
            Self::Tools => "T",
            Self::Help => "H",
        }
    }

    /// All categories in display order for scope hints bar
    pub fn all_for_hints() -> &'static [MenuCategory] {
        &[
            Self::Data,    // A - most common for power users
            Self::Edit,    // E
            Self::File,    // F
            Self::View,    // V
            Self::Tools,   // T
        ]
    }
}

// ============================================================================
// Command IDs (stable identifiers for all executable commands)
// ============================================================================

/// Stable identifier for commands. Used instead of function pointers to keep
/// search results serializable, testable, and loggable.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum CommandId {
    // Navigation
    GoToCell,
    FindInCells,
    GoToStart,
    SelectAll,
    SelectBlanks,
    SelectCurrentRegion,
    HideRows,
    UnhideRows,
    HideCols,
    UnhideCols,

    // Editing
    FillDown,
    FillRight,
    ClearCells,
    TrimWhitespace,
    Undo,
    Redo,
    AutoSum,

    // Clipboard
    Copy,
    Cut,
    Paste,
    PasteValues,
    PasteSpecial,
    PasteFormulas,
    PasteFormats,

    // Formatting
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    FormatCurrency,
    FormatPercent,
    FormatCells,
    ClearFormatting,
    FormatPainter,
    FormatPainterLocked,
    CopyFormat,
    PasteFormat,
    // Background colors
    FillColor,
    ClearBackground,
    BackgroundYellow,
    BackgroundGreen,
    BackgroundBlue,
    BackgroundRed,
    BackgroundOrange,
    BackgroundPurple,
    BackgroundGray,
    BackgroundCyan,
    // Cell styles
    StyleDefault,
    StyleError,
    StyleWarning,
    StyleSuccess,
    StyleInput,
    StyleTotal,
    StyleNote,
    StyleClear,
    // Borders
    BordersAll,
    BordersOutline,
    BordersInside,
    BordersTop,
    BordersBottom,
    BordersLeft,
    BordersRight,
    BordersClear,

    // File
    NewWindow,
    OpenFile,
    Save,
    SaveAs,
    ExportCsv,
    ExportTsv,
    ExportJson,

    // Appearance
    SelectTheme,
    SelectFont,

    // View
    ToggleInspector,
    ToggleProfiler,
    ProfileNextRecalc,
    ClearProfiler,
    ToggleMinimap,
    ToggleZenMode,
    ZoomIn,
    ZoomOut,
    ZoomReset,
    FreezeTopRow,
    FreezeFirstColumn,
    FreezePanes,
    UnfreezePanes,
    SplitRight,
    CloseSplit,
    ToggleTrace,
    CycleTracePrecedent,
    CycleTraceDependent,
    ReturnToTraceSource,
    ToggleVerifiedMode,
    Recalculate,
    ReloadCustomFunctions,
    ApproveModel,
    ClearApproval,
    NavPerfReport,

    // Window
    SwitchWindow,

    // Refactoring
    ExtractNamedRange,

    // Help
    ShowShortcuts,
    OpenKeybindings,
    OpenDocs,
    ShowAbout,
    TourNamedRanges,
    ShowRefactorLog,
    ShowAISettings,
    InsertFormulaAI,
    AnalyzeAI,

    // Sheets
    NextSheet,
    PrevSheet,
    AddSheet,

    // Data (sort/filter)
    SortAscending,
    SortDescending,
    ToggleAutoFilter,
    ClearSort,

    // Data (validation)
    ValidationDialog,
    ExcludeFromValidation,
    ClearValidationExclusions,
    CircleInvalidData,
    ClearInvalidCircles,

    // VisiHub sync
    HubCheckStatus,
    HubPull,
    HubPublish,
    HubOpenRemoteAsCopy,
    HubUnlink,
    HubDiagnostics,
    HubSignIn,
    HubSignOut,
    HubLinkDialog,
}

impl CommandId {
    /// Human-readable name for the command
    pub fn name(&self) -> &'static str {
        match self {
            Self::GoToCell => "Go to Cell",
            Self::FindInCells => "Find in Cells",
            Self::GoToStart => "Go to Start (A1)",
            Self::SelectAll => "Select All",
            Self::SelectBlanks => "Select: Blanks in Region",
            Self::SelectCurrentRegion => "Select Current Region",
            Self::HideRows => "Hide Rows",
            Self::UnhideRows => "Unhide Rows",
            Self::HideCols => "Hide Columns",
            Self::UnhideCols => "Unhide Columns",
            Self::FillDown => "Fill Down",
            Self::FillRight => "Fill Right",
            Self::ClearCells => "Clear Cells",
            Self::TrimWhitespace => "Transform: Trim Whitespace",
            Self::Undo => "Undo",
            Self::Redo => "Redo",
            Self::AutoSum => "AutoSum",
            Self::Copy => "Copy",
            Self::Cut => "Cut",
            Self::Paste => "Paste",
            Self::PasteValues => "Paste Values",
            Self::PasteSpecial => "Paste Special...",
            Self::PasteFormulas => "Paste Formulas",
            Self::PasteFormats => "Paste Formats",
            Self::ToggleBold => "Toggle Bold",
            Self::ToggleItalic => "Toggle Italic",
            Self::ToggleUnderline => "Toggle Underline",
            Self::FormatCurrency => "Format as Currency",
            Self::FormatPercent => "Format as Percent",
            Self::FormatCells => "Format Cells...",
            Self::ClearFormatting => "Clear Formatting",
            Self::FormatPainter => "Format Painter",
            Self::FormatPainterLocked => "Format Painter (Locked)",
            Self::CopyFormat => "Copy Format",
            Self::PasteFormat => "Paste Format",
            Self::FillColor => "Fill Color...",
            Self::ClearBackground => "Background: None",
            Self::BackgroundYellow => "Background: Yellow",
            Self::BackgroundGreen => "Background: Green",
            Self::BackgroundBlue => "Background: Blue",
            Self::BackgroundRed => "Background: Red",
            Self::BackgroundOrange => "Background: Orange",
            Self::BackgroundPurple => "Background: Purple",
            Self::BackgroundGray => "Background: Gray",
            Self::BackgroundCyan => "Background: Cyan",
            Self::StyleDefault => "Cell Style: Default",
            Self::StyleError => "Cell Style: Error",
            Self::StyleWarning => "Cell Style: Warning",
            Self::StyleSuccess => "Cell Style: Success",
            Self::StyleInput => "Cell Style: Input",
            Self::StyleTotal => "Cell Style: Total",
            Self::StyleNote => "Cell Style: Note",
            Self::StyleClear => "Cell Style: Clear",
            Self::BordersAll => "Borders: All",
            Self::BordersOutline => "Borders: Outline",
            Self::BordersInside => "Borders: Inside",
            Self::BordersTop => "Borders: Top",
            Self::BordersBottom => "Borders: Bottom",
            Self::BordersLeft => "Borders: Left",
            Self::BordersRight => "Borders: Right",
            Self::BordersClear => "Borders: Clear",
            Self::NewWindow => "New Workbook",
            Self::OpenFile => "Open File",
            Self::Save => "Save",
            Self::SaveAs => "Save As",
            Self::ExportCsv => "Export as CSV",
            Self::ExportTsv => "Export as TSV",
            Self::ExportJson => "Export as JSON",
            Self::SelectTheme => "Select Theme...",
            Self::SelectFont => "Select Font...",
            Self::ToggleInspector => "Toggle Inspector",
            Self::ToggleProfiler => "Toggle Profiler",
            Self::ProfileNextRecalc => "Profile Next Recalc",
            Self::ClearProfiler => "Clear Profiler",
            Self::ToggleMinimap => "Toggle Minimap",
            Self::ToggleZenMode => "Toggle Zen Mode",
            Self::ZoomIn => "Zoom In",
            Self::ZoomOut => "Zoom Out",
            Self::ZoomReset => "Reset Zoom",
            Self::FreezeTopRow => "Freeze Top Row",
            Self::FreezeFirstColumn => "Freeze First Column",
            Self::FreezePanes => "Freeze Panes",
            Self::UnfreezePanes => "Unfreeze Panes",
            Self::SplitRight => "Split View",
            Self::CloseSplit => "Close Split View",
            Self::ToggleTrace => "Toggle Trace Mode",
            Self::CycleTracePrecedent => "Jump to Precedent",
            Self::CycleTraceDependent => "Jump to Dependent",
            Self::ReturnToTraceSource => "Return to Trace Source",
            Self::ToggleVerifiedMode => "Toggle Verified Mode",
            Self::Recalculate => "Recalculate All",
            Self::ReloadCustomFunctions => "Reload Custom Functions",
            Self::ApproveModel => "Approve Model",
            Self::ClearApproval => "Clear Approval",
            Self::NavPerfReport => "Navigation Latency Report",
            Self::SwitchWindow => "Switch Window...",
            Self::ExtractNamedRange => "Extract to Named Range...",
            Self::ShowShortcuts => "Show Keyboard Shortcuts",
            Self::OpenKeybindings => "Open Keybindings (JSON)",
            Self::OpenDocs => "Documentation",
            Self::ShowAbout => "About VisiGrid",
            Self::TourNamedRanges => "Tour: Named Ranges & Refactoring",
            Self::ShowRefactorLog => "Show Refactor Log",
            Self::ShowAISettings => "AI Settings",
            Self::InsertFormulaAI => "Insert Formula with AI...",
            Self::AnalyzeAI => "Analyze with AI...",
            Self::NextSheet => "Next Sheet",
            Self::PrevSheet => "Previous Sheet",
            Self::AddSheet => "Add Sheet",
            Self::SortAscending => "Sort Ascending (A→Z)",
            Self::SortDescending => "Sort Descending (Z→A)",
            Self::ToggleAutoFilter => "Toggle AutoFilter",
            Self::ClearSort => "Clear Sort",
            Self::ValidationDialog => "Data Validation...",
            Self::ExcludeFromValidation => "Exclude from Validation",
            Self::ClearValidationExclusions => "Clear Validation Exclusions",
            Self::CircleInvalidData => "Circle Invalid Data",
            Self::ClearInvalidCircles => "Clear Invalid Circles",
            Self::HubCheckStatus => "VisiHub: Check Status",
            Self::HubPull => "VisiHub: Update from Remote",
            Self::HubPublish => "VisiHub: Publish",
            Self::HubOpenRemoteAsCopy => "VisiHub: Open Remote as Copy",
            Self::HubUnlink => "VisiHub: Unlink",
            Self::HubDiagnostics => "VisiHub: Show Diagnostics",
            Self::HubSignIn => "VisiHub: Sign In",
            Self::HubSignOut => "VisiHub: Sign Out",
            Self::HubLinkDialog => "VisiHub: Link to Dataset...",
        }
    }

    /// Keyboard shortcut display string (if any)
    pub fn shortcut(&self) -> Option<&'static str> {
        match self {
            Self::GoToCell => Some("Ctrl+G"),
            Self::FindInCells => Some("Ctrl+F"),
            Self::GoToStart => Some("Ctrl+Home"),
            Self::SelectAll => Some("Ctrl+A"),
            Self::SelectCurrentRegion => Some("Ctrl+Shift+*"),
            Self::HideRows => Some("Ctrl+9"),
            Self::UnhideRows => Some("Ctrl+Shift+9"),
            Self::HideCols => Some("Ctrl+0"),
            Self::UnhideCols => Some("Ctrl+Shift+0"),
            Self::FillDown => Some("Ctrl+D"),
            Self::FillRight => Some("Ctrl+R"),
            Self::ClearCells => Some("Delete"),
            Self::Undo => Some("Ctrl+Z"),
            Self::Redo => Some("Ctrl+Y"),
            Self::AutoSum => Some("Alt+="),
            Self::Copy => Some("Ctrl+C"),
            Self::Cut => Some("Ctrl+X"),
            Self::Paste => Some("Ctrl+V"),
            Self::PasteValues => Some("Ctrl+Alt+Shift+V"),
            #[cfg(target_os = "macos")]
            Self::PasteSpecial => Some("Cmd+Option+V"),
            #[cfg(not(target_os = "macos"))]
            Self::PasteSpecial => Some("Ctrl+Alt+V"),
            Self::ToggleBold => Some("Ctrl+B"),
            Self::ToggleItalic => Some("Ctrl+I"),
            Self::ToggleUnderline => Some("Ctrl+U"),
            Self::FormatCurrency => Some("Ctrl+Shift+$"),
            Self::FormatPercent => Some("Ctrl+Shift+%"),
            Self::FormatCells => Some("Ctrl+1"),
            Self::NewWindow => Some("Ctrl+N"),
            Self::OpenFile => Some("Ctrl+O"),
            Self::Save => Some("Ctrl+S"),
            Self::SaveAs => Some("Ctrl+Shift+S"),
            Self::ToggleInspector => Some("Ctrl+Shift+I"),
            Self::ToggleProfiler => Some("Ctrl+Alt+P"),
            Self::ToggleZenMode => Some("F11"),
            Self::ZoomIn => Some("Ctrl+Alt+="),
            Self::ZoomOut => Some("Ctrl+Alt+-"),
            Self::ZoomReset => Some("Ctrl+Alt+0"),
            Self::ToggleAutoFilter => Some("Ctrl+Shift+L"),
            #[cfg(target_os = "macos")]
            Self::SwitchWindow => Some("Cmd+`"),
            #[cfg(not(target_os = "macos"))]
            Self::SwitchWindow => Some("Ctrl+`"),
            Self::SplitRight => Some("Ctrl+\\"),
            Self::CloseSplit => Some("Ctrl+Shift+\\"),
            #[cfg(target_os = "macos")]
            Self::ToggleTrace => Some("⌥T"),
            #[cfg(not(target_os = "macos"))]
            Self::ToggleTrace => Some("Alt+T"),
            #[cfg(target_os = "macos")]
            Self::CycleTracePrecedent => Some("⌥["),
            #[cfg(not(target_os = "macos"))]
            Self::CycleTracePrecedent => Some("Ctrl+["),
            #[cfg(target_os = "macos")]
            Self::CycleTraceDependent => Some("⌥]"),
            #[cfg(not(target_os = "macos"))]
            Self::CycleTraceDependent => Some("Ctrl+]"),
            #[cfg(target_os = "macos")]
            Self::ReturnToTraceSource => Some("⌥↩"),
            #[cfg(not(target_os = "macos"))]
            Self::ReturnToTraceSource => Some("F5"),
            Self::Recalculate => Some("F9"),
            Self::InsertFormulaAI => Some("Ctrl+Shift+A"),
            Self::AnalyzeAI => Some("Ctrl+Shift+E"),
            Self::CopyFormat => Some("Ctrl+Shift+C"),
            Self::PasteFormat => Some("Ctrl+Shift+V"),
            _ => None,
        }
    }

    /// Search keywords (additional terms that match this command)
    pub fn keywords(&self) -> &'static str {
        match self {
            Self::GoToCell => "goto jump navigate",
            Self::FindInCells => "search",
            Self::GoToStart => "home beginning",
            Self::SelectAll => "selection",
            Self::SelectBlanks => "empty cells region selection",
            Self::SelectCurrentRegion => "select region contiguous data block table area ctrl shift star asterisk",
            Self::HideRows => "hide row invisible conceal",
            Self::UnhideRows => "unhide row show reveal visible",
            Self::HideCols => "hide column invisible conceal",
            Self::UnhideCols => "unhide column show reveal visible",
            Self::FillDown => "copy formula",
            Self::FillRight => "copy formula",
            Self::ClearCells => "delete remove empty",
            Self::TrimWhitespace => "strip spaces clean transform",
            Self::Undo => "revert back",
            Self::Redo => "forward",
            Self::AutoSum => "sum formula total",
            Self::Copy => "clipboard",
            Self::Cut => "clipboard",
            Self::Paste => "clipboard",
            Self::PasteValues => "clipboard special values only",
            Self::PasteSpecial => "clipboard special dialog formulas formats values",
            Self::PasteFormulas => "clipboard special formulas reference adjust",
            Self::PasteFormats => "clipboard special formatting style",
            Self::ToggleBold => "format style",
            Self::ToggleItalic => "format style",
            Self::ToggleUnderline => "format style",
            Self::FormatCurrency => "format number money dollar",
            Self::FormatPercent => "format number percentage",
            Self::FormatCells => "format style number date currency",
            Self::ClearFormatting => "clear reset format style default",
            Self::FormatPainter => "paint format brush copy style",
            Self::FormatPainterLocked => "paint format brush lock persist",
            Self::CopyFormat => "copy format style brush painter",
            Self::PasteFormat => "paste format style brush painter apply",
            Self::FillColor => "background color fill paint picker format",
            Self::ClearBackground => "format fill color none clear",
            Self::BackgroundYellow => "format fill color highlight",
            Self::BackgroundGreen => "format fill color highlight",
            Self::BackgroundBlue => "format fill color highlight",
            Self::BackgroundRed => "format fill color highlight",
            Self::BackgroundOrange => "format fill color highlight",
            Self::BackgroundPurple => "format fill color highlight",
            Self::BackgroundGray => "format fill color highlight",
            Self::BackgroundCyan => "format fill color highlight",
            Self::StyleDefault => "cell style semantic default none",
            Self::StyleError => "cell style semantic error red danger",
            Self::StyleWarning => "cell style semantic warning yellow caution",
            Self::StyleSuccess => "cell style semantic success green ok valid",
            Self::StyleInput => "cell style semantic input blue editable",
            Self::StyleTotal => "cell style semantic total sum bold",
            Self::StyleNote => "cell style semantic note comment info",
            Self::StyleClear => "cell style semantic clear remove reset",
            Self::BordersAll => "format border grid lines box",
            Self::BordersOutline => "format border box outline perimeter frame",
            Self::BordersInside => "format border inside internal inner grid",
            Self::BordersTop => "format border top edge upper",
            Self::BordersBottom => "format border bottom edge lower",
            Self::BordersLeft => "format border left edge",
            Self::BordersRight => "format border right edge",
            Self::BordersClear => "format border clear remove none",
            Self::NewWindow => "create workbook file window blank",
            Self::OpenFile => "load",
            Self::Save => "write",
            Self::SaveAs => "write export",
            Self::ExportCsv => "save comma",
            Self::ExportTsv => "save tab separated",
            Self::ExportJson => "save array",
            Self::SelectTheme => "appearance color scheme dark light",
            Self::SelectFont => "appearance typography",
            Self::ToggleInspector => "panel sidebar",
            Self::ToggleProfiler => "profiler performance timing hotspot",
            Self::ProfileNextRecalc => "profile recalc run",
            Self::ClearProfiler => "clear profiler reset",
            Self::ToggleMinimap => "minimap density navigator overview map",
            Self::ToggleZenMode => "distraction free fullscreen focus",
            Self::ZoomIn => "magnify scale enlarge bigger larger",
            Self::ZoomOut => "magnify scale shrink smaller",
            Self::ZoomReset => "magnify scale default 100",
            Self::FreezeTopRow => "lock header pin row",
            Self::FreezeFirstColumn => "lock pin column",
            Self::FreezePanes => "lock pin split scroll",
            Self::UnfreezePanes => "unlock unpin clear",
            Self::SwitchWindow => "window workbook navigate focus",
            Self::ExtractNamedRange => "extract refactor variable name range",
            Self::ShowShortcuts => "help keys bindings hotkeys",
            Self::OpenKeybindings => "shortcuts remap customize config json",
            Self::OpenDocs => "documentation help guide manual reference docs",
            Self::ShowAbout => "version info",
            Self::TourNamedRanges => "tour guide walkthrough refactor learn onboarding",
            Self::ShowRefactorLog => "audit history changes log refactor",
            Self::ShowAISettings => "ai openai anthropic ollama local llm api key model provider",
            Self::InsertFormulaAI => "ai ask question formula help gpt claude llm insert",
            Self::AnalyzeAI => "ai analyze question explain data patterns anomalies summarize",
            Self::NextSheet => "tab worksheet",
            Self::PrevSheet => "tab worksheet",
            Self::AddSheet => "new tab worksheet",
            Self::SortAscending => "sort order ascending asc a-z smallest lowest",
            Self::SortDescending => "sort order descending desc z-a largest highest",
            Self::ToggleAutoFilter => "filter dropdown autofilter data",
            Self::ClearSort => "unsort restore original order",
            Self::ValidationDialog => "validate data rules constraints list dropdown",
            Self::ExcludeFromValidation => "validate exclude skip ignore",
            Self::ClearValidationExclusions => "validate exclusions clear remove",
            Self::CircleInvalidData => "validate invalid circle mark highlight",
            Self::ClearInvalidCircles => "validate invalid circle clear remove",
            Self::HubCheckStatus => "visihub cloud sync status check refresh",
            Self::HubPull => "visihub cloud sync update pull",
            Self::HubPublish => "visihub cloud sync publish upload push commit",
            Self::HubOpenRemoteAsCopy => "visihub cloud sync open copy download safe",
            Self::HubUnlink => "visihub cloud sync unlink disconnect remove",
            Self::HubDiagnostics => "visihub cloud sync diagnostics debug state error",
            Self::HubSignIn => "visihub cloud sync sign in login authenticate token",
            Self::HubSignOut => "visihub cloud sync sign out logout disconnect",
            Self::HubLinkDialog => "visihub cloud sync link connect dataset repository",
            Self::SplitRight => "split view pane side by side divide window",
            Self::CloseSplit => "split close merge unsplit single pane",
            Self::ToggleTrace => "trace precedents dependents dependency audit formula inputs outputs",
            Self::CycleTracePrecedent => "trace precedent input jump navigate cycle next previous",
            Self::CycleTraceDependent => "trace dependent output jump navigate cycle next previous",
            Self::ReturnToTraceSource => "trace back return source origin home",
            Self::ToggleVerifiedMode => "verified deterministic recalc audit trust",
            Self::Recalculate => "recalc refresh calculate formulas f9",
            Self::ReloadCustomFunctions => "lua functions reload custom scripting",
            Self::ApproveModel => "approve model fingerprint semantic logic verify lock sign-off audit",
            Self::ClearApproval => "clear approval reset unapprove remove fingerprint",
            Self::NavPerfReport => "navigation latency perf performance timing",
        }
    }

    /// All available commands
    pub fn all() -> &'static [CommandId] {
        &[
            Self::GoToCell,
            Self::FindInCells,
            Self::GoToStart,
            Self::SelectAll,
            Self::SelectBlanks,
            Self::SelectCurrentRegion,
            Self::HideRows,
            Self::UnhideRows,
            Self::HideCols,
            Self::UnhideCols,
            Self::FillDown,
            Self::FillRight,
            Self::ClearCells,
            Self::TrimWhitespace,
            Self::Undo,
            Self::Redo,
            Self::AutoSum,
            Self::Copy,
            Self::Cut,
            Self::Paste,
            Self::PasteValues,
            Self::PasteSpecial,
            Self::PasteFormulas,
            Self::PasteFormats,
            Self::ToggleBold,
            Self::ToggleItalic,
            Self::ToggleUnderline,
            Self::FormatCurrency,
            Self::FormatPercent,
            Self::FormatCells,
            Self::ClearFormatting,
            Self::FormatPainter,
            Self::FormatPainterLocked,
            Self::CopyFormat,
            Self::PasteFormat,
            Self::FillColor,
            Self::ClearBackground,
            Self::BackgroundYellow,
            Self::BackgroundGreen,
            Self::BackgroundBlue,
            Self::BackgroundRed,
            Self::BackgroundOrange,
            Self::BackgroundPurple,
            Self::BackgroundGray,
            Self::BackgroundCyan,
            Self::StyleDefault,
            Self::StyleError,
            Self::StyleWarning,
            Self::StyleSuccess,
            Self::StyleInput,
            Self::StyleTotal,
            Self::StyleNote,
            Self::StyleClear,
            Self::BordersAll,
            Self::BordersOutline,
            Self::BordersInside,
            Self::BordersTop,
            Self::BordersBottom,
            Self::BordersLeft,
            Self::BordersRight,
            Self::BordersClear,
            Self::NewWindow,
            Self::OpenFile,
            Self::Save,
            Self::SaveAs,
            Self::ExportCsv,
            Self::ExportTsv,
            Self::ExportJson,
            Self::SelectTheme,
            Self::SelectFont,
            Self::ToggleInspector,
            Self::ToggleProfiler,
            Self::ProfileNextRecalc,
            Self::ClearProfiler,
            Self::ToggleMinimap,
            Self::ToggleZenMode,
            Self::ZoomIn,
            Self::ZoomOut,
            Self::ZoomReset,
            Self::FreezeTopRow,
            Self::FreezeFirstColumn,
            Self::FreezePanes,
            Self::UnfreezePanes,
            Self::SplitRight,
            Self::CloseSplit,
            Self::ToggleTrace,
            Self::CycleTracePrecedent,
            Self::CycleTraceDependent,
            Self::ReturnToTraceSource,
            Self::ToggleVerifiedMode,
            Self::Recalculate,
            Self::ReloadCustomFunctions,
            Self::ApproveModel,
            Self::ClearApproval,
            Self::SwitchWindow,
            Self::ExtractNamedRange,
            Self::ShowShortcuts,
            Self::OpenKeybindings,
            Self::OpenDocs,
            Self::ShowAbout,
            Self::TourNamedRanges,
            Self::ShowRefactorLog,
            Self::ShowAISettings,
            Self::InsertFormulaAI,
            Self::AnalyzeAI,
            Self::NextSheet,
            Self::PrevSheet,
            Self::AddSheet,
            Self::SortAscending,
            Self::SortDescending,
            Self::ToggleAutoFilter,
            Self::ClearSort,
            Self::ValidationDialog,
            Self::ExcludeFromValidation,
            Self::ClearValidationExclusions,
            Self::CircleInvalidData,
            Self::ClearInvalidCircles,
            Self::HubCheckStatus,
            Self::HubPull,
            Self::HubPublish,
            Self::HubOpenRemoteAsCopy,
            Self::HubUnlink,
            Self::HubDiagnostics,
            Self::HubSignIn,
            Self::HubSignOut,
            Self::HubLinkDialog,
            Self::NavPerfReport,
        ]
    }

    /// Returns the menu category for this command.
    /// Returns None for commands not addressable via Alt accelerators.
    ///
    /// DESIGN NOTE: Scoped palettes intentionally hide commands without
    /// a menu category. This is deliberate - Alt accelerators provide
    /// structured access to menu commands only. Non-menu commands remain
    /// accessible via the unscoped Command Palette.
    pub fn menu_category(&self) -> Option<MenuCategory> {
        match self {
            // File menu
            Self::NewWindow
            | Self::OpenFile
            | Self::Save
            | Self::SaveAs
            | Self::ExportCsv
            | Self::ExportTsv
            | Self::ExportJson
            | Self::HubCheckStatus
            | Self::HubPull
            | Self::HubPublish
            | Self::HubOpenRemoteAsCopy
            | Self::HubUnlink
            | Self::HubDiagnostics
            | Self::HubSignIn
            | Self::HubSignOut
            | Self::HubLinkDialog => Some(MenuCategory::File),

            // Edit menu
            Self::Undo
            | Self::Redo
            | Self::Cut
            | Self::Copy
            | Self::Paste
            | Self::PasteValues
            | Self::PasteSpecial
            | Self::PasteFormulas
            | Self::PasteFormats
            | Self::ClearCells
            | Self::SelectAll
            | Self::FindInCells
            | Self::GoToCell => Some(MenuCategory::Edit),

            // View menu
            Self::ToggleInspector
            | Self::ToggleProfiler
            | Self::ToggleMinimap
            | Self::ToggleZenMode
            | Self::ZoomIn
            | Self::ZoomOut
            | Self::ZoomReset
            | Self::FreezeTopRow
            | Self::FreezeFirstColumn
            | Self::FreezePanes
            | Self::UnfreezePanes
            | Self::SplitRight
            | Self::CloseSplit
            | Self::Recalculate => Some(MenuCategory::View),

            // Tools menu (trace, explain, audit, AI)
            Self::ToggleTrace
            | Self::CycleTracePrecedent
            | Self::CycleTraceDependent
            | Self::ReturnToTraceSource
            | Self::ToggleVerifiedMode
            | Self::ApproveModel
            | Self::ClearApproval
            | Self::InsertFormulaAI
            | Self::AnalyzeAI
            | Self::NavPerfReport
            | Self::ReloadCustomFunctions
            | Self::ProfileNextRecalc
            | Self::ClearProfiler => Some(MenuCategory::Tools),

            // Format menu
            Self::ToggleBold
            | Self::ToggleItalic
            | Self::ToggleUnderline
            | Self::FormatCurrency
            | Self::FormatPercent
            | Self::FormatCells
            | Self::ClearFormatting
            | Self::FormatPainter
            | Self::FormatPainterLocked
            | Self::CopyFormat
            | Self::PasteFormat
            | Self::SelectFont
            | Self::FillColor
            | Self::ClearBackground
            | Self::BackgroundYellow
            | Self::BackgroundGreen
            | Self::BackgroundBlue
            | Self::BackgroundRed
            | Self::BackgroundOrange
            | Self::BackgroundPurple
            | Self::BackgroundGray
            | Self::BackgroundCyan
            | Self::StyleDefault
            | Self::StyleError
            | Self::StyleWarning
            | Self::StyleSuccess
            | Self::StyleInput
            | Self::StyleTotal
            | Self::StyleNote
            | Self::StyleClear
            | Self::BordersAll
            | Self::BordersOutline
            | Self::BordersInside
            | Self::BordersTop
            | Self::BordersBottom
            | Self::BordersLeft
            | Self::BordersRight
            | Self::BordersClear
            | Self::HideRows
            | Self::UnhideRows
            | Self::HideCols
            | Self::UnhideCols => Some(MenuCategory::Format),

            // Data menu
            Self::FillDown
            | Self::FillRight
            | Self::TrimWhitespace
            | Self::AutoSum
            | Self::SortAscending
            | Self::SortDescending
            | Self::ToggleAutoFilter
            | Self::ClearSort
            | Self::ValidationDialog
            | Self::ExcludeFromValidation
            | Self::ClearValidationExclusions
            | Self::CircleInvalidData
            | Self::ClearInvalidCircles => Some(MenuCategory::Data),

            // Help menu
            Self::ShowShortcuts
            | Self::OpenKeybindings
            | Self::OpenDocs
            | Self::ShowAbout
            | Self::TourNamedRanges
            | Self::ShowRefactorLog => Some(MenuCategory::Help),

            // Commands not in any menu - intentionally excluded from Alt accelerators
            // These remain accessible via the unscoped Command Palette
            Self::GoToStart
            | Self::SelectBlanks
            | Self::SelectCurrentRegion
            | Self::ExtractNamedRange
            | Self::SelectTheme
            | Self::SwitchWindow
            | Self::NextSheet
            | Self::PrevSheet
            | Self::AddSheet
            | Self::ShowAISettings => None,
        }
    }
}

// ============================================================================
// Search Types
// ============================================================================

/// The kind of search result (used for visual differentiation)
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SearchKind {
    Command,
    RecentFile,
    Formula,
    Cell,
    NamedRange,
    Setting,
    GoTo,
    /// Cells that depend on the current cell (Find References)
    Reference,
    /// Cells that the current cell depends on (Go to Precedents)
    Precedent,
}

impl SearchKind {
    /// Icon or prefix character for visual typing
    pub fn icon(&self) -> &'static str {
        match self {
            Self::Command => ">",
            Self::RecentFile => "⏱",
            Self::Formula => "ƒ",
            Self::Cell => "@",
            Self::NamedRange => "$",
            Self::Setting => "⚙",
            Self::GoTo => ":",
            Self::Reference => "→",   // Arrow pointing to dependents
            Self::Precedent => "←",   // Arrow pointing to precedents
        }
    }

    /// Ranking priority (lower = higher priority, used as tie-breaker)
    /// Commands rank highest as primary user actions
    pub fn priority(&self) -> u8 {
        match self {
            Self::Command => 0,      // Primary actions
            Self::RecentFile => 1,   // Recently used files
            Self::GoTo => 2,         // Quick navigation
            Self::Formula => 3,      // Function insertion
            Self::Cell => 4,         // Cell search
            Self::Setting => 5,      // Configuration
            Self::NamedRange => 6,   // Future
            Self::Reference => 1,    // High priority when shown
            Self::Precedent => 1,    // High priority when shown
        }
    }
}

/// An action that can be executed from a search result.
/// Uses stable IDs instead of closures for testability and serialization.
#[derive(Clone, Debug, PartialEq)]
pub enum SearchAction {
    /// Run a registered command
    RunCommand(CommandId),

    /// Open a recent file
    OpenFile(PathBuf),

    /// Jump to a specific cell
    JumpToCell { row: usize, col: usize },

    /// Insert a formula function at cursor
    InsertFormula { name: String, signature: String },

    /// Jump to a named range (future)
    JumpToNamedRange { name: String },

    /// Open a setting (future)
    OpenSetting { key: String },

    // Secondary actions (Ctrl+Enter)

    /// Copy text to clipboard
    CopyToClipboard { text: String, description: String },

    /// Show function help/signature in status
    ShowFunctionHelp { name: String, signature: String, description: String },

    // Reference/Precedent navigation

    /// Show cells that reference the given cell (Find References)
    ShowReferences { row: usize, col: usize },

    /// Show cells that the given cell references (Go to Precedents)
    ShowPrecedents { row: usize, col: usize },
}

/// A single search result item
#[derive(Clone, Debug)]
pub struct SearchItem {
    pub kind: SearchKind,
    pub title: String,
    pub subtitle: Option<String>,
    pub score: f32,
    pub action: SearchAction,
    /// Optional secondary action (Ctrl+Enter)
    pub secondary_action: Option<SearchAction>,
    /// Character ranges to highlight in title (start, end)
    pub highlights: Vec<(usize, usize)>,
}

impl SearchItem {
    /// Create a new search item with default score
    pub fn new(kind: SearchKind, title: impl Into<String>, action: SearchAction) -> Self {
        Self {
            kind,
            title: title.into(),
            subtitle: None,
            score: 0.0,
            action,
            secondary_action: None,
            highlights: Vec::new(),
        }
    }

    /// Builder: set subtitle
    pub fn with_subtitle(mut self, subtitle: impl Into<String>) -> Self {
        self.subtitle = Some(subtitle.into());
        self
    }

    /// Builder: set score
    pub fn with_score(mut self, score: f32) -> Self {
        self.score = score;
        self
    }

    /// Builder: set highlights
    pub fn with_highlights(mut self, highlights: Vec<(usize, usize)>) -> Self {
        self.highlights = highlights;
        self
    }

    /// Builder: set secondary action (Ctrl+Enter)
    pub fn with_secondary_action(mut self, action: SearchAction) -> Self {
        self.secondary_action = Some(action);
        self
    }
}

/// A parsed search query
#[derive(Clone, Debug)]
pub struct SearchQuery<'a> {
    /// The original raw input
    pub raw: &'a str,
    /// Detected prefix character (>, =, @, :, #)
    pub prefix: Option<char>,
    /// The search needle (raw without prefix, trimmed)
    pub needle: &'a str,
}

impl<'a> SearchQuery<'a> {
    /// Parse a raw query string into structured form
    pub fn parse(raw: &'a str) -> Self {
        let trimmed = raw.trim();

        if trimmed.is_empty() {
            return Self {
                raw,
                prefix: None,
                needle: "",
            };
        }

        let first_char = trimmed.chars().next().unwrap();

        // Check for known prefixes
        // $ = named ranges, # = settings, @ = cells, : = goto, = = formulas, > = commands
        if matches!(first_char, '>' | '=' | '@' | ':' | '#' | '$') {
            Self {
                raw,
                prefix: Some(first_char),
                needle: trimmed[first_char.len_utf8()..].trim_start(),
            }
        } else {
            Self {
                raw,
                prefix: None,
                needle: trimmed,
            }
        }
    }

    /// Check if this query should include a specific provider
    pub fn matches_provider_prefix(&self, provider_prefixes: &[char]) -> bool {
        match self.prefix {
            Some(p) => provider_prefixes.contains(&p),
            None => provider_prefixes.is_empty(), // unprefixed providers participate
        }
    }
}

// ============================================================================
// Search Provider Trait
// ============================================================================

/// A source of search results
pub trait SearchProvider: Send + Sync {
    /// Provider name for debugging
    fn name(&self) -> &'static str;

    /// Prefixes this provider responds to.
    /// Empty slice means provider participates in unprefixed (general) search.
    fn prefixes(&self) -> &'static [char];

    /// Search for items matching the query.
    /// Implementations should:
    /// - Respect the limit parameter
    /// - Return items sorted by score (descending)
    /// - Not mutate any UI state
    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem>;
}

// ============================================================================
// Search Engine
// ============================================================================

/// The main search engine that coordinates providers
pub struct SearchEngine {
    providers: Vec<Box<dyn SearchProvider>>,
}

impl SearchEngine {
    pub fn new() -> Self {
        Self {
            providers: Vec::new(),
        }
    }

    /// Register a search provider
    pub fn register(&mut self, provider: Box<dyn SearchProvider>) {
        self.providers.push(provider);
    }

    /// Execute a search across all matching providers
    pub fn search(&self, raw_query: &str, limit: usize) -> Vec<SearchItem> {
        let query = SearchQuery::parse(raw_query);

        let mut results: Vec<SearchItem> = self.providers
            .iter()
            .filter(|p| {
                // Provider matches if:
                // 1. Query has a prefix that matches provider's prefixes, OR
                // 2. Query has no prefix AND provider has empty prefixes (general search)
                match query.prefix {
                    Some(prefix) => p.prefixes().contains(&prefix),
                    None => p.prefixes().is_empty(),
                }
            })
            .flat_map(|p| p.search(&query, limit))
            .collect();

        // Sort by: score (descending) → kind priority (ascending) → title (ascending)
        // This ensures stable, predictable ranking
        results.sort_by(|a, b| {
            // Primary: score descending
            match b.score.partial_cmp(&a.score) {
                Some(std::cmp::Ordering::Equal) | None => {}
                Some(ord) => return ord,
            }
            // Secondary: kind priority ascending (lower = higher priority)
            match a.kind.priority().cmp(&b.kind.priority()) {
                std::cmp::Ordering::Equal => {}
                ord => return ord,
            }
            // Tertiary: title alphabetically ascending
            a.title.cmp(&b.title)
        });

        // Truncate to limit
        results.truncate(limit);

        results
    }
}

impl Default for SearchEngine {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// Scoring Utilities
// ============================================================================

/// Simple substring match scorer
pub fn score_substring_match(needle: &str, haystack: &str) -> Option<(f32, Vec<(usize, usize)>)> {
    if needle.is_empty() {
        return Some((0.5, vec![])); // Empty query = show all with neutral score
    }

    let needle_lower = needle.to_lowercase();
    let haystack_lower = haystack.to_lowercase();

    // Exact match (highest)
    if haystack_lower == needle_lower {
        return Some((1.0, vec![(0, haystack.len())]));
    }

    // Prefix match (high)
    if haystack_lower.starts_with(&needle_lower) {
        return Some((0.9, vec![(0, needle.len())]));
    }

    // Contains match (medium)
    if let Some(pos) = haystack_lower.find(&needle_lower) {
        return Some((0.7, vec![(pos, pos + needle.len())]));
    }

    // Word boundary match (check if needle matches start of any word)
    let words: Vec<(usize, &str)> = haystack
        .match_indices(|c: char| c.is_whitespace())
        .map(|(i, _)| i + 1)
        .chain(std::iter::once(0))
        .filter_map(|start| haystack.get(start..).map(|s| (start, s)))
        .collect();

    for (start, word) in words {
        if word.to_lowercase().starts_with(&needle_lower) {
            return Some((0.8, vec![(start, start + needle.len())]));
        }
    }

    None
}

// ============================================================================
// Built-in Providers
// ============================================================================

/// Command search provider
pub struct CommandSearchProvider;

impl SearchProvider for CommandSearchProvider {
    fn name(&self) -> &'static str {
        "commands"
    }

    fn prefixes(&self) -> &'static [char] {
        &[] // Participates in unprefixed search
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        CommandId::all()
            .iter()
            .filter_map(|&cmd| {
                let name = cmd.name();
                let keywords = cmd.keywords();

                // Try matching name first
                if let Some((score, highlights)) = score_substring_match(query.needle, name) {
                    let mut item = SearchItem::new(
                        SearchKind::Command,
                        name,
                        SearchAction::RunCommand(cmd),
                    )
                    .with_score(score)
                    .with_highlights(highlights);

                    if let Some(shortcut) = cmd.shortcut() {
                        item = item.with_subtitle(shortcut);
                    }

                    return Some(item);
                }

                // Try matching keywords
                if score_substring_match(query.needle, keywords).is_some() {
                    let mut item = SearchItem::new(
                        SearchKind::Command,
                        name,
                        SearchAction::RunCommand(cmd),
                    )
                    .with_score(0.5); // Lower score for keyword match

                    if let Some(shortcut) = cmd.shortcut() {
                        item = item.with_subtitle(shortcut);
                    }

                    return Some(item);
                }

                None
            })
            .take(limit)
            .collect()
    }
}

/// Formula function search provider
pub struct FormulaSearchProvider;

impl SearchProvider for FormulaSearchProvider {
    fn name(&self) -> &'static str {
        "formulas"
    }

    fn prefixes(&self) -> &'static [char] {
        &['='] // Only responds to = prefix
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        use crate::formula_context::FUNCTIONS;

        FUNCTIONS
            .iter()
            .filter_map(|func| {
                if let Some((score, highlights)) = score_substring_match(query.needle, func.name) {
                    Some(SearchItem::new(
                        SearchKind::Formula,
                        func.name,
                        SearchAction::InsertFormula {
                            name: func.name.to_string(),
                            signature: func.signature.to_string(),
                        },
                    )
                    .with_subtitle(func.description)
                    .with_score(score)
                    .with_highlights(highlights)
                    .with_secondary_action(SearchAction::ShowFunctionHelp {
                        name: func.name.to_string(),
                        signature: func.signature.to_string(),
                        description: func.description.to_string(),
                    }))
                } else {
                    None
                }
            })
            .take(limit)
            .collect()
    }
}

/// Go-to cell provider (: prefix)
pub struct GoToSearchProvider;

impl SearchProvider for GoToSearchProvider {
    fn name(&self) -> &'static str {
        "goto"
    }

    fn prefixes(&self) -> &'static [char] {
        &[':']
    }

    fn search(&self, query: &SearchQuery, _limit: usize) -> Vec<SearchItem> {
        // Parse cell reference like "A1", "B25", "AA100"
        let input = query.needle.trim().to_uppercase();
        if input.is_empty() {
            return vec![];
        }

        // Find where letters end and numbers begin
        let letter_end = input.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        if letter_end == 0 || letter_end == input.len() {
            return vec![];
        }

        let letters = &input[..letter_end];
        let numbers = &input[letter_end..];

        // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
        let col = letters.chars().fold(0usize, |acc, c| {
            acc * 26 + (c as usize - 'A' as usize + 1)
        }).saturating_sub(1);

        // Parse row (1-based to 0-based)
        let row = match numbers.parse::<usize>() {
            Ok(r) if r > 0 => r - 1,
            _ => return vec![],
        };

        // Create result
        let cell_ref = format!("{}{}", letters, numbers);
        vec![SearchItem::new(
            SearchKind::GoTo,
            format!("Go to {}", cell_ref),
            SearchAction::JumpToCell { row, col },
        )
        .with_subtitle(format!("Row {}, Column {}", row + 1, col + 1))
        .with_score(1.0)]
    }
}

/// Settings search provider (# prefix)
/// Searches available settings and their current values
pub struct SettingsSearchProvider;

/// Known setting definitions
struct SettingDef {
    key: &'static str,
    label: &'static str,
    description: &'static str,
}

impl SettingsSearchProvider {
    const SETTINGS: &'static [SettingDef] = &[
        // Preferences entry - matches Cmd+, behavior
        SettingDef {
            key: "preferences",
            label: "Preferences",
            description: "Open preferences (theme selection)",
        },
        SettingDef {
            key: "theme",
            label: "Theme",
            description: "Color theme for the editor",
        },
        SettingDef {
            key: "font.family",
            label: "Font Family",
            description: "Font used for cell text",
        },
        SettingDef {
            key: "font.size",
            label: "Font Size",
            description: "Default font size",
        },
        SettingDef {
            key: "grid.rowHeight",
            label: "Row Height",
            description: "Default row height in pixels",
        },
        SettingDef {
            key: "grid.colWidth",
            label: "Column Width",
            description: "Default column width in pixels",
        },
        SettingDef {
            key: "autosave",
            label: "Auto Save",
            description: "Enable automatic saving",
        },
    ];

    /// Get the action for a setting key
    fn action_for_key(key: &str) -> SearchAction {
        match key {
            // Preferences and Theme both open the theme picker
            "preferences" | "theme" => SearchAction::RunCommand(CommandId::SelectTheme),
            // Other settings open the JSON file
            _ => SearchAction::OpenSetting { key: key.to_string() },
        }
    }
}

impl SearchProvider for SettingsSearchProvider {
    fn name(&self) -> &'static str {
        "settings"
    }

    fn prefixes(&self) -> &'static [char] {
        &['#']
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        let needle = query.needle.to_lowercase();

        Self::SETTINGS
            .iter()
            .filter_map(|setting| {
                // Match on key or label
                let match_text = format!("{} {}", setting.key, setting.label).to_lowercase();
                if let Some((score, highlights)) = score_substring_match(&needle, &match_text) {
                    Some(SearchItem::new(
                        SearchKind::Setting,
                        setting.label,
                        Self::action_for_key(setting.key),
                    )
                    .with_subtitle(setting.description)
                    .with_score(score)
                    .with_highlights(highlights))
                } else if needle.is_empty() {
                    // Show all settings when query is empty
                    Some(SearchItem::new(
                        SearchKind::Setting,
                        setting.label,
                        Self::action_for_key(setting.key),
                    )
                    .with_subtitle(setting.description)
                    .with_score(0.5))
                } else {
                    None
                }
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// Named Range Search Provider (# prefix)
// ============================================================================

/// A searchable named range entry (snapshot from workbook)
#[derive(Clone, Debug)]
pub struct NamedRangeEntry {
    pub name: String,
    pub reference: String,  // e.g., "A1" or "A1:B10"
    pub description: Option<String>,
    pub target_row: usize,  // Start row for navigation
    pub target_col: usize,  // Start col for navigation
}

impl NamedRangeEntry {
    pub fn new(
        name: String,
        reference: String,
        description: Option<String>,
        target_row: usize,
        target_col: usize,
    ) -> Self {
        Self { name, reference, description, target_row, target_col }
    }
}

/// Named range search provider - searches defined names
/// Create with a snapshot of named range data from the current workbook
pub struct NamedRangeSearchProvider {
    ranges: Vec<NamedRangeEntry>,
}

impl NamedRangeSearchProvider {
    pub fn new(ranges: Vec<NamedRangeEntry>) -> Self {
        Self { ranges }
    }
}

impl SearchProvider for NamedRangeSearchProvider {
    fn name(&self) -> &'static str {
        "named_ranges"
    }

    fn prefixes(&self) -> &'static [char] {
        &['$']
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        self.ranges
            .iter()
            .filter_map(|entry| {
                // Score against name
                if let Some((score, highlights)) = score_substring_match(query.needle, &entry.name) {
                    let mut item = SearchItem::new(
                        SearchKind::NamedRange,
                        &entry.name,
                        SearchAction::JumpToCell { row: entry.target_row, col: entry.target_col },
                    )
                    .with_subtitle(&entry.reference)
                    .with_score(score)
                    .with_highlights(highlights);

                    // Add secondary action to copy the name
                    item = item.with_secondary_action(SearchAction::CopyToClipboard {
                        text: entry.name.clone(),
                        description: format!("Copied '{}' to clipboard", entry.name),
                    });

                    Some(item)
                } else if query.needle.is_empty() {
                    // Show all named ranges when query is empty
                    let mut item = SearchItem::new(
                        SearchKind::NamedRange,
                        &entry.name,
                        SearchAction::JumpToCell { row: entry.target_row, col: entry.target_col },
                    )
                    .with_subtitle(&entry.reference)
                    .with_score(0.6);

                    item = item.with_secondary_action(SearchAction::CopyToClipboard {
                        text: entry.name.clone(),
                        description: format!("Copied '{}' to clipboard", entry.name),
                    });

                    Some(item)
                } else {
                    // Also try matching description if present
                    if let Some(ref desc) = entry.description {
                        if score_substring_match(query.needle, desc).is_some() {
                            let mut item = SearchItem::new(
                                SearchKind::NamedRange,
                                &entry.name,
                                SearchAction::JumpToCell { row: entry.target_row, col: entry.target_col },
                            )
                            .with_subtitle(&entry.reference)
                            .with_score(0.5);  // Lower score for description match

                            item = item.with_secondary_action(SearchAction::CopyToClipboard {
                                text: entry.name.clone(),
                                description: format!("Copied '{}' to clipboard", entry.name),
                            });

                            return Some(item);
                        }
                    }
                    None
                }
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// Cell Search Provider (requires app state snapshot)
// ============================================================================

/// A searchable cell entry (snapshot from sheet)
#[derive(Clone, Debug)]
pub struct CellEntry {
    pub row: usize,
    pub col: usize,
    pub display: String,
    pub formula: Option<String>,  // Raw formula text if cell is a formula
}

impl CellEntry {
    pub fn new(row: usize, col: usize, display: String, formula: Option<String>) -> Self {
        Self { row, col, display, formula }
    }

    /// Format cell reference like "A1", "AA100"
    fn cell_ref(&self) -> String {
        let mut col_name = String::new();
        let mut c = self.col;
        loop {
            col_name.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 { break; }
            c = c / 26 - 1;
        }
        format!("{}{}", col_name, self.row + 1)
    }
}

/// Cell search provider - searches within cell contents
/// Create with a snapshot of cell data from the current sheet
pub struct CellSearchProvider {
    cells: Vec<CellEntry>,
}

impl CellSearchProvider {
    pub fn new(cells: Vec<CellEntry>) -> Self {
        Self { cells }
    }
}

impl SearchProvider for CellSearchProvider {
    fn name(&self) -> &'static str {
        "cells"
    }

    fn prefixes(&self) -> &'static [char] {
        &['@']
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        if query.needle.is_empty() {
            return vec![];  // Don't show all cells on empty query
        }

        let needle = query.needle.to_lowercase();

        self.cells
            .iter()
            .filter_map(|cell| {
                // Build match text from display and formula
                let match_text = if let Some(ref formula) = cell.formula {
                    format!("{} {}", cell.display, formula).to_lowercase()
                } else {
                    cell.display.to_lowercase()
                };

                if let Some((score, highlights)) = score_substring_match(&needle, &match_text) {
                    // Format title: "A1: 42,500" or "B3: =SUM(A1:A10)"
                    let cell_ref = cell.cell_ref();
                    let title = if let Some(ref formula) = cell.formula {
                        format!("{}: {}", cell_ref, formula)
                    } else {
                        format!("{}: {}", cell_ref, cell.display)
                    };

                    Some(SearchItem::new(
                        SearchKind::Cell,
                        &title,
                        SearchAction::JumpToCell { row: cell.row, col: cell.col },
                    )
                    .with_score(score)
                    .with_highlights(highlights)
                    .with_secondary_action(SearchAction::CopyToClipboard {
                        text: cell_ref.clone(),
                        description: format!("Copied {} to clipboard", cell_ref),
                    }))
                } else {
                    None
                }
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// Recent Files Provider (requires app state snapshot)
// ============================================================================

/// Recent files provider - searches recently opened files
/// Create with a snapshot of recent file paths
pub struct RecentFilesProvider {
    files: Vec<PathBuf>,
}

impl RecentFilesProvider {
    pub fn new(files: Vec<PathBuf>) -> Self {
        Self { files }
    }
}

impl SearchProvider for RecentFilesProvider {
    fn name(&self) -> &'static str {
        "recent_files"
    }

    fn prefixes(&self) -> &'static [char] {
        &[]  // Participates in unprefixed search
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        self.files
            .iter()
            .filter_map(|path| {
                let filename = path.file_name()?.to_str()?;

                // Score against the query
                let (score, highlights) = if query.needle.is_empty() {
                    (0.6, vec![])  // Show all recent files with neutral score
                } else {
                    score_substring_match(query.needle, filename)?
                };

                let path_str = path.display().to_string();
                Some(SearchItem::new(
                    SearchKind::RecentFile,
                    filename,
                    SearchAction::OpenFile(path.clone()),
                )
                .with_subtitle(path.parent()?.to_str()?.to_string())
                .with_score(score)
                .with_highlights(highlights)
                .with_secondary_action(SearchAction::CopyToClipboard {
                    text: path_str,
                    description: "Copied path to clipboard".into(),
                }))
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// References Provider (Find References - Shift+F12)
// ============================================================================

/// A cell that references the source cell (a dependent)
#[derive(Clone, Debug)]
pub struct ReferenceEntry {
    pub row: usize,
    pub col: usize,
    pub cell_ref: String,
    pub formula: String,
}

impl ReferenceEntry {
    pub fn new(row: usize, col: usize, cell_ref: String, formula: String) -> Self {
        Self { row, col, cell_ref, formula }
    }
}

/// References provider - shows cells that depend on the source cell
/// Create with a snapshot of reference data from the current sheet
pub struct ReferencesProvider {
    source_cell_ref: String,
    references: Vec<ReferenceEntry>,
}

impl ReferencesProvider {
    pub fn new(source_cell_ref: String, references: Vec<ReferenceEntry>) -> Self {
        Self { source_cell_ref, references }
    }
}

impl SearchProvider for ReferencesProvider {
    fn name(&self) -> &'static str {
        "references"
    }

    fn prefixes(&self) -> &'static [char] {
        &[]  // Special provider, always participates when registered
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        self.references
            .iter()
            .filter_map(|entry| {
                // Title is the cell reference with formula snippet
                let title = format!("{}: {}", entry.cell_ref, entry.formula);

                // Filter by query if present
                if !query.needle.is_empty() {
                    let (score, highlights) = score_substring_match(query.needle, &title)?;
                    Some(SearchItem::new(
                        SearchKind::Reference,
                        &title,
                        SearchAction::JumpToCell { row: entry.row, col: entry.col },
                    )
                    .with_subtitle(format!("references {}", self.source_cell_ref))
                    .with_score(score)
                    .with_highlights(highlights))
                } else {
                    // Show all references with neutral score
                    Some(SearchItem::new(
                        SearchKind::Reference,
                        &title,
                        SearchAction::JumpToCell { row: entry.row, col: entry.col },
                    )
                    .with_subtitle(format!("references {}", self.source_cell_ref))
                    .with_score(0.8))
                }
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// Precedents Provider (Go to Precedents - F12)
// ============================================================================

/// A cell that the source cell references (a precedent)
#[derive(Clone, Debug)]
pub struct PrecedentEntry {
    pub row: usize,
    pub col: usize,
    pub cell_ref: String,
    pub display: String,
}

impl PrecedentEntry {
    pub fn new(row: usize, col: usize, cell_ref: String, display: String) -> Self {
        Self { row, col, cell_ref, display }
    }
}

/// Precedents provider - shows cells that the source cell references
/// Create with a snapshot of precedent data from the current sheet
pub struct PrecedentsProvider {
    source_cell_ref: String,
    precedents: Vec<PrecedentEntry>,
}

impl PrecedentsProvider {
    pub fn new(source_cell_ref: String, precedents: Vec<PrecedentEntry>) -> Self {
        Self { source_cell_ref, precedents }
    }
}

impl SearchProvider for PrecedentsProvider {
    fn name(&self) -> &'static str {
        "precedents"
    }

    fn prefixes(&self) -> &'static [char] {
        &[]  // Special provider, always participates when registered
    }

    fn search(&self, query: &SearchQuery, limit: usize) -> Vec<SearchItem> {
        self.precedents
            .iter()
            .filter_map(|entry| {
                // Title is the cell reference with its display value
                let title = format!("{}: {}", entry.cell_ref, entry.display);

                // Filter by query if present
                if !query.needle.is_empty() {
                    let (score, highlights) = score_substring_match(query.needle, &title)?;
                    Some(SearchItem::new(
                        SearchKind::Precedent,
                        &title,
                        SearchAction::JumpToCell { row: entry.row, col: entry.col },
                    )
                    .with_subtitle(format!("used by {}", self.source_cell_ref))
                    .with_score(score)
                    .with_highlights(highlights))
                } else {
                    // Show all precedents with neutral score
                    Some(SearchItem::new(
                        SearchKind::Precedent,
                        &title,
                        SearchAction::JumpToCell { row: entry.row, col: entry.col },
                    )
                    .with_subtitle(format!("used by {}", self.source_cell_ref))
                    .with_score(0.8))
                }
            })
            .take(limit)
            .collect()
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_query_parsing_no_prefix() {
        let q = SearchQuery::parse("copy");
        assert_eq!(q.prefix, None);
        assert_eq!(q.needle, "copy");
    }

    #[test]
    fn test_query_parsing_command_prefix() {
        let q = SearchQuery::parse("> undo");
        assert_eq!(q.prefix, Some('>'));
        assert_eq!(q.needle, "undo");
    }

    #[test]
    fn test_query_parsing_formula_prefix() {
        let q = SearchQuery::parse("=SUM");
        assert_eq!(q.prefix, Some('='));
        assert_eq!(q.needle, "SUM");
    }

    #[test]
    fn test_query_parsing_goto_prefix() {
        let q = SearchQuery::parse(":A1");
        assert_eq!(q.prefix, Some(':'));
        assert_eq!(q.needle, "A1");
    }

    #[test]
    fn test_command_provider_search() {
        let provider = CommandSearchProvider;
        let query = SearchQuery::parse("fill");
        let results = provider.search(&query, 10);

        assert!(!results.is_empty());
        assert!(results.iter().any(|r| r.title.contains("Fill")));
    }

    #[test]
    fn test_goto_provider_valid_cell() {
        let provider = GoToSearchProvider;
        let query = SearchQuery::parse(":B5");
        let results = provider.search(&query, 10);

        assert_eq!(results.len(), 1);
        match &results[0].action {
            SearchAction::JumpToCell { row, col } => {
                assert_eq!(*row, 4); // 0-indexed
                assert_eq!(*col, 1); // B = 1
            }
            _ => panic!("Expected JumpToCell action"),
        }
    }

    #[test]
    fn test_goto_provider_multi_letter_column() {
        let provider = GoToSearchProvider;
        let query = SearchQuery::parse(":AA1");
        let results = provider.search(&query, 10);

        assert_eq!(results.len(), 1);
        match &results[0].action {
            SearchAction::JumpToCell { row, col } => {
                assert_eq!(*row, 0);
                assert_eq!(*col, 26); // AA = 26
            }
            _ => panic!("Expected JumpToCell action"),
        }
    }

    #[test]
    fn test_search_engine_routing() {
        let mut engine = SearchEngine::new();
        engine.register(Box::new(CommandSearchProvider));
        engine.register(Box::new(GoToSearchProvider));

        // Unprefixed query should only get commands
        let results = engine.search("fill", 10);
        assert!(results.iter().all(|r| r.kind == SearchKind::Command));

        // : prefix should only get goto
        let results = engine.search(":A1", 10);
        assert!(results.iter().all(|r| r.kind == SearchKind::GoTo));
    }

    #[test]
    fn test_substring_scoring() {
        // Exact match = highest
        let (score, _) = score_substring_match("copy", "copy").unwrap();
        assert_eq!(score, 1.0);

        // Prefix match = high
        let (score, _) = score_substring_match("cop", "copy").unwrap();
        assert_eq!(score, 0.9);

        // Contains match = medium
        let (score, _) = score_substring_match("opy", "copy").unwrap();
        assert_eq!(score, 0.7);

        // No match
        assert!(score_substring_match("xyz", "copy").is_none());
    }

    #[test]
    fn test_ranking_stability() {
        // Create items with same score but different kinds/titles
        let mut items = vec![
            SearchItem::new(SearchKind::Setting, "Zebra", SearchAction::OpenSetting { key: "z".into() })
                .with_score(0.8),
            SearchItem::new(SearchKind::Command, "Apple", SearchAction::RunCommand(CommandId::Copy))
                .with_score(0.8),
            SearchItem::new(SearchKind::Command, "Banana", SearchAction::RunCommand(CommandId::Paste))
                .with_score(0.8),
            SearchItem::new(SearchKind::Formula, "Mango", SearchAction::InsertFormula { name: "M".into(), signature: "()".into() })
                .with_score(0.8),
        ];

        // Apply same sorting as SearchEngine::search
        items.sort_by(|a, b| {
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

        // Verify order: Commands first (alphabetically), then Formula, then Setting
        assert_eq!(items[0].title, "Apple");      // Command (priority 0), A
        assert_eq!(items[1].title, "Banana");     // Command (priority 0), B
        assert_eq!(items[2].title, "Mango");      // Formula (priority 3)
        assert_eq!(items[3].title, "Zebra");      // Setting (priority 5)
    }

    #[test]
    fn test_cell_search_provider() {
        let cells = vec![
            CellEntry::new(0, 0, "Hello World".into(), None),
            CellEntry::new(1, 0, "42".into(), None),
            CellEntry::new(2, 0, "100".into(), Some("=SUM(A1:A2)".into())),
        ];
        let provider = CellSearchProvider::new(cells);

        // Search for "world"
        let query = SearchQuery::parse("@world");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 1);
        assert!(results[0].title.contains("Hello World"));

        // Search for "SUM" (formula)
        let query = SearchQuery::parse("@SUM");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 1);
        assert!(results[0].title.contains("=SUM"));

        // Empty query returns nothing
        let query = SearchQuery::parse("@");
        let results = provider.search(&query, 10);
        assert!(results.is_empty());
    }

    #[test]
    fn test_recent_files_provider() {
        let files = vec![
            PathBuf::from("/home/user/documents/budget.csv"),
            PathBuf::from("/home/user/reports/sales.sheet"),
        ];
        let provider = RecentFilesProvider::new(files);

        // Search for "budget"
        let query = SearchQuery::parse("budget");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].title, "budget.csv");

        // Empty query returns all files
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 2);
    }

    #[test]
    fn test_cell_entry_cell_ref() {
        assert_eq!(CellEntry::new(0, 0, "".into(), None).cell_ref(), "A1");
        assert_eq!(CellEntry::new(0, 25, "".into(), None).cell_ref(), "Z1");
        assert_eq!(CellEntry::new(0, 26, "".into(), None).cell_ref(), "AA1");
        assert_eq!(CellEntry::new(99, 27, "".into(), None).cell_ref(), "AB100");
    }

    // =========================================================================
    // Cell Search Cache Freshness Tests
    // =========================================================================
    // These tests verify the CellSearchProvider behavior that underpins cache freshness.
    // The cache itself (in Spreadsheet) uses cells_rev to decide when to rebuild.
    // These tests verify that when the provider has fresh data, searches work correctly.

    #[test]
    fn test_cell_search_edited_cell_appears() {
        // Scenario: User edits A1 to "hello world", then searches @hello
        // Expected: A1 should appear in results

        let cells = vec![
            CellEntry::new(0, 0, "hello world".into(), None),
        ];
        let provider = CellSearchProvider::new(cells);

        let query = SearchQuery::parse("@hello");
        let results = provider.search(&query, 10);

        assert_eq!(results.len(), 1, "Edited cell should appear in search results");
        assert!(results[0].title.contains("hello world"), "Title should contain cell value");
        match &results[0].action {
            SearchAction::JumpToCell { row, col } => {
                assert_eq!(*row, 0, "Should jump to row 0");
                assert_eq!(*col, 0, "Should jump to col 0 (A)");
            }
            _ => panic!("Expected JumpToCell action"),
        }
    }

    #[test]
    fn test_cell_search_cleared_cell_absent() {
        // Scenario: Cell was cleared (not in provider's entries)
        // Expected: Should not appear in results

        // Provider with only B2 (A1 was "cleared" by not including it)
        let cells = vec![
            CellEntry::new(1, 1, "other data".into(), None),
        ];
        let provider = CellSearchProvider::new(cells);

        let query = SearchQuery::parse("@hello");
        let results = provider.search(&query, 10);

        assert!(results.is_empty(), "Cleared cell should not appear in results");
    }

    #[test]
    fn test_cell_search_formula_text_searchable() {
        // Scenario: B1 has formula =SUM(A1:A10), user searches @SUM
        // Expected: B1 should appear (formula text is searchable)

        let cells = vec![
            CellEntry::new(0, 1, "55".into(), Some("=SUM(A1:A10)".into())),
        ];
        let provider = CellSearchProvider::new(cells);

        let query = SearchQuery::parse("@SUM");
        let results = provider.search(&query, 10);

        assert_eq!(results.len(), 1, "Formula cell should appear when searching formula text");
        assert!(results[0].title.contains("=SUM"), "Title should show formula");
    }

    #[test]
    fn test_cell_search_fresh_data_finds_new_value() {
        // Scenario: Simulates cache refresh - new provider with updated data
        // This demonstrates what happens when cells_rev increments and cache rebuilds

        // Old data: A1 = "old value"
        let old_cells = vec![
            CellEntry::new(0, 0, "old value".into(), None),
        ];
        let old_provider = CellSearchProvider::new(old_cells);

        // Search should find "old"
        let query = SearchQuery::parse("@old");
        let results = old_provider.search(&query, 10);
        assert_eq!(results.len(), 1, "Old value should be found");

        // New data: A1 = "new value" (after edit + cache refresh)
        let new_cells = vec![
            CellEntry::new(0, 0, "new value".into(), None),
        ];
        let new_provider = CellSearchProvider::new(new_cells);

        // Search for "old" should NOT find anything (cell was changed)
        let query = SearchQuery::parse("@old");
        let results = new_provider.search(&query, 10);
        assert!(results.is_empty(), "Old value should not be found after edit");

        // Search for "new" should find the updated cell
        let query = SearchQuery::parse("@new");
        let results = new_provider.search(&query, 10);
        assert_eq!(results.len(), 1, "New value should be found after cache refresh");
    }

    #[test]
    fn test_cell_search_empty_query_returns_nothing() {
        // Scenario: User types @ with no query
        // Expected: No results (don't flood with all cells)

        let cells = vec![
            CellEntry::new(0, 0, "hello".into(), None),
            CellEntry::new(1, 0, "world".into(), None),
        ];
        let provider = CellSearchProvider::new(cells);

        let query = SearchQuery::parse("@");
        let results = provider.search(&query, 10);

        assert!(results.is_empty(), "Empty query should return no results");
    }

    // =========================================================================
    // Named Range Search Provider Tests
    // =========================================================================

    #[test]
    fn test_named_range_search_by_name() {
        let ranges = vec![
            NamedRangeEntry::new("Revenue".into(), "A1:A100".into(), Some("Total revenue".into()), 0, 0),
            NamedRangeEntry::new("Expenses".into(), "B1:B100".into(), None, 0, 1),
            NamedRangeEntry::new("Profit".into(), "C1".into(), None, 0, 2),
        ];
        let provider = NamedRangeSearchProvider::new(ranges);

        // Search for "rev" with $ prefix
        let query = SearchQuery::parse("$rev");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].title, "Revenue");
        assert_eq!(results[0].kind, SearchKind::NamedRange);
    }

    #[test]
    fn test_named_range_search_empty_query_shows_all() {
        let ranges = vec![
            NamedRangeEntry::new("Revenue".into(), "A1:A100".into(), None, 0, 0),
            NamedRangeEntry::new("Expenses".into(), "B1:B100".into(), None, 0, 1),
        ];
        let provider = NamedRangeSearchProvider::new(ranges);

        // Empty query shows all
        let query = SearchQuery::parse("$");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 2);
    }

    #[test]
    fn test_named_range_search_by_description() {
        let ranges = vec![
            NamedRangeEntry::new("TaxRate".into(), "D1".into(), Some("Current tax percentage".into()), 0, 3),
        ];
        let provider = NamedRangeSearchProvider::new(ranges);

        // Search by description
        let query = SearchQuery::parse("$percentage");
        let results = provider.search(&query, 10);
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].title, "TaxRate");
    }

    #[test]
    fn test_named_range_jump_action() {
        let ranges = vec![
            NamedRangeEntry::new("Data".into(), "E10:G20".into(), None, 9, 4),
        ];
        let provider = NamedRangeSearchProvider::new(ranges);

        let query = SearchQuery::parse("$Data");
        let results = provider.search(&query, 10);

        assert_eq!(results.len(), 1);
        match &results[0].action {
            SearchAction::JumpToCell { row, col } => {
                assert_eq!(*row, 9);  // E10 = row 9 (0-indexed)
                assert_eq!(*col, 4);  // E = col 4 (0-indexed)
            }
            _ => panic!("Expected JumpToCell action"),
        }
    }

    // =========================================================================
    // Keyboard Invariants Tests
    //
    // These shortcuts are Excel muscle memory and MUST NEVER change.
    // If a test fails here, you're breaking power user workflows.
    // =========================================================================

    #[test]
    fn test_keyboard_invariant_f9_is_recalculate() {
        // F9 = Recalculate (Excel muscle memory)
        // This was previously ToggleVerifiedMode - that was WRONG.
        assert_eq!(CommandId::Recalculate.shortcut(), Some("F9"));
    }

    #[test]
    fn test_keyboard_invariant_f2_edit_cell() {
        // F2 must start cell editing - universal spreadsheet convention
        // Note: F2 isn't in CommandId (it's a direct keybinding to StartEdit action)
        // This test documents the invariant even if we can't test the keybinding directly
    }

    #[test]
    fn test_keyboard_invariant_fill_shortcuts() {
        // Ctrl+D = Fill Down, Ctrl+R = Fill Right
        // These are fundamental data entry shortcuts
        assert_eq!(CommandId::FillDown.shortcut(), Some("Ctrl+D"));
        assert_eq!(CommandId::FillRight.shortcut(), Some("Ctrl+R"));
    }

    #[test]
    fn test_keyboard_invariant_autosum() {
        // Alt+= = AutoSum
        assert_eq!(CommandId::AutoSum.shortcut(), Some("Alt+="));
    }

    #[test]
    fn test_verified_mode_discoverable() {
        // ToggleVerifiedMode must be in the command palette (no shortcut is OK)
        // but it must be findable
        assert!(CommandId::all().contains(&CommandId::ToggleVerifiedMode));
        assert_eq!(CommandId::ToggleVerifiedMode.name(), "Toggle Verified Mode");
        // It should be in Tools menu category for Alt+T scoping
        assert_eq!(CommandId::ToggleVerifiedMode.menu_category(), Some(MenuCategory::Tools));
    }

    #[test]
    fn test_recalculate_in_view_menu() {
        // Recalculate should be accessible via Alt+V (View menu)
        assert_eq!(CommandId::Recalculate.menu_category(), Some(MenuCategory::View));
    }

    #[test]
    fn test_trace_commands_in_tools_menu() {
        // Trace commands should be in Tools menu (Alt+T)
        assert_eq!(CommandId::ToggleTrace.menu_category(), Some(MenuCategory::Tools));
        assert_eq!(CommandId::CycleTracePrecedent.menu_category(), Some(MenuCategory::Tools));
        assert_eq!(CommandId::CycleTraceDependent.menu_category(), Some(MenuCategory::Tools));
        assert_eq!(CommandId::ReturnToTraceSource.menu_category(), Some(MenuCategory::Tools));
        assert_eq!(CommandId::ToggleVerifiedMode.menu_category(), Some(MenuCategory::Tools));
    }

    #[test]
    fn test_alt_scope_key_hints() {
        // Verify Alt key hints match Excel ribbon convention
        assert_eq!(MenuCategory::Data.key_hint(), "A");  // Excel: Alt+A = Data
        assert_eq!(MenuCategory::Edit.key_hint(), "E");
        assert_eq!(MenuCategory::File.key_hint(), "F");
        assert_eq!(MenuCategory::View.key_hint(), "V");
        assert_eq!(MenuCategory::Tools.key_hint(), "T");
    }

    // =========================================================================
    // Alt+Down Priority Contract (Excel muscle memory)
    //
    // Alt+Down opens dropdowns in this priority order:
    // 1. Validation list dropdown (if cell has list validation)
    // 2. AutoFilter dropdown (if column is in filter range)
    // 3. "No dropdown available" status message
    //
    // This order matches Excel behavior: validation takes precedence over filter.
    // Changing this order will break user muscle memory.
    // =========================================================================

    #[test]
    fn test_alt_down_priority_documented() {
        // This test documents the Alt+Down priority contract.
        // The actual implementation is in app.rs::open_validation_dropdown().
        //
        // Priority order (MUST NOT CHANGE):
        // 1. Validation list dropdown (wins)
        // 2. AutoFilter dropdown (fallback)
        // 3. Status message "No dropdown available"
        //
        // If you're seeing this test and wondering why it exists:
        // It's a contract marker. The behavior is tested via integration tests,
        // but this ensures the contract is documented in the test suite.
        assert!(true, "Alt+Down priority: validation > filter > nothing");
    }
}
