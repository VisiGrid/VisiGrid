use gpui::actions;

// Navigation actions
actions!(navigation, [
    MoveUp,
    MoveDown,
    MoveLeft,
    MoveRight,
    JumpUp,
    JumpDown,
    JumpLeft,
    JumpRight,
    MoveToStart,
    MoveToEnd,
    PageUp,
    PageDown,
    GoToCell,
    JumpToActiveCell,  // Ctrl+Backspace: Scroll view to show active cell
    FindInCells,
    FindNext,
    FindPrev,
    FindReplace,    // Ctrl+H: Find and Replace dialog
    ReplaceNext,    // Replace current match and find next
    ReplaceAll,     // Replace all matches
    // IDE-style navigation (Shift+F12 / F12)
    FindReferences,    // Shift+F12: Find cells that reference current cell
    GoToPrecedents,    // F12: Go to cells that current cell references
    // Validation failure navigation (F8 / Shift+F8)
    NextInvalidCell,   // F8: Jump to next invalid cell
    PrevInvalidCell,   // Shift+F8: Jump to previous invalid cell
]);

// Editing actions
actions!(editing, [
    StartEdit,
    ConfirmEdit,
    ConfirmEditInPlace,  // Ctrl+Enter - confirms without moving
    ConfirmEditUp,       // Shift+Enter - confirms and moves up
    CancelEdit,
    TabNext,
    TabPrev,
    DeleteChar,
    BackspaceChar,
    DeleteCell,
    FillDown,
    FillRight,
    TrimWhitespace,
    InsertRowsOrCols,    // Ctrl+= / Ctrl++ - insert rows/cols based on selection
    DeleteRowsOrCols,    // Ctrl+- - delete rows/cols based on selection
    // Edit mode cursor movement
    EditCursorLeft,
    EditCursorRight,
    EditCursorHome,
    EditCursorEnd,
    // Edit mode text selection (Shift+Arrow)
    EditSelectLeft,
    EditSelectRight,
    EditSelectHome,
    EditSelectEnd,
    // Edit mode word navigation (Ctrl+Arrow)
    EditWordLeft,
    EditWordRight,
    // Edit mode word selection (Ctrl+Shift+Arrow)
    EditSelectWordLeft,
    EditSelectWordRight,
    // F4 reference cycling
    CycleReference,
    // Alt+= AutoSum
    AutoSum,
    // IDE-style Rename Symbol
    RenameSymbol,
    // Create Named Range from selection
    CreateNamedRange,
]);

// Selection actions
actions!(selection, [
    SelectAll,
    SelectBlanks,   // Select blank cells in current selection
    SelectRow,      // Shift+Space - select entire row
    SelectColumn,   // Ctrl+Space - select entire column
    ExtendUp,
    ExtendDown,
    ExtendLeft,
    ExtendRight,
    ExtendJumpUp,
    ExtendJumpDown,
    ExtendJumpLeft,
    ExtendJumpRight,
]);

// Clipboard actions
actions!(clipboard, [
    Copy,
    Cut,
    Paste,
    PasteValues,
    PasteSpecial,    // Ctrl+Alt+V / Cmd+Option+V - opens Paste Special dialog
    PasteFormulas,   // Paste formulas only (with reference adjustment)
    PasteFormats,    // Paste formatting only (no values)
]);

// File actions
actions!(file, [
    NewWindow,         // Ctrl+N: Open a new window with blank workbook (safe)
    NewInPlace,        // Replace current workbook in-place (dangerous, not bound by default)
    OpenFile,
    Save,
    SaveAs,
    ExportCsv,
    ExportTsv,
    ExportJson,
    ExportXlsx,
    ExportProvenance,  // Phase 9A: Export history as Lua script
    CloseWindow,
    Quit,
]);

// View actions
actions!(view, [
    ToggleCommandPalette,
    ToggleInspector,
    ShowHistoryPanel,  // Ctrl+H - opens inspector with History tab
    ShowFormatPanel,  // Ctrl+1 - opens inspector with Format tab
    ShowPreferences,  // Cmd+, on macOS - currently routes to theme picker
    OpenKeybindings,  // Open keybindings.json for editing
    ToggleProblems,
    ToggleZenMode,
    ToggleLuaConsole, // Ctrl+Shift+L - Lua scripting REPL
    ToggleFormulaView,
    ToggleShowZeros,
    ToggleVerifiedMode, // Toggle verified/deterministic recalc mode
    Recalculate,        // F9: Force full recalculation (Excel muscle memory)
    ZoomIn,
    ZoomOut,
    ZoomReset,
    ShowAbout,
    ShowLicense,
    ShowFontPicker,
    ShowKeyTips,        // Alt+Space (macOS) - Show keyboard accelerator hints
    // Freeze panes
    FreezeTopRow,      // Freeze first row only
    FreezeFirstColumn, // Freeze first column only
    FreezePanes,       // Freeze rows/cols above and left of current cell
    UnfreezePanes,     // Clear all freeze panes
    // Split view
    SplitRight,        // Ctrl+\ - Split view horizontally (50/50)
    CloseSplit,        // Close split view, keep active pane
    FocusOtherPane,    // Ctrl+] - Focus the other pane when split
    // Dependency tracing
    ToggleTrace,       // Alt+T - Toggle trace mode (highlight precedents/dependents)
    CycleTracePrecedent,  // Alt+[ - Jump to next precedent (shift for reverse)
    CycleTraceDependent,  // Alt+] - Jump to next dependent (shift for reverse)
    ReturnToTraceSource,  // F5 (Win/Linux), Alt+Enter (macOS) - Jump back to trace source
]);

// Window actions (macOS standard Window menu)
actions!(window, [
    Minimize,          // Cmd+M: Minimize current window
    Zoom,              // Toggle window zoom (maximize/restore)
    BringAllToFront,   // Bring all app windows to front
    SwitchWindow,      // Cmd+` / Ctrl+`: Open window switcher palette
]);

// Format actions
actions!(format, [
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    AlignLeft,
    AlignCenter,
    AlignRight,
    FormatCurrency,
    FormatPercent,
    // Background colors
    ClearBackground,
    BackgroundYellow,
    BackgroundGreen,
    BackgroundBlue,
    BackgroundRed,
    BackgroundOrange,
    BackgroundPurple,
    BackgroundGray,
    BackgroundCyan,
    // Borders
    BordersAll,      // Apply borders to all edges of selected cells
    BordersOutline,  // Apply borders only to outer perimeter of selection
    BordersClear,    // Clear all borders from selected cells
]);

// Undo/Redo
actions!(history, [
    Undo,
    Redo,
]);

// Menu bar actions (Alt+letter accelerators)
actions!(menu, [
    OpenFileMenu,
    OpenEditMenu,
    OpenViewMenu,
    OpenInsertMenu,
    OpenFormatMenu,
    OpenDataMenu,
    OpenHelpMenu,
    CloseMenu,
]);

// Sheet navigation actions
actions!(sheets, [
    NextSheet,
    PrevSheet,
    AddSheet,
]);

// Data actions (sort/filter/validation)
actions!(data, [
    SortAscending,   // Sort current column ascending
    SortDescending,  // Sort current column descending
    ToggleAutoFilter, // Ctrl+Shift+F - enable/disable AutoFilter
    ClearSort,       // Remove current sort (restore original order)
    ShowDataValidation,  // Data → Validation... dialog
    OpenValidationDropdown,  // Alt+Down - open validation dropdown for current cell
    ExcludeFromValidation,   // Exclude selection from validation
    ClearValidationExclusions, // Clear exclusions in selection
]);

// AI actions
actions!(ai, [
    AskAI,              // Ctrl+Shift+A - Ask AI about selected data
]);

// Command palette actions
actions!(palette, [
    PaletteUp,
    PaletteDown,
    PaletteExecute,
    PalettePreview,   // Shift+Enter - preview without closing
    PaletteCancel,    // Escape - cancel and restore
]);

// Alt accelerator actions (open Command Palette scoped to menu)
// These are opt-in on macOS via settings. On Windows/Linux, native Alt menus exist.
// Key mappings match Excel ribbon tabs: A=Data, E=Edit, F=File, V=View, T=Tools
actions!(accelerators, [
    AltFile,    // Alt+F - opens palette scoped to File commands
    AltEdit,    // Alt+E - opens palette scoped to Edit commands
    AltView,    // Alt+V - opens palette scoped to View commands
    AltFormat,  // Alt+O - opens palette scoped to Format commands (legacy Excel)
    AltData,    // Alt+A - opens palette scoped to Data commands (Excel ribbon: Alt+A)
    AltTools,   // Alt+T - opens palette scoped to Tools (trace, explain, audit)
    AltHelp,    // Alt+H - opens palette scoped to Help
]);

// Default app prompt actions (macOS title bar chip)
actions!(default_app, [
    SetDefaultApp,        // Set VisiGrid as default for current file type
    DismissDefaultPrompt, // Dismiss the prompt (forever for this file type)
]);

// VisiHub sync actions
actions!(hub, [
    HubCheckStatus,       // Check sync status with VisiHub
    HubPull,              // Pull latest version from VisiHub (only if local is clean)
    HubOpenRemoteAsCopy,  // Open remote version as a new file (always safe)
    HubPublish,           // Publish local changes to VisiHub (explicit only, never automatic)
    HubLink,              // Link current file to a VisiHub dataset
    HubUnlink,            // Remove VisiHub link from current file
    HubSignIn,            // Sign in to VisiHub (opens browser)
    HubSignOut,           // Sign out from VisiHub
    HubDiagnostics,       // Show hub sync diagnostics (state, errors, etc.)
]);
