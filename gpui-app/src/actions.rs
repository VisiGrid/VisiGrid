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
    HideRows,            // Ctrl+9 - hide selected rows
    UnhideRows,          // Ctrl+Shift+9 - unhide rows adjacent to selection
    HideCols,            // Ctrl+0 - hide selected columns
    UnhideCols,          // Ctrl+Shift+0 - unhide columns adjacent to selection
    InsertDate,          // Ctrl+; - insert current date into active cell
    InsertTime,          // Ctrl+Shift+; - insert current time into active cell
    CopyFormulaAbove,    // Ctrl+' - copy formula from cell above
    CopyValueAbove,      // Ctrl+Shift+" - copy display value from cell above
    InsertNewline,       // Alt+Enter - insert newline in cell
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
    // Text transforms
    TransformUppercase,
    TransformLowercase,
    TransformTitleCase,
    TransformSentenceCase,
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
    ExtendToStart,  // Ctrl+Shift+Home - extend selection to A1
    ExtendToEnd,    // Ctrl+Shift+End - extend selection to last used cell
    SelectCurrentRegion,  // Ctrl+Shift+* - select contiguous data region around active cell
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
    QuickOpen,  // Ctrl+K / Cmd+K - Open palette scoped to recent files
    ToggleInspector,
    ShowHistoryPanel,  // Ctrl+H - opens inspector with History tab
    ShowFormatPanel,  // Ctrl+1 - opens inspector with Format tab
    ShowPreferences,  // Cmd+, on macOS - currently routes to theme picker
    OpenKeybindings,  // Open keybindings.json for editing
    ToggleProblems,
    ToggleZenMode,
    ToggleLuaConsole, // Alt+F11 - Lua scripting REPL (matches Excel VBA Editor)
    ToggleTerminal,   // Ctrl+` - PTY terminal panel
    ToggleFormatBar,
    ToggleFormulaView,
    ToggleMinimap,
    ToggleShowZeros,
    ToggleVerifiedMode, // Toggle verified/deterministic recalc mode
    ToggleProfiler,
    ProfileNextRecalc,
    ClearProfiler,
    Recalculate,        // F9: Force full recalculation (Excel muscle memory)
    ApproveModel,       // Approve current semantic state (capture fingerprint)
    ClearApproval,      // Clear the approved state
    NavPerfReport,      // Show navigation latency report (VISIGRID_PERF=nav)
    ZoomIn,
    ZoomOut,
    ZoomReset,
    OpenContextMenu,     // Shift+F10 - open cell context menu
    ShowAbout,
    ShowLicense,
    ShowFontPicker,
    ShowColorPicker,    // Fill Color picker modal
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
    // Debug overlays (dev only)
    ToggleDebugGridAlignment,  // Cmd+Alt+Shift+G - Toggle pixel-alignment debug lines
]);

// Window actions (macOS standard Window menu)
actions!(window, [
    Minimize,          // Cmd+M: Minimize current window
    Zoom,              // Toggle window zoom (maximize/restore)
    BringAllToFront,   // Bring all app windows to front
    SwitchWindow,      // Cmd+` / Ctrl+`: Open window switcher palette
    NextWindow,        // Ctrl+Tab: Cycle to next window
    PrevWindow,        // Ctrl+Shift+Tab: Cycle to previous window
]);

// Format actions
actions!(format, [
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    ToggleStrikethrough,
    AlignLeft,
    AlignCenter,
    AlignRight,
    FormatCurrency,
    FormatPercent,
    FormatDate,        // Ctrl+Shift+# - apply date format
    FormatNumber,      // Ctrl+Shift+! - apply number (comma) format
    FormatGeneral,     // Ctrl+Shift+~ - apply general format
    FormatScientific,  // Ctrl+Shift+^ - apply scientific format
    FormatTime,        // Ctrl+Shift+@ - apply time format
    ClearFormatting,  // Reset all format properties to default
    FormatPainter,            // Activate Format Painter mode (single-shot)
    FormatPainterLocked,      // Activate Format Painter in locked mode (stays active)
    CopyFormat,               // Copy format from active cell (Ctrl+Shift+C)
    PasteFormat,              // Paste format to selection (Ctrl+Shift+V)
    CancelFormatPainter,      // Cancel Format Painter mode
    OpenNumberFormatEditor, // Open number format editor (Ctrl+1 escalation)
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
    BordersInside,   // Apply borders only to internal edges
    BordersTop,      // Apply border to top edge of selection
    BordersBottom,   // Apply border to bottom edge of selection
    BordersLeft,     // Apply border to left edge of selection
    BordersRight,    // Apply border to right edge of selection
    BordersClear,    // Clear all borders from selected cells
    // Merge cells
    MergeCells,      // Merge selected cells into one
    UnmergeCells,    // Unmerge selected merged cells
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
    OpenDiffResults,         // Open newest diff-*.json as a Diff Results sheet
]);

// Terminal actions — consume keys in Terminal context to prevent bubbling to Spreadsheet
actions!(terminal, [
    TerminalInput,  // No-op action that absorbs keys when terminal is focused
]);

// AI actions
actions!(ai, [
    InsertFormula,       // Ctrl+Shift+A - Insert Formula with AI
    Analyze,             // Ctrl+Shift+E - Analyze with AI (read-only)
]);

// Script editor actions
actions!(scripting, [
    ToggleScriptView,    // Ctrl+Shift+E / Cmd+Shift+E - Open/close script editor
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

// Debugger actions (Lua debug session control)
actions!(debugger, [
    DebugStartOrContinue,       // F5: start if idle, continue if paused
    DebugStepIn,                // F11
    DebugStepOver,              // F10
    DebugStepOut,               // Shift+F11
    DebugStop,                  // Shift+F5
    DebugToggleBreakpoint,      // F9 (no-op until Phase 5 gutter)
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
