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
    FindInCells,
    FindNext,
    FindPrev,
    FindReplace,    // Ctrl+H: Find and Replace dialog
    ReplaceNext,    // Replace current match and find next
    ReplaceAll,     // Replace all matches
    // IDE-style navigation (Shift+F12 / F12)
    FindReferences,    // Shift+F12: Find cells that reference current cell
    GoToPrecedents,    // F12: Go to cells that current cell references
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
]);

// File actions
actions!(file, [
    NewFile,
    OpenFile,
    Save,
    SaveAs,
    ExportCsv,
    ExportTsv,
    ExportJson,
    ExportXlsx,
    CloseWindow,
    Quit,
]);

// View actions
actions!(view, [
    ToggleCommandPalette,
    ToggleInspector,
    ShowFormatPanel,  // Ctrl+1 - opens inspector with Format tab
    ShowPreferences,  // Cmd+, on macOS - currently routes to theme picker
    OpenKeybindings,  // Open keybindings.json for editing
    ToggleProblems,
    ToggleZenMode,
    ToggleLuaConsole, // Ctrl+Shift+L - Lua scripting REPL
    ToggleFormulaView,
    ToggleShowZeros,
    ZoomIn,
    ZoomOut,
    ZoomReset,
    ShowAbout,
    ShowFontPicker,
    // Freeze panes
    FreezeTopRow,      // Freeze first row only
    FreezeFirstColumn, // Freeze first column only
    FreezePanes,       // Freeze rows/cols above and left of current cell
    UnfreezePanes,     // Clear all freeze panes
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

// Data actions (sort/filter)
actions!(data, [
    SortAscending,   // Sort current column ascending
    SortDescending,  // Sort current column descending
    ToggleAutoFilter, // Ctrl+Shift+F - enable/disable AutoFilter
    ClearSort,       // Remove current sort (restore original order)
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
actions!(accelerators, [
    AltFile,    // Alt+F - opens palette scoped to File commands
    AltEdit,    // Alt+E - opens palette scoped to Edit commands
    AltView,    // Alt+V - opens palette scoped to View commands
    AltFormat,  // Alt+O - opens palette scoped to Format commands (legacy Excel)
    AltData,    // Alt+D - opens palette scoped to Data commands
    AltHelp,    // Alt+H - opens palette scoped to Home/Format (modern Excel 2010+)
]);

// Default app prompt actions (macOS title bar chip)
actions!(default_app, [
    SetDefaultApp,        // Set VisiGrid as default for current file type
    DismissDefaultPrompt, // Dismiss the prompt (forever for this file type)
]);
