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
    // IDE-style navigation (Shift+F12 / F12)
    FindReferences,    // Shift+F12: Find cells that reference current cell
    GoToPrecedents,    // F12: Go to cells that current cell references
]);

// Editing actions
actions!(editing, [
    StartEdit,
    ConfirmEdit,
    ConfirmEditInPlace,  // Ctrl+Enter - confirms without moving
    CancelEdit,
    TabNext,
    TabPrev,
    DeleteChar,
    BackspaceChar,
    DeleteCell,
    FillDown,
    FillRight,
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
]);

// File actions
actions!(file, [
    NewFile,
    OpenFile,
    Save,
    SaveAs,
    ExportCsv,
]);

// View actions
actions!(view, [
    ToggleCommandPalette,
    ToggleInspector,
    ShowFormatPanel,  // Ctrl+1 - opens inspector with Format tab
    ToggleProblems,
    ToggleZenMode,
    ToggleFormulaView,
    ZoomIn,
    ZoomOut,
    ZoomReset,
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

// Command palette actions
actions!(palette, [
    PaletteUp,
    PaletteDown,
    PaletteExecute,
    PalettePreview,   // Shift+Enter - preview without closing
    PaletteCancel,    // Escape - cancel and restore
]);
