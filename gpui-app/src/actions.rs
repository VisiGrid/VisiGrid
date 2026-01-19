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
    ToggleProblems,
    ToggleZenMode,
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
