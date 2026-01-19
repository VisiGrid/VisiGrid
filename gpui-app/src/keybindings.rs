use gpui::{App, KeyBinding};
use crate::actions::*;

/// Register all keybindings for the application
pub fn register(cx: &mut App) {
    cx.bind_keys([
        // Navigation (in Spreadsheet context)
        KeyBinding::new("up", MoveUp, Some("Spreadsheet")),
        KeyBinding::new("down", MoveDown, Some("Spreadsheet")),
        KeyBinding::new("left", MoveLeft, Some("Spreadsheet")),
        KeyBinding::new("right", MoveRight, Some("Spreadsheet")),
        KeyBinding::new("ctrl-up", JumpUp, Some("Spreadsheet")),
        KeyBinding::new("ctrl-down", JumpDown, Some("Spreadsheet")),
        KeyBinding::new("ctrl-left", JumpLeft, Some("Spreadsheet")),
        KeyBinding::new("ctrl-right", JumpRight, Some("Spreadsheet")),
        KeyBinding::new("ctrl-home", MoveToStart, Some("Spreadsheet")),
        KeyBinding::new("ctrl-end", MoveToEnd, Some("Spreadsheet")),
        KeyBinding::new("pageup", PageUp, Some("Spreadsheet")),
        KeyBinding::new("pagedown", PageDown, Some("Spreadsheet")),
        KeyBinding::new("ctrl-g", GoToCell, Some("Spreadsheet")),
        KeyBinding::new("ctrl-f", FindInCells, Some("Spreadsheet")),
        KeyBinding::new("f3", FindNext, Some("Spreadsheet")),
        KeyBinding::new("shift-f3", FindPrev, Some("Spreadsheet")),

        // Editing
        KeyBinding::new("f2", StartEdit, Some("Spreadsheet")),
        KeyBinding::new("enter", ConfirmEdit, Some("Spreadsheet")),
        KeyBinding::new("ctrl-enter", ConfirmEditInPlace, Some("Spreadsheet")),
        KeyBinding::new("tab", TabNext, Some("Spreadsheet")),
        KeyBinding::new("shift-tab", TabPrev, Some("Spreadsheet")),
        KeyBinding::new("escape", CancelEdit, Some("Spreadsheet")),
        KeyBinding::new("backspace", BackspaceChar, Some("Spreadsheet")),
        KeyBinding::new("delete", DeleteChar, Some("Spreadsheet")),
        KeyBinding::new("delete", DeleteCell, Some("Spreadsheet")),
        KeyBinding::new("ctrl-d", FillDown, Some("Spreadsheet")),
        KeyBinding::new("ctrl-r", FillRight, Some("Spreadsheet")),
        // Edit mode cursor (Home/End only - left/right handled in MoveLeft/MoveRight)
        KeyBinding::new("home", EditCursorHome, Some("Spreadsheet")),
        KeyBinding::new("end", EditCursorEnd, Some("Spreadsheet")),
        // F4 reference cycling (Excel behavior)
        KeyBinding::new("f4", CycleReference, Some("Spreadsheet")),

        // Selection
        KeyBinding::new("ctrl-a", SelectAll, Some("Spreadsheet")),
        KeyBinding::new("shift-up", ExtendUp, Some("Spreadsheet")),
        KeyBinding::new("shift-down", ExtendDown, Some("Spreadsheet")),
        KeyBinding::new("shift-left", ExtendLeft, Some("Spreadsheet")),
        KeyBinding::new("shift-right", ExtendRight, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-up", ExtendJumpUp, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-down", ExtendJumpDown, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-left", ExtendJumpLeft, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-right", ExtendJumpRight, Some("Spreadsheet")),

        // Clipboard
        KeyBinding::new("ctrl-c", Copy, Some("Spreadsheet")),
        KeyBinding::new("ctrl-x", Cut, Some("Spreadsheet")),
        KeyBinding::new("ctrl-v", Paste, Some("Spreadsheet")),

        // File
        KeyBinding::new("ctrl-n", NewFile, Some("Spreadsheet")),
        KeyBinding::new("ctrl-o", OpenFile, Some("Spreadsheet")),
        KeyBinding::new("ctrl-s", Save, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-s", SaveAs, Some("Spreadsheet")),

        // View
        KeyBinding::new("ctrl-shift-p", ToggleCommandPalette, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-i", ToggleInspector, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-m", ToggleProblems, Some("Spreadsheet")),
        KeyBinding::new("f11", ToggleZenMode, Some("Spreadsheet")),

        // Format
        KeyBinding::new("ctrl-b", ToggleBold, Some("Spreadsheet")),
        KeyBinding::new("ctrl-i", ToggleItalic, Some("Spreadsheet")),
        KeyBinding::new("ctrl-u", ToggleUnderline, Some("Spreadsheet")),

        // History
        KeyBinding::new("ctrl-z", Undo, Some("Spreadsheet")),
        KeyBinding::new("ctrl-y", Redo, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-z", Redo, Some("Spreadsheet")),

        // Menu accelerators (Alt+letter, Excel 2003 style)
        KeyBinding::new("alt-f", OpenFileMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-e", OpenEditMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-v", OpenViewMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-i", OpenInsertMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-o", OpenFormatMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-d", OpenDataMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-h", OpenHelpMenu, Some("Spreadsheet")),

        // Sheet navigation
        KeyBinding::new("ctrl-pagedown", NextSheet, Some("Spreadsheet")),
        KeyBinding::new("ctrl-pageup", PrevSheet, Some("Spreadsheet")),
        KeyBinding::new("shift-f11", AddSheet, Some("Spreadsheet")),
    ]);
}
