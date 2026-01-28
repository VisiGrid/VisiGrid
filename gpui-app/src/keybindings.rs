use gpui::{App, KeyBinding};
use crate::actions::*;
use crate::settings::ModifierStyle;

/// Get the primary modifier key string based on platform and user preference
/// On macOS: "platform" -> "cmd", "ctrl" -> "ctrl"
/// On Windows/Linux: always "ctrl"
fn primary_mod(style: ModifierStyle) -> &'static str {
    #[cfg(target_os = "macos")]
    {
        match style {
            ModifierStyle::Platform => "cmd",
            ModifierStyle::Ctrl => "ctrl",
        }
    }
    #[cfg(not(target_os = "macos"))]
    {
        let _ = style; // suppress unused warning
        "ctrl"
    }
}

/// Build a keybinding string with the primary modifier
fn kb(style: ModifierStyle, key: &str) -> String {
    format!("{}-{}", primary_mod(style), key)
}

/// Build a keybinding string with primary modifier + shift
fn kb_shift(style: ModifierStyle, key: &str) -> String {
    format!("{}-shift-{}", primary_mod(style), key)
}

/// Register all keybindings for the application
pub fn register(cx: &mut App, modifier_style: ModifierStyle) {
    let m = modifier_style;

    // Build keybinding strings based on modifier preference
    let mut bindings: Vec<KeyBinding> = vec![
        // Navigation (in Spreadsheet context)
        KeyBinding::new("up", MoveUp, Some("Spreadsheet")),
        KeyBinding::new("down", MoveDown, Some("Spreadsheet")),
        KeyBinding::new("left", MoveLeft, Some("Spreadsheet")),
        KeyBinding::new("right", MoveRight, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "up"), JumpUp, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "down"), JumpDown, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "left"), JumpLeft, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "right"), JumpRight, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "home"), MoveToStart, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "end"), MoveToEnd, Some("Spreadsheet")),
        KeyBinding::new("pageup", PageUp, Some("Spreadsheet")),
        KeyBinding::new("pagedown", PageDown, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "g"), GoToCell, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "backspace"), JumpToActiveCell, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "f"), FindInCells, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "h"), FindReplace, Some("Spreadsheet")),
        KeyBinding::new("f3", FindNext, Some("Spreadsheet")),
        KeyBinding::new("shift-f3", FindPrev, Some("Spreadsheet")),
        // IDE-style navigation
        KeyBinding::new("shift-f12", FindReferences, Some("Spreadsheet")),
        KeyBinding::new("f12", GoToPrecedents, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "r"), RenameSymbol, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "n"), CreateNamedRange, Some("Spreadsheet")),
        // Validation failure navigation
        KeyBinding::new("f8", NextInvalidCell, Some("Spreadsheet")),
        KeyBinding::new("shift-f8", PrevInvalidCell, Some("Spreadsheet")),

        // Editing
        KeyBinding::new("f2", StartEdit, Some("Spreadsheet")),
        KeyBinding::new("enter", ConfirmEdit, Some("Spreadsheet")),
        KeyBinding::new("shift-enter", ConfirmEditUp, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "enter"), ConfirmEditInPlace, Some("Spreadsheet")),
        KeyBinding::new("tab", TabNext, Some("Spreadsheet")),
        KeyBinding::new("shift-tab", TabPrev, Some("Spreadsheet")),
        KeyBinding::new("escape", CancelEdit, Some("Spreadsheet")),
        KeyBinding::new("backspace", BackspaceChar, Some("Spreadsheet")),
        KeyBinding::new("delete", DeleteChar, Some("Spreadsheet")),
        KeyBinding::new("delete", DeleteCell, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "d"), FillDown, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "r"), FillRight, Some("Spreadsheet")),
        // Insert/Delete rows/cols (context-sensitive)
        KeyBinding::new(&kb(m, "="), InsertRowsOrCols, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "-"), DeleteRowsOrCols, Some("Spreadsheet")),
        // Edit mode cursor (Home/End only - left/right handled in MoveLeft/MoveRight)
        KeyBinding::new("home", EditCursorHome, Some("Spreadsheet")),
        KeyBinding::new("end", EditCursorEnd, Some("Spreadsheet")),
        // Edit mode word navigation (Alt+Arrow on macOS, Ctrl+Arrow elsewhere)
        KeyBinding::new("alt-left", EditWordLeft, Some("Spreadsheet")),
        KeyBinding::new("alt-right", EditWordRight, Some("Spreadsheet")),
        KeyBinding::new("alt-shift-left", EditSelectWordLeft, Some("Spreadsheet")),
        KeyBinding::new("alt-shift-right", EditSelectWordRight, Some("Spreadsheet")),
        // F4 reference cycling (Excel behavior)
        KeyBinding::new("f4", CycleReference, Some("Spreadsheet")),
        // Alt+= AutoSum (Excel behavior)
        KeyBinding::new("alt-=", AutoSum, Some("Spreadsheet")),
        // Alt+Down - open validation dropdown (Excel behavior)
        KeyBinding::new("alt-down", OpenValidationDropdown, Some("Spreadsheet")),

        // AI
        KeyBinding::new(&kb_shift(m, "a"), AskAI, Some("Spreadsheet")),

        // Selection
        KeyBinding::new(&kb(m, "a"), SelectAll, Some("Spreadsheet")),
        KeyBinding::new("shift-space", SelectRow, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "space"), SelectColumn, Some("Spreadsheet")),
        KeyBinding::new("shift-up", ExtendUp, Some("Spreadsheet")),
        KeyBinding::new("shift-down", ExtendDown, Some("Spreadsheet")),
        KeyBinding::new("shift-left", ExtendLeft, Some("Spreadsheet")),
        KeyBinding::new("shift-right", ExtendRight, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "up"), ExtendJumpUp, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "down"), ExtendJumpDown, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "left"), ExtendJumpLeft, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "right"), ExtendJumpRight, Some("Spreadsheet")),

        // Clipboard
        KeyBinding::new(&kb(m, "c"), Copy, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "x"), Cut, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "v"), Paste, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "v"), PasteValues, Some("Spreadsheet")),

        // File
        // Note: NewWindow is handled at App level (main.rs) to open a new window
        // We still bind it here so the keybinding shows in help, but the action
        // propagates up to the App-level handler
        KeyBinding::new(&kb(m, "n"), NewWindow, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "o"), OpenFile, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "s"), Save, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "s"), SaveAs, Some("Spreadsheet")),

        // View
        KeyBinding::new(&kb_shift(m, "p"), ToggleCommandPalette, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "i"), ToggleInspector, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "y"), ShowHistoryPanel, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "1"), ShowFormatPanel, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "m"), ToggleProblems, Some("Spreadsheet")),
        KeyBinding::new("f11", ToggleZenMode, Some("Spreadsheet")),
        KeyBinding::new("f9", Recalculate, Some("Spreadsheet")),  // Excel: force recalculate
        KeyBinding::new(&kb_shift(m, "l"), ToggleLuaConsole, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "`"), ToggleFormulaView, Some("Spreadsheet")),

        // Borders (Excel: Ctrl+Shift+& = outline, Ctrl+Shift+_ = clear)
        KeyBinding::new(&kb_shift(m, "7"), BordersOutline, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "-"), BordersClear, Some("Spreadsheet")),

        // Split view
        KeyBinding::new(&kb(m, "\\"), SplitRight, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "\\"), CloseSplit, Some("Spreadsheet")),
        // Split pane focus: Cmd+] on macOS, Ctrl+\ on Windows/Linux (to avoid Ctrl+] Excel conflict)
        #[cfg(target_os = "macos")]
        KeyBinding::new("cmd-]", FocusOtherPane, Some("Spreadsheet")),
        #[cfg(not(target_os = "macos"))]
        KeyBinding::new("ctrl-`", FocusOtherPane, Some("Spreadsheet")),

        // Dependency tracing
        // Toggle: Alt+T (Option+T on macOS) - universal
        KeyBinding::new("alt-t", ToggleTrace, Some("Spreadsheet")),
        // Jump to precedent/dependent: Ctrl+[/] on Windows/Linux (Excel), Alt+[/] on macOS
        #[cfg(target_os = "macos")]
        KeyBinding::new("alt-[", CycleTracePrecedent, Some("Spreadsheet")),
        #[cfg(target_os = "macos")]
        KeyBinding::new("alt-]", CycleTraceDependent, Some("Spreadsheet")),
        #[cfg(not(target_os = "macos"))]
        KeyBinding::new("ctrl-[", CycleTracePrecedent, Some("Spreadsheet")),
        #[cfg(not(target_os = "macos"))]
        KeyBinding::new("ctrl-]", CycleTraceDependent, Some("Spreadsheet")),
        // Return to trace source: F5 on Windows/Linux (Excel), Alt+Enter on macOS
        #[cfg(target_os = "macos")]
        KeyBinding::new("alt-enter", ReturnToTraceSource, Some("Spreadsheet")),
        #[cfg(not(target_os = "macos"))]
        KeyBinding::new("f5", ReturnToTraceSource, Some("Spreadsheet")),

        // Format
        KeyBinding::new(&kb(m, "b"), ToggleBold, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "i"), ToggleItalic, Some("Spreadsheet")),

        // Number formats (Mod+Shift+4 = $, Mod+Shift+5 = %)
        KeyBinding::new(&kb_shift(m, "4"), FormatCurrency, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "5"), FormatPercent, Some("Spreadsheet")),

        // History
        KeyBinding::new(&kb(m, "z"), Undo, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "y"), Redo, Some("Spreadsheet")),
        KeyBinding::new(&kb_shift(m, "z"), Redo, Some("Spreadsheet")),

        // Note: Alt+letter menu accelerators are registered separately via
        // register_menu_accelerators() or register_alt_accelerators() based on setting

        // Sheet navigation
        KeyBinding::new(&kb(m, "pagedown"), NextSheet, Some("Spreadsheet")),
        KeyBinding::new(&kb(m, "pageup"), PrevSheet, Some("Spreadsheet")),
        KeyBinding::new("shift-f11", AddSheet, Some("Spreadsheet")),

        // Data operations (sort/filter)
        KeyBinding::new(&kb_shift(m, "f"), ToggleAutoFilter, Some("Spreadsheet")),

        // Command palette (in CommandPalette context)
        KeyBinding::new("up", PaletteUp, Some("CommandPalette")),
        KeyBinding::new("down", PaletteDown, Some("CommandPalette")),
        KeyBinding::new("enter", PaletteExecute, Some("CommandPalette")),
        KeyBinding::new("shift-enter", PalettePreview, Some("CommandPalette")),
        KeyBinding::new("escape", PaletteCancel, Some("CommandPalette")),

        // Find dialog (in FindDialog context)
        KeyBinding::new(&kb(m, "f"), FindInCells, Some("FindDialog")),
        KeyBinding::new(&kb(m, "h"), FindReplace, Some("FindDialog")),
        KeyBinding::new("enter", ReplaceNext, Some("FindDialog")),
        KeyBinding::new(&kb(m, "enter"), ReplaceAll, Some("FindDialog")),
        KeyBinding::new("shift-enter", FindPrev, Some("FindDialog")),
        KeyBinding::new("f3", FindNext, Some("FindDialog")),
        KeyBinding::new("shift-f3", FindPrev, Some("FindDialog")),
        KeyBinding::new("escape", CancelEdit, Some("FindDialog")),
    ];

    // Platform-specific bindings
    #[cfg(target_os = "macos")]
    {
        // macOS always uses Cmd for these system shortcuts regardless of preference
        bindings.push(KeyBinding::new("cmd-,", ShowPreferences, Some("Spreadsheet")));
        bindings.push(KeyBinding::new("cmd-w", CloseWindow, Some("Spreadsheet")));
        bindings.push(KeyBinding::new("cmd-m", Minimize, Some("Spreadsheet")));
        bindings.push(KeyBinding::new("cmd-q", Quit, Some("Spreadsheet")));
        // Window switcher: Cmd+` (backtick)
        bindings.push(KeyBinding::new("cmd-`", SwitchWindow, Some("Spreadsheet")));

        // Ctrl+U starts edit on Mac (F2 is often brightness)
        bindings.push(KeyBinding::new("ctrl-u", StartEdit, Some("Spreadsheet")));

        // Note: On Mac, the Delete key (above Return) sends "backspace"
        // BackspaceChar handler now handles both edit mode (character delete)
        // and navigation mode (clear cell contents), so no separate DeleteCell binding needed
    }

    #[cfg(not(target_os = "macos"))]
    {
        // Preferences shortcut
        bindings.push(KeyBinding::new(&kb(m, ","), ShowPreferences, Some("Spreadsheet")));

        // Quit shortcut (Ctrl+Q on Windows/Linux)
        bindings.push(KeyBinding::new("ctrl-q", Quit, Some("Spreadsheet")));

        // Window switcher: Ctrl+` (backtick)
        bindings.push(KeyBinding::new("ctrl-`", SwitchWindow, Some("Spreadsheet")));

        // Ctrl+U for underline on non-Mac
        bindings.push(KeyBinding::new(&kb(m, "u"), ToggleUnderline, Some("Spreadsheet")));
    }

    cx.bind_keys(bindings);
}

/// Register menu dropdown accelerators (Alt+letter opens dropdown menu)
///
/// This is the default behavior when Alt accelerators setting is disabled.
/// Opens Excel 2003-style dropdown menus from the menu bar.
pub fn register_menu_accelerators(cx: &mut App) {
    cx.bind_keys([
        KeyBinding::new("alt-f", OpenFileMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-e", OpenEditMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-v", OpenViewMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-i", OpenInsertMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-o", OpenFormatMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-d", OpenDataMenu, Some("Spreadsheet")),
        KeyBinding::new("alt-h", OpenHelpMenu, Some("Spreadsheet")),
    ]);
}

/// Register Alt accelerator keybindings (opt-in, macOS only)
///
/// IMPORTANT: Alt is never stateful. We only bind complete chords (alt-f),
/// never Alt keydown/keyup. This prevents ghost states and ensures
/// Option key works normally for character composition when disabled.
///
/// These keybindings open the Command Palette scoped to a menu category.
/// Key mappings match Excel ribbon tabs for muscle memory:
/// - Alt+A -> Data (Excel ribbon: Alt+A)
/// - Alt+E -> Edit
/// - Alt+F -> File
/// - Alt+V -> View
/// - Alt+T -> Tools (trace, explain, audit)
/// - Alt+O -> Format (legacy Excel 2003)
/// - Alt+H -> Help
#[cfg(target_os = "macos")]
pub fn register_alt_accelerators(cx: &mut App) {
    cx.bind_keys([
        KeyBinding::new("alt-a", AltData, Some("Spreadsheet")),    // Excel ribbon: A=Data
        KeyBinding::new("alt-e", AltEdit, Some("Spreadsheet")),
        KeyBinding::new("alt-f", AltFile, Some("Spreadsheet")),
        KeyBinding::new("alt-v", AltView, Some("Spreadsheet")),
        KeyBinding::new("alt-t", AltTools, Some("Spreadsheet")),   // Tools: trace, explain
        KeyBinding::new("alt-o", AltFormat, Some("Spreadsheet")),
        KeyBinding::new("alt-h", AltHelp, Some("Spreadsheet")),
    ]);
}

/// Stub for non-macOS platforms (Alt accelerators not needed - native Alt menus exist)
#[cfg(not(target_os = "macos"))]
pub fn register_alt_accelerators(_cx: &mut App) {
    // On Windows/Linux, native Alt menus exist, so this is a no-op
}
