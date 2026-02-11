//! User-customizable keybindings
//!
//! Loads keybindings from `~/.config/visigrid/keybindings.json`
//! Users can override default shortcuts or add new ones.
//!
//! Format:
//! ```json
//! {
//!   "ctrl+shift+d": "selection.duplicate",
//!   "ctrl+;": "cell.insertDate"
//! }
//! ```

use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

use gpui::{App, KeyBinding};
use serde::{Deserialize, Serialize};

use crate::actions::*;

/// Get the path to the keybindings file
pub fn keybindings_path() -> Option<PathBuf> {
    dirs::config_dir().map(|p| p.join("visigrid").join("keybindings.json"))
}

/// User keybindings file format
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct UserKeybindings {
    /// Map of key combo -> action name
    #[serde(flatten)]
    pub bindings: HashMap<String, String>,
}

/// Load user keybindings from disk
pub fn load_user_keybindings() -> UserKeybindings {
    keybindings_path()
        .and_then(|p| fs::read_to_string(p).ok())
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

/// Save default keybindings template to disk (for user reference)
pub fn save_default_keybindings() {
    if let Some(path) = keybindings_path() {
        if path.exists() {
            return; // Don't overwrite existing file
        }
        if let Some(parent) = path.parent() {
            let _ = fs::create_dir_all(parent);
        }

        // Create OS-specific template
        #[cfg(target_os = "macos")]
        let template = r#"{
  "// VisiGrid Keybindings": "Add your custom shortcuts here",
  "// Format": "key-combo: action.name",
  "// Modifiers": "cmd, ctrl, alt, shift (use cmd for most shortcuts on macOS)",
  "//": "",
  "// Examples (uncomment to use):": "",
  "// cmd-shift-d": "edit.filldown",
  "// cmd-;": "edit.trim",
  "// alt-enter": "edit.confirminplace",
  "// cmd-shift-b": "selection.blanks",
  "//": "",
  "// Common actions:": "",
  "// navigation": "up, down, left, right, jumpup, goto, find",
  "// edit": "start, confirm, cancel, filldown, fillright, trim, undo, redo",
  "// selection": "all, blanks, row, column, extendup/down/left/right",
  "// clipboard": "copy, cut, paste",
  "// file": "new, open, save, saveas, exportcsv, exportjson",
  "// format": "bold, italic, underline",
  "// view": "palette, inspector, formulas"
}"#;

        #[cfg(target_os = "windows")]
        let template = r#"{
  "// VisiGrid Keybindings": "Add your custom shortcuts here",
  "// Format": "key-combo: action.name",
  "// Modifiers": "ctrl, alt, shift",
  "//": "",
  "// Examples (uncomment to use):": "",
  "// ctrl-shift-d": "edit.filldown",
  "// ctrl-;": "edit.trim",
  "// alt-enter": "edit.confirminplace",
  "// ctrl-shift-b": "selection.blanks",
  "//": "",
  "// Common actions:": "",
  "// navigation": "up, down, left, right, jumpup, goto, find",
  "// edit": "start, confirm, cancel, filldown, fillright, trim, undo, redo",
  "// selection": "all, blanks, row, column, extendup/down/left/right",
  "// clipboard": "copy, cut, paste",
  "// file": "new, open, save, saveas, exportcsv, exportjson",
  "// format": "bold, italic, underline",
  "// view": "palette, inspector, formulas"
}"#;

        #[cfg(target_os = "linux")]
        let template = r#"{
  "// VisiGrid Keybindings": "Add your custom shortcuts here",
  "// Format": "key-combo: action.name",
  "// Modifiers": "ctrl, alt, shift, super",
  "//": "",
  "// Examples (uncomment to use):": "",
  "// ctrl-shift-d": "edit.filldown",
  "// ctrl-;": "edit.trim",
  "// alt-enter": "edit.confirminplace",
  "// ctrl-shift-b": "selection.blanks",
  "//": "",
  "// Common actions:": "",
  "// navigation": "up, down, left, right, jumpup, goto, find",
  "// edit": "start, confirm, cancel, filldown, fillright, trim, undo, redo",
  "// selection": "all, blanks, row, column, extendup/down/left/right",
  "// clipboard": "copy, cut, paste",
  "// file": "new, open, save, saveas, exportcsv, exportjson",
  "// format": "bold, italic, underline",
  "// view": "palette, inspector, formulas"
}"#;

        let _ = fs::write(path, template);
    }
}

/// Open the keybindings file in the system editor
#[cfg(target_os = "linux")]
pub fn open_keybindings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = keybindings_path() {
        // Ensure file exists with template
        if !path.exists() {
            save_default_keybindings();
        }

        if let Ok(editor) = env::var("VISUAL").or_else(|_| env::var("EDITOR")) {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            std::process::Command::new("xdg-open").arg(&path).spawn()?;
        }
    }
    Ok(())
}

#[cfg(target_os = "macos")]
pub fn open_keybindings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = keybindings_path() {
        if !path.exists() {
            save_default_keybindings();
        }

        if let Ok(editor) = env::var("VISUAL").or_else(|_| env::var("EDITOR")) {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            std::process::Command::new("open").arg(&path).spawn()?;
        }
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub fn open_keybindings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = keybindings_path() {
        if !path.exists() {
            save_default_keybindings();
        }

        if let Ok(editor) = env::var("EDITOR") {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            std::process::Command::new("cmd")
                .args(["/C", "start", "", path.to_str().unwrap_or("")])
                .spawn()?;
        }
    }
    Ok(())
}

/// Register user keybindings (call after default keybindings)
/// User bindings take precedence over defaults.
pub fn register_user_keybindings(cx: &mut App) {
    let user_bindings = load_user_keybindings();

    let mut bindings: Vec<KeyBinding> = Vec::new();

    for (key_combo, action_name) in user_bindings.bindings {
        // Skip comment keys
        if key_combo.starts_with("//") {
            continue;
        }

        // Convert key combo format: "ctrl+shift+d" -> "ctrl-shift-d"
        let key = key_combo.replace('+', "-").to_lowercase();

        // Try to create a binding for this action
        if let Some(binding) = create_binding(&key, &action_name) {
            bindings.push(binding);
        }
    }

    if !bindings.is_empty() {
        cx.bind_keys(bindings);
    }
}

/// Create a KeyBinding from a key combo and action name
fn create_binding(key: &str, action: &str) -> Option<KeyBinding> {
    let context = Some("Spreadsheet");

    // Map action names to actual actions
    // Format: "namespace.action" or just "action"
    match action.to_lowercase().as_str() {
        // Navigation
        "navigation.up" | "move.up" => Some(KeyBinding::new(key, MoveUp, context)),
        "navigation.down" | "move.down" => Some(KeyBinding::new(key, MoveDown, context)),
        "navigation.left" | "move.left" => Some(KeyBinding::new(key, MoveLeft, context)),
        "navigation.right" | "move.right" => Some(KeyBinding::new(key, MoveRight, context)),
        "navigation.jumpup" | "jump.up" => Some(KeyBinding::new(key, JumpUp, context)),
        "navigation.jumpdown" | "jump.down" => Some(KeyBinding::new(key, JumpDown, context)),
        "navigation.jumpleft" | "jump.left" => Some(KeyBinding::new(key, JumpLeft, context)),
        "navigation.jumpright" | "jump.right" => Some(KeyBinding::new(key, JumpRight, context)),
        "navigation.start" | "goto.start" => Some(KeyBinding::new(key, MoveToStart, context)),
        "navigation.end" | "goto.end" => Some(KeyBinding::new(key, MoveToEnd, context)),
        "navigation.pageup" | "page.up" => Some(KeyBinding::new(key, PageUp, context)),
        "navigation.pagedown" | "page.down" => Some(KeyBinding::new(key, PageDown, context)),
        "navigation.goto" | "goto.cell" => Some(KeyBinding::new(key, GoToCell, context)),
        "navigation.find" | "find.cells" => Some(KeyBinding::new(key, FindInCells, context)),
        "navigation.findnext" | "find.next" => Some(KeyBinding::new(key, FindNext, context)),
        "navigation.findprev" | "find.prev" => Some(KeyBinding::new(key, FindPrev, context)),
        "navigation.references" | "find.references" => Some(KeyBinding::new(key, FindReferences, context)),
        "navigation.precedents" | "goto.precedents" => Some(KeyBinding::new(key, GoToPrecedents, context)),

        // Editing
        "edit.start" | "cell.edit" => Some(KeyBinding::new(key, StartEdit, context)),
        "edit.confirm" | "cell.confirm" => Some(KeyBinding::new(key, ConfirmEdit, context)),
        "edit.confirminplace" => Some(KeyBinding::new(key, ConfirmEditInPlace, context)),
        "edit.cancel" | "cell.cancel" => Some(KeyBinding::new(key, CancelEdit, context)),
        "edit.tabnext" | "cell.tabnext" => Some(KeyBinding::new(key, TabNext, context)),
        "edit.tabprev" | "cell.tabprev" => Some(KeyBinding::new(key, TabPrev, context)),
        "edit.delete" | "cell.delete" => Some(KeyBinding::new(key, DeleteCell, context)),
        "edit.deletechar" => Some(KeyBinding::new(key, DeleteChar, context)),
        "edit.backspace" => Some(KeyBinding::new(key, BackspaceChar, context)),
        "edit.filldown" | "fill.down" => Some(KeyBinding::new(key, FillDown, context)),
        "edit.fillright" | "fill.right" => Some(KeyBinding::new(key, FillRight, context)),
        "edit.insertrows" | "insert.rows" => Some(KeyBinding::new(key, InsertRowsOrCols, context)),
        "edit.deleterows" | "delete.rows" => Some(KeyBinding::new(key, DeleteRowsOrCols, context)),
        "edit.autosum" | "formula.autosum" => Some(KeyBinding::new(key, AutoSum, context)),
        "edit.cyclereference" | "formula.cycle" => Some(KeyBinding::new(key, CycleReference, context)),
        "edit.rename" | "refactor.rename" => Some(KeyBinding::new(key, RenameSymbol, context)),
        "edit.createnamedrange" | "range.create" => Some(KeyBinding::new(key, CreateNamedRange, context)),
        "edit.trim" | "transform.trim" => Some(KeyBinding::new(key, TrimWhitespace, context)),
        "transform.uppercase" | "transform.upper" => Some(KeyBinding::new(key, TransformUppercase, context)),
        "transform.lowercase" | "transform.lower" => Some(KeyBinding::new(key, TransformLowercase, context)),
        "transform.titlecase" | "transform.title" => Some(KeyBinding::new(key, TransformTitleCase, context)),
        "transform.sentencecase" | "transform.sentence" => Some(KeyBinding::new(key, TransformSentenceCase, context)),

        // Selection
        "selection.all" | "select.all" => Some(KeyBinding::new(key, SelectAll, context)),
        "selection.blanks" | "select.blanks" => Some(KeyBinding::new(key, SelectBlanks, context)),
        "selection.row" | "select.row" => Some(KeyBinding::new(key, SelectRow, context)),
        "selection.column" | "select.column" => Some(KeyBinding::new(key, SelectColumn, context)),
        "selection.extendup" | "extend.up" => Some(KeyBinding::new(key, ExtendUp, context)),
        "selection.extenddown" | "extend.down" => Some(KeyBinding::new(key, ExtendDown, context)),
        "selection.extendleft" | "extend.left" => Some(KeyBinding::new(key, ExtendLeft, context)),
        "selection.extendright" | "extend.right" => Some(KeyBinding::new(key, ExtendRight, context)),
        "selection.extendjumpup" | "extend.jumpup" => Some(KeyBinding::new(key, ExtendJumpUp, context)),
        "selection.extendjumpdown" | "extend.jumpdown" => Some(KeyBinding::new(key, ExtendJumpDown, context)),
        "selection.extendjumpleft" | "extend.jumpleft" => Some(KeyBinding::new(key, ExtendJumpLeft, context)),
        "selection.extendjumpright" | "extend.jumpright" => Some(KeyBinding::new(key, ExtendJumpRight, context)),

        // Clipboard
        "clipboard.copy" | "edit.copy" => Some(KeyBinding::new(key, Copy, context)),
        "clipboard.cut" | "edit.cut" => Some(KeyBinding::new(key, Cut, context)),
        "clipboard.paste" | "edit.paste" => Some(KeyBinding::new(key, Paste, context)),
        "clipboard.pastevalues" | "edit.pastevalues" => Some(KeyBinding::new(key, PasteValues, context)),

        // File
        "file.new" => Some(KeyBinding::new(key, NewWindow, context)),
        "file.open" => Some(KeyBinding::new(key, OpenFile, context)),
        "file.save" => Some(KeyBinding::new(key, Save, context)),
        "file.saveas" => Some(KeyBinding::new(key, SaveAs, context)),
        "file.exportcsv" | "export.csv" => Some(KeyBinding::new(key, ExportCsv, context)),
        "file.exporttsv" | "export.tsv" => Some(KeyBinding::new(key, ExportTsv, context)),
        "file.exportjson" | "export.json" => Some(KeyBinding::new(key, ExportJson, context)),
        "file.close" | "window.close" => Some(KeyBinding::new(key, CloseWindow, context)),
        "file.quit" | "app.quit" => Some(KeyBinding::new(key, Quit, context)),

        // View
        "view.palette" | "palette.toggle" => Some(KeyBinding::new(key, ToggleCommandPalette, context)),
        "view.inspector" | "inspector.toggle" => Some(KeyBinding::new(key, ToggleInspector, context)),
        "view.format" | "format.panel" => Some(KeyBinding::new(key, ShowFormatPanel, context)),
        "view.problems" | "problems.toggle" => Some(KeyBinding::new(key, ToggleProblems, context)),
        "view.zen" | "zen.toggle" => Some(KeyBinding::new(key, ToggleZenMode, context)),
        "view.formulas" | "formulas.toggle" => Some(KeyBinding::new(key, ToggleFormulaView, context)),
        "view.zeros" | "zeros.toggle" => Some(KeyBinding::new(key, ToggleShowZeros, context)),
        "view.preferences" | "preferences.show" => Some(KeyBinding::new(key, ShowPreferences, context)),
        "view.about" | "about.show" => Some(KeyBinding::new(key, ShowAbout, context)),
        "view.font" | "font.picker" => Some(KeyBinding::new(key, ShowFontPicker, context)),

        // Format
        "format.bold" | "style.bold" => Some(KeyBinding::new(key, ToggleBold, context)),
        "format.italic" | "style.italic" => Some(KeyBinding::new(key, ToggleItalic, context)),
        "format.underline" | "style.underline" => Some(KeyBinding::new(key, ToggleUnderline, context)),
        "format.alignleft" | "align.left" => Some(KeyBinding::new(key, AlignLeft, context)),
        "format.aligncenter" | "align.center" => Some(KeyBinding::new(key, AlignCenter, context)),
        "format.alignright" | "align.right" => Some(KeyBinding::new(key, AlignRight, context)),
        "format.currency" => Some(KeyBinding::new(key, FormatCurrency, context)),
        "format.percent" => Some(KeyBinding::new(key, FormatPercent, context)),

        // History
        "history.undo" | "edit.undo" => Some(KeyBinding::new(key, Undo, context)),
        "history.redo" | "edit.redo" => Some(KeyBinding::new(key, Redo, context)),

        // Sheets
        "sheet.next" | "sheets.next" => Some(KeyBinding::new(key, NextSheet, context)),
        "sheet.prev" | "sheets.prev" => Some(KeyBinding::new(key, PrevSheet, context)),
        "sheet.add" | "sheets.add" => Some(KeyBinding::new(key, AddSheet, context)),

        // Menu accelerators
        "menu.file" => Some(KeyBinding::new(key, OpenFileMenu, context)),
        "menu.edit" => Some(KeyBinding::new(key, OpenEditMenu, context)),
        "menu.view" => Some(KeyBinding::new(key, OpenViewMenu, context)),
        "menu.insert" => Some(KeyBinding::new(key, OpenInsertMenu, context)),
        "menu.format" => Some(KeyBinding::new(key, OpenFormatMenu, context)),
        "menu.data" => Some(KeyBinding::new(key, OpenDataMenu, context)),
        "menu.help" => Some(KeyBinding::new(key, OpenHelpMenu, context)),

        _ => None,
    }
}

/// Get all available action names (for documentation/autocomplete)
pub fn available_actions() -> Vec<&'static str> {
    vec![
        // Navigation
        "navigation.up", "navigation.down", "navigation.left", "navigation.right",
        "navigation.jumpup", "navigation.jumpdown", "navigation.jumpleft", "navigation.jumpright",
        "navigation.start", "navigation.end", "navigation.pageup", "navigation.pagedown",
        "navigation.goto", "navigation.find", "navigation.findnext", "navigation.findprev",
        "navigation.references", "navigation.precedents",
        // Editing
        "edit.start", "edit.confirm", "edit.confirminplace", "edit.cancel",
        "edit.delete", "edit.filldown", "edit.fillright", "edit.autosum",
        "edit.rename", "edit.createnamedrange", "edit.trim",
        // Selection
        "selection.all", "selection.blanks", "selection.row", "selection.column",
        "selection.extendup", "selection.extenddown", "selection.extendleft", "selection.extendright",
        // Clipboard
        "clipboard.copy", "clipboard.cut", "clipboard.paste", "clipboard.pastevalues",
        // File
        "file.new", "file.open", "file.save", "file.saveas",
        "file.exportcsv", "file.exporttsv", "file.exportjson",
        // View
        "view.palette", "view.inspector", "view.format", "view.zen", "view.formulas",
        // Format
        "format.bold", "format.italic", "format.underline",
        // History
        "history.undo", "history.redo",
        // Sheets
        "sheet.next", "sheet.prev", "sheet.add",
    ]
}
