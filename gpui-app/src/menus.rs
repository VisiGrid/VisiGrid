//! Native macOS menu bar using GPUI's set_menus API
//!
//! This module is only compiled on macOS. It creates the native menu bar
//! that appears at the top of the screen, dispatching the same Actions
//! used by keybindings and toolbar buttons.

use gpui::{App, Menu, MenuItem};

use crate::actions::*;

/// Set up the native macOS application menu bar
pub fn set_app_menus(cx: &mut App) {
    cx.set_menus(vec![
        // App menu (VisiGrid)
        Menu {
            name: "VisiGrid".into(),
            items: vec![
                MenuItem::action("About VisiGrid", ShowAbout),
                MenuItem::separator(),
                MenuItem::action("Preferences...", ShowPreferences),
                MenuItem::separator(),
                // Note: Hide/Show items would go here if GPUI supports them
                MenuItem::action("Quit VisiGrid", Quit),
            ],
        },
        // File menu
        Menu {
            name: "File".into(),
            items: vec![
                MenuItem::action("New", NewFile),
                MenuItem::action("Open...", OpenFile),
                MenuItem::separator(),
                MenuItem::action("Save", Save),
                MenuItem::action("Save As...", SaveAs),
                MenuItem::separator(),
                MenuItem::action("Export as CSV...", ExportCsv),
                MenuItem::separator(),
                MenuItem::action("Close", CloseWindow),
            ],
        },
        // Edit menu
        Menu {
            name: "Edit".into(),
            items: vec![
                MenuItem::action("Undo", Undo),
                MenuItem::action("Redo", Redo),
                MenuItem::separator(),
                MenuItem::action("Cut", Cut),
                MenuItem::action("Copy", Copy),
                MenuItem::action("Paste", Paste),
                MenuItem::separator(),
                MenuItem::action("Delete", DeleteCell),
                MenuItem::action("Select All", SelectAll),
                MenuItem::separator(),
                MenuItem::action("Edit Cell", StartEdit),
                MenuItem::separator(),
                MenuItem::action("Find...", FindInCells),
                MenuItem::action("Go To...", GoToCell),
            ],
        },
        // View menu
        Menu {
            name: "View".into(),
            items: vec![
                MenuItem::action("Command Palette", ToggleCommandPalette),
                MenuItem::separator(),
                MenuItem::action("Inspector", ToggleInspector),
                MenuItem::separator(),
                MenuItem::action("Show Formulas", ToggleFormulaView),
                MenuItem::action("Show Zeros", ToggleShowZeros),
            ],
        },
        // Format menu
        Menu {
            name: "Format".into(),
            items: vec![
                MenuItem::action("Bold", ToggleBold),
                MenuItem::action("Italic", ToggleItalic),
                MenuItem::action("Underline", ToggleUnderline),
                MenuItem::separator(),
                MenuItem::action("Font...", ShowFontPicker),
            ],
        },
        // Data menu
        Menu {
            name: "Data".into(),
            items: vec![
                MenuItem::action("Fill Down", FillDown),
                MenuItem::action("Fill Right", FillRight),
            ],
        },
        // Help menu
        Menu {
            name: "Help".into(),
            items: vec![
                MenuItem::action("About VisiGrid", ShowAbout),
            ],
        },
    ]);
}
