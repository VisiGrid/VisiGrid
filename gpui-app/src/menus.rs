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
                MenuItem::action("Hide VisiGrid", HideApp),
                MenuItem::action("Hide Others", HideOthers),
                MenuItem::action("Show All", ShowAll),
                MenuItem::separator(),
                MenuItem::action("Quit VisiGrid", Quit),
            ],
            disabled: false,
        },
        // File menu
        Menu {
            name: "File".into(),
            items: vec![
                MenuItem::action("New Workbook", NewWindow),
                MenuItem::action("Open...", OpenFile),
                MenuItem::separator(),
                MenuItem::action("Save", Save),
                MenuItem::action("Save As...", SaveAs),
                MenuItem::separator(),
                MenuItem::action("Export as CSV...", ExportCsv),
                MenuItem::action("Export as TSV...", ExportTsv),
                MenuItem::action("Export as JSON...", ExportJson),
                MenuItem::action("Export to Excel (.xlsx)...", ExportXlsx),
                MenuItem::action("Export PDF...", ExportPdf),
                MenuItem::action("Print...", PrintPreview),
                MenuItem::separator(),
                MenuItem::action("Export Provenance Script (.lua)...", ExportProvenance),
                MenuItem::separator(),
                MenuItem::action("Close Window", CloseWindow),
            ],
            disabled: false,
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
                MenuItem::action("Paste Special...", PasteSpecial),
                MenuItem::separator(),
                MenuItem::action("Delete", DeleteCell),
                MenuItem::action("Select All", SelectAll),
                MenuItem::separator(),
                MenuItem::action("Edit Cell", StartEdit),
                MenuItem::action("Add/Edit Comment…", AddEditComment),
                MenuItem::separator(),
                MenuItem::action("Find...", FindInCells),
                MenuItem::action("Go To...", GoToCell),
            ],
            disabled: false,
        },
        // View menu
        Menu {
            name: "View".into(),
            items: vec![
                MenuItem::action("Command Palette", ToggleCommandPalette),
                MenuItem::separator(),
                MenuItem::action("Inspector", ToggleInspector),
                MenuItem::action("Comments", ToggleCommentsSidebar),
                MenuItem::separator(),
                MenuItem::action("Show Formulas", ToggleFormulaView),
                MenuItem::action("Show Zeros", ToggleShowZeros),
                MenuItem::action("Format Bar", ToggleFormatBar),
            ],
            disabled: false,
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
            disabled: false,
        },
        // Data menu
        Menu {
            name: "Data".into(),
            items: vec![
                MenuItem::action("Sort Ascending (A→Z)", SortAscending),
                MenuItem::action("Sort Descending (Z→A)", SortDescending),
                MenuItem::action("Clear Sort", ClearSort),
                MenuItem::separator(),
                MenuItem::action("Toggle AutoFilter", ToggleAutoFilter),
                MenuItem::separator(),
                MenuItem::action("Fill Down", FillDown),
                MenuItem::action("Fill Right", FillRight),
                MenuItem::separator(),
                MenuItem::action("Validation...", ShowDataValidation),
                MenuItem::separator(),
                MenuItem::action("Create Table...", CreateTable),
                MenuItem::action("Insert PivotTable...", InsertPivotTable),
                MenuItem::action("PivotTable Fields", EditPivotFields),
                MenuItem::action("Refresh PivotTable", RefreshPivot),
                MenuItem::action("Refresh All PivotTables", RefreshAllPivots),
                MenuItem::action("Delete PivotTable", DeletePivot),
            ],
            disabled: false,
        },
        // Window menu
        Menu {
            name: "Window".into(),
            items: vec![
                MenuItem::action("Minimize", Minimize),
                MenuItem::action("Zoom", Zoom),
                MenuItem::separator(),
                MenuItem::action("Switch Window...", SwitchWindow),
                MenuItem::separator(),
                MenuItem::action("Bring All to Front", BringAllToFront),
            ],
            disabled: false,
        },
        // Help menu
        Menu {
            name: "Help".into(),
            items: {
                let mut items = vec![
                    MenuItem::action("Keyboard Shortcuts...", ShowKeyTips),
                    MenuItem::separator(),
                    MenuItem::action("About VisiGrid", ShowAbout),
                ];
                // License entry is commercial-only: shipped OSS builds don't
                // consult a license, and a paste-a-key dialog in the MAS
                // build reads as out-of-app purchasing (guideline 3.1.1).
                if cfg!(feature = "commercial") {
                    items.push(MenuItem::action("License...", ShowLicense));
                }
                items
            },
            disabled: false,
        },
    ]);
}
