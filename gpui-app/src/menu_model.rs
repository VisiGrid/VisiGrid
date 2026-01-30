//! Menu model — typed descriptors for the in-app menu bar.
//!
//! This module is the single source of truth for menu structure, item counts,
//! and action dispatch. It lives outside `views/` so both `app.rs` (navigation)
//! and `views/menu_bar.rs` (rendering) can import it without a layering leak.

#![allow(dead_code)]

use gpui::*;
use crate::app::Spreadsheet;
use crate::formatting::BorderApplyMode;
use crate::mode::Menu;

/// Typed action enum for keyboard-navigable menu items.
/// Prevents usize drift between item count and execution.
#[derive(Clone, Copy)]
pub enum MenuAction {
    NewWorkbook, Open, Save, SaveAs,
    ExportCsv, ExportTsv, ExportJson, ExportXlsx,
    Undo, Redo, Cut, Copy, Paste, PasteValues, Delete, Find, GoTo,
    CommandPalette, Inspector, ZoomIn, ZoomOut, ZoomReset,
    ShowFormulas, ShowZeros, FormatBar,
    FreezeTopRow, FreezeFirstCol, FreezePanes, UnfreezePanes,
    Bold, Italic, Underline, Font,
    BgColor(Option<[u8; 4]>),
    BorderAll, BorderOutline, BorderClear,
    MergeCells, UnmergeCells,
    Validation, ExcludeValidation, ClearExclusions,
    FillDown, FillRight,
    CircleInvalid, ClearCircles,
    InsertFormulaAI, AnalyzeAI,
    About, License,
}

/// Menu entry descriptor — single source of truth for menu structure.
pub enum MenuEntry {
    Item { label: &'static str, shortcut: Option<&'static str>, action: MenuAction },
    Color { label: &'static str, color: Option<[u8; 4]>, action: MenuAction },
    Separator,
    Label(&'static str),
    Disabled(&'static str),
}

pub fn file_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "New Workbook", shortcut: Some("Ctrl+N"), action: MenuAction::NewWorkbook },
        MenuEntry::Item { label: "Open...", shortcut: Some("Ctrl+O"), action: MenuAction::Open },
        MenuEntry::Item { label: "Save", shortcut: Some("Ctrl+S"), action: MenuAction::Save },
        MenuEntry::Item { label: "Save As...", shortcut: Some("Ctrl+Shift+S"), action: MenuAction::SaveAs },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Export as CSV...", shortcut: None, action: MenuAction::ExportCsv },
        MenuEntry::Item { label: "Export as TSV...", shortcut: None, action: MenuAction::ExportTsv },
        MenuEntry::Item { label: "Export as JSON...", shortcut: None, action: MenuAction::ExportJson },
        MenuEntry::Item { label: "Export to Excel (.xlsx)...", shortcut: None, action: MenuAction::ExportXlsx },
    ]
}

pub fn edit_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Undo", shortcut: Some("Ctrl+Z"), action: MenuAction::Undo },
        MenuEntry::Item { label: "Redo", shortcut: Some("Ctrl+Y"), action: MenuAction::Redo },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Cut", shortcut: Some("Ctrl+X"), action: MenuAction::Cut },
        MenuEntry::Item { label: "Copy", shortcut: Some("Ctrl+C"), action: MenuAction::Copy },
        MenuEntry::Item { label: "Paste", shortcut: Some("Ctrl+V"), action: MenuAction::Paste },
        MenuEntry::Item { label: "Paste Values", shortcut: Some("Ctrl+Shift+V"), action: MenuAction::PasteValues },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Delete", shortcut: Some("Del"), action: MenuAction::Delete },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Find...", shortcut: Some("Ctrl+F"), action: MenuAction::Find },
        MenuEntry::Item { label: "Go To...", shortcut: Some("Ctrl+G"), action: MenuAction::GoTo },
    ]
}

pub fn view_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Command Palette", shortcut: Some("Ctrl+Shift+P"), action: MenuAction::CommandPalette },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Inspector", shortcut: Some("Ctrl+Shift+I"), action: MenuAction::Inspector },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Zoom In", shortcut: Some("Ctrl+Shift+="), action: MenuAction::ZoomIn },
        MenuEntry::Item { label: "Zoom Out", shortcut: Some("Ctrl+Shift+-"), action: MenuAction::ZoomOut },
        MenuEntry::Item { label: "Reset Zoom", shortcut: Some("Ctrl+0"), action: MenuAction::ZoomReset },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Show Formulas", shortcut: Some("Ctrl+`"), action: MenuAction::ShowFormulas },
        MenuEntry::Item { label: "Show Zeros", shortcut: None, action: MenuAction::ShowZeros },
        MenuEntry::Item { label: "Format Bar", shortcut: None, action: MenuAction::FormatBar },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Freeze Top Row", shortcut: None, action: MenuAction::FreezeTopRow },
        MenuEntry::Item { label: "Freeze First Column", shortcut: None, action: MenuAction::FreezeFirstCol },
        MenuEntry::Item { label: "Freeze Panes", shortcut: None, action: MenuAction::FreezePanes },
        MenuEntry::Item { label: "Unfreeze Panes", shortcut: None, action: MenuAction::UnfreezePanes },
    ]
}

pub fn insert_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Disabled("Rows"),
        MenuEntry::Disabled("Columns"),
        MenuEntry::Separator,
        MenuEntry::Disabled("Function..."),
    ]
}

pub fn format_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Bold", shortcut: Some("Ctrl+B"), action: MenuAction::Bold },
        MenuEntry::Item { label: "Italic", shortcut: Some("Ctrl+I"), action: MenuAction::Italic },
        MenuEntry::Item { label: "Underline", shortcut: Some("Ctrl+U"), action: MenuAction::Underline },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Font...", shortcut: None, action: MenuAction::Font },
        MenuEntry::Disabled("Cells..."),
        MenuEntry::Separator,
        MenuEntry::Label("Background Color"),
        MenuEntry::Color { label: "None", color: None, action: MenuAction::BgColor(None) },
        MenuEntry::Color { label: "Yellow", color: Some([255, 255, 0, 255]), action: MenuAction::BgColor(Some([255, 255, 0, 255])) },
        MenuEntry::Color { label: "Green", color: Some([198, 239, 206, 255]), action: MenuAction::BgColor(Some([198, 239, 206, 255])) },
        MenuEntry::Color { label: "Blue", color: Some([189, 215, 238, 255]), action: MenuAction::BgColor(Some([189, 215, 238, 255])) },
        MenuEntry::Color { label: "Red", color: Some([255, 199, 206, 255]), action: MenuAction::BgColor(Some([255, 199, 206, 255])) },
        MenuEntry::Color { label: "Orange", color: Some([255, 235, 156, 255]), action: MenuAction::BgColor(Some([255, 235, 156, 255])) },
        MenuEntry::Color { label: "Purple", color: Some([204, 192, 218, 255]), action: MenuAction::BgColor(Some([204, 192, 218, 255])) },
        MenuEntry::Color { label: "Gray", color: Some([217, 217, 217, 255]), action: MenuAction::BgColor(Some([217, 217, 217, 255])) },
        MenuEntry::Color { label: "Cyan", color: Some([183, 222, 232, 255]), action: MenuAction::BgColor(Some([183, 222, 232, 255])) },
        MenuEntry::Separator,
        MenuEntry::Label("Borders"),
        MenuEntry::Item { label: "All Borders", shortcut: None, action: MenuAction::BorderAll },
        MenuEntry::Item { label: "Outline", shortcut: None, action: MenuAction::BorderOutline },
        MenuEntry::Item { label: "Clear Borders", shortcut: None, action: MenuAction::BorderClear },
        MenuEntry::Separator,
        MenuEntry::Label("Merge"),
        MenuEntry::Item { label: "Merge Cells", shortcut: Some("Ctrl+Shift+M"), action: MenuAction::MergeCells },
        MenuEntry::Item { label: "Unmerge Cells", shortcut: Some("Ctrl+Shift+U"), action: MenuAction::UnmergeCells },
        MenuEntry::Separator,
        MenuEntry::Disabled("Row Height..."),
        MenuEntry::Disabled("Column Width..."),
    ]
}

pub fn data_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Validation...", shortcut: None, action: MenuAction::Validation },
        MenuEntry::Item { label: "Exclude from Validation", shortcut: None, action: MenuAction::ExcludeValidation },
        MenuEntry::Item { label: "Clear Validation Exclusions", shortcut: None, action: MenuAction::ClearExclusions },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Fill Down", shortcut: Some("Ctrl+D"), action: MenuAction::FillDown },
        MenuEntry::Item { label: "Fill Right", shortcut: Some("Ctrl+R"), action: MenuAction::FillRight },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Circle Invalid Data", shortcut: None, action: MenuAction::CircleInvalid },
        MenuEntry::Item { label: "Clear Invalid Circles", shortcut: None, action: MenuAction::ClearCircles },
        MenuEntry::Separator,
        MenuEntry::Disabled("Sort..."),
        MenuEntry::Disabled("Filter"),
        MenuEntry::Separator,
        MenuEntry::Item { label: "Insert Formula with AI", shortcut: Some("Ctrl+Shift+A"), action: MenuAction::InsertFormulaAI },
        MenuEntry::Item { label: "Analyze with AI", shortcut: Some("Ctrl+Shift+E"), action: MenuAction::AnalyzeAI },
    ]
}

pub fn help_menu_entries() -> Vec<MenuEntry> {
    let mut entries = vec![
        MenuEntry::Item { label: "About VisiGrid", shortcut: None, action: MenuAction::About },
        MenuEntry::Separator,
    ];
    entries.push(MenuEntry::Item {
        label: if visigrid_license::is_pro() { "Manage License..." } else { "Enter License..." },
        shortcut: None,
        action: MenuAction::License,
    });
    entries
}

pub fn menu_entries(menu: Menu) -> Vec<MenuEntry> {
    match menu {
        Menu::File => file_menu_entries(),
        Menu::Edit => edit_menu_entries(),
        Menu::View => view_menu_entries(),
        Menu::Insert => insert_menu_entries(),
        Menu::Format => format_menu_entries(),
        Menu::Data => data_menu_entries(),
        Menu::Help => help_menu_entries(),
    }
}

/// Count selectable items (Item + Color) in a menu.
pub fn menu_item_count(menu: Menu) -> usize {
    menu_entries(menu).iter().filter(|e| matches!(e, MenuEntry::Item { .. } | MenuEntry::Color { .. })).count()
}

/// Execute a menu action by selectable-item index.
pub fn execute_menu_action(app: &mut Spreadsheet, menu: Menu, index: usize, window: &mut Window, cx: &mut Context<Spreadsheet>) {
    let entries = menu_entries(menu);
    let mut selectable_idx = 0;
    for entry in &entries {
        match entry {
            MenuEntry::Item { action, .. } | MenuEntry::Color { action, .. } => {
                if selectable_idx == index {
                    dispatch_action(app, *action, window, cx);
                    return;
                }
                selectable_idx += 1;
            }
            _ => {}
        }
    }
}

fn dispatch_action(app: &mut Spreadsheet, action: MenuAction, window: &mut Window, cx: &mut Context<Spreadsheet>) {
    match action {
        MenuAction::NewWorkbook => cx.dispatch_action(&crate::actions::NewWindow),
        MenuAction::Open => app.open_file(cx),
        MenuAction::Save => app.save(cx),
        MenuAction::SaveAs => app.save_as(cx),
        MenuAction::ExportCsv => app.export_csv(cx),
        MenuAction::ExportTsv => app.export_tsv(cx),
        MenuAction::ExportJson => app.export_json(cx),
        MenuAction::ExportXlsx => app.export_xlsx(cx),
        MenuAction::Undo => app.undo(cx),
        MenuAction::Redo => app.redo(cx),
        MenuAction::Cut => app.cut(cx),
        MenuAction::Copy => app.copy(cx),
        MenuAction::Paste => app.paste(cx),
        MenuAction::PasteValues => app.paste_values(cx),
        MenuAction::Delete => app.delete_selection(cx),
        MenuAction::Find => app.show_find(cx),
        MenuAction::GoTo => app.show_goto(cx),
        MenuAction::CommandPalette => app.toggle_palette(cx),
        MenuAction::Inspector => { app.inspector_visible = !app.inspector_visible; cx.notify(); }
        MenuAction::ZoomIn => app.zoom_in(cx),
        MenuAction::ZoomOut => app.zoom_out(cx),
        MenuAction::ZoomReset => app.zoom_reset(cx),
        MenuAction::ShowFormulas => app.toggle_show_formulas(cx),
        MenuAction::ShowZeros => app.toggle_show_zeros(cx),
        MenuAction::FormatBar => app.toggle_format_bar(cx),
        MenuAction::FreezeTopRow => app.freeze_top_row(cx),
        MenuAction::FreezeFirstCol => app.freeze_first_column(cx),
        MenuAction::FreezePanes => app.freeze_panes(cx),
        MenuAction::UnfreezePanes => app.unfreeze_panes(cx),
        MenuAction::Bold => app.toggle_bold(cx),
        MenuAction::Italic => app.toggle_italic(cx),
        MenuAction::Underline => app.toggle_underline(cx),
        MenuAction::Font => app.show_font_picker(window, cx),
        MenuAction::BgColor(color) => app.set_background_color(color, cx),
        MenuAction::BorderAll => app.apply_borders(BorderApplyMode::All, cx),
        MenuAction::BorderOutline => app.apply_borders(BorderApplyMode::Outline, cx),
        MenuAction::BorderClear => app.apply_borders(BorderApplyMode::Clear, cx),
        MenuAction::MergeCells => app.merge_cells(cx),
        MenuAction::UnmergeCells => app.unmerge_cells(cx),
        MenuAction::Validation => app.show_validation_dialog(cx),
        MenuAction::ExcludeValidation => app.exclude_from_validation(cx),
        MenuAction::ClearExclusions => app.clear_validation_exclusions(cx),
        MenuAction::FillDown => app.fill_down(cx),
        MenuAction::FillRight => app.fill_right(cx),
        MenuAction::CircleInvalid => app.circle_invalid_data(cx),
        MenuAction::ClearCircles => app.clear_invalid_circles(cx),
        MenuAction::InsertFormulaAI => app.show_ask_ai(cx),
        MenuAction::AnalyzeAI => app.show_analyze(cx),
        MenuAction::About => app.show_about(cx),
        MenuAction::License => app.show_license(cx),
    }
}
