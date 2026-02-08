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
    ShowFormulas, ShowZeros, FormatBar, Minimap,
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
    Item { label: &'static str, shortcut: Option<&'static str>, action: MenuAction, accel: Option<char> },
    Color { label: &'static str, color: Option<[u8; 4]>, action: MenuAction, accel: Option<char> },
    Separator,
    Label(&'static str),
    Disabled(&'static str),
}

pub fn file_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "New Workbook", shortcut: Some("Ctrl+N"), action: MenuAction::NewWorkbook, accel: None },
        MenuEntry::Item { label: "Open...", shortcut: Some("Ctrl+O"), action: MenuAction::Open, accel: None },
        MenuEntry::Item { label: "Save", shortcut: Some("Ctrl+S"), action: MenuAction::Save, accel: Some('s') },
        MenuEntry::Item { label: "Save As...", shortcut: Some("Ctrl+Shift+S"), action: MenuAction::SaveAs, accel: Some('a') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Export as CSV...", shortcut: None, action: MenuAction::ExportCsv, accel: Some('c') },
        MenuEntry::Item { label: "Export as TSV...", shortcut: None, action: MenuAction::ExportTsv, accel: Some('t') },
        MenuEntry::Item { label: "Export as JSON...", shortcut: None, action: MenuAction::ExportJson, accel: Some('j') },
        MenuEntry::Item { label: "Export to Excel (.xlsx)...", shortcut: None, action: MenuAction::ExportXlsx, accel: Some('x') },
    ]
}

pub fn edit_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Undo", shortcut: Some("Ctrl+Z"), action: MenuAction::Undo, accel: None },
        MenuEntry::Item { label: "Redo", shortcut: Some("Ctrl+Y"), action: MenuAction::Redo, accel: None },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Cut", shortcut: Some("Ctrl+X"), action: MenuAction::Cut, accel: Some('t') },
        MenuEntry::Item { label: "Copy", shortcut: Some("Ctrl+C"), action: MenuAction::Copy, accel: Some('c') },
        MenuEntry::Item { label: "Paste", shortcut: Some("Ctrl+V"), action: MenuAction::Paste, accel: Some('p') },
        MenuEntry::Item { label: "Paste Values", shortcut: Some("Ctrl+Shift+V"), action: MenuAction::PasteValues, accel: Some('v') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Delete", shortcut: Some("Del"), action: MenuAction::Delete, accel: None },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Find...", shortcut: Some("Ctrl+F"), action: MenuAction::Find, accel: None },
        MenuEntry::Item { label: "Go To...", shortcut: Some("Ctrl+G"), action: MenuAction::GoTo, accel: None },
    ]
}

pub fn view_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Command Palette", shortcut: Some("Ctrl+Shift+P"), action: MenuAction::CommandPalette, accel: None },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Inspector", shortcut: Some("Ctrl+Shift+I"), action: MenuAction::Inspector, accel: None },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Zoom In", shortcut: Some("Ctrl+Shift+="), action: MenuAction::ZoomIn, accel: Some('+') },
        MenuEntry::Item { label: "Zoom Out", shortcut: Some("Ctrl+Shift+-"), action: MenuAction::ZoomOut, accel: Some('-') },
        MenuEntry::Item { label: "Reset Zoom", shortcut: Some("Ctrl+0"), action: MenuAction::ZoomReset, accel: Some('0') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Show Formulas", shortcut: Some("Ctrl+`"), action: MenuAction::ShowFormulas, accel: Some('f') },
        MenuEntry::Item { label: "Show Zeros", shortcut: None, action: MenuAction::ShowZeros, accel: Some('z') },
        MenuEntry::Item { label: "Format Bar", shortcut: None, action: MenuAction::FormatBar, accel: Some('b') },
        MenuEntry::Item { label: "Minimap", shortcut: None, action: MenuAction::Minimap, accel: Some('m') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Freeze Top Row", shortcut: None, action: MenuAction::FreezeTopRow, accel: Some('t') },
        MenuEntry::Item { label: "Freeze First Column", shortcut: None, action: MenuAction::FreezeFirstCol, accel: Some('l') },
        MenuEntry::Item { label: "Freeze Panes", shortcut: None, action: MenuAction::FreezePanes, accel: Some('p') },
        MenuEntry::Item { label: "Unfreeze Panes", shortcut: None, action: MenuAction::UnfreezePanes, accel: None },
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
        MenuEntry::Item { label: "Bold", shortcut: Some("Ctrl+B"), action: MenuAction::Bold, accel: Some('b') },
        MenuEntry::Item { label: "Italic", shortcut: Some("Ctrl+I"), action: MenuAction::Italic, accel: Some('i') },
        MenuEntry::Item { label: "Underline", shortcut: Some("Ctrl+U"), action: MenuAction::Underline, accel: Some('u') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Font...", shortcut: None, action: MenuAction::Font, accel: Some('f') },
        MenuEntry::Disabled("Cells..."),
        MenuEntry::Separator,
        MenuEntry::Label("Background Color"),
        MenuEntry::Color { label: "None", color: None, action: MenuAction::BgColor(None), accel: Some('n') },
        MenuEntry::Color { label: "Yellow", color: Some([255, 255, 0, 255]), action: MenuAction::BgColor(Some([255, 255, 0, 255])), accel: Some('y') },
        MenuEntry::Color { label: "Green", color: Some([198, 239, 206, 255]), action: MenuAction::BgColor(Some([198, 239, 206, 255])), accel: Some('g') },
        MenuEntry::Color { label: "Blue", color: Some([189, 215, 238, 255]), action: MenuAction::BgColor(Some([189, 215, 238, 255])), accel: Some('l') },
        MenuEntry::Color { label: "Red", color: Some([255, 199, 206, 255]), action: MenuAction::BgColor(Some([255, 199, 206, 255])), accel: Some('r') },
        MenuEntry::Color { label: "Orange", color: Some([255, 235, 156, 255]), action: MenuAction::BgColor(Some([255, 235, 156, 255])), accel: Some('o') },
        MenuEntry::Color { label: "Purple", color: Some([204, 192, 218, 255]), action: MenuAction::BgColor(Some([204, 192, 218, 255])), accel: Some('p') },
        MenuEntry::Color { label: "Gray", color: Some([217, 217, 217, 255]), action: MenuAction::BgColor(Some([217, 217, 217, 255])), accel: Some('a') },
        MenuEntry::Color { label: "Cyan", color: Some([183, 222, 232, 255]), action: MenuAction::BgColor(Some([183, 222, 232, 255])), accel: Some('c') },
        MenuEntry::Separator,
        MenuEntry::Label("Borders"),
        MenuEntry::Item { label: "All Borders", shortcut: None, action: MenuAction::BorderAll, accel: Some('d') },
        MenuEntry::Item { label: "Outline", shortcut: None, action: MenuAction::BorderOutline, accel: Some('t') },
        MenuEntry::Item { label: "Clear Borders", shortcut: None, action: MenuAction::BorderClear, accel: Some('e') },
        MenuEntry::Separator,
        MenuEntry::Label("Merge"),
        MenuEntry::Item { label: "Merge Cells", shortcut: Some("Ctrl+Shift+M"), action: MenuAction::MergeCells, accel: Some('m') },
        MenuEntry::Item { label: "Unmerge Cells", shortcut: Some("Ctrl+Shift+U"), action: MenuAction::UnmergeCells, accel: Some('x') },
        MenuEntry::Separator,
        MenuEntry::Disabled("Row Height..."),
        MenuEntry::Disabled("Column Width..."),
    ]
}

pub fn data_menu_entries() -> Vec<MenuEntry> {
    vec![
        MenuEntry::Item { label: "Validation...", shortcut: None, action: MenuAction::Validation, accel: None },
        MenuEntry::Item { label: "Exclude from Validation", shortcut: None, action: MenuAction::ExcludeValidation, accel: None },
        MenuEntry::Item { label: "Clear Validation Exclusions", shortcut: None, action: MenuAction::ClearExclusions, accel: None },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Fill Down", shortcut: Some("Ctrl+D"), action: MenuAction::FillDown, accel: Some('d') },
        MenuEntry::Item { label: "Fill Right", shortcut: Some("Ctrl+R"), action: MenuAction::FillRight, accel: Some('r') },
        MenuEntry::Separator,
        MenuEntry::Item { label: "Circle Invalid Data", shortcut: None, action: MenuAction::CircleInvalid, accel: Some('i') },
        MenuEntry::Item { label: "Clear Invalid Circles", shortcut: None, action: MenuAction::ClearCircles, accel: Some('l') },
        MenuEntry::Separator,
        MenuEntry::Disabled("Sort..."),
        MenuEntry::Disabled("Filter"),
        MenuEntry::Separator,
        MenuEntry::Item { label: "Insert Formula with AI", shortcut: Some("Ctrl+Shift+A"), action: MenuAction::InsertFormulaAI, accel: Some('f') },
        MenuEntry::Item { label: "Analyze with AI", shortcut: Some("Ctrl+Shift+E"), action: MenuAction::AnalyzeAI, accel: Some('a') },
    ]
}

pub fn help_menu_entries() -> Vec<MenuEntry> {
    let mut entries = vec![
        MenuEntry::Item { label: "About VisiGrid", shortcut: None, action: MenuAction::About, accel: None },
        MenuEntry::Separator,
    ];
    entries.push(MenuEntry::Item {
        label: if visigrid_license::is_pro() { "Manage License..." } else { "Enter License..." },
        shortcut: None,
        action: MenuAction::License,
        accel: None,
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

/// Resolve the accelerator character for a menu entry.
/// Uses explicit `accel` if set, otherwise falls back to first letter of label.
pub fn resolve_accel(label: &str, accel: Option<char>) -> char {
    accel.unwrap_or_else(|| label.chars().next().unwrap_or(' ')).to_ascii_lowercase()
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
        MenuAction::Minimap => { app.minimap_visible = !app.minimap_visible; cx.notify(); }
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

/// Debug: verify no duplicate accelerator keys within a menu.
/// Call during development to catch accel collisions.
#[cfg(debug_assertions)]
pub fn debug_assert_unique_accels(menu: Menu) {
    let entries = menu_entries(menu);
    let mut seen = std::collections::HashMap::new();
    for entry in &entries {
        match entry {
            MenuEntry::Item { label, accel, .. } | MenuEntry::Color { label, accel, .. } => {
                let ch = resolve_accel(label, *accel);
                if let Some(prev_label) = seen.insert(ch, *label) {
                    debug_assert!(
                        false,
                        "Duplicate accel '{}' in {:?} menu: \"{}\" and \"{}\"",
                        ch, menu, prev_label, label
                    );
                }
            }
            _ => {}
        }
    }
}

/// Run accel uniqueness checks for all menus (call once at startup in debug builds).
#[cfg(debug_assertions)]
pub fn debug_assert_all_accels() {
    use crate::mode::Menu;
    for menu in [Menu::File, Menu::Edit, Menu::View, Menu::Format, Menu::Data, Menu::Help] {
        debug_assert_unique_accels(menu);
    }
}
