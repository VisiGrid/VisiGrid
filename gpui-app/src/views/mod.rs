mod about_dialog;
pub mod command_palette;
mod export_report_dialog;
mod find_dialog;
mod font_picker;
mod formula_bar;
mod goto_dialog;
mod grid;
mod headers;
pub mod impact_preview;
mod import_overlay;
mod import_report_dialog;
pub mod inspector_panel;
#[cfg(feature = "pro")]
mod lua_console;
#[cfg(not(feature = "pro"))]
mod lua_console_stub;
#[cfg(not(feature = "pro"))]
use lua_console_stub as lua_console;
pub mod license_dialog;
mod preferences_panel;
pub mod refactor_log;
mod menu_bar;
mod status_bar;
mod theme_picker;
mod tour;

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CELL_HEIGHT, HEADER_WIDTH, CreateNameFocus};
use crate::actions::*;
use crate::mode::{Mode, InspectorTab};
use crate::theme::TokenKey;

pub fn render_spreadsheet(app: &mut Spreadsheet, window: &mut Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let editing = app.mode.is_editing();
    let show_goto = app.mode == Mode::GoTo;
    let show_find = app.mode == Mode::Find;
    let show_command = app.mode == Mode::Command;
    let show_font_picker = app.mode == Mode::FontPicker;
    let show_theme_picker = app.mode == Mode::ThemePicker;
    let show_about = app.mode == Mode::About;
    let show_rename_symbol = app.mode == Mode::RenameSymbol;
    let show_create_named_range = app.mode == Mode::CreateNamedRange;
    let show_edit_description = app.mode == Mode::EditDescription;
    let show_tour = app.mode == Mode::Tour;
    let show_impact_preview = app.mode == Mode::ImpactPreview;
    let show_refactor_log = app.mode == Mode::RefactorLog;
    let show_extract_named_range = app.mode == Mode::ExtractNamedRange;
    let show_import_report = app.mode == Mode::ImportReport;
    let show_export_report = app.mode == Mode::ExportReport;
    let show_preferences = app.mode == Mode::Preferences;
    let show_license = app.mode == Mode::License;
    let show_import_overlay = app.import_overlay_visible;
    let show_name_tooltip = app.should_show_name_tooltip(cx) && app.mode == Mode::Navigation;
    let show_f2_tip = app.should_show_f2_tip(cx);  // Show immediately on trigger, not gated on mode
    let show_inspector = app.inspector_visible;
    let zen_mode = app.zen_mode;

    div()
        .relative()
        .key_context("Spreadsheet")
        .track_focus(&app.focus_handle)
        // Navigation actions (formula mode: insert references, edit mode: move cursor, nav mode: move selection)
        .on_action(cx.listener(|this, _: &MoveUp, _, cx| {
            // Lua console: history prev
            if this.lua_console.visible {
                this.lua_console.history_prev();
                cx.notify();
                return;
            }
            // Autocomplete navigation takes priority
            if this.autocomplete_visible {
                this.autocomplete_up(cx);
                return;
            }
            match this.mode {
                Mode::Command => this.palette_up(cx),
                Mode::FontPicker => this.font_picker_up(cx),
                Mode::ThemePicker => this.theme_picker_up(cx),
                Mode::Formula => this.formula_move_ref(-1, 0, cx),
                _ => this.move_selection(-1, 0, cx),
            }
        }))
        .on_action(cx.listener(|this, _: &MoveDown, _, cx| {
            // Lua console: history next
            if this.lua_console.visible {
                this.lua_console.history_next();
                cx.notify();
                return;
            }
            // Autocomplete navigation takes priority
            if this.autocomplete_visible {
                this.autocomplete_down(cx);
                return;
            }
            match this.mode {
                Mode::Command => this.palette_down(cx),
                Mode::FontPicker => this.font_picker_down(cx),
                Mode::ThemePicker => this.theme_picker_down(cx),
                Mode::Formula => this.formula_move_ref(1, 0, cx),
                _ => this.move_selection(1, 0, cx),
            }
        }))
        .on_action(cx.listener(|this, _: &MoveLeft, _, cx| {
            // Lua console: cursor left
            if this.lua_console.visible {
                this.lua_console.cursor_left();
                cx.notify();
                return;
            }
            if this.mode.is_formula() {
                this.formula_move_ref(0, -1, cx);
            } else if this.mode.is_editing() {
                this.move_edit_cursor_left(cx);
            } else {
                this.move_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &MoveRight, _, cx| {
            // Lua console: cursor right
            if this.lua_console.visible {
                this.lua_console.cursor_right();
                cx.notify();
                return;
            }
            if this.mode.is_formula() {
                this.formula_move_ref(0, 1, cx);
            } else if this.mode.is_editing() {
                this.move_edit_cursor_right(cx);
            } else {
                this.move_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(-1, 0, cx);
            } else {
                this.jump_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(1, 0, cx);
            } else {
                this.jump_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpLeft, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_left(cx);
            } else {
                this.jump_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpRight, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_right(cx);
            } else {
                this.jump_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &MoveToStart, _, cx| {
            this.selected = (0, 0);
            this.scroll_row = 0;
            this.scroll_col = 0;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &MoveToEnd, _, cx| {
            this.selected = (crate::app::NUM_ROWS - 1, crate::app::NUM_COLS - 1);
            this.scroll_row = crate::app::NUM_ROWS.saturating_sub(this.visible_rows());
            this.scroll_col = crate::app::NUM_COLS.saturating_sub(this.visible_cols());
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &PageUp, _, cx| {
            this.page_up(cx);
        }))
        .on_action(cx.listener(|this, _: &PageDown, _, cx| {
            this.page_down(cx);
        }))
        // Selection extension (formula mode: extend range reference)
        .on_action(cx.listener(|this, _: &ExtendUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(-1, 0, cx);
            } else {
                this.extend_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(1, 0, cx);
            } else {
                this.extend_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendLeft, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_left(cx);
            } else {
                this.extend_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendRight, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_right(cx);
            } else {
                this.extend_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(-1, 0, cx);
            } else {
                this.extend_jump_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(1, 0, cx);
            } else {
                this.extend_jump_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpLeft, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_left(cx);
            } else {
                this.extend_jump_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpRight, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_right(cx);
            } else {
                this.extend_jump_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectAll, _, cx| {
            if this.mode == Mode::Edit {
                this.select_all_edit(cx);
            } else {
                this.select_all(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectBlanks, _, cx| {
            if !this.mode.is_editing() {
                this.select_blanks(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectRow, _, cx| {
            if !this.mode.is_editing() {
                let row = this.selected.0;
                this.select_row(row, false, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectColumn, _, cx| {
            if !this.mode.is_editing() {
                let col = this.selected.1;
                this.select_col(col, false, cx);
            }
        }))
        // File actions
        .on_action(cx.listener(|this, _: &NewFile, _, cx| {
            this.new_file(cx);
        }))
        .on_action(cx.listener(|this, _: &OpenFile, _, cx| {
            this.open_file(cx);
        }))
        .on_action(cx.listener(|this, _: &Save, _, cx| {
            this.save(cx);
        }))
        .on_action(cx.listener(|this, _: &SaveAs, _, cx| {
            this.save_as(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportCsv, _, cx| {
            this.export_csv(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportTsv, _, cx| {
            this.export_tsv(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportJson, _, cx| {
            this.export_json(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportXlsx, _, cx| {
            this.export_xlsx(cx);
        }))
        // Clipboard actions
        .on_action(cx.listener(|this, _: &Copy, _, cx| {
            this.copy(cx);
        }))
        .on_action(cx.listener(|this, _: &Cut, _, cx| {
            this.cut(cx);
        }))
        .on_action(cx.listener(|this, _: &Paste, _, cx| {
            this.paste(cx);
        }))
        .on_action(cx.listener(|this, _: &DeleteCell, _, cx| {
            if !this.mode.is_editing() {
                this.delete_selection(cx);
            }
        }))
        // Insert/Delete rows/columns (Ctrl+= / Ctrl+-)
        .on_action(cx.listener(|this, _: &InsertRowsOrCols, _, cx| {
            if !this.mode.is_editing() {
                this.insert_rows_or_cols(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteRowsOrCols, _, cx| {
            if !this.mode.is_editing() {
                this.delete_rows_or_cols(cx);
            }
        }))
        // Undo/Redo
        .on_action(cx.listener(|this, _: &Undo, _, cx| {
            this.undo(cx);
        }))
        .on_action(cx.listener(|this, _: &Redo, _, cx| {
            this.redo(cx);
        }))
        // Editing actions
        .on_action(cx.listener(|this, _: &StartEdit, _, cx| {
            this.start_edit(cx);
            // On macOS, show tip about enabling F2 (catches Ctrl+U and menu-driven edit)
            this.maybe_show_f2_tip(cx);
        }))
        .on_action(cx.listener(|this, _: &ConfirmEdit, window, cx| {
            // Lua console handles its own Enter
            if this.lua_console.visible {
                crate::views::lua_console::execute_console(this, cx);
                return;
            }
            // If autocomplete is visible, Enter accepts the suggestion
            if this.autocomplete_visible {
                this.autocomplete_accept(cx);
                return;
            }
            // Handle Enter key based on current mode
            match this.mode {
                Mode::ThemePicker => this.theme_picker_execute(cx),
                Mode::FontPicker => this.font_picker_execute(cx),
                Mode::Command => this.palette_execute(cx),
                Mode::GoTo => this.confirm_goto(cx),
                _ => this.confirm_edit(cx),
            }
        }))
        .on_action(cx.listener(|this, _: &ConfirmEditUp, _, cx| {
            // Lua console handles its own Shift+Enter (insert newline)
            if this.lua_console.visible {
                this.lua_console.insert("\n");
                cx.notify();
                return;
            }
            // Shift+Enter: confirm and move up (or just move up in nav mode)
            match this.mode {
                Mode::ThemePicker | Mode::FontPicker | Mode::Command | Mode::GoTo => {
                    // Shift+Enter does nothing special in these modes
                }
                _ => this.confirm_edit_up(cx),
            }
        }))
        .on_action(cx.listener(|this, _: &CancelEdit, window, cx| {
            // Lua console handles its own Escape
            if this.lua_console.visible {
                this.lua_console.hide();
                window.focus(&this.focus_handle, cx);
                cx.notify();
                return;
            }
            // Import overlay takes priority - dismiss it but let import continue
            if this.import_overlay_visible {
                this.dismiss_import_overlay(cx);
                return;
            }
            if this.open_menu.is_some() {
                this.close_menu(cx);
            } else if this.mode == Mode::Command {
                this.hide_palette(cx);
            } else if this.mode == Mode::GoTo {
                this.hide_goto(cx);
            } else if this.mode == Mode::Find {
                this.hide_find(cx);
            } else if this.mode == Mode::FontPicker {
                this.hide_font_picker(cx);
            } else if this.mode == Mode::ThemePicker {
                this.hide_theme_picker(cx);
            } else if this.mode == Mode::About {
                this.hide_about(cx);
            } else if this.mode == Mode::RenameSymbol {
                this.hide_rename_symbol(cx);
            } else if this.mode == Mode::CreateNamedRange {
                this.hide_create_named_range(cx);
            } else if this.mode == Mode::EditDescription {
                this.hide_edit_description(cx);
            } else if this.mode == Mode::Tour {
                this.hide_tour(cx);
            } else if this.mode == Mode::ImpactPreview {
                this.hide_impact_preview(cx);
            } else if this.mode == Mode::RefactorLog {
                this.hide_refactor_log(cx);
            } else if this.mode == Mode::ExtractNamedRange {
                this.hide_extract_named_range(cx);
            } else if this.mode == Mode::ImportReport {
                this.hide_import_report(cx);
            } else if this.mode == Mode::Preferences {
                this.hide_preferences(cx);
            } else if this.mode == Mode::License {
                this.hide_license(cx);
            } else if this.inspector_visible && this.mode == Mode::Navigation {
                // Esc closes inspector panel when in navigation mode
                this.inspector_visible = false;
                cx.notify();
            } else {
                this.cancel_edit(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabNext, _, cx| {
            // If autocomplete is visible, Tab accepts the suggestion
            if this.autocomplete_visible {
                this.autocomplete_accept(cx);
                return;
            }
            if this.mode.is_editing() {
                this.confirm_edit_and_move_right(cx);
            } else {
                this.move_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabPrev, _, cx| {
            if this.mode.is_editing() {
                this.confirm_edit_and_move_left(cx);
            } else {
                this.move_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &BackspaceChar, _, cx| {
            // Lua console handles its own backspace
            if this.lua_console.visible {
                this.lua_console.backspace();
                cx.notify();
                return;
            }
            if this.mode == Mode::Command {
                this.palette_backspace(cx);
            } else if this.mode == Mode::GoTo {
                this.goto_backspace(cx);
            } else if this.mode == Mode::Find {
                this.find_backspace(cx);
            } else {
                this.backspace(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteChar, _, cx| {
            // Lua console handles its own delete
            if this.lua_console.visible {
                this.lua_console.delete();
                cx.notify();
                return;
            }
            this.delete_char(cx);
        }))
        .on_action(cx.listener(|this, _: &FillDown, _, cx| {
            this.fill_down(cx);
        }))
        .on_action(cx.listener(|this, _: &FillRight, _, cx| {
            this.fill_right(cx);
        }))
        .on_action(cx.listener(|this, _: &AutoSum, _, cx| {
            this.autosum(cx);
        }))
        .on_action(cx.listener(|this, _: &TrimWhitespace, _, cx| {
            this.trim_whitespace(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleFormulaView, _, cx| {
            this.toggle_show_formulas(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleShowZeros, _, cx| {
            this.toggle_show_zeros(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleInspector, _, cx| {
            this.inspector_visible = !this.inspector_visible;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ToggleZenMode, _, cx| {
            this.zen_mode = !this.zen_mode;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ToggleLuaConsole, window, cx| {
            // Pro feature gate
            if !visigrid_license::is_feature_enabled("lua") {
                this.status_message = Some("Lua scripting requires VisiGrid Pro".to_string());
                cx.notify();
                return;
            }
            this.lua_console.toggle();
            if this.lua_console.visible {
                window.focus(&this.console_focus_handle, cx);
            }
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ShowFormatPanel, _, cx| {
            this.inspector_visible = true;
            this.inspector_tab = InspectorTab::Format;
            cx.notify();
        }))
        // Edit mode cursor movement
        .on_action(cx.listener(|this, _: &EditCursorLeft, _, cx| {
            this.move_edit_cursor_left(cx);
        }))
        .on_action(cx.listener(|this, _: &EditCursorRight, _, cx| {
            this.move_edit_cursor_right(cx);
        }))
        .on_action(cx.listener(|this, _: &EditCursorHome, _, cx| {
            if this.mode.is_editing() {
                this.move_edit_cursor_home(cx);
            } else {
                // Navigation mode: go to first column of current row
                this.selected.1 = 0;
                this.selection_end = None;
                this.scroll_col = 0;
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &EditCursorEnd, _, cx| {
            if this.mode.is_editing() {
                this.move_edit_cursor_end(cx);
            } else {
                // Navigation mode: go to last column of current row
                this.selected.1 = crate::app::NUM_COLS - 1;
                this.selection_end = None;
                this.scroll_col = crate::app::NUM_COLS.saturating_sub(this.visible_cols());
                cx.notify();
            }
        }))
        // F4 reference cycling
        .on_action(cx.listener(|this, _: &CycleReference, _, cx| {
            this.cycle_reference(cx);
        }))
        .on_action(cx.listener(|this, _: &ConfirmEditInPlace, _, cx| {
            // Lua console handles its own Ctrl+Enter (execute)
            if this.lua_console.visible {
                crate::views::lua_console::execute_console(this, cx);
                return;
            }
            this.confirm_edit_in_place(cx);
        }))
        // Formatting
        .on_action(cx.listener(|this, _: &ToggleBold, _, cx| {
            this.toggle_bold(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleItalic, _, cx| {
            this.toggle_italic(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleUnderline, _, cx| {
            this.toggle_underline(cx);
        }))
        .on_action(cx.listener(|this, _: &FormatCurrency, _, cx| {
            this.format_currency(cx);
        }))
        .on_action(cx.listener(|this, _: &FormatPercent, _, cx| {
            this.format_percent(cx);
        }))
        // Go To dialog
        .on_action(cx.listener(|this, _: &GoToCell, _, cx| {
            this.show_goto(cx);
        }))
        // Find dialog
        .on_action(cx.listener(|this, _: &FindInCells, _, cx| {
            this.show_find(cx);
        }))
        .on_action(cx.listener(|this, _: &FindNext, _, cx| {
            this.find_next(cx);
        }))
        .on_action(cx.listener(|this, _: &FindPrev, _, cx| {
            this.find_prev(cx);
        }))
        // IDE-style navigation (Find References / Go to Precedents)
        .on_action(cx.listener(|this, _: &FindReferences, _, cx| {
            // In edit mode, check if cursor is on a named range
            if this.mode.is_editing() {
                if let Some(name) = this.named_range_at_cursor() {
                    this.show_named_range_references(&name, cx);
                    return;
                }
            }
            // Fall back to cell references
            let (row, col) = this.selected;
            this.show_references(row, col, cx);
        }))
        .on_action(cx.listener(|this, _: &GoToPrecedents, _, cx| {
            // In edit mode, check if cursor is on a named range
            if this.mode.is_editing() {
                if let Some(name) = this.named_range_at_cursor() {
                    this.go_to_named_range_definition(&name, cx);
                    return;
                }
            }
            // Fall back to cell precedents
            let (row, col) = this.selected;
            this.show_precedents(row, col, cx);
        }))
        .on_action(cx.listener(|this, _: &RenameSymbol, _, cx| {
            this.show_rename_symbol(None, cx);
        }))
        .on_action(cx.listener(|this, _: &CreateNamedRange, _, cx| {
            this.show_create_named_range(cx);
        }))
        // Command palette
        .on_action(cx.listener(|this, _: &ToggleCommandPalette, _, cx| {
            this.toggle_palette(cx);
        }))
        // View actions (for native macOS menus)
        .on_action(cx.listener(|this, _: &ShowAbout, _, cx| {
            this.show_about(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowFontPicker, _, cx| {
            this.show_font_picker(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowPreferences, _, cx| {
            this.show_preferences(cx);
        }))
        .on_action(cx.listener(|this, _: &OpenKeybindings, _, cx| {
            this.open_keybindings(cx);
        }))
        .on_action(cx.listener(|_this, _: &CloseWindow, window, _cx| {
            // Close the current window (not quit the app)
            window.remove_window();
        }))
        // Menu bar (Alt+letter accelerators)
        .on_action(cx.listener(|this, _: &OpenFileMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::File, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenEditMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Edit, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenViewMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::View, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenInsertMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Insert, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenFormatMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Format, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenDataMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Data, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenHelpMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Help, cx);
        }))
        // Sheet navigation
        .on_action(cx.listener(|this, _: &NextSheet, _, cx| {
            this.next_sheet(cx);
        }))
        .on_action(cx.listener(|this, _: &PrevSheet, _, cx| {
            this.prev_sheet(cx);
        }))
        .on_action(cx.listener(|this, _: &AddSheet, _, cx| {
            this.add_sheet(cx);
        }))
        // Character input (handles editing, goto, find, and command modes)
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            // Route keys to Lua console when visible
            if this.lua_console.visible {
                crate::views::lua_console::handle_console_key_from_main(this, event, window, cx);
                return;
            }

            // Handle sheet context menu (close on any key)
            if this.sheet_context_menu.is_some() {
                this.hide_sheet_context_menu(cx);
                if event.keystroke.key == "escape" {
                    return;
                }
            }

            // Exit zen mode with Escape
            if this.zen_mode && event.keystroke.key == "escape" {
                this.zen_mode = false;
                cx.notify();
                return;
            }

            // Handle sheet rename mode
            if this.renaming_sheet.is_some() {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.cancel_sheet_rename(cx);
                        return;
                    }
                    "enter" => {
                        this.confirm_sheet_rename(cx);
                        return;
                    }
                    "backspace" => {
                        this.sheet_rename_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for rename
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            this.sheet_rename_input_char(c, cx);
                        }
                        return;
                    }
                }
                return;
            }

            // Handle Formula Autocomplete (highest priority when visible)
            if this.autocomplete_visible {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.autocomplete_dismiss(cx);
                        return;
                    }
                    "enter" | "tab" => {
                        this.autocomplete_accept(cx);
                        return;
                    }
                    "up" => {
                        this.autocomplete_up(cx);
                        return;
                    }
                    "down" => {
                        this.autocomplete_down(cx);
                        return;
                    }
                    "shift-tab" => {
                        // Dismiss autocomplete on Shift+Tab (spec: no accept)
                        this.autocomplete_dismiss(cx);
                        return;
                    }
                    _ => {
                        // Other keys: let them pass through to normal handling
                        // but the input will update autocomplete
                    }
                }
            }

            // Handle Command Palette mode
            if this.mode == Mode::Command {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_palette(cx);
                        return;
                    }
                    "enter" => {
                        if event.keystroke.modifiers.control {
                            // Ctrl+Enter = secondary action (copy path, copy ref, show help)
                            this.palette_execute_secondary(cx);
                        } else if event.keystroke.modifiers.shift {
                            // Shift+Enter = preview (apply without closing)
                            this.palette_preview(cx);
                        } else {
                            // Plain Enter = execute and close
                            this.palette_execute(cx);
                        }
                        return;
                    }
                    "up" => {
                        this.palette_up(cx);
                        return;
                    }
                    "down" => {
                        this.palette_down(cx);
                        return;
                    }
                    "backspace" => {
                        this.palette_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for palette
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars() {
                            this.palette_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Font Picker mode
            if this.mode == Mode::FontPicker {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_font_picker(cx);
                        return;
                    }
                    "enter" => {
                        this.font_picker_execute(cx);
                        return;
                    }
                    "up" => {
                        this.font_picker_up(cx);
                        return;
                    }
                    "down" => {
                        this.font_picker_down(cx);
                        return;
                    }
                    "backspace" => {
                        this.font_picker_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for font picker
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars() {
                            this.font_picker_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Theme Picker mode
            if this.mode == Mode::ThemePicker {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_theme_picker(cx);
                        return;
                    }
                    "enter" => {
                        this.theme_picker_execute(cx);
                        return;
                    }
                    "up" => {
                        this.theme_picker_up(cx);
                        return;
                    }
                    "down" => {
                        this.theme_picker_down(cx);
                        return;
                    }
                    "backspace" => {
                        this.theme_picker_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for theme picker
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars() {
                            this.theme_picker_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Rename Symbol mode
            if this.mode == Mode::RenameSymbol {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_rename_symbol(cx);
                        return;
                    }
                    "enter" => {
                        this.confirm_rename_symbol(cx);
                        return;
                    }
                    "backspace" => {
                        this.rename_symbol_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for rename
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            this.rename_symbol_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Create Named Range mode
            if this.mode == Mode::CreateNamedRange {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_create_named_range(cx);
                        return;
                    }
                    "enter" => {
                        this.confirm_create_named_range(cx);
                        return;
                    }
                    "backspace" => {
                        this.create_name_backspace(cx);
                        return;
                    }
                    "tab" => {
                        this.create_name_tab(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for create name
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            this.create_name_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Edit Description mode
            if this.mode == Mode::EditDescription {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_edit_description(cx);
                        return;
                    }
                    "enter" => {
                        this.apply_edit_description(cx);
                        return;
                    }
                    "backspace" => {
                        this.edit_description_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for description
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            this.edit_description_insert_char(c, cx);
                        }
                        return;
                    }
                }
            }

            // Handle Tour mode
            if this.mode == Mode::Tour {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_tour(cx);
                        return;
                    }
                    "enter" | "right" => {
                        if this.tour_step < 3 {
                            this.tour_next(cx);
                        } else {
                            this.tour_done(cx);
                        }
                        return;
                    }
                    "left" => {
                        this.tour_back(cx);
                        return;
                    }
                    _ => {}
                }
                return; // Consume all keystrokes in tour mode
            }

            // Handle Impact Preview mode
            if this.mode == Mode::ImpactPreview {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_impact_preview(cx);
                        return;
                    }
                    "enter" => {
                        this.apply_impact_preview(cx);
                        return;
                    }
                    _ => {}
                }
                return; // Consume all keystrokes in impact preview mode
            }

            // Handle Refactor Log mode
            if this.mode == Mode::RefactorLog {
                match event.keystroke.key.as_str() {
                    "escape" | "enter" => {
                        this.hide_refactor_log(cx);
                        return;
                    }
                    _ => {}
                }
                return; // Consume all keystrokes in refactor log mode
            }

            // Handle Extract Named Range mode
            if this.mode == Mode::ExtractNamedRange {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_extract_named_range(cx);
                        return;
                    }
                    "enter" => {
                        this.confirm_extract_named_range(cx);
                        return;
                    }
                    "backspace" => {
                        match this.extract_focus {
                            CreateNameFocus::Name => this.extract_name_backspace(cx),
                            CreateNameFocus::Description => this.extract_description_backspace(cx),
                        }
                        return;
                    }
                    "tab" => {
                        this.extract_tab(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for extract name/description
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            match this.extract_focus {
                                CreateNameFocus::Name => this.extract_name_insert_char(c, cx),
                                CreateNameFocus::Description => this.extract_description_insert_char(c, cx),
                            }
                        }
                        return;
                    }
                }
                return; // Consume all keystrokes in extract mode
            }

            // Handle License mode
            if this.mode == Mode::License {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_license(cx);
                        return;
                    }
                    "enter" => {
                        if !this.license_input.is_empty() {
                            this.apply_license(cx);
                        }
                        return;
                    }
                    "backspace" => {
                        this.license_backspace(cx);
                        return;
                    }
                    _ => {}
                }

                // Handle text input for license
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| !c.is_control()) {
                            this.license_insert_char(c, cx);
                        }
                        return;
                    }
                }
                return; // Consume all keystrokes in license mode
            }

            // Handle About mode (consume keystrokes, only escape closes)
            if this.mode == Mode::About {
                if event.keystroke.key == "escape" {
                    this.hide_about(cx);
                }
                return; // Consume all keystrokes in about mode
            }

            // Handle Preferences mode (consume keystrokes, only escape closes)
            if this.mode == Mode::Preferences {
                if event.keystroke.key == "escape" {
                    this.hide_preferences(cx);
                }
                return; // Consume all keystrokes in preferences mode
            }

            // Handle Hint mode (Vimium-style jump navigation)
            if this.mode == Mode::Hint {
                let key = event.keystroke.key.as_str();
                if this.apply_hint_key(key, cx) {
                    return;
                }
                // Unhandled key in hint mode - exit without action
                this.exit_hint_mode(cx);
                return;
            }

            // Handle Navigation mode with keyboard hints or vim mode enabled
            // (before Names tab filter and before regular text input)
            if this.mode == Mode::Navigation
                && !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
                && !event.keystroke.modifiers.shift
            {
                let key = event.keystroke.key.as_str();
                let hints_enabled = this.keyboard_hints_enabled(cx);
                let vim_enabled = this.vim_mode_enabled(cx);

                // 'g' enters command/hint mode (if hints OR vim mode enabled)
                // - With hints: shows cell labels + g-commands (gg)
                // - With vim only: just g-commands (gg), no labels
                if (hints_enabled || vim_enabled) && key == "g" {
                    this.enter_hint_mode_with_labels(hints_enabled, cx);
                    return;
                }

                // Vim keys (if vim mode enabled)
                if vim_enabled {
                    if this.apply_vim_key(key, cx) {
                        return;
                    }
                    // For non-vim keys in vim mode, don't start edit - just ignore
                    if key.len() == 1 && key.chars().next().map(|c| c.is_alphabetic()).unwrap_or(false) {
                        // In vim mode, typing letters doesn't start edit
                        // Show status hint for unknown vim key
                        this.status_message = Some(format!("Unknown vim key: {}", key));
                        cx.notify();
                        return;
                    }
                }
            }

            // Handle Names tab filter input (when inspector visible + Names tab + Navigation mode)
            if this.inspector_visible
                && this.inspector_tab == InspectorTab::Names
                && this.mode == Mode::Navigation
            {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        if !this.names_filter_query.is_empty() {
                            // Clear filter first
                            this.names_filter_query.clear();
                            cx.notify();
                        } else {
                            // Close inspector
                            this.inspector_visible = false;
                            cx.notify();
                        }
                        return;
                    }
                    "backspace" => {
                        if !this.names_filter_query.is_empty() {
                            this.names_filter_query.pop();
                            cx.notify();
                        }
                        return;
                    }
                    "/" => {
                        // "/" focuses filter (already focused, this just prevents "/" from being typed)
                        return;
                    }
                    _ => {}
                }

                // Handle text input for filter (only alphanumeric and underscore)
                if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control
                        && !event.keystroke.modifiers.alt
                        && !event.keystroke.modifiers.platform
                    {
                        for c in key_char.chars().filter(|c| c.is_alphanumeric() || *c == '_' || *c == '.') {
                            this.names_filter_query.push(c);
                        }
                        cx.notify();
                        return;
                    }
                }
            }

            // Handle GoTo mode
            if this.mode == Mode::GoTo {
                if event.keystroke.key == "enter" {
                    this.confirm_goto(cx);
                    return;
                } else if event.keystroke.key == "escape" {
                    this.hide_goto(cx);
                    return;
                } else if event.keystroke.key == "backspace" {
                    this.goto_backspace(cx);
                    return;
                }
            }

            // Handle Find mode
            if this.mode == Mode::Find {
                if event.keystroke.key == "escape" {
                    this.hide_find(cx);
                    return;
                } else if event.keystroke.key == "backspace" {
                    this.find_backspace(cx);
                    return;
                }
            }

            if let Some(key_char) = &event.keystroke.key_char {
                if !event.keystroke.modifiers.control
                    && !event.keystroke.modifiers.alt
                    && !event.keystroke.modifiers.platform
                {
                    // Filter out control characters - let them be handled by actions instead
                    let printable_chars: String = key_char.chars()
                        .filter(|c| !c.is_control())
                        .collect();

                    if !printable_chars.is_empty() {
                        match this.mode {
                            Mode::GoTo => {
                                for c in printable_chars.chars() {
                                    this.goto_insert_char(c, cx);
                                }
                            }
                            Mode::Find => {
                                for c in printable_chars.chars() {
                                    this.find_insert_char(c, cx);
                                }
                            }
                            _ => {
                                for c in printable_chars.chars() {
                                    this.insert_char(c, cx);
                                }
                            }
                        }
                    }
                }
            }
        }))
        // Mouse wheel scrolling
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
            let delta = event.delta.pixel_delta(px(CELL_HEIGHT));
            // Convert pixel delta to row/col delta (negative Y = scroll up)
            let dy: f32 = delta.y.into();
            let dx: f32 = delta.x.into();
            let delta_rows = (-dy / CELL_HEIGHT).round() as i32;
            let delta_cols = (-dx / CELL_HEIGHT).round() as i32;
            if delta_rows != 0 || delta_cols != 0 {
                this.scroll(delta_rows, delta_cols, cx);
            }
        }))
        // Mouse move for resize and header selection dragging
        .on_mouse_move(cx.listener(|this, event: &MouseMoveEvent, _, cx| {
            // Handle Lua console resize drag
            if this.lua_console.resizing {
                let y: f32 = event.position.y.into();
                let delta = this.lua_console.resize_start_y - y; // Inverted: dragging up increases height
                let new_height = (this.lua_console.resize_start_height + delta)
                    .max(crate::scripting::MIN_CONSOLE_HEIGHT)
                    .min(crate::scripting::MAX_CONSOLE_HEIGHT);
                this.lua_console.height = new_height;
                cx.notify();
                return; // Don't process other drags while resizing console
            }
            // Handle column resize drag
            if let Some(col) = this.resizing_col {
                let x: f32 = event.position.x.into();
                let delta = x - this.resize_start_pos;
                let new_width = this.resize_start_size + delta;
                this.set_col_width(col, new_width);
                cx.notify();
            }
            // Handle row resize drag
            if let Some(row) = this.resizing_row {
                let y: f32 = event.position.y.into();
                let delta = y - this.resize_start_pos;
                let new_height = this.resize_start_size + delta;
                this.set_row_height(row, new_height);
                cx.notify();
            }
            // Handle column header selection drag
            if this.dragging_col_header {
                let x: f32 = event.position.x.into();
                if let Some(col) = this.col_from_window_x(x) {
                    this.continue_col_header_drag(col, cx);
                }
            }
            // Handle row header selection drag
            if this.dragging_row_header {
                let y: f32 = event.position.y.into();
                if let Some(row) = this.row_from_window_y(y) {
                    this.continue_row_header_drag(row, cx);
                }
            }
        }))
        // Mouse up to end resize and header selection drag
        .on_mouse_up(MouseButton::Left, cx.listener(|this, _, _, cx| {
            // End Lua console resize
            if this.lua_console.resizing {
                this.lua_console.resizing = false;
                cx.notify();
            }
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
                this.resizing_col = None;
                this.resizing_row = None;
                cx.notify();
            }
            // End header selection drag
            if this.dragging_row_header {
                this.end_row_header_drag(cx);
            }
            if this.dragging_col_header {
                this.end_col_header_drag(cx);
            }
        }))
        .flex()
        .flex_col()
        .size_full()
        .bg(app.token(TokenKey::AppBg))
        // Hide in-app menu bar on macOS (uses native menu bar instead)
        // Also hide in zen mode
        .when(!cfg!(target_os = "macos") && !zen_mode, |div| {
            div.child(menu_bar::render_menu_bar(app, cx))
        })
        .when(!zen_mode, |div| {
            div.child(formula_bar::render_formula_bar(app, cx))
        })
        .child(headers::render_column_headers(app, cx))
        .child(grid::render_grid(app, window, cx))
        // Lua console panel (above status bar)
        .child(lua_console::render_lua_console(app, cx))
        .when(!zen_mode, |div| {
            div.child(status_bar::render_status_bar(app, editing, cx))
        })
        .when(show_goto, |div| {
            div.child(goto_dialog::render_goto_dialog(app))
        })
        .when(show_find, |div| {
            div.child(find_dialog::render_find_dialog(app))
        })
        .when(show_command, |div| {
            div.child(command_palette::render_command_palette(app, cx))
        })
        .when(show_font_picker, |div| {
            div.child(font_picker::render_font_picker(app, cx))
        })
        .when(show_theme_picker, |div| {
            div.child(theme_picker::render_theme_picker(app, cx))
        })
        .when(show_preferences, |div| {
            div.child(preferences_panel::render_preferences_panel(app, cx))
        })
        .when(show_about, |div| {
            div.child(about_dialog::render_about_dialog(app, cx))
        })
        .when(show_license, |div| {
            div.child(license_dialog::render_license_dialog(app, cx))
        })
        .when(show_import_report, |div| {
            div.child(import_report_dialog::render_import_report_dialog(app, cx))
        })
        .when(show_export_report, |div| {
            div.child(export_report_dialog::render_export_report_dialog(app, cx))
        })
        // Import overlay: shows during background Excel imports (after 150ms delay)
        .when(show_import_overlay, |div| {
            div.child(import_overlay::render_import_overlay(app, cx))
        })
        .when(show_rename_symbol, |div| {
            div.child(render_rename_symbol_dialog(app))
        })
        .when(show_create_named_range, |div| {
            div.child(render_create_named_range_dialog(app))
        })
        .when(show_edit_description, |div| {
            div.child(render_edit_description_dialog(app))
        })
        .when(show_tour, |div| {
            div.child(tour::render_tour(app, cx))
        })
        .when(show_impact_preview, |div| {
            div.child(impact_preview::render_impact_preview(app, cx))
        })
        .when(show_refactor_log, |div| {
            div.child(refactor_log::render_refactor_log(app, cx))
        })
        .when(show_extract_named_range, |div| {
            div.child(render_extract_named_range_dialog(app))
        })
        // Name tooltip (one-time first-run hint)
        .when(show_name_tooltip, |div| {
            div.child(tour::render_name_tooltip(app, cx))
        })
        // F2 function key tip (macOS only, shown when editing via non-F2 path)
        .when(show_f2_tip, |div| {
            div.child(tour::render_f2_tooltip(app, cx))
        })
        // Menu dropdown (only on non-macOS where we have in-app menu)
        .when(!cfg!(target_os = "macos") && app.open_menu.is_some(), |div| {
            div.child(menu_bar::render_menu_dropdown(app, cx))
        })
        // Inspector panel (right-side drawer)
        .when(show_inspector, |div| {
            div.child(inspector_panel::render_inspector_panel(app, cx))
        })
        // Formula autocomplete popup (rendered at top level to avoid clipping)
        .when(app.autocomplete_visible, |div| {
            let suggestions = app.autocomplete_suggestions();
            let selected = app.autocomplete_selected;
            // Calculate popup position below the active cell
            let popup_x = HEADER_WIDTH + app.col_x_offset(app.selected.1);
            let popup_y = CELL_HEIGHT * 3.0 + app.row_y_offset(app.selected.0) + app.row_height(app.selected.0);
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let selection_bg = app.token(TokenKey::SelectionBg);
            div.child(formula_bar::render_autocomplete_popup(
                &suggestions,
                selected,
                popup_x,
                popup_y,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                selection_bg,
                cx,
            ))
        })
        // Formula signature help (rendered at top level)
        .when_some(app.signature_help(), |div, sig_info| {
            // Calculate popup position below the active cell
            let popup_x = HEADER_WIDTH + app.col_x_offset(app.selected.1);
            let popup_y = CELL_HEIGHT * 3.0 + app.row_y_offset(app.selected.0) + app.row_height(app.selected.0);
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let accent = app.token(TokenKey::Accent);
            div.child(formula_bar::render_signature_help(
                &sig_info,
                popup_x,
                popup_y,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                accent,
            ))
        })
        // Formula error banner (rendered at top level)
        .when_some(app.formula_error(), |div, error_info| {
            // Calculate popup position below the active cell
            let popup_x = HEADER_WIDTH + app.col_x_offset(app.selected.1);
            let popup_y = CELL_HEIGHT * 3.0 + app.row_y_offset(app.selected.0) + app.row_height(app.selected.0);
            let error_bg = app.token(TokenKey::ErrorBg);
            let error_color = app.token(TokenKey::Error);
            let panel_border = app.token(TokenKey::PanelBorder);
            div.child(formula_bar::render_error_banner(&error_info, popup_x, popup_y, error_bg, error_color, panel_border))
        })
        // Hover documentation popup (when not editing and hovering over formula bar)
        .when_some(app.hover_function.filter(|_| !app.mode.is_editing() && !app.autocomplete_visible), |div, func| {
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let accent = app.token(TokenKey::Accent);
            div.child(formula_bar::render_hover_docs(func, panel_bg, panel_border, text_primary, text_muted, accent))
        })
}

/// Render the rename symbol dialog (Ctrl+Shift+R)
fn render_rename_symbol_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let selection_bg = app.token(TokenKey::SelectionBg);

    let affected_count = app.rename_affected_cells.len();
    let select_all = app.rename_select_all;
    let has_error = app.rename_validation_error.is_some();

    // Build affected cells preview
    let cells_preview: Vec<String> = app.rename_affected_cells
        .iter()
        .take(8)
        .map(|(row, col)| {
            let col_letter = col_to_letter(*col);
            format!("{}{}", col_letter, row + 1)
        })
        .collect();

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child(format!("Rename '{}'", app.rename_original_name))
                )
                // Input field
                .child(
                    div()
                        .px_2()
                        .py_1()
                        .bg(hsla(0.0, 0.0, 0.0, 0.2))
                        .rounded_sm()
                        .border_1()
                        .border_color(if has_error { error_color } else { panel_border })
                        .text_color(text_primary)
                        .child(
                            // Show text with selection highlight if select_all is active
                            div()
                                .when(select_all, |d| d.bg(selection_bg).rounded_sm().px_1())
                                .child(app.rename_new_name.clone())
                        )
                )
                // Validation error (if any)
                .when_some(app.rename_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Affected cells count
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child(format!(
                            "{} formula{} will be updated",
                            affected_count,
                            if affected_count == 1 { "" } else { "s" }
                        ))
                )
                // Preview of affected cells (show list of cell refs)
                .when(affected_count > 0, |d| {
                    let preview = cells_preview.join(", ");
                    let more = if affected_count > 8 {
                        format!(" ...and {} more", affected_count - 8)
                    } else {
                        String::new()
                    };
                    d.child(
                        div()
                            .text_color(text_muted)
                            .text_xs()
                            .child(format!("{}{}", preview, more))
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Enter to confirm • Escape to cancel")
                )
        )
}

/// Render the edit description dialog
fn render_edit_description_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child(format!("Edit description for '{}'", app.edit_description_name))
                )
                // Input field
                .child(
                    div()
                        .px_2()
                        .py_2()
                        .bg(hsla(0.0, 0.0, 0.0, 0.2))
                        .rounded_sm()
                        .border_1()
                        .border_color(panel_border)
                        .text_color(if app.edit_description_value.is_empty() { text_muted } else { text_primary })
                        .min_h(px(60.0))
                        .child(if app.edit_description_value.is_empty() {
                            "Enter a description...".to_string()
                        } else {
                            app.edit_description_value.clone()
                        })
                )
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Enter to save • Escape to cancel")
                )
        )
}

/// Convert column index to letter(s) (0 = A, 25 = Z, 26 = AA, etc.)
fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

/// Render the create named range dialog (Ctrl+Shift+N)
fn render_create_named_range_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let has_error = app.create_name_validation_error.is_some();
    let name_focused = app.create_name_focus == CreateNameFocus::Name;
    let desc_focused = app.create_name_focus == CreateNameFocus::Description;

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("Create Named Range")
                )
                // Target (read-only display)
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Target:")
                        )
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_sm()
                                .child(app.create_name_target.clone())
                        )
                )
                // Name input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Name:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if name_focused && has_error {
                                    error_color
                                } else if name_focused {
                                    accent
                                } else {
                                    panel_border
                                })
                                .text_color(text_primary)
                                .child(if app.create_name_name.is_empty() && !name_focused {
                                    "(required)".to_string()
                                } else {
                                    app.create_name_name.clone()
                                })
                        )
                )
                // Description input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Description:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if desc_focused { accent } else { panel_border })
                                .text_color(if app.create_name_description.is_empty() { text_muted } else { text_primary })
                                .child(if app.create_name_description.is_empty() {
                                    "(optional)".to_string()
                                } else {
                                    app.create_name_description.clone()
                                })
                        )
                )
                // Validation error (if any)
                .when_some(app.create_name_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Tab to switch fields • Enter to confirm • Escape to cancel")
                )
        )
}

/// Render the extract named range dialog
fn render_extract_named_range_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let has_error = app.extract_validation_error.is_some();
    let name_focused = app.extract_focus == CreateNameFocus::Name;
    let desc_focused = app.extract_focus == CreateNameFocus::Description;

    // Format occurrence message
    let occurrence_msg = if app.extract_affected_cells.len() == 1 {
        format!("Will replace {} occurrence in 1 formula", app.extract_occurrence_count)
    } else {
        format!("Will replace {} occurrences in {} formulas", app.extract_occurrence_count, app.extract_affected_cells.len())
    };

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(420.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("Extract to Named Range")
                )
                // Range preview (read-only)
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Range:")
                        )
                        .child(
                            div()
                                .text_color(accent)
                                .text_sm()
                                .font_weight(FontWeight::MEDIUM)
                                .child(app.extract_range_literal.clone())
                        )
                )
                // Name input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Name:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(if app.extract_select_all && name_focused {
                                    accent.opacity(0.3)  // Selection highlight
                                } else {
                                    hsla(0.0, 0.0, 0.0, 0.2)
                                })
                                .rounded_sm()
                                .border_1()
                                .border_color(if name_focused && has_error {
                                    error_color
                                } else if name_focused {
                                    accent
                                } else {
                                    panel_border
                                })
                                .text_color(text_primary)
                                .child(if app.extract_name.is_empty() && !name_focused {
                                    "(required)".to_string()
                                } else {
                                    app.extract_name.clone()
                                })
                        )
                )
                // Description input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Description:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if desc_focused { accent } else { panel_border })
                                .text_color(if app.extract_description.is_empty() { text_muted } else { text_primary })
                                .child(if app.extract_description.is_empty() {
                                    "(optional)".to_string()
                                } else {
                                    app.extract_description.clone()
                                })
                        )
                )
                // Occurrence count
                .child(
                    div()
                        .px_2()
                        .py_2()
                        .bg(hsla(0.0, 0.0, 0.0, 0.15))
                        .rounded_sm()
                        .text_color(text_muted)
                        .text_xs()
                        .child(occurrence_msg)
                )
                // Validation error (if any)
                .when_some(app.extract_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Tab to switch fields • Enter to extract • Escape to cancel")
                )
        )
}
