mod about_dialog;
pub mod command_palette;
mod hub_dialogs;
mod export_report_dialog;
mod filter_dropdown;
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
mod validation_dialog;
mod validation_dropdown_view;

use std::time::Duration;
use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CELL_HEIGHT, CreateNameFocus};
use crate::search::MenuCategory;
use crate::actions::*;
use crate::formatting::BorderApplyMode;
use crate::mode::{Mode, InspectorTab};
use crate::theme::TokenKey;

pub fn render_spreadsheet(app: &mut Spreadsheet, window: &mut Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    // Check if validation dropdown source has changed (fingerprint mismatch)
    app.check_dropdown_staleness(cx);

    // Auto-dismiss rewind success banner after 5 seconds
    if app.rewind_success.should_dismiss() {
        app.rewind_success.hide();
    }

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
    let show_hub_paste_token = app.mode == Mode::HubPasteToken;
    let show_hub_link = app.mode == Mode::HubLink;
    let show_hub_publish_confirm = app.mode == Mode::HubPublishConfirm;
    let show_validation_dialog = app.mode == Mode::ValidationDialog;
    let show_rewind_confirm = app.rewind_confirm.visible;
    let show_rewind_success = app.rewind_success.visible;
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
            // Validation dropdown navigation takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.move_up();
                    cx.notify();
                }
                return;
            }
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
            // Validation dropdown navigation takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.move_down();
                    cx.notify();
                }
                return;
            }
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
        .on_action(cx.listener(|this, _: &MoveLeft, window, cx| {
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
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &MoveRight, window, cx| {
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
            this.update_edit_scroll(window);
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
        .on_action(cx.listener(|this, _: &JumpLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_left(cx);
            } else {
                this.jump_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &JumpRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_right(cx);
            } else {
                this.jump_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &MoveToStart, _, cx| {
            this.view_state.selected = (0, 0);
            this.view_state.scroll_row = 0;
            this.view_state.scroll_col = 0;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &MoveToEnd, _, cx| {
            this.view_state.selected = (crate::app::NUM_ROWS - 1, crate::app::NUM_COLS - 1);
            this.view_state.scroll_row = crate::app::NUM_ROWS.saturating_sub(this.visible_rows());
            this.view_state.scroll_col = crate::app::NUM_COLS.saturating_sub(this.visible_cols());
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &PageUp, _, cx| {
            // Validation dropdown takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.page_up(10);
                    cx.notify();
                }
                return;
            }
            this.page_up(cx);
        }))
        .on_action(cx.listener(|this, _: &PageDown, _, cx| {
            // Validation dropdown takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.page_down(10);
                    cx.notify();
                }
                return;
            }
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
        .on_action(cx.listener(|this, _: &ExtendLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_left(cx);
            } else {
                this.extend_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &ExtendRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_right(cx);
            } else {
                this.extend_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
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
        .on_action(cx.listener(|this, _: &ExtendJumpLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_left(cx);
            } else {
                this.extend_jump_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_right(cx);
            } else {
                this.extend_jump_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &SelectAll, window, cx| {
            if this.mode == Mode::Edit {
                this.select_all_edit(cx);
            } else {
                this.select_all(cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &SelectBlanks, _, cx| {
            if !this.mode.is_editing() {
                this.select_blanks(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectRow, _, cx| {
            if !this.mode.is_editing() {
                let row = this.view_state.selected.0;
                this.select_row(row, false, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectColumn, _, cx| {
            if !this.mode.is_editing() {
                let col = this.view_state.selected.1;
                this.select_col(col, false, cx);
            }
        }))
        // File actions
        .on_action(cx.listener(|this, _: &NewFile, window, cx| {
            this.new_file(cx);
            this.update_title_if_needed(window);
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
        .on_action(cx.listener(|this, _: &ExportProvenance, _, cx| {
            this.export_provenance(cx);
        }))
        // VisiHub sync actions
        .on_action(cx.listener(|this, _: &HubCheckStatus, _, cx| {
            this.hub_check_status(cx);
        }))
        .on_action(cx.listener(|this, _: &HubPull, _, cx| {
            this.hub_pull(cx);
        }))
        .on_action(cx.listener(|this, _: &HubOpenRemoteAsCopy, _, cx| {
            this.hub_open_remote_as_copy(cx);
        }))
        .on_action(cx.listener(|this, _: &HubUnlink, _, cx| {
            this.hub_unlink(cx);
        }))
        .on_action(cx.listener(|this, _: &HubDiagnostics, _, cx| {
            this.hub_diagnostics(cx);
        }))
        .on_action(cx.listener(|this, _: &HubSignIn, _, cx| {
            this.hub_sign_in(cx);
        }))
        .on_action(cx.listener(|this, _: &HubSignOut, _, cx| {
            this.hub_sign_out(cx);
        }))
        .on_action(cx.listener(|this, _: &HubLink, _, cx| {
            this.hub_show_link_dialog(cx);
        }))
        .on_action(cx.listener(|this, _: &HubPublish, _, cx| {
            this.hub_publish(cx);
        }))
        // Clipboard actions
        .on_action(cx.listener(|this, _: &Copy, _, cx| {
            this.copy(cx);
        }))
        .on_action(cx.listener(|this, _: &Cut, window, cx| {
            this.cut(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &Paste, window, cx| {
            // Special handling for HubPasteToken mode - paste into token input
            if this.mode == Mode::HubPasteToken {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        this.hub_token_paste(&text, cx);
                    }
                }
                return;
            }
            // Special handling for CreateNamedRange mode - paste into focused field
            if this.mode == Mode::CreateNamedRange {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        for c in text.chars().filter(|c| !c.is_control()) {
                            this.create_name_insert_char(c, cx);
                        }
                    }
                }
                return;
            }
            this.paste(cx);
            this.update_edit_scroll(window);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &PasteValues, window, cx| {
            this.paste_values(cx);
            this.update_edit_scroll(window);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &DeleteCell, window, cx| {
            if !this.mode.is_editing() {
                this.delete_selection(cx);
                this.update_title_if_needed(window);
            }
        }))
        // Insert/Delete rows/columns (Ctrl+= / Ctrl+-)
        .on_action(cx.listener(|this, _: &InsertRowsOrCols, window, cx| {
            if !this.mode.is_editing() {
                this.insert_rows_or_cols(cx);
                this.update_title_if_needed(window);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteRowsOrCols, window, cx| {
            if !this.mode.is_editing() {
                this.delete_rows_or_cols(cx);
                this.update_title_if_needed(window);
            }
        }))
        // Undo/Redo
        .on_action(cx.listener(|this, _: &Undo, window, cx| {
            this.undo(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &Redo, window, cx| {
            this.redo(cx);
            this.update_title_if_needed(window);
        }))
        // Editing actions
        .on_action(cx.listener(|this, _: &StartEdit, window, cx| {
            this.start_edit(cx);
            this.update_edit_scroll(window);
            // On macOS, show tip about enabling F2 (catches Ctrl+U and menu-driven edit)
            this.maybe_show_f2_tip(cx);
        }))
        .on_action(cx.listener(|this, _: &ConfirmEdit, window, cx| {
            // Validation dropdown: Enter commits selected item
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open() {
                    if let Some(value) = state.selected_item() {
                        let value = value.to_string();
                        this.commit_validation_value(&value, cx);
                    } else {
                        // No items visible - just close
                        this.close_validation_dropdown(
                            crate::validation_dropdown::DropdownCloseReason::Escape,
                            cx,
                        );
                    }
                }
                return;
            }
            // Lua console handles its own Enter
            if this.lua_console.visible {
                crate::views::lua_console::execute_console(this, cx);
                return;
            }
            // If autocomplete is visible, Enter accepts the suggestion
            if this.autocomplete_visible {
                this.autocomplete_accept(cx);
                this.update_title_if_needed(window);
                return;
            }
            // Handle Enter key based on current mode
            match this.mode {
                Mode::ThemePicker => this.theme_picker_execute(cx),
                Mode::FontPicker => this.font_picker_execute(cx),
                Mode::Command => this.palette_execute(cx),
                Mode::GoTo => this.confirm_goto(cx),
                Mode::CreateNamedRange => this.confirm_create_named_range(cx),
                _ => {
                    this.confirm_edit(cx);
                    this.update_title_if_needed(window);
                }
            }
        }))
        .on_action(cx.listener(|this, _: &ConfirmEditUp, window, cx| {
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
                _ => {
                    this.confirm_edit_up(cx);
                    this.update_title_if_needed(window);
                }
            }
        }))
        .on_action(cx.listener(|this, _: &CancelEdit, window, cx| {
            // Validation dropdown: Escape closes without committing
            if this.is_validation_dropdown_open() {
                this.close_validation_dropdown(
                    crate::validation_dropdown::DropdownCloseReason::Escape,
                    cx,
                );
                return;
            }
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
            } else if this.filter_dropdown_col.is_some() {
                // Esc closes filter dropdown
                this.close_filter_dropdown(cx);
            } else if this.is_previewing() {
                // Esc exits preview mode
                this.exit_preview(cx);
            } else if this.inspector_visible && this.mode == Mode::Navigation {
                // Esc closes inspector panel when in navigation mode
                this.inspector_visible = false;
                this.history_highlight_range = None;  // Clear history highlight
                cx.notify();
            } else if this.history_highlight_range.is_some() {
                // Esc clears history highlight when nothing else to dismiss
                this.history_highlight_range = None;
                this.selected_history_id = None;
                cx.notify();
            } else {
                this.cancel_edit(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabNext, window, cx| {
            // Dialog modes handle Tab themselves
            if this.mode == Mode::CreateNamedRange {
                this.create_name_tab(cx);
                return;
            }
            // If autocomplete is visible, Tab accepts the suggestion
            if this.autocomplete_visible {
                this.autocomplete_accept(cx);
                this.update_title_if_needed(window);
                return;
            }
            if this.mode.is_editing() {
                this.confirm_edit_and_move_right(cx);
                this.update_title_if_needed(window);
            } else {
                this.move_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabPrev, window, cx| {
            if this.mode.is_editing() {
                this.confirm_edit_and_move_left(cx);
                this.update_title_if_needed(window);
            } else {
                this.move_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &BackspaceChar, window, cx| {
            // Dialog modes handle backspace themselves
            if this.mode == Mode::CreateNamedRange {
                this.create_name_backspace(cx);
                return;
            }
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
            } else if this.mode.is_editing() {
                this.backspace(cx);
                this.update_edit_scroll(window);
            } else {
                // Navigation mode: backspace clears selected cells (like Delete key)
                this.delete_selection(cx);
                this.update_title_if_needed(window);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteChar, window, cx| {
            // Lua console handles its own delete
            if this.lua_console.visible {
                this.lua_console.delete();
                cx.notify();
                return;
            }
            this.delete_char(cx);
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &FillDown, window, cx| {
            this.fill_down(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &FillRight, window, cx| {
            this.fill_right(cx);
            this.update_title_if_needed(window);
        }))
        // Data operations (sort/filter)
        .on_action(cx.listener(|this, _: &SortAscending, window, cx| {
            this.sort_by_current_column(visigrid_engine::filter::SortDirection::Ascending, cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &SortDescending, window, cx| {
            this.sort_by_current_column(visigrid_engine::filter::SortDirection::Descending, cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &ToggleAutoFilter, _, cx| {
            this.toggle_auto_filter(cx);
        }))
        .on_action(cx.listener(|this, _: &ClearSort, window, cx| {
            this.clear_sort(cx);
            this.update_title_if_needed(window);
        }))
        // Data validation
        .on_action(cx.listener(|this, _: &ShowDataValidation, _, cx| {
            // TODO: Show data validation dialog
            this.status_message = Some("Data validation dialog not yet implemented".to_string());
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &OpenValidationDropdown, _, cx| {
            this.open_validation_dropdown(cx);
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
        .on_action(cx.listener(|this, _: &ToggleVerifiedMode, _, cx| {
            this.toggle_verified_mode(cx);
        }))
        // Zoom
        .on_action(cx.listener(|this, _: &ZoomIn, _, cx| {
            this.zoom_in(cx);
        }))
        .on_action(cx.listener(|this, _: &ZoomOut, _, cx| {
            this.zoom_out(cx);
        }))
        .on_action(cx.listener(|this, _: &ZoomReset, _, cx| {
            this.zoom_reset(cx);
        }))
        // Freeze panes
        .on_action(cx.listener(|this, _: &FreezeTopRow, _, cx| {
            this.freeze_top_row(cx);
        }))
        .on_action(cx.listener(|this, _: &FreezeFirstColumn, _, cx| {
            this.freeze_first_column(cx);
        }))
        .on_action(cx.listener(|this, _: &FreezePanes, _, cx| {
            this.freeze_panes(cx);
        }))
        .on_action(cx.listener(|this, _: &UnfreezePanes, _, cx| {
            this.unfreeze_panes(cx);
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
        .on_action(cx.listener(|this, _: &ShowHistoryPanel, _, cx| {
            this.inspector_visible = true;
            this.inspector_tab = InspectorTab::History;
            cx.notify();
        }))
        // Edit mode cursor movement
        .on_action(cx.listener(|this, _: &EditCursorLeft, window, cx| {
            this.move_edit_cursor_left(cx);
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &EditCursorRight, window, cx| {
            this.move_edit_cursor_right(cx);
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &EditCursorHome, window, cx| {
            if this.mode.is_editing() {
                this.move_edit_cursor_home(cx);
                this.update_edit_scroll(window);
            } else {
                // Navigation mode: go to first column of current row
                this.view_state.selected.1 = 0;
                this.view_state.selection_end = None;
                this.view_state.scroll_col = 0;
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &EditCursorEnd, window, cx| {
            if this.mode.is_editing() {
                this.move_edit_cursor_end(cx);
                this.update_edit_scroll(window);
            } else {
                // Navigation mode: go to last column of current row
                this.view_state.selected.1 = crate::app::NUM_COLS - 1;
                this.view_state.selection_end = None;
                this.view_state.scroll_col = crate::app::NUM_COLS.saturating_sub(this.visible_cols());
                cx.notify();
            }
        }))
        // Edit mode word navigation (Alt+Arrow)
        .on_action(cx.listener(|this, _: &EditWordLeft, window, cx| {
            if this.mode.is_editing() {
                this.edit_selection_anchor = None;
                this.edit_cursor = this.prev_word_start(this.edit_cursor);
                this.update_edit_scroll(window);
                this.ensure_formula_bar_caret_visible(window);
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &EditWordRight, window, cx| {
            if this.mode.is_editing() {
                this.edit_selection_anchor = None;
                this.edit_cursor = this.next_word_end(this.edit_cursor);
                this.update_edit_scroll(window);
                this.ensure_formula_bar_caret_visible(window);
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &EditSelectWordLeft, window, cx| {
            if this.mode.is_editing() {
                if this.edit_selection_anchor.is_none() {
                    this.edit_selection_anchor = Some(this.edit_cursor);
                }
                this.edit_cursor = this.prev_word_start(this.edit_cursor);
                this.update_edit_scroll(window);
                this.ensure_formula_bar_caret_visible(window);
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &EditSelectWordRight, window, cx| {
            if this.mode.is_editing() {
                if this.edit_selection_anchor.is_none() {
                    this.edit_selection_anchor = Some(this.edit_cursor);
                }
                this.edit_cursor = this.next_word_end(this.edit_cursor);
                this.update_edit_scroll(window);
                this.ensure_formula_bar_caret_visible(window);
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
        .on_action(cx.listener(|this, _: &ToggleBold, window, cx| {
            this.toggle_bold(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &ToggleItalic, window, cx| {
            this.toggle_italic(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &ToggleUnderline, window, cx| {
            this.toggle_underline(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &FormatCurrency, window, cx| {
            this.format_currency(cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &FormatPercent, window, cx| {
            this.format_percent(cx);
            this.update_title_if_needed(window);
        }))
        // Background colors
        .on_action(cx.listener(|this, _: &ClearBackground, window, cx| {
            this.set_background_color(None, cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundYellow, window, cx| {
            this.set_background_color(Some([255, 255, 0, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundGreen, window, cx| {
            this.set_background_color(Some([198, 239, 206, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundBlue, window, cx| {
            this.set_background_color(Some([189, 215, 238, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundRed, window, cx| {
            this.set_background_color(Some([255, 199, 206, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundOrange, window, cx| {
            this.set_background_color(Some([255, 235, 156, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundPurple, window, cx| {
            this.set_background_color(Some([204, 192, 218, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundGray, window, cx| {
            this.set_background_color(Some([217, 217, 217, 255]), cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BackgroundCyan, window, cx| {
            this.set_background_color(Some([183, 222, 232, 255]), cx);
            this.update_title_if_needed(window);
        }))
        // Borders
        .on_action(cx.listener(|this, _: &BordersAll, window, cx| {
            this.apply_borders(BorderApplyMode::All, cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BordersOutline, window, cx| {
            this.apply_borders(BorderApplyMode::Outline, cx);
            this.update_title_if_needed(window);
        }))
        .on_action(cx.listener(|this, _: &BordersClear, window, cx| {
            this.apply_borders(BorderApplyMode::Clear, cx);
            this.update_title_if_needed(window);
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
        .on_action(cx.listener(|this, _: &FindReplace, _, cx| {
            this.show_find_replace(cx);
        }))
        .on_action(cx.listener(|this, _: &ReplaceNext, _, cx| {
            this.replace_next(cx);
        }))
        .on_action(cx.listener(|this, _: &ReplaceAll, _, cx| {
            this.replace_all(cx);
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
            let (row, col) = this.view_state.selected;
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
            let (row, col) = this.view_state.selected;
            this.show_precedents(row, col, cx);
        }))
        .on_action(cx.listener(|this, _: &RenameSymbol, _, cx| {
            this.show_rename_symbol(None, cx);
        }))
        .on_action(cx.listener(|this, _: &CreateNamedRange, _, cx| {
            this.show_create_named_range(cx);
        }))
        // Validation failure navigation (F8 / Shift+F8)
        .on_action(cx.listener(|this, _: &NextInvalidCell, _, cx| {
            this.next_invalid_cell(cx);
        }))
        .on_action(cx.listener(|this, _: &PrevInvalidCell, _, cx| {
            this.prev_invalid_cell(cx);
        }))
        // Command palette
        .on_action(cx.listener(|this, _: &ToggleCommandPalette, _, cx| {
            this.toggle_palette(cx);
        }))
        // View actions (for native macOS menus)
        .on_action(cx.listener(|this, _: &ShowAbout, _, cx| {
            this.show_about(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowLicense, _, cx| {
            this.show_license(cx);
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
        .on_action(cx.listener(|this, _: &CloseWindow, window, _cx| {
            // Commit any pending edit before closing
            this.commit_pending_edit();
            // Close the current window (not quit the app)
            window.remove_window();
        }))
        .on_action(cx.listener(|this, _: &Quit, _, cx| {
            // Commit any pending edit before quitting
            this.commit_pending_edit();
            // Propagate to global quit handler (saves session and quits)
            cx.propagate();
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
        // Alt accelerators (open Command Palette scoped to menu category)
        // These only fire when enabled via settings and in Navigation/Command mode
        .on_action(cx.listener(|this, _: &AltFile, _, cx| {
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::File, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltEdit, _, cx| {
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::Edit, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltView, _, cx| {
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::View, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltFormat, _, cx| {
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::Format, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltData, _, cx| {
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::Data, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltHelp, _, cx| {
            // Alt+H = Home (formatting) in modern Excel, not Help
            // Help is accessible via F1 or Command Palette
            if this.mode == Mode::Navigation || this.mode == Mode::Command {
                this.apply_menu_scope(MenuCategory::Format, cx);
            }
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

            // Handle filter dropdown keys
            if this.filter_dropdown_col.is_some() {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.close_filter_dropdown(cx);
                        return;
                    }
                    "enter" => {
                        this.apply_filter_dropdown(cx);
                        return;
                    }
                    "backspace" => {
                        this.filter_search_text.pop();
                        cx.notify();
                        return;
                    }
                    _ => {
                        // Type into search box
                        if let Some(ch) = &event.keystroke.key_char {
                            if !event.keystroke.modifiers.control && !event.keystroke.modifiers.alt {
                                this.filter_search_text.push_str(ch);
                                cx.notify();
                                return;
                            }
                        }
                    }
                }
            }

            // Handle validation dropdown keys (MUST come before edit-mode triggers)
            if this.is_validation_dropdown_open() {
                let modifiers = crate::validation_dropdown::KeyModifiers {
                    control: event.keystroke.modifiers.control,
                    alt: event.keystroke.modifiers.alt,
                    shift: event.keystroke.modifiers.shift,
                    platform: event.keystroke.modifiers.platform,
                };

                // Route key through dropdown first
                if this.route_dropdown_key_event(&event.keystroke.key, modifiers, cx) {
                    return;
                }

                // Handle character input for filtering
                if let Some(key_char) = &event.keystroke.key_char {
                    for ch in key_char.chars() {
                        if this.route_dropdown_char_event(ch, cx) {
                            return;
                        }
                    }
                }
            }

            // Cancel fill handle drag with Escape
            if this.is_fill_dragging() && event.keystroke.key == "escape" {
                this.cancel_fill_drag(cx);
                return;
            }

            // Exit zen mode with Escape
            if this.zen_mode && event.keystroke.key == "escape" {
                this.zen_mode = false;
                cx.notify();
                return;
            }

            // F1 hold-to-peek: show context help while F1 is held
            if event.keystroke.key == "f1" {
                this.f1_help_visible = true;
                cx.notify();
                return;
            }

            // Space hold-to-peek: preview workbook state before selected history entry
            if event.keystroke.key == "space"
                && !this.mode.is_editing()
                && !this.is_previewing()
                && this.selected_history_id.is_some()
                && this.history_highlight_range.is_some()
            {
                if let Err(e) = this.enter_preview(cx) {
                    this.status_message = Some(format!("Preview failed: {}", e));
                    cx.notify();
                }
                return;
            }

            // Space+Arrow scrubbing: while previewing, Up/Down navigates history entries
            // Makes the feature visceral: "scrub the timeline"
            if this.is_previewing() && (event.keystroke.key == "up" || event.keystroke.key == "down") {
                let direction = if event.keystroke.key == "up" { -1i32 } else { 1 };
                this.scrub_preview(direction, cx);
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
                return; // Consume all keystrokes in create named range mode
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
                            // Exit preview when closing inspector
                            if this.is_previewing() {
                                this.exit_preview(cx);
                            }
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
                } else if event.keystroke.key == "tab" {
                    // Tab toggles focus between find and replace inputs
                    this.find_toggle_focus(cx);
                    return;
                }
            }

            // Handle HubPasteToken mode
            if this.mode == Mode::HubPasteToken {
                if event.keystroke.key == "escape" {
                    this.hub_cancel_sign_in(cx);
                    return;
                } else if event.keystroke.key == "enter" {
                    this.hub_complete_sign_in(cx);
                    return;
                } else if event.keystroke.key == "backspace" {
                    this.hub_token_backspace(cx);
                    return;
                }
            }

            // Handle HubLink mode
            if this.mode == Mode::HubLink {
                if event.keystroke.key == "escape" {
                    this.hub_cancel_link(cx);
                    return;
                } else if event.keystroke.key == "backspace" {
                    this.hub_dataset_backspace(cx);
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
                            Mode::HubPasteToken => {
                                for c in printable_chars.chars() {
                                    this.hub_token_insert_char(c, cx);
                                }
                            }
                            Mode::HubLink => {
                                for c in printable_chars.chars() {
                                    this.hub_dataset_insert_char(c, cx);
                                }
                            }
                            _ => {
                                for c in printable_chars.chars() {
                                    this.insert_char(c, cx);
                                }
                                this.update_edit_scroll(window);
                            }
                        }
                    }
                }
            }
        }))
        // F1 hold-to-peek: hide help when F1 is released
        // Space hold-to-peek: exit preview when Space is released
        .on_key_up(cx.listener(|this, event: &KeyUpEvent, _, cx| {
            if event.keystroke.key == "f1" && this.f1_help_visible {
                this.f1_help_visible = false;
                cx.notify();
            }
            if event.keystroke.key == "space" && this.is_previewing() {
                this.exit_preview(cx);
            }
        }))
        // Mouse wheel scrolling (or zoom with Ctrl/Cmd)
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
            // Check for zoom modifier (Ctrl on Linux/Windows, Cmd on macOS)
            // macOS: use .platform (Cmd), others: use .control (Ctrl)
            #[cfg(target_os = "macos")]
            let zoom_modifier = event.modifiers.platform;
            #[cfg(not(target_os = "macos"))]
            let zoom_modifier = event.modifiers.control;

            let delta = event.delta.pixel_delta(px(CELL_HEIGHT));
            let dy: f32 = delta.y.into();

            if zoom_modifier {
                // Zoom: Ctrl/Cmd + wheel
                this.zoom_wheel(dy, cx);
                return; // Don't scroll
            }

            // Normal scrolling
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
        // On macOS with transparent titlebar, render a draggable title bar area
        // This allows window dragging and double-click to zoom
        .when(cfg!(target_os = "macos"), |d| {
            let titlebar_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);

            // Typography: Zed-style approach
            // Primary text: default text color (let macOS handle inactive dimming)
            // Secondary text: muted + slight fade for extra restraint
            let primary_color = app.token(TokenKey::TextPrimary);
            let secondary_color = app.token(TokenKey::TextMuted).opacity(0.85);

            let title_primary = app.document_meta.title_primary(app.history.is_dirty());
            let title_secondary = app.document_meta.title_secondary();

            // Chrome scrim: subtle gradient fade from titlebar into content
            let scrim_top = titlebar_bg.opacity(0.12);
            let scrim_bottom = titlebar_bg.opacity(0.0);

            // Check if we should show the default app prompt
            // Also check timer for success state auto-hide
            app.check_default_app_prompt_timer(cx);

            let show_default_prompt = app.should_show_default_app_prompt(cx);
            let prompt_state = app.default_app_prompt_state;
            let prompt_file_type = app.get_prompt_file_type();

            // Mark shown when transitioning from Hidden to visible
            // This records timestamp for cool-down and prevents session spam
            if show_default_prompt && prompt_state == crate::app::DefaultAppPromptState::Hidden {
                app.on_default_app_prompt_shown(cx);
            }

            d.child(
                div()
                    .id("macos-title-bar")
                    .w_full()
                    .h(px(34.0))  // Match Zed's titlebar height
                    .flex_shrink_0()
                    .bg(titlebar_bg)
                    .border_b_1()
                    .border_color(panel_border.opacity(0.5))  // Subtle hairline
                    .window_control_area(WindowControlArea::Drag)
                    // Double-click to zoom (macOS native behavior)
                    .on_click(|event, window, _cx| {
                        if event.click_count() == 2 {
                            window.titlebar_double_click();
                        }
                    })
                    // Left padding for traffic lights
                    .pl(px(72.0))
                    .pr(px(12.0))  // Right padding for prompt chip
                    .flex()
                    .items_center()
                    .justify_between()  // Push left content and right prompt apart
                    // Left side: document title
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap(px(6.0))
                            // Primary: filename + dirty indicator
                            .child(
                                div()
                                    .text_color(primary_color)
                                    .text_size(px(12.0))  // Smaller, calmer
                                    .child(title_primary)
                            )
                            // Secondary: provenance (quieter, extra small)
                            .when_some(title_secondary, |el, secondary| {
                                el.child(
                                    div()
                                        .text_color(secondary_color)
                                        .text_size(px(10.0))  // XSmall like Zed
                                        .child(secondary)
                                )
                            })
                    )
                    // Right side: default app prompt chip (state-based rendering)
                    .when(show_default_prompt, |el| {
                        use crate::app::DefaultAppPromptState;

                        let border_color = panel_border.opacity(0.2);
                        let muted = app.token(TokenKey::TextMuted);
                        let chip_bg = titlebar_bg.opacity(0.5);  // Muted background like inactive tab

                        // File type name for scoped messaging
                        let file_type_name = prompt_file_type
                            .map(|ft| ft.short_name())
                            .unwrap_or(".csv files");

                        match prompt_state {
                            // Success state: brief confirmation then auto-hide
                            DefaultAppPromptState::Success => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .px(px(8.0))
                                        .py(px(3.0))
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        .child(
                                            div()
                                                .text_color(muted)
                                                .text_size(px(10.0))
                                                .child("Default set")
                                        )
                                )
                            }

                            // Needs Settings state: user must complete in System Settings
                            DefaultAppPromptState::NeedsSettings => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        .child(
                                            div()
                                                .id("default-app-open-settings")
                                                .flex()
                                                .items_center()
                                                .gap(px(6.0))
                                                .px(px(8.0))
                                                .py(px(3.0))
                                                .cursor_pointer()
                                                .hover(|s| s.bg(panel_border.opacity(0.08)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    this.open_default_app_settings(cx);
                                                }))
                                                .child(
                                                    div()
                                                        .text_color(muted)
                                                        .text_size(px(10.0))
                                                        .child("Finish in System Settings")
                                                )
                                                .child(
                                                    div()
                                                        .text_color(muted.opacity(0.8))
                                                        .text_size(px(10.0))
                                                        .child("Open")
                                                )
                                        )
                                )
                            }

                            // Normal showing state: prompt to set default
                            _ => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        // Main clickable area
                                        .child(
                                            div()
                                                .id("default-app-set")
                                                .flex()
                                                .items_center()
                                                .gap(px(6.0))
                                                .px(px(8.0))
                                                .py(px(3.0))
                                                .cursor_pointer()
                                                .hover(|s| s.bg(panel_border.opacity(0.08)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    this.set_as_default_app(cx);
                                                }))
                                                // Label: "Open .csv files with VisiGrid"
                                                .child(
                                                    div()
                                                        .text_color(muted.opacity(0.8))
                                                        .text_size(px(10.0))
                                                        .child(format!("Open {} with VisiGrid", file_type_name))
                                                )
                                                // Button: subtle outline style
                                                .child(
                                                    div()
                                                        .text_color(muted)
                                                        .text_size(px(10.0))
                                                        .child("Make default")
                                                )
                                        )
                                        // Close button (dismiss forever)
                                        .child(
                                            div()
                                                .border_l_1()
                                                .border_color(border_color)
                                                .child(
                                                    div()
                                                        .id("default-app-dismiss")
                                                        .px(px(6.0))
                                                        .py(px(3.0))
                                                        .cursor_pointer()
                                                        .text_color(muted.opacity(0.4))  // Low contrast until hover
                                                        .text_size(px(10.0))
                                                        .hover(|s| s.bg(panel_border.opacity(0.08)).text_color(muted.opacity(0.8)))
                                                        .on_click(cx.listener(|this, _, _, cx| {
                                                            this.dismiss_default_app_prompt(cx);
                                                        }))
                                                        .child("\u{2715}")  // ✕
                                                )
                                        )
                                )
                            }
                        }
                    })
            )
            // Chrome scrim: subtle top fade for visual separation
            .child(
                div()
                    .w_full()
                    .h(px(8.0))  // Slightly smaller scrim
                    .flex_shrink_0()
                    .bg(linear_gradient(
                        180.0,
                        linear_color_stop(scrim_top, 0.0),
                        linear_color_stop(scrim_bottom, 1.0),
                    ))
            )
        })
        // Hide in-app menu bar on macOS (uses native menu bar instead)
        // Also hide in zen mode
        .when(!cfg!(target_os = "macos") && !zen_mode, |d| {
            d.child(menu_bar::render_menu_bar(app, cx))
        })
        .when(!zen_mode, |div| {
            div.child(formula_bar::render_formula_bar(app, window, cx))
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
        .when(show_validation_dialog, |div| {
            div.child(validation_dialog::render_validation_dialog(app, cx))
        })
        .when(show_rewind_confirm, |div| {
            div.child(render_rewind_confirm_dialog(app, cx))
        })
        .when(show_rewind_success, |div| {
            div.child(render_rewind_success_banner(app, cx))
        })
        .when(show_hub_paste_token, |div| {
            div.child(hub_dialogs::render_paste_token_dialog(app, cx))
        })
        .when(show_hub_link, |div| {
            div.child(hub_dialogs::render_link_dialog(app, cx))
        })
        .when(show_hub_publish_confirm, |div| {
            div.child(hub_dialogs::render_publish_confirm_dialog(app, cx))
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
            div.child(render_create_named_range_dialog(app, cx))
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
        // Filter dropdown popup
        .when_some(filter_dropdown::render_filter_dropdown(app, cx), |div, dropdown| {
            div.child(dropdown)
        })
        // Validation dropdown popup (list validation)
        .when_some(validation_dropdown_view::render_validation_dropdown(app, cx), |div, dropdown| {
            div.child(dropdown)
        })
        // Inspector panel (right-side drawer)
        .when(show_inspector, |div| {
            div.child(inspector_panel::render_inspector_panel(app, cx))
        })
        // NOTE: Autocomplete, signature help, and error banner popups are now rendered
        // in the grid overlay layer (grid.rs::render_popup_overlay) where they can be
        // positioned relative to the cell rect without menu/formula bar offset math.
        // Hover documentation popup (when not editing and hovering over formula bar)
        .when_some(app.hover_function.filter(|_| !app.mode.is_editing() && !app.autocomplete_visible), |div, func| {
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let accent = app.token(TokenKey::Accent);
            div.child(formula_bar::render_hover_docs(func, panel_bg, panel_border, text_primary, text_muted, accent))
        })
        // F1 hold-to-peek context help overlay
        .when(app.f1_help_visible, |div| {
            div.child(render_f1_help_overlay(app))
        })
}

/// Render the F1 hold-to-peek context help overlay
fn render_f1_help_overlay(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);

    // Build content based on context
    let content = if let Some(sig_info) = app.signature_help() {
        // In formula mode with a function: show full signature help
        let func = sig_info.function;
        let current_arg = sig_info.current_arg;

        let params: Vec<_> = func.parameters.iter().enumerate().map(|(i, param)| {
            let is_current = i == current_arg;
            div()
                .flex()
                .items_center()
                .gap_3()
                .py(px(4.0))
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(if is_current { accent } else { text_muted })
                        .font_weight(if is_current { FontWeight::SEMIBOLD } else { FontWeight::NORMAL })
                        .min_w(px(90.0))
                        .child(param.name)
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_disabled)
                        .child(param.description)
                )
        }).collect();

        div()
            .flex()
            .flex_col()
            // Header: function name
            .child(
                div()
                    .px_3()
                    .py(px(10.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .child(
                        div()
                            .text_size(px(14.0))
                            .text_color(text_primary)
                            .font_weight(FontWeight::SEMIBOLD)
                            .child(func.name)
                    )
            )
            // Signature
            .child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .text_size(px(13.0))
                    .text_color(text_muted)
                    .child(func.signature)
            )
            // Description
            .child(
                div()
                    .px_3()
                    .pb(px(8.0))
                    .text_size(px(12.0))
                    .text_color(text_disabled)
                    .child(func.description)
            )
            // Parameters
            .when(!func.parameters.is_empty(), |d| {
                d.child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .flex_col()
                        .children(params)
                )
            })
    } else if app.is_multi_selection() {
        // Multi-cell selection: show range stats
        let ((min_row, min_col), (max_row, max_col)) = app.selection_range();
        let start_ref = app.cell_ref_at(min_row, min_col);
        let end_ref = app.cell_ref_at(max_row, max_col);
        let range_ref = format!("{}:{}", start_ref, end_ref);

        // Calculate stats
        let mut count = 0usize;
        let mut numeric_count = 0usize;
        let mut sum = 0.0f64;
        let mut min_val: Option<f64> = None;
        let mut max_val: Option<f64> = None;

        for row in min_row..=max_row {
            for col in min_col..=max_col {
                let display = app.sheet().get_display(row, col);
                if !display.is_empty() {
                    count += 1;
                    // Try to parse as number (handles both values and formula results)
                    let clean = display.replace(',', "").replace('$', "").replace('%', "");
                    if let Ok(num) = clean.parse::<f64>() {
                        numeric_count += 1;
                        sum += num;
                        min_val = Some(min_val.map_or(num, |m| m.min(num)));
                        max_val = Some(max_val.map_or(num, |m| m.max(num)));
                    }
                }
            }
        }

        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);
        let average = if numeric_count > 0 { Some(sum / numeric_count as f64) } else { None };

        // Helper to format numbers with thousands separators
        let fmt_num = |n: f64| -> String {
            let base = if n.fract() == 0.0 {
                format!("{}", n as i64)
            } else {
                format!("{:.2}", n)
            };
            // Add thousands separators
            let parts: Vec<&str> = base.split('.').collect();
            let int_part = parts[0];
            let dec_part = parts.get(1);
            let negative = int_part.starts_with('-');
            let digits: String = int_part.chars().filter(|c| c.is_ascii_digit()).collect();
            let with_commas: String = digits
                .chars()
                .rev()
                .enumerate()
                .map(|(i, c)| if i > 0 && i % 3 == 0 { format!(",{}", c) } else { c.to_string() })
                .collect::<Vec<_>>()
                .into_iter()
                .rev()
                .collect();
            let result = if negative { format!("-{}", with_commas) } else { with_commas };
            if let Some(dec) = dec_part {
                format!("{}.{}", result, dec)
            } else {
                result
            }
        };

        let mut content = div()
            .flex()
            .flex_col()
            // Header: range reference
            .child(
                div()
                    .px_3()
                    .py(px(10.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .child(
                        div()
                            .text_size(px(14.0))
                            .text_color(text_primary)
                            .font_weight(FontWeight::SEMIBOLD)
                            .child(range_ref)
                    )
            )
            // Cell count
            .child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .justify_between()
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_muted)
                            .child("Cells")
                    )
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_primary)
                            .child(format!("{}", cell_count))
                    )
            );

        // Count (non-empty)
        if count > 0 {
            content = content.child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .justify_between()
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_muted)
                            .child("Count")
                    )
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_primary)
                            .child(format!("{}", count))
                    )
            );
        }

        // Numeric stats
        if numeric_count > 0 {
            content = content
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Sum")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(sum))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Average")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(average.unwrap()))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Min")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(min_val.unwrap()))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Max")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(max_val.unwrap()))
                        )
                );
        }

        content
    } else {
        // Single cell inspector
        let (row, col) = app.view_state.selected;
        let cell_ref = app.cell_ref_at(row, col);
        let raw_value = app.sheet().get_raw(row, col);
        let display_value = app.sheet().get_display(row, col);
        let is_formula = raw_value.starts_with('=');
        let format = app.sheet().get_format(row, col);

        // Get dependents (always useful)
        let dependents = inspector_panel::get_dependents(app, row, col);

        // Build format badges
        let mut format_badges: Vec<&str> = Vec::new();
        match &format.number_format {
            visigrid_engine::cell::NumberFormat::Number { .. } => format_badges.push("Number"),
            visigrid_engine::cell::NumberFormat::Currency { .. } => format_badges.push("Currency"),
            visigrid_engine::cell::NumberFormat::Percent { .. } => format_badges.push("Percent"),
            visigrid_engine::cell::NumberFormat::Date { .. } => format_badges.push("Date"),
            visigrid_engine::cell::NumberFormat::Time => format_badges.push("Time"),
            visigrid_engine::cell::NumberFormat::DateTime => format_badges.push("DateTime"),
            visigrid_engine::cell::NumberFormat::General => {}
        }
        if format.bold { format_badges.push("Bold"); }
        if format.italic { format_badges.push("Italic"); }
        if format.underline { format_badges.push("Underline"); }

        // Get precedents for formulas
        let precedents = if is_formula {
            inspector_panel::get_precedents(&raw_value)
        } else {
            Vec::new()
        };

        if is_formula {
            // Formula cell: value-first status card design
            // Goal: reassurance + confidence, not explanation

            // Get depth for complexity label (only when verified mode is on)
            let depth = if app.verified_mode {
                if let Some(report) = &app.last_recalc_report {
                    use visigrid_engine::cell_id::CellId;
                    let sheet_id = app.sheet().id;
                    let cell_id = CellId::new(sheet_id, row, col);
                    report.get_cell_info(&cell_id).map(|info| info.depth)
                } else {
                    None
                }
            } else {
                None
            };

            let complexity_label = match depth {
                Some(1) => "Simple formula".to_string(),
                Some(2) => "2 layers deep".to_string(),
                Some(d) if d <= 4 => format!("{} layers deep", d),
                Some(d) => format!("Complex ({} layers)", d),
                None => "Formula".to_string(),
            };

            // Softer divider color (reduced opacity)
            let divider = panel_border.opacity(0.5);

            let mut content = div()
                .flex()
                .flex_col()
                // Header: cell ref + context-aware verified badge
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(cell_ref.clone())
                        )
                        // No verification badges in F1 Peek - that's Pro Inspector territory
                        // F1 = "what is this cell", Pro Inspector = "why you can trust this"
                )
                // VALUE - the hero (large, prominent)
                .child(
                    div()
                        .px_3()
                        .py(px(12.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(18.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child(if display_value.is_empty() { "(empty)".to_string() } else { display_value.clone() })
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mt(px(2.0))
                                .child(complexity_label)
                        )
                )
                // Formula (secondary, smaller)
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(raw_value.clone())
                        )
                );

            // Uses (precedents) - human language
            if !precedents.is_empty() {
                let prec_refs: Vec<String> = precedents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let prec_text = if precedents.len() > 6 {
                    format!("{} +{} more", prec_refs.join(", "), precedents.len() - 6)
                } else {
                    prec_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Uses")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(prec_text)
                        )
                );
            }

            // Feeds (dependents) - human language
            if !dependents.is_empty() {
                let dep_refs: Vec<String> = dependents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let dep_text = if dependents.len() > 6 {
                    format!("{} +{} more", dep_refs.join(", "), dependents.len() - 6)
                } else {
                    dep_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Feeds")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(dep_text)
                        )
                );
            }

            // Format section (if any formatting applied)
            if !format_badges.is_empty() {
                let badges: Vec<_> = format_badges.iter().map(|label| {
                    div()
                        .px(px(8.0))
                        .py(px(3.0))
                        .bg(divider)
                        .rounded(px(4.0))
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .child(*label)
                }).collect();

                content = content.child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mb(px(6.0))
                                .child("Format")
                        )
                        .child(
                            div()
                                .flex()
                                .gap(px(6.0))
                                .children(badges)
                        )
                );
            }

            content
        } else {
            // Simple value cell: value-first compact card
            let type_label = if raw_value.is_empty() {
                "Empty cell"
            } else if raw_value.parse::<f64>().is_ok() {
                "Number"
            } else if raw_value == "TRUE" || raw_value == "FALSE" {
                "Boolean"
            } else {
                "Text"
            };

            // Softer divider for value cells too
            let divider = panel_border.opacity(0.5);

            let mut content = div()
                .flex()
                .flex_col()
                // Header: cell ref
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(cell_ref)
                        )
                )
                // VALUE - the hero
                .child(
                    div()
                        .px_3()
                        .py(px(12.0))
                        .child(
                            div()
                                .text_size(px(18.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child(if display_value.is_empty() { "(empty)".to_string() } else { display_value })
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mt(px(2.0))
                                .child(type_label)
                        )
                );

            // Feeds (dependents) - only if this value is used elsewhere
            if !dependents.is_empty() {
                let dep_refs: Vec<String> = dependents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let dep_text = if dependents.len() > 6 {
                    format!("{} +{} more", dep_refs.join(", "), dependents.len() - 6)
                } else {
                    dep_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Feeds")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(dep_text)
                        )
                );
            } else if raw_value.is_empty() {
                // Empty cell with no dependents - positive framing
                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Independent cell")
                        )
                );
            }

            content
        }
    };

    // Position overlay near the selection
    // Calculate position based on selection and scroll
    let ((_min_row, _min_col), (max_row, max_col)) = app.selection_range();

    // Calculate pixel position of selection end (bottom-right of selection)
    // Account for: header width, scroll position, cell dimensions
    let header_w = app.metrics.header_w;
    let header_h = app.metrics.header_h;
    let cell_w = app.metrics.cell_w;
    let cell_h = app.metrics.cell_h;

    // Menu bar + formula bar height (approximate)
    let top_offset = 24.0 + 32.0 + header_h; // menu + formula bar + column headers

    // X position: right edge of selection, offset from scroll
    let col_offset = (max_col as f32 - app.view_state.scroll_col as f32 + 1.0) * cell_w;
    let overlay_x = header_w + col_offset + 8.0; // 8px gap from selection

    // Y position: below the selection
    let row_offset = (max_row as f32 - app.view_state.scroll_row as f32 + 1.0) * cell_h;
    let overlay_y = top_offset + row_offset + 4.0; // 4px gap below selection

    // Clamp to reasonable bounds (don't go off screen)
    let overlay_x = overlay_x.max(header_w + 20.0);
    let overlay_y = overlay_y.max(top_offset + 20.0);

    div()
        .absolute()
        .inset_0()
        .child(
            div()
                .absolute()
                .left(px(overlay_x))
                .top(px(overlay_y))
                .w(px(240.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .child(content)
        )
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
fn render_create_named_range_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let has_error = app.create_name_validation_error.is_some();
    let name_focused = app.create_name_focus == CreateNameFocus::Name;
    let desc_focused = app.create_name_focus == CreateNameFocus::Description;

    // Centered dialog overlay - blocks clicks from reaching grid below
    div()
        .id("create-named-range-overlay")
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(|_this, _event, _window, _cx| {
            // Consume click to prevent it reaching grid below
        }))
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
                                .id("create-name-input")
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
                                .text_color(if app.create_name_name.is_empty() && !name_focused { text_muted } else { text_primary })
                                .cursor_text()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.create_name_focus = CreateNameFocus::Name;
                                    cx.notify();
                                }))
                                .flex()
                                .items_center()
                                .child(if app.create_name_name.is_empty() && !name_focused {
                                    "(required)".to_string()
                                } else {
                                    app.create_name_name.clone()
                                })
                                .when(name_focused, |d| {
                                    d.child(
                                        div()
                                            .w(px(1.0))
                                            .h(px(14.0))
                                            .bg(text_primary)
                                            .with_animation(
                                                "name-cursor-blink",
                                                Animation::new(Duration::from_millis(530))
                                                    .repeat()
                                                    .with_easing(pulsating_between(0.0, 1.0)),
                                                |this, delta| {
                                                    let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                                    this.opacity(opacity)
                                                },
                                            )
                                    )
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
                                .id("create-desc-input")
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if desc_focused { accent } else { panel_border })
                                .text_color(if app.create_name_description.is_empty() && !desc_focused { text_muted } else { text_primary })
                                .cursor_text()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.create_name_focus = CreateNameFocus::Description;
                                    cx.notify();
                                }))
                                .flex()
                                .items_center()
                                .child(if app.create_name_description.is_empty() && !desc_focused {
                                    "(optional)".to_string()
                                } else {
                                    app.create_name_description.clone()
                                })
                                .when(desc_focused, |d| {
                                    d.child(
                                        div()
                                            .w(px(1.0))
                                            .h(px(14.0))
                                            .bg(text_primary)
                                            .with_animation(
                                                "desc-cursor-blink",
                                                Animation::new(Duration::from_millis(530))
                                                    .repeat()
                                                    .with_easing(pulsating_between(0.0, 1.0)),
                                                |this, delta| {
                                                    let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                                    this.opacity(opacity)
                                                },
                                            )
                                    )
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

/// Render the rewind confirmation dialog (Phase 8C)
fn render_rewind_confirm_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let discard_count = app.rewind_confirm.discard_count;
    let target_summary = app.rewind_confirm.target_summary.clone();
    let sheet_name = app.rewind_confirm.sheet_name.clone();
    let location = app.rewind_confirm.location.clone();
    let replay_count = app.rewind_confirm.replay_count;
    let build_ms = app.rewind_confirm.build_ms;

    // Build location badge text: "Sheet1!A1:B10" or just "A1:B10" or just "Sheet1"
    let location_badge = match (&sheet_name, &location) {
        (Some(sheet), Some(loc)) => Some(format!("{}!{}", sheet, loc)),
        (Some(sheet), None) => Some(sheet.clone()),
        (None, Some(loc)) => Some(loc.clone()),
        (None, None) => None,
    };

    // Dark red colors for destructive action
    let danger_bg = hsla(0.0, 0.8, 0.3, 0.15);       // Dark red background
    let danger_border = hsla(0.0, 0.8, 0.4, 0.3);    // Dark red border
    let danger_button = hsla(0.0, 0.8, 0.4, 1.0);    // Red button
    let danger_button_hover = hsla(0.0, 0.8, 0.5, 1.0); // Brighter red on hover

    // Modal backdrop
    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            // Dialog box
            div()
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(420.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_4()
                // Header with warning icon
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(20.0))
                                .text_color(error_color)
                                .child("\u{26A0}") // Warning triangle
                        )
                        .child(
                            div()
                                .text_size(px(16.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Rewind History")
                        )
                )
                // Target info with location badge
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        // Location badge (if available)
                        .when_some(location_badge, |el, badge| {
                            el.child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .px_2()
                                            .py(px(2.0))
                                            .bg(accent.opacity(0.15))
                                            .border_1()
                                            .border_color(accent.opacity(0.3))
                                            .rounded_sm()
                                            .text_size(px(11.0))
                                            .font_weight(FontWeight::MEDIUM)
                                            .text_color(accent)
                                            .child(SharedString::from(badge))
                                    )
                                    .child(
                                        div()
                                            .text_size(px(11.0))
                                            .text_color(text_muted)
                                            .child("Target location")
                                    )
                            )
                        })
                        // Main message
                        .child(
                            div()
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .child(SharedString::from(format!(
                                    "This will permanently discard {} action{}.",
                                    discard_count,
                                    if discard_count == 1 { "" } else { "s" }
                                )))
                        )
                        // Target action summary
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!(
                                    "Rewind to state before: \"{}\"",
                                    target_summary
                                )))
                        )
                )
                // Destructive warning box
                .child(
                    div()
                        .px_3()
                        .py_2()
                        .bg(danger_bg)
                        .border_1()
                        .border_color(danger_border)
                        .rounded_sm()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(error_color)
                                .child("This action cannot be undone. The discarded changes will be permanently lost.")
                        )
                )
                // Performance info (subtle)
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted.opacity(0.7))
                        .child(SharedString::from(format!(
                            "Preview: {} actions replayed in {}ms",
                            replay_count,
                            build_ms
                        )))
                )
                // Buttons
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("rewind-cancel-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .bg(panel_border.opacity(0.3))
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .cursor_pointer()
                                .hover(|s| s.bg(panel_border.opacity(0.5)))
                                .child("Cancel")
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.cancel_rewind(cx);
                                }))
                        )
                        .child(
                            div()
                                .id("rewind-confirm-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .bg(danger_button)
                                .text_size(px(13.0))
                                .text_color(gpui::white())
                                .cursor_pointer()
                                .hover(|s| s.bg(danger_button_hover))
                                .child(SharedString::from(format!(
                                    "Discard {} action{}",
                                    discard_count,
                                    if discard_count == 1 { "" } else { "s" }
                                )))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.confirm_rewind(cx);
                                }))
                        )
                )
        )
}

/// Render the rewind success banner (Phase 8C)
/// Shows briefly after rewind with "Copy details" button
fn render_rewind_success_banner(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let success_color = hsla(0.35, 0.7, 0.5, 1.0); // Green

    let discarded_count = app.rewind_success.discarded_count;
    let target_summary = app.rewind_success.target_summary.clone();

    // Top banner that slides in from top
    div()
        .absolute()
        .top_0()
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .pt_2()
        .child(
            div()
                .id("rewind-success-banner")
                .bg(panel_bg)
                .border_1()
                .border_color(success_color.opacity(0.5))
                .rounded_md()
                .shadow_lg()
                .px_4()
                .py_2()
                .flex()
                .items_center()
                .gap_3()
                // Success icon
                .child(
                    div()
                        .text_size(px(16.0))
                        .text_color(success_color)
                        .child("\u{2713}") // Checkmark
                )
                // Message
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .child(
                            div()
                                .text_size(px(13.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text_primary)
                                .child(SharedString::from(format!(
                                    "Rewound. Discarded {} action{}. Undo not available.",
                                    discarded_count,
                                    if discarded_count == 1 { "" } else { "s" }
                                )))
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!(
                                    "Before: \"{}\"",
                                    target_summary
                                )))
                        )
                )
                // Copy audit button
                .child(
                    div()
                        .id("copy-rewind-audit-btn")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.3))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
                        .child("Copy audit")
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.copy_rewind_details(cx);
                        }))
                )
                // Dismiss button
                .child(
                    div()
                        .id("dismiss-rewind-banner-btn")
                        .px_2()
                        .py_1()
                        .text_size(px(14.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.text_color(text_primary))
                        .child("\u{2715}") // X mark
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.dismiss_rewind_banner(cx);
                        }))
                )
        )
}
