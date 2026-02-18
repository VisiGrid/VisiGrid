use gpui::*;
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::actions::*;
use crate::mode::Mode;

pub(crate) fn bind(
    el: Div,
    cx: &mut Context<Spreadsheet>,
) -> Div {
    el
        // Clipboard actions
        .on_action(cx.listener(|this, _: &Copy, window, cx| {
            // Terminal handles its own copy (Cmd+C on Mac, Ctrl+Shift+C on Linux)
            if this.terminal_has_focus(window) { return; }
            // Script view: copy script content
            if this.script.open {
                let text = &this.script.buffer.text;
                if !text.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(text.clone()));
                }
                return;
            }
            if this.lua_console.visible {
                let input = &this.lua_console.input_buffer.text;
                if !input.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(input.clone()));
                }
                return;
            }
            this.copy(cx);
        }))
        .on_action(cx.listener(|this, _: &Cut, window, cx| {
            // Terminal handles its own input
            if this.terminal_has_focus(window) { return; }
            // Script view: cut script content
            if this.script.open {
                let text = &this.script.buffer.text;
                if !text.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(text.clone()));
                    this.script.buffer.clear();
                    cx.notify();
                }
                return;
            }
            if this.lua_console.visible {
                let input = &this.lua_console.input_buffer.text;
                if !input.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(input.clone()));
                    this.lua_console.input_buffer.clear();
                    cx.notify();
                }
                return;
            }
            this.cut(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &Paste, window, cx| {
            // Terminal handles its own paste (Cmd+V on Mac, Ctrl+Shift+V on Linux)
            if this.terminal_has_focus(window) { return; }
            // Script view: paste into script buffer
            if this.script.open {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        this.script.buffer.insert(&text);
                        this.script.buffer.ensure_cursor_visible(40);
                    }
                }
                cx.notify();
                return;
            }
            // Lua console: paste into console input
            if this.lua_console.visible {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        this.lua_console.insert(&text);
                    }
                }
                cx.notify();
                return;
            }
            // AI dialogs handle their own paste
            if this.mode == Mode::AISettings {
                this.ai_settings_paste(cx);
                return;
            }
            if this.mode == Mode::AiDialog {
                this.ask_ai_paste(cx);
                return;
            }
            // Sheet rename: paste text
            if this.renaming_sheet.is_some() {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        // Filter out control chars and newlines
                        for c in text.chars().filter(|c| !c.is_control() && *c != '\n' && *c != '\r') {
                            this.sheet_rename_input_char(c, cx);
                        }
                    }
                }
                return;
            }
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
            // Color picker: paste into hex input
            if this.mode == Mode::ColorPicker {
                this.color_picker_paste(cx);
                return;
            }
            // Find/Replace dialog: paste into find or replace input
            if this.mode == Mode::Find {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        // Filter out newlines but allow other chars
                        for c in text.chars().filter(|c| *c != '\n' && *c != '\r') {
                            this.find_insert_char(c, cx);
                        }
                    }
                }
                return;
            }
            // GoTo dialog: paste into input
            if this.mode == Mode::GoTo {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        for c in text.chars().filter(|c| !c.is_control()) {
                            this.goto_insert_char(c, cx);
                        }
                    }
                }
                return;
            }
            this.paste(cx);
            this.update_edit_scroll(window);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &PasteValues, window, cx| {
            this.paste_values(cx);
            this.update_edit_scroll(window);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &PasteSpecial, _, cx| {
            this.show_paste_special(cx);
        }))
        .on_action(cx.listener(|this, _: &PasteFormulas, window, cx| {
            this.paste_formulas(cx);
            this.update_edit_scroll(window);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &PasteFormats, window, cx| {
            this.paste_formats(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &DeleteCell, window, cx| {
            if this.guard_terminal_focus(window, cx, "DeleteCell") { return; }
            // Format bar editing consumes actions before Spreadsheet editing.
            if this.ui.format_bar.size_editing { return; }
            if !this.mode.is_editing() {
                this.delete_selection(cx);
                this.update_title_if_needed(window, cx);
            }
        }))
        // Insert/Delete rows/columns (Ctrl+= / Ctrl+-)
        .on_action(cx.listener(|this, _: &InsertRowsOrCols, window, cx| {
            if !this.mode.is_editing() {
                this.insert_rows_or_cols(cx);
                this.update_title_if_needed(window, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteRowsOrCols, window, cx| {
            if !this.mode.is_editing() {
                this.delete_rows_or_cols(cx);
                this.update_title_if_needed(window, cx);
            }
        }))
        // Hide/Unhide rows/columns (Ctrl+9/0, Ctrl+Shift+9/0)
        .on_action(cx.listener(|this, _: &HideRows, _, cx| {
            this.hide_rows(cx);
        }))
        .on_action(cx.listener(|this, _: &UnhideRows, _, cx| {
            this.unhide_rows(cx);
        }))
        .on_action(cx.listener(|this, _: &HideCols, _, cx| {
            this.hide_cols(cx);
        }))
        .on_action(cx.listener(|this, _: &UnhideCols, _, cx| {
            this.unhide_cols(cx);
        }))
        // Undo/Redo
        .on_action(cx.listener(|this, _: &Undo, window, cx| {
            this.undo(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &Redo, window, cx| {
            this.redo(cx);
            this.update_title_if_needed(window, cx);
        }))
        // Editing actions
        // F2: Toggle Navigation ↔ Edit, or toggle Caret/Point in Formula mode
        .on_action(cx.listener(|this, _: &StartEdit, window, cx| {
            if this.guard_terminal_focus(window, cx, "StartEdit") { return; }
            if this.mode.is_formula() {
                // In Formula mode: F2 toggles between Caret and Point submode
                this.toggle_formula_nav_mode(cx);
                return;
            }
            if this.mode.is_editing() {
                // Edit → Navigation: cancel without committing (same as Escape)
                this.cancel_edit(cx);
                return;
            }
            this.start_edit(cx);
            this.update_edit_scroll(window);
            // On macOS, show tip about enabling F2 (catches Ctrl+U and menu-driven edit)
            this.maybe_show_f2_tip(cx);
        }))
        .on_action(cx.listener(|this, _: &ConfirmEdit, window, cx| {
            if this.guard_terminal_focus(window, cx, "ConfirmEdit") { return; }
            // Close-confirm dialog: Enter activates focused button
            if this.close_confirm_visible {
                match this.close_confirm_focused {
                    0 => {
                        this.close_confirm_visible = false;
                        cx.notify();
                    }
                    1 => {
                        this.close_confirm_visible = false;
                        this.prepare_close(cx);
                        window.remove_window();
                    }
                    _ => {
                        this.close_confirm_visible = false;
                        let saved = this.save_and_close(cx);
                        if saved {
                            this.prepare_close(cx);
                            window.remove_window();
                        }
                    }
                }
                return;
            }
            if this.open_menu.is_some() {
                this.menu_execute_highlighted(window, cx);
                return;
            }
            // Format bar editing consumes actions before Spreadsheet editing.
            // gpui dispatches keybinding actions before on_key_down, so the
            // format bar's own key handler never sees Enter/Esc/Backspace.
            // These guards are the ONLY way format bar input works.
            if this.ui.format_bar.size_editing {
                crate::views::format_bar::commit_font_size(this, cx);
                return;
            }
            // Name box (cell selector) editing
            if this.name_box_editing {
                this.confirm_name_box(cx);
                return;
            }
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AiDialog) {
                return;
            }
            // Sheet rename: Enter confirms
            if this.renaming_sheet.is_some() {
                this.confirm_sheet_rename(cx);
                return;
            }
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
                this.update_title_if_needed(window, cx);
                return;
            }
            // Handle Enter key based on current mode
            match this.mode {
                Mode::ColorPicker => this.color_picker_execute(window, cx),
                Mode::ThemePicker => this.theme_picker_execute(window, cx),
                Mode::FontPicker => this.font_picker_execute(cx),
                Mode::Command => this.palette_execute(window, cx),
                Mode::GoTo => this.confirm_goto(cx),
                Mode::CreateNamedRange => this.confirm_create_named_range(cx),
                _ => {
                    this.confirm_edit_enter(cx);
                    this.update_title_if_needed(window, cx);
                }
            }
        }))
        .on_action(cx.listener(|this, _: &ConfirmEditUp, window, cx| {
            if this.guard_terminal_focus(window, cx, "ConfirmEditUp") { return; }
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
                    this.confirm_edit_up_enter(cx);
                    this.update_title_if_needed(window, cx);
                }
            }
        }))
        .on_action(cx.listener(|this, _: &CancelEdit, window, cx| {
            if this.guard_terminal_focus(window, cx, "CancelEdit") { return; }
            // Close-confirm dialog: Escape dismisses
            if this.close_confirm_visible {
                this.close_confirm_visible = false;
                cx.notify();
                return;
            }
            // Format bar editing consumes actions before Spreadsheet editing.
            // See ConfirmEdit guard above for rationale.
            if this.ui.format_bar.size_editing {
                this.ui.format_bar.size_editing = false;
                this.ui.format_bar.size_dropdown = false;
                this.ui.format_bar.size_replace_next = false;
                cx.notify();
                return;
            }
            // Name box (cell selector) editing
            if this.name_box_editing {
                this.cancel_name_box_edit(cx);
                return;
            }
            // Sheet rename: Escape cancels
            if this.renaming_sheet.is_some() {
                this.cancel_sheet_rename(cx);
                return;
            }
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
            } else if this.mode == Mode::ColorPicker {
                this.hide_color_picker(cx);
            } else if this.mode == Mode::FormatPainter {
                this.cancel_format_painter(cx);
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
            } else if this.mode == Mode::ExplainDiff {
                this.close_explain_diff(cx);
            } else if this.history_context_menu_entry_id.is_some() {
                this.hide_history_context_menu(cx);
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
            } else if this.profiler_visible && this.mode == Mode::Navigation {
                // Esc closes profiler panel when in navigation mode
                this.profiler_visible = false;
                window.focus(&this.focus_handle, cx);
                cx.notify();
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
            } else if this.clipboard_visual_range.is_some() && this.mode == Mode::Navigation {
                // Esc clears copy/cut border overlay
                this.clipboard_visual_range = None;
                cx.notify();
            } else {
                this.cancel_edit(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabNext, window, cx| {
            if this.guard_terminal_focus(window, cx, "TabNext") { return; }
            // Close-confirm dialog traps Tab
            if this.close_confirm_visible {
                this.close_confirm_focused = (this.close_confirm_focused + 1) % 3;
                cx.notify();
                return;
            }
            // Format bar editing consumes actions before Spreadsheet editing.
            // See ConfirmEdit guard above for rationale.
            if this.ui.format_bar.size_editing {
                crate::views::format_bar::commit_font_size(this, cx);
                window.focus(&this.focus_handle, cx);
                return;
            }
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AiDialog) {
                return;
            }
            // Dialog modes handle Tab themselves
            if this.mode == Mode::CreateNamedRange {
                this.create_name_tab(cx);
                return;
            }
            // If autocomplete is visible, Tab accepts the suggestion
            if this.autocomplete_visible {
                this.autocomplete_accept(cx);
                this.update_title_if_needed(window, cx);
                return;
            }
            if this.mode.is_editing() {
                this.confirm_edit_and_tab_right(cx);
                this.update_title_if_needed(window, cx);
            } else {
                // Nav mode: Tab also sets origin for tab-chain return
                if this.tab_chain_origin_col.is_none() {
                    this.tab_chain_origin_col = Some(this.view_state.selected.1);
                }
                this.move_selection(0, 1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabPrev, window, cx| {
            if this.guard_terminal_focus(window, cx, "TabPrev") { return; }
            // Close-confirm dialog traps Shift+Tab
            if this.close_confirm_visible {
                this.close_confirm_focused = if this.close_confirm_focused == 0 { 2 } else { this.close_confirm_focused - 1 };
                cx.notify();
                return;
            }
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AiDialog) {
                return;
            }
            if this.mode.is_editing() {
                this.confirm_edit_and_tab_left(cx);
                this.update_title_if_needed(window, cx);
            } else {
                // Nav mode: Shift+Tab also sets origin for tab-chain return
                if this.tab_chain_origin_col.is_none() {
                    this.tab_chain_origin_col = Some(this.view_state.selected.1);
                }
                this.move_selection(0, -1, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &BackspaceChar, window, cx| {
            if this.guard_terminal_focus(window, cx, "BackspaceChar") { return; }
            // Format bar editing consumes actions before Spreadsheet editing.
            // See ConfirmEdit guard above for rationale.
            if this.ui.format_bar.size_editing {
                this.ui.format_bar.size_replace_next = false;
                this.ui.format_bar.size_input.pop();
                cx.notify();
                return;
            }
            // Name box (cell selector) editing
            if this.name_box_editing {
                this.name_box_backspace(cx);
                return;
            }
            // AI dialogs handle their own backspace
            if this.mode == Mode::AISettings {
                this.ai_settings_backspace(cx);
                return;
            }
            if this.mode == Mode::AiDialog {
                this.ask_ai_backspace(cx);
                return;
            }
            // Sheet rename: handle backspace
            if this.renaming_sheet.is_some() {
                this.sheet_rename_backspace(cx);
                return;
            }
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
            if this.mode == Mode::ColorPicker {
                this.color_picker_handle_key("backspace", None, false, window, cx);
            } else if this.mode == Mode::Command {
                this.palette_backspace(cx);
            } else if this.mode == Mode::GoTo {
                this.goto_backspace(cx);
            } else if this.mode == Mode::Find {
                this.find_backspace(cx);
            } else if this.mode == Mode::ThemePicker {
                this.theme_picker_backspace(cx);
            } else if this.mode == Mode::FontPicker {
                this.font_picker_backspace(cx);
            } else if this.mode == Mode::RenameSymbol {
                this.rename_symbol_backspace(cx);
            } else if this.mode == Mode::EditDescription {
                this.edit_description_backspace(cx);
            } else if this.mode == Mode::ExtractNamedRange {
                match this.extract_focus {
                    CreateNameFocus::Name => this.extract_name_backspace(cx),
                    CreateNameFocus::Description => this.extract_description_backspace(cx),
                }
            } else if this.mode == Mode::License {
                this.license_backspace(cx);
            } else if this.mode.is_editing() {
                this.backspace(cx);
                this.update_edit_scroll(window);
            } else if this.mode == Mode::Navigation {
                // Single text cell: enter edit mode and delete last char
                let is_single_cell = this.view_state.selection_end.is_none()
                    || this.view_state.selection_end == Some(this.view_state.selected);
                let is_text = {
                    use visigrid_engine::cell::CellValue;
                    let (row, col) = this.view_state.selected;
                    matches!(this.sheet(cx).get_cell(row, col).value, CellValue::Text(_))
                };

                if is_single_cell && is_text {
                    this.start_edit(cx);
                    if !this.edit_value.is_empty() {
                        this.backspace(cx);
                    }
                    this.update_edit_scroll(window);
                } else {
                    this.delete_selection(cx);
                    this.update_title_if_needed(window, cx);
                }
            } else {
                this.delete_selection(cx);
                this.update_title_if_needed(window, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &DeleteChar, window, cx| {
            if this.guard_terminal_focus(window, cx, "DeleteChar") { return; }
            // Format bar editing consumes actions before Spreadsheet editing.
            if this.ui.format_bar.size_editing { return; }
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AiDialog) {
                return;
            }
            // Sheet rename: handle delete
            if this.renaming_sheet.is_some() {
                this.sheet_rename_delete(cx);
                return;
            }
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
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &FillRight, window, cx| {
            this.fill_right(cx);
            this.update_title_if_needed(window, cx);
        }))
        // Insert date/time and copy from above
        .on_action(cx.listener(|this, _: &InsertDate, window, cx| {
            this.insert_date(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &InsertTime, window, cx| {
            this.insert_time(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &CopyFormulaAbove, window, cx| {
            this.copy_formula_above(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &CopyValueAbove, window, cx| {
            this.copy_value_above(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &InsertNewline, _, cx| {
            this.insert_newline(cx);
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
            // Ask AI dialog: Cmd/Ctrl+Enter submits
            if this.mode == Mode::AiDialog {
                this.ask_ai_submit(cx);
                return;
            }
            // Lua console handles its own Ctrl+Enter (execute)
            if this.lua_console.visible {
                crate::views::lua_console::execute_console(this, cx);
                return;
            }
            this.confirm_edit_in_place(cx);
        }))
}
