pub mod command_palette;
mod find_dialog;
mod formula_bar;
mod goto_dialog;
mod grid;
mod headers;
mod menu_bar;
mod status_bar;

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CELL_HEIGHT};
use crate::actions::*;
use crate::mode::Mode;

pub fn render_spreadsheet(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let editing = app.mode.is_editing();
    let show_goto = app.mode == Mode::GoTo;
    let show_find = app.mode == Mode::Find;
    let show_command = app.mode == Mode::Command;

    div()
        .relative()
        .key_context("Spreadsheet")
        .track_focus(&app.focus_handle)
        // Navigation actions
        .on_action(cx.listener(|this, _: &MoveUp, _, cx| {
            this.move_selection(-1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &MoveDown, _, cx| {
            this.move_selection(1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &MoveLeft, _, cx| {
            this.move_selection(0, -1, cx);
        }))
        .on_action(cx.listener(|this, _: &MoveRight, _, cx| {
            this.move_selection(0, 1, cx);
        }))
        .on_action(cx.listener(|this, _: &JumpUp, _, cx| {
            this.jump_selection(-1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &JumpDown, _, cx| {
            this.jump_selection(1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &JumpLeft, _, cx| {
            this.jump_selection(0, -1, cx);
        }))
        .on_action(cx.listener(|this, _: &JumpRight, _, cx| {
            this.jump_selection(0, 1, cx);
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
        // Selection extension
        .on_action(cx.listener(|this, _: &ExtendUp, _, cx| {
            this.extend_selection(-1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendDown, _, cx| {
            this.extend_selection(1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendLeft, _, cx| {
            this.extend_selection(0, -1, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendRight, _, cx| {
            this.extend_selection(0, 1, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpUp, _, cx| {
            this.extend_jump_selection(-1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpDown, _, cx| {
            this.extend_jump_selection(1, 0, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpLeft, _, cx| {
            this.extend_jump_selection(0, -1, cx);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpRight, _, cx| {
            this.extend_jump_selection(0, 1, cx);
        }))
        .on_action(cx.listener(|this, _: &SelectAll, _, cx| {
            this.select_all(cx);
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
        }))
        .on_action(cx.listener(|this, _: &ConfirmEdit, _, cx| {
            this.confirm_edit(cx);
        }))
        .on_action(cx.listener(|this, _: &CancelEdit, _, cx| {
            if this.open_menu.is_some() {
                this.close_menu(cx);
            } else if this.mode == Mode::Command {
                this.hide_palette(cx);
            } else if this.mode == Mode::GoTo {
                this.hide_goto(cx);
            } else if this.mode == Mode::Find {
                this.hide_find(cx);
            } else {
                this.cancel_edit(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &TabNext, _, cx| {
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
            this.delete_char(cx);
        }))
        .on_action(cx.listener(|this, _: &FillDown, _, cx| {
            this.fill_down(cx);
        }))
        .on_action(cx.listener(|this, _: &FillRight, _, cx| {
            this.fill_right(cx);
        }))
        .on_action(cx.listener(|this, _: &ConfirmEditInPlace, _, cx| {
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
        // Command palette
        .on_action(cx.listener(|this, _: &ToggleCommandPalette, _, cx| {
            this.toggle_palette(cx);
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
        // Character input (handles editing, goto, find, and command modes)
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
            // Handle Command Palette mode
            if this.mode == Mode::Command {
                match event.keystroke.key.as_str() {
                    "escape" => {
                        this.hide_palette(cx);
                        return;
                    }
                    "enter" => {
                        this.palette_execute(cx);
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
        // Mouse move for resize dragging
        .on_mouse_move(cx.listener(|this, event: &MouseMoveEvent, _, cx| {
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
        }))
        // Mouse up to end resize
        .on_mouse_up(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
                this.resizing_col = None;
                this.resizing_row = None;
                cx.notify();
            }
        }))
        .flex()
        .flex_col()
        .size_full()
        .bg(rgb(0x1e1e1e))
        .child(menu_bar::render_menu_bar(app, cx))
        .child(formula_bar::render_formula_bar(app))
        .child(headers::render_column_headers(app, cx))
        .child(grid::render_grid(app, cx))
        .child(status_bar::render_status_bar(app, editing))
        .when(show_goto, |div| {
            div.child(goto_dialog::render_goto_dialog(app))
        })
        .when(show_find, |div| {
            div.child(find_dialog::render_find_dialog(app))
        })
        .when(show_command, |div| {
            div.child(command_palette::render_command_palette(app, cx))
        })
        .when(app.open_menu.is_some(), |div| {
            div.child(menu_bar::render_menu_dropdown(app, cx))
        })
}
