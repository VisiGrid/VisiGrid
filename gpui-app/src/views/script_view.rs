//! Full-panel script editor view.
//!
//! When open, replaces the grid slot in the main layout with a full-height
//! syntax-highlighted Lua code editor. The console panel stays below.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::scripting::{ConsoleTab, LuaTokenType, OutputEntry};
use crate::theme::TokenKey;

use super::code_render;

/// Render the full-panel script editor.
pub fn render_script_view(
    app: &mut Spreadsheet,
    _window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Prime caches (tokens + line offsets) for the script buffer
    app.script.buffer.tokens();

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);

    let line_height = 16.0_f32;
    let gutter_width = 36.0_f32;

    let buffer = &app.script.buffer;
    let input = &buffer.text;
    let cursor = buffer.cursor;
    let scroll_offset = buffer.scroll_offset;
    let tokens = buffer.cached_tokens();
    let line_offsets = buffer.cached_line_offsets();

    let cursor_line = input[..cursor].matches('\n').count();
    let total_lines = line_offsets.len();

    // Token color closure
    let token_color = |tt: LuaTokenType| -> Hsla {
        code_render::lua_token_color(app, tt)
    };

    // Build visible lines
    // We don't know actual available height at build time, so render a generous window.
    // Virtual scroll: show lines from scroll_offset to scroll_offset + estimated visible.
    // Use window height to estimate visible lines.
    let window_height: f32 = app.window_size.height.into();
    // Rough estimate: subtract toolbar (32), formula bar (30), format bar (28), column headers (24),
    // status bar (24), console if visible (lua_console.height), some padding.
    let console_h = if app.lua_console.visible { app.lua_console.height } else { 0.0 };
    let overhead = 32.0 + 30.0 + 28.0 + 24.0 + 24.0 + console_h + 40.0; // toolbar + extras
    let available = (window_height - overhead).max(100.0);
    let max_visible = (available / line_height).ceil() as usize;

    let visible_start = scroll_offset.min(total_lines.saturating_sub(1));
    let visible_end = (visible_start + max_visible + 2).min(total_lines); // +2 for partial lines

    let mut editor_lines = div()
        .flex_1()
        .overflow_hidden()
        .pt(px(4.0))
        .pb(px(4.0))
        .flex()
        .flex_col()
        .text_size(px(11.0))
        .font_family("monospace")
        .bg(editor_bg);

    for line_idx in visible_start..visible_end {
        let (line_start, line_end) = line_offsets[line_idx];
        let line_text = &input[line_start..line_end];
        let line_num = line_idx + 1;
        let is_cursor_line = line_idx == cursor_line;
        let cursor_in_line = if is_cursor_line {
            Some(cursor - line_start)
        } else {
            None
        };

        let line_el = code_render::render_code_line(
            line_text, line_start, line_end,
            tokens, cursor_in_line,
            line_num, is_cursor_line,
            gutter_width, line_height,
            text_primary, text_muted, accent,
            &token_color,
        );

        editor_lines = editor_lines.child(line_el);
    }

    div()
        .id("script-view")
        .key_context("ScriptView")
        .track_focus(&app.script_view_focus_handle)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            window.focus(&this.script_view_focus_handle, cx);
            cx.notify();
        }))
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            handle_script_view_key(this, event, window, cx);
            cx.stop_propagation();
        }))
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
            let delta = event.delta.pixel_delta(px(16.0));
            let dy: f32 = delta.y.into();
            let scroll_lines = (-dy / 16.0).round() as i32;
            if scroll_lines != 0 {
                let total_lines = this.script.buffer.text.split('\n').count();
                let current = this.script.buffer.scroll_offset as i32;
                let new_offset = (current + scroll_lines)
                    .max(0)
                    .min(total_lines.saturating_sub(1) as i32);
                this.script.buffer.scroll_offset = new_offset as usize;
                cx.notify();
            }
        }))
        .flex()
        .flex_col()
        .flex_1()
        .min_h(px(0.0))
        .bg(editor_bg)
        // Toolbar
        .child(
            render_toolbar(app, panel_bg, panel_border, text_primary, text_muted, accent, cx)
        )
        // Editor area (fills remaining space)
        .child(editor_lines)
}

/// Render the script view toolbar.
fn render_toolbar(
    app: &Spreadsheet,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .h(px(28.0))
        .w_full()
        .flex_shrink_0()
        .bg(panel_bg)
        .border_b_1()
        .border_color(panel_border)
        .flex()
        .items_center()
        .justify_between()
        .px_2()
        // Left: label + filename
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(8.0))
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child("Script")
                )
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .child("untitled.lua")
                )
        )
        // Right: Run + Close buttons
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(6.0))
                // Run button
                .child(
                    div()
                        .id("script-run-btn")
                        .px(px(8.0))
                        .py(px(2.0))
                        .bg(accent)
                        .rounded_sm()
                        .text_size(px(10.0))
                        .text_color(rgb(0xffffff))
                        .cursor_pointer()
                        .hover(|s| s.bg(accent.opacity(0.85)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            run_script(this, cx);
                        }))
                        .child("\u{25B6} Run")
                )
                // Close button
                .child(
                    div()
                        .id("script-close-btn")
                        .px(px(4.0))
                        .py(px(2.0))
                        .rounded_sm()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                            this.script.open = false;
                            if this.script.opened_from_console {
                                window.focus(&this.console_focus_handle, cx);
                            } else {
                                window.focus(&this.focus_handle, cx);
                            }
                            cx.notify();
                        }))
                        .child("\u{2715}")
                )
        )
}

/// Handle keyboard input in the script view.
fn handle_script_view_key(
    app: &mut Spreadsheet,
    event: &KeyDownEvent,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    let key = &event.keystroke.key;
    let ctrl = event.keystroke.modifiers.control;
    #[cfg(target_os = "macos")]
    let ctrl = ctrl || event.keystroke.modifiers.platform;

    match key.as_str() {
        "enter" => {
            if ctrl {
                // Ctrl+Enter: run script
                run_script(app, cx);
            } else {
                // Enter: insert newline
                app.script.buffer.insert("\n");
                app.script.buffer.ensure_cursor_visible(40);
            }
        }
        "escape" => {
            app.script.open = false;
            if app.script.opened_from_console {
                window.focus(&app.console_focus_handle, cx);
            } else {
                window.focus(&app.focus_handle, cx);
            }
        }
        "backspace" => {
            app.script.buffer.backspace();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "delete" => {
            app.script.buffer.delete();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "left" => {
            app.script.buffer.cursor_left();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "right" => {
            app.script.buffer.cursor_right();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "up" => {
            app.script.buffer.cursor_up();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "down" => {
            app.script.buffer.cursor_down();
            app.script.buffer.ensure_cursor_visible(40);
        }
        "home" => {
            if ctrl {
                app.script.buffer.cursor_buffer_home();
            } else {
                app.script.buffer.cursor_home();
            }
            app.script.buffer.ensure_cursor_visible(40);
        }
        "end" => {
            if ctrl {
                app.script.buffer.cursor_buffer_end();
            } else {
                app.script.buffer.cursor_end();
            }
            app.script.buffer.ensure_cursor_visible(40);
        }
        "pageup" => {
            // Move cursor up by ~38 lines (visible - 2)
            for _ in 0..38 {
                app.script.buffer.cursor_up();
            }
            app.script.buffer.ensure_cursor_visible(40);
        }
        "pagedown" => {
            for _ in 0..38 {
                app.script.buffer.cursor_down();
            }
            app.script.buffer.ensure_cursor_visible(40);
        }
        "l" if ctrl => {
            // Ctrl+L: clear console output
            app.lua_console.clear_output();
        }
        "tab" => {
            // Insert 2 spaces for indentation
            app.script.buffer.insert("  ");
            app.script.buffer.ensure_cursor_visible(40);
        }
        _ => {
            // Insert character
            if let Some(key_char) = &event.keystroke.key_char {
                if !ctrl && !event.keystroke.modifiers.alt {
                    app.script.buffer.insert(key_char);
                    app.script.buffer.ensure_cursor_visible(40);
                }
            }
        }
    }

    cx.notify();
}

/// Execute the script buffer content.
fn run_script(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    let script_text = app.script.buffer.text.clone();
    if script_text.trim().is_empty() {
        return;
    }

    // Auto-show console if hidden, switch to Run tab
    if !app.lua_console.visible {
        app.lua_console.show();
    }
    app.lua_console.active_tab = ConsoleTab::Run;

    // Begin group
    app.lua_console.begin_group();

    // Echo script input (truncated if long)
    let echo = truncate_script_echo(&script_text, 5);
    app.lua_console.push_output(OutputEntry::input(echo));

    // Delegate to existing execute body
    super::lua_console::execute_console_body(app, script_text, cx);

    // End group
    app.lua_console.end_group();
    cx.notify();
}

/// Truncate script text to first N lines + "..." for echoing in console output.
fn truncate_script_echo(text: &str, max_lines: usize) -> String {
    let lines: Vec<&str> = text.lines().collect();
    if lines.len() <= max_lines {
        text.to_string()
    } else {
        let mut result = lines[..max_lines].join("\n");
        result.push_str(&format!("\n... ({} more lines)", lines.len() - max_lines));
        result
    }
}
