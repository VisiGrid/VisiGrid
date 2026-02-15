//! Lua console panel view.
//!
//! A REPL panel for executing Lua scripts. Shows at the bottom of the window
//! when toggled with Ctrl+Shift+L.
//!
//! ## Virtual Scroll
//!
//! The console uses virtual scrolling to handle large outputs efficiently.
//! PageUp/PageDown navigate through output history. Ctrl+Home/End jump to
//! start/end. The view auto-scrolls to bottom when new output arrives.

use std::sync::atomic::Ordering;
use std::sync::mpsc::TryRecvError;

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::actions::*;
use crate::app::Spreadsheet;
use crate::scripting::{
    ConsoleState, ConsoleTab, DebugAction, DebugConfig, DebugEventPayload, DebugSessionState,
    LuaCellValue, LuaEvalResult, LuaOp, LuaTokenType, OutputEntry, OutputKind, SheetSnapshot,
    VarPathSegment, MAX_CONSOLE_HEIGHT, DEBUG_OUTPUT_CAP, spawn_debug_session, CONSOLE_SOURCE,
};
use crate::scripting::examples::{EXAMPLES, get_example, find_example};
use crate::theme::TokenKey;
use crate::ui::render_locked_feature_panel;

/// Render the Lua console panel (if visible)
pub fn render_lua_console(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let console = &app.lua_console;

    if !console.visible {
        return div().into_any_element();
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);
    let editor_bg = app.token(TokenKey::EditorBg);

    // Viewport-relative max height: cap at 60% of window height
    let window_height: f32 = app.window_size.height.into();
    let effective_max = if window_height > 0.0 {
        MAX_CONSOLE_HEIGHT.min(window_height * 0.6)
    } else {
        MAX_CONSOLE_HEIGHT  // Window size not yet known; use absolute max
    };

    let console_height = if console.is_maximized {
        effective_max
    } else {
        console.height
    };

    let current_tab = console.active_tab;
    let is_maximized = console.is_maximized;
    let has_output = !console.output.is_empty();

    div()
        .id("lua-console-panel")
        .key_context("LuaConsole")
        .track_focus(&app.console_focus_handle)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            window.focus(&this.console_focus_handle, cx);
            cx.notify();
        }))
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            handle_console_key_from_main(this, event, window, cx);
            cx.stop_propagation();
        }))
        .flex_shrink_0()
        .h(px(console_height))
        .bg(panel_bg)
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .child(
            // Resize handle at top
            div()
                .id("lua-console-resize")
                .h(px(4.0))
                .w_full()
                .cursor(CursorStyle::ResizeUpDown)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, event: &MouseDownEvent, _, cx| {
                    this.lua_console.resizing = true;
                    this.lua_console.resize_start_y = event.position.y.into();
                    this.lua_console.resize_start_height = this.lua_console.height;
                    cx.notify();
                }))
        )
        .child(
            // Tab bar + toolbar header
            render_console_tab_bar(
                current_tab, is_maximized, has_output,
                text_primary, text_muted, selection_bg, accent, panel_border,
                cx,
            )
        )
        .when(current_tab == ConsoleTab::Run, |d| {
            d.child(
                // Run tab: output area with virtual scroll
                render_run_tab_content(console, text_primary, text_muted, accent, error_color, panel_border, editor_bg, cx)
            )
            .child(
                // Input area
                render_input_bar(app, panel_border, editor_bg, text_primary, accent, cx)
            )
        })
        .when(current_tab == ConsoleTab::Debug, |d| {
            d.child(
                render_debug_tab_content(app, text_primary, text_muted, accent, error_color, panel_border, editor_bg, cx)
            )
        })
        .into_any_element()
}

/// Render the Run tab content area (output + scroll indicator)
fn render_run_tab_content(
    console: &crate::scripting::ConsoleState,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    error_color: Hsla,
    panel_border: Hsla,
    editor_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex_1()
        .flex()
        .flex_col()
        .overflow_hidden()
        .child(
            div()
                .flex_1()
                .overflow_hidden()
                .px_2()
                .py_1()
                .when(console.output.is_empty(), |d| {
                    d.child(
                        div()
                            .flex()
                            .flex_col()
                            .gap(px(2.0))
                            .py_2()
                            .text_xs()
                            .text_color(text_muted.opacity(0.5))
                            .child("Enter run \u{00B7} Shift+Enter newline \u{00B7} Ctrl+L clear")
                            .child("help \u{2192} commands \u{00B7} examples \u{2192} scripts")
                    )
                })
                .children({
                    let visible = console.visible_output();
                    let base = console.view_start;
                    let prev_gid = if base > 0 { console.group_id_at(base - 1) } else { 0 };

                    visible.iter().enumerate().map(move |(i, entry)| {
                        let prev = if i == 0 { prev_gid } else { visible[i - 1].group_id };
                        let is_group_start = entry.group_id != 0
                            && entry.group_id != prev
                            && (base + i) > 0;
                        render_output_entry_grouped(
                            entry, "run", base + i, is_group_start,
                            text_primary, text_muted, accent, error_color, panel_border,
                        )
                    }).collect::<Vec<_>>()
                })
        )
        .when(console.scroll_info().is_some(), |d| {
            d.child(
                div()
                    .h(px(16.0))
                    .px_2()
                    .flex()
                    .items_center()
                    .justify_between()
                    .border_t_1()
                    .border_color(panel_border)
                    .bg(editor_bg)
                    .child(
                        div()
                            .text_xs()
                            .text_color(text_muted)
                            .child(console.scroll_info().unwrap_or_default())
                    )
                    .child(
                        div()
                            .flex()
                            .gap_1()
                            .child(
                                scroll_button("\u{25B2}", console.can_scroll_up(), text_muted, text_primary, cx, |this, cx| {
                                    this.lua_console.scroll_page_up();
                                    cx.notify();
                                })
                            )
                            .child(
                                scroll_button("\u{25BC}", console.can_scroll_down(), text_muted, text_primary, cx, |this, cx| {
                                    this.lua_console.scroll_page_down();
                                    cx.notify();
                                })
                            )
                    )
            )
        })
}

/// Render the input bar (elevated mini-editor with line numbers and syntax highlighting)
fn render_input_bar(
    app: &Spreadsheet,
    panel_border: Hsla,
    editor_bg: Hsla,
    text_primary: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let text_muted = app.token(TokenKey::TextMuted);
    let line_count = app.lua_console.input_buffer.text.lines().count().max(1);
    let line_height = 16.0_f32;
    let padding = 12.0_f32;
    let max_lines = 12;
    let visible_lines = line_count.min(max_lines);
    let bar_height = (visible_lines as f32) * line_height + padding;

    div()
        .h(px(bar_height))
        .mx_2()
        .mb_1()
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .bg(editor_bg)
        .overflow_hidden()
        .child(
            render_input_area(app, editor_bg, text_primary, text_muted, accent, line_height, max_lines, cx)
        )
}

/// Render the tab bar with toolbar buttons
fn render_console_tab_bar(
    current_tab: ConsoleTab,
    is_maximized: bool,
    has_output: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    _accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let run_active = current_tab == ConsoleTab::Run;
    let debug_active = current_tab == ConsoleTab::Debug;

    div()
        .flex()
        .items_center()
        .justify_between()
        .border_b_1()
        .border_color(panel_border)
        // Left side: tabs
        .child(
            div()
                .flex()
                .child(
                    render_tab_button(
                        "console-tab-run", "Run", run_active,
                        text_primary, text_muted, selection_bg, panel_border, cx,
                        |this, cx| {
                            this.lua_console.active_tab = ConsoleTab::Run;
                            cx.notify();
                        },
                    )
                )
                .child(
                    render_tab_button(
                        "console-tab-debug", "Debug", debug_active,
                        text_primary, text_muted, selection_bg, panel_border, cx,
                        |this, cx| {
                            this.lua_console.active_tab = ConsoleTab::Debug;
                            cx.notify();
                        },
                    )
                )
        )
        // Right side: toolbar buttons
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(2.0))
                .pr_1()
                .child(
                    console_toolbar_btn(
                        "console-expand-btn",
                        "\u{2197}",  // ↗ expand arrow
                        true,
                        text_muted, text_primary, panel_border,
                        cx,
                        |this, cx| {
                            this.script.open = true;
                            this.script.opened_from_console = true;
                            // Copy REPL input if script is empty and input is worth expanding
                            if this.script.buffer.is_empty()
                                && !this.lua_console.input_buffer.text.is_empty()
                                && (this.lua_console.input_buffer.text.contains('\n')
                                    || this.lua_console.input_buffer.text.len() > 80)
                            {
                                this.script.buffer.set_text(this.lua_console.input_buffer.text.clone());
                            }
                            cx.notify();
                        },
                    )
                )
                .child(
                    console_toolbar_btn(
                        "console-clear-btn",
                        "Clear",
                        has_output,
                        text_muted, text_primary, panel_border,
                        cx,
                        |this, cx| {
                            this.lua_console.clear_output();
                            cx.notify();
                        },
                    )
                )
                .child(
                    console_toolbar_btn(
                        "console-maximize-btn",
                        if is_maximized { "Restore" } else { "Maximize" },
                        true,
                        text_muted, text_primary, panel_border,
                        cx,
                        move |this, cx| {
                            let window_h: f32 = this.window_size.height.into();
                            let eff_max = MAX_CONSOLE_HEIGHT.min(window_h * 0.6);
                            this.lua_console.toggle_maximize(eff_max);
                            cx.notify();
                        },
                    )
                )
                .child(
                    console_toolbar_btn(
                        "console-close-btn",
                        "\u{2715}",
                        true,
                        text_muted, text_primary, panel_border,
                        cx,
                        |this, cx| {
                            this.lua_console.hide();
                            cx.notify();
                        },
                    )
                )
        )
}

/// Render a single tab button in the tab bar
fn render_tab_button<F>(
    id: &'static str,
    label: &'static str,
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: F,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    div()
        .id(id)
        .px_3()
        .py(px(6.0))
        .text_size(px(14.0))
        .text_color(if is_active { text_primary } else { text_muted })
        .font_weight(if is_active { FontWeight::MEDIUM } else { FontWeight::NORMAL })
        .bg(if is_active { selection_bg.opacity(0.3) } else { gpui::transparent_black() })
        .border_b_2()
        .border_color(if is_active { text_primary } else { gpui::transparent_black() })
        .cursor_pointer()
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            on_click(this, cx);
        }))
        .child(label)
}

/// Toolbar button helper for console header
fn console_toolbar_btn<F>(
    id: &'static str,
    label: &'static str,
    enabled: bool,
    text_muted: Hsla,
    text_primary: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: F,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    let btn = div()
        .id(id)
        .px(px(6.0))
        .py(px(3.0))
        .rounded(px(3.0))
        .text_size(px(14.0));

    if enabled {
        btn
            .text_color(text_muted)
            .cursor_pointer()
            .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                on_click(this, cx);
            }))
            .child(label)
    } else {
        btn
            .text_color(text_muted.opacity(0.3))
            .child(label)
    }
}

/// Render a single output entry
fn render_output_entry(
    entry: &OutputEntry,
    tab: &str,
    index: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    error_color: Hsla,
) -> Stateful<Div> {
    let (prefix, color) = match entry.kind {
        OutputKind::Input => ("> ", text_muted),
        OutputKind::Result => ("", accent),
        OutputKind::Print => ("", text_primary),
        OutputKind::Error => ("", error_color),
        OutputKind::System | OutputKind::Stats => ("", text_muted),
    };

    div()
        .id(ElementId::Name(format!("lua-{}-output-{}", tab, index).into()))
        .text_xs()
        .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
        .text_color(color)
        .child(format!("{}{}", prefix, entry.text))
}

/// Render a single output entry with group-aware visual treatment (Run tab only).
fn render_output_entry_grouped(
    entry: &OutputEntry,
    tab: &str,
    index: usize,
    is_group_start: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    error_color: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let id = ElementId::Name(format!("lua-{}-output-{}", tab, index).into());

    match entry.kind {
        OutputKind::Input => {
            // Code card: left accent border, split multiline.
            // Separator border lives on the outer wrapper so border_color
            // doesn't conflict with the accent left border on the inner card.
            let mut card = div()
                .border_l_2()
                .border_color(accent.opacity(0.3))
                .pl(px(8.0))
                .py(px(2.0))
                .text_xs()
                .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                .text_color(text_muted);

            for line in entry.text.split('\n') {
                card = card.child(div().child(line.to_string()));
            }

            let mut wrapper = div().id(id);
            if is_group_start {
                wrapper = wrapper
                    .mt(px(6.0))
                    .border_t_1()
                    .border_color(panel_border.opacity(0.3))
                    .pt(px(4.0));
            }
            wrapper.child(card)
        }
        OutputKind::Stats => {
            // Dimmed metadata footer
            let mut wrapper = div()
                .id(id)
                .text_xs()
                .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                .text_color(text_muted.opacity(0.45))
                .child(entry.text.clone());

            if is_group_start {
                wrapper = wrapper
                    .mt(px(6.0))
                    .border_t_1()
                    .border_color(panel_border.opacity(0.3))
                    .pt(px(4.0));
            }
            wrapper
        }
        _ => {
            // Result/Print/Error/System — same logic as render_output_entry
            let (prefix, color) = match entry.kind {
                OutputKind::Result => ("", accent),
                OutputKind::Print => ("", text_primary),
                OutputKind::Error => ("", error_color),
                OutputKind::System => ("", text_muted),
                _ => ("", text_primary),
            };

            let mut wrapper = div()
                .id(id)
                .text_xs()
                .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                .text_color(color)
                .child(format!("{}{}", prefix, entry.text));

            if is_group_start {
                wrapper = wrapper
                    .mt(px(6.0))
                    .border_t_1()
                    .border_color(panel_border.opacity(0.3))
                    .pt(px(4.0));
            }
            wrapper
        }
    }
}

/// Render a scroll button (▲ or ▼)
fn scroll_button<F>(
    label: &'static str,
    enabled: bool,
    text_muted: Hsla,
    text_primary: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: F,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    let color = if enabled { text_muted } else { text_muted.opacity(0.3) };

    div()
        .id(ElementId::Name(format!("scroll-btn-{}", label).into()))
        .px_1()
        .text_xs()
        .text_color(color)
        .when(enabled, |div| {
            div
                .cursor_pointer()
                .hover(|s| s.text_color(text_primary))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    on_click(this, cx);
                }))
        })
        .child(label)
}

/// Render the input area with syntax highlighting, line numbers, and cursor
fn render_input_area(
    app: &Spreadsheet,
    _editor_bg: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    line_height: f32,
    max_lines: usize,
    _cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let console = &app.lua_console;
    let input = &console.input_buffer.text;
    let cursor = console.input_buffer.cursor;
    let scroll_offset = console.input_buffer.scroll_offset;

    // Color map for token types
    let color_keyword = app.token(TokenKey::FormulaFunction);  // violet
    let color_boolean = app.token(TokenKey::FormulaBoolean);   // cyan
    let color_string = app.token(TokenKey::FormulaString);     // green
    let color_number = app.token(TokenKey::FormulaNumber);     // amber
    let color_comment = text_muted;
    let color_operator = app.token(TokenKey::FormulaOperator);

    let token_color = |tt: LuaTokenType| -> Hsla {
        match tt {
            LuaTokenType::Keyword => color_keyword,
            LuaTokenType::Boolean => color_boolean,
            LuaTokenType::String => color_string,
            LuaTokenType::Number => color_number,
            LuaTokenType::Comment => color_comment,
            LuaTokenType::Operator | LuaTokenType::Punctuation => color_operator,
            LuaTokenType::Identifier => text_primary,
        }
    };

    // Get cached tokens and line offsets (primed by earlier tokens() call)
    let tokens = console.input_buffer.cached_tokens();
    let line_offsets = console.input_buffer.cached_line_offsets();

    let cursor_line = input[..cursor].matches('\n').count();
    let total_lines = line_offsets.len();
    let visible_start = scroll_offset.min(total_lines.saturating_sub(1));
    let visible_end = (visible_start + max_lines).min(total_lines);

    let gutter_width = 28.0_f32;

    let mut container = div()
        .id("lua-input")
        .flex_1()
        .overflow_hidden()
        .pt(px(7.0))
        .pb(px(7.0))
        .flex()
        .flex_col()
        .text_size(px(13.0))
        .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
        .max_h(px(max_lines as f32 * line_height));

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

        let line_el = super::code_render::render_code_line(
            line_text, line_start, line_end,
            tokens, cursor_in_line,
            line_num, is_cursor_line,
            gutter_width, line_height,
            text_primary, text_muted, accent,
            &token_color,
        );

        container = container.child(line_el);
    }

    container
}

/// Handle keyboard input in the console (called from main key handler)
pub fn handle_console_key_from_main(app: &mut Spreadsheet, event: &KeyDownEvent, window: &mut Window, cx: &mut Context<Spreadsheet>) {
    let key = &event.keystroke.key;
    let ctrl = event.keystroke.modifiers.control;
    let shift = event.keystroke.modifiers.shift;

    match key.as_str() {
        "enter" => {
            if ctrl {
                // Ctrl+Enter: always execute (power user shortcut)
                execute_console(app, cx);
            } else if shift {
                if app.lua_console.active_tab == ConsoleTab::Debug {
                    // Shift+Enter on Debug tab: start debug session
                    start_debug_session(app, cx);
                } else {
                    // Shift+Enter on Run tab: insert newline for multiline input
                    app.lua_console.insert("\n");
                }
            } else {
                // Enter: execute
                execute_console(app, cx);
            }
        }
        "escape" => {
            // Cancel / close and restore focus to main grid
            app.lua_console.hide();
            window.focus(&app.focus_handle, cx);
        }
        "backspace" => {
            app.lua_console.backspace();
        }
        "delete" => {
            app.lua_console.delete();
        }
        "left" => {
            app.lua_console.cursor_left();
        }
        "right" => {
            app.lua_console.cursor_right();
        }
        "home" => {
            if ctrl {
                // Ctrl+Home: scroll output to top + cursor to buffer start
                app.lua_console.scroll_to_start();
                app.lua_console.cursor_buffer_home();
            } else {
                app.lua_console.cursor_home();
            }
        }
        "end" => {
            if ctrl {
                // Ctrl+End: scroll output to bottom + cursor to buffer end
                app.lua_console.scroll_to_end();
                app.lua_console.cursor_buffer_end();
            } else {
                app.lua_console.cursor_end();
            }
        }
        "up" => {
            app.lua_console.history_prev();
        }
        "down" => {
            app.lua_console.history_next();
        }
        "pageup" => {
            // PageUp: scroll output up
            app.lua_console.scroll_page_up();
        }
        "pagedown" => {
            // PageDown: scroll output down
            app.lua_console.scroll_page_down();
        }
        "l" if ctrl => {
            // Ctrl+L: clear output
            app.lua_console.clear_output();
        }
        _ => {
            // Insert character
            if let Some(key_char) = &event.keystroke.key_char {
                if !ctrl && !event.keystroke.modifiers.alt {
                    app.lua_console.insert(key_char);
                }
            }
        }
    }

    cx.notify();
}

/// Execute the current input (public for action handlers).
///
/// Group lifetime is managed here: begin before the body, end after it returns.
/// The body can early-return on any path without leaking group state.
pub fn execute_console(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    let input = app.lua_console.consume_input();
    if input.trim().is_empty() {
        return;
    }

    app.lua_console.begin_group();
    app.lua_console.push_output(OutputEntry::input(&input));

    execute_console_body(app, input, cx);

    app.lua_console.end_group();
    cx.notify();
}

/// Inner body — all early returns are safe because the caller manages the group.
pub(crate) fn execute_console_body(app: &mut Spreadsheet, input: String, cx: &mut Context<Spreadsheet>) {
    use std::time::Instant;

    let trimmed = input.trim();

    if trimmed == "clear" {
        app.lua_console.clear_output();
        return;
    }

    if trimmed == "help" {
        app.lua_console.push_output(OutputEntry::system("Commands:"));
        app.lua_console.push_output(OutputEntry::system("  clear      - Clear output"));
        app.lua_console.push_output(OutputEntry::system("  examples   - List example scripts"));
        app.lua_console.push_output(OutputEntry::system("  example N  - Run example N (by number or name)"));
        app.lua_console.push_output(OutputEntry::system("  show N     - Show code for example N"));
        app.lua_console.push_output(OutputEntry::system("  help       - Show this help"));
        app.lua_console.push_output(OutputEntry::system(""));
        app.lua_console.push_output(OutputEntry::system("Keyboard:"));
        app.lua_console.push_output(OutputEntry::system("  Enter        - Execute"));
        app.lua_console.push_output(OutputEntry::system("  Shift+Enter  - Newline (multiline input)"));
        app.lua_console.push_output(OutputEntry::system("  Up/Down      - History navigation"));
        app.lua_console.push_output(OutputEntry::system("  PageUp/Down  - Scroll output"));
        app.lua_console.push_output(OutputEntry::system("  Ctrl+L       - Clear output"));
        app.lua_console.push_output(OutputEntry::system("  Escape       - Close console"));
        return;
    }

    if trimmed == "examples" {
        app.lua_console.push_output(OutputEntry::system("Example scripts (use 'example N' to load):"));
        for (i, example) in EXAMPLES.iter().enumerate() {
            app.lua_console.push_output(OutputEntry::system(
                format!("  {}. {} - {}", i + 1, example.name, example.description)
            ));
        }
        return;
    }

    // show N - preview example code without running
    if trimmed.starts_with("show ") {
        let arg = trimmed.strip_prefix("show ").unwrap().trim();

        let example = if let Ok(n) = arg.parse::<usize>() {
            get_example(n.saturating_sub(1))
        } else {
            find_example(arg)
        };

        if let Some(ex) = example {
            app.lua_console.push_output(OutputEntry::system(
                format!("-- {} --", ex.name)
            ));
            for line in ex.code.lines() {
                app.lua_console.push_output(OutputEntry::print(line.to_string()));
            }
            app.lua_console.push_output(OutputEntry::system(
                format!("Type 'example {}' to run", arg)
            ));
        } else {
            app.lua_console.push_output(OutputEntry::error(
                format!("Unknown example: '{}'. Type 'examples' to see list.", arg)
            ));
        }
        return;
    }

    // example N - run example directly
    let code_to_run = if trimmed.starts_with("example ") {
        let arg = trimmed.strip_prefix("example ").unwrap().trim();

        let example = if let Ok(n) = arg.parse::<usize>() {
            get_example(n.saturating_sub(1))
        } else {
            find_example(arg)
        };

        if let Some(ex) = example {
            app.lua_console.push_output(OutputEntry::system(
                format!("Running '{}'...", ex.name)
            ));
            Some(ex.code.to_string())
        } else {
            app.lua_console.push_output(OutputEntry::error(
                format!("Unknown example: '{}'. Type 'examples' to see list.", arg)
            ));
            return;
        }
    } else {
        None
    };

    // Use example code if running an example, otherwise use the user input
    let code = code_to_run.unwrap_or_else(|| input.clone());

    // Create snapshot of current sheet for Lua to read from
    let snapshot = SheetSnapshot::from_sheet(app.sheet(cx));
    let sheet_index = app.sheet_index(cx);

    // Compute selection bounds (normalize to start <= end)
    let (anchor_row, anchor_col) = app.view_state.selected;
    let (end_row, end_col) = app.view_state.selection_end.unwrap_or(app.view_state.selected);
    let selection = (
        anchor_row.min(end_row),
        anchor_col.min(end_col),
        anchor_row.max(end_row),
        anchor_col.max(end_col),
    );

    // Compute fingerprint before execution (for run records)
    let fingerprint_before = visigrid_io::native::compute_semantic_fingerprint(app.wb(cx));

    // Time the execution
    let start = Instant::now();

    // Evaluate with sheet access and selection
    let result = app.lua_runtime.eval_with_sheet_and_selection(&code, Box::new(snapshot), selection);

    let elapsed = start.elapsed();

    // Add output
    for line in &result.output {
        app.lua_console.push_output(OutputEntry::print(line.clone()));
    }

    if let Some(ref returned) = result.returned {
        app.lua_console.push_output(OutputEntry::result(returned.clone()));
    }

    if let Some(ref error) = result.error {
        app.lua_console.push_output(OutputEntry::error(error.clone()));
    }

    // Apply operations if any (single undo entry for all)
    if result.has_mutations() {
        let (changes, format_patches) = apply_lua_ops(app, sheet_index, &result.ops, cx);
        let cells_modified = changes.len() as i64;
        let has_values = !changes.is_empty();
        let has_formats = !format_patches.is_empty();

        // Build run record for ad-hoc console writes (provenance tracking)
        if has_values {
            build_console_run_record(
                app, cx, &code, &fingerprint_before, &changes,
                sheet_index, result.ops.len() as i64, elapsed, result.cells_read,
            );
        }

        if has_values && has_formats {
            use crate::history::{UndoAction, FormatActionKind};
            let group = UndoAction::Group {
                actions: vec![
                    UndoAction::Values { sheet_index, changes },
                    UndoAction::Format {
                        sheet_index,
                        patches: format_patches,
                        kind: FormatActionKind::CellStyle,
                        description: "Lua: set cell styles".into(),
                    },
                ],
                description: "Lua script".into(),
            };
            app.history.record_action_with_provenance(group, None);
            app.is_modified = true;
        } else if has_values {
            app.history.record_batch(sheet_index, changes);
            app.is_modified = true;
        } else if has_formats {
            app.history.record_format(
                sheet_index,
                format_patches,
                crate::history::FormatActionKind::CellStyle,
                "Lua: set cell styles".into(),
            );
            app.is_modified = true;
        }
    }

    // Show execution stats
    let stats = format!(
        "ops: {} | cells: {} | time: {:.1}ms",
        result.ops.len(),
        result.mutations,
        elapsed.as_secs_f64() * 1000.0
    );
    app.lua_console.push_output(OutputEntry::stats(stats));

    // Show cycle banner if Lua script introduced circular references
    app.maybe_show_cycle_banner(cx);
}

/// Build a run record for an ad-hoc console execution that produced mutations.
/// This provides provenance tracking even for one-liner console commands.
fn build_console_run_record(
    app: &mut Spreadsheet,
    cx: &mut gpui::Context<Spreadsheet>,
    code: &str,
    fingerprint_before: &str,
    changes: &[crate::history::CellChange],
    sheet_index: usize,
    ops_count: i64,
    elapsed: std::time::Duration,
    cells_read: usize,
) {
    use visigrid_io::scripting::{
        RunRecord, PatchLine, compute_script_hash, compute_diff_hash,
        build_diff_summary, compute_run_fingerprint, canonicalize_source,
    };

    let fingerprint_after = visigrid_io::native::compute_semantic_fingerprint(app.wb(cx));

    // Build PatchLines from CellChanges
    let mut patch_lines: Vec<PatchLine> = changes.iter().map(|c| {
        PatchLine {
            t: "cell".to_string(),
            sheet: sheet_index,
            r: c.row as u32,
            c: c.col as u32,
            k: if c.new_value.starts_with('=') { "formula".to_string() } else { "value".to_string() },
            old: if c.old_value.is_empty() { None } else { Some(c.old_value.clone()) },
            new: if c.new_value.is_empty() { None } else { Some(c.new_value.clone()) },
        }
    }).collect();

    // Sort for deterministic hashing
    patch_lines.sort_by(|a, b| {
        a.t.cmp(&b.t)
            .then(a.sheet.cmp(&b.sheet))
            .then(a.r.cmp(&b.r))
            .then(a.c.cmp(&b.c))
            .then(a.k.cmp(&b.k))
    });

    let diff_hash = if patch_lines.is_empty() { None } else { Some(compute_diff_hash(&patch_lines)) };

    let sheet_names: Vec<String> = app.wb(cx).sheet_names().iter().map(|s| s.to_string()).collect();
    let diff_summary = build_diff_summary(&patch_lines, &sheet_names);

    let canonical = canonicalize_source(code);
    let script_hash = compute_script_hash(code);

    let mut record = RunRecord {
        run_id: uuid::Uuid::new_v4().to_string(),
        run_fingerprint: String::new(), // computed below
        script_name: "(console)".to_string(),
        script_hash,
        script_source: canonical,
        script_origin: r#"{"kind":"Console"}"#.to_string(),
        capabilities_used: "SheetRead,SheetWriteValues,SheetWriteFormulas".to_string(),
        params: None,
        fingerprint_before: fingerprint_before.to_string(),
        fingerprint_after,
        diff_hash,
        diff_summary,
        cells_read: cells_read as i64,
        cells_modified: changes.len() as i64,
        ops_count,
        duration_ms: elapsed.as_millis() as i64,
        ran_at: chrono::Utc::now().to_rfc3339(),
        ran_by: None,
        status: "ok".to_string(),
        error: None,
    };

    record.run_fingerprint = compute_run_fingerprint(&record);
    app.pending_run_records.push(record);
}

/// Apply Lua operations to the sheet and return undo changes.
/// Uses batched, tracked mutations so dependents recalculate once at the end.
/// Returns (value_changes, format_patches) for separate undo tracking.
fn apply_lua_ops(
    app: &mut Spreadsheet,
    sheet_index: usize,
    ops: &[LuaOp],
    cx: &mut gpui::Context<Spreadsheet>,
) -> (Vec<crate::history::CellChange>, Vec<crate::history::CellFormatPatch>) {
    use crate::history::CellChange;
    use crate::history::CellFormatPatch;
    use visigrid_engine::cell::CellStyle;

    if ops.is_empty() {
        return (Vec::new(), Vec::new());
    }

    app.workbook.update(cx, |wb, _| {
        let mut guard = wb.batch_guard();
        let mut changes = Vec::new();
        let mut format_patches = Vec::new();

        for op in ops {
            match op {
                LuaOp::SetValue { row, col, value } => {
                    let row = *row as usize;
                    let col = *col as usize;

                    let old_value = guard.sheet(sheet_index)
                        .map(|s| s.get_raw(row, col))
                        .unwrap_or_default();

                    let new_value = lua_cell_value_to_string(value);
                    guard.set_cell_value_tracked(sheet_index, row, col, &new_value);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value,
                    });
                }
                LuaOp::SetFormula { row, col, formula } => {
                    let row = *row as usize;
                    let col = *col as usize;

                    let old_value = guard.sheet(sheet_index)
                        .map(|s| s.get_raw(row, col))
                        .unwrap_or_default();

                    guard.set_cell_value_tracked(sheet_index, row, col, formula);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: formula.clone(),
                    });
                }
                LuaOp::SetCellStyle { r1, c1, r2, c2, style } => {
                    let cell_style = CellStyle::from_int(*style as i32);
                    for row in (*r1 as usize)..=(*r2 as usize) {
                        for col in (*c1 as usize)..=(*c2 as usize) {
                            let before = guard.sheet(sheet_index)
                                .map(|s| s.get_format(row, col))
                                .unwrap_or_default();
                            if let Some(s) = guard.sheet_mut(sheet_index) {
                                s.set_cell_style(row, col, cell_style);
                            }
                            let after = guard.sheet(sheet_index)
                                .map(|s| s.get_format(row, col))
                                .unwrap_or_default();
                            if before != after {
                                format_patches.push(CellFormatPatch { row, col, before, after });
                            }
                        }
                    }
                }
            }
        }

        (changes, format_patches)
    })
}

/// Convert LuaCellValue to a string suitable for sheet.set_value()
fn lua_cell_value_to_string(value: &LuaCellValue) -> String {
    match value {
        LuaCellValue::Nil => String::new(),  // Empty clears the cell
        LuaCellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            }
        }
        LuaCellValue::String(s) => s.clone(),
        LuaCellValue::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        LuaCellValue::Error(e) => format!("#ERROR: {}", e),
    }
}

// =============================================================================
// Debug session integration (Phase 4)
// =============================================================================

/// Drain debug events from the background thread and apply them to state.
///
/// Called every render frame from `views/mod.rs`, before `render_lua_console`.
/// Two-pass design avoids split borrows: pass 1 drains `event_rx` (needs `&mut session`),
/// pass 2 processes payloads (needs `&mut app`).
/// Refresh the token cache if the input has changed. Called from the render
/// site (views/mod.rs) immediately before `render_lua_console` — so the work
/// is driven by rendering, not by debug event pumping.
pub fn refresh_input_tokens(app: &mut Spreadsheet) {
    if app.lua_console.visible && app.lua_console.active_tab == ConsoleTab::Run {
        app.lua_console.input_buffer.tokens();
    }
}

pub fn pump_debug_events(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    const MAX_EVENTS_PER_TICK: usize = 200;

    // Keep viewport_lines fresh for on_debug_paused scroll calculations.
    // Runs every frame, so resize/maximize/window-resize are picked up immediately.
    if app.lua_console.debug_session.is_some() {
        let console_height = if app.lua_console.is_maximized {
            let window_h: f32 = app.window_size.height.into();
            MAX_CONSOLE_HEIGHT.min(window_h * 0.6)
        } else {
            app.lua_console.height
        };
        // Subtract tab bar (~28), resize handle (4), controls bar (24),
        // debug output strip (~40), borders/padding (~4) ≈ 100px overhead
        let available = (console_height - 100.0).max(28.0);
        app.lua_console.debug_source_viewport_lines = (available / 14.0).floor().max(1.0) as usize;
    }

    // Pass 1: drain events into a local vec
    let mut events = Vec::new();
    if let Some(ref mut session) = app.lua_console.debug_session {
        let session_id = session.id;
        let cancelled = session.cancel.load(Ordering::Relaxed);
        for _ in 0..MAX_EVENTS_PER_TICK {
            match session.event_rx.try_recv() {
                Ok(event) if event.session_id == session_id => {
                    events.push(event.payload);
                }
                Ok(_) => continue, // stale session_id — discard
                Err(TryRecvError::Empty) => break,
                Err(TryRecvError::Disconnected) => {
                    if !cancelled {
                        events.push(DebugEventPayload::Error(
                            "debugger crashed unexpectedly".to_string(),
                        ));
                    }
                    break;
                }
            }
        }
    }

    if events.is_empty() {
        return;
    }

    // Pass 2: process events (full &mut app available)
    for payload in events {
        match payload {
            DebugEventPayload::OutputChunk(chunk) => {
                for line in chunk.lines() {
                    app.lua_console
                        .debug_output
                        .push(OutputEntry::print(line.to_string()));
                }
                // Ring buffer cap
                if app.lua_console.debug_output.len() > DEBUG_OUTPUT_CAP {
                    let drain = app.lua_console.debug_output.len() - DEBUG_OUTPUT_CAP;
                    app.lua_console.debug_output.drain(..drain);
                }
            }
            DebugEventPayload::Paused(snapshot) => {
                // Normalize 1-indexed Lua line to 0-indexed at the boundary
                let paused_line_0 = snapshot.line.saturating_sub(1);
                let total_lines = app.lua_console.input_buffer.text.lines().count();
                app.lua_console.on_debug_paused(paused_line_0, total_lines);

                if let Some(ref mut s) = app.lua_console.debug_session {
                    s.state = DebugSessionState::Paused;
                }
                app.lua_console.debug_snapshot = Some(snapshot);
            }
            DebugEventPayload::Completed(result) => {
                handle_debug_completed(app, result, cx);
            }
            DebugEventPayload::Error(msg) => {
                app.lua_console
                    .debug_output
                    .push(OutputEntry::error(msg.clone()));
                app.lua_console
                    .push_output_ungrouped(OutputEntry::error(format!("[debug] {}", msg)));
                app.lua_console.debug_session = None;
                app.lua_console.debug_snapshot = None;
            }
            DebugEventPayload::FrameVars { frame_index, locals, upvalues } => {
                app.lua_console.frame_vars_cache.insert(frame_index, (locals, upvalues));
            }
            DebugEventPayload::VariableExpanded { frame_index, path, children } => {
                let key = ConsoleState::var_expansion_key(frame_index, &path);
                app.lua_console.expanded_vars.insert(key, children);
            }
        }
    }
    cx.notify();
}

/// Handle a Completed event from the debug thread.
fn handle_debug_completed(
    app: &mut Spreadsheet,
    result: LuaEvalResult,
    cx: &mut Context<Spreadsheet>,
) {
    if result.cancelled {
        app.lua_console
            .push_output_ungrouped(OutputEntry::system("[debug] session stopped"));
    } else if let Some(ref error) = result.error {
        app.lua_console
            .push_output_ungrouped(OutputEntry::error(format!("[debug] error: {}", error)));
        app.lua_console
            .debug_output
            .push(OutputEntry::error(error.clone()));
    } else {
        // Apply ops
        let start_sheet_index = app
            .lua_console
            .debug_session
            .as_ref()
            .map(|s| s.start_sheet_index)
            .unwrap_or(0);

        let sheet_count = app.wb(cx).sheet_count();
        let target_index = if start_sheet_index < sheet_count {
            start_sheet_index
        } else {
            let current = app.sheet_index(cx);
            app.lua_console.debug_output.push(OutputEntry::system(
                "[debug] target sheet no longer exists; applied to current sheet",
            ));
            app.lua_console.push_output_ungrouped(OutputEntry::system(
                "[debug] target sheet no longer exists; applied to current sheet",
            ));
            current
        };

        if result.has_mutations() {
            let (changes, format_patches) = apply_lua_ops(app, target_index, &result.ops, cx);
            let has_values = !changes.is_empty();
            let has_formats = !format_patches.is_empty();

            if has_values && has_formats {
                use crate::history::{FormatActionKind, UndoAction};
                let group = UndoAction::Group {
                    actions: vec![
                        UndoAction::Values {
                            sheet_index: target_index,
                            changes,
                        },
                        UndoAction::Format {
                            sheet_index: target_index,
                            patches: format_patches,
                            kind: FormatActionKind::CellStyle,
                            description: "Lua debug: set cell styles".into(),
                        },
                    ],
                    description: "Lua debug".into(),
                };
                app.history.record_action_with_provenance(group, None);
                app.is_modified = true;
            } else if has_values {
                app.history.record_batch(target_index, changes);
                app.is_modified = true;
            } else if has_formats {
                app.history.record_format(
                    target_index,
                    format_patches,
                    crate::history::FormatActionKind::CellStyle,
                    "Lua debug: set cell styles".into(),
                );
                app.is_modified = true;
            }
        }

        let stats = format!(
            "[debug] session completed (ops: {} | cells: {})",
            result.ops.len(),
            result.mutations
        );
        app.lua_console.push_output_ungrouped(OutputEntry::stats(stats));
    }

    app.lua_console.debug_session = None;
    app.lua_console.debug_snapshot = None;
}

/// Start a debug session from the current console input.
pub fn start_debug_session(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    let input = app.lua_console.input_buffer.text.clone(); // Don't consume — keep for re-debug
    if input.trim().is_empty() {
        return;
    }

    let snapshot = SheetSnapshot::from_sheet(app.sheet(cx));
    let sheet_index = app.sheet_index(cx);
    let (anchor_row, anchor_col) = app.view_state.selected;
    let (end_row, end_col) = app
        .view_state
        .selection_end
        .unwrap_or(app.view_state.selected);
    let selection = (
        anchor_row.min(end_row),
        anchor_col.min(end_col),
        anchor_row.max(end_row),
        anchor_col.max(end_col),
    );

    let config = DebugConfig {
        code: input,
        snapshot,
        selection,
        breakpoints: app.lua_console.breakpoints.clone(),
    };

    let session = spawn_debug_session(config);
    app.lua_console.start_debug_session(session, sheet_index);
    app.lua_console
        .push_output_ungrouped(OutputEntry::system("[debug] session started"));
    cx.notify();
}

/// Check whether the debug tab is active, visible, and focused.
fn is_debug_active(app: &Spreadsheet, window: &Window) -> bool {
    app.lua_console.visible
        && app.lua_console.active_tab == ConsoleTab::Debug
        && app.console_focus_handle.is_focused(window)
}

/// Render the Debug tab content with action handlers.
fn render_debug_tab_content(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    error_color: Hsla,
    panel_border: Hsla,
    _editor_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let has_session = app.lua_console.debug_session.is_some();
    let session_state = app
        .lua_console
        .debug_session
        .as_ref()
        .map(|s| s.state);
    let is_paused = session_state == Some(DebugSessionState::Paused);
    let text_inverse = app.token(TokenKey::TextInverse);

    // Build inner content first (needs cx for click handlers inside debug_ui)
    let content: AnyElement = if has_session {
        div()
            .flex_1()
            .flex()
            .flex_col()
            .overflow_hidden()
            // Controls bar (fixed 24px)
            .child(debug_ui::render_debug_controls(
                app, session_state, is_paused,
                text_primary, text_muted, accent, error_color, panel_border, cx,
            ))
            // Main content: source pane + side pane
            .child(
                div()
                    .flex_1()
                    .flex()
                    .flex_row()
                    .overflow_hidden()
                    .child(debug_ui::render_source_pane(
                        app, is_paused,
                        text_primary, text_muted, accent, panel_border, cx,
                    ))
                    .child(
                        div()
                            .w(px(200.0))
                            .flex_shrink_0()
                            .flex()
                            .flex_col()
                            .border_l_1()
                            .border_color(panel_border)
                            .overflow_hidden()
                            .child(debug_ui::render_call_stack_pane(
                                app,
                                text_primary, text_muted, accent, panel_border, cx,
                            ))
                            .child(debug_ui::render_variables_pane(
                                app,
                                text_primary, text_muted, accent, panel_border, cx,
                            ))
                    )
            )
            // Debug output (bottom strip)
            .child(debug_ui::render_debug_output(
                &app.lua_console.debug_output,
                text_primary, text_muted, accent, error_color, panel_border,
            ))
            .into_any_element()
    } else if !visigrid_license::is_feature_enabled("lua_tooling") {
        let preview = debug_ui::render_locked_preview(text_muted, accent, panel_border);
        match render_locked_feature_panel(
            "Lua Debugger",
            "Set breakpoints, step through code, inspect variables, and trace execution in your Lua scripts.",
            preview,
            app.locked_panels_dismissed,
            panel_border, text_primary, text_muted, accent, text_inverse,
            cx,
        ) {
            Some(panel) => panel,
            None => debug_ui::render_idle_help(text_muted).into_any_element(),
        }
    } else {
        debug_ui::render_idle_help(text_muted).into_any_element()
    };

    // Outer div with action handlers (cx borrows released from content building above)
    div()
        .id("debug-tab-content")
        .key_context("LuaDebug")
        .flex_1()
        .flex()
        .flex_col()
        .overflow_hidden()
        // Action handlers
        .on_action(cx.listener(|this, _: &DebugStartOrContinue, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            if let Some(ref s) = this.lua_console.debug_session {
                if s.state == DebugSessionState::Paused {
                    this.lua_console.send_debug_action(DebugAction::Continue);
                    this.lua_console.set_debug_running();
                }
            } else {
                start_debug_session(this, cx);
            }
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &DebugStepOver, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            if let Some(ref s) = this.lua_console.debug_session {
                if s.state == DebugSessionState::Paused {
                    this.lua_console.send_debug_action(DebugAction::StepOver);
                    this.lua_console.set_debug_running();
                    cx.notify();
                }
            }
        }))
        .on_action(cx.listener(|this, _: &DebugStepIn, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            if let Some(ref s) = this.lua_console.debug_session {
                if s.state == DebugSessionState::Paused {
                    this.lua_console.send_debug_action(DebugAction::StepIn);
                    this.lua_console.set_debug_running();
                    cx.notify();
                }
            }
        }))
        .on_action(cx.listener(|this, _: &DebugStepOut, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            if let Some(ref s) = this.lua_console.debug_session {
                if s.state == DebugSessionState::Paused {
                    this.lua_console.send_debug_action(DebugAction::StepOut);
                    this.lua_console.set_debug_running();
                    cx.notify();
                }
            }
        }))
        .on_action(cx.listener(|this, _: &DebugStop, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            this.lua_console.stop_debug_session();
            this.lua_console.push_output_ungrouped(OutputEntry::system("[debug] session stopped"));
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &DebugToggleBreakpoint, window, cx| {
            if !is_debug_active(this, window) {
                return;
            }
            // Toggle breakpoint at current paused line
            if let Some(ref snap) = this.lua_console.debug_snapshot {
                let line = snap.line;
                let bp = (CONSOLE_SOURCE.to_string(), line);
                if this.lua_console.breakpoints.contains(&bp) {
                    this.lua_console.breakpoints.remove(&bp);
                    this.lua_console.send_debug_action(DebugAction::RemoveBreakpoint {
                        source: CONSOLE_SOURCE.to_string(), line,
                    });
                } else {
                    this.lua_console.breakpoints.insert(bp);
                    this.lua_console.send_debug_action(DebugAction::AddBreakpoint {
                        source: CONSOLE_SOURCE.to_string(), line,
                    });
                }
                cx.notify();
            }
        }))
        .child(content)
}

/// A small button for debug actions (Continue, Step, Stop, etc.)
fn debug_action_btn<F>(
    id: &'static str,
    label: &'static str,
    enabled: bool,
    text_muted: Hsla,
    text_primary: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: F,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    let btn = div()
        .id(id)
        .px(px(6.0))
        .py(px(2.0))
        .rounded(px(3.0))
        .text_size(px(14.0))
        .border_1()
        .border_color(panel_border);

    if enabled {
        btn
            .text_color(accent)
            .cursor_pointer()
            .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                on_click(this, cx);
            }))
            .child(label)
    } else {
        btn
            .text_color(text_muted.opacity(0.3))
            .child(label)
    }
}

// =============================================================================
// Debug panel rendering helpers (Phase 5)
// =============================================================================

mod debug_ui {
    use super::*;
    use crate::scripting::debugger::Variable;

    const MAX_UI_DEPTH: usize = 4;
    const MAX_RENDER_CHILDREN: usize = 30;
    const MAX_DEBUG_OUTPUT_LINES: usize = 50;

    /// Idle help text shown when no debug session is active.
    pub(super) fn render_idle_help(text_muted: Hsla) -> Div {
        let kbd_bg = text_muted.opacity(0.1);
        let kbd_border = text_muted.opacity(0.2);

        let kbd = |key: &'static str, label: &'static str| {
            div()
                .flex()
                .items_center()
                .gap(px(4.0))
                .child(
                    div()
                        .px(px(5.0))
                        .py(px(1.0))
                        .rounded(px(3.0))
                        .bg(kbd_bg)
                        .border_1()
                        .border_color(kbd_border)
                        .text_size(px(14.0))
                        .font_weight(FontWeight::MEDIUM)
                        .child(key)
                )
                .child(
                    div()
                        .text_size(px(14.0))
                        .text_color(text_muted.opacity(0.6))
                        .child(label)
                )
        };

        div()
            .flex()
            .flex_col()
            .gap(px(8.0))
            .py_3()
            .px_3()
            .text_color(text_muted.opacity(0.7))
            .child(
                div()
                    .text_size(px(13.0))
                    .child("Enter a script in the Run tab, then press F5 or Shift+Enter to debug.")
            )
            .child(
                div()
                    .flex()
                    .flex_wrap()
                    .gap(px(8.0))
                    .child(kbd("F5", "Continue"))
                    .child(kbd("F10", "Step Over"))
                    .child(kbd("F11", "Step In"))
                    .child(kbd("Shift+F11", "Step Out"))
                    .child(kbd("Shift+F5", "Stop"))
                    .child(kbd("F9", "Breakpoint"))
            )
    }

    /// Preview skeleton for the locked feature panel.
    pub(super) fn render_locked_preview(
        text_muted: Hsla,
        accent: Hsla,
        panel_border: Hsla,
    ) -> AnyElement {
        div()
            .flex()
            .flex_col()
            .gap_1()
            .p_2()
            // Fake controls bar
            .child(
                div()
                    .flex()
                    .items_center()
                    .gap_1()
                    .child(div().h(px(6.0)).w(px(60.0)).rounded_sm().bg(accent.opacity(0.15)))
                    .child(div().h(px(6.0)).w(px(30.0)).rounded_sm().bg(panel_border.opacity(0.3)))
                    .child(div().h(px(6.0)).w(px(30.0)).rounded_sm().bg(panel_border.opacity(0.3)))
                    .child(div().h(px(6.0)).w(px(30.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            )
            // Fake source lines
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(1.0))
                    .child(div().h(px(6.0)).w(px(140.0)).rounded_sm().bg(text_muted.opacity(0.08)))
                    .child(div().h(px(6.0)).w(px(100.0)).rounded_sm().bg(accent.opacity(0.1)))
                    .child(div().h(px(6.0)).w(px(120.0)).rounded_sm().bg(text_muted.opacity(0.08)))
            )
            .into_any_element()
    }

    /// Controls bar: status label + step buttons (fixed 24px height).
    pub(super) fn render_debug_controls(
        app: &Spreadsheet,
        _session_state: Option<DebugSessionState>,
        is_paused: bool,
        text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        error_color: Hsla,
        panel_border: Hsla,
        cx: &mut Context<Spreadsheet>,
    ) -> impl IntoElement {
        let status_text: String;
        let status_color: Hsla;

        if is_paused {
            status_color = accent;
            if let Some(ref snap) = app.lua_console.debug_snapshot {
                let reason = match snap.reason {
                    crate::scripting::PauseReason::Breakpoint => "breakpoint",
                    crate::scripting::PauseReason::StepIn => "step in",
                    crate::scripting::PauseReason::StepOver => "step over",
                    crate::scripting::PauseReason::StepOut => "step out",
                    crate::scripting::PauseReason::Entry => "entry",
                };
                // Include source name for diagnostic visibility
                let source = snap.call_stack.first()
                    .and_then(|f| f.source.as_deref())
                    .unwrap_or("?");
                status_text = format!("Paused at {}:{} ({})", source, snap.line, reason);
            } else {
                status_text = "Paused".to_string();
            }
        } else {
            status_text = "Running...".to_string();
            status_color = accent;
        }

        div()
            .h(px(24.0))
            .flex_shrink_0()
            .flex()
            .items_center()
            .justify_between()
            .px_2()
            .border_b_1()
            .border_color(panel_border)
            // Left: status
            .child(
                div()
                    .text_size(px(14.0))
                    .text_color(status_color)
                    .child(status_text)
            )
            // Right: buttons
            .child(
                div()
                    .flex()
                    .items_center()
                    .gap_1()
                    .child(debug_action_btn(
                        "debug-continue", "\u{25B6} F5", is_paused,
                        text_muted, text_primary, accent, panel_border, cx,
                        |this, cx| {
                            this.lua_console.send_debug_action(DebugAction::Continue);
                            this.lua_console.set_debug_running();
                            cx.notify();
                        },
                    ))
                    .child(debug_action_btn(
                        "debug-step-over", "\u{2192} F10", is_paused,
                        text_muted, text_primary, accent, panel_border, cx,
                        |this, cx| {
                            this.lua_console.send_debug_action(DebugAction::StepOver);
                            this.lua_console.set_debug_running();
                            cx.notify();
                        },
                    ))
                    .child(debug_action_btn(
                        "debug-step-in", "\u{2193} F11", is_paused,
                        text_muted, text_primary, accent, panel_border, cx,
                        |this, cx| {
                            this.lua_console.send_debug_action(DebugAction::StepIn);
                            this.lua_console.set_debug_running();
                            cx.notify();
                        },
                    ))
                    .child(debug_action_btn(
                        "debug-step-out", "\u{2191} S+F11", is_paused,
                        text_muted, text_primary, accent, panel_border, cx,
                        |this, cx| {
                            this.lua_console.send_debug_action(DebugAction::StepOut);
                            this.lua_console.set_debug_running();
                            cx.notify();
                        },
                    ))
                    .child(debug_action_btn(
                        "debug-stop", "\u{25A0} S+F5", true,
                        text_muted, text_primary, error_color, panel_border, cx,
                        |this, cx| {
                            this.lua_console.stop_debug_session();
                            this.lua_console.push_output_ungrouped(OutputEntry::system("[debug] session stopped"));
                            cx.notify();
                        },
                    ))
            )
    }

    /// Source pane: gutter with breakpoints + code lines.
    pub(super) fn render_source_pane(
        app: &Spreadsheet,
        _is_paused: bool,
        text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        panel_border: Hsla,
        cx: &mut Context<Spreadsheet>,
    ) -> impl IntoElement {
        let source = &app.lua_console.input_buffer.text;
        let lines: Vec<&str> = source.lines().collect();
        let total_lines = lines.len().max(1);
        let scroll = app.lua_console.debug_source_scroll;
        let viewport_lines = app.lua_console.debug_source_viewport_lines;

        let snapshot = app.lua_console.debug_snapshot.as_ref();
        let paused_line_1 = snapshot.map(|s| s.line).unwrap_or(0); // 1-indexed

        let breakpoints = &app.lua_console.breakpoints;

        // Compute visible range
        let start = scroll.min(total_lines.saturating_sub(1));
        let end = (start + viewport_lines + 2).min(total_lines); // +2 for partial lines

        div()
            .id("debug-source-pane")
            .flex_1()
            .flex()
            .flex_col()
            .overflow_hidden()
            .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
            .text_size(px(13.0))
            .children((start..end).map(|line_idx| {
                let line_num = line_idx + 1; // 1-indexed for display + breakpoint matching
                let line_text = lines.get(line_idx).copied().unwrap_or("");
                let is_current = line_num == paused_line_1;
                let has_bp = breakpoints.contains(&(CONSOLE_SOURCE.to_string(), line_num));

                div()
                    .id(ElementId::Name(format!("src-line-{}", line_num).into()))
                    .h(px(14.0))
                    .flex()
                    .items_center()
                    .when(is_current, |d| d.bg(accent.opacity(0.15)))
                    // Gutter (fixed 36px)
                    .child(
                        div()
                            .w(px(36.0))
                            .flex_shrink_0()
                            .flex()
                            .items_center()
                            .h_full()
                            .cursor_pointer()
                            .on_mouse_down(MouseButton::Left, {
                                let source_name = CONSOLE_SOURCE.to_string();
                                cx.listener(move |this, _, _, cx| {
                                    let bp = (source_name.clone(), line_num);
                                    if this.lua_console.breakpoints.contains(&bp) {
                                        this.lua_console.breakpoints.remove(&bp);
                                        this.lua_console.send_debug_action(DebugAction::RemoveBreakpoint {
                                            source: CONSOLE_SOURCE.to_string(), line: line_num,
                                        });
                                    } else {
                                        this.lua_console.breakpoints.insert(bp);
                                        this.lua_console.send_debug_action(DebugAction::AddBreakpoint {
                                            source: CONSOLE_SOURCE.to_string(), line: line_num,
                                        });
                                    }
                                    cx.notify();
                                })
                            })
                            // Breakpoint dot area (12px)
                            .child(
                                div()
                                    .w(px(12.0))
                                    .flex()
                                    .items_center()
                                    .justify_center()
                                    .text_size(px(8.0))
                                    .when(has_bp, |d| {
                                        d.text_color(rgb(0xE51400)) // Red dot
                                            .child("\u{25CF}") // ●
                                    })
                                    .when(!has_bp, |d| {
                                        d.text_color(gpui::transparent_black())
                                            .hover(|s| s.text_color(text_muted.opacity(0.4)))
                                            .child("\u{25CB}") // ○
                                    })
                            )
                            // Line number (right-aligned, 20px)
                            .child(
                                div()
                                    .w(px(20.0))
                                    .text_size(px(9.0))
                                    .text_color(text_muted.opacity(0.5))
                                    .flex()
                                    .justify_end()
                                    .pr(px(4.0))
                                    .child(format!("{}", line_num))
                            )
                    )
                    // Separator
                    .child(
                        div()
                            .w(px(1.0))
                            .h_full()
                            .bg(panel_border.opacity(0.3))
                    )
                    // Code text
                    .child(
                        div()
                            .flex_1()
                            .pl(px(4.0))
                            .overflow_hidden()
                            .text_color(text_primary)
                            .when(is_current, |d| d.font_weight(FontWeight::MEDIUM))
                            .child(line_text.to_string())
                    )
            }))
    }

    /// Call stack pane: frame list with selection.
    pub(super) fn render_call_stack_pane(
        app: &Spreadsheet,
        text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        panel_border: Hsla,
        cx: &mut Context<Spreadsheet>,
    ) -> impl IntoElement {
        let snapshot = app.lua_console.debug_snapshot.as_ref();
        let selected_frame = app.lua_console.selected_frame;

        div()
            .flex_shrink_0()
            .flex()
            .flex_col()
            .border_b_1()
            .border_color(panel_border)
            .max_h(px(100.0))
            .overflow_hidden()
            // Section header
            .child(
                div()
                    .px_1()
                    .py(px(2.0))
                    .text_size(px(9.0))
                    .text_color(text_muted.opacity(0.6))
                    .font_weight(FontWeight::SEMIBOLD)
                    .child("CALL STACK")
            )
            .children(
                snapshot
                    .map(|snap| &snap.call_stack[..])
                    .unwrap_or(&[])
                    .iter()
                    .enumerate()
                    .map(|(i, frame)| {
                        let is_selected = i == selected_frame;
                        let source = frame.source.as_deref().unwrap_or("?");
                        let label = format!("{}:{}", source, frame.line);

                        div()
                            .id(ElementId::Name(format!("call-frame-{}", i).into()))
                            .px_1()
                            .py(px(1.0))
                            .flex()
                            .items_center()
                            .gap(px(4.0))
                            .text_size(px(14.0))
                            .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                            .cursor_pointer()
                            .when(is_selected, |d| d.bg(accent.opacity(0.15)))
                            .hover(|s| s.bg(accent.opacity(0.08)))
                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                this.lua_console.select_frame(i);
                                cx.notify();
                            }))
                            .child(
                                div().text_color(text_primary).child(label)
                            )
                            .when(frame.function_name.is_some(), |d| {
                                let name = frame.function_name.as_deref().unwrap_or("");
                                d.child(
                                    div()
                                        .text_color(text_muted)
                                        .italic()
                                        .child(format!("in {}", name))
                                )
                            })
                            .when(frame.is_tail_call, |d| {
                                d.child(
                                    div()
                                        .text_size(px(8.0))
                                        .text_color(text_muted.opacity(0.5))
                                        .child("[tail]")
                                )
                            })
                    })
            )
    }

    /// Variables pane: locals + upvalues with lazy expansion.
    pub(super) fn render_variables_pane(
        app: &Spreadsheet,
        text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        _panel_border: Hsla,
        cx: &mut Context<Spreadsheet>,
    ) -> impl IntoElement {
        let snapshot = app.lua_console.debug_snapshot.as_ref();
        let selected_frame = app.lua_console.selected_frame;
        let expanded_vars = &app.lua_console.expanded_vars;

        // Get variables for selected frame
        let (locals, upvalues): (&[Variable], &[Variable]) = if selected_frame == 0 {
            match snapshot {
                Some(snap) => (&snap.locals, &snap.upvalues),
                None => (&[], &[]),
            }
        } else {
            match app.lua_console.frame_vars_cache.get(&selected_frame) {
                Some((l, u)) => (l.as_slice(), u.as_slice()),
                None => (&[], &[]),
            }
        };

        let frame_loading = selected_frame != 0
            && !app.lua_console.frame_vars_cache.contains_key(&selected_frame)
            && snapshot.is_some();

        let mut rows: Vec<AnyElement> = Vec::new();

        if frame_loading {
            rows.push(
                div()
                    .px_1()
                    .text_size(px(14.0))
                    .text_color(text_muted)
                    .italic()
                    .child(format!("Loading frame {} vars\u{2026}", selected_frame))
                    .into_any_element(),
            );
        } else {
            // Locals section
            if !locals.is_empty() {
                rows.push(
                    div()
                        .px_1()
                        .py(px(2.0))
                        .text_size(px(9.0))
                        .text_color(text_muted.opacity(0.6))
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("LOCALS")
                        .into_any_element(),
                );
                for (i, var) in locals.iter().enumerate() {
                    let path = vec![VarPathSegment::Local(i)];
                    render_variable_rows(
                        var, 0, &path, selected_frame, expanded_vars,
                        text_primary, text_muted, accent, cx, &mut rows,
                    );
                }
            }

            // Upvalues section
            if !upvalues.is_empty() {
                rows.push(
                    div()
                        .px_1()
                        .py(px(2.0))
                        .text_size(px(9.0))
                        .text_color(text_muted.opacity(0.6))
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("UPVALUES")
                        .into_any_element(),
                );
                for (i, var) in upvalues.iter().enumerate() {
                    let path = vec![VarPathSegment::Upvalue(i)];
                    render_variable_rows(
                        var, 0, &path, selected_frame, expanded_vars,
                        text_primary, text_muted, accent, cx, &mut rows,
                    );
                }
            }
        }

        div()
            .flex_1()
            .flex()
            .flex_col()
            .overflow_hidden()
            // Section header
            .child(
                div()
                    .px_1()
                    .py(px(2.0))
                    .text_size(px(9.0))
                    .text_color(text_muted.opacity(0.6))
                    .font_weight(FontWeight::SEMIBOLD)
                    .child(format!("VARIABLES (frame {})", selected_frame))
            )
            .children(rows)
    }

    /// Recursively render a variable and its expanded children.
    fn render_variable_rows(
        var: &Variable,
        depth: usize,
        path: &[VarPathSegment],
        frame_index: usize,
        expanded_vars: &std::collections::HashMap<String, Vec<Variable>>,
        text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        cx: &mut Context<Spreadsheet>,
        rows: &mut Vec<AnyElement>,
    ) {
        let key = ConsoleState::var_expansion_key(frame_index, path);
        let is_expanded = expanded_vars.contains_key(&key);
        let indent = depth as f32 * 12.0;

        // Disclosure triangle
        let disclosure = if var.expandable {
            if is_expanded { "\u{25BC} " } else { "\u{25B6} " }
        } else {
            "  "
        };

        let row_id = format!("var-{}", key);
        let path_owned = path.to_vec();
        let expandable = var.expandable;

        rows.push(
            div()
                .id(ElementId::Name(row_id.into()))
                .pl(px(indent + 4.0))
                .pr(px(2.0))
                .py(px(1.0))
                .flex()
                .items_center()
                .text_size(px(14.0))
                .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                .when(expandable, |d| {
                    d.cursor_pointer()
                        .hover(|s| s.bg(accent.opacity(0.06)))
                })
                .when(expandable, |d| {
                    let key_clone = key.clone();
                    let path_clone = path_owned.clone();
                    d.on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        if this.lua_console.expanded_vars.contains_key(&key_clone) {
                            // Collapse
                            this.lua_console.expanded_vars.remove(&key_clone);
                        } else {
                            // Expand — request children from backend
                            this.lua_console.send_debug_action(DebugAction::ExpandVariable {
                                frame_index,
                                path: path_clone.clone(),
                            });
                        }
                        cx.notify();
                    }))
                })
                .child(
                    div().text_color(text_muted).child(disclosure.to_string())
                )
                .child(
                    div().text_color(accent).child(var.name.clone())
                )
                .child(
                    div().text_color(text_muted.opacity(0.5)).child(" = ")
                )
                .child(
                    div().text_color(text_primary).child(var.value.clone())
                )
                .into_any_element(),
        );

        // Render expanded children
        if is_expanded && depth < MAX_UI_DEPTH {
            if let Some(children) = expanded_vars.get(&key) {
                let show_count = children.len().min(MAX_RENDER_CHILDREN);
                let remaining = children.len().saturating_sub(MAX_RENDER_CHILDREN);

                for child in children.iter().take(show_count) {
                    let mut child_path = path.to_vec();
                    // Determine key type from child name
                    if let Ok(int_key) = child.name.trim_start_matches('[').trim_end_matches(']').parse::<i64>() {
                        child_path.push(VarPathSegment::KeyInt(int_key));
                    } else {
                        child_path.push(VarPathSegment::KeyString(child.name.clone()));
                    }
                    render_variable_rows(
                        child, depth + 1, &child_path, frame_index, expanded_vars,
                        text_primary, text_muted, accent, cx, rows,
                    );
                }

                if remaining > 0 {
                    let more_indent = (depth + 1) as f32 * 12.0;
                    rows.push(
                        div()
                            .pl(px(more_indent + 4.0))
                            .py(px(1.0))
                            .text_size(px(14.0))
                            .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                            .text_color(text_muted.opacity(0.5))
                            .italic()
                            .child(format!("... +{} more", remaining))
                            .into_any_element(),
                    );
                }
            }
        }
    }

    /// Debug output area: last N lines in small muted text.
    pub(super) fn render_debug_output(
        output: &[OutputEntry],
        _text_primary: Hsla,
        text_muted: Hsla,
        accent: Hsla,
        error_color: Hsla,
        panel_border: Hsla,
    ) -> Div {
        if output.is_empty() {
            return div();
        }

        let start = output.len().saturating_sub(MAX_DEBUG_OUTPUT_LINES);
        let visible = &output[start..];

        div()
            .flex_shrink_0()
            .max_h(px(80.0))
            .overflow_hidden()
            .border_t_1()
            .border_color(panel_border)
            .px_1()
            .py(px(2.0))
            .children(
                visible.iter().enumerate().map(|(i, entry)| {
                    let color = match entry.kind {
                        OutputKind::Error => error_color,
                        OutputKind::Result => accent,
                        OutputKind::System | OutputKind::Stats => text_muted,
                        _ => text_muted.opacity(0.7),
                    };
                    div()
                        .id(ElementId::Name(format!("dbg-out-{}", start + i).into()))
                        .text_size(px(9.0))
                        .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                        .text_color(color)
                        .child(entry.text.clone())
                })
            )
    }
}
