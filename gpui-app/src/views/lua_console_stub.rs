//! Stub Lua console for Free edition.
//!
//! Shows an upgrade prompt instead of the full Lua REPL.

use gpui::*;

use crate::app::Spreadsheet;
use crate::scripting::{OutputEntry, OutputKind};
use crate::theme::TokenKey;

/// Render the Lua console panel (upgrade prompt in Free edition)
pub fn render_lua_console(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let console = &app.lua_console;

    if !console.visible {
        return div().into_any_element();
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);

    let console_height = console.height;

    div()
        .id("lua-console-panel")
        .key_context("LuaConsole")
        .track_focus(&app.console_focus_handle)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            window.focus(&this.console_focus_handle, cx);
            cx.notify();
        }))
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            // Only handle Escape to close
            if event.keystroke.key.as_str() == "escape" {
                this.lua_console.hide();
                window.focus(&this.focus_handle, cx);
            }
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
            // Header bar
            div()
                .h(px(24.0))
                .px_2()
                .flex()
                .items_center()
                .justify_between()
                .border_b_1()
                .border_color(panel_border)
                .child(
                    div()
                        .text_xs()
                        .text_color(text_muted)
                        .child("Lua Console")
                )
                .child(
                    // Close button
                    div()
                        .id("lua-console-close")
                        .px_1()
                        .cursor_pointer()
                        .text_xs()
                        .text_color(text_muted)
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.lua_console.hide();
                            cx.notify();
                        }))
                        .child("×")
                )
        )
        .child(
            // Output area - shows upgrade message
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
                        .children(
                            console.visible_output().iter().enumerate().map(|(i, entry)| {
                                render_output_entry(entry, console.view_start + i, text_primary, text_muted, accent, error_color)
                            })
                        )
                )
        )
        .child(
            // Footer with upgrade link
            div()
                .h(px(32.0))
                .px_2()
                .flex()
                .items_center()
                .justify_center()
                .border_t_1()
                .border_color(panel_border)
                .child(
                    div()
                        .id("upgrade-link")
                        .px_3()
                        .py_1()
                        .rounded_md()
                        .bg(accent)
                        .text_xs()
                        .text_color(rgb(0xffffff))
                        .cursor_pointer()
                        .hover(|s| s.opacity(0.9))
                        .on_mouse_down(MouseButton::Left, cx.listener(|_this, _, _, _cx| {
                            let _ = open::that("https://visigrid.com/pro");
                        }))
                        .child("Upgrade to Pro")
                )
        )
        .into_any_element()
}

/// Render a single output entry
fn render_output_entry(
    entry: &OutputEntry,
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
        OutputKind::System => ("", text_muted),
    };

    div()
        .id(ElementId::Name(format!("lua-output-{}", index).into()))
        .text_xs()
        .font_family("monospace")
        .text_color(color)
        .child(format!("{}{}", prefix, entry.text))
}

/// Stub - no-op in Free edition
pub fn execute_console(_app: &mut Spreadsheet, _cx: &mut Context<Spreadsheet>) {
    // No-op - Pro feature
}

/// Stub - no-op in Free edition
pub fn handle_console_key_from_main(_app: &mut Spreadsheet, _event: &KeyDownEvent, _window: &mut Window, _cx: &mut Context<Spreadsheet>) {
    // No-op - Pro feature
}
