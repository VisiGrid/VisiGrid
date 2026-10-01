use crate::{app::Spreadsheet, theme::TokenKey, ui::modal_overlay};
use gpui::prelude::FluentBuilder;
use gpui::*;

/// A focused cell destination picker. Navigation remains owned by dialogs.rs.
pub fn render_goto_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let empty = app.goto_input.trim().is_empty();
    let destination = Spreadsheet::parse_goto_destination(&app.goto_input);
    let valid = destination.is_ok();
    let invalid = !empty && !valid;
    let feedback = if empty {
        "Use a column letter and row number, such as B25.".to_string()
    } else {
        match destination {
            Ok((row, col)) => format!("Ready to jump to {}", app.cell_ref_at(row, col)),
            Err(message) => message.to_string(),
        }
    };

    let content = div()
        .w(px(460.0))
        .p_6()
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(border)
        .rounded_md()
        .shadow_lg()
        .flex()
        .flex_col()
        .gap_5()
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text)
                        .child("Go to cell"),
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(muted)
                        .child("Jump to a cell on the current sheet."),
                ),
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(muted)
                                .child("Cell reference"),
                        )
                        .child(
                            div()
                                .px_2()
                                .py_1()
                                .rounded_sm()
                                .bg(app.token(TokenKey::EditorBg))
                                .text_size(px(11.0))
                                .text_color(muted)
                                .child(format!("Current cell  {}", app.cell_ref())),
                        ),
                )
                .child(
                    div()
                        .id("goto-reference-input")
                        .h(px(44.0))
                        .px_3()
                        .bg(app.token(TokenKey::EditorBg))
                        .border_1()
                        .rounded_sm()
                        .border_color(if invalid {
                            app.token(TokenKey::Error)
                        } else {
                            accent
                        })
                        .flex()
                        .items_center()
                        .overflow_hidden()
                        .cursor_text()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|s, _, window, cx| {
                                cx.stop_propagation();
                                window.focus(&s.focus_handle, cx);
                            }),
                        )
                        .child(
                            div()
                                .min_w_0()
                                .overflow_hidden()
                                .text_ellipsis()
                                .text_size(px(16.0))
                                .text_color(text)
                                .child(app.goto_input.clone()),
                        )
                        .child(div().w(px(1.0)).h(px(18.0)).flex_shrink_0().bg(accent))
                        .when(empty, |s| {
                            s.child(
                                div()
                                    .ml_1()
                                    .text_size(px(16.0))
                                    .text_color(muted)
                                    .child("e.g. B25"),
                            )
                        }),
                ),
        )
        .child(
            div()
                .min_h(px(32.0))
                .text_size(px(12.0))
                .text_color(if invalid {
                    app.token(TokenKey::Error)
                } else {
                    muted
                })
                .child(feedback),
        )
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .pt_4()
                .border_t_1()
                .border_color(border)
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(muted)
                        .child("Enter to go · Esc to cancel"),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .id("goto-cancel")
                                .px_3()
                                .py_2()
                                .rounded_md()
                                .text_size(px(13.0))
                                .text_color(text)
                                .cursor_pointer()
                                .hover(|s| s.bg(app.token(TokenKey::SelectionBg)))
                                .on_mouse_down(
                                    MouseButton::Left,
                                    cx.listener(|s, _, _, cx| {
                                        cx.stop_propagation();
                                        s.hide_goto(cx);
                                    }),
                                )
                                .child("Cancel"),
                        )
                        .child(
                            div()
                                .id("goto-submit")
                                .px_4()
                                .py_2()
                                .rounded_md()
                                .text_size(px(13.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .bg(if valid {
                                    accent
                                } else {
                                    app.token(TokenKey::EditorBg)
                                })
                                .text_color(if valid {
                                    app.token(TokenKey::TextInverse)
                                } else {
                                    app.token(TokenKey::TextDisabled)
                                })
                                .when(valid, |s| {
                                    s.cursor_pointer().hover(|s| s.bg(accent.opacity(0.85)))
                                })
                                .on_mouse_down(
                                    MouseButton::Left,
                                    cx.listener(move |s, _, _, cx| {
                                        cx.stop_propagation();
                                        if valid {
                                            s.confirm_goto(cx);
                                        }
                                    }),
                                )
                                .child("Go to cell"),
                        ),
                ),
        );
    modal_overlay("goto-dialog", |s, cx| s.hide_goto(cx), content, cx)
}
