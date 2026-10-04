//! Conditional-format rule authoring, with live worksheet preview.

use gpui::{prelude::FluentBuilder, *};

use crate::{
    app::Spreadsheet, cond_format_ui::format_range_label, theme::TokenKey, ui::modal_overlay,
};

pub(crate) fn render_add_cond_format_dialog(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let bg = app.token(TokenKey::PanelBg);
    let input_bg = app.token(TokenKey::EditorBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);
    let editing = app.cf_draft.as_ref().is_some_and(|d| d.editing.is_some());
    let scope = if editing {
        "Edits apply to the full stored rule, including hidden cells."
    } else if app.cf_target.len() > 1 {
        "Highlight visible selected cells. Separate ranges appear as separate rules."
    } else {
        "Highlight selected cells when a rule is true."
    };
    let empty = app.cf_input.trim().is_empty();
    let error = app.cf_input_error.clone();
    let range_label = format_range_label(&app.cf_target);
    let example_ref = app
        .cf_draft
        .as_ref()
        .map(|d| app.cell_ref_at(d.anchor.0, d.anchor.1))
        .unwrap_or_else(|| "A1".into());
    let total: usize = app
        .cf_target
        .iter()
        .map(|r| (r.end_row - r.start_row + 1).saturating_mul(r.end_col - r.start_col + 1))
        .fold(0usize, usize::saturating_add);
    let (feedback_title, feedback_detail, feedback_color) = match (&error, app.cf_preview_matches) {
        (Some(message), _) => ("Check your rule".to_string(), message.clone(), error_color),
        (None, Some((matching, scanned))) => (
            format!("{} of {} cells match", matching, scanned),
            if scanned < total {
                format!(
                    "Live preview on the worksheet · first {} cells checked.",
                    scanned
                )
            } else {
                "Live preview on the worksheet. Add the rule to keep this formatting.".into()
            },
            accent,
        ),
        _ => (
            "Preview appears as you type".into(),
            "Enter a formula and a style to highlight matching cells on the worksheet.".into(),
            muted,
        ),
    };

    let content = div()
        .w(px(560.0))
        .p_6()
        .bg(bg)
        .border_1()
        .border_color(border)
        .rounded_lg()
        .shadow_lg()
        .flex()
        .flex_col()
        .gap_5()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
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
                        .child(if editing {
                            "Edit conditional formatting"
                        } else {
                            "Conditional formatting"
                        }),
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(muted)
                        .whitespace_normal()
                        .child(scope),
                ),
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_3()
                .px_3()
                .py_2()
                .bg(input_bg)
                .rounded_md()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(muted)
                        .child("Applies to"),
                )
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .text_size(px(13.0))
                        .text_color(text)
                        .font_family("IBM Plex Mono")
                        .whitespace_normal()
                        .child(range_label),
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
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text)
                                .child("Rule"),
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(muted)
                                .child("Formula → style"),
                        ),
                )
                .child(
                    div()
                        .id("conditional-format-rule-input")
                        .h(px(44.0))
                        .px_3()
                        .bg(input_bg)
                        .border_1()
                        .rounded_md()
                        .border_color(if error.is_some() { error_color } else { accent })
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
                                .font_family("IBM Plex Mono")
                                .text_size(px(14.0))
                                .text_color(text)
                                .child(app.cf_input.clone()),
                        )
                        .child(div().w(px(1.0)).h(px(18.0)).flex_shrink_0().bg(accent))
                        .when(empty, |d| {
                            d.child(
                                div()
                                    .ml_1()
                                    .text_size(px(14.0))
                                    .font_family("IBM Plex Mono")
                                    .text_color(muted)
                                    .child(format!("={} > 100 -> warning", example_ref)),
                            )
                        }),
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(muted)
                        .whitespace_normal()
                        .child(if editing {
                            "Separate the formula and style with ->. Relative references start at each range’s top-left cell.".into()
                        } else {
                            format!("Write the formula for {example_ref} (stored address), followed by -> and a style.")
                        }),
                ),
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .text_size(px(12.0))
                .text_color(muted)
                .child(
                    div()
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(text)
                        .child("Style options"),
                )
                .child("warning · success · error · input · total · note")
                .child(div().whitespace_normal().child(
                    "Or use bold, bg=#FFEB3B, fg=#333333, or like(Z1) to copy a cell’s style.",
                )),
        )
        .child(
            div()
                .min_h(px(72.0))
                .px_3()
                .py_3()
                .rounded_md()
                .bg(feedback_color.opacity(0.07))
                .border_1()
                .border_color(feedback_color.opacity(0.18))
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(feedback_color)
                        .child(feedback_title),
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(if error.is_some() { error_color } else { muted })
                        .whitespace_normal()
                        .child(feedback_detail),
                ),
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
                        .child("Enter to apply · Esc to cancel"),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .id("conditional-format-cancel")
                                .px_3()
                                .py_2()
                                .rounded_md()
                                .text_size(px(13.0))
                                .text_color(text)
                                .cursor_pointer()
                                .hover(move |s| s.bg(input_bg))
                                .on_mouse_down(
                                    MouseButton::Left,
                                    cx.listener(|s, _, _, cx| {
                                        cx.stop_propagation();
                                        s.hide_add_cond_format(cx);
                                    }),
                                )
                                .child("Cancel"),
                        )
                        .child(
                            div()
                                .id("conditional-format-submit")
                                .px_4()
                                .py_2()
                                .rounded_md()
                                .text_size(px(13.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .bg(if empty { input_bg } else { accent })
                                .text_color(app.token(if empty {
                                    TokenKey::TextDisabled
                                } else {
                                    TokenKey::TextInverse
                                }))
                                .when(!empty, |d| {
                                    d.cursor_pointer()
                                        .hover(move |s| s.bg(accent.opacity(0.85)))
                                })
                                .on_mouse_down(
                                    MouseButton::Left,
                                    cx.listener(move |s, _, _, cx| {
                                        cx.stop_propagation();
                                        if !empty {
                                            s.confirm_add_cond_format(cx);
                                        }
                                    }),
                                )
                                .child(if editing { "Save rule" } else { "Add rule" }),
                        ),
                ),
        );
    modal_overlay(
        "conditional-format-dialog",
        |s, cx| s.hide_add_cond_format(cx),
        content,
        cx,
    )
}
