//! Add Conditional Format dialog: one text input, typed rule syntax.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;

pub(crate) fn render_add_cond_format_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let range_label = app
        .cf_target
        .first()
        .map(|r| {
            format!(
                "{}{}:{}{}",
                col_letter(r.start_col),
                r.start_row + 1,
                col_letter(r.end_col),
                r.end_row + 1
            )
        })
        .unwrap_or_default();

    let input_display = if app.cf_input.is_empty() {
        None
    } else {
        Some(app.cf_input.clone())
    };
    let error = app.cf_input_error.clone();

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(520.0))
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
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_sm()
                                .font_weight(FontWeight::SEMIBOLD)
                                .child("Conditional Format")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_sm()
                                .child(range_label)
                        )
                )
                // Input line
                .child(
                    div()
                        .px_2()
                        .py_1()
                        .border_1()
                        .border_color(if error.is_some() { error_color } else { accent })
                        .rounded_sm()
                        .text_sm()
                        .font_family("IBM Plex Mono")
                        .when_some(input_display, |d, text| {
                            d.text_color(text_primary).child(text)
                        })
                        .when(app.cf_input.is_empty(), |d| {
                            d.text_color(text_muted)
                                .child("=A1>100 -> warning")
                        })
                )
                // Error or hint line
                .child(match error {
                    Some(e) => div().text_color(error_color).text_sm().child(e),
                    None => div()
                        .text_color(text_muted)
                        .text_sm()
                        .child(
                            "predicate -> style · styles: warning, error, success, note, \
                             bold, bg=#RRGGBB, fg=#RRGGBB, like(Z1) · Enter to add · Esc to cancel",
                        ),
                })
        )
}

fn col_letter(mut col: usize) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    s
}
