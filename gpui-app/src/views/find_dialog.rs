use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the Find/Replace dialog overlay
pub fn render_find_dialog(app: &Spreadsheet) -> impl IntoElement {
    let result_info = if app.find_results.is_empty() {
        if app.find_input.is_empty() {
            String::new()
        } else {
            "No matches".to_string()
        }
    } else {
        format!("{} of {}", app.find_index + 1, app.find_results.len())
    };

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let app_bg = app.token(TokenKey::AppBg);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);

    let title = if app.find_replace_mode {
        "Find and Replace"
    } else {
        "Find"
    };

    // Input border color based on focus
    let find_border = if !app.find_focus_replace { accent } else { panel_border };
    let replace_border = if app.find_focus_replace { accent } else { panel_border };

    // Instructions based on mode
    let instructions = if app.find_replace_mode {
        "Tab switch, Enter replace, Ctrl+Enter all, Esc close"
    } else {
        "F3 next, Shift+F3 prev, Escape to close"
    };

    div()
        .key_context("FindDialog")
        .absolute()
        .top_2()
        .right_2()
        .w(px(340.0))
        .bg(panel_bg)
        .border_1()
        .border_color(accent)
        .rounded_md()
        .p_3()
        .flex()
        .flex_col()
        .gap_2()
        // Title
        .child(
            div()
                .text_color(text_primary)
                .font_weight(FontWeight::MEDIUM)
                .text_sm()
                .child(title)
        )
        // Find input row
        .child(
            div()
                .flex()
                .gap_2()
                .items_center()
                .child(
                    div()
                        .w(px(60.0))
                        .text_color(text_muted)
                        .text_sm()
                        .child("Find:")
                )
                .child(
                    div()
                        .flex_1()
                        .h(px(28.0))
                        .bg(app_bg)
                        .border_1()
                        .border_color(find_border)
                        .rounded_sm()
                        .px_2()
                        .flex()
                        .items_center()
                        .text_color(text_primary)
                        .text_sm()
                        .child(if !app.find_focus_replace {
                            format!("{}|", app.find_input)
                        } else {
                            app.find_input.clone()
                        })
                )
                .child(
                    div()
                        .w(px(60.0))
                        .text_color(text_muted)
                        .text_sm()
                        .text_right()
                        .child(result_info)
                )
        )
        // Replace input row (only in replace mode)
        .when(app.find_replace_mode, |el| {
            el.child(
                div()
                    .flex()
                    .gap_2()
                    .items_center()
                    .child(
                        div()
                            .w(px(60.0))
                            .text_color(text_muted)
                            .text_sm()
                            .child("Replace:")
                    )
                    .child(
                        div()
                            .flex_1()
                            .h(px(28.0))
                            .bg(app_bg)
                            .border_1()
                            .border_color(replace_border)
                            .rounded_sm()
                            .px_2()
                            .flex()
                            .items_center()
                            .text_color(text_primary)
                            .text_sm()
                            .child(if app.find_focus_replace {
                                format!("{}|", app.replace_input)
                            } else {
                                app.replace_input.clone()
                            })
                    )
                    .child(
                        div()
                            .w(px(60.0))
                    )
            )
        })
        // Buttons row (only in replace mode)
        .when(app.find_replace_mode, |el| {
            el.child(
                div()
                    .flex()
                    .gap_2()
                    .justify_end()
                    .child(
                        div()
                            .px_3()
                            .py_1()
                            .bg(app_bg)
                            .border_1()
                            .border_color(panel_border)
                            .rounded_sm()
                            .text_sm()
                            .text_color(text_primary)
                            .cursor_pointer()
                            .hover(|s| s.bg(panel_border))
                            .child("Replace")
                    )
                    .child(
                        div()
                            .px_3()
                            .py_1()
                            .bg(app_bg)
                            .border_1()
                            .border_color(panel_border)
                            .rounded_sm()
                            .text_sm()
                            .text_color(text_primary)
                            .cursor_pointer()
                            .hover(|s| s.bg(panel_border))
                            .child("Replace All")
                    )
            )
        })
        // Instructions
        .child(
            div()
                .text_color(text_disabled)
                .text_xs()
                .child(instructions)
        )
        // Contextual tip: suggest Replace when in Find-only mode with content
        .when(!app.find_replace_mode && !app.find_input.is_empty(), |el| {
            el.child(
                div()
                    .text_color(text_muted)
                    .text_xs()
                    .child("Tip: Ctrl+H to replace")
            )
        })
}
