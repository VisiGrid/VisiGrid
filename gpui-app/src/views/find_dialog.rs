use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the Find dialog overlay
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

    div()
        .absolute()
        .top_2()
        .right_2()
        .w(px(300.0))
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
                .child("Find")
        )
        // Input field with result count
        .child(
            div()
                .flex()
                .gap_2()
                .child(
                    div()
                        .flex_1()
                        .h(px(28.0))
                        .bg(app_bg)
                        .border_1()
                        .border_color(panel_border)
                        .rounded_sm()
                        .px_2()
                        .flex()
                        .items_center()
                        .text_color(text_primary)
                        .text_sm()
                        .child(format!("{}|", app.find_input))
                )
                .child(
                    div()
                        .text_color(text_muted)
                        .text_sm()
                        .flex()
                        .items_center()
                        .child(result_info)
                )
        )
        // Instructions
        .child(
            div()
                .text_color(text_disabled)
                .text_xs()
                .child("F3 next, Shift+F3 prev, Escape to close")
        )
}
