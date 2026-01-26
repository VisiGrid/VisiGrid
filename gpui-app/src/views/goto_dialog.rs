use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::modal_backdrop;

/// Render the Go To cell dialog overlay
pub fn render_goto_dialog(app: &Spreadsheet) -> impl IntoElement {
    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let app_bg = app.token(TokenKey::AppBg);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);

    modal_backdrop(
        "goto-dialog",
        div()
            .w(px(300.0))
            .bg(panel_bg)
            .border_1()
            .border_color(accent)
            .rounded_md()
            .p_4()
            .flex()
            .flex_col()
            .gap_2()
            // Title
            .child(
                    div()
                        .text_color(text_primary)
                        .font_weight(FontWeight::MEDIUM)
                        .child("Go To Cell")
                )
                // Input field
                .child(
                    div()
                        .w_full()
                        .h(px(32.0))
                        .bg(app_bg)
                        .border_1()
                        .border_color(panel_border)
                        .rounded_sm()
                        .px_2()
                        .flex()
                        .items_center()
                        .text_color(text_primary)
                        .child(format!("{}|", app.goto_input))
                )
                // Help text
                .child(
                    div()
                        .text_color(text_muted)
                        .text_sm()
                        .child("Enter cell reference (e.g., A1, B25, AA100)")
                )
                // Instructions
                .child(
                    div()
                        .text_color(text_disabled)
                        .text_xs()
                        .child("Enter to confirm, Escape to cancel")
            )
    )
}
