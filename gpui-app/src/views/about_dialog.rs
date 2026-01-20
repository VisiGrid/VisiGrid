use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the About VisiGrid dialog overlay
pub fn render_about_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_about(cx);
        }))
        .child(
            div()
                .w(px(340.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_xl()
                .overflow_hidden()
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .child(
                    // Content
                    div()
                        .p_6()
                        .flex()
                        .flex_col()
                        .items_center()
                        .gap_4()
                        // Logo/Title
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .items_center()
                                .gap_1()
                                .child(
                                    div()
                                        .text_size(px(24.0))
                                        .font_weight(FontWeight::BOLD)
                                        .text_color(text_primary)
                                        .child("VisiGrid")
                                )
                                .child(
                                    div()
                                        .text_size(px(12.0))
                                        .text_color(text_muted)
                                        .child("A modern spreadsheet application")
                                )
                        )
                        // Version info
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .items_center()
                                .gap_1()
                                .child(
                                    div()
                                        .text_size(px(13.0))
                                        .text_color(text_primary)
                                        .child("Version 0.1.0")
                                )
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .child("Built with GPUI")
                                )
                        )
                        // Features
                        .child(
                            div()
                                .w_full()
                                .py_3()
                                .border_t_1()
                                .border_b_1()
                                .border_color(panel_border)
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(feature_item("Native Rust performance", text_muted))
                                .child(feature_item("Excel-compatible formulas", text_muted))
                                .child(feature_item("CSV and native file formats", text_muted))
                                .child(feature_item("Themeable interface", text_muted))
                        )
                        // Copyright
                        .child(
                            div()
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .child("MIT License")
                        )
                )
                // Close button footer
                .child(
                    div()
                        .w_full()
                        .px_4()
                        .py_3()
                        .bg(panel_bg)
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_center()
                        .child(
                            div()
                                .id("about-close-btn")
                                .px_6()
                                .py(px(6.0))
                                .bg(accent)
                                .rounded_md()
                                .cursor_pointer()
                                .text_size(px(12.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(app.token(TokenKey::TextInverse))
                                .hover(|s| s.opacity(0.9))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_about(cx);
                                }))
                                .child("Close")
                        )
                )
        )
}

fn feature_item(text: &'static str, color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .gap_2()
        .text_size(px(11.0))
        .text_color(color)
        .child(
            div()
                .w(px(4.0))
                .h(px(4.0))
                .rounded_full()
                .bg(color.opacity(0.5))
        )
        .child(text)
}
