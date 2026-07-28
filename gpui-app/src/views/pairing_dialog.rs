//! Pairing approval dialog: a client (MCP host, CLI) asked to control this
//! workbook. Approving issues a persistent token; denying sends it away.

use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

pub(crate) fn render_pairing_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    let client_name = app
        .pairing_prompt
        .as_ref()
        .map(|p| p.client_name.clone())
        .unwrap_or_default();

    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            div()
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(440.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                .child(
                    div()
                        .text_lg()
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_primary)
                        .child("Allow control of this spreadsheet?"),
                )
                .child(
                    div()
                        .text_sm()
                        .text_color(text_primary)
                        .child(format!("\u{201c}{}\u{201d} is asking to read and edit this workbook.", client_name)),
                )
                .child(
                    div()
                        .text_sm()
                        .text_color(text_muted)
                        .child("Its edits appear live and land in your undo history. Access persists across restarts until revoked (vgrid pair --revoke)."),
                )
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("pairing-deny-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .border_1()
                                .border_color(panel_border)
                                .text_sm()
                                .text_color(text_primary)
                                .cursor_pointer()
                                .hover(|s| s.bg(hsla(0.0, 0.0, 0.5, 0.15)))
                                .child("Deny")
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.respond_pairing(false, cx);
                                })),
                        )
                        .child(
                            div()
                                .id("pairing-allow-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .bg(accent)
                                .text_sm()
                                .text_color(hsla(0.0, 0.0, 1.0, 1.0))
                                .cursor_pointer()
                                .hover(|s| s.opacity(0.85))
                                .child("Allow")
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.respond_pairing(true, cx);
                                })),
                        ),
                ),
        )
}
