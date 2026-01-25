//! Import overlay for background Excel imports
//!
//! Shows a lightweight overlay when Excel imports take longer than 150ms.
//! Allows user to dismiss with ESC (import continues in background).

use gpui::*;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the import overlay
pub fn render_import_overlay(app: &Spreadsheet, _cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    let filename = app.import_filename.as_deref().unwrap_or("Excel file");

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        .child(
            div()
                .w(px(320.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_xl()
                .overflow_hidden()
                .flex()
                .flex_col()
                .p_4()
                .gap_3()
                // Title
                .child(
                    div()
                        .text_size(px(14.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(format!("Importing {}...", filename))
                )
                // Indeterminate progress bar
                .child(
                    div()
                        .w_full()
                        .h(px(4.0))
                        .rounded_full()
                        .bg(accent.opacity(0.2))
                        .overflow_hidden()
                        .child(
                            // Animated shimmer effect using a gradient
                            div()
                                .w(px(80.0))
                                .h_full()
                                .bg(accent)
                                .rounded_full()
                        )
                )
                // Info text
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child("This may take a few seconds for large files.")
                )
                // ESC hint
                .child(
                    div()
                        .mt_1()
                        .text_size(px(10.0))
                        .text_color(text_muted.opacity(0.7))
                        .child("Press Esc to dismiss")
                )
        )
}
