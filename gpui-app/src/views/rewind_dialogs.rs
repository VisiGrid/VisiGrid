use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the rewind confirmation dialog (Phase 8C)
pub(crate) fn render_rewind_confirm_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let discard_count = app.rewind_confirm.discard_count;
    let target_summary = app.rewind_confirm.target_summary.clone();
    let sheet_name = app.rewind_confirm.sheet_name.clone();
    let location = app.rewind_confirm.location.clone();
    let replay_count = app.rewind_confirm.replay_count;
    let build_ms = app.rewind_confirm.build_ms;

    // Build location badge text: "Sheet1!A1:B10" or just "A1:B10" or just "Sheet1"
    let location_badge = match (&sheet_name, &location) {
        (Some(sheet), Some(loc)) => Some(format!("{}!{}", sheet, loc)),
        (Some(sheet), None) => Some(sheet.clone()),
        (None, Some(loc)) => Some(loc.clone()),
        (None, None) => None,
    };

    // Dark red colors for destructive action
    let danger_bg = hsla(0.0, 0.8, 0.3, 0.15);       // Dark red background
    let danger_border = hsla(0.0, 0.8, 0.4, 0.3);    // Dark red border
    let danger_button = hsla(0.0, 0.8, 0.4, 1.0);    // Red button
    let danger_button_hover = hsla(0.0, 0.8, 0.5, 1.0); // Brighter red on hover

    // Modal backdrop
    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            // Dialog box
            div()
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(420.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_4()
                // Header with warning icon
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(20.0))
                                .text_color(error_color)
                                .child("\u{26A0}") // Warning triangle
                        )
                        .child(
                            div()
                                .text_size(px(16.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Rewind History")
                        )
                )
                // Target info with location badge
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        // Location badge (if available)
                        .when_some(location_badge, |el, badge| {
                            el.child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .px_2()
                                            .py(px(2.0))
                                            .bg(accent.opacity(0.15))
                                            .border_1()
                                            .border_color(accent.opacity(0.3))
                                            .rounded_sm()
                                            .text_size(px(11.0))
                                            .font_weight(FontWeight::MEDIUM)
                                            .text_color(accent)
                                            .child(SharedString::from(badge))
                                    )
                                    .child(
                                        div()
                                            .text_size(px(11.0))
                                            .text_color(text_muted)
                                            .child("Target location")
                                    )
                            )
                        })
                        // Main message
                        .child(
                            div()
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .child(SharedString::from(format!(
                                    "This will permanently discard {} action{}.",
                                    discard_count,
                                    if discard_count == 1 { "" } else { "s" }
                                )))
                        )
                        // Target action summary
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!(
                                    "Rewind to state before: \"{}\"",
                                    target_summary
                                )))
                        )
                )
                // Destructive warning box
                .child(
                    div()
                        .px_3()
                        .py_2()
                        .bg(danger_bg)
                        .border_1()
                        .border_color(danger_border)
                        .rounded_sm()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(error_color)
                                .child("This action cannot be undone. The discarded changes will be permanently lost.")
                        )
                )
                // Performance info (subtle)
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted.opacity(0.7))
                        .child(SharedString::from(format!(
                            "Preview: {} actions replayed in {}ms",
                            replay_count,
                            build_ms
                        )))
                )
                // Buttons
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("rewind-cancel-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .bg(panel_border.opacity(0.3))
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .cursor_pointer()
                                .hover(|s| s.bg(panel_border.opacity(0.5)))
                                .child("Cancel")
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.cancel_rewind(cx);
                                }))
                        )
                        .child(
                            div()
                                .id("rewind-confirm-btn")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .bg(danger_button)
                                .text_size(px(13.0))
                                .text_color(gpui::white())
                                .cursor_pointer()
                                .hover(|s| s.bg(danger_button_hover))
                                .child(SharedString::from(format!(
                                    "Discard {} action{}",
                                    discard_count,
                                    if discard_count == 1 { "" } else { "s" }
                                )))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.confirm_rewind(cx);
                                }))
                        )
                )
        )
}

/// Render the rewind success banner (Phase 8C)
/// Shows briefly after rewind with "Copy details" button
pub(crate) fn render_rewind_success_banner(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let success_color = hsla(0.35, 0.7, 0.5, 1.0); // Green

    let discarded_count = app.rewind_success.discarded_count;
    let target_summary = app.rewind_success.target_summary.clone();

    // Top banner that slides in from top
    div()
        .absolute()
        .top_0()
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .pt_2()
        .child(
            div()
                .id("rewind-success-banner")
                .bg(panel_bg)
                .border_1()
                .border_color(success_color.opacity(0.5))
                .rounded_md()
                .shadow_lg()
                .px_4()
                .py_2()
                .flex()
                .items_center()
                .gap_3()
                // Success icon
                .child(
                    div()
                        .text_size(px(16.0))
                        .text_color(success_color)
                        .child("\u{2713}") // Checkmark
                )
                // Message
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .child(
                            div()
                                .text_size(px(13.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text_primary)
                                .child(SharedString::from(format!(
                                    "Rewound. Discarded {} action{}. Undo not available.",
                                    discarded_count,
                                    if discarded_count == 1 { "" } else { "s" }
                                )))
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!(
                                    "Before: \"{}\"",
                                    target_summary
                                )))
                        )
                )
                // Copy audit button
                .child(
                    div()
                        .id("copy-rewind-audit-btn")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.3))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
                        .child("Copy audit")
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.copy_rewind_details(cx);
                        }))
                )
                // Dismiss button
                .child(
                    div()
                        .id("dismiss-rewind-banner-btn")
                        .px_2()
                        .py_1()
                        .text_size(px(14.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.text_color(text_primary))
                        .child("\u{2715}") // X mark
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.dismiss_rewind_banner(cx);
                        }))
                )
        )
}
