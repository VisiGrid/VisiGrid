use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

const VERSION: &str = env!("CARGO_PKG_VERSION");
const GIT_COMMIT: &str = env!("GIT_COMMIT_HASH");

/// Render the About VisiGrid dialog overlay
pub fn render_about_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let diagnostics = format!(
        "VisiGrid {}\nCommit: {}\nOS: {} {}\nBuild: {}\nSheet format: v{}",
        VERSION,
        GIT_COMMIT,
        std::env::consts::OS,
        std::env::consts::ARCH,
        if cfg!(debug_assertions) { "debug" } else { "release" },
        visigrid_io::NATIVE_FORMAT_VERSION
    );

    // Body content
    let body = div()
        // Title
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(22.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_primary)
                        .child("VisiGrid")
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(text_muted)
                        .child("A fast, local-first spreadsheet")
                )
        )
        // Philosophy
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .line_height(rems(1.5))
                .child("Built for analysts, engineers, and anyone who lives in spreadsheets. Your data stays on your machine, in files you control.")
        )
        // Principles
        .child(
            div()
                .py_3()
                .border_t_1()
                .border_b_1()
                .border_color(panel_border)
                .flex()
                .flex_col()
                .gap_2()
                .child(principle_item("Local-first", "files, not cloud documents", text_primary, text_muted))
                .child(principle_item("Keyboard-driven", "speed matters", text_primary, text_muted))
                .child(principle_item("Deterministic", "formulas do exactly what they say", text_primary, text_muted))
                .child(principle_item("Transparent", "no telemetry, no background sync", text_primary, text_muted))
        )
        // Status
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Early preview. Core functionality is stable; advanced features in progress.")
        )
        // Diagnostics section
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
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .font_family("monospace")
                                .child(format!("v{} ({})", VERSION, GIT_COMMIT))
                        )
                        .child(
                            div()
                                .id("copy-diagnostics-btn")
                                .px_2()
                                .py_1()
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .rounded(px(3.0))
                                .border_1()
                                .border_color(panel_border)
                                .cursor_pointer()
                                .hover(|s| s.bg(panel_border))
                                .on_mouse_down(MouseButton::Left, {
                                    let diagnostics = diagnostics.clone();
                                    move |_, _, cx| {
                                        cx.write_to_clipboard(ClipboardItem::new_string(diagnostics.clone()));
                                        cx.stop_propagation();
                                    }
                                })
                                .child("Copy diagnostics")
                        )
                )
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("Open source (AGPL-3.0)")
                )
        );

    // Footer with centered close button
    let footer = div()
        .flex()
        .justify_center()
        .child(
            Button::new("about-close-btn", "Close")
                .primary(accent, text_inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.hide_about(cx);
                }))
        );

    modal_overlay(
        "about-dialog",
        |this, cx| this.hide_about(cx),
        DialogFrame::new(body, panel_bg, panel_border)
            .size(DialogSize::Md)
            .footer(footer),
        cx,
    )
}

fn principle_item(title: &'static str, description: &'static str, title_color: Hsla, desc_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_baseline()
        .gap_2()
        .child(
            div()
                .text_size(px(11.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(title_color)
                .child(title)
        )
        .child(
            div()
                .text_size(px(11.0))
                .text_color(desc_color)
                .child(format!("— {}", description))
        )
}
