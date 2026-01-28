//! KeyTips overlay for macOS (Option+Space accelerators)
//!
//! Shows a small overlay with accelerator hints when user presses Option+Space.
//! Pressing a letter key opens the scoped command palette for that category.
//!
//! STABLE MAPPING: F/E/V/O/D/T/H letters are locked. Commands may be added
//! to categories, but category letters will not change.

use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the KeyTips overlay (shown after Option double-tap on macOS)
pub fn render_keytips_overlay(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    // KeyTips entries: (label, key)
    // Note: Matches available MenuCategory variants (no Insert category yet)
    let entries = [
        ("File", 'F'),
        ("Edit", 'E'),
        ("View", 'V'),
        ("Format", 'O'),
        ("Data", 'D'),
        ("Tools", 'T'),
        ("Help", 'H'),
    ];

    // Build the entries as a horizontal row of kbd-style chips
    let items = entries.iter().map(|(label, key)| {
        div()
            .flex()
            .items_center()
            .gap_1()
            .child(
                div()
                    .text_size(px(12.0))
                    .text_color(text_primary)
                    .child(*label)
            )
            .child(
                // kbd-style chip for the accelerator key
                div()
                    .px(px(6.0))
                    .py(px(2.0))
                    .bg(accent.opacity(0.15))
                    .border_1()
                    .border_color(accent.opacity(0.3))
                    .rounded(px(3.0))
                    .text_size(px(11.0))
                    .font_weight(FontWeight::MEDIUM)
                    .text_color(accent)
                    .child(key.to_string())
            )
    }).collect::<Vec<_>>();

    // Centered overlay card
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(80.0))  // Position below toolbar area
        .child(
            div()
                .id("keytips-overlay")
                .px_4()
                .py_3()
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_lg()
                .flex()
                .flex_wrap()
                .gap_3()
                .children(items)
                .child(
                    // Hint text
                    div()
                        .w_full()
                        .pt_2()
                        .border_t_1()
                        .border_color(panel_border)
                        .mt_1()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .text_center()
                        .child("Press a key · Enter repeats last · Esc cancels")
                )
                // Click anywhere on the overlay dismisses it
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.dismiss_keytips(cx);
                }))
        )
}
