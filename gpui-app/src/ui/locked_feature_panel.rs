//! Reusable locked feature panel — shows a dimmed preview + CTA for Free users.
//!
//! Used in the inspector panel to replace hidden Pro features with visible-but-locked
//! previews that drive early-access signups.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::ui::Button;

/// Render an inline locked feature panel with skeleton preview, description, and CTA.
///
/// Returns `None` if the user has dismissed locked panels this session.
pub fn render_locked_feature_panel(
    title: &str,
    description: &str,
    preview: AnyElement,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    text_inverse: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Option<AnyElement> {
    // Session dismiss: user clicked "×" on any locked panel
    if cx.entity().read(cx).locked_panels_dismissed {
        return None;
    }

    let trial = visigrid_license::trial_info();
    let title_owned: SharedString = title.to_string().into();
    let desc_owned: SharedString = description.to_string().into();

    Some(div()
        .mt_2()
        .rounded(px(6.0))
        .bg(panel_border.opacity(0.15))
        .border_1()
        .border_color(panel_border.opacity(0.3))
        .flex()
        .flex_col()
        .overflow_hidden()
        // Title bar: lock icon + title + Pro badge + dismiss ×
        .child(
            div()
                .px_3()
                .py(px(6.0))
                .flex()
                .items_center()
                .gap_2()
                .border_b_1()
                .border_color(panel_border.opacity(0.2))
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child("\u{1F512}")
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(text_primary)
                        .flex_1()
                        .child(title_owned)
                )
                .child(
                    div()
                        .px(px(6.0))
                        .py(px(2.0))
                        .rounded(px(8.0))
                        .bg(accent.opacity(0.15))
                        .text_size(px(9.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(accent)
                        .child("Pro")
                )
                // Session dismiss "×"
                .child(
                    div()
                        .id("locked-panel-dismiss")
                        .ml_1()
                        .px(px(4.0))
                        .py(px(2.0))
                        .rounded_sm()
                        .text_size(px(11.0))
                        .text_color(text_muted.opacity(0.5))
                        .cursor_pointer()
                        .hover(|s| s.text_color(text_muted))
                        .child("\u{00D7}")
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.locked_panels_dismissed = true;
                            cx.notify();
                        }))
                )
        )
        // Preview area (caller-provided, rendered at 20% opacity)
        .child(
            div()
                .px_3()
                .pt_2()
                .opacity(0.2)
                .child(preview)
        )
        // Description
        .child(
            div()
                .px_3()
                .py_2()
                .text_size(px(11.0))
                .text_color(text_muted)
                .line_height(rems(1.5))
                .child(desc_owned)
        )
        // CTA area
        .child(render_cta(trial, accent, text_inverse, text_muted, cx))
        .into_any_element())
}

fn render_cta(
    trial: Option<visigrid_license::TrialInfo>,
    accent: Hsla,
    text_inverse: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let mut cta = div()
        .px_3()
        .pb_3()
        .flex()
        .items_center()
        .gap_2();

    match trial {
        // Trial active — just show badge, no button
        Some(ref info) if !info.expired => {
            cta = cta.child(
                div()
                    .px(px(8.0))
                    .py(px(3.0))
                    .rounded(px(8.0))
                    .bg(accent.opacity(0.12))
                    .text_size(px(10.0))
                    .text_color(accent)
                    .child(SharedString::from(format!(
                        "Trial active \u{2014} {} days left",
                        info.remaining_days
                    )))
            );
        }
        // Trial expired or no trial — early access CTA
        _ => {
            cta = cta
                .child(
                    Button::new("locked-early-access-btn", "Request Early Access")
                        .primary(accent, text_inverse)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.show_license(cx);
                        }))
                )
                .child(
                    div()
                        .id("locked-whats-pro-btn")
                        .text_size(px(10.0))
                        .text_color(accent)
                        .cursor_pointer()
                        .hover(|s| s.opacity(0.8))
                        .child("What\u{2019}s Pro?")
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.show_license(cx);
                        }))
                );
        }
    }

    cta
}
