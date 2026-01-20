//! Tour modal for Named Ranges & Refactoring walkthrough

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Tour step content
struct TourStep {
    title: &'static str,
    body: &'static str,
    hint: &'static str,
}

const TOUR_STEPS: [TourStep; 4] = [
    TourStep {
        title: "Give meaning to your data",
        body: "Instead of remembering cell addresses like A1:A100, you can give important data a name.\n\nSelect any range and press Ctrl+Shift+N to create a named range.",
        hint: "Try it now — select a range and press Ctrl+Shift+N",
    },
    TourStep {
        title: "Use names like variables",
        body: "Named ranges work anywhere you write formulas.\n\nTry typing:\n=SUM(Revenue)\n\nAutocomplete will suggest your named range automatically.",
        hint: "Names are case-insensitive and safer than cell references",
    },
    TourStep {
        title: "Refactor safely",
        body: "Need to change a name later? No problem.\n\nRename a named range and every formula updates automatically — as a single undoable change.",
        hint: "Try Ctrl+Shift+R on a named range to rename it",
    },
    TourStep {
        title: "Jump to definitions and references",
        body: "Named ranges are symbols you can navigate.\n\n• F12 — jump to where a name is defined\n• Shift+F12 — see everywhere it's used",
        hint: "You can also search names instantly with $ in the command palette",
    },
];

/// Render the tour modal
pub fn render_tour(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    let step = app.tour_step;
    let step_content = &TOUR_STEPS[step];
    let is_first = step == 0;
    let is_last = step == 3;

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(440.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Header
                .child(
                    div()
                        .px_5()
                        .py_3()
                        .bg(accent.opacity(0.1))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child("Tour: Named Ranges & Refactoring")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child(format!("Step {} of 4", step + 1))
                        )
                )
                // Content
                .child(
                    div()
                        .px_5()
                        .py_4()
                        .flex()
                        .flex_col()
                        .gap_3()
                        // Title
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_base()
                                .font_weight(FontWeight::SEMIBOLD)
                                .child(step_content.title)
                        )
                        // Body
                        .child(
                            div()
                                .text_color(text_primary.opacity(0.9))
                                .text_sm()
                                .line_height(rems(1.5))
                                .child(step_content.body)
                        )
                        // Hint
                        .child(
                            div()
                                .mt_2()
                                .px_3()
                                .py_2()
                                .bg(panel_border.opacity(0.3))
                                .rounded_md()
                                .text_color(text_muted)
                                .text_xs()
                                .child(step_content.hint)
                        )
                )
                // Footer with buttons
                .child(
                    div()
                        .px_5()
                        .py_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        // Progress dots
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_1()
                                .child(render_progress_dot(step == 0, accent, text_muted))
                                .child(render_progress_dot(step == 1, accent, text_muted))
                                .child(render_progress_dot(step == 2, accent, text_muted))
                                .child(render_progress_dot(step == 3, accent, text_muted))
                        )
                        // Buttons
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                // Back button
                                .child(
                                    div()
                                        .id("tour-back")
                                        .px_3()
                                        .py_1()
                                        .rounded_md()
                                        .text_sm()
                                        .cursor(if is_first { CursorStyle::default() } else { CursorStyle::PointingHand })
                                        .text_color(if is_first { text_muted.opacity(0.5) } else { text_muted })
                                        .when(!is_first, |d| {
                                            d.hover(|s| s.bg(panel_border.opacity(0.5)))
                                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                                    this.tour_back(cx);
                                                }))
                                        })
                                        .child("Back")
                                )
                                // Next/Done button
                                .child(
                                    div()
                                        .id("tour-next")
                                        .px_4()
                                        .py_1()
                                        .rounded_md()
                                        .text_sm()
                                        .cursor_pointer()
                                        .bg(accent)
                                        .text_color(hsla(0.0, 0.0, 1.0, 1.0))
                                        .hover(|s| s.bg(accent.opacity(0.85)))
                                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                            if is_last {
                                                this.tour_done(cx);
                                            } else {
                                                this.tour_next(cx);
                                            }
                                        }))
                                        .child(if is_last { "Done" } else { "Next" })
                                )
                        )
                )
        )
}

fn render_progress_dot(active: bool, accent: Hsla, muted: Hsla) -> impl IntoElement {
    div()
        .w(px(6.0))
        .h(px(6.0))
        .rounded_full()
        .bg(if active { accent } else { muted.opacity(0.3) })
}

/// Render the one-time name tooltip
pub fn render_name_tooltip(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    // Position near top-center
    div()
        .absolute()
        .top(px(60.0))
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .child(
            div()
                .px_4()
                .py_3()
                .bg(panel_bg)
                .border_1()
                .border_color(accent.opacity(0.5))
                .rounded_lg()
                .shadow_lg()
                .flex()
                .items_center()
                .gap_3()
                // Lightbulb icon (using text)
                .child(
                    div()
                        .text_base()
                        .child("💡")
                )
                // Text
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_0()
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_sm()
                                .font_weight(FontWeight::MEDIUM)
                                .child("Tip: Give this range a name")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child("Press Ctrl+Shift+N to create a named range")
                        )
                )
                // Dismiss button
                .child(
                    div()
                        .id("dismiss-tooltip")
                        .ml_2()
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .text_color(text_muted)
                        .text_xs()
                        .cursor_pointer()
                        .hover(|s| s.bg(panel_border.opacity(0.5)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.dismiss_name_tooltip(cx);
                        }))
                        .child("×")
                )
        )
}

/// Render the F2 function key tip (macOS only)
/// Suggests enabling standard function keys so F2 works for edit
pub fn render_f2_tooltip(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    // Position near top-center
    div()
        .absolute()
        .top(px(60.0))
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .child(
            div()
                .px_4()
                .py_3()
                .bg(panel_bg)
                .border_1()
                .border_color(accent.opacity(0.5))
                .rounded_lg()
                .shadow_lg()
                .flex()
                .flex_col()
                .gap_2()
                // Header row with icon and dismiss
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                .child(
                                    div()
                                        .text_base()
                                        .child("⌨️")
                                )
                                .child(
                                    div()
                                        .text_color(text_primary)
                                        .text_sm()
                                        .font_weight(FontWeight::MEDIUM)
                                        .child("Tip: Enable F2 for faster editing")
                                )
                        )
                        // Dismiss X
                        .child(
                            div()
                                .id("dismiss-f2-tip-x")
                                .px_2()
                                .py_1()
                                .rounded_sm()
                                .text_color(text_muted)
                                .text_xs()
                                .cursor_pointer()
                                .hover(|s| s.bg(panel_border.opacity(0.5)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_f2_tip(cx);
                                }))
                                .child("×")
                        )
                )
                // Explanation
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .line_height(rems(1.4))
                        .child("To use F2 like other spreadsheets, enable standard function keys:")
                )
                // Steps
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .line_height(rems(1.4))
                        .child("System Settings → Keyboard → \"Use F1, F2, etc. keys as standard function keys\"")
                )
                // Alternative hint
                .child(
                    div()
                        .text_color(text_muted.opacity(0.8))
                        .text_xs()
                        .italic()
                        .child("Or hold Fn while pressing F2.")
                )
                // Action buttons
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_end()
                        .gap_2()
                        .mt_1()
                        // "Got it" button (hides tip for now)
                        .child(
                            div()
                                .id("f2-tip-got-it")
                                .px_3()
                                .py_1()
                                .rounded_md()
                                .text_xs()
                                .cursor_pointer()
                                .text_color(text_muted)
                                .hover(|s| s.bg(panel_border.opacity(0.5)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_f2_tip(cx);
                                }))
                                .child("Got it")
                        )
                        // "Don't show again" button
                        .child(
                            div()
                                .id("f2-tip-dismiss")
                                .px_3()
                                .py_1()
                                .rounded_md()
                                .text_xs()
                                .cursor_pointer()
                                .bg(accent.opacity(0.15))
                                .text_color(accent)
                                .hover(|s| s.bg(accent.opacity(0.25)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.dismiss_f2_tip(cx);
                                }))
                                .child("Don't show again")
                        )
                )
        )
}
