//! Conditional Formatting rules panel — right-side drawer.
//!
//! One rule per line in the typed syntax, no modals: toggle with the
//! checkbox, reorder with ↑/↓ (precedence: later rules win per property),
//! Edit reopens the quick-add dialog pre-filled (with live preview),
//! ✕ deletes. Every mutation is a single undo step.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::cond_format_ui::{format_range_label, style_to_text};
use crate::theme::TokenKey;

const PANEL_WIDTH: f32 = 380.0;

pub(crate) fn render_cf_rules_panel(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);
    let editor_bg = app.token(TokenKey::EditorBg);

    let rules: Vec<_> = app.sheet(cx).cond_formats.iter().cloned().collect();
    let rule_count = rules.len();

    div()
        .id("cf-rules-panel")
        .absolute()
        .right_0()
        .top_0()
        .h_full()
        .w(px(PANEL_WIDTH))
        .bg(panel_bg)
        .border_l_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_mouse_up(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        // Header
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .px_3()
                .py_2()
                .border_b_1()
                .border_color(panel_border)
                .child(
                    div()
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(format!(
                            "Conditional Formatting ({} rule{})",
                            rule_count,
                            if rule_count == 1 { "" } else { "s" }
                        ))
                )
                .child(
                    div()
                        .id("cf-panel-close")
                        .px_2()
                        .cursor_pointer()
                        .text_color(text_muted)
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            cx.stop_propagation();
                            this.cf_panel_visible = false;
                            cx.notify();
                        }))
                        .child("\u{2715}")
                )
        )
        // Add rule for current selection
        .child(
            div()
                .id("cf-panel-add")
                .mx_3()
                .mt_2()
                .mb_1()
                .px_2()
                .py_1()
                .rounded_sm()
                .border_1()
                .border_color(accent.opacity(0.45))
                .bg(editor_bg)
                .cursor_pointer()
                .text_sm()
                .text_color(text_muted)
                .hover(|s| s.border_color(gpui::hsla(0.6, 0.6, 0.6, 0.8)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.show_add_cond_format(cx);
                }))
                .child("+ Add rule for selection\u{2026}")
        )
        // Precedence note
        .when(rule_count > 1, |d| {
            d.child(
                div()
                    .px_3()
                    .py_1()
                    .text_xs()
                    .text_color(text_muted.opacity(0.7))
                    .child("Later rules win when properties conflict")
            )
        })
        // Rules list
        .child(
            div()
                .flex_1()
                .overflow_hidden()
                .flex()
                .flex_col()
                .px_2()
                .py_1()
                .gap_1()
                .when(rule_count == 0, |d| {
                    d.child(
                        div()
                            .px_2()
                            .py_3()
                            .text_sm()
                            .text_color(text_muted.opacity(0.6))
                            .child("No rules on this sheet. Select a range and add one — e.g. =A1>100 -> bad")
                    )
                })
                .children(rules.iter().enumerate().map(|(idx, rule)| {
                    let id = rule.id;
                    let enabled = rule.enabled;
                    let is_first = idx == 0;
                    let is_last = idx + 1 == rule_count;
                    let range_label = format_range_label(&rule.ranges);
                    let rule_text = format!("{} \u{2192} {}", rule.predicate, style_to_text(&rule.style));
                    let parse_err = rule.parse_error().map(|e| e.to_string());

                    div()
                        .id(ElementId::Name(format!("cf-rule-{}", id).into()))
                        .rounded_sm()
                        .border_1()
                        .border_color(panel_border.opacity(0.6))
                        .bg(editor_bg.opacity(if enabled { 1.0 } else { 0.5 }))
                        .px_2()
                        .py_1()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        // Row 1: checkbox, range, action buttons
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                .child(
                                    // Enable/disable checkbox
                                    div()
                                        .id(ElementId::Name(format!("cf-toggle-{}", id).into()))
                                        .w(px(14.0))
                                        .h(px(14.0))
                                        .rounded_sm()
                                        .border_1()
                                        .border_color(if enabled { accent } else { text_muted })
                                        .when(enabled, |d| d.bg(accent.opacity(0.8)))
                                        .cursor_pointer()
                                        .flex()
                                        .items_center()
                                        .justify_center()
                                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                            cx.stop_propagation();
                                            this.toggle_cf_rule(id, cx);
                                        }))
                                        .when(enabled, |d| {
                                            d.child(div().text_xs().text_color(gpui::white()).child("\u{2713}"))
                                        })
                                )
                                .child(
                                    div()
                                        .text_xs()
                                        .text_color(text_muted)
                                        .child(range_label)
                                )
                                .child(div().flex_1())
                                .child(rule_button(format!("cf-up-{}", id), "\u{2191}", !is_first, text_muted, text_primary, cx,
                                    move |this, cx| this.move_cf_rule(id, -1, cx)))
                                .child(rule_button(format!("cf-down-{}", id), "\u{2193}", !is_last, text_muted, text_primary, cx,
                                    move |this, cx| this.move_cf_rule(id, 1, cx)))
                                .child(rule_button(format!("cf-edit-{}", id), "Edit", true, text_muted, text_primary, cx,
                                    move |this, cx| this.edit_cf_rule(id, cx)))
                                .child(rule_button(format!("cf-del-{}", id), "\u{2715}", true, text_muted, error_color, cx,
                                    move |this, cx| this.delete_cf_rule(id, cx)))
                        )
                        // Row 2: the rule itself, typed syntax
                        .child(
                            div()
                                .text_sm()
                                .font_family(crate::views::terminal_panel::TERM_FONT_FAMILY)
                                .text_color(if enabled { text_primary } else { text_muted })
                                .child(rule_text)
                        )
                        // Parse error badge (rule is inert)
                        .when_some(parse_err, |d, err| {
                            d.child(
                                div()
                                    .text_xs()
                                    .text_color(error_color)
                                    .child(format!("\u{26A0} {}", err))
                            )
                        })
                }))
        )
        // Footer hint
        .child(
            div()
                .px_3()
                .py_2()
                .border_t_1()
                .border_color(panel_border)
                .text_xs()
                .text_color(text_muted.opacity(0.7))
                .child("Rules apply live \u{00B7} every change is one undo step")
        )
}

fn rule_button(
    id: String,
    label: &'static str,
    enabled: bool,
    color: Hsla,
    hover_color: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> Stateful<Div> {
    div()
        .id(ElementId::Name(id.into()))
        .px_1()
        .text_xs()
        .text_color(if enabled { color } else { color.opacity(0.3) })
        .when(enabled, |d| {
            d.cursor_pointer()
                .hover(move |s| s.text_color(hover_color))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation();
                    on_click(this, cx);
                }))
        })
        .child(label)
}
