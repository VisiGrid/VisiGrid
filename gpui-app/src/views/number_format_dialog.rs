//! Number Format Editor dialog (Ctrl+1 escalation from Format tab)
//!
//! Provides editing of Number and Currency format settings:
//! - Type selection (General, Number, Currency, Percent, Date)
//! - Decimal places (0-10)
//! - Thousands separator toggle
//! - Negative number style (Minus, Parens, RedMinus, RedParens)
//! - Currency symbol (for Currency type)
//! - Live preview of positive, negative, and zero values

use gpui::*;
use gpui::prelude::FluentBuilder;
use visigrid_engine::cell::NegativeStyle;
use crate::app::{Spreadsheet, NumberFormatEditorType};
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

/// Render the Number Format Editor dialog overlay
pub fn render_number_format_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let state = &app.number_format_editor;
    let format_type = state.format_type;
    let show_numeric_controls = matches!(format_type, NumberFormatEditorType::Number | NumberFormatEditorType::Currency | NumberFormatEditorType::Percent);
    let show_negative_styles = matches!(format_type, NumberFormatEditorType::Number | NumberFormatEditorType::Currency);
    let show_thousands = matches!(format_type, NumberFormatEditorType::Number | NumberFormatEditorType::Currency);
    let show_symbol = format_type == NumberFormatEditorType::Currency;

    // --- Type pills ---
    let type_pills = div()
        .flex()
        .flex_wrap()
        .gap_1()
        .child(type_pill("General", NumberFormatEditorType::General, format_type, text_primary, text_muted, accent, panel_border, cx))
        .child(type_pill("Number", NumberFormatEditorType::Number, format_type, text_primary, text_muted, accent, panel_border, cx))
        .child(type_pill("Currency", NumberFormatEditorType::Currency, format_type, text_primary, text_muted, accent, panel_border, cx))
        .child(type_pill("Percent", NumberFormatEditorType::Percent, format_type, text_primary, text_muted, accent, panel_border, cx))
        .child(type_pill("Date", NumberFormatEditorType::Date, format_type, text_primary, text_muted, accent, panel_border, cx));

    // --- Currency symbol presets ---
    let current_symbol = if state.currency_symbol.is_empty() { "$".to_string() } else { state.currency_symbol.clone() };
    let symbol_row = div()
        .flex()
        .items_center()
        .gap_2()
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .w(px(90.0))
                .child("Symbol")
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_1()
                .child(symbol_preset_btn("$", &current_symbol, text_primary, accent, panel_border, cx))
                .child(symbol_preset_btn("\u{20AC}", &current_symbol, text_primary, accent, panel_border, cx)) // €
                .child(symbol_preset_btn("\u{00A3}", &current_symbol, text_primary, accent, panel_border, cx)) // £
                .child(symbol_preset_btn("\u{00A5}", &current_symbol, text_primary, accent, panel_border, cx)) // ¥
                .child(
                    div()
                        .id("nf-symbol-current")
                        .px_2()
                        .py(px(3.0))
                        .rounded(px(3.0))
                        .border_1()
                        .border_color(panel_border)
                        .bg(panel_bg)
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .min_w(px(30.0))
                        .flex()
                        .justify_center()
                        .child(current_symbol)
                )
        );

    // --- Decimals control ---
    let decimals = state.decimals;
    let decimals_row = div()
        .flex()
        .items_center()
        .gap_2()
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .w(px(90.0))
                .child("Decimals")
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_1()
                .child(
                    small_button("-", decimals > 0, text_primary, text_muted, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            if this.number_format_editor.decimals > 0 {
                                this.number_format_editor.decimals -= 1;
                                cx.notify();
                            }
                        }))
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .w(px(24.0))
                        .flex()
                        .justify_center()
                        .child(decimals.to_string())
                )
                .child(
                    small_button("+", decimals < 10, text_primary, text_muted, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            if this.number_format_editor.decimals < 10 {
                                this.number_format_editor.decimals += 1;
                                cx.notify();
                            }
                        }))
                )
        );

    // --- Thousands toggle ---
    let thousands = state.thousands;
    let thousands_row = div()
        .flex()
        .items_center()
        .gap_2()
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .w(px(90.0))
                .child("Thousands")
        )
        .child(
            checkbox("nf-thousands", thousands, accent, panel_border, text_primary)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.number_format_editor.thousands = !this.number_format_editor.thousands;
                    cx.notify();
                }))
        );

    // --- Negative style radio buttons ---
    let negative = state.negative;
    // Build preview snippets for each negative style using current decimals/thousands/symbol
    let neg_previews = build_negative_previews(state);
    let negative_section = div()
        .flex()
        .flex_col()
        .gap(px(2.0))
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .mb(px(2.0))
                .child("Negative style")
        )
        .child(negative_radio(NegativeStyle::Minus, &neg_previews[0], negative, text_primary, text_muted, accent, panel_border, cx))
        .child(negative_radio(NegativeStyle::Parens, &neg_previews[1], negative, text_primary, text_muted, accent, panel_border, cx))
        .child(negative_radio(NegativeStyle::RedMinus, &neg_previews[2], negative, text_primary, text_muted, accent, panel_border, cx))
        .child(negative_radio(NegativeStyle::RedParens, &neg_previews[3], negative, text_primary, text_muted, accent, panel_border, cx));

    // --- Preview section ---
    let preview_pos = state.preview();
    let preview_neg = state.preview_negative();
    let preview_zero = state.preview_zero();
    let is_red = state.negative.is_red();

    let preview_section = div()
        .flex()
        .flex_col()
        .gap(px(2.0))
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Preview")
        )
        .child(
            div()
                .flex()
                .gap_3()
                .px_3()
                .py_2()
                .rounded(px(4.0))
                .border_1()
                .border_color(panel_border)
                .bg(panel_bg)
                .text_size(px(12.0))
                .child(
                    div().text_color(text_primary).child(preview_pos)
                )
                .child(
                    div()
                        .text_color(if is_red { gpui::Hsla::from(gpui::rgb(0xCC0000)) } else { text_primary })
                        .child(preview_neg)
                )
                .child(
                    div().text_color(text_muted).child(preview_zero)
                )
        );

    // --- Body assembly ---
    let body = div()
        .flex()
        .flex_col()
        .gap_2()
        .child(type_pills)
        .child(div().h(px(1.0)).bg(panel_border).w_full())
        .when(show_symbol, |el| el.child(symbol_row))
        .when(show_numeric_controls, |el| el.child(decimals_row))
        .when(show_thousands, |el| el.child(thousands_row))
        .when(show_negative_styles, |el| el.child(negative_section))
        .child(div().h(px(1.0)).bg(panel_border).w_full())
        .child(preview_section);

    // Header — show "Based on active cell" hint for multi-cell selections
    let is_multi_cell = app.all_selection_ranges().iter().any(|((r1, c1), (r2, c2))| r1 != r2 || c1 != c2)
        || app.all_selection_ranges().len() > 1;
    let header = div()
        .flex()
        .flex_col()
        .gap(px(2.0))
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child("Number Format")
        )
        .when(is_multi_cell, |d| {
            d.child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted)
                    .child("Based on active cell")
            )
        });

    // Footer
    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(
            Button::new("nf-cancel", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.close_number_format_editor(cx);
                }))
        )
        .child(
            Button::new("nf-apply", "Apply")
                .primary(accent, text_inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.apply_number_format_editor(cx);
                }))
        );

    // Keyboard handler
    let dialog_content = div()
        .on_key_down(cx.listener(move |this, event: &KeyDownEvent, _, cx| {
            match event.keystroke.key.as_str() {
                "escape" => {
                    this.close_number_format_editor(cx);
                }
                "enter" => {
                    this.apply_number_format_editor(cx);
                }
                // Arrow up/down: cycle negative styles
                "up" if show_negative_styles => {
                    let cur = this.number_format_editor.negative.to_int();
                    this.number_format_editor.negative = NegativeStyle::from_int((cur + 3) % 4);
                    cx.notify();
                }
                "down" if show_negative_styles => {
                    let cur = this.number_format_editor.negative.to_int();
                    this.number_format_editor.negative = NegativeStyle::from_int((cur + 1) % 4);
                    cx.notify();
                }
                // +/- or =/- keys: adjust decimals
                "+" | "=" if show_numeric_controls => {
                    if this.number_format_editor.decimals < 10 {
                        this.number_format_editor.decimals += 1;
                        cx.notify();
                    }
                }
                "-" if show_numeric_controls => {
                    if this.number_format_editor.decimals > 0 {
                        this.number_format_editor.decimals -= 1;
                        cx.notify();
                    }
                }
                // Space: toggle thousands separator
                "space" if show_thousands => {
                    this.number_format_editor.thousands = !this.number_format_editor.thousands;
                    cx.notify();
                }
                // Tab: cycle type forward, Shift+Tab: cycle backward
                "tab" => {
                    use NumberFormatEditorType::*;
                    let types = [General, Number, Currency, Percent, Date];
                    let cur = types.iter().position(|t| *t == this.number_format_editor.format_type).unwrap_or(0);
                    let next = if event.keystroke.modifiers.shift {
                        (cur + types.len() - 1) % types.len()
                    } else {
                        (cur + 1) % types.len()
                    };
                    this.number_format_editor.switch_type(types[next]);
                    cx.notify();
                }
                _ => {}
            }
        }))
        .child(
            DialogFrame::new(body, panel_bg, panel_border)
                .size(DialogSize::Lg)
                .header(header)
                .footer(footer)
        );

    modal_overlay(
        "number-format-dialog",
        |this, cx| this.close_number_format_editor(cx),
        dialog_content,
        cx,
    )
}

// --- Helper components ---

fn type_pill(
    label: &str,
    pill_type: NumberFormatEditorType,
    current: NumberFormatEditorType,
    text_primary: Hsla,
    _text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let is_active = pill_type == current;
    let label_owned = SharedString::from(label.to_string());
    div()
        .id(SharedString::from(format!("nf-type-{}", label)))
        .px_2()
        .py(px(3.0))
        .rounded(px(4.0))
        .cursor_pointer()
        .text_size(px(11.0))
        .when(is_active, |d| {
            d.bg(accent.opacity(0.15))
                .text_color(accent)
                .font_weight(FontWeight::MEDIUM)
        })
        .when(!is_active, |d| {
            d.text_color(text_primary)
                .hover(|s| s.bg(panel_border.opacity(0.5)))
        })
        .child(label_owned)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.number_format_editor.switch_type(pill_type);
            cx.notify();
        }))
}

fn small_button(
    label: &str,
    enabled: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let label_owned = SharedString::from(label.to_string());
    div()
        .id(SharedString::from(format!("nf-btn-{}", label)))
        .w(px(22.0))
        .h(px(22.0))
        .rounded(px(3.0))
        .border_1()
        .border_color(panel_border)
        .flex()
        .items_center()
        .justify_center()
        .text_size(px(12.0))
        .text_color(if enabled { text_primary } else { text_muted })
        .cursor(if enabled { CursorStyle::PointingHand } else { CursorStyle::Arrow })
        .when(enabled, |d| d.hover(|s| s.bg(panel_border.opacity(0.5))))
        .child(label_owned)
}

fn checkbox(
    id: &str,
    checked: bool,
    accent: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
) -> Stateful<Div> {
    div()
        .id(SharedString::from(id.to_string()))
        .w(px(16.0))
        .h(px(16.0))
        .rounded(px(3.0))
        .border_1()
        .border_color(if checked { accent } else { panel_border })
        .when(checked, |d| d.bg(accent.opacity(0.15)))
        .flex()
        .items_center()
        .justify_center()
        .cursor_pointer()
        .text_size(px(10.0))
        .text_color(text_primary)
        .when(checked, |d| d.child("\u{2713}")) // ✓
}

fn negative_radio(
    style: NegativeStyle,
    preview: &str,
    current: NegativeStyle,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let is_selected = style == current;
    let is_red = style.is_red();
    let label_color = if is_red { gpui::Hsla::from(gpui::rgb(0xCC0000)) } else { text_primary };
    let suffix = if is_red { "  (red)" } else { "" };

    div()
        .id(SharedString::from(format!("nf-neg-{}", style.to_int())))
        .flex()
        .items_center()
        .gap_2()
        .px_2()
        .py(px(3.0))
        .rounded(px(3.0))
        .cursor_pointer()
        .when(is_selected, |d| d.bg(accent.opacity(0.08)))
        .when(!is_selected, |d| d.hover(|s| s.bg(panel_border.opacity(0.3))))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.number_format_editor.negative = style;
            cx.notify();
        }))
        // Radio indicator
        .child(
            div()
                .w(px(14.0))
                .h(px(14.0))
                .rounded_full()
                .border_1()
                .border_color(if is_selected { accent } else { text_muted })
                .flex()
                .items_center()
                .justify_center()
                .when(is_selected, |d| {
                    d.child(
                        div()
                            .w(px(8.0))
                            .h(px(8.0))
                            .rounded_full()
                            .bg(accent)
                    )
                })
        )
        // Label
        .child(
            div()
                .flex()
                .items_center()
                .gap_1()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(label_color)
                        .child(preview.to_string())
                )
                .when(!suffix.is_empty(), |d| {
                    d.child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child(suffix)
                    )
                })
        )
}

fn symbol_preset_btn(
    symbol: &str,
    current: &str,
    text_primary: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let is_active = symbol == current;
    let sym = symbol.to_string();
    div()
        .id(SharedString::from(format!("nf-sym-{}", symbol)))
        .w(px(28.0))
        .h(px(24.0))
        .rounded(px(3.0))
        .border_1()
        .border_color(if is_active { accent } else { panel_border })
        .when(is_active, |d| d.bg(accent.opacity(0.15)))
        .flex()
        .items_center()
        .justify_center()
        .text_size(px(12.0))
        .text_color(text_primary)
        .cursor_pointer()
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .child(SharedString::from(symbol.to_string()))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // "$" is stored as empty string (default)
            this.number_format_editor.currency_symbol = if sym == "$" {
                String::new()
            } else {
                sym.clone()
            };
            cx.notify();
        }))
}

/// Build preview snippets for all 4 negative styles using current editor state
fn build_negative_previews(state: &crate::app::NumberFormatEditorState) -> [String; 4] {
    use visigrid_engine::cell::CellValue;
    use visigrid_engine::cell::NumberFormat;

    let sample = state.preview_value.abs().max(1234.56);
    let styles = [
        NegativeStyle::Minus,
        NegativeStyle::Parens,
        NegativeStyle::RedMinus,
        NegativeStyle::RedParens,
    ];

    styles.map(|neg| {
        let fmt = match state.format_type {
            NumberFormatEditorType::Number => NumberFormat::Number {
                decimals: state.decimals,
                thousands: state.thousands,
                negative: neg,
            },
            NumberFormatEditorType::Currency => NumberFormat::Currency {
                decimals: state.decimals,
                thousands: state.thousands,
                negative: neg,
                symbol: if state.currency_symbol.is_empty() { None } else { Some(state.currency_symbol.clone()) },
            },
            _ => NumberFormat::Number {
                decimals: state.decimals,
                thousands: state.thousands,
                negative: neg,
            },
        };
        CellValue::format_number(-sample, &fmt)
    })
}
