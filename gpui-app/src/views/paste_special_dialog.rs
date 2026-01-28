//! Paste Special dialog (Ctrl+Alt+V)
//!
//! Provides selection between paste modes:
//! - All (normal paste)
//! - Values (computed values only)
//! - Formulas (with reference adjustment)
//! - Formats (cell formatting only)

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, PasteType};
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

/// Render the Paste Special dialog overlay
pub fn render_paste_special_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let selected = app.paste_special_dialog.selected;

    // Build the list of paste type options
    let options = PasteType::all()
        .iter()
        .map(|paste_type| {
            let is_selected = *paste_type == selected;
            let pt = *paste_type;

            div()
                .id(SharedString::from(format!("paste-type-{:?}", paste_type)))
                .px_3()
                .py_2()
                .rounded(px(4.0))
                .cursor_pointer()
                .when(is_selected, |d| {
                    d.bg(accent.opacity(0.15))
                        .border_1()
                        .border_color(accent)
                })
                .when(!is_selected, |d| {
                    d.border_1()
                        .border_color(panel_border.opacity(0.0))
                        .hover(|s| s.bg(panel_border.opacity(0.5)))
                })
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.paste_special_dialog.selected = pt;
                    cx.notify();
                }))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                    if event.click_count == 2 {
                        this.apply_paste_special(cx);
                    }
                }))
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_3()
                        // Radio button indicator
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
                        // Label and description
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap(px(2.0))
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .gap_1()
                                        .child(
                                            div()
                                                .text_size(px(13.0))
                                                .font_weight(if is_selected { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                                                .text_color(text_primary)
                                                .child(paste_type.label())
                                        )
                                        // Accelerator hint
                                        .child(
                                            div()
                                                .text_size(px(11.0))
                                                .text_color(text_muted)
                                                .child(format!("({})", paste_type.accelerator()))
                                        )
                                )
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .child(paste_type.description())
                                )
                        )
                )
        })
        .collect::<Vec<_>>();

    // Body content
    let body = div()
        .flex()
        .flex_col()
        .gap_1()
        .children(options);

    // Header
    let header = div()
        .text_size(px(14.0))
        .font_weight(FontWeight::MEDIUM)
        .text_color(text_primary)
        .child("Paste Special");

    // Footer with buttons
    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(
            Button::new("paste-special-cancel", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.hide_paste_special(cx);
                }))
        )
        .child(
            Button::new("paste-special-ok", "Paste")
                .primary(accent, text_inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.apply_paste_special(cx);
                }))
        );

    // Wrap in keyboard handler for arrow/enter/escape and letter accelerators
    let dialog_content = div()
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
            match event.keystroke.key.as_str() {
                "escape" => {
                    this.hide_paste_special(cx);
                }
                "enter" => {
                    this.apply_paste_special(cx);
                }
                "up" => {
                    this.paste_special_up(cx);
                }
                "down" => {
                    this.paste_special_down(cx);
                }
                // Letter accelerators (Excel-style, work with or without Alt)
                "a" => {
                    this.paste_special_dialog.selected = PasteType::All;
                    this.apply_paste_special(cx);
                }
                "v" => {
                    this.paste_special_dialog.selected = PasteType::Values;
                    this.apply_paste_special(cx);
                }
                "f" => {
                    this.paste_special_dialog.selected = PasteType::Formulas;
                    this.apply_paste_special(cx);
                }
                "o" => {
                    // fOrmats - Excel convention for Formats
                    this.paste_special_dialog.selected = PasteType::Formats;
                    this.apply_paste_special(cx);
                }
                _ => {}
            }
        }))
        .child(
            DialogFrame::new(body, panel_bg, panel_border)
                .size(DialogSize::Md)
                .header(header)
                .footer(footer)
        );

    modal_overlay(
        "paste-special-dialog",
        |this, cx| this.hide_paste_special(cx),
        dialog_content,
        cx,
    )
}
