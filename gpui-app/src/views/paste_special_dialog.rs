//! Paste Special dialog (Ctrl+Alt+V)
//!
//! Provides selection between paste modes:
//! - All (contents and formatting, as copied)
//! - Values (computed values only)
//! - Formulas (with reference adjustment)
//! - Formats (cell formatting only)
//!
//! Keys are routed by the app, not this view: letters in `key_handler`,
//! Enter/Escape/Up/Down through their actions.

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

    let dialog = &app.paste_special_dialog;
    let selected = dialog.selected;

    // Where the paste lands: the active cell, extended by the clipboard's size
    let (view_row, col) = app.view_state.selected;
    let row = app.row_view.view_to_data(view_row);
    let start = app.cell_ref_at(row, col);
    let target = if dialog.rows * dialog.cols > 1 {
        format!("{}:{}", start, app.cell_ref_at(row + dialog.rows - 1, col + dialog.cols - 1))
    } else {
        start
    };
    let cells = if dialog.rows * dialog.cols == 1 {
        "1 cell".to_string()
    } else {
        format!("{} \u{00d7} {} cells", dialog.rows, dialog.cols)
    };
    let source = if dialog.from_visigrid { "" } else { " of text" };
    let summary = format!("{}{} \u{2192} {}", cells, source, target);

    let options = PasteType::all()
        .iter()
        .map(|paste_type| {
            let pt = *paste_type;
            let is_selected = pt == selected;
            let enabled = dialog.is_enabled(pt);
            let description = match pt {
                PasteType::All if !dialog.from_visigrid => "Contents only: text from another app has no formatting",
                PasteType::Formats if !enabled => "Needs cells copied in VisiGrid",
                _ => pt.description(),
            };
            let label_color = if enabled { text_primary } else { text_muted.opacity(0.6) };

            div()
                .id(SharedString::from(format!("paste-type-{:?}", pt)))
                .px_3()
                .py(px(7.0))
                .rounded(px(6.0))
                .border_1()
                .flex()
                .items_center()
                .gap_3()
                .when(is_selected, |d| d.bg(accent.opacity(0.12)).border_color(accent.opacity(0.6)))
                .when(!is_selected, |d| d.border_color(panel_border.opacity(0.0)))
                .when(enabled && !is_selected, |d| d.cursor_pointer().hover(|s| s.bg(panel_border.opacity(0.4))))
                .when(enabled, |d| {
                    d.on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                        this.paste_special_dialog.selected = pt;
                        if event.click_count == 2 {
                            this.apply_paste_special(cx);
                        } else {
                            cx.notify();
                        }
                    }))
                })
                // Radio indicator
                .child(
                    div()
                        .flex_none()
                        .size(px(14.0))
                        .rounded_full()
                        .border_1()
                        .border_color(if is_selected { accent } else { text_muted.opacity(if enabled { 1.0 } else { 0.4 }) })
                        .flex()
                        .items_center()
                        .justify_center()
                        .when(is_selected, |d| d.child(div().size(px(8.0)).rounded_full().bg(accent)))
                )
                // Label and description
                .child(
                    div()
                        .flex_1()
                        .flex()
                        .flex_col()
                        .gap(px(1.0))
                        .child(
                            div()
                                .text_size(px(13.0))
                                .font_weight(if is_selected { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                                .text_color(label_color)
                                .child(pt.label())
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted.opacity(if enabled { 1.0 } else { 0.6 }))
                                .child(description)
                        )
                )
                // Accelerator keycap
                .child(
                    div()
                        .flex_none()
                        .min_w(px(20.0))
                        .px(px(5.0))
                        .py(px(1.0))
                        .rounded(px(4.0))
                        .border_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_center()
                        .text_size(px(11.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(text_muted.opacity(if enabled { 1.0 } else { 0.4 }))
                        .child(pt.accelerator().to_string())
                )
        })
        .collect::<Vec<_>>();

    let body = div()
        .flex()
        .flex_col()
        .gap_1()
        .children(options)
        .child(
            div()
                .mt_2()
                .px_1()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Ctrl+V keeps formatted cells as they are; blank cells take the copied formatting.")
        );

    let header = div()
        .flex()
        .items_center()
        .justify_between()
        .gap_4()
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child("Paste Special")
        )
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child(summary)
        );

    let footer = div()
        .flex()
        .items_center()
        .justify_between()
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("\u{2191}\u{2193} choose \u{00b7} Enter paste \u{00b7} Esc cancel")
        )
        .child(
            div()
                .flex()
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
                )
        );

    let dialog_content = DialogFrame::new(body, panel_bg, panel_border)
        .size(DialogSize::Lg)
        .header(header)
        .footer(footer);

    modal_overlay(
        "paste-special-dialog",
        |this, cx| this.hide_paste_special(cx),
        dialog_content,
        cx,
    )
}
