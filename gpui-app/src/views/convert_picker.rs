//! Convert format picker dialog (palette → Convert)
//!
//! Provides selection between output formats:
//! - CSV, TSV, JSON, XLSX
//! Inserts the conversion command into the terminal without auto-running.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

/// Format option metadata
struct FormatOption {
    label: &'static str,
    ext: &'static str,
    description: &'static str,
}

const FORMAT_OPTIONS: &[FormatOption] = &[
    FormatOption { label: "CSV", ext: "csv", description: "Comma-separated values" },
    FormatOption { label: "TSV", ext: "tsv", description: "Tab-separated values" },
    FormatOption { label: "JSON", ext: "json", description: "JSON array of objects" },
    FormatOption { label: "XLSX", ext: "xlsx", description: "Excel workbook" },
];

/// Render the Convert format picker dialog overlay
pub fn render_convert_picker(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let selected = app.convert_picker_selected as usize;

    let options = FORMAT_OPTIONS
        .iter()
        .enumerate()
        .map(|(idx, opt)| {
            let is_selected = idx == selected;

            div()
                .id(SharedString::from(format!("convert-fmt-{}", opt.ext)))
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
                    this.convert_picker_selected = idx as u8;
                    cx.notify();
                }))
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_3()
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
                        // Label and description
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap(px(2.0))
                                .child(
                                    div()
                                        .text_size(px(13.0))
                                        .font_weight(if is_selected { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                                        .text_color(text_primary)
                                        .child(opt.label)
                                )
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .child(opt.description)
                                )
                        )
                )
        })
        .collect::<Vec<_>>();

    let body = div()
        .flex()
        .flex_col()
        .gap_1()
        .children(options);

    let header = div()
        .text_size(px(14.0))
        .font_weight(FontWeight::MEDIUM)
        .text_color(text_primary)
        .child("Convert to...");

    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(
            Button::new("convert-cancel", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.cancel_convert_picker(cx);
                }))
        )
        .child(
            Button::new("convert-ok", "Insert Command")
                .primary(accent, text_inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                    this.execute_convert(window, cx);
                }))
        );

    let dialog_content = div()
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            match event.keystroke.key.as_str() {
                "escape" => {
                    this.cancel_convert_picker(cx);
                }
                "enter" => {
                    this.execute_convert(window, cx);
                }
                "up" => {
                    if this.convert_picker_selected > 0 {
                        this.convert_picker_selected -= 1;
                        cx.notify();
                    }
                }
                "down" => {
                    if (this.convert_picker_selected as usize) < FORMAT_OPTIONS.len() - 1 {
                        this.convert_picker_selected += 1;
                        cx.notify();
                    }
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
        "convert-picker-dialog",
        |this, cx| this.cancel_convert_picker(cx),
        dialog_content,
        cx,
    )
}
