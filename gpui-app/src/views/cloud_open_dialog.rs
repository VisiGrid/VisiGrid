//! Cloud sheet picker (File > Open Cloud)
//!
//! Lists the signed-in user's cloud sheets, most recently edited first, and
//! opens the chosen one. Keyboard: Up/Down to move, Enter to open, Escape to
//! cancel (handled in key_handler.rs and the edit/nav actions).

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

/// Render the cloud sheet picker overlay
pub fn render_cloud_open_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let selected = app.cloud_selected_sheet;

    let rows = app.cloud_sheets_list
        .iter()
        .enumerate()
        .map(|(idx, sheet)| {
            let is_selected = selected == Some(idx);
            let detail = sheet_detail(sheet.last_edited_at.as_deref(), sheet.byte_size);

            div()
                .id(SharedString::from(format!("cloud-sheet-{}", sheet.id)))
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
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                    this.cloud_selected_sheet = Some(idx);
                    if event.click_count >= 2 {
                        this.cloud_open_selected(cx);
                    }
                    cx.notify();
                }))
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
                                .child(sheet.name.clone())
                        )
                        .when(!detail.is_empty(), |d| {
                            d.child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(text_muted)
                                    .child(detail)
                            )
                        })
                )
        })
        .collect::<Vec<_>>();

    let body = if rows.is_empty() {
        div()
            .id("cloud-sheet-list")
            .py_4()
            .text_size(px(13.0))
            .text_color(text_muted)
            .child("No cloud sheets yet. Use File > Move to Cloud to add this workbook, or create one at app.visigrid.app.")
    } else {
        div()
            .id("cloud-sheet-list")
            .flex()
            .flex_col()
            .gap_1()
            .max_h(px(360.0))
            .overflow_y_scroll()
            .children(rows)
    };

    let header = div()
        .text_size(px(14.0))
        .font_weight(FontWeight::MEDIUM)
        .text_color(text_primary)
        .child("Open Cloud Sheet");

    let has_selection = selected.is_some();
    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(
            Button::new("cloud-open-cancel", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.cloud_picker_cancel(cx);
                }))
        )
        .when(has_selection, |d| {
            d.child(
                Button::new("cloud-open-ok", "Open")
                    .primary(accent, text_inverse)
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        this.cloud_open_selected(cx);
                    }))
            )
        });

    let dialog_content = div()
        .child(
            DialogFrame::new(body, panel_bg, panel_border)
                .size(DialogSize::Md)
                .header(header)
                .footer(footer)
        );

    modal_overlay(
        "cloud-open-dialog",
        |this, cx| this.cloud_picker_cancel(cx),
        dialog_content,
        cx,
    )
}

/// "Edited 2026-09-27 14:05 · 12 KB" from the server's ISO timestamp and size.
fn sheet_detail(last_edited_at: Option<&str>, byte_size: Option<i64>) -> String {
    let mut parts = Vec::new();
    if let Some(ts) = last_edited_at {
        // "2026-09-27T14:05:33.123Z" → "2026-09-27 14:05"
        let short: String = ts.chars().take(16).collect::<String>().replacen('T', " ", 1);
        if !short.is_empty() {
            parts.push(format!("Edited {}", short));
        }
    }
    if let Some(bytes) = byte_size.filter(|b| *b > 0) {
        parts.push(format_size(bytes));
    }
    parts.join(" · ")
}

fn format_size(bytes: i64) -> String {
    if bytes < 1024 {
        format!("{} B", bytes)
    } else if bytes < 1024 * 1024 {
        format!("{} KB", bytes / 1024)
    } else {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    }
}

#[cfg(test)]
mod tests {
    // Not `super::*`: that glob brings in gpui's own `test` attribute.
    use super::sheet_detail;

    #[test]
    fn detail_shows_edit_time_and_size() {
        assert_eq!(
            sheet_detail(Some("2026-09-27T14:05:33.123Z"), Some(12 * 1024)),
            "Edited 2026-09-27 14:05 · 12 KB"
        );
    }

    #[test]
    fn detail_omits_what_the_server_did_not_send() {
        assert_eq!(sheet_detail(None, None), "");
        assert_eq!(sheet_detail(None, Some(0)), "");
        assert_eq!(sheet_detail(Some("2026-09-27T14:05:33Z"), None), "Edited 2026-09-27 14:05");
    }
}
