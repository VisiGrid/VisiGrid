//! Transform Diff Preview dialog (Pro)
//!
//! Shows a table of Cell | Before | After for all cells that will be modified.
//! Capped at MAX_PREVIEW_DISPLAY_ROWS for rendering. Apply commits without
//! recomputation; Cancel clears preview state.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::transforms::{TransformPreview, MAX_PREVIEW_DISPLAY_ROWS};
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize};

/// Render the Transform Diff Preview dialog overlay.
pub fn render_transform_diff_dialog(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);

    let preview = app.transform_preview.as_ref().unwrap();
    let total = preview.affected_count();
    let display_count = total.min(MAX_PREVIEW_DISPLAY_ROWS);
    let overflow = total.saturating_sub(MAX_PREVIEW_DISPLAY_ROWS);
    let op_label = preview.op.label();

    // Header
    let header = div()
        .flex()
        .flex_col()
        .gap_1()
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child(format!("Transform Preview: {}", op_label))
        )
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child(if total == 1 {
                    "1 cell will be modified".to_string()
                } else {
                    format!("{} cells will be modified", total)
                })
        );

    // Table header row
    let table_header = div()
        .flex()
        .gap_1()
        .pb_1()
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .w(px(60.0))
                .flex_shrink_0()
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_muted)
                .child("Cell")
        )
        .child(
            div()
                .flex_1()
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_muted)
                .child("Before")
        )
        .child(
            div()
                .flex_1()
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_muted)
                .child("After")
        );

    // Table rows (capped at MAX_PREVIEW_DISPLAY_ROWS)
    let rows = preview.diffs.iter()
        .take(display_count)
        .enumerate()
        .map(|(i, diff)| {
            let cell_label = cell_ref(diff.row, diff.col);
            let stripe = if i % 2 == 1 {
                panel_border.opacity(0.15)
            } else {
                hsla(0.0, 0.0, 0.0, 0.0)
            };

            div()
                .flex()
                .gap_1()
                .py(px(2.0))
                .px(px(2.0))
                .rounded(px(2.0))
                .bg(stripe)
                .child(
                    div()
                        .w(px(60.0))
                        .flex_shrink_0()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(accent)
                        .child(cell_label)
                )
                .child(
                    div()
                        .flex_1()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .overflow_x_hidden()
                        .child(truncate_display(&diff.before, 40))
                )
                .child(
                    div()
                        .flex_1()
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .overflow_x_hidden()
                        .child(truncate_display(&diff.after, 40))
                )
        })
        .collect::<Vec<_>>();

    // Body: table header + rows + overflow message
    let mut body = div()
        .flex()
        .flex_col()
        .gap(px(2.0))
        .child(table_header)
        .children(rows);

    if overflow > 0 {
        body = body.child(
            div()
                .pt_2()
                .text_size(px(11.0))
                .text_color(text_muted)
                .italic()
                .child(format!("+ {} more change{}", overflow, if overflow == 1 { "" } else { "s" }))
        );
    }

    // Footer with Cancel + Apply buttons
    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(
            Button::new("transform-preview-cancel", "Cancel")
                .secondary(panel_border, text_muted)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.cancel_transform_preview(cx);
                }))
        )
        .child(
            Button::new("transform-preview-apply", "Apply")
                .primary(accent, text_inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.confirm_transform_preview(cx);
                }))
        );

    // Wrap in key handler
    let dialog_content = div()
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
            match event.keystroke.key.as_str() {
                "escape" => this.cancel_transform_preview(cx),
                "enter" => this.confirm_transform_preview(cx),
                _ => {}
            }
        }))
        .child(
            DialogFrame::new(body, panel_bg, panel_border)
                .header(header)
                .footer(footer)
                .width(px(520.0))
                .max_height(px(480.0))
        );

    modal_overlay(
        "transform-diff-dialog",
        |this: &mut Spreadsheet, cx: &mut Context<Spreadsheet>| {
            this.cancel_transform_preview(cx);
        },
        dialog_content,
        cx,
    )
}

// ============================================================================
// Helpers
// ============================================================================

fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 { break; }
        n = n / 26 - 1;
    }
    result
}

fn cell_ref(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Truncate a string for display, appending "..." if too long.
fn truncate_display(s: &str, max_chars: usize) -> String {
    if s.chars().count() <= max_chars {
        s.to_string()
    } else {
        let truncated: String = s.chars().take(max_chars - 1).collect();
        format!("{}…", truncated)
    }
}
