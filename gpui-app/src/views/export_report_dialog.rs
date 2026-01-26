//! Export Report dialog for Excel exports
//!
//! Shows detailed statistics and warnings after exporting to xlsx.
//! Displayed automatically when export has warnings (formula conversions, precision limits).

use gpui::*;
use gpui::prelude::FluentBuilder;
use visigrid_io::xlsx::ExportResult;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button, DialogFrame, DialogSize, dialog_header_with_subtitle};

/// Render the Export Report dialog overlay
pub fn render_export_report_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let Some(export_result) = &app.export_result else {
        return div().into_any_element();
    };

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Warn);
    let text_inverse = app.token(TokenKey::TextInverse);

    let filename = app.export_filename.as_deref().unwrap_or("unknown file").to_string();

    // Body content
    let body = div()
        .child(render_summary_section(export_result, text_primary, text_muted))
        .child(render_formula_conversions(export_result, text_primary, text_muted, warning_color))
        .child(render_precision_warnings(export_result, text_primary, text_muted, warning_color));

    // Footer with buttons
    let copy_button = if export_result.has_warnings() {
        let report = export_result.full_report_with_context(&filename);
        Button::new("export-report-copy-btn", "Copy Details")
            .secondary(panel_border, text_muted)
            .on_mouse_down(MouseButton::Left, move |_, _, cx| {
                cx.write_to_clipboard(ClipboardItem::new_string(report.clone()));
                cx.stop_propagation();
            })
            .into_any_element()
    } else {
        div().into_any_element()
    };

    let close_button = Button::new("export-report-close-btn", "Close")
        .primary(accent, text_inverse)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_export_report(cx);
        }));

    let footer = div()
        .flex()
        .justify_between()
        .child(copy_button)
        .child(close_button);

    modal_overlay(
        "export-report-dialog",
        |this, cx| this.hide_export_report(cx),
        DialogFrame::new(body, panel_bg, panel_border)
            .size(DialogSize::Lg)
            .max_height(px(500.0))
            .header(dialog_header_with_subtitle("Export Report", filename, text_primary, text_muted))
            .footer(footer),
        cx,
    ).into_any_element()
}

fn render_summary_section(er: &ExportResult, text_primary: Hsla, text_muted: Hsla) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child("Summary")
        )
        .child(
            div()
                .flex()
                .flex_wrap()
                .gap_x_4()
                .gap_y_1()
                .child(stat_item("Sheets", er.sheets_exported, text_muted))
                .child(stat_item("Cells", er.cells_exported, text_muted))
                .child(stat_item("Formulas", er.formulas_exported, text_muted))
        )
        .child(
            div()
                .mt_1()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child(format!("Export time: {} ms", er.export_duration_ms))
        )
}

fn stat_item(label: &'static str, value: usize, color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .gap_1()
        .text_size(px(11.0))
        .text_color(color)
        .child(format!("{}:", label))
        .child(
            div()
                .font_weight(FontWeight::MEDIUM)
                .child(format_number(value))
        )
}

fn render_formula_conversions(
    er: &ExportResult,
    text_primary: Hsla,
    text_muted: Hsla,
    warning_color: Hsla,
) -> impl IntoElement {
    if er.converted_formulas.is_empty() {
        return div().into_any_element();
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(warning_color)
                        .child(format!("Formulas Converted to Values ({})", er.converted_formulas.len()))
                )
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("These formulas could not be exported to Excel and were replaced with their computed values:")
        )
        .child(
            div()
                .mt_1()
                .max_h(px(150.0))
                .overflow_hidden()
                .border_1()
                .border_color(warning_color.opacity(0.3))
                .rounded_sm()
                .bg(warning_color.opacity(0.05))
                .p_2()
                .flex()
                .flex_col()
                .gap_1()
                .children(er.converted_formulas.iter().take(20).map(|cf| {
                    div()
                        .flex()
                        .items_baseline()
                        .gap_2()
                        .text_size(px(10.0))
                        .child(
                            div()
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text_primary)
                                .child(format!("{}!{}", cf.sheet, cf.address))
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .child(format!("{} -> {}", cf.formula, cf.value))
                        )
                }))
                .when(er.converted_formulas.len() > 20, |this| {
                    this.child(
                        div()
                            .mt_1()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .italic()
                            .child(format!("... and {} more", er.converted_formulas.len() - 20))
                    )
                })
        )
        .into_any_element()
}

fn render_precision_warnings(
    er: &ExportResult,
    text_primary: Hsla,
    text_muted: Hsla,
    warning_color: Hsla,
) -> impl IntoElement {
    if er.precision_warnings.is_empty() {
        return div().into_any_element();
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(warning_color)
                        .child(format!("Large Numbers Exported as Text ({})", er.precision_warnings.len()))
                )
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("These numbers exceed Excel's 15-digit precision limit and were exported as text to preserve accuracy:")
        )
        .child(
            div()
                .mt_1()
                .max_h(px(100.0))
                .overflow_hidden()
                .border_1()
                .border_color(warning_color.opacity(0.3))
                .rounded_sm()
                .bg(warning_color.opacity(0.05))
                .p_2()
                .flex()
                .flex_col()
                .gap_1()
                .children(er.precision_warnings.iter().take(10).map(|pw| {
                    div()
                        .flex()
                        .items_baseline()
                        .gap_2()
                        .text_size(px(10.0))
                        .child(
                            div()
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text_primary)
                                .child(format!("{}!{}", pw.sheet, pw.address))
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .child(pw.value.clone())
                        )
                }))
                .when(er.precision_warnings.len() > 10, |this| {
                    this.child(
                        div()
                            .mt_1()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .italic()
                            .child(format!("... and {} more", er.precision_warnings.len() - 10))
                    )
                })
        )
        .into_any_element()
}

/// Format a number with thousand separators for readability
fn format_number(n: usize) -> String {
    if n < 1000 {
        return n.to_string();
    }

    let s = n.to_string();
    let mut result = String::new();
    for (i, c) in s.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    result.chars().rev().collect()
}
