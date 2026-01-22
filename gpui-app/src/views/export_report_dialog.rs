//! Export Report dialog for Excel exports
//!
//! Shows detailed statistics and warnings after exporting to xlsx.
//! Displayed automatically when export has warnings (formula conversions, precision limits).

use gpui::*;
use gpui::prelude::FluentBuilder;
use visigrid_io::xlsx::ExportResult;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the Export Report dialog overlay
pub fn render_export_report_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Warn);

    let export_result = match &app.export_result {
        Some(er) => er,
        None => return div().into_any_element(),
    };

    let filename = app.export_filename.as_deref().unwrap_or("unknown file");

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_export_report(cx);
            cx.stop_propagation();
        }))
        .child(
            div()
                .w(px(500.0))
                .max_h(px(500.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_xl()
                .overflow_hidden()
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .flex()
                .flex_col()
                // Header
                .child(
                    div()
                        .px_4()
                        .py_3()
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(14.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Export Report")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(filename.to_string())
                        )
                )
                // Content (scrollable)
                .child(
                    div()
                        .flex_1()
                        .overflow_hidden()
                        .p_4()
                        .flex()
                        .flex_col()
                        .gap_4()
                        // Summary section
                        .child(render_summary_section(export_result, text_primary, text_muted))
                        // Formula conversions section
                        .child(render_formula_conversions(export_result, text_primary, text_muted, warning_color))
                        // Precision warnings section
                        .child(render_precision_warnings(export_result, text_primary, text_muted, warning_color))
                )
                // Footer with close button
                .child(
                    div()
                        .w_full()
                        .px_4()
                        .py_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            // Copy report button (if there are warnings)
                            if export_result.has_warnings() {
                                let report = export_result.full_report_with_context(filename);
                                div()
                                    .id("export-report-copy-btn")
                                    .px_3()
                                    .py(px(6.0))
                                    .border_1()
                                    .border_color(panel_border)
                                    .rounded_md()
                                    .cursor_pointer()
                                    .text_size(px(12.0))
                                    .text_color(text_muted)
                                    .hover(|s| s.bg(panel_border.opacity(0.3)))
                                    .on_mouse_down(MouseButton::Left, move |_, _, cx| {
                                        cx.write_to_clipboard(ClipboardItem::new_string(report.clone()));
                                        cx.stop_propagation();
                                    })
                                    .child("Copy Details")
                                    .into_any_element()
                            } else {
                                div().into_any_element()
                            }
                        )
                        .child(
                            div()
                                .id("export-report-close-btn")
                                .px_4()
                                .py(px(6.0))
                                .bg(accent)
                                .rounded_md()
                                .cursor_pointer()
                                .text_size(px(12.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(app.token(TokenKey::TextInverse))
                                .hover(|s| s.opacity(0.9))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_export_report(cx);
                                }))
                                .child("Close")
                        )
                )
        )
        .into_any_element()
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
