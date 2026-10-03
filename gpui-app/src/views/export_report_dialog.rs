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
    if let Some(review) = &app.pending_xlsx_export {
        return render_export_review(app, review, cx);
    }
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
        .child(render_precision_warnings(export_result, text_primary, text_muted, warning_color))
        .child(div().flex().flex_col().gap_2().children(export_result.warnings.iter().map(|warning|
            div().text_size(px(12.0)).text_color(warning_color).child(warning.clone()))));

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
                .child(stat_item("Tables", er.tables_exported, text_muted))
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

/// Review losses in the app's scrollable modal instead of the platform prompt,
/// whose Linux fallback clips long paragraphs.
fn render_export_review(app: &Spreadsheet, review: &crate::xlsx_export::ExportReview, cx: &mut Context<Spreadsheet>) -> AnyElement {
    use visigrid_io::xlsx::ExportOrder;
    let warnings = review.warnings();
    let panel = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let warn = app.token(TokenKey::Warn);
    let width: f32 = app.window_size.width.into();
    let height: f32 = app.window_size.height.into();
    let cancel = Button::new("xlsx-review-cancel", "Cancel")
        .secondary(border, muted)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.hide_export_report(cx)));
    let label = if review.has_sort {
        match review.order() { ExportOrder::Sorted => "Export sorted…", ExportOrder::Stored => "Export stored order…" }
    } else { "Continue to export…" };
    let export = Button::new("xlsx-review-export", label)
        .primary(app.token(TokenKey::Accent), app.token(TokenKey::TextInverse))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.confirm_xlsx_export(cx)));
    modal_overlay("xlsx-export-review", |this, cx| this.hide_export_report(cx),
        div().w(px(620.0_f32.min((width - 48.0).max(280.0))))
            .max_h(px(640.0_f32.min((height - 64.0).max(240.0))))
            .bg(panel).border_1().border_color(border).rounded_lg().shadow_xl()
            .overflow_hidden().flex().flex_col()
            .child(div().px_5().py_4().flex_shrink_0().border_b_1().border_color(border)
                .child(div().text_size(px(16.0)).font_weight(FontWeight::SEMIBOLD).text_color(text).child("Review Excel export"))
                .child(div().mt_1().text_size(px(12.0)).text_color(muted).child(if review.has_sort { "Choose how Table rows are saved and review changes to the Excel copy." } else { "Review changes to the Excel copy before choosing a destination." })))
            .child(div().id("xlsx-export-review-scroll").min_h_0().overflow_y_scroll()
                .p_5().flex().flex_col().gap_3()
                .when(review.has_sort, |body| body
                    .child(div().text_size(px(12.0)).font_weight(FontWeight::SEMIBOLD).text_color(text).child("Table row order"))
                    .children([
                        ("xlsx-order-sorted", ExportOrder::Sorted, "Export sorted", "Open with records in the current sort order. Supported formulas follow their records."),
                        ("xlsx-order-stored", ExportOrder::Stored, "Keep stored order", "Keep cells and formulas at their stored addresses. Use Reapply in Excel to display the saved sort."),
                    ].into_iter().map(|(id, order, title, description)| {
                        let selected = review.order() == order;
                        let disabled = order == ExportOrder::Sorted && review.sorted.is_err();
                        let accent = app.token(TokenKey::Accent);
                        div().id(id).p_3().rounded_md().border_1()
                            .border_color(if selected { accent } else { border })
                            .bg(if selected { accent.opacity(0.08) } else { panel })
                            .when(disabled, |card| card.opacity(0.5))
                            .when(!disabled, |card| card.cursor_pointer()
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| this.select_xlsx_export_order(order, cx))))
                            .child(div().flex().justify_between()
                                .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(text).child(title))
                                .when(selected, |row| row.child(div().text_size(px(11.0)).text_color(accent).child("Selected"))))
                            .child(div().mt_1().text_size(px(12.0)).text_color(muted).child(description))
                    }))
                    .when_some(review.sorted.as_ref().err(), |body, error| body.child(
                        div().p_3().rounded_md().bg(warn.opacity(0.05))
                            .text_size(px(12.0)).text_color(text)
                            .child(div().font_weight(FontWeight::SEMIBOLD).child("Sorted export unavailable"))
                            .child(div().mt_1().child(error.clone())))))
                .children(warnings.iter().map(|warning| div().p_3().rounded_md()
                    .bg(warn.opacity(0.05)).border_1().border_color(warn.opacity(0.3))
                    .text_size(px(12.0)).text_color(text).child(warning.clone()))))
            .child(div().px_5().py_4().flex_shrink_0().border_t_1().border_color(border)
                .child(div().mb_3().text_size(px(12.0)).text_color(muted).child("Your original workbook stays unchanged."))
                .child(div().flex().justify_end().gap_2().child(cancel).child(export))), cx).into_any_element()
}
