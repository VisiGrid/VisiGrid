//! Import Report dialog for Excel file imports
//!
//! Shows detailed statistics and warnings after importing xlsx/xls/xlsb/ods files.

use gpui::*;
use visigrid_io::xlsx::ImportResult;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button};

/// Render the Import Report dialog overlay
pub fn render_import_report_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let Some(import_result) = &app.import_result else {
        return div().into_any_element();
    };

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);

    let filename = app.import_filename.as_deref().unwrap_or("unknown file");

    modal_overlay(
        "import-report-dialog",
        |this, cx| this.hide_import_report(cx),
        div()
            .w(px(500.0))
            .max_h(px(600.0))
            .bg(panel_bg)
            .border_1()
            .border_color(panel_border)
            .rounded_lg()
            .shadow_xl()
            .overflow_hidden()
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
                                .child("Import Report")
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
                        .child(render_summary_section(import_result, text_primary, text_muted))
                        // Quality section (if there are issues)
                        .child(render_quality_section(import_result, text_primary, text_muted, warning_color, error_color))
                        // Per-sheet table
                        .child(render_sheet_table(import_result, text_primary, text_muted, text_disabled, panel_border))
                        // Warnings section
                        .child(render_warnings_section(import_result, text_muted, warning_color))
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
                        .justify_end()
                        .child(
                            Button::new("import-report-close-btn", "Close")
                                .primary(accent, app.token(TokenKey::TextInverse))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_import_report(cx);
                                }))
                        )
                ),
        cx,
    ).into_any_element()
}

fn render_summary_section(ir: &ImportResult, text_primary: Hsla, text_muted: Hsla) -> impl IntoElement {
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
                .child(stat_item("Sheets", ir.sheets_imported, text_muted))
                .child(stat_item("Cells", ir.cells_imported, text_muted))
                .child(stat_item("Formulas", ir.formulas_imported, text_muted))
                .child(stat_item("Dates/Times", ir.dates_imported, text_muted))
        )
        .child(
            div()
                .mt_1()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child(format!("Import time: {} ms", ir.import_duration_ms))
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

fn render_quality_section(
    ir: &ImportResult,
    text_primary: Hsla,
    text_muted: Hsla,
    warning_color: Hsla,
    error_color: Hsla,
) -> impl IntoElement {
    let has_issues = ir.formulas_failed > 0 || ir.formulas_with_unknowns > 0 || !ir.unsupported_functions.is_empty();

    if !has_issues {
        return div().into_any_element();
    }

    let top_funcs = ir.top_unsupported_functions(8);

    // Build children conditionally
    let mut children: Vec<AnyElement> = Vec::new();

    // Header
    children.push(
        div()
            .text_size(px(12.0))
            .font_weight(FontWeight::MEDIUM)
            .text_color(text_primary)
            .child("Quality")
            .into_any_element()
    );

    // Parse errors
    if ir.formulas_failed > 0 {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(format!("Formula parse errors: {}", ir.formulas_failed))
                .into_any_element()
        );
    }

    // Formulas with unknown functions
    if ir.formulas_with_unknowns > 0 {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(warning_color)
                .child(format!("Formulas with unsupported functions: {}", ir.formulas_with_unknowns))
                .into_any_element()
        );
    }

    // Top unsupported functions list
    if !top_funcs.is_empty() {
        children.push(
            div()
                .mt_1()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("Top unsupported functions:")
                )
                .child(
                    div()
                        .flex()
                        .flex_wrap()
                        .gap_1()
                        .children(top_funcs.iter().map(|(name, count)| {
                            div()
                                .px_2()
                                .py(px(2.0))
                                .bg(warning_color.opacity(0.15))
                                .rounded_sm()
                                .text_size(px(10.0))
                                .text_color(warning_color)
                                .child(format!("{} ({})", name, count))
                        }))
                )
                .into_any_element()
        );
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .children(children)
        .into_any_element()
}

fn render_sheet_table(
    ir: &ImportResult,
    text_primary: Hsla,
    text_muted: Hsla,
    text_disabled: Hsla,
    border_color: Hsla,
) -> impl IntoElement {
    if ir.sheet_stats.is_empty() {
        return div().into_any_element();
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child("Per-Sheet Details")
        )
        .child(
            div()
                .border_1()
                .border_color(border_color)
                .rounded_sm()
                .overflow_hidden()
                // Header row
                .child(
                    div()
                        .flex()
                        .bg(border_color.opacity(0.3))
                        .text_size(px(10.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(text_muted)
                        .child(div().w(px(120.0)).px_2().py_1().child("Sheet"))
                        .child(div().w(px(60.0)).px_2().py_1().text_right().child("Cells"))
                        .child(div().w(px(60.0)).px_2().py_1().text_right().child("Formulas"))
                        .child(div().w(px(50.0)).px_2().py_1().text_right().child("Errors"))
                        .child(div().w(px(60.0)).px_2().py_1().text_right().child("Unknown"))
                        .child(div().w(px(60.0)).px_2().py_1().child("Truncated"))
                )
                // Data rows
                .children(ir.sheet_stats.iter().enumerate().map(|(idx, stats)| {
                    let row_bg = if idx % 2 == 0 {
                        hsla(0.0, 0.0, 0.0, 0.0)
                    } else {
                        border_color.opacity(0.1)
                    };

                    let truncated = stats.truncated_rows > 0 || stats.truncated_cols > 0;

                    div()
                        .flex()
                        .bg(row_bg)
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(
                            div()
                                .w(px(120.0))
                                .px_2()
                                .py_1()
                                .overflow_hidden()
                                .text_ellipsis()
                                .child(stats.name.clone())
                        )
                        .child(div().w(px(60.0)).px_2().py_1().text_right().child(format_number(stats.cells_imported)))
                        .child(div().w(px(60.0)).px_2().py_1().text_right().child(format_number(stats.formulas_imported)))
                        .child(
                            div()
                                .w(px(50.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.formulas_with_errors > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.formulas_with_errors))
                        )
                        .child(
                            div()
                                .w(px(60.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.formulas_with_unknowns > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.formulas_with_unknowns))
                        )
                        .child(
                            div()
                                .w(px(60.0))
                                .px_2()
                                .py_1()
                                .text_color(if truncated { text_primary } else { text_disabled })
                                .child(if truncated { "Yes" } else { "-" })
                        )
                }))
        )
        .into_any_element()
}

fn render_warnings_section(ir: &ImportResult, text_muted: Hsla, warning_color: Hsla) -> impl IntoElement {
    if ir.warnings.is_empty() {
        return div().into_any_element();
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(warning_color)
                .child("Warnings")
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(ir.warnings.iter().map(|warning| {
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(format!("- {}", warning))
                }))
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
