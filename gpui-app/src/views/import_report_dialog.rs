//! Import Report dialog for Excel file imports
//!
//! Shows detailed statistics and warnings after importing xlsx/xls/xlsb/ods files.

use gpui::*;
use gpui::prelude::FluentBuilder;
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
    let text_inverse = app.token(TokenKey::TextInverse);
    let warning_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);
    let success_color = app.token(TokenKey::CellStyleSuccessText);

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
                        .child(render_quality_section(import_result, text_primary, text_muted, warning_color, error_color, success_color, accent))
                        // Current calculation mode (live workbook state)
                        .child(render_calculation_mode_section(app, cx, import_result, text_primary, text_muted, warning_color, error_color, success_color))
                        // Per-sheet table
                        .child(render_sheet_table(import_result, text_primary, text_muted, text_disabled, panel_border))
                        // Warnings section
                        .child(render_warnings_section(import_result, text_muted, warning_color))
                )
                // Footer with cycle action buttons and close button
                .child({
                    let has_unresolved = app.has_unresolved_cycles(cx);
                    let is_xlsx = app.current_file.as_ref()
                        .and_then(|p| p.extension()).and_then(|e| e.to_str())
                        .map_or(false, |e| matches!(e.to_lowercase().as_str(), "xlsx" | "xls" | "xlsm" | "xlsb" | "ods"));

                    div()
                        .w_full()
                        .px_4()
                        .py_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_end()
                        .gap_2()
                        // "Turn on iterative calculation..." button (when unresolved cycles exist)
                        .when(has_unresolved, |d| {
                            d.child(
                                Button::new("enable-iter-btn", "Turn on iterative calculation\u{2026}")
                                    .primary(accent, text_inverse)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.enable_iteration_and_recalc(cx);
                                    }))
                            )
                        })
                        // "Freeze cycle values" button (when unresolved cycles exist)
                        .when(has_unresolved, |d| {
                            let can_freeze = is_xlsx && !import_result.freeze_applied;
                            let btn = Button::new("freeze-cycles-btn", "Freeze cycle values")
                                .disabled(!can_freeze)
                                .secondary(panel_border, text_primary);
                            if can_freeze {
                                d.child(
                                    btn.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.reimport_with_freeze(cx);
                                    }))
                                )
                            } else {
                                d.child(btn)
                            }
                        })
                        .child(
                            Button::new("import-report-close-btn", "Close")
                                .secondary(panel_border, text_primary)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_import_report(cx);
                                }))
                        )
                }),
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
                .when(ir.shared_formula_groups > 0, |d| {
                    d.child(stat_item("Shared groups", ir.shared_formula_groups, text_muted))
                })
                .when(ir.formula_cells_without_values > 0, |d| {
                    d.child(stat_item("Formula backfill", ir.formula_cells_without_values, text_muted))
                })
                .when(ir.value_cells_backfilled > 0, |d| {
                    d.child(stat_item("Value backfill", ir.value_cells_backfilled, text_muted))
                })
                .when(ir.styles_imported > 0, |d| {
                    d.child(stat_item("Styled cells", ir.styles_imported, text_muted))
                        .child(stat_item("Unique styles", ir.unique_styles, text_muted))
                })
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
    success_color: Hsla,
    accent: Hsla,
) -> impl IntoElement {
    let has_issues = ir.formulas_failed > 0
        || ir.formulas_with_unknowns > 0
        || !ir.unsupported_functions.is_empty()
        || ir.recalc_errors > 0
        || ir.recalc_circular > 0
        || ir.freeze_applied;

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

        // Show sample formulas that failed to parse
        if !ir.parse_error_samples.is_empty() {
            children.push(
                div()
                    .ml_2()
                    .flex()
                    .flex_col()
                    .gap(px(2.0))
                    .children(ir.parse_error_samples.iter().take(5).map(|sample| {
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child(sample.clone())
                    }))
                    .into_any_element()
            );
        }
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

    // Post-recalc errors
    if ir.recalc_errors > 0 {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(format!("Formula errors (post-recalc): {}", ir.recalc_errors))
                .into_any_element()
        );
    }

    // Circular references (at import)
    if ir.recalc_circular > 0 && !ir.freeze_applied {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(format!("Circular references (at import): {}", ir.recalc_circular))
                .into_any_element()
        );
    }

    // Freeze results — before/after delta
    if ir.freeze_applied {
        let total_cycles = ir.cycles_frozen + ir.cycles_no_cached;
        let remaining = ir.cycles_no_cached;
        let delta_msg = format!(
            "#CYCLE! cells: {} \u{2192} {} (frozen {})",
            total_cycles, remaining, ir.cycles_frozen
        );
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(success_color)
                .child(delta_msg)
                .into_any_element()
        );
        if remaining > 0 {
            children.push(
                div()
                    .text_size(px(10.0))
                    .text_color(warning_color)
                    .child(format!("{} cycle cells had no cached value in the XLSX and remain as #CYCLE!.", remaining))
                    .into_any_element()
            );
        }
        children.push(
            div()
                .flex()
                .items_center()
                .gap_1()
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(warning_color)
                        .child("Frozen cells use Excel\u{2019}s cached values and will not recalculate if you edit their inputs.")
                )
                .child(
                    div()
                        .id("freeze-explain-link")
                        .text_size(px(10.0))
                        .text_color(accent)
                        .cursor_pointer()
                        .hover(|s| s.underline())
                        .on_click(|_, _, cx| {
                            let _ = open::that(crate::docs_links::DOCS_CIRCULAR_REFS);
                        })
                        .child("Learn more")
                )
                .into_any_element()
        );
    }

    // Error examples (top N concrete cells)
    if !ir.recalc_error_examples.is_empty() {
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
                        .child("Examples:")
                )
                .children(ir.recalc_error_examples.iter().map(|ex| {
                    let detail = if let Some(ref f) = ex.formula {
                        format!("{}!{}  {}  {}  {}", ex.sheet, ex.address, ex.kind, ex.error, f)
                    } else {
                        format!("{}!{}  {}  {}", ex.sheet, ex.address, ex.kind, ex.error)
                    };
                    div()
                        .text_size(px(10.0))
                        .text_color(error_color)
                        .child(detail)
                }))
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("These errors occur when imported formulas cannot be evaluated correctly.")
                )
                .into_any_element()
        );
    } else if ir.recalc_errors > 0 || ir.recalc_circular > 0 {
        // Explanatory text without examples (shouldn't happen, but defensive)
        children.push(
            div()
                .mt_1()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("These errors occur when imported formulas cannot be evaluated correctly.")
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
                        .child(div().w(px(90.0)).px_2().py_1().child("Sheet"))
                        .child(div().w(px(55.0)).px_2().py_1().text_right().child("Cells"))
                        .child(div().w(px(55.0)).px_2().py_1().text_right().child("Formulas"))
                        .child(div().w(px(45.0)).px_2().py_1().text_right().child("Parse"))
                        .child(div().w(px(50.0)).px_2().py_1().text_right().child("Recalc"))
                        .child(div().w(px(40.0)).px_2().py_1().text_right().child("Circ"))
                        .child(div().w(px(50.0)).px_2().py_1().text_right().child("Unknown"))
                        .child(div().w(px(45.0)).px_2().py_1().child("Trunc"))
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
                                .w(px(90.0))
                                .px_2()
                                .py_1()
                                .overflow_hidden()
                                .text_ellipsis()
                                .child(stats.name.clone())
                        )
                        .child(div().w(px(55.0)).px_2().py_1().text_right().child(format_number(stats.cells_imported)))
                        .child(div().w(px(55.0)).px_2().py_1().text_right().child(format_number(stats.formulas_imported)))
                        .child(
                            div()
                                .w(px(45.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.formulas_with_errors > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.formulas_with_errors))
                        )
                        .child(
                            div()
                                .w(px(50.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.recalc_errors > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.recalc_errors))
                        )
                        .child(
                            div()
                                .w(px(40.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.recalc_circular > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.recalc_circular))
                        )
                        .child(
                            div()
                                .w(px(50.0))
                                .px_2()
                                .py_1()
                                .text_right()
                                .text_color(if stats.formulas_with_unknowns > 0 { text_primary } else { text_disabled })
                                .child(format_number(stats.formulas_with_unknowns))
                        )
                        .child(
                            div()
                                .w(px(45.0))
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

/// Render the "Current calculation mode" section reflecting live workbook state.
fn render_calculation_mode_section(
    app: &Spreadsheet,
    cx: &App,
    ir: &ImportResult,
    text_primary: Hsla,
    _text_muted: Hsla,
    warning_color: Hsla,
    error_color: Hsla,
    success_color: Hsla,
) -> impl IntoElement {
    let iterative = app.wb(cx).iterative_enabled();
    let rr = app.last_recalc_report.as_ref();
    let converged = rr.map_or(false, |r| r.converged);
    let scc_count = rr.map_or(0, |r| r.scc_count);
    let iters = rr.map_or(0, |r| r.iterations_performed);
    let cycle_count = app.current_cycle_count(cx);
    let unresolved = app.has_unresolved_cycles(cx);

    // Only show if there's something to report
    let has_content = (iterative && cycle_count > 0) || ir.freeze_applied || unresolved;
    if !has_content {
        return div().into_any_element();
    }

    let mut children: Vec<AnyElement> = Vec::new();

    children.push(
        div()
            .text_size(px(12.0))
            .font_weight(FontWeight::MEDIUM)
            .text_color(text_primary)
            .child("Current Calculation Mode")
            .into_any_element()
    );

    if iterative && cycle_count > 0 && converged {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(success_color)
                .child(format!(
                    "Iterative calculation: {} cycle cells in {} groups \u{2014} converged in {} iterations",
                    cycle_count, scc_count, iters
                ))
                .into_any_element()
        );
    } else if iterative && cycle_count > 0 && !converged {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(warning_color)
                .child(format!(
                    "Iterative calculation: {} cycle cells in {} groups \u{2014} did not converge (max iterations hit)",
                    cycle_count, scc_count
                ))
                .into_any_element()
        );
    } else if ir.freeze_applied {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(warning_color)
                .child(format!("Cycle values frozen: {} (Excel cached)", ir.cycles_frozen))
                .into_any_element()
        );
    } else if unresolved {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(format!("Circular references: {} cells (#CYCLE!)", cycle_count))
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
