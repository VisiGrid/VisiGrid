//! A readable import outcome, compatibility findings, and calculation diagnostics.

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_io::xlsx::{ImportResult, SheetStats};

use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button};

#[derive(Clone, Copy)]
struct Colors {
    panel: Hsla,
    surface: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
    accent: Hsla,
    warning: Hsla,
    error: Hsla,
    success: Hsla,
}

const RECALC_EXPLANATION: &str = "These errors were found after recalculation. They may already exist in the source workbook; this report does not establish that importing caused them.";

pub fn render_import_report_dialog(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let Some(ir) = &app.import_result else {
        return div().into_any_element();
    };
    let c = Colors {
        panel: app.token(TokenKey::PanelBg),
        surface: app.token(TokenKey::EditorBg),
        border: app.token(TokenKey::PanelBorder),
        text: app.token(TokenKey::TextPrimary),
        muted: app.token(TokenKey::TextMuted),
        accent: app.token(TokenKey::Accent),
        warning: app.token(TokenKey::Warn),
        error: app.token(TokenKey::Error),
        success: app.token(TokenKey::CellStyleSuccessText),
    };
    let inverse = app.token(TokenKey::TextInverse);
    let filename = app.import_filename.as_deref().unwrap_or("Unknown file");
    let notes = compatibility_notes(ir);
    let calculations = calculation_notes(ir);
    let has_findings = !notes.is_empty() || !calculations.is_empty();
    let report_text = report_text(ir, filename);
    let window_height: f32 = app.window_size.height.into();
    let window_width: f32 = app.window_size.width.into();
    let expanded = app.import_report_details_expanded;
    let has_unresolved = app.has_unresolved_cycles(cx);
    let is_excel_source = app
        .current_file
        .as_ref()
        .and_then(|p| p.extension())
        .and_then(|e| e.to_str())
        .is_some_and(|e| {
            matches!(
                e.to_lowercase().as_str(),
                "xlsx" | "xls" | "xlsm" | "xlsb" | "ods"
            )
        });

    modal_overlay("import-report-dialog", |this, cx| this.hide_import_report(cx),
        div().w(px(760.0_f32.min((window_width - 48.0).max(280.0))))
            .max_h(px(820.0_f32.min((window_height - 64.0).max(240.0))))
            .bg(c.panel).border_1().border_color(c.border).rounded_lg().shadow_xl()
            .overflow_hidden().flex().flex_col()
            .child(div().px_5().py_4().flex_shrink_0().border_b_1().border_color(c.border)
                .flex().flex_col().gap_1()
                .child(div().text_size(px(16.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child("Import report"))
                .child(div().text_size(px(12.0)).text_color(c.muted).truncate().child(filename.to_string())))
            .child(div().id("import-report-scroll").min_h_0().overflow_y_scroll()
                .p_5().flex().flex_col().gap_4()
                .child(card(c).child(
                    div().flex().flex_col().gap_2()
                        .child(div().text_size(px(18.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text)
                            .child(format!("Imported {}", counted(ir.sheets_imported, "sheet", "sheets"))))
                        .child(div().text_size(px(12.0)).text_color(if has_findings { c.warning } else { c.muted })
                            .child(outcome_detail(ir, !notes.is_empty())))
                ))
                .child(div().flex().flex_wrap().gap_2()
                    .children([
                        ("Sheets", format_number(ir.sheets_imported)),
                        ("Cells", format_number(ir.cells_imported)),
                        ("Formulas", format_number(ir.formulas_imported)),
                        ("Import time", format!("{} ms", format_number(ir.import_duration_ms))),
                    ].into_iter().map(|(label, value)| {
                        div().flex_1().min_w(px(125.0)).p_3().rounded_md().border_1().border_color(c.border).bg(c.surface)
                            .child(div().text_size(px(11.0)).text_color(c.muted).child(label))
                            .child(div().mt_1().text_size(px(20.0)).font_weight(FontWeight::MEDIUM).text_color(c.text).child(value))
                    })))
                .when(!notes.is_empty(), |d| d.child(
                    card(c).border_color(c.warning.opacity(0.45))
                        .child(section_title("Compatibility & import warnings", c))
                        .child(body_text("These findings can affect what is retained or understood from the source file.", c.muted))
                        .children(notes.iter().map(|note| body_text(note.clone(), c.text)))
                        .when(!ir.parse_error_samples.is_empty(), |d| d.child(
                            div().flex().flex_col().gap_2().mt_1()
                                .child(body_text("Formula parse examples", c.muted))
                                .children(ir.parse_error_samples.iter().take(5).map(|sample| code_line(sample.clone(), c)))
                        ))
                ))
                .when(!calculations.is_empty(), |d| d.child(
                    card(c).child(section_title("Calculation results", c))
                        .children(calculations.iter().map(|note| body_text(note.clone(), c.text)))
                        .when(ir.recalc_errors > 0 || ir.recalc_circular > 0, |d| d.child(body_text(RECALC_EXPLANATION, c.muted)))
                        .children(ir.recalc_error_examples.iter().enumerate().map(|(i, ex)| {
                            let sheet_name = ex.sheet.clone();
                            let address = ex.address.clone();
                            div().p_3().rounded_md().bg(c.panel).border_1().border_color(c.border)
                                .flex().flex_col().gap_2()
                                .child(div().flex().flex_wrap().items_center().justify_between().gap_2()
                                    .child(div().id(SharedString::from(format!("import-error-cell-{i}")))
                                        .text_size(px(12.0)).font_weight(FontWeight::MEDIUM).text_color(c.accent)
                                        .cursor_pointer().hover(|s| s.underline())
                                        .child(format!("{} · {} ↗", ex.sheet, ex.address))
                                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
                                            cx.stop_propagation();
                                            let target = this.wb(cx).sheets().iter().position(|s| s.name == sheet_name)
                                                .zip(Spreadsheet::parse_cell_ref(&address));
                                            if let Some((index, (row, col))) = target {
                                                let in_bounds = this.wb(cx).sheet(index).is_some_and(|s| row < s.rows && col < s.cols);
                                                if in_bounds {
                                                    if !this.activate_sheet(index, cx) {
                                                        return;
                                                    }
                                                    let Some(view_row) = this.data_to_view(row, cx) else {
                                                        this.status_message = Some("That imported cell is hidden by the current filter.".into());
                                                        cx.notify();
                                                        return;
                                                    };
                                                    this.clear_selection_state();
                                                    this.view_state.additional_selections.clear();
                                                    this.view_state.selected = (view_row, col);
                                                    this.ensure_cell_visible(view_row, col);
                                                    this.hide_import_report(cx);
                                                    window.focus(&this.focus_handle, cx);
                                                    return;
                                                }
                                            }
                                            this.status_message = Some("That imported cell is no longer available at its original location.".into());
                                            cx.notify();
                                        })))
                                    .child(div().px_2().py_1().rounded_sm().bg(c.error.opacity(0.08))
                                        .text_size(px(12.0)).text_color(c.error).child(ex.error.clone())))
                                .when_some(ex.formula.clone(), |d, formula| d.child(code_line(formula, c)))
                        }))
                        .when(ir.freeze_applied, |d| d.child(
                            div().id("freeze-explain-link").text_size(px(12.0)).text_color(c.accent)
                                .cursor_pointer().hover(|s| s.underline()).child("About frozen cycle values ↗")
                                .on_mouse_down(MouseButton::Left, |_, _, _| {
                                    let _ = open::that(crate::docs_links::DOCS_CIRCULAR_REFS);
                                })
                        ))
                ))
                .child(render_calculation_mode_section(app, cx, ir, c.text, c.muted, c.warning, c.error, c.success))
                .child(render_sheet_table(ir, c))
                .child(card(c)
                    .child(div().id("import-diagnostics-toggle").flex().justify_between().items_center()
                        .cursor_pointer().text_size(px(12.0)).text_color(c.accent)
                        .child(if expanded { "Hide import diagnostics" } else { "Show import diagnostics" })
                        .child(if expanded { "−" } else { "+" })
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.import_report_details_expanded = !this.import_report_details_expanded;
                            cx.notify();
                        })))
                    .when(expanded, |d| d.children(diagnostics(ir).into_iter().map(|(label, value)| {
                        div().flex().justify_between().gap_3().text_size(px(12.0))
                            .child(div().text_color(c.muted).child(label))
                            .child(div().text_color(c.text).child(value))
                    }))))
            )
            .child(div().px_5().py_3().flex_shrink_0().border_t_1().border_color(c.border).bg(c.surface)
                .flex().flex_col().gap_3()
                .when(has_unresolved, |d| d.child(
                    div().flex().flex_wrap().gap_2()
                        .child(Button::new("enable-iter-btn", "Turn on iterative calculation…").secondary(c.border, c.text)
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.enable_iteration_and_recalc(cx))))
                        .child(Button::new("freeze-cycles-btn", "Freeze cycle values")
                            .disabled(!is_excel_source || ir.freeze_applied).secondary(c.border, c.text)
                            .when(is_excel_source && !ir.freeze_applied, |b| b.on_mouse_down(MouseButton::Left,
                                cx.listener(|this, _, _, cx| this.reimport_with_freeze(cx)))))
                ))
                .child(div().flex().flex_wrap().justify_between().gap_2()
                    .child(Button::new("copy-import-report", "Copy report").secondary(c.border, c.text)
                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                            cx.write_to_clipboard(ClipboardItem::new_string(report_text.clone()));
                            this.status_message = Some("Import report copied to clipboard".into());
                            cx.notify();
                        })))
                    .child(Button::new("import-report-close-btn", "Done").primary(c.accent, inverse)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.hide_import_report(cx)))))
            ), cx).into_any_element()
}

fn card(c: Colors) -> Div {
    div()
        .flex_shrink_0()
        .p_4()
        .rounded_md()
        .border_1()
        .border_color(c.border)
        .bg(c.surface)
        .flex()
        .flex_col()
        .gap_2()
}

fn section_title(title: &'static str, c: Colors) -> Div {
    div()
        .text_size(px(13.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(c.text)
        .child(title)
}

fn body_text(text: impl Into<SharedString>, color: Hsla) -> Div {
    div()
        .text_size(px(12.0))
        .text_color(color)
        .child(text.into())
}

fn code_line(text: String, c: Colors) -> Div {
    div()
        .text_size(px(12.0))
        .font_family("monospace")
        .text_color(c.text)
        .child(text)
}

fn counted(n: usize, singular: &str, plural: &str) -> String {
    format!(
        "{} {}",
        format_number(n),
        if n == 1 { singular } else { plural }
    )
}

fn outcome_detail(ir: &ImportResult, has_compatibility_notes: bool) -> String {
    if ir.truncated {
        return "Some data was not loaded. Review the import warnings below.".into();
    }
    if has_compatibility_notes {
        return "Review compatibility findings before saving back to Excel.".into();
    }
    if ir.recalc_errors > 0 {
        return format!(
            "{} found after recalculation. See the examples below.",
            counted(ir.recalc_errors, "formula error", "formula errors")
        );
    }
    if ir.recalc_circular > 0 || ir.freeze_applied {
        return "Review circular references and calculation settings below.".into();
    }
    "No issues reported by the importer.".into()
}

fn compatibility_notes(ir: &ImportResult) -> Vec<String> {
    let mut notes = Vec::new();
    if ir.truncated {
        notes.push("Data was truncated to fit the supported sheet dimensions. The imported workbook is incomplete.".into());
    }
    if ir.formulas_failed > 0 {
        notes.push(format!(
            "{} could not be parsed.",
            counted(ir.formulas_failed, "formula", "formulas")
        ));
    }
    if ir.formulas_with_unknowns > 0 {
        notes.push(format!(
            "{} {} unsupported functions.",
            counted(ir.formulas_with_unknowns, "formula", "formulas"),
            if ir.formulas_with_unknowns == 1 {
                "uses"
            } else {
                "use"
            }
        ));
    }
    if !ir.unsupported_functions.is_empty() {
        let mut functions: Vec<_> = ir.unsupported_functions.iter().collect();
        functions.sort_by(|a, b| a.0.cmp(b.0));
        notes.push(format!(
            "Unsupported functions: {}.",
            functions
                .into_iter()
                .map(|(name, count)| format!("{name} ({count})"))
                .collect::<Vec<_>>()
                .join(", ")
        ));
    }
    let dropped = ir.merges_dropped_overlap + ir.merges_dropped_invalid;
    if dropped > 0 {
        notes.push(format!(
            "{} dropped (overlapping or invalid).",
            counted(dropped, "merged region", "merged regions")
        ));
    }
    if ir.validations_skipped > 0 {
        notes.push(format!(
            "{} skipped.",
            counted(
                ir.validations_skipped,
                "validation rule",
                "validation rules"
            )
        ));
    }
    for warning in &ir.warnings {
        if !notes.contains(warning) {
            notes.push(warning.clone());
        }
    }
    for feature in &ir.unsupported_format_features {
        let warning = format!("Unsupported formatting: {feature}");
        if !notes.contains(&warning) {
            notes.push(warning);
        }
    }
    notes
}

fn calculation_notes(ir: &ImportResult) -> Vec<String> {
    let mut notes = Vec::new();
    if ir.recalc_errors > 0 {
        notes.push(format!(
            "{} found after recalculation.",
            counted(ir.recalc_errors, "formula error", "formula errors")
        ));
    }
    if ir.recalc_circular > 0 {
        notes.push(format!(
            "{} reported at import.",
            counted(
                ir.recalc_circular,
                "circular reference",
                "circular references"
            )
        ));
    }
    if ir.freeze_applied {
        notes.push(format!(
            "Cycle cells: {} → {} remaining; {} frozen to Excel's cached values.",
            ir.cycles_frozen + ir.cycles_no_cached,
            ir.cycles_no_cached,
            ir.cycles_frozen
        ));
        notes.push("Frozen cells will not recalculate when their inputs change.".into());
        if ir.cycles_no_cached > 0 {
            notes.push(format!(
                "{} had no cached value: #CYCLE! remains.",
                counted(ir.cycles_no_cached, "cycle cell", "cycle cells")
            ));
        }
    }
    notes
}

fn sheet_findings(stats: &SheetStats) -> String {
    let mut notes = Vec::new();
    if stats.formulas_with_errors > 0 {
        notes.push(counted(
            stats.formulas_with_errors,
            "parse failure",
            "parse failures",
        ));
    }
    if stats.formulas_with_unknowns > 0 {
        notes.push(counted(
            stats.formulas_with_unknowns,
            "unsupported formula",
            "unsupported formulas",
        ));
    }
    if stats.recalc_errors > 0 {
        notes.push(counted(
            stats.recalc_errors,
            "formula error",
            "formula errors",
        ));
    }
    if stats.recalc_circular > 0 {
        notes.push(counted(
            stats.recalc_circular,
            "circular reference",
            "circular references",
        ));
    }
    if stats.truncated_rows > 0 || stats.truncated_cols > 0 {
        notes.push(format!(
            "Truncated: {} rows, {} columns",
            stats.truncated_rows, stats.truncated_cols
        ));
    }
    if notes.is_empty() {
        "None reported".into()
    } else {
        notes.join(" · ")
    }
}

fn render_sheet_table(ir: &ImportResult, c: Colors) -> impl IntoElement {
    if ir.sheet_stats.is_empty() {
        return div().into_any_element();
    }
    card(c)
        .child(section_title("By sheet", c))
        .child(
            div()
                .id("import-sheet-table-scroll")
                .overflow_x_scroll()
                .child(
                    div()
                        .min_w(px(610.0))
                        .rounded_md()
                        .border_1()
                        .border_color(c.border)
                        .overflow_hidden()
                        .child(
                            div()
                                .flex()
                                .px_3()
                                .py_2()
                                .gap_3()
                                .bg(c.panel)
                                .text_size(px(11.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(c.muted)
                                .child(div().w(px(150.0)).flex_shrink_0().child("Sheet"))
                                .child(
                                    div()
                                        .w(px(75.0))
                                        .flex_shrink_0()
                                        .text_right()
                                        .child("Cells"),
                                )
                                .child(
                                    div()
                                        .w(px(80.0))
                                        .flex_shrink_0()
                                        .text_right()
                                        .child("Formulas"),
                                )
                                .child(div().flex_1().child("Findings")),
                        )
                        .children(ir.sheet_stats.iter().map(|stats| {
                            let findings = sheet_findings(stats);
                            div()
                                .flex()
                                .items_start()
                                .px_3()
                                .py_3()
                                .gap_3()
                                .border_t_1()
                                .border_color(c.border)
                                .text_size(px(12.0))
                                .text_color(c.text)
                                .child(div().w(px(150.0)).flex_shrink_0().child(stats.name.clone()))
                                .child(
                                    div()
                                        .w(px(75.0))
                                        .flex_shrink_0()
                                        .text_right()
                                        .child(format_number(stats.cells_imported)),
                                )
                                .child(
                                    div()
                                        .w(px(80.0))
                                        .flex_shrink_0()
                                        .text_right()
                                        .child(format_number(stats.formulas_imported)),
                                )
                                .child(
                                    div()
                                        .flex_1()
                                        .min_w_0()
                                        .text_color(if findings == "None reported" {
                                            c.muted
                                        } else {
                                            c.warning
                                        })
                                        .child(findings),
                                )
                        })),
                ),
        )
        .into_any_element()
}

fn diagnostics(ir: &ImportResult) -> Vec<(&'static str, String)> {
    vec![
        ("Dates / times", format_number(ir.dates_imported)),
        ("Styled cells", format_number(ir.styles_imported)),
        ("Unique styles", format_number(ir.unique_styles)),
        ("Merged regions imported", format_number(ir.merges_imported)),
        (
            "Conditional format rules imported",
            format_number(ir.cond_formats_imported),
        ),
        (
            "Validation rules imported",
            format_number(ir.validations_imported),
        ),
        (
            "Shared formula groups",
            format_number(ir.shared_formula_groups),
        ),
        (
            "Formulas recovered from XML",
            format_number(ir.formula_cells_without_values),
        ),
        (
            "Values recovered from XML",
            format_number(ir.value_cells_backfilled),
        ),
        (
            "Date system",
            if ir.is_1904_system {
                "1904".into()
            } else {
                "1900".into()
            },
        ),
    ]
}

/// Copy includes collapsed diagnostics and all collected examples, not just visible rows.
fn report_text(ir: &ImportResult, filename: &str) -> String {
    let notes = compatibility_notes(ir);
    let mut lines = vec![
        "VisiGrid Import Report".into(),
        format!("File: {filename}"),
        format!(
            "Imported {} sheets, {} cells, {} formulas in {} ms",
            ir.sheets_imported, ir.cells_imported, ir.formulas_imported, ir.import_duration_ms
        ),
        outcome_detail(ir, !notes.is_empty()),
        "\nCompatibility & import warnings".into(),
    ];
    if notes.is_empty() {
        lines.push("None reported".into());
    } else {
        lines.extend(notes);
    }
    lines.extend(
        ir.parse_error_samples
            .iter()
            .map(|s| format!("Parse example: {s}")),
    );
    lines.push("\nCalculation results at import".into());
    lines.extend(calculation_notes(ir));
    if ir.recalc_errors > 0 || ir.recalc_circular > 0 {
        lines.push(RECALC_EXPLANATION.into());
    }
    for ex in &ir.recalc_error_examples {
        lines.push(format!(
            "{}!{}: {} {} {}",
            ex.sheet,
            ex.address,
            ex.kind,
            ex.error,
            ex.formula.as_deref().unwrap_or("")
        ));
    }
    lines.push("\nBy sheet".into());
    for stats in &ir.sheet_stats {
        lines.push(format!(
            "{}: {} cells, {} formulas; {}",
            stats.name,
            stats.cells_imported,
            stats.formulas_imported,
            sheet_findings(stats)
        ));
    }
    lines.push("\nImport diagnostics".into());
    lines.extend(
        diagnostics(ir)
            .into_iter()
            .map(|(label, value)| format!("{label}: {value}")),
    );
    lines.join("\n")
}

fn format_number(n: impl ToString) -> String {
    let s = n.to_string();
    let mut result = String::new();
    for (i, c) in s.chars().enumerate() {
        if i > 0 && (s.len() - i) % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    result
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
            .into_any_element(),
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
                .child(format!(
                    "Cycle values frozen: {} (Excel cached)",
                    ir.cycles_frozen
                ))
                .into_any_element(),
        );
    } else if unresolved {
        children.push(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(format!(
                    "Circular references: {} cells (#CYCLE!)",
                    cycle_count
                ))
                .into_any_element(),
        );
    }

    div()
        .flex()
        .flex_col()
        .gap_2()
        .children(children)
        .into_any_element()
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_io::xlsx::RecalcErrorExample;

    #[::core::prelude::v1::test]
    fn division_by_zero_is_a_calculation_finding_not_an_import_loss() {
        let ir = ImportResult {
            sheets_imported: 1,
            recalc_errors: 1,
            recalc_error_examples: vec![RecalcErrorExample {
                sheet: "Edge cases".into(),
                address: "B13".into(),
                kind: "error",
                error: "#DIV/0!".into(),
                formula: Some("=1/0".into()),
            }],
            ..Default::default()
        };
        assert!(compatibility_notes(&ir).is_empty());
        let text = report_text(&ir, "test.xlsx");
        assert!(text.contains("1 formula error found after recalculation"));
        assert!(text.contains("may already exist in the source workbook"));
        assert!(text.contains("Edge cases!B13: error #DIV/0! =1/0"));
        assert!(!text.contains("could not be parsed"));
    }

    #[::core::prelude::v1::test]
    fn copy_keeps_losses_sheet_findings_and_collapsed_diagnostics() {
        let mut ir = ImportResult {
            truncated: true,
            formulas_failed: 1,
            formulas_with_unknowns: 2,
            validations_skipped: 3,
            merges_dropped_overlap: 1,
            warnings: vec!["Unsupported formatting: color scale".into()],
            unsupported_format_features: vec!["color scale".into()],
            shared_formula_groups: 7,
            value_cells_backfilled: 9,
            sheet_stats: vec![SheetStats {
                name: "Monthly report".into(),
                formulas_with_errors: 1,
                formulas_with_unknowns: 2,
                truncated_rows: 10,
                ..Default::default()
            }],
            ..Default::default()
        };
        ir.unsupported_functions.insert("MISSING_FUNC".into(), 2);
        let text = report_text(&ir, "test.xlsx");
        assert!(text.contains("Some data was not loaded"));
        assert!(text.contains("1 formula could not be parsed"));
        assert!(text.contains("MISSING_FUNC (2)"));
        assert!(text.contains("3 validation rules skipped"));
        assert!(text.contains("1 merged region dropped"));
        assert_eq!(
            text.matches("Unsupported formatting: color scale").count(),
            1
        );
        assert!(text.contains("Monthly report:"));
        assert!(text.contains("Truncated: 10 rows, 0 columns"));
        assert!(text.contains("Shared formula groups: 7"));
        assert!(text.contains("Values recovered from XML: 9"));
    }

    #[::core::prelude::v1::test]
    fn frozen_cycles_remain_visible_with_their_recalculation_limitation() {
        let ir = ImportResult {
            freeze_applied: true,
            cycles_frozen: 3,
            cycles_no_cached: 1,
            ..Default::default()
        };
        let text = report_text(&ir, "cycles.xlsx");
        assert!(text.contains("3 frozen to Excel's cached values"));
        assert!(text.contains("will not recalculate"));
        assert!(text.contains("1 cycle cell had no cached value"));
        assert!(!text.contains("No issues reported"));
    }
}
