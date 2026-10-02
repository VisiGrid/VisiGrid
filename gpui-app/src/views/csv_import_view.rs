//! CSV import: the banner after opening a file, and the import settings dialog.
//!
//! State and actions live in `csv_import_ui.rs`.

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::formula::eval::Value;
use visigrid_io::csv::{ColumnDecision, ColumnRule, DateOrder, Resolved};

use crate::app::Spreadsheet;
use crate::csv_import_ui::{CsvDialogState, FIRST_COLUMN_FOCUS, REMEMBER_FOCUS, SETTINGS};
use crate::theme::TokenKey;
use crate::ui::{dialog_header_with_subtitle, modal_overlay, Button, DialogFrame};

#[derive(Clone, Copy)]
struct Colors {
    panel: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
    accent: Hsla,
    inverse: Hsla,
    warn: Hsla,
    error: Hsla,
    grid_bg: Hsla,
}

impl Colors {
    fn new(app: &Spreadsheet) -> Self {
        Self {
            panel: app.token(TokenKey::PanelBg),
            border: app.token(TokenKey::PanelBorder),
            text: app.token(TokenKey::TextPrimary),
            muted: app.token(TokenKey::TextMuted),
            accent: app.token(TokenKey::Accent),
            inverse: app.token(TokenKey::TextInverse),
            warn: app.token(TokenKey::Warn),
            error: app.token(TokenKey::Error),
            grid_bg: app.token(TokenKey::EditorBg),
        }
    }
}

fn plural(n: usize) -> &'static str {
    if n == 1 { "" } else { "s" }
}

fn thousands(n: usize) -> String {
    let s = n.to_string();
    let mut out = String::new();
    for (i, ch) in s.chars().enumerate() {
        if i > 0 && (s.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    out
}

// ============================================================================
// Banner
// ============================================================================

/// One thing the banner reports, with what can be done about it.
struct Finding {
    tone: Hsla,
    /// "!" for warnings, "i" for information.
    mark: &'static str,
    title: String,
    body: String,
    actions: Vec<(&'static str, String, FindingAction)>,
}

#[derive(Clone, Copy)]
enum FindingAction {
    ShowCells,
    ReviewFormulas,
    Settings,
    ReimportChanged,
    SaveAs,
}

/// "zip, sku and card"
fn name_list(names: &[String]) -> String {
    match names.len() {
        0 => String::new(),
        1 => names[0].clone(),
        n if n <= 4 => format!("{} and {}", names[..n - 1].join(", "), names[n - 1]),
        n => format!("{} and {} more", names[..3].join(", "), n - 3),
    }
}

fn findings(doc: &crate::csv_import_ui::CsvDocState, dirty: bool, c: &Colors) -> Vec<Finding> {
    let mut out = Vec::new();
    if let Some(change) = &doc.disk_change {
        let mut body = change.describe(!doc.options.no_header);
        if dirty {
            body.push_str(" Re-importing replaces your unsaved edits.");
        }
        out.push(Finding {
            tone: c.accent,
            mark: "!",
            title: format!("{} changed on disk", doc.file_name()),
            body,
            actions: vec![
                ("csv-reimport-changed", if dirty { "Discard edits and re-import".into() } else { "Re-import with the same settings".into() }, FindingAction::ReimportChanged),
                ("csv-review-settings", "Review settings…".into(), FindingAction::Settings),
            ],
        });
    }
    if doc.truncated() {
        out.push(Finding {
            tone: c.error,
            mark: "!",
            // Data rows, like the header line (the header row is not one)
            title: format!(
                "Only the first {} of {} rows fit",
                thousands(doc.rows_loaded - (doc.rows_in_file - doc.data_rows())),
                thousands(doc.data_rows())
            ),
            body: format!(
                "A sheet holds {} rows. The other {} are still in the file, so VisiGrid will not write over it.",
                thousands(doc.rows_loaded),
                thousands(doc.rows_in_file - doc.rows_loaded)
            ),
            actions: vec![("csv-save-as", "Save as a new file…".into(), FindingAction::SaveAs)],
        });
    }
    let formulas = doc.formula_cells.len();
    if formulas > 0 && doc.formulas_changed {
        out.push(Finding {
            tone: c.warn,
            mark: "!",
            title: format!("The formulas in {} changed since you approved them", doc.file_name()),
            body: format!("{formulas} cell{} starting with = {} left as text. Review them before evaluating again.", plural(formulas), if formulas == 1 { "was" } else { "were" }),
            actions: vec![
                ("csv-show-cells", "Show cells".into(), FindingAction::ShowCells),
                ("csv-review", "Review and evaluate…".into(), FindingAction::ReviewFormulas),
            ],
        });
    } else if formulas > 0 {
        out.push(Finding {
            tone: c.warn,
            mark: "!",
            title: format!("{formulas} cell{} start{} with = and {} left as text", plural(formulas), if formulas == 1 { "s" } else { "" }, if formulas == 1 { "was" } else { "were" }),
            body: "Formulas in a CSV can send your data elsewhere. Evaluate them only if you trust where the file came from.".into(),
            actions: vec![
                ("csv-show-cells", "Show cells".into(), FindingAction::ShowCells),
                ("csv-review", "Review and evaluate…".into(), FindingAction::ReviewFormulas),
            ],
        });
    }
    if doc.unreadable > 0 {
        out.push(Finding {
            tone: c.warn,
            mark: "!",
            title: format!("{} value{} did not fit the chosen type", thousands(doc.unreadable), plural(doc.unreadable)),
            body: format!("In {}. They stayed as text.", name_list(&doc.unreadable_columns)),
            actions: vec![("csv-types-2", "Change column types…".into(), FindingAction::Settings)],
        });
    }
    if doc.kept_as_text > 0 {
        out.push(Finding {
            tone: c.accent,
            mark: "i",
            title: format!("{} value{} kept as text", thousands(doc.kept_as_text), plural(doc.kept_as_text)),
            body: format!("So leading zeros and long IDs stay exact in {}.", name_list(&doc.text_columns)),
            actions: vec![("csv-types", "Change column types…".into(), FindingAction::Settings)],
        });
    }
    if doc.used_saved_settings {
        out.push(Finding {
            tone: c.accent,
            mark: "i",
            title: "Opened with your saved settings".into(),
            body: "You chose these for files with the same column names.".into(),
            actions: vec![("csv-saved", "Import settings…".into(), FindingAction::Settings)],
        });
    }
    out
}

/// Top-centre banner after a CSV import that needs a look: one row per
/// finding, warnings first, each with its own action. A clean import shows
/// nothing but the status line.
pub(crate) fn render_csv_banner(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let c = Colors::new(app);
    let Some(doc) = app.current_csv() else {
        return div().into_any_element();
    };
    let card = div()
        .id("csv-banner")
        .w(px(560.0))
        .bg(c.panel)
        .border_1()
        .rounded_md()
        .shadow_lg()
        .overflow_hidden()
        .flex()
        .flex_col()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation());

    let card = if doc.reviewing_formulas {
        render_formula_review(app, doc, &c, card, cx)
    } else {
        render_findings(doc, app.is_dirty(), &c, card, cx)
    };

    div()
        .absolute()
        .top_0()
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .pt_2()
        .child(card)
        .into_any_element()
}

fn dismiss_button(id: &'static str, c: &Colors, cx: &mut Context<Spreadsheet>, close_review: bool) -> impl IntoElement {
    let (muted, text) = (c.muted, c.text);
    div()
        .id(id)
        .px_2()
        .text_size(px(13.0))
        .text_color(muted)
        .cursor_pointer()
        .hover(move |s| s.text_color(text))
        .child("✕")
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            if close_review {
                this.csv_review_formulas(false, cx);
            } else {
                this.dismiss_csv_banner(cx);
            }
            cx.stop_propagation();
        }))
}

fn render_findings(
    doc: &crate::csv_import_ui::CsvDocState,
    dirty: bool,
    c: &Colors,
    card: Stateful<Div>,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let list = findings(doc, dirty, c);
    let checks = list.iter().filter(|f| f.mark == "!").count();
    let detail = match checks {
        0 => format!("{} rows", thousands(doc.data_rows())),
        n => format!("{} rows · {n} thing{} to check", thousands(doc.data_rows()), plural(n)),
    };

    let mut card = card.border_color(c.border).child(
        div()
            .flex()
            .items_center()
            .gap_2()
            .px(px(14.0))
            .py(px(9.0))
            .border_b_1()
            .border_color(c.border.opacity(0.6))
            .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child(format!("Opened {}", doc.file_name())))
            .child(div().text_size(px(12.0)).text_color(c.muted).child(detail))
            .child(div().flex_1())
            .child(dismiss_button("csv-banner-dismiss", c, cx, false)),
    );

    for f in list {
        let warning = f.mark == "!";
        let has_actions = !f.actions.is_empty();
        let mut actions = div().flex().gap(px(14.0)).mt(px(4.0));
        for (id, label, action) in f.actions {
            let color = if matches!(action, FindingAction::ReviewFormulas) { c.warn } else { c.accent };
            actions = actions.child(
                div()
                    .id(id)
                    .text_size(px(12.0))
                    .text_color(color)
                    .cursor_pointer()
                    .hover(|s| s.underline())
                    .child(label)
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        match action {
                            FindingAction::ShowCells => this.csv_show_formula_cells(cx),
                            FindingAction::ReviewFormulas => this.csv_review_formulas(true, cx),
                            FindingAction::Settings => this.show_csv_import_dialog(cx),
                            FindingAction::ReimportChanged => this.csv_reimport_changed(cx),
                            FindingAction::SaveAs => this.save_as(cx),
                        }
                        cx.stop_propagation();
                    })),
            );
        }
        card = card.child(
            div()
                .flex()
                .gap(px(10.0))
                .px(px(14.0))
                .py(px(10.0))
                .when(warning, |d| d.bg(f.tone.opacity(0.06)).border_l(px(3.0)).border_color(f.tone).pl(px(11.0)))
                .child(
                    div()
                        .mt(px(1.0))
                        .size(px(18.0))
                        .flex_shrink_0()
                        .flex()
                        .items_center()
                        .justify_center()
                        .rounded(px(4.0))
                        .bg(f.tone.opacity(0.16))
                        .text_color(f.tone)
                        .text_size(px(11.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .child(f.mark),
                )
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .flex()
                        .flex_col()
                        .gap(px(3.0))
                        .child(div().text_size(px(12.0)).font_weight(FontWeight::SEMIBOLD).text_color(if warning { f.tone } else { c.text }).child(f.title))
                        .child(div().text_size(px(11.0)).line_height(px(16.0)).text_color(c.muted).child(f.body))
                        .when(has_actions, |d| d.child(actions)),
                ),
        );
    }
    card
}

/// The review step's confirm button. With unsaved edits it names the cost.
fn evaluate_label(total: usize, dirty: bool) -> String {
    if dirty {
        format!("Discard edits and evaluate {total}")
    } else {
        format!("Evaluate {total} formula{}", plural(total))
    }
}

/// The step before evaluating: the formulas themselves, with any that reach
/// outside the sheet called out. "Keep as text" comes first.
fn render_formula_review(
    app: &Spreadsheet,
    doc: &crate::csv_import_ui::CsvDocState,
    c: &Colors,
    card: Stateful<Div>,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    const SHOWN: usize = 6;
    let total = doc.formula_cells.len();
    let sheet = app.sheet(cx);
    let formulas: Vec<(String, String)> = doc
        .formula_cells
        .iter()
        .map(|&(r, col)| (app.cell_ref_at(r, col), sheet.get_raw(r, col)))
        .collect();
    let outside = formulas.iter().filter(|(_, f)| crate::csv_import_ui::reaches_outside(f)).count();
    // Evaluating re-imports the file, which replaces the sheet and its undo
    // history: say so when there is something to lose, as the dialog does
    let dirty = app.is_dirty();

    let mut list = div()
        .rounded(px(4.0))
        .border_1()
        .border_color(c.border)
        .bg(c.grid_bg)
        .flex()
        .flex_col();
    for (i, (cell, formula)) in formulas.iter().take(SHOWN).enumerate() {
        let flagged = crate::csv_import_ui::reaches_outside(formula);
        list = list.child(
            div()
                .flex()
                .gap(px(12.0))
                .px(px(10.0))
                .py(px(6.0))
                .when(i > 0, |d| d.border_t_1().border_color(c.border.opacity(0.5)))
                .font_family("IBM Plex Mono")
                .text_size(px(11.0))
                .child(div().w(px(36.0)).flex_shrink_0().text_color(c.muted).child(cell.clone()))
                .child(div().flex_1().min_w_0().truncate().text_color(if flagged { c.warn } else { c.text }).child(formula.clone())),
        );
    }
    if total > SHOWN {
        list = list.child(
            div().px(px(10.0)).py(px(6.0)).border_t_1().border_color(c.border.opacity(0.5))
                .text_size(px(11.0)).text_color(c.muted).child(format!("and {} more", total - SHOWN)),
        );
    }

    card.border_color(c.warn.opacity(0.55))
        .px(px(14.0))
        .py(px(12.0))
        .gap(px(8.0))
        .child(
            div()
                .flex()
                .items_center()
                .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text)
                    .child(format!("Evaluate {total} formula{} from {}?", plural(total), doc.file_name())))
                .child(div().flex_1())
                .child(dismiss_button("csv-review-close", c, cx, true)),
        )
        .child(list)
        .child(div().text_size(px(11.0)).line_height(px(16.0)).text_color(if outside > 0 { c.warn } else { c.muted }).child(match outside {
            0 => "None of these reach outside the sheet.".to_string(),
            n => format!("{n} formula{} reach{} outside the sheet (a link, a web request or an import).", plural(n), if n == 1 { "es" } else { "" }),
        }))
        .when(dirty, |d| {
            d.child(div().text_size(px(11.0)).line_height(px(16.0)).text_color(c.error)
                .child("Evaluating re-imports the file: your unsaved edits to this sheet will be lost."))
        })
        .child(
            div()
                .flex()
                .justify_end()
                .gap_2()
                .child(
                    Button::new("csv-keep-text", "Keep as text")
                        .secondary(c.accent, c.text)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.csv_review_formulas(false, cx);
                            cx.stop_propagation();
                        })),
                )
                .child(
                    Button::new("csv-evaluate", evaluate_label(total, dirty))
                        .secondary(c.warn.opacity(0.7), c.warn)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.csv_evaluate_formulas(cx);
                            cx.stop_propagation();
                        })),
                ),
        )
}

// ============================================================================
// Import settings dialog
// ============================================================================

const COL_W: f32 = 150.0;
const HEAD_H: f32 = 104.0;
const ROW_H: f32 = 24.0;

/// Text on a column's type pill.
fn pill_label(d: &ColumnDecision) -> String {
    match d.rule {
        ColumnRule::Auto => format!("Auto · {}", d.resolved.label()),
        ColumnRule::Date(o) => format!("Date {}", o.label()),
        other => other.label(),
    }
}

fn order_name(o: DateOrder) -> &'static str {
    match o {
        DateOrder::Ymd => "YYYY-MM-DD",
        DateOrder::Dmy => "day/month/year",
        DateOrder::Mdy => "month/day/year",
    }
}

/// The note under a column's pill, and whether it warns.
fn column_note(d: &ColumnDecision) -> (String, bool) {
    match d.rule {
        ColumnRule::Skip => ("Not imported".into(), false),
        ColumnRule::Auto => match d.reason {
            Some(reason) => (format!("{}: {reason}", d.resolved.label()), false),
            None if d.resolved == Resolved::Empty => ("Empty in the preview".into(), false),
            None => (d.resolved.label().to_string(), false),
        },
        ColumnRule::Number if d.changed_if_number > 0 => (
            format!("Number would change {} value{}", d.changed_if_number, plural(d.changed_if_number)),
            true,
        ),
        ColumnRule::Date(o) if d.unreadable > 0 => (
            format!("{} value{} not {}; kept as text", d.unreadable, plural(d.unreadable), order_name(o)),
            true,
        ),
        ColumnRule::Date(o) => (format!("Read as {}", order_name(o)), false),
        _ => ("As chosen".into(), false),
    }
}

/// Values in the preview the chosen types would change or could not read.
fn changed_count(cols: &[ColumnDecision]) -> usize {
    cols.iter()
        .map(|d| match d.rule {
            ColumnRule::Number => d.changed_if_number,
            ColumnRule::Date(_) => d.unreadable,
            _ => 0,
        })
        .sum()
}

pub(crate) fn render_csv_import_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(state) = &app.csv_dialog else {
        return div().into_any_element();
    };
    let c = Colors::new(app);
    let width = (f32::from(app.window_size.width) - 48.0).clamp(640.0, 1180.0);
    let height = (f32::from(app.window_size.height) - 48.0).clamp(420.0, 760.0);
    let body_h = height - 128.0;

    let file = state.path.file_name().and_then(|n| n.to_str()).unwrap_or("file").to_string();
    let subtitle = match (app.current_csv(), &state.preview) {
        (Some(doc), Ok(p)) => format!(
            "{file} · {} rows · {} column{}",
            thousands(doc.data_rows()),
            p.columns.len(),
            plural(p.columns.len())
        ),
        _ => file.clone(),
    };

    let body = div()
        .flex()
        .h(px(body_h))
        .min_h_0()
        .child(render_settings(state, &c, cx))
        .child(render_preview(state, &c, body_h, cx));

    let changed = state.preview.as_ref().map_or(0, |p| changed_count(&p.columns));
    let text_cols = state.preview.as_ref().map_or(0, |p| {
        p.columns.iter().filter(|d| matches!(d.resolved, Resolved::Text | Resolved::Mixed)).count()
    });
    let summary = match &state.preview {
        Err(e) => format!("Cannot read the file with these settings: {e}"),
        Ok(_) if changed > 0 => format!("{changed} value{} in the preview would change", plural(changed)),
        Ok(_) => format!("No values change · {text_cols} column{} as text", plural(text_cols)),
    };
    let dirty = app.is_dirty();

    let footer = div()
        .flex()
        .items_center()
        .gap_3()
        .child(
            div()
                .flex()
                .flex_col()
                .gap(px(2.0))
                .child(div().text_size(px(12.0)).text_color(if changed > 0 || state.preview.is_err() { c.warn } else { c.muted }).child(summary))
                .when(dirty, |d| {
                    d.child(div().text_size(px(11.0)).text_color(c.warn).child("Re-importing replaces the sheet; your unsaved edits will be lost."))
                }),
        )
        .child(div().flex_1())
        .child(div().text_size(px(11.0)).text_color(c.muted).child("Tab moves · Space changes · Enter re-imports · Esc cancels"))
        .child(
            Button::new("csv-cancel", "Cancel")
                .secondary(c.border, c.text)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.close_csv_import_dialog(cx);
                    cx.stop_propagation();
                })),
        )
        .child(
            Button::new("csv-reimport", if dirty { "Discard edits and re-import" } else { "Re-import" })
                .disabled(state.preview.is_err())
                .primary(c.accent, c.inverse)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    let ok = this.csv_dialog.as_ref().is_some_and(|s| s.preview.is_ok());
                    if ok {
                        this.csv_dialog_apply(cx);
                    }
                    cx.stop_propagation();
                })),
        );

    modal_overlay(
        "csv-import-dialog",
        |this, cx| this.close_csv_import_dialog(cx),
        DialogFrame::new(body, c.panel, c.border)
            .width(px(width))
            .max_height(px(height))
            .header(dialog_header_with_subtitle("Import settings", subtitle, c.text, c.muted))
            .footer(footer),
        cx,
    )
    .into_any_element()
}

fn render_settings(state: &CsvDialogState, c: &Colors, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let c = *c;
    let mut col = div()
        .w(px(284.0))
        .flex_shrink_0()
        .h_full()
        .pr_4()
        .mr_4()
        .border_r_1()
        .border_color(c.border)
        .flex()
        .flex_col()
        .gap_3()
        .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child("File"));

    for (index, label) in SETTINGS.iter().enumerate() {
        let (value, hint) = state.setting(index);
        col = col.child(
            div()
                .flex()
                .flex_col()
                .gap(px(4.0))
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .gap_2()
                        .child(div().text_size(px(12.0)).text_color(c.text).child(*label))
                        .child(
                            Button::new(ElementId::Name(format!("csv-setting-{index}").into()), format!("{value} ›"))
                                .secondary(if state.focus == index { c.accent } else { c.border }, c.text)
                                .w(px(154.0))
                                .px_2()
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                    this.csv_dialog_cycle(index, cx);
                                    cx.stop_propagation();
                                })),
                        ),
                )
                .when(!hint.is_empty(), |d| d.child(div().text_size(px(11.0)).text_color(c.muted).child(hint))),
        );
    }

    let remember_focused = state.focus == REMEMBER_FOCUS;
    col = col.child(
        div()
            .id("csv-remember")
            .mt_1()
            .flex()
            .items_start()
            .gap_2()
            .cursor_pointer()
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                this.csv_dialog_cycle(REMEMBER_FOCUS, cx);
                cx.stop_propagation();
            }))
            .child(
                div()
                    .mt(px(1.0))
                    .size(px(14.0))
                    .flex_shrink_0()
                    .flex()
                    .items_center()
                    .justify_center()
                    .rounded(px(3.0))
                    .border_1()
                    .border_color(if remember_focused { c.accent } else { c.border })
                    .when(state.remember, |d| d.bg(c.accent).text_color(c.inverse))
                    .text_size(px(10.0))
                    .child(if state.remember { "✓" } else { "" }),
            )
            .child(
                div()
                    .flex_1()
                    .min_w_0()
                    .flex()
                    .flex_col()
                    .gap(px(2.0))
                    .child(div().text_size(px(12.0)).text_color(c.text).child("Remember for files with these column names"))
                    .child(div().text_size(px(11.0)).text_color(c.muted).child("They open the same way next time. Formulas are still asked about each time.")),
            ),
    );

    col.child(div().flex_1()).child(
        div()
            .p(px(10.0))
            .rounded_md()
            .border_1()
            .border_color(c.border)
            .bg(c.grid_bg)
            .flex()
            .flex_col()
            .gap(px(6.0))
            .child(div().text_size(px(11.0)).text_color(c.muted).child("Same import from the command line"))
            .child(div().font_family("IBM Plex Mono").text_size(px(11.0)).text_color(c.text).child(state.cli_line())),
    )
}

fn render_preview(state: &CsvDialogState, c: &Colors, body_h: f32, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let c = *c;
    let heading = div()
        .flex()
        .items_center()
        .gap_3()
        .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child("Preview"))
        .child(div().text_size(px(11.0)).text_color(c.muted).child("Click a column's type to change it. Auto never guesses dates."));

    let preview = match &state.preview {
        Ok(p) => p,
        Err(e) => {
            return div()
                .flex_1()
                .min_w_0()
                .flex()
                .flex_col()
                .gap_2()
                .child(heading)
                .child(div().text_size(px(12.0)).text_color(c.warn).child(e.clone()));
        }
    };

    let first_row = usize::from(!state.options.no_header);
    let grid_h = body_h - 30.0;
    let rows_fit = (((grid_h - HEAD_H) / ROW_H).floor() as usize).max(1);
    let last_row = (first_row + rows_fit).min(preview.sample.len());

    let mut columns = div()
        .id("csv-preview-columns")
        .flex()
        .size_full()
        .overflow_x_scroll()
        .track_scroll(&state.columns_scroll);

    let mut skipped_before = 0usize;
    for (ci, d) in preview.columns.iter().enumerate() {
        let focus = FIRST_COLUMN_FOCUS + ci;
        let focused = state.focus == focus;
        let dest = (d.rule != ColumnRule::Skip).then(|| d.source_index - skipped_before);
        if d.rule == ColumnRule::Skip {
            skipped_before += 1;
        }
        let (note, warns) = column_note(d);
        let chosen = d.rule != ColumnRule::Auto;
        let (pill_border, pill_bg) = if warns {
            (c.warn.opacity(0.6), c.warn.opacity(0.10))
        } else if chosen {
            (c.accent.opacity(0.6), c.accent.opacity(0.12))
        } else {
            (c.border, c.panel)
        };

        let mut column = div()
            .w(px(COL_W))
            .flex_shrink_0()
            .flex()
            .flex_col()
            .border_r_1()
            .border_color(c.border.opacity(0.6))
            .when(d.rule == ColumnRule::Skip, |d| d.opacity(0.55))
            .child(
                div()
                    .h(px(HEAD_H))
                    .p(px(8.0))
                    .flex()
                    .flex_col()
                    .gap(px(5.0))
                    .bg(c.panel)
                    .border_b_1()
                    .border_color(c.border)
                    .child(div().flex_shrink_0().font_family("IBM Plex Mono").text_size(px(12.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).truncate().child(d.name.clone()))
                    .child(
                        div()
                            .id(ElementId::Name(format!("csv-pill-{ci}").into()))
                            .flex_shrink_0()
                            .self_start()
                            .px(px(8.0))
                            .py(px(2.0))
                            .rounded(px(10.0))
                            .border_1()
                            .border_color(if focused { c.accent } else { pill_border })
                            .bg(pill_bg)
                            .text_size(px(11.0))
                            .text_color(c.text)
                            .cursor_pointer()
                            .child(format!("{} ›", pill_label(d)))
                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                this.csv_dialog_cycle(focus, cx);
                                cx.stop_propagation();
                            })),
                    )
                    .child(div().text_size(px(11.0)).line_height(px(14.0)).max_h(px(28.0)).overflow_hidden().text_color(if warns { c.warn } else { c.muted }).child(note)),
            );

        for row in first_row..last_row {
            let raw = preview.sample[row].get(d.source_index).cloned().unwrap_or_default();
            let (shown, color, bg, italic, right) = match dest {
                None => (raw.clone(), c.muted, None, true, false),
                Some(col) => {
                    let value = preview.sheet.get_computed_value(row, col);
                    let shown = preview.sheet.get_formatted_display(row, col);
                    let is_number = matches!(value, Value::Number(_));
                    let typed = matches!(d.rule, ColumnRule::Number | ColumnRule::Date(_));
                    if typed && !raw.is_empty() && !is_number {
                        (shown, c.error, Some(c.error.opacity(0.08)), false, false)
                    } else if d.rule == ColumnRule::Number && is_number && shown != raw {
                        (shown, c.warn, Some(c.warn.opacity(0.08)), false, true)
                    } else if raw.starts_with('=') && !state.options.evaluate_formulas {
                        (shown, c.muted, None, true, false)
                    } else {
                        (shown, c.text, None, false, is_number)
                    }
                }
            };
            column = column.child(
                div()
                    .h(px(ROW_H))
                    .px(px(8.0))
                    .flex()
                    .items_center()
                    .when(right, |d| d.justify_end())
                    .border_b_1()
                    .border_color(c.border.opacity(0.4))
                    .when_some(bg, |d, bg| d.bg(bg))
                    .font_family("IBM Plex Mono")
                    .text_size(px(12.0))
                    .text_color(color)
                    .when(italic, |d| d.italic())
                    .overflow_hidden()
                    .whitespace_nowrap()
                    .child(shown),
            );
        }
        columns = columns.child(column);
    }

    div()
        .flex_1()
        .min_w_0()
        .flex()
        .flex_col()
        .gap(px(10.0))
        .child(heading)
        .child(
            div()
                .h(px(grid_h))
                .rounded_md()
                .border_1()
                .border_color(c.border)
                .bg(c.grid_bg)
                .overflow_hidden()
                .child(columns),
        )
}

#[cfg(test)]
mod tests {
    use super::evaluate_label;

    #[test]
    fn evaluate_button_names_the_cost_of_unsaved_edits() {
        assert_eq!(evaluate_label(2, false), "Evaluate 2 formulas");
        assert_eq!(evaluate_label(1, false), "Evaluate 1 formula");
        assert_eq!(evaluate_label(2, true), "Discard edits and evaluate 2");
    }
}
