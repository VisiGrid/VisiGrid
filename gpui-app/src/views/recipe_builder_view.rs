//! The recipe builder: source settings, steps, the selected step's settings
//! and a preview. State and keys live in `recipe_builder.rs`.

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_io::csv_import::ColumnRule;
use visigrid_io::recipe::Step;

use crate::app::Spreadsheet;
use crate::recipe_builder::{
    cli_line, filter_op_label, missing_label, on_error_label, rule_label, step_kind, step_missing, type_label,
    EditorRow, Pane, RecipeBuilder, TotalPart, ADD_KINDS,
};
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
    crate::duckdb_import::number(n as u64)
}

fn with_builder(f: impl Fn(&mut RecipeBuilder) + 'static) -> impl Fn(&mut Spreadsheet, &MouseDownEvent, &mut Window, &mut Context<Spreadsheet>) + 'static {
    move |this, _, _, cx| {
        if let Some(b) = this.recipe_builder.as_mut() {
            b.error = None;
            f(b);
        }
        cx.stop_propagation();
        cx.notify();
    }
}

fn section_title(label: &str, c: &Colors, focused: bool) -> impl IntoElement {
    div()
        .text_size(px(12.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(if focused { c.accent } else { c.text })
        .child(label.to_string())
}

pub(crate) fn render_recipe_builder(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(b) = &app.recipe_builder else {
        return div().into_any_element();
    };
    let c = Colors::new(app);
    let width = (f32::from(app.window_size.width) - 48.0).clamp(760.0, 1400.0);
    let height = (f32::from(app.window_size.height) - 48.0).clamp(480.0, 860.0);
    let body_h = height - 150.0;

    let file = b.source_path.file_name().and_then(|n| n.to_str()).unwrap_or("file").to_string();
    let recipe_name = b
        .recipe_path
        .as_ref()
        .and_then(|p| p.file_name())
        .and_then(|n| n.to_str())
        .map(str::to_string)
        .unwrap_or_else(|| "not saved yet".into());
    let subtitle = format!("{file} · {recipe_name}{}", if b.dirty && b.recipe_path.is_some() { " · unsaved changes" } else { "" });

    let body = div()
        .flex()
        .h(px(body_h))
        .min_h_0()
        .child(render_source(b, &c, cx))
        .child(render_steps(b, &c, cx))
        .child(render_right(b, &c, body_h, cx));

    modal_overlay(
        "recipe-builder",
        |this, cx| this.close_recipe_builder(cx),
        DialogFrame::new(body, c.panel, c.border)
            .width(px(width))
            .max_height(px(height))
            .header(dialog_header_with_subtitle("Import recipe", subtitle, c.text, c.muted))
            .footer(render_footer(app, b, &c, cx)),
        cx,
    )
    .into_any_element()
}

// ----------------------------------------------------------------------------
// Source
// ----------------------------------------------------------------------------

fn render_source(b: &RecipeBuilder, c: &Colors, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let focused = b.pane == Pane::Source;
    let mut col = div()
        .w(px(250.0))
        .flex_shrink_0()
        .h_full()
        .p(px(14.0))
        .border_r_1()
        .border_color(c.border)
        .flex()
        .flex_col()
        .gap(px(12.0))
        .child(section_title(&format!("Source · {}", b.recipe.source.label()), c, focused));
    for (i, label) in b.source_rows().iter().enumerate() {
        let (value, hint) = b.source_value(i);
        let is_file = b.source_row_is_file(i);
        let active = focused && b.source_focus == i;
        let (accent, border, text) = (c.accent, c.border, c.text);
        col = col.child(
            div()
                .flex()
                .flex_col()
                .gap(px(3.0))
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .gap_2()
                        .child(div().text_size(px(12.0)).text_color(c.text).child(label.to_string()))
                        .child(
                            div()
                                .id(("recipe-source", i))
                                .max_w(px(150.0))
                                .overflow_hidden()
                                .text_ellipsis()
                                .whitespace_nowrap()
                                .px(px(8.0))
                                .py(px(3.0))
                                .rounded(px(4.0))
                                .border_1()
                                .border_color(if active { accent } else { border })
                                .text_size(px(12.0))
                                .text_color(text)
                                .cursor_pointer()
                                .child(if is_file { format!("{value} …") } else { format!("{value} ›") })
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                    cx.stop_propagation();
                                    if is_file {
                                        this.recipe_builder_choose_file(cx);
                                        return;
                                    }
                                    if let Some(b) = this.recipe_builder.as_mut() {
                                        b.pane = Pane::Source;
                                        b.source_focus = i;
                                        b.change_source(i, false);
                                    }
                                    cx.notify();
                                })),
                        ),
                )
                .when(!hint.is_empty(), |d| {
                    d.child(div().text_size(px(11.0)).line_height(px(15.0)).text_color(c.muted).child(hint))
                }),
        );
    }
    // The columns the file has, saved with the recipe for next month's check
    let mut list = div().flex().flex_col().gap(px(3.0));
    let saved = b.recipe.source.columns();
    for name in &b.file_columns {
        let new = !saved.is_empty() && !saved.iter().any(|s| s.eq_ignore_ascii_case(name));
        list = list.child(
            div()
                .text_size(px(12.0))
                .font_family("IBM Plex Mono")
                .text_color(if new { c.warn } else { c.text })
                .child(if new { format!("{name}  (new)") } else { name.clone() }),
        );
    }
    for name in saved.iter().filter(|s| !b.file_columns.iter().any(|f| f.eq_ignore_ascii_case(s))) {
        list = list.child(div().text_size(px(12.0)).font_family("IBM Plex Mono").text_color(c.error).child(format!("{name}  (missing)")));
    }
    col.child(
        div()
            .id("recipe-file-columns")
            .flex_1()
            .min_h_0()
            .overflow_y_scroll()
            .pt(px(10.0))
            .border_t_1()
            .border_color(c.border.opacity(0.6))
            .flex()
            .flex_col()
            .gap(px(6.0))
            .child(div().text_size(px(12.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child(if b.recipe.source.is_remote() { "Columns in the report" } else { "Columns in the file" }))
            .child(list)
            .child(
                div()
                    .text_size(px(11.0))
                    .line_height(px(15.0))
                    .text_color(c.muted)
                    .child(if b.recipe.source.is_remote() {
                        "Saved with the recipe. Each refresh is checked against them before anything runs."
                    } else {
                        "Saved with the recipe. Next month's file is checked against them before anything runs."
                    }),
            ),
    )
}

// ----------------------------------------------------------------------------
// Steps
// ----------------------------------------------------------------------------

fn render_steps(b: &RecipeBuilder, c: &Colors, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let focused = b.pane == Pane::Steps;
    let report = b.full.as_ref();
    // Direct children of the scroll container, so scroll_to_item can keep
    // the selected step and the add menu in view
    let mut list: Vec<AnyElement> = Vec::new();

    // The source itself, as "step 0": preview straight from the file
    let on = b.selected.is_none();
    list.push(step_row(
        "recipe-step-source".into(),
        "0".into(),
        match b.snapshot.as_ref().map_or(1, |s| s.file_count()) {
            1 if b.recipe.source.is_remote() => "Read the report".into(),
            1 => "Read the file".into(),
            n => format!("Read {n} files and append them"),
        },
        report.map(|r| format!("{} rows · {} columns", thousands(r.source_rows), b.file_columns.len())).unwrap_or_default(),
        // Appended files missing a column: worth seeing, not a failure
        report.filter(|r| !r.warnings.is_empty()).map(|r| r.warnings.join(" · ")),
        false,
        on,
        focused,
        c,
        cx,
        None,
    ).into_any_element());
    for (i, step) in b.recipe.steps.iter().enumerate() {
        let r = report.and_then(|r| r.steps.get(i));
        let meta = match r {
            Some(r) => format!(
                "{} → {} rows · {} ms · missing column: {}",
                thousands(r.rows_in),
                thousands(r.rows_out),
                r.millis,
                missing_label(step_missing(step)).to_lowercase()
            ),
            None => String::new(),
        };
        let failed = r.is_some_and(|r| r.failed);
        let note = r.and_then(|r| {
            if r.failed {
                let mut reasons: Vec<String> = Vec::new();
                if !r.missing.is_empty() {
                    reasons.push(format!("not in the file: {}", r.missing.join(", ")));
                }
                if let Some(n) = &r.note {
                    reasons.push(n.clone());
                }
                Some(if reasons.is_empty() { "fails the run".into() } else { reasons.join(" · ") })
            } else {
                r.note.clone()
            }
        });
        list.push(step_row(
            format!("recipe-step-{i}"),
            (i + 1).to_string(),
            step.describe(),
            meta,
            note,
            failed,
            b.selected == Some(i),
            focused,
            c,
            cx,
            Some(i),
        ).into_any_element());
    }

    let mut add = div()
        .id("recipe-add-step")
        .mt(px(4.0))
        .px(px(10.0))
        .py(px(7.0))
        .rounded(px(6.0))
        .border_1()
        .border_dashed()
        .border_color(c.border)
        .text_size(px(12.0))
        .text_color(c.accent)
        .cursor_pointer()
        .flex()
        .justify_between()
        .child("+ Add step")
        .child(div().text_color(c.muted).child("A"))
        .on_mouse_down(MouseButton::Left, cx.listener(with_builder(|b| {
            if b.add_menu.is_some() {
                b.add_menu = None;
            } else {
                b.open_add_menu();
            }
        })));
    if b.full.is_none() {
        add = add.opacity(0.5);
    }
    list.push(add.into_any_element());
    if let Some(sel) = b.add_menu {
        let mut menu = div().p(px(4.0)).rounded(px(6.0)).border_1().border_color(c.border).bg(c.grid_bg).flex().flex_col();
        for (k, (label, help)) in ADD_KINDS.iter().enumerate() {
            let on = k == sel;
            let (accent, text, muted) = (c.accent, c.text, c.muted);
            menu = menu.child(
                div()
                    .id(("recipe-add-kind", k))
                    .flex()
                    .justify_between()
                    .gap_2()
                    .px(px(8.0))
                    .py(px(5.0))
                    .rounded(px(4.0))
                    .text_size(px(12.0))
                    .cursor_pointer()
                    .when(on, |d| d.bg(accent.opacity(0.18)))
                    .child(div().text_color(text).child(format!("{}  {}", add_kind_key(k), label)))
                    .child(div().text_color(muted).child(help.to_string()))
                    .on_mouse_down(MouseButton::Left, cx.listener(with_builder(move |b| b.add_step(k)))),
            );
        }
        list.push(menu.into_any_element());
    }

    div()
        .w(px(370.0))
        .flex_shrink_0()
        .h_full()
        .border_r_1()
        .border_color(c.border)
        .flex()
        .flex_col()
        .child(
            div()
                .flex()
                .items_baseline()
                .gap_2()
                .px(px(14.0))
                .pt(px(14.0))
                .pb(px(8.0))
                .child(section_title("Steps", c, focused))
                .child(div().text_size(px(11.0)).text_color(c.muted).child("click a step to preview its result")),
        )
        .child(
            div()
                .id("recipe-steps")
                .flex_1()
                .min_h_0()
                .overflow_y_scroll()
                .track_scroll(&b.steps_scroll)
                .px(px(10.0))
                .pb(px(8.0))
                .flex()
                .flex_col()
                .gap(px(4.0))
                .children(list),
        )
        .child(
            div()
                .px(px(14.0))
                .py(px(8.0))
                .border_t_1()
                .border_color(c.border.opacity(0.6))
                .text_size(px(11.0))
                .line_height(px(15.0))
                .text_color(c.muted)
                .child("↑↓ choose · Enter edit · A add · Delete remove · Ctrl+↑↓ reorder"),
        )
}

#[allow(clippy::too_many_arguments)]
fn step_row(
    id: String,
    badge: String,
    text: String,
    meta: String,
    note: Option<String>,
    failed: bool,
    on: bool,
    pane_focused: bool,
    c: &Colors,
    cx: &mut Context<Spreadsheet>,
    index: Option<usize>,
) -> impl IntoElement {
    let (accent, error, text_c, muted, inverse) = (c.accent, c.error, c.text, c.muted, c.inverse);
    let badge_bg = if failed { error } else if on { accent } else { accent.opacity(0.18) };
    div()
        .id(SharedString::from(id))
        .flex()
        .gap(px(10.0))
        .px(px(10.0))
        .py(px(8.0))
        .rounded(px(6.0))
        .border_1()
        .border_color(if on && pane_focused { accent.opacity(0.8) } else if on { accent.opacity(0.4) } else { transparent_black() })
        .when(on, |d| d.bg(accent.opacity(0.12)))
        .cursor_pointer()
        .child(
            div()
                .size(px(20.0))
                .flex_shrink_0()
                .rounded_full()
                .bg(badge_bg)
                .text_color(if on || failed { inverse } else { text_c })
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .flex()
                .items_center()
                .justify_center()
                .child(badge),
        )
        .child(
            div()
                .flex_1()
                .min_w_0()
                .flex()
                .flex_col()
                .gap(px(3.0))
                .child(div().text_size(px(13.0)).line_height(px(17.0)).text_color(text_c).child(text))
                .when(!meta.is_empty(), |d| d.child(div().text_size(px(11.0)).text_color(muted).child(meta)))
                .when_some(note, |d, n| {
                    d.child(div().text_size(px(11.0)).text_color(if failed { error } else { accent }).child(n))
                }),
        )
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, e: &MouseDownEvent, _, cx| {
            cx.stop_propagation();
            if let Some(b) = this.recipe_builder.as_mut() {
                b.add_menu = None;
                b.select(index);
                b.pane = if index.is_some() && e.click_count >= 2 { Pane::Editor } else { Pane::Steps };
            }
            cx.notify();
        }))
}

// ----------------------------------------------------------------------------
// The selected step's settings, and the preview
// ----------------------------------------------------------------------------

fn render_right(b: &RecipeBuilder, c: &Colors, body_h: f32, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let mut right = div().flex_1().min_w_0().h_full().p(px(14.0)).flex().flex_col().gap(px(10.0));
    if let Err(e) = &b.snapshot {
        return right.child(div().text_size(px(12.0)).text_color(c.error).child(format!("Can't read {}: {e}", b.source_path.display())));
    }
    if let (Some(i), Some(step)) = (b.selected, b.step()) {
        right = right.child(render_editor(b, i, step, c, (body_h * 0.42).max(150.0), cx));
    }
    right.child(render_preview(b, c))
}

fn render_editor(b: &RecipeBuilder, index: usize, step: &Step, c: &Colors, max_h: f32, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let focused = b.pane == Pane::Editor;
    let hint = match step {
        Step::Select { .. } => "Checked columns are kept, in the order you check them.",
        Step::Remove { .. } => "Checked columns are dropped.",
        Step::Rename { .. } => "Type a new name; leave it empty to keep the name.",
        Step::Types { .. } => "Every value is checked against its type on every run.",
        Step::Trim { .. } => "None checked: every column.",
        Step::Dedupe { .. } => "None checked: rows must match in every column.",
        Step::Filter { .. } => "Rows where the condition holds are kept.",
        Step::Group { .. } => "One row per group of the checked columns, with the totals below.",
        Step::Unpivot { .. } => "Checked columns stay; every other column becomes rows, including ones added later.",
        Step::Sort { .. } => "Click a column to sort by it, again for descending. Earlier columns decide first; empty values go last.",
        Step::FillDown { .. } => "Empty cells in the checked columns take the value above.",
        Step::Replace { .. } => "None checked: every column. An empty Find in whole cells replaces empty cells.",
        Step::Split { .. } => "The new columns replace it. The last one keeps the rest, so nothing is lost.",
    };
    let mut rows = div().flex().flex_col();
    for (r, row) in b.editor_rows().iter().enumerate() {
        let on = focused && b.editor_focus == r;
        // A checkbox for keep/remove/trim/dedupe rows: Some(checked)
        let mut checkbox: Option<bool> = None;
        let (label, value, value_is_text): (String, String, bool) = match (row, step) {
            (EditorRow::Column { name, present }, Step::Rename { .. }) => {
                let new = b.row_text(row).unwrap_or_default();
                (column_label(name, *present), if new.is_empty() && !on { "—".into() } else { new }, true)
            }
            (EditorRow::Column { name, present }, Step::Types { columns, .. }) => {
                let t = columns.iter().find(|(k, _)| k.eq_ignore_ascii_case(name)).map(|(_, v)| v.clone()).unwrap_or_else(|| "auto".into());
                (column_label(name, *present), type_label(&t), false)
            }
            (EditorRow::Column { name, present }, _) => {
                checkbox = Some(b.column_checked(name));
                (column_label(name, *present), String::new(), false)
            }
            (EditorRow::FilterColumn, Step::Filter { column, .. }) => ("Column".into(), column.clone(), false),
            (EditorRow::FilterOp, Step::Filter { op, .. }) => ("Condition".into(), filter_op_label(*op).into(), false),
            (EditorRow::FilterValue, Step::Filter { value, .. }) => ("Value".into(), value.clone(), true),
            (EditorRow::OnError, Step::Types { on_error, .. }) => ("If a value doesn't fit".into(), on_error_label(*on_error).into(), false),
            (EditorRow::Missing, _) => ("If a column is missing".into(), missing_label(step_missing(step)).into(), false),
            (EditorRow::GroupBy { name, present }, Step::Group { by, .. }) => {
                checkbox = Some(by.iter().any(|c| c.eq_ignore_ascii_case(name)));
                (format!("Group by {}", column_label(name, *present)), String::new(), false)
            }
            (EditorRow::Keep { name, present }, Step::Unpivot { keep, .. }) => {
                checkbox = Some(keep.iter().any(|c| c.eq_ignore_ascii_case(name)));
                (format!("Keep {}", column_label(name, *present)), String::new(), false)
            }
            (EditorRow::Total { index, part }, Step::Group { totals, .. }) => {
                let Some(t) = totals.get(*index) else { continue };
                match part {
                    TotalPart::Func => (format!("Total {}", index + 1), t.func.label().to_string(), false),
                    TotalPart::Column => ("    of column".into(), if t.column.is_empty() { "—".into() } else { t.column.clone() }, false),
                    TotalPart::Name => ("    named".into(), t.name.clone(), true),
                    TotalPart::Remove => ("    Remove this total".into(), "Remove".into(), false),
                }
            }
            (EditorRow::AddTotal, _) => ("+ Add a total".into(), "Add".into(), false),
            (EditorRow::NamesTo, Step::Unpivot { names_to, .. }) => ("Column names go in".into(), names_to.clone(), true),
            (EditorRow::ValuesTo, Step::Unpivot { values_to, .. }) => ("Their values go in".into(), values_to.clone(), true),
            (EditorRow::SortBy { name, present }, Step::Sort { by, .. }) => {
                let place = by.iter().position(|k| k.column.eq_ignore_ascii_case(name));
                let value = match place {
                    None => "—".to_string(),
                    Some(i) => format!("{} · {}", i + 1, if by[i].descending { "Descending" } else { "Ascending" }),
                };
                (column_label(name, *present), value, false)
            }
            (EditorRow::ReplaceFind, Step::Replace { find, .. }) => ("Find".into(), find.clone(), true),
            (EditorRow::ReplaceWith, Step::Replace { with, .. }) => ("Replace with".into(), with.clone(), true),
            (EditorRow::ReplacePart, Step::Replace { part, .. }) => (
                "Match".into(),
                if *part { "Text inside cells".into() } else { "Whole cell".into() },
                false,
            ),
            (EditorRow::ReplaceCase, Step::Replace { match_case, .. }) => (
                "Case".into(),
                if *match_case { "Must match".into() } else { "Ignored".into() },
                false,
            ),
            (EditorRow::SplitColumn, Step::Split { column, .. }) => ("Column".into(), column.clone(), false),
            (EditorRow::SplitBy, Step::Split { by, .. }) => ("At each".into(), by.clone(), true),
            (EditorRow::SplitInto { index }, Step::Split { into, .. }) => {
                (format!("New column {}", index + 1), into.get(*index).cloned().unwrap_or_default(), true)
            }
            (EditorRow::AddSplitPiece, _) => ("+ Add a column".into(), "Add".into(), false),
            (EditorRow::DropEmpty, Step::Unpivot { drop_empty, .. }) => (
                "Empty values".into(),
                if *drop_empty { "Leave out".into() } else { "Keep as empty rows".into() },
                false,
            ),
            _ => continue,
        };
        let (accent, border, text, muted) = (c.accent, c.border, c.text, c.muted);
        let caret = on && value_is_text;
        rows = rows.child(
            div()
                .id(("recipe-editor-row", r))
                .flex()
                .items_center()
                .justify_between()
                .gap_2()
                .px(px(8.0))
                .py(px(4.0))
                .rounded(px(4.0))
                .when(on, |d| d.bg(accent.opacity(0.12)))
                .cursor_pointer()
                .child(
                    div()
                        .min_w_0()
                        .flex()
                        .items_center()
                        .gap(px(8.0))
                        .when_some(checkbox, |d, checked| {
                            // Drawn, not a glyph: the bundled UI font has no ballot boxes
                            d.child(
                                div()
                                    .size(px(13.0))
                                    .flex_shrink_0()
                                    .rounded(px(3.0))
                                    .border_1()
                                    .border_color(if checked { accent } else { muted })
                                    .when(checked, |d| d.bg(accent))
                                    .flex()
                                    .items_center()
                                    .justify_center()
                                    .text_size(px(10.0))
                                    .text_color(c.inverse)
                                    .when(checked, |d| d.child("✓")),
                            )
                        })
                        .child(div().min_w_0().overflow_hidden().text_ellipsis().whitespace_nowrap().text_size(px(12.0)).text_color(text).child(label)),
                )
                .when(!value.is_empty() || caret, |d| {
                    d.child(
                        div()
                            .flex_shrink_0()
                            .min_w(px(if value_is_text { 160.0 } else { 0.0 }))
                            .px(px(8.0))
                            .py(px(2.0))
                            .rounded(px(4.0))
                            .border_1()
                            .border_color(if on { accent } else { border })
                            .text_size(px(12.0))
                            .text_color(if value == "—" { muted } else { text })
                            .when(caret && b.text_selected && !value.is_empty(), |d| d.bg(accent.opacity(0.25)))
                            .flex()
                            .items_center()
                            .child(if value_is_text || caret { value } else { format!("{value} ›") })
                            .when(caret, |d| d.child(div().ml(px(1.0)).w(px(1.0)).h(px(14.0)).bg(accent))),
                    )
                })
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation();
                    if let Some(b) = this.recipe_builder.as_mut() {
                        b.pane = Pane::Editor;
                        b.activate_row(r, false);
                    }
                    cx.notify();
                })),
        );
    }
    div()
        .flex_shrink_0()
        .max_h(px(max_h))
        .flex()
        .flex_col()
        .gap(px(6.0))
        .p(px(10.0))
        .rounded(px(6.0))
        .border_1()
        .border_color(if focused { c.accent.opacity(0.6) } else { c.border })
        .child(
            div()
                .flex()
                .items_baseline()
                .gap_2()
                .child(section_title(&format!("Step {} · {}", index + 1, step_kind(step)), c, focused))
                .child(div().text_size(px(11.0)).text_color(c.muted).child(hint)),
        )
        .child(div().id("recipe-editor-rows").flex_1().min_h_0().overflow_y_scroll().track_scroll(&b.editor_scroll).child(rows))
        .child(div().text_size(px(11.0)).text_color(c.muted).child(if focused {
            "↑↓ choose · Space or ←→ change · type to edit names and values · Esc back to steps"
        } else {
            "Enter or Tab to change these settings"
        }))
}

fn column_label(name: &str, present: bool) -> String {
    if present { name.to_string() } else { format!("{name}  (not in the file)") }
}

fn render_preview(b: &RecipeBuilder, c: &Colors) -> impl IntoElement {
    let Some(p) = &b.preview else { return div() };
    let title = match b.selected {
        None if b.recipe.source.is_remote() => "Preview of the report".to_string(),
        None => "Preview of the file".to_string(),
        Some(i) => format!("Preview after step {}", i + 1),
    };
    let meta = format!(
        "{} row{} · {} column{} · counts from the whole {}{}",
        thousands(p.total_rows),
        plural(p.total_rows),
        p.columns.len(),
        plural(p.columns.len()),
        if b.recipe.source.is_remote() { "report" } else { "file" },
        if p.total_rows > p.rows.len() { format!(" · showing the first {}", p.rows.len()) } else { String::new() }
    );
    const W: f32 = 140.0;
    let mut header = div().flex().bg(c.panel).border_b_1().border_color(c.border);
    for col in &p.columns {
        let typed = !matches!(col.rule, ColumnRule::Auto);
        header = header.child(
            div()
                .w(px(W))
                .flex_shrink_0()
                .p(px(7.0))
                .border_r_1()
                .border_color(c.border.opacity(0.5))
                .flex()
                .flex_col()
                .gap(px(4.0))
                .child(
                    div()
                        .overflow_hidden()
                        .text_ellipsis()
                        .whitespace_nowrap()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(c.text)
                        .child(col.name.clone()),
                )
                .child(
                    div()
                        .px(px(7.0))
                        .rounded(px(9.0))
                        .border_1()
                        .border_color(if typed { c.accent.opacity(0.6) } else { c.border })
                        .when(typed, |d| d.bg(c.accent.opacity(0.12)))
                        .text_size(px(11.0))
                        .text_color(c.text)
                        .child(rule_label(col.rule)),
                ),
        );
    }
    let mut body = div().flex().flex_col();
    for row in &p.rows {
        let mut line = div().flex().border_b_1().border_color(c.border.opacity(0.35));
        for (i, v) in row.iter().enumerate() {
            let number = matches!(p.columns.get(i).map(|c| c.rule), Some(ColumnRule::Number));
            let formula = v.starts_with('=');
            line = line.child(
                div()
                    .w(px(W))
                    .flex_shrink_0()
                    .px(px(7.0))
                    .py(px(4.0))
                    .border_r_1()
                    .border_color(c.border.opacity(0.35))
                    .overflow_hidden()
                    .text_ellipsis()
                    .whitespace_nowrap()
                    .text_size(px(12.0))
                    .font_family("IBM Plex Mono")
                    .text_color(if formula { c.muted } else { c.text })
                    .when(number, |d| d.flex().justify_end())
                    .when(v != v.trim(), |d| d.bg(c.warn.opacity(0.08)))
                    .child(if v.is_empty() { " ".to_string() } else { v.clone() }),
            );
        }
        body = body.child(line);
    }
    div()
        .flex_1()
        .min_h_0()
        .flex()
        .flex_col()
        .gap(px(8.0))
        .child(
            div()
                .flex()
                .items_baseline()
                .gap_2()
                .child(section_title(&title, c, false))
                .child(div().text_size(px(11.0)).text_color(c.muted).child(meta)),
        )
        .child(
            div()
                .id("recipe-preview")
                .flex_1()
                .min_h_0()
                .overflow_scroll()
                .rounded(px(6.0))
                .border_1()
                .border_color(c.border)
                .bg(c.grid_bg)
                .child(div().flex().flex_col().child(header).child(body)),
        )
}

// ----------------------------------------------------------------------------
// Footer
// ----------------------------------------------------------------------------

fn render_footer(app: &Spreadsheet, b: &RecipeBuilder, c: &Colors, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let (summary, tone) = match (&b.full, &b.error) {
        (_, Some(e)) => (e.clone(), c.error),
        (None, _) if b.recipe.source.is_remote() => ("VisiBooks can't be read with these settings.".to_string(), c.error),
        (None, _) => ("The file can't be read with these settings.".to_string(), c.error),
        (Some(r), _) if !r.ok => {
            let first = r.failures.first().cloned().unwrap_or_else(|| "a check failed".into());
            (format!("This recipe would not publish: {first}"), c.error)
        }
        (Some(r), _) => (
            format!(
                "Loads {} of {} rows. Refresh re-runs these steps on the next export; if a check fails, the Table keeps its last good result.",
                thousands(r.rows),
                thousands(r.source_rows)
            ),
            c.muted,
        ),
    };
    let primary = match b.link_table.and_then(|id| app.wb(cx).table(id).map(|(_, t)| t.name.clone())) {
        Some(name) => format!("Save and refresh {name}"),
        None => "Save and load into Table".into(),
    };
    let can_run = b.full.as_ref().is_some_and(|r| r.ok);
    div()
        .flex()
        .flex_col()
        .gap(px(8.0))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .px(px(10.0))
                .py(px(6.0))
                .rounded(px(6.0))
                .border_1()
                .border_color(c.border)
                .bg(c.grid_bg)
                .child(div().text_size(px(11.0)).text_color(c.muted).child("Same recipe, unattended:"))
                .child(
                    div()
                        .flex_1()
                        .text_size(px(12.0))
                        .font_family("IBM Plex Mono")
                        .text_color(c.text)
                        .child(cli_line(b.recipe_path.as_deref(), &b.source_path)),
                )
                .when_some(b.recipe_path.clone(), |d, path| {
                    let (accent, muted) = (c.accent, c.muted);
                    d.child(
                        div()
                            .id("recipe-edit-text")
                            .text_size(px(11.0))
                            .text_color(muted)
                            .cursor_pointer()
                            .hover(move |s| s.text_color(accent))
                            .child("Open as text")
                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                cx.stop_propagation();
                                this.edit_recipe_file(&path, cx);
                            })),
                    )
                }),
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_3()
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .child(div().text_size(px(12.0)).text_color(tone).child(summary))
                        .when(b.confirm_discard, |d| {
                            d.child(div().text_size(px(11.0)).text_color(c.warn).child("Unsaved changes. Esc again discards them."))
                        }),
                )
                .child(div().text_size(px(11.0)).text_color(c.muted).child("Tab moves · Ctrl+S saves · Ctrl+Enter runs"))
                .child(
                    Button::new("recipe-cancel", "Cancel")
                        .secondary(c.border, c.text)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            cx.stop_propagation();
                            this.close_recipe_builder(cx);
                        })),
                )
                .child(
                    Button::new("recipe-save", "Save recipe")
                        .secondary(c.border, c.text)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            cx.stop_propagation();
                            this.recipe_builder_save(false, cx);
                        })),
                )
                .child(
                    Button::new("recipe-run", primary)
                        .disabled(!can_run)
                        .primary(c.accent, c.inverse)
                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                            cx.stop_propagation();
                            if can_run {
                                this.recipe_builder_save(true, cx);
                            }
                        })),
                ),
        )
}

/// The key that picks add-menu entry `k`: 1-9, then a, b, c, d.
fn add_kind_key(k: usize) -> char {
    if k < 9 { (b'1' + k as u8) as char } else { (b'a' + (k - 9) as u8) as char }
}
