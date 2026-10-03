//! Import recipes: the strip above a sheet with a recipe-backed Table, and
//! the banner shown when a run did not publish.
//!
//! State and actions live in `recipe_ui.rs`.

use std::collections::BTreeMap;
use std::path::PathBuf;

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_io::recipe::OnError;

use crate::app::Spreadsheet;
use crate::recipe_ui::{file_name, when, RecipeBlocked, RecipeTarget, RECIPE_STRIP_HEIGHT};
use crate::theme::TokenKey;

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
        }
    }
}

fn plural(n: usize) -> &'static str {
    if n == 1 { "" } else { "s" }
}

/// A small bordered button; `primary` fills it with the accent.
fn action(
    id: impl Into<ElementId>,
    label: impl Into<SharedString>,
    primary: bool,
    c: &Colors,
    cx: &mut Context<Spreadsheet>,
    f: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    let (accent, border, text, inverse) = (c.accent, c.border, c.text, c.inverse);
    div()
        .id(id.into())
        .flex_shrink_0()
        .px(px(10.0))
        .py(px(4.0))
        .rounded(px(4.0))
        .border_1()
        .text_size(px(12.0))
        .cursor_pointer()
        .when(primary, |d| d.bg(accent).border_color(accent).text_color(inverse).font_weight(FontWeight::SEMIBOLD))
        .when(!primary, |d| d.border_color(border).text_color(text).hover(move |s| s.border_color(accent)))
        .child(label.into())
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            cx.stop_propagation();
            f(this, cx);
        }))
}

fn key_cap(label: &'static str, c: &Colors) -> impl IntoElement {
    div()
        .px(px(5.0))
        .rounded(px(4.0))
        .border_1()
        .border_color(c.border)
        .text_size(px(10.0))
        .text_color(c.muted)
        .child(label)
}

// ============================================================================
// Strip
// ============================================================================

/// One line above the grid: which recipe the Table comes from, its last
/// good refresh, and Refresh / Edit recipe.
pub(crate) fn render_recipe_strip(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(table) = app.recipe_strip_table(cx) else {
        return div().into_any_element();
    };
    let Some(source) = table.source.clone() else {
        return div().into_any_element();
    };
    let c = Colors::new(app);
    let blocked_here = app
        .recipe_blocked
        .as_ref()
        .is_some_and(|b| b.target == RecipeTarget::Table(table.id));
    let recipe_path = PathBuf::from(&source.recipe);
    let detail = match &source.refreshed {
        Some(s) if blocked_here => format!(
            "last good refresh {} from {} · {} rows · the latest refresh didn't publish",
            when(&s.at),
            file_name(&s.source),
            s.rows
        ),
        Some(s) => format!("refreshed {} from {} · {} rows", when(&s.at), file_name(&s.source), s.rows),
        None => "not refreshed yet".into(),
    };
    let running = app.recipe_run_in_progress;
    let edit_path = recipe_path.clone();
    let table_id = table.id;
    div()
        .id("recipe-strip")
        .h(px(RECIPE_STRIP_HEIGHT))
        .flex_shrink_0()
        .flex()
        .items_center()
        .gap(px(10.0))
        .px_3()
        .bg(c.panel)
        .border_b_1()
        .border_color(c.border)
        .text_size(px(12.0))
        .child(
            div()
                .flex_shrink_0()
                .px(px(7.0))
                .py(px(1.0))
                .rounded(px(4.0))
                .border_1()
                .border_color(c.accent.opacity(0.5))
                .text_color(c.accent)
                .font_weight(FontWeight::SEMIBOLD)
                .child(format!("Table {}", table.name)),
        )
        .child(
            div()
                .flex_1()
                .min_w_0()
                .overflow_hidden()
                .text_ellipsis()
                .whitespace_nowrap()
                .text_color(if blocked_here { c.warn } else { c.muted })
                .child(format!("from recipe {} · {detail}", file_name(&source.recipe))),
        )
        .child({
            let path = recipe_path.clone();
            action("recipe-strip-choose", "Choose file…", false, &c, cx, move |this, cx| {
                this.recipe_choose_source(path.clone(), RecipeTarget::Table(table_id), cx)
            })
        })
        .child(action("recipe-strip-edit", "Edit recipe", false, &c, cx, move |this, cx| {
            this.open_recipe_builder(&edit_path, Some(table_id), cx)
        }))
        .child(
            div()
                .id("recipe-strip-refresh")
                .flex_shrink_0()
                .flex()
                .items_center()
                .gap(px(6.0))
                .px(px(10.0))
                .py(px(3.0))
                .rounded(px(4.0))
                .bg(c.accent)
                .text_color(c.inverse)
                .font_weight(FontWeight::SEMIBOLD)
                .cursor_pointer()
                .child(if running { "Running…" } else { "Refresh" })
                .child(div().text_size(px(10.0)).font_weight(FontWeight::NORMAL).child("Alt+F5"))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.refresh_recipe_table(cx);
                })),
        )
        .into_any_element()
}

// ============================================================================
// Blocked banner
// ============================================================================

struct Problem {
    title: String,
    body: String,
    actions: Vec<AnyElement>,
}

fn problems(b: &RecipeBlocked, c: &Colors, cx: &mut Context<Spreadsheet>) -> Vec<Problem> {
    let mut out = Vec::new();
    let path = b.recipe_path.clone();
    let link = match b.target {
        RecipeTarget::Table(id) => Some(id),
        RecipeTarget::NewWorkbook => None,
    };
    let edit = |id: String, c: &Colors, cx: &mut Context<Spreadsheet>| {
        let path = path.clone();
        action(id, "Edit recipe", false, c, cx, move |this, cx| this.open_recipe_builder(&path, link, cx)).into_any_element()
    };
    let mut covered_steps = Vec::new();

    // Columns a step needs that the file no longer has: one problem per
    // column, naming every step that needs it. A column a failed rename
    // would have produced is a consequence, not a second problem.
    let failed: Vec<_> = b.report.steps.iter().filter(|s| s.failed && !s.missing.is_empty()).collect();
    let mut cascade: Vec<String> = Vec::new();
    for step in &failed {
        if let Some(visigrid_io::recipe::Step::Rename { columns, .. }) = b.recipe.steps.get(step.index - 1) {
            for (from, to) in columns {
                if step.missing.iter().any(|m| m.eq_ignore_ascii_case(from)) {
                    cascade.push(to.to_lowercase());
                }
            }
        }
    }
    let mut by_name: Vec<(String, Vec<usize>)> = Vec::new();
    for step in &failed {
        covered_steps.push(step.index);
        for name in &step.missing {
            if cascade.contains(&name.to_lowercase()) {
                continue;
            }
            match by_name.iter_mut().find(|(n, _)| n.eq_ignore_ascii_case(name)) {
                Some((_, steps)) => steps.push(step.index),
                None => by_name.push((name.clone(), vec![step.index])),
            }
        }
    }
    for (name, steps) in &by_name {
        let first = b.report.steps.iter().find(|s| s.index == steps[0]).unwrap();
        {
            let step = first;
            let renamed = b
                .report
                .drift
                .possibly_renamed
                .iter()
                .find(|(old, _)| old.eq_ignore_ascii_case(name))
                .map(|(_, new)| new.clone());
            let mut actions = Vec::new();
            let body = match &renamed {
                Some(new) => {
                    let (old, new2) = (name.clone(), new.clone());
                    let is_primary = crate::recipe_ui::primary_rename(b).is_some_and(|(o, _)| o.eq_ignore_ascii_case(name));
                    actions.push(
                        action(
                            format!("recipe-use-{}-{name}", step.index),
                            if is_primary { format!("Use {new}  ↵") } else { format!("Use {new}") },
                            true,
                            c,
                            cx,
                            move |this, cx| this.recipe_fix_rename(old.clone(), new2.clone(), cx),
                        )
                        .into_any_element(),
                    );
                    format!("New column {new} sits where it was. Possibly renamed. Nothing changes until you choose; \"Use {new}\" saves that into the recipe.")
                }
                None if !b.report.drift.new.is_empty() => {
                    format!("Columns new in this file: {}.", b.report.drift.new.join(", "))
                }
                None => "The file has no column by that name.".into(),
            };
            actions.push(edit(format!("recipe-edit-missing-{}-{name}", step.index), c, cx));
            let title = if steps.len() == 1 {
                format!("Column {name} is not in this file (step {}, {})", step.index, step.description)
            } else {
                let list: Vec<String> = steps.iter().map(|i| i.to_string()).collect();
                format!("Column {name} is not in this file (steps {})", list.join(", "))
            };
            out.push(Problem { title, body, actions });
        }
    }

    // Values that don't fit a type, per step
    let mut by_step: BTreeMap<usize, Vec<&visigrid_io::recipe::CellError>> = BTreeMap::new();
    for e in &b.report.errors {
        by_step.entry(e.step).or_default().push(e);
    }
    let mut types_shown = false;
    // Reports keep the first 1,000 located errors; the count is exact
    let capped_single = b.report.errors.len() < b.report.error_count && by_step.len() == 1;
    for (step, errors) in by_step {
        let Some(report) = b.report.steps.iter().find(|s| s.index == step && s.failed) else { continue };
        covered_steps.push(step);
        types_shown = true;
        let columns: Vec<&str> = {
            let mut v: Vec<&str> = errors.iter().map(|e| e.column.as_str()).collect();
            v.dedup();
            v
        };
        let count = if capped_single { b.report.error_count } else { errors.len() };
        let reason = errors[0].reason.clone();
        let examples: Vec<String> = errors.iter().take(3).map(|e| format!("line {} \"{}\"", e.line, e.value)).collect();
        out.push(Problem {
            title: format!(
                "{count} value{} in {} {} {reason} (step {step}, {})",
                plural(count),
                columns.join(", "),
                if count == 1 { "is" } else { "are" },
                report.description
            ),
            body: examples.join(" · "),
            actions: vec![
                action(format!("recipe-keep-{step}"), "Keep as text", false, c, cx, move |this, cx| {
                    this.recipe_fix_on_error(step, OnError::KeepText, cx)
                })
                .into_any_element(),
                action(format!("recipe-blank-{step}"), "Leave blank", false, c, cx, move |this, cx| {
                    this.recipe_fix_on_error(step, OnError::Blank, cx)
                })
                .into_any_element(),
                edit(format!("recipe-edit-types-{step}"), c, cx),
            ],
        });
    }

    // Anything else the run reported
    for f in &b.report.failures {
        let tied = covered_steps.iter().any(|i| f.starts_with(&format!("step {i} (")));
        if tied || (types_shown && f.contains("did not fit")) {
            continue;
        }
        out.push(Problem {
            title: f.clone(),
            body: String::new(),
            actions: vec![edit(format!("recipe-edit-other-{}", out.len()), c, cx)],
        });
    }

    if let Some(reason) = &b.refused {
        out.push(Problem {
            title: "The result couldn't be placed".into(),
            body: reason.clone(),
            actions: vec![action("recipe-retry-refused", "Try again", false, c, cx, |this, cx| this.recipe_retry(cx)).into_any_element()],
        });
    }
    out
}

pub(crate) fn render_recipe_banner(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(b) = app.recipe_blocked.as_ref() else {
        return div().into_any_element();
    };
    let c = Colors::new(app);
    let list = problems(b, &c, cx);
    let headline = match (b.target, &b.last_good) {
        (RecipeTarget::Table(_), Some(s)) => format!(
            "Refresh didn't publish. {} still shows the result from {}.",
            b.table_name,
            file_name(&s.source)
        ),
        (RecipeTarget::Table(_), None) => format!("Refresh didn't publish. {} is unchanged.", b.table_name),
        (RecipeTarget::NewWorkbook, _) => format!("{} didn't run cleanly. Nothing was opened.", file_name(&b.recipe_path.display().to_string())),
    };
    let detail = format!(
        "{} · {} problem{} · nothing in the sheet changed",
        file_name(&b.report.source),
        list.len(),
        plural(list.len())
    );
    let (muted, text) = (c.muted, c.text);

    let mut card = div()
        .id("recipe-banner")
        .w_full()
        .max_w(px(640.0))
        .max_h(px(520.0))
        .overflow_y_scroll()
        .bg(c.panel)
        .border_1()
        .border_color(c.error.opacity(0.6))
        .rounded_md()
        .shadow_lg()
        .flex()
        .flex_col()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .px(px(14.0))
                .py(px(10.0))
                .border_l(px(3.0))
                .border_color(c.error)
                .bg(c.error.opacity(0.07))
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .child(div().text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child(headline))
                        .child(div().text_size(px(11.0)).text_color(c.muted).child(detail)),
                )
                .child(
                    div()
                        .id("recipe-banner-dismiss")
                        .flex()
                        .items_center()
                        .gap(px(6.0))
                        .text_size(px(13.0))
                        .text_color(muted)
                        .cursor_pointer()
                        .hover(move |s| s.text_color(text))
                        .child(key_cap("Esc", &c))
                        .child("✕")
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            cx.stop_propagation();
                            this.dismiss_recipe_blocked(cx);
                        })),
                ),
        );

    for (i, p) in list.into_iter().enumerate() {
        card = card.child(
            div()
                .id(("recipe-problem", i))
                .flex()
                .gap(px(10.0))
                .px(px(14.0))
                .py(px(10.0))
                .border_t_1()
                .border_color(c.border.opacity(0.6))
                .child(div().mt(px(5.0)).size(px(8.0)).flex_shrink_0().rounded_full().bg(c.error))
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .flex()
                        .flex_col()
                        .gap(px(4.0))
                        .child(div().text_size(px(12.0)).font_weight(FontWeight::SEMIBOLD).text_color(c.text).child(p.title))
                        .when(!p.body.is_empty(), |d| {
                            d.child(div().text_size(px(11.0)).line_height(px(16.0)).text_color(c.muted).child(p.body))
                        })
                        .child(div().flex().flex_wrap().gap(px(8.0)).mt(px(2.0)).children(p.actions)),
                ),
        );
    }

    if let Some(e) = &b.save_error {
        card = card.child(
            div()
                .px(px(14.0))
                .py(px(8.0))
                .text_size(px(11.0))
                .text_color(c.error)
                .child(format!("Couldn't save the fix to the recipe: {e}")),
        );
    }
    card = card.child(
        div()
            .flex()
            .items_center()
            .gap(px(10.0))
            .px(px(14.0))
            .py(px(8.0))
            .border_t_1()
            .border_color(c.border.opacity(0.6))
            .child(
                div()
                    .flex_1()
                    .min_w_0()
                    .text_size(px(11.0))
                    .line_height(px(16.0))
                    .text_color(c.muted)
                    .child("Each fix is saved into the recipe and retried against the same copy of the file, so what you checked is what gets loaded. From the command line the same run exits with code 70."),
            )
            .child({
                let (path, target) = (b.recipe_path.clone(), b.target);
                action("recipe-choose", "Choose file…", false, &c, cx, move |this, cx| {
                    this.recipe_choose_source(path.clone(), target, cx)
                })
            })
            .child(action("recipe-retry", "Read the file again", false, &c, cx, |this, cx| this.recipe_retry(cx))),
    );

    div()
        .absolute()
        .top_0()
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .pt_2()
        .px_4()
        .child(card)
        .into_any_element()
}
