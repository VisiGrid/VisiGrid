//! Pivot field list. Mouse controls share the keyboard's draft operations;
//! the workbook changes only when the user applies the draft.

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::pivot::Aggregation;

use crate::app::Spreadsheet;
use crate::pivot_ui::{PivotPanel, PivotPanelItem, PivotPanelMode};
use crate::theme::TokenKey;
use crate::ui::Button;

pub(crate) const PANEL_WIDTH: f32 = 380.0;

#[derive(Clone, Copy)]
struct Colors {
    surface: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
    accent: Hsla,
}

/// Compact, local drawer control. Draft actions are inert during computation.
fn control(
    id: impl Into<ElementId>,
    label: impl Into<SharedString>,
    enabled: bool,
    colors: Colors,
    action: impl Fn(&mut PivotPanel) + 'static,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let label: SharedString = label.into();
    div()
        .id(id)
        .px_2()
        .py_1()
        .rounded_sm()
        .text_size(px(11.0))
        .text_color(if enabled {
            colors.text
        } else {
            colors.muted.opacity(0.5)
        })
        .when(enabled, |d| {
            d.cursor_pointer()
                .hover(|s| s.bg(colors.accent.opacity(0.12)))
        })
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |this, _, _, cx| {
                cx.stop_propagation();
                if enabled {
                    if let Some(panel) = this.pivot_panel.as_mut().filter(|p| !p.busy) {
                        action(panel);
                        cx.notify();
                    }
                }
            }),
        )
        .child(label)
}

fn badge(label: impl Into<SharedString>, colors: Colors) -> Div {
    let label: SharedString = label.into();
    div()
        .flex_shrink_0()
        .px(px(6.0))
        .py(px(2.0))
        .rounded_sm()
        .bg(colors.accent.opacity(0.09))
        .text_size(px(10.0))
        .text_color(colors.text)
        .child(label)
}

fn well(title: &str, key: &str, description: &str, content: Div, colors: Colors) -> Div {
    div()
        .flex()
        .flex_col()
        .flex_shrink_0()
        .rounded_md()
        .border_1()
        .border_color(colors.border)
        .bg(colors.surface)
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .px_3()
                .py_2()
                .border_b_1()
                .border_color(colors.border)
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(colors.text)
                        .child(title.to_string()),
                )
                .child(badge(key.to_string(), colors)),
        )
        .child(content)
        .when(!description.is_empty(), |d| {
            d.child(
                div()
                    .px_3()
                    .py_2()
                    .text_size(px(11.0))
                    .text_color(colors.muted)
                    .child(description.to_string()),
            )
        })
}

pub(crate) fn render_pivot_panel(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let colors = Colors {
        surface: app.token(TokenKey::EditorBg),
        border: app.token(TokenKey::PanelBorder),
        text: app.token(TokenKey::TextPrimary),
        muted: app.token(TokenKey::TextMuted),
        accent: app.token(TokenKey::Accent),
    };
    let Some(panel) = app.pivot_panel.as_ref() else {
        return div().id("pivot-panel-empty");
    };
    let current = panel.current();
    let is_new = matches!(panel.mode, PivotPanelMode::New);
    let (title, stale) = match &panel.mode {
        PivotPanelMode::New => ("New pivot table".to_string(), false),
        PivotPanelMode::Edit { pivot_id } => match app.wb(cx).find_pivot(*pivot_id) {
            Some((_, t)) => (t.name.clone(), app.wb(cx).is_pivot_stale(t)),
            None => ("Pivot table".to_string(), false),
        },
    };
    let source_sheet = app
        .wb(cx)
        .sheet_by_id(panel.source.sheet_id)
        .map(|s| s.name.clone())
        .unwrap_or_else(|| "?".into());

    let has_tables = app.wb(cx).tables().next().is_some();
    let mut sources = div().flex().flex_col().gap_1().mt_2();
    if panel.source_menu {
        sources = sources.child(div().text_size(px(11.0)).text_color(colors.muted)
            .child("Choose a Table. Changing sources clears the draft fields."));
        for (index, (sheet_id, table)) in app.wb(cx).tables().enumerate() {
            let id = table.id;
            let name = app.wb(cx).sheet_by_id(sheet_id).map(|s| s.name.as_str()).unwrap_or("?");
            sources = sources.child(div().id(SharedString::from(format!("pivot-source-{}", id.0)))
                .px_2().py_2().rounded_sm().cursor_pointer().text_size(px(12.0)).text_color(colors.text)
                .when(panel.source_cursor == index, |d| d.bg(colors.accent.opacity(0.12)))
                .hover(|s| s.bg(colors.accent.opacity(0.08)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation(); this.pivot_choose_table(id, cx);
                }))
                .child(format!("{} · {} · {} rows", table.name, name, table.range.data_rows())));
        }
    }

    let row = |id: String,
               label: String,
               detail: Option<String>,
               item: PivotPanelItem,
               cx: &mut Context<Spreadsheet>| {
        let selected = current == Some(item);
        let assigned = !matches!(item, PivotPanelItem::Field(_));
        div()
            .id(SharedString::from(id.clone()))
            .flex()
            .items_center()
            .gap_2()
            .min_h(px(32.0))
            .px_2()
            .mx_1()
            .rounded_sm()
            .cursor_pointer()
            .text_size(px(12.0))
            .text_color(colors.text)
            .when(selected, |d| d.bg(colors.accent.opacity(0.16)))
            .hover(|s| s.bg(colors.accent.opacity(0.08)))
            .on_mouse_down(
                MouseButton::Left,
                cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation();
                    if let Some(p) = this.pivot_panel.as_mut().filter(|p| !p.busy) {
                        p.move_cursor_to(item);
                    }
                    cx.notify();
                }),
            )
            .child(
                div()
                    .flex_1()
                    .min_w_0()
                    .overflow_hidden()
                    .text_ellipsis()
                    .child(label),
            )
            .when_some(detail, |d, text| d.child(badge(text, colors)))
            .when(assigned, |d| {
                d.child(control(
                    SharedString::from(format!("{id}-remove")),
                    "×",
                    !panel.busy,
                    colors,
                    move |p| {
                        p.move_cursor_to(item);
                        p.remove_current();
                        p.message = None;
                    },
                    cx,
                ))
            })
    };

    let mut fields = div().flex().flex_col().py_1();
    for (i, header) in panel.headers.iter().enumerate() {
        let mut tags = Vec::new();
        if panel.draft.rows.iter().any(|f| f.offset as usize == i) {
            tags.push("Rows");
        }
        if panel
            .draft
            .column
            .as_ref()
            .is_some_and(|f| f.offset as usize == i)
        {
            tags.push("Column");
        }
        if panel
            .draft
            .values
            .iter()
            .any(|v| v.field.offset as usize == i)
        {
            tags.push("Values");
        }
        fields = fields.child(row(
            format!("pivot-field-{i}"),
            header.clone(),
            (!tags.is_empty()).then(|| tags.join(" · ")),
            PivotPanelItem::Field(i),
            cx,
        ));
    }
    let field_selected = matches!(current, Some(PivotPanelItem::Field(_))) && !panel.busy;
    let mut assign = div()
        .flex()
        .items_center()
        .justify_between()
        .px_2()
        .py_2()
        .border_t_1()
        .border_color(colors.border);
    for (target, label) in [
        ('r', "+ Rows  R"),
        ('c', "+ Column  C"),
        ('v', "+ Values  V"),
    ] {
        assign = assign.child(
            control(
                SharedString::from(format!("pivot-assign-{target}")),
                label,
                field_selected,
                colors,
                move |p| p.message = p.assign(target),
                cx,
            )
            .border_1()
            .border_color(colors.border),
        );
    }

    let mut rows = div().flex().flex_col().py_1();
    for (k, f) in panel.draft.rows.iter().enumerate() {
        rows = rows.child(row(
            format!("pivot-row-{k}"),
            f.header.clone(),
            None,
            PivotPanelItem::Row(k),
            cx,
        ));
    }
    let mut column = div().flex().flex_col().py_1();
    if let Some(f) = &panel.draft.column {
        column = column.child(row(
            "pivot-column".into(),
            f.header.clone(),
            None,
            PivotPanelItem::Column,
            cx,
        ));
    }
    let mut values = div().flex().flex_col().py_1();
    for (k, v) in panel.draft.values.iter().enumerate() {
        values = values.child(row(
            format!("pivot-value-{k}"),
            v.field.header.clone(),
            Some(v.aggregation.label().into()),
            PivotPanelItem::Value(k),
            cx,
        ));
    }
    if let Some(PivotPanelItem::Value(k)) = current {
        let selected = panel.draft.values[k].aggregation;
        let mut calculations = div().flex().flex_wrap().gap_1().px_2().py_2();
        for agg in Aggregation::ALL {
            calculations = calculations.child(
                control(
                    SharedString::from(format!("pivot-aggregation-{}", agg.label())),
                    agg.label(),
                    !panel.busy,
                    colors,
                    move |p| {
                        p.set_aggregation(k, agg);
                        p.message = None;
                    },
                    cx,
                )
                .border_1()
                .border_color(if agg == selected {
                    colors.accent
                } else {
                    colors.border
                })
                .when(agg == selected, |d| d.bg(colors.accent.opacity(0.1))),
            );
        }
        values = values
            .child(
                div()
                    .mx_3()
                    .mt_2()
                    .text_size(px(11.0))
                    .text_color(colors.muted)
                    .child("Summarize by"),
            )
            .child(calculations)
            .when(panel.draft.values.len() > 1, |d| {
                d.child(
                    control(
                        "pivot-aggregation-all",
                        "Apply calculation to all values  =",
                        !panel.busy,
                        colors,
                        |p| p.message = p.apply_aggregation_to_all(),
                        cx,
                    )
                    .mx_2()
                    .mb_1(),
                )
            });
    }

    let mut layout = div()
        .flex()
        .flex_col()
        .gap_2()
        .child(well(
            "Rows",
            "R",
            if panel.draft.rows.is_empty() {
                "Choose a field to group records into rows."
            } else {
                ""
            },
            rows,
            colors,
        ))
        .child(well(
            "Column",
            "C · optional",
            if panel.draft.column.is_none() {
                "Choose a field to compare groups side by side."
            } else {
                ""
            },
            column,
            colors,
        ))
        .child(well(
            "Values",
            "V",
            if panel.draft.values.is_empty() {
                "Choose a field to summarize."
            } else {
                ""
            },
            values,
            colors,
        ));
    if let Some(item @ (PivotPanelItem::Row(k) | PivotPanelItem::Value(k))) = current {
        let count = if matches!(item, PivotPanelItem::Row(_)) {
            panel.draft.rows.len()
        } else {
            panel.draft.values.len()
        };
        layout = layout.child(
            div()
                .flex()
                .items_center()
                .gap_1()
                .child(
                    div()
                        .flex_1()
                        .text_size(px(11.0))
                        .text_color(colors.muted)
                        .child("Selected field"),
                )
                .child(control(
                    "pivot-move-up",
                    "↑ Move up",
                    k > 0 && !panel.busy,
                    colors,
                    |p| p.reorder(true),
                    cx,
                ))
                .child(control(
                    "pivot-move-down",
                    "↓ Move down",
                    k + 1 < count && !panel.busy,
                    colors,
                    |p| p.reorder(false),
                    cx,
                )),
        );
    }

    let can_apply = !panel.busy && !panel.draft.is_empty();
    let hint = match current {
        Some(PivotPanelItem::Field(_)) => "↑ ↓ Select field · R / C / V Add to layout",
        Some(PivotPanelItem::Value(_)) => "← → Calculation · Alt+↑ ↓ Reorder · Del Remove",
        _ => "Alt+↑ ↓ Reorder · Del Remove · Esc Close",
    };
    div()
        .id("pivot-panel")
        .absolute()
        .right_0()
        .top_0()
        .h_full()
        .w(px(PANEL_WIDTH))
        .bg(app.token(TokenKey::PanelBg))
        .border_l_1()
        .border_color(colors.border)
        .flex()
        .flex_col()
        .overflow_hidden()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_mouse_up(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .px_4()
                .py_3()
                .flex_shrink_0()
                .border_b_1()
                .border_color(colors.border)
                .child(
                    div()
                        .flex_1()
                        .min_w_0()
                        .child(
                            div()
                                .text_size(px(14.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(colors.text)
                                .child(title),
                        )
                        .child(
                            div()
                                .mt_1()
                                .text_size(px(11.0))
                                .text_color(colors.muted)
                                .child("Group and summarize your data"),
                        ),
                )
                .child(
                    div()
                        .id("pivot-panel-close")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_color(colors.muted)
                        .hover(|s| s.bg(colors.border).text_color(colors.text))
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|this, _, _, cx| {
                                cx.stop_propagation();
                                this.close_pivot_panel(cx);
                            }),
                        )
                        .child("×"),
                ),
        )
        .child(
            div()
                .id("pivot-panel-lists")
                .flex_1()
                .min_h_0()
                .overflow_y_scroll()
                .flex()
                .flex_col()
                .gap_4()
                .p_4()
                .child(
                    div()
                        .p_3()
                        .flex_shrink_0()
                        .rounded_md()
                        .border_1()
                        .border_color(colors.border)
                        .bg(colors.surface)
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .justify_between()
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .font_weight(FontWeight::MEDIUM)
                                        .text_color(colors.muted)
                                        .child("SOURCE DATA"),
                                )
                                .child(badge(format!("{} rows", panel.source.data_rows()), colors)),
                        )
                        .child(
                            div()
                                .mt_2()
                                .text_size(px(12.0))
                                .text_color(colors.text)
                                .child(panel.table_name.clone().unwrap_or_else(|| format!("{}!{}", source_sheet, panel.source_label()))),
                        )
                        .when(panel.table_name.is_some(), |d| d.child(div().mt_1().text_size(px(11.0)).text_color(colors.muted)
                            .child("All Table records · filters ignored. Refresh includes new rows.")))
                        .when(has_tables, |d| d.child(control(
                            "pivot-source-chooser", if panel.source_menu { "Close source list" } else { "Choose Table…  S" },
                            !panel.busy, colors, |p| p.source_menu = !p.source_menu, cx)))
                        .child(sources)
                        .when(stale, |d| {
                            d.child(
                                div()
                                    .mt_2()
                                    .text_size(px(11.0))
                                    .text_color(app.token(TokenKey::Warn))
                                    .child("Source changed · Apply to refresh"),
                            )
                        })
                        .when_some(panel.growth, |d, last| {
                            d.child(div().mt_2().child(control(
                                "pivot-include-rows",
                                format!(
                                    "Include {} new rows below source  X",
                                    last - panel.source.end_row
                                ),
                                !panel.busy,
                                colors,
                                |p| {
                                    if let Some(last) = p.growth.take() {
                                        p.source.end_row = last;
                                        p.message = Some(
                                            "Source extended. Apply to use the new rows.".into(),
                                        );
                                    }
                                },
                                cx,
                            )))
                        }),
                )
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .flex_shrink_0()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(colors.text)
                                .child(format!("Source fields · {}", panel.headers.len())),
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(colors.muted)
                                .child("Select a field, then add it to the layout."),
                        )
                        .child(
                            div()
                                .rounded_md()
                                .border_1()
                                .border_color(colors.border)
                                .bg(colors.surface)
                                .overflow_hidden()
                                .child(fields)
                                .child(assign),
                        ),
                )
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .flex_shrink_0()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(colors.text)
                                .child("Pivot layout"),
                        )
                        .child(layout),
                ),
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .p_4()
                .flex_shrink_0()
                .border_t_1()
                .border_color(colors.border)
                .bg(colors.surface)
                .when_some(panel.message.clone(), |d, m| {
                    d.child(
                        div()
                            .p_2()
                            .rounded_sm()
                            .bg(colors.accent.opacity(0.08))
                            .text_size(px(12.0))
                            .text_color(colors.text)
                            .child(m),
                    )
                })
                .child(div().text_size(px(11.0)).text_color(colors.muted).child(
                    if panel.draft.is_empty() {
                        "Add a field to Rows, Column or Values to begin."
                    } else if is_new {
                        "Output will be created on a new sheet."
                    } else {
                        "Apply updates the pivot and refreshes its results."
                    },
                ))
                .child(
                    Button::new(
                        "pivot-panel-apply",
                        if panel.busy {
                            "Computing…"
                        } else if is_new {
                            "Create pivot table"
                        } else {
                            "Apply changes"
                        },
                    )
                    .disabled(!can_apply)
                    .primary(colors.accent, app.token(TokenKey::TextInverse))
                    .w_full()
                    .flex()
                    .items_center()
                    .justify_center()
                    .gap_2()
                    .when(can_apply, |d| {
                        d.on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|this, _, _, cx| {
                                cx.stop_propagation();
                                this.pivot_apply(cx);
                            }),
                        )
                    })
                    .child(div().text_size(px(10.0)).opacity(0.8).child("↵")),
                )
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(colors.muted)
                        .child(hint),
                ),
        )
}
