use crate::{app::Spreadsheet, table_ui::TableDialogKind, theme::TokenKey};
use gpui::{prelude::*, *};
use visigrid_engine::{
    filter::SortDirection,
    table::{DataTable, TableId},
    table_view::TableViewSpec,
};

fn summary(table: &DataTable, spec: &TableViewSpec) -> String {
    let mut parts = Vec::new();
    if let Some(sort) = &spec.sort {
        let column = table
            .columns
            .iter()
            .find(|c| c.id == sort.column)
            .map_or("?", |c| c.name.as_str());
        parts.push(format!(
            "{} {}",
            column,
            if sort.direction == SortDirection::Ascending {
                "↑"
            } else {
                "↓"
            }
        ));
    }
    if !spec.filters.is_empty() {
        parts.push(format!(
            "{} filtered column{}",
            spec.filters.len(),
            if spec.filters.len() == 1 { "" } else { "s" }
        ));
    }
    if parts.is_empty() {
        parts.push("All records · stored order".into());
    }
    if !spec.show_filter_buttons {
        parts.push("Filter buttons hidden".into());
    }
    parts.join(" · ")
}

fn action(
    id: &'static str,
    label: &'static str,
    app: &Spreadsheet,
    f: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .id(id)
        .px_3()
        .py_2()
        .rounded_md()
        .border_1()
        .border_color(app.token(TokenKey::PanelBorder))
        .text_size(px(12.0))
        .text_color(app.token(TokenKey::TextPrimary))
        .cursor_pointer()
        .hover(|s| s.bg(app.token(TokenKey::SelectionBg)))
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |s, _, _, cx| {
                cx.stop_propagation();
                f(s, cx);
            }),
        )
        .child(label)
}

pub(crate) fn render(app: &Spreadsheet, id: TableId, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(table) = app.wb(cx).table(id).map(|(_, t)| t.clone()) else {
        return div().into_any_element();
    };
    let d = app.table_dialog.as_ref().unwrap();
    let selected = d.field.min(table.saved_views.len().saturating_sub(1));
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let current = crate::table_saved_views::current_spec(app.wb(cx), id).ok();
    let mut list = div()
        .id("saved-view-list")
        .max_h(px(280.0))
        .overflow_y_scroll()
        .flex()
        .flex_col()
        .gap_2();
    for (index, saved) in table.saved_views.iter().enumerate() {
        list = list.child(
            div()
                .id(("saved-view", index))
                .flex_shrink_0()
                .p_3()
                .rounded_md()
                .border_1()
                .border_color(if index == selected { accent } else { border })
                .bg(if index == selected {
                    accent.opacity(0.08)
                } else {
                    app.token(TokenKey::EditorBg)
                })
                .cursor_pointer()
                .flex()
                .flex_col()
                .gap_1()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(move |s, e: &MouseDownEvent, _, cx| {
                        cx.stop_propagation();
                        if let Some(d) = &mut s.table_dialog {
                            d.field = index;
                            d.error = None;
                        }
                        if e.click_count == 2 {
                            s.apply_named_table_view(id, index, cx);
                        }
                        cx.notify();
                    }),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .gap_3()
                        .child(
                            div()
                                .min_w_0()
                                .overflow_hidden()
                                .text_ellipsis()
                                .text_size(px(14.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(text)
                                .child(saved.name.clone()),
                        )
                        .when(current.as_ref() == Some(&saved.view), |s| {
                            s.child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(muted)
                                    .child("Matches current"),
                            )
                        }),
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(muted)
                        .child(summary(&table, &saved.view)),
                ),
        );
    }
    if table.saved_views.is_empty() {
        list = list.child(div().py_5().text_size(px(13.0)).text_color(muted)
            .child("No saved views yet. Set up this Table's sorting and filters, then save the current view."));
    }
    let content = div()
        .w(px(560.0))
        .p_6()
        .rounded_md()
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(border)
        .shadow_lg()
        .flex()
        .flex_col()
        .gap_4()
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text)
                        .child("Saved Table views"),
                )
                .child(action(
                    "saved-view-new",
                    "Save current…",
                    app,
                    move |s, cx| s.open_table_dialog(TableDialogKind::SaveView(id), cx),
                    cx,
                )),
        )
        .child(div().text_size(px(13.0)).text_color(muted).child(format!(
            "{} · Reuse sorting and filters. Manually hidden rows stay separate.",
            table.name
        )))
        .child(list)
        .when(!table.saved_views.is_empty(), |s| {
            s.child(
                div()
                    .flex()
                    .gap_2()
                    .child(action(
                        "saved-view-rename",
                        "Rename…",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::RenameView(id, selected), cx)
                        },
                        cx,
                    ))
                    .child(action(
                        "saved-view-update",
                        "Update from current…",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::UpdateView(id, selected), cx)
                        },
                        cx,
                    ))
                    .child(action(
                        "saved-view-delete",
                        "Delete…",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::DeleteView(id, selected), cx)
                        },
                        cx,
                    )),
            )
        })
        .when_some(d.error.clone(), |s, e| {
            s.child(
                div()
                    .text_size(px(12.0))
                    .text_color(app.token(TokenKey::Error))
                    .child(e),
            )
        })
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .pt_4()
                .border_t_1()
                .border_color(border)
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(muted)
                        .child("↑ ↓ to select · Enter to apply · Esc to close"),
                )
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .child(action(
                            "saved-view-close",
                            "Close",
                            app,
                            |s, cx| {
                                s.table_dialog = None;
                                cx.notify();
                            },
                            cx,
                        ))
                        .when(!table.saved_views.is_empty(), |s| {
                            s.child(
                                div()
                                    .id("saved-view-apply")
                                    .px_4()
                                    .py_2()
                                    .rounded_md()
                                    .bg(accent)
                                    .text_color(app.token(TokenKey::TextInverse))
                                    .text_size(px(13.0))
                                    .font_weight(FontWeight::SEMIBOLD)
                                    .cursor_pointer()
                                    .hover(|s| s.bg(accent.opacity(0.85)))
                                    .on_mouse_down(
                                        MouseButton::Left,
                                        cx.listener(move |s, _, _, cx| {
                                            cx.stop_propagation();
                                            s.apply_named_table_view(id, selected, cx);
                                        }),
                                    )
                                    .child("Apply view"),
                            )
                        }),
                ),
        );
    crate::ui::modal_overlay(
        "saved-table-views",
        |s, cx| {
            s.table_dialog = None;
            cx.notify();
        },
        content,
        cx,
    )
    .into_any_element()
}
