//! Table-scoped sort and value filter controls; opening/cancelling is read-only.
use crate::{app::Spreadsheet, theme::TokenKey};
use gpui::{prelude::FluentBuilder, *};

fn action(
    label: &str,
    command: &'static str,
    enabled: bool,
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let hover = app.token(TokenKey::SelectionBg);
    div()
        .id(command)
        .px_3()
        .py_1()
        .rounded_sm()
        .text_sm()
        .text_color(app.token(if enabled {
            TokenKey::TextPrimary
        } else {
            TokenKey::TextMuted
        }))
        .when(enabled, |d| d.cursor_pointer().hover(move |s| s.bg(hover)))
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |s, _, _, cx| {
                cx.stop_propagation();
                if enabled {
                    s.table_menu_action(command, cx);
                }
            }),
        )
        .child(label.to_string())
}

pub(crate) fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> Option<impl IntoElement> {
    let menu = app.table_filter_dropdown.as_ref()?;
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let hover = app.token(TokenKey::SelectionBg);
    let width: f32 = app.window_size.width.into();
    let height: f32 = app.window_size.height.into();
    let left = menu.anchor.0.min((width - 308.0).max(8.0)).max(8.0);
    let top = menu.anchor.1.min((height - 530.0).max(8.0)).max(8.0);
    let query = menu.search.to_lowercase();
    let matches: Vec<_> = menu
        .values
        .iter()
        .enumerate()
        .filter(|(_, v)| v.display.to_lowercase().contains(&query))
        .collect();
    let matching_ids: Vec<_> = matches.iter().map(|(i, _)| *i).collect();
    let matching_clear = matching_ids.clone();
    let spec = app
        .sheet(cx)
        .table_view_spec()
        .filter(|s| s.table == menu.table);
    let owns_view = spec.is_some();
    let sorted = spec.is_some_and(|s| s.sort.is_some());
    let filtered = spec.is_some_and(|s| !s.filters.is_empty());
    let has_records = app
        .sheet(cx)
        .tables()
        .iter()
        .find(|t| t.id == menu.table)
        .is_some_and(|t| t.range.data_rows() > 0);
    let can_apply = !menu.limited && has_records;
    Some(div().id("table-filter-backdrop").absolute().inset_0()
        .on_mouse_down(MouseButton::Left, cx.listener(|s, _, _, cx| { s.table_filter_dropdown = None; cx.stop_propagation(); cx.notify(); }))
        .child(div().id("table-filter-menu").absolute().left(px(left)).top(px(top)).w(px(300.0))
            .max_h(px((height - top - 8.0).max(100.0))).overflow_y_scroll().flex().flex_col()
            .bg(bg).border_1().border_color(border).rounded_lg().shadow_lg().text_color(text)
            .on_mouse_down(MouseButton::Left, |_,_,cx| cx.stop_propagation())
            .child(div().px_3().pt_3().pb_2().border_b_1().border_color(border)
                .child(div().text_xs().text_color(muted).child(menu.table_name.clone()))
                .child(div().text_sm().font_weight(FontWeight::SEMIBOLD).overflow_hidden().text_ellipsis().child(menu.name.clone())))
            .child(div().p_1().flex().flex_col()
                .child(action("↑  Sort ascending", "ascending", has_records, app, cx))
                .child(action("↓  Sort descending", "descending", has_records, app, cx))
                .child(action("Clear sort", "clear-sort", sorted, app, cx)))
            .child(div().px_3().py_2().border_t_1().border_color(border).flex().flex_col().gap_2()
                .child(div().flex().justify_between().text_xs().text_color(muted)
                    .child("FILTER BY VALUES").child(format!("{} selected", menu.checked.len())))
                .child(div().px_2().py_2().rounded_sm().bg(app.token(TokenKey::AppBg)).border_1().border_color(border)
                    .text_sm().text_color(if menu.search.is_empty() { muted } else { text })
                    .child(if menu.search.is_empty() { "Type to search values…".to_string() } else { menu.search.clone() }))
                .child(div().flex().gap_3().text_xs().text_color(accent)
                    .child(div().id("table-values-all").cursor_pointer().child(if query.is_empty() { "Select all" } else { "Select matches" })
                        .on_mouse_down(MouseButton::Left, cx.listener(move |s,_,_,cx| {
                            if let Some(m) = &mut s.table_filter_dropdown { m.checked.extend(matching_ids.iter().copied()); }
                            cx.stop_propagation(); cx.notify();
                        })))
                    .child(div().id("table-values-none").cursor_pointer().child(if query.is_empty() { "Select none" } else { "Deselect matches" })
                        .on_mouse_down(MouseButton::Left, cx.listener(move |s,_,_,cx| {
                            if let Some(m) = &mut s.table_filter_dropdown { for i in &matching_clear { m.checked.remove(i); } }
                            cx.stop_propagation(); cx.notify();
                        })))))
            .child(div().id("table-value-list").h(px(160.0)).overflow_y_scroll().border_t_1().border_b_1().border_color(border)
                .when(matches.is_empty(), |d| d.child(div().p_3().text_sm().text_color(muted).child("No matching values")))
                .children(matches.into_iter().map(|(i,v)| {
                    let checked = menu.checked.contains(&i);
                    div().id(ElementId::NamedInteger("table-value".into(),i as u64)).flex().items_center().gap_2().px_3().py_1().cursor_pointer().hover(move |s| s.bg(hover))
                        .on_mouse_down(MouseButton::Left, cx.listener(move |s,_,_,cx| {
                            if let Some(m) = &mut s.table_filter_dropdown { if !m.checked.remove(&i) { m.checked.insert(i); } }
                            cx.stop_propagation(); cx.notify();
                        }))
                        .child(div().w(px(14.0)).h(px(14.0)).flex_shrink_0().border_1().border_color(if checked { accent } else { muted }).rounded_sm()
                            .flex().items_center().justify_center().text_size(px(10.0)).when(checked, |d| d.bg(accent).child("✓")))
                        .child(div().flex_1().text_sm().overflow_hidden().text_ellipsis().child(v.display.clone()))
                        .child(div().text_xs().text_color(muted).child(v.count.to_string()))
                })))
            .when(menu.limited, |d| d.child(div().px_3().py_2().text_xs().text_color(muted).child("More than 500 unique values. Sorting is available; value filtering is unavailable for this column.")))
            .when_some(menu.error.clone(), |d, error| d.child(div().px_3().py_2().text_xs().text_color(accent).child(error)))
            .child(div().p_1().flex().flex_col()
                .child(action("Clear all Table filters", "clear-filters", filtered, app, cx))
                .child(action("Clear view · enable editing", "clear-view", owns_view, app, cx)))
            .child(div().px_3().py_2().border_t_1().border_color(border).flex().items_center().gap_2()
                .child(div().flex_1().text_xs().text_color(muted).child("Esc to cancel"))
                .child(div().id("table-filter-cancel").px_2().py_1().text_sm().cursor_pointer().child("Cancel")
                    .on_mouse_down(MouseButton::Left,cx.listener(|s,_,_,cx| { s.table_filter_dropdown=None; cx.stop_propagation(); cx.notify(); })))
                .child(div().id("table-filter-apply").px_3().py_1().text_sm().rounded_sm().bg(accent).when(!can_apply,|d| d.opacity(0.4)).cursor_pointer().child("Apply")
                    .on_mouse_down(MouseButton::Left,cx.listener(move |s,_,_,cx| { if can_apply { s.table_menu_action("apply",cx); } cx.stop_propagation(); }))))))
}
