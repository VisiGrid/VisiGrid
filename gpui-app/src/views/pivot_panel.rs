//! Pivot table field list — right-side drawer.
//!
//! Keyboard first: ↑/↓ move through the source fields and the Rows, Column
//! and Values wells; R/C/V assign the field under the cursor; ←/→ change a
//! value field's aggregation; = applies it to every value field; Alt+↑/↓
//! reorder; Delete removes; X includes rows appended below the source; Enter
//! applies; Esc closes. The mouse works too. Nothing changes in the workbook
//! until Apply, and each Apply is one undo step.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::pivot_ui::{PivotPanelItem, PivotPanelMode};
use crate::theme::TokenKey;

pub(crate) const PANEL_WIDTH: f32 = 340.0;

pub(crate) fn render_pivot_panel(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);
    let editor_bg = app.token(TokenKey::EditorBg);

    let Some(panel) = app.pivot_panel.clone() else {
        return div().id("pivot-panel-empty");
    };
    let items = panel.items();
    let cursor_item = items.get(panel.cursor).copied();

    // Title, staleness and source line.
    let (title, stale) = match &panel.mode {
        PivotPanelMode::New => ("New pivot table".to_string(), false),
        PivotPanelMode::Edit { pivot_id } => match app.wb(cx).find_pivot(*pivot_id) {
            Some((_, t)) => (format!("{} fields", t.name), app.wb(cx).is_pivot_stale(t)),
            None => ("Pivot table".to_string(), false),
        },
    };
    let source_sheet = app
        .wb(cx)
        .sheet_by_id(panel.source.sheet_id)
        .map(|s| s.name.clone())
        .unwrap_or_else(|| "?".into());
    let source_line = format!(
        "Source: {}!{} · {} row{}",
        source_sheet,
        panel.source_label(),
        panel.source.data_rows(),
        if panel.source.data_rows() == 1 { "" } else { "s" }
    );

    let row = |id: SharedString, label: String, detail: Option<String>, item: PivotPanelItem, cx: &mut Context<Spreadsheet>| {
        let selected = cursor_item == Some(item);
        div()
            .id(id)
            .flex()
            .items_center()
            .justify_between()
            .px_3()
            .py(px(3.0))
            .mx_2()
            .rounded_sm()
            .cursor_pointer()
            .text_sm()
            .text_color(text_primary)
            .when(selected, |d| d.bg(accent.opacity(0.18)))
            .hover(|s| s.bg(accent.opacity(0.08)))
            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                cx.stop_propagation();
                if let Some(p) = this.pivot_panel.as_mut() {
                    if let Some(i) = p.items().iter().position(|x| *x == item) {
                        p.cursor = i;
                    }
                }
                cx.notify();
            }))
            .child(div().overflow_hidden().child(label))
            .when_some(detail, |d, t| d.child(div().text_xs().text_color(text_muted).child(t)))
    };

    let section = |title: &str| {
        div()
            .px_3()
            .pt_2()
            .pb_1()
            .text_xs()
            .font_weight(FontWeight::SEMIBOLD)
            .text_color(text_muted)
            .child(title.to_string())
    };

    // Source fields, with where each is used.
    let mut fields = div().flex().flex_col();
    for (i, h) in panel.headers.iter().enumerate() {
        let mut tags = Vec::new();
        if panel.draft.rows.iter().any(|f| f.offset as usize == i) {
            tags.push("row");
        }
        if panel.draft.column.as_ref().is_some_and(|f| f.offset as usize == i) {
            tags.push("column");
        }
        if panel.draft.values.iter().any(|v| v.field.offset as usize == i) {
            tags.push("value");
        }
        let detail = (!tags.is_empty()).then(|| tags.join(", "));
        fields = fields.child(row(format!("pivot-field-{i}").into(), h.clone(), detail, PivotPanelItem::Field(i), cx));
    }

    let mut rows_well = div().flex().flex_col();
    for (k, f) in panel.draft.rows.iter().enumerate() {
        rows_well = rows_well.child(row(format!("pivot-row-{k}").into(), f.header.clone(), None, PivotPanelItem::Row(k), cx));
    }
    let mut column_well = div().flex().flex_col();
    if let Some(f) = &panel.draft.column {
        column_well = column_well.child(row("pivot-column".into(), f.header.clone(), None, PivotPanelItem::Column, cx));
    }
    let mut values_well = div().flex().flex_col();
    for (k, v) in panel.draft.values.iter().enumerate() {
        values_well = values_well.child(row(
            format!("pivot-value-{k}").into(),
            v.header(),
            Some("\u{2190} \u{2192}".into()),
            PivotPanelItem::Value(k),
            cx,
        ));
    }
    let empty_hint = |text: &str| div().px_3().py(px(2.0)).text_xs().text_color(text_muted.opacity(0.7)).child(text.to_string());

    div()
        .id("pivot-panel")
        .absolute()
        .right_0()
        .top_0()
        .h_full()
        .w(px(PANEL_WIDTH))
        .bg(panel_bg)
        .border_l_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_mouse_up(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        // Header
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .px_3()
                .py_2()
                .border_b_1()
                .border_color(panel_border)
                .child(div().text_sm().font_weight(FontWeight::SEMIBOLD).text_color(text_primary).child(title))
                .child(
                    div()
                        .id("pivot-panel-close")
                        .px_2()
                        .cursor_pointer()
                        .text_color(text_muted)
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            cx.stop_propagation();
                            this.close_pivot_panel(cx);
                        }))
                        .child("\u{2715}"),
                ),
        )
        // Source and freshness
        .child(div().px_3().pt_2().text_xs().text_color(text_muted).child(source_line))
        .when(stale, |d| {
            d.child(div().px_3().text_xs().text_color(error_color).child("Out of date: the source changed. Enter or Alt+F5 refreshes."))
        })
        .when_some(panel.growth, |d, last| {
            let n = last - panel.source.end_row;
            d.child(div().px_3().text_xs().text_color(accent).child(format!(
                "{n} new row{} below the source. X includes {}.",
                if n == 1 { "" } else { "s" },
                if n == 1 { "it" } else { "them" }
            )))
        })
        // Lists
        .child(
            div()
                .id("pivot-panel-lists")
                .flex_1()
                .overflow_y_scroll()
                .flex()
                .flex_col()
                .child(section("FIELDS"))
                .child(fields)
                .child(section("ROWS"))
                .when(panel.draft.rows.is_empty(), |d| d.child(empty_hint("Press R on a field")))
                .child(rows_well)
                .child(section("COLUMN"))
                .when(panel.draft.column.is_none(), |d| d.child(empty_hint("Press C on a field (optional)")))
                .child(column_well)
                .child(section("VALUES"))
                .when(panel.draft.values.is_empty(), |d| d.child(empty_hint("Press V on a field")))
                .child(values_well),
        )
        // Message
        .when_some(panel.message.clone(), |d, m| {
            d.child(div().px_3().py_1().text_xs().text_color(text_primary).child(m))
        })
        // Apply
        .child(
            div()
                .id("pivot-panel-apply")
                .mx_3()
                .my_2()
                .px_2()
                .py_1()
                .rounded_sm()
                .border_1()
                .border_color(accent.opacity(0.6))
                .bg(editor_bg)
                .cursor_pointer()
                .text_sm()
                .text_center()
                .text_color(text_primary)
                .when(panel.busy, |d| d.opacity(0.5))
                .hover(|s| s.border_color(accent))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.pivot_apply(cx);
                }))
                .child(match (&panel.mode, panel.busy) {
                    (_, true) => "Computing\u{2026}",
                    (PivotPanelMode::New, _) => "Create pivot table (Enter)",
                    (PivotPanelMode::Edit { .. }, _) => "Apply (Enter)",
                }),
        )
        // Keys
        .child(
            div()
                .px_3()
                .pb_2()
                .text_xs()
                .text_color(text_muted.opacity(0.8))
                .child("\u{2191}\u{2193} move \u{00b7} R row \u{00b7} C column \u{00b7} V value \u{00b7} \u{2190}\u{2192} aggregation \u{00b7} = all \u{00b7} Alt+\u{2191}\u{2193} reorder \u{00b7} Del remove \u{00b7} Esc close"),
        )
}
