use crate::{
    app::Spreadsheet,
    table_ui::{parse_range, range_label, TableDialogKind},
    theme::TokenKey,
};
use gpui::prelude::FluentBuilder;
use gpui::*;

fn button(
    id: &'static str,
    label: impl Into<SharedString>,
    app: &Spreadsheet,
    action: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .id(id)
        .px_2()
        .py_1()
        .rounded_sm()
        .cursor_pointer()
        .text_color(app.token(TokenKey::TextPrimary))
        .hover(|s| s.bg(app.token(TokenKey::SelectionBg)))
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |this, _, _, cx| {
                cx.stop_propagation();
                action(this, cx);
            }),
        )
        .child(label.into())
}

pub(crate) fn render_table_controls(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> AnyElement {
    let Some(table) = app.table_under_cursor(cx) else {
        return div().into_any_element();
    };
    let id = table.id;
    div()
        .id("table-controls")
        .h(px(crate::table_ui::TABLE_CONTROLS_HEIGHT))
        .flex()
        .items_center()
        .gap_2()
        .px_3()
        .py_1()
        .flex_shrink_0()
        .bg(app.token(TokenKey::PanelBg))
        .border_b_1()
        .border_color(app.token(TokenKey::PanelBorder))
        .text_size(px(12.0))
        .child(
            div()
                .max_w(px(180.0))
                .overflow_hidden()
                .text_ellipsis()
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(app.token(TokenKey::TextPrimary))
                .child(format!("Table · {}", table.name)),
        )
        .child(
            div()
                .max_w(px(190.0))
                .overflow_hidden()
                .text_ellipsis()
                .text_color(app.token(TokenKey::TextMuted))
                .child(format!(
                    "{} · {} records",
                    range_label(table.range),
                    table.range.data_rows()
                )),
        )
        .child(div().flex_1())
        .child(button(
            "table-rename",
            "Rename",
            app,
            move |s, cx| s.open_table_dialog(TableDialogKind::Rename(id), cx),
            cx,
        ))
        .child(button(
            "table-resize",
            "Resize",
            app,
            move |s, cx| s.open_table_dialog(TableDialogKind::Resize(id), cx),
            cx,
        ))
        .child(button(
            "table-banding",
            if table.style.banded_rows {
                "✓ Banded rows"
            } else {
                "Banded rows"
            },
            app,
            move |s, cx| s.toggle_table_banding(id, cx),
            cx,
        ))
        .child(button(
            "table-convert",
            "Convert to range…",
            app,
            move |s, cx| s.open_table_dialog(TableDialogKind::Convert(id), cx),
            cx,
        ))
        .into_any_element()
}

pub(crate) fn render_table_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(d) = &app.table_dialog else {
        return div().into_any_element();
    };
    let title = match d.kind {
        TableDialogKind::Create => "Create Table",
        TableDialogKind::Rename(_) => "Rename Table",
        TableDialogKind::Resize(_) => "Resize Table",
        TableDialogKind::Convert(_) => "Convert to range",
    };
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let mut fields = div().flex().flex_col().gap_3();
    for (index, label, value) in [(0, "Table name", &d.name), (1, "Range", &d.range)] {
        let show = match d.kind {
            TableDialogKind::Create => true,
            TableDialogKind::Rename(_) => index == 0,
            TableDialogKind::Resize(_) => index == 1,
            TableDialogKind::Convert(_) => false,
        };
        if !show {
            continue;
        }
        let active = d.field == index;
        fields = fields.child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(div().text_size(px(12.0)).text_color(muted).child(label))
                .child(
                    div()
                        .id(("table-field", index))
                        .border_1()
                        .rounded_sm()
                        .border_color(if active { accent } else { border })
                        .px_3()
                        .py_2()
                        .text_color(text)
                        .min_h(px(34.0))
                        .cursor_text()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(move |s, _, _, cx| {
                                cx.stop_propagation();
                                if let Some(d) = s.table_dialog.as_mut() {
                                    d.field = index;
                                    d.select_all = true;
                                }
                                cx.notify();
                            }),
                        )
                        .child(
                            div()
                                .when(active && d.select_all, |s| {
                                    s.bg(app.token(TokenKey::SelectionBg))
                                })
                                .child(value.clone()),
                        ),
                ),
        );
    }
    let mut preview = div()
        .flex()
        .flex_col()
        .gap_1()
        .text_size(px(12.0))
        .text_color(muted);
    match d.kind {
        TableDialogKind::Create => {
            preview = preview.child("The first row supplies the column headers. Your data stays in place.");
            match parse_range(&d.range).and_then(|r|app.wb(cx).preview_table_headers(d.sheet,r).map(|h|(r,h))) {
                Ok((range,headers)) => {
                    preview = preview.child(format!("{} records · {} columns",range.data_rows(),headers.len()));
                    for (offset,name) in headers.iter().enumerate().take(6) {
                        let old=app.wb(cx).sheet_by_id(d.sheet).map(|s|s.get_display(range.start_row,range.start_col+offset)).unwrap_or_default();
                        preview=preview.child(if old==*name {name.clone()} else {format!("{} → {name}",if old.is_empty(){"(blank)"}else{&old})});
                    }
                    if headers.len()>6 {preview=preview.child(format!("…and {} more columns",headers.len()-6));}
                }
                Err(e)=>preview=preview.child(div().text_color(app.token(TokenKey::Error)).child(e)),
            }
        }
        TableDialogKind::Rename(_)=>preview=preview.child("Formulas that reference this Table will follow the new name."),
        TableDialogKind::Resize(_)=>preview=preview.child("Keep the top-left cell fixed. Cells released by shrinking stay in place; references to removed columns become #REF!."),
        TableDialogKind::Convert(_)=>preview=preview.child(format!("Convert {} ({}) to ordinary cells? Structured references become fixed cell references. Table banding disappears; explicit formatting is kept. You can undo this change.",d.name,d.range)),
    }
    let content = div()
        .w(px(470.0))
        .max_h(px(620.0))
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(border)
        .rounded_md()
        .p_5()
        .flex()
        .flex_col()
        .gap_4()
        .child(
            div()
                .text_size(px(18.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text)
                .child(title),
        )
        .child(fields)
        .child(preview)
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
                .gap_2()
                .justify_end()
                .child(button(
                    "table-dialog-cancel",
                    "Cancel",
                    app,
                    |s, cx| {
                        s.table_dialog = None;
                        cx.notify();
                    },
                    cx,
                ))
                .child(button(
                    "table-dialog-submit",
                    title,
                    app,
                    |s, cx| s.submit_table_dialog(cx),
                    cx,
                )),
        )
        .child(
            div()
                .text_size(px(11.0))
                .text_color(muted)
                .child("Tab to switch fields · Enter to apply · Esc to cancel"),
        );
    crate::ui::modal_overlay(
        "table-dialog",
        |s, cx| {
            s.table_dialog = None;
            cx.notify();
        },
        content,
        cx,
    )
    .into_any_element()
}
