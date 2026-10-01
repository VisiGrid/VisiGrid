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
        .children({
            let (row, col) = app.view_state.selected;
            let row = app.row_view.view_to_data(row);
            let column = &table.columns[col - table.range.start_col];
            let mut controls: Vec<AnyElement> = Vec::new();
            if column.formula.is_some() {
                let exception = app.sheet(cx).is_calculated_exception(row, col);
                controls.push(
                    div()
                        .text_color(app.token(TokenKey::TextMuted))
                        .child(format!(
                            "{} · {}",
                            if exception { "Override" } else { "Formula" },
                            column.name
                        ))
                        .into_any_element(),
                );
                if exception {
                    controls.push(
                        button(
                            "table-restore-formula",
                            "Restore column formula",
                            app,
                            move |s, cx| s.restore_column_formula(id, cx),
                            cx,
                        )
                        .into_any_element(),
                    );
                }
                controls.push(
                    button(
                        "table-edit-formula",
                        "Edit column formula",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::ColumnFormula(id, col, false), cx)
                        },
                        cx,
                    )
                    .into_any_element(),
                );
                controls.push(
                    button(
                        "table-replace-formula",
                        "Replace entire column",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::ColumnFormula(id, col, true), cx)
                        },
                        cx,
                    )
                    .into_any_element(),
                );
            } else if row > table.range.start_row
                && app.sheet(cx).get_raw(row, col).starts_with('=')
            {
                controls.push(
                    button(
                        "table-use-formula",
                        "Use formula for entire column",
                        app,
                        move |s, cx| {
                            s.open_table_dialog(TableDialogKind::ColumnFormula(id, col, true), cx)
                        },
                        cx,
                    )
                    .into_any_element(),
                );
            }
            controls
        })
        .child(button(
            "table-add-row",
            "Add row",
            app,
            move |s, cx| s.add_table_row(id, cx),
            cx,
        ))
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
        TableDialogKind::ColumnFormula(_, _, true) => "Use formula for entire column",
        TableDialogKind::ColumnFormula(_, _, false) => "Edit column formula",
    };
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let creating = d.kind == TableDialogKind::Create;
    let mut fields = div().flex().gap_3().when(!creating, |s| s.flex_col());
    for (index, label, value) in [(0, "Table name", &d.name), (1, "Range", &d.range)] {
        let label = if matches!(d.kind, TableDialogKind::ColumnFormula(..)) {
            "Column formula"
        } else {
            label
        };
        let show = match d.kind {
            TableDialogKind::Create => true,
            TableDialogKind::Rename(_) | TableDialogKind::ColumnFormula(..) => index == 0,
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
                .flex_1().min_w_0()
                .gap_2()
                .child(div().text_size(px(12.0)).text_color(muted).child(if creating && index == 1 { "Source range" } else { label }))
                .child(
                    div()
                        .id(("table-field", index))
                        .border_1()
                        .rounded_sm()
                        .border_color(if active { accent } else { border })
                        .px_3()
                        .py_2()
                        .text_color(text)
                        .min_h(px(38.0))
                        .bg(app.token(TokenKey::EditorBg))
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
                            div().overflow_hidden().text_ellipsis()
                                .when(active && d.select_all, |s| {
                                    s.bg(app.token(TokenKey::SelectionBg))
                                })
                                .child(value.clone()),
                        ),
                ),
        );
    }
    if d.kind == TableDialogKind::Create {
        fields = div().flex().flex_col().gap_4().child(fields).child(div()
            .id("table-has-headers")
            .flex().items_center().gap_3().p_3().rounded_md().border_1()
            .bg(app.token(TokenKey::EditorBg))
            .border_color(if d.field == 2 { accent } else { border })
            .cursor_pointer().text_color(text)
            .on_mouse_down(MouseButton::Left, cx.listener(|s, _, _, cx| {
                cx.stop_propagation();
                if let Some(d) = s.table_dialog.as_mut() {
                    d.has_headers = !d.has_headers;
                    d.field = 2;
                    d.error = None;
                }
                cx.notify();
            }))
            .child(div().size(px(14.0)).flex().items_center().justify_center()
                .border_1().rounded_sm().border_color(if d.has_headers { accent } else { border })
                .text_size(px(11.0)).text_color(accent)
                .child(if d.has_headers { "✓" } else { "" }))
            .child(div().flex().flex_col().gap_1()
                .child(div().text_size(px(13.0)).font_weight(FontWeight::MEDIUM).child("My data has headers"))
                .child(div().text_size(px(12.0)).text_color(muted).child(if d.has_headers {
                    "Use the first row as column names."
                } else { "Keep every selected row as data." }))));
    }
    let mut preview = div()
        .flex()
        .flex_col()
        .gap_1()
        .text_size(px(12.0))
        .text_color(muted);
    match d.kind {
        TableDialogKind::Create => {
            match parse_range(&d.range).and_then(|r| app.wb(cx)
                .preview_table_creation(d.sheet, r, d.has_headers).map(|(result, headers)| (r, result, headers))) {
                Ok((source, range, headers)) => {
                    let record_count = range.data_rows();
                    let shown_columns = headers.len().min(3);
                    let sheet = app.wb(cx).sheet_by_id(d.sheet).unwrap();
                    let cell = |value: String, header: bool| div()
                        .flex_1().min_w_0().px_3().py_2()
                        .overflow_hidden().text_ellipsis()
                        .text_color(if header { text } else { muted })
                        .when(header, |s| s.font_weight(FontWeight::MEDIUM))
                        .child(value);
                    let mut grid = div().rounded_md().border_1().border_color(border)
                        .overflow_hidden().bg(app.token(TokenKey::EditorBg))
                        .child(div().flex().bg(accent.opacity(0.10))
                            .children(headers.iter().take(shown_columns).map(|name| cell(name.clone(), true))));
                    for offset in 0..record_count.min(2) {
                        let row = source.start_row + usize::from(d.has_headers) + offset;
                        grid = grid.child(div().flex().border_t_1().border_color(border)
                            .children((0..shown_columns).map(|col| cell(sheet.get_display(row, source.start_col + col), false))));
                    }
                    if record_count == 0 {
                        grid = grid.child(div().px_3().py_2().border_t_1().border_color(border)
                            .child("No records yet. Add rows after creating the Table."));
                    }
                    preview = preview.gap_3()
                        .child(div().flex().items_center().justify_between()
                            .child(div().text_color(text).font_weight(FontWeight::MEDIUM).child("Preview"))
                            .child(format!("{} {} · {} {}", record_count, if record_count == 1 { "record" } else { "records" },
                                headers.len(), if headers.len() == 1 { "column" } else { "columns" })))
                        .child(grid)
                        .child(div().flex().flex_wrap().gap_2().justify_between()
                            .child(format!("Table range  {}", range_label(range)))
                            .when(headers.len() > shown_columns || record_count > 2, |s| s.child(format!(
                                "Preview: {} columns · {} records", shown_columns, record_count.min(2)))));
                    let changed: Vec<String> = headers.iter().enumerate().filter_map(|(offset, name)| {
                        let old = sheet.get_display(source.start_row, source.start_col + offset);
                        (d.has_headers && old != *name).then(|| format!("{} → {name}", if old.is_empty() { "(blank)" } else { &old }))
                    }).collect();
                    if !changed.is_empty() {
                        preview = preview.child(div().overflow_hidden().text_ellipsis().child(format!("Column names adjusted: {}{}", changed.iter().take(3).cloned().collect::<Vec<_>>().join(", "),
                            if changed.len() > 3 { " …" } else { "" })));
                    }
                    if !d.has_headers {
                        preview = preview.child(div().p_3().rounded_md().bg(accent.opacity(0.06))
                            .border_1().border_color(border).flex().flex_col().gap_1()
                            .child(div().text_color(text).font_weight(FontWeight::MEDIUM).child(format!("Insert a header row at row {}", source.start_row + 1)))
                            .child("All cells on this worksheet at or below that row move down, including data outside the Table. Column names are generated automatically."));
                    }
                }
                Err(e) => preview = preview.child(div().p_3().rounded_md().border_1()
                    .border_color(app.token(TokenKey::Error).opacity(0.4))
                    .text_color(app.token(TokenKey::Error)).child(e)),
            }
        }
        TableDialogKind::Rename(_)=>preview=preview.child("Formulas that reference this Table will follow the new name."),
        TableDialogKind::Resize(_)=>preview=preview.child("Keep the top-left cell fixed. Cells released by shrinking stay in place; references to removed columns become #REF!."),
        TableDialogKind::ColumnFormula(id,col,replace) => {
            if let Some((sheet,table)) = app.wb(cx).table(id) {
                let sheet = app.wb(cx).sheet_by_id(sheet).unwrap();
                let exceptions = (table.range.start_row+1..=table.range.end_row).filter(|r|sheet.is_calculated_exception(*r,col)).count();
                let populated = (table.range.start_row+1..=table.range.end_row).filter(|r|!sheet.get_raw(*r,col).is_empty()).count();
                preview = preview.child(format!("{} · {} · {} records",table.name,table.columns[col-table.range.start_col].name,table.range.data_rows()));
                preview = preview.child(if replace {format!("Replace {populated} existing values/formulas and fill all {} records, including {exceptions} overrides. One undo step.",table.range.data_rows())}
                    else {format!("Update {} formula cells. Preserve {exceptions} overrides, including cleared cells.",table.range.data_rows()-exceptions)});
                preview = preview.child(format!("Formula shown at row {}. New rows use this rule; cell edits remain overrides.", d.range.parse::<usize>().unwrap_or(0)+1));
            }
        }
        TableDialogKind::Convert(_)=>preview=preview.child(format!("Convert {} ({}) to ordinary cells? Structured references become fixed cell references. Table banding disappears; explicit formatting is kept. You can undo this change.",d.name,d.range)),
    }
    let content = div()
        .w(px(if creating { 560.0 } else { 470.0 }))
        .max_h(px(720.0))
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(border)
        .rounded_md()
        .p_6()
        .shadow_lg()
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
        .when(creating, |s| s.child(div().mt(px(-8.0)).text_size(px(13.0)).text_color(muted)
            .child("Organize your data with named columns and room to grow.")))
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
            div().flex().flex_wrap().gap_3().items_center().justify_between().pt_4().border_t_1().border_color(border)
                .child(div().text_size(px(11.0)).text_color(muted).child("Enter to apply · Esc to cancel"))
                .child(div().flex().items_center().gap_2()
                    .child(button("table-dialog-cancel", "Cancel", app, |s, cx| {
                        s.table_dialog = None;
                        cx.notify();
                    }, cx))
                    .child(div().id("table-dialog-submit").px_4().py_2().rounded_md()
                        .bg(accent).text_color(app.token(TokenKey::TextInverse))
                        .text_size(px(13.0)).font_weight(FontWeight::SEMIBOLD).cursor_pointer()
                        .hover(|s| s.bg(accent.opacity(0.85)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|s, _, _, cx| {
                            cx.stop_propagation();
                            s.submit_table_dialog(cx);
                        }))
                        .child(title))),
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
