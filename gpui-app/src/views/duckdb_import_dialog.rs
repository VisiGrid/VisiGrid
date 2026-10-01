use gpui::prelude::FluentBuilder;
use gpui::*;

use crate::app::Spreadsheet;
use crate::duckdb_import::{number, Focus, PREVIEW_COLUMNS};
use crate::theme::TokenKey;
use crate::ui::{modal_overlay, Button};

pub fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(d) = &app.duckdb_import else {
        return div().into_any_element();
    };
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let inverse = app.token(TokenKey::TextInverse);
    let warning = app.token(TokenKey::Warn);
    let width = (f32::from(app.window_size.width) - 32.0).clamp(280.0, 860.0);
    let height = (f32::from(app.window_size.height) - 48.0)
        .max(280.0)
        .min(660.0);
    let compact = width < 650.0;
    let filename = d
        .path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .into_owned();
    let can_import = d.selection_error().is_none() && d.error.is_none();
    let all = d.entries.iter().filter(|e| e.importable()).count();
    let all_checked = all > 0 && d.selected.len() == all;
    let visible = d.visible();

    let table_rows = visible
        .iter()
        .map(|&index| {
            let entry = &d.entries[index];
            let active = d.active == Some(index);
            let checked = d.selected.contains(&index);
            let enabled = entry.importable() && !d.importing;
            let detail = entry
                .info
                .as_ref()
                .map(|i| {
                    format!(
                        "{} rows · {} columns",
                        number(i.row_count),
                        number(i.columns.len() as u64)
                    )
                })
                .unwrap_or_else(|| "Metadata unavailable".into());
            div()
                .id(SharedString::from(format!("duckdb-table-{index}")))
                .w_full()
                .flex()
                .items_center()
                .gap_2()
                .p_2()
                .rounded_sm()
                .border_1()
                .border_color(if active && d.focus == Focus::Tables {
                    accent
                } else {
                    border.opacity(0.0)
                })
                .when(active, |el| el.bg(accent.opacity(0.12)))
                .child(
                    div()
                        .id(SharedString::from(format!("duckdb-check-{index}")))
                        .w(px(22.0))
                        .h(px(28.0))
                        .flex_shrink_0()
                        .flex()
                        .items_center()
                        .justify_center()
                        .when(enabled, |el| {
                            el.cursor_pointer().on_mouse_down(
                                MouseButton::Left,
                                cx.listener(move |this, _, _, cx| {
                                    this.duckdb_toggle(index, cx);
                                    if let Some(d) = &mut this.duckdb_import {
                                        d.focus = Focus::Tables;
                                    }
                                    this.duckdb_preview(index, cx);
                                    cx.stop_propagation();
                                }),
                            )
                        })
                        .opacity(if enabled { 1.0 } else { 0.4 })
                        .child(
                            div()
                                .w(px(14.0))
                                .h(px(14.0))
                                .border_1()
                                .rounded(px(2.0))
                                .border_color(if checked { accent } else { muted })
                                .bg(if checked { accent } else { bg })
                                .text_color(inverse)
                                .text_size(px(11.0))
                                .line_height(px(12.0))
                                .child(if checked { "✓" } else { "" }),
                        ),
                )
                .child(
                    div()
                        .id(SharedString::from(format!("duckdb-preview-{index}")))
                        .flex_1()
                        .min_w_0()
                        .cursor_pointer()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(move |this, _, _, cx| {
                                if let Some(d) = &mut this.duckdb_import {
                                    d.focus = Focus::Tables;
                                }
                                this.duckdb_preview(index, cx);
                            }),
                        )
                        .child(
                            div()
                                .text_size(px(13.0))
                                .text_color(if active { accent } else { text })
                                .overflow_hidden()
                                .child(entry.name.clone()),
                        )
                        .child(div().text_size(px(11.0)).text_color(muted).child(detail))
                        .when(entry.problem().is_some(), |el| {
                            el.child(div().text_size(px(11.0)).text_color(warning).child(
                                if entry.error.is_some() {
                                    "Cannot read table"
                                } else {
                                    "Exceeds import limit"
                                },
                            ))
                        }),
                )
        })
        .collect::<Vec<_>>();

    let sidebar = div()
        .w(if compact { px(width - 2.0) } else { px(235.0) })
        .flex_shrink_0()
        .flex()
        .flex_col()
        .min_h_0()
        .gap_2()
        .p_3()
        .when(compact, |el| {
            el.h(px(285.0)).border_b_1().border_color(border)
        })
        .when(!compact, |el| el.border_r_1().border_color(border))
        .child(
            div()
                .text_size(px(12.0))
                .text_color(muted)
                .child(format!("Tables · {}", d.entries.len())),
        )
        .child(
            div()
                .id("duckdb-search")
                .px_2()
                .py_2()
                .border_1()
                .rounded_sm()
                .border_color(if d.focus == Focus::Search {
                    accent
                } else {
                    border
                })
                .text_size(px(12.0))
                .text_color(if d.search.is_empty() { muted } else { text })
                .when(d.focus == Focus::Search && d.search_selected, |el| {
                    el.bg(accent.opacity(0.2))
                })
                .cursor_text()
                .overflow_hidden()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, window, cx| {
                        if let Some(d) = &mut this.duckdb_import {
                            d.focus = Focus::Search;
                        }
                        this.focus_handle.focus(window, cx);
                        cx.notify();
                    }),
                )
                .child(if d.search.is_empty() {
                    "Find a table…".into()
                } else {
                    format!(
                        "{}{}",
                        d.search,
                        if d.focus == Focus::Search { "│" } else { "" }
                    )
                }),
        )
        .child(
            div()
                .id("duckdb-select-all")
                .flex()
                .items_center()
                .gap_2()
                .py_1()
                .border_1()
                .rounded_sm()
                .border_color(if d.focus == Focus::All {
                    accent
                } else {
                    border.opacity(0.0)
                })
                .text_size(px(12.0))
                .cursor_pointer()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        if let Some(d) = &mut this.duckdb_import {
                            d.focus = Focus::All;
                        }
                        this.duckdb_toggle_all(cx);
                    }),
                )
                .child(
                    div()
                        .w(px(14.0))
                        .h(px(14.0))
                        .flex_shrink_0()
                        .rounded(px(2.0))
                        .border_1()
                        .border_color(accent)
                        .bg(if d.selected.is_empty() { bg } else { accent })
                        .text_color(inverse)
                        .text_size(px(11.0))
                        .line_height(px(12.0))
                        .child(if all_checked {
                            "✓"
                        } else if d.selected.is_empty() {
                            ""
                        } else {
                            "−"
                        }),
                )
                .child("Select importable tables"),
        )
        .child(
            div()
                .id("duckdb-table-list")
                .flex_1()
                .min_h_0()
                .overflow_y_scroll()
                .track_scroll(&d.scroll)
                .when(compact, |el| el.max_h(px(130.0)))
                .children(table_rows)
                .when(visible.is_empty() && !d.loading, |el| {
                    el.child(
                        div()
                            .py_3()
                            .text_size(px(12.0))
                            .text_color(muted)
                            .child("No matching tables"),
                    )
                }),
        )
        .child(
            div()
                .text_size(px(11.0))
                .text_color(muted)
                .child("Each selected table becomes a worksheet. Views are not included."),
        );

    let mut preview = div()
        .flex_1()
        .when(compact, |el| el.flex_none().h(px(410.0)))
        .min_w_0()
        .min_h_0()
        .p_4()
        .flex()
        .flex_col()
        .gap_3();
    if d.loading {
        preview = preview.child(
            div()
                .text_color(muted)
                .child("Reading table names and sizes…"),
        );
    } else if let Some(entry) = d.active.and_then(|i| d.entries.get(i)) {
        let detail = entry
            .info
            .as_ref()
            .map(|info| {
                let shown = d.preview.as_ref().map(|p| p.rows.len()).unwrap_or(0);
                if d.preview_loading {
                    format!(
                        "{} rows · {} columns",
                        number(info.row_count),
                        info.columns.len()
                    )
                } else {
                    format!(
                        "Preview: {} of {} rows · {} columns",
                        shown,
                        number(info.row_count),
                        info.columns.len()
                    )
                }
            })
            .unwrap_or_default();
        preview = preview.child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(14.0))
                        .font_weight(FontWeight::MEDIUM)
                        .child(entry.name.clone()),
                )
                .child(div().text_size(px(12.0)).text_color(muted).child(detail)),
        );
        if d.preview_loading {
            preview = preview.child(
                div()
                    .text_size(px(12.0))
                    .text_color(muted)
                    .child("Loading preview…"),
            );
        } else if let Some(error) = &d.preview_error {
            preview = preview.child(
                div()
                    .id("duckdb-preview-error")
                    .max_h(px(160.0))
                    .overflow_y_scroll()
                    .text_size(px(12.0))
                    .text_color(warning)
                    .child(error.clone()),
            );
        } else if let (Some(data), Some(info)) = (&d.preview, &entry.info) {
            let cell_width = 155.0;
            let grid_width = cell_width * data.columns as f32;
            let headers = info
                .columns
                .iter()
                .take(data.columns)
                .map(|c| {
                    div()
                        .w(px(cell_width))
                        .flex_shrink_0()
                        .px_2()
                        .py_2()
                        .border_r_1()
                        .border_color(border)
                        .overflow_hidden()
                        .flex()
                        .flex_col()
                        .gap(px(2.0))
                        .child(div().text_size(px(12.0)).child(c.name.clone()))
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(muted)
                                .child(c.data_type.clone()),
                        )
                })
                .collect::<Vec<_>>();
            let rows = data
                .rows
                .iter()
                .enumerate()
                .map(|(i, row)| {
                    div()
                        .flex()
                        .w(px(grid_width))
                        .h(px(29.0))
                        .border_b_1()
                        .border_color(border)
                        .when(i % 2 == 1, |el| el.bg(border.opacity(0.13)))
                        .children(row.iter().map(|value| {
                            div()
                                .w(px(cell_width))
                                .h_full()
                                .flex_shrink_0()
                                .px_2()
                                .py_1()
                                .border_r_1()
                                .border_color(border)
                                .overflow_hidden()
                                .text_size(px(12.0))
                                .child(
                                    value
                                        .chars()
                                        .take(200)
                                        .collect::<String>()
                                        .replace(['\n', '\r'], " "),
                                )
                        }))
                })
                .collect::<Vec<_>>();
            preview = preview.child(div().id("duckdb-preview-grid").min_h_0().overflow_scroll()
                .border_1().border_color(border).rounded_sm()
                .child(div().w(px(grid_width)).flex().flex_col()
                    .child(div().flex().border_b_1().border_color(border).bg(border.opacity(0.18)).children(headers))
                    .children(rows)))
                .when(data.rows.is_empty(), |el| el.child(div().text_size(px(12.0)).text_color(muted)
                    .child("No records. Import creates a sheet with column headers.")))
                .when(info.columns.len() > PREVIEW_COLUMNS, |el| el.child(div().text_size(px(11.0)).text_color(muted)
                    .child(format!("Showing the first {PREVIEW_COLUMNS} columns. Import includes all {} columns.", info.columns.len()))));
        }
        if let Some(problem) = entry.problem() {
            preview = preview.child(
                div()
                    .p_2()
                    .rounded_sm()
                    .bg(warning.opacity(0.1))
                    .text_size(px(12.0))
                    .text_color(warning)
                    .child(problem),
            );
        }
        preview = preview.child(div().text_size(px(11.0)).text_color(muted)
            .child("Large integers import as text when needed to preserve every digit. Values are copied; formulas are not created."));
    }

    let summary = if d.importing {
        "Importing selected tables…".into()
    } else if let Some(error) = &d.error {
        error.clone()
    } else if let Some(error) = d.selection_error().filter(|_| !d.loading) {
        error
    } else {
        format!(
            "{} selected · {} rows to import",
            d.selected.len(),
            number(d.selected_rows())
        )
    };
    let content = div().w(px(width)).h(px(height)).bg(bg).border_1().border_color(border).rounded_lg()
        .shadow_xl().flex().flex_col().overflow_hidden().text_color(text)
        .child(div().px_4().py_3().flex_shrink_0().border_b_1().border_color(border).flex().justify_between().gap_3()
            .child(div().flex_1().min_w_0().flex().flex_col().gap_1()
                .child(div().text_size(px(15.0)).font_weight(FontWeight::MEDIUM).child("Import from DuckDB"))
                .child(div().text_size(px(12.0)).text_color(muted).overflow_hidden().child(filename)))
            .child(div().text_size(px(11.0)).text_color(muted).child("Read-only source")))
        .child(div().id("duckdb-import-body").flex().flex_1().min_h_0()
            .when(compact, |el| el.flex_col().overflow_y_scroll())
            .when(!compact, |el| el.overflow_hidden()).child(sidebar).child(preview))
        .child(div().px_4().py_3().flex_shrink_0().border_t_1().border_color(border).flex().flex_col().gap_2()
            .child(div().id("duckdb-import-summary").max_h(px(100.0)).overflow_y_scroll().text_size(px(12.0)).text_color(if d.error.is_some() || (d.selected.len()>0 && d.selection_error().is_some() && !d.importing) { warning } else { text }).child(summary))
            .child(div().flex().items_center().justify_between().gap_2().when(compact, |el| el.flex_col().items_start())
                .child(div().flex_1().min_w_0().when(compact, |el| el.flex_none()).text_size(px(11.0)).text_color(muted).child("Copies values. Your database stays unchanged."))
                .child(div().flex().flex_shrink_0().gap_2()
                    .child(Button::new("duckdb-cancel", if d.importing { "Cancel import" } else { "Cancel" })
                        .secondary(if d.focus == Focus::Cancel { accent } else { border }, muted)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.cancel_duckdb_import(cx))))
                    .child(Button::new("duckdb-import", format!("Import {} table{}", d.selected.len(), if d.selected.len()==1 { "" } else { "s" }))
                        .disabled(!can_import).primary(accent, inverse)
                        .border_2().border_color(if d.focus == Focus::Import { text } else { accent })
                        .when(can_import, |el| el.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| this.duckdb_confirm_import(cx))))))))
        .child(div().px_4().pb_2().flex_shrink_0().text_size(px(11.0)).text_color(muted)
            .child("↑ ↓ Preview · Space Select · Tab Next control · Ctrl/Cmd+Enter Import · Esc Cancel"));
    modal_overlay(
        "duckdb-import-dialog",
        |this, cx| this.cancel_duckdb_import(cx),
        content,
        cx,
    )
    .into_any_element()
}
