//! Compact, zoom-aware Table controls. Icons do not depend on the UI font.
use crate::{app::Spreadsheet, theme::TokenKey};
use gpui::{prelude::FluentBuilder, *};
use visigrid_engine::{
    filter::SortDirection,
    table::{TableColumnId, TableId},
};

#[derive(Clone, Copy)]
enum Mark {
    Chevron,
    Ascending,
    Descending,
    Filter,
}

fn mark(kind: Mark, color: Hsla, zoom: f32) -> impl IntoElement {
    canvas(
        |_, _, _| (),
        move |bounds, _, window, _| {
            let p = |x, y| bounds.origin + point(px(x * zoom), px(y * zoom));
            let mut path = PathBuilder::stroke(px(1.4 * zoom));
            match kind {
                Mark::Chevron => {
                    path.move_to(p(2.5, 4.5));
                    path.line_to(p(6.0, 8.0));
                    path.line_to(p(9.5, 4.5));
                }
                Mark::Ascending | Mark::Descending => {
                    let (tip, tail) = if matches!(kind, Mark::Ascending) {
                        (2.0, 10.0)
                    } else {
                        (10.0, 2.0)
                    };
                    let shoulder = if matches!(kind, Mark::Ascending) {
                        5.0
                    } else {
                        7.0
                    };
                    path.move_to(p(6.0, tail));
                    path.line_to(p(6.0, tip));
                    path.move_to(p(2.8, shoulder));
                    path.line_to(p(6.0, tip));
                    path.line_to(p(9.2, shoulder));
                }
                Mark::Filter => {
                    path.move_to(p(1.5, 2.5));
                    path.line_to(p(10.5, 2.5));
                    path.line_to(p(7.0, 6.5));
                    path.line_to(p(7.0, 10.0));
                    path.line_to(p(5.0, 9.0));
                    path.line_to(p(5.0, 6.5));
                    path.close();
                }
            }
            if let Ok(path) = path.build() {
                window.paint_path(path, color);
            }
        },
    )
    .size(px(12.0 * zoom))
    .flex_shrink_0()
}

pub(super) fn width(sort: Option<SortDirection>, filtered: bool, zoom: f32) -> f32 {
    (24.0 + if sort.is_some() { 12.0 } else { 0.0 } + if filtered { 12.0 } else { 0.0 }) * zoom
}

pub(super) struct HeaderControl {
    pub table: TableId,
    pub col: usize,
    pub field: Option<TableColumnId>,
    pub label: String,
    pub sort: Option<SortDirection>,
    pub filtered: bool,
}

pub(super) fn render(
    app: &Spreadsheet,
    control: HeaderControl,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let HeaderControl {
        table,
        col,
        field,
        label,
        sort,
        filtered,
    } = control;
    let zoom = app.metrics.zoom;
    let accent = app.token(TokenKey::Accent);
    let ink = app.token(TokenKey::CellText);
    let active = sort.is_some() || filtered;
    let open = app
        .table_filter_dropdown
        .as_ref()
        .is_some_and(|m| m.table == table && Some(m.column) == field);
    let color = if active || open {
        accent
    } else {
        ink.opacity(0.72)
    };
    let border = if active || open {
        accent.opacity(0.35)
    } else {
        ink.opacity(0.14)
    };
    let bg = if open {
        accent.opacity(0.2)
    } else if active {
        accent.opacity(0.1)
    } else {
        transparent_black()
    };
    let mut state = Vec::new();
    if let Some(direction) = sort {
        state.push(if direction == SortDirection::Ascending {
            "Sorted ascending"
        } else {
            "Sorted descending"
        });
    }
    if filtered {
        state.push("Filtered");
    }
    let state = state.join(" · ");
    let tooltip_bg = app.token(TokenKey::PanelBg);
    let tooltip_border = app.token(TokenKey::PanelBorder);
    let tooltip_ink = app.token(TokenKey::TextPrimary);
    let tooltip_muted = app.token(TokenKey::TextMuted);
    div()
        .id(ElementId::Name(
            format!("table-header-menu-{}-{col}", table.0).into(),
        ))
        .absolute()
        .right(px(2.0 * zoom))
        .top(px(2.0 * zoom))
        .bottom(px(2.0 * zoom))
        .w(px(width(sort, filtered, zoom)))
        .flex()
        .items_center()
        .justify_center()
        .rounded(px(4.0 * zoom))
        .border_1()
        .border_color(border)
        .bg(bg)
        .cursor_pointer()
        .hover(move |s| s.bg(accent.opacity(0.18)).border_color(accent.opacity(0.5)))
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |s, event: &MouseDownEvent, window, cx| {
                cx.stop_propagation();
                s.focus_handle.focus(window, cx);
                s.open_table_filter(
                    table,
                    col,
                    (
                        event.position.x.into(),
                        f32::from(event.position.y) + 14.0 * zoom,
                    ),
                    cx,
                );
            }),
        )
        .tooltip(move |_, cx| {
            cx.new(|_| HeaderTooltip {
                label: label.clone().into(),
                state: state.clone().into(),
                bg: tooltip_bg,
                border: tooltip_border,
                ink: tooltip_ink,
                muted: tooltip_muted,
            })
            .into()
        })
        .when_some(sort, |d, sort| {
            d.child(mark(
                if sort == SortDirection::Ascending {
                    Mark::Ascending
                } else {
                    Mark::Descending
                },
                color,
                zoom,
            ))
        })
        .when(filtered, |d| d.child(mark(Mark::Filter, color, zoom)))
        .child(mark(Mark::Chevron, color, zoom))
}

struct HeaderTooltip {
    label: SharedString,
    state: SharedString,
    bg: Hsla,
    border: Hsla,
    ink: Hsla,
    muted: Hsla,
}
impl Render for HeaderTooltip {
    fn render(&mut self, _: &mut Window, _: &mut Context<Self>) -> impl IntoElement {
        div()
            .px_3()
            .py_2()
            .rounded_md()
            .border_1()
            .border_color(self.border)
            .bg(self.bg)
            .shadow_md()
            .flex()
            .flex_col()
            .gap_1()
            .text_size(px(12.0))
            .child(
                div()
                    .font_weight(FontWeight::MEDIUM)
                    .text_color(self.ink)
                    .child(self.label.clone()),
            )
            .when(!self.state.is_empty(), |d| {
                d.child(div().text_color(self.muted).child(self.state.clone()))
            })
            .child(
                div()
                    .text_color(self.muted)
                    .child("Sort and filter · Alt+↓"),
            )
    }
}
