use gpui::*;
use crate::app::{Spreadsheet, CELL_HEIGHT, HEADER_WIDTH};
use crate::theme::TokenKey;

/// Render the column header row (A, B, C, ...) with resize handles
pub fn render_column_headers(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let scroll_col = app.scroll_col;
    let visible_cols = app.visible_cols();
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);

    div()
        .flex()
        .flex_shrink_0()
        .h(px(CELL_HEIGHT))
        .bg(header_bg)
        // Corner cell (empty) - can be used for select-all
        .child(
            div()
                .flex_shrink_0()
                .w(px(HEADER_WIDTH))
                .h_full()
                .border_1()
                .border_color(header_border)
        )
        // Column headers with resize handles
        .children(
            (0..visible_cols).map(move |i| {
                let col = scroll_col + i;
                let col_width = app.col_width(col);
                render_column_header(app, col, col_width, cx)
            })
        )
}

/// Render a single column header with resize handle
fn render_column_header(app: &Spreadsheet, col: usize, width: f32, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let header_text = app.token(TokenKey::HeaderTextMuted);
    let accent = app.token(TokenKey::Accent);

    div()
        .flex_shrink_0()
        .w(px(width))
        .h_full()
        .relative()
        .flex()
        .items_center()
        .justify_center()
        .border_1()
        .border_color(header_border)
        .bg(header_bg)
        .text_color(header_text)
        .text_sm()
        .child(Spreadsheet::col_letter(col))
        // Resize handle on the right edge
        .child(
            div()
                .id(ElementId::NamedInteger("col-resize".into(), col as u64))
                .absolute()
                .right_0()
                .top_0()
                .w(px(6.0))
                .h_full()
                .cursor(CursorStyle::ResizeLeftRight)
                // Hover highlight
                .hover(move |s| s.bg(accent.opacity(0.25)))
                // Mouse down to start resize, double-click to auto-fit
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                    if event.click_count == 2 {
                        // Double-click to auto-fit
                        this.auto_fit_col_width(col, cx);
                    } else {
                        // Start resize drag
                        this.resizing_col = Some(col);
                        let x: f32 = event.position.x.into();
                        this.resize_start_pos = x;
                        this.resize_start_size = this.col_width(col);
                        cx.notify();
                    }
                }))
        )
}

/// Render a row header (1, 2, 3, ...) with resize handle
pub fn render_row_header(app: &Spreadsheet, row: usize, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let row_height = app.row_height(row);
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let header_text = app.token(TokenKey::HeaderTextMuted);
    let accent = app.token(TokenKey::Accent);

    div()
        .flex_shrink_0()
        .w(px(HEADER_WIDTH))
        .h(px(row_height))
        .relative()
        .flex()
        .items_center()
        .justify_center()
        .bg(header_bg)
        .border_1()
        .border_color(header_border)
        .text_color(header_text)
        .text_sm()
        .child(format!("{}", row + 1))
        // Resize handle on the bottom edge
        .child(
            div()
                .id(ElementId::NamedInteger("row-resize".into(), row as u64))
                .absolute()
                .bottom_0()
                .left_0()
                .w_full()
                .h(px(4.0))
                .cursor(CursorStyle::ResizeUpDown)
                .hover(move |s| s.bg(accent.opacity(0.25)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                    if event.click_count == 2 {
                        // Double-click to auto-fit
                        this.auto_fit_row_height(row, cx);
                    } else {
                        // Start resize drag
                        this.resizing_row = Some(row);
                        let y: f32 = event.position.y.into();
                        this.resize_start_pos = y;
                        this.resize_start_size = this.row_height(row);
                        cx.notify();
                    }
                }))
        )
}
