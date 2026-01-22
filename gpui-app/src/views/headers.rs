use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the column header row (A, B, C, ...) with resize handles
pub fn render_column_headers(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let scroll_col = app.scroll_col;
    let visible_cols = app.visible_cols();
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let metrics = &app.metrics;

    div()
        .flex()
        .flex_shrink_0()
        .h(px(metrics.header_h))  // Scaled header height
        .bg(header_bg)
        // Corner cell - click to select all
        .child(
            div()
                .id("select-all-corner")
                .flex_shrink_0()
                .w(px(metrics.header_w))  // Scaled header width
                .h_full()
                .border_1()
                .border_color(header_border)
                .cursor_pointer()
                .hover(|s| s.bg(selection_bg.opacity(0.3)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.select_all(cx);
                }))
        )
        // Column headers with resize handles
        .children(
            (0..visible_cols).map(move |i| {
                let col = scroll_col + i;
                // Pass scaled width for rendering
                let col_width = metrics.col_width(app.col_width(col));
                let is_selected = app.is_col_header_selected(col);
                render_column_header(app, col, col_width, is_selected, cx)
            })
        )
}

/// Render a single column header with resize handle and selection support
fn render_column_header(
    app: &Spreadsheet,
    col: usize,
    width: f32,
    is_selected: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let header_text = app.token(TokenKey::HeaderTextMuted);
    let accent = app.token(TokenKey::Accent);
    let selection_bg = app.token(TokenKey::SelectionBg);

    div()
        .id(ElementId::NamedInteger("col-header".into(), col as u64))
        .flex_shrink_0()
        .w(px(width))
        .h_full()
        .relative()
        .flex()
        .items_center()
        .justify_center()
        .border_1()
        .border_color(header_border)
        .when(is_selected, |div| div.bg(selection_bg.opacity(0.5)))
        .when(!is_selected, |div| div.bg(header_bg))
        .text_color(header_text)
        .text_sm()
        .cursor_pointer()
        .hover(|s| s.bg(selection_bg.opacity(0.3)))
        .child(Spreadsheet::col_letter(col))
        // Click handler for column selection
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Check if click is in the resize handle area (last 6px of column)
            // Skip selection handling if so - let the resize handle deal with it
            let click_x: f32 = event.position.x.into();
            // Use scaled header width (col_x_offset already returns scaled value)
            let col_start_x = this.metrics.header_w + this.col_x_offset(col);
            let col_end_x = col_start_x + width; // width is already scaled from caller
            let resize_area_start = col_end_x - 6.0;

            if click_x >= resize_area_start {
                // Click is on resize handle, don't change selection
                return;
            }

            if event.modifiers.shift {
                // Shift+click: extend selection
                this.select_col(col, true, cx);
            } else if event.modifiers.control || event.modifiers.platform {
                // Ctrl+click: add to selection
                this.ctrl_click_col(col, cx);
            } else {
                // Regular click: start drag selection
                this.start_col_header_drag(col, cx);
            }
        }))
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
                        // Double-click to auto-fit (all selected if part of selection)
                        this.auto_fit_selected_col_widths(col, cx);
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

/// Render a row header (1, 2, 3, ...) with resize handle and selection support
pub fn render_row_header(app: &Spreadsheet, row: usize, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    // Get scaled row height for rendering
    let row_height_scaled = app.metrics.row_height(app.row_height(row));
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let header_text = app.token(TokenKey::HeaderTextMuted);
    let accent = app.token(TokenKey::Accent);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let is_selected = app.is_row_header_selected(row);

    div()
        .id(ElementId::NamedInteger("row-header".into(), row as u64))
        .flex_shrink_0()
        .w(px(app.metrics.header_w))  // Scaled header width
        .h(px(row_height_scaled))     // Scaled row height
        .relative()
        .flex()
        .items_center()
        .justify_center()
        .when(is_selected, |div| div.bg(selection_bg.opacity(0.5)))
        .when(!is_selected, |div| div.bg(header_bg))
        .border_1()
        .border_color(header_border)
        .text_color(header_text)
        .text_size(px(app.metrics.font_size))  // Scaled font size
        .cursor_pointer()
        .hover(|s| s.bg(selection_bg.opacity(0.3)))
        .child(format!("{}", row + 1))
        // Click handler for row selection
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Check if click is in the resize handle area (bottom 4px of row)
            // Skip selection handling if so - let the resize handle deal with it
            let click_y: f32 = event.position.y.into();
            // row_y_offset returns scaled value
            let row_start_y = this.grid_layout.grid_body_origin.1 + this.row_y_offset(row);
            let row_end_y = row_start_y + this.metrics.row_height(this.row_height(row));
            let resize_area_start = row_end_y - 4.0;

            if click_y >= resize_area_start {
                // Click is on resize handle, don't change selection
                return;
            }

            if event.modifiers.shift {
                // Shift+click: extend selection
                this.select_row(row, true, cx);
            } else if event.modifiers.control || event.modifiers.platform {
                // Ctrl+click: add to selection
                this.ctrl_click_row(row, cx);
            } else {
                // Regular click: start drag selection
                this.start_row_header_drag(row, cx);
            }
        }))
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
                        // Double-click to auto-fit (all selected if part of selection)
                        this.auto_fit_selected_row_heights(row, cx);
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
