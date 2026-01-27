use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the filter dropdown button for a column header cell
/// Shows only when AutoFilter is enabled and the column is in the filter range
pub fn render_filter_button(
    app: &Spreadsheet,
    col: usize,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    // Only show if AutoFilter is enabled and column is in filter range
    if !app.filter_state.is_enabled() || !app.filter_state.contains_column(col) {
        return None;
    }

    let has_active_filter = app.column_has_filter(col);
    let accent = app.token(TokenKey::Accent);
    let text_muted = app.token(TokenKey::TextMuted);

    Some(
        div()
            .id(ElementId::NamedInteger("filter-btn".into(), col as u64))
            .absolute()
            .right(px(1.0))
            .bottom(px(1.0))
            .w(px(12.0))
            .h(px(12.0))
            .flex()
            .items_center()
            .justify_center()
            .cursor_pointer()
            .rounded_sm()
            .text_size(px(8.0))
            .when(has_active_filter, |d| d.text_color(accent).bg(accent.opacity(0.2)))
            .when(!has_active_filter, |d| d.text_color(text_muted).hover(|s| s.bg(text_muted.opacity(0.2))))
            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _: &MouseDownEvent, _, cx| {
                this.open_filter_dropdown(col, cx);
            }))
            .child("▼")
    )
}

/// Render the column header row (A, B, C, ...) with resize handles
///
/// With freeze panes active, renders:
/// 1. Frozen column headers (0 to frozen_cols-1) - always visible
/// 2. Divider line after frozen columns
/// 3. Scrollable column headers (scroll_col to scroll_col + scrollable_visible_cols)
pub fn render_column_headers(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let scroll_col = app.view_state.scroll_col;
    let visible_cols = app.visible_cols();
    let frozen_cols = app.view_state.frozen_cols;
    let header_bg = app.token(TokenKey::HeaderBg);
    let header_border = app.token(TokenKey::HeaderBorder);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let divider_color = app.token(TokenKey::PanelBorder);
    let metrics = &app.metrics;

    // Calculate scrollable region columns
    let scrollable_visible_cols = visible_cols.saturating_sub(frozen_cols);

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
        // Frozen column headers (always visible, cols 0 to frozen_cols-1)
        .when(frozen_cols > 0, |d| {
            d.children(
                (0..frozen_cols).map(|col| {
                    let col_width = metrics.col_width(app.col_width(col));
                    let is_selected = app.is_col_header_selected(col);
                    render_column_header(app, col, col_width, is_selected, cx)
                })
            )
        })
        // Divider after frozen columns
        .when(frozen_cols > 0, |d| {
            d.child(
                div()
                    .w(px(1.0))
                    .h_full()
                    .bg(divider_color)
            )
        })
        // Scrollable column headers (scroll_col to scroll_col + scrollable_visible_cols)
        .children(
            (0..scrollable_visible_cols).map(move |i| {
                let col = scroll_col + i;
                let col_width = metrics.col_width(app.col_width(col));
                let is_selected = app.is_col_header_selected(col);
                render_column_header(app, col, col_width, is_selected, cx)
            })
        )
}

/// Render sort indicator for a column if it's the sorted column
/// Returns None if not sorted, Some(element) with ▲ or ▼ if sorted
fn render_sort_indicator(app: &Spreadsheet, col: usize) -> Option<impl IntoElement> {
    let sort_state = app.display_sort_state()?;
    if sort_state.0 != col {
        return None;
    }

    let arrow = if sort_state.1 { "▲" } else { "▼" };
    let accent = app.token(TokenKey::Accent);

    Some(
        div()
            .text_size(px(8.0))
            .text_color(accent)
            .ml(px(2.0))
            .child(arrow)
    )
}

/// Width reserved for header chrome (filter button + padding)
/// This ensures text doesn't overlap with icons on narrow columns
const HEADER_CHROME_WIDTH: f32 = 16.0;

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

    // Reserve right padding when filter button is shown to prevent text/icon overlap
    let has_filter_button = app.filter_state.is_enabled() && app.filter_state.contains_column(col);

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
        // Column letter with optional sort indicator
        // Add right padding when filter button is shown to prevent overlap
        .child(
            div()
                .flex()
                .items_center()
                .overflow_hidden()
                .when(has_filter_button, |d| d.pr(px(HEADER_CHROME_WIDTH)))
                .child(Spreadsheet::col_letter(col))
                .when_some(render_sort_indicator(app, col), |d, indicator| d.child(indicator))
        )
        // Filter dropdown button (when AutoFilter is enabled)
        .when_some(render_filter_button(app, col, cx), |d, btn| d.child(btn))
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
