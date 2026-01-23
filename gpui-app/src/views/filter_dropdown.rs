//! Filter dropdown popup for AutoFilter

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the filter dropdown popup
/// Returns None if no dropdown is open
pub fn render_filter_dropdown(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    let col = app.filter_dropdown_col?;

    // Get unique values for this column
    let unique_values = app.filter_state.get_unique_values(col)?;
    let search_text = app.filter_search_text.to_lowercase();
    let checked_items = &app.filter_checked_items;

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let accent = app.token(TokenKey::Accent);
    let app_bg = app.token(TokenKey::AppBg);

    // Get header text for the column
    let header_text = if let Some(header_row) = app.filter_state.header_row() {
        app.sheet().get_display(header_row, col)
    } else {
        Spreadsheet::col_letter(col)
    };

    // Calculate position based on column position
    let col_x = app.metrics.header_w + app.col_x_offset(col);
    let header_h = app.metrics.header_h;

    // Filter values based on search text
    let filtered_values: Vec<(usize, &visigrid_engine::filter::UniqueValueEntry)> = unique_values
        .iter()
        .enumerate()
        .filter(|(_, entry)| {
            search_text.is_empty() || entry.display.to_lowercase().contains(&search_text)
        })
        .collect();

    // Count visible items and checked count
    let total_count = unique_values.len();
    let visible_count = filtered_values.len();
    let checked_count = checked_items.len();

    Some(
        div()
            .id("filter-dropdown-backdrop")
            .absolute()
            .inset_0()
            // Click outside to close (and stop propagation to grid)
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                this.close_filter_dropdown(cx);
                cx.stop_propagation();
            }))
            .child(
                div()
                    .id("filter-dropdown")
                    .absolute()
                    .left(px(col_x))
                    .top(px(header_h + 2.0))
                    .w(px(220.0))
                    .max_h(px(350.0))
                    .bg(panel_bg)
                    .border_1()
                    .border_color(panel_border)
                    .rounded_md()
                    .shadow_lg()
                    .flex()
                    .flex_col()
                    .overflow_hidden()
                    // Stop propagation for clicks inside panel
                    .on_mouse_down(MouseButton::Left, |_, _, cx| {
                        cx.stop_propagation();
                    })
                    // Header
                    .child(
                        div()
                            .px_3()
                            .py_2()
                            .border_b_1()
                            .border_color(panel_border)
                            .text_sm()
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_primary)
                            .child(format!("Filter: {}", header_text))
                    )
                    // Search box
                    .child(
                        div()
                            .px_3()
                            .py_2()
                            .border_b_1()
                            .border_color(panel_border)
                            .child(
                                div()
                                    .px_2()
                                    .py_1()
                                    .bg(app_bg)
                                    .border_1()
                                    .border_color(panel_border)
                                    .rounded_sm()
                                    .text_sm()
                                    .text_color(if app.filter_search_text.is_empty() { text_muted } else { text_primary })
                                    .child(
                                        if app.filter_search_text.is_empty() {
                                            "Search...".to_string()
                                        } else {
                                            app.filter_search_text.clone()
                                        }
                                    )
                            )
                    )
                    // Select All / Clear buttons
                    .child(
                        div()
                            .flex()
                            .px_3()
                            .py_1()
                            .gap_2()
                            .border_b_1()
                            .border_color(panel_border)
                            .child(
                                div()
                                    .id("filter-select-all")
                                    .px_2()
                                    .py_1()
                                    .text_xs()
                                    .text_color(accent)
                                    .cursor_pointer()
                                    .hover(|s| s.underline())
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.filter_select_all(cx);
                                        cx.stop_propagation();
                                    }))
                                    .child("Select All")
                            )
                            .child(
                                div()
                                    .id("filter-clear-all")
                                    .px_2()
                                    .py_1()
                                    .text_xs()
                                    .text_color(accent)
                                    .cursor_pointer()
                                    .hover(|s| s.underline())
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.filter_clear_all(cx);
                                        cx.stop_propagation();
                                    }))
                                    .child("Clear")
                            )
                    )
                    // Value count info
                    .child(
                        div()
                            .px_3()
                            .py_1()
                            .text_xs()
                            .text_color(text_muted)
                            .child(format!("{} of {} selected ({} shown)", checked_count, total_count, visible_count))
                    )
                    // Values list (or "No matches" message)
                    .child(
                        div()
                            .flex_1()
                            .overflow_hidden()
                            .max_h(px(180.0))
                            .flex()
                            .flex_col()
                            .when(filtered_values.is_empty(), |d| {
                                d.child(
                                    div()
                                        .px_3()
                                        .py_4()
                                        .text_sm()
                                        .text_color(text_muted)
                                        .italic()
                                        .child("No matches")
                                )
                            })
                            .when(!filtered_values.is_empty(), |d| {
                                d.children(
                                    filtered_values.into_iter().map(|(idx, entry)| {
                                        let is_checked = checked_items.contains(&idx);
                                        render_filter_item(idx, entry, is_checked, text_primary, text_muted, selection_bg, accent, cx)
                                    })
                                )
                            })
                    )
                    // Action buttons
                    .child(
                        div()
                            .flex()
                            .justify_end()
                            .gap_2()
                            .px_3()
                            .py_2()
                            .border_t_1()
                            .border_color(panel_border)
                            .child(
                                div()
                                    .id("filter-cancel")
                                    .px_3()
                                    .py_1()
                                    .text_sm()
                                    .text_color(text_muted)
                                    .cursor_pointer()
                                    .rounded_sm()
                                    .hover(|s| s.bg(selection_bg))
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.close_filter_dropdown(cx);
                                        cx.stop_propagation();
                                    }))
                                    .child("Cancel")
                            )
                            .child(
                                div()
                                    .id("filter-apply")
                                    .px_3()
                                    .py_1()
                                    .text_sm()
                                    .text_color(app_bg)
                                    .bg(accent)
                                    .cursor_pointer()
                                    .rounded_sm()
                                    .hover(|s| s.opacity(0.9))
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.apply_filter_dropdown(cx);
                                        cx.stop_propagation();
                                    }))
                                    .child("Apply")
                            )
                    )
            )
    )
}

/// Render a single filter item with checkbox
fn render_filter_item(
    idx: usize,
    entry: &visigrid_engine::filter::UniqueValueEntry,
    is_checked: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let display = entry.display.clone();
    let count = entry.count;

    div()
        .id(ElementId::NamedInteger("filter-item".into(), idx as u64))
        .flex()
        .items_center()
        .gap_2()
        .px_3()
        .py_1()
        .cursor_pointer()
        .hover(|s| s.bg(selection_bg.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.toggle_filter_item(idx, cx);
            cx.stop_propagation();
        }))
        // Checkbox (accent color when checked for better visibility)
        .child(
            div()
                .w(px(14.0))
                .h(px(14.0))
                .border_1()
                .rounded_sm()
                .flex()
                .items_center()
                .justify_center()
                .text_size(px(10.0))
                .when(is_checked, |d| {
                    d.bg(accent)
                        .border_color(accent)
                        .text_color(text_primary)
                        .child("✓")
                })
                .when(!is_checked, |d| d.border_color(text_muted))
        )
        // Value text
        .child(
            div()
                .flex_1()
                .text_sm()
                .text_color(text_primary)
                .overflow_hidden()
                .text_ellipsis()
                .child(display)
        )
        // Count
        .child(
            div()
                .text_xs()
                .text_color(text_muted)
                .child(format!("({})", count))
        )
}
