//! Validation dropdown popup for list validation
//!
//! Renders the dropdown for cells with list validation. The dropdown shows
//! filtered items, allows keyboard navigation, and type-to-filter.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::validation_dropdown::DropdownOpenState;

/// Maximum visible rows before scrolling
const MAX_VISIBLE_ROWS: usize = 10;

/// Minimum dropdown width
const MIN_WIDTH: f32 = 180.0;

/// Row height for list items
const ROW_HEIGHT: f32 = 24.0;

/// Render the validation dropdown popup
/// Returns None if no dropdown is open
pub fn render_validation_dropdown(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    let open_state = app.validation_dropdown.as_open()?;

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let app_bg = app.token(TokenKey::AppBg);

    // Get cell rect for positioning
    let (row, col) = open_state.cell;
    let cell_rect = app.cell_rect(row, col);

    // Calculate dropdown position
    // Anchor below the cell, with offset for headers
    let dropdown_x = app.metrics.header_w + cell_rect.x;
    let dropdown_y = app.metrics.header_h + cell_rect.y + cell_rect.height;

    // Calculate dropdown width (at least cell width or MIN_WIDTH)
    let dropdown_width = cell_rect.width.max(MIN_WIDTH);

    // Calculate available height below and above the cell
    let window_height: f32 = app.window_size.height.into();
    let space_below = window_height - dropdown_y - 8.0; // 8px margin
    let space_above = app.metrics.header_h + cell_rect.y - 8.0;

    // Calculate content height based on filtered items
    let filtered_count = open_state.filtered_count();
    let visible_rows = filtered_count.min(MAX_VISIBLE_ROWS);
    let content_height = (visible_rows as f32 * ROW_HEIGHT) + 56.0; // 56px for header + footer

    // Decide whether to flip above
    let (final_y, max_height) = if content_height <= space_below {
        // Fits below
        (dropdown_y, space_below)
    } else if content_height <= space_above {
        // Flip above
        let y = app.metrics.header_h + cell_rect.y - content_height.min(space_above);
        (y, space_above)
    } else {
        // Use whichever has more space
        if space_below >= space_above {
            (dropdown_y, space_below)
        } else {
            let y = app.metrics.header_h + cell_rect.y - space_above;
            (y, space_above)
        }
    };

    Some(
        div()
            .id("validation-dropdown-backdrop")
            .absolute()
            .inset_0()
            // Click outside to close (and stop propagation to grid)
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                this.close_validation_dropdown(
                    crate::validation_dropdown::DropdownCloseReason::ClickOutside,
                    cx,
                );
                cx.stop_propagation();
            }))
            .child(render_dropdown_panel(
                open_state,
                dropdown_x,
                final_y,
                dropdown_width,
                max_height,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                selection_bg,
                app_bg,
                cx,
            ))
    )
}

/// Render the actual dropdown panel
fn render_dropdown_panel(
    state: &DropdownOpenState,
    x: f32,
    y: f32,
    width: f32,
    max_height: f32,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    app_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let filtered_count = state.filtered_count();
    let filter_text = state.filter_text.clone();
    let selected_index = state.selected_index;
    let truncated = state.truncated;
    let scroll_offset = state.scroll_offset;

    // Collect visible items
    let visible_items: Vec<(usize, String)> = state
        .visible_items()
        .skip(scroll_offset)
        .take(MAX_VISIBLE_ROWS)
        .map(|(idx, s)| (idx, s.to_string()))
        .collect();

    div()
        .id("validation-dropdown")
        .absolute()
        .left(px(x))
        .top(px(y))
        .w(px(width))
        .max_h(px(max_height))
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
        // Filter text line (shows current filter or placeholder)
        .child(
            div()
                .px_2()
                .py_1()
                .border_b_1()
                .border_color(panel_border)
                .child(
                    div()
                        .px_2()
                        .py(px(4.0))
                        .bg(app_bg)
                        .border_1()
                        .border_color(panel_border)
                        .rounded_sm()
                        .text_sm()
                        .text_color(if filter_text.is_empty() { text_muted } else { text_primary })
                        .child(
                            if filter_text.is_empty() {
                                "Type to filter...".to_string()
                            } else {
                                filter_text
                            }
                        )
                )
        )
        // List items
        .child(
            div()
                .flex_1()
                .overflow_hidden()
                .max_h(px(MAX_VISIBLE_ROWS as f32 * ROW_HEIGHT))
                .flex()
                .flex_col()
                .children(
                    visible_items.into_iter().map(|(display_idx, item)| {
                        let is_selected = display_idx == selected_index;
                        let item_for_click = item.clone();
                        div()
                            .id(SharedString::from(format!("validation-item-{}", display_idx)))
                            .px_3()
                            .py(px(4.0))
                            .text_sm()
                            .text_color(text_primary)
                            .cursor_pointer()
                            .when(is_selected, |d| {
                                d.bg(selection_bg)
                            })
                            .hover(|d| d.bg(selection_bg.opacity(0.5)))
                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                // Commit this value
                                this.commit_validation_value(&item_for_click, cx);
                                cx.stop_propagation();
                            }))
                            .child(item)
                    })
                )
        )
        // Footer: item count and truncation indicator
        .child(
            div()
                .px_3()
                .py_1()
                .border_t_1()
                .border_color(panel_border)
                .flex()
                .justify_between()
                .text_xs()
                .text_color(text_muted)
                .child(format!("{} items", filtered_count))
                .when(truncated, |d| {
                    d.child(
                        div()
                            .text_color(text_muted.opacity(0.7))
                            .child("(truncated)")
                    )
                })
        )
}
