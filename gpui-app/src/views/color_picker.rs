//! Color picker modal for fill color selection.
//!
//! Reusable design: the picker is parameterized so it can later serve
//! text color and border color by swapping the `on_apply` closure.

use std::time::Duration;
use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::color_palette::{self, STANDARD_COLORS, rgba_to_hsla, parse_hex_color, to_hex};
use crate::theme::TokenKey;
use crate::ui::modal_overlay;

/// Swatch size in pixels
const SWATCH_SIZE: f32 = 22.0;
/// Gap between swatches
const SWATCH_GAP: f32 = 2.0;

/// Render the fill color picker as a modal overlay.
pub fn render_color_picker(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);

    // Get current cell's background color
    let (row, col) = app.view_state.selected;
    let current_bg = app.sheet(cx).get_background_color(row, col);

    // Build theme grid (6 rows × 10 cols)
    let grid = color_palette::theme_grid();

    // Hex input state
    let hex_input = app.ui.color_picker.hex_input.clone();
    let has_hex = !hex_input.is_empty();
    let parsed_preview = parse_hex_color(&hex_input);

    // Pre-build all swatch elements (can't pass cx into iterators)
    let mut grid_rows: Vec<Div> = Vec::with_capacity(6);
    for row_idx in 0..6 {
        let mut row_div = div().flex().gap(px(SWATCH_GAP));
        for col_idx in 0..10 {
            let color = grid[row_idx * 10 + col_idx];
            row_div = row_div.child(render_swatch(color, current_bg, accent, panel_border, cx));
        }
        grid_rows.push(row_div);
    }

    // Pre-build standard color swatches
    let mut standard_swatches: Vec<Stateful<Div>> = Vec::with_capacity(10);
    for &color in &STANDARD_COLORS {
        standard_swatches.push(render_swatch(color, current_bg, accent, panel_border, cx));
    }

    // Pre-build No Fill swatch (only when target allows clearing)
    let allow_none = app.ui.color_picker.target.allow_none();
    let no_fill = if allow_none {
        Some(render_no_fill_swatch(panel_border, text_muted, cx))
    } else {
        None
    };

    // Pre-build recent color swatches
    let has_recents = !app.ui.color_picker.recent_colors.is_empty();
    let mut recent_swatches: Vec<Stateful<Div>> = Vec::new();
    if has_recents {
        for &color in &app.ui.color_picker.recent_colors {
            recent_swatches.push(render_swatch(color, current_bg, accent, panel_border, cx));
        }
    }

    let dialog_content = div()
        .track_focus(&app.ui.color_picker.focus)
        .w(px(280.0))
        .bg(panel_bg)
        .rounded_md()
        .shadow_lg()
        .overflow_hidden()
        .flex()
        .flex_col()
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
                .child(
                    div()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .font_weight(FontWeight::MEDIUM)
                        .child(app.ui.color_picker.target.title())
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(render_color_chip(current_bg, panel_border))
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(10.0))
                                .child(SharedString::from(
                                    current_bg.map(to_hex).unwrap_or_else(|| "None".to_string())
                                ))
                        )
                )
        )
        // Theme grid (6 rows × 10 cols)
        .child(
            div()
                .px_3()
                .pt_2()
                .pb_1()
                .flex()
                .flex_col()
                .gap(px(SWATCH_GAP))
                .children(grid_rows)
        )
        // Separator
        .child(
            div()
                .mx_3()
                .my_1()
                .h(px(1.0))
                .bg(panel_border)
        )
        // Standard colors row
        .child(
            div()
                .px_3()
                .pb_1()
                .child(
                    div()
                        .text_color(text_muted)
                        .text_size(px(10.0))
                        .mb_1()
                        .child("Standard Colors")
                )
                .child(
                    div()
                        .flex()
                        .gap(px(SWATCH_GAP))
                        .children(standard_swatches)
                )
        )
        // No Fill swatch (conditional on target.allow_none())
        .when(allow_none, |el| {
            el.child(
                div()
                    .px_3()
                    .py_1()
                    .children(no_fill)
            )
        })
        // Recent colors row (if non-empty)
        .when(has_recents, |el| {
            el.child(
                div()
                    .px_3()
                    .pb_1()
                    .child(
                        div()
                            .text_color(text_muted)
                            .text_size(px(10.0))
                            .mb_1()
                            .child("Recent Colors")
                    )
                    .child(
                        div()
                            .flex()
                            .gap(px(SWATCH_GAP))
                            .children(recent_swatches)
                    )
            )
        })
        // Hex input
        .child(
            div()
                .px_3()
                .py_2()
                .border_t_1()
                .border_color(panel_border)
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_color(text_muted)
                        .text_size(px(11.0))
                        .child("Hex")
                )
                .child(
                    div()
                        .flex_1()
                        .flex()
                        .items_center()
                        .px_2()
                        .py(px(3.0))
                        .bg(panel_border.opacity(0.2))
                        .border_1()
                        .border_color(panel_border)
                        .rounded_sm()
                        .child(
                            div()
                                .text_color(if has_hex { text_primary } else { text_disabled })
                                .text_size(px(12.0))
                                .child(if has_hex {
                                    hex_input.clone()
                                } else {
                                    "#RRGGBB".to_string()
                                })
                        )
                        // Blinking cursor
                        .child(
                            div()
                                .w(px(1.0))
                                .h(px(13.0))
                                .bg(text_primary)
                                .ml(px(1.0))
                                .with_animation(
                                    "hex-cursor-blink",
                                    Animation::new(Duration::from_millis(530))
                                        .repeat()
                                        .with_easing(pulsating_between(0.0, 1.0)),
                                    |div, delta| {
                                        let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                        div.opacity(opacity)
                                    },
                                )
                        )
                )
                // Preview chip (shows parsed color if valid)
                .when(parsed_preview.is_some(), |el| {
                    el.child(render_color_chip(parsed_preview, panel_border))
                })
        )
        // Footer hints
        .child(
            div()
                .px_3()
                .py(px(5.0))
                .border_t_1()
                .border_color(panel_border)
                .text_size(px(10.0))
                .text_color(text_disabled)
                .child("Click \u{00b7} Shift+Click to keep open \u{00b7} Esc")
        );

    modal_overlay(
        "color-picker-dialog",
        |this, cx| this.hide_color_picker(cx),
        dialog_content,
        cx,
    )
}

/// Render a single color swatch with click/shift-click behavior.
fn render_swatch(
    color: [u8; 4],
    current_bg: Option<[u8; 4]>,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let is_current = current_bg == Some(color);
    let swatch_bg = rgba_to_hsla(color);

    let mut swatch = div()
        .id(ElementId::Name(SharedString::from(format!("swatch-{}-{}-{}", color[0], color[1], color[2]))))
        .size(px(SWATCH_SIZE))
        .rounded_sm()
        .cursor_pointer()
        .bg(swatch_bg);

    if is_current {
        swatch = swatch
            .border_2()
            .border_color(accent);
    } else {
        swatch = swatch
            .border_1()
            .border_color(panel_border)
            .hover(|s| s.border_color(accent));
    }

    swatch
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, window, cx| {
            cx.stop_propagation();
            this.apply_color_from_picker(Some(color), window, cx);
            if !event.modifiers.shift {
                this.hide_color_picker(cx);
            }
        }))
}

/// Render the "No Fill" swatch (clear background).
fn render_no_fill_swatch(
    panel_border: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .id("swatch-no-fill")
        .flex()
        .items_center()
        .gap_2()
        .cursor_pointer()
        .px_1()
        .py(px(2.0))
        .rounded_sm()
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, window, cx| {
            cx.stop_propagation();
            this.apply_color_from_picker(None, window, cx);
            if !event.modifiers.shift {
                this.hide_color_picker(cx);
            }
        }))
        .child(
            // White swatch with diagonal red line indicator
            div()
                .size(px(SWATCH_SIZE))
                .rounded_sm()
                .border_1()
                .border_color(panel_border)
                .bg(hsla(0.0, 0.0, 1.0, 1.0))
                .overflow_hidden()
                .child(
                    // Diagonal red line
                    div()
                        .absolute()
                        .w(px(30.0))
                        .h(px(2.0))
                        .bg(hsla(0.0, 0.85, 0.55, 1.0))
                        .top(px(10.0))
                        .left(px(-4.0))
                )
        )
        .child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("No Fill")
        )
}

/// Render a small color chip (preview square).
fn render_color_chip(color: Option<[u8; 4]>, border: Hsla) -> impl IntoElement {
    let bg = color
        .map(rgba_to_hsla)
        .unwrap_or(hsla(0.0, 0.0, 1.0, 1.0));

    div()
        .size(px(20.0))
        .rounded_sm()
        .border_1()
        .border_color(border)
        .bg(bg)
        .flex_shrink_0()
}
