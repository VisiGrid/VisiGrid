//! Command palette rendering
//!
//! This module provides the UI for the command palette overlay.
//! Search logic is handled by the search engine in `search.rs`.

use std::time::Duration;
use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::actions::{PaletteUp, PaletteDown, PaletteExecute, PalettePreview, PaletteCancel};
use crate::app::Spreadsheet;
use crate::search::SearchItem;
use crate::theme::TokenKey;

/// Render the command palette overlay
pub fn render_command_palette(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let results = app.palette_results();
    let selected_idx = app.palette_selected;
    let query = app.palette_query.clone();
    let has_query = !query.is_empty();
    let is_previewing = app.palette_previewing;
    let total_results = app.palette_total_results;
    let shown_results = results.len();
    let is_truncated = total_results > shown_results;

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let toolbar_hover = app.token(TokenKey::ToolbarButtonHoverBg);

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(100.0))
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        // Click outside to close
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_palette(cx);
        }))
        .child(
            div()
                .key_context("CommandPalette")
                .track_focus(&app.focus_handle)
                .w(px(500.0))
                .max_h(px(380.0))
                .bg(panel_bg)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Action handlers
                .on_action(cx.listener(|this, _: &PaletteUp, _, cx| {
                    this.palette_up(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteDown, _, cx| {
                    this.palette_down(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteExecute, _, cx| {
                    this.palette_execute(cx);
                }))
                .on_action(cx.listener(|this, _: &PalettePreview, _, cx| {
                    this.palette_preview(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteCancel, _, cx| {
                    this.hide_palette(cx);
                }))
                // Stop click propagation on the palette itself
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                // Search input
                .child(
                    div()
                        .flex()
                        .items_center()
                        .px_3()
                        .py(px(10.0))
                        .border_b_1()
                        .border_color(panel_border)
                        // Input area with cursor
                        .child(
                            div()
                                .flex_1()
                                .flex()
                                .items_center()
                                .child(
                                    div()
                                        .text_color(if has_query { text_primary } else { text_disabled })
                                        .text_size(px(13.0))
                                        .child(if has_query {
                                            query.clone()
                                        } else {
                                            "Execute a command...".to_string()
                                        })
                                )
                                // Blinking cursor
                                .child(
                                    div()
                                        .w(px(1.0))
                                        .h(px(14.0))
                                        .bg(text_primary)
                                        .ml(px(1.0))
                                        .with_animation(
                                            "cursor-blink",
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
                )
                // Result list
                .child({
                    let mut list = div()
                        .flex_1()
                        .overflow_hidden()
                        .py_1();

                    // Show help hints when query is empty
                    if !has_query {
                        list = list.child(
                            div()
                                .px_4()
                                .py_2()
                                .text_color(text_disabled)
                                .text_size(px(11.0))
                                .flex()
                                .flex_col()
                                .gap_1()
                                .child(
                                    div().flex().gap_4()
                                        .child(div().child(">  commands"))
                                        .child(div().child(":  go to cell"))
                                )
                                .child(
                                    div().flex().gap_4()
                                        .child(div().child("@  search cells"))
                                        .child(div().child("=  functions"))
                                )
                                .child(
                                    div().flex().gap_4()
                                        .child(div().child("#  settings"))
                                        .child(div().child("$  named ranges"))
                                )
                                .child(
                                    div()
                                        .mt_2()
                                        .text_color(text_muted.opacity(0.7))
                                        .italic()
                                        .child("Refactor spreadsheets like code.")
                                )
                        );
                    }

                    // Add search results
                    list = list.children(
                        results.iter().enumerate().map(|(idx, item)| {
                            let is_selected = idx == selected_idx;
                            render_search_item(item, is_selected, idx, text_primary, text_muted, selection_bg, toolbar_hover, cx)
                        })
                    );

                    if results.is_empty() && has_query {
                        list.child(
                            div()
                                .px_4()
                                .py_6()
                                .text_color(text_muted)
                                .text_size(px(14.0))
                                .child("No matching commands")
                        )
                    } else {
                        list
                    }
                })
                // Footer with hints
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(panel_border)
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        // Left side: preview indicator or truncation notice
                        .child(
                            div()
                                .when(is_previewing, |el| {
                                    el.child("Previewing - Esc to restore")
                                })
                                .when(!is_previewing && is_truncated, |el| {
                                    el.child(format!("Showing {} of {} matches", shown_results, total_results))
                                })
                        )
                        // Right side: keyboard hints
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_3()
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .gap_1()
                                        .child("Alt")
                                        .child(
                                            div()
                                                .text_color(text_disabled)
                                                .child("ctrl+↵")
                                        )
                                )
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .gap_1()
                                        .child("Preview")
                                        .child(
                                            div()
                                                .text_color(text_disabled)
                                                .child("shift+↵")
                                        )
                                )
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .gap_1()
                                        .child("Run")
                                        .child(
                                            div()
                                                .text_color(text_disabled)
                                                .child("↵")
                                        )
                                )
                        )
                )
        )
}

/// Render text with highlighted spans
fn render_highlighted_text(
    text: &str,
    highlights: &[(usize, usize)],
    normal_color: Hsla,
    highlight_color: Hsla,
) -> Div {
    if highlights.is_empty() {
        return div()
            .flex()
            .text_color(normal_color)
            .text_size(px(13.0))
            .child(text.to_string());
    }

    let mut container = div()
        .flex()
        .items_center()
        .text_size(px(13.0));

    let chars: Vec<char> = text.chars().collect();
    let mut pos = 0;

    for &(start, end) in highlights {
        // Clamp to valid range
        let start = start.min(chars.len());
        let end = end.min(chars.len());

        // Add text before highlight
        if pos < start {
            let segment: String = chars[pos..start].iter().collect();
            container = container.child(
                div().text_color(normal_color).child(segment)
            );
        }

        // Add highlighted text
        if start < end {
            let segment: String = chars[start..end].iter().collect();
            container = container.child(
                div()
                    .text_color(highlight_color)
                    .font_weight(FontWeight::SEMIBOLD)
                    .child(segment)
            );
        }

        pos = end;
    }

    // Add remaining text after last highlight
    if pos < chars.len() {
        let segment: String = chars[pos..].iter().collect();
        container = container.child(
            div().text_color(normal_color).child(segment)
        );
    }

    container
}

fn render_search_item(
    item: &SearchItem,
    is_selected: bool,
    idx: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let title = item.title.clone();
    let subtitle = item.subtitle.clone();
    let kind = item.kind;
    let highlights = item.highlights.clone();

    let bg_color = if is_selected { selection_bg } else { hsla(0.0, 0.0, 0.0, 0.0) };

    // Icon based on result kind (use the centralized icon from SearchKind)
    let icon = kind.icon();

    // Highlight color: brighter version of primary
    let highlight_color = hsla(0.55, 0.8, 0.65, 1.0);  // Bright cyan/teal for highlights

    let mut row = div()
        .id(ElementId::NamedInteger("palette-item".into(), idx as u64))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py(px(6.0))
        .cursor_pointer()
        .bg(bg_color)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // Select this item and execute
            this.palette_selected = idx;
            this.palette_execute(cx);
        }))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                // Kind icon
                .child(
                    div()
                        .text_color(text_muted)
                        .text_size(px(12.0))
                        .w(px(12.0))
                        .child(icon)
                )
                // Title with highlighted matches
                .child(
                    render_highlighted_text(&title, &highlights, text_primary, highlight_color)
                )
        );

    if !is_selected {
        row = row.hover(move |s| s.bg(hover_bg));
    }

    // Subtitle (shortcut or description)
    if let Some(sub) = subtitle {
        row = row.child(
            div()
                .text_color(text_muted)
                .text_size(px(12.0))
                .child(sub)
        );
    }

    row
}
