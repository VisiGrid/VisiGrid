use std::time::Duration;
use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the font picker overlay
pub fn render_font_picker(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let filtered = app.filter_fonts();
    let selected_idx = app.font_picker_selected;
    let query = app.font_picker_query.clone();
    let has_query = !query.is_empty();

    // Get current font for selected cell
    let (row, col) = app.view_state.selected;
    let current_font = app.sheet().get_font_family(row, col);
    let current_font_display = current_font.clone().unwrap_or_else(|| "(Default)".to_string());

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let toolbar_hover = app.token(TokenKey::ToolbarButtonHoverBg);
    let accent = app.token(TokenKey::Accent);

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
            this.hide_font_picker(cx);
        }))
        .child(
            div()
                .w(px(400.0))
                .max_h(px(420.0))
                .bg(panel_bg)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Stop click propagation on the picker itself
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
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
                                .child("Select Font")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(11.0))
                                .child(format!("Current: {}", current_font_display))
                        )
                )
                // Search input
                .child(
                    div()
                        .flex()
                        .items_center()
                        .px_3()
                        .py(px(10.0))
                        .border_b_1()
                        .border_color(panel_border)
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
                                            "Search fonts...".to_string()
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
                // Font list
                .child({
                    let list = div()
                        .flex_1()
                        .overflow_hidden()
                        .py_1()
                        .children(
                            filtered.iter().enumerate().take(12).map(|(idx, font_name)| {
                                let is_selected = idx == selected_idx;
                                let is_current = current_font.as_ref() == Some(font_name);
                                render_font_item(font_name, is_selected, is_current, idx, text_primary, selection_bg, toolbar_hover, accent, cx)
                            })
                        );
                    if filtered.is_empty() {
                        list.child(
                            div()
                                .px_4()
                                .py_6()
                                .text_color(text_muted)
                                .text_size(px(14.0))
                                .child("No matching fonts")
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
                        .child(
                            div()
                                .flex()
                                .gap_3()
                                .child("Select")
                                .child(
                                    div()
                                        .text_color(text_disabled)
                                        .child("enter")
                                )
                        )
                        .child(
                            div()
                                .flex()
                                .gap_3()
                                .child("Cancel")
                                .child(
                                    div()
                                        .text_color(text_disabled)
                                        .child("esc")
                                )
                        )
                )
        )
}

fn render_font_item(
    font_name: &str,
    is_selected: bool,
    is_current: bool,
    idx: usize,
    text_primary: Hsla,
    selection_bg: Hsla,
    hover_bg: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let font_str = font_name.to_string();
    let font_for_action = font_str.clone();

    let bg_color = if is_selected { selection_bg } else { hsla(0.0, 0.0, 0.0, 0.0) };

    let mut item = div()
        .id(ElementId::NamedInteger("font-item".into(), idx as u64))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py(px(6.0))
        .cursor_pointer()
        .bg(bg_color)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.apply_font_to_selection(&font_for_action, cx);
            this.hide_font_picker(cx);
        }))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        // Show font name in its own font family
                        .font_family(font_str.clone())
                        .child(font_str.clone())
                )
        );

    if !is_selected {
        item = item.hover(move |s| s.bg(hover_bg));
    }

    if is_current {
        item = item.child(
            div()
                .text_color(accent)
                .text_size(px(12.0))
                .child("✓")
        );
    }

    item
}
