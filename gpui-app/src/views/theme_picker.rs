use std::time::Duration;
use gpui::*;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the theme picker overlay
pub fn render_theme_picker(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let filtered = app.filter_themes();
    let selected_idx = app.theme_picker_selected;
    let query = app.theme_picker_query.clone();
    let has_query = !query.is_empty();

    // Get current theme name
    let current_theme_id = app.theme.meta.id;

    // Theme colors (use the preview theme if active for live preview effect)
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
            this.hide_theme_picker(cx);
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
                                .child("Preferences: Theme")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(11.0))
                                .child(format!("Current: {}", app.theme.meta.name))
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
                                            "Search themes...".to_string()
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
                // Theme list
                .child({
                    let list = div()
                        .flex_1()
                        .overflow_hidden()
                        .py_1()
                        .children(
                            filtered.iter().enumerate().take(12).map(|(idx, theme)| {
                                let is_selected = idx == selected_idx;
                                let is_current = theme.meta.id == current_theme_id;
                                render_theme_item(theme, is_selected, is_current, idx, text_primary, text_muted, selection_bg, toolbar_hover, accent, cx)
                            })
                        );
                    if filtered.is_empty() {
                        list.child(
                            div()
                                .px_4()
                                .py_6()
                                .text_color(text_muted)
                                .text_size(px(14.0))
                                .child("No matching themes")
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
                                .child("Apply")
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

fn render_theme_item(
    theme: &crate::theme::Theme,
    is_selected: bool,
    is_current: bool,
    idx: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    hover_bg: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let theme_id = theme.meta.id;
    let theme_name = theme.meta.name;
    let theme_appearance = format!("{:?}", theme.meta.appearance);

    let bg_color = if is_selected { selection_bg } else { hsla(0.0, 0.0, 0.0, 0.0) };

    // Get some theme colors for preview
    let preview_bg = theme.get(TokenKey::AppBg);
    let preview_accent = theme.get(TokenKey::Accent);
    let preview_text = theme.get(TokenKey::TextPrimary);

    let mut item = div()
        .id(ElementId::NamedInteger("theme-item".into(), idx as u64))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py(px(8.0))
        .cursor_pointer()
        .bg(bg_color)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.apply_theme_at_index(idx, cx);
        }))
        .child(
            div()
                .flex()
                .items_center()
                .gap_3()
                // Color preview swatches
                .child(
                    div()
                        .flex()
                        .gap_1()
                        .child(
                            div()
                                .w(px(16.0))
                                .h(px(16.0))
                                .rounded_sm()
                                .bg(preview_bg)
                                .border_1()
                                .border_color(preview_text.opacity(0.3))
                        )
                        .child(
                            div()
                                .w(px(16.0))
                                .h(px(16.0))
                                .rounded_sm()
                                .bg(preview_accent)
                        )
                        .child(
                            div()
                                .w(px(16.0))
                                .h(px(16.0))
                                .rounded_sm()
                                .bg(preview_text)
                        )
                )
                // Theme name and appearance
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_size(px(13.0))
                                .child(theme_name)
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(11.0))
                                .child(theme_appearance)
                        )
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
                .child("Active")
        );
    }

    item
}
