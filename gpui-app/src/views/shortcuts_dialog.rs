use std::time::Duration;
use gpui::*;
use crate::app::Spreadsheet;
use crate::shortcuts::{sections, CUSTOM_TITLE};
use crate::theme::TokenKey;

fn keycap(key: String, text: Hsla, border: Hsla) -> Div {
    div()
        .px(px(5.0))
        .min_w(px(18.0))
        .flex()
        .justify_center()
        .border_1()
        .border_color(border)
        .rounded(px(3.0))
        .text_size(px(11.0))
        .line_height(px(16.0))
        .font_weight(FontWeight::MEDIUM)
        .text_color(text)
        .child(key)
}

/// "Ctrl+Shift+L" → ["Ctrl", "Shift", "L"]; "Ctrl++" → ["Ctrl", "+"].
fn key_parts(keys: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut rest = keys;
    while !rest.is_empty() {
        let first = rest.chars().next().map_or(0, char::len_utf8);
        match rest[first..].find('+').map(|i| i + first) {
            Some(i) => {
                parts.push(rest[..i].to_string());
                rest = &rest[i + 1..];
            }
            None => {
                parts.push(rest.to_string());
                break;
            }
        }
    }
    parts
}

/// Render the Keyboard Shortcuts window (Help > Keyboard Shortcuts).
pub fn render_shortcuts_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let query = app.shortcuts_query.clone();
    let has_query = !query.is_empty();
    let found = sections(&app.shortcuts_custom, &query);
    let count: usize = found.iter().map(|s| s.rows.len()).sum();

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let hover_bg = app.token(TokenKey::ToolbarButtonHoverBg);
    let accent = app.token(TokenKey::Accent);

    let mut list = div()
        .id("shortcuts-list")
        .flex_1()
        .min_h(px(0.0))
        .overflow_y_scroll()
        .track_scroll(&app.shortcuts_scroll)
        .pb_2();

    for section in &found {
        list = list.child(
            div()
                .px_4()
                .pt(px(12.0))
                .pb(px(4.0))
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(if section.title == CUSTOM_TITLE { accent } else { text_muted })
                .child(section.title.to_uppercase()),
        );
        for row in &section.rows {
            let mut keys = div().flex().flex_shrink_0().gap(px(3.0));
            for part in key_parts(&row.keys) {
                keys = keys.child(keycap(part, text_primary, panel_border));
            }
            list = list.child(
                div()
                    .flex()
                    .items_center()
                    .justify_between()
                    .gap_4()
                    .px_4()
                    .py(px(4.0))
                    .hover(move |s| s.bg(hover_bg))
                    .child(
                        div()
                            .flex_1()
                            .min_w(px(0.0))
                            .overflow_hidden()
                            .text_ellipsis()
                            .whitespace_nowrap()
                            .text_size(px(13.0))
                            .text_color(text_primary)
                            .child(row.name.clone()),
                    )
                    .child(keys),
            );
        }
    }
    if found.is_empty() {
        list = list.child(
            div()
                .px_4()
                .py_6()
                .text_size(px(13.0))
                .text_color(text_muted)
                .child(format!("No shortcut matches \"{query}\"")),
        );
    }

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(70.0))
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        .occlude()
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_shortcuts(cx);
        }))
        .child(
            div()
                .w(px(560.0))
                .h(px(520.0))
                .max_h(relative(0.8))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                // Header
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .px_4()
                        .py_2()
                        .border_b_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_size(px(13.0))
                                .font_weight(FontWeight::MEDIUM)
                                .child("Keyboard Shortcuts"),
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(11.0))
                                .child(if count == 1 { "1 shortcut".to_string() } else { format!("{count} shortcuts") }),
                        ),
                )
                // Search
                .child(
                    div()
                        .flex()
                        .items_center()
                        .px_4()
                        .py(px(10.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .text_color(if has_query { text_primary } else { text_disabled })
                                .text_size(px(13.0))
                                .child(if has_query { query.clone() } else { "Search shortcuts...".to_string() }),
                        )
                        .child(
                            div()
                                .w(px(1.0))
                                .h(px(14.0))
                                .bg(text_primary)
                                .ml(px(1.0))
                                .with_animation(
                                    "shortcuts-cursor-blink",
                                    Animation::new(Duration::from_millis(530))
                                        .repeat()
                                        .with_easing(pulsating_between(0.0, 1.0)),
                                    |div, delta| div.opacity(if delta > 0.5 { 0.0 } else { 1.0 }),
                                ),
                        ),
                )
                .child(list)
                // Footer
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .px_4()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(panel_border)
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(
                            div()
                                .id("shortcuts-customize")
                                .cursor_pointer()
                                .text_color(accent)
                                .hover(|s| s.underline())
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_shortcuts(cx);
                                    this.open_keybindings(cx);
                                }))
                                .child("Change shortcuts (keybindings.json)"),
                        )
                        .child(
                            div()
                                .flex()
                                .gap_3()
                                .child("Close")
                                .child(div().text_color(text_disabled).child("esc")),
                        ),
                ),
        )
}

#[cfg(test)]
mod tests {
    use super::key_parts;

    #[test]
    fn keys_split_into_caps() {
        assert_eq!(key_parts("Ctrl+Shift+L"), ["Ctrl", "Shift", "L"]);
        assert_eq!(key_parts("Ctrl++"), ["Ctrl", "+"]);
        assert_eq!(key_parts("F9"), ["F9"]);
    }
}
