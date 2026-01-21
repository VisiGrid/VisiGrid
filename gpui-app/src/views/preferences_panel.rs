//! Preferences panel (Cmd+,)
//!
//! A thin preferences UI that proves the settings architecture.
//! Intentionally minimal: Appearance, Editing, Tips, Power.

use gpui::{*, BorrowAppContext};
use crate::app::Spreadsheet;
use crate::settings::{open_settings_file, Setting, EnterBehavior, SettingsStore};
use crate::theme::TokenKey;

/// Render the preferences panel overlay
pub fn render_preferences_panel(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);
    let editor_border = app.token(TokenKey::EditorBorder);

    // Current settings values (from global store)
    let user_settings = SettingsStore::global(cx).user_settings();
    let show_gridlines = match &user_settings.appearance.show_gridlines {
        Setting::Value(v) => *v,
        Setting::Inherit => true, // Default
    };

    let enter_behavior = match &user_settings.editing.enter_behavior {
        Setting::Value(v) => *v,
        Setting::Inherit => EnterBehavior::MoveDown,
    };

    let keyboard_hints = match &user_settings.navigation.keyboard_hints {
        Setting::Value(v) => *v,
        Setting::Inherit => false, // Default
    };

    let vim_mode = match &user_settings.navigation.vim_mode {
        Setting::Value(v) => *v,
        Setting::Inherit => false, // Default
    };

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(80.0))
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        // Click outside to close
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_preferences(cx);
        }))
        .child(
            div()
                .w(px(380.0))
                .bg(panel_bg)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Stop click propagation on the panel itself
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
                        .py_3()
                        .border_b_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_size(px(14.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .child("Preferences")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_size(px(11.0))
                                .child("Esc to close")
                        )
                )
                // Content
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .p_4()
                        .gap_5()
                        // =========================================================
                        // Appearance section
                        // =========================================================
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(section_header("APPEARANCE", text_primary))
                                // Theme row
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("Theme", text_muted))
                                        .child(
                                            div()
                                                .id("pref-theme-btn")
                                                .px_3()
                                                .py(px(4.0))
                                                .bg(accent.opacity(0.15))
                                                .rounded_md()
                                                .cursor_pointer()
                                                .text_size(px(11.0))
                                                .text_color(text_primary)
                                                .hover(|s| s.bg(accent.opacity(0.25)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    this.hide_preferences(cx);
                                                    this.show_theme_picker(cx);
                                                }))
                                                .child("Change theme...")
                                        )
                                )
                                // Show gridlines row
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("Show gridlines", text_muted))
                                        .child(
                                            div()
                                                .size(px(16.0))
                                                .rounded_sm()
                                                .border_1()
                                                .border_color(if show_gridlines { accent } else { editor_border })
                                                .bg(if show_gridlines { accent } else { editor_bg })
                                                .cursor_pointer()
                                                .flex()
                                                .items_center()
                                                .justify_center()
                                                .child(if show_gridlines {
                                                    div()
                                                        .text_size(px(10.0))
                                                        .text_color(gpui::white())
                                                        .child("✓")
                                                        .into_any_element()
                                                } else {
                                                    div().into_any_element()
                                                })
                                                .id("pref-gridlines-cb")
                                                .on_click(cx.listener(move |_this, _, _, cx| {
                                                    let new_value = !show_gridlines;
                                                    cx.update_global::<SettingsStore, _>(|store, _| {
                                                        store.user_settings_mut().appearance.show_gridlines = Setting::Value(new_value);
                                                        store.save();
                                                    });
                                                    cx.notify();
                                                }))
                                        )
                                )
                        )
                        // =========================================================
                        // Editing section
                        // =========================================================
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(section_header("EDITING", text_primary))
                                // After Enter row
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("After Enter", text_muted))
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_1()
                                                .child(enter_option("Down", EnterBehavior::MoveDown, enter_behavior, accent, text_primary, text_muted, cx))
                                                .child(enter_option("Right", EnterBehavior::MoveRight, enter_behavior, accent, text_primary, text_muted, cx))
                                                .child(enter_option("Stay", EnterBehavior::Stay, enter_behavior, accent, text_primary, text_muted, cx))
                                        )
                                )
                        )
                        // =========================================================
                        // Navigation section
                        // =========================================================
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(section_header("NAVIGATION", text_primary))
                                // Keyboard hints row
                                .child(
                                    div()
                                        .flex()
                                        .flex_col()
                                        .gap(px(2.0))
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .justify_between()
                                                .child(row_label("Keyboard hints", text_muted))
                                                .child(
                                                    div()
                                                        .size(px(16.0))
                                                        .rounded_sm()
                                                        .border_1()
                                                        .border_color(if keyboard_hints { accent } else { editor_border })
                                                        .bg(if keyboard_hints { accent } else { editor_bg })
                                                        .cursor_pointer()
                                                        .flex()
                                                        .items_center()
                                                        .justify_center()
                                                        .child(if keyboard_hints {
                                                            div()
                                                                .text_size(px(10.0))
                                                                .text_color(gpui::white())
                                                                .child("✓")
                                                                .into_any_element()
                                                        } else {
                                                            div().into_any_element()
                                                        })
                                                        .id("pref-keyboard-hints-cb")
                                                        .on_click(cx.listener(move |_this, _, _, cx| {
                                                            let new_value = !keyboard_hints;
                                                            cx.update_global::<SettingsStore, _>(|store, _| {
                                                                store.user_settings_mut().navigation.keyboard_hints = Setting::Value(new_value);
                                                                store.save();
                                                            });
                                                            cx.notify();
                                                        }))
                                                )
                                        )
                                        .child(
                                            div()
                                                .text_size(px(10.0))
                                                .text_color(text_muted.opacity(0.7))
                                                .child("Press g then type letters to jump to any cell")
                                        )
                                )
                                // Vim mode row
                                .child(
                                    div()
                                        .flex()
                                        .flex_col()
                                        .gap(px(2.0))
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .justify_between()
                                                .child(row_label("Vim mode", text_muted))
                                                .child(
                                                    div()
                                                        .size(px(16.0))
                                                        .rounded_sm()
                                                        .border_1()
                                                        .border_color(if vim_mode { accent } else { editor_border })
                                                        .bg(if vim_mode { accent } else { editor_bg })
                                                        .cursor_pointer()
                                                        .flex()
                                                        .items_center()
                                                        .justify_center()
                                                        .child(if vim_mode {
                                                            div()
                                                                .text_size(px(10.0))
                                                                .text_color(gpui::white())
                                                                .child("✓")
                                                                .into_any_element()
                                                        } else {
                                                            div().into_any_element()
                                                        })
                                                        .id("pref-vim-mode-cb")
                                                        .on_click(cx.listener(move |_this, _, _, cx| {
                                                            let new_value = !vim_mode;
                                                            cx.update_global::<SettingsStore, _>(|store, _| {
                                                                store.user_settings_mut().navigation.vim_mode = Setting::Value(new_value);
                                                                store.save();
                                                            });
                                                            cx.notify();
                                                        }))
                                                )
                                        )
                                        .child(
                                            div()
                                                .text_size(px(10.0))
                                                .text_color(text_muted.opacity(0.7))
                                                .child("Navigate with h/j/k/l, press i to edit")
                                        )
                                )
                        )
                        // =========================================================
                        // Tips section
                        // =========================================================
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(section_header("TIPS", text_primary))
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("Dismissed tips", text_muted))
                                        .child(
                                            div()
                                                .id("pref-reset-tips-btn")
                                                .px_3()
                                                .py(px(4.0))
                                                .bg(accent.opacity(0.15))
                                                .rounded_md()
                                                .cursor_pointer()
                                                .text_size(px(11.0))
                                                .text_color(text_primary)
                                                .hover(|s| s.bg(accent.opacity(0.25)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    cx.update_global::<SettingsStore, _>(|store, _| {
                                                        store.user_settings_mut().reset_all_tips();
                                                        store.save();
                                                    });
                                                    this.status_message = Some("All tips have been reset".to_string());
                                                    cx.notify();
                                                }))
                                                .child("Reset all tips")
                                        )
                                )
                        )
                        // =========================================================
                        // Advanced section
                        // =========================================================
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_2()
                                .child(section_header("ADVANCED", text_primary))
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("Settings file", text_muted))
                                        .child(
                                            div()
                                                .id("pref-open-json-btn")
                                                .px_3()
                                                .py(px(4.0))
                                                .bg(accent.opacity(0.15))
                                                .rounded_md()
                                                .cursor_pointer()
                                                .text_size(px(11.0))
                                                .text_color(text_primary)
                                                .hover(|s| s.bg(accent.opacity(0.25)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    if let Err(e) = open_settings_file() {
                                                        this.status_message = Some(format!("Failed to open settings: {}", e));
                                                    } else {
                                                        this.status_message = Some("Opened settings.json in system editor".to_string());
                                                    }
                                                    this.hide_preferences(cx);
                                                }))
                                                .child("Open settings.json")
                                        )
                                )
                        )
                )
        )
}

/// Section header (e.g., "APPEARANCE")
fn section_header(title: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .text_size(px(10.0))
        .text_color(text_color)
        .font_weight(FontWeight::SEMIBOLD)
        .child(title)
}

/// Row label
fn row_label(label: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .text_size(px(12.0))
        .text_color(text_color)
        .child(label)
}

/// Enter behavior option button
fn enter_option(
    label: &'static str,
    behavior: EnterBehavior,
    current: EnterBehavior,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = current == behavior;
    let bg = if is_selected { accent.opacity(0.2) } else { gpui::transparent_black() };
    let text = if is_selected { text_primary } else { text_muted };

    div()
        .id(SharedString::from(format!("enter-{}", label.to_lowercase())))
        .px_2()
        .py(px(3.0))
        .rounded_sm()
        .bg(bg)
        .cursor_pointer()
        .text_size(px(11.0))
        .text_color(text)
        .hover(|s| s.bg(accent.opacity(0.1)))
        .on_click(cx.listener(move |_this, _, _, cx| {
            cx.update_global::<SettingsStore, _>(|store, _| {
                store.user_settings_mut().editing.enter_behavior = Setting::Value(behavior);
                store.save();
            });
            cx.notify();
        }))
        .child(label)
}
