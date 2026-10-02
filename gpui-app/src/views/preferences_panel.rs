//! Preferences panel (Cmd+,)
//!
//! Grouped application settings, with persistent navigation and a bounded content pane.

use crate::app::Spreadsheet;
use crate::settings::{
    open_settings_file, AltAccelerators, EnterBehavior, ModifierStyle, Setting, SettingsStore,
};
use crate::theme::TokenKey;
use crate::ui::cell_size_input::CellSizeField;
use gpui::{prelude::FluentBuilder, BorrowAppContext, *};

#[derive(Clone, Copy, Default, PartialEq, Eq)]
pub enum PreferencesPage {
    #[default]
    Appearance,
    Editing,
    Advanced,
}

/// Render the preferences panel overlay
pub fn render_preferences_panel(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
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

    let show_comment_previews = crate::comments::previews_enabled(cx);

    let enter_behavior = match &user_settings.editing.enter_behavior {
        Setting::Value(v) => *v,
        Setting::Inherit => EnterBehavior::MoveDown,
    };

    let paste_values_by_default = match &user_settings.editing.paste_values_by_default {
        Setting::Value(v) => *v,
        Setting::Inherit => false, // Default: Ctrl+V pastes everything
    };

    let keyboard_hints = match &user_settings.navigation.keyboard_hints {
        Setting::Value(v) => *v,
        Setting::Inherit => false, // Default
    };

    let vim_mode = match &user_settings.navigation.vim_mode {
        Setting::Value(v) => *v,
        Setting::Inherit => false, // Default
    };

    let modifier_style = match &user_settings.navigation.modifier_style {
        Setting::Value(v) => *v,
        Setting::Inherit => ModifierStyle::Platform,
    };

    let alt_accelerators = match &user_settings.navigation.alt_accelerators {
        Setting::Value(v) => *v,
        Setting::Inherit => AltAccelerators::Disabled,
    };

    let page = app.preferences_page;
    let content = div().flex().flex_col().gap(px(24.));
    let content = match page {
        PreferencesPage::Appearance => content.child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(section_header("Appearance", text_primary))
                // Theme row
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(row_label("Theme", text_primary))
                        .child(
                            div()
                                .id("pref-theme-btn")
                                .keyboard(app, "pref-theme-btn")
                                .px_3()
                                .py(px(6.0))
                                .bg(accent.opacity(0.15))
                                .rounded_md()
                                .cursor_pointer()
                                .text_size(px(11.0))
                                .text_color(text_primary)
                                .hover(|s| s.bg(accent.opacity(0.25)))
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.close_preferences_keyboard(window, cx);
                                    this.show_theme_picker(cx);
                                }))
                                .child("Choose theme…"),
                        ),
                )
                // Show gridlines row
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(row_label("Show gridlines", text_primary))
                        .child(
                            preference_toggle(
                                app,
                                "pref-gridlines-cb",
                                show_gridlines,
                                accent,
                                text_muted.opacity(0.25),
                            )
                            .on_click(cx.listener(
                                move |_this, _, _, cx| {
                                    let new_value = !show_gridlines;
                                    cx.update_global::<SettingsStore, _>(|store, _| {
                                        store.user_settings_mut().appearance.show_gridlines =
                                            Setting::Value(new_value);
                                        store.save();
                                    });
                                    cx.notify();
                                },
                            )),
                        ),
                )
                // Comment reading previews
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(row_label("Comment hover previews", text_primary))
                        .child(
                            preference_toggle(
                                app,
                                "pref-comment-previews-cb",
                                show_comment_previews,
                                accent,
                                text_muted.opacity(0.25),
                            )
                            .on_click(
                                cx.listener(|this, _, _, cx| this.toggle_comment_previews(cx)),
                            ),
                        ),
                )
                .child(
                    div()
                        .mt_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .pt_4()
                        .child(section_header("Cell defaults", text_primary)),
                )
                .child(cell_size_row(
                    app,
                    CellSizeField::ColumnWidth,
                    "Default column width",
                    cx,
                ))
                .child(cell_size_row(
                    app,
                    CellSizeField::RowHeight,
                    "Default row height",
                    cx,
                ))
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child("Applies to rows and columns without an explicit size."),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child("Pixels at 100% zoom"),
                        )
                        .child(
                            div()
                                .id("pref-reset-cell-sizes")
                                .keyboard(app, "pref-reset-cell-sizes")
                                .cursor_pointer()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .on_click(cx.listener(|this, _, _, cx| {
                                    crate::settings::update_user_settings(cx, |settings| {
                                        settings.appearance.default_column_width = Setting::Inherit;
                                        settings.appearance.default_row_height = Setting::Inherit;
                                    });
                                    this.ui.cell_size_input = Default::default();
                                    cx.notify();
                                }))
                                .child("Reset sizes"),
                        ),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(row_label("Default font", text_primary))
                        .child(
                            div()
                                .id("pref-font-family")
                                .keyboard(app, "pref-font-family")
                                .px_2()
                                .py(px(6.0))
                                .rounded_sm()
                                .border_1()
                                .border_color(editor_border)
                                .bg(editor_bg)
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .cursor_pointer()
                                .on_click(cx.listener(|this, _, window, cx| {
                                    if !this.apply_cell_size_input(cx) {
                                        return;
                                    }
                                    this.show_font_picker(window, cx);
                                    this.font_picker_for_default = true;
                                }))
                                .child(app.cell_font.family.clone()),
                        ),
                )
                .when(app.font_catalog.is_missing(&app.cell_font.family), |d| {
                    d.child(
                        div()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .child(format!(
                                "Font unavailable; using {}.",
                                crate::settings::DEFAULT_FONT_FAMILY
                            )),
                    )
                })
                .child(cell_size_row(
                    app,
                    CellSizeField::FontSize,
                    "Default font size",
                    cx,
                ))
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child("Applies to cells without an explicit font. Sizes are in points."),
                )
                .child(
                    div()
                        .id("pref-reset-font")
                        .keyboard(app, "pref-reset-font")
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(accent)
                        .on_click(cx.listener(|this, _, _, cx| {
                            crate::settings::update_user_settings(cx, |settings| {
                                settings.appearance.default_font_family = Setting::Inherit;
                                settings.appearance.default_font_size = Setting::Inherit;
                            });
                            this.ui.cell_size_input = Default::default();
                            cx.notify();
                        }))
                        .child("Reset font"),
                ),
        ),
        PreferencesPage::Editing => content
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(14.))
                    .child(section_header("Editing", text_primary))
                    // After Enter row
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(row_label("After Enter", text_primary))
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_1()
                                    .child(enter_option(
                                        app,
                                        "Down",
                                        EnterBehavior::MoveDown,
                                        enter_behavior,
                                        accent,
                                        text_primary,
                                        text_muted,
                                        cx,
                                    ))
                                    .child(enter_option(
                                        app,
                                        "Right",
                                        EnterBehavior::MoveRight,
                                        enter_behavior,
                                        accent,
                                        text_primary,
                                        text_muted,
                                        cx,
                                    ))
                                    .child(enter_option(
                                        app,
                                        "Stay",
                                        EnterBehavior::Stay,
                                        enter_behavior,
                                        accent,
                                        text_primary,
                                        text_muted,
                                        cx,
                                    )),
                            ),
                    )
                    // Ctrl+V row. The behaviour has existed since
                    // the purist cycle and was reachable only from
                    // the command palette — but "make paste-values
                    // the default" is a settings question, and
                    // people look for it in settings.
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(row_label("Default paste", text_primary))
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_1()
                                    .child(paste_default_option(
                                        app,
                                        "Contents",
                                        false,
                                        paste_values_by_default,
                                        accent,
                                        text_primary,
                                        text_muted,
                                        cx,
                                    ))
                                    .child(paste_default_option(
                                        app,
                                        "Values only",
                                        true,
                                        paste_values_by_default,
                                        accent,
                                        text_primary,
                                        text_muted,
                                        cx,
                                    )),
                            ),
                    ),
            )
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(14.))
                    .child(section_header("Keyboard", text_primary))
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
                                    .child(row_label("Keyboard hints", text_primary))
                                    .child(
                                        preference_toggle(
                                            app,
                                            "pref-keyboard-hints-cb",
                                            keyboard_hints,
                                            accent,
                                            text_muted.opacity(0.25),
                                        )
                                        .on_click(
                                            cx.listener(move |_this, _, _, cx| {
                                                let new_value = !keyboard_hints;
                                                cx.update_global::<SettingsStore, _>(|store, _| {
                                                    store
                                                        .user_settings_mut()
                                                        .navigation
                                                        .keyboard_hints = Setting::Value(new_value);
                                                    store.save();
                                                });
                                                cx.notify();
                                            }),
                                        ),
                                    ),
                            )
                            .child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(text_muted)
                                    .child("Press g then type letters to jump to any cell"),
                            ),
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
                                    .child(row_label("Vim mode", text_primary))
                                    .child(
                                        preference_toggle(
                                            app,
                                            "pref-vim-mode-cb",
                                            vim_mode,
                                            accent,
                                            text_muted.opacity(0.25),
                                        )
                                        .on_click(
                                            cx.listener(move |_this, _, _, cx| {
                                                let new_value = !vim_mode;
                                                cx.update_global::<SettingsStore, _>(|store, _| {
                                                    store.user_settings_mut().navigation.vim_mode =
                                                        Setting::Value(new_value);
                                                    store.save();
                                                });
                                                cx.notify();
                                            }),
                                        ),
                                    ),
                            )
                            .child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(text_muted)
                                    .child("Navigate with h/j/k/l, press i to edit"),
                            ),
                    )
                    // Modifier style row (macOS only)
                    .when(cfg!(target_os = "macos"), |d: Div| {
                        d.child(
                            div()
                                .flex()
                                .flex_col()
                                .gap(px(2.0))
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label("Shortcut key", text_primary))
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_1()
                                                .child(modifier_option(
                                                    app,
                                                    "Cmd",
                                                    ModifierStyle::Platform,
                                                    modifier_style,
                                                    accent,
                                                    text_primary,
                                                    text_muted,
                                                    cx,
                                                ))
                                                .child(modifier_option(
                                                    app,
                                                    "Ctrl",
                                                    ModifierStyle::Ctrl,
                                                    modifier_style,
                                                    accent,
                                                    text_primary,
                                                    text_muted,
                                                    cx,
                                                )),
                                        ),
                                )
                                .child(div().text_size(px(11.0)).text_color(text_muted).child(
                                    "Use Ctrl for Windows-style shortcuts. Requires restart.",
                                )),
                        )
                        // Alt accelerators toggle (within same macOS-only block)
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap(px(2.0))
                                .mt_2()
                                .child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .justify_between()
                                        .child(row_label(
                                            "Excel-style Option (Alt) shortcuts",
                                            text_primary,
                                        ))
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_1()
                                                .child(alt_accel_option(
                                                    app,
                                                    "Off",
                                                    AltAccelerators::Disabled,
                                                    alt_accelerators,
                                                    accent,
                                                    text_primary,
                                                    text_muted,
                                                    cx,
                                                ))
                                                .child(alt_accel_option(
                                                    app,
                                                    "On",
                                                    AltAccelerators::Enabled,
                                                    alt_accelerators,
                                                    accent,
                                                    text_primary,
                                                    text_muted,
                                                    cx,
                                                )),
                                        ),
                                )
                                .child(div().text_size(px(11.0)).text_color(text_muted).child(
                                    "Option+F for File, Option+E for Edit. Requires restart.",
                                )),
                        )
                    }),
            ),
        PreferencesPage::Advanced => content
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(14.))
                    .child(section_header("Tips", text_primary))
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(row_label("Show dismissed tips again", text_primary))
                            .child(
                                div()
                                    .id("pref-reset-tips-btn")
                                    .keyboard(app, "pref-reset-tips-btn")
                                    .px_3()
                                    .py(px(6.0))
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
                                        this.status_message =
                                            Some("All tips have been reset".to_string());
                                        cx.notify();
                                    }))
                                    .child("Reset all tips"),
                            ),
                    ),
            )
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(14.))
                    .child(section_header("AI assistant", text_primary))
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(row_label("Connection settings", text_primary))
                            .child(
                                div()
                                    .id("pref-ai-settings-btn")
                                    .keyboard(app, "pref-ai-settings-btn")
                                    .px_3()
                                    .py(px(6.0))
                                    .bg(accent.opacity(0.15))
                                    .rounded_md()
                                    .cursor_pointer()
                                    .text_size(px(11.0))
                                    .text_color(text_primary)
                                    .hover(|s| s.bg(accent.opacity(0.25)))
                                    .on_click(cx.listener(|this, _, window, cx| {
                                        this.close_preferences_keyboard(window, cx);
                                        this.show_ai_settings(cx);
                                    }))
                                    .child("Configure…"),
                            ),
                    ),
            )
            .child(
                div()
                    .flex()
                    .flex_col()
                    .gap(px(14.))
                    .child(section_header("Settings file", text_primary))
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(row_label("Edit preferences as JSON", text_primary))
                            .child(
                                div()
                                    .id("pref-open-json-btn")
                                    .keyboard(app, "pref-open-json-btn")
                                    .px_3()
                                    .py(px(6.0))
                                    .bg(accent.opacity(0.15))
                                    .rounded_md()
                                    .cursor_pointer()
                                    .text_size(px(11.0))
                                    .text_color(text_primary)
                                    .hover(|s| s.bg(accent.opacity(0.25)))
                                    .on_click(cx.listener(|this, _, window, cx| {
                                        if let Err(e) = open_settings_file() {
                                            this.status_message =
                                                Some(format!("Failed to open settings: {}", e));
                                        } else {
                                            this.status_message = Some(
                                                "Opened settings.json in system editor".to_string(),
                                            );
                                        }
                                        this.close_preferences_keyboard(window, cx);
                                    }))
                                    .child("Open settings.json"),
                            ),
                    ),
            ),
    };
    let navigation = div()
        .w(px(164.))
        .flex_shrink_0()
        .p_3()
        .flex()
        .flex_col()
        .gap_1()
        .border_r_1()
        .border_color(panel_border)
        .children(
            [
                (PreferencesPage::Appearance, "Appearance"),
                (PreferencesPage::Editing, "Editing & keyboard"),
                (PreferencesPage::Advanced, "Advanced"),
            ]
            .into_iter()
            .enumerate()
            .map(|(index, (target, label))| {
                div()
                    .id(("preferences-page", index))
                    .keyboard(app, SIDEBAR_IDS[index])
                    .px_3()
                    .py(px(10.))
                    .rounded_md()
                    .text_size(px(12.))
                    .cursor_pointer()
                    .text_color(if page == target { accent } else { text_muted })
                    .font_weight(if page == target {
                        FontWeight::SEMIBOLD
                    } else {
                        FontWeight::NORMAL
                    })
                    .when(page == target, |d| d.bg(accent.opacity(0.1)))
                    .hover(move |d| d.bg(accent.opacity(0.08)).text_color(text_primary))
                    .on_click(cx.listener(move |this, _, _, cx| {
                        if !this.apply_cell_size_input(cx) {
                            return;
                        }
                        this.preferences_page = target;
                        this.preferences_keyboard
                            .scroll
                            .set_offset(point(px(0.), px(0.)));
                        cx.notify();
                    }))
                    .child(label)
            }),
        );

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0., 0., 0., 0.4))
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(|this, _, window, cx| this.close_preferences_keyboard(window, cx)),
        )
        .child(
            div()
                .id("preferences-panel")
                .w(px(760.))
                .max_w((app.window_size.width - px(32.)).max(px(280.)))
                .h(px(640.))
                .max_h((app.window_size.height - px(32.)).max(px(200.)))
                .bg(panel_bg)
                .text_color(text_primary)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_xl()
                .overflow_hidden()
                .flex()
                .flex_col()
                .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
                .on_scroll_wheel(|_, _, cx| cx.stop_propagation())
                .child(
                    div()
                        .px(px(24.))
                        .py(px(18.))
                        .flex_shrink_0()
                        .flex()
                        .items_center()
                        .justify_between()
                        .border_b_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_1()
                                .child(
                                    div()
                                        .text_size(px(18.))
                                        .font_weight(FontWeight::SEMIBOLD)
                                        .child("Preferences"),
                                )
                                .child(
                                    div()
                                        .text_size(px(12.))
                                        .text_color(text_muted)
                                        .child("Changes save automatically. Enter applies a size."),
                                ),
                        )
                        .child(
                            div()
                                .id("preferences-close")
                                .keyboard(app, "preferences-close")
                                .px_2()
                                .h(px(28.))
                                .gap_2()
                                .flex()
                                .items_center()
                                .justify_center()
                                .rounded_md()
                                .text_size(px(20.))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(move |d| {
                                    d.bg(panel_border.opacity(0.35)).text_color(text_primary)
                                })
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.close_preferences_keyboard(window, cx)
                                }))
                                .child(keycap("Esc", text_muted, panel_border))
                                .child("×"),
                        ),
                )
                .child(
                    div().flex().flex_1().min_h_0().child(navigation).child(
                        div()
                            .id(("preferences-content", page as usize))
                            .track_scroll(&app.preferences_keyboard.scroll)
                            .flex_1()
                            .min_w_0()
                            .h_full()
                            .overflow_y_scroll()
                            .p(px(24.))
                            .child(content),
                    ),
                )
                .child(
                    div()
                        .px(px(24.))
                        .py(px(14.))
                        .flex_shrink_0()
                        .flex()
                        .items_center()
                        .justify_between()
                        .border_t_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .text_size(px(11.))
                                .text_color(text_muted)
                                .child("Tab to move · Space / Enter to activate"),
                        )
                        .child(
                            crate::ui::Button::new("preferences-done", "Done")
                                .primary(accent, gpui::white())
                                .keyboard(app, "preferences-done")
                                .on_click(cx.listener(|this, _, window, cx| {
                                    if this.apply_cell_size_input(cx) {
                                        this.close_preferences_keyboard(window, cx);
                                    }
                                })),
                        ),
                ),
        )
}

/// Section header (e.g., "APPEARANCE")
fn section_header(title: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .mb_2()
        .text_size(px(14.0))
        .text_color(text_color)
        .font_weight(FontWeight::SEMIBOLD)
        .child(title)
}

/// Row label
fn row_label(label: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .text_size(px(13.0))
        .text_color(text_color)
        .child(label)
}

/// Enter behavior option button
fn enter_option(
    app: &Spreadsheet,
    label: &'static str,
    behavior: EnterBehavior,
    current: EnterBehavior,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = current == behavior;
    let bg = if is_selected {
        accent.opacity(0.2)
    } else {
        gpui::transparent_black()
    };
    let text = if is_selected {
        text_primary
    } else {
        text_muted
    };

    div()
        .id(SharedString::from(format!(
            "enter-{}",
            label.to_lowercase()
        )))
        .keyboard(
            app,
            match behavior {
                EnterBehavior::MoveDown => "enter-down",
                EnterBehavior::MoveRight => "enter-right",
                EnterBehavior::Stay => "enter-stay",
            },
        )
        .px_2()
        .py(px(6.0))
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

/// Default-paste option button.
///
/// Whichever way this is set, the other paste is always one shortcut away:
/// Ctrl+Shift+V pastes values, and Paste Special (Ctrl+Alt+V) offers All,
/// Values, Formulas and Formats. The preference only decides which one the
/// bare Ctrl+V reaches for.
fn paste_default_option(
    app: &Spreadsheet,
    label: &'static str,
    value: bool,
    current: bool,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = current == value;
    let bg = if is_selected {
        accent.opacity(0.2)
    } else {
        gpui::transparent_black()
    };
    let text = if is_selected {
        text_primary
    } else {
        text_muted
    };

    div()
        .id(SharedString::from(format!("paste-default-{}", value)))
        .keyboard(
            app,
            if value {
                "paste-default-true"
            } else {
                "paste-default-false"
            },
        )
        .px_2()
        .py(px(6.0))
        .rounded_sm()
        .bg(bg)
        .cursor_pointer()
        .text_size(px(11.0))
        .text_color(text)
        .hover(|s| s.bg(accent.opacity(0.1)))
        .on_click(cx.listener(move |_this, _, _, cx| {
            cx.update_global::<SettingsStore, _>(|store, _| {
                store.user_settings_mut().editing.paste_values_by_default = Setting::Value(value);
                store.save();
            });
            cx.notify();
        }))
        .child(label)
}

/// Modifier style option button (macOS only)
fn modifier_option(
    app: &Spreadsheet,
    label: &'static str,
    style: ModifierStyle,
    current: ModifierStyle,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = current == style;
    let bg = if is_selected {
        accent.opacity(0.2)
    } else {
        gpui::transparent_black()
    };
    let text = if is_selected {
        text_primary
    } else {
        text_muted
    };

    div()
        .id(SharedString::from(format!(
            "modifier-{}",
            label.to_lowercase()
        )))
        .keyboard(
            app,
            if style == ModifierStyle::Platform {
                "modifier-cmd"
            } else {
                "modifier-ctrl"
            },
        )
        .px_2()
        .py(px(6.0))
        .rounded_sm()
        .bg(bg)
        .cursor_pointer()
        .text_size(px(11.0))
        .text_color(text)
        .hover(|s| s.bg(accent.opacity(0.1)))
        .on_click(cx.listener(move |this, _, _, cx| {
            cx.update_global::<SettingsStore, _>(|store, _| {
                store.user_settings_mut().navigation.modifier_style = Setting::Value(style);
                store.save();
            });
            // Notify user that restart is needed
            this.status_message = Some("Restart VisiGrid to apply shortcut key change".to_string());
            cx.notify();
        }))
        .child(label)
}

/// Alt accelerators option button (macOS only)
fn alt_accel_option(
    app: &Spreadsheet,
    label: &'static str,
    mode: AltAccelerators,
    current: AltAccelerators,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = current == mode;
    let bg = if is_selected {
        accent.opacity(0.2)
    } else {
        gpui::transparent_black()
    };
    let text = if is_selected {
        text_primary
    } else {
        text_muted
    };

    div()
        .id(SharedString::from(format!(
            "alt-accel-{}",
            label.to_lowercase()
        )))
        .keyboard(
            app,
            if mode == AltAccelerators::Enabled {
                "alt-accel-on"
            } else {
                "alt-accel-off"
            },
        )
        .px_2()
        .py(px(6.0))
        .rounded_sm()
        .bg(bg)
        .cursor_pointer()
        .text_size(px(11.0))
        .text_color(text)
        .hover(|s| s.bg(accent.opacity(0.1)))
        .on_click(cx.listener(move |this, _, _, cx| {
            cx.update_global::<SettingsStore, _>(|store, _| {
                store.user_settings_mut().navigation.alt_accelerators = Setting::Value(mode);
                store.save();
            });
            // Notify user that restart is needed
            this.status_message = Some("Restart VisiGrid to apply shortcut change".to_string());
            cx.notify();
        }))
        .child(label)
}

fn cell_size_row(
    app: &Spreadsheet,
    field: CellSizeField,
    label: &'static str,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let sizes = crate::settings::CellSizeDefaults::from_user(crate::settings::user_settings(cx));
    let input = &app.ui.cell_size_input;
    let editing = input.field == Some(field);
    let text = if editing {
        input.text.clone()
    } else {
        field.value(sizes, app.cell_font.size).to_string()
    };
    let accent = app.token(TokenKey::Accent);
    let row = div()
        .flex()
        .items_center()
        .justify_between()
        .child(row_label(label, app.token(TokenKey::TextPrimary)))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .id(match field {
                            CellSizeField::ColumnWidth => "pref-column-width",
                            CellSizeField::RowHeight => "pref-row-height",
                            CellSizeField::FontSize => "pref-font-size",
                        })
                        .keyboard(app, field_id(field))
                        .w(px(64.0))
                        .px_2()
                        .py(px(6.0))
                        .rounded_sm()
                        .border_1()
                        .border_color(if editing {
                            accent
                        } else {
                            app.token(TokenKey::EditorBorder)
                        })
                        .bg(app.token(TokenKey::EditorBg))
                        .text_size(px(12.0))
                        .text_color(app.token(TokenKey::TextPrimary))
                        .cursor_text()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(move |this, _, window, cx| {
                                this.begin_cell_size_input(field, cx);
                                window
                                    .focus(&this.preferences_keyboard.handles[field_id(field)], cx);
                                cx.stop_propagation();
                            }),
                        )
                        .child(
                            div()
                                .when(editing && input.all_selected, |d| d.bg(accent.opacity(0.3)))
                                .child(if editing && !input.all_selected {
                                    format!("{text}|")
                                } else {
                                    text
                                }),
                        ),
                )
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(app.token(TokenKey::TextMuted))
                        .child(field.unit()),
                )
                .when(editing, |d| {
                    d.child(
                        div()
                            .id("pref-apply-cell-size")
                            .keyboard(app, "pref-apply-cell-size")
                            .px_2()
                            .py(px(6.0))
                            .rounded_sm()
                            .bg(accent.opacity(0.15))
                            .text_size(px(11.0))
                            .text_color(app.token(TokenKey::TextPrimary))
                            .cursor_pointer()
                            .on_click(cx.listener(move |this, _, window, cx| {
                                this.apply_cell_size_input(cx);
                                window
                                    .focus(&this.preferences_keyboard.handles[field_id(field)], cx);
                            }))
                            .child("Apply"),
                    )
                }),
        );
    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(row)
        .when(editing, |d| {
            d.when_some(input.error.clone(), |d, error| {
                d.child(
                    div()
                        .text_size(px(11.))
                        .text_color(app.token(TokenKey::Error))
                        .child(error),
                )
            })
        })
}

/// A shared on/off control; settings handlers remain with their rows.
fn preference_toggle(
    app: &Spreadsheet,
    id: &'static str,
    enabled: bool,
    accent: Hsla,
    border: Hsla,
) -> Stateful<Div> {
    div()
        .id(id)
        .keyboard(app, id)
        .w(px(36.))
        .h(px(22.))
        .flex_shrink_0()
        .px(px(3.))
        .flex()
        .items_center()
        .rounded_full()
        .cursor_pointer()
        .bg(if enabled { accent } else { border })
        .when(enabled, |d| d.justify_end())
        .hover(|d| d.opacity(0.85))
        .child(
            div()
                .size(px(16.))
                .rounded_full()
                .bg(gpui::white())
                .shadow_sm(),
        )
}

const SIDEBAR_IDS: [&str; 3] = [
    "pref-page-appearance",
    "pref-page-editing",
    "pref-page-advanced",
];
const APPEARANCE_IDS: &[&str] = &[
    "pref-theme-btn",
    "pref-gridlines-cb",
    "pref-comment-previews-cb",
    "pref-column-width",
    "pref-row-height",
    "pref-reset-cell-sizes",
    "pref-font-family",
    "pref-font-size",
    "pref-reset-font",
];
const EDITING_IDS: &[&str] = &[
    "enter-down",
    "enter-right",
    "enter-stay",
    "paste-default-false",
    "paste-default-true",
    "pref-keyboard-hints-cb",
    "pref-vim-mode-cb",
];
const MAC_IDS: &[&str] = &[
    "modifier-cmd",
    "modifier-ctrl",
    "alt-accel-off",
    "alt-accel-on",
];
const ADVANCED_IDS: &[&str] = &[
    "pref-reset-tips-btn",
    "pref-ai-settings-btn",
    "pref-open-json-btn",
];

pub struct PreferencesKeyboard {
    handles: std::collections::HashMap<&'static str, FocusHandle>,
    bounds:
        std::rc::Rc<std::cell::RefCell<std::collections::HashMap<&'static str, Bounds<Pixels>>>>,
    scroll: ScrollHandle,
}
impl PreferencesKeyboard {
    pub fn new(cx: &mut App) -> Self {
        let ids = SIDEBAR_IDS
            .iter()
            .chain(APPEARANCE_IDS)
            .chain(EDITING_IDS)
            .chain(MAC_IDS)
            .chain(ADVANCED_IDS)
            .chain(
                [
                    "preferences-done",
                    "preferences-close",
                    "pref-apply-cell-size",
                ]
                .iter(),
            );
        Self {
            handles: ids.map(|id| (*id, cx.focus_handle())).collect(),
            bounds: Default::default(),
            scroll: ScrollHandle::new(),
        }
    }
}
fn field_id(field: CellSizeField) -> &'static str {
    match field {
        CellSizeField::ColumnWidth => "pref-column-width",
        CellSizeField::RowHeight => "pref-row-height",
        CellSizeField::FontSize => "pref-font-size",
    }
}
fn field_for_id(id: &str) -> Option<CellSizeField> {
    match id {
        "pref-column-width" => Some(CellSizeField::ColumnWidth),
        "pref-row-height" => Some(CellSizeField::RowHeight),
        "pref-font-size" => Some(CellSizeField::FontSize),
        _ => None,
    }
}
fn focus_order(page: PreferencesPage) -> Vec<&'static str> {
    let mut ids = vec![SIDEBAR_IDS[page as usize]];
    ids.extend_from_slice(match page {
        PreferencesPage::Appearance => APPEARANCE_IDS,
        PreferencesPage::Editing => EDITING_IDS,
        PreferencesPage::Advanced => ADVANCED_IDS,
    });
    if page == PreferencesPage::Editing && cfg!(target_os = "macos") {
        ids.extend_from_slice(MAC_IDS);
    }
    ids.extend_from_slice(&["preferences-done", "preferences-close"]);
    ids
}

trait PreferenceKeyboardElement {
    fn keyboard(self, app: &Spreadsheet, id: &'static str) -> Self;
}
impl PreferenceKeyboardElement for Stateful<Div> {
    fn keyboard(self, app: &Spreadsheet, id: &'static str) -> Self {
        let handle = app.preferences_keyboard.handles[id].clone();
        let bounds = app.preferences_keyboard.bounds.clone();
        let accent = app.token(TokenKey::Accent);
        self.relative().track_focus(&handle).child(
            canvas(
                move |rect, _, _| {
                    bounds.borrow_mut().insert(id, rect);
                },
                move |rect, _, window, _| {
                    if handle.is_focused(window) {
                        window.paint_quad(quad(
                            rect.dilate(px(2.)),
                            px(5.),
                            transparent_black(),
                            px(2.),
                            accent,
                            BorderStyle::Solid,
                        ));
                    }
                },
            )
            .absolute()
            .inset_0(),
        )
    }
}
fn keycap(label: &'static str, text: Hsla, border: Hsla) -> impl IntoElement {
    div()
        .px(px(5.))
        .py(px(2.))
        .border_1()
        .border_color(border)
        .rounded_sm()
        .text_size(px(10.))
        .text_color(text)
        .child(label)
}
impl Spreadsheet {
    fn close_preferences_keyboard(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.hide_preferences(cx);
        window.focus(&self.focus_handle, cx);
    }
    fn focus_preference(&mut self, id: &'static str, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(field) = field_for_id(id) {
            self.begin_cell_size_input(field, cx);
        }
        window.focus(&self.preferences_keyboard.handles[id], cx);
        if !SIDEBAR_IDS.contains(&id) && id != "preferences-done" && id != "preferences-close" {
            let bounds = self.preferences_keyboard.bounds.borrow().get(id).copied();
            if let Some(bounds) = bounds {
                let viewport = self.preferences_keyboard.scroll.bounds();
                let mut offset = self.preferences_keyboard.scroll.offset();
                if bounds.top() < viewport.top() + px(8.) {
                    offset.y += viewport.top() + px(8.) - bounds.top();
                } else if bounds.bottom() > viewport.bottom() - px(8.) {
                    offset.y -= bounds.bottom() - viewport.bottom() + px(8.);
                }
                self.preferences_keyboard.scroll.set_offset(offset);
            }
        }
        cx.notify();
    }
    pub fn preferences_key(
        &mut self,
        event: &KeyDownEvent,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        let key = event.keystroke.key.as_str();
        let modifiers = event.keystroke.modifiers;
        let focused = self
            .preferences_keyboard
            .handles
            .iter()
            .find(|(_, h)| h.is_focused(window))
            .map(|(id, _)| *id);
        if key == "escape" {
            self.close_preferences_keyboard(window, cx);
            cx.stop_propagation();
            return;
        }
        if key == "tab" && !modifiers.control && !modifiers.platform && !modifiers.alt {
            // Validate before moving focus; invalid input remains visible and editable.
            let ids = focus_order(self.preferences_page);
            if !self.apply_cell_size_input(cx) {
                cx.stop_propagation();
                return;
            }
            let index = focused.and_then(|id| ids.iter().position(|item| *item == id));
            let next = match (index, modifiers.shift) {
                (Some(i), true) => (i + ids.len() - 1) % ids.len(),
                (Some(i), false) => (i + 1) % ids.len(),
                (None, true) => ids.len() - 1,
                (None, false) => 0,
            };
            self.focus_preference(ids[next], window, cx);
            cx.stop_propagation();
            return;
        }
        if focused.is_none() || focused.is_some_and(|id| SIDEBAR_IDS.contains(&id)) {
            if matches!(key, "up" | "down") && !modifiers.modified() {
                let pages = [
                    PreferencesPage::Appearance,
                    PreferencesPage::Editing,
                    PreferencesPage::Advanced,
                ];
                let next = (self.preferences_page as usize + if key == "down" { 1 } else { pages.len() - 1 })
                    % pages.len();
                if self.apply_cell_size_input(cx) {
                    self.preferences_page = pages[next];
                    self.preferences_keyboard
                        .scroll
                        .set_offset(point(px(0.), px(0.)));
                    self.focus_preference(SIDEBAR_IDS[next], window, cx);
                }
                cx.stop_propagation();
            }
            return;
        }
        if let Some(field) = focused.and_then(field_for_id) {
            if self.ui.cell_size_input.field != Some(field) {
                self.begin_cell_size_input(field, cx);
            }
            let command = modifiers.control || modifiers.platform;
            if command && key == "v" {
                self.cell_size_input_paste(cx);
            } else if command && key == "a" {
                self.ui.cell_size_input.all_selected = true;
                cx.notify();
            } else {
                self.cell_size_input_key(
                    key,
                    event.keystroke.key_char.as_deref(),
                    command || modifiers.alt,
                    cx,
                );
            }
            cx.stop_propagation();
        }
        // Other controls use GPUI's native Enter/Space -> on_click handling.
    }
}
