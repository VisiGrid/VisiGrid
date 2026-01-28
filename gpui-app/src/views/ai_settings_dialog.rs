//! AI Settings dialog
//!
//! Configure AI provider, model, API keys, and privacy settings.
//! Uses the design system components from `ui/`.

use gpui::{*, prelude::FluentBuilder};
use std::time::Duration;
use crate::app::{Spreadsheet, AIProviderOption, AISettingsFocus, AITestStatus};
use crate::theme::TokenKey;
use crate::ui::{DialogFrame, DialogSize, Button, dialog_header_with_subtitle};

/// Render the AI settings dialog overlay
pub fn render_ai_settings_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);
    let editor_border = app.token(TokenKey::EditorBorder);
    let error_color = hsla(0.0, 0.7, 0.5, 1.0);
    let success_color = hsla(0.33, 0.7, 0.4, 1.0);

    let state = &app.ai_settings;
    let provider = state.provider;
    let model = state.model.clone();
    let endpoint = state.endpoint.clone();
    let privacy_mode = state.privacy_mode;
    let allow_proposals = state.allow_proposals;
    let focus = state.focus;
    let dropdown_open = state.provider_dropdown_open;
    let key_present = state.key_present;
    let key_source = state.key_source.clone();
    let key_input = state.key_input.clone();
    let effective_model = state.effective_model().to_string();
    let context_policy = state.context_policy();
    let needs_key = state.needs_api_key();
    let error = state.error.clone();
    let test_status = state.test_status.clone();

    // Backdrop (top-aligned, not centered)
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(60.0))
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_ai_settings(cx);
        }))
        .child(
            // Wrapper for focus tracking and keyboard handling
            div()
                .id("ai-settings-dialog")
                .key_context("AISettings")
                .track_focus(&app.focus_handle)
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    handle_key_event(this, event, cx);
                    cx.stop_propagation();
                }))
                .child(
                    DialogFrame::new(
                        render_body(
                            provider, &model, &endpoint, privacy_mode, allow_proposals,
                            focus, dropdown_open, key_present, &key_source, &key_input,
                            &effective_model, &context_policy, needs_key, &error, &test_status,
                            panel_bg, panel_border, text_primary, text_muted, accent,
                            editor_bg, editor_border, error_color, success_color, cx,
                        ),
                        panel_bg,
                        panel_border,
                    )
                    .size(DialogSize::Md)
                    .width(px(420.0))
                    .header(dialog_header_with_subtitle("AI Settings", "Esc to close", text_primary, text_muted))
                    .footer(render_footer(accent, panel_border, text_muted, cx))
                )
        )
}

/// Handle keyboard events for the dialog
fn handle_key_event(this: &mut Spreadsheet, event: &KeyDownEvent, cx: &mut Context<Spreadsheet>) {
    let key = &event.keystroke.key;
    let modifiers = &event.keystroke.modifiers;

    // Handle Ctrl+V / Cmd+V paste
    if (modifiers.control || modifiers.platform) && key.eq_ignore_ascii_case("v") {
        this.ai_settings_paste(cx);
        return;
    }

    match key.as_str() {
        "escape" => {
            if this.ai_settings.provider_dropdown_open {
                this.ai_settings.provider_dropdown_open = false;
                cx.notify();
            } else {
                this.hide_ai_settings(cx);
            }
        }
        "enter" => {
            if !this.ai_settings.provider_dropdown_open {
                if !this.ai_settings.key_input.is_empty() {
                    this.ai_settings_set_key(cx);
                } else {
                    this.apply_ai_settings(cx);
                }
            }
        }
        "tab" => {
            this.ai_settings_tab(event.keystroke.modifiers.shift, cx);
        }
        "backspace" => {
            this.ai_settings_backspace(cx);
        }
        _ => {
            if let Some(c) = event.keystroke.key_char.as_ref().and_then(|s| s.chars().next()) {
                if !modifiers.control && !modifiers.alt {
                    this.ai_settings_type_char(c, cx);
                }
            }
        }
    }
}

/// Render the dialog body content
fn render_body(
    provider: AIProviderOption,
    model: &str,
    endpoint: &str,
    privacy_mode: bool,
    allow_proposals: bool,
    focus: AISettingsFocus,
    dropdown_open: bool,
    key_present: bool,
    key_source: &str,
    key_input: &str,
    effective_model: &str,
    context_policy: &str,
    needs_key: bool,
    error: &Option<String>,
    test_status: &AITestStatus,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    editor_border: Hsla,
    error_color: Hsla,
    success_color: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_4()
        // Provider section
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(section_label("PROVIDER", text_primary))
                .child(render_provider_dropdown(
                    provider, dropdown_open, focus == AISettingsFocus::Provider,
                    panel_bg, panel_border, text_primary, text_muted, accent, editor_bg, cx,
                ))
        )
        // Model section
        .when(provider != AIProviderOption::None, |d| {
            d.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_2()
                    .child(section_label("MODEL", text_primary))
                    .child(render_text_field(
                        model, effective_model, focus == AISettingsFocus::Model,
                        editor_bg, editor_border, text_primary, text_muted, accent,
                        cx, AISettingsFocus::Model,
                    ))
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted.opacity(0.7))
                            .child(format!("Leave empty for default ({})", effective_model))
                    )
            )
        })
        // Endpoint section (Local only)
        .when(provider == AIProviderOption::Local, |d| {
            d.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_2()
                    .child(section_label("ENDPOINT", text_primary))
                    .child(render_text_field(
                        endpoint, "http://localhost:11434", focus == AISettingsFocus::Endpoint,
                        editor_bg, editor_border, text_primary, text_muted, accent,
                        cx, AISettingsFocus::Endpoint,
                    ))
            )
        })
        // API Key section
        .when(needs_key, |d| {
            d.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_2()
                    .child(section_label("API KEY", text_primary))
                    .child(render_key_section(
                        key_present, key_source, key_input, focus == AISettingsFocus::KeyInput,
                        panel_border, text_primary, text_muted, accent, editor_bg, editor_border,
                        success_color, cx,
                    ))
            )
        })
        // Privacy section
        .when(provider != AIProviderOption::None, |d| {
            d.child(render_privacy_section(
                privacy_mode, allow_proposals, context_policy,
                text_muted, accent, editor_bg, editor_border, cx,
            ))
        })
        // Validation section
        .when(provider != AIProviderOption::None, |d| {
            d.child(render_validation_section(
                test_status, text_primary, text_muted, accent, success_color, error_color, cx,
            ))
        })
        // Error message
        .when(error.is_some(), |d| {
            d.child(
                div()
                    .px_3()
                    .py_2()
                    .rounded_sm()
                    .bg(error_color.opacity(0.1))
                    .text_size(px(11.0))
                    .text_color(error_color)
                    .child(error.clone().unwrap_or_default())
            )
        })
}

/// Render the footer with Cancel and Save buttons
fn render_footer(
    accent: Hsla,
    panel_border: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .justify_end()
        .gap_2()
        .child(
            Button::new("ai-cancel-btn", "Cancel")
                .secondary(panel_border, text_muted)
                .on_click(cx.listener(|this, _, _, cx| {
                    this.hide_ai_settings(cx);
                }))
        )
        .child(
            Button::new("ai-save-btn", "Save")
                .primary(accent, gpui::white())
                .on_click(cx.listener(|this, _, _, cx| {
                    this.apply_ai_settings(cx);
                }))
        )
}

/// Section label (e.g., "PROVIDER")
fn section_label(title: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .text_size(px(10.0))
        .text_color(text_color)
        .font_weight(FontWeight::SEMIBOLD)
        .child(title)
}

/// Privacy section with checkboxes
fn render_privacy_section(
    privacy_mode: bool,
    allow_proposals: bool,
    context_policy: &str,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    editor_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let text_primary = text_muted; // Reuse for consistency

    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(section_label("PRIVACY", text_primary))
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .child("Privacy mode")
                )
                .child(render_checkbox(
                    privacy_mode, accent, editor_bg, editor_border, "ai-privacy-mode", cx,
                    |this, _, _, cx| {
                        this.ai_settings.privacy_mode = !this.ai_settings.privacy_mode;
                        cx.notify();
                    },
                ))
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text_muted.opacity(0.7))
                .child(context_policy.to_string())
        )
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .mt_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .child("Allow AI proposals")
                )
                .child(render_checkbox(
                    allow_proposals, accent, editor_bg, editor_border, "ai-allow-proposals", cx,
                    |this, _, _, cx| {
                        this.ai_settings.allow_proposals = !this.ai_settings.allow_proposals;
                        cx.notify();
                    },
                ))
        )
        .when(!allow_proposals, |d| {
            d.child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted.opacity(0.7))
                    .child("AI can explain but cannot propose cell changes")
            )
        })
        .when(allow_proposals, |d| {
            d.child(
                div()
                    .text_size(px(10.0))
                    .text_color(hsla(0.08, 0.8, 0.5, 1.0))
                    .child("AI can propose changes (requires confirmation)")
            )
        })
}

/// Validation section with test button
fn render_validation_section(
    test_status: &AITestStatus,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    success_color: Hsla,
    error_color: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(section_label("VALIDATION", text_primary))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .id("ai-test-btn")
                        .px_3()
                        .py(px(6.0))
                        .rounded_md()
                        .bg(accent.opacity(0.15))
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .hover(|s| s.bg(accent.opacity(0.25)))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.ai_settings_test_connection(cx);
                        }))
                        .child("Validate")
                )
                .child(render_test_status(test_status, text_muted, success_color, error_color))
        )
}

/// Provider dropdown
fn render_provider_dropdown(
    current: AIProviderOption,
    is_open: bool,
    has_focus: bool,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let options = [
        AIProviderOption::None,
        AIProviderOption::Local,
        AIProviderOption::OpenAI,
        AIProviderOption::Anthropic,
        AIProviderOption::Gemini,
        AIProviderOption::Grok,
    ];

    div()
        .relative()
        .child(
            div()
                .id("ai-provider-dropdown")
                .flex()
                .items_center()
                .justify_between()
                .px_3()
                .py(px(6.0))
                .rounded_md()
                .border_1()
                .border_color(if has_focus { accent } else { panel_border })
                .bg(editor_bg)
                .cursor_pointer()
                .text_size(px(12.0))
                .text_color(text_primary)
                .on_click(cx.listener(|this, _, _, cx| {
                    this.ai_settings.provider_dropdown_open = !this.ai_settings.provider_dropdown_open;
                    this.ai_settings.focus = AISettingsFocus::Provider;
                    cx.notify();
                }))
                .child(current.label())
                .child(
                    div()
                        .text_color(text_muted)
                        .child(if is_open { "^" } else { "v" })
                )
        )
        .when(is_open, |d| {
            d.child(
                div()
                    .id("ai-provider-dropdown-list")
                    .absolute()
                    .top(px(32.0))
                    .left_0()
                    .right_0()
                    .bg(panel_bg)
                    .border_1()
                    .border_color(panel_border)
                    .rounded_md()
                    .shadow_lg()
                    .py_1()
                    // Stop clicks from reaching the backdrop
                    .on_mouse_down(MouseButton::Left, |_, _, cx| {
                        cx.stop_propagation();
                    })
                    .children(options.iter().map(|&option| {
                        let is_selected = option == current;
                        div()
                            .id(ElementId::Name(format!("ai-provider-{:?}", option).into()))
                            .px_3()
                            .py(px(4.0))
                            .mx_1()
                            .rounded_sm()
                            .cursor_pointer()
                            .text_size(px(12.0))
                            .text_color(if is_selected { text_primary } else { text_muted })
                            .bg(if is_selected { accent.opacity(0.15) } else { gpui::transparent_black() })
                            .hover(|s| s.bg(accent.opacity(0.1)))
                            .on_click(cx.listener(move |this, _, _, cx| {
                                this.ai_settings.provider = option;
                                this.ai_settings.provider_dropdown_open = false;
                                this.ai_settings.load_from_config();
                                this.ai_settings.provider = option;
                                cx.notify();
                            }))
                            .child(option.label())
                    }))
            )
        })
}

/// Text field with placeholder
fn render_text_field(
    value: &str,
    placeholder: &str,
    has_focus: bool,
    editor_bg: Hsla,
    editor_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
    focus_target: AISettingsFocus,
) -> impl IntoElement {
    // Show placeholder only when empty AND not focused
    let show_placeholder = value.is_empty() && !has_focus;
    let display_value = if show_placeholder {
        placeholder.to_string()
    } else {
        value.to_string()
    };

    div()
        .id(ElementId::Name(format!("ai-field-{:?}", focus_target).into()))
        .flex()
        .items_center()
        .px_3()
        .py(px(6.0))
        .rounded_md()
        .border_1()
        .border_color(if has_focus { accent } else { editor_border })
        .bg(editor_bg)
        .cursor_text()
        .text_size(px(12.0))
        .text_color(if show_placeholder { text_muted.opacity(0.5) } else { text_primary })
        .on_click(cx.listener(move |this, _, _, cx| {
            this.ai_settings.focus = focus_target;
            cx.notify();
        }))
        .child(display_value)
        .when(has_focus, |el| {
            el.child(
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
        })
}

/// API key section with status and input
fn render_key_section(
    key_present: bool,
    key_source: &str,
    key_input: &str,
    has_focus: bool,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    editor_border: Hsla,
    success_color: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Status:")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(if key_present { success_color } else { text_muted })
                                .child(if key_present {
                                    format!("Set ({})", key_source)
                                } else {
                                    "Not set".to_string()
                                })
                        )
                )
                .when(key_present, |d| {
                    d.child(
                        div()
                            .id("ai-clear-key-btn")
                            .px_2()
                            .py(px(2.0))
                            .rounded_sm()
                            .cursor_pointer()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .hover(|s| s.bg(panel_border.opacity(0.3)))
                            .on_click(cx.listener(|this, _, _, cx| {
                                this.ai_settings_clear_key(cx);
                            }))
                            .child("Clear")
                    )
                })
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .flex_1()
                        .id("ai-key-input")
                        .flex()
                        .items_center()
                        .px_3()
                        .py(px(6.0))
                        .rounded_md()
                        .border_1()
                        .border_color(if has_focus { accent } else { editor_border })
                        .bg(editor_bg)
                        .cursor_text()
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.ai_settings.focus = AISettingsFocus::KeyInput;
                            cx.notify();
                        }))
                        .child(if key_input.is_empty() && !has_focus {
                            // Show placeholder only when empty and not focused
                            div()
                                .text_color(text_muted.opacity(0.5))
                                .child("Enter API key...")
                                .into_any_element()
                        } else if key_input.is_empty() {
                            // Empty but focused - show nothing (just cursor)
                            div().into_any_element()
                        } else {
                            // Has content - show masked
                            div().child("*".repeat(key_input.len().min(40))).into_any_element()
                        })
                        .when(has_focus, |el| {
                            el.child(
                                div()
                                    .w(px(1.0))
                                    .h(px(14.0))
                                    .bg(text_primary)
                                    .ml(px(1.0))
                                    .with_animation(
                                        "cursor-blink-key",
                                        Animation::new(Duration::from_millis(530))
                                            .repeat()
                                            .with_easing(pulsating_between(0.0, 1.0)),
                                        |div, delta| {
                                            let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                            div.opacity(opacity)
                                        },
                                    )
                            )
                        })
                )
                .child(
                    div()
                        .id("ai-set-key-btn")
                        .px_3()
                        .py(px(6.0))
                        .rounded_md()
                        .bg(accent.opacity(0.15))
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .hover(|s| s.bg(accent.opacity(0.25)))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.ai_settings_set_key(cx);
                        }))
                        .child("Set Key")
                )
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text_muted.opacity(0.7))
                .child("Keys are stored securely in system keychain")
        )
}

/// Checkbox widget
fn render_checkbox(
    checked: bool,
    accent: Hsla,
    editor_bg: Hsla,
    editor_border: Hsla,
    id: &'static str,
    cx: &mut Context<Spreadsheet>,
    on_click: impl Fn(&mut Spreadsheet, &ClickEvent, &mut Window, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    div()
        .id(id)
        .size(px(16.0))
        .rounded_sm()
        .border_1()
        .border_color(if checked { accent } else { editor_border })
        .bg(if checked { accent } else { editor_bg })
        .cursor_pointer()
        .flex()
        .items_center()
        .justify_center()
        .on_click(cx.listener(on_click))
        .child(if checked {
            div()
                .text_size(px(10.0))
                .text_color(gpui::white())
                .child("✓")
                .into_any_element()
        } else {
            div().into_any_element()
        })
}

/// Test status display
fn render_test_status(
    status: &AITestStatus,
    text_muted: Hsla,
    success_color: Hsla,
    error_color: Hsla,
) -> impl IntoElement {
    match status {
        AITestStatus::Idle => {
            div()
                .text_size(px(11.0))
                .text_color(text_muted.opacity(0.5))
                .child("Not tested")
        }
        AITestStatus::Testing => {
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Testing...")
        }
        AITestStatus::Success(msg) => {
            div()
                .text_size(px(11.0))
                .text_color(success_color)
                .child(msg.clone())
        }
        AITestStatus::Error(msg) => {
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(msg.clone())
        }
    }
}
