//! AI Dialog (Insert Formula + Analyze)
//!
//! Unified dialog for AI verbs. Branches on AiVerb to show
//! Insert Formula (single-cell write) or Analyze (read-only).
//! Uses the design system components from `ui/`.

use gpui::*;
use gpui::prelude::FluentBuilder;
use std::time::Duration;
use crate::app::{Spreadsheet, AiVerb, AskAIStatus, AskAIContextMode};
use crate::theme::TokenKey;
use crate::ui::{DialogFrame, DialogSize, Button};

/// Render the Ask AI dialog
pub fn render_ask_ai_dialog(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);
    let success_color = hsla(0.35, 0.7, 0.45, 1.0);
    let error_color = hsla(0.0, 0.7, 0.5, 1.0);
    let warning_color = hsla(0.12, 0.8, 0.5, 1.0);

    // Extract state
    let state = &app.ask_ai;
    let verb = state.verb;
    let question = state.question.clone();
    let context_summary = state.context_summary.clone();
    let context_mode = state.context_mode;
    let context_selector_open = state.context_selector_open;
    let sent_context = state.sent_context.clone();
    let sent_panel_expanded = state.sent_panel_expanded;
    let status = state.status.clone();
    let explanation = state.explanation.clone();
    let formula = state.formula.clone();
    let formula_valid = state.formula_valid;
    let formula_error = state.formula_error.clone();
    let response_text = state.response_text.clone();
    let warnings = state.warnings.clone();
    let error = state.error.clone();
    let last_insertion = state.last_insertion.clone();
    let has_raw_response = state.raw_response.is_some();

    let is_loading = state.is_loading();
    let has_response = matches!(status, AskAIStatus::Success);
    let has_error = matches!(status, AskAIStatus::Error(_));
    let can_insert = verb == AiVerb::InsertFormula && state.can_insert();
    let can_retry = state.can_retry() && (has_response || has_error);

    // Backdrop (centered)
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(rgba(0x00000080))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_ask_ai(cx);
        }))
        .child(
            // Wrapper for keyboard handling
            div()
                .id("ask-ai-dialog")
                .key_context("AskAI")
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
                            verb, panel_bg, panel_border, text_primary, text_muted, accent,
                            editor_bg, success_color, error_color, warning_color,
                            question, context_summary, context_mode, context_selector_open,
                            sent_context, sent_panel_expanded, explanation, formula.clone(),
                            formula_valid, formula_error, response_text.clone(),
                            warnings, error.clone(),
                            last_insertion, is_loading, has_response, has_error,
                            has_raw_response, cx,
                        ),
                        panel_bg,
                        panel_border,
                    )
                    .size(DialogSize::Lg)
                    .width(px(580.0))
                    .max_height(px(650.0))
                    .header(render_header(verb, text_primary, text_muted, panel_border, cx))
                    .footer(render_footer(
                        verb, panel_border, text_primary, text_muted, accent, success_color,
                        is_loading, can_insert, can_retry, has_error, has_response,
                        formula, response_text.clone(), cx,
                    ))
                )
        )
}

/// Handle keyboard events
fn handle_key_event(this: &mut Spreadsheet, event: &KeyDownEvent, cx: &mut Context<Spreadsheet>) {
    let key = event.keystroke.key.as_str();
    let modifiers = &event.keystroke.modifiers;

    // Handle Ctrl+V / Cmd+V paste
    if (modifiers.control || modifiers.platform) && key.eq_ignore_ascii_case("v") {
        this.ask_ai_paste(cx);
        return;
    }

    match key {
        "escape" => {
            if this.ask_ai.context_selector_open {
                this.ask_ai.context_selector_open = false;
                cx.notify();
            } else {
                this.hide_ask_ai(cx);
            }
        }
        "enter" => {
            if modifiers.platform || modifiers.control {
                this.ask_ai_submit(cx);
            }
        }
        "backspace" => {
            this.ask_ai_backspace(cx);
        }
        _ => {
            // Don't type if modifier keys are held (except shift)
            if modifiers.control || modifiers.platform || modifiers.alt {
                return;
            }
            if let Some(c) = event.keystroke.key_char.as_ref() {
                for ch in c.chars() {
                    this.ask_ai_type_char(ch, cx);
                }
            }
        }
    }
}

/// Render the header with title, contract badge, and close button
fn render_header(
    verb: AiVerb,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let (title, badge_text) = match verb {
        AiVerb::InsertFormula => (
            "Insert Formula",
            "Single-cell write  Formula will be inserted into active cell.",
        ),
        AiVerb::Analyze => (
            "Analyze",
            "Read-only  No cells will be modified.",
        ),
    };

    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .text_size(px(14.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(title)
                )
                .child(
                    div()
                        .id("ask-ai-close")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(14.0))
                        .text_color(text_muted)
                        .hover(|s| s.bg(panel_border))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.hide_ask_ai(cx);
                        }))
                        .child("×")
                )
        )
        // Execution contract badge
        .child(
            div()
                .px_2()
                .py_1()
                .rounded_sm()
                .bg(panel_border.opacity(0.3))
                .text_size(px(10.0))
                .text_color(text_muted)
                .child(badge_text)
        )
}

/// Render the dialog body content
#[allow(clippy::too_many_arguments)]
fn render_body(
    verb: AiVerb,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    success_color: Hsla,
    error_color: Hsla,
    warning_color: Hsla,
    question: String,
    context_summary: String,
    context_mode: AskAIContextMode,
    context_selector_open: bool,
    sent_context: Option<crate::app::AskAISentContext>,
    sent_panel_expanded: bool,
    explanation: Option<String>,
    formula: Option<String>,
    formula_valid: bool,
    formula_error: Option<String>,
    response_text: Option<String>,
    warnings: Vec<String>,
    error: Option<String>,
    last_insertion: Option<String>,
    is_loading: bool,
    has_response: bool,
    _has_error: bool,
    has_raw_response: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let mut content = div()
        .flex()
        .flex_col()
        .gap_3()
        .flex_1()
        .overflow_hidden();

    // Context selector
    content = content.child(render_context_selector(
        panel_bg, panel_border, text_primary, text_muted, accent,
        context_summary, context_mode, context_selector_open, cx,
    ));

    // Warnings
    if !warnings.is_empty() {
        content = content.child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(warnings.into_iter().map(move |w| {
                    div()
                        .text_size(px(11.0))
                        .text_color(warning_color)
                        .child(format!("⚠ {}", w))
                }))
        );
    }

    // Question input with blinking cursor
    let placeholder = match verb {
        AiVerb::InsertFormula => "What formula do you need?",
        AiVerb::Analyze => "What would you like to know about this data?",
    };
    let tip_text = match verb {
        AiVerb::InsertFormula => {
            if cfg!(target_os = "macos") {
                "Tip: Ask for a single spreadsheet formula when possible. Press \u{2318}+Enter to submit."
            } else {
                "Tip: Ask for a single spreadsheet formula when possible. Press Ctrl+Enter to submit."
            }
        }
        AiVerb::Analyze => {
            if cfg!(target_os = "macos") {
                "Ask about patterns, anomalies, summaries, or comparisons. Press \u{2318}+Enter to submit."
            } else {
                "Ask about patterns, anomalies, summaries, or comparisons. Press Ctrl+Enter to submit."
            }
        }
    };

    content = content.child(
        div()
            .flex()
            .flex_col()
            .gap_1()
            .child(section_label("QUESTION", text_muted))
            .child(
                div()
                    .id("ask-ai-question")
                    .w_full()
                    .min_h(px(60.0))
                    .px_3()
                    .py_2()
                    .rounded_md()
                    .border_1()
                    .border_color(accent)
                    .bg(editor_bg)
                    .flex()
                    .items_start()
                    .child(
                        div()
                            .text_size(px(13.0))
                            .text_color(if question.is_empty() { text_muted } else { text_primary })
                            .child(if question.is_empty() {
                                placeholder.to_string()
                            } else {
                                question.clone()
                            })
                    )
                    // Blinking cursor
                    .child(
                        div()
                            .w(px(1.0))
                            .h(px(16.0))
                            .bg(text_primary)
                            .ml(px(1.0))
                            .with_animation(
                                "ask-ai-cursor",
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
            // Prompt suggestion chips (Analyze only, hidden once user types)
            .when(verb == AiVerb::Analyze && question.is_empty() && !is_loading && !has_response, |el| {
                el.child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(prompt_chip("Summarize", "Summarize this data", accent, panel_border, text_primary, cx))
                        .child(prompt_chip("Find anomalies", "Are there any anomalies or outliers?", accent, panel_border, text_primary, cx))
                        .child(prompt_chip("Compare categories", "Compare the top categories", accent, panel_border, text_primary, cx))
                )
            })
            .child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted)
                    .child(tip_text)
            )
    );

    // Response section (branches on verb)
    if has_response {
        match verb {
            AiVerb::InsertFormula => {
                // Explanation
                if let Some(expl) = explanation {
                    content = content.child(
                        div()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .child(section_label("EXPLANATION", text_muted))
                            .child(
                                div()
                                    .w_full()
                                    .px_3()
                                    .py_2()
                                    .rounded_md()
                                    .bg(editor_bg.opacity(0.5))
                                    .text_size(px(12.0))
                                    .text_color(text_primary)
                                    .child(expl)
                            )
                    );
                }

                // Formula
                if let Some(f) = formula {
                    let validation_label = if formula_valid {
                        div()
                            .text_size(px(10.0))
                            .text_color(success_color)
                            .child("\u{2713} Valid")
                    } else if let Some(err) = formula_error {
                        div()
                            .text_size(px(10.0))
                            .text_color(error_color)
                            .child(format!("\u{2717} {}", err))
                    } else {
                        div()
                    };

                    content = content.child(
                        div()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(section_label("FORMULA", text_muted))
                                    .child(validation_label)
                            )
                            .child(
                                div()
                                    .id("ask-ai-formula")
                                    .w_full()
                                    .px_3()
                                    .py_2()
                                    .rounded_md()
                                    .border_1()
                                    .border_color(if formula_valid { success_color.opacity(0.5) } else { error_color.opacity(0.5) })
                                    .bg(editor_bg)
                                    .text_size(px(13.0))
                                    .font_family("monospace")
                                    .text_color(text_primary)
                                    .child(f)
                            )
                    );
                }

                // Insertion confirmation
                if let Some(insertion_msg) = last_insertion {
                    content = content.child(
                        div()
                            .px_3()
                            .py_2()
                            .rounded_sm()
                            .bg(success_color.opacity(0.15))
                            .text_size(px(11.0))
                            .text_color(success_color)
                            .child(format!("\u{2713} {}", insertion_msg))
                    );
                }
            }
            AiVerb::Analyze => {
                // Analysis text
                if let Some(analysis) = &response_text {
                    content = content.child(
                        div()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .child(section_label("ANALYSIS", text_muted))
                            .child(
                                div()
                                    .w_full()
                                    .px_3()
                                    .py_2()
                                    .rounded_md()
                                    .bg(editor_bg.opacity(0.5))
                                    .text_size(px(12.0))
                                    .text_color(text_primary)
                                    .child(analysis.clone())
                            )
                    );

                    // Formula blurring warning
                    if analysis.contains("=SUM(") || analysis.contains("=AVERAGE(") || analysis.contains("=IF(") || analysis.contains("=VLOOKUP(") {
                        content = content.child(
                            div()
                                .px_3()
                                .py_1()
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .child(if cfg!(target_os = "macos") {
                                    "Formula suggestions are not available in Analyze mode. Use Insert Formula (\u{2318}+Shift+A) instead."
                                } else {
                                    "Formula suggestions are not available in Analyze mode. Use Insert Formula (Ctrl+Shift+A) instead."
                                })
                        );
                    }
                }
                // No formula, no insertion — read-only contract
            }
        }

        // Sent to AI panel (shared by both verbs)
        if let Some(sent) = sent_context {
            content = content.child(render_sent_panel(
                verb, panel_border, text_primary, text_muted, sent, sent_panel_expanded, cx,
            ));
        }
    }

    // Loading indicator
    if is_loading {
        content = content.child(
            div()
                .flex()
                .items_center()
                .justify_center()
                .py_4()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .child("Thinking...")
                )
        );
    }

    // Error message
    if let Some(err) = error {
        content = content.child(render_error_panel(
            panel_border, text_primary, error_color, err, has_raw_response, cx,
        ));
    }

    content
}

/// Render the footer with action buttons (branches on verb)
#[allow(clippy::too_many_arguments)]
fn render_footer(
    verb: AiVerb,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    success_color: Hsla,
    is_loading: bool,
    can_insert: bool,
    can_retry: bool,
    has_error: bool,
    has_response: bool,
    formula: Option<String>,
    response_text: Option<String>,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let mut left_buttons = div().flex().items_center().gap_2();
    let mut right_buttons = div().flex().items_center().gap_2();

    // Primary submit button with keyboard shortcut hint
    let shortcut = if cfg!(target_os = "macos") { "\u{2318}\u{21b5}" } else { "Ctrl+\u{21b5}" };
    let (btn_label, btn_loading_label) = match verb {
        AiVerb::InsertFormula => (format!("Ask AI  {}", shortcut), "Asking...".to_string()),
        AiVerb::Analyze => (format!("Analyze  {}", shortcut), "Analyzing...".to_string()),
    };
    let btn_text = if is_loading { btn_loading_label } else { btn_label };
    let submit_btn = Button::new("ask-ai-submit", btn_text)
        .disabled(is_loading)
        .primary(accent, text_primary);

    let submit_btn = if is_loading {
        submit_btn
    } else {
        submit_btn.on_click(cx.listener(|this, _, _, cx| {
            this.ask_ai_submit(cx);
        }))
    };
    left_buttons = left_buttons.child(submit_btn);

    // Retry button
    if can_retry {
        left_buttons = left_buttons.child(
            Button::new("ask-ai-retry", "Retry")
                .secondary(panel_border, text_primary)
                .on_click(cx.listener(|this, _, _, cx| {
                    this.ask_ai_retry(cx);
                }))
        );
    }

    // Refine button
    if can_retry && !has_error {
        left_buttons = left_buttons.child(
            Button::new("ask-ai-refine", "Refine\u{2026}")
                .secondary(panel_border, text_muted)
                .on_click(cx.listener(|this, _, _, cx| {
                    this.ask_ai_refine(cx);
                }))
        );
    }

    // Right side buttons (verb-specific)
    match verb {
        AiVerb::InsertFormula => {
            // Insert Formula button
            if can_insert {
                right_buttons = right_buttons.child(
                    Button::new("ask-ai-insert", "Insert Formula")
                        .primary(success_color, text_primary)
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.ask_ai_insert_formula(cx);
                        }))
                );
            }

            // Copy formula button
            if let Some(f) = formula {
                right_buttons = right_buttons.child(
                    Button::new("ask-ai-copy", "Copy")
                        .secondary(panel_border, text_primary)
                        .on_click(cx.listener(move |_, _, _, cx| {
                            cx.write_to_clipboard(ClipboardItem::new_string(f.clone()));
                        }))
                );
            }
        }
        AiVerb::Analyze => {
            // Copy analysis text button
            if has_response {
                if let Some(text) = response_text {
                    right_buttons = right_buttons.child(
                        Button::new("ask-ai-copy", "Copy")
                            .secondary(panel_border, text_primary)
                            .on_click(cx.listener(move |_, _, _, cx| {
                                cx.write_to_clipboard(ClipboardItem::new_string(text.clone()));
                            }))
                    );
                }
            }
            // No Insert button — read-only contract
        }
    }

    // Close button
    right_buttons = right_buttons.child(
        Button::new("ask-ai-cancel", "Close")
            .secondary(panel_border, text_primary)
            .on_click(cx.listener(|this, _, _, cx| {
                this.hide_ask_ai(cx);
            }))
    );

    div()
        .flex()
        .items_center()
        .justify_between()
        .child(left_buttons)
        .child(right_buttons)
}

/// Context selector (dropdown rendered separately as overlay)
fn render_context_selector(
    _panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    context_summary: String,
    _context_mode: AskAIContextMode,
    _context_selector_open: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(section_label("CONTEXT", text_muted))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .child(context_summary)
                )
                .child(
                    div()
                        .id("ask-ai-change-context")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(accent)
                        .hover(|s| s.bg(panel_border))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.ask_ai_toggle_context_selector(cx);
                        }))
                        .child("Change…")
                )
        )
}

/// Render the context menu as a standalone overlay (called from main render)
pub fn render_ask_ai_context_menu(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    // Only show if Ask AI dialog is open and context selector is open
    if !matches!(app.mode, crate::mode::Mode::AiDialog) || !app.ask_ai.context_selector_open {
        return None;
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let accent = app.token(TokenKey::Accent);
    let context_mode = app.ask_ai.context_mode;

    // Position: center of screen, offset to align with the context section
    // Dialog is 580px wide, centered. Context section is near top-left.
    // We'll position the menu at the dialog's left edge + some offset
    let window_width: f32 = app.window_size.width.into();
    let window_height: f32 = app.window_size.height.into();
    let dialog_width = 580.0;
    let dialog_left = (window_width - dialog_width) / 2.0;

    // Position dropdown near the "Change..." button
    // Header is ~50px, context section starts after that
    // "Change..." button is after the context summary text
    let menu_x = dialog_left + 16.0 + 100.0; // padding + rough offset to button
    let menu_y = (window_height / 2.0) - 200.0 + 85.0; // rough vertical position

    Some(
        div()
            .id("ask-ai-context-menu-backdrop")
            .absolute()
            .inset_0()
            // Click outside to close
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                this.ask_ai.context_selector_open = false;
                cx.notify();
                cx.stop_propagation();
            }))
            .child(
                div()
                    .id("ask-ai-context-menu")
                    .absolute()
                    .left(px(menu_x))
                    .top(px(menu_y))
                    .w(px(180.0))
                    .bg(panel_bg)
                    .border_1()
                    .border_color(panel_border)
                    .rounded_md()
                    .shadow_lg()
                    .flex()
                    .flex_col()
                    .py_1()
                    // Stop propagation for clicks inside menu
                    .on_mouse_down(MouseButton::Left, |_, _, cx| {
                        cx.stop_propagation();
                    })
                    .child(context_menu_option(
                        "ask-ai-ctx-selection", "Current selection",
                        context_mode == AskAIContextMode::CurrentSelection,
                        text_primary, accent, panel_border,
                        AskAIContextMode::CurrentSelection, cx,
                    ))
                    .child(context_menu_option(
                        "ask-ai-ctx-region", "Current region",
                        context_mode == AskAIContextMode::CurrentRegion,
                        text_primary, accent, panel_border,
                        AskAIContextMode::CurrentRegion, cx,
                    ))
                    .child(context_menu_option(
                        "ask-ai-ctx-used", "Entire used range",
                        context_mode == AskAIContextMode::EntireUsedRange,
                        text_primary, accent, panel_border,
                        AskAIContextMode::EntireUsedRange, cx,
                    ))
            )
    )
}

fn context_menu_option(
    id: &'static str,
    label: &'static str,
    is_selected: bool,
    text_primary: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    mode: AskAIContextMode,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .id(id)
        .px_3()
        .py(px(10.0))
        .cursor_pointer()
        .text_size(px(12.0))
        .text_color(if is_selected { accent } else { text_primary })
        .hover(|s| s.bg(panel_border))
        .on_click(cx.listener(move |this, _, _, cx| {
            this.ask_ai_set_context_mode(mode, cx);
        }))
        .child(if is_selected {
            format!("• {}", label)
        } else {
            format!("  {}", label)
        })
}

fn render_sent_panel(
    verb: AiVerb,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    sent: crate::app::AskAISentContext,
    expanded: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let (contract_label, write_scope) = match verb {
        AiVerb::InsertFormula => ("Single-cell write v1", "Active cell"),
        AiVerb::Analyze => ("Read-only v1", "None"),
    };

    div()
        .flex()
        .flex_col()
        .border_1()
        .border_color(panel_border.opacity(0.5))
        .rounded_sm()
        .child(
            div()
                .id("ask-ai-sent-toggle")
                .flex()
                .items_center()
                .gap_2()
                .px_3()
                .py_2()
                .cursor_pointer()
                .hover(|s| s.bg(panel_border.opacity(0.3)))
                .on_click(cx.listener(|this, _, _, cx| {
                    this.ask_ai_toggle_sent_panel(cx);
                }))
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(if expanded { "\u{25bc}" } else { "\u{25b6}" })
                )
                .child(
                    div()
                        .text_size(px(10.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_muted)
                        .child("SENT TO AI")
                )
        )
        .when(expanded, |el| {
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .border_t_1()
                    .border_color(panel_border.opacity(0.5))
                    .flex()
                    .flex_col()
                    .gap_1()
                    .text_size(px(11.0))
                    .child(sent_row("Provider", &sent.provider, text_muted, text_primary))
                    .child(sent_row("Model", &sent.model, text_muted, text_primary))
                    .child(sent_row("Range", &sent.range_display, text_muted, text_primary))
                    .child(sent_row("Cells", &format!("{} ({} \u{00d7} {})", sent.total_cells, sent.rows_sent, sent.cols_sent), text_muted, text_primary))
                    .child(sent_row("Headers", if sent.headers_included { "Yes" } else { "No" }, text_muted, text_primary))
                    .child(sent_row("Truncation", sent.truncation.label(), text_muted, text_primary))
                    .child(sent_row("Privacy mode", if sent.privacy_mode { "On (values only)" } else { "Off" }, text_muted, text_primary))
                    .child(sent_row("Contract", contract_label, text_muted, text_primary))
                    .child(sent_row("Write scope", write_scope, text_muted, text_primary))
            )
        })
}

fn sent_row(label: &str, value: &str, label_color: Hsla, value_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .gap_2()
        .child(
            div()
                .w(px(90.0))
                .text_color(label_color)
                .child(format!("{}:", label))
        )
        .child(
            div()
                .text_color(value_color)
                .child(value.to_string())
        )
}

fn render_error_panel(
    panel_border: Hsla,
    text_primary: Hsla,
    error_color: Hsla,
    error: String,
    has_details: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .px_3()
        .py_2()
        .rounded_sm()
        .bg(error_color.opacity(0.1))
        .child(
            div()
                .text_size(px(11.0))
                .text_color(error_color)
                .child(error)
        )
        .when(has_details, |el| {
            el.child(
                div()
                    .id("ask-ai-copy-details")
                    .px_2()
                    .py_1()
                    .rounded_sm()
                    .cursor_pointer()
                    .text_size(px(10.0))
                    .text_color(text_primary)
                    .bg(panel_border.opacity(0.5))
                    .hover(|s| s.bg(panel_border))
                    .on_click(cx.listener(|this, _, _, cx| {
                        this.ask_ai_copy_details(cx);
                    }))
                    .child("Copy details for bug report")
            )
        })
}

fn section_label(text: &'static str, color: Hsla) -> impl IntoElement {
    div()
        .text_size(px(10.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(color)
        .child(text)
}

/// Prompt suggestion chip for Analyze mode
fn prompt_chip(
    label: &'static str,
    fills_with: &'static str,
    accent: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .id(SharedString::from(format!("analyze-chip-{}", label.to_lowercase().replace(' ', "-"))))
        .px_2()
        .py_1()
        .rounded_md()
        .border_1()
        .border_color(accent.opacity(0.3))
        .cursor_pointer()
        .text_size(px(11.0))
        .text_color(text_primary)
        .hover(|s| s.bg(panel_border).border_color(accent.opacity(0.6)))
        .on_click(cx.listener(move |this, _, _, cx| {
            this.ask_ai.question = fills_with.to_string();
            cx.notify();
        }))
        .child(label)
}
