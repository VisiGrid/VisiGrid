//! Data Validation dialog (Phase 4)
//!
//! Modal dialog for creating/editing data validation rules.
//! MVP supports: Any Value, List, Whole Number, Decimal

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, ValidationTypeOption, NumericOperatorOption, ValidationDialogFocus};
use crate::theme::TokenKey;

/// Render the data validation dialog overlay
pub fn render_validation_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);
    let error_color: Hsla = rgb(0xE53935).into();

    let state = &app.validation_dialog;
    let has_error = state.error.is_some();

    // Calculate target range display
    let range_display = state.target_range.as_ref().map(|r| {
        if r.start_row == r.end_row && r.start_col == r.end_col {
            app.cell_ref_at(r.start_row, r.start_col)
        } else {
            format!("{}:{}",
                app.cell_ref_at(r.start_row, r.start_col),
                app.cell_ref_at(r.end_row, r.end_col))
        }
    }).unwrap_or_else(|| "Selection".to_string());

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        // Click outside to close
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_validation_dialog(cx);
        }))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_lg()
                .overflow_hidden()
                // Stop click propagation
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                // Keyboard handling
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    let key = &event.keystroke.key;
                    match key.as_str() {
                        "escape" => {
                            // Close dropdowns first, then close dialog
                            if this.validation_dialog.type_dropdown_open || this.validation_dialog.operator_dropdown_open {
                                this.validation_dialog.type_dropdown_open = false;
                                this.validation_dialog.operator_dropdown_open = false;
                                cx.notify();
                            } else {
                                this.hide_validation_dialog(cx);
                            }
                        }
                        "enter" => {
                            if !this.validation_dialog.type_dropdown_open && !this.validation_dialog.operator_dropdown_open {
                                this.apply_validation_dialog(cx);
                            }
                        }
                        "tab" => {
                            this.validation_dialog_tab(event.keystroke.modifiers.shift, cx);
                        }
                        "backspace" => {
                            this.validation_dialog_backspace(cx);
                        }
                        _ => {
                            if let Some(c) = event.keystroke.key_char.as_ref().and_then(|s| s.chars().next()) {
                                if !event.keystroke.modifiers.control && !event.keystroke.modifiers.alt {
                                    this.validation_dialog_type_char(c, cx);
                                }
                            }
                        }
                    }
                    cx.stop_propagation();
                }))
                // Header
                .child(
                    div()
                        .h(px(48.0))
                        .px_4()
                        .flex()
                        .items_center()
                        .justify_between()
                        .border_b_1()
                        .border_color(panel_border)
                        .child(
                            div()
                                .text_base()
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Data Validation")
                        )
                        .child(
                            div()
                                .text_sm()
                                .text_color(text_muted)
                                .child(SharedString::from(range_display))
                        )
                )
                // Content
                .child(
                    div()
                        .p_4()
                        .flex()
                        .flex_col()
                        .gap_4()
                        // Validation Type selector
                        .child(render_type_selector(app, text_primary, text_muted, accent, editor_bg, panel_border, cx))
                        // Type-specific fields
                        .child(render_type_fields(app, text_primary, text_muted, accent, editor_bg, panel_border, cx))
                        // Ignore blank checkbox
                        .child(render_ignore_blank_checkbox(app, text_primary, accent, cx))
                        // Error message
                        .when(has_error, |el| {
                            el.child(
                                div()
                                    .text_sm()
                                    .text_color(error_color)
                                    .child(SharedString::from(state.error.clone().unwrap_or_default()))
                            )
                        })
                )
                // Footer with buttons
                .child(
                    div()
                        .px_4()
                        .py_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            // Clear button (only if has existing validation)
                            div()
                                .when(state.has_existing_validation, |el| {
                                    el.child(render_button("Clear All", false, text_primary, panel_border, cx, |this, cx| {
                                        this.clear_validation_dialog(cx);
                                    }))
                                })
                        )
                        .child(
                            div()
                                .flex()
                                .gap_2()
                                .child(render_button("Cancel", false, text_primary, panel_border, cx, |this, cx| {
                                    this.hide_validation_dialog(cx);
                                }))
                                .child(render_button("OK", true, text_primary, accent, cx, |this, cx| {
                                    this.apply_validation_dialog(cx);
                                }))
                        )
                )
        )
}

fn render_type_selector(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let state = &app.validation_dialog;
    let is_focused = state.focus == ValidationDialogFocus::TypeDropdown;
    let is_open = state.type_dropdown_open;

    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(
            div()
                .text_sm()
                .text_color(text_muted)
                .child("Allow:")
        )
        .child(
            div()
                .relative()
                .child(
                    div()
                        .id("validation-type-selector")
                        .h(px(32.0))
                        .px_2()
                        .bg(editor_bg)
                        .border_1()
                        .border_color(if is_focused { accent } else { panel_border })
                        .rounded_sm()
                        .flex()
                        .items_center()
                        .justify_between()
                        .cursor_pointer()
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.validation_dialog.type_dropdown_open = !this.validation_dialog.type_dropdown_open;
                            this.validation_dialog.operator_dropdown_open = false;
                            this.validation_dialog.focus = ValidationDialogFocus::TypeDropdown;
                            cx.notify();
                        }))
                        .child(
                            div()
                                .text_sm()
                                .text_color(text_primary)
                                .child(state.validation_type.label())
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .child("▼")
                        )
                )
                .when(is_open, |el| {
                    el.child(render_type_dropdown(app, text_primary, editor_bg, accent, panel_border, cx))
                })
        )
}

fn render_type_dropdown(
    app: &Spreadsheet,
    text_primary: Hsla,
    editor_bg: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let current = app.validation_dialog.validation_type;

    div()
        .absolute()
        .top(px(34.0))
        .left_0()
        .right_0()
        .bg(editor_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_md()
        .children(ValidationTypeOption::ALL.iter().map(|&opt| {
            let is_selected = opt == current;
            div()
                .id(ElementId::Name(format!("vtype-{:?}", opt).into()))
                .px_2()
                .py_1()
                .text_sm()
                .text_color(text_primary)
                .cursor_pointer()
                .when(is_selected, |el| el.bg(accent.opacity(0.2)))
                .hover(|s| s.bg(accent.opacity(0.1)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.validation_dialog.validation_type = opt;
                    this.validation_dialog.type_dropdown_open = false;
                    this.validation_dialog.error = None;
                    // Reset focus to appropriate field
                    this.validation_dialog.focus = match opt {
                        ValidationTypeOption::AnyValue => ValidationDialogFocus::TypeDropdown,
                        ValidationTypeOption::List => ValidationDialogFocus::Source,
                        _ => ValidationDialogFocus::OperatorDropdown,
                    };
                    cx.notify();
                }))
                .child(opt.label())
        }))
}

fn render_type_fields(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let state = &app.validation_dialog;

    div()
        .flex()
        .flex_col()
        .gap_3()
        .when(state.validation_type == ValidationTypeOption::List, |el| {
            el.child(render_list_fields(app, text_primary, text_muted, accent, editor_bg, panel_border, cx))
        })
        .when(state.validation_type == ValidationTypeOption::WholeNumber || state.validation_type == ValidationTypeOption::Decimal, |el| {
            el.child(render_numeric_fields(app, text_primary, text_muted, accent, editor_bg, panel_border, cx))
        })
}

fn render_list_fields(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let state = &app.validation_dialog;
    let is_focused = state.focus == ValidationDialogFocus::Source;

    div()
        .flex()
        .flex_col()
        .gap_3()
        // Source field
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_sm()
                        .text_color(text_muted)
                        .child("Source:")
                )
                .child(
                    div()
                        .id("validation-source-field")
                        .h(px(32.0))
                        .px_2()
                        .bg(editor_bg)
                        .border_1()
                        .border_color(if is_focused { accent } else { panel_border })
                        .rounded_sm()
                        .flex()
                        .items_center()
                        .cursor_text()
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.validation_dialog.focus = ValidationDialogFocus::Source;
                            this.validation_dialog.type_dropdown_open = false;
                            cx.notify();
                        }))
                        .child(
                            div()
                                .text_sm()
                                .text_color(if state.list_source.is_empty() { text_muted } else { text_primary })
                                .child(if state.list_source.is_empty() {
                                    "A1:A10, Yes,No,Maybe, or NamedRange".to_string()
                                } else if is_focused {
                                    format!("{}|", state.list_source)
                                } else {
                                    state.list_source.clone()
                                })
                        )
                )
        )
        // Show dropdown checkbox
        .child(render_show_dropdown_checkbox(app, text_primary, accent, cx))
}

fn render_numeric_fields(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    editor_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let state = &app.validation_dialog;
    let needs_two = state.numeric_operator.needs_two_values();
    let is_op_focused = state.focus == ValidationDialogFocus::OperatorDropdown;
    let is_v1_focused = state.focus == ValidationDialogFocus::Value1;
    let is_v2_focused = state.focus == ValidationDialogFocus::Value2;
    let is_op_open = state.operator_dropdown_open;

    div()
        .flex()
        .flex_col()
        .gap_3()
        // Operator selector
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_sm()
                        .text_color(text_muted)
                        .child("Data:")
                )
                .child(
                    div()
                        .relative()
                        .child(
                            div()
                                .id("validation-operator-selector")
                                .h(px(32.0))
                                .px_2()
                                .bg(editor_bg)
                                .border_1()
                                .border_color(if is_op_focused { accent } else { panel_border })
                                .rounded_sm()
                                .flex()
                                .items_center()
                                .justify_between()
                                .cursor_pointer()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.validation_dialog.operator_dropdown_open = !this.validation_dialog.operator_dropdown_open;
                                    this.validation_dialog.type_dropdown_open = false;
                                    this.validation_dialog.focus = ValidationDialogFocus::OperatorDropdown;
                                    cx.notify();
                                }))
                                .child(
                                    div()
                                        .text_sm()
                                        .text_color(text_primary)
                                        .child(state.numeric_operator.label())
                                )
                                .child(
                                    div()
                                        .text_color(text_muted)
                                        .child("▼")
                                )
                        )
                        .when(is_op_open, |el| {
                            el.child(render_operator_dropdown(app, text_primary, editor_bg, accent, panel_border, cx))
                        })
                )
        )
        // Value fields
        .child(
            div()
                .flex()
                .gap_2()
                // Value1 (Minimum for between)
                .child(
                    div()
                        .flex_1()
                        .flex()
                        .flex_col()
                        .gap_1()
                        .child(
                            div()
                                .text_sm()
                                .text_color(text_muted)
                                .child(if needs_two { "Minimum:" } else { "Value:" })
                        )
                        .child(
                            div()
                                .id("validation-value1-field")
                                .h(px(32.0))
                                .px_2()
                                .bg(editor_bg)
                                .border_1()
                                .border_color(if is_v1_focused { accent } else { panel_border })
                                .rounded_sm()
                                .flex()
                                .items_center()
                                .cursor_text()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.validation_dialog.focus = ValidationDialogFocus::Value1;
                                    this.validation_dialog.type_dropdown_open = false;
                                    this.validation_dialog.operator_dropdown_open = false;
                                    cx.notify();
                                }))
                                .child(
                                    div()
                                        .text_sm()
                                        .text_color(if state.value1.is_empty() { text_muted } else { text_primary })
                                        .child(if state.value1.is_empty() {
                                            "Number or cell ref".to_string()
                                        } else if is_v1_focused {
                                            format!("{}|", state.value1)
                                        } else {
                                            state.value1.clone()
                                        })
                                )
                        )
                )
                // Value2 (Maximum for between)
                .when(needs_two, |el| {
                    el.child(
                        div()
                            .flex_1()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .child(
                                div()
                                    .text_sm()
                                    .text_color(text_muted)
                                    .child("Maximum:")
                            )
                            .child(
                                div()
                                    .id("validation-value2-field")
                                    .h(px(32.0))
                                    .px_2()
                                    .bg(editor_bg)
                                    .border_1()
                                    .border_color(if is_v2_focused { accent } else { panel_border })
                                    .rounded_sm()
                                    .flex()
                                    .items_center()
                                    .cursor_text()
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.validation_dialog.focus = ValidationDialogFocus::Value2;
                                        this.validation_dialog.type_dropdown_open = false;
                                        this.validation_dialog.operator_dropdown_open = false;
                                        cx.notify();
                                    }))
                                    .child(
                                        div()
                                            .text_sm()
                                            .text_color(if state.value2.is_empty() { text_muted } else { text_primary })
                                            .child(if state.value2.is_empty() {
                                                "Number or cell ref".to_string()
                                            } else if is_v2_focused {
                                                format!("{}|", state.value2)
                                            } else {
                                                state.value2.clone()
                                            })
                                    )
                            )
                    )
                })
        )
}

fn render_operator_dropdown(
    app: &Spreadsheet,
    text_primary: Hsla,
    editor_bg: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let current = app.validation_dialog.numeric_operator;

    div()
        .absolute()
        .top(px(34.0))
        .left_0()
        .right_0()
        .bg(editor_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_md()
        .max_h(px(200.0))
        .overflow_hidden()
        .children(NumericOperatorOption::ALL.iter().map(|&opt| {
            let is_selected = opt == current;
            div()
                .id(ElementId::Name(format!("vop-{:?}", opt).into()))
                .px_2()
                .py_1()
                .text_sm()
                .text_color(text_primary)
                .cursor_pointer()
                .when(is_selected, |el| el.bg(accent.opacity(0.2)))
                .hover(|s| s.bg(accent.opacity(0.1)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.validation_dialog.numeric_operator = opt;
                    this.validation_dialog.operator_dropdown_open = false;
                    this.validation_dialog.error = None;
                    // Focus appropriate value field
                    this.validation_dialog.focus = ValidationDialogFocus::Value1;
                    cx.notify();
                }))
                .child(opt.label())
        }))
}

fn render_show_dropdown_checkbox(
    app: &Spreadsheet,
    text_primary: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_checked = app.validation_dialog.show_dropdown;

    div()
        .id("validation-show-dropdown")
        .flex()
        .items_center()
        .gap_2()
        .cursor_pointer()
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.validation_dialog.show_dropdown = !this.validation_dialog.show_dropdown;
            cx.notify();
        }))
        .child(
            div()
                .w(px(16.0))
                .h(px(16.0))
                .rounded_sm()
                .border_1()
                .border_color(if is_checked { accent } else { text_primary.opacity(0.3) })
                .bg(if is_checked { accent } else { gpui::transparent_black() })
                .flex()
                .items_center()
                .justify_center()
                .when(is_checked, |el| {
                    el.child(
                        div()
                            .text_xs()
                            .text_color(gpui::white())
                            .child("✓")
                    )
                })
        )
        .child(
            div()
                .text_sm()
                .text_color(text_primary)
                .child("In-cell dropdown")
        )
}

fn render_ignore_blank_checkbox(
    app: &Spreadsheet,
    text_primary: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_checked = app.validation_dialog.ignore_blank;

    div()
        .id("validation-ignore-blank")
        .flex()
        .items_center()
        .gap_2()
        .cursor_pointer()
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.validation_dialog.ignore_blank = !this.validation_dialog.ignore_blank;
            cx.notify();
        }))
        .child(
            div()
                .w(px(16.0))
                .h(px(16.0))
                .rounded_sm()
                .border_1()
                .border_color(if is_checked { accent } else { text_primary.opacity(0.3) })
                .bg(if is_checked { accent } else { gpui::transparent_black() })
                .flex()
                .items_center()
                .justify_center()
                .when(is_checked, |el| {
                    el.child(
                        div()
                            .text_xs()
                            .text_color(gpui::white())
                            .child("✓")
                    )
                })
        )
        .child(
            div()
                .text_sm()
                .text_color(text_primary)
                .child("Ignore blank")
        )
}

fn render_button(
    label: &'static str,
    primary: bool,
    text_color: Hsla,
    border_or_accent: Hsla,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    div()
        .id(ElementId::Name(format!("vd-btn-{}", label).into()))
        .px_4()
        .py(px(6.0))
        .rounded_sm()
        .text_sm()
        .cursor_pointer()
        .when(primary, |el| {
            el.bg(border_or_accent)
                .text_color(gpui::white())
                .hover(|s| s.bg(border_or_accent.opacity(0.8)))
        })
        .when(!primary, |el| {
            el.border_1()
                .border_color(border_or_accent)
                .text_color(text_color)
                .hover(|s| s.bg(border_or_accent.opacity(0.1)))
        })
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            action(this, cx);
        }))
        .child(label)
}
