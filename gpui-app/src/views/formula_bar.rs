use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CELL_HEIGHT};
use crate::theme::TokenKey;
use crate::formula_context::{tokenize_for_highlight, TokenType};

/// Render the formula bar (cell reference + formula/value input)
pub fn render_formula_bar(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let cell_ref = app.cell_ref();
    let editing = app.mode.is_editing();

    // Get the raw value (without cursor)
    let raw_value = if editing {
        app.edit_value.clone()
    } else {
        app.sheet().get_raw(app.selected.0, app.selected.1)
    };

    // Theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let app_bg = app.token(TokenKey::AppBg);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);

    let (input_bg, _input_text) = if editing {
        (selection_bg, text_primary)
    } else {
        (app_bg, text_primary)
    };

    // Build the formula display content
    let formula_content = build_formula_content(app, &raw_value, editing);

    div()
        .relative()
        .flex_shrink_0()
        .h(px(CELL_HEIGHT))
        .bg(panel_bg)
        .flex()
        .items_center()
        .border_b_1()
        .border_color(panel_border)
        // Cell reference label
        .child(
            div()
                .w(px(60.0))
                .h_full()
                .flex()
                .items_center()
                .justify_center()
                .border_r_1()
                .border_color(panel_border)
                .bg(app_bg)
                .text_color(text_primary)
                .text_sm()
                .font_weight(FontWeight::MEDIUM)
                .child(cell_ref)
        )
        // Function button (fx)
        .child(
            div()
                .w(px(30.0))
                .h_full()
                .flex()
                .items_center()
                .justify_center()
                .border_r_1()
                .border_color(panel_border)
                .text_color(text_muted)
                .text_sm()
                .child("fx")
        )
        // Formula/value input area - clickable to start editing
        .child(
            div()
                .id("formula-bar-input")
                .flex_1()
                .h_full()
                .flex()
                .items_center()
                .px_2()
                .bg(input_bg)
                .text_sm()
                .overflow_hidden()
                .cursor_text()
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    // Click to start editing if not already editing
                    if !this.mode.is_editing() {
                        this.start_edit(cx);
                    }
                }))
                .on_hover(cx.listener(move |this, hovering, _, cx| {
                    if *hovering {
                        // Extract first function from formula for hover docs
                        let raw_value = if this.mode.is_editing() {
                            &this.edit_value
                        } else {
                            &this.sheet().get_raw(this.selected.0, this.selected.1)
                        };
                        this.hover_function = extract_first_function(raw_value);
                    } else {
                        this.hover_function = None;
                    }
                    cx.notify();
                }))
                .child(formula_content)
        )
        // Note: Autocomplete, signature help, and error popups are rendered at the top level
        // in views/mod.rs to avoid being clipped by the formula bar's fixed height
}

/// Extract the first function from a formula for hover documentation
fn extract_first_function(formula: &str) -> Option<&'static crate::formula_context::FunctionInfo> {
    use crate::formula_context::{tokenize_for_highlight, TokenType, get_function};

    if !formula.starts_with('=') {
        return None;
    }

    let tokens = tokenize_for_highlight(formula);

    for (range, token_type) in tokens {
        if token_type == TokenType::Function {
            let func_name = &formula[range];
            return get_function(func_name);
        }
    }

    None
}

/// Build the formula content with syntax highlighting
fn build_formula_content(app: &Spreadsheet, raw_value: &str, editing: bool) -> AnyElement {
    let text_primary = app.token(TokenKey::TextPrimary);

    // Only highlight formulas (starting with '=')
    if !raw_value.starts_with('=') {
        // Plain text - just show with cursor if editing
        let display = if editing {
            let cursor_pos = app.edit_cursor;
            let chars: Vec<char> = raw_value.chars().collect();
            let before: String = chars.iter().take(cursor_pos).collect();
            let after: String = chars.iter().skip(cursor_pos).collect();
            format!("{}|{}", before, after)
        } else {
            raw_value.to_string()
        };
        return div().text_color(text_primary).child(display).into_any_element();
    }

    // Get syntax highlighting tokens
    let tokens = tokenize_for_highlight(raw_value);

    // Get theme colors for each token type
    let color_function = app.token(TokenKey::FormulaFunction);
    let color_cell_ref = app.token(TokenKey::FormulaCellRef);
    let color_number = app.token(TokenKey::FormulaNumber);
    let color_string = app.token(TokenKey::FormulaString);
    let color_boolean = app.token(TokenKey::FormulaBoolean);
    let color_operator = app.token(TokenKey::FormulaOperator);
    let color_parens = app.token(TokenKey::FormulaParens);
    let color_error = app.token(TokenKey::FormulaError);

    // Map TokenType to color
    let get_color = |token_type: &TokenType| -> Hsla {
        match token_type {
            TokenType::Function => color_function,
            TokenType::CellRef | TokenType::Range => color_cell_ref,
            TokenType::Number => color_number,
            TokenType::String => color_string,
            TokenType::Boolean => color_boolean,
            TokenType::Operator | TokenType::Comparison | TokenType::Comma | TokenType::Colon => color_operator,
            TokenType::Paren => color_parens,
            TokenType::Error => color_error,
            _ => text_primary,
        }
    };

    // Build text runs for styled text
    let mut runs: Vec<TextRun> = Vec::new();
    let mut last_end = 0;

    // Handle cursor insertion for editing mode
    let cursor_pos = if editing { app.edit_cursor } else { usize::MAX };

    // Build display text with cursor inserted
    let display_text: String = if editing {
        let chars: Vec<char> = raw_value.chars().collect();
        let before: String = chars.iter().take(cursor_pos).collect();
        let after: String = chars.iter().skip(cursor_pos).collect();
        format!("{}|{}", before, after)
    } else {
        raw_value.to_string()
    };

    // Process tokens and build runs
    for (range, token_type) in &tokens {
        // Fill gap before this token (if any) with default color
        if range.start > last_end {
            let gap_text = &raw_value[last_end..range.start];
            let gap_len = gap_text.len();
            // Adjust for cursor if it falls in this gap
            let adjusted_len = if editing && cursor_pos >= last_end && cursor_pos < range.start {
                gap_len + 1 // +1 for cursor character
            } else {
                gap_len
            };
            if adjusted_len > 0 {
                runs.push(TextRun {
                    len: adjusted_len,
                    font: Font::default(),
                    color: text_primary,
                    background_color: None,
                    underline: None,
                    strikethrough: None,
                });
            }
        }

        // Add this token's text run
        let token_text = &raw_value[range.clone()];
        let token_len = token_text.len();
        // Adjust for cursor if it falls in this token
        let adjusted_len = if editing && cursor_pos >= range.start && cursor_pos < range.end {
            token_len + 1 // +1 for cursor character
        } else if editing && cursor_pos == range.end {
            token_len // Cursor is right after this token
        } else {
            token_len
        };

        if adjusted_len > 0 {
            runs.push(TextRun {
                len: adjusted_len,
                font: Font::default(),
                color: get_color(token_type),
                background_color: None,
                underline: None,
                strikethrough: None,
            });
        }

        last_end = range.end;
    }

    // Handle remaining text after last token
    if last_end < raw_value.len() {
        let remaining = &raw_value[last_end..];
        let remaining_len = remaining.len();
        let adjusted_len = if editing && cursor_pos >= last_end {
            remaining_len + 1
        } else {
            remaining_len
        };
        if adjusted_len > 0 {
            runs.push(TextRun {
                len: adjusted_len,
                font: Font::default(),
                color: text_primary,
                background_color: None,
                underline: None,
                strikethrough: None,
            });
        }
    } else if editing && cursor_pos >= last_end {
        // Cursor is at the very end
        runs.push(TextRun {
            len: 1,
            font: Font::default(),
            color: text_primary,
            background_color: None,
            underline: None,
            strikethrough: None,
        });
    }

    // Ensure runs cover the entire display text
    let total_run_len: usize = runs.iter().map(|r| r.len).sum();
    if total_run_len < display_text.len() {
        runs.push(TextRun {
            len: display_text.len() - total_run_len,
            font: Font::default(),
            color: text_primary,
            background_color: None,
            underline: None,
            strikethrough: None,
        });
    }

    let shared_text: SharedString = display_text.into();
    StyledText::new(shared_text).with_runs(runs).into_any_element()
}

/// Render the autocomplete dropdown popup
pub fn render_autocomplete_popup(
    suggestions: &[&'static crate::formula_context::FunctionInfo],
    selected_index: usize,
    popup_x: f32,
    popup_y: f32,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Limit to 7 items as per spec
    let visible_items = suggestions.iter().take(7).enumerate();

    // Position below the active cell
    div()
        .absolute()
        .top(px(popup_y))
        .left(px(popup_x))
        .w(px(320.0))
        .max_h(px(220.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_md()
        .shadow_lg()
        .overflow_hidden()
        .py_1()
        .children(
            visible_items.map(|(idx, func)| {
                let is_selected = idx == selected_index;
                render_autocomplete_item(
                    func,
                    idx,
                    is_selected,
                    text_primary,
                    text_muted,
                    selection_bg,
                    cx,
                )
            })
        )
}

/// Render a single autocomplete item
fn render_autocomplete_item(
    func: &'static crate::formula_context::FunctionInfo,
    idx: usize,
    is_selected: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let bg_color = if is_selected {
        selection_bg
    } else {
        hsla(0.0, 0.0, 0.0, 0.0)
    };

    let func_name = func.name;
    let signature = func.signature;
    // Truncate signature if too long
    let display_sig = if signature.len() > 35 {
        format!("{}...", &signature[..32])
    } else {
        signature.to_string()
    };

    div()
        .id(ElementId::NamedInteger("autocomplete-item".into(), idx as u64))
        .flex()
        .items_center()
        .px_2()
        .py(px(4.0))
        .cursor_pointer()
        .bg(bg_color)
        .hover(|s| {
            if is_selected {
                s
            } else {
                s.bg(hsla(0.0, 0.0, 1.0, 0.05))
            }
        })
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.autocomplete_selected = idx;
            this.autocomplete_accept(cx);
        }))
        .child(
            div()
                .flex()
                .flex_col()
                .gap(px(2.0))
                .child(
                    div()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .font_weight(FontWeight::MEDIUM)
                        .child(func_name)
                )
                .child(
                    div()
                        .text_color(text_muted)
                        .text_size(px(11.0))
                        .child(display_sig)
                )
        )
}

/// Render the signature help tooltip
pub fn render_signature_help(
    sig_info: &crate::app::SignatureHelpInfo,
    popup_x: f32,
    popup_y: f32,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
) -> impl IntoElement {
    let func = sig_info.function;
    let current_arg = sig_info.current_arg;
    let params = func.parameters;

    // Build parameter list with current arg highlighted
    let param_elements: Vec<_> = params.iter().enumerate().map(|(idx, param)| {
        let is_current = idx == current_arg;
        let text_color = if is_current { accent } else { text_muted };
        let font_weight = if is_current { FontWeight::BOLD } else { FontWeight::NORMAL };

        let param_text = if param.optional {
            format!("[{}]", param.name)
        } else if param.repeatable {
            format!("{}...", param.name)
        } else {
            param.name.to_string()
        };

        div()
            .text_color(text_color)
            .font_weight(font_weight)
            .child(param_text)
    }).collect();

    // Get current parameter description if available
    let current_param_desc = params.get(current_arg).map(|p| p.description);

    // Position below the active cell
    div()
        .absolute()
        .top(px(popup_y))
        .left(px(popup_x))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_md()
        .shadow_lg()
        .px_3()
        .py_2()
        .max_w(px(400.0))
        .flex()
        .flex_col()
        .gap_1()
        // Function signature line
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(2.0))
                .child(
                    div()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .font_weight(FontWeight::MEDIUM)
                        .child(format!("{}(", func.name))
                )
                .children(
                    param_elements.into_iter().enumerate().map(|(idx, elem)| {
                        // Add comma separator between params (not before first)
                        if idx > 0 {
                            div()
                                .flex()
                                .items_center()
                                .child(
                                    div()
                                        .text_color(text_muted)
                                        .text_size(px(13.0))
                                        .child(", ")
                                )
                                .child(elem)
                        } else {
                            div().flex().child(elem)
                        }
                    })
                )
                .child(
                    div()
                        .text_color(text_primary)
                        .text_size(px(13.0))
                        .child(")")
                )
        )
        // Current parameter description
        .when_some(current_param_desc, |parent, desc| {
            parent.child(
                div()
                    .text_color(text_muted)
                    .text_size(px(11.0))
                    .child(desc)
            )
        })
}

/// Render the error banner
pub fn render_error_banner(
    error_info: &crate::app::FormulaErrorInfo,
    popup_x: f32,
    popup_y: f32,
    error_bg: Hsla,
    error_text: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    // Position below the active cell
    div()
        .absolute()
        .top(px(popup_y))
        .left(px(popup_x))
        .bg(error_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_md()
        .shadow_lg()
        .px_3()
        .py_2()
        .max_w(px(350.0))
        .flex()
        .items_center()
        .gap_2()
        // Error icon
        .child(
            div()
                .text_color(error_text)
                .text_size(px(14.0))
                .child("!")
        )
        // Error message
        .child(
            div()
                .text_color(error_text)
                .text_size(px(12.0))
                .child(error_info.message.clone())
        )
}

/// Render the hover documentation popup for a function
pub fn render_hover_docs(
    func: &'static crate::formula_context::FunctionInfo,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
) -> impl IntoElement {
    let params = func.parameters;

    // Build parameter descriptions
    let param_descs: Vec<_> = params.iter().map(|param| {
        let name = if param.optional {
            format!("[{}]", param.name)
        } else if param.repeatable {
            format!("{}...", param.name)
        } else {
            param.name.to_string()
        };
        (name, param.description)
    }).collect();

    // Position below the formula bar
    div()
        .absolute()
        .top(px(CELL_HEIGHT * 2.0))
        .left(px(90.0))  // After cell ref and fx button
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_md()
        .shadow_lg()
        .px_3()
        .py_2()
        .max_w(px(400.0))
        .flex()
        .flex_col()
        .gap_2()
        // Function name header
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_color(accent)
                        .text_size(px(14.0))
                        .font_weight(FontWeight::BOLD)
                        .child(func.name)
                )
        )
        // Signature
        .child(
            div()
                .text_color(text_primary)
                .text_size(px(12.0))
                .font_family("monospace")
                .child(func.signature)
        )
        // Description
        .child(
            div()
                .text_color(text_muted)
                .text_size(px(12.0))
                .child(func.description)
        )
        // Parameter list
        .when(!param_descs.is_empty(), |parent| {
            parent.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_1()
                    .mt_1()
                    .border_t_1()
                    .border_color(panel_border)
                    .pt_2()
                    .children(
                        param_descs.into_iter().map(|(name, desc)| {
                            div()
                                .flex()
                                .flex_col()
                                .gap(px(1.0))
                                .child(
                                    div()
                                        .text_color(text_primary)
                                        .text_size(px(11.0))
                                        .font_weight(FontWeight::MEDIUM)
                                        .child(name)
                                )
                                .child(
                                    div()
                                        .text_color(text_muted)
                                        .text_size(px(10.0))
                                        .pl_2()
                                        .child(desc)
                                )
                        })
                    )
            )
        })
}
