use std::time::Duration;
use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::theme::TokenKey;

/// Render the rename symbol dialog (Ctrl+Shift+R)
pub(crate) fn render_rename_symbol_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let selection_bg = app.token(TokenKey::SelectionBg);

    let affected_count = app.rename_affected_cells.len();
    let select_all = app.rename_select_all;
    let has_error = app.rename_validation_error.is_some();

    // Build affected cells preview
    let cells_preview: Vec<String> = app.rename_affected_cells
        .iter()
        .take(8)
        .map(|(row, col)| {
            let col_letter = col_to_letter(*col);
            format!("{}{}", col_letter, row + 1)
        })
        .collect();

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child(format!("Rename '{}'", app.rename_original_name))
                )
                // Input field
                .child(
                    div()
                        .px_2()
                        .py_1()
                        .bg(hsla(0.0, 0.0, 0.0, 0.2))
                        .rounded_sm()
                        .border_1()
                        .border_color(if has_error { error_color } else { panel_border })
                        .text_color(text_primary)
                        .child(
                            // Show text with selection highlight if select_all is active
                            div()
                                .when(select_all, |d| d.bg(selection_bg).rounded_sm().px_1())
                                .child(app.rename_new_name.clone())
                        )
                )
                // Validation error (if any)
                .when_some(app.rename_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Affected cells count
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child(format!(
                            "{} formula{} will be updated",
                            affected_count,
                            if affected_count == 1 { "" } else { "s" }
                        ))
                )
                // Preview of affected cells (show list of cell refs)
                .when(affected_count > 0, |d| {
                    let preview = cells_preview.join(", ");
                    let more = if affected_count > 8 {
                        format!(" ...and {} more", affected_count - 8)
                    } else {
                        String::new()
                    };
                    d.child(
                        div()
                            .text_color(text_muted)
                            .text_xs()
                            .child(format!("{}{}", preview, more))
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Enter to confirm • Escape to cancel")
                )
        )
}

/// Render the edit description dialog
pub(crate) fn render_edit_description_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child(format!("Edit description for '{}'", app.edit_description_name))
                )
                // Input field
                .child(
                    div()
                        .px_2()
                        .py_2()
                        .bg(hsla(0.0, 0.0, 0.0, 0.2))
                        .rounded_sm()
                        .border_1()
                        .border_color(panel_border)
                        .text_color(if app.edit_description_value.is_empty() { text_muted } else { text_primary })
                        .min_h(px(60.0))
                        .child(if app.edit_description_value.is_empty() {
                            "Enter a description...".to_string()
                        } else {
                            app.edit_description_value.clone()
                        })
                )
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Enter to save • Escape to cancel")
                )
        )
}

/// Convert column index to letter(s) (0 = A, 25 = Z, 26 = AA, etc.)
pub(crate) fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

/// Render the create named range dialog (Ctrl+Shift+N)
pub(crate) fn render_create_named_range_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let has_error = app.create_name_validation_error.is_some();
    let name_focused = app.create_name_focus == CreateNameFocus::Name;
    let desc_focused = app.create_name_focus == CreateNameFocus::Description;

    // Centered dialog overlay - blocks clicks from reaching grid below
    div()
        .id("create-named-range-overlay")
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(|_this, _event, _window, _cx| {
            // Consume click to prevent it reaching grid below
        }))
        .child(
            div()
                .w(px(400.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("Create Named Range")
                )
                // Target (read-only display)
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Target:")
                        )
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_sm()
                                .child(app.create_name_target.clone())
                        )
                )
                // Name input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Name:")
                        )
                        .child(
                            div()
                                .id("create-name-input")
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if name_focused && has_error {
                                    error_color
                                } else if name_focused {
                                    accent
                                } else {
                                    panel_border
                                })
                                .text_color(if app.create_name_name.is_empty() && !name_focused { text_muted } else { text_primary })
                                .cursor_text()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.create_name_focus = CreateNameFocus::Name;
                                    cx.notify();
                                }))
                                .flex()
                                .items_center()
                                .child(if app.create_name_name.is_empty() && !name_focused {
                                    "(required)".to_string()
                                } else {
                                    app.create_name_name.clone()
                                })
                                .when(name_focused, |d| {
                                    d.child(
                                        div()
                                            .w(px(1.0))
                                            .h(px(14.0))
                                            .bg(text_primary)
                                            .with_animation(
                                                "name-cursor-blink",
                                                Animation::new(Duration::from_millis(530))
                                                    .repeat()
                                                    .with_easing(pulsating_between(0.0, 1.0)),
                                                |this, delta| {
                                                    let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                                    this.opacity(opacity)
                                                },
                                            )
                                    )
                                })
                        )
                )
                // Description input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Description:")
                        )
                        .child(
                            div()
                                .id("create-desc-input")
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if desc_focused { accent } else { panel_border })
                                .text_color(if app.create_name_description.is_empty() && !desc_focused { text_muted } else { text_primary })
                                .cursor_text()
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.create_name_focus = CreateNameFocus::Description;
                                    cx.notify();
                                }))
                                .flex()
                                .items_center()
                                .child(if app.create_name_description.is_empty() && !desc_focused {
                                    "(optional)".to_string()
                                } else {
                                    app.create_name_description.clone()
                                })
                                .when(desc_focused, |d| {
                                    d.child(
                                        div()
                                            .w(px(1.0))
                                            .h(px(14.0))
                                            .bg(text_primary)
                                            .with_animation(
                                                "desc-cursor-blink",
                                                Animation::new(Duration::from_millis(530))
                                                    .repeat()
                                                    .with_easing(pulsating_between(0.0, 1.0)),
                                                |this, delta| {
                                                    let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                                    this.opacity(opacity)
                                                },
                                            )
                                    )
                                })
                        )
                )
                // Validation error (if any)
                .when_some(app.create_name_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Tab to switch fields • Enter to confirm • Escape to cancel")
                )
        )
}

/// Render the extract named range dialog
pub(crate) fn render_extract_named_range_dialog(app: &Spreadsheet) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);
    let accent = app.token(TokenKey::Accent);

    let has_error = app.extract_validation_error.is_some();
    let name_focused = app.extract_focus == CreateNameFocus::Name;
    let desc_focused = app.extract_focus == CreateNameFocus::Description;

    // Format occurrence message
    let occurrence_msg = if app.extract_affected_cells.len() == 1 {
        format!("Will replace {} occurrence in 1 formula", app.extract_occurrence_count)
    } else {
        format!("Will replace {} occurrences in {} formulas", app.extract_occurrence_count, app.extract_affected_cells.len())
    };

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(420.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                // Header
                .child(
                    div()
                        .text_color(text_primary)
                        .text_sm()
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("Extract to Named Range")
                )
                // Range preview (read-only)
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Range:")
                        )
                        .child(
                            div()
                                .text_color(accent)
                                .text_sm()
                                .font_weight(FontWeight::MEDIUM)
                                .child(app.extract_range_literal.clone())
                        )
                )
                // Name input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Name:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(if app.extract_select_all && name_focused {
                                    accent.opacity(0.3)  // Selection highlight
                                } else {
                                    hsla(0.0, 0.0, 0.0, 0.2)
                                })
                                .rounded_sm()
                                .border_1()
                                .border_color(if name_focused && has_error {
                                    error_color
                                } else if name_focused {
                                    accent
                                } else {
                                    panel_border
                                })
                                .text_color(text_primary)
                                .child(if app.extract_name.is_empty() && !name_focused {
                                    "(required)".to_string()
                                } else {
                                    app.extract_name.clone()
                                })
                        )
                )
                // Description input
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .items_center()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .w(px(70.0))
                                .child("Description:")
                        )
                        .child(
                            div()
                                .flex_1()
                                .px_2()
                                .py_1()
                                .bg(hsla(0.0, 0.0, 0.0, 0.2))
                                .rounded_sm()
                                .border_1()
                                .border_color(if desc_focused { accent } else { panel_border })
                                .text_color(if app.extract_description.is_empty() { text_muted } else { text_primary })
                                .child(if app.extract_description.is_empty() {
                                    "(optional)".to_string()
                                } else {
                                    app.extract_description.clone()
                                })
                        )
                )
                // Occurrence count
                .child(
                    div()
                        .px_2()
                        .py_2()
                        .bg(hsla(0.0, 0.0, 0.0, 0.15))
                        .rounded_sm()
                        .text_color(text_muted)
                        .text_xs()
                        .child(occurrence_msg)
                )
                // Validation error (if any)
                .when_some(app.extract_validation_error.clone(), |d, err| {
                    d.child(
                        div()
                            .text_color(error_color)
                            .text_xs()
                            .child(err)
                    )
                })
                // Instructions
                .child(
                    div()
                        .text_color(text_muted)
                        .text_xs()
                        .child("Tab to switch fields • Enter to extract • Escape to cancel")
                )
        )
}
