use gpui::*;
use crate::app::{Spreadsheet, CELL_HEIGHT};
use crate::theme::TokenKey;

/// Render the formula bar (cell reference + formula/value input)
pub fn render_formula_bar(app: &Spreadsheet) -> impl IntoElement {
    let cell_ref = app.cell_ref();
    let editing = app.mode.is_editing();

    // Show edit value when editing, otherwise show raw value
    let display_value = if editing {
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
    let text_inverse = app.token(TokenKey::TextInverse);

    let (input_bg, input_text) = if editing {
        (selection_bg, text_primary)
    } else {
        (app_bg, text_primary)
    };

    div()
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
        // Formula/value input area
        .child(
            div()
                .flex_1()
                .h_full()
                .flex()
                .items_center()
                .px_2()
                .text_color(input_text)
                .bg(input_bg)
                .text_sm()
                .overflow_hidden()
                .child(display_value)
        )
}
