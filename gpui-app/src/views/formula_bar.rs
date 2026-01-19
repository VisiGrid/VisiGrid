use gpui::*;
use crate::app::{Spreadsheet, CELL_HEIGHT};

/// Render the formula bar (cell reference + formula/value input)
pub fn render_formula_bar(app: &Spreadsheet) -> impl IntoElement {
    let cell_ref = app.cell_ref();
    let editing = app.mode.is_editing();

    // Show edit value when editing, otherwise show raw value
    let display_value = if editing {
        app.edit_value.clone()
    } else {
        app.sheet.get_raw(app.selected.0, app.selected.1)
    };

    div()
        .flex_shrink_0()
        .h(px(CELL_HEIGHT))
        .bg(rgb(0x252526))
        .flex()
        .items_center()
        .border_b_1()
        .border_color(rgb(0x3c3c3c))
        // Cell reference label
        .child(
            div()
                .w(px(60.0))
                .h_full()
                .flex()
                .items_center()
                .justify_center()
                .border_r_1()
                .border_color(rgb(0x3c3c3c))
                .bg(rgb(0x1e1e1e))
                .text_color(rgb(0xcccccc))
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
                .border_color(rgb(0x3c3c3c))
                .text_color(rgb(0x808080))
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
                .text_color(if editing { rgb(0xffffff) } else { rgb(0xcccccc) })
                .bg(if editing { rgb(0x264f78) } else { rgb(0x1e1e1e) })
                .text_sm()
                .overflow_hidden()
                .child(display_value)
        )
}
