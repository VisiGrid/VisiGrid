use gpui::*;
use crate::app::Spreadsheet;

/// Render the Find dialog overlay
pub fn render_find_dialog(app: &Spreadsheet) -> impl IntoElement {
    let result_info = if app.find_results.is_empty() {
        if app.find_input.is_empty() {
            String::new()
        } else {
            "No matches".to_string()
        }
    } else {
        format!("{} of {}", app.find_index + 1, app.find_results.len())
    };

    div()
        .absolute()
        .top_2()
        .right_2()
        .w(px(300.0))
        .bg(rgb(0x252526))
        .border_1()
        .border_color(rgb(0x007acc))
        .rounded_md()
        .p_3()
        .flex()
        .flex_col()
        .gap_2()
        // Title
        .child(
            div()
                .text_color(rgb(0xffffff))
                .font_weight(FontWeight::MEDIUM)
                .text_sm()
                .child("Find")
        )
        // Input field with result count
        .child(
            div()
                .flex()
                .gap_2()
                .child(
                    div()
                        .flex_1()
                        .h(px(28.0))
                        .bg(rgb(0x1e1e1e))
                        .border_1()
                        .border_color(rgb(0x3c3c3c))
                        .rounded_sm()
                        .px_2()
                        .flex()
                        .items_center()
                        .text_color(rgb(0xffffff))
                        .text_sm()
                        .child(format!("{}|", app.find_input))
                )
                .child(
                    div()
                        .text_color(rgb(0x808080))
                        .text_sm()
                        .flex()
                        .items_center()
                        .child(result_info)
                )
        )
        // Instructions
        .child(
            div()
                .text_color(rgb(0x606060))
                .text_xs()
                .child("F3 next, Shift+F3 prev, Escape to close")
        )
}
