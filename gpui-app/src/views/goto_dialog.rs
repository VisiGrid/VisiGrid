use gpui::*;
use crate::app::Spreadsheet;

/// Render the Go To cell dialog overlay
pub fn render_goto_dialog(app: &Spreadsheet) -> impl IntoElement {
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(rgba(0x00000080))  // Semi-transparent backdrop
        .child(
            div()
                .w(px(300.0))
                .bg(rgb(0x252526))
                .border_1()
                .border_color(rgb(0x007acc))
                .rounded_md()
                .p_4()
                .flex()
                .flex_col()
                .gap_2()
                // Title
                .child(
                    div()
                        .text_color(rgb(0xffffff))
                        .font_weight(FontWeight::MEDIUM)
                        .child("Go To Cell")
                )
                // Input field
                .child(
                    div()
                        .w_full()
                        .h(px(32.0))
                        .bg(rgb(0x1e1e1e))
                        .border_1()
                        .border_color(rgb(0x3c3c3c))
                        .rounded_sm()
                        .px_2()
                        .flex()
                        .items_center()
                        .text_color(rgb(0xffffff))
                        .child(format!("{}|", app.goto_input))
                )
                // Help text
                .child(
                    div()
                        .text_color(rgb(0x808080))
                        .text_sm()
                        .child("Enter cell reference (e.g., A1, B25, AA100)")
                )
                // Instructions
                .child(
                    div()
                        .text_color(rgb(0x606060))
                        .text_xs()
                        .child("Enter to confirm, Escape to cancel")
                )
        )
}
