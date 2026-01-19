use gpui::*;
use crate::app::Spreadsheet;

/// Render the bottom status bar
pub fn render_status_bar(app: &Spreadsheet, editing: bool) -> impl IntoElement {
    let cell_ref = app.cell_ref();
    let raw_value = app.sheet.get_raw(app.selected.0, app.selected.1);

    let bg_color = if editing {
        rgb(0x68217a)  // Purple when editing
    } else {
        rgb(0x007acc)  // Blue normally
    };

    let status_text = if editing {
        format!("EDITING {} | Enter to confirm, Escape to cancel", cell_ref)
    } else if let Some(msg) = &app.status_message {
        msg.clone()
    } else {
        format!(
            "Cell: {} | Value: {} | Enter to edit, Arrows to move",
            cell_ref,
            raw_value
        )
    };

    div()
        .flex_shrink_0()
        .h(px(24.0))
        .bg(bg_color)
        .flex()
        .items_center()
        .px_2()
        .text_color(rgb(0xffffff))
        .text_sm()
        .child(status_text)
}
