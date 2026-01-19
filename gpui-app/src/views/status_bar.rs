use gpui::*;
use crate::app::Spreadsheet;

/// Render the bottom status bar (Zed-inspired minimal design)
pub fn render_status_bar(app: &Spreadsheet, editing: bool) -> impl IntoElement {
    // Calculate selection stats if multiple cells selected
    let selection_stats = calculate_selection_stats(app);

    // Mode indicator
    let mode_text = if editing {
        "Edit"
    } else if app.status_message.is_some() {
        ""
    } else {
        "Ready"
    };

    div()
        .flex_shrink_0()
        .h(px(22.0))
        .bg(rgb(0x252526))
        .border_t_1()
        .border_color(rgb(0x3d3d3d))
        .flex()
        .items_center()
        .justify_between()
        .px_2()
        .text_color(rgb(0x858585))
        .text_xs()
        .child(
            // Left side: sheet tabs placeholder + mode
            div()
                .flex()
                .items_center()
                .gap_3()
                .child(
                    // Sheet tab (placeholder for multi-sheet)
                    div()
                        .px_2()
                        .py_px()
                        .bg(rgb(0x1e1e1e))
                        .border_1()
                        .border_color(rgb(0x3d3d3d))
                        .rounded_sm()
                        .text_color(rgb(0xcccccc))
                        .child("Sheet1")
                )
                .child(
                    // Status message or mode
                    div()
                        .text_color(rgb(0x858585))
                        .child(
                            if let Some(msg) = &app.status_message {
                                msg.clone()
                            } else {
                                mode_text.to_string()
                            }
                        )
                )
        )
        .child(
            // Right side: selection stats
            div()
                .flex()
                .items_center()
                .gap_4()
                .children(selection_stats)
        )
}

/// Calculate statistics for the current selection
fn calculate_selection_stats(app: &Spreadsheet) -> Vec<impl IntoElement> {
    let ((min_row, min_col), (max_row, max_col)) = app.selection_range();

    // Only show stats if more than one cell is selected
    let is_multi_select = min_row != max_row || min_col != max_col;
    if !is_multi_select {
        return vec![];
    }

    // Collect numeric values from selection
    let mut values: Vec<f64> = Vec::new();
    let mut count = 0usize;

    for row in min_row..=max_row {
        for col in min_col..=max_col {
            count += 1;
            let display = app.sheet.get_display(row, col);
            if let Ok(num) = display.parse::<f64>() {
                values.push(num);
            }
        }
    }

    if values.is_empty() {
        // No numeric values, just show count
        return vec![
            stat_item("Count", &count.to_string()),
        ];
    }

    let sum: f64 = values.iter().sum();
    let avg = sum / values.len() as f64;
    let min = values.iter().cloned().fold(f64::INFINITY, f64::min);
    let max = values.iter().cloned().fold(f64::NEG_INFINITY, f64::max);

    vec![
        stat_item("Sum", &format_number(sum)),
        stat_item("Average", &format_number(avg)),
        stat_item("Min", &format_number(min)),
        stat_item("Max", &format_number(max)),
        stat_item("Count", &count.to_string()),
    ]
}

fn stat_item(label: &str, value: &str) -> Div {
    div()
        .flex()
        .items_center()
        .gap_1()
        .child(
            div()
                .text_color(rgb(0x6a6a6a))
                .child(format!("{}:", label))
        )
        .child(
            div()
                .text_color(rgb(0xcccccc))
                .child(value.to_string())
        )
}

fn format_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e10 {
        format!("{:.0}", n)
    } else if n.abs() < 0.0001 || n.abs() >= 1e10 {
        format!("{:.2e}", n)
    } else {
        format!("{:.2}", n)
    }
}
