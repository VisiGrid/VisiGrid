use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;

/// Render the bottom status bar (Zed-inspired minimal design)
pub fn render_status_bar(app: &Spreadsheet, editing: bool, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
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

    // Get sheet information
    let sheet_names = app.workbook.sheet_names();
    let active_index = app.workbook.active_sheet_index();

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
            // Left side: sheet tabs + add button + mode
            div()
                .flex()
                .items_center()
                .gap_1()
                // Sheet tabs
                .children(
                    sheet_names.iter().enumerate().map(|(idx, name)| {
                        let is_active = idx == active_index;
                        sheet_tab(name, idx, is_active, cx)
                    })
                )
                // Add sheet button
                .child(
                    div()
                        .id("add-sheet-btn")
                        .px_1()
                        .py_px()
                        .cursor_pointer()
                        .text_color(rgb(0x858585))
                        .hover(|s| s.text_color(rgb(0xcccccc)).bg(rgb(0x3d3d3d)))
                        .rounded_sm()
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.add_sheet(cx);
                        }))
                        .child("+")
                )
                // Separator
                .child(
                    div()
                        .w(px(1.0))
                        .h(px(12.0))
                        .bg(rgb(0x3d3d3d))
                        .mx_2()
                )
                // Status message or mode
                .child(
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

/// Render a single sheet tab
fn sheet_tab(name: &str, index: usize, is_active: bool, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let name_owned = name.to_string();

    div()
        .id(ElementId::Name(format!("sheet-tab-{}", index).into()))
        .px_2()
        .py_px()
        .cursor_pointer()
        .rounded_sm()
        .when(is_active, |d| {
            d.bg(rgb(0x1e1e1e))
                .border_1()
                .border_color(rgb(0x3d3d3d))
                .text_color(rgb(0xcccccc))
        })
        .when(!is_active, |d| {
            d.text_color(rgb(0x858585))
                .hover(|s| s.bg(rgb(0x2d2d2d)).text_color(rgb(0xcccccc)))
        })
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.goto_sheet(index, cx);
        }))
        .child(name_owned)
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
            let display = app.sheet().get_display(row, col);
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
