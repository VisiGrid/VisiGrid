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
    let renaming_sheet = app.renaming_sheet;
    let rename_input = app.sheet_rename_input.clone();
    let context_menu_sheet = app.sheet_context_menu;

    div()
        .relative()
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
                        let is_renaming = renaming_sheet == Some(idx);
                        let name_str = name.to_string();
                        let input_str = rename_input.clone();
                        sheet_tab_wrapper(name_str, input_str, idx, is_active, is_renaming, cx)
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
        // Context menu overlay
        .when(context_menu_sheet.is_some(), |d| {
            d.child(render_sheet_context_menu(context_menu_sheet.unwrap(), cx))
        })
}

/// Wrapper to return consistent type for sheet tabs
fn sheet_tab_wrapper(
    name: String,
    rename_input: String,
    index: usize,
    is_active: bool,
    is_renaming: bool,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    if is_renaming {
        sheet_tab_editing(rename_input, index)
    } else {
        sheet_tab(name, index, is_active, cx)
    }
}

/// Render a single sheet tab (normal mode)
fn sheet_tab(name: String, index: usize, is_active: bool, cx: &mut Context<Spreadsheet>) -> Stateful<Div> {
    div()
        .id(ElementId::Name(format!("sheet-tab-{}", index).into()))
        .px_2()
        .py_px()
        .cursor_pointer()
        .rounded_sm()
        .when(is_active, |d: Stateful<Div>| {
            d.bg(rgb(0x1e1e1e))
                .border_1()
                .border_color(rgb(0x3d3d3d))
                .text_color(rgb(0xcccccc))
        })
        .when(!is_active, |d: Stateful<Div>| {
            d.text_color(rgb(0x858585))
                .hover(|s| s.bg(rgb(0x2d2d2d)).text_color(rgb(0xcccccc)))
        })
        // Click to switch sheet
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.goto_sheet(index, cx);
        }))
        // Right-click for context menu
        .on_mouse_down(MouseButton::Right, cx.listener(move |this, _, _, cx| {
            this.show_sheet_context_menu(index, cx);
        }))
        .child(name)
}

/// Render a sheet tab in editing/rename mode
fn sheet_tab_editing(current_value: String, index: usize) -> Stateful<Div> {
    let display_value = if current_value.is_empty() {
        " ".to_string()
    } else {
        current_value
    };

    div()
        .id(ElementId::Name(format!("sheet-tab-edit-{}", index).into()))
        .px_1()
        .py_px()
        .bg(rgb(0x1e1e1e))
        .border_1()
        .border_color(rgb(0x007acc))
        .rounded_sm()
        .child(
            div()
                .min_w(px(40.0))
                .text_color(rgb(0xcccccc))
                .child(display_value)
        )
}

/// Render the sheet context menu
fn render_sheet_context_menu(sheet_index: usize, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    div()
        .absolute()
        .bottom(px(24.0))
        .left(px(4.0 + (sheet_index as f32 * 70.0))) // Approximate position
        .w(px(120.0))
        .bg(rgb(0x252526))
        .border_1()
        .border_color(rgb(0x3d3d3d))
        .rounded_sm()
        .shadow_lg()
        .py_1()
        .child(
            // Insert option
            div()
                .id("ctx-insert")
                .px_3()
                .py_1()
                .cursor_pointer()
                .text_color(rgb(0xcccccc))
                .text_xs()
                .hover(|s| s.bg(rgb(0x094771)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.hide_sheet_context_menu(cx);
                    this.add_sheet(cx);
                }))
                .child("Insert")
        )
        .child(
            // Delete option
            div()
                .id("ctx-delete")
                .px_3()
                .py_1()
                .cursor_pointer()
                .text_color(rgb(0xcccccc))
                .text_xs()
                .hover(|s| s.bg(rgb(0x094771)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.delete_sheet(sheet_index, cx);
                }))
                .child("Delete")
        )
        .child(
            // Rename option
            div()
                .id("ctx-rename")
                .px_3()
                .py_1()
                .cursor_pointer()
                .text_color(rgb(0xcccccc))
                .text_xs()
                .hover(|s| s.bg(rgb(0x094771)))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.hide_sheet_context_menu(cx);
                    this.start_sheet_rename(sheet_index, cx);
                }))
                .child("Rename")
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
