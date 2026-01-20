use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Render the bottom status bar (Zed-inspired minimal design)
pub fn render_status_bar(app: &Spreadsheet, editing: bool, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    // Calculate selection stats if multiple cells selected
    let selection_stats = calculate_selection_stats(app);

    // Mode indicator - show contextual tip when user has named ranges (post-activation)
    let has_named_ranges = !app.workbook.list_named_ranges().is_empty();
    let mode_text = if editing {
        "Edit"
    } else if app.status_message.is_some() {
        ""
    } else if has_named_ranges {
        "Tip: Named ranges let you refactor spreadsheets safely."
    } else {
        "Ready"
    };

    // Get sheet information
    let sheet_names = app.workbook.sheet_names();
    let active_index = app.workbook.active_sheet_index();
    let renaming_sheet = app.renaming_sheet;
    let rename_input = app.sheet_rename_input.clone();
    let context_menu_sheet = app.sheet_context_menu;

    // Theme colors
    let status_bg = app.token(TokenKey::StatusBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_primary = app.token(TokenKey::TextPrimary);
    let panel_bg = app.token(TokenKey::PanelBg);

    div()
        .relative()
        .flex_shrink_0()
        .h(px(22.0))
        .bg(panel_bg)
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .items_center()
        .justify_between()
        .px_2()
        .text_color(text_muted)
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
                        sheet_tab_wrapper(app, name_str, input_str, idx, is_active, is_renaming, cx)
                    })
                )
                // Add sheet button
                .child(
                    div()
                        .id("add-sheet-btn")
                        .px_1()
                        .py_px()
                        .cursor_pointer()
                        .text_color(text_muted)
                        .hover(move |s| s.text_color(text_primary).bg(panel_border))
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
                        .bg(panel_border)
                        .mx_2()
                )
                // Status message or mode
                .child(render_status_message(app, mode_text, text_muted, cx))
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
            d.child(render_sheet_context_menu(app, context_menu_sheet.unwrap(), cx))
        })
}

/// Wrapper to return consistent type for sheet tabs
fn sheet_tab_wrapper(
    app: &Spreadsheet,
    name: String,
    rename_input: String,
    index: usize,
    is_active: bool,
    is_renaming: bool,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    if is_renaming {
        sheet_tab_editing(app, rename_input, index)
    } else {
        sheet_tab(app, name, index, is_active, cx)
    }
}

/// Render a single sheet tab (normal mode)
fn sheet_tab(app: &Spreadsheet, name: String, index: usize, is_active: bool, cx: &mut Context<Spreadsheet>) -> Stateful<Div> {
    let app_bg = app.token(TokenKey::AppBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let header_bg = app.token(TokenKey::HeaderBg);

    div()
        .id(ElementId::Name(format!("sheet-tab-{}", index).into()))
        .px_2()
        .py_px()
        .cursor_pointer()
        .rounded_sm()
        .when(is_active, move |d: Stateful<Div>| {
            d.bg(app_bg)
                .border_1()
                .border_color(panel_border)
                .text_color(text_primary)
        })
        .when(!is_active, move |d: Stateful<Div>| {
            d.text_color(text_muted)
                .hover(move |s| s.bg(header_bg).text_color(text_primary))
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
fn sheet_tab_editing(app: &Spreadsheet, current_value: String, index: usize) -> Stateful<Div> {
    let display_value = if current_value.is_empty() {
        " ".to_string()
    } else {
        current_value
    };

    let app_bg = app.token(TokenKey::AppBg);
    let accent = app.token(TokenKey::Accent);
    let text_primary = app.token(TokenKey::TextPrimary);

    div()
        .id(ElementId::Name(format!("sheet-tab-edit-{}", index).into()))
        .px_1()
        .py_px()
        .bg(app_bg)
        .border_1()
        .border_color(accent)
        .rounded_sm()
        .child(
            div()
                .min_w(px(40.0))
                .text_color(text_primary)
                .child(display_value)
        )
}

/// Render the sheet context menu
fn render_sheet_context_menu(app: &Spreadsheet, sheet_index: usize, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let selection_bg = app.token(TokenKey::SelectionBg);

    div()
        .absolute()
        .bottom(px(24.0))
        .left(px(4.0 + (sheet_index as f32 * 70.0))) // Approximate position
        .w(px(120.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
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
                .text_color(text_primary)
                .text_xs()
                .hover(move |s| s.bg(selection_bg))
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
                .text_color(text_primary)
                .text_xs()
                .hover(move |s| s.bg(selection_bg))
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
                .text_color(text_primary)
                .text_xs()
                .hover(move |s| s.bg(selection_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    this.hide_sheet_context_menu(cx);
                    this.start_sheet_rename(sheet_index, cx);
                }))
                .child("Rename")
        )
}

/// Render the status message, making it clickable if an import report is available
fn render_status_message(app: &Spreadsheet, mode_text: &str, text_muted: Hsla, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let has_import_result = app.import_result.is_some();
    let accent = app.token(TokenKey::Accent);

    let message = if let Some(msg) = &app.status_message {
        msg.clone()
    } else {
        mode_text.to_string()
    };

    // If there's an import result, make the message clickable
    if has_import_result && app.status_message.is_some() {
        div()
            .id("status-message")
            .text_color(accent)
            .cursor_pointer()
            .hover(|s| s.underline())
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                this.show_import_report(cx);
            }))
            .child(message)
            .into_any_element()
    } else {
        div()
            .text_color(text_muted)
            .child(message)
            .into_any_element()
    }
}

/// Calculate statistics for the current selection
fn calculate_selection_stats(app: &Spreadsheet) -> Vec<Div> {
    let ((min_row, min_col), (max_row, max_col)) = app.selection_range();
    let text_muted = app.token(TokenKey::TextMuted);
    let text_primary = app.token(TokenKey::TextPrimary);

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
            stat_item("Count", &count.to_string(), text_muted, text_primary),
        ];
    }

    let sum: f64 = values.iter().sum();
    let avg = sum / values.len() as f64;
    let min = values.iter().cloned().fold(f64::INFINITY, f64::min);
    let max = values.iter().cloned().fold(f64::NEG_INFINITY, f64::max);

    vec![
        stat_item("Sum", &format_number(sum), text_muted, text_primary),
        stat_item("Average", &format_number(avg), text_muted, text_primary),
        stat_item("Min", &format_number(min), text_muted, text_primary),
        stat_item("Max", &format_number(max), text_muted, text_primary),
        stat_item("Count", &count.to_string(), text_muted, text_primary),
    ]
}

fn stat_item(label: &str, value: &str, label_color: Hsla, value_color: Hsla) -> Div {
    div()
        .flex()
        .items_center()
        .gap_1()
        .child(
            div()
                .text_color(label_color)
                .child(format!("{}:", label))
        )
        .child(
            div()
                .text_color(value_color)
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
