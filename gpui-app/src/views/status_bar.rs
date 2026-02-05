use gpui::{*};
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::links::LinkTarget;
use crate::mode::Mode;
use crate::theme::TokenKey;
use crate::ui::popup;

/// Render the bottom status bar (Zed-inspired minimal design)
pub fn render_status_bar(app: &Spreadsheet, editing: bool, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    // Calculate selection stats if multiple cells selected
    let selection_stats = calculate_selection_stats(app, cx);

    // Detect link in current cell (only in navigation mode with single cell selected)
    // Don't show hint if multi-selection is active (Ctrl+Enter won't work anyway)
    let detected_link = if !editing && app.mode == Mode::Navigation && !app.is_multi_selection() {
        app.detected_link(cx)
    } else {
        None
    };

    // Mode indicator - show current mode (VIM/HINT/NAV/EDIT)
    let vim_enabled = app.vim_mode_enabled(cx);
    let mode_text = if app.mode == Mode::Hint {
        // Show hint buffer as user types
        if app.hint_state.buffer.is_empty() {
            "HINT"
        } else {
            "" // We'll show the buffer separately
        }
    } else if editing {
        "EDIT"
    } else if app.status_message.is_some() {
        ""
    } else if vim_enabled {
        "VIM"
    } else {
        "NAV"
    };

    // Hint buffer display (when typing in hint mode)
    let hint_buffer = if app.mode == Mode::Hint && !app.hint_state.buffer.is_empty() {
        Some(format!("HINT: {}", app.hint_state.buffer))
    } else {
        None
    };

    // Get sheet information (convert to owned Strings to break borrow on cx for closure usage)
    let sheet_names: Vec<String> = app.wb(cx).sheet_names().iter().map(|s| s.to_string()).collect();
    let active_index = app.wb(cx).active_sheet_index();
    let renaming_sheet = app.renaming_sheet;
    let context_menu_sheet = app.sheet_context_menu;

    // Theme colors
    let _status_bg = app.token(TokenKey::StatusBg);
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
                        sheet_tab_wrapper(app, name_str, idx, is_active, is_renaming, cx)
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
                .child(render_status_message(app, mode_text, hint_buffer.as_deref(), text_muted, cx))
        )
        .child(
            // Right side: multi-selection hint + link hint + selection stats
            div()
                .flex()
                .items_center()
                .gap_4()
                // Multi-selection hint (context-aware)
                .when(app.is_multi_selection(), |d| {
                    let hint = if editing {
                        "Enter to apply · Esc to cancel"
                    } else {
                        "Type to edit all · Ctrl+Enter to fill"
                    };
                    d.child(
                        div()
                            .flex()
                            .items_center()
                            .text_color(text_muted)
                            .child(hint)
                    )
                })
                // Link hint (when link detected in current cell)
                .when(detected_link.is_some(), |d| {
                    let link_type = match detected_link.as_ref() {
                        Some(LinkTarget::Url(_)) => "URL",
                        Some(LinkTarget::Email(_)) => "Email",
                        Some(LinkTarget::Path(_)) => "File",
                        None => "",
                    };
                    d.child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            .text_color(text_muted)
                            .child(format!("{}: Ctrl+Enter to open", link_type))
                    )
                })
                .children(selection_stats)
                // Trace mode indicator (Alt+T)
                .when(app.trace_enabled, |d| {
                    let trace_summary = app.trace_summary().unwrap_or_else(|| "Trace: (range selected)".to_string());
                    let trace_color = app.token(TokenKey::Ok); // Green semantic color
                    d.child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            .text_color(trace_color)
                            .child(trace_summary)
                    )
                })
                // Hub sync status indicator (always shown)
                .child(render_hub_indicator(app, cx))
                // Verified mode indicator
                .when(app.verified_mode, |d| {
                    d.child(render_verified_indicator(app, cx))
                })
                // Approval status indicator (semantic fingerprint)
                .when(app.approved_fingerprint.is_some(), |d| {
                    d.child(render_approval_indicator(app, cx))
                })
                // Manual calc mode indicator
                .when(!app.wb(cx).auto_recalc(), |d| {
                    let warn_color = app.token(TokenKey::Warn);
                    d.child(
                        div()
                            .flex()
                            .items_center()
                            .text_color(warn_color)
                            .child("MANUAL CALC")
                    )
                })
                // Filter indicator (when filtering is active)
                .when(app.row_view.is_filtered(), |d| {
                    let visible = app.row_view.visible_count();
                    let total = app.row_view.row_count();
                    d.child(
                        div()
                            .flex()
                            .items_center()
                            .text_color(text_muted)
                            .child(format!("{} of {} rows", visible, total))
                    )
                })
                // Zoom indicator
                .child(
                    div()
                        .flex()
                        .items_center()
                        .text_color(text_muted)
                        .child(app.zoom_display())
                )
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
    index: usize,
    is_active: bool,
    is_renaming: bool,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    if is_renaming {
        sheet_tab_editing(app, cx)
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
        // Click to switch sheet, double-click to rename
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            if event.click_count == 2 {
                // Double-click: start rename
                this.start_sheet_rename(index, cx);
            } else {
                // Single click: switch to sheet
                this.goto_sheet(index, cx);
            }
        }))
        // Right-click for context menu (macOS: also ctrl+click)
        .on_mouse_down(MouseButton::Right, cx.listener(move |this, _, _, cx| {
            this.show_sheet_context_menu(index, cx);
        }))
        .child(name)
}

/// Render a sheet tab in editing/rename mode
fn sheet_tab_editing(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> Stateful<Div> {
    let input = &app.sheet_rename_input;
    let cursor_pos = app.sheet_rename_cursor;
    let select_all = app.sheet_rename_select_all;
    let caret_visible = app.caret_visible;

    let app_bg = app.token(TokenKey::AppBg);
    let accent = app.token(TokenKey::Accent);
    let text_primary = app.token(TokenKey::TextPrimary);
    let selection_bg = app.token(TokenKey::SelectionBg);

    // Build content with cursor
    let content = if select_all {
        // Show all text highlighted (selected)
        div()
            .flex()
            .child(
                div()
                    .bg(selection_bg)
                    .text_color(text_primary)
                    .child(if input.is_empty() { " ".to_string() } else { input.clone() })
            )
    } else {
        // Show text with cursor at position
        let (before, after) = if cursor_pos <= input.len() {
            (input[..cursor_pos].to_string(), input[cursor_pos..].to_string())
        } else {
            (input.clone(), String::new())
        };

        div()
            .flex()
            .child(div().text_color(text_primary).child(before))
            .when(caret_visible, |d| {
                d.child(
                    div()
                        .w(px(1.0))
                        .h(px(14.0))
                        .bg(text_primary)
                )
            })
            .child(div().text_color(text_primary).child(after))
    };

    // Click outside to confirm (on the editing tab itself does nothing special)
    let index = app.renaming_sheet.unwrap_or(0);

    div()
        .id(ElementId::Name(format!("sheet-tab-edit-{}", index).into()))
        .px_1()
        .py_px()
        .bg(app_bg)
        .border_1()
        .border_color(accent)
        .rounded_sm()
        // Click on the tab while editing - deselect all and place cursor at end
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            if this.sheet_rename_select_all {
                this.sheet_rename_select_all = false;
                this.sheet_rename_cursor = this.sheet_rename_input.len();
                cx.notify();
            }
        }))
        .child(
            div()
                .min_w(px(40.0))
                .child(content)
        )
}

/// Render the sheet context menu
fn render_sheet_context_menu(app: &Spreadsheet, sheet_index: usize, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let selection_bg = app.token(TokenKey::SelectionBg);

    popup("sheet-context-menu", panel_bg, panel_border, |this, cx| this.hide_sheet_context_menu(cx), cx)
        .bottom(px(24.0))
        .left(px(4.0 + (sheet_index as f32 * 70.0))) // Approximate position
        .w(px(120.0))
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
fn render_status_message(
    app: &Spreadsheet,
    mode_text: &str,
    hint_buffer: Option<&str>,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let has_import_result = app.import_result.is_some();
    let accent = app.token(TokenKey::Accent);
    let warning = app.token(TokenKey::Warn);

    // Highest priority: preview mode banner
    if let Some(session) = app.preview_session() {
        return div()
            .flex()
            .items_center()
            .gap_2()
            .child(
                div()
                    .px_2()
                    .py_px()
                    .bg(warning.opacity(0.2))
                    .rounded_sm()
                    .text_color(warning)
                    .font_weight(FontWeight::MEDIUM)
                    .child("PREVIEW")
            )
            .child(
                div()
                    .text_color(text_muted)
                    .child(format!("Before \"{}\" — Release Space to return", session.action_summary))
            )
            .into_any_element();
    }

    // Priority: hint buffer > status message > mode text
    let message = if let Some(hint) = hint_buffer {
        hint.to_string()
    } else if let Some(msg) = &app.status_message {
        msg.clone()
    } else {
        mode_text.to_string()
    };

    // Show a copy icon when there's a status message (perf reports, etc.)
    let show_copy = app.status_message.is_some() && hint_buffer.is_none();
    let copy_msg = if show_copy { Some(message.clone()) } else { None };
    let panel_border = app.token(TokenKey::PanelBorder);

    // If there's an import result, make the message clickable
    let msg_el: AnyElement = if has_import_result && app.status_message.is_some() {
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
    } else if hint_buffer.is_some() {
        // Hint mode gets accent color to stand out
        div()
            .text_color(accent)
            .child(message)
            .into_any_element()
    } else {
        div()
            .text_color(text_muted)
            .child(message)
            .into_any_element()
    };

    let row = div()
        .flex()
        .items_center()
        .gap_1()
        .child(msg_el)
        .when(show_copy, |d| {
            d.child(
                div()
                    .id("status-copy-btn")
                    .cursor_pointer()
                    .px_1()
                    .rounded_sm()
                    .text_color(text_muted)
                    .hover(move |s| s.text_color(accent).bg(panel_border))
                    .on_mouse_down(MouseButton::Left, cx.listener(move |_this, _, _window, cx| {
                        if let Some(msg) = &copy_msg {
                            cx.write_to_clipboard(ClipboardItem::new_string(msg.clone()));
                        }
                    }))
                    .child("⧉")
            )
        });

    row.into_any_element()
}

/// Maximum number of cells to analyze for statistics (prevents UI freeze on large selections)
const MAX_STATS_CELLS: usize = 10_000;

/// Calculate statistics for the current selection
fn calculate_selection_stats(app: &Spreadsheet, cx: &App) -> Vec<Div> {
    let ((min_row, min_col), (max_row, max_col)) = app.selection_range();
    let text_muted = app.token(TokenKey::TextMuted);
    let text_primary = app.token(TokenKey::TextPrimary);

    // Only show stats if more than one cell is selected
    let is_multi_select = min_row != max_row || min_col != max_col;
    if !is_multi_select {
        return vec![];
    }

    // Calculate total cell count without iterating
    let row_count = max_row - min_row + 1;
    let col_count = max_col - min_col + 1;
    let total_cells = row_count * col_count;

    // For very large selections, just show the count to avoid freezing
    // Show a clear message so users don't think stats are broken
    if total_cells > MAX_STATS_CELLS {
        return vec![
            stat_item("Count", &format_large_number(total_cells), text_muted, text_primary),
            stat_item("", &format!("(stats disabled for selections > {} cells)", format_large_number(MAX_STATS_CELLS)), text_muted, text_muted),
        ];
    }

    // Collect numeric values from selection
    let mut values: Vec<f64> = Vec::new();
    let mut count = 0usize;

    for row in min_row..=max_row {
        for col in min_col..=max_col {
            count += 1;
            let display = app.sheet(cx).get_display(row, col);
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

/// Format a large number with thousands separators for readability
fn format_large_number(n: usize) -> String {
    let s = n.to_string();
    let mut result = String::new();
    for (i, c) in s.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    result.chars().rev().collect()
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

/// Render the verified mode indicator with hover tooltip
fn render_verified_indicator(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let success_color = app.token(TokenKey::Ok);
    let error_color = app.token(TokenKey::Error);
    let text_muted = app.token(TokenKey::TextMuted);
    let _panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);

    // Determine status based on last recalc report
    let (status_text, status_color, _has_issues) = if let Some(report) = &app.last_recalc_report {
        if report.had_cycles || !report.errors.is_empty() {
            ("Verified ⚠", error_color, true)
        } else {
            ("Verified ✓", success_color, false)
        }
    } else {
        ("Verified ✓", success_color, false)
    };

    // Build tooltip content
    let tooltip_content = if let Some(report) = &app.last_recalc_report {
        format!(
            "{}ms · {} cells · depth {}{}",
            report.duration_ms,
            report.cells_recomputed,
            report.max_depth,
            if report.unknown_deps_recomputed > 0 {
                format!(" · {} dynamic", report.unknown_deps_recomputed)
            } else {
                String::new()
            }
        )
    } else {
        "Values are current".to_string()
    };

    div()
        .id("verified-indicator")
        .flex()
        .items_center()
        .gap_1()
        .px_2()
        .py_px()
        .rounded_sm()
        .cursor_pointer()
        .text_color(status_color)
        .hover(move |s| s.bg(panel_border))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.toggle_verified_mode(cx);
        }))
        .child(status_text)
        // Show tooltip on hover with recalc stats
        .child(
            div()
                .text_color(text_muted)
                .text_xs()
                .child(format!("({})", tooltip_content))
        )
}

/// Render the semantic approval status indicator
fn render_approval_indicator(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::app::ApprovalStatus;

    let success_color = app.token(TokenKey::Ok);
    let warning_color = app.token(TokenKey::Warn);
    let text_muted = app.token(TokenKey::TextMuted);
    let panel_border = app.token(TokenKey::PanelBorder);

    let status = app.approval_status();
    let (status_text, status_color) = match status {
        ApprovalStatus::NotApproved => ("", text_muted), // Shouldn't render
        ApprovalStatus::Approved => ("Approved ✓", success_color),
        ApprovalStatus::Drifted => ("Drifted ⚠", warning_color),
    };

    // Build tooltip content - show click hint
    let tooltip = match status {
        ApprovalStatus::Approved => "Click to clear approval",
        ApprovalStatus::Drifted => "Click to approve new logic",
        ApprovalStatus::NotApproved => "",
    };

    div()
        .id("approval-indicator")
        .flex()
        .items_center()
        .gap_1()
        .px_2()
        .py_px()
        .rounded_sm()
        .cursor_pointer()
        .text_color(status_color)
        .hover(move |s| s.bg(panel_border))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            // Click to re-approve if drifted (shows confirmation), or clear if approved
            if this.is_approved() {
                this.clear_approval(cx);
            } else {
                // Shows confirmation dialog since we're drifted
                this.approve_model(None, cx);
            }
        }))
        .child(status_text)
        .child(
            div()
                .text_color(text_muted)
                .text_xs()
                .child(format!("({})", tooltip))
        )
}

/// Render the VisiHub sync status indicator
fn render_hub_indicator(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::hub::HubStatus;

    let text_muted = app.token(TokenKey::TextMuted);
    let panel_border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let success_color = app.token(TokenKey::Ok);
    let warning_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);

    // Determine icon and color based on status
    let (icon, color) = match app.hub_status {
        HubStatus::Unlinked => ("☁", text_muted),
        HubStatus::Idle => ("☁ ✓", success_color),
        HubStatus::Ahead => ("☁ ↑", accent),
        HubStatus::Behind => ("☁ ↓", accent),
        HubStatus::Diverged => ("☁ ↕", warning_color),
        HubStatus::Syncing => ("☁ ⟳", text_muted),
        HubStatus::Offline => ("☁ ✗", text_muted),
        HubStatus::Forbidden => ("☁ 🔒", error_color),
    };

    // Display text depends on whether linked
    let label_text = if app.hub_status == HubStatus::Unlinked {
        "Local".to_string()
    } else {
        // Display name from hub link
        let display_name = app.hub_link
            .as_ref()
            .map(|link| link.display_name())
            .unwrap_or_default();

        // When syncing, show the current activity instead of generic "Syncing..."
        let status_label = if app.hub_status == HubStatus::Syncing {
            app.hub_activity
                .map(|a| a.label())
                .unwrap_or(app.hub_status.label())
        } else {
            app.hub_status.label()
        };

        format!("{} · {}", display_name, status_label)
    };

    // Click action depends on status
    let status = app.hub_status;

    div()
        .id("hub-indicator")
        .flex()
        .items_center()
        .gap_1()
        .px_2()
        .py_px()
        .rounded_sm()
        .cursor_pointer()
        .text_color(color)
        .hover(move |s| s.bg(panel_border))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            match status {
                HubStatus::Unlinked => this.hub_show_link_dialog(cx),
                HubStatus::Behind => this.hub_pull(cx),
                HubStatus::Ahead | HubStatus::Diverged => this.hub_open_remote_as_copy(cx),
                _ => this.hub_check_status(cx),
            }
        }))
        .child(icon)
        .child(
            div()
                .text_color(text_muted)
                .text_xs()
                .child(label_text)
        )
}
