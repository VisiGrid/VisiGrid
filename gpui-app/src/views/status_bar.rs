use gpui::{*};
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::links::LinkTarget;
use crate::mode::Mode;
use crate::theme::TokenKey;
use crate::ui::popup;

/// Simple tooltip for status bar pills.
struct PillTooltip(SharedString);

impl Render for PillTooltip {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .px_2()
            .py_1()
            .rounded_sm()
            .bg(rgb(0x2d2d2d))
            .border_1()
            .border_color(rgb(0x3d3d3d))
            .text_size(px(11.0))
            .text_color(rgb(0xcccccc))
            .child(self.0.clone())
    }
}

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
    let selection_bg = app.token(TokenKey::SelectionBg);

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
                // Cloud/Hub sync status indicator (always shown)
                .child(render_cloud_indicator(app, cx))
                // Verified mode indicator
                .when(app.verified_mode, |d| {
                    d.child(render_verified_indicator(app, cx))
                })
                // Verification status indicator (semantic fingerprint)
                .when(app.semantic_verification.fingerprint.is_some(), |d| {
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
                // Iteration pill (highest precedence cycle indicator)
                .when(app.wb(cx).iterative_enabled(), |d| {
                    let ok_color = app.token(TokenKey::Ok);
                    let warn_color = app.token(TokenKey::Warn);
                    let converged = app.last_recalc_report.as_ref().map_or(true, |r| !r.had_cycles || r.converged);
                    let color = if converged { ok_color } else { warn_color };
                    let label = if converged { "ITER: ON \u{2713}" } else { "ITER: ON \u{26a0}" };
                    let tip: SharedString = if converged {
                        let scc = app.last_recalc_report.as_ref().map_or(0, |r| r.scc_count);
                        let iters = app.last_recalc_report.as_ref().map_or(0, |r| r.iterations_performed);
                        format!("Converged in {} iterations across {} cycle groups.", iters, scc).into()
                    } else {
                        "Did not converge \u{2014} cycle cells show #NUM!".into()
                    };
                    d.child(
                        div()
                            .id("iter-pill")
                            .flex().items_center()
                            .text_color(color)
                            .cursor_pointer()
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                this.show_import_report(cx);
                            }))
                            .tooltip(move |_window, cx| {
                                cx.new(|_| PillTooltip(tip.clone())).into()
                            })
                            .child(label)
                    )
                })
                // Frozen pill (show alongside iteration if both active, or alone)
                .when(app.import_result.as_ref().map_or(false, |r| r.freeze_applied && r.cycles_frozen > 0), |d| {
                    let warn_color = app.token(TokenKey::Warn);
                    let n = app.import_result.as_ref().map(|r| r.cycles_frozen).unwrap_or(0);
                    let tip: SharedString = format!("{} cells use Excel cached values (not recalculated).", n).into();
                    d.child(
                        div()
                            .id("frozen-pill")
                            .flex().items_center()
                            .text_color(warn_color)
                            .cursor_pointer()
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                this.show_import_report(cx);
                            }))
                            .tooltip(move |_window, cx| {
                                cx.new(|_| PillTooltip(tip.clone())).into()
                            })
                            .child(format!("FROZEN: {}", n))
                    )
                })
                // Cycles pill (only when strict mode — unresolved cycles, no freeze, no iteration)
                .when(
                    app.has_unresolved_cycles(cx)
                        && !app.import_result.as_ref().map_or(false, |r| r.freeze_applied)
                        && !app.wb(cx).iterative_enabled(),
                    |d| {
                    let error_color = app.token(TokenKey::Error);
                    let n = app.current_cycle_count(cx);
                    let tip: SharedString = format!("{} cells in circular references (#CYCLE!). Click for details.", n).into();
                    d.child(
                        div()
                            .id("cycles-pill")
                            .flex().items_center()
                            .text_color(error_color)
                            .cursor_pointer()
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                this.show_import_report(cx);
                            }))
                            .tooltip(move |_window, cx| {
                                cx.new(|_| PillTooltip(tip.clone())).into()
                            })
                            .child(format!("CYCLES: {}", n))
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
                // Separator before panel toggles
                .child(
                    div().w(px(1.0)).h(px(12.0)).bg(panel_border.opacity(0.3)).mx_1()
                )
                // Panel toggle icons
                .child(render_panel_toggles(app, text_muted, text_primary, selection_bg, panel_border, cx))
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

/// Render the semantic verification status indicator
fn render_approval_indicator(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::app::VerificationStatus;

    let success_color = app.token(TokenKey::Ok);
    let warning_color = app.token(TokenKey::Warn);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let panel_border = app.token(TokenKey::PanelBorder);

    let status = app.verification_status(cx);
    let is_drifted = status == VerificationStatus::Drifted;
    let is_verified = status == VerificationStatus::Verified;

    let (status_text, status_color): (&str, gpui::Hsla) = match status {
        VerificationStatus::Unverified => ("", text_muted), // Shouldn't render (filtered by caller)
        VerificationStatus::Verified => ("Verified ✓", success_color),
        VerificationStatus::Drifted => ("Drifted ⚠", warning_color),
    };

    // Build context string: label from verification (persisted)
    let label = app.semantic_verification.label.clone();
    let has_label = label.is_some();

    div()
        .id("approval-indicator")
        .flex()
        .items_center()
        .gap_1()
        .px_2()
        .py_px()
        .rounded_sm()
        .text_color(status_color)
        // Main status text - click to approve/clear
        .child(
            div()
                .cursor_pointer()
                .hover(move |s| s.bg(panel_border))
                .px_1()
                .rounded_sm()
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    if this.is_verified(cx) {
                        this.clear_approval(cx);
                    } else {
                        this.approve_model(None, cx);
                    }
                }))
                .child(status_text)
        )
        // Context: label (when verified)
        .when(has_label && is_verified, move |d| {
            d.child(
                div()
                    .text_color(text_muted)
                    .text_xs()
                    .child(format!("({})", label.as_deref().unwrap_or("")))
            )
        })
        // "Why?" link when drifted
        .when(is_drifted, |d| {
            d.child(
                div()
                    .id("approval-why-link")
                    .text_color(accent)
                    .text_xs()
                    .cursor_pointer()
                    .hover(|s| s.underline())
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        this.show_approval_drift(cx);
                    }))
                    .child("(Why?)")
            )
        })
}

/// Render the cloud/hub sync status indicator.
///
/// Priority: cloud_identity > hub_link > local.
/// If a file has a cloud identity, show cloud sync state.
/// If it has a hub_link (dataset), show the old hub indicator.
/// Otherwise show "Local".
fn render_cloud_indicator(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::cloud::CloudSyncState;
    use crate::hub::HubStatus;

    let text_muted = app.token(TokenKey::TextMuted);
    let panel_border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let success_color = app.token(TokenKey::Ok);
    let error_color = app.token(TokenKey::Error);

    // If cloud identity exists, show cloud sync state
    if let Some(ref identity) = app.cloud_identity {
        let (icon, color) = match app.cloud_sync_state {
            CloudSyncState::Local => ("☁", text_muted),
            CloudSyncState::Synced => ("☁ ✓", success_color),
            CloudSyncState::Dirty | CloudSyncState::Syncing => ("☁ ⟳", accent),
            CloudSyncState::Offline => ("☁ ✗", text_muted),
            CloudSyncState::Error => ("☁ !", error_color),
        };

        // Build label: "sheet_name" for synced, or state for others
        // Include account info from saved auth
        let account_slug = crate::hub::auth::load_auth()
            .and_then(|a| a.user_slug)
            .unwrap_or_default();

        let label_text = match app.cloud_sync_state {
            CloudSyncState::Synced => {
                if account_slug.is_empty() {
                    identity.sheet_name.clone()
                } else {
                    format!("{} · @{}", identity.sheet_name, account_slug)
                }
            }
            CloudSyncState::Syncing => "Syncing...".to_string(),
            CloudSyncState::Dirty => "Modified".to_string(),
            CloudSyncState::Offline => "Offline".to_string(),
            CloudSyncState::Error => app.cloud_last_error.clone().unwrap_or_else(|| "Error".to_string()),
            CloudSyncState::Local => identity.sheet_name.clone(),
        };

        let sync_state = app.cloud_sync_state;
        let public_id = identity.public_id.clone();

        return div()
            .id("cloud-indicator")
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
                match sync_state {
                    CloudSyncState::Offline | CloudSyncState::Error => this.cloud_retry_upload(cx),
                    CloudSyncState::Synced | CloudSyncState::Local | CloudSyncState::Dirty => {
                        // Open this specific sheet in the VisiGrid web app
                        let url = format!("https://app.visigrid.app/sheets/{}", public_id);
                        if let Err(e) = open::that(&url) {
                            this.status_message = Some(format!("Failed to open browser: {}", e));
                        }
                        cx.notify();
                    }
                    _ => {}
                }
            }))
            .child(icon)
            .child(
                div()
                    .text_color(text_muted)
                    .text_xs()
                    .child(label_text)
            );
    }

    // Fall back to hub_link indicator if dataset-linked
    if app.hub_link.is_some() {
        let warning_color = app.token(TokenKey::Warn);

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

        let display_name = app.hub_link
            .as_ref()
            .map(|link| link.display_name())
            .unwrap_or_default();

        let status_label = if app.hub_status == HubStatus::Syncing {
            app.hub_activity
                .map(|a| a.label())
                .unwrap_or(app.hub_status.label())
        } else {
            app.hub_status.label()
        };

        let label_text = format!("{} · {}", display_name, status_label);
        let status = app.hub_status;

        return div()
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
            );
    }

    // No cloud identity, no hub link — show "Local"
    let is_signed_in = crate::hub::auth::is_authenticated();

    div()
        .id("cloud-indicator")
        .flex()
        .items_center()
        .gap_1()
        .px_2()
        .py_px()
        .rounded_sm()
        .cursor_pointer()
        .text_color(text_muted)
        .hover(move |s| s.bg(panel_border))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            if is_signed_in {
                this.cloud_move_to_cloud(cx);
            } else {
                this.hub_sign_in(cx);
            }
        }))
        .child("☁")
        .child(
            div()
                .text_color(text_muted)
                .text_xs()
                .child("Local")
        )
}

/// Render the panel toggle icon buttons (Inspector, Profiler, Console, Minimap)
fn render_panel_toggles(
    app: &Spreadsheet,
    text_muted: Hsla,
    text_primary: Hsla,
    selection_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Div {
    div()
        .flex()
        .items_center()
        .gap(px(2.0))
        .child(panel_toggle_btn(
            "toggle-inspector",
            "\u{229E}",
            "Inspector (Ctrl+Shift+I)",
            app.inspector_visible,
            text_muted, text_primary, selection_bg, panel_border,
            cx,
            |this, _, cx| {
                this.inspector_visible = !this.inspector_visible;
                if this.inspector_visible { this.profiler_visible = false; }
                cx.notify();
            },
        ))
        .child(panel_toggle_btn(
            "toggle-profiler",
            "\u{23F1}",
            "Profiler (Ctrl+Alt+P)",
            app.profiler_visible,
            text_muted, text_primary, selection_bg, panel_border,
            cx,
            |this, _, cx| {
                this.profiler_visible = !this.profiler_visible;
                if this.profiler_visible { this.inspector_visible = false; }
                cx.notify();
            },
        ))
        .child(panel_toggle_btn(
            "toggle-console",
            "\u{276F}_",
            "Lua Console (Alt+F11)",
            app.lua_console.visible,
            text_muted, text_primary, selection_bg, panel_border,
            cx,
            |this, window, cx| {
                this.lua_console.toggle();
                if this.lua_console.visible {
                    window.focus(&this.console_focus_handle, cx);
                }
                cx.notify();
            },
        ))
        .child(panel_toggle_btn(
            "toggle-minimap",
            "\u{25A6}",
            "Minimap",
            app.minimap_visible,
            text_muted, text_primary, selection_bg, panel_border,
            cx,
            |this, _, cx| {
                this.minimap_visible = !this.minimap_visible;
                cx.notify();
            },
        ))
}

/// Render a single panel toggle button for the status bar
fn panel_toggle_btn(
    id: &'static str,
    icon: &'static str,
    tooltip_text: &'static str,
    active: bool,
    text_muted: Hsla,
    text_primary: Hsla,
    selection_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: impl Fn(&mut Spreadsheet, &mut Window, &mut Context<Spreadsheet>) + 'static,
) -> Stateful<Div> {
    let tip: SharedString = tooltip_text.into();
    div()
        .id(id)
        .px(px(4.0))
        .py(px(2.0))
        .text_size(px(11.0))
        .rounded(px(2.0))
        .cursor_pointer()
        .when(active, move |d: Stateful<Div>| {
            d.text_color(text_primary)
                .bg(selection_bg.opacity(0.3))
        })
        .when(!active, move |d: Stateful<Div>| {
            d.text_color(text_muted.opacity(0.5))
        })
        .hover(move |s| {
            s.text_color(text_primary)
                .bg(panel_border.opacity(0.5))
        })
        .tooltip(move |_window, cx| {
            cx.new(|_| PillTooltip(tip.clone())).into()
        })
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
            on_click(this, window, cx);
        }))
        .child(icon)
}
