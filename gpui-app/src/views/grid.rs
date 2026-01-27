use gpui::*;
use gpui::StyledText;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::fill::{FILL_HANDLE_BORDER, FILL_HANDLE_HIT_SIZE, FILL_HANDLE_VISUAL_SIZE};
use crate::mode::Mode;
use crate::settings::{user_settings, Setting};
use crate::theme::TokenKey;
use super::headers::render_row_header;
use super::formula_bar;
use visigrid_engine::cell::{Alignment, VerticalAlignment};
use visigrid_engine::formula::eval::Value;

/// Render the main cell grid with freeze pane support
///
/// When frozen_rows > 0 or frozen_cols > 0, renders 4 regions:
/// 1. Frozen corner (top-left, never scrolls)
/// 2. Frozen rows (top, scrolls horizontally only)
/// 3. Frozen cols (left, scrolls vertically only)
/// 4. Main grid (scrolls both directions)
pub fn render_grid(app: &mut Spreadsheet, window: &Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let scroll_row = app.view_state.scroll_row;
    let scroll_col = app.view_state.scroll_col;
    let selected = app.view_state.selected;
    let editing = app.mode.is_editing();
    let edit_value = app.edit_value.clone();
    let total_visible_rows = app.visible_rows();
    let total_visible_cols = app.visible_cols();
    let frozen_rows = app.view_state.frozen_rows;
    let frozen_cols = app.view_state.frozen_cols;

    // Read show_gridlines from global settings
    let show_gridlines = match &user_settings(cx).appearance.show_gridlines {
        Setting::Value(v) => *v,
        Setting::Inherit => true, // Default to showing gridlines
    };

    // Calculate scrollable region dimensions
    let scrollable_visible_rows = total_visible_rows.saturating_sub(frozen_rows);
    let scrollable_visible_cols = total_visible_cols.saturating_sub(frozen_cols);

    // Get divider color for freeze pane separators
    let divider_color = app.token(TokenKey::PanelBorder);

    // No freeze panes - simple single-region rendering
    if frozen_rows == 0 && frozen_cols == 0 {
        return div()
            .flex_1()
            .overflow_hidden()
            .relative()  // Enable absolute positioning for popup overlay
            .child(
                div()
                    .flex()
                    .flex_col()
                    .children(
                        (0..total_visible_rows).filter_map(|screen_row| {
                            // Get the view_row and data_row for this screen position
                            // This respects both sort order AND filter visibility
                            let visible_index = scroll_row + screen_row;
                            let (view_row, data_row) = app.nth_visible_row(visible_index)?;
                            Some(render_row(
                                view_row,
                                data_row,
                                scroll_col,
                                total_visible_cols,
                                selected,
                                editing,
                                &edit_value,
                                show_gridlines,
                                app,
                                window,
                                cx,
                            ))
                        })
                    )
            )
            // Popup overlay layer - positioned relative to grid, not window chrome
            .child(render_popup_overlay(app, cx))
            .into_any_element();
    }

    // Get metrics for scaled dimensions
    let metrics = &app.metrics;

    // Freeze panes active - render 4 regions
    div()
        .flex_1()
        .overflow_hidden()
        .relative()  // Enable absolute positioning for popup overlay
        .flex()
        .flex_col()
        // Top section: frozen corner + frozen rows
        .when(frozen_rows > 0, |d| {
            d.child(
                div()
                    .flex()
                    .flex_shrink_0()
                    .children(
                        (0..frozen_rows).map(|view_row| {
                            // Frozen rows: view_row == data_row (headers don't sort)
                            let data_row = app.view_to_data(view_row);
                            // Use scaled row height for rendering
                            let row_height = metrics.row_height(app.row_height(view_row));
                            div()
                                .flex()
                                .flex_shrink_0()
                                .h(px(row_height))
                                // Row header for frozen row
                                .child(render_row_header(app, view_row, cx))
                                // Frozen corner cells (cols 0..frozen_cols)
                                .when(frozen_cols > 0, |d| {
                                    d.children(
                                        (0..frozen_cols).map(|col| {
                                            let col_width = metrics.col_width(app.col_width(col));
                                            render_cell(view_row, data_row, col, col_width, row_height, selected, editing, &edit_value, show_gridlines, app, window, cx)
                                        })
                                    )
                                })
                                // Vertical divider after frozen cols (1px stays constant)
                                .when(frozen_cols > 0, |d| {
                                    d.child(
                                        div()
                                            .w(px(1.0))
                                            .h_full()
                                            .bg(divider_color)
                                    )
                                })
                                // Frozen row cells (cols scroll_col..scroll_col+scrollable_visible_cols)
                                .children(
                                    (0..scrollable_visible_cols).map(|visible_col| {
                                        let col = scroll_col + visible_col;
                                        let col_width = metrics.col_width(app.col_width(col));
                                        render_cell(view_row, data_row, col, col_width, row_height, selected, editing, &edit_value, show_gridlines, app, window, cx)
                                    })
                                )
                        })
                    )
            )
            // Horizontal divider after frozen rows (1px stays constant)
            .child(
                div()
                    .w_full()
                    .h(px(1.0))
                    .bg(divider_color)
            )
        })
        // Bottom section: frozen cols + main grid
        .child(
            div()
                .flex_1()
                .overflow_hidden()
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .children(
                            (0..scrollable_visible_rows).filter_map(|screen_row| {
                                // Get the view_row and data_row for this screen position
                                // Account for frozen rows + scroll position in visible index
                                let visible_index = frozen_rows + scroll_row + screen_row;
                                let (view_row, data_row) = app.nth_visible_row(visible_index)?;
                                let row_height = metrics.row_height(app.row_height(view_row));
                                Some(div()
                                    .flex()
                                    .flex_shrink_0()
                                    .h(px(row_height))
                                    // Row header
                                    .child(render_row_header(app, view_row, cx))
                                    // Frozen column cells (cols 0..frozen_cols)
                                    .when(frozen_cols > 0, |d| {
                                        d.children(
                                            (0..frozen_cols).map(|col| {
                                                let col_width = metrics.col_width(app.col_width(col));
                                                render_cell(view_row, data_row, col, col_width, row_height, selected, editing, &edit_value, show_gridlines, app, window, cx)
                                            })
                                        )
                                    })
                                    // Vertical divider after frozen cols (1px stays constant)
                                    .when(frozen_cols > 0, |d| {
                                        d.child(
                                            div()
                                                .w(px(1.0))
                                                .h_full()
                                                .bg(divider_color)
                                        )
                                    })
                                    // Main grid cells
                                    .children(
                                        (0..scrollable_visible_cols).map(|visible_col| {
                                            let col = scroll_col + visible_col;
                                            let col_width = metrics.col_width(app.col_width(col));
                                            render_cell(view_row, data_row, col, col_width, row_height, selected, editing, &edit_value, show_gridlines, app, window, cx)
                                        })
                                    ))
                            })
                        )
                )
        )
        // Popup overlay layer - positioned relative to grid, not window chrome
        .child(render_popup_overlay(app, cx))
        .into_any_element()
}

fn render_row(
    view_row: usize,
    data_row: usize,
    scroll_col: usize,
    visible_cols: usize,
    selected: (usize, usize),
    editing: bool,
    edit_value: &str,
    show_gridlines: bool,
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Use scaled dimensions for rendering (use view_row for consistent row heights)
    let row_height = app.metrics.row_height(app.row_height(view_row));

    div()
        .flex()
        .flex_shrink_0()
        .h(px(row_height))
        .child(render_row_header(app, view_row, cx))
        .children(
            (0..visible_cols).map(|visible_col| {
                let col = scroll_col + visible_col;
                let col_width = app.metrics.col_width(app.col_width(col));
                // view_row for selection/display, data_row for cell data access
                render_cell(view_row, data_row, col, col_width, row_height, selected, editing, edit_value, show_gridlines, app, window, cx)
            })
        )
}

fn render_cell(
    view_row: usize,
    data_row: usize,
    col: usize,
    col_width: f32,
    _row_height: f32,
    selected: (usize, usize),
    editing: bool,
    edit_value: &str,
    show_gridlines: bool,
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Selection uses view_row (what user sees/clicks)
    let is_selected = app.is_selected(view_row, col);
    let is_active = selected == (view_row, col);
    let is_editing = editing && is_active;
    let is_formula_ref = app.is_formula_ref(view_row, col);
    let formula_ref_color = app.formula_ref_color(view_row, col);  // Color index for multi-color refs
    let is_active_ref_target = app.is_active_ref_target(view_row, col);  // Live ref navigation target
    let is_inspector_hover = app.inspector_hover_cell == Some((view_row, col));  // Hover highlight from inspector

    // Check if cell is in trace path (Phase 3.5b)
    let sheet_id = app.sheet().id;
    let trace_position = app.inspector_trace_path.as_ref().and_then(|path| {
        path.iter().position(|cell| {
            cell.sheet == sheet_id && cell.row == view_row && cell.col == col
        }).map(|pos| {
            let is_start = pos == 0;
            let is_end = pos == path.len() - 1;
            (is_start, is_end)
        })
    });

    // Check if cell is in history highlight range (Phase 7A)
    let is_history_highlight = app.history_highlight_range.map_or(false, |(sheet_idx, sr, sc, er, ec)| {
        sheet_idx == app.workbook.active_sheet_index()
            && view_row >= sr && view_row <= er
            && col >= sc && col <= ec
    });

    // Check for hint mode and get hint label for this cell
    let hint_label = if app.mode == Mode::Hint {
        app.hint_state.labels.iter()
            .find(|h| h.row == view_row && h.col == col)
            .map(|h| {
                let buffer = &app.hint_state.buffer;
                let matches = h.label.starts_with(buffer);
                let is_unique = app.hint_state.matching_labels().len() == 1 && matches;
                (h.label.clone(), matches, is_unique)
            })
    } else {
        None
    };

    // Spill state detection - uses data_row (storage)
    let is_spill_parent = app.sheet().is_spill_parent(data_row, col);
    let is_spill_receiver = app.sheet().is_spill_receiver(data_row, col);
    let has_spill_error = app.sheet().has_spill_error(data_row, col);

    // Check for multi-edit preview (shows what each selected cell will receive)
    let multi_edit_preview = app.multi_edit_preview(view_row, col);
    let is_multi_edit_preview = multi_edit_preview.is_some();

    // Cell value: use data_row to access actual storage
    let value = if is_editing {
        edit_value.to_string()
    } else if let Some(preview) = multi_edit_preview {
        // Show the preview value for cells in multi-selection during editing
        preview
    } else if app.show_formulas() {
        app.sheet().get_raw(data_row, col)
    } else {
        let display = app.sheet().get_formatted_display(data_row, col);
        // Hide zero values if show_zeros is false
        if !app.show_zeros() && display == "0" {
            String::new()
        } else {
            display
        }
    };

    let format = app.sheet().get_format(data_row, col);
    let cell_row = view_row;  // For UI interactions (click, selection)
    let cell_col = col;

    // Determine border color based on cell state
    // Precedence: selection > ref_target > spill states > formula refs > gridlines
    let border_color = if has_spill_error {
        app.token(TokenKey::SpillBlockedBorder)
    } else if is_spill_parent {
        app.token(TokenKey::SpillBorder)
    } else if is_spill_receiver {
        app.token(TokenKey::SpillReceiverBorder)
    } else if is_active_ref_target && !is_selected && !is_active {
        // Ref target gets accent color ONLY if not also selected (selection wins)
        app.token(TokenKey::Accent)
    } else {
        cell_border(app, is_editing, is_active, is_selected, formula_ref_color)
    };

    let mut cell = div()
        .id(ElementId::Name(format!("cell-{}-{}", view_row, col).into()))
        .relative()  // Enable absolute positioning for selection overlay
        .size_full()
        .flex()
        .px_1()
        .overflow_hidden()
        .bg(cell_base_background(app, is_editing, format.background_color))
        .border_color(border_color);

    // Add selection/formula-ref overlay (semi-transparent, layered on top of cell background)
    // This allows custom background colors to show through the selection highlight
    if !is_editing && (is_active || is_selected || is_formula_ref) {
        if let Some(overlay_color) = selection_overlay_color(app, is_active, is_selected, formula_ref_color) {
            cell = cell.child(
                div()
                    .absolute()
                    .inset_0()
                    .bg(overlay_color)
            );
        }
    }

    // Inspector hover highlight (when hovering over a cell reference in the inspector panel)
    if is_inspector_hover && !is_selected && !is_active {
        let hover_color = app.token(TokenKey::Accent).opacity(0.3);
        cell = cell.child(
            div()
                .absolute()
                .inset_0()
                .bg(hover_color)
                .border_2()
                .border_color(app.token(TokenKey::Accent))
        );
    }

    // Trace path highlight (Phase 3.5b - when a trace is active)
    if let Some((is_start, is_end)) = trace_position {
        if !is_selected && !is_active {
            let accent = app.token(TokenKey::Accent);
            // Start/end cells get stronger emphasis
            let (bg_opacity, border_width) = if is_start || is_end {
                (0.25, px(2.0))
            } else {
                (0.12, px(1.0))
            };
            cell = cell.child(
                div()
                    .absolute()
                    .inset_0()
                    .bg(accent.opacity(bg_opacity))
                    .border(border_width)
                    .border_color(accent.opacity(0.6))
            );
        }
    }

    // History highlight (Phase 7A - when a history entry is selected)
    if is_history_highlight && !is_selected && !is_active {
        // Use orange/amber to distinguish from selection (blue) and formula refs (varied)
        let history_color: Hsla = rgb(0xf59e0b).into(); // Amber-500
        cell = cell.child(
            div()
                .absolute()
                .inset_0()
                .bg(history_color.opacity(0.25))
                .border_1()
                .border_color(history_color.opacity(0.7))
        );
    }

    // Apply horizontal alignment
    // When editing, always left-align so caret positioning works correctly
    // General alignment: numbers right-align, text/empty left-aligns (Excel behavior)
    cell = if is_editing {
        // Editing: always left-align for correct caret positioning
        cell.justify_start()
    } else {
        match format.alignment {
            Alignment::General => {
                let computed = app.sheet().get_computed_value(data_row, col);
                match computed {
                    Value::Number(_) => cell.justify_end(),
                    _ => cell.justify_start(),
                }
            }
            Alignment::Left => cell.justify_start(),
            Alignment::Center => cell.justify_center(),
            Alignment::Right => cell.justify_end(),
        }
    };

    // Apply vertical alignment
    // When editing, always top-align so caret positioning (fixed top offset) works correctly
    cell = if is_editing {
        cell.items_start()
    } else {
        match format.vertical_alignment {
            VerticalAlignment::Top => cell.items_start(),
            VerticalAlignment::Middle => cell.items_center(),
            VerticalAlignment::Bottom => cell.items_end(),
        }
    };

    // Only right+bottom borders for normal cells (thinner gridlines)
    // For selected cells, only draw outer edges of the selection (not interior borders)
    // For formula refs, only draw outer edges of the range (not interior borders)
    // Spill parent/blocked get 2px border, receiver gets 1px
    cell = if is_editing {
        // Editing cell gets full border
        cell.border_1()
    } else if is_selected {
        // Selected cells: only draw outer edges to avoid double borders
        let (top, right, bottom, left) = app.selection_borders(view_row, col);
        let mut c = cell;
        if top { c = c.border_t_1(); }
        if right { c = c.border_r_1(); }
        if bottom { c = c.border_b_1(); }
        if left { c = c.border_l_1(); }
        c
    } else if is_active_ref_target {
        // Active ref target: bright 2px border so user knows exactly where arrow keys are pointing
        // (only if not also selected - selection takes precedence and was handled above)
        let (top, right, bottom, left) = app.ref_target_borders(view_row, col);
        let mut c = cell;
        if top { c = c.border_t_2(); }
        if right { c = c.border_r_2(); }
        if bottom { c = c.border_b_2(); }
        if left { c = c.border_l_2(); }
        c
    } else if has_spill_error || is_spill_parent {
        // 2px solid border for spill parent and blocked cells
        cell.border_2()
    } else if is_spill_receiver {
        // 1px border for spill receivers (ideally dashed, but gpui doesn't support that)
        cell.border_1()
    } else if is_formula_ref {
        // Get which borders to draw for this formula ref cell
        let (top, right, bottom, left) = app.formula_ref_borders(view_row, col);
        let mut c = cell;
        if top { c = c.border_t_1(); }
        if right { c = c.border_r_1(); }
        if bottom { c = c.border_b_1(); }
        if left { c = c.border_l_1(); }
        c
    } else {
        // Check for user-defined borders first (stored per data cell)
        let (user_top, user_right, user_bottom, user_left) = app.cell_user_borders(data_row, col);
        let has_user_border = user_top || user_right || user_bottom || user_left;

        if has_user_border {
            // Draw user-defined borders (black, 1px)
            let border_color = rgb(0x000000);  // Black for user borders
            let mut c = cell.border_color(border_color);
            if user_top { c = c.border_t_1(); }
            if user_right { c = c.border_r_1(); }
            if user_bottom { c = c.border_b_1(); }
            if user_left { c = c.border_l_1(); }
            c
        } else if show_gridlines {
            // Normal gridlines (only when enabled in settings)
            // Don't draw gridlines toward selected cells (selection borders handle those edges)
            let cell_right_selected = app.is_selected(view_row, col + 1);
            let cell_below_selected = app.is_selected(view_row + 1, col);
            let mut c = cell;
            if !cell_right_selected { c = c.border_r_1(); }
            if !cell_below_selected { c = c.border_b_1(); }
            c
        } else {
            // No gridlines and no user borders - plain cell
            cell
        }
    };

    cell = cell
        .text_color(cell_text_color(app, is_editing, is_selected, is_multi_edit_preview))
        .text_size(px(app.metrics.font_size))  // Scaled font size for zoom
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Clicking on grid while renaming sheet: confirm and continue
            if this.renaming_sheet.is_some() {
                this.confirm_sheet_rename(cx);
            }
            // Don't handle clicks if modal/overlay is visible
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            // Don't handle clicks if we're resizing
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
                return;
            }
            // Don't handle clicks if fill handle drag was just started (child handler fired first)
            if this.is_fill_dragging() {
                return;
            }

            // If clicking a spill receiver, redirect to the spill parent
            let (target_row, target_col) = if let Some((parent_row, parent_col)) = this.sheet().get_spill_parent(cell_row, cell_col) {
                (parent_row, parent_col)
            } else {
                (cell_row, cell_col)
            };

            // Formula mode: clicks insert cell references, drag for range
            if this.mode.is_formula() {
                if event.modifiers.shift {
                    this.formula_shift_click_ref(target_row, target_col, cx);
                } else {
                    // Start drag for range selection in formula mode
                    this.formula_start_drag(target_row, target_col, cx);
                }
                return;
            }

            // Normal mode handling
            if event.click_count == 2 {
                // Double-click to edit
                this.select_cell(target_row, target_col, false, cx);
                this.start_edit(cx);
                // On macOS, show tip about enabling F2 (since they're using fallback)
                this.maybe_show_f2_tip(cx);
            } else if event.modifiers.shift {
                // Shift+click extends selection
                this.select_cell(target_row, target_col, true, cx);
            } else if event.modifiers.control || event.modifiers.platform {
                // Ctrl+click (or Cmd on Mac) for discontiguous selection
                this.start_ctrl_drag_selection(target_row, target_col, cx);
            } else {
                // Start drag selection
                this.start_drag_selection(target_row, target_col, cx);
            }
        }))
        .on_mouse_move(cx.listener(move |this, _event: &MouseMoveEvent, _, cx| {
            // Don't handle if modal/overlay is visible
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            // Continue fill handle drag if active (priority over selection drag)
            if this.is_fill_dragging() {
                this.continue_fill_drag(cell_row, cell_col, cx);
                return;
            }
            // Continue drag selection if active
            if this.dragging_selection {
                if this.mode.is_formula() {
                    this.formula_continue_drag(cell_row, cell_col, cx);
                } else {
                    this.continue_drag_selection(cell_row, cell_col, cx);
                }
            }
        }))
        .on_mouse_up(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // Don't handle if modal/overlay is visible
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            // End fill handle drag if active (commits the fill)
            if this.is_fill_dragging() {
                this.end_fill_drag(cx);
                return;
            }
            // End drag selection (works for both normal and formula mode)
            this.end_drag_selection(cx);
        }));

    // Build the text content with selection highlight (caret drawn as overlay)
    if is_editing {
        // edit_cursor and selection range are already byte offsets
        let cursor_byte = app.edit_cursor;
        let selection = app.edit_selection_range();

        // Use raw buffer - caret is drawn as overlay, not injected into text
        let display_text: SharedString = value.into();
        let total_bytes = display_text.len();
        let byte_index = cursor_byte.min(total_bytes);

        // Shape text to get caret position for rendering
        // For empty text, use a space as surrogate for baseline metrics
        let text_len = display_text.len();
        let shape_text: SharedString = if text_len == 0 {
            " ".into()
        } else {
            display_text.clone()
        };
        let shape_len = shape_text.len();
        let shaped = window.text_system().shape_line(
            shape_text,
            px(app.metrics.font_size),
            &[TextRun {
                len: shape_len,
                font: Font::default(),
                color: Hsla::default(),
                background_color: None,
                underline: None,
                strikethrough: None,
            }],
            None,
        );

        // Use persisted scroll offset (computed by ensure_caret_visible in render_grid)
        let padding = 4.0;
        let scroll_x = app.edit_scroll_x;
        let caret_x: f32 = if total_bytes == 0 {
            0.0
        } else {
            shaped.x_for_index(byte_index).into()
        };

        // Create styled text with selection highlighting
        let text_element = if let Some((sel_start_byte, sel_end_byte)) = selection {
            // Selection positions are already byte offsets
            let byte_sel_start = sel_start_byte.min(total_bytes);
            let byte_sel_end = sel_end_byte.min(total_bytes);

            let normal_color = cell_text_color(app, is_editing, is_selected, is_multi_edit_preview);
            let selection_bg = app.token(TokenKey::EditorSelectionBg);
            let selection_fg = app.token(TokenKey::EditorSelectionText);

            let mut runs = Vec::new();

            // Before selection
            if byte_sel_start > 0 {
                runs.push(TextRun {
                    len: byte_sel_start,
                    font: Font::default(),
                    color: normal_color,
                    background_color: None,
                    underline: None,
                    strikethrough: None,
                });
            }

            // Selected text
            if byte_sel_end > byte_sel_start {
                runs.push(TextRun {
                    len: byte_sel_end - byte_sel_start,
                    font: Font::default(),
                    color: selection_fg,
                    background_color: Some(selection_bg),
                    underline: None,
                    strikethrough: None,
                });
            }

            // After selection
            if total_bytes > byte_sel_end {
                runs.push(TextRun {
                    len: total_bytes - byte_sel_end,
                    font: Font::default(),
                    color: normal_color,
                    background_color: None,
                    underline: None,
                    strikethrough: None,
                });
            }

            // Debug assert: run lengths must sum to text length (prevents caret drift)
            debug_assert_eq!(
                runs.iter().map(|r| r.len).sum::<usize>(),
                total_bytes,
                "TextRun lengths don't match text length"
            );

            StyledText::new(display_text).with_runs(runs).into_any_element()
        } else {
            // No selection - plain text
            display_text.into_any_element()
        };

        // Wrap text in scrolling container (positioned with scroll offset)
        cell = cell.child(
            div()
                .absolute()
                .top_0()
                .bottom_0()
                .left(px(padding + scroll_x))
                .flex()
                .items_center()
                .child(text_element)
        );

        // Draw caret as overlay rect when visible
        if app.caret_visible && selection.is_none() {
            let caret_color = app.token(TokenKey::TextPrimary);
            let line_height = app.metrics.row_height(app.row_height(view_row)) - 4.0;

            // Caret position accounts for scroll offset
            let visual_caret_x = padding + caret_x + scroll_x;

            cell = cell.child(
                div()
                    .absolute()
                    .left(px(visual_caret_x))
                    .top(px(2.0))
                    .w(px(1.5))
                    .h(px(line_height))
                    .bg(caret_color)
            );
        }
    } else {
        // Not editing - show value with formatting using StyledText
        let text_content: SharedString = value.into();

        // Check if any formatting is applied
        let has_formatting = format.bold || format.italic || format.underline;

        if has_formatting {
            // Get base text style from window and apply cell formatting
            let mut text_style = window.text_style();
            text_style.color = cell_text_color(app, is_editing, is_selected, is_multi_edit_preview);

            // Note: Bold/italic font variants may not render on Linux due to gpui limitations
            // with cosmic-text font selection. Underline works because it's drawn separately.
            // See: https://github.com/zed-industries/zed - Linux text system TODOs
            if format.bold {
                text_style.font_weight = FontWeight::BOLD;
            }
            if format.italic {
                text_style.font_style = FontStyle::Italic;
            }
            if format.underline {
                text_style.underline = Some(UnderlineStyle {
                    thickness: px(1.),
                    ..Default::default()
                });
            }

            cell = cell.child(StyledText::new(text_content).with_default_highlights(&text_style, []));
        } else {
            // No formatting - just add text directly
            cell = cell.child(text_content);
        }
    }

    // Add hint badge overlay when in hint mode
    if let Some((label, matches, is_unique)) = hint_label {
        let badge_bg = if is_unique {
            // Unique match - highlight strongly
            app.token(TokenKey::HintBadgeUniqueBg)
        } else if matches {
            // Matches current buffer - brighter
            app.token(TokenKey::HintBadgeMatchBg)
        } else {
            // Doesn't match - muted
            app.token(TokenKey::HintBadgeBg)
        };

        let badge_text = if is_unique {
            app.token(TokenKey::HintBadgeUniqueText)
        } else if matches {
            app.token(TokenKey::HintBadgeMatchText)
        } else {
            app.token(TokenKey::HintBadgeText)
        };

        // Only show if matches or buffer is empty (show all at start)
        if matches || app.hint_state.buffer.is_empty() {
            cell = cell.child(
                div()
                    .absolute()
                    .top_0()
                    .left_0()
                    .px(px(2.0))
                    .py(px(1.0))
                    .bg(badge_bg)
                    .rounded_sm()
                    .text_xs()
                    .font_weight(FontWeight::BOLD)
                    .text_color(badge_text)
                    .child(label)
            );
        }
    }

    // Check if this cell is in fill preview range
    let is_fill_preview = app.is_fill_preview_cell(view_row, col);

    // Apply fill preview styling (dashed-like effect via different border color)
    if is_fill_preview {
        let fill_preview_bg = app.token(TokenKey::SelectionBg).opacity(0.4);
        let fill_preview_border = app.token(TokenKey::Accent);
        cell = cell.bg(fill_preview_bg).border_1().border_color(fill_preview_border);
    }

    // Phase 6C: Invalid cell corner triangle (data validation)
    // Small red corner mark in top-right indicates validation failure
    if app.is_cell_invalid(data_row, col) {
        let invalid_color = rgb(0xE53935);  // Material Red 600
        let mark_size = 6.0 * app.metrics.zoom;
        cell = cell.child(
            div()
                .absolute()
                .top_0()
                .right_0()
                .w(px(mark_size))
                .h(px(mark_size))
                .bg(invalid_color)
        );
    }

    // Determine if we should show the fill handle on this cell
    // Excel-style: show fill handle at bottom-right corner of selection
    // - For single cell: show on active cell
    // - For range selection: show on cell at (max_row, max_col)
    // - No additional selections (Ctrl+Click multi-select)
    // - Not editing, not in hint mode, not already fill dragging
    let is_fill_handle_cell = if app.view_state.selection_end.is_some() {
        // Range selection: fill handle goes on bottom-right corner
        let ((_min_row, _min_col), (max_row, max_col)) = app.selection_range();
        view_row == max_row && col == max_col
    } else {
        // Single cell: fill handle goes on active cell
        is_active
    };
    let show_fill_handle = is_fill_handle_cell
        && !is_editing
        && app.mode != Mode::Hint
        && app.view_state.additional_selections.is_empty()
        && !app.is_fill_dragging();

    // Wrap in relative container for absolute positioning (hint badges, fill handle)
    let mut wrapper = div()
        .relative()
        .flex_shrink_0()
        .w(px(col_width))
        .h_full()
        .child(cell);

    // Add fill handle at bottom-right corner of active cell
    // Excel-style: solid opaque square with contrasting border
    if show_fill_handle {
        // Use opaque Accent color (not semi-transparent SelectionBorder)
        let handle_fill = app.token(TokenKey::Accent);
        // Use cell background for border (white on light themes, dark on dark themes)
        let handle_border = app.token(TokenKey::CellBg);
        let zoom = app.metrics.zoom;
        let visual_size = FILL_HANDLE_VISUAL_SIZE * zoom;
        let border_width = FILL_HANDLE_BORDER * zoom;
        let hit_size = FILL_HANDLE_HIT_SIZE * zoom;
        // Offset to center the hit area on the corner, with visual centered inside
        let hit_offset = hit_size / 2.0;
        let visual_offset = (hit_size - visual_size) / 2.0;

        wrapper = wrapper.child(
            div()
                .id(ElementId::Name(format!("fill-handle-{}-{}", view_row, col).into()))
                .absolute()
                .bottom(px(-hit_offset))
                .right(px(-hit_offset))
                .w(px(hit_size))
                .h(px(hit_size))
                .cursor(CursorStyle::Crosshair)
                // Mouse down on fill handle starts fill drag (preempts selection)
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _event: &MouseDownEvent, _, cx| {
                    this.start_fill_drag(cx);
                }))
                // Visual handle (smaller, centered in hit area) with Excel-style appearance
                .child(
                    div()
                        .absolute()
                        .top(px(visual_offset))
                        .left(px(visual_offset))
                        .w(px(visual_size))
                        .h(px(visual_size))
                        .bg(handle_fill)
                        .border_color(handle_border)
                        .border(px(border_width))
                )
        );
    }

    wrapper
}

/// Returns the base background color for a cell (ignoring selection state).
/// Selection is rendered as a semi-transparent overlay on top.
fn cell_base_background(
    app: &Spreadsheet,
    is_editing: bool,
    custom_bg: Option<[u8; 4]>,
) -> Hsla {
    if is_editing {
        app.token(TokenKey::EditorBg)
    } else if let Some([r, g, b, a]) = custom_bg {
        // Custom background color from cell format (RGBA → Hsla)
        Hsla::from(gpui::Rgba {
            r: r as f32 / 255.0,
            g: g as f32 / 255.0,
            b: b as f32 / 255.0,
            a: a as f32 / 255.0,
        })
    } else {
        app.token(TokenKey::CellBg)
    }
}

/// Returns the selection overlay color (semi-transparent) for layering on top of cell background.
/// For formula refs, uses the color index to pick from the rotating palette.
/// Selection ALWAYS takes precedence over formula ref highlighting.
fn selection_overlay_color(app: &Spreadsheet, is_active: bool, is_selected: bool, formula_ref_color: Option<usize>) -> Option<Hsla> {
    // Selection overlay takes precedence over formula ref (user sees their selection clearly)
    if is_active {
        // Active cell gets slightly stronger highlight
        Some(app.token(TokenKey::SelectionBg).opacity(0.5))
    } else if is_selected {
        // Selected cells get standard overlay (selection wins over formula ref)
        Some(app.token(TokenKey::SelectionBg).opacity(0.4))
    } else if let Some(color_idx) = formula_ref_color {
        // Formula reference highlight with per-ref color (semi-transparent)
        Some(ref_color_hsla(color_idx).opacity(0.25))
    } else {
        None
    }
}

/// Get the HSLA color for a formula reference by index (0-7 rotating)
fn ref_color_hsla(color_idx: usize) -> Hsla {
    use crate::app::REF_COLORS;
    let color = REF_COLORS[color_idx % 8];
    rgb(color).into()
}

fn cell_border(app: &Spreadsheet, is_editing: bool, is_active: bool, is_selected: bool, formula_ref_color: Option<usize>) -> Hsla {
    // Selection border ALWAYS wins over formula ref (user needs to see what they've selected)
    if is_editing || is_active {
        app.token(TokenKey::CellBorderFocus)
    } else if is_selected {
        app.token(TokenKey::SelectionBorder)
    } else if let Some(color_idx) = formula_ref_color {
        // Formula ref border uses per-ref color
        ref_color_hsla(color_idx)
    } else {
        app.token(TokenKey::GridLines)
    }
}

fn cell_text_color(app: &Spreadsheet, is_editing: bool, is_selected: bool, is_multi_edit_preview: bool) -> Hsla {
    if is_editing {
        app.token(TokenKey::EditorText)
    } else if is_multi_edit_preview {
        // Show preview text in a muted/dimmed color to distinguish from the active edit cell
        app.token(TokenKey::TextMuted)
    } else if is_selected {
        // Use primary text color for selected cells to ensure contrast
        // SelectionBg is semi-transparent, so text should be visible,
        // but use a slightly brighter color to ensure readability
        app.token(TokenKey::TextPrimary)
    } else {
        app.token(TokenKey::CellText)
    }
}

/// Render the popup overlay layer for autocomplete, signature help, and error banners.
/// Popups are positioned relative to the active cell rect in grid coordinates (post-scroll).
/// When editing in formula bar, popup anchors to top of grid (just below formula bar).
fn render_popup_overlay(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::app::EditorSurface;

    let (viewport_w, viewport_h) = app.viewport_rect();
    let popup_gap = 4.0;

    // Popup dimensions (capped for consistent flip behavior)
    let popup_max_width = 320.0;
    let popup_max_height = 280.0; // Cap height for flip calculation

    // Determine anchor based on active editor surface
    let (popup_x, popup_y) = match app.active_editor {
        EditorSurface::FormulaBar => {
            // Editing in formula bar: anchor to top of grid (just below formula bar)
            // X: align roughly with formula bar text area (converted to grid coords)
            // Formula bar text starts at FORMULA_BAR_TEXT_LEFT (98px), grid starts at HEADER_WIDTH (50px)
            let x = (crate::app::FORMULA_BAR_TEXT_LEFT - crate::app::HEADER_WIDTH)
                .max(0.0)
                .min((viewport_w - popup_max_width).max(0.0));
            // Y: top of grid with small gap
            let y = popup_gap;
            (x, y)
        }
        EditorSurface::Cell => {
            // Editing in cell: anchor to active cell (existing behavior)
            let cell_rect = app.active_cell_rect();

            // X position: align with cell left, clamped to viewport bounds
            let x = cell_rect.x
                .max(0.0)
                .min((viewport_w - popup_max_width).max(0.0));

            // Y position: prefer below cell, flip above if no room
            let y_below = cell_rect.bottom() + popup_gap;
            let y_above = cell_rect.y - popup_gap - popup_max_height;
            let y_raw = if y_below + popup_max_height <= viewport_h {
                y_below
            } else if y_above >= 0.0 {
                y_above
            } else {
                y_below // No room either way, just show below
            };

            // Final clamp to prevent offscreen rendering
            let y = y_raw
                .max(0.0)
                .min((viewport_h - popup_max_height).max(0.0));

            (x, y)
        }
    };

    // Get theme colors
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let accent = app.token(TokenKey::Accent);
    let error_bg = app.token(TokenKey::ErrorBg);
    let error_color = app.token(TokenKey::Error);

    div()
        .absolute()
        .inset_0()
        // Formula autocomplete popup
        .when(app.autocomplete_visible, |div| {
            let suggestions = app.autocomplete_suggestions();
            let selected = app.autocomplete_selected;
            div.child(formula_bar::render_autocomplete_popup(
                &suggestions,
                selected,
                popup_x,
                popup_y,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                selection_bg,
                cx,
            ))
        })
        // Formula signature help popup
        .when_some(app.signature_help(), |div, sig_info| {
            div.child(formula_bar::render_signature_help(
                &sig_info,
                popup_x,
                popup_y,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                accent,
            ))
        })
        // Formula error banner
        .when_some(app.formula_error(), |div, error_info| {
            div.child(formula_bar::render_error_banner(
                &error_info,
                popup_x,
                popup_y,
                error_bg,
                error_color,
                panel_border,
            ))
        })
}
