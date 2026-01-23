use gpui::*;
use gpui::StyledText;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::fill::{FILL_HANDLE_BORDER, FILL_HANDLE_HIT_SIZE, FILL_HANDLE_VISUAL_SIZE};
use crate::mode::Mode;
use crate::settings::{user_settings, Setting};
use crate::theme::TokenKey;
use super::headers::render_row_header;
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
    let scroll_row = app.scroll_row;
    let scroll_col = app.scroll_col;
    let selected = app.selected;
    let editing = app.mode.is_editing();
    let edit_value = app.edit_value.clone();
    let total_visible_rows = app.visible_rows();
    let total_visible_cols = app.visible_cols();
    let frozen_rows = app.frozen_rows;
    let frozen_cols = app.frozen_cols;

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
            .into_any_element();
    }

    // Get metrics for scaled dimensions
    let metrics = &app.metrics;

    // Freeze panes active - render 4 regions
    div()
        .flex_1()
        .overflow_hidden()
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

    // Determine border color based on cell state (spill states take precedence over normal states)
    let border_color = if has_spill_error {
        app.token(TokenKey::SpillBlockedBorder)
    } else if is_spill_parent {
        app.token(TokenKey::SpillBorder)
    } else if is_spill_receiver {
        app.token(TokenKey::SpillReceiverBorder)
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

    // Apply horizontal alignment
    // General alignment: numbers right-align, text/empty left-aligns (Excel behavior)
    cell = match format.alignment {
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
    };

    // Apply vertical alignment
    cell = match format.vertical_alignment {
        VerticalAlignment::Top => cell.items_start(),
        VerticalAlignment::Middle => cell.items_center(),
        VerticalAlignment::Bottom => cell.items_end(),
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

    // Build the text content with cursor and selection highlight
    if is_editing {
        let cursor_pos = app.edit_cursor;
        let chars: Vec<char> = value.chars().collect();
        let selection = app.edit_selection_range();

        // Build display string with cursor
        let before_cursor: String = chars.iter().take(cursor_pos).collect();
        let after_cursor: String = chars.iter().skip(cursor_pos).collect();
        let display_text: SharedString = format!("{}|{}", before_cursor, after_cursor).into();

        // Create styled text with selection highlighting
        if let Some((sel_start, sel_end)) = selection {
            // Calculate display positions accounting for cursor character '|'
            // Cursor is inserted at cursor_pos, so positions after cursor shift by 1
            let (disp_sel_start, disp_sel_end) = if cursor_pos <= sel_start {
                // Cursor before or at selection start - selection shifts right by 1
                (sel_start + 1, sel_end + 1)
            } else if cursor_pos >= sel_end {
                // Cursor at or after selection end - selection unchanged
                (sel_start, sel_end)
            } else {
                // Cursor inside selection (shouldn't happen with our selection model)
                (sel_start, sel_end + 1)
            };

            // Convert char positions to byte positions for the display string
            let display_chars: Vec<char> = display_text.chars().collect();
            let byte_sel_start = display_chars.iter().take(disp_sel_start).collect::<String>().len();
            let byte_sel_end = display_chars.iter().take(disp_sel_end).collect::<String>().len();
            let total_bytes = display_text.len();

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

            cell = cell.child(StyledText::new(display_text).with_runs(runs));
        } else {
            // No selection - plain text with cursor
            cell = cell.child(display_text);
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

    // Determine if we should show the fill handle on this cell
    // Excel-style: show fill handle at bottom-right corner of selection
    // - For single cell: show on active cell
    // - For range selection: show on cell at (max_row, max_col)
    // - No additional selections (Ctrl+Click multi-select)
    // - Not editing, not in hint mode, not already fill dragging
    let is_fill_handle_cell = if app.selection_end.is_some() {
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
        && app.additional_selections.is_empty()
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
