use gpui::*;
use gpui::StyledText;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, REF_COLORS};
use crate::fill::{FILL_HANDLE_BORDER, FILL_HANDLE_HIT_SIZE, FILL_HANDLE_VISUAL_SIZE, FILL_HANDLE_HOVER_GLOW, FILL_HANDLE_INWARD_OVERLAP};
use crate::formula_refs::RefKey;
use crate::mode::Mode;
use crate::settings::{user_settings, Setting};
use crate::split_view::SplitSide;
use crate::theme::TokenKey;
use crate::trace::TraceRole;
use crate::workbook_view::WorkbookViewState;
use super::headers::render_row_header;
use super::formula_bar;
use visigrid_engine::cell::{Alignment, VerticalAlignment};
use visigrid_engine::formula::eval::Value;

/// Create a non-interactive overlay div (absolute-positioned, full cell coverage).
///
/// In gpui, elements without `.id()` do not participate in hit-testing and cannot
/// intercept mouse events. This helper enforces that contract: callers style the
/// returned `Div` but must **never** add `.id()` to it, or it will steal clicks
/// from the cell underneath.
///
/// Used for: user border overlays, trace path overlays, and any visual-only layer
/// that must not interfere with selection, drag, or editing.
fn non_interactive_overlay() -> Div {
    div()
        .absolute()
        .inset_0()
}

/// Get the view state for a specific pane.
/// - `None` or `Some(SplitSide::Left)` → main view_state
/// - `Some(SplitSide::Right)` → split_pane.view_state (falls back to main if no split)
fn get_pane_view_state(app: &Spreadsheet, pane_side: Option<SplitSide>) -> &WorkbookViewState {
    match pane_side {
        None | Some(SplitSide::Left) => &app.view_state,
        Some(SplitSide::Right) => app.split_pane.as_ref()
            .map(|p| &p.view_state)
            .unwrap_or(&app.view_state),
    }
}

/// Check if a cell is selected based on the pane's view state.
fn is_selected_in_pane(
    view_state: &WorkbookViewState,
    row: usize,
    col: usize,
) -> bool {
    view_state.is_selected(row, col)
}

/// Get selection borders for a cell based on pane view state.
/// Returns (top, right, bottom, left) indicating which edges need borders.
fn selection_borders_for_pane(
    view_state: &WorkbookViewState,
    row: usize,
    col: usize,
) -> (bool, bool, bool, bool) {
    // Check if this cell is selected
    if !view_state.is_selected(row, col) {
        return (false, false, false, false);
    }

    // Only draw borders on selection edges (not interior cell boundaries)
    let above_selected = row > 0 && view_state.is_selected(row - 1, col);
    let below_selected = view_state.is_selected(row + 1, col);
    let left_selected = col > 0 && view_state.is_selected(row, col - 1);
    let right_selected = view_state.is_selected(row, col + 1);

    (
        !above_selected, // top: draw if cell above is NOT selected
        !right_selected, // right: draw if cell to right is NOT selected
        !below_selected, // bottom: draw if cell below is NOT selected
        !left_selected,  // left: draw if cell to left is NOT selected
    )
}

/// Get selection range from view state
fn selection_range_for_pane(view_state: &WorkbookViewState) -> ((usize, usize), (usize, usize)) {
    let (r1, c1) = view_state.selected;
    let (r2, c2) = view_state.selection_end.unwrap_or(view_state.selected);
    let min_row = r1.min(r2);
    let max_row = r1.max(r2);
    let min_col = c1.min(c2);
    let max_col = c1.max(c2);
    ((min_row, min_col), (max_row, max_col))
}

/// Render the main cell grid with freeze pane support
///
/// When frozen_rows > 0 or frozen_cols > 0, renders 4 regions:
/// 1. Frozen corner (top-left, never scrolls)
/// 2. Frozen rows (top, scrolls horizontally only)
/// 3. Frozen cols (left, scrolls vertically only)
/// 4. Main grid (scrolls both directions)
///
/// The `pane_side` parameter specifies which split pane is being rendered:
/// - `None` = no split view, use main view_state
/// - `Some(SplitSide::Left)` = left pane, use main view_state
/// - `Some(SplitSide::Right)` = right pane, use split_pane.view_state
pub fn render_grid(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
    pane_side: Option<SplitSide>,
) -> impl IntoElement {
    // Verify cached_sheet_id is in sync (debug builds only)
    // This catches desync bugs early - if this fires, a code path changed
    // the active sheet without calling update_cached_sheet_id().
    app.debug_assert_sheet_cache_sync(cx);

    // Get view state based on which pane we're rendering
    let view_state = get_pane_view_state(app, pane_side);
    let scroll_row = view_state.scroll_row;
    let scroll_col = view_state.scroll_col;
    let frozen_rows = view_state.frozen_rows;
    let frozen_cols = view_state.frozen_cols;

    let editing = app.mode.is_editing();
    let edit_value = app.edit_value.clone();
    let total_visible_rows = app.visible_rows();
    let total_visible_cols = app.visible_cols();

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
                            let (view_row, data_row) = app.nth_visible_row(visible_index, cx)?;
                            let is_last_visible_row = screen_row == total_visible_rows - 1;
                            Some(render_row(
                                view_row,
                                data_row,
                                scroll_col,
                                total_visible_cols,
                                view_state,
                                pane_side,
                                editing,
                                &edit_value,
                                show_gridlines,
                                is_last_visible_row,
                                app,
                                window,
                                cx,
                            ))
                        })
                    )
            )
            // Merge overlays - unified merged cells drawn above cell grid
            .child(render_merge_overlays(app, window, cx, pane_side))
            // Text spill overlay - draws overflowing text above cell backgrounds
            .child(render_text_spill_overlay(app, window, cx, pane_side))
            // Dashed borders overlay for formula references (above cells, below popups)
            .child(render_formula_ref_borders(app, pane_side))
            // Popup overlay layer - positioned relative to grid, not window chrome
            .child(render_popup_overlay(app, cx))
            // Debug: draw 1px reference lines to verify pixel alignment (Cmd+Alt+Shift+G).
            // Red line at x=0 (grid origin), green line at first cell boundary.
            // If these shimmer while scrolling, the origin or cell widths are fractional.
            .when(app.debug_grid_alignment, |d| {
                let first_col_w = app.metrics.col_width(app.col_width(scroll_col));
                d.child(
                    div().absolute().left_0().top_0().w(px(1.0)).h_full()
                        .bg(gpui::rgb(0xff0000))
                ).child(
                    div().absolute().left(px(first_col_w + app.metrics.header_w)).top_0().w(px(1.0)).h_full()
                        .bg(gpui::rgb(0x00ff00))
                )
            })
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
                            let data_row = app.view_to_data(view_row, cx);
                            // Use scaled row height for rendering
                            let row_height = metrics.row_height(app.row_height(view_row));
                            div()
                                .flex()
                                .flex_shrink_0()
                                .h(px(row_height))
                                // Row header for frozen row
                                .child(render_row_header(app, view_row, cx))
                                // Frozen corner cells (cols 0..frozen_cols)
                                // No viewport boundary edges — dividers separate from scrollable regions
                                .when(frozen_cols > 0, |d| {
                                    d.children(
                                        (0..frozen_cols).map(|col| {
                                            let col_width = metrics.col_width(app.col_width(col));
                                            render_cell(view_row, data_row, col, col_width, row_height, view_state, pane_side, editing, &edit_value, show_gridlines, false, false, app, window, cx)
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
                                // Right boundary at viewport edge; no bottom boundary (divider below)
                                .children(
                                    (0..scrollable_visible_cols).map(|visible_col| {
                                        let col = scroll_col + visible_col;
                                        let col_width = metrics.col_width(app.col_width(col));
                                        let is_last_col = visible_col == scrollable_visible_cols - 1;
                                        render_cell(view_row, data_row, col, col_width, row_height, view_state, pane_side, editing, &edit_value, show_gridlines, false, is_last_col, app, window, cx)
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
                                let (view_row, data_row) = app.nth_visible_row(visible_index, cx)?;
                                let row_height = metrics.row_height(app.row_height(view_row));
                                let is_last_row = screen_row == scrollable_visible_rows - 1;
                                Some(div()
                                    .flex()
                                    .flex_shrink_0()
                                    .h(px(row_height))
                                    // Row header
                                    .child(render_row_header(app, view_row, cx))
                                    // Frozen column cells (cols 0..frozen_cols)
                                    // Bottom boundary at viewport edge; no right boundary (divider separates)
                                    .when(frozen_cols > 0, |d| {
                                        d.children(
                                            (0..frozen_cols).map(|col| {
                                                let col_width = metrics.col_width(app.col_width(col));
                                                render_cell(view_row, data_row, col, col_width, row_height, view_state, pane_side, editing, &edit_value, show_gridlines, is_last_row, false, app, window, cx)
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
                                    // Main grid cells — both bottom and right boundary at viewport edge
                                    .children(
                                        (0..scrollable_visible_cols).map(|visible_col| {
                                            let col = scroll_col + visible_col;
                                            let col_width = metrics.col_width(app.col_width(col));
                                            let is_last_col = visible_col == scrollable_visible_cols - 1;
                                            render_cell(view_row, data_row, col, col_width, row_height, view_state, pane_side, editing, &edit_value, show_gridlines, is_last_row, is_last_col, app, window, cx)
                                        })
                                    ))
                            })
                        )
                )
        )
        // Merge overlays - unified merged cells drawn above cell grid
        .child(render_merge_overlays(app, window, cx, pane_side))
        // Text spill overlay - draws overflowing text above cell backgrounds
        .child(render_text_spill_overlay(app, window, cx, pane_side))
        // Dashed borders overlay for formula references (above cells, below popups)
        .child(render_formula_ref_borders(app, pane_side))
        // Popup overlay layer - positioned relative to grid, not window chrome
        .child(render_popup_overlay(app, cx))
        .into_any_element()
}

fn render_row(
    view_row: usize,
    data_row: usize,
    scroll_col: usize,
    visible_cols: usize,
    view_state: &WorkbookViewState,
    pane_side: Option<SplitSide>,
    editing: bool,
    edit_value: &str,
    show_gridlines: bool,
    is_last_visible_row: bool,
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
                let is_last_visible_col = visible_col == visible_cols - 1;
                // view_row for selection/display, data_row for cell data access
                render_cell(view_row, data_row, col, col_width, row_height, view_state, pane_side, editing, edit_value, show_gridlines, is_last_visible_row, is_last_visible_col, app, window, cx)
            })
        )
}

fn render_cell(
    view_row: usize,
    data_row: usize,
    col: usize,
    col_width: f32,
    _row_height: f32,
    view_state: &WorkbookViewState,
    pane_side: Option<SplitSide>,
    editing: bool,
    edit_value: &str,
    show_gridlines: bool,
    is_last_visible_row: bool,
    is_last_visible_col: bool,
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> AnyElement {
    // Selection uses view_row (what user sees/clicks)
    // Use pane-specific view state for selection checks
    let is_selected = is_selected_in_pane(view_state, view_row, col);
    let is_active = view_state.selected == (view_row, col);
    let is_editing = editing && is_active;
    let is_formula_ref = app.is_formula_ref(view_row, col);
    let formula_ref_color = app.formula_ref_color(view_row, col);  // Color index for multi-color refs
    let is_active_ref_target = app.is_active_ref_target(view_row, col);  // Live ref navigation target
    let is_inspector_hover = app.inspector_hover_cell == Some((view_row, col));  // Hover highlight from inspector

    // Check if cell is in trace path (Phase 3.5b)
    let sheet_id = app.sheet(cx).id;
    let trace_position = app.inspector_trace_path.as_ref().and_then(|path| {
        path.iter().position(|cell| {
            cell.sheet == sheet_id && cell.row == view_row && cell.col == col
        }).map(|pos| {
            let is_start = pos == 0;
            let is_end = pos == path.len() - 1;
            (is_start, is_end)
        })
    });

    // Dependency trace highlighting (Alt+T)
    let trace_role = app.trace_role(sheet_id, view_row, col);

    // Check if cell is in history highlight range (Phase 7A)
    let is_history_highlight = app.history_highlight_range.map_or(false, |(sheet_idx, sr, sc, er, ec)| {
        sheet_idx == app.sheet_index(cx)
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

    // Merged cell: hidden cells suppress text and block editing
    let is_merge_hidden = app.sheet(cx).is_merge_hidden(data_row, col);
    let is_merge_origin = app.sheet(cx).is_merge_origin(data_row, col);

    // Merge spacer: overlay handles all rendering for merged cells.
    // Hidden cells always become spacers. Origin cells become spacers unless editing.
    let cell_row = view_row;
    let cell_col = col;
    if is_merge_hidden || (is_merge_origin && !(editing && is_active)) {
        return div()
            .id(ElementId::Name(format!("cell-{}-{}", view_row, col).into()))
            .relative()
            .flex_shrink_0()
            .w(px(col_width))
            .h_full()
            // Keep mouse_move for edge-case drag-through (overlay handles clicks on top)
            .on_mouse_move(cx.listener(move |this, _event: &MouseMoveEvent, _, cx| {
                if this.is_fill_dragging() {
                    this.continue_fill_drag(cell_row, cell_col, cx);
                    return;
                }
                if this.dragging_selection {
                    if this.mode.is_formula() {
                        this.formula_continue_drag(cell_row, cell_col, cx);
                    } else {
                        this.continue_drag_selection(cell_row, cell_col, cx);
                    }
                }
            }))
            .into_any_element();
    }

    // Spill state detection - uses data_row (storage)
    let is_spill_parent = app.sheet(cx).is_spill_parent(data_row, col);
    let is_spill_receiver = app.sheet(cx).is_spill_receiver(data_row, col);
    let has_spill_error = app.sheet(cx).has_spill_error(data_row, col);

    // Check for multi-edit preview (shows what each selected cell will receive)
    let multi_edit_preview = app.multi_edit_preview(view_row, col);
    let is_multi_edit_preview = multi_edit_preview.is_some();

    // Cell value: use data_row to access actual storage
    // Merge-hidden cells display nothing (text is shown only at the origin)
    let value = if is_merge_hidden {
        String::new()
    } else if is_editing {
        edit_value.to_string()
    } else if let Some(preview) = multi_edit_preview {
        // Show the preview value for cells in multi-selection during editing
        preview
    } else if app.show_formulas() {
        app.sheet(cx).get_raw(data_row, col)
    } else {
        let display = app.sheet(cx).get_formatted_display(data_row, col);
        // Hide zero values if show_zeros is false
        if !app.show_zeros() && display == "0" {
            String::new()
        } else {
            display
        }
    };

    let format = app.sheet(cx).get_format(data_row, col);

    // NOTE: Text spillover is rendered in a separate overlay pass (render_text_spill_overlay)
    // to avoid z-order issues where adjacent cells' backgrounds cover the spilled text.
    // Cells always clip their own content; spill is drawn on top of all cell backgrounds.

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
        .relative()  // Enable absolute positioning for selection overlay
        .size_full()
        .flex()
        .px_1()
        .overflow_hidden()  // Always clip; spill is rendered in overlay layer
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

    // Dependency trace highlight (Alt+T - shows precedents and dependents)
    // Only show for non-selected, non-active cells to avoid visual clutter
    if !is_selected && !is_active {
        match trace_role {
            TraceRole::Source => {
                // Source cell: strong accent border (selection handles the fill)
                let source_border = app.token(TokenKey::TraceSourceBorder);
                cell = cell.child(
                    div()
                        .absolute()
                        .inset_0()
                        .border_2()
                        .border_color(source_border)
                );
            }
            TraceRole::Precedent => {
                // Precedent (input): themed tint
                let precedent_bg = app.token(TokenKey::TracePrecedentBg);
                cell = cell.child(
                    div()
                        .absolute()
                        .inset_0()
                        .bg(precedent_bg)
                        .border_1()
                        .border_color(precedent_bg.opacity(0.6))
                );
            }
            TraceRole::Dependent => {
                // Dependent (output): themed tint
                let dependent_bg = app.token(TokenKey::TraceDependentBg);
                cell = cell.child(
                    div()
                        .absolute()
                        .inset_0()
                        .bg(dependent_bg)
                        .border_1()
                        .border_color(dependent_bg.opacity(0.6))
                );
            }
            TraceRole::None => {}
        }
    }

    // CenterAcrossSelection span computation:
    // If this cell has text + CenterAcrossSelection, compute the total span width.
    // If this cell is empty + CenterAcrossSelection and a cell to its left spans across it, suppress text.
    // Priority: merge rules win — merged cells never use CenterAcrossSelection.
    let is_in_merge = app.sheet(cx).get_merge(data_row, col).is_some();
    let center_across_span = if !is_editing && !is_in_merge && format.alignment == Alignment::CenterAcrossSelection {
        if !value.is_empty() {
            // Source cell: scan right for empty cells with CenterAcrossSelection
            Some(center_across_span_width(data_row, col, col_width, app, cx))
        } else {
            // Empty cell: check if a source cell to the left spans across us → suppress
            if is_center_across_continuation(data_row, col, app, cx) {
                Some(0.0) // sentinel: continuation cell, suppress text
            } else {
                None // isolated empty CenterAcross cell, render normally
            }
        }
    } else {
        None
    };

    // Apply horizontal alignment
    // When editing, always left-align so caret positioning works correctly
    // General alignment: numbers right-align, text/empty left-aligns (Excel behavior)
    cell = if is_editing {
        // Editing: always left-align for correct caret positioning
        cell.justify_start()
    } else {
        match format.alignment {
            Alignment::General => {
                let computed = app.sheet(cx).get_computed_value(data_row, col);
                match computed {
                    Value::Number(_) => cell.justify_end(),
                    _ => cell.justify_start(),
                }
            }
            Alignment::Left => cell.justify_start(),
            Alignment::Center => cell.justify_center(),
            // CenterAcrossSelection: text is positioned by absolute overlay, not flex alignment
            Alignment::CenterAcrossSelection => cell.justify_start(),
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
        // Selected cells: selection border on outer edges (accent blue)
        // Use pane-specific view state for selection border calculation
        let (top, right, bottom, left) = selection_borders_for_pane(view_state, view_row, col);
        let mut c = cell;
        if top { c = c.border_t_1(); }
        if right { c = c.border_r_1(); }
        if bottom { c = c.border_b_1(); }
        if left { c = c.border_l_1(); }

        // Interior gridlines (GridLines color via overlay child, since cell border_color
        // is already SelectionBorder for outer edges)
        if show_gridlines {
            let (top_in_merge, left_in_merge, _, _) = if let Some(merge) = app.sheet(cx).get_merge(data_row, col) {
                (data_row > merge.start.0, col > merge.start.1, data_row < merge.end.0, col < merge.end.1)
            } else {
                (false, false, false, false)
            };
            // Canonical top+left ownership (same rule as normal cells).
            // Skip edges where selection border is already drawn (!top / !left).
            let need_top = !top && data_row > 0 && !top_in_merge;
            let need_left = !left && col > 0 && !left_in_merge;
            if need_top || need_left {
                c = c.child(
                    non_interactive_overlay()
                        .border_color(app.token(TokenKey::GridLines))
                        .when(need_top, |d| d.border_t_1())
                        .when(need_left, |d| d.border_l_1())
                );
            }
        }
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
        // 1px border for spill receivers
        cell.border_1()
    } else {
        // Formula ref borders are drawn as dashed overlays (render_formula_ref_borders)
        // so formula ref cells fall through here for normal gridline/user border handling
        let sheet_has_borders = app.sheet(cx).has_any_borders;

        // Fast path: no gridlines and no borders on this sheet → skip all Tier 5 work
        if !show_gridlines && !sheet_has_borders {
            cell
        } else {
            // Resolve user borders only when the sheet actually has borders (avoids
            // calling cell_user_borders() on every visible cell every frame for sheets
            // that have never had border formatting).
            let (user_top, user_right, user_bottom, user_left) = if sheet_has_borders {
                app.cell_user_borders(
                    data_row, col, cx, is_last_visible_row, is_last_visible_col,
                )
            } else {
                (false, false, false, false)
            };
            let has_user_border = user_top || user_right || user_bottom || user_left;

            // Suppress interior gridlines within merged regions (computed once, used by gridlines)
            let (top_in_merge, left_in_merge, bottom_in_merge, right_in_merge) = if let Some(merge) = app.sheet(cx).get_merge(data_row, col) {
                (data_row > merge.start.0, col > merge.start.1, data_row < merge.end.0, col < merge.end.1)
            } else {
                (false, false, false, false)
            };

            // Gridlines: draw on edges WITHOUT resolved user borders.
            // Border suppression order per edge: user border → merge interior → boundary.
            let mut c = cell;
            #[cfg(debug_assertions)]
            if show_gridlines {
                app.debug_gridline_cells.set(app.debug_gridline_cells.get() + 1);
            }
            if show_gridlines {
                if !user_top && data_row > 0 && !top_in_merge { c = c.border_t_1(); }
                if !user_left && col > 0 && !left_in_merge { c = c.border_l_1(); }
                if !user_bottom && is_last_visible_row && !bottom_in_merge { c = c.border_b_1(); }
                if !user_right && is_last_visible_col && !right_in_merge { c = c.border_r_1(); }
            }

            // User borders: non-interactive overlay so explicit borders render on top of gridlines.
            // Uses non_interactive_overlay() — must never have .id() or it will steal cell events.
            #[cfg(debug_assertions)]
            if has_user_border {
                app.debug_userborder_cells.set(app.debug_userborder_cells.get() + 1);
            }
            if has_user_border {
                let border_color = app.token(TokenKey::UserBorder);
                c = c.child(
                    non_interactive_overlay()
                        .border_color(border_color)
                        .when(user_top, |d| d.border_t_1())
                        .when(user_right, |d| d.border_r_1())
                        .when(user_bottom, |d| d.border_b_1())
                        .when(user_left, |d| d.border_l_1())
                );
            }
            c
        }
    };

    cell = cell
        .text_color(cell_text_color(app, is_editing, is_selected, is_multi_edit_preview))
        .text_size(px(app.metrics.font_size));  // Scaled font size for zoom

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
        // Text spillover is handled by render_text_spill_overlay() in a separate pass.
        // If this cell will be rendered by the spill overlay, skip text here to avoid double-draw.
        let use_spill_overlay = should_use_spill_overlay(
            data_row,
            col,
            &value,
            col_width,
            format.alignment,
            is_selected,
            is_active,
            is_editing,
            app,
            window,
            cx,
        );

        // CenterAcrossSelection: continuation cells suppress text entirely
        let suppress_text = matches!(center_across_span, Some(w) if w == 0.0);

        if !use_spill_overlay && !suppress_text {
            let text_content: SharedString = value.clone().into();

            // Check if any formatting is applied (font_family IS formatting)
            let has_formatting = format.bold
                || format.italic
                || format.underline
                || format.strikethrough
                || format.font_family.is_some()
                || format.font_size.is_some()
                || format.font_color.is_some();

            // Build styled text element if needed
            let text_element: AnyElement = if has_formatting {
                // Build text style with ALL formatting properties
                let mut text_style = window.text_style();
                text_style.color = cell_text_color(app, is_editing, is_selected, is_multi_edit_preview);

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
                if format.strikethrough {
                    text_style.strikethrough = Some(StrikethroughStyle {
                        thickness: px(1.),
                        ..Default::default()
                    });
                }
                // Font family must be on text_style for StyledText to use it
                if let Some(ref family) = format.font_family {
                    text_style.font_family = family.clone().into();
                }
                // Font color: only apply for non-editing/non-selected cells
                if let Some(rgba) = format.font_color {
                    if !is_editing && !is_selected && !is_multi_edit_preview {
                        text_style.color = gpui::Hsla::from(gpui::Rgba {
                            r: rgba[0] as f32 / 255.0,
                            g: rgba[1] as f32 / 255.0,
                            b: rgba[2] as f32 / 255.0,
                            a: rgba[3] as f32 / 255.0,
                        });
                    }
                }

                let styled = StyledText::new(text_content).with_default_highlights(&text_style, []);
                // Font size: TextRun doesn't carry font_size, so we must set it
                // on a parent div to cascade via the element tree's text style.
                if let Some(size) = format.font_size {
                    div()
                        .text_size(px(size * app.metrics.zoom))
                        .child(styled)
                        .into_any_element()
                } else {
                    styled.into_any_element()
                }
            } else {
                text_content.into_any_element()
            };

            // CenterAcrossSelection source cell: render text centered across the full span
            if let Some(total_width) = center_across_span {
                if total_width > 0.0 {
                    // Absolute overlay that extends across multiple columns, centered
                    cell = cell.child(
                        div()
                            .absolute()
                            .top_0()
                            .bottom_0()
                            .left_0()
                            .w(px(total_width))
                            .flex()
                            .items_center()
                            .justify_center()
                            .overflow_hidden()
                            .child(text_element)
                    );
                } else {
                    cell = cell.child(text_element);
                }
            } else {
                cell = cell.child(text_element);
            }

        }
        // If use_spill_overlay is true, text is rendered by render_text_spill_overlay() instead
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

    // Apply fill preview styling: border-only (no fill) for clear action feedback
    // Uses darker/more opaque border than selection to signal "action in progress"
    if is_fill_preview {
        let fill_preview_border = app.token(TokenKey::Accent);
        cell = cell.border_1().border_color(fill_preview_border);
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
    // - Not editing, not in hint mode
    // - During drag: keep visible at source_end to anchor spatial understanding
    // Use pane-specific view state for selection info
    let is_fill_handle_cell = if app.is_fill_dragging() {
        // During drag: show at source_end (bottom-right of source range)
        app.fill_drag_source_end() == Some((view_row, col))
    } else if view_state.selection_end.is_some() {
        // Range selection: fill handle goes on bottom-right corner
        let ((_min_row, _min_col), (max_row, max_col)) = selection_range_for_pane(view_state);
        view_row == max_row && col == max_col
    } else {
        // Single cell: fill handle goes on active cell
        is_active
    };
    let show_fill_handle = is_fill_handle_cell
        && !is_editing
        && app.mode != Mode::Hint
        && view_state.additional_selections.is_empty();

    // Cell wrapper owns .id() and all mouse handlers. This ensures clicks on sub-pixel
    // border gaps between cells still register — no dead zones. The wrapper has exact
    // pixel dimensions (col_width × row_height) so adjacent wrappers in the flex row
    // are packed with no gaps, eliminating gridline click failures.
    let mut wrapper = div()
        .id(ElementId::Name(format!("cell-{}-{}", view_row, col).into()))
        .relative()
        .flex_shrink_0()
        .w(px(col_width))
        .h_full()
        // TODO: Replace with custom thick-plus cursor if gpui adds custom cursor image support
        .cursor(CursorStyle::Crosshair) // Grid interior = crosshair (Excel convention)
        .child(cell)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, window, cx| {
            // Don't handle grid clicks when format bar owns focus
            if this.ui.format_bar.is_active(window) {
                // Commit any pending font size edit, then close
                crate::views::format_bar::commit_font_size(this, cx);
                this.ui.format_bar.size_dropdown = false;
                cx.notify();
                return;
            }
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

            // Activate this pane if we're in split view
            this.activate_pane(pane_side, cx);

            // If clicking a spill receiver, redirect to the spill parent
            let (target_row, target_col) = if let Some((parent_row, parent_col)) = this.sheet(cx).get_spill_parent(cell_row, cell_col) {
                (parent_row, parent_col)
            } else {
                (cell_row, cell_col)
            };

            // If clicking a merged cell, redirect to merge origin and note merge extent
            let (target_row, target_col, merge_end) = if let Some(merge) = this.sheet(cx).get_merge(target_row, target_col) {
                (merge.start.0, merge.start.1, Some(merge.end))
            } else {
                (target_row, target_col, None)
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

            // Format Painter mode: apply captured format to clicked cell
            if this.mode == crate::mode::Mode::FormatPainter {
                this.select_cell(target_row, target_col, false, cx);
                this.apply_format_painter(cx);
                return;
            }

            // If editing, commit the current edit before navigating to another cell
            if this.mode.is_editing() {
                let active = this.active_view_state().selected;
                if (target_row, target_col) != active {
                    this.commit_pending_edit(cx);
                }
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

            // If the click target is a multi-cell merge, expand selection to cover full region
            if let Some(merge_br) = merge_end {
                if merge_br != (target_row, target_col) {
                    let (anchor_row, anchor_col) = this.active_view_state().selected;
                    // Pick the merge corner that, combined with anchor, covers the full region
                    let far_row = if anchor_row <= target_row { merge_br.0 } else { target_row };
                    let far_col = if anchor_col <= target_col { merge_br.1 } else { target_col };
                    this.active_view_state_mut().selection_end = Some((far_row, far_col));
                    cx.notify();
                }
            }
        }))
        .on_mouse_down(MouseButton::Right, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Don't handle right-clicks if modal/overlay is visible
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            // If right-clicking outside current selection, move active cell there
            if !this.is_selected(cell_row, cell_col) {
                this.select_cell(cell_row, cell_col, false, cx);
            }
            this.show_context_menu(
                crate::app::ContextMenuKind::Cell,
                event.position,
                cx,
            );
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
        .on_mouse_up(MouseButton::Left, cx.listener(move |this, event: &MouseUpEvent, _, cx| {
            // Don't handle if modal/overlay is visible
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            // End fill handle drag if active (commits the fill)
            // Ctrl modifier toggles copy/series behavior
            if this.is_fill_dragging() {
                let ctrl_held = event.modifiers.control || event.modifiers.platform;
                this.end_fill_drag(ctrl_held, cx);
                return;
            }
            // End drag selection (works for both normal and formula mode)
            this.end_drag_selection(cx);
        }));

    // Add fill handle at bottom-right corner of active cell
    // Excel-style: solid dark square that overlaps selection border (corner cap feel)
    if show_fill_handle {
        // Solid dark fill - darker than selection border for contrast
        // Use selection border color darkened, or fall back to accent
        let handle_fill = app.token(TokenKey::SelectionBorder);
        // Use cell background for border (white on light themes, dark on dark themes)
        let handle_border = app.token(TokenKey::CellBg);
        // Hover glow: subtle accent halo
        let hover_glow = app.token(TokenKey::Accent).opacity(0.25);
        let zoom = app.metrics.zoom;
        let visual_size = FILL_HANDLE_VISUAL_SIZE * zoom;
        let border_width = FILL_HANDLE_BORDER * zoom;
        let hit_size = FILL_HANDLE_HIT_SIZE * zoom;
        let glow_size = FILL_HANDLE_HOVER_GLOW * zoom;
        let inward_overlap = FILL_HANDLE_INWARD_OVERLAP * zoom;

        // Position hit area so visual handle overlaps inward by 1px
        // hit_offset positions the hit area center relative to cell corner
        // Reduce offset to move handle inward
        let hit_offset = (hit_size / 2.0) - inward_overlap;
        let visual_offset = (hit_size - visual_size) / 2.0;
        let glow_offset = visual_offset - glow_size;
        let glow_total_size = visual_size + glow_size * 2.0;

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
                // Hover glow layer (invisible until hover)
                .child(
                    div()
                        .absolute()
                        .top(px(glow_offset))
                        .left(px(glow_offset))
                        .w(px(glow_total_size))
                        .h(px(glow_total_size))
                        .rounded(px(2.0))
                        .hover(move |style| style.bg(hover_glow))
                )
                // Visual handle: solid, dark, overlaps border like a corner cap
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

    wrapper.into_any_element()
}

/// Resolve alignment for rendering, applying Excel-style General rules.
/// This MUST be used by both normal cell rendering and spill overlay to ensure consistency.
///
/// Excel rules:
/// - Explicit Left/Center/Right: use as-is
/// - General: text → Left, numbers/dates → Right
#[inline]
fn resolved_alignment(alignment: Alignment, is_number: bool) -> Alignment {
    match alignment {
        Alignment::General => {
            if is_number {
                Alignment::Right
            } else {
                Alignment::Left
            }
        }
        explicit => explicit,
    }
}

/// Check if a cell's alignment allows text spillover.
/// Only left-aligned text spills rightward. Center and right-aligned text does not spill.
/// General alignment spills only for non-numeric values (text left-aligns, numbers right-align).
#[inline]
fn should_alignment_spill(alignment: Alignment, is_number: bool) -> bool {
    match resolved_alignment(alignment, is_number) {
        Alignment::Left => true,
        Alignment::Center | Alignment::Right | Alignment::General | Alignment::CenterAcrossSelection => false,
    }
}

/// Compute the total width of a CenterAcrossSelection span starting at (row, col).
/// Scans rightward while adjacent cells are empty AND have CenterAcrossSelection alignment.
/// Returns the total span width in pixels (starting from this cell's width).
fn center_across_span_width(
    row: usize,
    col: usize,
    col_width: f32,
    app: &Spreadsheet,
    cx: &App,
) -> f32 {
    let mut total_width = col_width;
    let max_col = app.sheet(cx).cols.min(col + 50); // reasonable scan limit
    let mut check_col = col + 1;
    while check_col < max_col {
        // Stop at merged cells — merge rules take priority over CenterAcross
        if app.sheet(cx).get_merge(row, check_col).is_some() {
            break;
        }
        let adj_format = app.sheet(cx).get_format(row, check_col);
        if adj_format.alignment != Alignment::CenterAcrossSelection {
            break;
        }
        let adj_display = app.sheet(cx).get_formatted_display(row, check_col);
        if !adj_display.is_empty() {
            break; // adjacent cell has content — stop the span
        }
        let adj_width = app.metrics.col_width(app.col_width(check_col));
        total_width += adj_width;
        check_col += 1;
    }
    total_width
}

/// Check if an empty CenterAcrossSelection cell is a "continuation" of a span from the left.
/// Scans leftward for a non-empty cell with CenterAcrossSelection whose span reaches this column.
fn is_center_across_continuation(
    row: usize,
    col: usize,
    app: &Spreadsheet,
    cx: &App,
) -> bool {
    if col == 0 {
        return false;
    }
    let mut check_col = col;
    while check_col > 0 {
        check_col -= 1;
        // Stop at merged cells — merge rules take priority over CenterAcross
        if app.sheet(cx).get_merge(row, check_col).is_some() {
            return false;
        }
        let fmt = app.sheet(cx).get_format(row, check_col);
        if fmt.alignment != Alignment::CenterAcrossSelection {
            return false; // hit a non-CenterAcross cell — no span from the left
        }
        let display = app.sheet(cx).get_formatted_display(row, check_col);
        if !display.is_empty() {
            // Found a source cell. Check if its span reaches our column
            // by re-scanning rightward from the source.
            let mut scan_col = check_col + 1;
            let max_col = app.sheet(cx).cols.min(check_col + 50);
            while scan_col < max_col && scan_col <= col {
                // Stop at merged cells
                if app.sheet(cx).get_merge(row, scan_col).is_some() {
                    break;
                }
                let sf = app.sheet(cx).get_format(row, scan_col);
                if sf.alignment != Alignment::CenterAcrossSelection {
                    break;
                }
                let sd = app.sheet(cx).get_formatted_display(row, scan_col);
                if !sd.is_empty() {
                    break;
                }
                scan_col += 1;
            }
            return scan_col > col;
        }
        // Empty CenterAcross cell — keep scanning left
    }
    false
}

/// Calculate how much a cell's text should spill into adjacent cells (Excel-style overflow).
/// Returns Some(extra_pixels) if text should spill, None otherwise.
/// Only left-aligned text (or General alignment for text values) spills rightward.
///
/// # Spill Behavior (Excel-compatible)
/// - Text spills rightward into adjacent empty cells only
/// - Spill stops at the first non-empty cell
/// - Selected/active cells don't spill (handled by caller)
/// - Editing cells don't spill (handled by caller)
/// - Numbers never spill (they show #### if too wide, but we don't implement that yet)
fn calculate_text_spill(
    row: usize,
    col: usize,
    text: &str,
    cell_width: f32,
    alignment: Alignment,
    app: &Spreadsheet,
    window: &Window,
    cx: &App,
) -> Option<f32> {
    // Check if alignment allows spilling
    let is_number = matches!(app.sheet(cx).get_computed_value(row, col), Value::Number(_));
    if !should_alignment_spill(alignment, is_number) {
        return None;
    }

    // Shape text to get its pixel width
    let text_owned = text.to_string();
    let text_shared: SharedString = text_owned.into();
    let text_len = text_shared.len();
    if text_len == 0 {
        return None;
    }

    let shaped = window.text_system().shape_line(
        text_shared,
        px(app.metrics.font_size),
        &[TextRun {
            len: text_len,
            font: Font::default(),
            color: Hsla::default(),
            background_color: None,
            underline: None,
            strikethrough: None,
        }],
        None,
    );

    let text_width: f32 = shaped.width.into();
    let padding = 8.0; // px_1 = 4px each side
    let available_width = cell_width - padding;

    if text_width <= available_width {
        return None; // Text fits, no spill needed
    }

    // Calculate how many columns we can spill into
    let overflow_needed = text_width - available_width;
    let mut spill_width: f32 = 0.0;
    let mut check_col = col + 1;
    let max_col = col + 10; // Limit spillover to prevent runaway

    while spill_width < overflow_needed && check_col < max_col {
        // Check if adjacent cell is empty
        let adjacent_display = app.sheet(cx).get_formatted_display(row, check_col);
        if !adjacent_display.is_empty() {
            break; // Adjacent cell has content, stop spilling
        }

        // Add this column's width to spill area
        let adj_col_width = app.metrics.col_width(app.col_width(check_col));
        spill_width += adj_col_width;
        check_col += 1;
    }

    if spill_width > 0.0 {
        Some(spill_width)
    } else {
        None
    }
}

/// Check if a cell's text should be rendered by the spill overlay instead of the normal cell.
/// Use this in render_cell to avoid double-drawing text.
///
/// Returns true if:
/// - Text overflows the cell width
/// - Alignment allows spilling (left-aligned or General for text)
/// - There's at least one adjacent empty cell to spill into
///
/// When this returns true, render_cell should NOT draw the text - the spill overlay will handle it.
#[inline]
fn should_use_spill_overlay(
    data_row: usize,
    col: usize,
    text: &str,
    cell_width: f32,
    alignment: Alignment,
    is_selected: bool,
    is_active: bool,
    is_editing: bool,
    app: &Spreadsheet,
    window: &Window,
    cx: &App,
) -> bool {
    // Selected, active, or editing cells don't use spill overlay (z-order complexity)
    if is_selected || is_active || is_editing {
        return false;
    }

    // Check if this cell would spill (text overflows AND can spill into adjacent cells)
    calculate_text_spill(data_row, col, text, cell_width, alignment, app, window, cx).is_some()
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

fn cell_border(app: &Spreadsheet, is_editing: bool, is_active: bool, is_selected: bool, _formula_ref_color: Option<usize>) -> Hsla {
    // Selection border ALWAYS wins over formula ref (user needs to see what they've selected)
    // Formula ref borders are now drawn as dashed overlays (render_formula_ref_borders)
    if is_editing || is_active {
        app.token(TokenKey::CellBorderFocus)
    } else if is_selected {
        app.token(TokenKey::SelectionBorder)
    } else {
        // Formula refs use dashed border overlay, not per-cell solid borders
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

/// Render dashed borders around formula reference ranges.
/// Uses canvas to draw dashed rectangles around each unique referenced range.
/// Only renders when in formula edit mode with at least one reference.
/// Only shows in the active pane (where editing is happening).
fn render_formula_ref_borders(app: &Spreadsheet, pane_side: Option<SplitSide>) -> impl IntoElement {
    // Only render in formula edit mode
    if !app.mode.is_formula() && !(app.mode.is_editing() && app.is_formula_content()) {
        return div().into_any_element();
    }

    // Only show formula ref borders in the active pane
    let is_active_pane = match pane_side {
        None => true, // No split, always show
        Some(side) => side == app.split_active_side,
    };
    if !is_active_pane {
        return div().into_any_element();
    }

    let refs = &app.formula_highlighted_refs;
    if refs.is_empty() {
        return div().into_any_element();
    }

    // Collect unique ranges with their color indices
    // We deduplicate by RefKey so each range gets one border
    let mut seen_keys = std::collections::HashSet::new();
    let mut ranges: Vec<(RefKey, usize)> = Vec::new();

    for fref in refs {
        if seen_keys.insert(fref.key.clone()) {
            ranges.push((fref.key.clone(), fref.color_index));
        }
    }

    // Compute bounds for each range using the pane's view state
    let view_state = get_pane_view_state(app, pane_side);
    let scroll_row = view_state.scroll_row;
    let scroll_col = view_state.scroll_col;

    let range_bounds: Vec<(Bounds<Pixels>, Hsla)> = ranges
        .iter()
        .filter_map(|(key, color_idx)| {
            let (r1, c1, r2, c2) = match key {
                RefKey::Cell { row, col } => (*row, *col, *row, *col),
                RefKey::Range { r1, c1, r2, c2 } => (*r1, *c1, *r2, *c2),
            };

            // Check if any part of the range is visible
            let visible_rows = app.visible_rows();
            let visible_cols = app.visible_cols();

            // Skip if entirely off-screen
            if r2 < scroll_row || c2 < scroll_col {
                return None;
            }
            if r1 >= scroll_row + visible_rows || c1 >= scroll_col + visible_cols {
                return None;
            }

            // Get bounds of the range using cell_rect
            // cell_rect returns positions relative to the cell grid (after row header)
            // We need to add header_w offset since canvas covers entire grid div
            let top_left = app.cell_rect(r1, c1);
            let bottom_right = app.cell_rect(r2, c2);
            let header_w = app.metrics.header_w;

            let x = top_left.x + header_w;
            let y = top_left.y;
            let width = (bottom_right.x + bottom_right.width) - top_left.x;
            let height = (bottom_right.y + bottom_right.height) - top_left.y;

            let bounds = Bounds {
                origin: Point::new(px(x), px(y)),
                size: Size {
                    width: px(width),
                    height: px(height),
                },
            };

            // Ensure full opacity for borders (rgb() should produce opaque, but be explicit)
            let mut color: Hsla = rgb(REF_COLORS[*color_idx % 8]).into();
            color.a = 1.0;

            Some((bounds, color))
        })
        .collect();

    if range_bounds.is_empty() {
        return div().into_any_element();
    }

    // Use canvas to paint dashed borders
    canvas(
        // Prepaint: pass range_bounds to paint
        move |_bounds, _window, _cx| range_bounds,
        // Paint: draw dashed borders
        // paint_quad uses window-relative coordinates, so we add canvas origin
        move |canvas_bounds, range_bounds, window, _cx| {
            for (cell_bounds, color) in &range_bounds {
                // Offset by canvas origin to convert from grid-relative to window-relative
                let adjusted_bounds = Bounds {
                    origin: Point::new(
                        canvas_bounds.origin.x + cell_bounds.origin.x,
                        canvas_bounds.origin.y + cell_bounds.origin.y,
                    ),
                    size: cell_bounds.size,
                };
                let quad = gpui::quad(
                    adjusted_bounds,
                    px(0.0), // no corner radius
                    gpui::transparent_black(), // no fill (keep existing cell fills)
                    px(2.0), // 2px border width
                    *color,
                    BorderStyle::Dashed,
                );
                window.paint_quad(quad);
            }
        },
    )
    .absolute()
    .inset_0()
    .size_full()
    .into_any_element()
}

/// Pre-computed geometry for a visible merge overlay.
struct VisibleMerge {
    origin_row: usize,
    origin_col: usize,
    end_row: usize,
    end_col: usize,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
}

/// Collect merges that overlap the current viewport.
///
/// Pure geometry — no gpui types. Takes scroll position + viewport size,
/// returns merge regions with their pixel rects.
fn collect_visible_merges(
    app: &Spreadsheet,
    cx: &App,
    scroll_row: usize,
    scroll_col: usize,
    visible_rows: usize,
    visible_cols: usize,
) -> Vec<VisibleMerge> {
    let sheet = app.sheet(cx);
    if sheet.merged_regions.is_empty() {
        return Vec::new();
    }

    let metrics = &app.metrics;

    let mut out = Vec::new();
    for merge in &sheet.merged_regions {
        if !merge.overlaps_viewport(scroll_row, scroll_col, visible_rows, visible_cols) {
            continue;
        }

        // Use pixel_rect for rect computation — same math as col_x_offset/row_y_offset
        // but via the engine's pure geometry method (testable without gpui).
        // Note: col_x_offset applies pixel snapping; we do the same via metrics helpers.
        let (x, y, width, height) = merge.pixel_rect(
            scroll_row,
            scroll_col,
            |c| metrics.col_width(app.col_width(c)),
            |r| metrics.row_height(app.row_height(r)),
        );

        out.push(VisibleMerge {
            origin_row: merge.start.0,
            origin_col: merge.start.1,
            end_row: merge.end.0,
            end_col: merge.end.1,
            x,
            y,
            width,
            height,
        });
    }
    out
}

/// Build the gpui element for a single merge overlay.
///
/// Layers (bottom → top): background → selection tint → text → user borders → selection border.
/// Has `.id()` and mouse handlers — this IS the interactive surface for the merge.
fn render_merge_div(
    m: &VisibleMerge,
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
    view_state: &WorkbookViewState,
    pane_side: Option<SplitSide>,
    editing: bool,
    show_gridlines: bool,
    gridline_color: Hsla,
    sel_border_color: Hsla,
    user_border_color: Hsla,
) -> Stateful<Div> {
    let sheet = app.sheet(cx);
    let format = sheet.get_format(m.origin_row, m.origin_col);
    let bg = cell_base_background(app, false, format.background_color);
    let is_editing_this = editing && view_state.selected == (m.origin_row, m.origin_col);

    // 1. Background + position
    let mut merge_div = div()
        .id(ElementId::Name(
            format!("merge-{}-{}", m.origin_row, m.origin_col).into()
        ))
        .absolute()
        .left(px(m.x))
        .top(px(m.y))
        .w(px(m.width))
        .h(px(m.height))
        .overflow_hidden()
        .bg(bg)
        .flex();

    // Gridlines: all four perimeter sides (overlay covers underlying cells)
    if show_gridlines {
        merge_div = merge_div.border_color(gridline_color);
        if m.origin_row > 0 { merge_div = merge_div.border_t_1(); }
        if m.origin_col > 0 { merge_div = merge_div.border_l_1(); }
        merge_div = merge_div.border_b_1().border_r_1();
    }

    // 2. Selection tint
    let is_active = view_state.selected == (m.origin_row, m.origin_col)
        || (view_state.selected.0 >= m.origin_row && view_state.selected.0 <= m.end_row
            && view_state.selected.1 >= m.origin_col && view_state.selected.1 <= m.end_col);
    let is_selected = is_merge_in_selection(view_state, m.origin_row, m.origin_col, m.end_row, m.end_col);

    if !is_editing_this {
        if let Some(overlay_color) = selection_overlay_color(app, is_active, is_selected, None) {
            merge_div = merge_div.child(
                div().absolute().inset_0().bg(overlay_color)
            );
        }
    }

    // 3. Text (skip if editing origin — caret renders in inline cell)
    if !is_editing_this {
        let is_number = matches!(
            sheet.get_computed_value(m.origin_row, m.origin_col),
            Value::Number(_)
        );

        merge_div = match format.alignment {
            Alignment::General => {
                if is_number { merge_div.justify_end() } else { merge_div.justify_start() }
            }
            Alignment::Left => merge_div.justify_start(),
            Alignment::Center | Alignment::CenterAcrossSelection => merge_div.justify_center(),
            Alignment::Right => merge_div.justify_end(),
        };

        merge_div = match format.vertical_alignment {
            VerticalAlignment::Top => merge_div.items_start(),
            VerticalAlignment::Middle => merge_div.items_center(),
            VerticalAlignment::Bottom => merge_div.items_end(),
        };

        let value = if app.show_formulas() {
            sheet.get_raw(m.origin_row, m.origin_col)
        } else {
            let display = sheet.get_formatted_display(m.origin_row, m.origin_col);
            if !app.show_zeros() && display == "0" {
                String::new()
            } else {
                display
            }
        };

        if !value.is_empty() {
            merge_div = merge_div.px_1().child(
                render_merge_text(&value, &format, app, window)
            );
        }
    }

    // 4. User borders (non-interactive overlay — no .id())
    let (b_top, b_right, b_bottom, b_left) =
        sheet.resolve_merge_borders(
            sheet.get_merge(m.origin_row, m.origin_col).unwrap()
        );
    if b_top.is_set() || b_right.is_set() || b_bottom.is_set() || b_left.is_set() {
        merge_div = merge_div.child(
            non_interactive_overlay()
                .border_color(user_border_color)
                .when(b_top.is_set(), |d| d.border_t_1())
                .when(b_right.is_set(), |d| d.border_r_1())
                .when(b_bottom.is_set(), |d| d.border_b_1())
                .when(b_left.is_set(), |d| d.border_l_1())
        );
    }

    // 5. Selection border (on top of everything — wins over user borders visually)
    if is_selected || is_active {
        merge_div = merge_div.child(
            non_interactive_overlay()
                .border_color(sel_border_color)
                .border_1()
        );
    }

    // Mouse handlers: the overlay IS the interactive surface for merges.
    let origin_row = m.origin_row;
    let origin_col = m.origin_col;
    let end_row = m.end_row;
    let end_col = m.end_col;

    merge_div
        .cursor(CursorStyle::Crosshair) // Grid interior = crosshair
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            if this.renaming_sheet.is_some() {
                this.confirm_sheet_rename(cx);
            }
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
                return;
            }
            if this.is_fill_dragging() {
                return;
            }

            this.activate_pane(pane_side, cx);

            let target_row = origin_row;
            let target_col = origin_col;
            let merge_end = Some((end_row, end_col));

            if this.mode.is_formula() {
                if event.modifiers.shift {
                    this.formula_shift_click_ref(target_row, target_col, cx);
                } else {
                    this.formula_start_drag(target_row, target_col, cx);
                }
                return;
            }

            if this.mode == crate::mode::Mode::FormatPainter {
                this.select_cell(target_row, target_col, false, cx);
                this.apply_format_painter(cx);
                return;
            }

            // If editing, commit the current edit before navigating to this merge
            if this.mode.is_editing() {
                let active = this.active_view_state().selected;
                if (target_row, target_col) != active {
                    this.commit_pending_edit(cx);
                }
            }

            if event.click_count == 2 {
                this.select_cell(target_row, target_col, false, cx);
                this.start_edit(cx);
                this.maybe_show_f2_tip(cx);
            } else if event.modifiers.shift {
                this.select_cell(target_row, target_col, true, cx);
            } else if event.modifiers.control || event.modifiers.platform {
                this.start_ctrl_drag_selection(target_row, target_col, cx);
            } else {
                this.start_drag_selection(target_row, target_col, cx);
            }

            // Expand selection to cover full merge region
            if let Some(merge_br) = merge_end {
                if merge_br != (target_row, target_col) {
                    let (anchor_row, anchor_col) = this.active_view_state().selected;
                    let far_row = if anchor_row <= target_row { merge_br.0 } else { target_row };
                    let far_col = if anchor_col <= target_col { merge_br.1 } else { target_col };
                    this.active_view_state_mut().selection_end = Some((far_row, far_col));
                    cx.notify();
                }
            }
        }))
        .on_mouse_move(cx.listener(move |this, _event: &MouseMoveEvent, _, cx| {
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            if this.is_fill_dragging() {
                this.continue_fill_drag(origin_row, origin_col, cx);
                return;
            }
            if this.dragging_selection {
                if this.mode.is_formula() {
                    this.formula_continue_drag(origin_row, origin_col, cx);
                } else {
                    this.continue_drag_selection(origin_row, origin_col, cx);
                }
            }
        }))
        .on_mouse_up(MouseButton::Left, cx.listener(move |this, event: &MouseUpEvent, _, cx| {
            if this.inspector_visible || this.filter_dropdown_col.is_some() {
                return;
            }
            if this.is_fill_dragging() {
                let ctrl_held = event.modifiers.control || event.modifiers.platform;
                this.end_fill_drag(ctrl_held, cx);
                return;
            }
            this.end_drag_selection(cx);
        }))
}

/// Render merged cell overlays as absolutely-positioned elements.
///
/// The grid uses flex layout (rows of cells), which cannot support multi-row spans.
/// This overlay renders merged cells on top of the flex grid, covering the full
/// merged region. Origin and hidden cells in the flex grid become transparent
/// spacers (handled by render_cell), and this overlay paints the unified merged cell.
///
/// Z-order: cell grid → merge overlays → text spill → formula refs → popups
///
/// Each merge overlay has .id() and mouse handlers — it IS the interactive surface
/// for merged cells. Without these, clicks and drag selection through merges would break.
fn render_merge_overlays(
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
    pane_side: Option<SplitSide>,
) -> AnyElement {
    let view_state = get_pane_view_state(app, pane_side);
    let editing = app.mode.is_editing();

    let show_gridlines = match &user_settings(cx).appearance.show_gridlines {
        Setting::Value(v) => *v,
        Setting::Inherit => true,
    };

    let visible_merges = collect_visible_merges(
        app, cx,
        view_state.scroll_row, view_state.scroll_col,
        app.visible_rows(), app.visible_cols(),
    );

    if visible_merges.is_empty() {
        return div().into_any_element();
    }

    let header_width = crate::app::HEADER_WIDTH * app.metrics.zoom;
    let gridline_color = app.token(TokenKey::GridLines);
    let sel_border_color = app.token(TokenKey::SelectionBorder);
    let user_border_color = app.token(TokenKey::UserBorder);

    div()
        .absolute()
        .top_0()
        .bottom_0()
        .left(px(header_width))
        .right_0()
        .overflow_hidden()
        .cursor(CursorStyle::Crosshair) // Match child merge overlays to prevent flicker on transition
        .children(visible_merges.iter().map(|m| {
            render_merge_div(
                m, app, window, cx, view_state, pane_side, editing,
                show_gridlines, gridline_color, sel_border_color, user_border_color,
            )
        }))
        .into_any_element()
}

/// Check if a merge region overlaps the current selection.
fn is_merge_in_selection(
    view_state: &WorkbookViewState,
    merge_start_row: usize,
    merge_start_col: usize,
    merge_end_row: usize,
    merge_end_col: usize,
) -> bool {
    let (ar, ac) = view_state.selected;
    // Active cell is inside merge
    if ar >= merge_start_row && ar <= merge_end_row && ac >= merge_start_col && ac <= merge_end_col {
        return true;
    }
    // Selection range overlaps merge
    if let Some(sel_end) = view_state.selection_end {
        let (min_r, max_r) = (ar.min(sel_end.0), ar.max(sel_end.0));
        let (min_c, max_c) = (ac.min(sel_end.1), ac.max(sel_end.1));
        merge_end_row >= min_r && merge_start_row <= max_r
            && merge_end_col >= min_c && merge_start_col <= max_c
    } else {
        false
    }
}

/// Render formatted text for a merge overlay.
fn render_merge_text(
    value: &str,
    format: &visigrid_engine::cell::CellFormat,
    app: &Spreadsheet,
    window: &Window,
) -> AnyElement {
    let text_content: SharedString = value.to_string().into();
    let has_formatting = format.bold
        || format.italic
        || format.underline
        || format.strikethrough
        || format.font_family.is_some()
        || format.font_size.is_some()
        || format.font_color.is_some();

    let text_color = app.token(TokenKey::CellText);

    if has_formatting {
        let mut text_style = window.text_style();
        text_style.color = text_color;
        text_style.font_size = px(app.metrics.font_size).into();

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
        if format.strikethrough {
            text_style.strikethrough = Some(StrikethroughStyle {
                thickness: px(1.),
                ..Default::default()
            });
        }
        if let Some(ref family) = format.font_family {
            text_style.font_family = family.clone().into();
        }
        if let Some(rgba) = format.font_color {
            text_style.color = gpui::Hsla::from(gpui::Rgba {
                r: rgba[0] as f32 / 255.0,
                g: rgba[1] as f32 / 255.0,
                b: rgba[2] as f32 / 255.0,
                a: rgba[3] as f32 / 255.0,
            });
        }

        let styled = StyledText::new(text_content).with_default_highlights(&text_style, []);
        // Font size: TextRun doesn't carry font_size, so cascade via parent div.
        if let Some(size) = format.font_size {
            div()
                .text_size(px(size * app.metrics.zoom))
                .child(styled)
                .into_any_element()
        } else {
            styled.into_any_element()
        }
    } else {
        div()
            .text_color(text_color)
            .text_size(px(app.metrics.font_size))
            .child(text_content)
            .into_any_element()
    }
}

/// Render text that spills into adjacent empty cells (Excel-style overflow).
/// This is a separate overlay pass that draws AFTER all cell backgrounds,
/// ensuring spilled text isn't covered by neighboring cells.
///
/// Key invariants:
/// - Only left-aligned text (or General for text values) spills rightward
/// - Spill stops at first non-empty adjacent cell
/// - Selected/active/editing cells don't spill (simplifies z-order)
/// - This is paint-only: no hit testing (clicks go to underlying cells)
///
/// IMPORTANT: Spill text is rendered in a second pass as a paint-only overlay.
/// - Never attach IDs or handlers (hit testing must remain cell-based)
/// - Must render after cell backgrounds and before borders/selections
/// - Only for display mode (never during edit)
/// Breaking this will cause z-order, clipping, or selection bugs.
fn render_text_spill_overlay(
    app: &Spreadsheet,
    window: &Window,
    cx: &App,
    pane_side: Option<SplitSide>,
) -> impl IntoElement {
    // Get view state for this pane
    let view_state = get_pane_view_state(app, pane_side);
    let scroll_row = view_state.scroll_row;
    let scroll_col = view_state.scroll_col;
    let (selected_row, selected_col) = view_state.selected;

    let visible_rows = app.visible_rows();
    let visible_cols = app.visible_cols();
    let metrics = &app.metrics;

    // Collect spill runs: cells whose text overflows into adjacent empty cells
    // Use HashMap to dedupe by (data_row, col) - prevents double-draw with freeze panes
    struct SpillRun {
        x: f32,           // Screen X position (relative to grid)
        y: f32,           // Screen Y position
        base_width: f32,  // Original cell width (for alignment calculation)
        total_width: f32, // Cell width + spill width (paint region)
        height: f32,      // Cell height
        text: String,
        text_width: f32,  // Shaped text width for alignment
        text_color: Hsla,
        font_size: f32,
        alignment: Alignment, // For text positioning within base_width
        bold: bool,
        italic: bool,
        underline: bool,
    }

    let mut spill_runs: std::collections::HashMap<(usize, usize), SpillRun> = std::collections::HashMap::new();

    // Text color for non-selected cells
    let cell_text = app.token(TokenKey::CellText);

    // Scan visible cells for spill candidates
    for screen_row in 0..visible_rows {
        let visible_index = scroll_row + screen_row;

        // Get view_row and data_row for this screen position
        let Some((view_row, data_row)) = app.nth_visible_row(visible_index, cx) else {
            continue;
        };

        // Calculate Y position for this row
        let mut y: f32 = 0.0;
        for r in 0..screen_row {
            let idx = scroll_row + r;
            if let Some((vr, _)) = app.nth_visible_row(idx, cx) {
                y += metrics.row_height(app.row_height(vr));
            }
        }
        let row_height = metrics.row_height(app.row_height(view_row));

        for screen_col in 0..visible_cols {
            let col = scroll_col + screen_col;

            // Skip selected/active cells (they don't spill to avoid z-order complexity)
            if view_row == selected_row && col == selected_col {
                continue;
            }

            // GUARD: Merged cells must not enter the spill scan.
            // Merge text is rendered by the merge overlay (render_merge_overlays).
            // If this guard is removed, merge-origin text will be drawn TWICE:
            // once by the overlay and once by the spill layer, causing visual
            // artifacts (double text, wrong z-order, misaligned spill runs).
            // The predicate get_merge() is tested in engine tests:
            //   test_merge_get_merge_covers_all_cells
            //   test_merge_with_text_still_detected_by_get_merge
            if app.sheet(cx).get_merge(data_row, col).is_some() {
                continue;
            }

            // Get cell display value
            let display = app.sheet(cx).get_formatted_display(data_row, col);
            if display.is_empty() {
                continue;
            }

            // Get format and check if alignment allows spilling
            let format = app.sheet(cx).get_format(data_row, col);
            let is_number = matches!(
                app.sheet(cx).get_computed_value(data_row, col),
                Value::Number(_)
            );

            if !should_alignment_spill(format.alignment, is_number) {
                continue;
            }

            // Calculate cell width
            let col_width = metrics.col_width(app.col_width(col));

            // Shape text to get its pixel width
            let text_owned = display.clone();
            let text_shared: SharedString = text_owned.clone().into();
            let text_len = text_shared.len();
            if text_len == 0 {
                continue;
            }

            let effective_font_size = format.font_size.map(|s| s * metrics.zoom).unwrap_or(metrics.font_size);
            let shaped = window.text_system().shape_line(
                text_shared,
                px(effective_font_size),
                &[TextRun {
                    len: text_len,
                    font: Font::default(),
                    color: Hsla::default(),
                    background_color: None,
                    underline: None,
                    strikethrough: None,
                }],
                None,
            );

            let text_width: f32 = shaped.width.into();
            let padding = 8.0; // px_1 = 4px each side
            let available_width = col_width - padding;

            if text_width <= available_width {
                continue; // Text fits, no spill needed
            }

            // Calculate how many columns we can spill into
            let overflow_needed = text_width - available_width;
            let mut spill_width: f32 = 0.0;
            let mut check_col = col + 1;
            let max_col = col + 10; // Limit spillover

            while spill_width < overflow_needed && check_col < max_col {
                // Check if adjacent cell is empty
                let adjacent_display = app.sheet(cx).get_formatted_display(data_row, check_col);
                if !adjacent_display.is_empty() {
                    break; // Adjacent cell has content, stop spilling
                }

                // Add this column's width to spill area
                let adj_col_width = metrics.col_width(app.col_width(check_col));
                spill_width += adj_col_width;
                check_col += 1;
            }

            if spill_width <= 0.0 {
                continue; // Can't spill anywhere
            }

            // Calculate X position for this cell
            let mut x: f32 = 0.0;
            for c in scroll_col..col {
                x += metrics.col_width(app.col_width(c));
            }

            // Use entry to dedupe - keep first entry for each (data_row, col)
            // This prevents double-draw when freeze panes render same cell in multiple quadrants
            // IMPORTANT: Use resolved_alignment to get the effective alignment (General → Left for text)
            let effective_alignment = resolved_alignment(format.alignment, is_number);
            spill_runs.entry((data_row, col)).or_insert(SpillRun {
                x,
                y,
                base_width: col_width,           // Original cell width for alignment
                total_width: col_width + spill_width,  // Extended paint region
                height: row_height,
                text: text_owned,
                text_width,                      // For alignment calculation
                text_color: if let Some(rgba) = format.font_color {
                    gpui::Hsla::from(gpui::Rgba {
                        r: rgba[0] as f32 / 255.0,
                        g: rgba[1] as f32 / 255.0,
                        b: rgba[2] as f32 / 255.0,
                        a: rgba[3] as f32 / 255.0,
                    })
                } else {
                    cell_text
                },
                font_size: format.font_size.map(|s| s * metrics.zoom).unwrap_or(metrics.font_size),
                alignment: effective_alignment,  // Resolved alignment for text positioning
                bold: format.bold,
                italic: format.italic,
                underline: format.underline,
            });
        }
    }

    if spill_runs.is_empty() {
        return div().into_any_element();
    }

    // Render spill runs as absolutely positioned divs
    // These sit above cell backgrounds but have no hit testing (paint-only)
    //
    // IMPORTANT: Text is anchored to the ORIGINAL cell (base_width), not the spill area.
    // Spill only extends the paint region, not the alignment reference.
    // - Left-aligned: text starts at left edge of original cell
    // - Center-aligned: text centered within original cell (may extend right)
    // - Right-aligned: text right-aligned within original cell (shouldn't spill, but handle gracefully)
    let padding = 4.0; // Same as px_1() = 4px each side

    // The overlay must be clipped to the grid content area (excluding row headers).
    // Row headers are HEADER_WIDTH (50px) wide, rendered inside each row.
    // Position overlay at left=HEADER_WIDTH to align with cell content area.
    let header_width = crate::app::HEADER_WIDTH * app.metrics.zoom;

    div()
        .absolute()
        .top_0()
        .bottom_0()
        .left(px(header_width))  // Offset past row headers
        .right_0()
        .overflow_hidden()       // Clip spills at grid boundary
        .children(
            spill_runs.into_values().map(|run| {
                // Compute text x-offset based on alignment within BASE cell (not spill area)
                // This is the key: text position is calculated as if confined to original cell,
                // but we paint into the extended region.
                // Note: alignment is pre-resolved (General → Left for text), so we only see explicit values.
                let text_x_offset = match run.alignment {
                    Alignment::Left | Alignment::General => {
                        // Left-aligned: start at left edge + padding
                        padding
                    }
                    Alignment::Center | Alignment::CenterAcrossSelection => {
                        // Center within base cell (not spill area) - rare for spillable text
                        let available = run.base_width - (padding * 2.0);
                        padding + ((available - run.text_width) / 2.0).max(0.0)
                    }
                    Alignment::Right => {
                        // Right-aligned within base cell - shouldn't spill, but handle gracefully
                        (run.base_width - padding - run.text_width).max(padding)
                    }
                };

                let mut text_div = div()
                    .absolute()
                    .left(px(run.x))
                    .top(px(run.y))
                    .w(px(run.total_width))
                    .h(px(run.height))
                    .flex()
                    .items_center()
                    .overflow_hidden() // Clip to spill area
                    .text_color(run.text_color)
                    .text_size(px(run.font_size));

                // Apply formatting
                if run.bold {
                    text_div = text_div.font_weight(FontWeight::BOLD);
                }
                if run.italic {
                    text_div = text_div.italic();
                }

                // Position text with calculated offset (anchored to base cell)
                text_div.child(
                    div()
                        .pl(px(text_x_offset))
                        .child(run.text)
                )
            })
        )
        .into_any_element()
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

// Tests for spill logic are in visigrid-engine/src/sheet.rs (test_text_spill_*)
// because the binary crate doesn't support unit tests well.
//
