use gpui::*;
use gpui::StyledText;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use super::headers::render_row_header;

/// Render the main cell grid
pub fn render_grid(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let scroll_row = app.scroll_row;
    let scroll_col = app.scroll_col;
    let selected = app.selected;
    let editing = app.mode.is_editing();
    let edit_value = app.edit_value.clone();
    let visible_rows = app.visible_rows();
    let visible_cols = app.visible_cols();

    div()
        .flex_1()
        .overflow_hidden()
        .child(
            div()
                .flex()
                .flex_col()
                .children(
                    (0..visible_rows).map(|visible_row| {
                        let row = scroll_row + visible_row;
                        render_row(
                            row,
                            scroll_col,
                            visible_cols,
                            selected,
                            editing,
                            &edit_value,
                            app,
                            cx,
                        )
                    })
                )
        )
}

fn render_row(
    row: usize,
    scroll_col: usize,
    visible_cols: usize,
    selected: (usize, usize),
    editing: bool,
    edit_value: &str,
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let row_height = app.row_height(row);

    div()
        .flex()
        .flex_shrink_0()
        .h(px(row_height))
        .child(render_row_header(app, row, cx))
        .children(
            (0..visible_cols).map(|visible_col| {
                let col = scroll_col + visible_col;
                let col_width = app.col_width(col);
                render_cell(row, col, col_width, row_height, selected, editing, edit_value, app, cx)
            })
        )
}

fn render_cell(
    row: usize,
    col: usize,
    col_width: f32,
    _row_height: f32,
    selected: (usize, usize),
    editing: bool,
    edit_value: &str,
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = app.is_selected(row, col);
    let is_active = selected == (row, col);
    let is_editing = editing && is_active;
    let is_formula_ref = app.is_formula_ref(row, col);

    // Spill state detection
    let is_spill_parent = app.sheet().is_spill_parent(row, col);
    let is_spill_receiver = app.sheet().is_spill_receiver(row, col);
    let has_spill_error = app.sheet().has_spill_error(row, col);

    let value = if is_editing {
        edit_value.to_string()
    } else if app.show_formulas {
        app.sheet().get_raw(row, col)
    } else {
        app.sheet().get_display(row, col)
    };

    let format = app.sheet().get_format(row, col);
    let cell_row = row;
    let cell_col = col;

    // Determine border color based on cell state (spill states take precedence over normal states)
    let border_color = if has_spill_error {
        app.token(TokenKey::SpillBlockedBorder)
    } else if is_spill_parent {
        app.token(TokenKey::SpillBorder)
    } else if is_spill_receiver {
        app.token(TokenKey::SpillReceiverBorder)
    } else {
        cell_border(app, is_editing, is_active, is_selected, is_formula_ref)
    };

    let needs_full_border = is_editing || is_active || is_selected;

    let mut cell = div()
        .id(ElementId::Name(format!("cell-{}-{}", row, col).into()))
        .flex_shrink_0()
        .w(px(col_width))
        .h_full()
        .flex()
        .items_center()
        .px_1()
        .overflow_hidden()
        .bg(cell_background(app, is_editing, is_active, is_selected, is_formula_ref))
        .border_color(border_color);

    // Only right+bottom borders for normal cells (thinner gridlines)
    // Full border for selected/editing cells
    // For formula refs, only draw outer edges of the range (not interior borders)
    // Spill parent/blocked get 2px border, receiver gets 1px
    cell = if needs_full_border {
        cell.border_1()
    } else if has_spill_error || is_spill_parent {
        // 2px solid border for spill parent and blocked cells
        cell.border_2()
    } else if is_spill_receiver {
        // 1px border for spill receivers (ideally dashed, but gpui doesn't support that)
        cell.border_1()
    } else if is_formula_ref {
        // Get which borders to draw for this formula ref cell
        let (top, right, bottom, left) = app.formula_ref_borders(row, col);
        let mut c = cell;
        if top { c = c.border_t_1(); }
        if right { c = c.border_r_1(); }
        if bottom { c = c.border_b_1(); }
        if left { c = c.border_l_1(); }
        c
    } else {
        cell.border_r_1().border_b_1()
    };

    cell = cell
        .text_color(cell_text_color(app, is_editing))
        .text_sm()
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Don't handle clicks if we're resizing
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
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
            // End drag selection (works for both normal and formula mode)
            this.end_drag_selection(cx);
        }));

    // Apply formatting
    if format.bold {
        cell = cell.font_weight(FontWeight::BOLD);
    }
    if format.italic {
        cell = cell.italic();
    }
    if format.underline {
        cell = cell.underline();
    }
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

            let normal_color = cell_text_color(app, is_editing);
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

            cell.child(StyledText::new(display_text).with_runs(runs))
        } else {
            // No selection - plain text with cursor
            cell.child(display_text)
        }
    } else {
        // Not editing - show value, possibly with custom font
        let text_content: SharedString = value.into();

        if let Some(ref font_family) = format.font_family {
            let run = TextRun {
                len: text_content.len(),
                font: Font {
                    family: font_family.clone().into(),
                    features: FontFeatures::default(),
                    fallbacks: None,
                    weight: if format.bold { FontWeight::BOLD } else { FontWeight::NORMAL },
                    style: if format.italic { FontStyle::Italic } else { FontStyle::Normal },
                },
                color: cell_text_color(app, is_editing),
                background_color: None,
                underline: if format.underline {
                    Some(UnderlineStyle {
                        thickness: px(1.0),
                        color: None,
                        wavy: false,
                    })
                } else {
                    None
                },
                strikethrough: None,
            };
            cell.child(StyledText::new(text_content).with_runs(vec![run]))
        } else {
            cell.child(text_content)
        }
    }
}

fn cell_background(app: &Spreadsheet, is_editing: bool, is_active: bool, is_selected: bool, is_formula_ref: bool) -> Hsla {
    if is_editing {
        app.token(TokenKey::EditorBg)
    } else if is_formula_ref {
        // Apply opacity so cell content remains visible through the highlight
        app.token(TokenKey::RefHighlight1).opacity(0.25)
    } else if is_active {
        app.token(TokenKey::SelectionBg)
    } else if is_selected {
        app.token(TokenKey::SelectionBg)
    } else {
        app.token(TokenKey::CellBg)
    }
}

fn cell_border(app: &Spreadsheet, is_editing: bool, is_active: bool, is_selected: bool, is_formula_ref: bool) -> Hsla {
    if is_editing || is_active {
        app.token(TokenKey::CellBorderFocus)
    } else if is_formula_ref {
        app.token(TokenKey::RefHighlight1)
    } else if is_selected {
        app.token(TokenKey::SelectionBorder)
    } else {
        app.token(TokenKey::GridLines)
    }
}

fn cell_text_color(app: &Spreadsheet, is_editing: bool) -> Hsla {
    if is_editing {
        app.token(TokenKey::EditorText)
    } else {
        app.token(TokenKey::CellText)
    }
}
