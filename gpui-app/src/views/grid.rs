use gpui::*;
use gpui::StyledText;
use crate::app::Spreadsheet;
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

    let value = if is_editing {
        edit_value.to_string()
    } else {
        app.sheet().get_display(row, col)
    };

    let format = app.sheet().get_format(row, col);
    let cell_row = row;
    let cell_col = col;

    let mut cell = div()
        .id(ElementId::Name(format!("cell-{}-{}", row, col).into()))
        .flex_shrink_0()
        .w(px(col_width))
        .h_full()
        .flex()
        .items_center()
        .px_1()
        .overflow_hidden()
        .bg(cell_background(is_editing, is_active, is_selected))
        .border_1()
        .border_color(cell_border(is_editing, is_active, is_selected))
        .text_color(cell_text_color(is_editing))
        .text_sm()
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            // Don't handle clicks if we're resizing
            if this.resizing_col.is_some() || this.resizing_row.is_some() {
                return;
            }

            // Formula mode: clicks insert cell references
            if this.mode.is_formula() {
                if event.modifiers.shift {
                    this.formula_shift_click_ref(cell_row, cell_col, cx);
                } else {
                    this.formula_click_ref(cell_row, cell_col, cx);
                }
                return;
            }

            // Normal mode handling
            if event.click_count == 2 {
                // Double-click to edit
                this.select_cell(cell_row, cell_col, false, cx);
                this.start_edit(cx);
            } else if event.modifiers.shift {
                // Shift+click extends selection
                this.select_cell(cell_row, cell_col, true, cx);
            } else if event.modifiers.control || event.modifiers.platform {
                // Ctrl+click (or Cmd on Mac) for discontiguous selection
                this.start_ctrl_drag_selection(cell_row, cell_col, cx);
            } else {
                // Start drag selection
                this.start_drag_selection(cell_row, cell_col, cx);
            }
        }))
        .on_mouse_move(cx.listener(move |this, _event: &MouseMoveEvent, _, cx| {
            // Continue drag selection if active
            if this.dragging_selection {
                this.continue_drag_selection(cell_row, cell_col, cx);
            }
        }))
        .on_mouse_up(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // End drag selection
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
    // Build the text content with cursor at correct position
    let text_content: SharedString = if is_editing {
        let cursor_pos = app.edit_cursor;
        // Insert cursor character at the correct position
        let chars: Vec<char> = value.chars().collect();
        let before: String = chars.iter().take(cursor_pos).collect();
        let after: String = chars.iter().skip(cursor_pos).collect();
        format!("{}|{}", before, after).into()
    } else {
        value.into()
    };

    // Apply per-cell font family using StyledText with explicit runs
    // BUG: Font rendering not working - gpui may need different approach
    // See: https://github.com/anthropics/visigrid/issues/XXX
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
            color: cell_text_color(is_editing),
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

fn cell_background(is_editing: bool, is_active: bool, is_selected: bool) -> Hsla {
    if is_editing {
        rgb(0xffffff).into()  // White when editing
    } else if is_active {
        rgb(0x264f78).into()  // Blue for active cell
    } else if is_selected {
        rgba(0x264f7880).into()  // Lighter blue for selection range (50% alpha)
    } else {
        rgb(0x1e1e1e).into()  // Default dark
    }
}

fn cell_border(is_editing: bool, is_active: bool, is_selected: bool) -> Hsla {
    if is_editing || is_active {
        rgb(0x007acc).into()  // Blue border
    } else if is_selected {
        rgba(0x007acc80).into()  // 50% alpha
    } else {
        rgb(0x3d3d3d).into()  // Default gray
    }
}

fn cell_text_color(is_editing: bool) -> Hsla {
    if is_editing {
        rgb(0x000000).into()  // Black text when editing
    } else {
        rgb(0xd4d4d4).into()  // Light gray text
    }
}
