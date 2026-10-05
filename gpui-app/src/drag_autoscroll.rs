//! Auto-scroll while a drag leaves the grid.
//!
//! Selection, fill-handle, formula-reference and header drags extend only
//! when the pointer moves over a cell, so on their own they stop at the edge
//! of the viewport. While one of them is active and the pointer is outside
//! the grid body (in any direction, or outside the window), a timer scrolls
//! toward the pointer and extends the drag to the row or column that scrolled
//! into view — as Excel does. The further out the pointer, the faster it goes.

use std::time::Duration;

use gpui::*;

use crate::app::{Spreadsheet, NUM_ROWS};

/// Time between scroll steps while the pointer stays outside the grid.
const TICK: Duration = Duration::from_millis(50);
/// Pixels past the edge that add one more row or column per step.
const ACCEL_PX: f32 = 24.0;
const MAX_ROWS_PER_TICK: i32 = 20;
const MAX_COLS_PER_TICK: i32 = 5;

/// Rows or columns to move per step for a pointer `over` pixels past an
/// edge (negative = before the start edge). Zero when inside.
fn step_for(over: f32, max: i32) -> i32 {
    if over == 0.0 {
        return 0;
    }
    let n = (1 + (over.abs() / ACCEL_PX) as i32).min(max);
    if over > 0.0 { n } else { -n }
}

/// How far `pos` lies outside `[start, start + len)`: negative before it,
/// positive past it, zero inside.
fn overshoot(pos: f32, start: f32, len: f32) -> f32 {
    if pos < start {
        pos - start
    } else if pos > start + len {
        pos - (start + len)
    } else {
        0.0
    }
}

/// The (rows, cols) step for a pointer, given the grid body's origin and size
/// in window coordinates.
pub(crate) fn autoscroll_step(pointer: (f32, f32), origin: (f32, f32), size: (f32, f32)) -> (i32, i32) {
    (
        step_for(overshoot(pointer.1, origin.1, size.1), MAX_ROWS_PER_TICK),
        step_for(overshoot(pointer.0, origin.0, size.0), MAX_COLS_PER_TICK),
    )
}

impl Spreadsheet {
    /// A drag that should follow the pointer past the grid's edge.
    fn autoscrolling_drag_active(&self) -> bool {
        self.dragging_selection
            || self.is_fill_dragging()
            || self.dragging_row_header
            || self.dragging_col_header
    }

    /// Remember the cell a cell-level drag last reached, so a scroll in one
    /// direction keeps the other coordinate where the pointer left it.
    pub(crate) fn note_drag_cell(&mut self, row: usize, col: usize) {
        self.drag_last_cell = Some((row, col));
    }

    /// Called on every mouse move in the window. Starts the timer when a drag
    /// leaves the grid body and stops it when the pointer comes back, the
    /// drag ends, or the button is no longer held.
    pub(crate) fn update_drag_autoscroll(&mut self, event: &MouseMoveEvent, cx: &mut Context<Self>) {
        let held = event.pressed_button == Some(MouseButton::Left);
        if !held || !self.autoscrolling_drag_active() {
            self.stop_drag_autoscroll();
            return;
        }
        let pointer = (event.position.x.into(), event.position.y.into());
        self.drag_autoscroll_pointer = Some(pointer);
        let layout = &self.grid_layout;
        if autoscroll_step(pointer, layout.grid_body_origin, layout.viewport_size) == (0, 0) {
            self.stop_drag_autoscroll();
            return;
        }
        if self.drag_autoscroll_task.is_none() {
            // Step once now so the response is immediate, then on the timer
            self.drag_autoscroll_tick(cx);
            self.drag_autoscroll_task = Some(cx.spawn(async move |this, cx| loop {
                cx.background_executor().timer(TICK).await;
                let keep_going = this
                    .update(cx, |this, cx| this.drag_autoscroll_tick(cx))
                    .unwrap_or(false);
                if !keep_going {
                    break;
                }
            }));
        }
    }

    pub(crate) fn stop_drag_autoscroll(&mut self) {
        self.drag_autoscroll_task = None;
        self.drag_autoscroll_pointer = None;
    }

    /// The button came up somewhere other than a cell: off the grid, or
    /// outside the window after auto-scrolling. Finish a cell drag the way a
    /// cell's own mouse-up does, so a fill commits and a selection ends.
    /// A release over a cell has already done this, so it is a no-op then.
    pub(crate) fn release_drag_off_grid(&mut self, ctrl_held: bool, cx: &mut Context<Self>) {
        self.stop_drag_autoscroll();
        self.drag_last_cell = None;
        if self.is_fill_dragging() {
            self.end_fill_drag(ctrl_held, cx);
        } else if self.dragging_selection {
            self.end_drag_selection(cx);
            if self.mode == crate::mode::Mode::FormatPainter {
                self.apply_format_painter(cx);
            }
        }
    }

    /// One scroll step toward the pointer. Returns false when there is
    /// nothing left to do, which ends the timer.
    fn drag_autoscroll_tick(&mut self, cx: &mut Context<Self>) -> bool {
        let Some(pointer) = self.drag_autoscroll_pointer else { return false };
        if !self.autoscrolling_drag_active() {
            self.drag_autoscroll_pointer = None;
            return false;
        }
        let layout = &self.grid_layout;
        let (d_rows, d_cols) = autoscroll_step(pointer, layout.grid_body_origin, layout.viewport_size);
        if (d_rows, d_cols) == (0, 0) {
            self.drag_autoscroll_pointer = None;
            return false;
        }
        // Header drags only move along their own axis
        let d_rows = if self.dragging_col_header { 0 } else { d_rows };
        let d_cols = if self.dragging_row_header { 0 } else { d_cols };
        self.scroll(d_rows, d_cols, cx);

        let (row, col) = self.drag_autoscroll_target(d_rows, d_cols, cx);
        if self.is_fill_dragging() {
            self.continue_fill_drag(row, col, cx);
        } else if self.dragging_selection {
            if self.mode.is_formula() {
                self.formula_continue_drag(row, col, cx);
            } else {
                self.continue_drag_selection(row, col, cx);
            }
        } else if self.dragging_row_header {
            self.continue_row_header_drag(row, cx);
        } else if self.dragging_col_header {
            self.continue_col_header_drag(col, cx);
        }
        self.drag_last_cell = Some((row, col));
        true
    }

    /// The cell the drag should reach after a step: the row or column at the
    /// edge it is scrolling toward, and the last cell reached on the other axis.
    fn drag_autoscroll_target(&self, d_rows: i32, d_cols: i32, _cx: &App) -> (usize, usize) {
        let view_state = self.active_view_state();
        let (last_row, last_col) = self.drag_last_cell.or(view_state.selection_end).unwrap_or(view_state.selected);
        let rows = self.pane_rows(view_state);
        let cols = self.pane_cols(view_state);
        let row = match d_rows.signum() {
            1 => rows.body.last().map_or(last_row, |s| s.index),
            -1 => rows.body.first().map_or(last_row, |s| s.index),
            _ => last_row,
        };
        let col = match d_cols.signum() {
            1 => cols.body.last().map_or(last_col, |s| s.index),
            -1 => cols.body.first().map_or(last_col, |s| s.index),
            _ => last_col,
        };
        (row.min(NUM_ROWS - 1), col)
    }
}

#[cfg(test)]
mod tests {
    use super::autoscroll_step;

    const ORIGIN: (f32, f32) = (40.0, 100.0);
    const SIZE: (f32, f32) = (800.0, 600.0);

    #[test]
    fn inside_the_grid_does_not_scroll() {
        assert_eq!(autoscroll_step((400.0, 400.0), ORIGIN, SIZE), (0, 0));
        // Exactly on the edges is still inside
        assert_eq!(autoscroll_step((40.0, 100.0), ORIGIN, SIZE), (0, 0));
        assert_eq!(autoscroll_step((840.0, 700.0), ORIGIN, SIZE), (0, 0));
    }

    #[test]
    fn past_an_edge_scrolls_toward_it() {
        assert_eq!(autoscroll_step((400.0, 701.0), ORIGIN, SIZE), (1, 0)); // below
        assert_eq!(autoscroll_step((400.0, 99.0), ORIGIN, SIZE), (-1, 0)); // above
        assert_eq!(autoscroll_step((841.0, 400.0), ORIGIN, SIZE), (0, 1)); // right
        assert_eq!(autoscroll_step((39.0, 400.0), ORIGIN, SIZE), (0, -1)); // left
        assert_eq!(autoscroll_step((900.0, 750.0), ORIGIN, SIZE), (3, 3)); // diagonal
    }

    #[test]
    fn further_out_is_faster_up_to_a_cap() {
        let slow = autoscroll_step((400.0, 710.0), ORIGIN, SIZE).0;
        let fast = autoscroll_step((400.0, 800.0), ORIGIN, SIZE).0;
        assert!(fast > slow, "{fast} should exceed {slow}");
        assert_eq!(autoscroll_step((400.0, 10_000.0), ORIGIN, SIZE).0, 20);
        assert_eq!(autoscroll_step((10_000.0, 400.0), ORIGIN, SIZE).1, 5);
        // Outside the window entirely (negative coordinates) still works
        assert_eq!(autoscroll_step((400.0, -10_000.0), ORIGIN, SIZE).0, -20);
    }
}
