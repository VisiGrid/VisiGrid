//! Minimap view: a vertical density strip on the right edge of the grid.
//!
//! Shows where data exists across the used row range with a viewport rectangle
//! and click-to-jump / drag-to-scrub navigation.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, NUM_ROWS};
use crate::minimap::compute_buckets;
use crate::theme::TokenKey;

/// Minimap strip width in pixels.
const MINIMAP_WIDTH: f32 = 14.0;

/// Render the minimap strip.
///
/// Returns a narrow vertical div containing density bars and a viewport rectangle.
/// The minimap auto-recomputes when the sheet revision or height changes.
pub fn render_minimap(
    app: &mut Spreadsheet,
    _window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let accent = app.token(TokenKey::Accent);
    let panel_border = app.token(TokenKey::PanelBorder);
    let minimap_bg = app.token(TokenKey::PanelBg);

    // Use grid viewport height for minimap height
    let minimap_height = app.grid_layout.viewport_size.1;
    let height_px = minimap_height.max(1.0) as usize;

    // Recompute cache if stale (revision changed, sheet switched, or height changed)
    let sheet_id = app.cached_sheet_id().0;
    let cells_rev = app.cells_rev;
    if app.minimap_cache.last_revision != cells_rev
        || app.minimap_cache.last_sheet_id != sheet_id
        || app.minimap_cache.height_px != height_px
    {
        let cache = app.workbook.read(cx).active_sheet();
        app.minimap_cache = compute_buckets(cache, height_px, cells_rev, sheet_id);
    }

    let cache = &app.minimap_cache;
    let has_data = cache.has_data();
    let row_range = cache.row_range();

    // Build density segments: contiguous runs of non-zero buckets
    let mut segments: Vec<(f32, f32)> = Vec::new(); // (top_px, height_px)
    if has_data {
        let bucket_h = if cache.buckets.is_empty() {
            1.0
        } else {
            minimap_height / cache.buckets.len() as f32
        };
        let mut run_start: Option<usize> = None;
        for (i, &v) in cache.buckets.iter().enumerate() {
            if v > 0 {
                if run_start.is_none() {
                    run_start = Some(i);
                }
            } else if let Some(start) = run_start.take() {
                segments.push((start as f32 * bucket_h, (i - start) as f32 * bucket_h));
            }
        }
        if let Some(start) = run_start {
            segments.push((
                start as f32 * bucket_h,
                (cache.buckets.len() - start) as f32 * bucket_h,
            ));
        }
    }

    // Viewport rectangle position
    let (vp_top, vp_height) = if has_data && row_range > 0 {
        let scroll_row = app.view_state.scroll_row;
        let visible_rows = app.visible_rows();
        let used_min = cache.used_min;
        let used_max = cache.used_max;

        // Clamp scroll into used range
        let clamped_scroll = scroll_row.clamp(used_min, used_max);
        let vp_top = (clamped_scroll - used_min) as f32 / row_range as f32 * minimap_height;
        let vp_h = (visible_rows as f32 / row_range as f32 * minimap_height)
            .max(4.0); // Minimum 4px for visibility

        // Clamp to minimap bounds
        let vp_top = vp_top.min(minimap_height - vp_h.min(minimap_height));
        let vp_h = vp_h.min(minimap_height - vp_top);

        (vp_top, vp_h)
    } else {
        (0.0, 0.0)
    };

    let density_color = accent.opacity(0.35);
    let vp_fill = accent.opacity(0.12);
    let vp_border = accent.opacity(0.5);

    div()
        .id("minimap")
        .w(px(MINIMAP_WIDTH))
        .h_full()
        .flex_shrink_0()
        .relative()
        .bg(minimap_bg)
        .border_l_1()
        .border_color(panel_border)
        // Density segments
        .children(segments.into_iter().map(|(top, height)| {
            div()
                .absolute()
                .left_0()
                .right_0()
                .top(px(top))
                .h(px(height))
                .bg(density_color)
        }))
        .cursor_pointer()
        // Viewport rectangle (only when we have data)
        .when(has_data && vp_height > 0.0, |d| {
            d.child(
                div()
                    .absolute()
                    .left_0()
                    .right_0()
                    .top(px(vp_top))
                    .h(px(vp_height))
                    .bg(vp_fill)
                    .border_t_1()
                    .border_b_1()
                    .border_color(vp_border)
            )
        })
        // Mouse handlers
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
            if !this.minimap_cache.has_data() {
                return;
            }
            let y: f32 = event.position.y.into();
            // Convert window y to minimap-local y
            let minimap_top = this.grid_layout.grid_body_origin.1;
            let local_y = y - minimap_top;
            let mm_height = this.grid_layout.viewport_size.1;

            if local_y < 0.0 || local_y > mm_height || mm_height <= 0.0 {
                return;
            }

            let cache = &this.minimap_cache;
            let row_range = cache.row_range();
            if row_range == 0 {
                return;
            }

            // Check if click is inside viewport rectangle (drag) or outside (jump)
            let scroll_row = this.view_state.scroll_row;
            let visible_rows = this.visible_rows();
            let clamped_scroll = scroll_row.clamp(cache.used_min, cache.used_max);
            let vp_top_local = (clamped_scroll - cache.used_min) as f32 / row_range as f32 * mm_height;
            let vp_h = (visible_rows as f32 / row_range as f32 * mm_height).max(4.0);

            if local_y >= vp_top_local && local_y <= vp_top_local + vp_h {
                // Drag: anchor relative to viewport top
                this.minimap_dragging = true;
                this.minimap_drag_offset_y = local_y - vp_top_local;
            } else {
                // Jump: center viewport on clicked row
                let frac = local_y / mm_height;
                let target_row = cache.used_min + (frac * row_range as f32) as usize;
                minimap_scroll_to_row_centered(this, target_row, cx);
            }
            cx.notify();
        }))
        .on_mouse_move(cx.listener(move |this, event: &MouseMoveEvent, _, cx| {
            if !this.minimap_dragging {
                return;
            }
            let y: f32 = event.position.y.into();
            let minimap_top = this.grid_layout.grid_body_origin.1;
            let local_y = y - minimap_top;
            let mm_height = this.grid_layout.viewport_size.1;

            if mm_height <= 0.0 {
                return;
            }

            let cache = &this.minimap_cache;
            let row_range = cache.row_range();
            if row_range == 0 {
                return;
            }

            // Compute new viewport top from drag
            let new_vp_top = (local_y - this.minimap_drag_offset_y).clamp(0.0, mm_height);
            let frac = new_vp_top / mm_height;
            let target_row = cache.used_min + (frac * row_range as f32) as usize;

            let frozen_rows = this.view_state.frozen_rows;
            let max_scroll = NUM_ROWS.saturating_sub(this.visible_rows());
            this.view_state.scroll_row = target_row.clamp(frozen_rows, max_scroll);
            cx.notify();
        }))
        .on_mouse_up(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.minimap_dragging {
                this.minimap_dragging = false;
                cx.notify();
            }
        }))
}

/// Scroll so the target row is centered in the viewport.
fn minimap_scroll_to_row_centered(app: &mut Spreadsheet, target_row: usize, cx: &mut Context<Spreadsheet>) {
    let half = app.visible_rows() / 2;
    let max_scroll = NUM_ROWS.saturating_sub(app.visible_rows());
    let frozen_rows = app.view_state.frozen_rows;
    let new_scroll = target_row.saturating_sub(half).min(max_scroll).max(frozen_rows);
    app.view_state.scroll_row = new_scroll;
    cx.notify();
}
