use gpui::prelude::FluentBuilder;
use gpui::*;

use crate::app::Spreadsheet;
use crate::review_mode::OVERVIEW_BUCKET_COUNT;
use crate::theme::TokenKey;

pub const REVIEW_OVERVIEW_RAIL_WIDTH: f32 = 14.0;

pub fn render_review_overview_rail(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let Some(state) = app.review_mode.as_ref() else {
        return div().into_any_element();
    };
    let max_density = state
        .overview_buckets()
        .iter()
        .copied()
        .max()
        .unwrap_or(0)
        .max(1);
    let active_bucket = state.bucket_for_source_row(app.view_state.scroll_row);
    let accent = app.token(TokenKey::Accent);
    let border = app.token(TokenKey::PanelBorder);
    let panel_bg = app.token(TokenKey::PanelBg);

    div()
        .id("review-overview-rail")
        .flex()
        .flex_col()
        .flex_shrink_0()
        .w(px(REVIEW_OVERVIEW_RAIL_WIDTH))
        .h_full()
        .p(px(2.0))
        .gap(px(1.0))
        .bg(panel_bg)
        .border_l_1()
        .border_color(border)
        .children((0..OVERVIEW_BUCKET_COUNT).map(|bucket| {
            let count = state.overview_buckets()[bucket];
            let opacity = if count == 0 {
                0.0
            } else {
                0.22 + 0.68 * (count as f32 / max_density as f32).sqrt()
            };
            let target_row = state.source_row_for_bucket(bucket);
            div()
                .id(ElementId::Name(format!("review-density-{bucket}").into()))
                .flex_1()
                .w_full()
                .cursor_pointer()
                .bg(accent.opacity(opacity))
                .when(bucket == active_bucket, |marker| {
                    marker.border_1().border_color(accent)
                })
                .on_click(cx.listener(move |this, _, _, cx| {
                    this.view_state.selected = (target_row, 0);
                    this.view_state.selection_end = None;
                    this.view_state.scroll_row = target_row.saturating_sub(2);
                    cx.notify();
                }))
        }))
        .into_any_element()
}
