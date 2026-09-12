use gpui::prelude::FluentBuilder;
use gpui::*;

use crate::app::Spreadsheet;
use crate::review_mode::{OVERVIEW_BUCKET_COUNT, review_proposal_color};
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
        .map(|bucket| bucket.change_count)
        .max()
        .unwrap_or(0)
        .max(1);
    let on_source_sheet = app.wb(cx).active_sheet_id() == state.source_sheet_id;
    let active_bucket = on_source_sheet.then(|| {
        state.bucket_for_source_row(app.view_to_data(app.view_state.scroll_row, cx))
    });
    let selected_data_row = on_source_sheet
        .then(|| app.view_to_data(app.view_state.selected.0, cx));
    let selected_bucket = selected_data_row
        .filter(|row| !state.change_indices_at_source_row(*row).is_empty())
        .map(|row| state.bucket_for_source_row(row));
    let proposal = review_proposal_color(app);
    let warning = app.token(TokenKey::Warn);
    let error = app.token(TokenKey::Error);
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
            let overview = state.overview_buckets()[bucket];
            let opacity = if overview.change_count == 0 {
                0.0
            } else {
                0.22 + 0.68 * (overview.change_count as f32 / max_density as f32).sqrt()
            };
            let marker_color = if overview.has_new_formula_error {
                error
            } else if overview.has_deleted_row {
                warning
            } else {
                proposal
            };
            div()
                .id(ElementId::Name(format!("review-density-{bucket}").into()))
                .flex_1()
                .w_full()
                .cursor_pointer()
                .bg(marker_color.opacity(opacity))
                .when(active_bucket == Some(bucket), |marker| {
                    marker.border_1().border_color(border)
                })
                .when(selected_bucket == Some(bucket), |marker| {
                    marker.border_2().border_color(proposal)
                })
                .on_click(cx.listener(move |this, _, _, cx| {
                    this.navigate_review_bucket(bucket, cx);
                }))
        }))
        .into_any_element()
}
