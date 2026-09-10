use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::operation_plan::ProblemSeverity;

use crate::app::Spreadsheet;
use crate::review_mode::ReviewEndpoint;
use crate::terminal::state::PendingResult;
use crate::theme::TokenKey;

pub const REVIEW_ACTION_BAR_HEIGHT: f32 = 36.0;

pub fn render_review_action_bar(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let Some(state) = app.review_mode.as_ref() else {
        return div().into_any_element();
    };
    let Some(PendingResult::LuaPreview(preview)) = app.terminal.pending_result.as_ref() else {
        return div().into_any_element();
    };
    let Some(prepared) = preview.prepared_plan.as_ref() else {
        return div().into_any_element();
    };
    if prepared.plan().id != state.plan_id {
        return div().into_any_element();
    }

    let plan = prepared.plan();
    let title = plan.title.clone();
    let total_changes = plan.summary.total_changes();
    let blocking = plan
        .problems
        .iter()
        .any(|problem| problem.severity == ProblemSeverity::Blocking);
    let source_changed = app.workbook.read(cx).revision() != plan.source_revision;
    let can_apply = !blocking && !source_changed;

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);
    let warning = app.token(TokenKey::Warn);

    let endpoint_button =
        |id: &'static str, label: &'static str, endpoint: ReviewEndpoint, active: bool| {
            div()
                .id(id)
                .cursor_pointer()
                .px(px(8.0))
                .py(px(3.0))
                .rounded(px(3.0))
                .text_size(px(11.0))
                .text_color(if active { text_inverse } else { text_muted })
                .when(active, |button| button.bg(accent))
                .when(!active, |button| {
                    button.hover(|style| style.text_color(text_primary))
                })
                .on_click(cx.listener(move |this, _, _, cx| {
                    if let Some(review) = this.review_mode.as_mut() {
                        review.set_endpoint(endpoint);
                        cx.notify();
                    }
                }))
                .child(label)
        };

    div()
        .id("review-action-bar")
        .flex()
        .flex_row()
        .flex_shrink_0()
        .items_center()
        .justify_between()
        .h(px(REVIEW_ACTION_BAR_HEIGHT))
        .px(px(10.0))
        .bg(panel_bg)
        .border_t_1()
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(8.0))
                .child(
                    div()
                        .text_size(px(11.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(format!("Review changes · {title}")),
                )
                .when(source_changed, |row| {
                    row.child(
                        div()
                            .text_size(px(10.0))
                            .text_color(warning)
                            .child("Source changed · re-preview required"),
                    )
                }),
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(4.0))
                .child(
                    div()
                        .id("review-previous-change")
                        .cursor_pointer()
                        .px(px(6.0))
                        .py(px(3.0))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .hover(|style| style.text_color(text_primary))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.navigate_review_change(false, false, cx);
                        }))
                        .child("Previous ["),
                )
                .child(
                    div()
                        .id("review-next-change")
                        .cursor_pointer()
                        .px(px(6.0))
                        .py(px(3.0))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .hover(|style| style.text_color(text_primary))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.navigate_review_change(true, false, cx);
                        }))
                        .child("Next ]"),
                )
                .child(endpoint_button(
                    "review-before",
                    "Before",
                    ReviewEndpoint::Before,
                    state.endpoint == ReviewEndpoint::Before,
                ))
                .child(endpoint_button(
                    "review-after",
                    "After",
                    ReviewEndpoint::After,
                    state.endpoint == ReviewEndpoint::After,
                )),
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap(px(8.0))
                .child(
                    div()
                        .id("review-dismiss")
                        .cursor_pointer()
                        .px(px(7.0))
                        .py(px(3.0))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .hover(|style| style.text_color(text_primary))
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.dismiss_structured_result(cx);
                        }))
                        .child("Dismiss"),
                )
                .child(
                    div()
                        .id("review-apply")
                        .px(px(9.0))
                        .py(px(4.0))
                        .rounded(px(3.0))
                        .text_size(px(11.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(if can_apply { text_inverse } else { text_muted })
                        .bg(if can_apply {
                            accent
                        } else {
                            panel_border.opacity(0.35)
                        })
                        .when(can_apply, |button| {
                            button
                                .cursor_pointer()
                                .hover(|style| style.bg(accent.opacity(0.82)))
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.apply_lua_to_current_sheet(window, cx);
                                }))
                        })
                        .child(format!("Apply {total_changes} changes")),
                ),
        )
        .into_any_element()
}
