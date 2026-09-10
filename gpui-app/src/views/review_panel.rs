use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::operation_plan::{
    ChangeCause, ChangeKind, MaterializedChange, ProblemSeverity, VerificationEvidence,
    VerificationStatus,
};

use crate::app::Spreadsheet;
use crate::terminal::state::PendingResult;
use crate::theme::TokenKey;

pub const REVIEW_PANEL_WIDTH: f32 = 286.0;

pub fn render_review_panel(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
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

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let ok = app.token(TokenKey::Ok);
    let warning = app.token(TokenKey::Warn);
    let error = app.token(TokenKey::Error);
    let summary = &plan.summary;

    let group_rows: Vec<_> = plan
        .groups
        .iter()
        .map(|group| {
            key_value_row(
                group.title.clone(),
                format!("{} changes", state.group_change_count(&group.id)),
                text_primary,
                text_muted,
            )
            .into_any_element()
        })
        .collect();

    let verification_rows: Vec<_> = plan.verification.iter().map(|result| {
        let status_color = match result.status {
            VerificationStatus::Passed => ok,
            VerificationStatus::Failed => error,
            VerificationStatus::Unknown => warning,
            VerificationStatus::NotRun => text_muted,
        };
        let status_label = match result.status {
            VerificationStatus::Passed => "Passed",
            VerificationStatus::Failed => "Failed",
            VerificationStatus::Unknown => "Unknown",
            VerificationStatus::NotRun => "Not run",
        };
        let evidence = match &result.evidence {
            VerificationEvidence::NoNewFormulaErrors { new_errors } => {
                format!("{new_errors} new formula errors")
            }
            VerificationEvidence::RetainedTotal {
                expected, actual, tolerance, currency, difference, ..
            } => format!(
                "Expected {expected} {currency} · actual {actual} {currency} · difference {difference} · tolerance {tolerance}"
            ),
            VerificationEvidence::Unavailable { message } => message.clone(),
        };
        div()
            .flex().flex_col().gap(px(2.0)).py(px(5.0))
            .child(
                div().flex().items_center().justify_between()
                    .child(div().text_size(px(11.0)).text_color(text_primary)
                        .child(result.label.clone().unwrap_or_else(|| result.id.clone())))
                    .child(div().text_size(px(10.0)).font_weight(FontWeight::SEMIBOLD)
                        .text_color(status_color).child(status_label))
            )
            .child(div().text_size(px(10.0)).text_color(text_muted).child(evidence))
            .into_any_element()
    }).collect();

    let selected_view = app.view_state.selected;
    let selected_data_row = app.view_to_data(selected_view.0, cx);
    let selected_change = app.review_change_for_source_selection(
        app.sheet(cx).id,
        selected_data_row,
        selected_view.1,
    );

    div()
        .id("review-panel")
        .flex()
        .flex_col()
        .flex_shrink_0()
        .w(px(REVIEW_PANEL_WIDTH))
        .h_full()
        .overflow_y_scroll()
        .bg(panel_bg)
        .border_l_1()
        .border_color(panel_border)
        .child(section_title("Review", text_primary, panel_border))
        .child(
            div()
                .flex()
                .flex_col()
                .px(px(10.0))
                .py(px(7.0))
                .child(key_value_row(
                    "Changed cells",
                    summary.cells_changed.to_string(),
                    text_primary,
                    text_muted,
                ))
                .child(key_value_row(
                    "Cleared cells",
                    summary.cells_cleared.to_string(),
                    text_primary,
                    text_muted,
                ))
                .child(key_value_row(
                    "Formula changes",
                    summary.formulas_changed.to_string(),
                    text_primary,
                    text_muted,
                ))
                .child(key_value_row(
                    "Deleted rows",
                    summary.rows_deleted.to_string(),
                    text_primary,
                    text_muted,
                ))
                .child(key_value_row(
                    "Recalculated",
                    summary.recalculated_cells.to_string(),
                    text_primary,
                    text_muted,
                )),
        )
        .when(!group_rows.is_empty(), |panel| {
            panel
                .child(section_title("Groups", text_primary, panel_border))
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .px(px(10.0))
                        .py(px(5.0))
                        .children(group_rows),
                )
        })
        .when(!verification_rows.is_empty(), |panel| {
            panel
                .child(section_title("Verification", text_primary, panel_border))
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .px(px(10.0))
                        .py(px(4.0))
                        .children(verification_rows),
                )
        })
        .when(!plan.problems.is_empty(), |panel| {
            panel
                .child(section_title("Problems", text_primary, panel_border))
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap(px(4.0))
                        .px(px(10.0))
                        .py(px(6.0))
                        .children(plan.problems.iter().map(|problem| {
                            let color = if problem.severity == ProblemSeverity::Blocking {
                                error
                            } else {
                                warning
                            };
                            div()
                                .text_size(px(10.0))
                                .text_color(color)
                                .child(problem.message.clone())
                        })),
                )
        })
        .child(section_title("Selected change", text_primary, panel_border))
        .child(render_selected_change(
            selected_change,
            text_primary,
            text_muted,
            app.token(TokenKey::Accent),
        ))
        .into_any_element()
}

fn section_title(label: &'static str, text: Hsla, border: Hsla) -> Div {
    div()
        .px(px(10.0))
        .py(px(6.0))
        .border_t_1()
        .border_b_1()
        .border_color(border)
        .text_size(px(10.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(text)
        .child(label)
}

fn key_value_row(
    label: impl Into<SharedString>,
    value: impl Into<SharedString>,
    text: Hsla,
    muted: Hsla,
) -> Div {
    div()
        .flex()
        .items_center()
        .justify_between()
        .py(px(2.0))
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text)
                .child(label.into()),
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(muted)
                .child(value.into()),
        )
}

fn render_selected_change(
    change: Option<&MaterializedChange>,
    text: Hsla,
    muted: Hsla,
    accent: Hsla,
) -> AnyElement {
    let Some(change) = change else {
        return div()
            .px(px(10.0))
            .py(px(8.0))
            .text_size(px(10.0))
            .text_color(muted)
            .child("Select a highlighted cell or deleted row.")
            .into_any_element();
    };

    let coordinate = if change.kind == ChangeKind::RowDeleted {
        change
            .before_coordinate
            .map(|coordinate| format!("Row {}", coordinate.row + 1))
            .unwrap_or_else(|| "Row".into())
    } else {
        change
            .before_coordinate
            .or(change.after_coordinate)
            .map(|coordinate| a1(coordinate.row, coordinate.col))
            .unwrap_or_else(|| "Cell".into())
    };
    let kind = match change.kind {
        ChangeKind::Value => "Value",
        ChangeKind::Formula => "Formula",
        ChangeKind::Cleared => "Clear",
        ChangeKind::Format => "Format",
        ChangeKind::RowDeleted => "Row deletion",
    };
    let cause = match change.cause {
        ChangeCause::Direct => "direct",
        ChangeCause::Recalculated => "recalculated",
    };
    let before = if change.kind == ChangeKind::Formula {
        display_or_empty(&change.before.raw)
    } else {
        display_or_empty(&change.before.display)
    };
    let after = if change.kind == ChangeKind::RowDeleted {
        "deleted".into()
    } else if change.kind == ChangeKind::Formula {
        display_or_empty(&change.after.raw)
    } else {
        display_or_empty(&change.after.display)
    };

    div()
        .flex()
        .flex_col()
        .gap(px(5.0))
        .px(px(10.0))
        .py(px(8.0))
        .child(
            div()
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(accent)
                .child(format!("{coordinate} · {kind} · {cause}")),
        )
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text)
                .child(format!("{before} → {after}")),
        )
        .when_some(change.reason.as_ref(), |detail, reason| {
            detail.child(
                div()
                    .text_size(px(10.0))
                    .text_color(muted)
                    .child(format!("Reason: {}", reason.0)),
            )
        })
        .when(!change.sources.is_empty(), |detail| {
            detail.child(
                div()
                    .text_size(px(10.0))
                    .text_color(muted)
                    .child(format!("Sources: {}", change.sources.join(", "))),
            )
        })
        .into_any_element()
}

fn display_or_empty(value: &str) -> String {
    if value.is_empty() {
        "(empty)".into()
    } else {
        format!("\"{value}\"")
    }
}

fn a1(row: usize, col: usize) -> String {
    let mut letters = String::new();
    let mut col = col + 1;
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    format!("{letters}{}", row + 1)
}
