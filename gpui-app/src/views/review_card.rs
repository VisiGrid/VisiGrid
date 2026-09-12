use std::path::Path;

use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::operation_plan::{
    ChangeCause, ChangeKind, MaterializedChange, ProblemSeverity, VerificationEvidence,
    VerificationResult, VerificationStatus,
};

use crate::app::Spreadsheet;
use crate::review_mode::{review_proposal_color, ReviewEndpoint};
use crate::terminal::state::PendingResult;
use crate::theme::TokenKey;

pub const REVIEW_CARD_WIDTH: f32 = 360.0;
pub const REVIEW_CARD_HEIGHT_ESTIMATE: f32 = 292.0;
const REVIEW_CARD_GAP: f32 = 16.0;
const REVIEW_CARD_DATA_GUTTER_COLS: usize = 1;

pub fn render_review_card(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
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
    let total = state.review_item_count().max(1);
    let focused = state
        .focused_cell()
        .or_else(|| state.first_navigable_cell());
    let selected_change = focused.and_then(|(row, col)| {
        app.review_change_for_source_selection(state.source_sheet_id, row, col)
    });
    let position = focused
        .and_then(|(row, col)| state.review_position(row, col))
        .unwrap_or(0)
        .min(total - 1);

    let mut surface = app.token(TokenKey::PanelBg);
    // The review surface must remain legible over the grid watermark and cell
    // contents even when the active theme uses translucent panel surfaces.
    surface.a = 1.0;
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let inverse = app.token(TokenKey::TextInverse);
    let proposal = review_proposal_color(app);
    let ok = app.token(TokenKey::Ok);
    let warn = app.token(TokenKey::Warn);
    let error = app.token(TokenKey::Error);

    if state.collapsed {
        return collapsed_pill(total, proposal, surface, text, muted, cx).into_any_element();
    }

    let Some(change) = selected_change else {
        return div().into_any_element();
    };
    let apply_status = state.apply_status(prepared, app.workbook.read(cx));
    let can_apply = apply_status.can_apply();
    let is_mcp = plan.producer.kind.starts_with("mcp");
    let verification_failure = plan.verification.iter().find(|result| {
        matches!(
            result.status,
            VerificationStatus::Failed | VerificationStatus::Unknown
        )
    });
    let blocking_problem = plan
        .problems
        .iter()
        .find(|problem| problem.severity == ProblemSeverity::Blocking);

    let title = if apply_status.stale {
        "Proposal needs a refresh".to_string()
    } else if apply_status.blocking {
        "Proposal cannot be applied".to_string()
    } else {
        format!("{} ready to review", count_label(total))
    };
    let group = change.group_id.as_ref().and_then(|id| {
        plan.groups
            .iter()
            .find(|group| &group.id == id)
            .map(|group| group.title.clone())
    });
    let (before, after) = change_diff(change);
    let show_endpoint = matches!(
        change.kind,
        ChangeKind::Formula | ChangeKind::Cleared | ChangeKind::RowDeleted
    ) || change.cause == ChangeCause::Recalculated;
    let blocked = if apply_status.stale {
        Some((
            warn,
            "Sheet changed since this preview. The proposal stays frozen.".to_string(),
        ))
    } else if let Some(result) = verification_failure {
        Some((error, verification_failure_message(result)))
    } else if let Some(problem) = blocking_problem {
        Some((error, problem.message.clone()))
    } else {
        apply_status
            .disabled_reason()
            .map(|reason| (warn, reason.to_string()))
    };

    let first_empty_col = prepared
        .source_workbook()
        .sheet_by_id(state.source_sheet_id)
        .and_then(|sheet| sheet.cells_iter().map(|(&(_, col), _)| col).max())
        .map_or(0, |col| col.saturating_add(1));
    // Keep one empty grid column between populated data and the card. Text can
    // spill into the first empty column, so merely docking at its left edge can
    // visually clip values in the last populated column.
    let safe_dock_col = first_empty_col.saturating_add(REVIEW_CARD_DATA_GUTTER_COLS);
    let preferred_left = app.metrics.header_w + app.col_x_offset(safe_dock_col) + REVIEW_CARD_GAP;
    let preferred_fits =
        preferred_left + REVIEW_CARD_WIDTH + REVIEW_CARD_GAP <= app.grid_layout.viewport_size.0;
    let active = app.active_cell_rect();
    let fallback_top = if active.y + active.height
        > app.grid_layout.viewport_size.1 - REVIEW_CARD_HEIGHT_ESTIMATE - REVIEW_CARD_GAP
    {
        REVIEW_CARD_GAP
    } else {
        (app.grid_layout.viewport_size.1 - REVIEW_CARD_HEIGHT_ESTIMATE - REVIEW_CARD_GAP)
            .max(REVIEW_CARD_GAP)
    };
    let fallback_left = (app.grid_layout.viewport_size.0 - REVIEW_CARD_WIDTH - REVIEW_CARD_GAP)
        .max(REVIEW_CARD_GAP);
    let automatic_position = if preferred_fits {
        (preferred_left, REVIEW_CARD_GAP)
    } else {
        (fallback_left, fallback_top)
    };
    let (card_left, card_top) = state.card_position().unwrap_or(automatic_position);
    let card_left = card_left.clamp(
        8.0,
        (app.grid_layout.viewport_size.0 - REVIEW_CARD_WIDTH - 8.0).max(8.0),
    );
    let card_top = card_top.clamp(
        8.0,
        (app.grid_layout.viewport_size.1 - REVIEW_CARD_HEIGHT_ESTIMATE - 8.0).max(8.0),
    );
    let card_is_dragging = state.card_is_dragging();
    let grid_top = app.grid_layout.grid_body_origin.1;

    let endpoint = endpoint_control(state.endpoint, proposal, surface, border, muted, cx);
    let previous = nav_button(
        "review-previous-change",
        "←",
        false,
        position > 0,
        proposal,
        muted,
        cx,
    );
    let next = nav_button(
        "review-next-change",
        "→",
        true,
        position + 1 < total,
        proposal,
        muted,
        cx,
    );

    let mut card = div()
        .id("review-card")
        .absolute()
        .w(px(REVIEW_CARD_WIDTH))
        .p(px(12.0))
        .rounded(px(12.0))
        .bg(surface)
        .border_1()
        .border_color(if apply_status.stale {
            warn.opacity(0.65)
        } else if apply_status.blocking {
            error.opacity(0.65)
        } else {
            proposal.opacity(0.26)
        })
        .shadow_lg()
        .left(px(card_left))
        .top(px(card_top))
        .child(
            div()
                .flex()
                .items_start()
                .justify_between()
                .gap(px(10.0))
                .child(
                    div()
                        .id("review-card-drag-handle")
                        .flex_1()
                        .min_w(px(0.0))
                        .cursor(if card_is_dragging {
                            CursorStyle::ClosedHand
                        } else {
                            CursorStyle::OpenHand
                        })
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                                cx.stop_propagation();
                                if let Some(state) = this.review_mode.as_mut() {
                                    state.begin_card_drag(
                                        (event.position.x.into(), event.position.y.into()),
                                        (card_left, card_top),
                                        grid_top,
                                    );
                                    cx.notify();
                                }
                            }),
                        )
                        .child(
                            div()
                                .text_size(px(14.0))
                                .font_weight(FontWeight::BOLD)
                                .text_color(text)
                                .child(title),
                        )
                        .child(
                            div()
                                .mt(px(3.0))
                                .text_size(px(10.0))
                                .text_color(muted)
                                .child(format!(
                                    "{} · {}",
                                    change_address(app, change),
                                    producer_attribution(plan)
                                )),
                        ),
                )
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap(px(6.0))
                        .when(show_endpoint, |row| row.child(endpoint))
                        .child(
                            div()
                                .id("review-collapse")
                                .cursor_pointer()
                                .h(px(26.0))
                                .px(px(8.0))
                                .flex()
                                .items_center()
                                .gap(px(5.0))
                                .rounded(px(7.0))
                                .border_1()
                                .border_color(border.opacity(0.7))
                                .text_size(px(9.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(muted)
                                .hover(|style| style.bg(border.opacity(0.32)).text_color(text))
                                .on_click(cx.listener(|this, _, _, cx| {
                                    if let Some(state) = this.review_mode.as_mut() {
                                        state.toggle_collapsed();
                                        cx.notify();
                                    }
                                }))
                                .child("Collapse")
                                .child(div().text_size(px(8.0)).child("▾")),
                        ),
                ),
        )
        .child(
            div()
                .mt(px(6.0))
                .flex()
                .items_center()
                .gap(px(7.0))
                .child(previous)
                .child(
                    div()
                        .text_size(px(10.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text)
                        .child(format!("Change {} of {total}", position + 1)),
                )
                .child(next),
        )
        .when_some(group, |card, group| {
            card.child(
                div()
                    .mt(px(7.0))
                    .text_size(px(10.0))
                    .font_weight(FontWeight::SEMIBOLD)
                    .text_color(if change.kind == ChangeKind::RowDeleted {
                        error
                    } else {
                        proposal
                    })
                    .child(group),
            )
        })
        .child(hero_diff(
            before,
            after,
            change.kind == ChangeKind::RowDeleted,
            proposal,
            error,
            border,
            text,
            muted,
        ))
        .when_some(change.reason.as_ref(), |card, reason| {
            card.child(
                div()
                    .mt(px(6.0))
                    .text_size(px(10.0))
                    .text_color(muted)
                    .child(reason.0.clone()),
            )
        });

    if let Some((color, message)) = blocked {
        card = card.child(message_line(
            if apply_status.stale {
                MessageIcon::Warning
            } else {
                MessageIcon::Error
            },
            message,
            color,
            border,
        ));
    } else if let Some((icon, color, message)) =
        verification_message(plan.verification.first(), ok, muted, error)
    {
        card = card.child(message_line(icon, message, color, border));
    }

    card.child(
        div()
            .mt(px(10.0))
            .flex()
            .items_center()
            .justify_between()
            .gap(px(8.0))
            .when(is_mcp && !can_apply, |row| {
                row.child(
                    div()
                        .flex_1()
                        .text_size(px(9.0))
                        .text_color(muted)
                        .child("Waiting for the agent to resubmit"),
                )
            })
            .when(!is_mcp && !can_apply, |row| {
                row.child(
                    div()
                        .id("review-repreview")
                        .cursor_pointer()
                        .px(px(11.0))
                        .py(px(7.0))
                        .rounded(px(7.0))
                        .bg(warn)
                        .text_size(px(10.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(inverse)
                        .on_click(
                            cx.listener(|this, _, window, cx| this.preview_last_lua(window, cx)),
                        )
                        .child("Re-preview"),
                )
            })
            .child(
                div()
                    .flex()
                    .items_center()
                    .gap(px(7.0))
                    .child(
                        div()
                            .id("review-dismiss")
                            .cursor_pointer()
                            .px(px(10.0))
                            .py(px(5.0))
                            .rounded(px(7.0))
                            .border_1()
                            .border_color(border)
                            .hover(|style| style.bg(border.opacity(0.28)))
                            .on_click(
                                cx.listener(|this, _, _, cx| this.dismiss_structured_result(cx)),
                            )
                            .child(
                                div()
                                    .flex()
                                    .flex_col()
                                    .items_center()
                                    .child(
                                        div()
                                            .text_size(px(10.0))
                                            .font_weight(FontWeight::SEMIBOLD)
                                            .text_color(text)
                                            .child("Discard plan"),
                                    )
                                    .child(div().text_size(px(8.0)).text_color(muted).child("Esc")),
                            ),
                    )
                    .when(can_apply, |buttons| {
                        buttons.child(
                            div()
                                .id("review-apply")
                                .cursor_pointer()
                                .px(px(10.0))
                                .py(px(5.0))
                                .rounded(px(7.0))
                                .bg(proposal)
                                .hover(|style| style.bg(proposal.opacity(0.84)))
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.apply_lua_to_current_sheet(window, cx)
                                }))
                                .child(
                                    div()
                                        .flex()
                                        .flex_col()
                                        .items_center()
                                        .child(
                                            div()
                                                .text_size(px(10.0))
                                                .font_weight(FontWeight::SEMIBOLD)
                                                .text_color(inverse)
                                                .child(format!("Apply {}", count_label(total))),
                                        )
                                        .child(
                                            div()
                                                .text_size(px(8.0))
                                                .text_color(inverse.opacity(0.76))
                                                .child("Enter"),
                                        ),
                                ),
                        )
                    }),
            ),
    )
    .into_any_element()
}

fn collapsed_pill(
    total: usize,
    proposal: Hsla,
    surface: Hsla,
    text: Hsla,
    muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    div()
        .id("review-pill")
        .absolute()
        .right(px(REVIEW_CARD_GAP))
        .bottom(px(REVIEW_CARD_GAP))
        .flex()
        .items_center()
        .gap(px(8.0))
        .px(px(12.0))
        .h(px(34.0))
        .rounded(px(17.0))
        .bg(surface)
        .border_1()
        .border_color(proposal.opacity(0.48))
        .shadow_lg()
        .cursor_pointer()
        .on_click(cx.listener(|this, _, _, cx| {
            if let Some(state) = this.review_mode.as_mut() {
                state.toggle_collapsed();
                cx.notify();
            }
        }))
        .child(div().size(px(7.0)).rounded(px(4.0)).bg(proposal))
        .child(
            div()
                .text_size(px(10.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text)
                .child(format!("{} proposed", count_label(total))),
        )
        .child(div().text_size(px(9.0)).text_color(muted).child("▲"))
}

fn endpoint_control(
    endpoint: ReviewEndpoint,
    proposal: Hsla,
    surface: Hsla,
    border: Hsla,
    muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let before = endpoint_segment(
        "review-before",
        "Before",
        ReviewEndpoint::Before,
        endpoint == ReviewEndpoint::Before,
        proposal,
        surface,
        muted,
        cx,
    );
    let after = endpoint_segment(
        "review-after",
        "After",
        ReviewEndpoint::After,
        endpoint == ReviewEndpoint::After,
        proposal,
        surface,
        muted,
        cx,
    );
    div()
        .id("review-endpoint")
        .flex()
        .items_center()
        .p(px(2.0))
        .rounded(px(6.0))
        .border_1()
        .border_color(border)
        .bg(border.opacity(0.18))
        .child(before)
        .child(after)
}

fn endpoint_segment(
    id: &'static str,
    label: &'static str,
    value: ReviewEndpoint,
    active: bool,
    proposal: Hsla,
    surface: Hsla,
    muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    div()
        .id(id)
        .cursor_pointer()
        .px(px(6.0))
        .py(px(3.0))
        .rounded(px(4.0))
        .bg(if active {
            surface
        } else {
            gpui::transparent_black()
        })
        .text_size(px(9.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(if active { proposal } else { muted })
        .on_click(cx.listener(move |this, _, _, cx| {
            if let Some(state) = this.review_mode.as_mut() {
                state.set_endpoint(value);
                cx.notify();
            }
        }))
        .child(label)
}

fn nav_button(
    id: &'static str,
    label: &'static str,
    forward: bool,
    enabled: bool,
    proposal: Hsla,
    muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    div()
        .id(id)
        .cursor(if enabled {
            CursorStyle::PointingHand
        } else {
            CursorStyle::Arrow
        })
        .size(px(28.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded(px(7.0))
        .text_size(px(15.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(if enabled {
            proposal
        } else {
            muted.opacity(0.42)
        })
        .when(enabled, |button| {
            button
                .hover(|style| style.bg(proposal.opacity(0.09)))
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(move |this, _, _, cx| {
                        cx.stop_propagation();
                        this.navigate_review_change(forward, false, cx);
                    }),
                )
        })
        .child(label)
}

fn hero_diff(
    before: String,
    after: String,
    deletion: bool,
    proposal: Hsla,
    error: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
) -> Div {
    let after_color = if deletion { error } else { proposal };
    div()
        .mt(px(5.0))
        .flex()
        .items_center()
        .gap(px(8.0))
        .child(
            div()
                .min_w(px(0.0))
                .max_w(px(138.0))
                .overflow_hidden()
                .text_ellipsis()
                .px(px(8.0))
                .py(px(4.0))
                .rounded(px(6.0))
                .bg(border.opacity(0.18))
                .border_1()
                .border_color(border.opacity(0.62))
                .text_size(px(18.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text.opacity(0.82))
                .child(before),
        )
        .child(
            div()
                .text_size(px(19.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(muted)
                .child("→"),
        )
        .child(
            div()
                .min_w(px(0.0))
                .max_w(px(138.0))
                .overflow_hidden()
                .text_ellipsis()
                .px(px(8.0))
                .py(px(4.0))
                .rounded(px(6.0))
                .bg(after_color.opacity(0.11))
                .border_1()
                .border_color(after_color.opacity(0.34))
                .text_size(px(18.0))
                .font_weight(FontWeight::BOLD)
                .text_color(after_color)
                .child(after),
        )
}

#[derive(Clone, Copy)]
enum MessageIcon {
    Passed,
    Warning,
    Error,
    Pending,
}

fn message_line(icon: MessageIcon, message: String, color: Hsla, border: Hsla) -> Div {
    let glyph = match icon {
        MessageIcon::Passed => "✓",
        MessageIcon::Warning => "!",
        MessageIcon::Error => "×",
        MessageIcon::Pending => "·",
    };
    div()
        .mt(px(11.0))
        .pt(px(9.0))
        .flex()
        .items_center()
        .gap(px(6.0))
        .border_t_1()
        .border_color(border)
        .text_size(px(10.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(color)
        .child(
            div()
                .flex_shrink_0()
                .size(px(14.0))
                .rounded(px(7.0))
                .flex()
                .items_center()
                .justify_center()
                .bg(color.opacity(0.12))
                .border_1()
                .border_color(color.opacity(0.38))
                .text_size(px(9.0))
                .font_weight(FontWeight::BOLD)
                .child(glyph),
        )
        .child(message)
}

fn producer_attribution(plan: &visigrid_engine::operation_plan::OperationPlan) -> String {
    let via = plan
        .producer
        .source_path
        .as_deref()
        .and_then(|path| Path::new(path).file_name())
        .and_then(|name| name.to_str())
        .map(ToOwned::to_owned)
        .unwrap_or_else(|| {
            if plan.producer.kind.starts_with("mcp") {
                "MCP".into()
            } else {
                "Lua".into()
            }
        });
    format!("{} via {via}", plan.producer.name)
}

fn count_label(count: usize) -> String {
    if count == 1 {
        "1 change".into()
    } else {
        format!("{count} changes")
    }
}

fn change_address(app: &Spreadsheet, change: &MaterializedChange) -> String {
    if change.kind == ChangeKind::RowDeleted {
        let row = change
            .before_coordinate
            .map_or(1, |coordinate| coordinate.row + 1);
        let last_col = app
            .review_source_sheet(change.sheet_id)
            .and_then(|sheet| sheet.cells_iter().map(|(&(_, col), _)| col).max())
            .unwrap_or(0);
        return format!("Row {row} · A{row}:{}{}", column_letters(last_col), row);
    }
    let Some(coordinate) = change.before_coordinate.or(change.after_coordinate) else {
        return "Change".into();
    };
    let address = format!("{}{}", column_letters(coordinate.col), coordinate.row + 1);
    let header = app
        .review_source_sheet(change.sheet_id)
        .map(|sheet| sheet.get_formatted_display(0, coordinate.col))
        .filter(|header| !header.trim().is_empty());
    match header {
        Some(header) if coordinate.row > 0 => format!("{address} · {header}"),
        _ => address,
    }
}

fn change_diff(change: &MaterializedChange) -> (String, String) {
    let before = if change.kind == ChangeKind::Formula {
        change.before.raw.clone()
    } else {
        visible_text_value(&change.before.raw, &change.before.display)
    };
    let after = if change.kind == ChangeKind::Formula {
        change.after.raw.clone()
    } else {
        visible_text_value(&change.after.raw, &change.after.display)
    };
    (
        display_or_empty(&before),
        if change.kind == ChangeKind::RowDeleted {
            "Deleted row".into()
        } else {
            display_or_empty(&after)
        },
    )
}

fn visible_text_value(raw: &str, display: &str) -> String {
    if raw.trim() == raw {
        return display.to_string();
    }
    let visible = raw
        .replace('\t', "\\t")
        .replace('\r', "\\r")
        .replace('\n', "\\n")
        .replace(' ', "·");
    format!("\"{visible}\"")
}

fn verification_message(
    result: Option<&VerificationResult>,
    ok: Hsla,
    muted: Hsla,
    error: Hsla,
) -> Option<(MessageIcon, Hsla, String)> {
    let result = result?;
    let label = result.label.clone().unwrap_or_else(|| result.id.clone());
    match result.status {
        VerificationStatus::Passed => Some((MessageIcon::Passed, ok, label)),
        VerificationStatus::Failed | VerificationStatus::Unknown => Some((
            MessageIcon::Error,
            error,
            verification_failure_message(result),
        )),
        VerificationStatus::NotRun => {
            Some((MessageIcon::Pending, muted, format!("{label} not run")))
        }
    }
}

fn verification_failure_message(result: &VerificationResult) -> String {
    let label = result.label.clone().unwrap_or_else(|| result.id.clone());
    match &result.evidence {
        VerificationEvidence::RetainedTotal {
            expected,
            actual,
            currency,
            difference,
            ..
        } => {
            let symbol = currency_symbol(currency);
            format!("{label}: off by {symbol}{difference}. Expected {symbol}{expected}; preview is {symbol}{actual}.")
        }
        VerificationEvidence::NoNewFormulaErrors { new_errors } => {
            format!("{label}: {new_errors} new formula errors")
        }
        VerificationEvidence::Unavailable { message } => format!("{label}: {message}"),
    }
}

fn currency_symbol(currency: &str) -> &str {
    match currency {
        "USD" => "$",
        "EUR" => "€",
        "GBP" => "£",
        "JPY" => "¥",
        _ => "",
    }
}

fn display_or_empty(value: &str) -> String {
    if value.is_empty() {
        "(empty)".into()
    } else {
        value.to_string()
    }
}

fn column_letters(col: usize) -> String {
    let mut letters = String::new();
    let mut col = col + 1;
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    letters
}

#[cfg(test)]
mod tests {
    use super::{column_letters, count_label, currency_symbol, visible_text_value};

    #[test]
    fn review_copy_uses_correct_singular_and_plural() {
        assert_eq!(count_label(1), "1 change");
        assert_eq!(count_label(12), "12 changes");
    }

    #[test]
    fn review_addresses_and_currency_evidence_are_human_readable() {
        assert_eq!(column_letters(0), "A");
        assert_eq!(column_letters(25), "Z");
        assert_eq!(column_letters(26), "AA");
        assert_eq!(currency_symbol("USD"), "$");
        assert_eq!(currency_symbol("EUR"), "€");
    }

    #[test]
    fn review_diff_makes_edge_whitespace_visible() {
        assert_eq!(visible_text_value(" grace ", "grace"), "\"·grace·\"");
        assert_eq!(visible_text_value("Grace", "Grace"), "Grace");
        assert_eq!(visible_text_value("Grace\t", "Grace"), "\"Grace\\t\"");
    }
}
