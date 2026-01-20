//! Refactor Log modal - audit trail for named range operations

use gpui::*;
use gpui::prelude::FluentBuilder;
use std::time::SystemTime;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// A single refactor log entry
#[derive(Clone, Debug)]
pub struct RefactorLogEntry {
    pub timestamp: SystemTime,
    pub action: String,
    pub details: String,
    pub impact: Option<String>,
}

impl RefactorLogEntry {
    pub fn new(action: impl Into<String>, details: impl Into<String>) -> Self {
        Self {
            timestamp: SystemTime::now(),
            action: action.into(),
            details: details.into(),
            impact: None,
        }
    }

    pub fn with_impact(mut self, impact: impl Into<String>) -> Self {
        self.impact = Some(impact.into());
        self
    }

    /// Format timestamp as HH:MM:SS
    fn format_time(&self) -> String {
        let duration = self.timestamp
            .duration_since(SystemTime::UNIX_EPOCH)
            .unwrap_or_default();
        let secs = duration.as_secs();
        // Convert to local time (approximate - just show relative time)
        let hours = (secs / 3600) % 24;
        let mins = (secs / 60) % 60;
        let seconds = secs % 60;
        format!("{:02}:{:02}:{:02}", hours, mins, seconds)
    }
}

/// Render the refactor log modal
pub fn render_refactor_log(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    let entries = app.refactor_log.clone();
    let is_empty = entries.is_empty();

    // Centered dialog overlay
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .w(px(520.0))
                .max_h(px(450.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Header
                .child(
                    div()
                        .px_5()
                        .py_3()
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_base()
                                .font_weight(FontWeight::SEMIBOLD)
                                .child("Refactor Log")
                        )
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child(format!("{} entries", entries.len()))
                        )
                )
                // Log entries
                .child(
                    div()
                        .flex_1()
                        .overflow_hidden()
                        .max_h(px(320.0))
                        .when(is_empty, |d| {
                            d.child(
                                div()
                                    .flex()
                                    .flex_col()
                                    .items_center()
                                    .justify_center()
                                    .py_8()
                                    .gap_2()
                                    .child(
                                        div()
                                            .text_color(text_muted)
                                            .text_sm()
                                            .child("No refactoring operations yet.")
                                    )
                                    .child(
                                        div()
                                            .text_color(text_muted.opacity(0.7))
                                            .text_xs()
                                            .child("Create, rename, or delete named ranges to see activity here.")
                                    )
                            )
                        })
                        .when(!is_empty, |d| {
                            d.child(
                                div()
                                    .flex()
                                    .flex_col()
                                    .px_5()
                                    .py_2()
                                    .gap_2()
                                    .children(
                                        entries.iter().rev().take(50).map(|entry| {
                                            div()
                                                .flex()
                                                .flex_col()
                                                .py_2()
                                                .border_b_1()
                                                .border_color(panel_border.opacity(0.5))
                                                // Timestamp and action
                                                .child(
                                                    div()
                                                        .flex()
                                                        .items_center()
                                                        .justify_between()
                                                        .child(
                                                            div()
                                                                .text_color(text_primary)
                                                                .text_sm()
                                                                .font_weight(FontWeight::MEDIUM)
                                                                .child(entry.action.clone())
                                                        )
                                                        .child(
                                                            div()
                                                                .text_color(text_muted)
                                                                .text_xs()
                                                                .child(entry.format_time())
                                                        )
                                                )
                                                // Details
                                                .child(
                                                    div()
                                                        .text_color(accent)
                                                        .text_sm()
                                                        .child(entry.details.clone())
                                                )
                                                // Impact (if any)
                                                .when(entry.impact.is_some(), |d| {
                                                    d.child(
                                                        div()
                                                            .text_color(text_muted)
                                                            .text_xs()
                                                            .mt_1()
                                                            .child(entry.impact.clone().unwrap_or_default())
                                                    )
                                                })
                                        })
                                    )
                            )
                        })
                )
                // Footer
                .child(
                    div()
                        .px_5()
                        .py_3()
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child("Log resets when you close the file.")
                        )
                        .child(
                            div()
                                .id("refactor-log-close")
                                .px_4()
                                .py_1()
                                .rounded_md()
                                .text_sm()
                                .cursor_pointer()
                                .bg(accent)
                                .text_color(hsla(0.0, 0.0, 1.0, 1.0))
                                .hover(|s| s.opacity(0.85))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_refactor_log(cx);
                                }))
                                .child("Close")
                        )
                )
        )
}
