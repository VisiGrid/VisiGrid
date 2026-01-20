//! Impact Preview modal for showing rename/delete consequences before applying

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// The type of impact action being previewed
#[derive(Clone, Debug, PartialEq)]
pub enum ImpactAction {
    Rename { old_name: String, new_name: String },
    Delete { name: String },
}

impl ImpactAction {
    pub fn name(&self) -> &str {
        match self {
            ImpactAction::Rename { old_name, .. } => old_name,
            ImpactAction::Delete { name } => name,
        }
    }
}

/// A formula that will be affected by the action
#[derive(Clone, Debug)]
pub struct ImpactedFormula {
    pub cell_ref: String,
    pub formula: String,
}

/// Render the impact preview modal
pub fn render_impact_preview(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error = app.token(TokenKey::Error);

    let action = app.impact_preview_action.clone();
    let usages = app.impact_preview_usages.clone();
    let usage_count = usages.len();

    // Build title and subtitle based on action type
    let (title, subtitle, is_delete) = match &action {
        Some(ImpactAction::Rename { old_name, new_name }) => (
            format!("Rename \"{}\" → \"{}\"", old_name, new_name),
            if usage_count == 0 {
                "No formulas reference this name".to_string()
            } else if usage_count == 1 {
                "This will update 1 formula".to_string()
            } else {
                format!("This will update {} formulas", usage_count)
            },
            false,
        ),
        Some(ImpactAction::Delete { name }) => (
            format!("Delete \"{}\"", name),
            if usage_count == 0 {
                "No formulas reference this name".to_string()
            } else if usage_count == 1 {
                "Warning: 1 formula will show #NAME? error".to_string()
            } else {
                format!("Warning: {} formulas will show #NAME? error", usage_count)
            },
            true,
        ),
        None => ("Preview".to_string(), "No action".to_string(), false),
    };

    let has_usages = !usages.is_empty();
    let subtitle_color = if is_delete && has_usages { error } else { text_muted };

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
                .w(px(480.0))
                .max_h(px(400.0))
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
                        .flex_col()
                        .gap_1()
                        // Title
                        .child(
                            div()
                                .text_color(text_primary)
                                .text_base()
                                .font_weight(FontWeight::SEMIBOLD)
                                .child(title)
                        )
                        // Subtitle
                        .child(
                            div()
                                .text_color(subtitle_color)
                                .text_sm()
                                .child(subtitle)
                        )
                )
                // Scrollable list of affected formulas
                .when(has_usages, |d| {
                    d.child(
                        div()
                            .flex_1()
                            .overflow_hidden()
                            .max_h(px(200.0))
                            .px_5()
                            .py_2()
                            .flex()
                            .flex_col()
                            .gap_1()
                            .children(
                                usages.iter().take(50).map(|usage| {
                                    div()
                                        .flex()
                                        .items_center()
                                        .gap_2()
                                        .py_1()
                                        // Cell reference
                                        .child(
                                            div()
                                                .text_color(accent)
                                                .text_sm()
                                                .font_weight(FontWeight::MEDIUM)
                                                .min_w(px(50.0))
                                                .child(usage.cell_ref.clone())
                                        )
                                        // Formula
                                        .child(
                                            div()
                                                .text_color(text_muted)
                                                .text_sm()
                                                .overflow_hidden()
                                                .text_ellipsis()
                                                .child(usage.formula.clone())
                                        )
                                })
                            )
                    )
                })
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
                        // Undo reassurance
                        .child(
                            div()
                                .text_color(text_muted)
                                .text_xs()
                                .child("This change can be undone.")
                        )
                        // Buttons
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                // Cancel button
                                .child(
                                    div()
                                        .id("impact-cancel")
                                        .px_4()
                                        .py_1()
                                        .rounded_md()
                                        .text_sm()
                                        .cursor_pointer()
                                        .text_color(text_muted)
                                        .hover(|s| s.bg(panel_border.opacity(0.5)))
                                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                            this.hide_impact_preview(cx);
                                        }))
                                        .child("Cancel")
                                )
                                // Apply button
                                .child(
                                    div()
                                        .id("impact-apply")
                                        .px_4()
                                        .py_1()
                                        .rounded_md()
                                        .text_sm()
                                        .cursor_pointer()
                                        .bg(if is_delete && has_usages { error } else { accent })
                                        .text_color(hsla(0.0, 0.0, 1.0, 1.0))
                                        .hover(|s| s.opacity(0.85))
                                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                            this.apply_impact_preview(cx);
                                        }))
                                        .child(if is_delete { "Delete" } else { "Apply" })
                                )
                        )
                )
        )
}
