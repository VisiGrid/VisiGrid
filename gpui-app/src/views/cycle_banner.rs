//! Dismissible top banner for cycle/freeze/iteration status.
//!
//! Follows the same pattern as `rewind_dialogs::render_rewind_success_banner`:
//! absolute top-center, panel bg, border, shadow, dismiss X.

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::Button;

/// Render the cycle status banner (top-center, dismissible).
///
/// Three branches based on current mode:
/// - **Strict**: Red border, "#CYCLE!" cells, offers iteration + freeze buttons
/// - **Freeze**: Amber border, frozen values warning
/// - **Iteration**: Green (converged) or amber (not converged)
pub(crate) fn render_cycle_banner(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let text_inverse = app.token(TokenKey::TextInverse);
    let ok_color = app.token(TokenKey::Ok);
    let warn_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);

    let iterative = app.wb(cx).iterative_enabled();
    let freeze = app.import_result.as_ref().map_or(false, |r| r.freeze_applied);
    let cycle_count = app.current_cycle_count(cx);
    let converged = app.last_recalc_report.as_ref().map_or(false, |r| r.converged);
    let scc_count = app.last_recalc_report.as_ref().map_or(0, |r| r.scc_count);
    let iters = app.last_recalc_report.as_ref().map_or(0, |r| r.iterations_performed);
    let is_xlsx = app.current_file.as_ref()
        .and_then(|p| p.extension()).and_then(|e| e.to_str())
        .map_or(false, |e| matches!(e.to_lowercase().as_str(), "xlsx" | "xls" | "xlsm" | "xlsb" | "ods"));
    let frozen_count = app.import_result.as_ref().map_or(0, |r| r.cycles_frozen);

    // Determine branch: iteration > freeze > strict
    // Iteration wins when both iteration and freeze are active (frozen cells are
    // just static constants in iteration mode, not part of the iterative solve).
    let (border_color, title, body) = if iterative && cycle_count > 0 {
        // Branch C: Iteration active
        let frozen_note = if freeze && frozen_count > 0 {
            format!(" Also contains {} frozen cells.", frozen_count)
        } else {
            String::new()
        };
        if converged {
            (
                ok_color,
                "Iterative calculation enabled",
                format!("{} cycle groups resolved \u{2014} converged in {} iterations.{}", scc_count, iters, frozen_note),
            )
        } else {
            (
                warn_color,
                "Iterative calculation enabled",
                format!("Cycles did not converge \u{2014} cells show #NUM!.{}", frozen_note),
            )
        }
    } else if freeze {
        // Branch B: Freeze active (iteration off)
        (
            warn_color,
            "Cycle values frozen",
            format!("{} cells frozen to Excel cached values \u{2014} edits to their inputs won\u{2019}t update them.", frozen_count),
        )
    } else {
        // Branch A: Strict mode (unresolved cycles)
        (
            error_color,
            "Circular references detected",
            format!("{} cells show #CYCLE!. You can resolve them or keep them as-is.", cycle_count),
        )
    };

    let is_strict = !iterative && !freeze && cycle_count > 0;
    // Only show freeze buttons when freeze is active AND iteration is NOT
    // (iteration subsumes freeze — frozen cells are just constants).
    let is_freeze = freeze && !iterative;

    div()
        .absolute()
        .top_0()
        .left_0()
        .right_0()
        .flex()
        .justify_center()
        .pt_2()
        .child(
            div()
                .id("cycle-banner")
                .max_w(px(560.0))
                .bg(panel_bg)
                .border_1()
                .border_color(border_color.opacity(0.5))
                .rounded_md()
                .shadow_lg()
                .px_4()
                .py_2()
                .flex()
                .flex_col()
                .gap_2()
                // Title row with dismiss X
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(13.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child(title)
                        )
                        .child(
                            div()
                                .id("dismiss-cycle-banner-btn")
                                .px_2()
                                .py_1()
                                .text_size(px(14.0))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.text_color(text_primary))
                                .child("\u{2715}") // X mark
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.dismiss_cycle_banner(cx);
                                }))
                        )
                )
                // Body text
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(body)
                )
                // Action buttons row
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        // Strict mode: primary "Turn on iterative calculation..." + secondary "Freeze"
                        .when(is_strict, |d| {
                            d.child(
                                Button::new("banner-enable-iter-btn", "Turn on iterative calculation\u{2026}")
                                    .primary(accent, text_inverse)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.enable_iteration_and_recalc(cx);
                                    }))
                            )
                            .child({
                                let can_freeze = is_xlsx;
                                let btn = Button::new("banner-freeze-btn", "Freeze cycle values")
                                    .disabled(!can_freeze)
                                    .secondary(panel_border, text_primary);
                                if can_freeze {
                                    btn.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.dismiss_cycle_banner(cx);
                                        this.reimport_with_freeze(cx);
                                    }))
                                } else {
                                    btn
                                }
                            })
                        })
                        // Freeze mode: secondary "Turn on iterative calculation..."
                        .when(is_freeze, |d| {
                            d.child(
                                Button::new("banner-enable-iter-btn2", "Turn on iterative calculation\u{2026}")
                                    .secondary(panel_border, text_primary)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.enable_iteration_and_recalc(cx);
                                    }))
                            )
                        })
                        // "View details" link (all branches)
                        .child(
                            div()
                                .id("banner-view-details")
                                .text_size(px(11.0))
                                .text_color(accent)
                                .cursor_pointer()
                                .hover(|s| s.underline())
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.dismiss_cycle_banner(cx);
                                    this.show_import_report(cx);
                                }))
                                .child("View details")
                        )
                )
        )
}
