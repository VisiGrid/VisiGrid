use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::ui::render_locked_feature_panel;

pub const PANEL_WIDTH: f32 = 280.0;

/// Render the profiler panel (right-side drawer).
pub fn render_profiler_panel(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);

    div()
        .id("profiler-panel")
        .absolute()
        .right_0()
        .top_0()
        .bottom_0()
        .w(px(PANEL_WIDTH))
        .bg(panel_bg)
        .border_l_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .overflow_hidden()
        // Capture mouse events to prevent click-through to grid and backdrop
        .on_mouse_down(MouseButton::Left, |_, _, cx| {
            cx.stop_propagation();
        })
        .on_mouse_up(MouseButton::Left, |_, _, cx| {
            cx.stop_propagation();
        })
        // Header
        .child(render_profiler_header(panel_bg, panel_border, text_primary, text_muted, cx))
        // Scrollable content
        .child(
            div()
                .id("profiler-scroll")
                .flex_1()
                .overflow_y_scroll()
                .child(render_profiler_content(app, cx))
        )
}

/// Render the profiler panel header bar.
fn render_profiler_header(
    _panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .px_3()
        .py_2()
        .flex()
        .items_center()
        .justify_between()
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .text_size(px(13.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child("Performance Profiler")
        )
        .child(
            div()
                .id("profiler-close-btn")
                .px_2()
                .py_1()
                .cursor_pointer()
                .text_color(text_muted)
                .text_size(px(12.0))
                .hover(|s| s.text_color(text_primary))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                    this.profiler_visible = false;
                    window.focus(&this.focus_handle, cx);
                    cx.notify();
                }))
                .child("\u{2715}")  // ✕
        )
}

/// Render the profiler content sections.
fn render_profiler_content(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let is_pro = visigrid_license::is_feature_enabled("performance");
    let has_report = app.profiler_report.is_some();

    let mut content = div()
        .flex()
        .flex_col()
        .p_3()
        .gap_3();

    // Summary section (always visible — Free + Pro)
    content = content.child(render_summary_section(app, cx));

    if has_report {
        // Phase Timing section (Pro gated)
        content = content.child(render_phase_timing_section(app, is_pro, cx));

        // Hotspot Suspects section (Pro gated)
        content = content.child(render_hotspot_section(app, is_pro, cx));

        // Cycle Analysis section (Pro gated, conditional)
        if let Some(ref report) = app.profiler_report {
            if report.had_cycles {
                content = content.child(render_cycle_section(app, is_pro, cx));
            }
        }
    }

    content
}

/// Summary section — always visible for Free + Pro users.
fn render_summary_section(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    let mut section = div()
        .flex()
        .flex_col()
        .gap_2();

    // Section header
    section = section.child(
        div()
            .text_size(px(11.0))
            .font_weight(FontWeight::SEMIBOLD)
            .text_color(text_muted)
            .child("SUMMARY")
    );

    if let Some(ref report) = app.profiler_report {
        // Report summary line
        let duration_str = if report.duration_ms == 0 {
            "<1ms".to_string()
        } else {
            format!("{}ms", report.duration_ms)
        };
        let summary = format!(
            "{} \u{00B7} {} cells \u{00B7} depth {}",
            duration_str, report.cells_recomputed, report.max_depth
        );
        section = section.child(
            div()
                .text_size(px(12.0))
                .text_color(text_primary)
                .font_weight(FontWeight::MEDIUM)
                .child(SharedString::from(summary))
        );

        // Cycle status
        if report.had_cycles {
            let cycle_msg = if report.converged {
                format!("{} cycle cells (resolved)", report.cycle_cells)
            } else {
                format!("{} cycle cells (unresolved)", report.cycle_cells)
            };
            section = section.child(
                div()
                    .text_size(px(11.0))
                    .text_color(app.token(TokenKey::Warn))
                    .child(SharedString::from(cycle_msg))
            );
        }

        // Classification diagnosis — answers "why is this slow?"
        let classification = classify_report(report);
        let classification_color = match classification {
            Classification::Fast => app.token(TokenKey::Ok),
            Classification::ScaleBound => text_muted,
            Classification::StructureBound => app.token(TokenKey::Warn),
            Classification::LuaBound => accent,
        };
        section = section.child(
            div()
                .mt_1()
                .px_2()
                .py(px(4.0))
                .rounded_sm()
                .bg(classification_color.opacity(0.1))
                .border_l_2()
                .border_color(classification_color.opacity(0.6))
                .text_size(px(11.0))
                .text_color(text_primary)
                .child(SharedString::from(classification.label()))
        );
    } else {
        // No report
        section = section.child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child("No profiling data yet.")
        );
        section = section.child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("Click \"Recalc Now\" to profile.")
        );
    }

    // Action buttons
    let capture_next = app.profiler_capture_next;
    section = section.child(
        div()
            .flex()
            .flex_col()
            .gap_1()
            .mt_1()
            // Recalc Now button
            .child(
                div()
                    .id("profiler-recalc-btn")
                    .px_3()
                    .py(px(5.0))
                    .bg(accent)
                    .rounded_sm()
                    .text_size(px(11.0))
                    .text_color(rgb(0xffffff))
                    .cursor_pointer()
                    .hover(|s| s.bg(accent.opacity(0.85)))
                    .flex()
                    .items_center()
                    .justify_center()
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        this.profile_next_recalc(cx);
                    }))
                    .child("Recalc Now")
            )
            // Capture Next Recalc toggle (Pro)
            .child(
                div()
                    .id("profiler-capture-toggle")
                    .px_2()
                    .py(px(4.0))
                    .rounded_sm()
                    .border_1()
                    .border_color(if capture_next { accent } else { panel_border })
                    .bg(if capture_next { accent.opacity(0.1) } else { panel_border.opacity(0.1) })
                    .text_size(px(10.0))
                    .text_color(if capture_next { accent } else { text_muted })
                    .cursor_pointer()
                    .hover(|s| s.bg(panel_border.opacity(0.2)))
                    .flex()
                    .items_center()
                    .gap_1()
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        this.profiler_capture_next = !this.profiler_capture_next;
                        cx.notify();
                    }))
                    .child(if capture_next { "\u{25C9}" } else { "\u{25CB}" })
                    .child("Capture next recalc")
            )
            // Clear button (only when report exists)
            .when(app.profiler_report.is_some(), |d| {
                d.child(
                    div()
                        .id("profiler-clear-btn")
                        .px_2()
                        .py(px(3.0))
                        .rounded_sm()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .cursor_pointer()
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.profiler_report = None;
                            this.profiler_hotspots = Vec::new();
                            cx.notify();
                        }))
                        .child("Clear")
                )
            })
    );

    section
}

/// Format microseconds for display.
fn format_us(us: u64) -> String {
    if us == 0 {
        "<0.1ms".to_string()
    } else if us < 1000 {
        format!("{:.1}ms", us as f64 / 1000.0)
    } else if us < 1_000_000 {
        format!("{:.1}ms", us as f64 / 1000.0)
    } else {
        format!("{:.1}s", us as f64 / 1_000_000.0)
    }
}

/// Phase Timing section — Pro gated.
fn render_phase_timing_section(
    app: &mut Spreadsheet,
    is_pro: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    if !is_pro {
        let text_inverse = app.token(TokenKey::TextInverse);
        // Skeleton preview for locked panel
        let preview = div()
            .flex()
            .flex_col()
            .gap_1()
            .child(div().h(px(8.0)).w(px(120.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(8.0)).w(px(80.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(8.0)).w(px(160.0)).rounded_sm().bg(panel_border.opacity(0.3)));

        return match render_locked_feature_panel(
            "Phase Timing",
            "See where recalc time is spent: invalidation, topo sort, evaluation, and Lua functions.",
            preview.into_any_element(),
            app.locked_panels_dismissed,
            panel_border,
            text_primary,
            text_muted,
            accent,
            text_inverse,
            cx,
        ) {
            Some(el) => div().child(el),
            None => div(),
        };
    }

    let report = match &app.profiler_report {
        Some(r) => r,
        None => return div(),
    };

    let total_us = report.phase_invalidation_us + report.phase_topo_sort_us + report.phase_eval_us;
    let max_us = total_us.max(1); // Avoid division by zero

    let phases: Vec<(&str, u64, Hsla)> = {
        let mut v = vec![
            ("Invalidation", report.phase_invalidation_us, hsla(0.6, 0.6, 0.5, 1.0)),
            ("Topo Sort", report.phase_topo_sort_us, hsla(0.3, 0.6, 0.5, 1.0)),
            ("Evaluation", report.phase_eval_us, hsla(0.08, 0.7, 0.5, 1.0)),
        ];
        if report.phase_lua_total_us > 0 {
            v.push(("Lua Functions", report.phase_lua_total_us, hsla(0.75, 0.6, 0.5, 1.0)));
        }
        v
    };

    let mut section = div()
        .flex()
        .flex_col()
        .gap_2();

    section = section.child(
        div()
            .text_size(px(11.0))
            .font_weight(FontWeight::SEMIBOLD)
            .text_color(text_muted)
            .child("PHASE TIMING")
    );

    for (label, us, color) in &phases {
        let pct = if *us == 0 { 0.0 } else { *us as f32 / max_us as f32 };
        let bar_width = (pct * 180.0).max(2.0); // Min 2px for visibility

        section = section.child(
            div()
                .flex()
                .flex_col()
                .gap(px(2.0))
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(10.0))
                                .text_color(text_primary)
                                .child(*label)
                        )
                        .child(
                            div()
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format_us(*us)))
                        )
                )
                .child(
                    div()
                        .h(px(4.0))
                        .w_full()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.2))
                        .child(
                            div()
                                .h(px(4.0))
                                .w(px(bar_width))
                                .rounded_sm()
                                .bg(*color)
                        )
                )
        );
    }

    section
}

/// Hotspot Suspects section — Pro gated.
fn render_hotspot_section(
    app: &mut Spreadsheet,
    is_pro: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    if !is_pro {
        let text_inverse = app.token(TokenKey::TextInverse);
        let preview = div()
            .flex()
            .flex_col()
            .gap_1()
            .child(div().h(px(14.0)).w(px(200.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(14.0)).w(px(180.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(14.0)).w(px(160.0)).rounded_sm().bg(panel_border.opacity(0.3)));

        return match render_locked_feature_panel(
            "Hotspot Suspects",
            "Find cells with high fan-out, deep dependency chains, or dynamic references that slow recalc.",
            preview.into_any_element(),
            app.locked_panels_dismissed,
            panel_border,
            text_primary,
            text_muted,
            accent,
            text_inverse,
            cx,
        ) {
            Some(el) => div().child(el),
            None => div(),
        };
    }

    let hotspots = &app.profiler_hotspots;

    let mut section = div()
        .flex()
        .flex_col()
        .gap_2();

    let header_text = if hotspots.is_empty() {
        "HOTSPOT SUSPECTS".to_string()
    } else {
        format!("HOTSPOT SUSPECTS ({})", hotspots.len())
    };

    section = section.child(
        div()
            .text_size(px(11.0))
            .font_weight(FontWeight::SEMIBOLD)
            .text_color(text_muted)
            .child(SharedString::from(header_text))
    );

    if hotspots.is_empty() {
        section = section.child(
            div()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("No hotspot suspects found.")
        );
        return section;
    }

    // Build hotspot entries — collect data first to avoid borrow issues
    let entries: Vec<(String, f64, usize, usize, usize, bool, String, usize, usize, usize)> = hotspots.iter().map(|entry| {
        let cell_ref = format!("{}", entry.cell);
        let wb = app.workbook.read(cx);
        let sheet_index = wb.sheet_index_by_id(entry.cell.sheet).unwrap_or(0);
        // Get formula preview
        let formula_preview = wb.sheet_by_id(entry.cell.sheet)
            .map(|s| {
                let raw = s.get_raw(entry.cell.row, entry.cell.col);
                if raw.len() > 40 {
                    format!("{}...", &raw[..40])
                } else {
                    raw
                }
            })
            .unwrap_or_default();
        (
            cell_ref,
            entry.score,
            entry.fan_in,
            entry.fan_out,
            entry.depth,
            entry.has_unknown_deps,
            formula_preview,
            sheet_index,
            entry.cell.row,
            entry.cell.col,
        )
    }).collect();

    for (i, (cell_ref, score, fan_in, fan_out, depth, has_unknown, formula_preview, sheet_index, row, col)) in entries.into_iter().enumerate() {
        section = section.child(render_hotspot_entry(
            i,
            cell_ref,
            score,
            fan_in,
            fan_out,
            depth,
            has_unknown,
            formula_preview,
            sheet_index,
            row,
            col,
            panel_border,
            text_primary,
            text_muted,
            accent,
            cx,
        ));
    }

    section
}

/// Render a single hotspot entry.
fn render_hotspot_entry(
    index: usize,
    cell_ref: String,
    score: f64,
    fan_in: usize,
    fan_out: usize,
    depth: usize,
    has_unknown: bool,
    formula_preview: String,
    sheet_index: usize,
    row: usize,
    col: usize,
    panel_border: Hsla,
    _text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let id = SharedString::from(format!("hotspot-{}", index));
    let cell_ref_shared: SharedString = cell_ref.into();
    let stats_line = format!("in:{} out:{} depth:{}", fan_in, fan_out, depth);
    let stats_shared: SharedString = stats_line.into();
    let score_str: SharedString = format!("{:.0}", score).into();

    let mut entry = div()
        .id(id)
        .px_2()
        .py(px(4.0))
        .rounded_sm()
        .bg(panel_border.opacity(0.1))
        .cursor_pointer()
        .hover(|s| s.bg(panel_border.opacity(0.2)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // Navigate to the cell
            let current_sheet = this.workbook.read(cx).active_sheet_index();
            if current_sheet != sheet_index {
                this.goto_sheet(sheet_index, cx);
            }
            this.select_cell(row, col, false, cx);
        }))
        .flex()
        .flex_col()
        .gap(px(2.0))
        // Top row: cell ref + score badge
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(11.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(accent)
                        .child(cell_ref_shared)
                )
                .child(
                    div()
                        .px(px(4.0))
                        .py(px(1.0))
                        .rounded(px(6.0))
                        .bg(accent.opacity(0.15))
                        .text_size(px(9.0))
                        .text_color(accent)
                        .child(score_str)
                )
                .when(has_unknown, |d| {
                    d.child(
                        div()
                            .text_size(px(10.0))
                            .text_color(app_warn_color())
                            .child("\u{26A0}")  // ⚠
                    )
                })
        )
        // Stats line
        .child(
            div()
                .text_size(px(9.0))
                .text_color(text_muted)
                .child(stats_shared)
        );

    // Formula preview
    if !formula_preview.is_empty() {
        let preview_shared: SharedString = formula_preview.into();
        entry = entry.child(
            div()
                .text_size(px(9.0))
                .text_color(text_muted.opacity(0.7))
                .overflow_hidden()
                .child(preview_shared)
        );
    }

    entry
}

/// Fallback warn color (used in hotspot entries for unknown deps warning).
fn app_warn_color() -> Hsla {
    hsla(0.1, 0.8, 0.5, 1.0)
}

/// Cycle Analysis section — Pro gated, only shown if cycles detected.
fn render_cycle_section(
    app: &mut Spreadsheet,
    is_pro: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);

    let report = match &app.profiler_report {
        Some(r) => r,
        None => return div(),
    };

    if !is_pro {
        let accent = app.token(TokenKey::Accent);
        let text_inverse = app.token(TokenKey::TextInverse);
        let preview = div()
            .flex()
            .flex_col()
            .gap_1()
            .child(div().h(px(10.0)).w(px(140.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(10.0)).w(px(100.0)).rounded_sm().bg(panel_border.opacity(0.3)))
            .child(div().h(px(10.0)).w(px(120.0)).rounded_sm().bg(panel_border.opacity(0.3)));

        return match render_locked_feature_panel(
            "Cycle Analysis",
            "See SCC count, iteration depth, and convergence status for circular references.",
            preview.into_any_element(),
            app.locked_panels_dismissed,
            panel_border,
            text_primary,
            text_muted,
            accent,
            text_inverse,
            cx,
        ) {
            Some(el) => div().child(el),
            None => div(),
        };
    }

    let mut section = div()
        .flex()
        .flex_col()
        .gap_2();

    section = section.child(
        div()
            .text_size(px(11.0))
            .font_weight(FontWeight::SEMIBOLD)
            .text_color(text_muted)
            .child("CYCLE ANALYSIS")
    );

    let convergence_str = if report.converged {
        "Converged"
    } else {
        "Did not converge"
    };

    let convergence_color = if report.converged {
        app.token(TokenKey::Ok)
    } else {
        app.token(TokenKey::Error)
    };

    let details = vec![
        ("Cycle cells", format!("{}", report.cycle_cells)),
        ("SCCs", format!("{}", report.scc_count)),
        ("Max iterations", format!("{}", report.iterations_performed)),
    ];

    for (label, value) in details {
        section = section.child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(label)
                )
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_primary)
                        .child(SharedString::from(value))
                )
        );
    }

    // Convergence status
    section = section.child(
        div()
            .flex()
            .items_center()
            .justify_between()
            .child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted)
                    .child("Status")
            )
            .child(
                div()
                    .text_size(px(10.0))
                    .font_weight(FontWeight::MEDIUM)
                    .text_color(convergence_color)
                    .child(convergence_str)
            )
    );

    section
}

// ============================================================================
// Classification heuristic
// ============================================================================

enum Classification {
    Fast,
    ScaleBound,
    StructureBound,
    LuaBound,
}

impl Classification {
    fn label(&self) -> &'static str {
        match self {
            Self::Fast => "Fast — no bottleneck detected",
            Self::LuaBound => "Likely Lua-bound — custom functions dominate eval time",
            Self::StructureBound => "Likely structure-bound — deep chains or cycles present",
            Self::ScaleBound => "Likely scale-bound — many cells, modest depth",
        }
    }
}

fn classify_report(report: &visigrid_engine::recalc::RecalcReport) -> Classification {
    // Lua-bound: lua_total > 30% of evaluation phase
    if report.phase_eval_us > 0 && report.phase_lua_total_us > 0 {
        let lua_pct = report.phase_lua_total_us as f64 / report.phase_eval_us as f64;
        if lua_pct > 0.30 {
            return Classification::LuaBound;
        }
    }

    // Structure-bound: cycles exist (unresolved), or max depth is high relative to cell count
    if report.had_cycles && !report.converged {
        return Classification::StructureBound;
    }
    if report.cells_recomputed > 0 && report.max_depth > 0 {
        // depth/cells ratio: a flat sheet is ~0, a deeply chained sheet approaches 1
        let depth_ratio = report.max_depth as f64 / report.cells_recomputed as f64;
        if depth_ratio > 0.1 || report.max_depth > 50 {
            return Classification::StructureBound;
        }
    }

    // Fast: sub-millisecond or trivial
    if report.duration_ms <= 1 {
        return Classification::Fast;
    }

    // Scale-bound: many cells, modest depth, no cycles — it's just big
    Classification::ScaleBound
}
