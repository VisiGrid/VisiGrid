use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, SelectionFormatState, TriState};
use crate::mode::InspectorTab;
use crate::theme::TokenKey;
use visigrid_engine::formula::parser::{parse, extract_cell_refs};
use visigrid_engine::cell::{Alignment, VerticalAlignment, TextOverflow, NumberFormat, DateStyle};
use visigrid_engine::cell_id::CellId;

pub const PANEL_WIDTH: f32 = 280.0;

/// Render the inspector panel (right-side drawer)
pub fn render_inspector_panel(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let accent = app.token(TokenKey::Accent);

    // Get the cell to inspect (pinned or selected)
    let (row, col) = app.inspector_pinned.unwrap_or(app.view_state.selected);
    let cell_ref = app.cell_ref_at(row, col);
    let is_pinned = app.inspector_pinned.is_some();
    let current_tab = app.inspector_tab;

    div()
        .id("inspector-panel")
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
        // Clear hover highlight when mouse leaves inspector panel
        .on_mouse_up_out(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.inspector_hover_cell.is_some() {
                this.inspector_hover_cell = None;
                cx.notify();
            }
        }))
        // Header with title and close button
        .child(render_header(&cell_ref, is_pinned, text_primary, text_muted, panel_border, cx))
        // Tab bar
        .child(render_tab_bar(current_tab, text_primary, text_muted, selection_bg, panel_border, cx))
        // Content area
        .child(render_content(app, row, col, current_tab, text_primary, text_muted, accent, panel_border, cx))
}

fn render_header(
    cell_ref: &str,
    is_pinned: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let cell_ref_owned: SharedString = cell_ref.to_string().into();

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
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(13.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(SharedString::from(format!("Inspector: {}", cell_ref_owned)))
                )
                .when(is_pinned, |el| {
                    el.child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child("(pinned)")
                    )
                })
        )
        .child(
            div()
                .flex()
                .items_center()
                .gap_1()
                // Pin/Unpin button
                .child(
                    div()
                        .id("inspector-pin-btn")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(if is_pinned { text_primary } else { text_muted })
                        .hover(|s| s.bg(panel_border))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.toggle_inspector_pin(cx);
                        }))
                        .child(if is_pinned { "Unpin" } else { "Pin" })
                )
                // Close button
                .child(
                    div()
                        .id("inspector-close-btn")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .hover(|s| s.bg(panel_border))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            // Exit preview when closing inspector
                            if this.is_previewing() {
                                this.exit_preview(cx);
                            }
                            this.inspector_visible = false;
                            cx.notify();
                        }))
                        .child("X")
                )
        )
}

fn render_tab_bar(
    current_tab: InspectorTab,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .id("inspector-tab-inspector")
                .flex_1()
                .px_3()
                .py_2()
                .text_size(px(12.0))
                .text_color(if current_tab == InspectorTab::Inspector { text_primary } else { text_muted })
                .font_weight(if current_tab == InspectorTab::Inspector { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                .bg(if current_tab == InspectorTab::Inspector { selection_bg.opacity(0.3) } else { gpui::transparent_black() })
                .border_b_2()
                .border_color(if current_tab == InspectorTab::Inspector { text_primary } else { gpui::transparent_black() })
                .cursor_pointer()
                .hover(|s| s.bg(panel_border.opacity(0.5)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.inspector_tab = InspectorTab::Inspector;
                    cx.notify();
                }))
                .child("Inspector")
        )
        .child(
            div()
                .id("inspector-tab-format")
                .flex_1()
                .px_3()
                .py_2()
                .text_size(px(12.0))
                .text_color(if current_tab == InspectorTab::Format { text_primary } else { text_muted })
                .font_weight(if current_tab == InspectorTab::Format { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                .bg(if current_tab == InspectorTab::Format { selection_bg.opacity(0.3) } else { gpui::transparent_black() })
                .border_b_2()
                .border_color(if current_tab == InspectorTab::Format { text_primary } else { gpui::transparent_black() })
                .cursor_pointer()
                .hover(|s| s.bg(panel_border.opacity(0.5)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.inspector_tab = InspectorTab::Format;
                    cx.notify();
                }))
                .child("Format")
        )
        .child(
            div()
                .id("inspector-tab-names")
                .flex_1()
                .px_3()
                .py_2()
                .text_size(px(12.0))
                .text_color(if current_tab == InspectorTab::Names { text_primary } else { text_muted })
                .font_weight(if current_tab == InspectorTab::Names { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                .bg(if current_tab == InspectorTab::Names { selection_bg.opacity(0.3) } else { gpui::transparent_black() })
                .border_b_2()
                .border_color(if current_tab == InspectorTab::Names { text_primary } else { gpui::transparent_black() })
                .cursor_pointer()
                .hover(|s| s.bg(panel_border.opacity(0.5)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.inspector_tab = InspectorTab::Names;
                    cx.notify();
                }))
                .child("Names")
        )
        .child(
            div()
                .id("inspector-tab-history")
                .flex_1()
                .px_3()
                .py_2()
                .text_size(px(12.0))
                .text_color(if current_tab == InspectorTab::History { text_primary } else { text_muted })
                .font_weight(if current_tab == InspectorTab::History { FontWeight::MEDIUM } else { FontWeight::NORMAL })
                .bg(if current_tab == InspectorTab::History { selection_bg.opacity(0.3) } else { gpui::transparent_black() })
                .border_b_2()
                .border_color(if current_tab == InspectorTab::History { text_primary } else { gpui::transparent_black() })
                .cursor_pointer()
                .hover(|s| s.bg(panel_border.opacity(0.5)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.inspector_tab = InspectorTab::History;
                    cx.notify();
                }))
                .child("History")
        )
}

fn render_content(
    app: &mut Spreadsheet,
    row: usize,
    col: usize,
    current_tab: InspectorTab,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Pre-build tabs that need &mut access
    let names_content = if current_tab == InspectorTab::Names {
        Some(render_names_tab(app, text_primary, text_muted, accent, panel_border, cx))
    } else {
        None
    };

    let history_content = if current_tab == InspectorTab::History {
        Some(render_history_tab(app, text_primary, text_muted, accent, panel_border, cx))
    } else {
        None
    };

    div()
        .flex_1()
        .overflow_hidden()
        .child(match current_tab {
            InspectorTab::Inspector => render_inspector_tab(app, row, col, text_primary, text_muted, accent, panel_border, cx),
            InspectorTab::Format => render_format_tab(app, row, col, text_primary, text_muted, panel_border, accent, cx).into_any_element(),
            InspectorTab::Names => names_content.unwrap().into_any_element(),
            InspectorTab::History => history_content.unwrap().into_any_element(),
        })
}

fn render_inspector_tab(
    app: &Spreadsheet,
    row: usize,
    col: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> AnyElement {
    // Basic Inspector is now FREE - only advanced explainability features are Pro
    // Free: Identity, Inputs list, Outputs list, Spill info
    // Pro: Verification Certificate, Impact Summary, Mini DAG, Trust Metrics
    let is_pro = visigrid_license::is_feature_enabled("inspector");

    let raw_value = app.sheet(cx).get_raw(row, col);
    let display_value = app.sheet(cx).get_display(row, col);
    let is_formula = raw_value.starts_with('=');

    // Determine cell type
    let cell_type = if raw_value.is_empty() {
        "Empty"
    } else if is_formula {
        if display_value.starts_with('#') {
            "Error"
        } else {
            "Formula"
        }
    } else if raw_value.parse::<f64>().is_ok() {
        "Number"
    } else if raw_value == "TRUE" || raw_value == "FALSE" {
        "Boolean"
    } else {
        "Text"
    };

    // Get precedents (cells this formula depends on)
    let precedents = if is_formula {
        get_precedents(&raw_value)
    } else {
        Vec::new()
    };

    // Get dependents (cells that depend on this cell)
    let dependents = get_dependents(app, row, col, cx);

    // Check spill info
    let is_spill_parent = app.sheet(cx).is_spill_parent(row, col);
    let is_spill_receiver = app.sheet(cx).is_spill_receiver(row, col);
    let spill_parent = app.sheet(cx).get_spill_parent(row, col);

    let cell_address = app.cell_ref_at(row, col);
    let has_spill_info = is_spill_parent || is_spill_receiver;
    let has_no_deps = precedents.is_empty() && dependents.is_empty() && !is_formula && !has_spill_info;

    // Get verification data for Pro features
    let recalc_info = if is_pro && app.verified_mode {
        if let Some(report) = &app.last_recalc_report {
            let sheet_id = app.sheet(cx).id;
            let cell_id = CellId::new(sheet_id, row, col);
            report.get_cell_info(&cell_id).cloned()
        } else {
            None
        }
    } else {
        None
    };

    let mut content = div()
        .p_3()
        .flex()
        .flex_col()
        .gap_4();

    // ========== PRO: TRUST BLOCK (the certificate) ==========
    if is_pro && app.verified_mode && is_formula {
        let verification_section = if let Some(ref info) = recalc_info {
            // Complexity label based on depth
            let complexity_label = if info.has_unknown_deps {
                "Dynamic calculation"
            } else if info.depth == 1 {
                "Simple calculation"
            } else if info.depth <= 3 {
                "Multi-step calculation"
            } else if info.depth <= 6 {
                "Complex calculation"
            } else {
                "Complex system"
            };

            // Visual treatment by complexity (accent intensity)
            // Orange warning color as Hsla for opacity support
            let warning_color: Hsla = rgb(0xFFA500).into();
            let (block_bg, block_border, show_warning) = if info.has_unknown_deps {
                // Dynamic refs: warning tone
                (warning_color.opacity(0.08), warning_color.opacity(0.3), true)
            } else if info.depth >= 7 {
                // Complex system: warning tone
                (warning_color.opacity(0.08), warning_color.opacity(0.3), true)
            } else if info.depth >= 4 {
                // Complex: stronger emphasis
                (accent.opacity(0.12), accent.opacity(0.4), false)
            } else if info.depth >= 2 {
                // Multi-step: slight accent
                (accent.opacity(0.08), accent.opacity(0.3), false)
            } else {
                // Simple: calm, compact
                (accent.opacity(0.05), accent.opacity(0.2), false)
            };

            // Guarantee as bullet points
            let guarantee = if info.has_unknown_deps {
                "Fully recomputed • Dynamic refs"
            } else {
                "Fully recomputed • Cycle-free"
            };

            // Stats line
            let stats = format!(
                "{}ms • Depth {} • Position {} of {}",
                app.last_recalc_report.as_ref().map(|r| r.duration_ms).unwrap_or(0),
                info.depth,
                info.eval_order + 1,
                app.last_recalc_report.as_ref().map(|r| r.cells_recomputed).unwrap_or(0)
            );

            div()
                .p_3()
                .rounded(px(6.0))
                .bg(block_bg)
                .border_1()
                .border_color(block_border)
                .flex()
                .flex_col()
                .gap_1()
                // Line 1: ✓ VERIFIED (or ⚠ for warnings)
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(if show_warning { warning_color } else { accent })
                        .child(if show_warning { "⚠ VERIFIED" } else { "✓ VERIFIED" })
                )
                // Line 2: Complexity label
                .child(
                    div()
                        .text_size(px(13.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(text_primary)
                        .child(complexity_label)
                )
                // Line 3: Guarantee points (blank line effect via margin)
                .child(
                    div()
                        .mt_1()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(guarantee)
                )
                // Line 4: Stats
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted.opacity(0.7))
                        .child(stats)
                )
        } else {
            // Formula not in last recalc (possibly new or in cycle)
            div()
                .p_3()
                .rounded(px(6.0))
                .bg(text_muted.opacity(0.08))
                .border_1()
                .border_color(text_muted.opacity(0.3))
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_muted)
                        .child("Not in last verification")
                )
        };
        content = content.child(verification_section);
    }

    // ========== PRO: TRUST HEADER (Phase 3.5 — Impact + Risk Semantics) ==========
    // This is the "certificate" feel — shows impact and any trust issues
    if is_pro {
        let sheet_id = app.sheet(cx).id;
        let impact = app.wb(cx).compute_impact(sheet_id, row, col);

        // Cycle detection: use recalc report as source of truth when available
        // 1. Check if this cell itself is #CYCLE!
        // 2. Check if any upstream cell has #CYCLE! (via graph traversal)
        // 3. If verified mode, check recalc report's had_cycles flag
        let cell_is_cycle = display_value == "#CYCLE!";
        let upstream_has_cycle = app.wb(cx).has_cycle_in_upstream(sheet_id, row, col);
        let report_had_cycles = app.verified_mode
            && app.last_recalc_report.as_ref().map(|r| r.had_cycles).unwrap_or(false);

        // Determine risk state
        let has_dynamic = impact.has_unknown_in_chain;
        // Cell is affected by cycle if: it IS a cycle, its inputs have cycles, or workbook has cycles affecting it
        let has_cycle = cell_is_cycle || upstream_has_cycle || (report_had_cycles && upstream_has_cycle);
        let is_verifiable = !has_cycle;

        // Only show impact header if there are dependents or it's a formula
        let show_impact = impact.affected_cells > 0 || is_formula;

        if show_impact {
            // Warning color for risk states
            let warning_color: Hsla = rgb(0xFFA500).into();
            let error_color: Hsla = rgb(0xE53935).into();

            // Impact line text
            let impact_text = if impact.is_unbounded {
                "unbounded (dynamic refs)".to_string()
            } else if impact.affected_cells == 0 {
                "no downstream cells".to_string()
            } else if impact.affected_cells == 1 {
                format!("affects 1 cell • max depth {}", impact.max_depth)
            } else {
                format!("affects {} cells • max depth {}", impact.affected_cells, impact.max_depth)
            };

            // Build badges
            let mut badges: Vec<(&str, Hsla)> = Vec::new();
            if has_dynamic {
                badges.push(("Dynamic", warning_color));
            }
            if has_cycle {
                badges.push(("Cycle", error_color));
            }
            if !is_verifiable {
                badges.push(("Not verifiable", error_color));
            }

            let header_border = if !is_verifiable {
                error_color.opacity(0.3)
            } else if has_dynamic {
                warning_color.opacity(0.3)
            } else {
                accent.opacity(0.2)
            };

            let header_bg = if !is_verifiable {
                error_color.opacity(0.05)
            } else if has_dynamic {
                warning_color.opacity(0.05)
            } else {
                accent.opacity(0.03)
            };

            let impact_section = div()
                .p_2()
                .rounded(px(4.0))
                .bg(header_bg)
                .border_1()
                .border_color(header_border)
                .flex()
                .flex_col()
                .gap_1()
                // Impact line
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(10.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_muted)
                                .child("Impact:")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(if impact.is_unbounded { warning_color } else { text_primary })
                                .child(SharedString::from(impact_text))
                        )
                )
                // Risk badges (if any)
                .when(!badges.is_empty(), |el| {
                    el.child(
                        div()
                            .flex()
                            .gap_1()
                            .mt_1()
                            .children(badges.into_iter().map(|(label, color)| {
                                div()
                                    .px_2()
                                    .py(px(2.0))
                                    .rounded(px(3.0))
                                    .bg(color.opacity(0.15))
                                    .border_1()
                                    .border_color(color.opacity(0.4))
                                    .text_size(px(9.0))
                                    .font_weight(FontWeight::MEDIUM)
                                    .text_color(color)
                                    .child(label)
                            }))
                    )
                });

            content = content.child(impact_section);
        }
    }

    // ========== PRO: TRACE PATH (Phase 3.5b) ==========
    // Show the current trace path when active
    if is_pro {
        if let Some(ref trace_path) = app.inspector_trace_path {
            if !trace_path.is_empty() {
                let sheet_id = app.sheet(cx).id;
                let warning_color: Hsla = rgb(0xFFA500).into();

                // Build path string: A1 → B1 → C1 → ...
                let path_str: String = trace_path
                    .iter()
                    .filter(|cell| cell.sheet == sheet_id)
                    .map(|cell| app.cell_ref_at(cell.row, cell.col))
                    .collect::<Vec<_>>()
                    .join(" → ");

                // Truncate if too long
                let display_path = if path_str.len() > 60 {
                    let first = &trace_path[0];
                    let last = &trace_path[trace_path.len() - 1];
                    format!(
                        "{} → ... → {} ({} cells)",
                        app.cell_ref_at(first.row, first.col),
                        app.cell_ref_at(last.row, last.col),
                        trace_path.len()
                    )
                } else {
                    path_str
                };

                let trace_section = div()
                    .p_2()
                    .rounded(px(4.0))
                    .bg(accent.opacity(0.08))
                    .border_1()
                    .border_color(accent.opacity(0.3))
                    .flex()
                    .flex_col()
                    .gap_1()
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .justify_between()
                            .child(
                                div()
                                    .text_size(px(10.0))
                                    .font_weight(FontWeight::SEMIBOLD)
                                    .text_color(text_muted)
                                    .child("Trace")
                            )
                            .child(
                                div()
                                    .id("trace-clear-btn")
                                    .px_1()
                                    .rounded_sm()
                                    .cursor_pointer()
                                    .text_size(px(9.0))
                                    .text_color(text_muted)
                                    .hover(|s| s.bg(panel_border.opacity(0.5)))
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.clear_trace_path(cx);
                                    }))
                                    .child("Clear")
                            )
                    )
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_primary)
                            .child(SharedString::from(display_path))
                    )
                    .when(app.inspector_trace_incomplete, |el| {
                        el.child(
                            div()
                                .mt_1()
                                .px_2()
                                .py(px(2.0))
                                .rounded(px(3.0))
                                .bg(warning_color.opacity(0.15))
                                .text_size(px(9.0))
                                .text_color(warning_color)
                                .child("Trace incomplete (dynamic refs)")
                        )
                    });

                content = content.child(trace_section);
            }
        }
    }

    // ========== IDENTITY SECTION (Free) ==========
    content = content.child(
        section("Identity", panel_border, text_primary)
            .child(info_row("Address", &cell_address, text_muted, text_primary))
            .child(info_row("Type", cell_type, text_muted, text_primary))
            .when(is_formula, |el| {
                el.child(info_row_multiline("Formula", &raw_value, text_muted, text_primary))
            })
            .child(info_row_multiline(
                if is_formula { "Result" } else { "Value" },
                if display_value.is_empty() { "(empty)" } else { &display_value },
                text_muted,
                text_primary,
            ))
    );

    // ========== VALIDATION SECTION (Free) ==========
    // Show validation info if this cell has data validation
    if let Some(rule) = app.sheet(cx).validations.get(row, col) {
        use visigrid_engine::validation::{ValidationType, ListSource};

        let validation_type = match &rule.rule_type {
            // NOTE: No AnyValue case - that variant no longer exists
            ValidationType::List(source) => {
                let source_desc = match source {
                    ListSource::Inline(items) => format!("{} items", items.len()),
                    ListSource::Range(r) => r.clone(),
                    ListSource::NamedRange(n) => n.clone(),
                };
                format!("List ({})", source_desc)
            }
            ValidationType::WholeNumber(_) => "Whole number".to_string(),
            ValidationType::Decimal(_) => "Decimal".to_string(),
            ValidationType::Date(_) => "Date".to_string(),
            ValidationType::Time(_) => "Time".to_string(),
            ValidationType::TextLength(_) => "Text length".to_string(),
            ValidationType::Custom(f) => format!("Custom: {}", f),
        };

        let mut validation_section = section("Validation", panel_border, text_primary)
            .child(info_row("Type", &validation_type, text_muted, text_primary));

        // Show dropdown hint for list validation
        if rule.show_dropdown && matches!(rule.rule_type, ValidationType::List(_)) {
            validation_section = validation_section.child(
                div()
                    .flex()
                    .justify_between()
                    .py_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .child("Open dropdown")
                    )
                    .child(
                        div()
                            .px_2()
                            .py(px(2.0))
                            .rounded(px(3.0))
                            .bg(accent.opacity(0.1))
                            .border_1()
                            .border_color(accent.opacity(0.3))
                            .text_size(px(10.0))
                            .text_color(accent)
                            .child(if cfg!(target_os = "macos") { "Option+Down" } else { "Alt+Down" })
                    )
            );
        }

        // Phase 6C: Show validation status (Valid/Invalid with reason)
        if let Some(reason) = app.get_invalid_reason(row, col) {
            use visigrid_engine::validation::ValidationFailureReason;
            let error_color: Hsla = rgb(0xE53935).into();
            let reason_text = match reason {
                ValidationFailureReason::InvalidValue => "Value does not meet criteria",
                ValidationFailureReason::ConstraintBlank => "Constraint cell is empty",
                ValidationFailureReason::ConstraintNotNumeric => "Constraint cell is not numeric",
                ValidationFailureReason::InvalidReference => "Invalid reference",
                ValidationFailureReason::FormulaNotSupported => "Formula constraints not supported",
                ValidationFailureReason::ListEmpty => "List is empty",
                ValidationFailureReason::NotInList => "Value not in allowed list",
            };
            validation_section = validation_section.child(
                div()
                    .flex()
                    .justify_between()
                    .items_center()
                    .py_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .child("Status")
                    )
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            .child(
                                div()
                                    .w(px(6.0))
                                    .h(px(6.0))
                                    .rounded(px(3.0))
                                    .bg(error_color)
                            )
                            .child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(error_color)
                                    .child(SharedString::from(reason_text))
                            )
                    )
            );
        } else if !display_value.is_empty() {
            // Show "Valid" status for non-empty cells with validation
            let valid_color: Hsla = rgb(0x43A047).into();  // Material Green 600
            validation_section = validation_section.child(
                div()
                    .flex()
                    .justify_between()
                    .items_center()
                    .py_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .child("Status")
                    )
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            .child(
                                div()
                                    .w(px(6.0))
                                    .h(px(6.0))
                                    .rounded(px(3.0))
                                    .bg(valid_color)
                            )
                            .child(
                                div()
                                    .text_size(px(11.0))
                                    .text_color(valid_color)
                                    .child("Valid")
                            )
                    )
            );
        }

        content = content.child(validation_section);
    }

    // Spill info (if applicable)
    if has_spill_info {
        let mut spill_section = section("Array Spill", panel_border, text_primary);
        if is_spill_parent {
            spill_section = spill_section.child(info_row("Role", "Spill Parent", text_muted, text_primary));
        }
        if is_spill_receiver {
            spill_section = spill_section.child(info_row("Role", "Spill Receiver", text_muted, text_primary));
            if let Some((pr, pc)) = spill_parent {
                spill_section = spill_section.child(clickable_cell_row(
                    "Parent",
                    &app.cell_ref_at(pr, pc),
                    pr,
                    pc,
                    text_muted,
                    accent,
                    cx,
                ));
            }
        }
        content = content.child(spill_section);
    }

    // Mini DAG visualization (Pro feature)
    if is_pro && (!precedents.is_empty() || !dependents.is_empty()) {
        content = content.child(render_mini_dag(
            app,
            &cell_address,
            &precedents,
            &dependents,
            text_primary,
            text_muted,
            accent,
            panel_border,
            cx,
        ));
    }

    // Inputs section (formerly Precedents) - click triggers trace
    if !precedents.is_empty() {
        let mut prec_section = section("Inputs", panel_border, text_primary)
            .child(
                div()
                    .text_size(px(9.0))
                    .text_color(text_muted.opacity(0.7))
                    .mb_1()
                    .child("Click to trace path")
            );
        for (r, c) in precedents.iter().take(10) {
            prec_section = prec_section.child(traceable_cell_row(
                &app.cell_ref_at(*r, *c),
                *r,
                *c,
                row,  // inspected cell
                col,
                true,  // is_input
                text_muted,
                accent,
                cx,
            ));
        }
        if precedents.len() > 10 {
            prec_section = prec_section.child(
                div()
                    .text_size(px(11.0))
                    .text_color(text_muted)
                    .child(SharedString::from(format!("+ {} more...", precedents.len() - 10)))
            );
        }
        content = content.child(prec_section);
    }

    // Outputs section (formerly Dependents) - click triggers trace
    if !dependents.is_empty() {
        let mut dep_section = section("Outputs", panel_border, text_primary)
            .child(
                div()
                    .text_size(px(9.0))
                    .text_color(text_muted.opacity(0.7))
                    .mb_1()
                    .child("Click to trace path")
            );
        for (r, c) in dependents.iter().take(10) {
            dep_section = dep_section.child(traceable_cell_row(
                &app.cell_ref_at(*r, *c),
                *r,
                *c,
                row,  // inspected cell
                col,
                false,  // is_output
                text_muted,
                accent,
                cx,
            ));
        }
        if dependents.len() > 10 {
            dep_section = dep_section.child(
                div()
                    .text_size(px(11.0))
                    .text_color(text_muted)
                    .child(SharedString::from(format!("+ {} more...", dependents.len() - 10)))
            );
        }
        content = content.child(dep_section);
    }

    // Proof section (Pro feature, only when verified mode is enabled)
    if is_pro && app.verified_mode {
        if let Some(report) = &app.last_recalc_report {
            let sheet_id = app.sheet(cx).id;
            let cell_id = CellId::new(sheet_id, row, col);

            if let Some(info) = report.get_cell_info(&cell_id) {
                let mut recalc_section = section("Proof", panel_border, text_primary);

                // Depth
                recalc_section = recalc_section.child(info_row(
                    "Depth",
                    &format!("{}", info.depth),
                    text_muted,
                    text_primary,
                ));

                // Evaluation order
                recalc_section = recalc_section.child(info_row(
                    "Eval Order",
                    &format!("#{} of {}", info.eval_order + 1, report.cells_recomputed),
                    text_muted,
                    text_primary,
                ));

                // Recompute timestamp (relative time)
                let timestamp = format_relative_time(info.recompute_time);
                recalc_section = recalc_section.child(info_row(
                    "Recomputed",
                    &timestamp,
                    text_muted,
                    text_primary,
                ));

                // Has unknown deps indicator
                if info.has_unknown_deps {
                    recalc_section = recalc_section.child(info_row(
                        "Dynamic Refs",
                        "Yes (INDIRECT/OFFSET)",
                        text_muted,
                        text_primary,
                    ));
                }

                // Adjacent cells (evaluated before/after)
                let (prev_cell, next_cell) = report.get_adjacent_cells(&cell_id);

                if let Some(prev) = prev_cell {
                    if prev.sheet == sheet_id {
                        recalc_section = recalc_section.child(clickable_cell_row(
                            "After",
                            &app.cell_ref_at(prev.row, prev.col),
                            prev.row,
                            prev.col,
                            text_muted,
                            accent,
                            cx,
                        ));
                    }
                }

                if let Some(next) = next_cell {
                    if next.sheet == sheet_id {
                        recalc_section = recalc_section.child(clickable_cell_row(
                            "Before",
                            &app.cell_ref_at(next.row, next.col),
                            next.row,
                            next.col,
                            text_muted,
                            accent,
                            cx,
                        ));
                    }
                }

                content = content.child(recalc_section);
            } else if is_formula {
                // Formula cell but not in recalc report (might be cycle or new)
                content = content.child(
                    section("Proof", panel_border, text_primary)
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child("Not in last verification")
                        )
                );
            }
        }
    } else if is_pro && is_formula {
        // Hint to enable verified mode (Pro users only)
        content = content.child(
            div()
                .py_2()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("Enable Verified Mode (F9) for proof")
        );
    } else if !is_pro && is_formula && (!precedents.is_empty() || !dependents.is_empty()) {
        // Upsell for Free users with formulas
        content = content.child(
            div()
                .py_2()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("Upgrade to Pro for trust verification")
        );
    }

    // Empty state for non-formula cells with no dependencies
    if has_no_deps {
        content = content.child(
            div()
                .py_4()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("No dependencies")
        );
    }

    content.into_any_element()
}

/// Render a mini DAG visualization showing precedents → cell → dependents flow
fn render_mini_dag(
    app: &Spreadsheet,
    cell_address: &str,
    precedents: &[(usize, usize)],
    dependents: &[(usize, usize)],
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let max_show = 4; // Max cells to show in each row

    section("Dependency Graph", panel_border, text_primary)
        // Fan-in (precedents) row
        .when(!precedents.is_empty(), |el| {
            let prec_cells: Vec<_> = precedents.iter().take(max_show).collect();
            let extra_count = precedents.len().saturating_sub(max_show);

            el.child(
                div()
                    .flex()
                    .flex_wrap()
                    .gap_1()
                    .justify_center()
                    .child(
                        div()
                            .text_size(px(9.0))
                            .text_color(text_muted)
                            .mr_1()
                            .child("inputs")
                    )
                    .children(prec_cells.into_iter().map(|(r, c)| {
                        dag_cell_chip(app, *r, *c, text_muted, accent, cx)
                    }))
                    .when(extra_count > 0, |el| {
                        el.child(
                            div()
                                .text_size(px(9.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!("+{}", extra_count)))
                        )
                    })
            )
        })
        // Arrow down
        .when(!precedents.is_empty(), |el| {
            el.child(
                div()
                    .flex()
                    .justify_center()
                    .py_px()
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child("↓")
                    )
            )
        })
        // Current cell (highlighted)
        .child(
            div()
                .flex()
                .justify_center()
                .child(
                    div()
                        .px_2()
                        .py_1()
                        .rounded(px(4.0))
                        .bg(accent.opacity(0.2))
                        .border_1()
                        .border_color(accent)
                        .text_size(px(11.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(SharedString::from(cell_address.to_string()))
                )
        )
        // Arrow down
        .when(!dependents.is_empty(), |el| {
            el.child(
                div()
                    .flex()
                    .justify_center()
                    .py_px()
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child("↓")
                    )
            )
        })
        // Fan-out (dependents) row
        .when(!dependents.is_empty(), |el| {
            let dep_cells: Vec<_> = dependents.iter().take(max_show).collect();
            let extra_count = dependents.len().saturating_sub(max_show);

            el.child(
                div()
                    .flex()
                    .flex_wrap()
                    .gap_1()
                    .justify_center()
                    .child(
                        div()
                            .text_size(px(9.0))
                            .text_color(text_muted)
                            .mr_1()
                            .child("outputs")
                    )
                    .children(dep_cells.into_iter().map(|(r, c)| {
                        dag_cell_chip(app, *r, *c, text_muted, accent, cx)
                    }))
                    .when(extra_count > 0, |el| {
                        el.child(
                            div()
                                .text_size(px(9.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!("+{}", extra_count)))
                        )
                    })
            )
        })
}

/// A small clickable cell chip for the DAG visualization
fn dag_cell_chip(
    app: &Spreadsheet,
    row: usize,
    col: usize,
    text_muted: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let cell_ref = app.cell_ref_at(row, col);

    div()
        .id(SharedString::from(format!("dag-chip-{}-{}", row, col)))
        .px_1()
        .rounded(px(3.0))
        .bg(text_muted.opacity(0.1))
        .text_size(px(10.0))
        .text_color(accent)
        .cursor_pointer()
        .hover(|s| s.bg(accent.opacity(0.2)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.select_cell(row, col, false, cx);
        }))
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.inspector_hover_cell != Some((row, col)) {
                this.inspector_hover_cell = Some((row, col));
                cx.notify();
            }
        }))
        .on_mouse_up_out(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.inspector_hover_cell.is_some() {
                this.inspector_hover_cell = None;
                cx.notify();
            }
        }))
        .child(SharedString::from(cell_ref))
}

fn render_format_tab(
    app: &Spreadsheet,
    _row: usize,
    _col: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Get format state for entire selection (tri-state resolution)
    let state = app.selection_format_state(cx);

    div()
        .p_3()
        .flex()
        .flex_col()
        .gap_4()
        // Selection summary (with value/format state)
        .child(render_selection_summary(&state, text_muted, panel_border))
        // Value preview (raw + display)
        .child(render_value_preview(&state, text_primary, text_muted, panel_border))
        // Number format section
        .child(render_number_format_section(&state, text_primary, text_muted, accent, panel_border, cx))
        // Alignment section
        .child(render_alignment_section(&state, text_primary, text_muted, accent, panel_border, cx))
        // Text style toggles
        .child(render_text_style_section(&state, text_primary, text_muted, accent, panel_border, cx))
        // Background color section
        .child(render_background_color_section(&state, text_primary, text_muted, accent, panel_border, cx))
        // Font section
        .child(render_font_section(&state, text_primary, text_muted, panel_border, cx))
}

fn render_names_tab(
    app: &mut Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // First collect the named ranges (immutable borrow)
    let ranges: Vec<_> = app.filtered_named_ranges(cx);

    // Then get usage counts (mutable borrow for cache)
    let named_ranges: Vec<_> = ranges.into_iter()
        .map(|nr| {
            let usage_count = app.get_named_range_usage_count(&nr.name, cx);
            (nr, usage_count)
        })
        .collect();

    let filter_query = app.names_filter_query.clone();
    let all_names = app.wb(cx).list_named_ranges();
    let has_names = !all_names.is_empty();
    let has_filtered_results = !named_ranges.is_empty();
    let name_count = all_names.len();
    let selected_name = app.selected_named_range.clone();
    let is_pro = visigrid_license::is_feature_enabled("inspector");

    // Get detail info for selected named range (Pro feature)
    let selected_detail = if is_pro {
        selected_name.as_ref().and_then(|name| {
            get_named_range_detail(app, name, cx)
        })
    } else {
        None
    };

    // Build the list content based on state
    let list_content = if !has_names {
        // Empty state
        div()
            .flex_1()
            .flex()
            .flex_col()
            .py_4()
            .items_center()
            .gap_2()
            .child(
                div()
                    .text_size(px(11.0))
                    .text_color(text_muted)
                    .child("Names let you refactor formulas safely.")
            )
            .child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted.opacity(0.7))
                    .child("Press Ctrl+Shift+N to create one.")
            )
    } else if !has_filtered_results {
        // No matches state
        div()
            .flex_1()
            .py_4()
            .text_size(px(11.0))
            .text_color(text_muted)
            .child("No matches")
    } else {
        // Build named ranges list
        let mut list = div()
            .flex_1()
            .flex()
            .flex_col()
            .gap_1();

        for (nr, usage_count) in named_ranges.iter() {
            let name = nr.name.clone();
            let name_for_jump = nr.name.clone();
            let name_for_rename = nr.name.clone();
            let name_for_edit = nr.name.clone();
            let name_for_delete = nr.name.clone();
            let reference = nr.reference_string();
            let description = nr.description.clone();
            let usage = *usage_count;

            // Show description under reference if it exists
            let has_description = description.is_some();

            let is_selected = selected_name.as_ref() == Some(&name);
            let name_for_select = nr.name.clone();

            list = list.child(
                div()
                    .id(SharedString::from(format!("named-range-{}", name)))
                    .group("named-range-row")
                    .flex()
                    .items_center()
                    .justify_between()
                    .px_2()
                    .py_1()
                    .rounded_sm()
                    .cursor_pointer()
                    .bg(if is_selected { accent.opacity(0.15) } else { gpui::transparent_black() })
                    .hover(|s| s.bg(panel_border.opacity(0.3)))
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                        if event.click_count == 2 {
                            // Double-click: jump to range
                            this.jump_to_named_range(&name_for_jump, cx);
                        } else {
                            // Single-click: toggle selection and trace
                            if this.selected_named_range.as_ref() == Some(&name_for_select) {
                                this.selected_named_range = None;
                                this.clear_named_range_trace(cx);
                            } else {
                                this.selected_named_range = Some(name_for_select.clone());
                                this.trace_named_range(&name_for_select, cx);
                            }
                        }
                    }))
                    .child(
                        div()
                            .flex()
                            .flex_col()
                            .overflow_hidden()
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .text_size(px(11.0))
                                            .text_color(text_primary)
                                            .font_weight(FontWeight::MEDIUM)
                                            .child(SharedString::from(name.clone()))
                                    )
                                    // Usage count badge
                                    .when(usage > 0, |d| {
                                        d.child(
                                            div()
                                                .px_1()
                                                .py(px(1.0))
                                                .bg(accent.opacity(0.2))
                                                .rounded(px(4.0))
                                                .text_size(px(8.0))
                                                .text_color(accent)
                                                .child(SharedString::from(format!("{}", usage)))
                                        )
                                    })
                            )
                            .child(
                                div()
                                    .text_size(px(10.0))
                                    .text_color(text_muted)
                                    .child(SharedString::from(reference))
                            )
                            .when(has_description, |d| {
                                d.child(
                                    div()
                                        .text_size(px(9.0))
                                        .text_color(text_muted.opacity(0.7))
                                        .overflow_hidden()
                                        .child(SharedString::from(description.unwrap_or_default()))
                                )
                            })
                    )
                    // Action buttons (appear on hover)
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            .opacity(0.0)
                            .group_hover("named-range-row", |s| s.opacity(1.0))
                            // Rename button
                            .child(
                                div()
                                    .id(SharedString::from(format!("rename-{}", name)))
                                    .px_1()
                                    .py(px(2.0))
                                    .rounded_sm()
                                    .text_size(px(9.0))
                                    .text_color(text_muted)
                                    .hover(|s| s.bg(accent.opacity(0.2)).text_color(accent))
                                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                        this.show_rename_symbol(Some(&name_for_rename), cx);
                                    }))
                                    .child("Rename")
                            )
                            // Edit description button
                            .child(
                                div()
                                    .id(SharedString::from(format!("edit-{}", name)))
                                    .px_1()
                                    .py(px(2.0))
                                    .rounded_sm()
                                    .text_size(px(9.0))
                                    .text_color(text_muted)
                                    .hover(|s| s.bg(accent.opacity(0.2)).text_color(accent))
                                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                        this.show_edit_description(&name_for_edit, cx);
                                    }))
                                    .child("Edit")
                            )
                            // Delete button
                            .child(
                                div()
                                    .id(SharedString::from(format!("delete-{}", name)))
                                    .px_1()
                                    .py(px(2.0))
                                    .rounded_sm()
                                    .text_size(px(9.0))
                                    .text_color(text_muted)
                                    .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
                                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                        this.delete_named_range(&name_for_delete, cx);
                                    }))
                                    .child("X")
                            )
                    )
            );
        }
        list
    };

    div()
        .p_3()
        .flex()
        .flex_col()
        .gap_3()
        .h_full()
        // Header with count and create button
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(SharedString::from(format!("{} named ranges", name_count)))
                )
                .child(
                    div()
                        .id("names-create-btn")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(10.0))
                        .text_color(accent)
                        .border_1()
                        .border_color(accent.opacity(0.5))
                        .hover(|s| s.bg(accent.opacity(0.1)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                            this.show_create_named_range(cx);
                            window.focus(&this.focus_handle, cx);
                        }))
                        .child("+ Create")
                )
        )
        // Search input (placeholder for now - will wire up keyboard input later)
        .child(
            div()
                .flex()
                .items_center()
                .px_2()
                .py_1()
                .bg(panel_border.opacity(0.2))
                .border_1()
                .border_color(panel_border)
                .rounded_sm()
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .mr_2()
                        .child("$")
                )
                .child(
                    div()
                        .flex_1()
                        .text_size(px(11.0))
                        .text_color(if filter_query.is_empty() { text_muted } else { text_primary })
                        .child(SharedString::from(if filter_query.is_empty() {
                            "Filter names...".to_string()
                        } else {
                            filter_query.clone()
                        }))
                )
        )
        // Named ranges list
        .child(list_content)
        // Detail panel for selected named range (Pro feature)
        .when(selected_detail.is_some(), |el| {
            let detail = selected_detail.as_ref().unwrap();
            el.child(render_named_range_detail(detail, text_primary, text_muted, accent, panel_border))
        })
        // Pro upsell when not licensed and has selection
        .when(!is_pro && selected_name.is_some(), |el| {
            el.child(
                div()
                    .pt_2()
                    .border_t_1()
                    .border_color(panel_border)
                    .p_2()
                    .bg(panel_border.opacity(0.2))
                    .rounded_sm()
                    .text_size(px(10.0))
                    .text_color(text_muted)
                    .child("Pro: View value preview, depth, and verification status")
            )
        })
        // Keyboard hint
        .child(
            div()
                .pt_2()
                .border_t_1()
                .border_color(panel_border)
                .text_size(px(9.0))
                .text_color(text_muted.opacity(0.7))
                .child("Click: select | Double-click: jump | Ctrl+Shift+N: Create")
        )
}

/// Detail info for a named range (Phase 5)
#[derive(Clone, Debug)]
#[allow(dead_code)]
struct NamedRangeDetail {
    name: String,
    reference: String,
    description: Option<String>,
    /// Value preview (first cell value, or summary for ranges)
    value_preview: String,
    /// Number of cells in the range
    cell_count: usize,
    /// Usage count (formulas referencing this name by name)
    usage_count: usize,
    /// Dependents count (cells that depend on cells in this range, from dependency graph)
    dependents_count: usize,
    /// Dependency depth (max depth of cells in range)
    depth: Option<usize>,
    /// Whether verified mode is on and this range is verified
    is_verified: bool,
    /// Has dynamic refs (INDIRECT/OFFSET) that make verification uncertain
    has_dynamic_refs: bool,
}

/// Get detail info for a named range
fn get_named_range_detail(app: &mut Spreadsheet, name: &str, cx: &App) -> Option<NamedRangeDetail> {
    use visigrid_engine::named_range::NamedRangeTarget;
    use visigrid_engine::cell_id::CellId;

    // Extract all data from named range first (immutable borrow)
    let (reference, description, cells, cell_count, sheet_index) = {
        let nr = app.wb(cx).get_named_range(name)?;
        let reference = nr.reference_string();
        let description = nr.description.clone();

        let (cells, cell_count, sheet_index): (Vec<(usize, usize)>, usize, usize) = match &nr.target {
            NamedRangeTarget::Cell { sheet, row, col } => {
                (vec![(*row, *col)], 1, *sheet)
            }
            NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                let mut cells = Vec::new();
                for r in *start_row..=*end_row {
                    for c in *start_col..=*end_col {
                        cells.push((r, c));
                    }
                }
                let count = cells.len();
                (cells, count, *sheet)
            }
        };

        (reference, description, cells, cell_count, sheet_index)
    };

    // Now we can call mutable methods
    let usage_count = app.get_named_range_usage_count(name, cx);

    // Count dependents from the dependency graph (cells that depend on cells in this range)
    let dependents_count = {
        use std::collections::HashSet;
        // Get the SheetId from the sheet index
        let sheet_id = app.wb(cx).sheets()
            .get(sheet_index)
            .map(|s| s.id)
            .unwrap_or_else(|| app.sheet(cx).id);
        let mut all_dependents: HashSet<CellId> = HashSet::new();

        // Convert cells in range to CellIds for exclusion
        let range_cells: HashSet<CellId> = cells.iter()
            .map(|(r, c)| CellId::new(sheet_id, *r, *c))
            .collect();

        // Collect unique dependents from all cells in the range
        for (row, col) in &cells {
            let deps = app.wb(cx).get_dependents(sheet_id, *row, *col);
            for dep in deps {
                // Don't count cells within the range itself as dependents
                if !range_cells.contains(&dep) {
                    all_dependents.insert(dep);
                }
            }
        }
        all_dependents.len()
    };

    // Value preview: first cell value (or summary for multi-cell ranges)
    let value_preview = if cell_count == 1 {
        let (row, col) = cells[0];
        app.sheet(cx).get_display(row, col)
    } else {
        let first_val = app.sheet(cx).get_display(cells[0].0, cells[0].1);
        if first_val.is_empty() {
            format!("{} cells", cell_count)
        } else {
            format!("{} ... ({} cells)", first_val, cell_count)
        }
    };

    // Compute depth from dependency graph (works without Verified Mode)
    // Depth = max depth of precedent chain (0 for raw values, 1+ for formulas)
    let (depth, is_verified, has_dynamic_refs) = {
        // Get the SheetId from the sheet index
        let sheet_id = app.wb(cx).sheets()
            .get(sheet_index)
            .map(|s| s.id)
            .unwrap_or_else(|| app.sheet(cx).id);

        let mut max_depth = 0usize;
        let mut any_dynamic = false;

        for (row, col) in &cells {
            let cell_id = CellId::new(sheet_id, *row, *col);
            // Compute depth by traversing precedents
            let cell_depth = compute_cell_depth(app.wb(cx), cell_id, &mut std::collections::HashSet::new());
            max_depth = max_depth.max(cell_depth);

            // Check for dynamic refs in the formula
            let raw = app.sheet(cx).get_raw(*row, *col);
            if raw.starts_with('=') {
                if let Ok(expr) = visigrid_engine::formula::parser::parse(&raw[1..]) {
                    if visigrid_engine::formula::analyze::has_dynamic_deps(&expr) {
                        any_dynamic = true;
                    }
                }
            }
        }

        // is_verified only true if Verified Mode is on and we have a valid report
        let is_verified = app.verified_mode && app.last_recalc_report.is_some() && !any_dynamic;

        (Some(max_depth), is_verified, any_dynamic)
    };

    Some(NamedRangeDetail {
        name: name.to_string(),
        reference,
        description,
        value_preview,
        cell_count,
        usage_count,
        dependents_count,
        depth,
        is_verified: app.verified_mode && is_verified,
        has_dynamic_refs,
    })
}

/// Render detail panel for a selected named range
fn render_named_range_detail(
    detail: &NamedRangeDetail,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    div()
        .pt_2()
        .mt_2()
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .gap_2()
        // Header: name with verified badge
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .text_size(px(12.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text_primary)
                        .child(SharedString::from(detail.name.clone()))
                )
                .when(detail.is_verified, |el| {
                    el.child(
                        div()
                            .px_1()
                            .py(px(1.0))
                            .bg(accent.opacity(0.2))
                            .rounded(px(4.0))
                            .text_size(px(8.0))
                            .text_color(accent)
                            .child("Verified")
                    )
                })
                .when(detail.has_dynamic_refs, |el| {
                    el.child(
                        div()
                            .px_1()
                            .py(px(1.0))
                            .bg(hsla(0.08, 0.8, 0.5, 0.2))
                            .rounded(px(4.0))
                            .text_size(px(8.0))
                            .text_color(hsla(0.08, 0.8, 0.5, 1.0))
                            .child("Dynamic")
                    )
                })
        )
        // Value preview
        .child(
            div()
                .flex()
                .flex_col()
                .gap_0p5()
                .child(
                    div()
                        .text_size(px(9.0))
                        .text_color(text_muted)
                        .child("Value")
                )
                .child(
                    div()
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .overflow_hidden()
                        .child(SharedString::from(
                            if detail.value_preview.is_empty() {
                                "(empty)".to_string()
                            } else {
                                detail.value_preview.clone()
                            }
                        ))
                )
        )
        // Metrics row: depth, usage
        .child(
            div()
                .flex()
                .gap_4()
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .child(
                            div()
                                .text_size(px(9.0))
                                .text_color(text_muted)
                                .child("Depth")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_primary)
                                .child(SharedString::from(
                                    detail.depth.map(|d| d.to_string()).unwrap_or_else(|| "—".to_string())
                                ))
                        )
                )
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .child(
                            div()
                                .text_size(px(9.0))
                                .text_color(text_muted)
                                .child("Feeds")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_primary)
                                .child(SharedString::from(format!("{} cells", detail.dependents_count)))
                        )
                )
        )
        // Description (if present)
        .when(detail.description.is_some(), |el| {
            el.child(
                div()
                    .flex()
                    .flex_col()
                    .gap_0p5()
                    .child(
                        div()
                            .text_size(px(9.0))
                            .text_color(text_muted)
                            .child("Description")
                    )
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted.opacity(0.8))
                            .child(SharedString::from(detail.description.clone().unwrap_or_default()))
                    )
            )
        })
}

fn render_selection_summary(
    state: &SelectionFormatState,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    let cell_summary = if state.cell_count == 1 {
        "1 cell selected".to_string()
    } else {
        format!("{} cells selected", state.cell_count)
    };

    let value_state = match &state.raw_value {
        TriState::Uniform(v) if v.is_empty() => "Values: empty",
        TriState::Uniform(_) => "Values: uniform",
        TriState::Mixed => "Values: mixed",
        TriState::Empty => "Values: —",
    };

    let format_state = if state.number_format.is_mixed() || state.alignment.is_mixed() {
        "Formats: mixed"
    } else {
        "Formats: uniform"
    };

    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(
            div()
                .px_2()
                .py_1()
                .bg(panel_border.opacity(0.3))
                .rounded_sm()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child(SharedString::from(cell_summary))
        )
        .child(
            div()
                .flex()
                .gap_3()
                .px_2()
                .text_size(px(9.0))
                .text_color(text_muted)
                .child(value_state)
                .child(format_state)
        )
}

fn render_value_preview(
    state: &SelectionFormatState,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    let raw_display = match &state.raw_value {
        TriState::Uniform(v) if v.is_empty() => "(empty)".to_string(),
        TriState::Uniform(v) => v.clone(),
        TriState::Mixed => "(mixed)".to_string(),
        TriState::Empty => "—".to_string(),
    };

    let formatted_display = match (&state.raw_value, &state.display_value) {
        (TriState::Uniform(_), Some(d)) if d.is_empty() => "(empty)".to_string(),
        (TriState::Uniform(_), Some(d)) => d.clone(),
        (TriState::Mixed, _) => "(mixed)".to_string(),
        _ => "—".to_string(),
    };

    // Only show preview if there's something meaningful to show
    let show_preview = !matches!(state.raw_value, TriState::Empty);

    div()
        .when(show_preview, |el| {
            el.child(
                section("Value Preview", panel_border, text_primary)
                    .child(
                        div()
                            .flex()
                            .flex_col()
                            .gap_1()
                            // Raw value
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .w(px(50.0))
                                            .text_size(px(10.0))
                                            .text_color(text_muted)
                                            .child("Raw")
                                    )
                                    .child(
                                        div()
                                            .flex_1()
                                            .px_2()
                                            .py_1()
                                            .bg(panel_border.opacity(0.2))
                                            .rounded_sm()
                                            .text_size(px(10.0))
                                            .text_color(text_primary)
                                            .overflow_hidden()
                                            .child(SharedString::from(raw_display))
                                    )
                            )
                            // Formatted display
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .w(px(50.0))
                                            .text_size(px(10.0))
                                            .text_color(text_muted)
                                            .child("Display")
                                    )
                                    .child(
                                        div()
                                            .flex_1()
                                            .px_2()
                                            .py_1()
                                            .bg(panel_border.opacity(0.2))
                                            .rounded_sm()
                                            .text_size(px(10.0))
                                            .text_color(text_primary)
                                            .overflow_hidden()
                                            .child(SharedString::from(formatted_display))
                                    )
                            )
                    )
            )
        })
}

fn render_text_style_section(
    state: &SelectionFormatState,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let bold_active = matches!(state.bold, TriState::Uniform(true));
    let bold_mixed = state.bold.is_mixed();
    let italic_active = matches!(state.italic, TriState::Uniform(true));
    let italic_mixed = state.italic.is_mixed();
    let underline_active = matches!(state.underline, TriState::Uniform(true));
    let underline_mixed = state.underline.is_mixed();

    section("Text Style", panel_border, text_primary)
        .child(
            div()
                .flex()
                .gap_2()
                // Bold button
                .child(
                    format_toggle_btn("B", bold_active, bold_mixed, true, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            let state = this.selection_format_state(cx);
                            let new_value = !matches!(state.bold, TriState::Uniform(true));
                            this.set_bold(new_value, cx);
                        }))
                )
                // Italic button
                .child(
                    format_toggle_btn("I", italic_active, italic_mixed, false, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            let state = this.selection_format_state(cx);
                            let new_value = !matches!(state.italic, TriState::Uniform(true));
                            this.set_italic(new_value, cx);
                        }))
                )
                // Underline button
                .child(
                    format_toggle_btn("U", underline_active, underline_mixed, false, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            let state = this.selection_format_state(cx);
                            let new_value = !matches!(state.underline, TriState::Uniform(true));
                            this.set_underline(new_value, cx);
                        }))
                )
        )
}

fn format_toggle_btn(
    label: &'static str,
    is_active: bool,
    is_mixed: bool,
    is_bold_style: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    // Build base div without id first
    let mut btn = div()
        .w(px(32.0))
        .h(px(28.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(13.0));

    // Style based on state
    if is_active {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent)
            .text_color(text_primary);
    } else if is_mixed {
        btn = btn
            .bg(panel_border.opacity(0.3))
            .border_color(panel_border)
            .text_color(text_muted);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border)
            .text_color(text_muted);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    // Apply bold/italic/underline styling to the button label itself
    if is_bold_style {
        btn = btn.font_weight(FontWeight::BOLD);
    }
    if label == "I" {
        btn = btn.italic();
    }
    if label == "U" {
        btn = btn.underline();
    }

    // Add child and id last (id converts to Stateful<Div>)
    btn.child(if is_mixed { "—" } else { label })
        .id(SharedString::from(format!("format-btn-{}", label)))
}

fn render_background_color_section(
    _state: &SelectionFormatState,
    text_primary: Hsla,
    _text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Color palette: None + 8 colors
    let colors: &[(&str, Option<[u8; 4]>)] = &[
        ("None", None),
        ("Yellow", Some([255, 255, 0, 255])),
        ("Green", Some([198, 239, 206, 255])),
        ("Blue", Some([189, 215, 238, 255])),
        ("Red", Some([255, 199, 206, 255])),
        ("Orange", Some([255, 235, 156, 255])),
        ("Purple", Some([204, 192, 218, 255])),
        ("Gray", Some([217, 217, 217, 255])),
        ("Cyan", Some([183, 222, 232, 255])),
    ];

    section("Background", panel_border, text_primary)
        .child(
            div()
                .flex()
                .flex_wrap()
                .gap_1()
                .children(colors.iter().map(|(name, color)| {
                    let color_value = *color;
                    let swatch_bg = color.map(|[r, g, b, _]| {
                        Hsla::from(gpui::Rgba {
                            r: r as f32 / 255.0,
                            g: g as f32 / 255.0,
                            b: b as f32 / 255.0,
                            a: 1.0,
                        })
                    }).unwrap_or(hsla(0.0, 0.0, 1.0, 1.0));

                    div()
                        .id(SharedString::from(format!("bg-color-{}", name)))
                        .size(px(24.0))
                        .rounded_sm()
                        .border_1()
                        .border_color(panel_border)
                        .bg(swatch_bg)
                        .cursor_pointer()
                        .hover(|s| s.border_color(accent))
                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                            this.set_background_color(color_value, cx);
                        }))
                }))
        )
}

fn render_font_section(
    state: &SelectionFormatState,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let font_display = match &state.font_family {
        TriState::Uniform(Some(f)) => f.clone(),
        TriState::Uniform(None) => "(Default)".to_string(),
        TriState::Mixed => "(Mixed)".to_string(),
        TriState::Empty => "(Default)".to_string(),
    };

    section("Font", panel_border, text_primary)
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                // Font name display / button
                .child(
                    div()
                        .id("format-font-btn")
                        .flex_1()
                        .px_2()
                        .py_1()
                        .bg(panel_border.opacity(0.2))
                        .border_1()
                        .border_color(panel_border)
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .hover(|s| s.bg(panel_border.opacity(0.4)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.show_font_picker(cx);
                        }))
                        .child(SharedString::from(font_display))
                )
                // Clear font button (reset to default)
                .child(
                    div()
                        .id("format-font-clear")
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .cursor_pointer()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .hover(|s| s.bg(panel_border.opacity(0.4)))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_font_family_selection(None, cx);
                        }))
                        .child("Clear")
                )
        )
}

fn render_number_format_section(
    state: &SelectionFormatState,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let current_format = &state.number_format;

    // Extract current decimals, date style, and format type
    let (current_type, current_decimals, current_date_style) = match current_format {
        TriState::Uniform(NumberFormat::General) => ("General", None, None),
        TriState::Uniform(NumberFormat::Number { decimals }) => ("Number", Some(*decimals), None),
        TriState::Uniform(NumberFormat::Currency { decimals }) => ("Currency", Some(*decimals), None),
        TriState::Uniform(NumberFormat::Percent { decimals }) => ("Percent", Some(*decimals), None),
        TriState::Uniform(NumberFormat::Date { style }) => ("Date", None, Some(*style)),
        TriState::Uniform(NumberFormat::Time) => ("Time", None, None),
        TriState::Uniform(NumberFormat::DateTime) => ("DateTime", None, None),
        TriState::Mixed => ("Mixed", None, None),
        TriState::Empty => ("General", None, None),
    };

    // Show decimal control only for numeric formats
    let show_decimals = current_decimals.is_some();
    let decimals_display = current_decimals.map(|d| d.to_string()).unwrap_or_default();

    // Show date style control only for date format
    let show_date_style = current_date_style.is_some();

    section("Number Format", panel_border, text_primary)
        // Format type buttons
        .child(
            div()
                .flex()
                .flex_wrap()
                .gap_1()
                // General
                .child(
                    format_type_btn("General", current_type == "General", current_format.is_mixed(), text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_number_format_selection(NumberFormat::General, cx);
                        }))
                )
                // Number
                .child(
                    format_type_btn("Number", current_type == "Number", current_format.is_mixed(), text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_number_format_selection(NumberFormat::Number { decimals: 2 }, cx);
                        }))
                )
                // Currency
                .child(
                    format_type_btn("Currency", current_type == "Currency", current_format.is_mixed(), text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_number_format_selection(NumberFormat::Currency { decimals: 2 }, cx);
                        }))
                )
                // Percent
                .child(
                    format_type_btn("Percent", current_type == "Percent", current_format.is_mixed(), text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_number_format_selection(NumberFormat::Percent { decimals: 2 }, cx);
                        }))
                )
                // Date
                .child(
                    format_type_btn("Date", current_type == "Date", current_format.is_mixed(), text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.set_number_format_selection(NumberFormat::Date { style: DateStyle::Short }, cx);
                        }))
                )
        )
        // Date style control (only if Date format)
        .when(show_date_style, |el| {
            el.child(
                div()
                    .flex()
                    .items_center()
                    .gap_2()
                    .mt_1()
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child("Style")
                    )
                    .child(
                        div()
                            .flex()
                            .gap_1()
                            // Short: 1/18/2026
                            .child(
                                date_style_btn("Short", matches!(current_date_style, Some(DateStyle::Short)), text_primary, text_muted, accent, panel_border)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.set_number_format_selection(NumberFormat::Date { style: DateStyle::Short }, cx);
                                    }))
                            )
                            // Long: January 18, 2026
                            .child(
                                date_style_btn("Long", matches!(current_date_style, Some(DateStyle::Long)), text_primary, text_muted, accent, panel_border)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.set_number_format_selection(NumberFormat::Date { style: DateStyle::Long }, cx);
                                    }))
                            )
                            // ISO: 2026-01-18
                            .child(
                                date_style_btn("ISO", matches!(current_date_style, Some(DateStyle::Iso)), text_primary, text_muted, accent, panel_border)
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                        this.set_number_format_selection(NumberFormat::Date { style: DateStyle::Iso }, cx);
                                    }))
                            )
                    )
            )
        })
        // Decimal places control (only if applicable)
        .when(show_decimals, |el| {
            let current_dec = current_decimals.unwrap_or(2);

            el.child(
                div()
                    .flex()
                    .items_center()
                    .gap_2()
                    .mt_1()
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child("Decimals")
                    )
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_1()
                            // Decrease button
                            .child(
                                decimal_btn("-", current_dec > 0, text_primary, text_muted, panel_border)
                                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                        this.adjust_decimals_selection(-1, cx);
                                    }))
                            )
                            // Current value
                            .child(
                                div()
                                    .w(px(24.0))
                                    .text_size(px(11.0))
                                    .text_color(text_primary)
                                    .flex()
                                    .justify_center()
                                    .child(SharedString::from(decimals_display))
                            )
                            // Increase button
                            .child(
                                decimal_btn("+", current_dec < 10, text_primary, text_muted, panel_border)
                                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                        this.adjust_decimals_selection(1, cx);
                                    }))
                            )
                    )
            )
        })
}

fn format_type_btn(
    label: &'static str,
    is_active: bool,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let mut btn = div()
        .px_2()
        .py_1()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(10.0));

    if is_active && !is_mixed {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent)
            .text_color(text_primary);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border)
            .text_color(text_muted);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    btn.child(label)
        .id(SharedString::from(format!("num-fmt-{}", label)))
}

fn date_style_btn(
    label: &'static str,
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let mut btn = div()
        .px_2()
        .py_1()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(9.0));

    if is_active {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent)
            .text_color(text_primary);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border)
            .text_color(text_muted);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    btn.child(label)
        .id(SharedString::from(format!("date-style-{}", label)))
}

fn decimal_btn(
    label: &'static str,
    enabled: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let mut btn = div()
        .w(px(22.0))
        .h(px(20.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .border_1()
        .border_color(panel_border)
        .text_size(px(12.0));

    if enabled {
        btn = btn
            .cursor_pointer()
            .text_color(text_primary)
            .hover(|s| s.bg(panel_border.opacity(0.5)));
    } else {
        btn = btn
            .text_color(text_muted.opacity(0.5));
    }

    btn.child(label)
        .id(SharedString::from(format!("decimal-{}", label)))
}

fn render_alignment_section(
    state: &SelectionFormatState,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let h_align = &state.alignment;
    let v_align = &state.vertical_alignment;
    let text_overflow = &state.text_overflow;

    let wrap_active = matches!(text_overflow, TriState::Uniform(TextOverflow::Wrap));
    let wrap_mixed = text_overflow.is_mixed();

    section("Alignment", panel_border, text_primary)
        // Horizontal alignment row
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .w(px(50.0))
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("Horiz.")
                )
                .child(
                    div()
                        .flex()
                        .gap_1()
                        // Left
                        .child(
                            align_btn("L", matches!(h_align, TriState::Uniform(Alignment::Left)), h_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_alignment_selection(Alignment::Left, cx);
                                }))
                        )
                        // Center
                        .child(
                            align_btn("C", matches!(h_align, TriState::Uniform(Alignment::Center)), h_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_alignment_selection(Alignment::Center, cx);
                                }))
                        )
                        // Right
                        .child(
                            align_btn("R", matches!(h_align, TriState::Uniform(Alignment::Right)), h_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_alignment_selection(Alignment::Right, cx);
                                }))
                        )
                )
        )
        // Vertical alignment row
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .w(px(50.0))
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("Vert.")
                )
                .child(
                    div()
                        .flex()
                        .gap_1()
                        // Top
                        .child(
                            align_btn("T", matches!(v_align, TriState::Uniform(VerticalAlignment::Top)), v_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_vertical_alignment_selection(VerticalAlignment::Top, cx);
                                }))
                        )
                        // Middle
                        .child(
                            align_btn("M", matches!(v_align, TriState::Uniform(VerticalAlignment::Middle)), v_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_vertical_alignment_selection(VerticalAlignment::Middle, cx);
                                }))
                        )
                        // Bottom
                        .child(
                            align_btn("B", matches!(v_align, TriState::Uniform(VerticalAlignment::Bottom)), v_align.is_mixed(), text_primary, text_muted, accent, panel_border)
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.set_vertical_alignment_selection(VerticalAlignment::Bottom, cx);
                                }))
                        )
                )
        )
        // Wrap text toggle
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    div()
                        .w(px(50.0))
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child("Wrap")
                )
                .child(
                    wrap_toggle_btn(wrap_active, wrap_mixed, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                            let state = this.selection_format_state(cx);
                            let new_overflow = if matches!(state.text_overflow, TriState::Uniform(TextOverflow::Wrap)) {
                                TextOverflow::Clip
                            } else {
                                TextOverflow::Wrap
                            };
                            this.set_text_overflow_selection(new_overflow, cx);
                        }))
                )
        )
}

fn align_btn(
    label: &'static str,
    is_active: bool,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let mut btn = div()
        .w(px(24.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(10.0));

    if is_active && !is_mixed {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent)
            .text_color(text_primary);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border)
            .text_color(text_muted);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    btn.child(label)
        .id(SharedString::from(format!("align-{}", label)))
}

fn wrap_toggle_btn(
    is_active: bool,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
) -> Stateful<Div> {
    let mut btn = div()
        .px_2()
        .py_1()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(10.0));

    if is_active && !is_mixed {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent)
            .text_color(text_primary);
    } else if is_mixed {
        btn = btn
            .bg(panel_border.opacity(0.3))
            .border_color(panel_border)
            .text_color(text_muted);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border)
            .text_color(text_muted);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    let label = if is_mixed { "—" } else if is_active { "On" } else { "Off" };

    btn.child(label)
        .id("wrap-toggle")
}

// Helper: Section container
fn section(title: &'static str, _border_color: Hsla, text_color: Hsla) -> Div {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(11.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_color)
                .child(title)
        )
}

// Helper: Info row (label: value)
fn info_row(label: &'static str, value: &str, label_color: Hsla, value_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .gap_2()
        .child(
            div()
                .w(px(70.0))
                .text_size(px(11.0))
                .text_color(label_color)
                .child(label)
        )
        .child(
            div()
                .flex_1()
                .text_size(px(11.0))
                .text_color(value_color)
                .overflow_hidden()
                .child(SharedString::from(value.to_string()))
        )
}

// Helper: Info row with multiline support
fn info_row_multiline(label: &'static str, value: &str, label_color: Hsla, value_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_1()
        .child(
            div()
                .text_size(px(11.0))
                .text_color(label_color)
                .child(label)
        )
        .child(
            div()
                .px_2()
                .py_1()
                .bg(label_color.opacity(0.1))
                .rounded_sm()
                .text_size(px(11.0))
                .text_color(value_color)
                .overflow_hidden()
                .child(SharedString::from(value.to_string()))
        )
}

// Helper: Clickable cell reference with hover-to-highlight
fn clickable_cell_row(
    label: &'static str,
    cell_ref: &str,
    row: usize,
    col: usize,
    label_color: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let cell_ref_owned: SharedString = cell_ref.to_string().into();

    let mut base = div()
        .id(SharedString::from(format!("nav-{}-{}", row, col)))
        .flex()
        .items_center()
        .gap_2()
        .cursor_pointer()
        .hover(|s| s.bg(label_color.opacity(0.1)))
        .rounded_sm()
        .px_1()
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.select_cell(row, col, false, cx);
            this.ensure_visible(cx);
        }))
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.inspector_hover_cell != Some((row, col)) {
                this.inspector_hover_cell = Some((row, col));
                cx.notify();
            }
        }))
        .on_mouse_up_out(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.inspector_hover_cell.is_some() {
                this.inspector_hover_cell = None;
                cx.notify();
            }
        }));

    if !label.is_empty() {
        base = base.child(
            div()
                .w(px(50.0))
                .text_size(px(11.0))
                .text_color(label_color)
                .child(label)
        );
    }

    base.child(
        div()
            .text_size(px(11.0))
            .text_color(accent)
            .child(cell_ref_owned)
    )
}

/// Cell row that triggers path trace on click (Phase 3.5b).
/// `is_input` determines trace direction:
/// - true: trace from clicked cell TO inspected cell (input → selected)
/// - false: trace from inspected cell TO clicked cell (selected → output)
fn traceable_cell_row(
    cell_ref: &str,
    clicked_row: usize,
    clicked_col: usize,
    inspected_row: usize,
    inspected_col: usize,
    is_input: bool,
    label_color: Hsla,
    accent: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let cell_ref_owned: SharedString = cell_ref.to_string().into();

    div()
        .id(SharedString::from(format!("trace-{}-{}", clicked_row, clicked_col)))
        .flex()
        .items_center()
        .gap_2()
        .cursor_pointer()
        .hover(|s| s.bg(label_color.opacity(0.1)))
        .rounded_sm()
        .px_1()
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // Trigger trace instead of navigation
            if is_input {
                // Input: trace from input → inspected cell
                this.set_trace_path(clicked_row, clicked_col, inspected_row, inspected_col, true, cx);
            } else {
                // Output: trace from inspected cell → output
                this.set_trace_path(inspected_row, inspected_col, clicked_row, clicked_col, true, cx);
            }
        }))
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.inspector_hover_cell != Some((clicked_row, clicked_col)) {
                this.inspector_hover_cell = Some((clicked_row, clicked_col));
                cx.notify();
            }
        }))
        .on_mouse_up_out(MouseButton::Left, cx.listener(|this, _, _, cx| {
            if this.inspector_hover_cell.is_some() {
                this.inspector_hover_cell = None;
                cx.notify();
            }
        }))
        .child(
            div()
                .text_size(px(11.0))
                .text_color(accent)
                .child(cell_ref_owned)
        )
}

// Get precedents from a formula string
pub fn get_precedents(formula: &str) -> Vec<(usize, usize)> {
    if let Ok(expr) = parse(formula) {
        let mut refs = extract_cell_refs(&expr);
        refs.sort();
        refs.dedup();
        refs
    } else {
        Vec::new()
    }
}

/// Compute the depth of a cell by traversing its precedents.
/// Depth 0 = raw value (no formula), Depth 1+ = formula depth.
/// Uses memoization via visited set to avoid recomputation and cycles.
fn compute_cell_depth(
    workbook: &visigrid_engine::workbook::Workbook,
    cell_id: visigrid_engine::cell_id::CellId,
    visited: &mut std::collections::HashSet<visigrid_engine::cell_id::CellId>,
) -> usize {
    // Cycle detection
    if visited.contains(&cell_id) {
        return 0;
    }
    visited.insert(cell_id);

    // Get the cell's formula
    let raw = workbook.sheets()
        .iter()
        .find(|s| s.id == cell_id.sheet)
        .map(|s| s.get_raw(cell_id.row, cell_id.col))
        .unwrap_or_default();

    // Raw values have depth 0
    if !raw.starts_with('=') {
        return 0;
    }

    // Get precedents and find max depth
    let precedents = workbook.get_precedents(cell_id.sheet, cell_id.row, cell_id.col);
    if precedents.is_empty() {
        return 1; // Formula with no cell refs (e.g., =1+2)
    }

    let max_prec_depth = precedents
        .into_iter()
        .map(|prec| compute_cell_depth(workbook, prec, visited))
        .max()
        .unwrap_or(0);

    max_prec_depth + 1
}

// Get dependents (cells that reference the given cell)
// Uses the workbook's dependency graph for O(1) lookup instead of O(n) scan.
pub fn get_dependents(app: &Spreadsheet, row: usize, col: usize, cx: &App) -> Vec<(usize, usize)> {
    let sheet_id = app.sheet(cx).id;
    let deps = app.wb(cx).get_dependents(sheet_id, row, col);

    // Filter to same-sheet cells and convert to (row, col)
    let mut dependents: Vec<(usize, usize)> = deps
        .into_iter()
        .filter(|cell_id| cell_id.sheet == sheet_id)
        .map(|cell_id| (cell_id.row, cell_id.col))
        .collect();

    dependents.sort();
    dependents
}

/// Format a SystemTime as a relative time string (e.g., "2s ago", "just now")
fn format_relative_time(time: std::time::SystemTime) -> String {
    match time.elapsed() {
        Ok(elapsed) => {
            let secs = elapsed.as_secs();
            if secs < 2 {
                "just now".to_string()
            } else if secs < 60 {
                format!("{}s ago", secs)
            } else if secs < 3600 {
                format!("{}m ago", secs / 60)
            } else {
                format!("{}h ago", secs / 3600)
            }
        }
        Err(_) => "unknown".to_string(),
    }
}

// ============================================================================
// History Tab
// ============================================================================

/// Number of history entries visible in virtual scroll window
const HISTORY_VIEW_LEN: usize = 30;

fn render_history_tab(
    app: &mut Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::app::HistoryFilterMode;
    use crate::history::HistoryDisplayEntry;

    let all_entries = app.history.display_entries();
    let total_count = all_entries.len();
    let filter_query = app.history_filter_query.clone();
    let filter_mode = app.history_filter_mode;
    let active_sheet_idx = app.sheet_index(cx);
    let selected_id = app.selected_history_id;
    let view_start = app.history_view_start;
    let is_pro = visigrid_license::is_feature_enabled("inspector");

    // Filter entries by mode first
    let mode_filtered: Vec<HistoryDisplayEntry> = all_entries
        .into_iter()
        .filter(|e| match filter_mode {
            HistoryFilterMode::All => true,
            HistoryFilterMode::CurrentSheet => {
                e.sheet_index.map_or(true, |si| si == active_sheet_idx)
            }
            HistoryFilterMode::ValidationOnly => {
                e.label.to_lowercase().contains("validation") || e.label.to_lowercase().contains("exclusion")
            }
            HistoryFilterMode::DataEditsOnly => {
                e.label.starts_with("Edit")
                    || e.label.starts_with("Paste")
                    || e.label.starts_with("Fill")
                    || e.label.starts_with("Clear")
                    || e.label.starts_with("Sort")
            }
        })
        .collect();

    // Then filter by text query (case-insensitive substring)
    let entries: Vec<HistoryDisplayEntry> = if filter_query.is_empty() {
        mode_filtered
    } else {
        let q = filter_query.to_lowercase();
        mode_filtered
            .into_iter()
            .filter(|e| {
                e.label.to_lowercase().contains(&q) || e.scope.to_lowercase().contains(&q)
            })
            .collect()
    };

    let filtered_count = entries.len();
    let is_filtered = filter_mode != HistoryFilterMode::All || !filter_query.is_empty();

    // Virtual scroll: only render visible entries
    let view_start_clamped = view_start.min(entries.len().saturating_sub(1));
    let view_end = (view_start_clamped + HISTORY_VIEW_LEN).min(entries.len());
    let visible_entries: Vec<HistoryDisplayEntry> = entries[view_start_clamped..view_end].to_vec();
    let can_scroll_up = view_start_clamped > 0;
    let can_scroll_down = view_end < entries.len();

    // Find selected entry for detail view
    let selected_entry: Option<HistoryDisplayEntry> = selected_id
        .and_then(|id| entries.iter().find(|e| e.id == id).cloned());

    // Collect entry IDs for keyboard navigation (full list, not just visible)
    let entry_ids: Vec<u64> = entries.iter().map(|e| e.id).collect();
    let entry_ids_for_scroll = entry_ids.clone();
    // Collect highlight ranges keyed by entry ID for Enter-to-jump
    let entry_highlights: std::collections::HashMap<u64, Option<(usize, usize, usize, usize, usize)>> = entries.iter()
        .map(|e| (e.id, e.sheet_index.and_then(|si| e.affected_range.map(|(sr, sc, er, ec)| (si, sr, sc, er, ec)))))
        .collect();

    // Filter mode label for banner
    let filter_label = match filter_mode {
        HistoryFilterMode::All => None,
        HistoryFilterMode::CurrentSheet => Some("This Sheet"),
        HistoryFilterMode::ValidationOnly => Some("Validation"),
        HistoryFilterMode::DataEditsOnly => Some("Data Edits"),
    };

    div()
        .id("history-tab")
        .size_full()
        .flex()
        .flex_col()
        .on_key_down(cx.listener(move |this, event: &gpui::KeyDownEvent, _, cx| {
            let key = event.keystroke.key.as_str();
            match key {
                "up" | "down" => {
                    // Navigate selection up/down
                    if entry_ids.is_empty() {
                        return;
                    }
                    let current_idx = this.selected_history_id
                        .and_then(|id| entry_ids.iter().position(|&eid| eid == id));

                    let new_idx = match (key, current_idx) {
                        ("up", Some(idx)) if idx > 0 => Some(idx - 1),
                        ("up", Some(_)) => Some(0), // Already at top
                        ("up", None) => Some(0), // Select first
                        ("down", Some(idx)) if idx < entry_ids.len() - 1 => Some(idx + 1),
                        ("down", Some(idx)) => Some(idx), // Already at bottom
                        ("down", None) => Some(0), // Select first
                        _ => None,
                    };

                    if let Some(idx) = new_idx {
                        let new_id = entry_ids[idx];
                        this.selected_history_id = Some(new_id);
                        this.history_highlight_range = entry_highlights.get(&new_id).copied().flatten();
                        // Auto-scroll to keep selection visible
                        if idx < this.history_view_start {
                            this.history_view_start = idx;
                        } else if idx >= this.history_view_start + HISTORY_VIEW_LEN {
                            this.history_view_start = idx.saturating_sub(HISTORY_VIEW_LEN - 1);
                        }
                        cx.notify();
                    }
                }
                "enter" => {
                    // Jump to affected range and select it
                    if let Some(range) = this.history_highlight_range {
                        let (sheet_idx, start_row, start_col, end_row, end_col) = range;
                        // Switch to sheet if needed
                        if sheet_idx != this.sheet_index(cx) {
                            this.wb_mut(cx, |wb| wb.set_active_sheet(sheet_idx));
                            this.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
                        }
                        // Select the full affected range
                        this.view_state.selected = (start_row, start_col);
                        if start_row != end_row || start_col != end_col {
                            this.view_state.selection_end = Some((end_row, end_col));
                        } else {
                            this.view_state.selection_end = None;
                        }
                        // Scroll to make selection visible
                        this.ensure_cell_visible(start_row, start_col);
                        cx.notify();
                    }
                }
                _ => {}
            }
        }))
        // Filter input
        .child(
            div()
                .px_3()
                .py_2()
                .border_b_1()
                .border_color(panel_border)
                .child(
                    div()
                        .id("history-filter-input")
                        .px_2()
                        .py_1()
                        .w_full()
                        .bg(panel_border.opacity(0.3))
                        .rounded_sm()
                        .text_size(px(12.0))
                        .text_color(text_primary)
                        .child(if filter_query.is_empty() {
                            div().text_color(text_muted).child("Filter...")
                        } else {
                            div().child(filter_query.clone())
                        })
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            // Simple: click to clear filter (full text input would require more work)
                            this.history_filter_query.clear();
                            cx.notify();
                        }))
                        .on_key_down(cx.listener(|this, event: &gpui::KeyDownEvent, _, cx| {
                            if let Some(ch) = &event.keystroke.key_char {
                                this.history_filter_query.push_str(ch);
                                cx.notify();
                            } else if event.keystroke.key == "backspace" && !this.history_filter_query.is_empty() {
                                this.history_filter_query.pop();
                                cx.notify();
                            } else if event.keystroke.key == "escape" {
                                this.history_filter_query.clear();
                                cx.notify();
                            }
                        }))
                )
                // Filter mode chips
                .child(
                    div()
                        .mt_2()
                        .flex()
                        .flex_wrap()
                        .gap_1()
                        .children([
                            (HistoryFilterMode::All, "All"),
                            (HistoryFilterMode::CurrentSheet, "This Sheet"),
                            (HistoryFilterMode::ValidationOnly, "Validation"),
                            (HistoryFilterMode::DataEditsOnly, "Data"),
                        ].into_iter().map(|(mode, label)| {
                            let is_active = filter_mode == mode;
                            div()
                                .id(SharedString::from(format!("history-filter-{:?}", mode)))
                                .px_2()
                                .py(px(2.0))
                                .rounded_sm()
                                .text_size(px(10.0))
                                .cursor_pointer()
                                .when(is_active, |el| {
                                    el.bg(accent.opacity(0.3))
                                        .text_color(accent)
                                        .border_1()
                                        .border_color(accent.opacity(0.5))
                                })
                                .when(!is_active, |el| {
                                    el.bg(panel_border.opacity(0.2))
                                        .text_color(text_muted)
                                        .hover(|s| s.bg(panel_border.opacity(0.4)))
                                })
                                .child(label)
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                    this.history_filter_mode = mode;
                                    cx.notify();
                                }))
                        }))
                )
        )
        // Filter banner (when filtered)
        .when(is_filtered, |el| {
            el.child(
                div()
                    .px_3()
                    .py_1()
                    .bg(accent.opacity(0.1))
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .items_center()
                    .justify_between()
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_2()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child(SharedString::from(format!("Showing {} of {}", filtered_count, total_count)))
                            .when(filter_label.is_some(), |el| {
                                el.child(
                                    div()
                                        .px_1()
                                        .rounded_sm()
                                        .bg(accent.opacity(0.2))
                                        .text_color(accent)
                                        .child(filter_label.unwrap_or(""))
                                )
                            })
                    )
                    .child(
                        div()
                            .id("clear-history-filter")
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .cursor_pointer()
                            .hover(|s| s.text_color(text_primary))
                            .child("Clear")
                            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                this.history_filter_mode = crate::app::HistoryFilterMode::All;
                                this.history_filter_query.clear();
                                this.history_view_start = 0;
                                cx.notify();
                            }))
                    )
            )
        })
        .child(
            // Entry list with virtual scroll
            div()
                .id("history-entry-list")
                .flex_1()
                .overflow_hidden()
                .on_scroll_wheel(cx.listener(move |this, event: &gpui::ScrollWheelEvent, _, cx| {
                    let delta = event.delta.pixel_delta(px(24.0));
                    let dy: f32 = delta.y.into();
                    let scroll_lines = (-dy / 24.0).round() as i32;

                    if scroll_lines > 0 {
                        // Scroll down
                        let max_start = entry_ids_for_scroll.len().saturating_sub(HISTORY_VIEW_LEN);
                        this.history_view_start = (this.history_view_start + scroll_lines as usize).min(max_start);
                    } else if scroll_lines < 0 {
                        // Scroll up
                        this.history_view_start = this.history_view_start.saturating_sub((-scroll_lines) as usize);
                    }
                    cx.notify();
                }))
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .children(visible_entries.iter().map(|entry| {
                            render_history_entry(entry, selected_id, text_primary, text_muted, panel_border, cx)
                        }))
                        .when(filtered_count == 0, |el| {
                            el.child(
                                div()
                                    .p_4()
                                    .text_size(px(12.0))
                                    .text_color(text_muted)
                                    .child(if total_count == 0 {
                                        "No history yet"
                                    } else {
                                        "No matches"
                                    })
                            )
                        })
                )
        )
        // Scroll indicator (when list is scrollable)
        .when(filtered_count > HISTORY_VIEW_LEN, |el| {
            el.child(
                div()
                    .h(px(20.0))
                    .px_3()
                    .flex()
                    .items_center()
                    .justify_between()
                    .border_t_1()
                    .border_color(panel_border)
                    .bg(panel_border.opacity(0.1))
                    .child(
                        div()
                            .text_size(px(10.0))
                            .text_color(text_muted)
                            .child(SharedString::from(format!(
                                "{}-{} of {}",
                                view_start_clamped + 1,
                                view_end,
                                filtered_count
                            )))
                    )
                    .child(
                        div()
                            .flex()
                            .gap_2()
                            .child(
                                div()
                                    .id("history-scroll-up")
                                    .text_size(px(10.0))
                                    .cursor_pointer()
                                    .text_color(if can_scroll_up { text_primary } else { text_muted.opacity(0.3) })
                                    .child("▲")
                                    .when(can_scroll_up, |el| {
                                        el.on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                            this.history_view_start = this.history_view_start.saturating_sub(HISTORY_VIEW_LEN);
                                            cx.notify();
                                        }))
                                    })
                            )
                            .child(
                                div()
                                    .id("history-scroll-down")
                                    .text_size(px(10.0))
                                    .cursor_pointer()
                                    .text_color(if can_scroll_down { text_primary } else { text_muted.opacity(0.3) })
                                    .child("▼")
                                    .when(can_scroll_down, |el| {
                                        el.on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                                            let max_start = filtered_count.saturating_sub(HISTORY_VIEW_LEN);
                                            this.history_view_start = (this.history_view_start + HISTORY_VIEW_LEN).min(max_start);
                                            cx.notify();
                                        }))
                                    })
                            )
                    )
            )
        })
        // Detail panel for selected entry
        .when(selected_entry.is_some(), |el| {
            let entry = selected_entry.unwrap();
            el.child(render_history_detail(&entry, is_pro, text_primary, text_muted, accent, panel_border, cx))
        })
}

fn render_history_entry(
    entry: &crate::history::HistoryDisplayEntry,
    selected_id: Option<u64>,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_selected = selected_id == Some(entry.id);
    let entry_id = entry.id;
    let is_undoable = entry.is_undoable;
    let location = entry.location.clone();
    let sheet_idx = entry.sheet_index;
    let ai_source = entry.ai_source.clone();

    // Capture highlight info for click handler
    let highlight_range = entry.sheet_index.and_then(|si| {
        entry.affected_range.map(|(sr, sc, er, ec)| (si, sr, sc, er, ec))
    });

    // Format relative time
    let time_str = format_instant_relative(entry.timestamp);

    // AI badge color (purple/magenta for AI)
    let ai_badge_color = hsla(0.8, 0.6, 0.55, 1.0);

    div()
        .id(SharedString::from(format!("history-entry-{}", entry.id)))
        .px_3()
        .py_2()
        .flex()
        .flex_col()
        .gap_0p5()
        .cursor_pointer()
        .bg(if is_selected { panel_border.opacity(0.5) } else { gpui::transparent_black() })
        .border_b_1()
        .border_color(panel_border.opacity(0.3))
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            // Hide context menu on left click
            this.history_context_menu_entry_id = None;
            // Toggle selection and highlight
            if this.selected_history_id == Some(entry_id) {
                this.selected_history_id = None;
                this.history_highlight_range = None;
            } else {
                this.selected_history_id = Some(entry_id);
                this.history_highlight_range = highlight_range;
            }
            cx.notify();
        }))
        .on_mouse_down(MouseButton::Right, cx.listener(move |this, _, _, cx| {
            // Show context menu for this entry
            this.show_history_context_menu(entry_id, cx);
        }))
        // Top row: label + time
        .child(
            div()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .font_weight(FontWeight::MEDIUM)
                                .text_color(if is_undoable { text_primary } else { text_muted })
                                .child(SharedString::from(entry.label.clone()))
                        )
                        .when(!is_undoable, |el| {
                            el.child(
                                div()
                                    .text_size(px(10.0))
                                    .text_color(text_muted)
                                    .child("(undone)")
                            )
                        })
                        .when(ai_source.is_some(), |el| {
                            el.child(
                                div()
                                    .px_1()
                                    .rounded_sm()
                                    .bg(ai_badge_color.opacity(0.2))
                                    .text_size(px(9.0))
                                    .text_color(ai_badge_color)
                                    .child(SharedString::from(ai_source.clone().unwrap()))
                            )
                        })
                )
                .child(
                    div()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(SharedString::from(time_str))
                )
        )
        // Bottom row: scope or location
        .when(!entry.scope.is_empty() || location.is_some(), |el| {
            el.child(
                div()
                    .flex()
                    .items_center()
                    .justify_between()
                    .child(
                        // Scope (from provenance) or empty
                        div()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .when(!entry.scope.is_empty(), |el| {
                                el.child(SharedString::from(entry.scope.clone()))
                            })
                    )
                    // Location chip (clickable to jump without selecting)
                    .when(location.is_some(), |el| {
                        let loc = location.clone().unwrap();
                        let jump_range = highlight_range;
                        el.child(
                            div()
                                .id(SharedString::from(format!("history-loc-{}", entry_id)))
                                .px_1()
                                .rounded_sm()
                                .bg(panel_border.opacity(0.3))
                                .text_size(px(9.0))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.bg(panel_border.opacity(0.5)).text_color(text_primary))
                                .child(SharedString::from(format!(
                                    "{}{}",
                                    if sheet_idx.is_some() { format!("S{}!", sheet_idx.unwrap() + 1) } else { String::new() },
                                    loc
                                )))
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _: &MouseDownEvent, _, cx| {
                                    // Jump to location (entry selection happens via parent handler)
                                    if let Some((si, sr, sc, er, ec)) = jump_range {
                                        if si != this.sheet_index(cx) {
                                            this.wb_mut(cx, |wb| wb.set_active_sheet(si));
                                            this.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
                                        }
                                        this.view_state.selected = (sr, sc);
                                        if sr != er || sc != ec {
                                            this.view_state.selection_end = Some((er, ec));
                                        } else {
                                            this.view_state.selection_end = None;
                                        }
                                        this.ensure_cell_visible(sr, sc);
                                        cx.notify();
                                    }
                                }))
                        )
                    })
            )
        })
}

fn render_history_detail(
    entry: &crate::history::HistoryDisplayEntry,
    is_pro: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let lua_code = entry.lua.clone();
    let generated_lua = entry.generated_lua.clone();
    let summary = entry.summary.clone();
    let ai_source = entry.ai_source.clone();
    let has_changes = !entry.affected_cells.is_empty();
    let entry_is_undoable = entry.is_undoable;

    // AI badge color
    let ai_badge_color = hsla(0.8, 0.6, 0.55, 1.0);

    // Determine which Lua to use (explicit provenance takes priority)
    let copyable_lua = lua_code.clone().or_else(|| generated_lua.clone());
    let has_lua = copyable_lua.is_some();

    // Format cell address from row/col (e.g., "A1", "B2")
    fn cell_addr(row: usize, col: usize) -> String {
        let col_letter = if col < 26 {
            ((b'A' + col as u8) as char).to_string()
        } else {
            let first = (b'A' + (col / 26 - 1) as u8) as char;
            let second = (b'A' + (col % 26) as u8) as char;
            format!("{}{}", first, second)
        };
        format!("{}{}", col_letter, row + 1)
    }

    // Show up to 5 changes - clone data to own it
    let changes_to_show: Vec<(usize, usize, String, String)> = entry
        .affected_cells
        .iter()
        .take(5)
        .map(|(r, c, o, n)| (*r, *c, o.clone(), n.clone()))
        .collect();
    let more_count = entry.affected_cells.len().saturating_sub(5);

    div()
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .max_h(px(250.0))
        .overflow_hidden()
        // AI source (when this is an AI-generated mutation)
        .when(ai_source.is_some(), |el: Div| {
            let source_label = ai_source.clone().unwrap();
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .items_center()
                    .gap_2()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_muted)
                            .child("Source")
                    )
                    .child(
                        div()
                            .px_2()
                            .py_1()
                            .rounded_sm()
                            .bg(ai_badge_color.opacity(0.15))
                            .text_size(px(11.0))
                            .text_color(ai_badge_color)
                            .child(SharedString::from(source_label))
                    )
            )
        })
        // Action summary (when available)
        .when(summary.is_some(), |el: Div| {
            let summary_text = summary.clone().unwrap();
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .flex_col()
                    .gap_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_muted)
                            .child("Details")
                    )
                    .child(
                        div()
                            .text_size(px(11.0))
                            .text_color(text_primary)
                            .child(SharedString::from(summary_text))
                    )
            )
        })
        // Changes section (when there are affected cells)
        .when(has_changes, |el: Div| {
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .flex()
                    .flex_col()
                    .gap_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_muted)
                            .child("Changes")
                    )
                    .children(changes_to_show.into_iter().map(|(row, col, old, new)| {
                        let addr = cell_addr(row, col);
                        div()
                            .flex()
                            .items_center()
                            .gap_2()
                            .text_size(px(11.0))
                            .child(
                                div()
                                    .text_color(text_muted)
                                    .min_w(px(40.0))
                                    .child(SharedString::from(addr))
                            )
                            .child(
                                div()
                                    .text_color(text_muted)
                                    .max_w(px(80.0))
                                    .overflow_hidden()
                                    .child(SharedString::from(if old.is_empty() { "(empty)".to_string() } else { old }))
                            )
                            .child(
                                div()
                                    .text_color(text_muted)
                                    .child("→")
                            )
                            .child(
                                div()
                                    .text_color(text_primary)
                                    .max_w(px(80.0))
                                    .overflow_hidden()
                                    .child(SharedString::from(if new.is_empty() { "(empty)".to_string() } else { new }))
                            )
                    }))
                    .when(more_count > 0, |el| {
                        el.child(
                            div()
                                .text_size(px(10.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!("...and {} more", more_count)))
                        )
                    })
            )
        })
        // Provenance section (when there's Lua code)
        .when(lua_code.is_some(), |el: Div| {
            let code = lua_code.clone().unwrap();
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .flex()
                    .flex_col()
                    .gap_1()
                    .child(
                        div()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_muted)
                            .child("Provenance")
                    )
                    .child(
                        if is_pro {
                            div()
                                .p_2()
                                .rounded_md()
                                .bg(rgb(0x1a1a1a))
                                .border_1()
                                .border_color(panel_border)
                                .overflow_hidden()
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .font_family("monospace")
                                        .text_color(text_primary)
                                        .child(SharedString::from(code))
                                )
                                .into_any_element()
                        } else {
                            div()
                                .text_size(px(10.0))
                                .text_color(accent)
                                .child("View Lua with Pro")
                                .into_any_element()
                        }
                    )
            )
        })
        // Empty state
        .when(!has_changes && lua_code.is_none(), |el: Div| {
            el.child(
                div()
                    .px_3()
                    .py_2()
                    .text_size(px(11.0))
                    .text_color(text_muted)
                    .child("No details available")
            )
        })
        // Action buttons (Copy Lua + Rewind to here)
        .when(entry_is_undoable || has_lua, |el: Div| {
            el.child({
                let lua_for_copy = copyable_lua.clone();
                div()
                    .px_3()
                    .py_2()
                    .border_t_1()
                    .border_color(panel_border)
                    .flex()
                    .gap_2()
                    // Copy Lua button (when Lua is available)
                    .when(has_lua, |el: Div| {
                        el.child(
                            div()
                                .id("copy-lua-btn")
                                .px_3()
                                .py_1()
                                .rounded_md()
                                .bg(accent.opacity(0.15))
                                .border_1()
                                .border_color(accent.opacity(0.3))
                                .text_size(px(11.0))
                                .text_color(accent)
                                .cursor_pointer()
                                .hover(|s| s.bg(accent.opacity(0.25)))
                                .child("Copy Lua")
                                .on_mouse_down(MouseButton::Left, {
                                    let lua = lua_for_copy.clone();
                                    cx.listener(move |this, _, _, cx| {
                                        if let Some(ref code) = lua {
                                            cx.write_to_clipboard(gpui::ClipboardItem::new_string(code.clone()));
                                            this.status_message = Some("Lua copied to clipboard".to_string());
                                            cx.notify();
                                        }
                                    })
                                })
                        )
                    })
                    // Rewind to here button (when entry is undoable)
                    .when(entry_is_undoable, |el: Div| {
                        el.child(
                            div()
                                .id("rewind-to-here-btn")
                                .px_3()
                                .py_1()
                                .rounded_md()
                                .bg(hsla(0.0, 0.8, 0.3, 0.2))
                                .border_1()
                                .border_color(hsla(0.0, 0.8, 0.4, 0.5))
                                .text_size(px(11.0))
                                .text_color(hsla(0.0, 0.8, 0.7, 1.0))
                                .cursor_pointer()
                                .hover(|s| s.bg(hsla(0.0, 0.8, 0.3, 0.4)))
                                .child("Rewind to here...")
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.show_rewind_confirm(cx);
                                }))
                        )
                    })
            })
        })
}

/// Format an Instant as relative time (e.g., "12s ago", "5m ago")
fn format_instant_relative(instant: std::time::Instant) -> String {
    let elapsed = instant.elapsed();
    let secs = elapsed.as_secs();

    if secs < 2 {
        "just now".to_string()
    } else if secs < 60 {
        format!("{}s ago", secs)
    } else if secs < 3600 {
        format!("{}m ago", secs / 60)
    } else if secs < 86400 {
        format!("{}h ago", secs / 3600)
    } else {
        format!("{}d ago", secs / 86400)
    }
}

/// Render history entry context menu overlay (right-click menu)
pub fn render_history_context_menu(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    use crate::theme::TokenKey;

    let entry_id = app.history_context_menu_entry_id?;

    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);

    // Check entry is still valid
    let entry_index = app.history.global_index_for_id(entry_id)?;
    let is_in_undo_stack = app.history.entry_at(entry_index).is_some();

    // Menu items depend on whether entry is in undo stack (can diff) or redo stack
    let can_explain_diff = is_in_undo_stack && entry_index < app.history.undo_count().saturating_sub(1);

    Some(div()
        .id("history-context-menu")
        .absolute()
        .right(px(10.0))
        .top(px(100.0))
        .w(px(200.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_md()
        .shadow_lg()
        .flex()
        .flex_col()
        .py_1()
        // Click outside to close
        .on_mouse_down_out(cx.listener(|this, _, _, cx| {
            this.hide_history_context_menu(cx);
        }))
        // Explain changes since this
        .when(can_explain_diff, |el| {
            el.child(
                div()
                    .id("explain-diff-item")
                    .px_3()
                    .py_1()
                    .text_size(px(12.0))
                    .text_color(text_primary)
                    .cursor_pointer()
                    .hover(|s| s.bg(panel_border.opacity(0.3)))
                    .child("Explain changes since this...")
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        this.show_explain_diff(entry_id, cx);
                    }))
            )
        })
        .when(!can_explain_diff, |el| {
            el.child(
                div()
                    .px_3()
                    .py_1()
                    .text_size(px(12.0))
                    .text_color(text_muted)
                    .child("(No changes after this entry)")
            )
        })
    )
}

/// Render the Explain Differences dialog
pub fn render_explain_diff_dialog(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::theme::TokenKey;

    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);

    let report = match &app.diff_report {
        Some(r) => r.clone(),
        None => return div().into_any_element(),
    };

    let ai_only = app.diff_ai_only_filter;
    let selected_entry = app.diff_selected_entry;
    let ai_count = report.ai_touched_count();
    let has_ai_changes = ai_count > 0;

    // Get sheet names for formatting addresses
    let sheet_names: Vec<String> = app.wb(cx).sheet_names()
        .iter()
        .map(|s| s.to_string())
        .collect();

    // AI badge color
    let ai_badge_color = hsla(0.8, 0.6, 0.55, 1.0);

    // Filter changes based on AI filter
    let value_changes: Vec<&crate::diff::DiffEntry> = report.value_changes_filtered(ai_only);
    let formula_changes: Vec<&crate::diff::DiffEntry> = report.formula_changes_filtered(ai_only);
    let value_changes_empty = value_changes.is_empty();
    let formula_changes_empty = formula_changes.is_empty();

    // AI summary state
    let ai_summary = app.diff_ai_summary.clone();
    let ai_summary_loading = app.diff_ai_summary_loading;
    let ai_summary_error = app.diff_ai_summary_error.clone();

    // Check if AI summary is available (provider configured with ask capability)
    let ai_config = visigrid_config::ai::ResolvedAIConfig::load();
    let ai_summary_available = ai_config.provider.capabilities().ask;
    let ai_explain_available = ai_config.provider.capabilities().ask;

    // Entry explanation state
    let entry_explanations = app.diff_entry_explanations.clone();
    let explaining_entry = app.diff_explaining_entry;

    div()
        .id("explain-diff-dialog")
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.close_explain_diff(cx);
        }))
        .child(
            div()
                .id("explain-diff-content")
                .w(px(600.0))
                .max_h(px(500.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_lg()
                .shadow_lg()
                .flex()
                .flex_col()
                .overflow_hidden()
                // Header
                .child(
                    div()
                        .px_4()
                        .py_3()
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .flex()
                                .flex_col()
                                .gap_1()
                                .child(
                                    div()
                                        .text_size(px(14.0))
                                        .font_weight(FontWeight::SEMIBOLD)
                                        .text_color(text_primary)
                                        .child("Changes since")
                                )
                                .child(
                                    div()
                                        .text_size(px(12.0))
                                        .text_color(text_muted)
                                        .child(SharedString::from(format!(
                                            "\"{}\" ({} actions)",
                                            report.since_entry_label,
                                            report.entries_spanned
                                        )))
                                )
                                .child(
                                    div()
                                        .text_size(px(10.0))
                                        .text_color(text_muted.opacity(0.7))
                                        .child("Click to jump • Enter to close")
                                )
                        )
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                // Copy Report button
                                .child(
                                    div()
                                        .id("copy-diff-report")
                                        .px_2()
                                        .py_1()
                                        .rounded_sm()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .cursor_pointer()
                                        .bg(panel_border.opacity(0.2))
                                        .hover(|s| s.bg(panel_border.opacity(0.4)).text_color(text_primary))
                                        .child("Copy Report")
                                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _: &MouseDownEvent, _, cx| {
                                            this.copy_diff_report(cx);
                                        }))
                                )
                                // Close button
                                .child(
                                    div()
                                        .id("close-explain-diff")
                                        .px_2()
                                        .py_1()
                                        .rounded_sm()
                                        .text_size(px(12.0))
                                        .text_color(text_muted)
                                        .cursor_pointer()
                                        .hover(|s| s.bg(panel_border.opacity(0.3)).text_color(text_primary))
                                        .child("×")
                                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _: &MouseDownEvent, _, cx| {
                                            this.close_explain_diff(cx);
                                        }))
                                )
                        )
                )
                // Keyboard handling: Enter = jump + close
                .on_key_down(cx.listener(|this, event: &gpui::KeyDownEvent, _, cx| {
                    if event.keystroke.key == "enter" {
                        this.diff_jump_and_close(cx);
                    }
                }))
                // AI filter toggle
                .when(has_ai_changes, |el| {
                    el.child(
                        div()
                            .px_4()
                            .py_2()
                            .border_b_1()
                            .border_color(panel_border)
                            .flex()
                            .items_center()
                            .gap_2()
                            .child(
                                div()
                                    .id("ai-filter-toggle")
                                    .px_2()
                                    .py_1()
                                    .rounded_sm()
                                    .cursor_pointer()
                                    .bg(if ai_only { ai_badge_color.opacity(0.2) } else { panel_border.opacity(0.2) })
                                    .text_size(px(11.0))
                                    .text_color(if ai_only { ai_badge_color } else { text_muted })
                                    .hover(|s| s.bg(ai_badge_color.opacity(0.3)))
                                    .child(SharedString::from(format!("AI-touched only ({})", ai_count)))
                                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _: &MouseDownEvent, _, cx| {
                                        this.toggle_diff_ai_filter(cx);
                                    }))
                            )
                    )
                })
                // AI Summary section (collapsible)
                .when(ai_summary_available, |el| {
                    el.child(
                        div()
                            .px_4()
                            .py_2()
                            .border_b_1()
                            .border_color(panel_border)
                            .flex()
                            .flex_col()
                            .gap_2()
                            // Header row with button
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .justify_between()
                                    .child(
                                        div()
                                            .text_size(px(11.0))
                                            .font_weight(FontWeight::MEDIUM)
                                            .text_color(text_primary)
                                            .child("AI Summary")
                                    )
                                    .when(ai_summary.is_none() && !ai_summary_loading, |el| {
                                        el.child(
                                            div()
                                                .id("generate-summary-btn")
                                                .px_2()
                                                .py_1()
                                                .rounded_sm()
                                                .text_size(px(10.0))
                                                .text_color(ai_badge_color)
                                                .cursor_pointer()
                                                .bg(ai_badge_color.opacity(0.1))
                                                .hover(|s| s.bg(ai_badge_color.opacity(0.2)))
                                                .child("Generate Summary")
                                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _: &MouseDownEvent, _, cx| {
                                                    this.generate_diff_summary(cx);
                                                }))
                                        )
                                    })
                                    .when(ai_summary_loading, |el| {
                                        el.child(
                                            div()
                                                .text_size(px(10.0))
                                                .text_color(text_muted)
                                                .child("Generating...")
                                        )
                                    })
                                    .when(ai_summary.is_some(), |el| {
                                        el.child(
                                            div()
                                                .id("copy-summary-btn")
                                                .px_2()
                                                .py_1()
                                                .rounded_sm()
                                                .text_size(px(10.0))
                                                .text_color(text_muted)
                                                .cursor_pointer()
                                                .bg(panel_border.opacity(0.2))
                                                .hover(|s| s.bg(panel_border.opacity(0.4)).text_color(text_primary))
                                                .child("Copy")
                                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _: &MouseDownEvent, _, cx| {
                                                    this.copy_diff_summary(cx);
                                                }))
                                        )
                                    })
                            )
                            // Error display
                            .when(ai_summary_error.is_some(), |el| {
                                el.child(
                                    div()
                                        .text_size(px(11.0))
                                        .text_color(hsla(0.0, 0.7, 0.5, 1.0))  // Red
                                        .child(SharedString::from(ai_summary_error.clone().unwrap()))
                                )
                            })
                            // Summary text
                            .when(ai_summary.is_some(), |el| {
                                el.child(
                                    div()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .child(SharedString::from(ai_summary.clone().unwrap()))
                                )
                            })
                    )
                })
                // Content area
                .child(
                    div()
                        .flex_1()
                        .overflow_hidden()
                        .p_4()
                        .flex()
                        .flex_col()
                        .gap_4()
                        // Stats summary
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child(SharedString::from(format!(
                                    "{} value changes, {} formula changes, {} structural, {} format",
                                    report.value_changes.len(),
                                    report.formula_changes.len(),
                                    report.structural_changes.len() + report.named_range_changes.len() + report.validation_changes.len(),
                                    report.format_change_count
                                )))
                        )
                        // Value changes section
                        .when(!value_changes_empty, |el| {
                            let sheet_names_clone = sheet_names.clone();
                            let explanations_clone = entry_explanations.clone();
                            el.child(render_diff_section(
                                "Values",
                                value_changes,
                                sheet_names_clone,
                                selected_entry,
                                accent,
                                text_primary,
                                text_muted,
                                ai_badge_color,
                                panel_border,
                                ai_explain_available,
                                explanations_clone,
                                explaining_entry,
                                cx,
                            ))
                        })
                        // Formula changes section
                        .when(!formula_changes_empty, |el| {
                            let sheet_names_clone = sheet_names.clone();
                            let explanations_clone = entry_explanations.clone();
                            el.child(render_diff_section(
                                "Formulas",
                                formula_changes,
                                sheet_names_clone,
                                selected_entry,
                                accent,
                                text_primary,
                                text_muted,
                                ai_badge_color,
                                panel_border,
                                ai_explain_available,
                                explanations_clone,
                                explaining_entry,
                                cx,
                            ))
                        })
                        // Structural changes section
                        .when(!report.structural_changes.is_empty(), |el| {
                            el.child(render_structural_section(
                                &report.structural_changes,
                                text_primary,
                                text_muted,
                                panel_border,
                            ))
                        })
                        // Named range changes
                        .when(!report.named_range_changes.is_empty(), |el| {
                            el.child(render_named_range_section(
                                &report.named_range_changes,
                                text_primary,
                                text_muted,
                                panel_border,
                            ))
                        })
                        // Validation changes
                        .when(!report.validation_changes.is_empty(), |el| {
                            el.child(render_validation_section(
                                &report.validation_changes,
                                text_primary,
                                text_muted,
                                panel_border,
                            ))
                        })
                        // Empty state
                        .when(ai_only && value_changes_empty && formula_changes_empty, |el| {
                            el.child(
                                div()
                                    .p_4()
                                    .text_size(px(12.0))
                                    .text_color(text_muted)
                                    .child("No AI-touched changes in this range.")
                            )
                        })
                        .when(!ai_only && report.total_changes() == 0, |el| {
                            el.child(
                                div()
                                    .p_4()
                                    .text_size(px(12.0))
                                    .text_color(text_muted)
                                    .child("No changes since this point.")
                            )
                        })
                )
        )
        .into_any_element()
}

/// Render a section of cell changes (values or formulas)
fn render_diff_section(
    title: &str,
    entries: Vec<&crate::diff::DiffEntry>,
    sheet_names: Vec<String>,
    selected_entry: Option<(usize, usize, usize)>,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    ai_badge_color: Hsla,
    panel_border: Hsla,
    ai_explain_available: bool,
    entry_explanations: std::collections::HashMap<(usize, usize, usize), String>,
    explaining_entry: Option<(usize, usize, usize)>,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child(SharedString::from(format!("{} ({})", title, entries.len())))
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(entries.into_iter().take(50).map(|entry| {
                    let sheet_idx = entry.sheet_index;
                    let row = entry.row;
                    let col = entry.col;
                    let address = entry.full_address(&sheet_names);
                    let old_val = entry.old_value.clone();
                    let new_val = entry.new_value.clone();
                    let is_ai = entry.ai_touched;
                    let ai_source = entry.ai_source.clone();
                    let is_selected = selected_entry == Some((sheet_idx, row, col));
                    let key = (sheet_idx, row, col);
                    let explanation = entry_explanations.get(&key).cloned();
                    let is_explaining = explaining_entry == Some(key);
                    let has_explanation = explanation.is_some();

                    div()
                        .flex()
                        .flex_col()
                        .gap_1()
                        // Main entry row
                        .child(
                            div()
                                .id(SharedString::from(format!("diff-{}-{}-{}", sheet_idx, row, col)))
                                .px_2()
                                .py_1()
                                .rounded_sm()
                                .bg(if is_selected { accent.opacity(0.3) } else { panel_border.opacity(0.1) })
                                .border_1()
                                .border_color(if is_selected { accent } else { gpui::transparent_black() })
                                .flex()
                                .items_center()
                                .gap_2()
                                .cursor_pointer()
                                .hover(|s| s.bg(if is_selected { accent.opacity(0.4) } else { panel_border.opacity(0.3) }))
                                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _: &MouseDownEvent, _, cx| {
                                    this.diff_jump_to_cell(sheet_idx, row, col, cx);
                                }))
                                // Address
                                .child(
                                    div()
                                        .text_size(px(11.0))
                                        .font_weight(FontWeight::MEDIUM)
                                        .text_color(text_primary)
                                        .min_w(px(80.0))
                                        .child(SharedString::from(address))
                                )
                                // Change: old → new
                                .child(
                                    div()
                                        .flex_1()
                                        .text_size(px(11.0))
                                        .text_color(text_muted)
                                        .overflow_hidden()
                                        .child(SharedString::from(format!(
                                            "{} → {}",
                                            if old_val.is_empty() { "(empty)" } else { &old_val },
                                            if new_val.is_empty() { "(empty)" } else { &new_val }
                                        )))
                                )
                                // AI badge
                                .when(is_ai, |el| {
                                    el.child(
                                        div()
                                            .px_1()
                                            .rounded_sm()
                                            .bg(ai_badge_color.opacity(0.2))
                                            .text_size(px(9.0))
                                            .text_color(ai_badge_color)
                                            .child("AI")
                                    )
                                })
                                // Explain button (only on selected entry, if AI available, and not already explained)
                                .when(is_selected && ai_explain_available && !has_explanation && !is_explaining, |el| {
                                    let old_for_explain = old_val.clone();
                                    let new_for_explain = new_val.clone();
                                    let ai_source_for_explain = ai_source.clone();
                                    el.child(
                                        div()
                                            .id(SharedString::from(format!("explain-{}-{}-{}", sheet_idx, row, col)))
                                            .px_2()
                                            .py_px()
                                            .rounded_sm()
                                            .text_size(px(9.0))
                                            .text_color(ai_badge_color)
                                            .cursor_pointer()
                                            .bg(ai_badge_color.opacity(0.1))
                                            .hover(|s| s.bg(ai_badge_color.opacity(0.2)))
                                            .child("Explain")
                                            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _: &MouseDownEvent, _, cx| {
                                                this.explain_diff_entry(
                                                    sheet_idx,
                                                    row,
                                                    col,
                                                    old_for_explain.clone(),
                                                    new_for_explain.clone(),
                                                    is_ai,
                                                    ai_source_for_explain.clone(),
                                                    cx,
                                                );
                                            }))
                                    )
                                })
                                // Loading indicator
                                .when(is_explaining, |el| {
                                    el.child(
                                        div()
                                            .text_size(px(9.0))
                                            .text_color(text_muted)
                                            .child("...")
                                    )
                                })
                        )
                        // Inline explanation (below entry, if available)
                        .when(has_explanation, |el| {
                            let explanation_text = explanation.clone().unwrap_or_default();
                            let explanation_for_copy = explanation_text.clone();
                            el.child(
                                div()
                                    .ml_2()
                                    .px_2()
                                    .py_1()
                                    .rounded_sm()
                                    .bg(ai_badge_color.opacity(0.05))
                                    .border_l_2()
                                    .border_color(ai_badge_color.opacity(0.3))
                                    .flex()
                                    .items_start()
                                    .gap_2()
                                    // Explanation text
                                    .child(
                                        div()
                                            .flex_1()
                                            .text_size(px(10.0))
                                            .text_color(text_muted)
                                            .child(SharedString::from(explanation_text))
                                    )
                                    // Copy button
                                    .child(
                                        div()
                                            .id(SharedString::from(format!("copy-explain-{}-{}-{}", sheet_idx, row, col)))
                                            .px_1()
                                            .py_px()
                                            .rounded_sm()
                                            .text_size(px(9.0))
                                            .text_color(text_muted)
                                            .cursor_pointer()
                                            .bg(panel_border.opacity(0.2))
                                            .hover(|s| s.bg(panel_border.opacity(0.4)).text_color(text_primary))
                                            .child("Copy")
                                            .on_mouse_down(MouseButton::Left, cx.listener(move |_this, _: &MouseDownEvent, _, cx| {
                                                cx.write_to_clipboard(gpui::ClipboardItem::new_string(explanation_for_copy.clone()));
                                            }))
                                    )
                            )
                        })
                }))
        )
}

/// Render structural changes section
fn render_structural_section(
    changes: &[crate::diff::StructuralChange],
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child(SharedString::from(format!("Structural ({})", changes.len())))
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(changes.iter().map(|change| {
                    div()
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.1))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(SharedString::from(change.description()))
                }))
        )
}

/// Render named range changes section
fn render_named_range_section(
    changes: &[crate::diff::NamedRangeChange],
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child(SharedString::from(format!("Named Ranges ({})", changes.len())))
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(changes.iter().map(|change| {
                    div()
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.1))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(SharedString::from(change.description()))
                }))
        )
}

/// Render validation changes section
fn render_validation_section(
    changes: &[crate::diff::ValidationChange],
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(12.0))
                .font_weight(FontWeight::MEDIUM)
                .text_color(text_primary)
                .child(SharedString::from(format!("Validation ({})", changes.len())))
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .children(changes.iter().map(|change| {
                    div()
                        .px_2()
                        .py_1()
                        .rounded_sm()
                        .bg(panel_border.opacity(0.1))
                        .text_size(px(11.0))
                        .text_color(text_muted)
                        .child(SharedString::from(change.description()))
                }))
        )
}
