use gpui::*;
use gpui::prelude::FluentBuilder;
use visigrid_engine::cell::NumberFormat;
use visigrid_engine::cell_id::CellId;
use crate::app::Spreadsheet;
use crate::theme::TokenKey;
use crate::views::inspector_panel;

/// Render the F1 hold-to-peek context help overlay
pub(crate) fn render_f1_help_overlay(app: &Spreadsheet, cx: &App) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let accent = app.token(TokenKey::Accent);

    // Build content based on context
    let content = if let Some(sig_info) = app.signature_help() {
        // In formula mode with a function: show full signature help
        let func = sig_info.function;
        let current_arg = sig_info.current_arg;

        let params: Vec<_> = func.parameters.iter().enumerate().map(|(i, param)| {
            let is_current = i == current_arg;
            div()
                .flex()
                .items_center()
                .gap_3()
                .py(px(4.0))
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(if is_current { accent } else { text_muted })
                        .font_weight(if is_current { FontWeight::SEMIBOLD } else { FontWeight::NORMAL })
                        .min_w(px(90.0))
                        .child(param.name)
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(text_disabled)
                        .child(param.description)
                )
        }).collect();

        div()
            .flex()
            .flex_col()
            // Header: function name
            .child(
                div()
                    .px_3()
                    .py(px(10.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .child(
                        div()
                            .text_size(px(14.0))
                            .text_color(text_primary)
                            .font_weight(FontWeight::SEMIBOLD)
                            .child(func.name)
                    )
            )
            // Signature
            .child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .text_size(px(13.0))
                    .text_color(text_muted)
                    .child(func.signature)
            )
            // Description
            .child(
                div()
                    .px_3()
                    .pb(px(8.0))
                    .text_size(px(12.0))
                    .text_color(text_disabled)
                    .child(func.description)
            )
            // Parameters
            .when(!func.parameters.is_empty(), |d| {
                d.child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_t_1()
                        .border_color(panel_border)
                        .flex()
                        .flex_col()
                        .children(params)
                )
            })
    } else if app.is_multi_selection() {
        // Multi-cell selection: show range stats
        let ((min_row, min_col), (max_row, max_col)) = app.selection_range();
        let start_ref = app.cell_ref_at(min_row, min_col);
        let end_ref = app.cell_ref_at(max_row, max_col);
        let range_ref = format!("{}:{}", start_ref, end_ref);

        let cell_count = (max_row - min_row + 1) * (max_col - min_col + 1);

        // Calculate stats (skip for very large selections to avoid freezing)
        let mut count = 0usize;
        let mut numeric_count = 0usize;
        let mut sum = 0.0f64;
        let mut min_val: Option<f64> = None;
        let mut max_val: Option<f64> = None;

        if cell_count <= 10_000 {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let display = app.sheet(cx).get_display(row, col);
                    if !display.is_empty() {
                        count += 1;
                        // Try to parse as number (handles both values and formula results)
                        let clean = display.replace(',', "").replace('$', "").replace('%', "");
                        if let Ok(num) = clean.parse::<f64>() {
                            numeric_count += 1;
                            sum += num;
                            min_val = Some(min_val.map_or(num, |m| m.min(num)));
                            max_val = Some(max_val.map_or(num, |m| m.max(num)));
                        }
                    }
                }
            }
        }
        let average = if numeric_count > 0 { Some(sum / numeric_count as f64) } else { None };

        // Helper to format numbers with thousands separators
        let fmt_num = |n: f64| -> String {
            let base = if n.fract() == 0.0 {
                format!("{}", n as i64)
            } else {
                format!("{:.2}", n)
            };
            // Add thousands separators
            let parts: Vec<&str> = base.split('.').collect();
            let int_part = parts[0];
            let dec_part = parts.get(1);
            let negative = int_part.starts_with('-');
            let digits: String = int_part.chars().filter(|c| c.is_ascii_digit()).collect();
            let with_commas: String = digits
                .chars()
                .rev()
                .enumerate()
                .map(|(i, c)| if i > 0 && i % 3 == 0 { format!(",{}", c) } else { c.to_string() })
                .collect::<Vec<_>>()
                .into_iter()
                .rev()
                .collect();
            let result = if negative { format!("-{}", with_commas) } else { with_commas };
            if let Some(dec) = dec_part {
                format!("{}.{}", result, dec)
            } else {
                result
            }
        };

        let mut content = div()
            .flex()
            .flex_col()
            // Header: range reference
            .child(
                div()
                    .px_3()
                    .py(px(10.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .child(
                        div()
                            .text_size(px(14.0))
                            .text_color(text_primary)
                            .font_weight(FontWeight::SEMIBOLD)
                            .child(range_ref)
                    )
            )
            // Cell count
            .child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .justify_between()
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_muted)
                            .child("Cells")
                    )
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_primary)
                            .child(format!("{}", cell_count))
                    )
            );

        // Count (non-empty)
        if count > 0 {
            content = content.child(
                div()
                    .px_3()
                    .py(px(8.0))
                    .border_b_1()
                    .border_color(panel_border)
                    .flex()
                    .justify_between()
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_muted)
                            .child("Count")
                    )
                    .child(
                        div()
                            .text_size(px(12.0))
                            .text_color(text_primary)
                            .child(format!("{}", count))
                    )
            );
        }

        // Numeric stats
        if numeric_count > 0 {
            content = content
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Sum")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(sum))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Average")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(average.unwrap()))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(panel_border)
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Min")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(min_val.unwrap()))
                        )
                )
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .flex()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Max")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .child(fmt_num(max_val.unwrap()))
                        )
                );
        }

        content
    } else {
        // Single cell inspector
        let (row, col) = app.view_state.selected;
        let cell_ref = app.cell_ref_at(row, col);
        let raw_value = app.sheet(cx).get_raw(row, col);
        let display_value = app.sheet(cx).get_display(row, col);
        let is_formula = raw_value.starts_with('=');
        let format = app.sheet(cx).get_format(row, col);

        // Get dependents (always useful)
        let dependents = inspector_panel::get_dependents(app, row, col, cx);

        // Build format badges
        let mut format_badges: Vec<&str> = Vec::new();
        match &format.number_format {
            NumberFormat::Number { .. } => format_badges.push("Number"),
            NumberFormat::Currency { .. } => format_badges.push("Currency"),
            NumberFormat::Percent { .. } => format_badges.push("Percent"),
            NumberFormat::Date { .. } => format_badges.push("Date"),
            NumberFormat::Time => format_badges.push("Time"),
            NumberFormat::DateTime => format_badges.push("DateTime"),
            NumberFormat::Custom(_) => format_badges.push("Custom"),
            NumberFormat::General => {}
        }
        if format.bold { format_badges.push("Bold"); }
        if format.italic { format_badges.push("Italic"); }
        if format.underline { format_badges.push("Underline"); }

        // Get precedents for formulas
        let precedents = if is_formula {
            inspector_panel::get_precedents(&raw_value)
        } else {
            Vec::new()
        };

        if is_formula {
            // Formula cell: value-first status card design
            // Goal: reassurance + confidence, not explanation

            // Get depth for complexity label (only when verified mode is on)
            let depth = if app.verified_mode {
                if let Some(report) = &app.last_recalc_report {
                    let sheet_id = app.sheet(cx).id;
                    let cell_id = CellId::new(sheet_id, row, col);
                    report.get_cell_info(&cell_id).map(|info| info.depth)
                } else {
                    None
                }
            } else {
                None
            };

            let complexity_label = match depth {
                Some(1) => "Simple formula".to_string(),
                Some(2) => "2 layers deep".to_string(),
                Some(d) if d <= 4 => format!("{} layers deep", d),
                Some(d) => format!("Complex ({} layers)", d),
                None => "Formula".to_string(),
            };

            // Softer divider color (reduced opacity)
            let divider = panel_border.opacity(0.5);

            let mut content = div()
                .flex()
                .flex_col()
                // Header: cell ref + context-aware verified badge
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(cell_ref.clone())
                        )
                        // No verification badges in F1 Peek - that's Pro Inspector territory
                        // F1 = "what is this cell", Pro Inspector = "why you can trust this"
                )
                // VALUE - the hero (large, prominent)
                .child(
                    div()
                        .px_3()
                        .py(px(12.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(18.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child(if display_value.is_empty() { "(empty)".to_string() } else { display_value.clone() })
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mt(px(2.0))
                                .child(complexity_label)
                        )
                )
                // Formula (secondary, smaller)
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(raw_value.clone())
                        )
                );

            // Uses (precedents) - human language
            if !precedents.is_empty() {
                let prec_refs: Vec<String> = precedents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let prec_text = if precedents.len() > 6 {
                    format!("{} +{} more", prec_refs.join(", "), precedents.len() - 6)
                } else {
                    prec_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Uses")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(prec_text)
                        )
                );
            }

            // Feeds (dependents) - human language
            if !dependents.is_empty() {
                let dep_refs: Vec<String> = dependents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let dep_text = if dependents.len() > 6 {
                    format!("{} +{} more", dep_refs.join(", "), dependents.len() - 6)
                } else {
                    dep_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_b_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Feeds")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(dep_text)
                        )
                );
            }

            // Format section (if any formatting applied)
            if !format_badges.is_empty() {
                let badges: Vec<_> = format_badges.iter().map(|label| {
                    div()
                        .px(px(8.0))
                        .py(px(3.0))
                        .bg(divider)
                        .rounded(px(4.0))
                        .text_size(px(11.0))
                        .text_color(text_primary)
                        .child(*label)
                }).collect();

                content = content.child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mb(px(6.0))
                                .child("Format")
                        )
                        .child(
                            div()
                                .flex()
                                .gap(px(6.0))
                                .children(badges)
                        )
                );
            }

            content
        } else {
            // Simple value cell: value-first compact card
            let type_label = if raw_value.is_empty() {
                "Empty cell"
            } else if raw_value.parse::<f64>().is_ok() {
                "Number"
            } else if raw_value == "TRUE" || raw_value == "FALSE" {
                "Boolean"
            } else {
                "Text"
            };

            // Softer divider for value cells too
            let divider = panel_border.opacity(0.5);

            let mut content = div()
                .flex()
                .flex_col()
                // Header: cell ref
                .child(
                    div()
                        .px_3()
                        .py(px(8.0))
                        .border_b_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(cell_ref)
                        )
                )
                // VALUE - the hero
                .child(
                    div()
                        .px_3()
                        .py(px(12.0))
                        .child(
                            div()
                                .text_size(px(18.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child(if display_value.is_empty() { "(empty)".to_string() } else { display_value })
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .mt(px(2.0))
                                .child(type_label)
                        )
                );

            // Feeds (dependents) - only if this value is used elsewhere
            if !dependents.is_empty() {
                let dep_refs: Vec<String> = dependents.iter().take(6).map(|(r, c)| {
                    app.cell_ref_at(*r, *c)
                }).collect();
                let dep_text = if dependents.len() > 6 {
                    format!("{} +{} more", dep_refs.join(", "), dependents.len() - 6)
                } else {
                    dep_refs.join(", ")
                };

                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(divider)
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Feeds")
                        )
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(accent)
                                .child(dep_text)
                        )
                );
            } else if raw_value.is_empty() {
                // Empty cell with no dependents - positive framing
                content = content.child(
                    div()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(divider)
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_disabled)
                                .child("Independent cell")
                        )
                );
            }

            content
        }
    };

    // Position overlay near the selection
    // Calculate position based on selection and scroll
    let ((_min_row, _min_col), (max_row, max_col)) = app.selection_range();

    // Calculate pixel position of selection end (bottom-right of selection)
    // Account for: header width, scroll position, cell dimensions
    let header_w = app.metrics.header_w;
    let header_h = app.metrics.header_h;
    let cell_w = app.metrics.cell_w;
    let cell_h = app.metrics.cell_h;

    // Menu bar + formula bar height (approximate)
    let top_offset = 24.0 + 32.0 + header_h; // menu + formula bar + column headers

    // X position: right edge of selection, offset from scroll
    let col_offset = (max_col as f32 - app.view_state.scroll_col as f32 + 1.0) * cell_w;
    let overlay_x = header_w + col_offset + 8.0; // 8px gap from selection

    // Y position: below the selection
    let row_offset = (max_row as f32 - app.view_state.scroll_row as f32 + 1.0) * cell_h;
    let overlay_y = top_offset + row_offset + 4.0; // 4px gap below selection

    // Clamp to reasonable bounds (don't go off screen)
    let overlay_x = overlay_x.max(header_w + 20.0);
    let overlay_y = overlay_y.max(top_offset + 20.0);

    div()
        .absolute()
        .inset_0()
        .child(
            div()
                .absolute()
                .left(px(overlay_x))
                .top(px(overlay_y))
                .w(px(240.0))
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .child(content)
        )
}
