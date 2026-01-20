use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, SelectionFormatState, TriState};
use crate::mode::InspectorTab;
use crate::theme::TokenKey;
use visigrid_engine::formula::parser::{parse, extract_cell_refs};
use visigrid_engine::cell::{Alignment, VerticalAlignment, TextOverflow, NumberFormat, DateStyle};

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
    let (row, col) = app.inspector_pinned.unwrap_or(app.selected);
    let cell_ref = app.cell_ref_at(row, col);
    let is_pinned = app.inspector_pinned.is_some();
    let current_tab = app.inspector_tab;

    div()
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
    // Pre-build the names tab content while we have &mut access for usage cache
    let names_content = if current_tab == InspectorTab::Names {
        Some(render_names_tab(app, text_primary, text_muted, accent, panel_border, cx))
    } else {
        None
    };

    div()
        .flex_1()
        .overflow_hidden()
        .child(match current_tab {
            InspectorTab::Inspector => render_inspector_tab(app, row, col, text_primary, text_muted, accent, panel_border, cx).into_any_element(),
            InspectorTab::Format => render_format_tab(app, row, col, text_primary, text_muted, panel_border, accent, cx).into_any_element(),
            InspectorTab::Names => names_content.unwrap().into_any_element(),
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
) -> impl IntoElement {
    let raw_value = app.sheet().get_raw(row, col);
    let display_value = app.sheet().get_display(row, col);
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
    let dependents = get_dependents(app, row, col);

    // Check spill info
    let is_spill_parent = app.sheet().is_spill_parent(row, col);
    let is_spill_receiver = app.sheet().is_spill_receiver(row, col);
    let spill_parent = app.sheet().get_spill_parent(row, col);

    let cell_address = app.cell_ref_at(row, col);
    let has_spill_info = is_spill_parent || is_spill_receiver;
    let has_no_deps = precedents.is_empty() && dependents.is_empty() && !is_formula && !has_spill_info;

    let mut content = div()
        .p_3()
        .flex()
        .flex_col()
        .gap_4()
        // Identity section
        .child(
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

    // Precedents section
    if !precedents.is_empty() {
        let mut prec_section = section("Precedents (depends on)", panel_border, text_primary);
        for (r, c) in precedents.iter().take(10) {
            prec_section = prec_section.child(clickable_cell_row(
                "",
                &app.cell_ref_at(*r, *c),
                *r,
                *c,
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

    // Dependents section
    if !dependents.is_empty() {
        let mut dep_section = section("Dependents (used by)", panel_border, text_primary);
        for (r, c) in dependents.iter().take(10) {
            dep_section = dep_section.child(clickable_cell_row(
                "",
                &app.cell_ref_at(*r, *c),
                *r,
                *c,
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

    content
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
    let state = app.selection_format_state();

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
    let ranges: Vec<_> = app.filtered_named_ranges()
        .into_iter()
        .cloned()
        .collect();

    // Then get usage counts (mutable borrow for cache)
    let named_ranges: Vec<_> = ranges.into_iter()
        .map(|nr| {
            let usage_count = app.get_named_range_usage_count(&nr.name);
            (nr, usage_count)
        })
        .collect();

    let filter_query = app.names_filter_query.clone();
    let all_names = app.workbook.list_named_ranges();
    let has_names = !all_names.is_empty();
    let has_filtered_results = !named_ranges.is_empty();
    let name_count = all_names.len();

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
                    .hover(|s| s.bg(panel_border.opacity(0.3)))
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        this.jump_to_named_range(&name_for_jump, cx);
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
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.show_create_named_range(cx);
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
        // Keyboard hint
        .child(
            div()
                .pt_2()
                .border_t_1()
                .border_color(panel_border)
                .text_size(px(9.0))
                .text_color(text_muted.opacity(0.7))
                .child("Ctrl+Shift+N: Create | Ctrl+Shift+R: Rename")
        )
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
                            let state = this.selection_format_state();
                            let new_value = !matches!(state.bold, TriState::Uniform(true));
                            this.set_bold(new_value, cx);
                        }))
                )
                // Italic button
                .child(
                    format_toggle_btn("I", italic_active, italic_mixed, false, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            let state = this.selection_format_state();
                            let new_value = !matches!(state.italic, TriState::Uniform(true));
                            this.set_italic(new_value, cx);
                        }))
                )
                // Underline button
                .child(
                    format_toggle_btn("U", underline_active, underline_mixed, false, text_primary, text_muted, accent, panel_border)
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            let state = this.selection_format_state();
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
                            let state = this.selection_format_state();
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

// Helper: Clickable cell reference
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

// Get precedents from a formula string
fn get_precedents(formula: &str) -> Vec<(usize, usize)> {
    if let Ok(expr) = parse(formula) {
        let mut refs = extract_cell_refs(&expr);
        refs.sort();
        refs.dedup();
        refs
    } else {
        Vec::new()
    }
}

// Get dependents (cells that reference the given cell)
fn get_dependents(app: &Spreadsheet, row: usize, col: usize) -> Vec<(usize, usize)> {
    let mut dependents = Vec::new();

    // Iterate through all cells with formulas and check if they reference this cell
    for ((cell_row, cell_col), cell) in app.sheet().cells_iter() {
        let raw = cell.value.raw_display();
        if raw.starts_with('=') {
            if let Ok(expr) = parse(&raw) {
                let refs = extract_cell_refs(&expr);
                if refs.contains(&(row, col)) {
                    dependents.push((*cell_row, *cell_col));
                }
            }
        }
    }

    dependents.sort();
    dependents
}
