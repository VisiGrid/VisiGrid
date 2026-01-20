use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::mode::InspectorTab;
use crate::theme::TokenKey;
use visigrid_engine::formula::parser::{parse, extract_cell_refs};

const PANEL_WIDTH: f32 = 280.0;

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
}

fn render_content(
    app: &Spreadsheet,
    row: usize,
    col: usize,
    current_tab: InspectorTab,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    div()
        .flex_1()
        .overflow_hidden()
        .child(match current_tab {
            InspectorTab::Inspector => render_inspector_tab(app, row, col, text_primary, text_muted, accent, panel_border, cx).into_any_element(),
            InspectorTab::Format => render_format_tab(app, row, col, text_primary, text_muted, panel_border).into_any_element(),
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
    row: usize,
    col: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
) -> impl IntoElement {
    let format = app.sheet().get_format(row, col);

    div()
        .p_3()
        .flex()
        .flex_col()
        .gap_4()
        // Text formatting
        .child(
            section("Text Style", panel_border, text_primary)
                .child(info_row("Bold", if format.bold { "Yes" } else { "No" }, text_muted, text_primary))
                .child(info_row("Italic", if format.italic { "Yes" } else { "No" }, text_muted, text_primary))
                .child(info_row("Underline", if format.underline { "Yes" } else { "No" }, text_muted, text_primary))
        )
        // Font
        .child(
            section("Font", panel_border, text_primary)
                .child(info_row(
                    "Family",
                    format.font_family.as_deref().unwrap_or("(Default)"),
                    text_muted,
                    text_primary,
                ))
        )
        // Number format (future)
        .child(
            section("Number Format", panel_border, text_primary)
                .child(info_row("Format", "(Default)", text_muted, text_primary))
        )
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
