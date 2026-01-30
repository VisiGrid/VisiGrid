use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, TriState, CELL_HEIGHT};
use crate::theme::TokenKey;
use visigrid_engine::cell::{Alignment, VerticalAlignment};

pub const FORMAT_BAR_HEIGHT: f32 = 28.0;

/// Common font sizes for the dropdown.
const FONT_SIZES: &[u32] = &[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];

/// Default font size (matches engine default).
const DEFAULT_FONT_SIZE: u32 = 11;

/// Render the format bar (between formula bar and column headers).
pub fn render_format_bar(app: &mut Spreadsheet, window: &Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let toolbar_bg = app.token(TokenKey::HeaderBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);

    // Auto-commit font size if focus moved away while editing
    if app.ui.format_bar.size_editing && !app.ui.format_bar.size_focus.is_focused(window) {
        commit_font_size(app, cx);
    }

    let state = app.selection_format_state(cx);

    // Font family display
    let font_display: SharedString = match &state.font_family {
        TriState::Uniform(Some(f)) => f.clone().into(),
        TriState::Uniform(None) | TriState::Empty => "(Default)".into(),
        TriState::Mixed => "\u{2014}".into(), // em dash
    };
    let font_is_mixed = state.font_family.is_mixed();

    // Font size display
    let size_replace_next = app.ui.format_bar.size_replace_next;
    let size_display: SharedString = if app.ui.format_bar.size_editing {
        if size_replace_next {
            // Show value without caret — it's "selected all" (will be replaced on first digit)
            app.ui.format_bar.size_input.clone().into()
        } else {
            // Show value with caret
            format!("{}|", app.ui.format_bar.size_input).into()
        }
    } else {
        match &state.font_size {
            TriState::Uniform(Some(s)) => format!("{}", *s as u32).into(),
            TriState::Uniform(None) | TriState::Empty => format!("{}", DEFAULT_FONT_SIZE).into(),
            TriState::Mixed => "\u{2014}".into(),
        }
    };
    let size_is_mixed = state.font_size.is_mixed();
    let size_editing = app.ui.format_bar.size_editing;
    let size_dropdown = app.ui.format_bar.size_dropdown;

    // Bold / Italic / Underline tri-state
    let bold_active = matches!(state.bold, TriState::Uniform(true));
    let bold_mixed = state.bold.is_mixed();
    let italic_active = matches!(state.italic, TriState::Uniform(true));
    let italic_mixed = state.italic.is_mixed();
    let underline_active = matches!(state.underline, TriState::Uniform(true));
    let underline_mixed = state.underline.is_mixed();

    // Fill color chip
    let fill_chip_color = rgba_to_hsla(&state.background_color);

    // Text color underbar
    let text_color_hsla = rgba_to_hsla(&state.font_color);
    let text_color_is_mixed = state.font_color.is_mixed();

    // Alignment state
    let align_left = matches!(state.alignment, TriState::Uniform(Alignment::Left));
    let align_center = matches!(state.alignment, TriState::Uniform(Alignment::Center));
    let align_right = matches!(state.alignment, TriState::Uniform(Alignment::Right));

    let size_focus = app.ui.format_bar.size_focus.clone();

    div()
        .flex()
        .flex_shrink_0()
        .relative()
        .h(px(FORMAT_BAR_HEIGHT))
        .w_full()
        .bg(toolbar_bg)
        .border_b_1()
        .border_color(panel_border)
        .items_center()
        .px_2()
        .gap_1()
        // Font family button
        .child(render_font_family_btn(
            font_display, font_is_mixed, text_primary, text_muted, panel_border, cx,
        ))
        // Font size input
        .child(render_font_size_input(
            app, size_display, size_is_mixed, size_editing, size_replace_next, text_primary, text_muted, panel_border, size_focus, cx,
        ))
        // Separator
        .child(toolbar_separator(panel_border))
        // B / I / U toggle buttons
        .child(render_style_btn("B", bold_active, bold_mixed, text_primary, text_muted, accent, panel_border, cx))
        .child(render_style_btn("I", italic_active, italic_mixed, text_primary, text_muted, accent, panel_border, cx))
        .child(render_style_btn("U", underline_active, underline_mixed, text_primary, text_muted, accent, panel_border, cx))
        // Separator
        .child(toolbar_separator(panel_border))
        // Fill color button
        .child(render_fill_color_btn(fill_chip_color, panel_border, cx))
        // Text color button
        .child(render_text_color_btn(text_color_hsla, text_color_is_mixed, text_primary, text_muted, panel_border, cx))
        // Separator
        .child(toolbar_separator(panel_border))
        // Alignment buttons
        .child(render_align_btn(Alignment::Left, align_left, text_primary, text_muted, accent, panel_border, cx))
        .child(render_align_btn(Alignment::Center, align_center, text_primary, text_muted, accent, panel_border, cx))
        .child(render_align_btn(Alignment::Right, align_right, text_primary, text_muted, accent, panel_border, cx))
}

// ============================================================================
// Helper: RGBA tri-state to Hsla
// ============================================================================

/// Convert TriState<Option<[u8; 4]>> to Option<Hsla> for display.
/// Uniform(Some(c)) → Some(color), Uniform(None)/Empty → Some(white/default), Mixed → None.
fn rgba_to_hsla(tri: &TriState<Option<[u8; 4]>>) -> Option<Hsla> {
    match tri {
        TriState::Uniform(Some(c)) => {
            let [r, g, b, _] = *c;
            Some(Hsla::from(gpui::Rgba {
                r: r as f32 / 255.0,
                g: g as f32 / 255.0,
                b: b as f32 / 255.0,
                a: 1.0,
            }))
        }
        TriState::Uniform(None) | TriState::Empty => Some(hsla(0.0, 0.0, 1.0, 1.0)),
        TriState::Mixed => None,
    }
}

// ============================================================================
// Controls
// ============================================================================

/// Thin vertical separator between toolbar groups.
fn toolbar_separator(border: Hsla) -> impl IntoElement {
    div()
        .w(px(1.0))
        .h(px(16.0))
        .mx(px(2.0))
        .bg(border.opacity(0.5))
}

/// Font family clickable label — opens font picker.
fn render_font_family_btn(
    display: SharedString,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let text_color = if is_mixed { text_muted } else { text_primary };

    div()
        .id("fmt-font-family")
        .min_w(px(80.0))
        .max_w(px(120.0))
        .h(px(20.0))
        .px_2()
        .flex()
        .items_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(panel_border)
        .text_size(px(11.0))
        .text_color(text_color)
        .overflow_hidden()
        .when(is_mixed, |d| d.italic())
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            this.show_font_picker(window, cx);
        }))
        .tooltip(|_window, cx| {
            cx.new(|_| FormatBarTooltip("Font (Format \u{2192} Font...)")).into()
        })
        .child(display)
}

/// Font size editable input with dropdown arrow.
fn render_font_size_input(
    app: &Spreadsheet,
    display: SharedString,
    is_mixed: bool,
    is_editing: bool,
    is_selected_all: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    focus_handle: FocusHandle,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let text_color = if is_editing && is_selected_all { text_primary } else if is_mixed { text_muted } else { text_primary };
    let editor_bg = app.token(TokenKey::EditorBg);
    let accent = app.token(TokenKey::Accent);

    div()
        .flex()
        .items_center()
        .child(
            div()
                .id("fmt-font-size")
                .track_focus(&focus_handle)
                .w(px(38.0))
                .h(px(20.0))
                .px(px(4.0))
                .flex()
                .items_center()
                .justify_center()
                .rounded_l_sm()
                .border_1()
                .border_color(if is_editing { accent } else { panel_border })
                .bg(if is_editing { editor_bg } else { gpui::transparent_black() })
                .text_size(px(11.0))
                .text_color(text_color)
                .when(!is_editing, |d| d.cursor_pointer())
                .when(is_editing, |d| d.cursor_text())
                .when(is_mixed && !is_editing, |d| d.italic())
                .hover(|s| s.bg(panel_border.opacity(0.3)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                    cx.stop_propagation();
                    if !this.ui.format_bar.size_editing {
                        // Enter editing mode: populate buffer with current display value
                        let state = this.selection_format_state(cx);
                        this.ui.format_bar.size_input = match &state.font_size {
                            TriState::Uniform(Some(s)) => format!("{}", *s as u32),
                            TriState::Uniform(None) | TriState::Empty => format!("{}", DEFAULT_FONT_SIZE),
                            TriState::Mixed => String::new(),
                        };
                        this.ui.format_bar.size_editing = true;
                        this.ui.format_bar.size_dropdown = false;
                        this.ui.format_bar.size_replace_next = true;
                        window.focus(&this.ui.format_bar.size_focus, cx);
                        cx.notify();
                    }
                }))
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    if !this.ui.format_bar.size_editing {
                        return;
                    }
                    cx.stop_propagation();
                    match event.keystroke.key.as_str() {
                        "enter" => {
                            commit_font_size(this, cx);
                        }
                        "escape" => {
                            this.ui.format_bar.size_editing = false;
                            this.ui.format_bar.size_dropdown = false;
                            this.ui.format_bar.size_replace_next = false;
                            cx.notify();
                        }
                        "backspace" => {
                            this.ui.format_bar.size_replace_next = false;
                            this.ui.format_bar.size_input.pop();
                            cx.notify();
                        }
                        _ => {
                            if let Some(ch) = &event.keystroke.key_char {
                                // Only allow digits
                                if ch.chars().all(|c| c.is_ascii_digit()) {
                                    // First keypress after entering edit: replace entire value
                                    if this.ui.format_bar.size_replace_next {
                                        this.ui.format_bar.size_input.clear();
                                        this.ui.format_bar.size_replace_next = false;
                                    }
                                    this.ui.format_bar.size_input.push_str(ch);
                                    cx.notify();
                                }
                            }
                        }
                    }
                }))
                .tooltip(|_window, cx: &mut App| {
                    cx.new(|_| FormatBarTooltip("Font Size")).into()
                })
                .when(is_editing && is_selected_all, |d| {
                    // "Select all" visual: accent background on the text to show it will be replaced
                    d.child(
                        div()
                            .bg(accent.opacity(0.3))
                            .rounded(px(2.0))
                            .px(px(1.0))
                            .child(display.clone())
                    )
                })
                .when(!(is_editing && is_selected_all), |d| {
                    d.child(display.clone())
                })
        )
        // Dropdown arrow button
        .child(
            div()
                .id("fmt-font-size-dropdown")
                .w(px(14.0))
                .h(px(20.0))
                .flex()
                .items_center()
                .justify_center()
                .rounded_r_sm()
                .cursor_pointer()
                .border_t_1()
                .border_b_1()
                .border_r_1()
                .border_color(panel_border)
                .text_size(px(8.0))
                .text_color(text_muted)
                .hover(|s| s.bg(panel_border.opacity(0.3)))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    cx.stop_propagation();
                    this.ui.format_bar.size_dropdown = !this.ui.format_bar.size_dropdown;
                    this.ui.format_bar.size_editing = false;
                    cx.notify();
                }))
                .child("\u{25BC}") // ▼
        )
}

/// Font size dropdown overlay — must be rendered at root level (after grid)
/// so it paints above column headers and cells.
pub fn render_font_size_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let hover_bg = app.token(TokenKey::ToolbarButtonHoverBg);

    // Vertical offset from root div top to below the format bar.
    // macOS: titlebar(34) + formula_bar(CELL_HEIGHT=24) + format_bar(28) = 86
    // Other: menu_bar(28) + formula_bar(24) + format_bar(28) = 80
    let chrome_above: f32 = if cfg!(target_os = "macos") { 34.0 } else { crate::app::MENU_BAR_HEIGHT };
    let top_offset = chrome_above + CELL_HEIGHT + FORMAT_BAR_HEIGHT;

    // Horizontal offset: px_2 padding (8) + font family (~84) + gap_1 (4) = ~96
    let left_offset: f32 = 8.0 + 84.0 + 4.0;

    let mut dropdown = div()
        .id("fmt-size-dropdown-panel")
        .absolute()
        .top(px(top_offset))
        .left(px(left_offset))
        .w(px(54.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_lg()
        .py_1()
        .flex()
        .flex_col()
        .on_mouse_down_out(cx.listener(|this, _, _, cx| {
            this.ui.format_bar.size_dropdown = false;
            this.ui.format_bar.size_editing = false;
            cx.notify();
        }));

    for &size in FONT_SIZES {
        let size_f32 = size as f32;
        dropdown = dropdown.child(
            div()
                .id(SharedString::from(format!("fmt-size-{}", size)))
                .px_2()
                .py(px(2.0))
                .text_size(px(11.0))
                .text_color(text_primary)
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation();
                    this.set_font_size_selection(Some(size_f32), cx);
                    this.ui.format_bar.size_dropdown = false;
                    this.ui.format_bar.size_editing = false;
                    cx.notify();
                }))
                .child(SharedString::from(format!("{}", size)))
        );
    }

    // "Auto" option to clear font size
    dropdown.child(
        div()
            .id("fmt-size-auto")
            .px_2()
            .py(px(2.0))
            .text_size(px(11.0))
            .text_color(text_muted)
            .italic()
            .cursor_pointer()
            .hover(move |s| s.bg(hover_bg))
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                cx.stop_propagation();
                this.set_font_size_selection(None, cx);
                this.ui.format_bar.size_dropdown = false;
                this.ui.format_bar.size_editing = false;
                cx.notify();
            }))
            .child("Auto")
    )
}

/// Parse font size input text into a validated font size.
/// Returns `Some(size)` for valid integers 1..=400, `None` otherwise.
/// This is the pure-logic core of commit_font_size, extracted for testability.
pub(crate) fn parse_font_size_input(input: &str) -> Option<f32> {
    let trimmed = input.trim();
    if let Ok(size) = trimmed.parse::<u32>() {
        if size >= 1 && size <= 400 {
            return Some(size as f32);
        }
    }
    None
}

/// Commit the font size input value and exit editing mode.
pub fn commit_font_size(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    if !app.ui.format_bar.size_editing {
        return; // Guard against double-commit
    }
    let input = app.ui.format_bar.size_input.clone();
    app.ui.format_bar.size_editing = false;
    app.ui.format_bar.size_dropdown = false;
    app.ui.format_bar.size_replace_next = false;

    if let Some(size) = parse_font_size_input(&input) {
        app.set_font_size_selection(Some(size), cx);
    }
    // If parse fails or out of range, just revert (no change)
    cx.notify();
}

/// B / I / U toggle button for the format bar.
fn render_style_btn(
    label: &'static str,
    is_active: bool,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let mut btn = div()
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .text_size(px(12.0));

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

    // Apply styling to the label itself
    if label == "B" {
        btn = btn.font_weight(FontWeight::BOLD);
    }
    if label == "I" {
        btn = btn.italic();
    }
    if label == "U" {
        btn = btn.underline();
    }

    let display_label: &str = if is_mixed { "\u{2014}" } else { label };

    // Tooltip text with shortcut
    #[cfg(not(target_os = "macos"))]
    let tooltip_text = match label {
        "B" => "Bold (Ctrl+B)",
        "I" => "Italic (Ctrl+I)",
        "U" => "Underline (Ctrl+U)",
        _ => label,
    };

    #[cfg(target_os = "macos")]
    let tooltip_text = match label {
        "B" => "Bold (\u{2318}B)",
        "I" => "Italic (\u{2318}I)",
        "U" => "Underline (\u{2318}U)",
        _ => label,
    };

    btn
        .child(display_label)
        .id(SharedString::from(format!("fmt-style-{}", label)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            match label {
                "B" => this.toggle_bold(cx),
                "I" => this.toggle_italic(cx),
                "U" => this.toggle_underline(cx),
                _ => {}
            }
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
}

/// Fill color swatch button — opens color picker.
fn render_fill_color_btn(
    chip_color: Option<Hsla>,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let chip = render_color_chip(chip_color, panel_border, 16.0, 7.0);

    div()
        .id("fmt-fill-color")
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(panel_border)
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            this.show_color_picker(crate::color_palette::ColorTarget::Fill, window, cx);
        }))
        .tooltip(|_window, cx| {
            cx.new(|_| FormatBarTooltip("Fill Color")).into()
        })
        .child(chip)
}

/// Text color button — "A" with colored underbar, opens color picker.
fn render_text_color_btn(
    color: Option<Hsla>,
    is_mixed: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Underbar color: use the actual font color, or text_primary for Automatic/None
    let underbar_color = color.unwrap_or(text_primary);

    div()
        .id("fmt-text-color")
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .flex_col()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(panel_border)
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            this.show_color_picker(crate::color_palette::ColorTarget::Text, window, cx);
        }))
        .tooltip(|_window, cx| {
            cx.new(|_| FormatBarTooltip("Text Color")).into()
        })
        // "A" letter
        .child(
            div()
                .text_size(px(11.0))
                .font_weight(FontWeight::BOLD)
                .text_color(if is_mixed { text_muted } else { text_primary })
                .child(if is_mixed { "\u{2014}" } else { "A" })
        )
        // Colored underbar
        .child(
            if is_mixed {
                // Checkerboard underbar for mixed
                let dark = panel_border.opacity(0.5);
                let light = hsla(0.0, 0.0, 1.0, 1.0);
                div()
                    .w(px(16.0))
                    .h(px(3.0))
                    .rounded_sm()
                    .flex()
                    .child(div().w(px(4.0)).h(px(3.0)).bg(dark))
                    .child(div().w(px(4.0)).h(px(3.0)).bg(light))
                    .child(div().w(px(4.0)).h(px(3.0)).bg(dark))
                    .child(div().w(px(4.0)).h(px(3.0)).bg(light))
            } else {
                div()
                    .w(px(16.0))
                    .h(px(3.0))
                    .rounded_sm()
                    .bg(underbar_color)
            }
        )
}

/// Color chip: solid swatch or 2x2 checkerboard for mixed state.
fn render_color_chip(color: Option<Hsla>, panel_border: Hsla, size: f32, half: f32) -> Div {
    if let Some(bg) = color {
        div()
            .size(px(size))
            .rounded_sm()
            .border_1()
            .border_color(panel_border)
            .bg(bg)
    } else {
        let dark = panel_border.opacity(0.3);
        let light = hsla(0.0, 0.0, 1.0, 1.0);
        div()
            .size(px(size))
            .rounded_sm()
            .border_1()
            .border_color(panel_border)
            .flex()
            .flex_wrap()
            .child(div().w(px(half)).h(px(half)).bg(dark))
            .child(div().w(px(half)).h(px(half)).bg(light))
            .child(div().w(px(half)).h(px(half)).bg(light))
            .child(div().w(px(half)).h(px(half)).bg(dark))
    }
}

/// Alignment toggle button with 3-line icon.
fn render_align_btn(
    target: Alignment,
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let line_color = if is_active { text_primary } else { text_muted };

    let mut btn = div()
        .w(px(22.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1();

    if is_active {
        btn = btn
            .bg(accent.opacity(0.2))
            .border_color(accent);
    } else {
        btn = btn
            .bg(gpui::transparent_black())
            .border_color(panel_border);
    }

    btn = btn.hover(|s| s.bg(panel_border.opacity(0.5)));

    let (id, tooltip_text): (&str, &str) = match target {
        Alignment::Left => ("fmt-align-left", "Align Left"),
        Alignment::Center => ("fmt-align-center", "Align Center"),
        Alignment::Right => ("fmt-align-right", "Align Right"),
        _ => ("fmt-align-general", "Align General"),
    };

    btn
        .child(render_align_icon(target, line_color))
        .id(SharedString::from(id))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.set_alignment_selection(target, cx);
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
}

/// Render a 3-line alignment icon (12px wide, 10px tall).
pub(super) fn render_align_icon(alignment: Alignment, color: Hsla) -> impl IntoElement {
    // Three horizontal lines of varying width to convey alignment
    let (w1, w2, w3) = (px(12.0), px(8.0), px(10.0));
    let line_h = px(1.5);
    let gap = px(1.5);

    let align_fn = match alignment {
        Alignment::Left => |d: Div| d.items_start(),
        Alignment::Center => |d: Div| d.items_center(),
        Alignment::Right => |d: Div| d.items_end(),
        _ => |d: Div| d.items_start(),
    };

    align_fn(
        div()
            .w(px(14.0))
            .flex()
            .flex_col()
            .gap(gap)
    )
    .child(div().w(w1).h(line_h).bg(color).rounded_sm())
    .child(div().w(w2).h(line_h).bg(color).rounded_sm())
    .child(div().w(w3).h(line_h).bg(color).rounded_sm())
}

/// Render a 3-line vertical alignment icon (14px wide, 14px tall).
///
/// Three equal-width lines positioned at top, middle, or bottom of a fixed-height box.
pub(super) fn render_valign_icon(alignment: VerticalAlignment, color: Hsla) -> impl IntoElement {
    let line_w = px(12.0);
    let line_h = px(1.5);
    let gap = px(1.5);

    let justify_fn = match alignment {
        VerticalAlignment::Top => |d: Div| d.justify_start(),
        VerticalAlignment::Middle => |d: Div| d.justify_center(),
        VerticalAlignment::Bottom => |d: Div| d.justify_end(),
    };

    justify_fn(
        div()
            .w(px(14.0))
            .h(px(14.0))
            .flex()
            .flex_col()
            .items_center()
            .gap(gap)
    )
    .child(div().w(line_w).h(line_h).bg(color).rounded_sm())
    .child(div().w(line_w).h(line_h).bg(color).rounded_sm())
    .child(div().w(line_w).h(line_h).bg(color).rounded_sm())
}

// ============================================================================
// Tooltip
// ============================================================================

/// Minimal tooltip for format bar buttons.
struct FormatBarTooltip(&'static str);

impl Render for FormatBarTooltip {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .px_2()
            .py_1()
            .rounded_sm()
            .bg(rgb(0x2d2d2d))
            .border_1()
            .border_color(rgb(0x3d3d3d))
            .text_size(px(11.0))
            .text_color(rgb(0xcccccc))
            .child(self.0)
    }
}

// Tests for format bar state machine live in tests.rs (this file imports gpui::*
// which shadows the standard #[test] attribute with gpui::test).
