use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, TriState, CELL_HEIGHT};
use crate::mode::Mode;
use crate::theme::TokenKey;
use visigrid_engine::cell::{Alignment, CellStyle, VerticalAlignment, NumberFormat};

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

    // Bold / Italic / Underline tri-state (mixed state not used by Format dropdown)
    let bold_active = matches!(state.bold, TriState::Uniform(true));
    let italic_active = matches!(state.italic, TriState::Uniform(true));
    let underline_active = matches!(state.underline, TriState::Uniform(true));

    // Fill color chip
    let fill_chip_color = rgba_to_hsla(&state.background_color);

    // Text color underbar
    let text_color_hsla = rgba_to_hsla(&state.font_color);
    let text_color_is_mixed = state.font_color.is_mixed();

    // Alignment state
    let align_left = matches!(state.alignment, TriState::Uniform(Alignment::Left));
    let align_center = matches!(state.alignment, TriState::Uniform(Alignment::Center));
    let align_right = matches!(state.alignment, TriState::Uniform(Alignment::Right));

    // Sort/filter state for toolbar buttons
    let filter_active = app.filter_state.is_enabled();
    let selected_col = app.view_state.selected.1;
    let sort_state = app.display_sort_state(cx);
    let sort_asc_active = sort_state.map_or(false, |(col, asc)| asc && col == selected_col);
    let sort_desc_active = sort_state.map_or(false, |(col, asc)| !asc && col == selected_col);
    let number_format_menu_open = app.ui.format_bar.number_format_menu_open;
    let cell_style_menu_open = app.ui.format_bar.cell_style_menu_open;

    let size_focus = app.ui.format_bar.size_focus.clone();

    div()
        .flex()
        .flex_shrink_0()
        .relative()
        .h(px(FORMAT_BAR_HEIGHT))
        .w_full()
        .bg(toolbar_bg)
        // No bottom border - let spacing and background separation do the work
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
        // Format dropdown (replaces B/I/U buttons)
        .child(render_format_dropdown_btn(
            bold_active, italic_active, underline_active,
            align_left, align_center, align_right,
            text_primary, text_muted, accent, panel_border,
            app.ui.format_menu_open,
            cx,
        ))
        // Format Painter button
        .child(render_format_painter_btn(app.mode == Mode::FormatPainter, text_primary, text_muted, accent, panel_border, cx))
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
        // Separator — data tools
        .child(toolbar_separator(panel_border))
        // Sort Ascending
        .child(render_sort_btn(true, sort_asc_active, text_primary, text_muted, accent, panel_border, cx))
        // Sort Descending
        .child(render_sort_btn(false, sort_desc_active, text_primary, text_muted, accent, panel_border, cx))
        // Filter toggle
        .child(render_filter_btn(filter_active, text_primary, text_muted, accent, panel_border, cx))
        // Separator — functions/format
        .child(toolbar_separator(panel_border))
        // AutoSum
        .child(render_autosum_btn(text_primary, text_muted, panel_border, cx))
        // Number Format dropdown
        .child(render_number_format_btn(text_primary, text_muted, accent, panel_border, number_format_menu_open, cx))
        // Separator
        .child(toolbar_separator(panel_border))
        // Cell Styles dropdown
        .child(render_cell_style_btn(text_primary, text_muted, accent, panel_border, cell_style_menu_open, cx))
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

/// Compact "Format ▾" dropdown button (replaces separate B/I/U buttons).
fn render_format_dropdown_btn(
    bold_active: bool,
    italic_active: bool,
    underline_active: bool,
    _align_left: bool,
    _align_center: bool,
    _align_right: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    is_open: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Show indicator if any formatting is active
    let has_active = bold_active || italic_active || underline_active;
    let btn_bg = if is_open { accent.opacity(0.15) } else { gpui::transparent_black() };
    let btn_border = if is_open { accent.opacity(0.5) } else { panel_border };
    let btn_color = if is_open { accent } else if has_active { text_primary } else { text_muted };

    div()
        .id("fmt-format-dropdown")
        .flex()
        .items_center()
        .gap(px(3.0))
        .px(px(6.0))
        .h(px(22.0))
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .text_size(px(11.0))
        .text_color(btn_color)
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.toggle_format_menu(cx);
        }))
        .child("Format")
        .child(div().text_size(px(7.0)).child("▾"))
}

/// Format Painter toggle button (paintbrush icon).
/// Single click → single-shot. Shift+click → locked mode.
fn render_format_painter_btn(
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let btn_bg = if is_active { accent.opacity(0.2) } else { gpui::transparent_black() };
    let btn_border = if is_active { accent } else { panel_border };
    let btn_color = if is_active { text_primary } else { text_muted };

    #[cfg(target_os = "macos")]
    let tooltip_text = "Format Painter (\u{2318}\u{21e7}C to copy)";
    #[cfg(not(target_os = "macos"))]
    let tooltip_text = "Format Painter (Ctrl+Shift+C to copy)";

    div()
        .id("fmt-painter")
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .text_size(px(12.0))
        .text_color(btn_color)
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, event: &MouseDownEvent, _, cx| {
            if this.mode == Mode::FormatPainter {
                // Already active: cancel
                this.cancel_format_painter(cx);
            } else if event.modifiers.shift {
                // Shift+click: locked mode
                this.start_format_painter_locked(cx);
            } else {
                // Normal click: single-shot
                this.start_format_painter(cx);
            }
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
        // Paintbrush icon: 🖌 (U+1F58C)
        .child("\u{1f58c}\u{fe0e}")
}

/// Format dropdown panel - rendered at root level so it floats above grid.
pub fn render_format_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let hover_bg = app.token(TokenKey::ToolbarButtonHoverBg);

    let state = app.selection_format_state(cx);
    let bold_active = matches!(state.bold, TriState::Uniform(true));
    let italic_active = matches!(state.italic, TriState::Uniform(true));
    let underline_active = matches!(state.underline, TriState::Uniform(true));
    let align_left = matches!(state.alignment, TriState::Uniform(Alignment::Left));
    let align_center = matches!(state.alignment, TriState::Uniform(Alignment::Center));
    let align_right = matches!(state.alignment, TriState::Uniform(Alignment::Right));

    // Position below format bar, roughly under Format button
    // macOS: titlebar(34) + formula_bar(24) + format_bar(28) = 86
    let chrome_above: f32 = if cfg!(target_os = "macos") { 34.0 } else { crate::app::MENU_BAR_HEIGHT };
    let top_offset = chrome_above + CELL_HEIGHT + FORMAT_BAR_HEIGHT;
    // Horizontal: after font family + size + separator
    let left_offset: f32 = 8.0 + 84.0 + 54.0 + 4.0 + 4.0;

    div()
        .id("fmt-format-dropdown-panel")
        .absolute()
        .top(px(top_offset))
        .left(px(left_offset))
        .w(px(140.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_lg()
        .py_1()
        .flex()
        .flex_col()
        .on_mouse_down_out(cx.listener(|this, _, _, cx| {
            this.close_format_menu(cx);
        }))
        // Bold
        .child(
            div()
                .id("fmt-menu-bold")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.toggle_bold(cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if bold_active { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if bold_active { accent } else { text_primary }).child("Bold"))
                .child(div().text_size(px(10.0)).text_color(text_muted).child("⌘B"))
        )
        // Italic
        .child(
            div()
                .id("fmt-menu-italic")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.toggle_italic(cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if italic_active { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if italic_active { accent } else { text_primary }).child("Italic"))
                .child(div().text_size(px(10.0)).text_color(text_muted).child("⌘I"))
        )
        // Underline
        .child(
            div()
                .id("fmt-menu-underline")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.toggle_underline(cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if underline_active { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if underline_active { accent } else { text_primary }).child("Underline"))
                .child(div().text_size(px(10.0)).text_color(text_muted).child("⌘U"))
        )
        // Separator
        .child(div().h(px(1.0)).mx_2().my_1().bg(panel_border.opacity(0.5)))
        // Align Left
        .child(
            div()
                .id("fmt-menu-align-left")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.set_alignment_selection(Alignment::Left, cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if align_left { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if align_left { accent } else { text_primary }).child("Align Left"))
        )
        // Align Center
        .child(
            div()
                .id("fmt-menu-align-center")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.set_alignment_selection(Alignment::Center, cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if align_center { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if align_center { accent } else { text_primary }).child("Align Center"))
        )
        // Align Right
        .child(
            div()
                .id("fmt-menu-align-right")
                .flex()
                .items_center()
                .px_2()
                .py(px(4.0))
                .cursor_pointer()
                .hover(move |s| s.bg(hover_bg))
                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                    this.set_alignment_selection(Alignment::Right, cx);
                }))
                .child(div().w(px(16.0)).text_size(px(10.0)).text_color(accent).child(if align_right { "✓" } else { "" }))
                .child(div().flex_1().text_size(px(12.0)).text_color(if align_right { accent } else { text_primary }).child("Align Right"))
        )
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

// ============================================================================
// Sort / Filter / AutoSum / Number Format buttons
// ============================================================================

/// Sort button — "A→Z" (ascending) or "Z→A" (descending).
fn render_sort_btn(
    ascending: bool,
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let label = if ascending { "A\u{2192}Z" } else { "Z\u{2192}A" };
    let id: &str = if ascending { "fmt-sort-asc" } else { "fmt-sort-desc" };

    let btn_bg = if is_active { accent.opacity(0.2) } else { gpui::transparent_black() };
    let btn_border = if is_active { accent } else { panel_border };
    let btn_color = if is_active { text_primary } else { text_muted };

    #[cfg(not(target_os = "macos"))]
    let tooltip_text: &str = if ascending { "Sort rows (A\u{2192}Z)" } else { "Sort rows (Z\u{2192}A)" };
    #[cfg(target_os = "macos")]
    let tooltip_text: &str = if ascending { "Sort rows (A\u{2192}Z)" } else { "Sort rows (Z\u{2192}A)" };

    div()
        .id(SharedString::from(id))
        .w(px(34.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .text_size(px(10.0))
        .text_color(btn_color)
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
            use visigrid_engine::filter::SortDirection;
            let dir = if ascending { SortDirection::Ascending } else { SortDirection::Descending };
            this.sort_by_current_column(dir, cx);
            this.update_title_if_needed(window, cx);
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
        .child(label)
}

/// Filter toggle button — funnel icon (three bars of decreasing width).
fn render_filter_btn(
    is_active: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let btn_bg = if is_active { accent.opacity(0.2) } else { gpui::transparent_black() };
    let btn_border = if is_active { accent } else { panel_border };
    let line_color = if is_active { text_primary } else { text_muted };

    #[cfg(not(target_os = "macos"))]
    let tooltip_text = "Filter rows (Ctrl+Shift+F)";
    #[cfg(target_os = "macos")]
    let tooltip_text = "Filter rows (\u{2318}\u{21e7}F)";

    div()
        .id("fmt-filter")
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.toggle_auto_filter(cx);
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
        .child(render_funnel_icon(line_color))
}

/// Funnel icon: three centered horizontal bars of decreasing width.
fn render_funnel_icon(color: Hsla) -> impl IntoElement {
    let line_h = px(1.5);
    let gap = px(1.5);

    div()
        .w(px(14.0))
        .flex()
        .flex_col()
        .items_center()
        .gap(gap)
        .child(div().w(px(12.0)).h(line_h).bg(color).rounded_sm())
        .child(div().w(px(8.0)).h(line_h).bg(color).rounded_sm())
        .child(div().w(px(4.0)).h(line_h).bg(color).rounded_sm())
}

/// AutoSum button — "Σ" (sigma).
fn render_autosum_btn(
    _text_primary: Hsla,
    text_muted: Hsla,
    panel_border: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    #[cfg(not(target_os = "macos"))]
    let tooltip_text = "Auto Sum (Alt+=)";
    #[cfg(target_os = "macos")]
    let tooltip_text = "Auto Sum (\u{2325}=)";

    div()
        .id("fmt-autosum")
        .w(px(26.0))
        .h(px(22.0))
        .flex()
        .items_center()
        .justify_center()
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(panel_border)
        .bg(gpui::transparent_black())
        .text_size(px(13.0))
        .text_color(text_muted)
        .hover(|s| s.bg(panel_border.opacity(0.5)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.autosum(cx);
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
        .child("\u{03A3}")
}

/// Number Format dropdown button — "123 ▾".
fn render_number_format_btn(
    _text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    is_open: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let btn_bg = if is_open { accent.opacity(0.15) } else { gpui::transparent_black() };
    let btn_border = if is_open { accent.opacity(0.5) } else { panel_border };
    let btn_color = if is_open { accent } else { text_muted };

    #[cfg(not(target_os = "macos"))]
    let tooltip_text = "Number format (Ctrl+1)";
    #[cfg(target_os = "macos")]
    let tooltip_text = "Number format (\u{2318}1)";

    div()
        .id("fmt-number-format")
        .flex()
        .items_center()
        .gap(px(2.0))
        .px(px(4.0))
        .h(px(22.0))
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .text_size(px(11.0))
        .text_color(btn_color)
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.ui.format_bar.number_format_menu_open = !this.ui.format_bar.number_format_menu_open;
            this.ui.format_bar.cell_style_menu_open = false;
            cx.notify();
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip(tooltip_text)).into()
        })
        .child("123")
        .child(div().text_size(px(7.0)).child("\u{25be}"))
}

/// Number format quick-menu dropdown overlay — rendered at root level.
pub fn render_number_format_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let hover_bg = app.token(TokenKey::ToolbarButtonHoverBg);

    let chrome_above: f32 = if cfg!(target_os = "macos") { 34.0 } else { crate::app::MENU_BAR_HEIGHT };
    let top_offset = chrome_above + CELL_HEIGHT + FORMAT_BAR_HEIGHT;

    // Horizontal offset: sum of all preceding button widths + gaps + padding.
    // px_2 padding(8) + font family(84+4) + font size(52+4) + sep(5) + format dropdown(~60+4)
    // + painter(26+4) + sep(5) + fill(26+4) + text(26+4) + sep(5) + align×3(22×3+4×2+4)
    // + sep(5) + sort×2(34×2+4+4) + filter(26+4) + sep(5) + autosum(26+4)
    // ≈ 8 + 88 + 56 + 5 + 64 + 30 + 5 + 30 + 30 + 5 + 74 + 5 + 76 + 30 + 5 + 30 = 541
    let left_offset: f32 = 541.0;

    div()
        .id("fmt-number-format-dropdown")
        .absolute()
        .top(px(top_offset))
        .left(px(left_offset))
        .w(px(120.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_lg()
        .py_1()
        .flex()
        .flex_col()
        .on_mouse_down_out(cx.listener(|this, _, _, cx| {
            this.ui.format_bar.number_format_menu_open = false;
            cx.notify();
        }))
        // General
        .child(render_numfmt_item("General", "fmt-nf-general", text_primary, hover_bg, cx, |this, cx| {
            this.set_number_format_selection(NumberFormat::General, cx);
        }))
        // Number
        .child(render_numfmt_item("Number", "fmt-nf-number", text_primary, hover_bg, cx, |this, cx| {
            this.set_number_format_selection(NumberFormat::number(2), cx);
        }))
        // Currency
        .child(render_numfmt_item("Currency", "fmt-nf-currency", text_primary, hover_bg, cx, |this, cx| {
            this.format_currency(cx);
        }))
        // Percent
        .child(render_numfmt_item("Percent", "fmt-nf-percent", text_primary, hover_bg, cx, |this, cx| {
            this.format_percent(cx);
        }))
        // Separator
        .child(div().h(px(1.0)).mx_2().my_1().bg(panel_border.opacity(0.5)))
        // More...
        .child(render_numfmt_item("More\u{2026}", "fmt-nf-more", text_muted, hover_bg, cx, |this, cx| {
            this.open_number_format_editor(cx);
        }))
}

/// Single item row in the number format dropdown.
fn render_numfmt_item(
    label: &'static str,
    id: &'static str,
    text_color: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    div()
        .id(SharedString::from(id))
        .flex()
        .items_center()
        .px_2()
        .py(px(4.0))
        .cursor_pointer()
        .text_size(px(12.0))
        .text_color(text_color)
        .hover(move |s| s.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            action(this, cx);
            this.ui.format_bar.number_format_menu_open = false;
            cx.notify();
        }))
        .child(label)
}

/// Cell Styles dropdown button — "Styles ▾".
fn render_cell_style_btn(
    _text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    panel_border: Hsla,
    is_open: bool,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let btn_bg = if is_open { accent.opacity(0.15) } else { gpui::transparent_black() };
    let btn_border = if is_open { accent.opacity(0.5) } else { panel_border };
    let btn_color = if is_open { accent } else { text_muted };

    div()
        .id("fmt-cell-styles")
        .flex()
        .items_center()
        .gap(px(2.0))
        .px(px(4.0))
        .h(px(22.0))
        .rounded_sm()
        .cursor_pointer()
        .border_1()
        .border_color(btn_border)
        .bg(btn_bg)
        .text_size(px(11.0))
        .text_color(btn_color)
        .hover(|s| s.bg(panel_border.opacity(0.3)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.ui.format_bar.cell_style_menu_open = !this.ui.format_bar.cell_style_menu_open;
            this.ui.format_bar.number_format_menu_open = false;
            cx.notify();
        }))
        .tooltip(move |_window, cx| {
            cx.new(|_| FormatBarTooltip("Cell Styles")).into()
        })
        .child("Styles")
        .child(div().text_size(px(7.0)).child("\u{25be}"))
}

/// Cell styles quick-menu dropdown overlay — rendered at root level.
pub fn render_cell_style_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let hover_bg = app.token(TokenKey::ToolbarButtonHoverBg);

    let chrome_above: f32 = if cfg!(target_os = "macos") { 34.0 } else { crate::app::MENU_BAR_HEIGHT };
    let top_offset = chrome_above + CELL_HEIGHT + FORMAT_BAR_HEIGHT;

    // Horizontal offset: number format btn left (541) + btn width (~40) + sep (5) ≈ 586
    let left_offset: f32 = 586.0;

    div()
        .id("fmt-cell-style-dropdown")
        .absolute()
        .top(px(top_offset))
        .left(px(left_offset))
        .w(px(140.0))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .rounded_sm()
        .shadow_lg()
        .py_1()
        .flex()
        .flex_col()
        .on_mouse_down_out(cx.listener(|this, _, _, cx| {
            this.ui.format_bar.cell_style_menu_open = false;
            cx.notify();
        }))
        // Normal (clear style)
        .child(render_cell_style_item("Normal", "cs-normal", None, None, text_muted, hover_bg, false, CellStyle::None, cx))
        // Separator
        .child(div().h(px(1.0)).mx_2().my_1().bg(panel_border.opacity(0.5)))
        // Error
        .child(render_cell_style_item("Error", "cs-error",
            Some(app.token(TokenKey::CellStyleErrorBg)),
            Some(app.token(TokenKey::CellStyleErrorBorder)),
            text_primary, hover_bg, false, CellStyle::Error, cx))
        // Warning
        .child(render_cell_style_item("Warning", "cs-warning",
            Some(app.token(TokenKey::CellStyleWarningBg)),
            Some(app.token(TokenKey::CellStyleWarningBorder)),
            text_primary, hover_bg, false, CellStyle::Warning, cx))
        // Success
        .child(render_cell_style_item("Success", "cs-success",
            Some(app.token(TokenKey::CellStyleSuccessBg)),
            Some(app.token(TokenKey::CellStyleSuccessBorder)),
            text_primary, hover_bg, false, CellStyle::Success, cx))
        // Input
        .child(render_cell_style_item("Input", "cs-input",
            Some(app.token(TokenKey::CellStyleInputBg)),
            Some(app.token(TokenKey::CellStyleInputBorder)),
            text_primary, hover_bg, false, CellStyle::Input, cx))
        // Total
        .child(render_cell_style_item("Total", "cs-total",
            None,
            Some(app.token(TokenKey::CellStyleTotalBorder)),
            text_primary, hover_bg, true, CellStyle::Total, cx))
        // Note
        .child(render_cell_style_item("Note", "cs-note",
            Some(app.token(TokenKey::CellStyleNoteBg)),
            Some(app.token(TokenKey::CellStyleNoteText)),
            text_primary, hover_bg, false, CellStyle::Note, cx))
}

/// Single item row in the cell style dropdown.
/// Shows a mini cell-preview swatch: fill bg + colored border (realistic style preview).
fn render_cell_style_item(
    label: &'static str,
    id: &'static str,
    swatch_fill: Option<Hsla>,
    swatch_border: Option<Hsla>,
    label_color: Hsla,
    hover_bg: Hsla,
    bold: bool,
    style: CellStyle,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let swatch = match (swatch_fill, swatch_border) {
        (Some(fill), Some(border)) => {
            // Filled swatch with colored border — realistic mini preview
            div().w(px(14.0)).h(px(14.0)).rounded_sm().bg(fill).border_1().border_color(border)
        }
        (None, Some(border)) => {
            // Border-only swatch (Total style — top border indicator)
            div().w(px(14.0)).h(px(14.0)).rounded_sm().border_t_2().border_color(border)
        }
        (Some(fill), None) => {
            div().w(px(14.0)).h(px(14.0)).rounded_sm().bg(fill)
        }
        (None, None) => {
            // No swatch (Normal)
            div().w(px(14.0)).h(px(14.0))
        }
    };

    let label_div = div()
        .text_size(px(12.0))
        .text_color(label_color)
        .when(bold, |d| d.font_weight(FontWeight::BOLD))
        .child(label);

    div()
        .id(SharedString::from(id))
        .flex()
        .items_center()
        .gap(px(8.0))
        .px_2()
        .py(px(4.0))
        .cursor_pointer()
        .hover(move |s| s.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.set_cell_style_selection(style, cx);
            this.ui.format_bar.cell_style_menu_open = false;
            cx.notify();
        }))
        .child(swatch)
        .child(label_div)
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
