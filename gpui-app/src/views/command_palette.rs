//! Command palette rendering
//!
//! This module provides the UI for the command palette overlay.
//! Search logic is handled by the search engine in `search.rs`; grouping and
//! windowing live in `command_palette.rs`.

use std::time::Duration;
use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::actions::{
    PaletteUp, PaletteDown, PaletteExecute, PalettePreview, PaletteCancel,
    PalettePageUp, PalettePageDown, PaletteHome, PaletteEnd,
};
use crate::app::{Spreadsheet, PaletteScope};
use crate::command_palette::{is_open_from_disk, palette_row_has_second_line, wheel_rows};
use crate::search::{SearchAction, SearchItem, SearchKind, SearchQuery};
use crate::theme::TokenKey;

/// Colors the palette draws with, read once per frame.
#[derive(Clone, Copy)]
struct Palette {
    panel_bg: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
    disabled: Hsla,
    selection_bg: Hsla,
    selection_text: Hsla,
    hover: Hsla,
    accent: Hsla,
    ok: Hsla,
    cell_ref: Hsla,
    function: Hsla,
}

impl Palette {
    fn new(app: &Spreadsheet) -> Self {
        Self {
            panel_bg: app.token(TokenKey::PanelBg),
            border: app.token(TokenKey::PanelBorder),
            text: app.token(TokenKey::TextPrimary),
            muted: app.token(TokenKey::TextMuted),
            disabled: app.token(TokenKey::TextDisabled),
            selection_bg: app.token(TokenKey::SelectionBg),
            selection_text: app.token(TokenKey::SelectionText),
            hover: app.token(TokenKey::ToolbarButtonHoverBg),
            accent: app.token(TokenKey::Accent),
            ok: app.token(TokenKey::Ok),
            cell_ref: app.token(TokenKey::FormulaCellRef),
            function: app.token(TokenKey::FormulaFunction),
        }
    }

    /// Icon tint per kind: commands accent, files green, cells and ranges the
    /// cell-reference color, functions the function color.
    fn tint(&self, item: &SearchItem) -> Hsla {
        match item.kind {
            _ if is_open_from_disk(item) => self.muted,
            SearchKind::Command => self.accent,
            SearchKind::RecentFile => self.ok,
            SearchKind::Formula => self.function,
            SearchKind::Setting => self.muted,
            SearchKind::Cell
            | SearchKind::NamedRange
            | SearchKind::GoTo
            | SearchKind::Reference
            | SearchKind::Precedent => self.cell_ref,
        }
    }
}

/// Label in the scope pill.
fn scope_label(scope: &PaletteScope) -> &'static str {
    match scope {
        PaletteScope::Menu(cat) => cat.name(),
        PaletteScope::QuickOpen => "Open file",
    }
}

/// Render the command palette overlay
pub fn render_command_palette(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let c = Palette::new(app);
    let query = app.palette_query.clone();
    let has_query = !query.is_empty();
    let scope = app.palette_scope;

    let placeholder = match scope {
        Some(PaletteScope::QuickOpen) => "Search recent files…".to_string(),
        Some(PaletteScope::Menu(cat)) => format!("Search {} commands…", cat.name()),
        None => "Search commands, files and ranges…".to_string(),
    };

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(72.0))
        .bg(hsla(0.0, 0.0, 0.0, 0.4))
        // Click outside to close
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_palette(cx);
        }))
        .child(
            div()
                .key_context("CommandPalette")
                .track_focus(&app.focus_handle)
                .w(px(600.0))
                .bg(c.panel_bg)
                .border_1()
                .border_color(c.border)
                .rounded_lg()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Action handlers
                .on_action(cx.listener(|this, _: &PaletteUp, _, cx| {
                    this.palette_up(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteDown, _, cx| {
                    this.palette_down(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteExecute, window, cx| {
                    this.palette_execute(window, cx);
                }))
                .on_action(cx.listener(|this, _: &PalettePreview, _, cx| {
                    this.palette_preview(cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteCancel, _, cx| {
                    this.hide_palette(cx);
                }))
                .on_action(cx.listener(|this, _: &PalettePageUp, _, cx| {
                    this.palette_move(-(Spreadsheet::PALETTE_VISIBLE as isize), cx);
                }))
                .on_action(cx.listener(|this, _: &PalettePageDown, _, cx| {
                    this.palette_move(Spreadsheet::PALETTE_VISIBLE as isize, cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteHome, _, cx| {
                    this.palette_move(isize::MIN / 2, cx);
                }))
                .on_action(cx.listener(|this, _: &PaletteEnd, _, cx| {
                    this.palette_move(isize::MAX / 2, cx);
                }))
                // Stop click propagation on the palette itself
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .child(render_input(&c, scope, &query, &placeholder))
                .when(!has_query && scope.is_none(), |d| d.child(render_prefix_chips(&c, cx)))
                .child(render_results(app, &c, cx))
                .child(render_footer(app, &c))
        )
}

fn render_input(c: &Palette, scope: Option<PaletteScope>, query: &str, placeholder: &str) -> impl IntoElement {
    let has_query = !query.is_empty();
    div()
        .flex()
        .items_center()
        .gap_2()
        .px(px(14.0))
        .py(px(11.0))
        .border_b_1()
        .border_color(c.border)
        .when_some(scope, |d, scope| {
            d.child(
                div()
                    .px(px(8.0))
                    .py(px(2.0))
                    .rounded(px(5.0))
                    .bg(c.accent.opacity(0.16))
                    .text_size(px(12.0))
                    .text_color(c.text)
                    .font_weight(FontWeight::MEDIUM)
                    .child(scope_label(&scope))
            )
        })
        .child({
            // Blinking cursor: after the query, or before the placeholder,
            // where typing will start
            let cursor = div()
                .w(px(1.0))
                .h(px(15.0))
                .bg(c.text)
                .mx(px(1.0))
                .with_animation(
                    "cursor-blink",
                    Animation::new(Duration::from_millis(530))
                        .repeat()
                        .with_easing(pulsating_between(0.0, 1.0)),
                    |div, delta| {
                        let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                        div.opacity(opacity)
                    },
                );
            let text = div()
                .text_color(if has_query { c.text } else { c.disabled })
                .child(if has_query { query.to_string() } else { placeholder.to_string() });
            let line = div().flex_1().flex().items_center().text_size(px(14.0));
            if has_query { line.child(text).child(cursor) } else { line.child(cursor).child(text) }
        })
        .child(keycap("Esc", c))
}

/// A key drawn as a cap: "Esc", "Ctrl", "↵".
fn keycap(label: impl Into<SharedString>, c: &Palette) -> Div {
    div()
        .px(px(5.0))
        .rounded(px(4.0))
        .border_1()
        .border_color(c.border)
        .text_size(px(11.0))
        .text_color(c.muted)
        .child(label.into())
}

/// "Ctrl+Shift+L" → ["Ctrl", "Shift", "L"]; "Ctrl++" → ["Ctrl", "+"].
fn shortcut_keys(shortcut: &str) -> Vec<String> {
    let mut keys = Vec::new();
    let mut rest = shortcut;
    while !rest.is_empty() {
        // Skip the first character so a literal "+" key is never a separator
        let first = rest.chars().next().map_or(0, char::len_utf8);
        match rest[first..].find('+').map(|i| i + first) {
            Some(i) => {
                keys.push(rest[..i].to_string());
                rest = &rest[i + 1..];
            }
            None => {
                keys.push(rest.to_string());
                break;
            }
        }
    }
    keys
}

/// Clickable prefix chips, shown when nothing is typed.
fn render_prefix_chips(c: &Palette, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let hints: [(char, &str); 6] = [
        ('>', "commands"),
        (':', "go to cell"),
        ('$', "named range"),
        ('=', "function"),
        ('@', "cell value"),
        ('#', "setting"),
    ];

    let mut row = div()
        .flex()
        .items_center()
        .flex_wrap()
        .gap(px(6.0))
        .px(px(14.0))
        .py(px(7.0))
        .border_b_1()
        .border_color(c.border);

    for (ch, label) in hints {
        let hover = c.hover;
        row = row.child(
            div()
                .id(SharedString::from(format!("hint-{}", ch)))
                .flex()
                .items_center()
                .gap(px(5.0))
                .px(px(7.0))
                .py(px(2.0))
                .rounded(px(5.0))
                .border_1()
                .border_color(c.border)
                .cursor_pointer()
                .hover(move |s| s.bg(hover))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    cx.stop_propagation();
                    this.palette_insert_char(ch, cx);
                }))
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(c.text)
                        .font_weight(FontWeight::MEDIUM)
                        .child(ch.to_string())
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(c.muted)
                        .child(label)
                )
        );
    }

    row
}

fn render_results(app: &Spreadsheet, c: &Palette, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let results = app.palette_results();
    let offset = app.palette_scroll_offset.min(results.len());
    let end = app.palette_window_end(offset);

    let mut list = div()
        .id("palette-list")
        .flex()
        .flex_col()
        .py(px(4.0))
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
            let dy: f32 = event.delta.pixel_delta(px(Spreadsheet::PALETTE_ROW_H)).y.into();
            let rows = wheel_rows(&mut this.palette_wheel_px, dy, Spreadsheet::PALETTE_ROW_H);
            if rows != 0 {
                this.palette_scroll(rows, cx);
            }
            cx.stop_propagation();
        }));

    if results.is_empty() {
        return list.child(render_empty(app, c));
    }

    for idx in offset..end {
        if let Some(section) = app.palette_heading_at(idx, idx == offset) {
            list = list.child(if section.title.is_empty() {
                div()
                    .h(px(Spreadsheet::PALETTE_DIVIDER_H))
                    .flex()
                    .items_center()
                    .px(px(14.0))
                    .child(div().h(px(1.0)).w_full().bg(c.border))
            } else {
                div()
                    .h(px(Spreadsheet::PALETTE_HEADER_H))
                    .flex()
                    .items_end()
                    .gap_2()
                    .px(px(14.0))
                    .pb(px(5.0))
                    .child(
                        div()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::SEMIBOLD)
                            .text_color(c.muted)
                            .child(section.title.to_uppercase())
                    )
                    .when(section.len > 3, |d| {
                        d.child(
                            div()
                                .text_size(px(11.0))
                                .text_color(c.disabled)
                                .child(section.len.to_string())
                        )
                    })
            });
        }
        list = list.child(render_row(&results[idx], idx == app.palette_selected, idx, c, cx));
    }

    list
}

fn render_empty(app: &Spreadsheet, c: &Palette) -> impl IntoElement {
    let query = SearchQuery::parse(&app.palette_query);
    let (title, hint) = match query.prefix {
        Some(':') if query.needle.is_empty() => ("Type a cell".to_string(), "Like B5 or AA100"),
        Some(':') => (format!("\"{}\" is not a cell", query.needle), "Try a cell like B5"),
        _ if query.needle.is_empty() => ("Nothing here yet".to_string(), "Type to search, or a prefix: > : $ = @ #"),
        _ => (format!("No matches for \"{}\"", query.needle), "Try fewer letters, or a prefix: > : $ = @ #"),
    };
    div()
        .py(px(24.0))
        .flex()
        .flex_col()
        .items_center()
        .gap_1()
        .child(div().text_size(px(13.0)).text_color(c.text).child(title))
        .child(div().text_size(px(12.0)).text_color(c.muted).child(hint))
}

/// Render text with highlighted spans
fn render_highlighted_text(
    text: &str,
    highlights: &[(usize, usize)],
    normal_color: Hsla,
    highlight_color: Hsla,
) -> Div {
    let mut container = div()
        .flex()
        .items_center()
        .overflow_hidden()
        .whitespace_nowrap()
        .text_size(px(13.0));

    if highlights.is_empty() {
        return container.text_color(normal_color).child(text.to_string());
    }

    let chars: Vec<char> = text.chars().collect();
    let mut pos = 0;

    for &(start, end) in highlights {
        // Clamp to valid range
        let start = start.min(chars.len()).max(pos);
        let end = end.min(chars.len());

        if pos < start {
            let segment: String = chars[pos..start].iter().collect();
            container = container.child(div().text_color(normal_color).child(segment));
        }

        if start < end {
            let segment: String = chars[start..end].iter().collect();
            container = container.child(
                div()
                    .text_color(highlight_color)
                    .font_weight(FontWeight::SEMIBOLD)
                    .child(segment)
            );
        }

        pos = end.max(pos);
    }

    if pos < chars.len() {
        let segment: String = chars[pos..].iter().collect();
        container = container.child(div().text_color(normal_color).child(segment));
    }

    container
}

fn render_row(
    item: &SearchItem,
    is_selected: bool,
    idx: usize,
    c: &Palette,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let two_line = palette_row_has_second_line(item);
    let tint = c.tint(item);

    // On an opaque selection row (VisiCalc inverse video), all text flips to
    // the selection text color.
    let (title_color, dim, highlight) = if is_selected {
        (c.selection_text, c.selection_text.opacity(0.75), c.selection_text)
    } else {
        (c.text, c.muted, c.accent)
    };
    let (icon_bg, icon_fg) = if is_selected {
        (c.selection_text.opacity(0.15), c.selection_text)
    } else {
        (tint.opacity(0.15), tint)
    };

    let mut row = div()
        .id(ElementId::NamedInteger("palette-item".into(), idx as u64))
        .h(px(if two_line { Spreadsheet::PALETTE_ROW2_H } else { Spreadsheet::PALETTE_ROW_H }))
        .flex()
        .items_center()
        .gap(px(10.0))
        .px(px(14.0))
        .cursor_pointer()
        .when(is_selected, |d| d.bg(c.selection_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
            this.palette_selected = idx;
            this.palette_execute(window, cx);
        }))
        // Kind icon in a tinted square
        .child(
            div()
                .size(px(20.0))
                .flex_shrink_0()
                .flex()
                .items_center()
                .justify_center()
                .rounded(px(5.0))
                .bg(icon_bg)
                .text_color(icon_fg)
                .text_size(px(12.0))
                .child(item.kind.icon())
        )
        // Title, and a second line for files, functions and settings
        .child(
            div()
                .flex_1()
                .min_w_0()
                .flex()
                .flex_col()
                .child(render_highlighted_text(&item.title, &item.highlights, title_color, highlight))
                .when(two_line, |d| {
                    d.child(
                        div()
                            .text_size(px(11.0))
                            .text_color(dim)
                            .truncate()
                            .child(item.subtitle.clone().unwrap_or_default())
                    )
                })
        )
        .when_some(item.meta.clone(), |d, meta| {
            d.child(div().flex_shrink_0().text_size(px(11.0)).text_color(dim).child(meta))
        });

    // A command's shortcut, as key caps
    if item.kind == SearchKind::Command {
        if let Some(shortcut) = &item.subtitle {
            let mut keys = div().flex().flex_shrink_0().gap(px(3.0));
            for key in shortcut_keys(shortcut) {
                keys = keys.child(keycap(key, c).when(is_selected, |d| {
                    d.text_color(c.selection_text).border_color(c.selection_text.opacity(0.4))
                }));
            }
            row = row.child(keys);
        }
    }

    if !is_selected {
        let hover = c.hover;
        row = row.hover(move |s| s.bg(hover));
    }

    row
}

/// What Enter (and Ctrl/Shift+Enter) will do to the selected row.
fn footer_keys(item: &SearchItem) -> Vec<(&'static str, &'static str)> {
    let mut keys = match &item.action {
        _ if is_open_from_disk(item) => vec![("↵", "Browse")],
        SearchAction::OpenFile(_) => vec![("↵", "Open")],
        SearchAction::JumpToCell { .. } => vec![("↵", "Go"), ("Shift ↵", "Peek")],
        SearchAction::JumpToNamedRange { .. } => vec![("↵", "Go")],
        SearchAction::InsertFormula { .. } => vec![("↵", "Insert")],
        SearchAction::OpenSetting { .. } => vec![("↵", "Open setting")],
        SearchAction::ShowReferences { .. } | SearchAction::ShowPrecedents { .. } => vec![("↵", "Show")],
        _ => vec![("↵", "Run")],
    };
    match (&item.action, &item.secondary_action) {
        (SearchAction::OpenFile(_), Some(_)) => keys.push(("Ctrl ↵", "Copy path")),
        (SearchAction::InsertFormula { .. }, Some(_)) => keys.push(("Ctrl ↵", "Show help")),
        (_, Some(SearchAction::CopyToClipboard { .. })) => keys.push(("Ctrl ↵", "Copy reference")),
        (_, Some(_)) => keys.push(("Ctrl ↵", "More")),
        _ => {}
    }
    keys
}

fn render_footer(app: &Spreadsheet, c: &Palette) -> impl IntoElement {
    let has_query = !app.palette_query.is_empty();
    let mut keys: Vec<(&'static str, &'static str)> = if app.palette_previewing {
        vec![("Esc", "Restore"), ("↵", "Keep")]
    } else {
        app.palette_results()
            .get(app.palette_selected)
            .map(footer_keys)
            .unwrap_or_default()
    };
    if app.palette_scope.is_some() && !has_query {
        keys.push(("Backspace", "All commands"));
    }

    let count = app.palette_total_results;
    let count_text = if has_query || app.palette_scope.is_some() {
        if count == 1 { "1 result".to_string() } else { format!("{count} results") }
    } else {
        String::new()
    };

    let mut left = div().flex().items_center().gap(px(14.0));
    for (key, label) in keys {
        left = left.child(
            div()
                .flex()
                .items_center()
                .gap(px(6.0))
                .child(keycap(key, c))
                .child(label)
        );
    }

    div()
        .flex()
        .items_center()
        .justify_between()
        .px(px(14.0))
        .py(px(7.0))
        .border_t_1()
        .border_color(c.border)
        .text_size(px(11.0))
        .text_color(c.muted)
        .child(left)
        .child(div().child(count_text))
}

#[cfg(test)]
mod tests {
    use super::shortcut_keys;

    #[test]
    fn shortcuts_split_into_key_caps() {
        assert_eq!(shortcut_keys("Ctrl+Shift+L"), ["Ctrl", "Shift", "L"]);
        assert_eq!(shortcut_keys("Ctrl++"), ["Ctrl", "+"]);
        assert_eq!(shortcut_keys("Ctrl+-"), ["Ctrl", "-"]);
        assert_eq!(shortcut_keys("Alt+="), ["Alt", "="]);
        assert_eq!(shortcut_keys("F9"), ["F9"]);
        assert_eq!(shortcut_keys("Ctrl+Shift+*"), ["Ctrl", "Shift", "*"]);
    }
}
