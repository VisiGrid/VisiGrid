use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{ContextMenuKind, Spreadsheet};
use crate::theme::TokenKey;

/// Render the right-click context menu overlay, if one is open.
///
/// Returns `None` when no menu is active, so callers can use `.when_some()`.
pub fn render_context_menu(
    app: &Spreadsheet,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) -> Option<impl IntoElement> {
    let state = app.context_menu?;

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let selection_bg = app.token(TokenKey::SelectionBg);

    // Build menu items based on kind
    let items = match state.kind {
        ContextMenuKind::Cell => build_cell_menu(app, text_primary, text_muted, selection_bg, cx),
        ContextMenuKind::RowHeader => build_row_header_menu(app, text_primary, text_muted, selection_bg, cx),
        ContextMenuKind::ColHeader => build_col_header_menu(app, text_primary, text_muted, selection_bg, cx),
    };

    // Estimate menu height for clamping.
    // Items: ~24px each (py(5) + ~14px text). Separators: ~9px (1px + my_1 = 4+1+4).
    let (n_items, n_seps) = match state.kind {
        ContextMenuKind::Cell => (8, 3),
        ContextMenuKind::RowHeader => (4, 1),
        ContextMenuKind::ColHeader => (8, 3),
    };
    let menu_h: f32 = n_items as f32 * 24.0 + n_seps as f32 * 9.0 + 8.0; // + py_1 padding

    // Clamp position to keep menu within window bounds
    let viewport = window.viewport_size();
    let window_w: f32 = viewport.width.into();
    let window_h: f32 = viewport.height.into();
    let menu_w: f32 = 200.0;
    let x: f32 = state.position.x.into();
    let y: f32 = state.position.y.into();
    let x = x.min(window_w - menu_w).max(0.0);
    let y = y.min(window_h - menu_h).max(0.0);

    Some(
        div()
            .id("context-menu")
            .absolute()
            .left(px(x))
            .top(px(y))
            .w(px(menu_w))
            .bg(panel_bg)
            .border_1()
            .border_color(panel_border)
            .rounded_md()
            .shadow_lg()
            .flex()
            .flex_col()
            .py_1()
            .on_mouse_down_out(cx.listener(|this, _, _, cx| {
                this.hide_context_menu(cx);
            }))
            .children(items)
    )
}

/// A single clickable menu item with label and optional shortcut hint.
fn menu_item(
    id: &'static str,
    label: &'static str,
    shortcut: Option<&'static str>,
    enabled: bool,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> AnyElement {
    let el = div()
        .id(id)
        .px_3()
        .py(px(5.0))
        .flex()
        .justify_between()
        .items_center()
        .text_size(px(12.0));

    if enabled {
        el
            .text_color(text_primary)
            .cursor_pointer()
            .hover(move |s| s.bg(selection_bg.opacity(0.5)))
            .child(label)
            .when_some(shortcut, |el, hint| {
                el.child(
                    div()
                        .text_color(text_muted)
                        .text_size(px(11.0))
                        .child(hint)
                )
            })
            .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                this.hide_context_menu(cx);
                on_click(this, cx);
            }))
            .into_any_element()
    } else {
        el
            .text_color(text_muted.opacity(0.5))
            .child(label)
            .when_some(shortcut, |el, hint| {
                el.child(
                    div()
                        .text_color(text_muted.opacity(0.3))
                        .text_size(px(11.0))
                        .child(hint)
                )
            })
            .into_any_element()
    }
}

/// Horizontal separator line.
fn separator(panel_border: Hsla) -> AnyElement {
    div()
        .h(px(1.0))
        .bg(panel_border)
        .mx_2()
        .my_1()
        .into_any_element()
}

/// Build the cell/selection context menu items.
fn build_cell_menu(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Vec<AnyElement> {
    let panel_border = app.token(TokenKey::PanelBorder);
    let has_clipboard = app.internal_clipboard.is_some();

    // On macOS use symbol shortcuts, on other platforms use text
    #[cfg(target_os = "macos")]
    let (cut_hint, copy_hint, paste_hint, paste_v_hint, inspect_hint) = (
        "\u{2318}X", "\u{2318}C", "\u{2318}V", "\u{21E7}\u{2318}V", "\u{21E7}\u{2318}I",
    );
    #[cfg(not(target_os = "macos"))]
    let (cut_hint, copy_hint, paste_hint, paste_v_hint, inspect_hint) = (
        "Ctrl+X", "Ctrl+C", "Ctrl+V", "Ctrl+Shift+V", "Ctrl+Shift+I",
    );

    vec![
        menu_item("ctx-cut", "Cut", Some(cut_hint), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.cut(cx),
        ),
        menu_item("ctx-copy", "Copy", Some(copy_hint), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.copy(cx),
        ),
        menu_item("ctx-paste", "Paste", Some(paste_hint), has_clipboard, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.paste(cx),
        ),
        menu_item("ctx-paste-values", "Paste Values", Some(paste_v_hint), has_clipboard, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.paste_values(cx),
        ),
        separator(panel_border),
        menu_item("ctx-format-painter", "Format Painter", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.start_format_painter(cx),
        ),
        separator(panel_border),
        menu_item("ctx-clear-contents", "Clear Contents", Some("Del"), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_selection(cx),
        ),
        menu_item("ctx-clear-formats", "Clear Formats", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.clear_formatting_selection(cx),
        ),
        separator(panel_border),
        menu_item("ctx-inspect", "Inspect", Some(inspect_hint), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| { this.inspector_visible = !this.inspector_visible; cx.notify(); },
        ),
    ]
}

/// Build the row header context menu items.
fn build_row_header_menu(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Vec<AnyElement> {
    let panel_border = app.token(TokenKey::PanelBorder);

    vec![
        menu_item("ctx-insert-row", "Insert Row", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.insert_rows_or_cols(cx),
        ),
        menu_item("ctx-delete-row", "Delete Row", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_rows_or_cols(cx),
        ),
        separator(panel_border),
        menu_item("ctx-clear-contents", "Clear Contents", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_selection(cx),
        ),
        menu_item("ctx-clear-formats", "Clear Formats", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.clear_formatting_selection(cx),
        ),
    ]
}

/// Build the column header context menu items.
fn build_col_header_menu(
    app: &Spreadsheet,
    text_primary: Hsla,
    text_muted: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> Vec<AnyElement> {
    let panel_border = app.token(TokenKey::PanelBorder);
    let is_col_sel = app.is_col_selection();

    vec![
        menu_item("ctx-insert-col", "Insert Column", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.insert_rows_or_cols(cx),
        ),
        menu_item("ctx-delete-col", "Delete Column", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_rows_or_cols(cx),
        ),
        separator(panel_border),
        menu_item("ctx-clear-contents", "Clear Contents", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_selection(cx),
        ),
        menu_item("ctx-clear-formats", "Clear Formats", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.clear_formatting_selection(cx),
        ),
        separator(panel_border),
        menu_item("ctx-sort-asc", "Sort A\u{2192}Z", None, is_col_sel, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                use visigrid_engine::filter::SortDirection;
                this.sort_by_current_column(SortDirection::Ascending, cx);
            },
        ),
        menu_item("ctx-sort-desc", "Sort Z\u{2192}A", None, is_col_sel, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                use visigrid_engine::filter::SortDirection;
                this.sort_by_current_column(SortDirection::Descending, cx);
            },
        ),
    ]
}
