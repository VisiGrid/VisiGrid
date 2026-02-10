use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{ContextMenuKind, Spreadsheet};
use crate::theme::TokenKey;
use crate::ui::{popup, clamp_to_viewport};

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
        ContextMenuKind::Cell => (13, 5),
        ContextMenuKind::RowHeader => (4, 1),
        ContextMenuKind::ColHeader => (8, 3),
    };
    let menu_h: f32 = n_items as f32 * 24.0 + n_seps as f32 * 9.0 + 8.0; // + py_1 padding

    // Clamp position to keep menu within window bounds
    let viewport = window.viewport_size();
    let menu_w: f32 = 200.0;
    let x: f32 = state.position.x.into();
    let y: f32 = state.position.y.into();
    let (x, y) = clamp_to_viewport(
        x, y, menu_w, menu_h,
        viewport.width.into(), viewport.height.into(),
    );

    Some(
        popup("context-menu", panel_bg, panel_border, |this, cx| this.hide_context_menu(cx), cx)
            .left(px(x))
            .top(px(y))
            .w(px(menu_w))
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

    #[cfg(target_os = "macos")]
    let format_hint = "\u{2318}1";
    #[cfg(not(target_os = "macos"))]
    let format_hint = "Ctrl+1";

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
        menu_item("ctx-insert-row", "Insert Row", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                let ((min_row, _), (max_row, _)) = this.selection_range();
                let count = max_row - min_row + 1;
                this.insert_rows(min_row, count, cx);
            },
        ),
        menu_item("ctx-insert-col", "Insert Column", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                let ((_, min_col), (_, max_col)) = this.selection_range();
                let count = max_col - min_col + 1;
                this.insert_cols(min_col, count, cx);
            },
        ),
        menu_item("ctx-delete-row", "Delete Row", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                let ((min_row, _), (max_row, _)) = this.selection_range();
                let count = max_row - min_row + 1;
                this.delete_rows(min_row, count, cx);
            },
        ),
        menu_item("ctx-delete-col", "Delete Column", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                let ((_, min_col), (_, max_col)) = this.selection_range();
                let count = max_col - min_col + 1;
                this.delete_cols(min_col, count, cx);
            },
        ),
        separator(panel_border),
        menu_item("ctx-clear-contents", "Clear Contents", Some("Del"), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.delete_selection(cx),
        ),
        menu_item("ctx-clear-formats", "Clear Formats", None, true, text_primary, text_muted, selection_bg, cx,
            |this, cx| this.clear_formatting_selection(cx),
        ),
        separator(panel_border),
        menu_item("ctx-format-cells", "Format Cells...", Some(format_hint), true, text_primary, text_muted, selection_bg, cx,
            |this, cx| {
                this.inspector_visible = true;
                this.inspector_tab = crate::mode::InspectorTab::Format;
                cx.notify();
            },
        ),
        menu_item("ctx-inspect", "Cell Inspector", Some(inspect_hint), true, text_primary, text_muted, selection_bg, cx,
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

/// Expected cell context menu item IDs in display order.
/// This is the contract for Shift+F10 — the first 4 items MUST be clipboard operations.
/// Used by regression tests to prevent accidental reordering.
pub fn cell_context_menu_item_ids() -> &'static [&'static str] {
    &[
        "ctx-cut",
        "ctx-copy",
        "ctx-paste",
        "ctx-paste-values",
        // separator
        "ctx-insert-row",
        "ctx-insert-col",
        "ctx-delete-row",
        "ctx-delete-col",
        // separator
        "ctx-clear-contents",
        "ctx-clear-formats",
        // separator
        "ctx-format-cells",
        "ctx-inspect",
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
