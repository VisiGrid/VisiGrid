//! The ribbon is a presentation of existing commands, not another editing path.
use crate::app::{SelectionFormatState, Spreadsheet, TriState};
use crate::search::CommandId;
use crate::theme::TokenKey;
use crate::toolbar::{RibbonTab, RIBBON_BODY_HEIGHT, RIBBON_TABS_HEIGHT};
use gpui::{prelude::FluentBuilder, *};

#[derive(Clone, Copy, PartialEq, Eq)]
enum Item {
    Command(CommandId, &'static str),
    NumberFormat,
    Styles,
    FontSize,
    TextColor,
    InsertRows,
    InsertCols,
}

struct Group {
    label: &'static str,
    width: f32,
    items: &'static [Item],
    font: bool,
}

fn groups(tab: RibbonTab) -> Vec<Group> {
    use CommandId::*;
    use Item::Command as C;
    let group = |label, width, items| Group {
        label,
        width,
        items,
        font: false,
    };
    match tab {
        RibbonTab::Home => vec![
            group(
                "Clipboard",
                178.,
                &[
                    C(Paste, "Paste"),
                    C(Cut, "Cut"),
                    C(Copy, "Copy"),
                    C(PasteValues, "Values"),
                    C(PasteSpecial, "Paste special"),
                    C(FormatPainter, "Paint format"),
                ],
            ),
            Group {
                label: "Font",
                width: 250.,
                items: &[
                    C(SelectFont, "Font…"),
                    C(ToggleBold, "Bold"),
                    C(ToggleItalic, "Italic"),
                    C(ToggleUnderline, "Underline"),
                    C(FillColor, "Fill color"),
                    Item::FontSize,
                    Item::TextColor,
                ],
                font: true,
            },
            group(
                "Alignment",
                98.,
                &[
                    C(AlignLeft, "Left"),
                    C(AlignCenter, "Center"),
                    C(AlignRight, "Right"),
                ],
            ),
            group(
                "Number",
                112.,
                &[
                    Item::NumberFormat,
                    C(FormatCurrency, "Currency"),
                    C(FormatPercent, "Percent"),
                ],
            ),
            group(
                "Styles",
                142.,
                &[
                    Item::Styles,
                    C(AddConditionalFormat, "Conditional…"),
                    C(ManageConditionalFormats, "Manage rules"),
                ],
            ),
            group(
                "Editing",
                150.,
                &[
                    C(AutoSum, "AutoSum"),
                    C(FindInCells, "Find"),
                    C(FillDown, "Fill down"),
                    C(FillRight, "Fill right"),
                ],
            ),
        ],
        RibbonTab::Insert => vec![
            group(
                "Worksheet",
                220.,
                &[C(AddSheet, "Add sheet"), Item::InsertRows, Item::InsertCols],
            ),
            group("Analysis", 155., &[C(InsertPivotTable, "Pivot table…")]),
            group(
                "Names",
                200.,
                &[C(ExtractNamedRange, "Create named range…")],
            ),
        ],
        RibbonTab::Formulas => vec![
            group(
                "Calculate",
                150.,
                &[C(AutoSum, "AutoSum"), C(Recalculate, "Recalculate")],
            ),
            group(
                "Names",
                210.,
                &[C(ExtractNamedRange, "Create named range…")],
            ),
            group(
                "Inspect",
                220.,
                &[
                    C(ToggleTrace, "Trace dependencies"),
                    C(CycleTracePrecedent, "Next precedent"),
                    C(CycleTraceDependent, "Next dependent"),
                    C(ToggleInspector, "Inspector"),
                ],
            ),
        ],
        RibbonTab::Data => vec![
            group(
                "Sort & filter",
                220.,
                &[
                    C(SortAscending, "Sort A → Z"),
                    C(SortDescending, "Sort Z → A"),
                    C(ClearSort, "Clear sort"),
                    C(ToggleAutoFilter, "Filter"),
                ],
            ),
            group(
                "Clean & validate",
                190.,
                &[
                    C(TrimWhitespace, "Trim whitespace"),
                    C(ValidationDialog, "Data validation…"),
                ],
            ),
            group(
                "Pivot tables",
                265.,
                &[
                    C(InsertPivotTable, "Pivot table…"),
                    C(RefreshPivot, "Refresh pivot"),
                    C(RefreshAllPivots, "Refresh all pivots"),
                    C(EditPivotFields, "Pivot fields"),
                ],
            ),
        ],
        RibbonTab::View => vec![
            group(
                "Freeze",
                195.,
                &[
                    C(FreezeTopRow, "Freeze top row"),
                    C(FreezeFirstColumn, "Freeze first column"),
                    C(FreezePanes, "Freeze panes"),
                    C(UnfreezePanes, "Unfreeze"),
                ],
            ),
            group(
                "Zoom",
                120.,
                &[
                    C(ZoomIn, "Zoom in"),
                    C(ZoomOut, "Zoom out"),
                    C(ZoomReset, "100%"),
                ],
            ),
            group(
                "Panels",
                145.,
                &[
                    C(ToggleInspector, "Inspector"),
                    C(ToggleMinimap, "Minimap"),
                    C(ToggleCommentsSidebar, "Comments"),
                ],
            ),
            group(
                "Appearance",
                155.,
                &[C(SelectTheme, "Theme…"), C(ToggleZenMode, "Zen mode")],
            ),
        ],
    }
}

fn item_label(item: Item) -> &'static str {
    match item {
        Item::Command(_, label) => label,
        Item::NumberFormat => "Number format ▾",
        Item::Styles => "Cell styles ▾",
        Item::FontSize => "Font size",
        Item::TextColor => "Text color",
        Item::InsertRows => "Insert rows",
        Item::InsertCols => "Insert columns",
    }
}

fn disabled_reason(app: &Spreadsheet, item: Item) -> Option<&'static str> {
    // Keep modes/dialog focus owned by the existing editor. Ribbon commands
    // operate on cells; layout selection itself still preserves editing buffers.
    if !app.mode.is_navigation() && app.mode != crate::mode::Mode::FormatPainter {
        return Some("Finish the current edit or dialog first");
    }
    let reads_only = matches!(
        item,
        Item::Command(
            CommandId::Copy
                | CommandId::FindInCells
                | CommandId::ToggleInspector
                | CommandId::ToggleMinimap
                | CommandId::ToggleCommentsSidebar
                | CommandId::ToggleTrace
                | CommandId::CycleTracePrecedent
                | CommandId::CycleTraceDependent
                | CommandId::ZoomIn
                | CommandId::ZoomOut
                | CommandId::ZoomReset
                | CommandId::SelectTheme
                | CommandId::ToggleZenMode,
            _
        )
    );
    if !reads_only && !app.can_edit() {
        return Some("Return to an editable sheet to use this command");
    }
    match item {
        Item::InsertRows if !app.is_row_selection() => {
            Some("Select entire rows first (Shift+Space)")
        }
        Item::InsertCols if !app.is_col_selection() => {
            Some("Select entire columns first (Ctrl+Space)")
        }
        Item::Command(CommandId::Undo, _) if !app.history.can_undo() => Some("Nothing to undo"),
        Item::Command(CommandId::Redo, _) if !app.history.can_redo() => Some("Nothing to redo"),
        _ => None,
    }
}

fn invoke(
    app: &mut Spreadsheet,
    item: Item,
    x: f32,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    if let Some(reason) = disabled_reason(app, item) {
        app.status_message = Some(reason.into());
        cx.notify();
        return;
    }
    app.ui.desktop_keytips.clear();
    window.focus(&app.focus_handle, cx);
    if item != Item::FontSize {
        app.ui.ribbon.group_menu = None;
    }
    // The picker is rendered at root level and must remain anchored after the
    // temporary ribbon closes. Its vertical anchor is geometry, not a row index.
    app.ui.ribbon.temporary = item == Item::FontSize && app.ribbon_collapsed(cx);
    app.ui.format_bar.popup_x = x;
    match item {
        Item::Command(command, _) => app.dispatch_command(command, window, cx),
        Item::NumberFormat => {
            app.ui.format_bar.number_format_menu_open = true;
        }
        Item::Styles => {
            app.ui.format_bar.cell_style_menu_open = true;
        }
        Item::FontSize => super::format_bar::begin_font_size_edit(app, window, cx),
        Item::TextColor => {
            app.show_color_picker(crate::color_palette::ColorTarget::Text, window, cx)
        }
        Item::InsertRows | Item::InsertCols => app.insert_rows_or_cols(cx),
    }
    cx.notify();
}

fn command_button(
    app: &Spreadsheet,
    item: Item,
    id: String,
    state: &SelectionFormatState,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div> {
    let reason = disabled_reason(app, item);
    let accent = app.token(TokenKey::Accent);
    let active = match item {
        Item::Command(CommandId::ToggleAutoFilter, _) => app.filter_state.is_enabled(),
        Item::Command(CommandId::ToggleInspector, _) => app.inspector_visible,
        Item::Command(CommandId::ToggleMinimap, _) => app.minimap_visible,
        Item::Command(CommandId::ToggleCommentsSidebar, _) => app.comments_sidebar_visible,
        Item::Command(CommandId::ToggleTrace, _) => app.trace_enabled,
        Item::Command(CommandId::AlignLeft, _) => matches!(
            state.alignment,
            TriState::Uniform(visigrid_engine::cell::Alignment::Left)
        ),
        Item::Command(CommandId::AlignCenter, _) => matches!(
            state.alignment,
            TriState::Uniform(visigrid_engine::cell::Alignment::Center)
        ),
        Item::Command(CommandId::AlignRight, _) => matches!(
            state.alignment,
            TriState::Uniform(visigrid_engine::cell::Alignment::Right)
        ),
        _ => false,
    };
    let text = app.token(if reason.is_some() {
        TokenKey::TextDisabled
    } else {
        TokenKey::TextPrimary
    });
    let hover = app.token(TokenKey::ToolbarButtonHoverBg);
    let tooltip: SharedString = reason
        .map(str::to_owned)
        .unwrap_or_else(|| match item {
            Item::Command(command, _) => command.name().to_string(),
            _ => item_label(item).to_string(),
        })
        .into();
    div()
        .id(SharedString::from(id))
        .w_full()
        .min_w_0()
        .overflow_hidden()
        .h(px(21.))
        .px_1()
        .rounded_sm()
        .role(Role::Button)
        .aria_label(item_label(item))
        .tab_index(0)
        .focus(|s| s.bg(accent.opacity(0.2)))
        .when(active, |d| d.bg(accent.opacity(0.12)))
        .text_size(px(12.))
        .text_color(text)
        .whitespace_nowrap()
        .when(reason.is_none(), |d| {
            d.cursor_pointer().hover(move |s| s.bg(hover))
        })
        .on_click(cx.listener(move |this, event: &ClickEvent, window, cx| {
            cx.stop_propagation();
            invoke(this, item, f32::from(event.position().x), window, cx);
        }))
        .on_action(
            cx.listener(move |this, _: &crate::actions::ConfirmEdit, window, cx| {
                invoke(this, item, 8., window, cx);
            }),
        )
        .tooltip(move |_, cx| cx.new(|_| RibbonTooltip(tooltip.clone())).into())
        .flex()
        .items_center()
        .gap_1()
        .when(app.ui.desktop_keytips.ribbon(app.ui.ribbon.tab), |d| {
            let code = keytip_for(app.ui.ribbon.tab, item);
            d.when(
                code.to_ascii_lowercase()
                    .starts_with(&app.ui.desktop_keytips.prefix),
                |d| d.child(keytip_badge(app, code).flex_shrink_0()),
            )
        })
        .child(
            div()
                .flex_1()
                .min_w_0()
                .overflow_hidden()
                .text_ellipsis()
                .child(item_label(item)),
        )
}

/// Group prefixes are stable, unique within each tab, and separate from root letters.
fn group_prefix(label: &str) -> char {
    match label {
        "Clipboard" | "Calculate" | "Clean & validate" => 'C',
        "Font" | "Freeze" => 'F',
        "Alignment" | "Analysis" | "Appearance" => 'A',
        "Number" | "Names" => 'N',
        "Styles" | "Sort & filter" => 'S',
        "Editing" => 'E',
        "Worksheet" => 'W',
        "Inspect" => 'I',
        "Pivot tables" | "Panels" => 'P',
        "Zoom" => 'Z',
        _ => unreachable!("Every ribbon group needs a KeyTip prefix"),
    }
}

fn keytip_for(tab: RibbonTab, item: Item) -> String {
    for group in groups(tab) {
        if let Some(index) = group.items.iter().position(|i| *i == item) {
            return format!("{}{}", group_prefix(group.label), index + 1);
        }
    }
    unreachable!("KeyTip command must belong to the visible tab")
}

pub(super) fn keytip_badge(app: &Spreadsheet, code: String) -> Div {
    div()
        .px(px(3.))
        .h(px(15.))
        .text_size(px(10.))
        .line_height(px(15.))
        .rounded_sm()
        .border_1()
        .border_color(app.token(TokenKey::Accent))
        .bg(app.token(TokenKey::PanelBg))
        .text_color(app.token(TokenKey::TextPrimary))
        .child(code)
}

pub(super) fn font_keytip(app: &Spreadsheet, code: &str, control: impl IntoElement) -> AnyElement {
    if !app.ui.desktop_keytips.ribbon(RibbonTab::Home)
        || !code
            .to_ascii_lowercase()
            .starts_with(&app.ui.desktop_keytips.prefix)
    {
        return control.into_any_element();
    }
    div()
        .relative()
        .child(control)
        .child(
            keytip_badge(app, code.into())
                .absolute()
                .right_0()
                .top(px(-8.)),
        )
        .into_any_element()
}

/// KeyTips reach commands even when their group has collapsed. Typing the group
/// letter opens its existing overflow panel so the remaining digit is discoverable.
pub(crate) fn handle_keytip(
    app: &mut Spreadsheet,
    tab: RibbonTab,
    input: &str,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    if input.len() != 1 || !input.chars().all(|c| c.is_ascii_alphanumeric()) {
        return;
    }
    let candidate = format!("{}{}", app.ui.desktop_keytips.prefix, input);
    let entries = groups(tab);
    for (index, group) in entries.iter().enumerate() {
        for (row, item) in group.items.iter().enumerate() {
            let code = format!(
                "{}{}",
                group_prefix(group.label).to_ascii_lowercase(),
                row + 1
            );
            if candidate == code {
                invoke(app, *item, app.ui.ribbon.group_menu_x, window, cx);
                // Disabled commands retain hints, allowing another choice.
                app.ui.desktop_keytips.prefix.clear();
                cx.notify();
                return;
            }
        }
        if candidate == group_prefix(group.label).to_ascii_lowercase().to_string() {
            app.ui.desktop_keytips.prefix = candidate;
            let widths = group_widths(
                &entries.iter().map(|g| g.width).collect::<Vec<_>>(),
                f32::from(app.window_size.width) - 16.,
            );
            if widths[index] < group.width {
                app.ui.ribbon.group_menu = Some(index);
                app.ui.ribbon.group_menu_x = 8. + widths[..index].iter().sum::<f32>();
            }
            cx.notify();
            return;
        }
    }
    // An invalid digit resets the prefix; never replay it into the cell editor.
    app.ui.desktop_keytips.prefix.clear();
    cx.notify();
}

/// Collapse whole groups from right to left, keeping the user's spatial order.
fn group_widths(widths: &[f32], available: f32) -> Vec<f32> {
    let mut result = widths.to_vec();
    let mut total: f32 = widths.iter().sum();
    for index in (0..result.len()).rev() {
        if total <= available {
            break;
        }
        total -= result[index] - 76.;
        result[index] = 76.;
    }
    result
}

pub fn render_ribbon(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let text = app.token(TokenKey::TextPrimary);
    let active = app.ui.ribbon.tab;
    let collapsed = app.ribbon_collapsed(cx);
    let mut tabs = div()
        .id("ribbon-tabs")
        .tab_group()
        .tab_index(0)
        .tab_stop(false)
        .role(Role::TabList)
        .h(px(RIBBON_TABS_HEIGHT))
        .flex()
        .items_center()
        .px_2()
        .gap_1()
        .border_b_1()
        .border_color(border);
    for (tab_index, tab) in RibbonTab::ALL.into_iter().enumerate() {
        tabs = tabs.child(
            div()
                .id(SharedString::from(format!("ribbon-tab-{}", tab.label())))
                .tab_index(tab_index as isize)
                .role(Role::Tab)
                .aria_label(tab.label())
                .aria_selected(active == tab)
                .focus(|s| s.bg(accent.opacity(0.2)))
                .h_full()
                .px_3()
                .flex()
                .items_center()
                .cursor_pointer()
                .text_size(px(12.))
                .text_color(text)
                .border_b_2()
                .border_color(if active == tab { accent } else { bg })
                .when(active == tab, |d| {
                    d.text_color(accent).bg(accent.opacity(0.08))
                })
                .hover(|s| s.bg(border.opacity(0.2)))
                .on_click(cx.listener(move |this, event: &ClickEvent, window, cx| {
                    cx.stop_propagation();
                    if event.click_count() == 2 {
                        this.toggle_ribbon_collapsed(window, cx);
                    } else {
                        this.select_ribbon_tab(tab, window, cx);
                    }
                }))
                .on_action(
                    cx.listener(move |this, _: &crate::actions::ConfirmEdit, window, cx| {
                        this.select_ribbon_tab(tab, window, cx);
                        if this.ui.ribbon.temporary {
                            window.focus(&this.ui.ribbon.temporary_focus, cx);
                        }
                    }),
                )
                .child(tab.label())
                .relative()
                .when(app.ui.desktop_keytips.root(), |d| {
                    let key = crate::desktop_keytips::TAB_KEYS
                        .iter()
                        .find(|(_, t)| *t == tab)
                        .unwrap()
                        .0;
                    d.child(
                        keytip_badge(app, key.to_ascii_uppercase().to_string())
                            .absolute()
                            .right_0()
                            .bottom_0(),
                    )
                }),
        );
    }
    tabs = tabs.child(div().flex_1()).child(
        div()
            .id("ribbon-collapse")
            .tab_index(6)
            .role(Role::Button)
            .aria_label("Collapse or expand ribbon")
            .aria_expanded(!collapsed)
            .px_2()
            .cursor_pointer()
            .text_color(text)
            .text_size(px(12.))
            .child(if collapsed { "Expand" } else { "Collapse" })
            .on_click(cx.listener(|this, _, window, cx| {
                cx.stop_propagation();
                this.toggle_ribbon_collapsed(window, cx);
            }))
            .on_action(
                cx.listener(|this, _: &crate::actions::ConfirmEdit, window, cx| {
                    this.toggle_ribbon_collapsed(window, cx)
                }),
            ),
    );
    keyboard_container(div().track_focus(&app.ui.ribbon.focus), cx)
        .flex()
        .flex_col()
        .flex_shrink_0()
        .w_full()
        .bg(bg)
        .h(px(app.toolbar_geometry(cx).toolbar_height))
        .child(tabs)
        .when(!collapsed, |d| d.child(render_body(app, window, cx)))
}

fn render_body(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let border = app.token(TokenKey::PanelBorder);
    let muted = app.token(TokenKey::TextMuted);
    let tab = app.ui.ribbon.tab;
    let state = app.selection_format_state(cx);
    let entries = groups(tab);
    let widths = group_widths(
        &entries.iter().map(|g| g.width).collect::<Vec<_>>(),
        f32::from(app.window_size.width) - 16.,
    );
    let mut body = div()
        .tab_group()
        .tab_index(1)
        .tab_stop(false)
        .h(px(RIBBON_BODY_HEIGHT))
        .w_full()
        .flex()
        .px_2()
        .py_1()
        .bg(app.token(TokenKey::PanelBg))
        .border_b_1()
        .border_color(border)
        .on_scroll_wheel(|_, _, cx| cx.stop_propagation());
    for (index, (group, width)) in entries.iter().zip(widths).enumerate() {
        let mut el = div()
            .w(px(width))
            .h_full()
            .flex_shrink_0()
            .flex()
            .flex_col()
            .justify_between()
            .px_1()
            .border_r_1()
            .border_color(border);
        if width < group.width {
            el = el.child(
                div()
                    .id(SharedString::from(format!("ribbon-group-{index}")))
                    .tab_index(0)
                    .role(Role::Button)
                    .aria_label(group.label)
                    .size_full()
                    .flex()
                    .flex_col()
                    .justify_center()
                    .text_size(px(11.))
                    .text_color(app.token(TokenKey::TextPrimary))
                    .cursor_pointer()
                    .hover(|s| s.bg(border.opacity(0.2)))
                    .on_click(cx.listener(move |this, event: &ClickEvent, window, cx| {
                        cx.stop_propagation();
                        this.ui.ribbon.group_menu = if this.ui.ribbon.group_menu == Some(index) {
                            None
                        } else {
                            Some(index)
                        };
                        this.ui.ribbon.group_menu_x = f32::from(event.position().x);
                        if this.ui.ribbon.group_menu.is_some() {
                            window.focus(&this.ui.ribbon.group_focus, cx);
                        }
                        cx.notify();
                    }))
                    .child(group.label)
                    .child("▾")
                    .relative()
                    .when(app.ui.desktop_keytips.ribbon(tab), |d| {
                        d.child(
                            keytip_badge(app, group_prefix(group.label).to_string())
                                .absolute()
                                .right_0()
                                .top_0(),
                        )
                    }),
            );
        } else {
            if group.font && disabled_reason(app, group.items[0]).is_none() {
                el = el.child(super::format_bar::render_ribbon_font_controls(
                    app, &state, window, cx,
                ));
            } else {
                let mut columns = div().flex().gap_1();
                for (column, items) in group.items.chunks(3).enumerate() {
                    let mut col = div().flex().flex_col().flex_1().min_w_0();
                    for (row, item) in items.iter().enumerate() {
                        col = col.child(command_button(
                            app,
                            *item,
                            format!("ribbon-{index}-{column}-{row}"),
                            &state,
                            cx,
                        ));
                    }
                    columns = columns.child(col);
                }
                el = el.child(columns);
            }
            el = el.child(
                div()
                    .text_size(px(11.))
                    .text_color(muted)
                    .child(group.label),
            );
        }
        body = body.child(el);
    }
    body
}

pub fn render_temporary(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let top = app.toolbar_geometry(cx).toolbar_top + RIBBON_TABS_HEIGHT;
    keyboard_container(div().track_focus(&app.ui.ribbon.temporary_focus), cx)
        .id("ribbon-temporary")
        .absolute()
        .top(px(top))
        .left_0()
        .right_0()
        .shadow_lg()
        .on_mouse_down_out(cx.listener(|this, event: &MouseDownEvent, window, cx| {
            let top = this.toolbar_geometry(cx).toolbar_top;
            let y = f32::from(event.position.y);
            if y < top || y >= top + RIBBON_TABS_HEIGHT {
                if this.ui.format_bar.size_editing {
                    super::format_bar::commit_font_size(this, cx);
                    window.focus(&this.focus_handle, cx);
                }
                this.ui.ribbon.temporary = false;
                cx.notify();
            }
        }))
        .child(render_body(app, window, cx))
}

pub fn render_group_menu(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let index = app.ui.ribbon.group_menu.unwrap_or(0);
    let state = app.selection_format_state(cx);
    let entries = groups(app.ui.ribbon.tab);
    let mut panel = keyboard_container(div().track_focus(&app.ui.ribbon.group_focus), cx)
        .id("ribbon-group-menu")
        .absolute()
        .left(px(app
            .ui
            .ribbon
            .group_menu_x
            .min((f32::from(app.window_size.width) - 280.).max(0.))))
        .top(px(app.toolbar_geometry(cx).popup_top))
        .w(px(270.))
        .p_2()
        .shadow_lg()
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(app.token(TokenKey::PanelBorder))
        .on_mouse_down_out(cx.listener(|this, _, window, cx| {
            if this.ui.format_bar.size_editing {
                super::format_bar::commit_font_size(this, cx);
                window.focus(&this.focus_handle, cx);
            }
            this.ui.ribbon.group_menu = None;
            cx.notify();
        }))
        .flex()
        .flex_col();
    if let Some(group) = entries.get(index) {
        if group.font && disabled_reason(app, group.items[0]).is_none() {
            return panel.child(super::format_bar::render_ribbon_font_controls(
                app, &state, window, cx,
            ));
        }
        for (i, item) in group.items.iter().enumerate() {
            panel = panel.child(command_button(
                app,
                *item,
                format!("ribbon-overflow-{i}"),
                &state,
                cx,
            ));
        }
    }
    panel
}

fn keyboard_container(el: Div, cx: &mut Context<Spreadsheet>) -> Div {
    el.key_context("Ribbon")
        .tab_group()
        .tab_index(0)
        .tab_stop(false)
        .on_action(|_: &crate::actions::TabNext, window, cx| window.focus_next(cx))
        .on_action(|_: &crate::actions::TabPrev, window, cx| window.focus_prev(cx))
        .on_action(|_: &crate::actions::MoveRight, window, cx| window.focus_next(cx))
        .on_action(|_: &crate::actions::MoveLeft, window, cx| window.focus_prev(cx))
        .on_action(|_: &crate::actions::MoveDown, window, cx| window.focus_next(cx))
        .on_action(|_: &crate::actions::MoveUp, window, cx| window.focus_prev(cx))
        .on_action(
            cx.listener(|this, _: &crate::actions::CancelEdit, window, cx| {
                this.close_toolbar_popups();
                window.focus(&this.focus_handle, cx);
                cx.notify();
            }),
        )
        .on_key_down(|_, _, cx| cx.stop_propagation())
}

struct RibbonTooltip(SharedString);
impl Render for RibbonTooltip {
    fn render(&mut self, _: &mut Window, _: &mut Context<Self>) -> impl IntoElement {
        div().px_2().py_1().text_size(px(12.)).child(self.0.clone())
    }
}

#[cfg(test)]
mod tests {
    use super::{group_prefix, group_widths, groups, keytip_for, Item, RibbonTab};
    #[test]
    fn command_keytips_are_unique_and_have_no_ambiguous_prefixes() {
        for tab in RibbonTab::ALL {
            let mut codes = Vec::new();
            let mut prefixes = std::collections::HashSet::new();
            for group in groups(tab) {
                assert!(prefixes.insert(group_prefix(group.label)));
                assert!(group.items.len() <= 9);
                for item in group.items {
                    let code = keytip_for(tab, *item);
                    assert!(codes
                        .iter()
                        .all(|old: &String| !old.starts_with(&code) && !code.starts_with(old)));
                    codes.push(code);
                }
            }
        }
    }
    #[test]
    fn font_badges_match_the_shared_controls() {
        use crate::search::CommandId::*;
        for (item, code) in [
            (Item::Command(SelectFont, "Font…"), "F1"),
            (Item::Command(ToggleBold, "Bold"), "F2"),
            (Item::Command(ToggleItalic, "Italic"), "F3"),
            (Item::Command(ToggleUnderline, "Underline"), "F4"),
            (Item::Command(FillColor, "Fill color"), "F5"),
            (Item::FontSize, "F6"),
            (Item::TextColor, "F7"),
        ] {
            assert_eq!(keytip_for(RibbonTab::Home, item), code);
        }
    }

    #[test]
    fn whole_groups_collapse_without_reordering_or_growing_the_ribbon() {
        for tab in RibbonTab::ALL {
            let widths: Vec<_> = groups(tab).iter().map(|g| g.width).collect();
            for available in [600., 800., 1024., 1440.] {
                let resolved = group_widths(&widths, available);
                assert!(resolved.iter().sum::<f32>() <= available);
                assert!(resolved
                    .iter()
                    .zip(&widths)
                    .all(|(a, b)| *a == 76. || a == b));
            }
        }
    }
}
