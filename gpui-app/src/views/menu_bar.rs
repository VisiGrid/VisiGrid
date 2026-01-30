use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::formatting::BorderApplyMode;
use crate::mode::Menu;
use crate::theme::TokenKey;

pub const MENU_HEIGHT: f32 = 22.0;  // Compact chrome height
const DROPDOWN_WIDTH: f32 = 200.0;

/// Render the modern menu bar - compact chrome, not content
pub fn render_menu_bar(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let open_menu = app.open_menu;
    let header_bg = app.token(TokenKey::HeaderBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let selection_bg = app.token(TokenKey::SelectionBg);
    let toolbar_hover = app.token(TokenKey::ToolbarButtonHoverBg);

    // Menu text at ~75% opacity for chrome feel (hover restores full)
    let menu_text = text_primary.opacity(0.75);

    div()
        .flex()
        .flex_shrink_0()
        .h(px(MENU_HEIGHT))
        .w_full()
        .bg(header_bg)
        .border_b_1()
        .border_color(panel_border)
        .items_center()
        .px_1()
        .gap_0()
        .text_size(px(11.0))  // Smaller than content text
        .font_weight(FontWeight::NORMAL)  // Light weight for chrome
        .text_color(menu_text)
        // Group 1: File, Edit, View
        .child(menu_header("File", 'F', Menu::File, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        .child(menu_header("Edit", 'E', Menu::Edit, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        .child(menu_header("View", 'V', Menu::View, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        // Visual separator - extra space before Insert
        .child(div().w(px(8.0)))
        // Group 2: Insert, Format
        .child(menu_header("Insert", 'I', Menu::Insert, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        .child(menu_header("Format", 'O', Menu::Format, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        // Visual separator - extra space before Data
        .child(div().w(px(8.0)))
        // Group 3: Data, Help
        .child(menu_header("Data", 'D', Menu::Data, open_menu, text_primary, selection_bg, toolbar_hover, cx))
        .child(menu_header("Help", 'H', Menu::Help, open_menu, text_primary, selection_bg, toolbar_hover, cx))
}

/// Render the dropdown menu overlay (should be rendered at root level)
pub fn render_menu_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let menu = app.open_menu.unwrap();
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let selection_bg = app.token(TokenKey::SelectionBg);

    render_dropdown(menu, panel_bg, panel_border, text_primary, text_muted, text_disabled, selection_bg, cx)
}

/// Render a menu header button with underlined accelerator
fn menu_header(
    label: &'static str,
    accel: char,
    menu: Menu,
    open_menu: Option<Menu>,
    text_full: Hsla,  // Full opacity text for hover/active states
    selection_bg: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_open = open_menu == Some(menu);

    // Find the position of the accelerator character
    let accel_pos = label.chars().position(|c| c.to_ascii_uppercase() == accel);

    div()
        .id(ElementId::Name(format!("menu-{:?}", menu).into()))
        .flex()
        .items_center()
        .px_3()  // Slightly more horizontal padding
        .h_full()
        .cursor_pointer()
        .when(is_open, move |d: Stateful<Div>| d.bg(selection_bg).text_color(text_full))
        .hover(move |style: StyleRefinement| {
            if is_open {
                style
            } else {
                style.bg(hover_bg).text_color(text_full)  // Restore full opacity on hover
            }
        })
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.toggle_menu(menu, cx);
        }))
        // Hover to switch menus when one is open
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.open_menu.is_some() && this.open_menu != Some(menu) {
                this.open_menu = Some(menu);
                cx.notify();
            }
        }))
        .child(
            // Render label with underlined accelerator
            div()
                .flex()
                .children(
                    label.char_indices().map(|(i, c)| {
                        let should_underline = accel_pos == Some(i);
                        div()
                            .when(should_underline, |d: Div| d.underline())
                            .child(c.to_string())
                    })
                )
        )
}

/// Render the dropdown menu for the given menu type
fn render_dropdown(
    menu: Menu,
    panel_bg: Hsla,
    panel_border: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    text_disabled: Hsla,
    selection_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    // Offsets account for: px_1 (4px), px_3 per item (12px each side), 8px spacers
    let left_offset = match menu {
        Menu::File => 4.0,
        Menu::Edit => 48.0,      // File width ~44px
        Menu::View => 92.0,      // + Edit width ~44px
        Menu::Insert => 148.0,   // + View ~44px + 8px spacer
        Menu::Format => 204.0,   // + Insert ~56px
        Menu::Data => 276.0,     // + Format ~64px + 8px spacer
        Menu::Help => 328.0,     // + Data ~52px
    };

    let menu_content = match menu {
        Menu::File => render_file_menu(text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::Edit => render_edit_menu(text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::View => render_view_menu(text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::Insert => render_insert_menu(text_disabled, panel_border),
        Menu::Format => render_format_menu(text_primary, text_muted, text_disabled, selection_bg, panel_border, cx),
        Menu::Data => render_data_menu(text_primary, text_muted, text_disabled, selection_bg, panel_border, cx),
        Menu::Help => render_help_menu(text_primary, selection_bg, panel_border, cx),
    };

    div()
        .absolute()
        .top(px(MENU_HEIGHT))
        .left(px(left_offset))
        .w(px(DROPDOWN_WIDTH))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .shadow_lg()
        .py_1()
        .child(menu_content)
}

fn render_file_menu(text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("New Workbook", Some("Ctrl+N"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); cx.dispatch_action(&crate::actions::NewWindow); }))
        .child(menu_item("Open...", Some("Ctrl+O"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.open_file(cx); }))
        .child(menu_item("Save", Some("Ctrl+S"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.save(cx); }))
        .child(menu_item("Save As...", Some("Ctrl+Shift+S"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.save_as(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Export as CSV...", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_csv(cx); }))
        .child(menu_item("Export as TSV...", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_tsv(cx); }))
        .child(menu_item("Export as JSON...", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_json(cx); }))
        .child(menu_item("Export to Excel (.xlsx)...", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_xlsx(cx); }))
}

fn render_edit_menu(text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Undo", Some("Ctrl+Z"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.undo(cx); }))
        .child(menu_item("Redo", Some("Ctrl+Y"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.redo(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Cut", Some("Ctrl+X"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.cut(cx); }))
        .child(menu_item("Copy", Some("Ctrl+C"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.copy(cx); }))
        .child(menu_item("Paste", Some("Ctrl+V"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.paste(cx); }))
        .child(menu_item("Paste Values", Some("Ctrl+Shift+V"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.paste_values(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Delete", Some("Del"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.delete_selection(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Find...", Some("Ctrl+F"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_find(cx); }))
        .child(menu_item("Go To...", Some("Ctrl+G"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_goto(cx); }))
}

fn render_view_menu(text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Command Palette", Some("Ctrl+Shift+P"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_palette(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Inspector", Some("Ctrl+Shift+I"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.inspector_visible = !this.inspector_visible; cx.notify(); }))
        .child(menu_separator(border))
        .child(menu_item("Zoom In", Some("Ctrl+Shift+="), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_in(cx); }))
        .child(menu_item("Zoom Out", Some("Ctrl+Shift+-"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_out(cx); }))
        .child(menu_item("Reset Zoom", Some("Ctrl+0"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_reset(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Show Formulas", Some("Ctrl+`"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_show_formulas(cx); }))
        .child(menu_item("Show Zeros", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_show_zeros(cx); }))
        .child(menu_item("Format Bar", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_format_bar(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Freeze Top Row", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_top_row(cx); }))
        .child(menu_item("Freeze First Column", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_first_column(cx); }))
        .child(menu_item("Freeze Panes", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_panes(cx); }))
        .child(menu_item("Unfreeze Panes", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.unfreeze_panes(cx); }))
}

fn render_insert_menu(text_disabled: Hsla, border: Hsla) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item_disabled("Rows", text_disabled))
        .child(menu_item_disabled("Columns", text_disabled))
        .child(menu_separator(border))
        .child(menu_item_disabled("Function...", text_disabled))
}

fn render_format_menu(text_primary: Hsla, text_muted: Hsla, text_disabled: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Bold", Some("Ctrl+B"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_bold(cx); }))
        .child(menu_item("Italic", Some("Ctrl+I"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_italic(cx); }))
        .child(menu_item("Underline", Some("Ctrl+U"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_underline(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Font...", None, text_primary, text_muted, selection_bg, cx, |this, window, cx| { this.close_menu(cx); this.show_font_picker(window, cx); }))
        .child(menu_item_disabled("Cells...", text_disabled))
        .child(menu_separator(border))
        .child(menu_label("Background Color", text_muted))
        .child(color_menu_item("None", None, text_primary, selection_bg, cx))
        .child(color_menu_item("Yellow", Some([255, 255, 0, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Green", Some([198, 239, 206, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Blue", Some([189, 215, 238, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Red", Some([255, 199, 206, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Orange", Some([255, 235, 156, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Purple", Some([204, 192, 218, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Gray", Some([217, 217, 217, 255]), text_primary, selection_bg, cx))
        .child(color_menu_item("Cyan", Some([183, 222, 232, 255]), text_primary, selection_bg, cx))
        .child(menu_separator(border))
        .child(menu_label("Borders", text_muted))
        .child(menu_item("All Borders", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::All, cx); }))
        .child(menu_item("Outline", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::Outline, cx); }))
        .child(menu_item("Clear Borders", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::Clear, cx); }))
        .child(menu_separator(border))
        .child(menu_label("Merge", text_muted))
        .child(menu_item("Merge Cells", Some("Ctrl+Shift+M"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.merge_cells(cx); }))
        .child(menu_item("Unmerge Cells", Some("Ctrl+Shift+U"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.unmerge_cells(cx); }))
        .child(menu_separator(border))
        .child(menu_item_disabled("Row Height...", text_disabled))
        .child(menu_item_disabled("Column Width...", text_disabled))
}

fn render_data_menu(text_primary: Hsla, text_muted: Hsla, text_disabled: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Validation...", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_validation_dialog(cx); }))
        .child(menu_item("Exclude from Validation", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.exclude_from_validation(cx); }))
        .child(menu_item("Clear Validation Exclusions", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.clear_validation_exclusions(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Fill Down", Some("Ctrl+D"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.fill_down(cx); }))
        .child(menu_item("Fill Right", Some("Ctrl+R"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.fill_right(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Circle Invalid Data", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.circle_invalid_data(cx); }))
        .child(menu_item("Clear Invalid Circles", None, text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.clear_invalid_circles(cx); }))
        .child(menu_separator(border))
        .child(menu_item_disabled("Sort...", text_disabled))
        .child(menu_item_disabled("Filter", text_disabled))
        .child(menu_separator(border))
        .child(menu_item("Insert Formula with AI", Some("Ctrl+Shift+A"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_ask_ai(cx); }))
        .child(menu_item("Analyze with AI", Some("Ctrl+Shift+E"), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_analyze(cx); }))
}

fn render_help_menu(text_primary: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let has_license = visigrid_license::is_pro();

    div()
        .flex()
        .flex_col()
        .child(menu_item("About VisiGrid", None, text_primary, text_primary, selection_bg, cx, |this, _window, cx| {
            this.close_menu(cx);
            this.show_about(cx);
        }))
        .child(menu_separator(border))
        .when(has_license, |d| {
            d.child(menu_item("Manage License...", None, text_primary, text_primary, selection_bg, cx, |this, _window, cx| {
                this.close_menu(cx);
                this.show_license(cx);
            }))
        })
        .when(!has_license, |d| {
            d.child(menu_item("Enter License...", None, text_primary, text_primary, selection_bg, cx, |this, _window, cx| {
                this.close_menu(cx);
                this.show_license(cx);
            }))
        })
}

fn menu_item(
    label: &'static str,
    shortcut: Option<&'static str>,
    text_color: Hsla,
    shortcut_color: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Window, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    div()
        .id(ElementId::Name(format!("menuitem-{}", label).into()))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py(px(4.0))  // Slightly tighter vertical padding
        .mx_1()
        .rounded_sm()
        .text_size(px(12.0))  // Match menu bar scale
        .text_color(text_color)
        .cursor_pointer()
        .hover(move |style| style.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
            action(this, window, cx);
        }))
        .child(label)
        .when(shortcut.is_some(), move |d: Stateful<Div>| {
            d.child(
                div()
                    .text_color(shortcut_color)
                    .text_size(px(10.0))  // Smaller shortcuts
                    .child(shortcut.unwrap_or(""))
            )
        })
}

fn menu_item_disabled(label: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .px_3()
        .py(px(4.0))
        .mx_1()
        .text_size(px(12.0))
        .text_color(text_color)
        .child(label)
}

fn menu_separator(border_color: Hsla) -> impl IntoElement {
    div()
        .h(px(1.0))
        .mx_2()
        .my_1()
        .bg(border_color)
}

fn menu_label(label: &'static str, text_color: Hsla) -> impl IntoElement {
    div()
        .px_3()
        .py(px(2.0))
        .mx_1()
        .text_size(px(10.0))
        .text_color(text_color)
        .child(label)
}

fn color_menu_item(
    label: &'static str,
    color: Option<[u8; 4]>,
    text_color: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let swatch_color = color.map(|[r, g, b, _]| {
        Hsla::from(gpui::Rgba {
            r: r as f32 / 255.0,
            g: g as f32 / 255.0,
            b: b as f32 / 255.0,
            a: 1.0,
        })
    });

    div()
        .id(ElementId::Name(format!("color-{}", label).into()))
        .flex()
        .items_center()
        .gap_2()
        .px_3()
        .py(px(4.0))
        .mx_1()
        .rounded_sm()
        .text_size(px(12.0))
        .text_color(text_color)
        .cursor_pointer()
        .hover(move |style| style.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.close_menu(cx);
            this.set_background_color(color, cx);
        }))
        .child(
            div()
                .size(px(12.0))
                .rounded_sm()
                .border_1()
                .border_color(hsla(0.0, 0.0, 0.5, 0.3))
                .when_some(swatch_color, |d, c| d.bg(c))
                .when(swatch_color.is_none(), |d| d.bg(hsla(0.0, 0.0, 1.0, 1.0)))
        )
        .child(label)
}
