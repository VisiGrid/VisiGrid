use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::formatting::BorderApplyMode;
use crate::mode::Menu;
use crate::theme::TokenKey;

pub const MENU_HEIGHT: f32 = 22.0;  // Compact chrome height
const DROPDOWN_WIDTH: f32 = 260.0;

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

/// Render the dropdown menu overlay (should be rendered at root level).
/// Wraps the dropdown in a transparent backdrop that catches clicks below the menu bar
/// to dismiss the menu. The backdrop starts at MENU_HEIGHT so menu headers remain interactive
/// for hover-to-switch.
pub fn render_menu_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let menu = app.open_menu.unwrap();
    let highlight = app.menu_highlight;
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let text_disabled = app.token(TokenKey::TextDisabled);
    let selection_bg = app.token(TokenKey::SelectionBg);

    let dropdown = render_dropdown(menu, highlight, panel_bg, panel_border, text_primary, text_muted, text_disabled, selection_bg, cx);

    // Backdrop: covers everything below the menu bar to catch dismiss clicks
    div()
        .id("menu-backdrop")
        .absolute()
        .top(px(MENU_HEIGHT))
        .left_0()
        .right_0()
        .bottom_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.0))  // Transparent
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            cx.stop_propagation();
            this.close_menu(cx);
        }))
        .on_mouse_down(MouseButton::Right, cx.listener(|this, _, _, cx| {
            cx.stop_propagation();
            this.close_menu(cx);
        }))
        .on_mouse_down(MouseButton::Middle, cx.listener(|this, _, _, cx| {
            cx.stop_propagation();
            this.close_menu(cx);
        }))
        .child(dropdown)
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
                this.menu_highlight = None;
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
    highlight: Option<usize>,
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
        Menu::File => render_file_menu(highlight, text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::Edit => render_edit_menu(highlight, text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::View => render_view_menu(highlight, text_primary, text_muted, selection_bg, panel_border, cx),
        Menu::Insert => render_insert_menu(text_disabled, panel_border),
        Menu::Format => render_format_menu(highlight, text_primary, text_muted, text_disabled, selection_bg, panel_border, cx),
        Menu::Data => render_data_menu(highlight, text_primary, text_muted, text_disabled, selection_bg, panel_border, cx),
        Menu::Help => render_help_menu(highlight, text_primary, selection_bg, panel_border, cx),
    };

    // Dropdown content panel — stop propagation so clicks inside don't reach the backdrop
    div()
        .id("menu-dropdown-content")
        .absolute()
        .top_0()  // relative to backdrop which already starts at MENU_HEIGHT
        .left(px(left_offset))
        .w(px(DROPDOWN_WIDTH))
        .bg(panel_bg)
        .border_1()
        .border_color(panel_border)
        .shadow_lg()
        .py_1()
        .on_mouse_down(MouseButton::Left, cx.listener(|_, _, _, cx| {
            cx.stop_propagation();
        }))
        .on_mouse_down(MouseButton::Right, cx.listener(|_, _, _, cx| {
            cx.stop_propagation();
        }))
        .on_mouse_down(MouseButton::Middle, cx.listener(|_, _, _, cx| {
            cx.stop_propagation();
        }))
        .child(menu_content)
}

fn render_file_menu(highlight: Option<usize>, text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    div()
        .flex()
        .flex_col()
        .child(menu_item("New Workbook", Some("Ctrl+N"), 0, h(0), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); cx.dispatch_action(&crate::actions::NewWindow); }))
        .child(menu_item("Open...", Some("Ctrl+O"), 1, h(1), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.open_file(cx); }))
        .child(menu_item("Save", Some("Ctrl+S"), 2, h(2), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.save(cx); }))
        .child(menu_item("Save As...", Some("Ctrl+Shift+S"), 3, h(3), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.save_as(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Export as CSV...", None, 4, h(4), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_csv(cx); }))
        .child(menu_item("Export as TSV...", None, 5, h(5), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_tsv(cx); }))
        .child(menu_item("Export as JSON...", None, 6, h(6), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_json(cx); }))
        .child(menu_item("Export to Excel (.xlsx)...", None, 7, h(7), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.export_xlsx(cx); }))
}

fn render_edit_menu(highlight: Option<usize>, text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    div()
        .flex()
        .flex_col()
        .child(menu_item("Undo", Some("Ctrl+Z"), 0, h(0), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.undo(cx); }))
        .child(menu_item("Redo", Some("Ctrl+Y"), 1, h(1), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.redo(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Cut", Some("Ctrl+X"), 2, h(2), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.cut(cx); }))
        .child(menu_item("Copy", Some("Ctrl+C"), 3, h(3), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.copy(cx); }))
        .child(menu_item("Paste", Some("Ctrl+V"), 4, h(4), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.paste(cx); }))
        .child(menu_item("Paste Values", Some("Ctrl+Shift+V"), 5, h(5), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.paste_values(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Delete", Some("Del"), 6, h(6), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.delete_selection(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Find...", Some("Ctrl+F"), 7, h(7), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_find(cx); }))
        .child(menu_item("Go To...", Some("Ctrl+G"), 8, h(8), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_goto(cx); }))
}

fn render_view_menu(highlight: Option<usize>, text_primary: Hsla, text_muted: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    div()
        .flex()
        .flex_col()
        .child(menu_item("Command Palette", Some("Ctrl+Shift+P"), 0, h(0), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_palette(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Inspector", Some("Ctrl+Shift+I"), 1, h(1), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.inspector_visible = !this.inspector_visible; cx.notify(); }))
        .child(menu_separator(border))
        .child(menu_item("Zoom In", Some("Ctrl+Shift+="), 2, h(2), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_in(cx); }))
        .child(menu_item("Zoom Out", Some("Ctrl+Shift+-"), 3, h(3), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_out(cx); }))
        .child(menu_item("Reset Zoom", Some("Ctrl+0"), 4, h(4), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.zoom_reset(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Show Formulas", Some("Ctrl+`"), 5, h(5), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_show_formulas(cx); }))
        .child(menu_item("Show Zeros", None, 6, h(6), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_show_zeros(cx); }))
        .child(menu_item("Format Bar", None, 7, h(7), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_format_bar(cx); }))
        .child(menu_item("Minimap", None, 8, h(8), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.minimap_visible = !this.minimap_visible; cx.notify(); }))
        .child(menu_separator(border))
        .child(menu_item("Freeze Top Row", None, 9, h(9), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_top_row(cx); }))
        .child(menu_item("Freeze First Column", None, 10, h(10), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_first_column(cx); }))
        .child(menu_item("Freeze Panes", None, 11, h(11), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.freeze_panes(cx); }))
        .child(menu_item("Unfreeze Panes", None, 12, h(12), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.unfreeze_panes(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Approve Model", None, 13, h(13), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.approve_model(None, cx); }))
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

fn render_format_menu(highlight: Option<usize>, text_primary: Hsla, text_muted: Hsla, text_disabled: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    div()
        .flex()
        .flex_col()
        .child(menu_item("Bold", Some("Ctrl+B"), 0, h(0), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_bold(cx); }))
        .child(menu_item("Italic", Some("Ctrl+I"), 1, h(1), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_italic(cx); }))
        .child(menu_item("Underline", Some("Ctrl+U"), 2, h(2), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.toggle_underline(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Font...", None, 3, h(3), text_primary, text_muted, selection_bg, cx, |this, window, cx| { this.close_menu(cx); this.show_font_picker(window, cx); }))
        .child(menu_item_disabled("Cells...", text_disabled))
        .child(menu_separator(border))
        .child(menu_label("Background Color", text_muted))
        .child(color_menu_item("None", None, 4, h(4), text_primary, selection_bg, cx))
        .child(color_menu_item("Yellow", Some([255, 255, 0, 255]), 5, h(5), text_primary, selection_bg, cx))
        .child(color_menu_item("Green", Some([198, 239, 206, 255]), 6, h(6), text_primary, selection_bg, cx))
        .child(color_menu_item("Blue", Some([189, 215, 238, 255]), 7, h(7), text_primary, selection_bg, cx))
        .child(color_menu_item("Red", Some([255, 199, 206, 255]), 8, h(8), text_primary, selection_bg, cx))
        .child(color_menu_item("Orange", Some([255, 235, 156, 255]), 9, h(9), text_primary, selection_bg, cx))
        .child(color_menu_item("Purple", Some([204, 192, 218, 255]), 10, h(10), text_primary, selection_bg, cx))
        .child(color_menu_item("Gray", Some([217, 217, 217, 255]), 11, h(11), text_primary, selection_bg, cx))
        .child(color_menu_item("Cyan", Some([183, 222, 232, 255]), 12, h(12), text_primary, selection_bg, cx))
        .child(menu_separator(border))
        .child(menu_label("Borders", text_muted))
        .child(menu_item("All Borders", None, 13, h(13), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::All, cx); }))
        .child(menu_item("Outline", None, 14, h(14), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::Outline, cx); }))
        .child(menu_item("Clear Borders", None, 15, h(15), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.apply_borders(BorderApplyMode::Clear, cx); }))
        .child(menu_separator(border))
        .child(menu_label("Merge", text_muted))
        .child(menu_item("Merge Cells", Some("Ctrl+Shift+M"), 16, h(16), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.merge_cells(cx); }))
        .child(menu_item("Unmerge Cells", Some("Ctrl+Shift+U"), 17, h(17), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.unmerge_cells(cx); }))
        .child(menu_separator(border))
        .child(menu_item_disabled("Row Height...", text_disabled))
        .child(menu_item_disabled("Column Width...", text_disabled))
}

fn render_data_menu(highlight: Option<usize>, text_primary: Hsla, text_muted: Hsla, text_disabled: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    div()
        .flex()
        .flex_col()
        .child(menu_item("Validation...", None, 0, h(0), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_validation_dialog(cx); }))
        .child(menu_item("Exclude from Validation", None, 1, h(1), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.exclude_from_validation(cx); }))
        .child(menu_item("Clear Validation Exclusions", None, 2, h(2), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.clear_validation_exclusions(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Fill Down", Some("Ctrl+D"), 3, h(3), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.fill_down(cx); }))
        .child(menu_item("Fill Right", Some("Ctrl+R"), 4, h(4), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.fill_right(cx); }))
        .child(menu_separator(border))
        .child(menu_item("Circle Invalid Data", None, 5, h(5), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.circle_invalid_data(cx); }))
        .child(menu_item("Clear Invalid Circles", None, 6, h(6), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.clear_invalid_circles(cx); }))
        .child(menu_separator(border))
        .child(menu_item_disabled("Sort...", text_disabled))
        .child(menu_item_disabled("Filter", text_disabled))
        .child(menu_separator(border))
        .child(menu_item("Insert Formula with AI", Some("Ctrl+Shift+A"), 7, h(7), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_ask_ai(cx); }))
        .child(menu_item("Analyze with AI", Some("Ctrl+Shift+E"), 8, h(8), text_primary, text_muted, selection_bg, cx, |this, _window, cx| { this.close_menu(cx); this.show_analyze(cx); }))
}

fn render_help_menu(highlight: Option<usize>, text_primary: Hsla, selection_bg: Hsla, border: Hsla, cx: &mut Context<Spreadsheet>) -> Div {
    let h = |i: usize| highlight == Some(i);
    let has_license = visigrid_license::is_pro();
    let license_label = if has_license { "Manage License..." } else { "Enter License..." };

    div()
        .flex()
        .flex_col()
        .child(menu_item("About VisiGrid", None, 0, h(0), text_primary, text_primary, selection_bg, cx, |this, _window, cx| {
            this.close_menu(cx);
            this.show_about(cx);
        }))
        .child(menu_separator(border))
        .child(menu_item(license_label, None, 1, h(1), text_primary, text_primary, selection_bg, cx, |this, _window, cx| {
            this.close_menu(cx);
            this.show_license(cx);
        }))
}

fn menu_item(
    label: &'static str,
    shortcut: Option<&'static str>,
    selectable_index: usize,
    is_highlighted: bool,
    text_color: Hsla,
    shortcut_color: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Window, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    menu_item_with_accel(label, shortcut, None, selectable_index, is_highlighted, text_color, shortcut_color, hover_bg, cx, action)
}

fn menu_item_with_accel(
    label: &'static str,
    shortcut: Option<&'static str>,
    accel: Option<char>,
    selectable_index: usize,
    is_highlighted: bool,
    text_color: Hsla,
    shortcut_color: Hsla,
    hover_bg: Hsla,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Window, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    // Resolve the accelerator character and find its position in the label
    let accel_char = crate::menu_model::resolve_accel(label, accel);
    let accel_pos = label.chars().position(|c| c.to_ascii_lowercase() == accel_char);

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
        .when(is_highlighted, move |d: Stateful<Div>| d.bg(hover_bg))
        .hover(move |style| style.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, window, cx| {
            action(this, window, cx);
        }))
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.menu_highlight != Some(selectable_index) {
                this.menu_highlight = Some(selectable_index);
                cx.notify();
            }
        }))
        .child(
            // Render label with underlined accelerator character
            div()
                .flex()
                .flex_1()
                .children(
                    label.char_indices().map(move |(i, c)| {
                        let should_underline = accel_pos == Some(i);
                        div()
                            .when(should_underline, |d: Div| d.underline())
                            .child(c.to_string())
                    })
                )
        )
        .when(shortcut.is_some(), move |d: Stateful<Div>| {
            d.child(
                div()
                    .text_color(shortcut_color)
                    .text_size(px(10.0))  // Smaller shortcuts
                    .flex_shrink_0()
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
    selectable_index: usize,
    is_highlighted: bool,
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

    // Resolve accelerator for underline
    let accel_char = crate::menu_model::resolve_accel(label, None);
    let accel_pos = label.chars().position(|c| c.to_ascii_lowercase() == accel_char);

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
        .when(is_highlighted, move |d: Stateful<Div>| d.bg(hover_bg))
        .hover(move |style| style.bg(hover_bg))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            this.close_menu(cx);
            this.set_background_color(color, cx);
        }))
        .on_mouse_move(cx.listener(move |this, _, _, cx| {
            if this.menu_highlight != Some(selectable_index) {
                this.menu_highlight = Some(selectable_index);
                cx.notify();
            }
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
        .child(
            div()
                .flex()
                .children(
                    label.char_indices().map(move |(i, c)| {
                        let should_underline = accel_pos == Some(i);
                        div()
                            .when(should_underline, |d: Div| d.underline())
                            .child(c.to_string())
                    })
                )
        )
}
