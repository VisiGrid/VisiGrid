use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::Spreadsheet;
use crate::mode::Menu;

const MENU_HEIGHT: f32 = 24.0;
const DROPDOWN_WIDTH: f32 = 200.0;

/// Render the Excel 2003-style menu bar (header row only)
pub fn render_menu_bar(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let open_menu = app.open_menu;

    div()
        .flex()
        .flex_shrink_0()
        .h(px(MENU_HEIGHT))
        .w_full()
        .bg(rgb(0x2d2d2d))
        .border_b_1()
        .border_color(rgb(0x3d3d3d))
        .items_center()
        .px_1()
        .gap_0()
        .text_sm()
        .text_color(rgb(0xcccccc))
        .child(menu_header("File", 'F', Menu::File, open_menu, cx))
        .child(menu_header("Edit", 'E', Menu::Edit, open_menu, cx))
        .child(menu_header("View", 'V', Menu::View, open_menu, cx))
        .child(menu_header("Insert", 'I', Menu::Insert, open_menu, cx))
        .child(menu_header("Format", 'O', Menu::Format, open_menu, cx))
        .child(menu_header("Data", 'D', Menu::Data, open_menu, cx))
        .child(menu_header("Help", 'H', Menu::Help, open_menu, cx))
}

/// Render the dropdown menu overlay (should be rendered at root level)
pub fn render_menu_dropdown(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let menu = app.open_menu.unwrap();
    render_dropdown(menu, cx)
}

/// Render a menu header button with underlined accelerator
fn menu_header(
    label: &'static str,
    accel: char,
    menu: Menu,
    open_menu: Option<Menu>,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let is_open = open_menu == Some(menu);

    // Find the position of the accelerator character
    let accel_pos = label.chars().position(|c| c.to_ascii_uppercase() == accel);

    div()
        .id(ElementId::Name(format!("menu-{:?}", menu).into()))
        .flex()
        .items_center()
        .px_2()
        .h_full()
        .cursor_pointer()
        .when(is_open, |d: Stateful<Div>| d.bg(rgb(0x094771)))
        .hover(|style: StyleRefinement| {
            if is_open {
                style
            } else {
                style.bg(rgb(0x404040))
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
fn render_dropdown(menu: Menu, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let left_offset = match menu {
        Menu::File => 4.0,
        Menu::Edit => 40.0,
        Menu::View => 76.0,
        Menu::Insert => 118.0,
        Menu::Format => 166.0,
        Menu::Data => 226.0,
        Menu::Help => 270.0,
    };

    let menu_content = match menu {
        Menu::File => render_file_menu(cx),
        Menu::Edit => render_edit_menu(cx),
        Menu::View => render_view_menu(cx),
        Menu::Insert => render_insert_menu(),
        Menu::Format => render_format_menu(cx),
        Menu::Data => render_data_menu(cx),
        Menu::Help => render_help_menu(cx),
    };

    div()
        .absolute()
        .top(px(MENU_HEIGHT))
        .left(px(left_offset))
        .w(px(DROPDOWN_WIDTH))
        .bg(rgb(0x252526))
        .border_1()
        .border_color(rgb(0x3d3d3d))
        .shadow_lg()
        .py_1()
        .child(menu_content)
}

fn render_file_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("New", Some("Ctrl+N"), cx, |this, cx| { this.close_menu(cx); this.new_file(cx); }))
        .child(menu_item("Open...", Some("Ctrl+O"), cx, |this, cx| { this.close_menu(cx); this.open_file(cx); }))
        .child(menu_item("Save", Some("Ctrl+S"), cx, |this, cx| { this.close_menu(cx); this.save(cx); }))
        .child(menu_item("Save As...", Some("Ctrl+Shift+S"), cx, |this, cx| { this.close_menu(cx); this.save_as(cx); }))
        .child(menu_separator())
        .child(menu_item("Export as CSV...", None, cx, |this, cx| { this.close_menu(cx); this.export_csv(cx); }))
}

fn render_edit_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Undo", Some("Ctrl+Z"), cx, |this, cx| { this.close_menu(cx); this.undo(cx); }))
        .child(menu_item("Redo", Some("Ctrl+Y"), cx, |this, cx| { this.close_menu(cx); this.redo(cx); }))
        .child(menu_separator())
        .child(menu_item("Cut", Some("Ctrl+X"), cx, |this, cx| { this.close_menu(cx); this.cut(cx); }))
        .child(menu_item("Copy", Some("Ctrl+C"), cx, |this, cx| { this.close_menu(cx); this.copy(cx); }))
        .child(menu_item("Paste", Some("Ctrl+V"), cx, |this, cx| { this.close_menu(cx); this.paste(cx); }))
        .child(menu_separator())
        .child(menu_item("Delete", Some("Del"), cx, |this, cx| { this.close_menu(cx); this.delete_selection(cx); }))
        .child(menu_separator())
        .child(menu_item("Find...", Some("Ctrl+F"), cx, |this, cx| { this.close_menu(cx); this.show_find(cx); }))
        .child(menu_item("Go To...", Some("Ctrl+G"), cx, |this, cx| { this.close_menu(cx); this.show_goto(cx); }))
}

fn render_view_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Command Palette", Some("Ctrl+Shift+P"), cx, |this, cx| { this.close_menu(cx); this.toggle_palette(cx); }))
}

fn render_insert_menu() -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item_disabled("Rows"))
        .child(menu_item_disabled("Columns"))
        .child(menu_separator())
        .child(menu_item_disabled("Function..."))
}

fn render_format_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Bold", Some("Ctrl+B"), cx, |this, cx| { this.close_menu(cx); this.toggle_bold(cx); }))
        .child(menu_item("Italic", Some("Ctrl+I"), cx, |this, cx| { this.close_menu(cx); this.toggle_italic(cx); }))
        .child(menu_item("Underline", Some("Ctrl+U"), cx, |this, cx| { this.close_menu(cx); this.toggle_underline(cx); }))
        .child(menu_separator())
        .child(menu_item("Font...", None, cx, |this, cx| { this.close_menu(cx); this.show_font_picker(cx); }))
        .child(menu_item_disabled("Cells..."))
        .child(menu_separator())
        .child(menu_item_disabled("Row Height..."))
        .child(menu_item_disabled("Column Width..."))
}

fn render_data_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("Fill Down", Some("Ctrl+D"), cx, |this, cx| { this.close_menu(cx); this.fill_down(cx); }))
        .child(menu_item("Fill Right", Some("Ctrl+R"), cx, |this, cx| { this.close_menu(cx); this.fill_right(cx); }))
        .child(menu_separator())
        .child(menu_item_disabled("Sort..."))
        .child(menu_item_disabled("Filter"))
}

fn render_help_menu(cx: &mut Context<Spreadsheet>) -> Div {
    div()
        .flex()
        .flex_col()
        .child(menu_item("About VisiGrid", None, cx, |this, cx| {
            this.close_menu(cx);
            this.status_message = Some("VisiGrid - A modern spreadsheet application".to_string());
            cx.notify();
        }))
}

fn menu_item(
    label: &'static str,
    shortcut: Option<&'static str>,
    cx: &mut Context<Spreadsheet>,
    action: impl Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
) -> impl IntoElement {
    div()
        .id(ElementId::Name(format!("menuitem-{}", label).into()))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py_1()
        .mx_1()
        .rounded_sm()
        .text_sm()
        .text_color(rgb(0xcccccc))
        .cursor_pointer()
        .hover(|style| style.bg(rgb(0x094771)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            action(this, cx);
        }))
        .child(label)
        .when(shortcut.is_some(), |d: Stateful<Div>| {
            d.child(
                div()
                    .text_color(rgb(0x888888))
                    .text_xs()
                    .child(shortcut.unwrap_or(""))
            )
        })
}

fn menu_item_disabled(label: &'static str) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .px_3()
        .py_1()
        .mx_1()
        .text_sm()
        .text_color(rgb(0x666666))
        .child(label)
}

fn menu_separator() -> impl IntoElement {
    div()
        .h(px(1.0))
        .mx_2()
        .my_1()
        .bg(rgb(0x3d3d3d))
}
