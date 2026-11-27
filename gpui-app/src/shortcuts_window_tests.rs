//! The Keyboard Shortcuts window, driven through the real Spreadsheet (#198).
use crate::{app::Spreadsheet, mode::Mode, search::CommandId};
use gpui::{px, TestAppContext, WindowHandle};

fn setup(cx: &mut TestAppContext) -> WindowHandle<Spreadsheet> {
    cx.update(|cx| {
        crate::settings::init_settings_store(cx);
        crate::load_embedded_fonts(cx);
        cx.set_global(crate::session::SessionManager::new());
        cx.set_global(crate::window_registry::WindowRegistry::new());
        crate::keybindings::register(cx, crate::settings::ModifierStyle::Platform);
    });
    let view = cx.add_window(Spreadsheet::new);
    view.update(cx, |app, window, cx| window.focus(&app.focus_handle, cx)).unwrap();
    view
}

#[gpui::test]
fn the_palette_command_opens_a_searchable_list(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        app.dispatch_command(CommandId::ShowShortcuts, window, cx);
        assert_eq!(app.mode, Mode::Shortcuts);
    }).unwrap();
    cx.run_until_parked();

    // Arrow and page keys scroll the list; the grid behind does not move.
    cx.simulate_keystrokes(view.into(), "down down pagedown");
    view.update(cx, |app, _, _| {
        assert!(app.shortcuts_scroll.offset().y < px(0.0), "the list scrolls");
        assert_eq!(app.view_state.active_cell(), (0, 0));
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "pageup pageup pageup");
    view.update(cx, |app, _, _| assert_eq!(app.shortcuts_scroll.offset().y, px(0.0))).unwrap();

    // Typing searches instead of editing the cell or running a shortcut.
    cx.simulate_keystrokes(view.into(), "f i l l x backspace enter");
    view.update(cx, |app, _, _| {
        assert_eq!(app.mode, Mode::Shortcuts);
        assert_eq!(app.shortcuts_query, "fill");
    }).unwrap();

    cx.simulate_keystrokes(view.into(), "escape");
    view.update(cx, |app, _, _| {
        assert_eq!(app.mode, Mode::Navigation);
        assert!(app.shortcuts_query.is_empty());
    }).unwrap();
}

#[gpui::test]
fn the_help_menu_item_opens_the_same_window(cx: &mut TestAppContext) {
    let view = setup(cx);
    cx.dispatch_action(view.into(), crate::actions::ShowKeyboardShortcuts);
    view.update(cx, |app, _, _| assert_eq!(app.mode, Mode::Shortcuts)).unwrap();
}
