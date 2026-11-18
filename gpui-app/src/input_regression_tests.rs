//! Issue #173: exercise real key dispatch and the resize/undo path.
use crate::app::Spreadsheet;
use gpui::{TestAppContext, WindowHandle};

fn setup(cx: &mut TestAppContext) -> WindowHandle<Spreadsheet> {
    cx.update(|cx| {
        crate::settings::init_settings_store(cx);
        crate::load_embedded_fonts(cx);
        cx.set_global(crate::session::SessionManager::new());
        cx.set_global(crate::window_registry::WindowRegistry::new());
        crate::keybindings::register(cx, crate::settings::ModifierStyle::Platform);
    });
    cx.add_window(Spreadsheet::new)
}

#[gpui::test]
fn delete_routes_to_cell_contents_or_forward_text_deletion(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| {
            wb.set_cell_value_tracked(0, 0, 0, "abc");
            wb.set_cell_value_tracked(0, 1, 0, "keep");
        });
        app.start_edit(cx);
        app.edit_cursor = 1;
        cx.notify();
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "delete");
    view.update(cx, |app, _, cx| {
        assert_eq!(app.edit_value, "ac");
        assert_eq!(app.edit_cursor, 1);
        assert_eq!(
            app.sheet(cx).get_raw(0, 0),
            "abc",
            "draft must stay private"
        );
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "escape delete");
    view.update(cx, |app, _, cx| {
        assert_eq!(app.sheet(cx).get_raw(0, 0), "");
        assert_eq!(app.sheet(cx).get_raw(1, 0), "keep");
        app.undo(cx);
        assert_eq!(app.sheet(cx).get_raw(0, 0), "abc");
        app.view_state.selection_end = Some((1, 0));
        cx.notify();
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "delete");
    view.update(cx, |app, _, cx| {
        assert_eq!(app.sheet(cx).get_raw(0, 0), "");
        assert_eq!(app.sheet(cx).get_raw(1, 0), "");
        app.undo(cx);
        assert_eq!(app.sheet(cx).get_raw(0, 0), "abc");
        assert_eq!(app.sheet(cx).get_raw(1, 0), "keep");
    })
    .unwrap();
}

#[gpui::test]
fn selected_columns_resize_together_and_undo_restores_each_original(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let id = app.cached_sheet_id();
        app.view_state.selected = (0, 1);
        app.view_state.selection_end = Some((10, 2));
        app.view_state
            .additional_selections
            .push(((0, 4), Some((10, 4))));
        app.set_col_width(1, 80.);
        app.set_col_width(4, 150.);
        app.metrics.zoom = 2.;
        let untouched = app.col_width(3);
        app.begin_col_resize(1, 100., cx);
        app.update_col_resize(140., cx);
        app.update_col_resize(180., cx);
        app.finish_col_resize(cx);
        for col in [1, 2, 4] {
            assert_eq!(app.col_width(col), 120.);
        }
        assert_eq!(app.col_width(3), untouched);
        assert_eq!(app.history.undo_count(), 1);
        app.undo(cx);
        assert_eq!(app.col_width(1), 80.);
        assert_eq!(app.col_width(4), 150.);
        assert!(
            !app.col_widths[&id].contains_key(&2),
            "restore preference-based width"
        );
        app.redo(cx);
        for col in [1, 2, 4] {
            assert_eq!(app.col_width(col), 120.);
        }
        // Clicking another header boundary must not resize the current selection.
        app.begin_col_resize(6, 100., cx);
        app.update_col_resize(140., cx);
        app.finish_col_resize(cx);
        for col in [1, 2, 4] {
            assert_eq!(app.col_width(col), 120.);
        }
        assert_eq!(app.history.undo_count(), 2);
        let before = app.col_widths.clone();
        app.begin_col_resize(1, 100., cx);
        app.update_col_resize(140., cx);
        app.update_col_resize(100., cx);
        app.finish_col_resize(cx);
        assert_eq!(app.col_widths, before);
        assert_eq!(app.history.undo_count(), 2, "returning to start is a no-op");
    })
    .unwrap();
}
