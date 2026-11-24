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

#[gpui::test]
fn insert_count_dialog_validates_keys_and_inserts_in_one_undo_step(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| { wb.set_cell_value_tracked(0, 1, 0, "original"); });
        app.view_state.selected = (1, 0);
        app.show_insert_count(Axis::Row, false, cx);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "0 enter");
    view.update(cx, |app, _, cx| {
        assert!(app.count_dialog.is_some());
        assert_eq!(app.history.undo_count(), 0);
        assert_eq!(app.sheet(cx).get_raw(1, 0), "original");
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "backspace 5 enter");
    view.update(cx, |app, _, cx| {
        assert!(app.count_dialog.is_none());
        assert_eq!(app.sheet(cx).get_raw(6, 0), "original");
        assert_eq!(app.history.undo_count(), 1);
        app.undo(cx);
        assert_eq!(app.sheet(cx).get_raw(1, 0), "original");
        app.show_insert_count(Axis::Col, true, cx);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "3 escape");
    view.update(cx, |app, _, _| { assert!(app.count_dialog.is_none()); assert_eq!(app.history.undo_count(), 0); }).unwrap();
}

#[gpui::test]
fn insert_count_dialog_targets_all_four_sides_and_checks_remaining_capacity(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.selected = (4, 3);
        app.view_state.selection_end = Some((6, 5));
        for (axis, after, at, capacity) in [
            (Axis::Row, false, 4, crate::app::NUM_ROWS - 4),
            (Axis::Row, true, 7, crate::app::NUM_ROWS - 7),
            (Axis::Col, false, 3, crate::app::NUM_COLS - 3),
            (Axis::Col, true, 6, crate::app::NUM_COLS - 6),
        ] {
            app.show_insert_count(axis, after, cx);
            let d = app.count_dialog.as_mut().unwrap();
            assert_eq!(d.at, at);
            assert_eq!(d.max, capacity);
            for invalid in [String::new(), "0".into(), "-1".into(), (capacity + 1).to_string(), "9999999999999999999999999".into()] {
                d.input = invalid;
                assert!(d.count().is_none());
            }
            d.input = capacity.to_string();
            assert_eq!(d.count(), Some(capacity));
            d.input = "1".into();
            assert_eq!(d.count(), Some(1));
            app.close_count_dialog(cx);
        }
    }).unwrap();
}

#[gpui::test]
fn header_insert_menu_opens_count_picker_and_button_inserts(cx: &mut TestAppContext) {
    use gpui::{point, px, Modifiers};
    cx.update(|cx| {
        crate::settings::init_settings_store(cx);
        crate::load_embedded_fonts(cx);
        cx.set_global(crate::session::SessionManager::new());
        cx.set_global(crate::window_registry::WindowRegistry::new());
    });
    let (view, vcx) = cx.add_window_view(Spreadsheet::new);
    view.update(vcx, |app, cx| {
        app.select_row(2, false, cx);
        app.show_context_menu(crate::app::ContextMenuKind::RowHeader, point(px(80.), px(200.)), cx);
    });
    vcx.run_until_parked();
    let item = vcx.debug_bounds("ctx-insert-row-below").unwrap().center();
    vcx.simulate_click(item, Modifiers::none());
    vcx.run_until_parked();
    let bounds = vcx.debug_bounds("count-popover").unwrap();
    assert_eq!(bounds.origin, point(px(80.), px(200.)));
    assert!(bounds.size.width <= px(260.) && bounds.size.height <= px(120.));
    view.update(vcx, |app, cx| {
        let dialog = app.count_dialog.as_mut().expect("row insert dialog");
        assert_eq!(dialog.at, 3);
        dialog.input = "2".into();
        cx.notify();
    });
    vcx.run_until_parked();
    let confirm = vcx.debug_bounds("count-confirm").unwrap().center();
    vcx.simulate_click(confirm, Modifiers::none());
    vcx.run_until_parked();
    view.read_with(vcx, |app, _| {
        assert!(app.count_dialog.is_none());
        assert_eq!(app.history.undo_count(), 1);
    });
    view.update(vcx, |app, cx| app.show_insert_count(visigrid_engine::structural::Axis::Row, false, cx));
    vcx.run_until_parked();
    vcx.simulate_click(point(px(500.), px(400.)), Modifiers::none());
    view.read_with(vcx, |app, _| {
        assert!(app.count_dialog.is_none(), "clicking back on the sheet dismisses the popover");
        assert_eq!(app.history.undo_count(), 1, "dismissal does not insert");
    });
}
