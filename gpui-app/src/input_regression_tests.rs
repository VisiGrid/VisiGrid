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
        app.view_state.set_selection_end(Some((1, 0)));
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
        app.view_state.set_range((0, 1), (10, 2));
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
fn selected_rows_resize_together_and_undo_restores_each_original(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let id = app.cached_sheet_id();
        app.view_state.selected = (4, 0);
        app.view_state.selection_end = Some((1, 1));
        app.view_state.additional_selections.push(((3, 4), Some((5, 4))));
        app.view_state.additional_selections.push(((7, 4), None));
        app.sheet_row_heights_mut().extend([(1, 30.), (4, 62.), (6, 37.)]);
        app.hidden_rows.entry(id).or_default().insert(2);
        app.row_view = visigrid_engine::filter::RowView::new(10);
        app.row_view.apply_filter((0..10).map(|r| r != 3).collect());
        let originals = app.row_heights.clone();

        app.begin_row_resize(1, 100., cx);
        app.update_row_resize(114., cx);
        app.update_row_resize(120., cx);
        for row in [1, 2, 3, 4, 5, 7] {
            assert_eq!(app.row_height(row), 50.);
        }
        assert_eq!(app.displayed_row_height(2), 0.);
        assert_eq!(app.displayed_row_height(3), 0.);
        app.finish_row_resize(cx);
        assert_eq!(app.row_heights[&id].len(), 7);
        assert_eq!(app.row_height(6), 37.);
        assert_eq!(app.history.undo_count(), 1);
        app.undo(cx);
        assert_eq!(app.row_heights, originals);
        app.redo(cx);
        for row in [1, 2, 3, 4, 5, 7] {
            assert_eq!(app.row_heights[&id][&row], 50.);
        }
    }).unwrap();
}

#[gpui::test]
fn unselected_row_resizes_alone(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.selected = (1, 0);
        app.view_state.selection_end = Some((4, 2));
        let height = app.row_height(6);
        app.begin_row_resize(6, 80., cx);
        app.update_row_resize(88., cx);
        app.finish_row_resize(cx);
        let heights = &app.row_heights[&app.cached_sheet_id()];
        assert_eq!(heights.len(), 1);
        assert_eq!(heights[&6], height + 8.);
        assert_eq!(app.history.undo_count(), 1);
    }).unwrap();
}

#[gpui::test]
fn zero_distance_row_resize_preserves_default_and_explicit_heights(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.selection_end = Some((2, 0));
        app.sheet_row_heights_mut().insert(1, 48.);
        let before = app.row_heights.clone();
        app.begin_row_resize(0, 100., cx);
        app.finish_row_resize(cx);
        app.begin_row_resize(1, 100., cx);
        app.update_row_resize(130., cx);
        app.update_row_resize(100., cx);
        assert_eq!(app.row_height(1), 48.);
        assert_eq!(app.row_height(0), app.metrics.default_cell_sizes.row_height);
        app.finish_row_resize(cx);
        assert_eq!(app.row_heights, before);
        assert_eq!(app.history.undo_count(), 0);
    }).unwrap();
}

#[gpui::test]
fn row_resize_at_double_zoom_tracks_the_pointer(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.metrics.zoom = 2.;
        app.sheet_row_heights_mut().insert(1, 30.);
        app.begin_row_resize(1, 100., cx);
        app.update_row_resize(140., cx);
        assert_eq!(app.displayed_row_height(1), 100.);
        app.finish_row_resize(cx);
        assert_eq!(app.row_height(1), 50.);
    }).unwrap();
}

#[gpui::test]
fn active_table_view_keeps_its_single_row_resize(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.table_view_installed = true;
        app.view_state.selection_end = Some((4, 2));
        let height = app.row_height(1);
        app.begin_row_resize(1, 80., cx);
        app.update_row_resize(88., cx);
        app.finish_row_resize(cx);
        let heights = &app.row_heights[&app.cached_sheet_id()];
        assert_eq!(heights.len(), 1);
        assert_eq!(heights[&1], height + 8.);
    }).unwrap();
}

#[gpui::test]
fn whole_sheet_row_resize_previews_even_the_last_row_without_materializing_heights(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.select_all(cx);
        let before = app.row_heights.clone();
        let start = app.row_height(0);
        app.begin_row_resize(0, 100., cx);
        for delta in [10., 15., 20.] {
            app.update_row_resize(100. + delta, cx);
            assert_eq!(app.row_height(crate::app::NUM_ROWS - 1), start + delta);
            assert_eq!(app.row_heights, before);
        }
        app.update_row_resize(100., cx);
        app.finish_row_resize(cx);
        assert_eq!(app.row_heights, before);
        assert_eq!(app.history.undo_count(), 0);
    }).unwrap();
}

#[gpui::test]
fn insert_count_starts_selected_and_accepts_default_or_replacement(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| { wb.set_cell_value_tracked(0, 0, 0, "original"); });
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "ctrl-shift-enter");
    view.update(cx, |app, window, _| {
        assert!(app.focus_handle.is_focused(window));
        let dialog = app.count_dialog.as_ref().unwrap();
        assert_eq!(dialog.input, "1");
        assert!(dialog.input_selected);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "enter");
    view.update(cx, |app, _, cx| {
        assert!(app.count_dialog.is_none());
        assert_eq!(app.sheet(cx).get_raw(1, 0), "original");
        assert_eq!(app.history.undo_count(), 1);
        app.undo(cx);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "ctrl-shift-enter 4");
    view.update(cx, |app, _, _| {
        assert_eq!(app.count_dialog.as_ref().unwrap().input, "4");
        assert!(!app.count_dialog.as_ref().unwrap().input_selected);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "ctrl-a 5 enter");
    view.update(cx, |app, _, cx| {
        assert_eq!(app.sheet(cx).get_raw(5, 0), "original");
        assert_eq!(app.history.undo_count(), 1);
        app.show_insert_count(Axis::Col, false, cx);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "backspace enter");
    view.update(cx, |app, _, _| {
        assert!(app.count_dialog.as_ref().unwrap().input.is_empty());
        assert_eq!(app.history.undo_count(), 1);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "escape");
    view.update(cx, |app, _, cx| app.select_col(2, false, cx)).unwrap();
    cx.simulate_keystrokes(view.into(), "ctrl-shift-enter");
    view.update(cx, |app, _, _| {
        let dialog = app.count_dialog.as_ref().unwrap();
        assert!(matches!(dialog.axis, Axis::Col));
        assert_eq!(dialog.at, 2);
        assert_eq!(dialog.input, "1");
        assert!(dialog.input_selected);
    }).unwrap();
}

#[gpui::test]
fn insert_count_defaults_to_selected_rows_or_columns_and_confirms_once(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| { wb.set_cell_value_tracked(0, 10, 10, "marker"); });
    }).unwrap();
    for (axis, after) in [(Axis::Row, false), (Axis::Row, true), (Axis::Col, false), (Axis::Col, true)] {
        view.update(cx, |app, _, cx| {
            match axis {
                Axis::Row => { app.select_row(2, false, cx); app.select_row(4, true, cx); }
                Axis::Col => { app.select_col(2, false, cx); app.select_col(4, true, cx); }
            }
            app.show_insert_count(axis, after, cx);
            let dialog = app.count_dialog.as_ref().unwrap();
            assert_eq!(dialog.input, "3");
            assert!(dialog.input_selected);
        }).unwrap();
        cx.simulate_keystrokes(view.into(), "enter");
        view.update(cx, |app, _, cx| {
            let (row, col) = match axis { Axis::Row => (13, 10), Axis::Col => (10, 13) };
            assert_eq!(app.sheet(cx).get_raw(row, col), "marker");
            assert!(app.count_dialog.is_none());
            assert_eq!(app.history.undo_count(), 1);
            app.undo(cx);
            assert_eq!(app.sheet(cx).get_raw(10, 10), "marker");
        }).unwrap();
    }
    view.update(cx, |app, _, cx| {
        app.select_row(crate::app::NUM_ROWS - 4, false, cx);
        app.select_row(crate::app::NUM_ROWS - 2, true, cx);
        app.show_insert_count(Axis::Row, true, cx);
        let dialog = app.count_dialog.as_ref().unwrap();
        assert_eq!(dialog.input, "3", "preserve the selection even when capacity is smaller");
        assert_eq!(dialog.count(), None, "remaining-capacity validation still applies");
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "enter");
    view.update(cx, |app, _, _| {
        assert!(app.count_dialog.is_some());
        assert_eq!(app.history.undo_count(), 0);
    }).unwrap();
}

#[gpui::test]
fn insert_count_dialog_validates_keys_and_inserts_in_one_undo_step(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| { wb.set_cell_value_tracked(0, 1, 0, "original"); });
        app.view_state.set_single_cell((1, 0));
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
        app.view_state.set_range((4, 3), (6, 5));
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
            assert_eq!(d.input, "3");
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
        app.select_row(4, true, cx);
        app.show_context_menu(crate::app::ContextMenuKind::RowHeader, point(px(80.), px(200.)), cx);
    });
    vcx.run_until_parked();
    let item = vcx.debug_bounds("ctx-insert-row-below").unwrap().center();
    vcx.simulate_click(item, Modifiers::none());
    vcx.run_until_parked();
    let bounds = vcx.debug_bounds("count-popover").unwrap();
    assert_eq!(bounds.origin, point(px(80.), px(200.)));
    assert!(bounds.size.width <= px(208.) && bounds.size.height <= px(72.));
    view.read_with(vcx, |app, _| {
        let dialog = app.count_dialog.as_ref().expect("row insert dialog");
        assert_eq!(dialog.at, 5);
        assert_eq!(dialog.input, "3");
        assert!(dialog.input_selected);
    });
    vcx.simulate_keystrokes("2");
    view.read_with(vcx, |app, _| {
        assert_eq!(app.count_dialog.as_ref().unwrap().input, "2");
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

fn primary_key(cx: &mut TestAppContext, view: WindowHandle<Spreadsheet>, key: &str) {
    let modifier = crate::keybindings::primary_mod_label(crate::settings::ModifierStyle::Platform);
    cx.simulate_keystrokes(
        view.into(),
        &format!("{}-{key}", modifier.to_ascii_lowercase()),
    );
}

fn setup_paste(cx: &mut TestAppContext) -> WindowHandle<Spreadsheet> {
    let view = setup(cx);
    // Keep the copy border's wall-clock animation out of the GPUI test scheduler.
    view.update(cx, |_, window, _| window.activate_window())
        .unwrap();
    cx.run_until_parked();
    gpui::VisualTestContext::from_window(view.into(), cx).deactivate_window();
    view
}

#[gpui::test]
fn range_paste_uses_top_left_in_either_drag_direction_and_undo_restores_cells(
    cx: &mut TestAppContext,
) {
    let view = setup_paste(cx);
    cx.update(|cx| {
        cx.global_mut::<crate::settings::SettingsStore>()
            .user_settings_mut()
            .editing
            .paste_values_by_default = crate::settings::Setting::Value(false);
    });
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| {
            for row in 2..=6 {
                wb.set_cell_value_tracked(0, row, 1, &format!("source {row}"));
                wb.set_cell_value_tracked(0, row, 5, &format!("original {row}"));
            }
        });
        app.view_state.set_range((2, 1), (6, 1));
    })
    .unwrap();
    primary_key(cx, view, "c");

    // A single-cell destination expands to the pasted rectangle.
    view.update(cx, |app, _, _| {
        app.view_state.set_single_cell((2, 5));
    })
    .unwrap();
    primary_key(cx, view, "v");
    view.update(cx, |app, _, cx| {
        assert_eq!(app.sheet(cx).get_raw(6, 5), "source 6");
        assert_eq!(app.view_state.active_cell(), (2, 5));
        assert_eq!(app.view_state.selection_end(), Some((6, 5)));
        app.undo(cx);
    })
    .unwrap();

    for (anchor, end) in [(2, 6), (6, 2)] {
        view.update(cx, |app, _, _| {
            app.view_state.set_range((anchor, 5), (end, 5));
        })
        .unwrap();
        primary_key(cx, view, "v");
        view.update(cx, |app, _, cx| {
            for row in 2..=6 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), format!("source {row}"));
            }
            for row in 7..=10 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), "");
            }
            assert_eq!(app.view_state.active_cell(), (2, 5));
            assert_eq!(app.view_state.selection_end(), Some((6, 5)));
            assert_eq!(app.history.undo_count(), 1);
        })
        .unwrap();
        primary_key(cx, view, "z");
        view.update(cx, |app, _, cx| {
            for row in 2..=6 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), format!("original {row}"));
            }
        })
        .unwrap();
    }
}

#[gpui::test]
fn paste_special_uses_top_left_and_selects_the_source_sized_rectangle(cx: &mut TestAppContext) {
    use visigrid_engine::cell::CellFormat;

    let view = setup_paste(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.set_cell_value_tracked(0, 0, 0, "4");
            wb.set_cell_value_tracked(0, 0, 1, "=A1+1");
            wb.active_sheet_mut().set_format(
                0,
                1,
                CellFormat {
                    bold: true,
                    ..Default::default()
                },
            );
        });
        app.view_state.set_range((0, 0), (0, 1));
        app.copy(cx);
        // Exercise the internal-only paths with the clipboard ID they require.
        let copied = app.internal_clipboard.as_ref().unwrap();
        cx.write_to_clipboard(gpui::ClipboardItem::new_string_with_json_metadata(
            copied.raw_tsv.clone(),
            copied.id.to_string(),
        ));

        // Select H4:F3. Each paste should select exactly F3:G3.
        let select_target = |app: &mut Spreadsheet| {
            app.view_state.set_range((3, 7), (2, 5));
        };
        let check_selection = |app: &Spreadsheet| {
            assert_eq!(app.view_state.active_cell(), (2, 5));
            assert_eq!(app.view_state.selection_end(), Some((2, 6)));
        };
        select_target(app);
        app.paste_values(cx);
        assert_eq!(app.sheet(cx).get_raw(2, 6), "5");
        check_selection(app);
        app.undo(cx);

        select_target(app);
        app.paste_formulas(cx);
        assert_eq!(app.sheet(cx).get_raw(2, 6), "=F3+1");
        check_selection(app);
        app.undo(cx);

        select_target(app);
        app.paste_formats(cx);
        assert!(app.sheet(cx).get_format(2, 6).bold);
        check_selection(app);
    })
    .unwrap();
}

#[gpui::test]
fn one_cell_formula_fallback_fills_selection_while_values_and_formats_use_top_left(cx: &mut TestAppContext) {
    let view = setup_paste(cx);
    view.update(cx, |app, _, cx| {
        // Use the normal contents paste when Paste Formulas falls back.
        cx.global_mut::<crate::settings::SettingsStore>()
            .user_settings_mut()
            .editing
            .paste_values_by_default = crate::settings::Setting::Value(false);
        app.wb_mut(cx, |wb| {
            wb.set_cell_value_tracked(0, 0, 0, "=2+3");
            wb.active_sheet_mut().toggle_bold(0, 0);
            for row in 2..=4 {
                wb.set_cell_value_tracked(0, row, 5, "original");
            }
        });
        app.view_state.set_single_cell((0, 0));
        app.copy(cx);
        let select_target = |app: &mut Spreadsheet| {
            app.view_state.set_range((4, 5), (2, 5));
            app.view_state.additional_selections = vec![((3, 8), None)];
        };
        let check_target = |app: &Spreadsheet, cx: &gpui::App| {
            assert_eq!(app.view_state.active_cell(), (2, 5));
            assert_eq!(app.view_state.selection_end(), None);
            assert!(app.view_state.additional_selections.is_empty());
            for row in 3..=4 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), "original");
                assert!(!app.sheet(cx).get_format(row, 5).bold);
            }
            assert_eq!(app.sheet(cx).get_raw(3, 8), "");
            assert!(!app.sheet(cx).get_format(3, 8).bold);
        };

        let check_fill = |app: &Spreadsheet, cx: &gpui::App, formula: &str| {
            for row in 2..=4 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), formula);
            }
            assert_eq!(app.sheet(cx).get_raw(3, 8), formula);
            assert_eq!(app.view_state.active_cell(), (4, 5));
            assert_eq!(app.view_state.selection_end(), Some((2, 5)));
            assert_eq!(app.view_state.additional_selections.as_slice(), &[((3, 8), None)]);
        };

        // Use the clipboard exactly as Copy produced it.
        select_target(app);
        app.paste_formulas(cx);
        check_fill(app, cx, "=2+3");
        app.undo(cx);

        select_target(app);
        app.paste_values(cx);
        assert_eq!(app.sheet(cx).get_raw(2, 5), "5");
        check_target(app, cx);
        app.undo(cx);

        // Formats requires matching internal clipboard metadata.
        let copied = app.internal_clipboard.as_ref().unwrap();
        cx.write_to_clipboard(gpui::ClipboardItem::new_string_with_json_metadata(
            copied.raw_tsv.clone(),
            copied.id.to_string(),
        ));

        select_target(app);
        app.paste_formats(cx);
        assert!(app.sheet(cx).get_format(2, 5).bold);
        assert_eq!(app.sheet(cx).get_raw(2, 5), "original");
        check_target(app, cx);
        app.undo(cx);

        select_target(app);
        cx.write_to_clipboard(gpui::ClipboardItem::new_string("=8+1".into()));
        app.paste_formulas(cx);
        check_fill(app, cx, "=8+1");
    })
    .unwrap();
}

#[gpui::test]
fn whole_column_jump_right_stops_at_the_end_of_first_row_data(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| {
            for col in 1..=3 {
                wb.set_cell_value_tracked(0, 0, col, "1");
            }
            wb.set_cell_value_tracked(0, 0, 5, "after gap");
        });
        app.select_col(1, false, cx);
    })
    .unwrap();
    primary_key(cx, view, "shift-right");
    view.update(cx, |app, _, _| {
        assert_eq!(app.view_state.active_cell(), (0, 1));
        assert_eq!(
            app.view_state.selection_end(),
            Some((crate::app::NUM_ROWS - 1, 3))
        );
    })
    .unwrap();
}

#[gpui::test]
fn one_cell_contents_and_all_preserve_the_selection_they_fill(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        cx.global_mut::<crate::settings::SettingsStore>()
            .user_settings_mut()
            .editing
            .paste_values_by_default = crate::settings::Setting::Value(false);
        cx.write_to_clipboard(gpui::ClipboardItem::new_string("fill".into()));
        app.view_state.set_range((4, 5), (2, 5));
        app.view_state.additional_selections.push(((3, 8), None));

        for paste in [Spreadsheet::paste, Spreadsheet::paste_all] {
            paste(app, cx);
            for row in 2..=4 {
                assert_eq!(app.sheet(cx).get_raw(row, 5), "fill");
            }
            assert_eq!(app.sheet(cx).get_raw(3, 8), "fill");
            assert_eq!(app.view_state.active_cell(), (4, 5));
            assert_eq!(app.view_state.selection_end(), Some((2, 5)));
            assert_eq!(
                app.view_state.additional_selections.as_slice(),
                &[((3, 8), None)]
            );
            app.undo(cx);
        }
    })
    .unwrap();
}

fn check_perpendicular_jump(
    cx: &mut TestAppContext,
    gap: (usize, usize),
    first: &str,
    first_end: (usize, usize),
    second: &str,
) {
    let view = setup(cx);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.wb_mut(cx, |wb| {
            for row in 0..9 {
                for col in 0..6 {
                    wb.set_cell_value_tracked(0, row, col, "1");
                }
            }
            wb.set_cell_value_tracked(0, gap.0, gap.1, "");
        });
        app.view_state.set_single_cell((8, 5));
    })
    .unwrap();
    primary_key(cx, view, first);
    view.update(cx, |app, _, _| {
        assert_eq!(app.view_state.active_cell(), (8, 5));
        assert_eq!(app.view_state.selection_end(), Some(first_end));
    })
    .unwrap();
    primary_key(cx, view, second);
    view.update(cx, |app, _, _| {
        assert_eq!(app.view_state.active_cell(), (8, 5));
        assert_eq!(app.view_state.selection_end(), Some((0, 0)));
    })
    .unwrap();
}

#[gpui::test]
fn horizontal_jump_extends_along_the_anchor_row(cx: &mut TestAppContext) {
    check_perpendicular_jump(cx, (0, 3), "shift-up", (0, 5), "shift-left");
}

#[gpui::test]
fn vertical_jump_extends_along_the_anchor_column(cx: &mut TestAppContext) {
    check_perpendicular_jump(cx, (3, 0), "shift-left", (8, 0), "shift-up");
}

#[gpui::test]
fn reversed_header_paste_checks_and_renames_from_the_leftmost_header(cx: &mut TestAppContext) {
    use visigrid_engine::table::TableRange;

    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            for (col, name) in ["First", "Second", "Third"].iter().enumerate() {
                wb.set_cell_value_tracked(0, 0, col, name);
            }
            let range = TableRange {
                start_row: 0,
                end_row: 2,
                start_col: 0,
                end_col: 2,
            };
            wb.create_table(wb.active_sheet().id, range, "Orders")
                .unwrap();
        });
        app.view_state.set_range((0, 2), (0, 0));
        cx.write_to_clipboard(gpui::ClipboardItem::new_string("Item\tQty".into()));
        app.paste(cx);
        assert_eq!(app.sheet(cx).get_raw(0, 0), "Item");
        assert_eq!(app.sheet(cx).get_raw(0, 1), "Qty");
        assert_eq!(app.sheet(cx).get_raw(0, 2), "Third");
        assert_eq!(app.view_state.active_cell(), (0, 0));
        assert_eq!(app.view_state.selection_end(), Some((0, 1)));
    })
    .unwrap();
}

#[gpui::test]
fn enter_shift_enter_tab_shift_tab_cycle_2x3_range_and_wrap(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let range = ((1, 1), (2, 3));
        app.view_state.set_range((1, 1), (2, 3));

        // Enter: down then next column, wrapping to top-left
        let enter_expected = [(2, 1), (1, 2), (2, 2), (1, 3), (2, 3), (1, 1)];
        for expected in enter_expected {
            app.confirm_edit_enter(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Enter step");
            assert_eq!(app.selection_range(), range, "range preserved on Enter");
        }

        // Shift+Enter: reverse order, wrapping to bottom-right
        let shift_enter_expected = [(2, 3), (1, 3), (2, 2), (1, 2), (2, 1), (1, 1)];
        for expected in shift_enter_expected {
            app.confirm_edit_up_enter(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Shift+Enter step");
            assert_eq!(app.selection_range(), range, "range preserved on Shift+Enter");
        }

        // Tab: right then next row, wrapping to top-left
        let tab_expected = [(1, 2), (1, 3), (2, 1), (2, 2), (2, 3), (1, 1)];
        for expected in tab_expected {
            app.confirm_edit_and_tab_right(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Tab step");
            assert_eq!(app.selection_range(), range, "range preserved on Tab");
        }

        // Shift+Tab: reverse order, wrapping to bottom-right
        let shift_tab_expected = [(2, 3), (2, 2), (2, 1), (1, 3), (1, 2), (1, 1)];
        for expected in shift_tab_expected {
            app.confirm_edit_and_tab_left(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Shift+Tab step");
            assert_eq!(app.selection_range(), range, "range preserved on Shift+Tab");
        }
    })
    .unwrap();
}

#[gpui::test]
fn typing_then_enter_in_range_writes_only_active_cell_and_advances(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((1, 1), (2, 2));

        app.start_edit(cx);
        app.edit_value = "alpha".into();
        app.confirm_edit_enter(cx);
        assert_eq!(app.sheet(cx).get_raw(1, 1), "alpha");
        assert_eq!(app.sheet(cx).get_raw(2, 1), "");
        assert_eq!(app.sheet(cx).get_raw(1, 2), "");
        assert_eq!(app.sheet(cx).get_raw(2, 2), "");
        assert_eq!(app.view_state.active_cell(), (2, 1));
        assert_eq!(app.selection_range(), ((1, 1), (2, 2)));

        app.start_edit(cx);
        app.edit_value = "beta".into();
        app.confirm_edit_enter(cx);
        assert_eq!(app.sheet(cx).get_raw(2, 1), "beta");
        assert_eq!(app.sheet(cx).get_raw(1, 2), "");
        assert_eq!(app.view_state.active_cell(), (1, 2));
        assert_eq!(app.selection_range(), ((1, 1), (2, 2)));
    })
    .unwrap();
}

#[gpui::test]
fn cycle_skips_hidden_row_and_hidden_column(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let sheet_id = app.sheet(cx).id;
        app.hidden_rows.entry(sheet_id).or_default().insert(2);
        app.hidden_cols.entry(sheet_id).or_default().insert(2);

        app.view_state.set_range((1, 1), (3, 3));

        // Visible cells in (1,1)..=(3,3) are (1,1), (3,1), (1,3), (3,3)
        let enter_expected = [(3, 1), (1, 3), (3, 3), (1, 1)];
        for expected in enter_expected {
            app.confirm_edit_enter(cx);
            assert_eq!(app.view_state.active_cell(), expected);
            assert_eq!(app.selection_range(), ((1, 1), (3, 3)));
        }

        let tab_expected = [(1, 3), (3, 1), (3, 3), (1, 1)];
        for expected in tab_expected {
            app.confirm_edit_and_tab_right(cx);
            assert_eq!(app.view_state.active_cell(), expected);
            assert_eq!(app.selection_range(), ((1, 1), (3, 3)));
        }
    })
    .unwrap();
}

#[gpui::test]
fn cycle_treats_merged_area_as_one_stop_at_origin(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.active_sheet_mut()
                .add_merge(visigrid_engine::sheet::MergedRegion::new(1, 1, 2, 2))
                .expect("merge 2x2");
        });
        app.view_state.set_range((1, 1), (3, 3));

        // Enter order: (1,1) [merge origin] -> (3,1) -> (3,2) -> (1,3) -> (2,3) -> (3,3) -> (1,1)
        let enter_expected = [(3, 1), (3, 2), (1, 3), (2, 3), (3, 3), (1, 1)];
        for expected in enter_expected {
            app.confirm_edit_enter(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Enter with merge");
            assert_eq!(app.selection_range(), ((1, 1), (3, 3)));
        }

        // Tab order: (1,1) [merge origin] -> (1,3) -> (2,3) -> (3,1) -> (3,2) -> (3,3) -> (1,1)
        let tab_expected = [(1, 3), (2, 3), (3, 1), (3, 2), (3, 3), (1, 1)];
        for expected in tab_expected {
            app.confirm_edit_and_tab_right(cx);
            assert_eq!(app.view_state.active_cell(), expected, "Tab with merge");
            assert_eq!(app.selection_range(), ((1, 1), (3, 3)));
        }
    })
    .unwrap();
}

#[gpui::test]
fn single_cell_enter_and_tab_preserve_tab_chain_return(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_single_cell((1, 1));

        app.confirm_edit_and_tab_right(cx);
        assert_eq!(app.view_state.active_cell(), (1, 2));
        app.confirm_edit_and_tab_right(cx);
        assert_eq!(app.view_state.active_cell(), (1, 3));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (2, 1), "Enter returns to tab-chain origin column");
        assert_eq!(app.view_state.selection_end(), None);
    })
    .unwrap();
}

#[gpui::test]
fn additional_selections_keep_existing_collapse_behaviour(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((1, 1), (2, 2));
        app.view_state.additional_selections.push(((5, 5), None));

        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));
        assert_eq!(app.view_state.selection_end(), None);
        assert!(app.view_state.additional_selections.is_empty());
    })
    .unwrap();
}

#[gpui::test]
fn arrow_key_after_cycling_collapses_selection(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((1, 1), (2, 3));

        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));
        assert_eq!(app.selection_range(), ((1, 1), (2, 3)));

        app.move_selection(0, 1, cx);
        assert_eq!(app.view_state.active_cell(), (2, 2));
        assert_eq!(app.view_state.selection_end(), None);
        assert!(!app.view_state.is_cycling_in_range());
        assert_eq!(app.selection_range(), ((2, 2), (2, 2)));
    })
    .unwrap();
}

#[gpui::test]
fn shift_arrow_and_shift_click_after_cycling_extend_from_active_cell(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((1, 1), (3, 3));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));

        // Shift+Right extends from active cell (2, 1) -> (2, 1)..=(2, 2)
        app.extend_selection(0, 1, cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));
        assert!(!app.view_state.is_cycling_in_range());
        assert_eq!(app.selection_range(), ((2, 1), (2, 2)));

        // Re-select range, cycle to (2, 1), then Shift+Click (2, 3) -> (2, 1)..=(2, 3)
        app.view_state.set_range((1, 1), (3, 3));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));

        app.select_cell(2, 3, true, cx);
        assert_eq!(app.view_state.active_cell(), (2, 1));
        assert!(!app.view_state.is_cycling_in_range());
        assert_eq!(app.selection_range(), ((2, 1), (2, 3)));
    })
    .unwrap();
}

#[gpui::test]
fn paste_after_cycling_still_starts_at_range_top_left(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((1, 1), (2, 3));

        // Cycle Enter twice: (1,1) -> (2,1) -> (1,2)
        app.confirm_edit_enter(cx);
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 2));
        assert_eq!(app.selection_range(), ((1, 1), (2, 3)));

        cx.write_to_clipboard(gpui::ClipboardItem::new_string("X\tY".into()));
        app.paste(cx);
        assert_eq!(app.sheet(cx).get_raw(1, 1), "X");
        assert_eq!(app.sheet(cx).get_raw(1, 2), "Y");
        assert_eq!(app.sheet(cx).get_raw(1, 3), "");
    })
    .unwrap();
}

#[gpui::test]
fn cycling_in_whole_column_selection_moves_one_row_and_is_instant(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.select_col(2, false, cx);
        assert_eq!(app.view_state.active_cell(), (0, 2));
        let expected_range = app.selection_range();

        let start = std::time::Instant::now();
        app.confirm_edit_enter(cx);
        let elapsed = start.elapsed();

        assert_eq!(app.view_state.active_cell(), (1, 2));
        assert_eq!(app.selection_range(), expected_range);
        assert!(
            elapsed < std::time::Duration::from_millis(100),
            "whole-column cycling took {:?} (expected O(log N) / instant)",
            elapsed
        );
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_named_range_jump(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            let nr = visigrid_engine::named_range::NamedRange::range("Target", 0, 9, 23, 11, 24);
            wb.named_ranges_mut().set(nr).unwrap();
        });
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));

        app.jump_to_named_range("Target", cx);
        assert_eq!(app.view_state.active_cell(), (9, 23));
        assert_eq!(app.selection_range(), ((9, 23), (11, 24)));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_goto_merged_cell(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.active_sheet_mut()
                .add_merge(visigrid_engine::sheet::MergedRegion::new(5, 5, 6, 6))
                .unwrap();
        });
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));

        app.goto_input = "F6".into();
        app.confirm_goto(cx);
        assert_eq!(app.view_state.active_cell(), (5, 5));
        assert_eq!(app.selection_range(), ((5, 5), (6, 6)));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_fill_down_and_fill_handle(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.set_cell_value_tracked(0, 0, 0, "10");
            wb.set_cell_value_tracked(0, 4, 4, "1");
            wb.set_cell_value_tracked(0, 5, 4, "2");
        });
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));

        // Fill Down keeps the primary selection cleanly
        app.fill_down(cx);
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));
        assert_eq!(app.sheet(cx).get_raw(1, 0), "10");

        // Fill handle drag from E5:E6 down to E8
        app.select_cell(4, 4, false, cx);
        app.select_cell(5, 4, true, cx);
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (5, 4));
        app.start_fill_drag(cx);
        app.continue_fill_drag(7, 4, cx);
        app.end_fill_drag(false, cx);
        assert_eq!(app.selection_range(), ((4, 4), (7, 4)));
        assert_eq!(app.view_state.active_cell(), (4, 4));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_insert_and_delete_rows_and_cols(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));

        app.delete_rows(5, 2, cx);
        assert_eq!(app.view_state.active_cell(), (5, 0));
        assert_eq!(app.selection_range(), ((5, 0), (6, crate::app::NUM_COLS - 1)));

        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        app.delete_cols(5, 2, cx);
        assert_eq!(app.view_state.active_cell(), (0, 5));
        assert_eq!(app.selection_range(), ((0, 5), (crate::app::NUM_ROWS - 1, 6)));

        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        app.insert_rows(4, 1, cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
        app.insert_cols(4, 1, cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_select_row_col_all(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);

        app.select_row(5, false, cx);
        assert_eq!(app.view_state.active_cell(), (5, 0));
        assert_eq!(app.selection_range(), ((5, 0), (5, crate::app::NUM_COLS - 1)));

        app.confirm_edit_and_tab_right(cx);
        app.select_col(4, false, cx);
        assert_eq!(app.view_state.active_cell(), (0, 4));
        assert_eq!(app.selection_range(), ((0, 4), (crate::app::NUM_ROWS - 1, 4)));

        app.confirm_edit_enter(cx);
        app.select_all(cx);
        assert_eq!(app.view_state.active_cell(), (0, 0));
        assert_eq!(app.selection_range(), ((0, 0), (crate::app::NUM_ROWS - 1, crate::app::NUM_COLS - 1)));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_find_next(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.set_cell_value_tracked(0, 4, 4, "needle");
        });
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);

        app.show_find(cx);
        for c in "needle".chars() {
            app.find_insert_char(c, cx);
        }
        app.find_next(cx);
        assert_eq!(app.view_state.active_cell(), (4, 4));
        assert_eq!(app.selection_range(), ((4, 4), (4, 4)));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_undo_redo_selection(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));

        app.delete_rows(5, 2, cx);
        assert_eq!(app.selection_range(), ((5, 0), (6, crate::app::NUM_COLS - 1)));

        app.undo(cx);
        assert_eq!(app.selection_range().0.0, app.view_state.active_cell().0);
        app.redo(cx);
        assert_eq!(app.selection_range().0.0, app.view_state.active_cell().0);
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_rewind_preview_enter_and_exit(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        // Create one history entry at (5, 5)
        app.select_cell(5, 5, false, cx);
        app.start_edit(cx);
        app.edit_value = "42".into();
        app.confirm_edit(cx);

        // Select A1:C3 and cycle to (1, 0)
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));

        // Enter rewind preview for entry 0 with a multi-cell highlight range, then exit
        app.selected_history_id = app.history.display_entries().first().map(|e| e.id);
        app.history_highlight_range = Some((0, 5, 5, 6, 6));
        app.enter_preview(cx).unwrap();
        assert_eq!(app.view_state.active_cell(), (5, 5));
        assert_eq!(app.selection_range(), ((5, 5), (6, 6)));
        app.exit_preview(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));
        assert!(app.view_state.is_cycling_in_range());
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_sheet_switch_and_split_pane(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.add_sheet();
        });
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));

        // Split pane inherits selection independently; moving in split pane doesn't corrupt main
        app.split_right(cx);
        assert_eq!(app.active_view_state().active_cell(), (1, 0));
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));
        app.move_selection(0, 1, cx);
        assert_eq!(app.selection_range(), ((1, 1), (1, 1)));
        app.close_split(cx);

        // Switching sheets resets selection cleanly
        app.goto_sheet(1, cx);
        assert_eq!(app.view_state.active_cell(), (0, 0));
        assert_eq!(app.selection_range(), ((0, 0), (0, 0)));
    })
    .unwrap();
}

#[gpui::test]
fn stale_anchor_after_cycling_session_save_and_restore(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.session_window_id = 1;
        app.view_state.set_range((0, 0), (2, 2));
        app.confirm_edit_enter(cx);
        assert_eq!(app.view_state.active_cell(), (1, 0));
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));

        let snap = app.snapshot_cached(cx);
        app.select_cell(8, 8, false, cx);
        app.apply(&snap, cx);
        assert_eq!(app.selection_range(), ((0, 0), (2, 2)));
        assert_eq!(app.view_state.active_cell(), (0, 0));
    })
    .unwrap();
}
