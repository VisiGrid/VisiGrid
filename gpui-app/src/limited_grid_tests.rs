//! Headless acceptance checks exercise the actual Spreadsheet, not a second model.
use crate::{
    app::{Spreadsheet, NUM_COLS, NUM_ROWS},
    limited_grid::{content_bounds, GridExtent},
    settings::Setting,
};
use gpui::{TestAppContext, WindowHandle};

fn init(cx: &mut TestAppContext) {
    cx.update(|cx| {
        crate::settings::init_settings_store(cx);
        crate::load_embedded_fonts(cx);
        cx.set_global(crate::session::SessionManager::new());
        cx.set_global(crate::window_registry::WindowRegistry::new());
        crate::keybindings::register(cx, crate::settings::ModifierStyle::Platform);
    });
}
fn setup(cx: &mut TestAppContext) -> WindowHandle<Spreadsheet> {
    init(cx);
    let view = cx.add_window(Spreadsheet::new);
    view.update(cx, |app, window, cx| {
        window.focus(&app.focus_handle, cx);
        app.doc_settings.display.limited_grid = Setting::Value(true);
        app.sync_limited_grid(cx);
    })
    .unwrap();
    view
}
fn extent(app: &mut Spreadsheet, cx: &gpui::App) -> (usize, usize) {
    app.sync_limited_grid(cx);
    (app.grid_rows(), app.grid_cols())
}

#[gpui::test]
fn split_panes_keep_visible_grid_bodies_and_working_strips(cx: &mut TestAppContext) {
    use gpui::{px, Modifiers};
    init(cx);
    let (view, vcx) = cx.add_window_view(Spreadsheet::new);
    view.update(vcx, |app, cx| {
        app.doc_settings.display.limited_grid = Setting::Value(true);
        app.set_zoom(0.5, cx);
        app.view_state.selected = (1, 1);
        app.freeze_panes(cx);
        app.split_right(cx);
        app.sync_limited_grid(cx);
        cx.notify();
    });
    vcx.run_until_parked();
    for (body_selector, strip_selector) in [
        ("limited-grid-left", "grid-add-row-left"),
        ("limited-grid-right", "grid-add-row-right"),
    ] {
        let body = vcx.debug_bounds(body_selector).unwrap();
        let strip = vcx.debug_bounds(strip_selector).unwrap();
        assert!(body.size.height > px(300.), "split grid must fill its pane");
        assert!(body.contains(&strip.center()), "bottom strip must be inside the visible pane");
        vcx.simulate_click(strip.center(), Modifiers::none());
        vcx.run_until_parked();
    }
    view.update(vcx, |app, cx| {
        assert_eq!(extent(app, cx), (22, 11));
        assert_eq!(app.view_state.frozen_rows, 1);
        assert_eq!(app.view_state.frozen_cols, 1);
        assert_eq!(app.history.undo_count(), 1, "only freezing panes adds an undo step");
    });
}

#[gpui::test]
fn add_count_accepts_selected_default_and_replacement(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.show_add_count(Axis::Row, cx);
        let dialog = app.count_dialog.as_ref().unwrap();
        assert_eq!(dialog.input, "1");
        assert!(dialog.input_selected);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "enter");
    view.update(cx, |app, _, cx| {
        assert_eq!(extent(app, cx), (21, 11));
        app.show_add_count(Axis::Col, cx);
    }).unwrap();
    cx.simulate_keystrokes(view.into(), "4 enter");
    view.update(cx, |app, _, cx| {
        assert_eq!(extent(app, cx), (21, 15));
        assert_eq!(app.history.undo_count(), 0);
        assert!(app.count_dialog.is_none());
    }).unwrap();
}

#[gpui::test]
fn default_extent_strips_clicks_and_add_count(cx: &mut TestAppContext) {
    use gpui::Modifiers;
    init(cx);
    let (view, vcx) = cx.add_window_view(Spreadsheet::new);
    view.update(vcx, |app, cx| {
        app.doc_settings.display.limited_grid = Setting::Value(true);
        app.set_zoom(0.5, cx);
        app.sync_limited_grid(cx);
        assert_eq!(extent(app, cx), (20, 11));
        cx.notify();
    });
    vcx.run_until_parked();
    let bottom = vcx
        .debug_bounds("grid-add-row")
        .expect("bottom strip")
        .center();
    vcx.simulate_click(bottom, Modifiers::none());
    vcx.run_until_parked();
    let right = vcx
        .debug_bounds("grid-add-col")
        .expect("right strip")
        .center();
    vcx.simulate_click(right, Modifiers::none());
    vcx.run_until_parked();
    view.update(vcx, |app, cx| {
        assert_eq!(extent(app, cx), (21, 12));
        app.show_add_count(visigrid_engine::structural::Axis::Row, cx);
        app.count_dialog.as_mut().unwrap().input = "5".into();
        cx.notify();
    });
    vcx.run_until_parked();
    let confirm = vcx.debug_bounds("count-confirm").unwrap().center();
    vcx.simulate_click(confirm, Modifiers::none());
    vcx.run_until_parked();
    view.update(vcx, |app, cx| {
        assert_eq!(extent(app, cx), (26, 12));
        assert_eq!(app.history.undo_count(), 0, "extent changes are view state");
        let rows = app.pane_rows(&app.view_state);
        let cols = app.pane_cols(&app.view_state);
        assert_eq!(rows.hit(10000.), None);
        assert_eq!(cols.hit(10000.), None);
    });
    let blank = vcx.debug_bounds("grid-add-both").unwrap().center()
        + gpui::point(gpui::px(40.), gpui::px(40.));
    let selected = view.read_with(vcx, |app, _| app.view_state.selected);
    vcx.simulate_click(blank, Modifiers::none());
    vcx.simulate_event(gpui::ScrollEvent {
        position: blank,
        delta: gpui::ScrollDelta::Pixels(gpui::point(gpui::px(-100_000.), gpui::px(-100_000.))),
        modifiers: Default::default(),
        touch_phase: gpui::TouchPhase::Moved,
        source: gpui::ScrollSource::Pointer,
    });
    vcx.run_until_parked();
    view.read_with(vcx, |app, _| {
        assert_eq!(
            app.view_state.selected, selected,
            "blank canvas does not select a cell"
        );
        assert!(app.view_state.scroll_row < 26 && app.view_state.scroll_col < 12);
    });
    let corner = vcx.debug_bounds("grid-add-both").unwrap().center();
    vcx.simulate_click(corner, Modifiers::none());
    view.update(vcx, |app, cx| assert_eq!(extent(app, cx), (27, 13)));
}

#[gpui::test]
fn enter_on_content_grows_once_empty_row_stops_and_right_stops(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.view_state.selected = (19, 10);
        app.start_edit_clear(cx);
        app.edit_value = "ledger".into();
        app.edit_cursor = 6;
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "enter");
    view.update(cx, |app, _, cx| {
        assert_eq!(extent(app, cx), (21, 11));
        assert_eq!(app.view_state.selected, (20, 10));
        assert_eq!(app.sheet(cx).get_raw(19, 10), "ledger");
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "enter tab right down");
    view.update(cx, |app, _, cx| {
        app.flush_pending_nav_moves(cx);
        assert_eq!(extent(app, cx), (21, 11));
        assert_eq!(app.view_state.selected, (20, 10));
    })
    .unwrap();
}

#[gpui::test]
fn paste_and_spill_expand_without_limiting_the_model_and_undo_refreshes(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let text = (0..50)
            .map(|r| format!("{r}\tvalue\t3"))
            .collect::<Vec<_>>()
            .join("\n");
        cx.write_to_clipboard(gpui::ClipboardItem::new_string(text));
        app.paste(cx);
        assert_eq!(extent(app, cx), (50, 11));
        app.undo(cx);
        assert_eq!(extent(app, cx), (20, 11));
        app.redo(cx);
        assert_eq!(extent(app, cx), (50, 11));
        app.undo(cx);
        app.set_cell_value(0, 0, "=SEQUENCE(100)", cx);
        assert_eq!(extent(app, cx), (100, 11));
        assert_eq!(app.sheet(cx).get_display(99, 0), "100");
    })
    .unwrap();
}

#[gpui::test]
fn navigation_scroll_and_jumps_stay_inside(cx: &mut TestAppContext) {
    let view = setup(cx);
    cx.simulate_keystrokes(view.into(), "ctrl-down ctrl-end pagedown");
    view.update(cx, |app, _, cx| {
        app.flush_pending_nav_moves(cx);
        assert!(app.view_state.selected.0 < 20);
        app.scroll(i32::MAX, i32::MAX, cx);
        assert!(app.view_state.scroll_row < 20);
        assert!(app.view_state.scroll_col < 11);
        let before = app.view_state.selected;
        app.show_goto(cx);
        app.goto_input = "Z500".into();
        app.confirm_goto(cx);
        assert_eq!(app.view_state.selected, before);
        assert!(app
            .status_message
            .as_ref()
            .unwrap()
            .contains("Z500 is outside"));
        assert_eq!(extent(app, cx), (20, 11));
        app.name_box_input = "Z500".into();
        app.confirm_name_box(cx);
        assert_eq!(app.view_state.selected, before);
        assert_eq!(extent(app, cx), (20, 11));
    })
    .unwrap();
}

#[gpui::test]
fn shrink_freeze_structure_undo_and_whole_column_semantics(cx: &mut TestAppContext) {
    use visigrid_engine::structural::Axis;
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.set_cell_value(6, 2, "end", cx);
        app.shrink_grid_to_data(cx);
        assert_eq!(extent(app, cx), (7, 3));
        app.insert_rows(2, 5, cx);
        assert_eq!(extent(app, cx), (12, 3));
        app.undo(cx);
        assert_eq!(extent(app, cx), (7, 3));
        app.redo(cx);
        assert_eq!(extent(app, cx), (12, 3));
        app.delete_rows(2, 5, cx);
        assert_eq!(extent(app, cx), (7, 3));
        app.show_add_count(Axis::Col, cx);
        app.close_count_dialog(cx);
        // Formatting-only cells outside the extent retain existing whole-column semantics.
        app.with_active_sheet_mut(cx, |s| s.set_italic(999, 1, true));
        app.select_col(1, false, cx);
        assert_eq!(app.selection_range(), ((0, 1), (NUM_ROWS - 1, 1)));
        let limited_ranges = app.format_apply_ranges(cx);
        app.set_bold(true, cx);
        assert_eq!(extent(app, cx), (7, 3));
        app.toggle_limited_grid(cx);
        assert_eq!(extent(app, cx), (NUM_ROWS, NUM_COLS));
        assert_eq!(app.format_apply_ranges(cx), limited_ranges);
        assert!(app.sheet(cx).get_format(999, 1).bold);
        // Existing large-selection formatting deliberately stops at populated bounds.
        assert!(!app.sheet(cx).get_format(NUM_ROWS - 1, 1).bold);
        app.toggle_limited_grid(cx);
        app.with_active_sheet_mut(cx, |s| s.frozen_panes = (8, 4));
        app.sync_freeze_panes((8, 4));
        app.shrink_grid_to_data(cx);
        assert_eq!(extent(app, cx), (9, 5));
    })
    .unwrap();
}

#[gpui::test]
fn sidecar_round_trip_and_export_bytes_are_independent_of_view(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("limited.sheet");
        app.set_cell_value(6, 2, "saved", cx);
        app.grow_grid(5, 2, cx);
        visigrid_io::native::save_workbook(app.wb(cx), &path).unwrap();
        crate::settings::save_doc_settings(&path, &app.doc_settings).unwrap();
        let mut compared = false;
        for _ in 0..5 {
            let timestamp = chrono::Utc::now().timestamp();
            let before = visigrid_io::xlsx::export_to_buffer(app.wb(cx), None)
                .unwrap()
                .0;
            app.toggle_limited_grid(cx);
            let after = visigrid_io::xlsx::export_to_buffer(app.wb(cx), None)
                .unwrap()
                .0;
            app.toggle_limited_grid(cx);
            if chrono::Utc::now().timestamp() == timestamp {
                assert_eq!(before, after);
                compared = true;
                break;
            }
        }
        assert!(
            compared,
            "XLSX comparison needs a shared creation timestamp"
        );
        let restored = visigrid_io::native::load_workbook(&path).unwrap();
        app.wb_mut(cx, |wb| *wb = restored);
        app.doc_settings = crate::settings::load_doc_settings(&path);
        assert!(app.limited_grid_enabled(cx));
        assert_eq!(extent(app, cx), (25, 13));
        assert_eq!(app.sheet(cx).get_raw(6, 2), "saved");
        app.toggle_limited_grid(cx);
        assert_eq!(extent(app, cx), (NUM_ROWS, NUM_COLS));
        assert_eq!(app.sheet(cx).get_raw(6, 2), "saved");
    })
    .unwrap();
}

#[gpui::test]
fn large_sheet_bounds_are_cached_during_scroll(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.wb_mut(cx, |wb| {
            wb.begin_batch();
            for row in 0..300_000 {
                wb.set_cell_value_tracked(0, row, 0, "1");
            }
            wb.end_batch();
        });
        assert_eq!(extent(app, cx), (300_000, 11));
        let scans = app.limited_grid.scans;
        for _ in 0..50 {
            app.scroll(3, 0, cx);
            app.sync_limited_grid(cx);
            let _ = app.pane_rows(&app.view_state);
        }
        assert_eq!(
            app.limited_grid.scans, scans,
            "no cell scans on scroll/render"
        );
        app.set_cell_value(300_005, 0, "new", cx);
        assert_eq!(extent(app, cx).0, 300_006);
        assert_eq!(app.limited_grid.scans, scans + 1);
    })
    .unwrap();
}

#[test]
fn bounds_include_metadata_and_ignore_formatting() {
    use visigrid_engine::{sheet::MergedRegion, validation::ValidationRule};
    let mut wb = visigrid_engine::workbook::Workbook::new();
    let sheet = wb.active_sheet_mut();
    sheet.toggle_bold(999, 99);
    assert_eq!(content_bounds(sheet), GridExtent { rows: 1, cols: 1 });
    wb.set_cell_value_tracked(0, 2, 2, "=\"\"");
    let sheet = wb.active_sheet_mut();
    assert_eq!(content_bounds(sheet), GridExtent { rows: 3, cols: 3 });
    sheet
        .add_merge(MergedRegion {
            start: (5, 2),
            end: (7, 4),
        })
        .unwrap();
    assert_eq!(content_bounds(sheet), GridExtent { rows: 8, cols: 5 });
    sheet.set_comment(
        25,
        12,
        Some(visigrid_engine::cell::CellComment {
            text: "note".into(),
            author: "test".into(),
        }),
    );
    assert_eq!(content_bounds(sheet), GridExtent { rows: 26, cols: 13 });
    sheet.set_cell_validation(
        30,
        15,
        ValidationRule::list_inline(vec!["yes".into(), "no".into()]),
    );
    assert_eq!(
        content_bounds(sheet),
        GridExtent { rows: 26, cols: 13 },
        "validation alone does not extend content bounds"
    );
    wb.set_cell_value_tracked(0, 30, 15, "yes");
    assert_eq!(
        content_bounds(wb.active_sheet()),
        GridExtent { rows: 31, cols: 16 }
    );
}

#[gpui::test]
fn whole_column_validation_keeps_empty_sheet_at_starting_extent(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.with_active_sheet_mut(cx, |sheet| {
            sheet.set_validation(
                0,
                25,
                NUM_ROWS - 1,
                25,
                visigrid_engine::validation::ValidationRule::list_inline(vec![
                    "yes".into(),
                    "no".into(),
                ]),
            );
        });
        assert_eq!(extent(app, cx), (20, 11));
    })
    .unwrap();
}

#[test]
fn bounds_include_empty_table_body_totals_and_pivot_output() {
    use visigrid_engine::{
        pivot::{PivotSource, PivotTable},
        table::TableRange,
        workbook::Workbook,
    };
    let mut wb = Workbook::new();
    let id = wb.active_sheet_id();
    let table = wb
        .create_table(
            id,
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 50,
                end_col: 3,
            },
            "EmptyBody",
        )
        .unwrap()
        .table_id();
    assert_eq!(
        content_bounds(wb.active_sheet()),
        GridExtent { rows: 51, cols: 4 }
    );
    wb.set_table_totals_visible(table, true, Default::default())
        .unwrap();
    assert_eq!(
        content_bounds(wb.active_sheet()),
        GridExtent { rows: 52, cols: 4 }
    );
    wb.active_sheet_mut().pivots.push(PivotTable {
        id: 1,
        name: "Output".into(),
        source: PivotSource {
            table_id: None,
            sheet_id: id,
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0,
        },
        definition: Default::default(),
        anchor_row: 70,
        anchor_col: 5,
        extent: Some((10, 3)),
        last_refresh: None,
        stale: false,
        source_generation: None,
    });
    assert_eq!(
        content_bounds(wb.active_sheet()),
        GridExtent { rows: 80, cols: 8 }
    );
}

#[gpui::test]
fn sheet_ids_survive_reorder_and_deleted_extent_is_removed(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.grow_grid(5, 2, cx);
        let id = app.sheet(cx).id;
        app.add_sheet(cx);
        assert_eq!(extent(app, cx), (20, 11));
        app.move_sheet(0, 1, cx);
        app.goto_sheet(1, cx);
        assert_eq!(app.sheet(cx).id, id);
        assert_eq!(extent(app, cx), (25, 13));
        app.delete_sheet(1, cx);
        app.sync_limited_grid(cx);
        assert!(!app.doc_settings.display.sheet_extents.contains_key(&id.0));
    })
    .unwrap();
}

#[gpui::test]
fn metadata_changes_invalidate_bounds_and_empty_last_row_never_grows(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.with_active_sheet_mut(cx, |sheet| {
            sheet.set_comment(
                99,
                30,
                Some(visigrid_engine::cell::CellComment {
                    text: "note".into(),
                    author: "test".into(),
                }),
            )
        });
        assert_eq!(extent(app, cx), (100, 31));
        app.with_active_sheet_mut(cx, |sheet| sheet.set_comment(99, 30, None));
        assert_eq!(extent(app, cx), (20, 11));
        app.with_active_sheet_mut(cx, |sheet| {
            sheet.set_cell_validation(
                19,
                5,
                visigrid_engine::validation::ValidationRule::list_inline(vec![
                    "last row content in another column".into(),
                ]),
            );
        });
        app.view_state.selected = (19, 0);
        for _ in 0..20 {
            app.move_selection(1, 0, cx);
        }
        assert_eq!(extent(app, cx), (20, 11));
        assert_eq!(app.view_state.selected, (19, 0));
    })
    .unwrap();
    cx.simulate_keystrokes(view.into(), "enter enter");
    view.update(cx, |app, _, cx| {
        assert_eq!(
            extent(app, cx),
            (20, 11),
            "validation on an empty last row must not trigger growth"
        );
        assert_eq!(app.view_state.selected, (19, 0));
        app.set_cell_value(19, 5, "last row content in another column", cx);
        // Page Down must not grow even when a tiny viewport fits only one row.
        app.grid_layout.viewport_size.1 = 1.;
        app.page_down(cx);
        assert_eq!(extent(app, cx), (20, 11));
        app.move_selection(1, 0, cx);
        assert_eq!(extent(app, cx), (21, 11));
        assert_eq!(app.view_state.selected, (20, 0));
        app.toggle_limited_grid(cx);
        app.view_state.selected = (499, 25);
        app.start_edit_clear(cx);
        app.edit_value = "pending".into();
        app.edit_cursor = 7;
        app.toggle_limited_grid(cx);
        assert_eq!(
            app.sheet(cx).get_raw(499, 25),
            "pending",
            "toggle must not relocate a pending edit"
        );
        assert_eq!(extent(app, cx), (500, 26));
    })
    .unwrap();
}

#[gpui::test]
fn hidden_last_row_freeze_and_external_jump_paths_respect_extent(cx: &mut TestAppContext) {
    let view = setup(cx);
    view.update(cx, |app, _, cx| {
        app.select_row(19, false, cx);
        app.hide_rows(cx);
        app.view_state.selected = (0, 0);
        app.view_state.selection_end = None;
        app.jump_selection(1, 0, cx);
        assert_eq!(
            app.view_state.selected.0, 18,
            "Ctrl+Down stops at the last visible row"
        );
        app.move_selection(1, 0, cx);
        assert_eq!(app.view_state.selected.0, 18);
        app.with_active_sheet_mut(cx, |s| s.frozen_panes = (2, 1));
        app.sync_freeze_panes((2, 1));
        app.scroll(i32::MAX, i32::MAX, cx);
        let rows = app.pane_rows(&app.view_state);
        assert_eq!(
            rows.frozen.iter().map(|s| s.index).collect::<Vec<_>>(),
            vec![0, 1]
        );
        assert!(rows.body.iter().all(|s| s.index < 19));
        app.wb_mut(cx, |wb| {
            let other = wb.add_sheet();
            wb.define_name_for_cell("Outside", other, 499, 25).unwrap();
        });
        let before = (app.sheet_index(cx), app.view_state.selected);
        app.jump_to_named_range("Outside", cx);
        assert_eq!((app.sheet_index(cx), app.view_state.selected), before);
        assert!(app
            .status_message
            .as_ref()
            .unwrap()
            .contains("Z500 is outside"));
        app.set_cell_value(0, 0, "=Z500", cx);
        app.view_state.selected = (0, 0);
        app.cycle_trace_precedent(false, cx);
        assert_eq!(app.view_state.selected, (0, 0));
        assert!(app
            .status_message
            .as_ref()
            .unwrap()
            .contains("Z500 is outside"));
        assert_eq!(extent(app, cx), (20, 11));
    })
    .unwrap();
}
