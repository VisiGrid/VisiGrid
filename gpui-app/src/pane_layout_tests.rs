use super::*;
use crate::{
    history::{History, UndoAction},
    table_command_scope::{
        metadata_history_allowed, restore_freeze_panes, validate_freeze_history,
    },
    table_edit::tests::fixture,
};

#[test]
fn frozen_and_scrolling_rows_share_visibility_and_view_slot_boundaries() {
    let wb = fixture(true);
    let sheet = wb.active_sheet();
    let rows = sheet
        .build_saved_table_view(sheet.rows)
        .unwrap()
        .unwrap()
        .rows()
        .clone();
    let size = |r| {
        if rows.is_view_row_visible(r) {
            20.0
        } else {
            0.0
        }
    };
    let p = layout(30, 5, 5, 141.0, size);
    assert_eq!(
        p.frozen.iter().map(|s| s.index).collect::<Vec<_>>(),
        [0, 1, 2, 4]
    );
    assert_eq!(
        p.body.iter().map(|s| s.index).collect::<Vec<_>>(),
        [5, 6, 7]
    );
    assert_eq!(p.body_offset, 81.0);
    assert_eq!(rows.view_to_data(p.frozen[3].index), 5);
    assert_eq!(rows.view_to_data(p.body[0].index), 3);
    assert_eq!(p.hit(70.0), Some(4));
    assert_eq!(p.hit(80.5), None); // divider
    assert_eq!(p.hit(90.0), Some(5));
    for s in p.frozen.iter().chain(&p.body) {
        assert_eq!(s.offset, offset(s.index, 5, 5, size));
    }
}

#[test]
fn scrolling_starts_at_a_slot_not_a_filtered_rank() {
    let size = |r| if [1, 2, 4].contains(&r) { 0.0 } else { 20.0 };
    let p = layout(20, 0, 5, 60.0, size);
    assert_eq!(
        p.body.iter().map(|s| s.index).collect::<Vec<_>>(),
        [5, 6, 7]
    );
    assert_eq!(p.hit(0.0), Some(5));
    assert_eq!(p.hit(59.0), Some(7));
    assert_eq!(p.hit(-1.0), None);
    assert_eq!(offset(3, 0, 5, size), -20.0);
}

#[test]
fn hidden_columns_and_custom_extents_align_frozen_corners_and_body() {
    let size = |c| match c {
        0 => 40.0,
        1 | 4 => 0.0,
        2 => 80.0,
        _ => 60.0,
    };
    let p = layout(20, 3, 4, 211.0, size);
    assert_eq!(p.frozen.iter().map(|s| s.index).collect::<Vec<_>>(), [0, 2]);
    assert_eq!(p.body.iter().map(|s| s.index).collect::<Vec<_>>(), [5, 6]);
    assert_eq!(p.body_offset, 121.0);
    for s in p.frozen.iter().chain(&p.body) {
        assert_eq!(p.hit(s.offset + 1.0), Some(s.index));
        assert_eq!(s.offset, offset(s.index, 3, 4, size));
    }
}

#[test]
fn hidden_frozen_band_keeps_only_the_divider_and_does_not_skip_body_records() {
    let size = |r| if r < 4 { 0.0 } else { 20.0 };
    let p = layout(20, 4, 4, 61.0, size);
    assert!(p.frozen.is_empty());
    assert_eq!(p.body[0].index, 4);
    assert_eq!(p.body_offset, 1.0);
    assert_eq!(p.body.len(), 3);
}

#[test]
fn oversized_frozen_regions_are_bounded_by_the_viewport() {
    let p = layout(65_536, 60_000, 60_000, 400.0, |_| 20.0);
    assert_eq!(p.frozen.len(), 20);
    assert!(p.body.is_empty());
    assert_eq!(
        ensure_visible(65_536, 60_000, 60_000, 60_010, 400.0, |_| 20.0),
        60_000
    );
}

#[test]
fn keyboard_navigation_scrolls_by_actual_visible_extents() {
    let size = |r| if [3, 5, 7].contains(&r) { 0.0 } else { 20.0 };
    assert_eq!(ensure_visible(30, 4, 4, 8, 101.0, size), 6);
    assert_eq!(ensure_visible(30, 4, 9, 2, 101.0, size), 9); // frozen selection
    assert_eq!(ensure_visible(30, 4, 9, 6, 101.0, size), 6); // above body
    assert_eq!(ensure_visible(30, 4, 4, 5, 101.0, size), 4); // hidden selection
    let p = layout(30, 4, 6, 101.0, size);
    assert_eq!(p.body.iter().map(|s| s.index).collect::<Vec<_>>(), [6, 8]);
}

#[test]
fn wheel_scrolling_skips_hidden_slots_and_never_crosses_the_frozen_boundary() {
    let visible = |r| ![3, 5, 7].contains(&r);
    assert_eq!(scroll(20, 4, 4, 2, visible), 8);
    assert_eq!(scroll(20, 4, 8, -2, visible), 4);
    assert_eq!(scroll(20, 4, 4, -20, visible), 4);
    assert_eq!(scroll(20, 4, 18, 20, visible), 19);
    assert_eq!(scroll(20, 4, 8, 0, visible), 8);
    assert_eq!(scroll(20, 4, 5, 1, visible), 8); // hidden starting slot draws row 6 first
}

#[test]
fn freeze_metadata_preserves_criteria_values_totals_and_history() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let base = wb.clone();
    let sid = wb.active_sheet_id();
    let action = UndoAction::FreezePanesChanged {
        sheet_id: sid,
        old_frozen_rows: 0,
        old_frozen_cols: 0,
        new_frozen_rows: 5,
        new_frozen_cols: 2,
    };
    assert!(metadata_history_allowed(&wb, &action));
    for (forward, panes) in [(true, (5, 2)), (false, (0, 0)), (true, (5, 2))] {
        validate_freeze_history(&wb, &action, forward).unwrap();
        let rev = wb.revision();
        restore_freeze_panes(&mut wb, sid, panes).unwrap();
        assert!(wb.revision() > rev);
        assert_eq!(wb.active_sheet().frozen_panes, panes);
        assert_eq!(
            wb.active_sheet().table_view_spec(),
            base.active_sheet().table_view_spec()
        );
        assert_eq!(
            wb.active_sheet().get_raw(7, 3),
            base.active_sheet().get_raw(7, 3)
        );
        wb.active_sheet().build_saved_table_view(30).unwrap();
    }
    let revision = wb.revision();
    restore_freeze_panes(&mut wb, sid, (5, 2)).unwrap();
    assert_eq!(wb.revision(), revision);
    let mut history = History::new();
    history.record_action_with_provenance(action, None);
    for pos in [0, 1] {
        let preview = history
            .build_workbook_before(pos, Some(&base), 100, 10_000)
            .unwrap()
            .workbook;
        assert_eq!(
            preview.active_sheet().frozen_panes,
            if pos == 0 { (0, 0) } else { (5, 2) }
        );
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join("freeze.sheet");
        visigrid_io::native::save_workbook(&preview, &file).unwrap();
        let reopened = visigrid_io::native::load_workbook(&file).unwrap();
        assert_eq!(
            reopened.active_sheet().frozen_panes,
            preview.active_sheet().frozen_panes
        );
        assert_eq!(
            reopened.active_sheet().table_view_spec(),
            preview.active_sheet().table_view_spec()
        );
        let json = visigrid_io::json::export_workbook(&preview, &[], 0).unwrap();
        let loaded = visigrid_io::json::import_any(&json).unwrap().0;
        assert_eq!(
            loaded.active_sheet().frozen_panes,
            preview.active_sheet().frozen_panes
        );
    }
}

#[test]
fn invalid_grouped_freeze_destinations_refuse_before_any_metadata_changes() {
    let mut wb = fixture(true);
    let sid = wb.active_sheet_id();
    let action = UndoAction::Group {
        actions: vec![
            UndoAction::FreezePanesChanged {
                sheet_id: sid,
                old_frozen_rows: 0,
                old_frozen_cols: 0,
                new_frozen_rows: 5,
                new_frozen_cols: 0,
            },
            UndoAction::FreezePanesChanged {
                sheet_id: sid,
                old_frozen_rows: 5,
                old_frozen_cols: 0,
                new_frozen_rows: 31,
                new_frozen_cols: 0,
            },
        ],
        description: "Freeze".into(),
    };
    assert!(validate_freeze_history(&wb, &action, true).is_err());
    validate_freeze_history(&wb, &action, false).unwrap();
    assert!(restore_freeze_panes(&mut wb, sid, (0, 9)).is_err());
    assert_eq!(wb.active_sheet().frozen_panes, (0, 0));
}

#[test]
fn every_small_visibility_mask_keeps_pixels_slots_and_hits_in_agreement() {
    for mask in 0u32..128 {
        for boundary in 0..=7 {
            for scroll_start in boundary..=7 {
                let size = |r| {
                    if mask & (1u32 << r) == 0 {
                        0.0
                    } else {
                        10.0 + r as f32
                    }
                };
                let p = layout(7, boundary, scroll_start, 65.0, size);
                let mut indices = std::collections::BTreeSet::new();
                for s in p.frozen.iter().chain(&p.body) {
                    assert!(indices.insert(s.index));
                    assert!(size(s.index) > 0.0);
                    assert_eq!(s.offset, offset(s.index, boundary, scroll_start, size));
                    assert_eq!(p.hit(s.offset + s.size / 2.0), Some(s.index));
                }
                assert!(p.frozen.iter().all(|s| s.index < boundary));
                assert!(p
                    .body
                    .iter()
                    .all(|s| s.index >= boundary && s.index >= scroll_start));
            }
        }
    }
    assert_eq!(scroll(7, 7, 7, -1, |_| true), 7);
    assert_eq!(ensure_visible(7, 7, 7, 6, 65.0, |_| 10.0), 7);
}

#[test]
fn sparse_row_iterator_geometry_matches_the_dense_visibility_model() {
    for mask in 0u32..128 {
        let size = |r| if mask & (1u32 << r) == 0 { 0.0 } else { 20.0 };
        let dense = layout(7, 3, 4, 80.0, size);
        let sparse = from_slots(
            3,
            80.0,
            (0..3).filter(|&r| size(r) > 0.0).map(|r| (r, size(r))),
            (4..7).filter(|&r| size(r) > 0.0).map(|r| (r, size(r))),
        );
        assert_eq!(dense.frozen, sparse.frozen);
        assert_eq!(dense.body, sparse.body);
        assert_eq!(dense.body_offset, sparse.body_offset);
    }
}
