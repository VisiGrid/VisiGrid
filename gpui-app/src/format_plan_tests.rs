use super::*;
use crate::{
    history::{FormatActionKind, History},
    table_edit::tests::fixture,
};
use visigrid_engine::{
    cell::{BorderStyle, CellComment},
    sheet::SheetId,
};

fn rows(wb: &Workbook, index: usize) -> RowView {
    let s = wb.sheet(index).unwrap();
    s.build_saved_table_view(s.rows)
        .unwrap()
        .map(|v| v.rows().clone())
        .unwrap_or_else(|| RowView::new(s.rows))
}
fn patches(wb: &Workbook, ranges: &[Range], op: Operation) -> Result<Vec<CellFormatPatch>, String> {
    plan(
        wb,
        0,
        &rows(wb, 0),
        Some(&wb.active_sheet().manual_hidden_rows()),
        None,
        ranges,
        &op,
    )
}
fn action(patches: Vec<CellFormatPatch>) -> UndoAction {
    UndoAction::Format {
        sheet_index: 0,
        patches,
        kind: FormatActionKind::Bold,
        description: "Table formatting".into(),
    }
}
fn thin() -> CellBorder {
    CellBorder {
        style: BorderStyle::Thin,
        color: Some([20, 40, 60, 255]),
    }
}

#[test]
fn visible_formatting_tracks_canonical_records_and_preserves_formulas_comments_and_totals() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.active_sheet_mut().set_comment(
        5,
        3,
        Some(CellComment {
            text: "Keep".into(),
            author: "QA".into(),
        }),
    );
    let base = wb.clone();
    let p = patches(&wb, &[((3, 1), (6, 3))], Operation::Bold(true)).unwrap();
    assert_eq!(p.len(), 9);
    assert!(p.iter().all(|p| [3, 5, 6].contains(&p.row)));
    for forward in [true, false, true] {
        validate_history(&wb, &action(p.clone()), forward).unwrap();
        apply(&mut wb, 0, &p, forward);
        for r in [3, 5, 6] {
            assert_eq!(wb.active_sheet().get_format(r, 3).bold, forward);
        }
        assert!(!wb.active_sheet().get_format(4, 3).bold);
        assert_eq!(wb.active_sheet().get_raw(5, 3), "=C6*2");
        assert_eq!(
            wb.active_sheet().comment(5, 3),
            base.active_sheet().comment(5, 3)
        );
        assert_eq!(
            wb.active_sheet().get_display(7, 3),
            base.active_sheet().get_display(7, 3)
        );
        assert_eq!(
            wb.active_sheet().table_view_spec(),
            base.active_sheet().table_view_spec()
        );
    }
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let loaded = visigrid_io::json::import_any(&json).unwrap().0;
    assert!(loaded.active_sheet().get_format(5, 3).bold);
    assert!(!loaded.active_sheet().get_format(4, 3).bold);
    assert_eq!(
        loaded.active_sheet().comment(5, 3),
        base.active_sheet().comment(5, 3)
    );
    let mut history = History::new();
    history.record_action_with_provenance(action(p), None);
    for position in [0, 1] {
        let preview = history
            .build_workbook_before(position, Some(&base), 100, 10_000)
            .unwrap()
            .workbook;
        assert_eq!(preview.active_sheet().get_format(5, 3).bold, position == 1);
        assert_eq!(
            preview.active_sheet().get_display(7, 3),
            base.active_sheet().get_display(7, 3)
        );
    }
}

#[test]
fn every_format_property_and_painter_snapshot_targets_the_sorted_record() {
    let base = fixture(true);
    let ops = vec![
        Operation::Bold(true),
        Operation::Italic(true),
        Operation::Underline(true),
        Operation::Strike(true),
        Operation::Font(Some("Test font".into())),
        Operation::Size(Some(18.)),
        Operation::Color(Some([1, 2, 3, 255])),
        Operation::Background(Some([30, 40, 50, 255])),
        Operation::Align(Alignment::CenterAcrossSelection),
        Operation::Vertical(VerticalAlignment::Top),
        Operation::Overflow(TextOverflow::Wrap),
        Operation::Number(NumberFormat::Percent { decimals: 3 }),
        Operation::Style(CellStyle::Input),
        Operation::Replace(CellFormat {
            italic: true,
            font_size: Some(24.),
            ..Default::default()
        }),
    ];
    for op in ops {
        let mut wb = base.clone();
        let p = patches(&wb, &[((4, 2), (4, 2))], op).unwrap();
        assert_eq!(p.len(), 1);
        assert_eq!((p[0].row, p[0].col), (5, 2));
        apply(&mut wb, 0, &p, true);
        assert_eq!(
            wb.active_sheet().get_format(4, 2),
            base.active_sheet().get_format(4, 2)
        );
        assert_eq!(
            wb.active_sheet().get_format(3, 2),
            base.active_sheet().get_format(3, 2)
        );
        apply(&mut wb, 0, &p, false);
        assert_eq!(
            wb.active_sheet().get_format(5, 2),
            base.active_sheet().get_format(5, 2)
        );
    }
    let mut wb = base.clone();
    wb.active_sheet_mut()
        .set_background_color(5, 2, Some([1, 2, 3, 255]));
    wb.active_sheet_mut()
        .set_font_color(5, 2, Some([3, 2, 1, 255]));
    let p = patches(&wb, &[((4, 2), (4, 2))], Operation::Style(CellStyle::Note)).unwrap();
    assert!(p[0].after.background_color.is_none() && p[0].after.font_color.is_none());
    let p = patches(&wb, &[((4, 2), (4, 2))], Operation::Clear).unwrap();
    assert!(p[0].after.is_default());
}

#[test]
fn header_footer_and_merged_title_formats_do_not_write_protected_values() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.active_sheet_mut()
        .add_merge(visigrid_engine::sheet::MergedRegion {
            start: (0, 4),
            end: (0, 6),
        })
        .unwrap();
    let p = patches(
        &wb,
        &[((0, 4), (0, 6)), ((2, 1), (2, 3)), ((7, 1), (7, 3))],
        Operation::Italic(true),
    )
    .unwrap();
    let header = wb.active_sheet().get_raw(2, 1);
    let footer = wb.active_sheet().get_raw(7, 3);
    apply(&mut wb, 0, &p, true);
    assert!(wb.active_sheet().get_format(2, 1).italic && wb.active_sheet().get_format(7, 3).italic);
    assert_eq!(wb.active_sheet().get_raw(2, 1), header);
    assert_eq!(wb.active_sheet().get_raw(7, 3), footer);
    wb.active_sheet().build_saved_table_view(30).unwrap();
    apply(&mut wb, 0, &p, false);
    assert_eq!(wb.active_sheet().get_merge(0, 5).unwrap().start, (0, 4));
}

#[test]
fn hidden_rows_columns_overlapping_ranges_and_visible_counts_use_one_target_set() {
    let mut wb = fixture(true);
    wb.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    let rows = rows(&wb, 0);
    let hidden_cols = BTreeSet::from([2]);
    let p = plan(
        &wb,
        0,
        &rows,
        None,
        Some(&hidden_cols),
        &[((3, 1), (6, 3)), ((4, 1), (6, 3))],
        &Operation::Bold(true),
    )
    .unwrap();
    assert_eq!(
        p.iter().map(|p| (p.row, p.col)).collect::<Vec<_>>(),
        [(5, 1), (5, 3), (6, 1), (6, 3)]
    );
    assert_eq!(visible_count(&rows, None, 3, 6), 2);
    assert_eq!(visible_count(&rows, Some(&BTreeSet::from([5])), 3, 6), 1);
    assert!(patches(&wb, &[((3, 1), (3, 3))], Operation::Bold(true))
        .unwrap()
        .is_empty());
}

#[test]
fn border_geometry_uses_first_last_and_neighboring_visible_records() {
    let wb = fixture(true);
    let p = patches(
        &wb,
        &[((4, 1), (5, 3))],
        Operation::Borders(BorderApplyMode::Outline, thin()),
    )
    .unwrap();
    let get = |r, c| &p.iter().find(|p| (p.row, p.col) == (r, c)).unwrap().after;
    assert_eq!(get(5, 2).border_top, thin());
    assert!(!get(5, 2).border_bottom.is_set());
    assert_eq!(get(3, 2).border_bottom, thin());
    assert!(!get(3, 2).border_top.is_set());
    assert!(p.iter().all(|p| p.row != 4 && p.row != 6));
    let rows = rows(&wb, 0);
    assert_eq!(row_neighbor(&rows, None, 4, false), Some(2));
    assert_eq!(row_neighbor(&rows, None, 4, true), Some(3));
    assert_eq!(
        row_neighbor(&rows, Some(&BTreeSet::from([3])), 4, true),
        Some(6)
    );
    let p = patches(
        &wb,
        &[((4, 1), (5, 3))],
        Operation::Borders(BorderApplyMode::Inside, thin()),
    )
    .unwrap();
    assert_eq!(
        p.iter()
            .find(|p| (p.row, p.col) == (5, 3))
            .unwrap()
            .after
            .border_bottom,
        thin()
    );
    assert!(p.iter().find(|p| (p.row, p.col) == (3, 3)).is_none());
    for (mode, expected) in [(BorderApplyMode::Top, 5), (BorderApplyMode::Bottom, 3)] {
        let p = patches(&wb, &[((4, 1), (5, 3))], Operation::Borders(mode, thin())).unwrap();
        assert!(p.iter().all(|p| p.row == expected));
    }
}

#[test]
fn clearing_borders_updates_visible_neighbors_but_never_filtered_or_manual_hidden_cells() {
    let mut wb = fixture(true);
    for r in [2, 3, 4, 5, 6] {
        wb.active_sheet_mut()
            .set_borders(r, 2, thin(), thin(), thin(), thin());
    }
    let p = patches(
        &wb,
        &[((4, 2), (4, 2))],
        Operation::Borders(BorderApplyMode::Clear, thin()),
    )
    .unwrap();
    assert!(p.iter().any(|p| p.row == 2));
    assert!(p.iter().any(|p| p.row == 3));
    assert!(p.iter().all(|p| p.row != 4));
    apply(&mut wb, 0, &p, true);
    assert!(!wb.active_sheet().get_format(5, 2).has_any_border());
    assert!(!wb.active_sheet().get_format(2, 2).border_bottom.is_set());
    assert!(!wb.active_sheet().get_format(3, 2).border_top.is_set());
    assert_eq!(wb.active_sheet().get_format(4, 2).border_top, thin());
    apply(&mut wb, 0, &p, false);
    wb.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    let p = patches(
        &wb,
        &[((4, 2), (4, 2))],
        Operation::Borders(BorderApplyMode::Clear, thin()),
    )
    .unwrap();
    assert!(p.iter().all(|p| p.row != 3 && p.row != 4));
    assert!(p.iter().any(|p| p.row == 6));
    let p = patches(
        &wb,
        &[((4, 1), (6, 3))],
        Operation::Borders(BorderApplyMode::All, thin()),
    )
    .unwrap();
    assert!(p
        .iter()
        .all(|p| p.after.border_top == thin() && p.after.border_bottom == thin()));
}

#[test]
fn percent_conversion_checks_only_visible_targets_and_refuses_atomically() {
    let mut wb = fixture(true);
    wb.set_cell_text_tracked(0, 4, 3, "25%");
    let op = Operation::Number(NumberFormat::Percent { decimals: 2 });
    let p = patches(&wb, &[((3, 3), (6, 3))], op.clone()).unwrap();
    assert!(p.iter().all(|p| p.row != 4));
    wb.set_cell_text_tracked(0, 5, 3, "25%");
    assert!(patches(&wb, &[((0, 0), (0, 0)), ((3, 3), (6, 3))], op)
        .unwrap_err()
        .contains("convert"));
    assert!(wb.active_sheet().get_format(0, 0).is_default());
    assert_eq!(wb.active_sheet().get_raw(5, 3), "25%");
}

#[test]
fn invalid_adjacent_targets_bounds_and_recovery_refuse_without_partial_formatting() {
    let mut wb = fixture(true);
    assert!(patches(
        &wb,
        &[((0, 0), (0, 0)), ((4, 4), (4, 4))],
        Operation::Bold(true)
    )
    .is_err());
    assert!(wb.active_sheet().get_format(0, 0).is_default());
    assert!(patches(&wb, &[((4, 4), (4, 4))], Operation::Clear)
        .unwrap()
        .is_empty());
    assert!(patches(&wb, &[((0, 0), (30, 0))], Operation::Bold(true)).is_err());
    assert!(patches(&wb, &[((0, 0), (0, 8))], Operation::Bold(true)).is_err());
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(patches(&wb, &[((4, 1), (4, 1))], Operation::Bold(true)).is_err());
}

#[test]
fn formatting_undo_restores_absence_and_coalescing_keeps_original_before_state() {
    let mut wb = fixture(true);
    let base = wb.clone();
    assert!(wb.active_sheet().get_cell_opt(0, 5).is_none());
    let p = patches(&wb, &[((0, 5), (0, 5))], Operation::Bold(true)).unwrap();
    assert!(p[0].remove_cell_on_undo);
    let mut history = History::new();
    apply(&mut wb, 0, &p, true);
    history.record_format(0, p, FormatActionKind::Bold, "Bold".into());
    let p = patches(&wb, &[((0, 5), (0, 5))], Operation::Bold(false)).unwrap();
    apply(&mut wb, 0, &p, true);
    history.record_format(0, p, FormatActionKind::Bold, "Normal".into());
    let entry = history.undo().unwrap();
    let UndoAction::Format { patches, .. } = entry.action else {
        panic!("format history")
    };
    assert!(patches[0].remove_cell_on_undo);
    apply(&mut wb, 0, &patches, false);
    assert!(wb.active_sheet().get_cell_opt(0, 5).is_none());
    assert_eq!(
        wb.active_sheet().get_raw(3, 3),
        base.active_sheet().get_raw(3, 3)
    );
}

#[test]
fn history_can_revisit_hidden_records_and_preflights_all_group_destinations() {
    let mut wb = fixture(true);
    let p = patches(&wb, &[((4, 2), (4, 2))], Operation::Bold(true)).unwrap();
    apply(&mut wb, 0, &p, true);
    wb.active_sheet_mut()
        .set_manual_hidden_rows([5].into())
        .unwrap();
    validate_history(&wb, &action(p.clone()), false).unwrap();
    apply(&mut wb, 0, &p, false);
    assert!(!wb.active_sheet().get_format(5, 2).bold);
    let mut unsafe_p = p.clone();
    unsafe_p[0].col = 7;
    let group = UndoAction::Group {
        actions: vec![action(p), action(unsafe_p)],
        description: "Group".into(),
    };
    assert!(validate_history(&wb, &group, true).is_err());
    assert!(wb.active_sheet().get_format(5, 2).is_default());
    let mut history = History::new();
    history.record_action_with_provenance(group, None);
    assert!(history
        .build_workbook_before(1, Some(&wb), 100, 10_000)
        .is_err());
}

#[test]
fn another_sheet_and_large_selection_limits_keep_the_original_table_unchanged() {
    let mut wb = fixture(true);
    wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 2000, 100, "Report"));
    let rows = rows(&wb, 1);
    let p = plan(
        &wb,
        1,
        &rows,
        None,
        None,
        &[((5, 7), (6, 8))],
        &Operation::Bold(true),
    )
    .unwrap();
    apply(&mut wb, 1, &p, true);
    assert!(wb.sheet(1).unwrap().get_format(5, 7).bold);
    assert!(wb.active_sheet().get_format(5, 2).is_default());
    assert!(plan(
        &wb,
        1,
        &rows,
        None,
        None,
        &[((0, 0), (1999, 99))],
        &Operation::Bold(true)
    )
    .unwrap_err()
    .contains("100,000"));
    assert!(wb.sheet(1).unwrap().get_format(0, 0).is_default());
}

#[test]
fn decimal_changes_noops_and_side_borders_remain_visible_only() {
    let mut wb = fixture(true);
    wb.active_sheet_mut()
        .set_number_format(5, 2, NumberFormat::Percent { decimals: 2 });
    for (delta, expected) in [(1, 3), (-10, 0), (127, 10)] {
        let p = patches(&wb, &[((4, 2), (5, 2))], Operation::Decimals(delta)).unwrap();
        assert_eq!(p.len(), 1); // The other visible cell is General.
        assert_eq!(p[0].row, 5);
        assert_eq!(
            p[0].after.number_format,
            NumberFormat::Percent { decimals: expected }
        );
    }
    assert!(patches(&wb, &[((4, 2), (4, 2))], Operation::Decimals(0))
        .unwrap()
        .is_empty());
    let rev = wb.revision();
    apply(&mut wb, 0, &[], true);
    assert_eq!(wb.revision(), rev);
    for (mode, col) in [(BorderApplyMode::Left, 1), (BorderApplyMode::Right, 3)] {
        let p = patches(&wb, &[((3, 1), (6, 3))], Operation::Borders(mode, thin())).unwrap();
        assert_eq!(p.len(), 3);
        assert!(p.iter().all(|p| p.col == col && p.row != 4));
    }
    assert_eq!(
        col_neighbor(Some(&BTreeSet::from([2])), 1, 8, true),
        Some(3)
    );
    assert_eq!(
        col_neighbor(Some(&BTreeSet::from([2])), 3, 8, false),
        Some(1)
    );
}
