//! Regression coverage for replaying Table criteria and sparse cell history.
use crate::{
    history::{History, PreviewBuildError, UndoAction},
    rewind::{preview_rows, preview_selection},
    table_cell_history::TableCellsCommit,
    table_edit::{prepare_table_writes, tests::fixture, TableCellWrite},
};
use visigrid_engine::{filter::SortDirection, sheet::SheetId, workbook::Workbook};

fn record_cells(history: &mut History, wb: &mut Workbook, writes: Vec<TableCellWrite>) {
    let after = prepare_table_writes(wb, 0, &writes).unwrap();
    let commit = TableCellsCommit::capture(
        wb.active_sheet(),
        after.active_sheet(),
        writes.iter().map(|w| (w.row, w.col)),
    );
    history.record_action_with_provenance(
        UndoAction::TableCellsChanged {
            sheet_index: 0,
            commit: Box::new(commit),
            description: "Edit visible cells".into(),
        },
        None,
    );
    *wb = after;
}

fn record_view(
    history: &mut History,
    wb: &mut Workbook,
    spec: Option<visigrid_engine::table_view::TableViewSpec>,
) {
    let commit = wb.set_table_view_spec(SheetId(7), spec).unwrap();
    history.record_action_with_provenance(
        UndoAction::TableViewChanged {
            sheet_index: 0,
            commit: Box::new(commit),
            description: "Change Table view".into(),
        },
        None,
    );
}

fn build(history: &mut History, base: &Workbook, before: usize) -> crate::history::PreviewBuildResult {
    history
        .build_workbook_before(before, Some(base), 100, 10_000)
        .unwrap()
}

#[test]
fn loaded_base_projection_is_rebuilt_without_any_history() {
    let base = fixture(true);
    let preview = build(&mut History::new(), &base, 0);
    let rows = preview_rows(preview.view_state.per_sheet.first());
    assert_eq!(
        (4..=6).map(|r| rows.view_to_data(r)).collect::<Vec<_>>(),
        vec![5, 3, 6]
    );
    assert!(rows.data_to_view(4).is_none());
    assert_eq!(rows.view_to_data(8), 8);
    assert_eq!(base.active_sheet().get_display(0, 1), "100");
}

#[test]
fn scrubbing_reconstructs_criteria_edits_hidden_records_and_clear_view() {
    let base = fixture(true);
    let mut wb = base.clone();
    let mut history = History::new();
    let mut spec = wb.active_sheet().table_view_spec().unwrap().clone();
    spec.sort.as_mut().unwrap().direction = SortDirection::Descending;
    record_view(&mut history, &mut wb, Some(spec));
    record_cells(
        &mut history,
        &mut wb,
        vec![
            TableCellWrite::value(3, 1, "East".into()),
            TableCellWrite::value(5, 2, "99".into()),
        ],
    );
    record_view(&mut history, &mut wb, None);
    let fp = history.fingerprint();
    let before_edit = build(&mut history, &base, 1);
    let rows = preview_rows(before_edit.view_state.per_sheet.first());
    assert_eq!(
        (3..=5).map(|r| rows.view_to_data(r)).collect::<Vec<_>>(),
        vec![6, 3, 5]
    );
    assert_eq!(before_edit.workbook.active_sheet().get_display(0, 1), "100");
    let before_clear = build(&mut history, &base, 2);
    let rows = preview_rows(before_clear.view_state.per_sheet.first());
    assert!(rows.data_to_view(3).is_none());
    assert!(rows.data_to_view(4).is_none());
    assert_eq!(
        before_clear.workbook.active_sheet().get_display(0, 1),
        "179"
    );
    assert_eq!(
        before_clear.workbook.active_sheet().get_display(5, 3),
        "198"
    );
    let cleared = build(&mut history, &base, 3);
    assert!(cleared.view_state.per_sheet[0].table_rows.is_none());
    assert!(cleared.workbook.active_sheet().table_view_spec().is_none());
    assert_eq!(history.fingerprint(), fp);
    assert_eq!(base.active_sheet().get_raw(3, 1), "West");
    assert_eq!(wb.active_sheet().get_raw(5, 2), "99");
}

#[test]
fn empty_table_projection_and_canonical_highlight_have_safe_focus() {
    let base = fixture(true);
    let mut wb = base.clone();
    let mut history = History::new();
    record_cells(
        &mut history,
        &mut wb,
        [3, 5, 6]
            .into_iter()
            .map(|r| TableCellWrite::value(r, 1, String::new()))
            .collect(),
    );
    let preview = build(&mut history, &base, 1);
    let rows = preview_rows(preview.view_state.per_sheet.first());
    let range = base.active_sheet().tables()[0].range;
    assert_eq!(
        preview_selection(&rows, (3, 1), (6, 1), Some(range)),
        ((2, 1), (2, 1))
    );
    for row in 3..=6 {
        assert!(rows.data_to_view(row).is_none());
    }
    let before = build(&mut history, &base, 0);
    let rows = preview_rows(before.view_state.per_sheet.first());
    assert_eq!(
        preview_selection(&rows, (5, 2), (5, 2), Some(range)),
        ((4, 2), (4, 2))
    );
}

#[test]
fn stale_cell_or_criteria_history_refuses_preview_without_touching_base() {
    let base = fixture(true);
    let mut wb = base.clone();
    let mut history = History::new();
    record_cells(
        &mut history,
        &mut wb,
        vec![TableCellWrite::value(3, 2, "55".into())],
    );
    let mut wrong_base = base.clone();
    wrong_base.set_cell_value_tracked(0, 3, 2, "999");
    assert!(matches!(
        history.build_workbook_before(1, Some(&wrong_base), 100, 10_000),
        Err(PreviewBuildError::InvariantViolation(_))
    ));
    assert_eq!(wrong_base.active_sheet().get_raw(3, 2), "999");
    let mut history = History::new();
    record_view(&mut history, &mut wb, None);
    wrong_base.set_table_view_spec(SheetId(7), None).unwrap();
    assert!(matches!(
        history.build_workbook_before(1, Some(&wrong_base), 100, 10_000),
        Err(PreviewBuildError::InvariantViolation(_))
    ));
}

#[test]
fn later_rewind_replays_retained_prefix_and_skips_audit_without_double_edits() {
    let base = fixture(true);
    let mut wb = base.clone();
    let mut history = History::new();
    record_cells(
        &mut history,
        &mut wb,
        vec![TableCellWrite::value(3, 2, "55".into())],
    );
    record_cells(
        &mut history,
        &mut wb,
        vec![TableCellWrite::value(5, 2, "66".into())],
    );
    let preview = build(&mut history, &base, 1);
    let id = history.entry_at(1).unwrap().id;
    wb.restore_snapshot_monotonic(&preview.workbook);
    history.truncate_and_append_rewind(1, id, 1, "Edit".into(), 1, 0);
    record_cells(
        &mut history,
        &mut wb,
        vec![TableCellWrite::value(6, 2, "77".into())],
    );
    let again = build(&mut history, &base, 3);
    assert_eq!(again.workbook.active_sheet().get_raw(3, 2), "55");
    assert_eq!(again.workbook.active_sheet().get_raw(5, 2), "20");
    assert_eq!(again.workbook.active_sheet().get_raw(6, 2), "77");
    let original = build(&mut history, &base, 0);
    assert_eq!(original.workbook.active_sheet().get_raw(3, 2), "30");
}

#[test]
fn every_sheet_gets_its_own_projection_and_invalid_layout_is_refused() {
    let mut base = fixture(true);
    let other = visigrid_engine::sheet::Sheet::new_with_name(SheetId(99), 30, 8, "Other");
    assert!(base.restore_sheet(1, other));
    base.set_active_sheet(1);
    let preview = build(&mut History::new(), &base, 0);
    assert!(preview.view_state.per_sheet[0].table_rows.is_some());
    assert!(preview.view_state.per_sheet[1].table_rows.is_none());
    assert_eq!(preview.workbook.active_sheet_index(), 1);
    base.set_cell_value_tracked(0, 3, 0, "neighbor");
    assert!(matches!(
        History::new().build_workbook_before(0, Some(&base), 100, 10_000),
        Err(PreviewBuildError::InvariantViolation(_))
    ));
}
