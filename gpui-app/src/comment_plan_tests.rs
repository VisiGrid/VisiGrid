use super::*;
use crate::{
    history::{apply_comment_patches, CommentPatch, History},
    table_command_scope::metadata_history_allowed,
    table_edit::tests::fixture,
};
use visigrid_engine::{cell::CellComment, sheet::SheetId};

fn note(text: &str) -> Option<CellComment> {
    Some(CellComment {
        text: text.into(),
        author: "QA".into(),
    })
}

fn action(
    row: usize,
    col: usize,
    before: Option<CellComment>,
    after: Option<CellComment>,
) -> UndoAction {
    UndoAction::Comments {
        sheet_index: 0,
        patches: vec![CommentPatch {
            remove_cell_on_undo: false,
            row,
            col,
            before,
            after,
        }],
        description: "Comment".into(),
    }
}

fn replay(wb: &mut Workbook, action: &UndoAction, forward: bool) -> Result<(), String> {
    validate_history(wb, action, forward)?;
    if let UndoAction::Comments {
        sheet_index,
        patches,
        ..
    } = action
    {
        apply_comment_patches(wb, *sheet_index, patches, forward);
    }
    Ok(())
}

#[test]
fn sorted_comment_targets_record_and_preserves_totals_formula_format_and_hidden_notes() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.active_sheet_mut().set_comment(4, 1, note("Hidden East"));
    wb.active_sheet_mut().set_bold(5, 3, true);
    let base = wb.clone();
    let view = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    let row = view.rows().view_to_data(4);
    assert_eq!(row, 5);
    validate_edit(&wb, 0, row, 3, true, Some(wb.revision())).unwrap();
    let edit = action(row, 3, None, note("Sorted record"));
    assert!(metadata_history_allowed(&wb, &edit));
    for forward in [true, false, true] {
        replay(&mut wb, &edit, forward).unwrap();
        let sheet = wb.active_sheet();
        assert_eq!(
            sheet.comment(5, 3),
            if forward { note("Sorted record") } else { None }.as_ref()
        );
        assert!(sheet.comment(4, 3).is_none());
        assert_eq!(sheet.comment(4, 1), note("Hidden East").as_ref());
        assert_eq!(sheet.get_raw(5, 3), "=C6*2");
        assert!(sheet.get_format(5, 3).bold);
        assert_eq!(
            sheet.get_display(7, 3),
            base.active_sheet().get_display(7, 3)
        );
        assert_eq!(
            sheet.table_view_spec(),
            base.active_sheet().table_view_spec()
        );
        assert_eq!(
            sheet
                .build_saved_table_view(30)
                .unwrap()
                .unwrap()
                .rows()
                .view_to_data(4),
            5
        );
    }
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let loaded = visigrid_io::json::import_any(&json).unwrap().0;
    assert_eq!(
        loaded.active_sheet().comment(5, 3),
        note("Sorted record").as_ref()
    );
    assert_eq!(
        loaded.active_sheet().comment(4, 1),
        note("Hidden East").as_ref()
    );
    assert_eq!(loaded.active_sheet().get_raw(5, 3), "=C6*2");
    assert_eq!(
        loaded.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
}

#[test]
fn comment_history_replays_hidden_records_after_filter_changes_and_rewinds() {
    let mut base = fixture(true);
    let edit = action(3, 2, None, note("Keep with West"));
    let mut history = History::new();
    history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), edit.clone(), None);
    let mut wb = base.clone();
    replay(&mut wb, &edit, true).unwrap();
    wb.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    assert!(validate_edit(&wb, 0, 3, 2, false, None).is_err());
    for forward in [false, true] {
        replay(&mut wb, &edit, forward).unwrap();
        assert_eq!(wb.active_sheet().comment(3, 2).is_some(), forward);
        assert!(wb.active_sheet().manual_hidden_rows().contains(&3));
    }
    // A historical patch may also target a record hidden by the saved filter.
    let hidden_edit = action(4, 2, None, note("East"));
    assert!(validate_edit(&wb, 0, 4, 2, true, None).is_err());
    replay(&mut wb, &hidden_edit, true).unwrap();
    replay(&mut wb, &hidden_edit, false).unwrap();
    base.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    for position in [0, 1] {
        let preview = history
            .build_workbook_before(position, Some(&base), 100, 10_000)
            .unwrap()
            .workbook;
        assert_eq!(
            preview.active_sheet().comment(3, 2).is_some(),
            position == 1
        );
        assert!(preview.active_sheet().manual_hidden_rows().contains(&3));
        assert_eq!(preview.active_sheet().get_raw(4, 1), "East");
    }
}

#[test]
fn headers_totals_and_notes_above_below_support_comment_metadata() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    for (row, col) in [(0, 0), (2, 1), (7, 3), (8, 7)] {
        validate_edit(&wb, 0, row, col, true, None).unwrap();
        let raw = wb.active_sheet().get_raw(row, col);
        let add = action(row, col, None, note("Note"));
        replay(&mut wb, &add, true).unwrap();
        let change = action(row, col, note("Note"), note("Updated"));
        replay(&mut wb, &change, true).unwrap();
        let delete = action(row, col, note("Updated"), None);
        replay(&mut wb, &delete, true).unwrap();
        replay(&mut wb, &delete, false).unwrap();
        assert_eq!(
            wb.active_sheet().comment(row, col),
            note("Updated").as_ref()
        );
        assert_eq!(wb.active_sheet().get_raw(row, col), raw);
    }
    wb.active_sheet().build_saved_table_view(30).unwrap();
}

#[test]
fn adjacent_comments_are_refused_before_any_group_replay_in_both_directions() {
    let wb = fixture(true);
    for col in [0, 4, 7] {
        assert!(validate_edit(&wb, 0, 5, col, true, None)
            .unwrap_err()
            .contains("beside"));
        let unsafe_add = action(5, col, None, note("Adjacent"));
        let unsafe_restore = action(5, col, note("Adjacent"), None);
        for (child, forward) in [(unsafe_add, true), (unsafe_restore, false)] {
            let group = UndoAction::Group {
                actions: vec![action(5, 1, None, note("Safe")), child],
                description: "Atomic comments".into(),
            };
            assert!(metadata_history_allowed(&wb, &group));
            assert!(validate_history(&wb, &group, forward).is_err());
            assert!(wb.active_sheet().comment(5, 1).is_none());
        }
    }
    // Deletion removes an adjacent comment and never introduces unsafe content.
    validate_history(&wb, &action(5, 7, note("Old"), None), true).unwrap();
    let mut cleared = wb.clone();
    cleared.set_table_view_spec(SheetId(7), None).unwrap();
    validate_edit(&cleared, 0, 5, 7, true, None).unwrap();
}

#[test]
fn stale_drafts_missing_sheets_bounds_and_recovery_refuse_before_mutation() {
    let mut wb = fixture(true);
    let revision = wb.revision();
    validate_edit(&wb, 0, 3, 2, true, Some(revision)).unwrap();
    wb.bump_revision_for_structure();
    assert!(validate_edit(&wb, 0, 3, 2, true, Some(revision))
        .unwrap_err()
        .contains("draft"));
    assert!(validate_edit(&wb, 9, 3, 2, true, None).is_err());
    for (row, col) in [(30, 0), (0, 8), (usize::MAX, 1)] {
        assert!(validate_edit(&wb, 0, row, col, true, None).is_err());
        assert!(validate_history(&wb, &action(row, col, None, note("Outside")), true).is_err());
    }
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(validate_edit(&wb, 0, 3, 2, true, None).is_err());
    assert!(wb.active_sheet().comment(3, 2).is_none());
}

#[test]
fn other_sheet_comments_and_mixed_history_do_not_bypass_unsafe_commands() {
    let mut wb = fixture(true);
    wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Report"));
    validate_edit(&wb, 1, 5, 7, true, None).unwrap();
    let mut edit = action(5, 7, None, note("Report"));
    if let UndoAction::Comments { sheet_index, .. } = &mut edit {
        *sheet_index = 1;
    }
    assert!(metadata_history_allowed(&wb, &edit));
    replay(&mut wb, &edit, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().comment(5, 7), note("Report").as_ref());
    wb.set_active_sheet(1);
    replay(&mut wb, &action(3, 2, None, note("Table")), true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().comment(3, 2), note("Table").as_ref());
    let group = UndoAction::Group {
        actions: vec![
            edit,
            UndoAction::Values {
                sheet_index: 1,
                changes: vec![],
            },
        ],
        description: "Not metadata-only".into(),
    };
    assert!(!metadata_history_allowed(&wb, &group));
}

#[test]
fn merged_title_anchor_and_manual_hides_keep_the_same_rules_after_clearing_criteria() {
    let mut wb = fixture(true);
    wb.active_sheet_mut()
        .add_merge(visigrid_engine::sheet::MergedRegion {
            start: (0, 0),
            end: (0, 2),
        })
        .unwrap();
    validate_edit(&wb, 0, 0, 0, true, None).unwrap();
    assert!(validate_edit(&wb, 0, 0, 1, true, None).is_err());
    let edit = action(0, 0, None, note("Merged title"));
    for forward in [true, false, true] {
        replay(&mut wb, &edit, forward).unwrap();
        assert_eq!(wb.active_sheet().comment(0, 0).is_some(), forward);
        assert_eq!(wb.active_sheet().get_merge(0, 0).unwrap().end, (0, 2));
    }
    wb.set_table_view_spec(SheetId(7), None).unwrap();
    wb.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    assert!(validate_edit(&wb, 0, 3, 2, true, None).is_err());
    validate_edit(&wb, 0, 4, 2, true, None).unwrap();
}

#[test]
fn missing_history_sheet_and_unsafe_rewind_return_errors() {
    let wb = fixture(true);
    let mut missing = action(3, 2, None, note("Missing"));
    if let UndoAction::Comments { sheet_index, .. } = &mut missing {
        *sheet_index = 99;
    }
    assert!(!metadata_history_allowed(&wb, &missing));
    for forward in [true, false] {
        assert!(validate_history(&wb, &missing, forward).is_err());
    }
    let mut history = History::new();
    history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), action(5, 7, None, note("Unsafe")), None);
    assert!(history
        .build_workbook_before(1, Some(&wb), 100, 10_000)
        .is_err());
    assert!(wb.active_sheet().comment(5, 7).is_none());
}

#[test]
fn undoing_a_comment_on_an_absent_cell_keeps_earlier_sparse_history_replayable() {
    let before = fixture(true);
    let mut wb = before.clone();
    wb.clear_cell_tracked(0, 3, 3);
    assert!(wb.active_sheet().get_cell_opt(3, 3).is_none());
    let clear = crate::table_cell_history::TableCellsCommit::capture(
        before.active_sheet(),
        wb.active_sheet(),
        [(3, 3)],
    );
    let mut add = action(3, 3, None, note("Temporarily annotate empty record"));
    if let UndoAction::Comments { patches, .. } = &mut add {
        patches[0].remove_cell_on_undo = true;
    }
    for _ in 0..2 {
        replay(&mut wb, &add, true).unwrap();
        replay(&mut wb, &add, false).unwrap();
        assert!(wb.active_sheet().get_cell_opt(3, 3).is_none());
    }
    clear.replay(&mut wb, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(3, 3), "=C4*2");
    clear.replay(&mut wb, false).unwrap();
    assert!(wb.active_sheet().get_cell_opt(3, 3).is_none());
    // Newer cell content/formatting must never be cleared by metadata undo.
    replay(&mut wb, &add, true).unwrap();
    wb.set_cell_value_tracked(0, 3, 3, "Keep newer value");
    replay(&mut wb, &add, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(3, 3), "Keep newer value");
    wb.clear_cell_tracked(0, 3, 3);
    replay(&mut wb, &add, true).unwrap();
    wb.active_sheet_mut().set_bold(3, 3, true);
    replay(&mut wb, &add, false).unwrap();
    assert!(wb.active_sheet().get_format(3, 3).bold);
}

#[test]
fn comment_undo_on_a_spill_receiver_preserves_the_generated_value() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(2,1)");
    let was_absent = wb.active_sheet().get_cell_opt(1, 0).is_none();
    let mut add = action(1, 0, None, note("Spilled value"));
    if let UndoAction::Comments { patches, .. } = &mut add {
        patches[0].remove_cell_on_undo = was_absent;
    }
    replay(&mut wb, &add, true).unwrap();
    replay(&mut wb, &add, false).unwrap();
    assert_eq!(wb.active_sheet().get_cell_opt(1, 0).is_none(), was_absent);
    assert_eq!(wb.active_sheet().get_display(1, 0), "2");
    assert_eq!(wb.active_sheet().get_spill_parent(1, 0), Some((0, 0)));
}
