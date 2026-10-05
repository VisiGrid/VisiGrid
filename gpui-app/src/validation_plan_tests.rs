use super::*;
use crate::{
    history::History,
    table_edit::tests::fixture,
    validation_state::{ValidationDialogState, ValidationTypeOption},
};
use visigrid_engine::validation::{ListSource, NumericConstraint, ValidationRule, ValidationType};

fn rule() -> ValidationRule {
    ValidationRule::list_inline(vec!["30".into(), "40".into()])
}
fn selected(wb: &Workbook) -> Draft {
    let sheet = wb.active_sheet();
    let view = sheet.build_saved_table_view(sheet.rows).unwrap().unwrap();
    let (ranges, anchor) = crate::cond_format_ui::plan::metadata_targets(
        sheet,
        view.rows(),
        Some(&sheet.manual_hidden_rows()),
        None,
        &[((3, 2), (6, 2))],
        "validation",
    )
    .unwrap();
    Draft::new(wb, ranges, anchor).unwrap()
}
fn action(commit: Commit) -> UndoAction {
    UndoAction::ValidationChanged {
        sheet_index: 0,
        commit: Box::new(commit),
        description: "Set validation".into(),
    }
}

#[test]
fn filtered_targets_preserve_hidden_rules_values_formulas_and_criteria() {
    let mut wb = fixture(true);
    let old = ValidationRule::list_inline(vec!["10".into()]);
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(3, 2, 6, 2), old.clone());
    let draft = selected(&wb);
    assert_eq!(draft.anchor, (5, 2));
    assert_eq!(range_summary(&draft.ranges), "C4, C6:C7");
    let cells: Vec<_> = wb
        .active_sheet()
        .cells_iter()
        .map(|((r, c), _)| (r, c, wb.active_sheet().get_raw(r, c)))
        .collect();
    let spec = wb.active_sheet().table_view_spec().cloned();
    let original = wb.active_sheet().validations.clone();
    let rev = wb.revision();
    let commit = draft
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .unwrap();
    assert_eq!(wb.revision(), rev);
    assert_eq!(wb.active_sheet().validations, original);
    commit.apply(&mut wb, true).unwrap();
    assert_eq!(wb.revision(), rev + 1);
    assert_eq!(wb.active_sheet().validations.get(4, 2), Some(&old));
    for row in [3, 5, 6] {
        assert_eq!(wb.active_sheet().validations.get(row, 2), Some(&rule()));
    }
    for (r, c, raw) in cells {
        assert_eq!(wb.active_sheet().get_raw(r, c), raw);
    }
    assert_eq!(wb.active_sheet().table_view_spec(), spec.as_ref());
    commit.apply(&mut wb, false).unwrap();
    assert_eq!(wb.active_sheet().validations, original);
}

#[test]
fn manual_hides_and_overlapping_selections_resolve_once() {
    let wb = fixture(true);
    let sheet = wb.active_sheet();
    let rows = sheet.build_saved_table_view(30).unwrap().unwrap();
    let (ranges, anchor) = crate::cond_format_ui::plan::metadata_targets(
        sheet,
        rows.rows(),
        Some(&[5].into()),
        Some(&[1, 3].into()),
        &[((3, 1), (6, 3)), ((3, 2), (6, 2))],
        "validation",
    )
    .unwrap();
    assert_eq!(anchor, (3, 2));
    assert_eq!(
        ranges,
        vec![CellRange::single(3, 2), CellRange::single(6, 2)]
    );
    assert!(crate::cond_format_ui::plan::metadata_targets(
        sheet,
        rows.rows(),
        None,
        Some(&[2].into()),
        &[((3, 2), (6, 2))],
        "validation"
    )
    .is_err());
}

#[test]
fn stale_drafts_and_recovery_refuse_before_publishing() {
    let wb = fixture(true);
    let draft = selected(&wb);
    let mut changed = wb.clone();
    changed.bump_revision_for_structure();
    assert!(draft
        .prepare(&changed, ValidationEdit::Set(rule()))
        .unwrap_err()
        .contains("Copy your draft"));
    let mut changed = wb.clone();
    changed
        .active_sheet_mut()
        .validations
        .exclude(CellRange::single(0, 0));
    assert!(draft
        .prepare(&changed, ValidationEdit::Set(rule()))
        .is_err());
    let mut changed = wb.clone();
    changed.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(draft
        .prepare(&changed, ValidationEdit::Set(rule()))
        .is_err());
    let replaced = Workbook::from_sheets(vec![Sheet::new(SheetId(99), 30, 8)], 0);
    assert!(draft
        .prepare(&replaced, ValidationEdit::Set(rule()))
        .is_err());
}

#[test]
fn partial_exclusion_edits_and_clear_preserve_hidden_records() {
    let mut wb = fixture(true);
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(3, 2, 6, 2), rule());
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::new(3, 2, 6, 2));
    let before = wb.active_sheet().validations.clone();
    for edit in [ValidationEdit::ClearExclusions, ValidationEdit::Clear] {
        let commit = selected(&wb).prepare(&wb, edit).unwrap().unwrap();
        commit.apply(&mut wb, true).unwrap();
        assert!(wb.active_sheet().validations.is_excluded(4, 2));
        commit.apply(&mut wb, false).unwrap();
        assert_eq!(wb.active_sheet().validations, before);
    }
    assert!(selected(&wb)
        .prepare(&wb, ValidationEdit::Exclude)
        .unwrap()
        .is_none());
    assert!(selected(&wb)
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .is_none());
}

#[test]
fn adjacent_body_is_refused_but_header_and_other_sheet_metadata_are_allowed() {
    let wb = fixture(true);
    assert!(Draft::new(&wb, vec![CellRange::single(3, 0)], (3, 0))
        .unwrap_err()
        .contains("beside"));
    assert!(Draft::new(&wb, vec![CellRange::single(2, 1)], (2, 1)).is_ok());
    assert!(Draft::new(&wb, vec![CellRange::single(29, 7)], (29, 7)).is_ok());
    assert!(Draft::new(&wb, vec![CellRange::single(30, 7)], (30, 7)).is_err());
    let mut sheets = wb.sheets().to_vec();
    sheets.push(Sheet::new(SheetId(88), 30, 8));
    let other = Workbook::from_sheets(sheets, 1);
    assert!(other.has_table_criteria());
    assert!(Draft::new(&other, vec![CellRange::single(3, 0)], (3, 0)).is_ok());
}

#[test]
fn grouped_validation_preflight_tracks_sequence_and_refuses_a_late_failure() {
    let base = fixture(true);
    let mut wb = base.clone();
    let first = selected(&wb)
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .unwrap();
    first.apply(&mut wb, true).unwrap();
    let second = selected(&wb)
        .prepare(&wb, ValidationEdit::Exclude)
        .unwrap()
        .unwrap();
    second.apply(&mut wb, true).unwrap();
    let group = UndoAction::Group {
        actions: vec![action(first.clone()), action(second.clone())],
        description: "Two edits".into(),
    };
    validate_history(&base, &group, true).unwrap();
    validate_history(&wb, &group, false).unwrap();
    assert!(crate::table_command_scope::metadata_history_allowed(
        &wb, &group
    ));
    let bad = UndoAction::Group {
        actions: vec![action(first.clone()), action(first)],
        description: "Stale second step".into(),
    };
    let snapshot = base.active_sheet().validations.clone();
    assert!(validate_history(&base, &bad, true).is_err());
    assert_eq!(base.active_sheet().validations, snapshot);
    let mut history = History::new();
    history.record_action_with_provenance(bad, None);
    assert!(history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .is_err());
    // A legacy structural/value group cannot change metadata ahead of the guard.
    let mixed = UndoAction::Group {
        actions: vec![
            action(second),
            UndoAction::Values {
                sheet_index: 0,
                changes: vec![],
            },
        ],
        description: "Mixed".into(),
    };
    assert!(validate_history(&wb, &mixed, false).is_err());
}

#[test]
fn replay_and_rewind_keep_targets_after_criteria_changes_and_save_reopen() {
    let base = fixture(true);
    let mut wb = base.clone();
    let commit = selected(&wb)
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .unwrap();
    let edit = action(commit.clone());
    let mut history = History::new();
    history.record_action_with_provenance(edit.clone(), None);
    let preview = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap()
        .workbook;
    commit.apply(&mut wb, true).unwrap();
    assert_eq!(
        preview.active_sheet().validations,
        wb.active_sheet().validations
    );
    let loaded =
        visigrid_io::json::import_any(&visigrid_io::json::export_workbook(&wb, &[], 0).unwrap())
            .unwrap()
            .0;
    assert_eq!(
        loaded.active_sheet().validations,
        wb.active_sheet().validations
    );
    wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
    validate_history(&wb, &edit, false).unwrap();
    commit.apply(&mut wb, false).unwrap();
    assert!(wb.active_sheet().validations.is_empty());
    commit.apply(&mut wb, true).unwrap();
    assert!(wb.active_sheet().validations.get(4, 2).is_none());
    assert_eq!(wb.active_sheet().validations.get(5, 2), Some(&rule()));
    assert_eq!(
        history
            .build_workbook_before(0, Some(&base), 100, 10_000)
            .unwrap()
            .workbook
            .active_sheet()
            .validations,
        base.active_sheet().validations
    );
}

#[test]
fn validation_history_uses_stable_sheet_identity() {
    let wb = fixture(true);
    let commit = selected(&wb)
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .unwrap();
    let mut moved = Workbook::from_sheets(
        vec![Sheet::new(SheetId(99), 30, 8), wb.active_sheet().clone()],
        0,
    );
    commit.apply(&mut moved, true).unwrap();
    assert!(moved.active_sheet().validations.is_empty());
    assert_eq!(moved.sheet(1).unwrap().validations.get(5, 2), Some(&rule()));
    commit.apply(&mut moved, false).unwrap();
    assert!(moved.sheet(1).unwrap().validations.is_empty());
}

#[test]
fn imported_rule_types_and_ambiguous_list_sources_survive_unchanged_apply() {
    let originals = [
        ValidationRule::new(ValidationType::Custom("=LEN(A1)>0".into())),
        ValidationRule::new(ValidationType::Date(NumericConstraint::between(1.0, 20.0))),
        ValidationRule::list_inline(vec!["Yes".into()]),
        ValidationRule::list_inline(vec![]),
        ValidationRule::new(ValidationType::Decimal(NumericConstraint::between(
            visigrid_engine::validation::ConstraintValue::Formula("SUM(A1:A2)".into()),
            visigrid_engine::validation::ConstraintValue::CellRef("  Limits!B2  ".into()),
        ))),
        ValidationRule::list_inline(vec!["a,b".into(), "  日 本 語  ".into()]),
        ValidationRule::new(ValidationType::List(ListSource::NamedRange(
            "Choices".into(),
        ))),
    ];
    for mut original in originals {
        original.input_message = Some(visigrid_engine::validation::InputMessage::new(
            "Keep this",
            "  日 本 語  ",
        ));
        original.error_alert = Some(visigrid_engine::validation::ErrorAlert::warning(
            "Warning",
            "Custom error",
        ));
        let mut state = ValidationDialogState::default();
        state.load_from_rule(&original);
        assert_eq!(state.build_rule().unwrap(), Some(original.clone()));
        state.ignore_blank = !original.ignore_blank;
        let mut expected = original.clone();
        expected.ignore_blank = !expected.ignore_blank;
        assert_eq!(state.build_rule().unwrap(), Some(expected));
        state.validation_type = ValidationTypeOption::AnyValue;
        state.type_changed = true;
        assert!(state.build_rule().unwrap().is_none());
    }
}

#[test]
fn opening_an_existing_rule_keeps_its_captured_targets_and_private_draft() {
    let mut wb = fixture(true);
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(3, 2, 6, 2), rule());
    let before = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let draft = selected(&wb);
    let expected = draft.ranges.clone();
    let mut state = ValidationDialogState::default();
    state.open_draft(draft, Some(&rule()), true);
    assert_eq!(state.target_range, expected.first().copied());
    assert!(state.has_existing_validation);
    state.ignore_blank = false;
    let commit = state
        .draft
        .as_ref()
        .unwrap()
        .prepare(
            &wb,
            ValidationEdit::Set(state.build_rule().unwrap().unwrap()),
        )
        .unwrap()
        .unwrap();
    assert_eq!(
        visigrid_io::json::export_workbook(&wb, &[], 0).unwrap(),
        before
    );
    commit.apply(&mut wb, true).unwrap();
    assert!(
        !wb.active_sheet()
            .validations
            .get(5, 2)
            .unwrap()
            .ignore_blank
    );
    assert!(
        wb.active_sheet()
            .validations
            .get(4, 2)
            .unwrap()
            .ignore_blank
    );
    state.reset();
    assert!(state.draft.is_none());

    // An excluded imported rule is still editable, not an implicit Clear.
    let imported = ValidationRule::new(ValidationType::Custom("=LEN(C6)>0".into()));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::single(5, 2), imported.clone());
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(5, 2));
    let draft = Draft::new(&wb, vec![CellRange::single(5, 2)], (5, 2)).unwrap();
    let original = draft.anchor_rule();
    assert_eq!(original, Some(imported.clone()));
    state.open_draft(draft, original.as_ref(), true);
    assert!(state.anchor_excluded);
    assert!(state.preserves_imported_type());
    assert_eq!(state.build_rule().unwrap(), Some(imported));
    assert!(state
        .draft
        .as_ref()
        .unwrap()
        .prepare(
            &wb,
            ValidationEdit::Set(state.build_rule().unwrap().unwrap())
        )
        .unwrap()
        .is_none());
}

#[test]
fn validation_dropdown_targets_and_writes_follow_the_stored_record() {
    let mut wb = fixture(true);
    wb.active_sheet_mut().validations.set(
        CellRange::single(3, 1),
        ValidationRule::list_inline(vec!["Other list".into()]),
    );
    wb.active_sheet_mut().validations.set(
        CellRange::single(5, 1),
        ValidationRule::list_inline(vec!["East".into(), "West".into()]),
    );
    let view = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    // Filtering retains view slots; row 3 is a filtered-out slot, not the
    // first visible record. Capture the displayed position of stored row 5.
    let selected = (view.rows().data_to_view(5).unwrap(), 1);
    let target = DropdownTarget::capture(&wb, view.rows(), selected);
    assert_eq!(target.cell, (5, 1));
    assert!(target.is_current(&wb, view.rows(), selected));
    assert!(
        !DropdownTarget::capture(&wb, view.rows(), (3, 1)).is_current(&wb, view.rows(), (3, 1))
    );
    assert!(!target.is_current(&wb, view.rows(), (5, 1)));
    assert_eq!(
        wb.get_list_items(0, target.cell.0, target.cell.1)
            .unwrap()
            .items,
        vec!["East", "West"]
    );
    let after = crate::table_edit::prepare_table_writes(
        &wb,
        0,
        &[crate::table_edit::TableCellWrite::value(
            target.cell.0,
            target.cell.1,
            "East".into(),
        )],
    )
    .unwrap();
    assert_eq!(after.active_sheet().get_raw(5, 1), "East");
    assert_eq!(after.active_sheet().get_raw(3, 1), "West");
    assert_eq!(after.active_sheet().get_raw(4, 1), "East");
    assert!(after
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .data_to_view(5)
        .is_none());
    assert!(!target.is_current(&after, view.rows(), selected));
    wb.bump_revision_for_structure();
    assert!(!target.is_current(&wb, view.rows(), selected));
}

#[test]
fn native_reopen_preserves_edited_validation_and_totals_independently() {
    let mut wb = fixture(true);
    let table = wb.active_sheet().tables().iter().next().unwrap().id;
    wb.set_table_totals_visible(table, true, Default::default())
        .unwrap();
    let commit = selected(&wb)
        .prepare(&wb, ValidationEdit::Set(rule()))
        .unwrap()
        .unwrap();
    commit.apply(&mut wb, true).unwrap();
    let excluded = selected(&wb)
        .prepare(&wb, ValidationEdit::Exclude)
        .unwrap()
        .unwrap();
    excluded.apply(&mut wb, true).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("validation.sheet");
    visigrid_io::native::save_workbook(&wb, &path).unwrap();
    let loaded = visigrid_io::native::load_workbook(&path).unwrap();
    assert_eq!(
        loaded.active_sheet().validations,
        wb.active_sheet().validations
    );
    assert_eq!(
        loaded.active_sheet().get_raw(0, 1),
        wb.active_sheet().get_raw(0, 1)
    );
    assert_eq!(
        loaded.active_sheet().get_raw(7, 1),
        wb.active_sheet().get_raw(7, 1)
    );
    assert_eq!(loaded.active_sheet().tables(), wb.active_sheet().tables());
    assert_eq!(
        loaded.active_sheet().table_view_spec(),
        wb.active_sheet().table_view_spec()
    );
}

#[test]
fn visible_failure_navigation_uses_display_order_and_skips_hidden_cells() {
    let wb = fixture(true);
    let rows = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    let failures = [(3, 2), (4, 2), (5, 2), (6, 2), (5, 3)];
    // First displayed West record is stored row 5, then row 3, then row 6.
    let a = failure_target(
        rows.rows(),
        None,
        Some(&[3].into()),
        &failures,
        (2, 2),
        false,
    )
    .unwrap();
    assert_eq!(a.0, 2);
    assert_eq!(a.2, 1);
    assert_eq!(a.3, 3);
    let b = failure_target(rows.rows(), None, None, &failures, a.1, false).unwrap();
    assert_eq!(b.0, 4);
    let c = failure_target(
        rows.rows(),
        Some(&[5].into()),
        None,
        &failures,
        (2, 2),
        false,
    )
    .unwrap();
    assert_eq!(c.0, 0);
    assert!(failure_target(
        rows.rows(),
        Some(&[3, 5, 6].into()),
        None,
        &failures,
        (2, 2),
        false
    )
    .is_none());
    let last = failure_target(
        rows.rows(),
        None,
        Some(&[3].into()),
        &failures,
        (2, 2),
        true,
    )
    .unwrap();
    assert_eq!(last.0, 3);
}
