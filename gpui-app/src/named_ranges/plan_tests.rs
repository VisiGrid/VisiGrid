use super::*;
use crate::{
    history::{History, UndoAction},
    table_edit::tests::fixture,
};
use visigrid_engine::{
    cond_format::CondStyle,
    formula::names::rename_reference,
    named_range::NamedRangeTarget,
    table::TableTotal,
    validation::{CellRange, ListSource, ValidationRule, ValidationType},
};

#[test]
fn source_refactoring_changes_only_named_tokens() {
    for (source, old, new, expected) in [
        (
            "=( TaxRate + taxrate ) * 2",
            "TaxRate",
            "Tax",
            "=( Tax + Tax ) * 2",
        ),
        (
            "=TaxRate+\"TaxRate\"+TaxRateSheet!A1+'TaxRate'!A1+Sales[TaxRate]+Other.TaxRate",
            "TaxRate",
            "Tax",
            "=Tax+\"TaxRate\"+TaxRateSheet!A1+'TaxRate'!A1+Sales[TaxRate]+Other.TaxRate",
        ),
        ("=E + 1E3 + 1e-3", "E", "Scale", "=Scale + 1E3 + 1e-3"),
        ("=SUM(A : A)+A", "A", "Alpha", "=SUM(A : A)+Alpha"),
        (
            "=TaxRate+INDIRECT(\"TaxRate\")+\"éTaxRate\"",
            "TaxRate",
            "Tax",
            "=Tax+INDIRECT(\"TaxRate\")+\"éTaxRate\"",
        ),
        ("TaxRate>0", "TaxRate", "Tax", "Tax>0"),
        (
            "=TaxRate+TaxRate.Extra",
            "TaxRate",
            "tAxRaTe",
            "=tAxRaTe+TaxRate.Extra",
        ),
    ] {
        assert_eq!(rename_reference(source, old, new).unwrap(), expected);
    }
    assert!(rename_reference("=TaxRate+@", "TaxRate", "Tax").is_err());
    assert_eq!(
        rename_reference("=UNSUPPORTED(@)", "TaxRate", "Tax").unwrap(),
        "=UNSUPPORTED(@)"
    );
}

#[test]
fn create_selection_uses_canonical_endpoints_and_jump_projects_exact_records() {
    let wb = fixture(true);
    let rows = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .clone();
    let range = selection_range(&wb, &rows, (4, 2), (5, 2)).unwrap();
    assert_eq!(
        range.target,
        NamedRangeTarget::Range {
            sheet: 0,
            start_row: 3,
            end_row: 5,
            start_col: 2,
            end_col: 2
        }
    );
    assert_eq!(
        project_named_range(&rows, (3, 2), (5, 2)),
        vec![((4, 2), (5, 2))]
    );
    assert!(project_named_range(&rows, (4, 2), (4, 2)).is_empty());
    assert!(selection_range(&wb, &rows, (3, 2), (3, 2)).is_err());
    assert!(selection_range(&wb, &rows, (30, 2), (30, 2)).is_err());
    let one = selection_range(&wb, &rows, (4, 2), (4, 2)).unwrap();
    assert_eq!(
        one.target,
        NamedRangeTarget::Cell {
            sheet: 0,
            row: 5,
            col: 2
        }
    );
}

#[test]
fn create_delete_description_recalculate_atomically_and_replay() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "=TaxRate*2");
    let base = wb.clone();
    let range = NamedRange::cell("TaxRate", 0, 3, 2).with_description("Canonical record");
    let (created, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(range.clone()))
        .unwrap();
    assert!(wb.get_named_range("TaxRate").is_none());
    assert_eq!(created.active_sheet().get_display(15, 1), "60");
    wb.restore_snapshot_monotonic(&created);
    commit.replay(&mut wb, true).unwrap();
    assert!(wb.get_named_range("TaxRate").is_none());
    commit.replay(&mut wb, false).unwrap();
    let (described, description) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Description {
            before: range.clone(),
            description: Some("New description".into()),
        })
        .unwrap();
    wb.restore_snapshot_monotonic(&described);
    description.replay(&mut wb, true).unwrap();
    let (_, noop) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Description {
            before: range.clone(),
            description: range.description.clone(),
        })
        .unwrap();
    assert!(noop.is_empty());
    let (deleted, deletion) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Delete(range.clone()))
        .unwrap();
    assert!(deleted.active_sheet().get_display(15, 1).contains("#NAME?"));
    wb.restore_snapshot_monotonic(&deleted);
    deletion.replay(&mut wb, true).unwrap();
    assert_eq!(wb.active_sheet().get_display(15, 1), "60");
    assert_eq!(
        wb.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
    let mut history = History::new();
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(commit),
        description: "Create name".into(),
    });
    let replay = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap()
        .workbook;
    assert_eq!(replay.get_named_range("TaxRate"), Some(&range));
    assert_eq!(replay.active_sheet().get_display(15, 1), "60");
}

#[test]
fn rename_covers_hidden_cells_cross_sheet_rules_totals_validation_and_roundtrips() {
    let mut wb = fixture(true);
    wb.define_name_for_cell("TaxRate", 0, 15, 1).unwrap();
    wb.set_cell_value_tracked(0, 15, 1, "3");
    let id = wb.active_sheet().tables()[0].id;
    wb.set_calculated_column(id, 3, 3, "=[@Amount]*TaxRate", true)
        .unwrap();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        id,
        3,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=SUM(Sales[Result])*TaxRate".into()),
            label: None,
        },
    )
    .unwrap();
    wb.set_table_totals_visible(id, false, Default::default())
        .unwrap();
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=( TaxRate + 1 ) * 2");
    let region = CellRange {
        start_row: 10,
        end_row: 12,
        start_col: 1,
        end_col: 1,
    };
    wb.sheet_mut(other).unwrap().cond_formats.add(
        vec![region],
        "=TaxRate>0",
        CondStyle::Inline(Default::default()),
    );
    let mut validation = ValidationRule::list_range("TaxRate");
    validation.rule_type = ValidationType::List(ListSource::NamedRange("TaxRate".into()));
    wb.sheet_mut(other)
        .unwrap()
        .validations
        .set(region, validation);
    let base = wb.clone();
    let before = wb.get_named_range("TaxRate").unwrap().clone();
    let usages = wb.named_range_usages("TaxRate");
    assert!(usages.iter().any(|(label, _)| label.contains("totals")));
    assert!(usages.iter().any(|(label, _)| label.contains("validation")));
    let (candidate, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Rename {
            before,
            name: "Tax".into(),
        })
        .unwrap();
    assert_eq!(
        candidate.sheet(other).unwrap().get_raw(0, 0),
        "=( Tax + 1 ) * 2"
    );
    assert_eq!(candidate.active_sheet().get_raw(4, 3), "=[@Amount]*Tax"); // hidden East record
    assert_eq!(
        candidate.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=[@Amount]*Tax")
    );
    assert_eq!(
        candidate
            .table(id)
            .unwrap()
            .1
            .totals
            .as_ref()
            .unwrap()
            .columns[2]
            .formula
            .as_deref(),
        Some("=SUM(Sales[Result])*Tax")
    );
    assert!(candidate.named_range_usages("TaxRate").is_empty());
    assert_eq!(candidate.named_range_usages("Tax").len(), usages.len());
    wb.restore_snapshot_monotonic(&candidate);
    commit.replay(&mut wb, true).unwrap();
    assert_eq!(wb.named_range_usages("TaxRate"), usages);
    commit.replay(&mut wb, false).unwrap();
    assert_eq!(
        wb.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("names.sheet");
    visigrid_io::native::save_workbook(&wb, &file).unwrap();
    let native = visigrid_io::native::load_workbook(&file).unwrap();
    assert_eq!(
        native.named_range_usages("Tax"),
        wb.named_range_usages("Tax")
    );
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let loaded = visigrid_io::json::import_any(&json).unwrap().0;
    assert_eq!(
        loaded.named_range_usages("Tax"),
        wb.named_range_usages("Tax")
    );
    let file = dir.path().join("names.xlsx");
    visigrid_io::xlsx::export_with_order(&wb, &file, None, visigrid_io::xlsx::ExportOrder::Stored)
        .unwrap();
    let imported = visigrid_io::xlsx::import(&file).unwrap().0;
    assert_eq!(
        imported.get_named_range("Tax").unwrap().target,
        wb.get_named_range("Tax").unwrap().target
    );
    assert_eq!(imported.sheet(other).unwrap().get_display(0, 0), "8");
    wb.append_table_rows(id, 1, &[(7, 1, "West".into()), (7, 2, "5".into())])
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(7, 3), "15");
}

#[test]
fn name_collisions_stale_drafts_targets_and_replay_refuse_without_partial_changes() {
    let mut wb = fixture(true);
    let range = NamedRange::cell("TaxRate", 0, 3, 2);
    let (after, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(range.clone()))
        .unwrap();
    wb.restore_snapshot_monotonic(&after);
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(range.clone()))
        .is_err());
    let collision = wb.active_sheet().tables()[0].name.clone();
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Rename {
            before: range.clone(),
            name: collision
        })
        .is_err());
    let mut stale = range.clone();
    stale.description = Some("stale".into());
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Delete(stale))
        .is_err());
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(NamedRange::cell(
            "BadTarget",
            0,
            30,
            0
        )))
        .is_err());
    wb.set_cell_value_tracked(0, 20, 1, "Newer edit");
    assert!(commit.replay(&mut wb, true).is_err());
    assert_eq!(wb.get_named_range("TaxRate"), Some(&range));
    assert_eq!(wb.active_sheet().get_raw(20, 1), "Newer edit");
}

#[test]
fn rename_rejects_unparsed_dependencies_without_publishing_other_rewrites() {
    let mut wb = fixture(true);
    wb.define_name_for_cell("TaxRate", 0, 15, 1).unwrap();
    wb.set_cell_value_tracked(0, 15, 1, "3");
    wb.set_cell_value_tracked(0, 16, 1, "=TaxRate+1");
    wb.set_cell_value_tracked(0, 17, 1, "=TaxRate+@");
    let revision = wb.revision();
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Rename {
            before: wb.get_named_range("TaxRate").unwrap().clone(),
            name: "Tax".into()
        })
        .is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_raw(16, 1), "=TaxRate+1");
    assert!(wb.get_named_range("Tax").is_none());
}

#[test]
fn stale_dialogs_and_recovery_never_publish() {
    let mut wb = fixture(true);
    let range = NamedRange::cell("TaxRate", 0, 3, 2);
    let draft = NameDraft {
        revision: wb.revision(),
        range: range.clone(),
    };
    assert_eq!(draft.checked_range(&wb).unwrap(), range);
    wb.set_cell_value_tracked(0, 15, 1, "42");
    assert!(draft.checked_range(&wb).is_err());
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(range))
        .is_err());
    assert!(wb.get_named_range("TaxRate").is_none());
}

#[test]
fn adding_a_name_that_would_spill_into_a_table_view_is_atomic() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "6");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(TaxRate)");
    wb.active_sheet().build_saved_table_view(30).unwrap();
    let before = wb.clone();
    let result = wb.prepare_named_range_edit(&NamedRangeEdit::Create(NamedRange::cell(
        "TaxRate", 0, 15, 1,
    )));
    assert!(result.is_err());
    assert!(wb.get_named_range("TaxRate").is_none());
    assert_eq!(wb.revision(), before.revision());
    assert_eq!(
        wb.active_sheet().get_display(0, 0),
        before.active_sheet().get_display(0, 0)
    );
}

#[test]
fn name_replay_invalidates_unchanged_formula_dependents_without_totals() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "=TaxRate*2");
    let range = NamedRange::cell("TaxRate", 0, 3, 2);
    let (candidate, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(range))
        .unwrap();
    wb.restore_snapshot_monotonic(&candidate);
    let generation = wb.active_sheet().edit_generation();
    commit.replay(&mut wb, true).unwrap();
    assert!(wb.active_sheet().edit_generation() > generation);
    assert!(wb.active_sheet().get_display(15, 1).contains("#NAME?"));
    let generation = wb.active_sheet().edit_generation();
    commit.replay(&mut wb, false).unwrap();
    assert!(wb.active_sheet().edit_generation() > generation);
    assert_eq!(wb.active_sheet().get_display(15, 1), "60");
}

#[test]
fn a_new_name_cannot_destroy_formula_sources_by_introducing_a_cycle() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "=TaxRate+1");
    let revision = wb.revision();
    let error = wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(NamedRange::cell(
            "TaxRate", 0, 15, 1,
        )))
        .unwrap_err();
    assert!(error.contains("cycle"), "{error}");
    assert_eq!(wb.active_sheet().get_raw(15, 1), "=TaxRate+1");
    assert_eq!(wb.revision(), revision);
    assert!(wb.get_named_range("TaxRate").is_none());
}

#[test]
fn case_only_renames_and_exact_noops_preserve_bindings() {
    let mut wb = fixture(true);
    wb.define_name_for_cell("TaxRate", 0, 3, 2).unwrap();
    wb.set_cell_value_tracked(0, 15, 1, "=TaxRate+1");
    let range = wb.get_named_range("TaxRate").unwrap().clone();
    let (_, noop) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Rename {
            before: range.clone(),
            name: "TaxRate".into(),
        })
        .unwrap();
    assert!(noop.is_empty());
    let (candidate, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Rename {
            before: range,
            name: "TAXRATE".into(),
        })
        .unwrap();
    assert_eq!(
        candidate.get_named_range("taxrate").unwrap().name,
        "TAXRATE"
    );
    assert_eq!(candidate.active_sheet().get_raw(15, 1), "=TAXRATE+1");
    assert_eq!(candidate.active_sheet().get_display(15, 1), "31");
    wb.restore_snapshot_monotonic(&candidate);
    commit.replay(&mut wb, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(15, 1), "=TaxRate+1");
}

#[test]
fn definition_changes_rebuild_filter_membership_on_unchanged_formula_sources() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "East");
    wb.set_cell_value_tracked(0, 3, 1, "=IFERROR(TaxRate,\"West\")");
    assert!(wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .is_data_row_visible(3));
    let (candidate, commit) = wb
        .prepare_named_range_edit(&NamedRangeEdit::Create(NamedRange::cell(
            "TaxRate", 0, 15, 1,
        )))
        .unwrap();
    assert!(!candidate
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .is_data_row_visible(3));
    wb.restore_snapshot_monotonic(&candidate);
    commit.replay(&mut wb, true).unwrap();
    assert!(wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .is_data_row_visible(3));
    commit.replay(&mut wb, false).unwrap();
    assert!(!wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap()
        .rows()
        .is_data_row_visible(3));
    assert_eq!(
        wb.active_sheet().get_raw(3, 1),
        "=IFERROR(TaxRate,\"West\")"
    );
}
