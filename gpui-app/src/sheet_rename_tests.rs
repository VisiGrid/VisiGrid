use super::SheetRenameDraft;
use crate::{
    history::{History, UndoAction},
    table_edit::tests::fixture,
};
use visigrid_engine::{
    cond_format::CondStyle,
    table::TableTotal,
    validation::{CellRange, ListSource, ValidationRule, ValidationType},
};

#[test]
fn rename_rewrites_hidden_calculated_rules_totals_and_validation_with_history() {
    let mut wb = fixture(true);
    let inputs = wb.add_sheet_named("Inputs").unwrap();
    wb.set_cell_value_tracked(inputs, 0, 0, "2");
    wb.define_name_for_cell("InputRate", inputs, 0, 0).unwrap();
    let id = wb.active_sheet().tables()[0].id;
    wb.set_calculated_column(id, 3, 3, "=[@Amount]*Inputs!$A$1", true)
        .unwrap();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        id,
        3,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=SUM(Sales[Result])*Inputs!A1".into()),
            label: None,
        },
    )
    .unwrap();
    wb.set_cell_value_tracked(0, 14, 1, "=SUM(Sales[[#Totals],[Result]])");
    let total = wb.active_sheet().get_display(14, 1);
    wb.set_cell_value_tracked(0, 15, 1, "=Inputs!A1+InputRate");
    wb.set_cell_text_tracked(0, 16, 1, "cached");
    let mut frozen = wb.active_sheet().get_cell(16, 1);
    frozen.set_frozen_formula(Some("=Inputs!A1".into()));
    wb.restore_cell_tracked(0, 16, 1, Some(frozen)).unwrap();
    wb.set_cell_value_tracked(0, 17, 1, "=INDIRECT(\"Inputs!A1\")");
    let region = CellRange {
        start_row: 10,
        start_col: 1,
        end_row: 10,
        end_col: 1,
    };
    wb.active_sheet_mut().cond_formats.add(
        vec![region],
        "=Inputs!A1>0",
        CondStyle::Inline(Default::default()),
    );
    let mut rule = ValidationRule::list_range("Inputs!A1");
    rule.rule_type = ValidationType::List(ListSource::NamedRange("OFFSET(Inputs!A1,0,0)".into()));
    wb.active_sheet_mut().validations.set(region, rule);
    let base = wb.clone();
    let draft = SheetRenameDraft::capture(&wb, inputs).unwrap();
    let (candidate, commit) = draft.prepare(&wb, "New Inputs").unwrap();
    assert_eq!(wb.sheet(inputs).unwrap().name, "Inputs");
    assert_eq!(candidate.active_sheet_index(), 0);
    assert_eq!(
        candidate.active_sheet().get_raw(4, 3),
        "=[@Amount]*'New Inputs'!$A$1"
    ); // hidden East
    assert_eq!(
        candidate.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=[@Amount]*'New Inputs'!$A$1")
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
        Some("=SUM(Sales[Result])*'New Inputs'!A1")
    );
    assert_eq!(candidate.active_sheet().get_display(14, 1), total);
    assert_eq!(candidate.active_sheet().get_display(15, 1), "4");
    assert_eq!(
        candidate.active_sheet().get_cell(16, 1).frozen_formula(),
        Some("='New Inputs'!A1")
    );
    assert_eq!(
        candidate.active_sheet().get_raw(17, 1),
        "=INDIRECT(\"Inputs!A1\")"
    );
    assert!(candidate
        .active_sheet()
        .get_display(17, 1)
        .starts_with("#REF!"));
    assert_eq!(
        candidate
            .active_sheet()
            .cond_formats
            .iter()
            .next()
            .unwrap()
            .predicate,
        "='New Inputs'!A1>0"
    );
    assert_eq!(
        candidate
            .active_sheet()
            .validations
            .get(10, 1)
            .unwrap()
            .rule_type,
        ValidationType::List(ListSource::NamedRange("OFFSET('New Inputs'!A1,0,0)".into()))
    );
    assert_eq!(
        candidate.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
    assert_eq!(
        candidate.get_named_range("InputRate"),
        base.get_named_range("InputRate")
    );
    wb.restore_snapshot_monotonic(&candidate);
    commit.replay(&mut wb, true).unwrap();
    assert_eq!(wb.sheet(inputs).unwrap().name, "Inputs");
    assert_eq!(wb.active_sheet().get_raw(4, 3), "=[@Amount]*Inputs!$A$1");
    commit.replay(&mut wb, false).unwrap();
    let mut history = History::new();
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(commit),
        description: "Rename Inputs".into(),
    });
    let replay = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap()
        .workbook;
    assert_eq!(replay.sheet(inputs).unwrap().name, "New Inputs");
    assert_eq!(
        replay.active_sheet().get_raw(4, 3),
        "=[@Amount]*'New Inputs'!$A$1"
    );
    // Dormant totals and templates keep the new qualifier when materialized later.
    wb.set_table_totals_visible(id, false, Default::default())
        .unwrap();
    let draft = SheetRenameDraft::capture(&wb, inputs).unwrap();
    let (candidate, _) = draft.prepare(&wb, "Later Inputs").unwrap();
    wb.restore_snapshot_monotonic(&candidate);
    assert!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("'Later Inputs'!A1"));
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(14, 1), total);
    // Dynamic references were verified above. Footer movement still refuses
    // them; remove those independent fixtures before testing template append.
    wb.clear_cell_tracked(0, 17, 1);
    wb.active_sheet_mut().validations.remove(&region);
    wb.append_table_rows(id, 1, &[(7, 1, "West".into()), (7, 2, "5".into())])
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(7, 3), "10");
}

#[test]
fn renamed_table_sheet_roundtrips_with_criteria_and_preserves_names() {
    let wb = fixture(true);
    let name = wb.active_sheet().name.clone();
    let draft = SheetRenameDraft::capture(&wb, 0).unwrap();
    let (candidate, _) = draft.prepare(&wb, "Q4 Sales").unwrap();
    let dir = tempfile::tempdir().unwrap();
    let native = dir.path().join("renamed.sheet");
    visigrid_io::native::save_workbook(&candidate, &native).unwrap();
    let loaded = visigrid_io::native::load_workbook(&native).unwrap();
    assert_eq!(loaded.active_sheet().name, "Q4 Sales");
    assert_eq!(
        loaded.active_sheet().table_view_spec(),
        candidate.active_sheet().table_view_spec()
    );
    let json = visigrid_io::json::export_workbook(&candidate, &[], 0).unwrap();
    let loaded = visigrid_io::json::import_any(&json).unwrap().0;
    assert_eq!(loaded.active_sheet().name, "Q4 Sales");
    let xlsx = dir.path().join("renamed.xlsx");
    visigrid_io::xlsx::export_with_order(
        &candidate,
        &xlsx,
        None,
        visigrid_io::xlsx::ExportOrder::Stored,
    )
    .unwrap();
    let (loaded, report) = visigrid_io::xlsx::import(&xlsx).unwrap();
    assert_eq!(loaded.active_sheet().name, "Q4 Sales");
    assert_eq!(report.tables_imported, 1);
    assert!(loaded.active_sheet().table_view_spec().is_some());
    assert_eq!(wb.active_sheet().name, name);
}

#[test]
fn stale_and_recovery_rename_drafts_refuse_without_changing_labels() {
    let mut wb = fixture(true);
    let draft = SheetRenameDraft::capture(&wb, 0).unwrap();
    wb.set_cell_value_tracked(0, 15, 1, "1");
    assert!(draft.prepare(&wb, "New").is_err());
    let other = wb.add_sheet_named("Other").unwrap();
    let draft = SheetRenameDraft::capture(&wb, 0).unwrap();
    wb.rename_sheet(other, "Changed elsewhere");
    assert!(draft.prepare(&wb, "New").is_err());
    let draft = SheetRenameDraft::capture(&wb, 0).unwrap();
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(draft.prepare(&wb, "New").is_err());
    assert!(SheetRenameDraft::capture(&wb, 0).is_err());
    assert_ne!(wb.active_sheet().name, "New");
}

#[test]
fn renaming_cannot_introduce_an_unsafe_cross_sheet_spill() {
    let mut wb = fixture(true);
    let input = wb.add_sheet_named("Inputs").unwrap();
    wb.set_cell_value_tracked(input, 0, 0, "6");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(New!A1)");
    let draft = SheetRenameDraft::capture(&wb, input).unwrap();
    assert!(draft.prepare(&wb, "New").is_err());
    assert_eq!(wb.sheet(input).unwrap().name, "Inputs");
    assert_eq!(wb.active_sheet().get_raw(0, 0), "=SEQUENCE(New!A1)");
}
