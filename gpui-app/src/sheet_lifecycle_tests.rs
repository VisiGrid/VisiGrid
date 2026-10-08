use crate::{
    history::{History, UndoAction},
    table_edit::tests::fixture,
};
use visigrid_engine::{
    table::TableTotal,
    validation::{CellRange, ValidationRule},
};

#[test]
fn sheet_lifecycle_preserves_rules_criteria_and_history_across_input_deletion() {
    let mut wb = fixture(true);
    let inputs = wb.add_sheet_named("Inputs").unwrap();
    let input_id = wb.sheets()[inputs].id;
    wb.set_cell_value_tracked(inputs, 0, 0, "2");
    let table = wb.active_sheet().tables()[0].id;
    wb.set_calculated_column(table, 3, 3, "=[@Amount]*Inputs!$A$1", true)
        .unwrap();
    wb.set_table_totals_visible(table, true, Default::default())
        .unwrap();
    wb.set_table_total(
        table,
        3,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=SUM(Sales[Result])*Inputs!A1".into()),
            label: None,
        },
    )
    .unwrap();
    let region = CellRange {
        start_row: 15,
        start_col: 1,
        end_row: 15,
        end_col: 1,
    };
    wb.active_sheet_mut()
        .validations
        .set(region, ValidationRule::list_range("Inputs!A1"));
    let base = wb.clone();
    let (candidate, delete) = wb.prepare_sheet_delete(input_id).unwrap();
    assert!(candidate.table(table).unwrap().1.columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("#REF!"));
    assert_eq!(candidate.active_sheet().get_display(4, 3), "#REF!"); // hidden row
    assert_eq!(
        candidate.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
    let mut history = History::new();
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(delete.clone()),
        description: "Delete Inputs".into(),
    });
    let (added, add) = candidate.prepare_sheet_add(Some("New Input")).unwrap();
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(add.clone()),
        description: "Add new input".into(),
    });
    let replay = history
        .build_workbook_before(2, Some(&base), 100, 10_000)
        .unwrap();
    assert_eq!(replay.workbook.sheet_names(), added.sheet_names());
    assert_eq!(
        replay.workbook.active_sheet().get_raw(4, 3),
        added.active_sheet().get_raw(4, 3)
    );
    wb.restore_snapshot_monotonic(&added);
    add.replay(&mut wb, true).unwrap();
    delete.replay(&mut wb, true).unwrap();
    assert_eq!(wb.sheets()[inputs].id, input_id);
    assert_eq!(wb.active_sheet().get_display(4, 3), "20");
    assert_eq!(
        wb.active_sheet().validations.get(15, 1),
        base.active_sheet().validations.get(15, 1)
    );
    delete.replay(&mut wb, false).unwrap();
    add.replay(&mut wb, false).unwrap();
    assert_eq!(wb.sheet_names(), added.sheet_names());
}

#[test]
fn sheet_lifecycle_roundtrips_deleted_references_and_remaining_tables() {
    let mut wb = fixture(true);
    let before = wb.add_sheet_named("Inputs").unwrap();
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(before, 0, 0, "2");
    wb.set_cell_value_tracked(summary, 0, 0, "=Inputs!A1+1");
    wb.set_cell_value_tracked(summary, 1, 0, "9");
    wb.define_name_for_cell("SummaryValue", summary, 1, 0)
        .unwrap();
    let (candidate, _) = wb.prepare_sheet_delete(wb.sheets()[before].id).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let native = dir.path().join("deleted.sheet");
    visigrid_io::native::save_workbook(&candidate, &native).unwrap();
    let native = visigrid_io::native::load_workbook(&native).unwrap();
    let json = visigrid_io::json::export_workbook(&candidate, &[], 0).unwrap();
    let json = visigrid_io::json::import_any(&json).unwrap().0;
    let path = dir.path().join("deleted.xlsx");
    visigrid_io::xlsx::export_with_order(
        &candidate,
        &path,
        None,
        visigrid_io::xlsx::ExportOrder::Stored,
    )
    .unwrap();
    let (xlsx, _) = visigrid_io::xlsx::import(&path).unwrap();
    for loaded in [native, json, xlsx] {
        assert_eq!(loaded.sheet_names(), candidate.sheet_names());
        assert_eq!(loaded.sheets()[1].get_display(0, 0), "#REF!");
        assert_eq!(
            loaded.get_named_range("SummaryValue"),
            candidate.get_named_range("SummaryValue")
        );
        assert!(loaded.sheets()[0].table_view_spec().is_some());
        assert_eq!(loaded.sheets()[0].tables().len(), 1);
    }
}

#[test]
fn sheet_lifecycle_view_remapping_keeps_surviving_panes_and_resets_deleted_ones() {
    let mut wb = fixture(true);
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_active_sheet(other);
    let ids: Vec<_> = wb.sheets().iter().map(|s| s.id).collect();
    let (candidate, _) = wb.prepare_sheet_delete(ids[0]).unwrap();
    let mut surviving = crate::workbook_view::WorkbookViewState::default();
    surviving.active_sheet = other;
    surviving.selected = (12, 4);
    surviving.zoom_level = 1.5;
    crate::table_batch::remap_sheet_view(&mut surviving, &ids, &candidate);
    assert_eq!(surviving.active_sheet, 0);
    assert_eq!(surviving.selected, (12, 4));
    let mut removed = crate::workbook_view::WorkbookViewState::default();
    removed.selected = (8, 2);
    removed.zoom_level = 1.5;
    crate::table_batch::remap_sheet_view(&mut removed, &ids, &candidate);
    assert_eq!(removed.active_sheet, 0);
    assert_eq!(removed.selected, (0, 0));
    assert_eq!(removed.zoom_level, 1.5);
}

#[test]
fn sheet_lifecycle_rewind_keeps_prior_view_history_on_its_sheet() {
    use visigrid_engine::{
        sheet::{Sheet, SheetId},
        workbook::Workbook,
    };
    let base = Workbook::from_sheets(
        vec![
            Sheet::new_with_name(SheetId(10), 10, 8, "First"),
            Sheet::new_with_name(SheetId(11), 10, 8, "Survivor"),
        ],
        1,
    );
    let mut history = History::new();
    let mut order: Vec<_> = (0..10).collect();
    order.swap(0, 1);
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::SortApplied {
        sheet_index: 1,
        previous_row_order: (0..10).collect(),
        previous_sort_state: None,
        new_row_order: order.clone(),
        new_sort_state: (0, true),
    });
    let (_, delete) = base.prepare_sheet_delete(SheetId(10)).unwrap();
    history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(delete),
        description: "Delete first sheet".into(),
    });
    let replay = history
        .build_workbook_before(2, Some(&base), 100, 10_000)
        .unwrap();
    assert_eq!(replay.workbook.active_sheet_id(), SheetId(11));
    assert_eq!(replay.view_state.per_sheet.len(), 1);
    assert_eq!(replay.view_state.per_sheet[0].row_order, Some(order));
    assert_eq!(replay.view_state.per_sheet[0].sort, Some((0, true)));
}
