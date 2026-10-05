use visigrid_engine::{
    filter::SortDirection,
    named_range::NamedRangeTarget,
    sheet::{Sheet, SheetId},
    table::TableRange,
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};

fn filtered() -> Workbook {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 30, 8, "Data")], 0);
    wb.set_cell_value_tracked(0, 2, 1, "Amount");
    wb.set_cell_value_tracked(0, 3, 1, "30");
    wb.set_cell_value_tracked(0, 4, 1, "10");
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 4,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(SheetId(1), Some(spec)).unwrap();
    wb
}

#[test]
fn sheet_lifecycle_add_recalculates_and_replays_without_reusing_identity() {
    let mut wb = filtered();
    wb.set_cell_value_tracked(0, 9, 0, "=IFERROR(New!A1+1,99)");
    let (mut added, commit) = wb.prepare_sheet_add(Some("New")).unwrap();
    let new_id = added.sheets()[1].id;
    assert_eq!(added.sheet(0).unwrap().get_display(9, 0), "1");
    assert_eq!(
        added.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec()
    );
    assert!(!commit.is_empty());
    assert!(wb.capture_guarded_batch(&added).is_err());
    added.set_active_sheet(1);
    commit.replay(&mut added, true).unwrap();
    assert_eq!(added.sheet_count(), 1);
    assert_eq!(added.active_sheet_id(), SheetId(1));
    assert_eq!(added.sheet(0).unwrap().get_display(9, 0), "99");
    let (fresh, _) = added.prepare_sheet_add(None).unwrap();
    assert!(fresh.sheets()[1].id.0 > new_id.0);
    commit.replay(&mut added, false).unwrap();
    assert_eq!(added.sheets()[1].id, new_id);
    added.set_cell_value_tracked(1, 0, 0, "7");
    assert_eq!(added.sheet(0).unwrap().get_display(9, 0), "8");
    assert!(commit.replay(&mut added, true).is_err());
    assert_eq!(added.sheet_count(), 2);
}

#[test]
fn sheet_lifecycle_delete_rewrites_refs_and_reindexes_names_with_exact_undo() {
    let mut wb = filtered();
    let summary = wb.add_sheet_named("Summary").unwrap();
    let summary_id = wb.sheets()[summary].id;
    wb.set_cell_value_tracked(summary, 0, 0, "7");
    wb.define_name_for_cell("KeepMe", summary, 0, 0).unwrap();
    wb.define_name_for_cell("DeletedInput", 0, 3, 1).unwrap();
    let formula = "=( Data!B4 + SUM(Sales[Amount]) )*2";
    wb.set_cell_value_tracked(summary, 1, 0, formula);
    wb.set_cell_value_tracked(summary, 2, 0, "=KeepMe+1");
    wb.set_cell_value_tracked(summary, 3, 0, "=DeletedInput");
    wb.create_table(
        summary_id,
        TableRange {
            start_row: 5,
            start_col: 1,
            end_row: 6,
            end_col: 1,
        },
        "Kept",
    )
    .unwrap();
    wb.set_active_sheet(summary);
    let (mut deleted, commit) = wb.prepare_sheet_delete(SheetId(1)).unwrap();
    assert_eq!(deleted.active_sheet_id(), summary_id);
    assert_eq!(deleted.sheet_count(), 1);
    assert_eq!(
        deleted.active_sheet().get_raw(1, 0),
        "=(#REF!+SUM(#REF!))*2"
    );
    assert_eq!(deleted.active_sheet().get_display(1, 0), "#REF!");
    assert_eq!(deleted.active_sheet().get_display(2, 0), "8");
    assert!(deleted
        .active_sheet()
        .get_display(3, 0)
        .starts_with("#NAME?"));
    assert!(deleted.get_named_range("DeletedInput").is_none());
    assert!(matches!(
        deleted.get_named_range("KeepMe").unwrap().target,
        NamedRangeTarget::Cell { sheet: 0, .. }
    ));
    let recreated = deleted.prepare_sheet_add(Some("Data")).unwrap().0;
    assert_eq!(recreated.sheet(0).unwrap().get_display(1, 0), "#REF!");
    commit.replay(&mut deleted, true).unwrap();
    assert_eq!(deleted.sheet_count(), 2);
    assert_eq!(deleted.active_sheet_id(), summary_id);
    assert_eq!(deleted.sheet(summary).unwrap().get_raw(1, 0), formula);
    assert_eq!(deleted.sheet(summary).unwrap().get_display(1, 0), "140");
    assert!(deleted
        .define_name_for_cell("Sales", summary, 0, 0)
        .is_err());
    assert!(deleted.define_name_for_cell("Kept", summary, 0, 0).is_err());
    assert_eq!(
        deleted.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec()
    );
    assert!(matches!(
        deleted.get_named_range("KeepMe").unwrap().target,
        NamedRangeTarget::Cell { sheet: 1, .. }
    ));
    commit.replay(&mut deleted, false).unwrap();
    assert_eq!(deleted.active_sheet().get_display(1, 0), "#REF!");
}

#[test]
fn sheet_lifecycle_validates_add_and_delete_spills_on_other_sheets() {
    let mut wb = filtered();
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IFERROR(New!A1+6,1))");
    assert!(wb.prepare_sheet_add(Some("New")).is_err());
    assert_eq!(wb.sheet_count(), 1);
    wb.clear_cell_tracked(0, 0, 0);
    let input = wb.add_sheet_named("Input").unwrap();
    wb.set_cell_value_tracked(input, 0, 0, "1");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IFERROR(Input!A1,6))");
    assert!(wb.prepare_sheet_delete(wb.sheets()[input].id).is_err());
    assert_eq!(wb.sheet_count(), 2);
    assert_eq!(
        wb.sheet(0).unwrap().get_raw(0, 0),
        "=SEQUENCE(IFERROR(Input!A1,6))"
    );
}

#[test]
fn sheet_lifecycle_refuses_invalid_last_recovery_and_unparsed_changes() {
    let mut wb = filtered();
    assert!(wb.prepare_sheet_delete(SheetId(1)).is_err());
    for name in ["", "data", "12345678901234567890123456789012"] {
        assert!(wb.prepare_sheet_add(Some(name)).is_err());
    }
    let second = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(second, 0, 0, "=Data!A1+@");
    let revision = wb.revision();
    assert!(wb.prepare_sheet_delete(SheetId(1)).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet_count(), 2);
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(wb.prepare_sheet_add(None).is_err());
    assert!(wb.prepare_sheet_delete(wb.sheets()[second].id).is_err());
}

#[test]
fn sheet_lifecycle_keeps_pivot_definitions_when_removing_output_and_refuses_source_loss() {
    use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotField, PivotValueField};
    let mut wb = filtered();
    let table = wb.active_sheet().tables()[0].clone();
    let source = wb.table_pivot_source(table.id).unwrap();
    let (pivot, output) = wb
        .create_pivot(
            source,
            PivotDefinition {
                rows: vec![],
                column: None,
                values: vec![PivotValueField {
                    field: PivotField {
                        column_id: Some(table.columns[0].id),
                        offset: 0,
                        header: "Amount".into(),
                    },
                    aggregation: Aggregation::Sum,
                    number_format: None,
                }],
            },
        )
        .unwrap();
    assert!(wb
        .prepare_sheet_delete(SheetId(1))
        .unwrap_err()
        .contains("PivotTable"));
    let output_id = wb.sheets()[output].id;
    let output_value = wb.sheets()[output].get_display(1, 0);
    let (mut candidate, commit) = wb.prepare_sheet_delete(output_id).unwrap();
    assert!(candidate.find_pivot(pivot).is_none());
    commit.replay(&mut candidate, true).unwrap();
    assert_eq!(candidate.find_pivot(pivot).unwrap().1.source, source);
    assert_eq!(
        candidate.sheet_by_id(output_id).unwrap().get_display(1, 0),
        output_value
    );
    candidate.refresh_pivot(pivot).unwrap();
}

#[test]
fn sheet_lifecycle_handles_quoted_sheet_names_and_literal_text() {
    let mut wb = Workbook::new();
    let target = wb.add_sheet_named("O'Brien").unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "='O''Brien'!A1&\"O'Brien!A1\"");
    let id = wb.sheets()[target].id;
    let (candidate, _) = wb.prepare_sheet_delete(id).unwrap();
    assert_eq!(candidate.sheets()[0].get_raw(0, 0), "=#REF!&\"O'Brien!A1\"");
    wb.set_cell_value_tracked(0, 1, 0, "='O''Brien'!A1+@");
    assert!(wb.prepare_sheet_delete(id).is_err());
}
