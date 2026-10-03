use visigrid_engine::{
    filter::{ColumnFilter, FilterKey},
    pivot::{Aggregation, PivotDefinition, PivotField, PivotSource, PivotValueField},
    sheet::SheetId,
    structural::Axis,
    table::{TableId, TableRange},
    table_view::{TableFilter, TableViewSpec},
    workbook::Workbook,
};

fn fixture() -> (Workbook, TableId, PivotSource, PivotDefinition) {
    let mut wb = Workbook::new();
    for (r, row) in [["Region", "Amount"], ["West", "10"], ["East", "20"]]
        .iter()
        .enumerate()
    {
        for (c, value) in row.iter().enumerate() {
            wb.set_cell_value_tracked(0, r, c, value);
        }
    }
    let id = wb
        .create_table(
            wb.sheet(0).unwrap().id,
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let source = wb.table_pivot_source(id).unwrap();
    let columns = &wb.table(id).unwrap().1.columns;
    let field = |offset: usize| PivotField {
        column_id: Some(columns[offset].id),
        offset: offset as u32,
        header: columns[offset].name.clone(),
    };
    let def = PivotDefinition {
        rows: vec![field(0)],
        column: None,
        values: vec![PivotValueField {
            field: field(1),
            aggregation: Aggregation::Sum,
            number_format: None,
        }],
    };
    (wb, id, source, def)
}

fn total(wb: &Workbook, id: u64) -> String {
    let (si, pivot) = wb.find_pivot(id).unwrap();
    wb.sheet(si).unwrap().get_raw(
        (pivot.anchor_row + pivot.extent.unwrap().0 - 1) as usize,
        pivot.anchor_col as usize + 1,
    )
}

#[test]
fn refresh_follows_growth_rename_and_metadata_only_resize() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    assert_eq!(total(&wb, pivot), "30");
    wb.append_table_rows(id, 1, &[(3, 0, "West".into()), (3, 1, "7".into())])
        .unwrap();
    assert!(wb.is_pivot_stale(wb.find_pivot(pivot).unwrap().1));
    assert_eq!(total(&wb, pivot), "30", "explicit refresh only");
    wb.rename_table(id, "Orders").unwrap();
    wb.rename_table_columns(id, &["Area".into(), "Revenue".into()])
        .unwrap();
    wb.refresh_pivot(pivot).unwrap();
    assert_eq!(total(&wb, pivot), "37");
    let p = wb.find_pivot(pivot).unwrap().1;
    assert_eq!(p.definition.rows[0].header, "Area");
    assert_eq!(p.definition.values[0].field.header, "Revenue");
    assert_eq!(p.last_refresh.as_ref().unwrap().source_rows, 3);
    assert!(!wb.is_pivot_stale(p));
    assert_eq!(wb.pivot_source_growth(p), None);
    let mut range = wb.table(id).unwrap().1.range;
    range.end_row += 1;
    wb.resize_table(id, range).unwrap();
    assert!(wb.is_pivot_stale(wb.find_pivot(pivot).unwrap().1));
}

#[test]
fn filter_does_not_change_membership_or_freshness_and_refresh_rebuilds_views() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    let mut spec = TableViewSpec::new(id);
    spec.filters.push(TableFilter {
        column: wb.table(id).unwrap().1.columns[0].id,
        criteria: ColumnFilter {
            selected: Some(
                [FilterKey::Text("west".into()).normalized()]
                    .into_iter()
                    .collect(),
            ),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(source.sheet_id, Some(spec)).unwrap();
    assert!(!wb.is_pivot_stale(wb.find_pivot(pivot).unwrap().1));
    wb.refresh_pivot(pivot).unwrap();
    assert_eq!(total(&wb, pivot), "30");
    wb.set_cell_value_tracked(0, 2, 1, "40");
    wb.refresh_pivot(pivot).unwrap();
    assert_eq!(total(&wb, pivot), "50", "hidden records are included");
}

#[test]
fn inserted_columns_rebind_by_id_and_removed_column_cannot_rebind_by_name() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    wb.structural_edit(0, Axis::Col, 1, 1, false).unwrap();
    wb.refresh_pivot(pivot).unwrap();
    assert_eq!(total(&wb, pivot), "30");
    assert_eq!(
        wb.find_pivot(pivot).unwrap().1.definition.values[0]
            .field
            .offset,
        2
    );
    wb.structural_edit(0, Axis::Col, 2, 1, true).unwrap();
    let before = wb.find_pivot(pivot).unwrap().1.clone();
    wb.rename_table_columns(id, &["Region".into(), "Amount".into()])
        .unwrap();
    assert!(wb.refresh_pivot(pivot).unwrap_err().contains("missing"));
    assert_eq!(wb.find_pivot(pivot).unwrap().1, &before);
    assert_eq!(total(&wb, pivot), "30");
    assert!(wb.is_pivot_stale(&before));
}

#[test]
fn deleted_table_and_reused_name_do_not_retarget_pivot() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    let range = wb.table(id).unwrap().1.range;
    wb.remove_table(id).unwrap();
    wb.create_table(source.sheet_id, range, "Sales").unwrap();
    assert!(wb
        .refresh_pivot(pivot)
        .unwrap_err()
        .contains("no longer exists"));
    assert_eq!(total(&wb, pivot), "30");
}

#[test]
fn legacy_range_wire_unchanged_table_wire_cannot_be_misread_as_range() {
    #[derive(serde::Deserialize)]
    #[allow(dead_code)]
    struct OldSource {
        sheet_id: SheetId,
        start_row: u32,
        start_col: u32,
        end_row: u32,
        end_col: u32,
    }
    let (_, _, mut source, _) = fixture();
    let json = serde_json::to_string(&source).unwrap();
    assert!(serde_json::from_str::<OldSource>(&json).is_err());
    assert_eq!(serde_json::from_str::<PivotSource>(&json).unwrap(), source);
    source.table_id = None;
    let json = serde_json::to_string(&source).unwrap();
    assert!(serde_json::from_str::<OldSource>(&json).is_ok());
    assert_eq!(serde_json::from_str::<PivotSource>(&json).unwrap(), source);
}

#[test]
fn pivot_placement_through_filtered_body_band_is_atomic() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    let mut p = wb.find_pivot(pivot).unwrap().1.clone();
    p.id = wb.next_pivot_id();
    p.anchor_col = 5;
    p.extent = None;
    wb.set_table_view_spec(source.sheet_id, Some(TableViewSpec::new(id)))
        .unwrap();
    let (p, snapshot, generation) = wb.pivot_snapshot(&p).unwrap();
    let output = visigrid_engine::pivot::aggregate(&p.definition, &snapshot).unwrap();
    let commit = wb
        .prepare_pivot_commit(source.sheet_id, p, &output, generation, 0)
        .unwrap();
    let revision = wb.revision();
    assert!(
        wb.apply_pivot_state(&commit.after).is_err(),
        "a separate rectangle still shares the Table's body rows"
    );
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.pivots().len(), 1);
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 5), "");
    assert!(wb.sheet(0).unwrap().build_saved_table_view(20).is_ok());
}

#[test]
fn header_only_table_refresh_and_generation_checks() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    let p = wb.find_pivot(pivot).unwrap().1.clone();
    let generation = wb.pivot_source_generation(&p).unwrap();
    wb.rename_table(id, "Orders").unwrap();
    assert_ne!(
        wb.pivot_source_generation(&p),
        Some(generation),
        "metadata invalidates in-flight jobs"
    );
    let mut range = wb.table(id).unwrap().1.range;
    range.end_row = range.start_row;
    wb.resize_table(id, range).unwrap();
    wb.refresh_pivot(pivot).unwrap();
    assert_eq!(
        wb.find_pivot(pivot)
            .unwrap()
            .1
            .last_refresh
            .as_ref()
            .unwrap()
            .source_rows,
        0
    );
    assert_eq!(total(&wb, pivot), "", "empty sums preserve the existing pivot blank-result contract");
    assert_eq!(wb.sheet(1).unwrap().get_raw(2, 0), "", "obsolete output rows are cleared");
}

#[test]
fn refresh_rejects_recalculation_that_would_break_a_view_on_another_sheet() {
    let (mut wb, id, source, def) = fixture();
    let (pivot, _) = wb.create_pivot(source, def).unwrap();
    wb.set_cell_value_tracked(0, 0, 4, "=IF(Pivot!B4>35,SEQUENCE(3,1),0)");
    wb.set_table_view_spec(source.sheet_id, Some(TableViewSpec::new(id)))
        .unwrap();
    wb.set_cell_value_tracked(0, 2, 1, "40");
    let revision = wb.revision();
    assert!(
        wb.refresh_pivot(pivot).is_err(),
        "new result would spill alongside the filtered Table"
    );
    assert_eq!(wb.revision(), revision);
    assert_eq!(total(&wb, pivot), "30");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "0");
    assert!(wb.sheet(0).unwrap().build_saved_table_view(20).is_ok());
}
