use visigrid_engine::{
    pivot::{Aggregation, PivotDefinition, PivotField, PivotValueField},
    table::TableRange,
    workbook::Workbook,
};
use visigrid_io::{json, native};

fn fixture() -> (Workbook, u64) {
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
    let field = |i: usize| PivotField {
        column_id: Some(columns[i].id),
        offset: i as u32,
        header: columns[i].name.clone(),
    };
    let definition = PivotDefinition {
        rows: vec![field(0)],
        column: None,
        values: vec![PivotValueField {
            field: field(1),
            aggregation: Aggregation::Sum,
            number_format: None,
        }],
    };
    let (pivot, _) = wb.create_pivot(source, definition).unwrap();
    // Persist criteria too: neither source binding nor filter may be lost.
    let mut spec = visigrid_engine::table_view::TableViewSpec::new(id);
    spec.sort = Some(visigrid_engine::table_view::TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: visigrid_engine::filter::SortDirection::Descending,
    });
    wb.set_table_view_spec(source.sheet_id, Some(spec)).unwrap();
    (wb, pivot)
}

fn verify(original: &Workbook, mut loaded: Workbook, pivot: u64) {
    let p = loaded.find_pivot(pivot).unwrap().1;
    assert_eq!(
        p.definition,
        original.find_pivot(pivot).unwrap().1.definition
    );
    assert_eq!(p.source.sheet_id, loaded.sheet(0).unwrap().id);
    assert!(loaded.sheet(0).unwrap().table_view_spec().is_some());
    assert!(!loaded.is_pivot_stale(p));
    let id = p.source.table_id.unwrap();
    let source_sheet = p.source.sheet_id;
    loaded.set_cell_value_tracked(0, 1, 1, "11");
    assert!(
        loaded.is_pivot_stale(loaded.find_pivot(pivot).unwrap().1),
        "first edit after reopen is detected"
    );
    loaded.set_cell_value_tracked(0, 1, 1, "10");
    // Clear the view only to allow append, then test that the loaded binding grows.
    loaded.set_table_view_spec(source_sheet, None).unwrap();
    loaded
        .append_table_rows(id, 1, &[(3, 0, "West".into()), (3, 1, "7".into())])
        .unwrap();
    loaded
        .rename_table_columns(id, &["Area".into(), "Revenue".into()])
        .unwrap();
    assert!(loaded.is_pivot_stale(loaded.find_pivot(pivot).unwrap().1));
    loaded.refresh_pivot(pivot).unwrap();
    assert_eq!(loaded.sheet(1).unwrap().get_raw(3, 1), "37");
    assert_eq!(loaded.sheet(1).unwrap().get_display(0, 1), "Sum of Revenue");
}

#[test]
fn native_table_pivot_roundtrip() {
    let (wb, pivot) = fixture();
    let file = tempfile::NamedTempFile::with_suffix(".sheet").unwrap();
    native::save_workbook(&wb, file.path()).unwrap();
    verify(&wb, native::load_workbook(file.path()).unwrap(), pivot);
}

#[test]
fn desktop_full_save_preserves_table_and_range_pivots() {
    let (mut wb, pivot) = fixture();
    let mut source = wb.find_pivot(pivot).unwrap().1.source;
    source.table_id = None;
    let mut definition = wb.find_pivot(pivot).unwrap().1.definition.clone();
    for field in definition
        .rows
        .iter_mut()
        .chain(definition.values.iter_mut().map(|v| &mut v.field))
    {
        field.column_id = None;
    }
    let (range_pivot, _) = wb.create_pivot(source, definition).unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".sheet").unwrap();
    native::save_workbook_full(&wb, &native::CellMetadata::default(), &[], &[], file.path())
        .unwrap();
    let loaded = native::load_workbook(file.path()).unwrap();
    assert_eq!(loaded.pivots().len(), 2);
    assert!(loaded
        .find_pivot(range_pivot)
        .unwrap()
        .1
        .source
        .table_id
        .is_none());
    verify(&wb, loaded, pivot);
}

#[test]
fn full_json_table_pivot_roundtrip() {
    let (wb, pivot) = fixture();
    let json = json::export_workbook(&wb, &[], 0).unwrap();
    let (loaded, _, _) = json::import_any(&json).unwrap();
    verify(&wb, loaded, pivot);
}

#[test]
fn missing_field_stays_bound_after_reopen() {
    let (mut wb, pivot) = fixture();
    let p = wb.find_pivot(pivot).unwrap().1;
    let source = p.source;
    wb.set_table_view_spec(source.sheet_id, None).unwrap();
    wb.structural_edit(0, visigrid_engine::structural::Axis::Col, 1, 1, true)
        .unwrap();
    let json = json::export_workbook(&wb, &[], 0).unwrap();
    let (mut loaded, _, _) = json::import_any(&json).unwrap();
    assert!(loaded.is_pivot_stale(loaded.find_pivot(pivot).unwrap().1));
    assert!(loaded.refresh_pivot(pivot).unwrap_err().contains("missing"));
    assert_eq!(loaded.sheet(1).unwrap().get_raw(3, 1), "30");
}
