use visigrid_engine::{
    cell::{CellComment, CellFormat},
    cond_format::CondStyle,
    filter::{ColumnFilter, FilterKey, SortDirection},
    formula::eval::Value,
    sheet::{MergedRegion, Sheet, SheetId},
    structural::Axis,
    table::{TableColumnId, TableId, TableRange},
    table_view::{
        validate_table_view_layout, TableFilter, TableSort, TableView, TableViewSpec, ViewOwner,
    },
    validation::{CellRange, ValidationRule},
    workbook::Workbook,
};

fn fixture() -> (Workbook, TableViewSpec) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(1), 30, 12)], 0);
    for (offset, row) in [
        ["Key", "Amount", "Label"],
        ["b", "20", "B"],
        ["A", "10", "A"],
        ["a", "10", "A2"],
        ["c", "", "blank"],
        ["", "5", "empty key"],
    ]
    .into_iter()
    .enumerate()
    {
        for (col, value) in row.into_iter().enumerate() {
            wb.set_cell_value_tracked(0, offset + 2, col + 1, value);
        }
    }
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 7,
                end_col: 3,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_cell_value_tracked(0, 0, 0, "=SUM(Sales[Amount])");
    wb.set_cell_value_tracked(0, 12, 8, "Unrelated note below the Table");
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: column(&wb, id, 1),
        direction: SortDirection::Ascending,
    });
    (wb, spec)
}

fn column(wb: &Workbook, id: TableId, offset: usize) -> TableColumnId {
    wb.table(id).unwrap().1.columns[offset].id
}

fn selected(values: &[Value]) -> ColumnFilter {
    ColumnFilter {
        selected: Some(
            values
                .iter()
                .map(|v| FilterKey::from_value(v).normalized())
                .collect(),
        ),
        text_filter: None,
    }
}

fn build(wb: &Workbook, spec: TableViewSpec) -> TableView {
    TableView::build(wb.active_sheet(), spec, wb.active_sheet().rows, None).unwrap()
}

#[test]
fn sort_is_bounded_stable_in_both_directions_and_keeps_blanks_last() {
    let (wb, mut spec) = fixture();
    let before: Vec<_> = wb
        .active_sheet()
        .cells_iter()
        .map(|(pos, c)| (pos, c.raw_display()))
        .collect();
    let revision = wb.revision();
    for (direction, expected) in [
        (SortDirection::Ascending, [7, 4, 5, 3, 6]),
        (SortDirection::Descending, [3, 4, 5, 7, 6]),
    ] {
        spec.sort.as_mut().unwrap().direction = direction;
        let view = build(&wb, spec.clone());
        assert_eq!(&view.rows().row_order()[3..=7], &expected);
        for row in (0..3).chain(8..30) {
            assert_eq!(view.rows().view_to_data(row), row);
            assert!(view.rows().is_data_row_visible(row));
        }
        assert_eq!(view.rows().view_to_data(2), 2);
        assert!(view.rows().is_data_row_visible(2));
    }
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_display(0, 0), "45");
    assert_eq!(
        before,
        wb.active_sheet()
            .cells_iter()
            .map(|(pos, c)| (pos, c.raw_display()))
            .collect::<Vec<_>>()
    );
}

#[test]
fn multiple_filters_use_typed_keys_and_do_not_change_aggregates() {
    let (wb, mut spec) = fixture();
    spec.filters = vec![
        TableFilter {
            column: column(&wb, spec.table, 0),
            criteria: selected(&[Value::Text(" a ".into())]),
        },
        TableFilter {
            column: column(&wb, spec.table, 1),
            criteria: selected(&[Value::Number(10.0)]),
        },
    ];
    let view = build(&wb, spec);
    assert_eq!(view.visible_body_rows(4, 2).unwrap(), [4, 5]);
    assert!(!view.rows().is_data_row_visible(3));
    assert!(!view.rows().is_data_row_visible(7));
    assert_eq!(wb.active_sheet().get_display(0, 0), "45");
}

#[test]
fn numeric_text_booleans_errors_and_blanks_remain_distinct() {
    let (mut wb, mut spec) = fixture();
    wb.active_sheet_mut().set_text(3, 2, "10");
    wb.set_cell_value_tracked(0, 5, 2, "=1=1");
    assert_eq!(
        wb.active_sheet().get_computed_value(5, 2),
        Value::Boolean(true)
    );
    wb.set_cell_value_tracked(0, 7, 2, "=1/0");
    let id = column(&wb, spec.table, 1);
    for (value, expected) in [
        (Value::Text("10".into()), 3),
        (Value::Number(10.0), 4),
        (Value::Boolean(true), 5),
        (Value::Empty, 6),
        (Value::Error("#DIV/0!".into()), 7),
    ] {
        spec.filters = vec![TableFilter {
            column: id,
            criteria: selected(&[value]),
        }];
        let view = build(&wb, spec.clone());
        let visible: Vec<_> = (3..=7)
            .filter(|row| view.rows().is_data_row_visible(*row))
            .collect();
        assert_eq!(visible, [expected]);
    }
}

#[test]
fn hiding_buttons_and_clearing_sort_or_filters_are_independent() {
    let (wb, mut spec) = fixture();
    spec.filters.push(TableFilter {
        column: column(&wb, spec.table, 0),
        criteria: selected(&[Value::Text("b".into())]),
    });
    spec.show_filter_buttons = false;
    let view = build(&wb, spec.clone());
    assert!(view.rows().is_sorted());
    assert!(view.rows().is_filtered());
    let mut without_sort = spec.clone();
    without_sort.clear_sort();
    let unsorted = build(&wb, without_sort);
    assert!(!unsorted.rows().is_sorted());
    assert!(unsorted.rows().is_filtered());
    spec.clear_filters();
    let unfiltered = build(&wb, spec);
    assert!(unfiltered.rows().is_sorted());
    assert!(!unfiltered.rows().is_filtered());
}

#[test]
fn ids_survive_renames_column_insertion_and_spec_serialization() {
    let (mut wb, mut spec) = fixture();
    let id = spec.table;
    spec.filters.push(TableFilter {
        column: column(&wb, id, 1),
        criteria: selected(&[Value::Number(10.0)]),
    });
    let before = build(&wb, spec.clone());
    wb.rename_table(id, "Orders").unwrap();
    wb.rename_table_columns(id, &["Region".into(), "Revenue".into(), "Note".into()])
        .unwrap();
    wb.structural_edit(0, Axis::Col, 2, 1, false).unwrap();
    let rebound = before.rebuild(wb.active_sheet()).unwrap();
    assert_eq!(rebound.filters().sort.as_ref().unwrap().column, 3);
    assert!(rebound.filters().column_filters.contains_key(&3));
    assert_eq!(before.rows().visible_rows(), rebound.rows().visible_rows());
    assert_eq!(before.rows().row_order(), rebound.rows().row_order());
    let encoded = serde_json::to_string(&spec).unwrap();
    let restored: TableViewSpec = serde_json::from_str(&encoded).unwrap();
    assert_eq!(restored, spec);
    assert_eq!(
        build(&wb, restored).rows().row_order(),
        rebound.rows().row_order()
    );
}

#[test]
fn removed_fields_and_recreated_names_do_not_rebind() {
    let (mut wb, spec) = fixture();
    let before = build(&wb, spec.clone());
    wb.structural_edit(0, Axis::Col, 2, 1, true).unwrap();
    assert!(before
        .rebuild(wb.active_sheet())
        .unwrap_err()
        .contains("no longer exists"));
    wb.structural_edit(0, Axis::Col, 2, 1, false).unwrap();
    wb.rename_table_columns(spec.table, &["Key".into(), "Amount".into(), "Label".into()])
        .unwrap();
    assert!(before.rebuild(wb.active_sheet()).is_err());
    wb.remove_table(spec.table).unwrap();
    wb.create_table(SheetId(1), before.range(), "Sales")
        .unwrap();
    assert!(before
        .rebuild(wb.active_sheet())
        .unwrap_err()
        .contains("Table"));
}

#[test]
fn only_one_owner_can_be_active_and_entire_table_must_fit() {
    let (wb, spec) = fixture();
    for owner in [ViewOwner::Range, ViewOwner::Table(TableId(999))] {
        assert!(
            TableView::build(wb.active_sheet(), spec.clone(), 30, Some(owner))
                .unwrap_err()
                .contains("Clear the current")
        );
    }
    assert!(TableView::build(
        wb.active_sheet(),
        spec.clone(),
        30,
        Some(ViewOwner::Table(spec.table))
    )
    .is_ok());
    for extent in [0, 7, 31, usize::MAX] {
        assert!(TableView::build(wb.active_sheet(), spec.clone(), extent, None).is_err());
    }
    assert!(TableView::build(wb.active_sheet(), spec, 8, None).is_ok());
}

#[test]
fn header_only_missing_table_invalid_columns_and_duplicate_filters_are_rejected() {
    let (mut wb, mut spec) = fixture();
    spec.sort.as_mut().unwrap().column = TableColumnId(999);
    assert!(TableView::build(wb.active_sheet(), spec.clone(), 30, None).is_err());
    spec.clear_sort();
    let filter = TableFilter {
        column: column(&wb, spec.table, 0),
        criteria: ColumnFilter::default(),
    };
    spec.filters = vec![filter.clone(), filter];
    assert!(TableView::build(wb.active_sheet(), spec.clone(), 30, None)
        .unwrap_err()
        .contains("only one"));
    spec.clear_filters();
    let mut header = wb.table(spec.table).unwrap().1.range;
    header.end_row = header.start_row;
    wb.resize_table(spec.table, header).unwrap();
    assert!(TableView::build(wb.active_sheet(), spec.clone(), 30, None)
        .unwrap_err()
        .contains("no body rows"));
    spec.table = TableId(999);
    assert!(TableView::build(wb.active_sheet(), spec, 30, None).is_err());
}

#[test]
fn visible_paste_maps_once_and_rejects_overflow_before_producing_targets() {
    let (wb, mut spec) = fixture();
    spec.filters.push(TableFilter {
        column: column(&wb, spec.table, 1),
        criteria: selected(&[Value::Number(20.0), Value::Number(10.0)]),
    });
    let view = build(&wb, spec);
    assert_eq!(view.visible_body_rows(4, 3).unwrap(), [4, 5, 3]);
    assert!(view.visible_body_rows(4, 4).is_err());
    assert!(view.visible_body_rows(3, 1).is_err()); // hidden record
    assert!(view.visible_body_rows(2, 1).is_err()); // header
    assert!(view.visible_body_rows(8, 1).is_err()); // outside Table
    assert!(view.visible_body_rows(usize::MAX, 1).is_err());
    assert!(view.visible_body_rows(4, 0).unwrap().is_empty());
}

#[test]
fn edit_rebuild_retains_record_focus_or_reports_nearest_visible_row() {
    let (mut wb, mut spec) = fixture();
    let view = build(&wb, spec.clone());
    wb.set_cell_value_tracked(0, 3, 2, "1");
    let rebuilt = view.rebuild(wb.active_sheet()).unwrap();
    let focus = rebuilt.focus_record(3).unwrap();
    assert_eq!(focus.data_row, 3);
    assert_eq!(focus.view_row, 3);
    assert!(!focus.record_hidden);
    spec.filters.push(TableFilter {
        column: column(&wb, spec.table, 1),
        criteria: selected(&[Value::Number(10.0)]),
    });
    let filtered = build(&wb, spec);
    let focus = filtered.focus_record(3).unwrap();
    assert!(focus.record_hidden);
    assert!(filtered.rows().is_data_row_visible(focus.data_row));
    assert_eq!(filtered.rows().view_to_data(focus.view_row), focus.data_row);
    assert_eq!(view.rows().view_to_data(3), 7); // old snapshot not mutated
    assert!(filtered.focus_record(30).is_err());
}

#[test]
fn recalculated_cross_sheet_keys_change_order_and_visibility_on_rebuild() {
    let (mut wb, mut spec) = fixture();
    let inputs = wb.add_sheet_named("Inputs").unwrap();
    wb.set_cell_value_tracked(inputs, 0, 0, "15");
    wb.set_cell_value_tracked(0, 3, 2, "=Inputs!A1");
    spec.filters.push(TableFilter {
        column: column(&wb, spec.table, 1),
        criteria: selected(&[Value::Number(15.0), Value::Number(10.0)]),
    });
    let view = build(&wb, spec);
    assert_eq!(view.visible_body_rows(4, 3).unwrap(), [4, 5, 3]);
    wb.set_cell_value_tracked(inputs, 0, 0, "1");
    assert_eq!(wb.active_sheet().get_display(3, 2), "1");
    let rebuilt = view.rebuild(wb.active_sheet()).unwrap();
    assert_eq!(rebuilt.rows().view_to_data(3), 3);
    assert!(!rebuilt.rows().is_data_row_visible(3));
    assert_eq!(rebuilt.visible_body_rows(5, 2).unwrap(), [4, 5]);
    assert_eq!(wb.active_sheet().get_display(0, 0), "26");
}

#[test]
fn filtering_every_record_out_keeps_header_and_outside_rows_available() {
    let (wb, mut spec) = fixture();
    spec.filters.push(TableFilter {
        column: column(&wb, spec.table, 1),
        criteria: selected(&[]),
    });
    let view = build(&wb, spec);
    for row in 0..30 {
        assert_eq!(
            view.rows().is_data_row_visible(row),
            !(3..=7).contains(&row)
        );
    }
    assert!(view.visible_body_rows(3, 1).is_err());
    let focus = view.focus_record(4).unwrap();
    assert!(focus.record_hidden);
    assert_eq!(focus.data_row, 2); // nearest visible row is the Table header
}

#[test]
fn sparse_layout_guard_covers_adjacent_values_and_metadata() {
    type Mutate = fn(&mut Sheet);
    let cases: &[(&str, Mutate)] = &[
        ("value", |s| s.set_value(4, 5, "outside")),
        ("formula displaying blank", |s| {
            s.set_value(4, 5, "=IF(1=1,\"\",1)")
        }),
        ("cell format", |s| s.set_bold(4, 5, true)),
        ("style metadata", |s| s.set_style_id(4, 5, 9)),
        ("comment", |s| {
            s.set_comment(
                4,
                5,
                Some(CellComment {
                    text: "note".into(),
                    author: "test".into(),
                }),
            )
        }),
        ("validation", |s| {
            s.validations.set(
                CellRange::single(4, 5),
                ValidationRule::list_inline(vec!["x".into()]),
            )
        }),
        ("validation exclusion", |s| {
            s.validations.exclude(CellRange::single(4, 5))
        }),
        ("conditional format", |s| {
            s.cond_formats.add(
                vec![CellRange::single(4, 5)],
                "=1=1",
                CondStyle::Inline(Default::default()),
            );
        }),
        ("merge", |s| {
            s.set_merges(vec![MergedRegion::new(3, 5, 4, 5)]);
        }),
        ("row format", |s| {
            s.row_formats.insert(
                4,
                CellFormat {
                    bold: true,
                    ..Default::default()
                },
            );
        }),
        ("column format", |s| {
            s.col_formats.insert(
                5,
                CellFormat {
                    bold: true,
                    ..Default::default()
                },
            );
        }),
        ("spill into body", |s| s.set_value(1, 5, "=SEQUENCE(5)")),
    ];
    for (name, mutate) in cases {
        let (mut wb, spec) = fixture();
        let view = build(&wb, spec.clone());
        mutate(wb.active_sheet_mut());
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        if *name == "formula displaying blank" {
            assert_eq!(
                wb.active_sheet().get_computed_value(4, 5),
                Value::Text(String::new())
            );
        }
        assert!(
            validate_table_view_layout(wb.active_sheet(), spec.table).is_err(),
            "{name}"
        );
        assert!(view.rebuild(wb.active_sheet()).is_err(), "{name}");
        assert!(
            view.validate_mutation_ranges(wb.active_sheet(), &[CellRange::single(4, 2)])
                .is_err(),
            "{name}"
        );
    }
}

#[test]
fn other_tables_and_pivot_output_in_body_rows_are_refused() {
    let (mut wb, spec) = fixture();
    wb.create_table(
        SheetId(1),
        TableRange {
            start_row: 4,
            start_col: 5,
            end_row: 9,
            end_col: 6,
        },
        "Beside",
    )
    .unwrap();
    assert!(validate_table_view_layout(wb.active_sheet(), spec.table).is_err());
    let (mut wb, spec) = fixture();
    wb.active_sheet_mut()
        .pivots
        .push(visigrid_engine::pivot::PivotTable {
            id: 1,
            name: "Pivot1".into(),
            source: visigrid_engine::pivot::PivotSource {
                sheet_id: SheetId(1),
                start_row: 0,
                start_col: 0,
                end_row: 1,
                end_col: 1,
            },
            definition: Default::default(),
            anchor_row: 3,
            anchor_col: 7,
            extent: Some((2, 2)),
            last_refresh: None,
            stale: false,
            source_generation: None,
        });
    assert!(validate_table_view_layout(wb.active_sheet(), spec.table)
        .unwrap_err()
        .contains("pivot output"));
}

#[test]
fn titles_and_metadata_above_below_or_inside_the_table_are_safe() {
    let (mut wb, spec) = fixture();
    let sheet = wb.active_sheet_mut();
    sheet.set_merges(vec![
        MergedRegion::new(0, 5, 1, 6),
        MergedRegion::new(10, 5, 12, 6),
    ]);
    sheet.set_bold(4, 2, true);
    sheet.set_comment(
        1,
        5,
        Some(CellComment {
            text: "title".into(),
            author: String::new(),
        }),
    );
    sheet.validations.set(
        CellRange::new(3, 1, 7, 3),
        ValidationRule::list_inline(vec!["x".into()]),
    );
    sheet.cond_formats.add(
        vec![CellRange::single(4, 2)],
        "=1=1",
        CondStyle::Inline(Default::default()),
    );
    assert!(validate_table_view_layout(sheet, spec.table).is_ok());
    assert!(TableView::build(sheet, spec, 30, None).is_ok());
}

#[test]
fn batch_footprints_reject_adjacent_writes_and_stale_shapes_without_mutating() {
    let (mut wb, spec) = fixture();
    let view = build(&wb, spec.clone());
    let revision = wb.revision();
    let batch = [
        CellRange::single(4, 2),
        CellRange::single(5, 7),
        CellRange::single(6, 2),
    ];
    assert!(view
        .validate_mutation_ranges(wb.active_sheet(), &batch)
        .unwrap_err()
        .contains("Clear the Table"));
    assert_eq!(wb.revision(), revision);
    assert!(view
        .validate_mutation_ranges(
            wb.active_sheet(),
            &[
                CellRange::single(4, 2),
                CellRange::single(1, 7),
                CellRange::single(9, 7)
            ]
        )
        .is_ok());
    assert!(view
        .validate_mutation_ranges(wb.active_sheet(), &[CellRange::new(4, 0, 4, 2)])
        .is_err());
    assert!(view
        .validate_mutation_ranges(wb.active_sheet(), &[CellRange::single(30, 2)])
        .is_err());
    assert!(view.rebuild(&Sheet::new(SheetId(2), 30, 12)).is_err());
    let mut grown = view.range();
    grown.end_row = 8;
    wb.resize_table(spec.table, grown).unwrap();
    assert!(view
        .validate_mutation_ranges(wb.active_sheet(), &[CellRange::single(4, 2)])
        .unwrap_err()
        .contains("changed shape"));
    assert_eq!(view.rebuild(wb.active_sheet()).unwrap().range(), grown);
}

#[test]
fn records_beyond_ten_thousand_participate_in_sort_filter_and_unique_values() {
    let last = 20_002;
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(1), last + 3, 8)], 0);
    wb.active_sheet_mut().set_value(2, 1, "Amount");
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: last,
                end_col: 1,
            },
            "LargeRecords",
        )
        .unwrap()
        .table_id();
    wb.active_sheet_mut().set_value(3, 1, "10");
    wb.active_sheet_mut().set_value(last, 1, "1");
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: column(&wb, id, 0),
        direction: SortDirection::Ascending,
    });
    spec.filters.push(TableFilter {
        column: column(&wb, id, 0),
        criteria: selected(&[Value::Number(1.0)]),
    });
    let view = build(&wb, spec);
    assert_eq!(view.visible_body_rows(3, 1).unwrap(), [last]);
    assert!(!view.rows().is_data_row_visible(3));
    assert!(view.rows().is_data_row_visible(last + 1));
    let mut filters = view.filters().clone();
    let values = filters.build_unique_values(
        1,
        |r, c| wb.active_sheet().get_computed_value(r, c),
        usize::MAX,
    );
    assert_eq!(values.iter().map(|v| v.count).sum::<usize>(), 20_000);
    assert_eq!(values.len(), 3); // 1, 10 and blank; full body, including last record
}
