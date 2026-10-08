use visigrid_engine::{
    cell::CellValue,
    cond_format::CondStyle,
    named_range::NamedRange,
    table::{TableId, TableRange, TableTotal},
    validation::{CellRange, ValidationRule, ValidationType},
    workbook::Workbook,
};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    wb.rename_sheet(0, "Sales Data");
    for (r, values) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (c, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, r, c, value);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    (wb, id)
}

#[test]
fn cells_footer_spans_and_names_follow_but_larger_ranges_keep_their_bounds() {
    let (mut wb, id) = book();
    let other = wb.add_sheet_named("Other").unwrap();
    for (r, source) in [
        "=('Sales Data'!$B$5+1)*2",
        "=SUM('Sales Data'!$A5:B$5)",
        "=SUM('Sales Data'!B2:B5)",
        "='Sales Data'!B6",
        "=SUM('Sales Data'!B:B)",
        "=B5",
        "=\"B5\"",
        "=SUM('Sales Data'!B6:B4)",
    ]
    .iter()
    .enumerate()
    {
        wb.set_cell_value_tracked(other, r, 3, source);
    }
    wb.set_cell_value_tracked(other, 4, 1, "42");
    wb.named_ranges_mut()
        .set(NamedRange::cell("FooterCell", 0, 4, 1))
        .unwrap();
    wb.named_ranges_mut()
        .set(NamedRange::range("FooterSpan", 0, 4, 0, 4, 1))
        .unwrap();
    wb.named_ranges_mut()
        .set(NamedRange::range("SpatialRange", 0, 1, 1, 4, 1))
        .unwrap();
    wb.set_cell_value_tracked(other, 8, 3, "=FooterCell+SUM(FooterSpan)");
    let before = wb.clone();
    let commit = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    for (r, source) in [
        "=('Sales Data'!$B$6+1)*2",
        "=SUM('Sales Data'!$A6:B$6)",
        "=SUM('Sales Data'!B2:B5)",
        "='Sales Data'!B6",
        "=SUM('Sales Data'!B:B)",
        "=B5",
        "=\"B5\"",
        "=SUM('Sales Data'!B6:B4)",
    ]
    .iter()
    .enumerate()
    {
        assert_eq!(wb.sheet(other).unwrap().get_raw(r, 3), *source);
    }
    for (r, value) in [
        (0, "142"),
        (1, "70"),
        (2, "70"),
        (3, "70"),
        (4, "140"),
        (5, "42"),
        (6, "B5"),
        (8, "140"),
    ] {
        assert_eq!(wb.sheet(other).unwrap().get_display(r, 3), value, "row {r}");
    }
    assert_eq!(
        wb.named_ranges()
            .get("FooterCell")
            .unwrap()
            .reference_string(),
        "B6"
    );
    assert_eq!(
        wb.named_ranges()
            .get("FooterSpan")
            .unwrap()
            .reference_string(),
        "A6:B6"
    );
    assert_eq!(
        wb.named_ranges()
            .get("SpatialRange")
            .unwrap()
            .reference_string(),
        "B2:B5"
    );
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(
        wb.sheet(other).unwrap().get_raw(0, 3),
        before.sheet(other).unwrap().get_raw(0, 3)
    );
    assert_eq!(
        wb.named_ranges()
            .get("FooterCell")
            .unwrap()
            .reference_string(),
        "B5"
    );
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_display(8, 3), "140");
}

#[test]
fn formulas_rules_dormant_totals_validation_and_frozen_provenance_move_together() {
    let (mut wb, id) = book();
    wb.set_table_total(
        id,
        0,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=$B$5+5".into()),
            label: None,
        },
    )
    .unwrap();
    let other = wb.add_sheet_named("Other").unwrap();
    for (r, values) in [["Qty", "Result"], ["1", ""], ["2", ""]].iter().enumerate() {
        for (c, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(other, r, c, value);
        }
    }
    let other_table = wb
        .create_table(
            wb.sheet(other).unwrap().id,
            TableRange {
                start_row: 0,
                end_row: 2,
                start_col: 0,
                end_col: 1,
            },
            "OtherData",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(other_table, 1, 1, "='Sales Data'!$B$5+[@Qty]", true)
        .unwrap();
    wb.set_table_totals_visible(other_table, true, Default::default())
        .unwrap();
    wb.set_table_total(
        other_table,
        1,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("='Sales Data'!$B$5".into()),
            label: None,
        },
    )
    .unwrap();
    wb.set_table_totals_visible(other_table, false, Default::default())
        .unwrap();
    let range = CellRange {
        start_row: 10,
        end_row: 15,
        start_col: 4,
        end_col: 4,
    };
    let sheet = wb.sheet_mut(other).unwrap();
    let cf = sheet.cond_formats.add(
        vec![range],
        "='Sales Data'!$B$5>50",
        CondStyle::Inline(Default::default()),
    );
    sheet
        .validations
        .set(range, ValidationRule::list_range("'Sales Data'!$A$5:$B$5"));
    sheet.validations.exclude(CellRange {
        start_row: 11,
        end_row: 11,
        ..range
    });
    sheet.freeze_cell(9, 4, CellValue::Number(60.0), "='Sales Data'!$B$5".into());
    let before = wb.clone();
    let commit = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    assert_eq!(wb.active_sheet().get_raw(5, 0), "=$B$6+5");
    assert_eq!(wb.active_sheet().get_display(5, 0), "75");
    assert_eq!(
        wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[0]
            .formula
            .as_deref(),
        Some("=$B$6+5")
    );
    assert_eq!(
        wb.table(other_table).unwrap().1.columns[1]
            .formula
            .as_deref(),
        Some("='Sales Data'!$B$6+[@[Qty]]")
    );
    assert_eq!(wb.sheet(other).unwrap().get_display(1, 1), "71");
    assert_eq!(
        wb.table(other_table)
            .unwrap()
            .1
            .totals
            .as_ref()
            .unwrap()
            .columns[1]
            .formula
            .as_deref(),
        Some("='Sales Data'!$B$6")
    );
    let sheet = wb.sheet(other).unwrap();
    assert_eq!(
        sheet.cond_formats.get(cf).unwrap().predicate,
        "='Sales Data'!$B$6>50"
    );
    assert_eq!(
        sheet.validations.get(10, 4).unwrap().rule_type,
        ValidationType::List(visigrid_engine::validation::ListSource::Range(
            "'Sales Data'!$A$6:$B$6".into()
        ))
    );
    assert!(sheet.validations.get(11, 4).is_none());
    assert_eq!(
        sheet.get_cell(9, 4).frozen_formula(),
        Some("='Sales Data'!$B$6")
    );
    assert_eq!(sheet.get_display(9, 4), "60");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(
        serde_json::to_value(wb.saved_tables()).unwrap(),
        serde_json::to_value(before.saved_tables()).unwrap()
    );
    assert_eq!(
        wb.sheet(other).unwrap().get_cell(9, 4).frozen_formula(),
        Some("='Sales Data'!$B$5")
    );
    wb.apply_table_commit(&commit, false).unwrap();
    wb.set_table_totals_visible(other_table, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_display(3, 1), "70");
}

#[test]
fn shrink_and_compound_resize_relocate_only_the_surviving_footer_cells() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 10, 4, "=B5");
    let range = wb.table(id).unwrap().1.range;
    let grow = wb
        .resize_table(
            id,
            TableRange {
                end_row: 7,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B9");
    let shrink = wb
        .resize_table(
            id,
            TableRange {
                end_row: 5,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B7");
    wb.apply_table_commit(&shrink, true).unwrap();
    wb.apply_table_commit(&grow, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B5");
    wb.set_cell_value_tracked(0, 11, 4, "=A5");
    let narrowed = wb
        .resize_table(
            id,
            TableRange {
                end_col: 0,
                end_row: 6,
                ..range
            },
        )
        .unwrap();
    assert_eq!(
        wb.active_sheet().get_raw(10, 4),
        "=B5",
        "released footer stays put"
    );
    assert_eq!(wb.active_sheet().get_raw(11, 4), "=A8");
    wb.apply_table_commit(&narrowed, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(11, 4), "=A5");
}

#[test]
fn failures_never_publish_rewritten_references_and_new_input_keeps_its_coordinates() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 10, 4, "=B5");
    wb.set_cell_value_tracked(0, 5, 1, "Occupied");
    let revision = wb.revision();
    assert!(wb.append_table_rows(id, 1, &[]).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B5");
    wb.clear_cell_tracked(0, 5, 1);
    let commit = wb
        .append_table_rows(id, 1, &[(4, 0, "=B5".into()), (4, 1, "10".into())])
        .unwrap();
    assert_eq!(wb.active_sheet().get_raw(4, 0), "=B5");
    assert_eq!(wb.active_sheet().get_display(4, 0), "10");
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B6");
    let good = wb.clone();
    wb.set_cell_value_tracked(0, 20, 4, "=B6");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    wb.restore_snapshot_monotonic(&good);
    wb.apply_table_commit(&commit, true).unwrap();
    wb.set_cell_value_tracked(0, 20, 4, "=B5");
    assert!(wb.apply_table_commit(&commit, false).is_err());
}

#[test]
fn new_footer_references_also_block_replay_when_the_original_append_had_no_links() {
    let (mut wb, id) = book();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    let after = wb.clone();
    wb.set_cell_value_tracked(0, 10, 4, "=B5");
    assert!(wb.apply_table_commit(&commit, true).is_err());
    wb.restore_snapshot_monotonic(&after);
    wb.apply_table_commit(&commit, true).unwrap();
    wb.set_cell_value_tracked(0, 10, 4, "=B6");
    assert!(wb.apply_table_commit(&commit, false).is_err());

    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 10, 4, "=B6"); // Existing destination reference stays spatial.
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.active_sheet().get_display(10, 4), "60");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B6");
    assert_eq!(wb.active_sheet().get_display(10, 4), "");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(10, 4), "=B6");
}

#[test]
fn relocation_that_introduces_an_unsafe_spill_refuses_without_publishing() {
    let (mut wb, id) = book();
    let mut view = visigrid_engine::table_view::TableViewSpec::new(id);
    view.sort = Some(visigrid_engine::table_view::TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: visigrid_engine::filter::SortDirection::Ascending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(view))
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 4, "=IF(ROW(B5)=5,0,SEQUENCE(4))");
    assert_eq!(wb.active_sheet().get_display(0, 4), "0");
    let revision = wb.revision();
    assert!(wb.append_table_rows(id, 1, &[]).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(
        wb.active_sheet().get_raw(0, 4),
        "=IF(ROW(B5)=5,0,SEQUENCE(4))"
    );
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
}

#[test]
fn fixed_pivot_source_confined_to_footer_moves_and_replays_with_it() {
    use visigrid_engine::pivot::{
        Aggregation, PivotDefinition, PivotField, PivotSource, PivotValueField,
    };
    let (mut wb, id) = book();
    let source = PivotSource {
        table_id: None,
        sheet_id: wb.active_sheet_id(),
        start_row: 4,
        end_row: 4,
        start_col: 0,
        end_col: 1,
    };
    let definition = PivotDefinition {
        rows: Vec::new(),
        column: None,
        values: vec![PivotValueField {
            field: PivotField {
                offset: 1,
                header: "60".into(),
                column_id: None,
            },
            aggregation: Aggregation::Sum,
            number_format: None,
        }],
    };
    let (pivot, _) = wb.create_pivot(source, definition).unwrap();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.find_pivot(pivot).unwrap().1.source.start_row, 5);
    assert_eq!(wb.find_pivot(pivot).unwrap().1.source.end_row, 5);
    assert!(wb.is_pivot_stale(wb.find_pivot(pivot).unwrap().1));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.find_pivot(pivot).unwrap().1.source.start_row, 4);
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.find_pivot(pivot).unwrap().1.source.start_row, 5);
    wb.refresh_pivot(pivot).unwrap();
    assert!(!wb.is_pivot_stale(wb.find_pivot(pivot).unwrap().1));
}

#[test]
fn unsupported_formulas_and_predicates_do_not_block_unrelated_footer_movement() {
    let (mut wb, id) = book();
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=UNSUPPORTED(@)");
    wb.set_cell_value_tracked(other, 1, 0, "='Sales Data'!$B$5#+@A1");
    wb.active_sheet_mut().cond_formats.add(vec![CellRange::single(12, 8)], "=UNSUPPORTED(@)", CondStyle::Named(visigrid_engine::cell::CellStyle::Warning));
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "=UNSUPPORTED(@)");
    assert_eq!(wb.sheet(other).unwrap().get_raw(1, 0), "='Sales Data'!$B$6#+@A1");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_raw(1, 0), "='Sales Data'!$B$5#+@A1");
}

#[test]
fn shrinking_width_keeps_a_released_footer_cycle_at_its_stored_position() {
    let (mut wb, id) = book();
    wb.set_table_total(id, 1, TableTotal {
        function: Some("custom".into()), formula: Some("=B5".into()), label: None,
    }).unwrap();
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "='Sales Data'!A5");
    let commit = wb.resize_table(id, TableRange {
        start_row: 0, start_col: 0, end_row: 4, end_col: 0,
    }).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "#CYCLE!");
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "='Sales Data'!A6");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "#CYCLE!");
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "='Sales Data'!A5");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "#CYCLE!");
}
