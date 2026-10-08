use visigrid_engine::{
    cond_format::CondStyle,
    filter::SortDirection,
    table::{TableId, TableRange, TableTotal},
    table_view::{TableSort, TableViewSpec},
    validation::{CellRange, ListSource, ValidationRule, ValidationType},
    workbook::Workbook,
};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (row, values) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    let table = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(table, true, Default::default())
        .unwrap();
    (wb, table)
}

#[test]
fn dynamic_footer_bases_follow_but_constructed_coordinates_stay_literal() {
    let (mut wb, id) = book();
    wb.define_name_for_cell("Footer", 0, 4, 1).unwrap();
    let formulas = [
        "=OFFSET($B$5,0,0)",
        "=INDIRECT(\"B5\")",
        "=OFFSET($B$1,4,0)",
        "=SUM(OFFSET(INDIRECT(\"Footer\"),0,0))",
        "=INDIRECT(\"Sales[[#Totals],[Amount]]\")",
        "=OFFSET(OFFSET($B$5,0,0),0,0)",
        "=INDIRECT(\"B6\")",
        "=LET(base,OFFSET($B$5,0,0),SUM(base))",
    ];
    for (row, formula) in formulas.iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 4, formula);
    }
    let commit = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    for row in [0, 3, 4, 5, 6, 7] {
        assert_eq!(wb.active_sheet().get_display(row, 4), "70", "row {row}");
    }
    for row in [1, 2] {
        assert_eq!(wb.active_sheet().get_display(row, 4), "10");
    }
    assert_eq!(wb.active_sheet().get_raw(0, 4), "=OFFSET($B$6, 0, 0)");
    assert_eq!(wb.active_sheet().get_raw(1, 4), formulas[1]);
    assert_eq!(wb.active_sheet().get_raw(2, 4), formulas[2]);
    wb.apply_table_commit(&commit, true).unwrap();
    for (row, formula) in formulas.iter().enumerate() {
        assert_eq!(wb.active_sheet().get_raw(row, 4), *formula);
        assert_eq!(
            wb.active_sheet().get_display(row, 4),
            if row == 6 { "" } else { "60" }
        );
    }
    wb.apply_table_commit(&commit, false).unwrap();
    wb.set_cell_value_tracked(0, 1, 1, "15");
    for row in [0, 3, 4, 5, 6, 7] {
        assert_eq!(wb.active_sheet().get_display(row, 4), "75");
    }
    assert!(wb.apply_table_commit(&commit, true).is_err());
}

#[test]
fn dynamic_footer_rules_validation_and_frozen_text_follow_across_sheets() {
    let (mut wb, id) = book();
    let other = wb.add_sheet_named("Other").unwrap();
    let source = "OFFSET(Sheet1!$B$5,0,0)";
    wb.set_cell_value_tracked(other, 0, 0, &format!("={source}"));
    wb.set_cell_text_tracked(other, 1, 0, "cached");
    let mut frozen = wb.sheet(other).unwrap().get_cell(1, 0);
    frozen.set_frozen_formula(Some(format!("={source}")));
    wb.restore_cell_tracked(other, 1, 0, Some(frozen)).unwrap();
    let range = CellRange {
        start_row: 3,
        start_col: 0,
        end_row: 3,
        end_col: 0,
    };
    let mut rule = ValidationRule::list_range(source);
    rule.rule_type = ValidationType::List(ListSource::NamedRange(source.into()));
    wb.sheet_mut(other).unwrap().validations.set(range, rule);
    let cf = wb.sheet_mut(other).unwrap().cond_formats.add(
        vec![range],
        &format!("={source}>0"),
        CondStyle::Inline(Default::default()),
    );
    let other_id = wb
        .create_table(
            wb.sheets()[other].id,
            TableRange {
                start_row: 5,
                start_col: 0,
                end_row: 7,
                end_col: 1,
            },
            "OtherTable",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(other_id, 1, 6, &format!("={source}"), true)
        .unwrap();
    wb.set_table_totals_visible(other_id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        other_id,
        1,
        TableTotal {
            function: Some("custom".into()),
            formula: Some(format!("={source}")),
            label: None,
        },
    )
    .unwrap();
    wb.set_table_totals_visible(other_id, false, Default::default())
        .unwrap();
    let commit = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "70");
    assert_eq!(wb.sheet(other).unwrap().get_display(6, 1), "70");
    assert_eq!(
        wb.sheet(other).unwrap().get_cell(1, 0).frozen_formula(),
        Some("=OFFSET(Sheet1!$B$6, 0, 0)")
    );
    assert_eq!(wb.sheet(other).unwrap().get_display(1, 0), "cached");
    assert_eq!(wb.get_list_items(other, 3, 0).unwrap().items, ["70"]);
    assert_eq!(
        wb.sheet(other)
            .unwrap()
            .cond_formats
            .get(cf)
            .unwrap()
            .predicate,
        "=OFFSET(Sheet1!$B$6, 0, 0)>0"
    );
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.get_list_items(other, 3, 0).unwrap().items, ["60"]);
    wb.apply_table_commit(&commit, false).unwrap();
    wb.set_table_totals_visible(other_id, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_display(8, 1), "70");
}

#[test]
fn dynamic_footer_cycles_from_new_input_or_custom_totals_refuse_atomically() {
    let (mut wb, id) = book();
    let revision = wb.revision();
    assert!(wb
        .append_table_rows(id, 1, &[(4, 1, "=INDIRECT(\"B6\")".into())])
        .is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_display(4, 1), "60");
    wb.set_cell_value_tracked(0, 0, 4, "=INDIRECT(\"B5\")");
    assert!(wb
        .append_table_rows(id, 1, &[(4, 1, "=E1".into())])
        .is_err());
    assert_eq!(wb.active_sheet().get_display(0, 4), "60");
    wb.clear_cell_tracked(0, 0, 4);
    wb.set_table_total(
        id,
        1,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=INDIRECT(\"B6\")".into()),
            label: None,
        },
    )
    .unwrap();
    let range = wb.table(id).unwrap().1.range;
    assert!(wb
        .resize_table(
            id,
            TableRange {
                end_row: 4,
                ..range
            }
        )
        .is_err());
    assert_eq!(wb.table(id).unwrap().1.range, range);
}

#[test]
fn dynamic_footer_spill_growth_through_sorted_records_refuses_in_manual_mode() {
    let (mut wb, id) = book();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(view.clone()))
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 4, "=IF(INDIRECT(\"B5\")=60,0,SEQUENCE(4))");
    wb.set_auto_recalc(false);
    let revision = wb.revision();
    assert!(wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.active_sheet().get_display(0, 4), "0");
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&view));
    assert!(!wb.auto_recalc());
}

#[test]
fn dynamic_footer_resize_and_unrelated_references_are_supported() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 0, 4, "=OFFSET(B5,0,0)");
    wb.set_cell_value_tracked(0, 1, 4, "=INDIRECT(\"B5\")");
    wb.set_cell_value_tracked(0, 2, 4, "=SUM(OFFSET(B2,0,0,2))");
    let range = wb.table(id).unwrap().1.range;
    let wide = wb
        .resize_table(
            id,
            TableRange {
                end_row: 5,
                end_col: 2,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.active_sheet().get_raw(0, 4), "=OFFSET(B7, 0, 0)");
    assert_eq!(wb.active_sheet().get_display(0, 4), "60");
    assert_eq!(wb.active_sheet().get_display(2, 4), "30");
    wb.apply_table_commit(&wide, true).unwrap();
    wb.clear_cell_tracked(0, 3, 0);
    wb.clear_cell_tracked(0, 3, 1);
    let shrink = wb
        .resize_table(
            id,
            TableRange {
                end_row: 2,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.active_sheet().get_raw(0, 4), "=OFFSET(B4, 0, 0)");
    assert_eq!(wb.active_sheet().get_display(0, 4), "30");
    assert_eq!(wb.active_sheet().get_display(1, 4), "");
    wb.apply_table_commit(&shrink, true).unwrap();
    assert_eq!(wb.active_sheet().get_raw(0, 4), "=OFFSET(B5,0,0)");
}

#[test]
fn newly_added_dynamic_readers_refuse_replay_of_a_previously_unlinked_move() {
    let (mut wb, id) = book();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    wb.set_cell_value_tracked(0, 0, 4, "=INDIRECT(\"B6\")");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(5));
}

#[test]
fn new_dynamic_input_uses_final_coordinates_and_settles_in_manual_mode() {
    let (mut wb, id) = book();
    wb.set_auto_recalc(false);
    // B5 is the old footer but becomes the first appended record. The new
    // B6 formula must read that record, not be rewritten to the moved B7 footer.
    let commit = wb
        .append_table_rows(
            id,
            2,
            &[(4, 1, "10".into()), (5, 1, "=OFFSET(B5,0,0)".into())],
        )
        .unwrap();
    assert!(!wb.auto_recalc());
    assert_eq!(wb.active_sheet().get_raw(5, 1), "=OFFSET(B5,0,0)");
    assert_eq!(wb.active_sheet().get_display(5, 1), "10");
    assert_eq!(wb.active_sheet().get_display(6, 1), "80");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.active_sheet().get_display(4, 1), "60");
    assert_eq!(wb.active_sheet().get_raw(5, 1), "");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(6, 1), "80");
    assert!(!wb.auto_recalc());
}
