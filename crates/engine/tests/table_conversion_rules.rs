use visigrid_engine::{
    cell::CellStyle,
    cond_format::CondStyle,
    formula::parser,
    table::{TableId, TableRange},
    validation::{CellRange, ValidationRule, ValidationType},
    workbook::Workbook,
};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (r, values) in [
        ["Amount", "Limit"],
        ["10", "15"],
        ["20", "15"],
        ["30", "40"],
    ]
    .iter()
    .enumerate()
    {
        for (c, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, r + 1, c + 1, value);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 1,
                end_row: 4,
                start_col: 1,
                end_col: 2,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    (wb, id)
}
fn metadata(wb: &Workbook) -> serde_json::Value {
    serde_json::to_value((
        &wb.active_sheet().cond_formats,
        wb.active_sheet().validations.iter().collect::<Vec<_>>(),
        wb.active_sheet()
            .validations
            .exclusions_iter()
            .collect::<Vec<_>>(),
    ))
    .unwrap()
}
fn matched(wb: &Workbook) -> Vec<bool> {
    let sheet = wb.active_sheet();
    (0..8)
        .flat_map(|r| {
            (0..5).map(move |c| sheet.cond_formats.override_for_cell(r, c, sheet).is_some())
        })
        .collect()
}

#[test]
fn local_and_named_this_row_rules_keep_cell_context_across_table_boundaries() {
    for source in [
        "=IFERROR([@Amount]<[@Limit],FALSE)",
        "=IFERROR(Sales[@Amount]<Sales[@Limit],FALSE)",
    ] {
        let (mut wb, id) = book();
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::new(0, 0, 7, 4)],
            source,
            CondStyle::Named(CellStyle::Warning),
        );
        let before = metadata(&wb);
        let expected = matched(&wb);
        assert!(expected.iter().any(|x| *x));
        let commit = wb.remove_table(id).unwrap();
        assert_eq!(matched(&wb), expected, "{source}");
        assert!(wb
            .active_sheet()
            .cond_formats
            .iter()
            .all(|r| r.parse_error().is_none()));
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(metadata(&wb), before);
        wb.apply_table_commit(&commit, false).unwrap();
        assert_eq!(matched(&wb), expected);
    }
}

#[test]
fn conditional_format_overlap_anchors_order_and_disabled_state_survive() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(2, 1, 6, 3), CellRange::new(0, 0, 4, 4)],
        "=AND(SUM(Sales[Amount])=60,B3>15,ROW()>2)",
        CondStyle::Named(CellStyle::Warning),
    );
    let disabled = wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 0, 7, 4)],
        "=SUM(Sales[Amount])>0",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .cond_formats
        .get_mut(disabled)
        .unwrap()
        .enabled = false;
    let before = metadata(&wb);
    let expected = matched(&wb);
    let commit = wb.remove_table(id).unwrap();
    assert_eq!(matched(&wb), expected);
    let last = wb.active_sheet().cond_formats.iter().last().unwrap();
    assert!(!last.enabled);
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), before);
}

#[test]
fn validation_origins_fixed_native_refs_and_exclusions_keep_per_cell_results() {
    for imported in [false, true] {
        let (mut wb, id) = book();
        for row in 0..8 {
            wb.set_cell_value_tracked(0, row, 0, if row % 2 == 0 { "1" } else { "0" });
        }
        let mut rule = ValidationRule::custom(if imported {
            "=AND(Sales[@Amount]<Sales[@Limit],$A7=1)"
        } else {
            "=AND(Sales[@Amount]<Sales[@Limit],A1=1)"
        });
        if imported {
            rule.reference_origin = Some((6, 4));
        }
        rule.ignore_blank = false;
        wb.active_sheet_mut()
            .validations
            .set(CellRange::new(0, 0, 7, 4), rule);
        wb.active_sheet_mut()
            .validations
            .exclude(CellRange::single(3, 2));
        let before = metadata(&wb);
        let expected: Vec<_> = (0..8)
            .flat_map(|r| (0..5).map(move |c| (r, c)))
            .map(|(r, c)| wb.validate_cell(0, r, c).is_valid())
            .collect();
        assert!(expected.iter().any(|valid| *valid));
        let commit = wb.remove_table(id).unwrap();
        let actual: Vec<_> = (0..8)
            .flat_map(|r| (0..5).map(move |c| (r, c)))
            .map(|(r, c)| wb.validate_cell(0, r, c).is_valid())
            .collect();
        assert_eq!(actual, expected, "imported {imported}");
        assert!(wb.active_sheet().validations.get(3, 2).is_none());
        if !imported {
            let rule = wb.active_sheet().validations.get(2, 1).unwrap().at(2, 1);
            let ValidationType::Custom(source) = &rule.rule_type else {
                panic!()
            };
            assert!(source.contains("$A$1"));
            assert!(source.contains("$B3"));
        }
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(metadata(&wb), before);
    }
}

#[test]
fn overlapping_validation_precedence_and_cross_sheet_dropdowns_are_preserved() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().validations.set(
        CellRange::new(0, 0, 3, 3),
        ValidationRule::list_inline(vec!["10".into()]),
    );
    wb.active_sheet_mut().validations.set(
        CellRange::new(1, 1, 7, 4),
        ValidationRule::list_range("Sales[Amount]"),
    );
    let other = wb.add_sheet_named("Other").unwrap();
    wb.sheet_mut(other).unwrap().validations.set(
        CellRange::new(0, 0, 2, 0),
        ValidationRule::list_range("Sales[Amount]"),
    );
    let before = metadata(&wb);
    let expected: Vec<_> = (0..8)
        .flat_map(|r| (0..5).map(move |c| (r, c)))
        .map(|(r, c)| wb.get_list_items(0, r, c).map(|r| r.items))
        .collect();
    let commit = wb.remove_table(id).unwrap();
    let actual: Vec<_> = (0..8)
        .flat_map(|r| (0..5).map(move |c| (r, c)))
        .map(|(r, c)| wb.get_list_items(0, r, c).map(|r| r.items))
        .collect();
    assert_eq!(actual, expected);
    assert_eq!(
        wb.get_list_items(other, 1, 0).unwrap().items,
        ["10", "20", "30"]
    );
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), before);
}

#[test]
fn rule_conversion_is_bounded_and_invalid_metadata_is_atomic() {
    let (mut wb, id) = book();
    for _ in 0..4097 {
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(2, 1)],
            "=SUM(Sales[Amount])>0",
            CondStyle::Named(CellStyle::Warning),
        );
    }
    let before = metadata(&wb);
    let revision = wb.revision();
    assert!(wb.remove_table(id).unwrap_err().contains("too much"));
    assert_eq!(wb.revision(), revision);
    assert_eq!(metadata(&wb), before);
    assert!(wb.table(id).is_some());
}

#[test]
fn standard_error_literals_roundtrip_and_remain_catchable_inside_formulas() {
    let mut wb = Workbook::new();
    for error in [
        "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!",
    ] {
        let source = format!("=IFERROR({error},42)");
        let expr = parser::parse(&source).unwrap();
        assert!(parser::parse(&parser::format_parsed_expr(&expr)).is_ok());
        wb.set_cell_value_tracked(0, 0, 0, &source);
        assert_eq!(wb.active_sheet().get_display(0, 0), "42", "{error}");
        wb.set_cell_value_tracked(0, 0, 0, &format!("={}", error.to_lowercase()));
        assert_eq!(wb.active_sheet().get_display(0, 0), error);
    }
    for source in ["=#NAME?garbage", "=#N/A1", "=#UNKNOWN!", "=#値!"] {
        assert!(parser::parse(source).is_err(), "{source}");
    }
}

#[test]
fn whole_sheet_rules_and_this_row_column_spans_use_bounded_rectangles() {
    use visigrid_engine::sheet::{NUM_COLS, NUM_ROWS};
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 0, NUM_ROWS - 1, NUM_COLS - 1)],
        "=IFERROR(SUM(Sales[@])>25,FALSE)",
        CondStyle::Named(CellStyle::Warning),
    );
    let before = matched(&wb);
    wb.remove_table(id).unwrap();
    assert_eq!(matched(&wb), before);
    assert!(wb.active_sheet().cond_formats.len() <= 15);
    assert!(!wb
        .active_sheet()
        .has_cond_format(NUM_ROWS - 1, NUM_COLS - 1));
}

#[test]
fn local_rules_keep_another_tables_binding_and_stale_metadata_refuses_replay() {
    let (mut wb, id) = book();
    for (r, data) in [["Amount", "Limit"], ["100", "90"], ["20", "30"]]
        .iter()
        .enumerate()
    {
        for (c, value) in data.iter().enumerate() {
            wb.set_cell_value_tracked(0, r + 8, c + 1, value);
        }
    }
    let other = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 8,
                end_row: 10,
                start_col: 1,
                end_col: 2,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 0, 12, 4)],
        "=IFERROR([@Amount]<[@Limit],FALSE)",
        CondStyle::Named(CellStyle::Warning),
    );
    let commit = wb.remove_table(id).unwrap();
    assert!(!wb.active_sheet().has_cond_format(9, 1));
    assert!(wb.active_sheet().has_cond_format(10, 1));
    assert!(wb.table(other).is_some());
    wb.active_sheet_mut().validations.set(
        CellRange::single(20, 0),
        ValidationRule::list_inline(vec!["Changed".into()]),
    );
    let before = metadata(&wb);
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(metadata(&wb), before);
    assert!(wb.table(id).is_none());
}

#[test]
fn invalid_rule_geometry_refuses_conversion_without_partial_rewrites() {
    use visigrid_engine::sheet::NUM_ROWS;
    for origin in [false, true] {
        let (mut wb, id) = book();
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(2, 1)],
            "=[@Amount]>0",
            CondStyle::Named(CellStyle::Warning),
        );
        let mut rule = ValidationRule::list_range("Sales[Amount]");
        let range = if origin {
            rule.reference_origin = Some((NUM_ROWS, 0));
            CellRange::single(8, 0)
        } else {
            CellRange::single(NUM_ROWS, 0)
        };
        wb.active_sheet_mut().validations.set(range, rule);
        let before = metadata(&wb);
        let revision = wb.revision();
        assert!(wb.remove_table(id).unwrap_err().contains("outside"));
        assert_eq!(metadata(&wb), before);
        assert_eq!(wb.revision(), revision);
        assert!(wb.table(id).is_some());
    }
}
