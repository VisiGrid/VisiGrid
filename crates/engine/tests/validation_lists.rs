use visigrid_engine::{
    cell::NumberFormat,
    named_range::NamedRange,
    sheet::{Sheet, SheetId, NUM_ROWS},
    table::TableRange,
    validation::{CellRange, ListSource, ValidationRule, ValidationType, MAX_LIST_ITEMS},
    workbook::Workbook,
};

fn book() -> Workbook {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(1), NUM_ROWS, 20)], 0);
    for (r, text) in ["Red", "Green", "Blue"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, text);
    }
    wb
}
fn source(wb: &mut Workbook, text: &str) {
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 4, ValidationRule::list_range(text));
}
fn items(wb: &Workbook) -> Vec<String> {
    let result = wb.get_list_items(0, 0, 4).unwrap();
    assert!(result.source_error.is_none(), "{:?}", result.source_error);
    result.items
}

#[test]
fn dynamic_references_keep_value_labels_and_sparse_whole_columns() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 1, "0.25");
    wb.sheet_mut(0)
        .unwrap()
        .set_number_format(0, 1, NumberFormat::Percent { decimals: 0 });
    for formula in [
        "=OFFSET(A1,0,0,COUNTA(A:A),1)",
        "=INDIRECT(\"A1:A3\")",
        "=INDIRECT(\"A:A\")",
        "=OFFSET(A1,0,0,1048576,1)",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["Red", "Green", "Blue"], "{formula}");
    }
    source(&mut wb, "=OFFSET(B1,0,0)");
    assert_eq!(items(&wb), ["0.25"]);
    let sheet = wb.sheet(0).unwrap();
    assert_eq!(sheet.get_list_items(0, 4).unwrap().items, ["0.25"]);
}

#[test]
fn indirect_resolves_quoted_sheets_names_and_tables() {
    let mut wb = book();
    let other = wb.add_sheet_named("Options! O'Brien").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "West");
    wb.set_cell_value_tracked(other, 1, 0, "East");
    wb.named_ranges_mut()
        .set(NamedRange::range("Regions", other, 0, 0, 1, 0))
        .unwrap();
    for formula in [
        "=INDIRECT(\"'Options! O''Brien'!A1:A2\")",
        "=INDIRECT(\"Regions\")",
        "=OFFSET(Regions,0,0)",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["West", "East"], "{formula}");
    }
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        },
        "Colors",
    )
    .unwrap();
    for formula in [
        "=Colors[Red]",
        "=Colors",
        "=INDIRECT(\"Colors[Red]\")",
        "=OFFSET(Colors[Red],0,0)",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["Green", "Blue"], "{formula}");
    }
}

#[test]
fn dependent_reference_choices_and_relative_origins_use_the_target_record() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 2, "1");
    wb.set_cell_value_tracked(0, 1, 2, "2");
    let mut rule = ValidationRule::list_range("=OFFSET($A$1,C1-1,0,1,1)");
    rule.reference_origin = Some((0, 4));
    wb.sheet_mut(0)
        .unwrap()
        .validations
        .set(CellRange::new(0, 4, 1, 4), rule);
    assert_eq!(wb.get_list_items(0, 0, 4).unwrap().items, ["Red"]);
    assert_eq!(wb.get_list_items(0, 1, 4).unwrap().items, ["Green"]);
    assert!(wb.validate_cell_input(0, 1, 4, "Green").is_valid());
    assert!(wb.validate_cell_input(0, 1, 4, "Red").is_invalid());
    for (formula, expected) in [
        ("=IF(C1=1,A1:A2,A3)", vec!["Red", "Green"]),
        ("=CHOOSE(C1,A1:A2,A3)", vec!["Red", "Green"]),
        ("=INDEX(A1:A3,2)", vec!["Green"]),
        ("=INDEX(A1:B3,0,1)", vec!["Red", "Green", "Blue"]),
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), expected, "{formula}");
    }
}

#[test]
fn array_formulas_produce_choices_and_match_formatted_numeric_targets() {
    let mut wb = book();
    source(&mut wb, "=SORT(UNIQUE(A1:A3))");
    assert_eq!(items(&wb), ["Blue", "Green", "Red"]);
    source(&mut wb, "=FILTER(A1:A3,A1:A3<>\"Green\")");
    assert_eq!(items(&wb), ["Red", "Blue"]);
    source(&mut wb, "=TRANSPOSE(A1:A3)");
    assert_eq!(items(&wb), ["Red", "Green", "Blue"]);
    source(&mut wb, "=SEQUENCE(3,1,0.25,0.25)");
    wb.set_cell_value_tracked(0, 0, 4, "0.25");
    wb.sheet_mut(0)
        .unwrap()
        .set_number_format(0, 4, NumberFormat::Percent { decimals: 0 });
    assert!(wb.validate_cell(0, 0, 4).is_valid());
    assert!(wb.validate_cell_input(0, 0, 4, "0.5").is_valid());
    assert!(wb.validate_cell_input(0, 0, 4, "0.6").is_invalid());
}

#[test]
fn broken_sources_do_not_turn_into_permissive_empty_lists() {
    let mut wb = book();
    for formula in [
        "=MissingName",
        "=Missing!A1",
        "=INDIRECT(\"Missing\")",
        "=INDIRECT(\"A1+A2\")",
        "=INDIRECT(\"R1C1\",FALSE)",
        "=OFFSET(A1,-1,0)",
        "=OFFSET(A1,0,0,1e300)",
        "=SEQUENCE(2,2)",
        "=DOESNOTEXIST()",
        "=1/0",
        "=A0",
    ] {
        source(&mut wb, formula);
        let list = wb.get_list_items(0, 0, 4).unwrap();
        assert!(list.source_error.is_some(), "{formula}: {list:?}");
        assert!(list.items.is_empty());
        assert!(
            wb.validate_cell_input(0, 0, 4, "anything").is_invalid(),
            "{formula}"
        );
    }
    source(&mut wb, "=B1:B3");
    assert!(wb.get_list_items(0, 0, 4).unwrap().source_error.is_none());
    assert!(
        wb.validate_cell_input(0, 0, 4, "anything").is_valid(),
        "valid empty ranges retain the existing policy"
    );
}

#[test]
fn list_array_limits_are_checked_before_allocation_and_reset_afterward() {
    let mut wb = book();
    for formula in [
        "=SEQUENCE(1000000000)",
        "=SEQUENCE(1000000000,1000000000)",
        "=TRANSPOSE(A1:XFD1048576)",
        "=SORT(OFFSET(A1,0,0,1048576))",
        "=SEQUENCE(1000)+TRANSPOSE(SEQUENCE(1000))",
    ] {
        source(&mut wb, formula);
        let list = wb.get_list_items(0, 0, 4).unwrap();
        assert!(
            list.source_error
                .as_ref()
                .is_some_and(|e| e.contains("limit")),
            "{formula}: {list:?}"
        );
    }
    source(&mut wb, "=SEQUENCE(10001)");
    let list = wb.get_list_items(0, 0, 4).unwrap();
    assert!(list.source_error.is_none());
    assert_eq!(list.items.len(), MAX_LIST_ITEMS);
    assert!(list.is_truncated);
    source(&mut wb, "=SEQUENCE(2)");
    assert_eq!(items(&wb), ["1", "2"]);
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        4,
        ValidationRule::custom("=SEQUENCE(1000000000)"),
    );
    wb.set_cell_value_tracked(0, 0, 4, "1");
    assert!(wb.validate_cell(0, 0, 4).is_invalid());
}

#[test]
fn legacy_named_formula_sources_still_resolve_and_failures_change_fingerprints() {
    let mut wb = book();
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        4,
        ValidationRule::new(ValidationType::List(ListSource::NamedRange(
            "INDIRECT(\"A1:A3\")".into(),
        ))),
    );
    assert_eq!(items(&wb), ["Red", "Green", "Blue"]);
    let before = wb.get_list_items(0, 0, 4).unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "=1/0");
    let after = wb.get_list_items(0, 0, 4).unwrap();
    assert!(after.source_error.is_some());
    assert_ne!(before.source_fingerprint, after.source_fingerprint);
}

#[test]
fn reference_lists_include_virtual_spill_receivers_without_duplicates() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 1, "=SEQUENCE(3)");
    for formula in ["=B1:B3", "=INDIRECT(\"B1:B3\")", "=OFFSET(B1,0,0,3)"] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["1", "2", "3"], "{formula}");
    }
    wb.sheet_mut(0)
        .unwrap()
        .set_number_format(1, 1, NumberFormat::Percent { decimals: 0 });
    assert_eq!(items(&wb), ["1", "2", "3"]);
}

#[test]
fn source_fingerprint_includes_truncation_even_when_visible_choices_match() {
    let mut wb = book();
    source(&mut wb, "=SEQUENCE(10000)");
    let exact = wb.get_list_items(0, 0, 4).unwrap();
    source(&mut wb, "=SEQUENCE(10001)");
    let truncated = wb.get_list_items(0, 0, 4).unwrap();
    assert_eq!(exact.items, truncated.items);
    assert_ne!(exact.source_fingerprint, truncated.source_fingerprint);
}

#[test]
fn new_array_functions_keep_dropdown_context_and_allocation_limits() {
    let mut wb = book();
    for formula in [
        "=LET(values,INDIRECT(\"A1:A3\"),SORTBY(values,values))",
        "=LET(read,LAMBDA(target,SORT(INDIRECT(target))),read(\"A1:A3\"))",
        "=TEXTSPLIT(\"Blue,Green,Red\",\",\")",
        "={\"Blue\",\"Green\",\"Red\"}",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["Blue", "Green", "Red"], "{formula}");
    }
    // Each individual allocation is permitted, but the nested sorts together
    // must share the validation budget rather than reset it at LET boundaries.
    let mut formula = "SEQUENCE(100000)".to_string();
    for _ in 0..12 {
        formula = format!("LET(values,{formula},SORTBY(values,values))");
    }
    source(&mut wb, &format!("={formula}"));
    let result = wb.get_list_items(0, 0, 4).unwrap();
    assert!(result.source_error.as_deref().is_some_and(|e| e.contains("array limit")), "{:?}", result.source_error);

    wb.set_cell_text_tracked(0, 0, 1, &"x,".repeat(100_001));
    source(&mut wb, "=TEXTSPLIT(B1,\",\")");
    let result = wb.get_list_items(0, 0, 4).unwrap();
    assert!(result.source_error.as_deref().is_some_and(|e| e.contains("array limit")), "{:?}", result.source_error);
    source(&mut wb, "=TEXTSPLIT(\"Blue,Green\",\",\")");
    assert_eq!(items(&wb), ["Blue", "Green"]);
}

#[test]
fn namespaced_functions_evaluate_without_changing_literals_or_stored_source() {
    let mut wb = book();
    let source_text = "=_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))";
    source(&mut wb, source_text);
    assert_eq!(items(&wb), ["Blue", "Green", "Red"]);
    assert_eq!(
        wb.sheet(0)
            .unwrap()
            .validations
            .get(0, 4)
            .unwrap()
            .rule_type,
        ValidationType::List(ListSource::Range(source_text.into()))
    );
    source(&mut wb, "=\"_xlfn.UNIQUE(A1)\"");
    assert_eq!(items(&wb), ["_xlfn.UNIQUE(A1)"]);
    source(&mut wb, "=A1:A3");
    assert!(wb
        .validate_cell_input(0, 0, 4, "=SEQUENCE(1000000000)")
        .is_invalid());
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 4), "");
}

#[test]
fn computed_lists_share_named_cross_sheet_and_composed_reference_resolution() {
    let mut wb = book();
    let other = wb.add_sheet_named("Options! O'Brien").unwrap();
    for (r, text) in ["West", "East", "West"].iter().enumerate() {
        wb.set_cell_value_tracked(other, r, 0, text);
    }
    wb.define_name_for_range("Regions", other, 0, 0, 2, 0)
        .unwrap();
    for formula in [
        "=SORT(UNIQUE(INDIRECT(\"Regions\")))",
        "=SORT(UNIQUE(INDIRECT(\"'Options! O''Brien'!A1:A3\")))",
        "=SORT(UNIQUE(OFFSET(INDIRECT(\"Regions\"),0,0)))",
        "=SORT(UNIQUE(OFFSET(IF(TRUE,Regions,A1),0,0)))",
        "=SORT(UNIQUE(OFFSET(CHOOSE(2,A1,Regions),0,0)))",
        "=SORT(OFFSET(INDEX(Regions,0,1),0,0,2))",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["East", "West"], "{formula}");
    }
    wb.set_cell_value_tracked(other, 0, 0, "North");
    assert_eq!(items(&wb), ["East", "North"]);
    assert!(wb.validate_cell_input(0, 0, 4, "North").is_valid());
    assert!(wb.validate_cell_input(0, 0, 4, "West").is_invalid());
}

#[test]
fn computed_lists_resolve_indirect_table_columns_and_current_record_context() {
    let mut wb = book();
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 2,
            end_col: 0,
        },
        "Colors",
    )
    .unwrap();
    for formula in [
        "=SORT(INDIRECT(\"Colors[Red]\"))",
        "=SORT(OFFSET(INDIRECT(\"Colors[Red]\"),0,0))",
        "=SORT(OFFSET(Colors[Red],0,0))",
    ] {
        source(&mut wb, formula);
        assert_eq!(items(&wb), ["Blue", "Green"], "{formula}");
    }
    source(&mut wb, "=SORT(OFFSET(INDIRECT(\"A1:A3\"),ROW(),0,2))");
    assert_eq!(items(&wb), ["Blue", "Green"]);
    assert_eq!(
        wb.sheet(0).unwrap().get_list_items(0, 4).unwrap().items,
        ["Blue", "Green"]
    );
}

#[test]
fn nested_reference_failures_keep_errors_and_array_limits() {
    let mut wb = book();
    for formula in [
        "=SORT(INDIRECT(\"A1:A3\",FALSE))",
        "=SORT(INDIRECT(\"Missing!A1:A3\"))",
        "=SORT(INDIRECT(\"SUM(A1:A3)\"))",
        "=SORT(INDIRECT(\"[External.xlsx]Sheet1!A1\"))",
        "=SORT(OFFSET(INDIRECT(\"A1:A3\"),-1,0))",
        "=SORT(OFFSET(A:A,1,0))",
        "=SORT(OFFSET(INDIRECT(\"A1:A3\"),0,0,1048576))",
    ] {
        source(&mut wb, formula);
        let result = wb.get_list_items(0, 0, 4).unwrap();
        assert!(result.source_error.is_some(), "{formula}: {result:?}");
        assert!(
            wb.validate_cell_input(0, 0, 4, "Red").is_invalid(),
            "{formula}"
        );
    }
    source(&mut wb, "=SORT(INDIRECT(\"A1:A3\"))");
    assert_eq!(items(&wb), ["Blue", "Green", "Red"]);
}
