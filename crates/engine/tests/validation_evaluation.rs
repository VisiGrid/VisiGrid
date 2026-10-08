use visigrid_engine::{
    cell::{DateStyle, NumberFormat},
    formula::eval::Value,
    named_range::NamedRange,
    sheet::{Sheet, SheetId},
    table::TableRange,
    validation::{
        CellRange, ComparisonOperator, ConstraintValue, NumericConstraint, ValidationRule,
        ValidationType,
    },
    workbook::Workbook,
};

fn book() -> Workbook {
    Workbook::from_sheets(vec![Sheet::new(SheetId(1), 100, 20)], 0)
}

fn bound(source: &str) -> NumericConstraint {
    NumericConstraint {
        operator: ComparisonOperator::LessThanOrEqual,
        value1: ConstraintValue::Formula(source.into()),
        value2: None,
    }
}

#[test]
fn validation_uses_typed_numbers_including_bounds_without_display_rounding() {
    let mut wb = book();
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set_value(0, 0, "0.1256789");
    sheet.set_value(0, 1, "0.125678");
    sheet.set_number_format(0, 0, NumberFormat::Percent { decimals: 0 });
    sheet.set_number_format(0, 1, NumberFormat::Percent { decimals: 0 });
    let rule = ValidationRule::decimal(NumericConstraint {
        operator: ComparisonOperator::LessThanOrEqual,
        value1: ConstraintValue::CellRef("B1".into()),
        value2: None,
    });
    sheet.set_cell_validation(0, 0, rule.clone());
    assert!(sheet.validate_cell(0, 0).is_invalid());
    assert!(wb.validate_cell(0, 0, 0).is_invalid());
    wb.sheet_mut(0).unwrap().set_value(0, 1, "0.125679");
    assert!(wb.validate_cell(0, 0, 0).is_valid());
    wb.sheet_mut(0).unwrap().set_text(0, 0, "0.1");
    assert!(
        wb.validate_cell(0, 0, 0).is_invalid(),
        "numeric text is still text"
    );
    wb.sheet_mut(0).unwrap().set_value(0, 0, "0.1");
    wb.sheet_mut(0).unwrap().set_text(0, 1, "0.2");
    assert!(
        wb.validate_cell(0, 0, 0).is_invalid(),
        "numeric text bounds are not numbers"
    );
}

#[test]
fn validation_date_time_and_length_use_typed_values_and_formula_bounds() {
    let mut wb = book();
    for (row, input, rule, format) in [
        (
            0,
            "=DATE(2026,10,4)",
            ValidationType::Date(bound("DATE(2026,10,5)")),
            NumberFormat::Date {
                style: DateStyle::Iso,
            },
        ),
        (
            1,
            "=TIME(12,0,0)",
            ValidationType::Time(bound("TIME(13,0,0)")),
            NumberFormat::Time,
        ),
        (
            2,
            "0.125",
            ValidationType::TextLength(NumericConstraint::equal_to(5)),
            NumberFormat::Percent { decimals: 0 },
        ),
        (
            3,
            "日本",
            ValidationType::TextLength(bound("LEN(\"日本\")")),
            NumberFormat::General,
        ),
    ] {
        wb.set_cell_value_tracked(0, row, 0, input);
        wb.sheet_mut(0)
            .unwrap()
            .set_cell_validation(row, 0, ValidationRule::new(rule));
        wb.sheet_mut(0).unwrap().set_number_format(row, 0, format);
        assert!(wb.validate_cell(0, row, 0).is_valid(), "row {row}");
        assert!(wb.sheet(0).unwrap().validate_cell(row, 0).is_valid());
    }
    assert!(wb
        .validate_cell_input(0, 0, 0, "=DATE(2026,10,6)")
        .is_invalid());
    assert!(wb
        .validate_cell_input(0, 1, 0, "=TIME(14,0,0)")
        .is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "not a date").is_invalid());
    assert!(wb.validate_cell_input(0, 1, 0, "not a time").is_invalid());
    assert!(wb.validate_cell_input(0, 3, 0, "日本語").is_invalid());
}

#[test]
fn validation_proposed_input_recalculates_cross_sheet_and_named_dependents_privately() {
    let mut wb = book();
    let other = wb.add_sheet_named("Other Sheet").unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "2");
    wb.set_cell_value_tracked(other, 0, 0, "=Sheet1!A1*2");
    wb.named_ranges_mut()
        .set(NamedRange::cell("Doubled", other, 0, 0))
        .unwrap();
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        0,
        ValidationRule::custom("=AND(A1>0,Doubled<10,'Other Sheet'!A1=A1*2)"),
    );
    let before = format!("{wb:?}");
    let revision = wb.revision();
    assert!(wb.validate_cell_input(0, 0, 0, "4").is_valid());
    assert!(wb.validate_cell_input(0, 0, 0, "5").is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "-1").is_invalid());
    assert_eq!(
        wb.sheet(other).unwrap().get_computed_value(0, 0),
        Value::Number(4.0)
    );
    assert_eq!(wb.revision(), revision);
    assert_eq!(format!("{wb:?}"), before);
}

#[test]
fn validation_custom_formula_checks_proposed_strings_with_normal_cell_typing() {
    let mut wb = book();
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        0,
        ValidationRule::custom("=AND(ISTEXT(A1),LEN(A1)=3)"),
    );
    assert!(wb.validate_cell_input(0, 0, 0, "abc").is_valid());
    assert!(wb.validate_cell_input(0, 0, 0, "abcd").is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "123").is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "=\"123\"").is_valid());
    assert!(wb
        .sheet(0)
        .unwrap()
        .validate_cell_input(0, 0, "abc")
        .is_valid());
    assert!(wb
        .sheet(0)
        .unwrap()
        .validate_cell_input(0, 0, "123")
        .is_invalid());
}

#[test]
fn validation_custom_relative_origin_and_cell_context_follow_each_target() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 1, 0, "5");
    wb.set_cell_value_tracked(0, 2, 0, "10");
    let mut rule = ValidationRule::custom("=AND(B2<$A2,ROW()=ROW(B2),COLUMN()=2)");
    rule.reference_origin = Some((1, 1));
    wb.sheet_mut(0)
        .unwrap()
        .validations
        .set(CellRange::new(1, 1, 2, 1), rule);
    assert!(wb.validate_cell_input(0, 1, 1, "7").is_invalid());
    assert!(wb.validate_cell_input(0, 2, 1, "7").is_valid());
    wb.set_cell_value_tracked(0, 2, 1, "7");
    assert!(wb.validate_cell(0, 2, 1).is_valid());
    wb.sheet_mut(0).unwrap().set_cell_validation(
        5,
        1,
        ValidationRule::decimal(bound("ROW()+COLUMN()")),
    );
    assert!(wb.validate_cell_input(0, 5, 1, "8").is_valid());
    assert!(wb.validate_cell_input(0, 5, 1, "9").is_invalid());
}

#[test]
fn validation_proposed_input_updates_spilled_dependents_and_detects_cycles() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "2");
    wb.set_cell_value_tracked(0, 0, 2, "=SEQUENCE(2,1,A1,1)");
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 0, ValidationRule::custom("=C2<5"));
    assert!(wb.validate_cell_input(0, 0, 0, "3").is_valid());
    assert!(wb.validate_cell_input(0, 0, 0, "4").is_invalid());
    assert_eq!(
        wb.sheet(0).unwrap().get_computed_value(1, 2),
        Value::Number(3.0)
    );
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 0, ValidationRule::custom("=A1>0"));
    assert!(wb.validate_cell_input(0, 0, 0, "=A1+1").is_invalid());
}

#[test]
fn validation_formula_failures_and_arrays_are_not_silent_success() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "7");
    for source in [
        "=1/0",
        "=DOESNOTEXIST(A1)",
        "=",
        "=SEQUENCE(2)",
        "=\"not logical\"",
        "=1e309",
        "=#REF!",
    ] {
        wb.sheet_mut(0)
            .unwrap()
            .set_cell_validation(0, 0, ValidationRule::custom(source));
        assert!(wb.validate_cell(0, 0, 0).is_invalid(), "{source}");
    }
    for source in ["=1/0", "=SEQUENCE(2)", "=TRUE", "=\"10\"", "=B1", "=1e309"] {
        wb.sheet_mut(0)
            .unwrap()
            .set_cell_validation(0, 0, ValidationRule::decimal(bound(source)));
        assert!(wb.validate_cell(0, 0, 0).is_invalid(), "{source}");
    }
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 0, ValidationRule::custom("=SEQUENCE(1,1,1)"));
    assert!(wb.validate_cell(0, 0, 0).is_valid());
}

#[test]
fn validation_strict_input_and_typed_whole_numbers_reject_nonfinite_and_fractional() {
    let mut wb = book();
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        0,
        ValidationRule::whole_number(NumericConstraint::between(-10, 10)),
    );
    for input in ["NaN", "inf", "-inf", "1e309", "1e-1", "3.0", "3."] {
        assert!(
            wb.validate_cell_input(0, 0, 0, input).is_invalid(),
            "{input}"
        );
        assert!(
            wb.sheet(0)
                .unwrap()
                .validate_cell_input(0, 0, input)
                .is_invalid(),
            "{input}"
        );
    }
    assert!(wb.validate_cell_input(0, 0, 0, "=3.0").is_valid());
    assert!(wb.validate_cell_input(0, 0, 0, "=3.1").is_invalid());
    wb.set_cell_value_tracked(0, 0, 0, "3.0");
    wb.sheet_mut(0)
        .unwrap()
        .set_number_format(0, 0, NumberFormat::Percent { decimals: 2 });
    assert!(
        wb.validate_cell(0, 0, 0).is_valid(),
        "stored 3 is an integer regardless of display"
    );
}

#[test]
fn validation_blank_policy_exclusions_and_range_failures_remain_consistent() {
    let mut wb = book();
    wb.sheet_mut(0).unwrap().validations.set(
        CellRange::new(0, 0, 2, 0),
        ValidationRule::custom("=FALSE").with_ignore_blank(false),
    );
    assert_eq!(wb.validate_range(0, 0, 0, 2, 0).count, 3);
    wb.sheet_mut(0)
        .unwrap()
        .validations
        .exclude(CellRange::single(1, 0));
    assert!(wb.validate_cell_input(0, 1, 0, "").is_valid());
    assert_eq!(wb.validate_range(0, 0, 0, 2, 0).count, 2);
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(4, 0, ValidationRule::custom("=FALSE"));
    assert!(wb.validate_cell(0, 4, 0).is_valid());
    assert!(wb.validate_cell_input(0, 4, 0, "").is_valid());
}

#[test]
fn validation_structured_rules_use_stored_table_record_and_preserve_totals() {
    let mut wb = book();
    for (r, values) in [["Qty", "Limit"], ["3", "5"], ["8", "6"]]
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
                start_col: 0,
                end_row: 2,
                end_col: 1,
            },
            "Orders",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let mut spec = visigrid_engine::table_view::TableViewSpec::new(id);
    let column = wb.table(id).unwrap().1.columns[0].id;
    spec.sort = Some(visigrid_engine::table_view::TableSort {
        column,
        direction: visigrid_engine::filter::SortDirection::Descending,
    });
    spec.filters.push(visigrid_engine::table_view::TableFilter {
        column,
        criteria: visigrid_engine::filter::ColumnFilter {
            selected: None,
            text_filter: Some(visigrid_engine::filter::TextFilter {
                mode: visigrid_engine::filter::TextFilterMode::Equals,
                value: "3".into(),
                case_sensitive: false,
            }),
        },
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    wb.sheet_mut(0).unwrap().validations.set(
        CellRange::new(1, 0, 2, 0),
        ValidationRule::custom("=Orders[@Qty]<=Orders[@Limit]"),
    );
    let before = format!("{wb:?}");
    assert!(wb.validate_cell(0, 1, 0).is_valid());
    assert!(wb.validate_cell(0, 2, 0).is_invalid());
    assert!(wb.validate_cell_input(0, 2, 0, "5").is_valid());
    assert!(wb.validate_cell_input(0, 1, 0, "8").is_invalid());
    assert_eq!(format!("{wb:?}"), before);
}

#[test]
fn validation_formula_constraints_see_proposed_self_value() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "100");
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 0, ValidationRule::decimal(bound("A1/2")));
    assert!(wb.validate_cell_input(0, 0, 0, "5").is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "-5").is_valid());
    assert_eq!(
        wb.sheet(0).unwrap().get_computed_value(0, 0),
        Value::Number(100.0)
    );
}

#[test]
fn validation_named_list_sources_recalculate_against_proposed_input() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "old");
    wb.set_cell_value_tracked(0, 0, 1, "=A1");
    wb.named_ranges_mut()
        .set(NamedRange::cell("Options", 0, 0, 1))
        .unwrap();
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        0,
        ValidationRule::new(ValidationType::List(
            visigrid_engine::validation::ListSource::NamedRange("Options".into()),
        )),
    );
    assert!(wb.validate_cell_input(0, 0, 0, "new").is_valid());
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 1), "old");
}

#[test]
fn validation_preview_does_not_invoke_host_custom_functions() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    use visigrid_engine::{
        custom_fns,
        formula::eval::{EvalArg, EvalResult},
    };
    static CALLS: AtomicUsize = AtomicUsize::new(0);
    fn handler(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
        if name != "VALIDATIONTESTHOOK" {
            return None;
        }
        CALLS.fetch_add(1, Ordering::SeqCst);
        Some(EvalResult::Boolean(true))
    }
    struct ResetHooks;
    impl Drop for ResetHooks {
        fn drop(&mut self) {
            custom_fns::set_default_custom_fn_handler(None);
        }
    }
    custom_fns::set_default_custom_fn_handler(Some(handler));
    let _reset = ResetHooks;
    let mut wb = book();
    wb.set_cell_value_tracked(0, 0, 0, "1");
    wb.set_cell_value_tracked(0, 0, 1, "=VALIDATIONTESTHOOK(A1)");
    assert!(CALLS.load(Ordering::SeqCst) > 0);
    CALLS.store(0, Ordering::SeqCst);
    wb.sheet_mut(0)
        .unwrap()
        .set_cell_validation(0, 0, ValidationRule::custom("=B1"));
    assert!(wb.validate_cell_input(0, 0, 0, "2").is_invalid());
    wb.sheet_mut(0).unwrap().set_cell_validation(
        0,
        0,
        ValidationRule::custom("=VALIDATIONTESTHOOK(A1)"),
    );
    assert!(wb.validate_cell(0, 0, 0).is_invalid());
    assert!(wb.validate_cell_input(0, 0, 0, "2").is_invalid());
    assert_eq!(CALLS.load(Ordering::SeqCst), 0);
}

#[test]
fn typed_date_and_time_validation_accepts_calendar_text_and_checks_bounds() {
    let mut wb = book();
    for (row, kind, valid, invalid) in [
        (0, ValidationType::Date(bound("DATE(2024,2,29)")),
         vec!["1/15/2024", "2024-01-15", "2024/2/29", "45306"],
         vec!["2024-03-01", "2/30/2024", "2/29/2023", "1/32/2024", "text"]),
        (1, ValidationType::Time(bound("TIME(13,0,0)")),
         vec!["12:30", "12:30:15", "12:30 PM", "12:30 am", "0.5"],
         vec!["14:00", "24:00", "12:60", "13:00 PM", "text"]),
    ] {
        wb.active_sheet_mut().set_cell_validation(row, 0, ValidationRule::new(kind));
        for input in valid {
            assert!(wb.validate_cell_input(0, row, 0, input).is_valid(), "{input}");
            assert!(wb.active_sheet().validate_cell_input(row, 0, input).is_valid(), "{input}");
            wb.set_cell_value_tracked(0, row, 0, input);
            assert!(wb.validate_cell(0, row, 0).is_valid(), "stored {input}");
        }
        for input in invalid {
            assert!(wb.validate_cell_input(0, row, 0, input).is_invalid(), "{input}");
        }
    }
}

#[test]
fn structural_insertion_keeps_full_grid_validation_edge_anchored() {
    use visigrid_engine::{sheet::{NUM_ROWS, NUM_COLS}, structural::Axis};
    for (axis, range) in [
        (Axis::Row, CellRange::new(0, 0, NUM_ROWS - 1, 0)),
        (Axis::Col, CellRange::new(0, 0, 0, NUM_COLS - 1)),
    ] {
        let mut wb = Workbook::new();
        let rule = ValidationRule::decimal(NumericConstraint::between(0.0, 10.0));
        wb.active_sheet_mut().validations.set(range, rule.clone());
        wb.structural_edit(0, axis, 1, 1, false).unwrap();
        wb.structural_edit(0, axis, 1, 1, true).unwrap();
        let validations = &wb.active_sheet().validations;
        assert!(validations.effective_ranges().is_ok());
        assert_eq!(validations.len(), 1);
        assert_eq!(*validations.iter().next().unwrap().0, range);
    }
}

#[test]
fn clamping_colliding_validation_rules_refuses_before_any_mutation() {
    use visigrid_engine::{sheet::{NUM_ROWS, NUM_COLS}, structural::Axis};
    for axis in [Axis::Row, Axis::Col] {
        let mut wb = Workbook::new();
        let (first, second) = if axis == Axis::Row {
            (CellRange::new(0, 0, NUM_ROWS - 2, 0), CellRange::new(0, 0, NUM_ROWS - 1, 0))
        } else {
            (CellRange::new(0, 0, 0, NUM_COLS - 2), CellRange::new(0, 0, 0, NUM_COLS - 1))
        };
        wb.active_sheet_mut().validations.set(first, ValidationRule::custom("=TRUE"));
        wb.active_sheet_mut().validations.set(second, ValidationRule::custom("=FALSE"));
        wb.active_sheet_mut().validations.exclude(CellRange::single(3, 3));
        wb.set_cell_value_tracked(0, 3, 3, "keep");
        let before = serde_json::json!([wb.active_sheet().validations.iter().collect::<Vec<_>>(), wb.active_sheet().validations.exclusions_iter().collect::<Vec<_>>()]);
        assert!(wb.structural_edit(0, axis, 1, 1, false).unwrap_err().contains("collapse two validation rules"));
        // The lower-level range transformation also refuses collisions atomically.
        assert!(wb.active_sheet_mut().validations.shift_for_structural(1, 1, false, axis == Axis::Row).unwrap_err().contains("collapse two validation rules"));
        assert_eq!(serde_json::json!([wb.active_sheet().validations.iter().collect::<Vec<_>>(), wb.active_sheet().validations.exclusions_iter().collect::<Vec<_>>()]), before);
        assert_eq!(wb.active_sheet().get_raw(3, 3), "keep");
    }
}
