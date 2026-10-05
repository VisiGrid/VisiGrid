use super::*;
use std::io::Cursor;

#[test]
fn invalid_targets_and_unsupported_rules_are_reported_without_reanchoring() {
    let xml = r#"<worksheet><dataValidations>
      <dataValidation type="custom" sqref="A0 B2:B3"><formula1>A1&gt;0</formula1></dataValidation>
      <dataValidation type="future" sqref="C1"><formula1>1</formula1></dataValidation>
      <dataValidation type="whole" sqref="D1" operator="greaterThan"><formula1>0</formula1></dataValidation>
    </dataValidations></worksheet>"#;
    let report = parse_validations_report(xml).unwrap();
    assert_eq!(report.skipped, 2);
    assert_eq!(report.rules.len(), 1);
    assert_eq!(report.rules[0].range, CellRange::single(0, 3));
}

#[test]
fn discontiguous_import_uses_one_origin_even_when_range_order_differs() {
    let xml = r#"<worksheet><dataValidations><dataValidation type="whole" operator="greaterThan" sqref="C5:C6 B2:B3"><formula1>$A5</formula1></dataValidation></dataValidations></worksheet>"#;
    let rules = parse_validations_from_xml(xml).unwrap();
    assert_eq!(rules.len(), 2);
    for entry in &rules {
        assert_eq!(entry.rule.reference_origin, Some((4, 2)));
    }
    assert_eq!(
        rules[1].rule.at(1, 1).rule_type,
        ValidationType::WholeNumber(NumericConstraint::greater_than(ConstraintValue::CellRef(
            "=$A2".into()
        )))
    );
    assert!(
        parse_validations_from_xml(&xml.replace("$A5", "$A$5")).unwrap()[0]
            .rule
            .reference_origin
            .is_some()
    );
}

#[test]
fn excluded_origin_and_overlaps_export_only_effective_rules_with_rebased_formulas() {
    use visigrid_engine::workbook::Workbook;
    let mut wb = Workbook::new();
    let s = wb.active_sheet_mut();
    let mut rule = ValidationRule::whole_number(NumericConstraint::greater_than(
        ConstraintValue::CellRef("$A2".into()),
    ));
    rule.reference_origin = Some((1, 1));
    s.validations.set(CellRange::new(1, 1, 9, 1), rule);
    s.validations.set(
        CellRange::new(4, 1, 11, 1),
        ValidationRule::whole_number(NumericConstraint::greater_than(100)),
    );
    s.validations.exclude(CellRange::new(1, 1, 2, 1));
    s.validations.exclude(CellRange::single(5, 1));
    for r in 0..12 {
        s.set_value(r, 0, &((r + 1) * 10).to_string());
    }
    let before = s.validations.clone();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("exclusions.xlsx");
    let report = crate::xlsx::export(&wb, &path, None).unwrap();
    assert_eq!(
        (report.validations_exported, report.validations_skipped),
        (3, 0)
    );
    let rules = parse_sheet_validations(&path, "Sheet1").unwrap();
    let formulas: Vec<_> = rules
        .iter()
        .map(|v| (&v.range, &v.rule.rule_type))
        .collect();
    assert!(formulas
        .iter()
        .any(|(r, t)| **r == CellRange::new(3, 1, 4, 1)
            && **t
                == ValidationType::WholeNumber(NumericConstraint::greater_than(
                    ConstraintValue::CellRef("=$A4".into())
                ))));
    let (loaded, _) = crate::xlsx::import(&path).unwrap();
    for r in 0..12 {
        assert_eq!(
            loaded.active_sheet().validations.get(r, 1).is_some(),
            before.get(r, 1).is_some(),
            "row {r}"
        );
        for value in ["5", "55", "115"] {
            assert_eq!(
                loaded.validate_cell_input(0, r, 1, value).is_valid(),
                wb.validate_cell_input(0, r, 1, value).is_valid(),
                "row {r}, value {value}"
            );
        }
    }
    assert_eq!(
        wb.active_sheet().validations,
        before,
        "export must not edit source metadata"
    );
}

#[test]
fn native_fixed_references_stay_fixed_after_excel_round_trip() {
    use visigrid_engine::workbook::Workbook;
    let mut wb = Workbook::new();
    let s = wb.active_sheet_mut();
    s.set_value(0, 0, "10");
    s.set_value(1, 0, "100");
    s.validations.set(
        CellRange::new(0, 1, 9, 1),
        ValidationRule::whole_number(NumericConstraint::greater_than(ConstraintValue::CellRef(
            "A1".into(),
        ))),
    );
    s.validations.exclude(CellRange::single(0, 1));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("fixed.xlsx");
    crate::xlsx::export(&wb, &path, None).unwrap();
    let (mut loaded, _) = crate::xlsx::import(&path).unwrap();
    for r in 1..10 {
        assert!(loaded.validate_cell_input(0, r, 1, "50").is_valid());
        assert!(loaded
            .active_sheet()
            .validations
            .get(r, 1)
            .unwrap()
            .reference_origin
            .is_some());
    }
    loaded
        .structural_edit(0, visigrid_engine::structural::Axis::Row, 0, 1, false)
        .unwrap();
    for r in 2..11 {
        assert!(loaded.validate_cell_input(0, r, 1, "50").is_valid());
    }
}

#[test]
fn relative_origins_survive_native_json_and_exclusion_edit_replay() {
    use visigrid_engine::{validation::ValidationEdit, workbook::Workbook};
    let mut wb = Workbook::new();
    let mut rule = ValidationRule::custom("=A2>0");
    rule.reference_origin = Some((1, 1));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(1, 1, 9, 1), rule);
    let patch = wb
        .active_sheet()
        .validations
        .plan_edit(
            &[CellRange::new(1, 1, 3, 1)],
            ValidationEdit::Exclude,
            1048576,
            16384,
        )
        .unwrap();
    patch
        .apply(&mut wb.active_sheet_mut().validations, true)
        .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("origin.sheet");
    crate::native::save_workbook(&wb, &path).unwrap();
    let raw: String = rusqlite::Connection::open(&path)
        .unwrap()
        .query_row(
            "SELECT value FROM meta WHERE key='validations_0'",
            [],
            |r| r.get(0),
        )
        .unwrap();
    assert_eq!(
        serde_json::from_str::<serde_json::Value>(&raw).unwrap()["version"],
        2
    );
    let native = crate::native::load_workbook(&path).unwrap();
    assert!(native.read_only_reason().is_none());
    assert_eq!(
        native.active_sheet().validations,
        wb.active_sheet().validations
    );
    for json in [
        crate::json::export_full(wb.active_sheet()).unwrap(),
        crate::json::export_workbook(&wb, &[], 0).unwrap(),
    ] {
        let mut doc: serde_json::Value = serde_json::from_str(&json).unwrap();
        assert_eq!(doc["version"], 4);
        let mut loaded = crate::json::import_any(&json).unwrap().0;
        assert_eq!(
            loaded.active_sheet().validations,
            wb.active_sheet().validations
        );
        patch
            .apply(&mut loaded.active_sheet_mut().validations, false)
            .unwrap();
        assert!(loaded.active_sheet().validations.get(1, 1).is_some());
        assert_eq!(
            loaded
                .active_sheet()
                .validations
                .get(7, 1)
                .unwrap()
                .at(7, 1)
                .rule_type,
            ValidationType::Custom("=A8>0".into())
        );
        doc["version"] = 2.into();
        assert!(crate::json::import_any(&doc.to_string())
            .unwrap_err()
            .contains("v4"));
    }
}

#[test]
fn corrupt_origins_are_rejected_by_native_and_json_loaders() {
    use visigrid_engine::workbook::Workbook;
    let mut wb = Workbook::new();
    let mut rule = ValidationRule::custom("=A2>0");
    rule.reference_origin = Some((1, 1));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(1, 1, 9, 1), rule);
    let json = crate::json::export_full(wb.active_sheet()).unwrap();
    let mut doc: serde_json::Value = serde_json::from_str(&json).unwrap();
    doc["validations"][0]["rule"]["reference_origin"] = serde_json::json!([1048576, 0]);
    assert!(crate::json::import_any(&doc.to_string())
        .unwrap_err()
        .contains("reference origin"));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("corrupt.sheet");
    crate::native::save_workbook(&wb, &path).unwrap();
    let conn = rusqlite::Connection::open(&path).unwrap();
    let raw: String = conn
        .query_row(
            "SELECT value FROM meta WHERE key='validations_0'",
            [],
            |r| r.get(0),
        )
        .unwrap();
    let mut metadata: serde_json::Value = serde_json::from_str(&raw).unwrap();
    metadata["rules"][0][1]["reference_origin"] = serde_json::json!([0, 16384]);
    conn.execute(
        "UPDATE meta SET value=?1 WHERE key='validations_0'",
        [metadata.to_string()],
    )
    .unwrap();
    assert!(crate::native::load_workbook(&path)
        .unwrap()
        .read_only_reason()
        .unwrap()
        .contains("reference origin"));
}

fn xml_for(rule: &ValidationRule) -> String {
    let mut book = rust_xlsxwriter::Workbook::new();
    book.add_worksheet()
        .add_data_validation(1, 1, 8, 1, &rule_to_xlsx(rule).unwrap())
        .unwrap();
    let bytes = book.save_to_buffer().unwrap();
    read_zip_file(
        &mut ZipArchive::new(Cursor::new(bytes)).unwrap(),
        "xl/worksheets/sheet1.xml",
    )
    .unwrap()
}

#[test]
fn all_rule_types_and_operators_keep_their_serialized_bounds() {
    for operator in [
        ComparisonOperator::Between,
        ComparisonOperator::NotBetween,
        ComparisonOperator::EqualTo,
        ComparisonOperator::NotEqualTo,
        ComparisonOperator::GreaterThan,
        ComparisonOperator::GreaterThanOrEqual,
        ComparisonOperator::LessThan,
        ComparisonOperator::LessThanOrEqual,
    ] {
        let constraint = NumericConstraint {
            operator,
            value1: ConstraintValue::Number(2_147_483_649.5),
            value2: matches!(
                operator,
                ComparisonOperator::Between | ComparisonOperator::NotBetween
            )
            .then_some(ConstraintValue::Number(9_000_000_000.75)),
        };
        for kind in [
            ValidationType::WholeNumber(constraint.clone()),
            ValidationType::Decimal(constraint.clone()),
            ValidationType::Date(constraint.clone()),
            ValidationType::Time(constraint.clone()),
            ValidationType::TextLength(constraint.clone()),
        ] {
            let rule = ValidationRule::new(kind.clone());
            let xml = xml_for(&rule);
            assert!(xml.contains("<formula1>2147483649.5</formula1>"), "{xml}");
            let loaded = parse_validations_from_xml(&xml).unwrap();
            assert_eq!(loaded.len(), 1);
            assert_eq!(loaded[0].rule.rule_type, kind);
        }
    }
}

#[test]
fn formula_constraints_and_custom_predicates_keep_entities_and_unicode() {
    let expression = "=IF(A1<10,\"A&B 日本語\",\"C>D\")";
    for kind in [
        ValidationType::Custom(expression.into()),
        ValidationType::Decimal(NumericConstraint {
            operator: ComparisonOperator::GreaterThan,
            value1: ConstraintValue::Formula(expression.into()),
            value2: None,
        }),
    ] {
        let xml = xml_for(&ValidationRule::new(kind.clone()));
        assert!(xml.contains("&lt;") && xml.contains("&amp;"));
        assert_eq!(
            parse_validations_from_xml(&xml).unwrap()[0].rule.rule_type,
            kind
        );
    }
    let xml = r#"<worksheet><dataValidations><dataValidation type="custom" sqref="A1">
        <formula1><![CDATA[IF(A1<2,"x&y",]]>&quot;z&quot;)</formula1>
        </dataValidation></dataValidations></worksheet>"#;
    assert_eq!(
        parse_validations_from_xml(xml).unwrap()[0].rule.rule_type,
        ValidationType::Custom("=IF(A1<2,\"x&y\",\"z\")".into())
    );
}

#[test]
fn dropdown_visibility_is_not_inverted_twice() {
    for show in [false, true] {
        let rule =
            ValidationRule::list_inline(vec!["Yes".into(), "No".into()]).with_show_dropdown(show);
        let xml = xml_for(&rule);
        assert_eq!(xml.contains("showDropDown=\"1\""), !show);
        assert_eq!(
            parse_validations_from_xml(&xml).unwrap()[0]
                .rule
                .show_dropdown,
            show
        );
    }
}

#[test]
fn enabled_and_disabled_messages_keep_text_and_error_style() {
    for show in [false, true] {
        for style in [
            ErrorStyle::Stop,
            ErrorStyle::Warning,
            ErrorStyle::Information,
        ] {
            let mut rule = ValidationRule::list_inline(vec!["Yes".into()]);
            rule.input_message = Some(InputMessage {
                show,
                title: "A & B < C".into(),
                message: "Line 1\n日本語 & \"quotes\"".into(),
            });
            rule.error_alert = Some(ErrorAlert {
                show,
                style,
                title: "Keep this & that".into(),
                message: "It isn't <valid>".into(),
            });
            let xml = xml_for(&rule);
            let loaded = parse_validations_from_xml(&xml).unwrap();
            assert_eq!(loaded[0].rule.input_message, rule.input_message);
            assert_eq!(loaded[0].rule.error_alert, rule.error_alert);
        }
    }
}

#[test]
fn inline_strings_retain_whitespace_quotes_and_xml_characters() {
    let rule = ValidationRule::list_inline(vec![
        "  Yes  ".into(),
        "A&B".into(),
        "\"日本語\"".into(),
        "<no>".into(),
    ]);
    let loaded = parse_validations_from_xml(&xml_for(&rule)).unwrap();
    assert_eq!(loaded[0].rule.rule_type, rule.rule_type);
}

#[test]
fn malformed_or_non_finite_bounds_are_not_silently_changed() {
    for number in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let rule = ValidationRule::decimal(NumericConstraint::between(0.0, number));
        assert!(rule_to_xlsx(&rule).is_none());
    }
    let rule = ValidationRule::whole_number(NumericConstraint {
        operator: ComparisonOperator::Between,
        value1: ConstraintValue::Number(1.0),
        value2: None,
    });
    assert!(rule_to_xlsx(&rule).is_none());
    assert!(rule_to_xlsx(&ValidationRule::custom("=")).is_none());
    assert!(parse_constraint_value("NaN").is_none());
    assert!(parse_constraint_value("inf").is_none());
    assert!(parse_list_source("\"").is_none());
    let attrs = HashMap::from([("type".into(), "custom".into())]);
    assert!(parse_single_validation(&attrs, Some("="), None).is_none());
}

#[test]
fn unrepresentable_inline_lists_are_reported_as_skipped() {
    use visigrid_engine::workbook::Workbook;
    let mut wb = Workbook::new();
    wb.active_sheet_mut().set_value(0, 0, "Data");
    for (col, items) in [vec![], vec!["one,two".into()]].into_iter().enumerate() {
        wb.active_sheet_mut().validations.set(
            CellRange::new(1, col, 5, col),
            ValidationRule::list_inline(items),
        );
    }
    let dir = tempfile::tempdir().unwrap();
    let result = crate::xlsx::export(&wb, &dir.path().join("lists.xlsx"), None).unwrap();
    assert_eq!(
        (result.validations_exported, result.validations_skipped),
        (0, 2)
    );
}

#[test]
fn malformed_and_oversized_cell_references_do_not_panic_or_wrap() {
    for cell in [
        "A0",
        "XFE1",
        "A1048577",
        "中1",
        "éA1",
        "A-1",
        "A99999999999999999999999999",
    ] {
        assert!(parse_cell_ref(cell).is_none(), "{cell}");
    }
    assert!(parse_cell_ref(&format!("{}1", "Z".repeat(200))).is_none());
    assert_eq!(parse_cell_ref("$XFD$1048576"), Some((1_048_575, 16_383)));
}

#[test]
fn prefixed_elements_and_boolean_words_are_supported() {
    let xml = r#"<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <s:dataValidations><s:dataValidation type="list" allowBlank="true" showDropDown="true"
        showInputMessage="false" prompt="A &amp; B" showErrorMessage="false" sqref="A1 B2:B3">
        <s:formula1>"Yes,No"</s:formula1></s:dataValidation></s:dataValidations></s:worksheet>"#;
    let loaded = parse_validations_from_xml(xml).unwrap();
    assert_eq!(loaded.len(), 2);
    assert!(loaded[0].rule.ignore_blank);
    assert!(!loaded[0].rule.show_dropdown);
    assert_eq!(
        loaded[0].rule.input_message.as_ref().unwrap().message,
        "A & B"
    );
    assert!(!loaded[0].rule.input_message.as_ref().unwrap().show);
    assert!(!loaded[0].rule.error_alert.as_ref().unwrap().show);
}

#[test]
fn escaped_sheet_names_and_absolute_relationships_resolve() {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let parts = [
        (
            "xl/workbook.xml",
            r#"<workbook><sheets><sheet name="A &amp; B" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet><dataValidations><dataValidation type="custom" sqref="A1"><formula1>A1&gt;0</formula1></dataValidation></dataValidations></worksheet>"#,
        ),
    ];
    use std::io::Write;
    for (name, xml) in parts {
        writer
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(xml.as_bytes()).unwrap();
    }
    let bytes = writer.finish().unwrap().into_inner();
    let mut zip = ZipArchive::new(Cursor::new(bytes)).unwrap();
    assert_eq!(
        find_worksheet_xml_path(&mut zip, "A & B").unwrap(),
        "xl/worksheets/sheet1.xml"
    );
}

#[test]
fn all_new_types_survive_public_workbook_import_and_export() {
    use visigrid_engine::workbook::Workbook;
    let mut wb = Workbook::new();
    wb.active_sheet_mut().set_name("A & B");
    for (col, kind) in [
        ValidationType::Date(NumericConstraint::between(45_000.0, 46_000.0)),
        ValidationType::Time(NumericConstraint::between(0.25, 0.75)),
        ValidationType::TextLength(NumericConstraint::between(1.0, 50.0)),
        ValidationType::Custom("=AND(A1>0,A1<10)".into()),
    ]
    .into_iter()
    .enumerate()
    {
        wb.active_sheet_mut()
            .validations
            .set(CellRange::new(1, col, 5, col), ValidationRule::new(kind));
    }
    wb.active_sheet_mut().set_value(0, 0, "1");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("validation.xlsx");
    let exported = crate::xlsx::export(&wb, &path, None).unwrap();
    assert_eq!(
        (exported.validations_exported, exported.validations_skipped),
        (4, 0)
    );
    let (loaded, result) = crate::xlsx::import(&path).unwrap();
    assert_eq!(
        (result.validations_imported, result.validations_skipped),
        (4, 0)
    );
    for ((a, original), (b, imported)) in wb
        .active_sheet()
        .validations
        .iter()
        .zip(loaded.active_sheet().validations.iter())
    {
        assert_eq!(a, b);
        assert_eq!(
            original.for_xlsx_range(a.start_row, a.start_col).rule_type,
            imported.rule_type
        );
    }
}

#[test]
fn imported_date_time_and_custom_formulas_are_evaluated_after_xlsx_roundtrip() {
    use visigrid_engine::{
        cell::{DateStyle, NumberFormat},
        workbook::Workbook,
    };
    let xml = r#"<worksheet><dataValidations>
      <dataValidation type="date" operator="lessThanOrEqual" sqref="B2:B3"><formula1>DATE(2026,10,5)</formula1></dataValidation>
      <dataValidation type="time" operator="lessThanOrEqual" sqref="C2:C3"><formula1>TIME(13,0,0)</formula1></dataValidation>
      <dataValidation type="custom" sqref="D2:D3"><formula1>AND(D2&lt;$A2,ROW()=ROW(D2))</formula1></dataValidation>
      <dataValidation type="decimal" operator="lessThanOrEqual" sqref="E2:E3"><formula1>$A2/100</formula1></dataValidation>
    </dataValidations></worksheet>"#;
    let mut wb = Workbook::new();
    for (r, row) in [
        ["5", "=DATE(2026,10,4)", "=TIME(12,0,0)", "4", "0.049"],
        ["10", "=DATE(2026,10,6)", "=TIME(14,0,0)", "11", "0.101"],
    ]
    .iter()
    .enumerate()
    {
        for (c, value) in row.iter().enumerate() {
            wb.set_cell_value_tracked(0, r + 1, c, value);
        }
        let s = wb.sheet_mut(0).unwrap();
        s.set_number_format(
            r + 1,
            1,
            NumberFormat::Date {
                style: DateStyle::Iso,
            },
        );
        s.set_number_format(r + 1, 2, NumberFormat::Time);
        s.set_number_format(r + 1, 4, NumberFormat::Percent { decimals: 0 });
    }
    for entry in parse_validations_from_xml(xml).unwrap() {
        wb.sheet_mut(0)
            .unwrap()
            .validations
            .set(entry.range, entry.rule);
    }
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("evaluated-validation.xlsx");
    let report = crate::xlsx::export(&wb, &path, None).unwrap();
    assert_eq!(
        (report.validations_exported, report.validations_skipped),
        (4, 0)
    );
    let (loaded, _) = crate::xlsx::import(&path).unwrap();
    for book in [&wb, &loaded] {
        for c in 1..=4 {
            assert!(book.validate_cell(0, 1, c).is_valid(), "column {c}");
            assert!(book.validate_cell(0, 2, c).is_invalid(), "column {c}");
        }
        assert_eq!(book.validate_range(0, 1, 1, 2, 4).count, 4);
        assert!(book.validate_cell_input(0, 2, 3, "9").is_valid());
        assert!(book.validate_cell_input(0, 2, 3, "10").is_invalid());
        assert!(book.validate_cell_input(0, 2, 4, "0.099").is_valid());
        assert!(book.validate_cell_input(0, 2, 4, "0.101").is_invalid());
    }
}
