use super::*;
use std::io::Cursor;

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
        assert_eq!(original.rule_type, imported.rule_type);
    }
}
