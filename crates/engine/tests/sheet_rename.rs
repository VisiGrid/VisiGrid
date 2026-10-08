use visigrid_engine::{formula::sheets::rename_sheet_reference, workbook::Workbook};

#[test]
fn sheet_qualifiers_rewrite_without_touching_literals_or_table_headers() {
    for (source, old, new, expected) in [
        (
            "=( Data!A1 + DATA!$B$2 )*2",
            "Data",
            "New Data",
            "=( 'New Data'!A1 + 'New Data'!$B$2 )*2",
        ),
        (
            "=SUM(Data!A:A,Data!1:3)",
            "Data",
            "Δ-1",
            "=SUM('Δ-1'!A:A,'Δ-1'!1:3)",
        ),
        (
            "='O''Brien'!A1+\"O'Brien!A1\"+Sales[O''Brien!A1]",
            "O'Brien",
            "New'Name",
            "='New''Name'!A1+\"O'Brien!A1\"+Sales[O''Brien!A1]",
        ),
        (
            "=Other!A1+Data!B2+INDIRECT(\"Data!B2\")",
            "Data",
            "Next",
            "=Other!A1+'Next'!B2+INDIRECT(\"Data!B2\")",
        ),
        ("Data!A1>0", "Data", "Next", "'Next'!A1>0"),
    ] {
        assert_eq!(rename_sheet_reference(source, old, new).unwrap(), expected);
    }
    assert_eq!(rename_sheet_reference("=Data!A1+@", "Data", "Next").unwrap(), "=Next!A1+@");
    assert_eq!(
        rename_sheet_reference("=UNSUPPORTED(@)", "Data", "Next").unwrap(),
        "=UNSUPPORTED(@)"
    );
}

#[test]
fn high_level_rename_keeps_values_names_and_dependency_updates() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet_named("Data").unwrap();
    let id = wb.sheet(data).unwrap().id;
    wb.set_cell_value_tracked(data, 0, 0, "7");
    wb.define_name_for_cell("Input", data, 0, 0).unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "=( Data!A1 + Input )*2");
    let (mut renamed, commit) = wb.prepare_sheet_rename(id, "Data", "O'Brien-Δ").unwrap();
    assert_eq!(wb.sheet(data).unwrap().name, "Data");
    assert_eq!(
        renamed.sheet(0).unwrap().get_raw(0, 0),
        "=( 'O''Brien-Δ'!A1 + Input )*2"
    );
    assert_eq!(renamed.sheet(0).unwrap().get_display(0, 0), "28");
    assert_eq!(
        renamed.get_named_range("Input"),
        wb.get_named_range("Input")
    );
    commit.replay(&mut renamed, true).unwrap();
    assert_eq!(renamed.sheet(data).unwrap().name, "Data");
    assert_eq!(
        renamed.sheet(0).unwrap().get_raw(0, 0),
        "=( Data!A1 + Input )*2"
    );
    commit.replay(&mut renamed, false).unwrap();
    renamed.set_cell_value_tracked(data, 0, 0, "8");
    assert_eq!(renamed.sheet(0).unwrap().get_display(0, 0), "32");
    assert!(commit.replay(&mut renamed, true).is_err());
}

#[test]
fn label_only_and_case_only_renames_have_history_but_exact_noops_do_not() {
    let wb = Workbook::new();
    let id = wb.active_sheet_id();
    let (_, noop) = wb.prepare_sheet_rename(id, "Sheet1", " Sheet1 ").unwrap();
    assert!(noop.is_empty());
    let (mut candidate, commit) = wb.prepare_sheet_rename(id, "Sheet1", "sheet1").unwrap();
    assert!(!commit.is_empty());
    assert_eq!(commit.changed_cell_count(), 0);
    commit.replay(&mut candidate, true).unwrap();
    assert_eq!(candidate.active_sheet().name, "Sheet1");
    commit.replay(&mut candidate, false).unwrap();
    assert_eq!(candidate.active_sheet().name, "sheet1");
    assert!(
        wb.capture_guarded_batch(&candidate).is_err(),
        "ordinary batches must still refuse sheet identity changes"
    );
}

#[test]
fn invalid_names_refuse_and_unsupported_formula_syntax_is_preserved() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet_named("Data").unwrap();
    let id = wb.sheet(data).unwrap().id;
    wb.set_cell_value_tracked(0, 0, 0, "=Data!A1+1");
    wb.set_cell_value_tracked(0, 1, 0, "=Data!A1+@");
    let (renamed, _) = wb.prepare_sheet_rename(id, "Data", "Next").unwrap();
    assert_eq!(renamed.sheet(0).unwrap().get_raw(1, 0), "=Next!A1+@");
    assert_eq!(wb.sheet(0).unwrap().get_raw(0, 0), "=Data!A1+1");
    assert_eq!(wb.sheet(data).unwrap().name, "Data");
    for name in ["", "  ", "Sheet1", "12345678901234567890123456789012"] {
        assert!(wb.prepare_sheet_rename(id, "Data", name).is_err());
    }
    assert!(wb.prepare_sheet_rename(id, "Stale", "Next").is_err());
}

#[test]
fn a_new_name_cannot_turn_an_unresolved_reference_into_a_cycle() {
    let mut wb = Workbook::new();
    let id = wb.active_sheet_id();
    wb.set_cell_value_tracked(0, 0, 0, "=Future!A1");
    assert!(wb
        .prepare_sheet_rename(id, "Sheet1", "Future")
        .unwrap_err()
        .contains("cycle"));
    assert_eq!(wb.active_sheet().name, "Sheet1");
    assert_eq!(wb.active_sheet().get_raw(0, 0), "=Future!A1");
}

#[test]
fn unsupported_reference_islands_preserve_strings_headers_and_external_owners() {
    use visigrid_engine::formula::sheets::delete_sheet_references;
    for (source, expected) in [
        ("=Data!A1#+@B2", "=Next!A1#+@B2"),
        ("=SUM(Data!A1:Data!B2)+@", "=SUM(Next!A1:B2)+@"),
        ("=Data!A1+\"Data!A1\"+Sales[Data!A1]+@", "=Next!A1+\"Data!A1\"+Sales[Data!A1]+@"),
        ("=[Book.xlsx]Data!A1+Data!B2+@", "=[Book.xlsx]Data!A1+Next!B2+@"),
        ("=Data!A1+#REF!+@", "=Next!A1+#REF!+@"),
    ] {
        assert_eq!(rename_sheet_reference(source, "Data", "Next").unwrap(), expected);
    }
    assert_eq!(delete_sheet_references("=Data!A1:Data!B2+@", "Data", &[]).unwrap(), "=#REF!+@");
    assert_eq!(delete_sheet_references("=[Book.xlsx]Data!A1+Data!B2+@", "Data", &[]).unwrap(), "=[Book.xlsx]Data!A1+#REF!+@");
    assert!(delete_sheet_references("=Data:Other!A1+@", "Data", &[]).is_err());
    assert!(rename_sheet_reference("=Data!A1:Other!A2+@", "Data", "Next").is_err());
}
