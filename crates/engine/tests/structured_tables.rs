use visigrid_engine::{
    cell::CellValue,
    formula::{
        parser::{self, Expr},
        structured::{StructuredReference, TableSection},
    },
    sheet::{Sheet, SheetId},
    table::{TableId, TableRange},
    workbook::Workbook,
};
fn range(end_row: usize, end_col: usize) -> TableRange {
    TableRange {
        start_row: 0,
        start_col: 0,
        end_row,
        end_col,
    }
}
fn fixture() -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(
        vec![
            Sheet::new_with_name(SheetId(1), 100, 20, "Data"),
            Sheet::new_with_name(SheetId(2), 100, 20, "Report"),
        ],
        0,
    );
    for (c, h) in ["Qty", "Price", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, c, h);
    }
    for (r, q, p) in [(1, "2", "10"), (2, "3", "20")] {
        wb.set_cell_value_tracked(0, r, 0, q);
        wb.set_cell_value_tracked(0, r, 1, p);
    }
    let id = wb
        .create_table(SheetId(1), range(2, 2), "Sales")
        .unwrap()
        .table_id();
    wb.set_cell_value_tracked(0, 1, 2, "=[@Qty]*[@Price]");
    wb.set_cell_value_tracked(0, 2, 2, "=[@[Qty]]*[@[Price]]");
    (wb, id)
}
fn display(wb: &Workbook, s: usize, r: usize, c: usize) -> String {
    wb.sheet(s).unwrap().get_display(r, c)
}
fn eval(wb: &mut Workbook, formula: &str) -> String {
    wb.set_cell_value_tracked(1, 20, 10, formula);
    display(wb, 1, 20, 10)
}

#[test]
fn structured_grammar_and_escaping_roundtrip() {
    for text in [
        "Sales[Amount]",
        "Sales[#Data]",
        "Sales[#All]",
        "Sales[#Headers]",
        "Sales[[Qty]:[Amount]]",
        "Sales[[#Headers],[Amount]]",
        "[@Amount]",
        "[@[Unit Price]]",
        "[Amount]",
        "Sales[@Qty]",
        "Sales[[#This Row],[Qty]]",
        "Sales[[#This Row],[Qty]:[Amount]]",
        "[#This Row]",
        "_Sales[Amount]",
        "Sales[ [Qty]:[Amount] ]",
    ] {
        let parsed = parser::parse(&format!("={text}")).unwrap();
        let formatted = parser::format_parsed_expr(&parsed);
        assert_eq!(
            format!("{:?}", parsed),
            format!("{:?}", parser::parse(&formatted).unwrap()),
            "{text}"
        );
    }
    for name in [
        "Tax [net]",
        "#Data",
        "@Qty",
        "O'Brien",
        "金額",
        "A1",
        "a,b:c",
    ] {
        let reference = StructuredReference {
            table: Some("Sales".into()),
            section: TableSection::Data,
            columns: Some((name.into(), name.into())),
        };
        let Expr::StructuredRef(parsed) =
            parser::parse(&format!("={}", reference.format())).unwrap()
        else {
            panic!()
        };
        assert_eq!(parsed, reference);
    }
    for bad in [
        "Sales[]",
        "Sales[Amount",
        "Sales[[#Totals],[Amount]]",
        "Sales[[#Headers],[#Data]]",
        "Sales[[Qty],[Amount]]",
        "Sales[[Qty]:]",
        "Sales['q]",
    ] {
        assert!(parser::parse(&format!("={bad}")).is_err(), "{bad}");
    }
}

#[test]
fn structured_rows_and_aggregate_dependencies_recalculate_across_sheets() {
    let (mut wb, _) = fixture();
    assert_eq!(display(&wb, 0, 1, 2), "20");
    assert_eq!(display(&wb, 0, 2, 2), "60");
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Amount])");
    wb.set_cell_value_tracked(1, 0, 1, "=A1*2");
    assert_eq!(display(&wb, 1, 0, 0), "80");
    wb.set_cell_value_tracked(0, 2, 0, "4");
    assert_eq!(display(&wb, 0, 2, 2), "80");
    assert_eq!(display(&wb, 1, 0, 0), "100");
    assert_eq!(display(&wb, 1, 0, 1), "200");
    assert_eq!(eval(&mut wb, "=SUM(sales[[Qty]:[Price]])"), "36");
    assert!(eval(&mut wb, "=Sales[@Qty]").starts_with("#VALUE!"));
    assert!(eval(&mut wb, "=[@Qty]").starts_with("#VALUE!"));
    assert!(eval(&mut wb, "=SUM(Missing[Qty])").starts_with("#NAME?"));
}

#[test]
fn structured_resize_rebuilds_shape_and_value_dependencies_and_undo() {
    let (mut wb, id) = fixture();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Amount])");
    wb.set_cell_value_tracked(1, 0, 1, "=ROWS(Sales)");
    wb.set_cell_value_tracked(0, 3, 2, "7");
    let grow = wb.resize_table(id, range(3, 2)).unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "87");
    assert_eq!(display(&wb, 1, 0, 1), "3");
    wb.set_cell_value_tracked(0, 3, 2, "9");
    assert_eq!(display(&wb, 1, 0, 0), "89");
    wb.apply_table_commit(&grow, true).unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "80");
    assert_eq!(display(&wb, 1, 0, 1), "2");
    wb.set_cell_value_tracked(0, 3, 2, "100");
    assert_eq!(display(&wb, 1, 0, 0), "80");
}

#[test]
fn structured_renames_preserve_ids_formula_layout_literals_and_column_swaps() {
    let (mut wb, id) = fixture();
    let source = "= (SUM(Sales[Amount]) + 2) * 3 + IF(\"Sales[Amount]\"=\"Sales[Amount]\",1,0)";
    wb.set_cell_value_tracked(1, 0, 0, source);
    let rename = wb.rename_table(id, "Revenue").unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "247");
    assert_eq!(
        wb.sheet(1).unwrap().get_raw(0, 0),
        source.replacen("SUM(Sales[Amount])", "SUM(Revenue[Amount])", 1)
    );
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), source);
    let swap = wb
        .rename_table_columns(id, &["Price".into(), "Qty".into(), "Net [Amount]".into()])
        .unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "247");
    assert_eq!(display(&wb, 0, 1, 2), "20");
    assert!(wb
        .sheet(1)
        .unwrap()
        .get_raw(0, 0)
        .contains("Net '[Amount']"));
    wb.apply_table_commit(&swap, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), source);
}

#[test]
fn structured_deleted_column_is_permanently_ref_error_and_undo_restores_it() {
    let (mut wb, id) = fixture();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Amount])");
    let shrink = wb.resize_table(id, range(2, 1)).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=SUM(#REF!)");
    assert!(display(&wb, 1, 0, 0).starts_with("#REF!"));
    wb.apply_table_commit(&shrink, true).unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "80");
    wb.apply_table_commit(&shrink, false).unwrap();
    wb.resize_table(id, range(2, 2)).unwrap();
    assert!(display(&wb, 1, 0, 0).starts_with("#REF!"));
}

#[test]
fn structured_convert_to_range_and_stale_formula_undo_are_atomic() {
    let (mut wb, id) = fixture();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Amount])");
    let remove = wb.remove_table(id).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=SUM(Data!$C$2:$C$3)");
    assert_eq!(display(&wb, 1, 0, 0), "80");
    wb.apply_table_commit(&remove, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=SUM(Sales[Amount])");
    let rename = wb.rename_table(id, "Revenue").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=123");
    assert!(wb.apply_table_commit(&rename, true).is_err());
    assert!(wb.table_by_name("Revenue").is_some());
    assert_eq!(display(&wb, 1, 0, 0), "123");
}

#[test]
fn structured_empty_tables_are_zero_records_and_arrays_spill_outside_tables() {
    let (mut wb, id) = fixture();
    wb.resize_table(id, range(0, 2)).unwrap();
    assert_eq!(eval(&mut wb, "=SUM(Sales[Amount])"), "0");
    assert_eq!(eval(&mut wb, "=COUNT(Sales[Amount])"), "0");
    assert_eq!(eval(&mut wb, "=ROWS(Sales)"), "0");
    let empty_average = eval(&mut wb, "=AVERAGE()");
    assert_eq!(eval(&mut wb, "=AVERAGE(Sales[Amount])"), empty_average);
    assert_eq!(eval(&mut wb, r#"=COUNTIF(Sales[Amount],">0")"#), "0");
    assert_eq!(
        eval(&mut wb, r#"=SUMIF(Sales[Qty],">0",Sales[Amount])"#),
        "0"
    );
    wb.set_cell_value_tracked(1, 0, 0, "99");
    for function in ["SUMIF", "AVERAGEIF"] {
        assert!(
            eval(&mut wb, &format!("={function}(A1,\">0\",Sales[Amount])")).starts_with("#VALUE!")
        );
    }
    assert!(eval(&mut wb, "=Sales[Amount]").starts_with("#CALC!"));
    assert!(wb.remove_table(id).is_err());
    wb.resize_table(id, range(2, 2)).unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=Sales[Qty]");
    assert_eq!(display(&wb, 1, 0, 0), "2");
    assert_eq!(display(&wb, 1, 1, 0), "3");
    wb.set_cell_value_tracked(0, 1, 0, "4");
    assert_eq!(display(&wb, 1, 0, 0), "4");
}

#[test]
fn structured_range_consumers_share_cross_sheet_semantics() {
    let (mut wb, _) = fixture();
    for (formula, expected) in [
        ("=SUMPRODUCT(Sales[Qty],Sales[Price])", "80"),
        ("=COUNTIF(Sales[Qty],\">2\")", "1"),
        ("=SUMIF(Sales[Qty],\">2\",Sales[Amount])", "60"),
        ("=INDEX(Sales[Amount],2)", "60"),
        ("=COLUMNS(Sales)", "3"),
        ("=SUM(Sales)", "115"),
        ("=COUNTA(Sales[#Headers])", "3"),
    ] {
        assert_eq!(eval(&mut wb, formula), expected, "{formula}");
    }
    wb.set_cell_value_tracked(1, 0, 0, "=FILTER(Sales[Amount],Sales[Qty]>2)");
    assert_eq!(display(&wb, 1, 0, 0), "60");
}

#[test]
fn structured_copy_preserves_brackets_and_moves_only_a1_references() {
    let text = "=Sales[A1]+[@[B2]]+$C3+\"D4\"+Sales['#E5]";
    assert_eq!(
        parser::adjust_formula_refs(text, 2, 1),
        "=Sales[A1]+[@[B2]]+$C5+\"D4\"+Sales['#E5]"
    );
    let (mut wb, _) = fixture();
    let formula = parser::adjust_formula_refs("=[@Qty]*[@Price]", 1, 0);
    wb.set_cell_value_tracked(0, 2, 2, &formula);
    assert_eq!(display(&wb, 0, 2, 2), "60");
}

#[test]
fn structured_cycles_and_sheet_removal_are_guarded() {
    let (mut wb, _) = fixture();
    assert!(wb
        .check_formula_cycle(SheetId(1), 1, 2, "=SUM(Sales[Amount])")
        .is_err());
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Sales[Amount])");
    assert!(!wb.delete_sheet(0));
    assert!(wb.take_sheet(0).is_none());
    assert!(matches!(
        wb.sheet(0).unwrap().get_cell(1, 2).value,
        CellValue::Formula { .. }
    ));
}

#[test]
fn structured_move_and_formula_formatting_preserve_grouping() {
    use visigrid_engine::structural::Axis;
    let (mut wb, _) = fixture();
    wb.set_cell_value_tracked(1, 0, 0, "=(SUM(Sales[Amount])+2)*3");
    wb.structural_edit(0, Axis::Row, 0, 2, false).unwrap();
    assert_eq!(display(&wb, 1, 0, 0), "246");
    assert_eq!(display(&wb, 0, 3, 2), "20");
    wb.set_cell_value_tracked(0, 3, 0, "4");
    assert_eq!(display(&wb, 1, 0, 0), "306");
    for input in [
        "=(1+2)*3", "=1-(2-3)", "=1/(2/3)", "=(2^3)^2", "=2^(3^2)", "=1+(2+3)",
    ] {
        let ast = parser::parse(input).unwrap();
        let formatted = parser::format_parsed_expr(&ast);
        assert_eq!(
            format!("{:?}", ast),
            format!("{:?}", parser::parse(&formatted).unwrap()),
            "{input}"
        );
    }
    assert_eq!(eval(&mut wb, "=IF(FALSE,Missing[Amount],7)"), "7");
}

#[test]
fn structured_replay_rejects_new_dependent_formulas() {
    let (mut wb, id) = fixture();
    let rename = wb.rename_table(id, "Revenue").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=SUM(Revenue[Amount])");
    assert!(wb.apply_table_commit(&rename, true).is_err());
    assert!(wb.table_by_name("Revenue").is_some());
    assert_eq!(display(&wb, 1, 0, 0), "80");
}

#[test]
fn structured_creation_binds_existing_formulas_and_undo_restores_unbound_state() {
    let mut wb = Workbook::new();
    let sid = wb.active_sheet().id;
    wb.set_cell_value_tracked(0, 0, 0, "Amount");
    wb.set_cell_value_tracked(0, 1, 0, "12");
    wb.set_cell_value_tracked(0, 0, 4, "=SUM(Sales[Amount])");
    assert!(display(&wb, 0, 0, 4).starts_with("#NAME?"));
    let create = wb.create_table(sid, range(1, 0), "Sales").unwrap();
    assert_eq!(display(&wb, 0, 0, 4), "12");
    wb.apply_table_commit(&create, true).unwrap();
    assert!(display(&wb, 0, 0, 4).starts_with("#NAME?"));
    wb.apply_table_commit(&create, false).unwrap();
    assert_eq!(display(&wb, 0, 0, 4), "12");
}

#[test]
fn structured_default_table_name_is_not_an_off_grid_cell_reference() {
    let (mut wb, id) = fixture();
    wb.rename_table(id, "Table1").unwrap();
    assert_eq!(eval(&mut wb, "=SUM(Table1)"), "115");
    assert_eq!(eval(&mut wb, "=SUM(Table1[Amount])"), "80");
    assert_eq!(
        parser::adjust_formula_refs("=SUM(Table1)+A1", 1, 1),
        "=SUM(Table1)+B2"
    );
    assert_eq!(parser::adjust_formula_refs("=XFD1", 0, 1), "=#REF!");
}

#[test]
fn structured_rename_preserves_scientific_number_literals() {
    let (mut wb, id) = fixture();
    wb.rename_table(id, "E").unwrap();
    let source = "=1E+2+1e-2+SUM(E[Amount])";
    wb.set_cell_value_tracked(1, 0, 0, source);
    let before = display(&wb, 1, 0, 0);
    let rename = wb.rename_table(id, "Revenue").unwrap();
    assert_eq!(
        wb.sheet(1).unwrap().get_raw(0, 0),
        "=1E+2+1e-2+SUM(Revenue[Amount])"
    );
    assert_eq!(display(&wb, 1, 0, 0), before);
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), source);
}

#[test]
fn structured_references_work_with_standalone_sheet_evaluation() {
    let (mut wb, _) = fixture();
    let sheet = wb.sheet_mut(0).unwrap();
    sheet.set_value(1, 2, "=[@Qty]*[@Price]+1");
    assert_eq!(sheet.get_display(1, 2), "21");
    sheet.set_value(5, 0, "=SUM(Sales[Amount])");
    assert_eq!(sheet.get_display(5, 0), "81");
}
