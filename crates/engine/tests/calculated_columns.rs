use visigrid_engine::{
    sheet::{Sheet, SheetId},
    structural::Axis,
    table::{TableId, TableRange},
    workbook::Workbook,
};
fn range(end: usize) -> TableRange {
    TableRange {
        start_row: 0,
        start_col: 0,
        end_row: end,
        end_col: 2,
    }
}
fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 100, 20, "Data")], 0);
    for (c, s) in ["Qty", "Price", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, c, s);
    }
    for (r, q) in [(1, "2"), (2, "3"), (3, "4")] {
        wb.set_cell_value_tracked(0, r, 0, q);
        wb.set_cell_value_tracked(0, r, 1, "10");
    }
    let id = wb
        .create_table(SheetId(1), range(3), "Sales")
        .unwrap()
        .table_id();
    let other = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=SUM(Sales[Amount])");
    (wb, id)
}
#[test]
fn single_edit_fills_once_with_relative_absolute_and_structured_refs() {
    for source in ["=[@Qty]*[@Price]", "=A3*$B$2", "=$A3*B$2"] {
        let (mut wb, id) = book();
        let rev = wb.revision();
        let commit = wb
            .try_calculated_column(SheetId(1), 2, 2, source)
            .unwrap()
            .unwrap();
        assert_eq!(wb.revision(), rev + 1);
        assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "90");
        assert!(!wb.sheet(0).unwrap().is_calculated_exception(3, 2));
        wb.apply_table_commit(&commit, true).unwrap();
        assert!(wb.table(id).unwrap().1.columns[2].formula.is_none());
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "");
        wb.apply_table_commit(&commit, false).unwrap();
        wb.set_cell_value_tracked(0, 3, 0, "5");
        assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "100");
    }
}
#[test]
fn populated_column_and_bulk_writes_do_not_infer_rules() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 1, 2, "99");
    assert!(wb
        .try_calculated_column(SheetId(1), 2, 2, "=[@Qty]*10")
        .unwrap()
        .is_none());
    wb.set_cell_value_tracked(0, 2, 2, "=[@Qty]*10");
    wb.append_table_rows(id, 1, &[(4, 2, "=[@Qty]*10".into())])
        .unwrap();
    assert!(wb.table(id).unwrap().1.columns[2].formula.is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "99");
}
#[test]
fn exceptions_survive_template_edits_and_restore_and_replace_are_undoable() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 1, "=[@Qty]*10", true)
        .unwrap();
    wb.set_cell_value_tracked(0, 2, 2, "777");
    wb.clear_cell_tracked(0, 3, 2);
    assert!(wb.sheet(0).unwrap().is_calculated_exception(2, 2));
    assert!(wb.sheet(0).unwrap().is_calculated_exception(3, 2));
    let update = wb
        .set_calculated_column(id, 2, 1, "=[@Qty]*20", false)
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "40");
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "777");
    assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "");
    let restore = wb.restore_calculated_cell(id, 3, 2).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 2), "80");
    wb.apply_table_commit(&restore, true).unwrap();
    assert!(wb.sheet(0).unwrap().is_calculated_exception(3, 2));
    wb.apply_table_commit(&update, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "20");
    let replace = wb.set_calculated_column(id, 2, 1, "=A2*30", true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(2, 2), "90");
    wb.apply_table_commit(&replace, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "777");
}
#[test]
fn append_fills_omissions_but_respects_explicit_paste_blanks_and_values() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 1, "=A2*B2", true).unwrap();
    let append = wb
        .append_table_rows(
            id,
            3,
            &[
                (4, 0, "5".into()),
                (4, 1, "10".into()),
                (5, 2, "".into()),
                (6, 2, "999".into()),
            ],
        )
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "50");
    assert!(!wb.sheet(0).unwrap().is_calculated_exception(4, 2));
    assert!(wb.sheet(0).unwrap().is_calculated_exception(5, 2));
    assert!(wb.sheet(0).unwrap().is_calculated_exception(6, 2));
    wb.apply_table_commit(&append, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 3);
    wb.apply_table_commit(&append, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "=A5*B5");
}
#[test]
fn tab_edit_and_appended_typing_can_establish_rule_in_same_commit() {
    for row in [3, 4] {
        let (mut wb, id) = book();
        let rev = wb.revision();
        let commit = wb
            .append_table_rows_with_edit(id, 1, row, 2, "=[@Qty]*[@Price]")
            .unwrap();
        assert_eq!(wb.revision(), rev + 1);
        assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "20");
        assert_eq!(wb.sheet(0).unwrap().get_display(3, 2), "40");
        assert!(wb.table(id).unwrap().1.columns[2].formula.is_some());
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.range.end_row, 3);
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "");
    }
}
#[test]
fn rules_follow_schema_renames_even_when_header_only() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 1, "=[@Qty]*[@Price]", true)
        .unwrap();
    let rename = wb
        .rename_table_columns(id, &["Quantity".into(), "Price".into(), "Amount".into()])
        .unwrap();
    wb.append_table_rows(id, 1, &[(4, 0, "5".into()), (4, 1, "10".into())])
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "50");
    assert!(!wb.sheet(0).unwrap().is_calculated_exception(1, 2));
    wb.resize_table(id, range(0)).unwrap();
    wb.rename_table_columns(id, &["Units".into(), "Price".into(), "Amount".into()])
        .unwrap();
    assert!(wb.table(id).unwrap().1.columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("Units"));
    assert!(wb.apply_table_commit(&rename, true).is_err());
}
#[test]
fn external_rules_follow_rename_and_convert_with_guarded_undo() {
    let (mut wb, id) = book();
    let other = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 10,
                start_col: 5,
                end_row: 10,
                end_col: 5,
            },
            "SummaryRule",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(other, 5, 11, "=SUM(Sales[Qty])", true)
        .unwrap();
    let rename = wb.rename_table(id, "Orders").unwrap();
    assert!(wb.table(other).unwrap().1.columns[0]
        .formula
        .as_ref()
        .unwrap()
        .contains("Orders"));
    wb.apply_table_commit(&rename, true).unwrap();
    let convert = wb.remove_table(id).unwrap();
    assert!(!wb.table(other).unwrap().1.columns[0]
        .formula
        .as_ref()
        .unwrap()
        .contains("Sales"));
    wb.apply_table_commit(&convert, true).unwrap();
    wb.append_table_rows(other, 1, &[]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(11, 5), "9");
}
#[test]
fn row_history_preserves_rules_and_empty_exceptions() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 1, "=A2*B2", true).unwrap();
    wb.clear_cell_tracked(0, 2, 2);
    let insert = wb
        .prepare_table_row_history(0, 2, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_row_history(&insert, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "=A3*B3");
    assert!(wb.sheet(0).unwrap().is_calculated_exception(3, 2));
    wb.apply_table_row_history(&insert, true).unwrap();
    assert!(wb.sheet(0).unwrap().is_calculated_exception(2, 2));
    let delete = wb
        .prepare_table_row_history(0, 1, 3, true)
        .unwrap()
        .unwrap();
    wb.apply_table_row_history(&delete, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 0);
    wb.apply_table_row_history(&delete, true).unwrap();
    // Ordinary deleted-cell payload restoration happens after bounds/rules.
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "");
    wb.apply_table_row_history(&delete, false).unwrap();
    wb.append_table_rows(id, 1, &[(1, 0, "5".into()), (1, 1, "10".into())])
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "50");
}
#[test]
fn structural_move_keeps_rule_and_followers_equivalent() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 1, "=A2*$B$2", true)
        .unwrap();
    wb.structural_edit(0, Axis::Row, 0, 2, false).unwrap();
    assert!(!wb.sheet(0).unwrap().is_calculated_exception(3, 2));
    wb.structural_edit(0, Axis::Col, 0, 1, false).unwrap();
    assert!(!wb.sheet(0).unwrap().is_calculated_exception(3, 3));
    wb.append_table_rows(id, 1, &[(6, 1, "5".into())]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(6, 3), "50");
}
#[test]
fn arrays_invalid_formulas_and_stale_replays_refuse_without_partial_writes() {
    let (mut wb, id) = book();
    let rev = wb.revision();
    for formula in ["=SEQUENCE(3)", "=SUM("] {
        assert!(wb.set_calculated_column(id, 2, 1, formula, true).is_err());
        assert_eq!(wb.revision(), rev);
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "");
    }
    assert!(wb.set_calculated_column(id, 2, usize::MAX, "=1", true).is_err());
    assert_eq!(wb.revision(), rev);
    let commit = wb
        .set_calculated_column(id, 2, 1, "=[@Qty]*10", true)
        .unwrap();
    wb.set_cell_value_tracked(0, 2, 2, "77");
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "77");
}

#[test]
fn authored_origin_retains_relative_references_above_the_first_record() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 2, 3, "=A1", true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "=A1");
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "=#REF!");
    wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "=A2");
}

#[test]
fn row_edit_on_plain_sheet_keeps_external_rule_history() {
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(1, 3, 1, "8");
    wb.set_calculated_column(id, 2, 1, "=Summary!$B$4", true)
        .unwrap();
    let history = wb
        .prepare_table_row_history(1, 1, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert!(wb.table(id).unwrap().1.columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("$B$5"));
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(
        wb.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=Summary!$B$4")
    );
    wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "8");
}

#[test]
fn totals_rule_edits_preserve_overrides_and_footer_and_invalidate_dependents() {
    use visigrid_engine::{filter::{ColumnFilter, NormalizedFilterKey}, table_view::{TableFilter, TableViewSpec}};
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        let totals = wb.table(id).unwrap().1.totals.clone();
        let mut format = wb.sheet(0).unwrap().get_format(4, 2);
        format.bold = true;
        wb.sheet_mut(0).unwrap().set_format(4, 2, format.clone());
        let first = wb.set_calculated_column(id, 2, 1, "=[@Qty]*10", true).unwrap();
        assert!(first.is_calculated_change());
        wb.set_cell_value_tracked(0, 2, 2, "777");
        wb.clear_cell_tracked(0, 3, 2);
        let mut spec = TableViewSpec::new(id);
        spec.filters.push(TableFilter { column: wb.table(id).unwrap().1.columns[0].id,
            criteria: ColumnFilter { selected: Some([NormalizedFilterKey::Number(2.0.into()), NormalizedFilterKey::Number(4.0.into())].into()), text_filter: None }});
        wb.set_table_view_spec(SheetId(1), Some(spec.clone())).unwrap();
        let generation = wb.sheet(1).unwrap().edit_generation();
        let update = wb.set_calculated_column(id, 2, 1, "=[@Qty]*20", false).unwrap();
        assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "40");
        assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "777");
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "");
        assert!(wb.sheet(1).unwrap().edit_generation() > generation);
        assert_eq!(wb.table(id).unwrap().1.totals, totals);
        assert_eq!(wb.sheet(0).unwrap().get_format(4, 2), format);
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), if visible { "40" } else { "" });
        let replace = wb.set_calculated_column(id, 2, 1, "=[@Qty]*20", true).unwrap();
        assert!(replace.is_calculated_change(), "same-rule replacement still owns body writes");
        assert_eq!(wb.sheet(0).unwrap().get_display(2, 2), "60");
        assert_eq!(wb.sheet(0).unwrap().get_display(3, 2), "80");
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), if visible { "120" } else { "" });
        assert_eq!(wb.sheet(0).unwrap().table_view_spec(), Some(&spec));
        wb.apply_table_commit(&replace, true).unwrap();
        wb.apply_table_commit(&update, true).unwrap();
        assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "20");
        assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "777");
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "");
        wb.apply_table_commit(&update, false).unwrap();
        wb.apply_table_commit(&replace, false).unwrap();
        if !visible { wb.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "120");
        assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "=SUBTOTAL(109,[Amount])");
    }
}

#[test]
fn totals_inference_and_restore_never_fill_footer_and_reject_stale_history() {
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let footer = wb.sheet(0).unwrap().get_raw(4, 2);
    assert!(wb.try_calculated_column(SheetId(1), 4, 2, "=999").unwrap().is_none());
    assert!(wb.table(id).unwrap().1.columns[2].formula.is_none());
    let first = wb.try_calculated_column(SheetId(1), 2, 2, "=A3*10").unwrap().unwrap();
    assert!(first.is_calculated_change());
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "=A2*10");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), footer);
    assert!(wb.restore_calculated_cell(id, 4, 2).is_err());
    wb.set_cell_value_tracked(0, 2, 2, "777");
    let restore = wb.restore_calculated_cell(id, 2, 2).unwrap();
    assert!(restore.is_calculated_change());
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "90");
    wb.apply_table_commit(&restore, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "777");
    wb.apply_table_commit(&restore, false).unwrap();
    wb.set_cell_value_tracked(0, 2, 2, "888");
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&restore, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "888");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), footer);
}

#[test]
fn totals_rule_rejects_array_formulas_atomically_and_supports_empty_body_template() {
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let revision = wb.revision();
    assert!(wb.set_calculated_column(id, 2, 1, "=SEQUENCE(2)", true).is_err());
    assert_eq!(wb.revision(), revision);
    assert!(wb.table(id).unwrap().1.columns[2].formula.is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 2), "");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 2), "=SUBTOTAL(109,[Amount])");
    let mut empty = Workbook::new();
    empty.set_cell_value_tracked(0, 0, 0, "Value");
    let id = empty.create_table(empty.active_sheet_id(), TableRange { start_row: 0, end_row: 0, start_col: 0, end_col: 0 }, "Empty").unwrap().table_id();
    empty.set_table_totals_visible(id, true, Default::default()).unwrap();
    empty.set_calculated_column(id, 0, 1, "=7", true).unwrap();
    assert_eq!(empty.active_sheet().get_raw(1, 0), "=SUBTOTAL(109,[Value])");
    empty.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(empty.active_sheet().get_display(1, 0), "7");
    assert_eq!(empty.active_sheet().get_display(2, 0), "7");
}
