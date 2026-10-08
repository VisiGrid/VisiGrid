use visigrid_engine::{
    sheet::{Sheet, SheetId},
    structural::Axis,
    table::{TableId, TableRange},
    workbook::Workbook,
};

fn book(source: &str) -> (Workbook, TableId) {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 100, 20, "Data")], 0);
    for (c, name) in ["Qty", "Price", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 0, c, name);
    }
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 2,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    for r in 1..=3 {
        wb.set_cell_value_tracked(0, r, 0, "2");
        wb.set_cell_value_tracked(0, r, 1, "10");
    }
    wb.set_calculated_column(id, 2, 1, source, true).unwrap();
    wb.set_cell_value_tracked(0, 2, 2, "999");
    wb.clear_cell_tracked(0, 3, 2);
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[Amount])");
    wb.set_cell_value_tracked(summary, 1, 0, "=Data!$C$2");
    (wb, id)
}

#[test]
fn interior_insertion_preserves_ids_formulas_and_overrides_across_replay() {
    for source in ["=[@Qty]*[@Price]", "=A2*$B$2"] {
        let (mut wb, id) = book(source);
        let original = wb.table(id).unwrap().1.clone();
        let history = wb
            .prepare_table_column_history(0, 1, 2, false)
            .unwrap()
            .unwrap();
        let revision = wb.revision();
        wb.apply_table_column_history(&history, false).unwrap();
        assert_eq!(wb.revision(), revision + 1);
        let columns = wb.table(id).unwrap().1.columns.clone();
        assert_eq!(columns[0].id, original.columns[0].id);
        assert_eq!(columns[3].id, original.columns[1].id);
        assert_eq!(columns[4].id, original.columns[2].id);
        assert_eq!(wb.sheet(0).unwrap().get_raw(0, 1), "Column1");
        assert_eq!(wb.sheet(0).unwrap().get_raw(0, 2), "Column2");
        assert_eq!(wb.sheet(0).unwrap().get_display(1, 4), "20");
        assert!(!wb.sheet(0).unwrap().is_calculated_exception(1, 4));
        assert!(wb.sheet(0).unwrap().is_calculated_exception(2, 4));
        assert!(wb.sheet(0).unwrap().is_calculated_exception(3, 4));
        assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), "=Data!$E$2");
        wb.apply_table_column_history(&history, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns, original.columns);
        assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), "=Data!$C$2");
        wb.apply_table_column_history(&history, false).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns, columns);
        wb.append_table_rows(id, 1, &[(4, 0, "3".into()), (4, 3, "10".into())])
            .unwrap();
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 4), "30");
    }
}

#[test]
fn deletion_restores_first_middle_and_calculated_columns_with_original_identity() {
    for at in 0..3 {
        let (mut wb, id) = book("=[@Qty]*[@Price]");
        let original = wb.table(id).unwrap().1.clone();
        let cells = wb.sheet(0).unwrap().occupied_cells_in_cols(at, 1);
        let history = wb
            .prepare_table_column_history(0, at, 1, true)
            .unwrap()
            .unwrap();
        let rewrites = wb.apply_table_column_history(&history, false).unwrap();
        let after = wb.table(id).unwrap().1.columns.clone();
        assert_eq!(after.len(), 2);
        assert!(!after.iter().any(|c| c.id == original.columns[at].id));
        if at == 2 {
            assert!(wb.sheet(1).unwrap().get_raw(0, 0).contains("#REF!"));
        } else {
            assert!(after[1].formula.as_ref().unwrap().contains("#REF!"));
        }
        wb.apply_table_column_history(&history, true).unwrap();
        // Ordinary desktop column history restores the sparse deleted payload
        // and destructive reference rewrites after the inverse structural move.
        for (r, c, value, _) in &cells {
            wb.set_cell_value_tracked(0, *r, *c, value);
        }
        for (si, r, c, old, _) in &rewrites {
            wb.set_cell_value_tracked(*si, *r, *c, old);
        }
        assert_eq!(wb.table(id).unwrap().1.columns, original.columns);
        assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "20");
        assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "999");
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "");
        assert_eq!(wb.sheet(1).unwrap().get_display(0, 0), "1019");
        wb.apply_table_column_history(&history, false).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns, after);
    }
}

#[test]
fn deleted_references_do_not_rebind_when_column_name_is_reused() {
    let (mut wb, id) = book("=[@Qty]*[@Price]");
    wb.structural_edit(0, Axis::Col, 1, 1, true).unwrap();
    wb.structural_edit(0, Axis::Col, 1, 1, false).unwrap();
    wb.rename_table_columns(id, &["Qty".into(), "Price".into(), "Amount".into()])
        .unwrap();
    wb.set_cell_value_tracked(0, 1, 1, "10");
    assert!(wb.sheet(0).unwrap().get_raw(1, 2).contains("#REF!"));
    assert!(wb.table(id).unwrap().1.columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("#REF!"));
}

#[test]
fn undo_keeps_allocator_high_water_and_unique_generated_headers() {
    let (mut wb, id) = book("=A2*B2");
    wb.rename_table_columns(id, &["column1".into(), "Price".into(), "Amount".into()])
        .unwrap();
    let history = wb
        .prepare_table_column_history(0, 1, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&history, false).unwrap();
    let allocated = wb.table(id).unwrap().1.columns[1].id;
    assert_eq!(wb.table(id).unwrap().1.columns[1].name, "Column2");
    wb.apply_table_column_history(&history, true).unwrap();
    wb.structural_edit(0, Axis::Col, 1, 1, false).unwrap();
    assert!(wb.table(id).unwrap().1.columns[1].id.0 > allocated.0);
}

#[test]
fn rules_on_other_sheets_follow_plain_sheet_column_history_even_without_body_cells() {
    let (mut wb, id) = book("=Summary!B2");
    wb.resize_table(
        id,
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 2,
        },
    )
    .unwrap();
    let history = wb
        .prepare_table_column_history(1, 1, 1, true)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&history, false).unwrap();
    assert!(wb.table(id).unwrap().1.columns[2]
        .formula
        .as_ref()
        .unwrap()
        .contains("#REF!"));
    wb.apply_table_column_history(&history, true).unwrap();
    assert_eq!(
        wb.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=Summary!B2")
    );
    wb.apply_table_column_history(&history, false).unwrap();
}

#[test]
fn multiple_tables_shift_together_and_refusals_leave_everything_unchanged() {
    let (mut wb, id) = book("=A2*B2");
    let second = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 5,
                start_col: 0,
                end_row: 5,
                end_col: 2,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    let history = wb
        .prepare_table_column_history(0, 1, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&history, false).unwrap();
    assert_eq!(wb.table(second).unwrap().1.range.end_col, 3);
    wb.apply_table_column_history(&history, true).unwrap();
    let revision = wb.revision();
    assert!(wb.prepare_table_column_history(0, 0, 3, true).is_err());
    assert!(wb
        .prepare_table_column_history(0, usize::MAX, 1, false)
        .is_err());
    wb.set_cell_value_tracked(0, 10, 19, "edge");
    assert!(wb.prepare_table_column_history(0, 1, 1, false).is_err());
    assert_eq!(wb.revision(), revision + 1);
    wb.rename_table(id, "Orders").unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_column_history(&history, false).is_err());
    assert_eq!(wb.revision(), revision);
}

#[test]
fn external_structured_rules_and_header_only_local_rules_restore_after_deletion() {
    let (mut wb, id) = book("=[@Qty]*[@Price]");
    let external_sheet = wb.sheet(1).unwrap().id;
    let external = wb
        .create_table(
            external_sheet,
            TableRange {
                start_row: 5,
                start_col: 0,
                end_row: 5,
                end_col: 0,
            },
            "Totals",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(external, 0, 6, "=SUM(Sales[Price])", true)
        .unwrap();
    wb.resize_table(
        id,
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 2,
        },
    )
    .unwrap();
    let history = wb
        .prepare_table_column_history(0, 1, 1, true)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&history, false).unwrap();
    assert!(wb.table(id).unwrap().1.columns[1]
        .formula
        .as_ref()
        .unwrap()
        .contains("#REF!"));
    assert!(wb.table(external).unwrap().1.columns[0]
        .formula
        .as_ref()
        .unwrap()
        .contains("#REF!"));
    wb.apply_table_column_history(&history, true).unwrap();
    assert_eq!(
        wb.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=[@Qty]*[@Price]")
    );
    assert_eq!(
        wb.table(external).unwrap().1.columns[0].formula.as_deref(),
        Some("=SUM(Sales[Price])")
    );
    wb.set_calculated_column(external, 0, 6, "=1", false)
        .unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_column_history(&history, false).is_err());
    assert_eq!(wb.revision(), revision);
}

#[test]
fn boundary_insertions_shift_but_do_not_expand_and_clipped_deletions_undo_exactly() {
    let (mut wb, id) = book("=A2*B2");
    let columns = wb.table(id).unwrap().1.columns.clone();
    let shift = wb
        .prepare_table_column_history(0, 0, 3, false)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&shift, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.start_col, 3);
    assert_eq!(wb.table(id).unwrap().1.columns.len(), 3);
    let after = wb
        .prepare_table_column_history(0, 6, 1, false)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&after, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_col, 5);
    wb.apply_table_column_history(&after, true).unwrap();
    let clipped = wb
        .prepare_table_column_history(0, 1, 3, true)
        .unwrap()
        .unwrap();
    wb.apply_table_column_history(&clipped, false).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.start_col, 1);
    assert_eq!(wb.table(id).unwrap().1.columns[0].id, columns[1].id);
    wb.apply_table_column_history(&clipped, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.start_col, 3);
    assert_eq!(wb.table(id).unwrap().1.columns[0].id, columns[0].id);
}


#[test]
fn totals_columns_move_settings_cells_and_references_with_atomic_history() {
    use visigrid_engine::table::TableTotal;
    for visible in [true, false] {
        let (mut wb, id) = book("=[@Qty]*[@Price]");
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 1, TableTotal {
            function: Some("custom".into()), formula: Some("=SUM([Qty])+A2".into()), label: None,
        }).unwrap();
        if visible { wb.sheet_mut(0).unwrap().set_comment(4, 1, Some(visigrid_engine::cell::CellComment { text: "Custom footer".into(), author: "QA".into() })); }
        wb.sheet_mut(0).unwrap().toggle_bold(4, 1);
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        let original = wb.table(id).unwrap().1.clone();
        let generation = wb.sheet(1).unwrap().edit_generation();
        let history = wb.prepare_table_column_history(0, 1, 2, false).unwrap().unwrap();
        wb.apply_table_column_history(&history, false).unwrap();
        let table = wb.table(id).unwrap().1;
        assert_eq!(table.columns[3].id, original.columns[1].id);
        assert_eq!(table.totals.as_ref().unwrap().columns.len(), 5);
        assert_eq!(table.totals.as_ref().unwrap().columns[1], Default::default());
        assert_eq!(table.totals.as_ref().unwrap().columns[3].formula.as_deref(), Some("=SUM([Qty])+A2"));
        assert!(wb.sheet(1).unwrap().edit_generation() > generation);
        if visible {
            assert_eq!(wb.sheet(0).unwrap().get_display(4, 3), "8");
            assert_eq!(wb.sheet(0).unwrap().get_raw(4, 4), "=SUBTOTAL(109, [Amount])");
            assert!(wb.sheet(0).unwrap().get_format(4, 3).bold);
            assert_eq!(wb.sheet(0).unwrap().get_cell(4, 3).comment().map(|c| c.text.as_str()), Some("Custom footer"));
        }
        wb.apply_table_column_history(&history, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns, original.columns);
        assert_eq!(wb.table(id).unwrap().1.totals, original.totals);
        wb.apply_table_column_history(&history, false).unwrap();
        if !visible { wb.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 3), "8");
    }
}

#[test]
fn deleted_field_references_in_visible_and_dormant_totals_restore_on_undo() {
    use visigrid_engine::table::TableTotal;
    for visible in [true, false] {
        let (mut wb, id) = book("=[@Qty]*[@Price]");
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 2, TableTotal {
            function: Some("custom".into()), formula: Some("=SUM([Qty])+SUM([Price])+B2".into()), label: None,
        }).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        let original = wb.table(id).unwrap().1.clone();
        let history = wb.prepare_table_column_history(0, 1, 1, true).unwrap().unwrap();
        wb.apply_table_column_history(&history, false).unwrap();
        let table = wb.table(id).unwrap().1;
        assert_eq!(table.columns.len(), 2);
        assert_eq!(table.totals.as_ref().unwrap().columns.len(), 2);
        let source = table.totals.as_ref().unwrap().columns[1].formula.as_ref().unwrap();
        assert!(source.contains("SUM(#REF!)"), "{source}");
        assert!(source.ends_with("+#REF!"), "{source}");
        if visible { assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), *source); }
        wb.apply_table_column_history(&history, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.totals, original.totals);
        if visible { assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "46"); }
        wb.apply_table_column_history(&history, false).unwrap();
        let revision = wb.revision();
        wb.set_cell_value_tracked(1, 9, 0, "Later edit");
        assert!(wb.apply_table_column_history(&history, true).is_err());
        assert_eq!(wb.revision(), revision + 1);
        assert_eq!(wb.sheet(1).unwrap().get_raw(9, 0), "Later edit");
    }
}

#[test]
fn column_moves_rewrite_other_sheet_totals_and_keep_dormant_local_context() {
    use visigrid_engine::table::TableTotal;
    let (mut wb, id) = book("=[@Qty]*[@Price]");
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_cell_value_tracked(1, 3, 3, "Remote");
    wb.set_cell_value_tracked(1, 4, 3, "7");
    let other_sheet = wb.sheet(1).unwrap().id;
    let other = wb.create_table(other_sheet, TableRange { start_row: 3, end_row: 4, start_col: 3, end_col: 3 }, "Other").unwrap().table_id();
    wb.set_table_totals_visible(other, true, Default::default()).unwrap();
    wb.set_table_total(other, 3, TableTotal {
        function: Some("custom".into()), formula: Some("=SUM(Sales[Qty])+Data!A2+SUM([Remote])".into()), label: None,
    }).unwrap();
    wb.set_table_totals_visible(other, false, Default::default()).unwrap();
    let history = wb.prepare_table_column_history(0, 0, 2, false).unwrap().unwrap();
    wb.apply_table_column_history(&history, false).unwrap();
    assert_eq!(wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_deref(),
        Some("=SUM(Sales[Qty])+Data!C2+SUM([Remote])"));
    wb.apply_table_column_history(&history, true).unwrap();
    wb.set_table_totals_visible(other, true, Default::default()).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_display(5, 3), "15");
    let (mut deleted, _) = wb.prepare_guarded_structure(0, vec![visigrid_engine::workbook::StructureStep {
        axis: Axis::Col, at: 0, count: 1, delete: true,
    }]).unwrap();
    assert!(deleted.sheet(1).unwrap().get_raw(5, 3).contains("SUM(#REF!)"));
    assert!(deleted.table(other).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_ref().unwrap().contains("SUM(#REF!)"));
    let revision = deleted.revision();
    assert!(deleted.structural_edit(0, Axis::Col, 0, 2, true).is_err());
    assert_eq!(deleted.revision(), revision);
}
