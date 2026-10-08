use visigrid_engine::{
    filter::{ColumnFilter, NormalizedFilterKey},
    table::{TableId, TableRange, TableTotal},
    table_view::{TableFilter, TableViewSpec},
    workbook::Workbook,
};

fn custom(source: &str) -> TableTotal {
    TableTotal {
        function: Some("custom".into()),
        formula: Some(source.into()),
        label: None,
    }
}
fn sales() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 0,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let (wb, _) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [2].into())
        .unwrap();
    (wb, id)
}

#[test]
fn foreign_visible_and_dormant_totals_convert_with_local_context_and_active_criteria() {
    for cross_sheet in [false, true] {
        for visible in [false, true] {
            let (mut wb, id) = sales();
            let sheet = if cross_sheet {
                wb.add_sheet_named("Summary").unwrap()
            } else {
                0
            };
            let start = if cross_sheet { 0 } else { 8 };
            for (r, value) in ["Value", "5", "7"].iter().enumerate() {
                wb.set_cell_value_tracked(sheet, start + r, 0, value);
            }
            let other = wb
                .create_table(
                    wb.sheet(sheet).unwrap().id,
                    TableRange {
                        start_row: start,
                        end_row: start + 2,
                        start_col: 0,
                        end_col: 0,
                    },
                    "Other",
                )
                .unwrap()
                .table_id();
            let flags = wb.sheet(sheet).unwrap().manual_hidden_rows();
            wb.set_table_totals_visible(other, true, flags.clone())
                .unwrap();
            let source = "=SUM(Sales[[#Totals],[Amount]])+SUBTOTAL(109,[Value])+IF(\"Sales[Amount]\"=\"x\",100,1)";
            wb.set_table_total(other, 0, custom(source)).unwrap();
            let mut spec = TableViewSpec::new(other);
            spec.filters.push(TableFilter {
                column: wb.table(other).unwrap().1.columns[0].id,
                criteria: ColumnFilter {
                    selected: Some([NormalizedFilterKey::Number(5.0.into())].into()),
                    text_filter: None,
                },
            });
            wb.set_table_view_spec(wb.sheet(sheet).unwrap().id, Some(spec.clone()))
                .unwrap();
            assert_eq!(wb.sheet(sheet).unwrap().get_display(start + 3, 0), "46");
            if !visible {
                wb.set_table_totals_visible(other, false, flags.clone())
                    .unwrap();
            }
            let before = serde_json::to_value(wb.saved_tables()).unwrap();
            let generation = wb.sheet(sheet).unwrap().edit_generation();
            let commit = wb.remove_table(id).unwrap();
            let formula = wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0]
                .formula
                .as_ref()
                .unwrap();
            assert!(formula.contains("$A$5"));
            assert!(formula.contains("SUBTOTAL(109,[Value])"));
            assert!(formula.contains("\"Sales[Amount]\""));
            assert_eq!(wb.sheet(sheet).unwrap().table_view_spec(), Some(&spec));
            assert!(wb.sheet(sheet).unwrap().edit_generation() > generation);
            if visible {
                assert_eq!(wb.sheet(sheet).unwrap().get_display(start + 3, 0), "46");
            }
            wb.apply_table_commit(&commit, true).unwrap();
            assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
            wb.apply_table_commit(&commit, false).unwrap();
            if !visible {
                wb.set_table_totals_visible(other, true, flags).unwrap();
            }
            assert_eq!(wb.sheet(sheet).unwrap().get_display(start + 3, 0), "46");
            wb.set_cell_value_tracked(0, 1, 0, "100");
            assert_eq!(wb.sheet(sheet).unwrap().get_display(start + 3, 0), "136");
        }
    }
}

#[test]
fn dormant_footer_overlapping_another_table_keeps_its_own_local_references() {
    let mut wb = Workbook::new();
    for (r, value) in ["Amount", "1", "2"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, value);
    }
    let other = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 2,
                start_col: 0,
                end_col: 0,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(other, true, Default::default())
        .unwrap();
    wb.set_table_total(
        other,
        0,
        custom("=SUM([Amount])+IFERROR(SUM(Sales[Amount]),0)"),
    )
    .unwrap();
    wb.set_table_totals_visible(other, false, Default::default())
        .unwrap();
    for (r, value) in ["Amount", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r + 3, 0, value);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 3,
                end_row: 5,
                start_col: 0,
                end_col: 0,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let rename = wb.rename_table_columns(id, &["Revenue".into()]).unwrap();
    assert_eq!(
        wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0]
            .formula
            .as_deref(),
        Some("=SUM([Amount])+IFERROR(SUM(Sales[Revenue]),0)")
    );
    wb.apply_table_commit(&rename, true).unwrap();
    let before = serde_json::to_value(wb.saved_tables()).unwrap();
    let commit = wb.remove_table(id).unwrap();
    let source = wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0]
        .formula
        .as_ref()
        .unwrap();
    assert!(source.contains("SUM([Amount])"));
    assert!(source.contains("$A$5:$A$6"));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
    wb.apply_table_commit(&commit, false).unwrap();
    wb.clear_cell_tracked(0, 3, 0);
    wb.set_table_totals_visible(other, true, Default::default())
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(3, 0), "53");
}

#[test]
fn actual_footer_formulas_and_retained_settings_are_rewritten_independently() {
    use visigrid_engine::cell::CellComment;
    let (mut wb, id) = sales();
    let sheet = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(sheet, 0, 0, "Value");
    wb.set_cell_value_tracked(sheet, 1, 0, "1");
    let other = wb
        .create_table(
            wb.sheet(sheet).unwrap().id,
            TableRange {
                start_row: 0,
                end_row: 1,
                start_col: 0,
                end_col: 0,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(other, true, Default::default())
        .unwrap();
    wb.set_table_total(other, 0, custom("=SUM(Sales[Amount])+5"))
        .unwrap();
    wb.sheet_mut(sheet).unwrap().toggle_bold(2, 0);
    wb.sheet_mut(sheet).unwrap().set_comment(
        2,
        0,
        Some(CellComment {
            text: "Keep me".into(),
            author: "QA".into(),
        }),
    );
    let mut catalog = wb.saved_tables();
    catalog
        .sheets
        .iter_mut()
        .flat_map(|s| &mut s.tables)
        .find(|t| t.id == other)
        .unwrap()
        .totals
        .as_mut()
        .unwrap()
        .columns[0] = custom("=SUM(Sales[Amount])+7");
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let commit = wb.remove_table(id).unwrap();
    assert_eq!(wb.sheet(sheet).unwrap().get_display(2, 0), "65");
    assert!(
        wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0]
            .formula
            .as_ref()
            .unwrap()
            .ends_with("+7")
    );
    assert!(wb.sheet(sheet).unwrap().get_format(2, 0).bold);
    assert_eq!(
        wb.sheet(sheet).unwrap().comment(2, 0).unwrap().text,
        "Keep me"
    );
    wb.apply_table_commit(&commit, true).unwrap();
    wb.apply_table_commit(&commit, false).unwrap();
    wb.set_table_totals_visible(other, false, Default::default())
        .unwrap();
    assert!(
        wb.table(other).unwrap().1.totals.as_ref().unwrap().columns[0]
            .formula
            .as_ref()
            .unwrap()
            .ends_with("+5")
    );
    // Hiding preserves comments; clearing the retained comment is required to
    // show a footer again, as with any occupied destination.
    wb.sheet_mut(sheet).unwrap().set_comment(2, 0, None);
    wb.set_table_totals_visible(other, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(sheet).unwrap().get_display(2, 0), "65");
}

#[test]
fn changed_foreign_totals_and_unrepresentable_sources_refuse_atomically() {
    let (mut wb, id) = sales();
    let sheet = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(sheet, 0, 0, "Value");
    wb.set_cell_value_tracked(sheet, 1, 0, "1");
    let other = wb
        .create_table(
            wb.sheet(sheet).unwrap().id,
            TableRange {
                start_row: 0,
                end_row: 1,
                start_col: 0,
                end_col: 0,
            },
            "OtherData",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(other, true, Default::default())
        .unwrap();
    wb.set_table_total(other, 0, custom("=SUM(Sales[Amount])"))
        .unwrap();
    let commit = wb.remove_table(id).unwrap();
    wb.set_table_total(other, 0, custom("=123")).unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet(sheet).unwrap().get_display(2, 0), "123");
    let empty = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 8,
                end_row: 8,
                start_col: 0,
                end_col: 0,
            },
            "Empty",
        )
        .unwrap()
        .table_id();
    wb.set_table_total(other, 0, custom("=SUM(Empty)")).unwrap();
    let before = serde_json::to_value(wb.saved_tables()).unwrap();
    let revision = wb.revision();
    assert!(wb.remove_table(empty).unwrap_err().contains("empty"));
    assert_eq!(wb.revision(), revision);
    assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
    assert_eq!(wb.sheet(sheet).unwrap().get_raw(2, 0), "=SUM(Empty)");
}
