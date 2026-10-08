use visigrid_engine::{
    cell::CellStyle,
    cond_format::CondStyle,
    table::TableRange,
    validation::{CellRange, ValidationRule},
    workbook::Workbook,
};
use visigrid_io::{json, native, xlsx};

#[test]
fn converted_rule_context_survives_native_json_and_excel_without_a_table() {
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
            wb.set_cell_value_tracked(0, r, c, value);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 0, 4, 1)],
        "=IFERROR([@Amount]<[@Limit],FALSE)",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut().validations.set(
        CellRange::new(1, 0, 3, 0),
        ValidationRule::custom("=[@Amount]<[@Limit]"),
    );
    let other = wb.add_sheet_named("List").unwrap();
    wb.sheet_mut(other).unwrap().validations.set(
        CellRange::new(0, 0, 2, 0),
        ValidationRule::list_range("Sales[Amount]"),
    );
    wb.remove_table(id).unwrap();
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..3 {
        let path = dir.path().join(if mode == 2 {
            "converted.xlsx"
        } else {
            "converted.sheet"
        });
        let loaded = match mode {
            0 => {
                native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap();
                native::load_workbook(&path).unwrap()
            }
            1 => {
                json::import_any(&json::export_workbook(&wb, &[], 0).unwrap())
                    .unwrap()
                    .0
            }
            _ => {
                xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
                xlsx::import(&path).unwrap().0
            }
        };
        assert!(loaded.tables().next().is_none());
        for (r, expected) in [false, true, false, true, false].into_iter().enumerate() {
            assert_eq!(
                loaded.active_sheet().has_cond_format(r, 0),
                expected,
                "mode {mode}, row {r}"
            );
        }
        for (r, expected) in [(1, true), (2, false), (3, true)] {
            assert_eq!(
                loaded.validate_cell(0, r, 0).is_valid(),
                expected,
                "mode {mode}, row {r}"
            );
        }
        assert_eq!(
            loaded.get_list_items(other, 1, 0).unwrap().items,
            ["10", "20", "30"]
        );
    }
}

#[test]
fn converted_foreign_calculated_and_dormant_totals_rules_survive_reopening_and_append() {
    use visigrid_engine::table::TableTotal;
    let mut wb = Workbook::new();
    for (row, value) in ["Amount", "10", "20", "30"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, value);
    }
    let source = wb
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
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "Value");
    wb.set_cell_value_tracked(other, 1, 0, "5");
    let id = wb
        .create_table(
            wb.sheet(other).unwrap().id,
            TableRange {
                start_row: 0,
                end_row: 1,
                start_col: 0,
                end_col: 0,
            },
            "Summary",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(id, 0, 1, "=SUM(Sales[Amount])", true)
        .unwrap();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        id,
        0,
        TableTotal {
            function: Some("custom".into()),
            formula: Some("=SUM(Sales[Amount])+SUM([Value])".into()),
            label: None,
        },
    )
    .unwrap();
    wb.set_table_totals_visible(id, false, Default::default())
        .unwrap();
    wb.remove_table(source).unwrap();
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..3 {
        let path = dir.path().join(if mode == 2 {
            "totals.xlsx"
        } else {
            "totals.sheet"
        });
        let mut loaded = match mode {
            0 => {
                native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap();
                native::load_workbook(&path).unwrap()
            }
            1 => {
                json::import_any(&json::export_workbook(&wb, &[], 0).unwrap())
                    .unwrap()
                    .0
            }
            _ => {
                xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
                xlsx::import(&path).unwrap().0
            }
        };
        let id = loaded
            .tables()
            .find(|(_, t)| t.name == "Summary")
            .unwrap()
            .1
            .id;
        assert!(loaded.sheet(0).unwrap().tables().is_empty());
        assert!(!loaded.table(id).unwrap().1.totals.as_ref().unwrap().visible);
        loaded
            .set_table_totals_visible(id, true, Default::default())
            .unwrap();
        assert_eq!(
            loaded.sheet(other).unwrap().get_display(2, 0),
            "120",
            "mode {mode}"
        );
        loaded.append_table_rows(id, 1, &[]).unwrap();
        assert_eq!(
            loaded.sheet(other).unwrap().get_display(3, 0),
            "180",
            "mode {mode}"
        );
        loaded.set_cell_value_tracked(0, 1, 0, "100");
        assert_eq!(
            loaded.sheet(other).unwrap().get_display(3, 0),
            "450",
            "mode {mode}"
        );
    }
}
