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
