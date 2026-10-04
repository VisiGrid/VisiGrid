use visigrid_engine::{named_range::NamedRange, table::TableRange, workbook::Workbook};
use visigrid_io::{native, xlsx};

#[test]
fn names_descriptions_and_cross_sheet_formulas_roundtrip_in_both_modes() {
    let dir = tempfile::tempdir().unwrap();
    let mut wb = Workbook::new();
    wb.rename_sheet(0, "O'Brien & Sales");
    wb.set_cell_value_tracked(0, 0, 0, "10");
    wb.set_cell_value_tracked(0, 1, 0, "20");
    wb.named_ranges_mut()
        .set(
            NamedRange::cell("Selected", 0, 0, 0)
                .with_description("A & B <total> \"quoted\" 'label'"),
        )
        .unwrap();
    wb.define_name_for_range("Amounts", 0, 0, 0, 1, 0).unwrap();
    wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=Selected+SUM(Amounts)");
    for order in [xlsx::ExportOrder::Stored, xlsx::ExportOrder::Sorted] {
        let path = dir.path().join("names.xlsx");
        xlsx::export_with_order(&wb, &path, None, order).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert!(report.warnings.is_empty(), "{:?}", report.warnings);
        assert_eq!(
            loaded.named_ranges().get("Selected"),
            wb.named_ranges().get("Selected")
        );
        assert_eq!(
            loaded.named_ranges().get("Amounts"),
            wb.named_ranges().get("Amounts")
        );
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "40");
        loaded.set_cell_value_tracked(0, 0, 0, "15");
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "50");
        let native_path = dir.path().join("names.sheet");
        native::save_workbook(&loaded, &native_path).unwrap();
        assert_eq!(
            native::load_workbook(&native_path)
                .unwrap()
                .named_ranges()
                .get("Selected"),
            wb.named_ranges().get("Selected")
        );
        let (values, _) = xlsx::import_with_options(
            &path,
            &xlsx::ImportOptions {
                values_only: true,
                ..Default::default()
            },
        )
        .unwrap();
        assert_eq!(
            values.named_ranges().get("Amounts"),
            wb.named_ranges().get("Amounts")
        );
        assert_eq!(values.sheet(0).unwrap().get_display(0, 0), "10");
        assert!(!values.sheet(1).unwrap().get_raw(0, 0).starts_with('='));
    }
}

#[test]
fn independent_writer_names_import_without_flattening_scope_or_expression_types() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("foreign.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    excel
        .add_worksheet()
        .set_name("Data")
        .unwrap()
        .write_number(0, 0, 12)
        .unwrap();
    excel.add_worksheet().set_name("Other").unwrap();
    for (name, target) in [
        ("Good", "=Data!$A$1"),
        ("Rectangle", "=Data!$A$1:$B$3"),
        ("Relative", "=Data!A1"),
        ("Constant", "=42"),
        ("Expression", "=SUM(Data!$A$1:$A$3)"),
        ("WholeColumn", "=Data!$A:$A"),
        ("Missing", "=Absent!$A$1"),
        ("External", "='[Book.xlsx]Data'!$A$1"),
        ("Duplicate", "=Data!$A$1"),
        ("duplicate", "=Data!$B$1"),
        ("Scoped", "=Data!$A$1"),
        ("Other!Scoped", "=Other!$A$1"),
        ("Other!LocalOnly", "=Other!$A$1"),
    ] {
        excel.define_name(name, target).unwrap();
    }
    excel
        .worksheet_from_name("Other")
        .unwrap()
        .write_formula(
            0,
            0,
            rust_xlsxwriter::Formula::new("Good*2").set_result("24"),
        )
        .unwrap();
    excel.save(&path).unwrap();
    let (values, _) = xlsx::import_with_options(
        &path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(values.sheet(1).unwrap().get_display(0, 0), "24");
    assert!(values.named_ranges().get("Good").is_some());
    let (wb, report) = xlsx::import(&path).unwrap();
    assert_eq!(wb.named_ranges().list().len(), 2);
    assert_eq!(
        wb.named_ranges().get("Good").unwrap().reference_string(),
        "A1"
    );
    for name in [
        "Relative",
        "Constant",
        "Expression",
        "WholeColumn",
        "Missing",
        "External",
        "Duplicate",
        "duplicate",
        "Scoped",
        "LocalOnly",
    ] {
        assert!(wb.named_ranges().get(name).is_none(), "{name}");
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.contains(&format!("'{name}'"))),
            "{name}: {:?}",
            report.warnings
        );
    }
}

fn sorted_book() -> Workbook {
    use visigrid_engine::{
        filter::SortDirection,
        table_view::{TableSort, TableViewSpec},
    };
    let mut wb = Workbook::new();
    for (row, text) in ["Key", "3", "1", "2"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row, 0, text);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 0,
            },
            "Records",
        )
        .unwrap()
        .table_id();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(view))
        .unwrap();
    wb
}

#[test]
fn sorted_export_moves_named_cells_and_keeps_source_unchanged() {
    let mut wb = sorted_book();
    wb.define_name_for_cell("Selected", 0, 1, 0).unwrap();
    wb.define_name_for_range("Pair", 0, 2, 0, 3, 0).unwrap();
    wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=Selected+SUM(Pair)");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("sorted.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Sorted).unwrap();
    let (out, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(
        out.named_ranges()
            .get("Selected")
            .unwrap()
            .reference_string(),
        "A4"
    );
    assert_eq!(
        out.named_ranges().get("Pair").unwrap().reference_string(),
        "A2:A3"
    );
    assert_eq!(out.sheet(1).unwrap().get_display(0, 0), "6");
    assert_eq!(
        wb.named_ranges()
            .get("Selected")
            .unwrap()
            .reference_string(),
        "A2"
    );
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 0), "3");
}

#[test]
fn split_name_target_and_invalid_export_preserve_existing_destination() {
    let mut wb = sorted_book();
    wb.define_name_for_range("Split", 0, 1, 0, 2, 0).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("kept.xlsx");
    std::fs::write(&path, b"keep me").unwrap();
    assert!(
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Sorted)
            .unwrap_err()
            .contains("Defined name 'Split'")
    );
    assert_eq!(std::fs::read(&path).unwrap(), b"keep me");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (out, _) = xlsx::import(&path).unwrap();
    assert_eq!(
        out.named_ranges().get("Split").unwrap().reference_string(),
        "A2:A3"
    );
    std::fs::write(&path, b"keep me").unwrap();
    wb.named_ranges_mut()
        .set(NamedRange::cell("Broken", 99, 0, 0))
        .unwrap();
    assert!(
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored)
            .unwrap_err()
            .contains("missing target sheet")
    );
    assert_eq!(std::fs::read(&path).unwrap(), b"keep me");
}
