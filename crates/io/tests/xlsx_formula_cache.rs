use calamine::{Data, Reader};
use std::io::Read;
use visigrid_engine::{formula::eval::Value, workbook::Workbook};
use visigrid_io::xlsx;

fn xml(bytes: &[u8], part: &str) -> String {
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut text = String::new();
    zip.by_name(part)
        .unwrap()
        .read_to_string(&mut text)
        .unwrap();
    text
}
fn cell_xml<'a>(xml: &'a str, address: &str) -> &'a str {
    let start = xml.find(&format!("<c r=\"{address}\"")).unwrap();
    let end = xml[start..].find("</c>").unwrap() + start + 4;
    &xml[start..end]
}

#[test]
fn scalar_caches_preserve_numbers_text_booleans_errors_and_escapes() {
    let mut wb = Workbook::new();
    let formulas = [
        "=42.125",
        "=\"001\"",
        "=\"1e3\"",
        "=\"TRUE\"",
        "=\"#DIV/0!\"",
        "=1=1",
        "=1=0",
        "=1/0",
        "=\"\"",
        "=\"  \"",
        "=\"a & <b>\"",
        "=\"_x0001_\"",
        "=HOSTTEXT()",
        "=DATE(2026,10,4)",
    ];
    for (r, formula) in formulas.iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, formula);
    }
    wb.sheet_mut(0).unwrap().set_number_format(
        0,
        0,
        visigrid_engine::cell::NumberFormat::Percent { decimals: 2 },
    );
    // A host function can legitimately return XML control characters.
    wb.sheet(0)
        .unwrap()
        .cache_computed(12, 0, Value::Text("\u{1}\rx".into()));
    wb.sheet_mut(0).unwrap().set_text(0, 2, "_x0001_"); // Shared-string decoding must not happen twice.
    let before: Vec<_> = (0..formulas.len())
        .map(|r| wb.sheet(0).unwrap().get_cached_value(r, 0))
        .collect();
    let revision = wb.revision();
    let (bytes, report) =
        xlsx::export_to_buffer_with_order(&wb, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let document = xml(&bytes, "xl/worksheets/sheet1.xml");
    for (cell, kind, value) in [
        ("A1", "n", "42.125"),
        ("A2", "str", "001"),
        ("A3", "str", "1e3"),
        ("A4", "str", "TRUE"),
        ("A5", "str", "#DIV/0!"),
        ("A6", "b", "1"),
        ("A7", "b", "0"),
        ("A8", "e", "#DIV/0!"),
        ("A9", "str", ""),
        ("A10", "str", "  "),
    ] {
        let raw = cell_xml(&document, cell);
        assert!(raw.contains(&format!("t=\"{kind}\"")), "{raw}");
        assert!(raw.contains(&format!("<v>{value}</v>")), "{raw}");
    }
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("typed.xlsx");
    std::fs::write(&path, &bytes).unwrap();
    let mut excel: calamine::Xlsx<_> = calamine::open_workbook(&path).unwrap();
    let range = excel.worksheet_range("Sheet1").unwrap();
    assert_eq!(range.get_value((0, 0)), Some(&Data::Float(42.125)));
    assert_eq!(range.get_value((1, 0)), Some(&Data::String("001".into())));
    assert_eq!(range.get_value((5, 0)), Some(&Data::Bool(true)));
    assert_eq!(
        range.get_value((7, 0)),
        Some(&Data::Error(calamine::CellErrorType::Div0))
    );
    let (values, _) = xlsx::import_with_options(
        &path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    for (r, expected) in [
        (1, "001"),
        (2, "1e3"),
        (3, "TRUE"),
        (4, "#DIV/0!"),
        (7, "#DIV/0!"),
        (8, ""),
        (9, "  "),
        (10, "a & <b>"),
        (11, "_x0001_"),
        (12, "\u{1}\rx"),
    ] {
        assert_eq!(values.sheet(0).unwrap().get_raw(r, 0), expected, "row {r}");
    }
    assert_eq!(values.sheet(0).unwrap().get_raw(0, 2), "_x0001_");
    assert_eq!(
        values.sheet(0).unwrap().get_computed_value(1, 0),
        Value::Text("001".into())
    );
    assert_eq!(
        values.sheet(0).unwrap().get_computed_value(8, 0),
        Value::Text("".into())
    );
    assert_eq!(wb.revision(), revision);
    assert_eq!(
        (0..formulas.len())
            .map(|r| wb.sheet(0).unwrap().get_cached_value(r, 0))
            .collect::<Vec<_>>(),
        before
    );
    let (recomputed, _) = xlsx::import(&path).unwrap();
    assert_eq!(
        recomputed.sheet(0).unwrap().get_computed_value(0, 0),
        Value::Number(42.125)
    );
    assert_eq!(recomputed.sheet(0).unwrap().get_raw(1, 0), formulas[1]);
}

#[test]
fn absent_and_unrepresentable_caches_are_omitted_and_custom_results_are_kept() {
    let mut wb = Workbook::new();
    for r in 0..6 {
        wb.set_cell_value_tracked(0, r, 0, "=HOSTFUNCTION()");
    }
    let sheet = wb.sheet(0).unwrap();
    sheet.clear_cached(0, 0);
    sheet.cache_computed(1, 0, Value::Error("#CYCLE!".into()));
    sheet.cache_computed(2, 0, Value::Number(f64::NAN));
    sheet.cache_computed(3, 0, Value::Number(123.5));
    sheet.cache_computed(5, 0, Value::Empty);
    // Last row retains the engine's unknown-function error.
    let (bytes, report) =
        xlsx::export_to_buffer_with_order(&wb, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(
        report
            .warnings
            .iter()
            .any(|w| w.contains("3 formula results")
                && w.contains("1 not calculated")
                && w.contains("2 unsupported")),
        "{:?}",
        report.warnings
    );
    assert_eq!(
        report.warnings,
        xlsx::table_export_warnings_with_order(&wb, None, xlsx::ExportOrder::Stored).unwrap()
    );
    let document = xml(&bytes, "xl/worksheets/sheet1.xml");
    for address in ["A1", "A2", "A3"] {
        let c = cell_xml(&document, address);
        assert!(c.contains("<f>"));
        assert!(!c.contains("<v>"), "{c}");
    }
    assert!(cell_xml(&document, "A4").contains("<v>123.5</v>"));
    assert!(cell_xml(&document, "A6").contains("t=\"str\""));
    assert!(cell_xml(&document, "A6").contains("<v></v>"));
    assert!(cell_xml(&document, "A5").contains("t=\"e\""));
    assert!(cell_xml(&document, "A5").contains("<v>#NAME?</v>"));
    assert_eq!(
        wb.sheet(0).unwrap().get_cached_value(3, 0),
        Some(Value::Number(123.5))
    );
}

#[test]
fn sorted_table_caches_follow_rows_and_stored_totals_keep_calculated_results() {
    use visigrid_engine::{
        filter::SortDirection,
        table::TableRange,
        table_view::{TableSort, TableViewSpec},
    };
    let mut wb = Workbook::new();
    for (r, text) in ["Key", "3", "1", "2"].iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, text);
    }
    wb.set_cell_value_tracked(0, 0, 1, "Amount");
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(id, 1, 1, "=[@Key]*10", true)
        .unwrap();
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("table.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Sorted).unwrap();
    let (loaded, _) = xlsx::import_with_options(
        &path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(
        (1..4)
            .map(|r| loaded.sheet(0).unwrap().get_display(r, 1))
            .collect::<Vec<_>>(),
        ["10", "20", "30"]
    );
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.define_name_for_cell("Footer", 0, 4, 1).unwrap();
    wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(1, 0, 0, "=Footer");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (loaded, report) = xlsx::import_with_options(
        &path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().get_display(4, 1),
        "60",
        "{:?}",
        report.warnings
    );
    assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "60");
}
