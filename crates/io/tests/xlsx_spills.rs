use std::io::{Read, Write};
use visigrid_engine::{
    formula::eval::{Array2D, Value},
    workbook::Workbook,
};
use visigrid_io::{native, xlsx};

fn xml(path: &std::path::Path, part: &str) -> String {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut xml = String::new();
    zip.by_name(part).unwrap().read_to_string(&mut xml).unwrap();
    xml
}
fn replace(path: &std::path::Path, out: &std::path::Path, old: &str, new: &str) {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut writer = zip::ZipWriter::new(std::fs::File::create(out).unwrap());
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).unwrap();
        if entry.name() == "xl/worksheets/sheet1.xml" {
            let mut xml = String::new();
            entry.read_to_string(&mut xml).unwrap();
            assert!(xml.contains(old));
            writer
                .start_file(entry.name(), zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(xml.replace(old, new).as_bytes()).unwrap();
        } else {
            writer.raw_copy_file(entry).unwrap();
        }
    }
    writer.finish().unwrap();
}
#[test]
fn dynamic_range_caches_roundtrip_as_values_and_rebuild_as_live_spills() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "3");
    wb.set_cell_value_tracked(0, 2, 1, "=SEQUENCE(A1,2,10,5)");
    wb.set_cell_value_tracked(0, 8, 1, "Note below spill");
    wb.sheet_mut(0).unwrap().set_bold(3, 2, true);
    wb.sheet_mut(0).unwrap().set_comment(
        3,
        2,
        Some(visigrid_engine::cell::CellComment {
            text: "Receiver note".into(),
            author: "QA".into(),
        }),
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "35");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("spilled.xlsx");
    let report = xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let document = xml(&path, "xl/worksheets/sheet1.xml");
    assert!(document.contains("ref=\"B3:C5\""), "{document}");
    assert!(document.contains("t=\"array\""));
    assert!(xml(&path, "xl/metadata.xml").contains("dynamicArrayProperties"));
    let (values, _) = xlsx::import_with_options(
        &path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    for (r, c, v) in [
        (2, 1, "10"),
        (2, 2, "15"),
        (3, 1, "20"),
        (3, 2, "25"),
        (4, 1, "30"),
        (4, 2, "35"),
    ] {
        assert_eq!(values.sheet(0).unwrap().get_raw(r, c), v);
        assert!(!values.sheet(0).unwrap().is_spill_receiver(r, c));
    }
    let (mut loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_display(4, 2), "35");
    assert!(loaded.sheet(0).unwrap().is_spill_receiver(4, 2));
    assert!(loaded.sheet(0).unwrap().get_format(3, 2).bold);
    assert_eq!(
        loaded.sheet(0).unwrap().comment(3, 2).unwrap().text,
        "Receiver note"
    );
    loaded.set_cell_value_tracked(0, 0, 0, "2");
    assert_eq!(
        loaded.sheet(0).unwrap().get_spill_info(2, 1).unwrap().rows,
        2
    );
    assert!(loaded.sheet(0).unwrap().get_spill_value(4, 2).is_none());
    assert_eq!(loaded.sheet(0).unwrap().get_raw(8, 1), "Note below spill");
    let path = dir.path().join("spilled.sheet");
    native::save_workbook(&loaded, &path).unwrap();
    let loaded = native::load_workbook(&path).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 2), "25");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "35");
}
#[test]
fn heterogeneous_spill_caches_keep_types_and_unsupported_array_members_are_not_lost() {
    use calamine::{Data, Reader};
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 1, 1, "=HOSTARRAY()");
    let array = Array2D::from_vec(vec![
        vec![
            Value::Number(1.0),
            Value::Text("001".into()),
            Value::Boolean(true),
        ],
        vec![
            Value::Error("#DIV/0!".into()),
            Value::Text(" \u{1}\r_x0001_ ".into()),
            Value::Empty,
        ],
    ]);
    assert!(wb.sheet_mut(0).unwrap().apply_spill(1, 1, &array));
    wb.sheet(0)
        .unwrap()
        .cache_computed(1, 1, Value::Number(1.0));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("typed-array.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let mut excel: calamine::Xlsx<_> = calamine::open_workbook(&path).unwrap();
    let range = excel.worksheet_range("Sheet1").unwrap();
    assert_eq!(range.get_value((1, 2)), Some(&Data::String("001".into())));
    assert_eq!(range.get_value((1, 3)), Some(&Data::Bool(true)));
    assert_eq!(
        range.get_value((2, 1)),
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
    assert_eq!(values.sheet(0).unwrap().get_raw(2, 2), " \u{1}\r_x0001_ ");
    assert_eq!(values.sheet(0).unwrap().get_raw(2, 1), "#DIV/0!");
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(
        report
            .warnings
            .iter()
            .any(|w| w.contains("cached result cells were kept")),
        "{:?}",
        report.warnings
    );
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 2), "001");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 2), " \u{1}\r_x0001_ ");
}
#[test]
fn independent_array_fixture_rebuilds_and_bad_metadata_never_clears_other_formulas() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("external.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let ws = excel.add_worksheet();
    ws.write_dynamic_array_formula(
        1,
        1,
        2,
        2,
        rust_xlsxwriter::Formula::new("=SEQUENCE(2,2)").set_result("1"),
    )
    .unwrap();
    ws.write_number(1, 2, 2).unwrap();
    ws.write_number(2, 1, 3).unwrap();
    ws.write_number(2, 2, 4).unwrap();
    ws.write_formula(3, 1, "=9+9").unwrap();
    excel.save(&path).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_display(2, 2), "4");
    assert!(loaded.sheet(0).unwrap().is_spill_receiver(2, 2));
    let changed = dir.path().join("claimed.xlsx");
    replace(&path, &changed, "ref=\"B2:C3\"", "ref=\"B2:C4\"");
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert!(
        report
            .warnings
            .iter()
            .any(|w| w.contains("independent formulas")),
        "{:?}",
        report.warnings
    );
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 1), "18");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 2), "4");
    let oversized = dir.path().join("oversized.xlsx");
    replace(&path, &oversized, "ref=\"B2:C3\"", "ref=\"B2:XFD1048576\"");
    let (loaded, report) = xlsx::import(&oversized).unwrap();
    assert!(
        report.warnings.iter().any(|w| w.contains("5,000,000")),
        "{:?}",
        report.warnings
    );
    assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 2), "4");
}

#[test]
fn table_source_spills_follow_totals_and_append_after_roundtrip() {
    use visigrid_engine::table::TableRange;
    let mut wb = Workbook::new();
    for (r, v) in ["Amount", "30", "10"].into_iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, v);
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 0,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(
        0,
        0,
        4,
        "=SEQUENCE(ROWS(Sales[Amount]),1,SUM(Sales[[#Totals],[Amount]]))",
    );
    assert_eq!(
        wb.active_sheet().get_display(1, 4),
        "41",
        "anchor: {}",
        wb.active_sheet().get_display(0, 4)
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("table-array.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (mut loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let id = loaded.active_sheet().tables()[0].id;
    assert!(loaded.active_sheet().is_spill_receiver(1, 4));
    loaded
        .append_table_rows(id, 1, &[(3, 0, "20".into())])
        .unwrap();
    assert_eq!(loaded.active_sheet().get_display(4, 0), "60");
    assert_eq!(loaded.active_sheet().get_display(2, 4), "62");
    assert_eq!(loaded.active_sheet().get_spill_info(0, 4).unwrap().rows, 3);
}

#[test]
fn error_at_array_anchor_is_a_valid_spilled_result() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=1/0");
    wb.set_cell_value_tracked(0, 1, 0, "4");
    wb.set_cell_value_tracked(0, 0, 2, "=TRANSPOSE(A1:A2)");
    assert!(wb.active_sheet().get_spill_info(0, 2).is_some());
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("array-error.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert!(loaded
        .active_sheet()
        .get_display(0, 2)
        .starts_with("#DIV/0!"));
    assert!(loaded.active_sheet().is_spill_receiver(0, 3));
    assert_eq!(loaded.active_sheet().get_display(0, 3), "4");
}

#[test]
fn overlapping_arrays_and_legacy_scalar_arrays_preserve_cached_members() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("overlap.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let ws = excel.add_worksheet();
    ws.write_dynamic_array_formula(0, 0, 1, 1, "=SEQUENCE(2,2)")
        .unwrap();
    ws.write_dynamic_array_formula(1, 1, 2, 2, "=SEQUENCE(2,2)")
        .unwrap();
    ws.write_number(0, 1, 22).unwrap();
    ws.write_number(2, 2, 44).unwrap();
    ws.write_array_formula(4, 0, 5, 0, "=SUM(1,2)").unwrap();
    ws.write_number(5, 0, 33).unwrap();
    excel.save(&path).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert_eq!(
        report
            .warnings
            .iter()
            .filter(|w| w.contains("overlapping array"))
            .count(),
        2
    );
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("could not be recalculated")));
    assert_eq!(loaded.active_sheet().get_raw(0, 1), "22");
    assert_eq!(loaded.active_sheet().get_raw(2, 2), "44");
    assert_eq!(loaded.active_sheet().get_raw(5, 0), "33");
}

#[test]
fn independently_written_filter_with_combined_excel_namespace_rebuilds() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("filter.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let ws = excel.add_worksheet();
    for (r, value) in [1, 0, 2].into_iter().enumerate() {
        ws.write_number(r as u32, 0, value).unwrap();
    }
    ws.write_dynamic_array_formula(
        0,
        4,
        1,
        4,
        rust_xlsxwriter::Formula::new("=FILTER(A1:A3,A1:A3>0)").set_result("1"),
    )
    .unwrap();
    ws.write_number(1, 4, 2).unwrap();
    excel.save(&path).unwrap();
    assert!(xml(&path, "xl/worksheets/sheet1.xml").contains("_xlfn._xlws.FILTER"));
    let (mut loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(report.formulas_with_unknowns, 0);
    assert_eq!(loaded.active_sheet().get_display(1, 4), "2");
    assert!(loaded.active_sheet().is_spill_receiver(1, 4));
    loaded.set_cell_value_tracked(0, 1, 0, "3");
    assert_eq!(loaded.active_sheet().get_display(1, 4), "3");
    assert_eq!(loaded.active_sheet().get_display(2, 4), "2");
}

#[test]
fn spills_beyond_excel_bounds_refuse_without_replacing_the_destination() {
    let mut wb = Workbook::new();
    // More than u16::MAX columns must not wrap during writer conversion.
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(1,65537)");
    assert!(wb.active_sheet().get_spill_info(0, 0).is_some());
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("keep.xlsx");
    std::fs::write(&path, b"original file").unwrap();
    let error = xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap_err();
    assert!(error.contains("Excel's worksheet limits"), "{error}");
    assert_eq!(std::fs::read(&path).unwrap(), b"original file");
}

#[test]
fn blocked_lifted_array_retains_array_identity_and_blocker() {
    let mut wb = Workbook::new();
    for row in 0..3 { wb.set_cell_value_tracked(0, row, 1, "2"); }
    wb.set_cell_value_tracked(0, 1, 0, "blocker");
    wb.set_cell_value_tracked(0, 0, 0, "=B1:B3*2");
    assert!(wb.active_sheet().get_cell(0, 0).spill_error().is_some());
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("blocked.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let document = xml(&path, "xl/worksheets/sheet1.xml");
    assert!(document.contains("t=\"array\" ref=\"A1\""), "{document}");
    let (mut loaded, _) = xlsx::import(&path).unwrap();
    assert_eq!(loaded.active_sheet().get_raw(1, 0), "blocker");
    loaded.set_cell_value_tracked(0, 1, 0, "");
    assert_eq!(loaded.active_sheet().get_display(2, 0), "4");
}

#[test]
fn out_of_bounds_hidden_row_does_not_abort_import_or_survive_in_layout() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("valid.xlsx");
    let altered = dir.path().join("invalid-hidden.xlsx");
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "Keep");
    wb.active_sheet_mut().set_manual_hidden_rows([2].into()).unwrap();
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    replace(&path, &altered, "</sheetData>", "<row r=\"1048577\" hidden=\"1\"/></sheetData>");
    let (loaded, report) = xlsx::import(&altered).unwrap();
    assert_eq!(loaded.active_sheet().get_raw(0, 0), "Keep");
    assert_eq!(loaded.active_sheet().manual_hidden_rows(), [2].into());
    assert_eq!(report.imported_layouts[0].hidden_rows, vec![2]);
    assert!(report.warnings.iter().any(|w| w.contains("outside the worksheet")));
}
