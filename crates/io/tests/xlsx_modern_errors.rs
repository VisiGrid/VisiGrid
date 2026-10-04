use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
    path::Path,
};
use visigrid_engine::{
    formula::eval::{Array2D, Value},
    workbook::Workbook,
};
use visigrid_io::{native, xlsx};

fn xml(bytes: &[u8], part: &str) -> String {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut s = String::new();
    zip.by_name(part).unwrap().read_to_string(&mut s).unwrap();
    s
}
fn change(bytes: &[u8], parts: BTreeMap<&str, String>) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..zip.len() {
        let entry = zip.by_index(i).unwrap();
        if !parts.contains_key(entry.name()) {
            out.raw_copy_file(entry).unwrap();
        }
    }
    for (name, s) in parts {
        out.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        out.write_all(s.as_bytes()).unwrap();
    }
    out.finish().unwrap().into_inner()
}
fn values(path: &Path) -> (Workbook, xlsx::ImportResult) {
    xlsx::import_with_options(
        path,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap()
}

#[test]
fn blocked_spill_and_calc_errors_keep_rich_caches_and_readable_fallbacks() {
    use calamine::{CellErrorType, Data, Reader};
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 1, 1, "obstruction");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(3,2)");
    wb.set_cell_value_tracked(0, 0, 4, "1");
    wb.set_cell_value_tracked(0, 1, 4, "2");
    wb.set_cell_value_tracked(0, 0, 6, "=FILTER(E1:E2,E1:E2>10)");
    wb.set_cell_value_tracked(0, 4, 0, "=SEQUENCE(2,2)");
    assert!(wb.active_sheet().get_display(0, 0).starts_with("#SPILL!"));
    assert!(wb.active_sheet().get_display(0, 6).starts_with("#CALC!"));
    let error = wb.active_sheet().get_cell(0, 0);
    assert_eq!(
        error
            .spill_error()
            .unwrap()
            .dimensions
            .as_ref()
            .unwrap()
            .rows,
        3
    );
    let (bytes, report) =
        xlsx::export_to_buffer_with_order(&wb, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert!(
        xlsx::table_export_warnings_with_order(&wb, None, xlsx::ExportOrder::Stored)
            .unwrap()
            .is_empty()
    );
    let sheet = xml(&bytes, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("vm="));
    assert!(!sheet.contains("<v>#SPILL!</v>"));
    assert!(xml(&bytes, "xl/richData/rdrichvalue.xml").contains("<v>8</v><v>2</v><v>1</v>"));
    let meta = xml(&bytes, "xl/metadata.xml");
    assert!(meta.contains("XLDAPR") && meta.contains("XLRICHVALUE"));
    assert!(meta.contains("dynamicArrayProperties"));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("errors.xlsx");
    std::fs::write(&path, &bytes).unwrap();
    // An independent reader without rich error support still opens the workbook.
    let mut excel: calamine::Xlsx<_> = calamine::open_workbook(&path).unwrap();
    let range = excel.worksheet_range("Sheet1").unwrap();
    assert_eq!(
        range.get_value((0, 0)),
        Some(&Data::Error(CellErrorType::Value))
    );
    let (saved, report) = values(&path);
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(saved.active_sheet().get_raw(0, 0), "#SPILL!");
    assert_eq!(saved.active_sheet().get_raw(0, 6), "#CALC!");
    assert_eq!(saved.active_sheet().get_raw(5, 1), "4");
    let (mut live, _) = xlsx::import(&path).unwrap();
    assert!(live.active_sheet().get_display(0, 0).starts_with("#SPILL!"));
    live.clear_cell_tracked(0, 1, 1);
    assert_eq!(live.active_sheet().get_display(2, 1), "6");
    let native_path = dir.path().join("errors.sheet");
    native::save_workbook(&live, &native_path).unwrap();
    assert_eq!(
        native::load_workbook(&native_path)
            .unwrap()
            .active_sheet()
            .get_display(2, 1),
        "6"
    );
    assert_eq!(wb.active_sheet().get_raw(1, 1), "obstruction");
}

#[test]
fn modern_scalar_and_receiver_errors_roundtrip_without_converting_literal_text() {
    let labels = [
        "#SPILL!",
        "#CONNECT!",
        "#BLOCKED!",
        "#UNKNOWN!",
        "#CALC!",
        "#BUSY!",
        "#TIMEOUT!",
    ];
    let mut wb = Workbook::new();
    for (r, label) in labels.iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, "=CUSTOM_ERROR()");
        wb.active_sheet()
            .cache_computed(r, 0, Value::Error((*label).into()));
        wb.sheet_mut(0).unwrap().set_text_exact(r, 2, label);
    }
    wb.set_cell_value_tracked(0, 10, 0, "=CUSTOM_ARRAY()");
    let array = Array2D::from_vec(vec![vec![
        Value::Number(1.0),
        Value::Error("#CALC!".into()),
    ]]);
    assert!(wb.sheet_mut(0).unwrap().apply_spill(10, 0, &array));
    wb.active_sheet().cache_computed(10, 0, Value::Number(1.0));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("all-errors.xlsx");
    let report = xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let (loaded, report) = values(&path);
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    for (r, label) in labels.iter().enumerate() {
        assert_eq!(loaded.active_sheet().get_raw(r, 0), *label);
        assert_eq!(loaded.active_sheet().get_raw(r, 2), *label);
    }
    assert_eq!(loaded.active_sheet().get_raw(10, 1), "#CALC!");
    let bytes = std::fs::read(&path).unwrap();
    let rich = xml(&bytes, "xl/richData/rdrichvalue.xml");
    assert!(rich.contains("count=\"7\""), "{rich}");
}

// Independent metadata fixture with non-identity indexes, reordered structure
// keys and nonstandard part paths: self-roundtrips cannot prove these bindings.
fn independent() -> Vec<u8> {
    let mut excel = rust_xlsxwriter::Workbook::new();
    let ws = excel.add_worksheet();
    ws.write_formula(
        0,
        0,
        rust_xlsxwriter::Formula::new("=1+1").set_result("#VALUE!"),
    )
    .unwrap();
    ws.write_formula(
        1,
        0,
        rust_xlsxwriter::Formula::new("=2+2").set_result("#VALUE!"),
    )
    .unwrap();
    ws.write_string(0, 1, "#VALUE!").unwrap();
    let bytes = excel.save_to_buffer().unwrap();
    let sheet = xml(&bytes, "xl/worksheets/sheet1.xml")
        .replace("t=\"str\"", "t=\"e\"")
        .replace("r=\"A1\"", "r=\"A1\" vm=\"2\"")
        .replace("r=\"A2\"", "r=\"A2\" vm=\"1\"");
    let rels = xml(&bytes, "xl/_rels/workbook.xml.rels").replace("</Relationships>", r#"<Relationship Id="m" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="meta/error.xml"/><Relationship Id="d" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="/xl/data/errors.xml"/><Relationship Id="s" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="data/../schema/errors.xml"/></Relationships>"#);
    let content_types = xml(&bytes, "[Content_Types].xml").replace("</Types>", r#"<Override PartName="/xl/meta/error.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/><Override PartName="/xl/data/errors.xml" ContentType="application/vnd.ms-excel.rdRichValue+xml"/><Override PartName="/xl/schema/errors.xml" ContentType="application/vnd.ms-excel.rdRichValueStructure+xml"/></Types>"#);
    change(&bytes, [
        ("[Content_Types].xml", content_types),
        ("xl/worksheets/sheet1.xml", sheet), ("xl/_rels/workbook.xml.rels", rels),
        ("xl/schema/errors.xml", r#"<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_error"><k n="subType" t="i"/><k n="errorType" t="i"/></s></rvStructures>"#.into()),
        ("xl/data/errors.xml", r#"<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2"><rv s="0"><v>0</v><v>11</v></rv><rv s="0"><v>0</v><v>13</v></rv></rvData>"#.into()),
        ("xl/meta/error.xml", r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="2"><metadataType name="XLDAPR"/><metadataType name="XLRICHVALUE"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="2"><bk><extLst><ext uri="{3E2802C4-A4D2-4D8B-9148-E3BE6C30E623}"><rd:rvb i="1"/></ext></extLst></bk><bk><extLst><ext uri="{3E2802C4-A4D2-4D8B-9148-E3BE6C30E623}"><rd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="2"><bk><rc t="2" v="1"/></bk><bk><rc t="2" v="0"/></bk></valueMetadata></metadata>"#.into()),
    ].into())
}

#[test]
fn independent_metadata_resolves_relationships_keys_and_all_index_layers() {
    let bytes = independent();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("external.xlsx");
    std::fs::write(&path, &bytes).unwrap();
    let (loaded, report) = values(&path);
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.active_sheet().get_raw(0, 0), "#CALC!");
    assert_eq!(loaded.active_sheet().get_raw(1, 0), "#UNKNOWN!");
    assert_eq!(loaded.active_sheet().get_raw(0, 1), "#VALUE!");
    let (live, _) = xlsx::import(&path).unwrap();
    assert_eq!(live.active_sheet().get_display(0, 0), "2");
    assert_eq!(live.active_sheet().get_display(1, 0), "4");
}

#[test]
fn malformed_or_external_rich_metadata_keeps_fallback_cells_and_warns() {
    let bytes = independent();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("bad.xlsx");
    for parts in [
        [(
            "xl/meta/error.xml",
            xml(&bytes, "xl/meta/error.xml").replace("i=\"1\"", "i=\"99\""),
        )]
        .into(),
        [(
            "xl/_rels/workbook.xml.rels",
            xml(&bytes, "xl/_rels/workbook.xml.rels")
                .replace("Id=\"d\"", "Id=\"d\" TargetMode=\"External\""),
        )]
        .into(),
        [(
            "xl/data/errors.xml",
            xml(&bytes, "xl/data/errors.xml").replace("<v>13</v>", "<v>9999</v>"),
        )]
        .into(),
    ] {
        std::fs::write(&path, change(&bytes, parts)).unwrap();
        let (loaded, report) = values(&path);
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.contains("Modern Excel error metadata was not restored")),
            "{:?}",
            report.warnings
        );
        assert_eq!(loaded.active_sheet().get_raw(0, 0), "#VALUE!");
        assert_eq!(loaded.active_sheet().get_raw(1, 0), "#VALUE!");
    }
}

#[test]
fn escaped_error_numbers_and_absent_value_metadata_are_read_correctly() {
    let bytes = independent();
    let bytes = change(
        &bytes,
        [
            (
                "xl/data/errors.xml",
                xml(&bytes, "xl/data/errors.xml").replace("<v>13</v>", "<v>&#49;<![CDATA[3]]></v>"),
            ),
            (
                "xl/worksheets/sheet1.xml",
                xml(&bytes, "xl/worksheets/sheet1.xml").replace("vm=\"1\"", "vm=\"0\""),
            ),
        ]
        .into(),
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("entities.xlsx");
    std::fs::write(&path, &bytes).unwrap();
    let (loaded, report) = values(&path);
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.active_sheet().get_raw(0, 0), "#CALC!");
    assert_eq!(loaded.active_sheet().get_raw(1, 0), "#VALUE!");
}
