use std::{
    io::{Read, Write},
    path::Path,
};
use visigrid_engine::{
    cell::CellComment,
    sheet::SheetId,
    table::{TableId, TableRange, TableStyle},
    workbook::Workbook,
};
use visigrid_io::{native, xlsx};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (c, s) in ["Qty", "Price", "Amount"].iter().enumerate() {
        wb.set_cell_value_tracked(0, 2, c + 1, s);
    }
    for r in 3..8 {
        wb.set_cell_value_tracked(0, r, 1, &(r - 1).to_string());
        wb.set_cell_value_tracked(0, r, 2, "10");
    }
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 7,
                end_col: 3,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(id, 3, 4, "=[@Qty]*C5", true)
        .unwrap();
    wb.set_cell_value_tracked(0, 4, 3, "777");
    wb.clear_cell_tracked(0, 5, 3);
    wb.set_cell_value_tracked(0, 6, 3, "=1+2");
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[Amount])");
    wb.sheet_mut(0).unwrap().set_comment(
        4,
        3,
        Some(CellComment {
            text: "Manual override".into(),
            author: "QA".into(),
        }),
    );
    (wb, id)
}
fn xml(path: &Path, name: &str) -> String {
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut out = String::new();
    zip.by_name(name).unwrap().read_to_string(&mut out).unwrap();
    out
}
fn rewrite(source: &Path, dest: &Path, change: impl Fn(&str, String) -> (String, String)) {
    let mut input = zip::ZipArchive::new(std::fs::File::open(source).unwrap()).unwrap();
    let mut output = zip::ZipWriter::new(std::fs::File::create(dest).unwrap());
    for i in 0..input.len() {
        let mut entry = input.by_index(i).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let name = entry.name();
        let (name, bytes) = if name.ends_with(".xml") || name.ends_with(".rels") {
            let (name, data) = change(name, String::from_utf8(bytes).unwrap());
            (name, data.into_bytes())
        } else {
            (name.to_string(), bytes)
        };
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output.write_all(&bytes).unwrap();
    }
    output.finish().unwrap();
}
#[test]
fn roundtrip_preserves_rules_overrides_comments_and_cross_sheet_formulas() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("tables.xlsx");
    let (mut wb, _) = book();
    for pass in 0..3 {
        let report = if pass == 1 {
            let (bytes, report) = xlsx::export_to_buffer(&wb, None).unwrap();
            std::fs::write(&file, bytes).unwrap();
            report
        } else {
            xlsx::export(&wb, &file, None).unwrap()
        };
        assert_eq!(report.tables_exported, 1);
        assert!(!report.has_warnings(), "{:?}", report.warnings);
        let metadata = xml(&file, "xl/tables/table1.xml");
        assert!(metadata.contains("ref=\"B3:D8\""));
        assert!(
            metadata.contains(
                "<calculatedColumnFormula>[[#This Row],[Qty]]*C4</calculatedColumnFormula>"
            ),
            "{metadata}"
        );
        let (mut loaded, report) = xlsx::import(&file).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        assert_eq!(report.tables_skipped, 0);
        let (sid, table) = loaded.table_by_name("Sales").unwrap();
        let id = table.id;
        assert_eq!(table.range.start_row, 2);
        assert_eq!(
            table.columns[2].formula.as_deref(),
            Some("=[[#This Row],[Qty]]*C4")
        );
        assert_eq!(loaded.sheet(0).unwrap().get_display(3, 3), "20");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(5, 3), "");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(6, 3), "=1+2");
        assert_eq!(loaded.sheet(0).unwrap().get_display(7, 3), "60");
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "860");
        assert_eq!(
            loaded.sheet(0).unwrap().comment(4, 3).unwrap().text,
            "Manual override"
        );
        assert!(loaded.sheet(0).unwrap().is_calculated_exception(5, 3));
        assert!(!loaded.sheet(0).unwrap().is_calculated_exception(3, 3));
        // Appended records inherit the rule after a native save/reopen, too.
        let native_file = dir.path().join("tables.sheet");
        native::save_workbook(&loaded, &native_file).unwrap();
        loaded = native::load_workbook(&native_file).unwrap();
        assert_eq!(loaded.table(id).unwrap().0, sid);
        wb = loaded.clone();
        loaded
            .append_table_rows(id, 1, &[(8, 1, "7".into()), (8, 2, "10".into())])
            .unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 3), "70");
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "930");
    }
}
#[test]
fn imports_external_writer_tables_and_long_this_row_syntax() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("external.xlsx");
    let mut external = rust_xlsxwriter::Workbook::new();
    let ws = external.add_worksheet();
    ws.write_number(1, 0, 3).unwrap();
    ws.write_number(2, 0, 4).unwrap();
    ws.add_table(
        0,
        0,
        2,
        1,
        &rust_xlsxwriter::Table::new()
            .set_name("Items")
            .set_columns(&[
                rust_xlsxwriter::TableColumn::new().set_header("Qty"),
                rust_xlsxwriter::TableColumn::new()
                    .set_header("Total")
                    .set_formula("=[@[Qty]]*2"),
            ]),
    )
    .unwrap();
    external.save(&file).unwrap();
    let (mut wb, report) = xlsx::import(&file).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 1), "6");
    let id = wb.table_by_name("Items").unwrap().1.id;
    wb.append_table_rows(id, 1, &[(3, 0, "5".into())]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "10");
    let (values, report) = xlsx::import_with_options(
        &file,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(report.tables_imported, 1);
    assert!(values.table_by_name("Items").unwrap().1.columns[1]
        .formula
        .is_none());
    assert!(!values.sheet(0).unwrap().get_raw(1, 1).starts_with('='));
}
#[test]
fn relationships_bind_multiple_tables_without_part_number_assumptions() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("tables.xlsx");
    let changed = dir.path().join("renumbered.xlsx");
    let (mut wb, _) = book();
    wb.set_cell_value_tracked(1, 3, 0, "Other");
    wb.set_cell_value_tracked(1, 4, 0, "12");
    wb.create_table(
        wb.sheet(1).unwrap().id,
        TableRange {
            start_row: 3,
            start_col: 0,
            end_row: 4,
            end_col: 0,
        },
        "OtherData",
    )
    .unwrap();
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.replace("tables/table1.xml", "tables/custom-sales.xml"),
            data.replace("tables/table1.xml", "tables/custom-sales.xml"),
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 2, "{:?}", report.warnings);
    assert_eq!(
        loaded.table_by_name("Sales").unwrap().0,
        loaded.sheet(0).unwrap().id
    );
    assert_eq!(
        loaded.table_by_name("OtherData").unwrap().0,
        loaded.sheet(1).unwrap().id
    );
}
#[test]
fn unsupported_or_corrupt_tables_keep_cells_and_report_loss() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("tables.xlsx");
    let changed = dir.path().join("changed.xlsx");
    let (wb, _) = book();
    xlsx::export(&wb, &file, None).unwrap();
    for (before, after, reason) in [
        ("totalsRowShown=\"0\"", "totalsRowCount=\"1\"", "totals row"),
        (
            "totalsRowShown=\"0\"",
            "headerRowCount=\"0\"",
            "hidden headers",
        ),
        (
            "totalsRowShown=\"0\"",
            "tableType=\"queryTable\"",
            "external",
        ),
        ("count=\"3\"", "count=\"4\"", "inconsistent"),
        ("id=\"2\"", "id=\"1\"", "duplicate"),
        ("name=\"Qty\"", "name=\"Wrong\"", "header"),
        ("B3:D8", "B3:XFE8", "beyond"),
    ] {
        rewrite(&file, &changed, |name, data| {
            (
                name.into(),
                if name == "xl/tables/table1.xml" {
                    assert!(data.contains(before));
                    data.replace(before, after)
                } else {
                    data
                },
            )
        });
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_skipped, 1, "{reason}: {:?}", report.warnings);
        assert_eq!(loaded.tables().count(), 0);
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.to_lowercase().contains(reason)),
            "{reason}: {:?}",
            report.warnings
        );
    }
}
#[test]
fn header_only_export_refuses_before_touching_destination() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("existing.xlsx");
    std::fs::write(&file, b"keep me").unwrap();
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "Header");
    wb.create_table(
        SheetId(1),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 0,
        },
        "EmptyData",
    )
    .unwrap();
    assert!(xlsx::export(&wb, &file, None)
        .unwrap_err()
        .contains("only headers"));
    assert!(xlsx::export_to_buffer(&wb, None).is_err());
    assert_eq!(std::fs::read(file).unwrap(), b"keep me");
}
#[test]
fn xml_escaping_and_at_signs_in_headers_and_literals_survive() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("escaping.xlsx");
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "@Qty & more");
    wb.set_cell_value_tracked(0, 0, 1, "Output");
    wb.set_cell_value_tracked(0, 1, 0, "2");
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 1,
                end_col: 1,
            },
            "Escaped",
        )
        .unwrap()
        .table_id();
    wb.set_calculated_column(
        id,
        1,
        1,
        "=IF([@['@Qty & more]]<3,\"a@b & c\",\"no\")",
        true,
    )
    .unwrap();
    wb.set_table_style(id, TableStyle { banded_rows: false })
        .unwrap();
    xlsx::export(&wb, &file, None).unwrap();
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_display(1, 1), "a@b & c");
    assert!(!loaded.table_by_name("Escaped").unwrap().1.style.banded_rows);
    assert!(loaded.table_by_name("Escaped").unwrap().1.columns[1]
        .formula
        .as_ref()
        .unwrap()
        .contains("\"a@b & c\""));
}

#[test]
fn filter_losses_are_reported_and_all_stored_records_survive() {
    use visigrid_engine::{
        filter::SortDirection,
        table_view::{TableSort, TableViewSpec},
    };
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("filtered.xlsx");
    let changed = dir.path().join("excel-filter.xlsx");
    let (mut wb, id) = book();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Descending,
    });
    wb.set_table_view_spec(SheetId(1), Some(view)).unwrap();
    let warnings = xlsx::table_export_warnings(&wb).unwrap();
    assert_eq!(warnings.len(), 1);
    let report = xlsx::export(&wb, &file, None).unwrap();
    assert_eq!(report.warnings, warnings);
    assert!(report.full_report().contains("every record"));
    assert!(report
        .full_report_with_context("filtered.xlsx")
        .contains("every record"));
    rewrite(&file, &changed, |name, data| {
        (name.into(),match name {
        "xl/tables/table1.xml"=>data.replace("<autoFilter ref=\"B3:D8\"/>","<autoFilter ref=\"B3:D8\"><filterColumn colId=\"0\"><filters><filter val=\"2\"/></filters></filterColumn></autoFilter>"),
        "xl/worksheets/sheet1.xml"=>data.replace("<row r=\"5\"", "<row hidden=\"1\" r=\"5\""),
        _=>data,
    })
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert!(
        report
            .warnings
            .iter()
            .any(|w| w.contains("All records are shown")),
        "{:?}",
        report.warnings
    );
    assert!(!report.imported_layouts[0].hidden_rows.contains(&4));
    for r in 3..8 {
        assert_eq!(loaded.sheet(0).unwrap().get_raw(r, 1), (r - 1).to_string());
    }
    assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "860");
}

#[test]
fn invalid_rule_keeps_table_and_existing_cells_without_fill_rule() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("rule.xlsx");
    let changed = dir.path().join("bad-rule.xlsx");
    let (wb, _) = book();
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/tables/table1.xml" {
                data.replace("[[#This Row],[Qty]]*C4", "1+")
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert!(loaded.table_by_name("Sales").unwrap().1.columns[2]
        .formula
        .is_none());
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("automatic-fill rule")));
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 3), "20");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
}

#[test]
fn blank_first_record_keeps_rule_and_values_only_keeps_offset_cells() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("blank-first.xlsx");
    let (mut wb, _) = book();
    wb.clear_cell_tracked(0, 3, 3);
    xlsx::export(&wb, &file, None).unwrap();
    let (mut loaded, report) = xlsx::import(&file).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    let id = loaded.table_by_name("Sales").unwrap().1.id;
    assert!(loaded.table(id).unwrap().1.columns[2].formula.is_some());
    assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 3), "");
    loaded
        .append_table_rows(id, 1, &[(8, 1, "7".into()), (8, 2, "10".into())])
        .unwrap();
    assert_eq!(loaded.sheet(0).unwrap().get_display(8, 3), "70");
    let (values, report) = xlsx::import_with_options(
        &file,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(values.sheet(0).unwrap().get_raw(2, 3), "Amount");
    assert_eq!(values.sheet(0).unwrap().get_raw(7, 2), "10");
    assert_eq!(values.sheet(0).unwrap().get_raw(4, 3), "777");
    assert!(values.table_by_name("Sales").unwrap().1.columns[2]
        .formula
        .is_none());
}
