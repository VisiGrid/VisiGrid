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

#[test]
fn native_authored_totals_survive_native_and_excel_roundtrip() {
    use visigrid_engine::table::TableTotal;
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(id, 1, TableTotal { label: Some("Grand total".into()), ..Default::default() }).unwrap();
    wb.set_table_total(id, 2, TableTotal { function: Some("custom".into()), formula: Some("=SUM([Price])*2".into()), label: None }).unwrap();
    let expected = wb.sheet(0).unwrap().get_display(8, 3);
    let path = dir.path().join("native-totals.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let wb = native::load_workbook(&path).unwrap();
    let path = dir.path().join("native-totals.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let table_xml = xml(&path, "xl/tables/table1.xml");
    assert!(table_xml.contains("totalsRowCount=\"1\""));
    assert!(table_xml.contains("totalsRowLabel=\"Grand total\""));
    assert!(table_xml.contains("totalsRowFunction=\"sum\""));
    let (mut wb, _) = xlsx::import(&path).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 3), expected);
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 2), "100");
    let id = wb.tables().next().unwrap().1.id;
    wb.set_table_totals_visible(id, false, Default::default()).unwrap();
    let path = dir.path().join("dormant-native-totals.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (mut wb, _) = xlsx::import(&path).unwrap();
    let id = wb.tables().next().unwrap().1.id;
    assert!(wb.table(id).unwrap().1.totals_row().is_none());
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 3), expected);
    assert_eq!(wb.sheet(0).unwrap().get_display(8, 2), "100");
    assert_eq!(wb.sheet(0).unwrap().get_raw(8, 1), "Grand total");
}

#[test]
fn moved_totals_survive_native_and_excel_with_comments_and_custom_formulas() {
    use visigrid_engine::table::TableTotal;
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(id, 2, TableTotal { function: Some("custom".into()), formula: Some("=SUM([Price])*2".into()), label: None }).unwrap();
    wb.sheet_mut(0).unwrap().set_comment(8, 2, Some(CellComment { text: "Custom total".into(), author: "QA".into() }));
    wb.sheet_mut(0).unwrap().toggle_bold(8, 2);
    wb.append_table_rows(id, 2, &[(8, 1, "8".into()), (8, 2, "15".into()), (9, 1, "9".into()), (9, 2, "20".into())]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(10, 2), "170");
    let native_path = dir.path().join("moved.sheet");
    native::save_workbook(&wb, &native_path).unwrap();
    let wb = native::load_workbook(&native_path).unwrap();
    let path = dir.path().join("moved.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (loaded, _) = xlsx::import(&path).unwrap();
    assert_eq!(loaded.tables().next().unwrap().1.totals_row(), Some(10));
    assert_eq!(loaded.sheet(0).unwrap().get_display(10, 2), "170");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(10, 2), "=SUM([Price])*2");
    assert!(loaded.sheet(0).unwrap().get_format(10, 2).bold);
    assert_eq!(loaded.sheet(0).unwrap().comment(10, 2).unwrap().text, "Custom total");
    assert!(loaded.sheet(0).unwrap().comment(8, 2).is_none());
    assert_eq!(loaded.sheet(0).unwrap().get_display(9, 3), "180");
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
        (
            "totalsRowShown=\"0\"",
            "totalsRowCount=\"2\"",
            "totals-row count",
        ),
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
    wb.set_table_style(
        id,
        TableStyle {
            banded_rows: false,
            ..Default::default()
        },
    )
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
fn materialized_sort_and_checkbox_filters_import_with_all_records() {
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
        "xl/tables/table1.xml"=>data.replace("<autoFilter ref=\"B3:D8\"></autoFilter>","<autoFilter ref=\"B3:D8\"><filterColumn colId=\"0\"><filters><filter val=\"2\"/></filters></filterColumn></autoFilter>"),
        "xl/worksheets/sheet1.xml"=>data.replace("<row r=\"5\"", "<row hidden=\"1\" r=\"5\""),
        _=>data,
    })
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(
        loaded
            .sheet(0)
            .unwrap()
            .table_view_spec()
            .unwrap()
            .filters
            .len(),
        1
    );
    assert!(!report.imported_layouts[0].hidden_rows.contains(&4));
    for r in 3..8 {
        assert_eq!(loaded.sheet(0).unwrap().get_raw(r, 1), (9 - r).to_string());
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

fn select_values(wb: &mut Workbook, id: TableId, column: usize, rows: &[usize], buttons: bool) {
    use visigrid_engine::{
        filter::{ColumnFilter, FilterKey},
        table_view::{TableFilter, TableViewSpec},
    };
    let (sid, table) = wb.table(id).unwrap();
    let sheet = wb.sheet_by_id(sid).unwrap();
    let keys = rows
        .iter()
        .map(|r| {
            FilterKey::from_value(&sheet.get_computed_value(*r, table.range.start_col + column))
                .normalized()
        })
        .collect();
    let mut spec = TableViewSpec::new(id);
    spec.show_filter_buttons = buttons;
    spec.filters.push(TableFilter {
        column: table.columns[column].id,
        criteria: ColumnFilter {
            selected: Some(keys),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(sid, Some(spec)).unwrap();
}

#[test]
fn builtin_styles_and_flags_survive_excel_native_excel_and_old_native_defaults() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("styles.xlsx");
    let native_file = dir.path().join("styles.sheet");
    let (mut wb, id) = book();
    let old: TableStyle = serde_json::from_str(r#"{"banded_rows":false}"#).unwrap();
    assert!(!old.banded_rows);
    assert_eq!(old.excel_style.as_deref(), Some("TableStyleMedium2"));
    for name in [
        None,
        Some("TableStyleLight1"),
        Some("TableStyleLight21"),
        Some("TableStyleMedium28"),
        Some("TableStyleDark11"),
    ] {
        let style = TableStyle {
            excel_style: name.map(str::to_owned),
            banded_rows: false,
            banded_columns: true,
            first_column: true,
            last_column: true,
        };
        wb.set_table_style(id, style.clone()).unwrap();
        xlsx::export(&wb, &file, None).unwrap();
        let meta = xml(&file, "xl/tables/table1.xml");
        assert!(meta.contains("showColumnStripes=\"1\""));
        assert!(meta.contains("showFirstColumn=\"1\""));
        assert!(meta.contains("showLastColumn=\"1\""));
        let (loaded, report) = xlsx::import(&file).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        assert_eq!(loaded.table_by_name("Sales").unwrap().1.style, style);
        native::save_workbook(&loaded, &native_file).unwrap();
        let native_copy = native::load_workbook(&native_file).unwrap();
        assert_eq!(native_copy.table_by_name("Sales").unwrap().1.style, style);
        xlsx::export(&native_copy, &file, None).unwrap();
        let (again, _) = xlsx::import(&file).unwrap();
        assert_eq!(again.table_by_name("Sales").unwrap().1.style, style);
    }
    for name in [
        "TableStyleLight0",
        "TableStyleDark12",
        "TableStyleMedium29",
        "TableStyleLight01",
        "CustomStyle",
    ] {
        assert!(!TableStyle::is_builtin_excel_style(name));
        assert!(wb
            .set_table_style(
                id,
                TableStyle {
                    excel_style: Some(name.into()),
                    ..Default::default()
                }
            )
            .is_err());
    }
}

#[test]
fn independent_excel_styles_and_uniform_hidden_buttons_are_retained() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("external-style.xlsx");
    let out = dir.path().join("returned.xlsx");
    let mut external = rust_xlsxwriter::Workbook::new();
    let ws = external.add_worksheet();
    ws.write_number(1, 0, 10).unwrap();
    ws.add_table(
        0,
        0,
        2,
        1,
        &rust_xlsxwriter::Table::new()
            .set_name("External")
            .set_style(rust_xlsxwriter::TableStyle::Dark7)
            .set_banded_columns(true)
            .set_first_column(true)
            .set_last_column(true)
            .set_autofilter(false),
    )
    .unwrap();
    external.save(&file).unwrap();
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert_eq!(report.tables_imported, 1);
    let table = loaded.table_by_name("External").unwrap().1;
    assert_eq!(table.style.excel_style.as_deref(), Some("TableStyleDark7"));
    assert!(table.style.banded_columns && table.style.first_column && table.style.last_column);
    assert!(
        !loaded
            .sheet(0)
            .unwrap()
            .table_view_spec()
            .unwrap()
            .show_filter_buttons
    );
    xlsx::export(&loaded, &out, None).unwrap();
    assert!(!xml(&out, "xl/tables/table1.xml").contains("autoFilter"));
    assert!(xml(&out, "xl/tables/table1.xml").contains("name=\"TableStyleDark7\""));
}

#[test]
fn numeric_and_formula_checkbox_filters_roundtrip_with_hidden_rows_and_buttons() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("filtered.xlsx");
    let native_file = dir.path().join("filtered.sheet");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 2, &[3, 5], false); // formula result 20 and blank
    let expected = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
    for pass in 0..2 {
        if pass == 1 {
            wb.sheet_mut(0).unwrap().tab_color = Some([0x12, 0x34, 0x56, 255]);
        }
        let report = xlsx::export(&wb, &file, None).unwrap();
        assert!(report.warnings.is_empty(), "{:?}", report.warnings);
        assert_eq!(report.hidden_rows_exported, 3);
        let meta = xml(&file, "xl/tables/table1.xml");
        assert!(
            meta.contains("<filters blank=\"1\"><filter val=\"20\"/></filters>"),
            "{meta}"
        );
        assert_eq!(meta.matches("hiddenButton=\"1\"").count(), 3);
        let sheet = xml(&file, "xl/worksheets/sheet1.xml");
        assert_eq!(sheet.matches("hidden=\"1\"").count(), 3);
        assert_eq!(sheet.matches("<sheetPr ").count(), 1);
        assert!(sheet.contains("filterMode=\"1\""));
        if pass == 1 {
            assert!(sheet.contains("<tabColor rgb=\"FF123456\"/>"));
        }
        let (loaded, imported) = xlsx::import(&file).unwrap();
        assert!(imported.warnings.is_empty(), "{:?}", imported.warnings);
        assert_eq!(loaded.sheet(0).unwrap().table_view_spec(), Some(&expected));
        assert!(imported.imported_layouts[0].hidden_rows.is_empty());
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(5, 3), "");
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "860");
        native::save_workbook(&loaded, &native_file).unwrap();
        wb = native::load_workbook(&native_file).unwrap();
        assert_eq!(wb.sheet(0).unwrap().table_view_spec(), Some(&expected));
    }
}

#[test]
fn checkbox_text_variants_escaping_and_multiple_columns_survive() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("text.xlsx");
    let (mut wb, id) = book();
    for (r, text) in [
        (3, "West & <south>"),
        (4, "WEST & <SOUTH>"),
        (5, " West & <south> "),
        (6, "East"),
        (7, "East"),
    ] {
        wb.set_cell_value_tracked(0, r, 1, text);
    }
    select_values(&mut wb, id, 0, &[3], true);
    let mut spec = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
    spec.filters.push(visigrid_engine::table_view::TableFilter {
        column: wb.table(id).unwrap().1.columns[2].id,
        criteria: visigrid_engine::filter::ColumnFilter {
            selected: Some(
                [visigrid_engine::filter::NormalizedFilterKey::Blank]
                    .into_iter()
                    .collect(),
            ),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(SheetId(1), Some(spec.clone()))
        .unwrap();
    let report = xlsx::export(&wb, &file, None).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(report.hidden_rows_exported, 4);
    let before = xml(&file, "xl/tables/table1.xml");
    assert!(before.contains("West &amp; &lt;south&gt;"));
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().table_view_spec(), Some(&spec));
    xlsx::export(&loaded, &file, None).unwrap();
    assert_eq!(xml(&file, "xl/tables/table1.xml"), before);
}

#[test]
fn unsupported_or_malformed_filters_keep_whole_table_and_never_apply_partial_criteria() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("base.xlsx");
    let changed = dir.path().join("unsupported.xlsx");
    let (wb, _) = book();
    xlsx::export(&wb, &file, None).unwrap();
    for bad in [
        "<filterColumn colId=\"1\"><dynamicFilter type=\"aboveAverage\"/></filterColumn>",
        "<filterColumn colId=\"1\"><filters><dateGroupItem year=\"2026\" dateTimeGrouping=\"year\"/></filters></filterColumn>",
        "<filterColumn colId=\"1\"><customFilters><customFilter operator=\"greaterThan\" val=\"3\"/></customFilters></filterColumn>",
        "<filterColumn colId=\"99\"><filters><filter val=\"1\"/></filters></filterColumn>",
        "<filterColumn colId=\"0\"><filters><filter val=\"3\"/></filters></filterColumn>",
    ] {
        rewrite(&file, &changed, |name, data| (name.into(), if name == "xl/tables/table1.xml" {
            data.replace("</autoFilter>", &format!("<filterColumn colId=\"0\"><filters><filter val=\"2\"/></filters></filterColumn>{bad}</autoFilter>"))
        } else if name == "xl/worksheets/sheet1.xml" { data.replace("<row r=\"5\"", "<row hidden=\"1\" r=\"5\"") } else { data }));
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        assert_eq!(report.tables_skipped, 0);
        assert!(loaded.sheet(0).unwrap().table_view_spec().is_none());
        assert!(report.warnings.iter().any(|w| w.contains("saved filters/button settings were not imported")), "{:?}", report.warnings);
        assert!(report.imported_layouts[0].hidden_rows.is_empty());
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
    }
}

#[test]
fn unsupported_export_filters_warn_and_do_not_leave_hidden_rows_without_criteria() {
    use visigrid_engine::filter::{TextFilter, TextFilterMode};
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("unsupported-export.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], true);
    for mode in 0..3 {
        let mut spec = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
        spec.filters[0].criteria.text_filter = None;
        if mode == 0 {
            spec.filters[0].criteria.text_filter = Some(TextFilter {
                mode: TextFilterMode::Contains,
                value: "2".into(),
                case_sensitive: false,
            });
        }
        if mode == 1 {
            spec.filters[0].criteria.selected = Some(Default::default());
        }
        if mode == 2 {
            spec.filters[0].criteria.selected = Some(
                [visigrid_engine::filter::NormalizedFilterKey::Number(
                    2.0.into(),
                )]
                .into_iter()
                .collect(),
            );
            wb.sheet_mut(0).unwrap().set_text(4, 1, "2"); // same Excel filter text, different type
        }
        wb.set_table_view_spec(SheetId(1), Some(spec)).unwrap();
        let expected = xlsx::table_export_warnings(&wb).unwrap();
        let report = xlsx::export(&wb, &file, None).unwrap();
        assert_eq!(report.warnings, expected);
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.contains("filter criteria are not exported")),
            "mode {mode}: {:?}",
            report.warnings
        );
        assert_eq!(report.hidden_rows_exported, 0);
        assert!(!xml(&file, "xl/tables/table1.xml").contains("<filters"));
        assert!(!xml(&file, "xl/worksheets/sheet1.xml").contains("hidden=\"1\""));
    }
}

#[test]
fn importing_filters_respects_adjacent_cells_row_heights_and_freeze_boundaries() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("layout.xlsx");
    let changed = dir.path().join("blocked.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], true);
    xlsx::export(&wb, &file, None).unwrap();
    for mode in 0..3 {
        rewrite(&file, &changed, |name, data| {
            (
                name.into(),
                if name == "xl/worksheets/sheet1.xml" {
                    match mode {
                0 => {
                    let start = data.find("<row r=\"4\"").unwrap();
                    let end = start + data[start..].find("</row>").unwrap();
                    let mut data = data;
                    data.insert_str(end, "<c r=\"F4\" t=\"inlineStr\"><is><t>Keep me</t></is></c>");
                    data
                },
                1 => data.replace("<row r=\"4\"", "<row ht=\"24\" customHeight=\"1\" r=\"4\""),
                _ => data.replace("workbookViewId=\"0\"/>", "workbookViewId=\"0\"><pane ySplit=\"4\" topLeftCell=\"A5\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView>"),
            }
                } else {
                    data
                },
            )
        });
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(
            report.tables_imported, 1,
            "mode {mode}: {:?}",
            report.warnings
        );
        if mode == 2 {
            assert!(!loaded.sheet(0).unwrap().table_view_spec().unwrap().filters.is_empty());
            assert_eq!(loaded.sheet(0).unwrap().frozen_panes, (4, 0));
            continue;
        }
        assert!(
            loaded.sheet(0).unwrap().table_view_spec().is_none(),
            "mode {mode}"
        );
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.contains("saved sort/filter/button settings were not imported")),
            "mode {mode}: {:?}",
            report.warnings
        );
    }
}

#[test]
fn overlong_table_names_refuse_before_touching_export_destination() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("existing.xlsx");
    std::fs::write(&file, b"untouched").unwrap();
    let (mut wb, id) = book();
    wb.rename_table(id, &"Long".repeat(64)).unwrap();
    assert!(xlsx::export(&wb, &file, None)
        .unwrap_err()
        .contains("255-character"));
    assert!(xlsx::export_to_buffer(&wb, None).is_err());
    assert_eq!(std::fs::read(file).unwrap(), b"untouched");
}

#[test]
fn multiple_excel_filter_owners_fall_back_without_arbitrarily_choosing_one() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("two-tables.xlsx");
    let changed = dir.path().join("two-views.xlsx");
    let mut external = rust_xlsxwriter::Workbook::new();
    let ws = external.add_worksheet();
    for (row, name) in [(0, "First"), (5, "OtherData")] {
        ws.write_number(row + 1, 0, 1).unwrap();
        ws.write_number(row + 2, 0, 2).unwrap();
        ws.add_table(
            row,
            0,
            row + 2,
            0,
            &rust_xlsxwriter::Table::new().set_name(name),
        )
        .unwrap();
        ws.set_row_hidden(row + 2).unwrap();
    }
    external.save(&file).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name.starts_with("xl/tables/") {
                let pos = data.find("<autoFilter ").unwrap();
                let end = pos + data[pos..].find("/>").unwrap();
                let mut out = data;
                out.replace_range(end..end + 2, "><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>");
                out
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 2, "{:?}", report.warnings);
    assert!(loaded.sheet(0).unwrap().table_view_spec().is_none());
    assert_eq!(
        report
            .warnings
            .iter()
            .filter(|w| w.contains("More than one Table"))
            .count(),
        2
    );
    assert!(report.imported_layouts[0].hidden_rows.is_empty());
    assert_eq!(loaded.sheet(0).unwrap().get_raw(7, 0), "2");
}

#[test]
fn mixed_buttons_and_custom_styles_have_precise_warnings_without_losing_filters() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("base.xlsx");
    let changed = dir.path().join("mixed.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], true);
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/tables/table1.xml" {
                data.replace("TableStyleMedium2", "CompanyCustomStyle")
                    .replace("colId=\"0\"", "colId=\"0\" showButton=\"0\"")
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 1);
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("Custom Excel Table styles")));
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("mixed per-column")));
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec()
    );
}

#[test]
fn boolean_and_error_values_and_values_only_import_keep_filter_membership() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("typed.xlsx");
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 3, 1, "=1=1");
    wb.set_cell_value_tracked(0, 4, 1, "=1/0");
    select_values(&mut wb, id, 0, &[3, 4], true);
    let expected = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
    let report = xlsx::export(&wb, &file, None).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert!(xml(&file, "xl/tables/table1.xml").contains("val=\"#DIV/0!\""));
    let (loaded, _) = xlsx::import(&file).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().table_view_spec(), Some(&expected));
    // Values-only import consumes actual typed caches. Its static boolean/error
    // representation is text, so verify membership rather than identical key types.
    let (values, report) = xlsx::import_with_options(&file, &xlsx::ImportOptions { values_only: true, ..Default::default() }).unwrap();
    assert_eq!(values.sheet(0).unwrap().get_raw(3, 1), "TRUE");
    assert_eq!(values.sheet(0).unwrap().get_raw(4, 1), "#DIV/0!");
    let sheet = values.sheet(0).unwrap();
    let view = visigrid_engine::table_view::TableView::build(sheet, sheet.table_view_spec().unwrap().clone(), 20, None).unwrap();
    assert_eq!((3..8).filter(|r| view.rows().is_data_row_visible(*r)).collect::<Vec<_>>(), vec![3, 4], "{:?}", report.warnings);
    // Numeric calculated-column results and blank overrides retain their criteria.
    let (mut wb, id) = book();
    select_values(&mut wb, id, 2, &[3, 5], true);
    xlsx::export(&wb, &file, None).unwrap();
    let (loaded, report) = xlsx::import_with_options(
        &file,
        &xlsx::ImportOptions {
            values_only: true,
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec(),
        "{:?}",
        report.warnings
    );
    assert!(loaded.table_by_name("Sales").unwrap().1.columns[2]
        .formula
        .is_none());
}

#[test]
fn excel_string_escape_sequences_in_filter_values_are_decoded_once() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("escapes.xlsx");
    let changed = dir.path().join("encoded.xlsx");
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(0, 3, 1, "_x0041_");
    wb.set_cell_value_tracked(0, 4, 1, "A");
    select_values(&mut wb, id, 0, &[3], true);
    xlsx::export(&wb, &file, None).unwrap();
    assert!(xml(&file, "xl/tables/table1.xml").contains("val=\"_x005F_x0041_\""));
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec(),
        "{:?}",
        report.warnings
    );
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/tables/table1.xml" {
                data.replace("_x005F_x0041_", "_x0041_")
            } else {
                data
            },
        )
    });
    let (loaded, _) = xlsx::import(&changed).unwrap();
    let selected = loaded.sheet(0).unwrap().table_view_spec().unwrap().filters[0]
        .criteria
        .selected
        .as_ref()
        .unwrap();
    assert_eq!(
        selected,
        &[visigrid_engine::filter::NormalizedFilterKey::Text(
            "a".into()
        )]
        .into_iter()
        .collect()
    );
}

#[test]
fn imported_filter_preserves_manually_hidden_matching_records() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("base.xlsx");
    let changed = dir.path().join("manual-hidden.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], true);
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/worksheets/sheet1.xml" {
                data.replace("<row r=\"4\"", "<row hidden=\"1\" r=\"4\"")
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec()
    );
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert!(report.imported_layouts[0].hidden_rows.contains(&3));
    let sheet = loaded.sheet(0).unwrap();
    assert!(sheet.manual_hidden_rows().contains(&3));
    assert!(!sheet.build_saved_table_view(sheet.rows).unwrap().unwrap().rows().is_data_row_visible(3));
    let mut cleared = loaded.clone();
    cleared.set_table_view_spec(sheet.id, None).unwrap();
    assert!(cleared.sheet(0).unwrap().manual_hidden_rows().contains(&3));

}

#[test]
fn numeric_filter_spellings_match_numbers_even_with_a_matching_text_record() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("numbers.xlsx");
    let changed = dir.path().join("decimal-filter.xlsx");
    let (mut wb, id) = book();
    wb.sheet_mut(0).unwrap().set_text(4, 1, "2.0");
    select_values(&mut wb, id, 0, &[3, 4], true);
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/tables/table1.xml" {
                data.replace("<filter val=\"2\"/>", "")
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec(),
        "{:?}",
        report.warnings
    );
}

fn set_saved_sort(wb: &mut Workbook, id: TableId, offset: usize, descending: bool, buttons: bool) {
    use visigrid_engine::{
        filter::SortDirection,
        table_view::{TableSort, TableViewSpec},
    };
    let (sid, table) = wb.table(id).unwrap();
    let mut spec = wb
        .sheet_by_id(sid)
        .unwrap()
        .table_view_spec()
        .cloned()
        .unwrap_or_else(|| TableViewSpec::new(id));
    spec.sort = Some(TableSort {
        column: table.columns[offset].id,
        direction: if descending {
            SortDirection::Descending
        } else {
            SortDirection::Ascending
        },
    });
    spec.show_filter_buttons = buttons;
    wb.set_table_view_spec(sid, Some(spec)).unwrap();
}

#[test]
fn materialized_sort_roundtrips_both_directions_and_buttons_with_formula_results() {
    use visigrid_engine::table_view::TableView;
    let dir = tempfile::tempdir().unwrap();
    let plain = dir.path().join("plain.xlsx");
    let file = dir.path().join("sorted.xlsx");
    let native_file = dir.path().join("sorted.sheet");
    for descending in [false, true] {
        for buttons in [false, true] {
            let (mut wb, id) = book();
            xlsx::export(&wb, &plain, None).unwrap();
            set_saved_sort(&mut wb, id, 2, descending, buttons);
            for _ in 0..2 {
                let original = serde_json::to_value(wb.saved_tables()).unwrap();
                let report = xlsx::export(&wb, &file, None).unwrap();
                assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), original);
                assert_eq!(report.warnings.len(), 1);
                assert!(report.warnings[0].contains("saved sort order"));
                assert!(report.warnings[0].contains("Reapplying"));
                assert_eq!(
                    xml(&file, "xl/worksheets/sheet2.xml"),
                    xml(&plain, "xl/worksheets/sheet2.xml")
                );
                let meta = xml(&file, "xl/tables/table1.xml");
                assert!(
                    meta.contains("<sortState ref=\"B4:D8\"><sortCondition ref=\"D4:D8\""),
                    "{meta}"
                );
                assert!(meta.find("<sortState").unwrap() < meta.find("<tableColumns").unwrap());
                assert_eq!(meta.contains("<autoFilter"), buttons);
                let (mut loaded, report) = xlsx::import(&file).unwrap();
                assert!(report.warnings.is_empty(), "{:?}", report.warnings);
                let sheet = loaded.sheet(0).unwrap();
                let spec = sheet.table_view_spec().unwrap().clone();
                assert_eq!(Some(&spec), wb.sheet(0).unwrap().table_view_spec());
                let view = TableView::build(sheet, spec, 10, None).unwrap();
                let rows: Vec<_> = (3..8).map(|r| view.rows().view_to_data(r)).collect();
                assert_eq!(rows, vec![3, 4, 5, 6, 7]);
                let expected = if descending {
                    ["777", "60", "20", "3", ""]
                } else {
                    ["3", "20", "60", "777", ""]
                };
                for (r, value) in (3..8).zip(expected) {
                    assert_eq!(sheet.get_display(r, 3), value);
                }
                assert_eq!(
                    sheet
                        .comment(if descending { 3 } else { 6 }, 3)
                        .unwrap()
                        .text,
                    "Manual override"
                );
                assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "860");
                let imported = serde_json::to_value(loaded.saved_tables()).unwrap();
                native::save_workbook(&loaded, &native_file).unwrap();
                loaded = native::load_workbook(&native_file).unwrap();
                assert_eq!(
                    serde_json::to_value(loaded.saved_tables()).unwrap(),
                    imported
                );
                let mut appended = loaded.clone();
                appended
                    .append_table_rows(id, 1, &[(8, 1, "7".into()), (8, 2, "10".into())])
                    .unwrap();
                assert_eq!(appended.sheet(0).unwrap().get_display(8, 3), "70");
                assert_eq!(appended.sheet(1).unwrap().get_display(0, 0), "930");
                wb = loaded;
            }
        }
    }
}

#[test]
fn independent_sort_states_bind_offsets_to_column_ids_at_either_ooxml_location() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("external.xlsx");
    let changed = dir.path().join("external-sorted.xlsx");
    let mut external = rust_xlsxwriter::Workbook::new();
    let ws = external.add_worksheet();
    ws.write_number(3, 2, 20).unwrap();
    ws.write_number(4, 2, 10).unwrap();
    ws.add_table(
        2,
        2,
        4,
        3,
        &rust_xlsxwriter::Table::new()
            .set_name("Orders")
            .set_style(rust_xlsxwriter::TableStyle::Medium2),
    )
    .unwrap();
    external.save(&file).unwrap();
    for nested in [false, true] {
        for header_in_range in [false, true] {
            let state = format!("<sortState ref=\"$C${}:$D$5\"><sortCondition ref=\"$C$4:$C$5\" sortBy=\"value\"/></sortState>", if header_in_range { 3 } else { 4 });
            rewrite(&file, &changed, |name, data| {
                (
                    name.into(),
                    if name == "xl/tables/table1.xml" {
                        let data = data.replace("<tableColumn id=\"1\"", "<tableColumn id=\"101\"");
                        if nested {
                            data.replace(
                                "<autoFilter ref=\"C3:D5\"/>",
                                &format!("<autoFilter ref=\"C3:D5\">{state}</autoFilter>"),
                            )
                        } else {
                            data.replace("<tableColumns", &format!("{state}<tableColumns"))
                        }
                    } else {
                        data
                    },
                )
            });
            let (wb, report) = xlsx::import(&changed).unwrap();
            assert!(report.warnings.is_empty(), "{:?}", report.warnings);
            let table = wb.table_by_name("Orders").unwrap().1;
            assert_eq!(table.columns[0].id.0, 101);
            let sort = wb
                .sheet(0)
                .unwrap()
                .table_view_spec()
                .unwrap()
                .sort
                .as_ref()
                .unwrap();
            assert_eq!(sort.column, table.columns[0].id);
            assert_eq!(
                sort.direction,
                visigrid_engine::filter::SortDirection::Ascending
            );
            assert_eq!(wb.sheet(0).unwrap().get_raw(3, 2), "20");
        }
    }
}

#[test]
fn unsupported_and_malformed_sorts_keep_filters_instead_of_applying_only_the_first_key() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("filtered.xlsx");
    let changed = dir.path().join("sort.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], true);
    xlsx::export(&wb, &file, None).unwrap();
    let state = |attrs: &str, children: &str| {
        format!("<sortState ref=\"B4:D8\" {attrs}>{children}</sortState>")
    };
    let key = "<sortCondition ref=\"B4:B8\"/>";
    let cases = [
        state("caseSensitive=\"1\"", key),
        state("columnSort=\"true\"", key),
        state("sortMethod=\"stroke\"", key),
        state("", &format!("{key}<sortCondition ref=\"C4:C8\"/>")),
        state(
            "",
            "<sortCondition ref=\"B4:B8\" customList=\"East,West\"/>",
        ),
        state(
            "",
            "<sortCondition ref=\"B4:B8\" sortBy=\"cellColor\" dxfId=\"1\"/>",
        ),
        state(
            "",
            "<sortCondition ref=\"B4:B8\" sortBy=\"icon\" iconId=\"0\"/>",
        ),
        state("", "<sortCondition ref=\"A4:A8\"/>"),
        state("", "<sortCondition ref=\"B4:C8\"/>"),
        state("", "<sortCondition ref=\"B5:B8\"/>"),
        state("", "<sortCondition ref=\"B4:B7\"/>"),
        state("", "<sortCondition ref=\"B4:B8\" descending=\"bad\"/>"),
        state("", ""),
        state("", &format!("{key}<extLst/>")),
        format!("{}{}", state("", key), state("", key)),
        "<sortState ref=\"B4:D7\"><sortCondition ref=\"B4:B7\"/></sortState>".into(),
        "<sortState><sortCondition ref=\"B4:B8\"/></sortState>".into(),
    ];
    for state in cases {
        rewrite(&file, &changed, |name, data| {
            (
                name.into(),
                if name == "xl/tables/table1.xml" {
                    data.replace("<tableColumns", &format!("{state}<tableColumns"))
                } else {
                    data
                },
            )
        });
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_imported, 1);
        assert_eq!(
            loaded.sheet(0).unwrap().table_view_spec(),
            wb.sheet(0).unwrap().table_view_spec(),
            "{state}: {:?}",
            report.warnings
        );
        assert!(
            report
                .warnings
                .iter()
                .any(|w| w.contains("saved Excel sorting was not imported")),
            "{state}: {:?}",
            report.warnings
        );
        assert!(report.imported_layouts[0].hidden_rows.is_empty());
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "860");
    }
}

#[test]
fn sort_and_filters_are_independent_when_filter_metadata_is_unsupported() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("sorted.xlsx");
    let changed = dir.path().join("advanced-filter.xlsx");
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    xlsx::export(&wb, &file, None).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name == "xl/tables/table1.xml" {
                data.replace(
                    "</autoFilter>",
                    "<filterColumn colId=\"0\"><top10 val=\"2\"/></filterColumn></autoFilter>",
                )
            } else if name == "xl/worksheets/sheet1.xml" {
                data.replace("<row r=\"5\"", "<row hidden=\"1\" r=\"5\"")
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec(),
        wb.sheet(0).unwrap().table_view_spec(),
        "{:?}",
        report.warnings
    );
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("saved filters/button settings were not imported")));
    assert!(report.imported_layouts[0].hidden_rows.is_empty());
    // The export-side fallback also retains the independent sort definition.
    select_values(&mut wb, id, 0, &[], true);
    set_saved_sort(&mut wb, id, 0, true, true);
    let report = xlsx::export(&wb, &file, None).unwrap();
    assert_eq!(report.warnings.len(), 2);
    assert!(xml(&file, "xl/tables/table1.xml").contains("<sortState"));
    assert!(!xml(&file, "xl/tables/table1.xml").contains("<filters"));
}

#[test]
fn clearing_sort_removes_excel_metadata_but_retains_filters_and_buttons() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("cleared.xlsx");
    let (mut wb, id) = book();
    select_values(&mut wb, id, 0, &[3], false);
    set_saved_sort(&mut wb, id, 2, true, false);
    xlsx::export(&wb, &file, None).unwrap();
    let (mut wb, report) = xlsx::import(&file).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let mut spec = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
    assert!(spec.sort.is_some());
    spec.sort = None;
    wb.set_table_view_spec(SheetId(1), Some(spec.clone()))
        .unwrap();
    assert!(xlsx::export(&wb, &file, None).unwrap().warnings.is_empty());
    assert!(!xml(&file, "xl/tables/table1.xml").contains("sortState"));
    let (loaded, _) = xlsx::import(&file).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().table_view_spec(), Some(&spec));
}

#[test]
fn sort_only_import_refuses_unsafe_layouts_and_preserves_manual_hidden_rows() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("sort-only.xlsx");
    let changed = dir.path().join("unsafe-sort.xlsx");
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    xlsx::export(&wb, &file, None).unwrap();
    for mode in 0..4 {
        rewrite(&file, &changed, |name, data| {
            (
                name.into(),
                if name == "xl/worksheets/sheet1.xml" {
                    match mode {
                0 => data.replace("<row r=\"4\"", "<row r=\"4\" hidden=\"1\""),
                1 => data.replace("<row r=\"4\"", "<row r=\"4\" ht=\"24\" customHeight=\"1\""),
                2 => data.replace("workbookViewId=\"0\"/>", "workbookViewId=\"0\"><pane ySplit=\"4\" topLeftCell=\"A5\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView>"),
                _ => {
                    let start = data.find("<row r=\"4\"").unwrap();
                    let end = start + data[start..].find("</row>").unwrap();
                    let mut data = data;
                    data.insert_str(end, "<c r=\"F4\" t=\"inlineStr\"><is><t>Keep me</t></is></c>");
                    data
                }
            }
                } else {
                    data
                },
            )
        });
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_imported, 1);
        if mode == 0 {
            let sheet = loaded.sheet(0).unwrap();
            assert!(sheet.table_view_spec().unwrap().sort.is_some());
            assert!(report.imported_layouts[0].hidden_rows.contains(&3));
            assert!(!sheet.build_saved_table_view(sheet.rows).unwrap().unwrap().rows().is_data_row_visible(3));
        } else if mode == 2 {
            assert!(loaded.sheet(0).unwrap().table_view_spec().unwrap().sort.is_some());
            assert_eq!(loaded.sheet(0).unwrap().frozen_panes, (4, 0));
        } else {
            assert!(loaded.sheet(0).unwrap().table_view_spec().is_none(), "mode {mode}");
            assert!(report.warnings.iter().any(|w| w.contains("saved sort/filter/button settings were not imported")), "mode {mode}: {:?}", report.warnings);
        }
        assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 1), "6");
    }
}

#[test]
fn multiple_sorted_tables_respect_the_single_view_owner_rule() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("two-sorts.xlsx");
    let changed = dir.path().join("external.xlsx");
    let mut external = rust_xlsxwriter::Workbook::new();
    let ws = external.add_worksheet();
    for (row, name) in [(0, "FirstData"), (5, "OtherData")] {
        ws.write_number(row + 1, 0, 2).unwrap();
        ws.write_number(row + 2, 0, 1).unwrap();
        ws.add_table(
            row,
            0,
            row + 2,
            0,
            &rust_xlsxwriter::Table::new().set_name(name),
        )
        .unwrap();
    }
    external.save(&file).unwrap();
    rewrite(&file, &changed, |name, data| {
        (
            name.into(),
            if name.starts_with("xl/tables/") {
                let r = if name.ends_with("table1.xml") {
                    "A2:A3"
                } else {
                    "A7:A8"
                };
                data.replace("<tableColumns", &format!("<sortState ref=\"{r}\"><sortCondition ref=\"{r}\"/></sortState><tableColumns"))
            } else {
                data
            },
        )
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 2);
    assert!(loaded.sheet(0).unwrap().table_view_spec().is_none());
    assert_eq!(
        report
            .warnings
            .iter()
            .filter(|w| w.contains("More than one Table"))
            .count(),
        2
    );
}

fn authored_snapshot(wb: &Workbook) -> serde_json::Value {
    serde_json::json!({
        "tables": wb.saved_tables(),
        "sheets": wb.sheets().iter().map(|s| {
            let mut cells: Vec<_> = s.cells_iter().map(|((r,c), cell)|
                (r, c, format!("{:?}", cell.to_cell()), s.get_display(r,c))).collect();
            cells.sort_by_key(|(r,c,_,_)| (*r,*c));
            (s.name.clone(), cells)
        }).collect::<Vec<_>>()
    })
}

#[test]
fn sorted_export_follows_records_for_mixed_absolute_cross_sheet_refs_formats_and_comments() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("materialized.xlsx");
    let (mut wb, id) = book();
    wb.sheet_mut(0).unwrap().set_bold(3, 1, true);
    wb.set_cell_value_tracked(
        1,
        1,
        0,
        "=Sheet1!$B$4+Sheet1!B$4+Sheet1!$B4+(Sheet1!B4+Sheet1!C4)*2",
    );
    wb.set_cell_value_tracked(1, 2, 0, "=SUM(Sheet1!B4:D8)");
    wb.set_cell_value_tracked(1, 3, 0, "=SUM(Sheet1!B:B)");
    select_values(&mut wb, id, 0, &[3, 7], false);
    set_saved_sort(&mut wb, id, 0, true, false);
    let original = authored_snapshot(&wb);
    let report = xlsx::export(&wb, &file, None).unwrap();
    assert_eq!(report.warnings, xlsx::table_export_warnings(&wb).unwrap());
    assert_eq!(authored_snapshot(&wb), original);
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 1), "6");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(7, 1), "2");
    assert!(loaded.sheet(0).unwrap().get_format(7, 1).bold);
    assert_eq!(
        loaded.sheet(0).unwrap().comment(6, 3).unwrap().text,
        "Manual override"
    );
    assert_eq!(
        loaded.sheet(1).unwrap().get_raw(1, 0),
        "=Sheet1!$B$8+Sheet1!B$8+Sheet1!$B8+(Sheet1!B8+Sheet1!C8)*2"
    );
    for r in 0..4 {
        assert_eq!(
            loaded.sheet(1).unwrap().get_display(r, 0),
            wb.sheet(1).unwrap().get_display(r, 0)
        );
    }
    let meta = xml(&file, "xl/tables/table1.xml");
    assert!(meta.contains("<sortState"));
    assert!(meta.contains("hiddenButton=\"1\""));
    let sheet_xml = xml(&file, "xl/worksheets/sheet1.xml");
    for r in [5, 6, 7] {
        assert!(
            sheet_xml.contains(&format!("<row r=\"{r}\" spans=\"2:4\" hidden=\"1\"")),
            "{sheet_xml}"
        );
    }
    // Clearing criteria keeps the new physical order, and all five records remain.
    let mut loaded = loaded;
    loaded.set_table_view_spec(SheetId(1), None).unwrap();
    for r in 3..8 {
        assert_eq!(loaded.sheet(0).unwrap().get_raw(r, 1), (9 - r).to_string());
    }
}

#[test]
fn sorted_file_and_buffer_exports_agree_and_repeated_exports_do_not_mutate_source() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("file.xlsx");
    let buffer = dir.path().join("buffer.xlsx");
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    let original = authored_snapshot(&wb);
    let report = xlsx::export(&wb, &file, None).unwrap();
    let (bytes, buffered) = xlsx::export_to_buffer(&wb, None).unwrap();
    std::fs::write(&buffer, bytes).unwrap();
    assert_eq!(report.warnings, buffered.warnings);
    for part in [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/tables/table1.xml",
        "xl/comments1.xml",
    ] {
        assert_eq!(xml(&file, part), xml(&buffer, part));
    }
    assert_eq!(authored_snapshot(&wb), original);
}

fn assert_sorted_export_refuses(
    wb: &Workbook,
    expected: &str,
    layouts: Option<&[xlsx::ExportLayout]>,
) {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("keep.xlsx");
    std::fs::write(&file, b"existing destination").unwrap();
    let original = authored_snapshot(wb);
    let error = xlsx::export(wb, &file, layouts).unwrap_err();
    assert!(error.contains(expected), "{error}");
    assert!(error.contains("no file was written"), "{error}");
    assert_eq!(std::fs::read(&file).unwrap(), b"existing destination");
    assert!(xlsx::export_to_buffer(wb, layouts)
        .unwrap_err()
        .contains(expected));
    if layouts.is_none() {
        assert!(xlsx::table_export_warnings(wb)
            .unwrap_err()
            .contains(expected));
    }
    assert_eq!(authored_snapshot(wb), original);
}

#[test]
fn sorted_export_refuses_discontiguous_ranges_and_coordinate_or_dynamic_functions_atomically() {
    for (source, expected) in [
        ("=SUM(Sheet1!B4:B5)", "discontiguous"),
        ("=SUM(Sheet1!$B$4:B8)", "discontiguous"),
        ("=SUM(Sheet1!B8:B4)", "Reversed ranges"),
        ("=A2", "Cannot export Tables"),
        ("=ROW(Sheet1!B4)", "Function ROW"),
        ("=INDIRECT(\"Sheet1!B4\")", "Function INDIRECT"),
        ("=OFFSET(Sheet1!B4,1,0)", "Function OFFSET"),
        ("=INDEX(Sales[Qty],1)", "Function INDEX"),
        ("=RAND()", "Function RAND"),
        ("=UnknownName", "Unresolved named reference"),
        ("=1+", "Unexpected"),
    ] {
        let (mut wb, id) = book();
        wb.set_cell_value_tracked(1, 1, 0, source);
        set_saved_sort(&mut wb, id, 2, false, true);
        assert_sorted_export_refuses(&wb, expected, None);
    }
}

#[test]
fn sorted_export_refuses_nonuniform_calculated_rules_before_writing() {
    let (mut wb, id) = book();
    wb.set_calculated_column(id, 3, 3, "=B3", true).unwrap();
    set_saved_sort(&mut wb, id, 0, true, true);
    assert_sorted_export_refuses(&wb, "different fill rules", None);
}

#[test]
fn sorted_export_refuses_unsafe_host_layout_before_writing() {
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    for mode in 0..3 {
        let mut layout = xlsx::ExportLayout::default();
        match mode {
            0 => {
                layout.row_heights.insert(4, 30.0);
            }
            1 => layout.hidden_rows.push(4),
            _ => layout.autofilter_range = Some((2, 1, 7, 3)),
        }
        assert_sorted_export_refuses(&wb, "row layout", Some(&[layout]));
    }
}

#[test]
fn sorted_export_keeps_ties_text_numbers_and_blanks_in_view_order() {
    use visigrid_engine::table_view::TableView;
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("types.xlsx");
    for descending in [false, true] {
        let (mut wb, id) = book();
        wb.set_cell_value_tracked(0, 3, 1, "5");
        wb.set_cell_value_tracked(0, 4, 1, "5");
        wb.set_cell_text_tracked(0, 5, 1, "001");
        wb.clear_cell_tracked(0, 6, 1);
        wb.set_cell_text_tracked(0, 7, 1, "alpha");
        set_saved_sort(&mut wb, id, 0, descending, true);
        let sheet = wb.sheet(0).unwrap();
        let view =
            TableView::build(sheet, sheet.table_view_spec().unwrap().clone(), 8, None).unwrap();
        let records: Vec<_> = (3..8)
            .map(|r| {
                let original = view.rows().view_to_data(r);
                (sheet.get_raw(original, 1), sheet.get_display(original, 3))
            })
            .collect();
        xlsx::export(&wb, &file, None).unwrap();
        let (loaded, _) = xlsx::import(&file).unwrap();
        for (r, (qty, amount)) in (3..8).zip(records) {
            assert_eq!(loaded.sheet(0).unwrap().get_raw(r, 1), qty);
            assert_eq!(loaded.sheet(0).unwrap().get_display(r, 3), amount);
        }
    }
}

#[test]
fn sorted_export_rewrites_references_between_two_sorted_sheets_in_one_plan() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("two-sheets.xlsx");
    let (mut wb, id) = book();
    let sid = wb.add_sheet_named("Second").unwrap();
    for (r, c, value) in [
        (0, 0, "Key"),
        (0, 1, "Linked"),
        (1, 0, "2"),
        (2, 0, "1"),
        (1, 1, "=Sheet1!B4"),
        (2, 1, "=Sheet1!B8"),
    ] {
        wb.set_cell_value_tracked(sid, r, c, value);
    }
    let sid_id = wb.sheet(sid).unwrap().id;
    let second = wb
        .create_table(
            sid_id,
            TableRange {
                start_row: 0,
                end_row: 2,
                start_col: 0,
                end_col: 1,
            },
            "SecondTable",
        )
        .unwrap()
        .table_id();
    wb.set_cell_value_tracked(1, 1, 0, "=Second!$B$2");
    set_saved_sort(&mut wb, id, 0, true, true);
    set_saved_sort(&mut wb, second, 0, false, true);
    let snapshot = authored_snapshot(&wb);
    xlsx::export(&wb, &file, None).unwrap();
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(sid).unwrap().get_raw(1, 1), "=Sheet1!B4");
    assert_eq!(loaded.sheet(sid).unwrap().get_display(1, 1), "6");
    assert_eq!(loaded.sheet(sid).unwrap().get_raw(2, 1), "=Sheet1!B8");
    assert_eq!(loaded.sheet(1).unwrap().get_raw(1, 0), "=Second!$B$3");
    assert_eq!(loaded.sheet(1).unwrap().get_display(1, 0), "2");
    assert_eq!(authored_snapshot(&wb), snapshot);
}

#[test]
fn sorted_export_preserves_contiguous_aggregate_membership_even_when_reversed() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("subset.xlsx");
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(1, 1, 0, "=SUM(Sheet1!$B$4:$B$5)");
    set_saved_sort(&mut wb, id, 0, true, true);
    xlsx::export(&wb, &file, None).unwrap();
    let (loaded, _) = xlsx::import(&file).unwrap();
    assert_eq!(
        loaded.sheet(1).unwrap().get_raw(1, 0),
        "=SUM(Sheet1!$B$7:$B$8)"
    );
    assert_eq!(loaded.sheet(1).unwrap().get_display(1, 0), "5");
}

#[test]
fn sorted_export_refuses_stale_results_without_changing_live_caches() {
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    wb.set_auto_recalc(false);
    wb.set_cell_value_tracked(0, 3, 2, "11");
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 3), "20");
    assert_sorted_export_refuses(&wb, "changes result", None);
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 3), "20");
}

#[test]
fn sorted_export_refuses_validation_conditional_format_and_spills() {
    use visigrid_engine::{
        cell::CellStyle,
        cond_format::CondStyle,
        validation::{CellRange, ValidationRule},
    };
    for mode in 0..3 {
        let (mut wb, id) = book();
        set_saved_sort(&mut wb, id, 0, true, true);
        let expected = match mode {
            0 => {
                wb.sheet_mut(0).unwrap().set_validation(
                    3,
                    1,
                    3,
                    1,
                    ValidationRule::list_inline(vec!["2".into()]),
                );
                "Validation"
            }
            1 => {
                wb.sheet_mut(0).unwrap().cond_formats.add(
                    vec![CellRange::single(3, 1)],
                    "=B4>0",
                    CondStyle::Named(CellStyle::Success),
                );
                "conditional-format"
            }
            _ => {
                wb.set_cell_value_tracked(1, 2, 0, "=SEQUENCE(2,1)");
                "Spilled formulas"
            }

        };
        assert_sorted_export_refuses(&wb, expected, None);
    }
}

#[test]
fn already_sorted_tables_and_unsorted_workbooks_do_not_trigger_materialization_guards() {
    let dir = tempfile::tempdir().unwrap();
    let plain = dir.path().join("plain.xlsx");
    let sorted = dir.path().join("sorted.xlsx");
    let (mut wb, id) = book();
    wb.set_cell_value_tracked(1, 1, 0, "=ROW(Sheet1!B4)");
    xlsx::export(&wb, &plain, None).unwrap();
    set_saved_sort(&mut wb, id, 0, false, true); // Qty is already ascending.
    xlsx::export(&wb, &sorted, None).unwrap();
    assert_eq!(
        xml(&plain, "xl/worksheets/sheet1.xml"),
        xml(&sorted, "xl/worksheets/sheet1.xml")
    );
    assert_eq!(
        xml(&plain, "xl/worksheets/sheet2.xml"),
        xml(&sorted, "xl/worksheets/sheet2.xml")
    );
}

#[test]
fn stored_order_export_keeps_cells_formulas_comments_filters_and_saved_sort_metadata() {
    use xlsx::ExportOrder;
    let dir = tempfile::tempdir().unwrap();
    let plain = dir.path().join("plain.xlsx");
    let stored = dir.path().join("stored.xlsx");
    let buffered = dir.path().join("buffer.xlsx");
    for buttons in [false, true] {
        let (mut wb, id) = book();
        select_values(&mut wb, id, 0, &[3, 7], buttons);
        xlsx::export(&wb, &plain, None).unwrap();
        set_saved_sort(&mut wb, id, 0, true, buttons);
        let before = authored_snapshot(&wb);
        let report = xlsx::export_with_order(&wb, &stored, None, ExportOrder::Stored).unwrap();
        assert_eq!(
            report.warnings,
            xlsx::table_export_warnings_with_order(&wb, None, ExportOrder::Stored).unwrap()
        );
        assert!(report
            .warnings
            .iter()
            .any(|w| w.contains("use Reapply") && w.contains("stored coordinates")));
        let (bytes, other) =
            xlsx::export_to_buffer_with_order(&wb, None, ExportOrder::Stored).unwrap();
        std::fs::write(&buffered, bytes).unwrap();
        assert_eq!(report.warnings, other.warnings);
        for part in [
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
            "xl/comments1.xml",
        ] {
            assert_eq!(xml(&plain, part), xml(&stored, part));
            assert_eq!(xml(&stored, part), xml(&buffered, part));
        }
        assert!(xml(&stored, "xl/tables/table1.xml").contains("<sortState"));
        let (loaded, report) = xlsx::import(&stored).unwrap();
        assert!(report.warnings.is_empty(), "{:?}", report.warnings);
        assert_eq!(
            loaded.sheet(0).unwrap().table_view_spec(),
            wb.sheet(0).unwrap().table_view_spec()
        );
        for r in 3..8 {
            assert_eq!(loaded.sheet(0).unwrap().get_raw(r, 1), (r - 1).to_string());
        }
        assert_eq!(
            loaded.sheet(0).unwrap().comment(4, 3).unwrap().text,
            "Manual override"
        );
        assert_eq!(authored_snapshot(&wb), before);
    }
}

#[test]
fn stored_order_is_explicit_fallback_for_unsupported_materialization_with_loss_warnings() {
    use visigrid_engine::{
        cell::CellStyle,
        cond_format::CondStyle,
        validation::{CellRange, ValidationRule},
    };
    use xlsx::ExportOrder;
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("stored.xlsx");
    for mode in 0..4 {
        let (mut wb, id) = book();
        match mode {
            0 => {
                wb.set_cell_value_tracked(1, 1, 0, "=ROW(Sheet1!B4)");
            }
            1 => {
                wb.sheet_mut(0).unwrap().cond_formats.add(
                    vec![CellRange::single(3, 1)],
                    "=B4>0",
                    CondStyle::Named(CellStyle::Success),
                );
            }
            2 => {
                wb.sheet_mut(0).unwrap().set_cell_validation(
                    3,
                    1,
                    ValidationRule::list_inline(vec!["2".into()]),
                );
            }
            _ => {
                wb.define_name_for_cell("FirstQty", 0, 3, 1).unwrap();
                wb.set_cell_value_tracked(1, 1, 0, "=FirstQty");
            }
        }
        set_saved_sort(&mut wb, id, 0, true, true);
        let before = authored_snapshot(&wb);
        if mode == 3 {
            xlsx::export(&wb, &file, None).unwrap();
        } else {
            assert!(xlsx::export(&wb, &file, None).is_err());
        }
        let report = xlsx::export_with_order(&wb, &file, None, ExportOrder::Stored).unwrap();
        if mode == 1 {
            assert!(report
                .warnings
                .iter()
                .any(|w| w.contains("Conditional formatting")));
        }
        if mode == 3 {
            assert!(!report.warnings.iter().any(|w| w.contains("Named-range definitions")));
        }
        let (loaded, _) = xlsx::import(&file).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 1), "2");
        if mode == 0 {
            assert_eq!(loaded.sheet(1).unwrap().get_raw(1, 0), "=ROW(Sheet1!B4)");
        }
        if mode == 2 {
            assert!(loaded.sheet(0).unwrap().has_validation(3, 1));
        }
        if mode == 3 {
            assert_eq!(loaded.sheet(1).unwrap().get_raw(1, 0), "=FirstQty");
        }
        assert_eq!(authored_snapshot(&wb), before);
    }
}

#[test]
fn stored_order_preserves_host_layout_that_cannot_move_with_a_sort() {
    use xlsx::ExportOrder;
    let (mut wb, id) = book();
    set_saved_sort(&mut wb, id, 0, true, true);
    let mut layout = xlsx::ExportLayout::default();
    layout.row_heights.insert(4, 30.0);
    layout.hidden_rows.push(4);
    layout.frozen_rows = 4;
    let layouts = [layout];
    assert!(
        xlsx::table_export_warnings_with_order(&wb, Some(&layouts), ExportOrder::Sorted).is_err()
    );
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("layout.xlsx");
    xlsx::export_with_order(&wb, &file, Some(&layouts), ExportOrder::Stored).unwrap();
    let (loaded, report) = xlsx::import(&file).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 1), "2");
    // ExportLayout uses pixels; imported OOXML row heights are points.
    assert_eq!(report.imported_layouts[0].row_heights.get(&4), Some(&22.5));
    assert!(report.imported_layouts[0].hidden_rows.contains(&4));
    assert_eq!(report.imported_layouts[0].frozen_rows, 4);
    assert!(report
        .warnings
        .iter()
        .any(|w| w.contains("saved sort/filter/button settings were not imported")));
}

#[test]
fn stored_order_does_not_bypass_table_writer_refusals() {
    use xlsx::ExportOrder;
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("keep.xlsx");
    std::fs::write(&file, b"untouched").unwrap();
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "Key");
    wb.create_table(
        SheetId(1),
        TableRange {
            start_row: 0,
            end_row: 0,
            start_col: 0,
            end_col: 0,
        },
        "EmptyTable",
    )
    .unwrap();
    assert!(
        xlsx::table_export_warnings_with_order(&wb, None, ExportOrder::Stored)
            .unwrap_err()
            .contains("only headers")
    );
    assert!(xlsx::export_with_order(&wb, &file, None, ExportOrder::Stored).is_err());
    assert!(xlsx::export_to_buffer_with_order(&wb, None, ExportOrder::Stored).is_err());
    assert_eq!(std::fs::read(&file).unwrap(), b"untouched");
}

#[test]
fn headless_fallback_preserves_stored_records_and_reports_unsupported_metadata() {
    use visigrid_engine::{
        cell::CellStyle,
        cond_format::CondStyle,
        validation::{CellRange, ValidationRule},
    };
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("fallback.xlsx");
    for validation in [true, false] {
        let (mut wb, id) = book();
        if validation {
            wb.sheet_mut(0).unwrap().set_cell_validation(
                3,
                1,
                ValidationRule::list_inline(vec!["2".into()]),
            );
        } else {
            wb.sheet_mut(0).unwrap().cond_formats.add(
                vec![CellRange::single(3, 1)],
                "=B4>0",
                CondStyle::Named(CellStyle::Success),
            );
        }
        set_saved_sort(&mut wb, id, 0, true, true);
        let before = authored_snapshot(&wb);
        assert!(xlsx::export_to_buffer(&wb, None).is_err());
        let (bytes, report) = xlsx::export_to_buffer_with_stored_fallback(&wb, None).unwrap();
        assert!(report.warnings[0].contains("Exported in stored order instead"));
        assert!(report.warnings[0].contains(if validation {
            "Validation"
        } else {
            "conditional-format"
        }));
        if !validation {
            assert!(report
                .warnings
                .iter()
                .any(|w| w.contains("semantic styles")));
        }
        std::fs::write(&file, bytes).unwrap();
        let (loaded, _) = xlsx::import(&file).unwrap();
        for row in 3..8 {
            assert_eq!(
                loaded.sheet(0).unwrap().get_raw(row, 1),
                wb.sheet(0).unwrap().get_raw(row, 1)
            );
            assert_eq!(
                loaded.sheet(0).unwrap().get_raw(row, 3),
                wb.sheet(0)
                    .unwrap()
                    .get_raw(row, 3)
                    .replace("[@Qty]", "[[#This Row],[Qty]]")
            );
            assert_eq!(
                loaded.sheet(0).unwrap().get_display(row, 3),
                wb.sheet(0).unwrap().get_display(row, 3)
            );
        }
        assert!(loaded
            .sheet(0)
            .unwrap()
            .table_view_spec()
            .unwrap()
            .sort
            .is_some());
        if validation {
            assert!(loaded.sheet(0).unwrap().has_validation(3, 1));
        }
        assert_eq!(authored_snapshot(&wb), before);
    }
}

fn external_totals_file(path: &Path, hidden: bool) {
    use rust_xlsxwriter::{Formula, Table, TableColumn, TableFunction, Workbook as Excel};
    let mut wb = Excel::new();
    let sheet = wb.add_worksheet();
    sheet.write_string(1, 0, "West").unwrap();
    sheet.write_number(1, 1, 10).unwrap();
    sheet.write_number(1, 2, 2).unwrap();
    sheet.write_string(2, 0, "East").unwrap();
    sheet.write_number(2, 1, 20).unwrap();
    sheet.write_number(2, 2, 3).unwrap();
    if hidden {
        sheet.set_row_hidden(2).unwrap();
    }
    sheet
        .add_table(
            0,
            0,
            3,
            2,
            &Table::new()
                .set_name("Sales")
                .set_total_row(true)
                .set_columns(&[
                    TableColumn::new()
                        .set_header("Region")
                        .set_total_label("Grand & total"),
                    TableColumn::new()
                        .set_header("Amount")
                        .set_total_function(TableFunction::Sum),
                    TableColumn::new()
                        .set_header("Qty")
                        .set_total_function(TableFunction::Custom(Formula::new("SUM([Qty])*2"))),
                ]),
        )
        .unwrap();
    sheet
        .write_formula(0, 5, Formula::new("Sales[[#Totals],[Amount]]"))
        .unwrap();
    sheet
        .write_formula(0, 6, Formula::new("SUM(Sales[Amount])"))
        .unwrap();
    wb.save(path).unwrap();
}

#[test]
fn excel_totals_roundtrip_preserves_body_footer_formulas_metadata_and_native_version() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("external.xlsx");
    let native_path = dir.path().join("totals.sheet");
    external_totals_file(&input, false);
    let (mut wb, report) = xlsx::import(&input).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(report.tables_skipped, 0);
    assert_eq!(wb.saved_tables().version, 5);
    let table = wb.tables().next().unwrap().1;
    assert_eq!(table.range.data_rows(), 2);
    assert_eq!(table.totals_row(), Some(3));
    assert_eq!(
        table.totals.as_ref().unwrap().columns[1]
            .function
            .as_deref(),
        Some("sum")
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "30");
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 2), "10");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 5), "30");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 6), "30");
    // Body editing recalculates both totals and dependent #Totals formulas.
    wb.set_cell_value_tracked(0, 1, 1, "15");
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "35");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 5), "35");
    native::save_workbook(&wb, &native_path).unwrap();
    wb = native::load_workbook(&native_path).unwrap();
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let (loaded, _, _) = visigrid_io::json::import_any(&json).unwrap();
    assert_eq!(loaded.saved_tables().version, 5);
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 1), "35");
    assert_eq!(loaded.sheet(0).unwrap().get_display(0, 5), "35");
    wb = loaded;
    for _ in 0..2 {
        let output = dir.path().join("output.xlsx");
        xlsx::export_with_order(&wb, &output, None, xlsx::ExportOrder::Stored).unwrap();
        let metadata = xml(&output, "xl/tables/table1.xml");
        assert!(metadata.contains("ref=\"A1:C4\""), "{metadata}");
        assert!(metadata.contains("totalsRowCount=\"1\""), "{metadata}");
        assert!(metadata.contains("totalsRowFunction=\"sum\""), "{metadata}");
        assert!(
            metadata.contains("totalsRowLabel=\"Grand &amp; total\""),
            "{metadata}"
        );
        assert!(
            metadata.contains("<totalsRowFormula>SUM([Qty])*2</totalsRowFormula>"),
            "{metadata}"
        );
        assert!(metadata.contains("<autoFilter ref=\"A1:C3\""), "{metadata}");
        let (loaded, report) = xlsx::import(&output).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        assert_eq!(loaded.sheet(0).unwrap().get_display(3, 1), "35");
        assert_eq!(loaded.sheet(0).unwrap().get_display(0, 5), "35");
        assert_eq!(loaded.sheet(0).unwrap().get_display(3, 2), "10");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 0), "Grand & total");
        wb = loaded;
    }
}

#[test]
fn imported_totals_remain_visible_and_recalculate_when_filter_fields_change() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("external.xlsx");
    external_totals_file(&input, false);
    let (mut wb, _) = xlsx::import(&input).unwrap();
    let (sid, table) = wb.tables().next().unwrap();
    let id = table.id;
    let column = table.columns[0].id;
    let mut spec = visigrid_engine::table_view::TableViewSpec::new(id);
    spec.filters.push(visigrid_engine::table_view::TableFilter {
        column,
        criteria: visigrid_engine::filter::ColumnFilter {
            selected: Some(
                [visigrid_engine::filter::NormalizedFilterKey::Text(
                    "west".into(),
                )]
                .into(),
            ),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(sid, Some(spec)).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "10");
    let view = wb
        .sheet(0)
        .unwrap()
        .build_saved_table_view(5)
        .unwrap()
        .unwrap();
    assert!(
        view.rows().data_to_view(3).is_some(),
        "footer stays visible"
    );
    let filtered_path = dir.path().join("filtered-totals.xlsx");
    xlsx::export_with_order(&wb, &filtered_path, None, xlsx::ExportOrder::Stored).unwrap();
    let (roundtrip, report) = xlsx::import(&filtered_path).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(roundtrip.sheet(0).unwrap().get_display(3, 1), "10");
    assert!(
        roundtrip
            .tables()
            .next()
            .unwrap()
            .1
            .totals
            .as_ref()
            .unwrap()
            .hidden_rows
            .is_empty(),
        "filter masks must not become manual hiding"
    );
    // Changing a different column still invalidates the subtotal.
    wb.set_cell_value_tracked(0, 2, 0, "West");
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "30");
    wb.set_table_view_spec(sid, None).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "30");
    let output = dir.path().join("filtered.xlsx");
    xlsx::export_with_order(&wb, &output, None, xlsx::ExportOrder::Stored).unwrap();
}

#[test]
fn totals_keep_manual_hidden_rows_and_refuse_unsafe_authoring() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("hidden.xlsx");
    external_totals_file(&input, true);
    let (mut wb, report) = xlsx::import(&input).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "10");
    let native_path = dir.path().join("hidden.sheet");
    native::save_workbook(&wb, &native_path).unwrap();
    assert!(native::load_layout(&native_path).hidden_rows[&0].contains(&2));
    wb = native::load_workbook(&native_path).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "10");
    let (sid, t) = wb.tables().next().unwrap();
    let id = t.id;
    let range = t.range;
    let resize = wb.resize_table(id, TableRange { end_row: 4, ..range }).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(5, 1), "10");
    wb.apply_table_commit(&resize, true).unwrap();
    let append = wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "10");
    wb.apply_table_commit(&append, true).unwrap();
    let converted = wb.remove_table(id).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 1), "10");
    assert!(wb.sheet(0).unwrap().manual_hidden_rows().contains(&2));
    wb.apply_table_commit(&converted, true).unwrap();
    assert!(wb
        .sheet(0)
        .unwrap()
        .table_structural_error(true, 2, 1, false)
        .is_none());
    assert!(wb.sheet(0).unwrap().table_value_write_error(3, 1).is_some());
    assert!(wb
        .create_table(
            sid,
            TableRange {
                start_row: 3,
                end_row: 4,
                start_col: 0,
                end_col: 2
            },
            "Overlap"
        )
        .is_err());
    let output = dir.path().join("output.xlsx");
    xlsx::export_with_order(&wb, &output, None, xlsx::ExportOrder::Stored).unwrap();
    let (loaded, report) = xlsx::import(&output).unwrap();
    assert!(report.imported_layouts[0].hidden_rows.contains(&2));
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 1), "10");
}

#[test]
fn empty_body_and_dormant_totals_metadata_roundtrip() {
    use rust_xlsxwriter::{Table, TableColumn as Column, TableFunction};
    let dir = tempfile::tempdir().unwrap();
    for empty in [true, false] {
        let changed = dir.path().join("external.xlsx");
        let mut external = rust_xlsxwriter::Workbook::new();
        let sheet = external.add_worksheet();
        if !empty {
            sheet.write_string(1, 0, "West").unwrap();
            sheet.write_number(1, 1, 10).unwrap();
            sheet.write_number(1, 2, 2).unwrap();
            sheet.write_string(2, 0, "East").unwrap();
            sheet.write_number(2, 1, 20).unwrap();
            sheet.write_number(2, 2, 3).unwrap();
        }
        sheet
            .add_table(
                0,
                0,
                if empty { 1 } else { 2 },
                2,
                &Table::new()
                    .set_name("Sales")
                    .set_total_row(empty)
                    .set_columns(&[
                        Column::new().set_header("Region").set_total_label("Total"),
                        Column::new()
                            .set_header("Amount")
                            .set_total_function(TableFunction::Sum),
                        Column::new()
                            .set_header("Qty")
                            .set_total_function(TableFunction::Custom("SUM([Qty])*2".into())),
                    ]),
            )
            .unwrap();
        external.save(&changed).unwrap();
        let (wb, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let t = wb.tables().next().unwrap().1;
        assert!(t.totals.is_some(), "dormant function settings must survive");
        assert_eq!(t.totals_row(), empty.then_some(1));
        if empty {
            assert_eq!(t.range.data_rows(), 0);
            assert_eq!(wb.sheet(0).unwrap().get_display(1, 1), "0");
            assert_eq!(wb.sheet(0).unwrap().get_display(1, 2), "0");
        }
        let output = dir.path().join("out.xlsx");
        let report = xlsx::export_with_order(&wb, &output, None, xlsx::ExportOrder::Stored).unwrap();
        let metadata = xml(&output, "xl/tables/table1.xml");
        if empty {
            let warning = "Table Sales has no records; its totals row is exported as values.";
            assert!(report.warnings.iter().any(|w| w == warning), "{:?}", report.warnings);
            assert!(!metadata.contains("totalsRow"), "{metadata}");
            // The filter covers the footer row exported as data, as the Table does.
            assert!(metadata.contains("<autoFilter ref=\"A1:C2\""), "{metadata}");
            assert!(!report.warnings.iter().any(|w| w.contains("filter range")), "{:?}", report.warnings);
            let sheet = xml(&output, "xl/worksheets/sheet1.xml");
            assert!(!sheet.contains("<f"), "{sheet}");
            let (loaded, report) = xlsx::import(&output).unwrap();
            assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
            assert!(!report.warnings.iter().any(|w| w.contains("filter range")), "{:?}", report.warnings);
            assert!(loaded.tables().next().unwrap().1.totals_row().is_none());
            assert_eq!(loaded.sheet(0).unwrap().get_display(1, 0), "Total");
            assert_eq!(loaded.sheet(0).unwrap().get_display(1, 1), "0");
            assert_eq!(loaded.sheet(0).unwrap().get_display(1, 2), "0");
            assert!(!loaded.sheet(0).unwrap().get_raw(1, 1).starts_with('='));
        } else {
            assert!(metadata.contains("totalsRowFunction"), "{metadata}");
            let (loaded, report) = xlsx::import(&output).unwrap();
            assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
            assert_eq!(loaded.tables().next().unwrap().1.totals, t.totals);
        }
    }
}

#[test]
fn totals_sort_export_uses_explicit_stored_order_or_headless_fallback() {
    use visigrid_engine::{
        filter::SortDirection,
        table_view::{TableSort, TableViewSpec},
    };
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("external.xlsx");
    external_totals_file(&input, false);
    let (mut wb, _) = xlsx::import(&input).unwrap();
    let (sid, t) = wb.tables().next().unwrap();
    let mut spec = TableViewSpec::new(t.id);
    spec.sort = Some(TableSort {
        column: t.columns[1].id,
        direction: SortDirection::Descending,
    });
    wb.set_table_view_spec(sid, Some(spec)).unwrap();
    let output = dir.path().join("out.xlsx");
    assert!(xlsx::export_with_order(&wb, &output, None, xlsx::ExportOrder::Sorted).is_err());
    assert!(!output.exists());
    let (bytes, report) = xlsx::export_to_buffer_with_stored_fallback(&wb, None).unwrap();
    std::fs::write(&output, bytes).unwrap();
    assert!(
        report.warnings.iter().any(|w| w.contains("stored")),
        "{:?}",
        report.warnings
    );
    let (loaded, report) = xlsx::import(&output).unwrap();
    assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 1), "10");
    assert_eq!(loaded.sheet(0).unwrap().get_display(3, 1), "30");
    assert_eq!(
        loaded.sheet(0).unwrap().table_view_spec().unwrap().sort,
        wb.sheet(0).unwrap().table_view_spec().unwrap().sort
    );
}

#[test]
fn renamed_totals_keep_custom_settings_and_calculated_rules_through_native_and_xlsx() {
    use visigrid_engine::table::TableTotal;
    for visible in [true, false] {
        let dir = tempfile::tempdir().unwrap();
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 2, TableTotal { function: Some("custom".into()),
            formula: Some("=SUM([Price])+SUM(Sales[Price])".into()), label: None }).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        wb.rename_table(id, "Orders").unwrap();
        wb.rename_table_columns(id, &["Units".into(), "Cost".into(), "Revenue".into()]).unwrap();
        let path = dir.path().join("rename.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        let path = dir.path().join("rename.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let xml = xml(&path, "xl/tables/table1.xml");
        assert!(xml.contains("name=\"Orders\""), "{xml}");
        assert!(xml.contains("SUM([Cost])+SUM(Orders[Cost])"), "{xml}");
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let id = loaded.tables().next().unwrap().1.id;
        if !visible { loaded.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 2), "100");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(8, 2), "=SUM([Cost])+SUM(Orders[Cost])");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
        assert!(loaded.table(id).unwrap().1.columns[2].formula.as_ref().unwrap().contains("Units"));
        loaded.set_table_totals_visible(id, false, Default::default()).unwrap();
        loaded.set_table_totals_visible(id, true, Default::default()).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 2), "100");
    }
}

#[test]
fn edited_calculated_rules_with_totals_roundtrip_and_fill_after_append() {
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        wb.set_calculated_column(id, 3, 3, "=[@Qty]*[@Price]*2", false).unwrap();
        let rule = wb.table(id).unwrap().1.columns[2].formula.clone();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("edited.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        assert_eq!(wb.table(id).unwrap().1.columns[2].formula, rule);
        let path = dir.path().join("edited.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let id = loaded.tables().next().unwrap().1.id;
        assert_eq!(loaded.table(id).unwrap().1.columns[2].formula.as_deref(), Some("=[[#This Row],[Qty]]*[[#This Row],[Price]]*2"));
        assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(5, 3), "");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(6, 3), "=1+2");
        if !visible { loaded.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        let total = loaded.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap();
        loaded.append_table_rows(id, 1, &[(8, 1, "5".into()), (8, 2, "10".into())]).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 3), "100");
        assert_eq!(loaded.sheet(0).unwrap().get_display(9, 3).parse::<f64>().unwrap(), total + 100.0);
    }
}

#[test]
fn totals_column_structure_roundtrips_through_native_and_excel() {
    use visigrid_engine::{structural::Axis, table::TableTotal};
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 3, TableTotal {
            function: Some("custom".into()), formula: Some("=SUM([Qty])+C4".into()), label: None,
        }).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        wb.structural_edit(0, Axis::Col, 2, 1, false).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("columns.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        let path = dir.path().join("columns.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let id = loaded.tables().next().unwrap().1.id;
        let table = loaded.table(id).unwrap().1;
        assert_eq!(table.columns.len(), 4);
        assert_eq!(table.totals.as_ref().unwrap().columns[3].formula.as_deref(), Some("=SUM([Qty])+D4"));
        assert_eq!(table.totals.as_ref().unwrap().visible, visible);
        if !visible { loaded.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        assert_eq!(loaded.sheet(0).unwrap().get_raw(8, 4), "=SUM([Qty])+D4");
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 4), "30");
        loaded.structural_edit(0, Axis::Col, 2, 1, true).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_raw(8, 3), "=SUM([Qty])+C4");
    }
}


#[test]
fn totals_worksheet_row_edits_keep_native_and_excel_metadata_aligned() {
    use visigrid_engine::{structural::Axis, table::TableTotal};
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 3, TableTotal {
            function: Some("custom".into()), formula: Some("=SUM([Qty])+C4".into()), label: None,
        }).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        wb.structural_edit(0, Axis::Row, 0, 2, false).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("rows.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        let path = dir.path().join("rows.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let id = loaded.tables().next().unwrap().1.id;
        let table = loaded.table(id).unwrap().1;
        assert_eq!(table.range.start_row, 4);
        assert_eq!(table.range.end_row, 9);
        assert_eq!(table.totals.as_ref().unwrap().columns[2].formula.as_deref(), Some("=SUM([Qty])+C6"));
        assert_eq!(table.totals.as_ref().unwrap().visible, visible);
        if !visible { loaded.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        assert_eq!(loaded.sheet(0).unwrap().get_display(10, 3), "30");
        loaded.structural_edit(0, Axis::Row, 10, 1, false).unwrap();
        assert_eq!(loaded.table(id).unwrap().1.totals_row(), Some(11));
        assert_eq!(loaded.sheet(0).unwrap().get_display(11, 3), "30");
        assert!(loaded.sheet(0).unwrap().get_raw(10, 3).starts_with('='));
    }
}


#[test]
fn horizontally_resized_totals_survive_native_and_xlsx_roundtrips() {
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        wb.set_cell_value_tracked(0, 2, 4, "Extra");
        wb.set_cell_value_tracked(0, 3, 4, "17");
        let original = wb.table(id).unwrap().1.range;
        wb.resize_table(id, TableRange { end_col: 4, ..original }).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("width.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        let path = dir.path().join("width.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        let id = loaded.tables().next().unwrap().1.id;
        assert_eq!(loaded.table(id).unwrap().1.columns.len(), 4);
        let totals = loaded.table(id).unwrap().1.totals.as_ref().unwrap();
        assert_eq!(totals.visible, visible);
        assert_eq!(totals.columns.len(), 4);
        assert_eq!(totals.columns[2].function.as_deref(), Some("sum"));
        assert_eq!(totals.columns[3].function.as_deref().unwrap_or("none"), "none");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 4), "17");
        if !visible { loaded.set_table_totals_visible(id, true, Default::default()).unwrap(); }
        loaded.resize_table(id, original).unwrap();
        assert_eq!(loaded.table(id).unwrap().1.totals.as_ref().unwrap().columns.len(), 3);
        assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 4), "17");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(8, 4), "");
    }
}

#[test]
fn copied_totals_keep_independent_bindings_through_native_and_excel() {
    use visigrid_engine::table::TableTotal;
    let dir = tempfile::tempdir().unwrap();
    for visible in [true, false] {
        let (mut wb, id) = book();
        wb.set_table_totals_visible(id, true, Default::default()).unwrap();
        wb.set_table_total(id, 3, TableTotal {
            function: Some("custom".into()), formula: Some("=SUM(Sales[Amount])+SUM([Price])".into()), label: None,
        }).unwrap();
        let original = wb.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap();
        if !visible { wb.set_table_totals_visible(id, false, Default::default()).unwrap(); }
        let (wb, index) = wb.prepare_sheet_copy(&wb, SheetId(1), "Copied data").unwrap();
        let path = dir.path().join("copied-totals.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let wb = native::load_workbook(&path).unwrap();
        let path = dir.path().join("copied-totals.xlsx");
        xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
        let (mut loaded, report) = xlsx::import(&path).unwrap();
        assert!(report.warnings.is_empty(), "{:?}", report.warnings);
        let table = loaded.sheet(index).unwrap().tables()[0].clone();
        assert_eq!(table.name, "Sales_Copy");
        assert_ne!(table.id, loaded.sheet(0).unwrap().tables()[0].id);
        assert_eq!(table.totals.as_ref().unwrap().visible, visible);
        assert_eq!(table.totals.as_ref().unwrap().columns[2].formula.as_deref(), Some("=SUM(Sales_Copy[Amount])+SUM([Price])"));
        assert_eq!(loaded.sheet(index).unwrap().comment(4, 3).unwrap().text, "Manual override");
        if !visible {
            loaded.set_table_totals_visible(table.id, true, Default::default()).unwrap();
            let original_id = loaded.sheet(0).unwrap().tables()[0].id;
            loaded.set_table_totals_visible(original_id, true, Default::default()).unwrap();
        }
        assert_eq!(loaded.sheet(index).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original);
        loaded.set_cell_value_tracked(index, 4, 3, "999");
        assert_eq!(loaded.sheet(index).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original + 222.0);
        assert_eq!(loaded.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original);
    }
}

#[test]
fn large_plain_sheets_do_not_enter_table_parser_or_hide_later_tables() {
    let dir = tempfile::tempdir().unwrap();
    for unrelated_links in [false, true] {
        let path = dir.path().join("large-plain-source.xlsx");
        let changed = dir.path().join("large-plain.xlsx");
        let mut file = rust_xlsxwriter::Workbook::new();
        let plain = file.add_worksheet();
        plain.set_name("Large plain").unwrap();
        plain.write_string(0, 0, "Keep ordinary cells").unwrap();
        if unrelated_links { plain.write_url(2, 0, "https://example.com").unwrap(); }
        let sheet = file.add_worksheet();
        sheet.set_name("With Tables").unwrap();
        sheet.write_number(1, 0, 17).unwrap();
        sheet.add_table(0, 0, 1, 0, &rust_xlsxwriter::Table::new().set_name("LaterData")).unwrap();
        file.save(&path).unwrap();
        rewrite(&path, &changed, |name, data| {
            (name.into(), if name == "xl/worksheets/sheet1.xml" {
                data.replace("</worksheet>", &format!("<!--{}--></worksheet>", " ".repeat(33 * 1024 * 1024)))
            } else { data })
        });
        let (loaded, report) = xlsx::import(&changed).unwrap();
        assert_eq!(report.tables_imported, 1, "{:?}", report.warnings);
        assert_eq!(report.tables_skipped, 0);
        assert!(report.warnings.iter().all(|w| !w.contains("Table metadata")), "{:?}", report.warnings);
        assert_eq!(loaded.sheet(0).unwrap().get_raw(0, 0), "Keep ordinary cells");
        assert_eq!(loaded.sheet(1).unwrap().get_raw(1, 0), "17");
        assert!(loaded.table_by_name("LaterData").is_some());
    }
}

#[test]
fn table_import_count_budget_is_global_and_keeps_cells_and_valid_definitions() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("many-tables.xlsx");
    let mut file = rust_xlsxwriter::Workbook::new();
    for (start, count) in [(0, 513), (513, 512)] {
        let sheet = file.add_worksheet();
        for i in 0..count {
            let row = i * 3;
            sheet.write_number(row + 1, 0, (start + i) as f64).unwrap();
            sheet.add_table(row, 0, row + 1, 0, &rust_xlsxwriter::Table::new().set_name(format!("ImportedData{}", start + i))).unwrap();
        }
    }
    file.save(&path).unwrap();
    for values_only in [false, true] {
        let (loaded, report) = xlsx::import_with_options(&path, &xlsx::ImportOptions { values_only, ..Default::default() }).unwrap();
        assert_eq!(report.tables_imported, 1024, "{:?}", report.warnings);
        assert_eq!(report.tables_skipped, 1);
        assert_eq!(loaded.tables().count(), 1024);
        assert_eq!(loaded.sheet(1).unwrap().get_raw(511 * 3 + 1, 0), "1024");
        assert!(loaded.table_by_name("ImportedData1023").is_some());
        assert!(loaded.table_by_name("ImportedData1024").is_none());
        assert!(report.warnings.iter().any(|w| w.contains("workbook limit of 1024")));
        let path = dir.path().join(format!("many-{values_only}.sheet"));
        native::save_workbook(&loaded, &path).unwrap();
        assert_eq!(native::load_workbook(&path).unwrap().tables().count(), 1024);
    }
    let damaged = dir.path().join("many-damaged.xlsx");
    rewrite(&path, &damaged, |name, data| {
        (name.into(), if name == "xl/tables/table1.xml" { data.replace("name=\"Column1\"", "name=\"Wrong header\"") } else { data })
    });
    let (loaded, report) = xlsx::import(&damaged).unwrap();
    assert_eq!(report.tables_imported, 1023);
    assert_eq!(report.tables_skipped, 2); // A corrupt definition still consumes an attempt.
    assert!(loaded.table_by_name("ImportedData1023").is_some());
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "0");
}

#[test]
fn table_metadata_byte_budget_skips_large_parts_but_keeps_later_small_tables() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("metadata-source.xlsx");
    let changed = dir.path().join("metadata-budget.xlsx");
    let mut file = rust_xlsxwriter::Workbook::new();
    let sheet = file.add_worksheet();
    for i in 0..4 {
        sheet.write_number(i * 3 + 1, 0, (i + 1) as f64).unwrap();
        sheet.add_table(i * 3, 0, i * 3 + 1, 0, &rust_xlsxwriter::Table::new().set_name(format!("BudgetData{i}"))).unwrap();
    }
    file.save(&path).unwrap();
    rewrite(&path, &changed, |name, data| {
        (name.into(), if ["xl/tables/table1.xml", "xl/tables/table2.xml", "xl/tables/table3.xml"].contains(&name) {
            data.replace("</table>", &format!("<!--{}--></table>", " ".repeat(24 * 1024 * 1024)))
        } else { data })
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 3, "{:?}", report.warnings);
    assert_eq!(report.tables_skipped, 1);
    assert!(report.warnings.iter().any(|w| w.contains("64 MiB Table metadata budget")));
    assert!(loaded.table_by_name("BudgetData2").is_none());
    assert!(loaded.table_by_name("BudgetData3").is_some());
    assert_eq!(loaded.active_sheet().get_raw(7, 0), "3");

    // Lie about Table uncompressed sizes while retaining the compressed bytes
    // and CRCs. Limits must charge decompression, not just ZIP declarations.
    let mut bytes = std::fs::read(&changed).unwrap();
    let end = bytes.windows(4).rposition(|b| b == b"PK\x05\x06").unwrap();
    let mut at = u32::from_le_bytes(bytes[end + 16..end + 20].try_into().unwrap()) as usize;
    while bytes.get(at..at + 4) == Some(b"PK\x01\x02") {
        let name_len = u16::from_le_bytes(bytes[at + 28..at + 30].try_into().unwrap()) as usize;
        let extra_len = u16::from_le_bytes(bytes[at + 30..at + 32].try_into().unwrap()) as usize;
        let comment_len = u16::from_le_bytes(bytes[at + 32..at + 34].try_into().unwrap()) as usize;
        if bytes[at + 46..at + 46 + name_len].starts_with(b"xl/tables/table") {
            let local = u32::from_le_bytes(bytes[at + 42..at + 46].try_into().unwrap()) as usize;
            bytes[at + 24..at + 28].copy_from_slice(&1u32.to_le_bytes());
            bytes[local + 22..local + 26].copy_from_slice(&1u32.to_le_bytes());
        }
        at += 46 + name_len + extra_len + comment_len;
    }
    let forged = dir.path().join("forged-metadata-sizes.xlsx");
    std::fs::write(&forged, bytes).unwrap();
    let (loaded, report) = xlsx::import(&forged).unwrap();
    // A ZIP reader may reject a size mismatch itself. Either layer must keep
    // these declarations from admitting all four Tables past the byte budget.
    assert!(report.tables_imported <= 2, "{:?}", report.warnings);
    assert_eq!(report.tables_skipped, 4 - report.tables_imported);
    assert!(!report.warnings.is_empty());
    assert_eq!(loaded.active_sheet().get_raw(10, 0), "4");
}

#[test]
fn unreadable_table_sheet_does_not_discard_valid_tables_on_later_sheets() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("two-table-sheets.xlsx");
    let changed = dir.path().join("malformed-table-sheet.xlsx");
    let mut file = rust_xlsxwriter::Workbook::new();
    for name in ["MalformedData", "LaterValidData"] {
        let sheet = file.add_worksheet();
        sheet.set_name(name).unwrap();
        sheet.write_number(1, 0, 42).unwrap();
        sheet.add_table(0, 0, 1, 0, &rust_xlsxwriter::Table::new().set_name(name)).unwrap();
    }
    file.save(&path).unwrap();
    rewrite(&path, &changed, |name, data| {
        (name.into(), if name == "xl/worksheets/sheet1.xml" {
            // Large worksheets are now valid; exercise failure isolation with
            // an actually unreadable Table link, while keeping the cell XML.
            assert!(data.contains("<tablePart r:id="));
            data.replace("<tablePart r:id=", "<tablePart missingId=")
        } else { data })
    });
    let (loaded, report) = xlsx::import(&changed).unwrap();
    assert_eq!(report.tables_imported, 1);
    assert_eq!(report.tables_skipped, 1);
    assert!(loaded.table_by_name("LaterValidData").is_some());
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "42");
    assert!(report.warnings.iter().any(|w| w.contains("on MalformedData") && w.contains("Missing Table relationship")));
}

#[test]
fn changed_manual_visibility_recalculates_and_roundtrips_native_and_excel() {
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let original = wb.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap();
    let (hidden, _) = wb.prepare_table_row_visibility(SheetId(1), [4].into_iter().collect()).unwrap();
    assert_eq!(hidden.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original - 777.0);
    let path = dir.path().join("manual-visibility.sheet");
    native::save_workbook(&hidden, &path).unwrap();
    let hidden = native::load_workbook(&path).unwrap();
    let path = dir.path().join("manual-visibility.xlsx");
    xlsx::export_with_order(&hidden, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(loaded.sheet(0).unwrap().get_raw(4, 3), "777");
    assert_eq!(loaded.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original - 777.0);
    assert!(loaded.sheet(0).unwrap().tables()[0].totals.as_ref().unwrap().hidden_rows.contains(&4));
    let (shown, _) = loaded.prepare_table_row_visibility(loaded.sheet(0).unwrap().id, Default::default()).unwrap();
    assert_eq!(shown.sheet(0).unwrap().get_display(8, 3).parse::<f64>().unwrap(), original);
}

#[test]
fn relocated_footer_links_and_names_survive_native_and_stored_excel_roundtrips() {
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let name = wb.sheet(0).unwrap().name.clone();
    wb.set_cell_value_tracked(1, 1, 0, &format!("='{name}'!$D$9"));
    wb.define_name_for_cell("SalesTotal", 0, 8, 3).unwrap();
    wb.set_cell_value_tracked(1, 2, 0, "=SalesTotal");
    wb.append_table_rows(id, 1, &[(8, 1, "2".into()), (8, 2, "10".into())]).unwrap();
    let expected = wb.sheet(0).unwrap().get_display(9, 3);
    assert_eq!(wb.sheet(1).unwrap().get_display(1, 0), expected);
    assert_eq!(wb.sheet(1).unwrap().get_display(2, 0), expected);
    let path = dir.path().join("footer-links.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let wb = native::load_workbook(&path).unwrap();
    assert_eq!(wb.named_ranges().get("SalesTotal").unwrap().reference_string(), "D10");
    assert_eq!(wb.sheet(1).unwrap().get_display(2, 0), expected);
    let path = dir.path().join("footer-links.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (mut loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert!(loaded.sheet(1).unwrap().get_raw(1, 0).ends_with("!$D$10"));
    assert_eq!(loaded.sheet(1).unwrap().get_display(1, 0), expected);
    assert_eq!(loaded.sheet(1).unwrap().get_display(2, 0), expected);
    let id = loaded.sheet(0).unwrap().tables()[0].id;
    loaded.append_table_rows(id, 1, &[]).unwrap();
    assert!(loaded.sheet(1).unwrap().get_raw(1, 0).ends_with("!$D$11"));
    assert_eq!(loaded.named_ranges().get("SalesTotal").unwrap().reference_string(), "D11");
    assert_eq!(loaded.sheet(1).unwrap().get_display(2, 0), expected);
    assert_eq!(loaded.sheet(1).unwrap().get_display(1, 0), expected);
}

#[test]
fn freeze_boundaries_roundtrip_through_sorted_and_stored_table_exports() {
    let dir = tempfile::tempdir().unwrap();
    for host in [false, true] {
        for order in [xlsx::ExportOrder::Sorted, xlsx::ExportOrder::Stored] {
            let (mut wb, id) = book();
            set_saved_sort(&mut wb, id, 0, true, true);
            wb.sheet_mut(0).unwrap().frozen_panes = (4, 2);
            let before = authored_snapshot(&wb);
            let mut layout = xlsx::ExportLayout::default();
            layout.frozen_rows = 5;
            layout.frozen_cols = 3;
            let layouts = [layout];
            let file = dir.path().join(format!("freeze-{host}-{order:?}.xlsx"));
            xlsx::export_with_order(&wb, &file, host.then_some(&layouts[..]), order).unwrap();
            let (loaded, report) = xlsx::import(&file).unwrap();
            let expected = if host { (5, 3) } else { (4, 2) };
            assert_eq!(loaded.sheet(0).unwrap().frozen_panes, expected);
            assert_eq!((report.imported_layouts[0].frozen_rows, report.imported_layouts[0].frozen_cols), expected);
            assert!(loaded.sheet(0).unwrap().table_view_spec().unwrap().sort.is_some());
            assert_eq!(loaded.sheet(0).unwrap().get_raw(3, 1), if order == xlsx::ExportOrder::Sorted { "6" } else { "2" });
            assert!(!report.warnings.iter().any(|w| w.contains("settings were not imported")));
            assert_eq!(authored_snapshot(&wb), before);
            assert_eq!(wb.sheet(0).unwrap().frozen_panes, (4, 2));
        }
    }
}

/// Formulas are saved with their computed results. Excel recalculates on
/// open, but readers of saved results (pandas, previews, recipes) used to
/// see 0 for every formula.
#[test]
fn exported_formulas_carry_their_computed_results() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "7");
    wb.set_cell_value_tracked(0, 0, 1, "=A1*3");
    wb.set_cell_value_tracked(0, 0, 2, "=\"id-\"&A1");
    wb.set_cell_value_tracked(0, 0, 3, "=A1>5");
    wb.set_cell_value_tracked(0, 0, 4, "=1/0");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("results.xlsx");
    xlsx::export(&wb, &path, None).unwrap();
    let sheet = xml(&path, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("<f>A1*3</f><v>21</v>"), "{sheet}");
    assert!(sheet.contains("<v>id-7</v>"), "{sheet}");
    assert!(sheet.contains("t=\"b\"") && sheet.contains("<v>1</v>"), "{sheet}");
    assert!(sheet.contains("<v>#DIV/0!</v>"), "{sheet}");
    // And they import as formulas again, with the same results
    let (back, _) = xlsx::import(&path).unwrap();
    assert_eq!(back.sheet(0).unwrap().get_display(0, 1), "21");
}

/// Worksheet parts hold every cell. A sheet over the 32 MB cap for small
/// metadata parts used to abort the whole Table pass ("XLSX part … is too
/// large"), dropping every Table in the workbook, even on sheets without one.
#[test]
fn large_worksheet_parts_do_not_block_table_import() {
    let dir = tempfile::tempdir().unwrap();
    let (wb, _) = book();
    let small = dir.path().join("small.xlsx");
    xlsx::export(&wb, &small, None).unwrap();
    let big = dir.path().join("big.xlsx");
    let padding = format!("<!--{}-->", " ".repeat(33 * 1024 * 1024));
    rewrite(&small, &big, |name, data| {
        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
            let at = data.find("?>").map_or(0, |i| i + 2);
            (name.into(), format!("{}{}{}", &data[..at], padding, &data[at..]))
        } else {
            (name.into(), data)
        }
    });
    let (imported, result) = xlsx::import(&big).unwrap();
    assert!(!result.warnings.iter().any(|w| w.contains("too large")), "{:?}", result.warnings);
    assert_eq!(result.tables_imported, 1, "{:?}", result.warnings);
    assert_eq!(imported.tables().count(), 1);
}

#[test]
fn dynamic_footer_references_survive_native_json_and_stored_excel() {
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.define_name_for_cell("Footer", 0, 8, 3).unwrap();
    wb.set_cell_value_tracked(1, 1, 0, "=OFFSET(Sheet1!$D$9,0,0)");
    wb.set_cell_value_tracked(1, 2, 0, "=INDIRECT(\"Footer\")");
    wb.set_cell_value_tracked(1, 3, 0, "=INDIRECT(\"Sheet1!D9\")");
    wb.append_table_rows(id, 1, &[(8, 1, "2".into()), (8, 2, "10".into())]).unwrap();
    let expected = wb.sheet(0).unwrap().get_display(9, 3);
    let path = dir.path().join("dynamic-footer.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let native = native::load_workbook(&path).unwrap();
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let (json, _, _) = visigrid_io::json::import_any(&json).unwrap();
    let path = dir.path().join("dynamic-footer.xlsx");
    xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    let (excel, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    for mut loaded in [native, json, excel] {
        assert_eq!(loaded.sheet(1).unwrap().get_raw(1, 0), "=OFFSET(Sheet1!$D$10, 0, 0)");
        assert_eq!(loaded.sheet(1).unwrap().get_raw(3, 0), "=INDIRECT(\"Sheet1!D9\")");
        for row in [1, 2] {
            assert_eq!(loaded.sheet(1).unwrap().get_display(row, 0), expected);
        }
        assert_eq!(loaded.sheet(1).unwrap().get_display(3, 0), "20");
        let id = loaded.sheet(0).unwrap().tables()[0].id;
        let commit = loaded.append_table_rows(id, 1, &[(9, 1, "3".into()), (9, 2, "10".into())]).unwrap();
        let new_total = loaded.sheet(0).unwrap().get_display(10, 3);
        assert_ne!(expected, new_total);
        for row in [1, 2] {
            assert_eq!(loaded.sheet(1).unwrap().get_display(row, 0), new_total);
        }
        assert_eq!(loaded.sheet(1).unwrap().get_display(3, 0), "20");
        loaded.apply_table_commit(&commit, true).unwrap();
        assert_eq!(loaded.sheet(1).unwrap().get_display(2, 0), expected);
        loaded.apply_table_commit(&commit, false).unwrap();
        assert_eq!(loaded.sheet(1).unwrap().get_display(2, 0), new_total);
    }
}

#[test]
fn moved_footer_rules_survive_native_json_and_stored_excel() {
    use visigrid_engine::{cond_format::CondStyle, cell::CellStyle, validation::{CellRange, ValidationRule, ValidationType}};
    let dir = tempfile::tempdir().unwrap();
    let (mut wb, id) = book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let mut validation = ValidationRule::custom("=D9>0");
    validation.reference_origin = Some((8, 3));
    wb.sheet_mut(0).unwrap().validations.set(CellRange::single(8, 3), validation);
    wb.sheet_mut(0).unwrap().cond_formats.add(vec![CellRange::single(8, 3)], "=D9>0", CondStyle::Named(CellStyle::Warning));
    wb.append_table_rows(id, 1, &[(8, 1, "2".into()), (8, 2, "10".into())]).unwrap();
    let path = dir.path().join("footer-rules.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let native = native::load_workbook(&path).unwrap();
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let (json, _, _) = visigrid_io::json::import_any(&json).unwrap();
    let path = dir.path().join("footer-rules.xlsx");
    let exported = xlsx::export_with_order(&wb, &path, None, xlsx::ExportOrder::Stored).unwrap();
    assert!(exported.warnings.iter().any(|w| w.contains("semantic styles")));
    let (excel, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    for mut loaded in [native, json, excel] {
        let check = |wb: &Workbook, row: usize| {
            let sheet = wb.sheet(0).unwrap();
            let expected = format!("=D{}>0", row + 1);
            let rule = sheet.validations.get(row, 3).unwrap().at(row, 3);
            assert_eq!(rule.rule_type, ValidationType::Custom(expected.clone()));
            assert!(!sheet.validations.has_validation(row - 1, 3));
            let predicates: Vec<_> = sheet.cond_formats.iter().filter_map(|r| r.predicate_at(row, 3)).collect();
            assert_eq!(predicates, [expected]);
            assert!(!sheet.cond_formats.any_rule_covers(row - 1, 3));
        };
        check(&loaded, 9);
        let id = loaded.sheet(0).unwrap().tables()[0].id;
        let commit = loaded.append_table_rows(id, 1, &[]).unwrap();
        check(&loaded, 10);
        loaded.apply_table_commit(&commit, true).unwrap();
        check(&loaded, 9);
        loaded.apply_table_commit(&commit, false).unwrap();
        check(&loaded, 10);
    }
}
