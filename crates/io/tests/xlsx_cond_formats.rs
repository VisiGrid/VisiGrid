use std::io::Read;
use visigrid_engine::{
    cell::{Alignment, BorderStyle, CellBorder, CellFormatOverride, CellStyle, NumberFormat},
    cond_format::CondStyle,
    sheet::{NUM_COLS, NUM_ROWS},
    validation::CellRange,
    workbook::Workbook,
};
use visigrid_io::xlsx::{self, ExportOrder};
fn xml(bytes: &[u8], part: &str) -> String {
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut text = String::new();
    zip.by_name(part)
        .unwrap()
        .read_to_string(&mut text)
        .unwrap();
    text
}
fn roundtrip(wb: &Workbook) -> (Workbook, Vec<u8>, xlsx::ExportResult) {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("rules.xlsx");
    let (bytes, exported) =
        xlsx::export_to_buffer_with_order(wb, None, ExportOrder::Stored).unwrap();
    std::fs::write(&path, &bytes).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    (loaded, bytes, exported)
}
#[test]
fn differential_formats_preserve_false_flags_borders_numbers_and_escaping() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "2");
    let mut style = CellFormatOverride {
        bold: Some(false),
        italic: Some(true),
        underline: Some(false),
        strikethrough: Some(false),
        font_color: Some(Some([12, 34, 56, 255])),
        background_color: Some(Some([200, 210, 220, 255])),
        number_format: Some(NumberFormat::Custom("0.00\" &lt; Δ\n\"".into())),
        border_left: Some(CellBorder {
            style: BorderStyle::Medium,
            color: Some([11, 22, 33, 255]),
        }),
        border_bottom: Some(CellBorder::default()),
        ..Default::default()
    };
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::single(0, 0)],
        "=A1>1",
        CondStyle::Inline(style.clone()),
    );
    let (loaded, bytes, report) = roundtrip(&wb);
    assert!(report.warnings.is_empty());
    let imported = loaded
        .active_sheet()
        .cond_formats
        .iter()
        .next()
        .unwrap()
        .style
        .as_override();
    // An explicit no-border differential carries the default border color.
    style.border_bottom.as_mut().unwrap().color = Some([0, 0, 0, 255]);
    assert_eq!(imported, style);
    let styles = xml(&bytes, "xl/styles.xml");
    assert!(styles.contains("<b val=\"0\"/>") && styles.contains("<u val=\"none\"/>"));
    assert!(styles.contains("&amp;lt;"));
    let sheet = xml(&bytes, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("type=\"expression\"") && sheet.contains("dxfId=\"0\""));
    assert!(!sheet.contains("stopIfTrue=\"1\""));
}
#[test]
fn overlapping_ranges_keep_native_anchors_and_later_property_precedence() {
    let mut wb = Workbook::new();
    for row in 0..12 {
        for col in 0..5 {
            wb.set_cell_value_tracked(0, row, col, &(row + col).to_string());
        }
    }
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 0, 6, 2), CellRange::new(3, 1, 10, 3)],
        "=B1>4",
        CondStyle::Inline(CellFormatOverride {
            bold: Some(true),
            background_color: Some(Some([200, 0, 0, 255])),
            ..Default::default()
        }),
    );
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(2, 0, 8, 3)],
        "=$B3>5",
        CondStyle::Inline(CellFormatOverride {
            bold: Some(false),
            italic: Some(true),
            ..Default::default()
        }),
    );
    let (loaded, bytes, report) = roundtrip(&wb);
    assert!(report.warnings.is_empty());
    for row in 0..12 {
        for col in 0..5 {
            let original =
                wb.active_sheet()
                    .cond_formats
                    .override_for_cell(row, col, wb.active_sheet());
            let imported = loaded.active_sheet().cond_formats.override_for_cell(
                row,
                col,
                loaded.active_sheet(),
            );
            assert_eq!(original, imported, "cell {row},{col}");
        }
    }
    let sheet = xml(&bytes, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains("priority=\"1\"") && sheet.contains("priority=\"2\""));
}
#[test]
fn metadata_only_whole_columns_empty_styles_and_like_snapshots_roundtrip() {
    let mut wb = Workbook::new();
    let range = CellRange::new(0, 0, NUM_ROWS - 1, 0);
    wb.active_sheet_mut().cond_formats.add(
        vec![range],
        "=A1<>\"<&Δ>\"",
        CondStyle::Like {
            source: (4, 7),
            snapshot: CellFormatOverride {
                font_color: Some(Some([1, 2, 3, 255])),
                ..Default::default()
            },
        },
    );
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::single(0, NUM_COLS - 1)],
        "=TRUE",
        CondStyle::Inline(Default::default()),
    );
    let (loaded, bytes, report) = roundtrip(&wb);
    assert!(report.warnings.is_empty());
    assert_eq!(loaded.active_sheet().cond_formats.len(), 2);
    assert!(loaded
        .active_sheet()
        .cond_formats
        .iter()
        .any(|r| r.ranges == [range] && r.predicate == "=A1<>\"<&Δ>\""));
    assert!(
        bytes.len() < 20_000,
        "whole-column coverage never enumerates cells"
    );
}
#[test]
fn partial_style_losses_and_inert_rules_are_reported_in_preflight_and_export() {
    let mut wb = Workbook::new();
    let ranges = vec![CellRange::single(0, 0)];
    let off = wb.active_sheet_mut().cond_formats.add(
        ranges.clone(),
        "=TRUE",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .cond_formats
        .get_mut(off)
        .unwrap()
        .enabled = false;
    wb.active_sheet_mut().cond_formats.add(
        ranges.clone(),
        "=SUM((",
        CondStyle::Inline(Default::default()),
    );
    wb.active_sheet_mut().cond_formats.add(
        ranges.clone(),
        "=TRUE",
        CondStyle::Named(CellStyle::Success),
    );
    wb.active_sheet_mut().cond_formats.add(
        ranges,
        "=TRUE",
        CondStyle::Inline(CellFormatOverride {
            bold: Some(true),
            font_family: Some(Some("Custom".into())),
            alignment: Some(Alignment::Center),
            background_color: Some(Some([1, 2, 3, 128])),
            font_color: Some(None),
            ..Default::default()
        }),
    );
    let expected = xlsx::table_export_warnings_with_order(&wb, None, ExportOrder::Stored).unwrap();
    let (loaded, _, report) = roundtrip(&wb);
    assert_eq!(expected, report.warnings);
    for message in [
        "disabled",
        "unparseable",
        "semantic styles",
        "font family",
        "transparent",
        "inheritance",
    ] {
        assert!(
            report.warnings.iter().any(|w| w.contains(message)),
            "{message}: {:?}",
            report.warnings
        );
    }
    assert_eq!(loaded.active_sheet().cond_formats.len(), 2);
}
#[test]
fn invalid_bounds_refuse_before_replacing_an_existing_file() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("keep.xlsx");
    std::fs::write(&path, b"keep me").unwrap();
    let mut wb = Workbook::new();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::single(NUM_ROWS, 0)],
        "=TRUE",
        CondStyle::Inline(Default::default()),
    );
    assert!(xlsx::table_export_warnings_with_order(&wb, None, ExportOrder::Stored).is_err());
    assert!(xlsx::export_with_order(&wb, &path, None, ExportOrder::Stored).is_err());
    assert_eq!(std::fs::read(&path).unwrap(), b"keep me");
}
#[test]
fn excel_shared_sqref_anchor_is_rebased_for_native_range_anchors() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("shared.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let sheet = excel.add_worksheet();
    sheet
        .add_conditional_format(
            0,
            0,
            1,
            0,
            &rust_xlsxwriter::ConditionalFormatFormula::new()
                .set_rule("=B1>0")
                .set_multi_range("A1:A2 C3:C4")
                .set_format(rust_xlsxwriter::Format::new().set_bold()),
        )
        .unwrap();
    excel.save(&path).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    let rules = &loaded.active_sheet().cond_formats;
    let sources = |r, c| {
        rules
            .iter()
            .filter_map(|rule| rule.predicate_at(r, c))
            .collect::<Vec<_>>()
    };
    assert_eq!(sources(0, 0), ["=B1>0"]);
    assert_eq!(sources(2, 2), ["=D3>0"]);
    assert_eq!(sources(3, 2), ["=D4>0"]);
}

#[test]
fn forbidden_xml_in_rule_style_refuses_without_replacing_the_destination() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("keep.xlsx");
    std::fs::write(&path, b"original").unwrap();
    let mut wb = Workbook::new();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::single(0, 0)],
        "=TRUE",
        CondStyle::Inline(CellFormatOverride {
            number_format: Some(NumberFormat::Custom("bad\u{1}".into())),
            ..Default::default()
        }),
    );
    assert!(
        xlsx::export_with_order(&wb, &path, None, ExportOrder::Stored)
            .unwrap_err()
            .contains("XML")
    );
    assert_eq!(std::fs::read(&path).unwrap(), b"original");
}

#[test]
fn empty_differential_border_edges_clear_only_the_present_sides() {
    use visigrid_io::xlsx_styles::{parse_dxfs, ThemePalette};
    let styles = parse_dxfs(
        r#"<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:dxfs count="3"><x:dxf/>
        <x:dxf><x:font/><x:border><x:left/><x:right style="thin">
        <x:color rgb="FF123456"/></x:right></x:border></x:dxf>
        <x:dxf><x:border/><x:font><x:b val="0"/></x:font></x:dxf>
        </x:dxfs></x:styleSheet>"#,
        &ThemePalette::default(),
    );
    assert_eq!(styles.len(), 3);
    assert!(styles[0].is_empty());
    assert_eq!(styles[1].extra.border_left, Some(CellBorder::default()));
    assert_eq!(
        styles[1].extra.border_right,
        Some(CellBorder {
            style: BorderStyle::Thin,
            color: Some([0x12, 0x34, 0x56, 255]),
        })
    );
    assert_eq!(styles[1].extra.border_top, None);
    assert_eq!(styles[1].extra.border_bottom, None);
    assert_eq!(
        styles[1].font_color, None,
        "empty font cannot capture border color"
    );
    assert_eq!(styles[2].extra, CellFormatOverride::default());
    assert_eq!(styles[2].bold, Some(false));
    assert!(styles.iter().all(|dxf| !dxf.has_unmapped));
}

#[test]
fn interleaved_ranges_keep_their_styles_and_global_precedence() {
    let mut wb = Workbook::new();
    for (col, bold, italic) in [(0, Some(true), None), (1, None, Some(true)), (0, Some(false), None)] {
        wb.active_sheet_mut().cond_formats.add(vec![CellRange::new(0, col, 2, col)], "=TRUE",
            CondStyle::Inline(CellFormatOverride { bold, italic, ..Default::default() }));
    }
    let (loaded, _, _) = roundtrip(&wb);
    for col in 0..2 {
        let expected = wb.active_sheet().cond_formats.override_for_cell(0, col, wb.active_sheet());
        let actual = loaded.active_sheet().cond_formats.override_for_cell(0, col, loaded.active_sheet());
        assert_eq!(actual, expected, "column {col}");
    }
}

#[test]
fn structured_conditional_rules_warn_instead_of_exporting_invalid_excel_syntax() {
    let mut wb = Workbook::new();
    wb.active_sheet_mut().cond_formats.add(vec![CellRange::new(0, 0, 2, 0)], "=Sales[Amount]>0",
        CondStyle::Inline(CellFormatOverride { bold: Some(true), ..Default::default() }));
    let (bytes, report) = xlsx::export_to_buffer_with_order(&wb, None, ExportOrder::Stored).unwrap();
    assert!(!xml(&bytes, "xl/worksheets/sheet1.xml").contains("<conditionalFormatting"));
    assert!(report.warnings.iter().any(|w| w.contains("structured references")));
    assert_eq!(wb.active_sheet().cond_formats.iter().count(), 1);
}

#[test]
fn interleaved_overlapping_ranges_keep_global_rule_precedence() {
    let mut wb = Workbook::new();
    for (col, bold, italic) in [(0, Some(true), None), (1, Some(false), None), (0, None, Some(true))] {
        wb.active_sheet_mut().cond_formats.add(vec![CellRange::new(0, col, 2, col + 1)], "=TRUE",
            CondStyle::Inline(CellFormatOverride { bold, italic, ..Default::default() }));
    }
    let (loaded, _, _) = roundtrip(&wb);
    for col in 0..3 {
        assert_eq!(loaded.active_sheet().cond_formats.override_for_cell(0, col, loaded.active_sheet()),
            wb.active_sheet().cond_formats.override_for_cell(0, col, wb.active_sheet()), "column {col}");
    }
}

#[test]
fn refused_row_insert_under_full_column_rules_still_saves_and_reopens_in_every_format() {
    use visigrid_engine::{structural::Axis, validation::{NumericConstraint, ValidationRule}};
    use visigrid_io::{native, json};
    let mut wb = Workbook::new();
    let range = CellRange::new(0, 0, NUM_ROWS - 1, 0);
    wb.active_sheet_mut().validations.set(range, ValidationRule::decimal(NumericConstraint::between(0.0, 10.0)));
    wb.active_sheet_mut().cond_formats.add(vec![range], "=A1>0", CondStyle::Named(CellStyle::Warning));
    let before = format!("{wb:?}");
    let error = wb.structural_edit(0, Axis::Row, 1, 1, false).unwrap_err();
    assert!(error.contains("conditional format") && error.contains("last row"), "{error}");
    assert_eq!(format!("{wb:?}"), before);
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("full-column.sheet");
    native::save_workbook(&wb, &path).unwrap();
    let native_copy = native::load_workbook(&path).unwrap();
    let (excel_copy, _, _) = roundtrip(&wb);
    let encoded = json::export_workbook(&wb, &[], 0).unwrap();
    let (json_copy, _, _) = json::import_any(&encoded).unwrap();
    for copy in [&native_copy, &excel_copy, &json_copy] {
        assert_eq!(*copy.active_sheet().validations.iter().next().unwrap().0, range);
        assert!(copy.active_sheet().validations.effective_ranges().is_ok());
        assert_eq!(copy.active_sheet().cond_formats.iter().next().unwrap().ranges, vec![range]);
    }
}
