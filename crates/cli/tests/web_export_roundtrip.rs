// The web grid's Export and Open run `vgrid convert` on the server: the live
// workbook's visigrid-json (CollabClient.snapshot) → XLSX, and a user's XLSX →
// visigrid-json. Every property the grid edits must survive that trip.
//
// Run with: cargo test -p visigrid-cli --test web_export_roundtrip

use std::process::Command;

use visigrid_engine::cell::{Alignment, BorderStyle, CellBorder, NumberFormat, TextOverflow};
use visigrid_engine::sheet::MergedRegion;
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::SheetLayout;

fn vgrid(args: &[&str]) {
    let out = Command::new(env!("CARGO_BIN_EXE_vgrid")).args(args).output().unwrap();
    assert!(out.status.success(), "vgrid {args:?}: {}", String::from_utf8_lossy(&out.stderr));
}

#[test]
fn every_grid_property_survives_json_to_xlsx_and_back() {
    let mut wb = Workbook::new();
    let s = wb.sheet_mut(0).unwrap();
    s.set_value(0, 0, "Item");
    s.set_value(0, 1, "Price");
    s.set_value(1, 0, "Widget");
    s.set_value(1, 1, "12.5");
    s.set_value(2, 1, "=B2*2");
    let mut f = s.get_format(0, 0);
    f.bold = true;
    f.italic = true;
    f.underline = true;
    f.font_family = Some("Arial".into());
    f.font_size = Some(14.0);
    f.font_color = Some([0xC0, 0x10, 0x20, 255]);
    f.background_color = Some([0xFF, 0xEB, 0x3B, 255]);
    f.alignment = Alignment::Center;
    f.text_overflow = TextOverflow::Wrap;
    f.border_bottom = CellBorder { style: BorderStyle::Medium, color: None };
    f.border_top = CellBorder { style: BorderStyle::Thin, color: None };
    s.set_format(0, 0, f);
    let mut money = s.get_format(1, 1);
    money.number_format = NumberFormat::Custom("$#,##0.00".into());
    s.set_format(1, 1, money.clone());
    s.set_format(2, 1, money);
    s.add_merge(MergedRegion::new(5, 0, 5, 2)).unwrap();
    let second = wb.add_sheet_named("Data").unwrap();
    wb.sheet_mut(second).unwrap().set_value(0, 0, "=Sheet1!B3+1");

    let mut layout = SheetLayout::default();
    layout.col_widths.insert(0, 150.0);
    layout.row_heights.insert(3, 36.0);
    layout.hidden_rows.insert(4);
    layout.hidden_cols.insert(3);
    layout.frozen_rows = 1;
    layout.frozen_cols = 1;
    let layouts = vec![layout.clone(), SheetLayout::default()];

    let dir = tempfile::tempdir().unwrap();
    let json_in = dir.path().join("in.json");
    let xlsx = dir.path().join("out.xlsx");
    let json_out = dir.path().join("back.json");
    std::fs::write(&json_in, visigrid_io::json::export_workbook(&wb, &layouts, 0).unwrap()).unwrap();
    vgrid(&["convert", "-f", "json-full", "-t", "xlsx", json_in.to_str().unwrap(), "-o", xlsx.to_str().unwrap()]);
    vgrid(&["convert", "-f", "xlsx", "-t", "json-full", xlsx.to_str().unwrap(), "-o", json_out.to_str().unwrap()]);
    let (back, back_layouts, _) = visigrid_io::json::import_any(&std::fs::read_to_string(&json_out).unwrap()).unwrap();

    assert_eq!(back.sheets().iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Sheet1", "Data"]);
    let (a, b) = (&wb.sheets()[0], &back.sheets()[0]);
    for (r, c, shown) in [(0, 0, "Item"), (0, 1, "Price"), (1, 0, "Widget"), (1, 1, "$12.50"), (2, 1, "$25.00")] {
        assert_eq!(b.get_raw(r, c), a.get_raw(r, c), "raw at {r},{c}");
        assert_eq!(b.get_formatted_display(r, c), shown, "display at {r},{c}");
    }
    let (fa, fb) = (a.get_format(0, 0), b.get_format(0, 0));
    assert_eq!((fb.bold, fb.italic, fb.underline), (true, true, true));
    assert_eq!(fb.font_family, fa.font_family);
    assert_eq!(fb.font_size, fa.font_size);
    assert_eq!(fb.font_color, fa.font_color);
    assert_eq!(fb.background_color, fa.background_color);
    assert_eq!(fb.alignment, Alignment::Center);
    assert_eq!(fb.text_overflow, TextOverflow::Wrap);
    assert_eq!(fb.border_bottom.style, BorderStyle::Medium);
    assert_eq!(fb.border_top.style, BorderStyle::Thin);
    assert_eq!(b.merged_regions.len(), 1);
    assert_eq!((b.merged_regions[0].start, b.merged_regions[0].end), ((5, 0), (5, 2)));
    let l = &back_layouts[0];
    assert!((l.col_widths[&0] - 150.0).abs() < 1.0, "width {:?}", l.col_widths);
    assert!((l.row_heights[&3] - 36.0).abs() < 1.0, "height {:?}", l.row_heights);
    assert!(l.hidden_rows.contains(&4) && l.hidden_cols.contains(&3));
    assert_eq!((l.frozen_rows, l.frozen_cols), (1, 1));
    // A formula on another sheet still reads across.
    assert_eq!(back.sheets()[1].get_raw(0, 0), "=Sheet1!B3+1");
}

#[test]
fn csv_export_writes_one_sheet_values_as_displayed() {
    let mut wb = Workbook::new();
    let s = wb.sheet_mut(0).unwrap();
    s.set_value(0, 0, "Name, with comma");
    s.set_value(0, 1, "=2+3");
    let other = wb.add_sheet_named("Other").unwrap();
    wb.sheet_mut(other).unwrap().set_value(0, 0, "second");
    let dir = tempfile::tempdir().unwrap();
    let json_in = dir.path().join("in.json");
    let csv = dir.path().join("out.csv");
    std::fs::write(&json_in, visigrid_io::json::export_workbook(&wb, &[SheetLayout::default(), SheetLayout::default()], 0).unwrap()).unwrap();
    vgrid(&["convert", "-f", "json-full", "-t", "csv", "--sheet", "Other", json_in.to_str().unwrap(), "-o", csv.to_str().unwrap()]);
    assert_eq!(std::fs::read_to_string(&csv).unwrap().trim(), "second");
    vgrid(&["convert", "-f", "json-full", "-t", "csv", "--sheet", "Sheet1", json_in.to_str().unwrap(), "-o", csv.to_str().unwrap()]);
    assert_eq!(std::fs::read_to_string(&csv).unwrap().trim(), "\"Name, with comma\",5");
}
