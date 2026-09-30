//! Headless QA driver. Outputs are deliberately separate from the input.
//! cargo run -p visigrid-io --example xlsx_probe -- INPUT OUTPUT_DIRECTORY
use std::{env, fs, path::PathBuf};
use visigrid_io::{native, xlsx};

fn main() -> Result<(), String> {
    let args: Vec<_> = env::args_os().collect();
    if args.len() != 3 {
        return Err("usage: xlsx_probe INPUT OUTPUT_DIRECTORY".into());
    }
    let source = PathBuf::from(&args[1]);
    let output = PathBuf::from(&args[2]);
    fs::create_dir_all(&output).map_err(|e| e.to_string())?;
    if source.to_str() == Some("--generate") {
        return generate_basics(&output);
    }
    let (mut wb, imported) = xlsx::import(&source)?;
    if source.file_name().and_then(|s| s.to_str()) == Some("basics.xlsx") {
        use visigrid_engine::formula::eval::Value;
        assert!(matches!(wb.sheets()[0].get_computed_value(7, 1), Value::Number(n) if n == 2469.0));
        assert!(matches!(wb.sheets()[0].get_computed_value(8, 1), Value::Number(n) if n == 1241.5));
    }
    let layouts: Vec<_> = imported
        .imported_layouts
        .iter()
        .map(|l| l.to_export_layout())
        .collect();
    let mut log = format!("IMPORT\n{imported:#?}\n");
    let exported = xlsx::export(&wb, &output.join("unchanged.xlsx"), Some(&layouts))?;
    log.push_str(&format!("EXPORT\n{exported:#?}\n"));
    let (reopened, _) = xlsx::import(&output.join("unchanged.xlsx"))?;
    assert_eq!(wb.sheets().len(), reopened.sheets().len());

    let path = output.join("native.sheet");
    native::save_workbook_full(&wb, &Default::default(), &[], &[], &path)?;
    let mut layout = native::SheetLayout {
        col_widths: Default::default(),
        row_heights: Default::default(),
        hidden_rows: Default::default(),
        hidden_cols: Default::default(),
    };
    for (i, l) in imported.imported_layouts.iter().enumerate() {
        let l = l.to_sheet_layout();
        layout
            .col_widths
            .insert(i, l.col_widths.into_iter().collect());
        layout
            .row_heights
            .insert(i, l.row_heights.into_iter().collect());
        layout
            .hidden_rows
            .insert(i, l.hidden_rows.into_iter().collect());
        layout
            .hidden_cols
            .insert(i, l.hidden_cols.into_iter().collect());
    }
    native::save_layout(&path, &layout)?;
    let restored = native::load_workbook(&path)?;
    let restored_layout = native::load_layout(&path);
    let native_layouts: Vec<_> = restored
        .sheets()
        .iter()
        .enumerate()
        .map(|(i, s)| xlsx::ExportLayout {
            col_widths: restored_layout
                .col_widths
                .get(&i)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect(),
            row_heights: restored_layout
                .row_heights
                .get(&i)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect(),
            hidden_rows: restored_layout
                .hidden_rows
                .get(&i)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect(),
            hidden_cols: restored_layout
                .hidden_cols
                .get(&i)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .collect(),
            frozen_rows: s.frozen_panes.0,
            frozen_cols: s.frozen_panes.1,
            ..Default::default()
        })
        .collect();
    xlsx::export(
        &restored,
        &output.join("native-roundtrip.xlsx"),
        Some(&native_layouts),
    )?;
    xlsx::import(&output.join("native-roundtrip.xlsx"))?;

    // Replace one designated cell; retain its style and leave other sheets alone.
    wb.set_cell_value_tracked(0, 0, 0, "VisiGrid QA edit");
    xlsx::export(&wb, &output.join("edited.xlsx"), Some(&layouts))?;
    let (edited, _) = xlsx::import(&output.join("edited.xlsx"))?;
    assert_eq!(edited.sheets()[0].get_raw(0, 0), "VisiGrid QA edit");
    fs::write(output.join("diagnostics.txt"), log).map_err(|e| e.to_string())?;
    Ok(())
}

fn generate_basics(output: &std::path::Path) -> Result<(), String> {
    use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook};
    let mut book = Workbook::new();
    let title = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::RGB(0x245A81));
    let money = Format::new().set_num_format("$#,##0.00");
    let date = Format::new().set_num_format("yyyy-mm-dd");
    let percent = Format::new().set_num_format("0.00%");
    let border = Format::new()
        .set_border(FormatBorder::Thin)
        .set_background_color(Color::RGB(0xE2F0D9));
    let s = book.add_worksheet();
    s.set_name("Basics").map_err(|e| e.to_string())?;
    s.write_string_with_format(0, 0, "VisiGrid round-trip basics", &title)
        .map_err(|e| e.to_string())?;
    s.write_string(1, 0, "Leading zeros")
        .map_err(|e| e.to_string())?;
    s.write_string(1, 1, "00123").map_err(|e| e.to_string())?;
    s.write_number_with_format(2, 1, 1234.5, &money)
        .map_err(|e| e.to_string())?;
    s.write_number_with_format(3, 1, 46295.0, &date)
        .map_err(|e| e.to_string())?;
    s.write_number_with_format(4, 1, 0.125, &percent)
        .map_err(|e| e.to_string())?;
    s.write_boolean(5, 1, true).map_err(|e| e.to_string())?;
    s.write_string(6, 1, "café — 日本語 & <text>")
        .map_err(|e| e.to_string())?;
    s.write_formula(7, 1, "=B3*2").map_err(|e| e.to_string())?;
    s.write_formula(8, 1, "='Other sheet'!A1+B3")
        .map_err(|e| e.to_string())?;
    s.merge_range(10, 0, 10, 2, "Merged green cells", &border)
        .map_err(|e| e.to_string())?;
    s.set_column_width(0, 30.0).map_err(|e| e.to_string())?;
    s.set_column_width(1, 24.0).map_err(|e| e.to_string())?;
    s.set_row_height(0, 28.0).map_err(|e| e.to_string())?;
    s.set_row_hidden(12).map_err(|e| e.to_string())?;
    s.set_column_hidden(4).map_err(|e| e.to_string())?;
    s.set_freeze_panes(1, 1).map_err(|e| e.to_string())?;
    book.add_worksheet()
        .set_name("Other sheet")
        .map_err(|e| e.to_string())?
        .write_number(0, 0, 7.0)
        .map_err(|e| e.to_string())?;
    book.save(output.join("basics.xlsx"))
        .map_err(|e| e.to_string())
}
