//! Headless QA of the same capture/renderer used by File > Export PDF.
use std::path::Path;
use visigrid_print::snapshot::{capture, SheetView};
use visigrid_print::{AxisItem, PageSettings, GRID_UNIT_PT};

fn main() -> Result<(), String> {
    let args: Vec<_> = std::env::args().collect();
    if args.len() != 4 {
        return Err("Usage: export_pdf INPUT.xlsx OUTPUT.pdf SHEET_NAME".into());
    }
    let (workbook, imported) = visigrid_io::xlsx::import(Path::new(&args[1]))?;
    let index = workbook
        .sheets()
        .iter()
        .position(|s| s.name == args[3])
        .ok_or("Sheet not found")?;
    let sheet = workbook.sheet(index).unwrap();
    let layout = imported.imported_layouts[index].to_sheet_layout();
    let view = SheetView {
        rows: (0..sheet.rows)
            .filter(|r| !layout.hidden_rows.contains(r))
            .map(|r| AxisItem {
                source_index: r,
                size_pt: f64::from(*layout.row_heights.get(&r).unwrap_or(&28.0)) * GRID_UNIT_PT,
            })
            .collect(),
        columns: (0..sheet.cols)
            .filter(|c| !layout.hidden_cols.contains(c))
            .map(|c| AxisItem {
                source_index: c,
                size_pt: f64::from(*layout.col_widths.get(&c).unwrap_or(&96.0)) * GRID_UNIT_PT,
            })
            .collect(),
    };
    let snapshot = capture(sheet, &view, None, "IBM Plex Sans", 11.0)?;
    let settings = PageSettings {
        footer: true,
        ..Default::default()
    };
    let output = visigrid_print::pdf::render(&snapshot, &settings)?;
    visigrid_print::pdf::save_atomic(Path::new(&args[2]), &output.bytes)?;
    println!(
        "{}: {} pages at {:.1}%, {} clipped cells, {} small-text cells, substituted fonts: {:?}",
        sheet.name,
        output.pages,
        output.scale * 100.0,
        output.clipped_cells,
        output.small_text_cells,
        output.substituted_fonts
    );
    println!(
        "Clipped source cells: {}",
        output.clipped_addresses.join(", ")
    );
    Ok(())
}
