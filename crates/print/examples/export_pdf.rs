//! Headless QA of the same capture/renderer used by File > Export PDF.
use std::path::Path;
use visigrid_engine::print_setup::PrintRows;
use visigrid_print::snapshot::{capture, SheetView};
use visigrid_print::{AxisItem, GRID_UNIT_PT};

fn main() -> Result<(), String> {
    let args: Vec<_> = std::env::args().collect();
    if args.len() < 4 {
        return Err("Usage: export_pdf INPUT.xlsx|INPUT.sheet OUTPUT.pdf SHEET_NAME [--gridlines] [--repeat-top=N]".into());
    }
    let mut gridlines = false;
    let mut repeat = None;
    for flag in &args[4..] {
        if flag == "--gridlines" {
            gridlines = true;
        } else if let Some(value) = flag.strip_prefix("--repeat-top=") {
            repeat = Some(value.parse::<usize>().map_err(|_| "Invalid header count")?);
        } else {
            return Err(format!("Unknown option: {flag}"));
        }
    }
    let input = Path::new(&args[1]);
    let native = input
        .extension()
        .is_some_and(|e| e.eq_ignore_ascii_case("sheet"));
    let (workbook, layouts) = if native {
        let wb = visigrid_io::native::load_workbook(input)?;
        let layout = visigrid_io::native::load_layout(input);
        let layouts = (0..wb.sheet_count())
            .map(|i| visigrid_io::json::SheetLayout {
                col_widths: layout
                    .col_widths
                    .get(&i)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .collect(),
                row_heights: layout
                    .row_heights
                    .get(&i)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .collect(),
                hidden_rows: layout
                    .hidden_rows
                    .get(&i)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .collect(),
                hidden_cols: layout
                    .hidden_cols
                    .get(&i)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .collect(),
                ..Default::default()
            })
            .collect::<Vec<_>>();
        (wb, layouts)
    } else {
        let (wb, imported) = visigrid_io::xlsx::import(input)?;
        (
            wb,
            imported
                .imported_layouts
                .iter()
                .map(|l| l.to_sheet_layout())
                .collect(),
        )
    };
    let index = workbook
        .sheets()
        .iter()
        .position(|s| s.name == args[3])
        .ok_or("Sheet not found")?;
    let sheet = workbook.sheet(index).unwrap();
    let layout = &layouts[index];
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
    let mut saved = sheet.print_setup.clone();
    if !native {
        saved.page_numbers = true;
    }
    if gridlines {
        saved.gridlines = true;
    }
    let snapshot = if let Some(area) = saved.area {
        visigrid_print::setup::capture_area(sheet, &view, area, "IBM Plex Sans", 11.0)?
    } else {
        capture(sheet, &view, None, "IBM Plex Sans", 11.0)?
    };
    if let Some(count) = repeat {
        if count >= snapshot.layout.rows.len() {
            return Err("Header rows must leave body content".into());
        }
        saved.repeat_rows = (count > 0).then(|| PrintRows {
            start: snapshot.layout.rows[0].source_index,
            end: snapshot.layout.rows[count - 1].source_index,
        });
    }
    let settings = visigrid_print::setup::page_settings(&snapshot, &saved)?;
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
