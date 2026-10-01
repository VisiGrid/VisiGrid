#![cfg(feature = "preview")]
use std::sync::Arc;
use visigrid_engine::cell::{CellBorder, CellFormat};
use visigrid_engine::sheet::{MergedRegion, Sheet, SheetId};
use visigrid_print::preview::{rasterize, PreviewPage};
use visigrid_print::snapshot::{capture, Scope, SheetView};
use visigrid_print::{paginate, pdf, AxisItem, PageSettings, Paper, Scale};

fn view() -> SheetView {
    SheetView {
        rows: (0..3)
            .map(|source_index| AxisItem {
                source_index,
                size_pt: 40.0,
            })
            .collect(),
        columns: (0..3)
            .map(|source_index| AxisItem {
                source_index,
                size_pt: 100.0,
            })
            .collect(),
    }
}
fn settings(gridlines: bool) -> PageSettings {
    PageSettings {
        gridlines,
        paper: Paper::Letter,
        scale: Scale::Actual,
        ..Default::default()
    }
}
fn pixel(page: &PreviewPage, x: u32, y: u32) -> &[u8] {
    let offset = ((y * page.width + x) * 4) as usize;
    &page.rgba[offset..offset + 4]
}
fn white_near(page: &PreviewPage, x: u32, y: u32) -> bool {
    (x - 1..=x + 1).all(|x| (y - 1..=y + 1).all(|y| pixel(page, x, y) == [255, 255, 255, 255]))
}

#[test]
fn preview_uses_pdf_geometry_and_gridlines_respect_merges_fills_and_borders() {
    let mut sheet = Sheet::new(SheetId(1), 1000, 10);
    sheet.add_merge(MergedRegion::new(0, 0, 0, 1)).unwrap();
    sheet.set_format(
        1,
        1,
        CellFormat {
            background_color: Some([255, 255, 255, 255]),
            ..Default::default()
        },
    );
    sheet.set_format(
        1,
        2,
        CellFormat {
            background_color: Some([255, 255, 255, 255]),
            ..Default::default()
        },
    );
    sheet.set_format(
        2,
        2,
        CellFormat {
            border_bottom: CellBorder::thin(),
            ..Default::default()
        },
    );
    // Gridlines must not pull distant formatting into the selected print scope.
    sheet.set_format(
        999,
        8,
        CellFormat {
            background_color: Some([0, 0, 0, 255]),
            ..Default::default()
        },
    );
    let snapshot = capture(
        &sheet,
        &view(),
        Some(Scope {
            rows: 0..3,
            columns: 0..3,
        }),
        "IBM Plex Sans",
        11.0,
    )
    .unwrap();
    let on = settings(true);
    let off = settings(false);
    let on_plan = paginate(&snapshot.layout, &on).unwrap();
    let off_plan = paginate(&snapshot.layout, &off).unwrap();
    assert_eq!(on_plan.pages(), off_plan.pages());
    assert_eq!(on_plan.scale(), off_plan.scale());
    let output = Arc::new(pdf::render(&snapshot, &on).unwrap());
    let original = output.bytes.clone();
    let page = rasterize(output.clone(), 0, 1584).unwrap(); // exactly 2 px/pt
    let plain = rasterize(Arc::new(pdf::render(&snapshot, &off).unwrap()), 0, 1584).unwrap();
    assert_eq!((page.width, page.height), (1224, 1584));
    // Unmerged blank-cell boundary at x=136pt, y=130pt.
    assert!(!white_near(&page, 272, 260));
    assert!(white_near(&plain, 272, 260));
    // Interior of A1:B1 merge has no line at x=136pt.
    assert!(white_near(&page, 272, 100));
    // Both cells either side of this boundary have explicit white fills.
    assert!(white_near(&page, 472, 180));
    // Explicit bottom border remains in both outputs, even when gridlines are off.
    assert!(!white_near(&page, 550, 312));
    assert!(!white_near(&plain, 550, 312));
    // Empty sheet beyond the scope and the page margin are not a printed grid.
    assert!(white_near(&page, 800, 400));
    assert!(white_near(&page, 40, 40));
    let small = rasterize(output.clone(), 0, 792).unwrap();
    assert_eq!((small.width, small.height), (612, 792));
    assert_eq!(
        output.bytes, original,
        "preview resolution must not modify exported bytes"
    );
}

#[test]
fn preview_renders_shaped_text_and_rejects_invalid_requests() {
    let mut sheet = Sheet::new(SheetId(1), 3, 3);
    sheet.set_value(0, 0, "Café résumé");
    let snapshot = capture(&sheet, &view(), None, "IBM Plex Sans", 11.0).unwrap();
    let output = Arc::new(pdf::render(&snapshot, &settings(false)).unwrap());
    let page = rasterize(output.clone(), 0, 792).unwrap();
    assert!((40..130).any(|x| (40..70).any(|y| pixel(&page, x, y)[0] < 200)));
    assert!(rasterize(output.clone(), 1, 792).is_err());
    assert!(rasterize(output, 0, 2401).is_err());
    assert!(rasterize(Arc::new(b"not a PDF".to_vec()), 0, 792).is_err());
    assert!(!PageSettings::default().gridlines);
}

#[test]
fn center_across_selection_centers_over_the_whole_span() {
    use visigrid_engine::cell::Alignment;
    let mut sheet = Sheet::new(SheetId(1), 3, 3);
    sheet.set_value(0, 0, "Quarterly Revenue");
    for c in 0..3 {
        sheet.set_alignment(0, c, Alignment::CenterAcrossSelection);
    }
    let snapshot = capture(&sheet, &view(), None, "IBM Plex Sans", 11.0).unwrap();
    let page = rasterize(Arc::new(pdf::render(&snapshot, &settings(false)).unwrap()), 0, 1584).unwrap();
    // Columns are 100pt from a 36pt margin: A 72-272px, B 272-472px, C 472-672px.
    // Row 1 spans 72-152px. Centered over A:C, the text sits around x=372px.
    let ink = |x0: u32, x1: u32| (x0..x1).any(|x| (80..145).any(|y| pixel(&page, x, y)[0] < 160));
    assert!(ink(300, 450), "text is drawn over column B, the middle of the span");
    assert!(!ink(80, 260), "nothing at the left of column A, so it is not centered in A alone");
}
