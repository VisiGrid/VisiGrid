#![cfg(feature = "pdf")]
use visigrid_engine::cell::{CellFormat, NumberFormat, TextOverflow};
use visigrid_engine::sheet::{MergedRegion, Sheet, SheetId};
use visigrid_print::snapshot::{capture, Scope, SheetView};
use visigrid_print::{AxisItem, PageSettings};

fn view(rows: &[usize], cols: &[usize]) -> SheetView {
    SheetView {
        rows: rows
            .iter()
            .map(|&source_index| AxisItem {
                source_index,
                size_pt: 24.0,
            })
            .collect(),
        columns: cols
            .iter()
            .map(|&source_index| AxisItem {
                source_index,
                size_pt: 100.0,
            })
            .collect(),
    }
}

#[test]
fn bounds_ignore_hidden_values_empty_results_and_distant_formatting() {
    let mut sheet = Sheet::new(SheetId(1), 1001, 10);
    sheet.set_value(2, 1, "visible");
    sheet.set_value(9, 1, "hidden row");
    sheet.set_value(2, 9, "hidden column");
    sheet.set_value(499, 1, "=IF(1=1,\"\",\"unused\")");
    sheet.set_value(500, 1, "0");
    sheet.set_number_format(500, 1, NumberFormat::Custom(";;;".into()));
    sheet.set_format(
        1000,
        4,
        CellFormat {
            background_color: Some([255, 0, 0, 255]),
            ..Default::default()
        },
    );
    let v = view(
        &(0..1001).filter(|&r| r != 9).collect::<Vec<_>>(),
        &[0, 1, 2, 3, 4],
    );
    let snapshot = capture(&sheet, &v, None, "IBM Plex Sans", 11.0).unwrap();
    assert_eq!(snapshot.layout.rows.len(), 1);
    assert_eq!(snapshot.layout.columns.len(), 1);
    assert_eq!(snapshot.cells[0].source, (2, 1));
    assert_eq!(snapshot.cells[0].text, "visible");
}

#[test]
fn capture_preserves_view_order_and_does_not_follow_later_edits() {
    let mut sheet = Sheet::new(SheetId(1), 4, 2);
    sheet.set_value(0, 0, "first");
    sheet.set_value(1, 0, "second");
    sheet.set_value(2, 0, "third");
    let snapshot = capture(&sheet, &view(&[2, 0], &[0]), None, "IBM Plex Sans", 11.0).unwrap();
    sheet.set_value(2, 0, "changed");
    assert_eq!(
        snapshot
            .cells
            .iter()
            .map(|c| c.text.as_str())
            .collect::<Vec<_>>(),
        ["third", "first"]
    );
}

#[test]
fn selection_expands_to_merge_but_never_includes_hidden_text() {
    let mut sheet = Sheet::new(SheetId(1), 5, 5);
    sheet.set_value(0, 0, "title");
    sheet.add_merge(MergedRegion::new(0, 0, 1, 2)).unwrap();
    let scope = || {
        Some(Scope {
            rows: 1..2,
            columns: 1..2,
        })
    };
    let snapshot = capture(
        &sheet,
        &view(&[0, 1, 2], &[0, 1, 2]),
        scope(),
        "IBM Plex Sans",
        11.0,
    )
    .unwrap();
    assert_eq!(snapshot.layout.rows.len(), 2);
    assert_eq!(snapshot.layout.columns.len(), 3);
    assert_eq!(snapshot.cells.len(), 1);
    assert_eq!(snapshot.cells[0].text, "title");
    let hidden = capture(
        &sheet,
        &view(&[1, 2], &[0, 1, 2]),
        Some(Scope {
            rows: 0..1,
            columns: 0..3,
        }),
        "IBM Plex Sans",
        11.0,
    )
    .unwrap();
    assert!(hidden.cells.iter().all(|c| c.text.is_empty()));
}

#[test]
fn sorted_merge_rejected_instead_of_losing_data() {
    let mut sheet = Sheet::new(SheetId(1), 5, 5);
    sheet.set_value(0, 0, "title");
    sheet.add_merge(MergedRegion::new(0, 0, 1, 2)).unwrap();
    let result = capture(
        &sheet,
        &view(&[0, 2, 1], &[0, 1, 2]),
        None,
        "IBM Plex Sans",
        11.0,
    );
    assert!(result.unwrap_err().contains("sort"));
}

#[test]
fn styled_blank_selection_exports_and_empty_active_sheet_does_not() {
    let sheet = Sheet::new(SheetId(1), 3, 3);
    let v = view(&[0, 1, 2], &[0, 1, 2]);
    assert!(capture(&sheet, &v, None, "IBM Plex Sans", 11.0).is_err());
    assert_eq!(
        capture(
            &sheet,
            &v,
            Some(Scope {
                rows: 0..2,
                columns: 0..2
            }),
            "IBM Plex Sans",
            11.0
        )
        .unwrap()
        .cells
        .len(),
        4
    );
}

#[test]
fn rendered_pdf_reports_clipping_and_is_atomically_saved() {
    let mut sheet = Sheet::new(SheetId(1), 3, 3);
    sheet.set_value(0, 0, "Café € Ελληνικά");
    sheet.set_value(
        1,
        0,
        "A very long wrapped text which cannot fit into one short row.",
    );
    sheet.set_format(
        1,
        0,
        CellFormat {
            text_overflow: TextOverflow::Wrap,
            ..Default::default()
        },
    );
    let snap = capture(&sheet, &view(&[0, 1], &[0]), None, "IBM Plex Sans", 11.0).unwrap();
    let pdf = visigrid_print::pdf::render(&snap, &PageSettings::default()).unwrap();
    assert!(pdf.bytes.starts_with(b"%PDF-"));
    assert!(pdf.clipped_cells > 0);
    assert_eq!(pdf.pages, 1);
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("existing.pdf");
    std::fs::write(&path, b"old content").unwrap();
    visigrid_print::pdf::save_atomic(&path, &pdf.bytes).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), pdf.bytes);
    let failure = visigrid_print::pdf::save_atomic(&dir.path().join("missing/sub.pdf"), &pdf.bytes);
    assert!(failure.is_err());
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    assert!(
        visigrid_print::pdf::render_cancellable(&snap, &PageSettings::default(), || true)
            .unwrap_err()
            .contains("cancelled")
    );
}
