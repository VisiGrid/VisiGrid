//! Large sheets as row bands: a banded export, loaded band by band in any
//! order, is the same workbook as the inline one.
use visigrid_engine::cell::CellFormat;
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::{bands, export_workbook, import_any, SheetLayout};

fn big() -> Workbook {
    let mut s = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    s.set_name("Data");
    let bold = CellFormat { bold: true, ..Default::default() };
    // 150,000 rows x 2 = 300,000 cells: over the threshold, three bands.
    for r in 0..150_000 {
        s.set_value_deferred(r, 0, &format!("{}", r % 997));
        s.set_value_deferred(r, 1, &format!("t{r}"));
    }
    s.set_format(70_000, 1, bold);
    // Formulas reaching across bands, and one that spills.
    s.set_value_deferred(0, 3, "=SUM(A1:A150000)");
    s.set_value_deferred(140_000, 3, "=A1+A70001");
    s.set_value_deferred(5, 5, "=SEQUENCE(3)");
    let small = {
        let mut t = Sheet::new(SheetId(2), NUM_ROWS, NUM_COLS);
        t.set_name("Small");
        t.set_value_deferred(0, 0, "=Data!D1");
        t
    };
    let mut wb = Workbook::from_sheets(vec![s, small], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

fn shown(wb: &Workbook, sheet: usize, r: usize, c: usize) -> (String, String, bool) {
    let s = &wb.sheets()[sheet];
    (s.get_raw(r, c), s.get_formatted_display(r, c), s.get_format(r, c).bold)
}

#[test]
fn banded_export_round_trips_in_any_band_order() {
    let wb = big();
    let layouts = vec![SheetLayout::default(), SheetLayout::default()];
    let (manifest, out) = bands::export_banded(&wb, &layouts, 0).unwrap();
    assert_eq!(out.len(), 10, "150,000 rows make ten 16,384-row bands");
    assert!(out.iter().all(|b| b.sheet == 0), "the small sheet stays inline");
    assert!(manifest.len() < 4096, "the manifest carries no cells: {} bytes", manifest.len());
    let refs = bands::manifest_bands(&manifest).unwrap();
    assert_eq!(refs.len(), 10);

    let (mut loaded, _, _) = import_any(&manifest).unwrap();
    assert_eq!(loaded.sheets()[0].cells_iter().count(), 0);
    for b in out.iter().rev() {
        bands::apply(&mut loaded, &b.data, Some(&b.reference.key)).unwrap();
    }
    bands::finish(&mut loaded);

    for &(sheet, r, c) in &[(0, 0, 0), (0, 70_000, 1), (0, 149_999, 1), (0, 0, 3), (0, 140_000, 3), (0, 5, 5), (0, 7, 5), (1, 0, 0)] {
        assert_eq!(shown(&loaded, sheet, r, c), shown(&wb, sheet, r, c), "cell {sheet}:{r}:{c}");
    }
    assert_eq!(loaded.sheets()[0].cells_iter().count(), wb.sheets()[0].cells_iter().count());
    // Inline and banded agree as documents too.
    assert_eq!(export_workbook(&loaded, &layouts, 0).unwrap(), export_workbook(&wb, &layouts, 0).unwrap());
}

#[test]
fn a_band_is_checked_against_its_key() {
    let wb = big();
    let (_, out) = bands::export_banded(&wb, &[SheetLayout::default(), SheetLayout::default()], 0).unwrap();
    let mut other = Workbook::new();
    let err = bands::apply(&mut other, &out[0].data, Some(&out[1].reference.key)).unwrap_err();
    assert!(err.contains("does not match"), "{err}");
}

#[test]
fn small_workbooks_are_not_banded() {
    let mut wb = Workbook::new();
    wb.sheet_mut(0).unwrap().set_value(0, 0, "1");
    let (manifest, out) = bands::export_banded(&wb, &[SheetLayout::default()], 0).unwrap();
    assert!(out.is_empty());
    let v = |s: &str| serde_json::from_str::<serde_json::Value>(s).unwrap();
    assert_eq!(v(&manifest), v(&export_workbook(&wb, &[SheetLayout::default()], 0).unwrap()));
}
