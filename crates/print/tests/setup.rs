#![cfg(feature = "pdf")]
use visigrid_engine::{
    print_setup::*,
    sheet::{MergedRegion, Sheet, SheetId},
};
use visigrid_print::{
    paginate, setup,
    snapshot::{capture, SheetView},
    AxisItem,
};
fn view(rows: impl Iterator<Item = usize>) -> SheetView {
    SheetView {
        rows: rows
            .map(|source_index| AxisItem {
                source_index,
                size_pt: 25.0,
            })
            .collect(),
        columns: (0..4)
            .map(|source_index| AxisItem {
                source_index,
                size_pt: 80.0,
            })
            .collect(),
    }
}
#[test]
fn headers_repeat_on_every_page_and_hidden_rows_stay_hidden() {
    let mut sheet = Sheet::new(SheetId(1), 120, 4);
    sheet.set_value(0, 0, "Header");
    sheet.set_value(100, 3, "last");
    let snapshot = capture(
        &sheet,
        &view((0..101).filter(|&r| r != 1)),
        None,
        "IBM Plex Sans",
        11.0,
    )
    .unwrap();
    let saved = PrintSetup {
        repeat_rows: Some(PrintRows { start: 0, end: 2 }),
        ..Default::default()
    };
    let settings = setup::page_settings(&snapshot, &saved).unwrap();
    assert_eq!(settings.repeat_rows, 2);
    let plan = paginate(&snapshot.layout, &settings).unwrap();
    assert!(plan.pages().len() > 1);
    for i in 0..plan.pages().len() {
        assert!(plan.rect(i, 0..1, 0..1).is_some());
        assert!(plan.rect(i, 1..2, 0..1).is_some());
    }
}
#[test]
fn area_includes_blank_cells_expands_merges_and_excludes_other_sources_when_sorted() {
    let mut sheet = Sheet::new(SheetId(1), 20, 4);
    sheet.add_merge(MergedRegion::new(3, 0, 3, 2)).unwrap();
    sheet.set_value(3, 0, "merged");
    sheet.set_value(9, 0, "outside");
    let area = PrintArea {
        start_row: 3,
        end_row: 5,
        start_col: 1,
        end_col: 1,
    };
    let snapshot = setup::capture_area(
        &sheet,
        &view([9, 3, 5, 4, 0].into_iter()),
        area,
        "IBM Plex Sans",
        11.0,
    )
    .unwrap();
    assert_eq!(
        snapshot
            .layout
            .rows
            .iter()
            .map(|r| r.source_index)
            .collect::<Vec<_>>(),
        vec![3, 5, 4]
    );
    assert_eq!(snapshot.layout.columns.len(), 3);
    assert!(snapshot.cells.iter().any(|c| c.text == "merged"));
    assert!(!snapshot.cells.iter().any(|c| c.text == "outside"));
    let saved = PrintSetup {
        repeat_rows: Some(PrintRows { start: 3, end: 4 }),
        ..Default::default()
    };
    assert!(
        setup::page_settings(&snapshot, &saved).is_err(),
        "discontiguous sorted headers cannot silently repeat the wrong rows"
    );
}
#[test]
fn header_boundary_cannot_split_merge_or_consume_whole_print_area() {
    let mut sheet = Sheet::new(SheetId(1), 20, 4);
    sheet.add_merge(MergedRegion::new(0, 0, 1, 0)).unwrap();
    sheet.set_value(0, 0, "Title");
    sheet.set_value(10, 3, "last");
    let snapshot = capture(&sheet, &view(0..20), None, "IBM Plex Sans", 11.0).unwrap();
    for end in [0, 10] {
        let saved = PrintSetup {
            repeat_rows: Some(PrintRows { start: 0, end }),
            ..Default::default()
        };
        let settings = setup::page_settings(&snapshot, &saved).unwrap();
        assert!(paginate(&snapshot.layout, &settings).is_err());
    }
}
