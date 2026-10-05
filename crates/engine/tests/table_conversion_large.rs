use visigrid_engine::{
    cell::CellComment,
    sheet::{Sheet, SheetId},
    table::TableRange,
    workbook::Workbook,
};

#[test]
fn mixed_formula_and_full_cell_history_rebuilds_spills_and_dependencies() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(2)");
    wb.set_cell_value_tracked(0, 0, 2, "before");
    wb.set_cell_value_tracked(0, 0, 4, "=SUM(A1:A4)");
    let mut candidate = wb.clone();
    candidate.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(4)");
    candidate.set_cell_value_tracked(0, 0, 2, "=A4*2");
    candidate.active_sheet_mut().set_comment(
        0,
        2,
        Some(CellComment {
            text: "new formula".into(),
            author: "tester".into(),
        }),
    );
    let history = wb.capture_guarded_batch(&candidate).unwrap();
    history.replay(&mut wb, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(3, 0), "4");
    assert_eq!(wb.active_sheet().get_display(0, 2), "8");
    assert_eq!(wb.active_sheet().get_display(0, 4), "10");
    assert!(wb.active_sheet().is_spill_receiver(3, 0));
    history.replay(&mut wb, true).unwrap();
    assert_eq!(wb.active_sheet().get_display(1, 0), "2");
    assert!(!wb.active_sheet().is_spill_receiver(3, 0));
    assert_eq!(wb.active_sheet().get_display(3, 0), "");
    assert_eq!(wb.active_sheet().get_raw(0, 2), "before");
    assert!(wb.active_sheet().comment(0, 2).is_none());
    assert_eq!(wb.active_sheet().get_display(0, 4), "3");
    history.replay(&mut wb, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(0, 2), "8");
    assert_eq!(wb.active_sheet().get_display(0, 4), "10");
    assert_eq!(wb.active_sheet().comment(0, 2).unwrap().text, "new formula");
}

#[test]
fn large_conversion_round_trips_beyond_the_ordinary_batch_limit() {
    let rows = 50_001;
    let mut sheet = Sheet::new(SheetId(1), rows + 2, 4);
    for (col, name) in ["Value", "Double", "PlusOne"].iter().enumerate() {
        sheet.set_value_deferred(0, col, name);
    }
    for row in 1..=rows {
        sheet.set_value_deferred(row, 0, "10");
        sheet.set_value_deferred(row, 1, "=[@Value]*2");
        sheet.set_value_deferred(row, 2, "=[@Value]+1");
    }
    sheet.set_comment(
        rows,
        1,
        Some(CellComment {
            text: "last record".into(),
            author: "tester".into(),
        }),
    );
    let mut wb = Workbook::from_sheets(vec![sheet], 0);
    let id = wb
        .create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                end_row: rows,
                start_col: 0,
                end_col: 2,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let original = wb.clone();
    let history = wb.remove_table(id).unwrap();
    assert!(wb.table(id).is_none());
    for row in [1, rows / 2, rows] {
        assert!(!wb.active_sheet().get_raw(row, 1).contains('@'));
        assert_eq!(wb.active_sheet().get_display(row, 1), "20");
        assert_eq!(wb.active_sheet().get_display(row, 2), "11");
    }
    // Ordinary guarded batches keep their existing bound. Only conversion,
    // which legitimately rewrites every calculated record, is exempt.
    assert!(original
        .capture_guarded_batch(&wb)
        .unwrap_err()
        .contains("100,000"));
    wb.apply_table_commit(&history, true).unwrap();
    assert!(wb.table(id).is_some());
    assert_eq!(wb.active_sheet().get_raw(rows, 1), "=[@Value]*2");
    wb.apply_table_commit(&history, false).unwrap();
    assert_eq!(
        wb.active_sheet().comment(rows, 1).unwrap().text,
        "last record"
    );
    assert_eq!(wb.active_sheet().get_display(rows, 2), "11");
    // A source-only history patch must still reject altered presentation.
    wb.active_sheet_mut().set_comment(
        rows,
        1,
        Some(CellComment {
            text: "later edit".into(),
            author: "tester".into(),
        }),
    );
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&history, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert!(wb.table(id).is_none());
    assert_eq!(
        wb.active_sheet().comment(rows, 1).unwrap().text,
        "later edit"
    );
}
