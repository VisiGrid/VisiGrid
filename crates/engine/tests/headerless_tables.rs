use visigrid_engine::{
    cell::{CellComment, CellFormat},
    sheet::{MergedRegion, Sheet, SheetId},
    table::TableRange,
    workbook::Workbook,
};

fn range(sr: usize, sc: usize, er: usize, ec: usize) -> TableRange {
    TableRange {
        start_row: sr,
        start_col: sc,
        end_row: er,
        end_col: ec,
    }
}
fn book() -> Workbook {
    let mut wb = Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 100, 20, "Data")], 0);
    wb.set_cell_value_tracked(0, 2, 1, "2");
    wb.set_cell_value_tracked(0, 2, 2, "=B3*2");
    wb.set_cell_value_tracked(0, 3, 1, "3");
    wb.set_cell_value_tracked(0, 3, 2, "=B4*2");
    wb.set_cell_value_tracked(0, 2, 8, "neighbor");
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=Data!$B$3");
    wb
}

#[test]
fn generated_headers_keep_values_formats_comments_and_cross_sheet_references() {
    let mut wb = book();
    let mut format = CellFormat::default();
    format.bold = true;
    wb.sheet_mut(0).unwrap().set_format(2, 1, format.clone());
    let comment = CellComment {
        text: "keep me".into(),
        author: "QA".into(),
    };
    wb.sheet_mut(0)
        .unwrap()
        .set_comment(2, 1, Some(comment.clone()));
    let rev = wb.revision();
    let preview = wb
        .preview_table_creation(SheetId(1), range(2, 1, 3, 2), false)
        .unwrap();
    assert_eq!(
        preview,
        (range(2, 1, 4, 2), vec!["Column1".into(), "Column2".into()])
    );
    assert_eq!(wb.revision(), rev);
    let commit = wb
        .create_table_without_headers(SheetId(1), range(2, 1, 3, 2), "Sales")
        .unwrap();
    assert_eq!(wb.revision(), rev + 1);
    assert_eq!(commit.inserted_header_row(), Some(2));
    let id = commit.table_id();
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 2);
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 1), "Column1");
    assert_eq!(wb.sheet(0).unwrap().get_raw(3, 1), "2");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "3");
    assert_eq!(wb.sheet(0).unwrap().get_display(3, 2), "4");
    assert_eq!(wb.sheet(0).unwrap().get_raw(3, 8), "neighbor");
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=Data!$B$4");
    assert_eq!(wb.sheet(0).unwrap().get_format(3, 1), format);
    assert_eq!(wb.sheet(0).unwrap().comment(3, 1), Some(&comment));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.revision(), rev + 2);
    assert!(wb.table(id).is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 2), "=B3*2");
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 8), "neighbor");
    assert_eq!(wb.sheet(1).unwrap().get_raw(0, 0), "=Data!$B$3");
    assert_eq!(wb.sheet(0).unwrap().comment(2, 1), Some(&comment));
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.revision(), rev + 3);
    assert_eq!(wb.table(id).unwrap().1.range, preview.0);
    assert_eq!(wb.sheet(0).unwrap().get_format(3, 1), format);
}

#[test]
fn adjacent_calculated_table_grows_and_its_generated_row_is_removed_on_undo() {
    let mut wb = book();
    let other = wb
        .create_table(SheetId(1), range(0, 5, 5, 6), "Other")
        .unwrap()
        .table_id();
    wb.set_calculated_column(other, 6, 1, "=F2*3", true)
        .unwrap();
    wb.clear_cell_tracked(0, 3, 6);
    let schema = wb.table(other).unwrap().1.clone();
    let commit = wb
        .create_table_without_headers(SheetId(1), range(2, 1, 3, 2), "Sales")
        .unwrap();
    assert_eq!(wb.table(other).unwrap().1.range.end_row, 6);
    assert!(!wb.sheet(0).unwrap().is_calculated_exception(2, 6));
    assert!(wb.sheet(0).unwrap().is_calculated_exception(4, 6));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.table(other).unwrap().1, &schema);
    assert!(wb.sheet(0).unwrap().is_calculated_exception(3, 6));
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.table(other).unwrap().1.range.end_row, 6);
}

#[test]
fn first_row_and_single_record_creation_preserve_blank_records() {
    let mut wb = book();
    for data in [range(0, 0, 0, 0), range(0, 0, 1, 0)] {
        let commit = wb
            .create_table_without_headers(SheetId(1), data, "Blank")
            .unwrap();
        assert_eq!(
            wb.table(commit.table_id()).unwrap().1.range.data_rows(),
            data.data_rows() + 1
        );
        assert_eq!(wb.sheet(0).unwrap().get_raw(0, 0), "Column1");
        assert_eq!(wb.sheet(0).unwrap().get_raw(1, 0), "");
        wb.apply_table_commit(&commit, true).unwrap();
    }
}

#[test]
fn preflight_refusals_do_not_insert_a_row_or_consume_an_identity() {
    for failure in 0..7 {
        let mut wb = book();
        match failure {
            0 => {
                wb.set_cell_value_tracked(0, 99, 19, "edge");
            }
            1 => {
                wb.create_table(SheetId(1), range(2, 1, 3, 2), "Other")
                    .unwrap();
            }
            2 => {
                wb.sheet_mut(0)
                    .unwrap()
                    .add_merge(MergedRegion {
                        start: (2, 1),
                        end: (2, 2),
                    })
                    .unwrap();
            }
            3 => {
                wb.clear_cell_tracked(0, 3, 1);
                wb.set_cell_value_tracked(0, 2, 1, "=SEQUENCE(2)");
                assert!(wb.sheet(0).unwrap().get_cell(2, 1).spill_info().is_some());
            }
            5 => {
                wb.sheet_mut(0).unwrap().set_comment(
                    99,
                    19,
                    Some(CellComment {
                        text: "edge note".into(),
                        author: "QA".into(),
                    }),
                );
            }
            6 => {
                let mut format = CellFormat::default();
                format.bold = true;
                wb.sheet_mut(0).unwrap().set_format(99, 19, format);
            }
            _ => {}
        }
        let rev = wb.revision();
        let catalog = serde_json::to_string(&wb.saved_tables()).unwrap();
        assert!(
            wb.create_table_without_headers(
                SheetId(1),
                range(2, 1, 3, 2),
                if failure == 4 { "A1" } else { "Sales" }
            )
            .is_err(),
            "failure case {failure}"
        );
        assert_eq!(wb.revision(), rev);
        assert_eq!(serde_json::to_string(&wb.saved_tables()).unwrap(), catalog);
        assert_eq!(wb.sheet(0).unwrap().get_raw(2, 8), "neighbor");
    }
    let wb = book();
    assert!(wb
        .preview_table_creation(SheetId(1), range(99, 0, 99, 0), false)
        .is_err());
}

#[test]
fn a_spill_created_by_moving_a_formula_rolls_back_the_entire_operation() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 1, 0, "=IF(ROW()=3,SEQUENCE(2),1)");
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 0), "1");
    let rev = wb.revision();
    let catalog = serde_json::to_string(&wb.saved_tables()).unwrap();
    assert!(wb
        .create_table_without_headers(SheetId(1), range(1, 0, 1, 0), "Sales")
        .is_err());
    assert_eq!(wb.revision(), rev);
    assert_eq!(serde_json::to_string(&wb.saved_tables()).unwrap(), catalog);
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 0), "1");
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 8), "neighbor");
}

#[test]
fn stale_replay_refuses_atomically_including_changes_outside_the_new_table() {
    for failure in 0..3 {
        let mut wb = book();
        let commit = wb
            .create_table_without_headers(SheetId(1), range(2, 1, 3, 2), "Sales")
            .unwrap();
        match failure {
            0 => {
                wb.set_cell_value_tracked(0, 2, 9, "new note");
            }
            1 => {
                wb.sheet_mut(0).unwrap().set_comment(
                    2,
                    9,
                    Some(CellComment {
                        text: "note".into(),
                        author: "QA".into(),
                    }),
                );
            }
            _ => {
                wb.set_cell_value_tracked(1, 0, 0, "=123");
            }
        }
        let rev = wb.revision();
        assert!(wb.apply_table_commit(&commit, true).is_err());
        assert_eq!(wb.revision(), rev);
        assert!(wb.table(commit.table_id()).is_some());
        assert_eq!(wb.sheet(0).unwrap().get_raw(3, 8), "neighbor");
    }
    let mut wb = book();
    let commit = wb
        .create_table_without_headers(SheetId(1), range(2, 1, 3, 2), "Sales")
        .unwrap();
    wb.apply_table_commit(&commit, true).unwrap();
    wb.create_table(SheetId(1), range(10, 5, 12, 6), "Sales")
        .unwrap();
    let rev = wb.revision();
    assert!(wb.apply_table_commit(&commit, false).is_err());
    assert_eq!(wb.revision(), rev);
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 8), "neighbor");
}

#[test]
fn creation_binds_existing_structured_references_and_undo_restores_unbound_sources() {
    let mut wb = book();
    wb.set_cell_value_tracked(1, 1, 0, "=SUM(Sales[Column1])+Data!B3");
    let before = wb.sheet(1).unwrap().get_raw(1, 0);
    let commit = wb
        .create_table_without_headers(SheetId(1), range(2, 1, 3, 2), "Sales")
        .unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_display(1, 0), "7");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_raw(1, 0), before);
    assert!(wb.sheet(1).unwrap().get_display(1, 0).starts_with('#'));
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(1).unwrap().get_display(1, 0), "7");
}
