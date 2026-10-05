use visigrid_engine::{
    cell::NumberFormat,
    filter::SortDirection,
    sheet::MergedRegion,
    table::{TableId, TableRange},
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};

fn book(hidden: &[usize]) -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (row, values) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let (wb, _) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), hidden.iter().copied().collect())
        .unwrap();
    (wb, id)
}

#[test]
fn hidden_footer_moves_keep_row_positions_and_subtotal_semantics() {
    for hidden in [vec![4], vec![6], vec![4, 6], vec![2, 4, 6, 20]] {
        let (mut wb, id) = book(&hidden);
        wb.set_cell_value_tracked(0, 0, 4, "=B5");
        wb.set_cell_value_tracked(0, 1, 4, "=SUBTOTAL(9,Sales[Amount])");
        wb.set_cell_value_tracked(0, 2, 4, "=SUM(Sales[Amount])");
        wb.set_cell_value_tracked(0, 4, 3, "Beside old footer");
        wb.set_cell_value_tracked(0, 6, 3, "Beside new footer");
        let mut footer = wb.active_sheet().get_cell(4, 1);
        footer.set_comment(Some(visigrid_engine::cell::CellComment {
            text: "Keep total note".into(),
            author: "QA".into(),
        }));
        std::sync::Arc::make_mut(&mut footer.format).number_format =
            NumberFormat::Percent { decimals: 1 };
        wb.active_sheet_mut()
            .set_format(4, 1, (*footer.format).clone());
        wb.active_sheet_mut()
            .set_comment(4, 1, footer.comment().cloned());
        let original = wb.clone();
        let commit = wb
            .append_table_rows(id, 2, &[(4, 1, "7".into()), (5, 1, "11".into())])
            .unwrap();
        let expected =
            78 - if hidden.contains(&4) { 7 } else { 0 } - if hidden.contains(&2) { 20 } else { 0 };
        assert_eq!(wb.active_sheet().get_display(0, 4), expected.to_string());
        assert_eq!(wb.active_sheet().get_display(1, 4), "78");
        assert_eq!(wb.active_sheet().get_display(2, 4), "78");
        assert_eq!(wb.active_sheet().get_raw(0, 4), "=B7");
        assert_eq!(
            wb.active_sheet().manual_hidden_rows(),
            hidden.iter().copied().collect()
        );
        assert_eq!(wb.active_sheet().get_cell(6, 1).format, footer.format);
        assert_eq!(wb.active_sheet().get_cell(6, 1).comment(), footer.comment());
        assert_eq!(wb.active_sheet().get_raw(4, 3), "Beside old footer");
        assert_eq!(wb.active_sheet().get_raw(6, 3), "Beside new footer");
        wb.apply_table_commit(&commit, true).unwrap();
        assert_eq!(
            wb.active_sheet().manual_hidden_rows(),
            original.active_sheet().manual_hidden_rows()
        );
        assert_eq!(wb.active_sheet().get_raw(0, 4), "=B5");
        assert_eq!(wb.active_sheet().get_cell(4, 1).format, footer.format);
        assert_eq!(wb.active_sheet().get_raw(5, 1), "");
        wb.apply_table_commit(&commit, false).unwrap();
        assert_eq!(wb.active_sheet().get_display(0, 4), expected.to_string());
    }
}

#[test]
fn shrinking_and_compound_resize_preserve_hidden_rows_and_released_records() {
    let (mut wb, id) = book(&[2, 4, 6]);
    wb.clear_cell_tracked(0, 2, 0);
    wb.clear_cell_tracked(0, 2, 1);
    let range = wb.table(id).unwrap().1.range;
    let shrink = wb
        .resize_table(
            id,
            TableRange {
                end_row: 1,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(2));
    assert_eq!(wb.active_sheet().get_display(2, 1), "10");
    assert_eq!(wb.active_sheet().get_raw(3, 1), "30");
    assert_eq!(wb.active_sheet().get_raw(4, 1), "");
    assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4, 6].into());
    wb.apply_table_commit(&shrink, true).unwrap();
    assert_eq!(wb.active_sheet().get_display(4, 1), "40");
    wb.apply_table_commit(&shrink, false).unwrap();
    wb.apply_table_commit(&shrink, true).unwrap();
    wb.set_cell_value_tracked(0, 4, 2, "Included side record");
    let wide = wb
        .resize_table(
            id,
            TableRange {
                end_row: 5,
                end_col: 2,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(6));
    assert_eq!(wb.active_sheet().get_raw(4, 2), "Included side record");
    assert_eq!(wb.active_sheet().manual_hidden_rows(), [2, 4, 6].into());
    wb.apply_table_commit(&wide, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range, range);
    wb.apply_table_commit(&wide, false).unwrap();
    assert_eq!(wb.active_sheet().get_display(6, 1), "40");
}

#[test]
fn sparse_move_replay_rejects_changed_visibility_and_keeps_saved_sort() {
    let (mut wb, id) = book(&[4, 5]);
    let mut spec = TableViewSpec::new(id);
    spec.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[1].id,
        direction: SortDirection::Descending,
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec.clone()))
        .unwrap();
    let commit = wb
        .append_table_rows(id, 1, &[(4, 1, "100".into())])
        .unwrap();
    assert_eq!(wb.active_sheet().get_display(5, 1), "60");
    let view = wb
        .active_sheet()
        .build_saved_table_view(30)
        .unwrap()
        .unwrap();
    assert!(view.rows().data_to_view(4).is_none());
    assert!(view.rows().data_to_view(5).is_none());
    let (mut changed, visibility) = wb
        .prepare_table_row_visibility(wb.active_sheet_id(), [5].into())
        .unwrap();
    assert_eq!(changed.active_sheet().get_display(5, 1), "160");
    let revision = changed.revision();
    assert!(changed.apply_table_commit(&commit, true).is_err());
    assert_eq!(changed.revision(), revision);
    visibility.replay(&mut changed, true).unwrap();
    changed.apply_table_commit(&commit, true).unwrap();
    assert_eq!(changed.active_sheet().table_view_spec(), Some(&spec));
    changed.apply_table_commit(&commit, false).unwrap();
    assert_eq!(changed.active_sheet().manual_hidden_rows(), [4, 5].into());
}

#[test]
fn hidden_destinations_still_refuse_occupied_and_merged_cells_atomically() {
    for merged in [false, true] {
        let (mut wb, id) = book(&[4, 5]);
        if merged {
            wb.active_sheet_mut()
                .add_merge(MergedRegion::new(5, 0, 5, 1))
                .unwrap();
        } else {
            wb.set_cell_value_tracked(0, 5, 1, "Keep hidden data");
        }
        let revision = wb.revision();
        assert!(wb.append_table_rows(id, 1, &[(4, 1, "77".into())]).is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
        assert_eq!(wb.active_sheet().get_display(4, 1), "60");
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [4, 5].into());
    }
}
