use visigrid_engine::{
    cell::CellStyle,
    cond_format::CondStyle,
    sheet::{MergedRegion, NUM_COLS, NUM_ROWS},
    structural::Axis,
    validation::{CellRange, ValidationRule},
    workbook::Workbook,
};

#[test]
fn insertion_refuses_each_edge_obstruction_before_mutation() {
    for axis in [Axis::Row, Axis::Col] {
        for kind in [
            "merged range",
            "validation range",
            "conditional format",
            "non-empty cell",
        ] {
            for count in [2, 3] {
                let mut wb = Workbook::new();
                let (r, c) = if axis == Axis::Row {
                    (NUM_ROWS - 2, 0)
                } else {
                    (0, NUM_COLS - 2)
                };
                let range = CellRange::new(r.saturating_sub(1), c.saturating_sub(1), r, c);
                match kind {
                    "merged range" => wb
                        .active_sheet_mut()
                        .add_merge(MergedRegion::new(range.start_row, range.start_col, r, c))
                        .unwrap(),
                    "validation range" => {
                        wb.active_sheet_mut()
                            .validations
                            .set(range, ValidationRule::custom("=TRUE"));
                    }
                    "conditional format" => {
                        wb.active_sheet_mut().cond_formats.add(
                            vec![range],
                            "=TRUE",
                            CondStyle::Named(CellStyle::Warning),
                        );
                    }
                    _ => {
                        wb.set_cell_value_tracked(0, r, c, "keep");
                    }
                }
                let before = format!("{wb:?}");
                let error = wb.structural_edit(0, axis, 0, count, false).unwrap_err();
                assert!(error.contains(kind), "{error}");
                assert!(
                    error.contains(if axis == Axis::Row { "1048575" } else { "XFC" }),
                    "{error}"
                );
                assert_eq!(format!("{wb:?}"), before);
            }
        }
    }
}

#[test]
fn edge_rules_and_merges_preserve_exact_row_and_column_history() {
    for axis in [Axis::Row, Axis::Col] {
        for edge in [true, false] {
            let mut wb = Workbook::new();
            let last = if axis == Axis::Row { NUM_ROWS - 1 } else { NUM_COLS - 1 };
            let end = last - usize::from(!edge);
            let range = if axis == Axis::Row { CellRange::new(1, 0, end, 0) } else { CellRange::new(0, 1, 0, end) };
            wb.active_sheet_mut().validations.set(range, ValidationRule::custom("=TRUE"));
            wb.active_sheet_mut().cond_formats.add(vec![range], "=TRUE", CondStyle::Named(CellStyle::Warning));
            // Keep merges small, near the edge, to avoid a million-cell index.
            let merge = if axis == Axis::Row { MergedRegion::new(end-2, 2, end, 3) } else { MergedRegion::new(2, end-2, 3, end) };
            wb.active_sheet_mut().add_merge(merge).unwrap();
            let metadata = |w: &Workbook| serde_json::json!([w.active_sheet().validations.iter().collect::<Vec<_>>(), w.active_sheet().validations.exclusions_iter().collect::<Vec<_>>(), w.active_sheet().cond_formats, w.active_sheet().merged_regions]);
            let before = metadata(&wb);
            if axis == Axis::Row {
                let h = wb.prepare_table_row_history(0, 1, 1, false).unwrap().unwrap();
                wb.apply_table_row_history(&h, false).unwrap();
                let after = metadata(&wb);
                wb.apply_table_row_history(&h, true).unwrap();
                assert_eq!(metadata(&wb), before);
                wb.apply_table_row_history(&h, false).unwrap();
                assert_eq!(metadata(&wb), after);
            } else {
                let h = wb.prepare_table_column_history(0, 1, 1, false).unwrap().unwrap();
                wb.apply_table_column_history(&h, false).unwrap();
                let after = metadata(&wb);
                wb.apply_table_column_history(&h, true).unwrap();
                assert_eq!(metadata(&wb), before);
                wb.apply_table_column_history(&h, false).unwrap();
                assert_eq!(metadata(&wb), after);
            }
        }
    }
}

#[test]
fn deleting_a_last_row_rule_round_trips_and_totals_allow_an_edge_merge() {
    let mut wb = Workbook::new();
    let last = NUM_ROWS - 1;
    let range = CellRange::new(last, 0, last, 0);
    wb.active_sheet_mut().validations.set(range, ValidationRule::custom("=TRUE"));
    wb.active_sheet_mut().cond_formats.add(vec![range], "=TRUE", CondStyle::Named(CellStyle::Warning));
    wb.active_sheet_mut().add_merge(MergedRegion::new(last, 2, last, 3)).unwrap();
    let metadata = |w: &Workbook| serde_json::json!([
        w.active_sheet().validations.iter().collect::<Vec<_>>(),
        w.active_sheet().cond_formats,
        w.active_sheet().merged_regions,
    ]);
    let before = metadata(&wb);
    let deleted = wb.prepare_table_row_history(0, last, 1, true).unwrap().unwrap();
    wb.apply_table_row_history(&deleted, false).unwrap();
    assert!(wb.active_sheet().validations.iter().next().is_none());
    assert!(wb.active_sheet().merged_regions.is_empty());
    wb.apply_table_row_history(&deleted, true).unwrap();
    assert_eq!(metadata(&wb), before);

    // A one-cell rule on the last row is clipped by any insert, so the totals
    // case starts clean. The merge already touches the edge and must not.
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 5, "Amount");
    wb.set_cell_value_tracked(0, 1, 5, "10");
    let id = wb.create_table(wb.active_sheet_id(), visigrid_engine::table::TableRange { start_row: 0, start_col: 5, end_row: 1, end_col: 5 }, "Sales").unwrap().table_id();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.active_sheet_mut().add_merge(MergedRegion::new(last - 2, 7, last, 7)).unwrap();
    let inserted = wb.prepare_table_row_history(0, 4, 1, false).unwrap().unwrap();
    wb.apply_table_row_history(&inserted, false).unwrap();
    assert!(wb.active_sheet().merged_regions.iter().any(|m| m.end.0 == last && m.start.1 == 7));
    wb.apply_table_row_history(&inserted, true).unwrap();
    assert!(wb.active_sheet().merged_regions.iter().any(|m| m.start.0 == last - 2 && m.end.0 == last && m.start.1 == 7));
}
