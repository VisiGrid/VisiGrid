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
            for count in [1, 2] {
                let mut wb = Workbook::new();
                let (r, c) = if axis == Axis::Row {
                    (NUM_ROWS - 1, 0)
                } else {
                    (0, NUM_COLS - 1)
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
                    error.contains(if axis == Axis::Row { "1048576" } else { "XFD" }),
                    "{error}"
                );
                assert_eq!(format!("{wb:?}"), before);
            }
        }
    }
}
