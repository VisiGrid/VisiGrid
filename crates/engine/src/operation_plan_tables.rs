//! Table guards for normalized Review operations, expressed in source coordinates.
use super::{PlanError, PlannedOp, PlannedOperation};
use crate::{
    structural::Axis,
    validation::CellRange,
    workbook::{shift_structure_index, StructureStep, Workbook},
};

pub(super) fn steps(ops: &[PlannedOperation]) -> Vec<StructureStep> {
    ops.iter()
        .filter_map(|p| match p.operation {
            PlannedOp::DeleteRows { at, count } => Some(StructureStep {
                axis: Axis::Row,
                at,
                count,
                delete: true,
            }),
            _ => None,
        })
        .collect()
}

pub(super) fn validate_targets(
    wb: &Workbook,
    index: usize,
    ops: &[PlannedOperation],
    after: bool,
) -> Result<(), PlanError> {
    let steps = steps(ops);
    let rows = wb.sheet(index).ok_or(PlanError::SourceSheetMissing)?.rows;
    let mut touched = 0usize;
    for planned in ops {
        let (range, values) = match &planned.operation {
            PlannedOp::SetCellValue { coordinate, .. }
            | PlannedOp::SetCellFormula { coordinate, .. }
            | PlannedOp::ClearCell { coordinate } => (
                CellRange {
                    start_row: coordinate.row,
                    end_row: coordinate.row,
                    start_col: coordinate.col,
                    end_col: coordinate.col,
                },
                true,
            ),
            PlannedOp::SetCellStyle { range, .. } => (
                CellRange {
                    start_row: range.start.row,
                    end_row: range.end.row,
                    start_col: range.start.col,
                    end_col: range.end.col,
                },
                false,
            ),
            _ => continue,
        };
        touched = touched.saturating_add(
            (range.end_row - range.start_row + 1)
                .saturating_mul(range.end_col - range.start_col + 1),
        );
        if touched > 100_000 {
            return Err(PlanError::InvalidOperation(
                "A reviewed Table batch may address at most 100,000 cells.".into(),
            ));
        }
        if !after {
            wb.validate_automation_range(index, range, values)
                .map_err(PlanError::InvalidOperation)?;
        } else {
            for source_row in range.start_row..=range.end_row {
                let row = steps
                    .iter()
                    .try_fold(source_row, |r, step| shift_structure_index(r, *step, rows));
                if let Some(row) = row {
                    wb.validate_automation_range(
                        index,
                        CellRange {
                            start_row: row,
                            end_row: row,
                            ..range
                        },
                        values,
                    )
                    .map_err(PlanError::InvalidOperation)?;
                }
            }
        }
    }
    Ok(())
}
