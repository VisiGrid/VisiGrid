//! Canonical-coordinate automation guards shared by desktop and headless hosts.
use super::Workbook;
use crate::validation::CellRange;

impl Workbook {
    pub fn has_table_criteria(&self) -> bool {
        self.sheets().iter().any(|s| {
            s.table_view_spec()
                .is_some_and(|v| v.sort.is_some() || !v.filters.is_empty())
        })
    }

    pub fn validate_saved_table_views(&self) -> Result<(), String> {
        super::guarded_structure::validate_views(self)
    }

    /// Automation explicitly addresses stored records, including hidden ones.
    /// Never translate these coordinates through a display permutation.
    pub fn validate_automation_range(
        &self,
        index: usize,
        range: CellRange,
        values: bool,
    ) -> Result<(), String> {
        self.ensure_writable()?;
        let sheet = self
            .sheet(index)
            .ok_or("The script sheet no longer exists.")?;
        if range.start_row > range.end_row
            || range.start_col > range.end_col
            || range.end_row >= sheet.rows
            || range.end_col >= sheet.cols
        {
            return Err("The operation is outside the worksheet.".into());
        }
        let count = (range.end_row - range.start_row + 1)
            .checked_mul(range.end_col - range.start_col + 1)
            .ok_or("Operation is too large.")?;
        if count > 100_000 {
            return Err(
                "An operation with Table criteria may target at most 100,000 cells.".into(),
            );
        }
        if let Some(spec) = sheet
            .table_view_spec()
            .filter(|spec| spec.sort.is_some() || !spec.filters.is_empty())
        {
            let table = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The saved Table no longer exists.")?;
            let body = table.range;
            if body.data_rows() > 0
                && range.end_row > body.start_row
                && range.start_row <= body.end_row
                && (range.start_col < body.start_col || range.end_col > body.end_col)
            {
                return Err(
                    "Clear Table criteria before changing adjacent cells in its body rows.".into(),
                );
            }
        }
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                if values {
                    if let Some(error) = sheet.table_value_write_error(row, col) {
                        return Err(error);
                    }
                }
                if sheet.is_pivot_owned(row, col) {
                    return Err("Cannot change PivotTable output.".into());
                }
                if sheet.get_spill_parent(row, col).is_some() {
                    return Err("Cannot change a spill receiver; edit its source formula.".into());
                }
                if sheet
                    .get_merge(row, col)
                    .is_some_and(|m| m.start != (row, col))
                {
                    return Err("Cannot change a covered merged cell.".into());
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        sheet::{Sheet, SheetId},
        table::TableRange,
        table_view::TableViewSpec,
    };

    #[test]
    fn buttons_without_criteria_do_not_guard_adjacent_cells() {
        let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 30, 8)], 0);
        wb.set_cell_value_tracked(0, 2, 1, "Amount");
        wb.set_cell_value_tracked(0, 3, 1, "10");
        let id = wb
            .create_table(
                SheetId(7),
                TableRange {
                    start_row: 2,
                    end_row: 3,
                    start_col: 1,
                    end_col: 1,
                },
                "Sales",
            )
            .unwrap()
            .table_id();
        wb.set_table_view_spec(SheetId(7), Some(TableViewSpec::new(id)))
            .unwrap();
        assert!(wb
            .validate_automation_range(
                0,
                CellRange {
                    start_row: 3,
                    end_row: 3,
                    start_col: 3,
                    end_col: 3,
                },
                true
            )
            .is_ok());
    }
}
