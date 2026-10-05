//! Move a footer within its columns, without inserting worksheet rows.
use super::{HeaderCell, TableCommit, Workbook};
use crate::{
    cell::Cell,
    sheet::SheetId,
    table::{DataTable, TableRange},
};

#[derive(Debug, Clone)]
pub(super) struct FooterCell {
    pub row: usize,
    pub col: usize,
    pub before: Cell,
    pub after: Cell,
    pub before_present: bool,
    pub after_present: bool,
}

#[derive(Debug, Clone)]
pub(super) struct FooterMove {
    pub before_row: usize,
    pub after_row: usize,
    pub cells: Vec<FooterCell>,
    references_guarded: bool,
}

impl FooterMove {
    pub fn owns(&self, row: usize, col: usize) -> bool {
        (row == self.before_row || row == self.after_row)
            && self
                .cells
                .first()
                .zip(self.cells.last())
                .is_some_and(|(first, last)| col >= first.col && col <= last.col)
    }
    pub fn values(&self) -> Vec<(HeaderCell, HeaderCell)> {
        self.cells
            .iter()
            .map(|c| {
                (
                    HeaderCell {
                        row: c.row,
                        col: c.col,
                        value: c.before.value.clone(),
                    },
                    HeaderCell {
                        row: c.row,
                        col: c.col,
                        value: c.after.value.clone(),
                    },
                )
            })
            .collect()
    }
}

impl Workbook {
    fn validate_footer_geometry(
        &self,
        sheet_id: SheetId,
        table: &DataTable,
        rows: [usize; 2],
    ) -> Result<(), String> {
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        for row in rows {
            self.validate_table_region(
                sheet_id,
                TableRange {
                    start_row: row,
                    end_row: row,
                    ..table.range
                },
                Some(table.id),
            )?;
            if table
                .totals
                .as_ref()
                .is_some_and(|t| t.hidden_rows.contains(&row))
            {
                return Err(
                    "The totals row cannot move from or into a manually hidden row.".into(),
                );
            }
            for col in table.range.start_col..=table.range.end_col {
                if sheet.has_validation(row, col) || sheet.cond_formats.any_rule_covers(row, col) {
                    return Err("Clear conditional formatting or validation on the old and new footer cells before moving totals.".into());
                }
            }
        }
        Ok(())
    }

    pub(super) fn prepare_footer_move(
        &self,
        sheet_id: SheetId,
        old: &DataTable,
        new: &DataTable,
    ) -> Result<Option<FooterMove>, String> {
        let (Some(before_row), Some(after_row)) = (old.totals_row(), new.totals_row()) else {
            return Ok(None);
        };
        if before_row == after_row {
            return Ok(None);
        }
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        new.validate(sheet.rows, sheet.cols)?;
        self.validate_footer_geometry(sheet_id, old, [before_row, after_row])?;
        self.validate_empty_table_append(sheet_id, TableRange { start_row: after_row, end_row: after_row, ..new.range })
            .map_err(|_| "The new totals row contains data or comments. Clear that row before resizing; existing records are never overwritten by totals.".to_string())?;
        let mut cells = Vec::with_capacity(old.columns.len() * 2);
        for col in old.range.start_col..=old.range.end_col {
            let source = sheet.get_cell(before_row, col);
            let mut released = sheet.empty_table_footer_cell(before_row, col);
            if new.range.contains(before_row, col) && old.range.data_rows() > 0 {
                released.format = std::sync::Arc::new(sheet.get_format(old.range.end_row, col));
            }
            cells.push(FooterCell {
                row: before_row,
                col,
                before: source.clone(),
                after: released,
                before_present: sheet.get_cell_opt(before_row, col).is_some(),
                after_present: new.range.contains(before_row, col),
            });
            cells.push(FooterCell {
                row: after_row,
                col,
                before: sheet.get_cell(after_row, col),
                after: source,
                before_present: sheet.get_cell_opt(after_row, col).is_some(),
                after_present: true,
            });
        }
        let references_guarded = self.clone().relocate_footer_references(old.id, new.range.end_row, None)?;
        Ok(Some(FooterMove {
            references_guarded,
            before_row,
            after_row,
            cells,
        }))
    }

    pub(super) fn validate_footer_move_replay(
        &self,
        commit: &TableCommit,
        movement: &FooterMove,
        undo: bool,
    ) -> Result<(), String> {
        let table = if undo {
            commit.after_table()
        } else {
            commit.before_table()
        }
        .unwrap();
        self.validate_footer_geometry(
            commit.sheet_id,
            table,
            [movement.before_row, movement.after_row],
        )?;
        if !movement.references_guarded {
            let target = if undo { commit.before_table() } else { commit.after_table() }.unwrap();
            if self.clone().relocate_footer_references(commit.table_id(), target.range.end_row, Some(commit))? {
                return Err("New footer references require a fresh Table operation.".into());
            }
        }
        let sheet = self.sheet_by_id(commit.sheet_id).unwrap();
        for patch in &movement.cells {
            let expected = if undo { &patch.after } else { &patch.before };
            let current = sheet.get_cell(patch.row, patch.col);
            if current.format != expected.format
                || current.comment() != expected.comment()
                || current.style_id() != expected.style_id()
                || current.frozen_formula() != expected.frozen_formula()
            {
                return Err(
                    "Footer formatting or comments changed since this operation was prepared."
                        .into(),
                );
            }
        }
        Ok(())
    }
}
