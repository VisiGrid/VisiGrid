//! Headerless creation composes a whole-row insertion and an ordinary Table
//! commit. Stage on the engine's copy-on-write workbook so either both succeed
//! or neither changes the document. History retains only sparse metadata.
use super::{TableCommit, TableRowHistory};
use crate::{
    cell::CellFormat, cell_id::CellId, print_setup::PrintSetup, sheet::SheetId, structural::Axis,
    table::TableRange, workbook::Workbook,
};

#[derive(Debug, Clone, serde::Serialize)]
pub(super) struct HeaderInsertion {
    pub at: usize,
    rows: Option<TableRowHistory>,
    guarded: Option<Box<crate::workbook::GuardedStructureCommit>>,
    print_before: PrintSetup,
    inserted_cells: Vec<(usize, usize, String, CellFormat)>,
    formulas: Vec<(CellId, String, String)>,
}

impl Workbook {
    /// Mutation-free preview. Headerless ranges refer to selected DATA cells;
    /// the returned rectangle includes a new header above the shifted data.
    pub fn preview_table_creation(
        &self,
        sheet: SheetId,
        range: TableRange,
        has_headers: bool,
    ) -> Result<(TableRange, Vec<String>), String> {
        if has_headers {
            return self.preview_table_headers(sheet, range).map(|h| (range, h));
        }
        self.validate_table_region(sheet, range, None)?;
        let index = self
            .sheet_index_by_id(sheet)
            .ok_or("Sheet no longer exists.")?;
        let mut result = range;
        result.end_row = result
            .end_row
            .checked_add(1)
            .ok_or("No room for a header row.")?;
        result.validate(self.sheets[index].rows, self.sheets[index].cols)?;
        self.validate_structural_edit(index, Axis::Row, range.start_row, 1, false)?;
        let source = &self.sheets[index];
        let last = source.rows - 1;
        if !source.occupied_cells_in_rows(last, 1).is_empty()
            || source.comments().any(|((row, _), _)| row == last)
            || source.row_formats.contains_key(&last)
        {
            return Err(
                "Inserting a header would push cells, comments or formatting off the sheet.".into(),
            );
        }
        Ok((
            result,
            (1..=range.width()).map(|i| format!("Column{i}")).collect(),
        ))
    }

    pub fn create_table_without_headers(
        &mut self,
        sheet: SheetId,
        range: TableRange,
        name: &str,
    ) -> Result<TableCommit, String> {
        self.validate_table_name_available(name, None)?;
        let (result, _) = self.preview_table_creation(sheet, range, false)?;
        let index = self.sheet_index_by_id(sheet).unwrap();
        let mut staged = self.clone();
        let rows = staged.prepare_table_row_history(index, range.start_row, 1, false)?;
        let print_before = staged.sheets[index].print_setup.clone();
        let rewrites = if let Some(rows) = &rows {
            staged.apply_table_row_history(rows, false)?
        } else {
            staged.structural_edit(index, Axis::Row, range.start_row, 1, false)?
        };
        let formulas = rewrites
            .into_iter()
            .map(|(s, r, c, b, a)| (CellId::new(staged.sheets[s].id, r, c), b, a))
            .collect();
        let mut commit = staged.create_table(sheet, result, name)?;
        // Totals use complete sparse structural history. Capture creation and
        // insertion together so the intermediate header cells are not mistaken
        // for a stale row-history replay when creation is undone.
        let guarded = if self.tables().any(|(_, t)| t.totals.is_some()) {
            Some(Box::new(self.capture_guarded_batch(&staged)?))
        } else { None };
        commit.header_insertion = Some(Box::new(HeaderInsertion {
            at: range.start_row,
            rows: if guarded.is_none() { rows } else { None },
            guarded,
            print_before,
            formulas,
            inserted_cells: staged.sheets[index].occupied_cells_in_rows(range.start_row, 1),
        }));
        staged.set_revision_after_atomic_commit(self.revision().saturating_add(1));
        *self = staged;
        Ok(commit)
    }

    pub(super) fn apply_headerless_table_commit(
        &mut self,
        commit: &TableCommit,
        undo: bool,
    ) -> Result<(), String> {
        let header = commit.header_insertion.as_ref().unwrap();
        if let Some(guarded) = &header.guarded {
            let mut candidate = guarded.candidate(self, undo)?;
            candidate.refresh_table_name_reservations();
            self.restore_snapshot_monotonic(&candidate);
            return Ok(());
        }
        let index = self
            .sheet_index_by_id(commit.sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        if header.rows.is_none()
            && self.sheets[index]
                .tables()
                .iter()
                .any(|t| !undo || t.id != commit.id)
        {
            return Err("Tables changed since header insertion.".into());
        }
        if undo {
            let sheet = &self.sheets[index];
            if sheet.occupied_cells_in_rows(header.at, 1) != header.inserted_cells
                || sheet.comments().any(|((r, _), _)| r == header.at)
                || sheet.row_formats.contains_key(&header.at)
            {
                return Err("The inserted header row changed. Undo those edits first.".into());
            }
        }
        for (cell, before, after) in &header.formulas {
            let row = cell.row
                + usize::from(undo && cell.sheet == commit.sheet_id && cell.row >= header.at);
            let current = self
                .sheet_by_id(cell.sheet)
                .ok_or("Formula sheet no longer exists.")?
                .get_raw(row, cell.col);
            if current != *if undo { after } else { before } {
                return Err(
                    "A formula changed since header insertion. Undo those edits first.".into(),
                );
            }
        }
        let mut staged = self.clone();
        if undo {
            // Removing the new Table first releases its protected header so
            // the inverse structural operation can remove just the new row.
            staged.apply_table_commit_inner(commit, true)?;
            if let Some(rows) = &header.rows {
                staged.apply_table_row_history(rows, true)?;
            } else {
                staged.structural_edit(index, Axis::Row, header.at, 1, true)?;
            }
            staged.sheets[index].print_setup = header.print_before.clone();
            for (cell, before, _) in &header.formulas {
                staged
                    .sheet_by_id_mut(cell.sheet)
                    .unwrap()
                    .set_value(cell.row, cell.col, before);
            }
            staged.rebuild_dep_graph();
            staged.recompute_full_ordered();
        } else {
            let mut data = commit.after.table.as_ref().unwrap().range;
            data.end_row -= 1;
            staged.preview_table_creation(commit.sheet_id, data, false)?;
            if let Some(rows) = &header.rows {
                staged.apply_table_row_history(rows, false)?;
            } else {
                staged.structural_edit(index, Axis::Row, header.at, 1, false)?;
            }
            staged.apply_table_commit_inner(commit, false)?;
        }
        staged.set_revision_after_atomic_commit(self.revision().saturating_add(1));
        *self = staged;
        Ok(())
    }
}
