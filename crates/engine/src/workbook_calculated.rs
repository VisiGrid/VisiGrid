//! Materialized calculated columns. Rules live in schema; divergent cell
//! contents are exceptions, including explicitly empty cells.
use super::{HeaderCell, TableCommit};
use crate::formula::parser::{bind_expr, parse};
use crate::{
    cell::CellValue,
    sheet::SheetId,
    table::{TableColumnId, TableId},
    workbook::Workbook,
};

#[derive(Debug, Clone, serde::Serialize)]
pub(crate) struct RuleChange {
    pub sheet: SheetId,
    pub table: TableId,
    pub column: TableColumnId,
    pub before: String,
    pub before_origin: usize,
    pub after_origin: usize,
    pub after: String,
}

impl Workbook {
    /// Only direct single-cell authoring establishes a rule implicitly. File
    /// loading, paste, and normal setters never infer calculated columns.
    pub fn try_calculated_column(
        &mut self,
        sheet: SheetId,
        row: usize,
        col: usize,
        source: &str,
    ) -> Result<Option<TableCommit>, String> {
        if !source.starts_with('=') {
            return Ok(None);
        }
        let sheet = self.sheet_by_id(sheet).ok_or("Sheet no longer exists.")?;
        let Some(table) = sheet.table_at(row, col) else {
            return Ok(None);
        };
        if row <= table.range.start_row || row > table.range.end_row
            || table.columns[col - table.range.start_col].formula.is_some()
        {
            return Ok(None);
        }
        if (table.range.start_row + 1..=table.range.end_row)
            .any(|r| r != row && !sheet.get_raw(r, col).is_empty())
        {
            return Ok(None);
        }
        self.set_calculated_column(table.id, col, row, source, true)
            .map(Some)
    }

    /// Explicit template edit updates followers. Replacing exceptions requires
    /// replace_all and a desktop preview of the affected count.
    pub fn set_calculated_column(
        &mut self,
        id: TableId,
        col: usize,
        source_row: usize,
        source: &str,
        replace_all: bool,
    ) -> Result<TableCommit, String> {
        if !source.starts_with('=') {
            return Err("Enter a formula beginning with =.".into());
        }
        parse(source).map_err(|e| format!("Invalid column formula: {e}"))?;
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        if col < old.range.start_col || col > old.range.end_col {
            return Err("Column is outside the Table.".into());
        }
        if source_row >= self.sheet_by_id(sheet_id).unwrap().rows {
            return Err("The formula origin is outside the sheet.".into());
        }
        self.validate_calculated_formula(sheet_id, source_row, col, source)?;
        if !replace_all && old.columns[col - old.range.start_col].formula.is_none() {
            return Err("Use formula for entire column to establish a new rule.".into());
        }
        let mut new = old.clone();
        new.columns[col - old.range.start_col].formula = Some(source.into());
        new.columns[col - old.range.start_col].formula_origin = source_row
            .checked_sub(old.range.start_row)
            .ok_or("The formula origin must be at or below the header.")?;
        let sheet = self.sheet_by_id(sheet_id).unwrap();
        let mut cells = Vec::new();
        for row in old.range.start_row + 1..=old.range.end_row {
            if !replace_all && sheet.is_calculated_exception(row, col) {
                continue;
            }
            let formula = new.formula_at(row, col).unwrap();
            self.validate_calculated_formula(sheet_id, row, col, &formula)?;
            cells.push((
                HeaderCell {
                    row,
                    col,
                    value: sheet.get_cell(row, col).value,
                },
                HeaderCell {
                    row,
                    col,
                    value: CellValue::from_input(&formula),
                },
            ));
        }
        let mut commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        commit.calculated_edit = true;
        commit.cells = cells;
        self.capture_table_cell_absence(&mut commit);
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    pub fn restore_calculated_cell(
        &mut self,
        id: TableId,
        row: usize,
        col: usize,
    ) -> Result<TableCommit, String> {
        let (sheet_id, table) = self.table(id).ok_or("Table no longer exists.")?;
        if row <= table.range.start_row || !table.range.contains(row, col) {
            return Err("Select a Table body cell.".into());
        }
        let formula = table
            .formula_at(row, col)
            .ok_or("This column has no formula rule.")?;
        self.validate_calculated_formula(sheet_id, row, col, &formula)?;
        let mut commit =
            self.table_commit(sheet_id, id, Some(table.clone()), Some(table.clone()))?;
        commit.calculated_edit = true;
        commit.cells.push((
            HeaderCell {
                row,
                col,
                value: self.sheet_by_id(sheet_id).unwrap().get_cell(row, col).value,
            },
            HeaderCell {
                row,
                col,
                value: CellValue::from_input(&formula),
            },
        ));
        self.capture_table_cell_absence(&mut commit);
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    pub(super) fn validate_calculated_formula(
        &self,
        sheet: SheetId,
        row: usize,
        col: usize,
        formula: &str,
    ) -> Result<(), String> {
        use crate::formula::eval::{evaluate, EvalResult};
        let ast = parse(formula).map_err(|e| format!("Invalid column formula: {e}"))?;
        let bound = bind_expr(&ast, |name| self.sheet_id_by_name(name));
        let lookup = crate::workbook::WorkbookLookup::with_cell_context(self, sheet, row, col);
        if matches!(evaluate(&bound, &lookup), EvalResult::Array(_)) {
            return Err("A calculated-column formula must return one value per row; arrays cannot spill inside a Table.".into());
        }
        Ok(())
    }
}

impl Workbook {
    pub(super) fn schema_rule_changes(
        &self,
        owner: SheetId,
        before: Option<&crate::table::DataTable>,
        after: Option<&crate::table::DataTable>,
    ) -> Result<Vec<RuleChange>, String> {
        let mut changes = Vec::new();
        let Some(before) = before else {
            return Ok(changes);
        };
        for (sheet, table) in self.tables() {
            // Removing a rule's own Table materializes its cells and drops it.
            if table.id == before.id && after.is_none() {
                continue;
            }
            for (offset, column) in table.columns.iter().enumerate() {
                let Some(source) = &column.formula else {
                    continue;
                };
                let row = table.range.start_row + column.formula_origin;
                let col = table.range.start_col + offset;
                // Local references must also bind in header-only Tables.
                let mut context = before.clone();
                if context.id == table.id {
                    context.range.end_row = context.range.end_row.max(row);
                }
                let mut target = after.cloned();
                if let Some(target) = &mut target {
                    if target.id == table.id {
                        target.range.end_row = target.range.end_row.max(row);
                    }
                }
                let formula = self.rewrite_table_formula_source(
                    owner,
                    &context,
                    target.as_ref(),
                    sheet,
                    row,
                    col,
                    source,
                )?;
                if formula != *source {
                    changes.push(RuleChange {
                        sheet,
                        table: table.id,
                        column: column.id,
                        before: source.clone(),
                        before_origin: column.formula_origin,
                        after_origin: column.formula_origin,
                        after: formula,
                    });
                }
            }
        }
        Ok(changes)
    }

    pub(super) fn validate_rule_changes(
        &self,
        changes: &[RuleChange],
        undo: bool,
    ) -> Result<(), String> {
        for change in changes {
            let current = self
                .table(change.table)
                .and_then(|(_, t)| t.columns.iter().find(|c| c.id == change.column))
                .and_then(|c| c.formula.as_ref().map(|f| (f, c.formula_origin)));
            let expected = if undo {
                (&change.after, change.after_origin)
            } else {
                (&change.before, change.before_origin)
            };
            if current != Some(expected) {
                return Err(
                    "A calculated-column rule changed since this operation was prepared.".into(),
                );
            }
        }
        Ok(())
    }

    pub(crate) fn apply_rule_changes(&mut self, changes: &[RuleChange], undo: bool) {
        for change in changes {
            if let Some(column) = self
                .sheet_by_id_mut(change.sheet)
                .and_then(|s| s.data_tables.iter_mut().find(|t| t.id == change.table))
                .and_then(|t| t.columns.iter_mut().find(|c| c.id == change.column))
            {
                column.formula_origin = if undo {
                    change.before_origin
                } else {
                    change.after_origin
                };
                column.formula = Some(if undo {
                    change.before.clone()
                } else {
                    change.after.clone()
                });
            }
        }
    }
}

impl Workbook {
    pub(crate) fn structural_rule_changes(
        &self,
        index: usize,
        axis: crate::structural::Axis,
        at: usize,
        count: usize,
        delete: bool,
    ) -> Vec<RuleChange> {
        use crate::structural::{adjust_formula_text, shift_span, Axis, StructuralEdit};
        let edit = StructuralEdit {
            sheet_name: self.sheets[index].name.clone(),
            axis,
            at,
            count,
            delete,
        };
        let edited = self.sheets[index].id;
        let mut changes = Vec::new();
        for (sheet, table) in self.tables() {
            let mut new_range = table.range;
            if sheet == edited {
                let (a, b) = if axis == Axis::Row {
                    (new_range.start_row, new_range.end_row)
                } else {
                    (new_range.start_col, new_range.end_col)
                };
                if let Some((a, b)) = shift_span(a, b, at, count, delete) {
                    if axis == Axis::Row {
                        new_range.start_row = a;
                        new_range.end_row = b;
                    } else {
                        new_range.start_col = a;
                        new_range.end_col = b;
                    }
                }
            }
            for (offset, column) in table.columns.iter().enumerate() {
                let Some(source) = &column.formula else {
                    continue;
                };
                // Keep the authored origin when it survives. When it is
                // deleted, project to the header (a surviving virtual record)
                // so relative same-row references survive even at the grid edge.
                let origin_row = table.range.start_row + column.formula_origin;
                let removed = sheet == edited
                    && axis == Axis::Row
                    && delete
                    && origin_row >= at
                    && origin_row < at + count;
                let mut virtual_row = if removed {
                    table.range.start_row
                } else {
                    origin_row
                };
                let mut projected = table
                    .formula_at(virtual_row, table.range.start_col + offset)
                    .unwrap();
                if removed && projected.contains("#REF!") && !source.contains("#REF!") {
                    virtual_row = table.range.end_row + 1;
                    projected = table
                        .formula_at(virtual_row, table.range.start_col + offset)
                        .unwrap();
                }
                let rewritten =
                    adjust_formula_text(&projected, &edit, &self.sheet_by_id(sheet).unwrap().name)
                        .unwrap_or(projected);
                let mut moved_row = virtual_row;
                if sheet == edited && axis == Axis::Row && at <= virtual_row {
                    moved_row = if delete {
                        virtual_row.saturating_sub(count.min(virtual_row - at))
                    } else {
                        virtual_row + count
                    };
                }
                let formula = rewritten;
                let origin = moved_row.saturating_sub(new_range.start_row);
                if formula != *source || origin != column.formula_origin {
                    changes.push(RuleChange {
                        sheet,
                        table: table.id,
                        column: column.id,
                        before: source.clone(),
                        before_origin: column.formula_origin,
                        after_origin: origin,
                        after: formula,
                    });
                }
            }
        }
        changes
    }

    pub(crate) fn fill_inserted_calculated_rows(&mut self, index: usize, at: usize, count: usize) {
        let mut writes = Vec::new();
        for table in self.sheets[index].tables() {
            for row in at.max(table.range.start_row + 1)..(at + count).min(table.range.end_row + 1)
            {
                for col in table.range.start_col..=table.range.end_col {
                    if let Some(formula) = table.formula_at(row, col) {
                        writes.push((row, col, formula));
                    }
                }
            }
        }
        for (row, col, formula) in writes {
            self.sheets[index].write_table_header(row, col, CellValue::from_input(&formula));
        }
    }
}
