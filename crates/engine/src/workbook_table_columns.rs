//! Worksheet column schema and sparse history. Ordinary edits retain schema
//! and rules; totals delegate to guarded history to restore protected footers.
use super::{calculated::RuleChange, same_schema, TotalsReferenceChange};
use crate::{
    cell::CellValue,
    sheet::{Sheet, SheetId},
    structural::Axis,
    table::{DataTable, TableColumn, TableColumnId},
    workbook::Workbook,
};

#[derive(Debug, Clone)]
pub struct TableColumnHistory {
    sheet: SheetId,
    at: usize,
    count: usize,
    delete: bool,
    before: Vec<DataTable>,
    after: Vec<DataTable>,
    rules: Vec<RuleChange>,
    guarded: Option<Box<crate::workbook::GuardedStructureCommit>>,
}

impl Sheet {
    pub(crate) fn tables_after_column_edit(
        &self,
        at: usize,
        count: usize,
        delete: bool,
    ) -> Result<Vec<DataTable>, String> {
        let mut tables = self.tables().to_vec();
        for table in &mut tables {
            let start = table.range.start_col;
            let end = table.range.end_col;
            if delete {
                if let Some(totals) = &mut table.totals {
                    totals.columns = totals.columns.iter().enumerate()
                        .filter(|(i, _)| start + i < at || start + i >= at + count)
                        .map(|(_, total)| total.clone()).collect();
                }
                table.columns = table
                    .columns
                    .iter()
                    .enumerate()
                    .filter(|(i, _)| start + i < at || start + i >= at + count)
                    .map(|(_, c)| c.clone())
                    .collect();
                if table.columns.is_empty() {
                    return Err(format!(
                        "Cannot delete every column of {}. Convert it to a range first.",
                        table.name
                    ));
                }
            } else if at > start && at <= end {
                let mut used: std::collections::HashSet<_> = table
                    .columns
                    .iter()
                    .map(|c| c.name.to_lowercase())
                    .collect();
                let mut columns = Vec::new();
                let mut number = 1;
                for _ in 0..count {
                    while used.contains(&format!("column{number}")) {
                        number += 1;
                    }
                    let name = format!("Column{number}");
                    used.insert(name.to_lowercase());
                    let id = TableColumnId(table.next_column_id);
                    table.next_column_id = table
                        .next_column_id
                        .checked_add(1)
                        .ok_or("Column IDs exhausted.")?;
                    columns.push(TableColumn {
                        id,
                        name,
                        formula: None,
                        formula_origin: 1,
                    });
                }
                table.columns.splice(at - start..at - start, columns);
                if let Some(totals) = &mut table.totals {
                    totals.columns.splice(at - start..at - start,
                        std::iter::repeat_with(Default::default).take(count));
                }
            }
            let (start, end) = crate::structural::shift_span(start, end, at, count, delete)
                .ok_or("Cannot remove a Table's last column.")?;
            table.range.start_col = start;
            table.range.end_col = end;
            table.validate(self.rows, self.cols)?;
        }
        self.validate_table_view_schema(&tables)?;
        Ok(tables)
    }

    pub(crate) fn install_column_tables(&mut self, mut tables: Vec<DataTable>) {
        for table in &mut tables {
            let next = self.table_column_allocators.entry(table.id.0).or_insert(1);
            *next = (*next).max(table.next_column_id);
            table.next_column_id = *next;
        }
        self.data_tables = tables;
        self.mark_table_changed();
    }

    pub(crate) fn sync_table_headers(&mut self) {
        let headers: Vec<_> =
            self.tables()
                .iter()
                .flat_map(|t| {
                    t.columns.iter().enumerate().map(move |(i, c)| {
                        (t.range.start_row, t.range.start_col + i, c.name.clone())
                    })
                })
                .collect();
        for (r, c, name) in headers {
            self.write_table_header(r, c, CellValue::Text(name));
        }
    }
}

impl Workbook {
    pub(crate) fn structural_totals_changes(
        &self, index: usize, axis: Axis, at: usize, count: usize, delete: bool,
        before: &[DataTable], after: &[DataTable],
    ) -> Result<Vec<TotalsReferenceChange>, String> {
        let owner = &self.sheets[index];
        let edit = crate::structural::StructuralEdit {
            sheet_name: owner.name.clone(), axis, at, count, delete,
        };
        // Dormant totals need the same local structured-reference context as a
        // visible footer, including when a deleted field becomes #REF!.
        let visible = |tables: &[DataTable]| tables.iter().cloned().map(|mut t| {
            if let Some(totals) = &mut t.totals { totals.visible = true; }
            t
        }).collect::<Vec<_>>();
        let contexts = visible(before);
        let targets = visible(after);
        let mut changes = Vec::new();
        for (sheet, table) in self.tables() {
            let Some(old) = &table.totals else { continue; };
            let target = if sheet == owner.id {
                after.iter().find(|t| t.id == table.id).unwrap()
            } else { table };
            let mut new = target.totals.clone().unwrap();
            for (i, column) in target.columns.iter().enumerate() {
                let Some(source) = &mut new.columns[i].formula else { continue; };
                let old_col = table.columns.iter().position(|c| c.id == column.id).unwrap();
                let rewritten = if axis == Axis::Col {
                    self.rewrite_column_schema_source(owner.id, &contexts, &targets,
                        sheet, table.range.end_row + 1, table.range.start_col + old_col, source)?
                } else { source.clone() };
                *source = crate::structural::adjust_formula_text(&rewritten, &edit,
                    &self.sheet_by_id(sheet).unwrap().name).unwrap_or(rewritten);
            }
            if new != *old {
                changes.push(TotalsReferenceChange { sheet, table: table.id, before: old.clone(), after: new });
            }
        }
        Ok(changes)
    }

    pub(crate) fn rewrite_column_schema_source(
        &self,
        owner: SheetId,
        before: &[DataTable],
        after: &[DataTable],
        sheet: SheetId,
        row: usize,
        col: usize,
        source: &str,
    ) -> Result<String, String> {
        let mut source = source.to_string();
        for old in before {
            let new = after.iter().find(|t| t.id == old.id).unwrap();
            if new.columns.len() >= old.columns.len() {
                continue;
            }
            let mut new = new.clone();
            // Bind local selectors at their pre-move coordinates. A structural
            // move keeps their Table context even when the old cell is outside
            // the new rectangle.
            new.range = old.range;
            source = self.rewrite_table_formula_source(
                owner,
                old,
                Some(&new),
                sheet,
                row,
                col,
                &source,
            )?;
        }
        Ok(source)
    }

    pub(crate) fn column_rule_changes(
        &self,
        index: usize,
        at: usize,
        count: usize,
        delete: bool,
        before: &[DataTable],
        after: &[DataTable],
    ) -> Result<Vec<RuleChange>, String> {
        let structural = self.structural_rule_changes(index, Axis::Col, at, count, delete);
        let owner = self.sheets[index].id;
        let mut changes = Vec::new();
        for (sheet, table) in self.tables() {
            for (i, column) in table.columns.iter().enumerate() {
                let Some(source) = &column.formula else {
                    continue;
                };
                if sheet == owner
                    && delete
                    && (at..at + count).contains(&(table.range.start_col + i))
                {
                    continue;
                }
                let moved = structural
                    .iter()
                    .find(|r| r.table == table.id && r.column == column.id);
                let adjusted = moved.map_or(source.as_str(), |r| r.after.as_str());
                // Header-only Tables still provide a virtual body-row context.
                let formula = self.rewrite_column_schema_source(
                    owner,
                    before,
                    after,
                    sheet,
                    table.range.start_row,
                    table.range.start_col + i,
                    adjusted,
                )?;
                if formula != *source {
                    changes.push(RuleChange {
                        sheet,
                        table: table.id,
                        column: column.id,
                        before: source.clone(),
                        after: formula,
                        before_origin: column.formula_origin,
                        after_origin: column.formula_origin,
                    });
                }
            }
        }
        Ok(changes)
    }

    pub fn prepare_table_column_history(
        &self,
        index: usize,
        at: usize,
        count: usize,
        delete: bool,
    ) -> Result<Option<TableColumnHistory>, String> {
        self.validate_structural_edit(index, Axis::Col, at, count, delete)?;
        let sheet = &self.sheets[index];
        if self.tables().any(|(_, t)| t.totals.is_some()) {
            let (_, guarded) = self.prepare_guarded_structure(index, vec![crate::workbook::StructureStep {
                axis: Axis::Col, at, count, delete,
            }])?;
            return Ok(Some(TableColumnHistory {
                sheet: sheet.id, at, count, delete, before: Vec::new(), after: Vec::new(),
                rules: Vec::new(), guarded: Some(Box::new(guarded)),
            }));
        }
        let before = sheet.tables().to_vec();
        let mut after = sheet.tables_after_column_edit(at, count, delete)?;
        let rules = self.column_rule_changes(index, at, count, delete, &before, &after)?;
        if before.is_empty() && rules.is_empty() {
            return Ok(None);
        }
        for table in &mut after {
            for change in rules.iter().filter(|r| r.table == table.id) {
                table
                    .columns
                    .iter_mut()
                    .find(|c| c.id == change.column)
                    .unwrap()
                    .formula = Some(change.after.clone());
            }
        }
        Ok(Some(TableColumnHistory {
            sheet: sheet.id,
            at,
            count,
            delete,
            before,
            after,
            rules,
            guarded: None,
        }))
    }

    pub fn validate_table_column_history(
        &self,
        history: &TableColumnHistory,
        undo: bool,
    ) -> Result<(), String> {
        if let Some(guarded) = &history.guarded {
            return guarded.candidate(self, undo).map(|_| ());
        }
        let index = self
            .sheet_index_by_id(history.sheet)
            .ok_or("Table sheet no longer exists.")?;
        let expected = if undo {
            &history.after
        } else {
            &history.before
        };
        let current = self.sheets[index].tables();
        let target = if undo { &history.before } else { &history.after };
        self.sheets[index].validate_table_view_schema(target)?;
        if current.len() != expected.len()
            || expected
                .iter()
                .any(|t| !current.iter().any(|c| same_schema(t, c)))
        {
            return Err("Tables changed since the column operation was prepared.".into());
        }
        self.validate_rule_changes(&history.rules, undo)?;
        self.validate_structural_edit(
            index,
            Axis::Col,
            history.at,
            history.count,
            history.delete != undo,
        )
    }

    pub fn apply_table_column_history(
        &mut self,
        history: &TableColumnHistory,
        undo: bool,
    ) -> Result<Vec<(usize, usize, usize, String, String)>, String> {
        if let Some(guarded) = &history.guarded {
            let candidate = guarded.candidate(self, undo)?;
            self.restore_snapshot_monotonic(&candidate);
            return Ok(Vec::new());
        }
        self.validate_table_column_history(history, undo)?;
        let index = self.sheet_index_by_id(history.sheet).unwrap();
        let rewrites = self.structural_edit_with_rules(
            index,
            Axis::Col,
            history.at,
            history.count,
            history.delete != undo,
            false,
        )?;
        let target = if undo {
            &history.before
        } else {
            &history.after
        };
        self.sheets[index].install_column_tables(target.clone());
        self.sheets[index].sync_table_headers();
        self.apply_rule_changes(&history.rules, undo);
        self.rebuild_dep_graph();
        self.recompute_full_ordered();
        Ok(rewrites)
    }
}
