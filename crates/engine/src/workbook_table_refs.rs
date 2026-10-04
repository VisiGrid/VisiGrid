//! Schema edits rewrite only reference-token spans. Names bind against the old
//! schema, then follow stable column IDs into the new schema. Formula literals,
//! whitespace, parentheses and unrelated references are left byte-for-byte intact.
use super::Workbook;
use crate::cell::ValueRef;
use crate::cell_id::CellId;
use crate::formula::parser::{format_expr, BoundExpr, Expr};
use crate::formula::structured::{self, StructuredReference};
use crate::sheet::SheetId;
use crate::table::{DataTable, TableId, TableTotals};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TableFormulaChange {
    pub cell: CellId,
    pub before: String,
    pub after: String,
}

/// Compare the input schemas before derived formula rewrites are applied.
pub(crate) fn names_only(before: &DataTable, after: &DataTable) -> bool {
    if before.columns.len() != after.columns.len() {
        return false;
    }
    let mut normalized = after.clone();
    normalized.name = before.name.clone();
    for (old, new) in before.columns.iter().zip(&mut normalized.columns) {
        new.name = old.name.clone();
    }
    normalized == *before
}

#[derive(Debug, Clone)]
pub(crate) struct TotalsReferenceChange {
    pub sheet: SheetId,
    pub table: TableId,
    pub before: TableTotals,
    pub after: TableTotals,
}

impl Workbook {
    fn reference_targets_table(
        &self,
        reference: &StructuredReference,
        sheet: SheetId,
        row: usize,
        col: usize,
        owner_sheet: SheetId,
        table: &DataTable,
    ) -> bool {
        match &reference.table {
            Some(name) => name.eq_ignore_ascii_case(&table.name),
            None => sheet == owner_sheet && table.full_range().contains(row, col),
        }
    }

    pub(crate) fn table_reference_sources(
        &self,
        owner_sheet: SheetId,
        table: &DataTable,
    ) -> Vec<(CellId, String)> {
        let mut sources = Vec::new();
        for sheet in self.sheets() {
            for ((row, col), cell) in sheet.cells_iter() {
                if sheet.id == owner_sheet
                    && row == table.range.start_row
                    && table.full_range().contains(row, col)
                {
                    continue;
                }
                let ValueRef::Formula {
                    source,
                    ast: Some(_),
                    ..
                } = cell.value()
                else {
                    continue;
                };
                if structured::source_references(source)
                    .iter()
                    .any(|(_, _, r)| {
                        self.reference_targets_table(r, sheet.id, row, col, owner_sheet, table)
                    })
                {
                    sources.push((CellId::new(sheet.id, row, col), source.to_string()));
                }
            }
        }
        sources.sort_by_key(|(cell, _)| (cell.sheet.0, cell.row, cell.col));
        sources
    }

    pub(crate) fn table_formula_changes(
        &self,
        owner_sheet: SheetId,
        before: Option<&DataTable>,
        after: Option<&DataTable>,
        allow_totals: bool,
    ) -> Result<Vec<TableFormulaChange>, String> {
        let Some(before) = before else {
            return Ok(Vec::new());
        };
        let mut changes = Vec::new();
        for sheet in self.sheets() {
            for table in sheet.tables().iter().filter(|t| t.id != before.id) {
                if let Some(totals) = table.totals.as_ref().filter(|_| !allow_totals) {
                    for (offset, total) in totals.columns.iter().enumerate() {
                        if let Some(source) = &total.formula {
                            if self.rewrite_table_formula_source(owner_sheet, before, after, sheet.id,
                                table.range.end_row + 1, table.range.start_col + offset, source)? != *source {
                                return Err("This schema change would rewrite totals metadata. Totals reference rewriting is not supported yet.".into());
                            }
                        }
                    }
                }
            }
            for ((row, col), cell) in sheet.cells_iter() {
                let ValueRef::Formula {
                    source,
                    ast: Some(_),
                    ..
                } = cell.value()
                else {
                    continue;
                };
                let rewritten = self.rewrite_table_formula_source(
                    owner_sheet,
                    before,
                    after,
                    sheet.id,
                    row,
                    col,
                    source,
                )?;
                if rewritten != source {
                    if !allow_totals && sheet.table_at(row, col).is_some_and(|t| t.totals_row() == Some(row) && t.id != before.id) {
                        return Err("This schema change would rewrite a totals formula. Totals reference rewriting is not supported yet.".into());
                    }
                    changes.push(TableFormulaChange {
                        cell: CellId::new(sheet.id, row, col),
                        before: source.to_string(),
                        after: rewritten,
                    });
                }
            }
        }
        changes.sort_by_key(|c| (c.cell.sheet.0, c.cell.row, c.cell.col));
        Ok(changes)
    }

    pub(crate) fn totals_reference_changes(
        &self,
        owner: SheetId,
        before: &DataTable,
        after: &DataTable,
    ) -> Result<Vec<TotalsReferenceChange>, String> {
        let mut changes = Vec::new();
        for (sheet, table) in self.tables() {
            let Some(totals) = &table.totals else { continue; };
            let mut rewritten = if table.id == before.id {
                after.totals.clone().unwrap()
            } else { totals.clone() };
            // Dormant footers still bind local references to their owning Table.
            let mut context = before.clone();
            let mut target = after.clone();
            if table.id == before.id {
                context.totals.as_mut().unwrap().visible = true;
                target.totals.as_mut().unwrap().visible = true;
            }
            for (offset, total) in rewritten.columns.iter_mut().enumerate() {
                if let Some(source) = &mut total.formula {
                    let old_offset = if table.id == before.id {
                        table.columns.iter().position(|c| c.id == after.columns[offset].id)
                            .ok_or("A new totals column cannot already have a custom formula.")?
                    } else { offset };
                    *source = self.rewrite_table_formula_source(
                        owner, &context, Some(&target), sheet,
                        table.range.end_row + 1, table.range.start_col + old_offset, source,
                    )?;
                }
            }
            if rewritten != *totals {
                changes.push(TotalsReferenceChange {
                    sheet, table: table.id, before: totals.clone(), after: rewritten,
                });
            }
        }
        Ok(changes)
    }

    pub(crate) fn rewrite_table_formula_source(
        &self,
        owner_sheet: SheetId,
        before: &DataTable,
        after: Option<&DataTable>,
        sheet: SheetId,
        row: usize,
        col: usize,
        source: &str,
    ) -> Result<String, String> {
        let mut replacements = Vec::new();
        for (start, end, reference) in structured::source_references(source) {
            if !self.reference_targets_table(&reference, sheet, row, col, owner_sheet, before) {
                continue;
            }
            let replacement = if let Some(after) = after {
                let mut rewritten = reference.clone();
                if let Some((a, b)) = &reference.columns {
                    let mut renamed = Vec::new();
                    let mut removed = false;
                    for name in [a, b] {
                        if let Some(old) = before.column_by_name(name) {
                            if let Some(new) = after.columns.iter().find(|c| c.id == old.id) {
                                renamed.push(if old.name == new.name {
                                    name.clone()
                                } else {
                                    new.name.clone()
                                });
                            } else {
                                removed = true;
                                break;
                            }
                        } else {
                            renamed.push(name.clone());
                        }
                    }
                    if removed {
                        replacements.push((start, end, "#REF!".into()));
                        continue;
                    }
                    rewritten.columns = Some((renamed[0].clone(), renamed[1].clone()));
                }
                if reference.table.is_some() && before.name != after.name {
                    rewritten.table = Some(after.name.clone());
                }
                // A released row must not silently adopt another table's
                // local context when that rectangle is reused later.
                if reference.table.is_none() && !after.full_range().contains(row, col) {
                    rewritten.table = Some(after.name.clone());
                }
                if rewritten == reference {
                    continue;
                }
                if !source[start..end].contains('[') {
                    after.name.clone()
                } else {
                    rewritten.format()
                }
            } else {
                // Convert to Range preserves formulas as absolute A1 refs.
                // An empty body has no lossless A1 representation.
                let resolved = structured::resolve_region(
                    before,
                    owner_sheet,
                    sheet,
                    Some((row, col)),
                    &reference,
                );
                match resolved {
                            Expr::Range { .. } | Expr::CellRef { .. } => format_expr(&resolved, |id| self.sheet_by_id(id).map(|s| s.name.clone()))[1..].to_string(),
                            Expr::EmptyRange { .. } => return Err("Cannot convert a referenced empty table to a range; it has no A1 representation.".into()),
                            _ => return Err("Cannot convert this table while it has unresolved structured references.".into()),
                        }
            };
            replacements.push((start, end, replacement));
        }
        let mut rewritten = source.to_string();
        for (start, end, replacement) in replacements.into_iter().rev() {
            rewritten.replace_range(start..end, &replacement);
        }
        Ok(rewritten)
    }

    /// Sheet history does not yet capture cross-sheet Table-reference rewrites.
    /// Refuse removal of a referenced Table sheet instead of allowing a same-
    /// named future table to adopt those references.
    pub(crate) fn has_external_table_references(&self, sheet_id: SheetId) -> bool {
        let Some(owner) = self.sheet_by_id(sheet_id) else {
            return false;
        };
        if owner.tables().is_empty() {
            return false;
        }
        self.sheets()
            .iter()
            .filter(|s| s.id != sheet_id)
            .any(|sheet| {
                sheet.tables().iter().any(|t| {
                    let rules = t.columns.iter().filter_map(|c| c.formula.as_deref());
                    let totals = t.totals.iter().flat_map(|totals|
                        totals.columns.iter().filter_map(|c| c.formula.as_deref()));
                    rules.chain(totals).any(|source| {
                        structured::source_references(source).iter().any(|(_, _, r)| {
                            r.table.as_ref().is_some_and(|name| {
                                owner.tables().iter().any(|owned| owned.name.eq_ignore_ascii_case(name))
                            })
                        })
                    })
                }) || sheet.cells_iter().any(|((row, col), cell)| {
                    let ValueRef::Formula {
                        source,
                        ast: Some(_),
                        ..
                    } = cell.value()
                    else {
                        return false;
                    };
                    structured::source_references(source)
                        .iter()
                        .any(|(_, _, r)| {
                            owner.tables().iter().any(|t| {
                                self.reference_targets_table(r, sheet.id, row, col, sheet_id, t)
                            })
                        })
                })
            })
    }

    /// Current range descriptor for inspection and dependency tracking.
    pub fn resolve_table_reference_at(
        &self,
        reference: &StructuredReference,
        sheet: SheetId,
        row: usize,
        col: usize,
    ) -> BoundExpr {
        use crate::formula::eval::CellLookup;
        super::WorkbookLookup::with_cell_context(self, sheet, row, col)
            .resolve_table_reference(reference, Some((row, col)))
    }
}
