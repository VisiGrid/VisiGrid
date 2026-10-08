//! References to moved footer cells follow their content. Larger worksheet
//! ranges retain their authored bounds. Called only on an atomic candidate.
#[path = "workbook_table_footer_metadata.rs"]
mod metadata;

use super::{TableCommit, Workbook};
use crate::{
    cell::{CellValue, ValueRef},
    formula::parser::{self, Expr, ParsedExpr},
    named_range::NamedRangeTarget,
    sheet::UnboundSheetRef,
    table::{TableId, TableRange},
};

struct Movement {
    owner: String,
    footer: TableRange,
    destination: usize,
}

impl Movement {
    fn contains(&self, start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> bool {
        self.footer.contains(start_row, start_col) && self.footer.contains(end_row, end_col)
    }

    fn adjust(&self, expr: &mut ParsedExpr, local: bool) -> bool {
        let same_sheet = |sheet: &UnboundSheetRef| match sheet {
            UnboundSheetRef::Current => local,
            UnboundSheetRef::Named(name) => name.eq_ignore_ascii_case(&self.owner),
        };
        match expr {
            Expr::CellRef {
                sheet, row, col, ..
            } if same_sheet(sheet) && self.footer.contains(*row, *col) => {
                *row = self.destination;
                true
            }
            Expr::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } if same_sheet(sheet) && self.contains(*start_row, *start_col, *end_row, *end_col) => {
                *start_row = self.destination;
                *end_row = self.destination;
                true
            }
            Expr::Function { args, .. } => {
                let mut changed = false;
                for arg in args {
                    changed |= self.adjust(arg, local);
                }
                changed
            }
            Expr::BinaryOp { left, right, .. } => {
                self.adjust(left, local) | self.adjust(right, local)
            }
            // Whole rows/columns and ranges extending beyond the moved cells
            // keep worksheet-coordinate semantics. Strings are never rewritten.
            _ => false,
        }
    }

    fn rewrite(
        &self,
        source: &mut String,
        local: bool,
        guarded: &mut bool,
    ) -> Result<bool, String> {
        let equals = source.starts_with('=');
        let input = if equals {
            source.clone()
        } else {
            format!("={source}")
        };
        let mut expr = match parser::parse(&input) {
            Ok(expr) => expr,
            Err(_) => {
                let result = crate::formula::source_refs::rewrite(source, |expr| self.adjust(expr, local))?;
                let changed = result != *source;
                *guarded |= changed;
                *source = result;
                return Ok(changed);
            }
        };
        // Explicit OFFSET bases follow the moved cells. Constructed addresses
        // and offsets retain their authored meaning; runtime dependencies are
        // rebuilt/settled on the final candidate, with guarded history even if
        // no source token changes (for example INDIRECT("B5")).
        *guarded |= crate::formula::analyze::has_dynamic_deps(&expr);
        let future = Movement {
            owner: self.owner.clone(),
            footer: TableRange {
                start_row: self.destination,
                end_row: self.destination,
                ..self.footer
            },
            destination: self.footer.start_row,
        };
        *guarded |= future.adjust(&mut expr.clone(), local);
        if !self.adjust(&mut expr, local) {
            return Ok(false);
        }
        *guarded = true;
        let result = parser::format_parsed_expr(&expr);
        *source = if equals {
            result
        } else {
            result.trim_start_matches('=').to_string()
        };
        Ok(true)
    }
}

impl Workbook {
    pub(super) fn validate_footer_relocation(&mut self, before: &Workbook, id: TableId) -> Result<(), String> {
        let new_cycles = {
            let mut allowed = before.dep_graph.find_cycle_members();
            if let (Some((sheet, old)), Some((_, new))) = (before.table(id), self.table(id)) {
                if let (Some(from), Some(to)) = (old.totals_row(), new.totals_row()) {
                    allowed = allowed.into_iter().map(|mut cell| {
                        if cell.sheet == sheet && cell.row == from
                            && (old.range.start_col..=old.range.end_col).contains(&cell.col)
                            && (new.range.start_col..=new.range.end_col).contains(&cell.col) {
                            cell.row = to;
                        }
                        cell
                    }).collect();
                }
            }
            !self.dep_graph.find_cycle_members().is_subset(&allowed)
        };
        if new_cycles || self.incremental_errors.iter().any(|e| e.error.contains("not settled")) {
            return Err("Moving totals would create a cycle or an unsettled calculation. Nothing was changed.".into());
        }
        for sheet in &self.sheets {
            if before.sheet_by_id(sheet.id).is_some_and(|b| b.edit_generation() == sheet.edit_generation()) { continue; }
            sheet.build_saved_table_view(sheet.rows)?;
        }
        Ok(())
    }

    pub(super) fn relocate_footer_references(
        &mut self,
        id: TableId,
        new_end_row: usize,
        owned: Option<&TableCommit>,
    ) -> Result<(bool, Vec<crate::cell_id::CellId>), String> {
        let (owner_id, table) = self.table(id).ok_or("Table no longer exists.")?;
        let Some(row) = table.totals_row() else {
            return Ok((false, Vec::new()));
        };
        let destination = new_end_row
            .checked_add(1)
            .ok_or("Totals row exceeds the worksheet boundary.")?;
        if destination == row {
            return Ok((false, Vec::new()));
        }
        let owner_index = self.sheet_index_by_id(owner_id).unwrap();
        let owner = self.sheet_by_id(owner_id).unwrap();
        if destination >= owner.rows.min(crate::sheet::NUM_ROWS) {
            return Err("Totals row exceeds the worksheet boundary.".into());
        }
        let movement = Movement {
            owner: owner.name.clone(),
            footer: TableRange {
                start_row: row,
                end_row: row,
                ..table.range
            },
            destination,
        };
        let mut guarded = false;
        let owned_cells: std::collections::HashSet<_> = owned
            .into_iter()
            .flat_map(|c| {
                c.cells
                    .iter()
                    .map(|(cell, _)| (c.sheet_id, cell.row, cell.col))
            })
            .collect();
        let mut sources: rustc_hash::FxHashSet<_> = self.volatile_cells.iter().copied().chain(self.table_readers.get(&id).into_iter().flat_map(|r| r.iter().copied())).collect();
        for col in movement.footer.start_col..=movement.footer.end_col {
            for row in [row, destination] {
                sources.extend(self.dep_graph.dependents(crate::cell_id::CellId::new(owner_id, row, col)));
            }
        }
        for sheet in &self.sheets {
            sources.extend(sheet.exceptional_reference_sources().into_iter().map(|(r,c)| crate::cell_id::CellId::new(sheet.id,r,c)));
        }
        // Update original formulas before the footer moves and new records are
        // authored. Relative/absolute flags and calculated origins are retained.
        for sheet in &mut self.sheets {
            let local = sheet.id == owner_id;
            let mut cells = Vec::new();
            for source in sources.iter().filter(|c| c.sheet == sheet.id) {
                let (r, c) = (source.row, source.col);
                let Some(cell) = sheet.get_cell_opt(r, c) else { continue; };
                if owned_cells.contains(&(sheet.id, r, c)) {
                    continue;
                }
                let mut formula = match cell.value() {
                    ValueRef::Formula { source, .. } => Some(source.to_string()),
                    _ => None,
                };
                let mut frozen = cell.frozen_formula().map(str::to_string);
                let mut changed = false;
                if let Some(source) = &mut formula {
                    changed |= movement.rewrite(source, local, &mut guarded)?;
                }
                if let Some(source) = &mut frozen {
                    changed |= movement.rewrite(source, local, &mut guarded)?;
                }
                if changed {
                    let mut image = cell.to_cell();
                    if let Some(source) = formula {
                        image.value = CellValue::from_input(&source);
                    }
                    image.set_frozen_formula(frozen);
                    cells.push((r, c, image));
                }
            }
            for (r, c, image) in cells {
                sheet.restore_history_cell(r, c, Some(image));
            }
            let mut metadata_changed = false;
            for table in &mut sheet.data_tables {
                for column in &mut table.columns {
                    if let Some(source) = &mut column.formula {
                        metadata_changed |= movement.rewrite(source, local, &mut guarded)?;
                    }
                }
                if let Some(totals) = &mut table.totals {
                    for column in &mut totals.columns {
                        if let Some(source) = &mut column.formula {
                            metadata_changed |= movement.rewrite(source, local, &mut guarded)?;
                        }
                    }
                }
            }
            metadata_changed |= movement.rewrite_rule_stores(sheet, local, &mut guarded)?;
            for pivot in &mut sheet.pivots {
                let source = &mut pivot.source;
                if source.table_id.is_none()
                    && source.sheet_id == owner_id
                    && source.start_row as usize == destination
                    && source.end_row as usize == destination
                    && source.start_col as usize >= movement.footer.start_col
                    && source.end_col as usize <= movement.footer.end_col
                {
                    guarded = true;
                }
                if source.table_id.is_none()
                    && source.sheet_id == owner_id
                    && movement.contains(
                        source.start_row as usize,
                        source.start_col as usize,
                        source.end_row as usize,
                        source.end_col as usize,
                    )
                {
                    guarded = true;
                    source.start_row = destination as u32;
                    source.end_row = destination as u32;
                    pivot.stale = true;
                    pivot.source_generation = None;
                    metadata_changed = true;
                }
            }
            if metadata_changed {
                sheet.mark_table_changed();
            }
        }
        let names: Vec<_> = self.named_ranges.list().into_iter().cloned().collect();
        for mut name in names {
            let future = TableRange {
                start_row: destination,
                end_row: destination,
                ..movement.footer
            };
            guarded |= match name.target {
                NamedRangeTarget::RefError => false,
                NamedRangeTarget::Cell { sheet, row, col } => {
                    sheet == owner_index && future.contains(row, col)
                }
                NamedRangeTarget::Range {
                    sheet,
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                } => {
                    sheet == owner_index
                        && future.contains(start_row, start_col)
                        && future.contains(end_row, end_col)
                }
            };
            let changed = match &mut name.target {
                NamedRangeTarget::Cell { sheet, row, col }
                    if *sheet == owner_index && movement.footer.contains(*row, *col) =>
                {
                    *row = destination;
                    true
                }
                NamedRangeTarget::Range {
                    sheet,
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                } if *sheet == owner_index
                    && movement.contains(*start_row, *start_col, *end_row, *end_col) =>
                {
                    *start_row = destination;
                    *end_row = destination;
                    true
                }
                _ => false,
            };
            if changed {
                guarded = true;
                self.named_ranges.set(name)?;
            }
        }
        Ok((guarded, sources.into_iter().collect()))
    }
}
