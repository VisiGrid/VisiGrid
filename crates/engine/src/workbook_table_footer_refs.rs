//! References to moved footer cells follow their content. Larger worksheet
//! ranges retain their authored bounds. Called only on an atomic candidate.
use super::{TableCommit, Workbook};
use crate::{
    cell::{CellValue, ValueRef},
    formula::parser::{self, Expr, ParsedExpr},
    named_range::NamedRangeTarget,
    sheet::UnboundSheetRef,
    table::{TableId, TableRange},
    validation::{ConstraintValue, ListSource, ValidationType},
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
        let mut expr = parser::parse(&input)
            .map_err(|_| "A formula cannot be checked for totals-row references. Resolve it before moving the footer.".to_string())?;
        if crate::formula::analyze::has_dynamic_deps(&expr) {
            return Err("INDIRECT or OFFSET references cannot be checked safely before moving totals. Use explicit or structured references first.".into());
        }
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
    pub(super) fn relocate_footer_references(
        &mut self,
        id: TableId,
        new_end_row: usize,
        owned: Option<&TableCommit>,
    ) -> Result<bool, String> {
        let (owner_id, table) = self.table(id).ok_or("Table no longer exists.")?;
        let Some(row) = table.totals_row() else {
            return Ok(false);
        };
        let destination = new_end_row
            .checked_add(1)
            .ok_or("Totals row exceeds the worksheet boundary.")?;
        if destination == row {
            return Ok(false);
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
        // Update original formulas before the footer moves and new records are
        // authored. Relative/absolute flags and calculated origins are retained.
        for sheet in &mut self.sheets {
            let local = sheet.id == owner_id;
            let mut cells = Vec::new();
            for ((r, c), cell) in sheet.cells_iter() {
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
            let rule_ids: Vec<_> = sheet.cond_formats.iter().map(|r| r.id).collect();
            for id in rule_ids {
                let rule = sheet.cond_formats.get_mut(id).unwrap();
                if movement.rewrite(&mut rule.predicate, local, &mut guarded)? {
                    rule.reparse();
                    metadata_changed = true;
                }
            }
            let validations: Vec<_> = sheet
                .validations
                .iter()
                .map(|(r, rule)| (*r, rule.clone()))
                .collect();
            for (range, mut rule) in validations {
                let mut changed = false;
                match &mut rule.rule_type {
                    ValidationType::Custom(source)
                    | ValidationType::List(ListSource::Range(source)) => {
                        changed |= movement.rewrite(source, local, &mut guarded)?
                    }
                    ValidationType::List(_) => {}
                    ValidationType::WholeNumber(c)
                    | ValidationType::Decimal(c)
                    | ValidationType::Date(c)
                    | ValidationType::Time(c)
                    | ValidationType::TextLength(c) => {
                        for value in std::iter::once(&mut c.value1).chain(c.value2.iter_mut()) {
                            if let ConstraintValue::CellRef(source)
                            | ConstraintValue::Formula(source) = value
                            {
                                changed |= movement.rewrite(source, local, &mut guarded)?;
                            }
                        }
                    }
                }
                if changed {
                    sheet.validations.set(range, rule);
                    metadata_changed = true;
                }
            }
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
        Ok(guarded)
    }
}
