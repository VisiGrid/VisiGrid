//! Move a footer within its columns, without inserting worksheet rows.
use super::{HeaderCell, TableCommit, Workbook};
use crate::{
    cell::{Cell, ValueRef},
    formula::parser::{Expr, ParsedExpr, RangeAxis},
    sheet::{SheetId, UnboundSheetRef},
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

fn points_at(expr: &ParsedExpr, local: bool, owner: &str, footer: TableRange) -> bool {
    let same_sheet = |sheet: &UnboundSheetRef| match sheet {
        UnboundSheetRef::Current => local,
        UnboundSheetRef::Named(name) => name.eq_ignore_ascii_case(owner),
    };
    match expr {
        Expr::CellRef {
            sheet, row, col, ..
        } => same_sheet(sheet) && footer.contains(*row, *col),
        Expr::Range {
            sheet,
            start_row,
            end_row,
            start_col,
            end_col,
            ..
        } => {
            same_sheet(sheet)
                && footer.intersects(TableRange {
                    start_row: *start_row.min(end_row),
                    end_row: *start_row.max(end_row),
                    start_col: *start_col.min(end_col),
                    end_col: *start_col.max(end_col),
                })
        }
        // Whole-column references continue to include the footer at its new row.
        Expr::WholeRange {
            sheet,
            axis: RangeAxis::Row,
            start,
            end,
            ..
        } => same_sheet(sheet) && *start.min(end) <= footer.start_row && *start.max(end) >= footer.end_row,
        Expr::Function { args, .. } => args.iter().any(|e| points_at(e, local, owner, footer)),
        Expr::BinaryOp { left, right, .. } => {
            points_at(left, local, owner, footer) || points_at(right, local, owner, footer)
        }
        _ => false,
    }
}

impl Workbook {
    fn validate_footer_links(
        &self,
        sheet_id: SheetId,
        footer: TableRange,
        owned: Option<&TableCommit>,
    ) -> Result<(), String> {
        let owner = self
            .sheet_by_id(sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        let index = self.sheet_index_by_id(sheet_id).unwrap();
        for (_, pivot) in self.pivots() {
            let source = &pivot.source;
            if source.table_id.is_none()
                && source.sheet_id == sheet_id
                && footer.intersects(TableRange {
                    start_row: source.start_row as usize,
                    end_row: source.end_row as usize,
                    start_col: source.start_col as usize,
                    end_col: source.end_col as usize,
                })
            {
                return Err(format!("{} uses a fixed pivot source that includes a footer position. Use a Table-backed source or change its range before moving totals.", pivot.name));
            }
        }
        for name in self.named_ranges.list() {
            use crate::named_range::NamedRangeTarget;
            let (sheet, range) = match name.target {
                NamedRangeTarget::Cell { sheet, row, col } => (
                    sheet,
                    TableRange {
                        start_row: row,
                        end_row: row,
                        start_col: col,
                        end_col: col,
                    },
                ),
                NamedRangeTarget::Range {
                    sheet,
                    start_row,
                    end_row,
                    start_col,
                    end_col,
                } => (
                    sheet,
                    TableRange {
                        start_row,
                        end_row,
                        start_col,
                        end_col,
                    },
                ),
            };
            if sheet == index && range.intersects(footer) {
                return Err(format!("Named range '{}' includes the totals row. Change its reference before moving the footer.", name.name));
            }
        }
        let owned_cells: std::collections::HashSet<_> = owned
            .into_iter()
            .flat_map(|c| {
                c.cells
                    .iter()
                    .map(|(cell, _)| (c.sheet_id, cell.row, cell.col))
            })
            .collect();
        for sheet in self.sheets() {
            let check = |source: &str| -> Result<(), String> {
                let source = if source.starts_with('=') {
                    source.to_string()
                } else {
                    format!("={source}")
                };
                let expr = crate::formula::parser::parse(&source)
                    .map_err(|_| "A formula cannot be checked for totals-row references. Resolve it before moving the footer.".to_string())?;
                if crate::formula::analyze::has_dynamic_deps(&expr) {
                    return Err("INDIRECT or OFFSET references cannot be checked safely before moving totals. Use explicit or structured references first.".into());
                }
                if points_at(&expr, sheet.id == sheet_id, &owner.name, footer) {
                    return Err("A fixed cell/range reference includes the totals row. Use a structured #Totals reference before moving the footer.".into());
                }
                Ok(())
            };
            for ((row, col), cell) in sheet.cells_iter() {
                if owned_cells.contains(&(sheet.id, row, col)) {
                    continue;
                }
                if let ValueRef::Formula { source, .. } = cell.value() {
                    check(source)?;
                }
                if let Some(source) = cell.frozen_formula() {
                    check(source)?;
                }
            }
            for table in sheet.tables() {
                for column in &table.columns {
                    if let Some(source) = &column.formula {
                        check(source)?;
                    }
                }
                if let Some(totals) = &table.totals {
                    for total in &totals.columns {
                        if let Some(source) = &total.formula {
                            check(source)?;
                        }
                    }
                }
            }
            for rule in sheet.cond_formats.iter() {
                check(&rule.predicate)?;
            }
            use crate::validation::{ConstraintValue, ListSource, ValidationType};
            for (_, rule) in sheet.validations.iter() {
                match &rule.rule_type {
                    ValidationType::Custom(source)
                    | ValidationType::List(ListSource::Range(source)) => check(source)?,
                    ValidationType::List(_) => {}
                    ValidationType::WholeNumber(c)
                    | ValidationType::Decimal(c)
                    | ValidationType::Date(c)
                    | ValidationType::Time(c)
                    | ValidationType::TextLength(c) => {
                        for value in std::iter::once(&c.value1).chain(c.value2.iter()) {
                            if let ConstraintValue::CellRef(source)
                            | ConstraintValue::Formula(source) = value
                            {
                                check(source)?;
                            }
                        }
                    }
                }
            }
        }
        Ok(())
    }

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
        self.validate_footer_links(
            sheet_id,
            TableRange {
                start_row: before_row,
                end_row: before_row,
                ..old.range
            },
            None,
        )?;
        self.validate_footer_links(
            sheet_id,
            TableRange {
                start_row: after_row,
                end_row: after_row,
                ..new.range
            },
            None,
        )?;
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
        Ok(Some(FooterMove {
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
        for row in [movement.before_row, movement.after_row] {
            self.validate_footer_links(
                commit.sheet_id,
                TableRange {
                    start_row: row,
                    end_row: row,
                    ..table.range
                },
                Some(commit),
            )?;
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
