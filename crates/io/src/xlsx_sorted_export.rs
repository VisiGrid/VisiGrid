//! Materialize saved Table sorts in an isolated export copy. Every accepted
//! reference follows its original records; unsupported coordinate semantics
//! refuse before any destination is opened.
use crate::xlsx::ExportLayout;
use std::{
    borrow::Cow,
    collections::{BTreeMap, HashMap},
};
use visigrid_engine::{
    cell::ValueRef,
    formula::{
        parser::{self, Expr, ParsedExpr, RangeAxis},
        structured::TableSection,
    },
    sheet::UnboundSheetRef,
    named_range::NamedRangeTarget,
    table::TableRange,
    table_view::TableView,
    workbook::Workbook,
};

struct Permutation {
    range: TableRange,
    /// Canonical body-row offset -> exported worksheet row, including hidden records.
    destinations: Vec<usize>,
}
impl Permutation {
    fn row(&self, row: usize, col: usize) -> usize {
        if row > self.range.start_row
            && row <= self.range.end_row
            && col >= self.range.start_col
            && col <= self.range.end_col
        {
            self.destinations[row - self.range.start_row - 1]
        } else {
            row
        }
    }
}
struct Plan<'a> {
    wb: &'a Workbook,
    sheets: HashMap<usize, Permutation>,
}
impl Plan<'_> {
    fn row(&self, sheet: usize, row: usize, col: usize) -> usize {
        self.sheets.get(&sheet).map_or(row, |p| p.row(row, col))
    }
    fn sheet(&self, current: usize, reference: &UnboundSheetRef) -> Result<usize, String> {
        match reference {
            UnboundSheetRef::Current => Ok(current),
            UnboundSheetRef::Named(name) => self
                .wb
                .sheets()
                .iter()
                .position(|s| Some(s.id) == self.wb.sheet_id_by_name(name))
                .ok_or_else(|| format!("Unresolved sheet reference {name}")),
        }
    }
    // Ranges must keep their ordered shape. Direct aggregate arguments may
    // instead keep the same rectangular membership with a different row order.
    fn range(&self, sheet: usize, r: TableRange, aggregate: bool) -> Result<TableRange, String> {
        if r.start_row > r.end_row || r.start_col > r.end_col {
            return Err("Reversed ranges are not supported for materialized sorting".into());
        }
        let Some(p) = self.sheets.get(&sheet) else {
            return Ok(r);
        };
        if r.end_row <= p.range.start_row
            || r.start_row > p.range.end_row
            || r.end_col < p.range.start_col
            || r.start_col > p.range.end_col
        {
            return Ok(r);
        }
        let start = r.start_row.max(p.range.start_row + 1);
        let end = r.end_row.min(p.range.end_row);
        if aggregate
            && (start..=end).all(|row| {
                let dst = p.row(row, r.start_col.max(p.range.start_col));
                dst >= r.start_row && dst <= r.end_row
            })
        {
            return Ok(r);
        }
        if aggregate
            && r.start_row > p.range.start_row
            && r.end_row <= p.range.end_row
            && r.start_col >= p.range.start_col
            && r.end_col <= p.range.end_col
        {
            let mapped =
                &p.destinations[r.start_row - p.range.start_row - 1..r.end_row - p.range.start_row];
            let min = *mapped.iter().min().unwrap();
            let max = *mapped.iter().max().unwrap();
            if max - min == r.end_row - r.start_row {
                return Ok(TableRange {
                    start_row: min,
                    end_row: max,
                    ..r
                });
            }
        }
        let delta = p.row(r.start_row, r.start_col) as isize - r.start_row as isize;
        if (r.start_col < p.range.start_col
            || r.end_col > p.range.end_col
            || r.start_row <= p.range.start_row
            || r.end_row > p.range.end_row)
            && delta != 0
        {
            return Err("A range would split across moved and stationary cells".into());
        }
        if !(start..=end).all(|row| {
            p.row(row, r.start_col.max(p.range.start_col)) as isize - row as isize == delta
        }) {
            return Err("A range would become discontiguous or change record order".into());
        }
        Ok(TableRange {
            start_row: (r.start_row as isize + delta) as usize,
            end_row: (r.end_row as isize + delta) as usize,
            ..r
        })
    }
    fn expr(
        &self,
        expr: &mut ParsedExpr,
        sheet: usize,
        row: usize,
        col: usize,
        aggregate: bool,
    ) -> Result<(), String> {
        match expr {
            Expr::CellRef {
                sheet: target,
                row: r,
                col: c,
                ..
            } => {
                *r = self.row(self.sheet(sheet, target)?, *r, *c);
            }
            Expr::Range {
                sheet: target,
                start_row,
                end_row,
                start_col,
                end_col,
                start_row_abs,
                end_row_abs,
                ..
            } => {
                let r = self.range(
                    self.sheet(sheet, target)?,
                    TableRange {
                        start_row: *start_row,
                        end_row: *end_row,
                        start_col: *start_col,
                        end_col: *end_col,
                    },
                    aggregate && start_row_abs == end_row_abs,
                )?;
                *start_row = r.start_row;
                *end_row = r.end_row;
            }
            Expr::WholeRange {
                sheet: target,
                axis,
                start,
                end,
                start_abs,
                end_abs,
                ..
            } => {
                let sid = self.sheet(sheet, target)?;
                let s = self.wb.sheet(sid).unwrap();
                let r = match axis {
                    RangeAxis::Column => TableRange {
                        start_row: 0,
                        end_row: s.rows - 1,
                        start_col: *start,
                        end_col: *end,
                    },
                    RangeAxis::Row => TableRange {
                        start_row: *start,
                        end_row: *end,
                        start_col: 0,
                        end_col: s.cols - 1,
                    },
                };
                let mapped = self.range(
                    sid,
                    r,
                    aggregate && (*axis == RangeAxis::Column || start_abs == end_abs),
                )?;
                if mapped != r {
                    return Err("An entire-row/column reference would move".into());
                }
            }
            Expr::StructuredRef(reference) => {
                let (sid, table) = if let Some(name) = &reference.table {
                    let (id, t) = self
                        .wb
                        .table_by_name(name)
                        .ok_or("Unresolved Table reference")?;
                    (self.wb.sheets().iter().position(|s| s.id == id).unwrap(), t)
                } else {
                    (
                        sheet,
                        self.wb
                            .sheet(sheet)
                            .unwrap()
                            .table_at(row, col)
                            .ok_or("Unresolved local Table reference")?,
                    )
                };
                if reference.section == TableSection::ThisRow {
                    // A named this-row reference to a different Table follows a
                    // worksheet row, not the record moved by this Table's sort.
                    if sid != sheet
                        || self.row(sheet, row, col) != self.row(sid, row, table.range.start_col)
                    {
                        return Err(
                            "A this-row reference would resolve to a different record".into()
                        );
                    }
                } else if self
                    .sheets
                    .get(&sid)
                    .is_some_and(|p| p.range == table.range)
                    && reference.section != TableSection::Headers
                    && !aggregate
                {
                    return Err(
                        "An order-sensitive Table reference requires stored-order export".into(),
                    );
                }
            }
            Expr::Function { name, args } => {
                let name = name.to_ascii_uppercase();
                let reducer = matches!(
                    name.as_str(),
                    "SUM"
                        | "AVERAGE"
                        | "COUNT"
                        | "COUNTA"
                        | "COUNTBLANK"
                        | "MIN"
                        | "MAX"
                        | "PRODUCT"
                        | "AND"
                        | "OR"
                );
                // Deliberate allowlist: INDIRECT/OFFSET, coordinate functions,
                // volatile functions, lookups and custom functions cannot be
                // proven safe merely by rewriting their visible references.
                let scalar = matches!(
                    name.as_str(),
                    "IF" | "IFERROR"
                        | "IFNA"
                        | "NOT"
                        | "ABS"
                        | "ROUND"
                        | "ROUNDUP"
                        | "ROUNDDOWN"
                        | "INT"
                        | "TRUNC"
                        | "MOD"
                        | "POWER"
                        | "SQRT"
                        | "SIGN"
                        | "EXP"
                        | "LN"
                        | "LOG"
                        | "LOG10"
                        | "LEN"
                        | "LOWER"
                        | "UPPER"
                        | "TRIM"
                        | "LEFT"
                        | "RIGHT"
                        | "MID"
                        | "SUBSTITUTE"
                        | "REPLACE"
                        | "CONCATENATE"
                        | "TEXT"
                        | "VALUE"
                        | "ISNUMBER"
                        | "ISTEXT"
                        | "ISBLANK"
                        | "ISERROR"
                        | "ISNA"
                        | "DATE"
                        | "YEAR"
                        | "MONTH"
                        | "DAY"
                        | "TRUE"
                        | "FALSE"
                );
                if !reducer && !scalar {
                    return Err(format!(
                        "Function {name} is not supported for materialized sorting"
                    ));
                }
                for arg in args {
                    self.expr(arg, sheet, row, col, reducer)?;
                }
            }
            Expr::BinaryOp { left, right, .. } => {
                self.expr(left, sheet, row, col, false)?;
                self.expr(right, sheet, row, col, false)?;
            }
            Expr::NamedRange(name) => {
                if self.wb.named_ranges().get(name).is_none() {
                    return Err(format!("Unresolved named reference {name}"));
                }
                // Targets are mapped once below. Keep ordered shape even for
                // aggregate uses: another consumer may use INDEX on this name.
            }
            Expr::ReferenceError(_) | Expr::RefError | Expr::EmptyRange { .. } => {
                return Err("A formula has an unresolved reference".into())
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Boolean(_) | Expr::Empty => {}
        }
        Ok(())
    }
    fn formula(
        &self,
        source: &str,
        sheet: usize,
        row: usize,
        col: usize,
    ) -> Result<String, String> {
        let mut expr = parser::parse(source)?;
        let original = expr.clone();
        self.expr(&mut expr, sheet, row, col, false)?;
        Ok(if expr == original {
            source.into()
        } else {
            parser::format_parsed_expr(&expr)
        })
    }
}

pub(crate) fn prepare<'a>(
    wb: &'a Workbook,
    layouts: Option<&[ExportLayout]>,
) -> Result<Cow<'a, Workbook>, String> {
    prepare_inner(wb, layouts).map_err(|reason| format!("Cannot export Tables in sorted order: {reason}. Choose stored-order export or save a native .sheet file; no file was written."))
}
pub(super) fn prepare_inner<'a>(
    wb: &'a Workbook,
    layouts: Option<&[ExportLayout]>,
) -> Result<Cow<'a, Workbook>, String> {
    let mut plan = Plan {
        wb,
        sheets: HashMap::new(),
    };
    for (sid, sheet) in wb.sheets().iter().enumerate() {
        let Some(spec) = sheet.table_view_spec().filter(|v| v.sort.is_some()) else {
            continue;
        };
        let table = sheet
            .tables()
            .iter()
            .find(|t| t.id == spec.table)
            .ok_or("Missing sorted Table")?;
        let mut sort_only = spec.clone();
        sort_only.filters.clear(); // All records participate, including filtered-out records.
        let view = TableView::build(sheet, sort_only, table.range.end_row + 1, None)?;
        let destinations: Vec<_> = (table.range.start_row + 1..=table.range.end_row)
            .map(|r| view.rows().data_to_view_unchecked(r))
            .collect();
        if destinations
            .iter()
            .enumerate()
            .all(|(i, r)| *r == table.range.start_row + 1 + i)
        {
            continue;
        }
        if wb.tables().any(|(_, t)| t.totals.is_some()) {
            return Err("Tables with totals metadata currently require stored-order export".into());
        }
        if sheet.manual_hidden_rows().iter().any(|row| *row > table.range.start_row && *row <= table.range.end_row) {
            return Err(format!("Table {} has manual row visibility that requires stored-order export", table.name));
        }
        if let Some(layout) = layouts.and_then(|l| l.get(sid)) {
            if layout
                .row_heights
                .keys()
                .chain(layout.hidden_rows.iter())
                .any(|r| *r > table.range.start_row && *r <= table.range.end_row)

                || layout.autofilter_range.is_some()
            {
                return Err(format!(
                    "Table {} has row layout or worksheet filters that cannot follow the sort",
                    table.name
                ));
            }
        }
        if sheet
            .row_formats
            .keys()
            .any(|r| *r > table.range.start_row && *r <= table.range.end_row)
        {
            return Err(format!("Table {} has inherited row formatting", table.name));
        }
        plan.sheets.insert(
            sid,
            Permutation {
                range: table.range,
                destinations,
            },
        );
    }
    if plan.sheets.is_empty() {
        return Ok(Cow::Borrowed(wb));
    }
    let mut names = Vec::new();
    for original in wb.named_ranges().list() {
        let mut name = original.clone();
        match &mut name.target {
            NamedRangeTarget::Cell { sheet, row, col } => *row = plan.row(*sheet, *row, *col),
            NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                let mapped = plan.range(*sheet, TableRange {
                    start_row: *start_row, start_col: *start_col,
                    end_row: *end_row, end_col: *end_col,
                }, false).map_err(|e| format!("Defined name '{}': {e}", name.name))?;
                *start_row = mapped.start_row;
                *end_row = mapped.end_row;
            }
        }
        names.push(name);
    }
    let mut changes = BTreeMap::new();
    let mut images = Vec::new();
    for (sid, sheet) in wb.sheets().iter().enumerate() {
        if !sheet.validations.is_empty() || !sheet.cond_formats.is_empty() {
            return Err(
                "Validation or conditional-format rules require stored-order export".into(),
            );
        }
        for ((r, c), cell) in sheet.cells_iter() {
            if cell.is_spill_receiver() || sheet.is_spill_parent(r, c) {
                return Err("Spilled formulas require stored-order export".into());
            }
            let dst = plan.row(sid, r, c);
            let formula = if let ValueRef::Formula { source, .. } = cell.value() {
                Some(plan.formula(source, sid, r, c).map_err(|e| {
                    format!(
                        "{}!{}{}: {e}",
                        sheet.name,
                        crate::xlsx::col_to_letter(c),
                        r + 1
                    )
                })?)
            } else {
                None
            };
            let rewritten = formula.as_ref().is_some_and(|s| *s != sheet.get_raw(r, c));
            if dst != r || rewritten {
                let mut image = cell.to_cell();
                if let Some(source) = formula {
                    image.set(&source);
                }
                changes.insert((sid, r, c), None);
                images.push(((sid, dst, c), image));
            }
        }
    }
    for (coord, image) in images {
        changes.insert(coord, Some(image));
    }
    let mut catalog = wb.saved_tables();
    for entry in &mut catalog.sheets {
        for table in &mut entry.tables {
            for (offset, column) in table.columns.iter_mut().enumerate() {
                let Some(source) = &column.formula else {
                    continue;
                };
                let col = table.range.start_col + offset;
                let origin = table.range.start_row + column.formula_origin;
                let mapped = plan.formula(source, entry.sheet, origin, col)?;
                let first = table.range.start_row + 1;
                let base = parser::adjust_formula_refs(
                    &mapped,
                    first as i32 - plan.row(entry.sheet, origin, col) as i32,
                    0,
                );
                for row in first..=table.range.end_row {
                    let old = parser::adjust_formula_refs(source, row as i32 - origin as i32, 0);
                    let mapped = plan.formula(&old, entry.sheet, row, col)?;
                    let expected = parser::adjust_formula_refs(
                        &base,
                        plan.row(entry.sheet, row, col) as i32 - first as i32,
                        0,
                    );
                    if parser::parse(&mapped)? != parser::parse(&expected)? {
                        return Err(format!(
                            "Calculated column {}[{}] would need different fill rules per record",
                            table.name, column.name
                        ));
                    }
                }
                column.formula = Some(base);
                column.formula_origin = 1;
            }
        }
    }
    let mut out = wb.clone();
    out.set_auto_recalc(false);
    for name in names { out.named_ranges_mut().set(name)?; }
    for ((sid, row, col), image) in changes {
        out.restore_cell_tracked(sid, row, col, image)?;
    }
    out.restore_tables(catalog)?;
    out.rebuild_dep_graph();
    let report = out.recompute_full_ordered();
    if report.had_cycles || report.cycle_cells > 0 {
        return Err("Circular formulas require stored-order export".into());
    }
    // Catch any supported formula whose dependency behavior was not preserved,
    // including stale source caches. Never silently publish changed results.
    for (sid, sheet) in wb.sheets().iter().enumerate() {
        for ((r, c), cell) in sheet.cells_iter() {
            if matches!(cell.value(), ValueRef::Formula { .. })
                && sheet.get_computed_value(r, c)
                    != out
                        .sheet(sid)
                        .unwrap()
                        .get_computed_value(plan.row(sid, r, c), c)
            {
                return Err(format!(
                    "{}!{}{} changes result after sorting",
                    sheet.name,
                    crate::xlsx::col_to_letter(c),
                    r + 1
                ));
            }
        }
    }
    out.validate_table_view_specs()?;
    Ok(Cow::Owned(out))
}
