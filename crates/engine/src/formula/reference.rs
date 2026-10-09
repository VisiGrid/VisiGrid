//! Reference-producing expressions shared by formula evaluation and dropdowns.
//! Keep the rectangle until a caller needs values; sparse dropdowns must not
//! materialize a whole column just to read its occupied cells.
use super::{
    eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution},
    parser::{BoundExpr, Expr, RangeAxis},
};
use crate::sheet::{SheetRef, NUM_COLS, NUM_ROWS};

pub(crate) struct ReferenceRange {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}
impl ReferenceRange {
    fn new(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> Self {
        Self {
            start_row,
            start_col,
            end_row,
            end_col,
        }
    }
}

pub(crate) fn resolve<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> Result<Source, String> {
    resolve_expr(expr, lookup, 0)
}

pub(crate) fn evaluate_reference<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> EvalResult {
    match resolve(expr, lookup) {
        Ok(Source::Reference(sheet, range)) => {
            lookup.record_dynamic_reference(
                &sheet,
                range.start_row,
                range.start_col,
                range.end_row,
                range.end_col,
            );
            super::eval_lookup::range_to_result(
                lookup,
                &sheet,
                range.start_row,
                range.start_col,
                range.end_row - range.start_row + 1,
                range.end_col - range.start_col + 1,
            )
        }
        Ok(Source::Value(value)) => value,
        Err(error) => EvalResult::Error(error),
    }
}

pub(crate) enum Source {
    Reference(SheetRef, ReferenceRange),
    Value(EvalResult),
}

fn reference(
    sheet: SheetRef,
    r0: usize,
    c0: usize,
    r1: usize,
    c1: usize,
) -> Result<Source, String> {
    if matches!(sheet, SheetRef::RefError { .. }) {
        return Err("#REF!".into());
    }
    if r0 >= NUM_ROWS
        || r1 >= NUM_ROWS
        || c0 >= NUM_COLS
        || c1 >= NUM_COLS
    {
        return Err("#REF! Invalid reference".into());
    }
    Ok(Source::Reference(
        sheet,
        ReferenceRange::new(r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1)),
    ))
}

fn scalar_number<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> Result<f64, String> {
    let value = evaluate(expr, lookup).to_number()?;
    if !value.is_finite() {
        return Err("#NUM! Non-finite reference argument".into());
    }
    Ok(value.trunc())
}

fn resolve_expr<L: CellLookup>(
    expr: &BoundExpr,
    lookup: &L,
    depth: usize,
) -> Result<Source, String> {
    if depth >= 64 {
        return Err("Reference nesting limit exceeded".into());
    }
    let resolve = |e: &BoundExpr| resolve_expr(e, lookup, depth + 1);
    match expr {
        Expr::CellRef {
            sheet, row, col, ..
        } => reference(sheet.clone(), *row, *col, *row, *col),
        Expr::Range {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => reference(sheet.clone(), *start_row, *start_col, *end_row, *end_col),
        Expr::WholeRange {
            sheet,
            axis,
            start,
            end,
            ..
        } => match axis {
            RangeAxis::Row => reference(sheet.clone(), *start, 0, *end, NUM_COLS - 1),
            RangeAxis::Column => reference(sheet.clone(), 0, *start, NUM_ROWS - 1, *end),
        },
        Expr::NamedRange(name) if !lookup.is_table_name(name) => {
            if let Some(target) = lookup.resolve_named_reference(name) {
                return resolve(&target);
            }
            // Single-sheet adapters may expose only the legacy name API.
            match lookup.resolve_named_range(name) {
                Some(NamedRangeResolution::Cell { row, col }) => {
                    reference(SheetRef::Current, row, col, row, col)
                }
                Some(NamedRangeResolution::Range {
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                }) => reference(SheetRef::Current, start_row, start_col, end_row, end_col),
                None => Err(format!("#NAME? Unknown reference name: {name}")),
            }
        }
        Expr::NamedRange(_) | Expr::StructuredRef(_) => {
            resolve(&crate::formula::structured::resolve(expr, lookup))
        }
        Expr::EmptyRange { .. } => Ok(Source::Value(EvalResult::Empty)),
        Expr::Function { name, args } if name.eq_ignore_ascii_case("INDIRECT") => {
            if args.is_empty() || args.len() > 2 {
                return Err("INDIRECT requires one or two arguments".into());
            }
            if let Some(style) = args.get(1) {
                if !evaluate(style, lookup).to_bool()? {
                    return Err("#REF! R1C1 references are not supported".into());
                }
            }
            let result = evaluate(&args[0], lookup);
            if let EvalResult::Error(error) = result {
                return Err(error);
            }
            if result.dimensions() != (1, 1) {
                return Err("INDIRECT requires one reference string".into());
            }
            let indirect = lookup.bind_reference_text(&result.to_text())?;
            // INDIRECT accepts a reference or name, never evaluates formula text.
            if !matches!(
                indirect,
                Expr::CellRef { .. }
                    | Expr::Range { .. }
                    | Expr::WholeRange { .. }
                    | Expr::NamedRange(_)
                    | Expr::StructuredRef(_)
            ) {
                return Err("#REF! INDIRECT did not produce a reference".into());
            }
            // Text that is neither an address nor a defined name is #REF! in
            // Excel. Binding reads it as a name, whose own error is #NAME?.
            resolve(&indirect).map_err(|e| if e.starts_with("#NAME?") { "#REF!".into() } else { e })
        }
        Expr::Function { name, args } if name.eq_ignore_ascii_case("OFFSET") => {
            if !(3..=5).contains(&args.len()) {
                return Err("OFFSET requires three to five arguments".into());
            }
            let Source::Reference(sheet, base) = resolve(&args[0])? else {
                return Err("OFFSET requires a reference".into());
            };
            let dr = scalar_number(&args[1], lookup)?;
            let dc = scalar_number(&args[2], lookup)?;
            let h = match args.get(3) {
                Some(e) => scalar_number(e, lookup)?,
                None => (base.end_row - base.start_row + 1) as f64,
            };
            let w = match args.get(4) {
                Some(e) => scalar_number(e, lookup)?,
                None => (base.end_col - base.start_col + 1) as f64,
            };
            let r = base.start_row as f64 + dr;
            let c = base.start_col as f64 + dc;
            if r < 0.0
                || c < 0.0
                || h < 1.0
                || w < 1.0
                || r + h > NUM_ROWS as f64
                || c + w > NUM_COLS as f64
            {
                return Err("#REF! OFFSET is outside the worksheet".into());
            }
            reference(
                sheet,
                r as usize,
                c as usize,
                (r + h - 1.0) as usize,
                (c + w - 1.0) as usize,
            )
        }
        Expr::Function { name, args } if name.eq_ignore_ascii_case("IF") => {
            if !(2..=3).contains(&args.len()) {
                return Err("IF requires two or three arguments".into());
            }
            let condition = evaluate(&args[0], lookup).to_bool()?;
            match args.get(if condition { 1 } else { 2 }) {
                Some(branch) => resolve(branch),
                None => Ok(Source::Value(EvalResult::Boolean(false))),
            }
        }
        Expr::Function { name, args } if name.eq_ignore_ascii_case("CHOOSE") => {
            let index = args.first().ok_or("CHOOSE requires an index")?;
            let index = scalar_number(index, lookup)?;
            if index < 1.0 || index >= args.len() as f64 {
                return Err("#VALUE! CHOOSE index is outside the choices".into());
            }
            resolve(&args[index as usize])
        }
        Expr::Function { name, args }
            if name.eq_ignore_ascii_case("INDEX") && (2..=3).contains(&args.len()) =>
        {
            let base = resolve(&args[0])?;
            let Source::Reference(sheet, range) = base else {
                return Ok(Source::Value(evaluate(expr, lookup)));
            };
            let h = range.end_row - range.start_row + 1;
            let w = range.end_col - range.start_col + 1;
            let row = scalar_number(&args[1], lookup)?;
            let col = match args.get(2) {
                Some(e) => scalar_number(e, lookup)?,
                None => {
                    if h == 1 {
                        row
                    } else {
                        1.0
                    }
                }
            };
            let row = if args.len() == 2 && h == 1 { 1.0 } else { row };
            if row < 0.0 || col < 0.0 || row > h as f64 || col > w as f64 {
                return Err("#REF! INDEX is outside the source".into());
            }
            let (r0, r1) = if row == 0.0 {
                (range.start_row, range.end_row)
            } else {
                let r = range.start_row + row as usize - 1;
                (r, r)
            };
            let (c0, c1) = if col == 0.0 {
                (range.start_col, range.end_col)
            } else {
                let c = range.start_col + col as usize - 1;
                (c, c)
            };
            reference(sheet, r0, c0, r1, c1)
        }
        _ => Ok(Source::Value(evaluate(expr, lookup))),
    }
}
