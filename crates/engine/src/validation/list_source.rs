//! Resolve dropdown expressions without materializing reference ranges. This
//! retains unformatted value labels and sparse whole-column behavior for dynamic refs.
use super::{CellRange, ResolvedList, MAX_LIST_ITEMS};
use crate::formula::{
    eval::{evaluate, CellLookup, EvalResult, Value},
    parser::{BoundExpr, Expr, RangeAxis},
};
use crate::sheet::{SheetRef, NUM_COLS, NUM_ROWS};

pub(crate) fn resolve<L: CellLookup>(
    source: &str,
    lookup: &L,
    bind: impl Fn(&str) -> Result<BoundExpr, String>,
    read: impl Fn(&SheetRef, &CellRange) -> ResolvedList,
) -> ResolvedList {
    crate::formula::eval_budget::validation(|| {
        let value = bind(source).and_then(|expr| resolve_expr(&expr, lookup, &bind, 0));
        match value {
            Ok(Source::Reference(sheet, range)) => read(&sheet, &range),
            Ok(Source::Value(value)) => values_to_list(value),
            Err(error) => ResolvedList::failed(error),
        }
    })
}

enum Source {
    Reference(SheetRef, CellRange),
    Value(EvalResult),
}

fn reference(
    sheet: SheetRef,
    r0: usize,
    c0: usize,
    r1: usize,
    c1: usize,
) -> Result<Source, String> {
    if matches!(sheet, SheetRef::RefError { .. })
        || r0 >= NUM_ROWS
        || r1 >= NUM_ROWS
        || c0 >= NUM_COLS
        || c1 >= NUM_COLS
    {
        return Err("#REF! Invalid validation list reference".into());
    }
    Ok(Source::Reference(
        sheet,
        CellRange::new(r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1)),
    ))
}

fn scalar_number<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> Result<f64, String> {
    let value = evaluate(expr, lookup).to_number()?;
    if !value.is_finite() {
        return Err("#NUM! Non-finite list reference argument".into());
    }
    Ok(value.trunc())
}

fn resolve_expr<L: CellLookup>(
    expr: &BoundExpr,
    lookup: &L,
    bind: &impl Fn(&str) -> Result<BoundExpr, String>,
    depth: usize,
) -> Result<Source, String> {
    if depth >= 64 {
        return Err("Validation list reference nesting limit exceeded".into());
    }
    let resolve = |e: &BoundExpr| resolve_expr(e, lookup, bind, depth + 1);
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
            let target = lookup
                .resolve_named_reference(name)
                .ok_or_else(|| format!("#NAME? Unknown validation list source: {name}"))?;
            resolve(&target)
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
                    return Err("R1C1 validation list references are not supported".into());
                }
            }
            let result = evaluate(&args[0], lookup);
            if let EvalResult::Error(error) = result {
                return Err(error);
            }
            if result.dimensions() != (1, 1) {
                return Err("INDIRECT requires one reference string".into());
            }
            let indirect = bind(&result.to_text())?;
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
            resolve(&indirect)
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

fn values_to_list(value: EvalResult) -> ResolvedList {
    let mut items = Vec::new();
    let mut push = |value: &Value| -> Result<(), String> {
        match value {
            Value::Error(e) => return Err(e.clone()),
            Value::Number(n) if !n.is_finite() => return Err("#NUM! Non-finite list value".into()),
            _ => {}
        }
        let text = value.to_text();
        if !text.is_empty() {
            items.push(text);
        }
        Ok(())
    };
    match value {
        EvalResult::Array(array) => {
            if array.rows() > 1 && array.cols() > 1 {
                return ResolvedList::failed(
                    "Validation list formula must return one row or column",
                );
            }
            for r in 0..array.rows() {
                for c in 0..array.cols() {
                    if let Err(error) = push(array.get(r, c).unwrap_or(&Value::Empty)) {
                        return ResolvedList::failed(error);
                    }
                }
            }
        }
        scalar => {
            if let Err(error) = push(&scalar.to_value()) {
                return ResolvedList::failed(error);
            }
        }
    }
    // The evaluator bounds temporary arrays before they allocate. Keep only one
    // extra item here so the shared constructor can report dropdown truncation.
    items.truncate(MAX_LIST_ITEMS + 1);
    ResolvedList::from_items(items)
}
