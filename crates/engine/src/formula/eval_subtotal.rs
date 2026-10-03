//! Filter-aware subtotal evaluation; stored Table criteria determine visibility.
use super::{
    eval::{evaluate, CellLookup, EvalResult, NamedRangeResolution, Value},
    parser::{BoundExpr, Expr},
};
use crate::sheet::SheetRef;

pub(crate) fn contains_subtotal<S>(expr: &Expr<S>) -> bool {
    match expr {
        Expr::Function { name, args } => {
            name.eq_ignore_ascii_case("SUBTOTAL") || args.iter().any(contains_subtotal)
        }
        Expr::BinaryOp { left, right, .. } => contains_subtotal(left) || contains_subtotal(right),
        _ => false,
    }
}

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str,
    args: &[BoundExpr],
    lookup: &L,
) -> Option<EvalResult> {
    if name != "SUBTOTAL" {
        return None;
    }
    Some(run(args, lookup).unwrap_or_else(EvalResult::Error))
}
fn run<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> Result<EvalResult, String> {
    if args.len() < 2 {
        return Err("#VALUE! SUBTOTAL needs a function number and references".into());
    }
    let code = evaluate(&args[0], lookup).to_number()?;
    if !code.is_finite() {
        return Err("#VALUE! Invalid SUBTOTAL function number".into());
    }
    let code = code.trunc() as i32;
    let ignore_hidden = (101..=111).contains(&code);
    let code = if (101..=111).contains(&code) {
        code - 100
    } else {
        code
    };
    if !(1..=11).contains(&code) {
        return Err("#VALUE! Invalid SUBTOTAL function number".into());
    }
    let mut nums = Vec::new();
    let mut count = 0usize;
    for arg in &args[1..] {
        let (sheet, r0, c0, r1, c1) = match arg {
            Expr::EmptyRange { .. } => continue,
            Expr::CellRef {
                sheet, row, col, ..
            } => (sheet, *row, *col, *row, *col),
            Expr::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => (sheet, *start_row, *start_col, *end_row, *end_col),
            Expr::NamedRange(name) => match lookup.resolve_named_range(name) {
                Some(NamedRangeResolution::Cell { row, col }) => {
                    (&SheetRef::Current, row, col, row, col)
                }
                Some(NamedRangeResolution::Range {
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                }) => (&SheetRef::Current, start_row, start_col, end_row, end_col),
                None => return Err(format!("#NAME? Unknown range {name}")),
            },
            Expr::RefError => return Err("#REF!".into()),
            Expr::ReferenceError(e) => return Err(e.clone()),
            _ => return Err("#VALUE! SUBTOTAL requires cell or bounded range references".into()),
        };
        if matches!(sheet, SheetRef::RefError { .. }) {
            return Err("#REF!".into());
        }
        let (rows, cols) = lookup.data_bounds(sheet);
        if rows == 0 || cols == 0 {
            continue;
        }
        for row in r0..=r1.min(rows - 1) {
            for col in c0..=c1.min(cols - 1) {
                if lookup.subtotal_skip_cell(sheet, row, col, ignore_hidden) {
                    continue;
                }
                let value = match sheet {
                    SheetRef::Current => lookup.get_typed(row, col),
                    SheetRef::Id(id) => lookup.get_typed_sheet(*id, row, col),
                    SheetRef::RefError { .. } => return Err("#REF!".into()),
                };
                if !matches!(value, Value::Empty) {
                    count += 1;
                }
                match value {
                    Value::Number(n) => nums.push(n),
                    Value::Error(e) if code != 2 && code != 3 => return Err(e),
                    _ => {}
                }
            }
        }
    }
    let n = nums.len();
    let sum: f64 = nums.iter().sum();
    let result = match code {
        1 if n == 0 => return Err("#DIV/0!".into()),
        1 => sum / n as f64,
        2 => n as f64,
        3 => count as f64,
        4 => nums.iter().copied().reduce(f64::max).unwrap_or(0.0),
        5 => nums.iter().copied().reduce(f64::min).unwrap_or(0.0),
        6 => {
            if n == 0 {
                0.0
            } else {
                nums.iter().product()
            }
        }
        7 | 8 | 10 | 11 => {
            let sample = code == 7 || code == 10;
            if n <= usize::from(sample) {
                return Err("#DIV/0!".into());
            }
            let mean = sum / n as f64;
            let variance = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>()
                / (n - usize::from(sample)) as f64;
            if code <= 8 {
                variance.sqrt()
            } else {
                variance
            }
        }
        9 => sum,
        _ => unreachable!(),
    };
    Ok(EvalResult::Number(result))
}
