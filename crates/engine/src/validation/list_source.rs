//! Resolve dropdown expressions without materializing reference ranges. This
//! retains unformatted value labels and sparse whole-column behavior for dynamic refs.
use super::{CellRange, ResolvedList, MAX_LIST_ITEMS};
use crate::formula::{
    eval::{CellLookup, EvalResult, Value},
    reference::{self, Source},
};
use crate::sheet::SheetRef;

pub(crate) fn resolve<L: CellLookup>(
    source: &str,
    lookup: &L,
    read: impl Fn(&SheetRef, &CellRange) -> ResolvedList,
) -> ResolvedList {
    crate::formula::eval_budget::validation(|| {
        let value = lookup
            .bind_reference_text(source)
            .and_then(|expr| reference::resolve(&expr, lookup));
        match value {
            Ok(Source::Reference(sheet, range)) => read(
                &sheet,
                &CellRange::new(
                    range.start_row,
                    range.start_col,
                    range.end_row,
                    range.end_col,
                ),
            ),
            Ok(Source::Value(value)) => values_to_list(value),
            Err(error) => ResolvedList::failed(error),
        }
    })
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
