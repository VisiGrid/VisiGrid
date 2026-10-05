//! Validation reads cached typed values. Only the hypothetical-input API needs
//! a private recalculation, so normal desktop marker refresh never copies the
//! workbook per cell.
use super::{Workbook, WorkbookLookup};
use crate::formula::{
    eval::{evaluate, EvalResult, Value},
    parser::{bind_expr, parse},
};
use crate::validation::{
    evaluation::{input_value, invalid, resolve_constraint, validate_rule},
    *,
};

impl Workbook {
    fn validation_formula(&self, sheet: usize, row: usize, col: usize, source: &str) -> EvalResult {
        let Some(id) = self.sheet_id_at_idx(sheet) else {
            return EvalResult::Error("#REF!".into());
        };
        let formula = format!("={}", source.trim().trim_start_matches('='));
        crate::formula::eval_budget::validation(|| match parse(&formula) {
            Ok(expr) => evaluate(
                &bind_expr(&expr, |name| self.sheet_id_by_name(name)),
                &WorkbookLookup::with_cell_context(self, id, row, col),
            ),
            Err(error) => EvalResult::Error(error),
        })
    }

    /// Resolve a numeric constraint in workbook context. Without a target
    /// coordinate ROW()/COLUMN() use A1; validation itself passes its target.
    pub fn resolve_constraint_value(
        &self,
        sheet: usize,
        value: &ConstraintValue,
    ) -> Result<f64, ConstraintResolveError> {
        resolve_constraint(value, &|source| {
            self.validation_formula(sheet, 0, 0, source)
        })
    }

    fn validate_typed_value(
        &self,
        sheet: usize,
        row: usize,
        col: usize,
        value: &Value,
        list_text: &str,
    ) -> ValidationResult {
        let Some(rule) = self
            .sheets
            .get(sheet)
            .and_then(|s| s.validations.get(row, col))
        else {
            return ValidationResult::Valid;
        };
        validate_rule(
            &rule.at(row, col),
            value,
            list_text,
            |source| self.validation_formula(sheet, row, col, source),
            |_| {
                self.get_list_items(sheet, row, col)
                    .unwrap_or_else(ResolvedList::empty)
            },
        )
    }

    /// Check an existing cell using its typed computed value. This does not
    /// recalculate: callers must settle ordinary workbook edits first.
    pub fn validate_cell(&self, sheet: usize, row: usize, col: usize) -> ValidationResult {
        let Some(s) = self.sheets.get(sheet) else {
            return ValidationResult::Valid;
        };
        if s.validations.get(row, col).is_none() {
            return ValidationResult::Valid;
        }
        self.validate_typed_value(
            sheet,
            row,
            col,
            &s.get_computed_value(row, col),
            &s.get_display(row, col),
        )
    }

    /// Check proposed input without changing cells, caches, history or revision.
    /// Formula-backed rules use a private recalculated candidate, including
    /// indirect dependencies and spills. Prefer validate_cell after a write or
    /// for batch marker refresh; this preview can cost one workbook recalc.
    pub fn validate_cell_input(
        &self,
        sheet: usize,
        row: usize,
        col: usize,
        input: &str,
    ) -> ValidationResult {
        let Some(s) = self.sheets.get(sheet) else {
            return ValidationResult::Valid;
        };
        let Some(rule) = s.validations.get(row, col) else {
            return ValidationResult::Valid;
        };
        let rule = rule.at(row, col);
        if rule.ignore_blank && input.trim().is_empty() {
            return ValidationResult::Valid;
        }
        let typed = match input_value(&rule, input) {
            Ok(value) => value,
            Err(result) => return result,
        };
        let formula_input = input.trim_start().starts_with('=');
        if !rule.has_reference_sources()
            && !matches!(
                rule.rule_type,
                ValidationType::List(ListSource::NamedRange(_))
            )
            && !formula_input
        {
            return self.validate_typed_value(sheet, row, col, &typed, input);
        }
        if row >= s.rows
            || col >= s.cols
            || s.is_pivot_owned(row, col)
            || s.table_value_write_error(row, col).is_some()
            || s.get_merge(row, col).is_some_and(|m| m.start != (row, col))
            || s.is_spill_receiver(row, col)
        {
            return invalid(
                &rule,
                "Cannot preview input in a protected, merged or spilled cell",
            );
        }
        let mut candidate = self.clone();
        candidate.sheets[sheet].set_value_deferred(row, col, input);
        candidate.rebuild_dep_graph();
        // Validation must not invoke host custom functions while previewing.
        crate::formula::eval_budget::validation(|| candidate.recompute_full_ordered_inner(None));
        if formula_input {
            candidate.validate_cell(sheet, row, col)
        } else {
            candidate.validate_typed_value(sheet, row, col, &typed, input)
        }
    }
}

impl Workbook {
    pub(super) fn resolve_validation_list(
        &self,
        sheet: usize,
        row: usize,
        col: usize,
        source: &str,
    ) -> ResolvedList {
        let Some(id) = self.sheet_id_at_idx(sheet) else {
            return ResolvedList::failed("#REF! Missing worksheet");
        };
        crate::validation::list_source::resolve(
            source,
            &WorkbookLookup::with_cell_context(self, id, row, col),
            |source| {
                parse(&format!("={}", source.trim().trim_start_matches('=')))
                    .map(|expr| bind_expr(&expr, |name| self.sheet_id_by_name(name)))
            },
            |target, range| {
                let target = match target {
                    crate::sheet::SheetRef::Current => Some(id),
                    crate::sheet::SheetRef::Id(id) => Some(*id),
                    crate::sheet::SheetRef::RefError { .. } => None,
                };
                target
                    .and_then(|id| self.sheet_by_id(id))
                    .map(|sheet| sheet.resolve_list_cells(range))
                    .unwrap_or_else(|| ResolvedList::failed("#REF! Missing list worksheet"))
            },
        )
    }
}
