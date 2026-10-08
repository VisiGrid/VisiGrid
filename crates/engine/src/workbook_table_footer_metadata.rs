//! Move rule coverage with footer cells, preserving surrounding anchors and
//! precedence. Work is bounded by rectangles and reference boundaries, never
//! by the number of worksheet cells covered by a rule.
use super::Movement;
use crate::{
    cond_format::{CondFormatRule, CondFormatStore},
    formula::parser::{self, Expr, ParsedExpr},
    sheet::{Sheet, UnboundSheetRef, NUM_COLS, NUM_ROWS},
    validation::{CellRange, ValidationRule, ValidationStore},
};
use std::collections::BTreeSet;

const MAX_FRAGMENTS: usize = 4096;
struct Budget(usize);
impl Budget {
    fn step(&mut self, n: usize) -> Result<(), String> {
        self.0 = self.0.checked_sub(n).ok_or_else(limit)?;
        Ok(())
    }
}
fn limit() -> String {
    "Moving totals would create too many rule fragments. Simplify the conditional formatting or validation first.".into()
}
fn subtract(r: CellRange, c: CellRange) -> Vec<CellRange> {
    if !r.overlaps(&c) {
        return vec![r];
    }
    let (top, bottom, left, right) = (
        r.start_row.max(c.start_row),
        r.end_row.min(c.end_row),
        r.start_col.max(c.start_col),
        r.end_col.min(c.end_col),
    );
    let mut out = Vec::new();
    if r.start_row < top {
        out.push(CellRange {
            end_row: top - 1,
            ..r
        });
    }
    if bottom < r.end_row {
        out.push(CellRange {
            start_row: bottom + 1,
            ..r
        });
    }
    if r.start_col < left {
        out.push(CellRange::new(top, r.start_col, bottom, left - 1));
    }
    if right < r.end_col {
        out.push(CellRange::new(top, right + 1, bottom, r.end_col));
    }
    out
}
fn without(
    mut ranges: Vec<CellRange>,
    cuts: &[CellRange],
    budget: &mut Budget,
) -> Result<Vec<CellRange>, String> {
    for &cut in cuts {
        budget.step(ranges.len())?;
        ranges = ranges.into_iter().flat_map(|r| subtract(r, cut)).collect();
        if ranges.len() > MAX_FRAGMENTS {
            return Err(limit());
        }
    }
    Ok(ranges)
}
fn sources(rule: &ValidationRule) -> Vec<String> {
    let mut result = Vec::new();
    rule.clone().map_references(|s| {
        result.push(s.to_owned());
        s.to_owned()
    });
    result
}
fn rebase(source: &str, row: i32, col: i32) -> Result<String, String> {
    let input = format!("={}", source.trim_start_matches('='));
    if parser::parse(&input).is_ok() {
        return Ok(parser::adjust_formula_refs(source, row, col));
    }
    let mut result = source.to_owned();
    for (span, expr) in crate::formula::source_refs::references(source)?.into_iter().rev() {
        let input = parser::format_parsed_expr(&expr);
        let adjusted = parser::adjust_formula_refs(&input, row, col);
        result.replace_range(span, adjusted.trim_start_matches('='));
    }
    Ok(result)
}
impl Movement {
    fn source_range(&self) -> CellRange {
        CellRange::new(
            self.footer.start_row,
            self.footer.start_col,
            self.footer.end_row,
            self.footer.end_col,
        )
    }
    fn destination_range(&self) -> CellRange {
        CellRange {
            start_row: self.destination,
            end_row: self.destination,
            ..self.source_range()
        }
    }
    fn spanning(&self, range: CellRange) -> bool {
        range.start_row < self.footer.start_row && range.end_row >= self.footer.end_row
    }
    fn mapped(&self, r: CellRange, local: bool, coverage: CellRange) -> Option<CellRange> {
        // A rule covering both sides of the footer describes worksheet space,
        // like A:A. Moving a footer must not punch a hole in that coverage.
        if !local { return Some(r); }
        if self.spanning(coverage) {
            // Body + footer coverage follows the footer boundary. Rules that
            // extend past it describe worksheet space and retain their extent.
            if coverage.end_row == self.footer.end_row
                && r.start_col >= self.footer.start_col && r.end_col <= self.footer.end_col
            {
                if r.start_row == self.footer.start_row {
                    return Some(CellRange {
                        start_row: r.start_row.min(self.destination),
                        end_row: self.destination,
                        ..r
                    });
                }
                // Shrinking releases body records in place; keep their rules.
                return Some(r);
            }
            return Some(r);
        }
        if self.source_range().contains(r.start_row, r.start_col) {
            Some(CellRange {
                start_row: self.destination,
                end_row: self.destination,
                ..r
            })
        } else if self.destination_range().contains(r.start_row, r.start_col) {
            None
        } else {
            Some(r)
        }
    }

    /// Split where a relative reference enters/leaves the old or new footer,
    /// or the worksheet bounds. One rebased predicate then represents every
    /// cell of a fragment, including rules whose stored anchor is far away.
    fn pieces(
        &self,
        range: CellRange,
        origin: Option<(usize, usize)>,
        sources: &[String],
        local: bool,
        budget: &mut Budget,
    ) -> Result<Vec<CellRange>, String> {
        let mut rows = BTreeSet::from([range.start_row, range.end_row + 1]);
        let mut cols = BTreeSet::from([range.start_col, range.end_col + 1]);
        fn cut(set: &mut BTreeSet<usize>, lo: usize, hi: usize, value: i64) {
            if value > lo as i64 && value <= hi as i64 {
                set.insert(value as usize);
            }
        }
        if local {
            for r in [self.source_range(), self.destination_range()] {
                cut(
                    &mut rows,
                    range.start_row,
                    range.end_row,
                    r.start_row as i64,
                );
                cut(
                    &mut rows,
                    range.start_row,
                    range.end_row,
                    r.end_row as i64 + 1,
                );
                cut(
                    &mut cols,
                    range.start_col,
                    range.end_col,
                    r.start_col as i64,
                );
                cut(
                    &mut cols,
                    range.start_col,
                    range.end_col,
                    r.end_col as i64 + 1,
                );
            }
        }
        if let Some((ar, ac)) = origin {
            let mut visit =
                |sheet: &UnboundSheetRef, row: usize, col: usize, ra: bool, ca: bool| {
                    let same = match sheet {
                        UnboundSheetRef::Current => local,
                        UnboundSheetRef::Named(name) => name.eq_ignore_ascii_case(&self.owner),
                    };
                    if !ra {
                        let delta = ar as i64 - row as i64;
                        for r in [0, NUM_ROWS] {
                            cut(&mut rows, range.start_row, range.end_row, delta + r as i64);
                        }
                        if same {
                            for r in [
                                self.footer.start_row,
                                self.footer.start_row + 1,
                                self.destination,
                                self.destination + 1,
                            ] {
                                cut(&mut rows, range.start_row, range.end_row, delta + r as i64);
                            }
                        }
                    }
                    if !ca {
                        let delta = ac as i64 - col as i64;
                        for c in [0, NUM_COLS] {
                            cut(&mut cols, range.start_col, range.end_col, delta + c as i64);
                        }
                        if same {
                            for c in [self.footer.start_col, self.footer.end_col + 1] {
                                cut(&mut cols, range.start_col, range.end_col, delta + c as i64);
                            }
                        }
                    }
                };
            fn refs(
                expr: &ParsedExpr,
                f: &mut impl FnMut(&UnboundSheetRef, usize, usize, bool, bool),
                b: &mut Budget,
            ) -> Result<(), String> {
                b.step(1)?;
                match expr {
                    Expr::CellRef {
                        sheet,
                        row,
                        col,
                        row_abs,
                        col_abs,
                    } => f(sheet, *row, *col, *row_abs, *col_abs),
                    Expr::Range {
                        sheet,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        start_row_abs,
                        start_col_abs,
                        end_row_abs,
                        end_col_abs,
                    } => {
                        f(
                            sheet,
                            *start_row,
                            *start_col,
                            *start_row_abs,
                            *start_col_abs,
                        );
                        f(sheet, *end_row, *end_col, *end_row_abs, *end_col_abs);
                    }
                    Expr::Function { args, .. } => {
                        for arg in args {
                            refs(arg, f, b)?;
                        }
                    }
                    Expr::BinaryOp { left, right, .. } => {
                        refs(left, f, b)?;
                        refs(right, f, b)?;
                    }
                    _ => {}
                }
                Ok(())
            }
            for source in sources {
                let input = if source.starts_with('=') {
                    source.clone()
                } else {
                    format!("={source}")
                };
                match parser::parse(&input) {
                    Ok(expr) => refs(&expr, &mut visit, budget)?,
                    Err(_) => for (_, expr) in crate::formula::source_refs::references(source)? {
                        refs(&expr, &mut visit, budget)?;
                    },
                }
            }
        }
        let count = (rows.len() - 1).saturating_mul(cols.len() - 1);
        if count > MAX_FRAGMENTS {
            return Err(limit());
        }
        budget.step(count)?;
        let rows: Vec<_> = rows.into_iter().collect();
        let cols: Vec<_> = cols.into_iter().collect();
        Ok(rows
            .windows(2)
            .flat_map(|r| {
                cols.windows(2)
                    .map(move |c| CellRange::new(r[0], c[0], r[1] - 1, c[1] - 1))
            })
            .collect())
    }

    fn rewrite_validation(
        &self,
        rule: &mut ValidationRule,
        local: bool,
        guarded: &mut bool,
    ) -> Result<bool, String> {
        let mut result = Ok(false);
        rule.map_references(|s| {
            let mut source = s.to_owned();
            match self.rewrite(&mut source, local, guarded) {
                Ok(changed) => {
                    if let Ok(any) = &mut result {
                        *any |= changed;
                    }
                }
                Err(error) => result = Err(error),
            }
            source
        });
        result
    }
    fn validation_pieces(
        &self,
        range: CellRange,
        coverage: CellRange,
        rule: &ValidationRule,
        local: bool,
        guarded: &mut bool,
        budget: &mut Budget,
    ) -> Result<Vec<(CellRange, ValidationRule)>, String> {
        let pieces = self.pieces(range, rule.reference_origin, &sources(rule), local, budget)?;
        let mut result = Vec::new();
        let mut changed = false;
        for piece in pieces {
            let Some(target) = self.mapped(piece, local, coverage) else {
                changed = true;
                continue;
            };
            let mut rebased = rule.clone();
            if let Some((row, col)) = rule.reference_origin {
                let mut error = None;
                rebased.map_references(|source| {
                    match rebase(source, piece.start_row as i32 - row as i32, piece.start_col as i32 - col as i32) {
                        Ok(result) => result,
                        Err(e) => { error = Some(e); source.into() },
                    }
                });
                if let Some(error) = error { return Err(error); }
                rebased.reference_origin = Some((piece.start_row, piece.start_col));
            }
            let rewritten = self.rewrite_validation(&mut rebased, local, guarded)?;
            if rewritten || target != piece {
                changed = true;
                if rebased.reference_origin.is_some() {
                    rebased.reference_origin = Some((target.start_row, target.start_col));
                }
                result.push((target, rebased));
            } else {
                result.push((target, rule.clone()));
            }
        }
        if changed {
            *guarded = true;
            Ok(result)
        } else {
            Ok(vec![(range, rule.clone())])
        }
    }
    fn validations(
        &self,
        store: &ValidationStore,
        local: bool,
        guarded: &mut bool,
        budget: &mut Budget,
    ) -> Result<ValidationStore, String> {
        let rules: Vec<_> = store.iter().map(|(r, rule)| (*r, rule.clone())).collect();
        let mut affected = BTreeSet::new();
        for (i, (range, rule)) in rules.iter().enumerate() {
            let transformed = self.validation_pieces(*range, *range, rule, local, guarded, budget)?;
            if transformed != vec![(*range, rule.clone())] {
                affected.insert(i);
            }
        }
        // Preserve first-rule precedence after splitting/rekeying an overlap
        // component. Unconnected definitions remain byte-for-byte unchanged.
        let mut pending: Vec<_> = affected.iter().copied().collect();
        while let Some(i) = pending.pop() {
            for (j, (other, _)) in rules.iter().enumerate() {
                budget.step(1)?;
                if !affected.contains(&j) && rules[i].0.overlaps(other) {
                    affected.insert(j);
                    pending.push(j);
                }
            }
        }
        let mut result = store.clone();
        for &i in &affected {
            result.remove(&rules[i].0);
        }
        let mut earlier = Vec::new();
        let mut fragments = 0;
        for i in affected {
            let (range, rule) = &rules[i];
            for piece in without(vec![*range], &earlier, budget)? {
                for (target, rule) in self.validation_pieces(piece, *range, rule, local, guarded, budget)? {
                    fragments += 1;
                    if fragments > MAX_FRAGMENTS {
                        return Err(limit());
                    }
                    result.set(target, rule);
                }
            }
            earlier.push(*range);
        }
        if local {
            let affected: Vec<_> = store
                .exclusions_iter()
                .copied()
                .filter(|range| {
                    range.overlaps(&self.source_range())
                        || range.overlaps(&self.destination_range())
                })
                .collect();
            for range in &affected {
                result.remove_exclusion(range);
            }
            for range in affected {
                *guarded = true;
                for piece in self.pieces(range, None, &[], local, budget)? {
                    if let Some(target) = self.mapped(piece, local, range) {
                        result.exclude(target);
                        fragments += 1;
                    }
                    if fragments > MAX_FRAGMENTS {
                        return Err(limit());
                    }
                }
            }
        }
        Ok(result)
    }
    fn conditional_formats(
        &self,
        store: &CondFormatStore,
        local: bool,
        guarded: &mut bool,
        budget: &mut Budget,
    ) -> Result<CondFormatStore, String> {
        let mut next_id = store.fragment_id_start()?;
        let mut result = store.clone();
        let mut position = 0;
        let mut fragments = 0;
        for rule in store.iter() {
            let mut changed = false;
            let mut transformed = Vec::new();
            let mut earlier = Vec::new();
            for range in &rule.ranges {
                for unique in without(vec![*range], &earlier, budget)? {
                    for piece in self.pieces(
                        unique,
                        Some((range.start_row, range.start_col)),
                        std::slice::from_ref(&rule.predicate),
                        local,
                        budget,
                    )? {
                        let Some(target) = self.mapped(piece, local, *range) else {
                            changed = true;
                            continue;
                        };
                        let mut source = rebase(
                            &rule.predicate,
                            (piece.start_row - range.start_row) as i32,
                            (piece.start_col - range.start_col) as i32,
                        )?;
                        let rewritten = self.rewrite(&mut source, local, guarded)?;
                        changed |= rewritten || target != piece;
                        transformed.push((target, source));
                        if transformed.len() > MAX_FRAGMENTS {
                            return Err(limit());
                        }
                    }
                }
                earlier.push(*range);
            }
            if !changed {
                position += 1;
                continue;
            }
            *guarded = true;
            result.remove(rule.id);
            for (n, (target, source)) in transformed.into_iter().enumerate() {
                fragments += 1;
                if fragments > MAX_FRAGMENTS {
                    return Err(limit());
                }
                let id = if n == 0 {
                    rule.id
                } else {
                    let id = next_id;
                    next_id = next_id
                        .checked_add(1)
                        .ok_or("Conditional-format rule IDs are exhausted. Nothing was changed.")?;
                    id
                };
                let mut fragment =
                    CondFormatRule::new(id, vec![target], source, rule.style.clone());
                fragment.enabled = rule.enabled;
                result.insert_at(position, fragment);
                position += 1;
            }
        }
        Ok(result)
    }
    pub(super) fn rewrite_rule_stores(
        &self,
        sheet: &mut Sheet,
        local: bool,
        guarded: &mut bool,
    ) -> Result<bool, String> {
        let valid = |r: &CellRange| {
            r.start_row <= r.end_row
                && r.start_col <= r.end_col
                && r.end_row < sheet.rows.min(NUM_ROWS)
                && r.end_col < sheet.cols.min(NUM_COLS)
        };
        if sheet.validations.iter().any(|(r, rule)| {
            !valid(r)
                || rule.reference_origin.is_some_and(|(r, c)| {
                    r >= sheet.rows.min(NUM_ROWS) || c >= sheet.cols.min(NUM_COLS)
                })
        }) || sheet.validations.exclusions_iter().any(|r| !valid(r))
            || sheet
                .cond_formats
                .iter()
                .flat_map(|rule| &rule.ranges)
                .any(|r| !valid(r))
        {
            return Err(
                "A formatting or validation range is outside the worksheet. Nothing was changed."
                    .into(),
            );
        }
        let mut budget = Budget(1_000_000);
        let before_guard = *guarded;
        // Use a local flag: earlier cell/name links must not mask a rule change.
        let mut rules_guarded = false;
        let validations =
            self.validations(&sheet.validations, local, &mut rules_guarded, &mut budget)?;
        let formats =
            self.conditional_formats(&sheet.cond_formats, local, &mut rules_guarded, &mut budget)?;
        if rules_guarded {
            sheet.validations = validations;
            sheet.cond_formats = formats;
        }
        *guarded = before_guard || rules_guarded;
        Ok(rules_guarded)
    }
}
