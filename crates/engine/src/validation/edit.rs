//! Exact worksheet-range edits. Callers resolve visible selections to stored
//! ranges before planning; this module never interprets screen coordinates.
use super::{CellRange, ConstraintValue, ValidationRule, ValidationStore, ValidationType};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet};

const MAX_FRAGMENTS: usize = 4096;
const MAX_WORK: usize = 1_000_000;

#[derive(Debug, Clone)]
pub enum ValidationEdit {
    /// Replace rules in the target cells. Explicit exclusions remain in force.
    Set(ValidationRule),
    Clear,
    Exclude,
    ClearExclusions,
}

/// Sparse, reversible metadata changes with a guard for the entire validation
/// store. A newly introduced overlapping rule must not silently change replay.
/// Cells, Table definitions and other sheet metadata are not retained here.
#[derive(Debug, Clone)]
pub struct ValidationPatch {
    rules: Vec<(CellRange, Option<ValidationRule>, Option<ValidationRule>)>,
    exclusions: Vec<(CellRange, bool, bool)>,
    before: [u8; 32],
    after: [u8; 32],
}

impl ValidationPatch {
    pub fn is_empty(&self) -> bool {
        self.rules.is_empty() && self.exclusions.is_empty()
    }

    pub fn changed_ranges(&self) -> impl Iterator<Item = &CellRange> {
        self.rules
            .iter()
            .map(|(r, _, _)| r)
            .chain(self.exclusions.iter().map(|(r, _, _)| r))
    }

    /// Validate before writing anything. `forward = false` restores the exact
    /// original overlapping ranges, including formerly shadowed definitions.
    pub fn apply(&self, store: &mut ValidationStore, forward: bool) -> Result<(), String> {
        let expected = if forward { self.before } else { self.after };
        if fingerprint(store)? != expected {
            return Err("Validation rules changed since this edit. Nothing was applied.".into());
        }
        for (range, before, after) in &self.rules {
            let value = if forward { after } else { before };
            if let Some(rule) = value {
                store.rules.insert(*range, rule.clone());
            } else {
                store.rules.remove(range);
            }
        }
        for (range, before, after) in &self.exclusions {
            if if forward { *after } else { *before } {
                store.exclusions.insert(*range);
            } else {
                store.exclusions.remove(range);
            }
        }
        Ok(())
    }
}

impl ValidationStore {
    /// Plan an atomic edit to exact stored cells. Partial edits split rectangles
    /// rather than deleting whole overlapping rules or exclusions. Only an
    /// affected overlap component is normalized: its effective first-rule
    /// precedence is preserved outside the targets. Validation references keep
    /// their existing fixed worksheet meaning (unlike relative CF predicates).
    ///
    /// Planning is bounded by rectangle count and intersection work, not by
    /// worksheet area. A full-column rule does not require per-cell expansion.
    pub fn plan_edit(
        &self,
        targets: &[CellRange],
        edit: ValidationEdit,
        rows: usize,
        cols: usize,
    ) -> Result<ValidationPatch, String> {
        let mut budget = Budget(MAX_WORK);
        if targets.len() > MAX_FRAGMENTS {
            return Err(limit_error());
        }
        for range in targets
            .iter()
            .chain(self.rules.keys())
            .chain(self.exclusions.iter())
        {
            if range.start_row > range.end_row
                || range.start_col > range.end_col
                || range.end_row >= rows
                || range.end_col >= cols
            {
                return Err("A validation range is outside the worksheet.".into());
            }
        }
        let targets = disjoint(targets.iter().copied(), &mut budget)?;
        let mut next = self.clone();
        match edit {
            ValidationEdit::Set(rule) => {
                if !self.targets_have_rule(&targets, &rule, &mut budget)? {
                    next.rules = self.without_targets(&targets, &mut budget)?;
                    for range in targets {
                        next.rules.insert(range, rule.clone());
                    }
                }
            }
            ValidationEdit::Clear => {
                next.rules = self.without_targets(&targets, &mut budget)?;
            }
            ValidationEdit::Exclude => {
                // Keep existing definitions; only add uncovered portions.
                let cuts: Vec<_> = self.exclusions.iter().copied().collect();
                for range in targets {
                    next.exclusions
                        .extend(subtract_all(vec![range], &cuts, &mut budget)?);
                }
            }
            ValidationEdit::ClearExclusions => {
                next.exclusions.clear();
                for &range in &self.exclusions {
                    next.exclusions
                        .extend(subtract_all(vec![range], &targets, &mut budget)?);
                    if next.exclusions.len() > self.exclusions.len().saturating_add(MAX_FRAGMENTS) {
                        return Err(limit_error());
                    }
                }
            }
        }
        let mut rules = Vec::new();
        for range in self
            .rules
            .keys()
            .chain(next.rules.keys())
            .copied()
            .collect::<BTreeSet<_>>()
        {
            let before = self.rules.get(&range);
            let after = next.rules.get(&range);
            if before != after {
                rules.push((range, before.cloned(), after.cloned()));
            }
        }
        let mut exclusions = Vec::new();
        for range in self.exclusions.union(&next.exclusions) {
            let before = self.exclusions.contains(range);
            let after = next.exclusions.contains(range);
            if before != after {
                exclusions.push((*range, before, after));
            }
        }
        if rules.len().saturating_add(exclusions.len()) > MAX_FRAGMENTS {
            return Err(limit_error());
        }
        Ok(ValidationPatch {
            rules,
            exclusions,
            before: fingerprint(self)?,
            after: fingerprint(&next)?,
        })
    }

    fn targets_have_rule(
        &self,
        targets: &[CellRange],
        rule: &ValidationRule,
        budget: &mut Budget,
    ) -> Result<bool, String> {
        let mut remaining = targets.to_vec();
        for (&range, existing) in &self.rules {
            let mut next = Vec::new();
            for piece in remaining {
                budget.step()?;
                if piece.overlaps(&range) && existing != rule {
                    return Ok(false);
                }
                subtract(piece, range, &mut next);
                if next.len() > MAX_FRAGMENTS {
                    return Err(limit_error());
                }
            }
            remaining = next;
            if remaining.is_empty() {
                return Ok(true);
            }
        }
        Ok(remaining.is_empty())
    }

    fn without_targets(
        &self,
        targets: &[CellRange],
        budget: &mut Budget,
    ) -> Result<BTreeMap<CellRange, ValidationRule>, String> {
        // Fragment keys can sort differently from their parent keys. Include
        // every connected overlap before normalizing so a formerly lower
        // priority rule cannot win outside the edited cells after a split.
        let mut affected = BTreeSet::new();
        let mut pending = Vec::new();
        for &range in self.rules.keys() {
            for target in targets {
                budget.step()?;
                if range.overlaps(target) {
                    affected.insert(range);
                    pending.push(range);
                    break;
                }
            }
        }
        while let Some(range) = pending.pop() {
            for &other in self.rules.keys() {
                budget.step()?;
                if !affected.contains(&other) && range.overlaps(&other) {
                    affected.insert(other);
                    pending.push(other);
                    if affected.len() > MAX_FRAGMENTS {
                        return Err(limit_error());
                    }
                }
            }
        }
        let mut result = self.rules.clone();
        let mut earlier = Vec::new();
        let mut fragments = 0usize;
        for range in &affected {
            result.remove(range);
        }
        for range in affected {
            let pieces = subtract_all(vec![range], targets, budget)?;
            let pieces = subtract_all(pieces, &earlier, budget)?;
            fragments = fragments.saturating_add(pieces.len());
            if fragments > MAX_FRAGMENTS {
                return Err(limit_error());
            }
            for piece in pieces {
                result.insert(piece, self.rules[&range].clone());
            }
            earlier.push(range);
        }
        Ok(result)
    }
}

fn fingerprint(store: &ValidationStore) -> Result<[u8; 32], String> {
    // serde_json encodes all non-finite numbers as null. Refuse them instead of
    // allowing different invalid constraints to share a stale-history guard.
    for rule in store.rules.values() {
        match &rule.rule_type {
            ValidationType::WholeNumber(c)
            | ValidationType::Decimal(c)
            | ValidationType::Date(c)
            | ValidationType::Time(c)
            | ValidationType::TextLength(c) => {
                for value in std::iter::once(&c.value1).chain(c.value2.iter()) {
                    if matches!(value, ConstraintValue::Number(n) if !n.is_finite()) {
                        return Err("Validation constraints must contain finite numbers.".into());
                    }
                }
            }
            _ => {}
        }
    }
    // JSON maps cannot use CellRange keys. Serialize the deterministic ordered
    // entries as pairs, matching our native persistence representation.
    let pairs: Vec<_> = store.rules.iter().collect();
    let bytes = serde_json::to_vec(&(pairs, &store.exclusions)).map_err(|e| e.to_string())?;
    Ok(Sha256::digest(bytes).into())
}

struct Budget(usize);
impl Budget {
    fn step(&mut self) -> Result<(), String> {
        self.0 = self.0.checked_sub(1).ok_or_else(limit_error)?;
        Ok(())
    }
}
fn limit_error() -> String {
    "This edit would split too many validation ranges. Select a smaller range.".into()
}
fn disjoint(
    ranges: impl Iterator<Item = CellRange>,
    budget: &mut Budget,
) -> Result<Vec<CellRange>, String> {
    let mut result = Vec::new();
    for range in ranges {
        let pieces = subtract_all(vec![range], &result, budget)?;
        result.extend(pieces);
        if result.len() > MAX_FRAGMENTS {
            return Err(limit_error());
        }
    }
    Ok(result)
}
fn subtract_all(
    mut pieces: Vec<CellRange>,
    cuts: &[CellRange],
    budget: &mut Budget,
) -> Result<Vec<CellRange>, String> {
    for &cut in cuts {
        let mut next = Vec::new();
        for range in pieces {
            budget.step()?;
            subtract(range, cut, &mut next);
            if next.len() > MAX_FRAGMENTS {
                return Err(limit_error());
            }
        }
        pieces = next;
        if pieces.is_empty() {
            break;
        }
    }
    Ok(pieces)
}
fn subtract(range: CellRange, cut: CellRange, result: &mut Vec<CellRange>) {
    if !range.overlaps(&cut) {
        result.push(range);
        return;
    }
    let top = range.start_row.max(cut.start_row);
    let bottom = range.end_row.min(cut.end_row);
    let left = range.start_col.max(cut.start_col);
    let right = range.end_col.min(cut.end_col);
    if range.start_row < top {
        result.push(CellRange::new(
            range.start_row,
            range.start_col,
            top - 1,
            range.end_col,
        ));
    }
    if bottom < range.end_row {
        result.push(CellRange::new(
            bottom + 1,
            range.start_col,
            range.end_row,
            range.end_col,
        ));
    }
    if range.start_col < left {
        result.push(CellRange::new(top, range.start_col, bottom, left - 1));
    }
    if right < range.end_col {
        result.push(CellRange::new(top, right + 1, bottom, range.end_col));
    }
}

#[cfg(test)]
mod tests;
