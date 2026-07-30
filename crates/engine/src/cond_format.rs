//! Conditional formatting rules
//!
//! A rule is a formula plus a style: for each cell in the rule's target
//! ranges, the predicate formula is evaluated with Excel-style relative
//! anchoring, and matching cells get the rule's style merged over their
//! base format.
//!
//! Spec: planning/visigrid/features/conditional-formatting-spec.md
//!
//! ## Anchoring
//!
//! The predicate is written relative to the target range's top-left cell.
//! For each cell, relative reference components shift by the cell's offset
//! from that anchor; absolute (`$`) components do not. `A2:D500` with
//! `=$C2>100` tests C2 on row 2, C3 on row 3, in every column.
//!
//! ## Precedence
//!
//! Rules are an ordered list. All matching rules apply; later rules win
//! per property (a later `bg` beats an earlier `bg`; unrelated properties
//! from earlier rules survive).
//!
//! ## Errors
//!
//! A predicate that fails to parse leaves the rule inert (`parse_error()`
//! is surfaced in the rules panel). A predicate that evaluates to an error
//! for a given cell is a no-match for that cell — never a grid error.

use serde::{Deserialize, Serialize};

use crate::cell::{CellFormat, CellFormatOverride, CellStyle};
use crate::formula::eval::{evaluate, CellLookup};
use crate::formula::parser::{bind_expr_same_sheet, parse, Expr, ParsedExpr};
use crate::validation::CellRange;

// ============================================================================
// Style
// ============================================================================

/// The style a rule applies to matching cells.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CondStyle {
    /// One of the built-in semantic cell styles (error, warning, success, …).
    Named(CellStyle),
    /// Explicit property overrides (bg, fg, bold, …).
    Inline(CellFormatOverride),
    /// Format-by-example: the template cell's explicit format properties,
    /// snapshotted when the rule was created. `source` is kept so the rules
    /// panel can display `like(Z1)`.
    Like {
        source: (usize, usize),
        snapshot: CellFormatOverride,
    },
}

impl CondStyle {
    /// The format override this style contributes to a matching cell.
    pub fn as_override(&self) -> CellFormatOverride {
        match self {
            CondStyle::Named(style) => CellFormatOverride {
                cell_style: Some(*style),
                ..Default::default()
            },
            CondStyle::Inline(ov) => ov.clone(),
            CondStyle::Like { snapshot, .. } => snapshot.clone(),
        }
    }
}

// ============================================================================
// Rule
// ============================================================================

/// A single conditional formatting rule.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CondFormatRule {
    /// Stable across edits and reorders; used by undo and the inspector.
    pub id: u64,
    /// Target ranges. A cell in any of them is subject to the rule.
    pub ranges: Vec<CellRange>,
    /// Predicate formula source, anchored at each range's top-left cell.
    pub predicate: String,
    /// Style applied to matching cells.
    pub style: CondStyle,
    /// Disabled rules are kept but never match.
    pub enabled: bool,
    /// Parsed predicate, built once at construction/edit time.
    /// None = predicate failed to parse (rule is inert).
    #[serde(skip)]
    ast: Option<ParsedExpr>,
    /// Parse error message when `ast` is None, for the rules panel.
    #[serde(skip)]
    parse_error: Option<String>,
}

impl PartialEq for CondFormatRule {
    fn eq(&self, other: &Self) -> bool {
        self.id == other.id
            && self.ranges == other.ranges
            && self.predicate == other.predicate
            && self.style == other.style
            && self.enabled == other.enabled
    }
}

impl CondFormatRule {
    pub fn new(id: u64, ranges: Vec<CellRange>, predicate: impl Into<String>, style: CondStyle) -> Self {
        let predicate = predicate.into();
        let (ast, parse_error) = match parse(&predicate) {
            Ok(ast) => (Some(ast), None),
            Err(e) => (None, Some(e)),
        };
        Self { id, ranges, predicate, style, enabled: true, ast, parse_error }
    }

    /// Re-parse the predicate. Called after deserialization (the AST is not
    /// serialized) and after predicate edits.
    pub fn reparse(&mut self) {
        match parse(&self.predicate) {
            Ok(ast) => {
                self.ast = Some(ast);
                self.parse_error = None;
            }
            Err(e) => {
                self.ast = None;
                self.parse_error = Some(e);
            }
        }
    }

    /// Parse error for the rules panel, if the predicate is invalid.
    pub fn parse_error(&self) -> Option<&str> {
        self.parse_error.as_deref()
    }

    /// Does this rule target the given cell?
    pub fn covers(&self, row: usize, col: usize) -> bool {
        self.ranges.iter().any(|r| r.contains(row, col))
    }

    /// The anchor (top-left of the first range containing the cell).
    fn anchor_for(&self, row: usize, col: usize) -> Option<(usize, usize)> {
        self.ranges
            .iter()
            .find(|r| r.contains(row, col))
            .map(|r| (r.start_row, r.start_col))
    }

    /// The predicate as it evaluates for a specific cell — relative refs
    /// shifted to that cell's offset from the range anchor. This is the
    /// inspector's "why is this cell styled" answer: `=A1>100` on A1:A10
    /// shows as `=A7>100` when inspecting row 7.
    pub fn predicate_at(&self, row: usize, col: usize) -> Option<String> {
        let ast = self.ast.as_ref()?;
        let (anchor_row, anchor_col) = self.anchor_for(row, col)?;
        let dr = row as i64 - anchor_row as i64;
        let dc = col as i64 - anchor_col as i64;
        let shifted = offset_expr(ast, dr, dc)?;
        let bound = bind_expr_same_sheet(&shifted);
        // format_expr includes the leading '='
        Some(crate::formula::parser::format_expr(&bound, |name| Some(name.to_string())))
    }

    /// Evaluate the predicate for a cell. True = the rule's style applies.
    ///
    /// No-match on: disabled rule, cell outside all ranges, unparseable
    /// predicate, reference shifted out of bounds, or an error result.
    pub fn matches<L: CellLookup>(&self, row: usize, col: usize, lookup: &L) -> bool {
        if !self.enabled {
            return false;
        }
        let Some(ast) = &self.ast else { return false };
        let Some((anchor_row, anchor_col)) = self.anchor_for(row, col) else {
            return false;
        };

        let dr = row as i64 - anchor_row as i64;
        let dc = col as i64 - anchor_col as i64;
        let Some(shifted) = offset_expr(ast, dr, dc) else {
            // Relative reference shifted out of bounds — Excel shows #REF!;
            // for formatting purposes that's simply a no-match.
            return false;
        };

        let bound = bind_expr_same_sheet(&shifted);
        evaluate(&bound, lookup).to_bool().unwrap_or(false)
    }
}

/// Shift the relative components of every reference in `expr` by (dr, dc).
/// Absolute components are unchanged. Returns None if any reference would
/// move out of bounds (negative).
fn offset_expr(expr: &ParsedExpr, dr: i64, dc: i64) -> Option<ParsedExpr> {
    fn shift(v: usize, abs: bool, d: i64) -> Option<usize> {
        if abs || d == 0 {
            return Some(v);
        }
        let shifted = v as i64 + d;
        usize::try_from(shifted).ok()
    }

    Some(match expr {
        Expr::Number(n) => Expr::Number(*n),
        Expr::Text(s) => Expr::Text(s.clone()),
        Expr::Boolean(b) => Expr::Boolean(*b),
        Expr::Empty => Expr::Empty,
        Expr::RefError => Expr::RefError,
        Expr::NamedRange(name) => Expr::NamedRange(name.clone()),
        Expr::CellRef { sheet, col, row, col_abs, row_abs } => Expr::CellRef {
            sheet: sheet.clone(),
            col: shift(*col, *col_abs, dc)?,
            row: shift(*row, *row_abs, dr)?,
            col_abs: *col_abs,
            row_abs: *row_abs,
        },
        Expr::Range {
            sheet,
            start_col, start_row, end_col, end_row,
            start_col_abs, start_row_abs, end_col_abs, end_row_abs,
        } => Expr::Range {
            sheet: sheet.clone(),
            start_col: shift(*start_col, *start_col_abs, dc)?,
            start_row: shift(*start_row, *start_row_abs, dr)?,
            end_col: shift(*end_col, *end_col_abs, dc)?,
            end_row: shift(*end_row, *end_row_abs, dr)?,
            start_col_abs: *start_col_abs,
            start_row_abs: *start_row_abs,
            end_col_abs: *end_col_abs,
            end_row_abs: *end_row_abs,
        },
        Expr::Function { name, args } => Expr::Function {
            name: name.clone(),
            args: args
                .iter()
                .map(|a| offset_expr(a, dr, dc))
                .collect::<Option<Vec<_>>>()?,
        },
        Expr::BinaryOp { op, left, right } => Expr::BinaryOp {
            op: *op,
            left: Box::new(offset_expr(left, dr, dc)?),
            right: Box::new(offset_expr(right, dr, dc)?),
        },
    })
}

// ============================================================================
// Store
// ============================================================================

/// Ordered conditional formatting rules for one sheet.
///
/// Position in the list is precedence: all matching rules apply, later
/// rules win per property.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CondFormatStore {
    rules: Vec<CondFormatRule>,
    next_id: u64,
}

impl CondFormatStore {
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a rule at the end (highest precedence). Returns its id.
    pub fn add(&mut self, ranges: Vec<CellRange>, predicate: impl Into<String>, style: CondStyle) -> u64 {
        let id = self.next_id;
        self.next_id += 1;
        self.rules.push(CondFormatRule::new(id, ranges, predicate, style));
        id
    }

    /// Remove a rule by id. Returns it if it existed.
    pub fn remove(&mut self, id: u64) -> Option<CondFormatRule> {
        let idx = self.rules.iter().position(|r| r.id == id)?;
        Some(self.rules.remove(idx))
    }

    /// Re-insert a previously removed rule at a position (for undo).
    pub fn insert_at(&mut self, index: usize, rule: CondFormatRule) {
        let index = index.min(self.rules.len());
        self.next_id = self.next_id.max(rule.id + 1);
        self.rules.insert(index, rule);
    }

    /// Move the rule with `id` to `new_index` (clamped). Returns false if absent.
    pub fn reorder(&mut self, id: u64, new_index: usize) -> bool {
        let Some(idx) = self.rules.iter().position(|r| r.id == id) else {
            return false;
        };
        let rule = self.rules.remove(idx);
        let new_index = new_index.min(self.rules.len());
        self.rules.insert(new_index, rule);
        true
    }

    pub fn get(&self, id: u64) -> Option<&CondFormatRule> {
        self.rules.iter().find(|r| r.id == id)
    }

    pub fn get_mut(&mut self, id: u64) -> Option<&mut CondFormatRule> {
        self.rules.iter_mut().find(|r| r.id == id)
    }

    pub fn iter(&self) -> impl Iterator<Item = &CondFormatRule> {
        self.rules.iter()
    }

    pub fn len(&self) -> usize {
        self.rules.len()
    }

    pub fn is_empty(&self) -> bool {
        self.rules.is_empty()
    }

    /// Does any enabled rule target this cell? Cheap pre-check so render
    /// paths can skip predicate evaluation entirely for uncovered cells.
    pub fn any_rule_covers(&self, row: usize, col: usize) -> bool {
        self.rules.iter().any(|r| r.enabled && r.covers(row, col))
    }

    /// Merged override for a cell from all matching rules, later rules
    /// winning per property. None if no rule matches.
    pub fn override_for_cell<L: CellLookup>(
        &self,
        row: usize,
        col: usize,
        lookup: &L,
    ) -> Option<CellFormatOverride> {
        let mut merged: Option<CellFormatOverride> = None;
        for rule in &self.rules {
            if rule.matches(row, col, lookup) {
                let ov = rule.style.as_override();
                merged = Some(match merged {
                    None => ov,
                    Some(mut acc) => {
                        acc.merge_from(&ov);
                        acc
                    }
                });
            }
        }
        merged
    }

    /// Effective format for a cell: base format plus all matching rules.
    pub fn effective_format<L: CellLookup>(
        &self,
        row: usize,
        col: usize,
        base: &CellFormat,
        lookup: &L,
    ) -> CellFormat {
        match self.override_for_cell(row, col, lookup) {
            Some(ov) => {
                let mut format = base.clone();
                ov.apply_to(&mut format);
                format
            }
            None => base.clone(),
        }
    }

    /// Re-parse all predicates. Call after deserialization.
    pub fn reparse_all(&mut self) {
        for rule in &mut self.rules {
            rule.reparse();
        }
    }

    /// Rules whose predicates matched a cell — for the inspector's
    /// "why is this cell yellow?" answer.
    pub fn matching_rules<L: CellLookup>(&self, row: usize, col: usize, lookup: &L) -> Vec<u64> {
        self.rules
            .iter()
            .filter(|r| r.matches(row, col, lookup))
            .map(|r| r.id)
            .collect()
    }

    // ------------------------------------------------------------------
    // Row/column structure changes (grid-line semantics, matching merges)
    // ------------------------------------------------------------------
    //
    // Only target ranges shift; predicate text is not rewritten. A rule
    // whose ranges are all deleted keeps an empty range list and simply
    // stops matching (still visible in the rules panel).

    pub fn insert_rows(&mut self, at_row: usize, count: usize) {
        self.shift_insert(at_row, count, true);
    }

    pub fn insert_cols(&mut self, at_col: usize, count: usize) {
        self.shift_insert(at_col, count, false);
    }

    pub fn delete_rows(&mut self, start_row: usize, count: usize) {
        self.shift_delete(start_row, count, true);
    }

    pub fn delete_cols(&mut self, start_col: usize, count: usize) {
        self.shift_delete(start_col, count, false);
    }

    fn shift_insert(&mut self, at: usize, count: usize, rows: bool) {
        for rule in &mut self.rules {
            for r in &mut rule.ranges {
                let (start, end) = if rows {
                    (&mut r.start_row, &mut r.end_row)
                } else {
                    (&mut r.start_col, &mut r.end_col)
                };
                if at <= *start {
                    // Insertion at or before range → shift whole range
                    *start += count;
                    *end += count;
                } else if at <= *end {
                    // Insertion inside range → expand
                    *end += count;
                }
            }
        }
    }

    fn shift_delete(&mut self, del_start: usize, count: usize, rows: bool) {
        let del_end = del_start + count; // exclusive
        for rule in &mut self.rules {
            rule.ranges.retain_mut(|r| {
                let (start, end) = if rows {
                    (&mut r.start_row, &mut r.end_row)
                } else {
                    (&mut r.start_col, &mut r.end_col)
                };
                if del_end <= *start {
                    // Deletion entirely before → shift toward origin
                    *start -= count;
                    *end -= count;
                    true
                } else if del_start > *end {
                    // Deletion entirely after → no effect
                    true
                } else if del_start <= *start && del_end > *end {
                    // Deletion engulfs range → drop it
                    false
                } else if del_start <= *start {
                    // Deletion clips leading edge
                    *start = del_start;
                    *end -= count;
                    true
                } else if del_end > *end {
                    // Deletion clips trailing edge
                    *end = del_start - 1;
                    true
                } else {
                    // Deletion entirely inside → shrink
                    *end -= count;
                    true
                }
            });
        }
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{Sheet, SheetId};

    fn range(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> CellRange {
        CellRange { start_row, start_col, end_row, end_col }
    }

    fn sheet_with_column(values: &[&str]) -> Sheet {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        for (i, v) in values.iter().enumerate() {
            sheet.set_value(i, 0, v);
        }
        sheet
    }

    #[test]
    fn relative_anchoring_shifts_per_row() {
        // A1=5, A2=15, A3=25; rule on A1:A3: =A1>10
        let sheet = sheet_with_column(&["5", "15", "25"]);
        let mut store = CondFormatStore::new();
        store.add(
            vec![range(0, 0, 2, 0)],
            "=A1>10",
            CondStyle::Named(CellStyle::Warning),
        );

        assert!(!store.any_rule_covers(5, 5));
        assert!(store.override_for_cell(0, 0, &sheet).is_none(), "A1=5 no match");
        assert!(store.override_for_cell(1, 0, &sheet).is_some(), "A2=15 matches");
        assert!(store.override_for_cell(2, 0, &sheet).is_some(), "A3=25 matches");
    }

    #[test]
    fn absolute_refs_do_not_shift() {
        // Threshold in C1 ($C$1=10). A1=5, A2=15.
        let mut sheet = sheet_with_column(&["5", "15"]);
        sheet.set_value(0, 2, "10");
        let mut store = CondFormatStore::new();
        store.add(
            vec![range(0, 0, 1, 0)],
            "=A1>$C$1",
            CondStyle::Named(CellStyle::Error),
        );

        assert!(store.override_for_cell(0, 0, &sheet).is_none());
        assert!(store.override_for_cell(1, 0, &sheet).is_some());
    }

    #[test]
    fn column_anchored_predicate_applies_across_columns() {
        // Rule on A1:C2 keyed on column A: =$A1>10. Whole row highlights.
        let mut sheet = sheet_with_column(&["5", "15"]);
        sheet.set_value(0, 1, "x");
        sheet.set_value(1, 1, "y");
        let mut store = CondFormatStore::new();
        store.add(
            vec![range(0, 0, 1, 2)],
            "=$A1>10",
            CondStyle::Named(CellStyle::Success),
        );

        assert!(store.override_for_cell(0, 1, &sheet).is_none(), "row 1 no match");
        assert!(store.override_for_cell(1, 1, &sheet).is_some(), "row 2 col B matches");
        assert!(store.override_for_cell(1, 2, &sheet).is_some(), "row 2 col C matches");
    }

    #[test]
    fn later_rule_wins_per_property() {
        let sheet = sheet_with_column(&["15"]);
        let mut store = CondFormatStore::new();
        // Rule 1: bold + red bg. Rule 2: blue bg only.
        store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Inline(CellFormatOverride {
                bold: Some(true),
                background_color: Some(Some([255, 0, 0, 255])),
                ..Default::default()
            }),
        );
        store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Inline(CellFormatOverride {
                background_color: Some(Some([0, 0, 255, 255])),
                ..Default::default()
            }),
        );

        let ov = store.override_for_cell(0, 0, &sheet).unwrap();
        assert_eq!(ov.bold, Some(true), "bold from rule 1 survives");
        assert_eq!(
            ov.background_color,
            Some(Some([0, 0, 255, 255])),
            "bg from rule 2 wins"
        );

        let base = CellFormat::default();
        let effective = store.effective_format(0, 0, &base, &sheet);
        assert!(effective.bold);
        assert_eq!(effective.background_color, Some([0, 0, 255, 255]));
    }

    #[test]
    fn predicate_error_is_no_match() {
        let sheet = sheet_with_column(&["abc"]);
        let mut store = CondFormatStore::new();
        // 1/0 → error → no match, not a grid error
        store.add(
            vec![range(0, 0, 0, 0)],
            "=1/0",
            CondStyle::Named(CellStyle::Error),
        );
        assert!(store.override_for_cell(0, 0, &sheet).is_none());
    }

    #[test]
    fn unparseable_predicate_is_inert_with_error() {
        let sheet = sheet_with_column(&["1"]);
        let mut store = CondFormatStore::new();
        let id = store.add(
            vec![range(0, 0, 0, 0)],
            "=SUM((",
            CondStyle::Named(CellStyle::Error),
        );
        assert!(store.get(id).unwrap().parse_error().is_some());
        assert!(store.override_for_cell(0, 0, &sheet).is_none());
    }

    #[test]
    fn out_of_bounds_shift_is_no_match() {
        // Rule anchored so a relative ref would go negative for the first cell:
        // range B1:B2 with predicate =A0-style is impossible to write, so use
        // a ref one row above the anchor: predicate =A1 evaluated for... trick:
        // anchor at row 1 (B2:B3), predicate references row above anchor via
        // relative ref written as =A1 — for the anchor cell dr=0 (fine), and
        // we instead check dc: range at col 0 with predicate referencing a
        // column left of A is unwritable. So simulate via disabled? Instead:
        // predicate =A1 with range starting at (0,0), evaluated at (0,0) is
        // fine; out-of-bounds requires offset_expr directly.
        let ast = parse("=A1").unwrap();
        assert!(offset_expr(&ast, -1, 0).is_none(), "negative row shift refuses");
        assert!(offset_expr(&ast, 0, -1).is_none(), "negative col shift refuses");
        assert!(offset_expr(&ast, 1, 1).is_some());
    }

    #[test]
    fn disabled_rule_never_matches() {
        let sheet = sheet_with_column(&["15"]);
        let mut store = CondFormatStore::new();
        let id = store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Named(CellStyle::Warning),
        );
        assert!(store.override_for_cell(0, 0, &sheet).is_some());
        store.get_mut(id).unwrap().enabled = false;
        assert!(store.override_for_cell(0, 0, &sheet).is_none());
    }

    #[test]
    fn reorder_changes_precedence() {
        let sheet = sheet_with_column(&["15"]);
        let mut store = CondFormatStore::new();
        let red = store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Inline(CellFormatOverride {
                background_color: Some(Some([255, 0, 0, 255])),
                ..Default::default()
            }),
        );
        store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Inline(CellFormatOverride {
                background_color: Some(Some([0, 0, 255, 255])),
                ..Default::default()
            }),
        );
        // Blue is later → wins. Move red to the end → red wins.
        assert_eq!(
            store.override_for_cell(0, 0, &sheet).unwrap().background_color,
            Some(Some([0, 0, 255, 255]))
        );
        store.reorder(red, 1);
        assert_eq!(
            store.override_for_cell(0, 0, &sheet).unwrap().background_color,
            Some(Some([255, 0, 0, 255]))
        );
    }

    #[test]
    fn serde_roundtrip_reparses() {
        let sheet = sheet_with_column(&["15"]);
        let mut store = CondFormatStore::new();
        store.add(
            vec![range(0, 0, 0, 0)],
            "=A1>10",
            CondStyle::Named(CellStyle::Warning),
        );

        let json = serde_json::to_string(&store).unwrap();
        let mut restored: CondFormatStore = serde_json::from_str(&json).unwrap();
        // AST is not serialized; rules are inert until reparse_all.
        assert!(restored.override_for_cell(0, 0, &sheet).is_none());
        restored.reparse_all();
        assert!(restored.override_for_cell(0, 0, &sheet).is_some());
    }

    #[test]
    fn insert_rows_shifts_and_expands_ranges() {
        let mut store = CondFormatStore::new();
        let id = store.add(vec![range(5, 0, 10, 3)], "=A1", CondStyle::Named(CellStyle::Note));

        // Insert above → whole range shifts down
        store.insert_rows(2, 3);
        assert_eq!(store.get(id).unwrap().ranges[0], range(8, 0, 13, 3));

        // Insert inside → range expands
        store.insert_rows(10, 2);
        assert_eq!(store.get(id).unwrap().ranges[0], range(8, 0, 15, 3));

        // Insert below → no effect
        store.insert_rows(50, 5);
        assert_eq!(store.get(id).unwrap().ranges[0], range(8, 0, 15, 3));
    }

    #[test]
    fn delete_rows_clips_shrinks_and_drops_ranges() {
        let mut store = CondFormatStore::new();
        let id = store.add(vec![range(5, 0, 10, 3)], "=A1", CondStyle::Named(CellStyle::Note));

        // Delete entirely above → shift up
        store.delete_rows(0, 2);
        assert_eq!(store.get(id).unwrap().ranges[0], range(3, 0, 8, 3));

        // Delete clipping the top (rows 2-4 → deletes range rows 3-4)
        store.delete_rows(2, 3);
        assert_eq!(store.get(id).unwrap().ranges[0], range(2, 0, 5, 3));

        // Delete inside → shrink
        store.delete_rows(3, 1);
        assert_eq!(store.get(id).unwrap().ranges[0], range(2, 0, 4, 3));

        // Delete clipping the bottom
        store.delete_rows(4, 10);
        assert_eq!(store.get(id).unwrap().ranges[0], range(2, 0, 3, 3));

        // Delete engulfing → range dropped, rule kept
        store.delete_rows(0, 20);
        assert!(store.get(id).unwrap().ranges.is_empty());
        assert!(!store.get(id).unwrap().covers(0, 0));
    }

    #[test]
    fn delete_cols_shifts_ranges() {
        let mut store = CondFormatStore::new();
        let id = store.add(vec![range(0, 5, 3, 8)], "=A1", CondStyle::Named(CellStyle::Note));
        store.delete_cols(0, 2);
        assert_eq!(store.get(id).unwrap().ranges[0], range(0, 3, 3, 6));
        store.insert_cols(4, 1);
        assert_eq!(store.get(id).unwrap().ranges[0], range(0, 3, 3, 7));
    }

    #[test]
    fn text_predicates_and_functions_work() {
        let mut sheet = Sheet::new(SheetId(1), 100, 100);
        sheet.set_value(0, 0, "urgent: fix the build");
        sheet.set_value(1, 0, "later");
        let mut store = CondFormatStore::new();
        store.add(
            vec![range(0, 0, 1, 0)],
            "=ISNUMBER(FIND(\"urgent\", A1))",
            CondStyle::Named(CellStyle::Error),
        );
        assert!(store.override_for_cell(0, 0, &sheet).is_some());
        assert!(store.override_for_cell(1, 0, &sheet).is_none());
    }
}
