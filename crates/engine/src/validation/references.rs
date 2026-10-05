use super::{ConstraintValue, ListSource, ValidationRule, ValidationType};
use std::borrow::Cow;

impl ValidationRule {
    /// Formula-backed imports need a source position even when every A1
    /// component is absolute: those addresses still follow structural edits.
    pub fn has_reference_sources(&self) -> bool {
        let mut found = false;
        let mut probe = self.clone();
        probe.map_references(|source| {
            found = true;
            source.to_owned()
        });
        found
    }

    pub(crate) fn adjust_relative_sources_for_structural(
        &mut self,
        edit: &crate::structural::StructuralEdit,
        sheet_name: &str,
    ) {
        if self.reference_origin.is_none() {
            return;
        }
        self.map_references(|source| {
            let has_equals = source.starts_with('=');
            let formula = if has_equals {
                source.to_owned()
            } else {
                format!("={source}")
            };
            match crate::structural::adjust_formula_text(&formula, edit, sheet_name) {
                Some(adjusted) if !has_equals => {
                    adjusted.strip_prefix('=').unwrap_or(&adjusted).to_owned()
                }
                Some(adjusted) => adjusted,
                None => source.to_owned(),
            }
        });
    }

    pub fn has_relative_references(&self) -> bool {
        let mut relative = false;
        let mut probe = self.clone();
        probe.map_references(|s| {
            relative |= crate::formula::parser::absolutize_formula_refs(s)
                != crate::formula::parser::adjust_formula_refs(s, 0, 0);
            s.to_owned()
        });
        relative
    }

    /// Resolve source text at a target without mutating the stored rule. Copies
    /// keep an explicit origin, so splitting ranges cannot change their meaning.
    pub fn at(&self, row: usize, col: usize) -> Cow<'_, Self> {
        let Some((r, c)) = self.reference_origin else {
            return Cow::Borrowed(self);
        };
        if (r, c) == (row, col) {
            return Cow::Borrowed(self);
        }
        let mut rule = self.clone();
        rule.map_references(|source| {
            match (
                i64::try_from(row)
                    .ok()
                    .zip(i64::try_from(r).ok())
                    .and_then(|(a, b)| i32::try_from(a - b).ok()),
                i64::try_from(col)
                    .ok()
                    .zip(i64::try_from(c).ok())
                    .and_then(|(a, b)| i32::try_from(a - b).ok()),
            ) {
                (Some(dr), Some(dc)) => crate::formula::parser::adjust_formula_refs(source, dr, dc),
                _ => "#REF!".into(),
            }
        });
        rule.reference_origin = Some((row, col));
        Cow::Owned(rule)
    }

    /// Produce a rule expressed at one XLSX range's first cell. Native rules
    /// keep fixed addresses; imported rules keep relative and mixed references.
    pub fn for_xlsx_range(&self, row: usize, col: usize) -> Self {
        if self.reference_origin.is_some() {
            return self.at(row, col).into_owned();
        }
        let mut rule = self.clone();
        rule.map_references(crate::formula::parser::absolutize_formula_refs);
        rule
    }

    fn map_references(&mut self, mut map: impl FnMut(&str) -> String) {
        match &mut self.rule_type {
            ValidationType::Custom(s) | ValidationType::List(ListSource::Range(s)) => *s = map(s),
            ValidationType::List(_) => {}
            ValidationType::WholeNumber(c)
            | ValidationType::Decimal(c)
            | ValidationType::Date(c)
            | ValidationType::Time(c)
            | ValidationType::TextLength(c) => {
                for v in std::iter::once(&mut c.value1).chain(c.value2.iter_mut()) {
                    if let ConstraintValue::CellRef(s) | ConstraintValue::Formula(s) = v {
                        *s = map(s);
                    }
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        validation::{CellRange, NumericConstraint, ValidationEdit, ValidationStore},
        workbook::Workbook,
    };

    #[test]
    fn imported_sources_use_strict_addresses_and_quoted_sheet_names() {
        let mut wb = Workbook::new();
        let source = wb.add_sheet_named("Limits! O'Brien").unwrap();
        wb.sheet_mut(source).unwrap().set_value(1, 0, "10");
        wb.sheet_mut(source).unwrap().set_value(2, 0, "20");
        let mut rule = ValidationRule::whole_number(NumericConstraint::greater_than(
            ConstraintValue::CellRef("='Limits! O''Brien'!$A2".into()),
        ));
        rule.reference_origin = Some((1, 1));
        wb.sheet_mut(0)
            .unwrap()
            .validations
            .set(CellRange::new(1, 1, 2, 1), rule);
        assert!(wb.validate_cell_input(0, 1, 1, "15").is_valid());
        assert!(wb.validate_cell_input(0, 2, 1, "15").is_invalid());
        let mut rule = ValidationRule::list_range("='Limits! O''Brien'!$A2:$A2");
        rule.reference_origin = Some((1, 2));
        wb.sheet_mut(0)
            .unwrap()
            .validations
            .set(CellRange::new(1, 2, 2, 2), rule);
        assert_eq!(wb.get_list_items(0, 2, 2).unwrap().items, ["20"]);
        let sheet = wb.sheet(0).unwrap();
        for invalid in ["A0", "A1048577", "XFE1", "A1+B2", "A1B2", "中1", "'X'!A1"] {
            assert!(sheet.parse_cell_ref(invalid).is_none(), "{invalid}");
        }
        assert!(sheet
            .parse_cell_ref(&format!("{}1", "Z".repeat(200)))
            .is_none());
        assert_eq!(
            sheet.parse_cell_ref("=$XFD$1048576"),
            Some((1048575, 16383))
        );
    }

    #[test]
    fn huge_sparse_list_ranges_are_resolved_in_row_order_with_bounded_items() {
        let mut wb = Workbook::new();
        let sheet = wb.active_sheet_mut();
        sheet.set_value(1048575, 0, "Last");
        sheet.set_value(0, 0, "First");
        assert_eq!(sheet.resolve_range_to_list("A:A").items, ["First", "Last"]);
        assert_eq!(
            sheet.resolve_range_to_list("A1:XFD1048576").items,
            ["First", "Last"]
        );
        for i in 0..super::super::MAX_LIST_ITEMS + 2 {
            sheet.set_value(i, 1, &format!("{i}"));
        }
        let resolved = sheet.resolve_range_to_list("B:B");
        assert!(resolved.is_truncated);
        assert_eq!(resolved.items.len(), super::super::MAX_LIST_ITEMS);
        assert_eq!(resolved.items[0], "0");
        assert_eq!(resolved.items.last().unwrap(), "9999");
    }

    #[test]
    fn relative_mixed_and_fixed_sources_retain_their_meaning() {
        let text = "=IF(A1>0, SUM($B2,C$3,$D$4,'Other Sheet'!E5,A:A,$2:4),\"A1\")+Table1[A1]";
        let fixed = ValidationRule::custom(text);
        assert!(fixed.has_relative_references());
        assert_eq!(fixed.at(10, 10).rule_type, fixed.rule_type);
        assert_eq!(
            fixed.for_xlsx_range(10, 10).rule_type,
            ValidationType::Custom(
                "=IF($A$1>0, SUM($B$2,$C$3,$D$4,'Other Sheet'!$E$5,$A:$A,$2:$4),\"A1\")+Table1[A1]"
                    .into()
            )
        );
        let mut relative = fixed.clone();
        relative.reference_origin = Some((1, 1));
        assert_eq!(
            relative.at(3, 4).rule_type,
            ValidationType::Custom(
                "=IF(D3>0, SUM($B4,F$3,$D$4,'Other Sheet'!H7,D:D,$2:6),\"A1\")+Table1[A1]".into()
            )
        );
        assert_eq!(relative.at(3, 4).reference_origin, Some((3, 4)));
        assert_eq!(relative.rule_type, fixed.rule_type);
        assert!(
            !ValidationRule::custom("=$A$1+SUM($B:$B)+\"C2\"+Table1[A1]").has_relative_references()
        );
        assert_eq!(
            crate::formula::parser::absolutize_formula_refs("=日本A1+\\B2+C3+Table1[A1]"),
            "=日本A1+\\B2+$C$3+Table1[A1]"
        );
    }

    #[test]
    fn numeric_checks_and_dropdowns_resolve_each_target_cell() {
        let mut wb = Workbook::new();
        let sheet = wb.active_sheet_mut();
        sheet.set_value(1, 0, "5");
        sheet.set_value(2, 0, "10");
        let mut rule = ValidationRule::whole_number(NumericConstraint::greater_than(
            ConstraintValue::CellRef("$A2".into()),
        ));
        rule.reference_origin = Some((1, 1));
        sheet.validations.set(CellRange::new(1, 1, 2, 1), rule);
        let mut list = ValidationRule::list_range("=$A2:$A2");
        list.reference_origin = Some((1, 2));
        sheet.validations.set(CellRange::new(1, 2, 2, 2), list);
        assert!(sheet.validate_cell_input(1, 1, "7").is_valid());
        assert!(sheet.validate_cell_input(2, 1, "7").is_invalid());
        assert_eq!(sheet.get_list_items(2, 2).unwrap().items, ["10"]);
        assert!(sheet.validate_cell_input(2, 2, "5").is_invalid());
        assert!(wb.validate_cell_input(0, 1, 1, "7").is_valid());
        assert!(wb.validate_cell_input(0, 2, 1, "7").is_invalid());
        assert_eq!(wb.get_list_items(0, 2, 2).unwrap().items, ["10"]);
        let limits = wb.add_sheet_named("Limits").unwrap();
        wb.sheet_mut(limits).unwrap().set_value(1, 0, "Yes");
        wb.sheet_mut(limits).unwrap().set_value(2, 0, "No");
        let mut cross_sheet = ValidationRule::list_range("=Limits!$A2:$A2");
        cross_sheet.reference_origin = Some((1, 3));
        wb.sheet_mut(0)
            .unwrap()
            .validations
            .set(CellRange::new(1, 3, 2, 3), cross_sheet);
        assert!(wb.validate_cell_input(0, 1, 3, "Yes").is_valid());
        assert!(wb.validate_cell_input(0, 2, 3, "Yes").is_invalid());
        assert!(wb.validate_cell_input(0, 2, 3, "No").is_valid());
        let sheet = wb.sheet_mut(0).unwrap();
        sheet.validations.set(
            CellRange::single(0, 4),
            ValidationRule::new(ValidationType::TextLength(NumericConstraint::equal_to(2))),
        );
        assert!(sheet.validate_cell_input(0, 4, "日本").is_valid());
        assert!(wb.validate_cell_input(0, 0, 4, "日本").is_valid());
    }

    #[test]
    fn splitting_excluding_and_replaying_keep_the_original_origin() {
        let mut store = ValidationStore::new();
        let mut rule = ValidationRule::custom("=A2>0");
        rule.reference_origin = Some((1, 1));
        store.set(CellRange::new(1, 1, 9, 1), rule.clone());
        store.exclude(CellRange::single(1, 1));
        store.exclude(CellRange::single(5, 1));
        let original = store.clone();
        let patch = store
            .plan_edit(&[CellRange::single(3, 1)], ValidationEdit::Clear, 100, 10)
            .unwrap();
        patch.apply(&mut store, true).unwrap();
        for (range, rule) in store.effective_ranges().unwrap() {
            assert_eq!(rule.reference_origin, Some((1, 1)));
            assert_eq!(
                rule.for_xlsx_range(range.start_row, range.start_col)
                    .rule_type,
                ValidationType::Custom(format!("=A{}>0", range.start_row + 1))
            );
        }
        assert_eq!(
            store.get(8, 1).unwrap().at(8, 1).rule_type,
            ValidationType::Custom("=A9>0".into())
        );
        patch.apply(&mut store, false).unwrap();
        assert_eq!(store, original);
        patch.apply(&mut store, true).unwrap();
        assert!(store.get(3, 1).is_none());
    }

    #[test]
    fn effective_ranges_match_exclusion_and_overlap_precedence() {
        let mut store = ValidationStore::new();
        for (range, text) in [
            (CellRange::new(0, 0, 4, 4), "first"),
            (CellRange::new(1, 1, 5, 5), "second"),
            (CellRange::new(3, 0, 5, 3), "third"),
        ] {
            store.set(range, ValidationRule::list_inline(vec![text.into()]));
        }
        store.exclude(CellRange::new(1, 0, 2, 5));
        store.exclude(CellRange::new(0, 2, 5, 3));
        let effective = store.effective_ranges().unwrap();
        for r in 0..6 {
            for c in 0..6 {
                let rules: Vec<_> = effective
                    .iter()
                    .filter(|(range, _)| range.contains(r, c))
                    .collect();
                assert!(rules.len() <= 1);
                assert_eq!(rules.first().map(|(_, rule)| rule), store.get(r, c));
            }
        }
        let mut full_column = ValidationStore::new();
        full_column.set(
            CellRange::new(0, 0, crate::sheet::NUM_ROWS - 1, 0),
            ValidationRule::custom("=A1>0"),
        );
        full_column.exclude(CellRange::new(1, 0, crate::sheet::NUM_ROWS - 2, 0));
        assert_eq!(full_column.effective_ranges().unwrap().len(), 2);
    }

    #[test]
    fn invalid_origins_and_excessive_fragmentation_are_refused() {
        let mut store = ValidationStore::new();
        let mut rule = ValidationRule::custom("=A1>0");
        rule.reference_origin = Some((100, 0));
        assert!(store
            .plan_edit(
                &[CellRange::single(0, 0)],
                ValidationEdit::Set(rule),
                10,
                10
            )
            .is_err());
        store.set(
            CellRange::new(0, 0, 10000, 0),
            ValidationRule::custom("=A1>0"),
        );
        for r in (1..10000).step_by(2) {
            store.exclude(CellRange::single(r, 0));
        }
        assert!(store.effective_ranges().is_err());
    }

    #[test]
    fn many_disjoint_rules_do_not_hit_the_fragment_budget() {
        for columns in [false, true] {
            let mut store = ValidationStore::new();
            for i in 0..10_000 {
                let (r, c) = if columns { (0, i) } else { (i, 0) };
                store.set(
                    CellRange::single(r, c),
                    ValidationRule::list_inline(vec!["Y".into()]),
                );
            }
            assert_eq!(store.effective_ranges().unwrap().len(), 10_000);
        }
    }
    #[test]
    fn structural_edits_rebase_at_surviving_cells_and_follow_cross_sheet_sources() {
        use crate::structural::Axis;
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(1, 0, "10");
        wb.active_sheet_mut().set_value(2, 0, "20");
        let mut rule = ValidationRule::whole_number(NumericConstraint::greater_than(
            ConstraintValue::CellRef("$A2".into()),
        ));
        rule.reference_origin = Some((1, 1));
        wb.active_sheet_mut()
            .validations
            .set(CellRange::new(1, 1, 3, 1), rule);
        let other = wb.add_sheet_named("Other").unwrap();
        let mut rule = ValidationRule::whole_number(NumericConstraint::greater_than(
            ConstraintValue::CellRef("Sheet1!$A2".into()),
        ));
        rule.reference_origin = Some((1, 1));
        wb.sheet_mut(other)
            .unwrap()
            .validations
            .set(CellRange::new(1, 1, 3, 1), rule);
        wb.structural_edit(0, Axis::Row, 0, 1, false).unwrap();
        assert!(wb.validate_cell_input(0, 2, 1, "15").is_valid());
        assert!(wb.validate_cell_input(0, 3, 1, "15").is_invalid());
        assert!(wb.validate_cell_input(other, 1, 1, "15").is_valid());
        assert!(wb.validate_cell_input(other, 2, 1, "15").is_invalid());
        let before = wb.clone();
        let (after, commit) = wb
            .prepare_guarded_structure(
                0,
                vec![crate::workbook::StructureStep {
                    axis: Axis::Row,
                    at: 2,
                    count: 1,
                    delete: true,
                }],
            )
            .unwrap();
        wb = after;
        assert!(wb.validate_cell_input(0, 2, 1, "15").is_invalid());
        assert!(wb.validate_cell_input(0, 2, 1, "25").is_valid());
        let rule = wb.sheet(0).unwrap().validations.get(2, 1).unwrap();
        assert_eq!(rule.reference_origin, Some((2, 1)));
        assert_eq!(
            rule.rule_type,
            ValidationType::WholeNumber(NumericConstraint::greater_than(ConstraintValue::CellRef(
                "$A3".into()
            )))
        );
        let after = wb.clone();
        commit.replay(&mut wb, true).unwrap();
        for i in 0..2 {
            assert_eq!(
                wb.sheet(i).unwrap().validations,
                before.sheet(i).unwrap().validations
            );
        }
        commit.replay(&mut wb, false).unwrap();
        for i in 0..2 {
            assert_eq!(
                wb.sheet(i).unwrap().validations,
                after.sheet(i).unwrap().validations
            );
        }
    }
}
