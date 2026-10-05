use super::*;
use crate::validation::{ComparisonOperator, ConstraintValue, NumericConstraint, ValidationType};

fn rule(text: &str) -> ValidationRule {
    ValidationRule::list_inline(vec![text.into()])
}

#[test]
fn exact_edits_match_a_cell_oracle_even_when_fragment_order_changes() {
    // A late fragment of the first rule sorts after the second rule. Naively
    // splitting the two original rectangles changes precedence at (3, 2).
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, 4, 4), rule("first"));
    original.set(CellRange::new(1, 1, 5, 5), rule("second"));
    original.set(CellRange::new(3, 0, 5, 3), rule("third"));
    original.exclude(CellRange::single(3, 3));
    let replacement = rule("new");
    for top in 0..6 {
        for left in 0..6 {
            for bottom in top..6 {
                for right in left..6 {
                    let cut = CellRange::new(top, left, bottom, right);
                    for set in [false, true] {
                        let edit = if set {
                            ValidationEdit::Set(replacement.clone())
                        } else {
                            ValidationEdit::Clear
                        };
                        let patch = original.plan_edit(&[cut], edit, 6, 6).unwrap();
                        let mut next = original.clone();
                        patch.apply(&mut next, true).unwrap();
                        for row in 0..6 {
                            for col in 0..6 {
                                let expected =
                                    if cut.contains(row, col) && !original.is_excluded(row, col) {
                                        if set {
                                            Some(&replacement)
                                        } else {
                                            None
                                        }
                                    } else {
                                        original.get(row, col)
                                    };
                                assert_eq!(
                                    next.get(row, col),
                                    expected,
                                    "{cut:?} at {row},{col}; set={set}"
                                );
                            }
                        }
                        let after = next.clone();
                        patch.apply(&mut next, false).unwrap();
                        assert_eq!(next, original, "Undo restores shadowed definitions too");
                        patch.apply(&mut next, true).unwrap();
                        assert_eq!(next, after);
                    }
                }
            }
        }
    }
}

#[test]
fn discontiguous_targets_leave_hidden_rows_and_columns_unchanged() {
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, 9, 4), rule("old"));
    let targets = [
        CellRange::new(1, 1, 1, 2),
        CellRange::single(4, 1),
        CellRange::single(8, 1),
    ];
    let mut next = original.clone();
    let patch = original
        .plan_edit(&targets, ValidationEdit::Set(rule("new")), 10, 5)
        .unwrap();
    patch.apply(&mut next, true).unwrap();
    for row in 0..10 {
        for col in 0..5 {
            assert_eq!(
                next.get(row, col),
                Some(&rule(if targets.iter().any(|r| r.contains(row, col)) {
                    "new"
                } else {
                    "old"
                }))
            );
        }
    }
    patch.apply(&mut next, false).unwrap();
    assert_eq!(next, original);
}

#[test]
fn partial_exclusion_clear_preserves_unselected_excluded_cells() {
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, 8, 8), rule("old"));
    original.exclude(CellRange::new(1, 1, 6, 6));
    original.exclude(CellRange::new(3, 3, 7, 7));
    let targets = [CellRange::new(2, 2, 4, 4), CellRange::new(3, 3, 5, 5)];
    for clear in [false, true] {
        let mut next = original.clone();
        let op = if clear {
            ValidationEdit::ClearExclusions
        } else {
            ValidationEdit::Exclude
        };
        let patch = original.plan_edit(&targets, op, 9, 9).unwrap();
        patch.apply(&mut next, true).unwrap();
        assert_eq!(next.rules, original.rules);
        for row in 0..9 {
            for col in 0..9 {
                let targeted = targets.iter().any(|r| r.contains(row, col));
                assert_eq!(
                    next.is_excluded(row, col),
                    if targeted {
                        !clear
                    } else {
                        original.is_excluded(row, col)
                    }
                );
            }
        }
        patch.apply(&mut next, false).unwrap();
        assert_eq!(next, original);
    }
}

#[test]
fn stale_replay_refuses_atomically_including_new_overlapping_metadata() {
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, 9, 0), rule("old"));
    let patch = original
        .plan_edit(&[CellRange::single(4, 0)], ValidationEdit::Clear, 10, 10)
        .unwrap();
    let mut stale = original.clone();
    stale.set(CellRange::new(2, 0, 8, 0), rule("added later"));
    let before_refusal = stale.clone();
    assert!(patch.apply(&mut stale, true).is_err());
    assert_eq!(stale, before_refusal);
    patch.apply(&mut original, true).unwrap();
    original.exclude(CellRange::single(5, 0));
    let before_refusal = original.clone();
    assert!(patch.apply(&mut original, false).is_err());
    assert_eq!(original, before_refusal);
}

#[test]
fn splitting_keeps_fixed_formula_references_and_unrelated_definitions() {
    let mut original = ValidationStore::new();
    let numeric = ValidationRule::new(ValidationType::Decimal(NumericConstraint {
        operator: ComparisonOperator::Between,
        value1: ConstraintValue::CellRef("'Limits'!$B$2".into()),
        value2: Some(ConstraintValue::Formula("=MAX(Limits!C1:C5)".into())),
    }));
    original.set(CellRange::new(0, 0, 9, 0), numeric.clone());
    // This overlapping component is unrelated and must retain both original
    // definitions, even the shadowed one.
    original.set(CellRange::new(0, 5, 9, 8), rule("unrelated"));
    original.set(CellRange::new(1, 6, 8, 7), rule("shadowed"));
    let patch = original
        .plan_edit(&[CellRange::single(4, 0)], ValidationEdit::Clear, 10, 10)
        .unwrap();
    let mut next = original.clone();
    patch.apply(&mut next, true).unwrap();
    assert_eq!(next.get(5, 0), Some(&numeric));
    assert_eq!(
        next.rules.get(&CellRange::new(1, 6, 8, 7)),
        Some(&rule("shadowed"))
    );
    assert!(patch.changed_ranges().all(|r| r.end_col == 0));
}

#[test]
fn empty_duplicate_and_already_excluded_edits_are_noops() {
    let mut original = ValidationStore::new();
    let range = CellRange::new(1, 1, 9, 1);
    original.set(range, rule("old"));
    original.exclude(range);
    for (targets, edit) in [
        (vec![], ValidationEdit::Clear),
        (vec![range, range], ValidationEdit::Set(rule("old"))),
        (
            vec![CellRange::single(4, 1)],
            ValidationEdit::Set(rule("old")),
        ),
        (vec![CellRange::single(4, 1)], ValidationEdit::Exclude),
        (vec![CellRange::single(4, 4)], ValidationEdit::Clear),
        (
            vec![CellRange::single(4, 4)],
            ValidationEdit::ClearExclusions,
        ),
    ] {
        assert!(original
            .plan_edit(&targets, edit, 10, 10)
            .unwrap()
            .is_empty());
    }
}

#[test]
fn full_column_rules_are_split_without_expanding_cells() {
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, usize::MAX - 1, 0), rule("all"));
    let patch = original
        .plan_edit(
            &[CellRange::single(50, 0)],
            ValidationEdit::Clear,
            usize::MAX,
            1,
        )
        .unwrap();
    let mut next = original.clone();
    patch.apply(&mut next, true).unwrap();
    assert_eq!(next.len(), 2);
    assert_eq!(next.get(usize::MAX - 1, 0), Some(&rule("all")));
    assert_eq!(next.get(50, 0), None);
}

#[test]
fn malformed_ranges_and_excessive_work_refuse_without_mutation() {
    let mut original = ValidationStore::new();
    original.set(CellRange::new(0, 0, 9, 9), rule("all"));
    let before = original.clone();
    for range in [
        CellRange::single(10, 0),
        CellRange {
            start_row: 9,
            end_row: 1,
            start_col: 0,
            end_col: 0,
        },
    ] {
        assert!(original
            .plan_edit(&[range], ValidationEdit::Clear, 10, 10)
            .is_err());
    }
    assert!(original
        .plan_edit(
            &vec![CellRange::single(0, 0); MAX_FRAGMENTS + 1],
            ValidationEdit::Clear,
            10,
            10
        )
        .is_err());
    let targets: Vec<_> = (0..2000).map(|r| CellRange::single(r * 2, 0)).collect();
    assert!(original
        .plan_edit(&targets, ValidationEdit::Exclude, 5000, 10)
        .is_err());
    assert_eq!(original, before);
}

#[test]
fn non_finite_constraints_cannot_bypass_the_replay_guard() {
    let original = ValidationStore::new();
    for number in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let invalid = ValidationRule::new(ValidationType::Decimal(NumericConstraint::between(
            0.0, number,
        )));
        assert!(original
            .plan_edit(
                &[CellRange::single(0, 0)],
                ValidationEdit::Set(invalid),
                10,
                10
            )
            .is_err());
    }
    let patch = original
        .plan_edit(
            &[CellRange::single(0, 0)],
            ValidationEdit::Set(rule("valid")),
            10,
            10,
        )
        .unwrap();
    let mut stale = original.clone();
    stale.set(
        CellRange::single(1, 0),
        ValidationRule::new(ValidationType::Decimal(NumericConstraint::between(
            0.0,
            f64::INFINITY,
        ))),
    );
    assert!(patch.apply(&mut stale, true).is_err());
    assert!(stale.get(0, 0).is_none());
}
